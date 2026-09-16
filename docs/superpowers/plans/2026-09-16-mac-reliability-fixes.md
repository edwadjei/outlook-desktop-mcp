# macOS Reliability Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the macOS Outlook MCP server's reply, save-attachment, sent-folder listing and liveness behaviour correct and bounded, and let `search_emails` filter by recipient and sender.

**Architecture:** Every tool builds an AppleScript string and runs it through `bridge.run()` (bounded at 120 s), or reads Outlook's SQLite profile database through `OutlookDB` (read-only). The fixes correct two AppleScript commands against Outlook's dictionary, put a deadline on database queries, defer the startup trust check to first use, add a `ping` tool, confirm the Sent Items copy after sends, and extend the database search with recipient and sender filters.

**Tech Stack:** Python 3.13, `mcp` FastMCP (1.x), `osascript`, `sqlite3`. Tests are plain scripts with a `check()` helper (no pytest).

**Spec:** `docs/superpowers/specs/2026-09-16-mac-reliability-fixes-design.md`

## Global Constraints

- Python to run everything: `~/.mcp-venvs/outlook-desktop/bin/python` (the `mcp` package is only installed there).
- Gate, green before every commit:
  `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py && ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py && ~/.mcp-venvs/outlook-desktop/bin/pip wheel -q --no-deps -w /tmp/odm-wheel "$PWD" && rm -rf /tmp/odm-wheel`
- Windows code (`server.py`, `com_bridge.py`) is not touched.
- Commit author: `git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit ...`. Subject as named per task. No AI references, no `Co-Authored-By` trailer. Stage by explicit path only.
- Outlook runs one AppleScript at a time: before any live probe run `ps -axo pid,etime,command | grep "[o]sascript -e"` and wait for it to clear. Never kill another session's script.
- Live sends go only to `Edward.Adjei@mtn.com` (the signed-in test account; self-sends are allowed). Run the live test at most once per task.
- AppleScript record fields are read from a local variable (`set senderRec to sender of m`, then `address of senderRec`), never through a chained specifier.
- Existing checks must keep passing: 118 in `tests/mac_batch_test.py`, 68 in `tests/mac_db_test.py` before Task 1.

---

### Task 1: `reply_email` uses the dictionary's `reply to` command; live send/reply test

**Files:**
- Modify: `src/outlook_desktop_mcp/server_mac.py` (function `reply_email`, ~lines 714-772)
- Modify: `tests/mac_batch_test.py` (add test + `main()` entry)
- Create: `tests/mac_live_test.py`

**Interfaces:**
- Consumes: `server_mac.reply_email(entry_id, body, reply_all=False, html_body="")` returns `"Reply sent to '<subject>' (reply_all=<bool>)"` or `"Error replying to email: ..."`.
- Produces: `tests/mac_live_test.py` with `wait_for_count(folder, subject, minimum)` and `check()`; Task 3 and Task 5 extend this file.

- [ ] **Step 1: Write the failing shape test**

Append to `tests/mac_batch_test.py` before `def main():`:

```python
def _outlook_running():
    import subprocess
    return subprocess.run(["pgrep", "-x", "Microsoft Outlook"],
                          capture_output=True).returncode == 0


def _compiles(script):
    """Compile (never run) an AppleScript with osacompile. Returns (ok, stderr)."""
    import subprocess
    with tempfile.TemporaryDirectory() as d:
        src = os.path.join(d, "s.applescript")
        with open(src, "w") as fh:
            fh.write(script)
        proc = subprocess.run(["osacompile", "-o", os.path.join(d, "s.scpt"), src],
                              capture_output=True, text=True)
        return proc.returncode == 0, proc.stderr.strip()


def test_reply_email_uses_dictionary_reply_command():
    log("--- reply_email uses 'reply to' with the 'reply to all' parameter ---")
    for reply_all in (False, True):
        fake = FakeBridge(output="Subj")
        server_mac.bridge = fake
        result = asyncio.run(server_mac.reply_email(entry_id="1", body="Test", reply_all=reply_all))
        script = fake.scripts[0]
        check(f"reply_all={reply_all}: reported success", result.startswith("Reply sent"), result)
        check(f"reply_all={reply_all}: no invented 'reply all to' command",
              "reply all to" not in script)
        expected = ("set replyMsg to reply to m with reply to all without opening window"
                    if reply_all else "set replyMsg to reply to m without opening window")
        check(f"reply_all={reply_all}: reply command", expected in script, script[:400])
        if _outlook_running():
            ok, err = _compiles(script)
            check(f"reply_all={reply_all}: script compiles against Outlook", ok, err)
        else:
            log("  SKIP: compile check (Outlook not running)")
```

Add `import tempfile` to the imports at the top of the file and `test_reply_email_uses_dictionary_reply_command()` as the last call in `main()`.

- [ ] **Step 2: Run it to verify it fails**

Run: `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py 2>&1 | grep -E "FAIL|checks passed"`
Expected: `FAIL: reply_all=True: no invented 'reply all to' command`, `FAIL: reply_all=True: reply command`, `FAIL: reply_all=True: script compiles against Outlook` (when Outlook is running); total shows 3 failures.

- [ ] **Step 3: Fix the generated command**

In `reply_email`, replace

```python
    reply_cmd = "reply all to" if reply_all else "reply to"
```

with

```python
    # Outlook's dictionary: `reply to <message>` with boolean parameters
    # `reply to all` and `opening window`. There is no `reply all to` command.
    reply_opts = ("with reply to all without opening window" if reply_all
                  else "without opening window")
```

and replace the script line

```
    set replyMsg to {reply_cmd} m without opening window
```

with

```
    set replyMsg to reply to m {reply_opts}
```

- [ ] **Step 4: Run the gate**

Run the gate command from Global Constraints. Expected: all checks pass (118 + new ones in batch; 68 in db), wheel builds.

- [ ] **Step 5: Write the live test**

Create `tests/mac_live_test.py`:

```python
"""
Outlook Desktop MCP - macOS live send/reply test
=================================================
Opt-in end-to-end test against the running Outlook app. Sends one seed
message to a scratch address that delivers back into this mailbox, waits
for the inbox copy, replies to it with reply_all=False and reply_all=True,
and asserts every sent item shows up in Sent Items.

Environment:
  OUTLOOK_MCP_LIVE_SCRATCH   address to send to (required; otherwise SKIP)
  OUTLOOK_MCP_LIVE_INLINE_ID message id with an inline image (optional)

Run: ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py
Requires Outlook for Mac in legacy mode with the account signed in.
"""
import sys
import os
import json
import asyncio
import time
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac, outlook_db
from outlook_desktop_mcp.outlook_db import OutlookDB

SCRATCH = os.environ.get("OUTLOOK_MCP_LIVE_SCRATCH", "").strip()
INLINE_ID = os.environ.get("OUTLOOK_MCP_LIVE_INLINE_ID", "").strip()
ROUND_TRIP = 120  # seconds allowed for an Exchange round trip


def log(msg):
    print(msg, file=sys.stderr, flush=True)


passed = 0
total = 0


def check(name, condition, detail=""):
    global passed, total
    total += 1
    if condition:
        passed += 1
        log(f"  PASS: {name}")
    else:
        log(f"  FAIL: {name} {detail}")


async def matches(folder, subject):
    """Messages in folder whose subject (ignoring Re:/Fw:) equals subject."""
    raw = await server_mac.search_emails(query=subject, folder=folder, count=20)
    rows = json.loads(raw)
    if not isinstance(rows, list):
        return []
    return [r for r in rows if r.get("subject", "").strip().lower() == subject.lower()]


async def wait_for_count(folder, subject, minimum, timeout=ROUND_TRIP):
    """Poll until at least `minimum` matching messages are in folder."""
    deadline = time.time() + timeout
    while True:
        rows = await matches(folder, subject)
        if len(rows) >= minimum or time.time() >= deadline:
            return rows
        await asyncio.sleep(3)


async def start_server():
    await server_mac.bridge.start()
    path = outlook_db.locate()
    if path and await server_mac._trust_db(OutlookDB(path)):
        server_mac.db = OutlookDB(path)
    log(f"  database: {'trusted' if server_mac.db else 'not used'}")


async def run():
    await start_server()
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    subject = f"[MCP live test {stamp}] reply regression"

    log("--- seed message ---")
    result = await server_mac.send_email(
        to=SCRATCH, subject=subject,
        body="Seed for the reply regression test. Safe to delete.\n\nBr,\nEdward.")
    check("send_email reported success", result.startswith("Email sent"), result)
    sent = await wait_for_count("sent", subject, 1)
    check("seed appears in Sent Items", len(sent) >= 1)
    received = await wait_for_count("inbox", subject, 1)
    check("seed delivered to inbox", len(received) >= 1)
    if not received:
        return
    seed_id = received[0]["entry_id"]

    log("--- replies ---")
    expected_sent = 1
    for reply_all in (False, True):
        result = await server_mac.reply_email(
            entry_id=seed_id, reply_all=reply_all,
            body=f"Reply with reply_all={reply_all}. Safe to delete.\n\nBr,\nEdward.")
        check(f"reply_all={reply_all}: reply_email reported success",
              result.startswith("Reply sent"), result)
        expected_sent += 1
        sent = await wait_for_count("sent", subject, expected_sent)
        check(f"reply_all={reply_all}: reply appears in Sent Items",
              len(sent) >= expected_sent, f"{len(sent)} in Sent Items")


def main():
    if not SCRATCH:
        log("SKIP: set OUTLOOK_MCP_LIVE_SCRATCH to run the live test")
        return
    asyncio.run(run())
    log("=" * 50)
    log(f"{passed}/{total} checks passed")
    if passed != total:
        sys.exit(1)


if __name__ == "__main__":
    main()
```

- [ ] **Step 6: Run the live test once**

Check no other AppleScript is running, then:
`OUTLOOK_MCP_LIVE_SCRATCH=Edward.Adjei@mtn.com ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py`
Expected: `7/7 checks passed`. Record the exact output in the report. If delivery takes longer than 120 s the inbox check fails; rerun once only after checking Outlook is online.

- [ ] **Step 7: Commit**

```bash
git add src/outlook_desktop_mcp/server_mac.py tests/mac_batch_test.py tests/mac_live_test.py
git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit -m "macOS: fix reply_email reply-all command; add live send/reply test"
```

---

### Task 2: Bounded database queries, deferred trust check, `ping` tool

**Files:**
- Modify: `src/outlook_desktop_mcp/outlook_db.py` (`OutlookDB.__init__`, `_query`, module constants)
- Modify: `src/outlook_desktop_mcp/server_mac.py` (module state after `bridge = AppleScriptBridge()`, `_db_folder_id`, `_trust_db`, new `_ensure_db`, new `ping` tool, `startup`, `main`)
- Modify: `tests/mac_db_test.py`, `tests/mac_batch_test.py`, `tests/mac_live_test.py`, `README.md`

**Interfaces:**
- Consumes: `AppleScriptBridge.run(script, timeout=SCRIPT_TIMEOUT)`; `STARTUP_TIMEOUT` (10 s) from `applescript_bridge`.
- Produces: `OutlookDB(path, timeout=DB_TIMEOUT)`; `outlook_db.DB_TIMEOUT` (float, env `OUTLOOK_MCP_DB_TIMEOUT`, default 20); `server_mac._ensure_db() -> OutlookDB | None`; `server_mac.startup()` coroutine; `server_mac._trust_db(candidate) -> bool | None`; module globals `_db_candidate`, `_db_state` (`"none" | "unchecked" | "trusted" | "untrusted"`); tool `ping()`.

- [ ] **Step 1: Failing test for the query deadline**

Append to `tests/mac_db_test.py` (before `def main():`) and add `import time` at the top:

```python
def test_query_deadline_interrupts_long_query():
    log("--- a query past the deadline is interrupted and raises OutlookDBError ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d), timeout=0.2)
        t0 = time.time()
        try:
            db._query("WITH RECURSIVE c(x) AS (SELECT 1 UNION ALL SELECT x + 1 FROM c) "
                      "SELECT COUNT(*) FROM c")
            check("raised OutlookDBError", False, "query finished")
        except outlook_db.OutlookDBError as e:
            check("raised OutlookDBError", "exceeded" in str(e), str(e))
        check("returned promptly", time.time() - t0 < 2.0, f"{time.time() - t0:.2f}s")
        check("normal query still works", db.resolve_folder("inbox") == 115)
    check("default deadline is 20s", outlook_db.DB_TIMEOUT == 20.0, str(outlook_db.DB_TIMEOUT))
```

Add the call to `main()`.

- [ ] **Step 2: Run to verify it fails**

Run: `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py 2>&1 | tail -3`
Expected: `TypeError: OutlookDB.__init__() got an unexpected keyword argument 'timeout'`.

- [ ] **Step 3: Implement the deadline**

In `outlook_db.py` add `import threading` and, after `_BUSY_RETRIES = 3`:

```python
def _env_seconds(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw:
        try:
            return float(raw)
        except ValueError:
            logger.warning("Invalid %s=%r; using default %ss", name, raw, default)
    return default


# Deadline for one SQL query. A query that runs past it is interrupted and
# reported as OutlookDBError so the caller falls back to AppleScript instead
# of hanging the tool call.
DB_TIMEOUT = _env_seconds("OUTLOOK_MCP_DB_TIMEOUT", 20)
```

Change the constructor and `_query`:

```python
    def __init__(self, path: str, timeout: float = DB_TIMEOUT):
        self.path = path
        self.timeout = timeout

    def _query(self, sql: str, params=()) -> list[sqlite3.Row]:
        last = None
        for attempt in range(_BUSY_RETRIES):
            con = None
            timer = None
            try:
                con = self._connect()
                timer = threading.Timer(self.timeout, _interrupt, (con,))
                timer.daemon = True
                timer.start()
                return con.execute(sql, params).fetchall()
            except sqlite3.OperationalError as e:
                msg = str(e).lower()
                if "interrupted" in msg:
                    raise OutlookDBError(
                        f"Outlook database query exceeded {self.timeout:g}s"
                    ) from e
                if "locked" in msg or "busy" in msg:
                    last = e
                    time.sleep(0.2 * (attempt + 1))
                    continue
                raise OutlookDBError(f"Outlook database query failed: {e}") from e
            except sqlite3.DatabaseError as e:
                raise OutlookDBError(f"Outlook database unusable: {e}") from e
            finally:
                if timer is not None:
                    timer.cancel()
                if con is not None:
                    con.close()
        raise OutlookDBError(f"Outlook database busy: {last}")
```

and add the module-level helper (near `_iso`):

```python
def _interrupt(con: sqlite3.Connection) -> None:
    """Timer callback: abort the running statement on this connection."""
    try:
        con.interrupt()
    except sqlite3.ProgrammingError:
        pass  # connection already closed
```

- [ ] **Step 4: Run the db suite**

Run: `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py 2>&1 | grep -E "FAIL|checks passed"`
Expected: all pass (68 + 4).

- [ ] **Step 5: Failing tests for deferred trust and `ping`**

Replace `test_server_trust_check` in `tests/mac_db_test.py` with:

```python
def test_server_trust_check():
    log("--- trust check compares database inbox with AppleScript ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        fake = FakeBridge(output="Inbox|||4")
        server_mac.bridge = fake
        check("agreeing counts -> trusted", asyncio.run(server_mac._trust_db(db)) is True)
        check("probe asks for folder id 115", "mail folder id 115" in fake.scripts[0])
        server_mac.bridge = FakeBridge(output="Inbox|||6")
        check("small drift -> trusted", asyncio.run(server_mac._trust_db(db)) is True)
        server_mac.bridge = FakeBridge(output="Inbox|||900")
        check("large drift -> untrusted", asyncio.run(server_mac._trust_db(db)) is False)

        class ErrBridge(FakeBridge):
            async def run(self, script, timeout=None):
                raise RuntimeError("AppleScript error: folder not found")

        server_mac.bridge = ErrBridge()
        check("folder id unknown to AppleScript -> untrusted",
              asyncio.run(server_mac._trust_db(db)) is False)

        class SlowBridge(FakeBridge):
            def __init__(self):
                super().__init__()
                self.timeouts = []

            async def run(self, script, timeout=None):
                self.timeouts.append(timeout)
                raise RuntimeError(f"AppleScript timed out after {timeout}s")

        slow = SlowBridge()
        server_mac.bridge = slow
        check("AppleScript timeout -> undecided (None)", asyncio.run(server_mac._trust_db(db)) is None)
        check("probe uses the startup timeout", slow.timeouts == [server_mac.STARTUP_TIMEOUT],
              str(slow.timeouts))


def _reset_db_state():
    server_mac.db = None
    server_mac._db_candidate = None
    server_mac._db_state = "none"
    server_mac._db_lock = None


def test_server_trust_check_is_deferred_to_first_use():
    log("--- trust check runs on first database use, not at startup ---")
    with tempfile.TemporaryDirectory() as d:
        path = build_fixture(d)
        os.environ[outlook_db.ENV_VAR] = path
        try:
            _reset_db_state()

            class VersionBridge(FakeBridge):
                async def run(self, script, timeout=None):
                    self.scripts.append(script)
                    return "16.93.2"

            vb = VersionBridge()
            server_mac.bridge = vb
            asyncio.run(server_mac.startup())
            check("startup ran only the version probe", len(vb.scripts) == 1 and "version" in vb.scripts[0])
            check("startup leaves db unset", server_mac.db is None)
            check("state is unchecked", server_mac._db_state == "unchecked")

            # First use: AppleScript busy -> timeout -> stays unchecked, tool still answers.
            class SlowBridge(FakeBridge):
                async def run(self, script, timeout=None):
                    self.scripts.append(script)
                    raise RuntimeError(f"AppleScript timed out after {timeout}s")

            server_mac.bridge = SlowBridge()
            fid = asyncio.run(server_mac._db_folder_id("inbox"))
            check("timeout keeps state unchecked", server_mac._db_state == "unchecked")
            check("no folder id while undecided", fid is None)

            # Next use: AppleScript agrees -> trusted, and stays trusted.
            server_mac.bridge = FakeBridge(output="Inbox|||4")
            fid = asyncio.run(server_mac._db_folder_id("inbox"))
            check("agreement -> trusted", server_mac._db_state == "trusted" and server_mac.db is not None)
            check("folder id resolved from database", fid == 115)
            server_mac.bridge = FakeBridge(output="Inbox|||900")
            fid = asyncio.run(server_mac._db_folder_id("inbox"))
            check("trusted verdict is cached", fid == 115 and len(server_mac.bridge.scripts) == 0)

            # Mismatch on first use -> untrusted, cached, AppleScript not asked again.
            _reset_db_state()
            server_mac.bridge = VersionBridge()
            asyncio.run(server_mac.startup())
            bad = FakeBridge(output="Inbox|||900")
            server_mac.bridge = bad
            asyncio.run(server_mac._db_folder_id("inbox"))
            check("mismatch -> untrusted", server_mac._db_state == "untrusted")
            asyncio.run(server_mac._db_folder_id("inbox"))
            check("untrusted verdict is cached", len(bad.scripts) == 1)
        finally:
            del os.environ[outlook_db.ENV_VAR]
            _reset_db_state()


def test_ping_tool():
    log("--- ping reports Outlook reachability and database state ---")
    _reset_db_state()
    fake = FakeBridge(output="16.93.2")
    server_mac.bridge = fake
    result = json.loads(asyncio.run(server_mac.ping()))
    check("ok when Outlook answers", result["ok"] is True, str(result))
    check("reports version", result["outlook_version"] == "16.93.2")
    check("reports db state", result["db"] == "none")
    check("reports uptime", isinstance(result["uptime_s"], int))
    check("reports server version", isinstance(result["server_version"], str) and result["server_version"])
    check("applescript_ms measured", isinstance(result["applescript_ms"], int))

    class DeadBridge(FakeBridge):
        def __init__(self):
            super().__init__()
            self.timeouts = []

        async def run(self, script, timeout=None):
            self.timeouts.append(timeout)
            raise RuntimeError("AppleScript timed out after 10s")

    dead = DeadBridge()
    server_mac.bridge = dead
    result = json.loads(asyncio.run(server_mac.ping()))
    check("not ok when Outlook is silent", result["ok"] is False)
    check("error carried", "timed out" in result.get("error", ""))
    check("ping uses the short startup timeout", dead.timeouts == [server_mac.STARTUP_TIMEOUT])
    names = [t.name for t in asyncio.run(server_mac.mcp.list_tools())]
    check("ping is registered as a tool", "ping" in names)
```

Add `test_server_trust_check_is_deferred_to_first_use()` and `test_ping_tool()` to `main()`.

- [ ] **Step 6: Run to verify they fail**

Run: `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py 2>&1 | grep -E "FAIL|Error|checks passed" | head`
Expected: failures/AttributeError on `STARTUP_TIMEOUT`, `_db_candidate`, `startup`, `ping`.

- [ ] **Step 7: Implement deferred trust, `ping`, `startup`**

In `server_mac.py`: add `import time` and change the bridge import to
`from outlook_desktop_mcp.applescript_bridge import AppleScriptBridge, STARTUP_TIMEOUT`.

Replace the block from `db: OutlookDB | None = None` through the end of `_db_folder_id` with:

```python
# Read-only handle on legacy Outlook's profile database. Set only after
# _ensure_db() has confirmed it is the store AppleScript is reading.
# None means every tool uses AppleScript only.
db: OutlookDB | None = None

# Candidate database located at startup, and the trust decision for it:
#   "none"      no database file found
#   "unchecked" found; trust check not yet completed (retried on next use)
#   "trusted"   check passed; `db` is set
#   "untrusted" check failed (stale store or unusable file); never retried
_db_candidate: OutlookDB | None = None
_db_state = "none"
_db_lock: asyncio.Lock | None = None
_STARTED_AT = time.monotonic()

# Startup trust check: the database inbox count may differ from
# AppleScript's by this much (mail arriving while we check) and still be
# trusted. A stale database left behind by a switch to New Outlook drifts
# far beyond this.
_TRUST_DRIFT_MIN = 50
_TRUST_DRIFT_FRACTION = 0.02


async def _ensure_db() -> OutlookDB | None:
    """Return the trusted database handle, running the trust check on first use.

    The check needs one short AppleScript call. Outlook serialises scripts
    across every client, so doing it at startup could stall past the MCP
    client's connect timeout; doing it here keeps startup instant. A timeout
    leaves the decision open for the next call; a real mismatch is final.
    """
    global db, _db_state, _db_lock
    if db is not None or _db_candidate is None or _db_state != "unchecked":
        return db
    if _db_lock is None:
        _db_lock = asyncio.Lock()
    async with _db_lock:
        if _db_state == "unchecked":
            verdict = await _trust_db(_db_candidate)
            if verdict is True:
                db = _db_candidate
                _db_state = "trusted"
                logger.info("Using Outlook profile database for list/search: %s", db.path)
            elif verdict is False:
                _db_state = "untrusted"
    return db


async def _db_folder_id(folder: str) -> int | None:
    """Resolve a folder name via the profile database, or None if unavailable."""
    handle = await _ensure_db()
    if handle is None:
        return None
    try:
        return await asyncio.to_thread(handle.resolve_folder, folder)
    except OutlookDBError as e:
        logger.warning("Outlook database folder lookup failed; using AppleScript: %s", e)
        return None
```

In `_trust_db`, change the signature to `-> bool | None`, run the probe as
`raw = await bridge.run(script, timeout=STARTUP_TIMEOUT)`, and replace the `except Exception as e:` branch with:

```python
    except Exception as e:
        if "timed out" in str(e):
            logger.warning("Outlook database trust check timed out; will retry on next use")
            return None
        logger.warning("Outlook database trust check failed (folder id %s): %s", fid, e)
        return False
```

Add the `ping` tool immediately before the `# TOOL 1: send_email` banner:

```python
# =====================================================================
# TOOL 0: ping
# =====================================================================

def _server_version() -> str:
    try:
        from importlib.metadata import version
        return version("outlook-desktop-mcp")
    except Exception:
        return "unknown"


@mcp.tool()
async def ping() -> str:
    """Check cheaply whether the server can reach Outlook.

    Call this before starting a long task, or when another tool has stopped
    answering. It runs one trivial AppleScript with a 10-second limit and
    never raises.

    Returns:
        JSON with ok (bool), outlook_version, applescript_ms, db
        ("trusted", "untrusted", "unchecked" or "none"), db_path,
        server_version, uptime_s, and error when ok is false.
    """
    started = time.monotonic()
    result = {
        "ok": False,
        "outlook_version": None,
        "applescript_ms": None,
        "db": _db_state,
        "db_path": _db_candidate.path if _db_candidate is not None else None,
        "server_version": _server_version(),
        "uptime_s": int(time.monotonic() - _STARTED_AT),
    }
    try:
        result["outlook_version"] = await bridge.run(
            'tell application "Microsoft Outlook" to get version',
            timeout=STARTUP_TIMEOUT,
        )
        result["ok"] = True
    except Exception as e:
        result["error"] = str(e)
    result["applescript_ms"] = int((time.monotonic() - started) * 1000)
    return json.dumps(result, indent=2)
```

Replace `main()` with:

```python
async def startup() -> None:
    """Verify Outlook answers and locate the profile database.

    The database trust check is deferred to first use (see _ensure_db) so
    startup never waits behind another client's AppleScript.
    """
    global _db_candidate, _db_state
    logger.info("Starting Outlook Desktop MCP server (macOS)...")
    await bridge.start()
    path = outlook_db.locate()
    if path is None:
        logger.info("No Outlook profile database found; list/search use AppleScript")
    else:
        _db_candidate = OutlookDB(path)
        _db_state = "unchecked"
        logger.info("Outlook profile database found (trust check on first use): %s", path)
    logger.info("AppleScript bridge ready. Starting MCP stdio transport...")


def main():
    asyncio.run(startup())
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()
```

Update `tests/mac_live_test.py`: replace the body of `start_server()` with

```python
async def start_server():
    await server_mac.startup()
    await server_mac._ensure_db()
    log(f"  database: {server_mac._db_state}")
```

- [ ] **Step 8: Run the gate**

Run the gate. Expected: batch 118 + Task 1 checks, db suite 68 + 4 + 8 + 12 + trust cases, wheel builds. Fix any failure in the code, not the test.

- [ ] **Step 9: README**

In the "### Email" tools table add a row after `list_emails`:
`| \`ping\` | no | yes | Liveness check: Outlook version, AppleScript round-trip time, database state; answers within 10 s |`.
In "### macOS: AppleScript Bridge" add a paragraph:

> **Every call is bounded.** Each AppleScript runs under `OUTLOOK_MCP_SCRIPT_TIMEOUT` (default 120 s) and each profile-database query under `OUTLOOK_MCP_DB_TIMEOUT` (default 20 s); a query past its deadline is interrupted and the tool falls back to AppleScript. A list or search call therefore never runs longer than about 260 s (database, batched script, legacy script), and most finish in well under a second. The database trust check runs on the first list or search call rather than at startup, so the server connects instantly even when Outlook is busy serving another client's script. Use `ping` to confirm Outlook is answering before a long task.

- [ ] **Step 10: Commit**

```bash
git add src/outlook_desktop_mcp/outlook_db.py src/outlook_desktop_mcp/server_mac.py tests/mac_db_test.py tests/mac_batch_test.py tests/mac_live_test.py README.md
git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit -m "macOS: bound database queries, defer trust check, add ping tool"
```

---

### Task 3: Confirm the Sent Items copy after send and reply; stable sent ordering

**Files:**
- Modify: `src/outlook_desktop_mcp/outlook_db.py` (`list_messages` ordering, new `normalize_subject`, new `find_sent_copy`)
- Modify: `src/outlook_desktop_mcp/server_mac.py` (new `_sent_copy_id`, `_confirmation_suffix`; `send_email`, `reply_email`, `list_emails` docstring)
- Modify: `tests/mac_db_test.py`, `tests/mac_batch_test.py`, `tests/mac_live_test.py`, `README.md`

**Interfaces:**
- Consumes: `server_mac._ensure_db()`, `server_mac._folder_ref(folder)` (returns `mail folder id N` or a keyword), `OutlookDB.resolve_folder`, `OutlookDB._row_to_summary`.
- Produces: `outlook_db.normalize_subject(str) -> str`; `OutlookDB.find_sent_copy(folder_id, subject, since) -> dict | None`; `server_mac._sent_copy_id(subject, since) -> str | None`; `server_mac._SENT_CONFIRM_TIMEOUT = 10.0`, `_SENT_CONFIRM_INTERVAL = 0.5`. Send/reply confirmations end with ` (Sent Items id N)` or ` (Sent Items copy not visible yet; verify with search_emails)`.

- [ ] **Step 1: Failing db tests**

Append to `tests/mac_db_test.py`:

```python
def test_list_messages_tie_breaks_on_record_id():
    log("--- equal timestamps are ordered by record id, newest first ---")
    mail = MAIL + [
        (301, 115, "Same second A", "A", "a@x.com", "", 0, 0, T0 + 500, 0, 0),
        (302, 115, "Same second B", "B", "b@x.com", "", 0, 0, T0 + 500, 0, 0),
    ]
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d, mail=mail))
        ids = [r["entry_id"] for r in db.list_messages(115, 2)]
        check("newest record first among equals", ids == ["302", "301"], str(ids))


def test_normalize_subject():
    log("--- normalize_subject strips reply and forward prefixes ---")
    check("RE:", outlook_db.normalize_subject("RE: budget") == "budget")
    check("nested", outlook_db.normalize_subject("Re: FW: Fwd: budget") == "budget")
    check("plain", outlook_db.normalize_subject("  budget  ") == "budget")
    check("inner Re kept", outlook_db.normalize_subject("Re: about the Re: thing") == "about the Re: thing")


def test_find_sent_copy():
    log("--- find_sent_copy returns the newest matching sent row after `since` ---")
    mail = MAIL + [
        (401, 127, "budget", "Me", "me@x.com", "older reply", 1, 0, T0 + 1000, 0, 0),
        (402, 127, "budget", "Me", "me@x.com", "newer reply", 1, 0, T0 + 2000, 0, 0),
        (403, 127, "budget", "Me", "me@x.com", "deleted", 1, 0, T0 + 3000, 1, 0),
    ]
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d, mail=mail))
        row = db.find_sent_copy(127, "RE: budget", since=T0 + 1500)
        check("newest live row after since", row is not None and row["entry_id"] == "402", str(row))
        check("older rows ignored", db.find_sent_copy(127, "budget", since=T0 + 2500) is None)
        check("case-insensitive subject", db.find_sent_copy(127, "BUDGET", since=T0) is not None)
        check("other folder ignored", db.find_sent_copy(115, "budget", since=T0) is None)


def test_server_send_email_confirms_sent_copy():
    log("--- send_email and reply_email report the Sent Items id ---")
    with tempfile.TemporaryDirectory() as d:
        mail = MAIL + [(501, 127, "Hello there", "Me", "me@x.com", "", 1, 0, int(time.time()) + 5, 0, 0)]
        server_mac.db = OutlookDB(build_fixture(d, mail=mail))
        server_mac.bridge = FakeBridge(output="")
        result = asyncio.run(server_mac.send_email(to="a@x.com", subject="Hello there", body="hi"))
        check("send confirmation carries the Sent Items id", result.endswith("(Sent Items id 501)"), result)
        server_mac.bridge = FakeBridge(output="RE: Hello there")
        result = asyncio.run(server_mac.reply_email(entry_id="1", body="hi"))
        check("reply confirmation carries the Sent Items id", result.endswith("(Sent Items id 501)"), result)

        old = server_mac._SENT_CONFIRM_TIMEOUT
        server_mac._SENT_CONFIRM_TIMEOUT = 0.6
        try:
            t0 = time.time()
            result = asyncio.run(server_mac.send_email(to="a@x.com", subject="Never lands", body="hi"))
            check("missing copy reported", result.endswith("(Sent Items copy not visible yet; verify with search_emails)"), result)
            check("gives up after the confirm timeout", 0.5 < time.time() - t0 < 3.0, f"{time.time() - t0:.2f}s")
        finally:
            server_mac._SENT_CONFIRM_TIMEOUT = old
    server_mac.db = None
```

Add all four to `main()`.

- [ ] **Step 2: Failing batch test (AppleScript fallback for the lookup)**

Append to `tests/mac_batch_test.py`:

```python
def test_send_email_looks_up_sent_copy_without_db():
    log("--- send_email falls back to an AppleScript Sent Items lookup ---")
    server_mac.db = None

    class TwoStep(FakeBridge):
        async def run(self, script, timeout=None):
            self.scripts.append(script)
            return "" if len(self.scripts) == 1 else "777"

    fake = TwoStep()
    server_mac.bridge = fake
    result = asyncio.run(server_mac.send_email(to="a@x.com", subject="Hi", body="x"))
    check("send then lookup", len(fake.scripts) == 2, str(len(fake.scripts)))
    if len(fake.scripts) == 2:
        check("lookup reads sent items", "sent items" in fake.scripts[1])
        check("lookup compares the subject", 'subject of m is "Hi"' in fake.scripts[1], fake.scripts[1])
    check("confirmation carries id", result.endswith("(Sent Items id 777)"), result)
```

Add to `main()`. In BOTH `tests/mac_batch_test.py` and `tests/mac_db_test.py`, make the first line of `main()`:

```python
    server_mac._SENT_CONFIRM_TIMEOUT = 0.0  # unit tests never wait for Outlook's Sent Items write
```

(`test_server_send_email_confirms_sent_copy` raises it to 0.6 for its timing check and restores it.) With a zero timeout the lookup still runs exactly once, so shape checks on the lookup script keep working. Any existing send/reply check that asserts `len(fake.scripts) == 1` must become `== 2` (send script, then the Sent Items lookup); `fake.scripts[0]` is still the send script.

Run both suites; expected failures: `AttributeError` on `normalize_subject`, `find_sent_copy`, and unmatched suffixes.

- [ ] **Step 3: Implement in `outlook_db.py`**

Add `import re` and:

```python
_PREFIX_RE = re.compile(r"^\s*(?:(?:re|fw|fwd|aw|sv|wg)\s*:\s*)+", re.IGNORECASE)


def normalize_subject(subject: str) -> str:
    """Subject without leading reply/forward prefixes, as Outlook stores it."""
    return _PREFIX_RE.sub("", subject or "").strip()
```

Change `list_messages` ordering to `"ORDER BY Message_TimeReceived DESC, Record_RecordID DESC LIMIT ?"` and add:

```python
    def find_sent_copy(self, folder_id: int, subject: str, since: int) -> dict | None:
        """Newest live row in folder with this normalized subject received at or after `since`.

        Used right after a send: Outlook keeps the message in the Outbox for
        a few seconds and then writes a new Sent Items row, so the caller
        polls this until it returns a row.
        """
        rows = self._query(
            f"SELECT {_MESSAGE_COLUMNS} FROM Mail "
            f"WHERE Record_FolderID = ? AND {_LIVE_ROWS} "
            "AND Message_NormalizedSubject = ? COLLATE NOCASE "
            "AND IFNULL(Message_TimeReceived, 0) >= ? "
            "ORDER BY Message_TimeReceived DESC, Record_RecordID DESC LIMIT 1",
            (folder_id, normalize_subject(subject), int(since)),
        )
        return self._row_to_summary(rows[0]) if rows else None
```

- [ ] **Step 4: Implement in `server_mac.py`**

Add after `_TRUST_DRIFT_FRACTION`:

```python
# After `send`, Outlook parks the message in the Outbox for a few seconds
# and then writes a new Sent Items record with a new id. Send tools poll
# for that record so their confirmation can name it.
_SENT_CONFIRM_TIMEOUT = 10.0
_SENT_CONFIRM_INTERVAL = 0.5


async def _sent_copy_id(subject: str, since: int) -> str | None:
    """Id of the Sent Items copy of a message sent at `since`, or None if not visible in time."""
    handle = await _ensure_db()
    deadline = time.monotonic() + _SENT_CONFIRM_TIMEOUT
    while True:
        try:
            if handle is not None:
                fid = await asyncio.to_thread(handle.resolve_folder, "sent")
                row = None
                if fid is not None:
                    row = await asyncio.to_thread(handle.find_sent_copy, fid, subject, since)
                if row is not None:
                    return row["entry_id"]
            else:
                folder_ref = await _folder_ref("sent")
                raw = await bridge.run(f'''tell application "Microsoft Outlook"
    set f to {folder_ref}
    set n to count of messages of f
    if n > 20 then set n to 20
    repeat with i from 1 to n
        set m to message i of f
        if subject of m is "{escape(subject)}" then return (id of m as text)
    end repeat
    return ""
end tell''')
                if raw.strip():
                    return raw.strip()
        except (OutlookDBError, RuntimeError) as e:
            logger.warning("Sent Items lookup failed: %s", e)
            return None
        if time.monotonic() >= deadline:
            return None
        await asyncio.sleep(_SENT_CONFIRM_INTERVAL)


def _confirmation_suffix(sent_id: str | None) -> str:
    if sent_id:
        return f" (Sent Items id {sent_id})"
    return " (Sent Items copy not visible yet; verify with search_emails)"
```

In `send_email`, replace the `try:` block with:

```python
    try:
        since = int(time.time())
        await bridge.run(script)
        sent_id = await _sent_copy_id(subject, since)
        return f"Email sent: '{subject}' to {to}{_confirmation_suffix(sent_id)}"
    except Exception as e:
        return f"Error sending email: {e}"
```

In `reply_email`, replace the `try:` block with:

```python
    try:
        since = int(time.time())
        subject = await bridge.run(script)
        sent_id = await _sent_copy_id(subject, since)
        return f"Reply sent to '{subject}' (reply_all={reply_all}){_confirmation_suffix(sent_id)}"
    except Exception as e:
        return f"Error replying to email: {e}"
```

Update the `Returns:` lines of both docstrings to mention the Sent Items id. In the `list_emails` docstring add to the `folder` argument text:

```
            Note: a message you just sent spends a few seconds in "outbox"
            before it appears in "sent"; the id in the send tool's
            confirmation, or search_emails, is the reliable check.
```

- [ ] **Step 5: Run the gate**

Expected: all pass. The `test_send_email_looks_up_sent_copy_without_db` check `'subject of m is "Hi"'` must match the generated script exactly.

- [ ] **Step 6: Extend the live test**

In `tests/mac_live_test.py` `run()`, after each `check(... reported success ...)` for the seed and for both replies add:

```python
        import re
        m = re.search(r"\(Sent Items id (\d+)\)", result)
        check("confirmation names the Sent Items id", m is not None, result)
        if m:
            copy = json.loads(await server_mac.read_email(entry_id=m.group(1)))
            check("Sent Items id readable", copy.get("subject", "").lower().endswith(subject.lower()), str(copy)[:200])
```

(for the seed use `result` from `send_email`; for replies the reply result). Run the live test once: `OUTLOOK_MCP_LIVE_SCRATCH=Edward.Adjei@mtn.com ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py`. Expected: all checks pass; record output.

- [ ] **Step 7: README**

In "#### Fast list and search on macOS" add a bullet:

> - A message you just sent sits in the Outbox for a few seconds, then Outlook writes a new Sent Items record with a new id. `send_email` and `reply_email` wait up to 10 s for that record and name it in their confirmation (`(Sent Items id N)`); `list_emails(folder="sent")` called inside that window will not show it yet. Verify sends by the confirmation id or `search_emails`, not by the top of `list_emails`.

- [ ] **Step 8: Commit**

```bash
git add src/outlook_desktop_mcp/outlook_db.py src/outlook_desktop_mcp/server_mac.py tests/mac_db_test.py tests/mac_batch_test.py tests/mac_live_test.py README.md
git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit -m "macOS: confirm Sent Items copy after send and reply; stable sent ordering"
```

---

### Task 4: `search_emails` by recipient and sender; document the preview limit

**Files:**
- Modify: `src/outlook_desktop_mcp/outlook_db.py` (`search_messages`)
- Modify: `src/outlook_desktop_mcp/server_mac.py` (`search_emails`)
- Modify: `tests/mac_db_test.py` (fixture columns, tests), `README.md`

**Interfaces:**
- Consumes: `OutlookDB._query`, `_like_pattern`, `_row_to_summary`; `server_mac._db_folder_id`.
- Produces: `OutlookDB.search_messages(folder_id, query, count, recipient="", sender="")`; `server_mac.search_emails(query="", folder="inbox", count=10, recipient="", sender="")`.

- [ ] **Step 1: Extend the fixture**

In `tests/mac_db_test.py` `build_fixture`, add three columns to `CREATE TABLE Mail` after `Message_Hidden INTEGER`:

```sql
            Message_ToRecipientAddressList TEXT,
            Message_CCRecipientAddressList TEXT,
            Message_DisplayTo TEXT
```

Change the insert to `INSERT INTO Mail VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)` and make it pad short tuples:

```python
    con.executemany("INSERT INTO Mail VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    [tuple(row) + (None,) * (14 - len(row)) for row in mail])
```

Update the `MAIL` comment line to `# id, folder, subject, sender name, sender addr, preview, read, att, time, del, hidden[, to addrs, cc addrs, display to]`.

- [ ] **Step 2: Failing tests**

Append:

```python
def test_search_by_recipient_and_sender():
    log("--- search filters by recipient and sender ---")
    mail = MAIL + [
        (601, 115, "New Service Integration - Kunim", "Derrick", "derrick@x.com", "please approve",
         0, 0, T0 + 600, 0, 0, "edward@x.com", "isdemand@x.com", "Edward Adjei"),
        (602, 115, "Weekly digest", "News", "news@x.com", "approve nothing",
         0, 0, T0 + 700, 0, 0, "all@x.com", "", "Everyone"),
        (603, 115, "Approval request", "Derrick", "derrick@x.com", "second one",
         0, 0, T0 + 800, 0, 0, "bob@x.com", "edward@x.com", "Bob"),
    ]
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d, mail=mail))
        ids = lambda rows: [r["entry_id"] for r in rows]
        check("recipient matches to and cc", ids(db.search_messages(115, "", 10, recipient="edward")) == ["603", "601"])
        check("recipient matches display name", ids(db.search_messages(115, "", 10, recipient="Adjei")) == ["601"])
        check("query AND recipient", ids(db.search_messages(115, "approv", 10, recipient="edward")) == ["603", "601"])
        check("sender filter", ids(db.search_messages(115, "", 10, sender="derrick")) == ["603", "601"])
        check("all three combined", ids(db.search_messages(115, "kunim", 10, recipient="edward", sender="derrick")) == ["601"])
        check("nothing given -> empty", db.search_messages(115, "", 10) == [])
        check("rows without recipient columns filled do not match", ids(db.search_messages(115, "", 10, recipient="x.com")) == ["603", "602", "601"])


def test_server_search_emails_filters():
    log("--- search_emails passes recipient and sender to the database ---")
    with tempfile.TemporaryDirectory() as d:
        mail = MAIL + [(701, 115, "Approve me please", "Derrick", "derrick@x.com", "", 0, 0, T0 + 900, 0, 0,
                        "edward@x.com", "", "Edward")]
        server_mac.db = OutlookDB(build_fixture(d, mail=mail))
        server_mac.bridge = FakeBridge(output="")
        rows = json.loads(asyncio.run(server_mac.search_emails(query="approve", recipient="edward")))
        check("filtered hit", [r["entry_id"] for r in rows] == ["701"], str(rows))
        rows = json.loads(asyncio.run(server_mac.search_emails(recipient="edward")))
        check("recipient alone works", [r["entry_id"] for r in rows] == ["701"], str(rows))
        result = json.loads(asyncio.run(server_mac.search_emails()))
        check("no criteria -> error", "error" in result, str(result))
        # Filters need the database; the AppleScript fallback cannot honour them.
        server_mac.db = None
        result = json.loads(asyncio.run(server_mac.search_emails(query="approve", recipient="edward")))
        check("filters without database -> error, not a silent subject search", "error" in result and "recipient" in result["error"], str(result))
    server_mac.db = None
```

Add both to `main()`. Run: expected `TypeError` on the `recipient` keyword.

- [ ] **Step 3: Implement `search_messages`**

Replace the method:

```python
    def search_messages(self, folder_id: int, query: str, count: int,
                        recipient: str = "", sender: str = "") -> list[dict]:
        """Case-insensitive substring search.

        `query` matches subject, sender name, sender address and the preview
        (first 255 characters of the body). `recipient` matches the To and
        CC address lists and the display-To names. `sender` matches sender
        name and address. All given criteria must match.
        """
        clauses: list[str] = []
        params: list = [folder_id]
        query, recipient, sender = query.strip(), recipient.strip(), sender.strip()
        if query:
            p = _like_pattern(query)
            clauses.append(
                "(Message_NormalizedSubject LIKE ? ESCAPE '\\' OR "
                "Message_SenderList LIKE ? ESCAPE '\\' OR "
                "Message_SenderAddressList LIKE ? ESCAPE '\\' OR "
                "Message_Preview LIKE ? ESCAPE '\\')"
            )
            params += [p] * 4
        if recipient:
            p = _like_pattern(recipient)
            clauses.append(
                "(IFNULL(Message_ToRecipientAddressList, '') LIKE ? ESCAPE '\\' OR "
                "IFNULL(Message_CCRecipientAddressList, '') LIKE ? ESCAPE '\\' OR "
                "IFNULL(Message_DisplayTo, '') LIKE ? ESCAPE '\\')"
            )
            params += [p] * 3
        if sender:
            p = _like_pattern(sender)
            clauses.append(
                "(IFNULL(Message_SenderList, '') LIKE ? ESCAPE '\\' OR "
                "IFNULL(Message_SenderAddressList, '') LIKE ? ESCAPE '\\')"
            )
            params += [p] * 2
        if not clauses:
            return []
        params.append(max(0, int(count)))
        rows = self._query(
            f"SELECT {_MESSAGE_COLUMNS} FROM Mail "
            f"WHERE Record_FolderID = ? AND {_LIVE_ROWS} AND " + " AND ".join(clauses) +
            " ORDER BY Message_TimeReceived DESC, Record_RecordID DESC LIMIT ?",
            params,
        )
        return [self._row_to_summary(r) for r in rows]
```

- [ ] **Step 4: Implement `search_emails`**

Change the signature to

```python
async def search_emails(
    query: str = "",
    folder: str = "inbox",
    count: int = 10,
    recipient: str = "",
    sender: str = "",
) -> str:
```

Replace the docstring with:

```python
    """Search for emails in Outlook.

    On legacy Outlook for Mac the search runs against Outlook's local
    message index. `query` matches the subject, the sender name and
    address, and the preview, which is only the FIRST 255 CHARACTERS of
    the body. Full bodies are not indexed: a keyword that appears deeper
    in a message is not found. To find pending requests reliably, filter
    by `recipient` (and `sender`) and read candidates with read_email.
    Reply/forward prefixes ("Re:", "FW:") are ignored.

    If the index is unavailable, `query` falls back to AppleScript
    filtering on subject only, and `recipient`/`sender` return an error
    rather than silently searching without them.

    Args:
        query: Substring for subject, sender and preview. May be empty
            when recipient or sender is given.
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.
        recipient: Substring matched against To and CC addresses and the
            displayed To names, e.g. "edward" or "isdemand".
        sender: Substring matched against the sender name and address.

    Returns:
        JSON array of matching email summaries, newest first, or an error.
    """
```

Replace the start of the body (through the database attempt) with:

```python
    if not (query.strip() or recipient.strip() or sender.strip()):
        return json.dumps({"error": "Provide at least one of query, recipient, sender"})
    fid = await _db_folder_id(folder)
    if fid is not None:
        try:
            rows = await asyncio.to_thread(db.search_messages, fid, query, count, recipient, sender)
            return json.dumps(rows, indent=2, default=str)
        except OutlookDBError as e:
            logger.warning("Outlook database search failed; using AppleScript: %s", e)
    if recipient.strip() or sender.strip():
        return json.dumps({"error": "recipient and sender filters need Outlook's message index, "
                                    "which is unavailable right now; retry, or search with query only"})
```

The rest of the function (the AppleScript scripts) stays as it is.

- [ ] **Step 5: Run the gate**

Expected: all pass, including the existing `test_search_matches_subject_sender_preview`.

- [ ] **Step 6: README**

Change the `search_emails` row in the Email table to:
`| \`search_emails\` | yes | yes | Search by keyword; macOS also filters by \`recipient\` and \`sender\`. Keyword matching covers subject, sender and the first 255 characters of the body only |`.
In "#### Fast list and search on macOS" add a bullet:

> - `search_emails` keyword matching covers the subject, sender and the message preview (first 255 characters). Bodies are not indexed, so a keyword sweep can miss requests whose key word sits lower in the message. Combine it with `recipient="<your name or address>"` to list everything addressed to you, then `read_email` the candidates.

- [ ] **Step 7: Commit**

```bash
git add src/outlook_desktop_mcp/outlook_db.py src/outlook_desktop_mcp/server_mac.py tests/mac_db_test.py README.md
git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit -m "macOS: search_emails by recipient and sender; document preview limit"
```

---

### Task 5: `save_attachment` saves inline images

**Files:**
- Modify: `src/outlook_desktop_mcp/server_mac.py` (`save_attachment`)
- Modify: `tests/mac_batch_test.py`, `tests/mac_live_test.py`

**Interfaces:**
- Consumes: `bridge.run`, `DELIM`, `escape`.
- Produces: `save_attachment` returns `{"status": "saved", "filename", "path", "bytes"}` or an error string.

- [ ] **Step 1: Failing test**

Append to `tests/mac_batch_test.py`:

```python
def test_save_attachment_uses_posix_file_and_verifies_output():
    log("--- save_attachment saves through a POSIX file reference ---")
    with tempfile.TemporaryDirectory() as d:
        class Writes(FakeBridge):
            async def run(self, script, timeout=None):
                self.scripts.append(script)
                with open(os.path.join(d, "image003.png"), "wb") as fh:
                    fh.write(b"\x89PNG fake")
                return f"image003.png{DELIM}{d}/image003.png"

        fake = Writes()
        server_mac.bridge = fake
        result = json.loads(asyncio.run(server_mac.save_attachment(entry_id="1", attachment_index=1, save_directory=d)))
        script = fake.scripts[0]
        check("only one script", len(fake.scripts) == 1)
        check("save uses POSIX file", "save a in (POSIX file savePath)" in script, script)
        check("no string-path save", "save a in savePath\n" not in script)
        check("dead placeholder script gone", "__PLACEHOLDER__" not in script)
        check("saved status", result.get("status") == "saved", str(result))
        check("path returned", result.get("path") == os.path.join(d, "image003.png"), str(result))
        check("byte count returned", result.get("bytes") == 9, str(result))

        server_mac.bridge = FakeBridge(output=f"ghost.png{DELIM}{d}/ghost.png")
        result = asyncio.run(server_mac.save_attachment(entry_id="1", attachment_index=1, save_directory=d))
        check("missing file reported as error", result.startswith("Error saving attachment") and "ghost.png" in result, result)
```

Add to `main()`. Run: expected failures on the POSIX-file shape, placeholder, `bytes`, and the missing-file case.

- [ ] **Step 2: Implement**

Replace everything in `save_attachment` from `# Use POSIX path for AppleScript` through `return f"Error saving attachment: {e}"` with:

```python
    # `save ... in` takes a file object. A POSIX path *string* is rejected
    # with -2700 for inline images (which have no `file` of their own), so
    # build the reference with `POSIX file`.
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    if attCount < {attachment_index} then return "ERROR:Only " & attCount & " attachment(s)"
    set a to item {attachment_index} of attList
    set aname to name of a
    set savePath to "{escape(save_directory)}/" & aname
    save a in (POSIX file savePath)
    return aname & "{DELIM}" & savePath
end tell'''

    try:
        raw = await bridge.run(script)
        if raw.startswith("ERROR:"):
            return raw
        parts = raw.split(DELIM)
        filename = parts[0].strip() if parts else "unknown"
        save_path = os.path.join(save_directory, filename)
        if not os.path.isfile(save_path) or os.path.getsize(save_path) == 0:
            return (f"Error saving attachment: Outlook reported success but no file "
                    f"was written at {save_path}")
        return json.dumps({
            "status": "saved",
            "filename": filename,
            "path": save_path,
            "bytes": os.path.getsize(save_path),
        }, indent=2, default=str)
    except Exception as e:
        return f"Error saving attachment: {e}"
```

Update the docstring `Returns:` to `JSON with status, filename, path and bytes, or an error.` and note in the description that inline (pasted) images are supported.

- [ ] **Step 3: Run the gate**

Expected: all pass.

- [ ] **Step 4: Live check**

Add to `tests/mac_live_test.py` `run()`, at the end:

```python
    if INLINE_ID:
        log("--- inline image save ---")
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            raw = await server_mac.save_attachment(entry_id=INLINE_ID, attachment_index=1, save_directory=d)
            try:
                saved = json.loads(raw)
            except ValueError:
                saved = {"error": raw}
            check("inline attachment saved", saved.get("status") == "saved", raw)
            check("file has bytes", saved.get("bytes", 0) > 0, raw)
```

Run: `OUTLOOK_MCP_LIVE_SCRATCH=Edward.Adjei@mtn.com OUTLOOK_MCP_LIVE_INLINE_ID=197468 ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py`. Expected: all checks pass, including the two inline ones. (Message 197468 in the inbox carries `image003.png` inline.)

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/server_mac.py tests/mac_batch_test.py tests/mac_live_test.py
git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit -m "macOS: save inline attachments via POSIX file reference"
```
