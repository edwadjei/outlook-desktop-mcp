# Standing rules for this task (read first)

You are the implementer for exactly one task of the delivery plan. This brief is your
requirements; use its values verbatim. Do not read other tasks' briefs or the whole plan.

- Work only in the repo checkout you were pointed at; never `cd` elsewhere; never touch git config.
- Test-driven: write the failing test first, run it and record the failing output (RED), then the
  minimal code, then the passing output (GREEN). Both go in your report.
- The gate (below) must be green before the commit. No stubs, no placeholders, no mock data in
  shipped code. Never weaken a test to make it pass; if a test in the brief is wrong, say so in the
  report under DONE_WITH_CONCERNS and fix the test to what the brief clearly intends.
- One commit for the task, exactly the subject given, staged by explicit path. No AI references,
  no Co-Authored-By trailer. Never push, never merge, never open a PR.
- Live probes: check `ps -axo pid,etime,command | grep "[o]sascript -e"` first; wait for other
  scripts to finish; never kill them. Outlook must be running in legacy mode.
- Write the report file named in your dispatch with: status (DONE | DONE_WITH_CONCERNS |
  NEEDS_CONTEXT | BLOCKED), RED and GREEN command+output, the gate output, the commit hash, and
  any concern or deviation. Then return a three-line status.

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
