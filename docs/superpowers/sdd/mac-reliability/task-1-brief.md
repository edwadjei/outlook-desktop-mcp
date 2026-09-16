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
