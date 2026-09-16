# Task 2 report: bounded database queries, deferred trust check, `ping` tool

Status: DONE_WITH_CONCERNS

Commit: `571ce18ec627d77d96f1f358b1252ff4252b62de`
`macOS: bound database queries, defer trust check, add ping tool` (author edwadjei <edd.net49@gmail.com>)

Files in the commit: `src/outlook_desktop_mcp/outlook_db.py`, `src/outlook_desktop_mcp/server_mac.py`,
`tests/mac_db_test.py`, `tests/mac_live_test.py`, `README.md`. (`tests/mac_batch_test.py` was staged as the
brief's command lists it but has no changes; see Deviations.)

## Baseline before any edit

```
$ ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py 2>&1 | tail -1
126/126 checks passed
$ ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py 2>&1 | tail -1
68/68 checks passed
```

`mcp` 1.29.1 is installed; `FastMCP.list_tools` is a coroutine (`inspect.iscoroutinefunction` -> True), so
the brief's `asyncio.run(server_mac.mcp.list_tools())` registration check was used unchanged.

## RED 1: query deadline (Steps 1-2)

Command: `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py 2>&1 | tail -3`

```
  File ".../tests/mac_db_test.py", line 410, in test_query_deadline_interrupts_long_query
    db = OutlookDB(build_fixture(d), timeout=0.2)
TypeError: OutlookDB.__init__() got an unexpected keyword argument 'timeout'
```

Why expected: the constructor had no `timeout` parameter and `DB_TIMEOUT` did not exist; this is the
failure the brief predicts in Step 2.

## GREEN 1 (Steps 3-4)

Command: `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py 2>&1 | grep -E "FAIL|checks passed"`

```
72/72 checks passed
```

(68 + 4: raised OutlookDBError with "exceeded" in the message, returned in < 2 s at a 0.2 s deadline, a normal
query on the same handle still works, `DB_TIMEOUT == 20.0`.)

## RED 2: deferred trust check and `ping` (Steps 5-6)

Command: `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py 2>&1 | grep -E "FAIL|Error|checks passed" | head`

```
  FAIL: AppleScript timeout -> undecided (None)
AttributeError: module 'outlook_desktop_mcp.server_mac' has no attribute 'STARTUP_TIMEOUT'
```

Why expected: `_trust_db` returned `False` on any exception (no `None` path) and `server_mac` did not import
`STARTUP_TIMEOUT`; `startup`, `_db_candidate`, `_db_state` and `ping` did not exist either, which the later
tests would have hit in turn.

## Intermediate failure after Step 7 (test fixed, see Deviations)

First run after implementing `server_mac.py` verbatim:

```
  File ".../src/outlook_desktop_mcp/server_mac.py", line 2269, in startup
    await bridge.start()
AttributeError: 'VersionBridge' object has no attribute 'start'
```

The test file's pre-existing `FakeBridge` had no `start()`; the brief's `startup()` (verbatim) calls
`bridge.start()`. Fixed in the test, not the server (details below).

## GREEN 2 (Step 7 complete)

Command: `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py 2>&1 | grep -E "FAIL|Error|checks passed"`

```
94/94 checks passed
```

Breakdown of the 26 new checks: 4 (deadline) + 2 (new trust cases: timeout -> `None`, probe uses
`STARTUP_TIMEOUT`) + 10 (`test_server_trust_check_is_deferred_to_first_use`) + 10 (`test_ping_tool`).
The brief's Step 8 note says "68 + 4 + 8 + 12 + trust cases"; the tests as written in the brief contain 10 and
10, so 94 is the correct total for the brief's own test code.

Live test still skips without the scratch variable (not run with it, per the dispatch):

```
$ env -u OUTLOOK_MCP_LIVE_SCRATCH ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py 2>&1 | tail -3
SKIP: set OUTLOOK_MCP_LIVE_SCRATCH to run the live test
exit=0
```

## Gate (Step 8), run on the final tree before the commit

Command (from the brief):
`~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py && ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py && ~/.mcp-venvs/outlook-desktop/bin/pip wheel -q --no-deps -w /tmp/odm-wheel "$PWD" && rm -rf /tmp/odm-wheel`

```
126/126 checks passed
94/94 checks passed
gate exit=0
```

Wheel built: `outlook_desktop_mcp-0.3.0-py3-none-any.whl` (verified on an earlier run of the same command;
the final gate run above returned exit 0 for the full chain).

## What changed

- `outlook_db.py`: `import threading`; `_env_seconds`; `DB_TIMEOUT` (env `OUTLOOK_MCP_DB_TIMEOUT`, default
  20); `_interrupt` timer callback; `OutlookDB(path, timeout=DB_TIMEOUT)`; `_query` arms a daemon
  `threading.Timer` per attempt, maps `sqlite3.OperationalError` containing "interrupted" to
  `OutlookDBError("... exceeded Ns")`, and always cancels the timer and closes the connection in `finally`.
- `server_mac.py`: `import time`; imports `STARTUP_TIMEOUT`; module state `_db_candidate`, `_db_state`,
  `_db_lock`, `_STARTED_AT`; new `_ensure_db()` (lock-guarded first-use trust check, `None` verdict leaves
  state `unchecked` for retry, `False` is final); `_db_folder_id` goes through `_ensure_db`; `_trust_db`
  returns `bool | None`, probes under `STARTUP_TIMEOUT`, returns `None` on "timed out"; new `ping` tool
  (TOOL 0) before `send_email`; `startup()` coroutine replaces the nested `_start`; `main()` calls it.
- `tests/mac_db_test.py`: `import time`; deadline test; trust-check test extended with the timeout cases;
  `_reset_db_state`; deferred-trust test; `ping` test; `FakeBridge.start()` (see Deviations); all wired
  into `main()`.
- `tests/mac_live_test.py`: `start_server()` now `await server_mac.startup(); await server_mac._ensure_db()`
  and logs `_db_state`; the two imports that only `start_server` used (`outlook_db`, `OutlookDB`) removed.
- `README.md`: `ping` row after `list_emails` in the Email tools table; "Every call is bounded." paragraph
  in the macOS AppleScript Bridge section (placed after the "Key differences from Windows" list, before the
  `outlook_db.py` subsection).

## Deviations from the brief and concerns

1. **Test fix (reason for DONE_WITH_CONCERNS).** The brief's `test_server_trust_check_is_deferred_to_first_use`
   drives `server_mac.startup()` with a `FakeBridge` subclass, and the brief's `startup()` calls
   `bridge.start()`, but the test file's existing `FakeBridge` only defines `run`/`run_lines`. The test's
   intent is explicit ("startup ran only the version probe", asserted through `vb.scripts`), so I added an
   `async def start(self)` to `FakeBridge` that mirrors the real `AppleScriptBridge.start()`: one call to
   `self.run('tell application "Microsoft Outlook" to get version', timeout=server_mac.STARTUP_TIMEOUT)`.
   The server code is exactly the brief's; only the fake gained the method the real bridge already has.
   No existing check changed meaning; the 68 baseline checks still pass.
2. **`tests/mac_batch_test.py` unchanged.** The brief lists it under Files and in the `git add` line but no
   step asks for a batch-test change (the new behaviour is covered by the db suite, which is where `ping`
   and the deferred trust check are tested). It was staged as instructed and has no diff; the 126 Task 1
   checks still pass.
3. **Dead imports removed from `tests/mac_live_test.py`.** After the brief's replacement of `start_server()`,
   `outlook_db` and `OutlookDB` were unreferenced in that file, so the import line was reduced to
   `from outlook_desktop_mcp import server_mac`. Not in the brief; harmless and keeps the live test clean.
4. **Check-count note.** Step 8 expects "68 + 4 + 8 + 12 + trust cases"; the brief's own test code yields
   68 + 4 + 2 + 10 + 10 = 94. No test was dropped or weakened.
5. **No live probe was run.** Nothing in this task needed osascript; the live test was executed only in its
   SKIP path.

Not in scope, left alone: `_sent_copy_id` (Task 3) does not exist; Windows code untouched; no push, no
branch, no PR.
