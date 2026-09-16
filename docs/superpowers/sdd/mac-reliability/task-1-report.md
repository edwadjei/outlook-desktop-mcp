# Task 1 report: `reply_email` uses the dictionary's `reply to` command; live send/reply test

**Status: DONE_WITH_CONCERNS**

Concern (one deviation from the brief, detailed under "Deviations"): a pre-existing check in
`tests/mac_batch_test.py` asserted the invented `reply all to m` command and had to be corrected
to the dictionary form. No other deviation; all brief values used verbatim.

Commit: `f04de6c6b8e4f49fac3c08b0b817096af84d1b87`
Subject: `macOS: fix reply_email reply-all command; add live send/reply test`
Author: `edwadjei <edd.net49@gmail.com>`, no trailers.
Files: `src/outlook_desktop_mcp/server_mac.py`, `tests/mac_batch_test.py`, `tests/mac_live_test.py`
(staged by explicit path; this report is not staged).

## RED

Command:

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py 2>&1 | grep -E "FAIL|checks passed"
```

Output (excerpt):

```
  PASS: reply_all=False: reported success
  PASS: reply_all=False: no invented 'reply all to' command
  PASS: reply_all=False: reply command
  PASS: reply_all=False: script compiles against Outlook
  PASS: reply_all=True: reported success
  FAIL: reply_all=True: no invented 'reply all to' command
  FAIL: reply_all=True: reply command tell application "Microsoft Outlook"
  FAIL: reply_all=True: script compiles against Outlook /var/folders/64/.../s.applescript:4: error: Expected end of line, etc. but found identifier. (-2741)
123/126 checks passed
```

Why expected: exactly the three failures the brief predicts, all on the `reply_all=True` branch.
The old code emitted `set replyMsg to reply all to m without opening window`; `reply all to` is
not a command in Outlook's dictionary, and `osacompile` against the running Outlook confirms it
(syntax error at line 4). The `reply_all=False` branch already used `reply to`, so it passed.

## GREEN

Command (same as RED, after the fix in `reply_email`):

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py 2>&1 | grep -E "FAIL|checks passed"
```

Output:

```
  PASS: reply_all=False: reported success
  PASS: reply_all=False: no invented 'reply all to' command
  PASS: reply_all=False: reply command
  PASS: reply_all=False: script compiles against Outlook
  PASS: reply_all=True: reported success
  PASS: reply_all=True: no invented 'reply all to' command
  PASS: reply_all=True: reply command
  PASS: reply_all=True: script compiles against Outlook
126/126 checks passed
```

(The first GREEN run showed `125/126` with `FAIL: reply all command` from the pre-existing check
described under "Deviations"; after correcting that check the file is 126/126.)

Fix applied in `reply_email` (verbatim from the brief):

```python
    # Outlook's dictionary: `reply to <message>` with boolean parameters
    # `reply to all` and `opening window`. There is no `reply all to` command.
    reply_opts = ("with reply to all without opening window" if reply_all
                  else "without opening window")
```

and the script line `set replyMsg to reply to m {reply_opts}`.

## Gate

Command (from Global Constraints, run unpiped so exit codes are honoured):

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py && ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py && ~/.mcp-venvs/outlook-desktop/bin/pip wheel -q --no-deps -w /tmp/odm-wheel "$PWD" && rm -rf /tmp/odm-wheel
```

Output:

```
gate rc=0
tests/mac_batch_test.py: 126/126 checks passed   (118 before Task 1 + 8 new)
tests/mac_db_test.py:    68/68 checks passed
pip wheel: built, no errors (only the pip self-update notice on stderr)
```

Also confirmed the live test's opt-out path: running `tests/mac_live_test.py` without
`OUTLOOK_MCP_LIVE_SCRATCH` prints `SKIP: set OUTLOOK_MCP_LIVE_SCRATCH to run the live test` and
exits 0.

## Live test

Pre-check: `ps -axo pid,etime,command | grep "[o]sascript -e"` returned nothing (no other
AppleScript running). Outlook 16.93.2 running in legacy mode. Run once only.

Command:

```
OUTLOOK_MCP_LIVE_SCRATCH=Edward.Adjei@mtn.com ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py
```

Full output:

```
Wed Sep 16 08:28:58 GMT 2026
2026-09-16 08:28:58,907 [outlook_desktop_mcp.applescript_bridge] INFO: AppleScript bridge ready. Outlook version: 16.93.2
2026-09-16 08:28:59,871 [outlook_desktop_mcp] INFO: Outlook database trusted (inbox id 115, 33058 messages)
  database: trusted
--- seed message ---
  PASS: send_email reported success
  PASS: seed appears in Sent Items
  PASS: seed delivered to inbox
--- replies ---
  PASS: reply_all=False: reply_email reported success
  PASS: reply_all=False: reply appears in Sent Items
  PASS: reply_all=True: reply_email reported success
  PASS: reply_all=True: reply appears in Sent Items
==================================================
7/7 checks passed
live rc=0
Wed Sep 16 08:29:22 GMT 2026
```

7/7 on the first run; the whole round trip took 24 s. Three messages now sit in the test
mailbox (Inbox and Sent Items) with subject
`[MCP live test 2026-09-16 08:28:58] reply regression`, all marked "Safe to delete".

## Deviations from the brief

1. **Corrected a pre-existing check that asserted the bug.** `tests/mac_batch_test.py`,
   `test_reply_email_inserts_html_after_body_tag`, line 321, had
   `check("reply all command", "reply all to m" in script)`. The brief does not mention it, but
   it asserts the exact invented command this task removes, so it failed once the fix landed
   (first GREEN run: `125/126`, `FAIL: reply all command`). I changed it to
   `check("reply all command", "reply to m with reply to all" in script)`. This is not a weakening:
   the check still asserts that reply-all is requested, now in the form Outlook's dictionary
   accepts and that `osacompile` verifies. Cost if wrong: none beyond the new test already covering
   the same command shape more strictly.

2. No other deviation. Step order, test code, fix code, live test file, live send address, and
   commit subject/author are verbatim from the brief. The live test ran exactly once.

## Notes for the reviewer

- The new `_compiles()` helper compiles via `osacompile`, which resolves the `Microsoft Outlook`
  terminology from the running app; the compile checks are skipped (with a log line) when Outlook
  is not running, so the batch test still runs cleanly on a machine without Outlook.
- `tests/mac_live_test.py` exposes `wait_for_count(folder, subject, minimum)` and `check()` as
  the brief's interface for Tasks 3 and 5. `INLINE_ID` is read but unused in this task, as
  specified.
