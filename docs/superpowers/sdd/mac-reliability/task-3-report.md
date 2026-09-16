# Task 3 report: Confirm the Sent Items copy after send and reply; stable sent ordering

Status: DONE

Branch `edw/mac-reliability-fixes`, base `a5f1826`, commit `b3541d0b71405d99ec005b66bde5a6f998c663ab`.

## RED

Tests added first (Step 1 in `tests/mac_db_test.py`: `test_list_messages_tie_breaks_on_record_id`,
`test_normalize_subject`, `test_find_sent_copy`, `test_server_send_email_confirms_sent_copy`;
Step 2 in `tests/mac_batch_test.py`: `test_send_email_looks_up_sent_copy_without_db`), all
registered in `main()`, and `server_mac._SENT_CONFIRM_TIMEOUT = 0.0` made the first line of
`main()` in both files.

Command:

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py
```

Failing output (excerpt):

```
--- equal timestamps are ordered by record id, newest first ---
  FAIL: newest record first among equals ['301', '302']
--- normalize_subject strips reply and forward prefixes ---
Traceback (most recent call last):
  File "/Users/edwadjei/projects/outlook-desktop-mcp/.claude/worktrees/agitated-yonath-70ffa9/tests/mac_db_test.py", line 632, in <module>
    main()
    ~~~~^^
  File "/Users/edwadjei/projects/outlook-desktop-mcp/.claude/worktrees/agitated-yonath-70ffa9/tests/mac_db_test.py", line 621, in main
    test_normalize_subject()
    ~~~~~~~~~~~~~~~~~~~~~~^^
  File "/Users/edwadjei/projects/outlook-desktop-mcp/.claude/worktrees/agitated-yonath-70ffa9/tests/mac_db_test.py", line 548, in test_normalize_subject
    check("RE:", outlook_db.normalize_subject("RE: budget") == "budget")
                 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^
AttributeError: module 'outlook_desktop_mcp.outlook_db' has no attribute 'normalize_subject'
```

Why expected: `list_messages` ordered by `Message_TimeReceived DESC` only, so equal timestamps
came back in insertion order (`['301', '302']`); `normalize_subject` did not exist yet, so the
suite aborted with `AttributeError` before reaching `find_sent_copy` and the confirmation checks.

Command:

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py
```

Failing output (excerpt):

```
--- send_email falls back to an AppleScript Sent Items lookup ---
  FAIL: send then lookup 1
  FAIL: confirmation carries id Email sent: 'Hi' to a@x.com
==================================================
126/128 checks passed
```

Why expected: `send_email` ran only the send script (1 script, no Sent Items lookup) and returned
the bare `Email sent: 'Hi' to a@x.com` with no `(Sent Items id N)` suffix. 126/128 = the 126
pre-existing checks passing, the 2 new ones failing (the 2 lookup-shape checks are skipped when
the script count is not 2).

## GREEN

Implemented per the brief: `outlook_db.py` gained `import re`, `_PREFIX_RE`,
`normalize_subject`, `OutlookDB.find_sent_copy`, and `list_messages` now orders by
`Message_TimeReceived DESC, Record_RecordID DESC`. `server_mac.py` gained
`_SENT_CONFIRM_TIMEOUT = 10.0`, `_SENT_CONFIRM_INTERVAL = 0.5`, `_sent_copy_id`,
`_confirmation_suffix`; `send_email` and `reply_email` record `since`, run the send, poll for the
Sent Items copy, and append the suffix; both `Returns:` docstrings and the `list_emails` `folder`
docstring were updated.

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py
```

```
--- equal timestamps are ordered by record id, newest first ---
  PASS: newest record first among equals
--- normalize_subject strips reply and forward prefixes ---
  PASS: RE:
  PASS: nested
  PASS: plain
  PASS: inner Re kept
--- find_sent_copy returns the newest matching sent row after `since` ---
  PASS: newest live row after since
  PASS: older rows ignored
  PASS: case-insensitive subject
  PASS: other folder ignored
--- send_email and reply_email report the Sent Items id ---
  PASS: send confirmation carries the Sent Items id
  PASS: reply confirmation carries the Sent Items id
  PASS: missing copy reported
  PASS: gives up after the confirm timeout
==================================================
107/107 checks passed
```

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py
```

```
--- send_email falls back to an AppleScript Sent Items lookup ---
  PASS: send then lookup
  PASS: lookup reads sent items
  PASS: lookup compares the subject
  PASS: confirmation carries id
==================================================
130/130 checks passed
```

## Gate

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py && ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py && ~/.mcp-venvs/outlook-desktop/bin/pip wheel -q --no-deps -w /tmp/odm-wheel "$PWD" && rm -rf /tmp/odm-wheel
```

Exit 0. Check counts:

```
130/130 checks passed
107/107 checks passed
```

(batch: 126 before this task + 4 new = 130; db: 94 before this task + 13 new = 107; wheel built.)

## Live test

Pre-check `ps -axo pid,etime,command | grep "[o]sascript -e"` returned nothing. Run exactly once.

```
OUTLOOK_MCP_LIVE_SCRATCH=Edward.Adjei@mtn.com ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py
```

Exit 0. Full output:

```
2026-09-16 08:46:40,616 [outlook_desktop_mcp] INFO: Starting Outlook Desktop MCP server (macOS)...
2026-09-16 08:46:40,666 [outlook_desktop_mcp.applescript_bridge] INFO: AppleScript bridge ready. Outlook version: 16.93.2
2026-09-16 08:46:40,720 [outlook_desktop_mcp] INFO: Outlook profile database found (trust check on first use): /Users/edwadjei/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite
2026-09-16 08:46:40,720 [outlook_desktop_mcp] INFO: AppleScript bridge ready. Starting MCP stdio transport...
2026-09-16 08:46:41,530 [outlook_desktop_mcp] INFO: Outlook database trusted (inbox id 115, 33064 messages)
2026-09-16 08:46:41,530 [outlook_desktop_mcp] INFO: Using Outlook profile database for list/search: /Users/edwadjei/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite
  database: trusted
  PASS: database trusted for the live test
--- seed message ---
  PASS: send_email reported success
  PASS: confirmation names the Sent Items id
  PASS: Sent Items id readable
  PASS: seed appears in Sent Items
  PASS: seed delivered to inbox
--- replies ---
  PASS: reply_all=False: reply_email reported success
  PASS: confirmation names the Sent Items id
  PASS: Sent Items id readable
  PASS: reply_all=False: reply appears in Sent Items
  PASS: reply_all=True: reply_email reported success
  PASS: confirmation names the Sent Items id
  PASS: Sent Items id readable
  PASS: reply_all=True: reply appears in Sent Items
==================================================
14/14 checks passed
```

All three sends (seed, reply_all=False, reply_all=True) produced a confirmation naming a Sent
Items id, and `read_email` on each id returned a message whose subject ends with the test
subject.

## Commit

```
b3541d0b71405d99ec005b66bde5a6f998c663ab macOS: confirm Sent Items copy after send and reply; stable sent ordering
```

Staged by explicit path: `src/outlook_desktop_mcp/outlook_db.py`,
`src/outlook_desktop_mcp/server_mac.py`, `tests/mac_db_test.py`, `tests/mac_batch_test.py`,
`tests/mac_live_test.py`, `README.md`. Author `edwadjei <edd.net49@gmail.com>`. No trailer, no
AI reference, per the brief. Nothing pushed, no branch created, no PR.

## Deviations and notes

- Ruling carried from the Task 1 review, applied as directed: `start_server()` in
  `tests/mac_live_test.py` now checks `server_mac._db_state == "trusted"` after `_ensure_db()`,
  and `run()` returns immediately after `start_server()` when the state is not `trusted`, so the
  test fails fast instead of polling `wait_for_count` for minutes on the AppleScript path.
- Step 6 shape (minor, structural only): the brief shows the Sent-Items-id checks inlined after
  each success check with a local `import re`. I put `import re` at module top and factored the
  two checks into one helper `check_sent_copy(result, subject)` called at the same three points
  (seed send, reply_all=False, reply_all=True). Check names, regex, and `read_email` assertion are
  verbatim from the brief; the live output above shows the three pairs.
- `len(fake.scripts) == 1` assertions: none of the existing send/reply tests assert a script
  count (they only read `fake.scripts[0]`), and the ones that do assert `== 1` are list/search
  tests, which do not run the lookup. So no existing assertion needed changing; `fake.scripts[0]`
  is still the send script everywhere.
- `_sent_copy_id` returns `None` (no fallback to AppleScript) when the database path raises
  `OutlookDBError`, exactly as specified; the user then sees the "copy not visible yet" suffix.
  Noting it only so the reviewer knows it is deliberate.
- The report file is not staged or committed.

## Fix round 1

Finding: on the AppleScript fallback path, `reply_email` passed the original message's subject to
`_sent_copy_id`, but the Sent Items copy of a reply carries `RE: <subject>`, so the lookup never
matched a first reply and waited the full 10 s. Ruling: return the subject actually sent.

Change: in `reply_email`'s AppleScript, `return msubject` became `return subject of replyMsg`;
the `set msubject to subject of m` line was dropped (nothing else in that script used it). The
`Returns:` docstring now says the confirmation names Outlook's reply subject. No `or subject of m
is "RE: ..."` clause was added. The DB path is unchanged (`normalize_subject` already strips the
prefix) and now receives the same real subject.

### RED

New check in `test_reply_email_uses_dictionary_reply_command` (both `reply_all` values):
`"return subject of replyMsg" in script and "return msubject" not in script`.

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py
```

```
  FAIL: reply_all=False: returns the subject actually sent yHtml
  FAIL: reply_all=True: returns the subject actually sent yHtml
130/132 checks passed
```

Expected: the script still ended with `return msubject`.

### GREEN

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py
```

```
  PASS: reply_all=False: returns the subject actually sent
  PASS: reply_all=True: returns the subject actually sent
132/132 checks passed
```

The two "script compiles against Outlook" checks for the reply script also passed (Outlook was
running), so `return subject of replyMsg` compiles against the Outlook dictionary. Existing
FakeBridge reply tests return a canned subject and needed no change.

### Gate

Exit 0.

```
132/132 checks passed
107/107 checks passed
```

(batch 130 + 2 new = 132; db unchanged at 107; wheel built.) Live test not re-run, per the ruling.

### Commit

```
4d5b9dd6bb2faea6b508e2f96d046b34cb39fb26 macOS: reply confirmation uses the subject actually sent
```

Staged by explicit path: `src/outlook_desktop_mcp/server_mac.py`, `tests/mac_batch_test.py`.
A concurrent, unrelated modification to `docs/superpowers/sdd/mac-reliability/progress.md`
(not mine) was left untouched and unstaged, as was this report.

## Fix round 2

Finding: reading `subject of replyMsg` after `send replyMsg` races the Outbox-to-Sent Items
transit, which deletes the outgoing record (a -1728 after a successful send would turn a sent
reply into an error response and invite a duplicate send). Ruling: read the subject into a
variable before sending.

Change: the final tell block of `reply_email`'s script is now exactly

```
tell application "Microsoft Outlook"
    set content of replyMsg to newContent
    set sentSubject to subject of replyMsg
    send replyMsg
    return sentSubject
end tell
```

Outlook assigns the Re: prefix when the reply is created, so the pre-send value equals the sent
value; `_sent_copy_id` still receives the real sent subject on both paths.

### RED

The fix-round-1 check in `test_reply_email_uses_dictionary_reply_command` was replaced by one
asserting `set sentSubject to subject of replyMsg` is present and appears before `send replyMsg`,
`return sentSubject` is present, and neither `return subject of replyMsg` nor `return msubject`
remains.

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py
```

```
  FAIL: reply_all=False: subject read into a variable before send    set newContent to origContent & replyHtml
  FAIL: reply_all=True: subject read into a variable before send    set newContent to origContent & replyHtml
130/132 checks passed
```

Expected: the script still ended with `send replyMsg` / `return subject of replyMsg`.

### GREEN

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py
```

```
  PASS: reply_all=False: subject read into a variable before send
  PASS: reply_all=False: script compiles against Outlook
  PASS: reply_all=True: subject read into a variable before send
  PASS: reply_all=True: script compiles against Outlook
132/132 checks passed
```

The "script compiles against Outlook" checks confirm the reordered block compiles against the
Outlook dictionary. FakeBridge reply tests needed no change.

### Gate

Exit 0.

```
132/132 checks passed
107/107 checks passed
```

(batch 132, db 107, wheel built.) Live test not re-run, per the ruling.

### Commit

```
b844fa694b7e34b987142705e1dba3a7927d029e macOS: read reply subject before send
```

Staged by explicit path: `src/outlook_desktop_mcp/server_mac.py`, `tests/mac_batch_test.py`.
The externally modified `progress.md`, the `task-3-*-diff.patch` files, and this report were left
unstaged.
