# Task 5 report: `save_attachment` saves inline images

Status: DONE

Branch `edw/mac-reliability-fixes`, base `52cf125`, commit `3a769906369f02d484d0e2fe947d05bc40bb63ff`.

## RED

Command:

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py
```

Failing output (excerpt, exit 1):

```
--- save_attachment saves through a POSIX file reference ---
  PASS: only one script
  FAIL: save uses POSIX file tell application "Microsoft Outlook"
    set m to message id 1
    set attList to attachments of m
    set attCount to count of attList
    if attCount < 1 then return "ERROR:Only " & attCount & " attachment(s)"
    set a to item 1 of attList
    set aname to name of a
    set savePath to "/var/folders/64/2pcwgjsj2td71drk_sdhgr6c0000gn/T/tmpoc7l20rp/" & aname
    save a in savePath
    return aname & "|||" & savePath
end tell
  FAIL: no string-path save 
  PASS: dead placeholder script gone
  PASS: saved status
  PASS: path returned
  FAIL: byte count returned {'status': 'saved', 'filename': 'image003.png', 'path': '/var/folders/64/2pcwgjsj2td71drk_sdhgr6c0000gn/T/tmpoc7l20rp/image003.png'}
  FAIL: missing file reported as error {
  "status": "saved",
  "filename": "ghost.png",
  "path": "/var/folders/64/2pcwgjsj2td71drk_sdhgr6c0000gn/T/tmpoc7l20rp/ghost.png"
}
==================================================
136/140 checks passed
```

Why expected: the old `save_attachment` shipped `save a in savePath` with a plain
string path (rejected with -2700 for inline images), returned no `bytes`, and
trusted Outlook's success without checking the file exists. The four failures
are exactly those three properties plus the missing-file case. Note: the
`dead placeholder script gone` check already passed at RED, because the first
(dead) script containing `__PLACEHOLDER__` was overwritten in Python before it
ever reached `bridge.run`; the test still guards against it coming back.

## GREEN

Command:

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py
```

Output (excerpt, exit 0):

```
--- save_attachment saves through a POSIX file reference ---
  PASS: only one script
  PASS: save uses POSIX file
  PASS: no string-path save
  PASS: dead placeholder script gone
  PASS: saved status
  PASS: path returned
  PASS: byte count returned
  PASS: missing file reported as error
==================================================
140/140 checks passed
```

## Gate

Command:

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py && ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py && ~/.mcp-venvs/outlook-desktop/bin/pip wheel -q --no-deps -w /tmp/odm-wheel "$PWD" && rm -rf /tmp/odm-wheel
```

Exit 0. Check counts:

```
140/140 checks passed
118/118 checks passed
```

(132 -> 140 in `mac_batch_test.py`, the 8 new checks from this task; `mac_db_test.py` unchanged at 118; wheel built.)

## Live check

Pre-check: `ps -axo pid,etime,command | grep "[o]sascript -e"` returned nothing;
Outlook 16.93.2 running. Run once (attempt 1 of the 2 allowed).

Command:

```
OUTLOOK_MCP_LIVE_SCRATCH=Edward.Adjei@mtn.com OUTLOOK_MCP_LIVE_INLINE_ID=197468 ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py
```

Full output (exit 0):

```
2026-09-16 09:07:41,361 [outlook_desktop_mcp] INFO: Starting Outlook Desktop MCP server (macOS)...
2026-09-16 09:07:41,403 [outlook_desktop_mcp.applescript_bridge] INFO: AppleScript bridge ready. Outlook version: 16.93.2
2026-09-16 09:07:41,460 [outlook_desktop_mcp] INFO: Outlook profile database found (trust check on first use): /Users/edwadjei/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite
2026-09-16 09:07:41,460 [outlook_desktop_mcp] INFO: AppleScript bridge ready. Starting MCP stdio transport...
2026-09-16 09:07:42,082 [outlook_desktop_mcp] INFO: Outlook database trusted (inbox id 115, 33070 messages)
2026-09-16 09:07:42,082 [outlook_desktop_mcp] INFO: Using Outlook profile database for list/search: /Users/edwadjei/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite
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
--- inline image save ---
  PASS: inline attachment saved
  PASS: file has bytes
==================================================
16/16 checks passed
```

## Commit

`3a769906369f02d484d0e2fe947d05bc40bb63ff` — `macOS: save inline attachments via POSIX file reference`

Staged by explicit path: `src/outlook_desktop_mcp/server_mac.py`,
`tests/mac_batch_test.py`, `tests/mac_live_test.py`. This report is not staged.

## Deviations from the brief

None in code or tests; the replacement block, the new test and the live block
are verbatim from the brief. Two observations, neither a deviation:

- The brief predicted a RED failure on the placeholder check; it passed at RED
  for the reason given above. The check is kept as a regression guard.
- Docstring wording for the inline-image note was my own sentence, as the brief
  gave only the intent ("note ... that inline (pasted) images are supported").
