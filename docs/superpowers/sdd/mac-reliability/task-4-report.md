# Task 4 report: `search_emails` by recipient and sender; document the preview limit

Status: DONE

Commit: `1f14d0a` — `macOS: search_emails by recipient and sender; document preview limit`
(branch `edw/mac-reliability-fixes`, base `8380a2f`; files: `src/outlook_desktop_mcp/outlook_db.py`,
`src/outlook_desktop_mcp/server_mac.py`, `tests/mac_db_test.py`, `README.md`)

## Pre-flight

- Baseline before any change: `tests/mac_batch_test.py` 132/132, `tests/mac_db_test.py` 107/107.
- Read-only sanity check of the real profile database (no write):
  `sqlite3 -readonly ".../Main Profile/Data/Outlook.sqlite" "SELECT name FROM pragma_table_info('Mail') WHERE name IN (...)"`
  returned all six columns the new SQL touches: `Message_ToRecipientAddressList`,
  `Message_CCRecipientAddressList`, `Message_DisplayTo`, `Message_SenderList`,
  `Message_SenderAddressList`, `Message_Preview`.
- No live AppleScript probe was needed for this task; `tests/mac_live_test.py` was not run.

## RED

Fixture widened to 14 columns (with `(None,) * (14 - len(row))` padding so the existing 11-field
tuples keep working), `MAIL` comment updated, the two brief tests appended and added to `main()`.

Command:

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py
```

Output (excerpt):

```
--- search filters by recipient and sender ---
Traceback (most recent call last):
  File ".../tests/mac_db_test.py", line 681, in <module>
    main()
  File ".../tests/mac_db_test.py", line 671, in main
    test_search_by_recipient_and_sender()
  File ".../tests/mac_db_test.py", line 611, in test_search_by_recipient_and_sender
    check("recipient matches to and cc", ids(db.search_messages(115, "", 10, recipient="edward")) == ["603", "601"])
TypeError: OutlookDB.search_messages() got an unexpected keyword argument 'recipient'
```

Why expected: the 3-argument `search_messages(folder_id, query, count)` does not accept the new
`recipient` keyword, which is exactly the failure the brief names for Step 2. All 107 pre-existing
checks passed before the traceback, confirming the fixture widening did not disturb them.

## GREEN

`search_messages` replaced with the brief's version (independent `query` / `recipient` / `sender`
LIKE clauses ANDed together, `IFNULL` on the nullable recipient and sender columns, stable
`ORDER BY Message_TimeReceived DESC, Record_RecordID DESC`). `search_emails` given the new
signature and docstring, the "at least one criterion" guard, the five-argument `to_thread` call,
and the "filters need the index" error before the AppleScript fallback.

Command:

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py
```

Output (excerpt):

```
--- search filters by recipient and sender ---
  PASS: recipient matches to and cc
  PASS: recipient matches display name
  PASS: query AND recipient
  PASS: sender filter
  PASS: all three combined
  PASS: nothing given -> empty
  PASS: rows without recipient columns filled do not match
--- search_emails passes recipient and sender to the database ---
  PASS: filtered hit
  PASS: recipient alone works
  PASS: no criteria -> error
  PASS: filters without database -> error, not a silent subject search
==================================================
118/118 checks passed
```

## Gate

Command (run after the README edit, so it reflects the committed tree):

```
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py && \
~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py && \
~/.mcp-venvs/outlook-desktop/bin/pip wheel -q --no-deps -w /tmp/odm-wheel "$PWD" && rm -rf /tmp/odm-wheel
```

Output:

```
132/132 checks passed
118/118 checks passed
GATE GREEN
```

(`tests/mac_batch_test.py` unchanged at 132; `tests/mac_db_test.py` 107 -> 118, +11 new checks.
`pip wheel` printed only its routine "new release of pip available" notice.)

## Deviations from the brief

- `test_server_search_emails_filters` calls `_reset_db_state()` at its start (before assigning
  `server_mac.db`) and again at the end, in place of the brief's trailing `server_mac.db = None`.
  This was requested in the dispatch so that no `_db_candidate` left by an earlier test can affect
  the "filters without database" check. Behaviourally identical for the assertions in the brief.
- The brief lists Step 5 (gate) before Step 6 (README). I made the README edit first and then ran
  the gate once, so that the gate result covers the exact tree that was committed. The db suite
  had already been run green (118/118) immediately after the implementation, before the README edit.

No other deviations. No stubs, no placeholders, no weakened tests. Windows code untouched.
