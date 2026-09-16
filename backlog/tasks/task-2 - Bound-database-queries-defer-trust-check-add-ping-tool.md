---
id: TASK-2
title: 'Bound database queries, defer trust check, add ping tool'
status: Done
assignee: []
created_date: '2026-09-16 08:26'
updated_date: '2026-09-16 08:43'
labels:
  - macos
  - bug
dependencies: []
priority: high
ordinal: 2000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Database-served list/search calls hung for 30 minutes with no error; startup stalled past the client's 30 s connect timeout while Outlook served another client's script. Interrupt any SQL query past 20 s, run the trust check on first use with a 10 s AppleScript limit, and add a ping tool.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 OutlookDB queries past OUTLOOK_MCP_DB_TIMEOUT (default 20 s) raise OutlookDBError within the deadline
- [x] #2 startup() runs only the version probe; the trust check runs on first database use; a timeout is retried, a mismatch is cached
- [x] #3 ping tool returns ok, outlook_version, applescript_ms, db state, server_version, uptime_s within 10 s and never raises
- [x] #4 README documents the bounds and the ping tool
- [x] #5 Gate green
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Commit 571ce18. Reviewer PASS/PASS. Gate 126/126, 94/94. Live: startup 0.10 s, ping 42 ms, trust on first use 0.96 s.
<!-- SECTION:FINAL_SUMMARY:END -->
