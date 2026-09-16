---
id: TASK-2
title: 'Bound database queries, defer trust check, add ping tool'
status: To Do
assignee: []
created_date: '2026-09-16 08:26'
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
- [ ] #1 OutlookDB queries past OUTLOOK_MCP_DB_TIMEOUT (default 20 s) raise OutlookDBError within the deadline
- [ ] #2 startup() runs only the version probe; the trust check runs on first database use; a timeout is retried, a mismatch is cached
- [ ] #3 ping tool returns ok, outlook_version, applescript_ms, db state, server_version, uptime_s within 10 s and never raises
- [ ] #4 README documents the bounds and the ping tool
- [ ] #5 Gate green
<!-- AC:END -->
