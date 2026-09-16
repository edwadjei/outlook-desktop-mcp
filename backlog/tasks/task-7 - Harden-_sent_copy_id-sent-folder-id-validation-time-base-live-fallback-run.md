---
id: TASK-7
title: >-
  Harden _sent_copy_id (sent folder, id validation, time base, live fallback
  run)
status: To Do
assignee: []
created_date: '2026-09-16 09:17'
labels:
  - macos
  - follow-up
dependencies: []
priority: medium
ordinal: 7000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Follow-up from the final review (Task 3 minors 2-4, N4). With a trusted database and an unresolvable sent folder the loop re-resolves for 10 s instead of falling back; the AppleScript fallback returns the id unvalidated (add an isdigit guard); find_sent_copy relies on Message_TimeReceived being local creation time (document or add slack); the AppleScript fallback has never been run live — run the live test once with OUTLOOK_MCP_DB_PATH=/nonexistent.
<!-- SECTION:DESCRIPTION:END -->
