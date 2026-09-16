---
id: TASK-3
title: Confirm Sent Items copy after send and reply; stable sent ordering
status: Done
assignee: []
created_date: '2026-09-16 08:26'
updated_date: '2026-09-16 09:00'
labels:
  - macos
  - bug
dependencies: []
priority: medium
ordinal: 3000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
A sent message sits in the Outbox for a few seconds before Outlook writes a new Sent Items record with a new id, so list_emails('sent') right after a send misses it. send_email and reply_email wait up to 10 s for that record and name its id; list ordering tie-breaks on record id; docs explain the window.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 send_email and reply_email confirmations end with '(Sent Items id N)' when the copy is found within 10 s, else the documented fallback text
- [x] #2 OutlookDB.find_sent_copy finds the newest live row with the normalized subject received at or after since
- [x] #3 list_messages orders by time received then record id, newest first
- [x] #4 Live test verifies each confirmation id is readable via read_email
- [x] #5 README and list_emails docstring document the Outbox window
- [x] #6 Gate green
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Commits b3541d0, 4d5b9dd, b844fa6. Reviewer PASS/PASS after two fix rounds. Gate 132/132, 107/107. Live test 14/14.
<!-- SECTION:FINAL_SUMMARY:END -->
