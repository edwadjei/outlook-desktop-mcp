---
id: TASK-4
title: search_emails by recipient and sender; document preview limit
status: Done
assignee: []
created_date: '2026-09-16 08:26'
updated_date: '2026-09-16 09:06'
labels:
  - macos
  - enhancement
dependencies: []
priority: medium
ordinal: 4000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Keyword search only sees subject, sender and the first 255 characters of the body, so requests with the key word deeper in the body are missed. Add recipient and sender filters over the database's recipient columns and state the limit clearly.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 OutlookDB.search_messages accepts recipient and sender, ANDed with query; query may be empty
- [x] #2 search_emails exposes recipient and sender; with no criteria returns an error; filters without the database return an error instead of a silent subject search
- [x] #3 Docstring and README state the 255-character preview limit and the recommended recipient filter
- [x] #4 Gate green
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Commit 1f14d0a. Reviewer PASS/PASS. Gate 132/132, 118/118. Message 197678 found via recipient filter in 3.6 ms.
<!-- SECTION:FINAL_SUMMARY:END -->
