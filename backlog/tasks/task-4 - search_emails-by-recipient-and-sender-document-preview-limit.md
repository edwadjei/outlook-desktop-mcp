---
id: TASK-4
title: search_emails by recipient and sender; document preview limit
status: To Do
assignee: []
created_date: '2026-09-16 08:26'
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
- [ ] #1 OutlookDB.search_messages accepts recipient and sender, ANDed with query; query may be empty
- [ ] #2 search_emails exposes recipient and sender; with no criteria returns an error; filters without the database return an error instead of a silent subject search
- [ ] #3 Docstring and README state the 255-character preview limit and the recommended recipient filter
- [ ] #4 Gate green
<!-- AC:END -->
