---
id: TASK-1
title: 'reply_email: use dictionary reply-to command; live send/reply test'
status: To Do
assignee: []
created_date: '2026-09-16 08:26'
labels:
  - macos
  - bug
dependencies: []
priority: high
ordinal: 1000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Reply-all replies fail to compile because 'reply all to' is not an Outlook command. Generate 'reply to m with reply to all' per the dictionary and add an opt-in live test that sends a seed and two replies to a scratch address and asserts they land in Sent Items.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 reply_email with reply_all=True generates 'reply to m with reply to all without opening window' and reply_all=False generates 'reply to m without opening window'
- [ ] #2 Both generated scripts compile with osacompile when Outlook is running
- [ ] #3 tests/mac_live_test.py sends a seed and two replies and asserts 3 Sent Items copies; passes once live
- [ ] #4 Gate green: mac_batch_test, mac_db_test, wheel build
<!-- AC:END -->
