---
id: TASK-5
title: save_attachment saves inline images
status: To Do
assignee: []
created_date: '2026-09-16 08:26'
labels:
  - macos
  - bug
dependencies: []
priority: low
ordinal: 5000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Saving an inline (pasted) image fails with -2700 because the path is passed as a string; Outlook's save command takes a file object. Use POSIX file and verify the file was written.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Generated script contains 'save a in (POSIX file savePath)' and the dead placeholder script is removed
- [ ] #2 Result JSON has status, filename, path, bytes; a missing file is reported as an error
- [ ] #3 Live test saves an inline image from a real message when OUTLOOK_MCP_LIVE_INLINE_ID is set
- [ ] #4 Gate green
<!-- AC:END -->
