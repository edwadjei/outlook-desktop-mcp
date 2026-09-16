---
id: TASK-5
title: save_attachment saves inline images
status: Done
assignee: []
created_date: '2026-09-16 08:26'
updated_date: '2026-09-16 09:10'
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
- [x] #1 Generated script contains 'save a in (POSIX file savePath)' and the dead placeholder script is removed
- [x] #2 Result JSON has status, filename, path, bytes; a missing file is reported as an error
- [x] #3 Live test saves an inline image from a real message when OUTLOOK_MCP_LIVE_INLINE_ID is set
- [x] #4 Gate green
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Commit 3a76990. Reviewer PASS/PASS. Gate 140/140, 118/118. Live 16/16 incl. inline save from 197468.
<!-- SECTION:FINAL_SUMMARY:END -->
