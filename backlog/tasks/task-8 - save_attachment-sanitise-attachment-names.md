---
id: TASK-8
title: 'save_attachment: sanitise attachment names'
status: To Do
assignee: []
created_date: '2026-09-16 09:17'
labels:
  - macos
  - follow-up
dependencies: []
priority: low
ordinal: 8000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Follow-up from the final review (Task 5 minor 1). An attachment name containing '/' or '../' is used unsanitised in both the AppleScript path and the Python join, so a hostile name could write outside save_directory while the tool reports saved. Apply os.path.basename on both sides and reject empty names.
<!-- SECTION:DESCRIPTION:END -->
