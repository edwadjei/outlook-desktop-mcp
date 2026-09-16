---
id: TASK-6
title: 'Trust check: treat a busy or interrupted probe as undecided'
status: To Do
assignee: []
created_date: '2026-09-16 09:17'
labels:
  - macos
  - follow-up
dependencies: []
priority: medium
ordinal: 6000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Follow-up from the mac reliability delivery final review (N2 + Task 2 minor 3). _trust_db returns False when inbox_probe raises OutlookDBError for a transient busy/exceeded condition, and _ensure_db caches that as untrusted for the process lifetime; concurrent first calls under a busy Outlook each re-run the 10 s probe. Return None for busy/interrupted (keep False for schema/corrupt) and share one timeout outcome across queued callers.
<!-- SECTION:DESCRIPTION:END -->
