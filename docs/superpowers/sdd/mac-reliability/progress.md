# Ledger — macOS reliability fixes

Spec: `docs/superpowers/specs/2026-09-16-mac-reliability-fixes-design.md`
Plan: `docs/superpowers/plans/2026-09-16-mac-reliability-fixes.md`
Branch: `edw/mac-reliability-fixes` (from `edw/mac-read-email-recipients` @ a96510c, itself on `edw/mac-sqlite-index` @ d1bb405)
Remote: `fork` (github.com/edwadjei/outlook-desktop-mcp)

## Pre-flight

- 2026-09-16 08:00 Evidence gathered before design (see spec "Evidence"): dictionary check for `reply to`,
  osacompile reproduction of the reply-all syntax error, POSIX-file save probe on message 197468,
  self-send probe showing the Outbox→Sent Items transit (record 197800 → 197801), MCP client logs
  for 2026-09-15 (three 1800 s hangs on database-served tools, no server warning) and 2026-09-16
  08:03 (connect timeout while the trust check waited on AppleScript).
- Conflict scan: no other worktree edits `server_mac.py`, `outlook_db.py` or the mac tests. Two MCP
  server processes (pids 70460, 70653, this session's failed connects) and one from the 2026-09-15
  session (pid 88411, cwd agentic-business-analysis) are alive and idle; none hold scripts or the
  database. Left running.
- Board initialised (`backlog init`, integration mode none). Cards task-1 … task-5.

## Rulings

Ruling: proceed without the brainstorming approval gate — the owner asked to "fix all" and is not
present; the report is unusually precise, and every choice below is recorded here for reversal —
cost if wrong: rework of one task.

Ruling: `send_email`/`reply_email` wait up to 10 s for the Sent Items record and name its id — it
makes "reply sent" verifiable from the confirmation alone and removes the Outbox-window trap behind
Bug 3 — cost if wrong: up to 10 s extra latency per send, and a misleading "not visible yet" note
when Outlook is slow (search_emails still finds it).

Ruling: the trust check moves from startup to first database use with a 10 s AppleScript limit;
a timeout is retried on the next call, a mismatch is final — startup must never wait behind another
client's script (2026-09-16 08:03 connect timeout) — cost if wrong: the first list/search call after
start pays the 0.2 s check, and a persistently busy Outlook keeps list/search on the slow AppleScript
path until it frees up.

Ruling: database queries are bounded by `Connection.interrupt()` from a 20 s timer rather than by a
tool-level wrapper — both choke points (bridge, database) bounded makes every tool bounded by
construction, and a wrapper around 25 tool registrations risks the schema generation — cost if
wrong: a future tool that waits on something other than the bridge or the database would be
unbounded; none exists today.

Ruling: the 2026-09-15 1800 s hangs are treated as "unbounded database path" without a proven
mechanism — the process was killed before inspection, the surviving later process is idle, and the
only unbounded await on that path is the SQL call — cost if wrong: the hang recurs; the new `ping`
tool and per-query deadline will then localise it within 20 s instead of 30 min.

Ruling: `search_emails` filters (`recipient`, `sender`) require the database; the AppleScript
fallback returns an error rather than silently searching subject only — a silent downgrade is what
turned a pending request into "queue clear" — cost if wrong: callers on a non-database profile get
an error for filtered searches and must use `query` alone.

Ruling: full-body search is out of scope — bodies are not in the profile database and AppleScript
`content contains` over 33k messages runs for minutes — cost if wrong: a keyword deep in a body is
still missed; the docs say so and prescribe `recipient` + `read_email`.

Ruling: Bug 2 stays last (owner's later note "lowest priority") but is done — the fix is one line
and the live check is cheap — cost if wrong: none.

Ruling: unit test mains set `_SENT_CONFIRM_TIMEOUT = 0` — otherwise every send in the FakeBridge
suites waits the full 10 s — cost if wrong: none (an explicit timing test covers the wait).

## Task log

### Task 1 — reply_email dictionary command; live test
- BASE 096d540 → HEAD f04de6c. Implementer (fresh) DONE_WITH_CONCERNS: corrected a pre-existing
  check (tests/mac_batch_test.py:321) that asserted the invented `reply all to m`; accepted.
- Reviewer (fresh): SPEC PASS, CODE PASS. Gate verified independently 126/126, 68/68; live run
  verified in the profile database (Sent Items ids 197807, 197810, 197812).
- Deferred minors: (1) live test's Re:-insensitive matching exists only on the database path; on
  the AppleScript fallback the reply checks would false-FAIL — carry to Task 3/5 (assert the database
  is trusted up front or strip prefixes); (2) reply_all True/False indistinguishable in assertions;
  (3) a poll on the AppleScript fallback can overrun ROUND_TRIP by up to 243 s (bounded).
- Live test left three "Safe to delete" messages in the test mailbox (subject
  `[MCP live test 2026-09-16 08:28:59] reply regression`).

### Task 2 — bounded database queries, deferred trust check, ping
- BASE df9e88c → HEAD 571ce18. Implementer (fresh) DONE_WITH_CONCERNS: added `FakeBridge.start()`
  to the db test fake (mirrors the real bridge's single version probe); trimmed unused imports in
  the live test; batch test unchanged (nothing in the brief required it). Accepted.
- Reviewer (fresh): SPEC PASS, CODE PASS. Independently: gate 126/126, 94/94; live startup 0.10 s
  with state unchecked; ping ok 42 ms; first list call ran the trust check in 0.96 s → trusted;
  5 concurrent first calls → one probe; runaway query interrupted in 7 ms; 300 tight queries → no
  stray timer threads.
- Deferred minors: (1) busy back-off sleeps before the timer/connection cleanup; (2) list/search
  read the global `db` instead of the handle `_ensure_db` returned; (3) a burst of concurrent
  first calls under a busy Outlook each pay up to 10 s; (4) `_env_seconds` duplicates the bridge
  helper; (5) README's 260 s bound omits the ~6.6 s busy-retry path.

### Task 3 — Sent Items confirmation; stable sent ordering
- BASE a5f1826 → HEAD b3541d0. Implementer (fresh) DONE; live test 14/14 on one run (Sent Items
  rows 197820, 197823, 197825 verified by the reviewer in the database).
- Reviewer (fresh): SPEC PASS, CODE PASS with one Important finding: on the AppleScript fallback
  the reply lookup compares `subject of m` with the ORIGINAL subject, but a reply's Sent Items copy
  is "RE: <subject>", so on an untrusted-database profile every first reply waits 10 s and reports
  "not visible yet" (never a wrong id). Inherited from the brief's snippet.
- Ruling: fix in a Task 3 fix round by having the reply script return `subject of replyMsg` (the
  subject actually sent) instead of `subject of m`, so both lookup paths compare the real sent
  subject and the confirmation names it — one-line change, no new branch in the lookup — cost if
  wrong: the confirmation text shows "RE: subject" instead of the bare subject.
- Deferred minors: (2) with a trusted database and no resolvable sent folder the loop re-resolves
  for 10 s instead of falling back; (3) the fallback id is returned unvalidated (`isdigit` guard);
  (4) `find_sent_copy` relies on `Message_TimeReceived` being local creation time — comment it;
  (5) live test should log the matched id on PASS.
- Fix round 1 (4d5b9dd): reply script returns `subject of replyMsg`. Re-review: ADDRESSED, but new
  Important finding — reading the outgoing message after `send` races the Outbox transit that
  deletes its record (-1728 observed on a transited id), which would report an error for a reply
  that was sent.
- Ruling: fix round 2 reads the subject into `sentSubject` before `send` and returns it — Outlook
  sets the Re: prefix at creation, so the value is identical, and no post-send read remains —
  cost if wrong: none identified.
- Fix round 2 (b844fa6): subject read into `sentSubject` before `send`. Re-review: ADDRESSED, no
  new findings. Task 3 closed at b844fa6 (three commits: b3541d0, 4d5b9dd, b844fa6). Reply path
  changed after the live run; the close-out live run covers it.

### Task 4 — search_emails by recipient and sender
- BASE 8380a2f → HEAD 1f14d0a. Implementer (fresh) DONE; deviations: `_reset_db_state()` bracketing
  in the server test (requested at dispatch); gate run after the README edit.
- Reviewer (fresh): SPEC PASS, CODE PASS. Independently: gate 132/132, 118/118; real inbox
  (33,070 rows): `recipient="edward"` returns the previously missed request 197678 in 3.6 ms;
  `query="approv"` alone still misses it because "approv" is not in its 255-character preview,
  which is exactly the documented limit; LIKE escaping verified live with `%`, `_`, `\`.
- Deferred minors: (1) the "using AppleScript" warning is misleading when filters are given
  (an error is returned instead); (2) README now has two adjacent bullets on what search matches;
  (3) query clause uses bare sender columns while the sender clause wraps them in IFNULL.

### Task 5 — save_attachment saves inline images
- BASE 52cf125 → HEAD 3a76990. Implementer (fresh) DONE; live run 16/16 including the inline save
  from message 197468. The brief's predicted RED failure on the placeholder check did not occur
  because the dead script never reached the bridge; the check stays as a regression guard.
- Reviewer (fresh): SPEC PASS, CODE PASS. Independently: gate 140/140, 118/118; real save of
  attachment 2 of 197468 → valid 1057×391 PNG, 27,315 bytes.
- Deferred minors: (1) an attachment name containing "/" or "../" is used unsanitised in both the
  AppleScript path and the Python join; (2) `getsize` called twice; (3) unreachable `"unknown"`
  branch.

## Close-out

### Final whole-branch review (fresh reviewer, a96510c..7092c33)
- Spec coverage complete; one deliberate deviation from spec §4 text (filters return an error on the
  fallback instead of a `note`), recorded as a ruling; spec text to be amended in a follow-up.
- Cross-task seams checked: `_sent_copy_id` over the lazy trust check is safe (nothing escapes
  `_trust_db`); ping reflects state correctly; test isolation holds; live probe: startup 0.13 s,
  ping 40 ms, first list 1.2 s → trusted, filtered search 3 ms.
- Commit hygiene: 18 commits, all edwadjei <edd.net49@gmail.com>, no AI references, no trailers.
- Must-fix: README still says the trust check runs "at startup"; search table row reads as
  cross-platform; two overlapping search bullets. Doc-only.
- Triage of deferred minors: accept — T1(1)(2)(3), T2(1)(2)(4)(5), T3(5), T4(1)(2)(3), T5(2)(3).
  Follow-up — T2(3) serialised 10 s re-probes under a busy Outlook; T3(2)(3)(4) `_sent_copy_id`
  hardening (unresolvable sent folder, unvalidated fallback id, local-time assumption on
  `Message_TimeReceived`); T5(1) attachment-name sanitising; N2 a transient busy/interrupted
  probe is cached as untrusted for the process lifetime; N3 amend spec §4 text; N4 the AppleScript
  sent-copy fallback has never run live.
- Rulings: reviewer would reverse none; recommends amending "mismatch or unusable is final" so a
  busy/interrupted probe returns None (folded into follow-up N2).
- Ruling: untracked task-*-diff.patch files are deleted at close-out — they are reproducible from
  the commit ranges recorded in this ledger — cost if wrong: none.
- Verdict: READY once the README fix lands.

### Fix wave and final gate
- Fix wave (92ca39c, README only): re-review ADDRESSED ×3, no content lost. Must-fix list empty.
- Final gate on 92ca39c: tests/mac_batch_test.py 140/140 (118 at start), tests/mac_db_test.py
  118/118 (68 at start), wheel builds. Live end-to-end run 16/16 (seed, reply, reply-all, each
  confirmed by Sent Items id and read_email; inline image saved from 197468).
- Follow-up cards created: task-6 (trust-check resilience), task-7 (_sent_copy_id hardening),
  task-8 (attachment-name sanitising), task-9 (spec §4 amendment).
- Diff packages deleted (reproducible from the ranges above). Package reinstalled into
  ~/.mcp-venvs/outlook-desktop from this worktree at 92ca39c.
- Integration decision left to the owner: merge locally / push and open a PR / keep the branch.
