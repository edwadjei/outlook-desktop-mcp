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
