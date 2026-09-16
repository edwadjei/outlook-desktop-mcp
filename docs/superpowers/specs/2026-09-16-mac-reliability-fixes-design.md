# macOS reliability fixes — design

Date: 2026-09-16. Scope: the macOS server (`src/outlook_desktop_mcp/server_mac.py`,
`applescript_bridge.py`, `outlook_db.py`) and its tests. Windows code is untouched.

Source: four defects and one enhancement reported from live use on 2026-09-15 (Outlook for Mac
16.93.2, legacy mode, ~33k inbox items), plus the evidence gathered on 2026-09-16 below.

## Evidence (established before design)

| Item | Finding |
| --- | --- |
| Bug 1 `reply_email` | The generated script fails to *compile* when `reply_all=True`: `reply all to m` is not a command. Outlook's dictionary (`Outlook.sdef`, line 1838) defines `reply to <message>` with boolean parameters `opening window` and `reply to all`. `reply to m with reply to all without opening window` compiles. Error offset 127:130 lands on the `all` token, matching the report. Nothing reaches Outlook, so no draft or send occurs. |
| Bug 2 `save_attachment` | `save a in "<posix string>"` raises -2700 for inline images (`file` property is `missing value`). `save a in POSIX file "<path>"` succeeds and writes a valid PNG. |
| Bug 3 `list_emails("sent")` | After `send`, the message exists for ~3 s as a record in the Outbox folder (special type 2), then reappears in Sent Items under a *new* record id with `Message_TimeReceived` set. A list within that window cannot see it. `search_emails` later finds the Sent Items record. |
| Bug 4 hangs | Server log 2026-09-15: startup trusted the database at 11:32:35; `search_emails`, `list_emails`, `search_emails` each produced no response for 1800 s from ~14:48; no warning was logged; SIGINT did not stop the process, SIGTERM did. All three calls are served by the profile database path. The AppleScript path already has a 120 s bound; the database path has none beyond sqlite's 2 s busy timeout. |
| Bug 4 startup | Log 2026-09-16 08:03: bridge ready at +0.4 s, then no "database trusted" line before the client's 30 s connect timeout. The trust check runs an AppleScript `count of messages` with the 120 s script timeout; Outlook serialises scripts across sessions, so a busy Outlook stalls startup past the client's window. The same count takes 0.2 s on an idle Outlook. |
| Enhancement | The database `Mail` table holds `Message_ToRecipientAddressList`, `Message_CCRecipientAddressList`, `Message_DisplayTo` (recipients) and `Message_Preview` (max 255 characters). Full bodies are not in the database. A recipient + keyword query over the inbox takes ~1 s. |

## Design

### 1. `reply_email` (Bug 1)

Generate `reply to m with reply to all without opening window` when `reply_all` is true and
`reply to m without opening window` otherwise. No other change to the reply script.

### 2. Bounded calls and liveness (Bug 4)

- **Database deadline.** Every `OutlookDB._query` runs under a deadline, default 20 s
  (`OUTLOOK_MCP_DB_TIMEOUT` seconds overrides). Implementation: a `threading.Timer` that calls
  `Connection.interrupt()`; the interrupted query raises `OutlookDBError("Outlook database query
  exceeded 20s")`. Callers already fall back to AppleScript on `OutlookDBError`.
- **Lazy, bounded trust check.** Startup performs only the bridge `get version` probe (10 s). The
  database candidate is located but its trust check runs on first use, with the AppleScript part
  under `STARTUP_TIMEOUT` (10 s). A *timeout* leaves the decision open (retried on the next call);
  a *mismatch* or *unusable database* caches "untrusted" for the process lifetime. Concurrent
  first calls share one check via an `asyncio.Lock`.
- **`ping` tool.** New tool `ping()` returning JSON: `outlook_version`, `applescript_ms`,
  `db` (`trusted` / `untrusted` / `unchecked` / `none`), `db_path`, `server_version`,
  `uptime_s`, `ok` (bool). Uses a 10 s AppleScript timeout and never raises; on failure `ok` is
  false and `error` holds the message.
- **Bound of every tool.** Each tool makes a finite number of bridge and database calls, so with
  both choke points bounded every tool returns within a finite time (worst case in the list/search
  chain: 20 s database + 120 s fast script + 120 s legacy script). Documented in the README.

### 3. Send confirmation and sent-folder ordering (Bug 3)

- `send_email` and `reply_email` wait up to 10 s (poll every 0.5 s) for the Sent Items copy and
  append ` (Sent Items id N)` to their confirmation string. Lookup uses the database when trusted
  (newest live row in the sent folder whose subject matches), else AppleScript over
  `messages 1 thru 20 of sent items`. If not found in 10 s the confirmation says
  `(Sent Items copy not visible yet; verify with search_emails)`.
- `OutlookDB.list_messages` orders by `Message_TimeReceived DESC, Record_RecordID DESC`.
- `list_emails` docstring and README state that a message spends a few seconds in the Outbox
  before it appears in Sent Items, and that the id returned by the send tools or `search_emails`
  is the reliable verification.

### 4. `search_emails` by recipient and sender (Enhancement)

- New optional parameters on the macOS `search_emails`: `recipient: str = ""` (substring over
  `Message_ToRecipientAddressList`, `Message_CCRecipientAddressList`, `Message_DisplayTo`) and
  `sender: str = ""` (substring over `Message_SenderList`, `Message_SenderAddressList`). All given
  filters are ANDed with `query`. `query` may be empty when a filter is given.
- `query` keeps matching subject, sender and preview. The docstring and README state that the
  preview is the first 255 characters of the body and that full bodies are not indexed, so a
  keyword sweep must be paired with a recipient filter or with `read_email` on candidates.
- The AppleScript fallback ignores `recipient` and `sender` and says so in a `note` field only
  when they were given.

### 5. `save_attachment` (Bug 2)

Use `save a in (POSIX file savePath)`. After the script returns, verify the file exists and is
non-empty; otherwise return an error naming the path. Remove the dead first script.

## Tests

- `tests/mac_batch_test.py` (FakeBridge, no Outlook): script-shape checks for reply-all, ping,
  save_attachment, send/reply confirmation polling; registration check that `ping` is a tool.
- `tests/mac_db_test.py`: ordering tie-break, recipient/sender filters, deadline interrupt.
- New `tests/mac_live_test.py` (opt-in): requires `OUTLOOK_MCP_LIVE_SCRATCH=<address>` and a
  running legacy-mode Outlook. Sends one message to the scratch address, waits for the inbox copy,
  replies with `reply_all=False` and again with `reply_all=True`, and asserts each reply appears in
  Sent Items (by the id in the confirmation, then by `search_emails`). Also saves an inline image
  when `OUTLOOK_MCP_LIVE_INLINE_ID` is set. Exits 0 when the environment variables are absent.

## Non-goals

Full-body search (bodies are not in the database), Windows parity for the new parameters,
returning attachment bytes inline (saving now works), changing the MCP client's idle timeout.
