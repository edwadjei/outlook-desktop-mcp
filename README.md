# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)]()

**Turn your running Outlook Desktop into an MCP server.** No Microsoft Graph API, no Entra app registration, no OAuth tokens — just your local Outlook and the authentication you already have.

Any MCP client (Claude Code, Claude Desktop, etc.) can then send emails, manage your calendar, create tasks, handle attachments, and more — all through your existing Outlook session.

## Quick Start

**1. Install** (requires Python 3.12+):

```bash
pip install outlook-desktop-mcp
```

**2. Register with Claude Code:**

```bash
claude mcp add outlook-desktop -- outlook-desktop-mcp
```

**3. Open Outlook and start a Claude Code session.** That's it — tools are available immediately.

## How It Works — Platform Routing

When the server starts, it checks which operating system it is running on and takes one of two paths:

```
                        outlook-desktop-mcp starts
                                  |
                          sys.platform check
                         /                  \
                   "win32"                "darwin"
                      |                      |
              ┌───────┴────────┐    ┌────────┴────────┐
              │  server.py     │    │  server_mac.py   │
              │  COM Bridge    │    │  AppleScript     │
              │  (29 tools)    │    │  Bridge          │
              │                │    │  (22 tools)      │
              └───────┬────────┘    └────────┬─────────┘
                      |                      |
              OUTLOOK.EXE via         Microsoft Outlook
              COM / STA thread        via osascript
                      |                      |
              Exchange / M365         Exchange / M365
```

**Both paths use your locally running Outlook app and its existing authenticated session.** No cloud credentials, no Graph API tokens — the server inherits whatever account Outlook is signed into.

### Why two paths?

Windows Outlook (Classic) exposes a rich COM automation interface — the Outlook Object Model (`MSOUTL.OLB`). This has been the standard way to programmatically control Outlook on Windows for over 20 years. It provides deep access to mail rules, categories, MAPI properties, and the full folder hierarchy.

Mac Outlook does not support COM. Instead, it exposes an AppleScript dictionary that can be driven via the `osascript` command. The AppleScript interface covers the core operations — email, calendar, tasks — but does not expose rules, categories, or certain advanced MAPI features. This is a limitation of what Microsoft chose to include in Outlook for Mac's scripting dictionary, not a limitation of this project.

The server is structured as two parallel implementations with identical tool names and signatures, so MCP clients see the same interface regardless of platform. Tools that are not available on a given platform are simply not registered.

## Requirements

### Windows

- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office. The new "modern" Outlook (`olk.exe`) does **not** support COM
- **Python 3.12+** (x64 or ARM64)
- **Outlook must be running** when the MCP server starts

Both x64 and ARM64 Windows are supported. On ARM64, all dependencies (`pywin32`, `mcp`, `pydantic-core`, `cryptography`, `cffi`, `rpds-py`) have prebuilt `win_arm64` wheels — see the [ARM64 install notes](#arm64-windows) below for the one extra `pip` flag you need.

#### Outlook "Programmatic Access" security prompts

When the MCP server first touches Outlook's COM API, Outlook may show a dialog: *"A program is trying to access email address information stored in Outlook"* (the Object Model Guard / Programmatic Access prompt). This is Outlook's protection against malicious automation.

You have three options:

1. **Click "Allow access for 10 minutes"** every time you start a session. Fine for casual use.
2. **Get the antivirus status to "Valid"** in *File > Options > Trust Center > Trust Center Settings > Programmatic Access*. When that line reads `Valid`, the "Never warn me about suspicious activity" radio becomes selectable and the prompts go away. On most personal machines with current Defender this works out of the box.
3. **Apply the registry policy** in [`docs/suppress-outlook-oom-prompts.reg`](docs/suppress-outlook-oom-prompts.reg). From an **elevated** PowerShell or Command Prompt (Win+X → *Terminal (Admin)*), run:

   ```powershell
   reg import "C:\path\to\outlook-desktop-mcp\docs\suppress-outlook-oom-prompts.reg"
   ```

   Then fully quit Outlook (check Task Manager for stray `OUTLOOK.EXE` processes) and reopen it. This writes `AdminSecurityMode=3` and approves all `PromptOOM*` categories under `HKLM\Software\Policies\Microsoft\Office\16.0\Outlook\Security`, which Outlook honors regardless of AV status. On Intune/MDM-managed corporate devices, `HKCU\Software\Policies\...\Outlook` is locked and the HKLM keys may be overwritten on next policy sync — if `reg import` fails or the prompts come back, ask IT to push the equivalent settings via Group Policy.

> **Windows 11 ARM64 note:** Defender does not register with Outlook's `IOfficeAntiVirus` interface on ARM64, so Trust Center shows *"Antivirus status: Invalid"* and the *"Never warn me about suspicious activity"* radio stays greyed out **even when Outlook is launched as Administrator**. Option 2 is unavailable on ARM64; the `reg import` from option 3 is the only durable suppression path.

### macOS

- **Microsoft Outlook for Mac** — version 16.x or later
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

#### Required macOS permissions

The first time a tool runs, macOS will show **two permission prompts** that you must approve:

1. **Privacy & Automation** — a system dialog asks: *"python3.12 wants to control Microsoft Outlook"*. Click **Allow** to let the server send AppleScript commands to Outlook.

2. **Accessibility** — to read your Exchange/M365 inbox, the server uses macOS UI scripting (System Events). This requires Accessibility access for `python3.12`:
   - Open **System Settings > Privacy & Security > Accessibility**
   - Find **python3.12** in the list (it appears after the first prompt)
   - Toggle it **on**

   Without Accessibility enabled, calendar, tasks, and local folder tools will work, but listing Exchange inbox messages will return empty results.

Both permissions are one-time setup — macOS remembers them for future sessions.

## Available Tools by Platform

### Email

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `send_email` | yes | yes | Send an email with To/CC/BCC; `html_body` recommended, plain `body` auto-converted to HTML |
| `list_emails` | yes | yes | List recent emails from any folder, with optional unread filter |
| `ping` | no | yes | Liveness check: Outlook version, AppleScript round-trip time, database state; answers within 10 s |
| `read_email` | yes | yes | Read full email content by entry ID or subject search |
| `search_emails` | yes | yes | Full-text search across email subjects and bodies (macOS: subject, sender, and preview via the local message index) |
| `reply_email` | yes | yes | Reply or reply-all, preserving the conversation thread; accepts `html_body` |
| `mark_as_read` | yes | yes | Mark a specific email as read |
| `mark_as_unread` | yes | yes | Mark a specific email as unread |
| `move_email` | yes | yes | Move an email to Archive, Trash, or any folder |
| `list_folders` | yes | yes | Browse the folder hierarchy with item counts |

### Calendar

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_events` | yes | yes | List upcoming events within a date range |
| `get_event` | yes | yes | Read full event details by entry ID |
| `create_event` | yes | yes | Create a personal calendar appointment |
| `create_meeting` | yes | yes | Create a meeting and send invitations to attendees |
| `update_event` | yes | yes | Modify an existing event's subject, time, location, etc. |
| `delete_event` | yes | yes | Delete an appointment or cancel a meeting |
| `respond_to_meeting` | yes | — | Accept, decline, or tentatively accept a meeting invite |
| `search_events` | yes | yes | Search calendar events by keyword within a date range |

### Tasks

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_tasks` | yes | yes | List pending or completed tasks, sorted by due date |
| `get_task` | yes | yes | Read full task details including body and completion status |
| `create_task` | yes | yes | Create a new task with subject, due date, importance |
| `complete_task` | yes | yes | Mark a task as complete |
| `delete_task` | yes | yes | Remove a task |

### Attachments

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_attachments` | yes | yes | List all attachments on an email or calendar event |
| `save_attachment` | yes | yes | Download an attachment to a local directory |

### Categories, Rules, Out of Office (Windows only)

These tools rely on COM-specific APIs (MAPI property accessors, the Rules object model, and the Categories collection) that Outlook for Mac does not expose through AppleScript.

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_categories` | yes | — | List all available color categories in Outlook |
| `set_category` | yes | — | Set or clear categories on any email, event, or task |
| `list_rules` | yes | — | List all mail rules with enabled/disabled status |
| `toggle_rule` | yes | — | Enable or disable a mail rule by name |
| `get_out_of_office` | yes | — | Check whether Out of Office auto-reply is on or off |

**Total: 29 tools on Windows, 22 tools on macOS.**

## Architecture Details

### Windows: COM Bridge (`com_bridge.py`)

All Outlook COM operations run on a dedicated thread using the Single-Threaded Apartment (STA) model, as required by COM. The async MCP event loop dispatches tool calls to this thread via a queue and awaits results, keeping COM threading rules respected and the MCP protocol non-blocking.

```
MCP tool call (async)
  → bridge.call(func, args)
    → queued to STA thread
      → func(outlook, namespace, args) executes on COM thread
    → result returned via threading.Event
  → JSON response back to MCP client
```

Each tool's inner function receives the live `Outlook.Application` and `MAPI.Namespace` COM objects and works directly with the Outlook Object Model — `GetItemFromID`, `CreateItem`, `Items.Restrict` with DASL filters, and so on.

### macOS: AppleScript Bridge (`applescript_bridge.py`)

Each tool call builds an AppleScript string and executes it as a subprocess via `osascript`. There is no persistent connection — every call is stateless.

```
MCP tool call (async)
  → build AppleScript string
    → asyncio.create_subprocess_exec("osascript", "-e", script)
    → parse stdout text into structured data
  → JSON response back to MCP client
```

Each tool constructs a single AppleScript that fetches all needed data in one `osascript` call (no per-message subprocess loops). Results come back as delimited text, which the server parses into the same JSON structure the Windows server produces.

**Key differences from Windows:**

- Entry IDs on macOS are **numeric** (e.g. `42`), not hex strings. They identify items within their folder context.
- Folder references use AppleScript's **locale-independent keywords** (`inbox`, `sent items`, `drafts`, `deleted items`) rather than localized folder names.
- Search uses AppleScript's `whose` clause (e.g. `messages whose subject contains "query"`) instead of DASL filters, unless the local message index is available (see below).
- User input is escaped for safe embedding in AppleScript strings to prevent script injection.
- Dates passed to AppleScript are built by component assignment (`set year of d to 2026`, ...) rather than `date "..."` literals, which osascript parses according to the system locale and can silently turn `"2026-09-01 07:00"` into a date in 2007.

**Every call is bounded.** Each AppleScript runs under `OUTLOOK_MCP_SCRIPT_TIMEOUT` (default 120 s) and each profile-database query under `OUTLOOK_MCP_DB_TIMEOUT` (default 20 s); a query past its deadline is interrupted and the tool falls back to AppleScript. A list or search call therefore never runs longer than about 260 s (database, batched script, legacy script), and most finish in well under a second. The database trust check runs on the first list or search call rather than at startup, so the server connects instantly even when Outlook is busy serving another client's script. Use `ping` to confirm Outlook is answering before a long task.

#### Fast list and search on macOS (`outlook_db.py`)

AppleScript's `whose` clause makes Outlook walk every message in the folder, one Apple Event at a time. On a mailbox with tens of thousands of messages a single search exceeds the script timeout, and because Outlook runs one script at a time it stalls every other tool while it runs. Legacy Outlook for Mac keeps all message metadata in a SQLite database under its profile directory, and its AppleScript layer reads from that same store, so the numeric `Record_RecordID` in the database is exactly the `id` AppleScript uses.

`list_emails` and `search_emails` therefore query that database directly, read-only:

- Listing and searching a 33,000-message inbox takes milliseconds instead of minutes.
- Search matches subject, sender name, sender address, and the message preview, and ignores `Re:`/`FW:` prefixes (the database stores a normalized subject, so results show the subject without those prefixes).
- Folder names are resolved through the database's folder table, which also fixes the `inbox` keyword resolving to the empty local "On My Computer" store on Exchange profiles; action tools then address folders by id.
- The ids returned are the same ones `read_email`, `reply_email`, `move_email`, and the mark tools use through AppleScript.

Safety and fallback:

- The database is opened with `mode=ro`; nothing is ever written, and reads never block Outlook.
- At startup the server asks AppleScript for the message count of the folder the database calls the inbox. If the id is unknown to AppleScript or the counts diverge (a stale database left behind after switching to New Outlook), the database is not used.
- If the database is missing (New Outlook, no profile yet), busy, or has an unexpected schema, or a folder name is not found in it, the tool falls back to the AppleScript path above. Search then matches subject only.
- `OUTLOOK_MCP_DB_PATH` overrides the database location. The default is `~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/<profile>/Data/Outlook.sqlite`, preferring `Main Profile`.

`python tests/mac_db_test.py` covers the database layer and its integration against a fixture database, without Outlook.

#### Email bodies on macOS

Outlook's AppleScript `content` property is **HTML**. A body set with literal newlines is delivered as one run-on line, because the newlines are just HTML whitespace. The server therefore treats every body as HTML:

- **`html_body` (recommended)** for `send_email` and `reply_email`: pass an HTML fragment. `<p>`, `<br>`, `<strong>`, `<a href>`, and `<table border="1">` all render for recipients in Outlook desktop and OWA. Quotes, backslashes, and non-ASCII characters are escaped for AppleScript automatically; do not pre-escape.
- **Plain `body`** is converted server-side: `&`, `<`, `>` are entity-escaped, blank lines become `<p>` paragraphs, and single newlines become `<br>`. It is ignored when `html_body` is given.
- **Replies** insert the reply HTML immediately inside the `<body>` of Outlook's quoted-thread draft, so the original message and thread headers stay intact below the reply.
- **Outlook adds no signature.** AppleScript-created messages bypass Outlook's compose window, which is where client signatures are inserted; the Sent Items copy contains exactly the HTML you passed. A signature configured under Outlook > Preferences > Signatures is therefore never appended (and never duplicated). Include the sign-off (e.g. `<p>Regards,<br>Alex</p>`) in the body itself.
- **The mail server may add one anyway.** On an Exchange tenant with a transport-rule signature (common in corporate tenants), the corporate signature and disclaimer block is appended in transit and shows up only in the *received* copy, directly after your closing. Do not embed a full signature block in the body or recipients will see it twice.

`python tests/manual_send_test.py you@example.com` sends one HTML email and one plain-body email to the given address exercising paragraphs, bold, a clickable link column, and a bordered table, then reads the Sent Items copy back and reports whether anything was appended.

## Install from Source

### Windows (x64)

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 "mcp[cli]" -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

Register from source using the launcher script:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### Windows (ARM64)

The `[cli]` extra of `mcp` transitively pulls in `cryptography`, and pip's default resolver may pick a version that lacks a `win_arm64` wheel — which then fails to build because it requires a Rust toolchain plus OpenSSL. Install without the `cli` extra and force wheels-only resolution:

```powershell
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
& "C:\Program Files\Python312-arm64\python.exe" -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install --only-binary=:all: pywin32 mcp
python -m pip install --no-deps -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

The base `mcp` package is sufficient for running the stdio server — the `[cli]` extra is only needed for the `mcp` developer CLI tools (`mcp dev`, `mcp inspector`), which aren't used at runtime.

Register from source the same way as x64:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### macOS

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install "mcp[cli]" -e .
```

Register from source:

```bash
claude mcp add outlook-desktop -- /path/to/outlook-desktop-mcp/.venv/bin/python -m outlook_desktop_mcp
```

## Usage Examples

Once registered, just talk to Claude naturally:

- *"Show me my 10 most recent inbox emails"*
- *"Read the email from Taylor about MLADS"*
- *"Send an email to alice@example.com about the project update"*
- *"What's on my calendar this week?"*
- *"Create a meeting with bob@example.com tomorrow at 2pm for 30 minutes"*
- *"Save the attachment from that email to my Downloads folder"*
- *"Create a task to review the quarterly report, due Friday, high importance"*
- *"Mark that email as read and move it to archive"*

Windows-only examples:

- *"What categories do I have? Set this email to 'Follow-up'"*
- *"List my mail rules"*
- *"Am I set as Out of Office?"*

## Why Not Microsoft Graph?

| | Microsoft Graph | outlook-desktop-mcp |
|---|---|---|
| Entra app registration | Required | Not needed |
| Admin consent | Required for mail permissions | Not needed |
| OAuth token management | You handle refresh tokens | Not needed |
| Tenant configuration | Required | Not needed |
| Works offline / cached | No | Yes (reads from local cache) |
| Setup time | 30-60 minutes | 2 minutes |
| Auth requirement | **Your own OAuth flow** | **Outlook is open** |

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    entrypoint.py            # Platform detection → routes to correct server
    server.py                # Windows MCP server (29 tools, COM automation)
    server_mac.py            # macOS MCP server (22 tools, AppleScript)
    com_bridge.py            # Async-to-COM threading bridge (Windows)
    applescript_bridge.py    # Async osascript execution (macOS)
    outlook_db.py            # Read-only Outlook profile database queries (macOS)
    tools/
      _folder_constants.py   # Outlook enums and constants (Windows)
    utils/
      formatting.py          # Email/event/task data extraction (Windows)
      errors.py              # COM error formatting (Windows)
      applescript_helpers.py # AppleScript escaping, date formatting (macOS)
  tests/
    phase1_com_test.py       # Email COM validation
    phase3_mcp_test.py       # Email MCP test
    calendar_com_test.py     # Calendar COM validation
    calendar_mcp_test.py     # Calendar MCP test
    extras_com_test.py       # Tasks/attachments/categories/rules/OOF COM test
    extras_mcp_test.py       # Tasks/attachments/categories/rules/OOF MCP test
    mac_batch_test.py        # macOS batched AppleScript generation (no Outlook needed)
    mac_db_test.py           # macOS profile database layer (no Outlook needed)
  outlook-desktop-mcp.cmd   # Windows launcher script
  pyproject.toml
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the branching strategy and development setup.

## License

See [LICENSE](LICENSE) file.
