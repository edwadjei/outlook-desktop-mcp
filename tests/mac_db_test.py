"""
Outlook Desktop MCP - macOS SQLite Index Test
==============================================
Unit tests for the read-only Outlook profile database layer (outlook_db)
and its integration into server_mac. Runs without Outlook: a fixture
database with the same schema subset is built in a temp directory, and
the AppleScript bridge is replaced with a fake.

Run: python tests/mac_db_test.py
"""
import sys
import os
import json
import asyncio
import sqlite3
import tempfile
import time
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import outlook_db
from outlook_desktop_mcp.outlook_db import OutlookDB


def log(msg):
    print(msg, file=sys.stderr, flush=True)


passed = 0
total = 0


def check(name, condition, detail=""):
    global passed, total
    total += 1
    if condition:
        passed += 1
        log(f"  PASS: {name}")
    else:
        log(f"  FAIL: {name} {detail}")


# --- Fixture database -------------------------------------------------

EXCHANGE = 60129542145  # account UID of the Exchange mailbox
LOCAL = 0               # account UID of the "On My Computer" placeholders

FOLDERS = [
    # id, account, parent, special, class, name
    (1, LOCAL, 7, 1, 0, "Placeholder_Inbox_Placeholder"),
    (3, LOCAL, 7, 8, 0, "Placeholder_Sent_Items_Placeholder"),
    (101, EXCHANGE, -2, 99, 0, ""),
    (109, EXCHANGE, 101, 15, 0, "Archive"),
    (112, EXCHANGE, 101, 9, 0, "Deleted Items"),
    (114, EXCHANGE, 101, 10, 0, "Drafts"),
    (115, EXCHANGE, 101, 1, 0, "Inbox"),
    (116, EXCHANGE, 115, 0, 0, "Dev Portal Approvals"),
    (124, EXCHANGE, 101, 12, 0, "Junk Email"),
    (127, EXCHANGE, 101, 8, 0, "Sent Items"),
    (106, EXCHANGE, 103, 4, 2, "Calendar"),
]

T0 = 1788400000  # 2026-09-03 local

MAIL = [
    # id, folder, subject, sender name, sender addr, preview, read, att, time, del, hidden[, to addrs, cc addrs, display to]
    (201, 115, "Quarterly budget report", "Alice", "alice@x.com", "Numbers attached", 0, 1, T0 + 300, 0, 0),
    (202, 115, "Lunch plans", "Bob", "bob@x.com", "Pizza?", 1, 0, T0 + 200, 0, 0),
    (203, 115, "Sandbox access", "Carol", "carol@x.com", "Your sandbox is ready", 0, 0, T0 + 100, 0, 0),
    (204, 115, "Old newsletter", "News", "news@x.com", "Budget tips inside", 1, 0, T0, 0, 0),
    (205, 115, "Deleted locally", "Dave", "dave@x.com", "gone", 0, 0, T0 + 400, 1, 0),
    (206, 116, "Approval needed", "Portal", "portal@x.com", "Approve me", 0, 0, T0 + 50, 0, 0),
    (207, 127, "Re: budget", "Me", "me@x.com", "Sent reply", 1, 0, T0 + 10, 0, 0),
]


def build_fixture(dir_path, mail=MAIL, folders=FOLDERS):
    path = os.path.join(dir_path, "Outlook.sqlite")
    con = sqlite3.connect(path)
    con.executescript("""
        CREATE TABLE Folders (
            Record_RecordID INTEGER PRIMARY KEY,
            Record_AccountUID INTEGER,
            Folder_ParentID INTEGER,
            Folder_SpecialFolderType INTEGER,
            Folder_FolderClass INTEGER,
            Folder_Name TEXT
        );
        CREATE TABLE Mail (
            Record_RecordID INTEGER PRIMARY KEY,
            Record_FolderID INTEGER,
            Message_NormalizedSubject TEXT,
            Message_SenderList TEXT,
            Message_SenderAddressList TEXT,
            Message_Preview TEXT,
            Message_ReadFlag INTEGER,
            Message_HasAttachment INTEGER,
            Message_TimeReceived INTEGER,
            Message_MarkedForDelete INTEGER,
            Message_Hidden INTEGER,
            Message_ToRecipientAddressList TEXT,
            Message_CCRecipientAddressList TEXT,
            Message_DisplayTo TEXT
        );
        CREATE INDEX MailIndex_TimeWindow ON Mail (Record_FolderID, Message_TimeReceived DESC);
    """)
    con.executemany("INSERT INTO Folders VALUES (?,?,?,?,?,?)", folders)
    con.executemany("INSERT INTO Mail VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    [tuple(row) + (None,) * (14 - len(row)) for row in mail])
    con.commit()
    con.close()
    return path


# --- outlook_db unit tests --------------------------------------------

def test_locate_missing_file():
    log("--- locate returns None when the database file does not exist ---")
    path = outlook_db.locate({"OUTLOOK_MCP_DB_PATH": "/nonexistent/Outlook.sqlite"})
    check("missing file -> None", path is None, str(path))


def test_locate_env_override():
    log("--- locate honours OUTLOOK_MCP_DB_PATH ---")
    with tempfile.TemporaryDirectory() as d:
        fixture = build_fixture(d)
        path = outlook_db.locate({"OUTLOOK_MCP_DB_PATH": fixture})
        check("override path returned", path == fixture, str(path))


def test_resolve_builtin_folders_prefer_exchange():
    log("--- resolve_folder maps built-in names to the Exchange folder ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        check("inbox -> Exchange Inbox (115), not local placeholder (1)",
              db.resolve_folder("inbox") == 115, str(db.resolve_folder("inbox")))
        check("sent -> 127", db.resolve_folder("sent") == 127)
        check("sentmail alias", db.resolve_folder("sentmail") == 127)
        check("trash -> Deleted Items 112", db.resolve_folder("trash") == 112)
        check("spam -> Junk Email 124", db.resolve_folder("spam") == 124)
        check("drafts -> 114", db.resolve_folder("drafts") == 114)
        check("archive -> 109", db.resolve_folder("archive") == 109)
        check("INBOX case-insensitive", db.resolve_folder("  INBOX ") == 115)


def test_resolve_custom_folder_by_name():
    log("--- resolve_folder finds custom folders by name ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        check("exact name", db.resolve_folder("Dev Portal Approvals") == 116)
        check("case-insensitive name", db.resolve_folder("dev portal approvals") == 116)
        check("path form uses last segment", db.resolve_folder("Inbox/Dev Portal Approvals") == 116)
        check("unknown -> None", db.resolve_folder("No Such Folder") is None)
        check("calendar (non-mail class) not resolved as mail folder",
              db.resolve_folder("Calendar") is None)


def test_resolve_falls_back_to_local_account():
    log("--- resolve_folder uses the local store when no Exchange folder exists ---")
    local_only = [f for f in FOLDERS if f[1] == LOCAL]
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d, folders=local_only))
        check("inbox -> local placeholder 1", db.resolve_folder("inbox") == 1)


def test_list_messages_newest_first_and_mapped():
    log("--- list_messages returns newest first with mapped fields ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        rows = db.list_messages(115, count=2)
        check("count limited", len(rows) == 2, str(len(rows)))
        check("newest first (201 before 202)", [r["entry_id"] for r in rows] == ["201", "202"],
              str([r["entry_id"] for r in rows]))
        r = rows[0]
        check("entry_id is a string", isinstance(r["entry_id"], str))
        check("subject", r["subject"] == "Quarterly budget report")
        check("sender address", r["sender"] == "alice@x.com")
        check("sender name", r["sender_name"] == "Alice")
        check("unread derived from read flag", r["unread"] is True)
        check("read message -> unread False", rows[1]["unread"] is False)
        check("has_attachments", r["has_attachments"] is True and r["attachment_count"] == 1)
        check("no attachments", rows[1]["has_attachments"] is False and rows[1]["attachment_count"] == 0)
        expected = datetime.fromtimestamp(T0 + 300).isoformat()
        check("received_time is local ISO", r["received_time"] == expected, r["received_time"])


def test_list_messages_excludes_deleted_and_other_folders():
    log("--- list_messages skips locally deleted rows and other folders ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        ids = [r["entry_id"] for r in db.list_messages(115, count=50)]
        check("deleted row 205 excluded", "205" not in ids, str(ids))
        check("subfolder row 206 excluded", "206" not in ids, str(ids))
        check("four inbox rows", len(ids) == 4, str(ids))


def test_list_messages_unread_only():
    log("--- list_messages unread_only filters on read flag ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        ids = [r["entry_id"] for r in db.list_messages(115, count=50, unread_only=True)]
        check("only unread rows", ids == ["201", "203"], str(ids))


def test_search_matches_subject_sender_preview():
    log("--- search_messages matches subject, sender, and preview ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        ids = [r["entry_id"] for r in db.search_messages(115, "budget", count=10)]
        check("subject + preview matches, newest first", ids == ["201", "204"], str(ids))
        ids = [r["entry_id"] for r in db.search_messages(115, "CAROL", count=10)]
        check("sender name, case-insensitive", ids == ["203"], str(ids))
        ids = [r["entry_id"] for r in db.search_messages(115, "bob@x.com", count=10)]
        check("sender address", ids == ["202"], str(ids))
        ids = [r["entry_id"] for r in db.search_messages(115, "budget", count=1)]
        check("count limit", ids == ["201"], str(ids))
        ids = [r["entry_id"] for r in db.search_messages(127, "budget", count=10)]
        check("scoped to folder", ids == ["207"], str(ids))


def test_search_escapes_like_wildcards():
    log("--- search_messages treats % and _ literally ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        check("percent matches nothing", db.search_messages(115, "%", count=10) == [])
        check("underscore matches nothing", db.search_messages(115, "_", count=10) == [])
        check("empty query matches nothing", db.search_messages(115, "   ", count=10) == [])


def test_search_excludes_deleted():
    log("--- search_messages skips locally deleted rows ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        check("deleted row 205 not found", db.search_messages(115, "gone", count=10) == [])


def test_inbox_probe():
    log("--- inbox_probe reports the Exchange inbox id and row count ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        check("probe -> (115, 4)", db.inbox_probe() == (115, 4), str(db.inbox_probe()))


def test_schema_mismatch_raises_clean_error():
    log("--- a database without the expected tables raises OutlookDBError ---")
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "Outlook.sqlite")
        sqlite3.connect(path).executescript("CREATE TABLE Other (x);")
        db = OutlookDB(path)
        try:
            db.resolve_folder("inbox")
            check("raised", False)
        except outlook_db.OutlookDBError:
            check("raised OutlookDBError", True)


def test_read_only_never_writes():
    log("--- the connection is read-only ---")
    with tempfile.TemporaryDirectory() as d:
        path = build_fixture(d)
        db = OutlookDB(path)
        before = os.stat(path).st_mtime_ns
        db.list_messages(115, count=5)
        db.search_messages(115, "budget", count=5)
        check("file untouched", os.stat(path).st_mtime_ns == before)
        try:
            db._connect().execute("DELETE FROM Mail")
            check("write rejected", False)
        except sqlite3.OperationalError:
            check("write rejected", True)


# --- server_mac integration -------------------------------------------

from outlook_desktop_mcp import server_mac
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM


class FakeBridge:
    def __init__(self, output=""):
        self.scripts = []
        self.output = output

    async def start(self):
        # Mirrors AppleScriptBridge.start(): one version probe under STARTUP_TIMEOUT.
        await self.run('tell application "Microsoft Outlook" to get version',
                       timeout=server_mac.STARTUP_TIMEOUT)

    async def run(self, script, timeout=None):
        self.scripts.append(script)
        return self.output

    async def run_lines(self, script, timeout=None):
        return [l for l in (await self.run(script)).split("\n") if l.strip()]


def test_server_list_emails_uses_db():
    log("--- list_emails serves from the database when trusted ---")
    with tempfile.TemporaryDirectory() as d:
        server_mac.db = OutlookDB(build_fixture(d))
        fake = FakeBridge()
        server_mac.bridge = fake
        result = json.loads(asyncio.run(server_mac.list_emails(folder="inbox", count=3)))
        check("no AppleScript run", fake.scripts == [], str(fake.scripts)[:100])
        check("three rows", len(result) == 3, str(len(result)))
        check("first is newest", result and result[0]["entry_id"] == "201")
        check("same JSON shape as AppleScript path",
              result and set(result[0]) == {"entry_id", "subject", "sender", "sender_name",
                                            "received_time", "unread", "has_attachments",
                                            "attachment_count"}, str(result[:1]))
        result = json.loads(asyncio.run(server_mac.list_emails(folder="inbox", unread_only=True)))
        check("unread_only via db", [r["entry_id"] for r in result] == ["201", "203"])
    server_mac.db = None


def test_server_search_emails_uses_db():
    log("--- search_emails serves from the database when trusted ---")
    with tempfile.TemporaryDirectory() as d:
        server_mac.db = OutlookDB(build_fixture(d))
        fake = FakeBridge()
        server_mac.bridge = fake
        result = json.loads(asyncio.run(server_mac.search_emails(query="budget", folder="inbox")))
        check("no AppleScript run", fake.scripts == [])
        check("matches", [r["entry_id"] for r in result] == ["201", "204"], str(result))
        result = json.loads(asyncio.run(server_mac.search_emails(query="zzz-none")))
        check("no match -> empty list, no fallback", result == [] and fake.scripts == [])
    server_mac.db = None


def test_server_unknown_folder_falls_back_to_applescript():
    log("--- folders the database cannot resolve go to AppleScript ---")
    with tempfile.TemporaryDirectory() as d:
        server_mac.db = OutlookDB(build_fixture(d))
        rec = DELIM.join(["9", "Hi", "a@b", "A", "2026-08-27 10:00:00", "true", "0"]) + RECORD_DELIM
        fake = FakeBridge(output=rec)
        server_mac.bridge = fake
        result = json.loads(asyncio.run(server_mac.list_emails(folder="Mystery")))
        check("AppleScript ran", len(fake.scripts) == 1)
        check("name lookup in script", 'mail folder "Mystery"' in fake.scripts[0])
        check("AppleScript result returned", result and result[0]["entry_id"] == "9")
    server_mac.db = None


def test_server_db_error_falls_back_to_applescript():
    log("--- a database failure falls back to AppleScript ---")
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "Outlook.sqlite")
        sqlite3.connect(path).executescript("CREATE TABLE Other (x);")
        server_mac.db = OutlookDB(path)
        rec = DELIM.join(["9", "Hi", "a@b", "A", "2026-08-27 10:00:00", "true", "0"]) + RECORD_DELIM
        fake = FakeBridge(output=rec)
        server_mac.bridge = fake
        result = json.loads(asyncio.run(server_mac.list_emails(folder="inbox")))
        check("AppleScript ran", len(fake.scripts) == 1)
        check("result from AppleScript", result and result[0]["entry_id"] == "9")
    server_mac.db = None


def test_server_no_db_uses_applescript():
    log("--- with no database the AppleScript path is unchanged ---")
    server_mac.db = None
    rec = DELIM.join(["9", "Hi", "a@b", "A", "2026-08-27 10:00:00", "true", "0"]) + RECORD_DELIM
    fake = FakeBridge(output=rec)
    server_mac.bridge = fake
    result = json.loads(asyncio.run(server_mac.search_emails(query="Hi")))
    check("AppleScript ran", len(fake.scripts) == 1)
    check("whose clause present", 'whose subject contains "Hi"' in fake.scripts[0])
    check("parsed", result and result[0]["entry_id"] == "9")


def test_server_folder_ref_uses_db_id():
    log("--- AppleScript folder references use the resolved folder id ---")
    with tempfile.TemporaryDirectory() as d:
        server_mac.db = OutlookDB(build_fixture(d))
        fake = FakeBridge(output="Subj")
        server_mac.bridge = fake
        asyncio.run(server_mac.move_email(entry_id="201", target_folder="archive"))
        check("move targets mail folder id 109", "move m to mail folder id 109" in fake.scripts[0],
              fake.scripts[0])
        fake.scripts.clear()
        asyncio.run(server_mac.move_email(entry_id="201", target_folder="Mystery"))
        check("unknown folder keeps name lookup", 'mail folder "Mystery"' in fake.scripts[0])
    server_mac.db = None
    fake = FakeBridge(output="Subj")
    server_mac.bridge = fake
    asyncio.run(server_mac.move_email(entry_id="201", target_folder="inbox"))
    check("no db -> inbox keyword", "move m to inbox" in fake.scripts[0], fake.scripts[0])


def test_server_trust_check():
    log("--- trust check compares database inbox with AppleScript ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d))
        fake = FakeBridge(output="Inbox|||4")
        server_mac.bridge = fake
        check("agreeing counts -> trusted", asyncio.run(server_mac._trust_db(db)) is True)
        check("probe asks for folder id 115", "mail folder id 115" in fake.scripts[0])
        server_mac.bridge = FakeBridge(output="Inbox|||6")
        check("small drift -> trusted", asyncio.run(server_mac._trust_db(db)) is True)
        server_mac.bridge = FakeBridge(output="Inbox|||900")
        check("large drift -> untrusted", asyncio.run(server_mac._trust_db(db)) is False)

        class ErrBridge(FakeBridge):
            async def run(self, script, timeout=None):
                raise RuntimeError("AppleScript error: folder not found")

        server_mac.bridge = ErrBridge()
        check("folder id unknown to AppleScript -> untrusted",
              asyncio.run(server_mac._trust_db(db)) is False)

        class SlowBridge(FakeBridge):
            def __init__(self):
                super().__init__()
                self.timeouts = []

            async def run(self, script, timeout=None):
                self.timeouts.append(timeout)
                raise RuntimeError(f"AppleScript timed out after {timeout}s")

        slow = SlowBridge()
        server_mac.bridge = slow
        check("AppleScript timeout -> undecided (None)", asyncio.run(server_mac._trust_db(db)) is None)
        check("probe uses the startup timeout", slow.timeouts == [server_mac.STARTUP_TIMEOUT],
              str(slow.timeouts))


def _reset_db_state():
    server_mac.db = None
    server_mac._db_candidate = None
    server_mac._db_state = "none"
    server_mac._db_lock = None


def test_server_trust_check_is_deferred_to_first_use():
    log("--- trust check runs on first database use, not at startup ---")
    with tempfile.TemporaryDirectory() as d:
        path = build_fixture(d)
        os.environ[outlook_db.ENV_VAR] = path
        try:
            _reset_db_state()

            class VersionBridge(FakeBridge):
                async def run(self, script, timeout=None):
                    self.scripts.append(script)
                    return "16.93.2"

            vb = VersionBridge()
            server_mac.bridge = vb
            asyncio.run(server_mac.startup())
            check("startup ran only the version probe", len(vb.scripts) == 1 and "version" in vb.scripts[0])
            check("startup leaves db unset", server_mac.db is None)
            check("state is unchecked", server_mac._db_state == "unchecked")

            # First use: AppleScript busy -> timeout -> stays unchecked, tool still answers.
            class SlowBridge(FakeBridge):
                async def run(self, script, timeout=None):
                    self.scripts.append(script)
                    raise RuntimeError(f"AppleScript timed out after {timeout}s")

            server_mac.bridge = SlowBridge()
            fid = asyncio.run(server_mac._db_folder_id("inbox"))
            check("timeout keeps state unchecked", server_mac._db_state == "unchecked")
            check("no folder id while undecided", fid is None)

            # Next use: AppleScript agrees -> trusted, and stays trusted.
            server_mac.bridge = FakeBridge(output="Inbox|||4")
            fid = asyncio.run(server_mac._db_folder_id("inbox"))
            check("agreement -> trusted", server_mac._db_state == "trusted" and server_mac.db is not None)
            check("folder id resolved from database", fid == 115)
            server_mac.bridge = FakeBridge(output="Inbox|||900")
            fid = asyncio.run(server_mac._db_folder_id("inbox"))
            check("trusted verdict is cached", fid == 115 and len(server_mac.bridge.scripts) == 0)

            # Mismatch on first use -> untrusted, cached, AppleScript not asked again.
            _reset_db_state()
            server_mac.bridge = VersionBridge()
            asyncio.run(server_mac.startup())
            bad = FakeBridge(output="Inbox|||900")
            server_mac.bridge = bad
            asyncio.run(server_mac._db_folder_id("inbox"))
            check("mismatch -> untrusted", server_mac._db_state == "untrusted")
            asyncio.run(server_mac._db_folder_id("inbox"))
            check("untrusted verdict is cached", len(bad.scripts) == 1)
        finally:
            del os.environ[outlook_db.ENV_VAR]
            _reset_db_state()


def test_ping_tool():
    log("--- ping reports Outlook reachability and database state ---")
    _reset_db_state()
    fake = FakeBridge(output="16.93.2")
    server_mac.bridge = fake
    result = json.loads(asyncio.run(server_mac.ping()))
    check("ok when Outlook answers", result["ok"] is True, str(result))
    check("reports version", result["outlook_version"] == "16.93.2")
    check("reports db state", result["db"] == "none")
    check("reports uptime", isinstance(result["uptime_s"], int))
    check("reports server version", isinstance(result["server_version"], str) and result["server_version"])
    check("applescript_ms measured", isinstance(result["applescript_ms"], int))

    class DeadBridge(FakeBridge):
        def __init__(self):
            super().__init__()
            self.timeouts = []

        async def run(self, script, timeout=None):
            self.timeouts.append(timeout)
            raise RuntimeError("AppleScript timed out after 10s")

    dead = DeadBridge()
    server_mac.bridge = dead
    result = json.loads(asyncio.run(server_mac.ping()))
    check("not ok when Outlook is silent", result["ok"] is False)
    check("error carried", "timed out" in result.get("error", ""))
    check("ping uses the short startup timeout", dead.timeouts == [server_mac.STARTUP_TIMEOUT])
    names = [t.name for t in asyncio.run(server_mac.mcp.list_tools())]
    check("ping is registered as a tool", "ping" in names)


def test_query_deadline_interrupts_long_query():
    log("--- a query past the deadline is interrupted and raises OutlookDBError ---")
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d), timeout=0.2)
        t0 = time.time()
        try:
            db._query("WITH RECURSIVE c(x) AS (SELECT 1 UNION ALL SELECT x + 1 FROM c) "
                      "SELECT COUNT(*) FROM c")
            check("raised OutlookDBError", False, "query finished")
        except outlook_db.OutlookDBError as e:
            check("raised OutlookDBError", "exceeded" in str(e), str(e))
        check("returned promptly", time.time() - t0 < 2.0, f"{time.time() - t0:.2f}s")
        check("normal query still works", db.resolve_folder("inbox") == 115)
    check("default deadline is 20s", outlook_db.DB_TIMEOUT == 20.0, str(outlook_db.DB_TIMEOUT))


def test_list_messages_tie_breaks_on_record_id():
    log("--- equal timestamps are ordered by record id, newest first ---")
    mail = MAIL + [
        (301, 115, "Same second A", "A", "a@x.com", "", 0, 0, T0 + 500, 0, 0),
        (302, 115, "Same second B", "B", "b@x.com", "", 0, 0, T0 + 500, 0, 0),
    ]
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d, mail=mail))
        ids = [r["entry_id"] for r in db.list_messages(115, 2)]
        check("newest record first among equals", ids == ["302", "301"], str(ids))


def test_normalize_subject():
    log("--- normalize_subject strips reply and forward prefixes ---")
    check("RE:", outlook_db.normalize_subject("RE: budget") == "budget")
    check("nested", outlook_db.normalize_subject("Re: FW: Fwd: budget") == "budget")
    check("plain", outlook_db.normalize_subject("  budget  ") == "budget")
    check("inner Re kept", outlook_db.normalize_subject("Re: about the Re: thing") == "about the Re: thing")


def test_find_sent_copy():
    log("--- find_sent_copy returns the newest matching sent row after `since` ---")
    mail = MAIL + [
        (401, 127, "budget", "Me", "me@x.com", "older reply", 1, 0, T0 + 1000, 0, 0),
        (402, 127, "budget", "Me", "me@x.com", "newer reply", 1, 0, T0 + 2000, 0, 0),
        (403, 127, "budget", "Me", "me@x.com", "deleted", 1, 0, T0 + 3000, 1, 0),
    ]
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d, mail=mail))
        row = db.find_sent_copy(127, "RE: budget", since=T0 + 1500)
        check("newest live row after since", row is not None and row["entry_id"] == "402", str(row))
        check("older rows ignored", db.find_sent_copy(127, "budget", since=T0 + 2500) is None)
        check("case-insensitive subject", db.find_sent_copy(127, "BUDGET", since=T0) is not None)
        check("other folder ignored", db.find_sent_copy(115, "budget", since=T0) is None)


def test_server_send_email_confirms_sent_copy():
    log("--- send_email and reply_email report the Sent Items id ---")
    with tempfile.TemporaryDirectory() as d:
        mail = MAIL + [(501, 127, "Hello there", "Me", "me@x.com", "", 1, 0, int(time.time()) + 5, 0, 0)]
        server_mac.db = OutlookDB(build_fixture(d, mail=mail))
        server_mac.bridge = FakeBridge(output="")
        result = asyncio.run(server_mac.send_email(to="a@x.com", subject="Hello there", body="hi"))
        check("send confirmation carries the Sent Items id", result.endswith("(Sent Items id 501)"), result)
        server_mac.bridge = FakeBridge(output="RE: Hello there")
        result = asyncio.run(server_mac.reply_email(entry_id="1", body="hi"))
        check("reply confirmation carries the Sent Items id", result.endswith("(Sent Items id 501)"), result)

        old = server_mac._SENT_CONFIRM_TIMEOUT
        server_mac._SENT_CONFIRM_TIMEOUT = 0.6
        try:
            t0 = time.time()
            result = asyncio.run(server_mac.send_email(to="a@x.com", subject="Never lands", body="hi"))
            check("missing copy reported", result.endswith("(Sent Items copy not visible yet; verify with search_emails)"), result)
            check("gives up after the confirm timeout", 0.5 < time.time() - t0 < 3.0, f"{time.time() - t0:.2f}s")
        finally:
            server_mac._SENT_CONFIRM_TIMEOUT = old
    server_mac.db = None


def test_search_by_recipient_and_sender():
    log("--- search filters by recipient and sender ---")
    mail = MAIL + [
        (601, 115, "New Service Integration - Kunim", "Derrick", "derrick@x.com", "please approve",
         0, 0, T0 + 600, 0, 0, "edward@x.com", "isdemand@x.com", "Edward Adjei"),
        (602, 115, "Weekly digest", "News", "news@x.com", "approve nothing",
         0, 0, T0 + 700, 0, 0, "all@x.com", "", "Everyone"),
        (603, 115, "Approval request", "Derrick", "derrick@x.com", "second one",
         0, 0, T0 + 800, 0, 0, "bob@x.com", "edward@x.com", "Bob"),
    ]
    with tempfile.TemporaryDirectory() as d:
        db = OutlookDB(build_fixture(d, mail=mail))
        ids = lambda rows: [r["entry_id"] for r in rows]
        check("recipient matches to and cc", ids(db.search_messages(115, "", 10, recipient="edward")) == ["603", "601"])
        check("recipient matches display name", ids(db.search_messages(115, "", 10, recipient="Adjei")) == ["601"])
        check("query AND recipient", ids(db.search_messages(115, "approv", 10, recipient="edward")) == ["603", "601"])
        check("sender filter", ids(db.search_messages(115, "", 10, sender="derrick")) == ["603", "601"])
        check("all three combined", ids(db.search_messages(115, "kunim", 10, recipient="edward", sender="derrick")) == ["601"])
        check("nothing given -> empty", db.search_messages(115, "", 10) == [])
        check("rows without recipient columns filled do not match", ids(db.search_messages(115, "", 10, recipient="x.com")) == ["603", "602", "601"])


def test_server_search_emails_filters():
    log("--- search_emails passes recipient and sender to the database ---")
    _reset_db_state()
    with tempfile.TemporaryDirectory() as d:
        mail = MAIL + [(701, 115, "Approve me please", "Derrick", "derrick@x.com", "", 0, 0, T0 + 900, 0, 0,
                        "edward@x.com", "", "Edward")]
        server_mac.db = OutlookDB(build_fixture(d, mail=mail))
        server_mac.bridge = FakeBridge(output="")
        rows = json.loads(asyncio.run(server_mac.search_emails(query="approve", recipient="edward")))
        check("filtered hit", [r["entry_id"] for r in rows] == ["701"], str(rows))
        rows = json.loads(asyncio.run(server_mac.search_emails(recipient="edward")))
        check("recipient alone works", [r["entry_id"] for r in rows] == ["701"], str(rows))
        result = json.loads(asyncio.run(server_mac.search_emails()))
        check("no criteria -> error", "error" in result, str(result))
        # Filters need the database; the AppleScript fallback cannot honour them.
        server_mac.db = None
        result = json.loads(asyncio.run(server_mac.search_emails(query="approve", recipient="edward")))
        check("filters without database -> error, not a silent subject search", "error" in result and "recipient" in result["error"], str(result))
    _reset_db_state()


def main():
    server_mac._SENT_CONFIRM_TIMEOUT = 0.0  # unit tests never wait for Outlook's Sent Items write
    test_locate_missing_file()
    test_locate_env_override()
    test_resolve_builtin_folders_prefer_exchange()
    test_resolve_custom_folder_by_name()
    test_resolve_falls_back_to_local_account()
    test_list_messages_newest_first_and_mapped()
    test_list_messages_excludes_deleted_and_other_folders()
    test_list_messages_unread_only()
    test_search_matches_subject_sender_preview()
    test_search_escapes_like_wildcards()
    test_search_excludes_deleted()
    test_inbox_probe()
    test_schema_mismatch_raises_clean_error()
    test_read_only_never_writes()
    test_server_list_emails_uses_db()
    test_server_search_emails_uses_db()
    test_server_unknown_folder_falls_back_to_applescript()
    test_server_db_error_falls_back_to_applescript()
    test_server_no_db_uses_applescript()
    test_server_folder_ref_uses_db_id()
    test_server_trust_check()
    test_server_trust_check_is_deferred_to_first_use()
    test_ping_tool()
    test_query_deadline_interrupts_long_query()
    test_list_messages_tie_breaks_on_record_id()
    test_normalize_subject()
    test_find_sent_copy()
    test_server_send_email_confirms_sent_copy()
    test_search_by_recipient_and_sender()
    test_server_search_emails_filters()

    log("=" * 50)
    log(f"{passed}/{total} checks passed")
    if passed != total:
        sys.exit(1)


if __name__ == "__main__":
    main()
