"""
Outlook Profile Database (macOS)
================================
Read-only access to the SQLite database that legacy Outlook for Mac keeps
under its profile directory. Outlook's own AppleScript layer reads from the
same store, so the numeric ``Record_RecordID`` values here are exactly the
``id`` values AppleScript uses for messages and folders. That lets list and
search run as indexed SQL queries (milliseconds on a 30k-message inbox)
instead of AppleScript ``whose`` scans (minutes), while every action tool
keeps using the ids through AppleScript unchanged.

Nothing here ever writes. The connection is opened with ``mode=ro`` and a
short busy timeout so Outlook's own writes are never blocked.
"""
import glob
import logging
import os
import re
import sqlite3
import threading
import time
from datetime import datetime
from urllib.parse import quote

logger = logging.getLogger("outlook_desktop_mcp.outlook_db")

ENV_VAR = "OUTLOOK_MCP_DB_PATH"
DEFAULT_GLOB = (
    "~/Library/Group Containers/UBF8T346G9.Office/Outlook/"
    "Outlook 15 Profiles/*/Data/Outlook.sqlite"
)

# Folder_SpecialFolderType codes observed in the Folders table.
_SPECIAL_TYPES = {
    "inbox": 1,
    "outbox": 2,
    "sent": 8,
    "sentmail": 8,
    "sent items": 8,
    "deleted": 9,
    "deleted items": 9,
    "trash": 9,
    "drafts": 10,
    "junk": 12,
    "spam": 12,
    "junk mail": 12,
    "junk email": 12,
    "archive": 15,
}

_MAIL_CLASS = 0  # Folder_FolderClass for mail folders (calendar=2, contacts=1, ...)

_MESSAGE_COLUMNS = (
    "Record_RecordID, Message_NormalizedSubject, Message_SenderList, "
    "Message_SenderAddressList, Message_ReadFlag, Message_HasAttachment, "
    "Message_TimeReceived"
)
_LIVE_ROWS = "IFNULL(Message_MarkedForDelete, 0) = 0"
_BUSY_RETRIES = 3

_PREFIX_RE = re.compile(r"^\s*(?:(?:re|fw|fwd|aw|sv|wg)\s*:\s*)+", re.IGNORECASE)


def normalize_subject(subject: str) -> str:
    """Subject without leading reply/forward prefixes, as Outlook stores it."""
    return _PREFIX_RE.sub("", subject or "").strip()


def _env_seconds(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw:
        try:
            return float(raw)
        except ValueError:
            logger.warning("Invalid %s=%r; using default %ss", name, raw, default)
    return default


# Deadline for one SQL query. A query that runs past it is interrupted and
# reported as OutlookDBError so the caller falls back to AppleScript instead
# of hanging the tool call.
DB_TIMEOUT = _env_seconds("OUTLOOK_MCP_DB_TIMEOUT", 20)


class OutlookDBError(RuntimeError):
    """The profile database is unusable (missing tables, locked, corrupt)."""


def locate(env=None) -> str | None:
    """Return the path of the profile database, or None if there is none.

    ``OUTLOOK_MCP_DB_PATH`` overrides discovery. Otherwise the standard
    profile directory is searched, preferring "Main Profile".
    """
    env = os.environ if env is None else env
    override = env.get(ENV_VAR)
    if override:
        return override if os.path.isfile(override) else None
    matches = sorted(glob.glob(os.path.expanduser(DEFAULT_GLOB)))
    if not matches:
        return None
    for path in matches:
        if "/Main Profile/" in path:
            return path
    return matches[0]


def _like_pattern(query: str) -> str:
    escaped = (
        query.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    )
    return f"%{escaped}%"


def _iso(ts) -> str:
    if not ts:
        return ""
    try:
        return datetime.fromtimestamp(int(ts)).isoformat()
    except (ValueError, OverflowError, OSError):
        return ""


def _interrupt(con: sqlite3.Connection) -> None:
    """Timer callback: abort the running statement on this connection."""
    try:
        con.interrupt()
    except sqlite3.ProgrammingError:
        pass  # connection already closed


class OutlookDB:
    """Read-only queries against one Outlook profile database."""

    def __init__(self, path: str, timeout: float = DB_TIMEOUT):
        self.path = path
        self.timeout = timeout

    def _connect(self) -> sqlite3.Connection:
        uri = f"file:{quote(self.path)}?mode=ro"
        con = sqlite3.connect(uri, uri=True, timeout=2.0)
        con.row_factory = sqlite3.Row
        return con

    def _query(self, sql: str, params=()) -> list[sqlite3.Row]:
        last = None
        for attempt in range(_BUSY_RETRIES):
            con = None
            timer = None
            try:
                con = self._connect()
                timer = threading.Timer(self.timeout, _interrupt, (con,))
                timer.daemon = True
                timer.start()
                return con.execute(sql, params).fetchall()
            except sqlite3.OperationalError as e:
                msg = str(e).lower()
                if "interrupted" in msg:
                    raise OutlookDBError(
                        f"Outlook database query exceeded {self.timeout:g}s"
                    ) from e
                if "locked" in msg or "busy" in msg:
                    last = e
                    time.sleep(0.2 * (attempt + 1))
                    continue
                raise OutlookDBError(f"Outlook database query failed: {e}") from e
            except sqlite3.DatabaseError as e:
                raise OutlookDBError(f"Outlook database unusable: {e}") from e
            finally:
                if timer is not None:
                    timer.cancel()
                if con is not None:
                    con.close()
        raise OutlookDBError(f"Outlook database busy: {last}")

    # --- folders -------------------------------------------------------

    def resolve_folder(self, name: str) -> int | None:
        """Map a user-facing folder name to its Record_RecordID.

        Built-in names use the special-folder code; anything else matches
        the folder name case-insensitively (the last segment of a
        "Parent/Child" path). Folders on a real account are preferred over
        the local "On My Computer" placeholders.
        """
        key = name.strip().lower()
        if not key:
            return None
        special = _SPECIAL_TYPES.get(key)
        if special is not None:
            rows = self._query(
                "SELECT Record_RecordID FROM Folders "
                "WHERE Folder_SpecialFolderType = ? AND Folder_FolderClass = ? "
                "ORDER BY (Record_AccountUID = 0), Record_RecordID LIMIT 1",
                (special, _MAIL_CLASS),
            )
        else:
            leaf = key.rsplit("/", 1)[-1].strip()
            rows = self._query(
                "SELECT Record_RecordID FROM Folders "
                "WHERE lower(Folder_Name) = ? AND Folder_FolderClass = ? "
                "ORDER BY (Record_AccountUID = 0), Record_RecordID LIMIT 1",
                (leaf, _MAIL_CLASS),
            )
        return int(rows[0][0]) if rows else None

    def inbox_probe(self) -> tuple[int, int] | None:
        """Return (inbox folder id, live message count) for the trust check."""
        fid = self.resolve_folder("inbox")
        if fid is None:
            return None
        rows = self._query(
            f"SELECT COUNT(*) FROM Mail WHERE Record_FolderID = ? AND {_LIVE_ROWS}",
            (fid,),
        )
        return fid, int(rows[0][0])

    # --- messages ------------------------------------------------------

    @staticmethod
    def _row_to_summary(row: sqlite3.Row) -> dict:
        has_att = bool(row["Message_HasAttachment"])
        return {
            "entry_id": str(row["Record_RecordID"]),
            "subject": (row["Message_NormalizedSubject"] or "").strip() or "(no subject)",
            "sender": (row["Message_SenderAddressList"] or "").strip(),
            "sender_name": (row["Message_SenderList"] or "").strip(),
            "received_time": _iso(row["Message_TimeReceived"]),
            "unread": not bool(row["Message_ReadFlag"]),
            "has_attachments": has_att,
            "attachment_count": 1 if has_att else 0,
        }

    def list_messages(self, folder_id: int, count: int, unread_only: bool = False) -> list[dict]:
        """Newest messages in a folder, in the same shape list_emails returns."""
        unread = " AND IFNULL(Message_ReadFlag, 0) = 0" if unread_only else ""
        rows = self._query(
            f"SELECT {_MESSAGE_COLUMNS} FROM Mail "
            f"WHERE Record_FolderID = ? AND {_LIVE_ROWS}{unread} "
            "ORDER BY Message_TimeReceived DESC, Record_RecordID DESC LIMIT ?",
            (folder_id, max(0, int(count))),
        )
        return [self._row_to_summary(r) for r in rows]

    def find_sent_copy(self, folder_id: int, subject: str, since: int) -> dict | None:
        """Newest live row in folder with this normalized subject received at or after `since`.

        Used right after a send: Outlook keeps the message in the Outbox for
        a few seconds and then writes a new Sent Items row, so the caller
        polls this until it returns a row.
        """
        rows = self._query(
            f"SELECT {_MESSAGE_COLUMNS} FROM Mail "
            f"WHERE Record_FolderID = ? AND {_LIVE_ROWS} "
            "AND Message_NormalizedSubject = ? COLLATE NOCASE "
            "AND IFNULL(Message_TimeReceived, 0) >= ? "
            "ORDER BY Message_TimeReceived DESC, Record_RecordID DESC LIMIT 1",
            (folder_id, normalize_subject(subject), int(since)),
        )
        return self._row_to_summary(rows[0]) if rows else None

    def search_messages(self, folder_id: int, query: str, count: int) -> list[dict]:
        """Case-insensitive substring search over subject, sender, and preview."""
        query = query.strip()
        if not query:
            return []
        pattern = _like_pattern(query)
        rows = self._query(
            f"SELECT {_MESSAGE_COLUMNS} FROM Mail "
            f"WHERE Record_FolderID = ? AND {_LIVE_ROWS} AND ("
            "Message_NormalizedSubject LIKE ? ESCAPE '\\' OR "
            "Message_SenderList LIKE ? ESCAPE '\\' OR "
            "Message_SenderAddressList LIKE ? ESCAPE '\\' OR "
            "Message_Preview LIKE ? ESCAPE '\\') "
            "ORDER BY Message_TimeReceived DESC LIMIT ?",
            (folder_id, pattern, pattern, pattern, pattern, max(0, int(count))),
        )
        return [self._row_to_summary(r) for r in rows]
