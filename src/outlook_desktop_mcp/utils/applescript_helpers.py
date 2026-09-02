"""Helpers for building and parsing AppleScript safely."""
import html
import re
from datetime import datetime


def escape(text: str) -> str:
    """Escape a string for safe embedding inside AppleScript double quotes.

    Handles backslashes, double quotes, and other special characters.
    """
    text = text.replace("\\", "\\\\")
    text = text.replace('"', '\\"')
    text = text.replace("\n", "\\n")
    text = text.replace("\r", "\\r")
    text = text.replace("\t", "\\t")
    return text


def text_to_html(text: str) -> str:
    """Convert plain text to minimal HTML for an Outlook message body.

    Outlook treats the `content` property of messages as HTML, so literal
    newlines collapse to whitespace. Blank lines become paragraph breaks,
    single newlines become <br>, and &, <, > are entity-escaped.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    if not text:
        return ""
    paragraphs = re.split(r"\n{2,}", text)
    return "".join(
        "<p>" + html.escape(p, quote=False).replace("\n", "<br>") + "</p>"
        for p in paragraphs
    )


def parse_date(text: str) -> str:
    """Parse an AppleScript date string to ISO 8601 format.

    AppleScript dates look like: "Sunday, March 22, 2026 at 2:00:00 PM"
    or various locale-specific formats. We attempt several common patterns.
    """
    text = text.strip()
    # Remove day name prefix if present (e.g., "Sunday, ")
    text = re.sub(r"^\w+day,\s*", "", text)
    # Remove " at " between date and time
    text = text.replace(" at ", " ")
    # Try common formats
    for fmt in (
        "%B %d, %Y %I:%M:%S %p",   # March 22, 2026 2:00:00 PM
        "%d. %B %Y %H:%M:%S",       # 22. mars 2026 14:00:00 (Norwegian)
        "%Y-%m-%d %H:%M:%S",        # 2026-03-22 14:00:00
        "%d/%m/%Y %H:%M:%S",        # 22/03/2026 14:00:00
        "%m/%d/%Y %H:%M:%S",        # 03/22/2026 14:00:00
    ):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.isoformat()
        except ValueError:
            continue
    # Fallback: return as-is
    return text


def date_component_lines(var: str, dt: datetime) -> str:
    """Build AppleScript that constructs a date in `var` from components.

    Unlike `date "..."` literals, component assignment does not depend on
    the system locale. Day is reset to 1 first so month assignment cannot
    overflow (e.g. current date Aug 31 -> month 2 would roll into March).
    """
    secs = dt.hour * 3600 + dt.minute * 60 + dt.second
    return (
        f"set {var} to current date\n"
        f"set day of {var} to 1\n"
        f"set year of {var} to {dt.year}\n"
        f"set month of {var} to {dt.month}\n"
        f"set day of {var} to {dt.day}\n"
        f"set time of {var} to {secs}"
    )


# Locale-independent AppleScript folder keywords
FOLDER_MAP = {
    "inbox": "inbox",
    "sent": "sent items",
    "sentmail": "sent items",
    "sent items": "sent items",
    "drafts": "drafts",
    "deleted": "deleted items",
    "deleted items": "deleted items",
    "trash": "deleted items",
    "junk": "junk mail",
    "spam": "junk mail",
    "outbox": "outbox",
}


def resolve_folder_ref(folder_name: str) -> str:
    """Map a user-facing folder name to an AppleScript folder reference.

    Returns an AppleScript expression like 'inbox' or 'mail folder "Archive"'.
    Built-in folders use locale-independent keywords; custom folders use name lookup.
    """
    key = folder_name.lower().strip()
    if key in FOLDER_MAP:
        return FOLDER_MAP[key]
    # Custom folder — search by name
    return f'mail folder "{escape(folder_name)}"'


# Delimiter used for structured AppleScript output
DELIM = "|||"
RECORD_DELIM = "==="
