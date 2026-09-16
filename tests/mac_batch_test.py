"""
Outlook Desktop MCP - macOS Batch AppleScript Test
===================================================
Unit tests for the batched (fast) AppleScript generation in server_mac.
Runs without Outlook: the AppleScript bridge is replaced with a fake.

Run: python tests/mac_batch_test.py
"""
import sys
import os
import json
import asyncio
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM


def log(msg):
    print(msg, file=sys.stderr, flush=True)


class FakeBridge:
    """Captures scripts and returns canned output."""

    def __init__(self, output="", fail_first_with=None):
        self.scripts = []
        self.output = output
        self.fail_first_with = fail_first_with

    async def run(self, script, timeout=None):
        self.scripts.append(script)
        if self.fail_first_with and len(self.scripts) == 1:
            raise RuntimeError(self.fail_first_with)
        return self.output

    async def run_lines(self, script, timeout=None):
        result = await self.run(script, timeout=timeout)
        return [line for line in result.split("\n") if line.strip()]


def email_record(mid="101", subject="Hello", sender="a@b.com",
                 sender_name="Alice", time="2026-08-27 10:00:00",
                 is_read="false", att="2"):
    return DELIM.join([mid, subject, sender, sender_name, time, is_read, att])


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


def test_list_emails_uses_batch_script():
    log("--- list_emails generates batch script, parses output ---")
    fake = FakeBridge(output=email_record() + RECORD_DELIM)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_emails(folder="inbox", count=40)))

    script = fake.scripts[0]
    check("single osascript call", len(fake.scripts) == 1)
    check("batch id fetch", "id of messages 1 thru maxCount" in script, script[:200])
    check("batch subject fetch", "subject of messages 1 thru maxCount" in script)
    check("no per-message subject fetch in loop", "set msubject to subject of m\n" not in script)
    check("no full-folder materialization", "set allMsgs to messages of folderRef" not in script)
    check("one email parsed", len(result) == 1, str(result))
    if result:
        check("subject parsed", result[0]["subject"] == "Hello")
        check("unread derived from is_read", result[0]["unread"] is True)
        check("attachment count parsed", result[0]["attachment_count"] == 2)


def test_list_emails_unread_uses_whose_filter():
    log("--- list_emails unread_only filters server-side ---")
    fake = FakeBridge(output=email_record() + RECORD_DELIM)
    server_mac.bridge = fake

    asyncio.run(server_mac.list_emails(folder="inbox", count=10, unread_only=True))
    script = fake.scripts[0]
    check("whose filter present", "whose is read is false" in script)
    check("batch fetch on filtered set", "id of (every message of folderRef whose is read is false)" in script)


def test_list_emails_falls_back_to_legacy():
    log("--- list_emails falls back to legacy loop on AppleScript error ---")
    fake = FakeBridge(output=email_record() + RECORD_DELIM,
                      fail_first_with="AppleScript error: can't batch")
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_emails(folder="inbox", count=5)))
    check("two attempts made", len(fake.scripts) == 2, f"{len(fake.scripts)} scripts")
    if len(fake.scripts) == 2:
        check("legacy loop used second", "set allMsgs to messages of folderRef" in fake.scripts[1])
    check("result parsed after fallback", len(result) == 1 and result[0]["subject"] == "Hello")


def test_list_emails_timeout_not_retried():
    log("--- list_emails does not retry after a timeout ---")
    fake = FakeBridge(fail_first_with="AppleScript timed out after 30s")
    server_mac.bridge = fake

    result = asyncio.run(server_mac.list_emails(folder="inbox", count=5))
    check("only one attempt after timeout", len(fake.scripts) == 1, f"{len(fake.scripts)} scripts")
    check("error surfaced", "timed out" in result)


def test_search_emails_batches_filtered_set():
    log("--- search_emails batches over whose-filtered set ---")
    fake = FakeBridge(output=email_record(subject="budget") + RECORD_DELIM)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.search_emails(query="budget", count=10)))
    script = fake.scripts[0]
    check("whose subject filter", 'whose subject contains "budget"' in script)
    check("batch fetch present", "id of (every message of folderRef whose subject contains" in script)
    check("no per-message property loop", "set msubject to subject of m\n" not in script)
    check("result parsed", len(result) == 1 and result[0]["subject"] == "budget")


def event_record(eid="7", subject="Standup", start="2026-08-27T09:00:00",
                 end="2026-08-27T09:15:00", loc="Room 1", org="Bob", allday="false"):
    return DELIM.join([eid, subject, start, end, loc, org, allday])


def test_list_events_batches():
    log("--- list_events filters by date server-side and batches fetch ---")
    fake = FakeBridge(output=event_record() + RECORD_DELIM)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_events(
        start_date="2026-08-27", end_date="2026-08-28", count=20)))
    script = fake.scripts[0]
    check("whose date filter",
          "whose start time ≥ rangeStart and start time ≤ rangeEnd" in script)
    check("locale-safe date construction", "set year of rangeStart to 2026" in script
          and "set month of rangeStart to 8" in script)
    check("iso date output", "«class isot»" in script)
    check("no per-event property loop", "set esubject to subject of e\n" not in script)
    check("no full calendar materialization", "set evts to calendar events\n" not in script)
    check("result parsed", len(result) == 1 and result[0]["subject"] == "Standup", str(result))


def test_list_events_filters_and_sorts_in_python():
    log("--- list_events drops out-of-range events and sorts by start ---")
    recs = RECORD_DELIM.join([
        event_record(eid="3", subject="Late", start="2026-08-27T15:00:00", end="2026-08-27T16:00:00"),
        event_record(eid="1", subject="Old holiday", start="2012-01-01T00:00:00", end="2012-01-02T00:00:00"),
        event_record(eid="2", subject="Early", start="2026-08-27T09:00:00", end="2026-08-27T09:15:00"),
    ]) + RECORD_DELIM
    fake = FakeBridge(output=recs)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_events(
        start_date="2026-08-27", end_date="2026-08-28", count=20)))
    check("out-of-range event dropped", all(e["subject"] != "Old holiday" for e in result), str(result))
    check("two in-range events kept", len(result) == 2, str(result))
    check("sorted by start time", [e["subject"] for e in result] == ["Early", "Late"], str(result))


def test_list_events_truncates_to_count():
    log("--- list_events returns at most count events after sorting ---")
    recs = RECORD_DELIM.join([
        event_record(eid=str(i), subject=f"E{i}", start=f"2026-08-27T{9 + i:02d}:00:00",
                     end=f"2026-08-27T{10 + i:02d}:00:00")
        for i in range(3)
    ]) + RECORD_DELIM
    fake = FakeBridge(output=recs)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_events(
        start_date="2026-08-27", end_date="2026-08-28", count=2)))
    check("truncated to count", len(result) == 2, str(result))
    check("earliest kept", [e["subject"] for e in result] == ["E0", "E1"], str(result))


def test_list_folders_batches_names():
    log("--- list_folders enumerates per account with batched fetches ---")
    recs = RECORD_DELIM.join([
        DELIM.join(["user@example.com", "Inbox", "1200", "4"]),
        DELIM.join(["user@example.com", "Inbox/Projects", "13413", "0"]),
        DELIM.join(["On My Computer", "Inbox", "0", "0"]),
    ]) + RECORD_DELIM
    fake = FakeBridge(output=recs)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_folders()))
    script = fake.scripts[0]
    check("enumerates exchange accounts", "exchange accounts" in script)
    check("enumerates local account", "on my computer" in script)
    check("batch name fetch", "name of mail folders of" in script)
    check("batch unread fetch", "unread count of mail folders of" in script)
    check("nameless containers skipped in script", "is not missing value" in script)
    check("no per-folder name fetch", "set fname to name of f\n" not in script)
    check("account tagged", result and result[0].get("account") == "user@example.com", str(result))
    check("subfolder path kept", any(f["name"] == "Inbox/Projects" for f in result), str(result))
    check("local account tagged", any(f["account"] == "On My Computer" for f in result), str(result))
    check("counts parsed", result and result[0]["item_count"] == 1200 and result[0]["unread_count"] == 4)


def test_list_folders_depth_and_legacy():
    log("--- list_folders honors max_depth and drops nameless legacy rows ---")
    fake = FakeBridge(output="")
    server_mac.bridge = fake
    asyncio.run(server_mac.list_folders(max_depth=1))
    check("depth 1 skips subfolder pass", "mail folders of (item i1 of refs1)" not in fake.scripts[0])

    fake = FakeBridge(output="")
    server_mac.bridge = fake
    asyncio.run(server_mac.list_folders(max_depth=2))
    check("depth 2 includes subfolder pass", "mail folders of (item i1 of refs1)" in fake.scripts[0])

    fake = FakeBridge(output="")
    server_mac.bridge = fake
    asyncio.run(server_mac.list_folders(max_depth=3))
    check("depth 3 recurses further", "mail folders of (item i2 of refs2)" in fake.scripts[0])
    check("subfolder path built", 'path1 & "/" & n2' in fake.scripts[0])

    # Legacy 3-field records (flat fallback script) still parse; nameless dropped.
    recs = RECORD_DELIM.join([
        DELIM.join(["Inbox", "10", "2"]),
        DELIM.join(["missing value", "0", "0"]),
    ]) + RECORD_DELIM
    fake = FakeBridge(output=recs)
    server_mac.bridge = fake
    result = json.loads(asyncio.run(server_mac.list_folders()))
    check("legacy record parsed", any(f["name"] == "Inbox" and f["item_count"] == 10 for f in result), str(result))
    check("nameless row dropped", all(f["name"] != "missing value" for f in result), str(result))


def test_list_tasks_batches():
    log("--- list_tasks batches task fetch ---")
    rec = DELIM.join(["55", "Pay invoice", "2026-09-01 00:00:00", "not completed", "priority normal"])
    fake = FakeBridge(output=rec + RECORD_DELIM)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_tasks(count=20)))
    script = fake.scripts[0]
    check("batch id fetch on filtered tasks", "id of (every task whose todo flag is not completed)" in script)
    check("no per-task property loop", "set tname to name of t\n" not in script)
    check("result parsed", len(result) == 1 and result[0]["subject"] == "Pay invoice")


def test_text_to_html_conversion():
    log("--- text_to_html: paragraphs, line breaks, entity escaping ---")
    from outlook_desktop_mcp.utils.applescript_helpers import text_to_html

    html = text_to_html("Hi team,\n\nSee <config> & \"notes\".\nSecond line\n\nRegards,\nAlex")
    check("blank line becomes paragraph, single newline becomes <br>",
          html == ("<p>Hi team,</p>"
                   "<p>See &lt;config&gt; &amp; \"notes\".<br>Second line</p>"
                   "<p>Regards,<br>Alex</p>"), html)
    check("no raw newlines survive", "\n" not in html)
    check("CRLF normalised", text_to_html("a\r\n\r\nb") == "<p>a</p><p>b</p>", text_to_html("a\r\n\r\nb"))
    check("runs of blank lines collapse to one paragraph break",
          text_to_html("a\n\n\n\nb") == "<p>a</p><p>b</p>", text_to_html("a\n\n\n\nb"))
    check("empty input stays empty", text_to_html("") == "")
    check("unicode preserved", text_to_html("It\u2019s \u2014 caf\u00e9") == "<p>It\u2019s \u2014 caf\u00e9</p>")


def test_send_email_plain_body_becomes_html():
    log("--- send_email converts plain body to HTML so newlines survive ---")
    fake = FakeBridge(output="")
    server_mac.bridge = fake

    asyncio.run(server_mac.send_email(
        to="a@b.com", subject="Hi", body="Line one\n\nLine two\nLine three"))
    script = fake.scripts[0]
    check("content is HTML paragraphs",
          'content:"<p>Line one</p><p>Line two<br>Line three</p>"' in script, script)
    check("no escaped raw newlines in content", "\\n" not in script)
    check("html content property never used", "html content" not in script)


def test_send_email_html_body_uses_content_property():
    log("--- send_email html_body goes into content (html content is not a valid property) ---")
    fake = FakeBridge(output="")
    server_mac.bridge = fake

    html = '<p>Hi \u2014 caf\u00e9,</p><table border="1"><tr><td><a href="https://x.y/K-1">K-1</a></td></tr></table>'
    result = asyncio.run(server_mac.send_email(
        to="a@b.com", subject="Hi", body="fallback", html_body=html))
    script = fake.scripts[0]
    check("html content property never used", "html content" not in script)
    check("html placed in content with quotes escaped",
          'content:"<p>Hi \u2014 caf\u00e9,</p><table border=\\"1\\"><tr><td><a href=\\"https://x.y/K-1\\">K-1</a></td></tr></table>"' in script,
          script)
    check("plain body not sent when html given", "fallback" not in script)
    check("success message", result.startswith("Email sent"), result)


def test_send_email_escapes_backslash_and_quotes_in_subject():
    log("--- send_email escapes AppleScript-special characters ---")
    fake = FakeBridge(output="")
    server_mac.bridge = fake

    asyncio.run(server_mac.send_email(
        to="a@b.com", subject='Re: "path\\file"', body="x"))
    script = fake.scripts[0]
    check("subject escaped", 'subject:"Re: \\"path\\\\file\\""' in script, script)


def test_reply_email_inserts_html_after_body_tag():
    log("--- reply_email inserts HTML reply inside quoted-thread body ---")
    fake = FakeBridge(output="Re: thing")
    server_mac.bridge = fake

    result = asyncio.run(server_mac.reply_email(
        entry_id="42", body="Thanks.\n\nRegards,\nAlex", reply_all=True))
    script = fake.scripts[0]
    check("reply all command", "reply to m with reply to all" in script)
    check("original content read back", "set origContent to content of replyMsg" in script)
    check("plain body converted to HTML",
          'set replyHtml to "<p>Thanks.</p><p>Regards,<br>Alex</p>"' in script, script)
    check("insertion after <body", 'offset of "<body" in origContent' in script)
    check("prepend fallback when no body tag", "replyHtml & origContent" in script)
    check("no return-char joining", "& return & return &" not in script)
    check("reply sent", "send replyMsg" in script)
    check("result reports subject", "Re: thing" in result, result)

    fake = FakeBridge(output="Re: thing")
    server_mac.bridge = fake
    asyncio.run(server_mac.reply_email(entry_id="42", body="ignored", html_body="<p><strong>Done</strong></p>"))
    script = fake.scripts[0]
    check("html_body used verbatim", 'set replyHtml to "<p><strong>Done</strong></p>"' in script, script)
    check("plain body dropped when html given", "ignored" not in script)


def _has_components(script, var, year, month, day, secs):
    return (f"set year of {var} to {year}" in script
            and f"set month of {var} to {month}" in script
            and f"set day of {var} to {day}" in script
            and f"set time of {var} to {secs}" in script)


def test_create_event_uses_date_components():
    log("--- create_event builds dates from components, not date literals ---")
    fake = FakeBridge(output=DELIM.join(["9", "Standup", "x", "y"]))
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.create_event(
        subject="Standup", start="2026-09-01 07:00", end="2026-09-01T07:30:00")))
    script = fake.scripts[0]
    check("no date literal", 'date "' not in script, script)
    check("start components", _has_components(script, "startDT", 2026, 9, 1, 7 * 3600), script)
    check("end components", _has_components(script, "endDT", 2026, 9, 1, 7 * 3600 + 1800), script)
    check("properties reference variables", "start time:startDT, end time:endDT" in script, script)
    check("result parsed", result["entry_id"] == "9" and result["status"] == "created", str(result))


def test_create_meeting_uses_date_components():
    log("--- create_meeting builds dates from components ---")
    fake = FakeBridge(output="9")
    server_mac.bridge = fake

    asyncio.run(server_mac.create_meeting(
        subject="Sync", start="2026-12-24 15:45", end="2026-12-24 16:00",
        required_attendees="a@b.com; c@d.com"))
    script = fake.scripts[0]
    check("no date literal", 'date "' not in script, script)
    check("start components", _has_components(script, "startDT", 2026, 12, 24, 15 * 3600 + 45 * 60), script)
    check("end components", _has_components(script, "endDT", 2026, 12, 24, 16 * 3600), script)
    check("attendees added", script.count("make new required attendee") == 2)


def test_update_event_uses_date_components():
    log("--- update_event builds dates from components ---")
    fake = FakeBridge(output=DELIM.join(["9", "S", "x", "y", "Room"]))
    server_mac.bridge = fake

    asyncio.run(server_mac.update_event(entry_id="9", start="2026-03-05 09:00"))
    script = fake.scripts[0]
    check("no date literal", 'date "' not in script, script)
    check("start components", _has_components(script, "startDT", 2026, 3, 5, 9 * 3600), script)
    check("start assigned from variable", "set start time of e to startDT" in script, script)
    check("end untouched when not given", "endDT" not in script)

    fake = FakeBridge(output=DELIM.join(["9", "S", "x", "y", "Room"]))
    server_mac.bridge = fake
    asyncio.run(server_mac.update_event(entry_id="9", end="2026-03-05 10:00"))
    check("end assigned from variable", "set end time of e to endDT" in fake.scripts[0], fake.scripts[0])


def test_create_task_uses_date_components():
    log("--- create_task builds due date from components ---")
    fake = FakeBridge(output="5")
    server_mac.bridge = fake

    asyncio.run(server_mac.create_task(subject="Pay", due_date="2026-09-30"))
    script = fake.scripts[0]
    check("no date literal", 'date "' not in script, script)
    check("due components", _has_components(script, "dueDT", 2026, 9, 30, 0), script)
    check("due assigned from variable", "due date:dueDT" in script, script)


def test_script_timeout_env_override():
    log("--- SCRIPT_TIMEOUT honors OUTLOOK_MCP_SCRIPT_TIMEOUT env var ---")
    import importlib
    from outlook_desktop_mcp import applescript_bridge

    os.environ["OUTLOOK_MCP_SCRIPT_TIMEOUT"] = "45"
    try:
        importlib.reload(applescript_bridge)
        check("env override applied", applescript_bridge.SCRIPT_TIMEOUT == 45.0,
              str(applescript_bridge.SCRIPT_TIMEOUT))
    finally:
        del os.environ["OUTLOOK_MCP_SCRIPT_TIMEOUT"]
        importlib.reload(applescript_bridge)
    check("default raised above 30s", applescript_bridge.SCRIPT_TIMEOUT >= 60,
          str(applescript_bridge.SCRIPT_TIMEOUT))


def test_read_email_extracts_sender_and_recipients():
    log("--- read_email reads sender/recipient records via local variables ---")
    record = DELIM.join([
        "195322", "Hello", "a@b.com", "Alice", "2026-09-04 10:00:00",
        "true", "0", "r1@b.com; ", "c1@b.com; c2@b.com; ", "body text",
    ])
    for kwargs in ({"entry_id": "195322"}, {"subject_search": "Hello"}):
        fake = FakeBridge(output=record)
        server_mac.bridge = fake
        result = json.loads(asyncio.run(server_mac.read_email(**kwargs)))
        script = fake.scripts[0]
        label = "entry_id" if "entry_id" in kwargs else "subject_search"
        check(f"{label}: sender record stored before reading address",
              "set senderRec to sender of m" in script
              and "set msender to address of senderRec" in script
              and "set msenderName to name of senderRec" in script)
        check(f"{label}: no chained sender specifier",
              "address of sender of m" not in script
              and "name of sender of m" not in script)
        check(f"{label}: recipient address read from email address record",
              "set ea to email address of r" in script
              and "address of ea" in script)
        check(f"{label}: no direct address of recipient",
              "address of r &" not in script)
        check(f"{label}: to parsed", result.get("to") == "r1@b.com;", str(result))
        check(f"{label}: cc parsed", result.get("cc") == "c1@b.com; c2@b.com;", str(result))
        check(f"{label}: sender parsed", result.get("sender") == "a@b.com", str(result))


def _check_legacy_sender_shape(label, script):
    check(f"{label}: legacy sender record stored before reading address",
          "set senderRec to sender of m" in script
          and "set msender to address of senderRec" in script
          and "set msenderName to name of senderRec" in script)
    check(f"{label}: legacy script has no chained sender specifier",
          "address of sender of m" not in script
          and "name of sender of m" not in script)


def test_list_emails_legacy_reads_sender_via_variable():
    log("--- list_emails legacy fallback reads sender via local variable ---")
    fake = FakeBridge(output=email_record() + RECORD_DELIM,
                      fail_first_with="Microsoft Outlook got an error: boom")
    server_mac.bridge = fake
    result = json.loads(asyncio.run(server_mac.list_emails(folder="inbox", count=5)))
    check("legacy fallback ran", len(fake.scripts) == 2)
    if len(fake.scripts) == 2:
        _check_legacy_sender_shape("list_emails", fake.scripts[1])
    check("sender parsed", result and result[0]["sender"] == "a@b.com", str(result))


def test_search_emails_legacy_reads_sender_via_variable():
    log("--- search_emails legacy fallback reads sender via local variable ---")
    fake = FakeBridge(output=email_record() + RECORD_DELIM,
                      fail_first_with="Microsoft Outlook got an error: boom")
    server_mac.bridge = fake
    result = json.loads(asyncio.run(server_mac.search_emails(query="Hello", count=5)))
    check("legacy fallback ran", len(fake.scripts) == 2)
    if len(fake.scripts) == 2:
        _check_legacy_sender_shape("search_emails", fake.scripts[1])
    check("sender parsed", result and result[0]["sender"] == "a@b.com", str(result))


def test_get_event_reads_attendee_address_via_variable():
    log("--- get_event reads attendee address from email address record ---")
    record = DELIM.join([
        "7", "Standup", "2026-09-04 09:00:00", "2026-09-04 09:30:00",
        "Room 1", "org@b.com", "false", "body", "x@b.com; y@b.com; ",
    ])
    fake = FakeBridge(output=record)
    server_mac.bridge = fake
    result = json.loads(asyncio.run(server_mac.get_event(entry_id="7")))
    script = fake.scripts[0]
    check("attendee email address record stored",
          "set ea to email address of a" in script
          and "address of ea" in script)
    check("no direct address of attendee", "address of a &" not in script)
    check("attendees parsed", result.get("attendees") == "x@b.com; y@b.com;", str(result))


def _outlook_running():
    import subprocess
    return subprocess.run(["pgrep", "-x", "Microsoft Outlook"],
                          capture_output=True).returncode == 0


def _compiles(script):
    """Compile (never run) an AppleScript with osacompile. Returns (ok, stderr)."""
    import subprocess
    with tempfile.TemporaryDirectory() as d:
        src = os.path.join(d, "s.applescript")
        with open(src, "w") as fh:
            fh.write(script)
        proc = subprocess.run(["osacompile", "-o", os.path.join(d, "s.scpt"), src],
                              capture_output=True, text=True)
        return proc.returncode == 0, proc.stderr.strip()


def test_reply_email_uses_dictionary_reply_command():
    log("--- reply_email uses 'reply to' with the 'reply to all' parameter ---")
    for reply_all in (False, True):
        fake = FakeBridge(output="Subj")
        server_mac.bridge = fake
        result = asyncio.run(server_mac.reply_email(entry_id="1", body="Test", reply_all=reply_all))
        script = fake.scripts[0]
        check(f"reply_all={reply_all}: reported success", result.startswith("Reply sent"), result)
        check(f"reply_all={reply_all}: no invented 'reply all to' command",
              "reply all to" not in script)
        expected = ("set replyMsg to reply to m with reply to all without opening window"
                    if reply_all else "set replyMsg to reply to m without opening window")
        check(f"reply_all={reply_all}: reply command", expected in script, script[:400])
        read_at = script.find("set sentSubject to subject of replyMsg")
        send_at = script.find("send replyMsg")
        check(f"reply_all={reply_all}: subject read into a variable before send",
              read_at != -1 and send_at != -1 and read_at < send_at
              and "return sentSubject" in script
              and "return subject of replyMsg" not in script
              and "return msubject" not in script,
              script[-250:])
        if _outlook_running():
            ok, err = _compiles(script)
            check(f"reply_all={reply_all}: script compiles against Outlook", ok, err)
        else:
            log("  SKIP: compile check (Outlook not running)")


def test_send_email_looks_up_sent_copy_without_db():
    log("--- send_email falls back to an AppleScript Sent Items lookup ---")
    server_mac.db = None

    class TwoStep(FakeBridge):
        async def run(self, script, timeout=None):
            self.scripts.append(script)
            return "" if len(self.scripts) == 1 else "777"

    fake = TwoStep()
    server_mac.bridge = fake
    result = asyncio.run(server_mac.send_email(to="a@x.com", subject="Hi", body="x"))
    check("send then lookup", len(fake.scripts) == 2, str(len(fake.scripts)))
    if len(fake.scripts) == 2:
        check("lookup reads sent items", "sent items" in fake.scripts[1])
        check("lookup compares the subject", 'subject of m is "Hi"' in fake.scripts[1], fake.scripts[1])
    check("confirmation carries id", result.endswith("(Sent Items id 777)"), result)


def main():
    server_mac._SENT_CONFIRM_TIMEOUT = 0.0  # unit tests never wait for Outlook's Sent Items write
    test_list_emails_uses_batch_script()
    test_list_emails_unread_uses_whose_filter()
    test_list_emails_falls_back_to_legacy()
    test_list_emails_timeout_not_retried()
    test_search_emails_batches_filtered_set()
    test_list_events_batches()
    test_list_events_filters_and_sorts_in_python()
    test_list_events_truncates_to_count()
    test_list_folders_batches_names()
    test_list_folders_depth_and_legacy()
    test_list_tasks_batches()
    test_text_to_html_conversion()
    test_send_email_plain_body_becomes_html()
    test_send_email_html_body_uses_content_property()
    test_send_email_escapes_backslash_and_quotes_in_subject()
    test_reply_email_inserts_html_after_body_tag()
    test_create_event_uses_date_components()
    test_create_meeting_uses_date_components()
    test_update_event_uses_date_components()
    test_create_task_uses_date_components()
    test_script_timeout_env_override()
    test_read_email_extracts_sender_and_recipients()
    test_list_emails_legacy_reads_sender_via_variable()
    test_search_emails_legacy_reads_sender_via_variable()
    test_get_event_reads_attendee_address_via_variable()
    test_reply_email_uses_dictionary_reply_command()
    test_send_email_looks_up_sent_copy_without_db()

    log("=" * 50)
    log(f"{passed}/{total} checks passed")
    if passed != total:
        sys.exit(1)


if __name__ == "__main__":
    main()
