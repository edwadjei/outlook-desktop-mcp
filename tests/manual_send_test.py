"""
Outlook Desktop MCP - macOS manual send test
============================================
Sends ONE real HTML email and ONE plain-body email through the running
Outlook app so a human can check the rendered formatting (paragraph spacing,
bold, clickable links, bordered table, unicode) in Outlook desktop and OWA.
Then reads the Sent Items copy back and reports whether Outlook added
anything client-side.

Requires Outlook for Mac running in legacy mode with an account signed in.

Run: python tests/manual_send_test.py recipient@example.com
"""
import sys
import os
import asyncio
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac
from outlook_desktop_mcp.applescript_bridge import AppleScriptBridge
from outlook_desktop_mcp.utils.applescript_helpers import escape

STAMP = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
SUBJECT = f"[MCP format test {STAMP}] partner handover"

ROWS = [
    ("Acme Payments", "USSD", "*7890#", "2026-08-12", "PROJ-1041"),
    ("Northwind Ltd", "SMS Bulk", "1900", "2026-08-21", "PROJ-1077"),
    ("Contoso Fintech", "API", "n/a", "2026-08-29", "PROJ-1102"),
]

table_rows = "".join(
    "<tr>"
    f"<td>{name}</td><td>{svc}</td><td>{code}</td><td>{live}</td>"
    f'<td><a href="https://example.atlassian.net/browse/{key}">{key}</a></td>'
    "</tr>"
    for name, svc, code, live, key in ROWS
)

HTML_BODY = (
    "<p>Hi team,</p>"
    "<p>Please find below the partner go-lives for August — this is a "
    "formatting test sent via the outlook-desktop MCP server, so the content "
    "is <strong>sample data only</strong>.</p>"
    "<p>Kindly confirm the ticket references are clickable and the table "
    "borders render so we can use this layout for the monthly handover.</p>"
    '<table border="1" cellpadding="4" cellspacing="0">'
    "<tr><th>Partner</th><th>Service Type</th><th>Short Code</th>"
    "<th>Go-Live Date</th><th>Ticket</th></tr>"
    f"{table_rows}"
    "</table>"
    "<p>Expected checks: three separate paragraphs above, bold phrase, "
    "unicode (it’s, café, —), and this closing block on two lines.</p>"
    "<p>Regards,<br>Alex</p>"
)

PLAIN_BODY = (
    "Hi team,\n\n"
    "This is the PLAIN-BODY variant of the same test: the server converts it "
    "to HTML, so these paragraphs must arrive separated by blank lines.\n\n"
    "Kindly confirm the line breaks survived & that <angle brackets> are shown "
    "literally.\n\n"
    "Regards,\nAlex"
)


def log(msg):
    print(msg, file=sys.stderr, flush=True)


async def sent_copy_report(bridge, subject):
    """Find the sent copy and return its HTML content."""
    script = f'''tell application "Microsoft Outlook"
    set hits to (every message of sent items whose subject is "{escape(subject)}")
    if (count of hits) is 0 then return "NOT FOUND"
    set c to content of item 1 of hits
    return c
end tell'''
    for _ in range(10):
        content = await bridge.run(script)
        if content != "NOT FOUND":
            return content
        await asyncio.sleep(2)
    return "NOT FOUND"


async def main(recipient):
    bridge = AppleScriptBridge()
    await bridge.start()
    server_mac.bridge = bridge

    log(f"Sending HTML test email to {recipient} ...")
    log(await server_mac.send_email(to=recipient, subject=SUBJECT, body="fallback", html_body=HTML_BODY))

    plain_subject = SUBJECT + " (plain body)"
    log(f"Sending plain-body test email to {recipient} ...")
    log(await server_mac.send_email(to=recipient, subject=plain_subject, body=PLAIN_BODY))

    log("Inspecting Sent Items copy for anything Outlook added client-side ...")
    content = await sent_copy_report(bridge, SUBJECT)
    if content == "NOT FOUND":
        log("  Sent copy not found yet; check Sent Items manually.")
    else:
        # Outlook wraps the fragment in <html><head>..</head><body>..</body></html>
        # and may add <tbody>; anything else would be a client-side signature.
        inner = content.split("<body>", 1)[-1].rsplit("</body>", 1)[0]
        inner = inner.replace("<tbody>", "").replace("</tbody>", "")
        log(f"  sent copy length: {len(content)}; body matches what we sent: {inner == HTML_BODY}")
        if inner != HTML_BODY:
            log("  extra content in sent copy:\n" + inner.replace(HTML_BODY, "")[:1500])
        log("  Note: a corporate signature/disclaimer added by the mail server in")
        log("  transit appears only in the RECEIVED copy, never in Sent Items.")

    log("")
    log("Now open both messages in Outlook desktop and OWA and verify:")
    log("  1. paragraphs separated by blank lines (both emails)")
    log("  2. bold phrase, clickable ticket links, bordered 5-column table (HTML email)")
    log("  3. unicode intact: curly apostrophe, café, em dash")
    log("  4. closing 'Regards,' / 'Alex' on two lines")
    log("  5. whether a signature block follows the closing (server-appended if so)")


if __name__ == "__main__":
    if len(sys.argv) != 2 or "@" not in sys.argv[1]:
        log("Usage: python tests/manual_send_test.py recipient@example.com")
        sys.exit(2)
    asyncio.run(main(sys.argv[1]))
