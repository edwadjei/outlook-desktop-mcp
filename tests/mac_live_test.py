"""
Outlook Desktop MCP - macOS live send/reply test
=================================================
Opt-in end-to-end test against the running Outlook app. Sends one seed
message to a scratch address that delivers back into this mailbox, waits
for the inbox copy, replies to it with reply_all=False and reply_all=True,
and asserts every sent item shows up in Sent Items.

Environment:
  OUTLOOK_MCP_LIVE_SCRATCH   address to send to (required; otherwise SKIP)
  OUTLOOK_MCP_LIVE_INLINE_ID message id with an inline image (optional)

Run: ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py
Requires Outlook for Mac in legacy mode with the account signed in.
"""
import sys
import os
import json
import asyncio
import time
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac, outlook_db
from outlook_desktop_mcp.outlook_db import OutlookDB

SCRATCH = os.environ.get("OUTLOOK_MCP_LIVE_SCRATCH", "").strip()
INLINE_ID = os.environ.get("OUTLOOK_MCP_LIVE_INLINE_ID", "").strip()
ROUND_TRIP = 120  # seconds allowed for an Exchange round trip


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


async def matches(folder, subject):
    """Messages in folder whose subject (ignoring Re:/Fw:) equals subject."""
    raw = await server_mac.search_emails(query=subject, folder=folder, count=20)
    rows = json.loads(raw)
    if not isinstance(rows, list):
        return []
    return [r for r in rows if r.get("subject", "").strip().lower() == subject.lower()]


async def wait_for_count(folder, subject, minimum, timeout=ROUND_TRIP):
    """Poll until at least `minimum` matching messages are in folder."""
    deadline = time.time() + timeout
    while True:
        rows = await matches(folder, subject)
        if len(rows) >= minimum or time.time() >= deadline:
            return rows
        await asyncio.sleep(3)


async def start_server():
    await server_mac.bridge.start()
    path = outlook_db.locate()
    if path and await server_mac._trust_db(OutlookDB(path)):
        server_mac.db = OutlookDB(path)
    log(f"  database: {'trusted' if server_mac.db else 'not used'}")


async def run():
    await start_server()
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    subject = f"[MCP live test {stamp}] reply regression"

    log("--- seed message ---")
    result = await server_mac.send_email(
        to=SCRATCH, subject=subject,
        body="Seed for the reply regression test. Safe to delete.\n\nBr,\nEdward.")
    check("send_email reported success", result.startswith("Email sent"), result)
    sent = await wait_for_count("sent", subject, 1)
    check("seed appears in Sent Items", len(sent) >= 1)
    received = await wait_for_count("inbox", subject, 1)
    check("seed delivered to inbox", len(received) >= 1)
    if not received:
        return
    seed_id = received[0]["entry_id"]

    log("--- replies ---")
    expected_sent = 1
    for reply_all in (False, True):
        result = await server_mac.reply_email(
            entry_id=seed_id, reply_all=reply_all,
            body=f"Reply with reply_all={reply_all}. Safe to delete.\n\nBr,\nEdward.")
        check(f"reply_all={reply_all}: reply_email reported success",
              result.startswith("Reply sent"), result)
        expected_sent += 1
        sent = await wait_for_count("sent", subject, expected_sent)
        check(f"reply_all={reply_all}: reply appears in Sent Items",
              len(sent) >= expected_sent, f"{len(sent)} in Sent Items")


def main():
    if not SCRATCH:
        log("SKIP: set OUTLOOK_MCP_LIVE_SCRATCH to run the live test")
        return
    asyncio.run(run())
    log("=" * 50)
    log(f"{passed}/{total} checks passed")
    if passed != total:
        sys.exit(1)


if __name__ == "__main__":
    main()
