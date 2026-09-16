# Standing rules for this task (read first)

You are the implementer for exactly one task of the delivery plan. This brief is your
requirements; use its values verbatim. Do not read other tasks' briefs or the whole plan.

- Work only in the repo checkout you were pointed at; never `cd` elsewhere; never touch git config.
- Test-driven: write the failing test first, run it and record the failing output (RED), then the
  minimal code, then the passing output (GREEN). Both go in your report.
- The gate (below) must be green before the commit. No stubs, no placeholders, no mock data in
  shipped code. Never weaken a test to make it pass; if a test in the brief is wrong, say so in the
  report under DONE_WITH_CONCERNS and fix the test to what the brief clearly intends.
- One commit for the task, exactly the subject given, staged by explicit path. No AI references,
  no Co-Authored-By trailer. Never push, never merge, never open a PR.
- Live probes: check `ps -axo pid,etime,command | grep "[o]sascript -e"` first; wait for other
  scripts to finish; never kill them. Outlook must be running in legacy mode.
- Write the report file named in your dispatch with: status (DONE | DONE_WITH_CONCERNS |
  NEEDS_CONTEXT | BLOCKED), RED and GREEN command+output, the gate output, the commit hash, and
  any concern or deviation. Then return a three-line status.

## Global Constraints

- Python to run everything: `~/.mcp-venvs/outlook-desktop/bin/python` (the `mcp` package is only installed there).
- Gate, green before every commit:
  `~/.mcp-venvs/outlook-desktop/bin/python tests/mac_batch_test.py && ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_db_test.py && ~/.mcp-venvs/outlook-desktop/bin/pip wheel -q --no-deps -w /tmp/odm-wheel "$PWD" && rm -rf /tmp/odm-wheel`
- Windows code (`server.py`, `com_bridge.py`) is not touched.
- Commit author: `git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit ...`. Subject as named per task. No AI references, no `Co-Authored-By` trailer. Stage by explicit path only.
- Outlook runs one AppleScript at a time: before any live probe run `ps -axo pid,etime,command | grep "[o]sascript -e"` and wait for it to clear. Never kill another session's script.
- Live sends go only to `Edward.Adjei@mtn.com` (the signed-in test account; self-sends are allowed). Run the live test at most once per task.
- AppleScript record fields are read from a local variable (`set senderRec to sender of m`, then `address of senderRec`), never through a chained specifier.
- Existing checks must keep passing: 118 in `tests/mac_batch_test.py`, 68 in `tests/mac_db_test.py` before Task 1.

---

### Task 3: Confirm the Sent Items copy after send and reply; stable sent ordering

**Files:**
- Modify: `src/outlook_desktop_mcp/outlook_db.py` (`list_messages` ordering, new `normalize_subject`, new `find_sent_copy`)
- Modify: `src/outlook_desktop_mcp/server_mac.py` (new `_sent_copy_id`, `_confirmation_suffix`; `send_email`, `reply_email`, `list_emails` docstring)
- Modify: `tests/mac_db_test.py`, `tests/mac_batch_test.py`, `tests/mac_live_test.py`, `README.md`

**Interfaces:**
- Consumes: `server_mac._ensure_db()`, `server_mac._folder_ref(folder)` (returns `mail folder id N` or a keyword), `OutlookDB.resolve_folder`, `OutlookDB._row_to_summary`.
- Produces: `outlook_db.normalize_subject(str) -> str`; `OutlookDB.find_sent_copy(folder_id, subject, since) -> dict | None`; `server_mac._sent_copy_id(subject, since) -> str | None`; `server_mac._SENT_CONFIRM_TIMEOUT = 10.0`, `_SENT_CONFIRM_INTERVAL = 0.5`. Send/reply confirmations end with ` (Sent Items id N)` or ` (Sent Items copy not visible yet; verify with search_emails)`.

- [ ] **Step 1: Failing db tests**

Append to `tests/mac_db_test.py`:

```python
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
```

Add all four to `main()`.

- [ ] **Step 2: Failing batch test (AppleScript fallback for the lookup)**

Append to `tests/mac_batch_test.py`:

```python
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
```

Add to `main()`. In BOTH `tests/mac_batch_test.py` and `tests/mac_db_test.py`, make the first line of `main()`:

```python
    server_mac._SENT_CONFIRM_TIMEOUT = 0.0  # unit tests never wait for Outlook's Sent Items write
```

(`test_server_send_email_confirms_sent_copy` raises it to 0.6 for its timing check and restores it.) With a zero timeout the lookup still runs exactly once, so shape checks on the lookup script keep working. Any existing send/reply check that asserts `len(fake.scripts) == 1` must become `== 2` (send script, then the Sent Items lookup); `fake.scripts[0]` is still the send script.

Run both suites; expected failures: `AttributeError` on `normalize_subject`, `find_sent_copy`, and unmatched suffixes.

- [ ] **Step 3: Implement in `outlook_db.py`**

Add `import re` and:

```python
_PREFIX_RE = re.compile(r"^\s*(?:(?:re|fw|fwd|aw|sv|wg)\s*:\s*)+", re.IGNORECASE)


def normalize_subject(subject: str) -> str:
    """Subject without leading reply/forward prefixes, as Outlook stores it."""
    return _PREFIX_RE.sub("", subject or "").strip()
```

Change `list_messages` ordering to `"ORDER BY Message_TimeReceived DESC, Record_RecordID DESC LIMIT ?"` and add:

```python
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
```

- [ ] **Step 4: Implement in `server_mac.py`**

Add after `_TRUST_DRIFT_FRACTION`:

```python
# After `send`, Outlook parks the message in the Outbox for a few seconds
# and then writes a new Sent Items record with a new id. Send tools poll
# for that record so their confirmation can name it.
_SENT_CONFIRM_TIMEOUT = 10.0
_SENT_CONFIRM_INTERVAL = 0.5


async def _sent_copy_id(subject: str, since: int) -> str | None:
    """Id of the Sent Items copy of a message sent at `since`, or None if not visible in time."""
    handle = await _ensure_db()
    deadline = time.monotonic() + _SENT_CONFIRM_TIMEOUT
    while True:
        try:
            if handle is not None:
                fid = await asyncio.to_thread(handle.resolve_folder, "sent")
                row = None
                if fid is not None:
                    row = await asyncio.to_thread(handle.find_sent_copy, fid, subject, since)
                if row is not None:
                    return row["entry_id"]
            else:
                folder_ref = await _folder_ref("sent")
                raw = await bridge.run(f'''tell application "Microsoft Outlook"
    set f to {folder_ref}
    set n to count of messages of f
    if n > 20 then set n to 20
    repeat with i from 1 to n
        set m to message i of f
        if subject of m is "{escape(subject)}" then return (id of m as text)
    end repeat
    return ""
end tell''')
                if raw.strip():
                    return raw.strip()
        except (OutlookDBError, RuntimeError) as e:
            logger.warning("Sent Items lookup failed: %s", e)
            return None
        if time.monotonic() >= deadline:
            return None
        await asyncio.sleep(_SENT_CONFIRM_INTERVAL)


def _confirmation_suffix(sent_id: str | None) -> str:
    if sent_id:
        return f" (Sent Items id {sent_id})"
    return " (Sent Items copy not visible yet; verify with search_emails)"
```

In `send_email`, replace the `try:` block with:

```python
    try:
        since = int(time.time())
        await bridge.run(script)
        sent_id = await _sent_copy_id(subject, since)
        return f"Email sent: '{subject}' to {to}{_confirmation_suffix(sent_id)}"
    except Exception as e:
        return f"Error sending email: {e}"
```

In `reply_email`, replace the `try:` block with:

```python
    try:
        since = int(time.time())
        subject = await bridge.run(script)
        sent_id = await _sent_copy_id(subject, since)
        return f"Reply sent to '{subject}' (reply_all={reply_all}){_confirmation_suffix(sent_id)}"
    except Exception as e:
        return f"Error replying to email: {e}"
```

Update the `Returns:` lines of both docstrings to mention the Sent Items id. In the `list_emails` docstring add to the `folder` argument text:

```
            Note: a message you just sent spends a few seconds in "outbox"
            before it appears in "sent"; the id in the send tool's
            confirmation, or search_emails, is the reliable check.
```

- [ ] **Step 5: Run the gate**

Expected: all pass. The `test_send_email_looks_up_sent_copy_without_db` check `'subject of m is "Hi"'` must match the generated script exactly.

- [ ] **Step 6: Extend the live test**

In `tests/mac_live_test.py` `run()`, after each `check(... reported success ...)` for the seed and for both replies add:

```python
        import re
        m = re.search(r"\(Sent Items id (\d+)\)", result)
        check("confirmation names the Sent Items id", m is not None, result)
        if m:
            copy = json.loads(await server_mac.read_email(entry_id=m.group(1)))
            check("Sent Items id readable", copy.get("subject", "").lower().endswith(subject.lower()), str(copy)[:200])
```

(for the seed use `result` from `send_email`; for replies the reply result). Run the live test once: `OUTLOOK_MCP_LIVE_SCRATCH=Edward.Adjei@mtn.com ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py`. Expected: all checks pass; record output.

- [ ] **Step 7: README**

In "#### Fast list and search on macOS" add a bullet:

> - A message you just sent sits in the Outbox for a few seconds, then Outlook writes a new Sent Items record with a new id. `send_email` and `reply_email` wait up to 10 s for that record and name it in their confirmation (`(Sent Items id N)`); `list_emails(folder="sent")` called inside that window will not show it yet. Verify sends by the confirmation id or `search_emails`, not by the top of `list_emails`.

- [ ] **Step 8: Commit**

```bash
git add src/outlook_desktop_mcp/outlook_db.py src/outlook_desktop_mcp/server_mac.py tests/mac_db_test.py tests/mac_batch_test.py tests/mac_live_test.py README.md
git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit -m "macOS: confirm Sent Items copy after send and reply; stable sent ordering"
```
