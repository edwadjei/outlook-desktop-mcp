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

### Task 4: `search_emails` by recipient and sender; document the preview limit

**Files:**
- Modify: `src/outlook_desktop_mcp/outlook_db.py` (`search_messages`)
- Modify: `src/outlook_desktop_mcp/server_mac.py` (`search_emails`)
- Modify: `tests/mac_db_test.py` (fixture columns, tests), `README.md`

**Interfaces:**
- Consumes: `OutlookDB._query`, `_like_pattern`, `_row_to_summary`; `server_mac._db_folder_id`.
- Produces: `OutlookDB.search_messages(folder_id, query, count, recipient="", sender="")`; `server_mac.search_emails(query="", folder="inbox", count=10, recipient="", sender="")`.

- [ ] **Step 1: Extend the fixture**

In `tests/mac_db_test.py` `build_fixture`, add three columns to `CREATE TABLE Mail` after `Message_Hidden INTEGER`:

```sql
            Message_ToRecipientAddressList TEXT,
            Message_CCRecipientAddressList TEXT,
            Message_DisplayTo TEXT
```

Change the insert to `INSERT INTO Mail VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)` and make it pad short tuples:

```python
    con.executemany("INSERT INTO Mail VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    [tuple(row) + (None,) * (14 - len(row)) for row in mail])
```

Update the `MAIL` comment line to `# id, folder, subject, sender name, sender addr, preview, read, att, time, del, hidden[, to addrs, cc addrs, display to]`.

- [ ] **Step 2: Failing tests**

Append:

```python
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
    server_mac.db = None
```

Add both to `main()`. Run: expected `TypeError` on the `recipient` keyword.

- [ ] **Step 3: Implement `search_messages`**

Replace the method:

```python
    def search_messages(self, folder_id: int, query: str, count: int,
                        recipient: str = "", sender: str = "") -> list[dict]:
        """Case-insensitive substring search.

        `query` matches subject, sender name, sender address and the preview
        (first 255 characters of the body). `recipient` matches the To and
        CC address lists and the display-To names. `sender` matches sender
        name and address. All given criteria must match.
        """
        clauses: list[str] = []
        params: list = [folder_id]
        query, recipient, sender = query.strip(), recipient.strip(), sender.strip()
        if query:
            p = _like_pattern(query)
            clauses.append(
                "(Message_NormalizedSubject LIKE ? ESCAPE '\\' OR "
                "Message_SenderList LIKE ? ESCAPE '\\' OR "
                "Message_SenderAddressList LIKE ? ESCAPE '\\' OR "
                "Message_Preview LIKE ? ESCAPE '\\')"
            )
            params += [p] * 4
        if recipient:
            p = _like_pattern(recipient)
            clauses.append(
                "(IFNULL(Message_ToRecipientAddressList, '') LIKE ? ESCAPE '\\' OR "
                "IFNULL(Message_CCRecipientAddressList, '') LIKE ? ESCAPE '\\' OR "
                "IFNULL(Message_DisplayTo, '') LIKE ? ESCAPE '\\')"
            )
            params += [p] * 3
        if sender:
            p = _like_pattern(sender)
            clauses.append(
                "(IFNULL(Message_SenderList, '') LIKE ? ESCAPE '\\' OR "
                "IFNULL(Message_SenderAddressList, '') LIKE ? ESCAPE '\\')"
            )
            params += [p] * 2
        if not clauses:
            return []
        params.append(max(0, int(count)))
        rows = self._query(
            f"SELECT {_MESSAGE_COLUMNS} FROM Mail "
            f"WHERE Record_FolderID = ? AND {_LIVE_ROWS} AND " + " AND ".join(clauses) +
            " ORDER BY Message_TimeReceived DESC, Record_RecordID DESC LIMIT ?",
            params,
        )
        return [self._row_to_summary(r) for r in rows]
```

- [ ] **Step 4: Implement `search_emails`**

Change the signature to

```python
async def search_emails(
    query: str = "",
    folder: str = "inbox",
    count: int = 10,
    recipient: str = "",
    sender: str = "",
) -> str:
```

Replace the docstring with:

```python
    """Search for emails in Outlook.

    On legacy Outlook for Mac the search runs against Outlook's local
    message index. `query` matches the subject, the sender name and
    address, and the preview, which is only the FIRST 255 CHARACTERS of
    the body. Full bodies are not indexed: a keyword that appears deeper
    in a message is not found. To find pending requests reliably, filter
    by `recipient` (and `sender`) and read candidates with read_email.
    Reply/forward prefixes ("Re:", "FW:") are ignored.

    If the index is unavailable, `query` falls back to AppleScript
    filtering on subject only, and `recipient`/`sender` return an error
    rather than silently searching without them.

    Args:
        query: Substring for subject, sender and preview. May be empty
            when recipient or sender is given.
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.
        recipient: Substring matched against To and CC addresses and the
            displayed To names, e.g. "edward" or "isdemand".
        sender: Substring matched against the sender name and address.

    Returns:
        JSON array of matching email summaries, newest first, or an error.
    """
```

Replace the start of the body (through the database attempt) with:

```python
    if not (query.strip() or recipient.strip() or sender.strip()):
        return json.dumps({"error": "Provide at least one of query, recipient, sender"})
    fid = await _db_folder_id(folder)
    if fid is not None:
        try:
            rows = await asyncio.to_thread(db.search_messages, fid, query, count, recipient, sender)
            return json.dumps(rows, indent=2, default=str)
        except OutlookDBError as e:
            logger.warning("Outlook database search failed; using AppleScript: %s", e)
    if recipient.strip() or sender.strip():
        return json.dumps({"error": "recipient and sender filters need Outlook's message index, "
                                    "which is unavailable right now; retry, or search with query only"})
```

The rest of the function (the AppleScript scripts) stays as it is.

- [ ] **Step 5: Run the gate**

Expected: all pass, including the existing `test_search_matches_subject_sender_preview`.

- [ ] **Step 6: README**

Change the `search_emails` row in the Email table to:
`| \`search_emails\` | yes | yes | Search by keyword; macOS also filters by \`recipient\` and \`sender\`. Keyword matching covers subject, sender and the first 255 characters of the body only |`.
In "#### Fast list and search on macOS" add a bullet:

> - `search_emails` keyword matching covers the subject, sender and the message preview (first 255 characters). Bodies are not indexed, so a keyword sweep can miss requests whose key word sits lower in the message. Combine it with `recipient="<your name or address>"` to list everything addressed to you, then `read_email` the candidates.

- [ ] **Step 7: Commit**

```bash
git add src/outlook_desktop_mcp/outlook_db.py src/outlook_desktop_mcp/server_mac.py tests/mac_db_test.py README.md
git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit -m "macOS: search_emails by recipient and sender; document preview limit"
```
