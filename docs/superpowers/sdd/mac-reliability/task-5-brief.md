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

### Task 5: `save_attachment` saves inline images

**Files:**
- Modify: `src/outlook_desktop_mcp/server_mac.py` (`save_attachment`)
- Modify: `tests/mac_batch_test.py`, `tests/mac_live_test.py`

**Interfaces:**
- Consumes: `bridge.run`, `DELIM`, `escape`.
- Produces: `save_attachment` returns `{"status": "saved", "filename", "path", "bytes"}` or an error string.

- [ ] **Step 1: Failing test**

Append to `tests/mac_batch_test.py`:

```python
def test_save_attachment_uses_posix_file_and_verifies_output():
    log("--- save_attachment saves through a POSIX file reference ---")
    with tempfile.TemporaryDirectory() as d:
        class Writes(FakeBridge):
            async def run(self, script, timeout=None):
                self.scripts.append(script)
                with open(os.path.join(d, "image003.png"), "wb") as fh:
                    fh.write(b"\x89PNG fake")
                return f"image003.png{DELIM}{d}/image003.png"

        fake = Writes()
        server_mac.bridge = fake
        result = json.loads(asyncio.run(server_mac.save_attachment(entry_id="1", attachment_index=1, save_directory=d)))
        script = fake.scripts[0]
        check("only one script", len(fake.scripts) == 1)
        check("save uses POSIX file", "save a in (POSIX file savePath)" in script, script)
        check("no string-path save", "save a in savePath\n" not in script)
        check("dead placeholder script gone", "__PLACEHOLDER__" not in script)
        check("saved status", result.get("status") == "saved", str(result))
        check("path returned", result.get("path") == os.path.join(d, "image003.png"), str(result))
        check("byte count returned", result.get("bytes") == 9, str(result))

        server_mac.bridge = FakeBridge(output=f"ghost.png{DELIM}{d}/ghost.png")
        result = asyncio.run(server_mac.save_attachment(entry_id="1", attachment_index=1, save_directory=d))
        check("missing file reported as error", result.startswith("Error saving attachment") and "ghost.png" in result, result)
```

Add to `main()`. Run: expected failures on the POSIX-file shape, placeholder, `bytes`, and the missing-file case.

- [ ] **Step 2: Implement**

Replace everything in `save_attachment` from `# Use POSIX path for AppleScript` through `return f"Error saving attachment: {e}"` with:

```python
    # `save ... in` takes a file object. A POSIX path *string* is rejected
    # with -2700 for inline images (which have no `file` of their own), so
    # build the reference with `POSIX file`.
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    if attCount < {attachment_index} then return "ERROR:Only " & attCount & " attachment(s)"
    set a to item {attachment_index} of attList
    set aname to name of a
    set savePath to "{escape(save_directory)}/" & aname
    save a in (POSIX file savePath)
    return aname & "{DELIM}" & savePath
end tell'''

    try:
        raw = await bridge.run(script)
        if raw.startswith("ERROR:"):
            return raw
        parts = raw.split(DELIM)
        filename = parts[0].strip() if parts else "unknown"
        save_path = os.path.join(save_directory, filename)
        if not os.path.isfile(save_path) or os.path.getsize(save_path) == 0:
            return (f"Error saving attachment: Outlook reported success but no file "
                    f"was written at {save_path}")
        return json.dumps({
            "status": "saved",
            "filename": filename,
            "path": save_path,
            "bytes": os.path.getsize(save_path),
        }, indent=2, default=str)
    except Exception as e:
        return f"Error saving attachment: {e}"
```

Update the docstring `Returns:` to `JSON with status, filename, path and bytes, or an error.` and note in the description that inline (pasted) images are supported.

- [ ] **Step 3: Run the gate**

Expected: all pass.

- [ ] **Step 4: Live check**

Add to `tests/mac_live_test.py` `run()`, at the end:

```python
    if INLINE_ID:
        log("--- inline image save ---")
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            raw = await server_mac.save_attachment(entry_id=INLINE_ID, attachment_index=1, save_directory=d)
            try:
                saved = json.loads(raw)
            except ValueError:
                saved = {"error": raw}
            check("inline attachment saved", saved.get("status") == "saved", raw)
            check("file has bytes", saved.get("bytes", 0) > 0, raw)
```

Run: `OUTLOOK_MCP_LIVE_SCRATCH=Edward.Adjei@mtn.com OUTLOOK_MCP_LIVE_INLINE_ID=197468 ~/.mcp-venvs/outlook-desktop/bin/python tests/mac_live_test.py`. Expected: all checks pass, including the two inline ones. (Message 197468 in the inbox carries `image003.png` inline.)

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/server_mac.py tests/mac_batch_test.py tests/mac_live_test.py
git -c user.name=edwadjei -c user.email=edd.net49@gmail.com commit -m "macOS: save inline attachments via POSIX file reference"
```
