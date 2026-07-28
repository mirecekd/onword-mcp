# onword-mcp

MCP server for live, surgical editing of the document currently open in Microsoft Word. Works on the running Word instance via COM automation (pywin32) - large-document friendly and safe for Teams/SharePoint co-authoring.

<div align="center">

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/mirecekdg) [!["PayPal.me"](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/donate/?business=LJ5ZF7Q9KMTRW&no_recurring=0&currency_code=USD)

</div>

## Why

Existing Word MCP servers either work on closed `.docx` files (useless for documents opened from Teams) or rewrite the whole `doc.Content.Text` at once, which breaks co-authoring sync and freezes Word on large documents.

onword-mcp mirrors the approach of the official Claude for Word add-in: it never sends the whole document anywhere. The LLM first reads a token-cheap outline (paragraph indexes, styles, list levels, previews), then reads or edits only the specific paragraphs it needs - one `Range` at a time, with `ScreenUpdating` off during writes.

## Requirements

- Windows (win32) with Microsoft Word installed and **running with a document open**
- [uv](https://docs.astral.sh/uv/getting-started/installation/) (`winget install astral-sh.uv` or `powershell -c "irm https://astral.sh/uv/install.ps1 | iex"`)

No Python installation needed - `uvx` downloads Python and all dependencies (fastmcp, pywin32) automatically on first run.

## Quick start on Windows: one desktop icon

If you want Word, the MCP server and (optionally) a reverse SSH tunnel started
by a single double-click, use the launcher in [`launcher/`](launcher/):

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File C:\tools\onword-mcp\launcher\Install-Shortcut.ps1 -SshTarget dev.domain.com -WithStop -WithStatus
```

This creates an **onword** icon on the Desktop that starts Word, the MCP server
on port 18347 and `ssh -N -R 18347:127.0.0.1:18347 <your-host>`, keeps the
tunnel alive, and shuts everything down when Word is closed. Host name and
ports live in a gitignored `launcher/onword.env` (template:
`onword.env.example`), never in the scripts. It can also auto-start with Word
via the optional `AutoExec.bas` macro. See
[launcher/README.md](launcher/README.md) for details.

To see whether it is running, and to stop it, run the two helper scripts
directly (the desktop icons are only shortcuts to them):

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File C:\tools\onword-mcp\launcher\onword-status.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File C:\tools\onword-mcp\launcher\onword-stop.ps1
```

## Running


### From a copied directory (no packaging needed)


Copy this whole directory to the Windows machine (e.g. `C:\tools\onword-mcp`) and run:

```bat
uvx --from C:\tools\onword-mcp onword-mcp --transport streamable-http --port 18347
```

That is all. `uvx` builds the package from the directory, creates an isolated environment with `pywin32` and `fastmcp`, and starts the server. Subsequent starts are instant (cached).

### One-line from GitHub (nothing to copy)

```bat
uvx --from git+https://github.com/mirecekd/onword-mcp onword-mcp --transport streamable-http
```

### One-line from PyPI (after publishing)
```bat
uvx onword-mcp --transport streamable-http
```

To publish: `uv build && uv publish` (needs a PyPI token).

### Transports and configuration

| Option | Env var | Default | Values |
|---|---|---|---|
| `--transport` | `MCP_TRANSPORT` | `stdio` | `stdio`, `streamable-http`, `http`, `sse` |
| `--port` | `MCP_PORT` | `18347` | any port |
| `--host` | `MCP_HOST` | `127.0.0.1` | bind address |

Env-var style (as used by some MCP clients on Windows):

```bat
cmd /c "set MCP_TRANSPORT=streamable-http&& set MCP_PORT=18347&& uvx --from C:\tools\onword-mcp onword-mcp"
```

With `streamable-http`, the MCP endpoint is `http://127.0.0.1:18347/mcp`.

## Troubleshooting

### `error: Failed to spawn: 'onword-mcp' ... Access is denied. (os error 5)`

The package built and installed fine - what failed is executing the small
console-script trampoline `onword-mcp.exe` that uv generates in its cache
(`%LOCALAPPDATA%\uv\cache\...`). On corporate Windows machines this freshly
created, unsigned exe is typically blocked or quarantined by the antivirus
(Defender, Cortex, CrowdStrike, ...) or by an AppLocker/SRP policy that
forbids running executables from the user profile.

Workarounds, in order of preference:

1. **Run as a module instead of the exe** (no trampoline exe involved):

   ```bat
   uvx --from C:\tools\onword-mcp python -m onword_mcp --transport streamable-http --port 18347
   ```

   Note: uvx caches the built environment by package version. If you updated
   the sources and get `No module named onword_mcp.__main__`, force a fresh
   build once with `uvx --no-cache --from ... python -m onword_mcp ...`.

2. **Add an AV/Defender exclusion** for the uv cache directory
   (`%LOCALAPPDATA%\uv`), or ask IT to allowlist it, then retry. You can
   also clear a possibly quarantined/corrupted cached build first:

   ```bat
   uv cache clean onword-mcp
   ```

3. **Move the uv cache** somewhere the policy allows executing from:

   ```bat
   set UV_CACHE_DIR=C:\tools\uv-cache
   uvx --from C:\tools\onword-mcp onword-mcp --transport streamable-http
   ```

4. **Use a plain venv** (no uv tool cache at all):

   ```bat
   cd C:\tools\onword-mcp
   uv venv && uv pip install -e .
   .venv\Scripts\python.exe -m onword_mcp --transport streamable-http --port 18347
   ```

## MCP client configuration

### Claude Desktop / Cline - stdio (client starts the server itself)

```json
{
  "mcpServers": {
    "onword": {
      "command": "uvx",
      "args": ["--from", "C:\\tools\\onword-mcp", "onword-mcp"]
    }
  }
}
```

### Claude Desktop / Cline - streamable-http (server already running)

Start the server first (see above), then:

```json
{
  "mcpServers": {
    "onword": {
      "type": "streamableHttp",
      "url": "http://127.0.0.1:18347/mcp"
    }
  }
}
```

## Tools

### Reading (token-cheap orientation)

| Tool | Purpose |
|---|---|
| `get_document_info` | name, path, pages, paragraph count, track changes state |
| `get_document_outline(start_index, limit)` | paginated structure: index, style, list level, preview |
| `read_paragraphs(start_index, count)` | full text of a paragraph block |
| `find_text(query)` | locate text, returns paragraph indexes + page numbers |
| `get_page_paragraphs(page_number)` | paragraphs on a given page ("the bullet on page 10") |
| `get_selection` | user's current cursor/selection |

### Text editing (surgical, co-authoring safe)

| Tool | Purpose |
|---|---|
| `replace_paragraph(index, new_text)` | replace one paragraph, keeps style |
| `insert_after_paragraph(index, text)` / `insert_before_paragraph` | insert new content |
| `append_to_document(text)` | append at the end without touching the rest |
| `delete_paragraphs(start_index, count)` | delete a block |
| `replace_text_in_paragraph(index, find, replace)` | substring replace inside one paragraph |
| `insert_at_selection(text)` | insert at user's cursor |
| `set_track_changes(enabled)` | edits become reviewable revisions |
| `save_document` | save / trigger co-authoring sync |

### Formatting

| Tool | Purpose |
|---|---|
| `get_paragraph_formatting(index)` | style, alignment, indents, list info, font |
| `list_document_styles` | available paragraph style names |
| `set_paragraph_style(index, style_name)` | apply Heading 1, Normal, List Bullet... |
| `list_indent(index)` / `list_outdent(index)` | move bullet right/left (Tab / Shift+Tab) |
| `set_list_level(index, level)` | exact list level 1-9 |
| `convert_to_list(index, count, numbered)` / `remove_list` | make/remove bullets or numbering |
| `set_paragraph_indent(index, left_cm, first_line_cm)` | indents in centimeters |
| `set_paragraph_alignment(index, alignment)` | left/center/right/justify |
| `format_text(index, find, bold, italic, ...)` | character formatting of a substring |

## Example workflow

> "There is a wrong bullet on page 10, move it one level right."

1. LLM calls `get_page_paragraphs(10)` - sees `[132] (List Paragraph) [list:bullet lvl:1]: Wrong bullet text...`
2. LLM calls `list_indent(132)` - bullet moves one level right, live in the open Word window.

## Architecture notes

- COM objects are never cached between tool calls. Each call does `CoInitialize()` + `GetActiveObject("Word.Application")` - cheap (running-object-table lookup) and avoids all cross-thread COM marshalling issues with the async HTTP server.
- All writes wrap the change in `ScreenUpdating = False` / `finally: True`.
- No whole-document reads or writes anywhere - everything is indexed by paragraph.

## Support

<div align="center">

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/mirecekdg) [!["PayPal.me"](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/donate/?business=LJ5ZF7Q9KMTRW&no_recurring=0&currency_code=USD)

</div>

## License

MIT (c) 2026 Miroslav Dvorak