# WordMcp — Microsoft Word MCP Server

An [MCP](https://modelcontextprotocol.io) server that lets AI assistants drive **Microsoft Word for Windows** through COM automation: open documents, read and edit text, manage paragraphs and tables, set document properties and export to PDF.

> **Windows only.** A local installation of Microsoft Word is required — this server automates the real application, it does not parse `.docx` files.

---

## Requirements

| | |
|---|---|
| OS | Windows 10/11 |
| Runtime | .NET 9 SDK or runtime |
| Office | Microsoft Word 2016 or newer (desktop, not Microsoft Store version) |

## Installation

```powershell
dotnet tool install --global WordMcp.McpServer
```

The tool is then available as `mcp-word`.

```powershell
mcp-word --version
mcp-word --help
```

## Client configuration

The server speaks **stdio**.

### VS Code / GitHub Copilot

`.vscode/mcp.json`:

```json
{
  "servers": {
    "word": {
      "type": "stdio",
      "command": "mcp-word"
    }
  }
}
```

### Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "mcp-word"
    }
  }
}
```

### Copilot CLI

```powershell
copilot mcp add word --command mcp-word
```

---

## Concepts

Every operation runs inside a **session**. A session owns one Word instance and one open document, identified by a `session_id` such as `word-a1b2c3d4e5f6g`.

```
file(open|create) ──► session_id ──► text / paragraph / table / document ──► file(save) ──► file(close)
```

* Paths must be **absolute** (`C:\Users\me\Documents\report.docx`).
* Supported input formats: `.docx`, `.docm`, `.doc`, `.dotx`, `.dotm`, `.rtf`.
* The document must not be open in Word already — WordMcp needs exclusive access.
* Word runs invisibly in the background and is terminated when the session closes.

---

## Tools

Five tools, each with an `action` parameter.

### `file` — session lifecycle

| Action | Purpose |
|---|---|
| `open` | Open an existing document and start a session |
| `create` | Create a new document at `path` |
| `save` | Save the open document |
| `close` | Save (optionally) and close the session |
| `list` | List all active sessions |
| `test` | Check whether Word can be automated on this machine |

```jsonc
file(action: "open", path: "C:\\Users\\me\\Documents\\report.docx")
// → { "sessionId": "word-a1b2c3d4e5f6g", "fileName": "report.docx", ... }
```

### `text` — content

| Action | Purpose |
|---|---|
| `get` | Read the whole text or a character range (`start`, `end`, `max_length`) |
| `append` | Append text, optionally as a new paragraph |
| `find` | Find a term; returns positions and surrounding context |
| `replace` | Replace occurrences (`match_case`, `match_whole_word`, `replace_all`) |
| `format` | Apply `bold`, `italic`, `underline`, `font_name`, `font_size`, `color` to a range |

Character positions come from `get` and `find` and are Word range offsets.

### `paragraph` — structure

| Action | Purpose |
|---|---|
| `list` | List paragraphs with index, text, style, alignment and outline level |
| `add` | Append a paragraph, optionally with `style` |
| `insert` | Insert a paragraph before a given index |
| `delete` | Delete a paragraph by index |
| `set-style` | Apply a style such as `Heading 1` |
| `set-alignment` | `left`, `center`, `right` or `justify` |

Paragraph indices are **1-based**, matching Word.

### `table` — tables

| Action | Purpose |
|---|---|
| `list` | List tables with dimensions and style |
| `create` | Create a table with `rows` × `columns` |
| `read` | Read all cells of a table as a row/column matrix |
| `set-cell` | Write a single cell (`row`, `column`, `text`) |
| `add-row` | Append a row |
| `delete-row` | Delete a row |
| `set-style` | Apply a table style such as `Table Grid` |

### `document` — metadata and export

| Action | Purpose |
|---|---|
| `get-info` | Word, character, paragraph, page, table and section counts |
| `get-properties` | Title, author, subject, keywords, comments, company |
| `set-properties` | Update those built-in properties |
| `export-pdf` | Export to PDF without touching the open document |
| `save-as` | Save a copy in another format |

---

## Responses

Every tool returns JSON. Failures are reported as structured payloads, never as a transport error:

```json
{
  "success": false,
  "isError": true,
  "tool": "text",
  "action": "Replace",
  "errorType": "KeyNotFoundException",
  "errorMessage": "Session 'word-unknown' not found."
}
```

---

## Known behaviour and pitfalls

* **`document(save-as)` also saves the original.** Word has no format-changing "save a copy" API. For any target other than PDF the server calls `SaveAs2(target)` and then `SaveAs2(original)`, which persists pending changes to the original file as a side effect. Use `export-pdf` when you need a side-effect-free export.
* **Colours are hex RGB** (`#0078D4`). The server converts to the BGR value Word expects.
* **Rights-protected documents** (IRM/AIP) are rejected before Word is launched.
* **A document open in Word blocks the session** — close it first.
* **Word dialogs stall automation.** If a call times out, check for an open dialog on the desktop.
* **Merged table cells** are returned as empty strings by `table(read)`.

---

## Building from source

```powershell
git clone https://github.com/trsdn/mcp-server-word.git
cd mcp-server-word
dotnet build WordMcp.sln -c Release
dotnet test WordMcp.sln --filter "Category!=RequiresWord"
```

### Project layout

| Project | Purpose |
|---|---|
| `src/WordMcp.ComInterop` | Word COM lifecycle: STA threading, sessions, OLE message filter, file validation |
| `src/WordMcp.Core` | Command implementations and result models |
| `src/WordMcp.McpServer` | stdio MCP server exposing the five tools |
| `tests/WordMcp.Core.Tests` | Unit tests, no Word required |
| `tests/WordMcp.McpServer.Tests` | Tool-layer unit tests, no Word required |

Tests that need a real Word installation are marked `[Trait("Category", "RequiresWord")]` and are excluded in CI.

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| `Word is not installed or not registered for COM` | Install Word desktop; the Microsoft Store version cannot be automated |
| `Could not load file or assembly 'office'` | `office.dll` was not found in the GAC — reinstall or repair Office |
| Operation times out | A Word dialog is waiting for input; close it and retry |
| `The file is already open in Word` | Close the document in the Word UI |

## License

MIT — see [LICENSE](LICENSE).
