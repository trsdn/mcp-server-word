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

To update or remove it later:

```powershell
dotnet tool update --global WordMcp.McpServer
dotnet tool uninstall --global WordMcp.McpServer
```

To run an unreleased build instead, pack it locally and install from the output folder:

```powershell
dotnet pack src\WordMcp.McpServer\WordMcp.McpServer.csproj -c Release -o artifacts
dotnet tool install --global --add-source .\artifacts WordMcp.McpServer
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

### The session service

Sessions normally live inside the MCP server process and disappear with it. `WordMcp.Service.exe`
is an optional background daemon that holds them instead, so a session survives a restarted client
and can be shared by several of them:

```
WordMcp.Service.exe --daemon [--idle-minutes 30]   # listen until idle or stopped
WordMcp.Service.exe --status                       # what is it doing?
WordMcp.Service.exe --stop                         # save open documents and exit
```

Starting it by hand is rarely necessary — a client configured to use it starts it on demand. The
pipe it listens on embeds your SID and is ACL'd to it, so sessions are never shared between
accounts. It exits on its own once no session has been open for the idle timeout.

---

## Tools

Fifteen tools, each with an `action` parameter.

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

### `image` — pictures

| Action | Purpose |
|---|---|
| `list` | List inline images with index, size, alt text and link state |
| `insert` | Insert a picture, optionally with `width`, `height`, `caption` and `alt_text` |
| `resize` | Resize by `width`/`height` or by `scale_percent` |
| `replace` | Swap the picture behind an index, keeping its size by default |
| `delete` | Delete an image by index |
| `set-alt-text` | Set the alternative text for accessibility |

### `field` — fields and tables of contents

| Action | Purpose |
|---|---|
| `list` | List all fields with index, type and field code |
| `insert-toc` | Insert a table of contents (`upper_heading_level`, `lower_heading_level`) |
| `update-toc` | Recalculate every table of contents |
| `update-all` | Update all fields, including those in headers and footers |
| `insert-page-number` | Add a page number to the header or footer |

### `section` — sections and page setup

| Action | Purpose |
|---|---|
| `list` | List all sections with start type, margins, page size and orientation |
| `add` | Insert a section break (`start_type`: `next-page`, `continuous`, `even-page`, `odd-page`) |
| `page-setup` | Set margins, `orientation` and `paper_size` for one section or the whole document |

### `header-footer` — headers and footers

| Action | Purpose |
|---|---|
| `get` | Read the header or footer of one section or of all sections |
| `set` | Write text, optionally with an `alignment` |
| `clear` | Empty the header or footer |

`kind` selects `header` or `footer`, `type` selects `primary`, `first-page` or `even-pages`.

### `style` — styles

| Action | Purpose |
|---|---|
| `list` | List styles; by default only the ones the document uses |
| `create` | Add a custom style, optionally based on an existing one |
| `modify` | Change font and paragraph formatting of a style |
| `delete` | Remove a custom style |

`style_type` selects `paragraph`, `character`, `table` or `list`. Pass `in_use_only: false` to
`list` for the full set, which is over 370 entries on a localized Word.

```jsonc
style(action: "create", session_id: "...", name: "Callout", base_style: "Normal")
style(action: "modify", session_id: "...", name: "Callout",
      font_size: 11, bold: true, color: "#C00000", space_after: 12)
```

---

### `list` — bullets and numbering

| Action | Purpose |
|---|---|
| `get` | Report the list formatting of the paragraphs, including the rendered bullet or number |
| `apply` | Format a paragraph range as a `bullet`, `number` or `outline-number` list |
| `set-level` | Set the list level of a paragraph range (1–9) |
| `restart` | Start the numbering over at a paragraph |
| `remove` | Strip the list formatting |

Omitting `end_index` applies the action to `start_index` alone.

```jsonc
list(action: "apply", session_id: "...", start_index: 2, end_index: 5, list_type: "number")
list(action: "set-level", session_id: "...", start_index: 3, end_index: 4, level: 2)
list(action: "restart", session_id: "...", start_index: 6)
```

---

### `comment` — review notes

| Action | Purpose |
|---|---|
| `list` | List the comments with author, date, text and the commented-on text |
| `add` | Attach a comment to a paragraph or to a phrase inside it |
| `resolve` | Mark a comment as done, or reopen it |
| `delete` | Remove a comment |

`add` comments the whole paragraph unless `anchor_text` names a phrase inside it. Indexes shift
after `delete`, so list again before deleting a second comment.

```jsonc
comment(action: "add", session_id: "...", paragraph_index: 4,
        text: "Source?", anchor_text: "fifteen percent")
comment(action: "list", session_id: "...", unresolved_only: true)
```

---

### `revision` — tracked changes

| Action | Purpose |
|---|---|
| `list` | List the tracked changes and report whether tracking is on |
| `accept` | Accept one revision, or all of them |
| `reject` | Reject one revision, or all of them |
| `set-tracking` | Turn change tracking on or off |

Omitting `index` on `accept`/`reject` handles the whole document, headers and footers included.

```jsonc
revision(action: "set-tracking", session_id: "...", enabled: true)
revision(action: "accept", session_id: "...")
```

---

### `bookmark` — stable references

| Action | Purpose |
|---|---|
| `list` | Bookmarks with name, paragraph index and a preview of the marked text |
| `add` | Bookmark a paragraph, a paragraph range or a phrase inside a paragraph |
| `get-text` | Read the full bookmarked text |
| `delete` | Remove a bookmark; the text stays |

Names must start with a letter and may only contain letters, digits and underscores. Bookmarks
survive edits elsewhere in the document, which makes them the reliable way to refer back to a
passage once paragraph indexes have shifted.

```jsonc
bookmark(action: "add", session_id: "...", name: "Intro", paragraph_index: 2)
bookmark(action: "add", session_id: "...", name: "Growth",
         paragraph_index: 4, anchor_text: "fifteen percent")
bookmark(action: "get-text", session_id: "...", name: "Intro")
```

---

### `screenshot` — see the page

| Action | Purpose |
|---|---|
| `page` | Render a page as a PNG |

Layout questions — page breaks, table widths, image placement, header positions — are far easier
to answer from the rendered page than from measurements. The PNG is written to a file and the path
returned; `include_image: true` additionally returns it inline as base64, which is only worth the
context when the image is going to be looked at.

`dpi` defaults to 150. Use 96 for a quick layout check and 300 for something close to print.

```jsonc
screenshot(action: "page", session_id: "...", page: 2)
screenshot(action: "page", session_id: "...", page: 1,
           output_path: "C:/temp/page1.png", dpi: 300, include_image: true)
```

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
* **Style names are English.** Built-in styles (`Heading 1`, `Title`, `Table Grid`, …) are translated to Word's language-independent style ids, so they work on localized installations. Any other name is passed to Word as-is, which is how custom and localized styles are addressed. Note that Word *reports* styles under their localized name (`Überschrift 1` on a German install), which is why `style(list)` returns both `name` and `english_name` — send `english_name` back when it is present.
* **Built-in styles cannot be deleted.** `style(delete)` rejects them with a clear message instead of passing Word's generic COM error through. A custom style that is still applied to a paragraph cannot be deleted either; set those paragraphs to another style first.
* **New documents are written directly, not via Word.** `file(create)` writes an empty `.docx`/`.docm` package itself and then opens it. Creating documents through Word instead is unreliable on machines signed in to Microsoft 365, because AutoSave claims the new document for OneDrive and silently ignores the requested local path.
* **Merged table cells** are returned as empty strings by `table(read)`.
* **Image sizes are in points, not pixels** (72 pt = 1 inch). `image(insert)` and `image(resize)` keep the aspect ratio unless `lock_aspect_ratio` is set to `false`, so passing only `width` scales the height along with it.
* **`image` only covers inline pictures.** Floating shapes, text boxes and charts are left untouched and do not appear in `image(list)`, so their presence does not shift image indexes.
* **A table of contents only lists heading paragraphs.** `field(insert-toc)` returns `entry_count: 0` for a document without heading styles — apply `Heading 1`/`Heading 2` via `paragraph(add|set-style)` first, then run `field(update-toc)`.
* **`field(update-all)` also walks headers and footers.** Word's `Document.Fields` covers the body only, which is why page numbers would otherwise never refresh.
* **`image(insert)` with a caption uses Word's caption numbering**, so the caption reads `Figure 1 <your text>` (localized on non-English installations) and participates in a table of figures.
* **All measurements are in points**, including page margins (1 cm = 28.35 pt, 1 inch = 72 pt).
* **`section(page-setup)` applies `paper_size` before the margins**, because changing the paper size resets them in Word. Without a `section_index` the setup is applied to every section.
* **Headers and footers are inherited between sections.** A new section shows the previous section's header until something is written to it. `header-footer(set)` with a `section_index` breaks that link automatically, so section 1 keeps its own text.
* **`first-page` and `even-pages` headers need a section switch.** `header-footer(set)` turns on `DifferentFirstPage` respectively `DifferentOddEvenPages` for you — without it Word stores the text but never renders it.
* **`list(apply)` starts a new list by default.** `continue_previous_list` is off, because picking up the numbering of an unrelated earlier list is rarely what was meant. Two numbered lists separated by plain paragraphs stay independent; use `list(restart)` when Word merges them anyway.
* **Only outline-number lists render distinct levels.** `list(set-level)` works on any list, but a plain `bullet` or `number` list shows the same marker on every level — the paragraphs are merely indented.
* **`comment(resolve)` often fails on Microsoft 365.** Word's modern comments treat every comment added through the API as an unposted draft, and a draft cannot be marked as done. The server reports that as a clear message; delete the comment instead. `comment(list)` returns `resolved: null` on installations that do not expose the state at all.
* **Comment and revision indexes shift.** Deleting a comment or accepting a single revision renumbers everything after it, so run `list` again between two such calls instead of reusing the old indexes.
* **`revision(accept|reject)` without an index also walks headers and footers.** Word's `Document.AcceptAllRevisions()` covers the body only, the same gap as with `field(update-all)`.
* **Tracked changes are recorded only while tracking is on.** `revision(set-tracking)` does not apply retroactively — turn it on before the edits you want recorded.
* **Bookmark names are restricted by Word.** They must start with a letter, may contain only letters, digits and underscores, and are limited to 40 characters. Spaces, hyphens, dots and non-ASCII letters are rejected before the call reaches Word, which would otherwise fail with a generic COM error.
* **Bookmarks are the stable way to refer to a passage.** Paragraph indexes shift with every insertion, bookmarks do not. Bookmark a passage once and use `bookmark(get-text)` to re-read it later.
* **`bookmark(add)` on a paragraph excludes the paragraph mark**, so `get-text` returns the text without a trailing newline. A bookmark over several paragraphs keeps the marks in between.
* **`screenshot(page)` renders through a PDF.** Word has no API that returns a page as an image, so the server exports the single page with `ExportAsFixedFormat` and rasterizes it. Unsaved changes are included, and the temporary PDF is deleted afterwards.
* **Page numbers come from a fresh repagination.** A document that was only ever edited through automation reports a stale page count, so `screenshot` repaginates first. That also means the page count reflects the current layout, not the one at open time.

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
| `src/WordMcp.Core` | Command interfaces, command implementations and result models |
| `src/WordMcp.Generators.Shared` | Source files shared by the generators; not a project of its own |
| `src/WordMcp.Generators.Mcp` | Roslyn source generator that emits the MCP tool classes |
| `src/WordMcp.McpServer` | stdio MCP server exposing the fifteen tools |
| `tests/WordMcp.Core.Tests` | Unit tests plus integration tests against a real Word |
| `tests/WordMcp.McpServer.Tests` | Tool-layer unit tests, no Word required |

### Generated tool layer

Fourteen of the fifteen tools are generated at build time. The command interfaces in
`src/WordMcp.Core/Commands` are the single source of truth for the wire contract:

- `[ServiceCategory("section", "Section")]` names the tool class, `WordSectionTool`.
- `[McpTool("section", Title = ..., Description = ...)]` supplies the tool name and the prompt the model reads.
- `[ServiceAction("page-setup")]` on each method becomes a value of the generated `WordSectionAction` enum.
- XML documentation on the interface parameters becomes the parameter descriptions in the MCP schema.

The generator merges the parameters of all actions into one method, so a parameter used by only
some actions is emitted as optional. To change the API, edit the interface — never the generated
code. `file` stays hand-written because it manages sessions rather than operating on one.

Inspect the emitted code under `src/WordMcp.McpServer/obj/generated`. The tests in
`GeneratedToolContractTests` compare the generated surface against the interfaces, so a mismatch
fails the build rather than reaching a client.

Tests that need a real Word installation are marked `[Trait("Category", "RequiresWord")]` and are excluded in CI. Failed integration runs can leave orphaned `WINWORD.EXE` processes behind, which slow down or block later runs — clean them up with `Get-Process WINWORD | Stop-Process -Force` before re-running.

### Further reading

| Document | Covers |
|---|---|
| [docs/architecture.md](docs/architecture.md) | The layers, the request flow and how the tool layer is generated |
| [docs/com-interop.md](docs/com-interop.md) | STA threading, releasing COM objects and the Word behaviour behind the pitfalls above |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Building, testing, adding a tool, cutting a release |
| [skills/word-mcp/SKILL.md](skills/word-mcp/SKILL.md) | An agent-facing guide to using the tools in the right order |

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| `Word is not installed or not registered for COM` | Install Word desktop; the Microsoft Store version cannot be automated |
| `Could not load file or assembly 'office'` | `office.dll` was not found in the GAC — reinstall or repair Office |
| Operation times out | A Word dialog is waiting for input; close it and retry |
| `The file is already open in Word` | Close the document in the Word UI |

## Contributing

Bug reports and feature requests go through the [issue templates](https://github.com/trsdn/mcp-server-word/issues/new/choose).
Pull requests are welcome; [CONTRIBUTING.md](CONTRIBUTING.md) covers how to build, how to run the
two halves of the test suite, and what it takes to add a tool.

Please read the [Code of Conduct](CODE_OF_CONDUCT.md) first.

**Do not file a public issue for a security problem** — report it privately as described in the
[security policy](SECURITY.md).

## License

MIT — see [LICENSE](LICENSE).
