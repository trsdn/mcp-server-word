# Word MCP Server

Automate Microsoft Word from chat. This extension registers the
[WordMcp](https://github.com/trsdn/mcp-server-word) MCP server with VS Code, which gives an agent
seventeen tools for reading, writing and formatting `.docx` documents through real Word.

Because it drives Word itself rather than rewriting the XML, the result is a document Word considers
its own: styles resolve, fields calculate, the table of contents matches the headings, and a PDF
export looks like one made by hand.

## Requirements

- **Windows** with **Microsoft Word 2016 or newer, desktop edition.** The Microsoft Store version of
  Office cannot be automated.
- **.NET 10 SDK or newer**, for the `dnx` command that fetches and runs the server.

The server is not bundled with the extension. It is resolved from nuget.org on first use, so it
updates without a new extension release.

## Setup

Install the extension, then open **MCP: List Servers** from the Command Palette and start **Word**.
Nothing else is configured; the server is discovered through the extension.

## Tools

| Tool | What it covers |
|---|---|
| `file` | Open, create, save and close documents; one session per document |
| `text` | Read, append, find, replace and format text |
| `paragraph` | List, add, insert, delete, style and align paragraphs |
| `table` | Create tables, read and write cells, add and delete rows |
| `document` | Counts, properties, PDF export, save as |
| `image` | Insert, resize, replace and delete pictures; alt text |
| `field` | Fields, page numbers and tables of contents |
| `section` | Section breaks, margins, orientation, paper size |
| `header-footer` | Headers and footers per section |
| `style` | List, create, modify and delete styles |
| `list` | Bullet, numbered and outline lists |
| `comment` | Read, add, delete and resolve comments |
| `revision` | Tracked changes: list, accept, reject, toggle tracking |
| `bookmark` | Stable references that survive edits |
| `footnote` | Footnotes and endnotes |
| `content-control` | The structured fields a Word template exposes |
| `screenshot` | Render a page as an image to check the layout |

## Settings

| Setting | Purpose |
|---|---|
| `wordMcp.serverVersion` | Pin a specific `WordMcp.McpServer` version. Empty means the latest. |
| `wordMcp.serviceMode` | `daemon` keeps sessions in a background service so they survive a restarted client. |

## Notes

- Paths must be absolute.
- A document already open in Word blocks the session; close it first.
- Word runs invisibly and is shut down when the session closes.

Full documentation, including the known Word behaviours worth reading before a long editing run,
lives in the [repository README](https://github.com/trsdn/mcp-server-word#readme).

## License

MIT
