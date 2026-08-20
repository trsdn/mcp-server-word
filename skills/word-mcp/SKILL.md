---
name: word-mcp
description: >
  Automate Microsoft Word on Windows through the WordMcp MCP server. Use when creating, reading,
  or editing Word documents: text, paragraphs, tables, images, styles, lists, sections, headers
  and footers, fields and tables of contents, comments, tracked changes and bookmarks. Can render
  a page as an image to check layout.
  Triggers: Word, document, docx, report, letter, memo, table of contents, tracked changes,
  comments, bookmarks, headers, footers.
---

# Word MCP Server Skill

Fifteen tools, each taking an `action` parameter. The tool descriptions are complete; this file
covers the workflow, the ordering rules that are not obvious, and the places where Word behaves
in a way nobody would guess.

## Preconditions

- Windows with Microsoft Word 2016 or newer, **desktop** edition. The Microsoft Store version
  cannot be automated.
- Absolute paths only: `C:\Users\me\Documents\report.docx`.
- The document must not already be open in Word. Ask the user to close it rather than guessing.
- Word runs invisibly and is shut down with the session.

## The one rule that matters

**Everything runs inside a session.** `file(open)` or `file(create)` returns a `session_id` that
every other call needs, and `file(close)` is what actually writes the document and releases Word.

```
file(open|create) → session_id → edit → file(save) → file(close, save: true)
```

A session left open holds a `WINWORD.EXE` process and an exclusive lock on the file. Close it even
when the work failed halfway through.

## Building a document

Order matters more than it looks:

1. `file(create)` — an empty document.
2. Content first: `paragraph(add)` with `style: "Heading 1"` for structure, `text(append)`,
   `table(create)`, `image(insert)`.
3. Formatting second: `style(create|modify)`, `list(apply)`, `paragraph(set-alignment)`.
4. Page setup: `section(page-setup)`, then `header-footer(set)`, then `field(insert-page-number)`.
5. Fields last: `field(insert-toc)` then `field(update-all)`. A table of contents inserted before
   the headings exist stays empty.
6. `file(close, save: true)`.

## Reading a document

- `document(get-info)` for the counts, `paragraph(list)` for the structure. Start there rather
  than pulling the entire text with `text(get)`.
- `text(find)` returns character positions that `text(format)` and `text(get)` accept.
- `screenshot(page)` answers layout questions — does the table fit, where does the page break, is
  the header on the right page — far more reliably than any measurement.

## Indexes, and how to avoid getting burned by them

Paragraph, table, image, comment and revision indexes are **1-based** and they **shift** on every
insertion or deletion. Two habits keep this from causing damage:

- Delete from the back to the front, or `list` again after every delete.
- When a passage has to be referred to more than once, `bookmark(add)` it. Bookmarks survive edits
  elsewhere; indexes do not.

## Recipes

**Report with a table of contents**

```
file(create, path: "C:\\temp\\report.docx")
paragraph(add, text: "Quarterly Report", style: "Title")
paragraph(add, text: "Summary", style: "Heading 1")
text(append, text: "Revenue grew by fifteen percent.", as_new_paragraph: true)
field(insert-toc)          # after the headings exist
field(update-all)
file(close, save: true)
```

**Review pass**

```
revision(set-tracking, enabled: true)     # before the edits, it is not retroactive
text(replace, find: "fifteen", replace: "seventeen")
comment(add, paragraph_index: 4, text: "Source?", anchor_text: "seventeen percent")
revision(list)
```

**Check the layout**

```
screenshot(page, page: 2)                 # dpi 150 by default; 96 is enough for a quick look
```

Only pass `include_image: true` when the image is actually going to be looked at — otherwise it
just fills the context window.

## Non-obvious behaviour

**Ordering**

- `field(insert-toc)` on a document without heading styles produces an empty table of contents.
  Apply `Heading 1`/`Heading 2` first, then `field(update-toc)`.
- `section(page-setup)` applies the paper size before the margins, because changing the paper size
  resets them. Without a `section_index` it hits every section.
- `revision(set-tracking)` is not retroactive.

**Scope**

- `field(update-all)` and `revision(accept|reject)` without an index deliberately walk headers and
  footers too; Word's own document-wide calls cover the body only.
- `image` covers inline pictures only. Text boxes, floating shapes and charts are invisible to it.
- Headers are inherited between sections until something is written to them. `header-footer(set)`
  with a `section_index` breaks that link for you.
- `first-page` and `even-pages` headers only render once the section switch is on, which
  `header-footer(set)` also handles.

**Traps**

- `document(save-as)` **also saves the original**, because Word has no format-changing "save a
  copy". Use `document(export-pdf)` when the original must stay untouched.
- `comment(resolve)` usually fails on Microsoft 365: comments added through the API are unposted
  drafts and a draft cannot be marked done. Delete the comment instead.
- `table(read)` returns merged cells as empty strings.
- Bookmark names must start with a letter and contain only letters, digits and underscores.
- Styles are addressed in English (`Heading 1`) but Word reports them localized. `style(list)`
  returns both `name` and `english_name` — send `english_name` back when it is there.
- Colours are hex RGB (`#0078D4`). Sizes and margins are in points (72 pt = 1 inch,
  1 cm = 28.35 pt).

## When something fails

Errors come back as JSON with `errorType` and `errorMessage`, not as a transport failure, so the
message is worth reading before retrying.

| Message | What to do |
|---|---|
| `Session '…' not found` | The session was closed. Open the document again. |
| `The file is already open in Word` | Ask the user to close it. |
| Call times out | A Word dialog is waiting for input on the desktop. |
| `Word is not installed or not registered for COM` | Run `file(action: "test")` to confirm, then stop; this cannot be worked around. |
