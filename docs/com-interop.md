# Working with Word through COM

Notes for anyone changing `WordMcp.ComInterop` or writing a new command. The user-facing quirks
live in the README under "Known behaviour and pitfalls"; this is the layer below that.

## Threading

Word's automation interface is single-threaded apartment. Getting this wrong does not produce a
clean error — it produces a `WINWORD.EXE` that never exits, or a call that fails once in fifty
runs.

Each `WordBatch` therefore owns:

* **One STA thread**, created with `SetApartmentState(ApartmentState.STA)` before `Start()`.
* **One work queue**, an unbounded `Channel<Action>`. `batch.Execute(...)` enqueues the delegate
  and waits; the STA thread runs them one at a time.
* **One OLE message filter**, registered on that thread and revoked on shutdown.

Everything follows from this:

* **All COM objects live on the STA thread.** A `Word.Document` handed out to a caller and touched
  from a thread-pool thread is a marshalling bug waiting for a slow machine to expose it.
* **Operations within a batch are serial.** For parallel work, open a second batch on a second
  document.
* **Never block the STA thread from inside an operation.** Waiting on another batch from within
  `Execute` deadlocks.

### The OLE message filter

Word rejects incoming calls while it is busy with something else, and .NET surfaces that as
`COMException` with `RPC_E_SERVERCALL_RETRYLATER`. `OleMessageFilter` intercepts the rejection and
retries for a bounded time before giving up, which turns a common transient failure into a short
pause. Without it, an operation running while Word is repaginating fails for no reason a user
could act on.

## Releasing COM objects

`ComUtilities.Release(ref obj)` calls `Marshal.ReleaseComObject` and nulls the reference. It
matters most inside loops:

```csharp
foreach (...)
{
    Word.Range? range = paragraph.Range;
    try { /* work */ }
    finally { ComUtilities.Release(ref range); }
}
```

A leaked runtime callable wrapper keeps `WINWORD.EXE` alive after the session ends. The symptom is
not an exception — it is a process in Task Manager and the next `file(open)` on the same document
failing because the file is locked. Ranges, paragraphs, tables and cells obtained during iteration
are the usual culprits.

`ComUtilities.TryQuitWord` swallows errors on purpose: shutdown must not throw over a Word that
has already gone away.

## Long dotted chains lie

`doc.Paragraphs[1].Range.Font.Bold = true` creates four wrappers and releases none of them. Split
the chain when it runs more than a couple of times.

## Where the collection hangs matters

Some Word collections live on the `Application`, not on the `Document` — `Application.Options`,
the dialogs, the recent files list. Reading one of those changes global Word state that outlives
the session, so treat anything on `Application` as shared and restore it if you change it.

The mirror image is that several document-wide operations silently cover the body only.
`Document.Fields.Update()` and `Document.AcceptAllRevisions()` both skip headers and footers, which
is why `field(update-all)` and `revision(accept)` iterate `Sections` and handle each section's
`Headers` and `Footers` on top of the document-wide call.

## Constants and conversions

Word constants are `int` in this codebase, declared in `ComInteropConstants`, not taken from the
interop enums — the enum names are long, inconsistently cased, and the numeric values are the
documented part of the API anyway.

String-to-constant conversion belongs in `WordConversions` and must happen **before**
`batch.Execute`, so an unknown value produces a message naming the valid options rather than a
COM error from inside Word. This is also what makes it testable without Word.

Style names go through `WordStyles`: built-in English names are translated to Word's
language-independent style ids, so `Heading 1` works on a German install. Anything unrecognised is
passed through, which is how custom and localized styles are addressed.

## Documents Word will not open cleanly

Checked before Word is launched, because Word's own failure mode is a modal dialog on a hidden
window — which looks like a hang:

* **Unsupported extension** — `FileAccessValidator.ValidateExtension`.
* **IRM/AIP protection** — an encrypted `.docx` is an OLE2 compound document rather than a ZIP, so
  the file signature gives it away. Legacy `.doc` is legitimately OLE2, hence the extension check
  first.
* **Already open** — an exclusive-access probe, reported as a message telling the user to close
  the document.

## Creating documents

`file(create)` writes an empty Open XML package itself (`EmptyDocumentFactory`) and then opens it.
Creating through `Documents.Add` plus `SaveAs` is unreliable on machines signed in to Microsoft
365: AutoSave claims the new document for OneDrive and quietly ignores the requested local path.

## Timeouts

Two of them, both in `ComInteropConstants`: a startup timeout for the Word instance and a
per-operation timeout (five minutes by default). An operation that exceeds it is almost always a
modal dialog waiting on an invisible desktop, so the error message says so.

## Debugging

* **Leftover processes**: `Get-Process WINWORD | Stop-Process -Force` before an integration run. A
  failed run leaves Word behind, and the next run then fails on a file lock instead of the real
  problem.
* **Watch it work**: batches take a `visible` flag. Seeing the document change makes an unexpected
  dialog obvious.
* **`HRESULT 0x800A...`** is Word saying no, not .NET failing. The number is rarely specific;
  `0x800A11FD` ("this command is not available") in particular covers everything from a wrong
  selection to a feature disabled by the document's protection state.

## When Word behaves oddly, write it down

Most of the hard-won knowledge here is behavioural, and none of it is discoverable from the type
signatures. Leave a comment at the line where the surprise bites, and add a bullet to the README
when a user of the server would trip over it too.
