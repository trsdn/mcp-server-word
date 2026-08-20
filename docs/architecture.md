# Architecture

Five projects, layered so that everything above the COM boundary is testable without Word running.

```
MCP client (VS Code, Claude Desktop, Copilot CLI)
        │  stdio, JSON-RPC
        ▼
WordMcp.McpServer          tool classes, JSON responses, DI host
        │                  ← generated at build time from the interfaces below
        ▼
WordMcp.Core               argument validation, command implementations, result models
        │
        ▼
WordMcp.ComInterop         session manager, STA thread, Word COM objects
        │
        ▼
WINWORD.EXE
```

`WordMcp.Generators.Mcp` and `WordMcp.Generators.Shared` are build-time only; they ship no runtime
code.

## The layers

### WordMcp.ComInterop

Owns everything that touches COM, and nothing else. `SessionManager` maps a `session_id` to an
`IWordBatch`; a batch owns one Word `Application`, one open `Document` and one STA thread. Also
here: the OLE message filter, extension and rights-protection validation, and the code that writes
an empty `.docx` package without going through Word.

Details are in [com-interop.md](com-interop.md).

### WordMcp.Core

One folder per tool under `Commands/`, each with an interface and an implementation:

```
Commands/Bookmark/IBookmarkCommands.cs    the contract the generator reads
Commands/Bookmark/BookmarkCommands.cs     the COM work
Models/BookmarkResults.cs                 the records that get serialized
```

The interface carries three kinds of attribute:

| Attribute | On | Purpose |
|---|---|---|
| `[ServiceCategory]` | interface | Groups the actions under one tool |
| `[McpTool]` | interface | Tool name and the description the client sees |
| `[ServiceAction]` | method | Action name and description |

Results are `record` types in `Models/`, serialized as camelCase with nulls omitted.

**Validation happens before the batch.** Every implementation checks its arguments, converts
strings to Word constants, and only then calls `batch.Execute`. That split is what makes the
validation testable on a machine without Word — the tests pass a `ThrowingBatch` whose `Execute`
throws, so "the call reached the batch" is the assertion that validation let the arguments
through.

### WordMcp.McpServer

The host: stdio transport, DI, logging to stderr from `Warning` upwards (stdout carries the
protocol, and MCP clients surface stderr as an error). `WordServices` exposes one static property
per command interface, and `WordToolsBase` provides the shared serialization and the `try`/`catch`
that turns any exception into a JSON error payload rather than a transport failure.

`WordFileTool` is the only hand-written tool, because session lifecycle does not fit the
generator's shape. The other fourteen are generated.

## The generated tool layer

A source generator reads the command interfaces and emits one MCP tool class per
`[ServiceCategory]`. For each tool it produces a method with an `action` parameter, the union of
all parameters of that tool's actions (each nullable, since a given action uses only some), a
switch over the action names, and the `Execute` wrapper.

The generator resolves the interface type to the `WordServices` property that provides it. **A new
command interface without a matching property on `WordServices` produces nothing** — no error, no
tool.

`EmitCompilerGeneratedFiles` is on, so the output lands in `src/WordMcp.McpServer/obj/generated`
and can be read like ordinary code. It is worth looking at once after adding a tool, to confirm
that the description a client sees is the one that was written.

Why generate: fifteen tools with close to seventy actions between them means the boilerplate is
where the bugs would be, and a hand-written tool class drifts from its interface silently.

## Request flow

```
tool call
  → generated tool method
      → WordToolsBase.Execute(tool, action, …)      catch → JSON error
          → WordToolsBase.Batch(session_id)         → SessionManager
              → XyzCommands.Action(batch, args)
                  → validate arguments              throws here on bad input
                  → batch.Execute(word => …)        marshalled to the STA thread
                      → Word COM
                  ← result record
          ← JSON
```

Failures never travel as transport errors. Anything that throws comes back as

```json
{ "success": false, "isError": true, "tool": "text", "action": "Replace",
  "errorType": "KeyNotFoundException", "errorMessage": "Session 'word-unknown' not found." }
```

which keeps a bad argument from looking like a broken server.

## Testing

| Project | Runs in CI | Covers |
|---|---|---|
| `WordMcp.Core.Tests` | validation tests only | argument validation, conversions, and — with Word — the real COM behaviour |
| `WordMcp.McpServer.Tests` | yes | that the generated tools match the interfaces |

Tests that need Word carry `[Trait("Category", "RequiresWord")]`; CI runs
`--filter "Category!=RequiresWord"` because GitHub runners have no Office. That makes the
integration suite a local gate, which is why the pull request template asks which Word version a
change was verified against.

## Versioning

`Directory.Build.props` holds the fallback version. A release passes `-p:Version=` derived from the
git tag, `AssemblyVersion` and `FileVersion` follow it, and a build target stamps it into a
generated copy of `.mcp/server.json` — the MCP registry rejects a manifest whose version disagrees
with the package. See [CONTRIBUTING.md](../CONTRIBUTING.md).
