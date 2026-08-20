# Architecture

Six projects, layered so that everything above the COM boundary is testable without Word running.

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

`WordMcp.Service` sits beside the MCP server and wraps `SessionManager` behind a data-only
request/response contract, so sessions can move out of the client's process. The MCP server routes
every session operation through it via `ServiceBridge`, in-process for now.

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
protocol, and MCP clients surface stderr as an error). `WordToolsBase` provides the shared
serialization and the `try`/`catch` that turns any exception into a JSON error payload rather than
a transport failure.

`WordFileTool` is the only hand-written tool, because session lifecycle does not fit the
generator's shape. The other fourteen are generated.

`ServiceBridge` is the seam to the session service, and everything crossing it is data. It exists
because the file tool and `WordMcpService` used to carry two copies of open, create, save, close,
list and test — and which copy a user hit depended on which entry point they came in through. A
tool now validates what only it can validate, hands the command to the service and serializes what
comes back. Path validation stays in the tool, so the caller who typed a relative path or a `.pptx`
gets the tool's wording rather than the service's.

The service reports failures as data; the tool layer formats them from exceptions.
`ServiceCommandException` carries the reported CLR type name across that gap, so a caller still
sees `FileNotFoundException` and not the wrapper.

Nothing in this project holds a COM object any more — not even indirectly. That is what makes
`WORDMCP_SERVICE_MODE=daemon` a working switch rather than a promise: the same tool call runs
in-process or travels down a pipe to another process, and no tool can tell which.

### WordMcp.Service

Holds the sessions so they can outlive a single client. `WordMcpService` owns the
`SessionManager` and exposes it only through one method:

```
ProcessAsync({ command, sessionId, args }) → { success, result, errorMessage, errorType }
```

Nothing in that contract is a COM object or a delegate, which is the whole point: the same
service can be called in-process or across a named pipe and the caller cannot tell the
difference. It is also the reason the boundary sits here and not at `IWordBatch` —
`IWordBatch.Execute` takes a delegate and cannot cross a process.

Three supporting pieces:

- **`ServiceProtocol`** — newline-delimited JSON. One request per line in, one response per line
  out. Serialization is compact, so a newline inside a value is escaped and never breaks framing.
- **`ServiceSecurity`** — the pipe name embeds the caller's SID and the pipe carries an ACL for
  that SID alone, so two users never meet. Clients pass `PipeOptions.CurrentUserOnly`, which makes
  Windows verify the server's owner before any byte moves, closing the reverse direction too.
- **`OrphanWordCleanup`** — terminates `WINWORD.EXE` processes *this service started* and that
  outlived their session. Only tracked process ids are ever considered, and only if the process is
  still named `WINWORD`, so a Word the user opened by hand is never at risk.

Idle behaviour is a pure predicate (`IsIdle`) over an injectable `TimeProvider` rather than a
timer, which keeps it testable without waiting. An open session always counts as active however
long ago it was touched: closing Word underneath a client that still holds a session id turns a
dormant workflow into a broken one.

### The daemon

`WordMcp.Service` is also an executable. `--daemon` runs `ServiceHost`, which accepts pipe
connections and feeds each line to `WordMcpService`; `--status` and `--stop` are thin clients
against a running one. `ServiceClient` is the other side, and starting it is the interesting part:

- **A missing service is not an error.** `SendAsync` first tries to connect; if nothing answers and
  auto-start is on, it spawns the daemon, waits for it to listen, and retries. Callers never have
  to know whether a service was already running.
- **One connection per request.** Locally that costs microseconds and buys reconnect for free: a
  daemon that exited on its idle timeout is simply started again by the next call, with no stale
  connection state to reason about.
- **Single instance via a named mutex, not the pipe.** Windows happily allows several servers on
  one pipe name, so two daemons would both appear to start and each client would see half the
  sessions. `WordMcp-host-{pipeName}` settles it before the pipe is ever created.

The accept loop parks in `WaitForConnectionAsync`, which no cancellation token can interrupt on
Windows, so the watchdog that notices shutdown or idleness wakes it with a throwaway connection.

Set `WORDMCP_SERVICE_MODE=daemon` to make the MCP server route every command through that pipe
instead of running the service in its own process. In-process stays the default because it is what
a single client wants: no second process, no pipe, no startup wait. The daemon earns its keep only
when several clients have to see the same open document.

One thing the daemon uncovered that no in-process run could: Word is driven through `dynamic`, and
the runtime binder needs `office.dll` to bind a member — a file every build copies next to the
executable but no package declares as a runtime dependency. A test host probes its own directory
and hides this; a plain executable does not, and the daemon could not open a single document until
`ComAssemblyResolver` taught it to look. It resolves only the Office interop assemblies, from the
application directory, and is installed by `WordBatch` itself so no host has to remember it.

## The generated tool layer

A source generator reads the command interfaces and emits **both halves of the process boundary**
from the same description, which is the only way they cannot drift apart:

- In `WordMcp.McpServer` it emits one MCP tool class per `[ServiceCategory]`: a method with an
  `action` parameter, the union of all parameters of that tool's actions (each nullable, since a
  given action uses only some), and a body that packs those parameters under their wire names and
  hands `category.action` to `ServiceBridge`.
- In `WordMcp.Service` it emits the dispatcher that unpacks them again, applies the same defaults,
  raises the same "x is required for this action." errors, and calls the command implementation.
  It finds that implementation by looking for the class implementing the interface, so adding a
  category needs no registry entry anywhere.

The generator picks its half from the assembly name of the compilation it runs in, and is
referenced as an analyzer by both projects.

The session is resolved lazily, only once the dispatcher recognises the command. A misspelled
command therefore reports that it is unknown instead of complaining about a missing `session_id`
the caller never needed to supply.

`EmitCompilerGeneratedFiles` is on for both projects, so the output lands in
`src/WordMcp.McpServer/obj/generated` and `src/WordMcp.Service/obj/generated` and can be read like
ordinary code. It is worth looking at once after adding a tool, to confirm that the description a
client sees is the one that was written.

Why generate: fifteen tools with close to seventy actions between them means the boilerplate is
where the bugs would be, and a hand-written tool class drifts from its interface silently.

## Request flow

```
tool call
  → generated tool method
      → WordToolsBase.Execute(tool, action, …)      catch → JSON error
          → ServiceBridge.Invoke("text.replace", session_id, { … })
              → in-process, or one line down a named pipe
                  → GeneratedToolDispatch                 → unknown command stops here
                      → session lookup                    → SessionManager
                      → XyzCommands.Action(batch, args)
                          → validate arguments            throws here on bad input
                          → batch.Execute(word => …)      marshalled to the STA thread
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
| `WordMcp.Service.Tests` | all but the lifecycle suite | protocol framing, pipe security, orphan cleanup, dispatch and idle rules; with Word, the session lifecycle |

Tests that need Word carry `[Trait("Category", "RequiresWord")]`; CI runs
`--filter "Category!=RequiresWord"` because GitHub runners have no Office. That makes the
integration suite a local gate, which is why the pull request template asks which Word version a
change was verified against.

## Versioning

`Directory.Build.props` holds the fallback version. A release passes `-p:Version=` derived from the
git tag, `AssemblyVersion` and `FileVersion` follow it, and a build target stamps it into a
generated copy of `.mcp/server.json` — the MCP registry rejects a manifest whose version disagrees
with the package. See [CONTRIBUTING.md](../CONTRIBUTING.md).
