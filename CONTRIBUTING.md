# Contributing

Thanks for looking. This is a spare-time project, so the fastest way to get a change in is to make
it easy to review.

## Before you start

Open an issue for anything larger than a bug fix. A new tool touches the source generator, the
README and the test suite, and it is worth agreeing on the shape before you write it.

## Setting up

Requirements:

* Windows 10 or 11
* .NET 9 SDK
* Microsoft Word 2016 or newer — the **desktop** version. The Microsoft Store version cannot be
  automated through COM.

```powershell
git clone https://github.com/trsdn/mcp-server-word
cd mcp-server-word
dotnet build
dotnet test --filter "Category!=RequiresWord"
```

## Running the tests

The suite splits in two.

```powershell
# Everything that runs without Word. This is what CI runs.
dotnet test --filter "Category!=RequiresWord"

# Everything, including the tests that drive a real Word instance.
dotnet test
```

The Word tests take roughly 35 seconds each, so a full run is around ten minutes. While iterating,
narrow it down:

```powershell
dotnet test --filter "FullyQualifiedName~BookmarkIntegrationTests"
```

A failed integration run can leave `WINWORD.EXE` behind, which slows down or blocks the next run.
Clear it first:

```powershell
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
```

## Adding a tool

The MCP tool classes are generated. [docs/architecture.md](docs/architecture.md) explains the
layering; [docs/com-interop.md](docs/com-interop.md) covers the COM rules a new implementation has
to follow. A new tool is three hand-written pieces:

1. **The interface** in `src/WordMcp.Core/Commands/<Name>/I<Name>Commands.cs`, carrying
   `[ServiceCategory]`, `[McpTool]` and one `[ServiceAction]` per action.
2. **The COM implementation** next to it.
3. **A property on `WordServices`** in `src/WordMcp.McpServer/WordServices.cs`. The generator maps
   the interface type to that property, so without it nothing is emitted.

The generated code lands in `src/WordMcp.McpServer/obj/generated` and is worth reading once, so the
tool description you wrote is the one a client actually sees.

Any new `Wd*` constant belongs in `src/WordMcp.ComInterop/ComInteropConstants.cs`, and any
string-to-enum mapping in `src/WordMcp.Core/Utilities/WordConversions.cs`.

### Tool descriptions are the interface

An assistant sees only the description. Say what an action does, what the arguments mean, and where
Word behaves in a way nobody would predict. The descriptions in the existing tools are long on
purpose.

## Conventions that the tests enforce

* **Validate arguments before `batch.Execute`.** Anything checked outside the batch is testable
  without Word, which is why the validation tests run in CI. Tests use `ThrowingBatch`: it throws
  `NotSupportedException`, so a test expecting that exception is asserting "validation let this
  through".
* **Convert strings to Word constants before the batch too**, so a bad value produces a clear
  message rather than a COM error from inside Word.
* **Integration tests create their own document** when they change document-wide state, and are
  marked `[Trait("Category", "RequiresWord")]`.
* **Warnings are errors.** So are the xUnit analyzers — `Assert.Single(x.Where(...))` will not
  compile.

## Write down what Word did

Most of the hard-won knowledge in this project is about Word behaving unexpectedly: a collection
hanging off the Application instead of the Document, a parameter that silently changes the scope of
an operation, a property that throws instead of returning empty.

When you hit one, leave a comment at the line where it bites — not only in the pull request, which
nobody reads a year later. If it is something a user of the server would trip over, add it to the
"Known behaviour and pitfalls" section of the README as well.

## Commits and pull requests

* Write commit messages that explain why, not what. The diff already says what.
* One logical change per pull request.
* Fill in the pull request template, especially which Word version you verified against — CI cannot
  run the integration tests, so that line is the only evidence they passed.

## Releasing (maintainers)

Releases are driven entirely by a tag. `Directory.Build.props` carries the fallback version for
local builds; the published version comes from the tag, and a build target stamps it into
`.mcp/server.json` so the MCP registry sees the same number as NuGet.

```powershell
git tag v0.2.0
git push origin v0.2.0
```

That runs `.github/workflows/release.yml`: build, test, pack, verify the stamped `server.json`,
push to nuget.org and open a GitHub release with generated notes. A tag containing a hyphen
(`v0.2.0-rc.1`) is published as a prerelease.

Publishing needs a `NUGET_API_KEY` repository secret. Without it the push step fails and the
release is not created — nothing half-published escapes. To rehearse the build without publishing,
run the workflow manually from the Actions tab and pass a version; that path packs and uploads an
artifact but never pushes.

## Reporting bugs
Use the [issue templates](https://github.com/trsdn/mcp-server-word/issues/new/choose). Include the
Word build number and UI language; localization is behind a surprising share of the bugs here.

Security problems do not go in the issue tracker — see [SECURITY.md](SECURITY.md).
