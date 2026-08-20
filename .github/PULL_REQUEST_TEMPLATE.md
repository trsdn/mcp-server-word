<!--
  Keep this short. The diff says what changed; this says why, and what you did to convince
  yourself it works.
-->

## What this changes

<!-- One or two sentences. -->

## Why

<!-- The problem being solved. Link the issue: "Closes #12" or "Refs #12". -->

## How it was verified

<!--
  Word automation fails in ways a compiler cannot catch, so say what you actually ran.

  - Word integration tests are excluded in CI. If you touched a command implementation, run them
    locally: dotnet test --filter "FullyQualifiedName~YourIntegrationTests"
  - Note the Word version and UI language you tested against - localization breaks style and
    caption handling in particular.
-->

- [ ] `dotnet build -c Release` is clean (warnings are errors in this repo)
- [ ] `dotnet test --filter "Category!=RequiresWord"` passes
- [ ] Word integration tests pass locally, or this change cannot affect them

Word version tested against:

## Checklist

- [ ] A new tool has an interface, a COM implementation and a `WordServices` property — the MCP
      tool class itself is generated, so it is not written by hand
- [ ] New behaviour has tests, including the failure paths
- [ ] Argument validation happens before `batch.Execute`, so it is testable without Word
- [ ] README updated if a tool, action or pitfall changed
- [ ] Anything surprising about Word's behaviour is written down as a comment at the point where it
      bites, not only in this description
