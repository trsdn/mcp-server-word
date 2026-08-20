# Security Policy

## Supported versions

This project is pre-1.0. Only the latest release receives security fixes.

| Version | Supported |
|---|---|
| 0.1.x | ✅ |
| < 0.1 | ❌ |

## Reporting a vulnerability

**Do not open a public issue for a security problem.**

Report it privately through GitHub:
[**Report a vulnerability**](https://github.com/trsdn/mcp-server-word/security/advisories/new).

Please include what you have — the affected version, what an attacker can achieve, and the steps
to reproduce it. A working proof of concept helps but is not required to file a report.

Expect an acknowledgement within seven days. This is a spare-time project, so a fix may take
longer than that; you will be kept informed either way, and credited in the advisory unless you
ask otherwise.

## Threat model

The server automates a local Microsoft Word installation on behalf of an MCP client. That shapes
what counts as a vulnerability here.

**In scope**

* A tool argument that escapes the intended document — writing to, reading from or deleting a path
  the caller did not name.
* Argument handling that reaches Word in a way the caller did not ask for, for example a path or
  field code that Word interprets rather than treats as data.
* A crafted document that makes the server execute code while it is being processed.
* Credentials or document content leaking into logs or error messages.

**Out of scope**

* An MCP client that sends destructive commands. The server executes what it is asked to; deciding
  what to ask is the client's job, which is why destructive actions are marked as such in the tool
  metadata.
* Macros in a `.docm` a user deliberately opens. The server does not run macros itself, but Word's
  own macro settings govern what happens on open.
* Anything requiring an attacker who already has interactive access to the machine, which is the
  same trust level the server itself runs at.
* Vulnerabilities in Microsoft Word — report those to
  [Microsoft MSRC](https://msrc.microsoft.com/report).

## Operating notes

* The server runs with the privileges of the user who starts it and can read and write any file
  that user can. Treat access to it as equivalent to shell access.
* Rights-protected documents (IRM/AIP) are rejected before Word is launched.
* Documents are automated through COM, so a Word instance is running for the lifetime of a session.
