# Security

## Security contact

Report suspected vulnerabilities to `mb@xato.net`.

Please include:

- The affected file or hosted URL.
- Steps to reproduce.
- Expected and actual behavior.
- Any proof-of-concept Markdown or manifest changes needed to reproduce.

## Runtime model

word-md is a static Office add-in. It has no project backend and stores no
server-side state. The task pane loads `office.js` from Microsoft and loads
the Markdown parser from the same hosted package as the add-in.

## Content handling

Pasted Markdown and loaded files stay in the Word task pane process. Preview
rendering constructs DOM nodes with text APIs instead of injecting parsed HTML.
Raw Markdown HTML is inserted into Word as plain text.

## Permission scope

The manifest declares `ReadWriteDocument` because the add-in inserts generated
Word content at the current selection. It does not intentionally read,
transmit, or persist document content.

## Release checks

Run `npm run check` before publishing a new manifest or hosted package.
The check verifies JavaScript syntax, CSP guardrails, self-hosted parser usage,
and the SHA-256 hash of the vendored `marked` bundle.

