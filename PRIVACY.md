# Privacy

word-md processes Markdown locally inside the Office task pane.

No pasted Markdown content, loaded file content, or generated Word content is
sent to the project maintainer or to a project backend. The project does not
set cookies, collect analytics, or maintain user accounts.

The task pane loads:

- `office.js` from Microsoft, required for Office add-ins.
- `vendor/marked.min.js` from this same hosted package.

The add-in requires `ReadWriteDocument` so it can insert content into the
active Word document. It does not intentionally read, store, or transmit the
document contents.
