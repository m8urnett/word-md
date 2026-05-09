# Word-md — Word Add-in

Word-md is a Microsoft Word task pane add-in that allows you paste markdown
into word as formatted text.

Paste Markdown, or load a local `.md` file, and word-md inserts it into the
active document as native Word content: real Heading 1-6 styles, Word lists,
Quote paragraphs, preformatted code blocks, hyperlinks, task-list glyphs, and
actual Word tables. The result is not a pasted web fragment. It behaves like
content authored directly in Word, so document templates, style changes, table
formatting, and downstream review workflows keep working.

The add-in is intentionally small and static. It has no project backend, no
analytics, and no runtime parser CDN dependency; Markdown is processed inside
the Office task pane and written to the current selection through Word.js.



## Style map

| Markdown                           | Word style / treatment                      |
| ---                                | ---                                         |
| `# H1` … `###### H6`               | Heading 1 … Heading 6                       |
| Paragraph                          | Normal                                      |
| `*` / `-` bullet list              | List Bullet (indent +18pt per nest level)   |
| `1.` ordered list                  | List Number                                 |
| `> blockquote`                     | Quote                                       |
| Triple-backtick fenced code        | HTML Preformatted (one para per line)       |
| `` `inline code` ``                | Consolas + light grey highlight             |
| `**bold**`, `*italic*`, `~~del~~`  | font.bold / italic / strikeThrough          |
| `[text](url)` / autolink           | Range.hyperlink                             |
| `\| pipe \| table \|`              | Real Word table, Grid Table 4 - Accent 1    |
| `---` horizontal rule              | Empty para with bottom border               |
| `- [x] task`                       | Checkbox glyph + List Bullet                |
| YAML frontmatter                   | Stripped (option, on by default)            |

## Sideload installation

For detailed instructions on setting up a trusted catalog for local sideloading, see [sideload.md](./sideload.md).

1. Host these files at any HTTPS endpoint. Easiest: GitHub Pages.
   - The bundled `manifest.xml` points at `https://m8urnett.github.io/word-md/`.
   - **Security Note:** Do not use the default URL for production or sensitive work. Hosting the manifest yourself ensures you control the code running in your environment. Fork this repo or host the static files on your own HTTPS origin.
2. **Word for Windows/Mac**: Insert ribbon → Get Add-ins → My Add-ins
   (top of dialog) → Manage My Add-ins → **Upload My Add-in** → pick
   `manifest.xml`.
3. **Word on the web**: Home → Add-ins → More Add-ins → **Upload My Add-in**.
4. The **word-md** group appears on the Home tab. Click **Paste** to open the pane.

## Enterprise deployment

This repo includes a small release-readiness wrapper:

- `vendor/marked.min.js` is pinned and self-hosted; the task pane does not load
  the Markdown parser from a third-party CDN at runtime.
- Preview rendering builds DOM nodes with text APIs instead of assigning parsed
  Markdown to `innerHTML`.
- Inserted hyperlinks are limited to `http:`, `https:`, and `mailto:` URLs.
- Markdown input is capped at 1 MB to avoid freezing the task pane or Word on
  oversized input.
- The vendored parser bundle is checked by SHA-256 in CI.
- `SECURITY.md`, `PRIVACY.md`, and `DEPLOYMENT.md` document the support,
  privacy, permission, and Microsoft 365 admin deployment posture.
- `npm run check` performs local static validation and JavaScript syntax checks.
- `.github/workflows/ci.yml` runs the same validation in CI.

Before publishing or assigning the add-in to users, run:

```powershell
npm run check
```

### Security notes

- Host the add-in on an organization-controlled HTTPS origin. The manifest's
  `SourceLocation` runs with `ReadWriteDocument`, so deployment control matters.
- For enterprise hosting, send Content Security Policy as an HTTP response
  header too. The checked-in pages include CSP meta tags with `object-src
  'none'`, but the meta tag should be treated as defense in depth, not the
  only policy enforcement point.
- Keep `office.js` loaded from Microsoft. Keep other runtime assets, including
  the Markdown parser, self-hosted and pinned.
- Review `vendor/marked.min.js` updates deliberately. `npm run check` verifies
  the current vendored bundle by SHA-256.
- Keep inserted hyperlink protocols restricted to `http:`, `https:`, and
  `mailto:` unless there is a documented enterprise reason to allow more.
- Treat the 1 MB Markdown limit as a safety control for the Office task pane.
  Raise it only after testing large documents on Word for Windows, Mac, and web.
- This add-in has no backend, telemetry, cookies, or project-managed content
  storage. If you add any of those, update `PRIVACY.md` and the deployment
  review checklist before assigning it to users.

## Keyboard

- **Ctrl/Cmd + Enter** in the textarea triggers Insert.

## Caveats

- **Blockquotes** apply Word's `Quote` style to their child blocks. Word's
  built-in Quote style has theme-dependent behavior, so the visual result can
  vary across templates.
- **Inline images** are inserted as alt-text placeholders. Embedding requires
  fetching the URL, converting to base64, and calling
  `Range.insertInlinePictureFromBase64`. CORS will reject most public URLs
  unless they ship permissive headers.
- **Input size** is capped at 1 MB to keep parsing and insertion responsive in
  the Office task pane.
- **Tables** use `Grid Table 4 - Accent 1`, which exists in the default
  Office theme set. Custom templates may not have it; the code falls back
  silently to the default table style.
- **Built-in styles** vary by template. The code uses `try/catch` so missing
  styles fall back to `Normal` rather than throwing.

## Tested against

- Word for Microsoft 365 (Windows, Mac, Web) — Word.js requirement set 1.3+.
- `marked` 12.0.2, vendored in `vendor/marked.min.js`.

## File layout

```
word-md/
├── manifest.xml
├── taskpane.html
├── taskpane.css
├── taskpane.js
├── vendor/
│   └── marked.min.js
├── scripts/
│   └── validate.mjs
├── icons/
│   ├── icon-16.png   (replace placeholder)
│   ├── icon-32.png
│   └── icon-80.png
├── DEPLOYMENT.md
├── PRIVACY.md
├── SECURITY.md
└── README.md
```
