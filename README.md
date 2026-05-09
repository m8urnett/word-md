# word-md — Word Add-in

A task pane that converts pasted Markdown (or a loaded `.md` file) to Word
content with **proper named styles applied** — Heading 1–6, List Bullet,
List Number, Quote, HTML Preformatted, real Word tables, hyperlinks.

## Why this works the way it does

The straightforward approach to this problem is to convert MD → HTML and call
`Range.insertHtml`. Word does interpret HTML, but its style mapping is
approximate: `<pre>` doesn't reliably land on **HTML Preformatted**,
`<blockquote>` styling depends on the active theme, and you can't address
custom or template-specific styles. The result feels like pasted-from-the-web
content, not native authored content.

This add-in instead walks `marked.lexer()`'s token tree and uses Word.js to
insert each block with an explicit named style. Inline runs (bold, italic,
code, link, strikethrough) are applied via `Range.font.*` after each
`insertText`. The output is indistinguishable from content typed by hand
using the ribbon — change the document's "Heading 2" definition and every
H2 in the inserted markdown updates with it.

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

## Sideload (development / personal use)

1. Host these files at any HTTPS endpoint. Easiest: GitHub Pages.
   - The bundled `manifest.xml` points at `https://m8urnett.github.io/word-md/`.
   - Push this folder to a repo, enable Pages, confirm that
     `https://m8urnett.github.io/word-md/taskpane.html` loads.
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
- `SECURITY.md`, `PRIVACY.md`, and `DEPLOYMENT.md` document the support,
  privacy, permission, and Microsoft 365 admin deployment posture.
- `npm run check` performs local static validation and JavaScript syntax checks.
- `.github/workflows/ci.yml` runs the same validation in CI.

Before publishing or assigning the add-in to users, run:

```powershell
npm run check
```

## Icons

Replace `icons/icon-{16,32,80}.svg` with PNGs at those exact sizes. Office
requires PNG for the ribbon. Keep them on transparent background so the
ribbon's hover state looks right.

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
