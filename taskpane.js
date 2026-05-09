/* Markdown Paste — Word task pane logic
 *
 * Strategy: parse Markdown to a token tree with `marked.lexer()`, then walk it
 * synchronously inside a single `Word.run` block, queueing Word.js operations
 * against a moving cursor (a `Range` proxy). Top-level blocks each get a
 * named Word style (Heading 1–6, List Bullet/Number, Quote, HTML Preformatted,
 * Normal). Inline runs (bold, italic, code, link, strikethrough) are applied
 * with `Range.font.*` after each `insertText`. One `context.sync()` at the end.
 */

(() => {
  "use strict";

  // Safety limit to prevent memory exhaustion during parsing or Word insertion.
  const MAX_MARKDOWN_BYTES = 1024 * 1024;
  const parser = new DOMParser();

  // ─────────────────────────────────────────────────────────────────────
  // Bootstrap
  // ─────────────────────────────────────────────────────────────────────

  // Initialize the add-in and bind UI events once the Office host is ready.
  Office.onReady(info => {
    if (info.host !== Office.HostType.Word) {
      setStatus("This add-in runs in Word.", "error");
      return;
    }

    document.getElementById("insert-btn").addEventListener("click", onInsert);
    document.getElementById("preview-btn").addEventListener("click", onPreview);
    document.getElementById("clear-btn").addEventListener("click", onClear);
    document.getElementById("file-input").addEventListener("change", onLoadFile);

    // Keyboard shortcut for power users: Ctrl/Cmd+Enter triggers the insert.
    document.getElementById("md-input").addEventListener("keydown", e => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        onInsert();
      }
    });

    // Configure the Markdown parser (Marked) with GitHub Flavored Markdown.
    if (window.marked) {
      marked.use({ gfm: true, breaks: false });
    }
  });

  // ─────────────────────────────────────────────────────────────────────
  // UI handlers
  // ─────────────────────────────────────────────────────────────────────

  // Handles local file loading, ensuring the file doesn't exceed the safety limit.
  function onLoadFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    if (file.size > MAX_MARKDOWN_BYTES) {
      e.target.value = "";
      setStatus(`File is too large. Limit is ${formatBytes(MAX_MARKDOWN_BYTES)}.`, "error");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      document.getElementById("md-input").value = reader.result;
      setStatus(`Loaded ${file.name} (${file.size.toLocaleString()} bytes).`, "ok");
    };
    reader.onerror = () => setStatus(`Could not read ${file.name}.`, "error");
    reader.readAsText(file);
  }

  // Generates a local HTML preview using safe DOM-building APIs.
  function onPreview() {
    setProcessing(true);
    try {
      const raw = document.getElementById("md-input").value;
      if (!isWithinSizeLimit(raw)) return;
      const md = preprocess(raw);
      // Use marked.lexer to get tokens, then build the preview DOM tree.
      const previewEl = document.getElementById("preview");
      previewEl.replaceChildren(window.marked ? renderPreview(marked.lexer(md)) : plainPreview(md));
      previewEl.hidden = false;
      setStatus("");
    } catch (err) {
      console.error(err);
      setStatus(`Preview error: ${err.message || err}`, "error");
    } finally {
      setProcessing(false);
    }
  }

  function onClear() {
    document.getElementById("md-input").value = "";
    document.getElementById("preview").hidden = true;
    setStatus("");
  }

  // Main entry point for inserting Markdown into the active Word document.
  async function onInsert() {
    const raw = document.getElementById("md-input").value;
    if (!raw.trim()) { setStatus("Nothing to insert.", "error"); return; }
    if (!window.marked) { setStatus("Markdown parser failed to load.", "error"); return; }
    if (!isWithinSizeLimit(raw)) return;

    let tokens;
    try {
      setProcessing(true);
      setStatus("Inserting…");
      const md = preprocess(raw);
      tokens = marked.lexer(md);
      
      // If the user checked "Replace selection", we delete the current range first.
      const replaceSel = document.getElementById("opt-replace").checked;

      await Word.run(async (context) => {
        const sel = context.document.getSelection();
        
        // We work with a 'cursor' (Range proxy) that moves forward as we insert content.
        if (replaceSel) sel.delete();

        let cursor = sel.getRange("End");
        for (const token of tokens) {
          cursor = insertBlock(cursor, token, /* listLevel */ 0, {});
        }
        await context.sync();
      });
      setStatus(`Inserted ${tokens.length} block${tokens.length === 1 ? "" : "s"}.`, "ok");
    } catch (err) {
      console.error(err);
      setStatus(`Error: ${err.message || err}`, "error");
    } finally {
      setProcessing(false);
    }
  }

  // ─────────────────────────────────────────────────────────────────────
  // Preprocessing
  // ─────────────────────────────────────────────────────────────────────

  // Applies optional transformations like stripping YAML frontmatter or converting to smart quotes.
  function preprocess(text) {
    let t = text;
    if (document.getElementById("opt-strip-frontmatter").checked) {
      t = t.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n?/, "");
    }
    // Smart quotes are applied via regex, avoiding changes inside code blocks.
    if (document.getElementById("opt-smart-quotes").checked) {
      t = smartQuotes(t);
    }
    return t;
  }

  function smartQuotes(s) {
    // Don't touch fenced or inline code
    const fences = [];
    s = s.replace(/```[\s\S]*?```/g, m => { fences.push(m); return `\u0000${fences.length - 1}\u0000`; });
    const codes = [];
    s = s.replace(/`[^`\n]+`/g, m => { codes.push(m); return `\u0001${codes.length - 1}\u0001`; });

    s = s
      .replace(/(^|[\s({\[\u2014\u2013])"/g, "$1\u201C")
      .replace(/"/g, "\u201D")
      .replace(/(^|[\s({\[\u2014\u2013])'/g, "$1\u2018")
      .replace(/'/g, "\u2019")
      .replace(/--/g, "\u2014")
      .replace(/\.\.\./g, "\u2026");

    s = s.replace(/\u0001(\d+)\u0001/g, (_, i) => codes[+i]);
    s = s.replace(/\u0000(\d+)\u0000/g, (_, i) => fences[+i]);
    return s;
  }

  // ─────────────────────────────────────────────────────────────────────
  // Block dispatch
  // ─────────────────────────────────────────────────────────────────────

  // Dispatches a Markdown block token to the appropriate Word.js insertion function.
  function insertBlock(anchor, token, listLevel, opts = {}) {
    switch (token.type) {
      case "heading":    return insertHeading(anchor, token, opts);
      case "paragraph":  return insertParagraph(anchor, token, opts);
      case "blockquote": return insertBlockquote(anchor, token);
      case "code":       return insertCode(anchor, token, opts);
      case "list":       return insertList(anchor, token, listLevel, opts);
      case "table":      return insertTable(anchor, token);
      case "hr":         return insertHr(anchor);
      case "space":      return anchor;
      case "html": {
        // Plain-text fallback for raw HTML; keep things predictable.
        const p = anchor.insertParagraph(stripHtml(token.text || ""), "After");
        p.style = opts.quote ? "Quote" : "Normal";
        return p.getRange("After");
      }
      default: {
        const p = anchor.insertParagraph(token.raw || token.text || "", "After");
        p.style = opts.quote ? "Quote" : "Normal";
        return p.getRange("After");
      }
    }
  }

  // Inserts a paragraph and applies Word's Heading 1-6 styles.
  function insertHeading(anchor, token, opts) {
    const depth = Math.min(Math.max(token.depth || 1, 1), 9);
    const p = anchor.insertParagraph("", "After");
    p.style = opts.quote ? "Quote" : `Heading ${depth}`;
    insertInlines(p, token.tokens || [{ type: "text", text: token.text }], {});
    return p.getRange("After");
  }

  // Inserts a standard paragraph with Normal style.
  function insertParagraph(anchor, token, opts) {
    const p = anchor.insertParagraph("", "After");
    p.style = opts.quote ? "Quote" : "Normal";
    insertInlines(p, token.tokens || [{ type: "text", text: token.text }], {});
    return p.getRange("After");
  }

  // Recursively inserts blocks within a blockquote, applying the "Quote" style.
  function insertBlockquote(anchor, token) {
    let cursor = anchor;
    const inner = token.tokens || [];
    for (const child of inner) {
      cursor = insertBlock(cursor, child, 0, { quote: true });
    }
    return cursor;
  }

  // Handles fenced code blocks by inserting one paragraph per line to preserve formatting.
  function insertCode(anchor, token, opts) {
    // Multi-line fenced block: one paragraph per line so line breaks render
    // correctly with HTML Preformatted style.
    const lines = (token.text || "").split(/\r?\n/);
    let cursor = anchor;
    for (const line of lines) {
      const p = cursor.insertParagraph(line, "After");
      // Try named style first; fall back to font if the style isn't in this template.
      try { p.style = opts.quote ? "Quote" : "HTML Preformatted"; }
      catch (_) { p.font.name = "Consolas"; p.font.size = 10; }
      cursor = p.getRange("After");
    }
    return cursor;
  }

  // Inserts lists (bulleted or numbered) with support for GFM checkboxes and nesting.
  function insertList(anchor, listToken, level, opts) {
    const styleName = listToken.ordered ? "List Number" : "List Bullet";
    let cursor = anchor;

    for (const item of listToken.items || []) {
      const { inlineTokens, nestedBlocks } = unpackListItem(item);

      const p = cursor.insertParagraph("", "After");
      try { p.style = opts.quote ? "Quote" : styleName; }
      catch (_) { p.style = "List Paragraph"; }

      // Word's "List Bullet 2/3" exists in some templates but not others.
      // Indent manually to support arbitrary nesting depth.
      if (level > 0) {
        try { p.leftIndent = level * 18; } catch (_) {}
      }

      if (item.task) {
        // GFM task list
        const r = p.insertText(item.checked ? "☒ " : "☐ ", "End");
        r.font.name = "Segoe UI Symbol";
      }

      insertInlines(p, inlineTokens, {});
      cursor = p.getRange("After");

      for (const nested of nestedBlocks) {
        cursor = insertBlock(cursor, nested, level + 1, opts);
      }
    }
    return cursor;
  }

  // Extracts inline tokens and nested block tokens from a list item's structure.
  function unpackListItem(item) {
    const inlineTokens = [];
    const nestedBlocks = [];
    for (const child of item.tokens || []) {
      if (child.type === "text") {
        if (child.tokens && child.tokens.length) inlineTokens.push(...child.tokens);
        else inlineTokens.push({ type: "text", text: child.text });
      } else if (child.type === "paragraph") {
        inlineTokens.push(...(child.tokens || []));
      } else {
        nestedBlocks.push(child);
      }
    }
    return { inlineTokens, nestedBlocks };
  }

  // Converts a Markdown table token into a native Word table with a built-in style.
  function insertTable(anchor, token) {
    const headerCells = (token.header || []).map(h => h.text || "");
    const dataRows    = (token.rows   || []).map(row => row.map(c => c.text || ""));
    const allRows     = [headerCells, ...dataRows];
    const cols        = headerCells.length;
    if (cols === 0) return anchor;

    const table = anchor.insertTable(allRows.length, cols, "After", allRows);
    try { table.styleBuiltIn = Word.BuiltInStyleName.gridTable4_Accent1; }
    catch (_) {
      try { table.style = "Grid Table 4 - Accent 1"; }
      catch (__) { /* leave default */ }
    }
    return table.getRange("After");
  }

  // Inserts a horizontal rule by applying a bottom border to an empty paragraph.
  function insertHr(anchor) {
    const p = anchor.insertParagraph("", "After");
    p.style = "Normal";
    // Bottom border on an empty paragraph reads as a horizontal rule in Word.
    try {
      const border = p.getBorder("Bottom");
      border.type = "Single";
      border.width = "Pt050";
      border.color = "#999999";
    } catch (_) {
      // Fallback: a row of em-dashes
      p.insertText("\u2014".repeat(40), "End");
    }
    return p.getRange("After");
  }

  // ─────────────────────────────────────────────────────────────────────
  // Inline runs
  // ─────────────────────────────────────────────────────────────────────

  // Iterates through inline tokens (bold, italic, links) and applies them to the current paragraph.
  function insertInlines(paragraph, tokens, format) {
    for (const t of tokens) {
      switch (t.type) {
        case "text":
        case "escape": {
          if (t.tokens && t.tokens.length) {
            insertInlines(paragraph, t.tokens, format);
          } else {
            const r = paragraph.insertText(decodeEntities(t.text || ""), "End");
            applyFormat(r, format);
          }
          break;
        }
        case "strong":   insertInlines(paragraph, t.tokens || [], { ...format, bold: true }); break;
        case "em":       insertInlines(paragraph, t.tokens || [], { ...format, italic: true }); break;
        case "del":      insertInlines(paragraph, t.tokens || [], { ...format, strike: true }); break;
        case "codespan": {
          const r = paragraph.insertText(decodeEntities(t.text || ""), "End");
          applyFormat(r, { ...format, code: true });
          break;
        }
        case "link": {
          const text = extractText(t) || t.href || "";
          const r = paragraph.insertText(text, "End");
          applyFormat(r, format);
          // Only allow safe protocols to prevent XSS via links.
          const href = getSafeUrl(t.href);
          if (href) {
            try { r.hyperlink = href; } catch (_) {}
          }
          break;
        }
        case "image": {
          // Embedding remote images requires fetch + base64; keep v1 simple.
          const alt = t.text || "";
          const r = paragraph.insertText(alt ? `[${alt}]` : `[image: ${t.href || ""}]`, "End");
          applyFormat(r, { ...format, italic: true });
          break;
        }
        case "br":       paragraph.insertText("\v", "End"); break;          // soft break
        case "html":     paragraph.insertText(stripHtml(t.text || ""), "End"); break;
        case "autolink": {
          const r = paragraph.insertText(t.text || t.href, "End");
          applyFormat(r, format);
          const href = getSafeUrl(t.href);
          if (href) { try { r.hyperlink = href; } catch (_) {} }
          break;
        }
        default: {
          if (t.tokens) insertInlines(paragraph, t.tokens, format);
          else if (t.text) {
            const r = paragraph.insertText(decodeEntities(t.text), "End");
            applyFormat(r, format);
          }
        }
      }
    }
  }

  // Helper to apply font-level formatting (bold, italic, etc.) to a Word Range.
  function applyFormat(range, fmt) {
    if (!fmt) return;
    if (fmt.bold)   range.font.bold = true;
    if (fmt.italic) range.font.italic = true;
    if (fmt.strike) range.font.strikeThrough = true;
    if (fmt.code) {
      range.font.name = "Consolas";
      try { range.font.highlightColor = "#F3F2F1"; } catch (_) {}
    }
  }

  // ─────────────────────────────────────────────────────────────────────
  // Helpers
  // ─────────────────────────────────────────────────────────────────────

  // Recursively flattens a token's children into a plain string.
  function extractText(token) {
    if (!token) return "";
    if (token.tokens) return token.tokens.map(extractText).join("");
    return token.text || "";
  }

  // Sanitizes raw HTML by extracting only the text content using DOMParser.
  function stripHtml(html) {
    const doc = parser.parseFromString(html, "text/html");
    return doc.body.textContent || "";
  }

  // Decodes HTML entities (e.g., &amp;) to plain characters for Word.js insertion.
  function decodeEntities(s) {
    if (!s || (s.indexOf("&") === -1)) return s;
    const doc = parser.parseFromString(s, "text/html");
    return doc.documentElement.textContent || s;
  }

  function plainPreview(text) {
    const pre = document.createElement("pre");
    pre.textContent = text;
    return pre;
  }

  // Builds a safe DocumentFragment for the task pane preview.
  function renderPreview(tokens) {
    const fragment = document.createDocumentFragment();
    for (const token of tokens) {
      fragment.appendChild(renderPreviewBlock(token));
    }
    return fragment;
  }

  // Maps Markdown tokens to safe HTML elements for the preview.
  function renderPreviewBlock(token) {
    switch (token.type) {
      case "heading": {
        const h = document.createElement(`h${Math.min(Math.max(token.depth || 1, 1), 6)}`);
        appendPreviewInlines(h, token.tokens || [{ type: "text", text: token.text }]);
        return h;
      }
      case "paragraph": {
        const p = document.createElement("p");
        appendPreviewInlines(p, token.tokens || [{ type: "text", text: token.text }]);
        return p;
      }
      case "blockquote": {
        const bq = document.createElement("blockquote");
        for (const child of token.tokens || []) {
          bq.appendChild(renderPreviewBlock(child));
        }
        return bq;
      }
      case "code": {
        const pre = document.createElement("pre");
        const code = document.createElement("code");
        code.textContent = token.text || "";
        pre.appendChild(code);
        return pre;
      }
      case "list": {
        const list = document.createElement(token.ordered ? "ol" : "ul");
        for (const item of token.items || []) {
          const li = document.createElement("li");
          const { inlineTokens, nestedBlocks } = unpackListItem(item);
          if (item.task) li.appendChild(document.createTextNode(item.checked ? "[x] " : "[ ] "));
          appendPreviewInlines(li, inlineTokens);
          for (const nested of nestedBlocks) {
            li.appendChild(renderPreviewBlock(nested));
          }
          list.appendChild(li);
        }
        return list;
      }
      case "table": {
        return renderPreviewTable(token);
      }
      case "hr": {
        return document.createElement("hr");
      }
      case "space": {
        return document.createTextNode("");
      }
      case "html": {
        const p = document.createElement("p");
        p.textContent = stripHtml(token.text || "");
        return p;
      }
      default: {
        const p = document.createElement("p");
        p.textContent = token.raw || token.text || "";
        return p;
      }
    }
  }

  // Generates the preview HTML table from Markdown table tokens.
  function renderPreviewTable(token) {
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    for (const cell of token.header || []) {
      const th = document.createElement("th");
      appendPreviewInlines(th, cell.tokens || [{ type: "text", text: cell.text }]);
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (const row of token.rows || []) {
      const tr = document.createElement("tr");
      for (const cell of row) {
        const td = document.createElement("td");
        appendPreviewInlines(td, cell.tokens || [{ type: "text", text: cell.text }]);
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    return table;
  }

  // Builds the inline content (text, strong, em, etc.) for a preview element using safe DOM APIs.
  function appendPreviewInlines(parent, tokens) {
    for (const token of tokens || []) {
      switch (token.type) {
        case "text":
        case "escape": {
          if (token.tokens && token.tokens.length) appendPreviewInlines(parent, token.tokens);
          else parent.appendChild(document.createTextNode(decodeEntities(token.text || "")));
          break;
        }
        case "strong": {
          const el = document.createElement("strong");
          appendPreviewInlines(el, token.tokens);
          parent.appendChild(el);
          break;
        }
        case "em": {
          const el = document.createElement("em");
          appendPreviewInlines(el, token.tokens);
          parent.appendChild(el);
          break;
        }
        case "del": {
          const el = document.createElement("s");
          appendPreviewInlines(el, token.tokens);
          parent.appendChild(el);
          break;
        }
        case "codespan": {
          const el = document.createElement("code");
          el.textContent = decodeEntities(token.text || "");
          parent.appendChild(el);
          break;
        }
        case "link":
        case "autolink": {
          const el = document.createElement("a");
          el.textContent = extractText(token) || token.text || token.href || "";
          const href = getSafeUrl(token.href);
          if (href) {
            el.href = href;
            el.rel = "noreferrer";
          }
          parent.appendChild(el);
          break;
        }
        case "image": {
          parent.appendChild(document.createTextNode(token.text ? `[${token.text}]` : `[image: ${token.href || ""}]`));
          break;
        }
        case "br": {
          parent.appendChild(document.createElement("br"));
          break;
        }
        case "html": {
          parent.appendChild(document.createTextNode(stripHtml(token.text || "")));
          break;
        }
        default: {
          if (token.tokens) appendPreviewInlines(parent, token.tokens);
          else if (token.text) parent.appendChild(document.createTextNode(decodeEntities(token.text)));
        }
      }
    }
  }

  // Validates a URL against an allowlist of protocols (http, https, mailto).
  function getSafeUrl(url) {
    if (!url) return false;
    try {
      const parsed = new URL(url.trim(), window.location.href);
      if (!["http:", "https:", "mailto:"].includes(parsed.protocol)) return "";
      return parsed.href;
    } catch (_) {
      return "";
    }
  }

  function isWithinSizeLimit(text) {
    const bytes = new Blob([text]).size;
    if (bytes <= MAX_MARKDOWN_BYTES) return true;
    setStatus(`Markdown is too large. Limit is ${formatBytes(MAX_MARKDOWN_BYTES)}.`, "error");
    return false;
  }

  // Toggles UI buttons during long-running operations.
  function setProcessing(isProcessing) {
    for (const id of ["insert-btn", "preview-btn", "clear-btn", "file-input"]) {
      document.getElementById(id).disabled = isProcessing;
    }
  }

  function formatBytes(bytes) {
    return `${Math.round(bytes / 1024 / 1024)} MB`;
  }

  function setStatus(msg, kind) {
    const el = document.getElementById("status");
    el.textContent = msg || "";
    el.className = "status" + (kind ? " " + kind : "");
  }
})();
