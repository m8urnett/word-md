import { existsSync, readFileSync, statSync } from "node:fs";
import { relative } from "node:path";
import { fileURLToPath } from "node:url";
import { createHash } from "node:crypto";

const root = new URL("../", import.meta.url);
const rootPath = fileURLToPath(root);
const read = (path) => readFileSync(new URL(path, root), "utf8");
const EXPECTED_MARKED_SHA256 = "15fabce5b65898b32b03f5ed25e9f891a729ad4c0d6d877110a7744aa847a894";
const failures = [];

function check(condition, message) {
  if (!condition) failures.push(message);
}

function checkFile(path, predicate, message) {
  const contents = read(path);
  check(predicate(contents), `${path}: ${message}`);
}

const markedBundleUrl = new URL("vendor/marked.min.js", root);
if (existsSync(markedBundleUrl)) {
  const markedBundle = readFileSync(markedBundleUrl);
  check(statSync(markedBundleUrl).isFile(), "vendor/marked.min.js must exist");
  check(markedBundle.length > 10000, "vendor/marked.min.js looks too small");
  check(
    createHash("sha256").update(markedBundle).digest("hex") === EXPECTED_MARKED_SHA256,
    "vendor/marked.min.js SHA-256 does not match the pinned marked 12.0.2 bundle"
  );
} else {
  check(false, "vendor/marked.min.js must exist");
}

checkFile("taskpane.html", (s) => s.includes('src="vendor/marked.min.js"'), "must load the vendored Markdown parser");
checkFile("taskpane.html", (s) => !s.includes("cdn.jsdelivr.net"), "must not load marked from jsDelivr");
checkFile("taskpane.html", (s) => s.includes("Content-Security-Policy"), "must declare a task pane CSP");
checkFile("taskpane.html", (s) => s.includes("object-src 'none'"), "task pane CSP must block object embeds");
checkFile("index.html", (s) => s.includes("object-src 'none'"), "landing page CSP must block object embeds");
checkFile("index.html", (s) => !s.includes("fonts.googleapis.com"), "must not load Google Fonts");
checkFile("index.html", (s) => !s.includes("jsDelivr"), "must not mention jsDelivr as a runtime dependency");
checkFile("taskpane.js", (s) => !s.includes("previewEl.innerHTML"), "preview must not inject Markdown HTML");
checkFile("taskpane.js", (s) => s.includes("renderPreview(marked.lexer(md))"), "preview must use the safe renderer");
checkFile("taskpane.js", (s) => s.includes('{ quote: true }'), "blockquotes must route through Quote styling");
checkFile("taskpane.js", (s) => s.includes("MAX_MARKDOWN_BYTES"), "must cap Markdown input size");
checkFile("taskpane.js", (s) => s.includes("getSafeUrl(t.href)"), "inserted hyperlinks must use protocol filtering");
checkFile("manifest.xml", (s) => s.includes("<Permissions>ReadWriteDocument</Permissions>"), "must explicitly declare document write permission");
checkFile("SECURITY.md", (s) => /security contact/i.test(s), "must include a security contact section");
checkFile("PRIVACY.md", (s) => s.includes("No pasted Markdown content"), "must document content handling");
checkFile("DEPLOYMENT.md", (s) => s.includes("Microsoft 365 admin center"), "must document tenant deployment");

const manifest = read("manifest.xml");
for (const tag of ["SourceLocation", "IconUrl", "HighResolutionIconUrl", "SupportUrl"]) {
  check(manifest.includes(`<${tag}`), `manifest.xml: missing ${tag}`);
}

if (failures.length) {
  console.error("Validation failed:");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log(`Validated ${relative(process.cwd(), rootPath) || "."}`);
