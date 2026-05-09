import { readFileSync, statSync } from "node:fs";
import { relative } from "node:path";
import { fileURLToPath } from "node:url";

const root = new URL("../", import.meta.url);
const rootPath = fileURLToPath(root);
const read = (path) => readFileSync(new URL(path, root), "utf8");
const exists = (path) => statSync(new URL(path, root)).isFile();
const failures = [];

function check(condition, message) {
  if (!condition) failures.push(message);
}

function checkFile(path, predicate, message) {
  const contents = read(path);
  check(predicate(contents), `${path}: ${message}`);
}

check(exists("vendor/marked.min.js"), "vendor/marked.min.js must exist");
check(statSync(new URL("vendor/marked.min.js", root)).size > 10000, "vendor/marked.min.js looks too small");

checkFile("taskpane.html", (s) => s.includes('src="vendor/marked.min.js"'), "must load the vendored Markdown parser");
checkFile("taskpane.html", (s) => !s.includes("cdn.jsdelivr.net"), "must not load marked from jsDelivr");
checkFile("taskpane.html", (s) => s.includes("Content-Security-Policy"), "must declare a task pane CSP");
checkFile("index.html", (s) => !s.includes("fonts.googleapis.com"), "must not load Google Fonts");
checkFile("index.html", (s) => !s.includes("jsDelivr"), "must not mention jsDelivr as a runtime dependency");
checkFile("taskpane.js", (s) => !s.includes("previewEl.innerHTML"), "preview must not inject Markdown HTML");
checkFile("taskpane.js", (s) => s.includes("renderPreview(marked.lexer(md))"), "preview must use the safe renderer");
checkFile("taskpane.js", (s) => s.includes('{ quote: true }'), "blockquotes must route through Quote styling");
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
