const fs = require("node:fs");
const path = require("node:path");

const pagePath = path.join(__dirname, "..", "app", "page.tsx");
const source = fs.readFileSync(pagePath, "utf8");

const checks = [
  {
    ok: !/<label\s+[^>]*className=\{`dropzone/.test(source),
    message: "Dropzone must not be a label; label activation can swallow file-picker changes."
  },
  {
    ok: /<button\s+[^>]*type="button"[\s\S]*className=\{`dropzone/.test(source),
    message: "Dropzone should be an explicit non-submit button."
  },
  {
    ok: /function openFilePicker\(\)[\s\S]*fileInputRef\.current\.value = ""[\s\S]*fileInputRef\.current\?\.click\(\)/.test(source),
    message: "openFilePicker should clear the native input then click it exactly through the ref."
  },
  {
    ok: /<input[\s\S]*type="file"[\s\S]*name="file"[\s\S]*onChange=\{handleFileSelect\}[\s\S]*onInput=\{handleFileInput\}/.test(source),
    message: "Hidden file input must remain wired to both change and input events."
  }
];

const failures = checks.filter((check) => !check.ok).map((check) => check.message);

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}

console.log("Dropzone upload dynamics guard passed.");
