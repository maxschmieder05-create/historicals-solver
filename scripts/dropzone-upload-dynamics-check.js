const fs = require("node:fs");
const path = require("node:path");

const pagePath = path.join(__dirname, "..", "app", "page.tsx");
const stylesPath = path.join(__dirname, "..", "app", "styles.css");
const source = fs.readFileSync(pagePath, "utf8");
const styles = fs.readFileSync(stylesPath, "utf8");

const checks = [
  {
    ok: !/<label\s+[^>]*className=\{`dropzone/.test(source),
    message: "Dropzone must not be a label; label activation can swallow file-picker changes."
  },
  {
    ok: !/<button[\s\S]*className=\{`dropzone/.test(source),
    message: "Dropzone must not be a button that depends on scripted file-input clicks."
  },
  {
    ok: /<div[\s\S]*className=\{`dropzone[\s\S]*<input[\s\S]*className="fileInput"[\s\S]*type="file"/.test(source),
    message: "Native file input must live inside the dropzone so direct clicks open the picker."
  },
  {
    ok: /<input[\s\S]*type="file"[\s\S]*name="file"[\s\S]*onChange=\{handleFileSelect\}[\s\S]*onInput=\{handleFileInput\}/.test(source),
    message: "Native file input must remain wired to both change and input events."
  },
  {
    ok: !/onClick=\{\(event\)\s*=>\s*\{[\s\S]*event\.currentTarget\.value = ""/.test(source),
    message: "File input must not clear itself from onClick because some browsers dispatch that after the picker closes."
  },
  {
    ok: /onPointerDown=\{\(event\)\s*=>\s*\{[\s\S]*prepareFilePicker\(event\.currentTarget\)/.test(source),
    message: "File input should clear before picker activation through pointer down."
  },
  {
    ok: /\.fileInput\s*\{[\s\S]*inset:\s*0;[\s\S]*width:\s*100%;[\s\S]*height:\s*100%;[\s\S]*opacity:\s*0;/.test(styles),
    message: "File input must stay as a full-size transparent overlay, not a clipped hidden input."
  }
];

const failures = checks.filter((check) => !check.ok).map((check) => check.message);

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}

console.log("Dropzone upload dynamics guard passed.");
