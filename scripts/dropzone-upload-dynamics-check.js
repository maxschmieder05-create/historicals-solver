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
    ok: /function hasTransferredFiles\(dataTransfer: DataTransfer \| null\)[\s\S]*dataTransfer\.types[\s\S]*includes\("Files"\)[\s\S]*dataTransfer\.items[\s\S]*item\.kind === "file"/.test(source),
    message: "File drag detection must accept both DataTransfer.types and DataTransfer.items file payloads."
  },
  {
    ok: /function workbookFileFromTransfer\(dataTransfer: DataTransfer \| null\)[\s\S]*Array\.from\(dataTransfer\.files \?\? \[\]\)\[0\][\s\S]*item\.getAsFile\(\)/.test(source),
    message: "Dropped workbook extraction must fall back to DataTransferItem.getAsFile when files is empty."
  },
  {
    ok: /input\.addEventListener\("change", handleNativeFileSelection\)/.test(source) && /input\.addEventListener\("input", handleNativeFileSelection\)/.test(source),
    message: "File input must keep native DOM listeners as a fallback for browser-specific picker event timing."
  },
  {
    ok: /const handleWorkbookSelected = useCallback/.test(source)
      && /handleWorkbookSelected\(workbookFileFromTransfer\(event\.dataTransfer\)\)/.test(source)
      && /handleWorkbookSelected\(input\.files\?\.item\(0\) \?\? undefined\)/.test(source),
    message: "File picker and drag/drop paths must share the same workbook-selection handler."
  },
  {
    ok: /handleWorkbookSelected\(workbookFileFromTransfer\(transfer\)\)/.test(source),
    message: "Window-level file drops must use the normalized drag payload extractor."
  },
  {
    ok: /const selectedFile =\s*file\s*\?\?\s*\(nativeFile instanceof File && nativeFile\.size > 0 \? nativeFile : null\)/.test(source),
    message: "Submit must prefer the displayed selected file state over the native input value."
  },
  {
    ok: /const resetFileInputAfterSelection = useCallback/.test(source)
      && /window\.setTimeout\(\(\) => \{\s*clearFileInput\(\);/.test(source)
      && /setFile\(nextFile\);\s*resetFileInputAfterSelection\(\);/.test(source),
    message: "A valid workbook selection must update visible state before deferring the native input reset for same-file reselection."
  },
  {
    ok: !/setFile\(nextFile\);\s*clearFileInput\(\);/.test(source),
    message: "Valid workbook selection must not clear the native input synchronously during the picker event."
  },
  {
    ok: !/fileKey\(nextFile\) === selectedFileKeyRef\.current/.test(source),
    message: "Focus-return sync must not skip a native file just because its key was seen before."
  },
  {
    ok: !/(onClick|onPointerDown|onKeyDown)=\{\(event\)\s*=>\s*\{[\s\S]*event\.currentTarget\.value = ""/.test(source),
    message: "File input must not clear itself while opening the picker; that can erase the selected file after Open."
  },
  {
    ok: !/prepareFilePicker/.test(source),
    message: "Dropzone should not use a picker-preparation helper that mutates the native input value."
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
