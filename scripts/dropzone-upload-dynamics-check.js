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
    ok: /<input[\s\S]*type="file"[\s\S]*name="file"[\s\S]*onChangeCapture=\{handleFileSelect\}[\s\S]*onInputCapture=\{handleFileInput\}[\s\S]*onChange=\{handleFileSelect\}[\s\S]*onInput=\{handleFileInput\}/.test(source),
    message: "Native file input must remain wired to capture and bubble change/input events."
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
    ok: /input\.addEventListener\("change", handleNativeFileSelection\)/.test(source)
      && /input\.addEventListener\("input", handleNativeFileSelection\)/.test(source)
      && /input\.addEventListener\("cancel", handleNativeFileSelection\)/.test(source),
    message: "File input must keep native DOM listeners as a fallback for browser-specific picker event timing, including same-file cancel events."
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
    ok: /const selectedFile =\s*selectedFileRef\.current\s*\?\?\s*file\s*\?\?\s*\(nativeFile instanceof File && nativeFile\.size > 0 \? nativeFile : null\)/.test(source),
    message: "Submit must prefer the synchronously stored selected file over the displayed state and native input value."
  },
  {
    ok: /const selectedFileRef = useRef<File \| null>\(null\);/.test(source)
      && /selectedFileRef\.current = nextFile;\s*setFile\(nextFile\);/.test(source),
    message: "A valid workbook selection must synchronously store the selected File before updating visible state."
  },
  {
    ok: /fileInputResetTimerRef\.current = window\.setTimeout\(\(\) => \{\s*clearFileInput\(\);\s*fileInputResetTimerRef\.current = null;\s*\}, 0\);/.test(source),
    message: "Valid workbook selection must defer the native input reset until after the selected File is stored."
  },
  {
    ok: !/fileKey\(nextFile\) === selectedFileKeyRef\.current/.test(source),
    message: "Focus-return sync must not skip a native file just because its key was seen before."
  },
  {
    ok: /function handlePageShow\(\)[\s\S]*syncInputSelectionSoon\(\);[\s\S]*syncInputSelectionSoon\(\);[\s\S]*window\.addEventListener\("pageshow", handlePageShow\)/.test(source),
    message: "Dropzone must resync from the native file input on mount and page-show after app reloads or browser restores."
  },
  {
    ok: !/prepareNativeFilePicker/.test(source)
      && !/function handleFilePicker(Pointer|Key)Down/.test(source)
      && !/onPointerDown=\{handleFilePickerPointerDown\}/.test(source)
      && !/onKeyDown=\{handleFilePickerKeyDown\}/.test(source),
    message: "Picker-opening pointer and keyboard events must not reset the native input before selection resolves."
  },
  {
    ok: !/(onClick|onPointerDown|onKeyDown)=\{\(event\)\s*=>\s*\{[\s\S]*event\.currentTarget\.value = ""/.test(source)
      && !/(onClick|onPointerDown|onKeyDown)=\{handleFilePicker/.test(source),
    message: "Native picker resets must stay out of picker-opening handlers that can erase picker results."
  },
  {
    ok: /selectedFileRef\.current = null;[\s\S]*setFile\(null\);[\s\S]*clearFileInput\(\);[\s\S]*not an \.xlsx workbook/.test(source),
    message: "Invalid workbook selections must clear both selected-file state and the native input."
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
