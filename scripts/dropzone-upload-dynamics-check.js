const fs = require("node:fs");
const path = require("node:path");

const pagePath = path.join(__dirname, "..", "app", "page.tsx");
const stylesPath = path.join(__dirname, "..", "app", "styles.css");
const source = fs.readFileSync(pagePath, "utf8");
const styles = fs.readFileSync(stylesPath, "utf8");

const checks = [
  {
    ok: /<input[\s\S]*className="fileInput"[\s\S]*type="file"[\s\S]*name="file"[\s\S]*accept=\{SUPPORTED_WORKBOOK_ACCEPT\}[\s\S]*onChange=\{handleFileSelect\}/.test(source),
    message: "A single native .xlsx file input must own picker selection."
  },
  {
    ok: /data-testid="workbook-dropzone"[\s\S]*<input[\s\S]*className="fileInput"[\s\S]*type="file"/.test(source)
      && /\.fileInput\s*\{[\s\S]*inset:\s*0;[\s\S]*z-index:\s*2;[\s\S]*width:\s*100%;[\s\S]*height:\s*100%;[\s\S]*opacity:\s*0;/.test(styles),
    message: "One native file input must cover the entire drop target so clicks, keyboard focus, and OS file drops use the browser upload control."
  },
  {
    ok: /function handleFileSelect\(event: ChangeEvent<HTMLInputElement>\)[\s\S]*handleWorkbookSelected\(event\.currentTarget\.files\?\.item\(0\) \?\? undefined\)/.test(source),
    message: "Native picker change must synchronously populate shared workbook state."
  },
  {
    ok: !/onChangeCapture=|onInputCapture=|onInput=\{handleFileInput\}|handleFilePickerActivation|syncInputSelectionSoon|inputSyncTimersRef|inputSyncFrameRef/.test(source),
    message: "Picker selection must not be raced by duplicate capture handlers or focus timers."
  },
  {
    ok: /onDragEnter=\{handleDrag\}/.test(source)
      && /onDragOver=\{handleDrag\}/.test(source)
      && /onDragLeave=\{handleDrag\}/.test(source)
      && /onDrop=\{handleDrop\}/.test(source)
      && !/onDropCapture=/.test(source),
    message: "The drop target must own one ordinary drag/drop event path."
  },
  {
    ok: /function workbookFileFromTransfer\(dataTransfer: DataTransfer \| null\)[\s\S]*listedFiles\.find\(isSupportedWorkbookFile\)[\s\S]*item\.getAsFile\(\)/.test(source)
      && /handleDroppedWorkbook\(event\.dataTransfer\)/.test(source),
    message: "Dropped files must be normalized and sent through the shared workbook selection handler."
  },
  {
    ok: /window\.addEventListener\("drop", handleWindowDrop\)/.test(source)
      && /handleDroppedWorkbook\(transfer\)/.test(source),
    message: "A file dropped anywhere on the page must still populate the workbook."
  },
  {
    ok: /const selectedFileRef = useRef<File \| null>\(null\);/.test(source)
      && /selectedFileRef\.current = nextFile;\s*setFile\(nextFile\);/.test(source),
    message: "A valid workbook must update synchronous submission state and visible React state."
  },
  {
    ok: /const selectedFile =\s*selectedFileRef\.current\s*\?\?\s*file\s*\?\?\s*\(nativeFile instanceof File && nativeFile\.size > 0 \? nativeFile : null\)/.test(source),
    message: "Submit must use the selected workbook even if a render has not completed yet."
  },
  {
    ok: /\.dropzone\s*>\s*:not\(\.fileInput\)\s*\{[\s\S]*pointer-events:\s*none;/.test(styles)
      && /<span className="browseCue" aria-hidden="true">/.test(source)
      && !/openWorkbookPicker|handleDropzoneClick|handleDropzoneKeyDown/.test(source),
    message: "Decorative dropzone content must not intercept the native upload control or depend on scripted picker activation."
  }
];

const failures = checks.filter((check) => !check.ok).map((check) => check.message);
if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}

console.log("Dropzone upload dynamics guard passed.");
