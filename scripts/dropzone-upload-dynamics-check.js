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
      && /<label className="dropzonePicker" htmlFor="model-template-file">/.test(source)
      && /\.fileInput\s*\{[\s\S]*width:\s*1px;[\s\S]*clip:\s*rect\(0, 0, 0, 0\);/.test(styles),
    message: "The visible picker must be an explicit label for one visually hidden native file input."
  },
  {
    ok: /function handleFileSelect\(event: ChangeEvent<HTMLInputElement>\)[\s\S]*const selected = event\.currentTarget\.files\?\.item\(0\);[\s\S]*handleWorkbookSelected\(selected\)/.test(source),
    message: "Native picker change must synchronously populate shared workbook state."
  },
  {
    ok: /handleWorkbookSelected\(selected\);[\s\S]*event\.currentTarget\.value = "";/.test(source)
      && !/prepareFilePicker|pickerSyncTimersRef|handlePickerReturn/.test(source),
    message: "The selected File must be retained before the native input resets for same-file reselection."
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
    ok: /<label className="dropzonePicker" htmlFor="model-template-file">[\s\S]*<span className="browseCue" aria-hidden="true">/.test(source)
      && !/openWorkbookPicker|handleDropzoneClick|handleDropzoneKeyDown/.test(source),
    message: "The full visible dropzone must activate its associated native input without scripted picker activation."
  }
];

const failures = checks.filter((check) => !check.ok).map((check) => check.message);
if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}

console.log("Dropzone upload dynamics guard passed.");
