const fs = require("node:fs");
const path = require("node:path");

const pagePath = path.join(__dirname, "..", "app", "page.tsx");
const stylesPath = path.join(__dirname, "..", "app", "styles.css");
const source = fs.readFileSync(pagePath, "utf8");
const styles = fs.readFileSync(stylesPath, "utf8");

const checks = [
  {
    ok: /<input[\s\S]*className="fileInput"[\s\S]*type="file"[\s\S]*name="file"[\s\S]*accept=\{SUPPORTED_WORKBOOK_ACCEPT\}[\s\S]*onClick=\{handleFilePickerOpen\}[\s\S]*onInput=\{handleFileSelect\}[\s\S]*onChange=\{handleFileSelect\}/.test(source),
    message: "A single native .xlsx file input must own picker selection."
  },
  {
    ok: /data-testid="workbook-dropzone"[\s\S]*<label className="dropzonePicker">[\s\S]*<input[\s\S]*className="fileInput"[\s\S]*type="file"/.test(source)
      && /\.fileInput\s*\{[\s\S]*width:\s*min\(100%, 520px\);[\s\S]*border:\s*1px solid var\(--line\);/.test(styles)
      && !/\.fileInput\s*\{[\s\S]*clip:\s*rect\(/.test(styles),
    message: "The browser-managed file control must remain visible so its native filename is always shown."
  },
  {
    ok: /function handleFileSelect\(event: SyntheticEvent<HTMLInputElement>\)[\s\S]*syncFileInputSelection\(event\.currentTarget\)/.test(source),
    message: "Native picker input/change events must synchronously populate shared workbook state."
  },
  {
    ok: /function handleFilePickerOpen\(event: SyntheticEvent<HTMLInputElement>\)[\s\S]*event\.currentTarget\.value = "";[\s\S]*syncFileInputSelectionSoon\(event\.currentTarget\)/.test(source),
    message: "The native input must reset before opening and resync after the file picker returns."
  },
  {
    ok: /input\.addEventListener\("input", handleNativeSelection\);[\s\S]*input\.addEventListener\("change", handleNativeSelection\);[\s\S]*window\.addEventListener\("focus", handleWindowFocus\);/.test(source),
    message: "Native file events and picker-return focus must provide fallbacks when synthetic change events are missed."
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
    ok: /function workbookNameFromInputValue\(value: string\)/.test(source)
      && /const selectedName = selected\?\.name \?\? workbookNameFromInputValue\(input\?\.value \?\? ""\);/.test(source)
      && /if \(selectedName\) setNativeFileName\(selectedName\);/.test(source)
      && /const selectedWorkbookName = file\?\.name \?\? nativeFileName;/.test(source),
    message: "The dropzone must display the native selected filename even while the browser is still exposing the File object."
  },
  {
    ok: /const selectedFile =\s*selectedFileRef\.current\s*\?\?\s*file\s*\?\?\s*\(nativeFile instanceof File && nativeFile\.size > 0 \? nativeFile : null\)/.test(source),
    message: "Submit must use the selected workbook even if a render has not completed yet."
  },
  {
    ok: /<label className="dropzonePicker">[\s\S]*<input[\s\S]*id="model-template-file"[\s\S]*type="file"/.test(source)
      && !/openWorkbookPicker|handleDropzoneClick|handleDropzoneKeyDown/.test(source),
    message: "The full visible dropzone must contain and activate the native picker without scripted activation."
  }
];

const failures = checks.filter((check) => !check.ok).map((check) => check.message);
if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}

console.log("Dropzone upload dynamics guard passed.");
