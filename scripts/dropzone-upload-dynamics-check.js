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
    ok: /function openWorkbookPicker\(\)[\s\S]*fileInputRef\.current\?\.click\(\);/.test(source)
      && /<button className="browseCue" type="button" onClick=\{openWorkbookPicker\}/.test(source),
    message: "The visible Choose workbook button must open the native input from a direct user action."
  },
  {
    ok: /function handleDropzoneClick\(event: MouseEvent<HTMLElement>\)[\s\S]*target\.closest\("button, input"\)[\s\S]*openWorkbookPicker\(\)/.test(source)
      && /function handleDropzoneKeyDown\(event: KeyboardEvent<HTMLElement>\)[\s\S]*event\.key !== "Enter"[\s\S]*event\.key !== " "[\s\S]*openWorkbookPicker\(\)/.test(source)
      && /data-testid="workbook-dropzone"[\s\S]*role="button"[\s\S]*onClick=\{handleDropzoneClick\}[\s\S]*onKeyDown=\{handleDropzoneKeyDown\}/.test(source),
    message: "The entire dropbox must open the picker by mouse or keyboard without double-opening from its nested controls."
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
    ok: /\.fileInput\s*\{[\s\S]*width:\s*1px;[\s\S]*height:\s*1px;[\s\S]*clip-path:\s*inset\(50%\)/.test(styles)
      && !/\.fileInput\s*\{[\s\S]*inset:\s*0;[\s\S]*width:\s*100%;[\s\S]*height:\s*100%/.test(styles),
    message: "The native input must be visually hidden instead of intercepting the entire drop target."
  }
];

const failures = checks.filter((check) => !check.ok).map((check) => check.message);
if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}

console.log("Dropzone upload dynamics guard passed.");
