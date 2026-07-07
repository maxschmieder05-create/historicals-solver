"use client";

import {
  ChangeEvent,
  DragEvent,
  FormEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState
} from "react";
import { ArrowDownToLine, CheckCircle2, FileCheck2, FileSpreadsheet, Loader2, Search, ShieldCheck, UploadCloud } from "lucide-react";

type FillSummary = {
  companyName: string;
  ticker: string;
  periods: string[];
  filledCells: number;
  commentsAdded: number;
  warnings: string[];
  debugLogPath?: string;
};

type FillError = {
  message: string;
  debugLogPath?: string;
};

const SUPPORTED_WORKBOOK_EXTENSIONS = [".xlsx", ".xlsm"] as const;
const SUPPORTED_WORKBOOK_ACCEPT = [
  ".xlsx",
  ".xlsm",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel.sheet.macroEnabled.12"
].join(",");

function isSupportedWorkbookFile(file: File) {
  const fileName = file.name.toLowerCase();
  return SUPPORTED_WORKBOOK_EXTENSIONS.some((extension) => fileName.endsWith(extension));
}

function sameWorkbookFile(left: File | null, right: File | null) {
  if (!left || !right) return false;
  return left.name === right.name && left.size === right.size && left.lastModified === right.lastModified;
}

function hasTransferredFiles(dataTransfer: DataTransfer | null) {
  if (!dataTransfer) return false;
  const types = Array.from(dataTransfer.types ?? []).map((type) => type.toLowerCase());
  if (types.includes("files") || types.includes("application/x-moz-file") || types.includes("public.file-url")) return true;
  return Array.from(dataTransfer.items ?? []).some((item) => item.kind === "file");
}

function workbookFileFromTransfer(dataTransfer: DataTransfer | null) {
  if (!dataTransfer) return null;
  const listedFiles = Array.from(dataTransfer.files ?? []);
  const listedWorkbook = listedFiles.find(isSupportedWorkbookFile);
  if (listedWorkbook) return listedWorkbook;
  if (listedFiles[0]) return listedFiles[0];
  const itemFiles: File[] = [];
  for (const item of Array.from(dataTransfer.items ?? [])) {
    if (item.kind !== "file") continue;
    const itemFile = item.getAsFile();
    if (itemFile) itemFiles.push(itemFile);
  }
  return itemFiles.find(isSupportedWorkbookFile) ?? itemFiles[0] ?? null;
}

function markWorkbookDropEffect(dataTransfer: DataTransfer | null) {
  if (!hasTransferredFiles(dataTransfer)) return false;
  if (dataTransfer) dataTransfer.dropEffect = "copy";
  return true;
}

export default function Home() {
  const [ticker, setTicker] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [summary, setSummary] = useState<FillSummary | null>(null);
  const [error, setError] = useState<FillError | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const selectedFileRef = useRef<File | null>(null);
  const dragDepthRef = useRef(0);
  const inputSyncFrameRef = useRef<number | null>(null);
  const inputSyncTimersRef = useRef<number[]>([]);

  const canSubmit = useMemo(() => !isSubmitting, [isSubmitting]);

  const clearFileInput = useCallback(() => {
    if (!fileInputRef.current) return;
    fileInputRef.current.value = "";
  }, []);

  function formatFileSize(bytes: number) {
    if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024)).toLocaleString()} KB`;
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
  }

  const handleWorkbookSelected = useCallback((nextFile?: File | null) => {
    if (!nextFile) return;
    const fileChanged = !sameWorkbookFile(selectedFileRef.current, nextFile);
    if (fileChanged) {
      setError(null);
      setSummary(null);
    }
    if (!isSupportedWorkbookFile(nextFile)) {
      selectedFileRef.current = null;
      setFile(null);
      clearFileInput();
      setError({ message: `${nextFile.name} is not a supported .xlsx or .xlsm workbook.` });
      return;
    }
    selectedFileRef.current = nextFile;
    setFile(nextFile);
  }, [clearFileInput]);

  const handleDroppedWorkbook = useCallback((dataTransfer: DataTransfer | null) => {
    const droppedFile = workbookFileFromTransfer(dataTransfer);
    if (!droppedFile) {
      setError({ message: "Drop an .xlsx or .xlsm workbook file." });
      return;
    }
    handleWorkbookSelected(droppedFile);
  }, [handleWorkbookSelected]);

  const clearPendingInputSync = useCallback(() => {
    if (inputSyncFrameRef.current !== null) {
      window.cancelAnimationFrame(inputSyncFrameRef.current);
      inputSyncFrameRef.current = null;
    }
    for (const timer of inputSyncTimersRef.current) window.clearTimeout(timer);
    inputSyncTimersRef.current = [];
  }, []);

  const syncInputSelection = useCallback((input: HTMLInputElement | null = fileInputRef.current) => {
    const nextFile = input?.files?.item(0) ?? undefined;
    if (!nextFile) return;
    handleWorkbookSelected(nextFile);
  }, [handleWorkbookSelected]);

  const syncInputSelectionSoon = useCallback((input: HTMLInputElement | null = fileInputRef.current) => {
    clearPendingInputSync();
    syncInputSelection(input);
    inputSyncFrameRef.current = window.requestAnimationFrame(() => {
      syncInputSelection(input);
      inputSyncFrameRef.current = null;
    });
    inputSyncTimersRef.current = [100, 300, 1000].map((delay) => window.setTimeout(() => syncInputSelection(input), delay));
  }, [clearPendingInputSync, syncInputSelection]);

  useEffect(() => {
    function handleWindowDragOver(event: globalThis.DragEvent) {
      if (!markWorkbookDropEffect(event.dataTransfer)) return;
      event.preventDefault();
      dragDepthRef.current = Math.max(1, dragDepthRef.current);
      setIsDragging(true);
    }

    function handleWindowDragLeave(event: globalThis.DragEvent) {
      if (event.clientX > 0 && event.clientY > 0 && event.clientX < window.innerWidth && event.clientY < window.innerHeight) return;
      dragDepthRef.current = 0;
      setIsDragging(false);
    }

    function handleWindowDrop(event: globalThis.DragEvent) {
      const transfer = event.dataTransfer;
      if (!hasTransferredFiles(transfer)) return;
      event.preventDefault();
      dragDepthRef.current = 0;
      setIsDragging(false);
      handleDroppedWorkbook(transfer);
    }

    function handleWindowFocus() {
      syncInputSelectionSoon();
    }

    function handleVisibilityChange() {
      if (document.visibilityState === "visible") syncInputSelectionSoon();
    }

    function handlePageShow() {
      syncInputSelectionSoon();
    }

    syncInputSelectionSoon();

    window.addEventListener("dragover", handleWindowDragOver);
    window.addEventListener("dragleave", handleWindowDragLeave);
    window.addEventListener("drop", handleWindowDrop);
    window.addEventListener("focus", handleWindowFocus);
    window.addEventListener("pageshow", handlePageShow);
    document.addEventListener("visibilitychange", handleVisibilityChange);
    return () => {
      window.removeEventListener("dragover", handleWindowDragOver);
      window.removeEventListener("dragleave", handleWindowDragLeave);
      window.removeEventListener("drop", handleWindowDrop);
      window.removeEventListener("focus", handleWindowFocus);
      window.removeEventListener("pageshow", handlePageShow);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
      clearPendingInputSync();
    };
  }, [clearPendingInputSync, handleDroppedWorkbook, syncInputSelectionSoon]);

  useEffect(() => {
    const input = fileInputRef.current;
    if (!input) return;

    const handleNativeFileSelection = () => {
      syncInputSelection(input);
    };

    input.addEventListener("change", handleNativeFileSelection);
    input.addEventListener("input", handleNativeFileSelection);
    input.addEventListener("cancel", handleNativeFileSelection);
    syncInputSelection(input);
    return () => {
      input.removeEventListener("change", handleNativeFileSelection);
      input.removeEventListener("input", handleNativeFileSelection);
      input.removeEventListener("cancel", handleNativeFileSelection);
    };
  }, [syncInputSelection]);

  function handleDrag(event: DragEvent<HTMLElement>) {
    event.preventDefault();
    event.stopPropagation();
    markWorkbookDropEffect(event.dataTransfer);
    if (event.type === "dragenter") {
      dragDepthRef.current += 1;
      setIsDragging(true);
      return;
    }
    if (event.type === "dragover") {
      setIsDragging(true);
      return;
    }
    if (event.type === "dragleave") {
      const nextTarget = event.relatedTarget;
      if (nextTarget instanceof Node && event.currentTarget.contains(nextTarget)) return;
      dragDepthRef.current = Math.max(0, dragDepthRef.current - 1);
      if (dragDepthRef.current === 0) setIsDragging(false);
    }
  }

  function handleDrop(event: DragEvent<HTMLElement>) {
    event.preventDefault();
    event.stopPropagation();
    markWorkbookDropEffect(event.dataTransfer);
    dragDepthRef.current = 0;
    setIsDragging(false);
    handleDroppedWorkbook(event.dataTransfer);
  }

  function pickInputFile(input: HTMLInputElement) {
    handleWorkbookSelected(input.files?.item(0) ?? undefined);
  }

  function handleFileSelect(event: ChangeEvent<HTMLInputElement>) {
    pickInputFile(event.currentTarget);
  }

  function handleFileInput(event: FormEvent<HTMLInputElement>) {
    pickInputFile(event.currentTarget);
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    const nativeFormData = new FormData(form);
    const query = String(nativeFormData.get("ticker") ?? ticker).trim();
    const nativeFile = nativeFormData.get("file");
    const selectedFile =
      selectedFileRef.current
        ?? file
        ?? (nativeFile instanceof File && nativeFile.size > 0 ? nativeFile : null)
        ?? fileInputRef.current?.files?.item(0)
        ?? null;
    if (!query) {
      setError({ message: "Enter a ticker or company name before filling." });
      return;
    }
    if (!selectedFile) {
      setError({ message: "Choose an .xlsx or .xlsm workbook before filling." });
      return;
    }
    if (!isSupportedWorkbookFile(selectedFile)) {
      setError({ message: `${selectedFile.name} is not a supported .xlsx or .xlsm workbook.` });
      return;
    }

    setIsSubmitting(true);
    setError(null);
    setSummary(null);

    const formData = new FormData();
    formData.append("ticker", query);
    formData.append("file", selectedFile);

    try {
      const response = await fetch("/api/fill-model", {
        method: "POST",
        body: formData
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => null);
        setError({
          message: payload?.error ?? "The workbook could not be filled.",
          debugLogPath: payload?.debugLogPath ?? response.headers.get("x-debug-log-path") ?? undefined
        });
        return;
      }

      const encoded = response.headers.get("x-fill-summary");
      if (encoded) {
        const parsedSummary = JSON.parse(decodeURIComponent(encoded)) as FillSummary;
        const debugLogPath = response.headers.get("x-debug-log-path") ?? parsedSummary.debugLogPath;
        setSummary(debugLogPath ? { ...parsedSummary, debugLogPath } : parsedSummary);
      }

      const blob = await response.blob();
      const downloadUrl = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = downloadUrl;
      a.download = response.headers.get("x-output-filename") ?? `${query.toUpperCase()}_historicals_filled.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(downloadUrl);
    } catch (caught) {
      setError({ message: caught instanceof Error ? caught.message : "Something went wrong." });
    } finally {
      setIsSubmitting(false);
    }
  }

  return (
    <main className="shell">
      <div className="topbar" aria-hidden="true">
        <div className="brandMark">
          <span />
          <span />
          <span />
          <span />
        </div>
      </div>
      <section className="workspace">
        <div className="intro">
          <p className="eyebrow">EDGAR to model</p>
          <h1>Historicals Solver</h1>
          <p>
            Drop a valuation template, enter a ticker or company name, and download a workbook with historical income
            statement and balance sheet cells populated from SEC company facts.
          </p>
          <div className="assurance">
            <span>
              <ShieldCheck aria-hidden="true" size={16} />
              SEC EDGAR sourced
            </span>
            <span>
              <CheckCircle2 aria-hidden="true" size={16} />
              Audit notes included
            </span>
          </div>
        </div>

        <form className="tool" action="/api/fill-model" method="post" encType="multipart/form-data" onSubmit={handleSubmit}>
          <div className="toolHeader">
            <FileSpreadsheet aria-hidden="true" size={22} />
            <div>
              <strong>Fill model historicals</strong>
              <span>Upload an Excel valuation template and choose the company.</span>
            </div>
          </div>

          <div className="field">
            <label htmlFor="ticker">Ticker or company name</label>
            <div className="searchBox">
              <Search aria-hidden="true" size={20} />
              <input
                id="ticker"
                name="ticker"
                value={ticker}
                onChange={(event) => setTicker(event.target.value)}
                placeholder="AAPL, Microsoft, Costco..."
                autoComplete="off"
              />
            </div>
          </div>

          <div
            className={`dropzone${isDragging ? " dragging" : ""}${file ? " hasFile" : ""}`}
            aria-label={file ? `Selected workbook ${file.name}. Choose a different workbook.` : "Choose Excel workbook"}
            aria-disabled={isSubmitting}
            onDragEnterCapture={handleDrag}
            onDragOverCapture={handleDrag}
            onDragLeaveCapture={handleDrag}
            onDropCapture={handleDrop}
          >
            <input
              id="model-template-file"
              ref={fileInputRef}
              className="fileInput"
              type="file"
              name="file"
              accept={SUPPORTED_WORKBOOK_ACCEPT}
              aria-label={file ? `Selected workbook ${file.name}. Choose a different workbook.` : "Choose Excel workbook"}
              disabled={isSubmitting}
              onChangeCapture={handleFileSelect}
              onInputCapture={handleFileInput}
              onChange={handleFileSelect}
              onInput={handleFileInput}
            />
            <span className="dropIcon">
              {file ? <FileCheck2 aria-hidden="true" size={30} /> : <UploadCloud aria-hidden="true" size={30} />}
            </span>
            <span className="dropTitle">{file ? "Workbook selected" : "Drop Excel model here"}</span>
            {file ? (
              <span className="selectedFile" aria-live="polite">
                <FileSpreadsheet aria-hidden="true" size={18} />
                <span>{file.name}</span>
                <small>{formatFileSize(file.size)}</small>
              </span>
            ) : (
              <small>Click to browse or drag in an .xlsx or .xlsm file</small>
            )}
            <span className="browseCue">{file ? "Choose different workbook" : "Choose workbook"}</span>
          </div>

          <button className="primary" type="submit" disabled={!canSubmit}>
            {isSubmitting ? <Loader2 className="spin" size={20} /> : <ArrowDownToLine size={20} />}
            {isSubmitting ? "Filling workbook" : "Fill and download"}
          </button>

          {error ? (
            <div className="errorPanel" role="alert">
              <strong>{error.message}</strong>
              {error.debugLogPath ? (
                <span>
                  Debug log: <code>{error.debugLogPath}</code>
                </span>
              ) : null}
            </div>
          ) : null}
        </form>

        {summary ? (
          <section className="summary" aria-live="polite">
            <div>
              <FileSpreadsheet aria-hidden="true" size={24} />
              <div>
                <strong>
                  {summary.companyName} ({summary.ticker})
                </strong>
                <span>
                  {summary.filledCells} cells filled across {summary.periods.length} periods. {summary.commentsAdded} comments
                  added for mapped or plugged rows.
                </span>
              </div>
            </div>
            {summary.warnings.length ? (
              <ul>
                {summary.warnings.slice(0, 6).map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            ) : null}
            {summary.debugLogPath ? (
              <p className="debugPath">
                Debug log: <code>{summary.debugLogPath}</code>
              </p>
            ) : null}
          </section>
        ) : null}
      </section>
    </main>
  );
}
