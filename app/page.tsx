"use client";

import {
  ChangeEvent,
  DragEvent,
  FormEvent,
  MouseEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState
} from "react";
import { ArrowDownToLine, CheckCircle2, FileCheck2, FileSpreadsheet, KeyRound, Loader2, Search, ShieldCheck, UploadCloud, X } from "lucide-react";

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

const SUPPORTED_WORKBOOK_EXTENSIONS = [".xlsx"] as const;
const SUPPORTED_WORKBOOK_ACCEPT = [
  ".xlsx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
  const [accessKey, setAccessKey] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [summary, setSummary] = useState<FillSummary | null>(null);
  const [error, setError] = useState<FillError | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const selectedFileRef = useRef<File | null>(null);
  const dragDepthRef = useRef(0);
  const pickerSyncTimersRef = useRef<number[]>([]);
  const activeRequestRef = useRef<AbortController | null>(null);

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
      setError({ message: `${nextFile.name} is not a supported .xlsx workbook.` });
      return;
    }
    selectedFileRef.current = nextFile;
    setFile(nextFile);
  }, [clearFileInput]);

  const handleDroppedWorkbook = useCallback((dataTransfer: DataTransfer | null) => {
    const droppedFile = workbookFileFromTransfer(dataTransfer);
    if (!droppedFile) {
      setError({ message: "Drop an .xlsx workbook file." });
      return;
    }
    handleWorkbookSelected(droppedFile);
  }, [handleWorkbookSelected]);

  const clearPickerSyncTimers = useCallback(() => {
    pickerSyncTimersRef.current.forEach((timer) => window.clearTimeout(timer));
    pickerSyncTimersRef.current = [];
  }, []);

  const syncFileInputSelection = useCallback((input: HTMLInputElement | null = fileInputRef.current) => {
    const selected = input?.files?.item(0);
    if (selected) handleWorkbookSelected(selected);
  }, [handleWorkbookSelected]);

  const scheduleFileInputSelectionSync = useCallback((input: HTMLInputElement | null = fileInputRef.current) => {
    clearPickerSyncTimers();
    pickerSyncTimersRef.current = [0, 100, 400].map((delay) =>
      window.setTimeout(() => syncFileInputSelection(input), delay)
    );
  }, [clearPickerSyncTimers, syncFileInputSelection]);

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

    window.addEventListener("dragover", handleWindowDragOver);
    window.addEventListener("dragleave", handleWindowDragLeave);
    window.addEventListener("drop", handleWindowDrop);
    return () => {
      window.removeEventListener("dragover", handleWindowDragOver);
      window.removeEventListener("dragleave", handleWindowDragLeave);
      window.removeEventListener("drop", handleWindowDrop);
    };
  }, [handleDroppedWorkbook]);

  useEffect(() => {
    const input = fileInputRef.current;
    if (!input) return;
    const handleNativeSelection = () => scheduleFileInputSelectionSync(input);
    const handlePickerReturn = () => scheduleFileInputSelectionSync(input);
    const handleVisibilityChange = () => {
      if (document.visibilityState === "visible") handlePickerReturn();
    };

    input.addEventListener("change", handleNativeSelection);
    input.addEventListener("input", handleNativeSelection);
    window.addEventListener("focus", handlePickerReturn);
    document.addEventListener("visibilitychange", handleVisibilityChange);
    return () => {
      input.removeEventListener("change", handleNativeSelection);
      input.removeEventListener("input", handleNativeSelection);
      window.removeEventListener("focus", handlePickerReturn);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
      clearPickerSyncTimers();
    };
  }, [clearPickerSyncTimers, scheduleFileInputSelectionSync]);

  useEffect(() => () => activeRequestRef.current?.abort(), []);

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

  function handleFileSelect(event: ChangeEvent<HTMLInputElement>) {
    handleWorkbookSelected(event.currentTarget.files?.item(0) ?? undefined);
  }

  function prepareFilePicker(event: MouseEvent<HTMLInputElement>) {
    event.currentTarget.value = "";
  }

  function cancelFill() {
    if (!activeRequestRef.current) return;
    activeRequestRef.current.abort();
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
      setError({ message: "Choose an .xlsx workbook before filling." });
      return;
    }
    if (!isSupportedWorkbookFile(selectedFile)) {
      setError({ message: `${selectedFile.name} is not a supported .xlsx workbook.` });
      return;
    }

    setIsSubmitting(true);
    setError(null);
    setSummary(null);

    const formData = new FormData();
    formData.append("ticker", query);
    formData.append("file", selectedFile);
    const controller = new AbortController();
    activeRequestRef.current = controller;
    const normalizedAccessKey = accessKey.trim();

    try {
      const response = await fetch("/api/fill-model", {
        method: "POST",
        body: formData,
        signal: controller.signal,
        headers: normalizedAccessKey ? { Authorization: `Bearer ${normalizedAccessKey}` } : undefined
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
      setError({
        message: controller.signal.aborted
          ? "Workbook fill cancelled. No output was downloaded."
          : caught instanceof Error
            ? caught.message
            : "Something went wrong."
      });
    } finally {
      if (activeRequestRef.current === controller) activeRequestRef.current = null;
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
            statement and balance sheet cells populated from SEC company facts for domestic 10-K/10-Q filers.
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
                disabled={isSubmitting}
              />
            </div>
          </div>

          <div className="field">
            <label htmlFor="deployment-access-key">Deployment access key <span className="optionalLabel">(if required)</span></label>
            <div className="searchBox">
              <KeyRound aria-hidden="true" size={20} />
              <input
                id="deployment-access-key"
                type="password"
                value={accessKey}
                onChange={(event) => setAccessKey(event.target.value)}
                placeholder="Enter the key provided by your administrator"
                autoComplete="off"
                spellCheck={false}
                disabled={isSubmitting}
              />
            </div>
            <small className="fieldHint">
              Leave blank unless this deployment is protected. The key is kept only in this page&apos;s memory, sent in the
              authorization header for this fill, and never added to the workbook.
            </small>
          </div>

          <div
            className={`dropzone${isDragging ? " dragging" : ""}${file ? " hasFile" : ""}`}
            data-testid="workbook-dropzone"
            onDragEnter={handleDrag}
            onDragOver={handleDrag}
            onDragLeave={handleDrag}
            onDrop={handleDrop}
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
              onClick={prepareFilePicker}
              onChange={handleFileSelect}
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
              <small>Click to browse or drag in an .xlsx file</small>
            )}
            <span className="browseCue" aria-hidden="true">
              {file ? "Choose different workbook" : "Choose workbook"}
            </span>
          </div>

          <div className="formActions">
            <button className="primary" type="submit" disabled={!canSubmit}>
              {isSubmitting ? <Loader2 className="spin" size={20} /> : <ArrowDownToLine size={20} />}
              {isSubmitting ? "Filling workbook" : "Fill and download"}
            </button>
            {isSubmitting ? (
              <button className="cancel" type="button" onClick={cancelFill}>
                <X aria-hidden="true" size={19} />
                Cancel fill
              </button>
            ) : null}
          </div>

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
                  added for mapped or derived rows.
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
