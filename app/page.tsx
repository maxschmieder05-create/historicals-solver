"use client";

import { ChangeEvent, DragEvent, FormEvent, KeyboardEvent, useEffect, useMemo, useRef, useState } from "react";
import { ArrowDownToLine, CheckCircle2, FileCheck2, FileSpreadsheet, Loader2, Search, ShieldCheck, UploadCloud } from "lucide-react";

type FillSummary = {
  companyName: string;
  ticker: string;
  periods: string[];
  filledCells: number;
  commentsAdded: number;
  warnings: string[];
};

export default function Home() {
  const [ticker, setTicker] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [summary, setSummary] = useState<FillSummary | null>(null);
  const [error, setError] = useState("");
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const canSubmit = useMemo(() => !isSubmitting, [isSubmitting]);

  useEffect(() => {
    function preventBrowserFileOpen(event: globalThis.DragEvent) {
      if (!event.dataTransfer?.types.includes("Files")) return;
      event.preventDefault();
    }

    window.addEventListener("dragover", preventBrowserFileOpen);
    window.addEventListener("drop", preventBrowserFileOpen);
    return () => {
      window.removeEventListener("dragover", preventBrowserFileOpen);
      window.removeEventListener("drop", preventBrowserFileOpen);
    };
  }, []);

  function syncFileInput(nextFile: File | null) {
    if (!fileInputRef.current) return;
    if (!nextFile) {
      fileInputRef.current.value = "";
      return;
    }
    const dataTransfer = new DataTransfer();
    dataTransfer.items.add(nextFile);
    fileInputRef.current.files = dataTransfer.files;
  }

  function formatFileSize(bytes: number) {
    if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024)).toLocaleString()} KB`;
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
  }

  function pickFile(nextFile?: File, options: { syncInput?: boolean } = {}) {
    setError("");
    setSummary(null);
    if (!nextFile) return;
    if (!nextFile.name.toLowerCase().endsWith(".xlsx")) {
      setFile(null);
      syncFileInput(null);
      setError(`${nextFile.name} is not an .xlsx workbook.`);
      return;
    }
    setFile(nextFile);
    if (options.syncInput) syncFileInput(nextFile);
  }

  function openFilePicker() {
    if (!fileInputRef.current) return;
    fileInputRef.current.value = "";
    fileInputRef.current.click();
  }

  function handleDrag(event: DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    event.stopPropagation();
    if (event.type === "dragenter" || event.type === "dragover") setIsDragging(true);
    if (event.type === "dragleave") {
      const nextTarget = event.relatedTarget;
      if (nextTarget instanceof Node && event.currentTarget.contains(nextTarget)) return;
      setIsDragging(false);
    }
    if (event.type === "drop") setIsDragging(false);
  }

  function handleDrop(event: DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(false);
    pickFile(event.dataTransfer.files?.[0], { syncInput: true });
  }

  function handleFileSelect(event: ChangeEvent<HTMLInputElement>) {
    pickFile(event.currentTarget.files?.item(0) ?? undefined);
  }

  function handleDropzoneKeyDown(event: KeyboardEvent<HTMLLabelElement>) {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    openFilePicker();
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    const nativeFormData = new FormData(form);
    const query = String(nativeFormData.get("ticker") ?? ticker).trim();
    const nativeFile = nativeFormData.get("file");
    const selectedFile =
      nativeFile instanceof File && nativeFile.size > 0
        ? nativeFile
        : file ?? fileInputRef.current?.files?.item(0) ?? null;
    if (!query) {
      setError("Enter a ticker or company name before filling.");
      return;
    }
    if (!selectedFile) {
      setError("Choose an .xlsx workbook before filling.");
      return;
    }
    if (!selectedFile.name.toLowerCase().endsWith(".xlsx")) {
      setError(`${selectedFile.name} is not an .xlsx workbook.`);
      return;
    }

    setIsSubmitting(true);
    setError("");
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
        throw new Error(payload?.error ?? "The workbook could not be filled.");
      }

      const encoded = response.headers.get("x-fill-summary");
      if (encoded) {
        setSummary(JSON.parse(decodeURIComponent(encoded)));
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
      setError(caught instanceof Error ? caught.message : "Something went wrong.");
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

          <label
            className={`dropzone${isDragging ? " dragging" : ""}${file ? " hasFile" : ""}`}
            role="button"
            tabIndex={0}
            aria-label={file ? `Selected workbook ${file.name}. Choose a different workbook.` : "Choose Excel workbook"}
            onKeyDown={handleDropzoneKeyDown}
            onDragEnter={handleDrag}
            onDragOver={handleDrag}
            onDragLeave={handleDrag}
            onDrop={handleDrop}
          >
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
            <span className="browseCue">Choose workbook</span>
            <input
              id="model-template-file"
              ref={fileInputRef}
              className="fileInput"
              type="file"
              name="file"
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              aria-hidden="true"
              tabIndex={-1}
              onChange={handleFileSelect}
            />
          </label>

          <button className="primary" type="submit" disabled={!canSubmit}>
            {isSubmitting ? <Loader2 className="spin" size={20} /> : <ArrowDownToLine size={20} />}
            {isSubmitting ? "Filling workbook" : "Fill and download"}
          </button>

          {error ? <p className="error">{error}</p> : null}
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
                {summary.warnings.slice(0, 4).map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            ) : null}
          </section>
        ) : null}
      </section>
    </main>
  );
}
