"use client";

import { ChangeEvent, DragEvent, FormEvent, KeyboardEvent, useEffect, useMemo, useRef, useState } from "react";
import { ArrowDownToLine, CheckCircle2, FileSpreadsheet, Loader2, Search, ShieldCheck, UploadCloud } from "lucide-react";

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

  function pickFile(nextFile?: File) {
    setError("");
    setSummary(null);
    if (!nextFile) return;
    if (!nextFile.name.toLowerCase().endsWith(".xlsx")) {
      setFile(null);
      setError(`${nextFile.name} is not an .xlsx workbook.`);
      return;
    }
    setFile(nextFile);
  }

  function handleDrag(event: DragEvent<HTMLDivElement>) {
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

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(false);
    pickFile(event.dataTransfer.files?.[0]);
  }

  function handleFileSelect(event: ChangeEvent<HTMLInputElement> | FormEvent<HTMLInputElement>) {
    pickFile(event.currentTarget.files?.item(0) ?? undefined);
  }

  function handleDropzoneKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    fileInputRef.current?.click();
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const selectedFile = file ?? fileInputRef.current?.files?.item(0) ?? null;
    if (!ticker.trim()) {
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
    formData.append("ticker", ticker.trim());
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
      a.download = response.headers.get("x-output-filename") ?? `${ticker.toUpperCase()}_historicals_filled.xlsx`;
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

        <form className="tool" onSubmit={handleSubmit}>
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
                value={ticker}
                onChange={(event) => setTicker(event.target.value)}
                placeholder="AAPL, Microsoft, Costco..."
                autoComplete="off"
              />
            </div>
          </div>

          <div
            className={isDragging ? "dropzone dragging" : "dropzone"}
            role="button"
            tabIndex={0}
            onClick={(event) => {
              if (event.target instanceof HTMLInputElement) return;
              fileInputRef.current?.click();
            }}
            onKeyDown={handleDropzoneKeyDown}
            onDragEnter={handleDrag}
            onDragOver={handleDrag}
            onDragLeave={handleDrag}
            onDrop={handleDrop}
          >
            <input
              id="model-template-file"
              ref={fileInputRef}
              className="filePicker"
              type="file"
              name="model-template-file"
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              aria-label="Upload Excel model template"
              onClick={(event) => {
                event.currentTarget.value = "";
              }}
              onInput={handleFileSelect}
              onChange={handleFileSelect}
            />
            <span className="dropIcon">
              <UploadCloud aria-hidden="true" size={30} />
            </span>
            <span className="dropTitle">{file ? file.name : "Drop Excel model here"}</span>
            <small>{file ? `${(file.size / 1024 / 1024).toFixed(2)} MB selected` : "Click to browse or drag in an .xlsx file"}</small>
          </div>

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
