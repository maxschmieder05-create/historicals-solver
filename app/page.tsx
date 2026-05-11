"use client";

import { ChangeEvent, DragEvent, FormEvent, useMemo, useState } from "react";
import { ArrowDownToLine, FileSpreadsheet, Loader2, Search, UploadCloud } from "lucide-react";

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

  const canSubmit = useMemo(() => Boolean(file && ticker.trim() && !isSubmitting), [file, ticker, isSubmitting]);

  function pickFile(nextFile?: File) {
    setError("");
    setSummary(null);
    if (!nextFile) return;
    if (!nextFile.name.toLowerCase().endsWith(".xlsx")) {
      setError("Upload an .xlsx workbook.");
      return;
    }
    setFile(nextFile);
  }

  function handleDrop(event: DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    setIsDragging(false);
    pickFile(event.dataTransfer.files?.[0]);
  }

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    pickFile(event.target.files?.[0]);
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!file || !ticker.trim()) return;

    setIsSubmitting(true);
    setError("");
    setSummary(null);

    const formData = new FormData();
    formData.append("ticker", ticker.trim());
    formData.append("file", file);

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
      <section className="workspace">
        <div className="intro">
          <p className="eyebrow">EDGAR to model</p>
          <h1>Historicals Solver</h1>
          <p>
            Drop the Owl Fund model, enter a ticker or company name, and download a workbook with the blue historical
            financial cells populated from SEC company facts.
          </p>
        </div>

        <form className="tool" onSubmit={handleSubmit}>
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

          <label
            className={isDragging ? "dropzone dragging" : "dropzone"}
            onDragOver={(event) => {
              event.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
          >
            <input type="file" accept=".xlsx" onChange={handleFileChange} />
            <UploadCloud aria-hidden="true" size={34} />
            <span>{file ? file.name : "Drop Excel model here"}</span>
            <small>{file ? `${(file.size / 1024 / 1024).toFixed(2)} MB selected` : "Only .xlsx files are accepted"}</small>
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
