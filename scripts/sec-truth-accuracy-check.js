#!/usr/bin/env node
// SEC-truth accuracy harness.
//
// For each ticker: resolve its CIK from SEC, fetch the actual reported annual
// figures from EDGAR (companyconcept), POST an input template to the running
// fill-model server, then score the filled Model sheet's headline lines against
// the SEC ground truth. Unlike the *-regression checks, this measures ACCURACY
// against the source of truth, not reproduction of a previously saved output.
//
// Usage:
//   FILL_API_URL=http://localhost:3004/api/fill-model \
//   node scripts/sec-truth-accuracy-check.js RSG AMZN NFLX THC
//
// Input template per ticker defaults to the matching verified example workbook
// (github/example-excels/verification/<TICKER>_historicals_filled.xlsx); override
// with TEMPLATE_<TICKER>=/path/to/template.xlsx or TEMPLATE_DEFAULT=...
const fs = require("fs");
const path = require("path");
const http = require("http");
const https = require("https");
const { URL } = require("url");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");
const API = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const TICKERS = (process.argv.slice(2).join(",") || process.env.ACCURACY_TICKERS || "RSG,AMZN,NFLX,THC")
  .split(",").map((t) => t.trim().toUpperCase()).filter(Boolean);
const TOL_PCT = Number(process.env.ACCURACY_TOLERANCE_PCT || 2);

function secUserAgent() {
  try {
    const env = fs.readFileSync(path.join(repoRoot, ".env.local"), "utf8");
    const m = env.match(/^SEC_USER_AGENT="?([^"\n]+)"?/m);
    if (m) return m[1];
  } catch {}
  return process.env.SEC_USER_AGENT || "HistoricalsSolver accuracy-check contact@example.com";
}
const SEC_HEADERS = { "User-Agent": secUserAgent(), "Accept-Encoding": "gzip, deflate" };

async function secJson(url) {
  const res = await fetch(url, { headers: SEC_HEADERS });
  if (!res.ok) return null;
  return res.json();
}

let tickerMapPromise = null;
async function resolveCik(ticker) {
  if (!tickerMapPromise) tickerMapPromise = secJson("https://www.sec.gov/files/company_tickers.json");
  const map = await tickerMapPromise;
  if (!map) return null;
  for (const row of Object.values(map)) {
    if (String(row.ticker).toUpperCase() === ticker) return String(row.cik_str).padStart(10, "0");
  }
  return null;
}

async function concept(cik, tag) {
  return secJson(`https://data.sec.gov/api/xbrl/companyconcept/CIK${cik}/us-gaap/${tag}.json`);
}
function annualFY(json, { instant = false } = {}) {
  const out = {};
  if (!json || !json.units) return out;
  const series = json.units.USD || json.units["USD/shares"] || Object.values(json.units)[0] || [];
  for (const f of series) {
    if (!f.form || !f.form.startsWith("10-K") || f.fp !== "FY") continue;
    if (instant) {
      if (f.end && !f.start) out[f.fy] = f.val;
    } else if (f.start && f.end) {
      const days = (new Date(f.end) - new Date(f.start)) / 86400000;
      if (days > 300) out[f.fy] = f.val;
    }
  }
  return out;
}

async function secTruth(cik) {
  const [rev1, rev2, ni, assets, eps] = await Promise.all([
    concept(cik, "RevenueFromContractWithCustomerExcludingAssessedTax"),
    concept(cik, "Revenues"),
    concept(cik, "NetIncomeLoss"),
    concept(cik, "Assets"),
    concept(cik, "EarningsPerShareDiluted"),
  ]);
  const rev = { ...annualFY(rev2), ...annualFY(rev1) };
  const niA = annualFY(ni);
  const asA = annualFY(assets, { instant: true });
  const epsA = annualFY(eps);
  const truth = {};
  for (const y of new Set([...Object.keys(rev), ...Object.keys(niA)])) {
    truth[y] = {
      rev: rev[y] != null ? rev[y] / 1e6 : null,
      ni: niA[y] != null ? niA[y] / 1e6 : null,
      assets: asA[y] != null ? asA[y] / 1e6 : null,
      eps: epsA[y] != null ? epsA[y] : null,
    };
  }
  return truth;
}

const num = (c) => {
  const v = c && c.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
};
const text = (c) => {
  const v = c && c.value;
  if (v == null) return "";
  if (typeof v === "object" && v.richText) return v.richText.map((r) => r.text).join("");
  if (typeof v === "object" && "result" in v) return String(v.result);
  return String(v);
};

function headline(ws) {
  if (!ws) return {};
  const fyCols = {};
  for (let c = 4; c <= 40; c++) {
    const a = num(ws.getCell(24, c)), b = num(ws.getCell(25, c));
    if (a && b && a === b && a >= 2015 && a <= 2035) fyCols[a] = c;
  }
  const find = (re) => {
    for (let r = 1; r <= ws.rowCount; r++) if (re.test(text(ws.getCell(r, 3)).trim())) return r;
    return null;
  };
  const rows = {
    rev: find(/^Revenue$/i),
    ni: find(/^Net Income \(Loss\)$/i),
    eps: find(/^GAAP EPS \(Diluted\)$/i),
    assets: find(/^Total Assets$/i),
  };
  const out = {};
  for (const [y, col] of Object.entries(fyCols)) {
    out[y] = {
      rev: rows.rev ? num(ws.getCell(rows.rev, col)) : null,
      ni: rows.ni ? num(ws.getCell(rows.ni, col)) : null,
      eps: rows.eps ? num(ws.getCell(rows.eps, col)) : null,
      assets: rows.assets ? num(ws.getCell(rows.assets, col)) : null,
    };
  }
  return out;
}

function templatePath(ticker) {
  return process.env[`TEMPLATE_${ticker}`] || process.env.TEMPLATE_DEFAULT ||
    path.join(repoRoot, "github", "example-excels", "verification", `${ticker}_historicals_filled.xlsx`);
}

// Use node:http (not global fetch) so there is no 300s headers timeout — fills can
// legitimately run for minutes. Multipart body is assembled by hand.
function postWorkbook(ticker, tmpl) {
  const buf = fs.readFileSync(tmpl);
  const boundary = "----hsAccuracyBoundary" + ticker;
  const head = Buffer.from(
    `--${boundary}\r\nContent-Disposition: form-data; name="ticker"\r\n\r\n${ticker}\r\n` +
    `--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="${path.basename(tmpl)}"\r\n` +
    `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`
  );
  const tail = Buffer.from(`\r\n--${boundary}--\r\n`);
  const body = Buffer.concat([head, buf, tail]);
  const u = new URL(API);
  const lib = u.protocol === "https:" ? https : http;
  const t0 = Date.now();
  return new Promise((resolve) => {
    const req = lib.request(
      { hostname: u.hostname, port: u.port, path: u.pathname, method: "POST",
        headers: { "Content-Type": `multipart/form-data; boundary=${boundary}`, "Content-Length": body.length } },
      (res) => {
        const chunks = [];
        res.on("data", (d) => chunks.push(d));
        res.on("end", () => {
          const secs = (Date.now() - t0) / 1000;
          const payload = Buffer.concat(chunks);
          if (res.statusCode !== 200) {
            let msg = payload.toString("utf8").slice(0, 400);
            try { msg = JSON.parse(payload.toString("utf8")).error || msg; } catch {}
            return resolve({ ok: false, status: res.statusCode, secs, error: msg });
          }
          const summaryHeader = res.headers["x-fill-summary"];
          let summary = null;
          try { summary = summaryHeader ? JSON.parse(decodeURIComponent(summaryHeader)) : null; } catch {}
          resolve({ ok: true, secs, buffer: payload, summary });
        });
      }
    );
    req.setTimeout(0); // no inactivity timeout
    req.on("error", (e) => resolve({ ok: false, status: 0, secs: (Date.now() - t0) / 1000, error: e.message }));
    req.end(body);
  });
}

function score(label, wb, sec) {
  if (sec == null) return null;
  if (wb == null) return { label, ok: false, blank: true, msg: `${label}: (blank) vs SEC ${fmt(sec)}` };
  const pct = sec ? ((wb - sec) / sec) * 100 : (wb === 0 ? 0 : Infinity);
  const ok = Math.abs(pct) <= TOL_PCT;
  const flag = ok ? "OK " : (Math.abs(pct) <= 10 ? "WARN" : "FAIL");
  return { label, ok, pct, msg: `[${flag}] ${label}: ${fmt(wb)} vs SEC ${fmt(sec)} (${pct >= 0 ? "+" : ""}${pct.toFixed(1)}%)` };
}
const fmt = (n) => (typeof n === "number" ? (Math.abs(n) >= 100 ? n.toFixed(0) : n.toFixed(2)) : String(n));

(async () => {
  console.log(`SEC-truth accuracy check  api=${API}  tolerance=±${TOL_PCT}%  tickers=${TICKERS.join(",")}\n`);
  const totals = { pass: 0, total: 0, failedTickers: [] };
  for (const ticker of TICKERS) {
    console.log(`══════════════════ ${ticker} ══════════════════`);
    const cik = await resolveCik(ticker);
    if (!cik) { console.log("  could not resolve CIK; skipping\n"); continue; }
    const truth = await secTruth(cik);
    const tmpl = templatePath(ticker);
    if (!fs.existsSync(tmpl)) { console.log(`  no template at ${tmpl}; skipping\n`); continue; }

    const r = await postWorkbook(ticker, tmpl);
    if (!r.ok) {
      totals.failedTickers.push(ticker);
      console.log(`  FILL FAILED (HTTP ${r.status}) in ${r.secs.toFixed(1)}s`);
      console.log(`    ${String(r.error).split(". Automatic")[0]}\n`);
      continue;
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(r.buffer);
    try { fs.mkdirSync(path.join(repoRoot, "tmp", "accuracy-out"), { recursive: true }); fs.writeFileSync(path.join(repoRoot, "tmp", "accuracy-out", `${ticker}_fresh.xlsx`), r.buffer); } catch {}
    const model = headline(wb.getWorksheet("Model"));
    const s = r.summary || {};
    console.log(`  FILL OK in ${r.secs.toFixed(1)}s  filledCells=${s.filledCells ?? "?"} warnings=${(s.warnings || []).length}`);

    // NOTE: the "GAAP EPS (Diluted)" row is formula-driven; ExcelJS reads stale cached
    // results (often template-leftover 2.89/5.34/7.78), so EPS is reported FYI-only and not
    // scored. Revenue / Net Income / Total Assets are hardcoded values and are scored.
    const years = Object.keys(truth).map(Number).filter((y) => model[y] && (truth[y].rev != null || truth[y].ni != null)).sort();
    for (const y of years.slice(-3)) {
      const lines = [
        score("Revenue ", model[y].rev, truth[y].rev),
        score("NetInc  ", model[y].ni, truth[y].ni),
        score("Assets  ", model[y].assets, truth[y].assets),
      ].filter(Boolean);
      console.log(`  FY${y}:`);
      for (const l of lines) { console.log("    " + l.msg); totals.total++; if (l.ok) totals.pass++; }
      if (model[y].eps != null) console.log(`    DilEPS row=${fmt(model[y].eps)} vs SEC ${truth[y].eps} — FYI only (formula/uncomputed, not scored)`);
    }
    console.log("");
  }
  const pctPass = totals.total ? ((totals.pass / totals.total) * 100).toFixed(0) : "—";
  console.log(`════ SUMMARY: ${totals.pass}/${totals.total} headline cells within ±${TOL_PCT}% (${pctPass}%)` +
    (totals.failedTickers.length ? `; FILL FAILED: ${totals.failedTickers.join(", ")}` : "") + " ════");
  process.exit(totals.failedTickers.length || (totals.total && totals.pass < totals.total) ? 1 : 0);
})().catch((e) => { console.error("FATAL", e); process.exit(2); });
