const fsp = require("node:fs/promises");
const { spawn } = require("node:child_process");
const path = require("node:path");

const DEFAULT_TIMEOUT_MS = Number(process.env.FILL_API_TIMEOUT_MS || 30 * 60 * 1000);

async function postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook, timeoutMs = DEFAULT_TIMEOUT_MS }) {
  await fsp.mkdir(path.dirname(outputWorkbook), { recursive: true });
  const { args, stdinText } = buildCurlInvocation({
    apiUrl,
    ticker,
    inputWorkbook,
    outputWorkbook,
    timeoutMs,
    accessKey: deploymentAccessKey()
  });
  await runCurl(args, timeoutMs, stdinText);
  await new Promise((resolve) => setTimeout(resolve, 1500));
}

function buildCurlInvocation({ apiUrl, ticker, inputWorkbook, outputWorkbook, timeoutMs = DEFAULT_TIMEOUT_MS, accessKey = "" }) {
  const normalizedAccessKey = String(accessKey).trim();
  if (/[\r\n\0]/.test(normalizedAccessKey)) {
    throw new Error("Fill API access keys cannot contain line breaks or null bytes.");
  }
  const stdinText = normalizedAccessKey
    ? `header = "Authorization: Bearer ${escapeCurlConfigValue(normalizedAccessKey)}"\n`
    : "";
  return {
    args: [
      ...(stdinText ? ["--config", "-"] : []),
      "--silent",
      "--show-error",
      "--fail-with-body",
      "--max-time",
      String(Math.ceil(timeoutMs / 1000)),
      "--output",
      outputWorkbook,
      "--form",
      `ticker=${ticker}`,
      "--form",
      `file=@${inputWorkbook}`,
      apiUrl
    ],
    stdinText
  };
}

function deploymentAccessKey() {
  return process.env.FILL_API_KEY?.trim() || process.env.HISTORICALS_API_KEY?.trim() || "";
}

function escapeCurlConfigValue(value) {
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function runCurl(args, timeoutMs, stdinText = "") {
  return new Promise((resolve, reject) => {
    const child = spawn("curl", args, { stdio: [stdinText ? "pipe" : "ignore", "ignore", "pipe"] });
    const stderr = [];
    const timer = setTimeout(() => {
      child.kill("SIGTERM");
      reject(new Error(`Fill API request timed out after ${Math.round(timeoutMs / 1000)}s`));
    }, timeoutMs + 5_000);

    child.stderr.on("data", (chunk) => stderr.push(chunk));
    if (stdinText && child.stdin) {
      child.stdin.on("error", () => undefined);
      child.stdin.end(stdinText);
    }
    child.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(new Error(Buffer.concat(stderr).toString("utf8").trim() || `curl exited with code ${code}`));
    });
  });
}

module.exports = {
  postWorkbook,
  __fillWorkbookApiTestHooks: {
    buildCurlInvocation,
    deploymentAccessKey
  }
};
