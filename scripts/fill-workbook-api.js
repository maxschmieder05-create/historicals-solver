const fsp = require("node:fs/promises");
const { spawn } = require("node:child_process");
const path = require("node:path");

const DEFAULT_TIMEOUT_MS = 15 * 60 * 1000;

async function postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook, timeoutMs = DEFAULT_TIMEOUT_MS }) {
  await fsp.mkdir(path.dirname(outputWorkbook), { recursive: true });
  await runCurl(
    [
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
    timeoutMs
  );
  await new Promise((resolve) => setTimeout(resolve, 1500));
}

function runCurl(args, timeoutMs) {
  return new Promise((resolve, reject) => {
    const child = spawn("curl", args, { stdio: ["ignore", "ignore", "pipe"] });
    const stderr = [];
    const timer = setTimeout(() => {
      child.kill("SIGTERM");
      reject(new Error(`Fill API request timed out after ${Math.round(timeoutMs / 1000)}s`));
    }, timeoutMs + 5_000);

    child.stderr.on("data", (chunk) => stderr.push(chunk));
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

module.exports = { postWorkbook };
