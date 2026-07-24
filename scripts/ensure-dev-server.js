#!/usr/bin/env node

const fs = require("node:fs");
const http = require("node:http");
const path = require("node:path");
const { spawn } = require("node:child_process");

const repoRoot = path.resolve(__dirname, "..");
const runtimeDirectory = path.join(repoRoot, "tmp");
const port = positiveInteger(process.env.PORT, 3000);
const host = process.env.DEV_SERVER_HOST || "127.0.0.1";
const statePath = path.join(runtimeDirectory, `dev-server-${port}.json`);
const logPath = path.join(runtimeDirectory, `dev-server-${port}.log`);
const nextBin = path.join(repoRoot, "node_modules", "next", "dist", "bin", "next");

function positiveInteger(value, fallback) {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : fallback;
}

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

function processIsRunning(pid) {
  if (!Number.isSafeInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return error?.code === "EPERM";
  }
}

function readState() {
  try {
    const parsed = JSON.parse(fs.readFileSync(statePath, "utf8"));
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function writeState(serverPid) {
  fs.mkdirSync(runtimeDirectory, { recursive: true });
  fs.writeFileSync(
    statePath,
    JSON.stringify(
      {
        supervisorPid: process.pid,
        serverPid,
        port,
        host,
        startedAt: new Date().toISOString(),
        logPath
      },
      null,
      2
    )
  );
}

function removeOwnedState() {
  const state = readState();
  if (state.supervisorPid === process.pid) fs.rmSync(statePath, { force: true });
}

function serverResponds(timeoutMs = 2_000) {
  return new Promise((resolve) => {
    const request = http.get({ host, port, path: "/", timeout: timeoutMs }, (response) => {
      response.resume();
      resolve(Boolean(response.statusCode && response.statusCode >= 200 && response.statusCode < 400));
    });
    request.on("timeout", () => request.destroy());
    request.on("error", () => resolve(false));
  });
}

async function waitForServer(timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (await serverResponds()) return true;
    await delay(500);
  }
  return false;
}

function logTail() {
  try {
    return fs.readFileSync(logPath, "utf8").split(/\r?\n/).slice(-24).join("\n").trim();
  } catch {
    return "";
  }
}

async function supervise() {
  fs.mkdirSync(runtimeDirectory, { recursive: true });
  let child = null;
  let stopping = false;

  const stop = (signal) => {
    if (stopping) return;
    stopping = true;
    if (child && processIsRunning(child.pid)) child.kill(signal);
  };

  process.on("SIGINT", () => stop("SIGINT"));
  process.on("SIGTERM", () => stop("SIGTERM"));
  process.on("exit", removeOwnedState);

  while (!stopping) {
    console.log(`[dev-supervisor] Starting Next.js on http://${host}:${port} at ${new Date().toISOString()}`);
    child = spawn(process.execPath, [nextBin, "dev", "--hostname", host, "--port", String(port)], {
      cwd: repoRoot,
      env: process.env,
      stdio: "inherit"
    });
    writeState(child.pid);
    const outcome = await new Promise((resolve) => {
      child.once("exit", (code, signal) => resolve({ code, signal }));
      child.once("error", (error) => resolve({ error }));
    });
    child = null;
    if (stopping) break;
    console.error(`[dev-supervisor] Next.js exited unexpectedly (${JSON.stringify(outcome)}); restarting in 1 second.`);
    await delay(1_000);
  }

  removeOwnedState();
}

async function stopServer() {
  const state = readState();
  if (!processIsRunning(state.supervisorPid)) {
    fs.rmSync(statePath, { force: true });
    console.log(`No managed dev server is running on port ${port}.`);
    return;
  }
  process.kill(state.supervisorPid, "SIGTERM");
  const deadline = Date.now() + 10_000;
  while (Date.now() < deadline && processIsRunning(state.supervisorPid)) await delay(100);
  if (processIsRunning(state.supervisorPid)) throw new Error(`Dev-server supervisor ${state.supervisorPid} did not stop.`);
  fs.rmSync(statePath, { force: true });
  console.log(`Stopped the managed dev server on port ${port}.`);
}

async function ensureServer() {
  if (await serverResponds()) {
    const state = readState();
    const managed = processIsRunning(state.supervisorPid);
    console.log(`Dev server is already responding at http://localhost:${port}${managed ? " (supervised)" : ""}.`);
    return;
  }

  const state = readState();
  if (processIsRunning(state.supervisorPid)) {
    console.log(`Dev-server supervisor ${state.supervisorPid} is running; waiting for port ${port}.`);
    if (await waitForServer(45_000)) {
      console.log(`Dev server is responding at http://localhost:${port}.`);
      return;
    }
    throw new Error(`The dev-server supervisor is running, but port ${port} did not become healthy.\n${logTail()}`);
  }

  fs.mkdirSync(runtimeDirectory, { recursive: true });
  const logDescriptor = fs.openSync(logPath, "a");
  const supervisor = spawn(process.execPath, [__filename, "--supervise"], {
    cwd: repoRoot,
    detached: true,
    env: process.env,
    stdio: ["ignore", logDescriptor, logDescriptor]
  });
  supervisor.unref();
  fs.closeSync(logDescriptor);

  if (!(await waitForServer(45_000))) {
    throw new Error(`Dev server did not become healthy on port ${port}.\n${logTail()}`);
  }
  console.log(`Started supervised dev server ${supervisor.pid} at http://localhost:${port}.`);
  console.log(`Log: ${logPath}`);
}

async function main() {
  if (process.argv.includes("--supervise")) return supervise();
  if (process.argv.includes("--stop")) return stopServer();
  return ensureServer();
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  });
}

module.exports = { positiveInteger, processIsRunning, serverResponds };
