import { NextRequest, NextResponse } from "next/server";
import { randomUUID, timingSafeEqual } from "node:crypto";
import { mkdir, open, readFile, stat, unlink, type FileHandle } from "node:fs/promises";
import { hostname, tmpdir } from "node:os";
import path from "node:path";
import { fillModelErrorDetails, fillModelWorkbook } from "../../../server/fill-model/fill-model-service";

export const runtime = "nodejs";
export const maxDuration = 900;

const MAX_UPLOAD_BYTES = positiveInteger(process.env.FILL_MODEL_MAX_UPLOAD_BYTES, 30 * 1024 * 1024);
const MAX_MULTIPART_OVERHEAD_BYTES = positiveInteger(process.env.FILL_MODEL_MAX_MULTIPART_OVERHEAD_BYTES, 1024 * 1024);
const MAX_REQUEST_BYTES = MAX_UPLOAD_BYTES + MAX_MULTIPART_OVERHEAD_BYTES;
const UPLOAD_READ_TIMEOUT_MS = positiveInteger(process.env.FILL_MODEL_UPLOAD_READ_TIMEOUT_MS, 120_000);
const MAX_CONCURRENT_FILLS = positiveInteger(process.env.FILL_MODEL_MAX_CONCURRENT_REQUESTS, 1);
let activeFillRequests = 0;

type FillRequestLock = { handle: FileHandle; path: string; token: string };

type FillRequestLockMetadata = {
  pid?: number;
  startedAt?: string;
  token?: string;
  hostname?: string;
};

class FillApiRequestError extends Error {
  readonly status: number;

  constructor(message: string, status: number) {
    super(message);
    this.name = "FillApiRequestError";
    this.status = status;
  }
}

export async function POST(request: NextRequest) {
  if (request.signal.aborted) return cancelledResponse();
  const configuredApiKey = process.env.HISTORICALS_API_KEY?.trim();
  if (!configuredApiKey && process.env.NODE_ENV === "production" && !allowsUnauthenticatedFill()) {
    return authenticationConfigurationResponse();
  }
  if (configuredApiKey && !hasValidBearerToken(request.headers.get("authorization"), configuredApiKey)) return unauthorizedResponse();
  const contentLengthHeader = request.headers.get("content-length");
  const contentLength = contentLengthHeader === null ? null : Number(contentLengthHeader);
  if (contentLength !== null && (!Number.isFinite(contentLength) || contentLength < 0)) {
    return jsonError("Invalid Content-Length header.", 400);
  }
  if (contentLength !== null && contentLength > MAX_REQUEST_BYTES) return oversizedUploadResponse();
  if (activeFillRequests >= MAX_CONCURRENT_FILLS) {
    return busyResponse();
  }

  let processLock: FillRequestLock | null = null;
  let fillSlotAcquired = false;
  try {
    processLock = MAX_CONCURRENT_FILLS === 1 ? await acquireFillRequestLock() : null;
    if (MAX_CONCURRENT_FILLS === 1 && !processLock) return busyResponse();
    activeFillRequests += 1;
    fillSlotAcquired = true;

    const formData = await readMultipartFormDataWithinLimit(request);
    const query = String(formData.get("ticker") ?? "").trim();
    const file = formData.get("file");

    if (!query) return jsonError("Enter a ticker or company name.", 400);
    if (!(file instanceof File)) return jsonError("Upload an .xlsx workbook.", 400);
    if (!/\.xlsx$/i.test(file.name)) {
      return jsonError("Only .xlsx workbooks are supported. Macro-enabled .xlsm files are rejected because VBA cannot yet be preserved safely.", 400);
    }
    if (file.size <= 0) return jsonError("The uploaded workbook is empty.", 400);
    if (file.size > MAX_UPLOAD_BYTES) return oversizedUploadResponse();

    const result = await fillModelWorkbook({
      query,
      workbookBuffer: await file.arrayBuffer(),
      workbookName: file.name,
      signal: request.signal
    });
    if (request.signal.aborted) throw cancelledRequestError();
    const responseBody = result.output.buffer.slice(result.output.byteOffset, result.output.byteOffset + result.output.byteLength) as ArrayBuffer;
    const exposeDebugPath = shouldExposeDebugPath();
    const responseSummary = browserFillSummary(result.summary, exposeDebugPath);

    return new NextResponse(responseBody, {
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": `attachment; filename="${result.outputName}"`,
        "x-output-filename": result.outputName,
        "x-fill-summary": encodeURIComponent(JSON.stringify(responseSummary)),
        "cache-control": "no-store",
        ...(exposeDebugPath && result.summary.debugLogPath ? { "x-debug-log-path": result.summary.debugLogPath } : {})
      }
    });
  } catch (error) {
    if (error instanceof FillApiRequestError) return jsonError(error.message, error.status);
    if (request.signal.aborted) return cancelledResponse();
    console.error(error);
    const { message, status, debugLogPath } = fillModelErrorDetails(error);
    return jsonError(message, status, shouldExposeDebugPath() ? debugLogPath : undefined);
  } finally {
    if (fillSlotAcquired) activeFillRequests = Math.max(0, activeFillRequests - 1);
    await releaseFillRequestLock(processLock);
  }
}

function browserFillSummary(summary: Record<string, unknown>, exposeDebugPath: boolean) {
  const warnings = Array.isArray(summary.warnings)
    ? summary.warnings.slice(0, 12).map((warning) => String(warning).replace(/[\r\n]+/g, " ").slice(0, 240))
    : [];
  return {
    companyName: summary.companyName,
    ticker: summary.ticker,
    periods: Array.isArray(summary.periods) ? summary.periods.slice(0, 80) : summary.periods,
    filledCells: summary.filledCells,
    commentsAdded: summary.commentsAdded,
    warnings,
    ...(exposeDebugPath && summary.debugLogPath ? { debugLogPath: summary.debugLogPath } : {})
  };
}

async function readMultipartFormDataWithinLimit(request: NextRequest) {
  const contentType = request.headers.get("content-type") ?? "";
  if (!/^multipart\/form-data\b/i.test(contentType)) {
    throw new FillApiRequestError("Upload the workbook using multipart/form-data.", 415);
  }
  if (!request.body) throw new FillApiRequestError("The upload request body is empty.", 400);

  const reader = request.body.getReader();
  const chunks: Uint8Array[] = [];
  let bytesRead = 0;
  const deadlineAt = Date.now() + UPLOAD_READ_TIMEOUT_MS;
  try {
    while (true) {
      const { done, value } = await readUploadChunk(reader, request.signal, deadlineAt);
      if (done) break;
      bytesRead += value.byteLength;
      if (bytesRead > MAX_REQUEST_BYTES) {
        void reader.cancel("upload request exceeded configured size limit").catch(() => undefined);
        throw oversizedUploadError();
      }
      chunks.push(value);
    }
  } finally {
    reader.releaseLock();
  }

  if (request.signal.aborted) throw cancelledRequestError();
  const body = new Uint8Array(bytesRead);
  let offset = 0;
  for (const chunk of chunks) {
    body.set(chunk, offset);
    offset += chunk.byteLength;
  }
  try {
    return await new Response(body, { headers: { "content-type": contentType } }).formData();
  } catch {
    throw new FillApiRequestError("The multipart workbook upload could not be parsed.", 400);
  }
}

function readUploadChunk(
  reader: ReadableStreamDefaultReader<Uint8Array>,
  signal: AbortSignal,
  deadlineAt: number
): Promise<ReadableStreamReadResult<Uint8Array>> {
  if (signal.aborted) return Promise.reject(cancelledRequestError());
  const remainingMs = deadlineAt - Date.now();
  if (remainingMs <= 0) return Promise.reject(uploadTimeoutError());

  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (callback: () => void) => {
      if (settled) return;
      settled = true;
      signal.removeEventListener("abort", onAbort);
      clearTimeout(timeout);
      callback();
    };
    const onAbort = () => {
      void reader.cancel("client disconnected").catch(() => undefined);
      finish(() => reject(cancelledRequestError()));
    };
    const timeout = setTimeout(() => {
      void reader.cancel("upload read timed out").catch(() => undefined);
      finish(() => reject(uploadTimeoutError()));
    }, remainingMs);
    signal.addEventListener("abort", onAbort, { once: true });
    reader.read().then(
      (result) => finish(() => resolve(result)),
      (error) => finish(() => reject(signal.aborted ? cancelledRequestError() : error))
    );
  });
}

function hasValidBearerToken(authorization: string | null, configuredApiKey: string) {
  const match = authorization?.match(/^Bearer\s+(.+)$/i);
  if (!match) return false;
  const provided = Buffer.from(match[1].trim());
  const expected = Buffer.from(configuredApiKey);
  return provided.length === expected.length && timingSafeEqual(provided, expected);
}

function positiveInteger(value: string | undefined, fallback: number) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? Math.floor(parsed) : fallback;
}

function uploadLimitLabel() {
  const megabytes = MAX_UPLOAD_BYTES / 1024 / 1024;
  return megabytes >= 1 ? `${Number.isInteger(megabytes) ? megabytes : megabytes.toFixed(1)} MB` : `${MAX_UPLOAD_BYTES} bytes`;
}

function oversizedUploadError() {
  return new FillApiRequestError(`Workbook upload exceeds the ${uploadLimitLabel()} limit.`, 413);
}

function oversizedUploadResponse() {
  const error = oversizedUploadError();
  return jsonError(error.message, error.status);
}

function uploadTimeoutError() {
  return new FillApiRequestError("The workbook upload timed out before the request body was received.", 408);
}

function cancelledRequestError() {
  return new FillApiRequestError("Workbook fill cancelled because the client disconnected.", 499);
}

function cancelledResponse() {
  const error = cancelledRequestError();
  return jsonError(error.message, error.status);
}

function unauthorizedResponse() {
  return NextResponse.json(
    { error: "A valid deployment access key is required." },
    {
      status: 401,
      headers: {
        "www-authenticate": 'Bearer realm="Historicals Solver"',
        "cache-control": "no-store"
      }
    }
  );
}

function authenticationConfigurationResponse() {
  return NextResponse.json(
    {
      error:
        "This production deployment is not configured for workbook uploads. Set HISTORICALS_API_KEY, or explicitly set ALLOW_UNAUTHENTICATED_FILL=true only for an isolated trusted environment."
    },
    { status: 503, headers: { "cache-control": "no-store" } }
  );
}

function allowsUnauthenticatedFill() {
  return /^(?:1|true|yes)$/i.test(process.env.ALLOW_UNAUTHENTICATED_FILL || "");
}

function busyResponse() {
  return NextResponse.json(
    { error: "Another workbook fill is already running. Retry after it completes." },
    { status: 429, headers: { "retry-after": "15" } }
  );
}

async function acquireFillRequestLock(): Promise<FillRequestLock | null> {
  const lockPath = process.env.FILL_MODEL_PROCESS_LOCK_PATH || path.join(tmpdir(), "historicals-solver", "fill-model.active.lock");
  await mkdir(path.dirname(lockPath), { recursive: true });
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      const token = randomUUID();
      const handle = await open(lockPath, "wx");
      try {
        await handle.writeFile(JSON.stringify({ pid: process.pid, hostname: hostname(), startedAt: new Date().toISOString(), token }));
      } catch (error) {
        await handle.close().catch(() => undefined);
        await unlink(lockPath).catch(() => undefined);
        throw error;
      }
      return { handle, path: lockPath, token };
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
      const staleMs = Number(process.env.FILL_MODEL_PROCESS_LOCK_STALE_MS || (maxDuration + 60) * 1_000);
      const [fileMetadata, lockSnapshot] = await Promise.all([
        stat(lockPath).catch(() => null),
        readFile(lockPath, "utf8").catch(() => "")
      ]);
      if (!fileMetadata) continue;
      const lockMetadata = parseFillRequestLockMetadata(lockSnapshot);
      const ownerIsOnThisHost = !lockMetadata.hostname || lockMetadata.hostname === hostname();
      const ownerIsRunning = ownerIsOnThisHost && typeof lockMetadata.pid === "number" && processIsRunning(lockMetadata.pid);
      const ownerIsDead = ownerIsOnThisHost && typeof lockMetadata.pid === "number" && !ownerIsRunning;
      const lockIsTooOld = Date.now() - fileMetadata.mtimeMs > staleMs;
      if (ownerIsRunning) return null;
      if (lockMetadata.hostname && !ownerIsOnThisHost && lockMetadata.token) return null;
      if (!ownerIsDead && !lockIsTooOld) return null;

      // Re-read before unlinking so a concurrently replaced lock is not removed based on stale metadata.
      const currentSnapshot = await readFile(lockPath, "utf8").catch(() => null);
      if (currentSnapshot === null) continue;
      if (currentSnapshot !== lockSnapshot) return null;
      await unlink(lockPath).catch(() => undefined);
    }
  }
  return null;
}

async function releaseFillRequestLock(lock: FillRequestLock | null) {
  if (!lock) return;
  await lock.handle.close().catch(() => undefined);
  const metadata = parseFillRequestLockMetadata(await readFile(lock.path, "utf8").catch(() => ""));
  if (metadata.token === lock.token) await unlink(lock.path).catch(() => undefined);
}

function parseFillRequestLockMetadata(value: string): FillRequestLockMetadata {
  try {
    const parsed = JSON.parse(value) as FillRequestLockMetadata;
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function processIsRunning(pid: number) {
  if (!Number.isSafeInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "EPERM";
  }
}

function shouldExposeDebugPath() {
  return process.env.NODE_ENV !== "production" || /^(1|true|yes)$/i.test(process.env.EXPOSE_FILL_DEBUG_PATHS || "");
}

function jsonError(error: string, status: number, debugLogPath?: string) {
  return NextResponse.json(
    debugLogPath ? { error, debugLogPath } : { error },
    {
      status,
      headers: {
        "cache-control": "no-store",
        ...(debugLogPath ? { "x-debug-log-path": debugLogPath } : {})
      }
    }
  );
}
