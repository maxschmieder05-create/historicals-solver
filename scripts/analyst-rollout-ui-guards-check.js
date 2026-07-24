const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const repoRoot = path.resolve(__dirname, "..");
const pageSource = fs.readFileSync(path.join(repoRoot, "app", "page.tsx"), "utf8");

const accessKeyIdIndex = pageSource.indexOf('id="deployment-access-key"');
const accessKeyInputStart = pageSource.lastIndexOf("<input", accessKeyIdIndex);
const accessKeyInputEnd = pageSource.indexOf("/>", accessKeyIdIndex);
const accessKeyInput =
  accessKeyIdIndex >= 0 && accessKeyInputStart >= 0 && accessKeyInputEnd >= 0
    ? pageSource.slice(accessKeyInputStart, accessKeyInputEnd + 2)
    : "";
assert.ok(accessKeyInput, "the analyst page must render a deployment access-key input");
assert.match(accessKeyInput, /type="password"/, "the access key must not be displayed as plain text");
assert.match(accessKeyInput, /autoComplete="off"/, "the page must not ask the browser to persist the deployment key");
assert.doesNotMatch(accessKeyInput, /\bname=/, "the access key must not be serialized into the multipart form body");
assert.doesNotMatch(pageSource, /(?:localStorage|sessionStorage|document\.cookie)/, "the page must not persist the deployment access key in browser storage");
assert.match(
  pageSource,
  /headers:\s*normalizedAccessKey\s*\?\s*\{ Authorization: `Bearer \$\{normalizedAccessKey\}` \}\s*:\s*undefined/,
  "the analyst-supplied key must be sent only as a bearer authorization header"
);
assert.match(pageSource, /never added to the workbook/i, "the page must explain how the deployment key is handled");

assert.match(pageSource, /const activeRequestRef = useRef<AbortController \| null>\(null\)/);
assert.match(pageSource, /signal:\s*controller\.signal/, "the browser fetch must use an AbortController signal");
assert.match(pageSource, /function cancelFill\(\)[\s\S]*activeRequestRef\.current\.abort\(\)/, "the cancel control must abort the active request");
assert.match(pageSource, /fillRequestErrorMessage\(caught, controller\.signal\.aborted\)/, "the page must classify cancellation from the active request signal");
assert.match(pageSource, /if \(wasCancelled\) return "Workbook fill cancelled/, "the page must show a clear cancellation outcome");
assert.match(pageSource, /type="button" onClick=\{cancelFill\}/, "cancel must be a non-submit button");
assert.match(
  pageSource,
  /function fillRequestErrorMessage\([\s\S]*failed to fetch[\s\S]*app server connection was lost[\s\S]*npm run dev:ensure/i,
  "browser network failures must be translated into an actionable local-server recovery message"
);

console.log("Analyst access-key and cancellation UI guards passed.");
