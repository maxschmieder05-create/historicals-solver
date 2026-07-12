const assert = require("node:assert/strict");
const { __fillWorkbookApiTestHooks } = require("./fill-workbook-api");

const originalFillApiKey = process.env.FILL_API_KEY;
const originalHistoricalsApiKey = process.env.HISTORICALS_API_KEY;

try {
  process.env.FILL_API_KEY = "client-specific-key";
  process.env.HISTORICALS_API_KEY = "server-fallback-key";
  assert.equal(__fillWorkbookApiTestHooks.deploymentAccessKey(), "client-specific-key");

  delete process.env.FILL_API_KEY;
  assert.equal(__fillWorkbookApiTestHooks.deploymentAccessKey(), "server-fallback-key");

  const invocation = __fillWorkbookApiTestHooks.buildCurlInvocation({
    apiUrl: "https://historicals.example/api/fill-model",
    ticker: "TEST",
    inputWorkbook: "/tmp/input.xlsx",
    outputWorkbook: "/tmp/output.xlsx",
    accessKey: 'secret-\\-"-value'
  });
  assert.deepEqual(invocation.args.slice(0, 2), ["--config", "-"]);
  assert.equal(invocation.args.some((argument) => argument.includes("secret-")), false, "the key must not be exposed in curl process arguments");
  assert.equal(invocation.stdinText, 'header = "Authorization: Bearer secret-\\\\-\\"-value"\n');

  const unauthenticated = __fillWorkbookApiTestHooks.buildCurlInvocation({
    apiUrl: "http://localhost:3000/api/fill-model",
    ticker: "TEST",
    inputWorkbook: "/tmp/input.xlsx",
    outputWorkbook: "/tmp/output.xlsx"
  });
  assert.equal(unauthenticated.args.includes("--config"), false);
  assert.equal(unauthenticated.stdinText, "");

  assert.throws(
    () =>
      __fillWorkbookApiTestHooks.buildCurlInvocation({
        apiUrl: "http://localhost:3000/api/fill-model",
        ticker: "TEST",
        inputWorkbook: "/tmp/input.xlsx",
        outputWorkbook: "/tmp/output.xlsx",
        accessKey: "unsafe\r\nInjected: header"
      }),
    /cannot contain line breaks/i
  );
} finally {
  if (originalFillApiKey === undefined) delete process.env.FILL_API_KEY;
  else process.env.FILL_API_KEY = originalFillApiKey;
  if (originalHistoricalsApiKey === undefined) delete process.env.HISTORICALS_API_KEY;
  else process.env.HISTORICALS_API_KEY = originalHistoricalsApiKey;
}

console.log("Fill-workbook API bearer-auth client guards passed.");
