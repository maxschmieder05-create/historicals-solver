const assert = require("node:assert/strict");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const {
  SEC_USER_AGENT_CONFIGURATION_MESSAGE,
  SecRegressionIdentityError,
  requireSecRegressionHeaders,
  validSecRegressionUserAgent
} = require("./sec-regression-identity");

const validUserAgent = "HistoricalsSolver SEC regression sec-operations@historicalssolver.dev";
const headers = requireSecRegressionHeaders({ SEC_USER_AGENT: `  ${validUserAgent}  ` });

assert.deepEqual(headers, { "User-Agent": validUserAgent });
assert.equal(Object.isFrozen(headers), true);
assert.equal(validSecRegressionUserAgent(validUserAgent), true);

for (const invalidUserAgent of [
  undefined,
  "",
  "contact@example.com",
  "HistoricalsSolver contact@example.com",
  "HistoricalsSolver analyst@localhost",
  "HistoricalsSolver analyst@your-domain.test",
  "HistoricalsSolver no-reply@historicalssolver.dev",
  "HistoricalsSolver noreply+sec@historicalssolver.dev",
  "HistoricalsSolver test@historicalssolver.dev",
  "HistoricalsSolver analyst@example.dev",
  "HistoricalsSolver analyst@historicalssolver.dev\r\nX-Injected: yes",
  "your-app sec-operations@historicalssolver.dev",
  "sec-operations@historicalssolver.dev"
]) {
  assert.equal(validSecRegressionUserAgent(invalidUserAgent), false, `expected rejection for ${String(invalidUserAgent)}`);
  assert.throws(
    () => requireSecRegressionHeaders({ SEC_USER_AGENT: invalidUserAgent }),
    (error) =>
      error instanceof SecRegressionIdentityError &&
      error.message.includes(SEC_USER_AGENT_CONFIGURATION_MESSAGE) &&
      error.message.includes("before this script can make a network request")
  );
}

const liveRegressionScripts = [
  "regression-basket-check.js",
  "balance-sheet-instant-regression-check.js",
  "cpb-actualized-forecast-period-check.js",
  "wmt-fiscal-year-regression-check.js",
  "balance-sheet-conservation-check.js",
  "goog-current-investments-regression-check.js"
];
for (const script of liveRegressionScripts) {
  const env = { ...process.env };
  delete env.SEC_USER_AGENT;
  const result = spawnSync(process.execPath, [path.join(__dirname, script)], {
    cwd: path.resolve(__dirname, ".."),
    env,
    encoding: "utf8"
  });
  assert.notEqual(result.status, 0, `${script} must fail closed without SEC_USER_AGENT`);
  assert.match(
    `${result.stdout}\n${result.stderr}`,
    /SEC_USER_AGENT is required for live SEC regression scripts.*before this script can make a network request/s,
    `${script} must report the identity configuration error before doing work`
  );
}

console.log("SEC regression identity checks passed.");
