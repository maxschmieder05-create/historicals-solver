const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const repoRoot = path.resolve(__dirname, "..");
const packageJson = require(path.join(repoRoot, "package.json"));
const managerSource = fs.readFileSync(path.join(repoRoot, "scripts", "ensure-dev-server.js"), "utf8");
const agentRules = fs.readFileSync(path.join(repoRoot, "AGENTS.md"), "utf8");
const { positiveInteger } = require("./ensure-dev-server");

assert.equal(packageJson.scripts["dev:ensure"], "node scripts/ensure-dev-server.js");
assert.equal(packageJson.scripts["dev:stop"], "node scripts/ensure-dev-server.js --stop");
assert.equal(positiveInteger("3010", 3000), 3010);
assert.equal(positiveInteger("invalid", 3000), 3000);
assert.match(managerSource, /detached:\s*true/, "the supervisor must be detached from the launching shell");
assert.match(managerSource, /supervisor\.unref\(\)/, "the launcher must not keep its short-lived parent attached");
assert.match(managerSource, /while \(!stopping\)[\s\S]*Next\.js exited unexpectedly/, "the supervisor must restart an unexpected Next.js exit");
assert.match(managerSource, /statusCode >= 200 && response\.statusCode < 400/, "HTTP error pages must not be treated as healthy");
assert.match(agentRules, /npm run dev:ensure/, "future Codex sessions must use the durable server launcher");

console.log("Dev-server reliability guard passed.");
