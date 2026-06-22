const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const helperPath = path.join(repoRoot, "server/fill-model/sec-accession.ts");
const source = fs.readFileSync(helperPath, "utf8");
const output = ts.transpileModule(source, {
  compilerOptions: {
    module: ts.ModuleKind.CommonJS,
    target: ts.ScriptTarget.ES2020
  }
}).outputText;

const mod = { exports: {} };
new Function("exports", "require", "module", output)(mod.exports, require, mod);

const { cikFromAccession, normalizeAccession } = mod.exports;

const koCanonical = "000002134424000017";

assert.equal(normalizeAccession("0000021344-24-000017"), koCanonical);
assert.equal(normalizeAccession("000002134424000017"), koCanonical);
assert.equal(normalizeAccession(" 0000021344-24-000017 "), koCanonical);
assert.equal(normalizeAccession("21344-24-17"), koCanonical);
assert.equal(normalizeAccession("2134424000017"), koCanonical);
assert.equal(normalizeAccession("https://www.sec.gov/Archives/edgar/data/21344/000002134424000017/ko-20240329.htm"), koCanonical);

const currentCompanyAccessions = new Set([normalizeAccession("0000021344-24-000017")]);
assert.equal(currentCompanyAccessions.has(normalizeAccession("000002134424000017")), true);
assert.equal(cikFromAccession("000002134424000017"), "0000021344");
assert.equal(cikFromAccession("2134424000017"), "0000021344");

console.log("SEC accession normalization check passed.");
