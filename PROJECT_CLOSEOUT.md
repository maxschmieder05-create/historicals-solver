# Project Closeout

## Decision

Historicals Solver was archived on July 26, 2026. There is no active development
backlog or implied commitment to finish the product.

The code is retained because repeated, targeted Codex and analyst intervention
can produce useful results. The project did not reach the stronger goal of
reliably accepting new companies and workbook templates without continuing
engineering supervision. It must therefore be described and used as an
operator-assisted prototype, not as an autonomous production system.

This closeout preserves all code and artifacts. Nothing was deleted, and no
attempt was made to rewrite history or conceal known failures.

## Archived baseline

- Shared branch: `codex/rsg-segment-historicals`
- Audited source commit: `ea3f67a263537770f21ea6b3438de4fc0de6b058`
- Source history at audit: 243 commits from May 11 through July 24, 2026
- Runtime used for closeout: Node.js 24.15.0 and npm 11.12.1
- Application: Next.js 14.2.35
- Source policy: SEC EDGAR is the source of truth for all financial values

The closeout documentation and ignore-file hardening are committed after the
audited source commit above.

## What is implemented

The repository contains a functioning local upload flow and a substantial
workbook translation pipeline:

- A browser UI for ticker/company input, deployment access key entry, `.xlsx`
  selection, cancellation, result download, and warning display.
- A guarded API route with upload limits, bearer authentication, request
  cancellation, single-process locking, and bounded concurrency.
- SEC company lookup, Company Facts access, filing-package parsing, accession
  normalization, retry/timeout handling, and optional bulk-archive caching.
- Period normalization and annual/quarterly fact handling.
- Income-statement, balance-sheet, cash-flow, and segment mapping logic.
- Structured line-item classification, balance-sheet row resolution, liability
  classification, source-ledger validation, audit comments, and fail-closed
  behavior for some ambiguous cases.
- Optional LLM accounting analysis through OpenRouter, including cost, timeout,
  model-fallback, validation, and telemetry controls.
- OOXML feature capture/restoration, package-safety checks, formula protection,
  and workbook recalculation settings.
- Approved-mapping and verified-gold-model mechanisms.
- Six committed workbook templates, 98 committed filled-output snapshots, and
  67 check scripts.

## Verification performed at closeout

The repository was clean and matched
`origin/codex/rsg-segment-historicals` before the audit.

The following passed on July 26, 2026:

- `npm test`, including TypeScript, ESLint, approved-mapping cache, normalized
  architecture, line-item classifier, liability rules, balance-sheet resolver,
  API guards, SEC runtime hardening, OOXML preservation, and package safety.
- `npm run build`, including the production Next.js compilation, lint, type
  validation, and route generation.
- `npm run dev:ensure`, followed by an HTTP 200 response from
  `http://localhost:3000`; the local server remains supervised.
- An additional self-contained sweep of 25 scripts covering ABT/ANET income
  reconciliation, UI/auth guards, cash-flow completion, formula normalization,
  gold-model scanning, IBM fail-closed behavior, LLM controls/workbench, SEC
  filing and bulk-cache behavior, segment formula/fallback behavior, source
  ledger status, and unusual SEC presentations.

This is meaningful component coverage. It is not evidence that an arbitrary
issuer/template pair will generate a correct workbook.

## Known unresolved defects and gaps

### Reproducible non-default regression failures

These scripts are not included in `npm test`, but failed during the closeout
audit:

1. `node scripts/amd-balance-sheet-liability-classification-check.js`
   fails at line 560. A treasury row that ties current SEC treasury plus
   employee-trust contra-equity is not receiving the expected
   narrow-primary-ledger advisory.
2. `node scripts/amd-income-statement-classification-check.js`
   fails at line 374. A globally empty income-statement assignment ledger does
   not fail closed for an SEC-reported requested period as the regression
   expects.

These should be treated as product defects or stale test/behavior contracts
until someone deliberately investigates and resolves the mismatch.

### The default quality gate is incomplete

`npm test` covers nine core script groups but omits many self-contained
regressions, including the two failures above. It also omits the live
company/template regressions. A green default suite is therefore necessary but
insufficient.

### End-to-end verification is not portable or hermetic

Thirty check scripts contain `FILL_API_URL` integration behavior or hard-coded
paths under one operator's `/Users/.../Downloads` or Desktop folders. They may
also require:

- a running fill server;
- an exported, valid `SEC_USER_AGENT`;
- live SEC availability and acceptable request timing;
- local input/gold workbooks that are not reproducibly provisioned; and
- OpenRouter credentials and currently available model identifiers.

Even `npm run test:regression-preflight` does not load `.env.local` itself; the
closeout invocation stopped because `SEC_USER_AGENT` was not exported in the
calling shell. The full live regression basket was not run during closeout.

### Generalization remains unproven

The commit history contains repeated fixes for individual issuers, statement
presentations, templates, and workbook-picker behavior. Many of those fixes
were generalized into reusable rules, but the aggregate evidence still shows a
reactive maintenance loop. New SEC presentations and workbook structures can
require another code change or Codex-guided repair.

The output must be considered draft analyst work. Review source-ledger status,
audit comments, blank cells, warnings, statement tie-outs, segment totals,
period alignment, sign conventions, and preserved formulas before use.

### Maintainability and repository cost

- `server/fill-model/fill-model-service.ts` is approximately 34,725 lines and
  combines many stages of the pipeline. Changes have a large regression surface.
- Application, server, and check-script code total roughly 71,700 lines.
- The working directory was about 5.4 GB during closeout because of local caches,
  dependencies, build output, and temporary runs.
- Tracked files were about 254 MB; `github/example-excels` alone was about
  240 MB. The 98 output workbooks are historical snapshots and include many
  numbered retries. They should not be mistaken for curated gold fixtures.
- There is no pinned Node version, automated CI workflow, deployment manifest,
  database, queue, or distributed lock. The default API configuration processes
  one fill at a time.

### External dependencies remain operational risks

Correct operation depends on SEC endpoint availability and identity/rate-limit
compliance. The intended analyst flow also depends on OpenRouter model access,
pricing, latency, and structured-response behavior. Configuration contains
fallback and cost controls, but those do not make the dependencies deterministic.

## Security and data handling

- Keep all financial data SEC-sourced. Do not substitute third-party finance
  sites or analyst estimates.
- Never commit `.env` files, API-key text files, local manifests containing
  absolute/private paths, debug logs, temporary workbooks, caches, or new filled
  outputs.
- Production upload access fails closed unless `HISTORICALS_API_KEY` is set or
  unauthenticated access is explicitly enabled.
- Debug paths are hidden in production unless explicitly exposed.
- Only `.xlsx` is accepted. Macro-enabled workbooks are rejected because VBA
  preservation is not supported safely.

The closeout expanded `.gitignore` to cover `.env*` except `.env.example` and to
ignore `api.txt` in every clone, rather than relying on a workstation-specific
Git exclude.

## If the project is ever reactivated

Start by fetching the shared branch and reading this document. Do not resume the
old open-ended repair loop. A credible reactivation should have a bounded target
and, at minimum:

1. Resolve the two failing self-contained regressions and add all deterministic
   checks to the default quality gate.
2. Replace workstation-specific fixture paths with versioned, minimal,
   rights-cleared fixtures or an explicit external-fixture provisioning process.
3. Establish a reproducible end-to-end regression basket with reviewed expected
   outputs and a clear pass/fail threshold.
4. Define the exact supported workbook templates, SEC statement types, issuer
   classes, and failure policy. Reject unsupported inputs explicitly.
5. Require human review for every produced workbook until measured accuracy
   justifies a different policy.
6. Break the monolithic fill service into independently testable pipeline stages.
7. Add CI, pin the supported Node/npm versions, and document deployment and
   rollback procedures before any production claim.

Until those conditions are met, the safest use is preservation, demonstration,
or narrowly scoped runs supervised by someone capable of checking the resulting
financial model.

## Lessons retained

- Passing examples are not a universal specification. A supported-input contract
  and curated regression set must precede broad automation claims.
- Company-specific fixes only become durable when encoded as statement- and
  concept-level rules and exercised across multiple issuers.
- Financial automation needs fail-closed validation and visible provenance from
  the beginning, not after mapping heuristics accumulate.
- Tests that depend on one workstation are operational notes, not a repeatable
  quality gate.
- LLM review can help resolve ambiguity, but it does not remove the need for
  deterministic accounting constraints, reconciliation, and human accountability.
- A project that works only through continuous expert or Codex intervention has
  useful components, but it is not yet a self-sustaining product.
