# Historicals Solver

> **Project status: archived / maintenance-only as of July 26, 2026.**
>
> The repository is preserved because it can produce useful SEC-backed workbook
> outputs with active analyst and Codex oversight. It is not a finished,
> autonomous, or generally reliable historicals product. See
> [PROJECT_CLOSEOUT.md](PROJECT_CLOSEOUT.md) for the verified state, unresolved
> defects, and reactivation requirements.

Historicals Solver is a local Next.js application that accepts a company ticker
or name and an `.xlsx` financial-model workbook. It retrieves financial data
from SEC EDGAR, maps reported facts into historical workbook periods, preserves
supported OOXML features, adds audit/provenance information, and returns a
filled workbook. An optional OpenRouter-backed accounting workflow reviews
ambiguous mappings.

## Preserved usage

Requirements:

- Node.js and npm. No Node version is pinned; the closeout verification used
  Node.js 24.15.0 and npm 11.12.1.
- A valid `SEC_USER_AGENT` containing the application identity and a real,
  monitored operator email address.
- An `OPENROUTER_API_KEY` for the intended LLM-led mapping path.

```bash
npm install
cp .env.example .env.local
# Configure SEC_USER_AGENT and any optional keys/settings in .env.local.
npm run dev:ensure
```

Then open <http://localhost:3000>, enter a ticker or company name, select an
`.xlsx` workbook, and review every warning, audit note, and source-ledger result
before relying on the output.

Operational constraints in the current API include:

- `.xlsx` only; `.xlsm` is rejected because VBA is not safely preserved.
- 30 MB uploads by default.
- One fill at a time by default.
- A maximum request duration of 900 seconds in the Next.js route.
- Production uploads fail closed unless `HISTORICALS_API_KEY` is configured or
  unauthenticated use is explicitly enabled for an isolated environment.

## Verification

The maintained default checks are:

```bash
npm test
npm run build
```

Those commands passed during closeout, but they do not prove end-to-end or
general-company reliability. Many additional scripts require live SEC access,
local workbook fixtures, a running server, and sometimes LLM credentials. Two
self-contained regression scripts also fail at the archived baseline; details
are in [PROJECT_CLOSEOUT.md](PROJECT_CLOSEOUT.md).

## Repository map

- `app/`: upload UI and workbook-fill API route.
- `server/fill-model/`: EDGAR retrieval, normalization, accounting mapping,
  validation, workbook writing, provenance, and OOXML preservation.
- `scripts/`: deterministic guards plus workstation- and network-dependent
  regression harnesses.
- `config/`: example approved-mapping, gold-model, and regression manifests.
- `github/templates/`: preserved workbook templates.
- `github/example-excels/`: historical output snapshots, not a current quality
  guarantee.

Financial statement values must continue to come from SEC EDGAR. Do not commit
credentials, local caches, temporary outputs, or newly generated workbooks.
