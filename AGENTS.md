# Codex Collaboration Rules

## Git Sync

- Use `codex/rsg-segment-historicals` as the shared working branch unless the user explicitly names another branch.
- Before starting new work, run `git fetch origin` and sync with `origin/codex/rsg-segment-historicals` so every Codex window starts from the latest shared code.
- When a task is complete, run the relevant checks, commit the completed source changes, and push the branch to `origin`.
- Do not leave completed source changes only in a local worktree. Other Codex windows should be able to sync by fetching/pulling the shared branch.
- Do not commit generated artifacts, temporary workbooks, local caches, or build outputs unless the user specifically asks for them.
- If another Codex window has pushed first, integrate the remote branch before pushing, resolving conflicts in favor of the combined verified behavior.
- Merge to `main` only from the shared branch after all active Codex windows have pushed their completed work and the combined branch is verified.
- The final output should come from the merged shared branch, not from an individual unsynced local window.

## Financial Data Source

- All financial statement data, model historicals, operating metrics, segment data, and other financial numbers must come from SEC EDGAR filings.
- Treat EDGAR as the source of truth. Start from SEC filings and SEC data endpoints linked through https://www.sec.gov/search-filings.
- Do not use third-party finance websites, scraped summaries, analyst estimates, or manually invented figures as source data for workbook outputs.
- If EDGAR does not provide a requested number directly, derive it only from EDGAR-sourced values and leave a note explaining the derivation.

## Local Server

- After source changes, make sure the dev server is running and verify `http://localhost:3000` responds successfully.
