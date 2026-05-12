# Codex Collaboration Rules

## Git Sync

- Use `codex/rsg-segment-historicals` as the shared working branch unless the user explicitly names another branch.
- When a task is complete, run the relevant checks, commit the completed source changes, and push the branch to `origin`.
- Do not leave completed source changes only in a local worktree. Other Codex windows should be able to sync by fetching/pulling the shared branch.
- Do not commit generated artifacts, temporary workbooks, local caches, or build outputs unless the user specifically asks for them.
- Merge to `main` only through an intentional merge or pull request after the shared branch is verified.

## Local Server

- After source changes, make sure the dev server is running and verify `http://localhost:3000` responds successfully.
