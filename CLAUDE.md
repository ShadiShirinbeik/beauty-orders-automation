# Beauty Orders Automation

Google Apps Script project bound to a Google Sheet of mock beauty-shop orders.
It automates a manual ops workflow: flag bad rows, build a summary sheet,
and email a daily digest. Code is pushed with clasp and deployed via GitHub Actions.

## Tech rules
- Plain JavaScript on the Apps Script V8 runtime. No TypeScript, no npm packages, no bundler.
- All code lives in `Code.js` until it grows; then split by concern.
- Batch reads/writes with `getValues()` / `setValues()`; never touch cells one by one in a loop.
- Small single-purpose functions with a short JSDoc comment.
- Secrets (clasp credentials, API keys) are never committed; they live in GitHub Secrets.

## Workflow
- Propose a short plan before non-trivial changes.
- After a change, state which function to run to verify it.
- Prefer simple, readable code over clever code.