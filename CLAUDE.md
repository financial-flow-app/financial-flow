# Coding rules for this repo

- Public repo: synthetic data only. No real names, numbers, accounts, income sources or organisations in
  code, tests, fixtures, commit messages, pull requests or issues.
- No network APIs in `src/` (fetch, XMLHttpRequest, WebSocket, EventSource, sendBeacon). The only exception
  is `sw.js`, and only for the app's own files at its own address. Tests enforce both.
- No outside scripts, fonts or analytics. Build the DOM with createElement (no HTML strings).
- Keep the Content-Security-Policy in `index.html` at least as strict as `test/hygiene.test.mjs` requires.
- The engine (`src/calc.js`) mirrors the workbook. Any change must keep `npm test` green, including
  parity with `test/parity/*.excel.json`. Re-record those only with `scripts/excel-parity.ps1` in Excel.
- Export must change only edited cells; everything else in the workbook stays byte-for-byte.
- After changing any app file, run `node scripts/stamp.mjs`. For a release, raise the version in
  `src/version.js` and `sw.js`, and add a CHANGELOG entry.
- Commit as `Claude <noreply@anthropic.com>`. Write files with LF line endings.
