# Financial Flow

A phone app for a monthly household-finance update and a dashboard, on top of an Excel workbook.
The workbook stays the record: the app imports it, you enter the month and your donations on the phone,
and the app exports the same workbook with your entries written in. If the app ever goes away, Excel carries on.

**Open the app:** https://financial-flow-app.github.io/financial-flow/ · **Guide:** [wiki](https://github.com/financial-flow-app/financial-flow/wiki) · **Versions:** [releases](https://github.com/financial-flow-app/financial-flow/releases) and [CHANGELOG](CHANGELOG.md)

## What it does

- **Dashboard:** net worth and its change, liquid cash, investments, amounts owed, money flow,
  tithe balance for the quarter, tax lines for the year, goals, accounts and health checks.
- **Month:** month-end balances, income, family-paid and one-time costs, investments and a note.
  Automatic values show in grey; type to override, or clear to go back to automatic.
- **Giving:** add, edit and delete donations.
- **File:** import a workbook; export it through the phone's share sheet (for example to a PC).

It reads workbooks in the FFOS 4.0 format (a specific layout with Settings, Month-End, Giving, History
and Calc sheets). Every number on the Dashboard is calculated exactly as the workbook calculates it.

## Privacy

- Your data stays on your phone, in the browser's storage for this app. It leaves only when you export.
- The page blocks every network connection for data (a strict Content-Security-Policy), loads no
  outside scripts, fonts or analytics, and has no server, account or bank connection.
- After the first visit it works offline. New versions wait until you tap **Update**.

More in [SECURITY.md](SECURITY.md).

## For developers

Plain HTML, CSS and JavaScript (ES modules), no dependencies and no build step.

```
npm test                       # all tests (Node 24)
node scripts/serve.mjs         # local test server on http://127.0.0.1:5180
node scripts/stamp.mjs         # after changing any app file: refresh the build stamp in sw.js
```

- `src/` the app: `zip.js` and `xlsx.js` (read and patch the workbook, changing only edited cells),
  `import.js`, `calc.js` (the workbook's calculations), `dashboard.js`, `entry.js`, `app.js` (screens),
  `store.js` (browser storage).
- `test/` tests. `test/parity/` holds Excel's own results for synthetic scenarios; the engine must match them.
- `scripts/` Excel checks (Windows, Excel installed), fixtures, icons, build stamp, test server.
- `fixtures/` the blank workbook template (synthetic, no real data).

Changes go through a pull request; the **Check** workflow must pass.

## License

All rights reserved. See [LICENSE](LICENSE).
