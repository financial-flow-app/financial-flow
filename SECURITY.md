# Security

Financial Flow has no server and no accounts: the workbook and your entries live only in your phone's
browser, until you export them. That leaves a few places where something could go wrong, and they are
what we protect:

- **The code that runs next to your numbers.** It comes only from this repository. `main` accepts
  changes only through a pull request whose automatic checks pass. The page tells the browser to run
  only the app's own files and to open no network connections for data.
- **Offline copies of the app.** A service worker keeps the app's own files on the phone. It loads
  files only from the app's own address, never your data, and a new version runs only after you tap Update.
- **Workbook files.** Text in a workbook is always treated as text, never as code. Export changes only
  the cells you edited.
- **Exported files.** An exported workbook holds your numbers. Keep it private.

## Reporting a problem

If you find a security or privacy problem, please don't post the details publicly. Use
**Security → Report a vulnerability** on this repository if it's available, or open an issue that says
only "security report, please contact me", and we'll reach you privately.

Please never include real numbers, names or workbook files in a report.
