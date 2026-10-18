# Changelog

## 0.8.1

- Safer updates: a new version's files are fetched under its build number, so GitHub's cache can't hand back an older copy, and they're kept only if their fingerprint matches the build. An update can no longer store a mix of old and new files; if something doesn't match, the version already on the phone keeps working and the next visit tries again.
- The browser check now cuts the network and reopens the app, and checks that it opens from the phone's copy with its data (on Chromium; Playwright's WebKit can't do this).

## 0.8.0

- Add accounts, investments and income sources inside a month in the Ledger (+ Add account, + Add investment, + Add income source). A new account takes an unused row of Settings, starts at $0 and isn't required, so no past month changes. The ⋯ beside each one renames it (goals that track an account follow), closes it in that month ($0 that month, then marked closed) or ends an income source (its Until date).
- The Ledger tab replaces Month and History: the month to fill next, this year's totals, the average month and every month (tap one to fill or fix it).
- No menu: Health checks open from the Dashboard's warning line, and the Workbook page (export, import, the app's version) from the chip in the top bar, which is always shown. The red dot is gone.

## 0.7.0

- Each tab has one job. Along the bottom: Dashboard, Month, Giving and Goals. At the top, ☰ More holds History, Import & export (was the File tab), Health checks and About.
- The Dashboard fits one phone screen: the most urgent health check with "+N more", net worth, liquid cash, owed, investments, tithe this quarter, a goals line and the accounts. Net worth opens History; tithe opens Giving; the goals line opens Goals.
- Goals has its own tab. History shows this year's totals, the average month and every month since the opening date. Giving shows the tithe balance and this year's tax-deductible giving at the top.
- The More button shows a red dot when an export is due; the phone's Back button returns from a page under More to its menu.

## 0.6.0

- Goals can be added, changed and deleted on the Dashboard. The form asks only for what the goal's type needs, with choices from your workbook. Export writes them into the goal rows of Settings, and nothing else there; other settings are still changed in Excel.

## 0.5.2

- The app turns with the phone. It used to stay upright when the phone was turned sideways. An app already on the home screen may take up to a day to pick this up; adding it to the home screen again makes it immediate (export first, to be safe).

## 0.5.1

- A ± button beside the “Moved into” fields, for phone number pads without a minus key.
- Export names show the time in 12-hour form, for example “FFOS 2026-10-18 12.01 AM.xlsx”.

## 0.5.0 (first release)

- Dashboard, Month, Giving and File screens for workbooks in the FFOS 4.0 format.
- Calculations identical to the workbook's, checked against Excel's own results.
- Export writes only the edited cells, names the file with the date and time, and stamps it, so an
  older copy is caught on import.
- Works offline; installs to the home screen; new versions wait for an Update tap.
- Reminder to export when entries have waited a month.
