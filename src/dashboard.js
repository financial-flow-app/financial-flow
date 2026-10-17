// The Dashboard: the workbook's Dashboard sheet, as data for the screen.
// Text lines are built exactly as the sheet builds them, so they can be checked against Excel.

import { round, dollars, YEAR, MONTH, DAYOF, DATE, EDATE, todaySerial } from './excel.js';

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
/** TEXT(serial, "mmmm d, yyyy") */
const longDate = (s) => `${MONTHS[MONTH(s) - 1]} ${DAYOF(s)}, ${YEAR(s)}`;
/** TEXT(serial, "m/d/yy") */
const mdyy = (s) => `${MONTH(s)}/${DAYOF(s)}/${String(YEAR(s) % 100).padStart(2, '0')}`;

/** The line under "Owed on cards & loans" (Dashboard J7). FFOS 3.9 shows the change; older workbooks a fixed line. */
function owedLine(m, trend) {
  const owed = m['Card & Loan Balances'];
  if (owed === 0) return { line: 'Nothing owed', tone: 'muted' };
  if (!trend) return { line: 'Paid from checking each month' };
  if (m['Latest done row (0 = none)'] === 0) return { line: 'As of opening', tone: 'muted' };
  const ch = m['Card & loan change vs previous'];
  const since = m['Previous done row'] === 0 ? 'at opening' : 'last month';
  if (ch < 0) return { line: `▼ ${dollars(Math.abs(ch))} less than ${since}`, tone: 'good' };
  if (ch > 0) return { line: `▲ ${dollars(ch)} more than ${since}`, tone: 'bad' };
  return { line: `Same as ${since}`, tone: 'muted' };
}

export function dashboardView(out, data, today = todaySerial()) {
  const m = out.metrics;
  const d = out.dashboard;
  const asOf = m['As-of date'];
  const nextStart = EDATE(DATE(YEAR(asOf), MONTH(asOf), 1), DAYOF(asOf + 1) === 1 ? 1 : 0);
  const nextIndex = out.months.findIndex((r) => r.MonthStart === nextStart);
  const change = m['Net worth change vs previous'];
  const avgCost = m['Average Monthly True Cost'];
  const avgN = data.settings.avgMonths;

  const latest = d.flow.latest, avg = d.flow.average, year = d.flow.thisYear;
  return {
    asOf,
    asOfText: longDate(asOf),
    next: { label: `${MONTHS[MONTH(nextStart) - 1]} ${YEAR(nextStart)}`, index: nextIndex },
    // The sheet's A2 line.
    headline: `As of ${longDate(asOf)}   ·   Next: fill ${MONTHS[MONTH(nextStart) - 1]} ${YEAR(nextStart)} on Month-End once its statements are out`,
    tiles: {
      netWorth: {
        value: m['Net Worth'],
        line: change === null
          ? `Starting point · opening ${mdyy(asOf)}`
          : `${change >= 0 ? '▲ ' : '▼ '}${dollars(Math.abs(change))}${m['Previous done row'] === 0 ? ' since opening' : ' vs last month'}`,
        tone: change === null ? 'muted' : change >= 0 ? 'good' : 'bad',
      },
      liquid: { value: m['Liquid Cash'], line: avgCost <= 0 ? '' : `${round(m['Liquid Cash'] / avgCost, 1).toFixed(1)} months of living costs` },
      investments: { value: m['Retirement + Investments'], line: `Contributed ${dollars(m['Contributions This Year'])} this year` },
      owed: { value: m['Card & Loan Balances'], ...owedLine(m, data.hasOwedTrend) },
    },
    flow: {
      headers: [`Latest month (${m['Latest done month label']})`, `Average (last ${avgN ?? ''} full)`, `This year (${YEAR(today)})`],
      rows: [
        ['Income', latest.income, avg.income, year.income],
        ['Household-paid living', latest.household, avg.household, year.household],
        ['Family-paid living', latest.family, avg.family, year.family],
        ['True cost of living', latest.trueCost, avg.trueCost, year.trueCost],
        ['   …of which one-time', latest.oneTime, null, year.oneTime],
        ['Tithe / charity given', latest.given, avg.given, year.given],
        ['Invested', latest.invested, null, year.invested],
        ['Saved (after giving)', latest.saved, null, year.saved],
      ],
    },
    tithe: {
      balance: m['Tithe Balance Now'],
      line: `This quarter: ${dollars(m['Tithe Given This Quarter'])} given of ${dollars(m['Tithe Owed This Quarter'] + 0)} owed`,
      given: m['Tithe Given This Quarter'],
      owed: m['Tithe Owed This Quarter'],
    },
    tax: d.tax, // null when the workbook has no tax option
    goals: d.goals,
    accounts: d.accounts,
    health: d.health,
  };
}
