// The calculation engine: the workbook's Calc and History sheets, rebuilt in JavaScript.
// Per-month values use the Calc sheet's column names (row 5), so each one can be checked
// against Excel cell by cell. Blank results ("" in Excel) are null.

import {
  isBlank, N, num, is, cmpValue, maxNums, minNums, round, mod,
  YEAR, MONTH, DAYOF, DATE, EOMONTH, EDATE, WEEKDAY, quarterStart,
  dollars, monYear, monthYear, mdyy,
} from './excel.js';
import { MONTH_COUNT, FIRST_MONTH } from './ffos.js';

const WEEKDAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
const ASSET = ['Checking', 'Savings', 'Cash'];
const LIABILITY = ['Card', 'Loan'];
const isAny = (v, list) => list.some((t) => is(v, t));
const sign = (type) => (isAny(type, ASSET) ? 1 : isAny(type, LIABILITY) ? -1 : 0);
const sum = (xs) => xs.reduce((a, b) => a + b, 0);
const QUARTERS = 60;
const FIRST_ROW = 6; // Calc row of the first month

export const METRIC_NAMES = [
  'Net Worth', 'Liquid Cash', 'Investments', 'Retirement', 'Retirement + Investments', 'Card & Loan Balances',
  'Average Monthly Income', 'Average Monthly Household-Paid', 'Average Monthly Family-Paid', 'Average Monthly True Cost',
  'Income This Year', 'Giving This Year', 'Tithe-able Income This Year', 'Contributions This Year',
  'Retirement Contributions This Year', 'Saved This Year', 'Tithe Owed This Quarter', 'Tithe Given This Quarter',
  'Tithe Balance Now', 'Average Monthly Giving', 'Specific Account',
];

/** The value Excel sees in a Month-End input cell: a typed constant or a custom formula's result. */
const cellValue = (c) => (c.kind === 'auto' ? null : c.value);
const typedConst = (c) => c.kind === 'typed';

/** Automatic income for one source in [windowStart, min(windowEnd, Until)]. */
function autoIncome(src, A, B, windowEnd) {
  if (isBlank(src.usual) || isBlank(src.frequency)) return 0;
  const end = minNums(windowEnd, isBlank(src.until) ? windowEnd : src.until);
  const start = maxNums(A, src.from);
  let times;
  if (is(src.frequency, 'Weekly')) {
    const wd = typeof src.day === 'string' ? WEEKDAYS.indexOf(src.day.toLowerCase()) + 1 : 0;
    if (wd === 0) times = 0;
    else if (end < start) times = 0;
    else times = Math.floor((end - start - mod(wd - WEEKDAY(start), 7)) / 7) + 1;
  } else {
    const pay = isBlank(src.day) ? B : DATE(YEAR(A), MONTH(A), minNums(src.day, DAYOF(B)));
    times = pay >= start && pay <= end ? 1 : 0;
  }
  return num(src.usual) * times;
}

/** Recurring donations counted in [max(A, From, notBefore), min(EffEnd, Until)]. */
// deductibleOnly: count only donations marked Tax-deductible = Yes (FFOS 3.8).
function recurringGiving(recurring, A, B, C, notBefore = null, deductibleOnly = false) {
  let total = 0;
  for (const r of recurring) {
    if (isBlank(r.amount) || is(r.household, 'No')) continue;
    if (deductibleOnly && !is(r.deductible, 'Yes')) continue;
    const pay = isBlank(r.day) ? B : DATE(YEAR(A), MONTH(A), minNums(r.day, DAYOF(B)));
    const start = notBefore === null ? maxNums(A, r.from) : maxNums(A, r.from, notBefore);
    const end = minNums(C, isBlank(r.until) ? C : r.until);
    total += num(r.amount) * (pay >= start && pay <= end ? 1 : 0);
  }
  return total;
}

export function calculate(data, today) {
  const { settings: st, accounts, investments, sources, recurring, goals } = data;
  const open = cmpValue(st.openingDate);
  const openNum = typeof st.openingDate === 'number' ? st.openingDate : 0;
  const giving = data.giving
    .filter((g) => typeof g.date === 'number')
    .map((g) => ({ date: g.date, amount: typeof g.amount === 'number' ? g.amount : 0, family: is(g.household, 'No'), deductible: is(g.deductible, 'Yes') }));
  const givingSum = (from, to, family, fromExclusive = false) =>
    sum(giving.filter((g) => (fromExclusive ? g.date > from : g.date >= from) && g.date <= to && g.family === family).map((g) => g.amount));

  const requiredCount = accounts.filter((a) => is(a.required, 'Yes') && !isBlank(a.name) && !is(a.active, 'No')).length;
  const openingNetCash = sum(accounts.map((a) => sign(a.type) * N(a.opening)));
  const openingOwed = sum(accounts.map((a) => (isAny(a.type, LIABILITY) ? N(a.opening) : 0))); // Calc!EB65 (FFOS 3.9)

  const months = [];
  for (let i = 0; i < MONTH_COUNT; i++) {
    const m = data.months[i];
    const prev = months[i - 1];
    const r = { ym: m.ym, row: FIRST_ROW + i };
    const A = DATE(FIRST_MONTH.year, FIRST_MONTH.month + i, 1);
    const B = EOMONTH(A, 0);
    const C = Math.min(B, today);
    const D = A > today ? 1 : 0;
    const E = B <= open ? 'pre' : A <= open ? 'open' : 'post';
    Object.assign(r, { MonthStart: A, MonthEnd: B, EffEnd: C, Future: D, Phase: E });

    // Income
    const inc = [];
    sources.forEach((src, s) => {
      const auto = autoIncome(src, A, B, C);
      r[`Auto${s}`] = auto;
      const v = cellValue(m.income[s]);
      inc[s] = isBlank(v) ? auto : N(v);
      r[`Inc${s}`] = inc[s];
      r[`PreAuto${s}`] = E !== 'open' || typedConst(m.income[s]) ? 0 : autoIncome(src, A, B, st.openingDate);
    });
    r.Income = sum(inc);
    r.TitheInc = B < cmpValue(st.titheStart) ? 0 : sum(inc.filter((_, s) => is(sources[s].tithe, 'Yes')));
    r.NotSureInc = sum(inc.filter((_, s) => is(sources[s].tithe, 'Not Sure')));
    // Tax (FFOS 3.8): income from taxable sources, and the part with no tax withheld.
    r.TaxableInc = sum(inc.filter((_, s) => is(sources[s].taxable, 'Yes')));
    r.UnwithheldInc = sum(inc.filter((_, s) => is(sources[s].taxable, 'Yes') && is(sources[s].withheld, 'No')));
    r.PreOpenInc = E === 'open' ? sum(sources.map((_, s) => r[`PreAuto${s}`])) + N(st.openingTypedIncome) : 0;
    r.IncSpend = E === 'post' ? r.Income : E === 'open' ? r.Income - r.PreOpenInc : 0;

    // Giving
    r.GiveLog = givingSum(A, B, false);
    r.GiveFam = givingSum(A, B, true);
    r.GiveRec = recurringGiving(recurring, A, B, C);
    // Deductible giving counts a marked row whoever paid; recurring ones only when household-funded.
    r.DeductGive = sum(giving.filter((g) => g.deductible && g.date >= A && g.date <= B).map((g) => g.amount))
      + recurringGiving(recurring, A, B, C, null, true);
    r.Given = r.GiveLog + r.GiveRec;
    r.GiveSpend = E === 'post' ? r.Given
      : E === 'open' ? givingSum(openNum, B, false, true) + recurringGiving(recurring, A, B, C, openNum + 1)
      : 0;

    // Deposits
    const dep = investments.map((_, k) => (E === 'pre' ? 0 : N(cellValue(m.deposits[k]))));
    dep.forEach((v, k) => (r[`Dep${k}`] = v));
    r.Contrib = sum(dep);
    r.ContribSpend = r.Contrib;
    r.RetDep = sum(dep.filter((_, k) => is(investments[k].type, 'Retirement')));

    // Account values: typed, else carried from last month (or the opening value)
    const carry = (k, opening) => (i === 0 || E === 'open' || prev.Phase === 'pre' ? N(opening) : N(prev[`Val${k}`]));
    accounts.forEach((a, k) => {
      const c = m.balances[k];
      r[`Val${k}`] = isBlank(a.name) || E === 'pre' ? null : typedConst(c) ? N(c.value) : carry(k, a.opening);
      r[`Typed${k}`] = typedConst(c) ? 1 : 0;
      r[`Miss${k}`] = E !== 'pre' && !isBlank(a.name) && !is(a.active, 'No') && is(a.required, 'Yes') && !typedConst(c) ? 1 : 0;
    });
    investments.forEach((inv, j) => {
      const k = 14 + j;
      const c = m.invValues[j];
      r[`Val${k}`] = isBlank(inv.name) || E === 'pre' ? null : typedConst(c) ? N(c.value) : carry(k, inv.opening) + dep[j];
      r[`Typed${k}`] = typedConst(c) ? 1 : 0;
    });
    r.MissingCount = sum(accounts.map((_, k) => r[`Miss${k}`]));
    r.Complete = E !== 'pre' && r.MissingCount === 0 && requiredCount > 0 ? 1 : 0;
    r.NegLiab = accounts.filter((a, k) => isAny(a.type, LIABILITY) && typedConst(m.balances[k]) && N(m.balances[k].value) < 0).length;

    // Family-paid
    r.FamAuto = isBlank(st.familyUsual) || D === 1 || E === 'pre' || A < (isBlank(st.familyStart) ? 0 : cmpValue(st.familyStart)) ? 0 : num(st.familyUsual);

    // Spending
    r.NetCash = r.Complete ? sum(accounts.map((a, k) => sign(a.type) * N(r[`Val${k}`]))) : null;
    if (E === 'open') { r.PrevNetCash = openingNetCash; r.PrevOK = 1; }
    else if (E === 'post') {
      const first = i === 0 || prev.Phase === 'pre';
      r.PrevNetCash = first ? openingNetCash : prev.NetCash;
      r.PrevOK = first ? 1 : prev.Complete;
    } else { r.PrevNetCash = null; r.PrevOK = 0; }

    // Cards and loans (FFOS 3.9): amount owed, the amount before, and rises in a row.
    r.Owed = r.Complete ? sum(accounts.map((a, k) => (isAny(a.type, LIABILITY) ? N(r[`Val${k}`]) : 0))) : null;
    r.PrevOwed = E === 'open' ? openingOwed
      : E === 'post' ? (i === 0 || prev.Phase === 'pre' ? openingOwed : prev.Owed)
      : null;
    r.Rise = r.Complete && r.PrevOK ? (round(r.Owed - r.PrevOwed, 2) > 0 ? 1 : 0) : 0;
    r.RiseStreak = r.Rise ? (i === 0 ? 1 : prev.RiseStreak + 1) : 0;
    r.HHSpend = r.Complete && r.PrevOK ? round(r.IncSpend - (r.NetCash - r.PrevNetCash) - r.ContribSpend - r.GiveSpend, 2) : null;
    const fam = cellValue(m.family);
    r.Family = isBlank(fam) ? r.FamAuto : N(fam);
    r.OneTime = N(cellValue(m.oneTime));
    r.TrueCost = r.HHSpend === null ? null : r.HHSpend + r.Family;
    r.NormalCost = r.TrueCost === null ? null : r.TrueCost - r.OneTime;
    r.InvValue = E === 'pre' || D === 1 ? null : sum(investments.map((_, j) => N(r[`Val${14 + j}`])));
    r.NetWorth = r.Complete ? r.NetCash + N(r.InvValue) : null;
    r.Saved = r.HHSpend === null ? null : r.IncSpend - r.HHSpend - r.GiveSpend;
    r.Full = E === 'post' && r.HHSpend !== null && B < today ? 1 : 0;
    r.Status = D === 1 ? null
      : E === 'pre' ? 'History only'
      : r.Complete ? (r.HHSpend === null ? '⚠ Previous month missing' : r.HHSpend < 0 ? '⚠ Check: spending < 0' : E === 'open' ? '✓ Done (partial)' : '✓ Done')
      : B >= today ? 'In progress' : '⚠ Needs balances';
    r.Year = YEAR(A);
    r.QStart = quarterStart(A);
    r.RowIfDone = r.Complete && r.HHSpend !== null && D === 0 ? r.row : 0;
    r.MissingPast = E !== 'pre' && B < today && !r.Complete ? 1 : 0;
    r.ChartNW = r.NetWorth === null ? '#N/A' : r.NetWorth;
    r.IncomeAvg = B < today && r.Income > 0 ? 1 : 0;
    months.push(r);
  }
  // Ranks count from the latest month backwards (1 = latest).
  let full = 0, incm = 0;
  for (let i = MONTH_COUNT - 1; i >= 0; i--) {
    const r = months[i];
    if (r.Full) full++;
    if (r.IncomeAvg) incm++;
    r.RankFull = r.Full ? full : 0;
    r.RankInc = r.IncomeAvg ? incm : 0;
  }

  // ---------- History: tithe by quarter ----------
  const ms = typeof st.titheStart === 'number' ? st.titheStart : 0;
  const msMonth = DATE(YEAR(ms), MONTH(ms), 1);
  const quarters = [];
  for (let q = 0; q < QUARTERS; q++) {
    const start = q === 0 ? quarterStart(ms) : EDATE(quarters[q - 1].start, 3);
    const prevF = q === 0 ? null : quarters[q - 1].balance;
    const carried = q === 0 ? N(st.openingTithe) : is(st.carryExtra, 'No') ? Math.max(0, prevF) : prevF;
    const due = round(num(st.titheRate) * sum(months.filter((r) => r.QStart === start).map((r) => r.TitheInc)), 2);
    const given = sum(months.filter((r) => r.QStart === start && r.MonthStart >= msMonth).map((r) => r.Given));
    const balance = round(carried + due - given, 2);
    const label = start > today ? null : `Q${Math.floor((MONTH(start) - 1) / 3) + 1} ${YEAR(start)}`;
    const status = label === null ? null : today < EDATE(start, 3) ? 'In progress' : balance <= 0 ? '✓ Settled' : 'Carried forward';
    quarters.push({ label, start, carried, due, given, balance, status });
  }
  const thisQ = quarters.find((q) => q.start === quarterStart(today));

  // ---------- Metrics (Calc!EA:EB) ----------
  const latestRow = Math.max(0, ...months.map((r) => r.RowIfDone));
  const latest = latestRow ? months[latestRow - FIRST_ROW] : null;
  const doneRows = months.map((r) => r.RowIfDone).sort((a, b) => b - a);
  const prevRow = doneRows[1] || 0;
  const thisYear = YEAR(today);
  const yearSum = (key, extra = () => true) => sum(months.filter((r) => r.Year === thisYear && extra(r) && typeof r[key] === 'number').map((r) => r[key]));
  const nAvg = num(st.avgMonths);
  const average = (key, rankKey) => {
    const cnt = months.filter((r) => r[rankKey] > 0).length;
    if (cnt === 0) return null;
    return sum(months.filter((r) => r[rankKey] > 0 && r[rankKey] <= nAvg).map((r) => r[key])) / minNums(st.avgMonths, cnt);
  };
  const avgHH = average('HHSpend', 'RankFull');
  const avgFam = average('Family', 'RankFull');
  const openYearThis = typeof st.openingDate === 'number' ? YEAR(st.openingDate) === thisYear : YEAR(0) === thisYear;

  // Account values now (Calc!EA40:EM59): 14 everyday accounts then 6 investments.
  const sixMonthsAgo = EDATE(today, -6);
  const rows = [];
  let rank = 0;
  const addRow = (acct, k, isInv) => {
    const named = !isBlank(acct.name);
    const active = named && !is(acct.active, 'No');
    const value = !named ? null : latestRow === 0 ? N(acct.opening) : N(latest[`Val${k}`]);
    if (active) rank++;
    const typedRecently = months.some((r) => r[`Typed${k}`] === 1 && r.MonthStart >= sixMonthsAgo);
    const staleType = isInv ? true : is(acct.type, 'Cash');
    rows.push({
      name: named ? acct.name : null,
      value,
      lifetime: isInv ? (named ? N(acct.contributed) + sum(months.filter((r) => r.Future === 0).map((r) => r[`Dep${k - 14}`])) : null) : null,
      stale: active && staleType && open < sixMonthsAgo && !typedRecently ? 1 : 0,
      rank: active ? rank : null,
      type: named ? acct.type : null,
      signed: !named ? null : isAny(acct.type, LIABILITY) ? -N(value) : N(value),
      estimate: !named ? null : is(acct.estimate, 'Yes') ? 'estimate' : null,
    });
  };
  accounts.forEach((a, k) => addRow(a, k, false));
  investments.forEach((inv, j) => addRow(inv, 14 + j, true));

  const byType = (list, offset, pred) => sum(list.map((a, k) => (pred(a.type) ? N(rows[offset + k].value) : 0)));
  const liquid = byType(accounts, 0, (t) => isAny(t, ASSET));
  const invNow = byType(investments, 14, (t) => is(t, 'Investment'));
  const retNow = byType(investments, 14, (t) => is(t, 'Retirement'));
  const liab = byType(accounts, 0, (t) => isAny(t, LIABILITY));
  const avgFamOut = avgFam === null ? N(st.fallbackFamily) : avgFam;
  const avgHHOut = avgHH === null ? N(st.fallbackHousehold) : avgHH;
  const openingNetWorth = openingNetCash + sum(investments.map((inv) => N(inv.opening)));
  const owedThisQ = thisQ ? thisQ.carried + thisQ.due : 0;

  const metrics = {
    'Latest done row (0 = none)': latestRow,
    'Net Worth': liquid + invNow + retNow - liab,
    'Liquid Cash': liquid,
    'Investments': invNow,
    'Retirement': retNow,
    'Retirement + Investments': invNow + retNow,
    'Card & Loan Balances': liab,
    'Average Monthly Income': average('Income', 'RankInc') ?? 0,
    'Average Monthly Household-Paid': avgHHOut,
    'Average Monthly Family-Paid': avgFamOut,
    'Average Monthly True Cost': avgHHOut + avgFamOut,
    'Income This Year': yearSum('Income'),
    'Giving This Year': yearSum('Given'),
    'Tithe-able Income This Year': yearSum('TitheInc'),
    'Contributions This Year': sum(investments.map((inv) => (!isBlank(inv.name) && openYearThis ? N(inv.contributedThisYear) : 0))) + yearSum('Contrib', (r) => r.Future === 0),
    'Retirement Contributions This Year': sum(investments.map((inv) => (!isBlank(inv.name) && openYearThis && is(inv.type, 'Retirement') ? N(inv.contributedThisYear) : 0))) + yearSum('RetDep', (r) => r.Future === 0),
    'Saved This Year': yearSum('Saved'),
    'Tithe Owed This Quarter': Math.max(0, owedThisQ),
    'Tithe Given This Quarter': thisQ ? thisQ.given : 0,
    'Tithe Balance Now': thisQ ? thisQ.balance : 0,
    'Average Monthly Giving': average('Given', 'RankInc') ?? 0,
    'Specific Account': 0,
    'Taxable Income Received This Year': yearSum('TaxableInc'),
    '…of Which No Tax Withheld': yearSum('UnwithheldInc'),
    'Tax-Deductible Giving This Year': yearSum('DeductGive'),
    'Opening net cash': openingNetCash,
    'Tithe owed this quarter (carried in + due)': owedThisQ,
    'Tithe given this quarter': thisQ ? thisQ.given : 0,
    'Liquid assets now': liquid,
    'Investments now': invNow,
    'Retirement now': retNow,
    'Liabilities now': liab,
    'As-of date': latestRow === 0 ? openNum : EOMONTH(latest.MonthStart, 0),
    'Latest done month label': latestRow === 0 ? `Opening ${mdyy(openNum)}` : monYear(latest.MonthStart),
    'Opening net worth': openingNetWorth,
    'Previous done row': prevRow,
    'Opening card & loan balances': openingOwed,
    'Card & loan balances at the previous done month (opening if none)': prevRow === 0 ? openingOwed : N(months[prevRow - FIRST_ROW].Owed),
    'Card & loan change vs previous': latestRow === 0 ? null
      : round(liab - (prevRow === 0 ? openingOwed : N(months[prevRow - FIRST_ROW].Owed)), 2),
    'Net worth change vs previous': latestRow === 0 ? null
      : N(latest.NetWorth) - (prevRow === 0 ? openingNetWorth : N(months[prevRow - FIRST_ROW].NetWorth)),
  };

  // ---------- Goals (Calc!EA70:EH89) ----------
  const metricByName = (name) => {
    if (typeof name !== 'string') return 0;
    const i = METRIC_NAMES.findIndex((n) => n.toLowerCase() === name.toLowerCase());
    return i < 0 ? 0 : metrics[METRIC_NAMES[i]];
  };
  const accountByName = (name) => {
    if (typeof name !== 'string' && typeof name !== 'number') return 0;
    const i = rows.findIndex((r) => r.name !== null && String(r.name).toLowerCase() === String(name).toLowerCase());
    return i < 0 ? 0 : rows[i].value;
  };
  const goalRows = goals.map((g, slot) => {
    const named = !isBlank(g.name);
    const base = isBlank(g.basedOn) ? 0 : is(g.basedOn, 'Specific Account') ? accountByName(g.account) : metricByName(g.basedOn);
    if (!named) return { slot, name: null, base, current: null, target: null, remaining: null, status: null, progress: null, note: null };
    const current = is(g.type, 'Tithe Balance') ? metrics['Tithe Given This Quarter']
      : is(g.currentFrom, 'Specific Account') ? accountByName(g.account) : metricByName(g.currentFrom);
    const target = is(g.type, 'Percentage') ? (base * N(g.target)) / 100
      : is(g.type, 'Multiple') ? base * N(g.target)
      : is(g.type, 'Tithe Balance') ? metrics['Tithe Owed This Quarter']
      : N(g.target);
    const payDown = is(g.type, 'Pay Down To');
    const remaining = payDown ? Math.max(0, current - target) : Math.max(0, target - current);
    const status = payDown ? (current <= target ? 'Met' : 'Above target') : current >= target - 0.005 ? 'Met' : 'Remaining';
    const progress = payDown ? (current <= 0 ? 1 : Math.min(1, target / current)) : target <= 0 ? 1 : current / target;
    let note = null;
    if (remaining > 0) {
      if (is(g.type, 'Save By Date')) {
        note = isBlank(g.targetDate) ? 'Set a target date'
          : `Need ${dollars(remaining / Math.max(1, (YEAR(g.targetDate) - YEAR(today)) * 12 + MONTH(g.targetDate) - MONTH(today) + 1))}/mo to ${monYear(g.targetDate)}`;
      } else if (is(g.type, 'Annual Contribution')) {
        note = `Need ${dollars(remaining / Math.max(1, 12 - MONTH(today) + 1))}/mo rest of year`;
      } else if (is(g.type, 'Tithe Balance')) note = 'This quarter';
    }
    return { slot, name: g.name, type: g.type, base, current, target, remaining, status, progress, note };
  });

  // ---------- Health checks (Calc!EA95:EA102) ----------
  const missingPast = months.filter((r) => r.MissingPast);
  const estimates = accounts.filter((a) => is(a.estimate, 'Yes')).length + investments.filter((inv) => is(inv.estimate, 'Yes')).length;
  const negMonths = months.filter((r) => typeof r.HHSpend === 'number' && r.HHSpend < 0).length;
  const notSure = yearSum('NotSureInc');
  const stale = sum(rows.map((r) => r.stale));
  const givingIncomplete = data.giving.filter((g) => isBlank(g.date) !== isBlank(g.amount)).length;
  const fullCount = months.filter((r) => r.RankFull > 0).length;
  const health = [
    requiredCount === 0 ? '⚠ No account is marked Required monthly (Settings → Everyday accounts)'
      : missingPast.length === 0 ? '✓ Every finished month has its required balances'
      : `⚠ Enter required balances for ${monthYear(missingPast[0].MonthStart)}${missingPast.length > 1 ? ` (+${missingPast.length - 1} more)` : ''}`,
    estimates === 0 ? '✓ No estimated opening values' : `ℹ ${estimates} opening values are estimates (Settings) — replace when you can`,
    negMonths === 0 ? '✓ No negative spending months' : `⚠ ${negMonths} month(s) with negative spending — usually missing income or a wrong balance`,
    sum(months.map((r) => r.NegLiab)) === 0 ? '✓ Card/loan balances are positive amounts' : '⚠ A card or loan balance is negative — enter the amount owed as a positive number',
    notSure === 0 ? '✓ No income waiting on a tithe decision' : `ℹ ${dollars(notSure)} income this year is marked Not Sure (not counted for tithe yet)`,
    stale === 0 ? '✓ Cash and investment values are fresh enough' : `ℹ ${stale} cash/investment value(s) not updated in 6 months`,
    givingIncomplete === 0 ? '✓ Giving log rows are complete' : '⚠ A Giving row has a date without an amount (or the reverse)',
    fullCount === 0 ? 'ℹ Averages use your fallback numbers until the first full month after opening is done'
      : `✓ Averages use ${minNums(st.avgMonths, fullCount)} real month(s)`,
  ];
  if (data.hasOwedTrend) {
    health.push((latestRow === 0 ? 0 : N(latest.RiseStreak)) >= 3
      ? '⚠ Card and loan balances have risen 3 months in a row — check they\'re being paid in full'
      : '✓ Card and loan balances aren\'t rising');
  }

  // ---------- Dashboard (the numbers it shows) ----------
  const at = (key) => (latest ? latest[key] : null);
  const dashboard = {
    asOf: metrics['As-of date'],
    latestLabel: metrics['Latest done month label'],
    netWorth: metrics['Net Worth'],
    netWorthChange: metrics['Net worth change vs previous'],
    changeSinceOpening: prevRow === 0,
    liquidCash: metrics['Liquid Cash'],
    monthsOfLiving: metrics['Average Monthly True Cost'] <= 0 ? null : metrics['Liquid Cash'] / metrics['Average Monthly True Cost'],
    investmentsAndRetirement: metrics['Retirement + Investments'],
    contributedThisYear: metrics['Contributions This Year'],
    owed: metrics['Card & Loan Balances'],
    flow: {
      latest: { income: at('Income'), household: at('HHSpend'), family: at('Family'), trueCost: at('TrueCost'), oneTime: at('OneTime'), given: at('Given'), invested: at('Contrib'), saved: at('Saved') },
      average: { income: metrics['Average Monthly Income'], household: metrics['Average Monthly Household-Paid'], family: metrics['Average Monthly Family-Paid'], trueCost: metrics['Average Monthly True Cost'], given: metrics['Average Monthly Giving'] },
      thisYear: { income: metrics['Income This Year'], household: yearSum('HHSpend'), family: yearSum('Family'), trueCost: yearSum('TrueCost'), oneTime: yearSum('OneTime'), given: metrics['Giving This Year'], invested: metrics['Contributions This Year'], saved: metrics['Saved This Year'] },
    },
    tax: data.hasTax ? { received: metrics['Taxable Income Received This Year'], unwithheld: metrics['…of Which No Tax Withheld'], deductible: metrics['Tax-Deductible Giving This Year'] } : null,
    tithe: { balance: metrics['Tithe Balance Now'], givenThisQuarter: metrics['Tithe Given This Quarter'], owedThisQuarter: metrics['Tithe Owed This Quarter'] },
    goals: goalRows.filter((g) => g.name !== null),
    accounts: rows.filter((r) => r.rank !== null).sort((a, b) => a.rank - b.rank),
    health,
  };

  return { today, months, quarters, metrics, accounts: rows, goals: goalRows, health, dashboard };
}
