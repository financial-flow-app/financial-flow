// Synthetic parity scenarios. Names, amounts and organisations are invented.
// Each one is written into the template, recalculated by Excel (scripts/excel-parity.ps1),
// and the app's engine must give the same numbers.

import { d, account, investment, source, recurringGift, goal, month, givingRows } from './helpers.mjs';

const settings = (...parts) => Object.assign({}, ...parts);

export const SCENARIOS = [
  {
    name: 'opening-month',
    about: 'Opening mid-April: typed pay, a card donation, an IRA deposit, a family-paid cost, a one-time cost.',
    settings: settings(
      { B5: d('2026-04-15'), B6: d('2026-04-01'), B7: 0.1, B8: 0, B9: 'Yes', B10: 3, B11: 2500, B12: 100, B13: 250, B14: 300, B15: d('2026-05-01') },
      account(0, 'Checking A', 'Checking', 5000, { required: 'Yes' }),
      account(1, 'Savings A', 'Savings', 10000),
      account(2, 'Card A', 'Card', 800, { required: 'Yes' }),
      account(3, 'Wallet', 'Cash', 200, { estimate: 'Yes' }),
      investment(0, 'IRA', 'Retirement', 20000, { contributed: 15000, thisYear: 1000 }),
      investment(1, 'Brokerage', 'Investment', 5000, { contributed: 4000 }),
      source(0, 'Salary', 'Yes', { usual: 1500, frequency: 'Weekly', day: 'Friday', from: '2026-01-01' }),
      source(1, 'Bonus', 'Yes'),
      source(2, 'Side work', 'Not Sure', { usual: 400, frequency: 'Monthly', day: 31 }),
      source(3, 'Rental', 'No', { usual: 900, frequency: 'Monthly', day: 1, until: '2026-08-31' }),
      recurringGift(0, 'Monthly pledge', 50, { day: 10, from: '2026-01-01' }),
      recurringGift(1, 'Family pledge', 25, { day: 20, household: 'No' }),
    ),
    monthEnd: [
      ...month('2026-04', { bal: { 0: 5400, 2: 950 }, inc: { 1: 1200 }, dep: { 0: 500 }, family: 450, oneTime: 300 }),
      ...month('2026-05', { bal: { 0: 5100, 1: 10100, 2: 700 }, inv: { 0: 21000 }, dep: { 1: 200 } }),
      ...month('2026-06', { bal: { 0: 5300, 2: 600 }, inc: { 0: 6000 } }),
    ],
    giving: givingRows([
      { date: '2026-04-10', amount: 100, recipient: 'Charity A', paidFrom: 'Card A' },
      { date: '2026-04-20', amount: 180, recipient: 'Charity B', paidFrom: 'Card A' },
      { date: '2026-05-03', amount: 60, recipient: 'Fund C', household: 'No' },
      { date: '2026-06-15', amount: 36, recipient: 'Charity A' },
    ]),
  },
  {
    name: 'blank-unchanged',
    about: 'Optional balances left blank carry forward; untyped investments carry forward plus deposits.',
    settings: settings(
      { B5: d('2026-03-31'), B6: d('2026-03-01'), B10: 3 },
      account(0, 'Checking A', 'Checking', 3000, { required: 'Yes' }),
      account(1, 'Savings A', 'Savings', 8000),
      account(2, 'Cash box', 'Cash', 150),
      account(3, 'Old account', 'Checking', 400, { active: 'No' }),
      investment(0, 'Index fund', 'Investment', 12000, { contributed: 9000 }),
      source(0, 'Pay', 'Yes', { usual: 2000, frequency: 'Monthly', day: 15, from: '2026-03-01' }),
    ),
    monthEnd: [
      ...month('2026-04', { bal: { 0: 3200 }, dep: { 0: 300 } }),
      ...month('2026-05', { bal: { 0: 3100, 1: 8100 }, dep: { 0: 300 } }),
      ...month('2026-06', { bal: { 0: 3400 } }),
      ...month('2026-07', { bal: { 0: 3350, 2: 90 }, inv: { 0: 13000 }, dep: { 0: 250 } }),
      ...month('2026-08', { bal: { 0: 9000 }, dep: { 0: -100 } }), // jump with no income: negative spending
    ],
    giving: givingRows([{ date: '2026-05-12', amount: 200, recipient: 'Charity A' }]),
  },
  {
    name: 'missing-balance',
    about: 'A required balance missing in June; negative card balance; estimates; an incomplete Giving row; lowercase yes.',
    settings: settings(
      { B5: d('2026-03-31'), B6: d('2026-03-01'), B10: 3 },
      account(0, 'Checking A', 'Checking', 4000, { required: 'Yes' }),
      account(1, 'Card A', 'Card', 500, { required: 'yes', estimate: 'Yes' }),
      account(2, 'Car loan', 'Loan', 9000),
      investment(0, 'Pension', 'Retirement', 30000, { estimate: 'Yes' }),
      source(0, 'Wages', 'yes', { usual: 800, frequency: 'Weekly', day: 'Monday', from: '2026-04-08', until: '2026-08-19' }),
      source(1, 'Gifts', 'Not Sure'),
    ),
    monthEnd: [
      ...month('2026-04', { bal: { 0: 4300, 1: 450, 2: 8800 }, inc: { 1: 75 } }),
      ...month('2026-05', { bal: { 0: 4500, 1: 400 } }),
      ...month('2026-06', { bal: { 0: 4700 } }),
      ...month('2026-07', { bal: { 0: 4600, 1: 380 } }),
      ...month('2026-08', { bal: { 0: 4900, 1: -20 } }),
      ...month('2026-10', { bal: { 0: 5000 } }),
    ],
    giving: givingRows([
      { date: '2026-04-02', amount: 40, recipient: 'Charity A' },
      { date: '2026-05-09', recipient: 'Charity B' },
    ]),
  },
  {
    name: 'three-full-months',
    about: 'Opening end of May; June, July and August complete; averages over the last 2 full months.',
    settings: settings(
      { B5: d('2026-05-31'), B6: d('2026-03-01'), B10: 2, B11: 3000, B12: 250, B14: 200 },
      account(0, 'Checking A', 'Checking', 2500, { required: 'Yes' }),
      account(1, 'Card A', 'Card', 300, { required: 'Yes' }),
      investment(0, '401k', 'Retirement', 40000, { contributed: 30000, thisYear: 2000 }),
      source(0, 'Job', 'Yes', { usual: 1100, frequency: 'Weekly', day: 'Monday', from: '2026-03-01' }),
      source(1, 'Tutoring', 'Not Sure', { usual: 300, frequency: 'Monthly', day: 15 }),
      recurringGift(0, 'Pledge', 120, { day: 1 }),
    ),
    monthEnd: [
      ...month('2026-03', { inc: { 1: 0 } }),
      ...month('2026-06', { bal: { 0: 2900, 1: 450 }, dep: { 0: 400 }, oneTime: 500 }),
      ...month('2026-07', { bal: { 0: 3100, 1: 380 }, dep: { 0: 400 }, family: 0 }),
      ...month('2026-08', { bal: { 0: 2800, 1: 610 }, dep: { 0: 400 }, inc: { 0: 4000 } }),
    ],
    giving: givingRows([
      { date: '2026-03-20', amount: 90, recipient: 'Charity A' },
      { date: '2026-06-05', amount: 150, recipient: 'Charity B' },
      { date: '2026-08-25', amount: 75.5, recipient: 'Charity A', household: 'No' },
    ]),
  },
  {
    name: 'tithe-quarters-no-carry',
    about: 'Tracking starts mid-February; giving on quarter edges; extra giving does not carry forward.',
    settings: settings(
      { B5: d('2026-02-28'), B6: d('2026-02-15'), B7: 0.1, B8: 120, B9: 'No', B10: 3 },
      account(0, 'Checking A', 'Checking', 6000, { required: 'Yes' }),
      source(0, 'Salary', 'Yes', { usual: 4000, frequency: 'Monthly', day: 25 }),
      source(1, 'Interest', 'No', { usual: 12.34, frequency: 'Monthly' }),
    ),
    monthEnd: [
      ...month('2026-03', { bal: { 0: 6500 } }),
      ...month('2026-04', { bal: { 0: 6900 } }),
      ...month('2026-05', { bal: { 0: 7000 } }),
      ...month('2026-06', { bal: { 0: 7600 } }),
      ...month('2026-07', { bal: { 0: 7700 } }),
      ...month('2026-08', { bal: { 0: 8300 } }),
      ...month('2026-09', { bal: { 0: 8800 } }),
    ],
    giving: givingRows([
      { date: '2026-03-31', amount: 500, recipient: 'Charity A' },
      { date: '2026-04-01', amount: 900, recipient: 'Charity B' },
      { date: '2026-06-30', amount: 1000, recipient: 'Charity A' },
      { date: '2026-07-01', amount: 333.33, recipient: 'Fund C' },
      { date: '2026-09-30', amount: 50, recipient: 'Fund C', household: 'No' },
      { date: '2026-10-05', amount: 120, recipient: 'Charity A' },
    ]),
  },
  {
    name: 'goals',
    about: 'One goal of every type, a specific-account goal and a metric-based percentage.',
    settings: settings(
      { B5: d('2026-03-31'), B6: d('2026-03-01'), B10: 3 },
      account(0, 'Checking A', 'Checking', 7000, { required: 'Yes' }),
      account(1, 'Emergency fund', 'Savings', 9000),
      account(2, 'Card A', 'Card', 1200, { required: 'Yes' }),
      investment(0, 'Roth IRA', 'Retirement', 15000, { contributed: 12000, thisYear: 500 }),
      investment(1, 'Taxable', 'Investment', 8000, { contributed: 7000 }),
      source(0, 'Salary', 'Yes', { usual: 5200, frequency: 'Monthly', day: 1 }),
      goal(0, 'Reach net worth', 'Fixed Amount', 50000, { currentFrom: 'Net Worth' }),
      goal(1, 'Save 20% of income', 'Percentage', 20, { basedOn: 'Income This Year', currentFrom: 'Saved This Year' }),
      goal(2, 'Six months cushion', 'Multiple', 6, { basedOn: 'Average Monthly True Cost', currentFrom: 'Liquid Cash' }),
      goal(3, 'Retirement deposits', 'Annual Contribution', 7000, { currentFrom: 'Retirement Contributions This Year' }),
      goal(4, 'Card under 500', 'Pay Down To', 500, { currentFrom: 'Card & Loan Balances' }),
      goal(5, 'Emergency fund 15k', 'Save By Date', 15000, { currentFrom: 'Specific Account', account: 'Emergency fund', targetDate: '2027-06-30' }),
      goal(6, 'Trip fund', 'Save By Date', 20000, { currentFrom: 'Specific Account', account: 'Checking A' }),
      goal(7, 'Tithe this quarter', 'Tithe Balance', null),
      goal(8, 'Taxable at 2x cost', 'Percentage', 150, { basedOn: 'Specific Account', account: 'Taxable', currentFrom: 'Investments' }),
    ),
    monthEnd: [
      ...month('2026-04', { bal: { 0: 7600, 2: 1100 }, dep: { 0: 400 } }),
      ...month('2026-05', { bal: { 0: 8200, 1: 9500, 2: 900 }, dep: { 0: 400 } }),
      ...month('2026-06', { bal: { 0: 8100, 2: 950 }, dep: { 0: 400, 1: 250 } }),
      ...month('2026-07', { bal: { 0: 8900, 1: 10200, 2: 700 }, inv: { 1: 8600 } }),
      ...month('2026-08', { bal: { 0: 9400, 2: 650 }, dep: { 0: 600 } }),
      ...month('2026-09', { bal: { 0: 9900, 1: 11000, 2: 600 } }),
    ],
    giving: givingRows([
      { date: '2026-07-15', amount: 400, recipient: 'Charity A' },
      { date: '2026-10-02', amount: 250, recipient: 'Charity B' },
    ]),
  },
  {
    name: 'tax',
    about: 'Tax marks: withheld, not withheld, not taxable, not sure; deductible donations, recurring and family-funded.',
    settings: settings(
      { B5: d('2026-03-31'), B6: d('2026-03-01'), B10: 3 },
      account(0, 'Checking A', 'Checking', 5000, { required: 'Yes' }),
      source(0, 'Salary', 'Yes', { usual: 3000, frequency: 'Monthly', day: 1, taxable: 'Yes', withheld: 'Yes' }),
      source(1, 'Side work', 'Yes', { usual: 500, frequency: 'Monthly', day: 20, taxable: 'Yes', withheld: 'No' }),
      source(2, 'Gifts', 'No', { taxable: 'No' }),
      source(3, 'Consulting', 'Not Sure', { usual: 250, frequency: 'Weekly', day: 'Wednesday', from: '2026-06-01', taxable: 'yes', withheld: 'no' }),
      source(4, 'Refund', 'No', { taxable: 'Not Sure', withheld: 'No' }),
      recurringGift(0, 'Pledge', 40, { day: 5, deductible: 'Yes' }),
      recurringGift(1, 'Family pledge', 30, { day: 12, household: 'No', deductible: 'Yes' }),
      recurringGift(2, 'Old pledge', 15, { day: 28, until: '2026-06-30', deductible: 'Yes' }),
    ),
    monthEnd: [
      ...month('2026-04', { bal: { 0: 5600 }, inc: { 2: 200, 4: 80 } }),
      ...month('2026-05', { bal: { 0: 6100 }, inc: { 1: 650 } }),
      ...month('2026-06', { bal: { 0: 6400 } }),
      ...month('2026-07', { bal: { 0: 7000 }, inc: { 0: 3100 } }),
      ...month('2026-08', { bal: { 0: 7300 } }),
      ...month('2026-09', { bal: { 0: 7900 } }),
    ],
    giving: givingRows([
      { date: '2026-04-08', amount: 100, recipient: 'Charity A', deductible: 'Yes' },
      { date: '2026-05-14', amount: 54, recipient: 'Charity B', deductible: 'No' },
      { date: '2026-06-02', amount: 18, recipient: 'Fund C' },
      { date: '2026-07-21', amount: 250, recipient: 'Charity A', household: 'No', deductible: 'Yes' },
      { date: '2026-09-30', amount: 75.25, recipient: 'Charity B', deductible: 'yes' },
    ]),
  },
  {
    name: 'cards-rising',
    about: 'Card and loan balances: a fall, then three rises in a row (health warning); the change line each month.',
    settings: settings(
      { B5: d('2026-03-31'), B6: d('2026-03-01'), B10: 3 },
      account(0, 'Checking A', 'Checking', 4000, { required: 'Yes' }),
      account(1, 'Card A', 'Card', 300, { required: 'Yes' }),
      account(2, 'Card B', 'card', 120, { required: 'Yes' }),
      account(3, 'Car loan', 'Loan', 5000),
      source(0, 'Salary', 'Yes', { usual: 3200, frequency: 'Monthly', day: 1 }),
    ),
    monthEnd: [
      ...month('2026-04', { bal: { 0: 4300, 1: 350, 2: 120 } }),
      ...month('2026-05', { bal: { 0: 4900, 1: 280, 2: 100, 3: 4800 } }),
      ...month('2026-06', { bal: { 0: 5100, 1: 400, 2: 100 } }),
      ...month('2026-07', { bal: { 0: 5000, 1: 450, 2: 110.5 } }),
      ...month('2026-08', { bal: { 0: 4700, 1: 520, 2: 110.5 } }),
    ],
    giving: givingRows([]),
  },
  {
    name: 'first-month',
    about: 'Only one complete month, so the card line compares with the opening balances.',
    settings: settings(
      { B5: d('2026-08-31'), B6: d('2026-03-01'), B10: 3 },
      account(0, 'Checking A', 'Checking', 2000, { required: 'Yes' }),
      account(1, 'Card A', 'Card', 400, { required: 'Yes' }),
    ),
    monthEnd: [...month('2026-09', { bal: { 0: 2300, 1: 250 } })],
    giving: givingRows([]),
  },
  {
    name: 'owed-same',
    about: 'The amount owed is unchanged since the opening: "Same as at opening".',
    settings: settings(
      { B5: d('2026-08-31'), B6: d('2026-03-01'), B10: 3 },
      account(0, 'Checking A', 'Checking', 2000, { required: 'Yes' }),
      account(1, 'Card A', 'Card', 400, { required: 'Yes' }),
    ),
    monthEnd: [...month('2026-09', { bal: { 0: 2100, 1: 400 } })],
    giving: givingRows([]),
  },
  {
    name: 'no-complete-month',
    about: 'Cards owed at opening and no month complete yet: "As of opening".',
    settings: settings(
      { B5: d('2026-08-31'), B6: d('2026-03-01'), B10: 3 },
      account(0, 'Checking A', 'Checking', 2000, { required: 'Yes' }),
      account(1, 'Card A', 'Card', 400, { required: 'Yes' }),
    ),
    monthEnd: [...month('2026-09', { bal: { 0: 2300 } })],
    giving: givingRows([]),
  },
];

// The same as the no-carry case, but extra giving carries forward as a credit.
SCENARIOS.push({
  ...SCENARIOS.find((s) => s.name === 'tithe-quarters-no-carry'),
  name: 'tithe-quarters-carry',
  about: 'As tithe-quarters-no-carry, but extra giving carries forward.',
});
SCENARIOS.at(-1).settings = { ...SCENARIOS.at(-1).settings, B9: 'Yes' };

// The goals scenario with goals edited through the app (spec section 11): the app's own export
// writes them into Settings, and Excel must compute the same goal results from what it wrote.
SCENARIOS.push({
  ...SCENARIOS.find((s) => s.name === 'goals'),
  name: 'goals-edited',
  about: 'As goals, then edited in the app: one added, two changed (one to another type), one deleted.',
  edit: (data, entry) => {
    entry.addGoal(data, { name: 'Checking cushion', type: 'Save By Date', target: 12000, currentFrom: 'Specific Account', account: 'Checking A', targetDate: d('2027-01-31') });
    entry.updateGoal(data, 1, { ...data.goals[1], target: 30, basedOn: 'Average Monthly Income' });
    entry.updateGoal(data, 2, { ...data.goals[2], type: 'Fixed Amount', target: 9000 });
    entry.deleteGoal(data, 4);
  },
});

// Accounts, investments and income sources set up through the app (spec section 13): the app's own export
// writes them into Settings and Month-End, and Excel must compute the same results from what it wrote.
SCENARIOS.push({
  ...SCENARIOS.find((s) => s.name === 'goals'),
  name: 'setup-edited',
  about: 'As goals, then in the app in October: a card and an investment added, a side job added, the salary ended, '
    + 'the emergency fund closed (its money moved to checking), and an investment renamed (a goal follows it).',
  edit: (data, entry) => {
    const oct = data.months.findIndex((m) => m.ym === '2026-10');
    entry.addAccount(data, { name: 'Card B', type: 'Card' });
    entry.setEntry(data, oct, 'balances', 3, 300);
    entry.addInvestment(data, { name: 'Brokerage', type: 'Investment' });
    entry.setEntry(data, oct, 'invValues', 2, 2000);
    entry.setEntry(data, oct, 'deposits', 2, 2000);
    entry.addSource(data, oct, { name: 'Side job', tithe: 'Yes', usual: 400, frequency: 'Monthly', day: 15 });
    entry.endSource(data, oct, 0);
    entry.closeAccount(data, oct, 'accounts', 1);
    entry.setEntry(data, oct, 'balances', 0, 21500);
    entry.setEntry(data, oct, 'balances', 2, 500);
    entry.renameSetup(data, 'investments', 1, 'Taxable account');
  },
});
