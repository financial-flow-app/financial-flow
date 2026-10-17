// Writes each synthetic scenario into the template, for Excel to calculate (scripts/excel-parity.ps1).
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { SCENARIOS } from '../test/scenarios/index.mjs';
import { buildScenario } from '../test/scenarios/build.mjs';

const template = readFileSync(new URL('../fixtures/ffos-template.xlsx', import.meta.url));
const out = new URL('../test-output/parity/', import.meta.url);
mkdirSync(out, { recursive: true });
for (const sc of SCENARIOS) {
  writeFileSync(new URL(`${sc.name}.xlsx`, out), await buildScenario(template, sc));
  console.log(`wrote test-output/parity/${sc.name}.xlsx`);
}
