// Writes the test scenario's exports to test-output/ for the Excel check.
import { readFileSync, writeFileSync, mkdirSync, copyFileSync } from 'node:fs';
import { runScenario } from '../test/scenario.mjs';

const out = new URL('../test-output/', import.meta.url);
mkdirSync(out, { recursive: true });
for (const name of ['ffos-template', 'ffos-template-excel']) {
  const src = new URL(`../fixtures/${name}.xlsx`, import.meta.url);
  const { pass1, pass2 } = await runScenario(readFileSync(src));
  copyFileSync(src, new URL(`${name}.base.xlsx`, out));
  writeFileSync(new URL(`${name}.pass1.xlsx`, out), pass1);
  writeFileSync(new URL(`${name}.pass2.xlsx`, out), pass2);
  console.log(`wrote ${name}.{base,pass1,pass2}.xlsx`);
}
