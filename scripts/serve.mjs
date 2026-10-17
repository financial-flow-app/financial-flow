// Local test server (never part of the app): serves this folder on http://127.0.0.1:5180.
// It also serves /__dev/fixture.js, a synthetic scenario workbook as base64, so a browser test
// can import a file without a file picker. Usage: node scripts/serve.mjs [port]

import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { extname, join, normalize, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { SCENARIOS } from '../test/scenarios/index.mjs';
import { buildScenario } from '../test/scenarios/build.mjs';

const root = fileURLToPath(new URL('..', import.meta.url));
const port = Number(process.argv[2] || 5180);
const types = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.svg': 'image/svg+xml', '.png': 'image/png',
  '.webmanifest': 'application/manifest+json', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

const template = readFileSync(new URL('../fixtures/ffos-template.xlsx', import.meta.url));

createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost');
  if (url.pathname === '/__dev/fixture.js') {
    const sc = SCENARIOS.find((s) => s.name === (url.searchParams.get('name') || 'opening-month'));
    const b64 = Buffer.from(await buildScenario(template, sc)).toString('base64');
    res.writeHead(200, { 'content-type': types['.js'], 'cache-control': 'no-store' });
    return res.end(`window.__fixture = ${JSON.stringify({ name: `${sc.name}.xlsx`, b64 })};`);
  }
  const path = normalize(join(root, decodeURIComponent(url.pathname === '/' ? '/index.html' : url.pathname)));
  if (!path.startsWith(root.endsWith(sep) ? root : root + sep)) { res.writeHead(403); return res.end(); }
  try {
    if (!(await stat(path)).isFile()) throw new Error('not a file');
    res.writeHead(200, { 'content-type': types[extname(path)] || 'application/octet-stream', 'cache-control': 'no-store' });
    res.end(await readFile(path));
  } catch {
    res.writeHead(404); res.end('not found');
  }
}).listen(port, '127.0.0.1', () => console.log(`http://127.0.0.1:${port}`));
