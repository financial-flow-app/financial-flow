// Draws the app icons as PNG files (no image libraries): a blue tile with three rising bars.
// Usage: node scripts/make-icons.mjs   -> icons/icon-192.png, icons/icon-512.png, icons/maskable-512.png
import { writeFileSync, mkdirSync } from 'node:fs';
import { deflateSync } from 'node:zlib';
import { crc32 } from '../src/zip.js';

const BLUE = [37, 99, 235];
const WHITE = [255, 255, 255];

/** Coverage (0..1) of pixel (x, y) for a size×size icon, sampled 4×4 for smooth edges. */
function draw(size, maskable) {
  const px = new Uint8Array(size * size * 4);
  const S = 4;
  // Shapes in a 0..1 square. Maskable icons fill the square; content stays in the centre 80%.
  const pad = maskable ? 0 : 0.04;
  const radius = maskable ? 0 : 0.22;
  const k = maskable ? 0.72 : 1; // shrink the bars into the maskable safe zone
  const bars = [[0.22, 0.58], [0.42, 0.42], [0.62, 0.26]].map(([x, top]) => ({
    x0: 0.5 + (x - 0.5) * k, x1: 0.5 + (x + 0.16 - 0.5) * k, y0: 0.5 + (top - 0.5) * k, y1: 0.5 + (0.78 - 0.5) * k,
  }));
  const inTile = (u, v) => {
    const a = pad, b = 1 - pad;
    if (u < a || u > b || v < a || v > b) return false;
    const r = radius;
    const cx = Math.min(Math.max(u, a + r), b - r), cy = Math.min(Math.max(v, a + r), b - r);
    return (u - cx) ** 2 + (v - cy) ** 2 <= r * r + 1e-12;
  };
  const inBar = (u, v) => bars.some((q) => u >= q.x0 && u <= q.x1 && v >= q.y0 && v <= q.y1);
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let tile = 0, bar = 0;
      for (let sy = 0; sy < S; sy++) {
        for (let sx = 0; sx < S; sx++) {
          const u = (x + (sx + 0.5) / S) / size, v = (y + (sy + 0.5) / S) / size;
          if (inTile(u, v)) { tile++; if (inBar(u, v)) bar++; }
        }
      }
      const n = S * S, i = (y * size + x) * 4;
      const t = tile / n, w = tile ? bar / tile : 0;
      for (let c = 0; c < 3; c++) px[i + c] = Math.round(BLUE[c] * (1 - w) + WHITE[c] * w);
      px[i + 3] = Math.round(255 * t);
    }
  }
  return px;
}

function png(size, rgba) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; // 8-bit RGBA
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) Buffer.from(rgba.buffer, y * size * 4, size * 4).copy(raw, y * (size * 4 + 1) + 1);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]);
}

const dir = new URL('../icons/', import.meta.url);
mkdirSync(dir, { recursive: true });
for (const [name, size, maskable] of [['icon-192.png', 192, false], ['icon-512.png', 512, false], ['maskable-512.png', 512, true]]) {
  writeFileSync(new URL(name, dir), png(size, draw(size, maskable)));
  console.log(`wrote icons/${name}`);
}
