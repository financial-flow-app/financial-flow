// Minimal zip reader/writer for .xlsx files.
// Unchanged entries are copied as their original bytes (headers and compressed data),
// so only the parts we edit are rewritten. Uses the platform's own deflate
// (CompressionStream / DecompressionStream), so there are no dependencies.

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_DESCRIPTOR = 0x08074b50;

const utf8 = new TextDecoder('utf-8');

function view(bytes) {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function findEocd(bytes) {
  const dv = view(bytes);
  const min = Math.max(0, bytes.length - 22 - 0xffff);
  for (let i = bytes.length - 22; i >= min; i--) {
    if (dv.getUint32(i, true) === SIG_EOCD) return i;
  }
  throw new Error('Not a zip file (no end-of-central-directory record).');
}

/** Parse a zip. Every entry keeps its raw local header, data, descriptor and central header. */
export function readZip(bytes) {
  // A plain Uint8Array view: Node's Buffer.slice() shares memory instead of copying,
  // and writeZip relies on slice() copying the headers it patches.
  bytes = ArrayBuffer.isView(bytes)
    ? new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    : new Uint8Array(bytes);
  const dv = view(bytes);
  const eocd = findEocd(bytes);
  const count = dv.getUint16(eocd + 10, true);
  const cdSize = dv.getUint32(eocd + 12, true);
  const cdOffset = dv.getUint32(eocd + 16, true);
  if (count === 0xffff || cdOffset === 0xffffffff) throw new Error('Zip64 files are not supported.');
  const entries = [];
  let p = cdOffset;
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== SIG_CENTRAL) throw new Error('Bad central directory.');
    const flags = dv.getUint16(p + 8, true);
    const method = dv.getUint16(p + 10, true);
    const crc = dv.getUint32(p + 16, true);
    const compSize = dv.getUint32(p + 20, true);
    const size = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const offset = dv.getUint32(p + 42, true);
    if (compSize === 0xffffffff || size === 0xffffffff || offset === 0xffffffff) {
      throw new Error('Zip64 entries are not supported.');
    }
    if (flags & 1) throw new Error('Encrypted zip entries are not supported.');
    const centralLen = 46 + nameLen + extraLen + commentLen;
    const central = bytes.subarray(p, p + centralLen);
    const name = utf8.decode(bytes.subarray(p + 46, p + 46 + nameLen));

    if (dv.getUint32(offset, true) !== SIG_LOCAL) throw new Error(`Bad local header for ${name}.`);
    const lNameLen = dv.getUint16(offset + 26, true);
    const lExtraLen = dv.getUint16(offset + 28, true);
    const dataStart = offset + 30 + lNameLen + lExtraLen;
    const local = bytes.subarray(offset, dataStart);
    const data = bytes.subarray(dataStart, dataStart + compSize);
    let descriptor = null;
    if (flags & 8) {
      const d = dataStart + compSize;
      const len = dv.getUint32(d, true) === SIG_DESCRIPTOR ? 16 : 12;
      descriptor = bytes.subarray(d, d + len);
    }
    entries.push({ name, flags, method, crc, compSize, size, local, data, descriptor, central });
    p += centralLen;
  }
  if (p !== cdOffset + cdSize) throw new Error('Central directory size mismatch.');
  const comment = bytes.subarray(eocd + 22, eocd + 22 + dv.getUint16(eocd + 20, true));
  return { entries, comment };
}

async function pipe(data, stream) {
  const out = await new Response(new Blob([data]).stream().pipeThrough(stream)).arrayBuffer();
  return new Uint8Array(out);
}

/** Uncompressed bytes of an entry. */
export async function entryBytes(entry) {
  if (entry.method === 0) return entry.data;
  if (entry.method !== 8) throw new Error(`Unsupported compression method ${entry.method} in ${entry.name}.`);
  const out = await pipe(entry.data, new DecompressionStream('deflate-raw'));
  if (out.length !== entry.size || crc32(out) !== entry.crc) {
    throw new Error(`Corrupt zip entry ${entry.name}.`);
  }
  return out;
}

export async function entryText(entry) {
  return utf8.decode(await entryBytes(entry));
}

let crcTable;
export function crc32(bytes) {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = crcTable[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** Headers for an entry that doesn't exist yet (deflate, sizes filled in on write, date 1980-01-01). */
function newEntry(name) {
  const nameBytes = new TextEncoder().encode(name);
  const local = new Uint8Array(30 + nameBytes.length);
  let dv = view(local);
  dv.setUint32(0, SIG_LOCAL, true);
  dv.setUint16(4, 20, true); // version needed: 2.0
  dv.setUint16(12, 0x21, true); // date 1980-01-01
  local.set(nameBytes, 30);
  dv.setUint16(26, nameBytes.length, true);
  const central = new Uint8Array(46 + nameBytes.length);
  dv = view(central);
  dv.setUint32(0, SIG_CENTRAL, true);
  dv.setUint16(4, 20, true); // made by
  dv.setUint16(6, 20, true); // needed
  dv.setUint16(14, 0x21, true);
  dv.setUint16(28, nameBytes.length, true);
  central.set(nameBytes, 46);
  return { name, flags: 0, method: 8, local, central, data: new Uint8Array(0), descriptor: null };
}

/**
 * Write a zip from the original entries.
 * changes: Map of entry name -> new uncompressed bytes (Uint8Array), or null to remove the entry.
 * Entries not in `changes` are copied byte-for-byte. Entry order is kept.
 * additions: Map of new entry name -> bytes, written after the existing entries.
 */
export async function writeZip(zip, changes = new Map(), additions = new Map()) {
  for (const name of changes.keys()) {
    if (!zip.entries.some((e) => e.name === name)) throw new Error(`No zip entry named ${name}.`);
  }
  for (const name of additions.keys()) {
    if (zip.entries.some((e) => e.name === name)) throw new Error(`Zip entry ${name} already exists.`);
  }
  const chunks = [];
  const centrals = [];
  let offset = 0;
  // New entries (additions) go at the end, with fresh headers built from a template entry.
  const added = [...additions.keys()].map((name) => newEntry(name));
  changes = new Map([...changes, ...additions]);
  for (const e of [...zip.entries, ...added]) {
    let local, data, descriptor, central;
    if (!changes.has(e.name)) {
      ({ local, data, descriptor } = e);
      central = e.central.slice();
    } else {
      const raw = changes.get(e.name);
      if (raw === null) continue;
      data = await pipe(raw, new CompressionStream('deflate-raw'));
      const crc = crc32(raw);
      const flags = e.flags & ~8; // sizes go in the header, no data descriptor
      local = e.local.slice();
      let dv = view(local);
      dv.setUint16(6, flags, true);
      dv.setUint16(8, 8, true);
      dv.setUint32(14, crc, true);
      dv.setUint32(18, data.length, true);
      dv.setUint32(22, raw.length, true);
      descriptor = null;
      central = e.central.slice();
      dv = view(central);
      dv.setUint16(8, flags, true);
      dv.setUint16(10, 8, true);
      dv.setUint32(16, crc, true);
      dv.setUint32(20, data.length, true);
      dv.setUint32(24, raw.length, true);
    }
    view(central).setUint32(42, offset, true);
    centrals.push(central);
    chunks.push(local, data);
    offset += local.length + data.length;
    if (descriptor) {
      chunks.push(descriptor);
      offset += descriptor.length;
    }
  }
  const cdOffset = offset;
  let cdSize = 0;
  for (const c of centrals) {
    chunks.push(c);
    cdSize += c.length;
  }
  const eocd = new Uint8Array(22);
  const dv = view(eocd);
  dv.setUint32(0, SIG_EOCD, true);
  dv.setUint16(8, centrals.length, true);
  dv.setUint16(10, centrals.length, true);
  dv.setUint32(12, cdSize, true);
  dv.setUint32(16, cdOffset, true);
  dv.setUint16(20, zip.comment.length, true);
  chunks.push(eocd, zip.comment);

  const total = chunks.reduce((n, c) => n + c.length, 0);
  const out = new Uint8Array(total);
  let p = 0;
  for (const c of chunks) {
    out.set(c, p);
    p += c.length;
  }
  return out;
}
