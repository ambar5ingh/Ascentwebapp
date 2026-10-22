// Build step: turn the raw extraction (build/raw) into the compact data the web app loads
// (site/data) and precompute the dependency graph (site/data/graph.bin).
//
//   node tools/build.mjs build/raw site/data
import fs from 'fs';
import path from 'path';
import { buildEngine } from '../site/js/engine/load.js';

const [src = 'build/raw', dst = 'site/data'] = process.argv.slice(2);
fs.mkdirSync(path.join(dst, 'sheets'), { recursive: true });

const wb = JSON.parse(fs.readFileSync(path.join(src, 'workbook.json')));
const texts = [];
const textId = new Map();
const tid = (t) => { let i = textId.get(t); if (i === undefined) { i = texts.length; texts.push(t); textId.set(t, i); } return i; };

const parseRef = (ref) => {
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(ref);
  const col = (s) => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const r1 = +m[2], c1 = col(m[1]);
  return m[3] ? [r1, c1, +m[4], col(m[3])] : [r1, c1, r1, c1];
};

const sheets = [];
wb.sheets.forEach((s, i) => {
  const d = JSON.parse(fs.readFileSync(path.join(src, s.file)));
  // normalise formulas: [r, c, textId, r0, c0, kind, spill]
  const formulas = [];
  for (const f of d.formulas) {
    const [r, c, text, sidx, kind, ref] = f;
    if (kind === 't') continue;
    let t = text, r0 = r, c0 = c;
    if (sidx !== null && sidx !== undefined && (text === null || text === undefined)) {
      const m = d.shared[sidx];
      if (!m) continue;
      [r0, c0, t] = m;
    }
    const rec = [r, c, tid(t)];
    if (r0 !== r || c0 !== c || kind) rec.push(r0, c0);
    if (kind) rec.push(kind, parseRef(ref));
    formulas.push(rec);
  }
  // compact cells: row runs [r, c0, v0, v1, ...]; styles as per-column defaults + exceptions
  const colCount = new Map();
  for (const cell of d.cells) {
    const c = cell[1], st = cell[3] || 0;
    let m = colCount.get(c); if (!m) colCount.set(c, m = new Map());
    m.set(st, (m.get(st) || 0) + 1);
  }
  const ps = {};
  for (const [c, m] of colCount) { let best = 0, bn = -1; for (const [st, n] of m) if (n > bn) { bn = n; best = st; } ps[c] = best; }
  const rows = [], sx = [];
  let run = null;
  for (const cell of d.cells) {
    const [r, c, v] = cell, st = cell[3] || 0;
    if (run && run[0] === r && run[1] + run.length - 2 === c) run.push(v);
    else { run = [r, c, v]; rows.push(run); }
    if (st !== ps[c]) sx.push(r, c, st);
  }
  const out = { ...d, rowInfo: d.rows, formulas, cells: undefined, rows, ps, sx };
  delete out.cells;
  delete out.shared;
  delete out.name;
  const file = 'sheets/s' + i + '.json';
  fs.writeFileSync(path.join(dst, file), JSON.stringify(out));
  s.file = file;
  sheets.push(out);
});
wb.texts = texts;
// Reference lists stay read-only. Metadata, the emission factors and GWP are editable in the
// app, so their cells are graph inputs too.
const READ_ONLY = ['City_master', 'ListsAndTables'];
wb.readOnlySheets = READ_ONLY;
// Spare blank rows at the end of the Metadata table, so data added in the app is part of
// Metadata[...] and the questionnaire lookups pick it up.
const META_SPARE = 2000;
const metaTable = (wb.tables || []).find((t) => t.name === 'Metadata');
if (metaTable && !wb.metadata) {
  wb.metadata = { first: metaTable.ref[0] + 1, last: metaTable.ref[2], spare: META_SPARE };
  metaTable.ref[2] += META_SPARE;
}
fs.writeFileSync(path.join(dst, 'workbook.json'), JSON.stringify(wb));

// media
fs.mkdirSync(path.join(dst, 'media'), { recursive: true });
if (fs.existsSync(path.join(src, 'media'))) for (const f of fs.readdirSync(path.join(src, 'media'))) fs.copyFileSync(path.join(src, 'media', f), path.join(dst, 'media', f));

// dependency graph
const eng = buildEngine(wb, sheets);
const st = eng.buildGraph();
console.log('graph', st);
const N = eng.cells.length, T = eng.nodeCount;
const inputs = [];
eng.inputIdx = null;
// collect precedent ranges again (deduplicated)
const readOnly = new Set(READ_ONLY.map((n) => eng.sheetIndex(n)));
const metaSi = eng.sheetIndex('Metadata');
for (const fc of eng.cells) {
  // the per-row Sr.no codes of Metadata only read their own row's keys, which the app never edits
  if (fc.sheet === metaSi) continue;
  const { ranges } = eng.precedents(fc);
  const seen = new Set();
  for (const [si, r1, c1, r2, c2] of ranges) {
    const k = si + ':' + r1 + ':' + c1 + ':' + r2 + ':' + c2;
    if (seen.has(k) || readOnly.has(si)) continue;
    seen.add(k);
    inputs.push(si, r1, c1, r2, c2, fc.id);
  }
}
let E = 0;
for (let i = 0; i < T; i++) E += eng.deps[i].length;
const header = [0x41534e54, N, T, E, inputs.length / 6, eng.volatile.length];
const buf = new Int32Array(header.length + N + (T + 1) + E + inputs.length + eng.volatile.length);
let p = 0;
for (const x of header) buf[p++] = x;
for (const id of eng.order) buf[p++] = id;
let off = 0;
for (let i = 0; i < T; i++) { buf[p++] = off; off += eng.deps[i].length; }
buf[p++] = off;
for (let i = 0; i < T; i++) for (const d of eng.deps[i]) buf[p++] = d;
for (const x of inputs) buf[p++] = x;
for (const x of eng.volatile) buf[p++] = x;
fs.writeFileSync(path.join(dst, 'graph.bin'), Buffer.from(buf.buffer));
console.log('graph.bin', (buf.byteLength / 1e6).toFixed(1), 'MB; inputs', inputs.length / 6, 'texts', texts.length);
