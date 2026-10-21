// Recalculate every formula with the JS engine and compare with Excel's cached results.
import fs from 'fs';
import { buildEngine, loadGraph } from '../site/js/engine/load.js';
import { isErr, Matrix } from '../site/js/engine/values.js';
const dir = process.argv[2] || 'site/data';
const wb = JSON.parse(fs.readFileSync(dir + '/workbook.json'));
const datas = wb.sheets.map((s) => JSON.parse(fs.readFileSync(dir + '/' + s.file)));
const graphBuf = fs.readFileSync(dir + '/graph.bin');
let t = Date.now();
const eng = buildEngine(wb, datas);
console.log('load', Date.now() - t, 'ms; formulas', eng.cells.length);
t = Date.now();
loadGraph(eng, graphBuf.buffer.slice(graphBuf.byteOffset, graphBuf.byteOffset + graphBuf.byteLength));
console.log('graph loaded', Date.now() - t, 'ms');
// snapshot cached values of formula cells (incl. spill ranges)
const cached = new Map();
for (const fc of eng.cells) {
  const sh = eng.sheets[fc.sheet];
  const area = fc.spill && fc.kind ? fc.spill : [fc.r, fc.c, fc.r, fc.c];
  for (let r = area[0]; r <= area[2]; r++) for (let c = area[1]; c <= area[3]; c++) cached.set(fc.sheet + ':' + r + ':' + c, sh.get(r, c));
}
t = Date.now();
eng.tracing = true;
eng.recalcAll();
console.log('recalcAll', Date.now() - t, 'ms');
const norm = (v) => (v === null || v === undefined ? 0 : isErr(v) ? v.code : v);
const eq = (a, b) => {
  a = norm(a); b = norm(b);
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  if (typeof a === 'string' && typeof b === 'string') return a === b;
  if (a === '' && b === 0) return true;
  return a === b || String(a) === String(b);
};
let bad = 0, total = 0;
const bySheet = {}, byText = new Map();
for (const [k, v] of cached) {
  total++;
  const [si, r, c] = k.split(':').map(Number);
  const nv = eng.sheets[si].get(r, c);
  if (!eq(v, nv)) {
    bad++;
    const name = eng.sheets[si].name;
    bySheet[name] = (bySheet[name] || 0) + 1;
    const fc = eng.sheets[si].formulas.get(r * 16384 + c);
    const key = fc ? fc.text.slice(0, 160) : '(spill)';
    if (!byText.has(key)) byText.set(key, { n: 0, ex: [] });
    const e = byText.get(key); e.n++;
    if (e.ex.length < 2) e.ex.push(`${name}!${c}:${r} excel=${JSON.stringify(norm(v))} js=${JSON.stringify(norm(nv))}`);
  }
}
console.log({ total, bad });
console.log(bySheet);
const arr = [...byText.entries()].sort((a, b) => b[1].n - a[1].n);
for (const [k, e] of arr.slice(0, +(process.argv[3] || 40))) console.log(e.n, '|', k.replace(/\s+/g, ' '), '\n    ', e.ex.join('\n     '));
if (bad > 100) process.exit(1);
