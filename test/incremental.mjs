// Change inputs incrementally, then compare with a full recalculation.
import fs from 'fs';
import { buildEngine, loadGraph } from '../site/js/engine/load.js';
import { isErr } from '../site/js/engine/values.js';
const dir = 'site/data';
const wb = JSON.parse(fs.readFileSync(dir + '/workbook.json'));
const datas = wb.sheets.map((s) => JSON.parse(fs.readFileSync(dir + '/' + s.file)));
const g = fs.readFileSync(dir + '/graph.bin');
const eng = buildEngine(wb, datas);
loadGraph(eng, g.buffer.slice(g.byteOffset, g.byteOffset + g.byteLength));
const q = eng.sheetIndex('Questionnaire');
const edits = JSON.parse(process.argv[2] || '[["Questionnaire",9,6,"District"],["Questionnaire",10,6,"Maharashtra"],["Questionnaire",11,6,"Pune"]]');
for (const [sn, r, c, v] of edits) {
  const t = Date.now();
  const ch = eng.setValues([[eng.sheetIndex(sn), r, c, v]]);
  console.log('edit', sn, r, c, v, '->', Date.now() - t, 'ms; formulas', eng.lastRecalc && eng.lastRecalc.formulas, 'changed cells', ch.length);
}
const snap = new Map();
for (const fc of eng.cells) { const sh = eng.sheets[fc.sheet]; snap.set(fc.id, sh.get(fc.r, fc.c)); }
const t = Date.now(); eng.recalcAll(); console.log('full', Date.now() - t, 'ms');
let bad = 0;
const norm = (v) => (v === null || v === undefined ? 0 : isErr(v) ? v.code : v);
for (const fc of eng.cells) {
  const a = norm(snap.get(fc.id)), b = norm(eng.sheets[fc.sheet].get(fc.r, fc.c));
  const ok = typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a)) : a === b;
  if (!ok) { bad++; if (bad < 15) console.log('diff', eng.sheets[fc.sheet].name, fc.r, fc.c, a, b, fc.text.slice(0, 80)); }
}
console.log('incremental vs full mismatches:', bad);
const show = (sn, a) => { const si = eng.sheetIndex(sn); console.log(sn, a, eng.sheets[si].get(...a)); };
show('Questionnaire', [16, 6]); show('Questionnaire', [17, 6]); show('Metadata helper', [16, 6]);
