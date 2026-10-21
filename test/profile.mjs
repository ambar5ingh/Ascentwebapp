import fs from 'fs';
import { buildEngine } from '../site/js/engine/load.js';
const dir = 'site/data';
const wb = JSON.parse(fs.readFileSync(dir + '/workbook.json'));
const datas = wb.sheets.map((s) => JSON.parse(fs.readFileSync(dir + '/' + s.file)));
const eng = buildEngine(wb, datas);
// time per sheet & per text without graph: evaluate in sheet order (values cached)
const per = new Map(), perSheet = {};
eng.epoch++;
for (const fc of eng.cells) {
  const t = process.hrtime.bigint();
  eng.evalCell(fc);
  const dt = Number(process.hrtime.bigint() - t) / 1e6;
  const k = fc.text.slice(0, 100);
  per.set(k, (per.get(k) || 0) + dt);
  const sn = eng.sheets[fc.sheet].name; perSheet[sn] = (perSheet[sn] || 0) + dt;
}
console.log(Object.entries(perSheet).sort((a, b) => b[1] - a[1]).slice(0, 12).map(([a, b]) => a + ' ' + b.toFixed(0)).join('\n'));
console.log([...per.entries()].sort((a, b) => b[1] - a[1]).slice(0, 15).map(([a, b]) => b.toFixed(0) + 'ms ' + a.replace(/\s+/g, ' ')).join('\n'));
