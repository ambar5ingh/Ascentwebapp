import fs from 'fs';
import { parse } from '../site/js/engine/parser.js';
const wb = JSON.parse(fs.readFileSync('site/data/workbook.json'));
let ok = 0, bad = 0; const errs = {};
const seen = new Set();
for (const s of wb.sheets) {
  const d = JSON.parse(fs.readFileSync('site/data/' + s.file));
  for (const f of d.formulas) {
    const t = f[2];
    if (!t || seen.has(t)) continue;
    seen.add(t);
    try { parse(t); ok++; } catch (e) { bad++; if (Object.keys(errs).length < 30) errs[s.name + '!' + f[0] + ',' + f[1]] = e.message + ' :: ' + t.slice(0, 200); }
  }
}
for (const n of wb.names) { try { parse(n.f); ok++; } catch (e) { bad++; errs['name ' + n.name] = e.message + ' :: ' + n.f.slice(0, 200); } }
console.log({ ok, bad }); console.log(errs);
