// Pivot tables on "Dashboard Data" (row fields + one summed data field, compact layout).
// Excel only refreshes them on demand; the web app refreshes them after every change so
// the dashboard charts that read them stay in step with the inputs.
import { numToText, isErr } from './engine/values.js';
import { parseAddr } from './format.js';

const label = (v) => (v === null || v === undefined || v === '' ? '(blank)' : typeof v === 'number' ? numToText(v) : isErr(v) ? v.code : String(v));
const cmp = (a, b) => (a === '(blank)') - (b === '(blank)') || a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true });

function range(ref) {
  const [a, b] = ref.replace(/\$/g, '').split(':');
  const [r1, c1] = parseAddr(a), [r2, c2] = b ? parseAddr(b) : [r1, c1];
  return { r1, c1, r2, c2 };
}

// Compute the cell writes [[si, r, c, v], ...] that bring every pivot up to date.
export function pivotWrites(app) {
  const eng = app.eng;
  const writes = [];
  for (const p of app.wb.pivots || []) {
    const si = eng.sheetIndex(p.sheet), ssi = eng.sheetIndex(p.src.sheet);
    if (si < 0 || ssi < 0 || p.data === null) continue;
    const src = range(p.src.ref), loc = range(p.ref);
    const hidden = Object.fromEntries(Object.entries(p.hidden || {}).map(([k, v]) => [k, new Set(v.map((x) => (x === null ? '(blank)' : x)))]));
    // aggregate
    const tree = new Map();
    let grand = 0, any = false;
    for (let r = src.r1 + 1; r <= src.r2; r++) {
      const keys = p.rows.map((f) => label(eng.getValue(ssi, r, src.c1 + f)));
      if (keys.some((k, i) => hidden[p.rows[i]] && hidden[p.rows[i]].has(k))) continue;
      const v = eng.getValue(ssi, r, src.c1 + p.data);
      const n = typeof v === 'number' ? v : 0;
      let node = tree;
      for (let i = 0; i < keys.length; i++) {
        let e = node.get(keys[i]);
        if (!e) { e = { sum: 0, kids: new Map() }; node.set(keys[i], e); }
        e.sum += n;
        node = e.kids;
      }
      grand += n; any = true;
    }
    const rows = [];
    const walk = (node, depth) => {
      for (const k of [...node.keys()].sort(cmp)) {
        const e = node.get(k);
        rows.push([k, e.sum]);
        if (depth + 1 < p.rows.length) walk(e.kids, depth + 1);
      }
    };
    walk(tree, 0);
    if (p.grandTotal) rows.push(['Grand Total', any ? grand : null]);
    // write below the header row, clearing anything left from a larger previous layout
    const start = loc.r1 + 1;
    const end = Math.max(loc.r2, start + rows.length - 1);
    for (let r = start; r <= end; r++) {
      const row = rows[r - start];
      writes.push([si, r, loc.c1, row ? row[0] : null], [si, r, loc.c1 + 1, row ? row[1] : null]);
    }
    p.ref = `${colLetters(loc.c1)}${loc.r1}:${colLetters(loc.c2)}${start + rows.length - 1}`;
  }
  return writes;
}

function colLetters(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
