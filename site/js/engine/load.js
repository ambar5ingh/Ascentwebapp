// Build an Engine from the extracted workbook JSON.
import { Engine } from './engine.js';
import { err, Rich } from './values.js';
import { parseRefText } from './parser.js';

export function decodeValue(v) {
  if (v && typeof v === 'object' && v.e) return err(v.e);
  return v;
}

function rangeOf(ref) {
  const r = parseRefText(ref);
  return [r.r1, r.c1, r.r2, r.c2];
}

// Load the precomputed dependency graph (graph.bin) into an engine
export function loadGraph(eng, arrayBuffer) {
  const a = new Int32Array(arrayBuffer);
  let p = 0;
  const [magic, N, T, E, nIn, nVol] = a.subarray(0, 6); p = 6;
  if (magic !== 0x41534e54 || N !== eng.cells.length) throw new Error('graph.bin does not match workbook');
  const order = a.subarray(p, p + N); p += N;
  const offs = a.subarray(p, p + T + 1); p += T + 1;
  const targets = a.subarray(p, p + E); p += E;
  const inputs = a.subarray(p, p + nIn * 6); p += nIn * 6;
  const vol = a.subarray(p, p + nVol);
  for (let k = 0; k < N; k++) eng.cells[order[k]].order = k;
  eng.order = order;
  eng.deps = { length: T, offs, targets };
  eng.depsOf = (id) => targets.subarray(offs[id], offs[id + 1]);
  eng.volatile = Array.from(vol);
  eng.nodeCount = T;
  const idx = eng.sheets.map(() => new Map());
  for (let k = 0; k < inputs.length; k += 6) {
    const si = inputs[k], r1 = inputs[k + 1], c1 = inputs[k + 2], r2 = inputs[k + 3], c2 = inputs[k + 4], id = inputs[k + 5];
    const m = idx[si];
    const cs = Math.min(c2, c1 + 1024);
    for (let c = c1; c <= cs; c++) { let l = m.get(c); if (!l) m.set(c, l = []); l.push(r1, r2, id); }
  }
  eng.inputIdx = idx;
}

export function buildEngine(wb, sheetDatas) {
  const eng = new Engine();
  wb.sheets.forEach((s) => eng.addSheet(s.name));
  eng.external = wb.external || {};
  for (const n of wb.names) eng.addName(n.name, n.f, n.sheet);
  for (const t of wb.tables) eng.addTable(t);
  const richVal = (vm) => {
    const i = wb.richMap[vm - 1];
    let rv = wb.rich[i];
    if (rv && rv['%cvi'] !== undefined) rv = wb.rich[parseInt(rv['%cvi'], 10)];
    return rv ? new Rich(rv) : null;
  };
  sheetDatas.forEach((d, si) => {
    const sh = eng.sheets[si];
    if (d.rows) {
      for (const run of d.rows) {
        const r = run[0], c0 = run[1];
        for (let k = 2; k < run.length; k++) {
          const v = run[k];
          if (v !== null && v !== undefined) sh.set(r, c0 + k - 2, decodeValue(v));
        }
        if (r > sh.maxRow) sh.maxRow = r;
        if (c0 + run.length - 3 > sh.maxCol) sh.maxCol = c0 + run.length - 3;
      }
    } else {
      for (const cell of d.cells) {
        const v = cell[2];
        if (v !== null && v !== undefined) sh.set(cell[0], cell[1], decodeValue(v));
        else { if (cell[0] > sh.maxRow) sh.maxRow = cell[0]; if (cell[1] > sh.maxCol) sh.maxCol = cell[1]; }
      }
    }
    if (d.rich) for (const [r, c, vm] of d.rich) { const rv = richVal(+vm); if (rv) sh.set(r, c, rv); }
    for (const f of d.formulas) {
      if (wb.texts) {
        // compact format: [r, c, textId, r0?, c0?, kind?, spill?]
        const [r, c, t, r0 = r, c0 = c, kind = '', spill = null] = f;
        eng.addFormula(si, r, c, wb.texts[t], r0, c0, kind, spill);
        continue;
      }
      const [r, c, text, sidx, kind, ref] = f;
      if (kind === 't') continue;
      let t = text, r0 = r, c0 = c;
      if (sidx !== null && sidx !== undefined && (text === null || text === undefined)) {
        const m = d.shared[sidx];
        if (!m) continue;
        [r0, c0, t] = m;
      }
      eng.addFormula(si, r, c, t, r0, c0, kind || '', kind ? rangeOf(ref) : null);
    }
  });
  return eng;
}
