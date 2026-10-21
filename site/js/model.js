// Per-sheet layout model: styles, merges, row/column sizes (with the hidden state that
// macros change), validation, conditional formats, links and comments.
import { colWidthPx } from './format.js';
import { KEY } from './engine/engine.js';

const MAXR = 1048576, MAXC = 16384;

// Hidden state of rows (or columns) as Excel macros change it: a base (the file's own
// hidden flags, or "all visible" after Rows.Hidden = False) plus ranges applied in order.
export class HiddenState {
  constructor(saved) {
    this.reset = saved ? !!saved.reset : false;
    this.ranges = saved ? saved.ranges.map((x) => x.slice()) : [];
  }
  showAll() { this.reset = true; this.ranges = []; }
  set(r1, r2, hidden) {
    this.ranges = this.ranges.filter(([a, b]) => !(a >= r1 && b <= r2));
    this.ranges.push([r1, r2, hidden ? 1 : 0]);
  }
  // undefined = use the file's flag
  lookup(r) {
    for (let k = this.ranges.length - 1; k >= 0; k--) {
      const [a, b, h] = this.ranges[k];
      if (r >= a && r <= b) return !!h;
    }
    return this.reset ? false : undefined;
  }
  toJSON() { return { reset: this.reset, ranges: this.ranges }; }
}

export class SheetModel {
  constructor(app, si) {
    this.app = app;
    this.si = si;
    this.meta = app.wb.sheets[si];
    const d = this.data = app.datas[si];
    this.name = this.meta.name;
    this.freezeRows = d.freeze ? d.freeze[0] : 0;
    this.freezeCols = d.freeze ? d.freeze[1] : 0;
    this.showGrid = d.showGrid !== false;
    this.noHeadings = !!d.noHeadings;
    this.zeroHeight = !!d.zeroHeight;
    this.defRowH = d.defRowH > 0 ? d.defRowH : 15;
    this.defColW = d.defColW || (d.baseColW ? d.baseColW + 0.83 : 8.83);
    this.rowInfo = d.rowInfo || {};
    this.cols = d.cols || [];
    this.ps = d.ps || {};
    // style exceptions and cell existence
    this.sx = new Map();
    for (let k = 0; k < d.sx.length; k += 3) this.sx.set(KEY(d.sx[k], d.sx[k + 1]), d.sx[k + 2]);
    this.runs = new Map(); // row -> [c0, c1, c0, c1, ...]
    let maxR = 0, maxC = 0;
    for (const run of d.rows) {
      const r = run[0], c0 = run[1], c1 = c0 + run.length - 3;
      let l = this.runs.get(r); if (!l) this.runs.set(r, l = []);
      l.push(c0, c1);
      if (r > maxR) maxR = r;
      if (c1 > maxC) maxC = c1;
    }
    // merges
    this.merges = d.merges || [];
    this.mergeIdx = new Map();
    this.merges.forEach((m, i) => {
      if ((m[2] - m[0] + 1) * (m[3] - m[1] + 1) > 20000) return;
      for (let r = m[0]; r <= m[2]; r++) for (let c = m[1]; c <= m[3]; c++) this.mergeIdx.set(KEY(r, c), i);
      if (m[2] > maxR) maxR = m[2];
      if (m[3] > maxC) maxC = m[3];
    });
    for (const k of Object.keys(this.rowInfo)) if (+k > maxR && !this.rowInfo[k].hidden) maxR = Math.max(maxR, +k);
    for (const dr of d.drawing || []) {
      if (dr.to) { maxR = Math.max(maxR, dr.to[2] + 1); maxC = Math.max(maxC, dr.to[0] + 1); }
      else if (dr.from) { maxR = Math.max(maxR, dr.from[2] + 2); maxC = Math.max(maxC, dr.from[0] + 2); }
    }
    for (const ct of d.controls || []) if (ct.to) { maxR = Math.max(maxR, ct.to[2] + 1); maxC = Math.max(maxC, ct.to[0] + 1); }
    for (const [r, c] of (d.comments || [])) { maxR = Math.max(maxR, r); maxC = Math.max(maxC, c); }
    const sh = app.eng.sheets[si];
    maxR = Math.max(maxR, sh.maxRow);
    maxC = Math.max(maxC, sh.maxCol);
    // column hidden to the end of the sheet?
    let colEnd = MAXC;
    for (const [a, b, , hid] of this.cols) if (hid && b >= MAXC) colEnd = Math.min(colEnd, a - 1);
    this.nCols = Math.min(colEnd, maxC + 30);
    this.nRows = Math.min(MAXR, maxR + (this.zeroHeight ? 2 : 150));
    this.colInfo = new Array(this.nCols + 2);
    for (const [a, b, w, hid, st] of this.cols) {
      for (let c = a; c <= Math.min(b, this.nCols + 1); c++) this.colInfo[c] = { w, hidden: !!hid, s: st };
    }
    this.buildDV();
    this.buildCF();
    this.links = new Map();
    for (const l of d.links || []) {
      const [r1, c1, r2, c2] = l.ref;
      if ((r2 - r1 + 1) * (c2 - c1 + 1) > 5000) continue;
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (!this.links.has(KEY(r, c))) this.links.set(KEY(r, c), l);
    }
    this.comments = new Map();
    for (const [r, c, text] of d.comments || []) this.comments.set(KEY(r, c), cleanComment(text));
  }

  // ------------------------------------------------------------- styles
  exists(r, c) {
    const l = this.runs.get(r);
    if (!l) return false;
    for (let k = 0; k < l.length; k += 2) if (c >= l[k] && c <= l[k + 1]) return true;
    return false;
  }
  styleId(r, c) {
    const k = KEY(r, c);
    const x = this.sx.get(k);
    if (x !== undefined) return x;
    if (this.exists(r, c)) return this.ps[c] || 0;
    const ri = this.rowInfo[r];
    if (ri && ri.s !== undefined) return ri.s;
    const ci = this.colInfo[c];
    if (ci && ci.s) return ci.s;
    return 0;
  }
  xf(r, c) { return this.app.wb.xfs[this.styleId(r, c)] || {}; }

  mergeAt(r, c) {
    const i = this.mergeIdx.get(KEY(r, c));
    return i === undefined ? null : this.merges[i];
  }

  // ------------------------------------------------------------- hidden rows/cols
  get rowState() { return this.app.rowState(this.si); }
  get colState() { return this.app.colState(this.si); }
  rowHidden(r) {
    const o = this.rowState.lookup(r);
    if (o !== undefined) return o;
    const ri = this.rowInfo[r];
    if (ri) return !!ri.hidden || ri.h === 0;
    return this.zeroHeight;
  }
  colHidden(c) {
    const o = this.colState.lookup(c);
    if (o !== undefined) return o;
    const ci = this.colInfo[c];
    return ci ? ci.hidden || ci.w === 0 : false;
  }
  rowHeightPt(r) {
    if (this.rowHidden(r)) return 0;
    const ri = this.rowInfo[r];
    if (ri && ri.h) return ri.h;
    return this.defRowH;
  }
  colWidthPx100(c) {
    if (this.colHidden(c)) return 0;
    const ci = this.colInfo[c];
    return colWidthPx(ci && ci.w ? ci.w : this.defColW);
  }

  // Compute pixel geometry at a zoom factor. rowY[r] = top of row r (1-based).
  layout(zoom) {
    this.zoom = zoom;
    const R = this.nRows, C = this.nCols;
    const rowY = this.rowY = new Float64Array(R + 2);
    const rowH = this.rowH = new Float64Array(R + 2);
    const rowPt = this.rowPt = new Float64Array(R + 2); // top in points (for Shape.Top)
    let y = 0, pt = 0;
    for (let r = 1; r <= R + 1; r++) {
      rowY[r] = y; rowPt[r] = pt;
      const hpt = r <= R ? this.rowHeightPt(r) : 0;
      const h = Math.round(hpt * 4 / 3 * zoom * 100) / 100;
      rowH[r] = h; y += h; pt += hpt;
    }
    const colX = this.colX = new Float64Array(C + 2);
    const colW = this.colW = new Float64Array(C + 2);
    const colPt = this.colPt = new Float64Array(C + 2);
    let x = 0; pt = 0;
    for (let c = 1; c <= C + 1; c++) {
      colX[c] = x; colPt[c] = pt;
      const w100 = c <= C ? this.colWidthPx100(c) : 0;
      const w = Math.round(w100 * zoom * 100) / 100;
      colW[c] = w; x += w; pt += w100 * 0.75;
    }
    this.totalW = x; this.totalH = y;
    this.frozenH = rowY[Math.min(this.freezeRows, R) + 1] || 0;
    this.frozenW = colX[Math.min(this.freezeCols, C) + 1] || 0;
  }
  // first row whose bottom is below y
  rowAtY(y) {
    const a = this.rowY; let lo = 1, hi = this.nRows;
    if (y < 0) return 1;
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (a[m] <= y) lo = m; else hi = m - 1; }
    // skip zero-height rows
    while (lo < this.nRows && this.rowH[lo] === 0) lo++;
    return lo;
  }
  colAtX(x) {
    const a = this.colX; let lo = 1, hi = this.nCols;
    if (x < 0) return 1;
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (a[m] <= x) lo = m; else hi = m - 1; }
    while (lo < this.nCols && this.colW[lo] === 0) lo++;
    return lo;
  }
  // anchor point -> px at current zoom ([col, colOffEmu, row, rowOffEmu], 0-based)
  anchorX(col, off) {
    const c = Math.min(col + 1, this.nCols + 1);
    const w = this.colW[c] || 0;
    return this.colX[c] + Math.min(off / 9525 * this.zoom, w);
  }
  anchorY(row, off) {
    const r = Math.min(row + 1, this.nRows + 1);
    const h = this.rowH[r] || 0;
    return this.rowY[r] + Math.min(off / 9525 * this.zoom, h);
  }
  anchorTopPt(row, off) {
    const r = Math.min(row + 1, this.nRows + 1);
    const hpt = this.rowHeightPt(r);
    return this.rowPt[r] + Math.min(off / 12700, hpt);
  }

  // ------------------------------------------------------------- validation
  buildDV() {
    this.dvRules = [];
    this.dvIdx = new Map();
    for (const dv of this.data.dv || []) {
      const rule = { ...dv, r0: dv.sqref[0][0], c0: dv.sqref[0][1] };
      this.dvRules.push(rule);
      for (const [r1, c1, r2, c2] of dv.sqref) {
        const n = (r2 - r1 + 1) * (c2 - c1 + 1);
        if (n > 50000) { (this.dvBig = this.dvBig || []).push([r1, c1, r2, c2, rule]); continue; }
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (!this.dvIdx.has(KEY(r, c))) this.dvIdx.set(KEY(r, c), rule);
      }
    }
  }
  dvAt(r, c) {
    const x = this.dvIdx.get(KEY(r, c));
    if (x) return x;
    if (this.dvBig) for (const [r1, c1, r2, c2, rule] of this.dvBig) if (r >= r1 && r <= r2 && c >= c1 && c <= c2) return rule;
    return null;
  }

  // ------------------------------------------------------------- conditional formats
  buildCF() {
    this.cfIdx = new Map();
    this.cfBig = [];
    for (const cf of this.data.cf || []) {
      const rules = cf.rules.slice().sort((a, b) => a.pri - b.pri).map((ru) => ({ ...ru, r0: cf.sqref[0][0], c0: cf.sqref[0][1] }));
      for (const [r1, c1, r2, c2] of cf.sqref) {
        const n = (r2 - r1 + 1) * (c2 - c1 + 1);
        if (n > 20000) { this.cfBig.push([r1, c1, r2, c2, rules]); continue; }
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
          const k = KEY(r, c);
          const l = this.cfIdx.get(k);
          if (l) l.push(...rules); else this.cfIdx.set(k, rules.slice());
        }
      }
    }
  }
  cfRulesAt(r, c) {
    let l = this.cfIdx.get(KEY(r, c)) || null;
    for (const [r1, c1, r2, c2, rules] of this.cfBig) if (r >= r1 && r <= r2 && c >= c1 && c <= c2) l = (l || []).concat(rules);
    return l;
  }
}

function cleanComment(t) {
  if (!t) return '';
  const m = /Comment:\s*([\s\S]*)$/.exec(t);
  if (t.startsWith('[Threaded comment]') && m) return m[1].trim();
  return t.trim();
}
