// Value types and coercion rules that mirror Excel.

export class XLError {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}
const ERRS = {};
export function err(code) {
  if (!ERRS[code]) ERRS[code] = new XLError(code);
  return ERRS[code];
}
export const E = {
  NA: err('#N/A'), VALUE: err('#VALUE!'), REF: err('#REF!'), DIV0: err('#DIV/0!'), NAME: err('#NAME?'),
  NUM: err('#NUM!'), NULL: err('#NULL!'), SPILL: err('#SPILL!'), CALC: err('#CALC!'),
};
export const isErr = (v) => v instanceof XLError;

// Linked data type (e.g. currency pair) cell value
export class Rich {
  constructor(fields) { this.fields = fields; }
  toString() { return this.fields._DisplayString || this.fields.Name || ''; }
}

// Missing function argument, e.g. IF(A1,,1)
export const MISSING = { missing: true, toString() { return ''; } };

export class Matrix {
  constructor(rows, cols, data) { this.rows = rows; this.cols = cols; this.data = data; }
  get(i, j) {
    // broadcast single row/col
    const r = this.rows === 1 ? 0 : i, c = this.cols === 1 ? 0 : j;
    if (r >= this.rows || c >= this.cols) return E.NA;
    return this.data[r * this.cols + c];
  }
  static of(rows) {
    const R = rows.length, C = R ? Math.max(...rows.map((r) => r.length)) : 0;
    const d = new Array(R * C);
    for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) d[i * C + j] = j < rows[i].length ? rows[i][j] : E.NA;
    return new Matrix(R, C, d);
  }
  map(fn) { return new Matrix(this.rows, this.cols, this.data.map(fn)); }
  toRows() {
    const out = [];
    for (let i = 0; i < this.rows; i++) out.push(this.data.slice(i * this.cols, (i + 1) * this.cols));
    return out;
  }
}

// Sparse 0/1 (or FALSE/TRUE) column vector produced by comparing a large column with a scalar.
// Behaves like a Matrix (data materialises lazily) but lets MATCH(1,(a=x)*(b=y),0) run fast.
export class IndexSet extends Matrix {
  constructor(n, idx, numeric) {
    super(n, 1, null);
    this.n = n; this.idx = idx; this.numeric = numeric;
  }
  get data() {
    if (!this._data) {
      const off = this.numeric ? 0 : false, on = this.numeric ? 1 : true;
      const d = new Array(this.n).fill(off);
      for (let k = 0; k < this.idx.length; k++) d[this.idx[k]] = on;
      this._data = d;
    }
    return this._data;
  }
  set data(v) { this._data = v; }
}

export function intersectSorted(a, b) {
  if (a.length > b.length) { const t = a; a = b; b = t; }
  const out = [];
  if (!a.length) return out;
  if (a.length * 16 < b.length) {
    // galloping: binary search each element of the small list
    let lo = 0;
    for (let i = 0; i < a.length; i++) {
      const x = a[i];
      let hi = b.length;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (b[mid] < x) lo = mid + 1; else hi = mid; }
      if (lo < b.length && b[lo] === x) out.push(x);
    }
    return out;
  }
  let i = 0, j = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) { out.push(a[i]); i++; j++; } else if (a[i] < b[j]) i++; else j++;
  }
  return out;
}
export function unionSorted(lists) {
  lists = lists.filter((l) => l.length);
  if (!lists.length) return [];
  if (lists.length === 1) return lists[0];
  let out = lists[0];
  for (let k = 1; k < lists.length; k++) {
    const a = out, b = lists[k], m = [];
    let i = 0, j = 0;
    while (i < a.length || j < b.length) {
      if (j >= b.length || (i < a.length && a[i] < b[j])) m.push(a[i++]);
      else if (i >= a.length || b[j] < a[i]) m.push(b[j++]);
      else { m.push(a[i]); i++; j++; }
    }
    out = m;
  }
  return out;
}

export class Ref {
  constructor(book, sheet, r1, c1, r2, c2) {
    this.book = book; this.sheet = sheet; this.r1 = r1; this.c1 = c1; this.r2 = r2; this.c2 = c2;
  }
  get rows() { return this.r2 - this.r1 + 1; }
  get cols() { return this.c2 - this.c1 + 1; }
  isCell() { return this.r1 === this.r2 && this.c1 === this.c2; }
  value(i = 0, j = 0) { return this.book.getValue(this.sheet, this.r1 + i, this.c1 + j); }
  toMatrix() {
    const R = this.rows, C = this.cols;
    const big = R * C >= 512;
    let key;
    if (big) {
      const sh0 = this.book.sheets[this.sheet];
      key = this.sheet + ':' + this.r1 + ':' + this.c1 + ':' + this.r2 + ':' + this.c2;
      const hit = this.book.matCache && this.book.matCache.get(key);
      if (hit && hit.version === sh0.version) return hit.m;
    }
    const d = new Array(R * C);
    const sh = this.book.sheets[this.sheet];
    for (let i = 0; i < R; i++) {
      const row = sh.vals[this.r1 + i];
      for (let j = 0; j < C; j++) {
        const v = row ? row[this.c1 + j] : undefined;
        d[i * C + j] = v === undefined ? null : v;
      }
    }
    const m = new Matrix(R, C, d);
    if (big && this.book.matCache) this.book.matCache.set(key, { version: this.book.sheets[this.sheet].version, m });
    return m;
  }
}

// ------------------------------------------------------------------ coercion
export function isBlank(v) { return v === null || v === undefined || v === MISSING; }

const NUM_RE = /^\s*[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?\s*$/;
export function parseNumber(s) {
  if (typeof s !== 'string') return null;
  let t = s.trim();
  if (t === '') return null;
  let pct = false;
  if (t.endsWith('%')) { pct = true; t = t.slice(0, -1); }
  t = t.replace(/^\$/, '').replace(/,(?=\d{3}(\D|$))/g, '');
  let neg = false;
  if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
  if (!NUM_RE.test(t)) return null;
  let n = parseFloat(t);
  if (neg) n = -n;
  return pct ? n / 100 : n;
}

export function toNum(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === MISSING) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof XLError) return v;
  if (typeof v === 'string') { const n = parseNumber(v); return n === null ? E.VALUE : n; }
  if (v instanceof Rich) return E.VALUE;
  return E.VALUE;
}

export function numToText(n) {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  if (!isFinite(n)) return '#NUM!';
  const abs = Math.abs(n);
  if (abs !== 0 && (abs >= 1e15 || abs < 1e-9)) {
    let s = n.toExponential(14).replace(/\.?0+e/, 'e');
    const [m, e] = s.split('e');
    const ex = parseInt(e, 10);
    return m + 'E' + (ex < 0 ? '-' : '+') + String(Math.abs(ex)).padStart(2, '0');
  }
  let s = n.toPrecision(15);
  if (s.includes('e')) return String(parseFloat(s));
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s;
}

export function toText(v) {
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numToText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v === null || v === undefined || v === MISSING) return '';
  if (v instanceof XLError) return v;
  if (v instanceof Rich) return v.toString();
  return String(v);
}

export function toBool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === undefined || v === MISSING) return false;
  if (v instanceof XLError) return v;
  if (typeof v === 'string') {
    const u = v.toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    return E.VALUE;
  }
  return E.VALUE;
}

// Excel compares numbers to ~15 significant digits
export function numEq(a, b) {
  if (a === b) return true;
  const d = Math.abs(a - b);
  const m = Math.max(Math.abs(a), Math.abs(b));
  return d <= m * 1e-15;
}

// type rank: number(1) < text(2) < bool(3)
function rank(v) {
  if (typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 3;
  return 0;
}

// returns -1,0,1 or error
export function compare(a, b) {
  if (a instanceof XLError) return a;
  if (b instanceof XLError) return b;
  if (a instanceof Rich) a = a.toString();
  if (b instanceof Rich) b = b.toString();
  const ab = isBlank(a), bb = isBlank(b);
  if (ab && bb) return 0;
  if (ab) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (bb) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (ra === 1) return numEq(a, b) ? 0 : a < b ? -1 : 1;
  if (ra === 2) {
    const x = a.toLowerCase(), y = b.toLowerCase();
    return x === y ? 0 : x < y ? -1 : 1;
  }
  return a === b ? 0 : a ? 1 : -1;
}

export function round15(n) {
  if (typeof n !== 'number' || !isFinite(n) || n === 0) return n;
  return parseFloat(n.toPrecision(15));
}
