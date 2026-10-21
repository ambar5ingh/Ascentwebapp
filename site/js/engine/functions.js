// Excel worksheet functions used by the ASCENT workbook (and common neighbours).
// Each entry: {args: kinds string, fn, lazy?, lift?}
//   kinds: 'v' scalar (ranges are implicitly intersected / lifted in array formulas)
//          'r' raw (reference, array or scalar passed through)
//          'a' array (references converted to Matrix)
//   The last kind repeats for extra arguments.

import {
  E, err, isErr, XLError, Matrix, Ref, Rich, MISSING, IndexSet, toNum, toText, toBool, compare, isBlank, parseNumber, numEq,
} from './values.js';

const SSF_ = () => (typeof globalThis !== 'undefined' && globalThis.SSF) || null;

// ------------------------------------------------------------------ helpers
function scalarOf(v, ctx, engine) {
  if (v instanceof Ref) {
    if (v.isCell()) return v.value();
    if (ctx.array) return v.toMatrix();
    return intersectRef(v, ctx);
  }
  if (Array.isArray(v)) return E.VALUE;
  return v;
}
function intersectRef(ref, ctx) {
  const row = ctx.row, col = ctx.col;
  if (ref.c1 === ref.c2 && row >= ref.r1 && row <= ref.r2) return ref.book.getValue(ref.sheet, row, ref.c1);
  if (ref.r1 === ref.r2 && col >= ref.c1 && col <= ref.c2) return ref.book.getValue(ref.sheet, ref.r1, col);
  if (row >= ref.r1 && row <= ref.r2 && col >= ref.c1 && col <= ref.c2) return ref.book.getValue(ref.sheet, row, col);
  return E.VALUE;
}
function first(v) {
  if (v instanceof Ref) return v.value(0, 0);
  if (v instanceof Matrix) return v.get(0, 0);
  return v;
}
// iterate leaf values: cb(value, fromRef)
function each(args, cb) {
  for (const a of args) {
    if (a === MISSING) { cb(MISSING, false); continue; }
    if (a instanceof Ref) {
      const sh = a.book.sheets[a.sheet];
      const r2 = Math.min(a.r2, sh.maxRow), c2 = Math.min(a.c2, sh.maxCol);
      for (let r = a.r1; r <= r2; r++) {
        const row = sh.vals[r]; if (!row) continue;
        for (let c = a.c1; c <= c2; c++) { const v = row[c]; if (v !== undefined && v !== null) cb(v, true); }
      }
    } else if (a instanceof Matrix) {
      for (const v of a.data) cb(v, true);
    } else if (Array.isArray(a)) {
      each(a, cb);
    } else cb(a, false);
  }
}
function toMatrix(v) {
  if (v instanceof Matrix) return v;
  if (v instanceof Ref) return v.toMatrix();
  return new Matrix(1, 1, [v]);
}
function numArgs(args, opts = {}) {
  // returns array of numbers or error
  const out = [];
  let e = null;
  each(args, (v, fromRef) => {
    if (e) return;
    if (isErr(v)) { e = v; return; }
    if (typeof v === 'number') out.push(v);
    else if (fromRef) {
      if (opts.countA && v !== null && v !== '' ) out.push(typeof v === 'boolean' ? (v ? 1 : 0) : 0);
    } else if (v === MISSING) { if (!opts.skipMissing) out.push(0); }
    else if (typeof v === 'boolean') out.push(v ? 1 : 0);
    else if (typeof v === 'string') { const n = parseNumber(v); if (n === null) e = E.VALUE; else out.push(n); }
    else if (v === null) out.push(0);
  });
  return e || out;
}
function rnd(x, d) {
  const m = Math.pow(10, d);
  const y = Math.abs(x) * m;
  return Math.sign(x) * Math.round(+y.toPrecision(15)) / m;
}

// criteria matching for SUMIF/COUNTIF family
function makeCriteria(crit) {
  if (crit instanceof Ref) crit = crit.value();
  if (crit instanceof Matrix) crit = crit.get(0, 0);
  if (isBlank(crit)) crit = 0;
  if (isErr(crit)) return () => false;
  if (typeof crit === 'number' || typeof crit === 'boolean') {
    return (v) => compare(v, crit) === 0 && !isBlank(v) && typeof v === typeof crit
      || (typeof v === 'string' && typeof crit === 'number' && parseNumber(v) !== null && numEq(parseNumber(v), crit));
  }
  let s = String(crit);
  let op = '=';
  const m = /^(<=|>=|<>|=|<|>)/.exec(s);
  if (m) { op = m[1]; s = s.slice(m[1].length); }
  const n = parseNumber(s);
  if (n !== null && s.trim() !== '') {
    return (v) => {
      let x = v;
      if (typeof x === 'string') { const p = parseNumber(x); if (p === null) return op === '<>'; x = p; }
      if (typeof x !== 'number') return op === '<>';
      switch (op) {
        case '=': return numEq(x, n);
        case '<>': return !numEq(x, n);
        case '<': return x < n && !numEq(x, n);
        case '>': return x > n && !numEq(x, n);
        case '<=': return x <= n || numEq(x, n);
        case '>=': return x >= n || numEq(x, n);
      }
      return false;
    };
  }
  const low = s.toLowerCase();
  if (s === '' && op === '=') return (v) => isBlank(v) || v === '';
  if (s === '' && op === '<>') return (v) => !(isBlank(v) || v === '');
  if (low === 'true' || low === 'false') {
    const b = low === 'true';
    return (v) => (op === '<>' ? v !== b : v === b);
  }
  const wild = /[*?]/.test(s);
  const re = wild ? wildcardRe(s) : null;
  return (v) => {
    if (isBlank(v)) v = '';
    if (typeof v !== 'string') return op === '<>';
    const x = v.toLowerCase();
    switch (op) {
      case '=': return wild ? re.test(v) : x === low;
      case '<>': return wild ? !re.test(v) : x !== low;
      case '<': return x < low;
      case '>': return x > low;
      case '<=': return x <= low;
      case '>=': return x >= low;
    }
    return false;
  };
}
function wildcardRe(s) {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '~' && i + 1 < s.length) { out += s[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
    if (ch === '*') out += '[\\s\\S]*';
    else if (ch === '?') out += '[\\s\\S]';
    else out += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp('^' + out + '$', 'i');
}

// iterate cells of a Ref/Matrix as flat list (row-major)
function cellsOf(v) {
  if (v instanceof Ref) {
    const out = new Array(v.rows * v.cols);
    const sh = v.book.sheets[v.sheet];
    let k = 0;
    for (let r = v.r1; r <= v.r2; r++) {
      const row = sh.vals[r];
      for (let c = v.c1; c <= v.c2; c++) { const x = row ? row[c] : undefined; out[k++] = x === undefined ? null : x; }
    }
    return { rows: v.rows, cols: v.cols, data: out };
  }
  if (v instanceof Matrix) return v;
  return { rows: 1, cols: 1, data: [v] };
}

function ifsAggregate(args, mode, valuesFirst) {
  // valuesFirst: SUMIFS(sum_range, r1, c1, ...); otherwise COUNTIFS(r1,c1,...)
  let target = null, i0 = 0;
  if (valuesFirst) { target = cellsOf(args[0]); i0 = 1; }
  const conds = [];
  for (let i = i0; i + 1 < args.length; i += 2) conds.push([cellsOf(args[i]), makeCriteria(args[i + 1])]);
  if (!conds.length) return E.VALUE;
  const R = conds[0][0].rows, C = conds[0][0].cols;
  for (const [rg] of conds) if (rg.rows !== R || rg.cols !== C) return E.VALUE;
  if (target && (target.rows !== R || target.cols !== C)) return E.VALUE;
  let sum = 0, cnt = 0, mx = -Infinity, mn = Infinity;
  const n = R * C;
  for (let k = 0; k < n; k++) {
    let ok = true;
    for (const [rg, f] of conds) { if (!f(rg.data[k])) { ok = false; break; } }
    if (!ok) continue;
    if (!target) { cnt++; continue; }
    const v = target.data[k];
    if (isErr(v)) return v;
    if (typeof v === 'number') { sum += v; cnt++; if (v > mx) mx = v; if (v < mn) mn = v; }
  }
  switch (mode) {
    case 'sum': return sum;
    case 'count': return cnt;
    case 'avg': return cnt ? sum / cnt : E.DIV0;
    case 'max': return cnt ? mx : 0;
    case 'min': return cnt ? mn : 0;
  }
  return E.VALUE;
}

// ------------------------------------------------------------------ lookups
function lookupEq(a, b) {
  // exact-match equality used by MATCH/VLOOKUP/XLOOKUP
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  if (typeof a === 'number' && typeof b === 'number') return numEq(a, b);
  if (isBlank(a) && isBlank(b)) return true;
  return a === b;
}

function vecOf(arr) {
  // returns {get(i), n, horizontal}
  if (arr instanceof Ref) {
    const horizontal = arr.rows === 1 && arr.cols > 1;
    const n = horizontal ? arr.cols : arr.rows;
    const sh = arr.book.sheets[arr.sheet];
    if (horizontal) {
      const row = sh.vals[arr.r1];
      return { n, horizontal, get: (i) => { const v = row ? row[arr.c1 + i] : undefined; return v === undefined ? null : v; } };
    }
    return { n, horizontal, get: (i) => { const row = sh.vals[arr.r1 + i]; const v = row ? row[arr.c1] : undefined; return v === undefined ? null : v; } };
  }
  if (arr instanceof Matrix) {
    const horizontal = arr.rows === 1 && arr.cols > 1;
    const n = horizontal ? arr.cols : arr.rows;
    return { n, horizontal, get: (i) => (horizontal ? arr.data[i] : arr.data[i * arr.cols]) };
  }
  return { n: 1, horizontal: false, get: () => arr };
}

function matchExact(val, vec, wildcard = true) {
  if (typeof val === 'string' && wildcard && /[*?~]/.test(val)) {
    const re = wildcardRe(val);
    for (let i = 0; i < vec.n; i++) { const v = vec.get(i); if (typeof v === 'string' && re.test(v)) return i; }
    return -1;
  }
  if (typeof val === 'string') {
    const low = val.toLowerCase();
    for (let i = 0; i < vec.n; i++) { const v = vec.get(i); if (typeof v === 'string' && v.toLowerCase() === low) return i; }
    return -1;
  }
  if (typeof val === 'number') {
    for (let i = 0; i < vec.n; i++) { const v = vec.get(i); if (typeof v === 'number' && (v === val || numEq(v, val))) return i; }
    return -1;
  }
  for (let i = 0; i < vec.n; i++) { if (lookupEq(vec.get(i), val)) return i; }
  return -1;
}

// Excel style approximate match (binary search); type 1 -> largest <= val, -1 -> smallest >= val (descending)
function matchApprox(val, vec, type) {
  const typeOf = (v) => (typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 0);
  const tv = typeOf(val);
  if (type === 1) {
    let lo = 0, hi = vec.n - 1, ans = -1;
    while (lo <= hi) {
      let mid = (lo + hi) >> 1;
      // skip blanks / other types like Excel by scanning down
      let m = mid;
      while (m >= lo && typeOf(vec.get(m)) !== tv) m--;
      if (m < lo) { lo = mid + 1; continue; }
      const c = compare(vec.get(m), val);
      if (c <= 0) { ans = m; lo = mid + 1; } else hi = m - 1;
    }
    return ans;
  }
  // descending
  let ans = -1;
  for (let i = 0; i < vec.n; i++) {
    const v = vec.get(i);
    if (typeOf(v) !== tv) continue;
    const c = compare(v, val);
    if (c >= 0) ans = i; else break;
  }
  return ans;
}

function indexInto(arr, r, c, engine) {
  // r, c are 1-based; 0 means entire row/col; returns Ref or value or Matrix
  if (arr instanceof Ref) {
    let R = arr.rows, C = arr.cols;
    if (r === 0 && c === 0) return arr;
    if (r === 0) { if (c > C) return E.REF; return new Ref(arr.book, arr.sheet, arr.r1, arr.c1 + c - 1, arr.r2, arr.c1 + c - 1); }
    if (c === 0) { if (r > R) return E.REF; return new Ref(arr.book, arr.sheet, arr.r1 + r - 1, arr.c1, arr.r1 + r - 1, arr.c2); }
    if (r > R || c > C || r < 1 || c < 1) return E.REF;
    return new Ref(arr.book, arr.sheet, arr.r1 + r - 1, arr.c1 + c - 1, arr.r1 + r - 1, arr.c1 + c - 1);
  }
  if (arr instanceof Matrix) {
    if (r === 0 && c === 0) return arr;
    if (r === 0) { if (c > arr.cols) return E.REF; const d = []; for (let i = 0; i < arr.rows; i++) d.push(arr.data[i * arr.cols + c - 1]); return new Matrix(arr.rows, 1, d); }
    if (c === 0) { if (r > arr.rows) return E.REF; return new Matrix(1, arr.cols, arr.data.slice((r - 1) * arr.cols, r * arr.cols)); }
    if (r > arr.rows || c > arr.cols || r < 1 || c < 1) return E.REF;
    return arr.data[(r - 1) * arr.cols + c - 1];
  }
  if (r <= 1 && c <= 1) return arr;
  return E.REF;
}

function intArg(v, def) {
  if (v === undefined || v === MISSING) return def;
  if (isBlank(v)) return 0;
  const n = toNum(v);
  if (isErr(n)) return n;
  return Math.trunc(n);
}

function textArg(v) { return toText(v); }

function dateSerial(y, m, d) {
  // Excel 1900 date system (with 1900 leap-year bug)
  const ms = Date.UTC(y, m - 1, d);
  let serial = ms / 86400000 + 25569;
  if (serial < 61) serial -= 1;
  return serial;
}
function serialToDate(n) {
  let s = Math.floor(n);
  if (s < 61) s += 1;
  return new Date((s - 25569) * 86400000);
}

function fmtText(v, fmt) {
  const ssf = SSF_();
  if (typeof v === 'string') {
    const n = parseNumber(v);
    if (n === null) return v;
    v = n;
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isBlank(v)) v = 0;
  if (ssf) { try { return ssf.format(fmt, v); } catch (e) { return toText(v); } }
  return toText(v);
}

function stackMatrices(args, vertical) {
  const ms = args.filter((a) => a !== MISSING).map((a) => (isErr(a) ? new Matrix(1, 1, [a]) : toMatrix(a)));
  if (vertical) {
    const C = Math.max(...ms.map((m) => m.cols));
    const rows = [];
    for (const m of ms) for (let i = 0; i < m.rows; i++) { const row = []; for (let j = 0; j < C; j++) row.push(j < m.cols ? m.data[i * m.cols + j] : E.NA); rows.push(row); }
    return Matrix.of(rows);
  }
  const R = Math.max(...ms.map((m) => m.rows));
  const rows = [];
  for (let i = 0; i < R; i++) { const row = []; for (const m of ms) for (let j = 0; j < m.cols; j++) row.push(i < m.rows ? m.data[i * m.cols + j] : E.NA); rows.push(row); }
  return Matrix.of(rows);
}

function blankTo0(v) { return v === null || v === undefined || v === MISSING ? 0 : v; }

// ------------------------------------------------------------------ table
export const FUNCS = {
  // ---------- logical
  IF: {
    lazy: true,
    fn(a, ctx, engine) {
      let c = scalarOf(a[0](ctx), ctx);
      if (c instanceof Matrix) {
        const t = a.length > 1 ? a[1](ctx) : true;
        const f = a.length > 2 ? a[2](ctx) : false;
        const tv = t instanceof Ref ? (t.isCell() ? t.value() : t.toMatrix()) : t;
        const fv = f instanceof Ref ? (f.isCell() ? f.value() : f.toMatrix()) : f;
        const R = Math.max(c.rows, tv instanceof Matrix ? tv.rows : 1, fv instanceof Matrix ? fv.rows : 1);
        const C = Math.max(c.cols, tv instanceof Matrix ? tv.cols : 1, fv instanceof Matrix ? fv.cols : 1);
        const out = new Array(R * C);
        for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) {
          const cv = c.get(i, j);
          let b = isErr(cv) ? cv : toBool(cv);
          if (isErr(b)) { out[i * C + j] = b; continue; }
          const src = b ? tv : fv;
          let v = src instanceof Matrix ? src.get(i, j) : src;
          if (v === MISSING) v = 0;
          out[i * C + j] = v;
        }
        return new Matrix(R, C, out);
      }
      if (isErr(c)) return c;
      const b = toBool(c);
      if (isErr(b)) return b;
      if (b) { if (a.length < 2) return true; const v = a[1](ctx); return v === MISSING ? 0 : v; }
      if (a.length < 3) return false;
      const v = a[2](ctx);
      return v === MISSING ? 0 : v;
    },
  },
  IFERROR: {
    lazy: true,
    fn(a, ctx) {
      let v = a[0](ctx);
      if (v instanceof Ref) v = v.isCell() ? v.value() : ctx.array ? v.toMatrix() : intersectRef(v, ctx);
      if (v instanceof IndexSet) return v;
      if (v instanceof Matrix) {
        let alt;
        const out = v.data.map((x) => {
          if (!isErr(x)) return x;
          if (alt === undefined) { alt = a.length > 1 ? a[1](ctx) : 0; alt = scalarOf(alt, ctx); if (alt === MISSING) alt = 0; }
          return alt instanceof Matrix ? alt.get(0, 0) : alt;
        });
        return new Matrix(v.rows, v.cols, out);
      }
      if (!isErr(v)) return v === MISSING ? 0 : v;
      const alt = a.length > 1 ? a[1](ctx) : 0;
      return alt === MISSING ? 0 : alt;
    },
  },
  IFNA: {
    lazy: true,
    fn(a, ctx) {
      let v = scalarOf(a[0](ctx), ctx);
      if (v instanceof Matrix) {
        const alt = scalarOf(a[1](ctx), ctx);
        return v.map((x) => (x === E.NA ? (alt instanceof Matrix ? alt.get(0, 0) : alt) : x));
      }
      if (v === E.NA) { const alt = a[1](ctx); return alt === MISSING ? 0 : alt; }
      return v;
    },
  },
  IFS: {
    lazy: true,
    fn(a, ctx) {
      for (let i = 0; i + 1 < a.length; i += 2) {
        const c = scalarOf(a[i](ctx), ctx);
        if (isErr(c)) return c;
        const b = toBool(c instanceof Matrix ? c.get(0, 0) : c);
        if (isErr(b)) return b;
        if (b) return a[i + 1](ctx);
      }
      return E.NA;
    },
  },
  SWITCH: {
    lazy: true,
    fn(a, ctx) {
      const v = scalarOf(a[0](ctx), ctx);
      if (isErr(v)) return v;
      let i = 1;
      for (; i + 1 < a.length; i += 2) {
        const w = scalarOf(a[i](ctx), ctx);
        if (compare(v, w) === 0) return a[i + 1](ctx);
      }
      return i < a.length ? a[i](ctx) : E.NA;
    },
  },
  CHOOSE: {
    lazy: true,
    fn(a, ctx) {
      const i = toNum(scalarOf(a[0](ctx), ctx));
      if (isErr(i)) return i;
      const k = Math.trunc(i);
      if (k < 1 || k >= a.length) return E.VALUE;
      const v = a[k](ctx);
      return v === MISSING ? 0 : v;
    },
  },
  AND: {
    args: 'r',
    fn(args) {
      let res = true, any = false, e = null;
      each(args, (v, fromRef) => {
        if (e) return;
        if (isErr(v)) { e = v; return; }
        if (v === MISSING) return;
        if (typeof v === 'string') { if (fromRef) return; const b = toBool(v); if (isErr(b)) { e = b; return; } any = true; if (!b) res = false; return; }
        if (isBlank(v)) return;
        any = true;
        if (!toBool(v)) res = false;
      });
      if (e) return e;
      return any ? res : E.VALUE;
    },
  },
  OR: {
    args: 'r',
    fn(args) {
      let res = false, any = false, e = null;
      each(args, (v, fromRef) => {
        if (e) return;
        if (isErr(v)) { e = v; return; }
        if (v === MISSING) return;
        if (typeof v === 'string') { if (fromRef) return; const b = toBool(v); if (isErr(b)) { e = b; return; } any = true; if (b) res = true; return; }
        if (isBlank(v)) return;
        any = true;
        if (toBool(v)) res = true;
      });
      if (e) return e;
      return any ? res : E.VALUE;
    },
  },
  XOR: {
    args: 'r',
    fn(args) { let n = 0; let e = null; each(args, (v) => { if (isErr(v)) e = e || v; else if (!isBlank(v) && typeof v !== 'string' && toBool(v)) n++; }); return e || n % 2 === 1; },
  },
  NOT: { args: 'v', fn([v]) { if (isErr(v)) return v; const b = toBool(v); return isErr(b) ? b : !b; } },
  TRUE: { args: '', fn() { return true; } },
  FALSE: { args: '', fn() { return false; } },

  // ---------- information
  ISNUMBER: { args: 'v', fn([v]) { return typeof v === 'number'; } },
  ISTEXT: { args: 'v', fn([v]) { return typeof v === 'string'; } },
  ISNONTEXT: { args: 'v', fn([v]) { return typeof v !== 'string'; } },
  ISLOGICAL: { args: 'v', fn([v]) { return typeof v === 'boolean'; } },
  ISERROR: { args: 'v', fn([v]) { return isErr(v); } },
  ISERR: { args: 'v', fn([v]) { return isErr(v) && v !== E.NA; } },
  ISNA: { args: 'v', fn([v]) { return v === E.NA; } },
  ISEVEN: { args: 'v', fn([v]) { const n = toNum(v); return isErr(n) ? n : Math.trunc(n) % 2 === 0; } },
  ISODD: { args: 'v', fn([v]) { const n = toNum(v); return isErr(n) ? n : Math.abs(Math.trunc(n) % 2) === 1; } },
  ISBLANK: {
    lazy: true,
    fn(a, ctx) {
      const v = a[0](ctx);
      if (v instanceof Ref) {
        if (v.isCell()) return isBlank(v.value());
        if (ctx.array) return v.toMatrix().map((x) => isBlank(x));
        return isBlank(intersectRef(v, ctx));
      }
      if (v instanceof Matrix) return v.map((x) => isBlank(x));
      return false;
    },
  },
  ISREF: { lazy: true, fn(a, ctx) { return a[0](ctx) instanceof Ref; } },
  NA: { args: '', fn() { return E.NA; } },
  N: { args: 'v', fn([v]) { if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? 1 : 0; if (isErr(v)) return v; return 0; } },
  T: { args: 'v', fn([v]) { return typeof v === 'string' ? v : isErr(v) ? v : ''; } },
  TYPE: { args: 'r', lift: false, fn([v]) { if (v instanceof Matrix || (v instanceof Ref && !v.isCell())) return 64; v = first(v); if (typeof v === 'number' || isBlank(v)) return 1; if (typeof v === 'string') return 2; if (typeof v === 'boolean') return 4; if (isErr(v)) return 16; return 1; } },
  ERROR_TYPE: { args: 'v', fn([v]) { const m = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7 }; return isErr(v) ? m[v.code] || 8 : E.NA; } },

  // ---------- math
  SUM: { args: 'r', fn(args) { const ns = numArgs(args, { skipMissing: true }); if (isErr(ns)) return ns; let s = 0; for (const n of ns) s += n; return s; } },
  PRODUCT: { args: 'r', fn(args) { const ns = numArgs(args, { skipMissing: true }); if (isErr(ns)) return ns; let s = 1; for (const n of ns) s *= n; return ns.length ? s : 0; } },
  AVERAGE: { args: 'r', fn(args) { const ns = numArgs(args, { skipMissing: true }); if (isErr(ns)) return ns; if (!ns.length) return E.DIV0; let s = 0; for (const n of ns) s += n; return s / ns.length; } },
  AVERAGEA: { args: 'r', fn(args) { const ns = numArgs(args, { countA: true, skipMissing: true }); if (isErr(ns)) return ns; if (!ns.length) return E.DIV0; return ns.reduce((a, b) => a + b, 0) / ns.length; } },
  MIN: { args: 'r', fn(args) { const ns = numArgs(args, { skipMissing: true }); if (isErr(ns)) return ns; return ns.length ? Math.min(...ns) : 0; } },
  MAX: { args: 'r', fn(args) { const ns = numArgs(args, { skipMissing: true }); if (isErr(ns)) return ns; return ns.length ? Math.max(...ns) : 0; } },
  MEDIAN: { args: 'r', fn(args) { const ns = numArgs(args, { skipMissing: true }); if (isErr(ns)) return ns; if (!ns.length) return E.NUM; ns.sort((a, b) => a - b); const m = ns.length >> 1; return ns.length % 2 ? ns[m] : (ns[m - 1] + ns[m]) / 2; } },
  LARGE: { args: 'rv', lift: false, fn([a, k]) { const ns = numArgs([a], { skipMissing: true }); if (isErr(ns)) return ns; k = toNum(first(k)); if (isErr(k)) return k; ns.sort((x, y) => y - x); return k >= 1 && k <= ns.length ? ns[Math.ceil(k) - 1] : E.NUM; } },
  SMALL: { args: 'rv', lift: false, fn([a, k]) { const ns = numArgs([a], { skipMissing: true }); if (isErr(ns)) return ns; k = toNum(first(k)); if (isErr(k)) return k; ns.sort((x, y) => x - y); return k >= 1 && k <= ns.length ? ns[Math.ceil(k) - 1] : E.NUM; } },
  COUNT: { args: 'r', fn(args) { let n = 0; each(args, (v, fromRef) => { if (typeof v === 'number') n++; else if (!fromRef && (typeof v === 'boolean' || (typeof v === 'string' && parseNumber(v) !== null))) n++; }); return n; } },
  COUNTA: { args: 'r', fn(args) { let n = 0; each(args, (v) => { if (v !== null && v !== undefined && v !== MISSING) n++; }); return n; } },
  COUNTBLANK: { args: 'r', fn([a]) { const m = cellsOf(a); let n = 0; for (const v of m.data) if (isBlank(v) || v === '') n++; return n; } },
  SUMPRODUCT: {
    args: 'a',
    fn(args) {
      const ms = args.map((a) => (a instanceof Matrix ? a : new Matrix(1, 1, [a])));
      const R = ms[0].rows, C = ms[0].cols;
      for (const m of ms) if (m.rows !== R || m.cols !== C) return E.VALUE;
      let s = 0;
      for (let k = 0; k < R * C; k++) {
        let p = 1;
        for (const m of ms) { const v = m.data[k]; if (isErr(v)) return v; p *= typeof v === 'number' ? v : 0; }
        s += p;
      }
      return s;
    },
  },
  SUMIF: {
    args: 'rvr', lift: false,
    fn([range, crit, sumRange]) {
      const rg = cellsOf(range);
      let target = rg;
      if (sumRange !== undefined && sumRange !== MISSING) {
        if (sumRange instanceof Ref) target = cellsOf(new Ref(sumRange.book, sumRange.sheet, sumRange.r1, sumRange.c1, sumRange.r1 + rg.rows - 1, sumRange.c1 + rg.cols - 1));
        else target = cellsOf(sumRange);
      }
      const f = makeCriteria(crit);
      let s = 0;
      for (let k = 0; k < rg.data.length; k++) {
        if (!f(rg.data[k])) continue;
        const v = target.data[k];
        if (isErr(v)) return v;
        if (typeof v === 'number') s += v;
      }
      return s;
    },
  },
  SUMIFS: { args: 'r', lift: false, fn(args) { return ifsAggregate(args, 'sum', true); } },
  COUNTIFS: { args: 'r', lift: false, fn(args) { return ifsAggregate(args, 'count', false); } },
  AVERAGEIFS: { args: 'r', lift: false, fn(args) { return ifsAggregate(args, 'avg', true); } },
  MAXIFS: { args: 'r', lift: false, fn(args) { return ifsAggregate(args, 'max', true); } },
  MINIFS: { args: 'r', lift: false, fn(args) { return ifsAggregate(args, 'min', true); } },
  COUNTIF: {
    args: 'rv', lift: false,
    fn([range, crit]) { const rg = cellsOf(range); const f = makeCriteria(crit); let n = 0; for (const v of rg.data) if (f(v)) n++; return n; },
  },
  AVERAGEIF: {
    args: 'rvr', lift: false,
    fn([range, crit, avgRange]) {
      const rg = cellsOf(range);
      let target = rg;
      if (avgRange !== undefined && avgRange !== MISSING) {
        if (avgRange instanceof Ref) target = cellsOf(new Ref(avgRange.book, avgRange.sheet, avgRange.r1, avgRange.c1, avgRange.r1 + rg.rows - 1, avgRange.c1 + rg.cols - 1));
        else target = cellsOf(avgRange);
      }
      const f = makeCriteria(crit);
      let s = 0, n = 0;
      for (let k = 0; k < rg.data.length; k++) {
        if (!f(rg.data[k])) continue;
        const v = target.data[k];
        if (isErr(v)) return v;
        if (typeof v === 'number') { s += v; n++; }
      }
      return n ? s / n : E.DIV0;
    },
  },
  SUBTOTAL: {
    args: 'vr', lift: false,
    fn([fnum, ...refs]) {
      const n = toNum(fnum); if (isErr(n)) return n;
      const code = Math.trunc(n) % 100;
      const map = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 9: 'SUM' };
      const f = FUNCS[map[code]];
      return f ? f.fn(refs) : E.VALUE;
    },
  },
  ROUND: { args: 'vv', fn([x, d]) { x = toNum(x); if (isErr(x)) return x; d = toNum(blankTo0(d)); if (isErr(d)) return d; return rnd(x, Math.trunc(d)); } },
  ROUNDUP: { args: 'vv', fn([x, d]) { x = toNum(x); if (isErr(x)) return x; d = Math.trunc(toNum(blankTo0(d))); const m = Math.pow(10, d); return Math.sign(x) * Math.ceil(+(Math.abs(x) * m).toPrecision(15)) / m; } },
  ROUNDDOWN: { args: 'vv', fn([x, d]) { x = toNum(x); if (isErr(x)) return x; d = Math.trunc(toNum(blankTo0(d))); const m = Math.pow(10, d); return Math.sign(x) * Math.floor(+(Math.abs(x) * m).toPrecision(15)) / m; } },
  TRUNC: { args: 'vv', fn([x, d]) { x = toNum(x); if (isErr(x)) return x; d = d === undefined ? 0 : Math.trunc(toNum(blankTo0(d))); const m = Math.pow(10, d); return Math.trunc(+(x * m).toPrecision(15)) / m; } },
  INT: { args: 'v', fn([x]) { x = toNum(x); return isErr(x) ? x : Math.floor(x); } },
  ABS: { args: 'v', fn([x]) { x = toNum(x); return isErr(x) ? x : Math.abs(x); } },
  SIGN: { args: 'v', fn([x]) { x = toNum(x); return isErr(x) ? x : Math.sign(x); } },
  SQRT: { args: 'v', fn([x]) { x = toNum(x); return isErr(x) ? x : x < 0 ? E.NUM : Math.sqrt(x); } },
  EXP: { args: 'v', fn([x]) { x = toNum(x); return isErr(x) ? x : Math.exp(x); } },
  LN: { args: 'v', fn([x]) { x = toNum(x); return isErr(x) ? x : x <= 0 ? E.NUM : Math.log(x); } },
  LOG10: { args: 'v', fn([x]) { x = toNum(x); return isErr(x) ? x : x <= 0 ? E.NUM : Math.log10(x); } },
  LOG: { args: 'vv', fn([x, b]) { x = toNum(x); if (isErr(x)) return x; b = b === undefined ? 10 : toNum(b); if (isErr(b)) return b; return x <= 0 || b <= 0 || b === 1 ? E.NUM : Math.log(x) / Math.log(b); } },
  POWER: { args: 'vv', fn([x, y]) { x = toNum(x); if (isErr(x)) return x; y = toNum(y); if (isErr(y)) return y; const r = Math.pow(x, y); return isFinite(r) ? r : E.NUM; } },
  MOD: { args: 'vv', fn([x, y]) { x = toNum(x); if (isErr(x)) return x; y = toNum(y); if (isErr(y)) return y; if (y === 0) return E.DIV0; return x - y * Math.floor(x / y); } },
  PI: { args: '', fn() { return Math.PI; } },
  CEILING: { args: 'vv', fn([x, s]) { x = toNum(x); if (isErr(x)) return x; s = s === undefined ? 1 : toNum(s); if (isErr(s)) return s; if (s === 0) return 0; return Math.ceil(+(x / s).toPrecision(15)) * s; } },
  'CEILING.MATH': { args: 'vv', fn([x, s]) { x = toNum(x); if (isErr(x)) return x; s = s === undefined ? 1 : Math.abs(toNum(s)); if (s === 0) return 0; return Math.ceil(+(x / s).toPrecision(15)) * s; } },
  FLOOR: { args: 'vv', fn([x, s]) { x = toNum(x); if (isErr(x)) return x; s = s === undefined ? 1 : toNum(s); if (isErr(s)) return s; if (s === 0) return E.DIV0; return Math.floor(+(x / s).toPrecision(15)) * s; } },
  'FLOOR.MATH': { args: 'vv', fn([x, s]) { x = toNum(x); if (isErr(x)) return x; s = s === undefined ? 1 : Math.abs(toNum(s)); if (s === 0) return 0; return Math.floor(+(x / s).toPrecision(15)) * s; } },
  RAND: { args: '', fn() { return Math.random(); } },

  // ---------- financial
  FV: {
    args: 'vvvvv',
    fn([rate, nper, pmt, pv, type]) {
      rate = toNum(rate); nper = toNum(nper); pmt = toNum(blankTo0(pmt)); pv = pv === undefined ? 0 : toNum(blankTo0(pv)); type = type === undefined ? 0 : toNum(blankTo0(type));
      for (const x of [rate, nper, pmt, pv, type]) if (isErr(x)) return x;
      if (rate === 0) return -(pv + pmt * nper);
      const f = Math.pow(1 + rate, nper);
      return -(pv * f + pmt * (1 + rate * type) * (f - 1) / rate);
    },
  },
  PV: {
    args: 'vvvvv',
    fn([rate, nper, pmt, fv, type]) {
      rate = toNum(rate); nper = toNum(nper); pmt = toNum(blankTo0(pmt)); fv = fv === undefined ? 0 : toNum(blankTo0(fv)); type = type === undefined ? 0 : toNum(blankTo0(type));
      for (const x of [rate, nper, pmt, fv, type]) if (isErr(x)) return x;
      if (rate === 0) return -(fv + pmt * nper);
      const f = Math.pow(1 + rate, nper);
      return -(fv + pmt * (1 + rate * type) * (f - 1) / rate) / f;
    },
  },
  PMT: {
    args: 'vvvvv',
    fn([rate, nper, pv, fv, type]) {
      rate = toNum(rate); nper = toNum(nper); pv = toNum(pv); fv = fv === undefined ? 0 : toNum(blankTo0(fv)); type = type === undefined ? 0 : toNum(blankTo0(type));
      for (const x of [rate, nper, pv, fv, type]) if (isErr(x)) return x;
      if (rate === 0) return -(pv + fv) / nper;
      const f = Math.pow(1 + rate, nper);
      return -(rate * (fv + pv * f)) / ((1 + rate * type) * (f - 1));
    },
  },
  NPV: {
    args: 'vr', lift: false,
    fn([rate, ...vals]) {
      rate = toNum(first(rate)); if (isErr(rate)) return rate;
      const ns = numArgs(vals, { skipMissing: true }); if (isErr(ns)) return ns;
      let s = 0; ns.forEach((v, i) => { s += v / Math.pow(1 + rate, i + 1); }); return s;
    },
  },

  // ---------- text
  CONCAT: { args: 'r', lift: false, fn(args) { let s = ''; let e = null; each(args, (v) => { if (e) return; if (isErr(v)) { e = v; return; } s += toText(v); }); return e || s; } },
  CONCATENATE: { args: 'v', fn(args) { let s = ''; for (const v of args) { if (isErr(v)) return v; s += toText(v); } return s; } },
  TEXTJOIN: {
    args: 'vvr', lift: false,
    fn([delim, ignoreEmpty, ...rest]) {
      const d = toText(first(delim)); const ig = toBool(first(ignoreEmpty));
      const parts = []; let e = null;
      each(rest, (v) => { if (e) return; if (isErr(v)) { e = v; return; } const t = toText(v); if (ig && t === '') return; parts.push(t); });
      return e || parts.join(d);
    },
  },
  LEN: { args: 'v', fn([v]) { if (isErr(v)) return v; return toText(v).length; } },
  TRIM: { args: 'v', fn([v]) { if (isErr(v)) return v; return toText(v).replace(/ +/g, ' ').trim(); } },
  UPPER: { args: 'v', fn([v]) { if (isErr(v)) return v; return toText(v).toUpperCase(); } },
  LOWER: { args: 'v', fn([v]) { if (isErr(v)) return v; return toText(v).toLowerCase(); } },
  PROPER: { args: 'v', fn([v]) { if (isErr(v)) return v; return toText(v).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, a, b) => a + b.toUpperCase()); } },
  LEFT: { args: 'vv', fn([v, n]) { if (isErr(v)) return v; n = n === undefined ? 1 : toNum(blankTo0(n)); if (isErr(n)) return n; return toText(v).slice(0, Math.max(0, n)); } },
  RIGHT: { args: 'vv', fn([v, n]) { if (isErr(v)) return v; n = n === undefined ? 1 : toNum(blankTo0(n)); if (isErr(n)) return n; const s = toText(v); return n <= 0 ? '' : s.slice(-n); } },
  MID: { args: 'vvv', fn([v, st, n]) { if (isErr(v)) return v; st = toNum(st); n = toNum(n); if (isErr(st)) return st; if (isErr(n)) return n; return toText(v).substr(st - 1, n); } },
  REPT: { args: 'vv', fn([v, n]) { if (isErr(v)) return v; n = toNum(n); if (isErr(n)) return n; return toText(v).repeat(Math.max(0, Math.trunc(n))); } },
  EXACT: { args: 'vv', fn([a, b]) { if (isErr(a)) return a; if (isErr(b)) return b; return toText(a) === toText(b); } },
  SUBSTITUTE: {
    args: 'vvvv',
    fn([text, oldT, newT, inst]) {
      for (const x of [text, oldT, newT]) if (isErr(x)) return x;
      const s = toText(text), o = toText(oldT), n = toText(newT);
      if (o === '') return s;
      if (inst === undefined || inst === MISSING) return s.split(o).join(n);
      const k = Math.trunc(toNum(inst));
      let idx = -1;
      for (let i = 0; i < k; i++) { idx = s.indexOf(o, idx + 1); if (idx < 0) return s; }
      return s.slice(0, idx) + n + s.slice(idx + o.length);
    },
  },
  REPLACE: { args: 'vvvv', fn([t, st, n, nt]) { const s = toText(t); st = toNum(st); n = toNum(n); return s.slice(0, st - 1) + toText(nt) + s.slice(st - 1 + n); } },
  SEARCH: {
    args: 'vvv',
    fn([find, within, st]) {
      if (isErr(find)) return find; if (isErr(within)) return within;
      const f = toText(find), w = toText(within);
      const start = st === undefined ? 1 : toNum(st);
      if (isErr(start)) return start;
      if (/[*?~]/.test(f)) {
        const re = new RegExp(wildcardRe(f).source.slice(1, -1), 'i');
        const m = re.exec(w.slice(start - 1));
        return m ? m.index + start : E.VALUE;
      }
      const i = w.toLowerCase().indexOf(f.toLowerCase(), start - 1);
      return i < 0 ? E.VALUE : i + 1;
    },
  },
  FIND: {
    args: 'vvv',
    fn([find, within, st]) {
      if (isErr(find)) return find; if (isErr(within)) return within;
      const start = st === undefined ? 1 : toNum(st);
      const i = toText(within).indexOf(toText(find), start - 1);
      return i < 0 ? E.VALUE : i + 1;
    },
  },
  VALUE: {
    args: 'v',
    fn([v]) {
      if (isErr(v)) return v;
      if (typeof v === 'number') return v;
      if (isBlank(v)) return 0;
      if (typeof v === 'boolean') return E.VALUE;
      const n = parseNumber(toText(v));
      return n === null ? E.VALUE : n;
    },
  },
  NUMBERVALUE: { args: 'v', fn([v]) { const n = parseNumber(toText(v)); return n === null ? E.VALUE : n; } },
  TEXT: { args: 'vv', fn([v, f]) { if (isErr(v)) return v; if (isErr(f)) return f; return fmtText(v, toText(f)); } },
  FIXED: { args: 'vvv', fn([v, d, nc]) { v = toNum(v); if (isErr(v)) return v; d = d === undefined ? 2 : toNum(d); const s = rnd(v, d).toFixed(Math.max(0, d)); return nc && toBool(nc) ? s : s.replace(/\B(?=(\d{3})+(?!\d))/g, ','); } },
  CHAR: { args: 'v', fn([v]) { v = toNum(v); return isErr(v) ? v : String.fromCharCode(v); } },
  CODE: { args: 'v', fn([v]) { const s = toText(v); return s.length ? s.charCodeAt(0) : E.VALUE; } },
  CLEAN: { args: 'v', fn([v]) { return toText(v).replace(/[\x00-\x1f]/g, ''); } },
  TEXTSPLIT: {
    args: 'vrrvvv', lift: false,
    fn([text, colDelim, rowDelim, ignoreEmpty, matchMode, padWith]) {
      if (isErr(text)) return text;
      const s = toText(text);
      const delims = (d) => {
        if (d === undefined || d === MISSING || isBlank(d)) return [];
        if (d instanceof Ref || d instanceof Matrix) return cellsOf(d).data.map(toText).filter((x) => x !== '');
        return [toText(d)];
      };
      const cd = delims(colDelim), rd = delims(rowDelim);
      const ig = ignoreEmpty !== undefined && ignoreEmpty !== MISSING && toBool(ignoreEmpty) === true;
      const ci = matchMode !== undefined && matchMode !== MISSING && toNum(matchMode) === 1;
      const pad = padWith === undefined || padWith === MISSING ? E.NA : padWith;
      const splitBy = (str, ds) => {
        if (!ds.length) return [str];
        const re = new RegExp(ds.map((x) => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'), ci ? 'gi' : 'g');
        let parts = str.split(re);
        if (ig) parts = parts.filter((p) => p !== '');
        return parts;
      };
      const rows = splitBy(s, rd).map((r) => splitBy(r, cd));
      const C = Math.max(...rows.map((r) => r.length));
      return Matrix.of(rows.map((r) => { const x = r.slice(); while (x.length < C) x.push(pad); return x; }));
    },
  },

  // ---------- lookup & reference
  MATCH: {
    args: 'vrv', lift: true,
    fn([val, arr, type]) {
      if (isErr(val)) return val;
      if (isBlank(val)) return E.NA;
      const t = type === undefined || type === MISSING ? 1 : toNum(type);
      if (isErr(t)) return t;
      if (arr instanceof Ref || arr instanceof Matrix) {
        if (arr.rows > 1 && arr.cols > 1) return E.NA;
      } else if (isErr(arr)) return arr;
      if (arr instanceof IndexSet && t === 0) {
        if ((arr.numeric && val === 1) || (!arr.numeric && val === true)) return arr.idx.length ? arr.idx[0] + 1 : E.NA;
      }
      const vec = vecOf(arr);
      let i;
      if (t === 0) i = matchExact(val, vec);
      else if (t > 0) i = matchApprox(val, vec, 1);
      else i = matchApprox(val, vec, -1);
      return i < 0 ? E.NA : i + 1;
    },
  },
  XMATCH: {
    args: 'vrvv',
    fn([val, arr, mode, search]) {
      if (isErr(val)) return val;
      const vec = vecOf(arr);
      const m = mode === undefined || mode === MISSING ? 0 : toNum(mode);
      const sm = search === undefined || search === MISSING ? 1 : toNum(search);
      const idxs = sm < 0 ? [...Array(vec.n).keys()].reverse() : [...Array(vec.n).keys()];
      let best = -1;
      for (const i of idxs) {
        const v = vec.get(i);
        const c = compare(v, val);
        if (isErr(c)) continue;
        if (m === 2 && typeof val === 'string') { if (wildcardRe(val).test(toText(v))) return i + 1; continue; }
        if (c === 0) return i + 1;
        if (m === -1 && c < 0 && (best < 0 || compare(v, vec.get(best)) > 0)) best = i;
        if (m === 1 && c > 0 && (best < 0 || compare(v, vec.get(best)) < 0)) best = i;
      }
      return best >= 0 ? best + 1 : E.NA;
    },
  },
  INDEX: {
    args: 'rvvv', lift: true,
    fn([arr, r, c, area]) {
      if (isErr(arr)) return arr;
      if (Array.isArray(arr) && arr.isUnion) { const k = area === undefined ? 1 : toNum(area); arr = arr[k - 1]; }
      const rGiven = r !== undefined && r !== MISSING, cGiven = c !== undefined && c !== MISSING;
      let R = rGiven ? toNum(blankTo0(r)) : 0;
      if (isErr(R)) return R;
      let C = cGiven ? toNum(blankTo0(c)) : 0;
      if (isErr(C)) return C;
      R = Math.trunc(R); C = Math.trunc(C);
      const rows = arr instanceof Ref || arr instanceof Matrix ? arr.rows : 1;
      const cols = arr instanceof Ref || arr instanceof Matrix ? arr.cols : 1;
      if (!cGiven) {
        if (rows === 1 && cols > 1) { if (R < 0) return E.VALUE; return indexInto(arr, 1, R); }
        C = cols === 1 ? 1 : 0;
      }
      if (R < 0 || C < 0) return E.VALUE;
      return indexInto(arr, R, C);
    },
  },
  VLOOKUP: {
    args: 'vrvv', lift: true,
    fn([val, table, col, approx]) {
      if (isErr(val)) return val;
      if (isErr(table)) return table;
      const ci = toNum(col); if (isErr(ci)) return ci;
      const ap = approx === undefined || approx === MISSING ? true : toBool(approx);
      const m = table instanceof Ref || table instanceof Matrix ? table : new Matrix(1, 1, [table]);
      if (ci < 1) return E.VALUE;
      if (ci > m.cols) return E.REF;
      const first_ = m instanceof Ref ? new Ref(m.book, m.sheet, m.r1, m.c1, m.r2, m.c1) : new Matrix(m.rows, 1, Array.from({ length: m.rows }, (_, i) => m.data[i * m.cols]));
      const vec = vecOf(first_);
      vec.horizontal = false;
      const i = ap ? matchApprox(val, vec, 1) : matchExact(val, vec);
      if (i < 0) return E.NA;
      const v = m instanceof Ref ? m.value(i, Math.trunc(ci) - 1) : m.data[i * m.cols + Math.trunc(ci) - 1];
      return v === null ? 0 : v;
    },
  },
  HLOOKUP: {
    args: 'vrvv', lift: true,
    fn([val, table, row, approx]) {
      if (isErr(val)) return val;
      const ri = toNum(row); if (isErr(ri)) return ri;
      const ap = approx === undefined || approx === MISSING ? true : toBool(approx);
      const m = table instanceof Ref || table instanceof Matrix ? table : new Matrix(1, 1, [table]);
      if (ri > m.rows) return E.REF;
      const firstRow = m instanceof Ref ? new Ref(m.book, m.sheet, m.r1, m.c1, m.r1, m.c2) : new Matrix(1, m.cols, m.data.slice(0, m.cols));
      const vec = vecOf(firstRow);
      const i = ap ? matchApprox(val, vec, 1) : matchExact(val, vec);
      if (i < 0) return E.NA;
      const v = m instanceof Ref ? m.value(Math.trunc(ri) - 1, i) : m.data[(Math.trunc(ri) - 1) * m.cols + i];
      return v === null ? 0 : v;
    },
  },
  LOOKUP: {
    args: 'vrr', lift: true,
    fn([val, look, result]) {
      if (isErr(val)) return val;
      const vec = vecOf(look);
      const i = matchApprox(val, vec, 1);
      if (i < 0) return E.NA;
      const res = result === undefined ? look : result;
      const rv = vecOf(res);
      const v = rv.get(i);
      return v === null ? 0 : v;
    },
  },
  XLOOKUP: {
    args: 'vrrrvv', lift: true,
    fn([val, look, ret, notFound, mode, search]) {
      if (isErr(val)) return val;
      const vec = vecOf(look);
      const m = mode === undefined || mode === MISSING ? 0 : toNum(mode);
      const sm = search === undefined || search === MISSING ? 1 : toNum(search);
      const order = sm < 0 ? [...Array(vec.n).keys()].reverse() : [...Array(vec.n).keys()];
      let hit = -1, best = -1;
      const isStrWild = m === 2 && typeof val === 'string';
      const re = isStrWild ? wildcardRe(val) : null;
      for (const i of order) {
        const v = vec.get(i);
        if (isStrWild) { if (typeof v === 'string' && re.test(v)) { hit = i; break; } continue; }
        if (lookupEq(v, val) && !(isBlank(v) && !isBlank(val))) { hit = i; break; }
        if (m === -1 || m === 1) {
          const c = compare(v, val);
          if (isErr(c) || isBlank(v)) continue;
          if (m === -1 && c < 0 && (best < 0 || compare(v, vec.get(best)) > 0)) best = i;
          if (m === 1 && c > 0 && (best < 0 || compare(v, vec.get(best)) < 0)) best = i;
        }
      }
      if (hit < 0) hit = best;
      if (hit < 0) return notFound === undefined || notFound === MISSING ? E.NA : (notFound instanceof Ref ? notFound.value() : notFound);
      // return row/col of ret
      if (ret instanceof Ref) {
        if (vec.horizontal) return new Ref(ret.book, ret.sheet, ret.r1, ret.c1 + hit, ret.r2, ret.c1 + hit);
        return new Ref(ret.book, ret.sheet, ret.r1 + hit, ret.c1, ret.r1 + hit, ret.c2);
      }
      if (ret instanceof Matrix) {
        if (vec.horizontal) return indexInto(ret, 0, hit + 1);
        return indexInto(ret, hit + 1, 0);
      }
      return ret;
    },
  },
  ROW: {
    lazy: true,
    fn(a, ctx) {
      if (!a.length || a[0] === undefined) return ctx.row;
      const v = a[0](ctx);
      if (v === MISSING) return ctx.row;
      if (!(v instanceof Ref)) return E.VALUE;
      if (v.rows === 1) return v.r1;
      return new Matrix(v.rows, 1, Array.from({ length: v.rows }, (_, i) => v.r1 + i));
    },
  },
  COLUMN: {
    lazy: true,
    fn(a, ctx) {
      if (!a.length) return ctx.col;
      const v = a[0](ctx);
      if (v === MISSING) return ctx.col;
      if (!(v instanceof Ref)) return E.VALUE;
      if (v.cols === 1) return v.c1;
      return new Matrix(1, v.cols, Array.from({ length: v.cols }, (_, i) => v.c1 + i));
    },
  },
  ROWS: { args: 'r', lift: false, fn([v]) { if (v instanceof Ref || v instanceof Matrix) return v.rows; return isErr(v) ? v : 1; } },
  COLUMNS: { args: 'r', lift: false, fn([v]) { if (v instanceof Ref || v instanceof Matrix) return v.cols; return isErr(v) ? v : 1; } },
  OFFSET: {
    args: 'rvvvv', lift: false,
    fn([ref, rows, cols, h, w]) {
      if (!(ref instanceof Ref)) return isErr(ref) ? ref : E.VALUE;
      const dr = toNum(blankTo0(rows)), dc = toNum(blankTo0(cols));
      if (isErr(dr)) return dr; if (isErr(dc)) return dc;
      const H = h === undefined || h === MISSING ? ref.rows : toNum(h);
      const W = w === undefined || w === MISSING ? ref.cols : toNum(w);
      if (isErr(H)) return H; if (isErr(W)) return W;
      const r1 = ref.r1 + Math.trunc(dr), c1 = ref.c1 + Math.trunc(dc);
      if (r1 < 1 || c1 < 1 || H < 1 || W < 1) return E.REF;
      return new Ref(ref.book, ref.sheet, r1, c1, r1 + Math.trunc(H) - 1, c1 + Math.trunc(W) - 1);
    },
  },
  INDIRECT: {
    args: 'vv', lift: false,
    fn([text, a1], ctx, engine) {
      if (isErr(text)) return text;
      const s = toText(text).trim();
      const m = /^(?:'?((?:[^'!]|'')+)'?!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/.exec(s);
      if (!m) {
        const n = engine.lookupName(s, ctx.sheet);
        if (n) { if (!n.compiled) n.compiled = engine.compileText(n.f); return n.compiled.fn({ ...ctx, dr: 0, dc: 0 }); }
        return E.REF;
      }
      const si = m[1] ? engine.sheetIndex(m[1].replace(/''/g, "'")) : ctx.sheet;
      if (si < 0) return E.REF;
      const parts = m[2].replace(/\$/g, '').toUpperCase().split(':');
      const cell = (p) => { const mm = /^([A-Z]+)(\d+)$/.exec(p); let n = 0; for (const ch of mm[1]) n = n * 26 + ch.charCodeAt(0) - 64; return [+mm[2], n]; };
      const [r1, c1] = cell(parts[0]);
      const [r2, c2] = parts[1] ? cell(parts[1]) : [r1, c1];
      return new Ref(engine, si, Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2));
    },
  },
  TRANSPOSE: { args: 'a', lift: false, fn([m]) { m = toMatrix(m); const d = new Array(m.rows * m.cols); for (let i = 0; i < m.rows; i++) for (let j = 0; j < m.cols; j++) d[j * m.rows + i] = m.data[i * m.cols + j]; return new Matrix(m.cols, m.rows, d); } },

  // ---------- dynamic arrays
  FILTER: {
    args: 'aar', lift: false,
    fn([arr, inc, ifEmpty]) {
      if (isErr(arr)) return arr;
      if (isErr(inc)) return inc;
      const m = toMatrix(arr), f = toMatrix(inc);
      const rows = [];
      if (f.cols === 1 && f.rows === m.rows) {
        for (let i = 0; i < m.rows; i++) {
          const v = f.data[i];
          if (isErr(v)) return v;
          if (toBool(isBlank(v) ? false : v) === true) rows.push(m.data.slice(i * m.cols, (i + 1) * m.cols));
        }
        if (!rows.length) return ifEmpty === undefined || ifEmpty === MISSING ? E.CALC : first(ifEmpty);
        return Matrix.of(rows);
      }
      if (f.rows === 1 && f.cols === m.cols) {
        const keep = [];
        for (let j = 0; j < m.cols; j++) { const v = f.data[j]; if (isErr(v)) return v; if (toBool(isBlank(v) ? false : v) === true) keep.push(j); }
        if (!keep.length) return ifEmpty === undefined || ifEmpty === MISSING ? E.CALC : first(ifEmpty);
        for (let i = 0; i < m.rows; i++) rows.push(keep.map((j) => m.data[i * m.cols + j]));
        return Matrix.of(rows);
      }
      return E.VALUE;
    },
  },
  UNIQUE: {
    args: 'avv', lift: false,
    fn([arr, byCol, once]) {
      const m = toMatrix(arr);
      const bc = byCol !== undefined && byCol !== MISSING && toBool(byCol) === true;
      const eo = once !== undefined && once !== MISSING && toBool(once) === true;
      const items = [];
      if (!bc) for (let i = 0; i < m.rows; i++) items.push(m.data.slice(i * m.cols, (i + 1) * m.cols));
      else for (let j = 0; j < m.cols; j++) items.push(Array.from({ length: m.rows }, (_, i) => m.data[i * m.cols + j]));
      const key = (row) => row.map((v) => (typeof v === 'string' ? 's' + v.toLowerCase() : isBlank(v) ? 'b' : typeof v + String(v))).join('\u0001');
      const counts = new Map();
      for (const it of items) { const k = key(it); counts.set(k, (counts.get(k) || 0) + 1); }
      const seen = new Set(), out = [];
      for (const it of items) {
        const k = key(it);
        if (seen.has(k)) continue;
        seen.add(k);
        if (eo && counts.get(k) > 1) continue;
        out.push(it.map((v) => (isBlank(v) ? 0 : v)));
      }
      if (!out.length) return E.CALC;
      if (!bc) return Matrix.of(out);
      const t = Matrix.of(out);
      return FUNCS.TRANSPOSE.fn([t]);
    },
  },
  VSTACK: { args: 'r', lift: false, fn(args) { return stackMatrices(args, true); } },
  HSTACK: { args: 'r', lift: false, fn(args) { return stackMatrices(args, false); } },
  SORT: {
    args: 'avvv', lift: false,
    fn([arr, idx, ord, byCol]) {
      const m = toMatrix(arr);
      const k = idx === undefined || idx === MISSING ? 1 : toNum(idx);
      const o = ord === undefined || ord === MISSING ? 1 : toNum(ord);
      const rows = m.toRows();
      rows.sort((a, b) => { const c = compare(a[k - 1], b[k - 1]); return (isErr(c) ? 0 : c) * o; });
      return Matrix.of(rows);
    },
  },
  SEQUENCE: {
    args: 'vvvv', lift: false,
    fn([r, c, s, st]) {
      r = toNum(r); c = c === undefined || c === MISSING ? 1 : toNum(c); s = s === undefined || s === MISSING ? 1 : toNum(s); st = st === undefined || st === MISSING ? 1 : toNum(st);
      const d = []; for (let i = 0; i < r * c; i++) d.push(s + i * st);
      return new Matrix(r, c, d);
    },
  },
  TOCOL: { args: 'a', lift: false, fn([m]) { m = toMatrix(m); return new Matrix(m.data.length, 1, m.data.slice()); } },
  TOROW: { args: 'a', lift: false, fn([m]) { m = toMatrix(m); return new Matrix(1, m.data.length, m.data.slice()); } },
  TAKE: {
    args: 'avv', lift: false,
    fn([m, r, c]) {
      m = toMatrix(m); let rows = m.toRows();
      if (r !== undefined && r !== MISSING) { r = toNum(r); rows = r >= 0 ? rows.slice(0, r) : rows.slice(r); }
      if (c !== undefined && c !== MISSING) { c = toNum(c); rows = rows.map((x) => (c >= 0 ? x.slice(0, c) : x.slice(c))); }
      return Matrix.of(rows);
    },
  },
  DROP: {
    args: 'avv', lift: false,
    fn([m, r, c]) {
      m = toMatrix(m); let rows = m.toRows();
      if (r !== undefined && r !== MISSING) { r = toNum(r); rows = r >= 0 ? rows.slice(r) : rows.slice(0, r); }
      if (c !== undefined && c !== MISSING) { c = toNum(c); rows = rows.map((x) => (c >= 0 ? x.slice(c) : x.slice(0, c))); }
      return Matrix.of(rows);
    },
  },
  CHOOSECOLS: { args: 'av', lift: false, fn([m, ...cols]) { m = toMatrix(m); const rows = m.toRows().map((r) => cols.map((c) => r[toNum(c) - 1])); return Matrix.of(rows); } },
  CHOOSEROWS: { args: 'av', lift: false, fn([m, ...rs]) { m = toMatrix(m); const rows = m.toRows(); return Matrix.of(rs.map((r) => rows[toNum(r) - 1])); } },

  // ---------- date
  TODAY: { args: '', fn() { const d = new Date(); return dateSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); } },
  NOW: { args: '', fn() { const d = new Date(); return dateSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; } },
  DATE: { args: 'vvv', fn([y, m, d]) { y = toNum(y); m = toNum(m); d = toNum(d); if (isErr(y)) return y; if (y < 1900) y += 1900; return dateSerial(y, m, d); } },
  YEAR: { args: 'v', fn([v]) { v = toNum(v); return isErr(v) ? v : serialToDate(v).getUTCFullYear(); } },
  MONTH: { args: 'v', fn([v]) { v = toNum(v); return isErr(v) ? v : serialToDate(v).getUTCMonth() + 1; } },
  DAY: { args: 'v', fn([v]) { v = toNum(v); return isErr(v) ? v : serialToDate(v).getUTCDate(); } },

  // ---------- linked data types
  _FV: {
    args: 'vv', lift: false,
    fn([v, field]) {
      if (!(v instanceof Rich)) return E.FIELD || err('#FIELD!');
      const f = toText(field);
      const raw = v.fields[f];
      if (raw === undefined) return err('#FIELD!');
      const n = parseNumber(raw);
      return n === null ? raw : n;
    },
  },
};

FUNCS['CONCATENATE'] = FUNCS.CONCATENATE;
FUNCS['STDEV'] = {
  args: 'r', fn(args) { const ns = numArgs(args, { skipMissing: true }); if (isErr(ns)) return ns; if (ns.length < 2) return E.DIV0; const m = ns.reduce((a, b) => a + b, 0) / ns.length; return Math.sqrt(ns.reduce((a, b) => a + (b - m) ** 2, 0) / (ns.length - 1)); },
};
FUNCS['STDEV.S'] = FUNCS.STDEV;
FUNCS['RANK'] = { args: 'vrv', lift: false, fn([v, ref, ord]) { const ns = numArgs([ref], { skipMissing: true }); v = toNum(v); const o = ord === undefined ? 0 : toNum(ord); let r = 1; for (const n of ns) if (o ? n < v : n > v) r++; return ns.includes(v) ? r : E.NA; } };
FUNCS['RANK.EQ'] = FUNCS.RANK;
