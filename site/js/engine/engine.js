// Spreadsheet calculation engine: compiles Excel formulas to closures, tracks
// dependencies and recalculates only what changed.

import { parse, parseStructSpec, walk, colToNum, numToCol } from './parser.js';
import {
  E, err, isErr, XLError, Matrix, Ref, Rich, MISSING, toNum, toText, compare, isBlank, IndexSet, intersectSorted, unionSorted,
} from './values.js';
import { FUNCS } from './functions.js';

const KEY = (r, c) => r * 16384 + c;

export class Sheet {
  constructor(index, name) {
    this.index = index;
    this.name = name;
    this.vals = [];          // vals[row][col]
    this.formulas = new Map(); // key -> FormulaCell
    this.maxRow = 0;
    this.maxCol = 0;
    this.version = 0;
  }
  get(r, c) { const row = this.vals[r]; if (!row) return null; const v = row[c]; return v === undefined ? null : v; }
  set(r, c, v) {
    let row = this.vals[r];
    if (!row) { row = this.vals[r] = []; }
    row[c] = v;
    this.version++;
    if (r > this.maxRow) this.maxRow = r;
    if (c > this.maxCol) this.maxCol = c;
  }
}

export class FormulaCell {
  constructor(id, sheet, r, c, text, r0, c0, kind, spill) {
    this.id = id; this.sheet = sheet; this.r = r; this.c = c;
    this.text = text;   // formula text (master text for shared formulas)
    this.r0 = r0; this.c0 = c0; // anchor of text (shared formula master)
    this.kind = kind;   // '' normal, 'a' CSE array, 'd' dynamic array
    this.spill = spill; // [r1,c1,r2,c2] for array formulas
    this.lastSpill = null;
    this.order = 0;
  }
}

function excelRound(x, d) { const m = Math.pow(10, d); return Math.sign(x) * Math.round(+(Math.abs(x) * m).toPrecision(15)) / m; }

export class Engine {
  constructor() {
    this.sheets = [];
    this.sheetByName = new Map();
    this.names = new Map();     // 'name' or 'sheetIdx!name' -> {f, compiled}
    this.tables = new Map();    // lower name -> table
    this.external = {};
    this.compiled = new Map();  // text -> {fn, refs}
    this.cells = [];            // FormulaCell by id
    this.epoch = 1;
    this.volatile = [];
    this.rich = [];
    this.tracing = false;
    this.matCache = new Map();
    this.colIndexCache = new Map();
  }

  // -------------------------------------------------------------- loading
  addSheet(name) {
    const s = new Sheet(this.sheets.length, name);
    this.sheets.push(s);
    this.sheetByName.set(name.toLowerCase(), s.index);
    return s;
  }
  sheetIndex(name) {
    const i = this.sheetByName.get(name.toLowerCase());
    return i === undefined ? -1 : i;
  }
  addName(name, formula, sheetIdx) {
    const key = (sheetIdx === undefined || sheetIdx === null ? '' : sheetIdx + '!') + name.toLowerCase();
    this.names.set(key, { name, f: formula, sheet: sheetIdx, compiled: null });
  }
  addTable(t) {
    this.tables.set(t.name.toLowerCase(), {
      ...t, sheetIdx: this.sheetIndex(t.sheet),
      colIndex: new Map(t.columns.map((c, i) => [normColName(c), i])),
    });
  }
  addFormula(sheetIdx, r, c, text, r0, c0, kind, spill) {
    const fc = new FormulaCell(this.cells.length, sheetIdx, r, c, text, r0, c0, kind || '', spill || null);
    this.cells.push(fc);
    this.sheets[sheetIdx].formulas.set(KEY(r, c), fc);
    return fc;
  }

  getValue(si, r, c) {
    const row = this.sheets[si].vals[r];
    if (!row) return null;
    const v = row[c];
    return v === undefined ? null : v;
  }

  // -------------------------------------------------------------- compilation
  compileText(text) {
    let c = this.compiled.get(text);
    if (c) return c;
    let ast;
    try { ast = parse(text); } catch (e) { ast = { t: 'err', v: '#NAME?', parseError: e.message }; }
    c = { ast, fn: null };
    markInvariant(ast);
    c.fn = this.compileNode(ast);
    this.compiled.set(text, c);
    return c;
  }

  compileNode(node) {
    const self = this;
    switch (node.t) {
      case 'num': { const v = node.v; return () => v; }
      case 'str': { const v = node.v; return () => v; }
      case 'bool': { const v = node.v; return () => v; }
      case 'err': { const v = err(node.v); return () => v; }
      case 'missing': return () => MISSING;
      case 'arr': {
        const m = Matrix.of(node.rows.map((r) => r.map((x) => (x && x.e ? err(x.e) : x))));
        return () => m;
      }
      case 'paren': return this.compileNode(node.a);
      case 'implicit': {
        const a = this.compileNode(node.a);
        return (ctx) => { const v = a(ctx); return v instanceof Ref ? intersect(v, ctx) : v instanceof Matrix ? v.get(0, 0) : v; };
      }
      case 'spill': {
        const a = this.compileNode(node.a);
        return (ctx) => {
          const v = a(ctx);
          if (!(v instanceof Ref)) return E.REF;
          const fc = self.sheets[v.sheet].formulas.get(KEY(v.r1, v.c1));
          if (!fc || !fc.lastSpill) return v;
          const s = fc.lastSpill;
          return new Ref(self, v.sheet, s[0], s[1], s[2], s[3]);
        };
      }
      case 'ref': return this.compileRef(node);
      case 'rangeop': {
        const a = this.compileNode(node.a), b = this.compileNode(node.b);
        return (ctx) => {
          const x = a(ctx), y = b(ctx);
          if (isErr(x)) return x;
          if (isErr(y)) return y;
          if (!(x instanceof Ref) || !(y instanceof Ref) || x.sheet !== y.sheet) return E.VALUE;
          return new Ref(self, x.sheet, Math.min(x.r1, y.r1), Math.min(x.c1, y.c1), Math.max(x.r2, y.r2), Math.max(x.c2, y.c2));
        };
      }
      case 'union': {
        const items = node.items.map((x) => this.compileNode(x));
        return (ctx) => { const refs = items.map((f) => f(ctx)); const u = refs; u.isUnion = true; return u; };
      }
      case 'name': return this.compileName(node);
      case 'struct': return this.compileStruct(node);
      case 'neg': {
        const a = this.compileNode(node.a);
        return (ctx) => unary(a(ctx), ctx, (v) => { const n = toNum(v); return isErr(n) ? n : -n; });
      }
      case 'pct': {
        const a = this.compileNode(node.a);
        return (ctx) => unary(a(ctx), ctx, (v) => { const n = toNum(v); return isErr(n) ? n : n / 100; });
      }
      case 'op': {
        const a = this.compileNode(node.a), b = this.compileNode(node.b);
        const op = OPS[node.op];
        let fn;
        if (node.op === '=') fn = (ctx) => { const x = a(ctx), y = b(ctx); return eqFast(self, x, y, ctx) || binary(x, y, ctx, op); };
        else if (node.op === '*') fn = (ctx) => { const x = a(ctx), y = b(ctx); return mulFast(x, y) || binary(x, y, ctx, op); };
        else fn = (ctx) => binary(a(ctx), b(ctx), ctx, op);
        return node.inv ? this.cacheWrap(fn) : fn;
      }
      case 'fn': {
        const fn = this.compileCall(node);
        return node.inv ? this.cacheWrap(fn) : fn;
      }
      default:
        return () => E.NAME;
    }
  }

  // cache results of nodes that do not depend on the formula's position, per recalculation pass
  cacheWrap(fn) {
    const self = this;
    let ep = -1, sh = -1, arr = null, val;
    return (ctx) => {
      if (ctx.lets) return fn(ctx);
      if (ep === self.epoch && sh === ctx.sheet && arr === ctx.array) return val;
      val = fn(ctx); ep = self.epoch; sh = ctx.sheet; arr = ctx.array;
      return val;
    };
  }

  compileRef(node) {
    const self = this;
    if (node.ext) {
      const book = this.external[node.ext];
      return () => {
        if (!book) return E.REF;
        const sheet = book.data[node.sheet] || {};
        if (node.r1 !== node.r2 || node.c1 !== node.c2) return E.REF;
        const v = sheet[numToCol(node.c1) + node.r1];
        return v === undefined ? null : v && v.e ? err(v.e) : v;
      };
    }
    let si = null;
    if (node.sheet !== null && node.sheet !== undefined) {
      si = this.sheetIndex(node.sheet);
      if (si < 0) return () => E.REF;
    }
    const { r1, c1, r2, c2, ar1, ac1, ar2, ac2 } = node;
    if (ar1 && ac1 && ar2 && ac2) {
      return (ctx) => new Ref(self, si === null ? ctx.sheet : si, r1, c1, r2, c2);
    }
    return (ctx) => {
      const R1 = ar1 ? r1 : r1 + ctx.dr, C1 = ac1 ? c1 : c1 + ctx.dc;
      const R2 = ar2 ? r2 : r2 + ctx.dr, C2 = ac2 ? c2 : c2 + ctx.dc;
      if (R1 < 1 || C1 < 1 || R2 < 1 || C2 < 1 || R1 > 1048576 || R2 > 1048576 || C1 > 16384 || C2 > 16384) return E.REF;
      return new Ref(self, si === null ? ctx.sheet : si, Math.min(R1, R2), Math.min(C1, C2), Math.max(R1, R2), Math.max(C1, C2));
    };
  }

  lookupName(name, sheetCtx) {
    const lower = name.toLowerCase();
    let n = this.names.get(sheetCtx + '!' + lower);
    if (!n) n = this.names.get(lower);
    return n;
  }

  compileName(node) {
    const self = this;
    const lname = node.name.toLowerCase();
    const explicitSheet = node.sheet ? this.sheetIndex(node.sheet) : null;
    return (ctx) => {
      if (ctx.lets && ctx.lets.has(lname)) return ctx.lets.get(lname);
      const n = self.lookupName(node.name, explicitSheet !== null ? explicitSheet : ctx.sheet);
      if (!n) return E.NAME;
      if (!n.compiled) {
        n.compiled = self.compileText(n.f);
      }
      // names are evaluated with no offset, in their own scope sheet if local
      const c2 = { sheet: n.sheet !== undefined && n.sheet !== null ? n.sheet : ctx.sheet, row: ctx.row, col: ctx.col, dr: 0, dc: 0, array: ctx.array, lets: null, fc: ctx.fc };
      return n.compiled.fn(c2);
    };
  }

  resolveStruct(tableName, spec, row) {
    const t = this.tables.get(tableName.toLowerCase());
    if (!t) return E.REF;
    const ps = parseStructSpec(spec);
    const [tr1, tc1, tr2, tc2] = t.ref;
    let r1 = tr1 + t.header, r2 = tr2 - t.totals;
    const items = ps.items.map((x) => x.toLowerCase());
    if (items.includes('#all')) { r1 = tr1; r2 = tr2; }
    else if (items.includes('#headers') && items.includes('#data')) { r1 = tr1; }
    else if (items.includes('#data') && items.includes('#totals')) { r2 = tr2; }
    else if (items.includes('#headers')) { r1 = r2 = tr1; }
    else if (items.includes('#totals')) { r1 = r2 = tr2; }
    else if (items.includes('#this row')) {
      if (row < tr1 + t.header || row > tr2 - t.totals) return E.VALUE;
      r1 = r2 = row;
    }
    let c1 = tc1, c2 = tc2;
    if (ps.cols) {
      const a = t.colIndex.get(normColName(ps.cols[0])), b = t.colIndex.get(normColName(ps.cols[1]));
      if (a === undefined || b === undefined) return E.REF;
      c1 = tc1 + Math.min(a, b); c2 = tc1 + Math.max(a, b);
    }
    return new Ref(this, t.sheetIdx, r1, c1, r2, c2);
  }

  compileStruct(node) {
    const self = this;
    const thisRow = /#this row|@/i.test(node.spec);
    if (!thisRow) {
      let cached = null;
      return () => { if (!cached) cached = self.resolveStruct(node.table, node.spec, 0); return cached; };
    }
    return (ctx) => self.resolveStruct(node.table, node.spec, ctx.row + 0);
  }

  compileCall(node) {
    const self = this;
    const name = node.name;
    if (name === 'LET') return this.compileLet(node);
    const spec = FUNCS[name];
    if (!spec) return () => E.NAME;
    const argFns = node.args.map((a) => this.compileNode(a));
    if (spec.lazy) {
      return (ctx) => spec.fn(argFns, ctx, self);
    }
    const kinds = spec.args || '';
    const rest = kinds.length ? kinds[kinds.length - 1] : 'v';
    const lift = spec.lift !== false;
    const n = argFns.length;
    return (ctx) => {
      const vals = new Array(n);
      let liftArr = null;
      for (let i = 0; i < n; i++) {
        const k = i < kinds.length ? kinds[i] : rest;
        let v = argFns[i](ctx);
        if (k === 'v') {
          if (v instanceof Ref) v = v.isCell() ? v.value() : ctx.array ? v.toMatrix() : intersect(v, ctx);
          if (v instanceof Matrix) {
            if (v.rows === 1 && v.cols === 1) v = v.data[0];
            else if (lift) liftArr = liftArr ? maxDims(liftArr, v) : v;
          }
        } else if (k === 'a') {
          if (v instanceof Ref) v = v.toMatrix();
        }
        vals[i] = v;
      }
      if (liftArr) {
        const R = liftArr.rows, C = liftArr.cols;
        const out = new Array(R * C);
        const a2 = vals.slice();
        for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) {
          for (let x = 0; x < n; x++) {
            const k = x < kinds.length ? kinds[x] : rest;
            a2[x] = k === 'v' && vals[x] instanceof Matrix ? vals[x].get(i, j) : vals[x];
          }
          let r = spec.fn(a2, ctx, self);
          if (r instanceof Ref) r = r.isCell() ? r.value() : r.toMatrix().get(0, 0);
          if (r instanceof Matrix) r = r.get(0, 0);
          out[i * C + j] = r;
        }
        return new Matrix(R, C, out);
      }
      return spec.fn(vals, ctx, self);
    };
  }

  compileLet(node) {
    const args = node.args;
    const pairs = [];
    for (let i = 0; i + 1 < args.length; i += 2) {
      if (args[i].t !== 'name') return () => E.VALUE;
      pairs.push([args[i].name.toLowerCase(), this.compileNode(args[i + 1])]);
    }
    const body = this.compileNode(args[args.length - 1]);
    return (ctx) => {
      const lets = new Map(ctx.lets || []);
      const c2 = { ...ctx, lets };
      for (const [nm, f] of pairs) lets.set(nm, f(c2));
      return body(c2);
    };
  }

  // -------------------------------------------------------------- evaluation
  evalCell(fc) {
    if (fc.dead) return [];
    const comp = this.compileText(fc.text);
    const ctx = { sheet: fc.sheet, row: fc.r, col: fc.c, dr: fc.r - fc.r0, dc: fc.c - fc.c0, array: fc.kind !== '', lets: null, fc };
    let res;
    try { res = comp.fn(ctx); } catch (e) { if (this.tracing) console.error('eval error', this.sheets[fc.sheet].name, fc.r, fc.c, e); res = E.VALUE; }
    const sh = this.sheets[fc.sheet];
    const changed = [];
    if (fc.kind === '' || !fc.spill || (fc.spill[0] === fc.spill[2] && fc.spill[1] === fc.spill[3] && fc.kind === 'a')) {
      let v = finalValue(res, ctx);
      if (fc.kind === 'd') return this.spillDynamic(fc, res, ctx);
      const old = sh.get(fc.r, fc.c);
      if (!same(old, v)) { sh.set(fc.r, fc.c, v); changed.push(fc.r, fc.c); }
      return changed;
    }
    if (fc.kind === 'd') return this.spillDynamic(fc, res, ctx);
    // CSE multi-cell array formula
    const [r1, c1, r2, c2] = fc.spill;
    let m = res instanceof Ref ? res.toMatrix() : res;
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      let v;
      if (m instanceof Matrix) {
        const i = r - r1, j = c - c1;
        if ((m.rows === 1 || i < m.rows) && (m.cols === 1 || j < m.cols)) v = m.get(i, j); else v = E.NA;
      } else v = m;
      v = v === null || v === undefined || v === MISSING ? 0 : v;
      const old = sh.get(r, c);
      if (!same(old, v)) { sh.set(r, c, v); changed.push(r, c); }
    }
    return changed;
  }

  spillDynamic(fc, res, ctx) {
    const sh = this.sheets[fc.sheet];
    const changed = [];
    let m = res instanceof Ref ? (res.isCell() ? res.value() : res.toMatrix()) : res;
    if (!(m instanceof Matrix)) m = new Matrix(1, 1, [m]);
    // clear previous spill
    const prev = fc.lastSpill || fc.spill || [fc.r, fc.c, fc.r, fc.c];
    const R = m.rows, C = m.cols;
    const next = [fc.r, fc.c, fc.r + R - 1, fc.c + C - 1];
    // blocked spill?
    let blocked = false;
    for (let i = 0; i < R && !blocked; i++) for (let j = 0; j < C; j++) {
      if (i === 0 && j === 0) continue;
      const k = KEY(fc.r + i, fc.c + j);
      const other = sh.formulas.get(k);
      if (other && other !== fc) { blocked = true; break; }
    }
    if (blocked) { m = new Matrix(1, 1, [E.SPILL]); next[2] = fc.r; next[3] = fc.c; }
    for (let r = prev[0]; r <= prev[2]; r++) for (let c = prev[1]; c <= prev[3]; c++) {
      if (r <= next[2] && c <= next[3] && r >= next[0] && c >= next[1]) continue;
      if (sh.formulas.get(KEY(r, c)) && !(r === fc.r && c === fc.c)) continue;
      if (sh.get(r, c) !== null) { sh.set(r, c, null); changed.push(r, c); }
    }
    for (let i = 0; i <= next[2] - next[0]; i++) for (let j = 0; j <= next[3] - next[1]; j++) {
      let v = m.get(i, j);
      v = v === null || v === undefined || v === MISSING ? 0 : v;
      if (v instanceof Ref) v = v.value();
      const r = fc.r + i, c = fc.c + j;
      if (!same(sh.get(r, c), v)) { sh.set(r, c, v); changed.push(r, c); }
    }
    fc.lastSpill = next;
    return changed;
  }

  // Evaluate an arbitrary formula in the context of a cell (used for validation lists,
  // conditional formats, chart series).
  evaluate(text, si, r = 1, c = 1, array = true) {
    const comp = this.compileText(text);
    const ctx = { sheet: si, row: r, col: c, dr: 0, dc: 0, array, lets: null };
    try { return comp.fn(ctx); } catch (e) { return E.VALUE; }
  }
  // Evaluate with relative offset (conditional formats are relative to top-left of range)
  evaluateRel(text, si, r0, c0, r, c, array = false) {
    const comp = this.compileText(text);
    const ctx = { sheet: si, row: r, col: c, dr: r - r0, dc: c - c0, array, lets: null };
    try { return comp.fn(ctx); } catch (e) { return E.VALUE; }
  }

  // -------------------------------------------------------------- dependency analysis
  // Precedent ranges of a formula cell: [[sheet,r1,c1,r2,c2], ...]; volatile flag
  precedents(fc) {
    const comp = this.compileText(fc.text);
    const out = [];
    let vol = false;
    const dr = fc.r - fc.r0, dc = fc.c - fc.c0;
    const visitNames = new Set();
    const visit = (ast, sheetCtx, odr, odc, rowCtx) => {
      walk(ast, (n) => {
        if (n.t === 'ref') {
          if (n.ext) return;
          const si = n.sheet ? this.sheetIndex(n.sheet) : sheetCtx;
          if (si < 0) return;
          const R1 = n.ar1 ? n.r1 : n.r1 + odr, C1 = n.ac1 ? n.c1 : n.c1 + odc;
          const R2 = n.ar2 ? n.r2 : n.r2 + odr, C2 = n.ac2 ? n.c2 : n.c2 + odc;
          out.push([si, Math.min(R1, R2), Math.min(C1, C2), Math.max(R1, R2), Math.max(C1, C2)]);
        } else if (n.t === 'struct') {
          const ref = this.resolveStruct(n.table, n.spec, rowCtx);
          if (ref instanceof Ref) out.push([ref.sheet, ref.r1, ref.c1, ref.r2, ref.c2]);
          else {
            // e.g. #This Row evaluated outside the table: depend on whole column
            const t = this.tables.get(n.table.toLowerCase());
            if (t) out.push([t.sheetIdx, t.ref[0], t.ref[1], t.ref[2], t.ref[3]]);
          }
        } else if (n.t === 'name') {
          const nm = this.lookupName(n.name, sheetCtx);
          if (nm && !visitNames.has(nm)) {
            visitNames.add(nm);
            const c = this.compileText(nm.f);
            visit(c.ast, nm.sheet !== undefined && nm.sheet !== null ? nm.sheet : sheetCtx, 0, 0, rowCtx);
          }
        } else if (n.t === 'fn') {
          if (n.name === 'OFFSET' || n.name === 'INDIRECT' || n.name === 'TODAY' || n.name === 'NOW' || n.name === 'RAND' || n.name === 'RANDBETWEEN' || n.name === 'CELL' || n.name === 'INFO') vol = true;
        }
      });
    };
    visit(comp.ast, fc.sheet, dr, dc, fc.r);
    return { ranges: out, volatile: vol };
  }

  // Build dependency graph and topological order over all formula cells
  buildGraph() {
    const t0 = Date.now();
    const N = this.cells.length;
    // owner map: cell -> formula id (anchor or spill owner)
    const owner = this.sheets.map(() => new Map());
    // per sheet, per column sorted rows of formula-owned cells
    const colIdx = this.sheets.map(() => new Map());
    const addOwned = (si, r, c, id) => {
      owner[si].set(KEY(r, c), id);
      let col = colIdx[si].get(c);
      if (!col) colIdx[si].set(c, col = []);
      col.push(r);
    };
    for (const fc of this.cells) {
      if (fc.spill && fc.kind) {
        const [r1, c1, r2, c2] = fc.spill;
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
          const f2 = this.sheets[fc.sheet].formulas.get(KEY(r, c));
          if (f2 && f2 !== fc) continue;
          addOwned(fc.sheet, r, c, fc.id);
        }
      } else addOwned(fc.sheet, fc.r, fc.c, fc.id);
    }
    for (const m of colIdx) for (const arr of m.values()) arr.sort((a, b) => a - b);

    // Nodes 0..N-1 are formula cells; nodes >= N are shared "range" nodes standing for
    // large precedent ranges (so 800 formulas reading a 55k-row column cost 55k+800 edges).
    const prec = [];                 // node -> precedent node ids
    const inputRanges = this.sheets.map(() => []); // [r1,c1,r2,c2,id]
    const rangeNodes = new Map();    // key -> node id
    this.volatile = [];
    let edges = 0;
    const ownersIn = (si, r1, c1, r2, c2, self, set) => {
      const cols = colIdx[si];
      for (let c = c1; c <= Math.min(c2, this.sheets[si].maxCol + 1); c++) {
        const rows = cols.get(c);
        if (!rows) continue;
        let lo = 0, hi = rows.length;
        while (lo < hi) { const mid = (lo + hi) >> 1; if (rows[mid] < r1) lo = mid + 1; else hi = mid; }
        for (let k = lo; k < rows.length && rows[k] <= r2; k++) {
          const id = owner[si].get(KEY(rows[k], c));
          if (id !== undefined && id !== self) set.add(id);
        }
      }
    };
    for (const fc of this.cells) {
      const { ranges, volatile } = this.precedents(fc);
      if (volatile) this.volatile.push(fc.id);
      const set = new Set();
      for (const [si, r1, c1, r2, c2] of ranges) {
        inputRanges[si].push([r1, c1, r2, c2, fc.id]);
        if ((c2 - c1 + 1) * (r2 - r1 + 1) <= 32) {
          for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
            const id = owner[si].get(KEY(r, c));
            if (id !== undefined && id !== fc.id) set.add(id);
          }
        } else {
          const key = si + ':' + r1 + ':' + c1 + ':' + r2 + ':' + c2;
          let rn = rangeNodes.get(key);
          if (rn === undefined) {
            rn = N + rangeNodes.size;
            rangeNodes.set(key, rn);
            const rs = new Set();
            ownersIn(si, r1, c1, r2, c2, -1, rs);
            prec[rn] = Array.from(rs);
            edges += rs.size;
          }
          if (prec[rn].length === 1) { if (prec[rn][0] !== fc.id) set.add(prec[rn][0]); }
          else if (prec[rn].length > 1) set.add(rn);
        }
      }
      prec[fc.id] = Array.from(set);
      edges += set.size;
    }
    const T = N + rangeNodes.size;
    for (let i = 0; i < T; i++) if (!prec[i]) prec[i] = [];
    // a range node containing the formula that reads it must not create a self-cycle
    // dependents
    const deps = Array.from({ length: T }, () => []);
    for (let i = 0; i < T; i++) for (const p of prec[i]) deps[p].push(i);
    // topological order (Kahn); cycles broken at the node with the fewest pending inputs
    const indeg = new Int32Array(T);
    for (let i = 0; i < T; i++) indeg[i] = prec[i].length;
    const order = [];
    const queue = [];
    for (let i = 0; i < T; i++) if (indeg[i] === 0) queue.push(i);
    let qh = 0;
    const done = new Uint8Array(T);
    let ndone = 0;
    for (;;) {
      while (qh < queue.length) {
        const i = queue[qh++];
        if (done[i]) continue;
        done[i] = 1; ndone++;
        if (i < N) order.push(i);
        for (const d of deps[i]) if (--indeg[d] === 0) queue.push(d);
      }
      if (ndone >= T) break;
      let best = -1;
      for (let i = 0; i < T; i++) if (!done[i] && (best < 0 || indeg[i] < indeg[best])) best = i;
      indeg[best] = 0; queue.push(best);
      this.cycles = (this.cycles || 0) + 1;
    }
    order.forEach((id, k) => { this.cells[id].order = k; });
    this.order = order;
    this.deps = deps;
    this.depsOf = (id) => deps[id];
    this.prec = prec;
    this.owner = owner;
    // input index: per sheet, per column bucket lists of [r1,r2,c1,c2,id]
    this.inputIdx = inputRanges.map((list) => {
      const m = new Map();
      for (const x of list) {
        const [r1, c1, r2, c2, id] = x;
        const cs = Math.min(c2, 1024);
        for (let c = c1; c <= cs; c++) {
          let a = m.get(c); if (!a) m.set(c, a = []);
          a.push(r1, r2, id);
        }
      }
      return m;
    });
    this.graphStats = { formulas: N, rangeNodes: rangeNodes.size, edges, ms: Date.now() - t0, cycles: this.cycles || 0, volatile: this.volatile.length };
    this.nodeCount = T;
    return this.graphStats;
  }

  // formulas that directly reference cell (si,r,c)
  directDependents(si, r, c) {
    const out = [];
    const a = this.inputIdx[si].get(c);
    if (a) for (let k = 0; k < a.length; k += 3) if (r >= a[k] && r <= a[k + 1]) out.push(a[k + 2]);
    return out;
  }

  // -------------------------------------------------------------- recalculation
  recalcAll() {
    this.epoch++;
    for (const id of this.order) this.evalCell(this.cells[id]);
  }

  // Set input values: [[si,r,c,value], ...]; returns list of changed cells [[si,r,c],...]
  // With force, a formula cell is replaced by the value (like a VBA copy/paste-values).
  setValues(list, force = false) {
    const changedCells = [];
    const dirty = new Set();
    for (const [si, r, c, v] of list) {
      const sh = this.sheets[si];
      const fc = sh.formulas.get(KEY(r, c));
      if (fc) {
        if (!force || fc.kind) continue; // formulas are not overwritten
        fc.dead = true;
        sh.formulas.delete(KEY(r, c));
      }
      const old = sh.get(r, c);
      if (same(old, v)) continue;
      sh.set(r, c, v);
      this.epoch++;
      changedCells.push([si, r, c]);
      for (const id of this.directDependents(si, r, c)) dirty.add(id);
    }
    if (!dirty.size && !changedCells.length) return changedCells;
    for (const id of this.volatile) dirty.add(id);
    this.propagate(dirty, changedCells);
    return changedCells;
  }

  propagate(dirty, changedCells) {
    this.epoch++;
    // close over dependents
    const stack = Array.from(dirty);
    while (stack.length) {
      const id = stack.pop();
      for (const d of this.depsOf(id)) if (!dirty.has(d)) { dirty.add(d); stack.push(d); }
    }
    const N = this.cells.length;
    const list = Array.from(dirty).filter((id) => id < N).sort((a, b) => this.cells[a].order - this.cells[b].order);
    for (const id of list) {
      const fc = this.cells[id];
      const ch = this.evalCell(fc);
      for (let k = 0; k < ch.length; k += 2) changedCells.push([fc.sheet, ch[k], ch[k + 1]]);
    }
    this.lastRecalc = { formulas: list.length };
  }
}

// ------------------------------------------------------------------ helpers
function normColName(s) { return String(s).replace(/_x000a_/gi, '').replace(/\s+/g, '').toLowerCase(); }

function same(a, b) {
  if (a === b) return true;
  if (a === null || a === undefined) return b === null || b === undefined;
  if (typeof a === 'number' && typeof b === 'number') return Number.isNaN(a) && Number.isNaN(b);
  return false;
}

export function intersect(ref, ctx) {
  if (ref.isCell()) return ref.value();
  if (ctx.fc === undefined && ctx.row === undefined) return ref.value();
  const row = ctx.row, col = ctx.col;
  if (ref.c1 === ref.c2) {
    if (row >= ref.r1 && row <= ref.r2) return ref.book.getValue(ref.sheet, row, ref.c1);
    return E.VALUE;
  }
  if (ref.r1 === ref.r2) {
    if (col >= ref.c1 && col <= ref.c2) return ref.book.getValue(ref.sheet, ref.r1, col);
    return E.VALUE;
  }
  if (row >= ref.r1 && row <= ref.r2 && col >= ref.c1 && col <= ref.c2) return ref.book.getValue(ref.sheet, row, col);
  return E.VALUE;
}

function finalValue(res, ctx) {
  if (res instanceof Ref) res = ctx.array ? (res.isCell() ? res.value() : res.value(0, 0)) : intersect(res, ctx);
  if (res instanceof Matrix) res = res.get(0, 0);
  if (Array.isArray(res)) return E.VALUE;
  if (res === null || res === undefined || res === MISSING) return 0;
  if (typeof res === 'number' && !isFinite(res)) return E.NUM;
  return res;
}

function maxDims(a, b) {
  const R = Math.max(a.rows, b.rows), C = Math.max(a.cols, b.cols);
  if (R === a.rows && C === a.cols) return a;
  if (R === b.rows && C === b.cols) return b;
  return { rows: R, cols: C };
}

// value -> sorted row indices index for a large single-column reference
function columnIndex(book, ref) {
  const sh = book.sheets[ref.sheet];
  const key = ref.sheet + ':' + ref.r1 + ':' + ref.c1 + ':' + ref.r2;
  const hit = book.colIndexCache.get(key);
  if (hit && hit.version === sh.version) return hit.idx;
  const idx = new Map();
  const blanks = [];
  for (let r = ref.r1; r <= ref.r2; r++) {
    const row = sh.vals[r];
    let v = row ? row[ref.c1] : undefined;
    const i = r - ref.r1;
    if (v === undefined || v === null) { blanks.push(i); continue; }
    let k;
    if (typeof v === 'number') k = 'n' + v;
    else if (typeof v === 'string') k = 's' + v.toLowerCase();
    else if (typeof v === 'boolean') k = 'b' + v;
    else if (v instanceof Rich) k = 's' + v.toString().toLowerCase();
    else continue; // errors are not indexed
    let l = idx.get(k); if (!l) idx.set(k, l = []); l.push(i);
  }
  const res = { map: idx, blanks, hasErr: false, memo: new Map() };
  for (let r = ref.r1; r <= ref.r2; r++) { const row = sh.vals[r]; const v = row ? row[ref.c1] : undefined; if (v instanceof XLError) { res.hasErr = true; break; } }
  book.colIndexCache.set(key, { version: sh.version, idx: res });
  return res;
}

function eqFast(book, x, y, ctx) {
  if (!ctx.array) return null;
  let ref = null, s;
  if (x instanceof Ref && x.cols === 1 && x.rows >= 256) { ref = x; s = y; }
  else if (y instanceof Ref && y.cols === 1 && y.rows >= 256) { ref = y; s = x; }
  else return null;
  if (s instanceof Ref) { if (!s.isCell()) return null; s = s.value(); }
  if (s instanceof Matrix || Array.isArray(s) || s instanceof XLError) return null;
  const ci = columnIndex(book, ref);
  if (ci.hasErr) return null;
  const mk = typeof s + ':' + (typeof s === 'string' ? s.toLowerCase() : String(s));
  const memo = ci.memo.get(mk);
  if (memo) return new IndexSet(ref.rows, memo, false);
  let lists;
  if (s === null || s === undefined || s === MISSING) lists = [ci.blanks, ci.map.get('n0') || [], ci.map.get('s') || [], ci.map.get('bfalse') || []];
  else if (typeof s === 'number') { if (!Number.isInteger(s) && String(s).length > 15) return null; lists = [ci.map.get('n' + s) || []]; if (s === 0) lists.push(ci.blanks); }
  else if (typeof s === 'string') { lists = [ci.map.get('s' + s.toLowerCase()) || []]; if (s === '') lists.push(ci.blanks); }
  else if (typeof s === 'boolean') { lists = [ci.map.get('b' + s) || []]; if (!s) lists.push(ci.blanks); }
  else return null;
  const idx = lists.length === 1 ? lists[0] : unionSorted(lists);
  ci.memo.set(mk, idx);
  return new IndexSet(ref.rows, idx, false);
}

function mulFast(x, y) {
  const ix = x instanceof IndexSet, iy = y instanceof IndexSet;
  if (!ix && !iy) return null;
  if (ix && iy) { if (x.n !== y.n) return null; return new IndexSet(x.n, intersectSorted(x.idx, y.idx), true); }
  const set = ix ? x : y;
  let s = ix ? y : x;
  if (s instanceof Ref) { if (!s.isCell()) return null; s = s.value(); }
  if (s instanceof Matrix || Array.isArray(s)) return null;
  const n = toNum(s);
  if (isErr(n)) return null;
  if (n === 1) return new IndexSet(set.n, set.idx, true);
  if (n === 0) return new IndexSet(set.n, [], true);
  return null;
}

function toOperand(v, ctx) {
  if (v instanceof Ref) {
    if (v.isCell()) return v.value();
    return ctx.array ? v.toMatrix() : intersect(v, ctx);
  }
  if (Array.isArray(v)) return E.VALUE;
  return v;
}

function unary(v, ctx, f) {
  v = toOperand(v, ctx);
  if (v instanceof Matrix) return v.map((x) => f(x));
  return f(v);
}

function binary(x, y, ctx, op) {
  x = toOperand(x, ctx); y = toOperand(y, ctx);
  const mx = x instanceof Matrix, my = y instanceof Matrix;
  if (!mx && !my) return op(x, y);
  const R = Math.max(mx ? x.rows : 1, my ? y.rows : 1), C = Math.max(mx ? x.cols : 1, my ? y.cols : 1);
  const out = new Array(R * C);
  for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) {
    const a = mx ? x.get(i, j) : x, b = my ? y.get(i, j) : y;
    out[i * C + j] = op(a, b);
  }
  return new Matrix(R, C, out);
}

function arith(f) {
  return (a, b) => {
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    const x = toNum(a); if (isErr(x)) return x;
    const y = toNum(b); if (isErr(y)) return y;
    return f(x, y);
  };
}
function cmp(f) {
  return (a, b) => { const c = compare(a, b); return isErr(c) ? c : f(c); };
}
const OPS = {
  '+': arith((x, y) => x + y),
  '-': arith((x, y) => x - y),
  '*': arith((x, y) => x * y),
  '/': arith((x, y) => (y === 0 ? E.DIV0 : x / y)),
  '^': arith((x, y) => { const r = Math.pow(x, y); return isFinite(r) && !Number.isNaN(r) ? r : E.NUM; }),
  '&': (a, b) => { if (isErr(a)) return a; if (isErr(b)) return b; return toText(a) + toText(b); },
  '=': cmp((c) => c === 0),
  '<>': cmp((c) => c !== 0),
  '<': cmp((c) => c < 0),
  '>': cmp((c) => c > 0),
  '<=': cmp((c) => c <= 0),
  '>=': cmp((c) => c >= 0),
};

// mark AST nodes whose value does not depend on the evaluating cell's position
function markInvariant(ast) {
  const rec = (n) => {
    let inv = true;
    switch (n.t) {
      case 'ref': inv = !!(n.ar1 && n.ac1 && n.ar2 && n.ac2); break;
      case 'struct': inv = !/#this row|@/i.test(n.spec); break;
      case 'name': inv = false; break; // could be a LET variable or relative name
      case 'num': case 'str': case 'bool': case 'err': case 'arr': case 'missing': inv = true; break;
      case 'op': { const a = rec(n.a), b = rec(n.b); inv = a && b; break; }
      case 'neg': case 'pct': case 'paren': case 'implicit': case 'spill': inv = rec(n.a); break;
      case 'rangeop': { const a = rec(n.a), b = rec(n.b); inv = a && b; break; }
      case 'fn': {
        let all = true;
        for (const a of n.args) if (!rec(a)) all = false;
        inv = all && !['ROW', 'COLUMN', 'OFFSET', 'INDIRECT', 'LET', 'TODAY', 'NOW', 'RAND', 'RANDBETWEEN', 'CELL'].includes(n.name);
        break;
      }
      case 'union': { let all = true; for (const a of n.items) if (!rec(a)) all = false; inv = all; break; }
      default: inv = false;
    }
    // only cache non-trivial nodes
    if (inv && (n.t === 'op' || n.t === 'fn')) n.inv = true;
    return inv;
  };
  rec(ast);
}

export { KEY, excelRound };
