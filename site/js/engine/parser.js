// Excel formula tokenizer + parser -> AST.
// Shared by the browser engine and the Node build/verify scripts.

export class ParseError extends Error {}

const ERR_LITERALS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#GETTING_DATA'];

export function colToNum(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}
export function numToCol(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
const COL_RE = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_RE = /^(\$?)(\d+)$/;

// ---------------------------------------------------------------- tokenizer
// token: {type, value, ...}
// types: num, str, bool, err, ref (raw text incl. sheet prefix), name, func, op, lparen, rparen,
//        comma, semicolon, lbrace, rbrace, struct, ws
export function tokenize(src) {
  const toks = [];
  let i = 0;
  const n = src.length;
  const isIdStart = (ch) => /[A-Za-z_\\À-￿]/.test(ch);
  const isIdChar = (ch) => /[A-Za-z0-9_.\\?À-￿]/.test(ch);
  while (i < n) {
    const ch = src[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t' || ch === ' ') {
      let j = i; while (j < n && /[ \n\r\t ]/.test(src[j])) j++;
      toks.push({ type: 'ws' });
      i = j; continue;
    }
    if (ch === '"') {
      let j = i + 1, s = '';
      for (;;) {
        if (j >= n) throw new ParseError('Unterminated string');
        if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
        s += src[j++];
      }
      toks.push({ type: 'str', value: s });
      i = j + 1; continue;
    }
    if (ch === '#') {
      const e = ERR_LITERALS.find((x) => src.startsWith(x, i) || src.toUpperCase().startsWith(x, i));
      if (e) {
        // #REF! may be followed by !A1 etc. inside a ref, treat as plain error
        toks.push({ type: 'err', value: e });
        i += e.length; continue;
      }
      throw new ParseError('Bad error literal at ' + i);
    }
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] || ''))) {
      // could be a number, or a row range like 1:3, or a cell like 1E3 (number)
      const m = /^(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i));
      // row range e.g. 5:10 (only when followed by ':' and digits and previous token isn't operand)
      const rr = /^\$?\d+:\$?\d+/.exec(src.slice(i));
      if (rr && !m[2] && !m[3]) {
        toks.push({ type: 'ref', value: rr[0], sheet: null });
        i += rr[0].length; continue;
      }
      toks.push({ type: 'num', value: parseFloat(m[0]) });
      i += m[0].length; continue;
    }
    if (ch === "'" || ch === '[' || isIdStart(ch) || ch === '$') {
      // sheet-qualified reference, external ref, name, function, bool, cell ref, structured ref
      let j = i;
      let sheet = null, ext = null;
      if (ch === '[') {
        // external workbook index like [1]Sheet1!A1   or structured ref inside table (handled elsewhere)
        const m = /^\[(\d+)\]/.exec(src.slice(i));
        if (m) { ext = m[1]; j = i + m[0].length; }
      }
      if (src[j] === "'") {
        let k = j + 1, s = '';
        for (;;) {
          if (k >= n) throw new ParseError('Unterminated sheet name');
          if (src[k] === "'") { if (src[k + 1] === "'") { s += "'"; k += 2; continue; } break; }
          s += src[k++];
        }
        if (src[k + 1] !== '!') throw new ParseError('Expected ! after quoted sheet name');
        sheet = s; j = k + 2;
        if (ext === null) { const m2 = /^\[(\d+)\](.*)$/.exec(sheet); if (m2) { ext = m2[1]; sheet = m2[2]; } }
      } else {
        // try unquoted sheet name followed by !
        const m = /^([A-Za-z0-9_.À-￿]+)!/.exec(src.slice(j));
        if (m) { sheet = m[1]; j += m[0].length; }
      }
      if (sheet !== null) {
        // read the reference part
        const m = /^(\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+|#REF!)/.exec(src.slice(j));
        if (m) {
          toks.push({ type: 'ref', value: m[0], sheet, ext });
          i = j + m[0].length; continue;
        }
        // sheet-scoped defined name
        const m3 = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(src.slice(j));
        if (m3) { toks.push({ type: 'name', value: m3[0], sheet, ext }); i = j + m3[0].length; continue; }
        throw new ParseError('Bad sheet reference at ' + i);
      }
      // identifier
      let k = i;
      while (k < n && (isIdChar(src[k]) || src[k] === '$')) k++;
      let word = src.slice(i, k);
      if (src[k] === '(') {
        toks.push({ type: 'func', value: word });
        i = k + 1; continue;
      }
      if (src[k] === '[') {
        // structured reference Table[...]
        let depth = 0, q = k;
        for (; q < n; q++) {
          if (src[q] === "'") { q++; continue; }
          if (src[q] === '[') depth++;
          else if (src[q] === ']') { depth--; if (depth === 0) break; }
        }
        toks.push({ type: 'struct', table: word, spec: src.slice(k, q + 1) });
        i = q + 1; continue;
      }
      const cm = CELL_RE.exec(word);
      if (cm && colToNum(cm[2].toUpperCase()) <= 16384) {
        const m2 = /^:(\$?[A-Za-z]{1,3}\$?\d+)(?![A-Za-z0-9_(\[])/.exec(src.slice(k));
        if (m2) { toks.push({ type: 'ref', value: word + m2[0], sheet: null }); i = k + m2[0].length; continue; }
        toks.push({ type: 'ref', value: word, sheet: null });
        i = k; continue;
      }
      if (COL_RE.test(word)) {
        const m2 = /^:(\$?[A-Za-z]{1,3})(?![A-Za-z0-9_(\[])/.exec(src.slice(k));
        if (m2) { toks.push({ type: 'ref', value: word + m2[0], sheet: null }); i = k + m2[0].length; continue; }
      }
      const up = word.toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') { toks.push({ type: 'bool', value: up === 'TRUE' }); i = k; continue; }
      toks.push({ type: 'name', value: word });
      i = k; continue;
    }
    if (ch === '(') { toks.push({ type: 'lparen' }); i++; continue; }
    if (ch === ')') { toks.push({ type: 'rparen' }); i++; continue; }
    if (ch === '{') { toks.push({ type: 'lbrace' }); i++; continue; }
    if (ch === '}') { toks.push({ type: 'rbrace' }); i++; continue; }
    if (ch === ',') { toks.push({ type: 'comma' }); i++; continue; }
    if (ch === ';') { toks.push({ type: 'semicolon' }); i++; continue; }
    if (ch === '<' || ch === '>') {
      if (src[i + 1] === '=' ) { toks.push({ type: 'op', value: ch + '=' }); i += 2; continue; }
      if (ch === '<' && src[i + 1] === '>') { toks.push({ type: 'op', value: '<>' }); i += 2; continue; }
      toks.push({ type: 'op', value: ch }); i++; continue;
    }
    if ('+-*/^&=%:@#'.includes(ch)) { toks.push({ type: 'op', value: ch }); i++; continue; }
    throw new ParseError('Unexpected character ' + ch + ' at ' + i);
  }
  // drop whitespace tokens (intersection operator not used in this workbook)
  return toks.filter((t) => t.type !== 'ws');
}

// ---------------------------------------------------------------- reference parsing
export function parseRefText(text) {
  // returns {r1,c1,r2,c2, ar1,ac1,ar2,ac2} (1-based); rows/cols may be whole
  if (text === '#REF!') return { err: true };
  const parts = text.split(':');
  const one = (p) => {
    let m = CELL_RE.exec(p);
    if (m) return { r: +m[4], c: colToNum(m[2].toUpperCase()), ar: !!m[3], ac: !!m[1] };
    m = COL_RE.exec(p);
    if (m) return { r: null, c: colToNum(m[2].toUpperCase()), ar: true, ac: !!m[1] };
    m = ROW_RE.exec(p);
    if (m) return { r: +m[2], c: null, ar: !!m[1], ac: true };
    throw new ParseError('Bad ref ' + p);
  };
  const a = one(parts[0]);
  const b = parts.length > 1 ? one(parts[1]) : a;
  const out = { r1: a.r, c1: a.c, r2: b.r, c2: b.c, ar1: a.ar, ac1: a.ac, ar2: b.ar, ac2: b.ac };
  if (a.r === null) { out.r1 = 1; out.r2 = 1048576; out.ar1 = out.ar2 = true; out.wholeCol = true; }
  if (a.c === null) { out.c1 = 1; out.c2 = 16384; out.ac1 = out.ac2 = true; out.wholeRow = true; }
  return out;
}

// ---------------------------------------------------------------- parser
const BINARY = {
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
  '&': 2,
  '+': 3, '-': 3,
  '*': 4, '/': 4,
  '^': 5,
};

export function parse(src) {
  if (src.startsWith('=')) src = src.slice(1);
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (type) => { const t = next(); if (!t || t.type !== type) throw new ParseError('Expected ' + type + ' got ' + (t ? t.type + ':' + t.value : 'EOF') + ' in ' + src.slice(0, 120)); return t; };

  function parseExpr(minPrec) {
    let left = parseUnary();
    for (;;) {
      const t = peek();
      if (!t || t.type !== 'op' || !(t.value in BINARY)) break;
      const prec = BINARY[t.value];
      if (prec < minPrec) break;
      next();
      // ^ is left-assoc in Excel
      const right = parseExpr(prec + 1);
      left = { t: 'op', op: t.value, a: left, b: right };
    }
    return left;
  }

  function parseUnary() {
    const t = peek();
    if (t && t.type === 'op' && (t.value === '-' || t.value === '+')) {
      next();
      const a = parseUnary();
      return t.value === '-' ? { t: 'neg', a } : a;
    }
    if (t && t.type === 'op' && t.value === '@') { next(); return { t: 'implicit', a: parseUnary() }; }
    return parsePostfix();
  }

  function parsePostfix() {
    let a = parseRange();
    for (;;) {
      const t = peek();
      if (t && t.type === 'op' && t.value === '%') { next(); a = { t: 'pct', a }; continue; }
      if (t && t.type === 'op' && t.value === '#') { next(); a = { t: 'spill', a }; continue; }
      break;
    }
    return a;
  }

  function parseRange() {
    let a = parsePrimary();
    for (;;) {
      const t = peek();
      if (t && t.type === 'op' && t.value === ':') {
        next();
        const b = parsePrimary();
        a = { t: 'rangeop', a, b };
        continue;
      }
      break;
    }
    return a;
  }

  function parseArgs() {
    const args = [];
    if (peek() && peek().type === 'rparen') { next(); return args; }
    for (;;) {
      const t = peek();
      if (t && (t.type === 'comma' || t.type === 'rparen')) args.push({ t: 'missing' });
      else args.push(parseExpr(1));
      const s = next();
      if (!s) throw new ParseError('Unterminated function call');
      if (s.type === 'rparen') break;
      if (s.type !== 'comma') throw new ParseError('Expected , or ) got ' + s.type);
    }
    return args;
  }

  function parsePrimary() {
    const t = next();
    if (!t) throw new ParseError('Unexpected end of formula');
    switch (t.type) {
      case 'num': return { t: 'num', v: t.value };
      case 'str': return { t: 'str', v: t.value };
      case 'bool': return { t: 'bool', v: t.value };
      case 'err': return { t: 'err', v: t.value };
      case 'ref': {
        const r = parseRefText(t.value);
        if (r.err) return { t: 'err', v: '#REF!' };
        return { t: 'ref', sheet: t.sheet, ext: t.ext || null, ...r };
      }
      case 'struct': return { t: 'struct', table: t.table, spec: t.spec };
      case 'name': return { t: 'name', name: t.value, sheet: t.sheet || null };
      case 'func': {
        let name = t.value.toUpperCase();
        name = name.replace(/^_XLFN\./, '').replace(/^_XLWS\./, '').replace(/^_XLL\./, '');
        const args = parseArgs();
        return { t: 'fn', name, args };
      }
      case 'lparen': {
        // could be a union (a,b) - only seen in chart refs
        const e = parseExpr(1);
        if (peek() && peek().type === 'comma') {
          const items = [e];
          while (peek() && peek().type === 'comma') { next(); items.push(parseExpr(1)); }
          expect('rparen');
          return { t: 'union', items };
        }
        expect('rparen');
        return { t: 'paren', a: e };
      }
      case 'lbrace': {
        const rows = [[]];
        for (;;) {
          const x = next();
          if (!x) throw new ParseError('Unterminated array');
          if (x.type === 'rbrace') break;
          if (x.type === 'comma') continue;
          if (x.type === 'semicolon') { rows.push([]); continue; }
          let neg = false;
          let y = x;
          if (y.type === 'op' && y.value === '-') { neg = true; y = next(); }
          let v;
          if (y.type === 'num') v = neg ? -y.value : y.value;
          else if (y.type === 'str') v = y.value;
          else if (y.type === 'bool') v = y.value;
          else if (y.type === 'err') v = { e: y.value };
          else throw new ParseError('Bad array constant');
          rows[rows.length - 1].push(v);
        }
        return { t: 'arr', rows };
      }
      default:
        throw new ParseError('Unexpected token ' + t.type + ' ' + (t.value || ''));
    }
  }

  const ast = parseExpr(1);
  if (p < toks.length) throw new ParseError('Trailing tokens in formula: ' + src.slice(0, 120));
  return ast;
}

// ---------------------------------------------------------------- structured reference spec parsing
// Returns {items:[...], cols:[c1,c2]|null}; items are '#All','#Data','#Headers','#Totals','#This Row'
export function parseStructSpec(spec) {
  // spec like "[Col]" or "[[#This Row],[Col]]" or "[[Col1]:[Col2]]" or "[#Data]" or "[]"
  let s = spec.trim();
  if (s.startsWith('[') && s.endsWith(']')) s = s.slice(1, -1);
  const unesc = (x) => x.replace(/'(.)/g, '$1');
  const parts = [];
  if (!s.includes('[')) {
    if (s === '') return { items: ['#Data'], cols: null };
    if (s.startsWith('#')) return { items: [s], cols: null };
    if (s.startsWith('@')) return { items: ['#This Row'], cols: [unesc(s.slice(1)), unesc(s.slice(1))] };
    return { items: ['#Data'], cols: [unesc(s), unesc(s)] };
  }
  // tokens of [ ... ] groups separated by , or :
  let i = 0;
  const groups = [];
  let seps = [];
  while (i < s.length) {
    const ch = s[i];
    if (ch === '[') {
      let j = i + 1, buf = '';
      while (j < s.length && s[j] !== ']') { if (s[j] === "'") { buf += s[j + 1]; j += 2; continue; } buf += s[j++]; }
      groups.push(buf); i = j + 1; continue;
    }
    if (ch === ',' || ch === ':') { seps.push(ch); i++; continue; }
    if (ch === '@') {
      // [@Col] or @[Col]
      groups.push('#This Row'); seps.push(','); i++; continue;
    }
    i++;
  }
  const items = [];
  const cols = [];
  groups.forEach((g, idx) => {
    if (g.startsWith('#')) items.push(g);
    else cols.push(g);
  });
  if (!items.length) items.push('#Data');
  const range = cols.length ? [cols[0], cols[cols.length - 1]] : null;
  return { items, cols: range };
}

// ---------------------------------------------------------------- utilities over AST
export function walk(ast, fn) {
  fn(ast);
  switch (ast.t) {
    case 'op': walk(ast.a, fn); walk(ast.b, fn); break;
    case 'neg': case 'pct': case 'paren': case 'implicit': case 'spill': walk(ast.a, fn); break;
    case 'rangeop': walk(ast.a, fn); walk(ast.b, fn); break;
    case 'fn': ast.args.forEach((x) => walk(x, fn)); break;
    case 'union': ast.items.forEach((x) => walk(x, fn)); break;
    default: break;
  }
}
