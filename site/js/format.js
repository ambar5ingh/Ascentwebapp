// Cell styles -> CSS, and value formatting with Excel number formats (via SSF).
import { isErr, Rich } from './engine/values.js';

const SSF = () => globalThis.SSF;

export function hex(c) {
  if (!c) return null;
  if (typeof c === 'string') return '#' + c;
  if (c.a !== undefined && c.a < 1) {
    const n = parseInt(c.c, 16);
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${c.a})`;
  }
  return '#' + c.c;
}

const FONT_MAP = {
  calibri: 'Calibri, Carlito, "Segoe UI", Arial, sans-serif',
  'calibri light': '"Calibri Light", Calibri, Carlito, "Segoe UI", Arial, sans-serif',
  arial: 'Arial, Helvetica, sans-serif',
  'arial narrow': '"Arial Narrow", Arial, sans-serif',
  'times new roman': '"Times New Roman", Times, serif',
  'segoe ui': '"Segoe UI", Calibri, Arial, sans-serif',
  aptos: 'Aptos, Calibri, Carlito, "Segoe UI", Arial, sans-serif',
  'aptos narrow': '"Aptos Narrow", Aptos, Calibri, Carlito, Arial, sans-serif',
  cambria: 'Cambria, Caladea, Georgia, serif',
};
export function fontFamily(name) {
  if (!name) return FONT_MAP.calibri;
  return FONT_MAP[name.toLowerCase()] || `"${name}", Calibri, Carlito, Arial, sans-serif`;
}

const BORDER = {
  thin: [1, 'solid'], hair: [1, 'dotted'], dotted: [1, 'dotted'], dashed: [1, 'dashed'], dashDot: [1, 'dashed'],
  dashDotDot: [1, 'dashed'], medium: [2, 'solid'], mediumDashed: [2, 'dashed'], mediumDashDot: [2, 'dashed'],
  mediumDashDotDot: [2, 'dashed'], slantDashDot: [2, 'dashed'], thick: [3, 'solid'], double: [3, 'double'],
};
export function borderCss(b) {
  if (!b) return null;
  const [w, st] = BORDER[b[0]] || [1, 'solid'];
  return `${w}px ${st} #${b[1] || '000000'}`;
}

// Produce a CSS rule body for an xf (cell style). zoom scales font size.
export function styleCss(xf, defFont, zoom) {
  const f = { ...defFont, ...(xf.font || {}) };
  const out = [];
  out.push(`font-family:${fontFamily(f.name)}`);
  out.push(`font-size:${((f.sz || 11) * 4 / 3 * zoom).toFixed(2)}px`);
  if (f.b) out.push('font-weight:bold');
  if (f.i) out.push('font-style:italic');
  const deco = [];
  if (f.u) deco.push('underline');
  if (f.s) deco.push('line-through');
  if (deco.length) out.push('text-decoration:' + deco.join(' '));
  if (f.color) out.push('color:#' + f.color);
  return out.join(';');
}

export function fillCss(xf) {
  return xf.fill ? '#' + xf.fill : null;
}

// ------------------------------------------------------------ number formatting
const fmtCache = new Map();
export function formatValue(v, fmt) {
  if (v === null || v === undefined) return '';
  if (isErr(v)) return v.code;
  if (v instanceof Rich) return v.toString();
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') {
    if (fmt && fmt.includes('@') && fmt !== '@') {
      try { return SSF().format(fmt, v); } catch (e) { return v; }
    }
    return v;
  }
  if (typeof v === 'number') {
    if (!fmt || fmt === 'General') return generalNumber(v);
    const k = fmt;
    try {
      let s = SSF().format(fmt, v);
      return s;
    } catch (e) {
      return generalNumber(v);
    }
  }
  return String(v);
}

export function generalNumber(v) {
  if (!isFinite(v)) return '#NUM!';
  if (v === 0) return '0';
  const a = Math.abs(v);
  if (a >= 1e11 || a < 1e-9) {
    let s = v.toExponential(5).replace(/\.?0+e/, 'e');
    const [m, e] = s.split('e');
    const ex = parseInt(e, 10);
    return m + 'E' + (ex < 0 ? '-' : '+') + String(Math.abs(ex)).padStart(2, '0');
  }
  // Excel General shows up to ~11 characters
  const intDigits = Math.max(1, Math.floor(Math.log10(a)) + 1);
  const dec = Math.max(0, 10 - intDigits);
  let s = v.toFixed(Math.min(dec, 9));
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  if (s === '-0') s = '0';
  return s;
}

// Colour from a number format section such as [Red]
export function formatColor(v, fmt) {
  if (typeof v !== 'number' || !fmt || !fmt.includes('[')) return null;
  const secs = fmt.split(';');
  const sec = v > 0 || secs.length === 1 ? secs[0] : v < 0 ? secs[1] : secs[2] || secs[0];
  const m = /\[(Red|Blue|Green|Black|White|Yellow|Magenta|Cyan)\]/i.exec(sec || '');
  if (!m) return null;
  return { red: '#FF0000', blue: '#0000FF', green: '#00B050', black: '#000000', white: '#FFFFFF', yellow: '#FFFF00', magenta: '#FF00FF', cyan: '#00FFFF' }[m[1].toLowerCase()];
}

// Excel column width (characters, as stored) -> pixels at 100%
export function colWidthPx(w) {
  return Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7);
}

export function colName(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
export function addr(r, c) { return colName(c) + r; }
export function parseAddr(a) {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(a.trim());
  if (!m) return null;
  let c = 0;
  for (const ch of m[1].toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64;
  return [+m[2], c];
}

// Formula text of a shared-formula cell: shift relative A1 references by (dr, dc).
export function shiftFormula(text, dr, dc) {
  if (!dr && !dc) return text;
  let out = '', i = 0;
  const n = text.length;
  const shiftSeg = (seg) => seg.replace(/(^|[^A-Za-z0-9_.$\]])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(!])/g, (all, pre, ac, col, ar, row) => {
    let c = 0;
    for (const ch of col) c = c * 26 + ch.charCodeAt(0) - 64;
    if (c > 16384) return all;
    const nc = ac ? c : c + dc, nr = ar ? +row : +row + dr;
    if (nc < 1 || nr < 1) return pre + '#REF!';
    return pre + ac + colName(nc) + ar + nr;
  });
  let seg = '';
  while (i < n) {
    const ch = text[i];
    if (ch === '"' || ch === "'" || ch === '[') {
      out += shiftSeg(seg); seg = '';
      const close = ch === '[' ? ']' : ch;
      let j = i + 1, depth = 1;
      while (j < n) {
        if (ch === '[' && text[j] === '[') depth++;
        if (text[j] === close) {
          if (ch !== '[' && text[j + 1] === close) { j += 2; continue; }
          if (ch === '[' && --depth > 0) { j++; continue; }
          break;
        }
        j++;
      }
      out += text.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    seg += ch; i++;
  }
  return out + shiftSeg(seg);
}
