// Renders a block of worksheet rows as a styled HTML table: merged cells become spans,
// header cells become table headings, unlocked cells become inputs (or dropdowns for
// data-validation lists) and formulas show their live results.
import { KEY } from '../engine/engine.js';
import { isErr, Rich } from '../engine/values.js';
import { formatValue } from '../format.js';

export const el = (tag, cls, text) => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined && text !== null) e.textContent = text;
  return e;
};

export function displayText(v, fmt) {
  if (v === null || v === undefined || v === '') return '';
  if (isErr(v)) return v.code;
  if (v instanceof Rich) return v.toString();
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return formatValue(v, fmt || 'General').trim();
}

function luminance(hex) {
  const n = parseInt(hex, 16);
  return (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
}
const isHeaderXf = (xf) => !!xf.fill && luminance(xf.fill) < 0.55;
// helper values the workbook hides with white text on a white cell
const hiddenText = (xf) => !!(xf.font && xf.font.color && luminance(xf.font.color) > 0.97 && (!xf.fill || luminance(xf.fill) > 0.9));

// the value a reader sees in a cell (null when blank or hidden)
export function shownValue(app, si, r, c) {
  const v = app.eng.getValue(si, r, c);
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'string' && !v.trim()) return null;
  if (hiddenText(app.modelFor(si).xf(r, c))) return null;
  return v;
}

// Rows worth showing: visible and holding a value or an input.
export function contentRows(app, si, r1, r2, c1, c2, { editable } = {}) {
  const m = app.modelFor(si), eng = app.eng;
  const out = [];
  for (let r = r1; r <= r2; r++) {
    if (m.rowHidden(r)) continue;
    for (let c = c1; c <= c2; c++) {
      if (m.colHidden(c)) continue;
      if (shownValue(app, si, r, c) !== null || (editable && editable(r, c))) { out.push(r); break; }
    }
  }
  return out;
}

// opts:
//   cols: [c1, c2]                 column range to show
//   editable(r, c, xf, fc) -> bool which cells get inputs (default: none)
//   onEdit(r, c, text) -> Promise<bool>
//   onPick(r, c, value) -> Promise  (dropdown choice; defaults to onEdit)
//   extra(r, c, td)                 hook to add controls into a cell
//   keep(r, c) -> bool              cell must stay separate (it gets extra controls)
//   rowClass(r) -> string
export function sheetTable(app, si, rows, opts) {
  const m = app.modelFor(si), eng = app.eng, sh = eng.sheets[si];
  const [c1, c2] = opts.cols;
  const cols = [];
  for (let c = c1; c <= c2; c++) if (!m.colHidden(c)) cols.push(c);
  const table = el('table', 'sv-table');
  const cg = el('colgroup');
  const total = cols.reduce((a, c) => a + m.colWidthPx100(c), 0) || 1;
  for (const c of cols) { const col = el('col'); col.style.width = (100 * m.colWidthPx100(c) / total).toFixed(2) + '%'; cg.appendChild(col); }
  table.appendChild(cg);
  const tbody = el('tbody');
  table.appendChild(tbody);
  const rowSet = new Set(rows);
  const editable = opts.editable || (() => false);

  for (const r of rows) {
    const tr = el('tr');
    if (opts.rowClass) { const k = opts.rowClass(r); if (k) tr.className = k; }
    // collect the cells of this row, honouring merges
    const cells = [];
    for (let i = 0; i < cols.length; i++) {
      const c = cols[i];
      const mg = m.mergeAt(r, c);
      if (mg && (mg[0] !== r || mg[1] !== c)) {
        // covered by a merge: either part of a span already emitted in this row, or a
        // vertical merge from a row above
        if (mg[0] === r) continue;
        if (rowSet.has(mg[0]) && cells.length && cells[cells.length - 1].covered) { cells[cells.length - 1].span++; continue; }
        cells.push({ c, span: 1, covered: true });
        continue;
      }
      let span = 1;
      if (mg) { for (let j = i + 1; j < cols.length && cols[j] <= mg[3]; j++) span++; }
      cells.push({ c, span, mg });
    }
    // title row: one bold text cell and nothing else
    const filled = cells.filter((x) => !x.covered && (shownValue(app, si, r, x.c) !== null || editable(r, x.c, m.xf(r, x.c), sh.formulas.get(KEY(r, x.c)))));
    if (filled.length === 1 && !opts.noTitles) {
      const x = filled[0], xf = m.xf(r, x.c);
      const v = eng.getValue(si, r, x.c);
      if (xf.font && xf.font.b && typeof v === 'string' && !editable(r, x.c, xf, sh.formulas.get(KEY(r, x.c))) && (x.span >= 4 || !isHeaderXf(xf))) {
        tr.className = (tr.className + ' sv-title-row').trim();
        const td = el('td', 'sv-title');
        td.colSpan = cols.length;
        td.appendChild(el('h3', null, String(v).trim()));
        if (opts.extra) opts.extra(r, x.c, td);
        tr.appendChild(td);
        tbody.appendChild(tr);
        continue;
      }
    }
    // like Excel, text spills over the empty cells to its right
    for (let i = 0; i < cells.length; i++) {
      const x = cells[i];
      if (x.covered || x.skip || typeof shownValue(app, si, r, x.c) !== 'string') continue;
      const xf = m.xf(r, x.c);
      if (isHeaderXf(xf) || editable(r, x.c, xf, sh.formulas.get(KEY(r, x.c)))) continue;
      for (let j = i + 1; j < cells.length; j++) {
        const y = cells[j];
        const yf = m.xf(r, y.c);
        if (y.covered || shownValue(app, si, r, y.c) !== null || editable(r, y.c, yf, sh.formulas.get(KEY(r, y.c))) || isHeaderXf(yf) || (opts.keep && opts.keep(r, y.c))) break;
        x.span += y.span; y.skip = true;
      }
    }
    for (const x of cells) {
      if (x.skip) continue;
      const c = x.c;
      const xf = m.xf(r, c);
      const fc = sh.formulas.get(KEY(r, c));
      const v = shownValue(app, si, r, c);
      let td;
      if (x.covered) {
        td = el('td', 'sv-empty');
      } else if (editable(r, c, xf, fc)) {
        td = el('td', 'sv-in');
        td.appendChild(makeInput(app, si, r, c, xf, eng.getValue(si, r, c), opts));
      } else if (isHeaderXf(xf) && hasValue(v)) {
        td = el('th', 'sv-head', displayText(v, xf.fmt));
      } else {
        const txt = displayText(v, xf.fmt);
        td = el('td', txt === '' ? 'sv-empty' : fc ? 'sv-calc' : 'sv-text');
        if (txt !== '') {
          const span = el('span', null, txt);
          td.appendChild(span);
          if (xf.font && xf.font.b) td.classList.add('sv-bold');
          if (typeof v === 'number') td.classList.add('sv-num');
          if (isErr(v)) td.classList.add('sv-err');
        }
      }
      if (x.span > 1) td.colSpan = x.span;
      const note = m.comments.get(KEY(r, c));
      if (note && !x.covered) {
        const i = el('span', 'sv-note', 'i');
        i.title = note;
        td.appendChild(i);
      }
      if (opts.extra && !x.covered) opts.extra(r, c, td);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  return table;
}

function hasValue(v) { return v !== null && v !== undefined && v !== ''; }

function makeInput(app, si, r, c, xf, v, opts) {
  const m = app.modelFor(si);
  const dv = m.dvAt(r, c);
  const commit = opts.onEdit;
  const pick = opts.onPick || ((rr, cc, val) => commit(rr, cc, val === null ? '' : String(val)));
  if (dv && dv.type === 'list') {
    const sel = el('select', 'sv-input sv-select');
    const items = app.listValues(si, r, c);
    const cur = v === null || v === undefined ? '' : v instanceof Rich ? v.toString() : String(v);
    sel.appendChild(new Option('— Select —', ''));
    let found = cur === '';
    for (const it of items) {
      const t = it instanceof Rich ? it.toString() : typeof it === 'number' ? formatValue(it, 'General') : String(it);
      const o = new Option(t, t);
      if (!found && t.toLowerCase() === cur.toLowerCase()) { o.selected = true; found = true; }
      sel.appendChild(o);
    }
    if (!found) { const o = new Option(cur, cur); o.selected = true; sel.appendChild(o); }
    sel.addEventListener('change', async () => {
      const i = sel.selectedIndex - 1;
      const val = i >= 0 && i < items.length ? (items[i] instanceof Rich ? items[i].toString() : items[i]) : sel.value || null;
      await pick(r, c, val);
    });
    if (dv.prompt) sel.title = dv.prompt;
    sel.dataset.addr = r + ',' + c;
    return sel;
  }
  const inp = el('input', 'sv-input');
  inp.type = 'text';
  inp.value = app.editText(v, xf);
  inp.dataset.orig = inp.value;
  inp.dataset.addr = r + ',' + c;
  if (dv && dv.prompt) inp.title = dv.prompt;
  if (dv && (dv.type === 'decimal' || dv.type === 'whole')) inp.inputMode = 'decimal';
  const save = async () => {
    if (inp.value === inp.dataset.orig) return;
    const ok = await commit(r, c, inp.value.trim());
    if (ok === false) inp.value = inp.dataset.orig;
    else inp.dataset.orig = inp.value;
  };
  inp.addEventListener('change', save);
  inp.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); inp.blur(); }
    if (e.key === 'Escape') { inp.value = inp.dataset.orig; inp.blur(); }
  });
  return inp;
}
