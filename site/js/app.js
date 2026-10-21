// ASCENT web app: loads the workbook model, runs the formula engine in the browser and
// renders sheets, shapes, charts and macros like the original Excel file.
import { buildEngine, loadGraph, decodeValue } from './engine/load.js';
import { KEY } from './engine/engine.js';
import { isErr, err, Rich, parseNumber, numToText, toText, Ref, Matrix } from './engine/values.js';
import { SheetModel, HiddenState } from './model.js';
import { Grid } from './grid.js';
import { Drawings } from './drawings.js';
import { formatValue, colName, parseAddr, shiftFormula } from './format.js';
import { runMacro, workbookOpen, questionnaireChange } from './macros.js';
import { requireLogin, logout } from './auth.js';
import { pivotWrites } from './pivots.js';

const STORE = 'ascent-state-v1';
const $ = (s, el = document) => el.querySelector(s);

function encodeVal(v) { return isErr(v) ? { e: v.code } : v instanceof Rich ? v.toString() : v === undefined ? null : v; }

class App {
  constructor() {
    this.models = [];
    this.views = [];
    this.layoutsDirty = new Set();
    this.undoStack = [];
    this.redoStack = [];
    this.history = [];
  }

  // ------------------------------------------------------------------ loading
  async load() {
    const bar = $('#load-bar'), txt = $('#load-text');
    const step = (p, t) => { bar.style.width = (p * 100).toFixed(0) + '%'; if (t) txt.textContent = t; };
    step(0.02, 'Loading workbook…');
    const wb = await (await fetch('data/workbook.json')).json();
    this.wb = wb;
    let done = 0;
    const total = wb.sheets.length + 1;
    const tick = () => { done++; step(0.05 + 0.75 * done / total, `Loading sheets (${Math.min(done, wb.sheets.length)}/${wb.sheets.length})…`); };
    const graphP = fetch('data/graph.bin').then((r) => r.arrayBuffer()).then((b) => { tick(); return b; });
    const datas = await Promise.all(wb.sheets.map((s) => fetch('data/' + s.file).then((r) => r.json()).then((d) => { tick(); return d; })));
    this.datas = datas;
    step(0.82, 'Building the calculation engine…');
    await new Promise((r) => setTimeout(r, 0));
    this.eng = buildEngine(wb, datas);
    loadGraph(this.eng, await graphP);
    this.readOnly = new Set((wb.readOnlySheets || []).map((n) => this.eng.sheetIndex(n)));
    step(0.9, 'Restoring your inputs…');
    await new Promise((r) => setTimeout(r, 0));
    this.restore();
    step(1, 'Ready');
  }

  // ------------------------------------------------------------------ state
  restore() {
    let st = null;
    try { st = JSON.parse(localStorage.getItem(STORE) || 'null'); } catch (e) { st = null; }
    this.applyState(st);
  }
  applyState(st) {
    const wb = this.wb;
    this.state = {
      edits: {}, forced: {}, rows: {}, cols: {}, vis: {}, shapes: {}, prot: {}, views: {}, zoom: {},
      ui: { formulaBar: false, tabs: false, headings: false },
    };
    if (st && st.v === 1) Object.assign(this.state, st, { ui: { ...this.state.ui, ...(st.ui || {}) } });
    this.state.v = 1;
    this.rowStates = wb.sheets.map((_, i) => new HiddenState(this.state.rows[i]));
    this.colStates = wb.sheets.map((_, i) => new HiddenState(this.state.cols[i]));
    // re-apply inputs
    const plain = [], forced = [];
    for (const [k, v] of Object.entries(this.state.edits)) {
      const [si, r, c] = k.split(',').map(Number);
      if (!wb.sheets[si]) continue;
      (this.state.forced[k] ? forced : plain).push([si, r, c, decodeValue(v)]);
    }
    if (forced.length) this.eng.setValues(forced, true);
    if (plain.length) this.eng.setValues(plain, false);
    this.refreshPivots();
    this.view = { ...this.state.ui };
    this.views = wb.sheets.map((_, i) => ({ topRow: 1, leftCol: 1, sel: [1, 1], ...(this.state.views[i] || {}) }));
  }
  snapshot() {
    const s = this.state;
    s.rows = {}; s.cols = {};
    this.rowStates.forEach((h, i) => { if (h.reset || h.ranges.length) s.rows[i] = h.toJSON(); });
    this.colStates.forEach((h, i) => { if (h.reset || h.ranges.length) s.cols[i] = h.toJSON(); });
    s.views = {};
    this.views.forEach((v, i) => { s.views[i] = { topRow: v.topRow, leftCol: v.leftCol, sel: v.sel, gridlines: v.gridlines }; });
    s.ui = { formulaBar: this.view.formulaBar, tabs: this.view.tabs, headings: this.view.headings };
    return s;
  }
  saveSoon() {
    clearTimeout(this.saveTimer);
    this.saveTimer = setTimeout(() => this.save(), 400);
  }
  save() {
    try { localStorage.setItem(STORE, JSON.stringify(this.snapshot())); } catch (e) { console.warn('save failed', e); }
  }

  // pivot tables follow their source data (Excel would need Refresh All)
  refreshPivots() {
    const w = pivotWrites(this);
    if (w.length) this.eng.setValues(w);
  }
  rowState(si) { return this.rowStates[si]; }
  colState(si) { return this.colStates[si]; }
  viewOf(si) { return this.views[si]; }
  modelFor(si) {
    if (!this.models[si]) {
      this.models[si] = new SheetModel(this, si);
      this.models[si].layout(this.zoomOf(si));
    }
    return this.models[si];
  }
  zoomOf(si) {
    const z = this.state.zoom[si];
    if (z) return z;
    return (this.wb.sheets[si].zoom || 100) / 100;
  }
  sheetVisible(si) {
    const v = this.state.vis[si];
    if (v !== undefined) return v;
    return this.wb.sheets[si].state === 'visible' || !this.wb.sheets[si].state;
  }
  setSheetVisible(si, vis) {
    if (!vis) {
      const others = this.wb.sheets.some((_, i) => i !== si && this.sheetVisible(i));
      if (!others) return;
    }
    this.state.vis[si] = !!vis;
    this.tabsDirty = true;
  }
  isProtected(si) {
    const p = this.state.prot[si];
    if (p !== undefined) return p;
    return !!this.wb.sheets[si].protected;
  }
  setProtected(si, on) { this.state.prot[si] = !!on; }
  layoutDirty(si) { this.layoutsDirty.add(si); }
  shapeOverride(si, name, create) {
    let s = this.state.shapes[si];
    if (!s) { if (!create) return null; s = this.state.shapes[si] = {}; }
    let o = s[name];
    if (!o && create) o = s[name] = {};
    return o || null;
  }
  findShape(si, name) {
    for (const d of this.datas[si].drawing || []) if (d.obj.name === name) return d.obj;
    return null;
  }
  parseAddr(a) {
    const p = parseAddr(String(a).replace(/\$/g, ''));
    if (!p) throw new Error('Bad address ' + a);
    return p;
  }

  // ------------------------------------------------------------------ UI
  async start() {
    this.drawings = new Drawings(this);
    this.grid = new Grid(this, $('#grid'));
    this.bindUI();
    workbookOpen(this);
    this.applyUIFlags();
    this.renderTabs();
    this.showSheet(this.activeSi, true);
    $('#loading').remove();
  }

  bindUI() {
    $('#btn-home').addEventListener('click', () => this.activateSheet(this.eng.sheetIndex('ASCENT India')));
    $('#btn-back').addEventListener('click', () => this.back());
    $('#btn-undo').addEventListener('click', () => this.undo());
    $('#btn-redo').addEventListener('click', () => this.redo());
    $('#zoom-in').addEventListener('click', () => this.zoomBy(0.1));
    $('#zoom-out').addEventListener('click', () => this.zoomBy(-0.1));
    $('#zoom-val').addEventListener('click', () => this.setZoom(this.activeSi, 1));
    $('#sheet-select').addEventListener('change', (e) => {
      const si = +e.target.value;
      if (!this.sheetVisible(si)) this.setSheetVisible(si, true);
      this.activateSheet(si);
      this.renderTabs();
    });
    for (const id of ['formulaBar', 'tabs', 'headings']) {
      const cb = $(`#opt-${id}`);
      cb.addEventListener('change', () => { this.view[id] = cb.checked; this.applyUIFlags(); this.grid.relayout(); this.saveSoon(); });
    }
    $('#opt-gridlines').addEventListener('change', (e) => { this.viewOf(this.activeSi).gridlines = e.target.checked; this.grid.invalidate(); this.saveSoon(); });
    $('#menu-export').addEventListener('click', () => this.exportInputs());
    $('#menu-import').addEventListener('click', () => $('#import-file').click());
    $('#import-file').addEventListener('change', (e) => this.importInputs(e.target.files[0]));
    $('#menu-csv').addEventListener('click', () => this.downloadCsv());
    $('#menu-reset').addEventListener('click', () => this.resetAll());
    $('#menu-protect').addEventListener('click', () => this.toggleProtect());
    $('#menu-logout').addEventListener('click', () => { logout(); location.reload(); });
    document.querySelectorAll('.menu > button').forEach((b) => b.addEventListener('click', (e) => {
      e.stopPropagation();
      const m = b.parentElement;
      const open = m.classList.contains('open');
      document.querySelectorAll('.menu.open').forEach((x) => x.classList.remove('open'));
      if (!open) m.classList.add('open');
    }));
    document.addEventListener('click', () => document.querySelectorAll('.menu.open').forEach((x) => x.classList.remove('open')));
    document.querySelectorAll('.menu-panel').forEach((p) => p.addEventListener('click', (e) => { if (e.target.type !== 'checkbox' && e.target.tagName !== 'LABEL') p.parentElement.classList.remove('open'); else e.stopPropagation(); }));
    const fx = $('#fx-input');
    fx.addEventListener('keydown', async (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const [r, c] = this.grid.sel;
        if (!fx.readOnly) { const ok = await this.userEdit(this.activeSi, r, c, fx.value); if (ok) this.grid.scroll.focus(); }
      } else if (e.key === 'Escape') { this.onSelect(this.activeSi, ...this.grid.sel); this.grid.scroll.focus(); }
    });
    $('#name-box').addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      this.gotoText(e.target.value.trim());
    });
    window.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && (e.key === '=' || e.key === '+')) { e.preventDefault(); this.zoomBy(0.1); }
      if ((e.ctrlKey || e.metaKey) && e.key === '-') { e.preventDefault(); this.zoomBy(-0.1); }
      if (e.ctrlKey && e.key === 'PageDown') { e.preventDefault(); this.cycleSheet(1); }
      if (e.ctrlKey && e.key === 'PageUp') { e.preventDefault(); this.cycleSheet(-1); }
    });
    $('#grid').addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      this.zoomBy(e.deltaY < 0 ? 0.05 : -0.05);
    }, { passive: false });
    window.addEventListener('paste', (e) => {
      if (document.activeElement !== this.grid.scroll) return;
      const t = (e.clipboardData && e.clipboardData.getData('text/plain')) || '';
      e.preventDefault();
      this.pasteText(t);
    });
    window.addEventListener('beforeunload', () => this.save());
  }

  applyUIFlags() {
    document.body.classList.toggle('show-fx', !!this.view.formulaBar);
    document.body.classList.toggle('show-tabs', !!this.view.tabs);
    $('#opt-formulaBar').checked = !!this.view.formulaBar;
    $('#opt-tabs').checked = !!this.view.tabs;
    $('#opt-headings').checked = !!this.view.headings;
  }

  renderTabs() {
    this.tabsDirty = false;
    const tabs = $('#tabs');
    tabs.innerHTML = '';
    const sel = $('#sheet-select');
    sel.innerHTML = '';
    const grpV = document.createElement('optgroup'); grpV.label = 'Sheets';
    const grpH = document.createElement('optgroup'); grpH.label = 'Hidden sheets';
    this.wb.sheets.forEach((s, i) => {
      const vis = this.sheetVisible(i);
      const o = document.createElement('option');
      o.value = i; o.textContent = s.name.trim();
      if (i === this.activeSi) o.selected = true;
      (vis ? grpV : grpH).appendChild(o);
      if (!vis) return;
      const b = document.createElement('button');
      b.className = 'tab' + (i === this.activeSi ? ' active' : '');
      b.textContent = s.name.trim();
      if (s.tabColor) b.style.setProperty('--tab', '#' + s.tabColor);
      b.addEventListener('click', () => this.activateSheet(i));
      tabs.appendChild(b);
    });
    sel.appendChild(grpV);
    if (grpH.children.length) sel.appendChild(grpH);
    const act = tabs.querySelector('.active');
    if (act) act.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  // Switch sheets (used by tabs, links and macros).
  activateSheet(si, opts = {}) {
    if (si < 0) return;
    if (this.activeSi !== undefined && this.activeSi !== si && this.grid.model) {
      this.grid.captureScroll();
      this.history.push(this.activeSi);
      if (this.history.length > 50) this.history.shift();
    }
    this.activeSi = si;
    if (!opts.fromMacro) this.showSheet(si, true);
  }
  back() {
    while (this.history.length) {
      const si = this.history.pop();
      if (this.sheetVisible(si) && si !== this.activeSi) { this.activeSi = si; this.showSheet(si, true); return; }
    }
  }
  cycleSheet(d) {
    const vis = this.wb.sheets.map((_, i) => i).filter((i) => this.sheetVisible(i));
    const k = vis.indexOf(this.activeSi);
    this.activateSheet(vis[(k + d + vis.length) % vis.length]);
  }
  showSheet(si, restore) {
    const m = this.modelFor(si);
    this.layoutsDirty.delete(si);
    this.grid.show(m, this.views[si]);
    const v = this.views[si];
    if (v.ensureSel) { this.grid.ensureVisible(...v.sel); v.ensureSel = false; }
    this.renderTabs();
    this.updateZoomLabel();
    $('#sheet-title').textContent = this.wb.sheets[si].name.trim();
    $('#opt-gridlines').checked = v.gridlines !== false && m.showGrid;
    this.onSelect(si, ...this.grid.sel);
    document.title = this.wb.sheets[si].name.trim() + ' · ASCENT';
    this.saveSoon();
    void restore;
  }

  // After macros change rows/visibility/shapes: refresh what is shown.
  refreshAfterChange(prevActive) {
    const si = this.activeSi;
    if (si !== prevActive || !this.grid.model || this.grid.model.si !== si) {
      if (!this.sheetVisible(si)) this.state.vis[si] = true;
      this.showSheet(si, true);
    } else {
      const v = this.views[si];
      this.layoutsDirty.delete(si);
      this.grid.relayout(true);
      if (v.ensureSel) { this.grid.sel = v.sel; this.grid.ensureVisible(...v.sel); v.ensureSel = false; }
      this.grid.sel = v.sel;
      this.grid.positionOverlay();
      this.grid.invalidate();
      this.onSelect(si, ...v.sel);
    }
    if (this.tabsDirty) this.renderTabs();
    this.saveSoon();
  }

  async runMacro(name, ctx) {
    if (this.macroBusy) return;
    this.macroBusy = true;
    document.body.classList.add('busy');
    const prev = this.activeSi;
    this.grid.captureScroll();
    try {
      await runMacro(this, name, ctx);
    } finally {
      this.macroBusy = false;
      document.body.classList.remove('busy');
      this.refreshAfterChange(prev);
    }
  }

  setZoom(si, z, opts = {}) {
    z = Math.max(0.1, Math.min(4, Math.round(z * 100) / 100));
    this.state.zoom[si] = z;
    if (opts.fromMacro) { this.layoutDirty(si); return; }
    if (si === this.activeSi) { this.grid.relayout(); this.updateZoomLabel(); }
    this.saveSoon();
  }
  zoomBy(d) { this.setZoom(this.activeSi, this.zoomOf(this.activeSi) + d); }
  updateZoomLabel() { $('#zoom-val').textContent = Math.round(this.zoomOf(this.activeSi) * 100) + '%'; }

  // ------------------------------------------------------------------ selection / formula bar
  onSelect(si, r, c) {
    const nb = $('#name-box');
    nb.value = colName(c) + r;
    const fx = $('#fx-input');
    const sh = this.eng.sheets[si];
    const fc = sh.formulas.get(KEY(r, c));
    const m = this.modelFor(si);
    if (fc) {
      let t = shiftFormula(fc.text, fc.r - fc.r0, fc.c - fc.c0).replace(/_xlfn\.|_xlws\.|_xlpm\./g, '');
      fx.value = (fc.kind === 'a' ? '{=' : '=') + t + (fc.kind === 'a' ? '}' : '');
    } else fx.value = this.editText(sh.get(r, c), m.xf(r, c));
    fx.readOnly = !!this.cannotEdit(si, r, c);
    const v = sh.get(r, c);
    const status = $('#status-cell');
    status.textContent = v === null ? '' : (typeof v === 'number' ? 'Value: ' + numToText(v) : '');
  }

  gotoText(t, link = false) {
    if (!t) return;
    let si = this.activeSi, a = t;
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(t);
    if (m) { si = this.eng.sheetIndex((m[1] || m[2]).replace(/''/g, "'")); a = m[3]; }
    if (si < 0) return;
    let p = parseAddr(a.replace(/\$/g, '').split(':')[0]);
    if (!p) {
      const nm = this.eng.names.get(a.toLowerCase()) || this.eng.names.get(si + '!' + a.toLowerCase());
      if (nm) return this.gotoText(nm.f, link);
      return;
    }
    if (!this.sheetVisible(si)) { this.alert('Microsoft Excel', 'Reference isn’t valid.'); return; }
    if (si !== this.activeSi) this.activateSheet(si);
    if (link && !this.grid.isCellVisible(p[0], p[1])) {
      // a hyperlink puts its target cell at the top-left of the window
      const g = this.grid, mm = g.model;
      const colVis = mm.colX[p[1]] >= g.scroll.scrollLeft && mm.colX[p[1]] + Math.min(mm.colW[p[1]], 50) <= g.scroll.scrollLeft + g.W - g.gx;
      g.captureScroll();
      g.applyScroll(p[0], colVis ? g.view.leftCol : p[1]);
      this.grid.select(p[0], p[1], { scroll: false });
    } else this.grid.select(p[0], p[1]);
    this.grid.scroll.focus();
  }

  followLink(url, fromSi) {
    if (!url) return;
    if (url.startsWith('#')) { this.gotoText(url.slice(1), true); return; }
    if (/^(https?:|mailto:)/i.test(url)) window.open(url, '_blank', 'noopener');
    void fromSi;
  }

  // ------------------------------------------------------------------ editing
  cannotEdit(si, r, c) {
    if (this.readOnly.has(si)) return { title: 'Read-only sheet', msg: 'This reference sheet is read-only in the web app.' };
    const sh = this.eng.sheets[si];
    const fc = sh.formulas.get(KEY(r, c));
    const m = this.modelFor(si);
    if (this.isProtected(si)) {
      const mg = m.mergeAt(r, c);
      const xf = mg ? m.xf(mg[0], mg[1]) : m.xf(r, c);
      if (!xf.unlocked) return { title: 'Microsoft Excel', msg: 'The cell or chart you’re trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.' };
    }
    if (fc && fc.kind) return { title: 'Microsoft Excel', msg: 'You can’t change part of an array.' };
    return null;
  }
  editText(v, xf) {
    if (v === null || v === undefined) return '';
    if (isErr(v)) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (v instanceof Rich) return v.toString();
    if (typeof v === 'number') {
      const fmt = (xf && xf.fmt) || '';
      if (/%/.test(fmt.split(';')[0].replace(/"[^"]*"/g, ''))) return numToText(Math.round(v * 100 * 1e10) / 1e10) + '%';
      return numToText(v);
    }
    return String(v);
  }
  parseInput(text, xf) {
    if (text === '') return { v: null };
    if (text.startsWith('=') && text.length > 1) return { error: 'Typing new formulas is not supported in the web app. Enter a value instead.' };
    const up = text.trim().toUpperCase();
    if (up === 'TRUE') return { v: true };
    if (up === 'FALSE') return { v: false };
    if (/^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/.test(up)) return { v: err(up) };
    const n = parseNumber(text);
    if (n !== null) {
      const fmt = (xf && xf.fmt) || '';
      const pctFmt = /%/.test(fmt.split(';')[0].replace(/"[^"]*"/g, ''));
      if (pctFmt && !text.trim().endsWith('%') && Math.abs(n) >= 1) return { v: n / 100 };
      return { v: n };
    }
    return { v: text };
  }

  // Data validation of a value about to be entered; returns true to accept.
  async validate(si, r, c, v) {
    const m = this.modelFor(si);
    const dv = m.dvAt(r, c);
    if (!dv || !dv.type || dv.type === 'none') return true;
    if (v === null) return true;
    let ok = true;
    const evalNum = (f) => {
      if (f === undefined) return null;
      let x = this.eng.evaluateRel(f, si, dv.r0, dv.c0, r, c);
      if (x instanceof Ref) x = x.value();
      if (x instanceof Matrix) x = x.get(0, 0);
      return typeof x === 'number' ? x : parseNumber(String(x));
    };
    const cmp = (x) => {
      const a = evalNum(dv.f1), b = evalNum(dv.f2);
      switch (dv.operator || 'between') {
        case 'between': return x >= a && x <= b;
        case 'notBetween': return x < a || x > b;
        case 'equal': return x === a;
        case 'notEqual': return x !== a;
        case 'greaterThan': return x > a;
        case 'lessThan': return x < a;
        case 'greaterThanOrEqual': return x >= a;
        case 'lessThanOrEqual': return x <= a;
        default: return true;
      }
    };
    if (dv.type === 'list') {
      const items = this.grid.model && this.grid.model.si === si ? this.grid.listValues(r, c) : [];
      const t = toText(v).toLowerCase();
      ok = items.some((it) => (typeof it === 'number' && typeof v === 'number' ? Math.abs(it - v) < 1e-9 : toText(it instanceof Rich ? it.toString() : it).toLowerCase() === t));
    } else if (dv.type === 'whole') ok = typeof v === 'number' && Number.isInteger(v) && cmp(v);
    else if (dv.type === 'decimal' || dv.type === 'date' || dv.type === 'time') ok = typeof v === 'number' && cmp(v);
    else if (dv.type === 'textLength') ok = cmp(toText(v).length);
    if (ok) return true;
    if (dv.showErrorMessage !== '1' && dv.showErrorMessage !== 1) return true;
    const title = dv.errorTitle || 'Microsoft Excel';
    const msg = dv.error || 'This value doesn’t match the data validation restrictions defined for this cell.';
    const style = dv.errorStyle || 'stop';
    if (style === 'stop') { await this.alert(title, msg); return false; }
    if (style === 'warning') return this.confirm(msg + '\n\nContinue?', title);
    await this.alert(title, msg);
    return true;
  }

  async userEdit(si, r, c, text) {
    const why = this.cannotEdit(si, r, c);
    if (why) { await this.alert(why.title, why.msg); return false; }
    const m = this.modelFor(si);
    const p = this.parseInput(text, m.xf(r, c));
    if (p.error) { await this.alert('ASCENT', p.error); return false; }
    if (!(await this.validate(si, r, c, p.v))) return false;
    this.applyCells([[si, r, c, p.v]], { user: true });
    await this.afterUserChange(si, [[r, c]]);
    return true;
  }
  async setFromList(si, r, c, v) {
    const why = this.cannotEdit(si, r, c);
    if (why) { await this.alert(why.title, why.msg); return; }
    if (v instanceof Rich) v = v.toString();
    this.applyCells([[si, r, c, v]], { user: true });
    await this.afterUserChange(si, [[r, c]]);
  }
  async pasteText(t) {
    const si = this.activeSi;
    const [r0, c0] = this.grid.sel;
    const rows = t.replace(/\r?\n$/, '').split(/\r?\n/).map((l) => l.split('\t'));
    const list = [];
    const m = this.modelFor(si);
    for (let i = 0; i < rows.length; i++) for (let j = 0; j < rows[i].length; j++) {
      const r = r0 + i, c = c0 + j;
      if (this.cannotEdit(si, r, c)) continue;
      const p = this.parseInput(rows[i][j], m.xf(r, c));
      if (p.error) continue;
      list.push([si, r, c, p.v]);
    }
    if (!list.length) { const why = this.cannotEdit(si, r0, c0); if (why) await this.alert(why.title, why.msg); return; }
    this.applyCells(list, { user: true });
    await this.afterUserChange(si, list.map((x) => [x[1], x[2]]));
  }

  // Write values into the engine, record them for persistence and undo.
  applyCells(list, { user = false, force = false, undo = true } = {}) {
    const eng = this.eng;
    const rec = [];
    const t0 = performance.now();
    const plain = [], forced = [];
    for (const [si, r, c, v] of list) {
      const sh = eng.sheets[si];
      const fc = sh.formulas.get(KEY(r, c));
      const old = sh.get(r, c);
      const mustForce = !!fc && (force || (user && !this.isProtected(si)));
      if (fc && !mustForce) continue;
      rec.push({ si, r, c, old, v, fc: mustForce ? fc : null, wasEdit: this.state.edits[[si, r, c].join(',')] });
      (mustForce ? forced : plain).push([si, r, c, v]);
      const k = [si, r, c].join(',');
      this.state.edits[k] = encodeVal(v);
      if (mustForce) this.state.forced[k] = 1;
    }
    if (forced.length) eng.setValues(forced, true);
    if (plain.length) eng.setValues(plain, false);
    this.refreshPivots();
    if (undo && rec.length) { this.undoStack.push(rec); if (this.undoStack.length > 100) this.undoStack.shift(); this.redoStack = []; }
    $('#status-calc').textContent = `Calculated in ${Math.round(performance.now() - t0)} ms`;
    this.afterRecalc();
    this.saveSoon();
    return rec;
  }
  macroWrite(list, opts = {}) { this.applyCells(list, { force: !!opts.force, undo: !opts.force }); }

  async afterUserChange(si, cells) {
    // Worksheet_Change of Questionnaire fires when F9 is part of the change
    if (si === this.eng.sheetIndex('Questionnaire') && cells.some(([r, c]) => r === 9 && c === 6)) {
      const prev = this.activeSi;
      this.grid.captureScroll();
      try { await questionnaireChange(this, [9, 6]); } catch (e) { console.error(e); }
      this.refreshAfterChange(prev);
    }
  }

  afterRecalc() {
    this.grid.textCache.clear();
    this.grid.invalidate();
    this.drawings.refresh();
    if (this.grid.model) this.onSelect(this.activeSi, ...this.grid.sel);
  }

  undo() { this.undoRedo(this.undoStack, this.redoStack, true); }
  redo() { this.undoRedo(this.redoStack, this.undoStack, false); }
  undoRedo(from, to, isUndo) {
    const rec = from.pop();
    if (!rec) return;
    const eng = this.eng;
    const back = [];
    for (const e of rec.slice().reverse()) {
      const k = [e.si, e.r, e.c].join(',');
      if (isUndo) {
        if (e.fc) {
          // restore the formula that was overwritten
          e.fc.dead = false;
          eng.sheets[e.si].formulas.set(KEY(e.r, e.c), e.fc);
          delete this.state.forced[k];
          delete this.state.edits[k];
          eng.propagate(new Set([e.fc.id]), []);
        } else {
          eng.setValues([[e.si, e.r, e.c, e.old]]);
          if (e.wasEdit === undefined) delete this.state.edits[k]; else this.state.edits[k] = e.wasEdit;
        }
      } else {
        eng.setValues([[e.si, e.r, e.c, e.v]], !!e.fc);
        this.state.edits[k] = encodeVal(e.v);
        if (e.fc) this.state.forced[k] = 1;
      }
      back.push(e);
    }
    to.push(rec);
    this.refreshPivots();
    this.afterRecalc();
    this.saveSoon();
    const last = rec[0];
    if (last.si === this.activeSi) this.grid.select(last.r, last.c);
    if (rec.some((e) => e.si === this.eng.sheetIndex('Questionnaire') && e.r === 9 && e.c === 6)) this.afterUserChange(last.si, [[9, 6]]);
    void back;
  }

  controlChanged(si, ct, all) {
    const radios = all.filter((c) => c.objectType === 'Radio');
    const link = (all.find((c) => c.fmlaLink) || ct).fmlaLink;
    if (!link) return;
    const [r, c] = this.parseAddr(link);
    const v = ct.objectType === 'Radio' ? radios.indexOf(ct) + 1 : !this.eng.getValue(si, r, c);
    this.applyCells([[si, r, c, v]], { force: true });
  }

  // ------------------------------------------------------------------ menu actions
  exportInputs() {
    const blob = new Blob([JSON.stringify(this.snapshot(), null, 1)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'ascent-inputs.json';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }
  async importInputs(file) {
    if (!file) return;
    try {
      const st = JSON.parse(await file.text());
      if (st.v !== 1 || !st.edits) throw new Error('Not an ASCENT inputs file');
      localStorage.setItem(STORE, JSON.stringify(st));
      location.reload();
    } catch (e) { this.alert('Import failed', String(e.message || e)); }
  }
  downloadCsv() {
    const si = this.activeSi, m = this.modelFor(si), sh = this.eng.sheets[si];
    const lines = [];
    for (let r = 1; r <= sh.maxRow; r++) {
      const row = [];
      for (let c = 1; c <= sh.maxCol; c++) {
        const v = sh.get(r, c);
        let t = v === null ? '' : formatValue(v, m.xf(r, c).fmt || 'General').trim();
        if (/[",\n]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
        row.push(t);
      }
      lines.push(row.join(',').replace(/,+$/, ''));
    }
    const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = this.wb.sheets[si].name.trim() + '.csv';
    a.click();
  }
  async resetAll() {
    if (!(await this.confirm('Clear all your inputs and restore the workbook to its original state?', 'Reset workbook'))) return;
    try { localStorage.removeItem(STORE); } catch (e) { /* ignore */ }
    this.state = null;
    window.onbeforeunload = null;
    this.save = () => {};
    location.reload();
  }
  async toggleProtect() {
    const si = this.activeSi;
    if (this.isProtected(si)) {
      const pw = await this.prompt('Unprotect Sheet', 'Password:');
      if (pw === null) return;
      if (pw !== 'Wrii@3' && pw !== 'Wrii1234') { await this.alert('Microsoft Excel', 'The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.'); return; }
      this.setProtected(si, false);
      this.toast('Sheet unprotected');
    } else { this.setProtected(si, true); this.toast('Sheet protected'); }
    this.onSelect(si, ...this.grid.sel);
    this.grid.positionOverlay();
    this.saveSoon();
  }

  // ------------------------------------------------------------------ dialogs
  dialog(title, msg, buttons, withInput) {
    return new Promise((resolve) => {
      const back = document.createElement('div');
      back.className = 'modal-back';
      const box = document.createElement('div');
      box.className = 'modal';
      box.setAttribute('role', 'dialog');
      const h = document.createElement('h2'); h.textContent = title;
      const p = document.createElement('p'); p.textContent = msg;
      box.append(h, p);
      let input = null;
      if (withInput) { input = document.createElement('input'); input.type = 'password'; box.appendChild(input); }
      const row = document.createElement('div'); row.className = 'modal-buttons';
      buttons.forEach(([label, val], i) => {
        const b = document.createElement('button');
        b.textContent = label;
        if (i === 0) b.className = 'primary';
        b.addEventListener('click', () => { back.remove(); resolve(withInput ? (val ? input.value : null) : val); });
        row.appendChild(b);
      });
      box.appendChild(row);
      back.appendChild(box);
      document.body.appendChild(back);
      back.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') { back.remove(); resolve(withInput ? null : buttons[buttons.length - 1][1]); }
        if (e.key === 'Enter') { e.preventDefault(); row.firstChild.click(); }
      });
      (input || row.firstChild).focus();
    }).finally(() => { if (this.grid) this.grid.scroll.focus({ preventScroll: true }); });
  }
  alert(title, msg) { return this.dialog(title, msg, [['OK', true]]); }
  confirm(msg, title = 'Microsoft Excel') { return this.dialog(title, msg, [['Yes', true], ['No', false]]); }
  prompt(title, msg) { return this.dialog(title, msg, [['OK', true], ['Cancel', false]], true); }
  toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(this.toastTimer);
    this.toastTimer = setTimeout(() => t.classList.remove('show'), 1800);
  }
}

// ---------------------------------------------------------------------- boot
const app = new App();
window.ascent = app;
const loaded = app.load().catch((e) => {
  console.error(e);
  $('#load-text').textContent = 'Could not load the workbook data: ' + e.message + '. Serve the site over http(s), not from file://.';
  throw e;
});
requireLogin($('#login')).then(async () => {
  $('#loading').style.display = 'flex';
  await loaded;
  await app.start();
});
