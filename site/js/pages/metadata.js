// Metadata: the workbook's secondary-data table (Census, Vahan, Livestock Census, ...) for
// the location chosen in the Questionnaire. Values can be corrected and new rows added; the
// questionnaire lookups ('Metadata helper') read the same table, and when the questionnaire
// uses secondary data the answers are refilled straight away.
import { el, displayText } from './sheetview.js';
import { copyMetadata, usesSecondaryData } from '../macros.js';
import { numToText } from '../engine/values.js';

// Metadata table columns (C..Q)
const COL = { code: 3, tier: 4, state: 5, district: 6, city: 7, village: 8, sector: 9, subsector: 10, question: 11, fuel: 12, value: 13, unit: 14, source: 15, origin: 16, year: 17 };
const PAGE = 50;

export class MetadataPage {
  constructor(app) {
    this.app = app;
    this.si = app.eng.sheetIndex('Metadata');
    this.qsi = app.eng.sheetIndex('Questionnaire');
    const md = app.wb.metadata || { first: 7, last: app.eng.sheets[this.si].maxRow, spare: 0 };
    this.first = md.first; this.last = md.last; this.end = md.last + md.spare;
    this.search = '';
    this.sector = '';
    this.page = 0;
    this.codes = this.readCodes();
  }

  v(r, c) { const x = this.app.eng.getValue(this.si, r, c); return x === undefined ? null : x; }
  q(r) { const x = this.app.eng.getValue(this.qsi, r, 6); return x === undefined || x === '' ? null : x; }

  // Meta_Code: code -> sector / subsector / question
  readCodes() {
    const t = (this.app.wb.tables || []).find((x) => x.name === 'Meta_Code');
    const out = [];
    if (!t) return out;
    const [r1, c1, r2] = t.ref;
    for (let r = r1 + 1; r <= r2; r++) {
      const code = this.v(r, c1);
      if (code === null) continue;
      out.push({ code, sector: this.v(r, c1 + 1), subsector: this.v(r, c1 + 2), question: this.v(r, c1 + 3) });
    }
    return out;
  }

  location() {
    return { tier: this.q(9), state: this.q(10), district: this.q(11), city: this.q(12), village: this.q(13) };
  }

  // rows of the table that belong to the location (same rule as the questionnaire lookups)
  matches(r, loc) {
    const eq = (a, b) => String(a === null ? '' : a).trim().toLowerCase() === String(b === null ? '' : b).trim().toLowerCase();
    if (this.v(r, COL.state) === null) return false;
    if (!eq(this.v(r, COL.state), loc.state)) return false;
    if (loc.tier === 'State') return eq(this.v(r, COL.tier), 'State');
    if (!eq(this.v(r, COL.district), loc.district)) return false;
    if (loc.tier === 'City' || loc.tier === 'Village' || loc.tier === 'Town') { if (!eq(this.v(r, COL.city), loc.city)) return false; }
    else if (this.v(r, COL.city) !== null) return false;
    if (loc.tier === 'Village' || loc.tier === 'Town') { if (!eq(this.v(r, COL.village), loc.village)) return false; }
    return true;
  }

  rows() {
    const loc = this.location();
    const out = [];
    if (!loc.state) return out;
    for (let r = this.first; r <= this.end; r++) if (this.matches(r, loc)) out.push(r);
    return out;
  }

  mount(host) { this.host = host; this.render(); }

  render() {
    const app = this.app, host = this.host;
    const scroll = host.scrollTop;
    const focus = document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.key : null;
    host.innerHTML = '';
    const wrap = el('div', 'page-wrap');
    const hero = el('section', 'hero hero-sm');
    const h1 = el('h1', 'hero-title');
    h1.appendChild(el('span', 'gradient-text', 'Metadata'));
    hero.append(h1, el('p', 'hero-sub', 'Secondary data for every state and district (Census, Vahan, Livestock Census and more). Add or correct values here and the Questionnaire fills itself from them.'));
    wrap.appendChild(hero);

    wrap.appendChild(this.locationCard());
    const all = this.rows();
    const loc = this.location();

    // sync card
    const sync = el('section', 'card sync-card');
    const on = usesSecondaryData(app);
    const lab = el('label', 'switch');
    const cb = el('input'); cb.type = 'checkbox'; cb.checked = on;
    cb.addEventListener('change', () => this.setAuto(cb.checked));
    lab.append(cb, el('span', 'slider'));
    const txt = el('div', 'sync-text');
    txt.appendChild(el('strong', null, 'Fill the Questionnaire from Metadata'));
    txt.appendChild(el('span', null, on
      ? 'On: the Questionnaire uses secondary data, and every change on this page (or a new location) refills it automatically.'
      : 'Off: the Questionnaire keeps your own answers. Turn on to copy the values below into it (same as answering “Yes” to “use secondary data”).'));
    sync.append(txt, lab);
    if (on) {
      const now = el('button', 'btn btn-ghost btn-sm', 'Refill now');
      now.addEventListener('click', () => this.autoFill(true));
      sync.appendChild(now);
    }
    wrap.appendChild(sync);

    // data card
    const card = el('section', 'card');
    const head = el('div', 'card-head');
    head.appendChild(el('h2', 'section-title', loc.state ? `Data for ${app.location().parts.join(', ')}` : 'Choose a location'));
    head.appendChild(el('span', 'grow'));
    const add = el('button', 'btn btn-primary btn-sm', '+ Add data');
    add.disabled = !loc.state;
    add.addEventListener('click', () => this.addDialog());
    head.appendChild(add);
    card.appendChild(head);

    const sectors = [...new Set(all.map((r) => this.v(r, COL.sector)).filter((x) => x))].sort();
    const filt = el('div', 'filter-row');
    const s = el('input', 'field-input');
    s.type = 'search'; s.placeholder = 'Search questions, fuels, sources…'; s.value = this.search; s.dataset.key = 'search';
    s.addEventListener('input', () => { this.search = s.value; this.page = 0; this.renderSoon(); });
    const sel = el('select', 'field-input');
    sel.appendChild(new Option('All sectors', ''));
    for (const x of sectors) { const o = new Option(x, x); o.selected = x === this.sector; sel.appendChild(o); }
    sel.addEventListener('change', () => { this.sector = sel.value; this.page = 0; this.render(); });
    filt.append(s, sel);
    card.appendChild(filt);

    const needle = this.search.trim().toLowerCase();
    const rows = all.filter((r) => (!this.sector || this.v(r, COL.sector) === this.sector) &&
      (!needle || [COL.sector, COL.subsector, COL.question, COL.fuel, COL.source, COL.unit].some((c) => String(this.v(r, c) || '').toLowerCase().includes(needle))));
    const pages = Math.max(1, Math.ceil(rows.length / PAGE));
    this.page = Math.min(this.page, pages - 1);
    const shown = rows.slice(this.page * PAGE, (this.page + 1) * PAGE);

    const tbl = el('table', 'data-table meta-table');
    const thead = el('thead'), htr = el('tr');
    for (const h of ['Sector', 'Sub-sector', 'Question', 'Fuel', 'Value', 'Unit', 'Source', 'Year', '']) htr.appendChild(el('th', null, h));
    thead.appendChild(htr); tbl.appendChild(thead);
    const tb = el('tbody');
    for (const r of shown) tb.appendChild(this.rowEl(r));
    tbl.appendChild(tb);
    const box = el('div', 'sv-scroll');
    box.appendChild(tbl);
    if (!rows.length) box.appendChild(el('p', 'empty', loc.state ? 'No metadata rows for this location yet. Use “Add data” to enter some.' : 'Pick a state (and district) above to see its data.'));
    card.appendChild(box);
    const pager = el('div', 'pager');
    pager.appendChild(el('span', 'muted', `${rows.length} row${rows.length === 1 ? '' : 's'}`));
    pager.appendChild(el('span', 'grow'));
    const pv = el('button', 'btn btn-ghost btn-sm', '‹'); pv.disabled = this.page === 0;
    pv.addEventListener('click', () => { this.page--; this.render(); });
    const nx = el('button', 'btn btn-ghost btn-sm', '›'); nx.disabled = this.page >= pages - 1;
    nx.addEventListener('click', () => { this.page++; this.render(); });
    pager.append(pv, el('span', 'muted', `Page ${this.page + 1} of ${pages}`), nx);
    card.appendChild(pager);
    wrap.appendChild(card);
    host.appendChild(wrap);
    host.scrollTop = scroll;
    if (focus) {
      const f = host.querySelector(`[data-key="${focus}"]`);
      if (f) { f.focus({ preventScroll: true }); if (f.type === 'search') f.setSelectionRange(f.value.length, f.value.length); }
    }
  }
  renderSoon() { clearTimeout(this.rt); this.rt = setTimeout(() => this.render(), 200); }

  locationCard() {
    const app = this.app;
    const card = el('section', 'card');
    card.appendChild(el('h2', 'section-title', 'Location'));
    card.appendChild(el('p', 'section-desc', 'This is the Questionnaire’s location, so changing it here changes it there too.'));
    const grid = el('div', 'field-grid');
    const loc = this.location();
    const fields = [[9, 'Government Tier'], [10, 'State'], [11, 'District'], [12, 'City'], [13, 'Village']];
    const show = { 9: true, 10: true, 11: loc.tier !== 'State', 12: loc.tier === 'City' || loc.tier === 'Village' || loc.tier === 'Town', 13: loc.tier === 'Village' || loc.tier === 'Town' };
    for (const [r, label] of fields) {
      if (!show[r]) continue;
      const g = el('div', 'field-group');
      g.appendChild(el('label', 'field-label', label));
      const items = app.listValues(this.qsi, r, 6);
      const cur = this.q(r);
      const wrapSel = el('div', 'select-wrap');
      let input;
      if (items.length) {
        input = el('select', 'field-input');
        input.appendChild(new Option('— Choose —', ''));
        let found = cur === null;
        for (const it of items) {
          const t = displayText(it);
          const o = new Option(t, t);
          if (!found && cur !== null && t.toLowerCase() === String(cur).toLowerCase()) { o.selected = true; found = true; }
          input.appendChild(o);
        }
        if (!found) { const o = new Option(String(cur), String(cur)); o.selected = true; input.appendChild(o); }
        input.addEventListener('change', () => this.setLocation(r, input.value || null));
      } else {
        input = el('input', 'field-input');
        input.value = cur === null ? '' : String(cur);
        input.addEventListener('change', () => this.setLocation(r, input.value.trim() || null));
      }
      wrapSel.appendChild(input);
      g.appendChild(wrapSel);
      grid.appendChild(g);
    }
    card.appendChild(grid);
    return card;
  }

  async setLocation(r, v) {
    const app = this.app;
    // the app clears the levels below the tier and refills from Metadata when that is on
    await app.setFromList(this.qsi, r, 6, v);
    this.page = 0;
    this.render();
  }

  rowEl(r) {
    const app = this.app;
    const tr = el('tr');
    const added = r > this.last;
    if (added) tr.className = 'row-added';
    for (const c of [COL.sector, COL.subsector, COL.question, COL.fuel]) tr.appendChild(el('td', null, displayText(this.v(r, c))));
    for (const c of [COL.value, COL.unit, COL.source, COL.year]) {
      const td = el('td', 'sv-in');
      const inp = el('input', 'sv-input' + (c === COL.value ? ' num' : ''));
      const v = this.v(r, c);
      inp.value = v === null ? '' : typeof v === 'number' ? numToText(v) : String(v);
      inp.dataset.key = r + ':' + c;
      const orig = this.app.originalValue(this.si, r, c);
      if (!added && app.state.edits[[this.si, r, c].join(',')] !== undefined && orig !== v) {
        td.classList.add('sv-changed');
        td.title = 'Original value: ' + (orig === null ? '(blank)' : displayText(orig));
      }
      inp.addEventListener('change', () => this.setCell(r, c, inp.value.trim()));
      inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') inp.blur(); });
      td.appendChild(inp);
      tr.appendChild(td);
    }
    const act = el('td', 'row-act');
    if (added) {
      const del = el('button', 'icon-btn', '✕');
      del.title = 'Delete this row';
      del.addEventListener('click', () => this.deleteRow(r));
      act.appendChild(del);
    } else if ([COL.value, COL.unit, COL.source, COL.year].some((c) => app.state.edits[[this.si, r, c].join(',')] !== undefined && this.app.originalValue(this.si, r, c) !== this.v(r, c))) {
      const rv = el('button', 'icon-btn', '↺');
      rv.title = 'Restore the original values';
      rv.addEventListener('click', () => this.revertRow(r));
      act.appendChild(rv);
    }
    tr.appendChild(act);
    return tr;
  }

  parse(c, text) {
    if (text === '') return null;
    if (c === COL.value || c === COL.year) {
      const p = this.app.parseInput(text, {});
      if (typeof p.v === 'number') return p.v;
      if (c === COL.value) return undefined;
    }
    return text;
  }

  async setCell(r, c, text) {
    const v = this.parse(c, text);
    if (v === undefined) { await this.app.alert('Metadata', 'The value must be a number.'); this.render(); return; }
    this.app.applyCells([[this.si, r, c, v]]);
    if (c === COL.value && usesSecondaryData(this.app)) this.autoFill(false);
    this.render();
  }

  revertRow(r) {
    const list = [COL.value, COL.unit, COL.source, COL.year].map((c) => [this.si, r, c, this.app.originalValue(this.si, r, c)]);
    this.app.applyCells(list);
    for (const [, , c] of list) delete this.app.state.edits[[this.si, r, c].join(',')];
    if (usesSecondaryData(this.app)) this.autoFill(false);
    this.render();
  }

  deleteRow(r) {
    const list = [];
    for (let c = COL.code; c <= COL.year; c++) list.push([this.si, r, c, null]);
    this.app.applyCells(list);
    for (const [, , c] of list) delete this.app.state.edits[[this.si, r, c].join(',')];
    this.app.saveSoon();
    if (usesSecondaryData(this.app)) this.autoFill(false);
    this.render();
  }

  freeRow() {
    for (let r = this.last + 1; r <= this.end; r++) {
      let empty = true;
      for (let c = COL.code; c <= COL.year; c++) if (this.v(r, c) !== null) { empty = false; break; }
      if (empty) return r;
    }
    return -1;
  }

  async setAuto(on) {
    const app = this.app;
    if (on) {
      if (!(await app.confirm('Copy the metadata values into the Questionnaire? Answers that have metadata will be replaced.', 'Use secondary data'))) { this.render(); return; }
      // the same answer as clicking "Yes" next to "use secondary data" in the Questionnaire
      const si = this.qsi;
      const o25 = app.shapeOverride(si, 'Oval 25', true), o26 = app.shapeOverride(si, 'Oval 26', true);
      o25.fill = '000000'; o26.fill = 'FFFFFF'; o26.line = '000000';
      this.autoFill(true);
    } else {
      await app.runMacro('DoNothing', { sheet: this.qsi, caller: 'Oval 26' });
    }
    app.saveSoon();
    this.render();
  }

  autoFill(announce) {
    const n = copyMetadata(this.app);
    if (announce || n) this.app.toast(`Questionnaire filled from Metadata (${n} value${n === 1 ? '' : 's'})`);
  }

  // ------------------------------------------------------------------ add dialog
  addDialog() {
    const app = this.app;
    const loc = this.location();
    const back = el('div', 'modal-back');
    const box = el('div', 'modal modal-wide');
    box.setAttribute('role', 'dialog');
    box.appendChild(el('h2', null, 'Add metadata'));
    box.appendChild(el('p', 'muted', `${loc.tier || ''} · ${app.location().parts.join(', ')}`));
    const grid = el('div', 'field-grid');
    const field = (label, input) => { const g = el('div', 'field-group'); g.append(el('label', 'field-label', label), input); grid.appendChild(g); return input; };
    const mkSel = () => el('select', 'field-input');
    const sector = field('Sector', mkSel());
    const subsector = field('Sub-sector', mkSel());
    const question = field('Question', mkSel());
    const fuel = field('Fuel (if any)', el('input', 'field-input'));
    const value = field('Value', el('input', 'field-input'));
    const unit = field('Unit', el('input', 'field-input'));
    const source = field('Source', el('input', 'field-input'));
    const year = field('Year', el('input', 'field-input'));
    value.inputMode = 'decimal'; year.inputMode = 'numeric';
    const fill = (sel, items) => { const keep = sel.value; sel.innerHTML = ''; for (const x of items) sel.appendChild(new Option(x, x)); if (items.includes(keep)) sel.value = keep; };
    const uniq = (a) => [...new Set(a.filter((x) => x !== null && x !== undefined).map(String))];
    const sync = () => {
      fill(subsector, uniq(this.codes.filter((x) => String(x.sector) === sector.value).map((x) => x.subsector ?? '—')));
      fill(question, uniq(this.codes.filter((x) => String(x.sector) === sector.value && String(x.subsector ?? '—') === subsector.value).map((x) => x.question)));
      // suggest the unit used elsewhere for this question
      for (let r = this.first; r <= this.last; r += 1) {
        if (this.v(r, COL.question) === question.value && this.v(r, COL.sector) === sector.value) { if (!unit.value) unit.value = displayText(this.v(r, COL.unit)); break; }
      }
    };
    fill(sector, uniq(this.codes.map((x) => x.sector)));
    sector.addEventListener('change', () => { unit.value = ''; sync(); });
    subsector.addEventListener('change', () => { unit.value = ''; sync(); });
    question.addEventListener('change', () => { unit.value = ''; sync(); });
    sync();
    box.appendChild(grid);
    const err = el('div', 'login-error');
    box.appendChild(err);
    const btns = el('div', 'modal-buttons');
    const ok = el('button', 'primary', 'Add');
    const cancel = el('button', null, 'Cancel');
    btns.append(ok, cancel);
    box.appendChild(btns);
    back.appendChild(box);
    document.body.appendChild(back);
    sector.focus();
    cancel.addEventListener('click', () => back.remove());
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') back.remove(); });
    ok.addEventListener('click', () => {
      const v = this.parse(COL.value, value.value.trim());
      if (v === undefined || v === null) { err.textContent = 'Enter a numeric value.'; value.focus(); return; }
      const code = this.codes.find((x) => String(x.sector) === sector.value && String(x.subsector ?? '—') === subsector.value && x.question === question.value);
      if (!code) { err.textContent = 'Choose a question.'; return; }
      const r = this.freeRow();
      if (r < 0) { err.textContent = 'The table is full.'; return; }
      const yr = this.parse(COL.year, year.value.trim());
      const t = (x) => (x === '' ? null : x);
      const cells = {
        [COL.code]: code.code, [COL.tier]: loc.tier === 'State' ? 'State' : 'District', [COL.state]: loc.state,
        [COL.district]: loc.tier === 'State' ? null : loc.district,
        [COL.city]: ['City', 'Village', 'Town'].includes(loc.tier) ? loc.city : null,
        [COL.village]: ['Village', 'Town'].includes(loc.tier) ? loc.village : null,
        [COL.sector]: code.sector, [COL.subsector]: code.subsector, [COL.question]: code.question,
        [COL.fuel]: t(fuel.value.trim()), [COL.value]: v, [COL.unit]: t(unit.value.trim()), [COL.source]: t(source.value.trim()) || 'Added in the web app', [COL.year]: yr,
      };
      app.applyCells(Object.entries(cells).map(([c, x]) => [this.si, r, +c, x]));
      back.remove();
      if (usesSecondaryData(app)) this.autoFill(true);
      else app.toast('Metadata row added');
      this.search = ''; this.sector = '';
      this.render();
    });
  }

  refresh() { this.render(); }
}
