// Emission Factors: the "Emission Factor" sheet (and the GWP values it uses) as editable
// tables. Every factor is a plain input cell of the workbook, so changing one recalculates
// the inventory, scenarios and dashboards.
import { el, sheetTable, contentRows, shownValue, displayText } from './sheetview.js';
import { KEY } from '../engine/engine.js';
import { numToText } from '../engine/values.js';

const EF = 'Emission Factor';

export class FactorsPage {
  constructor(app) {
    this.app = app;
    this.si = app.eng.sheetIndex(EF);
    this.gwp = app.eng.sheetIndex('GWP');
    this.sections = this.findSections();
    this.k = 0;
  }

  // a section starts at a row whose only content is a bold title in column C
  findSections() {
    const app = this.app, si = this.si, m = app.modelFor(si), sh = app.eng.sheets[si];
    const titles = [];
    for (let r = 1; r <= sh.maxRow; r++) {
      const v = app.eng.getValue(si, r, 3);
      if (typeof v !== 'string' || !v.trim()) continue;
      const xf = m.xf(r, 3);
      if (!xf.font || !xf.font.b) continue;
      const mg = m.mergeAt(r, 3);
      let other = false;
      for (let c = mg ? mg[3] + 1 : 4; c <= Math.min(sh.maxCol, 34); c++) { const x = app.eng.getValue(si, r, c); if (x !== null && x !== '') { other = true; break; } }
      if (!other) titles.push([r, v.trim()]);
    }
    return titles.map(([r, t], i) => ({ title: t, r1: r + 1, r2: i + 1 < titles.length ? titles[i + 1][0] - 1 : sh.maxRow }));
  }

  editable(si, r, c) {
    const sh = this.app.eng.sheets[si];
    if (sh.formulas.get(KEY(r, c))) return false;
    return typeof this.app.eng.getValue(si, r, c) === 'number';
  }

  mount(host) { this.host = host; this.render(); }

  render() {
    const app = this.app, host = this.host;
    const scroll = host.scrollTop;
    const focus = document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.addr : null;
    host.innerHTML = '';
    const wrap = el('div', 'page-wrap');
    const hero = el('section', 'hero hero-sm');
    const h1 = el('h1', 'hero-title');
    h1.appendChild(el('span', 'gradient-text', 'Emission Factors'));
    hero.append(h1, el('p', 'hero-sub', 'Update a factor when a new source is published. Every inventory, scenario and dashboard figure recalculates straight away, and your changes are kept in this browser.'));
    wrap.appendChild(hero);

    const changed = this.changedCount();
    const tools = el('div', 'page-tools');
    if (changed) tools.appendChild(el('span', 'pill pill-amber', `${changed} factor${changed > 1 ? 's' : ''} changed`));
    tools.appendChild(el('span', 'grow'));
    const reset = el('button', 'btn btn-ghost btn-sm', 'Restore factors');
    reset.disabled = !changed;
    reset.addEventListener('click', () => this.restore());
    tools.appendChild(reset);
    wrap.appendChild(tools);

    const layout = el('div', 'side-layout');
    const side = el('nav', 'side-nav');
    side.appendChild(el('div', 'side-label', 'Sector'));
    this.sections.forEach((s, i) => {
      const b = el('button', 'side-item' + (i === this.k ? ' active' : ''), s.title);
      b.addEventListener('click', () => { this.k = i; this.render(); });
      side.appendChild(b);
    });
    side.appendChild(el('div', 'side-label', 'Global'));
    const gb = el('button', 'side-item' + (this.k === -1 ? ' active' : ''), 'Global Warming Potential');
    gb.addEventListener('click', () => { this.k = -1; this.render(); });
    side.appendChild(gb);
    layout.appendChild(side);

    const card = el('section', 'card');
    if (this.k === -1) {
      card.appendChild(el('h2', 'section-title', 'Global Warming Potential (AR6)'));
      card.appendChild(el('p', 'section-desc', 'CO₂-equivalence of methane and nitrous oxide, used for every CH₄ and N₂O factor.'));
      const rows = contentRows(app, this.gwp, 1, 20, 3, 5);
      card.appendChild(this.table(this.gwp, rows, [3, 5]));
    } else {
      const s = this.sections[this.k];
      card.appendChild(el('h2', 'section-title', s.title));
      const efRows = [];
      for (let r = s.r1; r <= s.r2; r++) { const v = shownValue(app, this.si, r, 3); if (typeof v === 'string' && v.startsWith('EF_')) efRows.push(r); }
      if (efRows.length) {
        const g = (r, c) => displayText(app.eng.getValue(this.gwp, r, c));
        card.appendChild(el('p', 'section-desc', `Factors per unit of activity. CH₄ and N₂O are in their own units and converted to CO₂e with the GWP values (CH₄ ${g(4, 5)}, N₂O ${g(5, 5)}).`));
        card.appendChild(this.factorTable(efRows));
        layout.appendChild(card);
        wrap.appendChild(layout);
        host.appendChild(wrap);
        host.scrollTop = scroll;
        if (focus) { const f = host.querySelector(`[data-addr="${focus}"]`); if (f) f.focus({ preventScroll: true }); }
        return;
      }
      card.appendChild(el('p', 'section-desc', 'Edit the shaded values; the other columns are calculated.'));
      const rows = contentRows(app, this.si, s.r1, s.r2, 3, 34);
      const m = app.modelFor(this.si);
      let c2 = 3;
      for (const r of rows) for (let c = 34; c > c2; c--) { if (m.colHidden(c)) continue; const v = app.eng.getValue(this.si, r, c); if (v !== null && v !== '') { c2 = c; break; } }
      card.appendChild(this.table(this.si, rows, [3, c2]));
    }
    layout.appendChild(card);
    wrap.appendChild(layout);
    host.appendChild(wrap);
    host.scrollTop = scroll;
    if (focus) { const f = host.querySelector(`[data-addr="${focus}"]`); if (f) f.focus({ preventScroll: true }); }
  }

  // the standard factor table: activity, unit, CO2 / CH4 / N2O factors and the CO2e total
  factorTable(rows) {
    const app = this.app, si = this.si, m = app.modelFor(si);
    const cols = [[4, 'Activity'], [7, 'Unit'], [8, 'CO₂'], [9, 'CH₄'], [12, 'N₂O'], [15, 'Total CO₂e']];
    const t = el('table', 'data-table ef-table');
    const htr = el('tr');
    for (const [, h] of cols) htr.appendChild(el('th', null, h));
    const thead = el('thead'); thead.appendChild(htr); t.appendChild(thead);
    const tb = el('tbody');
    for (const r of rows) {
      const tr = el('tr');
      for (const [c] of cols) {
        if (c === 4 || c === 7) { tr.appendChild(el('td', c === 4 ? 'sv-bold' : 'muted', displayText(app.eng.getValue(si, r, c)))); continue; }
        const fc = app.eng.sheets[si].formulas.get(KEY(r, c));
        const v = app.eng.getValue(si, r, c);
        if (fc || c === 15) {
          const td = el('td', 'sv-calc sv-num');
          td.appendChild(el('span', null, displayText(v, m.xf(r, c).fmt && !/^_/.test(m.xf(r, c).fmt) ? m.xf(r, c).fmt : '0.0####')));
          tr.appendChild(td);
          continue;
        }
        const td = el('td', 'sv-in');
        const inp = el('input', 'sv-input num');
        inp.value = v === null ? '' : numToText(v);
        inp.dataset.orig = inp.value;
        inp.dataset.addr = r + ',' + c;
        inp.inputMode = 'decimal';
        inp.setAttribute('aria-label', `${displayText(app.eng.getValue(si, r, 4))} ${cols.find((x) => x[0] === c)[1]}`);
        inp.addEventListener('change', async () => { if (!(await this.edit(si, r, c, inp.value.trim()))) inp.value = inp.dataset.orig; });
        inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') inp.blur(); if (e.key === 'Escape') { inp.value = inp.dataset.orig; inp.blur(); } });
        const orig = app.originalValue(si, r, c);
        if (app.state.edits[[si, r, c].join(',')] !== undefined && orig !== v) {
          td.classList.add('sv-changed');
          td.title = 'Workbook value: ' + (orig === null ? '(blank)' : numToText(orig));
        }
        td.appendChild(inp);
        tr.appendChild(td);
      }
      tb.appendChild(tr);
    }
    t.appendChild(tb);
    const box = el('div', 'sv-scroll');
    box.appendChild(t);
    return box;
  }

  table(si, rows, cols) {
    const box = el('div', 'sv-scroll');
    const t = sheetTable(this.app, si, rows, {
      cols,
      noTitles: si === this.gwp,
      editable: (r, c) => this.editable(si, r, c),
      onEdit: (r, c, text) => this.edit(si, r, c, text),
      extra: (r, c, td) => {
        const orig = this.app.originalValue(si, r, c);
        const v = this.app.eng.getValue(si, r, c);
        if (this.app.state.edits[[si, r, c].join(',')] !== undefined && orig !== v) {
          td.classList.add('sv-changed');
          td.title = 'Workbook value: ' + (orig === null ? '(blank)' : numToText(orig));
        }
      },
    });
    box.appendChild(t);
    return box;
  }

  async edit(si, r, c, text) {
    const m = this.app.modelFor(si);
    const p = this.app.parseInput(text, m.xf(r, c));
    if (p.error || (p.v !== null && typeof p.v !== 'number')) { await this.app.alert('Emission factor', 'Enter a number.'); return false; }
    this.app.applyCells([[si, r, c, p.v]]);
    this.render();
    return true;
  }

  changedCount() {
    let n = 0;
    for (const k of Object.keys(this.app.state.edits)) {
      const [si, r, c] = k.split(',').map(Number);
      if ((si === this.si || si === this.gwp) && this.app.originalValue(si, r, c) !== this.app.eng.getValue(si, r, c)) n++;
    }
    return n;
  }

  async restore() {
    if (!(await this.app.confirm('Put every emission factor and GWP back to the workbook values?', 'Restore factors'))) return;
    const list = [];
    for (const k of Object.keys(this.app.state.edits)) {
      const [si, r, c] = k.split(',').map(Number);
      if (si === this.si || si === this.gwp) list.push([si, r, c, this.app.originalValue(si, r, c)]);
    }
    if (list.length) this.app.applyCells(list);
    for (const [si, r, c] of list) delete this.app.state.edits[[si, r, c].join(',')];
    this.app.saveSoon();
    this.render();
  }

  refresh() { this.render(); }
}
