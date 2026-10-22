// Questionnaire as a sectioned form. Each tab runs the workbook's own section macro
// (TaketoBuildingEnergy, ...), so the rows shown are exactly the ones Excel shows, and the
// cells are rendered as inputs bound to the calculation engine.
import { el, sheetTable, contentRows, shownValue } from './sheetview.js';
import { KEY } from '../engine/engine.js';

const Q = 'Questionnaire';
export const SECTIONS = [
  { title: 'Basic Information', short: 'Info', macro: 'Taketobasicinfo', shape: 'Rectangle: Rounded Corners 4', desc: 'Choose the government tier and location, then set population, climate and growth assumptions.' },
  { title: 'Buildings & Energy', short: 'Energy', macro: 'TaketoBuildingEnergy', shape: 'Rectangle: Rounded Corners 7', desc: 'Fuel and electricity used by residential, commercial, institutional and industrial buildings.' },
  { title: 'Electricity Generation', short: 'Electricity', macro: 'TaketoElectricity', shape: 'Rectangle: Rounded Corners 18', desc: 'Power generation inside the boundary (state level only).' },
  { title: 'Transportation', short: 'Transport', macro: 'TaketoTransportation', shape: 'Rectangle: Rounded Corners 1', desc: 'Fuel sales or vehicle kilometres travelled, by mode.' },
  { title: 'Solid Waste', short: 'Waste', macro: 'TaketoSolidWaste', shape: 'Rectangle: Rounded Corners 9', desc: 'Waste generation, composition and treatment.' },
  { title: 'Wastewater', short: 'Wastewater', macro: 'TaketoWastewater', shape: 'Rectangle: Rounded Corners 12', desc: 'Domestic and industrial wastewater treatment and discharge.' },
  { title: 'Industry (IPPU)', short: 'Industry', macro: 'Taketoippu', shape: 'Rectangle: Rounded Corners 16', desc: 'Industrial processes and product use.' },
  { title: 'Agriculture & Land (AFOLU)', short: 'Agriculture', macro: 'TaketoAFOLU', shape: 'Rectangle: Rounded Corners 14', desc: 'Livestock, crops, land use and forests.' },
];

export class QuestionnairePage {
  constructor(app) {
    this.app = app;
    this.si = app.eng.sheetIndex(Q);
    this.k = 0;
    try { this.k = Math.max(0, Math.min(SECTIONS.length - 1, +(localStorage.getItem('ascent-q-section') || 0))); } catch (e) { /* ignore */ }
  }

  sectionEnabled(k) {
    const s = SECTIONS[k];
    const obj = this.app.findShape(this.si, s.shape);
    const ov = this.app.shapeOverride(this.si, s.shape, false);
    const macro = ov && ov.onAction !== undefined ? ov.onAction : obj && obj.macro;
    return !!macro;
  }

  async mount(host) {
    this.host = host;
    if (!this.sectionEnabled(this.k)) this.k = 0;
    await this.runSection();
    this.render();
  }

  async runSection() {
    const app = this.app;
    app.activeSi = this.si;
    await app.runMacro(SECTIONS[this.k].macro, { sheet: this.si, caller: SECTIONS[this.k].shape });
  }

  async go(k) {
    if (k === this.k || !this.sectionEnabled(k)) return;
    this.k = k;
    try { localStorage.setItem('ascent-q-section', String(k)); } catch (e) { /* ignore */ }
    await this.runSection();
    this.render();
    this.host.scrollTop = 0;
  }

  editable(r, c, xf, fc) {
    if (!xf.unlocked || r < 7) return false;
    if (fc && fc.kind) return false;
    // unlocked cells without input shading hold labels (e.g. the transport option texts)
    return !!xf.fill || !!this.app.modelFor(this.si).dvAt(r, c);
  }

  ovals() {
    const out = [];
    const d = this.app.datas[this.si];
    const texts = [];
    for (const dr of d.drawing || []) {
      const o = dr.obj;
      if (o.kind !== 'sp' || !dr.from) continue;
      if (o.geom === 'ellipse' && o.macro) out.push({ name: o.name, macro: o.macro, r: dr.from[2] + 1, c: dr.from[0] + 1, obj: o });
      else if (!o.macro && o.text && o.text.paras) {
        const t = o.text.paras.map((p) => p.runs.map((x) => x.t).join('')).join(' ').trim();
        if (t && t.length < 20) texts.push({ r: dr.from[2] + 1, c: dr.from[0] + 1, t });
      }
    }
    for (const o of out) {
      const same = texts.filter((t) => t.r === o.r).sort((a, b) => Math.abs(a.c - o.c) - Math.abs(b.c - o.c));
      o.label = same.length ? same[0].t : '';
    }
    return out;
  }
  ovalOn(o) {
    const ov = this.app.shapeOverride(this.si, o.name, false);
    const fill = ov && ov.fill !== undefined && !ov.fillHidden ? ov.fill : o.obj.fill && (o.obj.fill.c || o.obj.fill);
    return typeof fill === 'string' && fill.toUpperCase() === '000000';
  }

  render() {
    const app = this.app, si = this.si, m = app.modelFor(si);
    const host = this.host;
    const focus = document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.addr : null;
    const scroll = host.scrollTop;
    host.innerHTML = '';
    const wrap = el('div', 'page-wrap');
    const hero = el('section', 'hero');
    const h1 = el('h1', 'hero-title');
    h1.appendChild(el('span', 'gradient-text', 'GHG Inventory'));
    hero.append(h1, el('p', 'hero-sub', 'Fill in the questionnaire to build the base-year GHG emissions inventory. Every answer feeds the same calculations as the ASCENT workbook.'));
    // section progress tabs
    const prog = el('div', 'progress-bar-wrap');
    const ps = el('div', 'progress-sections');
    SECTIONS.forEach((s, k) => {
      const b = el('button', 'ps' + (k === this.k ? ' active' : '') + (k < this.k ? ' done' : ''), s.short);
      const on = this.sectionEnabled(k);
      b.disabled = !on;
      if (!on) b.title = 'Not part of the inventory for this government tier';
      b.addEventListener('click', () => this.go(k));
      ps.appendChild(b);
    });
    const track = el('div', 'progress-track');
    const fill = el('div', 'progress-fill');
    fill.style.width = ((this.k + 1) / SECTIONS.length * 100).toFixed(1) + '%';
    track.appendChild(fill);
    prog.append(ps, track);
    hero.appendChild(prog);
    wrap.appendChild(hero);

    const s = SECTIONS[this.k];
    const sec = el('section', 'form-card');
    const head = el('div', 'section-header');
    head.appendChild(el('div', 'section-number', String(this.k + 1).padStart(2, '0')));
    const ht = el('div');
    ht.append(el('h2', 'section-title', s.title), el('p', 'section-desc', s.desc));
    head.appendChild(ht);
    sec.appendChild(head);

    const ed = (r, c, xf, fc) => this.editable(r, c, xf, fc);
    const ovals = this.ovals().filter((o) => !m.rowHidden(o.r));
    const ovalRows = new Set(ovals.map((o) => o.r));
    // location levels the tier does not use (the workbook hides them when the tier changes)
    const notUsed = new Set({ State: [11, 12, 13], District: [12, 13], City: [13], Village: [12], Town: [12] }[app.eng.getValue(si, 9, 6)] || []);
    const rows = contentRows(app, si, 7, 806, 3, 23, { editable: (r, c) => ed(r, c, m.xf(r, c), app.eng.sheets[si].formulas.get(KEY(r, c))) })
      .concat([...ovalRows]).filter((r, i, a) => a.indexOf(r) === i && !notUsed.has(r)).sort((a, b) => a - b);
    // columns actually used by the visible rows
    let maxC = 6;
    for (const r of rows) for (let c = 23; c > maxC; c--) {
      if (m.colHidden(c)) continue;
      if (shownValue(app, si, r, c) !== null || ed(r, c, m.xf(r, c), app.eng.sheets[si].formulas.get(KEY(r, c)))) { maxC = c; break; }
    }
    for (const o of ovals) maxC = Math.max(maxC, o.c);
    const table = sheetTable(app, si, rows, {
      cols: [3, maxC],
      editable: ed,
      onEdit: (r, c, text) => this.edit(r, c, text),
      onPick: (r, c, v) => this.pick(r, c, v),
      keep: (r, c) => ovals.some((o) => o.r === r && ((m.mergeAt(r, o.c) || [0, o.c])[1]) === c),
      extra: (r, c, td) => {
        if (!ovalRows.has(r)) return;
        const mine = ovals.filter((o) => {
          if (o.r !== r) return false;
          const mg = m.mergeAt(r, o.c);
          return (mg ? mg[1] : o.c) === c;
        });
        if (!mine.length) return;
        const grp = el('div', 'sv-options');
        for (const o of mine) {
          const lab = el('label', 'sv-option' + (this.ovalOn(o) ? ' on' : ''));
          const rb = el('input');
          rb.type = 'radio';
          rb.name = 'opt-' + r;
          rb.checked = this.ovalOn(o);
          rb.addEventListener('change', () => this.option(o));
          lab.appendChild(rb);
          if (o.label) lab.appendChild(el('span', null, o.label));
          grp.appendChild(lab);
        }
        td.prepend(grp);
      },
    });
    const scrollBox = el('div', 'sv-scroll');
    scrollBox.appendChild(table);
    sec.appendChild(scrollBox);

    const nav = el('div', 'form-nav');
    const prev = el('button', 'btn btn-ghost', '← Previous');
    const next = el('button', 'btn btn-primary', 'Next section →');
    const findNext = (d) => { let k = this.k + d; while (k >= 0 && k < SECTIONS.length && !this.sectionEnabled(k)) k += d; return k >= 0 && k < SECTIONS.length ? k : -1; };
    const pk = findNext(-1), nk = findNext(1);
    prev.disabled = pk < 0; next.disabled = nk < 0;
    prev.addEventListener('click', () => this.go(pk));
    next.addEventListener('click', () => this.go(nk));
    const res = el('a', 'btn btn-amber', 'See results');
    res.href = '#/results';
    nav.append(prev, el('span', 'grow'), res, next);
    sec.appendChild(nav);
    wrap.appendChild(sec);
    host.appendChild(wrap);
    host.scrollTop = scroll;
    if (focus) { const f = host.querySelector(`[data-addr="${focus}"]`); if (f) f.focus({ preventScroll: true }); }
  }

  async edit(r, c, text) {
    const ok = await this.app.userEdit(this.si, r, c, text);
    if (ok) await this.after(r, c);
    return ok;
  }
  async pick(r, c, v) {
    await this.app.setFromList(this.si, r, c, v);
    await this.after(r, c);
  }
  async after(r, c) {
    // the tier change macro shows every row again; restore this section's view like Excel
    if (r === 9 && c === 6) {
      if (!this.sectionEnabled(this.k)) this.k = 0;
      await this.runSection();
    }
    this.render();
  }
  async option(o) {
    await this.app.runMacro(o.macro, { sheet: this.si, caller: o.name });
    this.render();
  }
  refresh() { this.render(); }
}
