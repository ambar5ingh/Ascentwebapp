// A page made of tabs, each showing one worksheet as cards: its charts (live Plotly views)
// and its cells as a table, with inputs where the workbook has input cells.
import { el, sheetTable, contentRows, displayText } from './sheetview.js';
import { ChartView } from '../charts.js';
import { KEY } from '../engine/engine.js';

export const TIER_SHEETS = {
  State: ['Base- Emission Profile_State', 'Dashboard- BAU-State'],
  District: ['Base - Emission Profile_Distric', 'Dashboard- BAU-District'],
  City: ['Base- Emission Profile_City', 'Dashboard- BAU-City'],
  Village: ['Base- Emission Profile_Village', 'Dashboard- BAU-Village'],
  Town: ['Base- Emission Profile_Village', 'Dashboard- BAU-Village'],
};

export function chartsOf(app, si) {
  const out = [];
  const walk = (o) => { if (o.kind === 'graphicFrame' && o.chart !== undefined && app.wb.charts[o.chart]) out.push(o.chart); (o.children || []).forEach(walk); };
  for (const d of app.datas[si].drawing || []) walk(d.obj);
  return out;
}

// width in px (at 100% zoom) of each chart as placed on the sheet
function chartWidths(app, si) {
  const m = app.modelFor(si);
  if (!m.colX) m.layout(app.zoomOf(si));
  const out = new Map();
  const walk = (o, w) => { if (o.kind === 'graphicFrame' && o.chart !== undefined) out.set(o.chart, w); for (const ch of o.children || []) walk(ch, o.box && ch.box && o.box[2] ? w * ch.box[2] / (o.chbox || o.box)[2] : w); };
  for (const d of app.datas[si].drawing || []) {
    let w = 600;
    if (d.ext) w = d.ext[0] / 9525;
    else if (d.from && d.to) w = (m.anchorX(d.to[0], d.to[1]) - m.anchorX(d.from[0], d.from[1])) / m.zoom;
    walk(d.obj, w);
  }
  return out;
}

function fallbackTitle(spec, i) {
  const type = spec && spec.plots && spec.plots[0] && spec.plots[0].type;
  if (type === 'pie' || type === 'doughnut') return 'Share of emissions by sector';
  return i === 0 ? 'Emissions by sub-sector (tCO₂e)' : 'Emissions (tCO₂e)';
}

function chartTitle(spec, fallback) {
  const t = spec && spec.title;
  if (t && t.text) return t.text.replace(/\s+/g, ' ').trim();
  return fallback;
}

const INPUT_FILLS = new Set(['B4C7E7', 'B4C6E7', 'DAE3F3']);

export const tierOf = (app) => String(app.eng.getValue(app.eng.sheetIndex('Questionnaire'), 9, 6) || 'District');

// Subclasses define title, subtitle() and tabs() -> [{ title, sheet, desc, rows: [r1, r2],
// cols: [c1, c2], charts: (ids) => ids, intro, before(card), run: async () => {} }].
export class SheetTabsPage {
  constructor(app) {
    this.app = app;
    this.tab = 0;
    this.charts = [];
  }
  subtitle() {
    const loc = this.app.location().parts;
    return `${tierOf(this.app)} level · ${loc.length ? loc.join(', ') : 'no location selected'}`;
  }
  async mount(host) { this.host = host; await this.prepare(); this.render(); }
  unmount() { this.destroyCharts(); }
  destroyCharts() { for (const c of this.charts) c.destroy(); this.charts = []; if (this.ro) this.ro.disconnect(); }
  async prepare() { const t = this.tabs()[this.tab]; if (t && t.run) await t.run(); }
  async go(i) { this.tab = i; await this.prepare(); this.render(); this.host.scrollTop = 0; }

  editable(si, r, c) {
    const xf = this.app.modelFor(si).xf(r, c);
    const fc = this.app.eng.sheets[si].formulas.get(KEY(r, c));
    if (fc) return false;
    const dv = !!this.app.modelFor(si).dvAt(r, c);
    if (this.app.isProtected(si)) return !!xf.unlocked && (!!xf.fill || dv);
    // unprotected sheets: dropdowns and the workbook's blue input cells
    return dv || (!!xf.unlocked && !!xf.fill) || INPUT_FILLS.has(String(xf.fill).toUpperCase());
  }

  render() {
    const app = this.app, host = this.host;
    const scroll = host.scrollTop;
    const focus = document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.addr : null;
    this.destroyCharts();
    host.innerHTML = '';
    const wrap = el('div', 'page-wrap');
    const hero = el('section', 'hero hero-sm');
    const h1 = el('h1', 'hero-title');
    h1.appendChild(el('span', 'gradient-text', this.title));
    hero.append(h1, el('p', 'hero-sub', this.subtitle()));
    wrap.appendChild(hero);
    const tabs = this.tabs();
    if (this.tab >= tabs.length) this.tab = 0;
    if (tabs.length > 1) {
      const bar = el('div', 'subsection-tabs');
      tabs.forEach((t, i) => {
        const b = el('button', 'sub-tab' + (i === this.tab ? ' active' : ''), t.title);
        b.addEventListener('click', () => this.go(i));
        bar.appendChild(b);
      });
      wrap.appendChild(bar);
    }
    if (this.toolbar) this.toolbar(wrap);
    const t = tabs[this.tab];
    const si = t ? app.eng.sheetIndex(t.sheet) : -1;
    if (si < 0) { host.appendChild(wrap); return; }

    // the sheet's cells
    const m = app.modelFor(si), sh = app.eng.sheets[si];
    const [r1, r2] = t.rows || [1, Math.min(sh.maxRow, 400)];
    let [c1, c2] = t.cols || [1, Math.min(sh.maxCol, 40)];
    const ed = (r, c) => this.editable(si, r, c);
    const rows = contentRows(app, si, r1, r2, c1, c2, { editable: ed });
    if (!t.cols && rows.length) {
      let a = 99, b = 1;
      for (const r of rows) for (let c = c1; c <= c2; c++) {
        if (m.colHidden(c)) continue;
        if (displayText(app.eng.getValue(si, r, c)) !== '' || ed(r, c)) { a = Math.min(a, c); b = Math.max(b, c); }
      }
      c1 = a; c2 = b;
    }
    const tableCard = () => {
      const card = el('section', 'card');
      card.appendChild(el('h3', 'card-title', t.desc));
      if (t.intro) card.appendChild(el('p', 'section-desc', t.intro));
      if (t.before) t.before(card);
      const after = () => { if (t.after) t.after(card); };
      if (rows.length) {
        const box = el('div', 'sv-scroll');
        box.appendChild(sheetTable(app, si, rows, {
          cols: [c1, c2],
          editable: ed,
          onEdit: async (r, c, text) => { const ok = await app.userEdit(si, r, c, text); if (ok) this.render(); return ok; },
          onPick: async (r, c, v) => { await app.setFromList(si, r, c, v); this.render(); },
        }));
        card.appendChild(box);
      }
      after();
      return card;
    };
    if (t.tableFirst) wrap.appendChild(tableCard());

    // charts
    const ids = (t.charts || ((x) => x))(chartsOf(app, si));
    const widths = chartWidths(app, si);
    this.ro = new ResizeObserver((entries) => {
      for (const e of entries) {
        const cv = e.target._chart;
        if (!cv) continue;
        // keep the workbook's font sizes in proportion to the chart's size on the sheet
        cv.plainScale = Math.max(0.6, Math.min(1, e.contentRect.width / (cv.sheetWidth || 600)));
        cv.resize(e.contentRect.width, e.contentRect.height);
        cv.ensure();
      }
    });
    if (ids.length) {
      const grid = el('div', 'chart-grid');
      ids.forEach((id, i) => {
        const spec = app.wb.charts[id];
        const card = el('div', 'card chart-card');
        card.appendChild(el('h3', 'card-title', chartTitle(spec, fallbackTitle(spec, i))));
        const box = el('div', 'chart-box');
        const cv = new ChartView(app, si, id, box);
        cv.plain = true;
        cv.sheetWidth = widths.get(id);
        box._chart = cv;
        this.charts.push(cv);
        card.appendChild(box);
        grid.appendChild(card);
        this.ro.observe(box);
      });
      wrap.appendChild(grid);
    }
    if (!t.tableFirst && (rows.length || t.before)) wrap.appendChild(tableCard());
    host.appendChild(wrap);
    host.scrollTop = scroll;
    if (focus) { const f = host.querySelector(`[data-addr="${focus}"]`); if (f) f.focus({ preventScroll: true }); }
  }
  refresh() { if (this.host) this.render(); }
}
