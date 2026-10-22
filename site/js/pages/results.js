// Results: the emissions profile, BAU dashboard and scenario comparison of the selected
// government tier, as cards of live charts and tables read from the workbook.
import { el, sheetTable, contentRows, displayText } from './sheetview.js';
import { ChartView } from '../charts.js';

const TIER_SHEETS = {
  State: ['Base- Emission Profile_State', 'Dashboard- BAU-State'],
  District: ['Base - Emission Profile_Distric', 'Dashboard- BAU-District'],
  City: ['Base- Emission Profile_City', 'Dashboard- BAU-City'],
  Village: ['Base- Emission Profile_Village', 'Dashboard- BAU-Village'],
  Town: ['Base- Emission Profile_Village', 'Dashboard- BAU-Village'],
};

function chartsOf(app, si) {
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

export class ResultsPage {
  constructor(app) {
    this.app = app;
    this.tab = 0;
    this.charts = [];
  }
  tier() { return String(this.app.eng.getValue(this.app.eng.sheetIndex('Questionnaire'), 9, 6) || 'District'); }
  tabs() {
    const [base, dash] = TIER_SHEETS[this.tier()] || TIER_SHEETS.District;
    return [
      { title: 'Emissions profile', sheet: base, desc: 'Base-year emissions by sector and sub-sector.' },
      { title: 'BAU dashboard', sheet: dash, desc: 'Business-as-usual emissions for the interim and target years.' },
      { title: 'Scenario comparison', sheet: 'Dashboard- Scenario Comparison', desc: 'BAU, E&P and high-ambition scenarios against the target.' },
    ];
  }
  mount(host) { this.host = host; this.render(); }
  unmount() { this.destroyCharts(); }
  destroyCharts() { for (const c of this.charts) c.destroy(); this.charts = []; if (this.ro) this.ro.disconnect(); }

  render() {
    const app = this.app, host = this.host;
    this.destroyCharts();
    host.innerHTML = '';
    const wrap = el('div', 'page-wrap');
    const hero = el('section', 'hero hero-sm');
    const h1 = el('h1', 'hero-title');
    h1.appendChild(el('span', 'gradient-text', 'Results'));
    const loc = app.location().parts;
    hero.append(h1, el('p', 'hero-sub', `${this.tier()} level · ${loc.length ? loc.join(', ') : 'no location selected'}`));
    wrap.appendChild(hero);
    const tabs = this.tabs();
    const bar = el('div', 'subsection-tabs');
    tabs.forEach((t, i) => {
      const b = el('button', 'sub-tab' + (i === this.tab ? ' active' : ''), t.title);
      b.addEventListener('click', () => { this.tab = i; this.render(); });
      bar.appendChild(b);
    });
    wrap.appendChild(bar);
    const t = tabs[this.tab];
    const si = app.eng.sheetIndex(t.sheet);
    if (si < 0) { host.appendChild(wrap); return; }

    const ids = chartsOf(app, si);
    const widths = chartWidths(app, si);
    const grid = el('div', 'chart-grid');
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
    if (ids.length) wrap.appendChild(grid);

    // the sheet's own table(s)
    const m = app.modelFor(si);
    const sh = app.eng.sheets[si];
    const rows = contentRows(app, si, 1, Math.min(sh.maxRow, 400), 1, Math.min(sh.maxCol, 40));
    if (rows.length) {
      let c1 = 99, c2 = 1;
      for (const r of rows) for (let c = 1; c <= Math.min(sh.maxCol, 40); c++) {
        if (m.colHidden(c)) continue;
        if (displayText(app.eng.getValue(si, r, c)) !== '') { c1 = Math.min(c1, c); c2 = Math.max(c2, c); }
      }
      const card = el('div', 'card');
      card.appendChild(el('h3', 'card-title', t.desc));
      const box = el('div', 'sv-scroll');
      box.appendChild(sheetTable(app, si, rows, { cols: [c1, c2] }));
      card.appendChild(box);
      const open = el('a', 'btn btn-ghost btn-sm', 'Open this sheet in the workbook');
      open.href = '#/workbook';
      open.addEventListener('click', () => { app.pendingSheet = si; });
      card.appendChild(open);
      wrap.appendChild(card);
    }
    host.appendChild(wrap);
  }
  refresh() { if (!this.host) return; const keep = this.host.scrollTop; this.render(); this.host.scrollTop = keep; }
}
