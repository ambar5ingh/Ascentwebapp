// Excel charts drawn with Plotly in a ggplot2-like theme (grey panel, white grid),
// keeping the workbook's series colours, data labels and number formats.
import { formatValue } from './format.js';
import { Ref, Matrix, isErr, Rich } from './engine/values.js';

const PALETTE = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E', '636363', '997300', '255E91', '43682B'];
const FONT = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif';
const hex = (c) => (!c ? null : typeof c === 'string' ? '#' + c : c.c ? '#' + c.c : null);

function d3fmt(fmt) {
  if (!fmt || fmt === 'General') return '';
  const sec = fmt.split(';')[0];
  const dec = (/0\.(0+)/.exec(sec) || [, ''])[1].length;
  if (sec.includes('%')) return `.${dec}%`;
  if (sec.includes(',')) return `,.${dec}f`;
  if (/E\+/i.test(sec)) return `.${dec}e`;
  return `.${dec}f`;
}

export class ChartView {
  constructor(app, si, id, el) {
    this.app = app; this.si = si; this.id = id; this.el = el;
    this.spec = app.wb.charts[id];
    this.rendered = false;
    this.sig = '';
    this.w = 0; this.h = 0;
  }
  ensure() {
    if (this.rendered || !this.w || !this.h || !window.Plotly) return;
    this.rendered = true;
    this.render(true);
  }
  resize(w, h) {
    if (Math.abs(w - this.w) < 1 && Math.abs(h - this.h) < 1 && this.zoom === this.app.zoomOf(this.si)) return;
    this.w = w; this.h = h;
    this.zoom = this.app.zoomOf(this.si);
    if (this.rendered) this.render(true);
  }
  refresh() { if (this.rendered) this.render(false); }
  destroy() { if (this.rendered && window.Plotly) window.Plotly.purge(this.el); this.rendered = false; }

  // ---------------------------------------------------------------- data
  read(ref) {
    if (!ref) return null;
    const eng = this.app.eng;
    let v = eng.evaluate(ref.replace(/^=/, ''), this.si);
    if (v instanceof Ref) {
      const m = this.app.modelFor(v.sheet);
      const rows = [], cols = [];
      for (let r = v.r1; r <= v.r2; r++) if (!m.rowHidden(r)) rows.push(r);
      for (let c = v.c1; c <= v.c2; c++) if (!m.colHidden(c)) cols.push(c);
      const fmtAt = (r, c) => m.xf(r, c).fmt || 'General';
      return { ref: v, rows, cols, get: (r, c) => eng.getValue(v.sheet, r, c), fmtAt };
    }
    if (v instanceof Matrix) {
      const rows = [...Array(v.rows).keys()], cols = [...Array(v.cols).keys()];
      return { rows, cols, get: (r, c) => v.get(r, c), fmtAt: () => 'General' };
    }
    return { rows: [0], cols: [0], get: () => v, fmtAt: () => 'General' };
  }
  // flatten a data range into a list of points along its long axis
  points(d) {
    if (!d) return [];
    const out = [];
    if (d.rows.length >= d.cols.length) { for (const r of d.rows) for (const c of d.cols.slice(0, 1)) out.push(d.get(r, c)); }
    else for (const c of d.cols) out.push(d.get(d.rows[0], c));
    return out;
  }
  categories(d, n) {
    if (!d) return [...Array(n).keys()].map((i) => String(i + 1));
    const label = (v, r, c) => {
      if (v === null || v === undefined) return '';
      if (isErr(v)) return v.code;
      if (v instanceof Rich) return v.toString();
      if (typeof v === 'number') return formatValue(v, d.fmtAt(r, c)).trim();
      return String(v);
    };
    const byRow = d.rows.length >= d.cols.length;
    const pts = byRow ? d.rows : d.cols;
    const lev = byRow ? d.cols : d.rows;
    if (lev.length <= 1) return pts.map((p) => (byRow ? label(d.get(p, lev[0]), p, lev[0]) : label(d.get(lev[0], p), lev[0], p)));
    // multi-level categories: forward-fill outer levels, show "outer · inner"
    const fill = lev.map(() => '');
    return pts.map((p) => {
      const parts = lev.map((l, i) => {
        const v = byRow ? label(d.get(p, l), p, l) : label(d.get(l, p), l, p);
        if (v !== '') fill[i] = v;
        return i < lev.length - 1 ? fill[i] : v;
      });
      const inner = parts[parts.length - 1];
      const outer = parts.slice(0, -1).filter(Boolean);
      return inner && outer.length ? inner + ' (' + outer[outer.length - 1] + ')' : inner || outer.join(' · ');
    });
  }
  seriesName(s, i) {
    if (s.tx) {
      if (s.tx.ref) {
        const d = this.read(s.tx.ref);
        if (d) {
          const vals = [];
          for (const r of d.rows) for (const c of d.cols) { const v = d.get(r, c); if (v !== null && v !== '') vals.push(v instanceof Rich ? v.toString() : String(isErr(v) ? v.code : v)); }
          if (vals.length) return vals.join(' ');
        }
      }
      if (s.tx.v) return s.tx.v;
      if (s.tx.cache && s.tx.cache.length) return s.tx.cache.join(' ');
    }
    return 'Series' + (i + 1);
  }
  num(v) {
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    return null;
  }

  // ---------------------------------------------------------------- figure
  figure() {
    const sp = this.spec, z = this.zoom || 1;
    const pt = (s) => s * 4 / 3 * z;
    const traces = [];
    let barmode = 'group', barnorm = '', anyBar = false, nBarSeries = 0, gap = 150;
    let isPie = false, horiz = false;
    const cats0 = [];
    let sIndex = 0;
    for (const plot of sp.plots) {
      const type = plot.type;
      const plotLbl = plot.dLbls || {};
      for (const s of plot.series) {
        const i = s.idx !== undefined ? +s.idx : sIndex;
        sIndex++;
        const name = this.seriesName(s, i);
        const valD = this.read((s.val || s.yVal || {}).ref);
        const vals = this.points(valD).map((v) => this.num(v));
        const fmt = (s.val || s.yVal || {}).fmt || (valD && valD.rows.length ? valD.fmtAt(valD.rows[0], valD.cols[0]) : 'General');
        const lbl = s.dLbls || plotLbl;
        const showVal = lbl.showVal === 1 || lbl.showVal === '1';
        const txt = vals.map((v) => (v === null ? '' : formatValue(v, fmt).trim()));
        const color = hex(s.fill) || hex(s.line && s.line.c) || '#' + PALETTE[i % PALETTE.length];
        const lblFont = lbl.font ? { size: pt(lbl.font.sz || 9), color: hex(lbl.font.color) || '#404040', family: FONT } : { size: pt(9), color: '#404040', family: FONT };
        if (lbl.font && lbl.font.b) lblFont.weight = 'bold';
        if (type === 'pie' || type === 'doughnut') {
          isPie = true;
          const cats = this.categories(this.read((s.cat || {}).ref), vals.length);
          const colors = vals.map((_, k) => {
            const dp = (s.dPt || []).find((p) => +p.idx === k);
            return hex(dp && dp.fill) || '#' + PALETTE[k % PALETTE.length];
          });
          const info = [];
          if (lbl.showCatName) info.push('label');
          if (showVal) info.push('text');
          if (lbl.showPercent) info.push('percent');
          traces.push({
            type: 'pie', name, labels: cats, values: vals.map((v) => (v === null ? 0 : v)), text: txt, sort: false, direction: 'clockwise',
            rotation: +(plot.firstSliceAng || 0), hole: type === 'doughnut' ? 0.5 : 0,
            marker: { colors, line: { color: '#FFFFFF', width: 1.5 } },
            textinfo: info.length ? info.join('+') : 'none', textposition: lbl.pos === 'outEnd' ? 'outside' : 'auto',
            textfont: lblFont, hovertemplate: '%{label}<br>%{text} (%{percent})<extra></extra>',
          });
          continue;
        }
        if (type === 'scatter') {
          const xs = this.points(this.read((s.xVal || {}).ref)).map((v) => this.num(v));
          const mk = s.marker || {};
          traces.push({
            type: 'scatter', name, x: xs, y: vals, text: txt,
            mode: (s.line === null ? '' : 'lines') + (mk.symbol === 'none' ? '' : (s.line === null ? 'markers' : '+markers')) + (showVal ? '+text' : ''),
            line: { color, width: Math.max(1, pt((s.line && s.line.w) || 2.25) * 0.75) },
            marker: { color: hex(mk.fill) || color, size: Math.max(4, (mk.size || 5) * z * 1.3), symbol: mk.symbol === 'square' ? 'square' : mk.symbol === 'diamond' ? 'diamond' : mk.symbol === 'triangle' ? 'triangle-up' : 'circle' },
            textposition: 'top center', textfont: lblFont,
            hovertemplate: `${esc(name)}<br>%{x}: %{text}<extra></extra>`,
          });
          continue;
        }
        const cats = this.categories(this.read((s.cat || {}).ref), vals.length);
        if (!cats0.length) cats0.push(...cats);
        if (type === 'bar') {
          anyBar = true; nBarSeries++;
          horiz = plot.barDir === 'bar';
          if (plot.grouping === 'stacked') barmode = 'relative';
          else if (plot.grouping === 'percentStacked') { barmode = 'relative'; barnorm = 'percent'; }
          if (plot.gapWidth !== undefined) gap = +plot.gapWidth;
          const colors = vals.map((_, k) => { const dp = (s.dPt || []).find((p) => +p.idx === k); return hex(dp && dp.fill) || color; });
          traces.push({
            type: 'bar', name, orientation: horiz ? 'h' : 'v',
            x: horiz ? vals : cats, y: horiz ? cats : vals, text: showVal ? txt : undefined, customdata: txt,
            textposition: showVal ? (lbl.pos === 'inEnd' || lbl.pos === 'ctr' || lbl.pos === 'inBase' ? 'inside' : plot.grouping === 'stacked' ? 'inside' : 'outside') : 'none',
            textfont: lblFont, cliponaxis: false,
            marker: { color: colors, line: { color: hex(s.line && s.line.c) || 'rgba(0,0,0,0)', width: s.line && s.line.c ? 1 : 0 } },
            hovertemplate: `${esc(name)}<br>%{${horiz ? 'y' : 'x'}}: %{customdata}<extra></extra>`,
          });
          continue;
        }
        if (type === 'line' || type === 'area') {
          const mk = s.marker || {};
          const area = type === 'area';
          const stacked = plot.grouping === 'stacked' || plot.grouping === 'percentStacked';
          const t = {
            type: 'scatter', name, x: cats, y: vals, text: txt, customdata: txt,
            mode: area ? 'lines' : 'lines' + (mk.symbol === 'none' ? '' : '+markers') + (showVal ? '+text' : ''),
            line: { color: area ? (hex(s.line && s.line.c) || color) : color, width: area ? 1 : Math.max(1, pt((s.line && s.line.w) || 2.25) * 0.75), shape: s.smooth ? 'spline' : 'linear' },
            marker: { color: hex(mk.fill) || color, size: Math.max(4, (mk.size || 5) * z * 1.3), symbol: mk.symbol === 'square' ? 'square' : mk.symbol === 'diamond' ? 'diamond' : mk.symbol === 'triangle' ? 'triangle-up' : 'circle' },
            textposition: 'top center', textfont: lblFont, connectgaps: sp.dispBlanksAs === 'span',
            hovertemplate: `${esc(name)}<br>%{x}: %{customdata}<extra></extra>`,
          };
          if (area) {
            t.fillcolor = hex(s.fill) || color;
            if (stacked) { t.stackgroup = 'a' + sp.plots.indexOf(plot); if (plot.grouping === 'percentStacked') t.groupnorm = 'percent'; }
            else t.fill = 'tozeroy';
          }
          traces.push(t);
        }
      }
    }
    // axes
    const axes = Object.values(sp.axes || {});
    const catAx = axes.find((a) => a.type === 'catAx' || a.type === 'dateAx') || (axes.find((a) => a.axPos === 'b'));
    const valAx = axes.find((a) => a.type === 'valAx' && a !== catAx) || axes.find((a) => a.type === 'valAx');
    const axisStyle = (a, isVal) => {
      const o = {
        showgrid: !!(a && a.grid) || (isVal && !isPie), gridcolor: '#FFFFFF', gridwidth: Math.max(1, 1.2 * z),
        zeroline: false, showline: false, ticks: 'outside', tickcolor: '#333333', ticklen: 3 * z,
        tickfont: { family: FONT, size: pt((a && a.font && a.font.sz) || 9), color: hex(a && a.font && a.font.color) || '#4D4D4D' },
        automargin: true, visible: !(a && a.delete === '1'),
      };
      if (a && a.font && a.font.b) o.tickfont.weight = 'bold';
      if (isVal && a) {
        const f = d3fmt(a.fmt);
        if (f) o.tickformat = f;
        if (a.min !== undefined) o.range = [+a.min, a.max !== undefined ? +a.max : null];
        if (a.orientation === 'maxMin') o.autorange = 'reversed';
      }
      if (!isVal) { o.type = 'category'; o.tickangle = 'auto'; if (a && a.orientation === 'maxMin') o.autorange = 'reversed'; }
      return o;
    };
    const title = this.titleText();
    const layout = {
      width: this.w, height: this.h,
      paper_bgcolor: hex(sp.chartSpace && sp.chartSpace.fill) || '#FFFFFF',
      plot_bgcolor: '#EBEBEB',
      font: { family: FONT, size: pt(10), color: '#4D4D4D' },
      margin: { l: 8 * z + 10, r: 10 * z + 6, t: (title ? pt((sp.titleFont && sp.titleFont.sz) || 14) * 2.2 : 12 * z + 6), b: 8 * z + 8, pad: 2 },
      showlegend: !!sp.legend,
      hoverlabel: { font: { family: FONT } },
    };
    if (title) {
      const tf = sp.titleFont || {};
      layout.title = { text: esc(title), x: 0.5, xanchor: 'center', font: { family: FONT, size: pt(tf.sz || 14), color: hex(tf.color) || '#333333', weight: tf.b ? 'bold' : undefined } };
    }
    if (sp.legend) {
      const lf = sp.legend.font || {};
      const pos = sp.legend.pos || 'r';
      layout.legend = { font: { family: FONT, size: pt(lf.sz || 9), color: hex(lf.color) || '#4D4D4D' }, bgcolor: 'rgba(0,0,0,0)' };
      if (lf.b) layout.legend.font.weight = 'bold';
      if (pos === 'b') Object.assign(layout.legend, { orientation: 'h', x: 0.5, xanchor: 'center', y: -0.12, yanchor: 'top' });
      else if (pos === 't') Object.assign(layout.legend, { orientation: 'h', x: 0.5, xanchor: 'center', y: 1.02, yanchor: 'bottom' });
      else if (pos === 'l') Object.assign(layout.legend, { x: -0.15, xanchor: 'right' });
    }
    if (!isPie) {
      const xa = axisStyle(horiz ? valAx : catAx, horiz), ya = axisStyle(horiz ? catAx : valAx, !horiz);
      if (sp.plots[0].type === 'scatter') { xa.type = 'linear'; delete xa.tickangle; xa.showgrid = true; }
      layout.xaxis = xa; layout.yaxis = ya;
      if (horiz && !(catAx && catAx.orientation === 'maxMin')) layout.yaxis.autorange = true;
    }
    if (anyBar) {
      layout.barmode = barmode;
      if (barnorm) layout.barnorm = barnorm;
      const n = barmode === 'group' ? Math.max(1, nBarSeries) : 1;
      layout.bargap = Math.min(0.9, (gap / 100) / (n + gap / 100));
      layout.bargroupgap = 0;
    }
    return { traces, layout };
  }

  titleText() {
    const sp = this.spec;
    if (sp.title) {
      if (sp.title.ref) {
        const d = this.read(sp.title.ref);
        if (d) { const v = d.get(d.rows[0], d.cols[0]); if (v !== null) return isErr(v) ? v.code : String(v); }
      }
      if (sp.title.text) return sp.title.text;
    }
    if (!sp.autoTitleDeleted && sp.plots.length === 1 && sp.plots[0].series.length === 1 && sp.title !== undefined) {
      return this.seriesName(sp.plots[0].series[0], 0);
    }
    return '';
  }

  render(force) {
    if (!window.Plotly || !this.w || !this.h) return;
    let fig;
    try { fig = this.figure(); } catch (e) { console.error('chart', this.id, e); return; }
    const sig = JSON.stringify(fig);
    if (!force && sig === this.sig) return;
    this.sig = sig;
    const border = this.spec.chartSpace && this.spec.chartSpace.line && this.spec.chartSpace.line.c;
    this.el.style.border = border ? `1px solid ${hex(border.c ? border : border)}` : 'none';
    window.Plotly.react(this.el, fig.traces, fig.layout, {
      displaylogo: false, responsive: false, displayModeBar: 'hover',
      modeBarButtonsToRemove: ['select2d', 'lasso2d', 'autoScale2d'],
      toImageButtonOptions: { filename: this.id, scale: 2 },
    });
  }
}

function esc(s) { return String(s).replace(/[<>&]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c])); }
