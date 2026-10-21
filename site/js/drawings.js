// Shapes, pictures, groups, charts and form controls placed over the grid.
// Shapes are positioned from their cell anchors, so hiding rows moves/collapses them as in Excel.
import { fontFamily, formatValue } from './format.js';
import { Ref, Matrix, isErr } from './engine/values.js';
import { ChartView } from './charts.js';

const css = (c) => (!c ? null : typeof c === 'string' ? '#' + c : c.c ? (c.a !== undefined && c.a < 1 ? rgba(c.c, c.a) : '#' + c.c) : null);
function rgba(hex, a) {
  const n = parseInt(hex, 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
}

export class Drawings {
  constructor(app) {
    this.app = app;
    this.items = [];
    this.charts = [];
    this.observer = new IntersectionObserver((ents) => {
      for (const e of ents) if (e.isIntersecting && e.target._chart) e.target._chart.ensure();
    }, { rootMargin: '200px' });
  }

  // Build DOM for every drawing of the sheet into the grid panes.
  build(model, grid) {
    this.model = model;
    this.grid = grid;
    for (const ch of this.charts) ch.destroy();
    this.observer.disconnect();
    grid.innerBody.innerHTML = '';
    grid.innerFrozen.innerHTML = '';
    this.items = [];
    this.charts = [];
    const d = model.data;
    (d.drawing || []).forEach((dr, i) => {
      const el = this.makeObj(dr.obj, true);
      el.classList.add('drawing');
      el.style.zIndex = String(10 + i);
      const frozen = model.freezeRows && dr.from && dr.from[2] < model.freezeRows;
      (frozen ? grid.innerFrozen : grid.innerBody).appendChild(el);
      this.items.push({ dr, el, obj: dr.obj });
    });
    (d.controls || []).forEach((ct, i) => {
      if (!ct.from) return;
      const el = document.createElement('label');
      el.className = 'drawing form-control';
      el.style.zIndex = String(1000 + i);
      const inp = document.createElement('input');
      inp.type = ct.objectType === 'Radio' ? 'radio' : 'checkbox';
      inp.name = 'ctl-' + model.si;
      el.appendChild(inp);
      inp.addEventListener('change', () => this.app.controlChanged(model.si, ct, d.controls));
      const frozen = model.freezeRows && ct.from[2] < model.freezeRows;
      (frozen ? grid.innerFrozen : grid.innerBody).appendChild(el);
      this.items.push({ dr: { anchor: 'twoCellAnchor', from: ct.from, to: ct.to }, el, ctl: ct, input: inp });
    });
    this.position();
    this.refresh();
  }

  makeObj(obj, top) {
    let el;
    if (obj.kind === 'grpSp') {
      el = document.createElement('div');
      el.className = 'shape-group';
      el._children = obj.children.map((ch) => { const c = this.makeObj(ch, false); el.appendChild(c); return { obj: ch, el: c }; });
    } else if (obj.kind === 'pic') {
      el = document.createElement('div');
      el.className = 'pic';
      const img = document.createElement('img');
      img.src = 'data/' + obj.img;
      img.alt = obj.descr || obj.name || '';
      img.draggable = false;
      el.appendChild(img);
    } else if (obj.kind === 'graphicFrame') {
      el = document.createElement('div');
      el.className = 'chart';
      if (obj.chart && this.app.wb.charts[obj.chart]) {
        const cv = new ChartView(this.app, this.model.si, obj.chart, el);
        el._chart = cv;
        this.charts.push(cv);
        this.observer.observe(el);
      }
    } else {
      el = document.createElement('div');
      el.className = 'shape';
      const tx = document.createElement('div');
      tx.className = 'shape-text';
      el.appendChild(tx);
      el._text = tx;
    }
    el._obj = obj;
    const clickable = obj.macro || obj.href;
    if (clickable || obj.kind === 'sp' || obj.kind === 'pic' || obj.kind === 'grpSp') {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        this.click(obj, top ? obj.name : null);
      });
    }
    if (obj.name) el.dataset.name = obj.name;
    return el;
  }

  click(obj, topName) {
    const si = this.model.si;
    const ov = this.app.shapeOverride(si, obj.name, false);
    const macro = ov && ov.onAction !== undefined ? ov.onAction : obj.macro;
    if (macro) { this.app.runMacro(macro, { sheet: si, caller: obj.name }); return; }
    if (obj.href) { this.app.followLink(obj.href, si); return; }
  }

  // Place every drawing from its anchor using the current layout.
  position() {
    const m = this.model;
    for (const it of this.items) {
      const { dr, el } = it;
      let x, y, w, h;
      const box = it.obj && it.obj.box;
      if (dr.anchor === 'absoluteAnchor' && box) {
        x = box[0] / 9525 * m.zoom; y = box[1] / 9525 * m.zoom; w = box[2] / 9525 * m.zoom; h = box[3] / 9525 * m.zoom;
      } else {
        const f = dr.from;
        x = m.anchorX(f[0], f[1]); y = m.anchorY(f[2], f[3]);
        if (dr.anchor === 'oneCellAnchor' || dr.editAs === 'oneCell' || dr.editAs === 'absolute') {
          // "move but don't size with cells": keeps its size and stays visible when rows hide
          const ext = dr.ext || (box && box[2] ? [box[2], box[3]] : null);
          if (ext) { w = ext[0] / 9525 * m.zoom; h = ext[1] / 9525 * m.zoom; }
          else { w = m.anchorX(dr.to[0], dr.to[1]) - x; h = m.anchorY(dr.to[2], dr.to[3]) - y; }
          if (dr.editAs === 'absolute' && box) { x = box[0] / 9525 * m.zoom; y = box[1] / 9525 * m.zoom; }
        } else {
          const t = dr.to;
          w = m.anchorX(t[0], t[1]) - x; h = m.anchorY(t[2], t[3]) - y;
        }
      }
      const hiddenObj = it.obj && this.isHidden(it.obj);
      if (!(w > 0.5 && h > 0.5) || hiddenObj) { el.style.display = 'none'; continue; }
      el.style.display = '';
      Object.assign(el.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: h + 'px' });
      if (it.obj) this.styleObj(it.obj, el, w, h, null);
      else if (it.ctl) {
        const s = Math.max(10, Math.min(w, h) * 0.75);
        it.input.style.width = s + 'px'; it.input.style.height = s + 'px';
      }
      if (el._chart) el._chart.resize(w, h);
    }
  }

  isHidden(obj) {
    const ov = this.app.shapeOverride(this.model.si, obj.name, false);
    if (ov && ov.visible !== undefined) return !ov.visible;
    return !!obj.hidden;
  }

  // Apply fill/line/text style with any macro overrides.
  styleObj(obj, el, w, h, groupOv) {
    const m = this.model, z = m.zoom;
    const ov = this.app.shapeOverride(m.si, obj.name, false) || groupOv || {};
    if (obj.kind === 'grpSp') {
      const [, , cw, chh] = obj.chbox || obj.box;
      const [cx0, cy0] = obj.chbox || obj.box;
      const sx = cw ? w / cw : 1, sy = chh ? h / chh : 1;
      for (const ch of el._children) {
        const b = ch.obj.box || [0, 0, 0, 0];
        const cwid = b[2] * sx, chgt = b[3] * sy;
        if (this.isHidden(ch.obj)) { ch.el.style.display = 'none'; continue; }
        Object.assign(ch.el.style, { display: '', position: 'absolute', left: ((b[0] - cx0) * sx) + 'px', top: ((b[1] - cy0) * sy) + 'px', width: cwid + 'px', height: chgt + 'px' });
        this.styleObj(ch.obj, ch.el, cwid, chgt, this.app.shapeOverride(m.si, obj.name, false));
        if (ch.el._chart) ch.el._chart.resize(cwid, chgt);
      }
      if (obj.rot) el.style.transform = `rotate(${obj.rot}deg)`;
      return;
    }
    const clickable = (ov.onAction !== undefined ? ov.onAction : obj.macro) || obj.href;
    el.style.cursor = clickable ? 'pointer' : 'default';
    el.classList.toggle('clickable', !!clickable);
    if (obj.kind === 'pic') {
      const img = el.firstElementChild;
      if (obj.crop) {
        const [l, t, r, b] = obj.crop;
        const iw = w / Math.max(0.01, 1 - l - r), ih = h / Math.max(0.01, 1 - t - b);
        Object.assign(img.style, { position: 'absolute', width: iw + 'px', height: ih + 'px', left: (-l * iw) + 'px', top: (-t * ih) + 'px' });
      } else Object.assign(img.style, { position: 'absolute', left: 0, top: 0, width: '100%', height: '100%' });
      if (obj.line && obj.line.c) el.style.outline = `${Math.max(1, obj.line.w * z)}px solid ${css(obj.line.c)}`;
      if (obj.rot) el.style.transform = `rotate(${obj.rot}deg)`;
      return;
    }
    if (obj.kind === 'graphicFrame') return;
    // ---- shape
    let bg = 'transparent';
    if (ov.fillHidden) bg = 'transparent';
    else if (ov.fill) bg = '#' + ov.fill;
    else if (obj.grad) {
      const stops = obj.grad.stops.map(([p, c]) => `${css(c)} ${p}%`).join(', ');
      bg = `linear-gradient(${(obj.grad.ang || 0) + 90}deg, ${stops})`;
    } else if (obj.fill) bg = css(obj.fill) || 'transparent';
    el.style.background = bg;
    let border = 'none';
    const line = obj.line;
    const lineVisible = ov.lineVisible !== undefined ? ov.lineVisible : !!(line && line.c);
    if (lineVisible) {
      const lw = Math.max(0.75, ((line && line.w) || 1) * z);
      const lc = ov.line ? '#' + ov.line : css(line && line.c) || '#000';
      border = `${lw}px ${line && line.dash && line.dash !== 'solid' ? 'dashed' : 'solid'} ${lc}`;
    }
    el.style.border = border;
    el.style.boxSizing = 'border-box';
    const geom = obj.geom;
    el.style.borderRadius = geom === 'ellipse' ? '50%' : geom === 'roundRect' ? (Math.min(w, h) * 0.1667) + 'px' : '0';
    if (obj.rot) el.style.transform = `rotate(${obj.rot}deg)`;
    // text
    const tx = el._text;
    const t = obj.text;
    tx.innerHTML = '';
    if (t && (t.hasText || obj.textlink)) {
      const ins = t.ins || [9.6, 4.8, 9.6, 4.8];
      const k = 4 / 3 * z;
      el.style.overflow = t.clip ? 'hidden' : 'visible';
      Object.assign(tx.style, {
        paddingLeft: ins[0] * z + 'px', paddingTop: ins[1] * z + 'px', paddingRight: ins[2] * z + 'px', paddingBottom: ins[3] * z + 'px',
        justifyContent: t.anchor === 'ctr' ? 'center' : t.anchor === 'b' ? 'flex-end' : 'flex-start',
        whiteSpace: t.nowrap ? 'pre' : 'pre-wrap',
        overflow: t.clip ? 'hidden' : 'visible',
        visibility: ov.textHidden ? 'hidden' : 'visible',
      });
      if (obj.vert || obj.rot === 270) { /* rotated by the container */ }
      let paras = t.paras;
      if (obj.textlink) {
        const val = this.linkedText(obj.textlink);
        const r0 = (paras[0] && paras[0].runs[0]) || {};
        paras = [{ runs: [{ ...r0, t: val }], align: paras[0] ? paras[0].align : 'l', sz: paras[0] && paras[0].sz }];
      }
      for (const p of paras) {
        const pd = document.createElement('div');
        pd.className = 'shape-para';
        pd.style.textAlign = { ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[p.align] || 'left';
        if (!p.runs.length) {
          pd.innerHTML = '&nbsp;';
          pd.style.fontSize = ((p.sz || 11) * k) + 'px';
        }
        for (const r of p.runs) {
          const sp = document.createElement('span');
          sp.textContent = r.t;
          const sz = r.sz || p.sz || 11;
          sp.style.fontSize = (sz * k) + 'px';
          sp.style.fontFamily = fontFamily(r.font || 'Calibri');
          if (r.b) sp.style.fontWeight = 'bold';
          if (r.i) sp.style.fontStyle = 'italic';
          if (r.u) sp.style.textDecoration = 'underline';
          sp.style.color = ov.textColor ? '#' + ov.textColor : css(r.color) || css(obj.fontColor) || '#000';
          pd.appendChild(sp);
        }
        tx.appendChild(pd);
      }
    }
  }

  linkedText(ref) {
    const app = this.app;
    let v = app.eng.evaluate(ref.replace(/^=/, ''), this.model.si);
    let fmt = 'General';
    if (v instanceof Ref) {
      const si = v.sheet, r = v.r1, c = v.c1;
      fmt = (app.models[si] ? app.models[si].xf(r, c).fmt : app.modelFor(si).xf(r, c).fmt) || 'General';
      v = v.value();
    }
    if (v instanceof Matrix) v = v.get(0, 0);
    if (isErr(v)) return v.code;
    return formatValue(v, fmt);
  }

  // Refresh dynamic content after a recalculation: linked text, charts, control state.
  refresh() {
    if (!this.model) return;
    for (const it of this.items) {
      if (it.ctl) {
        const ct = it.ctl;
        const ctls = this.model.data.controls;
        const link = ctls.find((c) => c.fmlaLink) || ct;
        if (ct.objectType === 'Radio' && link.fmlaLink) {
          const v = this.app.eng.evaluate(link.fmlaLink, this.model.si);
          const val = v instanceof Ref ? v.value() : v;
          const radios = ctls.filter((c) => c.objectType === 'Radio');
          it.input.checked = typeof val === 'number' && radios[val - 1] === ct;
        } else if (ct.fmlaLink) {
          const v = this.app.eng.evaluate(ct.fmlaLink, this.model.si);
          it.input.checked = (v instanceof Ref ? v.value() : v) === true;
        }
        continue;
      }
      if (it.obj && hasTextLink(it.obj) && it.el.style.display !== 'none') {
        this.styleObj(it.obj, it.el, parseFloat(it.el.style.width), parseFloat(it.el.style.height), null);
      }
    }
    for (const ch of this.charts) ch.refresh();
  }
}

function hasTextLink(o) {
  if (o.textlink) return true;
  return !!(o.children && o.children.some(hasTextLink));
}
