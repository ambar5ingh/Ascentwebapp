// Canvas sheet renderer with frozen rows, merged cells, text overflow, conditional formats,
// selection, editing, validation dropdowns, comments and hyperlinks.
import { formatValue, formatColor, fontFamily, colName } from './format.js';
import { isErr, Rich, Ref, Matrix, toText, isBlank } from './engine/values.js';
import { KEY } from './engine/engine.js';

const GRID_COLOR = '#D9D9D9';
const SEL_COLOR = '#217346';
const HEAD_BG = '#F3F3F3';

export class Grid {
  constructor(app, root) {
    this.app = app;
    this.root = root;
    root.innerHTML = `
      <div class="grid-scroll" tabindex="0"><div class="grid-spacer"></div></div>
      <canvas class="grid-canvas"></canvas>
      <div class="pane pane-body"><div class="pane-inner"></div></div>
      <div class="pane pane-frozen"><div class="pane-inner"></div></div>
      <div class="grid-overlay">
        <textarea class="cell-editor" spellcheck="false" style="display:none"></textarea>
        <button class="dv-button" style="display:none" tabindex="-1" aria-label="Show list">&#9662;</button>
        <div class="dv-list" style="display:none" role="listbox"></div>
        <div class="tip" style="display:none"></div>
        <div class="comment-tip" style="display:none"></div>
      </div>`;
    this.scroll = root.querySelector('.grid-scroll');
    this.spacer = root.querySelector('.grid-spacer');
    this.canvas = root.querySelector('.grid-canvas');
    this.ctx = this.canvas.getContext('2d');
    this.paneBody = root.querySelector('.pane-body');
    this.paneFrozen = root.querySelector('.pane-frozen');
    this.innerBody = this.paneBody.firstElementChild;
    this.innerFrozen = this.paneFrozen.firstElementChild;
    this.editor = root.querySelector('.cell-editor');
    this.dvButton = root.querySelector('.dv-button');
    this.dvList = root.querySelector('.dv-list');
    this.tip = root.querySelector('.tip');
    this.commentTip = root.querySelector('.comment-tip');
    this.model = null;
    this.sel = [1, 1];
    this.editing = null;
    this.textCache = new Map();
    this.measureCache = new Map();
    this.cfCache = new Map();
    this.cfEpoch = -1;
    this.pending = false;
    this.bind();
  }

  // ------------------------------------------------------------------ setup
  bind() {
    const s = this.scroll;
    s.addEventListener('scroll', () => { this.onScroll(); });
    new ResizeObserver(() => this.resize()).observe(this.root);
    s.addEventListener('mousedown', (e) => this.onMouseDown(e));
    s.addEventListener('dblclick', (e) => this.onDblClick(e));
    s.addEventListener('mousemove', (e) => this.onMouseMove(e));
    s.addEventListener('mouseleave', () => { this.commentTip.style.display = 'none'; });
    s.addEventListener('keydown', (e) => this.onKey(e));
    // wheel over shapes/charts scrolls the sheet
    for (const p of [this.paneBody, this.paneFrozen]) {
      p.addEventListener('wheel', (e) => {
        if (e.ctrlKey) return;
        e.preventDefault();
        s.scrollBy(e.deltaX, e.deltaY);
      }, { passive: false });
    }
    this.editor.addEventListener('keydown', (e) => this.onEditorKey(e));
    this.editor.addEventListener('blur', () => { if (this.editing && !this.committing) this.commitEdit(0, 0); });
    this.dvButton.addEventListener('mousedown', (e) => { e.preventDefault(); this.toggleList(); });
    this.dvList.addEventListener('mousedown', (e) => {
      e.preventDefault();
      const it = e.target.closest('.dv-item');
      if (it) this.pickListItem(+it.dataset.i);
    });
    document.addEventListener('mousedown', (e) => {
      if (this.dvList.style.display !== 'none' && !this.dvList.contains(e.target) && e.target !== this.dvButton) this.dvList.style.display = 'none';
    });
  }

  get headings() { return this.app.view.headings && true; }
  get gx() { return this.headings ? Math.round(36 * Math.max(0.7, this.model.zoom)) : 0; }
  get gy() { return this.headings ? Math.round(20 * Math.max(0.7, this.model.zoom)) : 0; }

  show(model, view) {
    this.cancelEdit();
    this.dvList.style.display = 'none';
    this.model = model;
    this.view = view; // {topRow, leftCol, sel}
    this.sel = view.sel || [1, 1];
    this.textCache.clear();
    this.cfCache.clear();
    this.relayout(true);
  }

  // Recompute sizes (zoom or hidden rows changed); keeps the top row/left column.
  relayout(restore) {
    const m = this.model;
    if (!m) return;
    const v = this.view;
    if (!restore) this.captureScroll();
    m.layout(this.app.zoomOf(m.si));
    this.measureCache.clear();
    this.spacer.style.width = (m.totalW + this.gx + 40) + 'px';
    this.spacer.style.height = (m.totalH + this.gy + 40) + 'px';
    this.app.drawings.build(m, this);
    this.applyScroll(v.topRow || 1, v.leftCol || 1);
    this.resize();
  }

  captureScroll() {
    const m = this.model, v = this.view;
    if (!m || !m.rowY) return;
    const st = this.scroll.scrollTop, sl = this.scroll.scrollLeft;
    v.topRow = m.rowAtY(m.frozenH + st + 0.5);
    v.leftCol = m.colAtX(sl + 0.5);
  }
  // scroll so that topRow is the first scrolled row (Excel ActiveWindow.ScrollRow)
  applyScroll(topRow, leftCol) {
    const m = this.model;
    let r = Math.max(topRow, m.freezeRows + 1);
    while (r < m.nRows && m.rowH[r] === 0) r++;
    const st = Math.max(0, m.rowY[Math.min(r, m.nRows)] - m.frozenH);
    const c = Math.max(1, Math.min(leftCol, m.nCols));
    const sl = m.colX[c];
    this.scroll.scrollTop = st;
    this.scroll.scrollLeft = sl;
    this.view.topRow = r; this.view.leftCol = c;
    this.onScroll();
  }

  resize() {
    const s = this.scroll;
    const W = s.clientWidth, H = s.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    this.W = W; this.H = H;
    this.canvas.style.width = W + 'px'; this.canvas.style.height = H + 'px';
    this.canvas.width = Math.round(W * dpr); this.canvas.height = Math.round(H * dpr);
    this.dpr = dpr;
    this.layoutPanes();
    this.invalidate();
  }

  layoutPanes() {
    const m = this.model;
    if (!m) return;
    const gx = this.gx, gy = this.gy, fh = m.frozenH;
    const sl = this.scroll.scrollLeft, st = this.scroll.scrollTop;
    Object.assign(this.paneFrozen.style, { left: gx + 'px', top: gy + 'px', width: Math.max(0, this.W - gx) + 'px', height: fh + 'px' });
    Object.assign(this.paneBody.style, { left: gx + 'px', top: (gy + fh) + 'px', width: Math.max(0, this.W - gx) + 'px', height: Math.max(0, this.H - gy - fh) + 'px' });
    this.innerFrozen.style.transform = `translate(${-sl}px, 0px)`;
    this.innerBody.style.transform = `translate(${-sl}px, ${-st - fh}px)`;
  }

  onScroll() {
    this.layoutPanes();
    this.positionOverlay();
    this.invalidate();
    if (this.model) {
      clearTimeout(this.scrollSave);
      this.scrollSave = setTimeout(() => { this.captureScroll(); this.app.saveSoon(); }, 300);
    }
  }

  invalidate() {
    if (this.pending) return;
    this.pending = true;
    requestAnimationFrame(() => { this.pending = false; this.draw(); });
  }

  // ------------------------------------------------------------------ geometry
  // screen rect of a cell (or its merge) relative to the grid root
  cellRect(r, c) {
    const m = this.model;
    let r1 = r, c1 = c, r2 = r, c2 = c;
    const mg = m.mergeAt(r, c);
    if (mg) [r1, c1, r2, c2] = mg;
    const x = this.gx + m.colX[c1] - this.scroll.scrollLeft;
    const y = this.gy + m.rowY[r1] - (r1 <= m.freezeRows ? 0 : this.scroll.scrollTop);
    const w = m.colX[Math.min(c2, m.nCols) + 1] - m.colX[c1];
    const h = m.rowY[Math.min(r2, m.nRows) + 1] - m.rowY[r1];
    return { x, y, w, h, r1, c1, r2, c2 };
  }
  hit(clientX, clientY) {
    const m = this.model;
    const b = this.scroll.getBoundingClientRect();
    const x = clientX - b.left - this.gx, y = clientY - b.top - this.gy;
    if (x < 0 || y < 0) return null;
    const sx = x + this.scroll.scrollLeft;
    const sy = y < m.frozenH ? y : y + this.scroll.scrollTop;
    if (sx > m.totalW || sy > m.totalH) return null;
    let r = m.rowAtY(sy), c = m.colAtX(sx);
    if (y >= m.frozenH && r <= m.freezeRows) r = m.freezeRows + 1;
    return [r, c];
  }
  isCellVisible(r, c) {
    const rc = this.cellRect(r, c);
    const m = this.model;
    const top = r <= m.freezeRows ? this.gy : this.gy + m.frozenH;
    return rc.y >= top && rc.y + rc.h <= this.H && rc.x >= this.gx && rc.x + Math.min(rc.w, 50) <= this.W;
  }
  ensureVisible(r, c) {
    const m = this.model;
    const mg = m.mergeAt(r, c);
    if (mg) { r = mg[0]; c = mg[1]; }
    const s = this.scroll;
    const viewH = this.H - this.gy - m.frozenH, viewW = this.W - this.gx;
    if (r > m.freezeRows) {
      const top = m.rowY[r] - m.frozenH, bot = m.rowY[r] + m.rowH[r] - m.frozenH;
      if (top < s.scrollTop) s.scrollTop = top;
      else if (bot > s.scrollTop + viewH) s.scrollTop = Math.min(top, bot - viewH);
    }
    const left = m.colX[c], right = m.colX[c] + m.colW[c];
    if (left < s.scrollLeft) s.scrollLeft = left;
    else if (right > s.scrollLeft + viewW) s.scrollLeft = Math.min(left, right - viewW);
  }

  // ------------------------------------------------------------------ drawing
  draw() {
    const m = this.model;
    if (!m || !m.rowY) return;
    const ctx = this.ctx, W = this.W, H = this.H;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, W, H);
    if (this.app.eng.epoch !== this.cfEpoch) { this.cfCache.clear(); this.cfEpoch = this.app.eng.epoch; }
    const gx = this.gx, gy = this.gy, fh = m.frozenH;
    const sl = this.scroll.scrollLeft, st = this.scroll.scrollTop;
    const c1 = m.colAtX(sl), c2 = m.colAtX(sl + W - gx);
    if (H - gy - fh > 0) {
      const r1 = Math.max(m.freezeRows + 1, m.rowAtY(fh + st)), r2 = m.rowAtY(st + H - gy);
      this.drawRegion(r1, r2, c1, c2, gx - sl, gy - st, [gx, gy + fh, W - gx, H - gy - fh]);
    }
    if (m.freezeRows && fh > 0) {
      this.drawRegion(1, m.freezeRows, c1, c2, gx - sl, gy, [gx, gy, W - gx, fh]);
      ctx.fillStyle = '#BFBFBF';
      ctx.fillRect(gx, gy + fh - 0.5, W - gx, 1);
    }
    if (this.headings) this.drawHeadings(c1, c2);
    this.drawSelection();
  }

  drawRegion(r1, r2, c1, c2, ox, oy, clip) {
    const m = this.model, ctx = this.ctx, app = this.app;
    const sh = app.eng.sheets[m.si];
    ctx.save();
    ctx.beginPath();
    ctx.rect(clip[0], clip[1], clip[2], clip[3]);
    ctx.clip();
    const rows = [], cols = [];
    for (let r = r1; r <= Math.min(r2, m.nRows); r++) if (m.rowH[r] > 0) rows.push(r);
    for (let c = c1; c <= Math.min(c2, m.nCols); c++) if (m.colW[c] > 0) cols.push(c);
    const X = (c) => ox + m.colX[c], Y = (r) => oy + m.rowY[r];
    // gridlines
    if (m.showGrid && app.view.gridlines !== false) {
      ctx.fillStyle = GRID_COLOR;
      const x0 = X(cols[0] || c1), x1 = X((cols[cols.length - 1] || c2)) + (m.colW[cols[cols.length - 1]] || 0);
      const y0 = Y(rows[0] || r1), y1 = Y((rows[rows.length - 1] || r2)) + (m.rowH[rows[rows.length - 1]] || 0);
      for (const r of rows) ctx.fillRect(x0, Math.round(Y(r) + m.rowH[r]) - 1, x1 - x0, 1);
      for (const c of cols) ctx.fillRect(Math.round(X(c) + m.colW[c]) - 1, y0, 1, y1 - y0);
    }
    // fills
    const merges = new Set();
    const cells = []; // [r, c, xf, cf]
    for (const r of rows) {
      for (const c of cols) {
        const mi = m.mergeIdx.get(KEY(r, c));
        if (mi !== undefined) { merges.add(mi); continue; }
        const xf = m.xf(r, c);
        const cf = this.cfAt(r, c);
        const fill = (cf && cf.fill) || xf.fill;
        if (fill) {
          ctx.fillStyle = '#' + fill;
          const x0 = Math.round(X(c)), y0 = Math.round(Y(r));
          ctx.fillRect(x0, y0, Math.round(X(c) + m.colW[c]) - x0, Math.round(Y(r) + m.rowH[r]) - y0);
        }
        cells.push(r, c);
      }
    }
    for (const mi of merges) {
      const [mr1, mc1, mr2, mc2] = m.merges[mi];
      const xf = m.xf(mr1, mc1);
      const cf = this.cfAt(mr1, mc1);
      const fill = (cf && cf.fill) || xf.fill;
      const x = Math.round(X(mc1)), y = Math.round(Y(mr1));
      const w = Math.round(X(mc1) + m.colX[Math.min(mc2, m.nCols) + 1] - m.colX[mc1]) - x, h = Math.round(Y(mr1) + m.rowY[Math.min(mr2, m.nRows) + 1] - m.rowY[mr1]) - y;
      if (w <= 0 || h <= 0) continue;
      ctx.fillStyle = fill ? '#' + fill : '#FFFFFF';
      ctx.fillRect(x, y, w - (fill ? 0 : 1), h - (fill ? 0 : 1));
    }
    // text: also look a few columns to the left for overflowing text
    const lead = [];
    for (let c = Math.max(1, c1 - 12); c < c1; c++) if (m.colW[c] > 0) lead.push(c);
    for (const r of rows) {
      for (const c of lead) {
        if (m.mergeIdx.has(KEY(r, c))) continue;
        const v = sh.get(r, c);
        if (typeof v === 'string' && v) this.drawCellText(r, c, X, Y, sh);
      }
    }
    for (let k = 0; k < cells.length; k += 2) this.drawCellText(cells[k], cells[k + 1], X, Y, sh);
    for (const mi of merges) this.drawMergeText(m.merges[mi], X, Y, sh);
    // borders
    for (const r of rows) for (const c of cols) this.drawBorders(r, c, X, Y);
    // comment indicators
    if (m.comments.size) {
      ctx.fillStyle = '#C00000';
      for (const r of rows) for (const c of cols) {
        if (!m.comments.has(KEY(r, c))) continue;
        const x = X(c) + m.colW[c], y = Y(r);
        ctx.beginPath(); ctx.moveTo(x - 6, y); ctx.lineTo(x, y); ctx.lineTo(x, y + 6); ctx.closePath(); ctx.fill();
      }
    }
    ctx.restore();
  }

  cfAt(r, c) {
    const m = this.model;
    const k = KEY(r, c);
    if (this.cfCache.has(k)) return this.cfCache.get(k);
    const rules = m.cfRulesAt(r, c);
    let out = null;
    if (rules) {
      const eng = this.app.eng;
      const v = eng.getValue(m.si, r, c);
      for (const ru of rules) {
        let hitRule = false;
        if (ru.type === 'containsText') {
          hitRule = ru.text !== undefined && toText(isErr(v) ? '' : v).toLowerCase().includes(String(ru.text).toLowerCase());
        } else if (ru.type === 'notContainsBlanks') {
          hitRule = !isBlank(v) && String(toText(v)).trim() !== '';
        } else if (ru.type === 'containsBlanks') {
          hitRule = isBlank(v) || String(toText(v)).trim() === '';
        } else if (ru.type === 'expression' && ru.f && ru.f[0]) {
          let res = eng.evaluateRel(ru.f[0], m.si, ru.r0, ru.c0, r, c);
          if (res instanceof Ref) res = res.value();
          if (res instanceof Matrix) res = res.get(0, 0);
          hitRule = res === true || (typeof res === 'number' && res !== 0);
        }
        if (hitRule) {
          const dx = this.app.wb.dxfs[+ru.dxfId];
          if (dx) {
            out = out || {};
            if (dx.fill && out.fill === undefined) out.fill = dx.fill;
            if (dx.font) {
              if (dx.font.color && out.color === undefined) out.color = dx.font.color;
              if (dx.font.b !== undefined && out.b === undefined) out.b = dx.font.b;
              if (dx.font.i !== undefined && out.i === undefined) out.i = dx.font.i;
            }
          }
          if (ru.stopIfTrue) break;
        }
      }
    }
    this.cfCache.set(k, out);
    return out;
  }

  fontFor(xf, cf) {
    const f = xf.font || this.app.wb.defaultFont || {};
    const z = this.model.zoom;
    const sz = (f.sz || 11) * 4 / 3 * z;
    const b = cf && cf.b !== undefined ? cf.b : f.b;
    const i = cf && cf.i !== undefined ? cf.i : f.i;
    return { css: `${i ? 'italic ' : ''}${b ? 'bold ' : ''}${sz.toFixed(2)}px ${fontFamily(f.name)}`, sz, u: f.u && f.u !== 'none', color: f.color };
  }

  displayText(r, c, v, fmt) {
    const k = KEY(r, c);
    const hit = this.textCache.get(k);
    if (hit && hit.v === v && hit.fmt === fmt) return hit.t;
    const t = formatValue(v, fmt);
    this.textCache.set(k, { v, fmt, t });
    return t;
  }
  measure(font, text) {
    const k = font + '\u0001' + text;
    let w = this.measureCache.get(k);
    if (w === undefined) {
      this.ctx.font = font;
      w = this.ctx.measureText(text).width;
      if (this.measureCache.size > 50000) this.measureCache.clear();
      this.measureCache.set(k, w);
    }
    return w;
  }

  isEmptyCell(r, c, sh) {
    if (this.model.mergeIdx.has(KEY(r, c))) return false;
    const v = sh.get(r, c);
    if (v !== null && v !== '') return false;
    return !sh.formulas.has(KEY(r, c)) || v === null;
  }

  drawCellText(r, c, X, Y, sh) {
    const v = sh.get(r, c);
    if (v === null || v === undefined || v === '') return;
    const m = this.model;
    const xf = m.xf(r, c);
    const cf = this.cfAt(r, c);
    const x = X(c), y = Y(r), w = m.colW[c], h = m.rowH[r];
    this.paintText(r, c, v, xf, cf, x, y, w, h, sh, true);
  }

  drawMergeText(mg, X, Y, sh) {
    const [r1, c1, r2, c2] = mg;
    const m = this.model;
    const v = sh.get(r1, c1);
    if (v === null || v === undefined || v === '') return;
    const xf = m.xf(r1, c1);
    const cf = this.cfAt(r1, c1);
    const x = X(c1), y = Y(r1);
    const w = m.colX[Math.min(c2, m.nCols) + 1] - m.colX[c1], h = m.rowY[Math.min(r2, m.nRows) + 1] - m.rowY[r1];
    if (w <= 0 || h <= 0) return;
    this.paintText(r1, c1, v, xf, cf, x, y, w, h, sh, false);
  }

  paintText(r, c, v, xf, cf, x, y, w, h, sh, canOverflow) {
    const m = this.model, ctx = this.ctx, z = m.zoom;
    const fmt = xf.fmt || 'General';
    let text = this.displayText(r, c, v, fmt);
    if (text === '') return;
    const font = this.fontFor(xf, cf);
    const al = xf.align || {};
    let h_ = al.h || 'general';
    if (h_ === 'general') h_ = typeof v === 'number' ? 'right' : (typeof v === 'boolean' || isErr(v)) ? 'center' : 'left';
    if (h_ === 'centerContinuous' || h_ === 'distributed') h_ = 'center';
    if (h_ === 'fill' || h_ === 'justify') h_ = 'left';
    const va = al.v || 'bottom';
    const pad = Math.max(1, 2.5 * z);
    const indent = (al.ind || 0) * 9 * z;
    const color = (cf && cf.color) || formatColor(v, fmt) || font.color || '000000';
    ctx.font = font.css;
    ctx.fillStyle = color[0] === '#' ? color : '#' + color;
    const lineH = font.sz * 1.22;
    if (al.rot === 90 || al.rot === 180 || al.rot === 255) {
      ctx.save();
      ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
      ctx.translate(x + w / 2, y + h / 2);
      ctx.rotate(-Math.PI / 2);
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(text, 0, 0);
      ctx.restore();
      return;
    }
    if (al.wrap || text.includes('\n')) {
      const lines = this.wrapLines(text, font.css, Math.max(4, w - 2 * pad - indent), !!al.wrap);
      const total = lines.length * lineH;
      let ty;
      if (va === 'top') ty = y + pad * 0.6;
      else if (va === 'center' || va === 'justify' || va === 'distributed') ty = y + (h - total) / 2;
      else ty = y + h - total - pad * 0.6;
      ctx.save();
      ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
      ctx.textBaseline = 'top';
      for (let i = 0; i < lines.length; i++) {
        const ln = lines[i];
        const tw = this.measure(font.css, ln);
        let tx;
        if (h_ === 'center') tx = x + (w - tw) / 2;
        else if (h_ === 'right') tx = x + w - pad - tw - indent;
        else tx = x + pad + indent;
        const yy = ty + i * lineH + (lineH - font.sz) / 2;
        ctx.fillText(ln, tx, yy);
        if (font.u) ctx.fillRect(tx, yy + font.sz * 1.02, tw, Math.max(1, z));
      }
      ctx.restore();
      return;
    }
    let tw = this.measure(font.css, text);
    let size = font.sz;
    if (al.shrink && tw > w - 2 * pad && tw > 0) {
      const k = (w - 2 * pad) / tw;
      size = font.sz * k;
      ctx.font = font.css.replace(/[\d.]+px/, size.toFixed(2) + 'px');
      tw *= k;
    }
    // overflow into empty neighbours (text only)
    let cx0 = x, cx1 = x + w;
    const avail = w - 2 * pad - indent;
    if (tw > avail) {
      if (canOverflow && typeof v === 'string') {
        const need = tw - avail;
        if (h_ === 'left' || h_ === 'center') {
          let ext = 0, cc = c + 1;
          const lim = h_ === 'center' ? need / 2 : need;
          while (ext < lim && cc <= m.nCols && this.isEmptyCell(r, cc, this.app.eng.sheets[m.si])) { ext += m.colW[cc]; cc++; }
          cx1 += ext;
        }
        if (h_ === 'right' || h_ === 'center') {
          let ext = 0, cc = c - 1;
          const lim = h_ === 'center' ? need / 2 : need;
          while (ext < lim && cc >= 1 && this.isEmptyCell(r, cc, this.app.eng.sheets[m.si])) { ext += m.colW[cc]; cc--; }
          cx0 -= ext;
        }
      } else if (typeof v === 'number' && tw > (w - pad) * 1.12 && !/E[+-]/.test(text)) {
        text = '#'.repeat(Math.max(1, Math.floor((w - 2 * pad) / this.measure(font.css, '#'))));
        tw = this.measure(font.css, text);
      }
    }
    let tx;
    if (h_ === 'center') tx = x + (w - tw) / 2;
    else if (h_ === 'right') tx = x + w - pad - tw - indent;
    else tx = x + pad + indent;
    let ty, base;
    if (va === 'top') { base = 'top'; ty = y + pad * 0.5; }
    else if (va === 'center' || va === 'justify' || va === 'distributed') { base = 'middle'; ty = y + h / 2; }
    else { base = 'alphabetic'; ty = y + h - Math.max(2, size * 0.24); }
    ctx.textBaseline = base;
    const clipNeeded = tx < cx0 || tx + tw > cx1 || size > h;
    if (clipNeeded) { ctx.save(); ctx.beginPath(); ctx.rect(cx0, y, cx1 - cx0, h); ctx.clip(); }
    ctx.fillText(text, tx, ty);
    if (font.u) {
      const uy = base === 'top' ? ty + size : base === 'middle' ? ty + size * 0.45 : ty + Math.max(1, size * 0.1);
      ctx.fillRect(tx, uy, tw, Math.max(1, z));
    }
    if (clipNeeded) ctx.restore();
  }

  wrapLines(text, font, width, wrap) {
    const out = [];
    for (const para of String(text).split(/\r?\n/)) {
      if (!wrap) { out.push(para); continue; }
      const words = para.split(/(\s+)/);
      let line = '';
      for (const wd of words) {
        if (!wd) continue;
        const t = line + wd;
        if (line && this.measure(font, t.trimEnd()) > width) {
          out.push(line.trimEnd());
          line = /^\s+$/.test(wd) ? '' : wd;
          // break very long words
          while (this.measure(font, line) > width && line.length > 1) {
            let n = line.length - 1;
            while (n > 1 && this.measure(font, line.slice(0, n)) > width) n--;
            out.push(line.slice(0, n)); line = line.slice(n);
          }
        } else line = t;
      }
      out.push(line.trimEnd());
    }
    return out;
  }

  drawBorders(r, c, X, Y) {
    const m = this.model, ctx = this.ctx;
    const xf = m.xf(r, c);
    const b = xf.border;
    if (!b) return;
    const mg = m.mergeAt(r, c);
    const x = X(c), y = Y(r), w = m.colW[c], h = m.rowH[r];
    const line = (spec, x0, y0, len, horiz) => {
      const [style, color] = spec;
      const lw = style === 'medium' || style.startsWith('medium') || style === 'slantDashDot' ? 2 : style === 'thick' ? 3 : 1;
      ctx.fillStyle = '#' + (color || '000000');
      const dashed = /dash|dot|hair/i.test(style);
      if (style === 'double') {
        if (horiz) { ctx.fillRect(x0, y0 - 1.5, len, 1); ctx.fillRect(x0, y0 + 0.5, len, 1); }
        else { ctx.fillRect(x0 - 1.5, y0, 1, len); ctx.fillRect(x0 + 0.5, y0, 1, len); }
        return;
      }
      const off = Math.floor(lw / 2);
      if (!dashed) {
        if (horiz) ctx.fillRect(x0, Math.round(y0) - off - (lw === 1 ? 1 : 0), len, lw);
        else ctx.fillRect(Math.round(x0) - off - (lw === 1 ? 1 : 0), y0, lw, len);
        return;
      }
      const seg = /dot|hair/i.test(style) ? 1 : 3;
      for (let p = 0; p < len; p += seg * 2) {
        const l = Math.min(seg, len - p);
        if (horiz) ctx.fillRect(x0 + p, Math.round(y0) - 1, l, lw); else ctx.fillRect(Math.round(x0) - 1, y0 + p, lw, l);
      }
    };
    if (b.t && (!mg || r === mg[0])) line(b.t, x, y + (r === 1 ? 1 : 0), w, true);
    if (b.b && (!mg || r === mg[2])) line(b.b, x, y + h, w, true);
    if (b.l && (!mg || c === mg[1])) line(b.l, x + (c === 1 ? 1 : 0), y, h, false);
    if (b.r && (!mg || c === mg[3])) line(b.r, x + w, y, h, false);
  }

  drawHeadings(c1, c2) {
    const m = this.model, ctx = this.ctx, gx = this.gx, gy = this.gy;
    const sl = this.scroll.scrollLeft, st = this.scroll.scrollTop;
    ctx.save();
    ctx.fillStyle = HEAD_BG;
    ctx.fillRect(0, 0, this.W, gy);
    ctx.fillRect(0, 0, gx, this.H);
    ctx.font = `${Math.round(11 * Math.max(0.75, m.zoom))}px ${fontFamily('Calibri')}`;
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    const [sr, sc] = this.sel;
    for (let c = c1; c <= Math.min(c2, m.nCols); c++) {
      const w = m.colW[c]; if (!w) continue;
      const x = gx + m.colX[c] - sl;
      if (c === sc) { ctx.fillStyle = '#D2D2D2'; ctx.fillRect(x, 0, w, gy); }
      ctx.fillStyle = '#C8C8C8'; ctx.fillRect(x + w - 1, 0, 1, gy);
      ctx.fillStyle = '#444'; ctx.fillText(colName(c), x + w / 2, gy / 2);
    }
    const drawRow = (r, y) => {
      const h = m.rowH[r]; if (!h) return;
      if (r === sr) { ctx.fillStyle = '#D2D2D2'; ctx.fillRect(0, y, gx, h); }
      ctx.fillStyle = '#C8C8C8'; ctx.fillRect(0, y + h - 1, gx, 1);
      if (h > 8) { ctx.fillStyle = '#444'; ctx.fillText(String(r), gx / 2, y + h / 2); }
    };
    ctx.save(); ctx.beginPath(); ctx.rect(0, gy + m.frozenH, gx, this.H); ctx.clip();
    const r1 = Math.max(m.freezeRows + 1, m.rowAtY(m.frozenH + st)), r2 = m.rowAtY(st + this.H - gy);
    for (let r = r1; r <= Math.min(r2, m.nRows); r++) drawRow(r, gy + m.rowY[r] - st);
    ctx.restore();
    for (let r = 1; r <= m.freezeRows; r++) drawRow(r, gy + m.rowY[r]);
    ctx.fillStyle = '#BDBDBD';
    ctx.fillRect(0, gy - 1, this.W, 1); ctx.fillRect(gx - 1, 0, 1, this.H);
    ctx.restore();
  }

  drawSelection() {
    const m = this.model, ctx = this.ctx;
    const [r, c] = this.sel;
    if (m.rowHidden(r) || m.colHidden(c)) return;
    const rc = this.cellRect(r, c);
    if (rc.w <= 0 || rc.h <= 0) return;
    const inFrozen = rc.r1 <= m.freezeRows;
    ctx.save();
    ctx.beginPath();
    if (inFrozen) ctx.rect(this.gx, this.gy, this.W, m.frozenH);
    else ctx.rect(this.gx, this.gy + m.frozenH, this.W, this.H);
    ctx.clip();
    ctx.strokeStyle = SEL_COLOR;
    ctx.lineWidth = 2;
    ctx.strokeRect(rc.x, rc.y, rc.w - 1, rc.h - 1);
    ctx.fillStyle = SEL_COLOR;
    ctx.fillRect(rc.x + rc.w - 4, rc.y + rc.h - 4, 5, 5);
    ctx.restore();
  }

  // ------------------------------------------------------------------ interaction
  select(r, c, opts = {}) {
    const m = this.model;
    r = Math.max(1, Math.min(r, m.nRows));
    c = Math.max(1, Math.min(c, m.nCols));
    const mg = m.mergeAt(r, c);
    if (mg) { r = mg[0]; c = mg[1]; }
    this.sel = [r, c];
    this.view.sel = this.sel;
    if (opts.scroll !== false) this.ensureVisible(r, c);
    this.dvList.style.display = 'none';
    this.positionOverlay();
    this.invalidate();
    this.app.onSelect(m.si, r, c);
    this.app.saveSoon();
  }

  onMouseDown(e) {
    if (e.button !== 0) return;
    const h = this.hit(e.clientX, e.clientY);
    if (!h) return;
    if (this.editing) this.commitEdit(0, 0);
    e.preventDefault();
    this.scroll.focus({ preventScroll: true });
    this.select(h[0], h[1], { scroll: false });
    const m = this.model;
    const link = m.links.get(KEY(h[0], h[1]));
    if (link) this.app.followLink(link.url || link.location || '', m.si);
  }
  onDblClick(e) {
    const h = this.hit(e.clientX, e.clientY);
    if (!h) return;
    this.beginEdit(false);
  }
  onMouseMove(e) {
    const h = this.hit(e.clientX, e.clientY);
    const m = this.model;
    if (!h) { this.scroll.style.cursor = ''; this.commentTip.style.display = 'none'; return; }
    let [r, c] = h;
    const mg = m.mergeAt(r, c);
    if (mg) { r = mg[0]; c = mg[1]; }
    this.scroll.style.cursor = m.links.has(KEY(r, c)) ? 'pointer' : 'cell';
    const cm = m.comments.get(KEY(r, c));
    if (cm) {
      const rc = this.cellRect(r, c);
      this.commentTip.textContent = cm;
      Object.assign(this.commentTip.style, { display: 'block', left: (rc.x + rc.w + 6) + 'px', top: rc.y + 'px' });
    } else this.commentTip.style.display = 'none';
  }

  nextVisible(r, c, dr, dc) {
    const m = this.model;
    const mg = m.mergeAt(r, c);
    if (mg) { if (dr > 0) r = mg[2]; if (dc > 0) c = mg[3]; }
    do { r += dr; c += dc; } while (r >= 1 && r <= m.nRows && c >= 1 && c <= m.nCols && ((dr && m.rowHidden(r)) || (dc && m.colHidden(c))));
    return [Math.max(1, Math.min(r, m.nRows)), Math.max(1, Math.min(c, m.nCols))];
  }

  onKey(e) {
    if (this.editing) return;
    const [r, c] = this.sel;
    const k = e.key;
    const mod = e.ctrlKey || e.metaKey;
    const mv = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
    if (mv && e.altKey && k === 'ArrowDown') { e.preventDefault(); this.toggleList(); return; }
    if (mv) { e.preventDefault(); const [nr, nc] = this.nextVisible(r, c, ...mv); this.select(nr, nc); return; }
    if (k === 'Enter') { e.preventDefault(); const [nr, nc] = this.nextVisible(r, c, e.shiftKey ? -1 : 1, 0); this.select(nr, nc); return; }
    if (k === 'Tab') { e.preventDefault(); const [nr, nc] = this.nextVisible(r, c, 0, e.shiftKey ? -1 : 1); this.select(nr, nc); return; }
    if (k === 'PageDown' || k === 'PageUp') {
      e.preventDefault();
      const d = k === 'PageDown' ? 1 : -1;
      this.scroll.scrollTop += d * (this.H - this.gy - this.model.frozenH) * 0.9;
      return;
    }
    if (k === 'F2') { e.preventDefault(); this.beginEdit(false); return; }
    if (k === 'Delete' || k === 'Backspace') {
      e.preventDefault();
      if (k === 'Backspace') { this.beginEdit(true, ''); return; }
      this.app.userEdit(this.model.si, r, c, '');
      return;
    }
    if (mod && (k === 'z' || k === 'Z')) { e.preventDefault(); if (e.shiftKey) this.app.redo(); else this.app.undo(); return; }
    if (mod && (k === 'y' || k === 'Y')) { e.preventDefault(); this.app.redo(); return; }
    if (mod && (k === 'c' || k === 'C')) { e.preventDefault(); this.copy(); return; }
    if (mod && (k === 'v' || k === 'V')) { return; }
    if (k.length === 1 && !mod && !e.altKey) { this.beginEdit(true, k); e.preventDefault(); }
  }

  copy() {
    const [r, c] = this.sel;
    const v = this.app.eng.getValue(this.model.si, r, c);
    const xf = this.model.xf(r, c);
    const t = v === null ? '' : formatValue(v, xf.fmt || 'General').trim();
    if (navigator.clipboard) navigator.clipboard.writeText(t).catch(() => {});
    this.app.toast('Copied ' + colName(c) + r);
  }

  // ------------------------------------------------------------------ editing
  beginEdit(replace, initial) {
    const m = this.model;
    const [r, c] = this.sel;
    const why = this.app.cannotEdit(m.si, r, c);
    if (why) { this.app.alert(why.title || 'Microsoft Excel', why.msg); return; }
    const rc = this.cellRect(r, c);
    const v = this.app.eng.getValue(m.si, r, c);
    const xf = m.xf(r, c);
    let text = '';
    if (!replace) text = this.app.editText(v, xf);
    else if (initial !== undefined) text = initial;
    this.editing = { r, c };
    const ed = this.editor;
    const font = this.fontFor(xf, null);
    ed.value = text;
    Object.assign(ed.style, {
      display: 'block', font: font.css, left: rc.x + 'px', top: rc.y + 'px',
      minWidth: Math.max(rc.w, 60) + 'px', width: Math.max(rc.w, 60) + 'px', height: Math.max(rc.h, font.sz * 1.4) + 'px',
      background: xf.fill ? '#' + xf.fill : '#fff', color: '#' + ((xf.font && xf.font.color) || '000000'),
    });
    ed.focus();
    ed.setSelectionRange(ed.value.length, ed.value.length);
    this.positionDvButton();
  }
  cancelEdit() {
    if (!this.editing) return;
    this.editing = null;
    this.editor.style.display = 'none';
    this.scroll.focus({ preventScroll: true });
  }
  async commitEdit(dr, dc) {
    if (!this.editing) return;
    const { r, c } = this.editing;
    const text = this.editor.value;
    this.committing = true;
    this.editing = null;
    this.editor.style.display = 'none';
    this.committing = false;
    this.scroll.focus({ preventScroll: true });
    const ok = await this.app.userEdit(this.model.si, r, c, text);
    if (ok && (dr || dc)) { const [nr, nc] = this.nextVisible(r, c, dr, dc); this.select(nr, nc); }
  }
  onEditorKey(e) {
    if (e.key === 'Escape') { e.preventDefault(); this.cancelEdit(); return; }
    if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); this.commitEdit(e.shiftKey ? -1 : 1, 0); return; }
    if (e.key === 'Enter' && e.altKey) {
      e.preventDefault();
      const ed = this.editor, p = ed.selectionStart;
      ed.value = ed.value.slice(0, p) + '\n' + ed.value.slice(ed.selectionEnd);
      ed.setSelectionRange(p + 1, p + 1);
      return;
    }
    if (e.key === 'Tab') { e.preventDefault(); this.commitEdit(0, e.shiftKey ? -1 : 1); }
  }

  // ------------------------------------------------------------------ overlay (dropdown etc.)
  positionOverlay() {
    if (!this.model) return;
    const [r, c] = this.sel;
    const rc = this.cellRect(r, c);
    if (this.editing) {
      const e = this.cellRect(this.editing.r, this.editing.c);
      this.editor.style.left = e.x + 'px'; this.editor.style.top = e.y + 'px';
    }
    this.positionDvButton(rc);
    // input message
    const dv = this.model.dvAt(r, c);
    if (dv && dv.showInputMessage && dv.prompt && !this.model.rowHidden(r)) {
      this.tip.innerHTML = '';
      if (dv.promptTitle) { const b = document.createElement('b'); b.textContent = dv.promptTitle; this.tip.appendChild(b); }
      const p = document.createElement('div'); p.textContent = dv.prompt; this.tip.appendChild(p);
      Object.assign(this.tip.style, { display: 'block', left: (rc.x + Math.min(rc.w, 40)) + 'px', top: (rc.y + rc.h + 4) + 'px' });
    } else this.tip.style.display = 'none';
    if (this.dvList.style.display !== 'none') this.placeList(rc);
  }
  positionDvButton(rc) {
    const [r, c] = this.sel;
    rc = rc || this.cellRect(r, c);
    const dv = this.model.dvAt(r, c);
    const show = dv && dv.type === 'list' && dv.showDropDown !== '1' && !this.model.rowHidden(r) && rc.w > 0 && !this.app.cannotEdit(this.model.si, r, c);
    if (!show) { this.dvButton.style.display = 'none'; return; }
    const sz = Math.max(16, Math.min(22, rc.h));
    Object.assign(this.dvButton.style, { display: 'block', left: (rc.x + rc.w + 1) + 'px', top: (rc.y + rc.h - sz) + 'px', width: sz + 'px', height: sz + 'px' });
  }
  listValues(r, c) {
    const m = this.model, dv = m.dvAt(r, c);
    if (!dv || dv.type !== 'list' || !dv.f1) return [];
    const f = dv.f1;
    if (f.startsWith('"')) return f.slice(1, -1).split(',').map((s) => s.trim());
    let res = this.app.eng.evaluateRel(f, m.si, dv.r0, dv.c0, r, c, true);
    const out = [];
    const push = (v) => { if (v === null || v === undefined || v === '' ) return; out.push(v); };
    if (res instanceof Ref) res = res.isCell() ? res.value() : res.toMatrix();
    if (res instanceof Matrix) { for (const v of res.data) push(v); } else if (!isErr(res)) push(res);
    return out;
  }
  toggleList() {
    if (this.dvList.style.display !== 'none') { this.dvList.style.display = 'none'; return; }
    const [r, c] = this.sel;
    if (this.app.cannotEdit(this.model.si, r, c)) return;
    const vals = this.listValues(r, c);
    this.listItems = vals;
    const cur = this.app.eng.getValue(this.model.si, r, c);
    this.dvList.innerHTML = '';
    vals.forEach((v, i) => {
      const d = document.createElement('div');
      d.className = 'dv-item' + (cur !== null && toText(cur) === toText(v) ? ' current' : '');
      d.dataset.i = i;
      d.setAttribute('role', 'option');
      d.textContent = v instanceof Rich ? v.toString() : typeof v === 'number' ? formatValue(v, 'General') : toText(v);
      this.dvList.appendChild(d);
    });
    if (!vals.length) { const d = document.createElement('div'); d.className = 'dv-empty'; d.textContent = '(no items)'; this.dvList.appendChild(d); }
    this.dvList.style.display = 'block';
    this.placeList(this.cellRect(r, c));
    const curEl = this.dvList.querySelector('.current');
    if (curEl) curEl.scrollIntoView({ block: 'nearest' });
  }
  placeList(rc) {
    const z = this.model.zoom;
    const w = Math.max(rc.w + 20, 140);
    const below = this.H - (rc.y + rc.h);
    const maxH = Math.max(120, Math.min(320, below > 160 ? below - 8 : rc.y - 8));
    Object.assign(this.dvList.style, { left: rc.x + 'px', width: w + 'px', maxHeight: maxH + 'px', fontSize: Math.max(12, 14 * Math.min(1, z * 1.4)) + 'px' });
    const h = Math.min(this.dvList.scrollHeight, maxH);
    this.dvList.style.top = (below > h + 8 || rc.y < h ? rc.y + rc.h + 1 : rc.y - h - 1) + 'px';
  }
  async pickListItem(i) {
    this.dvList.style.display = 'none';
    const v = this.listItems[i];
    const [r, c] = this.sel;
    this.scroll.focus({ preventScroll: true });
    await this.app.setFromList(this.model.si, r, c, v);
  }
}
