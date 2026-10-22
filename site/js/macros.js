// Ports of the workbook's VBA (ThisWorkbook, Sheet46 = Questionnaire, Module1-4).
// Each macro uses a small VBA-like API over the app state: hidden rows/columns,
// sheet visibility, shape fill/line/text/alt-text/OnAction, cell values and scrolling.

const RGB = (r, g, b) => [r, g, b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
const WHITE = RGB(255, 255, 255), BLACK = RGB(0, 0, 0);
const NAVY = RGB(32, 56, 100), GREY = RGB(128, 128, 128), LIGHT = RGB(198, 217, 241), DARK = RGB(31, 78, 121);
const PWD_SHEETS = ['Questionnaire', 'A. Basic Information', 'D. Transportation', 'H. AFOLU', 'F. Waste water Emission',
  'E. Solid Waste', 'C. Electricity Generation', 'G. IPPU', 'Scenario Questionnaire', 'Base- Emission Profile_State',
  'Base- Emission Profile_City', 'Base- Emission Profile_Village', 'Base - Emission Profile_Distric', 'Dashboard- BAU-District',
  'Dashboard- BAU-City', 'Dashboard- BAU-State', 'Dashboard- BAU-Village'];

class VbaError extends Error {}

// VBA-like helpers bound to the app for one macro run
class Vba {
  constructor(app, caller, callerSheet) {
    this.app = app;
    this.caller = caller;
    this.callerSheet = callerSheet;
  }
  get active() { return this.app.activeSi; }
  si(name) {
    const i = this.app.eng.sheetIndex(name);
    if (i < 0) throw new VbaError(`Run-time error '9': Subscript out of range (${name})`);
    return i;
  }
  sheetOf(name) { return name === undefined ? this.active : this.si(name); }
  // Rows("a:b").Hidden = h   /  Rows.Hidden = False (no address)
  rows(addr, hidden, sheet) {
    const si = this.sheetOf(sheet);
    const st = this.app.rowState(si);
    if (addr === null) { if (hidden) st.set(1, 1048576, true); else st.showAll(); }
    else {
      const [a, b] = String(addr).split(':').map((x) => parseInt(x, 10));
      st.set(a, b || a, hidden);
    }
    this.app.layoutDirty(si);
  }
  cols(addr, hidden, sheet) {
    const si = this.sheetOf(sheet);
    const st = this.app.colState(si);
    if (addr === null) { if (hidden) st.set(1, 16384, true); else st.showAll(); }
    else {
      const num = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
      const [a, b] = String(addr).split(':');
      st.set(num(a), num(b || a), hidden);
    }
    this.app.layoutDirty(si);
  }
  activate(name) { this.app.activateSheet(this.si(name), { fromMacro: true }); }
  view(si = this.active) { return this.app.viewOf(si); }
  zoom(z) { this.app.setZoom(this.active, z / 100, { fromMacro: true }); }
  scrollRow(n) { this.view().topRow = n; }
  scrollColumn(n) { this.view().leftCol = n; }
  smallScrollDown(n) {
    const si = this.active, m = this.app.modelFor(si), v = this.view();
    let r = Math.max(v.topRow || 1, m.freezeRows + 1), k = 0;
    while (k < n && r < m.nRows) { r++; if (!m.rowHidden(r)) k++; }
    v.topRow = r;
  }
  select(addr) {
    const v = this.view();
    v.sel = this.app.parseAddr(addr);
    v.ensureSel = true;
  }
  gotoScroll(addr) {
    const v = this.view();
    const [r, c] = this.app.parseAddr(addr);
    v.sel = [r, c]; v.topRow = r; v.leftCol = c;
  }
  value(sheet, addr) {
    const si = this.sheetOf(sheet);
    const [r, c] = this.app.parseAddr(addr);
    return this.app.eng.getValue(si, r, c);
  }
  setValue(sheet, addr, v) {
    const si = this.sheetOf(sheet);
    const [r, c] = this.app.parseAddr(addr);
    this.app.macroWrite([[si, r, c, v]]);
  }
  visible(name, vis) { this.app.setSheetVisible(this.si(name), vis); }
  protect(name, on) { this.app.setProtected(this.si(name), on); }
  // Shape object for VBA-style property changes; throws like VBA if the shape does not exist
  shape(name, sheet) {
    const si = this.sheetOf(sheet);
    const obj = this.app.findShape(si, name);
    if (!obj) throw new VbaError(`The item with the specified name wasn't found (${name})`);
    return new ShapeRef(this.app, si, obj);
  }
  shapes(sheet) {
    const si = this.sheetOf(sheet);
    return (this.app.datas[si].drawing || []).map((d) => new ShapeRef(this.app, si, d.obj, d));
  }
}

class ShapeRef {
  constructor(app, si, obj, anchor) {
    this.app = app; this.si = si; this.obj = obj;
    this.anchor = anchor || (app.datas[si].drawing || []).find((d) => d.obj === obj);
    this.ov = app.shapeOverride(si, obj.name, true);
  }
  get name() { return this.obj.name; }
  get kind() { return this.obj.kind; }
  get isOval() { return this.obj.geom === 'ellipse'; }
  get alt() { return this.ov.descr !== undefined ? this.ov.descr : (this.obj.descr || ''); }
  set alt(v) { this.ov.descr = v; }
  get top() { // Shape.Top in points
    const m = this.app.modelFor(this.si);
    if (!m.rowPt) m.layout(this.app.zoomOf(this.si));
    const f = this.anchor.from;
    return m.anchorTopPt(f[2], f[3]);
  }
  set fill(c) { this.ov.fill = c; }
  set lineColor(c) { this.ov.line = c; }
  set lineVisible(v) { this.ov.lineVisible = v; }
  set textColor(c) { this.ov.textColor = c; }
  set onAction(s) { this.ov.onAction = s; }
}

// ------------------------------------------------------------------ shared routines
// Sheet46.ActivateShape (Questionnaire)
function activateShapeQ(sh, on) {
  if (on) {
    sh.ov.fillHidden = false; sh.fill = LIGHT; sh.lineVisible = true; sh.lineColor = WHITE;
    sh.textColor = WHITE; sh.ov.textHidden = false; sh.alt = 'L2';
  } else {
    sh.fill = WHITE; sh.ov.fillHidden = true; sh.lineVisible = false; sh.ov.textHidden = true; sh.alt = 'disabled';
  }
}
// Module1.ActivateInterimShape
function activateInterim(sh, on) {
  if (on) {
    sh.ov.fillHidden = false; sh.fill = LIGHT; sh.lineVisible = true; sh.lineColor = WHITE;
    sh.textColor = DARK; sh.ov.textHidden = false; sh.alt = '';
  } else {
    sh.fill = WHITE; sh.ov.fillHidden = true; sh.lineVisible = false; sh.ov.textHidden = true; sh.alt = 'disabled';
  }
}

// Module1.GetLayer / Module3.GetLayer (by Shape.Top in points)
function layerM1(top) {
  const t = [[17.27276, 1], [7.272756, 1], [99.32811, 2], [89.32811, 2], [167.2728, 3], [158.1818, 3], [187.5, 4], [154.0428, 4]];
  for (const [v, l] of t) if (Math.abs(top - v) < 2) return l;
  return 0;
}
function layerM3(top) {
  const t = [[8, 1], [9, 1], [10, 1], [95, 2], [82, 2], [88, 2]];
  for (const [v, l] of t) if (Math.abs(top - v) < 5) return l;
  return 0;
}

// Module1/Module3 Changeshapeclr_onclick: highlight the clicked navigation button by layer
function changeClrByTop(X, module) {
  let clicked;
  try { clicked = X.shape(X.caller); } catch (e) { return; } // caller is not on the active sheet
  const getLayer = module === 1 ? layerM1 : layerM3;
  const cl = getLayer(clicked.top);
  if (cl === 0) return;
  for (const sh of X.shapes()) {
    if (sh.kind === 'pic') continue;
    if (sh.alt === 'disabled') continue;
    if (module === 1) sh.lineColor = WHITE;
    const l = getLayer(sh.top);
    const same = sh.name === clicked.name;
    const otherText = module === 3 && cl === 3 ? DARK : WHITE;
    if (cl === 1) {
      if (l === 1) {
        if (same) { sh.fill = GREY; sh.lineColor = GREY; sh.textColor = WHITE; }
        else { sh.fill = NAVY; sh.textColor = WHITE; }
      } else if (l === 2) { sh.fill = LIGHT; sh.textColor = WHITE; }
    } else if (cl === 2) {
      if (l === 2) { if (same) { sh.fill = DARK; sh.textColor = WHITE; } else { sh.fill = LIGHT; sh.textColor = WHITE; } }
      else if (l === 3) { sh.fill = LIGHT; sh.textColor = WHITE; }
    } else if (cl === 3) {
      if (l === 3) { if (same) { sh.fill = DARK; sh.textColor = WHITE; } else { sh.fill = LIGHT; sh.textColor = otherText; } }
      else if (l === 4) { sh.fill = LIGHT; sh.textColor = otherText; }
    } else if (cl === 4) {
      if (l === 4) { if (same) { sh.fill = DARK; sh.textColor = WHITE; } else { sh.fill = LIGHT; sh.textColor = WHITE; } }
    }
  }
}

// Module2 Changeshapeclr_onclick: layers from the alternative text ("L1"/"L2")
function changeClrByAlt(X) {
  let clicked;
  try { clicked = X.shape(X.caller); } catch (e) { return; }
  const layerOf = (sh) => { const t = sh.alt.trim().toUpperCase(); return t.includes('L1') ? 1 : t.includes('L2') ? 2 : 0; };
  const cl = layerOf(clicked);
  if (cl === 0) return;
  for (const sh of X.shapes()) {
    if (sh.kind === 'pic' || sh.isOval) continue;
    if (sh.alt.trim().toUpperCase().includes('DISABLED')) continue;
    const l = layerOf(sh);
    const same = sh.name === clicked.name;
    if (cl === 1) {
      if (l === 1) {
        if (same) { sh.fill = GREY; sh.lineVisible = false; sh.textColor = WHITE; }
        else { sh.fill = NAVY; sh.textColor = WHITE; }
      } else if (l === 2) { sh.fill = LIGHT; sh.lineVisible = false; sh.textColor = WHITE; }
    } else if (cl === 2 && l === 2) {
      sh.fill = same ? DARK : LIGHT; sh.lineVisible = false; sh.textColor = WHITE;
    }
  }
}

const ovalLine = (X, names) => { for (const n of names) X.shape(n).lineColor = BLACK; };
const ovalFill = (X, map) => { for (const [n, c] of Object.entries(map)) X.shape(n).fill = c; };
const protectQ = (X) => X.protect('Questionnaire', true);

// ------------------------------------------------------------------ Module1 (scenario sheets)
function hidesectionsInScenario(X) {
  const v = X.value('Questionnaire', 'F9');
  const H = (a) => X.rows(a, true);
  if (v === 'Village' || v === 'Town') { ['181:200', '1034:1053', '1884:1903', '749:859', '1601:1711', '2450:2560'].forEach(H); }
  else if (v === 'District') { ['181:200', '1034:1053', '1884:1903'].forEach(H); }
  else if (v === 'City') { ['181:200', '1034:1053', '1884:1903', '749:859', '1601:1711', '2450:2560', '654:740', '1506:1592', '2355:2441'].forEach(H); }
}
function hideRowsInScenarioTranspo(X) {
  for (const s of ['E&P Scenario', 'High Ambition Scenario']) {
    for (const a of ['212:279', '281:323', '1064:1131', '1133:1175', '1914:1981', '1983:2026']) X.rows(a, false, s);
  }
  for (const s of ['E&P Scenario', 'High Ambition Scenario']) {
    const v = X.value(s, 'BF179');
    if (v === 2) ['212:279', '1064:1131', '1914:1981'].forEach((a) => X.rows(a, true, s));
    else if (v === 1) ['281:323', '1133:1175', '1983:2026'].forEach((a) => X.rows(a, true, s));
  }
}
function hideRowsBasedOnC(X) {
  const blocks = [[53, 134], [906, 987], [1756, 1837]];
  for (const s of ['E&P Scenario', 'High Ambition Scenario']) for (const [a, b] of blocks) X.rows(`${a}:${b}`, false, s);
  for (const [a, b] of blocks) {
    for (let r = a; r <= b; r++) {
      if (X.value('E&P Scenario', 'C' + r) === 1) X.rows(String(r), true, 'E&P Scenario');
      if (X.value('High Ambition Scenario', 'C' + r) === 1) X.rows(String(r), true, 'High Ambition Scenario');
    }
  }
}
const INTERIM = { y1: 'Rectangle: Rounded Corners 483', y2: 'Rectangle: Rounded Corners 484', ty: 'Rectangle: Rounded Corners 485',
  a1: 'Rectangle: Rounded Corners 478', a2: 'Rectangle: Rounded Corners 486', at: 'Rectangle: Rounded Corners 480' };
function epButtons(X) {
  activateInterim(X.shape(INTERIM.y1), true); activateInterim(X.shape(INTERIM.y2), true); activateInterim(X.shape(INTERIM.ty), true);
  activateInterim(X.shape(INTERIM.a1), false); activateInterim(X.shape(INTERIM.a2), false); activateInterim(X.shape(INTERIM.at), false);
  X.shape(INTERIM.a1).onAction = ''; X.shape(INTERIM.a2).onAction = ''; X.shape(INTERIM.at).onAction = '';
  X.shape(INTERIM.y1).onAction = 'IntermYear1'; X.shape(INTERIM.ty).onAction = 'TargetYear'; X.shape(INTERIM.y2).onAction = 'IntermYear2';
}
function ambButtons(X) {
  activateInterim(X.shape(INTERIM.y1), false); activateInterim(X.shape(INTERIM.y2), false); activateInterim(X.shape(INTERIM.ty), false);
  activateInterim(X.shape(INTERIM.a1), true); activateInterim(X.shape(INTERIM.a2), true); activateInterim(X.shape(INTERIM.at), true);
  X.shape(INTERIM.y2).onAction = ''; X.shape(INTERIM.y1).onAction = ''; X.shape(INTERIM.ty).onAction = '';
  X.shape(INTERIM.a1).onAction = 'IntermYear1'; X.shape(INTERIM.at).onAction = 'TargetYear'; X.shape(INTERIM.a2).onAction = 'IntermYear2';
}
// Module1 macros wrap their body in On Error GoTo ErrorHandler -> MsgBox
async function withHandler(X, fn) {
  try { await fn(); } catch (e) {
    if (e instanceof VbaError) await X.app.alert('Microsoft Excel', 'Error: ' + e.message);
    else throw e;
  }
}

const MACROS = {
  // ---------------------------------------------------------------- Module1
  EP_Scenario: (X) => withHandler(X, () => {
    X.cols(null, false); X.rows(null, false); X.zoom(55);
    X.activate('E&P Scenario'); X.scrollRow(1); X.scrollColumn(1);
    epButtons(X);
    hidesectionsInScenario(X); hideRowsInScenarioTranspo(X); hideRowsBasedOnC(X);
    changeClrByTop(X, 1);
    X.cols('P:Z', true);
  }),
  Ambitious_Scenario: (X) => withHandler(X, () => {
    X.activate('High Ambition Scenario'); X.scrollRow(1); X.scrollColumn(1);
    X.rows(null, false); X.cols(null, false); X.zoom(55);
    hidesectionsInScenario(X);
    ambButtons(X);
    hideRowsInScenarioTranspo(X); hideRowsBasedOnC(X);
    changeClrByTop(X, 1);
    X.cols('P:Z', true);
  }),
  IntermYear1: (X) => withHandler(X, () => {
    X.rows(null, false); X.zoom(55); X.scrollRow(1); X.scrollColumn(1);
    hidesectionsInScenario(X); hideRowsInScenarioTranspo(X); hideRowsBasedOnC(X);
    X.rows('823:2452', true);
    changeClrByTop(X, 1);
  }),
  IntermYear2: (X) => withHandler(X, () => {
    X.rows(null, false); X.zoom(55);
    hidesectionsInScenario(X); hideRowsInScenarioTranspo(X); hideRowsBasedOnC(X);
    X.rows('1712:2560', true); X.rows('8:861', true);
    changeClrByTop(X, 1);
  }),
  TargetYear: (X) => withHandler(X, () => {
    X.rows(null, false);
    hidesectionsInScenario(X); hideRowsInScenarioTranspo(X); hideRowsBasedOnC(X);
    X.rows('8:1711', true);
    changeClrByTop(X, 1);
  }),
  TaketoScenario_sheet: (X) => withHandler(X, () => {
    X.activate('E&P Scenario'); X.scrollRow(1); X.scrollColumn(1);
    X.rows(null, false);
    hideRowsInScenarioTranspo(X); hidesectionsInScenario(X); hideRowsBasedOnC(X);
    changeClrByTop(X, 1);
    epButtons(X);
  }),

  // ---------------------------------------------------------------- Module2 (Questionnaire sections)
  Taketobasicinfo: (X) => {
    protectQ(X); X.rows(null, false); X.rows('807:1048576', true);
    ovalLine(X, ['Oval 25', 'Oval 26']); ovalFill(X, { 'Oval 25': WHITE, 'Oval 26': WHITE });
    X.zoom(55); X.gotoScroll('D8');
    changeClrByAlt(X);
    X.rows('42:806', true);
  },
  TaketoBuildingEnergy: (X) => {
    protectQ(X); X.rows(null, false); X.rows('807:1048576', true); X.zoom(55); X.smallScrollDown(30);
    changeClrByAlt(X);
    X.rows('7:41', true); X.rows('139:806', true); X.select('F47');
  },
  TaketoTransportation: (X) => {
    protectQ(X); X.rows(null, false); X.rows('807:1048576', true); X.zoom(55); X.smallScrollDown(30);
    changeClrByAlt(X);
    ovalFill(X, { 'Oval 48': WHITE, 'Oval 47': WHITE });
    X.rows('7:162', true); X.rows('207:806', true);
    ovalLine(X, ['Oval 48', 'Oval 47']);
    const v = X.value('Questionnaire', 'H142');
    if (v === 2) { X.rows('170:175', true, 'Questionnaire'); X.shape('Oval 47').fill = BLACK; }
    else if (v === 1) { X.rows('177:206', true, 'Questionnaire'); X.shape('Oval 48').fill = BLACK; }
    const n151 = X.value(undefined, 'N151'), l153 = X.value(undefined, 'L153');
    if (isZero(n151) && isZero(l153)) { X.rows('177:206', false); X.rows('170:175', false); }
    X.select('D163');
  },
  TaketoSolidWaste: (X) => {
    protectQ(X); X.rows(null, false); X.rows('807:1048576', true); X.zoom(55); X.smallScrollDown(30);
    changeClrByAlt(X);
    X.rows('7:214', true); X.rows('271:806', true); X.select('F216');
  },
  TaketoWastewater: (X) => {
    protectQ(X); X.rows(null, false); X.rows('807:1048576', true); X.zoom(55); X.smallScrollDown(30);
    changeClrByAlt(X);
    ovalLine(X, ['Oval 3', 'Oval 1', 'Oval 2']);
    X.rows('7:296', true); X.rows('349:806', true); X.select('E298');
  },
  TaketoAFOLU: (X) => {
    protectQ(X); X.rows(null, false);
    const f = X.value('Questionnaire', 'F9');
    X.rows('807:1048576', true);
    ovalLine(X, ['Oval 3', 'Oval 1', 'Oval 2']); ovalFill(X, { 'Oval 3': WHITE, 'Oval 1': WHITE, 'Oval 2': WHITE });
    ovalLine(X, ['Oval 23', 'Oval 24']); ovalFill(X, { 'Oval 23': WHITE, 'Oval 24': WHITE });
    if (f === 'City') { X.rows('7:348', true); X.rows('376:806', true); }
    else { X.rows('7:688', true); X.select('H689'); }
    X.zoom(55); X.smallScrollDown(30);
    changeClrByAlt(X);
  },
  Taketoippu: (X) => {
    protectQ(X); X.rows(null, false); X.rows('807:1048576', true); X.zoom(55); X.smallScrollDown(30);
    changeClrByAlt(X);
    X.rows('7:387', true); X.rows('678:806', true); X.select('E393');
  },
  TaketoElectricity: (X) => {
    protectQ(X); X.rows(null, false); X.rows('807:1048576', true); X.zoom(55); X.smallScrollDown(30);
    changeClrByAlt(X);
    X.rows('7:138', true); X.rows('163:806', true); X.select('F139');
  },
  Hide_TransportMode_option1: (X) => {
    X.rows('170:175', false); X.rows('177:206', true);
    X.rows('37:427', true, 'D. Transportation'); X.rows('9:36', false, 'D. Transportation');
    X.setValue('E&P Scenario', 'BF179', 1); X.setValue('High Ambition Scenario', 'BF179', 1); X.setValue('Questionnaire', 'H142', 1);
    ovalFill(X, { 'Oval 48': BLACK, 'Oval 47': WHITE });
    X.select('D144');
  },
  Hide_TransportMode_option2: (X) => {
    X.rows('170:175', true); X.rows('177:206', false);
    X.rows('37:427', false, 'D. Transportation'); X.rows('9:36', true, 'D. Transportation');
    X.setValue('E&P Scenario', 'BF179', 2); X.setValue('High Ambition Scenario', 'BF179', 2); X.setValue('Questionnaire', 'H142', 2);
    ovalFill(X, { 'Oval 48': WHITE, 'Oval 47': BLACK });
    X.select('D149');
  },
  Hide_UrbanSettlement_option1: (X) => {
    X.rows('363:375', true); X.rows('357:361', false);
    ovalFill(X, { 'Oval 1': BLACK, 'Oval 2': WHITE, 'Oval 3': WHITE }); X.select('D358');
  },
  Hide_UrbanSettlement_option2: (X) => {
    X.rows('357:361', true); X.rows('363:364', false); X.rows('366:375', true);
    ovalFill(X, { 'Oval 2': BLACK, 'Oval 1': WHITE, 'Oval 3': WHITE }); X.select('D364');
  },
  Hide_UrbanSettlement_option3: (X) => {
    X.rows('366:375', false); X.rows('357:364', true);
    ovalFill(X, { 'Oval 3': BLACK, 'Oval 1': WHITE, 'Oval 2': WHITE }); X.select('D367');
  },
  Hide_AFOLU_option1: (X) => {
    X.rows('714:734', false); X.rows('735:737', true);
    ovalFill(X, { 'Oval 23': BLACK, 'Oval 24': WHITE }); X.select('F720');
  },
  Hide_AFOLU_option2: (X) => {
    X.rows('717:735', true); X.rows('736:737', false);
    ovalFill(X, { 'Oval 23': WHITE, 'Oval 24': BLACK }); X.select('H737');
  },

  // ---------------------------------------------------------------- Module3 (sheet navigation)
  taketoBuilding_sheet: (X) => {
    X.rows(null, false); X.activate('Base Year GHG Inventory'); X.scrollRow(1); X.scrollColumn(1);
    X.rows('99:1048576', true); changeClrByTop(X, 3); X.zoom(55);
  },
  TaketoTransportation_sheet: (X) => {
    X.rows(null, false);
    const v = X.value('Questionnaire', 'H142');
    if (v === 2) { X.activate('D. Transportation'); X.scrollRow(1); X.scrollColumn(1); }
    else if (v === 1) { X.activate('Base Year GHG Inventory'); X.rows(null, false); X.rows('7:87', true); X.rows('130:500', true); }
    changeClrByTop(X, 3);
  },
  TaketoSolidWaste_sheet: (X) => { X.activate('E. Solid Waste'); X.scrollRow(1); X.scrollColumn(1); changeClrByTop(X, 3); },
  TaketoWastewater_sheet: (X) => { X.rows(null, false); X.activate('F. Waste water Emission'); X.scrollRow(1); X.scrollColumn(1); changeClrByTop(X, 3); },
  TaketoAFOLU_sheet: (X) => {
    X.activate('H. AFOLU');
    const v = X.value('Questionnaire', 'F9');
    X.scrollRow(1); X.scrollColumn(1);
    if (v === 'City') { X.rows('7:253', true); X.rows('282:500', true); }
    changeClrByTop(X, 3);
  },
  TaketoIPPU_sheet: (X) => { X.activate('G. IPPU'); X.scrollRow(1); X.scrollColumn(1); changeClrByTop(X, 3); },
  Taketoelectricity_sheet: (X) => { X.activate('C. Electricity Generation'); X.scrollRow(1); X.scrollColumn(1); changeClrByTop(X, 3); },
  TaketoInventory_sheet: (X) => {
    X.protect('Questionnaire', false);
    X.rows(null, false, 'Questionnaire');
    X.activate('Questionnaire'); X.scrollRow(1); X.scrollColumn(1);
    const v = X.value('Questionnaire', 'F9');
    if (v === 'District') X.rows('12:13', true);
    else if (v === 'Village') X.rows('12', true);
    else if (v === 'City') X.rows('13', true);
    changeClrByTop(X, 3);
  },
  Taketoassumption_sheet: (X) => { X.activate('Assumption'); X.scrollRow(1); X.scrollColumn(1); changeClrByTop(X, 3); },
  TakeEF_sheet: (X) => { X.activate('Emission Factor'); X.scrollRow(1); X.scrollColumn(1); changeClrByTop(X, 3); },
  TakeBasicInfo_sheet: (X) => {
    X.protect('Questionnaire', false);
    X.activate('A. Basic Information'); X.scrollRow(1); X.scrollColumn(1); X.zoom(55); changeClrByTop(X, 3);
  },
  TaketoTargetSetting_sheet: (X) => {
    X.activate('Target Setting '); X.select('E14'); X.scrollRow(13); X.scrollColumn(1); changeClrByTop(X, 3);
  },
  TaketoGrowthFactor_sheet: (X) => {
    X.activate('Target Setting '); X.select('E12'); X.scrollRow(1); X.scrollColumn(1); changeClrByTop(X, 3);
  },
  TaketoFinalSettings_sheet: (X) => { X.activate('Dashboard- Scenario Comparison'); X.scrollRow(1); X.scrollColumn(1); changeClrByTop(X, 3); },
  TaketoConversionFactor_sheet: (X) => {
    X.activate('Cost and Conversion Factor'); X.scrollRow(1); X.scrollColumn(1); X.rows(null, false); X.rows('46:205', true); changeClrByTop(X, 3);
  },
  TaketoConversionFactor_sheet_costoverride: (X) => {
    X.activate('Cost and Conversion Factor'); X.scrollRow(1); X.scrollColumn(1); X.rows(null, false); X.rows('9:45', true); changeClrByTop(X, 3);
  },
  taketoDashboardBAU_sheet: (X) => {
    const v = X.value('Questionnaire', 'F9');
    // the VBA tests an undeclared variable for "Town", so Town falls through
    if (v === 'Village') X.activate('Dashboard- BAU-Village');
    else if (v === 'State') X.activate('Dashboard- BAU-State');
    else if (v === 'City') X.activate('Dashboard- BAU-City');
    else if (v === 'District') X.activate('Dashboard- BAU-District');
    changeClrByTop(X, 3);
  },
  TaketoBaseYearEmissions: (X) => {
    const v = X.value('Questionnaire', 'F9');
    if (v === 'Village') X.activate('Base- Emission Profile_Village');
    else if (v === 'State') X.activate('Base- Emission Profile_State');
    else if (v === 'City') X.activate('Base- Emission Profile_City');
    else if (v === 'District') X.activate('Base - Emission Profile_Distric');
    changeClrByTop(X, 3);
  },

  // ---------------------------------------------------------------- Module4
  DoNothing: (X) => {
    X.shape('Oval 25').fill = WHITE; X.shape('Oval 26').fill = BLACK; X.shape('Oval 25').lineColor = BLACK;
  },
  Copy_Metadata_to_Questionnaire_v7: async (X) => {
    if (!(await X.app.confirm('Are you sure?'))) return;
    X.shape('Oval 26', 'Questionnaire').fill = WHITE; X.shape('Oval 25', 'Questionnaire').fill = BLACK;
    X.shape('Oval 26', 'Questionnaire').lineColor = BLACK;
    copyMetadata(X.app);
  },
};

// Body of Copy_Metadata_to_Questionnaire_v7: 'Metadata helper'!B:N -> Questionnaire for the
// rows flagged 1 in column A, leaving the exempt cells alone. Returns the number of cells set.
export function copyMetadata(app) {
  const src = app.eng.sheetIndex('Metadata helper'), dst = app.eng.sheetIndex('Questionnaire');
  const exempt = new Set();
  const mark = (list) => {
    for (const part of list.split(',')) {
      const [a, b] = part.split(':');
      const [r1, c1] = app.parseAddr(a), [r2, c2] = b ? app.parseAddr(b) : [r1, c1];
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) exempt.add(r * 100 + c);
    }
  };
  mark('A15:A803'); mark('F21,F22,G29,H29,I29,J29,K29,K30,K31,K32,K33'); mark('H125:H130'); mark('E179,J186,G194');
  mark('D206:K206'); mark('D223:D233,J223:J225,K223:K225,K232,K233'); mark('D249:K249'); mark('J239:J248,K239:K248');
  mark('I331:I343,I345:I347'); mark('E548:E561'); mark('C599:C649');
  const eng = app.eng;
  const writes = [];
  let filled = 0;
  for (let r = 15; r <= 803; r++) {
    if (eng.getValue(dst, r, 1) !== 1) continue;
    for (let c = 2; c <= 14; c++) {
      if (exempt.has(r * 100 + c)) continue;
      let v = eng.getValue(src, r, c);
      // VBA writing "" leaves the cell empty
      if (v === undefined || v === '') v = null;
      if (c === 6 && v !== null && v !== '') filled++;
      writes.push([dst, r, c, v]);
    }
  }
  app.macroWrite(writes, { force: true });
  return filled;
}

// "Do you want to use secondary data?" is answered Yes (Oval 25 filled black)
export function usesSecondaryData(app) {
  const si = app.eng.sheetIndex('Questionnaire');
  const ov = app.shapeOverride(si, 'Oval 25', false);
  return !!ov && typeof ov.fill === 'string' && ov.fill.toUpperCase() === '000000';
}

function isZero(v) { return v === null || v === 0 || v === '' || v === false; }

// Worksheet_Change of "Questionnaire" (Sheet46), fired when F9 changes
export async function questionnaireChange(app, target) {
  const X = new Vba(app, null, app.eng.sheetIndex('Questionnaire'));
  const Q = 'Questionnaire';
  const qsi = X.si(Q);
  // the event runs with Questionnaire as the active sheet
  const v = X.value(Q, 'F9');
  const R = (a, h) => X.rows(a, h, Q);
  R(null, false); R('780:1048576', true);
  X.rows(null, false, 'Scenario Questionnaire');
  const s14 = X.shape('Rectangle: Rounded Corners 14', Q), s16 = X.shape('Rectangle: Rounded Corners 16', Q), s18 = X.shape('Rectangle: Rounded Corners 18', Q);
  activateShapeQ(s14, true); activateShapeQ(s16, true); activateShapeQ(s18, true);
  s14.onAction = 'TaketoAFOLU'; s16.onAction = 'Taketoippu'; s18.onAction = 'TaketoElectricity';
  for (const n of ['H. AFOLU', 'G. IPPU', 'C. Electricity Generation', 'Base- Emission Profile_State', 'Base- Emission Profile_City',
    'Base- Emission Profile_Village', 'Base - Emission Profile_Distric', 'Dashboard- BAU-City', 'Dashboard- BAU-Village',
    'Dashboard- BAU-District', 'Dashboard- BAU-State']) X.visible(n, true);
  const hide = (...names) => names.forEach((n) => X.visible(n, false));
  if (v === 'Village' || v === 'Town') {
    R('363:656', true); R('115:137', true); R('12:12', true);
    X.rows('134:173', true, 'Scenario Questionnaire'); X.rows('1025:1135', true, 'Scenario Questionnaire');
    hide('C. Electricity Generation', 'Base- Emission Profile_State', 'Base- Emission Profile_City', 'Base - Emission Profile_Distric',
      'Dashboard- BAU-State', 'Dashboard- BAU-City', 'Dashboard- BAU-District', 'G. IPPU');
    activateShapeQ(s16, false); activateShapeQ(s18, false);
    s16.onAction = ''; s18.onAction = '';
  } else if (v === 'City') {
    R('13:13', true); R('363:656', true); R('665:780', true); R('115:137', true);
    X.rows('134:173', true, 'Scenario Questionnaire'); X.rows('1024:1135', true, 'Scenario Questionnaire'); X.rows('770:1023', true, 'Scenario Questionnaire');
    hide('G. IPPU', 'C. Electricity Generation', 'Dashboard- BAU-State', 'Dashboard- BAU-Village', 'Dashboard- BAU-District',
      'Base- Emission Profile_State', 'Base - Emission Profile_Distric', 'Base- Emission Profile_Village');
    activateShapeQ(s16, false); activateShapeQ(s18, false);
    s16.onAction = ''; s18.onAction = '';
  } else if (v === 'District') {
    R('115:137', true); R('12:13', true);
    X.rows('134:173', true, 'Scenario Questionnaire');
    hide('C. Electricity Generation', 'Base- Emission Profile_City', 'Base- Emission Profile_Village', 'Base- Emission Profile_State',
      'Dashboard- BAU-City', 'Dashboard- BAU-Village', 'Dashboard- BAU-State');
    activateShapeQ(s18, false);
    s18.onAction = '';
  } else {
    R('11:13', true);
    hide('Base- Emission Profile_City', 'Base - Emission Profile_Distric', 'Base- Emission Profile_Village',
      'Dashboard- BAU-City', 'Dashboard- BAU-Village', 'Dashboard- BAU-District');
    s14.onAction = 'TaketoAFOLU'; s16.onAction = 'Taketoippu'; s18.onAction = 'TaketoElectricity';
  }
  void qsi; void target;
}

// ThisWorkbook.Workbook_Open
export function workbookOpen(app) {
  const X = new Vba(app, null, 0);
  X.activate('ASCENT India');
  X.select('A1'); X.scrollRow(1); X.scrollColumn(1);
  for (const n of PWD_SHEETS) { try { X.protect(n, true); } catch (e) { /* sheet missing */ } }
  // HideGridlinesInAllSheets: active window only
  app.viewOf(X.active).gridlines = false;
  X.zoom(55); X.scrollRow(1); X.scrollColumn(1);
  // HideSheetsBasedOnTabColor (red tabs)
  app.wb.sheets.forEach((s, i) => { if ((s.tabColor || '').toUpperCase() === 'FF0000') app.setSheetVisible(i, false); });
  X.visible('B. Building & Energy', false);
  X.visible('Navigation', false);
}

export function macroExists(name) { return !!resolve(name); }

function resolve(name) {
  let n = String(name || '').trim();
  n = n.replace(/^'[^']*'!/, '').replace(/^[^!]*!/, '');
  if (n.includes('.')) n = n.slice(n.lastIndexOf('.') + 1);
  const key = Object.keys(MACROS).find((k) => k.toLowerCase() === n.toLowerCase());
  return key ? MACROS[key] : null;
}

export async function runMacro(app, name, { caller, sheet } = {}) {
  const fn = resolve(name);
  if (!fn) {
    await app.alert('Microsoft Excel', `Cannot run the macro '${name}'. The macro may not be available in this workbook or all macros may be disabled.`);
    return;
  }
  const X = new Vba(app, caller, sheet);
  try {
    await fn(X);
  } catch (e) {
    if (e instanceof VbaError) await app.alert('Microsoft Visual Basic', e.message);
    else { console.error(e); await app.alert('Error', String(e && e.message || e)); }
  }
}
