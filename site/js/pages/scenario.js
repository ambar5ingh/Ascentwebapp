// Scenario Development: the Existing & Planned and High Ambition scenario sheets for 2030,
// 2040 or 2050. The workbook's own macros pick the rows to show, as its buttons do in Excel.
import { el } from './sheetview.js';
import { SheetTabsPage } from './sheetpage.js';

const YEARS = [['2030', 'IntermYear1'], ['2040', 'IntermYear2'], ['2050', 'TargetYear']];

export class ScenarioPage extends SheetTabsPage {
  constructor(app) { super(app); this.year = 0; }
  get title() { return 'Scenario Development'; }
  subtitle() { return 'Choose the actions for each sector and how far they go by each year. Emission reductions, energy savings and costs are calculated from your inventory.'; }
  tabs() {
    const run = (macro, sheet) => async () => {
      const app = this.app, si = app.eng.sheetIndex(sheet);
      app.activeSi = si;
      await app.runMacro(macro, { sheet: si });
      await app.runMacro(YEARS[this.year][1], { sheet: si });
    };
    return [
      { title: 'Existing & Planned Scenario', sheet: 'E&P Scenario', rows: [10, 2560], cols: [4, 15], tableFirst: true, desc: 'Existing and planned actions', run: run('EP_Scenario', 'E&P Scenario'), charts: () => [] },
      { title: 'High Ambition Scenario', sheet: 'High Ambition Scenario', rows: [10, 2560], cols: [4, 15], tableFirst: true, desc: 'High-ambition actions', run: run('Ambitious_Scenario', 'High Ambition Scenario'), charts: () => [] },
    ];
  }
  toolbar(wrap) {
    const bar = el('div', 'subsection-tabs year-tabs');
    YEARS.forEach(([y], i) => {
      const b = el('button', 'sub-tab' + (i === this.year ? ' active' : ''), y);
      b.addEventListener('click', async () => { this.year = i; await this.prepare(); this.render(); });
      bar.appendChild(b);
    });
    wrap.appendChild(bar);
  }
}
