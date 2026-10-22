// Emission Profile: base-year emissions, growth factor selection, the BAU dashboard and
// target setting, as in the workbook's Emissions Profile menu.
import { el, sheetTable, contentRows } from './sheetview.js';
import { SheetTabsPage, TIER_SHEETS, tierOf } from './sheetpage.js';

const TARGET = 'Target Setting ';
const titled = (app, re) => (ids) => ids.filter((id) => { const t = app.wb.charts[id].title; return t && t.text && re.test(t.text); });

export class ProfilePage extends SheetTabsPage {
  get title() { return 'Emission Profile'; }
  tabs() {
    const app = this.app;
    const [base, dash] = TIER_SHEETS[tierOf(app)] || TIER_SHEETS.District;
    return [
      { title: 'Base year Emissions', sheet: base, desc: 'Base-year emissions by sector and sub-sector.' },
      {
        title: 'Growth Factor Selection', sheet: TARGET, rows: [12, 12], cols: [3, 7], tableFirst: true,
        desc: 'Growth factor for the BAU projection',
        intro: 'Choose what the business-as-usual (BAU) emissions grow with. The BAU dashboard and target setting update straight away.',
        charts: () => [],
        after: (card) => {
          const si = app.eng.sheetIndex(TARGET);
          card.appendChild(el('h3', 'card-title spaced', 'Resulting BAU emissions'));
          const box = el('div', 'sv-scroll');
          box.appendChild(sheetTable(app, si, contentRows(app, si, 17, 21, 3, 4), { cols: [3, 4] }));
          card.appendChild(box);
        },
      },
      { title: 'Business As Usual (BAU) Dashboard', sheet: dash, desc: 'Business-as-usual emissions for the interim and target years.' },
      {
        title: 'Target Setting', sheet: TARGET, rows: [14, 21], cols: [3, 7], tableFirst: true,
        desc: 'Emission reduction targets',
        intro: 'Set the target for each interim year and the target year, either as a reduction percentage or as an emissions value.',
        charts: titled(app, /Target/i),
        before: (card) => this.targetMode(card),
      },
    ];
  }

  // the workbook's two option buttons, linked to K15 (1 = percentages, 2 = values)
  targetMode(card) {
    const app = this.app, si = app.eng.sheetIndex(TARGET);
    const cur = app.eng.getValue(si, 15, 11);
    const grp = el('div', 'sv-options target-mode');
    [[1, 'Targets as percentages (%)'], [2, 'Targets as values (MtCO₂e)']].forEach(([v, label]) => {
      const lab = el('label', 'sv-option' + (cur === v ? ' on' : ''));
      const rb = el('input');
      rb.type = 'radio'; rb.name = 'target-mode'; rb.checked = cur === v;
      rb.addEventListener('change', () => { app.applyCells([[si, 15, 11, v]]); this.render(); });
      lab.append(rb, el('span', null, label));
      grp.appendChild(lab);
    });
    card.appendChild(grp);
  }
}
