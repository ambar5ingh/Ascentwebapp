// Results: the workbook's "Final Results" dashboard comparing BAU, E&P and high-ambition
// scenarios against the target.
import { SheetTabsPage } from './sheetpage.js';

export class ResultsPage extends SheetTabsPage {
  get title() { return 'Results'; }
  tabs() {
    return [{ title: 'Scenario comparison', sheet: 'Dashboard- Scenario Comparison', desc: 'BAU, E&P and high-ambition scenarios against the target.' }];
  }
}
