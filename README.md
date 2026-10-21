# ASCENT India – web app

The WRI India **ASCENT (Beta V25)** Excel workbook as a static web app. Every sheet, cell style,
formula, dropdown, validation rule, navigation button, macro and chart from the `.xlsm` runs in the
browser. Nothing is sent to a server: inputs are saved in the browser's local storage.

## What is replicated

| Excel feature | Web app |
|---|---|
| 59 sheets, styles, merged cells, frozen panes, zoom, hidden rows/columns | Canvas renderer that draws each sheet like Excel |
| ~122,000 formulas (XLOOKUP, FILTER, LET, TEXTSPLIT, SUMIFS, OFFSET, …) | Formula engine in `site/js/engine` with live recalculation (checked against Excel's saved results: 123,122 of 123,139 cells match; the rest point at an external workbook) |
| Dropdown lists, data validation messages, input prompts | Dropdown button on the selected cell, same error/prompt text |
| Sheet protection (locked/unlocked cells) | Same rules; File → Protect/unprotect uses the workbook password |
| Navigation buttons, option ovals, form radio buttons | Clickable shapes running ports of the VBA |
| VBA: `Workbook_Open`, Questionnaire `Worksheet_Change` (tier F9), Module1–4 (scenario views, section navigation, transport/urban/AFOLU options, Copy Metadata) | `site/js/macros.js` |
| 45 charts (bar, line, area, pie, scatter) | Interactive Plotly charts in a ggplot2-style theme, with the workbook's colours and labels |
| Pivot tables feeding the dashboards | Recomputed after every change (Excel only refreshes them on Refresh All) |
| Conditional formatting, comments, hyperlinks | Supported |

View menu: formula bar, sheet tabs, headings and gridlines (hidden by default, as the workbook's
`Workbook_Open` hides them). File menu: export/import your inputs as JSON, download a sheet as CSV,
reset the workbook.

## Password

The app opens with a password screen (password: `Wrii1234`). This is a client-side gate on a static
site: it keeps casual visitors out but is not real security, because anyone who downloads the site
files can read the data. To change it, put the SHA-256 of the new password in `HASH` in
`site/js/auth.js`.

## Deploy with GitHub Actions (GitHub Pages)

1. Create a GitHub repository and push this folder to its `main` branch.
2. In the repository, open **Settings → Pages** and set **Source** to **GitHub Actions**.
3. The workflow `.github/workflows/deploy.yml` runs on every push. It extracts the workbook in
   `excel/`, builds `site/data`, checks the formula engine against Excel's results and publishes `site/`.
   The site URL appears on the workflow run and under Settings → Pages.

To update the model, replace the `.xlsm` in `excel/` and push; the workflow rebuilds everything.

## Run locally

`site/data` is already built, so you only need a static web server:

```bash
cd site
python3 -m http.server 8000      # then open http://localhost:8000
```

(Opening `index.html` directly from disk does not work, because browsers block loading the data files from `file://`.)

To rebuild the data yourself (Python 3 with `lxml`, Node 18+):

```bash
pip install -r tools/requirements.txt
python3 tools/extract.py excel/WRI_India_ASCENT_Beta_V25.xlsm build/raw
node --max-old-space-size=6000 tools/build.mjs build/raw site/data
node --max-old-space-size=6000 test/verify.mjs      # compare with Excel's saved values
```

## Differences from Excel

- Typing a new formula into a cell is not supported (values only); all of the workbook's own formulas work.
- `Metadata`, `City_master`, `ListsAndTables` and `GWP` are reference sheets and are read-only.
- The one external-workbook link has no saved data in the file, so cells that use it show the same errors Excel shows.
- The button that runs `DashbaseYear` shows Excel's "cannot run the macro" message, because that macro does not exist in the workbook.

## Layout

```
excel/            source workbook
tools/            extract.py (xlsx → JSON), build.mjs (compact data + dependency graph)
site/             the web app (index.html, css/, js/, vendor/plotly + ssf, data/)
test/             engine checks against Excel's cached values
.github/workflows deploy to GitHub Pages
```
