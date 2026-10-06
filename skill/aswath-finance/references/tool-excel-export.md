# Tool: export_xlsx.py — a valuation as a shareable Excel model

`scripts/export_xlsx.py` (needs `openpyxl`) turns the **same inputs JSON that `ginzu.py` runs** into a workbook with
investment-banking conventions: blue hard-coded inputs (each with its source), black formulas, green cross-sheet
links, named ranges (`ValuePerShare`, `SharePrice`, `WACC_Initial`, `OperatingAssets`, `AllChecksPass`), a Checks
tab and a disclaimer. Every calculation cell is a live Excel formula mirroring `ginzu.py` / `cost_of_capital.py`,
so the model recalculates when a colleague changes an input. Use it only after the valuation itself is done
(`valuation-engine.md`): it presents the engine's model, it does not replace it.

```
python scripts/export_xlsx.py inputs.json                 # -> inputs_valuation.xlsx next to the JSON
python scripts/export_xlsx.py inputs.json -o out.xlsx --verify   # Windows + Excel: recalc, cache values, report Checks
python scripts/export_xlsx.py --selftest                  # needs pycel (finance/.venv)
```

## Why trust it

`--selftest` rebuilds all 45 `ginzu.py` calibration and branch cases (31 published Damodaran valuations, 14 template
branches: failure book/fair value, reinvestment lags 0-3, NOL, trapped cash, options at price and at value, IPO
shares at value, riskfree after year 10, terminal overrides, R&D, driver paths, segments) and all 21
cost-of-capital cases (every beta, ERP and debt approach, preferred, industry average, distribution, direct) as
workbooks, recalculates each with an independent Excel formula engine (pycel) and requires the engine's value
(to 1e-7) and the published cost of capital (to 1e-6): 66/66. `--verify` repeats the tie-out in real Excel on the actual file,
including the data tables; the Checks tab does it every time the workbook is opened.

## What the workbook contains

| Tab | Content |
|---|---|
| Cover | Headline value, price, check status, contents, colour convention |
| Summary | Key outputs, EV-to-equity bridge, scenario values, football field (52-week range and analyst target if given), engine breakeven target margin |
| Inputs | Every input with its `_sources` entry; column D = live value of the active scenario |
| Scenarios | Selector (D5) that switches the whole model; value per scenario via a native Excel data table |
| Sensitivity | Flex cells + data tables: initial x terminal cost of capital, growth x margin, riskfree (value and WACC) |
| WACC | Damodaran's cost-of-capital worksheet for the approach in the `cost_of_capital` block (or the direct input) |
| DCF | Years 1-10 + terminal: growth, margin, tax drift, NOL, reinvestment (lag via CHOOSE), cost-of-capital drift, ROIC; valuation and equity bridge; R&D converter |
| Dilution | Only with options or IPO shares at value: the Ginzu circular reference solved by 60 explicit fixed-point steps (no iterative calculation) |
| Diagnostics | The six Ginzu Diagnostics questions with industry benchmarks (`industry_us`) |
| Checks | Tie-outs to the engine (value, WACC, scenarios, table centres, bridge, convergence) + the Ginzu rules as warnings |
| Sources | `_sources`, notes, methodology, data vintage, disclaimer |

## Optional `_workbook` block (ignored by ginzu.py)

```json
"_workbook": {"company": "NVIDIA Corp.", "ticker": "NVDA", "valuation_date": "2025-01-15", "currency": "$",
  "units": "$ in millions except per share",
  "market": {"week52_low": 60, "week52_high": 153, "analyst_target": 170},
  "scenarios": [{"name": "Target margin 50%", "set": {"target_margin": 0.50}},
                {"name": "10% chance of failure", "set": {"failure_probability": 0.10}}],
  "notes": ["Judgment call: ..."]}
```

Scenario keys must be scalar inputs of the base case (or one of the Ginzu defaults the Inputs tab always shows: NOL, IPO,
lag, failure, keep-effective-tax); every scenario value is also run through `ginzu.py` and tied out on Checks. Keep
market data that is not Damodaran's (analyst targets, trading range) labelled as third party.

## Limits

- **Segments** (`segments`, the Nvidia market x share structure) are exported as the fixed revenue-growth and margin
  paths `ginzu.segment_paths` derives; scenarios on inputs that feed the segment build are refused. Segment models
  with their own mechanics (SpaceX: `valuation/SPCX/spacex_model.py`) export the same way from their `build()`
  output; a live segment sheet needs a bespoke builder (`valuation/SPCX/build_workbook.py` is one).
- Banks and insurers (`financial_firms.py`) are not exported.
- The synthetic-rating table is the one for the firm type at export; its rows are on the WACC tab.
- Data tables are native Excel What-If tables: Excel computes them on open (Google Sheets does not support them).
  Without `--verify` the file carries no cached values, so previewers that do not calculate show blanks.
- Flex cells: the growth flex shifts years 1-5, the margin flex shifts the target margin (or the margin path in
  parallel), the riskfree flex moves the cost of debt and equity, riskfree after year 10, terminal growth (unless
  overridden) and the terminal cost of capital; a direct cost-of-capital input moves 1:1 with it.
