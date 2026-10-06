# Valuation engine: Damodaran's Ginzu model in code

`scripts/ginzu.py` is Damodaran's `fcffsimpleginzu.xlsx` FCFF model transcribed formula by formula;
`scripts/cost_of_capital.py` is the matching cost-of-capital worksheet; `data/ginzu_data.json` holds the Damodaran data
tables from the January 2026 template (mature ERP 4.23%, ERPs and default spreads for 192 countries,
21 regional ERPs, 96 US and 96 global industries, rating spreads, cost-of-capital distribution).

## Why trust it

`python scripts/ginzu.py --selftest` replays 45 cases: 31 of Damodaran's own published valuations (27 spreadsheets, 4 also
rebuilt from segment inputs) plus 14 template branch tests. Every one matches the published value per share (all exact except
the Feb-2024 Nvidia segment rebuild, 0.04%):

| Group | Damodaran spreadsheets (value/share) |
|---|---|
| Standard Ginzu | Amazon 2024 (155.72), Apple 2024 (176.79), Coca-Cola 2024 (65.16), Disney 2023 (87.52), Google 2024 (143.07), Intel 2024 (29.11), Meta 2024 (445.10), Microsoft 2024 (355.88), Netflix 2023 (238.08), Nike 2024 (91.75), Starbucks 2024 (57.60), Walgreens 2024 (7.53), template (7.19) |
| Nvidia (market x share segments) | Jun 2023 (237.55), Feb 2024 (436.34), Sep 2024 (86.68), Jan 2025 (77.51) |
| DIY / segment / IPO models | Tesla Nov 2021, Jan 2023, Oct 2023, Jan 2024, Jan 2025 (571.29 / 129.90 / 179.56 / 183.75 / 147.16); SpaceX IPO 2026 and update (503.20 / 97.83); Birkenstock IPO 2023 (41.33); Instacart IPO 2023 (29.32); Adani 2023 (947.79) |

The cost-of-capital results for all 19 files that use the standard worksheet (18 company valuations + the
template; detailed, industry-average incl. multi-business, and distribution approaches) are also reproduced
exactly by `calibrate.py`. `python scripts/cost_of_capital.py --selftest` replays 21 recalculations of the
current template (`--build-cases`, pycel) covering every worksheet choice: the four approaches, all five beta
options, all four ERP options, direct / actual-rating / synthetic (firm types 1 and 2) cost of debt, preferred
stock — 21/21 within 1e-6. Spreadsheets:
`pages.stern.nyu.edu/~adamodar/pc/blog/<Name>.xlsx`, linked from the posts.

Branches the published files never use were checked by recalculating the Ginzu template with an
independent Excel formula engine (`branch_test.py`, pycel) under each override: failure (book and
fair-value proceeds), reinvestment lags 0/2/3, effective tax kept, NOL, loss-making years, riskfree
after year 10, negative terminal growth, trapped cash, options at price, terminal cost of capital and
ROIC overrides — 14/14 exact. 

## Hard rules when valuing a company

1. **Never compute a DCF by hand or in ad-hoc code.** Write an inputs JSON and run `ginzu.py`.
2. **Every input needs a source**, recorded in a `_sources` object in the JSON (keys starting with `_`
   are ignored by the engine). Allowed sources:
   - company filings (10-K/10-Q/annual report, trailing 12 months) — name the filing and date;
   - Damodaran's data (`data/ginzu_data.json` via `cost_of_capital.py`) — "Damodaran data, Jan 2026";
   - Damodaran's own earlier valuation of the company (`--from-case`, `references/cases-*.md`) — cite the post;
   - a story judgment — say so explicitly and tie it to Damodaran's framework (story → 3P test → input) with a post citation.
   Never fill a number from memory. If a base-year figure is unavailable, ask the user or fetch it.
3. **Keep the Ginzu default assumptions** (stable growth = riskfree, terminal ROIC = cost of capital,
   cost of capital drifts to riskfree + mature ERP, no failure, one-year reinvestment lag, tax drifts to
   marginal) unless there is a stated, cited reason — each published model documents its overrides.
4. **Report the Ginzu Diagnostics, not new rules.** The engine prints only rules Damodaran states (growth ≤ riskfree,
   ROIC < cost of capital destroys value, the value/price 0.5×–2× flag) plus the Ginzu Diagnostics sheet
   (your forecasts vs industry averages with its questions). Answer those questions in the write-up.

## Workflow

1. **Start from Damodaran's work.** `python scripts/ginzu.py --from-case <Name>` lists/prints the published inputs
   (any name in the table above, e.g. NvidiaJan2025, Nvidia2024Sep, MSFT2024, Netflix2023). Also read the company's section in `references/cases-*.md`
   for the story and later revisions. State the last published value, price and date.
2. **Update the base year** (TTM): revenues, EBIT, R&D history (if capitalizing; life from the Ginzu lookup),
   book equity, book debt incl. lease liabilities, cash, cross-holdings, minority interests, shares,
   price, effective tax rate; marginal tax rate from `ginzu_data.json` → `countries[<country>].tax_rate`.
3. **Cost of capital** — put a `cost_of_capital` block in the inputs (or run `cost_of_capital.py`):
   ```json
   "cost_of_capital": {"approach": "detailed", "shares_outstanding": 24490, "price": 123,
     "beta": {"approach": "single_us", "industry": "Semiconductor"},
     "erp": {"approach": "countries", "weights": {"United States": 26966, "China": 10306, "Taiwan": 13405}},
     "debt": {"approach": "rating", "rating": "A2/A", "book_debt": 10225, "interest_expense": 249, "maturity": 3}}
   ```
   Industry names must match the data tables (`industries_us` keys); ERP weights are revenues by country
   (`countries` keys) or region (`regions` keys). Riskfree = the government bond rate in the valuation
   currency on the valuation date (cite it).
4. **Story → drivers.** Revenue growth (year 1, years 2–5) or `segments` (market × share, as in the
   Nvidia 2025 model); target operating margin and convergence year; sales-to-capital (anchor on the
   industry `Sales/Capital`, the company's recent ratio, and Damodaran's earlier choice); failure probability
   for young/distressed firms; terminal ROIC only above cost of capital with a moat story that passes Damodaran's tests.
5. **Run** `python scripts/ginzu.py inputs.json` (add `industry_us`, `revenues_last_year` for full
   Diagnostics). Then `--breakeven target_margin` (or `growth_years_2_5`, `sales_to_capital_1_5`,
   `initial_cost_of_capital`) to state what the market price implies.
6. **Write up**: inputs table with a source column, value vs price with date, the Diagnostics answers,
   breakeven, and which inputs move value most. Label clearly what is Damodaran's (data, defaults, prior inputs)
   and what is judgment.
7. **Share as a spreadsheet** (when asked for an Excel/banker-style model): `python scripts/export_xlsx.py
   inputs.json --verify` turns the same inputs file into a live-formula workbook; see `tool-excel-export.md`.

## Input reference

See the docstring at the top of `scripts/ginzu.py` (all fields, defaults, overrides) and of
`scripts/cost_of_capital.py` (approaches). Rates are decimals; money in one consistent unit.

## Limits (not in the engine)

- Operating-lease converter (post-2019 accounting already puts lease liabilities in debt — include them in `book_debt`).
- Banks and financial-services firms: Damodaran values them with FCFE/dividends and regulatory capital, not FCFF — use `financial_firms.py` (`tool-financial-firms.md`).
- Convertible debt in the cost of capital; segments with different sales-to-capital ratios need an explicit `reinvestment_path`.

## Refreshing (each January, when the template is updated)

Damodaran's spreadsheets are mirrored in the maintenance workspace's `finance/spreadsheets/` folder (post-linked
files; `tools\` = the website calculators) and the maintenance Python env is `finance\.venv` (openpyxl, xlrd, pycel).
Download `pages.stern.nyu.edu/~adamodar/pc/fcffsimpleginzu.xlsx` (and any new company spreadsheets) there,
then `python scripts/calibrate.py <files...>` (needs `openpyxl`) — it re-checks the engine and rewrites
`data/ginzu_data.json` and `data/calibration_cases.json` — and `python scripts/branch_test.py
<fcffsimpleginzu.xlsx>` (needs `pycel`) to re-verify every override branch, and `python scripts/cost_of_capital.py
--build-cases <fcffsimpleginzu.xlsx>` (needs `pycel`) to rebuild `data/coc_cases.json` for the new data. If the template's formulas
change, a calibration case will fail: update `ginzu.py` to match before using the new data.
