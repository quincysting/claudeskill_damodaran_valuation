# SPCX (SpaceX): valuation, Damodaran's way, as of 2026-09-25

**Value per share: $81.40 vs price $148.68.** The price is 183% of value. The inputs are Damodaran's June 2026 story for
SpaceX, updated with the Q2-2026 10-Q and current market data, and run through the aswath-finance engines, which are
calibrated against Damodaran's own spreadsheets. It is not investment advice.

## Result

| | Damodaran's June 2026 valuation | This update (2026-09-25) |
|---|---|---|
| Value per share | **$97.83** (reproduced exactly from the spreadsheet) | **$81.40** |
| Price | $135 (IPO offer) | $148.68 (close 2026-09-25) |
| Value of operating assets | $1,224bn | $1,012bn |
| Initial / terminal cost of capital | 8.37% / 8.25% | 10.42% / 8.86% |
| Riskfree rate (10-yr T-bond) | 4.56% | 5.17% |
| Cash / debt | $24.7bn + $75bn IPO proceeds / $22.9bn | $100.0bn (IPO received) / $39.4bn |
| Shares | 12,535m + IPO shares | 13,182m |

Damodaran called the $135 offer price "too richly priced for my tastes" and added that "as an investor, I would not sell
short, notwithstanding my value assessment", because mood and momentum can push prices anywhere
([Revisiting the SpaceX Valuation](https://aswathdamodaran.blogspot.com/2026/06/a-weeks-ago-i-assessed-value-of-spacex.html), 2026-06).
At $148.68 the gap has widened.

## What changed since June, and why the value fell

1. **Rates rose.** The 10-year Treasury went from 4.56% to 5.17%
   ([CNBC, 2026-09-25](https://www.cnbc.com/2026/09/25/treasury-yields-bonds-debt.html)). This alone is about −$8/share
   (sensitivity row "Riskfree back at June 4.56%": $89.59).
2. **Beta rose from 0.75 to 1.03.** Damodaran's method weights each business's industry beta by its revenues × EV/Sales.
   The June cost-of-capital sheet still used the April pre-prospectus weights (xAI at $100m, Launch $4.1bn, Starlink
   $11.4bn), even though the June input sheet had xAI at $3.2bn. With the actual trailing AI revenue ($5.1bn), the
   software beta gets its proper weight. Keeping the literal June weights gives $88.91.
3. **Terminal cost of capital 8.86%.** This is the June model's 8.25% steady-state choice moved up with the riskfree rate, the way
   Damodaran's industry-average method adjusts for riskfree changes. The literal 8.25% gives $95.89. The Ginzu template default
   (riskfree + 4.20% mature ERP = 9.37%) gives $72.52.
4. **Base year rolled forward** to the twelve months ending 2026-06-30. Revenue is $23.0bn (Space 4.1, Connectivity
   13.9, AI 5.1). Operating income is −$3.7bn and R&D is $12.2bn.

The story is unchanged. The engine uses Damodaran's 2036 revenue targets (Launch $40bn, Starlink $120bn, xAI $160bn, other
expansion $100bn), target margins (45% / 60% / 25% / 30%), sales-to-capital ratios and 15% terminal
ROIC.

## Sensitivity (one change at a time; `python spacex_model.py sep_2026_inputs.json --sensitivity`)

| Scenario | Value/share | vs price |
|---|---|---|
| Base case | 81.40 | −45.2% |
| June-model beta weights (xAI at $100m, stale) | 88.91 | −40.2% |
| Riskfree back at June 4.56% (terminal coc 8.25%) | 89.59 | −39.7% |
| Terminal cost of capital = June model's 8.25% | 95.89 | −35.5% |
| Terminal cost of capital = riskfree + mature ERP (9.37%) | 72.52 | −51.2% |
| xAI 2036 revenue $80bn (April 2026 view) | 73.14 | −50.8% |
| xAI 2036 revenue $240bn | 89.67 | −39.7% |
| Starlink target margin 50% / 70% | 75.14 / 87.67 | −49.5% / −41.0% |
| xAI sales-to-capital 0.5 in years 1–5 (closer to the H1-2026 capex pace) | 77.44 | −47.9% |
| All 2036 revenue targets ×0.75 / ×1.25 | 62.58 / 100.23 | −57.9% / −32.6% |

**What the price implies.** The 2036 revenue targets would need to be **1.89× larger**, $795bn of revenue in 2036
instead of $420bn, at the same margins and reinvestment efficiency. Damodaran's separate market-cap breakeven tool
(`market.py`, used in the trillion-dollar-market-cap and AI posts) points the same way. SpaceX's enterprise value of
$1.90trn needs about $832bn of revenue after 10 years, a 43% compounded annual growth rate from $23bn, at a 28.6%
after-tax operating margin and 15% ROIC. The market cap is $1.96trn and the enterprise value is $1.90trn after cash
and debt.

## Ginzu Diagnostics, answered

| Diagnostic | Figure | Comment |
|---|---|---|
| Growth vs industry | Forecast 86% next year and 55% in years 2–5, vs 11% for the industry and 23% last year | This is Damodaran's big-market story. The June post rejected the prospectus's $28trn TAM as fantasy, and its $420bn in 2036 is already far below the company's own aspirations. |
| Margins | Current R&D-adjusted margin 28% (a single year, distorted by R&D capitalization) → 8% in year 1 → 38% in year 10, vs 9% for the industry | Starlink's 60% and launch's 45% targets carry the value. |
| Reinvestment | Current sales-to-capital 0.27 | H1-2026 capex was $28.5bn, $15.8bn of it in AI ([CNBC](https://www.cnbc.com/2026/08/04/spacex-spcx-earnings-live-updates-q2-2026.html)), against about $10bn of year-1 reinvestment in the model. The reinvestment assumptions look optimistic next to the filings, which is a downside risk to this value. |
| Cost of capital | 10.4% initially, 8.9% steady state, vs 7.6% for aerospace/defense | The higher figure reflects the AI business. |
| Value / price | 55% | Close to the Ginzu "value seems low, revisit inputs" flag of 50%. No single reasonable input change closes the gap, though (see the breakeven). |

## Judgment calls and limits

- **Base segment margins** (8% / 10% / −5%) are Damodaran's June estimates. The 10-Q does not disclose segment profit.
- **Operating-country ERP weights** are from Damodaran's June sheet. The 10-Q has no geographic revenue split. The country ERPs are
  Damodaran's July-2026 update (US 4.42%, China 5.07%), and the rest of the world is the June sheet's 6.18%.
- **Debt rating** is the June model's Ba2/BB, with Damodaran's January-2026 spread table. Employee options are excluded, as in the June model;
  the 10-Q does not disclose option counts.
- **Share count** is taken from the 10-Q cover (13,182m, 2026-07-28). StockAnalysis shows 13.57bn shares
  ([link](https://stockanalysis.com/stocks/spcx/), 2026-09-25), which probably includes shares issued since. Using that
  count would lower value per share by about 2.9%, to about $79.
- **The story is Damodaran's June 2026 story.** Q2 showed AI revenue up 247% year on year and much heavier AI capex. These
  cut both ways (xAI's revenue path versus its reinvestment needs). They are shown as sensitivities, not folded into
  the base case.

## Files (this folder)

| File | What |
|---|---|
| `sep_2026_inputs.json` | Base-case inputs with a `_sources` entry for every input |
| `damodaran_june_2026_inputs.json` / `damodaran_june_2026_output.txt` | Damodaran's June 2026 model in segment form; reproduces its $97.83 |
| `valuation_output_2026-09-25.txt` | Engine output: year-by-year table, cost of capital, Ginzu Diagnostics |
| `sensitivity_2026-09-25.txt` | Scenario table and revenue breakeven |
| `breakeven_marketcap.json` / `_output.txt` | Damodaran's market-cap breakeven tool applied to SPCX |
| `spacex_model.py` | The June-model SpaceX segment mechanics (from SpaceX2026IPOUpdated.xlsx) feeding `ginzu.py`; `--calibrate` reproduces its $97.8277 |
| `SPCX_Valuation_2026-09-25.xlsx` | Shareable live-formula workbook (Summary, Inputs, Scenarios, Sensitivity, WACC, Segments, DCF, Diagnostics, Checks, Sources); its Checks tab ties it to the engine |
| `build_workbook.py` / `finalize_workbook.ps1` | Rebuild the workbook: `python build_workbook.py`, then the PowerShell script (Excel adds the data tables and recalculates) |

Reproduce: `python spacex_model.py sep_2026_inputs.json` (and `--sensitivity`); `--calibrate` needs `finance\.venv`.

**Sources.** SEC Q2-2026 10-Q ([spcx-20260630](https://www.sec.gov/Archives/edgar/data/0001181412/000162828026052535/spcx-20260630.htm)),
SEC Q2-2026 earnings release ([link](https://www.sec.gov/Archives/edgar/data/1181412/000162828026052515/earningsreleaseq22608042.htm)),
price ([CoinGecko](https://www.coingecko.com/en/stocks/spacex)), 10-year yield ([CNBC](https://www.cnbc.com/2026/09/25/treasury-yields-bonds-debt.html)),
IPO details ([Wikipedia](https://en.wikipedia.org/wiki/Initial_public_offering_of_SpaceX)). Damodaran's model:
[SpaceX2026IPOUpdated.xlsx](https://pages.stern.nyu.edu/~adamodar/pc/blog/SpaceX2026IPOUpdated.xlsx). Damodaran's posts:
[To a Trillion(s) Dollars and beyond](https://aswathdamodaran.blogspot.com/2026/04/to-trillion-dollars-and-beyond-spacex.html) (2026-04),
[Revisiting the SpaceX Valuation](https://aswathdamodaran.blogspot.com/2026/06/a-weeks-ago-i-assessed-value-of-spacex.html) (2026-06).
