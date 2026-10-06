# META (Meta Platforms): valuation, Damodaran's way, as of 2026-10-02

**Value per share: $757.32 vs price $728.08.** The price is 96% of value: Meta looks roughly fairly priced. The model
is Damodaran's Ginzu FCFF spreadsheet, transcribed in the aswath-finance engine. The story is Damodaran's February-2024
Meta story, re-based to filings through 2026-06-30. It is not investment advice.

Workbook: `META_Valuation_2026-10-02.xlsx`. It has live formulas, a scenario selector, data tables and a Checks tab
tied to the engine (all checks pass in Excel).

## Damodaran's record on Meta

| Valuation | Value | Price then | Post |
|---|---|---|---|
| April 2018 (after Cambridge Analytica) | ~$181 | $155 | [The Facebook Feeding Frenzy](https://aswathdamodaran.blogspot.com/2018/04/the-facebook-feeding-frenzy.html), 2018-04 |
| February 2024 | $445.10 (reproduced exactly) | $456.98 | [The Seven Samurai](https://aswathdamodaran.blogspot.com/2024/02/the-seven-samurai-how-big-tech-rescued.html), 2024-02 |
| **This update** | **$757.32** | **$728.08** | inputs below |

In 2022 Damodaran bought Meta (April) and added at $100 as the stock fell, while criticising the dual-class structure that leaves
outside shareholders unable to stop a bad strategic bet ([META Lesson 1](https://aswathdamodaran.blogspot.com/2022/11/meta-lesson-1-corporate-governance.html), 2022-11).
A May 2026 post named Meta's AI spending as the test of Zuckerberg's "next big thing" bets, after the failed Metaverse
spend ([An Ode to Restraint](https://aswathdamodaran.blogspot.com/2026/05/an-ode-to-restraint-lessons-from-tim.html), 2026-05).

## What changed since February 2024

- **Growth beat the story.** The 2024 model assumed 12% a year. Revenue grew 22% in 2024 and 22% in 2025, and is up
  27.7% over the last twelve months, to $228.2bn. Consensus has the next twelve months at about $281bn (+23%).
- **AI capital spending has collapsed reinvestment efficiency.** Over the last twelve months Meta reinvested about
  $94.8bn: capex of $92.4bn including finance-lease principal, less $22.7bn of depreciation, plus $25.1bn of net R&D
  capitalization. That bought $49.4bn of extra revenue, a sales-to-capital ratio of 0.52 against 1.5 in the 2024
  model. The company guides 2026 capex to $130-145bn.
- **Rates are higher.** The 10-year Treasury is at 5.28%, against 4.08% in February 2024.

The story's other levers are kept. Revenue grows 12% a year in years 2-5, the target margin is 40% by year 5, and the
terminal ROIC is 15%. The cost of capital uses the industry-average approach of the 2024 sheet (US Advertising).

## Inputs (every one has a `_sources` entry in `oct_2026_inputs.json`)

| Input | Value | Source |
|---|---|---|
| Revenue / operating income (TTM to 2026-06-30) | $228,247m / $86,926m | 10-K 2025 + 10-Q Q2 2026 |
| R&D (TTM; 3 prior years, 3-year life) | $71,634m; R&D-adjusted margin 49.1% | filings; life = Damodaran |
| Cash / debt incl. leases / equity investments | $90,260m / $112,318m / $30,157m | 10-Q balance sheet |
| Shares | 2,547.5m | 10-Q cover, 2026-07-24 |
| Price / riskfree | $728.08 / 5.28% (10-year Treasury) | StockAnalysis; Treasury, 2026-10-02 |
| Growth next year / years 2-5 | 23% / 12% | consensus anchor / **Damodaran** |
| Margin next year / target (year 5) | 46% / 40% | 2026 expense guidance / **Damodaran** |
| Sales-to-capital years 1-5 / 6-10 | 0.5 / 1.5 | TTM reinvestment (judgment) / **Damodaran** |
| Initial cost of capital | 8.51% | **Damodaran's method**: US Advertising industry average 7.81% + riskfree adjustment 0.70% |
| Terminal cost of capital / ROIC / growth | 9.48% / 15% / 5.28% | Ginzu default (riskfree + 4.20% mature ERP) / **Damodaran** / Ginzu default |
| Effective / marginal tax | 16% / 25% | company's 15-17% outlook / **Damodaran** |

## Scenarios (one change at a time; live in the workbook)

| Scenario | Value/share | vs price |
|---|---|---|
| Base case | 757.32 | +4.0% |
| Sales-to-capital 1.5 in years 1-5 (Damodaran Feb-2024) | 841.91 | +15.6% |
| Sales-to-capital 0.3 in years 1-5 (2026 capex pace) | 672.73 | −7.6% |
| Growth years 2-5 at 8% / 15% | 661.03 / 842.29 | −9.2% / +15.7% |
| Target margin 35% / 45% | 654.98 / 859.65 | −10.0% / +18.1% |
| Terminal ROIC = cost of capital (Ginzu default) | 588.70 | −19.1% |
| Next-year revenue = 4 × Q3 guidance midpoint ($250bn) | 674.73 | −7.3% |
| 10% chance of failure (youth-harm litigation, 50% proceeds) | 719.61 | −1.2% |

The cost-of-capital sensitivity runs from $570.27 to $1,069.55. The growth × margin table runs from $446.58 to
$1,359.26. The riskfree table runs from $915.45 (−1.5 points) to $619.43 (+1.5 points).

## What the price implies

The price sits close to value, so the breakevens are close to the base inputs:
- revenue growth of **10.9% a year in years 2-5** instead of 12%; or
- a **38.6% target margin** instead of 40%; or
- sales-to-capital of **0.41 in years 1-5** instead of 0.5, i.e. somewhat more AI spending per dollar of new revenue.

The valuation is most sensitive to whether the AI spending earns its keep. If the advantage fades and the terminal
ROIC falls to the cost of capital, value drops to $588.70.

## Ginzu Diagnostics, answered

| Diagnostic | Figure | Comment |
|---|---|---|
| Growth vs industry | 23% next year, 12% in years 2-5; industry 17.7%, last year 27.7% | Growth slows toward the 2024 story's 12% as the revenue base scales. |
| Dollar revenues | $228bn → $281bn → $442bn (year 5) → $648bn (year 10) | This requires keeping a dominant share of an ad market that keeps growing. |
| Margins | 49.1% now (R&D-adjusted) → 46% → 40%; industry 10.2% | AI depreciation (2026 capex $130-145bn) and rising R&D squeeze margins toward the 2024 target. |
| Reinvestment | Sales-to-capital 0.56 now, 0.5 in years 1-5, 1.5 in years 6-10; industry 3.85; ROIC 23% now → 21% in year 10 | Reinvestment absorbs 41% of the present value of operating income. The 6-10 rebound assumes the build-out ends. |
| Cost of capital | 8.51% → 9.48%; industry 7.8% | The industry-average approach puts the starting cost of capital below the Ginzu terminal default. |
| Value / price | 104% | Inside the Ginzu 0.5×–2× band: fairly priced. |

## Judgment calls and limits

- **Sales-to-capital of 0.5 in years 1-5** is the central judgment. It is anchored on trailing reinvestment and still
  understates 2026, when capex alone is guided to $130-145bn. The breakeven (0.41) and the scenarios bracket it.
- **Next-year growth** uses analysts' consensus (S&P Global via StockAnalysis): 2026 $254.2bn, 2027 $306.4bn. This
  is third-party data and a judgment.
- **Tax rate**: 16% follows the company's outlook. The trailing rate is distorted by a $15.93bn charge in Q3 2025
  (One Big Beautiful Bill Act) and an $8.03bn benefit in Q1 2026 (minimum-tax relief).
- **Cost of capital** follows the 2024 sheet's industry-average approach. A bottom-up beta reflecting the AI risk
  would raise it. The cost-of-capital table shows the effect: each +0.5 point on the initial rate costs about
  $20-23 a share.
- **Legal risk**: the 10-Q warns of youth-related trials that may result in material losses. This is not in the base
  case; scenario 9 prices a 10% chance of failure as an illustration.
- **Not included**: finance lease liabilities (in other liabilities), unvested RSUs and the small option pool.
  Dual-class control is flagged as a governance risk, not priced as a discount.

## Files (this folder)

| File | What |
|---|---|
| `oct_2026_inputs.json` | Inputs with a `_sources` entry for each, plus the workbook's scenarios |
| `valuation_output_2026-10-02.txt` | Engine output: year-by-year table, Ginzu Diagnostics, sales-to-capital breakeven |
| `META_Valuation_2026-10-02.xlsx` | Shareable workbook (Summary, Inputs, Scenarios, Sensitivity, WACC, DCF, Diagnostics, Checks, Sources) |

Reproduce: `python ../../skill/aswath-finance/scripts/ginzu.py oct_2026_inputs.json --breakeven sales_to_capital_1_5`;
workbook: `python ../../skill/aswath-finance/scripts/export_xlsx.py oct_2026_inputs.json -o META_Valuation_2026-10-02.xlsx --verify`.

**Sources.** SEC 10-Q Q2 2026 ([meta-20260630](https://www.sec.gov/Archives/edgar/data/1326801/000162828026050705/meta-20260630.htm)),
10-K 2025, Q2 2026 earnings release ([exhibit 99.1](https://www.sec.gov/Archives/edgar/data/1326801/000162828026050596/meta-06302026xexhibit991.htm)),
price, range, consensus and target ([StockAnalysis](https://stockanalysis.com/stocks/meta/forecast/), 2026-10-02),
10-year Treasury (home.treasury.gov daily par yield curve, 2026-10-02). Damodaran's model:
[Meta2024.xlsx](https://pages.stern.nyu.edu/~adamodar/pc/blog/Meta2024.xlsx). Damodaran's data: fcffsimpleginzu.xlsx
(Jan 2026) and ctrypremJuly26.xlsx (July 2026).
