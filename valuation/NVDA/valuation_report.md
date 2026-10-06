# NVDA (NVIDIA): valuation, Damodaran's way, as of 2026-10-02

**Value per share: $192.41 vs price $233.95.** The price is 122% of value (value is 82% of price). The model is Damodaran's
Ginzu FCFF spreadsheet, transcribed in the aswath-finance engine. The story is Damodaran's Nvidia story from September 2024
and January 2025, re-based to filings through 2026-07-26 and current market data. It is not investment advice.

Workbook: `NVDA_Valuation_2026-10-02.xlsx`. It has live formulas, a scenario selector, data tables and a Checks tab
tied to the engine (all checks pass in Excel).

## Damodaran's record on Nvidia

| Valuation | Value | Price then | Post |
|---|---|---|---|
| June 2023 | ~$240 | $409 | [AI's Winners, Losers and Wannabes](https://aswathdamodaran.blogspot.com/2023/06/ais-winners-losers-and-wannabes-nvidia.html), 2023-06 |
| September 2024 | ~$87 | $106 | [The Power of Expectations](https://aswathdamodaran.blogspot.com/2024/09/the-expectations-game-aftermath-of.html), 2024-09 |
| January 2025 (post-DeepSeek) | $77.51 (reproduced exactly) | $123 | [DeepSeek crashes the AI Party](https://aswathdamodaran.blogspot.com/2025/01/deepseek-crashes-ai-party-story-break.html), 2025-01 |
| **This update** | **$192.41** | **$233.95** | inputs below |

Damodaran found Nvidia overvalued each time while keeping part of a personal position (a quarter of the 2023 holding
as of January 2025). In December 2025 the post on trillion-dollar market caps used the market-cap breakeven tool
on Nvidia instead of a DCF, calling the revenue needed "possible and plausible", with the probable test running into competition
([Trillion Dollar Market Caps](https://aswathdamodaran.blogspot.com/2025/12/trillion-dollar-market-caps-fairy-tale.html), 2025-12).

## Why the story had to be re-based

The January 2025 model sized the whole **AI chip market at $300bn in 2035**, with Nvidia holding 60% of it. Nvidia's
data-center revenue for the twelve months to July 2026 is already **$277.8bn**. Total revenue is $303.0bn, up from
the $113.3bn base in that model, and FY2026 revenue ($215.9bn) beat its year-1 forecast ($141bn) by half. The
two-segment build (AI chips, auto chips) cannot be rolled forward without new market sizes. This update keeps the
story's levers and expresses them as standard Ginzu drivers, so every input stays live in the workbook:

- **Kept from the January 2025 model:** Nvidia's share falls to 60%. The target margin falls to 60% ("still software-like") as
  competition, customers and TSMC push for a share. The 20% terminal ROIC (lasting advantages), sales-to-capital of
  2.5, 3-year reinvestment lag, 25% marginal tax and A2/A rating are also kept.
- **Updated:** next year follows the latest information. Analysts' consensus puts the next twelve months at about
  $548bn (+81%), and the company guides Q3 FY2027 to $108bn. After that, AI chip spending has to be paid for by AI
  product and service revenue, which Damodaran puts at roughly $250bn today
  ([AI's Bar Mitzvah Moment](https://aswathdamodaran.blogspot.com/2026/08/ais-bar-mitzvah-moment-from-hype-hope.html), 2026-08).
  So growth slows to 5% a year in years 2-5 and converges to the riskfree rate. Year-10 revenue of $858bn is a 60%
  share of a roughly $1.35trn AI data-center market, plus edge computing.

## Inputs (every one has a `_sources` entry in `oct_2026_inputs.json`)

| Input | Value | Source |
|---|---|---|
| Revenue / operating income (TTM to 2026-07-26) | $302,970m / $197,579m | 10-K FY2026 + 10-Q Q2 FY2027 |
| R&D (TTM; 5 prior years, 5-year life) | $23,592m; R&D-adjusted margin 69.5% | filings; life = Damodaran |
| Cash / debt incl. leases / equity stakes | $56,586m / $38,860m / $93,940m | 10-Q balance sheet |
| Shares | 24,147m | 10-Q equity statement |
| Price / riskfree | $233.95 / 5.28% (10-year Treasury) | StockAnalysis; Treasury, 2026-10-02 |
| Growth next year / years 2-5 | 81% / 5% | judgment (consensus anchor / Damodaran's story) |
| Margin next year / target (year 5) | 68% / 60% | Q3 guidance / **Damodaran** |
| Sales-to-capital, reinvestment lag | 2.5, 3 years | **Damodaran** |
| Initial cost of capital | 12.19% | **Damodaran's method**: US semiconductor beta 1.50 (Jan-2026 data), ERP 4.60% from the July-2026 country ERPs weighted by revenue by customer HQ, A2/A |
| Terminal cost of capital / ROIC / growth | 9.08% / 20% / 5.28% | Damodaran's 8.50% moved with the riskfree rate / **Damodaran** / Ginzu default (= riskfree) |
| Effective / marginal tax | 16.05% / 25% | filings / **Damodaran** |

## Scenarios (one change at a time; live in the workbook)

| Scenario | Value/share | vs price |
|---|---|---|
| Base case | 192.41 | −17.8% |
| Next-year revenue = 4 × guided Q3 ($432bn) | 152.56 | −34.8% |
| Growth years 2-5 at 0% (AI capex digestion) | 154.72 | −33.9% |
| Growth years 2-5 at 11.2% (Damodaran's US semiconductor 5-year average) | 252.40 | +7.9% |
| Target margin 50% / 70% | 164.52 / 220.29 | −29.7% / −5.8% |
| Sales-to-capital 1.21 (Damodaran's US semiconductor average) | 188.63 | −19.4% |
| Terminal cost of capital 9.48% (Ginzu default: riskfree + mature ERP) | 180.48 | −22.9% |
| Terminal ROIC = cost of capital (Ginzu default, no lasting moat) | 144.01 | −38.4% |
| Riskfree at the Jan-2025 4.70% (terminal 8.50%) | 202.32 | −13.5% |
| Reinvestment lag 1 year (Ginzu default) | 192.78 | −17.6% |

The cost-of-capital sensitivity runs from $150.78 to $264.31. The growth × margin table runs from $110.40 to
$338.52. The riskfree table runs from $219.04 (−1.5 points) to $168.80 (+1.5 points).

## What the price implies

- **DCF breakeven.** Revenue growth of **9.5% a year in years 2-5** instead of 5%: revenue of $787bn in year 5 and
  $1.10trn in year 10, i.e. a 60% share of an AI market of about $1.75trn. Alternatively, a **75% target margin**
  instead of 60%, with everything else unchanged.
- **Damodaran's breakeven tool** (`breakeven_marketcap.json`), using the December-2025 Nvidia inputs: 53.01% net margin,
  64.44% ROE and a 9% cost of equity. At today's $5.65trn market cap, Nvidia needs **$546bn of revenue if it were
  mature today** and **$766bn after a five-year wait** (2% cash yield). The base case reaches $667bn in year 5.

## Ginzu Diagnostics, answered

| Diagnostic | Figure | Comment |
|---|---|---|
| Growth vs industry | 81% next year, then 5%; industry 11.2%, last year 83.4% | Year 1 follows consensus. After it, Nvidia is no longer "small", so growth falls below the industry average. |
| Dollar revenues | $303bn → $548bn → $667bn (year 5) → $858bn (year 10) | Year 10 = 60% of a ~$1.35trn AI data-center market. The price needs ~$1.10trn. |
| Margins | 69.5% now → 68% → 60% (year 5); industry 36.5% | The 60% target is still well above the chip industry. |
| Reinvestment | Sales-to-capital 2.5 vs 1.15 now (1.8 excluding equity stakes) and 1.21 for the industry; year-10 ROIC 96% | This is Damodaran's input; the January-2025 model also had year-10 ROIC above 100%. The industry ratio costs only $3.78 a share (3-year lag, slow growth). |
| Cost of capital | 12.19% → 9.08%; industry 10.6% | A higher semiconductor beta (1.50) and a riskfree rate up 0.58 points since the January-2025 sheet. |
| Value / price | 82% | Inside the Ginzu 0.5×–2× band. The gap comes from the story (growth after year 1), not from a modelling flag. |

## Judgment calls and limits

- **Next-year growth** uses analysts' consensus (S&P Global via StockAnalysis): FY2027 $411.6bn, FY2028 $683.3bn.
  This is third-party data and a judgment, not Damodaran's input. The guidance-only view (scenario 2) gives $152.56.
- **Years 2-5 growth (5%)** is the central judgment. It is the most powerful lever after the terminal ROIC. The
  breakeven (9.5%) and the scenarios bracket it.
- **Equity stakes of $93.9bn** (marketable $42.8bn, non-marketable $51.2bn, partly carried at cost plus upward
  adjustments) are added as cross holdings at balance-sheet value, about $3.89 a share.
- **ERP weights** use Nvidia's FY2026 revenue by customer headquarters. The 10-K estimates that 76% of data-center
  revenue from Taiwan-headquartered customers goes to US and European end customers. Moving that share to the US
  lowers the ERP to 4.53% and raises value by $1.15.
- **Rating**: A2/A from the January-2025 sheet. Debt is 0.6% of capital, so the rating barely matters.
- **Shares**: basic shares outstanding, as in the Nvidia sheets. Unvested RSUs are not counted, and no options are
  valued (the published model's choice).

## Files (this folder)

| File | What |
|---|---|
| `oct_2026_inputs.json` | Inputs with a `_sources` entry for each, plus the workbook's scenarios |
| `valuation_output_2026-10-02.txt` | Engine output: year-by-year table, Ginzu Diagnostics, growth breakeven |
| `breakeven_marketcap.json` / `_output.txt` | Damodaran's market-cap breakeven tool at $5.65trn |
| `NVDA_Valuation_2026-10-02.xlsx` | Shareable workbook (Summary, Inputs, Scenarios, Sensitivity, WACC, DCF, Diagnostics, Checks, Sources) |

Reproduce: `python ../../skill/aswath-finance/scripts/ginzu.py oct_2026_inputs.json --breakeven growth_years_2_5`;
workbook: `python ../../skill/aswath-finance/scripts/export_xlsx.py oct_2026_inputs.json -o NVDA_Valuation_2026-10-02.xlsx --verify`.

**Sources.** SEC 10-Q Q2 FY2027 ([nvda-20260726](https://www.sec.gov/Archives/edgar/data/1045810/000104581026000075/nvda-20260726.htm)),
10-K FY2026 ([nvda-20260125](https://www.sec.gov/Archives/edgar/data/1045810/000104581026000021/nvda-20260125.htm)),
Q2 FY2027 earnings release ([q2fy27pr](https://www.sec.gov/Archives/edgar/data/1045810/000104581026000073/q2fy27pr.htm)),
price, range, consensus and target ([StockAnalysis](https://stockanalysis.com/stocks/nvda/forecast/), 2026-10-02),
10-year Treasury (home.treasury.gov daily par yield curve, 2026-10-02). Damodaran's model:
[NvidiaJan2025.xlsx](https://pages.stern.nyu.edu/~adamodar/pc/blog/NvidiaJan2025.xlsx). Damodaran's data: fcffsimpleginzu.xlsx
(Jan 2026) and ctrypremJuly26.xlsx (July 2026).
