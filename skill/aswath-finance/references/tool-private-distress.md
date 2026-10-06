# Tool: private_distress.py — private-company and distress adjustments, Damodaran's way

`scripts/private_distress.py` answers, from Damodaran's own calculators (stdlib only):
- **Illiquidity discount** for a private firm (restricted-stock regression; bid-ask spread regression)
- **Minority vs control**: value of control, minority discount, expected value given a probability of change
- **Cost of capital for an undiversified owner**: total beta (or the sheet's "private company premium" alternative)
- **Probability of failure/distress** from a bond price (distress.xls) or from a rating / firm age (Ginzu table)
- **Value of equity in a distressed / money-losing firm**: fcffneg.xls 10-year FCFF model with NOL shield, or the
  Ginzu going-concern value weighted with a failure probability pulled from the template's table (`--ginzu`)
- **Normalized earnings** (normearn.xls) and the **complexity score** (complscore.xls)

## Sources
| File (pages.stern.nyu.edu/~adamodar/pc/...) | Sheet / cells transcribed |
|---|---|
| liqdisc.xls | 'Restricted Stock Regression'!B8, 'Bid-Ask Spread'!B9, Sheet2 (the discount table) |
| minoritydiscount.xls | Sheet1 B7:B9 |
| pvtdiscrate.xls | 'Total Beta approach'!B21:D26, 'Synthetic ratings'!D7:D10 + its 15-row table |
| distress.xls | Sheet1 rows 8-38 (the sheet solves p with Solver; we bisect) |
| fcffsimpleginzu.xlsx | 'Failure Rate worksheet' A5:K12 (rating x 1-10 yrs), O18:X28 (BLS 2022 survival by age) |
| normearn.xls | 'Earnings Normalizer & Output'!D2:D14, G19:G21 |
| complscore.xls | Sheet1 D3:E25 |
| fcffneg.xls | 'Valuation' rows 21-140, 'Option Valuation' (dilution-adjusted Black-Scholes) |

Tables live in `data/private_distress_data.json` (failure by rating and age: Ginzu template Jan 2026, identical in
all 19 of Damodaran's 2023-2026 valuation files that carry the sheet; pvtdiscrate's own synthetic-rating table).

Posts where Damodaran uses these ideas:
- [Valuing declining and distressed companies](https://aswathdamodaran.blogspot.com/2009/06/valuing-declining-and-distressed.html), 2009-06 — adjust DCF value for the likelihood of failure; bond prices as a clue.
- [Valuing Young Companies](https://aswathdamodaran.blogspot.com/2009/06/valuing-young-companies.html), 2009-06 — BLS survival data (Knaup & Piazza) as failure rates.
- [Asset selection & Valuation in Illiquid Markets](https://aswathdamodaran.blogspot.com/2010/12/asset-selection-valuation-in-illiquid.html), 2010-12 — illiquidity discount (restricted stock, bid-ask) vs a discount-rate premium.
- [Stay Private vs Going Public](https://aswathdamodaran.blogspot.com/2011/07/stay-private-vs-going-public-changing.html), 2011-07 — total beta: market beta 0.8 -> 2.4 for an undiversified owner.
- [Activist Value Investing](https://aswathdamodaran.blogspot.com/2012/06/activist-value-investing-be-your-own.html), 2012-06 and [Alibaba's Governance by Politburo](https://aswathdamodaran.blogspot.com/2014/09/alibabas-governance-by-politburo.html), 2014-09 — expected value = optimal x p(change) + status quo x (1 - p).
- [The Small Cap Premium: Where is the beef?](https://aswathdamodaran.blogspot.com/2015/04/the-small-cap-premium-fact-fiction-and.html), 2015-04 — never add a small-cap premium *and* an illiquidity discount; Damodaran uses no small-cap premium.
- [Myth 4.1](https://aswathdamodaran.blogspot.com/2016/11/myth-41-if-you-don-like-betas-or-modern.html), 2016-11 — use a total-risk measure if the marginal investor is not diversified.
- [Runaway Story or Meltdown in Motion? WeWork](https://aswathdamodaran.blogspot.com/2019/09/runaway-story-or-meltdown-in-motion.html), 2019-09 — 20% failure, proceeds 60% of fair value.
- [Dealing with Aging: Intel, Walgreens, Starbucks](https://aswathdamodaran.blogspot.com/2024/09/dealing-with-aging-daignosing-intel.html), 2024-09 — Walgreens (Ba2/BB): failure probability 11.78% = the table's BB 10-year value.

## Calibration (`python scripts/private_distress.py --selftest` -> 86/86)
| File | Sheet value | Ours |
|---|---|---|
| liqdisc.xls restricted stock ($209m rev, profitable) | 0.190955 | 0.190955 |
| liqdisc.xls bid-ask (cash/value 3%) | 0.117767 | 0.117767 |
| liqdisc.xls Sheet2 table, 16 revenue levels x profitable/loss | 0.1579 ... 0.3421 (4 dp) | 32/32 |
| minoritydiscount.xls discount / majority stake | 14.966% / 7,497 | 14.966% / 7,497 |
| pvtdiscrate.xls total beta / cost of equity / WACC | 2.5067 / 19.787% / 17.449% | same |
| distress.xls annual / 5-yr / 10-yr probability (price 653) | 13.532% / 51.662% / 76.635% | same |
| fcffsimpleginzu.xlsx BB 10-yr; Information age 6 (the sheet's worked examples) + 5 age cells | 11.78%; 11.70% | same |
| normearn.xls normalized EBIT (sector margin) | 1,788.67 | 1,788.67 |
| complscore.xls score | 51.5 | 51.5 |
| fcffneg.xls value/share, firm value, options | 11.0818, 15,917.3, 299.73 | same (sheet slips replicated) |
| Walgreens2024.xlsx Ginzu value with failure = BB table value | 7.5268 | 7.5268 |

Pycel branch tests (the template recalculated under inputs the source files never use), all matching: liqdisc (loss-making,
10% block, cash/volume), minority (30% stake), pvtdiscrate (premium "P", synthetic rating, coverage 1.1 + target debt
ratio, correlation 0.25), distress (other bond), normearn (approaches 1 and 2, changed history), complscore, fcffneg
(direct cost of equity, beta unchanged, debt ratio change, capex = 110% of depreciation, private book/target debt
ratio, no options, no NOL, year-1 EBIT > NOL, other stable growth/ROC) and 3 runs on a copy with the two slips fixed.

**Could not match as written**: minoritydiscount.xls B8 "value of minority stake" = `B2+B5*B3` (20,825, more than the
whole firm) is a typo; we return minority stake x status quo value (`"replicate_sheet_bugs": true` gives the sheet's B8).
fcffneg.xls has two slips — years 8-10 use the
prior year's EBITDA margin (J100:L100 point one row up) and year-1 tax compares the text label "NOL" (always 0).
Default is the corrected logic; `"replicate_sheet_bugs": true` reproduces the cached 11.08 (corrected: 11.14).
In both tools the report prints a `sheet_error` line with the as-written figure whenever it differs.
Adani2023 (10%) and InstacartIPO (5%) failure probabilities are Damodaran's judgment, not table values (both rated A2/A = 1.24%).

## Spreadsheet defaults and when Damodaran overrides them
- Restricted-stock discount: base 25% for a profitable $10m-revenue firm (B8 hard-codes 0.25; the B2 input is unlinked),
  block = 100%, `profitable` = 1/0 dummy. Bid-ask: cash/value and volume/value default 0.
- pvtdiscrate: total beta "T" = unlevered beta / correlation x (1 + (1-t) D/E); correlation 0.45 in the source file;
  "P" = relevered market beta x ERP + rf + premium (10%). Synthetic rating uses **the sheet's own old spread table**
  (A3/A- 0.95% ...); for today's spreads use `cost_of_capital.py` and pass `pre_tax_cost_of_debt`.
- Failure: 10-year horizon (the sheet's worked example); broad grades only (`Ba2/BB`, `BB+` -> BB; CCC, CC, C -> CCC/C).
  Age approach = S(age) - S(10) from BLS survival (age 0 = 1 - S(10)). Damodaran overrides with judgment (WeWork 20%, Adani 10%).
- Ginzu proceeds on failure stay a Ginzu input (`failure_proceeds_basis` B/V, `failure_proceeds_pct`).
- normearn approach 3 default sector margin = sum of 5-yr EBIT / sum of 5-yr revenues; approach 1 average EBIT is a
  typed input in the sheet (we default to the history mean only when it is omitted).
- fcffneg: 10-year horizon only (the template errors for shorter periods); EBITDA margin years 6-10 move linearly to the
  stable margin; capex growth yrs 9-10 = yr-10 revenue growth; depreciation growth yr t (7-10) = capex growth yr t-1;
  beta, debt ratio and pre-tax cost of debt move to their stable values over years 6-10 (the cost of debt path moves
  even when "change in stable period" is No — the sheet's H118:L118); stable reinvestment = g / ROC x EBIT(1-t).

## Input reference (JSON sections; combine several in one file)
`illiquidity {revenues, block, profitable, base_discount, cash_to_value, volume_to_value}` ·
`minority {optimal_value, status_quo_value, majority_stake, minority_stake, probability_of_change,
replicate_sheet_bugs}` ·
`private_coc {unlevered_beta, approach T|P, correlation, premium, riskfree_rate, erp, tax_rate, debt_to_capital,
pre_tax_cost_of_debt | synthetic {ebit, interest_expense}}` ·
`bond_distress {coupon_rate, maturity, riskfree_rate, price, face=1000}` ·
`failure {rating, years} | {sector, age}` (sectors: Agriculture, Mining, Utilities, Construction, Manufacturing,
Retail, Transportation, Information, Health Care, All Sectors) ·
`normalized_ebit {approach 1|2|3, average_ebit, average_roc, sector_margin, revenues, book_debt, book_equity, history}` ·
`complexity {businesses, one_time_pct, ..., off_balance_sheet_debt}` (all 21 answers, see `COMPLEXITY`) ·
`fcffneg {...}` (see the function docstring; `--selftest` case inputs are a full example).

## Example (Damodaran's own inputs)
`in.json`:
```json
{"private_coc": {"unlevered_beta": 1.02, "approach": "T", "correlation": 0.45, "riskfree_rate": 0.06,
                 "erp": 0.055, "tax_rate": 0.4, "debt_to_capital": 0.15, "pre_tax_cost_of_debt": 0.07},
 "illiquidity": {"revenues": 209, "profitable": true, "cash_to_value": 0.03},
 "bond_distress": {"coupon_rate": 0.12, "maturity": 8, "riskfree_rate": 0.05, "price": 653},
 "failure": {"rating": "Ba2/BB", "years": 10}}
```
`python scripts/private_distress.py in.json` -> total beta 2.507, cost of capital 17.45%; illiquidity discount 19.1%
(bid-ask 11.8%); distress 13.5%/yr (51.7% over 5 yrs); failure 11.78%.
Ginzu with a table failure rate: `python scripts/ginzu.py --from-case Walgreens2024 > w.json`, add
`"failure": {"rating": "Ba2/BB"}`, then `python scripts/private_distress.py --ginzu w.json` -> $7.53/share.

## Limits
- The regressions (Silber restricted stock; bid-ask spread) are Damodaran's old estimates, US data; the discount applies to
  the value after a total-beta/market-beta DCF — do not stack it with a small-cap or illiquidity premium (2015-04).
- distress.xls assumes a constant annual default probability, zero recovery and riskfree discounting; the sheet
  truncates at 25 years (we do not).
- The failure table is cumulative default rates for rated bonds (dates as of the Ginzu 2026 template); BLS survival
  is for all establishments, not venture-backed firms. Damodaran's "factors to consider" (size, access to capital, growth, unit
  economics) are judgment, not formulas.
- Rebuild: `.venv\Scripts\python.exe scripts\private_distress.py --build-cases <files>` (see the script docstring);
  pycel needs Excel "Save As .xlsx" copies of the .xls tools (kept in `build\tmp\private_distress\`).
