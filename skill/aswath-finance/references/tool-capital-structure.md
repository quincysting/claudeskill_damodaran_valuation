# Tool: capital_structure.py (optimal debt ratio, buybacks and cash return, done Damodaran's way)

**What it answers.** (1) What is a company's optimal debt ratio under Damodaran's cost-of-capital approach? At each debt ratio
from 0% to 90% it re-levers beta, rates the debt synthetically from interest coverage, and computes the cost of debt,
the WACC and the firm value. (2) How does value change if the company moves to that ratio: EV, value per share, and
value per remaining share after a debt-funded buyback? (3) What does a buyback do to beta, WACC, price, EPS and PE, and
how much value moves between selling and remaining shareholders? (4) How much cash can the firm return: FCFE before
and after debt cash flows, compared with dividends, buybacks and net cash returned? It also relevers betas (levbeta)
and gives a synthetic rating. Stdlib only. Every formula comes from the source spreadsheets.

## Sources and posts
| Source file (URL `pages.stern.nyu.edu/~adamodar/pc/<file>`) | What was transcribed | Post |
|---|---|---|
| capstru.xlsx (template: Disney inputs dated 2023-02-26; Jan 2026 rating and country tables) | 'Inputs', 'Optimal CS Worksheet' rows 4-97, 'Operating leases', 'Default Spreads and Ratios', 'Repurchase price Worksheet' | [Data Update 6 for 2023: A Wake up call for the Indebted?](https://aswathdamodaran.blogspot.com/2023/02/data-update-6-for-2023-wake-up-call-for.html), 2023-02; [Data Update 8 for 2025: Debt, Taxes and Default](https://aswathdamodaran.blogspot.com/2025/02/data-update-8-for-2025-debt-taxes-and.html), 2025-02; [Data Update 7 for 2026: Debt and Taxes](https://aswathdamodaran.blogspot.com/2026/02/data-update-7-for-2026-debt-and-taxes.html), 2026-02 |
| blog/AdaniCapStru.xlsx, blog/AdaniGroupCapStru.xlsx | Damodaran's applied cases (older layout, same formulas) | 2023-02 post above: Adani Enterprises "413,443 million, more than double its optimal debt of 185,309 million"; the group carries "about three times as much debt as it should" |
| buybacks.xls ('Value Effect', 'EPS Effect') | debt-funded buyback: relevered beta, WACC, EV, price, value transfer, EPS | [Stock Buybacks: They are big, they are back and they scare some people!](https://aswathdamodaran.blogspot.com/2014/09/stock-buybacks-they-are-big-they-are.html), 2014-09 |
| buybackcalculator.xlsx | EPS/PE effect when you supply the post-buyback price | [Data Update 8 for 2026: Time for Harvesting](https://aswathdamodaran.blogspot.com/2026/02/data-update-8-for-2026-time-for.html), 2026-02 |
| levbeta.xls, ratings.xls | Hamada relevering table; coverage → rating → cost of debt | ratings.xls: [Putting the D in the DCF: The Cost of Capital](https://aswathdamodaran.blogspot.com/2015/01/putting-d-in-dcf-cost-of-capital.html), 2015-01 (levbeta.xls is not linked from a post) |
| datasets/divfcfe*.xls ('Variables & FAQ') | Damodaran's FCFE definitions and cash-return ratios with the dataset's 'NA' rules | [Data Update 9 for 2025: Dividends and Buybacks](https://aswathdamodaran.blogspot.com/2025/03/data-update-9-for-2025-dividends-and.html), 2025-03; 2026-02 post above |

Earlier versions linked from the posts: capstru.xls in [January 2017 Data Update 8: The Dark and Light Sides of Debt](https://aswathdamodaran.blogspot.com/2017/02/january-2017-data-update-8-dark-and.html) (2017-02),
and capstru2018.xlsx, which added the interest-deduction cap, in [January 2018 Data Update 8: Debt and Taxes](https://aswathdamodaran.blogspot.com/2018/01/january-2018-data-update-8-debt-and.html) (2018-01).
For Apple in [Apple: News, Noise and Value](https://aswathdamodaran.blogspot.com/2013/04/apple-calm-after-storm.html) (2013-04) Damodaran re-ran
it with operating income halved. apv.xls (Hormel 2009) is a different method: APV with default probabilities and
bankruptcy costs. It is not transcribed. wacccalc.xls (today's cost of capital) is not transcribed either;
use `cost_of_capital.py` for that.

## Damodaran's method (worksheet → code)
1. **Today.** Market value of debt is given (B24/B25), or estimated as a bond from interest, book debt and maturity
   (B26/B27), or taken at book. Operating leases are PV'd at the pre-tax cost of debt; years past 5 run as an annuity
   over round(beyond / avg yr1-5) years, and EBITDA, depreciation and interest are adjusted. WACC uses the stated
   pre-tax cost of debt, or the synthetic one (B44). The implied growth rate is g = (EV·WACC − FCFF)/(EV + FCFF), capped at the riskfree rate.
2. **Each debt ratio r (0%..90% of current D+E).** β = βu(1+(1−t)·r/(1−r)). Interest = rate × debt if refinance (B43).
   Otherwise, below today's ratio, interest = (interest/book debt) × debt; above it, interest = today's interest +
   rate × (debt − MV of debt). Coverage = EBIT/interest gives the rating and pre-tax rate (rf + spread + country spread),
   and the rate feeds back into interest (circular, Excel iteration). Tax on interest = min(t, t·EBIT/interest, the
   cap-limited rate if interest deductions are capped at X% of EBITDA/EBIT).
3. **Value.** EV_r = EV_now × (1 + (WACC_now − WACC_r)/(WACC_r − g)). The optimum is the highest EV_r, which is
   not always the lowest WACC. With indirect bankruptcy costs on (B34), EBITDA drops with the rating (the sheet's 'Drop in
   EBITDA' column) and value = (EBIT(1−t) − (capex − dep))(1+g_implied)/(WACC_r − g_implied).
4. **Recap** ('Repurchase price Worksheet'): new debt / buyback price = shares bought; value per remaining share =
   (EV_opt + cash − debt_opt)/shares left.
5. **Buyback value effect:** debt rises by (cost − cash used). D/E after = (D + new debt)/(market cap − cost). The new
   debt's rate applies to all debt. EV' = EV + EV(WACC − WACC')/(WACC' − rf). Firm value = EV' + cash − cash used.
   Value transfer = (fair value − buyback price) × shares bought / shares left. **EPS:** NI' = NI − interest income on
   cash − new debt × rate × (1−t).
6. **FCFE** = NI + D&A − cap ex (incl. R&D and acquisitions) [− Δ working capital, per the 2026 post] + new debt −
   repayments. Ratios are 'NA' when the denominator is ≤ 0, except net cash returned / NI.

## Calibration (`python scripts/capital_structure.py --selftest` → 41/41, every cell within 0.05% / 1e-6)
| File | Sheet value | Ours |
|---|---|---|
| capstru.xlsx, Disney (185 cells: 10 debt ratios × 14 rows + summary + recap) | optimal 20%, WACC 9.5212%, EV 218,233.41, $101.27/sh | same |
| AdaniCapStru.xlsx (182 cells) | optimal 10%, debt 185,308 vs 413,443, EV 1,868,031, ₹1,359.97/sh | same |
| AdaniGroupCapStru.xlsx (182 cells; shares = 1, so value/share = equity value) | optimal 10%, EV 8,713,021 | same |
| capstru template recalculated **in Excel** under 16 overrides the source files never use: firm types 2/3, indirect bankruptcy costs, interest cap on EBITDA and on EBIT, refinance, synthetic current rating, leases (± synthetic), given MV of debt, book debt, operating income −40%, and 4 combinations | 16 × 185 cells | same |
| buybacks.xls | price 60 → 61.658, EPS 2.000 → 2.0472, value transfer 1.111 | same |
| buybackcalculator.xlsx (interest income left to our default = cash × 0.5%, the sheet's formula) | EPS 4.000 → 4.1579, PE 12.03 | same |
| levbeta.xls / ratings.xls | current β 1.4811; coverage 6.25 → A2/A, 3.78% | same |
| divfcfe (US, Global, Europe, Japan, China, India, Emerging) totals + 4 US 'NA' industries (18 cases) | e.g. US net cash returned/FCFE 1.0061 pre-debt, 0.7109 post-debt | same |

**Circularity; read this before quoting a rating.** At some debt ratios more than one rating is self-consistent: the rate
implies coverage that maps back to the same rating. Excel keeps whichever one its iteration history reached. The
engine lists them all (`self_consistent_ratings`, shown in the report) and defaults to the worst one, which is what
the template shows. Damodaran's Adani Enterprises file (10%: AA, not A+) and group file (20%, 70%) sit on other
self-consistent ratings, and so do 7 of the 16 Excel branches. Those cases pin the column with
`ratings_at` (the engine refuses a rating that is not self-consistent). The optimum matches the sheet's in all 19 files and
branches except one (IBC + type 2 + refinance: 0% by default, 10% at the sheet's rating). When a column is ambiguous the report
also prints `optimal_if_best_ratings`. Pycel cannot evaluate this sheet (its loop cells start at 0, and #DIV/0! is
absorbing), so the branches were recalculated by Excel 16 via COM (`build/tmp/capital_structure/branches.ps1`, which
saves copies and never saves the source file) and read like any of the source files. Rebuild with
`--build-cases spreadsheets/capstru.xlsx spreadsheets/Adani*CapStru.xlsx build/tmp/capital_structure/branches/*.xlsx
spreadsheets/buybackcalculator.xlsx spreadsheets/tools/{buybacks,levbeta,ratings}.xls spreadsheets/divfcfe*.xls`.

## Spreadsheet defaults, and when Damodaran overrides them
- Firm type 1 (large/safe coverage table). Damodaran switches to type 3 (infrastructure) for both Adani files.
- Existing debt keeps its rate (B43 No). The Adani files set **refinance = Yes**. Current cost of debt as stated (B44 No).
- No interest-deduction cap. This is the 2017 US tax-law option, "30% of EBITDA until 2022, EBIT after".
- No indirect bankruptcy costs. The Low/Medium/High magnitude input (B35) is **not wired**: the table is always
  'Medium'.
- Market data: the Adani files look up ERP, default spread and tax rate for the country. The template types them, and
  its 1.81% country default spread for a US company is a typed value; the US figure in Damodaran's Jan 2026 table is
  0.23%. When those are missing, the engine fills
  them from the Jan 2026 country table (`data/ginzu_data.json`, same as the template's).
- **Quirks kept on purpose:** (a) for type 3 the worksheet uses the utility coverage bounds with the type-2 spreads
  (it never reads the utility spreads); (b) EBITDA is grossed up for today's rating drop (F45) even with indirect
  costs off; (c) the 'Summary Table' sheet shows the EBIT-limit tax rate (row 68), not row 69. They differ only
  when the interest cap is on. (d) The Adani files carry the template's (Disney) lease table and
  buyback price (67.71), so their recap line (negative shares bought, ~₹385/share) is not meaningful; set your own
  `buyback_price` or drop it.

## Inputs (JSON; any subset of sections)
`optimal`: ebitda, depreciation, capex, interest_expense, book_debt, cash, shares_outstanding, price, beta,
riskfree_rate, pre_tax_cost_of_debt, [country → marginal_tax_rate, erp, country_default_spread], market_value_debt |
debt_maturity, leases {current_expense, commitments[5], beyond}, interest_limit {measure, max_pct},
indirect_bankruptcy_costs, firm_type, refinance, synthetic_current_rating, buyback_price, rating_tables, ratings_at.
`buyback`: price, shares_outstanding, shares_bought, buyback_price, cash, cash_used, net_income, new_debt_rate,
tax_rate, [interest_income_from_cash], then either post_buyback_price, or beta, debt, cost_of_debt, riskfree_rate, erp
[, fair_value]. `cash_return`: net_income, dividends, buybacks, stock_issuances, depreciation, capex,
change_in_working_capital, new_debt, debt_repaid (or fcfe_pre_debt / fcfe_post_debt). `levbeta`: beta, tax,
de_regression [, market_value_equity, book_debt, debt_maturity, interest_expense, cost_of_debt, de_target].

## Example (Damodaran's Adani Enterprises inputs)
```
python scripts/capital_structure.py --from-case AdaniCapStru > adani.json
python scripts/capital_structure.py adani.json          # add --json for machine output
```
→ current 22.3% debt, WACC 12.00%; optimum 10% (debt 185,308), WACC 11.78%, EV +56,061, ₹1,353.71/sh at the default
(worst) rating. At the file's cached AA rating it is ₹1,359.97; both are shown.

## Limits
- It uses 10% steps and a single year of EBITDA. Damodaran's FAQ says to use normalized income if the latest year is abnormal.
  Value assumes perpetual growth at g, and working capital is ignored in FCFF, as in the source sheet.
- The rating comes from coverage alone. Damodaran warns (FAQ Q5) that this can differ from the actual rating. For banks and
  other financials use the financial-firm table (ratings.xls) via `rating_tables`; capstru is for non-financial firms.
- Buyback EPS subtracts all interest income on cash, pre-tax, as the source sheets do.
- divfcfe's 'Payout' column is not dividends/NI for 5-12 industries in each file checked (US, Global, India), so only the
  four ratios above are calibrated.
