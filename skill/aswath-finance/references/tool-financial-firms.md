# Tool: valuing banks, insurers and other financial-service firms (`scripts/financial_firms.py`)

`ginzu.py` (FCFF) cannot value a bank: deposits and debt are raw material, not capital, so Damodaran values only the
**equity**, with cash flows to equity and a cost of equity. This engine transcribes the source equity models cell by
cell. Use it for: "value bank X Damodaran's way" (`bank_fcfe`), "what is its excess return on equity worth?"
(`excess_return`), "what does a DDM / FCFE model give?" (`ddm`, `fcfe`, stage models).

## Damodaran's framework (from the posts)
- Pre-2008 Damodaran valued banks with dividend discount models (trusting managers and regulators); after 2008 the approach is a
  bank FCFE: **FCFE = net income - change in regulatory (Tier 1) capital needed to grow**. It can go negative (new
  equity has to be raised, i.e. dilution). A simpler route: DDM with payout = 1 - g/ROE
  ([Breach of Trust: Bank Valuation after the banking crisis](https://aswathdamodaran.blogspot.com/2011/03/breach-of-trust-bank-valuation-after.html), 2011-03).
- Deutsche Bank: 10-year FCFE with a Tier 1 ratio rebuilt to peer levels, ROE recovering, a 10% "complete guess"
  probability of equity wipeout; $20.67 vs $13.33 price
  ([Deutsche Bank: A Greek Tragedy at a German Institution?](https://aswathdamodaran.blogspot.com/2016/10/deutsche-bank-greek-tragedy-at-german.html), 2016-10).
- Citi: 3% risk-adjusted asset growth, ROE 8.78% -> 9.50% (5-yr average), Tier 1 14.80% -> 15%, cost of equity
  11.67% (implied for the 25 largest banks): $68.58 vs $46.32, and Damodaran bought it
  ([Good (Bad) Banks and Good (Bad) Investments](https://aswathdamodaran.blogspot.com/2023/05/good-bad-banks-and-good-bad-investments.html), 2023-05).
  Good vs bad banks and failure risk (deposit stickiness, SVB):
  [Breach of Trust: Decoding the Banking Crisis](https://aswathdamodaran.blogspot.com/2023/05/breach-of-trust-decoding-banking-crisis.html), 2023-05.
- **Bank cost of equity, Damodaran's way**: implied from a stable DDM on peer banks, ke = g + (ROE - g)/(P/BV). With
  g 1.60%, P/BV 0.69/1.06/1.54 and ROE 5.85/9.91/14.84% (25th/median/75th pct of 616 global banks) this gives
  the 7.76/9.44/10.20% in DeutscheBank2016 Inputs!D39:F39 (2016-10 post). Same relation: P/BV = (ROE-g)/(ke-g).
- DDM is too restrictive for most firms, but fits a committed payer (Aramco) and fails when firms pay more than
  they can afford ([Aramco's long awaited IPO](https://aswathdamodaran.blogspot.com/2019/11/a-coming-out-party-for-worlds-most.html), 2019-11).
- Insurers / other financials: no insurer-specific file exists among Damodaran's tools; use the equity models
  (`ddm`, `fcfe`, `excess_return`).

## Models and sources
| model | source file (URL: pages.stern.nyu.edu/~adamodar/pc/...) | what it does |
|---|---|---|
| `bank_fcfe` | blog/CitiVal2023.xlsx; blog/DeutscheBank2016.xls (`template`) | NI - change in Tier 1 capital; ROE/Tier-1 paths; payout g/ROE in perpetuity; failure prob. |
| `excess_return` | eqexret.xls | book equity + PV of (ROE - ke) x beginning book equity |
| `ddm` | divginzu.xls (ships with Damodaran's JPMorgan May-2024 inputs) | 1-15 yr dividends, ROE x retention growth, 2nd-half glide to stable |
| `fcfe` | fcfeginzu.xls (ships with Levi Strauss May-2024) | NI x (1 - equity reinvestment rate), normalizations, R&D |
| `ddm_stable`/`_2stage`/`_3stage` | ddmst, ddm2st, ddm3st .xls | Gordon; 2-stage; 3-stage (linear transition) |
| `fcfe_stable`/`_2stage`/`_3stage` | fcfest, fcfe2st, fcfe3st .xls | FCFE = EPS - net capex(1-DR) - dWC(1-DR) |
| `fcfe_vs_ddm` | fcfevsddm.xls | FCFE vs DDM value; gap = PV of cash build-up |

## Calibration (`python scripts/financial_firms.py --selftest`: 81/81 cases, max error 1e-15)
| file | sheet value | ours |
|---|---|---|
| CitiVal2023 (May 2023) | 68.5841 /share | 68.5841 |
| DeutscheBank2016 (Oct 2016) | 22.9717, after 10% wipeout 20.6745 | 22.9717 / 20.6745 |
| divginzu (JPMorgan, May 2024) | 231.1608 (assets in place 172.96, stable 35.00, extraordinary 23.20) | same |
| fcfeginzu (Levi Strauss, May 2024) | 18.9239 | 18.9239 |
| eqexret | 74.9440 | 74.9440 |
| ddmst / ddm2st / ddm3st | 46.4510 / 177.5385 / 19.9455 | same |
| fcfest / fcfe2st / fcfe3st | 40.9716 / 51.5636 / 12.6080 | same |
| fcfevsddm | FCFE 1864.93, DDM 1855.68 | same |

Plus 69 branch cases: the source file recalculated with pycel under option switches the source files never use (failure
risk, DOJ-fine / Tier-1 / ROE / risk choices, normalized earnings and reinvestment approaches 0-3, R&D, direct
growth, no gradual glide, stable payout/ROE/beta overrides, 0/5/7/15-year horizons, direct cost of equity,
item-specific growth, debt mixes, stable capex rules, fixed transition payout/beta) - all exact. The two branches
that hit the sheet slips (below) run with `replicate_sheet_bugs: true`; the corrected defaults were checked in Excel
on copies with the slips fixed: fcfeginzu reported earnings 9.1100 (as written 9.3175), fcfeginzu 3 years 14.6706
(14.4306), fcfe2st WC 30% / debt 30%-20% 55.9308 (56.4369) - all equal to ours.
Rebuild: `--build-cases <source files> --xlsx-dir <Excel-converted .xlsx copies>` (pycel reads only .xlsx).

## Spreadsheet defaults and when Damodaran overrides
- Bank FCFE: stable growth = asset growth (Citi 3%) or inflation (Deutsche 1%); target ROE = 5-yr average (Citi)
  or 25th percentile, then the stable cost of equity by year 10 (Deutsche); failure probability 0 for large banks
  with sticky deposits, higher "for smaller banks, with sensitive deposits & dangerously low Tier 1 ratios".
- `ddm`: stable growth = riskfree rate (override is "dangerous territory"); stable payout = 1 - g/ROE_stable;
  inputs glide to stable in the second half of the high-growth period (`gradual`, default true).
- `fcfe`: stable equity reinvestment = g/ROE_stable; growth = non-cash ROE x equity reinvestment rate.
- Stage templates: stable cost of equity is always CAPM (beta_stable or beta), even if `cost_of_equity` is direct.

## Inputs (full lists in each function's docstring; rates decimal; `_`-prefixed keys ignored)
- `bank_fcfe`: tier1_capital, book_equity, risk_adjusted_assets, net_income, asset_growth, target_tier1_ratio,
  target_roe, stable_growth, shares_outstanding, price, cost_of_equity, [stable_cost_of_equity,
  failure_probability]; DeutscheBank2016 adds tier1_ratio_next_year, [target_roe_10, tier1_capital_adjustment].
- `excess_return`: net_income, book_equity(+_last_year), eps, dps, shares_outstanding, beta, riskfree_rate, erp,
  years, stable_growth, [roe_high, retention_high, roe_stable, stable_payout, beta_stable, normalize_earnings].
- `ddm`: book_equity_last_year, net_income, dividends, shares_outstanding, price, beta, riskfree_rate, erp, years,
  [growth, roe_high, retention_high, roe_stable, stable_growth, stable_payout, beta_stable, gradual].
- `fcfe`: net_income, interest_income_after_tax, book_equity(+_last_year), cash(+_last_year), shares_outstanding,
  capex, depreciation, change_wc, net_debt_issued, beta, riskfree_rate, erp, years, stable_growth, [rd,
  normalize_earnings, normalize_reinvestment, growth, roe_high, reinvestment_rate_high, roe_stable, beta_stable].
- `erp` may be a number or a `cost_of_capital.py` spec ({"approach": "country", "country": "Germany"}, Jan-2026).

## Example (Damodaran's Citi inputs)
```
python scripts/financial_firms.py --from-case CitiVal2023 > citi.json   # edit, then:
python scripts/financial_firms.py citi.json          # -> value/share 68.58, price/value 0.675, year table
```
`citi.json`: {"model": "bank_fcfe", "tier1_capital": 169145, "book_equity": 182194, "risk_adjusted_assets":
1142985, "net_income": 14845, "asset_growth": 0.03, "target_tier1_ratio": 0.15, "target_roe": 0.095,
"stable_growth": 0.03, "shares_outstanding": 1958.3, "price": 46.32, "cost_of_equity": 0.1167}

## Sheet errors: corrected by default, `"replicate_sheet_bugs": true` reproduces the source file (report prints both)
- fcfeginzu B38 (reported earnings): denominator B36 - B7 subtracts interest income twice (B36 is already net of
  it; the sheet's normalized branch subtracts it once). Year-1 net income grows at the year-2 rate (C14 = D5*(1+D12));
  matters only for 2-3 years with the glide, and the reason 1 year is rejected when replicating.
- fcfe2st row 121: high-growth working capital financed at the capex debt ratio (E95), while the sheet's current FCFE and
  terminal year use the working-capital ratio (E96); matters only when the two ratios differ.

## Limits and template quirks reproduced on purpose
- CitiVal2023 'current ROE' = net income / **Tier 1 capital** (the 8.78% Damodaran quotes); NI_t = ROE_t x prior-year
  book equity; DeutscheBank2016 uses same-year book equity and a terminal NI = NI_10 x (1+g).
- divginzu: year-1 earnings = current net income (not grown); its 0-year branch mixes per-share and total
  values, so `ddm` needs 1-15 years (use `ddm_stable` for a stable firm).
- fcfeginzu: its columns for years 11-15 compute FCFE as NI x reinvestment rate (a slip), so `fcfe` stops at 10;
  a typed stable ROE drives g/ROE even when switched off (not reproduced: we use ROE_high; the case builder passes
  the sheet's rate as `stable_reinvestment_rate`).
- Terminal values use MAX of net income / book equity over the growth years (same as last year when growing).
- fcfe3st's capex/depreciation stable rule grows depreciation at earnings growth x max cumulated transition growth.
- fcfeginzu's 'Normalized Reinvestment' sheet ships with another company's numbers: the branch tests verify
  formulas, not those values. Stage templates: at most 10 years per stage.
- No live data: bank Tier 1 capital, risk-adjusted assets and peer P/BV/ROE must come from the user or filings.
