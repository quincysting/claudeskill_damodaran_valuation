# Tool: `scripts/multiples.py`: fundamentals-justified multiples and pricing against peers

Deterministic transcription of Damodaran's multiple spreadsheets (stdlib only). It is calibrated against four of
the source files plus pycel branch tests of the templates: 198/198 checks within 0.05%.

## What it answers
- **What multiple do the fundamentals justify?**
  - PE (forward/trailing), PBV and PS from growth, payout (or ROE) and risk (cost of equity).
  - EV/EBIT(1−t), EV/EBIT, EV/Invested capital, EV/Sales from growth, reinvestment (or ROIC), cost of capital and
    after-tax margin. EV/EBITDA only through the identity EV/EBIT × (1 − D&A/EBITDA).
- **Is the stock cheap on multiple X relative to its fundamentals?**
  - `actual` gives actual ÷ justified.
  - `--solve FIELD:MULTIPLE` gives the growth, ROE or cost of equity that the market multiple implies.
  - `industry` prints Damodaran's Jan 2026 industry multiples and their drivers.
- **What would peers pay?** `comps` mode applies peer multiples (median or average) to the firm's metrics, as in
  Damodaran's Alibaba pricing.

Damodaran's rules: relative valuation done right = similar firms, a consistent multiple (equity value with equity metrics,
EV with operating metrics), matched timing (trailing vs forward), and **control for differences** by deriving the
multiple's drivers from a DCF.
- "Most stocks that look cheap deserve to be cheap."
- "Absolute rules don't belong in a relative world."

Sources: [January 2018 Data Update 10](https://aswathdamodaran.blogspot.com/2018/02/january-2018-data-update-10-price-is.html), 2018-02;
[January 2019 Data Update 9](https://aswathdamodaran.blogspot.com/2019/02/january-2019-data-update-9-pricing-game.html), 2019-02.
See `pricing-markets.md` §2 for the driver table (PE: growth, payout, risk; PBV: ROE; EV/Sales: margin; ...).

## Sources
| File | URL | Used in |
|---|---|---|
| eqmult.xls (tools page, saved Jan 2025) | https://pages.stern.nyu.edu/~adamodar/pc/eqmult.xls | spreadsh.htm calculator |
| firmmult.xls (tools page, saved Jan 2025) | https://pages.stern.nyu.edu/~adamodar/pc/firmmult.xls | spreadsh.htm calculator |
| EqMultCalculator.xlsx | https://pages.stern.nyu.edu/~adamodar/pc/blog/EqMultCalculator.xlsx | [Interest Rates, Earning Growth and Equity Value](https://aswathdamodaran.blogspot.com/2021/03/rates-growth-and-value-investment.html), 2021-03 |
| AlibabaRelValue.xls | https://pages.stern.nyu.edu/~adamodar/pc/blog/AlibabaRelValue.xls | [Alibaba's coming out party](https://aswathdamodaran.blogspot.com/2014/09/alibabas-coming-out-party-valued-right.html), 2014-09 |

The stable-growth form PBV = (ROE − g)/(r − g), inverted to get the implied cost of equity, is from
[Deutsche Bank: A Greek Tragedy?](https://aswathdamodaran.blogspot.com/2016/10/deutsche-bank-greek-tragedy-at-german.html), 2016-10.

**Equity multiples** (eqmult.xls D27): two-stage DDM per $1 of next year's net income.

`PE_fwd = payout·(1 − ((1+g)/(1+r))^n)/(r − g) + (1+g)^(n−1)·(1+gs)·payout_s/((rs − gs)(1+r)^n)`

- ROE = g/(1 − payout); PBV = PE_fwd × ROE (book = next year's net income / ROE).
- PS_fwd = PE_fwd × net margin; trailing multiples = forward × (1+g).
- EqMultCalculator values current EPS: PBV = (ROE·PV_high + ROE_s·PV_terminal)/EPS (`pbv_roe_trailing`).
- Its `life_cycle` preset: g = rf + 10% (Growth) / 1% (Mature) / −5% (Declining); cost of equity = rf + 5%;
  stable growth = rf; payout = 1 − g/ROE.

**EV multiples** (firmmult.xls C29): the same formula with (1 − reinvestment rate) in place of payout and the cost
of capital in place of r.
- ROIC = g/RR; EV/IC = EV/NOPAT_fwd × ROIC; EV/Sales = × after-tax margin; EV/EBIT = × (1 − t).
- `years = 0` gives the stable-growth multiples.

## Calibration (cached sheet outputs vs ours; all within 0.05%)
| File | Sheet value | Ours |
|---|---|---|
| eqmult.xls | PE 18.6785, PBV 2.9696, PS fwd 0.4670, PS trailing 0.5137, ROE 15.90% | same |
| firmmult.xls | EV/EBIT(1−t) 22.178 fwd / 24.308 trailing, EV/EBIT 13.307 / 14.585, EV/IC 3.549, EV/Sales 3.200 / 3.508 | same |
| EqMultCalculator.xlsx (Declining) | value/share 28.555, PE (trailing) 14.277, PBV 2.1416 | same |
| AlibabaRelValue.xls (Median) | 25 equity values, e.g. PE × Baidu 113,934; EV/EBITDA × all online 76,708 | same |
| pycel branches, eqmult | 8 scenarios: n, g/payout, stable inputs, rf, ERP (stable follows), beta, margin, g < r | all match |
| pycel branches, firmmult | 6 scenarios: n, g/RR, stable inputs, WACC, margin, tax | all match |
| pycel branches, EqMultCalculator | Growth (PE 31.94, PBV 4.79), Mature (PE 19.25, PBV 2.89), rf 4%/ROE 20%, 5 years | all match |
| eqmult/firmmult at **g = r exactly** | the sheet's `IF(g=r)` branch: PE 23.29, EV/NOPAT 26.00 | **16.71, 19.71**¹ ; the sheet's figure with `replicate_sheet_bugs` |

¹ The sheet's IF branch returns n × earnings (it omits the payout or (1 − RR), and the discounting). The engine uses the
limit of the main formula, which is continuous: 16.70966 / 16.70968 / 16.70969 at g = r − 1e-7, r, r + 1e-7.
At g = r the report prints a NOTE; `"replicate_sheet_bugs": true` reproduces the IF branch (all 15 cells match).

Other quirks in the source files, not transcribed:
- The side "PBV/PS/EV-to-IC equation" cells (eqmult G30, firmmult F37/F40) use (1+r)^5 or the high-growth r in the
  stable denominator, so they disagree with the main outputs.
- The Sheet2/Sheet3 data tables hold stale static values from older inputs.

## Spreadsheet defaults
- Stable ERP = ERP (eqmult D8 = C8).
- If no stable inputs are given: stable cost of equity = cost of equity, stable growth = riskfree, stable ROE = ROE
  (EqMultCalculator C15–C17).
- Peer statistic: median (the Alibaba default); `"statistic": "average"` is available. Damodaran warns that averages are
  skewed by outliers.
- Firm mode has no defaults: all stable inputs are required (firmmult's are plain inputs).

## Inputs (JSON; decimals)
```
equity: years, growth, payout | roe, cost_of_equity | riskfree_rate+beta+erp, stable_growth, stable_payout |
        stable_roe, stable_cost_of_equity | stable_beta (+stable_erp, stable_riskfree_rate), net_margin, eps,
        life_cycle ("Growth"|"Mature"|"Declining" with riskfree_rate, roe, years)
firm:   years, growth, reinvestment_rate | roic, cost_of_capital, stable_growth, stable_reinvestment_rate |
        stable_roic, stable_cost_of_capital, after_tax_operating_margin, tax_rate, da_to_ebitda
both:   actual {pe_forward|pe_trailing|pbv|ps_forward|ev_ebit_forward|ev_sales_forward|...: x}, industry, scope,
        replicate_sheet_bugs (the sheet's IF(g=r) branch)
comps:  firm {revenues, ebitda, ebit, net_income, book_equity, debt, cash, other_holdings, shares, market_cap},
        peers {group: {"PE"|"PBV"|"PS"|"EV/EBITDA"|"EV/EBIT"|"EV/Sales"|"EV/Invested Capital": x or [x, ...]}}
```

## Example (Damodaran's Deutsche Bank median bank, 2016-10)
```
python scripts/multiples.py db.json --solve cost_of_equity:pbv
{"mode":"equity","years":0,"growth":0.016,"roe":0.0991,"cost_of_equity":0.09,"stable_growth":0.016,
 "actual":{"pbv":1.06},"industry":"Bank (Money Center)"}
-> IMPLIED cost_of_equity for pbv = 1.06: 9.4396%   (Damodaran's post: 9.44%)
```

Other commands:
- `python scripts/multiples.py --selftest` replays all cases.
- `--build-cases <files>` rebuilds them (maintenance venv; the .xlsx Save-As copies of eqmult/firmmult drive pycel).

## Limits
- Two-stage models with a constant high-growth rate, as in the source sheets; no fade period.
- Comps mode applies given peer multiples. It does not screen peers, compute aggregates (sum of market cap / sum of
  earnings) from raw rows, or run sector regressions.
- EV/EBITDA is not in the source sheet. It comes only from the D&A identity.
- Industry context (`data/ginzu_data.json`, Jan 2026) holds Damodaran's US/Global industry averages, not the full
  pedata/pbvdata/psdata/vebitda datasets, which are not in `finance/spreadsheets`.
- `--solve` scans −30% to 60% and holds every other given input fixed. With `growth` solved, the payout (ROE given)
  or ROIC (RR given) moves with it.
