# Tool: `scripts/options.py` (option pricing and real options, Damodaran's way)

**Answers:** value of warrants and employee options with dilution (pre-tax, vesting-adjusted, after-tax); short-term
(discrete dividends) and long-term (dividend yield) Black-Scholes calls and puts; Black-Scholes inputs converted to a
binomial tree; the option to delay, expand or abandon a project; a patent or undeveloped natural-resource reserve as an
option; the value of financial flexibility; equity in a distressed firm as a call on firm value (plus implied debt
value, interest rate and risk-neutral default probability). Each tool is a cell-by-cell transcription of one of Damodaran's
calculators; stdlib only.

## Source files and where Damodaran uses them
| tool | source calculator (pages.stern.nyu.edu/~adamodar/pc/...) | used in |
|---|---|---|
| `warrant` | warrant.xls; the "Option value" sheet of every company Ginzu | $10 stock/$10 strike/5y/40%: $3.36 per option, [Options and Taxes: Is a "Facebook" tax next?](https://aswathdamodaran.blogspot.com/2012/02/options-and-taxes-is-facebook-tax-next.html), 2012-02; Tesla options netted out of equity, [Share Count Confusion](https://aswathdamodaran.blogspot.com/2018/07/share-count-confusion-dilution-employee.html), 2018-07 |
| `employee_options` | employeeoption.xls | same logic plus its maturity / tax / vesting switches |
| `long_term` | optlt.xls; older version buffettoptions.xls | BofA warrants: $4.30/option x 700m = ~$3bn (file uses sd 60%; the post text says 50%), [Buffett and Bank of America](https://aswathdamodaran.blogspot.com/2011/08/buffett-and-bank-of-america-poker-and.html), 2011-08 |
| `short_term` | optst.xls | listed options < 1 year with known dividends |
| `binomial` | bstobin.xls | BS misprices deep out-of-the-money and illiquid options: modify BS or use a binomial, [Buffett and Black-Scholes](https://aswathdamodaran.blogspot.com/2011/02/buffett-and-black-scholes.html), 2011-02 |
| `expand` | expand.xls | linked for "optionality", [If it is a strategic growth investment in China...](https://aswathdamodaran.blogspot.com/2014/03/if-it-is-strategic-growth-investment-in.html), 2014-03 |
| `delay`, `abandon`, `patent` | delay.xls, abandon.xls, project.xls | real options on projects and product patents |
| `natural_resource` | natres.xls | undeveloped reserves (Petrobras-type), 2011-02; [The School Bell Rings!](https://aswathdamodaran.blogspot.com/2016/09/the-school-bell-rings-its-time-for-class.html), 2016-09 |
| `flexibility` | flexval.xls | value of excess debt capacity (financial flexibility) |
| `equity_option` | equity.xls | distressed equity = buying options, [A Viral Market Meltdown IV](https://aswathdamodaran.blogspot.com/2020/03/a-viral-market-meltdown-iv-investing.html), 2020-03 |

Damodaran's rules on real options, from the posts (the tool does not enforce them; state them when you use it):
- Option value is an addendum to a DCF, not a substitute: value the business first, then the option; the test is
  **exclusivity** (proprietary technology, exclusive license) - no exclusivity, no option premium (2014-03).
- "Even when I have estimated a real option premium, I have never obtained an increase of more than 20-25% on my DCF
  value" ([Following up on Groupon](https://aswathdamodaran.blogspot.com/2011/11/following-up-on-groupon.html), 2011-11).
- Optionality in a user base is worth most with sticky, intense users and an unpredictable future
  ([User/Subscriber Economics](https://aswathdamodaran.blogspot.com/2017/07/usersubscriber-economics-value-dynamics.html), 2017-07);
  inputs are hard to get before the option takes form - Damodaran used Facebook's optionality (2018) as the tie-breaker, not
  as a number ([The Zomato IPO](https://aswathdamodaran.blogspot.com/2021/07/the-zomato-ipo-bet-on-big-markets-and.html), 2021-07).
- Employee options: value with an option pricing model, net out of equity, divide by shares outstanding including
  restricted shares; never treasury-stock or "fully diluted" shortcuts (2018-07). Long-dated options need a long-term
  variance and dilution/illiquidity corrections (2011-02).

## Calibration (`python options.py --selftest`: 54/54 within 0.05% / 1e-6)
Every case checks all the sheet's output cells (d1, N(d1), d2, N(d2), adjusted S, values), not just the headline.

| file (cached sheet values) | headline | sheet | ours |
|---|---|---|---|
| warrant.xls | value per option | 3.3610 | 3.3610 |
| employeeoption.xls | after-tax value, all options | 697.0477 | 697.0477 |
| optst.xls (replicate_sheet_bugs) / optlt.xls | call | 6.3908 / 377.4080 | 6.3908 / 377.4080 |
| buffettoptions.xls (2011, BofA) | call | 4.3024 | 4.3024 |
| bstobin.xls | up move (and full 5-period tree) | 1.19937 | 1.19937 |
| delay / expand / abandon.xls | option value | 51.0306 / 45.9137 / 10.5615 | same |
| project.xls / natres.xls | option value | 70,185.16 / 472,622.08 | 70,185.13 / 472,622.12 |
| flexval.xls | annual value of flexibility | 0.021170 | 0.021170 |
| equity.xls | equity as a call (debt 182,597, rate 14.80%, P(default) 81.0%) | 55,913.92 | 55,913.92 |
| Option value sheet: Adani 2023, Instacart IPO, SpaceX 2026 IPO, Tesla Nov21/Jan23/Oct23/Jan24/Jan25, fcffsimpleginzu | value per option | e.g. Tesla Jan25 109.4769, SpaceX 478.6927 | identical |

Plus 32 pycel recalculations of the templates under branches the source files never use: warrant dividend yield and deep
out-of-the-money dilution; employee maturity override / vesting / no tax / all on; optst with 0, 2 and 4 dividends; optlt
dividend yield; buffettoptions dividends and yield override; bstobin quarterly/monthly; delay, expand (cost of delay),
abandon, patent (cost-of-delay override), natres variants; flexval with a negative reinvestment year; equity with firm-input
variance, direct variance, another industry, dividends and a yield override. Skipped: SpaceX2026IPOUpdated,
BirkenstockIPO2023, NVIDIA2024Feb (their cached option cells are `#VALUE!`), and the other 2023-24 company Ginzus whose
option sheet holds the template placeholder (strike 1.29, 7.72m options), not Damodaran's inputs.

## Spreadsheet defaults and quirks (kept on purpose, except one sheet error)
- Warrant block: adjusted S = (S x shares + option value x options)/(shares + options), solved by iteration like the
  circular sheet; the company Ginzus feed the *estimated value per share* as S (Tesla) or the price.
- employeeoption: maturity override "No" (1.5y unused), **tax adjust "Yes" at 26%** (`tax_rate`, 0 to switch off),
  vesting "No" (90% unused). The sheet's after-tax value (B41) is taken from the pre-vesting total (B39), not B40 (kept as
  the sheet's figure; with vesting on, a NOTE prints the vesting-adjusted after-tax value).
- optst: rate used = 1 - EXP(-r); PV of dividends at the annual rate; at most 4 dividends; options under one year.
  **Sheet error, corrected by default:** the sheet's call (D53) and put (F54) discount the PV of dividends a second time, so
  they use a different S than its d1 and than Damodaran's stated method (S minus the PV of dividends). Default = S - PV(div)
  throughout (the default file: call 6.3839 vs the sheet's 6.3908); the report prints a NOTE with the sheet's figures;
  `"replicate_sheet_bugs": true` reproduces D53/F54 (set in the calibration cases that have dividends).
- optlt: takes the continuous rate directly (variance input 0.25 = 50% sd); the 2011 version converts an annual rate
  with 1 - EXP(-r) and turns annual dividends into a yield on the price (`riskfree_annual`, `annual_dividends`).
- Cost of delay (dividend yield): delay = 1/years of rights; abandon = 1/project life (and years <= life); patent =
  1/life unless overridden; natural resource = annual after-tax cash flow / (reserves x (price - unit cost)); expand = 0.
- Flexibility: 1-year option, no yield; value = [call(K = financeable %) - call(K = maximum %)] x (ROC - WACC)/WACC;
  the history sheet takes AVERAGE(reinvestment/value) and STDEV(LN ratio), counting a non-positive year as 0.
- Equity: strike = face value of all debt **plus coupons**, t = face-value-weighted duration; variance from
  (default) Damodaran's industry table (sd of firm value, `data/options_data.json`, equity.xls Jan 2026), firm inputs
  sqrt((1-w)^2 ss^2 + w^2 sb^2 + 2w(1-w) ss sb rho), or direct; yield = dividends / firm value unless overridden.

## Inputs (JSON, `"tool"` + fields; rates decimal; `sd` or `variance`)
- warrant / employee_options: price, strike, maturity, sd, riskfree, warrants, shares, dividend_yield (0);
  employee: maturity_override, tax_rate (0.26), vesting_probability
- short_term: price, sd, strike, days, riskfree, dividends [[dps, days], ...], replicate_sheet_bugs (false)
- long_term: price, strike, maturity, sd, riskfree | riskfree_annual, dividend_yield | annual_dividends
- binomial: price, sd, riskfree, dividend_yield, periods_per_year (1), periods (5)
- delay: pv_cash_flows, sd, investment, years, riskfree - expand: pv_expansion, sd, investment, years, riskfree,
  cost_of_delay - abandon: pv_cash_flows, sd, project_life, abandonment_value, years, riskfree
- patent: pv_cash_flows, sd, development_cost, years, riskfree, cost_of_delay
- natural_resource: reserves, price, unit_cost, sd, annual_cash_flow, development_cost, years, riskfree
- flexibility: reinvestment_pct, sd, financeable_pct, max_pct, riskfree, cost_of_capital, return_on_capital, history
- equity_option: firm_value, face_value_debt, duration, riskfree, dividends | dividend_yield, sd | industry | firm_inputs

## Example (Damodaran's Tesla Jan 2025 option inputs)
```
python scripts/options.py --from-case TeslaJan2025 > tsla_opts.json   # {"tool": "warrant", "price": 147.16, "strike": 40.41,
python scripts/options.py tsla_opts.json                              #  "maturity": 4, "sd": 0.3, "riskfree": 0.0422,
# -> adjusted S 143.50, value per option 109.48, all 344.5m options $37,715m  "warrants": 344.5, "shares": 3197}
python scripts/options.py --from-case equity.xls > eq.json && python scripts/options.py eq.json --json
```

## Limits
- Black-Scholes only (European, lognormal, constant variance); `binomial` returns the sheet's up/down moves and price tree,
  not an option value - bstobin.xls itself stops there. No early-exercise or barrier features.
- Industry lookup is exact-name (the sheet's VLOOKUP is approximate over rows 2-95); "Total Market" rows are allowed here.
- Inputs (PV of cash flows, variance, exclusivity) are the hard part; the tool gives no judgement on them - see the
  rules above. Rebuild cases: `--build-cases` the source .xls/.xlsx files plus Excel-saved .xlsx copies of the templates.
