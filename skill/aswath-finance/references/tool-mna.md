# Tool: `scripts/mna.py`: value of control, value of synergy, most an acquirer can pay

Deterministic transcription of Damodaran's own control and synergy spreadsheets. It uses stdlib only and is
calibrated against six of the source files plus pycel branch tests of the templates (132/132 checks within 0.05%).

## What it answers
- **Value of control**: status quo vs optimally run (investment, financing, dividend policy), plus the expected value
  with a probability of management change.
- **Value of synergy**: acquirer, target and combined firm. Operating synergy comes from higher margin/EBIT (cost
  savings), higher ROC, reinvestment rate or a longer growth period (growth). Financial synergy comes from a lower
  tax rate, lower cost of debt, or a higher debt ratio (debt capacity).
- **How much can an acquirer pay**: target (with control) + synergy, as total equity, per share and as a breakeven
  exchange ratio.

Damodaran's three steps ([Acquisition Alchemy](https://aswathdamodaran.blogspot.com/2012/12/acquisition-alchemy-creating-value-from.html), 2012-12): (1) status quo value of the target;
(2) optimal value, control = optimal − status quo; (3) combined firm with synergy, synergy = combined − (target
*with control* + acquirer). Never pay a premium above control + synergy; negotiate for a share of it.

## Sources
| File | URL | Used in |
|---|---|---|
| controlvalue.xls (tools page) | https://pages.stern.nyu.edu/~adamodar/pc/controlvalue.xls | [Acquisition Alchemy](https://aswathdamodaran.blogspot.com/2012/12/acquisition-alchemy-creating-value-from.html), 2012-12 |
| synergyvaluation.xls (tools page) | https://pages.stern.nyu.edu/~adamodar/pc/synergyvaluation.xls | spreadsh.htm calculator |
| KFTcontrolvalue.xls (Kraft) | https://pages.stern.nyu.edu/~adamodar/pc/blog/KFTcontrolvalue.xls | [Activist Value Investing](https://aswathdamodaran.blogspot.com/2012/06/activist-value-investing-be-your-own.html), 2012-06 |
| TWCComcast.xls | https://pages.stern.nyu.edu/~adamodar/pc/blog/TWCComcast.xls | [Comcast bids for Time Warner Cable](https://aswathdamodaran.blogspot.com/2014/02/comcast-bids-for-time-warner-cable.html), 2014-02 |
| SABControl.xls, InbevSABrestrusynergy.xls | https://pages.stern.nyu.edu/~adamodar/pc/blog/ | [Winning at a Loser's Game?](https://aswathdamodaran.blogspot.com/2015/10/winning-at-losers-game-control-synergy.html), 2015-10 |

The expected value of control, `optimal × p + status quo × (1 − p)`, is from
[Alibaba's Governance by Politburo](https://aswathdamodaran.blogspot.com/2014/09/alibabas-governance-by-politburo.html), 2014-09.

**Firm model** (every sheet): cost of equity = rf + β·ERP; cost of debt = pre-tax kd × (1 − marginal tax) (2015
sheets: kd = rf + default spread); WACC by D/C; g = after-tax ROC × reinvestment rate; NOPAT = EBIT × (1 − effective
tax); n-year PV of NOPAT × (1 − RR); TV = NOPAT·(1+g)^n·(1+rf)·(1 − rf/ROC_s)/(WACC − rf). Stable growth = riskfree;
ROC_s = WACC (Damodaran's stated default, zero excess returns); the 2015 deal sheets keep ROC_s = ROC (`"stable_roc": "roc"`).
**Betas:** optimal = status quo unlevered beta relevered at the optimal D/E and tax; combined = unlevered betas
weighted by firm value, relevered at the combined D/E (the tools template weights the *terminal* WACC by terminal
values: `terminal_weights`).

## Calibration (cached sheet outputs vs ours; all within 0.05%)
| File | Sheet value | Ours |
|---|---|---|
| controlvalue.xls | status quo 68,466.9; optimal 76,150.3; control 7,683.5 | same |
| synergyvaluation.xls | acquirer 458.19, target 271.50, combined 729.79, synergy 0.106 | same |
| KFTcontrolvalue.xls (Kraft) | 59,179.9 / 64,586.5 / control 5,406.7 ("$59bn, $65bn, $6bn" in the post) | same¹ |
| SABControl.xls | status quo 43,747.2, optimal 48,449.4, control 4,702.2; equity 50,916.3 / 55,618.5 | same² |
| InbevSABrestrusynergy.xls | InBev 211,952.8, SAB 50,065.4, combined 276,609.9, **synergy 14,591.8**, max 52,766.1 = $32.59/sh, exchange ratio 0.3158 | same |
| TWCComcast.xls | Comcast 176,573.9, TWC 72,827.2, equity 130,302.7 / 47,585.0, combined beta 1.0612, WACC 6.45% | same |
| TWCComcast.xls, combined | 254,226.0, **synergy 4,824.8** (the post's "$4.82 billion") | **260,317.0, 10,915.8**³ ; the sheet's figure with `replicate_sheet_bugs` |
| pycel branches, controlvalue | 8 scenarios: debt mix, taxes, ROC/RR, growth periods, EBIT×1.1 default, rf/ERP, ROC < WACC | all match |
| pycel branches, synergyvaluation | 7 scenarios: cost, growth, debt-capacity and tax synergy, dissimilar firms, rf/ERP | all match |

Notes on the calibration table:

¹ In Damodaran's Kraft file, C27 is typed as 20% over `=C16` (30%). Growth uses 20% while the cash flows reinvest 30%. The
case passes `expected_growth`.

² SABControl C6 relevers with the *effective* tax rate (a row shift in the extended sheet). The case passes the sheet's beta.

³ **Source formula error.** TWC E44 multiplies by `(1+D2)`, and D2 is blank. The combined terminal value therefore
loses the (1 + riskfree) growth factor. The corrected synergy is about $10.9bn, not $4.8bn. The engine uses the
correct formula (the one the InBev sheet uses). `"replicate_sheet_bugs": true` (synergy mode) drops the factor from the
combined TV and reproduces the sheet's 254,226.0 / 4,824.8 exactly (a NOTE is printed).

## Spreadsheet defaults and when Damodaran overrides them
- **Optimal inputs:** anything not given is inherited from the status quo. EBIT defaults to status quo × 1.1 (the
  template's C13), and the output flags it. In SABControl, Damodaran set the optimal values at industry benchmarks:
  pre-tax ROIC, reinvestment rate and D/C of the global industry averages.
  - With `industry` set, the control report prints the firm next to Damodaran's Jan 2026 industry averages, plus the five
    hints: debt mix; margin vs industry; ROC vs cost of capital; reinvestment when ROC is high; a longer growth
    period only with excess returns.
- **Combined firm** (when not given, per the 2015 worksheet's status quo column): revenues and EBIT summed; taxes
  EBIT-weighted; ERP revenue-weighted; D/E = total debt / total equity value; ROC = combined NOPAT / combined
  capital; RR = total reinvestment / total NOPAT; years = the target's. The combined cost of debt must be given if
  the two firms' costs differ (it depends on a new rating).
- **Synergy is measured against the target with control** (`target_optimal`), as in the InBev file.

## Inputs (JSON; decimals; one currency)
```
firm: beta | unlevered_beta, erp, pre_tax_cost_of_debt | default_spread | after_tax_cost_of_debt,
      tax_rate | marginal_tax_rate + effective_tax_rate, debt_to_capital | debt_to_equity,
      ebit, roc | pretax_roc, reinvestment_rate, years, [expected_growth, stable_roc (number|"roc"),
      stable_growth, stable_cost_of_capital, revenues, cash, debt, cross_holdings, minority_interests, shares, price]
control: {"mode":"control", riskfree_rate, erp, status_quo:{firm}, optimal:{changes}, probability_of_change, industry}
synergy: {"mode":"synergy", riskfree_rate, erp, acquirer:{firm}, target:{firm}, target_optimal:{changes},
          combined:{ebit | operating_margin, roc, reinvestment_rate, years, taxes, debt_to_capital,
                    pre_tax_cost_of_debt | default_spread, terminal_weights "terminal_value"|"current_value"},
          replicate_sheet_bugs (the TWC sheet's E44 slip)}
```

## Example (Damodaran's Kraft inputs, 2012)
```
python scripts/mna.py kraft.json
{"mode":"control","riskfree_rate":0.02,"erp":0.06,"probability_of_change":0.5,"industry":"Food Processing",
 "status_quo":{"beta":1.1,"pre_tax_cost_of_debt":0.04,"tax_rate":0.4,"debt_to_capital":0.28,"revenues":54365,
   "ebit":6657,"pretax_roc":0.1106,"reinvestment_rate":0.3,"years":5},
 "optimal":{"pre_tax_cost_of_debt":0.045,"debt_to_capital":0.4,"ebit":7322.7,"pretax_roc":0.125,
   "reinvestment_rate":0.3,"expected_growth":0.015}}
-> status quo 59,179.9; optimal 64,586.5; VALUE OF CONTROL 5,406.7; expected value (p = 50%) 61,883.2
```

Other commands:
- The inputs of every calibrated deal are in `data/mna_cases.json` (`"inputs"`). Copy one and edit it.
- `python scripts/mna.py --selftest` replays all cases.
- `--build-cases <files>` rebuilds them from the source spreadsheets. Maintenance venv; the `.xlsx` Save-As copies of the
  `.xls` drive pycel.

## Limits
- Two-stage model only: one high-growth period, then stable growth at the riskfree rate. No year-by-year paths,
  no R&D/lease conversion (the 2015 workbook does those on side sheets: pass the adjusted EBIT, ROC and RR), and no
  failure risk. For a full DCF of either firm use `ginzu.py`, then feed its value drivers here.
- Synergy is not split into components. Run once per synergy (e.g. margin only, then debt capacity) to see each one.
- Not transcribed:
  - **lboval.xls**: Excel 4.0 file (Excel's File Block refuses it; xlrd gives values only), a 10-year LBO debt
    schedule model, not simple.
  - **synergy.xls** (1998 FCFE version linked in the 2012-12 and 2014-03 posts): superseded by synergyvaluation.xls.
  - **controlvalue.xls's 'Industry averages' sheet**: undated; the engine uses Damodaran's Jan 2026 industry data instead.
- Accretion/dilution is deliberately absent. Damodaran calls it useless for judging deals (2012-12).
