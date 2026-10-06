# Tool: market.py — index intrinsic value, breakeven revenues, PE history

`scripts/market.py` (stdlib only) answers, with Damodaran's own spreadsheets transcribed cell by cell:
- **What is the S&P 500 worth the Damodaran way?** Consensus earnings → cash payout (dividends + buybacks) fading to the
  sustainable 1 − g/ROE → discounted at riskfree + ERP, stable growth = riskfree. Reports value, over/under-valuation,
  intrinsic trailing/forward PE and CAPE, and the **implied ERP** (the sheet's goal seek: the ERP at which value = index).
- **Is the market over/under-valued?** Value vs index at a chosen ERP, and the implied ERP next to Damodaran's yardsticks
  (ERP10 = average of the last decade, ERPLong = 1960-current average, ERPHigh/ERPLow, HUS = historical US premium).
- **What revenues must company X / an AI business / a group of leaders reach to justify its market cap or EV?**
  Breakeven revenue if mature today, after waiting N years (with a cash yield), the CAGR needed, and the sheet's two tables.
- **PE history of the index**: PE, normalized PE, CAPE, T.Bond PE, E/P vs T.Bond, period averages, 1960-2025.
- **How much of the market are the leaders?** Fearsome Five market cap vs all US equities, 2009-2024.
- For the published monthly ERP series, the cash-flow menu and currency riskfree rates use `implied_erp.py`
  (`tool-implied-erp.md`); template `2025-09` below is the same 'Impl premium calculator' sheet.

## Source files and the posts that use them (pages.stern.nyu.edu/~adamodar/pc/blog/<file>)
| File | Tool / template | Post |
|---|---|---|
| S&P500ValueJan2023.xlsx | index `2023-01` | [Data Update 2 for 2023: A Rocky Year for Equities!](https://aswathdamodaran.blogspot.com/2023/01/data-update-2-for-2023-rocky-year-for.html), 2023-01 |
| S&P500ValueJuly2023.xlsx | index `2023-07` | [Market Resilience or Investors In Denial?](https://aswathdamodaran.blogspot.com/2023/07/market-resilience-or-investors-in.html), 2023-07 |
| S&P500ValueOct2023.xlsx | index `2023-10` | [Market Bipolarity: Exuberance versus Exhaustion!](https://aswathdamodaran.blogspot.com/2023/10/market-bipolarit-exuberance-versus.html), 2023-10 |
| S&P500ValueJan2024.xlsx | index `2024-01` | [Data Update 2 for 2024: A Stock Comeback](https://aswathdamodaran.blogspot.com/2024/01/data-update-2-for-2024-stock-comeback.html), 2024-01 |
| S&PValueJan2025.xlsx | index `2024-01` | [Data Update 2 for 2025: The Party Continued](https://aswathdamodaran.blogspot.com/2025/01/data-update-2-for-2025-party-continued.html), 2025-01 |
| IntrinsicValueS&P500Sept25.xlsx (+ pc/implprem/ERPOct25, ERPJan26) | index `2025-09` | [A "Fairly Highly Valued" Market](https://aswathdamodaran.blogspot.com/2025/10/a-fairly-highly-valued-market-fed-chair.html), 2025-10; [Data Update 2 for 2026](https://aswathdamodaran.blogspot.com/2026/01/data-update-2-for-2026-equities-get.html), 2026-01 |
| MktCapBreakeven.xlsx | breakeven, equity | [Trillion Dollar Market Caps](https://aswathdamodaran.blogspot.com/2025/12/trillion-dollar-market-caps-fairy-tale.html), 2025-12 |
| BreakevenAIBusiness, BreakevenAnthropic, BreakevenEnterprisevalue.xlsx | breakeven, firm | [AI's Bar Mitzvah Moment](https://aswathdamodaran.blogspot.com/2026/08/ais-bar-mitzvah-moment-from-hype-hope.html), 2026-08 |
| PEHistorical2025.xlsx / PEHistory2026.xlsx | `--pe` | 2025-10 and 2026-01 posts above |
| FearsomeFiveMktCap.xlsx | `--leaders` | [Breaking up Big Tech: Cui Bono?](https://aswathdamodaran.blogspot.com/2024/10/breaking-up-big-tech-cui-bono.html), 2024-10 |

## Calibration (`python scripts/market.py --selftest`: 64/64, all exact except the post-text case)
| File | Sheet value | Ours |
|---|---|---|
| S&P500ValueJan2023 (index 3839.5) | 3648.39, over 5.24%, PE 16.62 / fwd 17.64 | same |
| S&P500ValueJuly2023 (4450.38) | 4369.68, over 1.85% | same |
| S&P500ValueOct2023 (4288.05) | 4147.04, over 3.40% | same |
| S&P500ValueJan2024 (4769.83) | 4365.12, over 9.27% | same |
| S&PValueJan2025 (5881.63) | 5261.80, over 11.78% | same |
| IntrinsicValueS&P500Sept25 (6688.46) | 5937.92 at ERP 4.25%; top-down growth 8.936% | same |
| ERPOct25 / ERPJan26 | 5024.28 / 6814.70; implied ERP 3.7682% / 4.2312% | same (ERP to 1e-9) |
| MktCapBreakeven ($5T, 52.41%, ROE 40%, 9%) | 509,624; 714,774 after 5 yrs at 2%; 80 + 56 table cells | same |
| BreakevenAIBusiness ($5T, 20%, ROIC 15%, 10%) | 1,966,783; 8,215,742 after 15 yrs; CAGR 26.22%; 99 + 56 cells | same |
| BreakevenAnthropic, BreakevenEnterprisevalue ($2T, 30%, 25%) | 457,875; 1,912,659; CAGR 25.29% | same |
| PEHistory2026 / PEHistorical2025 | 499 / 509 cells (every PE, normalized PE, CAPE, T.Bond PE, ratio, P/D, P/cash, period average) | same |
| FearsomeFiveMktCap | 77 cells (total, share, changes) | same |
| pycel recalculation of the templates | 42 branches: riskfree change gradual/immediate/end, earnings adjustment on/off, long-term growth, payout fixed, reported base, month; risk D4, direct rate, finite life | same |

The one text case: [Nvidia at $5 trillion](https://aswathdamodaran.blogspot.com/2025/12/trillion-dollar-market-caps-fairy-tale.html),
2025-12 — net margin 53.01%, ROE 64.44%, 9% cost of equity: the post's $483.38 billion and $677.97 billion (5 years, 2% cash
yield); ours 483.48 / 678.11 (0.02%, the post's rounded inputs). The AI post's round numbers also reproduce: Anthropic
$1.19T at 10 years, $1.91T at 15; the AI business $5.10T at 10, $8.22T at 15 ([2026-08](https://aswathdamodaran.blogspot.com/2026/08/ais-bar-mitzvah-moment-from-hype-hope.html)).

## Spreadsheet defaults and when Damodaran overrides them
- **Earnings**: analyst consensus for the next 2-3 years, which Damodaran calls unbiased in aggregate
  ([2024-01](https://aswathdamodaran.blogspot.com/2024/01/data-update-2-for-2024-stock-comeback.html)); growth then fades to the riskfree rate.
  Only Jan 2023 cut them (`earnings_adjustment` -10%).
- **Cash payout**: dividends + buybacks / earnings in the base year, moving to 1 − g/ROE by year 5.
- **ROE**: 20% (Jan/Jul 2023), 19.73% (Oct 2023), 17.04% (Jan 2024), 17.5% (Jan 2025), 10-year average 18.0% (2025-09).
- **Riskfree** = 10-year T.Bond that day; stable growth = riskfree. The "change riskfree after year 5" switch is never used.
- **ERP**: 5% in 2023-24 ("ERPDirect"), 4.5% in Jan 2025, 4.25% (1960-2024 average) since Sept 2025.
- **2025-09 month**: time to first cash flow = months left in the year / 12, prorated first cash flow; the file says
  October — set `month` to the current month.
- **Breakeven, equity**: cost of equity by risk class D1-D5 = 12/9/8/7/6% (default D2, 9%), stable growth 4%,
  5-year wait, 2% cash yield. **Firm**: cost of capital 12.5/10/8.5/7.5/6% (D2, 10%), 15-year wait, 0% yield.
  The sheet's "inflation" cell (2.5%) feeds no formula. The Nvidia post uses current margin/ROE, then lower ones.
- **Target level a year out** (not computed): value × (1 + expected return − dividend yield) ([2024-01](https://aswathdamodaran.blogspot.com/2024/01/data-update-2-for-2024-stock-comeback.html)).

## Inputs (JSON; keys starting with `_` are ignored)
- Index, all templates: `index_level`, `riskfree_rate`, `erp` (number or ERP10/ERPLong/ERPHigh/ERPLow/HUS/HG), `roe`,
  `cape_denominator` (optional).
- `2023-01` … `2024-01`: `base_earnings`, `earnings_forecasts` [E1, E2] (+E3 in 2023-07/2023-10), `base_payout`
  (2023-01) or `base_cash_return`, `earnings_adjustment` (omit = No), `riskfree_change`, `riskfree_after_5`,
  `riskfree_change_mode` gradual | immediate | end.
- `2025-09` (default): `base_earnings` (consensus, last year), `trailing_earnings` (reported LTM), `base_cash_flow`,
  `growth` or `consensus` [E0, E1, E2], `adjust_payout` (true), `long_term_growth` (riskfree), `month`, `default_spread`.
- Breakeven (`"model": "breakeven"`): `basis` equity | firm, `value_to_match`, `margin` (net or after-tax operating),
  `return_on_capital` (ROE or ROIC), `risk_class` or `discount_rate`, `stable_growth`, `years_to_maturity`,
  `cash_yield`, `current_revenues` (for CAGR), `finite_years` (see limits). For a group, sum the market caps and use
  the blended margin.

## Examples
```
python scripts/market.py --from-case "IntrinsicValueS&P500Sept25" > sp.json   # Damodaran's inputs, Sept 30 2025
python scripts/market.py sp.json       # 5,937.92 vs 6,688.46: overvalued 12.6%; implied ERP 3.78%
```
```
{"model": "breakeven", "basis": "firm", "value_to_match": 2000000, "margin": 0.30, "return_on_capital": 0.25,
 "risk_class": "D2", "years_to_maturity": 10, "cash_yield": 0, "current_revenues": 65000}
```
→ $457,875 million if mature today and $1,187,611 million after 10 years (CAGR 33.7%): Damodaran's Anthropic case. Other
modes are `--pe`, `--leaders` and `--json`.

## Limits and quirks found in the source files
- **Finite-life breakeven** (Anthropic/EV sheets, never used in the saved files): the sheet formula multiplies by
  ((1+g)/(1+r))^N, the value of a perpetuity deferred N years, instead of the growing-annuity factor
  1 − ((1+g)/(1+r))^N. The engine uses the growing-annuity factor by default and prints the discrepancy with the
  sheet's figure (`his_sheet_breakeven_revenue_now`);
  `"replicate_sheet_bugs": true` reproduces the source sheet (used by the calibration case). The sensitivity table always uses perpetuity.
- **Intrinsic CAPE** in the 2023-25 S&P files divides by AVERAGE('Buyback & Dividend computation'!C37:C46), i.e.
  nominal earnings for 2009-18 or 2008-17 rather than the latest inflation-adjusted decade. Pass your own `cape_denominator`.
- **Sept 2025 file**: its implied ERP cell (4.0127%) is stale against the saved inputs. It reproduces with
  `adjust_payout: false` (the 4.01% in the [2025-10 post](https://aswathdamodaran.blogspot.com/2025/10/a-fairly-highly-valued-market-fed-chair.html));
  with the saved inputs the implied ERP is 3.78%.
- **PE period rows**: the row labelled "2001-2010" averages 2001-2020 in both PE files, and the 2025 file's
  "EP - T.Bond Rate" column averages T.Bond − T.Bill. The cases use the sheet's ranges; `--pe` prints true periods.
- 'Value Rates Growth' (Jan/July 2023) is a pasted Sept 2022 table without inputs: not calibrated.
- The source workbooks contain circular references (Jan 2023: E24↔F24 through the adjustment IF). The pycel branch values
  sum the recalculated PV cells, because pycel's cycle mode leaves the SUM cell stale.
- It needs index-level inputs (level, consensus earnings, trailing dividends + buybacks, ROE); `--from-case` gives
  Damodaran's; `implied_erp.py` holds the monthly ERP series. It does not forecast; it values at today's rates.
