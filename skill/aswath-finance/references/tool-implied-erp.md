# Tool: implied_erp.py — implied ERP, intrinsic riskfree rate, currency riskfree rates

`scripts/implied_erp.py` (stdlib only) answers:
- **Implied ERP** for an index, given its level, the cash yield (dividends + buybacks), expected earnings growth and the
  riskfree rate. It also reports the ERP against a default-free US$ rate and values the index at any ERP.
- **What ERP did Damodaran publish on date X?** It looks up the daily crisis series (Sep–Dec 2008, 2020, Dec 2024–Apr 2025),
  the start-of-month table (Sep 2008–Jul 2026, with every variant) and the annual series (1960–2024).
- **Intrinsic riskfree rate** (inflation + real GDP growth; smoothed = 10-yr averages) against the T.Bond rate.
- **Riskfree rate in currency X**: government bond minus the sovereign default spread, or differential inflation.

## Damodaran's method (sheet 'Impl premium calculator', ERPJan23 … ERPJuly26)
- E0 = consensus earnings for the most recent year, E_t = E0(1+g)^t for years 1–5 (g = top-down analyst CAGR). Payout
  p0 = (dividends + buybacks, trailing 12m) / reported LTM earnings, constant or ("sustainable") linear to 1 − g_stable/ROE
  (10-yr average ROE) by year 5. CF_t = E_t·p_t; TV = E_5(1+g_stable)·p_T / (r − g_stable), g_stable = riskfree unless
  long-term growth is overridden. r = T.Bond + ERP; the ERP is goal-seeked so that PV = index.
- Since 2025 the first cash flow is prorated. With τ = months left in the year/12, CF_1·τ is discounted at τ and CF_t at
  τ+t−1. Older files use τ = 1.
- After the May 2025 Moody's downgrade (Aaa→Aa1), a second ERP is added: ERP + US default spread, measured against
  T.Bond − spread. Damodaran calls the T.Bond-based ERP the *mature market premium*. Numerically they are the same thing
  ([Data Update 4 for 2026](https://aswathdamodaran.blogspot.com/2026/02/data-update-4-for-2026-global.html), 2026-02).
- **Cash-flow menu:** C = trailing 12m (**default**); A10/A5 = average yield over the last 10/5 years × index
  ("Smoothed"); P10/P5 = average payout × LTM earnings; NORM = average payout × average 10-yr earnings ("Normalized");
  NC = net of stock issuance ("Net Cash Yield"). `--from-case` prints all of them under `_cash_choices`.
- **Other defaults:** base earnings are "Estimated". The ERP used to value the index is typed in (4.25% in Jan/Mar/Jul
  2026; 5% earlier).
- The daily crisis ERPs hold the start-of-month cash flows and growth fixed and update only the index and the T.Bond
  rate. `--from-case` followed by an edit of `index`/`riskfree` does the same.
- Intrinsic riskfree = inflation + real GDP growth. "Fed effect" = T.Bond − intrinsic.
- **Currency, sovereign method:** rf = local govt bond − default spread for the sovereign's rating.
- **Currency, inflation method:** rf_local = (1+rf_US$)(1+π_local)/(1+π_US$) − 1, using IMF expected inflation. The Jan
  2026 table uses the additive form rf_US$ + π_local − π_US$ and falls back to actual inflation when the forecast is NA.
  Since 2026 the US$ base is T.Bond − US default spread (4.18% − 0.23% = 3.95%).

## Sources and posts
| File(s) | Post |
|---|---|
| ERPJan23 | [Data Update 2 for 2023: A Rocky Year for Equities!](https://aswathdamodaran.blogspot.com/2023/01/data-update-2-for-2023-rocky-year-for.html), 2023-01 |
| ERPAug23 | [The Price of Risk: With Equity Risk Premiums, Caveat Emptor!](https://aswathdamodaran.blogspot.com/2023/08/the-price-of-risk-with-equity-risk.html), 2023-08 |
| ERPJan24 | [Data Update 2 for 2024: A Stock Comeback](https://aswathdamodaran.blogspot.com/2024/01/data-update-2-for-2024-stock-comeback.html), 2024-01 |
| ERPJan25 | [Data Update 2 for 2025: The Party Continued](https://aswathdamodaran.blogspot.com/2025/01/data-update-2-for-2025-party-continued.html), 2025-01 |
| ERPJuly25 | [Country Risk 2025: The Story behind the Numbers!](https://aswathdamodaran.blogspot.com/2025/07/country-risk-2025-story-behind-numbers.html), 2025-07 |
| ERPOct25 | [A "Fairly Highly Valued" Market](https://aswathdamodaran.blogspot.com/2025/10/a-fairly-highly-valued-market-fed-chair.html), 2025-10 |
| ERPJan26 | [Data Update 2 for 2026: Equities get tested, and pass again!](https://aswathdamodaran.blogspot.com/2026/01/data-update-2-for-2026-equities-get.html), 2026-01 |
| ERPMarch26 | [The Price of Risk: An Equity Risk Premium Monologue!](https://aswathdamodaran.blogspot.com/2026/03/the-price-of-risk-equity-risk-premium.html), 2026-03 |
| ERPJuly26, DiffInflationRiskfree26 | [Country Risk: Drivers, Measures and Investment Implications - 2026](https://aswathdamodaran.blogspot.com/2026/07/country-risk-drivers-measures-and.html), 2026-07 |
| TariffERPbyday, ERPbyDay2008Crisis | [Anatomy of a Market Crisis: Tariffs, Markets and the Economy!](https://aswathdamodaran.blogspot.com/2025/04/anatomy-of-market-crisis-tariffs-rock.html), 2025-04 |
| ERPbyDayCOVID | [A Viral Market Meltdown VI: The Price of Risk](https://aswathdamodaran.blogspot.com/2020/04/a-viral-market-meltdown-vi-price-of-risk.html), 2020-04 |
| intrinsicvsactualrate23 | [Data Update 3 for 2023: Inflation and Interest Rates](https://aswathdamodaran.blogspot.com/2023/01/data-update-3-for-2023-inflation-and.html), 2023-01 |
| intrinsicvsactualrate24, currencyriskfree2024 | [Data Update 3 for 2024: Interest Rates in 2023](https://aswathdamodaran.blogspot.com/2024/01/data-update-3-for-2024-interest-rates.html), 2024-01 |
| intrinsicvsactualrate25 | [Data Update 4 for 2025: Interest Rates, Inflation and Central Banks!](https://aswathdamodaran.blogspot.com/2025/01/data-update-4-for-2025-interest-rates.html), 2025-01 |
| IntrinsicRiskfree2026 | [Data Update 3 for 2026: The Trust Deficit](https://aswathdamodaran.blogspot.com/2026/01/data-update-3-for-2026-trust-deficit.html), 2026-01 |
| intrinsicvsactualrateSept26 | [Interest Rates and Stock Prices: An Old Debate Flares up!](https://aswathdamodaran.blogspot.com/2026/09/interest-rates-and-stock-prices-old.html), 2026-09 |
| currencyriskfree2025 | [Data Update 5 for 2025: It's a small world, after all!](https://aswathdamodaran.blogspot.com/2025/02/data-update-5-for-2025-its-small-world.html), 2025-02 |
| CurrencyRiskfreeJuly2024 | [Country Risk: My 2024 Data Update](https://aswathdamodaran.blogspot.com/2024/07/country-risk-my-2024-data-update.html), 2024-07 |
| DiffInflRiskfree2026 | [Data Update 4 for 2026: The Global Perspective!](https://aswathdamodaran.blogspot.com/2026/02/data-update-4-for-2026-global.html), 2026-02 |
| Currency&Inflation2023 | [Country Risk: A July 2023 Update!](https://aswathdamodaran.blogspot.com/2023/07/country-risk-july-2023-update.html), 2023-07 |

## Calibration (`--selftest`: 108/108)
Each ERP below is the sheet's goal-seeked cell, and our solution matches it within 5e-9. The second value is the ERP against
the default-free rate.

| File (τ) | Sheet ERP | Ours | File (τ) | Sheet ERP (vs default-free) | Ours |
|---|---|---|---|---|---|
| ERPJan23 (1) | 5.9409% | 5.9409% | ERPOct25 (5/12) | 3.7682% (4.0382%) | same |
| ERPAug23 (1) | 4.8297% | 4.8297% | ERPJan26 (1) | 4.2312% (4.4612%) | same |
| ERPJan24 (1) | 4.5994% | 4.5994% | ERPMarch26 (10/12) | 4.3700% (4.6000%) | same |
| ERPJan25 (1) | 4.3328% | 4.3328% | ERPJuly26 (0.5) | 4.1696% (4.3896%) | same |
| ERPJuly25 (0.5) | 4.2107% (4.6107%) | same | | | |

Additional checks in the selftest:
- The index value at the sheet's fair-value ERP matches in all 9 files, to 1e-15.
- Variants from Damodaran's own monthly table (T12m, sustainable payout, Smoothed, Normalized, Net Cash Yield), published at
  0.01%: 43/45 match within rounding.
- Branches the source files never use, recalculated with pycel on ERPJan26 and ERPJuly26: sustainable payout, stable growth
  3%, both at 2.5%, reported-LTM base, month, NORM cash. 12/12 match to 1e-15.
- Post figures: Oct 2025 as posted (cash C, October) gives 4.01%; value at 4.25% with sustainable payout ≈ 5940;
  Jan 2026 value at 6% = 4790.
- Intrinsic riskfree (latest year of each of the 5 files) and every row of 7 currency tables (898 rows): exact.

**Not reproducible from the saved files** (the selftest prints these as info):
- **ERPOct25 saved state** uses the net-cash-yield (NC) cash flow and month "August" (3.77%). Its post reports 4.01%,
  which reproduces with cash C and October.
- **ERPJuly26:** the post and the monthly table say 4.20% (4.42% against the default-free rate), but the file solves
  to 4.17%. Reaching 4.20% would need growth of 13.84% instead of the file's 13.69%.
- **ERPAug23:** the post's 4.44% matches no variant in the file; its monthly-table net cash yield (4.62%) vs 4.65%.
- **ERPJan26 post:** says 14834 at a 2% ERP; the model gives 14634, which looks like a typo.
- **ERPJuly25:** the post nets out a 0.27% default spread, while the file cell holds 0.40%.
- **2008 daily series:** from Oct 7 (cash flow 52.584, growth as in the file), 58/60 days match within input rounding.
  Oct 10 and Oct 13 are off by 0.01–0.02%. Before Oct 7 Damodaran used a base cash flow the file does not state.

## Inputs
- `erp`: `index`, `riskfree`, `cash_flow` (or `cash_yield`), `growth`; optional `earnings_ltm` + `earnings_base`
  (required with `earnings_ltm`; the sheet's default base is current-year *estimated* earnings, = `earnings_ltm` for its
  'Reported LTM' option) (payout
  = cash_flow/earnings_ltm), `month` or `months_to_year_end` (default January, τ = 1), `stable_growth` (default
  riskfree), `sustainable_payout` + `roe`, `default_spread`, `erp`, `value_at`.
- `intrinsic`: `inflation`, `real_growth`, `treasury_rate`, and optionally `inflation_10y`/`real_growth_10y`. Or
  `--year 2025` / `--year 2026-09` to use Damodaran's data.
- `currency`: either `bond_rate` with one of `default_spread`, `rating` or `country` (Jan 2026 spreads from
  ginzu_data.json), or `inflation_local` with optional `inflation_base`/`base_riskfree` (default: the Jul 2026 table,
  4.23% and 2.35%) and `additive`. `--lookup Brazil` shows every vintage.

```
python scripts/implied_erp.py --from-case ERPJuly26 > in.json   # Damodaran's 1 Jul 2026 inputs; edit index/riskfree
python scripts/implied_erp.py erp in.json          # -> ERP 4.1696%, vs default-free 4.3896%, values at 2-6%
python scripts/implied_erp.py history 2025-04-08   # -> daily 5.14%, April 1 table 4.61% (+variants), annual
python scripts/implied_erp.py intrinsic --year 2026-09   # -> 5.41% vs T.Bond 4.84%; smoothed 5.47%
python scripts/implied_erp.py currency --lookup Brazil
```

## Limits
- It does not choose growth for you. Damodaran uses the top-down analyst CAGR; bottom-up and fundamental growth
  (g = (1 − payout)·ROE) are in the source sheet but are not computed here.
- The COVID-adjusted ERP (2020 earnings drop plus recovery) is available only as data. Damodaran's COVID model file is not
  among the sources.
- The 1954–62 smoothed intrinsic rates differ by up to 5e-6 because the sheet's helper columns hold rounded copies.
- Kept as published: currencyriskfree2025 rows that look misaligned (Brazil 2.85%, Bulgaria 14.78%, Russia 2.88%,
  Romania 15.22%); the Jul 2024 bond sheet repeats Jan 2024 rates (whose header says "12/31/22"). `currency --lookup` prints a caution on these rows.
- Maintenance: `finance\.venv\Scripts\python.exe scripts/implied_erp.py --build-cases <the 23 spreadsheets>` (openpyxl,
  xlrd, pycel) rebuilds `data/implied_erp_cases.json` and `data/implied_erp_data.json`.
