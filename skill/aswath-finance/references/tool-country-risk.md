# Tool: country_risk.py (country risk and equity risk premiums, calculated the way Damodaran does it)

**What it answers.** It gives the ERP, CRP, default spread, rating, CDS and marginal tax rate for a country, either
the latest figures (July 2026) or as of a date. It also recomputes a CRP/ERP from a rating, a default spread or a
CDS spread (spread × relative equity volatility + mature ERP). It gives a company's revenue-weighted (or
production-weighted) ERP across countries and Damodaran's regions, and a local-currency riskfree rate (govt bond −
default spread). Every number comes from the ctryprem spreadsheets or is recomputed with their formulas.

## Sources
| File (`finance/spreadsheets`, URL = pages.stern.nyu.edu/~adamodar/pc/datasets/<file>) | Vintage | Post |
|---|---|---|
| ctrypremJuly23.xlsx (+ Ratings&CDS2023.xlsx, same spread table) | 2023-07-01 | [Country Risk: A July 2023 Update!](https://aswathdamodaran.blogspot.com/2023/07/country-risk-july-2023-update.html), 2023-07 |
| ctrypremJuly24.xlsx | 2024-07-01 | [Country Risk: My 2024 Data Update](https://aswathdamodaran.blogspot.com/2024/07/country-risk-my-2024-data-update.html), 2024-07 |
| ctrypremJuly25.xlsx | 2025-07-01 | [Country Risk 2025: The Story behind the Numbers!](https://aswathdamodaran.blogspot.com/2025/07/country-risk-2025-story-behind-numbers.html), 2025-07 |
| ctryprem.xlsx | 2026-01-01 (ratings updated Feb 16) | [Data Update 4 for 2026: The Global Perspective!](https://aswathdamodaran.blogspot.com/2026/02/data-update-4-for-2026-global.html), 2026-02 |
| ctrypremApr26.xlsx | 2026-04-01 | [Oil, War and the Global Economy](https://aswathdamodaran.blogspot.com/2026/04/oil-war-and-global-economy-markets.html), 2026-04 |
| ctrypremJuly26.xlsx (latest) | 2026-07-01 (spreads corrected July 9) | [Country Risk: Drivers, Measures and Investment Implications - The 2026 Edition!](https://aswathdamodaran.blogspot.com/2026/07/country-risk-drivers-measures-and.html), 2026-07 |
| countrytaxrates.xlsx (Tax Foundation; + global minimum tax) | 2026-01-05 | [Data Update 7 for 2026: Debt and Taxes](https://aswathdamodaran.blogspot.com/2026/02/data-update-7-for-2026-debt-and-taxes.html), 2026-02 |
| CountryTaxRates2025.xlsx (blog/; effective vs marginal, 2025) | 2025 | same post |

The change after the May 2025 downgrade is explained in [Sovereign Ratings, Default Risk and Markets: The Moody's
Downgrade Aftermath!](https://aswathdamodaran.blogspot.com/2025/06/sovereign-ratings-default-risk-and.html) (2025-06)
and in the 2025-07 and 2026-07 posts.

## Damodaran's method (sheet → function)
1. **Rating** ('Sovereign Ratings'): the sheet uses Moody's local-currency rating. If Moody's does not rate a country, it
   maps the S&P rating notch-for-notch to Moody's. In the 2026 files Nepal falls back to Fitch
   (`moodys_equivalent`).
2. **Rating → spread** ('Default Spreads for Ratings'): each update scales the previous table by
   (1 + median % change in sovereign CDS across the ~80 countries quoted at both dates). Damodaran also types a few overrides:
   `C` = 1750 bp, and Aa1 = 27 bp in July 2025 ("adjusted to reflect US addition to Aa1") (`update_spreads`).
3. **Scalar** ('Relative Equity Volatility') = the average of 5 annual standard deviations of daily S&P Emerging BMI
   returns, divided by the same average for the iShares EMB sovereign-bond ETF. It is a ratio of averages. Each
   standard deviation is STDEV × √(days) (`relative_volatility`). Values: 1.418 (Jul 23), 1.298 (Jul 24), 1.505
   (Jul 25), 1.523 (Jan 26), 1.534 (Apr 26), 1.554 (Jul 26).
4. **CRP = spread × scalar; ERP = mature ERP + CRP** (`premiums`, `country_row`). CDS version: CDS minus the
   benchmark CDS, floored at 0. The benchmark is the **US** in the 2023–24 files and **Switzerland** from July 2025.
   The CDS columns are always scaled; the "Yes/No" switch (E5) only affects the rating-based CRP.
5. **Mature ERP.** Up to the January 2025 update, it was the implied S&P 500 ERP (US = Aaa). From July 2025 it is
   the implied ERP over (T.Bond − US Aa1 spread), **minus the US default spread, unscaled**: 4.48% − 0.27% = 4.21%
   (Jul 25), 4.46% − 0.23% = 4.23% (Jan 26), 4.42% − 0.22% = 4.20% (Jul 26) (`mature_erp`). The FAQ text writes
   "− spread × relative volatility", but the cells and posts subtract the unscaled spread. From Jan 2026 the US row
   is the US implied ERP, with an unscaled CRP. In the **July 2025 file the US row is mature + scaled CRP = 4.62%**,
   not the 4.48% in the post; lookups and weights print a caution on that row.
6. **Unrated countries** ('PRS Worksheet'): the PRS composite score falls into a band (lower bound, rating). The ERP
   is that rating's ERP, CRP = ERP − mature, and the spread is CRP / scalar (`prs_premium`). Bands: <50 → C (1750 bp),
   50–55 Ca, … , 72–74 Baa2, … , >85 Aaa.
7. **Regions** ('Regional Weighted Averages'): GDP-weighted spread, CRP and tax; ERP = mature + weighted CRP. In the
   July 2024 file it is instead the weighted sum of the pasted ERPs, which carry a stale 4.11% mature ERP versus
   4.12%, 1 bp off (`region_premiums`).
8. **Companies**: Damodaran measures exposure by where the company operates, not where it is incorporated. The ERP is the
   revenue-weighted average of country ERPs, as in the Ginzu "countries" option. The method uses production weights for
   natural-resource firms and a blend for manufacturers (`company_erp`). Local riskfree rate = local govt bond −
   sovereign default spread, e.g. India July 2025: 6.32% − 2.16% = 4.16% (`local_riskfree`).

## Calibration (`python country_risk.py --selftest` → 7/7)
Each vintage starts from Damodaran's raw inputs: Moody's/S&P/Fitch ratings, the prior spread table, CDS at both dates,
raw CDS and benchmark, daily index prices, mature/US ERP, PRS scores and regional GDP. From these the tool
recomputes the scalar, the median CDS change, the 22-rating table, and for every country the rating, spread,
CRP, ERP, net CDS, CDS CRP and CDS ERP. It then recomputes the PRS frontier rows and the GDP-weighted
regions/Global. The tolerance is 1e-6 absolute or 0.05% relative, and every check passes.

| File | Checks | Scalar sheet / ours | India ERP sheet / ours | Baa3 bp sheet / ours |
|---|---|---|---|---|
| ctrypremJuly23 | 1,225 | 1.4183 / 1.4183 | 8.3346% / 8.3346% (post: 8.33%) | 235.11 / 235.11 |
| ctrypremJuly24 | 1,218 | 1.2980 / 1.2980 | 6.8043% / 6.8043% | 206.79 / 206.79 |
| ctrypremJuly25 | 1,225 | 1.5046 / 1.5046 | 7.4613% / 7.4613% (post: 7.46%) | 216.10 / 216.10 |
| ctryprem (Jan 26) | 1,226 | 1.5234 / 1.5234 | 7.0754% / 7.0754% | 186.78 / 186.78 |
| ctrypremApr26 | 1,229 | 1.5344 / 1.5344 | 7.9791% / 7.9791% | 209.15 / 209.15 |
| ctrypremJuly26 | 1,229 | 1.5545 / 1.5545 | 6.9202% / 6.9202% | 174.99 / 174.99 |
| July26, E5 = No (pycel recalculation of the template) | 30 | — | CRP = spread | — |

Quirks in the source files that are reproduced as inputs rather than "fixed":
- Curacao is typed Baa2 in July 2023.
- Guernsey and Jersey are typed Moody's NR in July 2025, so their S&P rating is used.
- Nepal uses Fitch.
- Russia is dropped from the July 2024 list.
- Stale pasted regional CRPs: Kenya and Turkey in Jan 26; US, Austria and Finland in July 25. The regional averages
  use the pasted values, as the source sheet does.

A 1 bp change in the mature ERP produces 262 mismatches, so the test is sensitive.

## Inputs (`python country_risk.py inputs.json [--json]`; rates as decimals; `_notes` keys ignored)
- `country`, `as_of` (YYYY-MM-DD; picks the latest vintage on or before that date; the earliest is 2023-07-01).
- `premium`: one of `rating`, `default_spread`, or `cds` (+ `cds_benchmark`). Optional `mature_erp`, `scalar`
  (defaults: the vintage's values) and `scale` (default true).
- `weights`: `{name: revenues}`. Names can be countries, Damodaran's regions (`Western Europe`, `Asia`, …, `Global`) or the sheet's
  groupings (`Emerging Markets`, `Latin America & Caribbean`, `Africa & Mid East`, `EMEA`, `Small Asia (No India,
  China & Japan)`, …).
- `riskfree`: `govt_bond` + `country` or `default_spread`.
- `tax`: country. On or after 2026-01-05 this returns the statutory rate from countrytaxrates.xlsx (with the
  global-minimum-tax variant); earlier dates return the vintage's rate. Both add the 2025 effective rate (aggregate,
  quartiles).
- Shortcut: `--country India [--as-of 2025-01-15]`. Name matching is case-insensitive and fuzzy, and US/USA/UK work.

Example (July 2025 numbers from Damodaran's post):
```
echo {"country":"Turkey","as_of":"2025-07-01","riskfree":{"govt_bond":0.0632,"country":"India","as_of":"2025-07-01"}} > q.json
python country_risk.py q.json
-> Turkey B1: spread 4.43%, CRP 6.66%, ERP 10.87% (post: 10.87%); CDS 3.80% net 3.63%, ERP 9.67%
-> Local riskfree = 6.32% - 2.16% = 4.16% (post: 4.16%)
```

## Limits
- Only six vintages, from 2023-07 to 2026-07. For earlier numbers, quote Damodaran's posts (references/country-risk.md). The
  Jan 2023 raw files (sovrrating2023, sovrCDS2023) are not turned into a vintage, because the scalar for that date
  is not in them.
- The data is Damodaran's, and ratings and CDS go stale between updates. For a live CDS or rating, use `premium` with the
  new input and the latest scalar and mature ERP.
- Not covered: country risk dimensions (CountryRiskDimensions2023), market pricing and statistics (Country2023,
  CountryRisk2024, countrypricing2025, countrystats). Those are descriptive data in the posts, not ERP inputs.
- For a currency conversion of the cost of equity (differential inflation), see references/country-risk.md. This
  tool does not forecast inflation.
- `--build-cases` rebuilds both JSON files from the source xlsx. It needs the maintenance venv (openpyxl; pycel for the
  branch case).
