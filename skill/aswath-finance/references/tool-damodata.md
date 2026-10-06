# damodata: Damodaran's datasets as queryable data

`scripts/damodata.py` + `data/damodata_data.json` (stdlib only). Use it for questions like: "median/aggregate operating
margin, beta, cost of capital, debt ratio, payout, tax rate, ROIC, EVA for industry X in region Y", "industries sorted
by metric", "arithmetic/geometric returns and risk premiums of stocks, T-bonds and T-bills over any period", and
"implied ERP by year or month". Every value is a source cell, under the source column name and in the source units (decimals; US $ millions
where the header says so). Every answer prints the **source file, URL and as-of date**. Quote those; don't round away
the source units.

## Source files (Jan 2026 website versions) and where Damodaran uses them
| Dataset key | Files (regions) | As of | Post |
|---|---|---|---|
| betas, wacc | betas/beta*.xls, wacc*.xls (US, Global, Europe, Japan, Emerging, China, India) | 2026-01-05 | [Data Update 5 for 2026: Risk and Hurdle Rates](https://aswathdamodaran.blogspot.com/2026/02/data-update-5-for-2026-risk-and-hurdle.html), 2026-02 |
| margin, EVA | margin*.xls, EVA*.xls (7 regions) | 2026-01-05 | [Data Update 7 for 2025: The End Game in Business!](https://aswathdamodaran.blogspot.com/2025/02/data-update-7-for-2025-end-game-in.html), 2025-02 |
| dbtfund, debtdetails | dbtfund*.xls, debtdetails*.xls (7 regions) | 2026-01-05 | [Data Update 8 for 2025: Debt, Taxes and Default](https://aswathdamodaran.blogspot.com/2025/02/data-update-8-for-2025-debt-taxes-and.html), 2025-02 |
| divfund, divfcfe | divfund*.xls, divfcfe*.xls (7 regions) | 2026-01-05 | [Data Update 9 for 2025: Dividends and Buybacks](https://aswathdamodaran.blogspot.com/2025/03/data-update-9-for-2025-dividends-and.html), 2025-03 |
| taxrate, mgnroc | taxrate.xls, mgnroc.xls (US only) | 2026-01-05 | [Data Update 1 for 2025: The Draw (and Danger) of Data](https://aswathdamodaran.blogspot.com/2025/01/data-update-1-for-2025-draw-and-danger.html), 2025-01 |
| IndustryMargins, IndustryReturns | Industry{Margins,Returns}{US,Global}.xlsx: aggregate + quartiles/median, ROE/ROIC, excess-return spreads | post 2026-02-16 | [Data Update 6 for 2026: In Search of Profitability!](https://aswathdamodaran.blogspot.com/2026/02/data-update-6-for-2026-in-search-of.html), 2026-02 |
| DebtIndustry | DebtIndustry{US,Global}2025.xlsx: coverage, Debt/EBITDA, debt/capital quartiles | post 2026-02-20 | [Data Update 7 for 2026: Debt and Taxes](https://aswathdamodaran.blogspot.com/2026/02/data-update-7-for-2026-debt-and-taxes.html), 2026-02 |
| DividendsIndustry, BuybacksIndustry | {Dividends,Buybacks}Industry{US,Global}.xlsx | post 2026-02-24 | [Data Update 8 for 2026: Time for Harvesting](https://aswathdamodaran.blogspot.com/2026/02/data-update-8-for-2026-time-for.html), 2026-02 |
| WACCdistn | WACCdistn.xlsx: cost-of-capital deciles/quartiles by industry (Global) | **Jan 2025** (post 2025-02-08) | [Data Update 6 for 2025: The Hurdle Rate Question](https://aswathdamodaran.blogspot.com/2025/02/data-update-6-for-2025-from-macro-to.html), 2025-02 |
| history | histretSP.xlsx 'Returns by year' 1928-2025 (+ inflation) | 2026-01-01 | [The Price of Risk: An Equity Risk Premium Monologue!](https://aswathdamodaran.blogspot.com/2026/03/the-price-of-risk-equity-risk-premium.html), 2026-03 |
| implied_erp | histimpl.xlsx 'Historical Impl Premiums' 1960-2025 | 2026-01 (data through end-2025) | same post, 2026-03 |
| erp_monthly | ERPbymonth.xlsx 'Historical ERP', Sep 2008 to Sep 2026 | 2026-09-01 | same post, 2026-03 |
| companies | indname.xlsx 'By industry': 48,156 firms, industry group and country | Jan 2026 | Data Update 1 for 2025 (above) |

Blog files have no date cell, so their as-of is the date of the post that links them. The xls datasets also carry Damodaran's
input settings (`columns <dataset>` prints them), e.g. wacc.xls US: T.Bond 3.95%, ERP 4.46%, marginal tax 25%;
waccGlobal.xls: ERP 5.63%, default spread 0.94%, tax 25.22%. The cost of capital and EVA columns were computed with
those settings. Rerun `cost_of_capital.py` if you need other inputs.

## Calibration (`python scripts/damodata.py --selftest`)
**536/536 cases match**, within 1e-6 absolute or 0.05% relative (incl. 6 erroneous source cells, reproduced as the sheet computes them).
- (a) Spot checks: 5 seeded-random cells per sheet (10 for two-sheet files), read directly from every one of the
  69 industry files (7 regions x 8 dataset families, taxrate, mgnroc and the 11 blog files), 8 cells each from
  histretSP, histimpl and ERPbymonth, and 18 company fields from indname. All return the source cell.
- (b) The sheets' summary rows, recomputed by `period_stats()` from the annual returns:

| Sheet cell (histretSP / histimpl) | Period | Sheet value | Ours |
|---|---|---|---|
| B121 S&P 500 arithmetic | 1928-2025 | 0.118547 | 0.118547 |
| B126 S&P 500 geometric | 1928-2025 | 0.100177 | 0.100177 |
| L122 / M122 ERP arith (bills / bonds) | 1928-2025 | 0.084397 / 0.070323 | same |
| L127 / M127 ERP geometric (bills / bonds) | 1928-2025 | 0.066461 / 0.054831 | same |
| M128 / M129 ERP geometric vs bonds | 1976-2025 / 2016-2025 | 0.057917 / 0.137866 | same |
| V126 / W126 real geometric S&P / small cap | 1928-2025 | 0.067763 / 0.086825 | same |
| AH124 real stocks-bills std error | 1976-2025 | 0.021924 | 0.021924 |
| T118, E17 'Historical ERP' / Damodaran's estimator | 1928-2025 | 0.054831 | 0.054831 |
| histimpl C78 / D78 avg implied ERP / + T.Bond | 1960-2025 | 0.042477 / 0.099268 | same |
| histimpl C80 / D80 | 2016-2025 | 0.04999 / 0.07841 | same |

The full set covers all 7 series x arithmetic/geometric x nominal/real x 3 periods, every premium and std-error cell,
and T for six years.

**Source-sheet errors (11 cells; ours is the corrected figure):**
- C126:C128 and Q127:Q129 (small-cap geometric return and premium). The sheet's column J is `=100*(1+C)` in every row instead of
  cumulating, so the sheet's 1928-2025 small-cap geometric figure is 0.16%. Ours is 11.98%. Its real column AD does
  cumulate, and ours matches its real small-cap geometric 8.68%. `history` prints a NOTE with the sheet's figure for the
  window, computed the sheet's way (JSON `his_sheet`); the selftest checks those 6 cells against the sheet, so they count as cases.
- N122, R122 (STDEV of 98 years / sqrt(97)) and AH123, AI123 (97 years / sqrt(98)), printed as `src-err`. Ours uses n
  on both sides; the sheet's figure = ours x sqrt(98/97) for N122/R122, STDEV(1929-2025)/sqrt(98) for AH123/AI123.
  The other std-error cells (N/O/R123-124, O122) actually cover 1975-2024 / 2015-2024 / 1928-2024 although labelled
  -2025, and they match exactly when run over those years. A request for "1976-2025" gets the consistent 1976-2025 number.
- histimpl D81 `=C82+K72` adds the 2025 ERP to the 2024 bond rate (`src-err`; ours 4.33% + 4.58%). B82 is labelled "End of 2024" but links to 2025.
- EVAGlobal.xls has a stale date cell (2025-01-05). Its firm counts and betas equal the Jan 2026 files, so as_of is set
  to 2026-01-05. histimpl's header also reads 2024-01-05, but its data runs through 2025.
- ERPbymonth text cells: '6,36%' (L214) is stored as 0.0636, '84,88' (G115) as 84.88 and '175.51*' as 175.51.
  'Ended' (O162) is stored as null.

## Damodaran's conventions
- The xls datasets hold aggregates: a sum over the industry, not a median. Medians and quartiles exist only in the blog files
  (IndustryMargins, IndustryReturns, DebtIndustry, DividendsIndustry, WACCdistn) and only for US and Global.
  For "median" answers, name the file. For other regions, say that only the aggregate exists.
- Premium = difference of averages. The geometric ERP is geo(stocks) - geo(bonds), not the geometric mean of annual
  differences. Real return = (1+r)/(1+inflation) - 1. Std error = STDEV(annual premium)/sqrt(n).
  Damodaran quotes the geometric stocks-minus-T.Bonds premium as the "historical ERP".
- Small Cap Premium = bottom-decile return minus the S&P 500 (the sheet's column R formula; the header text says the reverse).
- Implied ERP year t = value at the end of t (start of t+1). 'Implied ERP (FCFE)' is the headline, and the monthly
  headline is 'ERP (T12m)'. The sheet's period average skips years without an ERP, so 1960-2025 averages 1961-2025.
- Column names are as in the source files. Repeated or bare-year headers get the group header as a prefix ('Return on Equity: Median',
  'Unlevered beta corrected for cash - Over time: 2024', 'Cash Tax Rates: Aggregate tax rate').

## Input reference
`datasets` · `columns <dataset>` (names, Damodaran's definitions and settings) · `industry [name] [--region R] [--dataset D]
[--metric substring]` (no name = list the industries) · `rank <metric|dataset:column> [--region R] [--top N] [--asc]`
(Total Market rows are shown as reference and not ranked; an exact name found in several files resolves in the order
betas, wacc, margin, EVA, dbtfund, debtdetails, divfund, divfcfe, taxrate, mgnroc, then the blog files) ·
`history [--start Y --end Y | --year Y]` · `erp [--year Y | --start Y --end Y | --month YYYY-MM | --monthly
--start YYYY-MM --end YYYY-MM]` · `company <name or ticker>` · `--json` on any command.
Regions: US, Global, Europe, Japan, Emerging, China, India. Names match case-insensitively: an exact match first, else
a unique substring. An ambiguous name returns the list of candidates.

## Example (real output)
```
python scripts/damodata.py history --start 1976 --end 2025
  S&P 500 (includes dividends)  arith 13.17%  geom 11.92%  real 9.27% / 8.04%
  Stocks - T.Bonds              arith 6.56%   geom 5.79%   std err 2.62%
python scripts/damodata.py rank "Cost of Capital" --top 3   ->  Software (Internet) 0.1066, Semiconductor 0.1055,
  Semiconductor Equip 0.0989 ... Total Market 0.0696  [wacc.xls | as of 2026-01-05]
python scripts/damodata.py industry Semiconductor --region Global --metric margin
python scripts/damodata.py erp --start 2016 --end 2025   ->  average 4.999%, + T.Bond 7.841%
python scripts/damodata.py company NVDA                  ->  Semiconductor, United States
```

## Limits
- It is a snapshot of Damodaran's Jan 2026 files (Jan 2025 for WACCdistn; ERP monthly through Sep 2026). To refresh, download
  the files again and run `.venv\Scripts\python.exe scripts\damodata.py --build-cases ..\..\spreadsheets\*.xls*`.
- Not stored: IndustryMargins raw-sum sheets, indname Primary Sector/SIC, and the older ERP/industry files. The
  cost-of-debt lookup tables inside wacc/EVA are also left out; they are inputs already applied in the datasets.
- No company-level numbers: indname only maps a company to an industry and country.
- The data file is 3.7 MB, 2.4 MB of which is the company map.
