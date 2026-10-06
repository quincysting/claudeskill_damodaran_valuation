---
name: aswath-finance
description: Aswath Damodaran's finance knowledge from all 682 "Musings on Markets" posts (2008-2026) with post-level citations, plus deterministic engines transcribed from and verified against Damodaran's spreadsheets (1,300+ checks) - company valuation (Ginzu, 31 published valuations; exports a live Excel model), banks, cost of capital, implied ERP, country risk, optimal debt ratio and buybacks, M&A synergy and control, fundamental multiples, market value and breakeven revenue, private-firm and distress discounts, options and real options - and current Damodaran datasets (industry betas, margins, costs of capital; historical returns; ERP history). Use when valuing a company, bank, IPO or young business the Damodaran way (incl. an Excel model), estimating discount rates, ERP or country risk premiums, analyzing capital structure, payout or acquisitions, pricing on multiples, judging if a stock, market or asset is over- or undervalued, or when the user mentions Damodaran, Aswath, Musings on Markets, or value vs price.
---

# Aswath Finance

Damodaran's frameworks, numbers, cases and tools: *Musings on Markets* (682 posts, Sep 2008 – Sep 2026) distilled with a
citation on every claim, plus the source spreadsheets turned into tested code.

## How to answer

1. **Numbers come from engines and Damodaran's data, never from memory.** Any figure the answer computes or looks up (a value,
   ERP, CRP, beta, cost of capital, optimal debt ratio, synergy, justified multiple, discount, option value, industry
   statistic, historical return) must come from a `scripts/` engine or `damodata.py`. Read the matching `tool-*.md`
   first. Engines keep the spreadsheet defaults, print sources and as-of dates; `python scripts/check_all.py` verifies them all.
2. **Reasoning comes from the references.** Route with the tables below, grep before reading whole files, and cite posts
   inline: `([title](url), YYYY-MM)`. Every numeric claim there was checked against its post (`scripts/audit.py`);
   before quoting a figure the answer hinges on, grep the cited post.
3. Date every number and say whose it is: Damodaran's data/defaults/prior inputs vs judgment (labeled, tied to a cited post).
4. Attribute, don't impersonate ("Damodaran argues…"); flag where Damodaran's view changed (the "views evolved" section of each reference).
5. **Source-sheet errors**: where a source spreadsheet contradicts Damodaran's own method (documented in each `tool-*.md`), the
   engine uses the corrected math by default, prints the discrepancy, and `replicate_sheet_bugs` reproduces the sheet's figure.

## Engines (read the tool doc first)

| Question | Engine | Doc |
|---|---|---|
| Value a company (FCFF, young/segment/IPO firms), start from Damodaran's inputs | `ginzu.py` (`--from-case`) | [valuation-engine.md](references/valuation-engine.md) |
| Share a valuation as an Excel model (live formulas, scenarios, data tables, checks tied to the engine) | `export_xlsx.py` (`--verify`) | [tool-excel-export.md](references/tool-excel-export.md) |
| Value a bank, insurer or financial-service firm; DDM/FCFE/excess-return models | `financial_firms.py` | [tool-financial-firms.md](references/tool-financial-firms.md) |
| Cost of capital (industry beta, ERP by country/region, rating spread) | `cost_of_capital.py` | [valuation-engine.md](references/valuation-engine.md) |
| Implied ERP, Damodaran's published ERPs, intrinsic riskfree, currency riskfree | `implied_erp.py` | [tool-implied-erp.md](references/tool-implied-erp.md) |
| Country ERP/CRP/default spread/tax rate; revenue-weighted ERP | `country_risk.py` | [tool-country-risk.md](references/tool-country-risk.md) |
| Optimal debt ratio, buybacks, FCFE cash-return capacity, relevered beta, synthetic rating | `capital_structure.py` | [tool-capital-structure.md](references/tool-capital-structure.md) |
| Value of control, synergy, most an acquirer can pay | `mna.py` | [tool-mna.md](references/tool-mna.md) |
| Fundamentals-justified PE/PBV/PS/EV multiples, pricing vs peers | `multiples.py` | [tool-multiples.md](references/tool-multiples.md) |
| S&P 500 value and implied ERP; breakeven revenue for a market cap; PE history | `market.py` | [tool-market.md](references/tool-market.md) |
| Illiquidity/minority discounts, total beta, failure probability, distressed equity | `private_distress.py` | [tool-private-distress.md](references/tool-private-distress.md) |
| Warrants/employee options, Black-Scholes, real options (delay, expand, abandon), patents, reserves | `options.py` | [tool-options.md](references/tool-options.md) |
| Industry statistics by region, historical returns and premiums, ERP history, company → industry | `damodata.py` | [tool-damodata.md](references/tool-damodata.md) |

Valuing a company: start from Damodaran's own inputs if a published valuation exists (`ginzu.py --from-case <Name>`), update the base year from
filings, compute cost of capital from Damodaran's data, map story → drivers with a `_sources` entry for every input, keep the Ginzu
defaults unless a cited reason says otherwise, run it, answer the Ginzu Diagnostics, show `--breakeven`. Never hand-roll a DCF.

## Knowledge (reasoning, history, cases)

| Question about | Read |
|---|---|
| DCF mechanics: FCFF/FCFE, growth, reinvestment, terminal value, taxes, accounting fixes, options/SBC | [dcf-valuation.md](references/dcf-valuation.md) |
| ERP, riskfree rate, beta, cost of capital, hurdle rates, interest rates, inflation, the Fed | [discount-rates-macro.md](references/discount-rates-macro.md) |
| Country risk, currency choice, emerging markets, sovereign ratings, geopolitics | [country-risk.md](references/country-risk.md) |
| Damodaran's data updates: current benchmarks 2024–2026; methods and history 2009–2023 | [data-latest.md](references/data-latest.md), [data-history.md](references/data-history.md) |
| Debt, capital structure, dividends vs buybacks, cash, banks, distress | [capital-structure-payout.md](references/capital-structure-payout.md) |
| Corporate life cycle, M&A, CEOs, governance, ESG | [governance-lifecycle-ma.md](references/governance-lifecycle-ma.md) |
| Story → numbers, young firms, VC, IPOs/SPACs, big-market delusion, platforms, AI | [narrative-young-companies.md](references/narrative-young-companies.md) |
| Value vs price, multiples, bubbles, crises, market timing, crypto/gold/collectibles | [pricing-markets.md](references/pricing-markets.md) |
| Investment philosophy, value vs growth, active vs passive, behavioral finance, earnings | [investment-philosophy.md](references/investment-philosophy.md) |
| A company Damodaran valued (each file opens with a company index) | [cases-2008-2012.md](references/cases-2008-2012.md), [cases-2013-2017.md](references/cases-2013-2017.md), [cases-2018-2026.md](references/cases-2018-2026.md) |
| The source post for anything | [post-index.md](references/post-index.md) (all posts: date, title, labels) |

## Core through-lines

- **Value ≠ price.** Value comes from cash flows, growth and risk; price from mood, momentum and liquidity.
- **A valuation is a story told with numbers.** Test it (possible → plausible → probable), convert to inputs, revise when the narrative breaks, changes or shifts.
- **Growth adds value only when return on capital beats cost of capital.** A good company is not automatically a good investment.
- **Consistency beats precision.** Currency, inflation, riskfree rate, ERP and growth share one basis; stable growth ≤ riskfree; terminal reinvestment = g / ROC.
- **Use the implied (forward-looking) ERP**, not historical averages.
- **Country risk follows operations, not domicile.** Discrete risks (nationalization, war) are probability adjustments, not discount-rate add-ons.
- **The life-cycle stage decides what matters.** Young: story and survival. Mature: discipline. Declining: shrink and return cash.
- **Buybacks and dividends return cash; they don't create value.** Debt trades tax benefits against distress costs.
- **Markets are efficient enough for most investors.** Pick a philosophy that fits you; humility beats conviction plus leverage.

## Going deeper and staying current

- Answering needs only this folder. Maintenance uses an optional workspace around it (called `finance/` in these docs; see the repository README): the skill at `finance/skill/aswath-finance`, Damodaran's source spreadsheets in `finance/spreadsheets`, full post text in `finance/markdown/<YYYY-MM-slug>.md`.
- References are a snapshot through the 2026-09-10 post; engine data is from Damodaran's Jan–Jul 2026 files. Refresh steps are in the repository README and each tool doc (`--build-cases` re-verifies against the new source files).
- Damodaran's public work as of each post or file date, for education — not investment advice.
