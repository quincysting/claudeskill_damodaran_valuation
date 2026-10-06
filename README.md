# Aswath Finance: a Claude skill for valuing companies the Damodaran way

A [Claude skill](https://docs.anthropic.com/en/docs/agents-and-tools/agent-skills/overview) that answers corporate
finance and valuation questions with Aswath Damodaran's frameworks, numbers and spreadsheets. It has two parts:

- **Knowledge**: all 682 posts of Damodaran's blog *Musings on Markets* (Sep 2008 – Sep 2026), distilled into reference files
  with a citation on every claim. An audit script checks each numeric claim and quote against the source post.
- **Engines**: Damodaran's spreadsheets transcribed into tested Python. They value a company with the Ginzu FCFF model, estimate
  the cost of capital, ERP and country risk, analyse capital structure, M&A, multiples, options and distress, and
  export a live-formula Excel model. Every engine replays cases taken from the source spreadsheets.

Ask Claude something like *"Value Nvidia the Damodaran way and give me the spreadsheet"*. The skill tells Claude to
start from Damodaran's last published inputs, update the base year from filings, compute the cost of capital from
Damodaran's data, run the engine (never a hand-rolled DCF) and answer the Ginzu Diagnostics questions.

> **Disclaimer.** This is an independent project. It is not affiliated with, reviewed or endorsed by Aswath
> Damodaran or NYU Stern. It summarises Damodaran's public blog posts (with links) and uses data from the public website
> (pages.stern.nyu.edu/~adamodar). It is for education only and is not investment advice.

## Repository layout

```
skill/aswath-finance/        the skill: copy this folder into your skills directory
  SKILL.md                   entry point Claude reads (routing, rules, engine table)
  references/                knowledge files + one tool doc per engine
  scripts/                   the engines (Python), the Excel exporter and the test harness
  data/                      Damodaran's data tables (Jan-Jul 2026 files) and test cases from the source spreadsheets
valuation/                   three sample valuations built with the skill (see below)
  SPCX/  NVDA/  META/        inputs with sources, engine output, report and Excel workbook
```

The docs call the repository root the `finance/` workspace. Maintenance folders that are not included
(`spreadsheets/`, `markdown/`, `.venv/`) would sit next to `skill/` and `valuation/`.

## Install

**Claude Code** (personal skill, available in every project):

```bash
git clone https://github.com/quincysting/claudeskills_damodaran_valuation.git
cp -r claudeskills_damodaran_valuation/skill/aswath-finance ~/.claude/skills/
```

On Windows (PowerShell): `Copy-Item -Recurse claudeskills_damodaran_valuation\skill\aswath-finance $HOME\.claude\skills\`.
For a single project, copy the folder into `<project>/.claude/skills/` instead. Start a new session (or run
`/reload-skills`) and the skill triggers on valuation questions or a mention of Damodaran.

**Claude apps / API**: zip the `aswath-finance` folder and upload it as a custom skill where your plan supports it.
The engines need a code-execution environment with Python.

### Requirements

| Use | Needs |
|---|---|
| All engines (`ginzu.py`, `cost_of_capital.py`, ...) | Python 3.10+, standard library only |
| Excel export (`export_xlsx.py`) | `pip install openpyxl` |
| Exporter selftest (recalculates 66 workbooks) | `pip install pycel` |
| `--verify` (recalculate in real Excel) | Windows with Microsoft Excel |

## Using it

In Claude, just ask:

- *"Value Meta the Damodaran way."*
- *"What is Damodaran's current equity risk premium for Brazil?"*
- *"What would Nvidia need in revenue to justify its market cap?"*
- *"Build me the valuation spreadsheet so I can share it with my team."*

From the command line (inside `skill/aswath-finance/scripts`):

```bash
python ginzu.py --from-case Apple2024 > aapl.json         # Damodaran's inputs for one of 31 published valuations
python ginzu.py aapl.json --breakeven target_margin       # value, year-by-year table, Diagnostics, breakeven
python export_xlsx.py aapl.json --verify                  # live Excel model with a Checks tab tied to the engine
python country_risk.py --country Brazil                    # ERP, CRP, default spread (July 2026 Damodaran data)
python check_all.py                                        # every engine's selftest + the reference audit
```

Each engine's options are documented in `references/tool-*.md` and in the docstring at the top of the script.

### What's inside

| Question | Engine |
|---|---|
| Value a company (FCFF; young, segment and IPO firms) | `ginzu.py` |
| Share the valuation as an Excel model | `export_xlsx.py` |
| Value a bank or insurer (DDM, FCFE, excess returns) | `financial_firms.py` |
| Cost of capital (industry beta, ERP by country, rating spread) | `cost_of_capital.py` |
| Implied ERP, published Damodaran ERPs, currency riskfree rates | `implied_erp.py` |
| Country ERP / CRP / default spread / tax rate | `country_risk.py` |
| Optimal debt ratio, buybacks, synthetic rating | `capital_structure.py` |
| Value of control, synergy, most an acquirer can pay | `mna.py` |
| Fundamentals-justified multiples | `multiples.py` |
| S&P 500 value, breakeven revenue for a market cap | `market.py` |
| Illiquidity and minority discounts, distress | `private_distress.py` |
| Employee options, warrants, real options | `options.py` |
| Industry statistics, historical returns, ERP history | `damodata.py` |

## How it is verified

`python scripts/check_all.py` runs every engine's selftest against cases taken from Damodaran's own spreadsheets:

| Engine | Cases |
|---|---|
| `ginzu.py` | 45/45 (31 published valuations + 14 template branches) within 0.05% |
| `cost_of_capital.py` | 21/21 within 1e-6 |
| `export_xlsx.py` | 66/66 workbooks recalculate (pycel) to the engine |
| `damodata.py` | 536/536 |
| `multiples.py` / `mna.py` / `implied_erp.py` | 198/198, 132/132, 108/108 |
| `private_distress.py` / `financial_firms.py` / `market.py` | 86/86, 81/81, 64/64 |
| `options.py` / `capital_structure.py` / `country_risk.py` | 54/54, 41/41, 7/7 |
| `audit.py` (references vs posts) | 99.7% of numeric claims found in the cited post; 89/89 quotes verbatim |

Where a source spreadsheet contradicts Damodaran's own method, the engine uses the corrected math by default and
`replicate_sheet_bugs` reproduces the sheet's figure; each case is documented in the tool docs. Without `pycel` the
exporter selftest reports SKIP; without the post corpus (not distributed) the audit reports SKIP.

## Sample valuations (`valuation/`)

| Company | As of | Value / share | Price | Notes |
|---|---|---|---|---|
| SpaceX (SPCX) | 2026-09-25 | $81.40 | $148.68 | Damodaran's June-2026 segment model rolled forward; reproduces its $97.83 first |
| NVIDIA (NVDA) | 2026-10-02 | $192.41 | $233.95 | The Jan-2025 Damodaran story re-based to an AI market already larger than its 2035 forecast |
| Meta (META) | 2026-10-02 | $757.32 | $728.08 | The Feb-2024 Damodaran story; reinvestment re-anchored on 2026 AI capital spending |

Each folder has the inputs (with a `_sources` entry per input: filing, market data, Damodaran's data or labelled
judgment), the engine output, a written report with scenarios and breakevens, and the Excel workbook. Reproduce one
from its folder, for example `cd valuation/NVDA && python ../../skill/aswath-finance/scripts/ginzu.py oct_2026_inputs.json`.

These are point-in-time illustrations of the method, not recommendations. Analyst consensus figures in them are
third-party data (S&P Global via StockAnalysis), labelled as such.

## Data vintage and refreshing

- References: a snapshot through the 2026-09-10 post. Engine data: Damodaran's January–July 2026 files.
- To refresh, download Damodaran's updated spreadsheets from pages.stern.nyu.edu/~adamodar into `spreadsheets/` next to
  `skill/`, then rebuild each engine's data and cases with its `--build-cases` (or `calibrate.py`) step; each tool doc
  has a "Refreshing" section. The rebuild tools need `openpyxl`, `xlrd` and `pycel`.
- The reference audit needs the post corpus (one markdown file per post) in `markdown/` next to `skill/`, or set
  `ASWATH_CORPUS`.

## Credits

All frameworks, data and published valuations are Aswath Damodaran's, from the blog
[Musings on Markets](https://aswathdamodaran.blogspot.com) and the NYU Stern website. Read the original posts; the
reference files link to every one of them.
