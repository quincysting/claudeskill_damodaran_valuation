"""Relative valuation Damodaran's way: the multiple a firm's fundamentals justify, and pricing against peers.

"Every multiple embeds assumptions about growth, risk and reinvestment": derive it from a DCF, then compare
(pricing-markets.md, 2018-02 / 2019-02). Transcribed from Damodaran's spreadsheets:
  eqmult.xls    https://pages.stern.nyu.edu/~adamodar/pc/eqmult.xls    Sheet1 C4:D31 (saved Jan 2025)
      two-stage DDM per $1 of next year's net income -> PE (forward), PBV (book = NI next year / ROE),
      PS (forward, trailing); ROE = g / (1 - payout), cost of equity = rf + beta x ERP (stable ERP = ERP, D8)
  firmmult.xls  https://pages.stern.nyu.edu/~adamodar/pc/firmmult.xls  Sheet1 C4:F32 (saved Jan 2025)
      two-stage FCFF per $1 of next year's EBIT(1-t) -> EV/EBIT(1-t), EV/EBIT, EV/Invested capital (ROIC = g/RR),
      EV/Sales (after-tax operating margin); trailing = forward x (1+g)
  EqMultCalculator.xlsx  https://pages.stern.nyu.edu/~adamodar/pc/blog/EqMultCalculator.xlsx  (2021-03 post)
      life_cycle preset: g = rf + {Growth 10%, Mature 1%, Declining -5%}, cost of equity = rf + 5% in both stages,
      stable growth = rf, payout = 1 - g/ROE; value on current EPS; PBV = ROE-weighted PV / EPS (C21)
  AlibabaRelValue.xls    https://pages.stern.nyu.edu/~adamodar/pc/blog/AlibabaRelValue.xls   (2014-09 post)
      pricing from peers: equity value = PE x net income, PBV x book; EV multiples x metric + cash + other holdings
      - debt; invested capital = book equity + debt - cash
Industry context: Damodaran's Jan 2026 industry averages (data/ginzu_data.json, via cost_of_capital.py).

Formulas (g, r: high growth; gs, rs: stable; n years):
  PE_fwd  = payout (1 - ((1+g)/(1+r))^n)/(r-g) + (1+g)^(n-1) (1+gs) payout_s / ((rs-gs)(1+r)^n)
  EV/NOPAT_fwd = same with (1 - reinvestment rate) in place of payout and the cost of capital in place of r
  n = 0 gives the stable-growth forms, e.g. PBV = (ROE - g)/(r - g) (Deutsche Bank post, 2016-10).
At g = r the spreadsheet's IF branch returns n x earnings (no payout, no discounting); we use the limit of the main formula
(a NOTE is printed); "replicate_sheet_bugs": true reproduces that IF branch.

Usage:
  python multiples.py inputs.json [--json] [--solve FIELD:MULTIPLE]   e.g. --solve cost_of_equity:pbv solves the
                                     input that makes the justified multiple equal actual[MULTIPLE] (implied value)
  python multiples.py --selftest     replay data/multiples_cases.json
  python multiples.py --build-cases FILES...   (maintenance venv: xlrd/openpyxl/pycel; give X.xls with its Excel
                                     Save-As X.xlsx for eqmult/firmmult to add pycel branch tests)

Inputs (decimals):
  mode "equity": years, growth, payout | roe, cost_of_equity | riskfree_rate + beta + erp, stable_growth (default
      riskfree_rate), stable_payout | stable_roe (default roe), stable_cost_of_equity | stable_beta (+ stable_erp,
      stable_riskfree_rate; defaults erp, riskfree_rate) (default: cost_of_equity), net_margin (PS), eps (current,
      -> value per share), or life_cycle "Growth"|"Mature"|"Declining" + riskfree_rate + roe + years
  mode "firm": years, growth, reinvestment_rate | roic, cost_of_capital, stable_growth, stable_reinvestment_rate |
      stable_roic, stable_cost_of_capital, after_tax_operating_margin (EV/Sales), tax_rate (EV/EBIT),
      da_to_ebitda (optional: EV/EBITDA = EV/EBIT x (1 - D&A/EBITDA), an identity, not in the source sheet)
  both: actual {multiple: value} -> actual / justified; industry (Damodaran's industry name), scope "US"|"Global"
  mode "comps": firm {revenues, ebitda, ebit, net_income, book_equity, debt, cash, other_holdings, shares,
      market_cap}, peers {group: {"PE"|"PBV"|"PS"|"EV/EBITDA"|"EV/EBIT"|"EV/Sales"|"EV/Invested Capital": x or [x...]}},
      statistic "median" (the spreadsheet default) | "average" for lists
"""

from __future__ import annotations

import argparse
import json
import statistics
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
CASES = HERE.parent / "data" / "multiples_cases.json"
TOOLS, BLOG = "https://pages.stern.nyu.edu/~adamodar/pc/", "https://pages.stern.nyu.edu/~adamodar/pc/blog/"
LIFE_CYCLE = {"Declining": (-0.05, 0.05), "Growth": (0.10, 0.05), "Mature": (0.01, 0.05)}  # EqMultCalculator F15:H17
METRIC = {"PE": ("net_income", False), "PBV": ("book_equity", False), "PS": ("revenues", False),
          "EV/EBITDA": ("ebitda", True), "EV/EBIT": ("ebit", True), "EV/Sales": ("revenues", True),
          "EV/Invested Capital": ("invested_capital", True)}
DRIVERS = ["Trailing PE", "Price/Book", "EV/Sales", "EV/EBITDA", "EV/EBIT", "ROE", "Dividend Payout Ratio",
           "Annual Average Revenue growth - Last 5 years", "Pre-tax Operating Margin (Unadjusted)", "After-tax ROC",
           "Reinvestment Rate", "Cost of equity", "Cost of capital", "Average effective tax rate"]


def two_stage(first: float, g: float, r: float, n: float, stable: float, gs: float, rs: float,
              his_if: bool = False) -> tuple[float, float]:
    """(PV of the high-growth payouts, PV of the terminal value) per $1 of next year's earnings.
    his_if: at g = r reproduce the sheet's IF branch (n x earnings: no payout, no discounting) instead of the limit."""
    ann = (1 - ((1 + g) / (1 + r)) ** n) / (r - g) if r != g else n / (1 + r)
    hi = n if r == g and his_if else first * ann
    return hi, (1 + g) ** (n - 1) * (1 + gs) * stable / ((rs - gs) * (1 + r) ** n)


GR_NOTE = ("g = r: the source sheet's IF(g=r) branch returns n x earnings (no payout, no discounting); this uses the limit "
           "of the sheet's main formula. Set replicate_sheet_bugs=true to reproduce the source sheet")


def equity(p: dict) -> dict:
    p = dict(p)
    if p.get("life_cycle"):
        diff, prem = LIFE_CYCLE[p["life_cycle"]]
        p.setdefault("growth", p["riskfree_rate"] + diff)
        p.setdefault("cost_of_equity", p["riskfree_rate"] + prem)
    g, n = p["growth"], p["years"]
    r = p["cost_of_equity"] if "cost_of_equity" in p else p["riskfree_rate"] + p["beta"] * p["erp"]
    payout = p["payout"] if "payout" in p else 1 - g / p["roe"]
    roe = p.get("roe", g / (1 - payout) if payout != 1 else float("nan"))
    gs = p.get("stable_growth", p.get("riskfree_rate"))
    if "stable_cost_of_equity" in p:
        rs = p["stable_cost_of_equity"]
    elif "stable_beta" in p:
        rs = p.get("stable_riskfree_rate", p["riskfree_rate"]) + p["stable_beta"] * p.get("stable_erp", p["erp"])
    else:
        rs = r
    payout_s = p["stable_payout"] if "stable_payout" in p else 1 - gs / p.get("stable_roe", roe)
    roe_s = p.get("stable_roe", gs / (1 - payout_s) if payout_s != 1 else float("nan"))
    hi, term = two_stage(payout, g, r, n, payout_s, gs, rs, p.get("replicate_sheet_bugs", False))
    pe = hi + term
    out = {"cost_of_equity": r, "stable_cost_of_equity": rs, "growth": g, "stable_growth": gs, "payout": payout,
           "stable_payout": payout_s, "roe": roe, "stable_roe": roe_s, "pe_forward": pe, "pe_trailing": pe * (1 + g),
           "pbv": pe * roe, "pbv_roe_trailing": (roe * hi + roe_s * term) * (1 + g),
           "notes": [GR_NOTE] if g == r and not p.get("replicate_sheet_bugs") else []}
    if "net_margin" in p:
        out.update(ps_forward=pe * p["net_margin"], ps_trailing=pe * p["net_margin"] * (1 + g))
    if "eps" in p:
        out["value_per_share"] = out["pe_trailing"] * p["eps"]
    return out


def firm(p: dict) -> dict:
    g, n, r = p["growth"], p["years"], p["cost_of_capital"]
    rr = p["reinvestment_rate"] if "reinvestment_rate" in p else g / p["roic"]
    roic = p.get("roic", g / rr if rr else float("nan"))
    gs, rs = p["stable_growth"], p["stable_cost_of_capital"]
    rr_s = p["stable_reinvestment_rate"] if "stable_reinvestment_rate" in p else gs / p["stable_roic"]
    ev = sum(two_stage(1 - rr, g, r, n, 1 - rr_s, gs, rs, p.get("replicate_sheet_bugs", False)))
    out = {"reinvestment_rate": rr, "roic": roic, "stable_reinvestment_rate": rr_s,
           "stable_roic": gs / rr_s if rr_s else float("nan"),
           "ev_ebit_after_tax_forward": ev, "ev_ebit_after_tax_trailing": ev * (1 + g),
           "ev_invested_capital": ev * roic,
           "notes": [GR_NOTE] if g == r and not p.get("replicate_sheet_bugs") else []}
    if "tax_rate" in p:
        out.update(ev_ebit_forward=ev * (1 - p["tax_rate"]), ev_ebit_trailing=ev * (1 - p["tax_rate"]) * (1 + g))
        if "da_to_ebitda" in p:
            out.update(ev_ebitda_forward=out["ev_ebit_forward"] * (1 - p["da_to_ebitda"]),
                       ev_ebitda_trailing=out["ev_ebit_trailing"] * (1 - p["da_to_ebitda"]))
    if "after_tax_operating_margin" in p:
        m = p["after_tax_operating_margin"]
        out.update(ev_sales_forward=ev * m, ev_sales_trailing=ev * m * (1 + g))
    return out


def comps(p: dict) -> dict:
    f = dict(p["firm"])
    if "book_equity" in f:
        f.setdefault("invested_capital", f["book_equity"] + f.get("debt", 0.0) - f.get("cash", 0.0))
    stat = statistics.median if p.get("statistic", "median") == "median" else statistics.mean
    nonop = f.get("cash", 0.0) + f.get("other_holdings", 0.0) - f.get("debt", 0.0)
    res = {}
    for group, mults in p["peers"].items():
        for m, x in mults.items():
            key, is_ev = METRIC[m]
            if key in f:
                x = stat(x) if isinstance(x, list) else x
                res.setdefault(group, {})[m] = x * f[key] + (nonop if is_ev else 0.0)
    return {"comps": res}


def mode(p: dict) -> str:
    if "mode" in p:
        return p["mode"]
    return "comps" if "peers" in p else "firm" if "cost_of_capital" in p else "equity"


def run(p: dict) -> dict:
    out = {"equity": equity, "firm": firm, "comps": comps}[mode(p)](p)
    act = p.get("actual", {})
    out["actual_vs_justified"] = {k: v / out[k] for k, v in act.items() if k in out and out[k]}
    return out


def solve(p: dict, field: str, mult: str) -> float | None:
    """Value of `field` at which the justified multiple equals actual[mult] (grid scan + bisection)."""
    target = p["actual"][mult]

    def f(x: float) -> float | None:
        try:
            return run({**p, field: x})[mult] - target
        except (ZeroDivisionError, ValueError, OverflowError):
            return None
    def bisect(lo: float, flo: float, hi: float) -> float | None:
        for _ in range(100):
            mid = (lo + hi) / 2
            fm = f(mid)
            if fm is None:
                return None
            lo, flo, hi = (mid, fm, hi) if flo * fm > 0 else (lo, flo, mid)
        x = (lo + hi) / 2
        fx = f(x)
        return x if fx is not None and abs(fx) < 1e-6 * max(1.0, abs(target)) else None  # a pole is not a root

    prev = None
    for k in range(601):
        x = -0.3 + 0.9 * k / 600
        y = f(x)
        if y is not None and prev is not None and prev[1] * y <= 0:
            root = bisect(prev[0], prev[1], x)
            if root is not None:
                return root
        prev = None if y is None else (x, y)
    return None


def report(p: dict, out: dict) -> str:
    rows = []
    if "comps" in out:
        f = p["firm"]
        for group, vals in out["comps"].items():
            rows.append(group)
            for m, v in vals.items():
                extra = f"  per share {v / f['shares']:,.2f}" if f.get("shares") else ""
                extra += f"  market cap / value {f['market_cap'] / v:.0%}" if f.get("market_cap") else ""
                rows.append(f"   {m:22s} equity value {v:14,.1f}{extra}")
        rows.append("RULE: 'most stocks that look cheap deserve to be cheap' - check the driver (growth, risk, ROE, "
                    "margin) before calling a low multiple a bargain (2019-02)")
        return "\n".join(rows)
    for k, v in out.items():
        if isinstance(v, float):
            fmt = "{:,.4f}" if any(s in k for s in ("pe", "pbv", "ps_", "ev_")) else "{:.2%}"
            rows.append(f"{k:28s} {fmt.format(v)}")
    rows += [f"NOTE: {n}" for n in out.get("notes", [])]
    for k, v in out["actual_vs_justified"].items():
        rows.append(f"ACTUAL {k:21s} {p['actual'][k]:,.4f} = {v:.0%} of justified"
                    f" ({'above' if v > 1 else 'below'} what these fundamentals support)")
    if p.get("industry"):
        from cost_of_capital import industry, load
        ind = industry(load(), p["industry"], p.get("scope", "US"))
        rows.append(f"\nINDUSTRY {p['industry']} ({p.get('scope', 'US')}, Damodaran's Jan 2026 data)")
        rows += [f"   {d:46s} " + (f"{ind[d]:,.4f}" if isinstance(ind[d], (int, float)) else str(ind[d]))
                 for d in DRIVERS if ind.get(d) is not None]
    return "\n".join(rows)


# ---------------------------------------------------------------- self test
def pick(d: dict, path: str) -> float:
    for k in path.split("|"):
        d = d[k]
    return d


def selftest() -> int:
    bad = total = 0
    for c in json.loads(CASES.read_text(encoding="utf-8")):
        out = run(c["inputs"])
        for key, his in c["expected"].items():
            ours = pick(out, key)
            ok = abs(ours - his) <= max(5e-4 * abs(his), 1e-6)
            bad, total = bad + (not ok), total + 1
            print(f"{'ok ' if ok else 'BAD'} {c['name']:44s} {key:36s} Damodaran {his:14,.6f}  ours {ours:14,.6f}")
        for key, (his, why) in c.get("unmatched", {}).items():
            print(f"n/a {c['name']:44s} {key:36s} Damodaran {his:14,.6f}  ours {pick(out, key):14,.6f}  ({why})")
    print(f"{total - bad}/{total} checks within 0.05% of Damodaran's spreadsheets")
    return 1 if bad else 0


# ---------------------------------------------------------------- case builder (maintenance only)
def _col(a: str) -> tuple[int, int]:
    i = next(k for k, ch in enumerate(a) if ch.isdigit())
    c = 0
    for ch in a[:i]:
        c = c * 26 + ord(ch.upper()) - 64
    return int(a[i:]) - 1, c - 1


def _reader(path: Path):
    if path.suffix.lower() == ".xls":
        import xlrd
        b = xlrd.open_workbook(str(path))
        return lambda sh, a: b.sheet_by_name(sh).cell_value(*_col(a))
    import openpyxl
    w = openpyxl.load_workbook(str(path), data_only=True)
    return lambda sh, a: w[sh][a].value


def x_eqmult(get) -> tuple[dict, dict]:
    v = lambda a: get("Sheet1", a)
    keys = {"years": "C4", "growth": "C5", "payout": "C6", "riskfree_rate": "C7", "erp": "C8", "beta": "C9",
            "net_margin": "C10", "stable_growth": "D5", "stable_payout": "D6", "stable_riskfree_rate": "D7",
            "stable_erp": "D8", "stable_beta": "D9"}
    exp = {"pe_forward": "D28", "pbv": "D29", "ps_forward": "D30", "ps_trailing": "D31", "cost_of_equity": "D21",
           "stable_cost_of_equity": "D25", "roe": "G19"}
    return {"mode": "equity", **{k: v(a) for k, a in keys.items()}}, {k: v(a) for k, a in exp.items()}


def x_firmmult(get) -> tuple[dict, dict]:
    v = lambda a: get("Sheet1", a)
    keys = {"years": "C4", "growth": "C5", "reinvestment_rate": "C6", "cost_of_capital": "C7",
            "after_tax_operating_margin": "C8", "tax_rate": "C9", "stable_growth": "D5",
            "stable_reinvestment_rate": "D6", "stable_cost_of_capital": "D7"}
    exp = {"ev_ebit_after_tax_forward": "D32", "ev_ebit_after_tax_trailing": "D31", "ev_ebit_forward": "C32",
           "ev_ebit_trailing": "C31", "ev_invested_capital": "E31", "ev_sales_forward": "F32",
           "ev_sales_trailing": "F31", "roic": "F19"}
    return {"mode": "firm", **{k: v(a) for k, a in keys.items()}}, {k: v(a) for k, a in exp.items()}


def x_eqcalc(get) -> tuple[dict, dict]:
    v = lambda a: get("Equity Multiple", a)
    inp = {"mode": "equity", "life_cycle": v("C8"), "riskfree_rate": v("C6"), "roe": v("C7"), "years": v("C10"),
           "eps": v("C4")}
    exp = {"value_per_share": v("C19"), "pe_trailing": v("C20"), "pbv_roe_trailing": v("C21"), "growth": v("C11"),
           "payout": v("C12"), "cost_of_equity": v("C13"), "stable_payout": v("C16")}
    return inp, exp


def x_alibaba(get) -> tuple[dict, dict]:
    io = lambda a: get("Input & Output", a)
    table = {"Median": "Sector Medians", "Simple Average": "Sector Averages"}.get(io("B10"), "Sector Aggregates")
    firm_ = dict(zip(["revenues", "ebitda", "net_income", "book_equity", "debt", "cash", "other_holdings"],
                     [io(f"B{r}") for r in range(2, 9)]))
    peers, exp = {}, {}
    for col in "BCDEF":
        group = io(f"{col}12").strip()
        for r in range(13, 18):
            m = io(f"A{r}").strip()
            peers.setdefault(group, {})[m] = get(table, f"{col}{r - 11}")
            exp[f"comps|{group}|{m}"] = io(f"{col}{r}")
    return {"mode": "comps", "firm": firm_, "peers": peers}, exp


EXTRACT = {"eqmult": ("Sheet1", x_eqmult), "firmmult": ("Sheet1", x_firmmult),
           "eqmultcalculator": ("Equity Multiple", x_eqcalc), "alibabarelvalue": (None, x_alibaba)}
GR = "the sheet's IF(g=r) branch returns n x earnings (no payout, no discounting); we use the limit of its main formula"
BRANCHES = {  # pycel scenarios on Damodaran's templates (plain-value cells only)
    "eqmult": [("5-year growth", {"C4": 5}), ("20% growth, 20% payout", {"C5": 0.2, "C6": 0.2}),
               ("stable 4% growth, 50% payout", {"D5": 0.04, "D6": 0.5}), ("rf 4.5% / 4%", {"C7": 0.045, "D7": 0.04}),
               ("ERP 6% (stable follows)", {"C8": 0.06}), ("beta 1.4 -> 1.1", {"C9": 1.4, "D9": 1.1}),
               ("net margin 10%", {"C10": 0.1}), ("growth below cost of equity", {"C5": 0.06, "C6": 0.6}),
               ("g = cost of equity (sheet IF branch)", {"C5": 0.085}, GR)],
    "firmmult": [("5-year growth", {"C4": 5}), ("15% growth, 80% reinvestment", {"C5": 0.15, "C6": 0.8}),
                 ("stable 2%, RR 10%", {"D5": 0.02, "D6": 0.1}), ("WACC 10% / 9%", {"C7": 0.1, "D7": 0.09}),
                 ("margin 8%", {"C8": 0.08}), ("tax 25%", {"C9": 0.25}),
                 ("g = cost of capital (sheet IF branch)", {"C5": 0.0803}, GR)],
    "eqmultcalculator": [("Growth", {"C8": "Growth"}), ("Mature", {"C8": "Mature"}),
                         ("Growth, rf 4%, ROE 20%", {"C8": "Growth", "C6": 0.04, "C7": 0.2}),
                         ("Mature, 5 years, EPS 3", {"C8": "Mature", "C10": 5, "C4": 3})]}


def build_cases(files: list[str]) -> int:
    paths = [Path(f) for f in files]
    stems = {p.stem.lower() for p in paths if p.suffix.lower() == ".xls"}
    cases = []
    for p in paths:
        stem = p.stem.lower()
        if stem not in EXTRACT:
            print("skip (unknown layout):", p)
            continue
        url = (TOOLS if stem in ("eqmult", "firmmult") else BLOG) + p.name.replace(".xlsx", ".xls" if stem in stems else ".xlsx")
        sheet, extract = EXTRACT[stem]
        if not (p.suffix.lower() == ".xlsx" and stem in stems):
            inp, exp = extract(_reader(p))
            cases.append({"name": p.stem, "source": p.name, "url": url, "inputs": inp, "expected": exp})
        if p.suffix.lower() == ".xlsx" and stem in BRANCHES:
            import warnings
            warnings.filterwarnings("ignore")
            from pycel import ExcelCompiler
            for name, cells, *why in BRANCHES[stem]:
                xl = ExcelCompiler(filename=str(p))
                get = lambda sh, a: xl.evaluate(f"{sh}!{a}")
                extract(get)  # places the cells in pycel's cell map
                for a, val in cells.items():
                    xl.set_value(f"{sheet}!{a}", val)
                xl.recalculate()
                inp, exp = extract(get)
                un = {k: (exp.pop(k), why[0]) for k in list(exp) if why and k.startswith(("pe", "pbv", "ps", "ev"))}
                cases.append({"name": f"{p.stem} branch: {name}", "source": p.name, "url": url, "custom": ["branch"],
                              "inputs": inp, "expected": exp, **({"unmatched": un} if un else {})})
                if un:  # the same branch with the sheet's IF reproduced: every cell must match
                    cases.append({"name": f"{p.stem} branch: {name}, replicate_sheet_bugs", "source": p.name,
                                  "url": url, "custom": ["branch"], "inputs": {**inp, "replicate_sheet_bugs": True},
                                  "expected": {**exp, **{k: v for k, (v, _) in un.items()}}})
    CASES.write_text(json.dumps(cases, indent=1), encoding="utf-8")
    print(f"wrote {len(cases)} cases -> {CASES}")
    return selftest()


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("inputs", nargs="?")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--solve", metavar="FIELD:MULTIPLE")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--build-cases", nargs="+", metavar="FILE")
    a = ap.parse_args()
    if a.selftest:
        return selftest()
    if a.build_cases:
        return build_cases(a.build_cases)
    if not a.inputs:
        ap.print_help()
        return 1
    p = json.loads(Path(a.inputs).read_text(encoding="utf-8"))
    p = {k: v for k, v in p.items() if not k.startswith("_")}
    out = run(p)
    print(json.dumps(out, indent=1) if a.json else report(p, out))
    if a.solve:
        field, mult = a.solve.split(":")
        x = solve(p, field, mult)
        print(f"\nIMPLIED {field} for {mult} = {p['actual'][mult]}: " + ("none in [-30%, 60%]" if x is None else f"{x:.4%}"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
