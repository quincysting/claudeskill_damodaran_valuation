"""Damodaran's SpaceX segment model (SpaceX2026IPO.xlsx / SpaceX2026IPOUpdated.xlsx), transcribed from its
'Valuation output' rows 3-19, feeding the skill's calibrated Ginzu engine (ginzu.py) and cost-of-capital engine.

The model's segment mechanics (per segment: Launch/Space, Starlink/Connectivity, xAI/AI; plus "Other"):
  revenue  year 5 = base + (target_2036 - base)/3; years 1-4 move toward year 5 by 20%, 30%, 40%, 50% of the gap;
           years 6-9 move toward the 2036 target by 20%, 30%, 40%, 50%; year 10 = target.
           "Other" is zero through year 5, then (t-5)/5 of its 2036 target.
  margin   converges linearly from the segment's current margin to its target by the convergence year.
  reinvest (revenue change) / segment sales-to-capital (years 1-5 and 6-10 ratios).

Usage:
  python spacex_model.py inputs.json [--json]     value from a segment-inputs file (see *_inputs.json here)
  python spacex_model.py inputs.json --sensitivity one-at-a-time scenarios + revenue breakeven vs price
  python spacex_model.py --calibrate              rebuild the June 2026 workbook from raw inputs (needs finance/.venv)
"""
import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
FIN = HERE.parents[1]
sys.path.insert(0, str(FIN / "skill" / "aswath-finance" / "scripts"))
from cost_of_capital import compute as coc_compute, load as coc_load  # noqa: E402
from ginzu import diagnostics, report, value, warnings  # noqa: E402

STEPS = (0.2, 0.3, 0.4, 0.5)


def segment_revenue(base: float, target: float) -> list[float]:
    y5 = base + (target - base) / 3
    rev = [base]
    for f in STEPS:
        rev.append(rev[-1] + (y5 - rev[-1]) * f)
    rev.append(y5)
    for f in STEPS:
        rev.append(rev[-1] + (target - rev[-1]) * f)
    rev.append(target)
    return rev  # years 0..10


def other_revenue(target: float) -> list[float]:
    return [0.0] * 6 + [(t - 5) * target / 5 for t in range(6, 11)]


def margins(m0: float, target: float, conv: int) -> list[float]:
    return [m0] + [target if t > conv else target - (target - m0) / conv * (conv - t) for t in range(1, 11)]


def build(seg_inputs: dict) -> dict:
    """Segment inputs -> ginzu inputs (explicit revenue, margin and reinvestment paths)."""
    segs, conv = seg_inputs["segments"], seg_inputs["margin_convergence_year"]
    rev_tot, ebit_tot, reinv = [0.0] * 11, [0.0] * 11, [0.0] * 11
    for s in segs:
        rev = other_revenue(s["revenue_2036"]) if s.get("starts_after_year5") else segment_revenue(s["base_revenue"], s["revenue_2036"])
        mg = margins(s["base_margin"], s["target_margin"], conv)
        for t in range(11):
            rev_tot[t] += rev[t]
            ebit_tot[t] += rev[t] * mg[t] if t else 0.0
            if t:
                s2c = s["sales_to_capital_1_5"] if t <= 5 else s["sales_to_capital_6_10"]
                reinv[t] += (rev[t] - rev[t - 1]) / s2c
    inp = {k: v for k, v in seg_inputs.items() if k not in ("segments", "cost_of_capital_inputs") and not k.startswith("_")}
    inp.update(revenues=rev_tot[0], margin_convergence_year=conv,
               revenue_growth_path=[rev_tot[t] / rev_tot[t - 1] - 1 for t in range(1, 11)],
               margin_path=[ebit_tot[t] / rev_tot[t] for t in range(1, 11)],
               reinvestment_path=reinv[1:])
    if seg_inputs.get("cost_of_capital_inputs"):
        coc = coc_compute({"riskfree_rate": seg_inputs["riskfree_rate"], "marginal_tax_rate": seg_inputs["marginal_tax_rate"],
                           **seg_inputs["cost_of_capital_inputs"]}, coc_load())
        inp["initial_cost_of_capital"], inp["_coc"] = coc["cost_of_capital"], coc
    return inp


def calibrate() -> int:
    """Rebuild the two workbooks from their raw inputs; must reproduce the published value per share."""
    import warnings as w
    import openpyxl
    w.filterwarnings("ignore")
    bad = 0
    # Only the June 2026 workbook uses these mechanics; the April workbook (SpaceX2026IPO) moves revenues in equal
    # steps (1/10, 1/9, ...) toward 2036 and is reproduced by ginzu.py's calibration cases instead.
    for name in ("SpaceX2026IPOUpdated",):
        wb = openpyxl.load_workbook(FIN / "spreadsheets" / f"{name}.xlsx", data_only=True)
        ins, vo = wb["Input sheet"], wb["Valuation output"]
        lab = {str(ins.cell(r, 1).value).strip().lower(): r for r in range(1, 90) if ins.cell(r, 1).value}
        v = lambda p, col=2: ins.cell(next(r for k, r in lab.items() if k.startswith(p)), col).value
        names = ["launch", "starlink", "xai"]
        segs = [{"base_revenue": v(f"revenues ({n})"), "revenue_2036": v(f"revenues ({n}) in 2036"),
                 "base_margin": vo.cell(8 + i, 2).value, "target_margin": v(f"target operating margin ({n})"),
                 "sales_to_capital_1_5": vo.cell(49 + i, 3).value, "sales_to_capital_6_10": vo.cell(49 + i, 8).value}
                for i, n in enumerate(names)]
        segs.append({"starts_after_year5": True, "revenue_2036": v("revenues (other) in 2036"), "base_margin": vo["B11"].value,
                     "target_margin": v("target operating margin (other)"), "sales_to_capital_1_5": vo["C52"].value,
                     "sales_to_capital_6_10": vo["H52"].value})
        # base inputs and default-assumption overrides exactly as the calibrated case file records them
        case = next(c for c in json.loads((FIN / "skill/aswath-finance/data/calibration_cases.json").read_text())
                    if c["name"] == name)
        base = {k: x for k, x in case["inputs"].items() if not k.endswith("_path") and k != "revenues"}
        got = value(build({**base, "margin_convergence_year": v("year of convergence"), "segments": segs}))["value_per_share"]
        sheet_val = case["expected_value_per_share"]
        ok = abs(got / sheet_val - 1) < 5e-4
        bad += not ok
        print(f"{'ok ' if ok else 'BAD'} {name:22s} Damodaran {sheet_val:10.4f}  rebuilt from raw segment inputs {got:10.4f}")
    return 1 if bad else 0


def per_share(seg_inputs: dict) -> float:
    return value(build(seg_inputs))["value_per_share"]


def scale_targets(d: dict, k: float) -> dict:
    d = json.loads(json.dumps(d))
    for s in d["segments"]:
        s["revenue_2036"] *= k
    return d


def scenarios(base: dict) -> list[tuple[str, dict]]:
    """One-at-a-time changes to the base case (shared with build_workbook.py's live scenario table)."""
    def seg(d, i, **kw):
        d = json.loads(json.dumps(d))
        d["segments"][i].update(kw)
        return d

    def with_coc(d, **kw):
        d = json.loads(json.dumps(d))
        for k, v in kw.items():
            d["cost_of_capital_inputs"][k] = v
        return d

    june_beta = {"approach": "multi_global", "businesses": [
        {"industry": "Aerospace/Defense", "revenues": 4100}, {"industry": "Telecom. Services", "revenues": 11400},
        {"industry": "Software (System & Application)", "revenues": 100}]}
    return [
        ("Base case (sep_2026_inputs.json)", base),
        ("June-model beta weights (xAI at $100m, stale)", with_coc(base, beta=june_beta)),
        ("Riskfree back at June 4.56% (terminal coc 8.25%)", {**base, "riskfree_rate": 0.0456, "terminal_cost_of_capital": 0.0825}),
        ("Terminal cost of capital = June model's 8.25%", {**base, "terminal_cost_of_capital": 0.0825}),
        ("Terminal cost of capital = rf + mature ERP 9.37%", {**base, "terminal_cost_of_capital": 0.0937}),
        ("xAI 2036 revenue $80bn (April 2026 view)", seg(base, 2, revenue_2036=80000)),
        ("xAI 2036 revenue $240bn", seg(base, 2, revenue_2036=240000)),
        ("Starlink target margin 50%", seg(base, 1, target_margin=0.50)),
        ("Starlink target margin 70%", seg(base, 1, target_margin=0.70)),
        ("xAI sales-to-capital 0.5 in yrs 1-5 (H1-26 capex pace)", seg(base, 2, sales_to_capital_1_5=0.5)),
        ("All 2036 revenue targets x0.75", scale_targets(base, 0.75)),
        ("All 2036 revenue targets x1.25", scale_targets(base, 1.25)),
    ]


def breakeven_multiplier(base: dict) -> float:
    """Scale on every 2036 revenue target that makes value = price (bisection)."""
    lo, hi = 0.5, 5.0
    for _ in range(80):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if per_share(scale_targets(base, mid)) < base["price"] else (lo, mid)
    return lo


def sensitivity(base: dict) -> str:
    """One-at-a-time changes to the base case, all run through the same engines, plus a revenue breakeven."""
    rows = [f"{'Scenario':52s} {'value/share':>11s} {'vs price':>9s}"]
    for name, d in scenarios(base):
        v = per_share(d)
        rows.append(f"{name:52s} {v:11.2f} {v / base['price'] - 1:9.1%}")
    lo = breakeven_multiplier(base)
    total = sum(s["revenue_2036"] for s in base["segments"])
    rows.append(f"\nBREAKEVEN: 2036 revenue targets x{lo:.3f} (total ${total * lo / 1000:,.0f}bn vs the model's "
                f"${total / 1000:,.0f}bn, same margins and reinvestment efficiency) make value = ${base['price']:.2f}")
    return "\n".join(rows)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="?")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--calibrate", action="store_true")
    ap.add_argument("--sensitivity", action="store_true")
    a = ap.parse_args()
    if a.calibrate:
        return calibrate()
    seg_inputs = json.loads(Path(a.inputs).read_text(encoding="utf-8"))
    if a.sensitivity:
        print(sensitivity(seg_inputs))
        return 0
    inp = build(seg_inputs)
    out = value(inp)
    if a.json:
        print(json.dumps({**out, "rules": warnings(inp, out), "diagnostics": diagnostics(inp, out)}, indent=1, default=str))
    else:
        print(report(inp, out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
