"""Branch test: recalculate Damodaran's own template (pycel) under input overrides Damodaran's published files never
use, compare with ginzu.value(), and save the results to data/branch_cases.json for `ginzu.py --selftest`.

Usage: python branch_test.py <path/to/fcffsimpleginzu.xlsx>      (needs pycel + openpyxl)
"""
import json
import sys
import types
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).parent))

# pycel lacks Excel's NORMSDIST, which the source option sheet uses
extra = types.ModuleType("pycel_extra")
exec("import math\ndef normsdist(x):\n    return 0.5 * (1 + math.erf(x / math.sqrt(2)))", extra.__dict__)
sys.modules["pycel_extra"] = extra

from pycel import ExcelCompiler  # noqa: E402

from calibrate import read_case  # noqa: E402
from ginzu import CASES, value  # noqa: E402

X = Path(sys.argv[1])
base = read_case(X)["inputs"]
NOL = 474.8 + 256.6  # the template's NOL cell is a formula; overriding it is not possible in pycel

# (name, {Input sheet cell: value}, {our input key: value}); only plain-value cells are overridden
SCENARIOS = [
    ("baseline", {}, {}),
    ("failure V 20% @50%", {"B51": "Yes", "B52": 0.2, "B53": "V", "B54": 0.5},
     {"failure_probability": 0.2, "failure_proceeds_basis": "V", "failure_proceeds_pct": 0.5}),
    ("failure B 30% @40%", {"B51": "Yes", "B52": 0.3, "B53": "B", "B54": 0.4},
     {"failure_probability": 0.3, "failure_proceeds_basis": "B", "failure_proceeds_pct": 0.4}),
    ("lag 0", {"B56": "Yes", "B57": 0}, {"reinvestment_lag": 0}),
    ("lag 2", {"B56": "Yes", "B57": 2}, {"reinvestment_lag": 2}),
    ("lag 3", {"B56": "Yes", "B57": 3}, {"reinvestment_lag": 3}),
    ("keep effective tax", {"B59": "Yes"}, {"keep_effective_tax": True}),
    ("NOL (sheet 731.4)", {"B61": "Yes"}, {"nol": NOL}),
    ("loss-making + NOL", {"B12": -1500, "B61": "Yes"}, "LOSS"),
    ("riskfree after 10", {"B64": "Yes", "B65": 0.03}, {"riskfree_after_10": 0.03}),
    ("terminal growth -2%", {"B67": "Yes", "B68": -0.02}, {"terminal_growth": -0.02}),
    ("trapped cash", {"B70": "Yes", "B71": 5000, "B72": 0.1}, {"trapped_cash": 5000, "trapped_cash_tax_rate": 0.1}),
    ("options @price", {"B37": "Yes", "B38": 50, "B39": 60, "B40": 5, "B41": 0.3},
     {"options": {"number": 50, "strike": 60, "maturity": 5, "volatility": 0.3, "price_basis": "price"}}),
    ("coc+roic override", {"B45": "Yes", "B48": "Yes", "B49": 0.12},
     {"terminal_cost_of_capital": base["initial_cost_of_capital"], "terminal_roic": 0.12}),
]

bad, cases = 0, []
for name, cells, ours_over in SCENARIOS:
    xl = ExcelCompiler(filename=str(X), cycles={"iterations": 200, "tolerance": 1e-12}, plugins=["pycel_extra"])
    xl.evaluate("Valuation output!B33")  # builds the cell map
    for c, v in cells.items():
        xl.set_value("Input sheet!" + c, v)
    xl.recalculate()  # set_value alone does not refresh every dependent
    his = xl.evaluate("Valuation output!B33")
    if ours_over == "LOSS":  # the template's margin inputs are formulas of the base margin (I26): mirror them
        m = xl.evaluate("Input sheet!I26")
        ours_over = {"ebit": -1500, "nol": NOL, "margin_next_year": m, "target_margin": m}
    ours = value({**base, **ours_over})["value_per_share"]
    ok = abs(ours - his) <= 5e-4 * max(1.0, abs(his))
    bad += not ok
    print(f"{'ok ' if ok else 'BAD'} {name:22s} Damodaran {his:12.6f}  ours {ours:12.6f}")
    cases.append({"name": f"template branch: {name}", "custom": ["branch"], "inputs": {**base, **ours_over},
                  "expected_value_per_share": his})
print(f"{len(SCENARIOS) - bad}/{len(SCENARIOS)} branch scenarios match Damodaran's template")
out = CASES.parent / "branch_cases.json"
out.write_text(json.dumps(cases, indent=1, default=str), encoding="utf-8")
print(f"wrote {len(cases)} cases -> {out}")
