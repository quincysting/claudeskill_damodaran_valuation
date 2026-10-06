"""Selftest for export_xlsx.py: rebuild every ginzu calibration/branch case and every cost-of-capital case as a
workbook and recalculate it with an independent Excel formula engine (pycel). Run it through export_xlsx.py (selftest flag)."""
import json
import math
import sys
import tempfile
import types
from pathlib import Path

from export_xlsx import Book
from ginzu import CASES


def selftest() -> int:
    try:
        from pycel import ExcelCompiler
    except ImportError:
        print("SKIP export_xlsx selftest: needs pycel (pip install pycel; finance/.venv has it)")
        return 0
    import logging
    import warnings
    warnings.filterwarnings("ignore")
    logging.getLogger("pycel").setLevel(logging.CRITICAL)
    plug = types.ModuleType("_xl_plugin")
    plug.normsdist = lambda x: 0.5 * (1 + math.erf(x / math.sqrt(2)))
    plug.exact = lambda a, b: str(a) == str(b)  # Excel EXACT: case-sensitive, like Python ==
    sys.modules["_xl_plugin"] = plug
    tmp = Path(tempfile.mkdtemp()) / "t.xlsx"

    def evaluate(raw, cell):
        b = Book(raw)
        cells = b.build(tmp, tables=False)
        xl = ExcelCompiler(filename=str(tmp), plugins=["_xl_plugin"])
        try:
            for a in b.eval_order():
                xl.evaluate(a)
            return b, xl.evaluate(cells[cell])
        except Exception as e:  # noqa: BLE001 - report and continue with the next case
            return b, f"error: {str(e).splitlines()[-1][:160]}"

    cases = json.loads(CASES.read_text(encoding="utf-8")) + json.loads((CASES.parent / "branch_cases.json").read_text(encoding="utf-8"))
    bad = 0
    for c in cases:
        b, got = evaluate(c["inputs"], "value")
        eng = b.out["value_per_share"]
        ok = isinstance(got, float) and abs(got - eng) <= 1e-7 * max(1, abs(eng))
        got = got if isinstance(got, float) else float("nan")
        bad += not ok
        print(f"{'ok ' if ok else 'BAD'} {c['name']:34s} engine {eng:12.4f}  workbook {got:12.4f}")
    template = next(c["inputs"] for c in cases if c["name"] == "fcffsimpleginzu")
    coc = json.loads((CASES.parent / "coc_cases.json").read_text(encoding="utf-8"))["cases"]
    for c in coc:
        ci = {k: v for k, v in c["inputs"].items() if k not in ("riskfree_rate", "marginal_tax_rate")}
        raw = {**template, "riskfree_rate": c["inputs"]["riskfree_rate"], "marginal_tax_rate": c["inputs"]["marginal_tax_rate"],
               "cost_of_capital": ci}
        _, got = evaluate(raw, "wacc")
        ok = isinstance(got, float) and abs(got - c["expected"]) <= 1e-6
        got = got if isinstance(got, float) else float("nan")
        bad += not ok
        print(f"{'ok ' if ok else 'BAD'} coc: {c['name']:29s} Damodaran {c['expected']:.6%}  workbook {got:.6%}")
    n = len(cases) + len(coc)
    print(f"{n - bad}/{n} workbooks recalculate (pycel) to the engine's value / Damodaran's cost of capital")
    return 1 if bad else 0
