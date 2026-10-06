"""Acquisitions Damodaran's way: value of control, value of synergy, and the most an acquirer can pay.

Damodaran's three steps (Acquisition Alchemy, 2012-12): (1) value the target as is (status quo); (2) value it run optimally,
value of control = optimal - status quo; (3) value the combined firm with synergy built in, value of synergy =
combined - (target with control + acquirer). Most an acquirer can pay = target (with control) + synergy.

Transcribed from Damodaran's calculators (spreadsh.htm, URL https://pages.stern.nyu.edu/~adamodar/pc/<name>):
  controlvalue.xls      'Control value'!B2:C38  status quo vs optimal; optimal beta = status quo unlevered beta
                        relevered at the optimal D/C and tax rate; optimal EBIT defaults to status quo x 1.1 (C13)
  synergyvaluation.xls  'Sheet1'!B2:E38         acquirer, target, combined; combined beta = unlevered betas weighted
                        by firm value, relevered at combined D/C (terminal WACC: weighted by terminal values, E7:E24)
and the extended worksheets Damodaran used on real deals (URL https://pages.stern.nyu.edu/~adamodar/pc/blog/<name>):
  KFTcontrolvalue.xls (Kraft, 2012-06), SABControl.xls (SABMiller, 2015-10: effective vs marginal tax, cash, debt,
  cross holdings, stable ROC = ROC), TWCComcast.xls / InbevSABrestrusynergy.xls (the 'synergy worksheet, updated
  February 2015': ERP per firm, kd = riskfree + default spread, D/E, equity value, max price/share, exchange ratio).

Firm model (all sheets): ke = rf + beta ERP; kd = pre-tax kd (1 - marginal tax); WACC = ke (1-D/C) + kd D/C;
g = after-tax ROC x reinvestment rate; NOPAT = EBIT (1 - effective tax);
PV(high growth) = NOPAT (1-RR)(1+g)(1-(1+g)^n/(1+WACC)^n)/(WACC-g);
TV = NOPAT (1+g)^n (1+gs)(1-gs/ROC_s)/(WACC_s-gs), gs = riskfree, ROC_s = WACC_s (per Damodaran: zero excess returns
after the growth period); value of operating assets = PV + TV/(1+WACC)^n.

Usage:
  python mna.py inputs.json [--json]     mode "control" or "synergy" (see below)
  python mna.py --selftest               replay data/mna_cases.json (Damodaran's files + pycel template branches)
  python mna.py --build-cases FILES...   rebuild the cases from the source .xls/.xlsx (xlrd/openpyxl/pycel, maintenance venv);
                                         give both X.xls (the cached values) and X.xlsx (Excel Save-As copy, drives
                                         pycel branch tests) for controlvalue / synergyvaluation

Firm inputs (decimals; one currency unit):
  beta | unlevered_beta          levered beta, or unlevered (relevered at D/E with marginal tax). Give both when the
                                 sheet does (levered for the firm's own ke, unlevered for the combined-beta weighting)
  erp                            default: top-level erp
  pre_tax_cost_of_debt | default_spread (kd = riskfree + spread) | after_tax_cost_of_debt (override)
  tax_rate | marginal_tax_rate + effective_tax_rate   (marginal: debt shield, relevering; effective: NOPAT)
  debt_to_capital | debt_to_equity
  ebit, roc (after-tax) | pretax_roc, reinvestment_rate, years; expected_growth (override of roc x RR)
  stable_roc (number, or "roc" = keep the firm's ROC in perpetuity as in Damodaran's 2015 sheets; default stable WACC),
  stable_growth (default riskfree), stable_cost_of_capital (default WACC)
  optional: revenues, cash, debt, cross_holdings, minority_interests, shares, price  -> equity, value per share
Mode "control": riskfree_rate, erp, status_quo {firm}, optimal {only what changes; rest inherited},
  probability_of_change (optional -> expected value = status quo + p x control, Damodaran's 2014-09 Alibaba post)
Mode "synergy": riskfree_rate, erp, acquirer {firm}, target {firm}, target_optimal {optional: control},
  combined {what the merger changes: ebit | operating_margin (x combined revenues), roc, reinvestment_rate, years,
  tax rates, debt_to_capital, pre_tax_cost_of_debt | default_spread, terminal_weights "terminal_value" (tool
  template, default) | "current_value" (2015 worksheet)}. Omitted combined inputs follow the 2015 worksheet's
  status-quo column: revenues and EBIT summed, taxes EBIT-weighted, ERP revenue-weighted, D/E = total debt /
  total equity value, ROC = combined NOPAT / combined invested capital, RR = total reinvestment / total NOPAT,
  years = target's. The combined pre-tax cost of debt must be given when the two firms' differ (a new rating).
  replicate_sheet_bugs (synergy mode): true drops the (1 + riskfree) factor from the combined terminal value, as
  TWCComcast.xls E44 does (D2 blank); default = corrected formula (as in the template and InBev sheet).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
CASES = HERE.parent / "data" / "mna_cases.json"
TOOLS, BLOG = "https://pages.stern.nyu.edu/~adamodar/pc/", "https://pages.stern.nyu.edu/~adamodar/pc/blog/"
KD_KEYS = ("pre_tax_cost_of_debt", "default_spread", "after_tax_cost_of_debt")


def norm(f: dict) -> dict:
    """tax_rate -> marginal + effective; debt_to_capital -> debt_to_equity."""
    f = dict(f)
    t = f.pop("tax_rate", None)
    if t is not None:
        f.setdefault("marginal_tax_rate", t)
        f.setdefault("effective_tax_rate", t)
    if "marginal_tax_rate" in f:
        f.setdefault("effective_tax_rate", f["marginal_tax_rate"])
    dc = f.pop("debt_to_capital", None)
    if dc is not None:
        f.setdefault("debt_to_equity", dc / (1 - dc))
    return f


def firm(f: dict, rf: float) -> dict:
    """One column of the control / synergy sheets: two-stage FCFF value of operating assets."""
    f = norm(f)
    tm, te, de = f["marginal_tax_rate"], f["effective_tax_rate"], f["debt_to_equity"]
    dc = de / (1 + de)
    beta = f["beta"] if "beta" in f else f["unlevered_beta"] * (1 + (1 - tm) * de)
    bu = f.get("unlevered_beta", beta / (1 + (1 - tm) * de))
    ke = rf + beta * f["erp"]
    kd_pre = f["pre_tax_cost_of_debt"] if "pre_tax_cost_of_debt" in f else rf + f.get("default_spread", float("nan"))
    kd = f.get("after_tax_cost_of_debt", kd_pre * (1 - tm))
    wacc = ke * (1 - dc) + kd * dc
    roc = f["roc"] if "roc" in f else f["pretax_roc"] * (1 - te)
    rr, n = f["reinvestment_rate"], f["years"]
    g = f.get("expected_growth", roc * rr)
    nopat = f["ebit"] * (1 - te)
    annuity = (1 - ((1 + g) / (1 + wacc)) ** n) / (wacc - g) if wacc != g else n / (1 + wacc)  # limit at g = WACC
    pv = nopat * (1 - rr) * (1 + g) * annuity
    gs = f.get("stable_growth", rf)
    rs = f.get("stable_cost_of_capital", wacc)
    roc_s = roc if f.get("stable_roc") == "roc" else f.get("stable_roc", rs)
    tv_growth = 1.0 if f.get("replicate_sheet_bugs") else 1 + gs  # TWCComcast E44: (1+D2), D2 blank
    tv = nopat * (1 + g) ** n * tv_growth * (1 - gs / roc_s) / (rs - gs)
    out = {"beta": beta, "unlevered_beta": bu, "cost_of_equity": ke, "pre_tax_cost_of_debt": kd_pre,
           "after_tax_cost_of_debt": kd, "debt_to_capital": dc, "cost_of_capital": wacc, "roc": roc,
           "reinvestment_rate": rr, "growth": g, "years": n, "nopat": nopat, "pv_high_growth": pv,
           "terminal_value": tv, "stable_cost_of_capital": rs, "stable_roc": roc_s,
           "value": pv + tv / (1 + wacc) ** n}
    if any(k in f for k in ("cash", "debt", "cross_holdings", "minority_interests")):
        out["equity"] = (out["value"] + f.get("cash", 0.0) + f.get("cross_holdings", 0.0) - f.get("debt", 0.0)
                         - f.get("minority_interests", 0.0))
        if f.get("shares"):
            out["value_per_share"] = out["equity"] / f["shares"]
    return out


def optimal(base: dict, over: dict, base_out: dict) -> tuple[dict, list[str]]:
    """Optimal-firm inputs: what changes, the rest inherited from the status quo. Beta is relevered from the status quo
    unlevered beta; EBIT defaults to status quo x 1.1 (controlvalue.xls C13 = B13*1.1)."""
    b = {k: v for k, v in norm(base).items() if k not in ("beta", "expected_growth")}
    b["unlevered_beta"] = base_out["unlevered_beta"]
    o = norm(over)
    for grp in (("roc", "pretax_roc"), KD_KEYS, ("beta", "unlevered_beta")):
        if any(k in o for k in grp):
            for k in grp:
                b.pop(k, None)
    notes = []
    if "ebit" not in o:
        o["ebit"] = b["ebit"] * 1.1
        notes.append("optimal EBIT not given: the template default = status quo EBIT x 1.1")
    return {**b, **o}, notes


def common(inp: dict, f: dict) -> dict:
    return {**({"erp": inp["erp"]} if "erp" in inp else {}), **f}


def control(inp: dict) -> dict:
    rf = inp["riskfree_rate"]
    sq_in = common(inp, inp["status_quo"])
    sq = firm(sq_in, rf)
    opt_in, notes = optimal(sq_in, inp.get("optimal", {}), sq)
    opt = firm(opt_in, rf)
    out = {"status_quo": sq, "optimal": opt, "value_of_control": opt["value"] - sq["value"], "notes": notes}
    if "equity" in sq:
        out["control_equity"] = opt["equity"] - sq["equity"]
    p = inp.get("probability_of_change")
    if p is not None:  # Expected value = optimal x p + status quo x (1 - p)   (2014-09)
        out["expected_value"] = sq["value"] + p * out["value_of_control"]
        if "value_per_share" in sq:
            out["expected_value_per_share"] = sq["value_per_share"] + p * (opt["value_per_share"] - sq["value_per_share"])
    return out


def synergy(inp: dict) -> dict:
    rf = inp["riskfree_rate"]
    a_in, t_in = norm(common(inp, inp["acquirer"])), norm(common(inp, inp["target"]))
    A, T = firm(a_in, rf), firm(t_in, rf)
    out = {"acquirer": A, "target": T, "notes": []}
    if "target_optimal" in inp:  # Damodaran's Step 3: synergy is measured against the target WITH the control premium
        t_in, out["notes"] = optimal(t_in, inp["target_optimal"], T)
        out["target_status_quo"], T = T, firm(t_in, rf)
        out["target"], out["value_of_control"] = T, T["value"] - out["target_status_quo"]["value"]

    c = norm(inp.get("combined", {}))
    if inp.get("replicate_sheet_bugs"):  # applies to the combined column only, where the TWC sheet's slip sits
        c["replicate_sheet_bugs"] = True
        out["notes"].append("replicate_sheet_bugs: combined terminal value drops the (1 + riskfree) growth factor, "
                            "as TWCComcast.xls E44 does (=(1+D2), D2 blank); the default is the corrected TV")
    weights = c.pop("terminal_weights", "terminal_value")
    fs = (a_in, t_in)
    rev = [f.get("revenues") for f in fs]
    if "revenues" not in c and None not in rev:
        c["revenues"] = sum(rev)
    ebit_sum = a_in["ebit"] + t_in["ebit"]
    if "ebit" not in c:
        c["ebit"] = c["operating_margin"] * c["revenues"] if "operating_margin" in c else ebit_sum
    for k in ("marginal_tax_rate", "effective_tax_rate"):  # EBIT-weighted ('Return on capital'!D3)
        c.setdefault(k, (a_in["ebit"] * a_in[k] + t_in["ebit"] * t_in[k]) / ebit_sum)
    if "erp" not in c:  # revenue-weighted ('Synergy value'!D3)
        c["erp"] = a_in["erp"] if a_in["erp"] == t_in["erp"] else \
            (a_in["erp"] * rev[0] + t_in["erp"] * rev[1]) / (rev[0] + rev[1])
    if "debt_to_equity" not in c:  # total debt / total equity value ('Synergy value'!D23)
        if "debt" not in a_in or "debt" not in t_in:
            raise ValueError("combined.debt_to_capital is needed (or give debt for both firms)")
        c["debt_to_equity"] = (a_in["debt"] + t_in["debt"]) / (A["equity"] + T["equity"])
    if not any(k in c for k in KD_KEYS):
        if abs(A["pre_tax_cost_of_debt"] - T["pre_tax_cost_of_debt"]) > 1e-12:
            raise ValueError("combined.pre_tax_cost_of_debt (or default_spread) is needed: the firms' costs differ")
        c["pre_tax_cost_of_debt"] = A["pre_tax_cost_of_debt"]
    nopat = [A["nopat"], T["nopat"]]
    if "roc" not in c and "pretax_roc" not in c:  # combined NOPAT / combined capital ('Return on capital'!D9)
        c["roc"] = ebit_sum * (1 - c["effective_tax_rate"]) / (nopat[0] / A["roc"] + nopat[1] / T["roc"])
    c.setdefault("reinvestment_rate", (nopat[0] * A["reinvestment_rate"] + nopat[1] * T["reinvestment_rate"]) / sum(nopat))
    c.setdefault("years", t_in["years"])
    if "beta" not in c and "unlevered_beta" not in c:
        wsum = lambda key: (A["unlevered_beta"] * A[key] + T["unlevered_beta"] * T[key]) / (A[key] + T[key])
        c["unlevered_beta"] = wsum("value")
        if weights == "terminal_value" and "stable_cost_of_capital" not in c:  # synergyvaluation.xls E7, E22:E24
            first = firm(c, rf)
            bt = wsum("terminal_value") * (1 + (1 - c["marginal_tax_rate"]) * c["debt_to_equity"])
            dc = first["debt_to_capital"]
            c["stable_cost_of_capital"] = (rf + bt * c["erp"]) * (1 - dc) + first["after_tax_cost_of_debt"] * dc
    C = firm(c, rf)
    syn = C["value"] - (A["value"] + T["value"])
    out.update(combined=C, synergy=syn, max_firm_value=T["value"] + syn)
    if "equity" in T:
        out["max_value_to_pay"] = T["equity"] + syn
        if t_in.get("shares"):
            out["max_price_per_share"] = out["max_value_to_pay"] / t_in["shares"]
            if t_in.get("price"):
                out["max_premium_over_price"] = out["max_price_per_share"] / t_in["price"] - 1
            if "value_per_share" in A:
                out["breakeven_exchange_ratio"] = out["max_price_per_share"] / A["value_per_share"]
    return out


def run(inp: dict) -> dict:
    return control(inp) if inp.get("mode", "control") == "control" else synergy(inp)


# ---------------------------------------------------------------- report
ROWS = [("beta", "beta", "{:.3f}"), ("cost of equity", "cost_of_equity", "{:.2%}"),
        ("after-tax kd", "after_tax_cost_of_debt", "{:.2%}"), ("debt/capital", "debt_to_capital", "{:.1%}"),
        ("cost of capital", "cost_of_capital", "{:.2%}"), ("ROC (after tax)", "roc", "{:.2%}"),
        ("reinvestment rate", "reinvestment_rate", "{:.1%}"), ("growth", "growth", "{:.2%}"),
        ("years", "years", "{:g}"), ("PV high growth", "pv_high_growth", "{:,.1f}"),
        ("terminal value", "terminal_value", "{:,.1f}"), ("stable ROC", "stable_roc", "{:.2%}"),
        ("VALUE (operating)", "value", "{:,.1f}"), ("equity", "equity", "{:,.1f}"),
        ("value/share", "value_per_share", "{:,.2f}")]


HINTS = """DAMODARAN'S HINTS (controlvalue.xls E8:E14): 1. a different debt mix may lower the cost of capital; 2. compare the
pre-tax operating margin with the industry (cost cutting); 3. ROC below the cost of capital -> set it at least to the
cost of capital, above it -> check industry and trend; 4. high ROC with low reinvestment -> room to reinvest more;
5. competitive advantages -> a longer growth period (only with excess returns). Work within the realm of the possible."""


def industry_check(inp: dict, sq: dict) -> str:
    """SABControl.xls F17:G23: the firm next to its industry, the benchmarks Damodaran used for the optimal inputs."""
    from cost_of_capital import industry, load
    ind = industry(load(), inp["industry"], inp.get("scope", "US"))
    f = norm(inp["status_quo"])
    num = lambda x: x if isinstance(x, (int, float)) else None
    rows = [("pre-tax operating margin", f["ebit"] / f["revenues"] if f.get("revenues") else None,
             num(ind.get("Pre-tax Operating Margin (Unadjusted)"))),
            ("effective tax rate", f["effective_tax_rate"], num(ind.get("Average effective tax rate"))),
            ("after-tax ROIC", sq["roc"], num(ind.get("After-tax ROC"))),
            ("reinvestment rate", sq["reinvestment_rate"], num(ind.get("Reinvestment Rate"))),
            ("debt to capital", sq["debt_to_capital"], num(ind.get("Market Debt/Capital"))),
            ("cost of capital", sq["cost_of_capital"], num(ind.get("Cost of capital")))]
    pct = lambda x: "   n/a" if x is None else f"{x:6.1%}"
    return "\n".join([f"\nSTATUS QUO vs INDUSTRY {inp['industry']} ({inp.get('scope', 'US')}, Damodaran's Jan 2026 data)"]
                     + [f"   {a:26s} firm {pct(b)}   industry {pct(c)}" for a, b, c in rows])


def table(cols: dict) -> str:
    lines = [f"{'':20s}" + "".join(f"{k:>18s}" for k in cols)]
    for label, key, fmt in ROWS:
        if any(key in c for c in cols.values()):
            lines.append(f"{label:20s}" + "".join(f"{fmt.format(c[key]) if key in c else '':>18s}" for c in cols.values()))
    return "\n".join(lines)


def report(inp: dict, out: dict) -> str:
    if "status_quo" in out:
        s = [table({"status quo": out["status_quo"], "optimal": out["optimal"]}), "",
             f"VALUE OF CONTROL      {out['value_of_control']:,.1f}  "
             f"({out['value_of_control'] / out['status_quo']['value']:.1%} of status quo value)"]
        if "expected_value" in out:
            s.append(f"EXPECTED VALUE        {out['expected_value']:,.1f}  (p(management change) = "
                     f"{inp['probability_of_change']:.0%})" + (f"  per share {out['expected_value_per_share']:,.2f}"
                                                               if "expected_value_per_share" in out else ""))
        if inp.get("industry"):
            s.append(industry_check(inp, out["status_quo"]))
        s.append(HINTS)
    else:
        cols = {"acquirer": out["acquirer"]}
        if "target_status_quo" in out:
            cols["target (as is)"] = out["target_status_quo"]
        cols.update({"target": out["target"], "combined": out["combined"]})
        s = [table(cols), "",
             f"VALUE OF SYNERGY      {out['synergy']:,.1f}  = combined - (acquirer + target"
             + (" with control)" if "target_status_quo" in out else ")")]
        if "value_of_control" in out:
            s.append(f"VALUE OF CONTROL      {out['value_of_control']:,.1f}")
        s.append(f"MAX FIRM VALUE        {out['max_firm_value']:,.1f}  (target operating assets + synergy)")
        for k, lab, fmt in (("max_value_to_pay", "MAX EQUITY PRICE", "{:,.1f}"), ("max_price_per_share", "MAX PRICE/SHARE", "{:,.2f}"),
                            ("max_premium_over_price", "MAX PREMIUM", "{:.1%}"),
                            ("breakeven_exchange_ratio", "BREAKEVEN EXCH. RATIO", "{:.4f}")):
            if k in out:
                s.append(f"{lab:22s}{fmt.format(out[k])}")
        s.append("RULE: never pay a premium above control + synergy, and negotiate for a share of it (2012-12)")
    return "\n".join(s + [f"NOTE: {n}" for n in out["notes"]])


# ---------------------------------------------------------------- self test
def pick(d: dict, path: str) -> float:
    for k in path.split("."):
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
            print(f"{'ok ' if ok else 'BAD'} {c['name']:42s} {key:32s} Damodaran {his:14,.6f}  ours {ours:14,.6f}")
        for key, (his, why) in c.get("unmatched", {}).items():
            print(f"n/a {c['name']:42s} {key:32s} Damodaran {his:14,.6f}  ours {pick(out, key):14,.6f}  ({why})")
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


def _tool_firm(v, col: str) -> dict:
    keys = {"beta": 7, "pre_tax_cost_of_debt": 8, "tax_rate": 9, "debt_to_capital": 10, "revenues": 12, "ebit": 13,
            "pretax_roc": 15, "reinvestment_rate": 16, "years": 18}
    return {k: v(f"{col}{r}") for k, r in keys.items()}


def x_control(get) -> tuple[dict, dict, list]:  # controlvalue.xls, KFTcontrolvalue.xls
    v = lambda a: get("Control value", a)
    sq, opt = _tool_firm(v, "B"), _tool_firm(v, "C")
    del opt["beta"]  # C7 is the sheet's relevering formula
    inp = {"mode": "control", "riskfree_rate": v("B2"), "erp": v("B3"), "status_quo": sq, "optimal": opt}
    custom = []
    for col, key in (("B", "status_quo"), ("C", "optimal")):
        if abs(v(col + "27") - v(col + "16")) > 1e-12:  # the template: row 27 = row 16
            inp[key]["expected_growth"] = v(col + "28")
            custom.append(f"{col}27 typed as {v(col + '27')} over ={col}16 ({v(col + '16')}): growth uses "
                          f"{col}27, cash flows use {col}16 -> expected_growth passed")
    exp = {"status_quo.value": v("B33"), "optimal.value": v("C33"), "value_of_control": v("B38"),
           "optimal.beta": v("C7"), "optimal.cost_of_capital": v("C24")}
    return inp, exp, custom


def x_synergy_tool(get) -> tuple[dict, dict, list]:  # synergyvaluation.xls
    v = lambda a: get("Sheet1", a)
    comb = _tool_firm(v, "D")
    del comb["beta"]  # D7 is the sheet's combined-beta formula
    inp = {"mode": "synergy", "riskfree_rate": v("B2"), "erp": v("B3"), "acquirer": _tool_firm(v, "B"),
           "target": _tool_firm(v, "C"), "combined": comb}
    exp = {"acquirer.value": v("B33"), "target.value": v("C33"), "combined.value": v("D33"), "synergy": v("B38"),
           "combined.beta": v("D7"), "combined.stable_cost_of_capital": v("E24")}
    return inp, exp, []


def x_sab(get) -> tuple[dict, dict, list]:  # SABControl.xls (Damodaran's extended control sheet, 2015)
    v = lambda a: get("Control value", a)
    sq = {"unlevered_beta": v("D58"), "pre_tax_cost_of_debt": v("B7"), "marginal_tax_rate": v("B8"),
          "effective_tax_rate": v("B9"), "debt_to_capital": v("B10"), "revenues": v("B12"), "ebit": v("B14"),
          "pretax_roc": v("B16"), "reinvestment_rate": v("B17"), "years": v("B19"), "stable_roc": "roc",
          "cash": v("B35"), "cross_holdings": v("B36"), "debt": v("B37"), "minority_interests": v("B38")}
    opt = {"beta": v("C6"), "pre_tax_cost_of_debt": v("C7"), "marginal_tax_rate": v("C8"),
           "effective_tax_rate": v("C9"), "debt_to_capital": v("C10"), "ebit": v("C14"), "pretax_roc": v("C16"),
           "reinvestment_rate": v("C17"), "years": v("C19")}
    inp = {"mode": "control", "riskfree_rate": v("B2"), "erp": v("B3"), "status_quo": sq, "optimal": opt}
    exp = {"status_quo.value": v("B34"), "optimal.value": v("C34"), "value_of_control": v("B44"),
           "status_quo.equity": v("B39"), "optimal.equity": v("C39"), "status_quo.beta": v("B6")}
    return inp, exp, ["optimal beta C6 unlevers/relevers with the effective tax rate (rows shifted in the extended "
                      "sheet): passed as a levered beta", "stable ROC = ROC (B20 = B27): excess returns continue"]


LAYOUT_2015 = {  # 'Synergy value' rows of Damodaran's Feb-2015 synergy worksheet (InBev file has an extra minority row)
    "inbevsabrestrusynergy": dict(spread=17, beta=20, eff=22, de=23, rev=25, roc=28, rr=29, n=31, kd_at=36,
                                  wacc=37, value=46, cash=47, debt=48, eq=49, syn=54, mx=60, pps=61, xr=65),
    "twccomcast": dict(spread=16, beta=19, eff=21, de=22, rev=24, roc=27, rr=28, n=30, kd_at=35, wacc=36,
                       value=45, cash=46, debt=47, eq=48, syn=53, mx=59, pps=60, xr=64)}


def x_synergy_2015(get, stem: str) -> tuple[dict, dict, list]:
    L = LAYOUT_2015[stem]
    v = lambda a: get("Synergy value", a)
    inbev = stem.startswith("inbev")

    def one(c: str) -> dict:
        return {"erp": v(f"{c}3"), "marginal_tax_rate": v(f"{c}4"), "beta": v(f"{c}{L['beta']}"),
                "unlevered_beta": get("Business & Betas", "D9" if c == "B" else "D15"),
                "default_spread": v(f"{c}{L['spread']}"), "effective_tax_rate": v(f"{c}{L['eff']}"),
                "debt_to_equity": v(f"{c}{L['de']}"), "revenues": v(f"{c}{L['rev']}"),
                "ebit": get("Earnings & Cash Flows", f"{c}4") if inbev else v(f"{c}{L['rev'] + 1}"),
                "roc": v(f"{c}{L['roc']}"), "reinvestment_rate": v(f"{c}{L['rr']}"), "years": v(f"{c}{L['n']}"),
                "stable_roc": "roc", "cash": v(f"{c}{L['cash']}"), "debt": v(f"{c}{L['debt']}"), "shares": v(f"{c}11")}

    comb = {"marginal_tax_rate": v("E4"), "effective_tax_rate": v(f"E{L['eff']}"),
            "default_spread": v(f"E{L['spread']}"), "operating_margin": get("SYNERGY INPUTS", "E10"),
            "roc": v(f"E{L['roc']}"), "reinvestment_rate": v(f"E{L['rr']}"), "stable_roc": "roc",
            "terminal_weights": "current_value"}
    custom = ["stable ROC = ROC in perpetuity (row 6), TV growth = riskfree", "combined D/E, ERP, revenues, years "
              "left to the engine's defaults (the sheet's status-quo column rules)"]
    if not inbev:  # combined after-tax kd uses the effective tax rate (E35 = E20*(1-E21))
        comb["after_tax_cost_of_debt"] = v(f"E{L['kd_at']}")
        custom.append("combined after-tax kd E35 uses the effective tax rate: passed as after_tax_cost_of_debt")
    else:
        custom.append("SABMiller inputs are the restructured (optimal) values from SABControl.xls; EBIT is the "
                      "lease-adjusted 'Earnings & Cash Flows' row")
    inp = {"mode": "synergy", "riskfree_rate": v("B2"), "acquirer": one("B"), "target": one("C"), "combined": comb}
    exp = {"acquirer.value": v(f"B{L['value']}"), "target.value": v(f"C{L['value']}"),
           "acquirer.equity": v(f"B{L['eq']}"), "target.equity": v(f"C{L['eq']}"),
           "combined.beta": v(f"E{L['beta']}"), "combined.cost_of_capital": v(f"E{L['wacc']}"),
           "combined.value": v(f"E{L['value']}"), "synergy": v(f"B{L['syn']}"),
           "max_value_to_pay": v(f"B{L['mx']}"), "max_price_per_share": v(f"B{L['pps']}"),
           "breakeven_exchange_ratio": v(f"B{L['xr']}")}
    unmatched = {}
    if not inbev:  # E44 = E25*(1-E21)*(1+E40)^E30*(1+D2)*... : D2 is blank -> terminal growth factor 1, not 1+rf
        why = "the sheet's E44 multiplies by (1+D2) and D2 is blank: TV drops the (1+riskfree) growth factor"
        for k in ("combined.value", "synergy", "max_value_to_pay", "max_price_per_share", "breakeven_exchange_ratio"):
            unmatched[k] = (exp.pop(k), why)
    return inp, exp, custom, unmatched


# pycel branch scenarios on Damodaran's templates: plain-value cells only
BRANCHES = {
    "controlvalue": ("Control value", [
        ("baseline", {}), ("more debt, higher kd", {"C10": 0.5, "C8": 0.05}), ("tax change", {"B9": 0.35, "C9": 0.25}),
        ("better ROC, more reinvestment", {"C15": 0.2, "C16": 0.5}), ("growth periods 3 / 10", {"B18": 3, "C18": 10}),
        ("status quo EBIT 5000 (optimal EBIT default 1.1x)", {"B13": 5000}), ("rf 4%, ERP 5%", {"B2": 0.04, "B3": 0.05}),
        ("ROC below WACC", {"B15": 0.06, "B16": 0.5})]),
    "synergyvaluation": ("Sheet1", [
        ("baseline", {}), ("cost synergy EBIT 90", {"D13": 90}), ("growth synergy", {"D15": 0.2, "D16": 0.8}),
        ("debt capacity", {"D10": 0.3, "D8": 0.055}), ("tax synergy", {"D9": 0.25}),
        ("dissimilar firms", {"B7": 1.5, "C7": 0.7, "B10": 0.3, "C10": 0.05, "B9": 0.35, "C9": 0.2, "C13": 40,
                              "C15": 0.25, "C16": 0.4, "C18": 8, "B18": 3, "D18": 6, "D13": 95}),
        ("rf 3%, ERP 5.5%", {"B2": 0.03, "B3": 0.055})])}


def build_cases(files: list[str]) -> int:
    paths = [Path(f) for f in files]
    stems = {p.stem.lower() for p in paths if p.suffix.lower() == ".xls"}
    cases = []
    for p in paths:
        stem = p.stem.lower()
        url = (TOOLS if stem in ("controlvalue", "synergyvaluation") else BLOG) + p.stem + ".xls"
        if not (p.suffix.lower() == ".xlsx" and stem in stems):  # an .xlsx Save-As copy only drives pycel
            get = _reader(p)
            unmatched = {}
            if stem in ("controlvalue", "kftcontrolvalue"):
                inp, exp, custom = x_control(get)
            elif stem == "synergyvaluation":
                inp, exp, custom = x_synergy_tool(get)
            elif stem == "sabcontrol":
                inp, exp, custom = x_sab(get)
            elif stem in LAYOUT_2015:
                inp, exp, custom, unmatched = x_synergy_2015(get, stem)
            else:
                print("skip (unknown layout):", p)
                continue
            cases.append({"name": p.stem, "source": p.name, "url": url, "custom": custom, "inputs": inp,
                          "expected": exp, **({"unmatched": unmatched} if unmatched else {})})
            if unmatched:  # the sheet's slip reproduced: every cell must match
                cases.append({"name": p.stem + " (replicate_sheet_bugs)", "source": p.name, "url": url,
                              "custom": custom, "inputs": {**inp, "replicate_sheet_bugs": True},
                              "expected": {**exp, **{k: v for k, (v, _) in unmatched.items()}}})
        if p.suffix.lower() == ".xlsx" and stem in BRANCHES:
            import warnings
            warnings.filterwarnings("ignore")
            from pycel import ExcelCompiler
            sheet, scen = BRANCHES[stem]
            for name, cells in scen:
                xl = ExcelCompiler(filename=str(p))
                get = lambda sh, a: xl.evaluate(f"{sh}!{a}")
                extract = x_control if stem == "controlvalue" else x_synergy_tool
                extract(get)  # places every input/output cell in pycel's cell map
                for a, val in cells.items():
                    xl.set_value(f"{sheet}!{a}", val)
                xl.recalculate()
                inp, exp, _ = extract(get)
                if stem == "controlvalue" and "B13" in cells:
                    del inp["optimal"]["ebit"]  # exercise the 1.1x default
                cases.append({"name": f"{p.stem} branch: {name}", "source": p.name, "url": url, "custom": ["branch"],
                              "inputs": inp, "expected": exp})
    CASES.write_text(json.dumps(cases, indent=1), encoding="utf-8")
    print(f"wrote {len(cases)} cases -> {CASES}")
    return selftest()


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("inputs", nargs="?")
    ap.add_argument("--json", action="store_true")
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
    inp = json.loads(Path(a.inputs).read_text(encoding="utf-8"))
    inp = {k: v for k, v in inp.items() if not k.startswith("_")}
    out = run(inp)
    print(json.dumps(out, indent=1) if a.json else report(inp, out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
