"""Option pricing and real options, transcribed cell-by-cell from Damodaran's option calculators.

Sources (https://pages.stern.nyu.edu/~adamodar/pc/<file>; all single-sheet Black-Scholes blocks):
  tool              source file (sheet)                      cells transcribed
  warrant           warrant.xls (Warrant Valuation)          D8:D15 -> C21 adjusted S (circular), B26:B30, C32
                    + the 'Option value' sheet of Damodaran's company valuations (Tesla, SpaceX, Adani, Instacart...)
  employee_options  employeeoption.xls (Warrant Valuation)   B8:B15, F17:F22 switches -> B27:B41
  short_term        optst.xls (NewShTermOptions)             G19:G40, C29:D32 dividends -> H42:H44, C47:C51, D53, F54
  long_term         optlt.xls (LT Option)                    C3:C5, G3:G5 -> C8:C12, C14 call, G14 put
                    buffettoptions.xls (older version: annual rate -> G37 = 1-EXP(-r), dividends/price yield)
  binomial          bstobin.xls (Sheet1)                     B2:B5, B10 -> B7 up, B8 down, B11:G15 price tree
  delay             delay.xls                                F5:F11 -> G16 = 1/years, C19:C23, D25
  expand            expand.xls (ProjectDelay)                G6:G10, F12 -> C20:C24, D26
  abandon           abandon.xls (ProjectDelay)               F4:F11 -> G16 = 1/life, C19:C23, D25 (put)
  patent            project.xls                              F19:F33 -> G38 = 1/life or cost of delay, E47
  natural_resource  natres.xls (NewNatResOption)             F16:F29 -> C32 = reserves x (price - cost), G34, E43
  flexibility       flexval.xls (Reinvestment Needs,         B2:C11 -> D13, E14; F19:F31 -> C41:F45, E47
                    Value of Flexibility)
  equity_option     equity.xls (Equity as an Option)         F15:F42 -> G46 variance (Industry | Firm inputs |
                                                             Direct Input), G47, C50:C54, D56:D59
Every formula keeps the spreadsheets' quirks: optst converts the rate with 1-EXP(-r); employeeoption's after-tax value (B41) is not
vesting-adjusted (a NOTE prints the vesting-adjusted figure); the warrant block solves the dilution circularity
(adjusted S depends on the option value) by iteration, as Excel does. One sheet error is corrected by default: optst's
call value (D53/F54) discounts the PV of dividends a second time; replicate_sheet_bugs=true reproduces the sheet's figure.

Usage:
  python options.py inputs.json [--json]       {"tool": "<name>", ...inputs}; readable report or JSON
  python options.py --from-case delay          the source inputs for a calibration case, as a starting JSON
  python options.py --selftest                 replay data/options_cases.json (Damodaran's cached outputs + pycel branches)
  python options.py --build-cases FILES...     rebuild the cases: the source .xls/.xlsx files (cached values, xlrd/openpyxl);
        .xlsx copies of the templates (Excel "Save As", same stem) add pycel branch tests; equity.xls also
        rewrites data/options_data.json (Damodaran's industry standard deviations)

Inputs (decimals; rates annualized; "sd" = standard deviation of ln(value), or give "variance"):
  warrant / employee_options: price, strike, maturity, sd, riskfree, warrants, shares, dividend_yield (0)
      employee_options only (spreadsheet switches): maturity_override (sheet default "No"), tax_rate (sheet default "Yes", 0.26; 0 = no),
      vesting_probability (sheet default "No")
  short_term: price, sd, strike, days, riskfree (annual; the sheet uses 1-EXP(-r)), dividends [[dps, days], ...] (max 4),
      replicate_sheet_bugs (false; true = the sheet's D53/F54 double discounting of the dividends)
  long_term: price, strike, maturity, sd, riskfree (continuous) | riskfree_annual (older template),
      dividend_yield (0) | annual_dividends (yield = dividends / price)
  binomial: price, sd, riskfree, dividend_yield (0), periods_per_year (1), periods (5, as in the source sheet)
  delay: pv_cash_flows, sd, investment, years, riskfree              (cost of delay = 1/years)
  expand: pv_expansion, sd, investment, years, riskfree, cost_of_delay (0)
  abandon: pv_cash_flows, sd, project_life, abandonment_value, years, riskfree   (cost of delay = 1/life)
  patent: pv_cash_flows, sd, development_cost, years, riskfree, cost_of_delay (spreadsheet default 1/years)
  natural_resource: reserves, price, unit_cost, sd, annual_cash_flow, development_cost, years, riskfree
  flexibility: reinvestment_pct, sd, financeable_pct, max_pct, riskfree, cost_of_capital, return_on_capital
      and/or history {reinvestment: [...], firm_value: [...]} (the spreadsheet's Reinvestment Needs sheet)
  equity_option: firm_value, face_value_debt, duration, riskfree, dividends (0) | dividend_yield (override),
      one of: sd | industry (Damodaran's table, data/options_data.json) |
              firm_inputs {sd_stock, sd_bond, correlation, debt_ratio}
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
DATA = HERE.parent / "data"
CASES = DATA / "options_cases.json"
TABLE = DATA / "options_data.json"
URL = "https://pages.stern.nyu.edu/~adamodar/pc/"


def nd(x: float) -> float:  # Excel NORMSDIST
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))


def bs(S: float, K: float, t: float, var: float, r: float, y: float = 0.0) -> dict:
    """The spreadsheets' Black-Scholes block: d1 = (LN(S/K)+(r-y+var/2)t)/(var^.5 t^.5); call; put by parity (optlt G14)."""
    d1 = (math.log(S / K) + (r - y + var / 2) * t) / (var ** 0.5 * t ** 0.5)
    d2 = d1 - var ** 0.5 * t ** 0.5
    call = math.exp(-y * t) * S * nd(d1) - K * math.exp(-r * t) * nd(d2)
    return {"S": S, "K": K, "t": t, "variance": var, "r": r, "y": y, "d1": d1, "N(d1)": nd(d1),
            "d2": d2, "N(d2)": nd(d2), "call": call, "put": call - S * math.exp(-y * t) + K * math.exp(-r * t)}


def var_of(inp: dict) -> float:
    return inp["variance"] if "variance" in inp else inp["sd"] ** 2


def warrant(inp: dict, employee: bool = False) -> dict:
    """warrant.xls / employeeoption.xls: adjusted S = (S*shares + value*warrants)/(shares + warrants), circular."""
    t = inp.get("maturity_override") or inp["maturity"] if employee else inp["maturity"]
    n, w, var = inp["shares"], inp["warrants"], var_of(inp)
    val, out = 0.0, {}
    for _ in range(1000):  # Excel iterative calculation of the circular reference
        s_adj = (inp["price"] * n + val * w) / (n + w)
        out = bs(s_adj, inp["strike"], t, var, inp["riskfree"], inp.get("dividend_yield", 0.0))
        if abs(out["call"] - val) < 1e-13:
            break
        val = out["call"]
    out = {**out, "adjusted_price": out["S"], "value_per_option": out["call"], "value_all": out["call"] * w}
    if employee:
        out["value_vesting_adjusted"] = out["value_all"] * inp.get("vesting_probability", 1.0)
        out["value_after_tax"] = out["value_all"] * (1 - inp.get("tax_rate", 0.26))  # the sheet's B41 uses B39, not B40
        if inp.get("vesting_probability", 1.0) != 1.0 and inp.get("tax_rate", 0.26):
            out["note"] = (f"the sheet's B41 after-tax value is taken from the pre-vesting total (B39); vesting-adjusted "
                           f"after tax = {out['value_vesting_adjusted'] * (1 - inp.get('tax_rate', 0.26)):,.4f}")
    return out


def short_term(inp: dict) -> dict:
    """optst.xls: dividends subtracted as PV (at the annual rate); rate converted with 1-EXP(-r).
    The sheet's D53/F54 discount the PV of dividends (H44) a second time, so the call uses a different S than its d1 and
    Damodaran's stated method (S - PV of dividends). Default: S - PV(div) throughout; replicate_sheet_bugs=true: the sheet's D53/F54."""
    divs = inp.get("dividends", [])
    if len(divs) > 4:
        raise ValueError("optst.xls takes at most 4 dividends")
    r, S, t = inp["riskfree"], inp["price"], inp["days"] / 365
    rc = 1 - math.exp(-r)  # H42
    pv_div = sum(d / (1 + r) ** (days / 365) for d, days in divs)  # H44
    out = {**bs(S - pv_div, inp["strike"], t, var_of(inp), rc), "rate_used": rc, "pv_dividends": pv_div}
    s_his = S - pv_div / (1 + rc) ** t  # D53 discounts H44 again
    call = s_his * out["N(d1)"] - inp["strike"] * math.exp(-rc * t) * out["N(d2)"]
    his = {"call": call, "put": call - s_his + inp["strike"] * math.exp(-rc * t)}
    if inp.get("replicate_sheet_bugs"):
        return {**out, **his}
    if pv_div:
        out["note"] = (f"SHEET ERROR: the optst D53 discounts the PV of dividends a second time (S {s_his:.4f} in the "
                       f"call vs S - PV(div) {S - pv_div:.4f} in d1): the sheet's call {his['call']:.4f}, put {his['put']:.4f}. "
                       "Set replicate_sheet_bugs=true to reproduce the source sheet")
    return out


def long_term(inp: dict) -> dict:
    """optlt.xls (continuous rate, dividend yield); older version (buffettoptions.xls) converts an annual rate."""
    r = 1 - math.exp(-inp["riskfree_annual"]) if "riskfree_annual" in inp else inp["riskfree"]
    y = inp["annual_dividends"] / inp["price"] if "annual_dividends" in inp else inp.get("dividend_yield", 0.0)
    return bs(inp["price"], inp["strike"], inp["maturity"], var_of(inp), r, y)


def binomial(inp: dict) -> dict:
    """bstobin.xls: Black-Scholes inputs -> binomial up/down moves and the price tree."""
    sd, n = inp["sd"], inp.get("periods_per_year", 1)
    drift = (inp["riskfree"] - inp.get("dividend_yield", 0.0) - sd ** 2 / 2) * (1 / n)
    u, d = math.exp(sd * (1 / n) ** 0.5 + drift), math.exp(-sd * (1 / n) ** 0.5 + drift)
    tree = [[inp["price"]]]
    for _ in range(inp.get("periods", 5)):
        tree.append([x * u for x in tree[-1]] + [tree[-1][-1] * d])
    return {"up": u, "down": d, "tree": tree}


def delay(inp: dict) -> dict:
    out = bs(inp["pv_cash_flows"], inp["investment"], inp["years"], var_of(inp), inp["riskfree"], 1 / inp["years"])
    return {**out, "value": out["call"]}


def expand(inp: dict) -> dict:
    out = bs(inp["pv_expansion"], inp["investment"], inp["years"], var_of(inp), inp["riskfree"],
             inp.get("cost_of_delay", 0.0))
    return {**out, "value": out["call"]}


def abandon(inp: dict) -> dict:
    if inp["years"] > inp["project_life"]:
        raise ValueError("abandon.xls: years with the right to abandon must be <= project life")
    out = bs(inp["pv_cash_flows"], inp["abandonment_value"], inp["years"], var_of(inp), inp["riskfree"],
             1 / inp["project_life"])
    return {**out, "value": out["put"]}


def patent(inp: dict) -> dict:
    y = inp.get("cost_of_delay", 1 / inp["years"])  # G38 = IF(F30="No", 1/life, F31)
    out = bs(inp["pv_cash_flows"], inp["development_cost"], inp["years"], var_of(inp), inp["riskfree"], y)
    return {**out, "value": out["call"]}


def natural_resource(inp: dict) -> dict:
    S = inp["reserves"] * (inp["price"] - inp["unit_cost"])  # C32
    out = bs(S, inp["development_cost"], inp["years"], var_of(inp), inp["riskfree"], inp["annual_cash_flow"] / S)
    return {**out, "value": out["call"]}


def reinvestment_stats(history: dict) -> tuple[float, float]:
    """Reinvestment Needs sheet: AVERAGE(reinvestment/value), STDEV(LN(ratio)) with 0 when reinvestment <= 0."""
    ratios = [b / c for b, c in zip(history["reinvestment"], history["firm_value"])]
    logs = [math.log(x) if b > 0 else 0.0 for x, b in zip(ratios, history["reinvestment"])]
    mean = sum(logs) / len(logs)
    return sum(ratios) / len(ratios), (sum((x - mean) ** 2 for x in logs) / (len(logs) - 1)) ** 0.5


def flexibility(inp: dict) -> dict:
    """flexval.xls: [call(K = financeable) - call(K = maximum)] x (ROC - WACC)/WACC, 1-year, no yield."""
    out = {}
    if "history" in inp:
        out["avg_reinvestment_pct"], out["sd_ln_reinvestment"] = reinvestment_stats(inp["history"])
    S = inp.get("reinvestment_pct", out.get("avg_reinvestment_pct"))
    var = inp["sd"] ** 2 if "sd" in inp else out["sd_ln_reinvestment"] ** 2
    lo = bs(S, inp["financeable_pct"], 1, var, inp["riskfree"])
    hi = bs(S, inp["max_pct"], 1, var, inp["riskfree"])
    excess = inp["return_on_capital"] - inp["cost_of_capital"]
    return {**out, **{k: lo[k] for k in ("d1", "N(d1)", "d2", "N(d2)")}, "call_financeable": lo["call"],
            **{k + "_max": hi[k] for k in ("d1", "N(d1)", "d2", "N(d2)")}, "call_max": hi["call"],
            "value": (lo["call"] - hi["call"]) * excess / inp["cost_of_capital"]}


def industry_sd(name: str) -> float:
    rows = json.loads(TABLE.read_text(encoding="utf-8"))["industries"]
    hit = {k.lower(): v for k, v in rows.items()}.get(name.lower())
    if hit is None:
        raise KeyError(f"industry '{name}' not in Damodaran's table ({TABLE.name}); e.g. {sorted(rows)[:6]}")
    return hit["sd_firm_value"]


def equity_option(inp: dict) -> dict:
    """equity.xls: equity = call on firm value, strike = face value of debt (plus coupons), t = debt duration."""
    V = inp["firm_value"]
    if "sd" in inp or "variance" in inp:
        var = var_of(inp)
    elif "firm_inputs" in inp:  # E25
        f = inp["firm_inputs"]
        w, ss, sb = f["debt_ratio"], f["sd_stock"], f["sd_bond"]
        var = (1 - w) ** 2 * ss ** 2 + w ** 2 * sb ** 2 + 2 * w * (1 - w) * ss * sb * f["correlation"]
    else:
        var = industry_sd(inp["industry"]) ** 2
    y = inp["dividend_yield"] if "dividend_yield" in inp else inp.get("dividends", 0.0) / V
    out = bs(V, inp["face_value_debt"], inp["duration"], var, inp["riskfree"], y)
    debt = V - out["call"]
    return {**out, "value": out["call"], "value_of_debt": debt,
            "interest_rate_on_debt": (inp["face_value_debt"] / debt) ** (1 / inp["duration"]) - 1,
            "default_probability": 1 - out["N(d2)"]}


TOOLS = {  # name: (function, source file, headline output)
    "warrant": (warrant, "warrant.xls", "value_per_option"),
    "employee_options": (lambda i: warrant(i, True), "employeeoption.xls", "value_after_tax"),
    "short_term": (short_term, "optst.xls", "call"),
    "long_term": (long_term, "optlt.xls", "call"),
    "binomial": (binomial, "bstobin.xls", "up"),
    "delay": (delay, "delay.xls", "value"),
    "expand": (expand, "expand.xls", "value"),
    "abandon": (abandon, "abandon.xls", "value"),
    "patent": (patent, "project.xls", "value"),
    "natural_resource": (natural_resource, "natres.xls", "value"),
    "flexibility": (flexibility, "flexval.xls", "value"),
    "equity_option": (equity_option, "equity.xls", "value"),
}


def run(inp: dict) -> dict:
    out = TOOLS[inp["tool"]][0]({k: v for k, v in inp.items() if not k.startswith("_")})
    return out if inp["tool"] in ("short_term", "long_term", "abandon") else {k: v for k, v in out.items()
                                                                             if k != "put"}  # only where the sheet shows it


def flat(out: dict) -> dict:
    res = {}
    for k, v in out.items():
        if isinstance(v, list):
            for i, row in enumerate(v):
                for j, x in enumerate(row):
                    res[f"{k}.{i}.{j}"] = x
        else:
            res[k] = v
    return res


def report(inp: dict, out: dict) -> str:
    fn, src, head = TOOLS[inp["tool"]]
    rows = [f"{inp['tool']}  (source {src}, {URL}{src})"]
    for k, v in out.items():
        if k == "tree":
            rows += [f"  period {i}: " + "  ".join(f"{x:,.4f}" for x in r) for i, r in enumerate(v)]
        elif isinstance(v, (int, float)):
            rows.append(f"{'>> ' if k == head else '   '}{k:24s} {v:,.6f}")
        elif isinstance(v, str):
            rows.append(f"NOTE {v}")
    return "\n".join(rows)


def selftest() -> int:
    cases, bad = json.loads(CASES.read_text(encoding="utf-8")), 0
    for c in cases:
        got, worst = flat(run(c["inputs"])), 0.0
        for k, his in c["expected"].items():
            err = abs(got[k] - his)
            worst = max(worst, 0.0 if err <= 1e-6 else err / abs(his))
        ok = worst <= 5e-4
        bad += not ok
        head = TOOLS[c["inputs"]["tool"]][2]
        h = c["expected"].get(head, next(iter(c["expected"].values())))
        print(f"{'ok ' if ok else 'BAD'} {c['name']:44s} Damodaran {h:14.6f}  ours {got.get(head, 0):14.6f}  "
              f"worst of {len(c['expected'])} cells {worst:.2e}")
    print(f"{len(cases) - bad}/{len(cases)} cases match Damodaran's spreadsheets (0.05% relative / 1e-6 absolute)")
    return 1 if bad else 0


# ---------------------------------------------------------------- case building (maintenance only) ----------
# source cells -> our inputs (v reads a cell: "A1" on the main sheet or "Sheet!A1"); outputs: our key -> source cell
BS_OUT = lambda d1, n1, d2, n2: {"d1": d1, "N(d1)": n1, "d2": d2, "N(d2)": n2}
YES = lambda x: str(x).strip().lower() == "yes"


def _warrant_in(v, c):  # c = the 8 input cells (price ... shares)
    return dict(zip(["price", "strike", "maturity", "sd", "dividend_yield", "riskfree", "warrants", "shares"],
                    [v(x) for x in c]))


def _employee_in(v):
    inp = {"tool": "employee_options", **_warrant_in(v, [f"B{r}" for r in range(8, 16)]),
           "tax_rate": v("F20") if YES(v("F19")) else 0.0}
    if YES(v("F17")):
        inp["maturity_override"] = v("F18")
    if YES(v("F21")):
        inp["vesting_probability"] = v("F22")
    return inp


def _equity_in(v):
    inp = {"tool": "equity_option", "firm_value": v("F15"), "face_value_debt": v("E38"), "duration": v("E39"),
           "riskfree": v("F42"), "dividends": v("F32") or 0.0}
    a = str(v("F18")).strip().lower()
    if a == "industry":
        inp["industry"] = v("F27")
    elif a == "firm inputs":
        inp["firm_inputs"] = {"sd_stock": v("F21"), "sd_bond": v("F22"), "correlation": v("F23"),
                              "debt_ratio": v("F24")}
    else:
        inp["sd"] = v("F30")
    if YES(v("F34")):
        inp["dividend_yield"] = v("F35")
    return inp


TEMPLATES = {  # stem: (sheet, reader, outputs)
    "warrant": ("Warrant Valuation", lambda v: {"tool": "warrant", **_warrant_in(v, [f"D{r}" for r in range(8, 16)])},
                {"adjusted_price": "C21", **BS_OUT("B26", "B27", "B29", "B30"), "value_per_option": "C32"}),
    "employeeoption": ("Warrant Valuation", _employee_in,
                       {"adjusted_price": "B27", "t": "B29", **BS_OUT("B32", "B33", "B35", "B36"),
                        "value_per_option": "B38", "value_all": "B39", "value_vesting_adjusted": "B40",
                        "value_after_tax": "B41"}),
    "optst": ("NewShTermOptions", lambda v: {
        "tool": "short_term", "price": v("G19"), "sd": v("G21"), "strike": v("F35"), "days": v("F37"),
        "riskfree": v("G40"), "dividends": [[v(f"C{29 + i}"), v(f"D{29 + i}")] for i in range(4) if v("G24") > i],
        "replicate_sheet_bugs": v("G24") > 0},  # the sheet's D53/F54 (see short_term); no dividends = default path
              {"rate_used": "H42", "t": "D44", "pv_dividends": "H44", **BS_OUT("C47", "C48", "C50", "C51"),
               "call": "D53", "put": "F54"}),
    "optlt": ("LT Option ", lambda v: {
        "tool": "long_term", "price": v("C3"), "strike": v("C4"), "maturity": v("C5"), "riskfree": v("G3"),
        "variance": v("G4"), "dividend_yield": v("G5")},
              {**BS_OUT("C8", "C9", "C11", "C12"), "call": "C14", "put": "G14"}),
    "buffettoptions": ("Buffett's options", lambda v: {
        "tool": "long_term", "price": v("F19"), "sd": v("F21"), "strike": v("E29"), "maturity": v("E31"),
        "riskfree_annual": v("F34"),
        **({"dividend_yield": v("F26")} if YES(v("F25")) else {"annual_dividends": v("F23") or 0.0})},
                       {"r": "G37", "y": "G39", **BS_OUT("C42", "C43", "C45", "C46"), "call": "C48", "put": "G48"}),
    "bstobin": ("Sheet1", lambda v: {
        "tool": "binomial", "sd": v("B2"), "riskfree": v("B3"), "dividend_yield": v("B4"),
        "periods_per_year": v("B5"), "price": v("B10")},
                {"up": "B7", "down": "B8", **{f"tree.{i}.{j}": f"{'BCDEFG'[j]}{10 + i}" for i in range(6)
                                                for j in range(i + 1)}}),
    "delay": ("delay.xls", lambda v: {
        "tool": "delay", "pv_cash_flows": v("F5"), "sd": v("F6"), "investment": v("F7"), "years": v("F8"),
        "riskfree": v("F11")}, {"y": "G16", **BS_OUT("C19", "C20", "C22", "C23"), "value": "D25"}),
    "expand": ("ProjectDelay", lambda v: {
        "tool": "expand", "pv_expansion": v("G6"), "sd": v("G7"), "investment": v("G8"), "years": v("G9"),
        "cost_of_delay": v("G10") or 0.0, "riskfree": v("F12")},
               {**BS_OUT("C20", "C21", "C23", "C24"), "value": "D26"}),
    "abandon": ("ProjectDelay", lambda v: {
        "tool": "abandon", "pv_cash_flows": v("F4"), "sd": v("F5"), "project_life": v("F6"),
        "abandonment_value": v("F7"), "years": v("F8"), "riskfree": v("F11")},
                {"y": "G16", **BS_OUT("C19", "C20", "C22", "C23"), "value": "D25"}),
    "project": ("project.xls", lambda v: {
        "tool": "patent", "pv_cash_flows": v("F19"), "sd": v("F21"), "development_cost": v("F24"),
        "years": v("E26"), "riskfree": v("F33"), **({"cost_of_delay": v("F31")} if YES(v("F30")) else {})},
                {"y": "G38", **BS_OUT("C41", "C42", "C44", "C45"), "value": "E47"}),
    "natres": ("NewNatResOption", lambda v: {
        "tool": "natural_resource", "reserves": v("F16"), "price": v("F17"), "unit_cost": v("F18"), "sd": v("F20"),
        "annual_cash_flow": v("F22"), "development_cost": v("F25"), "years": v("E27"), "riskfree": v("F29")},
               {"S": "C32", "y": "G34", **BS_OUT("C37", "C38", "C40", "C41"), "value": "E43"}),
    "flexval": ("Value of Flexibility", lambda v: {
        "tool": "flexibility", "reinvestment_pct": v("F19"), "sd": v("F21"), "financeable_pct": v("F24"),
        "max_pct": v("F25"), "riskfree": v("F27"), "cost_of_capital": v("F30"), "return_on_capital": v("F31"),
        "history": {"reinvestment": [v(f"Reinvestment Needs!B{r}") for r in range(2, 12)],
                    "firm_value": [v(f"Reinvestment Needs!C{r}") for r in range(2, 12)]}},
                {"avg_reinvestment_pct": "Reinvestment Needs!D13", "sd_ln_reinvestment": "Reinvestment Needs!E14",
                 **BS_OUT("C41", "C42", "C44", "C45"), "d1_max": "F41", "N(d1)_max": "F42", "d2_max": "F44",
                 "N(d2)_max": "F45", "value": "E47"}),
    "equity": ("Equity as an Option", _equity_in,
               {"variance": "G46", "y": "G47", **BS_OUT("C50", "C51", "C53", "C54"), "value": "D56",
                "value_of_debt": "D57", "interest_rate_on_debt": "D58", "default_probability": "D59"}),
}
# option sheet of Damodaran's company valuations (two layouts): input cells, adjusted S, BS block, per option, all
OPTION_SHEETS = [([f"D{r}" for r in range(2, 10)], "C15", ("B20", "B21", "B23", "B24"), "C26", "D27"),
                 ([f"B{r}" for r in range(4, 12)], "B17", ("B22", "B23", "B25", "B26"), "B28", "B29")]
# pycel branch tests: overrides of plain-value input cells the default files never exercise
BRANCHES = {
    "warrant": [("dividend yield 3%", {"D12": 0.03}), ("out of the money, 40% dilution", {"D9": 18, "D14": 400})],
    "employeeoption": [("maturity override 1.5y", {"F17": "Yes"}), ("vesting 90%", {"F21": "Yes"}),
                       ("no tax adjustment", {"F19": "No"}),
                       ("all switches on", {"F17": "Yes", "F21": "Yes", "F18": 4, "F20": 0.35})],
    "optst": [("no dividends", {"G24": 0}), ("two dividends", {"G24": 2, "C30": 1.2, "D30": 40, "F37": 120}),
              ("four dividends", {"G24": 4, "C30": 1, "D30": 100, "C31": 1, "D31": 190, "C32": 1.1, "D32": 280,
                                  "F37": 300})],
    "optlt": [("dividend yield 2%", {"G5": 0.02}), ("in the money 5y", {"C4": 400, "C5": 5, "G4": 0.09})],
    "buffettoptions": [("dividends $0.20", {"F23": 0.2}), ("dividend yield override 1%", {"F25": "Yes", "F26": 0.01})],
    "bstobin": [("quarterly periods", {"B5": 4}), ("monthly, no yield", {"B5": 12, "B4": 0})],
    "delay": [("10 years, rf 4%", {"F8": 10, "F11": 0.04}), ("PV above cost", {"F5": 600})],
    "expand": [("cost of delay 3%", {"G10": 0.03}), ("5 years", {"G9": 5, "G8": 120})],
    "abandon": [("rights for full life", {"F8": 25}), ("salvage 220, life 12", {"F7": 220, "F6": 12, "F8": 6})],
    "project": [("cost of delay override 5%", {"F30": "Yes", "F31": 0.05}), ("20-year patent", {"E26": 20})],
    "natres": [("price 15, 20 years", {"F17": 15, "E27": 20}), ("high development cost", {"F25": 2000000})],
    "flexval": [("a negative reinvestment year", {"Reinvestment Needs!B5": -40}),
                ("narrow flexibility band", {"F24": 0.08, "F25": 0.12, "F31": 0.15})],
    "equity": [("firm inputs variance", {"F18": "Firm inputs"}), ("direct variance", {"F18": "Direct Input"}),
               ("industry Steel", {"F27": "Steel"}), ("dividends 5000", {"F32": 5000}),
               ("dividend yield override 2%", {"F34": "Yes", "F35": 0.02})],
}
CIRCULAR = {"warrant", "employeeoption"}


def build_cases(files: list[str]) -> int:
    import os
    import types
    import warnings
    warnings.filterwarnings("ignore")
    cases, bad = [], 0

    def add(name, inp, expected, src):
        nonlocal bad
        got = flat(run(inp))
        worst = max((0.0 if abs(got[k] - x) <= 1e-6 else abs(got[k] - x) / abs(x)) for k, x in expected.items())
        bad += worst > 5e-4
        print(f"{'ok ' if worst <= 5e-4 else 'BAD'} {name:44s} worst {worst:.2e}")
        cases.append({"name": name, "source": src, "inputs": inp, "expected": expected})

    for f in files:
        p = Path(f)
        stem = p.stem
        if p.suffix == ".xls" and stem in TEMPLATES:  # the cached source outputs
            import xlrd
            book = xlrd.open_workbook(str(p))
            sheet0, reader, outs = TEMPLATES[stem]

            def v(ref, book=book, sheet0=sheet0):
                sh, _, a = ref.rpartition("!")
                s = book.sheet_by_name(sh or sheet0)
                col = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed("".join(filter(str.isalpha, a)))))
                row = int("".join(filter(str.isdigit, a)))
                x = s.cell_value(row - 1, col - 1) if row <= s.nrows and col <= s.ncols else ""
                return None if x == "" else x
            if stem == "equity":  # Damodaran's industry standard deviations for the VLOOKUP
                s = book.sheet_by_name("Industry average standard dev")
                rows = {s.cell_value(r, 0): dict(zip(["firms", "sd_equity", "sd_firm_value", "e_to_v", "d_to_v"],
                                                     s.row_values(r, 1, 6))) for r in range(1, s.nrows)}
                TABLE.write_text(json.dumps({"source": "equity.xls, sheet 'Industry average standard dev' (US)",
                                             "url": URL + "equity.xls", "as_of": "2026-01 (file saved 2026-01-26)",
                                             "industries": rows}, indent=1), encoding="utf-8")
            add(f"{p.name} (source default)", reader(v), {k: v(c) for k, c in outs.items()}, URL + p.name)
        elif p.suffix == ".xlsx" and stem in BRANCHES:  # Excel-saved copy of a template: pycel branch tests
            extra = types.ModuleType("pycel_extra")
            exec("import math\ndef normsdist(x):\n    return x if isinstance(x, str) else "  # errors pass through
                 "0.5 * (1 + math.erf(x / math.sqrt(2)))\n"
                 "def stdev(*a):\n    import statistics\n    xs = [v for r in a for row in (r if isinstance(r, tuple)"
                 " else ((r,),)) for v in row if isinstance(v, (int, float)) and not isinstance(v, bool)]\n"
                 "    return statistics.stdev(xs)", extra.__dict__)  # pycel lacks NORMSDIST and STDEV
            sys.modules["pycel_extra"] = extra
            from pycel import ExcelCompiler
            sheet0, reader, outs = TEMPLATES[stem]
            for bname, cells in BRANCHES[stem]:
                xl = ExcelCompiler(filename=str(p), plugins=["pycel_extra"],
                                   cycles={"iterations": 1000, "tolerance": 1e-13} if stem in CIRCULAR else None)

                def q(ref, sheet0=sheet0):
                    sh, _, a = ref.rpartition("!")
                    return "'" + (sh or sheet0).replace("'", "''") + "'!" + a
                for c in outs.values():
                    xl.evaluate(q(c))
                for c, x in cells.items():
                    xl.set_value(q(c), x)
                if stem not in CIRCULAR:  # recalculate() resets the sheet's circular cells to #NUM!
                    xl.recalculate()
                pv = lambda ref, xl=xl: xl.evaluate(q(ref))
                add(f"{stem} branch: {bname}", reader(pv), {k: pv(c) for k, c in outs.items()},
                    f"pycel recalculation of {stem}.xls with {cells}")
        elif p.suffix == ".xlsx":  # the 'Option value' sheet of one of Damodaran's company valuations
            import openpyxl
            wb = openpyxl.load_workbook(str(p), data_only=True)
            if "Option value" not in wb.sheetnames:
                continue
            s = wb["Option value"]
            for ins, adj, block, per, total in OPTION_SHEETS:
                vals = [s[c].value for c in ins]
                if isinstance(vals[0], (int, float)):
                    break
            if not all(isinstance(x, (int, float)) for x in vals + [s[per].value]):
                print(f"skip {p.name}: cached option value is {s[per].value!r}")
                continue
            inp = {"tool": "warrant", **_warrant_in(lambda c: s[c].value, ins)}
            add(f"{p.name} option sheet", inp, {"adjusted_price": s[adj].value,
                                                **{k: s[c].value for k, c in BS_OUT(*block).items()},
                                                "value_per_option": s[per].value, "value_all": s[total].value},
                os.path.basename(f) + " 'Option value' sheet")
    CASES.write_text(json.dumps(cases, indent=1), encoding="utf-8")
    print(f"wrote {len(cases)} cases -> {CASES} ({bad} BAD)")
    return 1 if bad else 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("inputs", nargs="?")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--from-case")
    ap.add_argument("--build-cases", nargs="+", metavar="FILE")
    a = ap.parse_args()
    if a.selftest:
        return selftest()
    if a.build_cases:
        return build_cases(a.build_cases)
    if a.from_case:
        cases = json.loads(CASES.read_text(encoding="utf-8"))
        hit = [c for c in cases if a.from_case.lower() in c["name"].lower()]
        if not hit:
            print("cases:", ", ".join(c["name"] for c in cases))
            return 1
        print(json.dumps({"_source": f"Damodaran's own inputs: {hit[0]['source']}", **hit[0]["inputs"]}, indent=1))
        return 0
    if not a.inputs:
        ap.print_help()
        return 1
    inp = json.loads(Path(a.inputs).read_text(encoding="utf-8"))
    out = run(inp)
    print(json.dumps(out, indent=1) if a.json else report(inp, out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
