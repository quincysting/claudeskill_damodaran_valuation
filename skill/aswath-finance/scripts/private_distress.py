"""Private-company and distress adjustments, transcribed from Damodaran's own calculators.

Sources (https://pages.stern.nyu.edu/~adamodar/pc/<name>; Ginzu: .../pc/fcffsimpleginzu.xlsx):
  liqdisc.xls           'Restricted Stock Regression'!B8 (Silber regression, discount vs a $10m-revenue firm)
                        'Bid-Ask Spread'!B9 (spread regression as the illiquidity discount); Sheet2 = the source table
  minoritydiscount.xls  Sheet1 B7:B9 (value of control as a share of optimal value)
  pvtdiscrate.xls       'Total Beta approach'!B21:D26, 'Synthetic ratings'!D7:D10 (its own 29-row table)
  distress.xls          Sheet1 (annual default probability that prices a bond at its market price; the sheet uses Solver)
  fcffsimpleginzu.xlsx  'Failure Rate worksheet' (cumulative default by rating, 1-10 yrs; BLS survival by age)
  normearn.xls          'Earnings Normalizer & Output'!D14 (three ways to normalize EBIT)
  complscore.xls        Sheet1 E3:E25 (complexity score)
  fcffneg.xls           Valuation / Option Valuation sheets (10-year FCFF model for money-losing firms, NOL shield)

Tools (JSON sections; several may be combined in one file):
  illiquidity     revenues ($m), block (fraction of equity, 1.0), profitable (true), base_discount (0.25),
                  cash_to_value (0), volume_to_value (0)  -> restricted_stock_discount, bid_ask_discount
  minority        optimal_value, status_quo_value, majority_stake (0.51), minority_stake (0.49),
                  probability_of_change (optional; Damodaran's expected-value-of-control formula, 2012-06 post)
  private_coc     unlevered_beta, approach "T" total beta | "P" premium, premium, correlation, riskfree_rate,
                  erp, tax_rate, debt_to_capital, pre_tax_cost_of_debt | synthetic: {ebit, interest_expense}
  bond_distress   coupon_rate, maturity, riskfree_rate, price (per 1000 face) -> annual, 5y, 10y probability
  failure         rating ("BB", "Ba2/BB", ...) + years (10)  |  sector (BLS name) + age  -> failure_probability
  normalized_ebit approach 1 avg EBIT | 2 avg ROC x (book debt+equity) | 3 sector margin x revenues;
                  average_ebit, average_roc, sector_margin, revenues, book_debt, book_equity,
                  history: {revenues: [...], ebit: [...]}  (default sector margin = sum EBIT / sum revenues)
  complexity      answers keyed like COMPLEXITY below -> score by item and total
  fcffneg         see fcffneg() docstring

Usage:
  python private_distress.py inputs.json [--json]
  python private_distress.py --ginzu ginzu_inputs.json   # Ginzu inputs with "failure": {rating|sector+age}
  python private_distress.py --selftest                   # replay data/private_distress_cases.json
  python private_distress.py --build-cases <files>   (maintenance venv: xlrd, openpyxl, pycel)
      files: the source tools/*.xls (cached values), fcffsimpleginzu.xlsx (failure tables -> data json),
      Damodaran's Ginzu valuations (Walgreens2024.xlsx...: failure probability = table value for their rating), and
      Excel "Save As .xlsx" copies of the .xls tools (pycel recalculates them under SCENARIOS; pycel reads no .xls)
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
DATA = HERE.parent / "data" / "private_distress_data.json"
CASES = HERE.parent / "data" / "private_distress_cases.json"


def load_data() -> dict:
    return json.loads(DATA.read_text(encoding="utf-8"))


# ---------------- liqdisc.xls ----------------
def _silber(rev: float, block: float, dern: float) -> float:
    return (100 - math.exp(4.33 + 0.036 * math.log(rev) - 0.142 * math.log(block * 100) + 0.174 * dern)) / 100


def illiquidity(d: dict) -> dict:
    """Restricted-stock (Silber) discount relative to a profitable $10m-revenue firm at base_discount, and the
    bid-ask spread regression. The sheet's B8 hard-codes 0.25 (the B2 'base discount' label is not linked)."""
    rev, dern = d["revenues"], float(d.get("profitable", True))
    block = d.get("block", 1.0)
    restricted = d.get("base_discount", 0.25) - (_silber(10, block, 1) - _silber(rev, block, dern))
    spread = (0.145 - 0.0022 * math.log(rev) - 0.015 * dern - 0.016 * d.get("cash_to_value", 0.0)
              - 0.11 * d.get("volume_to_value", 0.0))
    return {"restricted_stock_discount": restricted, "bid_ask_discount": spread}


# ---------------- minoritydiscount.xls ----------------
def minority(d: dict) -> dict:
    opt, sq, ms = d["optimal_value"], d["status_quo_value"], d.get("minority_stake", 0.49)
    his = opt + ms * sq  # the sheet's B8 reads =B2+B5*B3 (more than the whole firm): a typo; a minority stake is worth status quo
    out = {"value_of_control": opt - sq, "minority_discount": (opt - sq) / opt,
           "majority_stake_value": d.get("majority_stake", 0.51) * opt,
           "minority_stake_value": his if d.get("replicate_sheet_bugs") else ms * sq}
    if not d.get("replicate_sheet_bugs"):
        out["sheet_error"] = (f"the sheet's B8 (=B2+B5*B3) gives {his:,.6g} for the minority stake, more than the whole firm; "
                              "replicate_sheet_bugs=true reproduces it")
    if "probability_of_change" in d:  # 2012-06 / 2014-09 posts: expected value of control
        p = d["probability_of_change"]
        out["expected_value"] = opt * p + sq * (1 - p)
    return out


# ---------------- pvtdiscrate.xls ----------------
def private_coc(d: dict, data: dict | None = None) -> dict:
    t, rf, erp, bu = d["tax_rate"], d["riskfree_rate"], d["erp"], d["unlevered_beta"]
    dc = d["debt_to_capital"]
    lever = 1 + (1 - t) * dc / (1 - dc)
    total_beta = bu / d["correlation"] * lever if d.get("correlation") else None
    if d.get("approach", "T").upper() == "T":
        ke = rf + erp * total_beta
    else:
        ke = bu * lever * erp + rf + d["premium"]
    out = {"debt_to_capital": dc, "market_beta": bu * lever, "total_beta": total_beta, "cost_of_equity": ke}
    if d.get("synthetic"):
        s = d["synthetic"]
        table = (data or load_data())["pvtdiscrate_synthetic_table"]
        cov = s["ebit"] / s["interest_expense"] if s["interest_expense"] > 0 else 10000000
        lo, hi, rating, spread = [row for row in table if row[0] <= cov][-1]  # VLOOKUP, approximate match
        kd = s.get("riskfree_rate", rf) + spread
        out.update(interest_coverage=cov, rating=rating, default_spread=spread)
    else:
        kd = d["pre_tax_cost_of_debt"]
    out.update(pre_tax_cost_of_debt=kd, after_tax_cost_of_debt=kd * (1 - t),
               cost_of_capital=(1 - dc) * ke + dc * kd * (1 - t))
    return out


# ---------------- distress.xls ----------------
def bond_value(coupon: float, maturity: int, rf: float, p: float, face: float = 1000.0) -> float:
    """The sheet's Sheet1 rows 8-32: coupons (and face at maturity) x (1-p)^t, discounted at the riskfree rate.
    The source sheet stops at year 25; we do not truncate."""
    return sum((coupon * face + (face if t == maturity else 0)) * (1 - p) ** t / (1 + rf) ** t
               for t in range(1, int(maturity) + 1))


def bond_distress(d: dict) -> dict:
    c, n, rf, price, face = d["coupon_rate"], d["maturity"], d["riskfree_rate"], d["price"], d.get("face", 1000.0)
    lo, hi = -0.99, 0.999999
    for _ in range(200):  # bisection replaces the sheet's Solver step; value falls as p rises
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if bond_value(c, n, rf, mid, face) > price else (lo, mid)
    p = (lo + hi) / 2
    return {"annual_probability": p, "probability_5y": 1 - (1 - p) ** 5, "probability_10y": 1 - (1 - p) ** 10,
            "riskless_value": bond_value(c, n, rf, 0.0, face)}


# ---------------- Ginzu 'Failure Rate worksheet' ----------------
def broad_rating(rating: str) -> str:
    """'Ba2/BB' or 'BB+' -> 'BB'; CCC, CC and C share the table's 'CCC/C' row."""
    r = rating.split("/")[-1].strip().upper().rstrip("+-")
    return "CCC/C" if r in ("CCC", "CC", "C", "CCC/C") else r


def failure(d: dict, data: dict | None = None) -> dict:
    data = data or load_data()
    if d.get("rating"):
        g, yrs = broad_rating(d["rating"]), d.get("years", 10)
        if not 1 <= yrs <= 10:
            raise ValueError(f"years={yrs}: the table covers 1-10 years")
        table = data["failure_by_rating"]
        if g not in table:
            raise KeyError(f"rating {d['rating']!r} -> {g!r} not in the table {list(table)}")
        return {"failure_probability": table[g][yrs - 1], "basis": f"{g} cumulative default, {yrs} yr"}
    surv, age = data["survival_by_age"][d["sector"]], d["age"]  # S(1..10); the sheet's cells: S(age) - S(10), S(0)=1
    s = 1.0 if age == 0 else surv[age - 1]
    return {"failure_probability": s - surv[9], "basis": f"BLS survival, {d['sector']}, age {age}"}


def ginzu_with_failure(inp: dict, data: dict | None = None) -> dict:
    """Ginzu inputs where 'failure' = {rating[, years]} or {sector, age} -> failure_probability filled in."""
    if "failure" not in inp:
        return inp
    rest = {k: v for k, v in inp.items() if k != "failure"}
    return {**rest, "failure_probability": failure(inp["failure"], data)["failure_probability"]}


# ---------------- normearn.xls ----------------
def normalized_ebit(d: dict) -> dict:
    a, h = d.get("approach", 3), d.get("history")
    margin = d.get("sector_margin")
    if margin is None and h:  # the sheet's D11 default = G21 = sum of EBIT / sum of revenues, last 5 years
        margin = sum(h["ebit"]) / sum(h["revenues"])
    if a == 1:
        avg = d.get("average_ebit", sum(h["ebit"]) / len(h["ebit"]) if h else None)
        val = avg
    elif a == 2:
        val = d["average_roc"] * (d["book_debt"] + d["book_equity"])
    else:
        val = margin * d["revenues"]
    return {"normalized_ebit": val, "sector_margin": margin}


# ---------------- complscore.xls ----------------
# key: (sheet row, weight, kind) ; kind n = number x weight, y = weight if "Yes", n0 = weight if "No"
COMPLEXITY = {
    "businesses": (3, 2, "n"), "one_time_pct": (4, 5, "n"), "unspecified_income_pct": (5, 5, "n"),
    "volatile_items_pct": (6, 5, "n"), "foreign_revenue_pct": (7, 3, "n"), "different_tax_books": (8, 3, "y"),
    "tax_haven_hq": (9, 3, "y"), "volatile_tax_rate": (10, 2, "y"), "volatile_capex": (11, 2, "y"),
    "frequent_acquisitions": (12, 4, "y"), "stock_acquisitions": (13, 4, "y"), "unspecified_wc": (14, 3, "y"),
    "volatile_wc": (15, 2, "y"), "off_balance_sheet_assets": (16, 3, "y"), "big_buybacks": (17, 3, "y"),
    "volatile_roc": (18, 5, "y"), "roc_above_industry": (19, 5, "y"), "emerging_market_pct": (21, 5, "n"),
    "debt_traded": (22, 2, "n0"), "has_rating": (23, 2, "n0"), "off_balance_sheet_debt": (24, 5, "y"),
}


def complexity(d: dict) -> dict:
    yes = lambda v: v is True or str(v).strip().lower() in ("yes", "y", "true")
    items = {}
    for k, (_, w, kind) in COMPLEXITY.items():  # every question must be answered (KeyError otherwise)
        v = d[k]
        items[k] = v * w if kind == "n" else (w if yes(v) == (kind == "y") else 0)
    items["businesses_cost_of_capital"] = d.get("businesses", 0) * 1  # the sheet's E20 = D3 x 1
    return {"items": items, "complexity_score": sum(items.values())}


# ---------------- fcffneg.xls ----------------
def _call(s: float, k: float, t: float, vol: float, r: float, y: float = 0.0) -> float:
    nd = lambda x: 0.5 * (1 + math.erf(x / math.sqrt(2)))
    d1 = (math.log(s / k) + (r - y + vol * vol / 2) * t) / (vol * math.sqrt(t))
    return math.exp(-y * t) * s * nd(d1) - k * math.exp(-r * t) * nd(d1 - vol * math.sqrt(t))


def warrants(price: float, shares: float, o: dict) -> float:
    """The template's 'Option Valuation' sheet: dilution-adjusted Black-Scholes (circular in Excel; iterated here)."""
    w, n = 0.0, o["number"]
    for _ in range(500):
        s = (price * shares + w * n) / (shares + n)
        new = _call(s, o["strike"], o["maturity"], o["volatility"], o["riskfree_rate"], o.get("dividend_yield", 0.0))
        if abs(new - w) < 1e-13:
            break
        w = new
    return w * n


def fcffneg(d: dict) -> dict:
    """10-year FCFF model for firms with negative earnings (fcffneg.xls, 'Valuation' sheet).
    Base: revenues, capital_spending, depreciation, tax_rate, cash, book_debt, book_equity, nol,
          public (true) + price, shares_outstanding, market_debt (= book_debt) | private: book_debt_ratio or debt_ratio
    Risk: cost_of_equity (direct) | beta, riskfree_rate, erp; cost_of_debt (pre-tax)
    Paths (years 1-10): revenue_growth[10]; ebitda_margin[5] (yrs 6-10 move linearly to the stable margin) or [10];
          capex_growth[8..10] (missing years = year-10 revenue growth, the sheet's E70:E71); depreciation_growth[6..10]
          (missing year t = capex growth of year t-1, the sheet's F68:F71); working_capital_pct (scalar or [10])
    Stable: stable_growth, stable_ebitda_margin, stable_working_capital_pct, stable_roc | stable_capex_to_depreciation;
          optional stable_beta, stable_debt_ratio, stable_cost_of_debt (beta, debt ratio and cost of debt move
          linearly to these over years 6-10, as in the sheet's rows 116-119)
    options: {number, strike, maturity, volatility, riskfree_rate, dividend_yield}  (valued at price)
    replicate_sheet_bugs: true reproduces two slips in the source file: years 8-10 use the previous year's EBITDA margin
          (J100:L100 point one row up) and year-1 taxes are always zero (C104 compares the text label 'NOL')."""
    T, g_s, bug = d["tax_rate"], d["stable_growth"], d.get("replicate_sheet_bugs", False)
    g = [0.0] + list(d["revenue_growth"])
    m = list(d["ebitda_margin"])
    M = d["stable_ebitda_margin"]
    if len(m) == 5:
        m += [m[4] - (m[4] - M) / 5 * k for k in range(1, 6)]
    m = [0.0] + m
    cg = list(d["capex_growth"])
    cg += [g[10]] * (10 - len(cg))
    dg = list(d["depreciation_growth"])
    dg += [cg[t - 1] for t in range(len(dg), 10)]
    cg, dg = [0.0] + cg, [0.0] + dg
    wc = d["working_capital_pct"]
    wc = [0.0] + (list(wc) if isinstance(wc, list) else [wc] * 10)

    rev, dep, capex, nol = [d["revenues"]], [d["depreciation"]], [d["capital_spending"]], [d["nol"]]
    ebit, tax, fcff, tr = [0.0], [0.0], [0.0], [0.0]
    for t in range(1, 11):
        rev.append(rev[-1] * (1 + g[t]))
        mu = m[t - 1] if bug and t >= 8 else m[t]
        dep.append(dep[-1] * (1 + dg[t]))
        capex.append(capex[-1] * (1 + cg[t]))
        e = rev[t] * mu - dep[t]
        prev = nol[-1]
        if bug and t == 1:
            x, n_ = 0.0, prev - e
        else:
            x = 0.0 if e <= 0 or prev > e else (e - prev) * T
            n_ = prev - e if prev > e else 0.0
        ebit.append(e), tax.append(x), nol.append(n_)
        tr.append(x / e if e else 0.0)  # the sheet's row 115 (Excel would show #DIV/0! at EBIT = 0)
        fcff.append(e - x + dep[t] - capex[t] - (rev[t] - rev[t - 1]) * wc[t])

    if d.get("public", True):
        mv_e, debt = d["price"] * d["shares_outstanding"], d.get("market_debt", d["book_debt"])
        eq_ratio = mv_e / (debt + mv_e)
    else:
        debt = d["book_debt"]
        eq_ratio = (d["book_equity"] / (d["book_debt"] + d["book_equity"]) if d.get("book_debt_ratio")
                    else 1 - d["debt_ratio"])
    dr0 = 1 - eq_ratio
    beta, sb, kd = d.get("beta", 0.0), d.get("stable_beta"), d["cost_of_debt"]
    skd, sdr = d.get("stable_cost_of_debt", kd), d.get("stable_debt_ratio")
    cum, pv = [1.0], [0.0]
    for t in range(1, 11):
        k = max(t - 5, 0)
        b = beta - (beta - sb) / 5 * k if sb is not None else beta
        ke = d["cost_of_equity"] if "cost_of_equity" in d else d["riskfree_rate"] + b * d["erp"]
        kd_t = (kd - (kd - skd) * k / 5) * (1 - tr[t])
        dr = dr0 - (dr0 - sdr) / 5 * k if sdr is not None else dr0
        cum.append(cum[-1] * (1 + ke * (1 - dr) + kd_t * dr))
        pv.append(fcff[t] / cum[t])

    # terminal year (the sheet's column M)
    r11 = rev[10] * (1 + g_s)
    dep11 = dep[10] * (1 + g_s)
    e11 = r11 * M - dep11
    x11 = 0.0 if e11 <= 0 or nol[10] > e11 else (e11 - nol[10]) * T
    dwc11 = (r11 - rev[10]) * d["stable_working_capital_pct"]
    nopat11 = e11 - x11
    cap11 = (g_s / d["stable_roc"] * nopat11 + dep11 - dwc11 if d.get("stable_roc")
             else dep11 * d["stable_capex_to_depreciation"])
    fcff11 = nopat11 + dep11 - cap11 - dwc11
    b_s = sb if sb is not None else beta
    ke_s = d["riskfree_rate"] + b_s * d["erp"] if "erp" in d else d["cost_of_equity"]
    e_s = 1 - sdr if sdr is not None else eq_ratio
    wacc_s = ke_s * e_s + skd * (1 - T) * (1 - e_s)
    tv = fcff11 / (wacc_s - g_s)

    firm = sum(pv[1:]) + tv / cum[10]
    equity = firm + d["cash"] - debt
    opts = warrants(d["price"], d["shares_outstanding"], d["options"]) if d.get("options") else 0.0
    out = {"pv_high_growth": sum(pv[1:]), "terminal_value": tv, "pv_terminal": tv / cum[10],
           "value_of_firm": firm, "value_of_equity": equity, "options_value": opts,
           "equity_in_common": equity - opts, "value_per_share": (equity - opts) / d["shares_outstanding"],
           "stable_cost_of_capital": wacc_s,
           "years": [{"t": t, "revenue": rev[t], "ebit": ebit[t], "tax": tax[t], "nol": nol[t], "fcff": fcff[t],
                      "cum_discount": cum[t]} for t in range(1, 11)] + [{"t": "terminal", "fcff": fcff11}]}
    his = None if bug else fcffneg({**d, "replicate_sheet_bugs": True})["value_per_share"]
    if his is not None and abs(his - out["value_per_share"]) > 1e-9 * abs(out["value_per_share"]):
        out["sheet_error"] = (f"Damodaran's fcffneg as written gives {his:.6g}/share (years 8-10 margin one row up; "
                              "year-1 tax always 0); replicate_sheet_bugs=true reproduces it")
    return out


TOOLS = {"illiquidity": illiquidity, "minority": minority, "private_coc": private_coc,
         "bond_distress": bond_distress, "failure": failure, "normalized_ebit": normalized_ebit,
         "complexity": complexity, "fcffneg": fcffneg}


# ---------------- --build-cases (maintenance: xlrd / openpyxl / pycel) ----------------
# Each extractor reads one of Damodaran's workbooks through get(sheet, cell) and returns cases. The same extractor
# runs on the source .xls (cached values) and on an Excel->.xlsx copy recalculated by pycel under SCENARIOS.
def _cells(get, sheet: str, cols: str, row: int) -> list:
    return [get(sheet, f"{c}{row}") for c in cols]


def _x_liqdisc(get, his: bool) -> list:
    R, B = "Restricted Stock Regression", "Bid-Ask Spread"
    base = {"block": get(R, "B4"), "profitable": get(R, "B5")}
    out = [{"name": "restricted stock discount", "tool": "illiquidity", "inputs": {"revenues": get(R, "B3"), **base},
            "expected": {"restricted_stock_discount": get(R, "B8")}},
           {"name": "bid-ask spread discount", "tool": "illiquidity",
            "inputs": {"revenues": get(B, "B3"), "profitable": get(B, "B4"), "cash_to_value": get(B, "B5"),
                       "volume_to_value": get(B, "B6")}, "expected": {"bid_ask_discount": get(B, "B9")}}]
    for r in range(2, 18) if his else []:  # Sheet2: the source table (4 dp), 100% block
        for col, prof in (("B", 1), ("C", 0)):
            out.append({"name": f"Sheet2 table rev {get('Sheet2', f'A{r}'):g} {'profitable' if prof else 'loss'}",
                        "tool": "illiquidity", "inputs": {"revenues": get("Sheet2", f"A{r}"), "profitable": prof},
                        "expected": {"restricted_stock_discount": get("Sheet2", f"{col}{r}")}})
    return out


def _x_minority(get, his: bool) -> list:
    s = "Sheet1"
    return [{"name": "minority discount", "tool": "minority",
             "inputs": {"optimal_value": get(s, "B2"), "status_quo_value": get(s, "B3"),
                        "majority_stake": get(s, "B4"), "minority_stake": get(s, "B5")},
             "expected": {"minority_discount": get(s, "B9"), "majority_stake_value": get(s, "B7")}}]


def _x_pvt(get, his: bool) -> list:
    T, S = "Total Beta approach", "Synthetic ratings"
    d = {"unlevered_beta": get(T, "B3"), "approach": get(T, "B4"), "premium": get(T, "B5"),
         "correlation": get(T, "B6"), "riskfree_rate": get(T, "B7"), "erp": get(T, "B8"), "tax_rate": get(T, "B13"),
         "debt_to_capital": get(T, "B17") if get(T, "B16") == "Y" else get(T, "B18")}
    exp = {"total_beta": get(T, "B22"), "cost_of_equity": get(T, "B26"), "after_tax_cost_of_debt": get(T, "C26"),
           "cost_of_capital": get(T, "D26")}
    if get(T, "B11") == "Y":
        d["synthetic"] = {"ebit": get(S, "F3"), "interest_expense": get(S, "F4"), "riskfree_rate": get(S, "F5")}
        exp["rating"] = get(S, "D8")
    else:
        d["pre_tax_cost_of_debt"] = get(T, "B12")
    return [{"name": "private cost of capital", "tool": "private_coc", "inputs": d, "expected": exp}]


def _x_distress(get, his: bool) -> list:
    s = "Sheet1"  # branches keep the sheet's D36 fixed, so the bond value D34 is the price that p produces
    d = {"coupon_rate": get(s, "C2"), "maturity": get(s, "C3"), "riskfree_rate": get(s, "C4"),
         "price": get(s, "C5") if his else get(s, "D34")}
    return [{"name": "probability of distress from bond price", "tool": "bond_distress", "inputs": d,
             "expected": {"annual_probability": get(s, "D36"), "probability_5y": get(s, "D37"),
                          "probability_10y": get(s, "D38")}}]


def _x_normearn(get, his: bool) -> list:
    E, M = "Earnings Normalizer & Output", "Master Inputs Start here"
    d = {"approach": get(E, "D2"), "average_ebit": get(E, "D5"), "average_roc": get(E, "D8"),
         "revenues": get(M, "B15"), "book_debt": get(M, "B18"), "book_equity": get(M, "B19"),
         "history": {"revenues": _cells(get, E, "BCDEF", 19), "ebit": _cells(get, E, "BCDEF", 20)}}
    return [{"name": f"normalized EBIT (approach {d['approach']:g})", "tool": "normalized_ebit", "inputs": d,
             "expected": {"normalized_ebit": get(E, "D14"), "sector_margin": get(E, "D11")}}]


def _x_compl(get, his: bool) -> list:
    d = {k: get("Sheet1", f"D{row}") for k, (row, _, _) in COMPLEXITY.items()}
    return [{"name": "complexity score", "tool": "complexity", "inputs": d,
             "expected": {"complexity_score": get("Sheet1", "E25")}}]


def _x_fcffneg(get, his: bool) -> list:
    V, O = "Valuation", "Option Valuation"
    yes = lambda c: str(get(V, c)).lower() == "yes"
    d = {"revenues": get(V, "D28"), "capital_spending": get(V, "D25"), "depreciation": get(V, "D26"),
         "tax_rate": get(V, "D27"), "cash": get(V, "D31"), "book_debt": get(V, "D32"), "book_equity": get(V, "D33"),
         "nol": get(V, "D34"), "public": yes("F37"), "price": get(V, "F39"), "shares_outstanding": get(V, "F40"),
         "market_debt": get(V, "F41"), "beta": get(V, "D52"), "riskfree_rate": get(V, "D53"), "erp": get(V, "D54"),
         "cost_of_debt": get(V, "E56"), "revenue_growth": [get(V, f"C{r}") for r in range(62, 72)],
         "ebitda_margin": [get(V, f"D{r}") for r in range(62, 67)],
         "capex_growth": [get(V, f"E{r}") for r in range(62, 70)],
         "depreciation_growth": [get(V, f"F{r}") for r in range(62, 68)], "working_capital_pct": get(V, "G62"),
         "stable_growth": get(V, "E73"), "stable_ebitda_margin": get(V, "E74"),
         "stable_working_capital_pct": get(V, "E75"), "replicate_sheet_bugs": True}
    if not d["public"]:
        d["book_debt_ratio"] = yes("F43")
        if not d["book_debt_ratio"]:
            d["debt_ratio"] = get(V, "F44")
    if yes("E49"):
        d["cost_of_equity"] = get(V, "E50")
    for flag, key, cell in (("E77", "stable_beta", "E78"), ("F80", "stable_debt_ratio", "F81"),
                            ("E83", "stable_cost_of_debt", "E84")):
        if yes(flag):
            d[key] = get(V, cell)
    d.update({"stable_roc": get(V, "F88")} if yes("F87") else {"stable_capex_to_depreciation": get(V, "F89")})
    if get(O, "D8"):
        d["options"] = {"number": get(O, "D8"), "strike": get(O, "D3"), "maturity": get(O, "D4"),
                        "volatility": get(O, "D5"), "dividend_yield": get(O, "D6"), "riskfree_rate": get(O, "D7")}
    return [{"name": "FCFF valuation, money-losing firm", "tool": "fcffneg", "inputs": d,
             "expected": {"value_per_share": get(V, "F140"), "value_of_firm": get(V, "F134"),
                          "pv_high_growth": get(V, "F132"), "options_value": get(V, "F138")}}]


EXTRACT = {"liqdisc": _x_liqdisc, "minoritydiscount": _x_minority, "pvtdiscrate": _x_pvt, "distress": _x_distress,
           "normearn": _x_normearn, "complscore": _x_compl, "fcffneg": _x_fcffneg}
_TB, _EN, _V = "Total Beta approach!", "Earnings Normalizer & Output!", "Valuation!"
# Damodaran's template recalculated (pycel) under inputs the source file never uses; only plain-value cells are overridden
SCENARIOS = {
    "liqdisc": [("rev 5, loss-making", {"Restricted Stock Regression!B3": 5, "Restricted Stock Regression!B5": 0,
                                        "Bid-Ask Spread!B3": 5}),
                ("rev 1000, 10% block", {"Restricted Stock Regression!B3": 1000, "Restricted Stock Regression!B4": 0.1,
                                         "Bid-Ask Spread!B3": 1000}),
                ("cash 20%, volume 5%", {"Bid-Ask Spread!B5": 0.2, "Bid-Ask Spread!B6": 0.05})],
    "minoritydiscount": [("30% stake", {"Sheet1!B2": 200, "Sheet1!B3": 150, "Sheet1!B4": 0.7, "Sheet1!B5": 0.3})],
    "pvtdiscrate": [("premium P", {_TB + "B4": "P"}), ("synthetic rating", {_TB + "B11": "Y"}),
                    ("synthetic cov 1.1 + target 30%", {_TB + "B11": "Y", "Synthetic ratings!F4": 9000,
                                                        _TB + "B16": "N", _TB + "B18": 0.3}),
                    ("correlation 0.25", {_TB + "B6": 0.25})],
    "distress": [("10% 20yr rf 4%", {"Sheet1!C2": 0.1, "Sheet1!C3": 20, "Sheet1!C4": 0.04}),
                 ("5% 3yr p 3%", {"Sheet1!C2": 0.05, "Sheet1!C3": 3, "Sheet1!D36": 0.03})],
    "normearn": [("approach 1", {_EN + "D2": 1}), ("approach 2", {_EN + "D2": 2}),
                 ("history changed", {_EN + "F19": 4000, _EN + "F20": -50, "Master Inputs Start here!B15": 5000})],
    "complscore": [("mostly No", {"Sheet1!D3": 4, "Sheet1!D8": "No", "Sheet1!D12": "No", "Sheet1!D22": "No",
                                  "Sheet1!D23": "No", "Sheet1!D24": "Yes", "Sheet1!D21": 0.8})],
    "fcffneg": [("direct cost of equity", {_V + "E49": "Yes", _V + "E50": 0.13}), ("beta unchanged", {_V + "E77": "No"}),
                ("debt ratio to 20%", {_V + "F80": "Yes", _V + "F81": 0.2}),
                ("capex = 110% dep.", {_V + "F87": "No"}), ("private, book D/E", {_V + "F37": "No", _V + "F43": "Yes"}),
                ("private, target 30%", {_V + "F37": "No", _V + "F43": "No", _V + "F44": 0.3}),
                ("no options", {"Option Valuation!D8": 0}), ("no NOL", {_V + "D34": 0}),
                ("yr-1 EBIT > NOL", {_V + "D62": 0.6, _V + "D34": 100}),
                ("g 4%, ROC 12%", {_V + "E73": 0.04, _V + "F88": 0.12})],
}
# the two source-sheet slips in fcffneg.xls, corrected in a copy to verify the default (replicate_sheet_bugs = false) path
FCFFNEG_FIX = {"J100": "=IF($E$46<J98,0,J99*(1-$D69))", "K100": "=IF($E$46<K98,0,K99*(1-$D70))",
               "L100": "=IF($E$46<L98,0,L99*(1-$D71))",
               "C104": "=IF(C103>0,IF(D34>C103,0,(C103-D34)*$D$27),0)", "C111": "=IF(D34>C103,D34-C103,0)"}


def _rc(addr: str) -> tuple[int, int]:
    col = "".join(ch for ch in addr if ch.isalpha())
    c = 0
    for ch in col:
        c = c * 26 + ord(ch.upper()) - 64
    return int(addr[len(col):]) - 1, c - 1


def _pycel_cases(path: Path, stem: str, fixed: bool = False) -> list:
    import types
    extra = types.ModuleType("pycel_extra")  # pycel lacks Excel's NORMSDIST (the template's option sheet)
    exec("import math\ndef normsdist(x):\n    return 0.5 * (1 + math.erf(x / math.sqrt(2)))", extra.__dict__)
    sys.modules["pycel_extra"] = extra
    from pycel import ExcelCompiler

    out = []
    for name, cells in [("baseline", {})] + SCENARIOS[stem] if not fixed else \
            [("baseline", {})] + [s for s in SCENARIOS[stem] if s[0] in ("yr-1 EBIT > NOL", "capex = 110% dep.")]:
        xl = ExcelCompiler(filename=str(path), cycles={"iterations": 200, "tolerance": 1e-12}, plugins=["pycel_extra"])
        get = lambda s, a: (lambda v: None if v in ("", None) else v)(xl.evaluate(f"{s}!{a}"))
        EXTRACT[stem](get, False)  # builds the cell map
        for c, v in cells.items():
            xl.set_value(c, v)
        xl.recalculate()
        for case in EXTRACT[stem](get, False):
            if fixed:
                case["inputs"]["replicate_sheet_bugs"] = False
            out.append({**case, "name": f"{stem} template{' (slips fixed)' if fixed else ''} {name}: {case['name']}"})
    return out


def _ginzu_tables(path: Path) -> dict:
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True)["Failure Rate worksheet"]
    v = lambda a: ws[a].value
    return {"failure_by_rating": {v(f"A{r}"): [v(f"{c}{r}") for c in "BCDEFGHIJK"] for r in range(6, 13)},
            "survival_by_age": {v(f"{c}18"): [v(f"{c}{r}") for r in range(19, 29)] for c in "OPQRSTUVWX"},
            "_cells": ws}


def build(files: list[str]) -> int:
    import warnings
    warnings.filterwarnings("ignore")
    data = json.loads(DATA.read_text(encoding="utf-8")) if DATA.exists() else {}
    cases, ginzu_files = [], []
    for f in map(Path, files):
        stem = f.stem.replace("_fixed", "")
        if f.suffix == ".xls" and stem in EXTRACT:
            import xlrd
            wb = xlrd.open_workbook(str(f))

            def get(sheet, addr, wb=wb):
                sh, (r, c) = wb.sheet_by_name(sheet), _rc(addr)
                v = sh.cell_value(r, c) if r < sh.nrows and c < sh.ncols else ""
                return None if v == "" else v
            cases += [{**c, "name": f"{f.name} (source file): {c['name']}"} for c in EXTRACT[stem](get, True)]
            if stem == "pvtdiscrate":
                data["pvtdiscrate_synthetic_table"] = [[get("Synthetic ratings", f"{c}{r}") for c in "ABCD"]
                                                       for r in range(15, 30)]
        elif f.name == "fcffsimpleginzu.xlsx":
            t = _ginzu_tables(f)
            ws = t.pop("_cells")
            data.update(t, source="Damodaran, fcffsimpleginzu.xlsx 'Failure Rate worksheet' (cumulative default "
                        "rates by rating; BLS 2022 survival rates by age); pvtdiscrate.xls 'Synthetic ratings'",
                        url="https://pages.stern.nyu.edu/~adamodar/pc/fcffsimpleginzu.xlsx",
                        as_of="Ginzu template January 2026 (same table in Damodaran's 2023-2026 valuation files)")
            his = [("A14 text: BB, 10 yr", {"rating": "BB", "years": 10}, ws["K10"].value),
                   ("A31 text: technology (Information), age 6", {"sector": "Information", "age": 6}, ws["J25"].value)]
            his += [(f"cell {c}{r}: {ws[c + '18'].value}, age {ws['B' + str(r)].value}",
                     {"sector": ws[c + "18"].value, "age": ws["B" + str(r)].value}, ws[f"{c}{r}"].value)
                    for c, r in (("L", 19), ("L", 22), ("C", 24), ("F", 27), ("K", 28))]
            cases += [{"name": f"fcffsimpleginzu.xlsx: {n}", "tool": "failure", "inputs": i,
                       "expected": {"failure_probability": e}} for n, i, e in his]
        elif f.suffix == ".xlsx" and stem in EXTRACT:
            if stem == "fcffneg" and "_fixed" not in f.stem:  # also build the corrected copy
                import openpyxl
                wb = openpyxl.load_workbook(f)
                for c, formula in FCFFNEG_FIX.items():
                    wb["Valuation"][c] = formula
                fixed = f.with_name("fcffneg_fixed.xlsx")
                wb.save(fixed)
                cases += _pycel_cases(fixed, stem, fixed=True)
            cases += _pycel_cases(f, stem)
        elif f.suffix == ".xlsx":
            ginzu_files.append(f)
    for f in ginzu_files:  # Damodaran's valuations whose failure probability is the table's value for their rating
        import openpyxl
        wb = openpyxl.load_workbook(f, data_only=True)
        coc = next(wb[s] for s in wb.sheetnames if s.lower().startswith("cost of capital"))
        rating = next(coc.cell(c.row, c.column + 1).value for row in coc.iter_rows() for c in row
                      if isinstance(c.value, str) and c.value.startswith("If actual rating, input the rating"))
        ins = wb["Input sheet"]
        p = next(ins.cell(c.row, 2).value for row in ins.iter_rows() for c in row
                 if isinstance(c.value, str) and c.value.startswith("If yes, enter the probability of failure"))
        table = failure({"rating": rating, "years": 10}, data)["failure_probability"]
        calib = {c["name"]: c for c in json.loads((DATA.parent / "calibration_cases.json").read_text(encoding="utf-8"))}
        if not p or abs(p - table) > 1e-12 or f.stem not in calib:
            print(f"skip {f.name}: failure probability {p} vs table {table:.4f} for {rating} (Damodaran's own judgment)")
            continue
        inputs = {k: v for k, v in calib[f.stem]["inputs"].items() if k != "failure_probability"}
        cases.append({"name": f"{f.name}: Ginzu value with failure from rating {rating}", "tool": "ginzu",
                      "inputs": {**inputs, "failure": {"rating": rating, "years": 10}},
                      "expected": {"value_per_share": calib[f.stem]["expected_value_per_share"]}})
    DATA.write_text(json.dumps(data, indent=1), encoding="utf-8")
    CASES.write_text(json.dumps(cases, indent=1), encoding="utf-8")
    print(f"wrote {len(cases)} cases -> {CASES}\nwrote data -> {DATA}")
    return selftest()


# ---------------- CLI ----------------
def report(name: str, out: dict) -> str:
    rows = [f"== {name}"]
    for k, v in out.items():
        if k == "years":
            rows.append("   year  revenue      EBIT       tax       NOL      FCFF")
            rows += [f"   {y['t']:>4}  {y['revenue']:9,.0f} {y['ebit']:9,.0f} {y['tax']:9,.0f} {y['nol']:9,.0f} "
                     f"{y['fcff']:9,.0f}" if "revenue" in y else f"   terminal FCFF {y['fcff']:,.0f}" for y in v]
        elif isinstance(v, dict):
            rows += [f"   {kk:28s} {vv:g}" for kk, vv in v.items()]
        elif isinstance(v, float):
            rows.append(f"   {k:28s} {v:.6g}")
        else:
            rows.append(f"   {k:28s} {v}")
    return "\n".join(rows)


def close(got, exp) -> bool:
    if isinstance(exp, str):
        return got == exp
    return abs(got - exp) <= max(1e-6, 5e-4 * abs(exp))  # spec: 0.05% relative, or 1e-6 absolute for rates


def selftest() -> int:
    cases = json.loads(CASES.read_text(encoding="utf-8"))
    bad = 0
    for c in cases:
        if c["tool"] == "ginzu":
            import ginzu
            out = ginzu.value(ginzu_with_failure(c["inputs"]))
        else:
            out = TOOLS[c["tool"]](c["inputs"])
        misses = [(k, v, out[k]) for k, v in c["expected"].items() if not close(out[k], v)]
        bad += bool(misses)
        shown = ", ".join(f"{k} {v:.6g}" if isinstance(v, float) else f"{k} {v}" for k, v in c["expected"].items())
        print(f"{'BAD' if misses else 'ok '} {c['name'][:58]:58s} {shown[:90]}")
        for k, v, got in misses:
            print(f"      {k}: Damodaran {v}  ours {got}")
    print(f"{len(cases) - bad}/{len(cases)} cases match Damodaran's spreadsheets (0.05% rel. or 1e-6 abs.)")
    return 1 if bad else 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="?")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--ginzu", action="store_true", help="inputs are Ginzu inputs with a 'failure' section")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--build-cases", nargs="+", metavar="FILE")
    a = ap.parse_args()
    if a.selftest:
        return selftest()
    if a.build_cases:
        return build(a.build_cases)
    inp = json.loads(Path(a.inputs).read_text(encoding="utf-8"))
    if a.ginzu:
        import ginzu
        inp = ginzu_with_failure({k: v for k, v in inp.items() if not k.startswith("_")})
        out = ginzu.value(inp)
        print(json.dumps(out, indent=1) if a.json else
              f"failure_probability {inp.get('failure_probability', 0):.4f}\n" + ginzu.report(inp, out))
        return 0
    results = {k: TOOLS[k](v) for k, v in inp.items() if k in TOOLS}
    if not results:
        print("no tool section found; expected one of:", ", ".join(TOOLS))
        return 1
    print(json.dumps(results, indent=1) if a.json else "\n".join(report(k, v) for k, v in results.items()))
    return 0


if __name__ == "__main__":
    sys.exit(main())
