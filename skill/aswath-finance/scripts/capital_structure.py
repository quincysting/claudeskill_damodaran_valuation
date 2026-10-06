"""Optimal debt ratio (cost-of-capital approach), buyback effects and FCFE cash-return capacity, transcribed from
Damodaran's own spreadsheets and checked against the source files (`--selftest`).

Sources (URL = https://pages.stern.nyu.edu/~adamodar/pc/<name>):
  capstru.xlsx            'Inputs', 'Optimal CS Worksheet' (rows 4-97), 'Operating leases', 'Default Spreads and
                          Ratios', 'Repurchase price Worksheet'. Template = Disney inputs (Feb 2023), Jan 2026 tables.
  blog/AdaniCapStru.xlsx, blog/AdaniGroupCapStru.xlsx   Damodaran's Feb 2023 applications (older layout, same logic)
  buybacks.xls            'Value Effect' + 'EPS Effect';  buybackcalculator.xlsx (2026) = its EPS sheet
  levbeta.xls, ratings.xls   Hamada relevering table; stand-alone synthetic rating
  datasets/divfcfe*.xls   Damodaran's FCFE and cash-return definitions (sheet 'Variables & FAQ')

optimal() mirrors the worksheet: at each debt ratio 0%..90% of current (D+E) it relevers beta, sets interest on
the new debt, finds the synthetic rating from interest coverage (circular: rate -> interest -> coverage -> rating,
iterated like Excel), then cost of debt, WACC and firm value = EV x (1 + (WACC_now - WACC)/(WACC - g)), g = implied
growth capped at the riskfree rate. Optimal = debt ratio with the highest value.

Usage:
  python capital_structure.py inputs.json [--json]     inputs: {"optimal": {...}, "buyback": {...}, "cash_return": {...}}
  python capital_structure.py --from-case AdaniEnterprises     the source inputs for a calibration case, as JSON
  python capital_structure.py --selftest                replay data/capital_structure_cases.json
  python capital_structure.py --build-cases <files...>  (maintenance, openpyxl/xlrd) rebuild data/ from the source files

optimal inputs (same units throughout; rates as decimals; defaults = the source template):
  ebitda, depreciation, capex, interest_expense, book_debt, cash, shares_outstanding, price, beta, riskfree_rate,
  pre_tax_cost_of_debt (B17), current_rating (label, display only)
  country            fills marginal_tax_rate / erp / country_default_spread from Damodaran's Jan 2026 country table if absent
  marginal_tax_rate, erp, country_default_spread
  market_value_debt  (B24/B25)  else debt_maturity -> the spreadsheet's PV estimate (B26/B27)  else book debt
  leases: {current_expense, commitments: [y1..y5], beyond}         (B28 + 'Operating leases' sheet)
  interest_limit: {measure: "EBITDA"|"EBIT", max_pct: 0.3}         (B30-B32; default no limit)
  indirect_bankruptcy_costs  false (B34; EBITDA drops with rating, the sheet's 'Drop in EBITDA' column)
  firm_type          1 large/safe | 2 small/risky | 3 infrastructure/utility (B42; default 1)
  refinance          false (B43: existing debt keeps its rate; true = all debt at the new rate)
  synthetic_current_rating  false (B44: current cost of debt from coverage, not B17)
  buyback_price      optional ('Repurchase price Worksheet' B4) -> value/share after a debt-funded recap
  rating_tables      optional {large, small, utility}: rows [low, high, rating, spread, drop]; default Damodaran's Jan 2026
  ratings_at         optional {"0.4": "Caa/CCC"}: pick another self-consistent rating at a debt ratio (see optimal())
buyback inputs: price, shares_outstanding, shares_bought, buyback_price, cash_used, net_income, cash,
  interest_income_from_cash (default cash x 0.5%, the spreadsheet calculator), new_debt_rate, tax_rate, and either
  post_buyback_price (calculator) or beta, debt, cost_of_debt, riskfree_rate, erp (buybacks.xls value effect);
  fair_value optional (value transfer)
cash_return inputs: net_income, dividends, buybacks, stock_issuances, depreciation, capex (incl. R&D, acquisitions),
  change_in_working_capital (0), new_debt (0), debt_repaid (0)  -- or fcfe_pre_debt / fcfe_post_debt directly
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
DATA = HERE.parent / "data" / "capital_structure_data.json"
CASES = HERE.parent / "data" / "capital_structure_cases.json"
GINZU_DATA = HERE.parent / "data" / "ginzu_data.json"
RATIOS = [round(0.1 * i, 1) for i in range(10)]
ITERATIONS = 100  # Excel's default iteration count for the circular source sheet


def lookup(x: float, rows: list) -> list:
    """Excel VLOOKUP(x, table, k) with approximate match: last row whose lower bound <= x."""
    hit = None
    for row in rows:
        if row[0] > x:
            break
        hit = row
    if hit is None:
        raise ValueError(f"coverage {x} below the rating table")
    return hit


def excel_round(x: float) -> int:
    return int(math.copysign(math.floor(abs(x) + 0.5), x))


def worksheet_table(tables: dict, firm_type: int, rf: float, country_spread: float) -> list:
    """'Optimal CS Worksheet'!D83:H97 -> [low, rating, pre-tax rate, EBITDA drop]. Coverage bounds follow the firm
    type, but rating/spread/drop come from table 1 for type 1 and table 2 for types 2 AND 3 (the spreadsheet formulas never
    read the utility spreads here)."""
    bounds = tables[{1: "large", 2: "small", 3: "utility"}[firm_type]]
    src = tables["large" if firm_type == 1 else "small"]
    return [[b[0], s[2], s[3] + rf + country_spread, s[4]] for b, s in zip(bounds, src)]


def sheet_rate(tables: dict, firm_type: int, coverage: float, rf: float, country_spread: float) -> float:
    """'Default Spreads and Ratios'!D10: rf + spread of the chosen table (type 3 uses its own utility spreads)."""
    return rf + lookup(coverage, tables[{1: "large", 2: "small", 3: "utility"}[firm_type]])[3] + country_spread


def lease_debt(le: dict, kd: float, ebit: float, interest: float, dep: float) -> dict:
    """'Operating leases' sheet: PV of commitments at the pre-tax cost of debt; years beyond 5 as an annuity."""
    c, beyond = le["commitments"], le.get("beyond", 0.0)
    n6 = excel_round(beyond / (sum(c[:5]) / 5))
    annual = (beyond / n6 if n6 > 0 else beyond) if beyond > 0 else 0.0
    pv = sum(x / (1 + kd) ** t for t, x in enumerate(c[:5], 1))
    pv += annual * (1 - (1 + kd) ** -n6) / kd / (1 + kd) ** 5 if n6 > 0 else beyond / (1 + kd) ** 6
    return {"debt": pv, "ebit": ebit + le["current_expense"] - pv / (5 + n6),
            "interest": interest + pv * kd, "depreciation": dep + pv / (5 + n6)}


def with_country(inp: dict) -> dict:
    """The spreadsheet's Inputs B15/B38/B39 are VLOOKUPs on the country table when linked (Adani files)."""
    need = [k for k in ("marginal_tax_rate", "erp", "country_default_spread") if k not in inp]
    if not need:
        return inp
    if "country" not in inp:
        raise KeyError(f"give {need} or a country")
    row = json.loads(GINZU_DATA.read_text(encoding="utf-8"))["countries"][inp["country"]]
    src = {"marginal_tax_rate": row["tax_rate"], "erp": row["erp"], "country_default_spread": row["default_spread"]}
    return {**{k: src[k] for k in need}, **inp}


def tax_rates(interest: float, ebit: float, ebitda: float, t: float, limit: dict | None) -> tuple[float, float]:
    """Rows 58 & 69: (tax on taxable income, tax rate used in the after-tax cost of debt)."""
    base = None if not limit else (ebit if limit["measure"] == "EBIT" else ebitda) * limit["max_pct"]
    capped = base is not None and interest >= base
    tax = (ebit - base) * t if capped else t * (ebit - interest)
    r67 = t * base / interest if capped else t
    r68 = t if interest < ebit else t * ebit / interest
    return tax, min(r67, r68)


def optimal(raw: dict, tables: dict | None = None) -> dict:
    inp = with_country(raw)
    tables = inp.get("rating_tables") or tables or json.loads(DATA.read_text(encoding="utf-8"))["tables"]
    t, rf, erp, cs = inp["marginal_tax_rate"], inp["riskfree_rate"], inp["erp"], inp["country_default_spread"]
    ft, ibc, limit = inp.get("firm_type", 1), inp.get("indirect_bankruptcy_costs", False), inp.get("interest_limit")
    table = worksheet_table(tables, ft, rf, cs)
    kd_in, interest_in, book = inp["pre_tax_cost_of_debt"], inp["interest_expense"], inp["book_debt"]

    # current market values (rows 4-8)
    e = inp["shares_outstanding"] * inp["price"]
    if "market_value_debt" in inp:
        d = inp["market_value_debt"]
    elif inp.get("debt_maturity"):
        n = inp["debt_maturity"]
        d = interest_in * (1 - (1 + kd_in) ** -n) / kd_in + book / (1 + kd_in) ** n
    else:
        d = book
    ebitda, dep, interest, lease = inp["ebitda"], inp["depreciation"], interest_in, 0.0
    if inp.get("leases"):
        kd_lease = kd_in
        for _ in range(ITERATIONS if inp.get("synthetic_current_rating") else 1):
            le = lease_debt(inp["leases"], kd_lease, ebitda - dep, interest_in, dep)
            if not inp.get("synthetic_current_rating"):
                break
            new = sheet_rate(tables, ft, le["ebit"] / le["interest"] if le["interest"] > 0 else 1e7, rf, cs)
            if abs(new - kd_lease) < 1e-15:
                break
            kd_lease = new
        lease, ebitda, dep, interest = le["debt"], le["ebit"] + le["depreciation"], le["depreciation"], le["interest"]
    debt = d + lease
    capex = inp["capex"]
    dc_now = debt / (debt + e)
    ev = e + debt - inp["cash"]
    cov_now = (ebitda - dep) / interest if interest > 0 else 1e7
    now_row = lookup(cov_now, table)
    ke_now = rf + inp["beta"] * erp
    kd_now = now_row[2] if inp.get("synthetic_current_rating") else kd_in
    wacc_now = ke_now * (1 - dc_now) + kd_now * (1 - t) * dc_now
    fcff = (ebitda - dep) * (1 - t) + dep - capex
    g_implied = (ev * wacc_now - fcff) / (ev + fcff) if fcff > 0 else None
    g = rf if g_implied is None or g_implied > rf else g_implied
    ebitda_adj = ebitda if debt == 0 else ebitda / (1 + now_row[3])  # F45: undo the drop at today's rating
    bu = inp["beta"] / (1 + (1 - t) * debt / e)

    rows = []
    for r in RATIOS:
        dollars = r * (debt + e)
        def state(row):  # rows 52-63 for one candidate rating
            eb = ebitda_adj * (1 + row[3]) if ibc else ebitda_adj
            if inp.get("refinance"):
                intr = row[2] * dollars
            elif r < dc_now:
                intr = interest_in / book * dollars
            else:
                intr = interest_in + row[2] * (dollars - d)
            return eb, eb - dep, intr, ((eb - dep - intr) / intr + 1 if r else None)

        # rate -> interest -> coverage -> rating -> rate is circular in the source sheet (Excel iteration). A rating is a
        # solution if the coverage it implies maps back to it; there can be several, and which one Excel shows
        # depends on its iteration history. Default: the worst (what Excel reaches iterating down from D, and what
        # the template shows); ratings_at {"0.4": "Caa/CCC"} picks another solution. 0% debt: AAA, no drop (B52:B66).
        aaa = [0, table[-1][1], table[-1][2], 0.0]
        fixed = [row for row in table if lookup(state(row)[3], table) is row] if r else [aaa]
        pin = inp.get("ratings_at", {}).get(f"{r:g}")
        row = fixed[0] if pin is None else next((x for x in fixed if x[1] == pin), None)
        if row is None:
            raise ValueError(f"{pin} is not a self-consistent rating at {r:.0%} debt: {[x[1] for x in fixed]}")
        _, rating, kd, drop = row
        eb, ebit, intr, cov = state(row)
        tax, tx = tax_rates(intr, ebit, eb, t, limit)
        beta = bu * (1 + (1 - (t if not r else tx)) * r / (1 - r))
        ke = rf + beta * erp
        wacc = ke * (1 - r) + kd * (1 - tx) * r
        if ibc and g_implied is not None:
            value = (ebit * (1 - t) - (capex - dep)) * (1 + g_implied) / (wacc - g_implied)
        else:
            value = ev * (1 + (wacc_now - wacc) / (wacc - g))
        rows.append({"debt_ratio": r, "debt": dollars, "beta": beta, "cost_of_equity": ke, "ebitda": eb, "ebit": ebit,
                     "interest": intr, "coverage": cov, "rating": rating, "pre_tax_cost_of_debt": kd, "tax_rate": tx,
                     "after_tax_cost_of_debt": kd * (1 - tx), "wacc": wacc, "firm_value": value,
                     "net_income": ebit - intr - tax})
        if len(fixed) > 1:
            rows[-1]["self_consistent_ratings"] = [x[1] for x in fixed]
    best = max(rows, key=lambda x: x["firm_value"])
    price_now = (ev + inp["cash"] - d - lease) / inp["shares_outstanding"]
    out = {"company": inp.get("company"), "current": {
        "debt": debt, "market_value_debt": d, "lease_debt": lease, "equity": e, "debt_ratio": dc_now,
        "enterprise_value": ev, "ebitda": ebitda, "depreciation": dep, "interest": interest, "coverage": cov_now,
        "synthetic_rating": now_row[1], "synthetic_rate": now_row[2], "rating": inp.get("current_rating"),
        "beta": inp["beta"], "cost_of_equity": ke_now, "after_tax_cost_of_debt": kd_now * (1 - t), "wacc": wacc_now,
        "fcff": fcff, "implied_growth": g_implied, "growth_used": g, "adjusted_ebitda": ebitda_adj,
        "unlevered_beta": bu, "value_per_share": price_now, "debt_to_ebitda": debt / ebitda},
        "table": rows, "optimal": {
        "debt_ratio": best["debt_ratio"], "debt": best["debt"], "beta": best["beta"],
        "cost_of_equity": best["cost_of_equity"], "after_tax_cost_of_debt": best["after_tax_cost_of_debt"],
        "wacc": best["wacc"], "rating": best["rating"], "coverage": best["coverage"],
        "enterprise_value": best["firm_value"], "change_in_value": best["firm_value"] - ev,
        "value_per_share": (best["firm_value"] - ev) / inp["shares_outstanding"] + price_now,
        "debt_to_ebitda": best["debt"] / inp["ebitda"]}}
    multi = {f"{x['debt_ratio']:g}": x["self_consistent_ratings"][-1] for x in rows if "self_consistent_ratings" in x}
    if multi and "ratings_at" not in inp:  # same model, best self-consistent rating wherever there are several
        alt = optimal({**raw, "ratings_at": multi}, tables)["optimal"]
        out["optimal_if_best_ratings"] = {k: alt[k] for k in ("debt_ratio", "wacc", "enterprise_value", "value_per_share")}
    if inp.get("buyback_price"):  # 'Repurchase price Worksheet'
        bought = (best["debt"] - debt) / inp["buyback_price"]
        left = inp["shares_outstanding"] - bought
        out["recap"] = {"new_debt": best["debt"] - debt, "shares_bought": bought, "shares_after": left,
                        "value_per_share": (best["firm_value"] + inp["cash"] - best["debt"]) / left}
    return out


def relever(beta: float, tax: float, de_from: float, de_to: float) -> float:
    """levbeta.xls / capstru row 50: unlever at one D/E, relever at another (Hamada)."""
    return beta / (1 + (1 - tax) * de_from) * (1 + (1 - tax) * de_to)


def buyback(b: dict) -> dict:
    """buybacks.xls 'Value Effect' + 'EPS Effect' (buybackcalculator.xlsx when post_buyback_price is given)."""
    p, n, k, px = b["price"], b["shares_outstanding"], b["shares_bought"], b["buyback_price"]
    cost, cash_used = k * px, b["cash_used"]
    new_debt, t = cost - cash_used, b["tax_rate"]
    out = {"cost": cost, "new_debt": new_debt, "warning": "cash used exceeds cash on hand" if cash_used > b["cash"] else None}
    if "post_buyback_price" in b:
        post = b["post_buyback_price"]
    else:
        debt, rf, erp = b["debt"], b["riskfree_rate"], b["erp"]
        de0, de1 = debt / (p * n), (debt + new_debt) / (p * n - cost)
        beta1 = relever(b["beta"], t, de0, de1)
        dc0, dc1 = de0 / (1 + de0), de1 / (1 + de1)
        kd1 = b["cost_of_debt"] if new_debt == 0 else b["new_debt_rate"]
        w0 = (rf + b["beta"] * erp) * (1 - dc0) + b["cost_of_debt"] * (1 - t) * dc0
        w1 = (rf + beta1 * erp) * (1 - dc1) + kd1 * (1 - t) * dc1
        ev0 = p * n + debt - b["cash"]
        ev1 = ev0 + ev0 * (w0 - w1) / (w1 - rf)
        equity1 = ev1 + b["cash"] - cash_used - (debt + new_debt)
        post = equity1 / (n - k)
        fair = b.get("fair_value", p)
        out.update({"beta": [b["beta"], beta1], "debt_to_capital": [dc0, dc1], "wacc": [w0, w1],
                    "enterprise_value": [ev0, ev1], "equity_value": [p * n, equity1],
                    "value_transfer_per_share": (fair - px) * k / (n - k),
                    "price_with_transfer": post + (fair - px) * k / (n - k)})
    ni0 = b["net_income"]
    ni1 = ni0 - b.get("interest_income_from_cash", b["cash"] * 0.005) - new_debt * b["new_debt_rate"] * (1 - t)
    eps0, eps1 = ni0 / n, ni1 / (n - k)
    out.update({"price": [p, post], "market_cap": [p * n, post * (n - k)], "net_income": [ni0, ni1],
                "eps": [eps0, eps1], "pe": [p / eps0, post / eps1]})
    return out


def cash_return(c: dict) -> dict:
    """Damodaran's divfcfe dataset definitions: FCFE = NI + D&A - cap ex (incl. R&D, acquisitions) [- chg WC] (+ net debt);
    ratios are 'NA' when the denominator is negative, except net cash returned/NI."""
    ni = c["net_income"]
    pre = c.get("fcfe_pre_debt")
    if pre is None:
        pre = ni + c["depreciation"] - c["capex"] - c.get("change_in_working_capital", 0.0)
    post = c.get("fcfe_post_debt")
    if post is None:
        post = pre + c.get("new_debt", 0.0) - c.get("debt_repaid", 0.0)
    gross = c["dividends"] + c["buybacks"]
    net = gross - c.get("stock_issuances", 0.0)
    ratio = lambda a, b: a / b if b > 0 else "NA"
    return {"fcfe_pre_debt": pre, "fcfe_post_debt": post, "cash_returned": gross, "net_cash_returned": net,
            "cash_return_to_net_income": ratio(gross, ni), "net_cash_returned_to_fcfe_pre_debt": ratio(net, pre),
            "net_cash_returned_to_fcfe_post_debt": ratio(net, post), "net_cash_returned_to_net_income": net / ni,
            "unreturned_fcfe_post_debt": post - net}



def levbeta(i: dict) -> dict:
    """levbeta.xls: unlever the regression beta at the average D/E, relever at 0%..90% debt-to-capital."""
    t = i["tax"]
    bu = relever(i["beta"], t, i["de_regression"], 0.0)
    out = {"unlevered_beta": bu, "betas": [relever(bu, t, 0.0, r / (1 - r)) for r in RATIOS]}
    if "market_value_equity" in i:
        n, k, bd = i["debt_maturity"], i["cost_of_debt"], i["book_debt"]
        mvd = i["interest_expense"] * (1 - (1 + k) ** -n) / k + bd / (1 + k) ** n
        out.update({"market_value_debt": mvd,
                    "current_beta": relever(bu, t, 0.0, (mvd + i.get("lease_debt", 0.0)) / i["market_value_equity"])})
    if "de_target" in i:
        out["target_beta"] = relever(bu, t, 0.0, i["de_target"])
    return out


def synthetic_rating(i: dict) -> dict:
    """ratings.xls / capstru 'Default Spreads and Ratios' D7:D10: coverage -> rating -> rf + spread (+ country)."""
    tables = i.get("rating_tables") or json.loads(DATA.read_text(encoding="utf-8"))["tables"]
    cov = i["ebit"] / i["interest_expense"] if i["interest_expense"] > 0 else 100000
    hit = lookup(cov, tables[i.get("table", "large")])
    return {"coverage": cov, "rating": hit[2], "spread": hit[3],
            "cost_of_debt": i["riskfree_rate"] + hit[3] + i.get("country_default_spread", 0.0)}


KINDS = {"optimal": optimal, "buyback": buyback, "cash_return": cash_return, "levbeta": levbeta,
         "rating": synthetic_rating}


# ---------------------------------------------------------------- report / selftest
def pct(x):
    return "  n/a " if x is None else f"{x:6.2%}"


def report_optimal(o: dict) -> str:
    c, b = o["current"], o["optimal"]
    rows = [f"{o.get('company') or ''}  current D/(D+E) {c['debt_ratio']:.2%}  WACC {c['wacc']:.2%}  "
            f"coverage {c['coverage']:.2f} (synthetic {c['synthetic_rating']})  growth used {c['growth_used']:.2%}",
            "D/(D+E)   beta  cost eq  cover  rating     kd pre  tax    kd aft   WACC      firm value"]
    for r in o["table"]:
        cov = "     -" if r["coverage"] is None else f"{r['coverage']:6.2f}"
        rows.append(f"  {r['debt_ratio']:4.0%}  {r['beta']:6.3f} {pct(r['cost_of_equity'])} {cov}  {r['rating']:9s} "
                    f"{pct(r['pre_tax_cost_of_debt'])} {r['tax_rate']:5.1%} {pct(r['after_tax_cost_of_debt'])} "
                    f"{pct(r['wacc'])} {r['firm_value']:14,.0f}{'  <- optimal' if r['debt_ratio'] == b['debt_ratio'] else ''}"
                    + (f"  [self-consistent: {', '.join(r['self_consistent_ratings'])}]" if "self_consistent_ratings" in r else ""))
    rows += ["", f"OPTIMAL {b['debt_ratio']:.0%}: debt {b['debt']:,.0f} vs {c['debt']:,.0f} now; WACC {b['wacc']:.2%} "
             f"vs {c['wacc']:.2%}; EV {b['enterprise_value']:,.0f} vs {c['enterprise_value']:,.0f} "
             f"({b['change_in_value']:+,.0f}); value/share {b['value_per_share']:,.2f} vs {c['value_per_share']:,.2f}"]
    if "optimal_if_best_ratings" in o:
        a = o["optimal_if_best_ratings"]
        rows.append(f"Taking the best self-consistent rating where several exist: optimal {a['debt_ratio']:.0%}, "
                    f"WACC {a['wacc']:.2%}, EV {a['enterprise_value']:,.0f}, value/share {a['value_per_share']:,.2f}")
    if "recap" in o:
        r = o["recap"]
        rows.append(f"RECAP (Damodaran's repurchase sheet): new debt {r['new_debt']:+,.0f} -> shares bought "
                    f"{r['shares_bought']:,.1f}; value per remaining share {r['value_per_share']:,.2f}")
    return "\n".join(rows)


def report(res: dict) -> str:
    out = [report_optimal(res["optimal"])] if "optimal" in res else []
    for k in [k for k in res if k != "optimal"]:
        out.append(k.upper() + "\n" + "\n".join(f"  {a:36s} {v}" for a, v in res[k].items() if v is not None))
    return "\n\n".join(out)


def flatten(o, prefix=""):
    if isinstance(o, dict):
        for k, v in o.items():
            yield from flatten(v, f"{prefix}{k}.")
    elif isinstance(o, list):
        for i, v in enumerate(o):
            yield from flatten(v, f"{prefix}{i}.")
    else:
        yield prefix[:-1], o


def compare(expected: dict, got: dict) -> list:
    have_all, bad = dict(flatten(got)), []
    for k, want in flatten(expected):
        have = have_all.get(k, "missing")
        num = all(isinstance(x, (int, float)) and not isinstance(x, bool) for x in (want, have))
        if not (abs(have - want) <= max(1e-6, 5e-4 * abs(want)) if num else have == want):
            bad.append(f"{k}: Damodaran {want} ours {have}")
    return bad


def selftest() -> int:
    cases, nbad = json.loads(CASES.read_text(encoding="utf-8")), 0
    for c in cases:
        bad = compare(c["expected"], KINDS[c["kind"]](c["inputs"]))
        nbad += bool(bad)
        cells = sum(1 for _ in flatten(c["expected"]))
        pins = c["inputs"].get("ratings_at")
        print(f"{'ok ' if not bad else 'BAD'} {c['name']:46s} {cells:4d} cells"
              + (f"  (the source Excel state on another self-consistent rating: {pins})" if pins else "")
              + "".join("\n      " + b for b in bad[:6]))
    print(f"{len(cases) - nbad}/{len(cases)} cases match Damodaran's spreadsheets (0.05% relative or 1e-6 absolute, every cell)")
    return 1 if nbad else 0


# ---------------------------------------------------------------- case builder (maintenance only)
def _capstru_case(get, name: str, source: str) -> dict:
    """get('Sheet!A1') -> cached/recalculated value. Template layout ('Optimal CS Worksheet') or the older Adani
    layout ('Optimal Capital Structure', rows from 42 down shifted up by one)."""
    try:
        tmpl = get("Optimal CS Worksheet!A47") == "D/(D+E)"
    except Exception:  # sheet missing -> the older layout
        tmpl = False
    sh, off = ("Optimal CS Worksheet", 0) if tmpl else ("Optimal Capital Structure", -1)
    I = lambda c: get(f"Inputs!{c}")
    O = lambda c, n: get(f"{sh}!{c}{n + off if n >= 42 else n}")
    yes = lambda c: str(I(c)).strip().lower() == "yes"
    inp = {"company": I("B7"), "country": I("B8"), "ebitda": I("B11"), "depreciation": I("B12"), "capex": I("B13"),
           "interest_expense": I("B14"), "marginal_tax_rate": I("B15"), "current_rating": I("B16"),
           "pre_tax_cost_of_debt": I("B17"), "shares_outstanding": I("B19"), "price": I("B20"), "beta": I("B21"),
           "cash": I("B22"), "book_debt": I("B23"), "riskfree_rate": I("B37"), "erp": I("B38"),
           "country_default_spread": I("B39"), "firm_type": int(I("B42")), "refinance": yes("B43"),
           "synthetic_current_rating": yes("B44"), "indirect_bankruptcy_costs": yes("B34")}
    if yes("B24"):
        inp["market_value_debt"] = I("B25")
    elif yes("B26"):
        inp["debt_maturity"] = I("B27")
    if yes("B28"):
        L = lambda c: get(f"Operating leases!{c}")
        inp["leases"] = {"current_expense": L("E7"), "commitments": [L(f"B{r}") for r in range(10, 15)], "beyond": L("B15")}
    if yes("B30"):
        inp["interest_limit"] = {"measure": I("B31"), "max_pct": I("B32")}
    inp["buyback_price"] = get("Repurchase price Worksheet!B4")
    S = lambda r: [get(f"Default Spreads and Ratios!{c}{r}") for c in "ABCDE"]
    inp["rating_tables"] = {"large": [S(r) for r in range(15, 30)], "small": [S(r) for r in range(34, 49)],
                            "utility": [S(r) for r in range(52, 67)]}
    num = lambda x: None if x in ("NA",) or isinstance(x, str) and x.strip() in ("", "—", "∞", "��") else x
    table = []
    for col in "BCDEFGHIJK":
        f = {k: O(col, n) for k, n in [("debt_ratio", 47), ("debt", 49), ("beta", 50), ("cost_of_equity", 51),
                                        ("ebitda", 53), ("ebit", 55), ("interest", 56), ("net_income", 59),
                                        ("coverage", 63), ("rating", 65), ("pre_tax_cost_of_debt", 66),
                                        ("tax_rate", 69), ("after_tax_cost_of_debt", 75), ("wacc", 76),
                                        ("firm_value", 78)]}
        f["coverage"] = num(f["coverage"]) if col != "B" else None
        table.append(f)
    cur = {k: num(O(c, n)) for k, (c, n) in {
        "market_value_debt": ("C", 5), "lease_debt": ("C", 7), "ebitda": ("I", 4), "depreciation": ("I", 5),
        "interest": ("I", 8), "debt": ("E", 13), "debt_ratio": ("E", 12), "enterprise_value": ("J", 14),
        "coverage": ("K", 25), "synthetic_rating": ("K", 26), "synthetic_rate": ("K", 27), "cost_of_equity": ("E", 15),
        "after_tax_cost_of_debt": ("E", 17), "wacc": ("E", 19), "fcff": ("J", 16), "implied_growth": ("J", 17),
        "growth_used": ("E", 20), "adjusted_ebitda": ("F", 45), "unlevered_beta": ("M", 6),
        "value_per_share": ("E", 22)}.items()}
    best = {k: O(c, n) for k, (c, n) in {"debt_ratio": ("F", 12), "debt": ("F", 13), "beta": ("F", 14),
            "cost_of_equity": ("F", 15), "after_tax_cost_of_debt": ("F", 17), "wacc": ("F", 19),
            "enterprise_value": ("F", 21), "value_per_share": ("F", 22)}.items()}
    if tmpl:  # the template's Output Summary (the Adani files use older display formulas here)
        cur["debt_to_ebitda"] = I("G13")
        best.update({"debt_to_ebitda": I("H13"), "coverage": num(I("H14"))})
    got = optimal(inp)["table"]  # where the source Excel iteration sits on another self-consistent rating, pin it
    pins = {f"{e['debt_ratio']:g}": e["rating"] for e, g in zip(table, got)
            if e["rating"] != g["rating"] and e["rating"] in g.get("self_consistent_ratings", [])}
    if pins:
        inp["ratings_at"] = pins
    R = lambda c: get(f"Repurchase price Worksheet!{c}")
    recap = {"new_debt": R("B8"), "shares_bought": R("B9"), "shares_after": R("B10"), "value_per_share": R("B17")}
    return {"name": name, "kind": "optimal", "source": source, "inputs": inp,
            "expected": {"current": cur, "table": table, "optimal": best, "recap": recap}}


def build(files: list) -> int:
    """Rebuild data/capital_structure_cases.json (and the rating tables in data/) from Damodaran's workbooks."""
    import warnings
    warnings.filterwarnings("ignore")
    import openpyxl
    import xlrd

    cases, url = [], "https://pages.stern.nyu.edu/~adamodar/pc/"
    for f in map(Path, files):
        low = f.name.lower()
        if low.endswith(".xlsx") and "capstru" in low:
            wb = openpyxl.load_workbook(f, data_only=True)
            get = lambda a, wb=wb: wb[a.rsplit("!", 1)[0]][a.rsplit("!", 1)[1]].value
            note = get("Inputs!D7")  # template copies recalculated in Excel under an override carry "branch: ..."
            if isinstance(note, str) and note.startswith("branch:"):
                what, cells = (x.strip() for x in note[7:].split("|"))
                name, src = f"capstru template branch: {what}", f"{url}capstru.xlsx recalculated in Excel, Inputs {cells}"
            else:
                blog = "blog/" if low.startswith("adani") else ""
                name, src = f"{f.stem} ({get('Inputs!B7')})", f"{url}{blog}{f.name}"
            case = _capstru_case(get, name, src)
            cases.append(case)
            if low == "capstru.xlsx":  # the current template: the rating tables for data/
                DATA.write_text(json.dumps({"source": "Damodaran capstru.xlsx, 'Default Spreads and Ratios' A15:E66 "
                                            "(rows: coverage low, high, rating, spread, EBITDA drop)",
                                            "url": url + "capstru.xlsx", "as_of": "January 2026 (template's tables)",
                                            "tables": case["inputs"]["rating_tables"]}, indent=1), encoding="utf-8")
        elif low == "buybacks.xls":
            v = xlrd.open_workbook(f).sheet_by_name("Value Effect")
            e = xlrd.open_workbook(f).sheet_by_name("EPS Effect")
            c = lambda s, a: s.cell_value(*_rc(a))
            V, E = (lambda a: c(v, a)), (lambda a: c(e, a))
            inp = {"price": V("B6"), "shares_outstanding": V("B7"), "beta": V("B8"), "debt": V("B10"), "cash": V("B11"),
                   "cost_of_debt": V("B12"), "net_income": V("B13"), "interest_income_from_cash": V("B14"),
                   "fair_value": V("B17"), "buyback_price": V("B19"), "shares_bought": V("B20"), "cash_used": V("B21"),
                   "new_debt_rate": V("B23"), "riskfree_rate": V("B25"), "erp": V("B26"), "tax_rate": V("B27")}
            two = lambda g, r: [g(f"B{r}"), g(f"C{r}")]
            exp = {"new_debt": V("B22"), "beta": two(V, 31), "debt_to_capital": two(V, 33), "wacc": two(V, 36),
                   "enterprise_value": two(V, 37), "equity_value": two(V, 40), "price": two(V, 42),
                   "value_transfer_per_share": V("B50"), "price_with_transfer": V("B51"), "market_cap": two(E, 21),
                   "net_income": two(E, 23), "eps": two(E, 24), "pe": two(E, 25)}
            cases.append({"name": "buybacks.xls (value + EPS effect)", "kind": "buyback", "source": url + f.name,
                          "inputs": inp, "expected": exp})
        elif low == "buybackcalculator.xlsx":
            s = openpyxl.load_workbook(f, data_only=True)["Sheet1"]
            S = lambda a: s[a].value
            inp = {"price": S("B2"), "cash": S("B4"), "net_income": S("B6"), "shares_outstanding": S("B7"),
                   "shares_bought": S("B10"), "buyback_price": S("B11"), "post_buyback_price": S("B12"),
                   "cash_used": S("B15"), "new_debt_rate": S("B17"), "tax_rate": S("B18")}  # B5 left to our default
            two = lambda r: [S(f"B{r}"), S(f"C{r}")]
            cases.append({"name": "buybackcalculator.xlsx (EPS effect)", "kind": "buyback", "source": url + f.name,
                          "inputs": inp, "expected": {"new_debt": S("B16"), "market_cap": two(20), "price": two(21),
                                                      "net_income": two(22), "eps": two(23), "pe": two(24)}})
        elif low == "levbeta.xls":
            s = xlrd.open_workbook(f).sheet_by_index(0)
            S = lambda a: s.cell_value(*_rc(a))
            inp = {"beta": S("E3"), "tax": S("E4"), "de_regression": S("E5"), "market_value_equity": S("E7"),
                   "book_debt": S("E8"), "debt_maturity": S("E9"), "interest_expense": S("E10"), "cost_of_debt": S("E11"),
                   "de_target": S("E33")}
            cases.append({"name": "levbeta.xls (Hamada table)", "kind": "levbeta", "source": url + f.name, "inputs": inp,
                          "expected": {"unlevered_beta": S("E17"), "betas": [S(f"C{r}") for r in range(21, 31)],
                                       "market_value_debt": S("E12"), "current_beta": S("E18"), "target_beta": S("E35")}})
        elif low == "ratings.xls":
            s = xlrd.open_workbook(f).sheet_by_index(0)
            S = lambda a: s.cell_value(*_rc(a))
            tab = lambda r0, c0: [[S(f"{c}{r}") for c in c0] + [0] for r in range(r0, r0 + 15)]
            inp = {"ebit": S("F6"), "interest_expense": S("F7"), "riskfree_rate": S("F8"),
                   "table": {1: "large", 2: "small", 3: "financial"}[int(S("C4"))],
                   "rating_tables": {"large": tab(19, "ABCD"), "small": tab(38, "ABCD"), "financial": tab(19, "FGHI")}}
            cases.append({"name": "ratings.xls (synthetic rating)", "kind": "rating", "source": url + f.name, "inputs": inp,
                          "expected": {"coverage": S("D10"), "rating": S("D11"), "spread": S("D12"), "cost_of_debt": S("D13")}})
        elif low.startswith("divfcfe"):
            s = xlrd.open_workbook(f).sheet_by_name("Industry Averages")
            pick = {"Total Market", "Grand Total", "Total Market (without financials)"}
            if low == "divfcfe.xls":  # plus rows that exercise the dataset's 'NA' rules
                pick |= {"Advertising", "Air Transport", "Chemical (Basic)", "Chemical (Specialty)"}
            for r in range(s.nrows):
                row = s.row_values(r)
                if row[0] not in pick:
                    continue
                C, D, F, H, I, J = row[2], row[3], row[5], row[7], row[8], row[9]
                cases.append({"name": f"{f.stem}: {row[0]}", "kind": "cash_return",
                              "source": f"{url}datasets/{f.name}", "inputs": {
                                  "net_income": D, "dividends": C, "buybacks": F - C, "stock_issuances": F - H,
                                  "fcfe_pre_debt": I, "fcfe_post_debt": J},
                              "expected": {"cash_return_to_net_income": row[6], "net_cash_returned_to_fcfe_pre_debt": row[10],
                                           "net_cash_returned_to_fcfe_post_debt": row[11],
                                           "net_cash_returned_to_net_income": row[12]}})
        else:
            print("skipped (not a source for this engine):", f)
    CASES.write_text(json.dumps(cases, indent=1, default=str), encoding="utf-8")
    print(f"wrote {len(cases)} cases -> {CASES}")
    return selftest()


def _rc(a: str) -> tuple[int, int]:
    import re
    col, row = re.match(r"([A-Z]+)(\d+)", a).groups()
    c = 0
    for ch in col:
        c = c * 26 + ord(ch) - 64
    return int(row) - 1, c - 1


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="?")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--from-case", help="print the source inputs for a calibration case (e.g. AdaniCapStru) as JSON")
    ap.add_argument("--build-cases", nargs="+", metavar="FILE")
    a = ap.parse_args()
    if a.selftest:
        return selftest()
    if a.build_cases:
        return build(a.build_cases)
    if a.from_case:
        cases = json.loads(CASES.read_text(encoding="utf-8"))
        hit = [c for c in cases if c["name"].lower().startswith(a.from_case.lower())]
        if not hit:
            print("cases:", ", ".join(c["name"] for c in cases))
            return 1
        c = hit[0]
        pins = c["inputs"].get("ratings_at")
        note = f" (the cached source table sits on other self-consistent ratings {pins}; add ratings_at to reproduce)" if pins else ""
        inp = {k: v for k, v in c["inputs"].items() if k != "ratings_at"}
        print(json.dumps({"_source": f"Damodaran's own inputs: {c['source']}{note}", c["kind"]: inp}, indent=1))
        return 0
    inp = json.loads(Path(a.inputs).read_text(encoding="utf-8"))
    res = {k: KINDS[k]({x: y for x, y in v.items() if not x.startswith("_")}) for k, v in inp.items() if k in KINDS}
    print(json.dumps(res, indent=1) if a.json else report(res))
    return 0


if __name__ == "__main__":
    sys.exit(main())
