"""Damodaran's equity valuation models for banks, insurers and other financial-service firms.

The FCFF engine (ginzu.py) cannot value a bank: debt is raw material, not capital, so Damodaran values only the
equity - with cash flows to equity, a cost of equity, and (for banks) reinvestment defined as the regulatory
capital needed to grow. Every formula below is transcribed from Damodaran's own spreadsheets (cell refs in comments)
and replayed against their cached outputs by --selftest.

Models ("model" key in the inputs JSON) and the source files:
  bank_fcfe      CitiVal2023.xlsx (default template) and DeutscheBank2016.xls (template "DeutscheBank2016"):
                 FCFE = net income - change in Tier 1 capital (risk-adjusted assets x Tier 1 ratio),
                 ROE and capital-ratio paths to targets, terminal payout g/ROE, failure/wipeout probability.
                 https://pages.stern.nyu.edu/~adamodar/pc/blog/CitiVal2023.xlsx (post 2023-05),
                 http://www.stern.nyu.edu/~adamodar/pc/blog/DeutscheBank2016.xls (post 2016-10)
  excess_return  eqexret.xls: equity = book equity + PV of (ROE - cost of equity) x beginning book equity.
  ddm            divginzu.xls (Damodaran's dividend 'Ginzu', shipped with the JPMorgan May-2024 inputs).
  fcfe           fcfeginzu.xls (Damodaran's FCFE 'Ginzu', shipped with Levi Strauss May-2024 inputs): FCFE =
                 net income x (1 - equity reinvestment rate), growth = non-cash ROE x reinvestment rate.
  ddm_stable, ddm_2stage, ddm_3stage       ddmst.xls, ddm2st.xls, ddm3st.xls
  fcfe_stable, fcfe_2stage, fcfe_3stage    fcfest.xls, fcfe2st.xls, fcfe3st.xls
  fcfe_vs_ddm    fcfevsddm.xls (reconciles FCFE and DDM values through the cash build-up)
  Tool URLs: https://pages.stern.nyu.edu/~adamodar/pc/<file>.

Usage:
  python financial_firms.py inputs.json [--json]     value + year table (inputs: see model docstrings)
  python financial_firms.py --from-case CitiVal2023   the source inputs for a case, as a starting JSON
  python financial_firms.py --selftest                replay data/financial_firms_cases.json
  python financial_firms.py --build-cases FILES... [--xlsx-dir DIR]
      read the source files (xlrd/openpyxl) + recalc option branches with pycel. pycel reads only .xlsx, so .xls
      templates need Excel-converted copies (same stem) in DIR.
Rates are decimals. "erp" may be a number or a cost_of_capital.py erp spec, e.g. {"approach": "country",
"country": "Germany"} (Damodaran's Jan-2026 data). Keys ending in _high/_stable/..., when omitted, take the template's
default (usually: same as the current/high-growth value).
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from statistics import mean
from typing import Any, Callable

HERE = Path(__file__).resolve().parent
CASES = HERE.parent / "data" / "financial_firms_cases.json"


# ---------------------------------------------------------------- shared pieces of the templates
def _ke(inp: dict) -> float:
    """High-growth cost of equity: direct input, else CAPM riskfree + beta x ERP."""
    return inp["cost_of_equity"] if "cost_of_equity" in inp else inp["riskfree_rate"] + inp["beta"] * inp["erp"]


def _ke_stable(inp: dict) -> float:
    """Stable cost of equity in the templates: always CAPM, with beta_stable if given (even when the
    high-growth cost of equity was entered directly)."""
    return inp["riskfree_rate"] + inp.get("beta_stable", inp["beta"]) * inp["erp"]


def _glide(hi: float, st: float, t: int, n: int, gradual: bool) -> float:
    """The templates' 'gradually adjust inputs in the second half' rule (eqexret/divginzu/fcfeginzu): year t < n/2 keeps
    the high-growth value, later years move linearly to the stable value, reached in year n."""
    return hi if not gradual or t < n / 2 else st + (hi - st) / (n / 2) * (n - t)


def _weighted_growth(inp: dict, fundamental: float) -> float:
    """Stage templates: EPS growth = weights x [historical 5-yr EPS CAGR, outside estimate, fundamental]."""
    wh, wo, wf = inp["growth_weights"]
    hist = (inp["eps"] / inp["eps_5y_ago"]) ** 0.2 - 1 if "eps_5y_ago" in inp else 0.0
    return wh * hist + wo * inp.get("outside_growth", 0.0) + wf * fundamental


def _roc_growth(inp: dict) -> tuple[float, float]:
    """fcfe2st/fcfe3st fundamental growth b x (ROC + D/E (ROC - i(1-t))) and current D/E."""
    de = inp["book_debt"] / inp["book_equity"] if inp.get("book_equity") else 0.0
    if "net_income" not in inp:
        return 0.0, de
    t = inp["tax_rate"]
    roc = (inp["net_income"] + inp["interest_expense"] * (1 - t)) / (
        inp["book_debt_last_year"] + inp["book_equity_last_year"])
    rate = inp["interest_expense"] / inp["book_debt"] if inp["book_debt"] else 0.0
    roc, de_h = inp.get("roc_high", roc), inp.get("de_high", de)
    b, i = inp.get("retention_high", 1 - inp["dps"] / inp["eps"]), inp.get("interest_rate_high", rate)
    return b * (roc + de_h * (roc - i * (1 - t))), de


def _check(n: int, lo: int, hi: int, why: str) -> int:
    if not lo <= n <= hi:
        raise ValueError(f"years={n}: the template supports {lo}-{hi} ({why})")
    return int(n)


# ---------------------------------------------------------------- bank FCFE (CitiVal2023 / DeutscheBank2016)
def bank_fcfe(inp: dict) -> dict:
    """Inputs: tier1_capital, book_equity, risk_adjusted_assets, net_income, asset_growth, target_tier1_ratio,
    target_roe, stable_growth, shares_outstanding, price, cost_of_equity, stable_cost_of_equity (= cost_of_equity),
    failure_probability (0). CitiVal2023: 5 years, linear paths, current ROE = net income / Tier 1 capital (the sheet's
    B9), NI_t = ROE_t x book equity_t-1. template "DeutscheBank2016": 10 years, tier1_ratio_next_year then linear
    to target in yr 10, ROE closes 1/3 of the gap to target_roe (yr 5) each year then linear to target_roe_10
    (= stable_cost_of_equity), NI_t = ROE_t x book equity_t, cost of equity linear to stable in yrs 6-10,
    tier1_capital_adjustment (the sheet's DOJ fine) hits Tier 1 and book equity."""
    db = inp.get("template", "CitiVal2023") == "DeutscheBank2016"
    n, g, gs = (10 if db else 5), inp["asset_growth"], inp["stable_growth"]
    adj = inp.get("tier1_capital_adjustment", 0.0)
    t1, be0, ni0, raa0 = inp["tier1_capital"] + adj, inp["book_equity"] + adj, inp["net_income"], inp["risk_adjusted_assets"]
    ke = inp["cost_of_equity"]
    ks = inp.get("stable_cost_of_equity", ke)
    raa = [raa0]
    for t in range(1, n + 1):
        raa.append(raa[-1] * (1 + g))
    raa.append(raa[n] * (1 + gs))
    r0, target = t1 / raa0, inp["target_tier1_ratio"]
    if db:  # C3 = next-year ratio, D3..K3 step (T - C3)/9, L3 = target
        r1 = inp["tier1_ratio_next_year"]
        ratio = [r0, r1] + [r1 + (target - r1) / 9 * (t - 1) for t in range(2, n)] + [r0 + (target - r0) / 10 * 10]
    else:
        ratio = [r0 + (target - r0) / n * t for t in range(n + 1)]
    ratio.append(ratio[n])
    cap = [ratio[t] * raa[t] for t in range(n + 2)]
    if db:
        r5, r10 = inp["target_roe"], inp.get("target_roe_10", ks)
        roe = [ni0 / be0]
        for _ in range(4):
            roe.append(roe[-1] - (roe[-1] - r5) / 3)
        roe.append(r5)
        for _ in range(5):
            roe.append(roe[-1] - (r5 - r10) / 5)
        roe.append(r10)
        kes = [ke] * 6
        for _ in range(5):
            kes.append(kes[-1] - (ke - ks) / 5)
    else:
        roe0 = ni0 / t1  # the sheet: 'Current ROE' = net income / Tier 1 capital
        roe = [roe0 + (inp["target_roe"] - roe0) / n * t for t in range(n + 1)]
        roe.append(roe[n])
        kes = [ke] * (n + 1)
    be, ni, rows, cum, pv = [be0], [ni0], [], 1.0, 0.0
    for t in range(1, n + 1):
        dcap = cap[t] - cap[t - 1]
        income = roe[t] * (be[-1] + dcap if db else be[-1])
        be.append(be[-1] + dcap)
        ni.append(income)
        cum *= 1 + kes[t]
        pv += (income - dcap) / cum
        rows.append({"t": t, "risk_adj_assets": raa[t], "tier1_ratio": ratio[t], "tier1_capital": cap[t],
                     "roe": roe[t], "net_income": income, "reg_capital_investment": dcap, "fcfe": income - dcap,
                     "book_equity": be[t], "cost_of_equity": kes[t]})
    ni_s = ni[n] * (1 + gs) if db else roe[n + 1] * be[n]
    fcfe_s = ni_s - gs / roe[n + 1] * ni_s
    tv = fcfe_s / (ks - gs)
    rows.append({"t": "terminal", "risk_adj_assets": raa[n + 1], "tier1_ratio": ratio[n + 1], "roe": roe[n + 1],
                 "net_income": ni_s, "reg_capital_investment": ni_s - fcfe_s, "fcfe": fcfe_s, "cost_of_equity": ks})
    equity = pv + tv / cum
    vps = equity / inp["shares_outstanding"]
    adj_vps = vps * (1 - inp.get("failure_probability", 0.0))
    return {"template": "DeutscheBank2016" if db else "CitiVal2023", "current_tier1_ratio": r0, "current_roe": roe[0],
            "pv_fcfe": pv, "terminal_value": tv, "pv_terminal": tv / cum, "equity_value": equity,
            "value_per_share": vps, "value_per_share_adjusted": adj_vps, "price": inp.get("price"),
            "price_to_value": inp["price"] / adj_vps if inp.get("price") else None, "years": rows}


# ---------------------------------------------------------------- excess return model (eqexret.xls)
def excess_return(inp: dict) -> dict:
    """Inputs: net_income, book_equity, book_equity_last_year, eps, dps, shares_outstanding, beta, riskfree_rate,
    erp, years (1-10), stable_growth; optional normalize_earnings {approach 1, net_income_history | approach 2,
    roe}, roe_high, retention_high, roe_stable, gradual (true), stable_payout, beta_stable, erp_stable."""
    n = _check(inp["years"], 1, 10, "10 year columns; 0 leaves no terminal value")
    norm, bv_last = inp.get("normalize_earnings"), inp["book_equity_last_year"]
    roe0 = (inp["net_income"] / bv_last if not norm else
            mean(norm["net_income_history"]) / bv_last if norm["approach"] == 1 else norm["roe"])  # Inputs!B18
    roe, ret = inp.get("roe_high", roe0), inp.get("retention_high", 1 - inp["dps"] / inp["eps"])
    ke = inp["riskfree_rate"] + inp["beta"] * inp["erp"]
    ks = inp["riskfree_rate"] + inp["beta_stable"] * inp.get("erp_stable", inp["erp"]) if "beta_stable" in inp else ke
    gs, roe_s = inp["stable_growth"], inp.get("roe_stable", roe)
    pay_s, gradual = inp.get("stable_payout", 1 - gs / roe_s), inp.get("gradual", True)
    bv, cum, pv, rows = inp["book_equity"], 1.0, 0.0, []
    for t in range(1, n + 1):
        k, r, p = (_glide(hi, st, t, n, gradual) for hi, st in ((ke, ks), (roe, roe_s), (1 - ret, pay_s)))
        ni, cost = bv * r, bv * k
        cum *= 1 + k
        pv += (ni - cost) / cum
        rows.append({"t": t, "beginning_bv": bv, "roe": r, "net_income": ni, "equity_cost": cost,
                     "excess_return": ni - cost, "cost_of_equity": k, "payout": p, "pv": (ni - cost) / cum})
        bv += ni * (1 - p)
    bv_t = max(x["beginning_bv"] for x in rows) * (1 + gs)  # the sheet's L15 = MAX(B15:K15)*(1+g)
    tv = bv_t * (roe_s - ks) / (ks - gs)
    rows.append({"t": "terminal", "beginning_bv": bv_t, "roe": roe_s, "net_income": bv_t * roe_s,
                 "equity_cost": bv_t * ks, "excess_return": bv_t * (roe_s - ks), "cost_of_equity": ks})
    equity = inp["book_equity"] + pv + tv / cum
    return {"cost_of_equity": ke, "roe_high": roe, "expected_growth": roe * ret, "terminal_value": tv,
            "pv_excess_returns": pv + tv / cum, "equity_value": equity,
            "value_per_share": equity / inp["shares_outstanding"], "years": rows}


# ---------------------------------------------------------------- dividend Ginzu (divginzu.xls)
def ddm(inp: dict) -> dict:
    """Inputs: book_equity_last_year, net_income, dividends (aggregate), shares_outstanding, price, beta,
    riskfree_rate, erp, years (1-15); optional growth (else ROE x retention), roe_high, retention_high,
    roe_stable, gradual (true), stable_growth (default = riskfree_rate), stable_payout, beta_stable."""
    n = _check(inp["years"], 1, 15, "15 year columns; the template's 0-year branch mixes per-share and total values")
    ni, rf = inp["net_income"], inp["riskfree_rate"]
    roe = inp.get("roe_high", ni / inp["book_equity_last_year"])
    ret = inp.get("retention_high", 1 - inp["dividends"] / ni)
    g, payout = inp.get("growth", roe * ret), 1 - ret
    ke, ks = rf + inp["beta"] * inp["erp"], rf + inp.get("beta_stable", inp["beta"]) * inp["erp"]
    gs, roe_s = inp.get("stable_growth", rf), inp.get("roe_stable", roe)
    pay_s, gradual = inp.get("stable_payout", 1 - gs / roe_s), inp.get("gradual", True)
    e, cum, pv, rows = ni, 1.0, 0.0, []
    for t in range(1, n + 1):
        gt, pt, kt, rt = (_glide(hi, st, t, n, gradual) for hi, st in ((g, gs), (payout, pay_s), (ke, ks), (roe, roe_s)))
        e = e * (1 + gt) if t > 1 else e  # the sheet's C14: year-1 earnings = current net income
        cum *= 1 + kt
        pv += e * pt / cum
        rows.append({"t": t, "growth": gt, "roe": rt, "net_income": e, "payout": pt, "dividends": e * pt,
                     "cost_of_equity": kt, "pv": e * pt / cum})
    tv = max(x["net_income"] for x in rows) * (1 + gs) * pay_s / (ks - gs)
    shares = inp["shares_outstanding"]
    vps, eps = (pv + tv / cum) / shares, ni / shares
    aip = eps / ks
    stable = eps * pay_s * (1 + gs) / (ks - gs) - aip
    return {"cost_of_equity": ke, "expected_growth": g, "payout_high": payout, "stable_payout": pay_s,
            "pv_dividends": pv, "terminal_value": tv, "pv_terminal": tv / cum, "equity_value": pv + tv / cum,
            "value_per_share": vps, "price": inp.get("price"),
            "value_assets_in_place": aip, "value_stable_growth": stable, "value_extraordinary_growth": vps - stable - aip,
            "years": rows}


# ---------------------------------------------------------------- FCFE Ginzu (fcfeginzu.xls)
def _normalized_reinvestment(inp: dict, nr: dict) -> float:
    """'Normalized Reinvestment' sheet -> Inputs!B21 + B22 - B23 (approach 0 = use current numbers)."""
    c, w, d = nr.get("net_capex", {}), nr.get("working_capital", {}), nr.get("net_debt", {})
    a = c.get("approach", 0)
    capex = (inp["capex"] - inp["depreciation"] if a == 0 else
             mean(c["net_capex_history"]) / mean(c["ebit_history"]) * c["ebit_history"][-1] if a == 1 else
             c["net_capex_to_ebit"] * c["ebit"])
    a = w.get("approach", 0)
    drev = w.get("revenues", 0) - w.get("revenues_last_year", 0)
    wc = (inp["change_wc"] if a == 0 else w["non_cash_wc"] / w["revenues"] * drev if a == 1 else w["wc_to_revenue"] * drev)
    a = d.get("approach", 0)
    debt = (inp["net_debt_issued"] if a == 0 else
            d["book_debt"] / (d["book_debt"] + inp["book_equity"]) * (capex + wc) if a == 1 else
            d["book_debt"] / (d["book_debt"] + inp["market_cap"]) * (capex + wc) if a == 2 else
            d["debt_ratio"] * (capex + wc))
    return capex + wc - debt


def fcfe(inp: dict) -> dict:
    """Inputs: net_income, interest_income_after_tax (0), book_equity, book_equity_last_year, cash,
    cash_last_year, shares_outstanding, capex, depreciation, change_wc, net_debt_issued, beta, riskfree_rate,
    erp, years (0 or 2-10), stable_growth; optional rd {life, current, past} (ginzu.py format),
    normalize_earnings {approach 1 net_income_history | 2 roe}, normalize_reinvestment {net_capex, working_capital,
    net_debt} (+ market_cap), growth, roe_high, reinvestment_rate_high, roe_stable, gradual (true),
    stable_reinvestment_rate (default stable_growth / roe_stable), beta_stable, price.
    replicate_sheet_bugs (false): true reproduces two slips in the source file - with reported earnings its B38 subtracts
    interest income twice (B36 - B7, B36 already net of B7), and year-1 net income grows at the year-2 rate
    (C14 = D5*(1+D12)). Default: interest income subtracted once, year 1 at its own growth rate."""
    n, bug = inp["years"], inp.get("replicate_sheet_bugs", False)
    if (n == 1 and bug) or not 0 <= n <= 10:
        raise ValueError(f"years={n}: the fcfeginzu template needs 0-10 (1 only without replicate_sheet_bugs: its year 1 "
                         "uses year-2 growth; its columns for years 11-15 compute FCFE as net income x reinvestment "
                         "rate, a template slip)")
    ni_raw, ii = inp["net_income"], inp.get("interest_income_after_tax", 0.0)
    rd = 0.0
    if inp.get("rd"):
        from ginzu import rd_adjustment
        rd = rd_adjustment(inp["rd"])[0]
    ne = inp.get("normalize_earnings")
    if ne:  # the sheet's B20 uses current book equity, its D4 last year's
        hist = ne["approach"] == 1 and mean(ne["net_income_history"])
        adj_ni = (hist or ne["roe"] * inp["book_equity"]) + rd
        ni_base = hist or ne["roe"] * inp["book_equity_last_year"]
    else:
        adj_ni, ni_base = ni_raw - ii + rd, ni_raw
    roe0 = adj_ni / (inp["book_equity_last_year"] - inp["cash_last_year"])  # non-cash ROE, B37
    nr = inp.get("normalize_reinvestment")
    err0 = (_normalized_reinvestment(inp, nr) / (ni_raw - ii) if nr else  # B38
            (inp["capex"] - inp["depreciation"] + inp["change_wc"] - inp["net_debt_issued"]) /
            (adj_ni - ii if ne or bug else adj_ni))  # reported branch: adj_ni is already net of interest income
    roe, err = inp.get("roe_high", roe0), inp.get("reinvestment_rate_high", err0)
    g, rf = inp.get("growth", roe * err), inp["riskfree_rate"]
    ke, ks = rf + inp["beta"] * inp["erp"], rf + inp.get("beta_stable", inp["beta"]) * inp["erp"]
    gs, roe_s = inp["stable_growth"], inp.get("roe_stable", roe)
    err_s, gradual = inp.get("stable_reinvestment_rate", gs / roe_s), inp.get("gradual", True)
    ni0 = ni_base - ii
    growth = [_glide(g, gs, t, n, gradual) for t in range(1, n + 1)]
    e, cum, pv, rows = ni0, 1.0, 0.0, []
    for t in range(1, n + 1):
        e *= 1 + growth[(max(t, 2) if bug else t) - 1]  # the sheet's C14 = D5*(1+D12): year 1 at the year-2 rate
        r, kt = _glide(err, err_s, t, n, gradual), _glide(ke, ks, t, n, gradual)
        cum *= 1 + kt
        pv += e * (1 - r) / cum
        rows.append({"t": t, "growth": growth[t - 1], "net_income": e, "reinvestment_rate": r, "fcfe": e * (1 - r),
                     "cost_of_equity": kt, "pv": e * (1 - r) / cum})
    tv = (max(x["net_income"] for x in rows) if n else ni0) * (1 + gs) * (1 - err_s) / (ks - gs)
    equity = pv + tv / cum + inp["cash"]
    vps = equity / inp["shares_outstanding"]
    out = {"cost_of_equity": ke, "non_cash_roe": roe0, "reinvestment_rate": err0, "expected_growth": g,
           "stable_reinvestment_rate": err_s, "pv_fcfe": pv, "terminal_value": tv, "pv_terminal": tv / cum,
           "equity_value": equity, "value_per_share": vps, "price": inp.get("price"), "years": rows}
    his = None if bug or n == 1 else fcfe({**inp, "replicate_sheet_bugs": True})["value_per_share"]
    if his is not None and abs(his - vps) > 1e-9 * abs(vps):
        out["sheet_error"] = (f"Damodaran's fcfeginzu as written gives {his:,.4f}/share (interest income subtracted twice "
                              "in B38 and/or year 1 at the year-2 growth rate); replicate_sheet_bugs=true reproduces it")
    return out


# ---------------------------------------------------------------- textbook stage models
def ddm_stable(inp: dict) -> dict:
    """ddmst.xls Gordon growth: eps, payout, growth, cost_of_equity | beta+riskfree_rate+erp."""
    ke, g = _ke(inp), inp["growth"]
    dps = inp["eps"] * inp["payout"]
    w = [m for c, m in ((g > 0.1, "This is high for a stable growth rate"),
                        (inp.get("beta", 0) > 1.5, "This Beta is high for a stable firm"),
                        (inp["payout"] < 0.2, "Payout ratio is low for a stable firm")) if c]
    return {"cost_of_equity": ke, "dps": dps, "value_per_share": dps * (1 + g) / (ke - g), "warnings": w}


def ddm_2stage(inp: dict) -> dict:
    """ddm2st.xls: eps, dps, years, cost_of_equity | beta, riskfree_rate, erp (stable phase is always CAPM),
    growth_weights [historical, outside, fundamental] with eps_5y_ago, outside_growth, net_income +
    book_equity_last_year (ROE); optional roe_high, retention_high, roe_stable, stable_growth, stable_payout,
    beta_stable."""
    eps, dps, n = inp["eps"], inp["dps"], inp["years"]
    ke = _ke(inp)
    roe = inp.get("roe_high", inp["net_income"] / inp["book_equity_last_year"] if "net_income" in inp else None)
    ret = inp.get("retention_high", 1 - dps / eps)
    g = _weighted_growth(inp, ret * roe if roe is not None else 0.0)
    gs, roe_s = inp["stable_growth"], inp.get("roe_stable", roe)
    pay, pay_s, ks = 1 - ret, inp.get("stable_payout", 1 - gs / roe_s), _ke_stable(inp)
    price_n = eps * (1 + g) ** n * (1 + gs) * pay_s / (ks - gs)
    pv_div = eps * pay * (1 + g) * (1 - (1 + g) ** n / (1 + ke) ** n) / (ke - g)
    value = pv_div + price_n / (1 + ke) ** n
    aip = dps / ks
    stable = dps * (1 + gs) / (ks - gs) - aip
    w = [m for c, m in ((eps < 0, "negative current EPS: this model will not work"),
                        (inp.get("eps_5y_ago", 1) < 0, "historical growth cannot be computed with negative EPS"),
                        (roe is not None and roe < gs, "high-growth ROE below stable growth: very low fundamental growth"),
                        (roe_s is not None and roe_s <= gs, "stable ROE <= stable growth: you cannot afford any dividends"),
                        (abs(sum(inp["growth_weights"]) - 1) > 1e-9, "growth weights do not add up to one"),
                        (gs > 0.1, "This is a high growth rate for a stable period")) if c]
    return {"cost_of_equity": ke, "expected_growth": g, "payout_high": pay, "stable_payout": pay_s,
            "stable_cost_of_equity": ks, "terminal_price": price_n, "pv_dividends": pv_div,
            "pv_terminal": price_n / (1 + ke) ** n, "value_per_share": value, "value_assets_in_place": aip,
            "value_stable_growth": stable, "value_extraordinary_growth": value - stable - aip, "warnings": w,
            "years": [{"t": t, "dividends": eps * (1 + g) * pay * (1 + g) ** (t - 1)} for t in range(1, n + 1)]}


def ddm_3stage(inp: dict) -> dict:
    """ddm3st.xls: ddm_2stage inputs + transition_years (linear growth decline); roe_stable (or stable_payout);
    optional transition_payout / transition_beta (else gradual to stable). Beta, riskfree_rate and erp needed."""
    eps, dps = inp["eps"], inp["dps"]
    n1 = _check(inp["years"], 0, 10, "10 high-growth columns")
    n2 = _check(inp["transition_years"], 0, 10, "10 transition columns")
    ke, rf, erp = _ke(inp), inp["riskfree_rate"], inp["erp"]
    bv = inp.get("book_equity_last_year", 0)
    roe = inp.get("roe_high", inp["net_income"] / bv if bv > 0 else None)  # the sheet's C47: "NA" -> 0 growth
    ret = inp.get("retention_high", 1 - dps / eps)
    g = _weighted_growth(inp, ret * roe if roe is not None else 0.0)
    gs, pay = inp["stable_growth"], 1 - ret
    pay_s = inp["stable_payout"] if "stable_payout" in inp else 1 - gs / inp["roe_stable"]
    ks = _ke_stable(inp)
    rows, pv1 = [], 0.0
    for t in range(1, n1 + 1):
        d = eps * (1 + g) * pay * (1 + g) ** (t - 1)
        pv1 += d / (1 + ke) ** t
        rows.append({"t": t, "growth": g, "earnings": eps * (1 + g) ** t, "payout": pay, "dividends": d,
                     "cost_of_equity": ke, "pv": d / (1 + ke) ** t})
    disc, e, gt, pt, bt, pv2 = (1 + ke) ** n1, eps * (1 + g) ** n1, g, pay, inp["beta"], 0.0
    for t in range(1, n2 + 1):
        gt -= (g - gs) / n2
        pt = pay + (pay_s - pay) / n2 * t if "transition_payout" not in inp else inp["transition_payout"]
        bt = bt - (inp["beta"] - (ks - rf) / erp) / n2 if "transition_beta" not in inp else inp["transition_beta"]
        e *= 1 + gt
        disc *= 1 + rf + erp * bt
        pv2 += e * pt / disc
        rows.append({"t": n1 + t, "growth": gt, "earnings": e, "payout": pt, "dividends": e * pt,
                     "cost_of_equity": rf + erp * bt, "pv": e * pt / disc})
    price = e * (1 + gs) * pay_s / (ks - gs)
    return {"cost_of_equity": ke, "expected_growth": g, "stable_payout": pay_s, "stable_cost_of_equity": ks,
            "pv_high_growth": pv1, "pv_transition": pv2, "terminal_price": price, "pv_terminal": price / disc,
            "value_per_share": pv1 + pv2 + price / disc, "years": rows}


def fcfe_stable(inp: dict) -> dict:
    """fcfest.xls: eps, capex, depreciation, change_wc, debt_ratio, growth, cost_of_equity | CAPM; optional
    offset_capex (net capex 0) or roe_stable (reinvestment from fundamentals: g/ROE x EPS)."""
    ke, g, eps, dr = _ke(inp), inp["growth"], inp["eps"], inp["debt_ratio"]
    wc = (1 - dr) * inp["change_wc"]
    capex = (0.0 if inp.get("offset_capex") else g / inp["roe_stable"] * eps - wc if "roe_stable" in inp
             else (1 - dr) * (inp["capex"] - inp["depreciation"]))
    f = eps - capex - wc
    w = [m for c, m in ((g > 0.1, "This is high for a stable growth rate"),
                        (inp.get("beta", 0) > 1.5, "This Beta is high for a stable firm"),
                        (inp["capex"] > 1.5 * inp["depreciation"], "Capital Spending seems high relative to depreciation"))
         if c]
    return {"cost_of_equity": ke, "fcfe": f, "value_per_share": f * (1 + g) / (ke - g), "warnings": w}


def _stable_capex(inp: dict, e_s: float, dwc_s: float, dep_term: float) -> float:
    """Stable-period net capex: offset (0) | from fundamentals (g/ROE x earnings - dWC) | capex/dep ratio."""
    if inp.get("offset_capex_stable"):
        return 0.0
    if "roe_stable" in inp:
        return inp["stable_growth"] / inp["roe_stable"] * e_s - dwc_s
    return (inp["capex_to_depreciation_stable"] - 1) * dep_term


def fcfe_2stage(inp: dict) -> dict:
    """fcfe2st.xls: eps, dps, capex, depreciation, revenues, working_capital, change_wc, years (0-10),
    cost_of_equity | CAPM, growth_weights + eps_5y_ago / outside_growth / fundamentals (net_income,
    interest_expense, book_debt, book_debt_last_year, book_equity, book_equity_last_year, tax_rate; overrides
    roc_high, de_high, retention_high, interest_rate_high), stable_growth, beta_stable; optional capex_growth +
    depreciation_growth + revenue_growth + revenue_growth_stable (else all grow with earnings), wc_to_revenue,
    debt_ratio_capex + debt_ratio_wc (else D/(D+E)), offset_capex_stable | roe_stable | capex_to_depreciation_stable.
    replicate_sheet_bugs (false): true reproduces the sheet's row 121, which finances high-growth working capital at the
    capex debt ratio (E95) although the sheet's current FCFE and terminal year use the working-capital ratio (E96)."""
    bug = inp.get("replicate_sheet_bugs", False)
    n = _check(inp["years"], 0, 10, "10 year columns")
    eps, capex, dep, rev = inp["eps"], inp["capex"], inp["depreciation"], inp["revenues"]
    ke, (fund, de) = _ke(inp), _roc_growth(inp)
    g, gs = _weighted_growth(inp, fund), inp["stable_growth"]
    same = "capex_growth" not in inp
    gc, gd, gr = (g, g, g) if same else (inp["capex_growth"], inp["depreciation_growth"], inp["revenue_growth"])
    gr_s = gs if same else inp["revenue_growth_stable"]
    wcp = inp.get("wc_to_revenue", max(inp["working_capital"] / rev, 0.0))
    de_h = inp.get("de_high", de)
    drc, drw = inp.get("debt_ratio_capex", de_h / (1 + de_h)), inp.get("debt_ratio_wc", de_h / (1 + de_h))
    f0 = eps - (capex - dep) * (1 - drc) - inp["change_wc"] * (1 - drw)
    rows, pv = [], 0.0
    for t in range(1, n + 1):
        e = eps * (1 + g) ** t
        nc = (capex * (1 + gc) ** t - dep * (1 + gd) ** t) * (1 - drc)
        dwc = rev * wcp * ((1 + gr) ** t - (1 + gr) ** (t - 1)) * (1 - (drc if bug else drw))  # the sheet's row 121: drc
        pv += (e - nc - dwc) / (1 + ke) ** t
        rows.append({"t": t, "earnings": e, "net_capex": nc, "chg_wc": dwc, "fcfe": e - nc - dwc,
                     "pv": (e - nc - dwc) / (1 + ke) ** t})
    e_s = eps * (1 + g) ** n * (1 + gs)
    dwc_s = rev * (1 + gr) ** n * gr_s * wcp * (1 - drw)
    nc_s = _stable_capex(inp, e_s, dwc_s, dep * (1 + gd) ** n * (1 + gs) * (1 - drc))
    ks = _ke_stable(inp)
    price = (e_s - nc_s - dwc_s) / (ks - gs)
    rows.append({"t": "terminal", "earnings": e_s, "net_capex": nc_s, "chg_wc": dwc_s, "fcfe": e_s - nc_s - dwc_s})
    value = pv + price / (1 + ke) ** n
    aip = f0 / ks
    stable = f0 * (1 + gs) / (ks - gs) - aip
    out = {"cost_of_equity": ke, "expected_growth": g, "current_fcfe": f0, "stable_cost_of_equity": ks,
           "pv_fcfe": pv, "terminal_price": price, "pv_terminal": price / (1 + ke) ** n, "value_per_share": value,
           "value_assets_in_place": aip, "value_stable_growth": stable,
           "value_extraordinary_growth": value - stable - aip, "years": rows}
    his = None if bug or drc == drw else fcfe_2stage({**inp, "replicate_sheet_bugs": True})["value_per_share"]
    if his is not None and abs(his - value) > 1e-9 * abs(value):
        out["sheet_error"] = (f"Damodaran's fcfe2st as written gives {his:,.4f}/share (row 121 finances working capital at "
                              "the capex debt ratio); replicate_sheet_bugs=true reproduces it")
    return out


def fcfe_3stage(inp: dict) -> dict:
    """fcfe3st.xls: fcfe_2stage inputs + transition_years (1-10), optional capex/depreciation/revenue
    _growth_transition, transition_beta (else gradual to stable). Debt mix default uses current D/E."""
    n1 = _check(inp["years"], 0, 10, "10 high-growth columns")
    n2 = _check(inp["transition_years"], 1, 10, "10 transition columns; 0 zeroes the template's terminal year")
    eps, capex, dep, rev = inp["eps"], inp["capex"], inp["depreciation"], inp["revenues"]
    ke, rf, erp = _ke(inp), inp["riskfree_rate"], inp["erp"]
    fund, de = _roc_growth(inp)
    g, gs = _weighted_growth(inp, fund), inp["stable_growth"]
    same = "capex_growth" not in inp
    gc, gd, gr = (g, g, g) if same else (inp["capex_growth"], inp["depreciation_growth"], inp["revenue_growth"])
    gr_s = gs if same else inp["revenue_growth_stable"]
    wcp = inp.get("wc_to_revenue", inp["working_capital"] / rev)
    drc, drw = inp.get("debt_ratio_capex", de / (1 + de)), inp.get("debt_ratio_wc", de / (1 + de))
    ks = _ke_stable(inp)
    rows, pv1 = [], 0.0
    for t in range(1, n1 + 1):
        e = eps * (1 + g) ** t
        nc = (capex * (1 + gc) ** t - dep * (1 + gd) ** t) * (1 - drc)
        dwc = (rev * (1 + gr) ** t - rev * (1 + gr) ** (t - 1)) * wcp * (1 - drw)
        pv1 += (e - nc - dwc) / (1 + ke) ** t
        rows.append({"t": t, "growth": g, "earnings": e, "net_capex": nc, "chg_wc": dwc, "fcfe": e - nc - dwc,
                     "cost_of_equity": ke, "pv": (e - nc - dwc) / (1 + ke) ** t})
    c0, d0, r0 = capex * (1 + gc) ** n1, dep * (1 + gd) ** n1, rev * (1 + gr) ** n1
    disc, e, gt, bt, cum, cmax, pv2 = (1 + ke) ** n1, eps * (1 + g) ** n1, g, inp["beta"], 0.0, float("-inf"), 0.0
    for t in range(1, n2 + 1):
        gt -= (g - gs) / n2
        cum_prev, cum = cum, (1 + cum) * (1 + gt) - 1
        cmax = max(cmax, cum)
        e *= 1 + gt
        if same:
            nc = (c0 * (1 + cum) - d0 * (1 + cum)) * (1 - drc)
            dwc = (r0 * (1 + cum) - r0 * (1 + cum_prev)) * wcp * (1 - drw)
        else:
            gct, gdt, grt = inp["capex_growth_transition"], inp["depreciation_growth_transition"], inp["revenue_growth_transition"]
            nc = (c0 * (1 + gct) ** t - d0 * (1 + gdt) ** t) * (1 - drc)
            dwc = (r0 * (1 + grt) ** t - r0 * (1 + grt) ** (t - 1)) * wcp * (1 - drw)
        bt = bt - (inp["beta"] - (ks - rf) / erp) / n2 if "transition_beta" not in inp else inp["transition_beta"]
        disc *= 1 + rf + erp * bt
        pv2 += (e - nc - dwc) / disc
        rows.append({"t": n1 + t, "growth": gt, "earnings": e, "net_capex": nc, "chg_wc": dwc, "fcfe": e - nc - dwc,
                     "cost_of_equity": rf + erp * bt, "pv": (e - nc - dwc) / disc})
    e_s, dwc_s = e * (1 + gs), dwc * (1 + gr_s)
    # the sheet's N147 ratio branch grows depreciation with earnings growth and MAX(cumulated transition growth)
    nc_s = _stable_capex(inp, e_s, dwc_s, dep * (1 + cmax) * (1 + g) ** n1 * (1 + gs) * (1 - drc))
    price = (e_s - nc_s - dwc_s) / (ks - gs)
    rows.append({"t": "terminal", "earnings": e_s, "net_capex": nc_s, "chg_wc": dwc_s, "fcfe": e_s - nc_s - dwc_s,
                 "cost_of_equity": ks})
    return {"cost_of_equity": ke, "expected_growth": g, "stable_cost_of_equity": ks, "pv_high_growth": pv1,
            "pv_transition": pv2, "terminal_price": price, "pv_terminal": price / disc,
            "value_per_share": pv1 + pv2 + price / disc, "years": rows}


def fcfe_vs_ddm(inp: dict) -> dict:
    """fcfevsddm.xls: net_income, dividends, capex, depreciation, revenues, working_capital, net_debt_cash_flow,
    years, growth (all items), beta, riskfree_rate, erp, stable_growth, roe_stable, beta_stable (opt),
    cash_reinvestment_rate (default: cost of equity)."""
    n = _check(inp["years"], 0, 10, "10 year columns")
    ni, g, gs = inp["net_income"], inp["growth"], inp["stable_growth"]
    ke = inp["riskfree_rate"] + inp["beta"] * inp["erp"]
    wcp, rc, pay = inp["working_capital"] / inp["revenues"], inp.get("cash_reinvestment_rate", ke), inp["dividends"] / ni
    rows, pvf, pvd, cash = [], 0.0, 0.0, 0.0
    for t in range(1, n + 1):
        e = ni * (1 + g) ** t
        nc = inp["capex"] * (1 + g) ** t - inp["depreciation"] * (1 + g) ** t
        dwc = inp["revenues"] * wcp * ((1 + g) ** t - (1 + g) ** (t - 1))
        f = e - nc - dwc + inp["net_debt_cash_flow"] * (1 + g) ** t
        cash = cash * (1 + rc) + f - e * pay
        pvf += f / (1 + ke) ** t
        pvd += e * pay / (1 + ke) ** t
        rows.append({"t": t, "net_income": e, "fcfe": f, "dividends": e * pay, "cash_buildup": cash})
    e_s = ni * (1 + g) ** n * (1 + gs)
    ks = _ke_stable(inp)
    price = (1 - gs / inp["roe_stable"]) * e_s / (ks - gs)
    buildup = max((r["cash_buildup"] for r in rows), default=0.0)  # the sheet's N52 = MAX(D52:M52)
    vf = pvf + price / (1 + ke) ** n
    vd = pvd + price / (1 + ke) ** n + buildup / (1 + ke) ** n
    return {"cost_of_equity": ke, "terminal_price": price, "cash_buildup": buildup, "value_fcfe": vf,
            "value_ddm": vd, "difference": vf - vd, "years": rows}


MODELS: dict[str, Callable[[dict], dict]] = {f.__name__: f for f in (
    bank_fcfe, excess_return, ddm, fcfe, ddm_stable, ddm_2stage, ddm_3stage, fcfe_stable, fcfe_2stage,
    fcfe_3stage, fcfe_vs_ddm)}


def value(inp: dict) -> dict:
    inp = {k: v for k, v in inp.items() if not k.startswith("_")}
    for k in ("erp", "erp_stable"):
        if isinstance(inp.get(k), dict):  # Damodaran's Jan-2026 country/region ERPs via cost_of_capital.py
            from cost_of_capital import erp, load
            inp[k] = erp(inp[k], load())
    if inp.get("model") not in MODELS:
        raise ValueError(f"model must be one of {sorted(MODELS)}")
    return MODELS[inp["model"]](inp)


def report(out: dict) -> str:
    fmt = lambda v: f"{v:,.4f}" if isinstance(v, float) and abs(v) < 1 else f"{v:,.2f}" if isinstance(v, float) else str(v)
    lines = [f"{k:28s} {fmt(v)}" for k, v in out.items() if k not in ("years", "warnings") and v is not None]
    rows = out.get("years") or []
    if rows:
        cols = list(rows[0])
        lines += ["", "  ".join(f"{c[:12]:>12s}" for c in cols)]
        lines += ["  ".join(f"{fmt(r.get(c, '')):>12s}" for c in cols) for r in rows]
    lines += [f"WARNING (source sheet): {w}" for w in out.get("warnings", [])]
    return "\n".join(lines)


def selftest() -> int:
    cases = json.loads(CASES.read_text(encoding="utf-8"))
    bad = 0
    for c in cases:
        out = value(c["inputs"])
        for k, his in c["expected"].items():
            ours = out[k]
            ok = abs(ours - his) <= (1e-6 if abs(his) < 1 else 0) + 5e-4 * abs(his)
            bad += not ok
            print(f"{'ok ' if ok else 'BAD'} {c['name'][:52]:52s} {k:26s} Damodaran {his:14.6f} ours {ours:14.6f}")
    print(f"{len(cases)} cases, {bad} mismatches (tolerance 0.05% relative)")
    return 1 if bad else 0


# ---------------------------------------------------------------- --build-cases: source cells -> our inputs
def _yes(x: Any) -> bool:
    return str(x).strip().lower() == "yes"


def _norm_earnings(v: Callable) -> dict:
    n = lambda c: v("Normalized Earnings!" + c)
    if n("B2") == 1:
        return {"approach": 1, "net_income_history": [x for x in (n(c + "5") for c in "BCDEF") if isinstance(x, (int, float))]}
    return {"approach": 2, "roe": n("B8")}


def _read_citi(v: Callable) -> dict:
    i = lambda c: v("Inputs!" + c)
    return {"model": "bank_fcfe", "template": "CitiVal2023", "tier1_capital": i("B2"), "book_equity": i("B3"),
            "risk_adjusted_assets": i("B4"), "net_income": i("B5"), "asset_growth": i("B12"),
            "target_tier1_ratio": i("B13"), "target_roe": i("B14"), "stable_growth": i("B16"),
            "shares_outstanding": i("B19"), "price": i("B20"), "cost_of_equity": i("B21"),
            "stable_cost_of_equity": i("B22"), "failure_probability": i("B25")}


def _read_deutsche(v: Callable) -> dict:
    i = lambda c: v("Inputs!" + c)
    return {"model": "bank_fcfe", "template": "DeutscheBank2016", "tier1_capital": i("B2"),
            "tier1_capital_adjustment": i("B4"), "book_equity": i("B5"), "risk_adjusted_assets": i("B6"),
            "net_income": i("B7"), "asset_growth": i("B14"), "tier1_ratio_next_year": i("B15"),
            "target_tier1_ratio": i("B17"), "target_roe": i("B18"), "target_roe_10": i("B20"), "stable_growth": i("B22"),
            "cost_of_equity": i("B26"), "stable_cost_of_equity": v("Expected FCFE!M14"), "price": i("B28"),
            "shares_outstanding": i("B29"), "failure_probability": i("B32")}


def _read_eqexret(v: Callable) -> dict:
    i = lambda c: v("Inputs!" + c)
    d = {"model": "excess_return", "net_income": i("B2"), "book_equity": i("B3"), "book_equity_last_year": i("C3"),
         "eps": i("B4"), "dps": i("B5"), "shares_outstanding": i("B6"), "beta": i("B10"), "riskfree_rate": i("B11"),
         "erp": i("B12"), "years": int(i("B15")), "stable_growth": i("B31"), "gradual": _yes(i("B28"))}
    if _yes(i("B7")):
        d["normalize_earnings"] = _norm_earnings(v)
    for flag, keys in (("B20", {"roe_high": "B22", "retention_high": "B23"}), ("B24", {"roe_stable": "B26"}),
                       ("B34", {"stable_payout": "B35"}), ("B37", {"beta_stable": "B38", "erp_stable": "B39"})):
        if _yes(i(flag)):
            d.update({k: i(c) for k, c in keys.items()})
    return d


def _read_divginzu(v: Callable) -> dict:
    i = lambda c: v("Inputs!" + c)
    d = {"model": "ddm", "book_equity_last_year": i("C8"), "net_income": i("B9"), "dividends": i("B10"),
         "shares_outstanding": i("B11"), "price": i("B12"), "beta": i("B15"), "riskfree_rate": i("B16"),
         "erp": i("B17"), "years": int(i("B20")), "gradual": _yes(i("B34"))}
    for flag, keys in (("B27", {"roe_high": "B29", "retention_high": "B30"}), ("B31", {"roe_stable": "B33"}),
                       ("B38", {"stable_growth": "B39"}), ("B42", {"stable_payout": "B43"}), ("B45", {"beta_stable": "B46"})):
        if _yes(i(flag)):
            d.update({k: i(c) for k, c in keys.items()})
    if not _yes(i("B22")):
        d["growth"] = i("B23")
    return d


def _read_fcfeginzu(v: Callable) -> dict:
    i = lambda c: v("Inputs!" + c)
    d = {"model": "fcfe", "net_income": i("B6"), "interest_income_after_tax": i("B7"), "book_equity": i("B9"),
         "book_equity_last_year": i("C9"), "cash": i("B10"), "cash_last_year": i("C10"), "market_cap": i("B11"),
         "shares_outstanding": i("B12"), "capex": i("B13"), "depreciation": i("B14"), "change_wc": i("B15"),
         "net_debt_issued": i("B16"), "beta": i("B26"), "riskfree_rate": i("B27"), "erp": i("B28"),
         "years": int(i("B31")), "stable_growth": i("B50"), "gradual": _yes(i("B47"))}
    if _yes(i("B8")):
        r = lambda c: v("R&D converter!" + c)
        d["rd"] = {"life": int(r("F6")), "current": r("F7"), "past": [r(f"B{11 + k}") or 0.0 for k in range(int(r("F6")))]}
    if _yes(i("B17")):
        d["normalize_earnings"] = _norm_earnings(v)
    if _yes(i("B18")):
        n = lambda c: v("Normalized Reinvestment!" + c)
        row = lambda r: [x for x in (n(c + r) for c in "BCDEF") if isinstance(x, (int, float))]
        nr, a, b, c = {}, int(n("B2")), int(n("B12")), int(n("B22"))
        if a:
            nr["net_capex"] = ({"approach": 1, "net_capex_history": row("5"), "ebit_history": row("6")} if a == 1 else
                               {"approach": 2, "net_capex_to_ebit": n("B9"), "ebit": n("F6")})
        if b:
            nr["working_capital"] = {"approach": b, "revenues": n("B15"), "revenues_last_year": n("B16"),
                                     **({"non_cash_wc": n("B14")} if b == 1 else {"wc_to_revenue": n("B19")})}
        if c:
            nr["net_debt"] = {"approach": c, "book_debt": n("B24")} if c in (1, 2) else {"approach": 3, "debt_ratio": n("B27")}
        d["normalize_reinvestment"] = nr
    if not _yes(i("B33")):
        d["growth"] = i("B34")
    if _yes(i("B39")):
        d.update(roe_high=i("B41"), reinvestment_rate_high=i("B42"))
    if _yes(i("B43")):
        d["roe_stable"] = i("B45")
    if _yes(i("B53")):
        d["stable_reinvestment_rate"] = i("B54")
    elif not _yes(i("B43")):  # the sheet's B52 = B50/B45 uses the typed stable ROE even when it is switched off
        d["stable_reinvestment_rate"] = i("B50") / i("B45")
    if _yes(i("B56")):
        d["beta_stable"] = i("B57")
    if not _yes(i("B17")) or d["years"] in (2, 3):  # where the sheet's B38 / C14 slips change the value
        d["replicate_sheet_bugs"] = True
    return d


def _stage_reader(sheet: str, model: str, cells: dict, flags: list) -> Callable:
    """Stage templates: plain cell map + [(flag cell, when yes?, {key: cell | constant})]."""
    def read(v: Callable) -> dict:
        s = lambda c: v(f"{sheet}!{c}")
        d: dict = {"model": model}
        for k, c in cells.items():
            d[k] = [s(x) for x in c] if isinstance(c, list) else s(c)
        for flag, want, keys in flags:
            if _yes(s(flag)) == want:
                d.update({k: (s(c) if isinstance(c, str) else c) for k, c in keys.items()})
        for k in ("years", "transition_years"):
            if k in d:
                d[k] = int(d[k])
        return {k: x for k, x in d.items() if x is not None}
    return read


def _capm(b: str, r: str, e: str) -> dict:
    return {"beta": b, "riskfree_rate": r, "erp": e}


TEMPLATES: dict[str, dict] = {
    "CitiVal2023": {"read": _read_citi, "sheet": "Inputs",
                    "out": {"value_per_share": "Expected FCFE!B17", "value_per_share_adjusted": "Expected FCFE!B19"},
                    "branches": [("failure 10%", {"B25": 0.1}), ("stable ke 10%, g 2%", {"B22": 0.10, "B16": 0.02}),
                                 ("asset growth 6%, Tier1 18%, ROE 12%", {"B12": 0.06, "B13": 0.18, "B14": 0.12})]},
    "DeutscheBank2016": {"read": _read_deutsche, "sheet": "Inputs",
                         "out": {"value_per_share": "Expected FCFE!B18", "value_per_share_adjusted": "Expected FCFE!B20"},
                         "branches": [("DOJ fine 14bn, median risk", {"B3": "D1. DOJ Assessment",
                                                                      "B25": "A2. Median (ke =9.43%)"}),
                                      ("median Tier1 target and ROE", {"B16": "C2. Median (Tier 1 = 12.96% )",
                                                                       "B19": "B2. Median (ROE = 9.91%)"}),
                                      ("no wipeout, 2% growth", {"B32": 0, "B14": 0.02, "B22": 0.02})]},
    "eqexret": {"read": _read_eqexret, "sheet": "Inputs", "out": {"value_per_share": "Excess Return Valuation!B29"},
                "branches": [("normalized: average net income", {"B7": "Yes"}),
                             ("normalized: ROE", {"B7": "Yes", "Normalized Earnings!B2": 2}),
                             ("no high/stable overrides", {"B20": "No", "B24": "No"}),
                             ("no gradual, payout 50%, beta unchanged", {"B28": "No", "B34": "Yes", "B35": 0.5, "B37": "No"}),
                             ("5 years", {"B15": 5}), ("7 years", {"B15": 7})]},
    "divginzu": {"read": _read_divginzu, "sheet": "Inputs",
                 "out": {"value_per_share": "Valuation Output!E31", "value_assets_in_place": "Valuation Output!D35",
                         "value_stable_growth": "Valuation Output!D36", "value_extraordinary_growth": "Valuation Output!D37"},
                 "branches": [("direct growth 8%", {"B22": "No", "B23": 0.08}),
                              ("no high/stable overrides", {"B27": "No", "B31": "No"}), ("no gradual", {"B34": "No"}),
                              ("g 3%, payout 55%, beta 0.9", {"B38": "Yes", "B39": 0.03, "B42": "Yes", "B43": 0.55,
                                                              "B45": "Yes", "B46": 0.9}),
                              ("15 years", {"B20": 15}), ("5 years", {"B20": 5}), ("7 years", {"B20": 7})]},
    "fcfeginzu": {"read": _read_fcfeginzu, "sheet": "Inputs", "out": {"value_per_share": "FCFE Valuation!E32"},
                  "branches": [("reported earnings", {"B17": "No"}), ("normalized ROE", {"Normalized Earnings!B2": 2}),
                               ("normalized reinvestment 1/1/1", {"B18": "Yes"}),
                               ("normalized reinvestment 2/2/2", {"B18": "Yes", "Normalized Reinvestment!B2": 2,
                                                                  "Normalized Reinvestment!B12": 2,
                                                                  "Normalized Reinvestment!B22": 2}),
                               ("normalized reinvestment 0/0/3", {"B18": "Yes", "Normalized Reinvestment!B2": 0,
                                                                  "Normalized Reinvestment!B12": 0,
                                                                  "Normalized Reinvestment!B22": 3}),
                               ("R&D capitalized", {"B8": "Yes"}), ("direct growth 8%", {"B33": "No", "B34": 0.08}),
                               ("high-growth overrides", {"B39": "Yes"}), ("stable ROE switched off", {"B43": "No"}),
                               ("no gradual, stable ERR 30%, beta 0.8", {"B47": "No", "B53": "Yes", "B54": 0.3, "B56": "Yes"}),
                               ("0 years", {"B31": 0}), ("5 years", {"B31": 5}), ("7 years", {"B31": 7})]},
    "ddmst": {"read": _stage_reader("ddmst.xls", "ddm_stable", {"eps": "D18", "payout": "D19", "growth": "D28",
                                                                 **_capm("D24", "D25", "D26")},
                                    [("F21", True, {"cost_of_equity": "D22"})]),
              "sheet": "ddmst.xls", "out": {"value_per_share": "ddmst.xls!F44"},
              "branches": [("direct ke 11%", {"F21": "Yes", "D22": 0.11}), ("g 4%", {"D28": 0.04})]},
    "ddm2st": {"read": _stage_reader("ddm2st.xls", "ddm_2stage", {
        "eps": "D21", "dps": "D22", "years": "E24", **_capm("D29", "D30", "D31"), "eps_5y_ago": "E34",
        "outside_growth": "E37", "net_income": "D41", "book_equity_last_year": "E42",
        "growth_weights": ["E57", "E58", "E59"], "stable_growth": "E61", "roe_stable": "C53"}, [
        ("E26", True, {"cost_of_equity": "E27"}), ("G47", True, {"roe_high": "C49", "retention_high": "C50"}),
        ("E64", True, {"stable_payout": "E65"}), ("E67", True, {"beta_stable": "E68"})]),
        "sheet": "ddm2st.xls",
        "out": {"value_per_share": "ddm2st.xls!F104", "value_assets_in_place": "ddm2st.xls!E108",
                "value_stable_growth": "ddm2st.xls!E109", "value_extraordinary_growth": "ddm2st.xls!E110"},
        "branches": [("direct ke 11%", {"E26": "Yes", "E27": 0.11}), ("high inputs off", {"G47": "No"}),
                     ("stable payout 60%, beta 1", {"E64": "Yes", "E65": 0.6, "E67": "Yes", "E68": 1.0}),
                     ("8 years, weights .2/.2", {"E24": 8, "E57": 0.2, "E58": 0.2})]},
    "ddm3st": {"read": _stage_reader("ddm3st.xls", "ddm_3stage", {
        "eps": "D22", "dps": "D23", **_capm("D28", "D29", "D30"), "years": "E33", "eps_5y_ago": "E36",
        "outside_growth": "E39", "net_income": "D43", "book_equity_last_year": "E44",
        "growth_weights": ["E59", "E60", "E61"], "transition_years": "E64", "stable_growth": "E73", "roe_stable": "C55"}, [
        ("E25", True, {"cost_of_equity": "E26"}), ("G49", True, {"roe_high": "C51", "retention_high": "C52"}),
        ("G66", False, {"transition_payout": "G67"}), ("G69", False, {"transition_beta": "G70"}),
        ("E76", True, {"stable_payout": "E77"}), ("E79", True, {"beta_stable": "E80"})]),
        "sheet": "ddm3st.xls", "out": {"value_per_share": "ddm3st.xls!F121"},
        "branches": [("direct ke 12%", {"E25": "Yes", "E26": 0.12}), ("fundamental growth", {"D43": 10, "E44": 50}),
                     ("high inputs on", {"G49": "Yes", "D43": 10, "E44": 50}),
                     ("fixed transition payout & beta", {"G66": "No", "G67": 0.5, "G69": "No", "G70": 1.0}),
                     ("stable payout & beta", {"E76": "Yes", "E77": 0.55, "E79": "Yes", "E80": 0.9}),
                     ("6 + 4 years", {"E33": 6, "E64": 4})]},
    "fcfest": {"read": _stage_reader("NewFCFEStableGrowth", "fcfe_stable", {
        "eps": "D21", "capex": "D23", "depreciation": "D24", "change_wc": "D25", "debt_ratio": "D31",
        **_capm("D36", "D37", "D38"), "growth": "D40"}, [
        ("G27", True, {"offset_capex": True}), ("G29", True, {"roe_stable": "G30"}),
        ("F33", True, {"cost_of_equity": "D34"})]),
        "sheet": "NewFCFEStableGrowth", "out": {"value_per_share": "NewFCFEStableGrowth!F61"},
        "branches": [("offset capex", {"G27": "Yes"}), ("reported capex", {"G29": "No"}),
                     ("direct ke 11%", {"F33": "Yes", "D34": 0.11})]},
    "fcfe2st": {"read": _stage_reader("NewFCFE2Stage", "fcfe_2stage", {
        "eps": "D21", "dps": "D22", "capex": "D23", "depreciation": "D24", "revenues": "D25", "working_capital": "D26",
        "change_wc": "D27", "years": "E29", **_capm("D34", "D35", "D36"), "eps_5y_ago": "E40", "outside_growth": "E43",
        "net_income": "D47", "interest_expense": "D48", "book_debt": "D49", "book_debt_last_year": "E49",
        "book_equity": "D50", "book_equity_last_year": "E50", "tax_rate": "D51",
        "growth_weights": ["E61", "E62", "E63"], "stable_growth": "E65"}, [
        ("E31", True, {"cost_of_equity": "E32"}),
        ("F55", True, {"roc_high": "C57", "de_high": "E57", "retention_high": "C58", "interest_rate_high": "E58"}),
        ("E68", True, {"beta_stable": "E69"}),
        ("F72", False, {"capex_growth": "C75", "depreciation_growth": "D75", "revenue_growth": "E75",
                        "revenue_growth_stable": "E76"}),
        ("F78", False, {"wc_to_revenue": "E79"}),
        ("F81", False, {"debt_ratio_capex": "E83", "debt_ratio_wc": "E84", "replicate_sheet_bugs": True}),
        ("F87", True, {"offset_capex_stable": True}), ("F88", True, {"roe_stable": "F89"}),
        ("F88", False, {"capex_to_depreciation_stable": "F90"})]),
        "sheet": "NewFCFE2Stage",
        "out": {"value_per_share": "NewFCFE2Stage!F132", "value_assets_in_place": "NewFCFE2Stage!E136",
                "value_stable_growth": "NewFCFE2Stage!E137", "value_extraordinary_growth": "NewFCFE2Stage!E138"},
        "branches": [("direct ke 13%", {"E31": "Yes", "E32": 0.13}), ("high inputs on", {"F55": "Yes"}),
                     ("items grow with earnings", {"F72": "Yes"}),
                     ("WC 30%, debt 30%/20%", {"F78": "No", "E79": 0.3, "F81": "No", "E83": 0.3, "E84": 0.2}),
                     ("offset stable capex", {"F87": "Yes"}), ("stable capex/dep ratio", {"F88": "No"}),
                     ("stable beta unchanged", {"E68": "No"}), ("8 years", {"E29": 8})]},
    "fcfe3st": {"read": _stage_reader("NewFCFE3Stage", "fcfe_3stage", {
        "eps": "D23", "dps": "D24", "capex": "D25", "depreciation": "D26", "revenues": "D27", "working_capital": "D28",
        "change_wc": "D29", **_capm("D34", "D35", "D36"), "years": "E40", "book_debt": "D52", "book_equity": "D53",
        "growth_weights": ["E68", "E69", "E70"], "transition_years": "E73", "stable_growth": "E79"}, [
        ("E31", True, {"cost_of_equity": "E32"}), ("E42", True, {"eps_5y_ago": "E43"}),
        ("E45", True, {"outside_growth": "E46"}),
        ("F48", True, {"net_income": "D50", "interest_expense": "D51", "book_debt_last_year": "E52",
                       "book_equity_last_year": "E53", "tax_rate": "D54"}),
        ("F58", True, {"roc_high": "C60", "de_high": "E60", "retention_high": "C61", "interest_rate_high": "E61"}),
        ("F75", False, {"transition_beta": "F76"}), ("E81", True, {"beta_stable": "E82"}),
        ("F85", False, {"capex_growth": "C88", "depreciation_growth": "D88", "revenue_growth": "E88",
                        "capex_growth_transition": "C89", "depreciation_growth_transition": "D89",
                        "revenue_growth_transition": "E89", "revenue_growth_stable": "E90"}),
        ("F92", False, {"wc_to_revenue": "E93"}), ("F95", False, {"debt_ratio_capex": "E97", "debt_ratio_wc": "E98"}),
        ("F101", True, {"offset_capex_stable": True}), ("F102", True, {"roe_stable": "F103"}),
        ("F102", False, {"capex_to_depreciation_stable": "F104"})]),
        "sheet": "NewFCFE3Stage", "out": {"value_per_share": "NewFCFE3Stage!F163"},
        "branches": [("direct ke 12%", {"E31": "Yes", "E32": 0.12}),
                     ("historical + fundamental growth", {"E42": "Yes", "F48": "Yes", "E68": 0.3, "E69": 0.4, "E70": 0.3}),
                     ("fundamental overrides", {"F48": "Yes", "F58": "Yes", "E69": 0.5, "E70": 0.5}),
                     ("own item growth rates", {"F85": "No"}), ("WC 30%", {"F92": "No", "E93": 0.3}),
                     ("current debt mix", {"F95": "Yes"}), ("offset stable capex", {"F101": "Yes"}),
                     ("stable capex/dep ratio", {"F102": "No"}), ("fixed transition beta", {"F75": "No", "F76": 1.0}),
                     ("stable beta unchanged", {"E81": "No"}), ("3 + 5 years", {"E40": 3, "E73": 5})]},
    "fcfevsddm": {"read": _stage_reader("NewFCFE2Stage", "fcfe_vs_ddm", {
        "net_income": "D3", "dividends": "D4", "capex": "D5", "depreciation": "D6", "revenues": "D7",
        "working_capital": "D8", "net_debt_cash_flow": "D9", "years": "E11", "growth": "E12",
        **_capm("D15", "D16", "D17"), "stable_growth": "E19", "roe_stable": "E20"}, [
        ("E21", True, {"beta_stable": "E22"}), ("H24", False, {"cash_reinvestment_rate": "H25"})]),
        "sheet": "NewFCFE2Stage", "out": {"value_fcfe": "NewFCFE2Stage!F65", "value_ddm": "NewFCFE2Stage!G65"},
        "branches": [("cash reinvested at ke", {"H24": "Yes"}), ("stable beta 1.2", {"E21": "Yes", "E22": 1.2}),
                     ("8 years", {"E11": 8})]},
}


def build_cases(files: list[str], xlsx_dir: str) -> int:
    import warnings
    warnings.filterwarnings("ignore")
    import openpyxl
    import xlrd
    from openpyxl.utils.cell import coordinate_to_tuple
    from pycel import ExcelCompiler, excelwrapper

    def _names(self: Any) -> dict:  # pycel 1.0b30 predates openpyxl 3.1's defined-names API
        if self.workbook is not None and self._defined_names is None:
            self._defined_names = {str(d.name): [(a, w) for w, a in d.destinations if w in self.workbook]
                                   for d in self.workbook.defined_names.values()}
        return self._defined_names
    excelwrapper.ExcelOpxWrapper.defined_names = property(_names)

    def case(name: str, spec: dict, v: Callable) -> dict:
        inputs = spec["read"](v)
        c = {"name": name, "inputs": inputs, "expected": {k: float(v(a)) for k, a in spec["out"].items()}}
        out = value(inputs)
        for k, his in c["expected"].items():
            ok = abs(out[k] - his) <= 5e-4 * abs(his) + 1e-9
            print(f"{'ok ' if ok else 'BAD'} {name[:55]:55s} {k:26s} Damodaran {his:14.6f} ours {out[k]:14.6f}")
        return c

    def xls_getter(path: Path) -> Callable:
        book = xlrd.open_workbook(str(path))

        def v(a: str) -> Any:
            sheet, cell = a.rsplit("!", 1)
            r, c = coordinate_to_tuple(cell)
            sh = book.sheet_by_name(sheet)
            x = sh.cell_value(r - 1, c - 1) if r <= sh.nrows and c <= sh.ncols else None
            return None if x == "" else x
        return v

    cases = []
    for f in map(Path, files):
        spec = TEMPLATES[f.stem]
        if f.suffix == ".xls":
            v, xlsx = xls_getter(f), Path(xlsx_dir) / f"{f.stem}.xlsx"
        else:
            wb, xlsx = openpyxl.load_workbook(f, data_only=True), f
            v = lambda a, wb=wb: wb[a.rsplit("!", 1)[0]][a.rsplit("!", 1)[1]].value
        cases.append(case(f"{f.stem} (source file)", spec, v))
        for name, cells in spec["branches"]:
            xl = ExcelCompiler(filename=str(xlsx))
            for a in spec["out"].values():
                xl.evaluate(a)
            for a, x in cells.items():
                a = a if "!" in a else f"{spec['sheet']}!{a}"
                xl.evaluate(a)  # pycel only overrides cells already in its map
                xl.set_value(a, x)
            xl.recalculate()
            cases.append(case(f"{f.stem} branch: {name}", spec, xl.evaluate))
    CASES.write_text(json.dumps(cases, indent=1), encoding="utf-8")
    print(f"wrote {len(cases)} cases -> {CASES}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("inputs", nargs="?")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--from-case", help="print the source inputs for a case (e.g. CitiVal2023) as JSON")
    ap.add_argument("--build-cases", nargs="+", metavar="FILE")
    ap.add_argument("--xlsx-dir", default=".", help="Excel-converted .xlsx copies of Damodaran's .xls files (pycel needs xlsx)")
    a = ap.parse_args()
    if a.selftest:
        return selftest()
    if a.build_cases:
        return build_cases(a.build_cases, a.xlsx_dir)
    if a.from_case:
        cases = json.loads(CASES.read_text(encoding="utf-8"))
        hit = [c for c in cases if c["name"].lower().startswith(a.from_case.lower())]
        if not hit:
            print("cases:", ", ".join(c["name"] for c in cases))
            return 1
        print(json.dumps({"_source": f"Damodaran's own inputs, {hit[0]['name']} -> {hit[0]['expected']}",
                          **hit[0]["inputs"]}, indent=1))
        return 0
    if not a.inputs:
        ap.print_help()
        return 1
    out = value(json.loads(Path(a.inputs).read_text(encoding="utf-8")))
    print(json.dumps(out, indent=1) if a.json else report(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
