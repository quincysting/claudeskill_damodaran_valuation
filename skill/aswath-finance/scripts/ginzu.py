"""Deterministic FCFF valuation that mirrors Damodaran's fcffsimpleginzu spreadsheet.

Every formula below is transcribed from the 'Valuation output', 'R& D converter' and
'Option value' sheets of Damodaran's fcffsimpleginzu.xlsx (Jan 2026 edition) and is checked by
calibrate.py against 26 of Damodaran's published company valuations + the Ginzu template. Same inputs -> same value.

Usage:
  python ginzu.py inputs.json                    # value per share + year-by-year table
  python ginzu.py inputs.json --json             # machine-readable output
  python ginzu.py inputs.json --breakeven target_margin   # solve input so value = price
  python ginzu.py --selftest                     # replay 31 published Damodaran valuations + 14 template branch tests
  python ginzu.py --from-case NvidiaJan2025      # Damodaran's own inputs for a company, as a starting JSON

Inputs (JSON, currency units consistent, e.g. $ millions; rates as decimals):
  base: revenues, ebit, book_equity, book_debt, cash, non_operating_assets,
        minority_interests, shares_outstanding, price, effective_tax_rate, marginal_tax_rate
  drivers: growth_next_year, margin_next_year, growth_years_2_5, target_margin,
        margin_convergence_year, sales_to_capital_1_5, sales_to_capital_6_10, riskfree_rate,
        initial_cost_of_capital  OR  cost_of_capital: {...}  (computed by cost_of_capital.py from Damodaran's data)
        mature_erp               default: the Damodaran data file (Jan 2026 template: 4.23%)
  context (optional, feeds the Ginzu Diagnostics): industry_us (Damodaran's US industry name), revenues_last_year
  optional (the Ginzu "default assumption" overrides; omit to keep the defaults):
        terminal_cost_of_capital   default riskfree(after yr 10) + mature_erp
        terminal_roic              default = terminal cost of capital (no excess returns)
        failure_probability, failure_proceeds_basis ("B" book capital | "V" fair value),
        failure_proceeds_pct       default no failure
        reinvestment_lag           default 1 (0-3)
        keep_effective_tax         default false (tax converges to marginal in yrs 6-10)
        nol                        default 0 (loss carried forward into year 1)
        riskfree_after_10          default = riskfree_rate
        terminal_growth            default = riskfree after year 10
        trapped_cash, trapped_cash_tax_rate
        ipo_proceeds               primary proceeds retained by the firm (Damodaran's IPO models add them to equity)
        ipo_shares_at_value        true = new shares issued at estimated value (Damodaran's SpaceX model)
        rd: {life, current, past: [yr-1, yr-2, ...]}    capitalize R&D
        options: {number, strike, maturity, volatility, price_basis}  employee options;
                 price_basis "price" (template default) or "value" (value at estimated value/share)
        revenue_growth_path / margin_path: 10 values, replace the generated paths
        reinvestment_path: 10 values (segment models whose segments have different sales-to-capital)
        segments: [{name, market: [now, year5, year10], share: [now, year10], margin: [now, target]}]
                  Damodaran's Nvidia 2025 structure; revenues = total, the rest uses the standard drivers
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
CASES = HERE.parent / "data" / "calibration_cases.json"


def rd_adjustment(rd: dict) -> tuple[float, float]:
    """R& D converter sheet -> (EBIT adjustment, value of research asset)."""
    n, past = rd["life"], rd.get("past", [])
    asset, amort = rd["current"], 0.0
    for k, x in enumerate(past[:n], start=1):  # year -k
        asset += x * (n - k) / n
        amort += x / n
    return rd["current"] - amort, asset


def option_value(price: float, shares: float, opt: dict, riskfree: float) -> float:
    """Dilution-adjusted Black-Scholes, iterated like Excel's circular reference."""
    k, t, var, n = opt["strike"], opt["maturity"], opt["volatility"] ** 2, opt["number"]
    nd = lambda x: 0.5 * (1 + math.erf(x / math.sqrt(2)))
    w = 0.0
    for _ in range(200):
        s = (price * shares + w * n) / (shares + n)
        d1 = (math.log(s / k) + (riskfree + var / 2) * t) / math.sqrt(var * t)
        d2 = d1 - math.sqrt(var * t)
        new = s * nd(d1) - k * math.exp(-riskfree * t) * nd(d2)
        if abs(new - w) < 1e-12:
            break
        w = new
    return w * n


def standard_paths(inp: dict, base_revenue: float, g_term: float) -> tuple[list[float], list[float]]:
    """Ginzu revenue and margin paths for years 1..10 (index 0 = base year)."""
    growth = [0.0, inp["growth_next_year"]] + [inp["growth_years_2_5"]] * 4
    for k in range(1, 6):
        growth.append(growth[5] - (growth[5] - g_term) / 5 * k)
    rev = [base_revenue]
    for t in range(1, 11):
        rev.append(rev[-1] * (1 + growth[t]))
    m1, tgt, conv = inp["margin_next_year"], inp["target_margin"], inp["margin_convergence_year"]
    margin = [0.0, m1] + [tgt if t > conv else tgt - (tgt - m1) / conv * (conv - t) for t in range(2, 11)]
    return rev, margin


def segment_paths(inp: dict, g_term: float) -> tuple[list[float], list[float]]:
    """Damodaran's Nvidia-2025-style segment model: each segment = total market x market share, with its own
    margin; the rest of the business follows the standard drivers. Returns total revenue growth and
    total margin paths (years 1..10). Market: straight line to the year-5 value, then to year 10;
    share: straight line over 10 years; margin: converges like the Ginzu margin."""
    conv = inp["margin_convergence_year"]
    seg_rev, seg_ebit = [0.0] * 11, [0.0] * 11
    for s in inp["segments"]:
        (m0, m5, m10), (s0, s10), (mg1, mgt) = s["market"], s["share"], s["margin"]
        for t in range(11):
            mkt = m0 + (m5 - m0) / 5 * t if t <= 5 else m5 + (m10 - m5) / 5 * (t - 5)
            r = mkt * (s0 - (s0 - s10) / 10 * t)
            mg = mg1 if t <= 1 else (mgt if t > conv else mgt - (mgt - mg1) / conv * (conv - t))
            seg_rev[t] += r
            seg_ebit[t] += r * mg if t else 0.0
    rest_rev, rest_margin = standard_paths(inp, inp["revenues"] - seg_rev[0], g_term)
    total = [rest_rev[t] + seg_rev[t] for t in range(11)]
    ebit = [rest_rev[t] * rest_margin[t] + seg_ebit[t] for t in range(11)]
    return [total[t] / total[t - 1] - 1 for t in range(1, 11)], [ebit[t] / total[t] for t in range(1, 11)]


def value(inp: dict) -> dict:
    def g(k: str, d: Any = None) -> Any:
        return inp.get(k, d)

    if g("cost_of_capital") or "mature_erp" not in inp:  # compute from Damodaran's data tables
        from cost_of_capital import compute, load
        data = load()
        inp = {"mature_erp": data["mature_erp"], **inp}
        if g("cost_of_capital"):
            coc = compute({"riskfree_rate": inp["riskfree_rate"], "marginal_tax_rate": inp["marginal_tax_rate"],
                           **inp["cost_of_capital"]}, data)
            inp = {**inp, "initial_cost_of_capital": coc["cost_of_capital"], "_coc": coc}
    return _value(inp)


def _value(inp: dict) -> dict:
    def g(k: str, d: Any = None) -> Any:
        return inp.get(k, d)

    if g("segments") and not g("revenue_growth_path"):
        rf10_ = g("riskfree_after_10", inp["riskfree_rate"])
        gp, mp = segment_paths(inp, g("terminal_growth", rf10_))
        inp = {**inp, "revenue_growth_path": gp, "margin_path": mp}
    rf = inp["riskfree_rate"]
    rf10 = g("riskfree_after_10", rf)
    g_term = g("terminal_growth", rf10)
    r_term = g("terminal_cost_of_capital") or rf10 + inp["mature_erp"]

    ebit0, ic0 = inp["ebit"], inp["book_equity"] + inp["book_debt"] - inp["cash"]
    if g("rd"):
        adj, asset = rd_adjustment(inp["rd"])
        ebit0, ic0 = ebit0 + adj, ic0 + asset

    # --- paths for years 1..10 (index 0 = base year, 11 = terminal) ---
    if g("revenue_growth_path"):
        growth = [0.0] + list(inp["revenue_growth_path"])
    else:
        growth = [0.0, inp["growth_next_year"]] + [inp["growth_years_2_5"]] * 4
        for k in range(1, 6):
            growth.append(growth[5] - (growth[5] - g_term) / 5 * k)
    growth.append(g_term)

    rev = [inp["revenues"]]
    for t in range(1, 12):
        rev.append(rev[-1] * (1 + growth[t]))

    margin = [ebit0 / inp["revenues"]]
    if g("margin_path"):
        margin += list(inp["margin_path"])
    else:
        m1, tgt, conv = inp["margin_next_year"], inp["target_margin"], inp["margin_convergence_year"]
        margin.append(m1)
        for t in range(2, 11):
            margin.append(tgt if t > conv else tgt - (tgt - m1) / conv * (conv - t))
    margin.append(margin[10])
    ebit = [ebit0] + [margin[t] * rev[t] for t in range(1, 12)]

    t_eff, t_mar = inp["effective_tax_rate"], inp["marginal_tax_rate"]
    t_term = t_eff if g("keep_effective_tax") else t_mar
    tax = [t_eff] * 6
    for _ in range(5):
        tax.append(tax[-1] + (t_term - tax[5]) / 5)
    tax.append(t_term)

    # after-tax operating income with NOL shield
    nol = [g("nol", 0.0)]
    nopat = [ebit[0] * (1 - tax[0]) if ebit[0] > 0 else ebit[0]]
    for t in range(1, 11):
        e, prev = ebit[t], nol[t - 1]
        nopat.append((e if e < prev else e - (e - prev) * tax[t]) if e > 0 else e)
        nol.append(prev - e if e < 0 else (prev - e if prev > e else 0.0))
    nopat.append(ebit[11] * (1 - tax[11]))

    coc = [0.0] + [inp["initial_cost_of_capital"]] * 5
    for _ in range(5):
        coc.append(coc[-1] - (coc[5] - r_term) / 5)
    coc.append(r_term)
    roic_term = g("terminal_roic", coc[10])

    if not g("reinvestment_path") and not ("sales_to_capital_1_5" in inp and "sales_to_capital_6_10" in inp):
        raise KeyError("sales_to_capital_1_5 and sales_to_capital_6_10 are required (the Ginzu template has no default)")
    s2c = [0.0] + [g("sales_to_capital_1_5", 1.0)] * 5 + [g("sales_to_capital_6_10", 1.0)] * 5  # 1.0 unused with a path
    lag = g("reinvestment_lag", 1)
    reinv = [0.0]
    for t in range(1, 11):
        if lag == 0:
            d = rev[t] - rev[t - 1]
        elif lag == 1:
            d = rev[t + 1] - rev[t]
        elif lag == 2:
            d = rev[t + 2] - rev[t + 1] if t <= 9 else rev[11] * growth[11]
        else:  # lag 3: year 9 uses terminal revenue growth, year 10 grows year 9's reinvestment
            d = rev[t + 3] - rev[t + 2] if t <= 8 else rev[11] * growth[11]
        reinv.append(reinv[9] * (1 + g_term) if (lag == 3 and t == 10) else d / s2c[t])
    if g("reinvestment_path"):  # segment models with a different sales-to-capital per segment
        reinv[1:11] = inp["reinvestment_path"]
    reinv.append(g_term / roic_term * nopat[11] if g_term > 0 else 0.0)

    fcff = [0.0] + [nopat[t] - reinv[t] for t in range(1, 12)]
    df, pv = [1.0], [0.0]
    for t in range(1, 11):
        df.append(df[-1] / (1 + coc[t]))
        pv.append(fcff[t] * df[t])

    tv = fcff[11] / (r_term - g_term)
    sum_pv = sum(pv[1:11]) + tv * df[10]

    p_fail = g("failure_probability", 0.0)
    basis = g("failure_proceeds_basis", "V")
    proceeds = ((inp["book_equity"] + inp["book_debt"]) if basis == "B" else sum_pv) * g("failure_proceeds_pct", 0.0)
    op_assets = sum_pv * (1 - p_fail) + proceeds * p_fail

    cash = inp["cash"]
    if g("trapped_cash"):
        cash -= inp["trapped_cash"] * (t_mar - inp["trapped_cash_tax_rate"])
    # IPO proceeds (Damodaran's IPO workbooks): added to equity; SpaceX-style also issues new shares at value
    ipo = g("ipo_proceeds", 0.0)
    equity = op_assets - inp["book_debt"] - inp["minority_interests"] + cash + inp["non_operating_assets"] + ipo
    n0 = inp["shares_outstanding"]
    new_shares = lambda v: ipo / v if g("ipo_shares_at_value") and v else 0.0
    opts, per_share = 0.0, equity / n0
    new = per_share
    at_value = bool(g("options")) and inp["options"].get("price_basis", "price") == "value"
    for _ in range(200):  # fixed point: options at value and/or IPO shares at value are circular
        n = n0 + new_shares(per_share)
        if g("options"):
            # Ginzu template values options at the market price; Damodaran's Tesla models at estimated value
            opts = option_value(per_share if at_value else inp["price"], n0, inp["options"], rf)
        new = (equity - opts) / n
        if abs(new - per_share) < 1e-10:
            break
        per_share = new
    per_share = new

    ic = [ic0]
    for t in range(1, 11):
        ic.append(ic[-1] + reinv[t])
    roic = [nopat[0] / ic0] + [nopat[t] / ic[t - 1] for t in range(1, 11)] + [roic_term]
    pv_nopat = sum(nopat[t] * df[t] for t in range(1, 11))

    return {
        "cost_of_capital_detail": g("_coc"), "mature_erp": inp["mature_erp"],
        "invested_capital": ic, "base_margin": margin[0], "base_roic": roic[0],
        "marginal_roic": (nopat[10] - nopat[0]) / (ic[10] - ic[0]) if ic[10] != ic[0] else None,
        "compounded_cost_of_capital": (1 / df[10]) ** 0.1 - 1,
        "pv_reinvestment_share": (pv_nopat - sum(pv[1:11])) / pv_nopat if pv_nopat else 0.0,
        "value_per_share": per_share, "price": inp["price"], "price_to_value": inp["price"] / per_share,
        "operating_assets": op_assets, "pv_terminal": tv * df[10], "pv_10yr": sum(pv[1:11]),
        "terminal_value": tv, "equity": equity, "options_value": opts,
        "terminal_growth": g_term, "terminal_cost_of_capital": r_term, "terminal_roic": roic_term,
        "years": [{"t": t, "growth": growth[t], "revenue": rev[t], "margin": margin[t], "ebit": ebit[t],
                   "tax": tax[t], "nopat": nopat[t], "reinvestment": reinv[t], "fcff": fcff[t],
                   "cost_of_capital": coc[t], "roic": roic[t]} for t in range(1, 12)],
    }


def warnings(inp: dict, out: dict) -> list[str]:
    """Only rules Damodaran states outright (Ginzu default-assumption text, Myth 5.3); no invented thresholds."""
    w = []
    rf10 = inp.get("riskfree_after_10", inp["riskfree_rate"])
    if out["terminal_growth"] > rf10 + 1e-12:
        w.append("stable growth exceeds the riskfree rate: the Ginzu template caps it there to prevent 'impossible' growth")
    if out["terminal_roic"] < out["terminal_cost_of_capital"] - 1e-12:
        w.append("terminal ROIC below cost of capital: growth destroys value forever (Myth 5.3, 2016-11)")
    ratio = out["value_per_share"] / inp["price"] if inp["price"] else 0
    if out["value_per_share"] < 0 or ratio > 2 or ratio < 0.5:  # Ginzu Diagnostics step 6 flags
        w.append(f"value is {ratio:.0%} of price - the Ginzu Diagnostics say re-check growth, target margin, "
                 "sales-to-capital and terminal ROIC (see diagnostics)")
    return w


def diagnostics(inp: dict, out: dict) -> str:
    """The Ginzu 'Diagnostics' sheet: your forecasts next to industry averages, with the questions to ask."""
    ind = None
    if inp.get("industry_us"):
        from cost_of_capital import industry, load
        ind = industry(load(), inp["industry_us"], "US")
    y = out["years"]
    pct = lambda v: "   n/a" if v is None else f"{v:6.1%}"
    last = inp.get("revenues_last_year")
    ic0 = out["invested_capital"][0]
    rows = [
        f"1 Revenue growth  industry {pct(ind and ind['Annual Average Revenue growth - Last 5 years'])} | "
        f"last yr {pct(last and inp['revenues'] / last - 1)} | next yr {pct(y[0]['growth'])} | yrs 2-5 {pct(y[1]['growth'])}",
        "    ask: growth above industry -> is the company small? far from last year's growth -> why?",
        f"2 Revenues        now {inp['revenues']:,.0f} | yr1 {y[0]['revenue']:,.0f} | yr5 {y[4]['revenue']:,.0f} | yr10 {y[9]['revenue']:,.0f}",
        "    ask: how big is the total market, what do its biggest firms earn, what share does year 10 imply?",
        f"3 Operating margin industry {pct(ind and ind['Pre-tax Operating Margin (Lease & R&D adjusted)'])} | "
        f"now {pct(out['base_margin'])} | yr1 {pct(y[0]['margin'])} | yr5 {pct(y[4]['margin'])} | yr10 {pct(y[9]['margin'])}",
        "    ask: industry margins? unit economics (cost of the extra unit)? what does competition look like?",
        f"4 Sales/capital   industry {f'{ind[chr(83) + 'ales/Capital']:.2f}' if ind else 'n/a'} | now {inp['revenues'] / ic0:.2f} | "
        f"yr1-5 {inp.get('sales_to_capital_1_5', 'path')} | yr6-10 {inp.get('sales_to_capital_6_10', 'path')}",
        f"  ROIC  now {pct(out['base_roic'])} | marginal yrs1-10 {pct(out['marginal_roic'])} | yr10 {pct(y[9]['roic'])} | "
        f"stable {pct(out['terminal_roic'])} | value effect of reinvestment {out['pv_reinvestment_share']:.0%} of PV(NOPAT)",
        "    ask: bounce-back or new growth? excess capacity? does efficiency change with scale? comfortable with yr-10 ROIC?",
        f"5 Cost of capital industry {pct(ind and ind['Cost of capital'])} | yrs 1-5 {pct(y[0]['cost_of_capital'])} | "
        f"compounded {pct(out['compounded_cost_of_capital'])} | stable {pct(out['terminal_cost_of_capital'])} | "
        f"failure {pct(inp.get('failure_probability', 0))}",
        "    ask: vs industry? why does it change over time? is the failure rate consistent with the company?",
        f"6 Value/price     {out['value_per_share'] / inp['price']:.0%}  (Ginzu flags: <50% or >200% -> revisit growth, "
        "margin, sales/capital, terminal ROIC)",
    ]
    return "\n".join(rows)


def breakeven(inp: dict, field: str, lo: float, hi: float) -> float | None:
    """Bisection on one input so that value per share = price."""
    f = lambda x: value({**inp, field: x})["value_per_share"] - inp["price"]
    flo, fhi = f(lo), f(hi)
    if flo * fhi > 0:
        return None
    for _ in range(100):
        mid = (lo + hi) / 2
        fm = f(mid)
        if abs(fm) < 1e-9:
            break
        lo, flo, hi = (mid, fm, hi) if flo * fm > 0 else (lo, flo, mid)
    return (lo + hi) / 2


def report(inp: dict, out: dict) -> str:
    rows = ["year  growth   revenue      margin  EBIT(1-t)   reinvest     FCFF      WACC   ROIC"]
    for y in out["years"]:
        t = "term" if y["t"] == 11 else f"{y['t']:>4d}"
        rows.append(f"{t}  {y['growth']:6.1%} {y['revenue']:11,.0f}  {y['margin']:6.1%} {y['nopat']:10,.0f} "
                    f"{y['reinvestment']:10,.0f} {y['fcff']:10,.0f}  {y['cost_of_capital']:6.2%} {y['roic']:6.1%}")
    coc = out.get("cost_of_capital_detail")
    if coc:
        rows.insert(0, "Cost of capital: " + ", ".join(
            f"{k} {v:.4g}" if isinstance(v, float) else f"{k} {v}" for k, v in coc.items() if v is not None) + "\n")
    rows += ["",
             f"PV of FCFF yrs 1-10   {out['pv_10yr']:,.0f}",
             f"PV of terminal value  {out['pv_terminal']:,.0f}  (TV {out['terminal_value']:,.0f}, "
             f"g {out['terminal_growth']:.2%}, r {out['terminal_cost_of_capital']:.2%}, ROIC {out['terminal_roic']:.2%})",
             f"Operating assets      {out['operating_assets']:,.0f}",
             f"Equity                {out['equity']:,.0f}   options {out['options_value']:,.0f}",
             f"VALUE PER SHARE       {out['value_per_share']:,.2f}   price {out['price']:,.2f}   "
             f"price/value {out['price_to_value']:.0%}"]
    rows += [f"RULE: {w}" for w in warnings(inp, out)]
    rows += ["", "DIAGNOSTICS (Ginzu 'Diagnostics' sheet)", diagnostics(inp, out)]
    return "\n".join(rows)


def selftest() -> int:
    cases = json.loads(CASES.read_text(encoding="utf-8"))
    branch = CASES.parent / "branch_cases.json"  # the Ginzu template recalculated under every override branch
    if branch.exists():
        cases += json.loads(branch.read_text(encoding="utf-8"))
    bad = 0
    for c in cases:
        got = value(c["inputs"])["value_per_share"]
        err = abs(got / c["expected_value_per_share"] - 1)
        bad += err > 0.0005
        print(f"{'ok ' if err <= 0.0005 else 'BAD'} {c['name']:28s} Damodaran {c['expected_value_per_share']:10.2f}  ours {got:10.2f}  diff {err:.3%}")
    print(f"{len(cases) - bad}/{len(cases)} cases within 0.05% of Damodaran's spreadsheets")
    return 1 if bad else 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="?")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--breakeven", choices=["target_margin", "growth_years_2_5", "sales_to_capital_1_5",
                                            "initial_cost_of_capital"])
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--from-case", help="print Damodaran's inputs for a calibration case (e.g. NvidiaJan2025) as JSON")
    a = ap.parse_args()
    if a.selftest:
        return selftest()
    if a.from_case:
        cases = json.loads(CASES.read_text(encoding="utf-8"))
        hit = [c for c in cases if c["name"].lower().startswith(a.from_case.lower())]
        if not hit:
            print("cases:", ", ".join(c["name"] for c in cases))
            return 1
        c = hit[-1]  # prefer the segments-helper variant (listed after the plain case)
        print(json.dumps({"_source": f"Damodaran's own inputs, {c['name']} (value/share {c['expected_value_per_share']:.2f})",
                          **c["inputs"]}, indent=1))
        return 0
    inp = json.loads(Path(a.inputs).read_text(encoding="utf-8"))
    inp = {k: v for k, v in inp.items() if not k.startswith("_")}  # allow _source/_why notes
    out = value(inp)
    if a.json:
        print(json.dumps({**out, "rules": warnings(inp, out), "diagnostics": diagnostics(inp, out)}, indent=1))
    else:
        print(report(inp, out))
    if a.breakeven:
        lo, hi = {"target_margin": (-0.5, 0.95), "growth_years_2_5": (-0.3, 1.5),
                  "sales_to_capital_1_5": (0.05, 20), "initial_cost_of_capital": (0.01, 0.4)}[a.breakeven]
        x = breakeven(inp, a.breakeven, lo, hi)
        print(f"\nBREAKEVEN {a.breakeven} for value = price: " + ("none in range" if x is None else f"{x:.4f}"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
