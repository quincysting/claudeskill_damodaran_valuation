"""The WACC tab: Damodaran's cost-of-capital worksheet (cost_of_capital.py) as live Excel formulas, every approach:
detailed (beta single/multi/direct, ERP direct/country/countries/regions, debt direct/rating/synthetic, preferred),
industry average, distribution, direct. Data values come from Damodaran's tables and are shown as blue inputs."""
from cost_of_capital import GROUPS, industry
from xlsx_common import (FX, M, NOTE, NUM2, NUM4, PCT, PCT2, BOLD, BOX, TOP, hdr, label, printable, put, section, title,
                         widths)


class WaccSheet:
    def __init__(self, book):
        self.inp, self.data, self.ref, self.src, self.PS = book.inp, book.data, book.ref, book.src, book.PS

    def build(self, ws):
        ci = self.inp.get("cost_of_capital")
        a = (ci or {}).get("approach", "direct")
        title(ws, "Cost of Capital", f"Damodaran's cost-of-capital worksheet, approach: {a}. Data: Damodaran ({self.data['updated']}).")
        widths(ws, {"B": 42, **{c: 13 for c in "CDEFGH"}, "I": 3, "J": 60, "K": 10, "L": 10, "M": 10, "N": 10})
        label(ws, "B5", "Riskfree rate (live, incl. flex)")
        put(ws, "C5", f"={self.ref('riskfree_rate')}+{FX['d_rf']}", PCT2)
        label(ws, "B6", "Marginal tax rate")
        put(ws, "C6", f"={self.ref('marginal_tax_rate')}", PCT)
        self.r = 8
        fn = {"direct": self._coc_direct, "industry_average": self._coc_industry, "distribution": self._coc_dist,
              "detailed": self._coc_detailed}[a]
        cell = fn(ws, ci)
        label(ws, f"B{cell[1:]}", "Cost of capital (initial)", bold=True)
        ws[cell].border, ws[cell].font = BOX, BOLD
        printable(ws, one_page=True)
        return f"WACC!${cell[0]}${cell[1:]}"

    def note(self, ws, r, text):
        put(ws, f"J{r}", text, kind="label").font = NOTE

    def row(self, ws, lab, v, fmt=None, note=None, bold=False):
        label(ws, f"B{self.r}", lab, bold=bold)
        put(ws, f"C{self.r}", v, fmt, bold=bold)
        if note:
            self.note(ws, self.r, note)
        self.r += 1
        return f"C{self.r - 1}"

    def _coc_direct(self, ws, ci):
        section(ws, self.r, "Cost of capital (direct input)", last="J")
        self.r += 1
        base = self.row(ws, "Cost of capital (input)", ci["cost_of_capital"] if ci else f"={self.ref('initial_cost_of_capital')}",
                        PCT2, self.src("cost_of_capital") if ci else None)
        adj = self.row(ws, "Riskfree change (flex; moves it 1:1)", f"=$C$5-{self.ref('riskfree_rate')}", PCT2)
        return self.row(ws, "", f"={base}+{adj}", PCT2)

    def _industry_table(self, ws, rows, scope, metric, head):
        if len(rows) == 1 and "revenues" not in rows[0]:  # single business: the table value, no weighting
            v = industry(self.data, rows[0]["industry"], scope)[metric]
            cell = self.row(ws, f"{head}: {rows[0]['industry']} ({scope})", v, NUM4 if "Beta" in metric else PCT2,
                            f"Damodaran {scope} industry averages ({self.data['updated']})")
            self.r += 1
            return cell
        hdr(ws, self.r, {"B": "Business (industry)", "C": "Revenues", "D": "EV/Sales", "E": "Estimated value", "F": "Weight",
                         "G": head, "H": "Weighted"})
        first = self.r + 1
        for b in rows:
            self.r += 1
            d = industry(self.data, b["industry"], scope)
            put(ws, f"B{self.r}", b["industry"])
            put(ws, f"C{self.r}", b.get("revenues", 1), M)
            put(ws, f"D{self.r}", d["EV/Sales"], NUM2)
            put(ws, f"E{self.r}", f"=C{self.r}*D{self.r}", M)
            put(ws, f"G{self.r}", d[metric], NUM4 if "Beta" in metric else PCT2)
            put(ws, f"H{self.r}", f"=F{self.r}*G{self.r}", NUM4 if "Beta" in metric else PCT2)
        last = self.r
        for r in range(first, last + 1):
            put(ws, f"F{r}", f"=E{r}/SUM($E${first}:$E${last})", PCT)
        self.note(ws, first, f"Industry {metric} and EV/Sales: Damodaran {scope} industry averages ({self.data['updated']}); "
                             "weights = revenues x EV/Sales (Damodaran's multi-business method)")
        self.r += 1
        label(ws, f"B{self.r}", "Total / weighted", bold=True)
        put(ws, f"H{self.r}", f"=SUM(H{first}:H{last})", NUM4 if "Beta" in metric else PCT2, bold=True).border = TOP
        self.r += 2
        return f"H{self.r - 2}"

    def _coc_industry(self, ws, ci):
        scope = ci.get("scope", "US")
        section(ws, self.r, f"Industry-average cost of capital ({scope})", last="J")
        self.r += 1
        rows = ci.get("businesses") or [{"industry": ci["industry"]}]
        base = self._industry_table(ws, rows, scope, "Cost of capital", "Industry cost of capital")
        drf = self.row(ws, "Riskfree rate in the industry data", self.data["industry_riskfree"], PCT2, "Damodaran data")
        adj = self.row(ws, "Riskfree adjustment (live riskfree - data riskfree)", f"=$C$5-{drf}", PCT2)
        return self.row(ws, "", f"={base}+{adj}", PCT2)

    def _coc_dist(self, ws, ci):
        section(ws, self.r, f"Cost of capital distribution: {ci['region']}, {ci['group']}", last="J")
        self.r += 1
        base = self.row(ws, f"{ci['region']} {ci['group']}", self.data["coc_distribution"][ci["region"]][GROUPS.index(ci["group"])],
                        PCT2, f"Damodaran cost-of-capital distribution ({self.data['updated']})")
        drf = self.data.get("distribution_riskfree")
        adj = self.row(ws, "Riskfree adjustment", f"=$C$5-{drf}" if drf else 0, PCT2,
                       None if drf else "The Damodaran data carries no riskfree for the distribution: no adjustment")
        return self.row(ws, "", f"={base}+{adj}", PCT2)

    def _coc_detailed(self, ws, ci):
        b, e, d, p = ci["beta"], ci["erp"], ci["debt"], ci.get("preferred") or {}
        section(ws, self.r, f"Beta ({b['approach']})", last="J")
        self.r += 1
        if b["approach"] == "direct":
            bu, bl_direct = None, self.row(ws, "Levered beta (direct input)", b["levered_beta"], NUM4, self.src("beta"))
        else:
            scope = "US" if b["approach"].endswith("_us") else "Global"
            rows = b.get("businesses") or [{"industry": b["industry"]}]
            bu = self._industry_table(ws, rows, scope, "Unlevered Beta", "Unlevered beta")
        section(ws, self.r, f"Equity risk premium ({e['approach']})", last="J")
        self.r += 1
        if e["approach"] == "direct":
            erp = self.row(ws, "ERP (direct input)", e["value"], PCT2, self.src("erp"))
        elif e["approach"] == "country":
            erp = self.row(ws, f"ERP, {e['country']}", self.data["countries"][e["country"]]["erp"], PCT2, "Damodaran country ERPs")
        else:
            hdr(ws, self.r, {"B": "Country / region", "C": "Revenues", "D": "Weight", "E": "ERP", "F": "Weighted"})
            first = self.r + 1
            table = self.data["countries"] if e["approach"] == "countries" else self.data["regions"]
            for nm, v in e["weights"].items():
                self.r += 1
                rev = v["revenues"] if isinstance(v, dict) else v
                prem = v["erp"] if isinstance(v, dict) else (table[nm]["erp"] if e["approach"] == "countries" else table[nm])
                put(ws, f"B{self.r}", nm)
                put(ws, f"C{self.r}", rev, M)
                put(ws, f"E{self.r}", prem, PCT2)
                put(ws, f"F{self.r}", f"=D{self.r}*E{self.r}", PCT2)
            last = self.r
            for r in range(first, last + 1):
                put(ws, f"D{r}", f"=C{r}/SUM($C${first}:$C${last})", PCT)
            given = all(isinstance(v, dict) for v in e["weights"].values())
            self.note(ws, first, "ERPs as given in the inputs (see Sources)" if given else
                      f"ERPs: Damodaran {e['approach']} table ({self.data['updated']}) unless given with the weight")
            self.r += 1
            label(ws, f"B{self.r}", "Weighted ERP", bold=True)
            erp = f"F{self.r}"
            put(ws, erp, f"=SUM(F{first}:F{last})", PCT2, bold=True).border = TOP
            self.r += 2
        section(ws, self.r, f"Cost of debt ({d['approach']}) and market value of debt", last="J")
        self.r += 1
        bd = self.row(ws, "Book value of debt", d["book_debt"], M, self.src("book_debt"))
        intr = self.row(ws, "Interest expense", d["interest_expense"], M)
        mat = self.row(ws, "Average maturity (years)", d.get("maturity", 0), NUM2)
        if d["approach"] == "direct":
            kd = self.row(ws, "Pre-tax cost of debt (direct input)", d["pre_tax_cost"], PCT2)
        elif d["approach"] == "rating":
            self.row(ws, "Bond rating", d["rating"])
            sp = self.row(ws, "Default spread for the rating", self.data["rating_spreads"][d["rating"]], PCT2, "Damodaran rating spreads")
            kd = self.row(ws, "Pre-tax cost of debt  = riskfree + spread", f"=$C$5+{sp}", PCT2)
        else:
            tab = self.data["synthetic_large" if d.get("firm_type", 2) == 1 else "synthetic_small"]
            t0 = self.r
            hdr(ws, t0, {"K": "Coverage from", "L": "to", "M": "Rating", "N": "Spread"})
            for k, (lo, hi, rt, s) in enumerate(tab, start=1):
                for c, v, fmt in (("K", lo, NUM2), ("L", hi, NUM2), ("M", rt, None), ("N", s, PCT2)):
                    put(ws, f"{c}{t0 + k}", v, fmt)
            rng = f"$K${t0 + 1}:$N${t0 + len(tab)}"
            eb = self.row(ws, "EBIT for interest coverage", d["ebit"], M)
            cov = self.row(ws, "Interest coverage ratio", f"=IF({intr}=0,1000000,IF({eb}<0,-100000,{eb}/{intr}))", NUM2,
                           f"Synthetic rating table (firm type {d.get('firm_type', 2)}), Damodaran, in K:N")
            self.row(ws, "Synthetic rating", f"=VLOOKUP({cov},{rng},3,TRUE)")
            sp = self.row(ws, "Default spread", f"=VLOOKUP({cov},{rng},4,TRUE)", PCT2)
            cds = self.row(ws, "Country default spread", self.data["countries"].get(d.get("country", ""), {}).get("default_spread", 0.0),
                           PCT2, f"Damodaran, {d.get('country', 'none')}")
            kd = self.row(ws, "Pre-tax cost of debt", f"=$C$5+{sp}+{cds}", PCT2)
        mvd = self.row(ws, "Market value of debt", f"=IF({mat}>0,{intr}*(1-(1+{kd})^(-{mat}))/{kd}+{bd}/(1+{kd})^{mat},{bd})", M)
        section(ws, self.r + 1, "Capital structure (market values)", last="J")
        self.r += 2
        px = self.row(ws, "Share price", ci["price"], self.PS)
        sh = self.row(ws, "Shares outstanding", ci["shares_outstanding"], M)
        mve = self.row(ws, "Market value of equity", f"={px}*{sh}", M)
        ps_, pp, pdv = (self.row(ws, lab, p.get(k, 0), fmt) for lab, k, fmt in (
            ("Preferred shares", "shares", M), ("Preferred price", "price", self.PS), ("Preferred dividend per share", "dividend", self.PS)))
        mvp = self.row(ws, "Market value of preferred", f"={ps_}*{pp}", M)
        cap = self.row(ws, "Total capital", f"={mve}+{mvd}+{mvp}", M)
        section(ws, self.r + 1, "Cost of capital", last="J")
        self.r += 2
        t = "$C$6"
        bl = self.row(ws, "Levered beta" + ("" if bu is None else "  = Bu x (1 + (1 - t) x D/E)"),
                      f"={bl_direct}" if bu is None else f"={bu}*(1+(1-{t})*{mvd}/{mve})", NUM4)
        ke = self.row(ws, "Cost of equity  = Rf + Bl x ERP", f"=$C$5+{bl}*{erp}", PCT2)
        kp = self.row(ws, "Cost of preferred", f"=IF({mvp}>0,{pdv}/{pp},0)", PCT2)
        for lab, f in (("Equity / capital", f"={mve}/{cap}"), ("Debt / capital", f"={mvd}/{cap}"), ("Preferred / capital", f"={mvp}/{cap}")):
            self.row(ws, lab, f, PCT)
        return self.row(ws, "", f"=({mve}*{ke}+{mvd}*{kd}*(1-{t})+{mvp}*{kp})/{cap}", PCT2)
