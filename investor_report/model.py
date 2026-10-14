"""Exxon vs Chevron oil-price scenario model. All $ in billions unless noted.
Data current to Oct 13, 2026. Sources listed in the report's References."""
B25 = 69.1        # 2025 Brent average, $/bbl (EIA via secondary reporting)
HESS_CFO24, HESS_B24, HESS_SENS = 6.353, 80.5, 0.10   # Hess 2024 10-K CFO ex-WC; 2024 Brent; assumed Hess $/bbl sensitivity
HESS_MISSING_SHARE = 0.55                              # deal closed Jul 18 2025, so ~55% of 2025 excluded
hess_adj = HESS_MISSING_SHARE * (HESS_CFO24 - HESS_SENS*(HESS_B24-B25))

co = {
 "XOM": dict(name="ExxonMobil", px=173.24, sh=14.5/3.48, base=59.7, sens=0.70, capex=28.0, div=17.2, bb=20.0, me_loss=4*1.3),
 "CVX": dict(name="Chevron",    px=217.73, sh=12.1/6.11, base=33.9+hess_adj, sens=0.60, capex=18.5, div=4*1.78*(12.1/6.11), bb=None, me_loss=0.0),
}
SCEN = [("Glut", 60, False), ("Peace (EIA 4Q27)", 74, False), ("EIA 2027 avg", 84, False), ("War persists", 105, True)]

def cfo(c, B, war): return c["base"] + c["sens"]*(B-B25) - (c["me_loss"] if war else 0.0)

def run():
    out = {}
    x = co["XOM"]; x["mcap"] = x["px"]*x["sh"]
    BB_PCT = x["bb"]/x["mcap"]   # Exxon plan as % of market value, applied to both for a like-for-like test
    c = co["CVX"]; c["mcap"] = c["px"]*c["sh"]; c["bb"] = BB_PCT*c["mcap"]
    for k,c in co.items():
        c["mcap"] = c["px"]*c["sh"]
        rows=[]
        for nm,B,war in SCEN:
            f = cfo(c,B,war); fcf = f-c["capex"]; res = fcf-c["div"]
            rows.append(dict(scen=nm,B=B,cfo=f,fcf=fcf,fcf_yld=100*fcf/c["mcap"],resid=res,resid_yld=100*res/c["mcap"],bb_cov=res/c["bb"]))
        be_div = B25 - (c["base"]-c["capex"]-c["div"])/c["sens"]
        be_bb  = B25 - (c["base"]-c["capex"]-c["div"]-c["bb"])/c["sens"]
        out[k]=dict(c=c,rows=rows,be_div=be_div,be_bb=be_bb)
    return out

if __name__=="__main__":
    print("Hess adjustment: +$%.2fB" % hess_adj)
    o=run()
    for k,v in o.items():
        c=v["c"]
        print(f"\n{k}: mcap {c['mcap']:.0f} shares {c['sh']:.3f} div {c['div']:.2f} div yld {100*c['div']/c['mcap']:.2f}% base CFO {c['base']:.1f} sens/mcap {100*c['sens']/c['mcap']:.3f}%")
        for r in v["rows"]:
            print("  {scen:18s} ${B:>3} CFO {cfo:5.1f} FCF {fcf:5.1f} FCFyld {fcf_yld:4.1f}% afterDiv {resid:5.1f} ({resid_yld:4.1f}%) buyback cover {bb_cov:4.2f}x".format(**r))
        print(f"  Breakeven capex+div ${v['be_div']:.0f}; +buyback(${c['bb']:.0f}B) ${v['be_bb']:.0f}")
        w,p = v["rows"][3], v["rows"][1]
        print(f"  War->Peace FCF change {p['fcf']-w['fcf']:.1f} ({100*(p['fcf']/w['fcf']-1):.0f}%)")
