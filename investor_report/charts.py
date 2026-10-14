import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt, numpy as np
from model import run, co
o = run()
BLUE, ORANGE, INK, MUTED, GRID = "#2a78d6", "#eb6834", "#0b0b0b", "#52514e", "#e4e3df"
plt.rcParams.update({"font.family":"serif","font.serif":["Palatino Linotype","Palatino","TeX Gyre Pagella","P052","DejaVu Serif"],
    "font.size":8.5,"axes.edgecolor":MUTED,"axes.labelcolor":INK,"xtick.color":MUTED,"ytick.color":MUTED,
    "axes.spines.top":False,"axes.spines.right":False})

# Figure 1: FCF yield by scenario
scen = [r["scen"]+f"\n${r['B']}" for r in o["XOM"]["rows"]]
x = np.arange(len(scen)); w = 0.36
fig, ax = plt.subplots(figsize=(5.5,2.4), dpi=300)
for i,(k,col) in enumerate([("XOM",BLUE),("CVX",ORANGE)]):
    vals = [r["fcf_yld"] for r in o[k]["rows"]]
    bars = ax.bar(x+(i-0.5)*w, vals, w-0.03, color=col, label=f"{co[k]['name']} ({k})", zorder=3)
    for b,v in zip(bars,vals): ax.text(b.get_x()+b.get_width()/2, v+0.15, f"{v:.1f}%", ha="center", va="bottom", fontsize=7, color=INK)
    dy = 100*co[k]["div"]/co[k]["mcap"]
    ax.axhline(dy, color=col, lw=1, ls=(0,(4,3)), zorder=2, label=f"{k} dividend yield ({dy:.1f}%)")
ax.set_xticks(x, scen); ax.set_xlim(-0.65, len(scen)-0.4); ax.set_ylabel("Free cash flow yield (%)")
ax.yaxis.grid(True, color=GRID, lw=0.6, zorder=0); ax.set_ylim(0,10.5)
ax.legend(frameon=False, loc="upper left", fontsize=7, ncol=2)
fig.tight_layout(); fig.savefig("fig1_fcf_yield.png"); plt.close(fig)

# Figure 2: Brent needed to fund commitments vs EIA path
fig, ax = plt.subplots(figsize=(5.5,2.6), dpi=300)
labels = ["Capex + dividend", "Capex + dividend + buyback"]
y = np.arange(2)
for i,(k,col) in enumerate([("XOM",BLUE),("CVX",ORANGE)]):
    vals = [o[k]["be_div"], o[k]["be_bb"]]
    bars = ax.barh(y+(0.5-i)*0.36, vals, 0.33, color=col, label=f"{co[k]['name']} ({k})", zorder=3)
    for b,v in zip(bars,vals): ax.text(v+0.8, b.get_y()+b.get_height()/2, f"${v:.0f}", va="center", fontsize=7, color=INK)
for p,lab,yy in [(105,"EIA 4Q26\n$105",1.55),(84,"EIA 2027\navg $84",1.55),(74,"EIA 4Q27\n$74",-0.45)]:
    ha = "right" if p==74 else "left"
    ax.axvline(p, color=MUTED, lw=0.8, ls=(0,(2,2)), zorder=1)
    ax.text(p+(1 if ha=="left" else -1), yy, lab, fontsize=6.3, color=MUTED, ha=ha, va="bottom" if yy>0 else "bottom")
ax.set_yticks(y, labels); ax.set_xlim(0,122); ax.set_ylim(-0.5,1.95)
ax.set_xlabel("Brent price needed ($/bbl), normalized 2025 cash flow base")
ax.xaxis.grid(True, color=GRID, lw=0.6, zorder=0)
ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5,-0.28), ncol=2, fontsize=7.5)
fig.tight_layout(); fig.savefig("fig2_breakevens.png"); plt.close(fig)
print("ok")
