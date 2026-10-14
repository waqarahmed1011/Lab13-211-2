"""Builds Investor_Analyst_Report.docx in JDF 2.2 format. Numbers come from model.py."""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from model import run, co, hess_adj

o = run()
X, C = o["XOM"], o["CVX"]
xr = {r["B"]: r for r in X["rows"]}
cr = {r["B"]: r for r in C["rows"]}
FONT = "Palatino Linotype"

doc = Document()
for s in doc.sections:
    s.page_width, s.page_height = Inches(8.5), Inches(11)
    s.top_margin, s.bottom_margin = Inches(1), Inches(1.5)
    s.left_margin = s.right_margin = Inches(1.5)
    s.footer_distance = Inches(1)

st = doc.styles["Normal"]
st.font.name = FONT
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.font.size = Pt(11)
pf = st.paragraph_format
pf.line_spacing = 1.26
pf.space_after = Pt(8.5)
pf.space_before = Pt(0)
pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

cp = doc.core_properties
cp.author = ""; cp.last_modified_by = ""; cp.title = ""; cp.subject = ""; cp.keywords = ""; cp.comments = ""

# page number footer
fp = doc.sections[0].footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run()
for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
    if tag:
        e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), tag); r._r.append(e)
    else:
        e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = text; r._r.append(e)


def runs(p, text, size=None):
    """Tiny markup: **bold**, *italic*."""
    import re
    for part in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not part:
            continue
        if part.startswith("**"):
            rr = p.add_run(part[2:-2]); rr.bold = True
        elif part.startswith("*"):
            rr = p.add_run(part[1:-1]); rr.italic = True
        else:
            rr = p.add_run(part)
        if size:
            rr.font.size = Pt(size)
    return p


def para(text, **kw):
    return runs(doc.add_paragraph(), text, **kw)


def h1(text):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(11); p.paragraph_format.space_after = Pt(8.5)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.keep_with_next = True
    rr = p.add_run(text.upper()); rr.bold = True


def h2(text):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(8.5); p.paragraph_format.space_after = Pt(8.5)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.keep_with_next = True
    rr = p.add_run(text); rr.bold = True


def caption(label, text, keep_next=False):
    p = doc.add_paragraph()
    f = p.paragraph_format
    f.left_indent = f.right_indent = Inches(1)
    f.line_spacing_rule = WD_LINE_SPACING.EXACTLY; f.line_spacing = Pt(14)
    f.space_after = Pt(8.5); f.keep_with_next = keep_next
    rr = p.add_run(label + "—"); rr.bold = True; rr.italic = True; rr.font.size = Pt(8.5)
    runs(p, text, size=8.5)


def figure(path, label, text, width=5.5):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0; p.paragraph_format.space_after = Pt(2); p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(path, width=Inches(width))
    caption(label, text)


def set_cell_border(cell, **edges):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for edge, val in edges.items():
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(val)); e.set(qn("w:color"), "000000")
        b.append(e)
    tcPr.append(b)


def table(label, text, header, rows, widths):
    caption(label, text, keep_next=True)
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    grid = t._tbl.tblGrid
    for gc, wdt in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(wdt * 1440)))
    for i, row in enumerate([header] + rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j); cell.width = Inches(widths[j])
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(1); p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.keep_with_next = (i < len(rows))
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            rr = p.add_run(val); rr.font.size = Pt(8.5); rr.bold = (i == 0)
            if i == 0:
                set_cell_border(cell, top=8, bottom=4)
            elif i == len(rows):
                set_cell_border(cell, bottom=8)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def f1(v): return f"{v:.1f}"

# ---------------------------------------------------------------- title block
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.line_spacing = 1.15
rr = p.add_run("ExxonMobil or Chevron When the War Premium Fades"); rr.font.size = Pt(17)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Investor / Analyst Report · Energy: U.S. Integrated Oil\nData and analysis current through October 13, 2026")

ab = doc.add_paragraph()
ab.paragraph_format.left_indent = ab.paragraph_format.right_indent = Inches(0.5)
rr = ab.add_run("Abstract—"); rr.bold = True; rr.italic = True
runs(ab, (
    "Brent crude trades near $105 because of the Iran conflict and disruption around the Strait of Hormuz, and "
    "ExxonMobil (XOM) and Chevron (CVX) shares are both up roughly 30 to 40 percent in 2026. The U.S. Energy "
    "Information Administration (EIA) expects most of that premium to fade by late 2027. This report asks which "
    "major is the better holding for an investor who wants integrated oil exposure without betting on the war "
    "lasting. Using a normalized cash flow model built on 2025 results and each company's disclosed oil price "
    f"sensitivity, Chevron offers the higher free cash flow (FCF) yield at the EIA's late 2027 price of $74 "
    f"({cr[74]['fcf_yld']:.1f}% versus {xr[74]['fcf_yld']:.1f}%) and in three of our four scenarios, even though the "
    "model gives it no credit for growth at Tengiz. Exxon is the lower-risk stock: its FCF falls "
    f"{100*(1-xr[74]['fcf']/xr[105]['fcf']):.0f}% from the war case to the peace case versus "
    f"{100*(1-cr[74]['fcf']/cr[105]['fcf']):.0f}% for Chevron, and it covers its dividend more comfortably at $60. "
    "**We prefer CVX over XOM.** The call flips toward Exxon if oil settles below about $67 or if Chevron's "
    "third quarter results fail to show the cash flow our conservative model assumes."))

# ---------------------------------------------------------------- 1
h1("1 The question and the recommendation")
para("**Investment question.** For an investor adding a position in a U.S. integrated oil major today, is "
     "ExxonMobil or Chevron the better risk-adjusted choice, given that current oil prices carry a war premium "
     "that most forecasters expect to shrink?")
para("**Recommendation: prefer Chevron over ExxonMobil.** This is a relative call, not a claim that either "
     "stock is cheap. Three findings drive it. First, after normalizing away the war windfall, Chevron generates "
     f"more free cash flow per dollar of market value: {cr[74]['fcf_yld']:.1f}% versus {xr[74]['fcf_yld']:.1f}% at "
     "$74 Brent, and more in every scenario except a glut below the EIA's path. Second, that result is "
     "conservative for Chevron. The model leaves out its Tengiz expansion, and management says its breakeven is "
     "below $50, about $10 under our estimate. Third, Chevron pays the higher dividend (3.3% versus 2.4%), and "
     "its dividend stays covered down to about $60 Brent. The cost of this choice is risk: Chevron is more "
     "exposed to falling oil, and Exxon has a hedge Chevron lacks. Roughly 20% of Exxon's 2025 production came "
     "from Qatar and the UAE [4], so if the conflict ends, oil falls but those barrels come back.")

# ---------------------------------------------------------------- 2
h1("2 Why this question, why now")
para("Brent closed at $71.32 on February 27, 2026, the last trading day before the conflict began, and passed "
     "$100 by mid-March [20]. It averaged about $104 in the second quarter [14] and $114 in September, and "
     "traded near $105 on October 13 [13][21]. The EIA's October outlook forecasts Brent at $105 in the fourth "
     "quarter of 2026, $84 on average in 2027, and $74 by the fourth quarter of 2027. That path assumes "
     "about 4.5 million barrels per day of Middle East output stays shut in for now, with most of it returning "
     "by mid-2027 [13]. The EIA's own baseline, in other words, is that today's price is temporary.")
para("Equity investors have priced the windfall aggressively. By late July, Chevron was up 29% for the year "
     "and Exxon about two points ahead of it [19]. With both companies reporting third quarter results at "
     "the end of October, the question for a new buyer is less about which company is better run and more "
     "about which one still earns its share price when $105 oil becomes $74 oil.")

# ---------------------------------------------------------------- 3
h1("3 The two companies")
para("Table 1 sets out the inputs used throughout. Market values use second quarter diluted share counts "
     "(for Exxon, implied by its reported EPS). Exxon is the larger, more diversified business, with "
     "bigger refining and chemicals operations. Chevron completed its $53 billion acquisition of Hess on July "
     "18, 2025, gaining a 30% stake in Guyana's Stabroek block [18]. Because Hess was only in Chevron's numbers "
     "for about five months of 2025, Chevron's 2025 cash flow understates its current scale; Section 4 adjusts for "
     "that.")
para("The exposures differ sharply. Exxon's Qatar LNG and UAE assets were disrupted from March, and the "
     "company's own earnings bridges show Middle East volumes cutting upstream earnings by $1.28 billion in "
     "the first half versus 2025, about $1.03 billion of it in the second quarter [4][5]. Chevron's direct regional exposure is the Saudi Arabia and Kuwait Partitioned Zone, about 1% of "
     "second quarter output, plus its Leviathan gas field off Israel, which was shut for 33 days and restarted "
     "on April 2 [8][17].")

table("Table 1", "Key inputs. Sources in brackets. Cash flow (CFO) is operating cash flow excluding working capital.",
      ["Item", "XOM", "CVX"],
      [["Share price, Oct 13, 2026 [16]", "$173.24", "$217.73"],
       ["Diluted shares, Q2 2026 [1][6]", f"{co['XOM']['sh']:.2f}B", f"{co['CVX']['sh']:.2f}B"],
       ["Market value", f"${co['XOM']['mcap']:.0f}B", f"${co['CVX']['mcap']:.0f}B"],
       ["Annual dividend at current rate [1][8]", f"${co['XOM']['div']:.1f}B", f"${co['CVX']['div']:.1f}B"],
       ["Dividend yield", f"{100*co['XOM']['div']/co['XOM']['mcap']:.1f}%", f"{100*co['CVX']['div']/co['CVX']['mcap']:.1f}%"],
       ["2025 CFO [2][7]", "$59.7B", "$34.9B"],
       ["2026 capex guidance [2][9]", "$27–29B", "$18–19B"],
       ["Buyback plan / pace [2][11]", "$20B", "$10–20B / ~$11B"],
       ["Mgmt. sensitivity per $1 Brent [3][9]", "~$0.70B", "~$0.60B"],
       ["Middle East share of output [4][17]", "~20%", "~1%"]],
      [3.1, 1.1, 1.3])
# ---------------------------------------------------------------- 4
h1("4 Method: normalizing cash flow")
h2("4.1 Why not just annualize the latest quarter")
para("The obvious shortcut is to multiply second quarter cash flow by four. That gives $102.6 billion for "
     "Exxon and $78.8 billion for Chevron. We tested whether oil prices alone explain those figures. Q2 Brent "
     "averaged about $103.50 versus $69.10 in 2025 [14][15]. Backing out that $34.40 difference with each "
     "company's own sensitivity still leaves $78.5 billion for Exxon and $58.2 billion for Chevron at 2025 oil "
     "prices, compared with actual 2025 cash flow of $59.7 billion and $34.9 billion. Roughly $19 billion "
     "(Exxon) and $23 billion (Chevron) of the run rate is unexplained by crude. Part of it is real growth, "
     "chiefly Hess for Chevron. Much of it is the war itself. Refining margins spiked, and Exxon's own guidance puts the "
     "earnings effect of each $1 per barrel of refining margin at about $800 million a year, more than its "
     "$700 million per $1 of Brent [5]. Those margins should normalize along with crude. Annualizing Q2 would treat a wartime "
     "quarter as the new normal, so we reject it as a base case.")
h2("4.2 The model")
para("We start from 2025 operating cash flow, a year with no war premium (Brent $69.10), and move it to each "
     "scenario price using the company's disclosed sensitivity: about $0.70 billion per $1 for Exxon [3] and "
     "$0.60 billion for Chevron [9]. Cash flow at price *P* is the 2025 base plus sensitivity times "
     "(*P* minus $69.10). For Chevron we add the missing part of Hess. Hess reported $6.35 billion of 2024 cash "
     "flow before working capital at roughly $80 Brent [12]. Adjusting that to $69 with an assumed $0.10 "
     f"billion per $1 (our assumption, not a sourced figure; zero or double would move the adjustment by about $0.6 billion and Chevron's breakevens by about $1) and taking the 55% of 2025 that preceded the deal gives +${hess_adj:.1f} billion, for a "
     f"Chevron base of ${co['CVX']['base']:.1f} billion. Free cash flow subtracts the capex guidance midpoint "
     "for Exxon ($28 billion) and $18.5 billion for Chevron, which guided to the low end of its range [9]. "
     "Exxon plans $20 billion of buybacks, 2.8% of its market value. To compare like with like, we test "
     f"Chevron at the same 2.8%, about ${co['CVX']['bb']:.0f} billion, which sits inside its $10 to $20 "
     "billion range and slightly above its first half pace of $5.5 billion [11].")
para("We run four Brent scenarios: a post-war glut at $60, the EIA's late 2027 price of $74, its 2027 "
     "average of $84, and the current $105 if the war persists. We treat $74 as one input, not a "
     "prediction. The EIA raised its fourth quarter 2026 forecast by $14 in a single month [13], so its "
     "forecasts are moving quickly, and the $60 and $84 cases bracket it on both sides. In the war case only, Exxon's cash flow is "
     "reduced by $4.1 billion a year (four times the Q2 loss) because its Middle East volumes stay offline. "
     "The 2025 base already includes those volumes, so the peace cases assume they return.")
para("**Assumptions to keep in view.** Both companies' sensitivities are management estimates. Exxon's is "
     "an upstream figure, and we treat it as a cash flow figure. The model gives neither company credit for "
     "growth after 2025 other than Hess. That understates both, and it understates Chevron more, because its "
     "Tengiz expansion ramped up during 2025.")

# ---------------------------------------------------------------- 5
h1("5 Results")
rows = []
for B, nm in [(60, "Glut, $60"), (74, "Peace, $74"), (84, "EIA 2027, $84"), (105, "War, $105")]:
    rows.append([nm, f"${xr[B]['fcf']:.1f}B", f"${xr[B]['resid']:.1f}B", f"${cr[B]['fcf']:.1f}B", f"${cr[B]['resid']:.1f}B"])
table("Table 2", "Annual free cash flow and cash left after dividends, before buybacks. Author's model.",
      ["Scenario", "XOM FCF", "XOM after div.", "CVX FCF", "CVX after div."], rows, [1.3, 1.0, 1.1, 1.0, 1.1])
figure("fig1_fcf_yield.png", "Figure 1",
       "Free cash flow yield (FCF divided by market value) under four Brent scenarios, with each stock's "
       "dividend yield. Author's model; inputs from Table 1.")
para(f"Figure 1 and Table 2 show the core trade-off. At $74 Chevron yields {cr[74]['fcf_yld']:.1f}% against "
     f"Exxon's {xr[74]['fcf_yld']:.1f}%, and its lead widens as oil rises: {cr[84]['fcf_yld']:.1f}% versus "
     f"{xr[84]['fcf_yld']:.1f}% at $84 and {cr[105]['fcf_yld']:.1f}% versus {xr[105]['fcf_yld']:.1f}% if the war "
     f"persists. Exxon wins only in the $60 glut ({xr[60]['fcf_yld']:.1f}% versus {cr[60]['fcf_yld']:.1f}%). The "
     f"other side of the trade is volatility. Moving from war to peace, Chevron's free cash flow falls from "
     f"${cr[105]['fcf']:.1f} billion to ${cr[74]['fcf']:.1f} billion (down "
     f"{100*(1-cr[74]['fcf']/cr[105]['fcf']):.0f}%), while Exxon's falls from ${xr[105]['fcf']:.1f} billion to "
     f"${xr[74]['fcf']:.1f} billion (down {100*(1-xr[74]['fcf']/xr[105]['fcf']):.0f}%). At $60, Exxon has "
     f"${xr[60]['resid']:.1f} billion left after its dividend, while Chevron roughly breaks even on its "
     f"${co['CVX']['div']:.1f} billion dividend.")
figure("fig2_breakevens.png", "Figure 2",
       "Brent price at which normalized cash flow covers capex and the dividend, and capex, dividend and "
       "buyback, against EIA forecast prices [13]. Author's model.")
para(f"Figure 2 restates the same results as breakevens. Exxon covers capex and its dividend at about "
     f"${X['be_div']:.0f} Brent and its full $20 billion buyback at about ${X['be_bb']:.0f}. Chevron needs "
     f"about ${C['be_div']:.0f} for capex and dividend, and ${C['be_bb']:.0f} to add a buyback of the same 2.8% size. "
     "Both full payout programs need more than the EIA's $74 late 2027 price, so investors should expect "
     "buybacks to shrink or be funded from the balance sheet. Exxon has done the latter before: in 2025 "
     "it distributed $37.2 billion against $26.1 billion of free cash flow [2]. Chevron's management "
     "targets a capex-plus-dividend breakeven below $50 [10]. Our estimate is about $10 higher, mostly "
     "because we exclude growth since 2025. If management is closer to right, Chevron's lead is larger than "
     "Figure 1 shows; if we are, it still leads at every price above about $67.")

# ---------------------------------------------------------------- 6
h1("6 Alternatives considered")
para("**The case for Exxon.** Exxon is the steadier stock. Its cash flow is less sensitive to oil (0.10% of "
     "market value per $1 of Brent versus 0.14% for Chevron), it falls less when the war premium fades, and its "
     "dividend is well covered at $60. Its Middle East volumes act as a hedge, since they return just as prices "
     "fall. An investor who expects oil below about $67 after the war, or who cares most about dividend safety, "
     "should prefer Exxon. We do not, because the EIA path and Chevron's own planning range ($60 to $80 for its "
     "buyback guidance [10]) both sit above that level, and Exxon's lower risk comes with a lower yield.")
para("**Owning neither.** Both stocks have rallied on a windfall the EIA calls temporary, and a peace-case "
     "FCF yield near 5% is not a margin of safety. An investor with no mandate to hold energy could reasonably "
     "wait for a lower entry point. Our recommendation is aimed at the investor who wants the exposure and "
     "has to choose between the two.")

# ---------------------------------------------------------------- 7
h1("7 Risks and what would change the call")
para("**Chevron's growth may disappoint.** The recommendation leans on Chevron's cash flow holding up after "
     "the Hess integration. If its third quarter cash flow, adjusted for Brent, implies a normalized base below "
     f"our ${co['CVX']['base']:.1f} billion, its yield advantage shrinks or disappears. This is the first thing we "
     "would check after the late October reports.")
para("**Exxon's hedge could pay out sooner.** A fast ceasefire would hit Chevron harder while returning Exxon's "
     "Middle East volumes. If those volumes never return, though, Exxon's peace-case yield falls to about 4.5% "
     "and Chevron's lead widens. A full-quarter Hormuz closure would cut Exxon's Middle East output by about "
     "750,000 barrels of oil equivalent a day versus 2025 [5].")
para("**Oil can overshoot downward.** Once the 4.5 million barrels per day of shut-in output returns, a glut "
     "below $60 is plausible. That is the scenario where Exxon wins, and it would hurt both stocks in "
     "absolute terms. Chevron's Tengiz output also depends on the CPC pipeline through the Black Sea, which "
     "has been intermittently disrupted [17].")
para("**Model risk.** The 2025 cash flow figures, the Hess adjustment and the sensitivities are the inputs "
     "the conclusion depends on most. A $1 billion error in a company's base moves its breakevens by about "
     "$1.4 (Exxon) to $1.7 (Chevron) a barrel.")

# ---------------------------------------------------------------- 8
h1("8 Conclusion")
para("At today's prices the market is valuing both majors partly on wartime cash flows. Once those cash flows "
     "are normalized, Chevron offers more free cash flow and a higher dividend per dollar invested, even before "
     "crediting growth our model leaves out. Exxon is the safer stock if oil collapses, and investors who "
     "expect that outcome should own it instead. For the investor who wants integrated oil exposure and "
     "accepts the EIA's path, we prefer Chevron, and we would revisit the call after both companies report "
     "third quarter results. We would not recommend either stock to an investor who does not need energy "
     "exposure until the price of entry reflects peace-case cash flows.")

# ---------------------------------------------------------------- references
doc.add_page_break()
h1("References")
refs = [
 "ExxonMobil Holdings Corp. (2026, July 31). ExxonMobil announces second-quarter 2026 results [Form 8-K, Ex. 99.1]. https://www.sec.gov/Archives/edgar/data/0002115436/000211543626000006/livef8k2q26991.htm",
 "ExxonMobil. (2026, January 30). ExxonMobil announces 2025 results. https://corporate.exxonmobil.com/news/news-releases/2026/0130-exxonmobil-announces-2025-results",
 "ExxonMobil. (2026). 2Q 2026 sensitivities and key assumptions. https://d1io3yog0oux5.cloudfront.net/_9c0226c1dc3efd2dbe7760ece1f33886/exxonmobil/db/2431/22314/file/Sensitivities+and+Key+Assumptions.pdf",
 "Exxon Mobil Corp. (2026, April 8). Form 8-K, first-quarter 2026 earnings considerations. https://www.sec.gov/Archives/edgar/data/34088/000003408826000056/f8k1q992040826.htm",
 "ExxonMobil Holdings Corp. (2026, July 31). 2Q 2026 earnings call presentation slides (slides 9, 13, 17 and sensitivities page). https://investor.exxonmobil.com",
 "Chevron Corp. (2026, July 31). Chevron reports second quarter 2026 results [Form 8-K, Ex. 99.1]. https://www.sec.gov/Archives/edgar/data/0000093410/000009341026000162/a06302026ex9918-k.htm",
 "Chevron Corp. (2026, January). Chevron reports fourth quarter 2025 results [Form 8-K, Ex. 99.1]. https://www.sec.gov/Archives/edgar/data/93410/000009341026000019/a12312025ex9918-k.htm",
 "Chevron Corp. (2026). Form 10-Q for the quarter ended June 30, 2026. https://www.sec.gov/Archives/edgar/data/0000093410/000009341026000167/cvx-20260630.htm",
 "Chevron Corp. (2026). Form 10-Q for the quarter ended March 31, 2026. https://www.sec.gov/Archives/edgar/data/0000093410/000009341026000113/cvx-20260331.htm",
 "Chevron Corp. (2025, November 12). 2025 Investor Day, edited transcript. https://chevroncorp.gcs-web.com/static-files/f8964224-9b9a-4be6-8644-9f96bd1e1fe4",
 "Chevron Corp. (2026, August 7). Chevron 2026 investor presentation. https://chevroncorp.gcs-web.com/static-files/6fe32bf1-951a-44c8-81ee-74cd626c0095",
 "Hess Corp. (2025). Form 10-K for fiscal year 2024. https://www.sec.gov/Archives/edgar/data/4447/000162828025008716/hes-20241231.htm",
 "U.S. Energy Information Administration. (2026, October 6). Short-Term Energy Outlook, October 2026. https://www.eia.gov/outlooks/steo/archives/oct26.pdf",
 "APA. (2026). World Bank releases average oil price data for Brent and WTI. https://en.apa.az/energy-and-industry/world-bank-releases-average-oil-price-data-for-brent-and-wti-519072",
 "Statista. (2026). Brent crude oil price annually 1976–2025. https://www.statista.com/statistics/262860/weekly-crude-oil-prices",
 "Morningstar. (2026, October 13). XOM and CVX stock quotes. https://www.morningstar.com/stocks/xnys/xom/quote ; https://www.morningstar.com/stocks/xnys/cvx/quote",
 "The Motley Fool. (2026, August 3). Chevron (CVX) Q2 2026 earnings call transcript. https://www.fool.com/earnings/call-transcripts/2026/08/03/chevron-cvx-q2-2026-earnings-call-transcript/",
 "World Oil. (2025, July 18). Chevron completes $53 billion Hess takeover following lengthy Exxon arbitration. https://worldoil.com/news/2025/7/18/chevron-completes-53-billion-hess-takeover-following-lengthy-exxon-arbitration/",
 "24/7 Wall St. (2026, July 23). At $100 per barrel, which oil stock has dominated in 2026; The Motley Fool. (2026, July 30). These 3 oil stocks surged in the first half of 2026. https://247wallst.com/investing/2026/07/23/at-100-per-barrel-which-oil-stock-has-dominated-in-2026-exxonmobil-chevron-or-bp/",
 "Politis. (2026). Tracking oil prices since the beginning of the war 100 days ago. https://en.politis.com.cy/news/world/1011001/tracking-oil-prices-since-the-beginning-of-the-war-100-days-ago",
 "Trading Economics. (2026, October 13). Brent crude oil. https://tradingeconomics.com/commodity/brent-crude-oil",
]
for i, r_ in enumerate(refs, 1):
    p = doc.add_paragraph(); p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Inches(0.3); p.paragraph_format.first_line_indent = Inches(-0.3)
    p.paragraph_format.space_after = Pt(4); p.paragraph_format.line_spacing = 1.1
    rr = p.add_run(f"[{i}] {r_}"); rr.font.size = Pt(9)

doc.save("Investor_Analyst_Report.docx")
print("saved")
