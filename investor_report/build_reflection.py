"""Builds Analyst_Reflection.docx in JDF 2.2 format, reusing the report's setup and helpers.
Text marked [[FILL ...]] must be completed from the student's own spot-check results before submission."""
src = open("build_report.py").read()
exec(src.split("# ---------------------------------------------------------------- title block")[0])

FILL = "[[FILL: "


def fill(text):
    p = doc.add_paragraph()
    rr = p.add_run(FILL + text + "]]"); rr.bold = True; rr.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)


p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.line_spacing = 1.15
rr = p.add_run("Analyst Reflection: ExxonMobil or Chevron When the War Premium Fades"); rr.font.size = Pt(17)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Companion to the Investor / Analyst Report · Written October 2026")

h1("1 How I used AI")
para("I used Claude (Anthropic), a general-purpose assistant, running in a coding environment with web search "
     "and Python. It was not a finance-specific product. It suggested the ExxonMobil versus Chevron "
     "scope, searched for current data (its built-in knowledge stopped before the period I analyzed), wrote and "
     "ran the Python scenario model and charts, and drafted most of the report text. I want to be plain about that last point: the AI wrote most of the prose. My role was "
     "deciding at the points where the analysis could go different ways, checking the numbers the "
     "conclusion depends on, and choosing the recommendation I was willing to sign.")

h1("2 Where I pushed back")
para("**Choosing the baseline.** The tool's first pass worked from the latest quarter, but Q2 2026 cash flow "
     "was about $19 billion (Exxon) and $24 billion (Chevron) a year higher than oil prices alone explain. Of "
     "the options it laid out, I rejected company guidance because Chevron's sub-$50 breakeven claim is what "
     "the report needed to test, not assume, and I rejected the Q2 run rate because it treats wartime "
     "refining margins as permanent. I chose 2025 results plus a Hess adjustment.")
para("**Challenging the draft.** After the first draft I had the AI list the weakest points it could see, and "
     "I made the call on each. The most important was Figure 2. It compared Exxon's breakeven with a $20 "
     "billion buyback against Chevron's with a $10 billion buyback, but those are 2.8% and 2.3% of market "
     "value. The smaller Chevron program made the two companies look closer than they are. I had the model "
     "test both at the same 2.8%, and Chevron's breakeven rose from $79 to $82, against $77 for Exxon. I also "
     "would not let one EIA number stand as \u201cthe\u201d peace price, because the EIA raised its fourth "
     "quarter forecast by $14 in a single month. The report now calls $74 one input and leans on the "
     "bracketing scenarios. The Hess adjustment used a $0.10 billion per $1 sensitivity the AI had simply "
     "assumed. I kept it but required the report to flag it and show that zero or double moves Chevron's "
     "breakevens by only about $1. Finally, I kept Exxon's Middle East recovery central despite the damaged "
     "Qatar trains, because peace implies an eventual restart and the risk section already shows what "
     "happens if it never comes (Exxon's yield at $74 falls to 4.5%).")

h1("3 How I improved the work")
para("**The question changed.** I started with a generic “which major is better” comparison. The first "
     "data pulls showed Brent averaging $114 in September because of the Iran conflict, and the EIA forecasting "
     "a fall to $74 by late 2027. That turned the question into which company still earns its share price "
     "after the war premium fades.")
para("**A better metric replaced a misleading one.** Annualized quarterly cash flow became a normalized "
     "model, and Chevron's apparent valuation discount disappeared: both stocks yield about 4.9% at $74.")
para("**Fairness fixes.** I added the missing months of Hess for Chevron (+$2.9 billion), modeled Exxon's "
     "Middle East output returning in peace, and stress-tested the case where it does not. I also corrected "
     "a label that called Exxon's sensitivity an after-tax cash flow figure, which its source does not support.")

h1("4 What I verified")
para("Verification mattered more than usual here, because the AI tool could not open SEC filings, EIA "
     "pages or company websites from its environment. Every figure reached it through search engine summaries, "
     "and those summaries sometimes disagreed. For example, Chevron's 2025 operating cash flow came back as both "
     "$33.9 billion and $34.9 billion. Chevron's Q2 free cash flow appeared as both $15.4 billion "
     "(company-adjusted) and $18.1 billion, and its market value ranged from $405 billion to $427 billion "
     "depending on the site. Q2 Brent was $104.40 (World Bank) or $102.93 (EIA), too small a gap "
     "to matter.")
para("Every derived number was recomputed in code rather than typed by hand. I then pulled the original "
     "documents myself: Chevron's Q2 2026 release and 10-Q, Exxon's 2025 results release, and Exxon's Q2 2026 "
     "earnings slides. I had the AI extract the tables from them and compare them with the draft. I read the "
     "Chevron sensitivity sentence in the 10-Q myself: about $600 million of after-tax earnings and cash "
     "flow per $1 of Brent, which matched. Exxon's 2025 cash flow ($59.7 billion), free cash flow ($26.1 "
     "billion) and distributions ($37.2 billion) all matched to the decimal, and its sensitivity page "
     "confirmed $700 million per $1 of Brent. That page also showed $800 million per $1 of refining margin, "
     "which I added to the report because it explains much of the Q2 cash surge.")
para("**The check that failed.** The draft said Exxon lost about $1.3 billion of upstream earnings to "
     "Middle East volumes in Q2, a figure from a news summary. Exxon's slides show $1.28 billion for the "
     "whole first half, with Q2 about $0.78 billion worse than Q1, which puts Q2 alone near $1.03 billion. "
     "The summary had relabeled a six-month number as a quarter. That number drives the war scenario, so "
     "the war-case deduction fell from $5.2 billion to $4.1 billion a year. The recommendation did not "
     "change. I could not verify Chevron's full-year 2025 cash flow, because the release I found covered Q2 "
     "2026, and I did not check the October 13 share prices, so both remain secondary-source figures.")

h1("5 The AI tool's weaknesses")
para("**Out-of-date knowledge presented confidently.** When I picked the topic, the tool described the "
     "comparison as having “one dominant driver (oil price)” and did not know a war had pushed Brent "
     "above $100. I noticed this when the first search showed September Brent at $114. It mattered because "
     "the original framing would have produced a stale report. The fix was a rule that every current fact "
     "must come from a dated 2026 source. That worked for facts, but the tool's framing still reflected its "
     "older picture of the industry until the data forced a rethink.")
para("**No access to primary sources.** The tool's environment blocked the SEC, EIA and company sites, so it "
     "worked from search summaries that conflicted, mixed up table columns, and sometimes cited secondary "
     "blogs. It flagged some conflicts but still built clean-looking tables from them. "
     "I gave priority to company-reported definitions, derived market values myself, and checked the key "
     "inputs by hand. That reduced the risk without removing it.")
para("**Treating a recent quarter as representative.** Its first instinct was to annualize Q2 2026, "
     "overstating normal cash flow by roughly a quarter (Section 2). The normalized model fixed this but "
     "swapped in a different assumption (Section 7).")
para("**False precision.** The model reports breakevens like “$62” to the dollar, but Chevron's "
     "figure depends partly on a Hess oil price sensitivity of $0.10 billion per $1 that the tool assumed and "
     "could not source. The report says a $1 billion error moves a breakeven by $1.40 to $1.70, so readers "
     "can see how soft those figures are.")

h1("6 Where my thinking changed")
para("I expected this to be a valuation contest that the cheaper stock would win. On headline numbers "
     "Chevron looked cheaper: a higher dividend yield, a slightly lower forward P/E, and more Q2 cash flow per "
     "dollar of market value. Normalizing the cash flows removed that discount. The deciding factor turned out "
     "to be one I had not thought about at all: Exxon's Middle East exposure. During the war it is a "
     "liability, since about 10% of upstream output was offline. In a peace scenario it works as a hedge, "
     "because production returns just as prices fall. I also came to see the recommendation as a relative call "
     "rather than a reason to buy energy.")

h1("7 What remains uncertain")
para("I have the least confidence in the 2025 baseline. It treats 2025 as normal and ignores both companies' "
     "growth since then, apart from Hess. That penalizes Chevron in particular, because its Tengiz expansion "
     "was ramping up during 2025. It is the most likely reason my Chevron breakeven ($62) is so far above "
     "management's sub-$50 claim. If the real figure is closer to $50, Chevron becomes the better value and "
     "my recommendation flips. The next step is the late October third quarter reports. I would take "
     "Chevron's Q3 operating cash flow, back out the oil price effect using its $0.60 billion sensitivity, "
     "and see whether the implied normal level is well above my $36.8 billion. I would also look for refining "
     "margin data, to see how much of the 2026 windfall came from refining rather than crude.")

h1("8 Would I stand behind this report?")
para("Yes, as a relative call. I would defend “prefer Exxon over Chevron for an investor who wants "
     "integrated oil exposure while the war premium fades,” because it rests on findings that do not "
     "depend on a precise oil forecast: Exxon's lower sensitivity per dollar of market value, its ability to "
     "cover the dividend at $60, and its Middle East volumes coming back. I would not put my name on a broader "
     "“buy energy now” call. Both stocks have rallied on cash flows the EIA expects to be temporary, "
     "and at $74 oil neither generates enough free cash flow to fund its current buyback. Before committing "
     "real capital I would also want the third quarter results in hand and every input in Table 1 checked "
     "against the original filing, not just the ones I spot-checked.")

doc.save("Analyst_Reflection.docx")
print("saved")
