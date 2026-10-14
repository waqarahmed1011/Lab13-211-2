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
para("**Choosing the baseline.** The most important decision in the report was what counts as a company's "
     "normal cash flow. The tool's first pass worked from the most recent quarter. When we tested that "
     "against the companies' own oil price sensitivities, Q2 2026 cash flow turned out to be about $19 billion "
     "(Exxon) and $24 billion (Chevron) higher per year than oil prices alone explain. The AI then laid out "
     "three options: annualize Q2, start from 2025 results with an adjustment for Chevron's Hess acquisition, or "
     "rely on company guidance such as Chevron's claim that it breaks even below $50 Brent. I rejected "
     "company guidance because that sub-$50 claim is exactly what the report needed to test, not assume. I "
     "also rejected the Q2 run rate because it treats wartime refining margins as permanent. I chose the 2025 "
     "base plus Hess. The result was that Chevron's breakeven came out near $62, well above its own claim, "
     "and the report now names that gap as its largest uncertainty instead of hiding it.")
fill("Add your second pushback example from the spot-checks, in your own words: which number you checked, "
     "what the AI or search summary said, what the source actually said, and what changed in the report. If "
     "every check matched, say that and describe instead a claim in the draft you questioned and why.")

h1("3 How I improved the work")
para("**The question changed.** I started with a generic “which major is better” comparison. The first "
     "data pulls showed Brent averaging $114 in September because of the Iran conflict, and the EIA forecasting "
     "a fall to $74 by late 2027. That turned the question into which company still earns its share price "
     "after the war premium fades.")
para("**A better metric replaced a misleading one.** Annualized quarterly cash flow was replaced by a "
     "normalized model: 2025 cash flow moved to each scenario price using each company's disclosed "
     "sensitivity. With that change, Chevron's apparent valuation discount disappeared. Both stocks yield "
     "about 4.9% at $74.")
para("**A fairness adjustment and a stress test were added.** Chevron's 2025 numbers include only about five "
     "months of Hess, so I added the missing portion (+$2.9 billion) using Hess's last annual report. I also "
     "chose to model Exxon's Middle East production returning in a peace scenario, which became the "
     "report's most distinctive argument. Then I tested the case where it does not return: Exxon's peace-case "
     "yield drops to about 4.4%, below Chevron's, and most of the preference disappears. That test is in the "
     "risk section.")
para("**A labeling error was corrected.** The draft called Exxon's sensitivity an after-tax cash flow "
     "figure; the source only supports an upstream earnings sensitivity.")

h1("4 What I verified")
para("Verification mattered more than usual here, because the AI tool could not open SEC filings, EIA "
     "pages or company websites from its environment. Every figure reached it through search engine summaries, "
     "and those summaries sometimes disagreed. For example, Chevron's 2025 operating cash flow came back as both "
     "$33.9 billion and $34.9 billion. Chevron's Q2 free cash flow appeared as both $15.4 billion "
     "(company-adjusted) and $18.1 billion, and its market value ranged from $405 billion to $427 billion "
     "depending on the site. Q2 Brent was $104.40 (World Bank) or $102.93 (EIA), too small a gap "
     "to matter.")
para("Every derived number was recomputed in code rather than typed by hand: market values from share "
     "counts implied by each company's reported EPS, free cash flow, and breakevens. I then checked the "
     "inputs the conclusion depends on most against the original documents:")
fill("List each spot-check you did (1 to 6 from SPOT_CHECKS.md): the number, the source you opened, and "
     "whether it matched. Give at least one in full detail: why it mattered, how you checked it, and what "
     "you found.")

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
