"""Builds Analyst_Reflection.docx in JDF 2.2 format, reusing the report's setup and helpers.
Text marked [[FILL ...]] must be completed from the student's own spot-check results before submission."""
src = open("build_report.py").read()
exec(src.split("# ---------------------------------------------------------------- title block")[0])

FILL = "[[FILL: "


def fill(text):
    p = doc.add_paragraph()
    rr = p.add_run(FILL + text + "]]"); rr.bold = True; rr.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)


p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.line_spacing = 1.15
rr = p.add_run("Analyst Reflection: Chevron Over ExxonMobil"); rr.font.size = Pt(17)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Companion to the Investor / Analyst Report · Written October 2026")

h1("1 How I used AI")
para("I used Claude (Anthropic), a general purpose AI assistant, not a finance-specific product. It suggested "
     "Exxon versus Chevron as a scope I could handle, searched the web for current data, wrote and ran the "
     "Python model and charts, and drafted most of the text. My job was making the calls where the analysis "
     "could go different ways, checking the numbers the conclusion depends on, and deciding what I was actually "
     "willing to recommend.")

h1("2 Where I pushed back")
para("**The baseline.** The AI's first pass worked off the latest quarter. When we tested it, Q2 2026 cash "
     "flow was about $19 billion (Exxon) and $23 billion (Chevron) a year higher then oil prices alone could "
     "explain. I didn't want to just trust company guidance either, because Chevron's claim that it breaks even "
     "below $50 is exactly the thing the report should be testing. So I went with 2025 results plus an "
     "adjustment for the Hess deal.")
para("**The buyback chart.** After the first draft I had the AI list the weakest parts of its own work and I "
     "decided each one. The big one was Figure 2. It compared Exxon with a $20 billion buyback against Chevron "
     "with a $10 billion one, but that's 2.8% of Exxon's value versus 2.3% of Chevron's, so it wasn't a fair "
     "comparison. I had both tested at 2.8%, which raised Chevron's breakeven by about $3. I also didn't "
     "want one EIA number treated as “the” peace price when the EIA moved its own forecast $14 in a "
     "month, so the report now calls $74 one input. The AI had just assumed a Hess oil sensitivity with no "
     "source. I kept it but made the report say so and show it only moves the breakevens about $1.")

h1("3 How I improved the work")
para("The question itself changed. I started out with a basic “which oil major is better” idea, "
     "then the first data showed Brent averaging $114 in September because of the war, with the EIA expecting "
     "$74 by late 2027. That turned it into: which company still earns its share price once the war premium "
     "is gone? I also replaced annualized quarterly cash flow with a normalized model, added the missing "
     "months of Hess for Chevron, modeled Exxon's Middle East output coming back in peace (and tested the "
     "case where it doesn't), and fixed a label that called Exxon's sensitivity an after-tax cash flow figure "
     "when it's source didn't say that.")

h1("4 What I verified")
para("Verification mattered alot here because the AI couldn't open the SEC, EIA or company websites from "
     "its environment. Everything came through search result summaries, and they didn't always agree. "
     "Chevron's 2025 cash flow showed up as both $33.9 and $34.9 billion, and its market value ranged from "
     "$405 to $427 billion depending on the site. So I downloaded the original documents myself: both "
     "companies' fourth quarter 2025 releases, Chevron's Q2 2026 release and 10-Q, and Exxon's Q2 2026 slides. "
     "The AI pulled the tables out and compared them to the draft, and I read Chevron's sensitivity sentence "
     "in the 10-Q myself ($600 million after tax per $1 of Brent, which matched). Exxon's 2025 cash flow, free "
     "cash flow and payouts all matched exactly.")
para("**Two checks failed.** The draft said Exxon lost $1.3 billion to Middle East volumes in Q2, but the "
     "slides show $1.28 billion for the whole first half, so Q2 was closer to $1.03 billion. A news summary had "
     "labeled a six-month number as one quarter. The bigger one was Chevron's 2025 cash flow excluding working "
     "capital. The draft used $33.9 billion, but Chevron's release says $34.9 billion. The $33.9 was the line "
     "right above it, cash flow including working capital. That $1 billion is what changed my recommendation "
     "(Section 6). I never checked the October 13 share prices, so those are still from secondary sources.")

h1("5 The AI tool's weaknesses")
para("**Outdated knowledge.** When we picked the topic the AI called it a comparison with “one dominant "
     "driver (oil price)” and had no idea a war had pushed oil over $100. I only noticed when the first "
     "search came back with $114 Brent. After that, every current fact had to come from a dated 2026 source.")
para("**No primary sources.** Since it couldn't open filings, it worked from summaries that sometimes grabbed "
     "the wrong table row or relabeled time periods, and it still built clean looking tables out of them. Both "
     "of my failed checks came from this. The fix was getting the documents myself, which worked for the "
     "numbers I checked but not for the ones I didn't.")
para("**Treating one quarter as normal.** Its first instinct was to annualize Q2, which would have overstated "
     "normal cash flow by around a quarter. **False precision** was the other one. It reports breakevens to "
     "the dollar even though some inputs are guesses, so the report now says how much each error moves them.")

h1("6 Where my thinking changed")
para("My recommendation changed twice. On headline numbers Chevron looked a lot cheaper. After normalizing, "
     "the draft showed the two tied at $74, and I leaned Exxon since it falls less when oil drops and gets "
     "it's Middle East barrels back in a peace scenario. Then my spot-check found Chevron's 2025 cash flow was "
     "$1 billion too low in the draft. With the right number Chevron yields more in three of four scenarios, "
     "even without credit for Tengiz, so I switched to Chevron. The lesson for me was about conviction. If one "
     "corrected number can flip the call, it's a narrow call, and the report should say that, which it now does.")

h1("7 What remains uncertain")
para("I'm least confident in the 2025 baseline. It treats 2025 as normal and ignores growth since then, apart "
     "from Hess, which hurts Chevron the most because Tengiz was ramping up during 2025. That's probably why my "
     "Chevron breakeven ($60) is about $10 above what management claims. If they're right, my call gets "
     "stronger. If Chevron's normal cash flow is below my $37.8 billion, the lead goes away. Next I would take "
     "Chevron's Q3 results at the end of October, back out the oil price effect, and see where it lands.")

h1("8 Would I stand behind this report?")
para("Yes, as a relative call, with moderate conviction. I'd defend “prefer Chevron over Exxon for an "
     "investor who wants oil exposure and accepts the EIA's price path,” because Chevron comes out ahead "
     "in three of four scenarios and pays a higher dividend under a model that is definately conservative "
     "toward it. But I would tell the committee the lead is narrow, that Exxon is safer if oil goes below "
     "about $67, and that the call already flipped once on one corrected number. I wouldn't sign a general "
     "“buy energy now” call, and before putting real money in I'd want the Q3 results and every "
     "input in Table 1 checked against the filings.")

doc.save("Analyst_Reflection.docx")
print("saved")
