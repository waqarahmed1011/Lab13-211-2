# Spot-checks, step by step (about 20 to 25 minutes)

**Setup (1 min):** open a notes app. For each check, write one line:
`#. what the report says | what the source says | match / mismatch`
Use a laptop if you can; finding text in PDFs on a phone is painful.

**General trick:** once a page or PDF is open, press **Ctrl+F** (Cmd+F on Mac) and type the search term I give you. Don't read the whole document.

---

## 1. Chevron 2025 cash flow excluding working capital (report says $33.9B). Most important check.
1. Google: `Chevron Reports Fourth Quarter 2025 Results`
2. Open the result from **chevron.com**, **businesswire.com** or **sec.gov** (the release is dated late January 2026).
3. Ctrl+F: `excluding working capital`
4. You'll find a row of numbers. The columns are usually ordered: 4Q 2025, 3Q 2025, 4Q 2024, then **Year 2025**, Year 2024. Take the **full-year 2025** figure. Check the column headers above the row to be sure.
5. Write it down. If it says 34.9 (or anything other than 33.9), that's a real mismatch. Tell me.

## 2. Exxon 2025 cash flow and payouts (report says CFO ex-WC $59.7B, FCF $26.1B, distributions $37.2B = $17.2B dividends + $20.0B buybacks)
1. Google: `ExxonMobil Announces 2025 Results`
2. Open the **corporate.exxonmobil.com** link (Jan 30, 2026).
3. Ctrl+F `free cash flow`, then check the first number is 26.1.
4. Ctrl+F `distributions`, then check 37.2, 17.2 and 20.0.
5. Ctrl+F `excluding working capital`. It's in a table near the bottom; take the full-year 2025 number (should be 59.7 billion, possibly shown as 59,698 million).

## 3. Exxon oil price sensitivity (report says ~$0.70B per $1 Brent, upstream)
1. Google: `ExxonMobil 2Q 2026 sensitivities and key assumptions`
2. Or go to **investor.exxonmobil.com**, then the quarterly earnings section, then **2Q 2026**, then the PDF named "Sensitivities and Key Assumptions".
3. Find the Upstream row for Brent / crude price (usually the first row).
4. Write down three things: the dollar figure, whether the page says **"after-tax"**, and whether it's **annual**.

## 4. Chevron oil price sensitivity (report says ~$600M per $1 Brent, after-tax earnings and cash flow)
1. Google: `Chevron 10-Q June 30 2026` and open the **sec.gov** link (filed around early August 2026).
   - If that fails: go to sec.gov, search "Chevron", click **Filings**, and pick the 10-Q dated August 2026.
2. Ctrl+F: `every dollar` (or `sensitivity`).
3. Write down the exact dollar amount and whether it says after-tax.

## 5. Share prices on Oct 13, 2026 (report says XOM $173.24, CVX $217.73)
1. Go to `finance.yahoo.com/quote/XOM/history`
2. Find the row for **Oct 13, 2026**. Use the **Close** column, not "Adj Close".
3. Repeat with `finance.yahoo.com/quote/CVX/history`.
4. CVX is the more likely mismatch: search results ranged from $211.98 to $217.73.

## 6. Exxon's Q2 Middle East hit (report says ~$1.3B to upstream earnings in Q2)
1. Google: `ExxonMobil second quarter 2026 earnings presentation` or go to investor.exxonmobil.com, then **2Q 2026**, then **Earnings slides** (PDF).
2. Ctrl+F: `Middle East`
3. Look for the upstream earnings "walk" or bar chart (earnings going from Q1 to Q2, broken into factors). Find the Middle East bar and write down its number. If there's no slide, Ctrl+F the press release from check 2's site instead.

---

## When you're done
Paste all six lines back to me. For example:
```
1. 33.9 | release says 34.9 | MISMATCH
2. all match
3. 700M | says "approx $700M, annual", doesn't say after-tax | partial
...
```
I'll rerun the model with the corrected numbers, update both documents, and write the verification section from what you found.

**If a page won't load or you can't find a number in 3 minutes, skip it and write "couldn't find."** That's an honest result too, and it belongs in the reflection.
