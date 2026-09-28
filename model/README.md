# PAMOJA simplified financial model (v0.3, draft)

`PAMOJA_Financial_Model_v0.3.xlsx` shows how the $24m fund flows from TAMISEMI to lenders, on to CMGs as loans, and back as repayments.
Yellow cells are inputs; everything else is a live formula. **All district data and rates are dummy placeholders.**

| Sheet | What it does |
|---|---|
| Read me | How to use it, the logic, what changed, and what is not yet modelled |
| Inputs | Fund size, $2m + $2m incentives set aside, Loan:Savings cap, interest, max tenure (12 months) and term, default rate, last lending month, re-lending %, split of interest |
| Districts | Per LGA: number of CMGs, savings, loan asked → capped loan, demand, allocation (scaled if demand > corpus) |
| Monthly cycle | 24 months (year 1 lending, year 2 run-off): new loans, re-lent repayments, repayments, cash, loans out; year-end snapshot |
| Fund flow | Totals for all loans made in year 1 until they close, split of interest, corpus check, **bar chart of money not lent / repaid / interest** |
| Open questions | Things to settle (also in `OPEN_QUESTIONS.md`) |
| Glossary | Acronyms and terms |

## v0.3: a GUI instead of typing into Excel

Every independent input (everything yellow on the Inputs and Districts sheets) can now be set from a browser
page instead of the spreadsheet: **[PAMOJA Model Inputs](https://claude.ai/artifact/HcwpiFMz2g1dmxUu5ueG2g)**.
It shows a live preview (corpus, district allocation, a bar chart, break-even default rate) as you change values.

How it feeds the Excel file: the page saves your numbers into the artifact's own small database (not a
download — downloads don't work inside these pages). Ask Claude to "pull the latest inputs and rebuild the
model" and it will fetch that saved data and regenerate the workbook from it. You can also click
"Copy config as text" on the page and paste that into the chat as a fallback.

To rebuild by hand from a saved config file:
```
pip install openpyxl
python3 model/build_model.py [config.json]   # config.json is optional; without it, uses built-in defaults
```
`config.json` shape:
```json
{"inputs": {"fund": 24000000, "ratio": 0.5, "...": "..."},
 "districts": [{"name": "District A", "cmgs": 2500, "savings": 4000, "loan_ask": 2500}]}
```
See `DEFAULT_INPUTS` / `DEFAULT_DISTRICTS` in `build_model.py` for every key.

## Earlier changes (v0.2)
- Lender discount removed; the lender keeps the margin (interest left after the other shares).
- Lenders can make new loans in any month up to a chosen last month; tenure up to 12 months.
- Optional re-lending of repayments.
- Monthly sheet runs 24 months, with a year-end (month 12) snapshot of cash vs money still on loan.
