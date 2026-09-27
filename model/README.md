# PAMOJA simplified financial model (v0.2, draft)

`PAMOJA_Financial_Model_v0.2.xlsx` shows how the $24m fund flows from TAMISEMI to lenders, on to CMGs as loans, and back as repayments.
Yellow cells are inputs; everything else is a live formula. **All district data and rates are dummy placeholders.**

| Sheet | What it does |
|---|---|
| Read me | How to use it, the logic, what changed, and what is not yet modelled |
| Inputs | Fund size, $2m + $2m incentives set aside, Loan:Savings cap, interest, max tenure (12 months) and term, default rate, last lending month, re-lending %, split of interest |
| Districts | Per LGA: number of CMGs, savings, loan asked → capped loan, demand, allocation (scaled if demand > corpus) |
| Monthly cycle | 24 months (year 1 lending, year 2 run-off): new loans, re-lent repayments, repayments, cash, loans out; year-end snapshot |
| Fund flow | Totals for all loans made in year 1 until they close, split of interest, corpus check |
| Open questions | Things to settle (also in `OPEN_QUESTIONS.md`) |
| Glossary | Acronyms and terms |

## Changes in v0.2
- Lender discount removed; the lender keeps the margin (interest left after the other shares).
- Lenders can make new loans in any month up to a chosen last month; tenure up to 12 months.
- Optional re-lending of repayments.
- Monthly sheet runs 24 months, with a year-end (month 12) snapshot of cash vs money still on loan.

To rebuild after editing the script: `pip install openpyxl && python3 model/build_model.py`
