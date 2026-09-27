# PAMOJA simplified financial model (v0.1, draft)

`PAMOJA_Financial_Model_v0.1.xlsx` shows how the $24m fund flows from TAMISEMI to lenders, on to CMGs as loans, and back as repayments.
Yellow cells are inputs; everything else is a live formula. **All district data and rates are dummy placeholders.**

| Sheet | What it does |
|---|---|
| Read me | How to use it, the logic, and what is not yet modelled |
| Inputs | Fund size, $2m + $2m incentives set aside, Loan:Savings cap, interest, term, default rate, split of interest |
| Districts | Per LGA: number of CMGs, savings, loan asked → capped loan, demand, allocation (scaled if demand > corpus) |
| Fund flow | One annual cycle as a waterfall, plus a check that the corpus is restored |
| Monthly cycle | 12-month timing of disbursements and repayments; warns if loans run past month 12 |
| Glossary | Acronyms and terms |

To rebuild after editing the script: `pip install openpyxl && python3 model/build_model.py`
