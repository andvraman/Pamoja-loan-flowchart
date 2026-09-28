# PAMOJA simplified financial model (v0.4, draft)

`PAMOJA_Financial_Model_v0.4.xlsx` shows how the $24m fund flows from TAMISEMI to lenders, on to CMGs as loans, and back as repayments.
Yellow cells are inputs; everything else is a live formula. **All district data and rates are dummy placeholders.**

| Sheet | What it does |
|---|---|
| Read me | How to use it, the logic, what changed, and what is not yet modelled |
| Inputs | Fund size, incentives, Loan:Savings cap, interest, tenure/term, default rate, lending timing, split of interest, **principal settlement and risk share of defaults** |
| Districts | Per LGA: number of CMGs, savings, loan asked → capped loan, demand, allocation (scaled if demand > corpus) |
| Monthly cycle | 24 months (year 1 lending, year 2 run-off): new loans, repayments, **TAMISEMI's share of repaid principal, lender's risk-share compensation**, cash, loans out; year-end snapshot |
| Fund flow | Totals for all loans made in year 1 until they close, split of interest, **risk share of defaults**, corpus check, bar chart of money not lent / repaid / interest |
| Open questions | Things to settle (also in `OPEN_QUESTIONS.md`) |
| Glossary | Acronyms and terms |

## v0.4: risk share and principal settlement

- **Region share** and **LGA share** of interest are now labelled **Region service fee** and **LGA service fee**.
- New: **TAMISEMI share of repaid principal** (default 100%) — the share of every repaid instalment's
  principal that returns to the TAMISEMI corpus. The rest is kept by the lender (a computed remainder).
- New: **Risk share of defaults** — how the loss on a defaulted loan splits between TAMISEMI and the lender.
  TAMISEMI absorbs its share against the corpus (input, default 50%); the lender absorbs the rest (a computed
  remainder) as its own loss, and pays that amount back into the fund as risk-share compensation.
- Fund flow now shows both settlements and a **Lender's net result** line (margin − its own risk share of
  defaults + any principal it keeps).
- The break-even default rate and the default budget needed now only have to cover TAMISEMI's risk share of
  defaults, not the whole default — so a 50/50 risk share roughly doubles how much default the same interest
  split can absorb.
- **Both are placeholders** — see question 12 in `OPEN_QUESTIONS.md`.

## The GUI: PAMOJA Model Inputs

Every independent input (everything yellow on the Inputs and Districts sheets, including the new risk-share
fields) can be set from a browser page instead of the spreadsheet:
**[PAMOJA Model Inputs](https://claude.ai/artifact/HcwpiFMz2g1dmxUu5ueG2g)**. It shows a live preview (corpus,
district allocation, a bar chart, break-even default rate, lender's net result, whether the corpus is
restored) as you change values.

It saves your numbers into the artifact's own small database (not a download — downloads don't work inside
these pages). Ask Claude to "pull the latest inputs and rebuild the model" and it will fetch that saved data
and regenerate the workbook. "Copy config as text" on the page is a manual fallback.

To rebuild by hand from a saved config file:
```
pip install openpyxl
python3 model/build_model.py [config.json]   # config.json is optional; without it, uses built-in defaults
```
See `DEFAULT_INPUTS` / `DEFAULT_DISTRICTS` in `build_model.py` for every key.

## Earlier changes
- **v0.3**: config-file-driven rebuild; Fund flow bar chart.
- **v0.2**: lender discount removed (lender keeps the margin); lending across the cycle up to 12-month tenure;
  optional re-lending; 24-month monthly sheet with a year-end snapshot.
