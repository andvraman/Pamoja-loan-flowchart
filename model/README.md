# PAMOJA simplified financial model (v0.5, draft)

`PAMOJA_Financial_Model_v0.5.xlsx` has two parts. Yellow cells are inputs; everything else is a live formula.
**All district data and rates are dummy placeholders.**

## Part 1: the annual $24m model
| Sheet | What it does |
|---|---|
| Read me | How to use it, the logic, what changed, and what is not yet modelled |
| Inputs | Fund size, incentives, Loan:Savings cap, interest, tenure/term, default rate, lending timing, split of interest, principal settlement and risk share of defaults |
| Districts | Per LGA: number of CMGs, savings, loan asked → capped loan, demand, allocation (scaled if demand > corpus) |
| Monthly cycle | 24 months (year 1 lending, year 2 run-off): new loans, repayments, TAMISEMI's share of repaid principal, lender's risk-share compensation, cash, loans out; year-end snapshot |
| Fund flow | Totals for all loans made in year 1, split of interest, risk share of defaults, corpus check, bar chart |

## Part 2 (new in v0.5): the 8-pilot-district weekly model
A separate, lighter analysis: how loans actually get issued week by week in 8 pilot districts, and whether the
corpus should be handed out upfront or in tranches. **Uses its own $6m placeholder pool, not yet reconciled
with the $20m above (question 13).**

| Sheet | What it does |
|---|---|
| Pilot districts | 8 districts: CMGs, savings/loan data, # FIs, weekly appraisal capacity, appraisal→disbursal lag, rollout start week |
| District timeline | Pick one district (dropdown) and see, week by week: appraisals, disbursals, Portfolio Outstanding, and when its pool runs out. A redeployment-mode toggle controls whether repayments are recycled, and by whom |
| Corpus rollout | All 8 districts side by side. A tranche % input compares giving each district its full allocation upfront vs. a smaller first tranche with the rest held as reserve at TAMISEMI — with idle-cash-at-TAMISEMI charted over time either way |

**What the placeholder run shows:** with an 80% appraisal→disbursal conversion rate, District 1 never actually
uses its full $800k allocation (only ~$640k gets disbursed, since only 80% of its 400 CMGs convert to loans) —
its pool is never "exhausted" in the strict sense. Once all of a district's CMGs have been appraised, recycled
repayments have nowhere to go and just sit as idle capital. With a 70% tranche, TAMISEMI keeps about $1.86m in
reserve once all 8 districts have rolled out (by week 17) — enough, in this placeholder run, to cover a
$1.5m call for a new region.

## Open questions and glossary
Both sheets now cover the whole model (annual + pilot). See `OPEN_QUESTIONS.md` for the full list — three new
ones (13-15) are about the pilot model's own placeholders and how it relates to the annual $20m corpus.

## The GUI: PAMOJA Model Inputs
[PAMOJA Model Inputs](https://claude.ai/artifact/HcwpiFMz2g1dmxUu5ueG2g) covers the annual model's inputs
(Inputs + Districts sheets) with a live preview. **It does not yet cover the pilot-district sheets** — those are
still edited directly in `PILOT_DISTRICTS` / `PILOT_CORPUS` etc. in `build_model.py`, or in the workbook itself.

To rebuild:
```
pip install openpyxl
python3 model/build_model.py [config.json]   # config.json only covers the annual model's inputs so far
```
