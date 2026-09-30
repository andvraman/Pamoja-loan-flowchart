"""Builds PAMOJA_Financial_Model_v0.5.xlsx (simple fund-flow model).
Run: python3 model/build_model.py [config.json]   (needs openpyxl)
All numbers in blue on yellow are inputs. Everything else is a formula.

config.json (optional) overrides the defaults below. It is the same file the
'PAMOJA Model Inputs' GUI (an Artifact page) saves. Shape:
  {"inputs": {"fund": 24000000, "ratio": 0.5, ...},
   "districts": [{"name": "District A", "cmgs": 2500, "savings": 4000, "loan_ask": 2500}, ...]}
Any key left out keeps its default. See DEFAULT_INPUTS / DEFAULT_DISTRICTS below for every key.
"""
import json
import sys
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

VERSION = "v0.5"
OUT = f"model/PAMOJA_Financial_Model_{VERSION}.xlsx"
MONTHS = 24          # year 1 (lending) + year 2 (loans made late in year 1 finish)

# ---------------------------------------------------------------- Config (GUI feeds this in)
DEFAULT_INPUTS = {
    "fund": 24_000_000, "linc": 2_000_000, "cinc": 2_000_000,
    "ratio": 0.5, "rate": 0.18, "maxterm": 12, "term": 6, "def": 0.05,
    "lastm": 12, "relend": 0.0,
    "sdef": 0.30, "stam": 0.10, "sreg": 0.05, "slga": 0.05,
    "ptam": 1.00,   # TAMISEMI share of repaid principal (remainder kept by lender)
    "rtam": 0.50,   # TAMISEMI risk share of defaults (remainder is the lender's risk share)
}
DEFAULT_DISTRICTS = [
    {"name": "District A", "cmgs": 2500, "savings": 4000, "loan_ask": 2500},
    {"name": "District B", "cmgs": 1800, "savings": 3000, "loan_ask": 2000},
    {"name": "District C", "cmgs": 3200, "savings": 5000, "loan_ask": 3000},
    {"name": "District D", "cmgs": 1500, "savings": 3500, "loan_ask": 1500},
    {"name": "District E", "cmgs": 2000, "savings": 4500, "loan_ask": 2500},
]

# 8 pilot districts for the weekly issuance timeline (Pilot districts / District timeline /
# Corpus rollout sheets). A separate, smaller placeholder pool from the annual $20m model above -
# not yet reconciled with it (see open question 13).
PILOT_CORPUS = 6_000_000
PILOT_CONVERSION = 0.80   # share of appraised CMGs that are actually approved AND accept the loan
PILOT_TERM_WEEKS = 26
PILOT_REDEPLOY = 1        # 0 = hold as cash, 1 = same FI redeploys, 2 = handed to a 2nd FI in district
PILOT_DISTRICTS = [
    {"name": "District 1", "cmgs": 400, "savings": 4000, "loan_ask": 2500, "fis": 2, "cap_per_fi": 6, "lag": 2, "rollout": 1},
    {"name": "District 2", "cmgs": 350, "savings": 3500, "loan_ask": 2200, "fis": 2, "cap_per_fi": 5, "lag": 2, "rollout": 1},
    {"name": "District 3", "cmgs": 500, "savings": 4500, "loan_ask": 2800, "fis": 3, "cap_per_fi": 5, "lag": 3, "rollout": 3},
    {"name": "District 4", "cmgs": 300, "savings": 3000, "loan_ask": 2000, "fis": 1, "cap_per_fi": 6, "lag": 2, "rollout": 5},
    {"name": "District 5", "cmgs": 450, "savings": 5000, "loan_ask": 3000, "fis": 2, "cap_per_fi": 5, "lag": 3, "rollout": 5},
    {"name": "District 6", "cmgs": 250, "savings": 3500, "loan_ask": 2200, "fis": 1, "cap_per_fi": 5, "lag": 2, "rollout": 9},
    {"name": "District 7", "cmgs": 380, "savings": 4000, "loan_ask": 2500, "fis": 2, "cap_per_fi": 4, "lag": 3, "rollout": 13},
    {"name": "District 8", "cmgs": 320, "savings": 3800, "loan_ask": 2300, "fis": 1, "cap_per_fi": 6, "lag": 2, "rollout": 17},
]
PILOT_WEEKS = 104   # 2-year horizon


def load_config(path):
    cfg_inputs, cfg_districts = dict(DEFAULT_INPUTS), list(DEFAULT_DISTRICTS)
    if path:
        with open(path) as fh:
            raw = json.load(fh)
        cfg_inputs.update({k: v for k, v in raw.get("inputs", {}).items() if k in DEFAULT_INPUTS})
        rows = raw.get("districts")
        if rows:
            cfg_districts = [{"name": r.get("name", ""), "cmgs": r.get("cmgs"),
                               "savings": r.get("savings"), "loan_ask": r.get("loan_ask")}
                              for r in rows if r.get("name")]
    return cfg_inputs, cfg_districts


CFG_PATH = sys.argv[1] if len(sys.argv) > 1 else None
CFG, CFG_DISTRICTS = load_config(CFG_PATH)

wb = Workbook()

INPUT = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor="1C5A3F")
SUB = PatternFill("solid", fgColor="DCEBE1")
GREY = PatternFill("solid", fgColor="EEEEEE")
BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
WHITE = Font(bold=True, color="FFFFFF")
thin = Side(style="thin", color="C6CDBE")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
USD = '"$"#,##0;[Red]-"$"#,##0'
PCT = "0.0%"
NUM = "#,##0"


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=14, color="1C5A3F")
    ws["A2"] = sub
    ws["A2"].font = Font(italic=True, color="5B6B60")


def header(ws, row, labels):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.fill, c.font, c.border = HEAD, WHITE, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")


def inp(c, fmt):
    c.fill, c.font, c.number_format, c.border = INPUT, BLUE, fmt, BOX


def section(ws, r, text, ncols=4):
    ws.cell(row=r, column=1, value=text).font = BOLD
    for c in range(1, ncols + 1):
        ws.cell(row=r, column=c).fill = SUB


def table(ws, start, rows, ncols=4):
    """rows: (key, label, value, fmt, unit/from, note). key=None -> section.
    Returns {key: row number} so formulas can refer to rows by name."""
    at = {}
    for r, (key, label, val, fmt, unit, note) in enumerate(rows, start):
        if key is None:
            section(ws, r, label, ncols)
            continue
        at[key] = r
        ws.cell(row=r, column=1, value=label)
        c = ws.cell(row=r, column=2)
        if callable(val):
            val = val(at)
        c.value = val
        if isinstance(val, str) and val.startswith("="):
            c.number_format, c.border = fmt, BOX
        else:
            inp(c, fmt)
        ws.cell(row=r, column=3, value=unit)
        ws.cell(row=r, column=4, value=note)
    return at


# ---------------------------------------------------------------- Open questions
QUESTIONS = [
    # (number, question, why it matters, added in)
    (1, "Is the Loan:Savings cap on the loan itself (loan must not exceed 0.5 x the CMG's savings)?",
     "Sets the loan cap per CMG on the Districts sheet.", "v0.1"),
    (2, "Is interest flat or on the reducing balance? What is TAMISEMI's interest ceiling?",
     "Drives all interest income, and so the default budget.", "v0.1"),
    (3, "Are the TAMISEMI, region and LGA amounts a % of interest, or a % of the loan amount?",
     "Changes how much interest is left for the default budget and the lender.", "v0.1"),
    (4, "Can repayments be lent out again in the same year? If yes, to whom: new CMGs or second loans to the same CMGs?",
     "Re-lending raises total lending but ties money up longer; the Districts sheet only allocates the first round.", "v0.1"),
    (5, "Are the $2m lender and $2m CMG incentive pots paid in full each year, or spread over several years?",
     "Decides whether the corpus stays at $20m every year.", "v0.1"),
    (6, "When does TAMISEMI's annual cycle end (e.g. financial year July-June)? At that date, must the corpus be back as CASH, or can loans still running count?",
     "With lending across the cycle and tenures up to 12 months, some money is always out on loan at year end.", "v0.2"),
    (7, "Do CMG savings cycles start in different months (staggered), or all at once?",
     "A loan must end before the CMG's share-out, so a 12-month loan only fits a CMG that borrows at the very start of its cycle.", "v0.2"),
    (8, "With the lender discount removed, is the lender's only income the interest margin plus the lender incentives? Does the lender pay anything for the TAMISEMI advance?",
     "Shows whether lending is worth it for the lender.", "v0.2"),
    (9, "Should the 15-working-day disbursement rule apply to each new advance across the cycle, not only the first?",
     "Sets how fast money can go out each month.", "v0.2"),
    (10, "Are defaulted loans written off at maturity, or chased into the next year (recoveries)?",
     "Recoveries would add money back to the corpus later.", "v0.2"),
    (11, "Who should be able to edit the 'PAMOJA Model Inputs' GUI, and does each district need its own editor?",
     "The GUI page is currently shared only with its owner; wider access needs a deliberate share decision.", "v0.3"),
    (12, "What should the risk share of defaults between TAMISEMI and the lender actually be, and should it "
     "vary by lender or by performance? Should the lender ever keep part of the repaid principal instead of "
     "returning all of it to TAMISEMI?",
     "Placeholders now: TAMISEMI absorbs 50% of every default, the lender the other 50%; the lender returns "
     "100% of repaid principal to TAMISEMI. Matches question 12 on the flowchart (trigger and timing of risk "
     "share are still open).", "v0.4"),
    (13, "How does the 8-district pilot pool ($6m placeholder) relate to the $24m annual model - is it part of "
     "the same $20m corpus, or a separate first tranche?",
     "The two are not reconciled yet: 'Pilot districts' has its own corpus figure and its own scale-down logic.", "v0.5"),
    (14, "Appraisal capacity, the appraisal->disbursal lag, and the conversion rate are all placeholders per "
     "district (and constant over time) - what are the real figures, and do they vary by season or by lender?",
     "These three numbers set the whole pace of the weekly cascade on 'District timeline'.", "v0.5"),
    (15, "Should recycled repayments be capped at what the district can still use (CMGs not yet served), or can "
     "they always be redeployed even if it means lending to the same CMGs again?",
     "In the placeholder run, once a district's CMGs are all appraised, recycled cash sits idle in the pool - "
     "see 'District timeline' for District 1.", "v0.5"),
]

# ---------------------------------------------------------------- Read me
rm = wb.active
rm.title = "Read me"
title(rm, f"PAMOJA simplified financial model - {VERSION} (draft)",
      "How $24m flows TAMISEMI -> Lenders -> CMGs as loans, and back as repayments.")
lines = [
    "",
    "HOW TO USE",
    "1. Change the yellow cells with blue numbers only. Everything else is a formula.",
    "2. Inputs: fund size, lending rules, lending timing and the split of interest.",
    "3. Districts: CMG numbers, savings and loan size per LGA.",
    "4. Monthly cycle: when loans go out and come back over 24 months (year 1 lending, year 2 run-off).",
    "5. Fund flow: totals for all loans made in the year, with a check that the corpus is restored.",
    "6. Pilot districts / District timeline / Corpus rollout: a separate, weekly view of loan issuance",
    "   for 8 pilot districts (see CHANGES IN v0.5 below).",
    "7. Open questions: things to settle. Glossary: acronyms and terms.",
    "",
    "IMPORTANT",
    "- All district data and rates are DUMMY placeholders. Replace them with real figures.",
    "- The split of interest is a first guess to confirm.",
    "",
    "CHANGES IN v0.2",
    "- Lender discount removed. The lender keeps the margin (what is left after the other shares).",
    "- Lenders can make new loans in any month of the year, up to a last month you choose.",
    "- Loan term can be up to 12 months (maximum tenure is an input).",
    "- Repayments can be lent out again (re-lending %, set to 0% to switch off).",
    "- Monthly sheet runs 24 months so loans made late in year 1 can finish in year 2.",
    "- Year-end snapshot: cash back in the fund vs money still out on loan at month 12.",
    "",
    "CHANGES IN v0.3",
    "- All independent inputs (Inputs + Districts) can now come from a config.json file,",
    "  built by the 'PAMOJA Model Inputs' GUI page instead of typed into Excel.",
    "  Run: python3 model/build_model.py config.json",
    "- Fund flow now has a bar chart: money not lent, principal repaid, interest collected.",
    "",
    "CHANGES IN v0.4",
    "- Region share and LGA share of interest renamed to Region service fee and LGA service fee.",
    "- New: TAMISEMI share of repaid principal (default 100%) - the rest is kept by the lender.",
    "- New: Risk share of defaults, split between TAMISEMI and the lender. TAMISEMI absorbs its",
    "  share against the corpus; the lender absorbs the remainder (a computed field) as its own loss,",
    "  and pays that amount back into the fund as risk-share compensation.",
    "- Fund flow shows both settlements and a 'Lender's net result' line (margin - own risk share",
    "  of defaults + principal kept back).",
    "- Break-even default rate and the default budget needed now only have to cover TAMISEMI's",
    "  risk share of defaults, not the whole default.",
    "",
    "CHANGES IN v0.5 - the pilot-district weekly model",
    "- Three new sheets model 8 pilot districts week by week (separate from the annual model above,",
    "  with its own $6m placeholder pool - see question 13):",
    "  * Pilot districts: CMG data, FIs, weekly appraisal capacity, appraisal->disbursal lag and",
    "    rollout start week for each of the 8 districts.",
    "  * District timeline: pick one district and see, week by week, appraisals, disbursals, the",
    "    loan book (Portfolio Outstanding), and when its pool would run out - with a toggle for",
    "    whether repayments are recycled and by whom.",
    "  * Corpus rollout: all 8 districts side by side, with a tranche % input comparing 'give each",
    "    district its full allocation upfront' against 'a smaller tranche now, rest held in reserve",
    "    at TAMISEMI' - showing idle cash at TAMISEMI over time either way.",
    "- Not every appraised CMG becomes a loan: a conversion rate (appraisal pass + CMG acceptance,",
    "  80% placeholder) reduces disbursals below capacity.",
    "",
    "SIMPLE LOGIC",
    "- Lending corpus = Total fund - Lender incentives - CMG incentives.",
    "- Loan per CMG = the smaller of: average loan asked for, or savings x maximum Loan:Savings ratio.",
    "- If total demand is more than the corpus, every district is scaled down by the same %.",
    "- Loans are repaid in equal monthly instalments over the term, starting the month after the loan.",
    "- Interest is flat: rate x loan x term, collected with each instalment.",
    "- Defaults: the default rate is taken off every instalment (no principal or interest).",
    "- Re-lending: a % of principal repaid in one month is lent out again the next month, up to the last lending month.",
    "- Of every instalment repaid, TAMISEMI's share of repaid principal returns to the fund; the rest",
    "  stays with the lender.",
    "- Of every default, TAMISEMI absorbs its risk share against the corpus; the lender absorbs the",
    "  rest and pays that amount back into the fund.",
    "- The corpus is restored if the default budget (from interest) covers TAMISEMI's risk share of",
    "  the defaulted principal, after the lender's risk-share payment and any principal it keeps.",
    "",
    "KEY TRADE-OFF",
    "- For all money to be back as CASH by month 12, the last loan must go out by month (12 - term).",
    "- With a 12-month term, that means lending only in month 0 - so lending across the cycle means",
    "  some of the corpus is still out on loan at year end. See question 6.",
    "",
    "NOT YET IN THE MODEL",
    "- Different terms per loan or district; staggered CMG cycles.",
    "- Multiple lenders per LGA; region-level roll-up.",
    "- Timing of incentive payments; LGA service fee formula; risk-share formula (Q12 in the flowchart).",
    "- Exchange rate (TZS / USD) and inflation.",
]
for i, t in enumerate(lines, 4):
    rm.cell(row=i, column=1, value=t)
    if t and t.upper() == t and not t.startswith("-"):
        rm.cell(row=i, column=1).font = BOLD
rm.column_dimensions["A"].width = 110

# ---------------------------------------------------------------- Inputs
ip = wb.create_sheet("Inputs")
title(ip, "Inputs", "Yellow cells are inputs. Dummy values - replace with agreed figures.")
header(ip, 4, ["Item", "Value", "Unit", "Note"])
I = table(ip, 5, [
    (None, "FUND", None, None, None, None),
    ("fund", "Total PAMOJA fund", CFG["fund"], USD, "USD", "From TAMISEMI"),
    ("linc", "Lender incentives (set aside)", CFG["linc"], USD, "USD", "Not lent out"),
    ("cinc", "CMG incentives (set aside)", CFG["cinc"], USD, "USD", "Not lent out"),
    ("corpus", "Lending corpus", lambda a: f"=B{a['fund']}-B{a['linc']}-B{a['cinc']}", USD, "USD",
     "Money available to lend"),
    (None, "LENDING RULES", None, None, None, None),
    ("ratio", "Maximum Loan:Savings ratio", CFG["ratio"], "0.00", "x", "Loan must not exceed this x CMG savings"),
    ("rate", "Interest rate (flat, per year)", CFG["rate"], PCT, "%", "Placeholder; must be at or below TAMISEMI ceiling"),
    ("maxterm", "Maximum loan tenure", CFG["maxterm"], "0", "months", "Policy limit"),
    ("term", "Loan term used in the model", CFG["term"], "0", "months", "Must be 1 to the maximum tenure"),
    ("termok", "Check: term within maximum",
     lambda a: f'=IF(AND(B{a["term"]}>=1,B{a["term"]}<=B{a["maxterm"]}),"OK","ERROR: term above maximum tenure")',
     "@", "", ""),
    ("def", "Portfolio default rate (budget)", CFG["def"], PCT, "%", "% of principal lent that is not repaid"),
    (None, "LENDING TIMING (across the cycle)", None, None, None, None),
    ("lastm", "Last month lenders can make new loans", CFG["lastm"], "0", "month", "1 to 12. Spread of new loans is set on 'Monthly cycle'"),
    ("relend", "% of principal repaid that is lent out again", CFG["relend"], PCT, "%", "0% = no re-lending (see question 4)"),
    ("cashm", "Latest month to lend if all cash must be back by month 12",
     lambda a: f"=MAX(0,12-B{a['term']})", "0", "month", "0 means only at the very start of the year"),
    (None, "SPLIT OF INTEREST COLLECTED", None, None, None, None),
    ("sdef", "Default budget (loss reserve)", CFG["sdef"], PCT, "%", "Returned to fund to cover defaults"),
    ("stam", "TAMISEMI share of interest", CFG["stam"], PCT, "%", "Placeholder"),
    ("sreg", "Region service fee", CFG["sreg"], PCT, "%", "Placeholder"),
    ("slga", "LGA service fee", CFG["slga"], PCT, "%", "Placeholder"),
    ("slen", "Lender margin (remainder)", lambda a: f"=1-SUM(B{a['sdef']}:B{a['slga']})", PCT, "%",
     "What is left for the lender"),
    ("splitok", "Check: split adds to 100% or less",
     lambda a: f'=IF(B{a["slen"]}>=0,"OK","ERROR: over 100%")', "@", "", ""),
    (None, "PRINCIPAL SETTLEMENT AND RISK SHARE", None, None, None, None),
    ("ptam", "TAMISEMI share of repaid principal", CFG["ptam"], PCT, "%",
     "Share of every repaid instalment returned to the TAMISEMI corpus"),
    ("plen", "Lender share of repaid principal (remainder)", lambda a: f"=1-B{a['ptam']}", PCT, "%",
     "Kept by the lender; computed from the TAMISEMI share above"),
    ("rtam", "TAMISEMI risk share of defaults", CFG["rtam"], PCT, "%",
     "Share of every default TAMISEMI absorbs against the corpus. Placeholder - see question 12"),
    ("rlen", "Lender risk share of defaults (remainder)", lambda a: f"=1-B{a['rtam']}", PCT, "%",
     "The lender's own loss on a default; computed from the TAMISEMI risk share above"),
])
for key in ("corpus", "slen", "termok", "splitok", "cashm", "plen", "rlen"):
    ip[f"B{I[key]}"].font = BOLD
for L, w in zip("ABCD", (48, 16, 8, 58)):
    ip.column_dimensions[L].width = w


def R(key):
    return f"Inputs!$B${I[key]}"


CORPUS, RATIO, RATE, TERM, DEF = R("corpus"), R("ratio"), R("rate"), R("term"), R("def")

# ---------------------------------------------------------------- Districts
ds = wb.create_sheet("Districts")
title(ds, "Allocation by district (LGA)",
      "Dummy data. Add real LGAs in the yellow cells; up to 20 rows. Allocates the first round of lending only.")
header(ds, 4, ["District (LGA)", "Number of CMGs", "Average savings per CMG (USD)",
               "Average loan asked per CMG (USD)", "Loan cap per CMG = savings x ratio",
               "Loan per CMG (smaller of the two)", "Loan:Savings ratio used",
               "Demand (USD)", "Allocation (USD)", "Share of corpus"])
ds.row_dimensions[4].height = 45
FIRST, LAST = 5, 24
T = LAST + 1
for r in range(FIRST, LAST + 1):
    i = r - FIRST
    row = CFG_DISTRICTS[i] if i < len(CFG_DISTRICTS) else None
    data = (row["name"], row["cmgs"], row["savings"], row["loan_ask"]) if row else (None, None, None, None)
    ds.cell(row=r, column=1, value=data[0]).fill = INPUT
    for c, v, f in ((2, data[1], NUM), (3, data[2], USD), (4, data[3], USD)):
        inp(ds.cell(row=r, column=c, value=v), f)
    ds.cell(row=r, column=5, value=f'=IF(B{r}="","",C{r}*{RATIO})').number_format = USD
    ds.cell(row=r, column=6, value=f'=IF(B{r}="","",MIN(D{r},E{r}))').number_format = USD
    ds.cell(row=r, column=7, value=f'=IF(B{r}="","",F{r}/C{r})').number_format = "0.00"
    ds.cell(row=r, column=8, value=f'=IF(B{r}="",0,B{r}*F{r})').number_format = USD
    ds.cell(row=r, column=9, value=f'=IF(B{r}="",0,H{r}*$B${T + 3})').number_format = USD
    ds.cell(row=r, column=10, value=f'=IF(B{r}="","",I{r}/{CORPUS})').number_format = PCT
ds.cell(row=T, column=1, value="TOTAL").font = BOLD
for c, f in ((2, NUM), (8, USD), (9, USD), (10, PCT)):
    L = col(c)
    cell = ds.cell(row=T, column=c, value=f"=SUM({L}{FIRST}:{L}{LAST})")
    cell.number_format, cell.font = f, BOLD
ds[f"A{T+2}"], ds[f"B{T+2}"] = "Lending corpus", f"={CORPUS}"
ds[f"A{T+3}"], ds[f"B{T+3}"] = "Scale factor (1 = full demand met)", f"=IF(H{T}=0,0,MIN(1,B{T+2}/H{T}))"
ds[f"A{T+4}"], ds[f"B{T+4}"] = "Corpus not used", f"=B{T+2}-I{T}"
ds[f"A{T+5}"], ds[f"B{T+5}"] = "Demand not met", f"=H{T}-I{T}"
for r, f in ((T + 2, USD), (T + 3, "0.00"), (T + 4, USD), (T + 5, USD)):
    ds[f"B{r}"].number_format, ds[f"B{r}"].font = f, BOLD
ds.column_dimensions["A"].width = 34
for c in range(2, 11):
    ds.column_dimensions[col(c)].width = 17
ds.freeze_panes = "B5"
ALLOC = f"Districts!$I${T}"

# ---------------------------------------------------------------- Monthly
mo = wb.create_sheet("Monthly cycle")
title(mo, "Monthly cycle (24 months: year 1 lending, year 2 run-off)",
      "Enter the % of the allocated amount lent as NEW loans each month (yellow, months 1-12). "
      "Repayments are equal monthly instalments over the loan term.")
LC = col(MONTHS + 2)                     # total column
header(mo, 4, ["Line"] + [f"M{m}" for m in range(1, MONTHS + 1)] + ["Total"])
M = {k: r for r, k in enumerate(
    ["month", "year", "pct", "new", "relent", "disb", "due", "repaid", "interest",
     "reserve", "tocorp", "lencomp", "out", "cash", "total", "backflag"], 5)}
labels = {
    "month": "Month number", "year": "Year",
    "pct": "% of allocation lent as NEW loans this month",
    "new": "New loans from the corpus", "relent": "Re-lent repayments",
    "disb": "Total loans made this month", "due": "Principal due",
    "repaid": "Principal repaid by CMGs (gross, after defaults)", "interest": "Interest collected",
    "reserve": "Default budget returned to fund",
    "tocorp": "TAMISEMI's share of repaid principal (to fund cash)",
    "lencomp": "Lender's risk-share compensation for defaults (to fund cash)",
    "out": "Loans outstanding (end of month)", "cash": "Fund cash on hand (end of month)",
    "total": "Cash + loans outstanding", "backflag": "(helper) month if cash is back to corpus",
}
for k, r in M.items():
    mo.cell(row=r, column=1, value=labels[k])
mo.cell(row=M["backflag"], column=1).font = Font(italic=True, color="8A978C")
pattern = [0.25, 0.20, 0.15, 0.10, 0.10, 0.10, 0.05, 0.05, 0, 0, 0, 0]
MR = f"$B${M['month']}"
for m in range(1, MONTHS + 1):
    c, p = col(m + 1), col(m)            # this column, previous column
    mo[f"{c}{M['month']}"] = m
    mo[f"{c}{M['month']}"].font = BOLD
    mo[f"{c}{M['year']}"] = 1 if m <= 12 else 2
    cell = mo[f"{c}{M['pct']}"]
    if m <= 12:
        cell.value = pattern[m - 1]
        inp(cell, PCT)
    else:
        cell.value, cell.fill, cell.number_format = 0, GREY, PCT
    mo[f"{c}{M['new']}"] = f"=IF({c}${M['month']}<={R('lastm')},{c}{M['pct']}*{ALLOC},0)"
    mo[f"{c}{M['relent']}"] = ("=0" if m == 1 else
                               f"=IF({c}${M['month']}<={R('lastm')},{p}{M['repaid']}*{R('relend')},0)")
    mo[f"{c}{M['disb']}"] = f"={c}{M['new']}+{c}{M['relent']}"
    # principal due in month m from loans made in earlier months d with 1 <= m-d <= term
    if m == 1:
        mo[f"{c}{M['due']}"] = "=0"
    else:
        mo[f"{c}{M['due']}"] = (f"=SUMIFS($B{M['disb']}:{p}{M['disb']},{MR}:{p}${M['month']},"
                                f"\">=\"&({c}${M['month']}-{TERM}))/{TERM}")
    mo[f"{c}{M['repaid']}"] = f"={c}{M['due']}*(1-{DEF})"
    mo[f"{c}{M['interest']}"] = f"={c}{M['repaid']}*{RATE}/12*{TERM}"
    mo[f"{c}{M['reserve']}"] = f"={c}{M['interest']}*{R('sdef')}"
    mo[f"{c}{M['tocorp']}"] = f"={c}{M['repaid']}*{R('ptam')}"
    mo[f"{c}{M['lencomp']}"] = f"=({c}{M['due']}-{c}{M['repaid']})*{R('rlen')}"
    prev_out = "0" if m == 1 else f"{p}{M['out']}"
    mo[f"{c}{M['out']}"] = f"={prev_out}+{c}{M['disb']}-{c}{M['due']}"
    prev_cash = CORPUS if m == 1 else f"{p}{M['cash']}"
    mo[f"{c}{M['cash']}"] = (f"={prev_cash}-{c}{M['disb']}+{c}{M['tocorp']}"
                             f"+{c}{M['lencomp']}+{c}{M['reserve']}")
    mo[f"{c}{M['total']}"] = f"={c}{M['cash']}+{c}{M['out']}"
    mo[f"{c}{M['backflag']}"] = f"=IF({c}{M['cash']}>={CORPUS}-0.5,{c}{M['month']},999)"
    for k in ("new", "relent", "disb", "due", "repaid", "interest", "reserve",
              "tocorp", "lencomp", "out", "cash", "total"):
        mo[f"{c}{M[k]}"].number_format = USD
    if m == 12:
        for k in M:
            mo[f"{c}{M[k]}"].border = Border(right=Side(style="thick", color="1C5A3F"))
first, lastc = "B", col(MONTHS + 1)
mo[f"{LC}{M['pct']}"] = f"=SUM(B{M['pct']}:M{M['pct']})"
mo[f"{LC}{M['pct']}"].number_format = PCT
for k in ("new", "relent", "disb", "due", "repaid", "interest", "reserve", "tocorp", "lencomp"):
    mo[f"{LC}{M[k]}"] = f"=SUM({first}{M[k]}:{lastc}{M[k]})"
    mo[f"{LC}{M[k]}"].number_format, mo[f"{LC}{M[k]}"].font = USD, BOLD
for k in ("disb", "cash", "out"):
    mo.cell(row=M[k], column=1).font = BOLD

S = M["backflag"] + 2
section(mo, S, "YEAR-END SNAPSHOT (end of month 12)", 4)
snap = [
    ("Fund cash on hand at month 12", f"=M{M['cash']}", USD),
    ("Money still out on loan at month 12 (repaid in year 2)", f"=M{M['out']}", USD),
    ("Cash + loans still out", f"=M{M['total']}", USD),
    ("Share of corpus back as cash at month 12", f"=M{M['cash']}/{CORPUS}", PCT),
    ("Total loans made in year 1 (incl. re-lending)", f"=SUM(B{M['disb']}:M{M['disb']})", USD),
    ("Times the corpus was lent in year 1 (turnover)", f"=SUM(B{M['disb']}:M{M['disb']})/{CORPUS}", "0.00"),
    ("First month the corpus is fully back as cash",
     f'=IF(MIN(B{M["backflag"]}:{lastc}{M["backflag"]})=999,"Not within 24 months (defaults not covered)",'
     f'MIN(B{M["backflag"]}:{lastc}{M["backflag"]}))', "0"),
    ("Fund cash at month 24 (all loans closed)", f"={lastc}{M['cash']}", USD),
    ("Check: new-loan % adds to 100%", f'=IF(ABS({LC}{M["pct"]}-1)<0.0001,"OK","ERROR: adjust row {M["pct"]}")', "@"),
    ("Check: no new loans after the last lending month",
     f'=IF(SUMPRODUCT((B{M["month"]}:M{M["month"]}>{R("lastm")})*B{M["pct"]}:M{M["pct"]})>0,'
     f'"WARNING: some % is after the last lending month and is ignored","OK")', "@"),
]
SN = {}
for i, (lab, f, fmt) in enumerate(snap, S + 1):
    mo[f"A{i}"], mo[f"B{i}"] = lab, f
    mo[f"B{i}"].number_format, mo[f"B{i}"].font = fmt, BOLD
    SN[lab] = i
mo[f"A{i + 2}"] = ("Note: fund cash counts TAMISEMI's share of repaid principal, the lender's risk-share "
                   "compensation for defaults, and the default-budget share of interest. Interest shares also go "
                   "to TAMISEMI, the region, the LGA and the lender; any repaid principal kept by the lender does "
                   "not come back to the fund. The thick line marks year end.")
mo.column_dimensions["A"].width = 52
for m in range(2, MONTHS + 3):
    mo.column_dimensions[col(m)].width = 13
mo.freeze_panes = "B5"


def MO(k):
    return f"'Monthly cycle'!${LC}${M[k]}"


# ---------------------------------------------------------------- Fund flow
ff = wb.create_sheet("Fund flow", index=3)
title(ff, "Fund flow: all loans made in year 1, until they close",
      "TAMISEMI -> Lenders -> CMGs -> back. All formulas; totals come from 'Monthly cycle'.")
header(ff, 4, ["Step", "Amount (USD)", "From -> To", "How it is worked out"])
F = table(ff, 5, [
    (None, "1. MONEY OUT", None, None, None, None),
    ("fund", "Total PAMOJA fund", f"={R('fund')}", USD, "TAMISEMI", ""),
    ("linc", "Lender incentives set aside", f"={R('linc')}", USD, "TAMISEMI -> Lenders", "Paid later, not lent"),
    ("cinc", "CMG incentives set aside", f"={R('cinc')}", USD, "TAMISEMI -> CMGs", "Paid later, not lent"),
    ("adv", "Advances to lenders (new loans from corpus)", f"={MO('new')}", USD, "TAMISEMI -> Lenders",
     "Sum over the months"),
    ("unused", "Corpus not lent (returned)", lambda a: f"={CORPUS}-B{a['adv']}", USD, "Lenders -> TAMISEMI",
     "Unused advance"),
    ("relent", "Repayments lent out again", f"={MO('relent')}", USD, "Lenders -> CMGs", "Re-lending"),
    ("loans", "Total loans to CMGs", lambda a: f"=B{a['adv']}+B{a['relent']}", USD, "Lenders -> CMGs", ""),
    (None, "2. MONEY BACK FROM CMGs", None, None, None, None),
    ("dft", "Principal defaulted", lambda a: f"=B{a['loans']}*{DEF}", USD, "", "Loans x default rate"),
    ("rep", "Principal repaid by CMGs (gross)", f"={MO('repaid')}", USD, "CMGs -> Lenders", "Loans - defaults"),
    ("tocorp", "TAMISEMI's share of repaid principal", f"={MO('tocorp')}", USD, "Lenders -> TAMISEMI",
     "Repaid gross x TAMISEMI share of repaid principal"),
    ("lenkeep", "Kept by the lender (repaid principal not returned)",
     lambda a: f"=B{a['rep']}-B{a['tocorp']}", USD, "Kept by lenders", "Repaid gross x lender share"),
    ("int", "Interest collected", f"={MO('interest')}", USD, "CMGs -> Lenders", "Flat: repaid x rate x term/12"),
    ("tot", "Total repayments", lambda a: f"=B{a['rep']}+B{a['int']}", USD, "CMGs -> Lenders", ""),
    (None, "3. SPLIT OF INTEREST", None, None, None, None),
    ("sdef", "Default budget (loss reserve)", lambda a: f"=B{a['int']}*{R('sdef')}", USD, "Lenders -> TAMISEMI", ""),
    ("stam", "TAMISEMI share of interest", lambda a: f"=B{a['int']}*{R('stam')}", USD, "Lenders -> TAMISEMI", ""),
    ("sreg", "Region service fee", lambda a: f"=B{a['int']}*{R('sreg')}", USD, "-> Region", ""),
    ("slga", "LGA service fee", lambda a: f"=B{a['int']}*{R('slga')}", USD, "-> LGA", ""),
    ("slen", "Lender margin (from interest)", lambda a: f"=B{a['int']}*{R('slen')}", USD, "Kept by lenders", "Remainder"),
    (None, "4. RISK SHARE OF DEFAULTS", None, None, None, None),
    ("dtam", "TAMISEMI's risk share of defaults (absorbed by TAMISEMI)",
     lambda a: f"=B{a['dft']}*{R('rtam')}", USD, "Loss to TAMISEMI", "Defaulted principal x TAMISEMI risk share"),
    ("dlen", "Lender's risk share of defaults (the lender's own loss)",
     lambda a: f"=B{a['dft']}*{R('rlen')}", USD, "Loss to lender", "Defaulted principal x lender risk share"),
    ("comp", "Lender's risk-share compensation paid into the fund",
     f"={MO('lencomp')}", USD, "Lenders -> TAMISEMI", "Should equal the lender's risk share above"),
    ("lennet", "Lender's net result (margin - own risk share - principal kept)",
     lambda a: f"=B{a['slen']}-B{a['dlen']}+B{a['lenkeep']}", USD, "", "Whether lending is worth it for the lender"),
    (None, "5. CORPUS CHECK (after all loans close)", None, None, None, None),
    ("start", "Corpus at start", f"={CORPUS}", USD, "", ""),
    ("lost", "Less: TAMISEMI's risk share of defaults", lambda a: f"=-B{a['dtam']}", USD, "", ""),
    ("lostp", "Less: repaid principal kept by the lender", lambda a: f"=-B{a['lenkeep']}", USD, "", ""),
    ("add", "Plus default budget added back", lambda a: f"=B{a['sdef']}", USD, "", "Covers TAMISEMI's risk share of defaults"),
    ("end", "Corpus at end", lambda a: f"=B{a['start']}+B{a['lost']}+B{a['lostp']}+B{a['add']}", USD, "", ""),
    ("gap", "Gap (end - start)", lambda a: f"=B{a['end']}-B{a['start']}", USD, "", "Negative = corpus NOT restored"),
    ("res", "Result", lambda a: f'=IF(B{a["gap"]}>=-0.5,"Corpus restored",'
                                '"Corpus NOT restored - raise interest share for defaults or rate")', "@", "", ""),
    ("need", "Default budget needed (% of interest)", lambda a: f"=IF(B{a['int']}=0,0,B{a['dtam']}/B{a['int']})",
     PCT, "", "Minimum share of interest to cover TAMISEMI's risk share of defaults"),
    ("be", "Break-even default rate",
     f"={R('sdef')}*{RATE}*{TERM}/12/({R('rtam')}+{R('sdef')}*{RATE}*{TERM}/12)", PCT, "",
     "Highest default rate the current split and risk share can cover"),
    ("tie", "Check: matches month-24 cash on 'Monthly cycle'",
     lambda a: f"=IF(ABS(B{a['end']}-'Monthly cycle'!{lastc}{M['cash']})<1,\"OK\",\"ERROR\")", "@", "", ""),
    (None, "6. YEAR-END TIMING (from 'Monthly cycle')", None, None, None, None),
    ("ye_cash", "Cash back in the fund at month 12", f"='Monthly cycle'!B{SN['Fund cash on hand at month 12']}",
     USD, "", ""),
    ("ye_out", "Still out on loan at month 12",
     f"='Monthly cycle'!B{SN['Money still out on loan at month 12 (repaid in year 2)']}", USD, "",
     "Comes back in year 2"),
    ("ye_back", "First month corpus fully back as cash",
     f"='Monthly cycle'!B{SN['First month the corpus is fully back as cash']}", "0", "", ""),
])
for k in ("end", "res", "loans", "lennet"):
    ff[f"A{F[k]}"].font = ff[f"B{F[k]}"].font = BOLD
for Lc, w in zip("ABCD", (44, 18, 24, 48)):
    ff.column_dimensions[Lc].width = w

# Chart source: pulls the 3 lines the chart shows into one contiguous range,
# so the chart survives if the rows above ever move. Rows 10, 15, 16 today:
# money not lent, principal repaid, interest collected.
CH = {"unused": "Money not lent (returned to TAMISEMI)", "rep": "Principal repaid",
      "int": "Interest collected"}
ch_row = F["res"] + 3
ff.cell(row=ch_row - 1, column=6, value="Chart data (auto - do not edit)").font = Font(italic=True, color="8A978C")
for i, (k, lab) in enumerate(CH.items()):
    r = ch_row + i
    ff.cell(row=r, column=6, value=lab)
    c = ff.cell(row=r, column=7, value=f"=B{F[k]}")
    c.number_format = USD
chart = BarChart()
chart.type, chart.title, chart.y_axis.title = "col", "Year 1: money not lent, repaid, and interest earned", "USD"
chart.y_axis.numFmt = USD
chart.x_axis.delete = False
chart.style = 10
data = Reference(ff, min_col=7, min_row=ch_row, max_row=ch_row + len(CH) - 1)
cats = Reference(ff, min_col=6, min_row=ch_row, max_row=ch_row + len(CH) - 1)
chart.add_data(data, titles_from_data=False)
chart.set_categories(cats)
chart.series[0].tx = None
chart.legend = None
chart.width, chart.height = 16, 9
ff.add_chart(chart, f"A{ch_row + len(CH) + 2}")

# ---------------------------------------------------------------- Pilot districts
pd_ = wb.create_sheet("Pilot districts")
title(pd_, "Pilot districts", "8 districts for the weekly issuance timeline. All data below is placeholder.")
header(pd_, 4, ["Item", "Value", "Unit", "Note"])
PA = table(pd_, 5, [
    (None, "PILOT CORPUS AND SHARED ASSUMPTIONS", None, None, None, None),
    ("pcorpus", "Pilot corpus (USD)", PILOT_CORPUS, USD, "USD",
     "A separate, smaller pool from the annual $20m model - not yet reconciled with it (question 13)"),
    ("pratio", "Maximum Loan:Savings ratio (from Inputs)", f"={R('ratio')}", "0.00", "x", "Reused from the Inputs sheet"),
    ("prate", "Interest rate (from Inputs)", f"={R('rate')}", PCT, "%", "Reused from the Inputs sheet"),
    ("pdef", "Portfolio default rate (from Inputs)", f"={R('def')}", PCT, "%", "Reused from the Inputs sheet"),
    ("pterm", "Term used in this weekly model", PILOT_TERM_WEEKS, "0", "weeks", "Separate from the annual model's monthly term"),
    ("pconv", "Appraisal -> disbursal conversion rate", PILOT_CONVERSION, PCT, "%",
     "Share of appraised CMGs that are approved AND accept the loan"),
    ("predeploy", "Redeployment mode (District timeline sheet)", PILOT_REDEPLOY, "0", "0/1/2",
     "0 = hold repayments as cash, 1 = same FI redeploys them, 2 = handed to a 2nd FI in the district (same maths, different label)"),
])
for key in ("pratio", "prate", "pdef"):
    pd_[f"B{PA[key]}"].font = BOLD

PFIRST = PA["predeploy"] + 3
header(pd_, PFIRST - 1, ["District", "Number of CMGs", "Avg savings/CMG (USD)", "Avg loan ask/CMG (USD)",
                         "Loan cap/CMG", "Loan/CMG", "Demand (USD)", "Full allocation (USD)",
                         "# FIs assigned", "Weekly capacity/FI (CMGs)", "Combined weekly capacity",
                         "Application -> disbursal lag (wks)", "Rollout start week"])
pd_.row_dimensions[PFIRST - 1].height = 40
PLAST = PFIRST + len(PILOT_DISTRICTS) - 1
PT = PLAST + 1
for i, d in enumerate(PILOT_DISTRICTS):
    r = PFIRST + i
    pd_.cell(row=r, column=1, value=d["name"]).fill = INPUT
    for c, v, f in ((2, d["cmgs"], NUM), (3, d["savings"], USD), (4, d["loan_ask"], USD),
                    (9, d["fis"], NUM), (10, d["cap_per_fi"], NUM), (12, d["lag"], "0"), (13, d["rollout"], "0")):
        inp(pd_.cell(row=r, column=c, value=v), f)
    pd_.cell(row=r, column=5, value=f"=C{r}*$B${PA['pratio']}").number_format = USD
    pd_.cell(row=r, column=6, value=f"=MIN(D{r},E{r})").number_format = USD
    pd_.cell(row=r, column=7, value=f"=B{r}*F{r}").number_format = USD
    pd_.cell(row=r, column=8, value=f"=G{r}*$B${PT + 3}").number_format = USD
    pd_.cell(row=r, column=11, value=f"=I{r}*J{r}").number_format = NUM
pd_.cell(row=PT, column=1, value="TOTAL").font = BOLD
for c, f in ((2, NUM), (7, USD), (8, USD), (11, NUM)):
    L = col(c)
    cell = pd_.cell(row=PT, column=c, value=f"=SUM({L}{PFIRST}:{L}{PLAST})")
    cell.number_format, cell.font = f, BOLD
pd_[f"A{PT+2}"], pd_[f"B{PT+2}"] = "Pilot corpus", f"=B{PA['pcorpus']}"
pd_[f"A{PT+3}"], pd_[f"B{PT+3}"] = "Scale factor (1 = full demand met)", f"=IF(G{PT}=0,0,MIN(1,B{PT+2}/G{PT}))"
pd_[f"A{PT+4}"], pd_[f"B{PT+4}"] = "Corpus not used", f"=B{PT+2}-H{PT}"
pd_[f"A{PT+5}"], pd_[f"B{PT+5}"] = "Demand not met", f"=G{PT}-H{PT}"
for r, f in ((PT + 2, USD), (PT + 3, "0.00"), (PT + 4, USD), (PT + 5, USD)):
    pd_[f"B{r}"].number_format, pd_[f"B{r}"].font = f, BOLD
pd_.column_dimensions["A"].width = 14
for c in range(2, 14):
    pd_.column_dimensions[col(c)].width = 15
pd_.freeze_panes = "B" + str(PFIRST)

PNAMES = f"'Pilot districts'!$A${PFIRST}:$A${PLAST}"
PCMGS = f"'Pilot districts'!$B${PFIRST}:$B${PLAST}"
PLPC = f"'Pilot districts'!$F${PFIRST}:$F${PLAST}"
PALLOC = f"'Pilot districts'!$H${PFIRST}:$H${PLAST}"
PCAP = f"'Pilot districts'!$K${PFIRST}:$K${PLAST}"
PLAG = f"'Pilot districts'!$L${PFIRST}:$L${PLAST}"

# ---------------------------------------------------------------- District timeline
dt = wb.create_sheet("District timeline")
title(dt, "District timeline: weekly loan issuance cascade",
      "Pick a pilot district below. Weeks run down the rows so the cascade charts naturally.")
dt["A4"] = "Pick a district:"
dt["A4"].font = BOLD
dt["B4"] = PILOT_DISTRICTS[0]["name"]
dt["B4"].fill, dt["B4"].font, dt["B4"].border = INPUT, BLUE, BOX
dv = DataValidation(type="list", formula1=f"={PNAMES}", allow_blank=False)
dt.add_data_validation(dv)
dv.add(dt["B4"])

header(dt, 6, ["Item", "Value", "Unit", "Note"])
SEL = "$B$4"
TL = table(dt, 7, [
    (None, "SELECTED DISTRICT'S PARAMETERS (looked up from 'Pilot districts')", None, None, None, None),
    ("cmgs", "Number of CMGs", f"=INDEX({PCMGS},MATCH({SEL},{PNAMES},0))", NUM, "CMGs", ""),
    ("lpc", "Loan per CMG (USD)", f"=INDEX({PLPC},MATCH({SEL},{PNAMES},0))", USD, "USD", ""),
    ("alloc", "Full allocation (USD)", f"=INDEX({PALLOC},MATCH({SEL},{PNAMES},0))", USD, "USD", ""),
    ("cap", "Combined weekly appraisal capacity", f"=INDEX({PCAP},MATCH({SEL},{PNAMES},0))", NUM, "CMGs/wk", ""),
    ("lag", "Application -> disbursal lag", f"=INDEX({PLAG},MATCH({SEL},{PNAMES},0))", "0", "weeks", ""),
    ("term", "Term used in this weekly model", f"='Pilot districts'!B{PA['pterm']}", "0", "weeks", "From 'Pilot districts'"),
    ("rate", "Interest rate", f"='Pilot districts'!B{PA['prate']}", PCT, "%", "From 'Pilot districts'"),
    ("def", "Portfolio default rate", f"='Pilot districts'!B{PA['pdef']}", PCT, "%", "From 'Pilot districts'"),
    ("conv", "Appraisal -> disbursal conversion rate", f"='Pilot districts'!B{PA['pconv']}", PCT, "%", "From 'Pilot districts'"),
    ("redeploy", "Redeployment mode (0/1/2)", f"='Pilot districts'!B{PA['predeploy']}", "0", "",
     "0 = hold cash, 1 = same FI redeploys, 2 = 2nd FI in district (same maths)"),
], ncols=4)

hdr = TL["redeploy"] + 2
header(dt, hdr, ["Week", "New CMGs appraised", "New CMGs disbursed", "Capital disbursed this wk",
                 "Cumulative capital disbursed", "Capital remaining in pool", "Principal due",
                 "Principal repaid", "Interest collected", "Recycled capital", "Portfolio Outstanding",
                 "Pool exhausted?"])
dt.row_dimensions[hdr].height = 40
WF, WL = hdr + 1, hdr + PILOT_WEEKS
Bc, Cc, Dc, Ec, Fc, Gc, Hc, Ic, Jc, Kc, Lc2 = "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"
for w in range(1, PILOT_WEEKS + 1):
    r = WF + w - 1
    dt[f"A{r}"] = w
    prevB = "0" if w == 1 else f"SUM($B${WF}:B{r-1})"
    dt[f"{Bc}{r}"] = f"=MAX(0,MIN($B${TL['cap']},$B${TL['cmgs']}-{prevB}))"
    proposed = f'IF(A{r}<=$B${TL["lag"]},0,N(OFFSET(B{r},-$B${TL["lag"]},0))*$B${TL["conv"]})'
    prevF = f"$B${TL['alloc']}" if w == 1 else f"F{r-1}"
    dt[f"{Cc}{r}"] = f"=MIN({proposed},IF($B${TL['lpc']}=0,0,{prevF}/$B${TL['lpc']}))"
    dt[f"{Dc}{r}"] = f"=C{r}*$B${TL['lpc']}"
    dt[f"{Ec}{r}"] = f"=D{r}" if w == 1 else f"=E{r-1}+D{r}"
    dt[f"{Fc}{r}"] = f"=$B${TL['alloc']}-E{r}+SUM($J${WF}:J{r})"
    if w == 1:
        dt[f"{Gc}{r}"] = "=0"
    else:
        dt[f"{Gc}{r}"] = f'=SUMIFS($D${WF}:D{r-1},$A${WF}:A{r-1},">="&(A{r}-$B${TL["term"]}))/$B${TL["term"]}'
    dt[f"{Hc}{r}"] = f"=G{r}*(1-$B${TL['def']})"
    dt[f"{Ic}{r}"] = f"=H{r}*$B${TL['rate']}/52*$B${TL['term']}"
    dt[f"{Jc}{r}"] = f"=IF($B${TL['redeploy']}>0,H{r},0)"
    dt[f"{Kc}{r}"] = f"=E{r}-SUM($H${WF}:H{r})"
    dt[f"{Lc2}{r}"] = f'=IF(F{r}<$B${TL["lpc"]},"Yes","No")'
    for c in "BC":
        dt[f"{c}{r}"].number_format = NUM
    for c in "DEFGHIJK":
        dt[f"{c}{r}"].number_format = USD
sumr = WL + 2
section(dt, sumr - 1, "SUMMARY", 4)
summary = [
    ("Weeks to exhaust the pool", f'=IFERROR(INDEX($A${WF}:$A${WL},MATCH("Yes",$L${WF}:$L${WL},0)),"Not within {PILOT_WEEKS} weeks")', "@"),
    ("Peak Portfolio Outstanding", f"=MAX($K${WF}:$K${WL})", USD),
    ("Week of peak Portfolio Outstanding", f"=INDEX($A${WF}:$A${WL},MATCH(MAX($K${WF}:$K${WL}),$K${WF}:$K${WL},0))", "0"),
    ("Total CMGs served", f"=SUM($C${WF}:C${WL})", NUM),
    ("Total capital disbursed", f"=E{WL}", USD),
    ("Total capital recycled", f"=SUM($J${WF}:J{WL})", USD),
]
for i, (lab, f, fmt) in enumerate(summary, sumr):
    dt[f"A{i}"], dt[f"B{i}"] = lab, f
    dt[f"B{i}"].number_format, dt[f"B{i}"].font = fmt, BOLD
dt[f"A{i+2}"] = ("To compare with vs without recycling: change the Redeployment mode on 'Pilot districts' "
                 "(0 = off) and re-check 'Weeks to exhaust the pool' and 'Total capital disbursed' above.")
dt.column_dimensions["A"].width = 34
for c in range(2, 13):
    dt.column_dimensions[col(c)].width = 15
dt.freeze_panes = f"B{WF}"

lchart = LineChart()
lchart.title, lchart.y_axis.title, lchart.x_axis.title = "Weekly cascade for the selected district", "USD", "Week"
lchart.y_axis.numFmt = USD
cats = Reference(dt, min_col=1, min_row=WF, max_row=WL)
data = Reference(dt, min_col=6, max_col=6, min_row=hdr, max_row=WL)
lchart.add_data(data, titles_from_data=True)
data2 = Reference(dt, min_col=11, max_col=11, min_row=hdr, max_row=WL)
lchart.add_data(data2, titles_from_data=True)
lchart.set_categories(cats)
lchart.width, lchart.height = 20, 10
dt.add_chart(lchart, f"A{i+4}")

# ---------------------------------------------------------------- Corpus rollout
cr = wb.create_sheet("Corpus rollout")
title(cr, "Corpus rollout: full allocation upfront vs staggered tranches",
      "All 8 pilot districts side by side. Does spending most of the pool on early districts strand later ones?")
header(cr, 4, ["Item", "Value", "Unit", "Note"])
CRA = table(cr, 5, [
    (None, "ROLLOUT ASSUMPTIONS", None, None, None, None),
    ("tranche", "Tranche advanced to each district upfront (rest held as TAMISEMI reserve)", 0.70, PCT, "%",
     "100% = the whole allocation goes out on day one, like the annual model"),
    ("need", "New region funding need (illustrative)", 1_500_000, USD, "USD",
     "Just to test whether the reserve would cover a call for funds from elsewhere"),
])
DFIRST = CRA["need"] + 3
header(cr, DFIRST - 1, ["District", "Full allocation (USD)", "Committed tranche (USD)", "Reserve contribution (USD)",
                        "Rollout start week", "Combined weekly capacity", "Loan/CMG (USD)", "Weeks to exhaust tranche"])
cr.row_dimensions[DFIRST - 1].height = 40
DLAST = DFIRST + len(PILOT_DISTRICTS) - 1
DTOTAL = DLAST + 1
for i in range(len(PILOT_DISTRICTS)):
    r, pr = DFIRST + i, PFIRST + i
    cr[f"A{r}"] = f"='Pilot districts'!A{pr}"
    cr[f"B{r}"] = f"='Pilot districts'!H{pr}"
    cr[f"C{r}"] = f"=B{r}*$B${CRA['tranche']}"
    cr[f"D{r}"] = f"=B{r}-C{r}"
    cr[f"E{r}"] = f"='Pilot districts'!M{pr}"
    cr[f"F{r}"] = f"='Pilot districts'!K{pr}"
    cr[f"G{r}"] = f"='Pilot districts'!F{pr}"
    cr[f"H{r}"] = f"=IF(F{r}*G{r}=0,0,ROUNDUP(C{r}/(F{r}*G{r}),0))"
    for c, f in ((2, USD), (3, USD), (4, USD), (5, "0"), (6, NUM), (7, USD), (8, "0")):
        cr.cell(row=r, column=c).number_format = f
cr[f"A{DTOTAL}"] = "TOTAL"
cr[f"A{DTOTAL}"].font = BOLD
for c, f in ((2, USD), (3, USD), (4, USD)):
    L = col(c)
    cell = cr.cell(row=DTOTAL, column=c, value=f"=SUM({L}{DFIRST}:{L}{DLAST})")
    cell.number_format, cell.font = f, BOLD
cr[f"A{DTOTAL+2}"] = "Reserve covers the new region funding need?"
cr[f"B{DTOTAL+2}"] = f'=IF(D{DTOTAL}>=B{CRA["need"]},"Yes","No")'
cr[f"B{DTOTAL+2}"].font = BOLD

GRID = DTOTAL + 4
section(cr, GRID - 1, "WEEKLY CASCADE - capital disbursed to CMGs, by district (capacity-limited, no recycling)", 12)
header(cr, GRID, ["Week"] + [""] * len(PILOT_DISTRICTS) + ["Total disbursed", "Idle cash at TAMISEMI"])
# fill the per-district header cells with formulas pulling the live district names
for i in range(len(PILOT_DISTRICTS)):
    c = cr.cell(row=GRID, column=2 + i, value=f"=A{DFIRST+i}")
    c.fill, c.font, c.border = HEAD, WHITE, BOX
GWF, GWL = GRID + 1, GRID + PILOT_WEEKS
for w in range(1, PILOT_WEEKS + 1):
    r = GWF + w - 1
    cr[f"A{r}"] = w
    for i in range(len(PILOT_DISTRICTS)):
        dc, drow = col(2 + i), DFIRST + i
        cr[f"{dc}{r}"] = (f"=IF(A{r}<$E${drow},0,MIN($C${drow},($F${drow}*$G${drow})*(A{r}-$E${drow}+1)))")
        cr[f"{dc}{r}"].number_format = USD
    lastdc = col(2 + len(PILOT_DISTRICTS) - 1)
    cr[f"{col(2+len(PILOT_DISTRICTS))}{r}"] = f"=SUM(B{r}:{lastdc}{r})"
    cr[f"{col(2+len(PILOT_DISTRICTS))}{r}"].number_format = USD
    cr[f"{col(3+len(PILOT_DISTRICTS))}{r}"] = (
        f"='Pilot districts'!$B${PA['pcorpus']}-SUMPRODUCT(($E${DFIRST}:$E${DLAST}<=A{r})*$C${DFIRST}:$C${DLAST})")
    cr[f"{col(3+len(PILOT_DISTRICTS))}{r}"].number_format = USD
cr.column_dimensions["A"].width = 30
for c in range(2, 4 + len(PILOT_DISTRICTS)):
    cr.column_dimensions[col(c)].width = 15
cr.freeze_panes = f"B{GWF}"

lchart2 = LineChart()
lchart2.title, lchart2.y_axis.title, lchart2.x_axis.title = "Capital disbursed by district", "USD", "Week"
lchart2.y_axis.numFmt = USD
cats2 = Reference(cr, min_col=1, min_row=GWF, max_row=GWL)
data3 = Reference(cr, min_col=2, max_col=1 + len(PILOT_DISTRICTS), min_row=GRID, max_row=GWL)
lchart2.add_data(data3, titles_from_data=True)
lchart2.set_categories(cats2)
lchart2.width, lchart2.height = 22, 10
cr.add_chart(lchart2, f"A{GWL+2}")

lchart3 = LineChart()
lchart3.title, lchart3.y_axis.title, lchart3.x_axis.title = "Idle cash sitting at TAMISEMI", "USD", "Week"
lchart3.y_axis.numFmt = USD
data4 = Reference(cr, min_col=3 + len(PILOT_DISTRICTS), max_col=3 + len(PILOT_DISTRICTS), min_row=GRID, max_row=GWL)
lchart3.add_data(data4, titles_from_data=True)
lchart3.set_categories(cats2)
lchart3.width, lchart3.height = 22, 10
cr.add_chart(lchart3, f"L{GWL+2}")

# ---------------------------------------------------------------- Open questions sheet
oq = wb.create_sheet("Open questions")
title(oq, "Open questions", "To settle before the model uses real figures.")
header(oq, 4, ["#", "Question", "Why it matters", "Added in", "Answer"])
for r, (n, q, why, v) in enumerate(QUESTIONS, 5):
    for c, val in enumerate((n, q, why, v, ""), 1):
        cell = oq.cell(row=r, column=c, value=val)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    oq.cell(row=r, column=5).fill = INPUT
for Lc, w in zip("ABCDE", (4, 60, 55, 9, 40)):
    oq.column_dimensions[Lc].width = w

# ---------------------------------------------------------------- Glossary
gl = wb.create_sheet("Glossary")
title(gl, "Glossary", "Acronyms and terms used in this model.")
header(gl, 4, ["Term", "Meaning"])
terms = [
    ("Advance", "Money TAMISEMI sends a lender ahead of lending, based on the CMG pool allocated to that lender."),
    ("Allocation", "The share of the lending corpus given to a district, based on its loan demand."),
    ("Annual cycle", "The CMG's savings year. Loans must be repaid before the share-out at the end."),
    ("Appraisal", "A lender's check of a CMG before deciding whether to lend, and how much. Capacity-limited on 'District timeline' (CMGs a lender can appraise per week)."),
    ("Break-even default rate", "The highest default rate the default budget and TAMISEMI's risk share can cover and still restore the corpus."),
    ("CMG", "Community Microfinance Group: a savings group whose members save and lend to each other."),
    ("Committed tranche", "On 'Corpus rollout': the share of a district's allocation advanced upfront, per the tranche % input."),
    ("Conversion rate", "On the pilot-district sheets: the share of appraised CMGs that are approved and go on to accept the loan."),
    ("Corpus", "The pool of money available to lend ($24m less the $4m set aside for incentives). The pilot-district sheets use a separate, smaller placeholder pool."),
    ("Default", "A loan, or part of a loan, that is not repaid."),
    ("Default budget / loss reserve", "The share of interest set aside to cover TAMISEMI's risk share of defaults."),
    ("Demand", "Number of CMGs x loan per CMG, before any scaling down."),
    ("Flat interest", "Interest worked out on the full loan amount for the whole term, not on the falling balance."),
    ("Incentive", "A payment to reward good results (lenders: lending well; CMGs: on-time repayment)."),
    ("Instalment", "One of the equal monthly repayments of a loan."),
    ("Lender margin", "The share of interest the lender keeps after the default budget and the TAMISEMI/region/LGA shares are paid, before its risk share of defaults."),
    ("Lender's net result", "Lender margin, less the lender's own risk share of defaults, plus any repaid principal the lender keeps instead of returning to TAMISEMI."),
    ("LGA", "Local Government Authority: the district or council."),
    ("Loan:Savings ratio", "Loan amount divided by the CMG's savings. Capped by the maximum ratio input."),
    ("Portfolio default rate", "Defaulted principal as a % of all loans made."),
    ("Portfolio Outstanding", "The running loan book: capital disbursed so far, less principal repaid so far (defaulted principal is not written off in this simple version)."),
    ("Principal", "The amount lent, not counting interest."),
    ("Redeployment mode", "On the pilot-district sheets: whether repayments are held as cash (0), recycled by the same FI (1), or handed to a 2nd FI in the district (2) - modelled the same way, reported separately."),
    ("Re-lending / revolving fund", "Lending repayments out again instead of holding them as cash."),
    ("Reserve (TAMISEMI)", "On 'Corpus rollout': the share of the pilot pool not yet advanced to any district, held back for later needs."),
    ("Risk share", "How the loss on a defaulted loan is split between TAMISEMI (absorbed against the corpus) and the lender (its own loss, repaid to the fund as compensation)."),
    ("Rollout start week", "The week a pilot district begins appraising CMGs, on the pilot-district sheets."),
    ("Run-off", "The period after lending stops, while loans already made are being repaid."),
    ("Scale factor", "The % of demand that can be met when demand is more than the corpus."),
    ("Service fee", "The Region and LGA shares of interest collected (previously called Region/LGA share)."),
    ("Share-out", "When the CMG shares its savings and profits among members at the end of the cycle."),
    ("TAMISEMI", "President's Office - Regional Administration and Local Government (Tanzania). Funds PAMOJA."),
    ("TAMISEMI share of repaid principal", "The % of every repaid instalment's principal that returns to the TAMISEMI corpus; the rest is kept by the lender."),
    ("Tenure / term", "How long a loan runs, in months, from disbursement to the last instalment."),
    ("Turnover", "Total loans made in a year divided by the corpus: how many times the money was lent."),
    ("TZS / USD", "Tanzanian shilling / United States dollar."),
]
for r, (a, b) in enumerate(terms, 5):
    gl.cell(row=r, column=1, value=a).font = BOLD
    gl.cell(row=r, column=2, value=b)
gl.column_dimensions["A"].width = 30
gl.column_dimensions["B"].width = 100

wb.save(OUT)
print("saved", OUT)

# Keep a plain-text copy of the open questions next to the model
with open("model/OPEN_QUESTIONS.md", "w") as fh:
    fh.write("# Open questions (PAMOJA financial model)\n\n"
             "Generated by `build_model.py`; also on the 'Open questions' sheet.\n\n"
             "| # | Question | Why it matters | Added in |\n|---|---|---|---|\n")
    for n, q, why, v in QUESTIONS:
        fh.write(f"| {n} | {q} | {why} | {v} |\n")
