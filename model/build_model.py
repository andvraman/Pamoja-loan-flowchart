"""Builds PAMOJA_Financial_Model_v0.1.xlsx (simple fund-flow model).
Run: python3 model/build_model.py   (needs openpyxl)
All numbers in blue on yellow are inputs. Everything else is a formula."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as col

OUT = "model/PAMOJA_Financial_Model_v0.1.xlsx"
wb = Workbook()

INPUT = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor="1C5A3F")
SUB = PatternFill("solid", fgColor="DCEBE1")
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


# ---------------------------------------------------------------- Read me
rm = wb.active
rm.title = "Read me"
title(rm, "PAMOJA simplified financial model - v0.1 (draft)",
      "How $24m flows TAMISEMI -> Lenders -> CMGs as loans, and back as repayments.")
lines = [
    "",
    "HOW TO USE",
    "1. Change the yellow cells with blue numbers only. Everything else is a formula.",
    "2. Inputs: fund size and assumptions. Districts: CMG numbers, savings and loan size per LGA.",
    "3. Fund flow: one annual cycle as a waterfall, with a check that the corpus is restored.",
    "4. Monthly cycle: when loans go out and come back over 12 months.",
    "",
    "IMPORTANT",
    "- All district data and rates are DUMMY placeholders. Replace them with real figures.",
    "- The split of interest (default budget, discounts, TAMISEMI, region, LGA) is a first guess to confirm.",
    "",
    "SIMPLE LOGIC (v0.1)",
    "- Lending corpus = Total fund - Lender incentives - CMG incentives.",
    "- Loan per CMG = the smaller of: average loan asked for, or savings x maximum Loan:Savings ratio.",
    "- Demand per district = number of CMGs x loan per CMG.",
    "- If total demand is more than the corpus, every district is scaled down by the same %.",
    "- Interest is flat: rate x loan x term. Defaulted loans pay no principal and no interest.",
    "- The corpus is restored if the default budget taken from interest covers the defaulted principal.",
    "",
    "NOT YET IN THE MODEL (next steps)",
    "- Re-lending repayments within the same cycle (revolving fund).",
    "- Multiple lenders per LGA; different terms per district; region-level roll-up.",
    "- Timing of incentive payments; LGA service fee formula; risk-share formula (Q12).",
    "- Exchange rate (TZS / USD) and inflation.",
    "",
    "See the 'Glossary' sheet for all acronyms and terms.",
]
for i, t in enumerate(lines, 4):
    rm.cell(row=i, column=1, value=t)
    if t.isupper() and t:
        rm.cell(row=i, column=1).font = BOLD
rm.column_dimensions["A"].width = 110

# ---------------------------------------------------------------- Inputs
ip = wb.create_sheet("Inputs")
title(ip, "Inputs", "Yellow cells are inputs. Dummy values - replace with agreed figures.")
header(ip, 4, ["Item", "Value", "Unit", "Note"])
rows = [
    ("FUND", None, None, None),
    ("Total PAMOJA fund", 24_000_000, USD, "From TAMISEMI"),
    ("Lender incentives (set aside)", 2_000_000, USD, "Not lent out"),
    ("CMG incentives (set aside)", 2_000_000, USD, "Not lent out"),
    ("Lending corpus", "=B6-B7-B8", USD, "Formula: money available to lend"),
    ("LENDING RULES", None, None, None),
    ("Maximum Loan:Savings ratio", 0.5, "0.00", "Loan must not exceed this x CMG savings"),
    ("Interest rate (flat, per year)", 0.18, PCT, "Placeholder; must be at or below TAMISEMI ceiling"),
    ("Loan term", 6, "0", "Months. Must end before the CMG share-out"),
    ("Portfolio default rate (budget)", 0.05, PCT, "% of principal lent that is not repaid"),
    ("SPLIT OF INTEREST COLLECTED", None, None, None),
    ("Default budget (loss reserve)", 0.30, PCT, "Returned to fund to cover defaults"),
    ("Lender discount", 0.15, PCT, "Placeholder"),
    ("TAMISEMI share", 0.10, PCT, "Placeholder"),
    ("Region share", 0.05, PCT, "Placeholder"),
    ("LGA share", 0.05, PCT, "Placeholder"),
    ("Lender net margin (remainder)", "=1-SUM(B16:B20)", PCT, "Formula: what is left for the lender"),
    ("Check: split adds to 100% or less", '=IF(B21>=0,"OK","ERROR: over 100%")', "@", ""),
]
for r, (a, b, fmt, n) in enumerate(rows, 5):
    ip.cell(row=r, column=1, value=a)
    if b is None:
        ip.cell(row=r, column=1).font = BOLD
        for c in range(1, 5):
            ip.cell(row=r, column=c).fill = SUB
        continue
    c = ip.cell(row=r, column=2, value=b)
    if isinstance(b, str) and b.startswith("="):
        c.number_format, c.border, c.font = fmt, BOX, BOLD
    else:
        inp(c, fmt)
    ip.cell(row=r, column=3, value={USD: "USD", PCT: "%"}.get(fmt, ""))
    ip.cell(row=r, column=4, value=n)
ip["C13"] = "months"
ip["C11"] = "x"
for L, w in zip("ABCD", (38, 16, 8, 55)):
    ip.column_dimensions[L].width = w

# Named refs used below
CORPUS, RATIO, RATE, TERM, DEF = "Inputs!$B$9", "Inputs!$B$11", "Inputs!$B$12", "Inputs!$B$13", "Inputs!$B$14"

# ---------------------------------------------------------------- Districts
ds = wb.create_sheet("Districts")
title(ds, "Allocation by district (LGA)",
      "Dummy data. Add real LGAs in the yellow cells; up to 20 rows.")
header(ds, 4, ["District (LGA)", "Number of CMGs", "Average savings per CMG (USD)",
               "Average loan asked per CMG (USD)", "Loan cap per CMG = savings x ratio",
               "Loan per CMG (smaller of the two)", "Loan:Savings ratio used",
               "Demand (USD)", "Allocation (USD)", "Share of corpus"])
ds.row_dimensions[4].height = 45
sample = [("District A", 2500, 4000, 2500), ("District B", 1800, 3000, 2000),
          ("District C", 3200, 5000, 3000), ("District D", 1500, 3500, 1500),
          ("District E", 2000, 4500, 2500)]
FIRST, LAST = 5, 24
for r in range(FIRST, LAST + 1):
    data = sample[r - FIRST] if r - FIRST < len(sample) else (None, None, None, None)
    ds.cell(row=r, column=1, value=data[0]).fill = INPUT
    for c, v, f in ((2, data[1], NUM), (3, data[2], USD), (4, data[3], USD)):
        inp(ds.cell(row=r, column=c, value=v), f)
    ds.cell(row=r, column=5, value=f'=IF(B{r}="","",C{r}*{RATIO})').number_format = USD
    ds.cell(row=r, column=6, value=f'=IF(B{r}="","",MIN(D{r},E{r}))').number_format = USD
    ds.cell(row=r, column=7, value=f'=IF(B{r}="","",F{r}/C{r})').number_format = "0.00"
    ds.cell(row=r, column=8, value=f'=IF(B{r}="",0,B{r}*F{r})').number_format = USD
    ds.cell(row=r, column=9, value=f'=IF(B{r}="",0,H{r}*$B$28)').number_format = USD
    ds.cell(row=r, column=10, value=f'=IF(B{r}="","",I{r}/{CORPUS})').number_format = PCT
T = LAST + 1
ds.cell(row=T, column=1, value="TOTAL").font = BOLD
for c, f in ((2, NUM), (8, USD), (9, USD), (10, PCT)):
    L = col(c)
    cell = ds.cell(row=T, column=c, value=f"=SUM({L}{FIRST}:{L}{LAST})")
    cell.number_format, cell.font = f, BOLD
ds["A27"], ds["B27"] = "Lending corpus", f"={CORPUS}"
ds["A28"], ds["B28"] = "Scale factor (1 = full demand met)", f"=IF(H{T}=0,0,MIN(1,B27/H{T}))"
ds["A29"], ds["B29"] = "Corpus not used", f"=B27-I{T}"
ds["A30"], ds["B30"] = "Demand not met", f"=H{T}-I{T}"
for a, f in (("B27", USD), ("B28", "0.00"), ("B29", USD), ("B30", USD)):
    ds[a].number_format, ds[a].font = f, BOLD
ds.column_dimensions["A"].width = 34
for c in range(2, 11):
    ds.column_dimensions[col(c)].width = 17
ds.freeze_panes = "B5"

# ---------------------------------------------------------------- Fund flow
ff = wb.create_sheet("Fund flow")
title(ff, "Fund flow for one annual cycle",
      "TAMISEMI -> Lenders -> CMGs -> back. All formulas.")
header(ff, 4, ["Step", "Amount (USD)", "From -> To", "How it is worked out"])
L = f"Districts!$I${T}"
flow = [
    ("1. MONEY OUT", None, None, None),
    ("Total PAMOJA fund", "=Inputs!B6", "TAMISEMI", ""),
    ("Lender incentives set aside", "=Inputs!B7", "TAMISEMI -> Lenders", "Paid later, not lent"),
    ("CMG incentives set aside", "=Inputs!B8", "TAMISEMI -> CMGs", "Paid later, not lent"),
    ("Advance to lenders (= loans made)", f"={L}", "TAMISEMI -> Lenders", "Sum of district allocations"),
    ("Corpus not lent (returned)", "=Inputs!B9-B9", "Lenders -> TAMISEMI", "Unused advance"),
    ("Loans to CMGs", "=B9", "Lenders -> CMGs", "Advance lent in full"),
    ("2. MONEY BACK FROM CMGs", None, None, None),
    ("Principal defaulted", f"=B11*{DEF}", "", "Loans x default rate"),
    ("Principal repaid", "=B11-B13", "CMGs -> Lenders", "Loans - defaults"),
    ("Interest collected", f"=B14*{RATE}*{TERM}/12", "CMGs -> Lenders", "Flat: repaid loans x rate x term/12"),
    ("Total repayments", "=B14+B15", "CMGs -> Lenders", ""),
    ("3. SPLIT OF INTEREST", None, None, None),
    ("Default budget (loss reserve)", "=B15*Inputs!B16", "Lenders -> TAMISEMI", ""),
    ("Lender discount", "=B15*Inputs!B17", "Kept by lenders", ""),
    ("TAMISEMI share", "=B15*Inputs!B18", "Lenders -> TAMISEMI", ""),
    ("Region share", "=B15*Inputs!B19", "-> Region", ""),
    ("LGA share", "=B15*Inputs!B20", "-> LGA", ""),
    ("Lender net margin", "=B15*Inputs!B21", "Kept by lenders", "Remainder"),
    ("4. CORPUS CHECK (end of cycle)", None, None, None),
    ("Corpus at start", "=Inputs!B9", "", ""),
    ("Principal repaid to TAMISEMI", "=B14", "Lenders -> TAMISEMI", ""),
    ("Corpus not lent", "=B10", "", ""),
    ("Default budget added back", "=B18", "", "Covers the defaulted principal"),
    ("Corpus at end", "=B26+B27+B28", "", ""),
    ("Gap (end - start)", "=B29-B25", "", "Negative = corpus NOT restored"),
    ("Result", '=IF(B30>=-0.5,"Corpus restored","Corpus NOT restored - raise interest share for defaults or rate")', "", ""),
    ("Default budget needed (% of interest)", "=IF(B15=0,0,B13/B15)", "", "Minimum share of interest to cover defaults"),
    ("Break-even default rate", f"=Inputs!B16*{RATE}*{TERM}/12/(1+Inputs!B16*{RATE}*{TERM}/12)", "", "Highest default rate the current split can cover"),
]
for r, (a, b, ft, n) in enumerate(flow, 5):
    ff.cell(row=r, column=1, value=a)
    if b is None:
        ff.cell(row=r, column=1).font = BOLD
        for c in range(1, 5):
            ff.cell(row=r, column=c).fill = SUB
        continue
    c = ff.cell(row=r, column=2, value=b)
    c.number_format, c.border = USD, BOX
    ff.cell(row=r, column=3, value=ft)
    ff.cell(row=r, column=4, value=n)
ff["B32"].number_format = PCT
ff["B33"].number_format = PCT
for a in ("A29", "B29", "A31", "B31"):
    ff[a].font = BOLD
for Lc, w in zip("ABCD", (40, 18, 24, 48)):
    ff.column_dimensions[Lc].width = w

# ---------------------------------------------------------------- Monthly
mo = wb.create_sheet("Monthly cycle")
title(mo, "Monthly cycle (12 months)",
      "Enter the % of total loans disbursed each month (yellow). Repayments are equal monthly instalments over the loan term.")
header(mo, 4, ["Line"] + [f"M{m}" for m in range(1, 13)] + ["Total"])
labels = ["Month number", "% of loans disbursed this month", "Loans disbursed",
          "Principal due", "Principal repaid (after defaults)", "Interest collected",
          "Loans outstanding (end of month)", "Fund cash on hand (end of month)",
          "Warning: loan still running after month 12"]
for i, t in enumerate(labels, 5):
    mo.cell(row=i, column=1, value=t)
pattern = [0.4, 0.3, 0.2, 0.1] + [0] * 8
for m in range(1, 13):
    c = col(m + 1)
    mo[f"{c}5"] = m
    inp(mo[f"{c}6"], PCT)
    mo[f"{c}6"] = pattern[m - 1]
    mo[f"{c}7"] = f"={c}6*Districts!$I${T}"
    # principal due in month m from loans made in earlier months d where 1 <= m-d <= term
    rng = "$B$5:$M$5"
    mo[f"{c}8"] = (f"=SUMPRODUCT(({c}$5-{rng}>=1)*({c}$5-{rng}<={TERM})*$B$7:$M$7)/{TERM}")
    mo[f"{c}9"] = f"={c}8*(1-{DEF})"
    mo[f"{c}10"] = f"={c}9*{RATE}/12*{TERM}"
    prev = "0" if m == 1 else f"{col(m)}11"
    mo[f"{c}11"] = f"={prev}+{c}7-{c}8"
    prevcash = "Inputs!$B$9" if m == 1 else f"{col(m)}12"
    mo[f"{c}12"] = f"={prevcash}-{c}7+{c}9+{c}10*Inputs!$B$16"
    for rr in (7, 8, 9, 10, 11, 12):
        mo[f"{c}{rr}"].number_format = USD
    mo[f"{c}5"].font = BOLD
mo["N6"] = "=SUM(B6:M6)"; mo["N6"].number_format = PCT
for rr in (7, 8, 9, 10):
    mo[f"N{rr}"] = f"=SUM(B{rr}:M{rr})"; mo[f"N{rr}"].number_format = USD
mo["B13"] = f'=IF(SUMPRODUCT(($B$5:$M$5+{TERM}>12)*($B$7:$M$7))>0,"YES - some loans end after the cycle. Move disbursement earlier or shorten the term.","No - all loans end within the cycle")'
mo["A15"] = "Latest month a loan can go out and still be repaid by month 12:"
mo["B15"] = f"=12-{TERM}"
mo["A16"] = "Cash on hand at end of month 12 (should equal the corpus if restored):"
mo["B16"] = "=M12"; mo["B16"].number_format = USD
mo["A17"] = "Check: % disbursed adds to 100%"
mo["B17"] = '=IF(ABS(N6-1)<0.0001,"OK","ERROR: adjust row 6")'
mo["A19"] = ("Note: Fund cash counts principal repaid plus the default budget share of interest. "
             "Other interest shares go to lenders, TAMISEMI, region and LGA.")
for a in ("A13", "A15", "A16", "A17", "B15", "B16", "B17", "B13"):
    mo[a].font = BOLD
mo.column_dimensions["A"].width = 44
for m in range(2, 15):
    mo.column_dimensions[col(m)].width = 13
mo.freeze_panes = "B5"

# ---------------------------------------------------------------- Glossary
gl = wb.create_sheet("Glossary")
title(gl, "Glossary", "Acronyms and terms used in this model.")
header(gl, 4, ["Term", "Meaning"])
terms = [
    ("Advance", "Money TAMISEMI sends a lender ahead of lending, based on the CMG pool allocated to that lender."),
    ("Allocation", "The share of the lending corpus given to a district, based on its loan demand."),
    ("Annual cycle", "The CMG's savings year. Loans must be repaid before the share-out at the end."),
    ("Break-even default rate", "The highest default rate the default budget can cover and still restore the corpus."),
    ("CMG", "Community Microfinance Group: a savings group whose members save and lend to each other."),
    ("Corpus", "The pool of money available to lend ($24m less the $4m set aside for incentives)."),
    ("Default", "A loan, or part of a loan, that is not repaid."),
    ("Default budget / loss reserve", "The share of interest set aside to cover defaults."),
    ("Demand", "Number of CMGs x loan per CMG, before any scaling down."),
    ("Discount", "A share of interest given back or kept to reward a party (e.g. lender discount)."),
    ("Flat interest", "Interest worked out on the full loan amount for the whole term, not on the falling balance."),
    ("Incentive", "A payment to reward good results (lenders: lending well; CMGs: on-time repayment)."),
    ("LGA", "Local Government Authority: the district or council."),
    ("Loan:Savings ratio", "Loan amount divided by the CMG's savings. Capped by the maximum ratio input."),
    ("Portfolio default rate", "Defaulted principal as a % of all loans made."),
    ("Principal", "The amount lent, not counting interest."),
    ("Revolving fund", "A fund where repayments are lent out again. Not yet in v0.1."),
    ("Scale factor", "The % of demand that can be met when demand is more than the corpus."),
    ("Share-out", "When the CMG shares its savings and profits among members at the end of the cycle."),
    ("TAMISEMI", "President's Office - Regional Administration and Local Government (Tanzania). Funds PAMOJA."),
    ("TZS / USD", "Tanzanian shilling / United States dollar."),
]
for r, (a, b) in enumerate(terms, 5):
    gl.cell(row=r, column=1, value=a).font = BOLD
    gl.cell(row=r, column=2, value=b)
gl.column_dimensions["A"].width = 30
gl.column_dimensions["B"].width = 100

wb.save(OUT)
print("saved", OUT)
