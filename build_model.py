"""
Build a 10-year acquisition model for a hypothetical GTA multi-tenant industrial building.

Every output in the workbook is a live Excel formula driven by the blue input cells on the
Assumptions sheet, so the model can be re-underwritten by changing inputs in Excel.

Usage:  python build_model.py   ->  GTA_Industrial_Acquisition_Model.xlsx
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

OUT = "GTA_Industrial_Acquisition_Model.xlsx"
FONT = "Arial"
YEARS = 11  # 10-year max hold + 1 year for forward exit NOI

# ---------- styles ----------
BLUE, GREEN, BLACK, WHITE = "0000FF", "008000", "000000", "FFFFFF"
NAVY = "1F3864"
f_in = Font(name=FONT, color=BLUE)
f_link = Font(name=FONT, color=GREEN)
f_calc = Font(name=FONT, color=BLACK)
f_bold = Font(name=FONT, bold=True)
f_title = Font(name=FONT, bold=True, size=14, color=NAVY)
f_head = Font(name=FONT, bold=True, color=WHITE)
f_note = Font(name=FONT, italic=True, size=9, color="595959")
fill_head = PatternFill("solid", fgColor=NAVY)
fill_key = PatternFill("solid", fgColor="FFFF00")
fill_sub = PatternFill("solid", fgColor="D9E1F2")
fill_grey = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="808080")
top_border = Border(top=thin)
box = Border(top=thin, bottom=thin, left=thin, right=thin)

USD = '$#,##0;($#,##0);"-"'
USD2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
NUM = '#,##0;(#,##0);"-"'
MULT = '0.00"x"'
YR = "0"

wb = Workbook()
ws_sum = wb.active
ws_sum.title = "Summary"
ws_a = wb.create_sheet("Assumptions")
ws_cf = wb.create_sheet("Cash Flow")
ws_d = wb.create_sheet("Debt")
ws_s = wb.create_sheet("Sensitivity")
ws_sc = wb.create_sheet("Sens_Calc")
ws_m = wb.create_sheet("Market Data")

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2


def put(ws, ref, value, font=f_calc, fmt=None, fill=None, bold=False, align=None, border=None):
    c = ws[ref]
    c.value = value
    c.font = Font(name=FONT, color=font.color, bold=bold or font.bold, italic=font.italic,
                  size=font.size)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = Alignment(horizontal=align)
    if border:
        c.border = border
    return c


def header_bar(ws, row, c1, c2, text):
    for col in range(c1, c2 + 1):
        ws.cell(row=row, column=col).fill = fill_head
    put(ws, f"{L(c1)}{row}", text, f_head)


# =====================================================================================
# ASSUMPTIONS
# =====================================================================================
A = {}  # name -> absolute reference string, e.g. "Assumptions!$C$13"
ws_a.column_dimensions["B"].width = 44
ws_a.column_dimensions["C"].width = 16
ws_a.column_dimensions["D"].width = 70
for col in "EFGHIJKLM":
    ws_a.column_dimensions[col].width = 11

put(ws_a, "B2", "Assumptions: GTA Multi-Tenant Industrial Acquisition", f_title)
put(ws_a, "B3", "Blue = hard-coded input (edit these). Black = formula. Yellow fill = key "
                "assumption. Green = link from another sheet. Sources on 'Market Data'.", f_note)


def section(row, text):
    header_bar(ws_a, row, 2, 4, text)


def inp(row, name, label, value, fmt, note="", key=False, formula=False):
    put(ws_a, f"B{row}", label)
    put(ws_a, f"C{row}", value, f_calc if formula else f_in, fmt, fill_key if key else None,
        align="right")
    if note:
        put(ws_a, f"D{row}", note, f_note).alignment = Alignment(indent=1)
    A[name] = f"Assumptions!$C${row}"


section(5, "Property")
put(ws_a, "B6", "Property")
put(ws_a, "C6", "Hypothetical 200,000 SF multi-tenant industrial building, Brampton, ON", f_in)
inp(7, "gla", "Gross leasable area (SF)", "=SUM($D$58:$D$61)", NUM,
    "Sum of the rent roll below", formula=True)
inp(8, "clear", "Clear height (ft)", 28, "0", "Class B product (30 ft+ is modern Class A)")
inp(9, "built", "Year built", 2005, YR)
inp(10, "start", "Analysis start year (Year 1)", 2027, YR, "Assumes closing on Jan 1 of Year 1")
inp(11, "hold", "Hold period (years, 1 to 10)", 10, "0",
    "Sale at end of this year, priced on the following year's NOI", key=True)

section(13, "Acquisition")
inp(14, "price", "Purchase price ($)", 54000000, USD,
    "About $270/SF; below CBRE GTA average asking price ($374.53/SF) for older Class B stock",
    key=True)
inp(15, "cc", "Closing costs (% of price)", 0.0275, PCT2,
    "Ontario land transfer tax (~2% top rate) plus legal, due diligence, financing fees")
inp(16, "ppsf", "Price per SF ($)", f"={A['price']}/{A['gla']}", USD2, formula=True)
inp(17, "goingin", "Going-in cap rate (Year 1 NOI / price)", "='Cash Flow'!$D$40/" + A["price"],
    PCT2, "Compare: CBRE Toronto industrial Class B 5.25% to 6.00% (Q3 2026)", formula=True)

section(19, "Market Leasing")
inp(20, "mr", "Market net rent, Year 1 ($/SF/yr)", 15.75, USD2,
    "GTA avg asking net rent $16.06 (Q3 2026); discounted for 28 ft clear", key=True)
inp(21, "renew", "Renewal probability", 0.65, PCT, "Industry-standard underwriting assumption")
inp(22, "down", "Downtime if tenant leaves (months)", 6, "0")
inp(23, "term", "New lease term (years)", 5, "0")
inp(24, "esc_new", "Annual escalation on new leases", 0.03, PCT)
inp(25, "ti_new", "Tenant improvements, new tenant ($/SF)", 5.00, USD2)
inp(26, "ti_ren", "Tenant improvements, renewal ($/SF)", 1.00, USD2)
inp(27, "lc_new", "Leasing commissions, new tenant ($/SF)", 6.00, USD2)
inp(28, "lc_ren", "Leasing commissions, renewal ($/SF)", 3.00, USD2)
inp(29, "lc_infl", "Leasing cost inflation (per year)", 0.025, PCT)
inp(30, "down_w", "Weighted downtime at rollover (months)",
    f"=(1-{A['renew']})*{A['down']}", "0.0", "(1 - renewal probability) x downtime",
    formula=True)
inp(31, "ti_w", "Weighted TI at rollover ($/SF)",
    f"={A['renew']}*{A['ti_ren']}+(1-{A['renew']})*{A['ti_new']}", USD2, formula=True)
inp(32, "lc_w", "Weighted leasing commissions at rollover ($/SF)",
    f"={A['renew']}*{A['lc_ren']}+(1-{A['renew']})*{A['lc_new']}", USD2, formula=True)

section(34, "Operating")
inp(35, "tmi", "Recoverable operating costs, TMI, Year 1 ($/SF)", 3.75, USD2,
    "GTA TMI averages $3.50 to $4.00/SF; recovered from tenants on net leases")
inp(36, "opex_g", "Operating cost growth (per year)", 0.03, PCT)
inp(37, "vac", "General vacancy and credit loss (% of PGR)", 0.02, PCT,
    "On top of modeled rollover downtime")
inp(38, "mgmt", "Non-recoverable management fee (% of EGI)", 0.01, PCT)
inp(39, "ga", "Non-recoverable G&A, Year 1 ($/SF)", 0.25, USD2)
inp(40, "reserve", "Capital reserve, Year 1 ($/SF)", 0.20, USD2,
    "Roof, paving, HVAC; grows with leasing cost inflation")

section(42, "Financing")
inp(43, "ltv", "Loan-to-value", 0.55, PCT, key=True)
inp(44, "loan", "Loan amount ($)", f"={A['ltv']}*{A['price']}", USD, formula=True)
inp(45, "rate", "Interest rate (fixed)", 0.055, PCT2,
    "5-yr GoC yield 3.59% (Oct 2026) plus ~190 bps lender spread", key=True)
inp(46, "amort", "Amortization (years)", 25, "0")
inp(47, "ds", "Annual debt service ($)",
    f"=PMT({A['rate']}/12,{A['amort']}*12,-{A['loan']})*12", USD,
    "Monthly payments, shown annually", formula=True)

section(49, "Exit")
inp(50, "exitcap", "Exit cap rate", 0.0575, PCT2,
    "25 bps above the going-in cap for an older building at sale", key=True)
inp(51, "sellc", "Selling costs (% of sale price)", 0.015, PCT)

# market rent growth by year
header_bar(ws_a, 53, 2, 13, "Market Rent Growth by Year")
put(ws_a, "B54", "Year", f_bold)
put(ws_a, "B55", "Market rent growth (from prior year)")
growth = [None, 0.005, 0.02] + [0.03] * 8
for i in range(YEARS):
    col = L(3 + i)
    put(ws_a, f"{col}54", i + 1, f_bold, "0", align="right")
    if growth[i] is None:
        put(ws_a, f"{col}55", "n/a", f_note, align="right")
    else:
        put(ws_a, f"{col}55", growth[i], f_in, PCT, align="right")
A["growth_row"] = "Assumptions!$C$55:$M$55"
put(ws_a, "B56", "Rents fell for 7 straight quarters to Q1 2026 and are near a floor, so "
                 "growth is held low for 2028-29 before returning to 3%.", f_note)

# rent roll
RR0 = 58  # first tenant row
header_bar(ws_a, 57, 2, 9, "")
heads = ["Tenant", "Industry", "SF", "Year 1 rent ($/SF)", "Annual escalation",
         "Lease expiry (end of year #)", "Expiry (calendar year)", "In-place vs market"]
for i, h in enumerate(heads):
    c = put(ws_a, f"{L(2 + i)}57", h, f_head, fill=fill_head)
    c.alignment = Alignment(horizontal="center", wrap_text=True)
ws_a.row_dimensions[57].height = 30
tenants = [
    ("Tenant A", "Third-party logistics", 80000, 13.25, 0.030, 2, "Signed 2021, before the rent spike: below market"),
    ("Tenant B", "Food distribution", 60000, 18.50, 0.030, 5, "Signed 2023 near the peak: above market"),
    ("Tenant C", "Auto parts", 35000, 15.00, 0.025, 4, ""),
    ("Tenant D", "E-commerce fulfillment", 25000, 17.25, 0.030, 7, "Signed 2024"),
]
for i, (t, ind, sf, rent, esc, exp, note) in enumerate(tenants):
    r = RR0 + i
    put(ws_a, f"B{r}", t, f_in)
    put(ws_a, f"C{r}", ind, f_in)
    put(ws_a, f"D{r}", sf, f_in, NUM)
    put(ws_a, f"E{r}", rent, f_in, USD2)
    put(ws_a, f"F{r}", esc, f_in, PCT)
    put(ws_a, f"G{r}", exp, f_in, "0", align="center")
    put(ws_a, f"H{r}", f"={A['start']}+G{r}-1", f_calc, YR, align="center")
    put(ws_a, f"I{r}", f"=E{r}/{A['mr']}-1", f_calc, PCT)
    if note:
        ws_a[f"B{r}"].comment = Comment(note, "Model")
RRn = RR0 + len(tenants) - 1
rt = RRn + 1
put(ws_a, f"B{rt}", "Total / weighted average", f_bold, border=top_border)
put(ws_a, f"D{rt}", f"=SUM(D{RR0}:D{RRn})", f_bold, NUM, border=top_border)
put(ws_a, f"E{rt}", f"=SUMPRODUCT(D{RR0}:D{RRn},E{RR0}:E{RRn})/D{rt}", f_bold, USD2,
    border=top_border)
put(ws_a, f"G{rt}", f"=SUMPRODUCT(D{RR0}:D{RRn},G{RR0}:G{RRn})/D{rt}", f_bold, "0.0",
    border=top_border, align="center")
put(ws_a, f"I{rt}", f"=E{rt}/{A['mr']}-1", f_bold, PCT, border=top_border)
put(ws_a, f"B{rt + 1}", "Total row shows weighted average lease term remaining (years) under "
                        "expiry. Hover over tenant names for lease notes.", f_note)
ws_a.freeze_panes = "B5"

# =====================================================================================
# CASH FLOW
# =====================================================================================
cf = ws_cf
cf.column_dimensions["B"].width = 42
cf.column_dimensions["C"].width = 14
for i in range(YEARS):
    cf.column_dimensions[L(4 + i)].width = 13
put(cf, "B2", "10-Year Cash Flow Projection ($)", f_title)
put(cf, "B3", "Year 11 is shown only to price the exit on forward NOI. Columns after the "
              "hold period drop out of the returns.", f_note)

YC = [L(4 + i) for i in range(YEARS)]  # D..N = years 1..11
FIRST, LAST = YC[0], YC[-1]
LAST10 = YC[9]  # M
put(cf, "B5", "Year", f_bold)
put(cf, "B6", "Calendar year", f_bold)
put(cf, "C5", 0, f_bold, "0", align="right")
put(cf, "C6", f"={A['start']}-1", f_bold, YR, align="right")
for i, col in enumerate(YC):
    put(cf, f"{col}5", i + 1 if i == 0 else f"={YC[i - 1]}5+1", f_bold, "0", align="right")
    put(cf, f"{col}6", f"={A['start']}+{col}$5-1", f_bold, YR, align="right")
for col in ["B", "C"] + YC:
    cf[f"{col}6"].border = Border(bottom=thin)
for col in ["B", "C"] + YC:
    cf[f"{col}6"].fill = fill_sub
    cf[f"{col}5"].fill = fill_sub


def row_label(r, text, bold=False, indent=0, fill=None):
    c = put(cf, f"B{r}", text, f_bold if bold else f_calc, fill=fill)
    c.alignment = Alignment(indent=indent)


def fill_row(r, fn, fmt=USD, bold=False, border=None, font=f_calc):
    for i, col in enumerate(YC):
        put(cf, f"{col}{r}", fn(col, i), font, fmt, bold=bold, border=border)


# market rent
row_label(8, "Market net rent ($/SF)", True)
fill_row(8, lambda c, i: f"={A['mr']}" if i == 0 else
         f"={YC[i - 1]}8*(1+INDEX({A['growth_row']},{c}$5))", USD2)
for i in range(1):
    cf[f"{YC[0]}8"].font = Font(name=FONT, color=GREEN)

# tenant blocks
R_PSF, R_FLAG = 10, 16
row_label(R_PSF - 1, "Contract rent by tenant ($/SF)", True)
row_label(R_FLAG - 1, "Rollover flag (1 = lease rolls this year)", True)
SFr, Rr, ESCr, EXr = [], [], [], []
for k in range(len(tenants)):
    rr = RR0 + k
    sf, r0, esc, ex = (f"Assumptions!$D${rr}", f"Assumptions!$E${rr}",
                       f"Assumptions!$F${rr}", f"Assumptions!$G${rr}")
    term = A["term"]
    row_label(R_PSF + k, f"=Assumptions!$B${rr}", indent=1)
    row_label(R_FLAG + k, f"=Assumptions!$B${rr}", indent=1)
    cf[f"B{R_PSF + k}"].font = Font(name=FONT, color=GREEN)
    cf[f"B{R_FLAG + k}"].font = Font(name=FONT, color=GREEN)

    def psf(c, i, sf=sf, r0=r0, esc=esc, ex=ex, term=term):
        start = f"({ex}+1+INT(({c}$5-{ex}-1)/{term})*{term})"
        return (f"=IF({c}$5<={ex},{r0}*(1+{esc})^({c}$5-1),"
                f"INDEX(${FIRST}$8:${LAST}$8,{start})*(1+{A['esc_new']})^({c}$5-{start}))")

    def flag(c, i, ex=ex, term=term):
        return f"=IF(AND({c}$5>{ex},MOD({c}$5-{ex}-1,{term})=0),1,0)"

    fill_row(R_PSF + k, psf, USD2)
    fill_row(R_FLAG + k, flag, "0")
SF_RANGE = f"Assumptions!$D${RR0}:$D${RRn}"

R = 21
row_label(R, "Revenue", True, fill=fill_sub)
for col in ["C"] + YC:
    cf[f"{col}{R}"].fill = fill_sub
psf_rng = lambda c: f"{c}{R_PSF}:{c}{R_PSF + len(tenants) - 1}"
flag_rng = lambda c: f"{c}{R_FLAG}:{c}{R_FLAG + len(tenants) - 1}"
occ = f"{A['down_w']}/12"

row_label(22, "Scheduled base rent", indent=1)
fill_row(22, lambda c, i: f"=SUMPRODUCT({SF_RANGE},{psf_rng(c)})")
row_label(23, "Less: rollover downtime", indent=1)
fill_row(23, lambda c, i: f"=-SUMPRODUCT({SF_RANGE},{psf_rng(c)},{flag_rng(c)})*{occ}")
row_label(24, "Operating cost recoveries", indent=1)
fill_row(24, lambda c, i: f"={A['tmi']}*(1+{A['opex_g']})^({c}$5-1)*"
                          f"(SUM({SF_RANGE})-SUMPRODUCT({SF_RANGE},{flag_rng(c)})*{occ})")
row_label(25, "Potential gross revenue (PGR)", True)
fill_row(25, lambda c, i: f"=SUM({c}22:{c}24)", bold=True, border=top_border)
row_label(26, "Less: general vacancy and credit loss", indent=1)
fill_row(26, lambda c, i: f"=-{c}25*{A['vac']}")
row_label(27, "Effective gross income (EGI)", True)
fill_row(27, lambda c, i: f"={c}25+{c}26", bold=True, border=top_border)

R = 29
row_label(R, "Operating Expenses", True, fill=fill_sub)
for col in ["C"] + YC:
    cf[f"{col}{R}"].fill = fill_sub
row_label(30, "Recoverable operating costs (TMI)", indent=1)
fill_row(30, lambda c, i: f"=-{A['tmi']}*(1+{A['opex_g']})^({c}$5-1)*{A['gla']}")
row_label(31, "Management fee (non-recoverable)", indent=1)
fill_row(31, lambda c, i: f"=-{c}27*{A['mgmt']}")
row_label(32, "G&A (non-recoverable)", indent=1)
fill_row(32, lambda c, i: f"=-{A['ga']}*(1+{A['opex_g']})^({c}$5-1)*{A['gla']}")
row_label(33, "Total operating expenses", True)
fill_row(33, lambda c, i: f"=SUM({c}30:{c}32)", bold=True, border=top_border)

# NOI sits on row 40 (referenced by Assumptions going-in cap); rows 34-39 kept for KPIs
row_label(35, "Occupancy (economic, after downtime and credit loss)", indent=1)
fill_row(35, lambda c, i: f"=1+({c}23+{c}26)/{c}22", PCT)
row_label(36, "NOI per SF ($)", indent=1)
fill_row(36, lambda c, i: f"={c}40/{A['gla']}", USD2)
row_label(37, "NOI growth", indent=1)
fill_row(37, lambda c, i: "" if i == 0 else f"={c}40/{YC[i - 1]}40-1", PCT)
for r in (35, 36, 37):
    for col in YC:
        cf[f"{col}{r}"].font = Font(name=FONT, italic=True, color="595959")

row_label(40, "Net operating income (NOI)", True)
fill_row(40, lambda c, i: f"={c}27+{c}33", bold=True, border=box)
for col in ["B"] + YC:
    cf[f"{col}40"].fill = fill_key if col == "B" else fill_grey

R = 42
row_label(R, "Capital Costs", True, fill=fill_sub)
for col in ["C"] + YC:
    cf[f"{col}{R}"].fill = fill_sub
infl = lambda c: f"(1+{A['lc_infl']})^({c}$5-1)"
row_label(43, "Tenant improvements", indent=1)
fill_row(43, lambda c, i: f"=-SUMPRODUCT({SF_RANGE},{flag_rng(c)})*{A['ti_w']}*{infl(c)}")
row_label(44, "Leasing commissions", indent=1)
fill_row(44, lambda c, i: f"=-SUMPRODUCT({SF_RANGE},{flag_rng(c)})*{A['lc_w']}*{infl(c)}")
row_label(45, "Capital reserve", indent=1)
fill_row(45, lambda c, i: f"=-{A['reserve']}*{A['gla']}*{infl(c)}")
row_label(46, "Cash flow before debt service", True)
fill_row(46, lambda c, i: f"={c}40+SUM({c}43:{c}45)", bold=True, border=top_border)

row_label(48, "Debt service", indent=1)
fill_row(48, lambda c, i: f"=IF({c}$5<={A['hold']},-{A['ds']},0)")
row_label(49, "Cash flow after debt service (hold period)", True)
fill_row(49, lambda c, i: f"=IF({c}$5<={A['hold']},{c}46+{c}48,0)", bold=True, border=top_border)

R = 51
row_label(R, "Sale (in the hold-period year)", True, fill=fill_sub)
for col in ["C"] + YC:
    cf[f"{col}{R}"].fill = fill_sub
noi_rng = f"${FIRST}$40:${LAST}$40"
row_label(52, "Forward NOI (year after sale)", indent=1)
fill_row(52, lambda c, i: f"=IF({c}$5={A['hold']},INDEX({noi_rng},{c}$5+1),0)")
row_label(53, "Gross sale price", indent=1)
fill_row(53, lambda c, i: f"={c}52/{A['exitcap']}")
row_label(54, "Selling costs", indent=1)
fill_row(54, lambda c, i: f"=-{c}53*{A['sellc']}")
row_label(55, "Loan repayment", indent=1)
fill_row(55, lambda c, i: f"=IF({c}$5={A['hold']},-INDEX(Debt!$D$11:$M$11,{c}$5),0)")
row_label(56, "Net sale proceeds to equity", True)
fill_row(56, lambda c, i: f"=SUM({c}53:{c}55)", bold=True, border=top_border)

R = 58
row_label(R, "Returns Cash Flows", True, fill=fill_sub)
for col in ["C"] + YC:
    cf[f"{col}{R}"].fill = fill_sub
row_label(59, "Unlevered cash flow", True)
put(cf, "C59", f"=-{A['price']}*(1+{A['cc']})", f_calc, USD, bold=True)
fill_row(59, lambda c, i: f"=IF({c}$5<={A['hold']},{c}46+{c}53+{c}54,0)", bold=True)
row_label(60, "Levered cash flow (equity)", True)
put(cf, "C60", f"=-({A['price']}*(1+{A['cc']})-{A['loan']})", f_calc, USD, bold=True)
fill_row(60, lambda c, i: f"=IF({c}$5<={A['hold']},{c}49+{c}56,0)", bold=True)
for col in YC:
    for r in (59, 60):
        cf[f"{col}{r}"].fill = fill_grey
cf[f"{LAST}5"].comment = Comment("Year 11 exists only to price the exit on forward NOI.",
                                 "Model")
cf.freeze_panes = "D7"

# =====================================================================================
# DEBT
# =====================================================================================
d = ws_d
d.column_dimensions["B"].width = 34
d.column_dimensions["C"].width = 14
DC = [L(4 + i) for i in range(10)]  # D..M = years 1..10
for c in DC:
    d.column_dimensions[c].width = 13
put(d, "B2", "Debt Schedule ($)", f_title)
put(d, "B4", "Loan amount")
put(d, "C4", f"={A['loan']}", f_link, USD)
put(d, "B5", "Interest rate")
put(d, "C5", f"={A['rate']}", f_link, PCT2)
put(d, "B6", "Monthly payment")
put(d, "C6", f"=PMT(C5/12,{A['amort']}*12,-C4)", f_calc, USD)
put(d, "B8", "Year", f_bold, fill=fill_sub)
put(d, "C8", "", fill=fill_sub)
for i, c in enumerate(DC):
    put(d, f"{c}8", i + 1, f_bold, "0", fill=fill_sub, align="right")
labels = ["Beginning balance", "Debt service", "Ending balance", "Interest paid",
          "Principal repaid", "DSCR (NOI / debt service)", "Debt yield (NOI / opening loan)"]
for k, lab in enumerate(labels):
    put(d, f"B{9 + k}", lab, f_bold if k in (2, 5) else f_calc)
for i, c in enumerate(DC):
    cfc = YC[i]
    put(d, f"{c}9", "=$C$4" if i == 0 else f"={DC[i - 1]}11", f_calc, USD)
    put(d, f"{c}10", "=$C$6*12", f_calc, USD)
    put(d, f"{c}11", f"=FV($C$5/12,12*{c}$8,$C$6,-$C$4)", f_calc, USD, bold=True)
    put(d, f"{c}12", f"={c}10-{c}13", f_calc, USD)
    put(d, f"{c}13", f"={c}9-{c}11", f_calc, USD)
    put(d, f"{c}14", f"='Cash Flow'!{cfc}40/{c}10", f_calc, MULT, bold=True)
    put(d, f"{c}15", f"='Cash Flow'!{cfc}40/{c}9", f_calc, PCT)
put(d, "B17", "Balances shown for all 10 years; the loan is repaid from sale proceeds in the "
              "hold-period year.", f_note)

# =====================================================================================
# SENSITIVITY (live: each grid cell is an IRR over its own cash-flow row on Sens_Calc)
# =====================================================================================
s, sc = ws_s, ws_sc
s.column_dimensions["B"].width = 26
for c in "CDEFGH":
    s.column_dimensions[c].width = 13
put(s, "B2", "Sensitivity Analysis", f_title)
put(s, "B3", "Every cell recalculates from the model (helper cash flows on 'Sens_Calc'). "
             "Center column/row = base case.", f_note)
put(s, "B5", "Purchase price step ($)")
put(s, "C5", 2000000, f_in, USD)
put(s, "B6", "Cap rate step")
put(s, "C6", 0.0025, f_in, PCT2)
put(s, "B7", "LTV step")
put(s, "C7", 0.05, f_in, PCT)
put(s, "B8", "Interest rate step")
put(s, "C8", 0.005, f_in, PCT2)

OFFS = [-2, -1, 0, 1, 2]
cfbds = lambda t: f"'Cash Flow'!{YC[t - 1]}$46"  # t = 1..10
noi_fwd = f"INDEX('Cash Flow'!${FIRST}$40:${LAST}$40,{A['hold']}+1)"

sc.column_dimensions["B"].width = 14
sc.column_dimensions["C"].width = 12
put(sc, "B2", "Helper cash flows for the Sensitivity grids (do not edit)", f_title)


def grid(top, title, row_lab, col_lab, row_axis, col_axis, row_fmt, col_fmt, sc_top, cf_fn):
    """row_axis/col_axis: lambda offset -> formula. cf_fn(rowref, colref, t) -> formula."""
    header_bar(s, top, 2, 8, title)
    put(s, f"B{top + 1}", f"{row_lab} (down) vs {col_lab} (across)", f_note)
    for j, o in enumerate(OFFS):
        put(s, f"{L(4 + j)}{top + 2}", col_axis(o), f_bold, col_fmt, fill_sub, align="center")
    put(s, f"C{top + 2}", "", fill=fill_sub)
    put(s, f"B{top + 2}", "", fill=fill_sub)
    # helper header
    put(sc, f"B{sc_top}", title, f_bold)
    put(sc, f"B{sc_top + 1}", row_lab, f_bold)
    put(sc, f"C{sc_top + 1}", col_lab, f_bold)
    for t in range(11):
        put(sc, f"{L(4 + t)}{sc_top + 1}", f"Yr {t}", f_bold, align="right")
    put(sc, f"O{sc_top + 1}", "IRR", f_bold, align="right")
    hr = sc_top + 2
    for i, ro in enumerate(OFFS):
        sr = top + 3 + i
        put(s, f"C{sr}", row_axis(ro), f_bold, row_fmt, fill_sub, align="center")
        for j, co in enumerate(OFFS):
            rowref, colref = f"Sensitivity!$C${sr}", f"Sensitivity!${L(4 + j)}${top + 2}"
            put(sc, f"B{hr}", f"={rowref}", f_link, row_fmt)
            put(sc, f"C{hr}", f"={colref}", f_link, col_fmt)
            for t in range(11):
                put(sc, f"{L(4 + t)}{hr}", cf_fn(f"$B{hr}", f"$C{hr}", t), f_calc, USD)
            put(sc, f"O{hr}", f"=IRR(D{hr}:N{hr})", f_calc, PCT2)
            cell = put(s, f"{L(4 + j)}{sr}", f"=Sens_Calc!$O${hr}", f_calc, PCT2, align="center",
                       border=box)
            if ro == 0 and co == 0:
                cell.fill = fill_key
                cell.font = Font(name=FONT, bold=True)
            hr += 1
    return hr


sale_net = lambda cap: f"{noi_fwd}/{cap}*(1-{A['sellc']})"
scale = lambda P: f"{P}/{A['price']}"
bal_hold = f"INDEX(Debt!$D$11:$M$11,{A['hold']})"


def lev_price_cap(P, cap, t):
    if t == 0:
        return f"=-{P}*(1+{A['cc']})+{A['ltv']}*{P}"
    return (f"=IF({t}<={A['hold']},{cfbds(t)}-{A['ds']}*{scale(P)},0)"
            f"+IF({t}={A['hold']},{sale_net(cap)}-{bal_hold}*{scale(P)},0)")


def unlev_price_cap(P, cap, t):
    if t == 0:
        return f"=-{P}*(1+{A['cc']})"
    return f"=IF({t}<={A['hold']},{cfbds(t)},0)+IF({t}={A['hold']},{sale_net(cap)},0)"


def lev_ltv_rate(ltv, r, t):
    loan = f"{ltv}*{A['price']}"
    pmt = f"PMT({r}/12,{A['amort']}*12,-{loan})"
    if t == 0:
        return f"=-{A['price']}*(1+{A['cc']})+{loan}"
    bal = f"FV({r}/12,12*{A['hold']},{pmt},-{loan})"
    return (f"=IF({t}<={A['hold']},{cfbds(t)}-{pmt}*12,0)"
            f"+IF({t}={A['hold']},{sale_net(A['exitcap'])}-{bal},0)")


price_axis = lambda o: f"={A['price']}+({o})*$C$5"
cap_axis = lambda o: f"={A['exitcap']}+({o})*$C$6"
ltv_axis = lambda o: f"={A['ltv']}+({o})*$C$7"
rate_axis = lambda o: f"={A['rate']}+({o})*$C$8"

nxt = grid(10, "Levered IRR: Purchase Price vs Exit Cap Rate", "Purchase price",
           "Exit cap", price_axis, cap_axis, USD, PCT2, 4, lev_price_cap)
nxt = grid(20, "Unlevered IRR: Purchase Price vs Exit Cap Rate", "Purchase price",
           "Exit cap", price_axis, cap_axis, USD, PCT2, nxt + 2, unlev_price_cap)
nxt = grid(30, "Levered IRR: Loan-to-Value vs Interest Rate", "Loan-to-value",
           "Interest rate", ltv_axis, rate_axis, PCT, PCT2, nxt + 2, lev_ltv_rate)
for t in range(11):
    sc.column_dimensions[L(4 + t)].width = 13
sc.column_dimensions["O"].width = 10
put(s, "B39", "Reading the grids: a lower purchase price or lower exit cap raises returns. "
              "In the financing grid, more leverage only helps while the property's yield "
              "beats the cost of debt.", f_note)

# =====================================================================================
# SUMMARY
# =====================================================================================
m = ws_sum
m.column_dimensions["B"].width = 38
m.column_dimensions["C"].width = 18
m.column_dimensions["D"].width = 4
m.column_dimensions["E"].width = 34
m.column_dimensions["F"].width = 18
put(m, "B2", "GTA Industrial Acquisition Model: Investment Summary", f_title)
put(m, "B3", "Hypothetical 200,000 SF multi-tenant industrial building in Brampton, ON. "
             "Built for learning; assumptions grounded in Q3 2026 GTA market data.", f_note)


def kv(ws, r, col, label, formula, fmt, bold=False):
    put(ws, f"{col}{r}", label, f_bold if bold else f_calc)
    vc = L(ord(col) - 64 + 1)
    put(ws, f"{vc}{r}", formula, f_calc, fmt, bold=bold, align="right")


header_bar(m, 5, 2, 3, "Property and Pricing")
kv(m, 6, "B", "Gross leasable area (SF)", f"={A['gla']}", NUM)
kv(m, 7, "B", "Purchase price", f"={A['price']}", USD)
kv(m, 8, "B", "Price per SF", f"={A['ppsf']}", USD2)
kv(m, 9, "B", "Going-in cap rate", f"={A['goingin']}", PCT2, True)
kv(m, 10, "B", "Exit cap rate", f"={A['exitcap']}", PCT2)
kv(m, 11, "B", "In-place rent vs market", f"=Assumptions!$I${rt}", PCT)
kv(m, 12, "B", "Weighted avg lease term, WALT (years)", f"=Assumptions!$G${rt}", "0.0")
kv(m, 13, "B", "Hold period (years)", f"={A['hold']}", "0")

header_bar(m, 5, 5, 6, "Sources and Uses")
kv(m, 6, "E", "Purchase price", f"={A['price']}", USD)
kv(m, 7, "E", "Closing costs", f"={A['price']}*{A['cc']}", USD)
kv(m, 8, "E", "Total uses", "=F6+F7", USD, True)
m["F8"].border = top_border
kv(m, 10, "E", "Senior loan", f"={A['loan']}", USD)
kv(m, 11, "E", "Equity", "=F8-F10", USD)
kv(m, 12, "E", "Total sources", "=F10+F11", USD, True)
m["F12"].border = top_border
kv(m, 13, "E", "Check (sources - uses)", "=F12-F8", USD)

header_bar(m, 15, 2, 3, "Returns")
kv(m, 16, "B", "Unlevered IRR", "=IRR('Cash Flow'!$C$59:$M$59)", PCT2, True)
kv(m, 17, "B", "Levered IRR", "=IRR('Cash Flow'!$C$60:$M$60)", PCT2, True)
kv(m, 18, "B", "Equity multiple",
   "=SUM('Cash Flow'!$D$60:$M$60)/-'Cash Flow'!$C$60", MULT, True)
kv(m, 19, "B", "Net profit to equity", "=SUM('Cash Flow'!$C$60:$M$60)", USD)
kv(m, 20, "B", "Avg cash-on-cash (hold period)",
   f"=SUMPRODUCT(('Cash Flow'!$D$5:$M$5<={A['hold']})*'Cash Flow'!$D$49:$M$49)"
   f"/{A['hold']}/-'Cash Flow'!$C$60", PCT2)
for r in (16, 17, 18):
    m[f"C{r}"].fill = fill_key

header_bar(m, 15, 5, 6, "Debt Metrics")
kv(m, 16, "E", "Loan-to-value", f"={A['ltv']}", PCT)
kv(m, 17, "E", "Interest rate", f"={A['rate']}", PCT2)
kv(m, 18, "E", "Year 1 DSCR", "=Debt!$D$14", MULT, True)
kv(m, 19, "E", "Minimum DSCR (hold period)",
   f"=MIN(INDEX(Debt!$D$14:$M$14,1):INDEX(Debt!$D$14:$M$14,{A['hold']}))", MULT)
kv(m, 20, "E", "Year 1 debt yield", "=Debt!$D$15", PCT2)

header_bar(m, 22, 2, 6, "Exit")
kv(m, 23, "B", "Gross sale price", "=SUM('Cash Flow'!$D$53:$M$53)", USD)
kv(m, 24, "B", "Sale price per SF", f"=C23/{A['gla']}", USD2)
kv(m, 25, "B", "Value created vs purchase price", f"=C23/{A['price']}-1", PCT)
kv(m, 23, "E", "Year 1 NOI", "='Cash Flow'!$D$40", USD)
kv(m, 24, "E", "NOI CAGR (Year 1 to exit NOI)",
   f"=(INDEX('Cash Flow'!$D$40:$N$40,{A['hold']}+1)/'Cash Flow'!$D$40)^(1/{A['hold']})-1",
   PCT2)

put(m, "B27", "Key risks to underwrite", f_bold)
risks = [
    "1. Rollover: Tenant A (40% of GLA) expires in Year 2 at a below-market rent, so the "
    "first big test is re-leasing it at market.",
    "2. Mark-to-market: Tenant B pays above today's market; its Year 5 renewal could cut rent.",
    "3. Supply: GTA availability has risen as new buildings deliver; slower rent recovery "
    "lowers returns (try 0% growth on Assumptions row 55).",
    "4. Rates: the loan rate is about equal to the going-in cap rate, so debt adds little cash "
    "yield at entry and returns rely on NOI growth. See the financing grid on 'Sensitivity'.",
]
for k, t in enumerate(risks):
    put(m, f"B{28 + k}", t, f_calc)

# chart: NOI and levered cash flow by year
ch = BarChart()
ch.type = "col"
ch.title = "NOI and Levered Cash Flow by Year ($)"
ch.y_axis.title = "$"
ch.y_axis.numFmt = "$#,##0"
ch.height, ch.width = 8, 22
data1 = Reference(ws_cf, min_col=3 + 1, max_col=3 + 10, min_row=40, max_row=40)
data2 = Reference(ws_cf, min_col=3 + 1, max_col=3 + 10, min_row=49, max_row=49)
cats = Reference(ws_cf, min_col=4, max_col=13, min_row=6, max_row=6)
from openpyxl.chart import Series
s1 = Series(data1, title="NOI")
s2 = Series(data2, title="Cash flow after debt service")
ch.series.append(s1)
ch.series.append(s2)
ch.set_categories(cats)
m.add_chart(ch, "B34")

# =====================================================================================
# MARKET DATA
# =====================================================================================
md = ws_m
widths = {"B": 34, "C": 22, "D": 46, "E": 40, "F": 60}
for k, v in widths.items():
    md.column_dimensions[k].width = v
put(md, "B2", "Market Data Behind the Assumptions", f_title)
put(md, "B3", "Figures gathered November 2026. The property and rent roll are hypothetical; "
              "market inputs are anchored to these published sources.", f_note)
hd = ["Metric", "Value", "Source", "Used for", "Link"]
for i, h in enumerate(hd):
    put(md, f"{L(2 + i)}5", h, f_head, fill=fill_head)
rows = [
    ("GTA average asking net rent, Q3 2026", "$16.06/SF (down 4.8% YoY)",
     "Avison Young via Canadian Mortgage Professional", "Market rent ($15.75 for 28 ft clear)",
     "https://www.mpamag.com/ca/specialty/commercial/gta-industrial-rents-fall-in-q3/595265"),
    ("Net rent by clear height, Q3 2026", "30 ft+: $16.84/SF; 18-21 ft: $14.86/SF",
     "Avison Young via Canadian Mortgage Professional", "Discount for a 28 ft building",
     "https://www.mpamag.com/ca/specialty/commercial/gta-industrial-rents-fall-in-q3/595265"),
    ("GTA rent trend", "7 straight quarterly declines to Q1 2026; ~9% below Q2 2024 peak",
     "WareCRE Toronto Industrial Report Q1 2026", "Low growth in 2028-29, then 3%",
     "https://warecre.com/cre-insights/market-property-insights/toronto-warehouse-market-report/"),
    ("GTA TMI (additional rent)", "$3.50 to $4.00/SF",
     "WareCRE Toronto Industrial Report Q1 2026", "Recoverable opex $3.75/SF",
     "https://warecre.com/cre-insights/market-property-insights/toronto-warehouse-market-report/"),
    ("GTA net absorption, Q3 2026", "2.0M SF; 5 straight positive quarters",
     "CBRE Toronto Industrial Figures Q3 2026", "Demand support for renewals",
     "https://www.cbre.com/insights/figures/toronto-industrial-figures-q3-2026"),
    ("Canada industrial availability, Q3 2026", "6.2%",
     "Altus Group Canadian Industrial Update Q3 2026", "Downtime and vacancy assumptions",
     "https://www.altusgroup.com/insights/canadian-industrial-market-update-q3-2026/"),
    ("Toronto industrial cap rates, Q3 2026", "Class A 5.00-5.25%; Class B 5.25-6.00%",
     "CBRE Canadian Cap Rates Q3 2026 (as summarized by RentalMath)",
     "Going-in ~5.5%; exit 5.75%", "https://rentalmath.ca/cap-rates-by-city-canada/toronto/"),
    ("Toronto industrial cap rates, Q3 2026", "5.10% to 5.50%",
     "Avison Young Q3 2026 (as summarized by Aion Appraisals)", "Cross-check on cap rates",
     "https://aionappraisals.ca/resources/ontario-cap-rates"),
    ("GTA average asking sale price, Q2 2026", "$374.53/SF",
     "CBRE Toronto Industrial Figures Q2 2026", "Price check ($270/SF for older Class B)",
     "https://mktgdocs.cbre.com/2299/972eca9b-6685-4432-ac7e-08b384b31d72-623765564/"
     "Toronto_Industrial_Figures_Q2_.pdf"),
    ("5-year Government of Canada yield", "3.59% (Oct 6, 2026)",
     "WealthNorth, citing Bank of Canada", "Loan rate 5.50% (~190 bps spread)",
     "https://wealthnorth.ca/mortgages/rates/canada-5-year-bond-yield/"),
    ("Ontario land transfer tax, commercial", "2.0% top marginal rate",
     "LendCity GTA commercial mortgage guide", "Closing costs 2.75%",
     "https://lendcity.ca/blog/commercial-mortgage-toronto-gta-guide/"),
    ("Judgment assumptions", "Renewal 65%, downtime 6 mo, TI/LC, reserves, fees",
     "Standard underwriting conventions; not from a single source",
     "Leasing and capital cost inputs", ""),
]
for k, row in enumerate(rows):
    r = 6 + k
    for i, v in enumerate(row):
        c = put(md, f"{L(2 + i)}{r}", v, f_calc)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    if row[4]:
        md[f"F{r}"].hyperlink = row[4]
        md[f"F{r}"].font = Font(name=FONT, color="0563C1", underline="single")

# validation on hold period
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="10",
                    showErrorMessage=True, error="Hold must be 1 to 10 years")
ws_a.add_data_validation(dv)
dv.add("C11")

for ws in wb.worksheets:
    ws.sheet_properties.pageSetUpPr = None
wb.calculation.fullCalcOnLoad = True
ws_sc.sheet_state = "visible"
wb.save(OUT)
print("saved", OUT)
