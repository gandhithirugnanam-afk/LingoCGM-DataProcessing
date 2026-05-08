"""
create_cgm_template.py
======================
Generates LingoCGM_CGM_Report.xlsx — a self-contained Excel workbook.

The user pastes their CGM data into the INPUT sheet (columns A & B).
All other sheets auto-recalculate using Excel formulas (requires Excel 365
or Excel 2021 for the Daily and AGP sheets which use FILTER/UNIQUE/SORT).

Sheets produced
---------------
  INPUT        — paste DateTime + Glucose data here (up to 20 000 readings)
  Summary      — scalar metrics: avg, SD, CV, GMI, eA1c, 5-tier TIR, fasting/nocturnal avg
  Daily        — per-day: count, avg glucose, TIR%, TBR%, TAR%
  AGP          — hourly 10th / 25th / median / 75th / 90th percentiles
  Time Blocks  — 3-hour window averages
  _Calc        — hidden helper (date & hour extraction per row; do not edit)

Run:  python create_cgm_template.py
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Constants ─────────────────────────────────────────────────────────────────

NROWS   = 20_000          # max data rows supported
DSTART  = 3               # first data row in both Input and _Calc (row 1=title, row 2=header)
DEND    = DSTART + NROWS - 1   # = 20 002

# ── Colour palette (matches dashboard) ───────────────────────────────────────
C = dict(
    accent  = "1D5FDB",
    green   = "059669",
    red     = "DC2626",
    amber   = "B45309",
    purple  = "7C3AED",
    bg      = "F6F8FB",
    surface = "FFFFFF",
    muted   = "7A8FAB",
    text    = "0E1624",
    text2   = "334155",
    border  = "DDE3ED",
    tbr2    = "7F1D1D",
    tbr1    = "DC2626",
    tir     = "059669",
    tar1    = "B45309",
    tar2    = "78350F",
    input_g = "ECFDF5",   # light green for glucose column
    input_d = "EFF4FF",   # light blue for datetime column
    hdr_alt = "F0F4F9",   # alternate section header
)

# ── Style helpers ─────────────────────────────────────────────────────────────

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="0E1624", size=11, italic=False, name="Calibri"):
    return Font(bold=bold, color=color, size=size, italic=italic, name=name)

def border():
    s = Side(style="thin", color="DDE3ED")
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="center", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

def col_w(ws, col_idx, width):
    ws.column_dimensions[get_column_letter(col_idx)].width = width

def hcell(ws, row, col, text, bg=None, fg="FFFFFF", bold=True, size=11,
          h_align="center", wrap=False):
    """Write a header-styled cell."""
    c = ws.cell(row=row, column=col, value=text)
    c.fill  = fill(bg or C["accent"])
    c.font  = font(bold=bold, color=fg, size=size)
    c.alignment = align(h_align, wrap=wrap)
    c.border = border()
    return c

def dcell(ws, row, col, value, bold=False, color=None, num_fmt=None,
          h_align="center", bg=None, italic=False):
    """Write a data-styled cell (value or formula string)."""
    c = ws.cell(row=row, column=col, value=value)
    c.font  = font(bold=bold, color=color or C["text"], italic=italic)
    c.alignment = align(h_align)
    c.border = border()
    if num_fmt:
        c.number_format = num_fmt
    if bg:
        c.fill = fill(bg)
    return c

def merge_hdr(ws, row, c1, c2, text, bg=None, fg="FFFFFF", size=12, h_align="left"):
    hcell(ws, row, c1, text, bg=bg, fg=fg, size=size, h_align=h_align)
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)

# ── Shorthand formula ranges ──────────────────────────────────────────────────

GLU  = f"_Calc!$B${DSTART}:$B${DEND}"
DATE = f"_Calc!$C${DSTART}:$C${DEND}"
HOUR = f"_Calc!$D${DSTART}:$D${DEND}"
N    = f"COUNT({GLU})"

def tir_formula(lo, hi):
    return f"=IFERROR(COUNTIFS({GLU},\">=\"&{lo},{GLU},\"<=\"&{hi})/{N}*100,0)"

def countif_f(crit):
    return f"=IFERROR(COUNTIF({GLU},{crit})/{N}*100,0)"

def countifs_f(c1, c2):
    return f"=IFERROR(COUNTIFS({GLU},\">=\"&{c1},{GLU},\"<\"&{c2})/{N}*100,0)"

def avgifs_hour(h_lo, h_hi):
    return (f"=IFERROR(AVERAGEIFS({GLU},{HOUR},\">=\"&{h_lo},"
            f"{HOUR},\"<\"&{h_hi}),\"—\")")

def pctile_hour(h, p):
    """Excel 365 FILTER-based percentile for a specific hour (0-23)."""
    return (f"=IFERROR(PERCENTILE(FILTER({GLU},{HOUR}={h}),{p}),\"—\")")

# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def build_input(wb):
    ws = wb.create_sheet("INPUT")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"
    ws.tab_color = C["accent"]

    # ── Row 1: title banner ──────────────────────────────────────────────────
    ws.row_dimensions[1].height = 36
    ws.merge_cells("A1:B1")
    t = ws["A1"]
    t.value = "LingoCGM — CGM Data Input"
    t.font  = font(bold=True, size=16, color="FFFFFF")
    t.fill  = fill(C["accent"])
    t.alignment = align("left")

    inst = ws["C1"]
    inst.value = (
        f"Paste your CGM data below (rows 3–{DEND:,}).  "
        "Column A = DateTime  |  Column B = Glucose (mg/dL).  "
        "All other sheets recalculate automatically.  "
        "Maximum 20 000 readings (~70 days at 5-min, ~208 days at 15-min).  "
        "Requires Excel 365 / Excel 2021."
    )
    inst.font  = font(size=9, color="FFFFFF", italic=True)
    inst.fill  = fill(C["accent"])
    inst.alignment = align("left", wrap=True)
    ws.merge_cells("C1:F1")

    # ── Row 2: column headers ────────────────────────────────────────────────
    ws.row_dimensions[2].height = 22
    hcell(ws, 2, 1, "DateTime",        bg=C["accent"],  size=12)
    hcell(ws, 2, 2, "Glucose (mg/dL)", bg="059669",     size=12)
    # Example hint
    hcell(ws, 2, 3, "← Paste data in columns A & B from row 3 onwards",
          bg=C["hdr_alt"], fg=C["muted"], bold=False, size=10, h_align="left")
    ws.merge_cells("C2:F2")

    # ── Column widths ────────────────────────────────────────────────────────
    col_w(ws, 1, 22)
    col_w(ws, 2, 18)
    col_w(ws, 3, 48)

    # ── Data rows: format A as datetime, B as number, light bg ──────────────
    for r in range(DSTART, DEND + 1):
        a = ws.cell(row=r, column=1)
        a.number_format = "YYYY-MM-DD HH:MM:SS"
        a.fill = fill("EFF4FF")
        a.alignment = align("left")
        a.border = border()

        b = ws.cell(row=r, column=2)
        b.number_format = "0.0"
        b.fill = fill("ECFDF5")
        b.alignment = align("center")
        b.border = border()

    # ── Conditional formatting: colour glucose by zone ───────────────────────
    from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
    from openpyxl.styles import PatternFill as PF

    glu_range = f"B{DSTART}:B{DEND}"
    ws.conditional_formatting.add(
        glu_range,
        CellIsRule(operator="lessThan",    formula=["70"],
                   fill=PF("solid", fgColor="FEF2F2"),
                   font=Font(color=C["red"],   bold=True, name="Calibri")))
    ws.conditional_formatting.add(
        glu_range,
        CellIsRule(operator="greaterThan", formula=["180"],
                   fill=PF("solid", fgColor="FFFBEB"),
                   font=Font(color=C["amber"], bold=True, name="Calibri")))
    ws.conditional_formatting.add(
        glu_range,
        FormulaRule(formula=[f"AND(B{DSTART}>=70,B{DSTART}<=180)"],
                    fill=PF("solid", fgColor="ECFDF5"),
                    font=Font(color=C["green"], name="Calibri")))


def build_calc(wb):
    """Hidden helper sheet: extracts valid glucose, date serial, and hour per row."""
    ws = wb.create_sheet("_Calc")
    ws.sheet_view.showGridLines = False
    ws.sheet_state = "hidden"

    # Header
    for col, hdr in enumerate(["DateTime", "Glucose_Valid", "Date_Serial", "Hour"], 1):
        ws.cell(row=1, column=col, value=hdr)

    # Formula rows (DSTART to DEND) — same row numbers as Input
    print(f"  Writing {NROWS:,} _Calc helper rows…", end="", flush=True)
    for r in range(DSTART, DEND + 1):
        ws.cell(row=r, column=1, value=f"=Input!A{r}")
        ws.cell(row=r, column=2, value=(
            f'=IF(ISNUMBER(Input!B{r}),'
            f'IF(AND(Input!B{r}>=20,Input!B{r}<=700),Input!B{r},""),"")'))
        ws.cell(row=r, column=3, value=f'=IF(B{r}="","",INT(A{r}))')
        ws.cell(row=r, column=4, value=f'=IF(B{r}="","",HOUR(A{r}))')
        if r % 5000 == 0:
            print(".", end="", flush=True)
    print(" done")

    col_w(ws, 1, 22); col_w(ws, 2, 14); col_w(ws, 3, 14); col_w(ws, 4, 8)


def build_summary(wb):
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False
    ws.tab_color = C["accent"]
    col_w(ws, 1, 36); col_w(ws, 2, 18); col_w(ws, 3, 14); col_w(ws, 4, 46)

    # Title
    ws.row_dimensions[1].height = 32
    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value = "LingoCGM — CGM Summary Report"
    t.font  = font(bold=True, size=16, color=C["accent"])
    t.alignment = align("left")

    ws.row_dimensions[2].height = 18
    ws.merge_cells("A2:D2")
    sub = ws["A2"]
    sub.value = "Edit the subject name here →"
    sub.font  = font(size=12, color=C["text2"], italic=True)
    sub.alignment = align("left")

    r = 4  # current row

    def section(title, bg=None):
        nonlocal r
        ws.row_dimensions[r].height = 22
        merge_hdr(ws, r, 1, 4, title, bg=bg or C["accent"], size=12)
        r += 1

    def metric_row(label, formula, unit="", note="", val_color=None,
                   bold_val=False, num_fmt="0.0", bg_val=None):
        nonlocal r
        ws.row_dimensions[r].height = 20
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = font(color=C["text2"]); lc.alignment = align("left"); lc.border = border()

        vc = ws.cell(row=r, column=2, value=formula)
        vc.font = font(bold=bold_val, color=val_color or C["text"], size=12)
        vc.alignment = align("center"); vc.border = border()
        vc.number_format = num_fmt
        if bg_val: vc.fill = fill(bg_val)

        uc = ws.cell(row=r, column=3, value=unit)
        uc.font = font(color=C["muted"], size=10); uc.alignment = align("center"); uc.border = border()

        nc = ws.cell(row=r, column=4, value=note)
        nc.font = font(color=C["muted"], size=10, italic=True)
        nc.alignment = align("left", wrap=True); nc.border = border()
        r += 1

    def blank():
        nonlocal r
        for col in range(1, 5):
            ws.cell(row=r, column=col).border = border()
        r += 1

    # ── OVERVIEW ─────────────────────────────────────────────────────────────
    section("OVERVIEW")
    metric_row("Subject / File",  "← enter name in cell B2",  note="Edit B2 above")
    metric_row("Total Readings",  f"={N}",                     unit="readings",
               num_fmt="0", bold_val=True)
    metric_row("Days of Data",
               f"=IFERROR(SUMPRODUCT((1/COUNTIFS({DATE},{DATE},{GLU},\"<>\"))*({GLU}<>\"\"),\"\"),\"—\")",
               unit="days",  num_fmt="0",
               note="Unique calendar days with ≥1 valid reading")
    metric_row("Date Range (start)",
               f"=IFERROR(TEXT(MIN({DATE}),\"YYYY-MM-DD\"),\"—\")",
               num_fmt="@", note="Earliest date in dataset")
    metric_row("Date Range (end)",
               f"=IFERROR(TEXT(MAX({DATE}),\"YYYY-MM-DD\"),\"—\")",
               num_fmt="@", note="Latest date in dataset")
    blank()

    # ── BASIC STATS ──────────────────────────────────────────────────────────
    section("BASIC STATISTICS")
    metric_row("Average Glucose",
               f"=IFERROR(AVERAGE({GLU}),\"—\")",
               unit="mg/dL", bold_val=True,
               note="Mean of all valid readings")
    metric_row("Minimum Glucose",
               f"=IFERROR(MIN({GLU}),\"—\")",
               unit="mg/dL", note="Lowest single reading")
    metric_row("Maximum Glucose",
               f"=IFERROR(MAX({GLU}),\"—\")",
               unit="mg/dL", note="Highest single reading")
    metric_row("Hypo Events (<70 mg/dL)",
               f"=IFERROR(COUNTIF({GLU},\"<70\"),\"—\")",
               unit="readings",  num_fmt="0",
               val_color=C["red"],
               note="Total readings below 70 mg/dL (ADA hypoglycaemia threshold)")
    blank()

    # ── GLYCEMIC VARIABILITY ─────────────────────────────────────────────────
    section("GLYCEMIC VARIABILITY")
    metric_row("Standard Deviation (SD)",
               f"=IFERROR(STDEV.P({GLU}),\"—\")",
               unit="mg/dL",
               note="Population SD — how spread glucose values are")
    metric_row("Coefficient of Variation (CV)",
               f"=IFERROR(STDEV.P({GLU})/AVERAGE({GLU})*100,\"—\")",
               unit="%",
               note="CV = SD ÷ Mean × 100.  Target: <36% (stable).  36–50%: moderate.  >50%: high variability")
    metric_row("MAGE",
               "— (see note)",
               num_fmt="@",
               note="Mean Amplitude of Glycemic Excursions cannot be calculated with Excel formulas. "
                    "Use the Python script (cgm_to_xls.py) to obtain MAGE.")
    blank()

    # ── ESTIMATED HBA1C ───────────────────────────────────────────────────────
    section("ESTIMATED HbA1c")
    metric_row("GMI — Glucose Management Indicator",
               f"=IFERROR(3.31+0.02392*AVERAGE({GLU}),\"—\")",
               unit="%",
               val_color=C["accent"], bold_val=True,
               note="ADA/EASD formula: 3.31 + 0.02392 × mean glucose (mg/dL).  ≥7.0%: Diabetic  5.7–6.9%: Pre-diabetic  <5.7%: Normal")
    metric_row("eA1c (legacy)",
               f"=IFERROR((AVERAGE({GLU})+46.7)/28.7,\"—\")",
               unit="%",
               note="Legacy estimate: (avg + 46.7) / 28.7.  GMI is preferred for CGM data.")
    blank()

    # ── 5-TIER TIR ────────────────────────────────────────────────────────────
    section("5-TIER TIME IN RANGE  (ADA/EASD Consensus Targets — T2D)")

    tier_rows = [
        ("TBR Level 2  (<54 mg/dL)",
         f"=IFERROR(COUNTIF({GLU},\"<54\")/{N}*100,0)",
         C["tbr2"], "Target <1%"),
        ("TBR Level 1  (54–69 mg/dL)",
         f"=IFERROR(COUNTIFS({GLU},\">=\"&54,{GLU},\"<\"&70)/{N}*100,0)",
         C["tbr1"], "Target <4%"),
        ("TIR  (70–180 mg/dL)",
         f"=IFERROR(COUNTIFS({GLU},\">=\"&70,{GLU},\"<=\"&180)/{N}*100,0)",
         C["tir"],  "Target >70%"),
        ("TAR Level 1  (181–250 mg/dL)",
         f"=IFERROR(COUNTIFS({GLU},\">\"&180,{GLU},\"<=\"&250)/{N}*100,0)",
         C["tar1"], "Target <25%"),
        ("TAR Level 2  (>250 mg/dL)",
         f"=IFERROR(COUNTIF({GLU},\">250\")/{N}*100,0)",
         C["tar2"], "Target <5%"),
    ]
    for label, formula, clr, target in tier_rows:
        ws.row_dimensions[r].height = 20
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = font(color=C["text2"]); lc.alignment = align("left"); lc.border = border()
        lc.fill = fill(C["hdr_alt"])
        vc = ws.cell(row=r, column=2, value=formula)
        vc.font = font(bold=True, color=clr, size=12)
        vc.number_format = "0.0"; vc.alignment = align("center"); vc.border = border()
        vc.fill = fill(C["hdr_alt"])
        uc = ws.cell(row=r, column=3, value="%")
        uc.font = font(color=C["muted"], size=10); uc.alignment = align("center"); uc.border = border()
        nc = ws.cell(row=r, column=4, value=target)
        nc.font = font(color=C["muted"], size=10, italic=True)
        nc.alignment = align("left"); nc.border = border()
        r += 1
    blank()

    # ── TIME-OF-DAY AVERAGES ──────────────────────────────────────────────────
    section("TIME-OF-DAY AVERAGES")
    metric_row("Fasting Average (6–10 AM)",
               avgifs_hour(6, 10),
               unit="mg/dL", num_fmt='0.0"  mg/dL"',
               note="Readings between 06:00–09:59. Key metformin efficacy marker.")
    metric_row("Nocturnal Average (0–4 AM)",
               avgifs_hour(0, 4),
               unit="mg/dL",
               note="Readings between 00:00–03:59. Elevated values suggest hepatic insulin resistance.")
    blank()

    # ── TOTALS SUMMARY BAR ────────────────────────────────────────────────────
    section("SIMPLIFIED TIR  (3-tier)", bg=C["text2"])
    metric_row("TBR Total  (<70 mg/dL)",
               f"=IFERROR(COUNTIF({GLU},\"<70\")/{N}*100,0)",
               unit="%", val_color=C["red"],  bold_val=True,
               note="All readings below target range. ADA target <4%.")
    metric_row("TIR  (70–180 mg/dL)",
               f"=IFERROR(COUNTIFS({GLU},\">=\"&70,{GLU},\"<=\"&180)/{N}*100,0)",
               unit="%", val_color=C["green"], bold_val=True,
               note="Readings within target range. ADA target >70%.")
    metric_row("TAR Total  (>180 mg/dL)",
               f"=IFERROR(COUNTIF({GLU},\">180\")/{N}*100,0)",
               unit="%", val_color=C["amber"], bold_val=True,
               note="All readings above target range. ADA target <25%.")


def build_daily(wb):
    ws = wb.create_sheet("Daily")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"
    ws.tab_color = C["green"]

    col_w(ws, 1, 16); col_w(ws, 2, 12); col_w(ws, 3, 20)
    col_w(ws, 4, 16); col_w(ws, 5, 16); col_w(ws, 6, 16); col_w(ws, 7, 14)

    # Title
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = "Daily Glucose Breakdown"
    t.font  = font(bold=True, size=14, color=C["accent"])
    t.alignment = align("left")

    ws.row_dimensions[2].height = 16
    ws.merge_cells("A2:G2")
    s = ws["A2"]
    s.value = ("Dates are auto-populated from your data (Excel 365 / 2021 required for UNIQUE/FILTER functions). "
               "Per-day metrics recalculate automatically.")
    s.font = font(size=9, color=C["muted"], italic=True)
    s.alignment = align("left", wrap=True)

    # Column headers
    ws.row_dimensions[3].height = 22
    for col, (hdr, w) in enumerate([
        ("Date",              16), ("Readings",    12),
        ("Avg Glucose",       20), ("TIR % (70–180)", 16),
        ("TBR % (<70)",       16), ("TAR % (>180)", 16),
        ("Status",            14)
    ], 1):
        hcell(ws, 3, col, hdr)

    # Row 4: spilled date list (Excel 365)
    date_formula = (
        f"=IFERROR(SORT(UNIQUE(FILTER({DATE},{DATE}<>\"\"))),"
        "\"No data — paste CGM data in the INPUT sheet\")"
    )
    dc = ws.cell(row=4, column=1, value=date_formula)
    dc.number_format = "YYYY-MM-DD"
    dc.font  = font(color=C["text2"])
    dc.alignment = align("center")
    dc.border = border()

    # Rows 4–369: formulas for B–G (covers up to 366 unique days)
    for r in range(4, 370):
        date_ref = f"A{r}"

        # Col B: reading count
        dcell(ws, r, 2,
              f'=IFERROR(IF({date_ref}="","",COUNTIFS({DATE},{date_ref},{GLU},"<>")),"")',
              num_fmt="0")

        # Col C: avg glucose
        avg_c = dcell(ws, r, 3,
               f'=IFERROR(IF({date_ref}="","",AVERAGEIFS({GLU},{DATE},{date_ref})),"")',
               num_fmt="0.0")

        # Col D: TIR
        dcell(ws, r, 4,
              f'=IFERROR(IF({date_ref}="","",COUNTIFS({GLU},">="&70,{GLU},"<="&180,{DATE},{date_ref})'
              f'/COUNTIFS({DATE},{date_ref},{GLU},"<>")*100),"")' ,
              num_fmt='0.0"%"')

        # Col E: TBR
        dcell(ws, r, 5,
              f'=IFERROR(IF({date_ref}="","",COUNTIFS({GLU},"<70",{DATE},{date_ref})'
              f'/COUNTIFS({DATE},{date_ref},{GLU},"<>")*100),"")',
              num_fmt='0.0"%"')

        # Col F: TAR
        dcell(ws, r, 6,
              f'=IFERROR(IF({date_ref}="","",COUNTIFS({GLU},">180",{DATE},{date_ref})'
              f'/COUNTIFS({DATE},{date_ref},{GLU},"<>")*100),"")',
              num_fmt='0.0"%"')

        # Col G: status text
        dcell(ws, r, 7,
              f'=IFERROR(IF({date_ref}="","",IF(D{r}>=70,"✓ Good",IF(E{r}>4,"Low glucose","High glucose"))),"")' ,
              num_fmt="@")

    # Conditional formatting on daily TIR column (D4:D369)
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill as PF
    ws.conditional_formatting.add("D4:D369",
        CellIsRule(operator="greaterThanOrEqual", formula=["70"],
                   fill=PF("solid", fgColor="ECFDF5"),
                   font=Font(color=C["green"], bold=True, name="Calibri")))
    ws.conditional_formatting.add("D4:D369",
        CellIsRule(operator="lessThan", formula=["50"],
                   fill=PF("solid", fgColor="FEF2F2"),
                   font=Font(color=C["red"], bold=True, name="Calibri")))


def build_agp(wb):
    ws = wb.create_sheet("AGP")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"
    ws.tab_color = C["purple"]

    col_w(ws, 1, 12); col_w(ws, 2, 10)
    for c in range(3, 9): col_w(ws, c, 18)

    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = "Ambulatory Glucose Profile (AGP) — Hourly Percentiles"
    t.font  = font(bold=True, size=14, color=C["accent"])
    t.alignment = align("left")

    ws.row_dimensions[2].height = 16
    ws.merge_cells("A2:H2")
    s = ws["A2"]
    s.value = ("Median glucose and interquartile range (25th–75th pct) by hour of day. "
               "Target range 70–180 mg/dL shown for reference. Requires Excel 365.")
    s.font = font(size=9, color=C["muted"], italic=True)
    s.alignment = align("left")

    ws.row_dimensions[3].height = 22
    hdrs = ["Hour", "Readings", "10th pct", "25th pct (IQR low)",
            "Median (50th)", "75th pct (IQR high)", "90th pct", "Zone"]
    for col, hdr in enumerate(hdrs, 1):
        hcell(ws, 3, col, hdr)

    for i, h in enumerate(range(24)):
        r = i + 4
        hour_label = f"{h:02d}:00"

        dcell(ws, r, 1, hour_label, bold=True, color=C["accent"])
        dcell(ws, r, 2,
              f"=IFERROR(COUNTIF({HOUR},{h}),0)", num_fmt="0")
        dcell(ws, r, 3, pctile_hour(h, 0.10), num_fmt="0.0")
        dcell(ws, r, 4, pctile_hour(h, 0.25), num_fmt="0.0")
        dcell(ws, r, 5, pctile_hour(h, 0.50), num_fmt="0.0",
              bold=True, color=C["accent"])
        dcell(ws, r, 6, pctile_hour(h, 0.75), num_fmt="0.0")
        dcell(ws, r, 7, pctile_hour(h, 0.90), num_fmt="0.0")

        # Zone based on median
        zone_formula = (
            f'=IFERROR(IF({pctile_hour(h,0.50)[1:]}="","—",'
            f'IF({pctile_hour(h,0.50)[1:]}<70,"Low",'
            f'IF({pctile_hour(h,0.50)[1:]}>180,"High","In Range"))),"—")'
        )
        dcell(ws, r, 8, zone_formula, num_fmt="@")

    # Reference lines note
    r2 = 28
    ws.merge_cells(f"A{r2}:H{r2}")
    note = ws.cell(row=r2, column=1,
                   value="Reference: Target range 70–180 mg/dL  |  IQR band = 25th–75th percentile  |  Solid line = median")
    note.font = font(size=9, color=C["muted"], italic=True)
    note.alignment = align("left")


def build_timeblocks(wb):
    ws = wb.create_sheet("Time Blocks")
    ws.sheet_view.showGridLines = False
    ws.tab_color = C["amber"]

    col_w(ws, 1, 16); col_w(ws, 2, 12); col_w(ws, 3, 22)
    col_w(ws, 4, 14); col_w(ws, 5, 44)

    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value = "Average Glucose by 3-Hour Time Block"
    t.font  = font(bold=True, size=14, color=C["accent"])
    t.alignment = align("left")

    ws.row_dimensions[2].height = 16
    ws.merge_cells("A2:E2")
    s = ws["A2"]
    s.value = "Average glucose per 3-hour window — identifies postprandial hyperglycaemia patterns and dawn phenomenon."
    s.font = font(size=9, color=C["muted"], italic=True)
    s.alignment = align("left")

    ws.row_dimensions[3].height = 22
    for col, hdr in enumerate(
            ["Time Window", "Readings", "Avg Glucose (mg/dL)", "Zone", "Clinical Relevance"], 1):
        hcell(ws, 3, col, hdr)

    blocks = [
        (0,  3,  "12–3 AM",  "Overnight baseline — hepatic glucose output"),
        (3,  6,  "3–6 AM",   "Dawn phenomenon window — early-morning glucose rise"),
        (6,  9,  "6–9 AM",   "Fasting / pre-breakfast — key metformin efficacy marker"),
        (9,  12, "9–12 PM",  "Post-breakfast — postprandial excursion"),
        (12, 15, "12–3 PM",  "Lunch / midday postprandial"),
        (15, 18, "3–6 PM",   "Afternoon — late postprandial / activity"),
        (18, 21, "6–9 PM",   "Dinner / evening postprandial"),
        (21, 24, "9–12 AM",  "Evening / pre-sleep"),
    ]

    for i, (h_lo, h_hi, label, relevance) in enumerate(blocks, 4):
        dcell(ws, i, 1, label, bold=True, h_align="left")
        dcell(ws, i, 2,
              f"=IFERROR(COUNTIFS({HOUR},\">=\"&{h_lo},{HOUR},\"<\"&{h_hi}),0)",
              num_fmt="0")
        avg_formula = avgifs_hour(h_lo, h_hi)
        dcell(ws, i, 3, avg_formula, num_fmt="0.0", bold=True)
        zone_formula = (
            f'=IFERROR(IF({avg_formula[1:]}="","—",'
            f'IF({avg_formula[1:]}<70,"Low",'
            f'IF({avg_formula[1:]}>180,"High","In Range"))),"—")'
        )
        dcell(ws, i, 4, zone_formula, num_fmt="@")
        rc = ws.cell(row=i, column=5, value=relevance)
        rc.font = font(color=C["muted"], size=10, italic=True)
        rc.alignment = align("left")
        rc.border = border()

    # Conditional formatting on avg glucose column (C4:C11)
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill as PF
    ws.conditional_formatting.add("C4:C11",
        CellIsRule(operator="greaterThan", formula=["180"],
                   fill=PF("solid", fgColor="FFFBEB"),
                   font=Font(color=C["amber"], bold=True, name="Calibri")))
    ws.conditional_formatting.add("C4:C11",
        CellIsRule(operator="lessThan",    formula=["70"],
                   fill=PF("solid", fgColor="FEF2F2"),
                   font=Font(color=C["red"], bold=True, name="Calibri")))
    ws.conditional_formatting.add("C4:C11",
        CellIsRule(operator="between",     formula=["70", "180"],
                   fill=PF("solid", fgColor="ECFDF5"),
                   font=Font(color=C["green"], bold=True, name="Calibri")))


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    out = "LingoCGM_CGM_Report.xlsx"
    print(f"Building {out}  (NROWS={NROWS:,})…")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default Sheet1

    print("  INPUT sheet…")
    build_input(wb)
    print("  _Calc sheet…")
    build_calc(wb)
    print("  Summary sheet…")
    build_summary(wb)
    print("  Daily sheet…")
    build_daily(wb)
    print("  AGP sheet…")
    build_agp(wb)
    print("  Time Blocks sheet…")
    build_timeblocks(wb)

    # Sheet order: INPUT first, then analytics, _Calc hidden at end
    sheet_order = ["INPUT", "Summary", "Daily", "AGP", "Time Blocks", "_Calc"]
    wb._sheets.sort(key=lambda s: sheet_order.index(s.title)
                    if s.title in sheet_order else 99)

    print(f"  Saving…", end="", flush=True)
    wb.save(out)
    print(f" done → {out}")

    import os
    size_kb = os.path.getsize(out) / 1024
    print(f"  File size: {size_kb:.0f} KB")
    print(f"\n  Sheets: {[s.title for s in wb.worksheets]}")
    print(f"  Data capacity: {NROWS:,} rows  "
          f"(≈ {NROWS//288} days at 5-min  /  ≈ {NROWS//96} days at 15-min)")
    print("\n  ✓  Open LingoCGM_CGM_Report.xlsx in Excel 365, paste CGM data")
    print("     in the INPUT sheet (col A = DateTime, col B = Glucose mg/dL).")
    print("     All other sheets recalculate automatically.")


if __name__ == "__main__":
    main()
