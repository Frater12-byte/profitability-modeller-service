from flask import Flask, request, jsonify
import base64, io, os, datetime, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ── palette ───────────────────────────────────────────────────────────────────
DARK_BLUE  = "1F3864"; MID_BLUE   = "2E5FA3"; LIGHT_BLUE = "BDD7EE"
PALE_BLUE  = "DEEAF1"; GOLD       = "FFF2CC"; GRN_HDR    = "375623"
LIGHT_GRN  = "E2EFDA"; WHITE      = "FFFFFF"; LIGHT_GREY = "F2F2F2"
DARK_GREY  = "595959"; AMBER      = "C55A11"; NAVY       = "1F3864"

def fill(h):  return PatternFill("solid", fgColor=h)
def hf(sz=10, bold=True, color=WHITE): return Font(name="Arial", size=sz, bold=bold, color=color)
def bf(sz=9,  bold=False, color="000000"): return Font(name="Arial", size=sz, bold=bold, color=color)
def bdr():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)
def bdr_thick():
    t = Side(style="medium", color="1F3864")
    n = Side(style="thin",   color="BFBFBF")
    return Border(left=t, right=t, top=n, bottom=n)

CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left",   vertical="center")
RGHT = Alignment(horizontal="right",  vertical="center")

AED  = '#,##0;(#,##0);"-"'
AED2 = '#,##0.00;(#,##0.00);"-"'
PCT  = '0.0%;(0.0%);"-"'
INT  = '#,##0;(#,##0);"-"'
GP_DELTA = '[Green]+#,##0;[Red]-#,##0;"-"'

# ── safe numeric extractor ────────────────────────────────────────────────────
def safe(v):
    if v is None or v in ("#VALUE!", "#REF!", "#N/A", "#DIV/0!"):
        return None
    if isinstance(v, datetime.datetime):
        return None
    if isinstance(v, str):
        try:
            return float(v.replace(",", ""))
        except (ValueError, AttributeError):
            return None
    return float(v) if isinstance(v, (int, float)) else None


# ── data loaders ──────────────────────────────────────────────────────────────
def load_d1(wb):
    """data_1: Agency Group / Customer breakdown.
       Cols (1-based): 1=AgencyGroup, 2=Customer, 3=Sales(TV), 10=GP, 12=GP%, 15=RNTs, 18=Bookings
    """
    ws = wb.active
    agencies = {}   # {name: {tv, gp, gp_pct}}
    customers = []  # [{agency, customer, tv, gp, gp_pct}]
    cur_agency = None

    for r in range(2, ws.max_row + 1):
        ag = ws.cell(r, 1).value
        cu = ws.cell(r, 2).value
        if ag:
            cur_agency = str(ag).strip()
        if not cu or not cur_agency:
            continue
        cu = str(cu).strip()
        tv  = safe(ws.cell(r, 3).value)
        gp  = safe(ws.cell(r, 10).value)
        gp_pct = safe(ws.cell(r, 12).value)

        if cu == "Total":
            # Agency-level summary row
            agencies[cur_agency] = dict(agency=cur_agency, tv=tv or 0, gp=gp or 0, gp_pct=gp_pct or 0)
        else:
            if gp is not None or tv is not None:  # skip zero-data rows
                customers.append(dict(
                    agency=cur_agency, customer=cu,
                    tv=tv or 0, gp=gp or 0, gp_pct=gp_pct or 0
                ))

    # Sort agencies by TV desc
    ag_list = sorted(agencies.values(), key=lambda x: -(x["tv"] or 0))
    # Sort customers by GP desc within agency
    cu_list = sorted(customers, key=lambda x: (x["agency"], -(x["gp"] or 0)))
    return ag_list, cu_list


def load_d2(wb):
    """data_2: Country/Destination breakdown.
       Cols (1-based): 1=Country, 2=Sales(TV), 9=GP, 11=GP%
    """
    ws = wb.active
    rows = []
    for r in range(2, ws.max_row + 1):
        co = ws.cell(r, 1).value
        if not co:
            continue
        co  = str(co).strip()
        tv  = safe(ws.cell(r, 2).value)
        gp  = safe(ws.cell(r, 9).value)
        gp_pct = safe(ws.cell(r, 11).value)
        if gp is not None or tv is not None:
            rows.append(dict(country=co, tv=tv or 0, gp=gp or 0, gp_pct=gp_pct or 0))

    return sorted(rows, key=lambda x: -(x["tv"] or 0))


# ── determine current YTD month from data ─────────────────────────────────────
def detect_ytd_months(ag_rows, seasonality_weights):
    """Return the month index (1-12) of the last month with data."""
    # Use today as a simple proxy; could be improved with actual data column scan
    today = datetime.date.today()
    return today.month


# ── sheet builders ─────────────────────────────────────────────────────────────
MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
WEIGHTS = [0.09, 0.08, 0.10, 0.10, 0.09, 0.08, 0.07, 0.07, 0.08, 0.09, 0.09, 0.06]

def build_seasonality(ws, today_str, ag_rows):
    """Rebuild the Seasonality sheet matching the real modeller exactly."""
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    # Row 1 – title banner
    ws.merge_cells("A1:H1")
    c = ws.cell(1, 1, f"SEASONALITY & EOY FORECAST  —  Booking Date 2026 YTD")
    c.font = hf(12); c.fill = fill(DARK_BLUE); c.alignment = CTR
    ws.row_dimensions[1].height = 26

    # Row 2 – section labels
    for col, label, span, clr in [
        (1,"Month",1,MID_BLUE),(2,"2026 Weights",2,MID_BLUE),
        (4,"TV & GP  ✎ editable",2,"7F6000"),(6,"EOY Totals",2,GRN_HDR),(8,"Note",1,DARK_GREY)
    ]:
        if span > 1:
            ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col+span-1)
        c = ws.cell(2, col, label)
        c.font = hf(9); c.fill = fill(clr); c.alignment = CTR
    ws.row_dimensions[2].height = 18

    # Row 3 – column headers
    hdrs = ["Month","Monthly Wt","Cumul.","TV (AED)","GP (AED)","EOY TV","EOY GP","Note"]
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(3, ci, h)
        c.font = hf(9); c.fill = fill(MID_BLUE); c.alignment = CTR; c.border = bdr()
    ws.row_dimensions[3].height = 28

    # Compute YTD totals from agency data
    ytd_tv = sum(r["tv"] for r in ag_rows)
    ytd_gp = sum(r["gp"] for r in ag_rows)
    today  = datetime.date.today()
    ytd_m  = today.month  # number of complete months

    # Rows 4-15 – months
    ytd_wt = sum(WEIGHTS[:ytd_m])

    for i, (mo, wt) in enumerate(zip(MONTHS, WEIGHTS)):
        r   = i + 4
        bg  = LIGHT_GREY if i % 2 == 0 else WHITE
        completed = (i < ytd_m)

        # A – Month
        c = ws.cell(r, 1, mo); c.font = bf(bold=True); c.fill = fill(bg)
        c.alignment = LEFT; c.border = bdr()

        # B – Weight (editable, gold)
        c = ws.cell(r, 2, wt); c.font = Font(name="Arial", size=9, color="0000FF")
        c.fill = fill(GOLD); c.alignment = CTR; c.number_format = PCT; c.border = bdr()

        # C – Cumulative
        c = ws.cell(r, 3, f"=SUM($B$4:B{r})")
        c.font = bf(color="006100"); c.alignment = CTR; c.number_format = PCT; c.border = bdr()

        if completed:
            # D/E – actual TV and GP (editable, gold)
            mo_tv = ytd_tv * (wt / ytd_wt) if ytd_wt else 0
            mo_gp = ytd_gp * (wt / ytd_wt) if ytd_wt else 0
            c = ws.cell(r, 4, round(mo_tv, 2))
            c.font = Font(name="Arial", size=9, color="0000FF")
            c.fill = fill(GOLD); c.alignment = RGHT; c.number_format = AED; c.border = bdr()
            c = ws.cell(r, 5, round(mo_gp, 2))
            c.font = Font(name="Arial", size=9, color="0000FF")
            c.fill = fill(GOLD); c.alignment = RGHT; c.number_format = AED; c.border = bdr()
            # F/G – EOY for actual months: actual ÷ ytd_wt × mo_wt
            c = ws.cell(r, 6, f"=IFERROR(D{r}/C{r}*B{r},0)")
            c.font = bf(bold=True, color="006100"); c.fill = fill(LIGHT_GRN)
            c.alignment = RGHT; c.number_format = AED; c.border = bdr()
            c = ws.cell(r, 7, f"=IFERROR(E{r}/C{r}*B{r},0)")
            c.font = bf(bold=True, color="006100"); c.fill = fill(LIGHT_GRN)
            c.alignment = RGHT; c.number_format = AED; c.border = bdr()
            note = f"●  Actual ÷ YTD wt × mo. wt"
        else:
            # D/E – empty for future months
            for col in [4, 5]:
                c = ws.cell(r, col, ""); c.fill = fill(bg); c.border = bdr()
            # F/G – forecast: YTD TV/GP ÷ ytd_cumulative × month_wt
            c = ws.cell(r, 6, f"=IFERROR(SUM(D$4:D${3+ytd_m})/C{3+ytd_m}*B{r},0)")
            c.font = bf(color="006100"); c.fill = fill(LIGHT_GRN)
            c.alignment = RGHT; c.number_format = AED; c.border = bdr()
            c = ws.cell(r, 7, f"=IFERROR(SUM(E$4:E${3+ytd_m})/C{3+ytd_m}*B{r},0)")
            c.font = bf(color="006100"); c.fill = fill(LIGHT_GRN)
            c.alignment = RGHT; c.number_format = AED; c.border = bdr()
            note = f"○  Forecast: YTD ÷ {int(ytd_wt*100)}% × {int(wt*100)}%"

        c = ws.cell(r, 8, note); c.font = bf(sz=8, color=DARK_GREY)
        c.alignment = LEFT; c.border = bdr()

    # Row 16 – TOTAL
    tr = 16
    ws.merge_cells(f"A{tr}:A{tr}")
    c = ws.cell(tr, 1, "TOTAL"); c.font = hf(9); c.fill = fill(DARK_BLUE)
    c.alignment = LEFT; c.border = bdr()
    for ci, fmt in [(2, PCT), (3, PCT)]:
        c = ws.cell(tr, ci, f"=SUM({get_column_letter(ci)}4:{get_column_letter(ci)}15)")
        c.font = hf(9); c.fill = fill(DARK_BLUE); c.alignment = CTR
        c.number_format = fmt; c.border = bdr()
    for ci, fmt in [(4, AED), (5, AED), (6, AED), (7, AED)]:
        c = ws.cell(tr, ci, f"=SUM({get_column_letter(ci)}4:{get_column_letter(ci)}15)")
        c.font = hf(9); c.fill = fill(DARK_BLUE); c.alignment = RGHT
        c.number_format = fmt; c.border = bdr()
    ws.cell(tr, 8, "").border = bdr()

    # Row 18 – summary note
    ytd_wt_pct = int(ytd_wt * 100)
    eoy_tv = ytd_tv / ytd_wt if ytd_wt else 0
    eoy_gp = ytd_gp / ytd_wt if ytd_wt else 0
    ws.merge_cells("A18:H18")
    note18 = (f"YTD ({'-'.join(MONTHS[:ytd_m])}): {ytd_tv:,.0f} AED TV  |  "
              f"YTD weight: {ytd_wt_pct}%  ->  EOY projection: {eoy_tv:,.0f} AED TV  /  {eoy_gp:,.0f} AED GP")
    c = ws.cell(18, 1, note18)
    c.font = bf(sz=8, color=DARK_GREY); c.alignment = LEFT

    # Column widths
    for ci, w in enumerate([10, 13, 10, 16, 16, 16, 16, 38], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def _kpi_banner(ws, ytd_tv, ytd_gp, ytd_wt):
    """Top KPI strip: YTD TV | YTD GP | YTD GP% | EOY TV | EOY GP | EOY GP%"""
    labels = ["YTD Total Value","","","YTD Gross Profit","","","YTD GP%","","",
              "EOY TV (Base)","","","EOY GP (Base)","","","EOY GP (Adjusted)"]
    clrs   = [MID_BLUE]*3 + [MID_BLUE]*3 + [MID_BLUE]*3 + [GRN_HDR]*3 + [GRN_HDR]*3 + ["7F3F00"]*3

    for ci, (lbl, clr) in enumerate(zip(labels, clrs), 1):
        c = ws.cell(1, ci, lbl)
        c.font = hf(9, bold=True); c.fill = fill(clr); c.alignment = CTR; c.border = bdr()
    ws.row_dimensions[1].height = 20

    # Row 2 – KPI values
    eoy_tv = ytd_tv / ytd_wt if ytd_wt else 0
    eoy_gp = ytd_gp / ytd_wt if ytd_wt else 0
    gp_pct = ytd_gp / ytd_tv  if ytd_tv else 0

    values = [ytd_tv,"","",ytd_gp,"","",gp_pct,"","",eoy_tv,"","",eoy_gp,"","",eoy_gp]
    fmts   = [AED,"","",AED,"","",PCT,"","",AED,"","",AED,"","",AED]
    for ci, (val, fmt) in enumerate(zip(values, fmts), 1):
        if val == "":
            continue
        c = ws.cell(2, ci, val)
        c.font = hf(10, bold=True); c.fill = fill(WHITE)
        c.alignment = RGHT; c.number_format = fmt; c.border = bdr()
    ws.row_dimensions[2].height = 22


def build_analysis_sheet(ws, title, rows, id_key, id_label,
                          agency_key=None, ytd_wt=0.27, today_str=""):
    """
    Builds Agency Groups / Customer / Destination analysis tab.
    Matches the real modeller layout exactly:
      - Row 1: KPI banner (TV | GP | GP% | EOY TV | EOY GP | Adj GP)
      - Row 2: blank KPI values
      - Row 3: blank separator
      - Row 4: blank
      - Row 5: section title
      - Row 6: sub-section colour bands
      - Row 7: column headers
      - Row 8+: data
      - After data: TOTAL row, notes, opportunity insights
    """
    ws.sheet_view.showGridLines = False

    n_id   = 2 if agency_key else 1      # number of identity columns
    N_COLS = n_id + 10                   # total columns (A…K or A…L)
    gl     = get_column_letter

    ytd_tv  = sum(r["tv"]  for r in rows) if not agency_key else sum(r["tv"] for r in rows)
    ytd_gp  = sum(r["gp"]  for r in rows) if not agency_key else sum(r["gp"] for r in rows)
    eoy_tv  = ytd_tv / ytd_wt  if ytd_wt else 0
    eoy_gp  = ytd_gp / ytd_wt  if ytd_wt else 0
    avg_gp_pct = ytd_gp / ytd_tv if ytd_tv else 0

    # ── KPI banner (rows 1-2) ──────────────────────────────────────────────
    kpi_specs = [
        ("YTD Total Value",  ytd_tv,  AED, MID_BLUE, 3),
        ("YTD Gross Profit", ytd_gp,  AED, MID_BLUE, 3),
        ("YTD GP%",          avg_gp_pct, PCT, MID_BLUE, 3),
        ("EOY TV (Base)",    eoy_tv,  AED, GRN_HDR, 3),
        ("EOY GP (Base)",    eoy_gp,  AED, GRN_HDR, 3),
        ("EOY GP (Adjusted)","", AED, "7F3F00", 3),
    ]
    col = 1
    for label, val, fmt, clr, span in kpi_specs:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col+span-1)
        c = ws.cell(1, col, label)
        c.font = hf(9); c.fill = fill(clr); c.alignment = CTR; c.border = bdr()
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col+span-1)
        if val != "":
            c2 = ws.cell(2, col, val)
        else:
            c2 = ws.cell(2, col, "")
        c2.font = hf(11, bold=True); c2.fill = fill(WHITE)
        c2.alignment = RGHT; c2.number_format = fmt; c2.border = bdr()
        col += span
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 24

    # ── Rows 3-4: blank spacer ─────────────────────────────────────────────
    for r in [3, 4]:
        ws.row_dimensions[r].height = 6

    # ── Row 5: section title ───────────────────────────────────────────────
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=N_COLS)
    c = ws.cell(5, 1, title)
    c.font = hf(13); c.fill = fill(DARK_BLUE); c.alignment = CTR
    ws.row_dimensions[5].height = 28

    # ── Row 6: sub-section colour bands ───────────────────────────────────
    if agency_key:  # Customer has 2 id cols
        bands = [(n_id,"",DARK_BLUE),(2,"◀  YTD ACTUALS",MID_BLUE),
                 (2,"✏  SCENARIO INPUTS","7F3F00"),(3,"▶  EOY BASE FCST",GRN_HDR),
                 (2,"★  ADJUSTED FCST","375E23")]
    else:
        bands = [(n_id,"",DARK_BLUE),(2,"◀  YTD ACTUALS",MID_BLUE),
                 (2,"✏  SCENARIO INPUTS","7F3F00"),(3,"▶  EOY BASE FCST",GRN_HDR),
                 (2,"★  ADJUSTED FCST","375E23")]
    col = 1
    for span, label, clr in bands:
        if span > 1:
            ws.merge_cells(start_row=6, start_column=col, end_row=6, end_column=col+span-1)
        c = ws.cell(6, col, label)
        c.font = hf(9); c.fill = fill(clr); c.alignment = CTR; c.border = bdr()
        col += span
    ws.row_dimensions[6].height = 18

    # ── Row 7: column headers ──────────────────────────────────────────────
    if agency_key:
        hdrs = [" Agency", id_label, "YTD TV (AED)", "YTD GP (AED)", "YTD GP%",
                "GP% Adj", "TV Chg %", "EOY TV", "EOY GP", "EOY GP%",
                "Adj. EOY GP", "GP Delta"]
    else:
        hdrs = [id_label, "YTD TV (AED)", "YTD GP (AED)", "YTD GP%",
                "GP% Adj", "TV Chg %", "EOY TV", "EOY GP", "EOY GP%",
                "Adj. EOY GP", "GP Delta"]
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(7, ci, h)
        c.font = hf(9); c.fill = fill(MID_BLUE); c.alignment = CTR; c.border = bdr()
    ws.row_dimensions[7].height = 32
    ws.freeze_panes = "A8"

    # ── Rows 8+: data ─────────────────────────────────────────────────────
    DATA_START = 8
    for ri, row in enumerate(rows):
        r  = DATA_START + ri
        bg = PALE_BLUE if ri % 2 == 0 else WHITE
        col = 1

        # Identity columns
        if agency_key:
            c = ws.cell(r, col, row.get(agency_key, ""))
            c.font = bf(); c.alignment = LEFT; c.fill = fill(bg); c.border = bdr(); col += 1
        c = ws.cell(r, col, row.get(id_key, ""))
        c.font = bf(bold=True); c.alignment = LEFT; c.fill = fill(bg); c.border = bdr(); col += 1

        tv_col  = col
        gp_col  = col + 1
        gpp_col = col + 2
        adj_col = col + 3
        tvc_col = col + 4
        eov_col = col + 5   # EOY TV
        eog_col = col + 6   # EOY GP
        eopc    = col + 7   # EOY GP%
        adj_gp  = col + 8   # Adj GP
        delta   = col + 9   # Delta

        # YTD TV
        c = ws.cell(r, tv_col, row.get("tv", 0) or 0)
        c.font = bf(); c.alignment = RGHT; c.fill = fill(bg)
        c.border = bdr(); c.number_format = AED

        # YTD GP
        c = ws.cell(r, gp_col, row.get("gp", 0) or 0)
        c.font = bf(); c.alignment = RGHT; c.fill = fill(bg)
        c.border = bdr(); c.number_format = AED

        # YTD GP%
        c = ws.cell(r, gpp_col, row.get("gp_pct", 0) or 0)
        c.font = bf(); c.alignment = CTR; c.fill = fill(bg)
        c.border = bdr(); c.number_format = PCT

        # GP% Adj (editable input, gold)
        c = ws.cell(r, adj_col, 0)
        c.font = Font(name="Arial", size=9, color="0000FF")
        c.alignment = CTR; c.fill = fill(GOLD); c.border = bdr(); c.number_format = PCT

        # TV Chg % (editable input, gold)
        c = ws.cell(r, tvc_col, 0)
        c.font = Font(name="Arial", size=9, color="0000FF")
        c.alignment = CTR; c.fill = fill(GOLD); c.border = bdr(); c.number_format = PCT

        # EOY TV = YTD_TV ÷ ytd_wt
        tv_l  = gl(tv_col); gp_l = gl(gp_col); gpp_l = gl(gpp_col)
        adj_l = gl(adj_col); tvc_l = gl(tvc_col)
        eov_l = gl(eov_col); eog_l = gl(eog_col); eop_l = gl(eopc)
        adjgp_l = gl(adj_gp)

        c = ws.cell(r, eov_col,
            f"=IFERROR({tv_l}{r}/IF({ytd_wt}=0,1,{ytd_wt})*(1+{tvc_l}{r}),0)")
        c.font = bf(bold=True, color="006100"); c.alignment = RGHT
        c.fill = fill(LIGHT_GRN); c.border = bdr(); c.number_format = AED

        # EOY GP = EOY_TV * GP%
        c = ws.cell(r, eog_col,
            f"=IFERROR({eov_l}{r}*{gpp_l}{r},0)")
        c.font = bf(bold=True, color="006100"); c.alignment = RGHT
        c.fill = fill(LIGHT_GRN); c.border = bdr(); c.number_format = AED

        # EOY GP%
        c = ws.cell(r, eopc,
            f"={gpp_l}{r}")
        c.font = bf(color="006100"); c.alignment = CTR
        c.fill = fill(LIGHT_GRN); c.border = bdr(); c.number_format = PCT

        # Adj. EOY GP = EOY_TV * (GP% + Adj)
        c = ws.cell(r, adj_gp,
            f"=IFERROR({eov_l}{r}*({gpp_l}{r}+{adj_l}{r}),0)")
        c.font = bf(bold=True); c.alignment = RGHT
        c.fill = fill(LIGHT_BLUE); c.border = bdr(); c.number_format = AED

        # GP Delta
        c = ws.cell(r, delta,
            f"={adjgp_l}{r}-{eog_l}{r}")
        c.font = bf(); c.alignment = RGHT
        c.fill = fill(LIGHT_BLUE); c.border = bdr(); c.number_format = GP_DELTA

    # ── TOTAL row ─────────────────────────────────────────────────────────
    tr = DATA_START + len(rows)
    ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=n_id)
    c = ws.cell(tr, 1, "TOTAL")
    c.font = hf(9); c.fill = fill(DARK_BLUE); c.alignment = LEFT; c.border = bdr()

    tv_col  = n_id + 1
    sum_cols   = {tv_col: AED, tv_col+1: AED, tv_col+5: AED, tv_col+6: AED,
                  tv_col+8: AED, tv_col+9: GP_DELTA}
    avg_cols   = {tv_col+2: PCT, tv_col+7: PCT}
    blank_cols = {tv_col+3, tv_col+4}   # input cols — leave blank in total

    for ci in range(tv_col, tv_col + 10):
        c = ws.cell(tr, ci)
        if ci in blank_cols:
            c.value = ""
        elif ci in sum_cols:
            c.value = f"=SUM({gl(ci)}{DATA_START}:{gl(ci)}{tr-1})"
            c.number_format = sum_cols[ci]
        elif ci in avg_cols:
            c.value = f"=IFERROR(AVERAGE({gl(ci)}{DATA_START}:{gl(ci)}{tr-1}),0)"
            c.number_format = avg_cols[ci]
        c.font = hf(9); c.fill = fill(DARK_BLUE); c.alignment = RGHT; c.border = bdr()
    ws.row_dimensions[tr].height = 18

    # ── Notes row ─────────────────────────────────────────────────────────
    nr = tr + 1
    ws.merge_cells(start_row=nr, start_column=1, end_row=nr, end_column=N_COLS)
    note = (f"YTD TV and GP are editable. EOY GP (Base) = YTD_TV ÷ YTD_factor × GP%  "
            f"(no adjustments).  TV Chg% shifts volume; GP% Adj shifts margin.  "
            f"Refreshed: {today_str}")
    c = ws.cell(nr, 1, note)
    c.font = bf(sz=8, color=DARK_GREY); c.alignment = LEFT
    ws.row_dimensions[nr].height = 14

    # ── Opportunity Insights ───────────────────────────────────────────────
    ir = nr + 2
    ws.merge_cells(start_row=ir, start_column=1, end_row=ir, end_column=N_COLS)
    c = ws.cell(ir, 1, f"  ★  OPPORTUNITY INSIGHTS  —  Where managers can act")
    c.font = hf(10); c.fill = fill(NAVY); c.alignment = LEFT
    ws.row_dimensions[ir].height = 20

    ir2 = ir + 1
    ws.merge_cells(start_row=ir2, start_column=1, end_row=ir2, end_column=N_COLS)
    c = ws.cell(ir2, 1,
        f"Company avg GP%: {avg_gp_pct*100:.1f}%   |   "
        "Panel A: Low-margin accounts where pushing GP% to avg would add the most GP.   "
        "Panel B: High-margin accounts where growing TV would be highly profitable.   "
        "Panel C: Accounts with the biggest absolute GP gap vs avg.")
    c.font = bf(sz=8, color=DARK_GREY); c.alignment = LEFT
    ws.row_dimensions[ir2].height = 14

    # Panel headers
    ir3 = ir2 + 1
    panel_hdrs = [("A — Improve Margin", 1, 3), ("B — Grow Volume", 4, 6), ("C — Biggest GP Gap", 7, N_COLS)]
    for label, c1, c2 in panel_hdrs:
        ws.merge_cells(start_row=ir3, start_column=c1, end_row=ir3, end_column=c2)
        c = ws.cell(ir3, c1, label)
        c.font = hf(9); c.fill = fill(MID_BLUE); c.alignment = CTR
    ws.row_dimensions[ir3].height = 18

    ir4 = ir3 + 1
    desc_hdrs = [
        ("TV is large but GP% is below avg. Raising GP% to avg = big win.", 1, 3),
        ("High-margin, low-volume. Growing TV would be highly profitable.", 4, 6),
        ("Largest gap between current GP and what avg margin would generate.", 7, N_COLS),
    ]
    for label, c1, c2 in desc_hdrs:
        ws.merge_cells(start_row=ir4, start_column=c1, end_row=ir4, end_column=c2)
        c = ws.cell(ir4, c1, label)
        c.font = bf(sz=8, color=DARK_GREY); c.alignment = LEFT
    ws.row_dimensions[ir4].height = 14

    # Panel column headers
    ir5 = ir4 + 1
    for ci, h in enumerate(["Name","YTD TV","GP%","Opportunity","Name","YTD TV","GP%",
                              "Name","YTD TV","GP%","Opportunity"], 1):
        if ci <= N_COLS:
            c = ws.cell(ir5, ci, h)
            c.font = hf(9); c.fill = fill(MID_BLUE); c.alignment = CTR; c.border = bdr()
    ws.row_dimensions[ir5].height = 18

    # Panel A: low-margin, large-TV (below avg GP%, sorted by potential GP upside)
    eligible = [r for r in rows if r["gp_pct"] < avg_gp_pct and r["tv"] > 0]
    panel_a = sorted(eligible, key=lambda x: -(x["tv"] * (avg_gp_pct - x["gp_pct"])))[:5]
    # Panel B: high-margin, lower TV
    panel_b = sorted([r for r in rows if r["gp_pct"] > avg_gp_pct],
                     key=lambda x: x["tv"])[:5]
    # Panel C: biggest abs GP gap
    panel_c = sorted(rows, key=lambda x: -(abs(x["tv"] * avg_gp_pct - x["gp"])))[:5]

    max_p = max(len(panel_a), len(panel_b), len(panel_c))
    for pi in range(max_p):
        pr = ir5 + 1 + pi
        bg = LIGHT_GREY if pi % 2 == 0 else WHITE

        if pi < len(panel_a):
            ra = panel_a[pi]
            upside = ra["tv"] * (avg_gp_pct - ra["gp_pct"])
            for ci, val, fmt in [(1, ra[id_key], None), (2, ra["tv"], AED),
                                  (3, ra["gp_pct"], PCT),
                                  (4, f"{upside:,.0f} AED GP upside", None)]:
                c = ws.cell(pr, ci, val)
                c.font = bf(sz=8); c.fill = fill(bg); c.border = bdr()
                if fmt: c.number_format = fmt

        if pi < len(panel_b):
            rb = panel_b[pi]
            for ci, val, fmt in [(5, rb[id_key], None), (6, rb["tv"], AED), (7, rb["gp_pct"], PCT)]:
                c = ws.cell(pr, ci, val)
                c.font = bf(sz=8); c.fill = fill(bg); c.border = bdr()
                if fmt: c.number_format = fmt

        if pi < len(panel_c):
            rc = panel_c[pi]
            gap = rc["tv"] * avg_gp_pct - rc["gp"]
            for ci, val, fmt in [(8, rc[id_key], None), (9, rc["tv"], AED),
                                  (10, rc["gp_pct"], PCT),
                                  (11, f"{gap:,.0f} AED {'gap to avg' if gap > 0 else 'above avg'}", None)]:
                if ci <= N_COLS:
                    c = ws.cell(pr, ci, val)
                    c.font = bf(sz=8); c.fill = fill(bg); c.border = bdr()
                    if fmt: c.number_format = fmt

        ws.row_dimensions[pr].height = 14

    # ── Column widths ──────────────────────────────────────────────────────
    if agency_key:
        widths = [16, 28, 14, 14, 9, 9, 9, 16, 14, 9, 14, 14]
    else:
        widths = [28, 14, 14, 9, 9, 9, 16, 14, 9, 14, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[gl(i)].width = w


def build_dashboard(ws, today_str, data_month, ag_rows, cu_rows, de_rows, ytd_wt):
    """Main dashboard tab."""
    ws.sheet_view.showGridLines = False

    ytd_tv = sum(r["tv"] for r in ag_rows)
    ytd_gp = sum(r["gp"] for r in ag_rows)
    gp_pct = ytd_gp / ytd_tv if ytd_tv else 0
    eoy_tv = ytd_tv / ytd_wt if ytd_wt else 0
    eoy_gp = ytd_gp / ytd_wt if ytd_wt else 0

    # Row 1: Title
    ws.merge_cells("A1:P1")
    c = ws.cell(1, 1,
        f"ELEVATE DMC — PROFITABILITY MODELLER 2026  |  Refreshed: {today_str}  |  "
        f"Booking Data through: {data_month}")
    c.font = hf(13); c.fill = fill(DARK_BLUE); c.alignment = CTR
    ws.row_dimensions[1].height = 30

    # Row 2: blank
    ws.row_dimensions[2].height = 8

    # Row 3: KPI labels
    kpi_labels = ["YTD Total Value","","","YTD Gross Profit","","","YTD GP%","","",
                  "EOY TV Forecast","","","EOY GP Forecast","","","EOY GP% Forecast"]
    clrs3 = [MID_BLUE]*9 + [GRN_HDR]*9
    for ci, (lbl, clr) in enumerate(zip(kpi_labels, clrs3), 1):
        c = ws.cell(3, ci, lbl)
        c.font = hf(9); c.fill = fill(clr); c.alignment = CTR; c.border = bdr()
    for start, end in [(1,3),(4,6),(7,9),(10,12),(13,15),(16,16)]:
        ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=end)
    ws.row_dimensions[3].height = 20

    # Row 4: KPI values
    kpi_vals = [ytd_tv,"","",ytd_gp,"","",gp_pct,"","",eoy_tv,"","",eoy_gp,"","",gp_pct]
    kpi_fmts = [AED,"","",AED,"","",PCT,"","",AED,"","",AED,"","",PCT]
    for ci, (val, fmt) in enumerate(zip(kpi_vals, kpi_fmts), 1):
        if val == "":
            c = ws.cell(4, ci, ""); c.border = bdr()
        else:
            c = ws.cell(4, ci, val)
            c.font = hf(12, bold=True); c.fill = fill(WHITE)
            c.alignment = RGHT; c.number_format = fmt; c.border = bdr()
    ws.row_dimensions[4].height = 28

    # Rows 5-6: blank
    for r in [5, 6]: ws.row_dimensions[r].height = 8

    # Row 7: How-to-use note
    ws.merge_cells("A7:P7")
    c = ws.cell(7, 1,
        "HOW TO USE: Edit TV/GP in Agency Groups / Customer / Destination tabs → KPI banner auto-updates.  "
        "GP% Adj & TV Chg % inputs let you model scenarios.  Seasonality weights drive EOY projections.")
    c.font = bf(sz=8, color=DARK_GREY); c.fill = fill(LIGHT_GREY); c.alignment = LEFT
    ws.row_dimensions[7].height = 14

    # Rows 8-10: blank
    for r in [8, 9, 10]: ws.row_dimensions[r].height = 8

    # Top-5 agency table (rows 11-18)
    ws.merge_cells("A11:F11")
    c = ws.cell(11, 1, "TOP AGENCY GROUPS  —  YTD Performance")
    c.font = hf(10); c.fill = fill(DARK_BLUE); c.alignment = CTR
    ws.row_dimensions[11].height = 20

    top_hdrs = ["Agency Group","YTD TV (AED)","YTD GP (AED)","YTD GP%","EOY TV","EOY GP"]
    for ci, h in enumerate(top_hdrs, 1):
        c = ws.cell(12, ci, h)
        c.font = hf(9); c.fill = fill(MID_BLUE); c.alignment = CTR; c.border = bdr()
    ws.row_dimensions[12].height = 22

    top5_ag = sorted(ag_rows, key=lambda x: -(x["tv"] or 0))[:5]
    for ri, row in enumerate(top5_ag):
        r  = 13 + ri
        bg = PALE_BLUE if ri % 2 == 0 else WHITE
        eoy_tv_row = (row["tv"] / ytd_wt) if ytd_wt else 0
        eoy_gp_row = (row["gp"] / ytd_wt) if ytd_wt else 0
        for ci, (val, fmt) in enumerate([
            (row["agency"], None),(row["tv"], AED),(row["gp"], AED),
            (row["gp_pct"], PCT),(eoy_tv_row, AED),(eoy_gp_row, AED)
        ], 1):
            c = ws.cell(r, ci, val)
            c.font = bf(); c.fill = fill(bg); c.border = bdr()
            if fmt: c.number_format = fmt
            c.alignment = RGHT if fmt else LEFT
        ws.row_dimensions[r].height = 16

    # Top-5 destinations table (cols H–M)
    ws.merge_cells("H11:M11")
    c = ws.cell(11, 8, "TOP DESTINATIONS  —  YTD Performance")
    c.font = hf(10); c.fill = fill(DARK_BLUE); c.alignment = CTR

    top_hdrs2 = ["Country","YTD TV (AED)","YTD GP (AED)","YTD GP%","EOY TV","EOY GP"]
    for ci, h in enumerate(top_hdrs2, 1):
        c = ws.cell(12, 7+ci, h)
        c.font = hf(9); c.fill = fill(MID_BLUE); c.alignment = CTR; c.border = bdr()

    top5_de = sorted(de_rows, key=lambda x: -(x["tv"] or 0))[:5]
    for ri, row in enumerate(top5_de):
        r  = 13 + ri
        bg = PALE_BLUE if ri % 2 == 0 else WHITE
        eoy_tv_row = (row["tv"] / ytd_wt) if ytd_wt else 0
        eoy_gp_row = (row["gp"] / ytd_wt) if ytd_wt else 0
        for ci, (val, fmt) in enumerate([
            (row["country"], None),(row["tv"], AED),(row["gp"], AED),
            (row["gp_pct"], PCT),(eoy_tv_row, AED),(eoy_gp_row, AED)
        ], 1):
            c = ws.cell(r, 7+ci, val)
            c.font = bf(); c.fill = fill(bg); c.border = bdr()
            if fmt: c.number_format = fmt
            c.alignment = RGHT if fmt else LEFT

    # Col widths
    for ci, w in enumerate([22,16,16,9,16,16,4,18,16,16,9,16,16], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w


# ── main rebuild ──────────────────────────────────────────────────────────────
def rebuild(d1_bytes, d2_bytes):
    wb1 = openpyxl.load_workbook(io.BytesIO(d1_bytes), data_only=True)
    wb2 = openpyxl.load_workbook(io.BytesIO(d2_bytes), data_only=True)

    ag_rows, cu_rows = load_d1(wb1)
    de_rows           = load_d2(wb2)

    today      = datetime.date.today()
    today_str  = today.strftime("%d %b %Y")
    ytd_m      = today.month
    ytd_wt     = sum(WEIGHTS[:ytd_m])
    data_month = MONTHS[ytd_m - 1] + " 2026"

    wb = openpyxl.Workbook()

    # Sheet order: Dashboard, Seasonality, Agency Groups, Customer, Destination
    db = wb.active; db.title = "Dashboard"
    ss = wb.create_sheet("Seasonality")
    ag = wb.create_sheet("Agency Groups")
    cu = wb.create_sheet("Customer")
    de = wb.create_sheet("Destination")

    build_dashboard(db, today_str, data_month, ag_rows, cu_rows, de_rows, ytd_wt)
    build_seasonality(ss, today_str, ag_rows)
    build_analysis_sheet(ag, "AGENCY GROUPS ANALYSIS", ag_rows,
                         id_key="agency", id_label="Agency Group",
                         ytd_wt=ytd_wt, today_str=today_str)
    build_analysis_sheet(cu, "CUSTOMER ANALYSIS", cu_rows,
                         id_key="customer", id_label="Customer",
                         agency_key="agency", ytd_wt=ytd_wt, today_str=today_str)
    build_analysis_sheet(de, "DESTINATION ANALYSIS", de_rows,
                         id_key="country", id_label="Country",
                         ytd_wt=ytd_wt, today_str=today_str)

    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.read()


# ── Flask routes ──────────────────────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "profitability-modeller"})


@app.route("/rebuild", methods=["POST"])
def rebuild_endpoint():
    """Accept base64-encoded files, return rebuilt modeller as base64."""
    try:
        body     = request.get_json()
        d1_bytes = base64.b64decode(body["data1_b64"])
        d2_bytes = base64.b64decode(body["data2_b64"])
        result   = rebuild(d1_bytes, d2_bytes)
        return jsonify({"modeller_b64": base64.b64encode(result).decode(), "status": "ok"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/rebuild-from-drive", methods=["POST"])
def rebuild_from_drive():
    """
    Read data_1.xlsx + data_2.xlsx from a Google Drive folder,
    rebuild the modeller, and save Profitability_Modeller_2026.xlsx back.

    Expects JSON body:
      {
        "access_token": "...",        # OAuth2 token with Drive scope
        "folder_id":    "...",        # Google Drive folder ID
        "output_name":  "Profitability_Modeller_2026.xlsx"  # optional
      }
    """
    import requests as req

    try:
        body         = request.get_json()
        token        = body["access_token"]
        folder_id    = body["folder_id"]
        output_name  = body.get("output_name", "Profitability_Modeller_2026.xlsx")

        headers = {"Authorization": f"Bearer {token}"}

        def find_file(name):
            q   = f"'{folder_id}' in parents and name='{name}' and trashed=false"
            r   = req.get(
                "https://www.googleapis.com/drive/v3/files",
                params={"q": q, "fields": "files(id,name)", "pageSize": 1},
                headers=headers, timeout=30
            )
            r.raise_for_status()
            files = r.json().get("files", [])
            if not files:
                raise FileNotFoundError(f"{name} not found in Drive folder")
            return files[0]["id"]

        def download_file(file_id):
            r = req.get(
                f"https://www.googleapis.com/drive/v3/files/{file_id}",
                params={"alt": "media"},
                headers=headers, timeout=60
            )
            r.raise_for_status()
            return r.content

        # Download source data
        d1_id = find_file("data_1.xlsx")
        d2_id = find_file("data_2.xlsx")
        d1    = download_file(d1_id)
        d2    = download_file(d2_id)

        # Rebuild
        result = rebuild(d1, d2)

        # Check if output file already exists in folder
        q_out = f"'{folder_id}' in parents and name='{output_name}' and trashed=false"
        r_out = req.get(
            "https://www.googleapis.com/drive/v3/files",
            params={"q": q_out, "fields": "files(id)", "pageSize": 1},
            headers=headers, timeout=30
        )
        r_out.raise_for_status()
        existing = r_out.json().get("files", [])

        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        if existing:
            # Update existing file (PATCH)
            file_id_out = existing[0]["id"]
            up = req.patch(
                f"https://www.googleapis.com/upload/drive/v3/files/{file_id_out}",
                params={"uploadType": "media"},
                headers={**headers, "Content-Type": mime},
                data=result, timeout=120
            )
        else:
            # Create new file
            import json
            meta = json.dumps({"name": output_name, "parents": [folder_id]})
            up = req.post(
                "https://www.googleapis.com/upload/drive/v3/files",
                params={"uploadType": "multipart"},
                headers=headers,
                files={
                    "metadata": ("meta", meta, "application/json"),
                    "file":     ("file", result, mime)
                },
                timeout=120
            )
        up.raise_for_status()

        return jsonify({
            "status":      "ok",
            "output_file": output_name,
            "folder_id":   folder_id,
            "refreshed":   datetime.date.today().strftime("%d %b %Y")
        })

    except FileNotFoundError as e:
        return jsonify({"error": str(e), "hint": "Make sure data_1.xlsx and data_2.xlsx are in the Drive folder"}), 404
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
