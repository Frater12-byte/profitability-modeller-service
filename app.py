from flask import Flask, request, jsonify
import base64, io, os, datetime, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ── colours ───────────────────────────────────────────────────────────────────
DARK_BLUE="1F3864"; MID_BLUE="2E5FA3"; LIGHT_BLUE="BDD7EE"; PALE_BLUE="DEEAF1"
GOLD="FFF2CC"; GRN_HDR="375623"; LIGHT_GRN="E2EFDA"; WHITE="FFFFFF"
LIGHT_GREY="F2F2F2"; DARK_GREY="595959"

def fill(h): return PatternFill("solid", fgColor=h)
def hf(sz=10,bold=True,color=WHITE): return Font(name="Arial",size=sz,bold=bold,color=color)
def bf(sz=9,bold=False,color="000000"): return Font(name="Arial",size=sz,bold=bold,color=color)
def bdr():
    s=Side(style="thin",color="BFBFBF")
    return Border(left=s,right=s,top=s,bottom=s)
CTR=Alignment(horizontal="center",vertical="center",wrap_text=True)
LEFT=Alignment(horizontal="left",vertical="center")
RGHT=Alignment(horizontal="right",vertical="center")

def safe(v):
    if v is None or v in ('#VALUE!','#REF!','#N/A'): return None
    if isinstance(v, datetime.datetime): return None
    return float(v) if isinstance(v,(int,float)) else None

def load_d1(wb):
    ws=wb.active; rows=[]; cur=None
    for r in range(2,ws.max_row+1):
        ag=ws.cell(r,1).value; cu=ws.cell(r,2).value
        if ag: cur=ag
        if not cu: continue
        row=dict(agency=cur,customer=cu,
                 gp=safe(ws.cell(r,10).value),gp_py=safe(ws.cell(r,11).value),
                 gp_pct=safe(ws.cell(r,12).value),rnts=safe(ws.cell(r,15).value),
                 bookings=safe(ws.cell(r,18).value))
        if cu=='Total': rows.append({**row,'type':'agency'})
        else: rows.append({**row,'type':'customer'})
    return [r for r in rows if r['type']=='agency'], [r for r in rows if r['type']=='customer']

def load_d2(wb):
    ws=wb.active; rows=[]
    for r in range(2,ws.max_row+1):
        co=ws.cell(r,1).value
        if not co: continue
        rows.append(dict(country=co,gp=safe(ws.cell(r,9).value),
                         gp_py=safe(ws.cell(r,10).value),gp_pct=safe(ws.cell(r,11).value),
                         rnts=safe(ws.cell(r,14).value),bookings=safe(ws.cell(r,17).value)))
    return rows

def build_seasonality(ws):
    ws.freeze_panes="A3"
    ws.merge_cells("A1:D1")
    c=ws.cell(1,1,"SEASONALITY ASSUMPTIONS"); c.font=hf(12); c.fill=fill(DARK_BLUE); c.alignment=CTR
    ws.row_dimensions[1].height=26
    for ci,h in enumerate(["Month","Monthly Factor","Cumulative YTD","Mark Current Month (1)"],1):
        c=ws.cell(2,ci,h); c.font=hf(9); c.fill=fill(MID_BLUE); c.alignment=CTR; c.border=bdr()
    ws.row_dimensions[2].height=32
    months=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    factors=[0.05,0.06,0.07,0.09,0.10,0.11,0.10,0.10,0.10,0.09,0.08,0.05]
    for i,(m,f) in enumerate(zip(months,factors)):
        r=i+3; bg=LIGHT_GREY if i%2==0 else WHITE
        c=ws.cell(r,1,m); c.font=bf(bold=True); c.alignment=LEFT; c.fill=fill(bg); c.border=bdr()
        c=ws.cell(r,2,f); c.font=Font(name="Arial",size=9,color="0000FF"); c.fill=fill(GOLD)
        c.alignment=CTR; c.number_format="0.0%"; c.border=bdr()
        c=ws.cell(r,3,f"=SUM($B$3:B{r})"); c.font=bf(color="006100"); c.alignment=CTR
        c.number_format="0.0%"; c.border=bdr()
        c=ws.cell(r,4,1 if m=="Sep" else ""); c.font=Font(name="Arial",size=9,color="0000FF")
        c.fill=fill(GOLD); c.alignment=CTR; c.border=bdr()
    ws.merge_cells("A16:C16"); ws.cell(16,1,"Current YTD Cumulative Factor:").font=bf(bold=True)
    ws.cell(16,1).alignment=LEFT
    ws.cell(16,4,"=IFERROR(INDEX(C3:C14,MATCH(1,D3:D14,0)),INDEX(C3:C14,MATCH(MAX(D3:D14),D3:D14,0)))")
    ws.cell(16,4).number_format="0.0%"; ws.cell(16,4).font=bf(bold=True,color="006100")
    for ci,w in enumerate([12,16,18,22],1): ws.column_dimensions[get_column_letter(ci)].width=w

def build_analysis(ws, title, rows, id_key, id_label, agency_key=None):
    n_id=2 if agency_key else 1; SREF="Seasonality!$D$16"
    total_cols=n_id+13
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=total_cols)
    c=ws.cell(1,1,title); c.font=hf(13); c.fill=fill(DARK_BLUE); c.alignment=CTR
    ws.row_dimensions[1].height=28
    sections=[(n_id,"",DARK_BLUE),(4,"YTD ACTUALS",MID_BLUE),(1,"SEASON",DARK_GREY),
              (3,"EOY FORECAST",GRN_HDR),(2,"SCENARIO INPUTS","7F3F00"),(3,"SCENARIO RESULTS","375E23")]
    col=1
    for span,label,clr in sections:
        if span>1: ws.merge_cells(start_row=2,start_column=col,end_row=2,end_column=col+span-1)
        c=ws.cell(2,col,label); c.font=hf(9); c.fill=fill(clr); c.alignment=CTR; col+=span
    ws.row_dimensions[2].height=18
    hdrs=([" Agency Group",id_label] if agency_key else [id_label])+[
        "YTD GP (AED)","YTD GP%","YTD RNTs","YTD Bookings","YTD Factor",
        "EOY GP Forecast","EOY GP%","EOY RNTs","GP% Adj (pp)","TV Change %",
        "Scenario GP","Scenario GP%","GP Delta vs Forecast"]
    for ci,h in enumerate(hdrs,1):
        c=ws.cell(3,ci,h); c.font=hf(9); c.fill=fill(MID_BLUE); c.alignment=CTR; c.border=bdr()
    ws.row_dimensions[3].height=32; ws.freeze_panes="A4"
    DATA_START=4; gl=get_column_letter
    for ri,row in enumerate(rows):
        r=DATA_START+ri; bg=PALE_BLUE if ri%2==0 else WHITE; col=1
        if agency_key:
            c=ws.cell(r,col,row.get(agency_key,"")); c.font=bf(); c.alignment=LEFT
            c.fill=fill(bg); c.border=bdr(); col+=1
        c=ws.cell(r,col,row.get(id_key,"")); c.font=bf(bold=True); c.alignment=LEFT
        c.fill=fill(bg); c.border=bdr(); col+=1
        gp_col=col; c=ws.cell(r,col,row.get('gp') or 0)
        c.font=bf(); c.alignment=RGHT; c.fill=fill(bg); c.border=bdr(); c.number_format='#,##0;(#,##0);"-"'; col+=1
        gp_pct_col=col; c=ws.cell(r,col,row.get('gp_pct') or 0)
        c.font=bf(); c.alignment=CTR; c.fill=fill(bg); c.border=bdr(); c.number_format='0.0%'; col+=1
        rnts_col=col; c=ws.cell(r,col,row.get('rnts') or 0)
        c.font=bf(); c.alignment=RGHT; c.fill=fill(bg); c.border=bdr(); c.number_format='#,##0;(#,##0);"-"'; col+=1
        c=ws.cell(r,col,row.get('bookings') or 0); c.font=bf(); c.alignment=RGHT
        c.fill=fill(bg); c.border=bdr(); c.number_format='#,##0;(#,##0);"-"'; col+=1
        sf_col=col; c=ws.cell(r,col,f"={SREF}"); c.font=Font(name="Arial",size=9,color="00763D")
        c.alignment=CTR; c.fill=fill(bg); c.border=bdr(); c.number_format='0.0%'; col+=1
        eoy_gp_col=col
        c=ws.cell(r,col,f"=IFERROR({gl(gp_col)}{r}/IF({gl(sf_col)}{r}=0,1,{gl(sf_col)}{r}),0)")
        c.font=bf(bold=True,color="006100"); c.alignment=RGHT; c.fill=fill(LIGHT_GRN)
        c.border=bdr(); c.number_format='#,##0;(#,##0);"-"'; col+=1
        c=ws.cell(r,col,f"={gl(gp_pct_col)}{r}"); c.font=bf(color="006100"); c.alignment=CTR
        c.fill=fill(LIGHT_GRN); c.border=bdr(); c.number_format='0.0%'; col+=1
        c=ws.cell(r,col,f"=IFERROR({gl(rnts_col)}{r}/IF({gl(sf_col)}{r}=0,1,{gl(sf_col)}{r}),0)")
        c.font=bf(color="006100"); c.alignment=RGHT; c.fill=fill(LIGHT_GRN)
        c.border=bdr(); c.number_format='#,##0;(#,##0);"-"'; col+=1
        gp_adj_col=col; c=ws.cell(r,col,0); c.font=Font(name="Arial",size=9,color="0000FF")
        c.alignment=CTR; c.fill=fill(GOLD); c.border=bdr(); c.number_format='0.0%;(0.0%);"-"'; col+=1
        tv_col=col; c=ws.cell(r,col,0); c.font=Font(name="Arial",size=9,color="0000FF")
        c.alignment=CTR; c.fill=fill(GOLD); c.border=bdr(); c.number_format='0.0%;(0.0%);"-"'; col+=1
        scen_gp_col=col
        f_scen=(f"=IFERROR({gl(eoy_gp_col)}{r}*(1+{gl(tv_col)}{r})"
                f"*({gl(gp_pct_col)}{r}+{gl(gp_adj_col)}{r})"
                f"/IF({gl(gp_pct_col)}{r}=0,1,{gl(gp_pct_col)}{r}),0)")
        c=ws.cell(r,col,f_scen); c.font=bf(bold=True); c.alignment=RGHT
        c.fill=fill(LIGHT_BLUE); c.border=bdr(); c.number_format='#,##0;(#,##0);"-"'; col+=1
        c=ws.cell(r,col,f"={gl(gp_pct_col)}{r}+{gl(gp_adj_col)}{r}"); c.font=bf()
        c.alignment=CTR; c.fill=fill(LIGHT_BLUE); c.border=bdr(); c.number_format='0.0%'; col+=1
        c=ws.cell(r,col,f"={gl(scen_gp_col)}{r}-{gl(eoy_gp_col)}{r}"); c.font=bf()
        c.alignment=RGHT; c.fill=fill(LIGHT_BLUE); c.border=bdr()
        c.number_format='[Green]+#,##0;[Red]-#,##0;"-"'
    tr=DATA_START+len(rows)
    ws.merge_cells(start_row=tr,start_column=1,end_row=tr,end_column=n_id)
    c=ws.cell(tr,1,"TOTAL"); c.font=hf(9); c.fill=fill(DARK_BLUE); c.alignment=LEFT; c.border=bdr()
    sum_map={n_id+1:'#,##0;(#,##0);"-"',n_id+3:'#,##0;(#,##0);"-"',n_id+4:'#,##0;(#,##0);"-"',
             n_id+6:'#,##0;(#,##0);"-"',n_id+8:'#,##0;(#,##0);"-"',
             n_id+11:'#,##0;(#,##0);"-"',n_id+13:'[Green]+#,##0;[Red]-#,##0;"-"'}
    for ci,fmt in sum_map.items():
        c=ws.cell(tr,ci,f"=SUM({gl(ci)}{DATA_START}:{gl(ci)}{tr-1})")
        c.font=hf(9); c.fill=fill(DARK_BLUE); c.alignment=RGHT; c.border=bdr(); c.number_format=fmt
    for ci,fmt in {n_id+2:'0.0%',n_id+7:'0.0%',n_id+12:'0.0%'}.items():
        c=ws.cell(tr,ci,f"=IFERROR(AVERAGE({gl(ci)}{DATA_START}:{gl(ci)}{tr-1}),0)")
        c.font=hf(9); c.fill=fill(DARK_BLUE); c.alignment=CTR; c.border=bdr(); c.number_format=fmt
    widths=([18] if agency_key else [])+[28,14,9,11,11,9,14,9,11,11,10,14,9,14]
    for i,w in enumerate(widths,1): ws.column_dimensions[gl(i)].width=w

def rebuild(d1_bytes, d2_bytes):
    wb1=openpyxl.load_workbook(io.BytesIO(d1_bytes),data_only=True)
    wb2=openpyxl.load_workbook(io.BytesIO(d2_bytes),data_only=True)
    ag_rows,cu_rows=load_d1(wb1)
    de_rows=load_d2(wb2)
    today=datetime.date.today().strftime("%d %b %Y")
    wb=openpyxl.Workbook()
    ss=wb.active; ss.title="Seasonality"; build_seasonality(ss)
    ag=wb.create_sheet("Agency"); build_analysis(ag,"AGENCY ANALYSIS",ag_rows,id_key='agency',id_label='Agency Group')
    cu=wb.create_sheet("Customer"); build_analysis(cu,"CUSTOMER ANALYSIS",cu_rows,id_key='customer',id_label='Customer',agency_key='agency')
    de=wb.create_sheet("Destination"); build_analysis(de,"DESTINATION ANALYSIS",de_rows,id_key='country',id_label='Country')
    # Simple dashboard
    db=wb.create_sheet("Dashboard",0)
    db.merge_cells("A1:P1")
    c=db.cell(1,1,f"PROFITABILITY MODELLER 2026 | Refreshed: {today}")
    c.font=hf(13); c.fill=fill(DARK_BLUE); c.alignment=CTR; db.row_dimensions[1].height=30
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return out.read()

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status':'ok','service':'profitability-modeller'})

@app.route('/rebuild', methods=['POST'])
def rebuild_endpoint():
    try:
        body=request.get_json()
        d1_bytes=base64.b64decode(body['data1_b64'])
        d2_bytes=base64.b64decode(body['data2_b64'])
        result_bytes=rebuild(d1_bytes, d2_bytes)
        return jsonify({'modeller_b64': base64.b64encode(result_bytes).decode(), 'status':'ok'})
    except Exception as e:
        return jsonify({'error':str(e)}), 500

if __name__=='__main__':
    port=int(os.environ.get('PORT',8080))
    app.run(host='0.0.0.0', port=port)
