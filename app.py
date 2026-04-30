"""
Y&S Tickets — Installment Projections
Streamlit web app: drop in a General Ledger, download the report.
"""

import io
import re
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Installment Projections",
    page_icon="🎟️",
    layout="centered",
)

# ── CUSTOM CSS ────────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Page background */
.stApp {
    background: #F5F3EF;
}

/* Hide default Streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }

/* Hero section */
.hero {
    text-align: center;
    padding: 3rem 0 2rem;
}
.hero h1 {
    font-family: 'DM Serif Display', serif;
    font-size: 2.8rem;
    color: #1F3864;
    margin: 0;
    line-height: 1.1;
    letter-spacing: -0.5px;
}
.hero h1 em {
    font-style: italic;
    color: #2E5EAA;
}
.hero p {
    font-size: 1.05rem;
    color: #666;
    margin-top: 0.6rem;
    font-weight: 300;
}

/* Upload card */
.upload-card {
    background: white;
    border-radius: 16px;
    padding: 2.5rem 2rem;
    box-shadow: 0 2px 20px rgba(31,56,100,0.08);
    border: 1.5px solid #E8E4DC;
    margin: 1.5rem 0;
}

/* Step labels */
.step-label {
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #2E5EAA;
    margin-bottom: 0.5rem;
}

/* Success box */
.success-box {
    background: #EEF4FB;
    border-left: 4px solid #2E5EAA;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    margin: 1rem 0;
}
.success-box .file-title {
    font-family: 'DM Serif Display', serif;
    font-size: 1.1rem;
    color: #1F3864;
    margin: 0;
}
.success-box .file-meta {
    font-size: 0.85rem;
    color: #555;
    margin: 0.2rem 0 0;
}

/* Stats row */
.stats-row {
    display: flex;
    gap: 1rem;
    margin: 1rem 0;
}
.stat-card {
    flex: 1;
    background: #F5F3EF;
    border-radius: 10px;
    padding: 0.9rem 1rem;
    text-align: center;
    border: 1px solid #E8E4DC;
}
.stat-card .stat-num {
    font-family: 'DM Serif Display', serif;
    font-size: 1.6rem;
    color: #1F3864;
    line-height: 1;
}
.stat-card .stat-lbl {
    font-size: 0.72rem;
    color: #888;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 0.2rem;
}

/* Divider */
.divider {
    border: none;
    border-top: 1px solid #E8E4DC;
    margin: 1.5rem 0;
}

/* Override Streamlit file uploader */
[data-testid="stFileUploader"] {
    background: #FAFAF8;
    border: 2px dashed #C5D8F0;
    border-radius: 12px;
    padding: 1rem;
}
[data-testid="stFileUploader"]:hover {
    border-color: #2E5EAA;
}

/* Download button */
[data-testid="stDownloadButton"] > button {
    background: #1F3864 !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.7rem 2rem !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    width: 100% !important;
    transition: background 0.2s !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #2E5EAA !important;
}

/* Error */
.err-box {
    background: #FFF0F0;
    border-left: 4px solid #C00000;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    color: #C00000;
    font-size: 0.9rem;
}

.footer {
    text-align: center;
    color: #aaa;
    font-size: 0.78rem;
    padding: 2rem 0 1rem;
}
</style>
""", unsafe_allow_html=True)


# ── CONSTANTS ─────────────────────────────────────────────────────────────────

PRIMARY_ACCOUNTS = {'Slash', 'Divvy Credit', 'Divvy Prefund', 'Wex Credit', 'Wex Prefund'}
DARK_BLUE = '1F3864'
MED_BLUE  = '2E5EAA'
TEAL      = '1F6B75'
ALT_ROW   = 'EEF4FB'
ALT_OTHER = 'F0F7F4'
WHITE     = 'FFFFFF'


# ── REPORT LOGIC ──────────────────────────────────────────────────────────────

def relabel(account):
    a = str(account).strip(); al = a.lower()
    if al.startswith('sp') or ' sp' in al: return 'Slash'
    if 'divvy cr' in al: return 'Divvy Credit'
    if 'divvy pf' in al: return 'Divvy Prefund'
    if 'wex (prefund)' in al or 'wex prefund' in al: return 'Wex Prefund'
    if 'wex cr' in al or 'wex (credit' in al: return 'Wex Credit'
    return a

def build_filename(df):
    company_name   = str(df.iloc[1].dropna().tolist()[0]).strip()
    date_range_str = str(df.iloc[2].dropna().tolist()[0]).strip()
    m = re.search(r'(\w+)\s+[\d\-]+,?\s*(\d{4})', date_range_str)
    current_month = pd.to_datetime(f'{m.group(1)} 1 {m.group(2)}')
    next_15 = (current_month + pd.offsets.MonthBegin(1)).replace(day=15)
    day = next_15.day
    suffix = {1:'st',2:'nd',3:'rd'}.get(day%10,'th') if day not in (11,12,13) else 'th'
    date_label = next_15.strftime(f'%B {day}{suffix}')
    company_title = company_name.title().replace('Y&s','Y&S')
    report_title = f'{company_title} - Installment Projections - {date_label}'
    return report_title, date_range_str, f'{report_title}.xlsx'

def load_transactions(df):
    mask = df[8].notna() & (df[8] != 'Amount') & df[0].isna()
    tx = df[mask].copy()
    tx[8] = pd.to_numeric(tx[8], errors='coerce')
    tx = tx.dropna(subset=[8])
    tx['labeled'] = tx[7].apply(relabel)
    tx = tx[tx[2].notna() & (tx[2].astype(str).str.strip() != '')]
    tx = tx[tx['labeled'] != 'Clearing Account']
    return tx

def hfont(sz=11): return Font(name='Arial', bold=True, color=WHITE, size=sz)
def cfont(sz=10): return Font(name='Arial', size=sz)
def tborder():
    s = Side(style='thin', color='BFBFBF')
    return Border(left=s, right=s, top=s, bottom=s)

def write_title(ws, report_title, subtitle, ncols):
    span = get_column_letter(ncols)
    ws.merge_cells(f'A1:{span}1'); c = ws['A1']; c.value = report_title
    c.font = Font(name='Arial', bold=True, color=WHITE, size=13)
    c.fill = PatternFill('solid', fgColor=DARK_BLUE)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f'A2:{span}2'); c = ws['A2']; c.value = subtitle
    c.font = Font(name='Arial', italic=True, color=WHITE, size=10)
    c.fill = PatternFill('solid', fgColor=MED_BLUE)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 8

def write_col_headers(ws, row, headers, bg):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(name='Arial', bold=True, color=WHITE, size=10)
        c.fill = PatternFill('solid', fgColor=bg)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = tborder()
    ws.row_dimensions[row].height = 20

def write_section_label(ws, row, label, color, ncols):
    ws.merge_cells(f'A{row}:{get_column_letter(ncols)}{row}')
    c = ws.cell(row=row, column=1, value=label)
    c.font = Font(name='Arial', bold=True, color=WHITE, size=10)
    c.fill = PatternFill('solid', fgColor=color)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    c.border = tborder()
    ws.row_dimensions[row].height = 20

def write_subtotal(ws, row, label, bf, cf_, gt, color):
    for col,(val,fmt) in enumerate(zip([label,bf,cf_,f'=B{row}/B{gt}'],[None,'$#,##0.00','#,##0','0.0%']),1):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name='Arial', bold=True, color=WHITE, size=10)
        c.fill = PatternFill('solid', fgColor=color); c.border = tborder()
        c.alignment = Alignment(horizontal='left' if col==1 else 'center', vertical='center')
        if fmt: c.number_format = fmt
    ws.row_dimensions[row].height = 20

def write_grand_total(ws, row, bf, cf_):
    for col,(val,fmt) in enumerate(zip(['Grand Total',bf,cf_,'100.0%'],[None,'$#,##0.00','#,##0',None]),1):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name='Arial', bold=True, color=WHITE, size=11)
        c.fill = PatternFill('solid', fgColor=DARK_BLUE); c.border = tborder()
        c.alignment = Alignment(horizontal='left' if col==1 else 'center', vertical='center')
        if fmt: c.number_format = fmt
    ws.row_dimensions[row].height = 24

def build_summary_by_account(wb, tx, report_title, date_range_str):
    acct = tx.groupby('labeled')[8].agg(['sum','count']).reset_index()
    acct.columns = ['Account','Total Spent','Transactions']
    acct = acct.sort_values('Total Spent', ascending=False).reset_index(drop=True)
    prim = acct[acct['Account'].isin(PRIMARY_ACCOUNTS)].sort_values('Total Spent',ascending=False).reset_index(drop=True)
    oth  = acct[~acct['Account'].isin(PRIMARY_ACCOUNTS)].sort_values('Total Spent',ascending=False).reset_index(drop=True)

    ws = wb.active; ws.title = 'Summary By Account'
    write_title(ws, report_title, f'{date_range_str}  |  Spending By Account', 4)

    r = 4
    write_section_label(ws, r, '▸  Divvy / Slash / Wex Accounts', MED_BLUE, 4); r+=1
    write_col_headers(ws, r, ['Account','Total Spent ($)','# Transactions','% Of Total'], '4472C4'); r+=1
    p_start = r
    for i, row in prim.iterrows():
        fill = ALT_ROW if i%2==0 else WHITE
        ws.cell(row=r,column=1,value=row['Account']); ws.cell(row=r,column=2,value=row['Total Spent']); ws.cell(row=r,column=3,value=int(row['Transactions']))
        for col in range(1,5):
            c=ws.cell(row=r,column=col); c.font=cfont(); c.fill=PatternFill('solid',fgColor=fill); c.border=tborder()
            c.alignment=Alignment(horizontal='left' if col==1 else 'center',vertical='center')
        ws.cell(row=r,column=2).number_format='$#,##0.00'; ws.cell(row=r,column=3).number_format='#,##0'
        ws.row_dimensions[r].height=18; r+=1
    p_end=r-1; p_sub=r; GT=p_sub+6+len(oth)
    for i in range(len(prim)):
        rr=p_start+i; c=ws.cell(row=rr,column=4,value=f'=B{rr}/B{GT}')
        c.font=cfont(); c.fill=PatternFill('solid',fgColor=ALT_ROW if i%2==0 else WHITE)
        c.border=tborder(); c.alignment=Alignment(horizontal='center',vertical='center'); c.number_format='0.0%'
    write_subtotal(ws,p_sub,'Subtotal — Divvy / Slash / Wex',f'=SUM(B{p_start}:B{p_end})',f'=SUM(C{p_start}:C{p_end})',GT,MED_BLUE)
    ws.row_dimensions[p_sub+1].height=10
    o_sec=p_sub+2; write_section_label(ws,o_sec,'▸  Other Accounts',TEAL,4)
    write_col_headers(ws,o_sec+1,['Account','Total Spent ($)','# Transactions','% Of Total'],'2E8B8F')
    o_start=o_sec+2; r=o_start
    for i,row in oth.iterrows():
        fill=ALT_OTHER if i%2==0 else WHITE
        ws.cell(row=r,column=1,value=row['Account']); ws.cell(row=r,column=2,value=row['Total Spent'])
        ws.cell(row=r,column=3,value=int(row['Transactions'])); ws.cell(row=r,column=4,value=f'=B{r}/B{GT}')
        for col in range(1,5):
            c=ws.cell(row=r,column=col); c.font=cfont(); c.fill=PatternFill('solid',fgColor=fill); c.border=tborder()
            c.alignment=Alignment(horizontal='left' if col==1 else 'center',vertical='center')
        ws.cell(row=r,column=2).number_format='$#,##0.00'; ws.cell(row=r,column=3).number_format='#,##0'; ws.cell(row=r,column=4).number_format='0.0%'
        ws.row_dimensions[r].height=18; r+=1
    o_end=r-1; o_sub=r
    write_subtotal(ws,o_sub,'Subtotal — Other Accounts',f'=SUM(B{o_start}:B{o_end})',f'=SUM(C{o_start}:C{o_end})',GT,TEAL)
    ws.row_dimensions[o_sub+1].height=10
    write_grand_total(ws,GT,f'=B{p_sub}+B{o_sub}',f'=C{p_sub}+C{o_sub}')
    for col,w in zip('ABCD',[26,18,16,14]): ws.column_dimensions[col].width=w

def build_summary_by_team(wb, tx, report_title, date_range_str):
    team = tx.groupby(4)[8].agg(['sum','count']).reset_index()
    team.columns=['Team','Total Spent','Transactions']
    team=team.sort_values('Total Spent',ascending=False).reset_index(drop=True)
    ws=wb.create_sheet('Summary By Team')
    write_title(ws,report_title,f'{date_range_str}  |  Spending By Team  |  {len(team)} Teams',4)
    write_col_headers(ws,4,['Team','Total Spent ($)','# Transactions','% Of Total'],MED_BLUE)
    ws.row_dimensions[4].height=22
    GT=5+len(team)
    for i,row in team.iterrows():
        r=i+5; fill=ALT_ROW if i%2==0 else WHITE
        ws.cell(row=r,column=1,value=row['Team']); ws.cell(row=r,column=2,value=row['Total Spent'])
        ws.cell(row=r,column=3,value=int(row['Transactions'])); ws.cell(row=r,column=4,value=f'=B{r}/B{GT}')
        for col in range(1,5):
            c=ws.cell(row=r,column=col); c.font=cfont(); c.fill=PatternFill('solid',fgColor=fill); c.border=tborder()
            c.alignment=Alignment(horizontal='left' if col==1 else 'center',vertical='center')
        ws.cell(row=r,column=2).number_format='$#,##0.00'; ws.cell(row=r,column=3).number_format='#,##0'
        ws.cell(row=r,column=4).number_format='0.0%'; ws.row_dimensions[r].height=18
    write_grand_total(ws,GT,f'=SUM(B5:B{GT-1})',f'=SUM(C5:C{GT-1})')
    for col,w in zip('ABCD',[32,18,16,14]): ws.column_dimensions[col].width=w
    ws.freeze_panes='A5'

def build_transactions(wb, tx, report_title, date_range_str):
    tdf=tx[[1,2,4,5,'labeled',8]].copy()
    tdf.columns=['Transaction Date','Transaction Type','Name','Description','Account','Amount']
    tdf['Transaction Date']=pd.to_datetime(tdf['Transaction Date'],errors='coerce')
    tdf=tdf.sort_values('Amount',ascending=False).reset_index(drop=True)
    ws=wb.create_sheet('Transactions')
    write_title(ws,report_title,f'{date_range_str}  |  {len(tdf):,} Transactions',6)
    for col,h in enumerate(['Date','Type','Name','Description','Account','Amount ($)'],1):
        c=ws.cell(row=4,column=col,value=h); c.font=hfont()
        c.fill=PatternFill('solid',fgColor=MED_BLUE); c.alignment=Alignment(horizontal='center',vertical='center'); c.border=tborder()
    ws.row_dimensions[4].height=22
    for i,row in tdf.iterrows():
        r=i+5; fill=ALT_ROW if i%2==0 else WHITE
        vals=[row['Transaction Date'].strftime('%m/%d/%Y') if pd.notna(row['Transaction Date']) else '',
              row['Transaction Type'],row['Name'],row['Description'],row['Account'],row['Amount']]
        for col,val in enumerate(vals,1):
            c=ws.cell(row=r,column=col,value=val); c.font=cfont(sz=9)
            c.fill=PatternFill('solid',fgColor=fill); c.border=tborder(); c.alignment=Alignment(horizontal='left',vertical='center')
        ws.cell(row=r,column=6).alignment=Alignment(horizontal='right',vertical='center')
        ws.cell(row=r,column=6).number_format='$#,##0.00'; ws.row_dimensions[r].height=16
    for i,w in enumerate([13,16,22,52,18,14],1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes='A5'

def generate_report_bytes(file_bytes):
    """Run full report pipeline, return (filename, xlsx_bytes, stats_dict)."""
    df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    report_title, date_range_str, filename = build_filename(df)
    tx = load_transactions(df)

    stats = {
        'transactions': len(tx),
        'accounts': tx['labeled'].nunique(),
        'teams': tx[4].nunique(),
        'total': tx[8].sum(),
        'date_range': date_range_str,
        'report_title': report_title,
    }

    wb = Workbook()
    build_summary_by_account(wb, tx, report_title, date_range_str)
    build_summary_by_team(wb, tx, report_title, date_range_str)
    build_transactions(wb, tx, report_title, date_range_str)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return filename, buf.read(), stats


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
    <h1>🎟️ Installment<br><em>Projections</em></h1>
    <p>Drop in a General Ledger export and get your formatted report instantly.</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="upload-card">', unsafe_allow_html=True)
st.markdown('<div class="step-label">Step 1 — Upload your file</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    label="General Ledger (.xlsx)",
    type=["xlsx"],
    label_visibility="collapsed",
)

if uploaded:
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    st.markdown('<div class="step-label">Step 2 — Your report</div>', unsafe_allow_html=True)

    with st.spinner("Building your report…"):
        try:
            filename, xlsx_bytes, stats = generate_report_bytes(uploaded.read())

            st.markdown(f"""
            <div class="success-box">
                <p class="file-title">📄 {filename}</p>
                <p class="file-meta">{stats['date_range']}</p>
            </div>
            """, unsafe_allow_html=True)

            st.markdown(f"""
            <div class="stats-row">
                <div class="stat-card">
                    <div class="stat-num">{stats['transactions']:,}</div>
                    <div class="stat-lbl">Transactions</div>
                </div>
                <div class="stat-card">
                    <div class="stat-num">{stats['accounts']}</div>
                    <div class="stat-lbl">Accounts</div>
                </div>
                <div class="stat-card">
                    <div class="stat-num">{stats['teams']}</div>
                    <div class="stat-lbl">Teams</div>
                </div>
                <div class="stat-card">
                    <div class="stat-num">${stats['total']:,.0f}</div>
                    <div class="stat-lbl">Total Spent</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.download_button(
                label="⬇️  Download Report",
                data=xlsx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.markdown(f'<div class="err-box">⚠️ Something went wrong: {e}</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
<div class="footer">
    3 tabs · Summary By Account · Summary By Team · Transactions
</div>
""", unsafe_allow_html=True)
