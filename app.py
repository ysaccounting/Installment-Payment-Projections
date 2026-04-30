"""
Y&S Tickets — Installment Projections
Streamlit web app: drop in a General Ledger, download the report.

Supports both General Ledger export formats automatically.
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

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #F5F3EF; }
#MainMenu, footer, header { visibility: hidden; }

.hero { text-align: center; padding: 3rem 0 2rem; }
.hero h1 {
    font-family: 'DM Serif Display', serif;
    font-size: 2.8rem; color: #1F3864; margin: 0;
    line-height: 1.1; letter-spacing: -0.5px;
}
.hero h1 em { font-style: italic; color: #2E5EAA; }
.hero p { font-size: 1.05rem; color: #666; margin-top: 0.6rem; font-weight: 300; }

.upload-card {
    background: white; border-radius: 16px; padding: 2.5rem 2rem;
    box-shadow: 0 2px 20px rgba(31,56,100,0.08);
    border: 1.5px solid #E8E4DC; margin: 1.5rem 0;
}
.step-label {
    font-size: 0.7rem; font-weight: 600; letter-spacing: 2px;
    text-transform: uppercase; color: #2E5EAA; margin-bottom: 0.5rem;
}
.success-box {
    background: #EEF4FB; border-left: 4px solid #2E5EAA;
    border-radius: 8px; padding: 1rem 1.2rem; margin: 1rem 0;
}
.success-box .file-title {
    font-family: 'DM Serif Display', serif;
    font-size: 1.1rem; color: #1F3864; margin: 0;
}
.success-box .file-meta { font-size: 0.85rem; color: #555; margin: 0.2rem 0 0; }

.stats-row { display: flex; gap: 1rem; margin: 1rem 0; }
.stat-card {
    flex: 1; background: #F5F3EF; border-radius: 10px;
    padding: 0.9rem 1rem; text-align: center; border: 1px solid #E8E4DC;
}
.stat-card .stat-num {
    font-family: 'DM Serif Display', serif;
    font-size: 1.6rem; color: #1F3864; line-height: 1;
}
.stat-card .stat-lbl {
    font-size: 0.72rem; color: #888;
    text-transform: uppercase; letter-spacing: 1px; margin-top: 0.2rem;
}
.divider { border: none; border-top: 1px solid #E8E4DC; margin: 1.5rem 0; }

[data-testid="stFileUploader"] {
    background: #FAFAF8; border: 2px dashed #C5D8F0; border-radius: 12px; padding: 1rem;
}
[data-testid="stFileUploader"]:hover { border-color: #2E5EAA; }

[data-testid="stDownloadButton"] > button {
    background: #1F3864 !important; color: white !important; border: none !important;
    border-radius: 10px !important; padding: 0.7rem 2rem !important;
    font-family: 'DM Sans', sans-serif !important; font-weight: 600 !important;
    font-size: 0.95rem !important; width: 100% !important; transition: background 0.2s !important;
}
[data-testid="stDownloadButton"] > button:hover { background: #2E5EAA !important; }

.err-box {
    background: #FFF0F0; border-left: 4px solid #C00000;
    border-radius: 8px; padding: 1rem 1.2rem; color: #C00000; font-size: 0.9rem;
}
.footer { text-align: center; color: #aaa; font-size: 0.78rem; padding: 2rem 0 1rem; }
</style>
""", unsafe_allow_html=True)


# ── CONSTANTS ─────────────────────────────────────────────────────────────────

PRIMARY_ACCOUNTS = {'Slash', 'Divvy Credit', 'Divvy Prefund', 'Wex Credit', 'Wex Prefund'}
EXCLUDE_ACCOUNTS = {'Clearing Account', 'Accounts Payable'}
DARK_BLUE='1F3864'; MED_BLUE='2E5EAA'; TEAL='1F6B75'
ALT_ROW='EEF4FB'; ALT_OTHER='F0F7F4'; WHITE='FFFFFF'


# ── VALIDATION ───────────────────────────────────────────────────────────────

def validate_gl(df):
    """Returns (is_valid, error_message). Catches accidentally uploaded output reports."""
    ncols = df.shape[1]
    if ncols < 9:
        return False, (
            f"This looks like a previously generated report, not a General Ledger export "
            f"(only {ncols} columns found). Please upload the raw QuickBooks GL export (.xlsx)."
        )
    for col in [8, 9]:
        if col < ncols and pd.to_numeric(df[col], errors='coerce').notna().sum() > 5:
            return True, None
    return False, (
        "This doesn't look like a General Ledger export. "
        "Please upload the raw QuickBooks GL export (.xlsx), not a previously generated report."
    )


# ── FORMAT DETECTION ─────────────────────────────────────────────────────────

def detect_format(df):
    return 'new' if str(df.iloc[1, 0]).strip() == 'Transaction Report' else 'old'


# ── METADATA / FILENAME ───────────────────────────────────────────────────────

def smart_title(s):
    ALLCAPS = {'llc','tl','ys','yss','mls','nba','nhl','mlb','nfl','tc','dep','lp','inc','ltd'}
    result = []
    for w in s.split():
        if '&' in w: result.append(w.upper() if len(w) <= 3 else w.title())
        elif w.lower() in ALLCAPS: result.append(w.upper())
        else: result.append(w.capitalize())
    return ' '.join(result)

def build_filename(df, fmt):
    if fmt == 'old':
        company_name   = str(df.iloc[1].dropna().tolist()[0]).strip()
        date_range_str = str(df.iloc[2].dropna().tolist()[0]).strip()
    else:
        company_name   = str(df.iloc[0].dropna().tolist()[0]).strip()
        date_range_str = str(df.iloc[2].dropna().tolist()[0]).strip()
    m = re.search(r'(\w+)\s+[\d\-]+,?\s*(\d{4})', date_range_str) or \
        re.search(r'(\w+)\s+(\d{4})', date_range_str)
    current_month = pd.to_datetime(f'{m.group(1)} 1 {m.group(2)}')
    next_15 = (current_month + pd.offsets.MonthBegin(1)).replace(day=15)
    day = next_15.day
    suffix = {1:'st',2:'nd',3:'rd'}.get(day%10,'th') if day not in (11,12,13) else 'th'
    date_label = next_15.strftime(f'%B {day}{suffix}')
    company_title = smart_title(company_name)
    report_title = f'{company_title} - Installment Projections - {date_label}'
    return report_title, date_range_str, f'{report_title}.xlsx'


# ── ACCOUNT RELABELING ───────────────────────────────────────────────────────

def relabel(account):
    a=str(account).strip(); al=a.lower()
    if 'slash plat' in al or al.startswith('sp') or ' sp' in al: return 'Slash'
    if 'divvy cr' in al or al == 'divvy (credit)': return 'Divvy Credit'
    if 'divvy pf' in al or al == 'divvy (prefund)': return 'Divvy Prefund'
    if 'wex (prefund)' in al or 'wex prefund' in al: return 'Wex Prefund'
    if 'wex cr' in al or 'wex (credit' in al: return 'Wex Credit'
    return a


# ── DATA LOADING ──────────────────────────────────────────────────────────────

def load_transactions(df, fmt):
    if fmt == 'old':
        amt_col=8; acct_col=7; type_col=2; name_col=4; desc_col=5; date_col=1
        mask = df[amt_col].notna() & (df[amt_col] != 'Amount') & df[0].isna()
    else:
        amt_col=9; acct_col=8; type_col=2; name_col=5; desc_col=6; date_col=1
        mask = (df[amt_col].notna() & df[0].isna() & df[date_col].notna() &
                ~df[date_col].astype(str).str.contains('Beginning Balance|Date', case=False, na=False))
    tx = df[mask].copy()
    tx[amt_col] = pd.to_numeric(tx[amt_col], errors='coerce')
    tx = tx.dropna(subset=[amt_col])
    tx['labeled'] = tx[acct_col].apply(relabel)
    tx = tx[tx[type_col].notna() & (tx[type_col].astype(str).str.strip() != '')]
    tx = tx[~tx['labeled'].isin(EXCLUDE_ACCOUNTS)]
    tx = tx[~tx[acct_col].isin(EXCLUDE_ACCOUNTS)]
    result = tx[[date_col, type_col, name_col, desc_col, 'labeled', amt_col]].copy()
    result.columns = ['Date','Type','Name','Description','Account','Amount']
    result['Date'] = pd.to_datetime(result['Date'], errors='coerce')
    return result


# ── STYLE HELPERS ─────────────────────────────────────────────────────────────

def hfont(sz=11): return Font(name='Arial', bold=True, color=WHITE, size=sz)
def cfont(sz=10): return Font(name='Arial', size=sz)
def tborder():
    s=Side(style='thin',color='BFBFBF'); return Border(left=s,right=s,top=s,bottom=s)

def write_title(ws, rt, sub, ncols):
    sp=get_column_letter(ncols)
    ws.merge_cells(f'A1:{sp}1'); c=ws['A1']; c.value=rt
    c.font=Font(name='Arial',bold=True,color=WHITE,size=13); c.fill=PatternFill('solid',fgColor=DARK_BLUE)
    c.alignment=Alignment(horizontal='center',vertical='center'); ws.row_dimensions[1].height=30
    ws.merge_cells(f'A2:{sp}2'); c=ws['A2']; c.value=sub
    c.font=Font(name='Arial',italic=True,color=WHITE,size=10); c.fill=PatternFill('solid',fgColor=MED_BLUE)
    c.alignment=Alignment(horizontal='center',vertical='center'); ws.row_dimensions[2].height=18; ws.row_dimensions[3].height=8

def write_col_headers(ws, row, hdrs, bg):
    for col,h in enumerate(hdrs,1):
        c=ws.cell(row=row,column=col,value=h); c.font=Font(name='Arial',bold=True,color=WHITE,size=10)
        c.fill=PatternFill('solid',fgColor=bg); c.alignment=Alignment(horizontal='center',vertical='center'); c.border=tborder()
    ws.row_dimensions[row].height=20

def write_sec_lbl(ws, row, lbl, color, ncols):
    ws.merge_cells(f'A{row}:{get_column_letter(ncols)}{row}'); c=ws.cell(row=row,column=1,value=lbl)
    c.font=Font(name='Arial',bold=True,color=WHITE,size=10); c.fill=PatternFill('solid',fgColor=color)
    c.alignment=Alignment(horizontal='left',vertical='center',indent=1); c.border=tborder(); ws.row_dimensions[row].height=20

def write_subtotal(ws, row, lbl, bf, cf_, gt, color):
    for col,(val,fmt) in enumerate(zip([lbl,bf,cf_,f'=B{row}/B{gt}'],[None,'$#,##0.00','#,##0','0.0%']),1):
        c=ws.cell(row=row,column=col,value=val); c.font=Font(name='Arial',bold=True,color=WHITE,size=10)
        c.fill=PatternFill('solid',fgColor=color); c.border=tborder()
        c.alignment=Alignment(horizontal='left' if col==1 else 'center',vertical='center')
        if fmt: c.number_format=fmt
    ws.row_dimensions[row].height=20

def write_grand_total(ws, row, bf, cf_):
    for col,(val,fmt) in enumerate(zip(['Grand Total',bf,cf_,'100.0%'],[None,'$#,##0.00','#,##0',None]),1):
        c=ws.cell(row=row,column=col,value=val); c.font=Font(name='Arial',bold=True,color=WHITE,size=11)
        c.fill=PatternFill('solid',fgColor=DARK_BLUE); c.border=tborder()
        c.alignment=Alignment(horizontal='left' if col==1 else 'center',vertical='center')
        if fmt: c.number_format=fmt
    ws.row_dimensions[row].height=24


# ── SHEET BUILDERS ────────────────────────────────────────────────────────────

def build_summary_by_account(wb, tx, rt, drs):
    acct=tx.groupby('Account')['Amount'].agg(['sum','count']).reset_index()
    acct.columns=['Account','Total Spent','Transactions']
    acct=acct.sort_values('Total Spent',ascending=False).reset_index(drop=True)
    prim=acct[acct['Account'].isin(PRIMARY_ACCOUNTS)].sort_values('Total Spent',ascending=False).reset_index(drop=True)
    oth=acct[~acct['Account'].isin(PRIMARY_ACCOUNTS)].sort_values('Total Spent',ascending=False).reset_index(drop=True)
    ws=wb.active; ws.title='Summary By Account'
    write_title(ws,rt,f'{drs}  |  Spending By Account',4)
    r=4; write_sec_lbl(ws,r,'▸  Divvy / Slash / Wex Accounts',MED_BLUE,4); r+=1
    write_col_headers(ws,r,['Account','Total Spent ($)','# Transactions','% Of Total'],'4472C4'); r+=1
    p_start=r
    for i,row in prim.iterrows():
        fill=ALT_ROW if i%2==0 else WHITE
        ws.cell(row=r,column=1,value=row['Account']); ws.cell(row=r,column=2,value=row['Total Spent']); ws.cell(row=r,column=3,value=int(row['Transactions']))
        for col in range(1,5):
            c=ws.cell(row=r,column=col); c.font=cfont(); c.fill=PatternFill('solid',fgColor=fill); c.border=tborder()
            c.alignment=Alignment(horizontal='left' if col==1 else 'center',vertical='center')
        ws.cell(row=r,column=2).number_format='$#,##0.00'; ws.cell(row=r,column=3).number_format='#,##0'; ws.row_dimensions[r].height=18; r+=1
    p_end=r-1; p_sub=r; GT=p_sub+6+len(oth)
    for i in range(len(prim)):
        rr=p_start+i; c=ws.cell(row=rr,column=4,value=f'=B{rr}/B{GT}')
        c.font=cfont(); c.fill=PatternFill('solid',fgColor=ALT_ROW if i%2==0 else WHITE)
        c.border=tborder(); c.alignment=Alignment(horizontal='center',vertical='center'); c.number_format='0.0%'
    write_subtotal(ws,p_sub,'Subtotal — Divvy / Slash / Wex',f'=SUM(B{p_start}:B{p_end})',f'=SUM(C{p_start}:C{p_end})',GT,MED_BLUE)
    ws.row_dimensions[p_sub+1].height=10
    o_sec=p_sub+2; write_sec_lbl(ws,o_sec,'▸  Other Accounts',TEAL,4)
    write_col_headers(ws,o_sec+1,['Account','Total Spent ($)','# Transactions','% Of Total'],'2E8B8F')
    o_start=o_sec+2; r=o_start
    for i,row in oth.iterrows():
        fill=ALT_OTHER if i%2==0 else WHITE
        ws.cell(row=r,column=1,value=row['Account']); ws.cell(row=r,column=2,value=row['Total Spent'])
        ws.cell(row=r,column=3,value=int(row['Transactions'])); ws.cell(row=r,column=4,value=f'=B{r}/B{GT}')
        for col in range(1,5):
            c=ws.cell(row=r,column=col); c.font=cfont(); c.fill=PatternFill('solid',fgColor=fill); c.border=tborder()
            c.alignment=Alignment(horizontal='left' if col==1 else 'center',vertical='center')
        ws.cell(row=r,column=2).number_format='$#,##0.00'; ws.cell(row=r,column=3).number_format='#,##0'; ws.cell(row=r,column=4).number_format='0.0%'; ws.row_dimensions[r].height=18; r+=1
    o_end=r-1; o_sub=r
    write_subtotal(ws,o_sub,'Subtotal — Other Accounts',f'=SUM(B{o_start}:B{o_end})',f'=SUM(C{o_start}:C{o_end})',GT,TEAL)
    ws.row_dimensions[o_sub+1].height=10; write_grand_total(ws,GT,f'=B{p_sub}+B{o_sub}',f'=C{p_sub}+C{o_sub}')
    for col,w in zip('ABCD',[26,18,16,14]): ws.column_dimensions[col].width=w

def build_summary_by_team(wb, tx, rt, drs):
    team=tx.groupby('Name')['Amount'].agg(['sum','count']).reset_index()
    team.columns=['Team','Total Spent','Transactions']
    team=team.sort_values('Total Spent',ascending=False).reset_index(drop=True)
    ws=wb.create_sheet('Summary By Team')
    write_title(ws,rt,f'{drs}  |  Spending By Team  |  {len(team)} Teams',4)
    write_col_headers(ws,4,['Team','Total Spent ($)','# Transactions','% Of Total'],MED_BLUE); ws.row_dimensions[4].height=22
    GT=5+len(team)
    for i,row in team.iterrows():
        r=i+5; fill=ALT_ROW if i%2==0 else WHITE
        ws.cell(row=r,column=1,value=row['Team']); ws.cell(row=r,column=2,value=row['Total Spent'])
        ws.cell(row=r,column=3,value=int(row['Transactions'])); ws.cell(row=r,column=4,value=f'=B{r}/B{GT}')
        for col in range(1,5):
            c=ws.cell(row=r,column=col); c.font=cfont(); c.fill=PatternFill('solid',fgColor=fill); c.border=tborder()
            c.alignment=Alignment(horizontal='left' if col==1 else 'center',vertical='center')
        ws.cell(row=r,column=2).number_format='$#,##0.00'; ws.cell(row=r,column=3).number_format='#,##0'; ws.cell(row=r,column=4).number_format='0.0%'; ws.row_dimensions[r].height=18
    write_grand_total(ws,GT,f'=SUM(B5:B{GT-1})',f'=SUM(C5:C{GT-1})')
    for col,w in zip('ABCD',[32,18,16,14]): ws.column_dimensions[col].width=w
    ws.freeze_panes='A5'

def build_transactions(wb, tx, rt, drs):
    tdf=tx.sort_values('Amount',ascending=False).reset_index(drop=True)
    ws=wb.create_sheet('Transactions')
    write_title(ws,rt,f'{drs}  |  {len(tdf):,} Transactions',6)
    for col,h in enumerate(['Date','Type','Name','Description','Account','Amount ($)'],1):
        c=ws.cell(row=4,column=col,value=h); c.font=hfont()
        c.fill=PatternFill('solid',fgColor=MED_BLUE); c.alignment=Alignment(horizontal='center',vertical='center'); c.border=tborder()
    ws.row_dimensions[4].height=22
    for i,row in tdf.iterrows():
        r=i+5; fill=ALT_ROW if i%2==0 else WHITE
        vals=[row['Date'].strftime('%m/%d/%Y') if pd.notna(row['Date']) else '',row['Type'],row['Name'],row['Description'],row['Account'],row['Amount']]
        for col,val in enumerate(vals,1):
            c=ws.cell(row=r,column=col,value=val); c.font=cfont(sz=9)
            c.fill=PatternFill('solid',fgColor=fill); c.border=tborder(); c.alignment=Alignment(horizontal='left',vertical='center')
        ws.cell(row=r,column=6).alignment=Alignment(horizontal='right',vertical='center')
        ws.cell(row=r,column=6).number_format='$#,##0.00'; ws.row_dimensions[r].height=16
    for i,w in enumerate([13,16,22,52,18,14],1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes='A5'


# ── PIPELINE ──────────────────────────────────────────────────────────────────

def generate_report_bytes(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    valid, err = validate_gl(df)
    if not valid:
        raise ValueError(err)
    fmt = detect_format(df)
    report_title, date_range_str, filename = build_filename(df, fmt)
    tx = load_transactions(df, fmt)
    stats = {
        'transactions': len(tx), 'accounts': tx['Account'].nunique(),
        'teams': tx['Name'].nunique(), 'total': float(tx['Amount'].sum()),
        'date_range': date_range_str, 'report_title': report_title,
    }
    wb = Workbook()
    build_summary_by_account(wb, tx, report_title, date_range_str)
    build_summary_by_team(wb, tx, report_title, date_range_str)
    build_transactions(wb, tx, report_title, date_range_str)
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
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

uploaded_files = st.file_uploader(
    label="General Ledger (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    n = len(uploaded_files)
    st.markdown(f'<div class="step-label">Step 2 — Your {"report" if n == 1 else f"{n} reports"}</div>', unsafe_allow_html=True)

    for uploaded in uploaded_files:
        with st.spinner(f"Building report for {uploaded.name}…"):
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
                    label=f"⬇️  Download — {filename}",
                    data=xlsx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=filename,
                )

            except Exception as e:
                st.markdown(f'<div class="err-box">⚠️ <strong>{uploaded.name}</strong>: {e}</div>', unsafe_allow_html=True)

            st.markdown('<hr class="divider">', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
<div class="footer">
    3 tabs · Summary By Account · Summary By Team · Transactions · Two export formats supported
</div>
""", unsafe_allow_html=True)
