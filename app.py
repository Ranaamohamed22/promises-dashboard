import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, datetime, time
import io
import base64
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# إعدادات الصفحة الأساسية
# ==========================================
st.set_page_config(page_title="SNB Promises Dashboard", layout="wide", initial_sidebar_state="collapsed")

if 'current_page' not in st.session_state:
    st.session_state.current_page = 'Overview'
if 'theme_choice' not in st.session_state:
    st.session_state.theme_choice = 'ثيم إجادة (Navy)'

# ==========================================
# إعدادات الثيم والألوان بالترتيب المطلوب
# ==========================================
ejada_navy_dark = "#1F4373"
ejada_navy_light = "#304D73"
ejada_grey_blue = "#586F8C"
ejada_border_light = "#99AABF"
white_bg = "#FFFFFF"  
snb_green = "#2CA549"  

if st.session_state.theme_choice == "ثيم إجادة (Navy)":
    bg_color = ejada_navy_dark 
    card_color = white_bg 
    border_color = ejada_border_light
    primary_color = ejada_navy_light 
    primary_text = "#FFFFFF"
    text_color = ejada_navy_dark
    sub_text_color = ejada_grey_blue
    gauge_bg = "#E2E8F0"
    nav_btn_bg = "#FFFFFF"
    nav_btn_text = ejada_navy_dark
    nav_btn_border = ejada_border_light
    kpi_title_color = "#586F8C"
    header_logo_color = "#FFFFFF"
else:
    bg_color = "#13131A"
    card_color = "#1C1C24"
    border_color = "#2D2D3B"
    primary_color = ejada_navy_light
    primary_text = "#FFFFFF"
    text_color = "#FFFFFF" 
    sub_text_color = "#A0AEC0"
    gauge_bg = "#2D2D3B"
    nav_btn_bg = "#1A1A1D"
    nav_btn_text = "#E0E0E0"
    nav_btn_border = "#333333"
    kpi_title_color = "#A0AEC0"
    header_logo_color = "#FFFFFF"

# ==========================================
# تصميم الـ UI الديناميكي (CSS)
# ==========================================
st.markdown(f"""
<style>
    [data-testid="collapsedControl"] {{display: none !important;}}
    [data-testid="stSidebar"] {{display: none !important;}}
    
    .stApp {{ background-color: {bg_color} !important; }}

    .block-container {{ padding-top: 3.8rem !important; padding-left: 1rem !important; padding-right: 1rem !important; max-width: 100% !important; }}
    
    div[data-testid="stButton"] > button[kind="secondary"] {{
        background-color: {nav_btn_bg} !important;
        color: {nav_btn_text} !important;
        border: 1px solid {nav_btn_border} !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }}
    div[data-testid="stButton"] > button[kind="secondary"] p {{ color: {nav_btn_text} !important; font-weight: 800 !important; }}
    div[data-testid="stButton"] > button[kind="secondary"]:hover, 
    div[data-testid="stButton"] > button[kind="secondary"]:focus {{ border-color: {ejada_grey_blue} !important; color: {ejada_navy_light} !important; }}
    
    div[data-testid="stButton"] > button[kind="primary"] {{
        background-color: {snb_green} !important; 
        color: #FFFFFF !important; 
        border: none !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.3) !important;
    }}
    div[data-testid="stButton"] > button[kind="primary"] p {{ color: #FFFFFF !important; font-weight: 800 !important; }}
    div[data-testid="stButton"] > button[kind="primary"]:focus {{ box-shadow: none !important; }}

    div.st-key-theme_toggle div[data-testid="stButton"] > button {{
        background: linear-gradient(90deg, {ejada_navy_dark} 50%, {ejada_grey_blue} 50%) !important;
        border: 2px solid #FFFFFF !important;
        border-radius: 50% !important; 
        width: 38px !important;
        height: 38px !important;
        min-width: 38px !important;
        min-height: 38px !important;
        padding: 0 !important;
        margin-top: 1px !important; 
        color: transparent !important; 
        box-shadow: 0 2px 6px rgba(0,0,0,0.4) !important;
    }}
    div.st-key-theme_toggle div[data-testid="stButton"] > button:hover {{ transform: scale(1.1) !important; }}

    .stTextInput input {{ background-color: {nav_btn_bg} !important; color: {text_color} !important; border: 1px solid {border_color} !important; border-radius: 8px !important; }}
    .stTextInput input::placeholder {{ color: {sub_text_color} !important; }}
    [data-testid="stFileUploader"] > section {{ background-color: {nav_btn_bg} !important; border: 1px dashed {border_color} !important; border-radius: 8px !important; }}
    [data-testid="stFileUploader"] section * {{ color: {text_color} !important; }}

    .kpi-card {{ background-color: {card_color}; border-radius: 12px; padding: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.2); border: 1px solid {border_color}; border-left-width: 6px; border-left-style: solid; text-align: left; direction: ltr; position: relative; }}
    .kpi-title {{ color: {kpi_title_color}; font-size: 13px; text-transform: uppercase; font-weight: 800; margin-bottom: 10px; }}
    .kpi-value {{ color: {text_color}; font-size: 30px; font-weight: 900; margin-bottom: 5px; }}
    .kpi-sub {{ color: {sub_text_color}; font-size: 12px; font-weight: 600; }}
    .kpi-icon {{ position: absolute; top: 20px; right: 20px; font-size: 24px; opacity: 0.9; }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 0. الصف العلوي: القمر + Overview
# ==========================================
top_cols = st.columns([0.4, 1.3, 8], gap="small")

with top_cols[0]:
    if st.button(" ", key="theme_toggle"): 
        if st.session_state.theme_choice == 'التصميم الأصلي (Dark)':
            st.session_state.theme_choice = 'ثيم إجادة (Navy)'
        else:
            st.session_state.theme_choice = 'التصميم الأصلي (Dark)'
        st.rerun()

with top_cols[1]:
    if st.button("📊 Overview", type="primary" if st.session_state.current_page == 'Overview' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Overview'
        st.rerun()

# ==========================================
# 1. الهيدر الرئيسي (إجادة + SNB)
# ==========================================
st.markdown(f"""
<div style="position: relative; width: 100%; min-height: 120px; margin-top: -60px; margin-bottom: 25px; pointer-events: none;">
    <div style="text-align: center; line-height: 1.1;">
        <div style="color: {header_logo_color}; font-size: 45px; font-weight: 900; font-family: Tahoma, sans-serif;">إجـــادة</div>
        <div style="color: {header_logo_color}; font-size: 20px; letter-spacing: 10px; font-weight: 900; margin-top: -8px; margin-right:-10px;">EJADA</div>
        <div style="color: {header_logo_color}; font-size: 18px; font-weight: bold; margin-top: 5px;">لتحصيل الديون</div>
    </div>
    <div style="position: absolute; top: 0; right: -15px; text-align: right; line-height: 1.2; direction: ltr;">
        <div style="white-space: nowrap;">
            <span style="color: {snb_green}; font-weight: 900; font-size: 26px; letter-spacing: 1px; vertical-align: middle;">SNB</span>
            <span style="color: {snb_green}; font-size: 22px; font-weight: bold; font-family: Tahoma, sans-serif; margin-left: 8px; vertical-align: middle;">البنك الأهلي السعودي</span>
        </div>
        <div style="color: {snb_green}; font-size: 17px; font-weight: 900; text-align: center; letter-spacing: 1.5px; margin-top: 2px;">
            (30-60-NPL)
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 2. شريط التنقل
# ==========================================
nav_cols = st.columns([1.5, 1.3, 1.3, 5], gap="small") 

with nav_cols[0]:
    if st.button("🤝 الوعود القائمة والمكسورة", type="primary" if st.session_state.current_page == 'Promises' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Promises'
        st.rerun()

with nav_cols[1]:
    if st.button("🚨 تقرير الإهمال", type="primary" if st.session_state.current_page == 'Neglect' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Neglect'
        st.rerun()

with nav_cols[2]:
    if st.button("📈 النشاط", type="primary" if st.session_state.current_page == 'Activity' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Activity'
        st.rerun()

with nav_cols[3]:
    st.markdown("")

st.markdown(f"<hr style='margin: 10px 0 15px 0; border-color: {border_color}; clear: both;'>", unsafe_allow_html=True)

# ==========================================
# دوال تقرير الوعود والإكسيل
# ==========================================
PIVOT_TEAM = 'Sales Team'
PIVOT_COLL = 'Salesperson'
PIVOT_NOTE = 'ملاحظات-التمويل'
PIVOT_ACC = 'Account Number'
PIVOT_AMT = 'Net Amount'
SUM_FMT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'

def build_pivot(df, order='desc'):
    need = [PIVOT_TEAM, PIVOT_COLL, PIVOT_NOTE, PIVOT_ACC, PIVOT_AMT]
    if df is None or df.empty or any(c not in df.columns for c in need):
        return {'notes': ['NPL'], 'teams': [], 'grand': {'by_note': {'NPL': (0, 0.0)}, 'cnt': 0, 'amt': 0.0}}

    d = df[need].copy()
    d['_c'] = d[PIVOT_ACC].notna().astype(int)
    d[PIVOT_AMT] = pd.to_numeric(d[PIVOT_AMT], errors='coerce').fillna(0.0)
    d[PIVOT_TEAM] = d[PIVOT_TEAM].fillna('(فارغ)').astype(str)
    d[PIVOT_COLL] = d[PIVOT_COLL].fillna('(فارغ)').astype(str)
    d[PIVOT_NOTE] = d[PIVOT_NOTE].fillna('(فارغ)').astype(str).str.strip()

    def _ordered(values):
        vals = list(pd.unique(values))
        return sorted(vals, reverse=True) if order == 'desc' else vals

    notes = sorted(d[PIVOT_NOTE].unique(), reverse=True)
    g = d.groupby([PIVOT_TEAM, PIVOT_COLL, PIVOT_NOTE]).agg(cnt=('_c', 'sum'), amt=(PIVOT_AMT, 'sum'))

    teams = []
    grand_by_note = {n: [0, 0.0] for n in notes}
    for team in _ordered(d[PIVOT_TEAM]):
        rows = []
        for coll in _ordered(d.loc[d[PIVOT_TEAM] == team, PIVOT_COLL]):
            by_note, t_cnt, t_amt = {}, 0, 0.0
            for n in notes:
                if (team, coll, n) in g.index:
                    c, a = int(g.loc[(team, coll, n), 'cnt']), float(g.loc[(team, coll, n), 'amt'])
                else:
                    c, a = 0, 0.0
                by_note[n] = (c, a)
                t_cnt += c
                t_amt += a
                grand_by_note[n][0] += c
                grand_by_note[n][1] += a
            rows.append({'coll': coll, 'by_note': by_note, 'cnt': t_cnt, 'amt': t_amt})
        teams.append({'team': team, 'rows': rows})

    grand = {
        'by_note': {n: (v[0], v[1]) for n, v in grand_by_note.items()},
        'cnt': sum(v[0] for v in grand_by_note.values()),
        'amt': sum(v[1] for v in grand_by_note.values()),
    }
    return {'notes': notes, 'teams': teams, 'grand': grand}

def pivot_to_frame(p):
    cols = ['اسم المشرف', 'اسم المحصل']
    for n in p['notes']:
        cols += [f'{n} | عدد الحسابات', f'{n} | متبقي المديونية']
    cols += ['Total عدد الحسابات', 'Total متبقي المديونية']
    out = []
    for t in p['teams']:
        for i, r in enumerate(t['rows']):
            row = [t['team'] if i == 0 else '', r['coll']]
            for n in p['notes']:
                row += [r['by_note'][n][0], round(r['by_note'][n][1], 2)]
            row += [r['cnt'], round(r['amt'], 2)]
            out.append(row)
    g = p['grand']
    row = ['الاجمالي', '']
    for n in p['notes']:
        row += [g['by_note'][n][0], round(g['by_note'][n][1], 2)]
    row += [g['cnt'], round(g['amt'], 2)]
    out.append(row)
    return pd.DataFrame(out, columns=cols)

def _prepare_for_output(df):
    d = df.copy()
    for c in ['Amount', 'Net Amount', 'Payment']:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors='coerce')
    for c in ['Debit Date', 'Assign Date', 'Follow up Due Date', 'Follow up Last Date', 'Expected Closing']:
        if c in d.columns:
            d[c] = pd.to_datetime(d[c], errors='coerce')
    return d

def _write_df_sheet(ws, df):
    if df is None or len(df.columns) == 0:
        return
    df = _prepare_for_output(df)
    bold = Font(bold=True)
    for j, c in enumerate(df.columns, 1):
        ws.cell(row=1, column=j, value=c).font = bold
        ws.column_dimensions[get_column_letter(j)].width = min(40, max(12, len(str(c)) + 4))
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if v is None or (not isinstance(v, (list, tuple)) and pd.isna(v)):
                continue
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            elif hasattr(v, 'item') and not isinstance(v, (str, bytes)):
                v = v.item()
            cell = ws.cell(row=i, column=j, value=v)
            if isinstance(v, datetime):
                cell.number_format = 'yyyy-mm-dd'
            elif isinstance(v, float):
                cell.number_format = '#,##0.00'

def _write_pivot_sheet(ws, p, title):
    notes = p['notes']
    n = len(notes)
    last_col = 2 + 2 * n + 2
    ws.sheet_view.rightToLeft = True

    dark = PatternFill('solid', fgColor='244062')
    light = PatternFill('solid', fgColor='C5D9F1')
    white_b = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    thin, thick = Side(style='thin'), Side(style='thick')
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def hdr(r, c, v=None, size=11):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name='Calibri', bold=True, color='FFFFFF', size=size)
        cell.fill = dark
        cell.alignment = center
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        return cell

    for r in (1, 2):
        for c in range(1, last_col + 1):
            hdr(r, c, None)
    ws.cell(row=1, column=1, value=title).font = Font(name='Calibri', bold=True, color='FFFFFF', size=18)
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=last_col)

    for c in range(1, last_col + 1):
        hdr(3, c)
        hdr(4, c)
        hdr(5, c)
    ws.cell(row=3, column=3, value=PIVOT_NOTE)
    ws.cell(row=3, column=4, value='Values')
    for i, note in enumerate(notes):
        c0 = 3 + 2 * i
        ws.cell(row=4, column=c0, value=note).font = Font(name='Calibri', bold=True, color='FFFFFF', size=14)
        ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 1)
        ws.cell(row=5, column=c0, value='عدد الحسابات')
        ws.cell(row=5, column=c0 + 1, value='متبقي المديونية')
    tc = 3 + 2 * n
    ws.cell(row=4, column=tc, value='Total عدد الحسابات')
    ws.merge_cells(start_row=4, start_column=tc, end_row=5, end_column=tc)
    ws.cell(row=4, column=tc + 1, value='Total متبقي المديونية')
    ws.merge_cells(start_row=4, start_column=tc + 1, end_row=5, end_column=tc + 1)
    ws.cell(row=5, column=1, value='اسم المشرف ')
    ws.cell(row=5, column=2, value='اسم المحصل ')

    def body(r, c, v, fmt=None, bold=False):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name='Calibri', bold=bold, size=11)
        cell.alignment = center
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        if fmt:
            cell.number_format = fmt
        return cell

    r = 6
    for t in p['teams']:
        start = r
        for row in t['rows']:
            body(r, 2, row['coll'])
            for i, note in enumerate(notes):
                body(r, 3 + 2 * i, row['by_note'][note][0])
                body(r, 4 + 2 * i, row['by_note'][note][1], SUM_FMT)
            body(r, tc, row['cnt'])
            body(r, tc + 1, row['amt'], SUM_FMT)
            r += 1
        a = body(start, 1, t['team'])
        a.fill = light
        a.border = Border(left=thick, right=thin, top=thin, bottom=thin)
        for rr in range(start + 1, r):
            cell = ws.cell(row=rr, column=1)
            cell.fill = light
            cell.border = Border(left=thick, right=thin, top=thin, bottom=thin)
        if r - 1 > start:
            ws.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)

    g = p['grand']
    for c in range(1, last_col + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = white_b
        cell.fill = dark
        cell.alignment = center
        cell.border = Border(left=thick if c == 1 else thin, right=thin, top=thin, bottom=thick)
    ws.cell(row=r, column=1, value='الاجمالي')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    for i, note in enumerate(notes):
        ws.cell(row=r, column=3 + 2 * i, value=g['by_note'][note][0])
        x = ws.cell(row=r, column=4 + 2 * i, value=g['by_note'][note][1])
        x.number_format = SUM_FMT
    ws.cell(row=r, column=tc, value=g['cnt'])
    x = ws.cell(row=r, column=tc + 1, value=g['amt'])
    x.number_format = SUM_FMT

    ws.column_dimensions['A'].width = 35.7
    ws.column_dimensions['B'].width = 30.7
    for c in range(3, 3 + 2 * n):
        ws.column_dimensions[get_column_letter(c)].width = 18.9
    ws.column_dimensions[get_column_letter(tc)].width = 16.7
    ws.column_dimensions[get_column_letter(tc + 1)].width = 17.6

def build_report_xlsx(active_df, broken_df, neglect_df=None):
    wb = Workbook()
    ws = wb.active
    ws.title = 'تقرير الوعود القائمة'
    _write_df_sheet(ws, active_df)

    ws = wb.create_sheet('ملخص تقرير الوعود القائمة')
    _write_pivot_sheet(ws, build_pivot(active_df, order='first'), 'تقرير الوعود القائمة البنك الاهلي السعودي SNB (NPL&60Dpd)')

    ws = wb.create_sheet('تقرير الوعود المكسورة')
    _write_df_sheet(ws, broken_df)

    ws = wb.create_sheet('ملخص تقرير الوعود المكسورة')
    _write_pivot_sheet(ws, build_pivot(broken_df, order='desc'), 'تقرير الوعود المكسورة البنك الاهلي SNB (NPL&60Dpd)')

    if neglect_df is not None and len(neglect_df.columns) > 0:
        ws = wb.create_sheet('تقرير الإهمال')
        _write_df_sheet(ws, neglect_df)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ==========================================
# إدارة الملفات لكل صفحة تقرير بشكل منفصل تماماً
# ==========================================
if 'page_files' not in st.session_state:
    st.session_state.page_files = {
        'Overview': None,
        'Promises': None,
        'Neglect': None,
        'Activity': None
    }

curr_page = st.session_state.current_page

# شريط التحكم العلوي: خانة رفع الملفات مكان تاريخ التقرير مع الجملة المطلوبة + شريط البحث
ctrl_cols = st.columns([5, 5])
with ctrl_cols[0]:
    st.markdown(f"<p style='color: {text_color}; font-weight: 800; margin-bottom: 2px; direction: rtl;'>ارفعي الملفات هنا ي رنا</p>", unsafe_allow_html=True)
    uploaded_files_top = st.file_uploader("ارفعي الملفات هنا ي رنا", type=["xlsx", "xls", "csv"], accept_multiple_files=True, label_visibility="collapsed", key=f"top_uploader_{curr_page}")
    if uploaded_files_top:
        st.session_state.page_files[curr_page] = uploaded_files_top
        st.rerun()
with ctrl_cols[1]:
    st.markdown("<div style='padding-top: 25px;'>", unsafe_allow_html=True)
    search_query = st.text_input("", placeholder="🔍 ابحثي هنا (برقم العميل، الاسم، أو أي بيان)...", label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown(f"<hr style='margin: 15px 0 25px 0; border-color: {border_color}; clear: both;'>", unsafe_allow_html=True)

# دالة معالجة بيانات أي صفحة
def process_page_data(uploaded_files):
    if not uploaded_files:
        return None, None, None, None, None, 0, 0, 0, 0, 0, pd.DataFrame()
    
    today = pd.Timestamp(date.today()).normalize()
    promises_dfs = []
    activity_dfs = []
    all_dfs = []

    for f in uploaded_files:
        try:
            temp_df = pd.read_csv(f, dtype=str) if f.name.endswith('.csv') else pd.read_excel(f, dtype=str)
            temp_df.columns = temp_df.columns.astype(str).str.strip()
            all_dfs.append(temp_df)
            if 'Follow up Due Date' in temp_df.columns and 'Follow up Last Date' in temp_df.columns:
                promises_dfs.append(temp_df)
            has_created = any(c.lower().strip() in ['created on', 'created_on', 'createdon', 'تاريخ الإنشاء', 'وقت المعالجة', 'date', 'التاريخ'] for c in temp_df.columns)
            if has_created or not ('Follow up Due Date' in temp_df.columns):
                activity_dfs.append(temp_df)
        except Exception as e:
            st.error(f"خطأ في قراءة الملف {f.name}: {e}")

    df_all = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
    df_promises = pd.concat(promises_dfs, ignore_index=True) if promises_dfs else pd.DataFrame()
    df_act_raw = pd.concat(activity_dfs, ignore_index=True) if activity_dfs else df_all.copy()

    def find_col_in(target_df, candidates):
        if target_df.empty:
            return None
        for c in candidates:
            for col in target_df.columns:
                if col.lower().strip() == c.lower().strip():
                    return col
        return None

    created_col = find_col_in(df_act_raw, ['created on', 'created_on', 'createdon', 'تاريخ الإنشاء', 'وقت المعالجة', 'date', 'التاريخ'])
    if created_col:
        df_act_raw['Created_DateTime'] = pd.to_datetime(df_act_raw[created_col], errors='coerce')
        df_act_raw['Date'] = df_act_raw['Created_DateTime'].dt.date
        df_act_raw['Time'] = df_act_raw['Created_DateTime'].dt.time

    collector_col = find_col_in(df_act_raw, ['collecter', 'collector', 'المحصل', 'اسم المحصل', 'الموظف', 'user', 'agent'])
    if not collector_col:
        collector_col = find_col_in(df_all, ['collecter', 'collector', 'المحصل', 'اسم المحصل', 'الموظف', 'user', 'agent'])

    id_col = find_col_in(df_act_raw, ['id numer', 'id number', 'id_number', 'id', 'رقم الهوية', 'رقم العميل', 'account id'])
    notes_col = find_col_in(df_act_raw, ['notes', 'note', 'الملاحظات', 'ملاحظات', 'feedback'])

    dedup_cols = [c for c in [collector_col, id_col, notes_col] if c is not None and c in df_act_raw.columns]
    activity_base_df = df_act_raw.drop_duplicates(subset=dedup_cols).copy() if len(dedup_cols) == 3 else df_act_raw.copy()

    active_promises_df = pd.DataFrame()
    broken_promises_df = pd.DataFrame()
    neglect_df = pd.DataFrame()
    total_unhandled = 0

    if not df_promises.empty and 'Follow up Due Date' in df_promises.columns and 'Follow up Last Date' in df_promises.columns:
        due_date = pd.to_datetime(df_promises['Follow up Due Date'], errors='coerce').dt.normalize()
        last_date = pd.to_datetime(df_promises['Follow up Last Date'], errors='coerce').dt.normalize()
        base_condition = (
            (df_promises['حالة المعالجة - التمويل'].astype(str).str.strip() == 'غير معالج') & 
            (df_promises['تحديث فئة العميل - التمويل'].astype(str).str.strip().str.lower() != 'writeoff') & 
            (df_promises['Classification'].astype(str).str.strip() == 'وعود السداد قيد التنفيذ')
        )
        active_promises_df = df_promises[base_condition & (due_date == today) & (last_date != today)].copy()
        broken_promises_df = df_promises[base_condition & (due_date < today)].copy()
        _diff = (last_date[broken_promises_df.index] - due_date[broken_promises_df.index]).dt.days
        broken_promises_df = broken_promises_df[_diff < 0]
        if not broken_promises_df.empty:
            broken_promises_df.insert(broken_promises_df.columns.get_loc('Follow up Last Date') + 1, 'الفرق ', _diff[_diff < 0].astype(int))

        neglect_base = df_promises[
            (df_promises['حالة المعالجة - التمويل'].astype(str).str.strip() == 'غير معالج') & 
            (df_promises['تحديث فئة العميل - التمويل'].astype(str).str.strip().str.lower() != 'writeoff')
        ].copy()
        n_due_date = pd.to_datetime(neglect_base['Follow up Due Date'], errors='coerce').dt.normalize()
        neglect_base = neglect_base[n_due_date != today]
        n_last_date = pd.to_datetime(neglect_base['Follow up Last Date'], errors='coerce').dt.normalize()
        neglect_base['Neglect_Days'] = (today - n_last_date).dt.days
        neglect_df = neglect_base[neglect_base['Neglect_Days'] > 3].copy()

        total_unhandled = len(df_promises[base_condition])

        if search_query:
            active_promises_df = active_promises_df[active_promises_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]
            broken_promises_df = broken_promises_df[broken_promises_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]
            neglect_df = neglect_df[neglect_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]

    output_buffer = build_report_xlsx(active_promises_df, broken_promises_df, neglect_df)
    report_bytes = output_buffer.getvalue()

    total_rows = len(df_promises) if not df_promises.empty else len(df_all)
    active_count = len(active_promises_df)
    broken_count = len(broken_promises_df)
    neglect_count = len(neglect_df)
    total_activity_count = len(activity_base_df)

    return active_promises_df, broken_promises_df, neglect_df, activity_base_df, report_bytes, total_rows, total_unhandled, active_count, broken_count, neglect_count, total_activity_count

# التحقق من الملفات الخاصة بالصفحة الحالية فقط
page_uploaded = st.session_state.page_files[curr_page]

if page_uploaded is None:
    # 1. لو في الصفحة الرئيسية (Overview)، نعرض جملة الترحيب المخصصة لرنا
    if curr_page == 'Overview':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 15px; text-align: center; padding: 60px 20px; direction: rtl; border: 1px solid {border_color}; margin-top: 20px; margin-bottom: 30px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
            <h2 style="color: {text_color}; margin-bottom: 15px; font-weight:900;">أهلاً بيكي يا رنا في منصة تقارير شركة إجادة لتحصيل الديون</h2>
            <p style="color: {sub_text_color}; font-size: 18px; font-weight: bold;">من شريط التنقل بالأعلى، اختاري أي تقرير لعرضه ورفع ملفاته الخاصة.</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info(f"📁 يرجى رفع الملفات الخاصة بصفحة [{curr_page}] من خانة الرفع بالأعلى للبدء.")
else:
    active_promises_df, broken_promises_df, neglect_df, activity_base_df, report_bytes, total_rows, total_unhandled, active_count, broken_count, neglect_count, total_activity_count = process_page_data(st.session_state.page_files[curr_page])

    # ==========================================
    # صفحة الـ Overview الشاملة
    # ==========================================
    if curr_page == 'Overview':
        st.markdown(f"<h2 style='text-align:center; margin:0; color:{ejada_navy_light}; font-weight:900;'>📊 لوحة القيادة الشاملة (Executive Overview)</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center; color:#FFFFFF; font-weight: bold; margin-bottom: 30px;'>ملخص أداء محفظة الديون ونشاط المحصلين</p>", unsafe_allow_html=True)
        
        st.markdown(f"""
        <div style="display: grid; grid-template-columns: repeat(6, 1fr); gap: 15px; margin-bottom: 30px;">
            <div class="kpi-card" style="border-left-color: {border_color}; padding: 15px;">
                <div class="kpi-icon">📋</div>
                <div class="kpi-title">إجمالي الحسابات</div>
                <div class="kpi-value">{total_rows:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #586F8C; padding: 15px;">
                <div class="kpi-icon">⏳</div>
                <div class="kpi-title">وعود قيد التنفيذ</div>
                <div class="kpi-value">{total_unhandled:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #304D73; padding: 15px;">
                <div class="kpi-icon">✅</div>
                <div class="kpi-title">وعود قائمة</div>
                <div class="kpi-value">{active_count:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #586F8C; padding: 15px;">
                <div class="kpi-icon">⚠️</div>
                <div class="kpi-title">وعود مكسورة</div>
                <div class="kpi-value">{broken_count:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #99AABF; padding: 15px;">
                <div class="kpi-icon">🚨</div>
                <div class="kpi-title">حالات إهمال</div>
                <div class="kpi-value">{neglect_count:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: {border_color}; padding: 15px;">
                <div class="kpi-icon">📞</div>
                <div class="kpi-title">إجمالي النشاط</div>
                <div class="kpi-value">{total_activity_count:,}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        chart_col1, chart_col2, chart_col3 = st.columns(3)

        with chart_col1:
            with st.container(border=True):
                fig_donut = go.Figure(go.Pie(
                    labels=['الوعود القائمة', 'الوعود المكسورة', 'حالات الإهمال'],
                    values=[active_count, broken_count, neglect_count],
                    hole=0.6,
                    marker_colors=["#1F4373", "#304D73", "#586F8C"],
                    textinfo='percent+label',
                    textfont_color="#FFFFFF"
                ))
                fig_donut.update_layout(
                    title=dict(text="توزيع الوعود وحالات الإهمال", x=0.5, xanchor='center', font=dict(color=ejada_navy_dark, size=15, weight='bold')),
                    paper_bgcolor='white', plot_bgcolor='white', font_color=ejada_navy_dark, 
                    margin=dict(t=50, b=10, l=10, r=10), showlegend=False
                )
                st.plotly_chart(fig_donut, use_container_width=True, config={'displayModeBar': False})

        with chart_col2:
            with st.container(border=True):
                total_extracted = active_count + broken_count
                gauge_val = int((active_count / total_extracted * 100)) if total_extracted > 0 else 0
                fig_gauge = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=gauge_val,
                    number={'suffix': "%", 'font': {'color': ejada_navy_dark, 'size': 34, 'weight': 'bold'}},
                    title_text="نسبة الوعود القائمة من الإجمالي",
                    gauge={
                        'axis': {'range': [0, 100], 'visible': False},
                        'bar': {'color': "#304D73"}, 'bgcolor': "#E2E8F0", 'borderwidth': 0,
                        'steps': [{'range': [0, 50], 'color': '#99AABF'}, {'range': [50, 100], 'color': '#586F8C'}]
                    }
                ))
                fig_gauge.update_layout(
                    title_font=dict(color=ejada_navy_dark, size=15, weight='bold'), title_x=0.5,
                    paper_bgcolor='white', font_color=ejada_navy_dark, margin=dict(l=20, r=20, t=50, b=20)
                )
                st.plotly_chart(fig_gauge, use_container_width=True, config={'displayModeBar': False})

        with chart_col3:
            with st.container(border=True):
                collector_col = find_col_in(activity_base_df, ['collecter', 'collector', 'المحصل', 'اسم المحصل'])
                if collector_col and collector_col in activity_base_df.columns and total_activity_count > 0:
                    top_colls = activity_base_df[collector_col].value_counts().head(5)
                    fig_bar = go.Figure(go.Bar(
                        x=top_colls.index, y=top_colls.values,
                        marker_color=["#1F4373", "#304D73", "#586F8C", "#99AABF", "#1F4373"],
                        text=top_colls.values, textposition='auto', textfont=dict(color="#FFFFFF", weight='bold')
                    ))
                    fig_bar.update_layout(
                        title=dict(text="أكثر 5 محصلين نشاطاً", x=0.5, xanchor='center', font=dict(color=ejada_navy_dark, size=15, weight='bold')),
                        paper_bgcolor='white', plot_bgcolor='white', font_color=ejada_navy_dark, margin=dict(t=50, b=10, l=10, r=10)
                    )
                    fig_bar.update_yaxes(showgrid=False, visible=False)
                    st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info("لا توجد بيانات نشاط كافية لعرض رسم بياني.")
                
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(label=f"📥 تحميل التقرير الكامل.xlsx", data=report_bytes, file_name=f"Daily_Promises_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    # ==========================================
    # صفحة الوعود القائمة والمكسورة (مدمجة)
    # ==========================================
    elif curr_page == 'Promises':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:#1F4373; font-weight:900;">🤝 إدارة ومتابعة الوعود (القائمة والمكسورة)</h3>
        </div>
        """, unsafe_allow_html=True)
        
        tab_active, tab_broken = st.tabs([f"✅ الوعود القائمة ({active_count:,})", f"⚠️ الوعود المكسورة ({broken_count:,})"])
        
        with tab_active:
            st.markdown(f"<h4 style='text-align:right; direction:rtl; color:#FFFFFF; font-weight:900;'>📊 ملخص تقرير الوعود القائمة (Pivot)</h4>", unsafe_allow_html=True)
            st.dataframe(pivot_to_frame(build_pivot(active_promises_df, order='first')), use_container_width=True, hide_index=True)
            st.markdown(f"<h4 style='text-align:right; direction:rtl; color:#FFFFFF; font-weight:900; margin-top:20px;'>📋 تفاصيل شيت الوعود القائمة</h4>", unsafe_allow_html=True)
            st.dataframe(active_promises_df, use_container_width=True, height=400)
            
        with tab_broken:
            st.markdown(f"<h4 style='text-align:right; direction:rtl; color:#FFFFFF; font-weight:900;'>📊 ملخص تقرير الوعود المكسورة (Pivot)</h4>", unsafe_allow_html=True)
            st.dataframe(pivot_to_frame(build_pivot(broken_promises_df, order='desc')), use_container_width=True, hide_index=True)
            st.markdown(f"<h4 style='text-align:right; direction:rtl; color:#FFFFFF; font-weight:900; margin-top:20px;'>📋 تفاصيل شيت الوعود المكسورة</h4>", unsafe_allow_html=True)
            st.dataframe(broken_promises_df, use_container_width=True, height=400)
            
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button("📥 تحميل التقرير الشامل (الإكسيل بالملخصات).xlsx", data=report_bytes, file_name=f"Daily_Promises_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_promises_all")
        
    elif curr_page == 'Neglect':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:#586F8C; font-weight:900;">🚨 تقرير الإهمال (تأخير أكثر من 3 أيام)</h3>
        </div>
        """, unsafe_allow_html=True)
        default_sub_states = ['وعد بسداد مبلغ المعالجة', 'واعد بسداد مبلغ المعالجه', 'رافض السداد', 'الجوال مغلق', 'مماطل', 'تم سداد جزء وليس مبلغ المعالجه', 'لا يجيب', 'متقاعد الراتب علي بنك آخر', 'يرغب بجدوله المديونيه']
        selected_sub_states = st.multiselect("فلاتر الـ Sub State المشمولة:", options=default_sub_states, default=default_sub_states)
        sub_col = 'Sub State' if 'Sub State' in neglect_df.columns else 'sub state' if 'sub state' in neglect_df.columns else None
        final_neglect_df = neglect_df[neglect_df[sub_col].isin(selected_sub_states)] if sub_col else neglect_df
        st.markdown(f"<p style='text-align:right; direction:rtl; color:#FFFFFF; font-weight: bold;'>العدد الإجمالي: <b style='color:#586F8C; font-size:18px;'>{len(final_neglect_df):,}</b> حساب مهمل</p>", unsafe_allow_html=True)
        st.dataframe(final_neglect_df, use_container_width=True, height=450)

    elif curr_page == 'Activity':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:{text_color}; font-weight:900;">📈 تقرير النشاط وموديل التقييم (ناجحة / مغطاة)</h3>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("📁 ارفعي موديل النشاط (شيت التقييم 1 و 0) - اختياري", expanded=True):
            st.markdown(f"<p style='font-size: 13px; color: {sub_text_color}; direction: rtl;'>التقييم (1) يعني 'ناجحة'، والتقييم (0) يعني 'مغطاة'.</p>", unsafe_allow_html=True)
            mapping_file = st.file_uploader("شيت موديل النشاط (اختياري)", type=["xlsx"], key="mapping_uploader")

        filter_act_col1, filter_act_col2, filter_act_col3 = st.columns([2, 1.5, 1.5])
        with filter_act_col1:
            shift_option = st.selectbox("⏰ اختاري فترة النشاط (الشفت):", ["كل اليوم بدون تحديد", "من 2:00 م إلى 12:00 منتصف الليل (14:00 - 00:00)", "من 12:00 ظهراً إلى 4:00 عصراً (12:00 - 16:00)", "من 4:00 عصراً إلى آخر اليوم (16:00 - 23:59)"])
        with filter_act_col2:
            exclude_last_hour = st.checkbox("⏳ استبعاد آخر ساعة من الفترة", value=False)
        with filter_act_col3:
            exclude_break = st.checkbox("☕ استبعاد البريك (12:00 - 12:30)", value=False)

        created_col = find_col_in(activity_base_df, ['created on', 'created_on', 'createdon', 'تاريخ الإنشاء', 'وقت المعالجة', 'date', 'التاريخ'])
        if created_col and 'Created_DateTime' in activity_base_df.columns and activity_base_df['Created_DateTime'].notna().any():
            act_df = activity_base_df.dropna(subset=['Created_DateTime']).copy()
            act_times = act_df['Created_DateTime'].dt.time
            if shift_option == "من 2:00 م إلى 12:00 منتصف الليل (14:00 - 00:00)":
                start_t, end_t = time(14, 0), time(23, 0) if exclude_last_hour else time(23, 59, 59)
                act_df = act_df[(act_times >= start_t) & (act_times <= end_t)]
            elif shift_option == "من 12:00 ظهراً إلى 4:00 عصراً (12:00 - 16:00)":
                start_t, end_t = time(12, 0), time(15, 0) if exclude_last_hour else time(16, 0)
                act_df = act_df[(act_times >= start_t) & (act_times <= end_t)]
            elif shift_option == "من 4:00 عصراً إلى آخر اليوم (16:00 - 23:59)":
                start_t, end_t = time(16, 0), time(23, 0) if exclude_last_hour else time(23, 59, 59)
                act_df = act_df[(act_times >= start_t) & (act_times <= end_t)]
            if exclude_break:
                curr_times = act_df['Created_DateTime'].dt.time
                act_df = act_df[~((curr_times >= time(12, 0)) & (curr_times <= time(12, 30)))]
        else:
            act_df = activity_base_df.copy()

        if mapping_file is not None:
            map_df = pd.read_excel(mapping_file)
            map_df.columns = map_df.columns.str.strip()
            common_cols = [c for c in map_df.columns if c in act_df.columns]
            if common_cols:
                join_key = common_cols[0] 
                target_status_col = [c for c in map_df.columns if c != join_key]
                if target_status_col:
                    act_df = act_df.merge(map_df[[join_key, target_status_col[0]]], on=join_key, how='left')
                    def classify_call(val):
                        val_str = str(val).strip()
                        if val_str in ['1', '1.0', '1,0']: return 'ناجحة'
                        elif val_str in ['0', '0.0', '0,0']: return 'مغطاة'
                        else: return 'غير مصنف'
                    actual_col = target_status_col[0]
                    if actual_col in act_df.columns:
                        act_df['Call_Status'] = act_df[actual_col].apply(classify_call)
                    else:
                        act_df['Call_Status'] = 'غير مصنف'
            else:
                act_df['Call_Status'] = 'غير مصنف'
        else:
            target_col = find_col_in(act_df, ['التقييم', 'التقيمي', 'evaluation', 'eval', 'تقييم', 'status', '
