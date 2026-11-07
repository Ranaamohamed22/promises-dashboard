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
nav_cols = st.columns([1.3, 1.3, 1.3, 1.3, 5], gap="small") 

with nav_cols[0]:
    if st.button("✅ الوعود القائمة", type="primary" if st.session_state.current_page == 'Active' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Active'
        st.rerun()

with nav_cols[1]:
    if st.button("⚠️ الوعود المكسورة", type="primary" if st.session_state.current_page == 'Broken' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Broken'
        st.rerun()

with nav_cols[2]:
    if st.button("🚨 تقرير الإهمال", type="primary" if st.session_state.current_page == 'Neglect' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Neglect'
        st.rerun()

with nav_cols[3]:
    if st.button("📈 النشاط", type="primary" if st.session_state.current_page == 'Activity' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Activity'
        st.rerun()

with nav_cols[4]:
    st.markdown("")

st.markdown(f"<hr style='margin: 10px 0 10px 0; border-color: {border_color}; clear: both;'>", unsafe_allow_html=True)

# ==========================================
# دوال تقرير الوعود (نفس شكل الشيت + نفس الـ Pivot Table)
# ==========================================
PIVOT_TEAM = 'Sales Team'
PIVOT_COLL = 'Salesperson'
PIVOT_NOTE = 'ملاحظات-التمويل'
PIVOT_ACC = 'Account Number'
PIVOT_AMT = 'Net Amount'
SUM_FMT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'


def build_pivot(df, order='desc'):
    """order='desc' ترتيب Z→A (زي شيت المكسورة) | order='first' ترتيب ظهور الصفوف في الشيت (زي شيت القائمة)
    Pivot: الصفوف = المشرف ثم المحصل | الأعمدة = ملاحظات-التمويل | القيم = عدد الحسابات + مجموع متبقي المديونية"""
    need = [PIVOT_TEAM, PIVOT_COLL, PIVOT_NOTE, PIVOT_ACC, PIVOT_AMT]
    if df is None or df.empty or any(c not in df.columns for c in need):
        return {'notes': ['NPL'], 'teams': [], 'grand': {'by_note': {'NPL': (0, 0.0)}, 'cnt': 0, 'amt': 0.0}}

    d = df[need].copy()
    d['_c'] = d[PIVOT_ACC].notna().astype(int)
    d[PIVOT_AMT] = pd.to_numeric(d[PIVOT_AMT], errors='coerce').fillna(0.0)
    d[PIVOT_TEAM] = d[PIVOT_TEAM].fillna('(فارغ)').astype(str)
    d[PIVOT_COLL] = d[PIVOT_COLL].fillna('(فارغ)').astype(str)
    d[PIVOT_NOTE] = d[PIVOT_NOTE].fillna('(فارغ)').astype(str).str.strip()

    # الترتيب في الشيت الأصلي تنازلي (Z→A): NPL ثم Dpd60
    def _ordered(values):
        vals = list(pd.unique(values))
        return sorted(vals, reverse=True) if order == 'desc' else vals

    notes = sorted(d[PIVOT_NOTE].unique(), reverse=True)  # NPL ثم Dpd60 دايماً
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
    """تحويل الـ Pivot لجدول للعرض على الشاشة"""
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
    """نفس أنواع الأعمدة في الشيت الأصلي: المبالغ أرقام، والتواريخ الأساسية تواريخ، والباقي نصوص"""
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
    """نفس تنسيق شيت الملخص: عنوان - رؤوس - مشرف/محصل - إجمالي"""
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

    # عنوان التقرير
    for r in (1, 2):
        for c in range(1, last_col + 1):
            hdr(r, c, None)
    ws.cell(row=1, column=1, value=title).font = Font(name='Calibri', bold=True, color='FFFFFF', size=18)
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=last_col)

    # الصفوف 3-5
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

    # صف الإجمالي
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


# ==========================================
# Pivot Table حقيقي داخل ملف الإكسيل (Cache + Pivot Table definition)
# ==========================================
import re
import zipfile
from xml.sax.saxutils import escape as _xml_escape

_ILLEGAL_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _xa(v):
    """نص آمن داخل attribute في XML"""
    return _xml_escape(_ILLEGAL_XML.sub('', str(v)), {'"': '&quot;'})


def _cell_kind(v):
    if isinstance(v, datetime):
        return 'd'
    if isinstance(v, bool):
        return 's'
    if isinstance(v, (int, float)):
        return 'n'
    return 's'


def _normalize_value(v):
    if v is None or (not isinstance(v, (list, tuple)) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, 'item') and not isinstance(v, (str, bytes)):
        return v.item()
    return v


def _pivot_parts(data_df, data_sheet, p, order, idx):
    """يبني ملفات الـ Pivot (cache definition + records + pivot table) لجدول بيانات واحد"""
    df = _prepare_for_output(data_df)
    cols = list(df.columns)
    n_rows = len(df)
    team_i, coll_i, note_i = cols.index(PIVOT_TEAM), cols.index(PIVOT_COLL), cols.index(PIVOT_NOTE)
    acc_i, amt_i = cols.index(PIVOT_ACC), cols.index(PIVOT_AMT)

    values = [[_normalize_value(v) for v in df[c].tolist()] for c in cols]

    # نوع كل عمود
    kinds = []
    for col_vals in values:
        ks = {_cell_kind(v) for v in col_vals if v is not None}
        if not ks:
            kinds.append('blank')
        elif ks == {'n'}:
            kinds.append('n')
        elif ks == {'d'}:
            kinds.append('d')
        else:
            kinds.append('s')

    # الأعمدة اللي على المحاور بتتخزن كـ shared items (بترتيب ظهورها في البيانات)
    shared_cols = {team_i, coll_i, note_i}
    shared_maps = {}
    for ci in shared_cols:
        vals = ['' if v is None else str(v).strip() if ci == note_i else str(v) for v in values[ci]]
        values[ci] = vals
        uniq = list(dict.fromkeys(vals))
        shared_maps[ci] = {u: k for k, u in enumerate(uniq)}

    # ---------- cache definition ----------
    fields_xml = []
    for ci, name in enumerate(cols):
        if ci in shared_cols:
            items = ''.join(f'<s v="{_xa(u)}"/>' for u in shared_maps[ci])
            fields_xml.append(f'<cacheField name="{_xa(name)}" numFmtId="0"><sharedItems count="{len(shared_maps[ci])}">{items}</sharedItems></cacheField>')
            continue
        col_vals = values[ci]
        has_blank = any(v is None for v in col_vals)
        k = kinds[ci]
        if k == 'blank':
            si = '<sharedItems containsNonDate="0" containsString="0" containsBlank="1"/>'
        elif k == 'n':
            nums = [float(v) for v in col_vals if v is not None]
            is_int = all(float(v).is_integer() for v in nums)
            attrs = 'containsSemiMixedTypes="0" containsString="0" containsNumber="1"' if not has_blank else 'containsString="0" containsBlank="1" containsNumber="1"'
            if is_int:
                attrs += ' containsInteger="1"'
            si = f'<sharedItems {attrs} minValue="{repr(min(nums))}" maxValue="{repr(max(nums))}"/>'
        elif k == 'd':
            ds = [v for v in col_vals if v is not None]
            attrs = 'containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0"' if not has_blank else 'containsNonDate="0" containsDate="1" containsString="0" containsBlank="1"'
            si = f'<sharedItems {attrs} minDate="{min(ds).strftime("%Y-%m-%dT%H:%M:%S")}" maxDate="{max(ds).strftime("%Y-%m-%dT%H:%M:%S")}"/>'
        else:
            si = '<sharedItems containsBlank="1"/>' if has_blank else '<sharedItems/>'
        fields_xml.append(f'<cacheField name="{_xa(name)}" numFmtId="0">{si}</cacheField>')

    last_col_letter = get_column_letter(len(cols))
    ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    cache_def = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition {ns} r:id="rId1" refreshedBy="Ejada" '
        f'refreshedDate="46330.5" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="{n_rows}">'
        f'<cacheSource type="worksheet"><worksheetSource ref="A1:{last_col_letter}{n_rows + 1}" sheet="{_xa(data_sheet)}"/></cacheSource>'
        f'<cacheFields count="{len(cols)}">{"".join(fields_xml)}</cacheFields></pivotCacheDefinition>'
    )

    # ---------- records ----------
    recs = []
    for ri in range(n_rows):
        cells = []
        for ci in range(len(cols)):
            v = values[ci][ri]
            if ci in shared_cols:
                cells.append(f'<x v="{shared_maps[ci][v]}"/>')
            elif v is None:
                cells.append('<m/>')
            elif kinds[ci] == 'n':
                cells.append(f'<n v="{repr(float(v)) if not float(v).is_integer() else int(float(v))}"/>')
            elif kinds[ci] == 'd':
                cells.append(f'<d v="{v.strftime("%Y-%m-%dT%H:%M:%S")}"/>')
            else:
                cells.append(f'<s v="{_xa(v if not isinstance(v, datetime) else v.strftime("%Y-%m-%d %H:%M:%S"))}"/>')
        recs.append('<r>' + ''.join(cells) + '</r>')
    cache_records = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords {ns} count="{n_rows}">{"".join(recs)}</pivotCacheRecords>'

    # ---------- pivot table ----------
    notes = p['notes']
    n_notes = len(notes)
    # ترتيب العناصر على المحاور (display order)
    team_order = [t['team'] for t in p['teams']]
    coll_order = []
    for t in p['teams']:
        for r in t['rows']:
            if r['coll'] not in coll_order:
                coll_order.append(r['coll'])
    note_pos = {n: k for k, n in enumerate(notes)}
    team_pos = {t: k for k, t in enumerate(team_order)}
    coll_pos = {c: k for k, c in enumerate(coll_order)}

    def items_xml(order_list, ci):
        return f'<items count="{len(order_list)}">' + ''.join(f'<item x="{shared_maps[ci][o]}"/>' for o in order_list) + '</items>'

    pf = []
    for ci in range(len(cols)):
        base = 'compact="0" outline="0" showAll="0" defaultSubtotal="0"'
        if ci == note_i:
            pf.append(f'<pivotField axis="axisCol" {base} sortType="descending">{items_xml(notes, ci)}</pivotField>')
        elif ci == team_i:
            pf.append(f'<pivotField name="اسم المشرف " axis="axisRow" {base}>{items_xml(team_order, ci)}</pivotField>')
        elif ci == coll_i:
            st_attr = ' sortType="descending"' if order == 'desc' else ''
            pf.append(f'<pivotField name="اسم المحصل " axis="axisRow" {base}{st_attr}>{items_xml(coll_order, ci)}</pivotField>')
        elif ci == acc_i:
            pf.append(f'<pivotField dataField="1" compact="0" outline="0" showAll="0" defaultSubtotal="0"/>')
        elif ci == amt_i:
            pf.append(f'<pivotField dataField="1" compact="0" outline="0" showAll="0" defaultSubtotal="0"/>')
        else:
            pf.append(f'<pivotField compact="0" outline="0" showAll="0" defaultSubtotal="0"/>')

    row_items = []
    for t in p['teams']:
        for k, r in enumerate(t['rows']):
            if k == 0:
                tp = team_pos[t['team']]
                tx = '<x/>' if tp == 0 else f'<x v="{tp}"/>'
                cp = coll_pos[r['coll']]
                cx = '<x/>' if cp == 0 else f'<x v="{cp}"/>'
                row_items.append(f'<i>{tx}{cx}</i>')
            else:
                row_items.append(f'<i r="1"><x v="{coll_pos[r["coll"]]}"/></i>')
    row_items.append('<i t="grand"><x/></i>')

    col_items = []
    for k in range(n_notes):
        first = '<x/>' if k == 0 else f'<x v="{k}"/>'
        col_items.append(f'<i>{first}<x/></i>')
        col_items.append('<i r="1" i="1"><x v="1"/></i>')
    col_items.append('<i t="grand"><x/></i>')
    col_items.append('<i t="grand" i="1"><x/></i>')

    n_body = len(row_items) - 1
    last_row = 5 + n_body + 1
    last_col = 2 + 2 * n_notes + 2
    ref = f'A3:{get_column_letter(last_col)}{last_row}'

    pivot_def = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'name="PivotTable{idx}" cacheId="{idx}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" '
        f'applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" grandTotalCaption="الاجمالي" missingCaption="0" '
        f'updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" mergeItem="1" createdVersion="8" indent="0" '
        f'compact="0" compactData="0" multipleFieldFilters="0" rowHeaderCaption="المحصل" colHeaderCaption="الفرع">'
        f'<location ref="{ref}" firstHeaderRow="1" firstDataRow="3" firstDataCol="2"/>'
        f'<pivotFields count="{len(cols)}">{"".join(pf)}</pivotFields>'
        f'<rowFields count="2"><field x="{team_i}"/><field x="{coll_i}"/></rowFields>'
        f'<rowItems count="{len(row_items)}">{"".join(row_items)}</rowItems>'
        f'<colFields count="2"><field x="{note_i}"/><field x="-2"/></colFields>'
        f'<colItems count="{len(col_items)}">{"".join(col_items)}</colItems>'
        f'<dataFields count="2"><dataField name="عدد الحسابات" fld="{acc_i}" subtotal="count" baseField="{coll_i}" baseItem="0"/>'
        f'<dataField name="متبقي المديونية" fld="{amt_i}" baseField="0" baseItem="0" numFmtId="3"/></dataFields>'
        f'<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
        f'</pivotTableDefinition>'
    )
    return cache_def, cache_records, pivot_def


def inject_pivot_tables(xlsx_bytes, specs):
    """specs: قائمة {summary_sheet, data_sheet, data_df, pivot, order} - بتحوّل شيتات الملخص لـ Pivot Table حقيقي"""
    zin = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    files = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    wb_xml = files['xl/workbook.xml'].decode('utf-8')
    wb_rels = files['xl/_rels/workbook.xml.rels'].decode('utf-8')
    ct = files['[Content_Types].xml'].decode('utf-8')

    # اسم الشيت -> ملف الشيت
    sheet_rid = {}
    for m in re.finditer(r'<sheet\b[^>]*?>', wb_xml):
        tag = m.group(0)
        nm = re.search(r'name="([^"]*)"', tag).group(1)
        rid = re.search(r'r:id="([^"]*)"', tag).group(1)
        sheet_rid[nm.replace('&amp;', '&')] = rid
    rid_target = {}
    for m in re.finditer(r'<Relationship\b[^>]*?>', wb_rels):
        tag = m.group(0)
        rid_target[re.search(r'Id="([^"]*)"', tag).group(1)] = re.search(r'Target="([^"]*)"', tag).group(1)

    pivot_caches_xml = ''
    new_rels = ''
    new_ct = ''
    for idx, sp in enumerate(specs, 1):
        cache_def, cache_records, pivot_def = _pivot_parts(sp['data_df'], sp['data_sheet'], sp['pivot'], sp['order'], idx)
        files[f'xl/pivotCache/pivotCacheDefinition{idx}.xml'] = cache_def.encode('utf-8')
        files[f'xl/pivotCache/pivotCacheRecords{idx}.xml'] = cache_records.encode('utf-8')
        files[f'xl/pivotCache/_rels/pivotCacheDefinition{idx}.xml.rels'] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords{idx}.xml"/></Relationships>'
        ).encode('utf-8')
        files[f'xl/pivotTables/pivotTable{idx}.xml'] = pivot_def.encode('utf-8')
        files[f'xl/pivotTables/_rels/pivotTable{idx}.xml.rels'] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition{idx}.xml"/></Relationships>'
        ).encode('utf-8')

        # ربط الشيت بالـ Pivot
        target = rid_target[sheet_rid[sp['summary_sheet']]]
        sheet_file = target.lstrip('/') if target.startswith('/') else 'xl/' + target
        rels_path = sheet_file.replace('worksheets/', 'worksheets/_rels/') + '.rels'
        rel = (f'<Relationship Id="rIdPv{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" '
               f'Target="../pivotTables/pivotTable{idx}.xml"/>')
        if rels_path in files:
            files[rels_path] = files[rels_path].decode('utf-8').replace('</Relationships>', rel + '</Relationships>').encode('utf-8')
        else:
            files[rels_path] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                + rel + '</Relationships>').encode('utf-8')

        new_rels += (f'<Relationship Id="rIdPc{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" '
                     f'Target="pivotCache/pivotCacheDefinition{idx}.xml"/>')
        pivot_caches_xml += f'<pivotCache cacheId="{idx}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdPc{idx}"/>'
        new_ct += (f'<Override PartName="/xl/pivotCache/pivotCacheDefinition{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
                   f'<Override PartName="/xl/pivotCache/pivotCacheRecords{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>'
                   f'<Override PartName="/xl/pivotTables/pivotTable{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>')

    if specs:
        wb_rels = wb_rels.replace('</Relationships>', new_rels + '</Relationships>')
        ct = ct.replace('</Types>', new_ct + '</Types>')
        if re.search(r'<calcPr\b[^>]*/>', wb_xml):
            wb_xml = re.sub(r'(<calcPr\b[^>]*/>)', r'\1<pivotCaches>' + pivot_caches_xml.replace('\\', '\\\\') + '</pivotCaches>', wb_xml, count=1)
        else:
            wb_xml = wb_xml.replace('</workbook>', '<pivotCaches>' + pivot_caches_xml + '</pivotCaches></workbook>')
        files['xl/workbook.xml'] = wb_xml.encode('utf-8')
        files['xl/_rels/workbook.xml.rels'] = wb_rels.encode('utf-8')
        files['[Content_Types].xml'] = ct.encode('utf-8')

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        # [Content_Types].xml لازم أول ملف
        zout.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
        for n, b in files.items():
            zout.writestr(n, b)
    out.seek(0)
    return out


def _can_pivot(df):
    need = [PIVOT_TEAM, PIVOT_COLL, PIVOT_NOTE, PIVOT_ACC, PIVOT_AMT]
    return (df is not None and len(df) > 0 and all(c in df.columns for c in need)
            and len(set(df.columns)) == len(df.columns))


def build_report_xlsx(active_df, broken_df, neglect_df=None):
    """نفس ملف التقرير: 4 شيتات (قائمة + ملخصها + مكسورة + ملخصها) وبعدهم تقرير الإهمال لو موجود"""
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

    # تحويل شيتات الملخص لـ Pivot Table حقيقي (لو فيه بيانات)
    specs = []
    if _can_pivot(active_df):
        specs.append({'summary_sheet': 'ملخص تقرير الوعود القائمة', 'data_sheet': 'تقرير الوعود القائمة',
                      'data_df': active_df, 'pivot': build_pivot(active_df, order='first'), 'order': 'first'})
    if _can_pivot(broken_df):
        specs.append({'summary_sheet': 'ملخص تقرير الوعود المكسورة', 'data_sheet': 'تقرير الوعود المكسورة',
                      'data_df': broken_df, 'pivot': build_pivot(broken_df, order='desc'), 'order': 'desc'})
    if specs:
        try:
            return inject_pivot_tables(buf.getvalue(), specs)
        except Exception:
            pass  # لو حصل أي خطأ يفضل الملف بالجداول العادية
    buf.seek(0)
    return buf


# ==========================================
# 3. شريط التحكم (رفع الملفات المتعددة والبحث)
# ==========================================
filter_cols = st.columns([3, 1.6, 5.4])
with filter_cols[0]:
    uploaded_files = st.file_uploader("", type=["xlsx", "xls", "csv"], accept_multiple_files=True, label_visibility="collapsed")
with filter_cols[1]:
    report_date = st.date_input("📅 تاريخ التقرير", value=date.today(), format="DD/MM/YYYY")
with filter_cols[2]:
    st.markdown("<div style='padding-top: 30px;'>", unsafe_allow_html=True)
    search_query = st.text_input("", placeholder="🔍 ابحثي هنا (برقم العميل، الاسم، أو أي بيان)...", label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown(f"<hr style='margin: 10px 0 25px 0; border-color: {border_color};'>", unsafe_allow_html=True)

# ==========================================
# 4. معالجة البيانات المجمعة
# ==========================================
if uploaded_files:
    today = pd.Timestamp(report_date).normalize()
    
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
            st.error(f"حدث خطأ في قراءة الملف {f.name}: {e}")

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
    if len(dedup_cols) == 3:
        activity_base_df = df_act_raw.drop_duplicates(subset=dedup_cols).copy()
    else:
        activity_base_df = df_act_raw.copy()

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
        # عمود "الفرق " بعد Follow up Last Date زي الشيت الأصلي
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

    # ملف التقرير: 4 شيتات (قائمة + ملخص + مكسورة + ملخص) + تقرير الإهمال
    output_buffer = build_report_xlsx(active_promises_df, broken_promises_df, neglect_df)
    report_bytes = output_buffer.getvalue()

    total_rows = len(df_promises) if not df_promises.empty else len(df_all)
    active_count = len(active_promises_df)
    broken_count = len(broken_promises_df)
    neglect_count = len(neglect_df)
    total_activity_count = len(activity_base_df)

    # ==========================================
    # صفحة الـ Overview الشاملة
    # ==========================================
    if st.session_state.current_page == 'Overview':
        st.markdown(f"<h2 style='text-align:center; margin:0; color:{ejada_navy_light}; font-weight:900;'>📊 لوحة القيادة الشاملة (Executive Overview)</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center; color:#FFFFFF; font-weight: bold; margin-bottom: 30px;'>ملخص أداء محفظة الديون ونشاط المحصلين - {today.date()}</p>", unsafe_allow_html=True)
        
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

        # 1. Donut Chart داخل st.container بخلفية بيضاء صريحة
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
                    paper_bgcolor='white', 
                    plot_bgcolor='white',
                    font_color=ejada_navy_dark, 
                    margin=dict(t=50, b=10, l=10, r=10),
                    showlegend=False
                )
                st.plotly_chart(fig_donut, use_container_width=True, config={'displayModeBar': False})

        # 2. Gauge Chart داخل st.container بخلفية بيضاء صريحة
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
                        'bar': {'color': "#304D73"},
                        'bgcolor': "#E2E8F0",
                        'borderwidth': 0,
                        'steps': [
                            {'range': [0, 50], 'color': '#99AABF'},
                            {'range': [50, 100], 'color': '#586F8C'}
                        ]
                    }
                ))
                fig_gauge.update_layout(
                    title_font=dict(color=ejada_navy_dark, size=15, weight='bold'),
                    title_x=0.5,
                    paper_bgcolor='white', 
                    font_color=ejada_navy_dark, 
                    margin=dict(l=20, r=20, t=50, b=20)
                )
                st.plotly_chart(fig_gauge, use_container_width=True, config={'displayModeBar': False})

        # 3. Bar Chart داخل st.container بخلفية بيضاء صريحة
        with chart_col3:
            with st.container(border=True):
                if collector_col and collector_col in activity_base_df.columns and total_activity_count > 0:
                    top_colls = activity_base_df[collector_col].value_counts().head(5)
                    fig_bar = go.Figure(go.Bar(
                        x=top_colls.index,
                        y=top_colls.values,
                        marker_color=["#1F4373", "#304D73", "#586F8C", "#99AABF", "#1F4373"],
                        text=top_colls.values,
                        textposition='auto',
                        textfont=dict(color="#FFFFFF", weight='bold')
                    ))
                    fig_bar.update_layout(
                        title=dict(text="أكثر 5 محصلين نشاطاً", x=0.5, xanchor='center', font=dict(color=ejada_navy_dark, size=15, weight='bold')),
                        paper_bgcolor='white', 
                        plot_bgcolor='white', 
                        font_color=ejada_navy_dark, 
                        margin=dict(t=50, b=10, l=10, r=10)
                    )
                    fig_bar.update_yaxes(showgrid=False, visible=False)
                    st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info("لا توجد بيانات نشاط كافية لعرض رسم بياني.")
                
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 25px; border: 1px solid {border_color}; box-shadow: 0 4px 10px rgba(0,0,0,0.2);">
            <h5 style="margin:0 0 20px 0; color:{text_color}; font-weight:900;">Breakdown Overview</h5>
            <div style="display: flex; align-items: center; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid {border_color}; direction: rtl;">
                <span style="width:30%; color:{text_color}; font-weight: bold; text-align: right;">الوعود القائمة اليوم</span>
                <div style="width: 55%; background-color: {gauge_bg}; height: 8px; border-radius: 4px; overflow: hidden;"><div style="width:{gauge_val}%; background-color:#1F4373; height:100%;"></div></div>
                <span style="color:#1F4373; font-weight: bold; font-size: 14px;">{gauge_val}%</span>
            </div>
            <div style="display: flex; align-items: center; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid {border_color}; direction: rtl;">
                <span style="width:30%; color:{text_color}; font-weight: bold; text-align: right;">الوعود المكسورة</span>
                <div style="width: 55%; background-color: {gauge_bg}; height: 8px; border-radius: 4px; overflow: hidden;"><div style="width:{100-gauge_val if total_extracted > 0 else 0}%; background-color:#304D73; height:100%;"></div></div>
                <span style="color:#304D73; font-weight: bold; font-size: 14px;">{100-gauge_val if total_extracted > 0 else 0}%</span>
            </div>
            <div style="display: flex; align-items: center; justify-content: space-between; padding: 12px 0; direction: rtl;">
                <span style="width:30%; color:{text_color}; font-weight: bold; text-align: right;">حالات الإهمال</span>
                <div style="width: 55%; background-color: {gauge_bg}; height: 8px; border-radius: 4px; overflow: hidden;"><div style="width:{int(neglect_count/total_rows*100) if total_rows > 0 else 0}%; background-color:#586F8C; height:100%;"></div></div>
                <span style="color:#586F8C; font-weight: bold; font-size: 14px;">{int(neglect_count/total_rows*100) if total_rows > 0 else 0}%</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(label=f"📥 تحميل التقرير الكامل - {today.date()}.xlsx", data=report_bytes, file_name=f"Daily_Promises_Report_{today.date()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    # ==========================================
    # باقي الصفحات 
    # ==========================================
    elif st.session_state.current_page == 'Active':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:#1F4373; font-weight:900;">✅ شيت الوعود القائمة ({active_count:,} حساب)</h3>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"<h4 style='text-align:right; direction:rtl; color:#FFFFFF; font-weight:900;'>📊 ملخص تقرير الوعود القائمة (Pivot)</h4>", unsafe_allow_html=True)
        st.dataframe(pivot_to_frame(build_pivot(active_promises_df, order='first')), use_container_width=True, hide_index=True)
        st.dataframe(active_promises_df, use_container_width=True, height=450)
        st.download_button("📥 تحميل تقرير الوعود (القائمة + المكسورة + الملخصات).xlsx", data=report_bytes, file_name=f"Daily_Promises_Report_{today.date()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_active")
        
    elif st.session_state.current_page == 'Broken':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:#304D73; font-weight:900;">⚠️ شيت الوعود المكسورة ({broken_count:,} حساب)</h3>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"<h4 style='text-align:right; direction:rtl; color:#FFFFFF; font-weight:900;'>📊 ملخص تقرير الوعود المكسورة (Pivot)</h4>", unsafe_allow_html=True)
        st.dataframe(pivot_to_frame(build_pivot(broken_promises_df)), use_container_width=True, hide_index=True)
        st.dataframe(broken_promises_df, use_container_width=True, height=450)
        st.download_button("📥 تحميل تقرير الوعود (القائمة + المكسورة + الملخصات).xlsx", data=report_bytes, file_name=f"Daily_Promises_Report_{today.date()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_broken")
        
    elif st.session_state.current_page == 'Neglect':
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

    # ==========================================
    # صفحة تقرير النشاط (Activity Report) 
    # ==========================================
    elif st.session_state.current_page == 'Activity':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:{text_color}; font-weight:900;">📈 تقرير النشاط وموديل التقييم (ناجحة / مغطاة)</h3>
            <p style="margin: 5px 0 0 0; color: {sub_text_color}; font-size: 14px;">تم حذف التكرار تلقائياً بناءً على ({collector_col or 'collecter'}, {id_col or 'id numer'}, {notes_col or 'notes'}).</p>
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
            st.info("ℹ️ يتم عرض كافة سجلات النشاط (لم يتم العثور على عمود وقت Created On لتطبيق فلتر الساعات).")
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
                    act_df['Call_Status'] = act_df[target_status_col[0]].apply(classify_call)
            else:
                st.warning("⚠️ لم يتم العثور على عمود مشترك بين شيت النشاط الأساسي وموديل النشاط المرفوع.")
                act_df['Call_Status'] = 'غير مصنف'
        else:
            target_col = find_col_in(act_df, ['التقييم', 'التقيمي', 'evaluation', 'eval', 'تقييم', 'status', 'الحالة', 'sub state', 'result', 'نتيجة المكالمة'])
            target_col = target_col if target_col else act_df.columns[0]
            available_vals = act_df[target_col].dropna().astype(str).unique().tolist()
            selected_success_vals = st.multiselect(f"حددي القيم التي تعتبر (ناجحة = 1) من عمود [{target_col}]:", options=available_vals, default=[v for v in available_vals if str(v).strip() in ['1', '1.0', 'ناجح', 'وعد', 'سداد', 'تم']])
            act_df['Call_Status'] = act_df[target_col].astype(str).apply(lambda x: 'ناجحة' if x in selected_success_vals else 'مغطاة')

        if search_query:
            act_df = act_df[act_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]

        total_activity = len(act_df)
        success_count = len(act_df[act_df['Call_Status'] == 'ناجحة'])
        covered_count = len(act_df[act_df['Call_Status'] == 'مغطاة'])
        success_rate = (success_count / total_activity * 100) if total_activity > 0 else 0

        if collector_col and collector_col in act_df.columns and total_activity > 0:
            top_collector_series = act_df[collector_col].value_counts()
            if not top_collector_series.empty:
                top_name = top_collector_series.index[0]
                top_calls = top_collector_series.values[0]
                top_pct = (top_calls / total_activity * 100) if total_activity > 0 else 0
                
                st.markdown(f"""
                <div style="background-color: {card_color}; border-radius: 12px; padding: 20px 30px; border: 2px solid #586F8C; direction:rtl; margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between;">
                    <div>
                        <h4 style="margin:0; color: #586F8C; font-weight:900;">🏆 بطل الشفت (الأكثر نشاطاً)</h4>
                        <p style="margin: 5px 0 0 0; color: {text_color}; font-size: 18px;">المحصل: <b style="color:#1F4373; font-size: 22px;">{top_name}</b></p>
                    </div>
                    <div style="text-align: left; background-color: {bg_color}; padding: 10px 20px; border-radius: 8px; border: 1px solid #586F8C;">
                        <h3 style="margin:0; color: #586F8C; font-weight:900;">{top_calls:,} مكالمة</h3>
                        <p style="margin: 0; color: #FFFFFF; font-size: 14px; font-weight: bold;">استحوذ على <span style="color:#99AABF;">{top_pct:.1f}%</span> من إجمالي شغل اليوم</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown(f"""
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 20px;">
            <div class="kpi-card" style="border-left-color: {border_color};">
                <div class="kpi-icon">📞</div>
                <div class="kpi-title">إجمالي المحاولات</div>
                <div class="kpi-value">{total_activity:,}</div>
                <div class="kpi-sub" style="color: {sub_text_color};">بعد إزالة التكرار</div>
            </div>
            <div class="kpi-card" style="border-left-color: #1F4373;">
                <div class="kpi-icon">✅</div>
                <div class="kpi-title">المكالمات الناجحة (1)</div>
                <div class="kpi-value">{success_count:,}</div>
                <div class="kpi-sub" style="color: #1F4373;">تقييم ناجح</div>
            </div>
            <div class="kpi-card" style="border-left-color: {border_color};">
                <div class="kpi-icon">🎧</div>
                <div class="kpi-title">المكالمات المغطاة (0)</div>
                <div class="kpi-value">{covered_count:,}</div>
                <div class="kpi-sub" style="color: {sub_text_color};">تقييم مغطاة</div>
            </div>
            <div class="kpi-card" style="border-left-color: #586F8C;">
                <div class="kpi-icon">🎯</div>
                <div class="kpi-title">نسبة النجاح الإجمالية</div>
                <div class="kpi-value">{success_rate:.1f}%</div>
                <div class="kpi-sub" style="color: #586F8C;">معدل الوصول للعملاء</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if collector_col and collector_col in act_df.columns:
            pivot_summary = act_df.groupby([collector_col, 'Call_Status']).size().unstack(fill_value=0)
            if 'ناجحة' not in pivot_summary.columns: pivot_summary['ناجحة'] = 0
            if 'مغطاة' not in pivot_summary.columns: pivot_summary['مغطاة'] = 0
            
            pivot_summary['إجمالي المحاولات'] = pivot_summary['ناجحة'] + pivot_summary['مغطاة']
            pivot_summary['نسبة النجاح %'] = (pivot_summary['ناجحة'] / pivot_summary['إجمالي المحاولات'] * 100).round(1).astype(str) + '%'
            pivot_summary = pivot_summary.sort_values(by='إجمالي المحاولات', ascending=False).reset_index()

            with st.expander("📊 تفصيل أداء المحصلين (ناجحة مقابل مغطاة)", expanded=True):
                st.dataframe(pivot_summary, use_container_width=True)

        extra_time_cols = [c for c in ['Date', 'Time', 'Call_Status'] if c in act_df.columns]
        display_cols = extra_time_cols + [c for c in act_df.columns if c not in ['Created_DateTime', 'Date', 'Time', 'Call_Status']]
        st.dataframe(act_df[display_cols], use_container_width=True, height=450)

        act_buf = io.BytesIO()
        act_df[display_cols].to_excel(act_buf, index=False, sheet_name='تقرير النشاط والتقييم')
        act_buf.seek(0)
        st.download_button("📥 تحميل تقرير النشاط مفلتر.xlsx", data=act_buf, file_name=f"Activity_Success_Report_{today.date()}.xlsx")

else:
    st.markdown(f"""
    <div style="background-color: {card_color}; border-radius: 15px; text-align: center; padding: 100px 20px; direction: rtl; border: 1px solid {border_color}; margin-top: 50px;">
        <h2 style="color: {text_color}; margin-bottom: 15px; font-weight:900;">أهلاً بيكي يا رنا في منصة تقارير شركة إجادة لتحصيل الديون</h2>
        <p style="color: {sub_text_color}; font-size: 18px; font-weight: bold;">من شريط التحكم بالأعلى، ارفعي شيت الإكسيل اليومي لعرض الداشبورد والتقارير.</p>
    </div>
    """, unsafe_allow_html=True)
