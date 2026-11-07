import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# 1. إعدادات الصفحة والهوية البصرية (SNB Theme)
# ==========================================
st.set_page_config(
    page_title="منصة تقارير الأداء والتحصيل | البنك الأهلي السعودي",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# إدارة الوضع الليلي والنهاري
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

# القائمة الجانبية للإعدادات ورفع الملفات
with st.sidebar:
    st.markdown("### ⚙️ إعدادات المنصة")
    st.session_state.dark_mode = st.toggle("🌙 الوضع الليلي (Dark Mode)", value=st.session_state.dark_mode)
    
    st.markdown("---")
    st.markdown("### 📂 رفع ملفات التقارير")
    uploaded_files = st.file_uploader(
        "ارفعي ملف أو أكثر (الوعود، الإهمال، النشاط، أو أي تقرير إضافي)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True
    )
    
    st.markdown("---")
    st.markdown("### 📅 تاريخ التقرير المرجعي")
    ref_date = st.date_input("تاريخ التقييم المرجعي", value=datetime.today())

# تحديد ألوان الثيم بناءً على الوضع المختار
if st.session_state.dark_mode:
    bg_color = "#0e1117"
    card_bg = "#1e2530"
    text_color = "#ffffff"
    sub_text = "#a0aec0"
    border_color = "#2d3748"
else:
    bg_color = "#f4f7f6"
    card_bg = "#ffffff"
    text_color = "#1a202c"
    sub_text = "#4a5568"
    border_color = "#e2e8f0"

snb_green = "#006C35"
snb_light_green = "#009147"
snb_gold = "#C5A059"

# تنسيقات CSS المتقدمة ودعم RTL
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Cairo', sans-serif !important;
        direction: rtl;
        text-align: right;
        background-color: {bg_color};
        color: {text_color};
    }}
    
    .main-header {{
        background: linear-gradient(135deg, {snb_green} 0%, {snb_light_green} 100%);
        padding: 25px 30px;
        border-radius: 15px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0, 108, 53, 0.2);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    
    .main-header h1 {{
        margin: 0;
        font-size: 28px;
        font-weight: 800;
        color: #ffffff !important;
    }}
    
    .main-header p {{
        margin: 5px 0 0 0;
        font-size: 15px;
        color: #e2e8f0;
    }}
    
    .kpi-card {{
        background-color: {card_bg};
        border-right: 6px solid {snb_green};
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        border-top: 1px solid {border_color};
        border-left: 1px solid {border_color};
        border-bottom: 1px solid {border_color};
        margin-bottom: 15px;
        transition: transform 0.2s ease;
    }}
    
    .kpi-card:hover {{
        transform: translateY(-3px);
    }}
    
    .kpi-title {{
        font-size: 14px;
        font-weight: 600;
        color: {sub_text};
        margin-bottom: 8px;
    }}
    
    .kpi-value {{
        font-size: 26px;
        font-weight: 800;
        color: {text_color};
    }}
    
    .kpi-sub {{
        font-size: 12px;
        color: {snb_gold};
        font-weight: 700;
        margin-top: 5px;
    }}
    
    /* ضبط التبويبات والجداول لليمين */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 10px;
        direction: rtl;
    }}
    .stTabs [data-baseweb="tab"] {{
        font-family: 'Cairo', sans-serif !important;
        font-weight: 700;
        font-size: 15px;
        padding: 10px 20px;
        border-radius: 8px 8px 0 0;
    }}
    div[data-testid="stSidebar"] {{
        direction: rtl;
        text-align: right;
    }}
</style>
""", unsafe_allow_html=True)

# الهيدر الرئيسي
st.markdown(f"""
<div class="main-header">
    <div>
        <h1>🏦 البنك الأهلي السعودي - SNB</h1>
        <p>المنصة الشاملة لتحليل المحافظ، الوعود القائمة والمكسورة، تقارير الإهمال، وتقييم النشاط (60-NPL)</p>
    </div>
    <div style="text-align: left;">
        <span style="background: rgba(255,255,255,0.2); padding: 6px 14px; border-radius: 20px; font-size: 13px; font-weight: bold;">
            تاريخ التقرير: {ref_date.strftime('%Y-%m-%d')}
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 2. دوال مساعدة لتحليل وتصنيف البيانات
# ==========================================
def find_col(df, keywords):
    """البحث الذكي عن اسم العمود بناءً على الكلمات المفتاحية"""
    for col in df.columns:
        col_str = str(col).strip().lower()
        for kw in keywords:
            if kw.lower() in col_str:
                return col
    return None

def parse_duration_to_minutes(val):
    """تحويل مدة المكالمة (سواء نص 00:05:30 أو رقم) إلى دقائق"""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    try:
        parts = val_str.split(':')
        if len(parts) == 3:
            return int(parts[0]) * 60 + int(parts[1]) + int(parts[2]) / 60.0
        elif len(parts) == 2:
            return int(parts[0]) + int(parts[1]) / 60.0
        else:
            return float(val_str)
    except:
        return 0.0

def render_kpi_card(title, value, subtitle="", border_color=snb_green):
    """رسم كارت مؤشر الأداء بتصميم موحد"""
    st.markdown(f"""
    <div class="kpi-card" style="border-right-color: {border_color};">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)

def convert_df_to_excel(dfs_dict):
    """تحويل مجموعة جداول إلى ملف إكسيل واحد متعدد الشيتات للتحميل"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            if not df.empty:
                safe_name = str(sheet_name)[:31].replace('/', '_').replace('\\', '_')
                df.to_excel(writer, index=False, sheet_name=safe_name)
    return output.getvalue()

# ==========================================
# 3. قراءة وفرز الملفات المرفوعة
# ==========================================
if not uploaded_files:
    st.info("👋 أهلاً بكِ يا رنا! قومي برفع ملفات الإكسيل (الوعود، الإهمال، النشاط، أو أي تقارير أخرى) من القائمة الجانبية لعرض المؤشرات والتقارير فوراً.")
    st.stop()

promises_list = []
activity_list = []
neglect_explicit_list = []
other_reports = {}

for file in uploaded_files:
    try:
        if file.name.endswith('.csv'):
            sheets_dict = {"Sheet1": pd.read_csv(file)}
        else:
            excel_file = pd.ExcelFile(file)
            sheets_dict = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
            
        for sheet_name, df in sheets_dict.items():
            if df.empty:
                continue
            # تنظيف أسماء الأعمدة
            df.columns = df.columns.astype(str).str.strip()
            df['المصدر (الملف/الشيت)'] = f"{file.name} - {sheet_name}"
            
            cols_joined = " ".join(df.columns).lower()
            name_joined = f"{file.name} {sheet_name}".lower()
            
            # 1. التحقق هل هو ملف/شيت نشاط ومكالمات؟
            if any(k in cols_joined for k in ['مدة المكالمة', 'مدة الاتصال', 'call duration', 'وقت المكالمة', 'التقيمي']) or \
               any(k in name_joined for k in ['نشاط', 'activity', 'مكالمات', 'calls', 'تقييم']):
                activity_list.append(df)
                
            # 2. التحقق هل هو شيت إهمال منفصل جاهز؟
            elif any(k in name_joined for k in ['إهمال', 'اهمال', 'neglect']) and not any(k in name_joined for k in ['وعود', 'سجل']):
                neglect_explicit_list.append(df)
                
            # 3. التحقق هل هو ملف/شيت محفظة ووعود؟
            elif any(k in cols_joined for k in ['تاريخ الوعد', 'مبلغ الوعد', 'حالة الوعد', 'الوعود', 'حالة المعالجة', 'تاريخ المتابعة', 'promise']) or \
                 any(k in name_joined for k in ['وعود', 'promise', 'npl', 'محفظة']):
                promises_list.append(df)
                
            # 4. أي تقرير إضافي آخر
            else:
                other_reports[f"{file.name} ({sheet_name})"] = df
                
    except Exception as e:
        st.error(f"⚠️ حدث خطأ أثناء قراءة الملف {file.name}: {e}")

df_promises_all = pd.concat(promises_list, ignore_index=True) if promises_list else pd.DataFrame()
df_activity_all = pd.concat(activity_list, ignore_index=True) if activity_list else pd.DataFrame()
df_neglect_ready = pd.concat(neglect_explicit_list, ignore_index=True) if neglect_explicit_list else pd.DataFrame()

# ==========================================
# 4. معالجة منطق الوعود (القائمة والمكسورة) والإهمال
# ==========================================
df_active_promises = pd.DataFrame()
df_broken_promises = pd.DataFrame()
df_neglect_extracted = pd.DataFrame()

collector_col_p = None
amount_col_p = None
product_col_p = None

if not df_promises_all.empty:
    collector_col_p = find_col(df_promises_all, ['المحصل', 'اسم المحصل', 'الموظف', 'collector', 'agent'])
    promise_date_col = find_col(df_promises_all, ['تاريخ الوعد', 'وعد الدفع', 'تاريخ السداد', 'promise date'])
    promise_status_col = find_col(df_promises_all, ['حالة الوعد', 'الحالة', 'حالة المعالجة', 'التصنيف', 'status', 'النتيجة'])
    amount_col_p = find_col(df_promises_all, ['مبلغ الوعد', 'قيمة الوعد', 'المبلغ', 'المديونية', 'الرصيد', 'amount', 'balance'])
    last_action_col = find_col(df_promises_all, ['تاريخ آخر إجراء', 'اخر اجراء', 'آخر متابعة', 'تاريخ المتابعة', 'تاريخ الاتصال', 'last action'])
    product_col_p = find_col(df_promises_all, ['المنتج', 'نوع المنتج', 'المرحلة', 'الفرع', 'product', 'bucket', 'stage'])

    # تحويل التواريخ
    if promise_date_col:
        df_promises_all['_p_date'] = pd.to_datetime(df_promises_all[promise_date_col], errors='coerce').dt.date
    else:
        df_promises_all['_p_date'] = pd.NaT

    if last_action_col:
        df_promises_all['_last_date'] = pd.to_datetime(df_promises_all[last_action_col], errors='coerce').dt.date
    else:
        df_promises_all['_last_date'] = pd.NaT

    # 1. تصنيف الوعود القائمة والمكسورة
    status_series = df_promises_all[promise_status_col].astype(str) if promise_status_col else pd.Series([""] * len(df_promises_all))
    
    # شرط الوعد القائم: مكتوب صراحة "قائم/ساري" أو يوجد تاريخ وعد في المستقبل أو اليوم
    cond_explicit_active = status_series.str.contains('قائم|قائمه|ساري|انتظار|pending|active', case=False, na=False)
    cond_explicit_broken = status_series.str.contains('مكسور|اخلاف|إخلاف|نكث|broken', case=False, na=False)
    
    has_p_date = df_promises_all['_p_date'].notna()
    cond_date_active = has_p_date & (df_promises_all['_p_date'] >= ref_date) & (~cond_explicit_broken)
    cond_date_broken = has_p_date & (df_promises_all['_p_date'] < ref_date) & (~status_series.str.contains('تم السداد|مسدد|مدفوع|paid', case=False, na=False))

    df_active_promises = df_promises_all[cond_explicit_active | cond_date_active].copy()
    df_broken_promises = df_promises_all[cond_explicit_broken | cond_date_broken].copy()

    # 2. استخراج الإهمال من المحفظة (الحسابات بدون متابعة أو المتأخرة عن المتابعة)
    cond_explicit_neglect = status_series.str.contains('إهمال|اهمال|بدون إجراء|لم يتم التواصل|neglect', case=False, na=False)
    if last_action_col:
        days_diff = pd.Series([(ref_date - d).days if pd.notna(d) else 999 for d in df_promises_all['_last_date']], index=df_promises_all.index)
        cond_date_neglect = (df_promises_all['_last_date'].isna()) | (days_diff > 5)
        df_neglect_extracted = df_promises_all[cond_explicit_neglect | cond_date_neglect].copy()
    else:
        df_neglect_extracted = df_promises_all[cond_explicit_neglect].copy()

    # تنظيف الأعمدة المؤقتة
    for temp_df in [df_promises_all, df_active_promises, df_broken_promises, df_neglect_extracted]:
        for c in ['_p_date', '_last_date']:
            if c in temp_df.columns:
                temp_df.drop(columns=[c], inplace=True)

# دمج الإهمال المستخرج مع شيتات الإهمال المنفصلة إن وجدت
df_neglect_final = pd.concat([df_neglect_ready, df_neglect_extracted], ignore_index=True) if (not df_neglect_ready.empty or not df_neglect_extracted.empty) else pd.DataFrame()

# ==========================================
# 5. معالجة منطق النشاط وتقييم المكالمات
# ==========================================
collector_col_a = None
if not df_activity_all.empty:
    collector_col_a = find_col(df_activity_all, ['المحصل', 'اسم المحصل', 'الموظف', 'collector', 'agent'])
    dur_col = find_col(df_activity_all, ['مدة المكالمة', 'مدة الاتصال', 'المدة', 'duration', 'وقت المكالمة'])
    call_res_col = find_col(df_activity_all, ['نتيجة المكالمة', 'حالة المكالمة', 'الحالة', 'النتيجة', 'التقيمي', 'result', 'status'])

    if dur_col:
        df_activity_all['مدة_المكالمة_بالدقائق'] = df_activity_all[dur_col].apply(parse_duration_to_minutes).round(2)
    else:
        df_activity_all['مدة_المكالمة_بالدقائق'] = 0.0

    # حساب الوقت الضائع (الناجحة > 20 دقيقة، غير الناجحة > 5 دقائق)
    def calc_wasted_time(row):
        dur = row['مدة_المكالمة_بالدقائق']
        res = str(row[call_res_col]).lower() if call_res_col else ""
        is_success = any(w in res for w in ['ناجح', 'تم الرد', 'تجاوب', 'وعد', 'سداد', 'answered', 'success', 'connected'])
        if is_success:
            return max(0.0, dur - 20.0)
        else:
            return max(0.0, dur - 5.0)

    df_activity_all['الوقت_الضائع_بالدقائق'] = df_activity_all.apply(calc_wasted_time, axis=1).round(2)

# ==========================================
# 6. فلاتر التصفية الشاملة في القائمة الجانبية
# ==========================================
with st.sidebar:
    st.markdown("---")
    st.markdown("### 🔍 فلاتر التصفية")
    
    # جمع أسماء المحصلين من كل الملفات
    all_collectors = set()
    if collector_col_p and collector_col_p in df_promises_all.columns:
        all_collectors.update(df_promises_all[collector_col_p].dropna().astype(str).unique())
    if collector_col_a and collector_col_a in df_activity_all.columns:
        all_collectors.update(df_activity_all[collector_col_a].dropna().astype(str).unique())
        
    selected_collectors = st.multiselect("👤 تصفية حسب المحصل:", sorted(list(all_collectors)))
    
    # تطبيق الفلتر على الجداول
    if selected_collectors:
        if collector_col_p:
            df_promises_all = df_promises_all[df_promises_all[collector_col_p].astype(str).isin(selected_collectors)]
            if not df_active_promises.empty:
                df_active_promises = df_active_promises[df_active_promises[collector_col_p].astype(str).isin(selected_collectors)]
            if not df_broken_promises.empty:
                df_broken_promises = df_broken_promises[df_broken_promises[collector_col_p].astype(str).isin(selected_collectors)]
        if not df_neglect_final.empty:
            col_neg = find_col(df_neglect_final, ['المحصل', 'اسم المحصل', 'الموظف', 'collector'])
            if col_neg:
                df_neglect_final = df_neglect_final[df_neglect_final[col_neg].astype(str).isin(selected_collectors)]
        if collector_col_a and not df_activity_all.empty:
            df_activity_all = df_activity_all[df_activity_all[collector_col_a].astype(str).isin(selected_collectors)]

# ==========================================
# 7. عرض لوحة مؤشرات الأداء الشاملة (KPIs)
# ==========================================
st.markdown("### 📊 الملخص التنفيذي لمؤشرات الأداء (KPIs)")

k1, k2, k3, k4, k5, k6 = st.columns(6)

# حساب الأرقام
total_portfolio = len(df_promises_all)
active_count = len(df_active_promises)
broken_count = len(df_broken_promises)
neglect_count = len(df_neglect_final)
total_calls = len(df_activity_all)
total_wasted_mins = df_activity_all['الوقت_الضائع_بالدقائق'].sum() if not df_activity_all.empty else 0.0

active_amt = pd.to_numeric(df_active_promises[amount_col_p], errors='coerce').sum() if (not df_active_promises.empty and amount_col_p) else 0
broken_amt = pd.to_numeric(df_broken_promises[amount_col_p], errors='coerce').sum() if (not df_broken_promises.empty and amount_col_p) else 0

with k1:
    render_kpi_card("📁 إجمالي سجلات المحفظة", f"{total_portfolio:,}", f"من {len(uploaded_files)} ملف مرفوع", snb_green)
with k2:
    sub_act = f"بقيمة: {active_amt:,.0f}" if active_amt > 0 else "وعود قيد المتابعة"
    render_kpi_card("🤝 الوعود القائمة", f"{active_count:,}", sub_act, "#10B981")
with k3:
    sub_brk = f"بقيمة: {broken_amt:,.0f}" if broken_amt > 0 else "تجاوزت تاريخ السداد"
    render_kpi_card("💔 الوعود المكسورة", f"{broken_count:,}", sub_brk, "#EF4444")
with k4:
    neg_rate = f"بنسبة {(neglect_count/total_portfolio*100):.1f}% من المحفظة" if total_portfolio > 0 else "حسابات متأخرة"
    render_kpi_card("⚠️ حالات الإهمال", f"{neglect_count:,}", neg_rate, "#F59E0B")
with k5:
    render_kpi_card("📞 إجمالي النشاط والمكالمات", f"{total_calls:,}", "إجمالي الحركات المسجلة", "#3B82F6")
with k6:
    render_kpi_card("⏱️ إجمالي الوقت الضائع", f"{total_wasted_mins:,.1f} د", f"يعادل {(total_wasted_mins/60):.1f} ساعة", "#8B5CF6")

# ==========================================
# 8. التبويبات التفصيلية والرسوم البيانية
# ==========================================
tab_overview, tab_promises, tab_neglect, tab_activity, tab_others = st.tabs([
    "📈 التحليل البياني الشامل",
    "🤝 الوعود القائمة والمكسورة",
    "⚠️ تقرير الإهمال",
    "📞 تقييم النشاط والوقت الضائع",
    f"🗂️ تقارير إضافية ({len(other_reports)})"
])

# --- تبويب 1: التحليل البياني الشامل ---
with tab_overview:
    c_chart1, c_chart2 = st.columns(2)
    
    with c_chart1:
        st.markdown("#### 📊 توزيع حالات الوعود والإهمال")
        if (active_count + broken_count + neglect_count) > 0:
            fig_pie = px.pie(
                names=["الوعود القائمة", "الوعود المكسورة", "حالات الإهمال"],
                values=[active_count, broken_count, neglect_count],
                color_discrete_sequence=[snb_green, "#EF4444", "#F59E0B"],
                hole=0.45
            )
            fig_pie.update_layout(font=dict(family="Cairo", size=13), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("لا توجد بيانات وعود أو إهمال كافية لعرض الرسم البياني.")

    with c_chart2:
        st.markdown("#### 👤 أداء المحصلين (الوعود القائمة مقابل المكسورة)")
        if collector_col_p and (not df_active_promises.empty or not df_broken_promises.empty):
            s_act = df_active_promises[collector_col_p].value_counts().rename("وعود قائمة")
            s_brk = df_broken_promises[collector_col_p].value_counts().rename("وعود مكسورة")
            df_col_summary = pd.concat([s_act, s_brk], axis=1).fillna(0).reset_index()
            df_col_summary.rename(columns={'index': 'المحصل', collector_col_p: 'المحصل'}, inplace=True)
            
            fig_bar = px.bar(
                df_col_summary,
                x='المحصل',
                y=['وعود قائمة', 'وعود مكسورة'],
                barmode='group',
                color_discrete_map={'وعود قائمة': snb_green, 'وعود مكسورة': '#EF4444'}
            )
            fig_bar.update_layout(font=dict(family="Cairo", size=12), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("ارفعي ملف الوعود متضمناً عمود 'المحصل' لعرض المقارنة.")

# --- تبويب 2: الوعود القائمة والمكسورة ---
with tab_promises:
    sub_t1, sub_t2, sub_t3 = st.tabs([
        f"✅ الوعود القائمة ({len(df_active_promises)})",
        f"❌ الوعود المكسورة ({len(df_broken_promises)})",
        f"📋 كافة سجلات المحفظة ({len(df_promises_all)})"
    ])
    with sub_t1:
        if not df_active_promises.empty:
            st.dataframe(df_active_promises, use_container_width=True)
        else:
            st.info("لا توجد وعود قائمة حالياً.")
    with sub_t2:
        if not df_broken_promises.empty:
            st.dataframe(df_broken_promises, use_container_width=True)
        else:
            st.info("لا توجد وعود مكسورة حالياً.")
    with sub_t3:
        if not df_promises_all.empty:
            st.dataframe(df_promises_all, use_container_width=True)
        else:
            st.info("لم يتم رفع ملفات المحفظة والوعود.")

# --- تبويب 3: تقرير الإهمال ---
with tab_neglect:
    if not df_neglect_final.empty:
        col_neg_name = find_col(df_neglect_final, ['المحصل', 'اسم المحصل', 'الموظف', 'collector'])
        if col_neg_name:
            st.markdown("#### ⚠️ توزيع حالات الإهمال حسب المحصل")
            neg_by_col = df_neglect_final[col_neg_name].value_counts().reset_index()
            neg_by_col.columns = ['المحصل', 'عدد حالات الإهمال']
            fig_neg = px.bar(neg_by_col, x='المحصل', y='عدد حالات الإهمال', color_discrete_sequence=['#F59E0B'])
            fig_neg.update_layout(font=dict(family="Cairo"), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_neg, use_container_width=True)
            
        st.markdown("#### 📋 جدول حالات الإهمال التفصيلي")
        st.dataframe(df_neglect_final, use_container_width=True)
    else:
        st.info("لا توجد حالات إهمال مسجلة في الملفات المرفوعة.")

# --- تبويب 4: تقييم النشاط والوقت الضائع ---
with tab_activity:
    if not df_activity_all.empty:
        if collector_col_a:
            st.markdown("#### 📞 ملخص نشاط المحصلين والوقت الضائع")
            act_summary = df_activity_all.groupby(collector_col_a).agg(
                إجمالي_المكالمات=(collector_col_a, 'count'),
                إجمالي_الدقائق=('مدة_المكالمة_بالدقائق', 'sum'),
                الوقت_الضائع_بالدقائق=('الوقت_الضائع_بالدقائق', 'sum')
            ).reset_index()
            act_summary['إجمالي_الدقائق'] = act_summary['إجمالي_الدقائق'].round(1)
            act_summary['الوقت_الضائع_بالدقائق'] = act_summary['الوقت_الضائع_بالدقائق'].round(1)
            st.dataframe(act_summary, use_container_width=True)
            
        st.markdown("#### 📋 سجل النشاط التفصيلي")
        st.dataframe(df_activity_all, use_container_width=True)
    else:
        st.info("لم يتم رفع ملف النشاط والمكالمات بعد. ارفعي ملف النشاط لعرض التقييم والوقت الضائع تلقائياً.")

# --- تبويب 5: أي تقارير إضافية أخرى ---
with tab_others:
    if other_reports:
        for rep_name, rep_df in other_reports.items():
            with st.expander(f"📄 تقرير: {rep_name} ({len(rep_df)} سجل)", expanded=True):
                st.dataframe(rep_df, use_container_width=True)
    else:
        st.info("أي ملف إكسيل إضافي ترفعينه لا يندرج تحت (الوعود أو النشاط أو الإهمال) سيظهر هنا تلقائياً بجدوله الكامل.")

# ==========================================
# 9. تصدير وتحميل التقرير المجمع النهائي
# ==========================================
st.markdown("---")
st.markdown("### 📥 تصدير التقرير الشامل الموحد (Excel)")

export_sheets = {
    "الوعود القائمة": df_active_promises,
    "الوعود المكسورة": df_broken_promises,
    "تقرير الإهمال": df_neglect_final,
    "تقرير النشاط": df_activity_all,
    "إجمالي المحفظة": df_promises_all
}
for k, v in other_reports.items():
    export_sheets[k] = v

excel_bytes = convert_df_to_excel(export_sheets)
st.download_button(
    label="⬇️ تحميل جميع التقارير المفلترة في ملف إكسيل واحد (SNB Report)",
    data=excel_bytes,
    file_name=f"تقرير_الأداء_والوعود_SNB_{ref_date.strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
