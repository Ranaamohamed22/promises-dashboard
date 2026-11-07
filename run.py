import pandas as pd
from datetime import date
import os
import glob

# تحديد مسار الفولدر الرئيسي تلقائياً عشان الكود يشتغل صح من أي مكان
base_dir = os.path.dirname(os.path.abspath(__file__))
input_folder = os.path.join(base_dir, 'input')
output_folder = os.path.join(base_dir, 'output')

def process_daily_sheets():
    # التأكد من وجود الفولدرات
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    
    # البحث عن أحدث ملف إكسيل داخل فولدر input (وتجاهل أي ملفات مؤقتة مفتوحة)
    excel_files = [
        f for f in glob.glob(os.path.join(input_folder, "*.xlsx")) 
        if not os.path.basename(f).startswith('~$')
    ]
    
    if not excel_files:
        print("مفيش أي ملف إكسيل موجود جوه فولدر input! حطي الشيت وجربي تاني.")
        return
        
    latest_file = max(excel_files, key=os.path.getctime)
    print(f"جاري العمل على الملف: {os.path.basename(latest_file)}")
    
    # 1. قراءة الشيت الأساسي وتنظيف أسماء العواميد من أي مسافات زيادة
    df = pd.read_excel(latest_file)
    df.columns = df.columns.str.strip()
    
    # تحديد تاريخ اليوم
    today = pd.to_datetime(date.today()).date()
    
    # تحويل عواميد التواريخ
    due_date = pd.to_datetime(df['Follow up Due Date'], errors='coerce').dt.date
    last_date = pd.to_datetime(df['Follow up Last Date'], errors='coerce').dt.date
    
    # ==========================================
    # الشروط المشتركة بين الشيتين
    # ==========================================
    base_condition = (
        (df['حالة المعالجة - التمويل'].astype(str).str.strip() == 'غير معالج') &
        (df['تحديث فئة العميل - التمويل'].astype(str).str.strip().str.lower() != 'writeoff') &
        (df['Classification'].astype(str).str.strip() == 'وعود السداد قيد التنفيذ')
    )
    
    # ==========================================
    # أولاً: استخراج "الوعود القائمة"
    # ==========================================
    active_promises_df = df[
        base_condition &
        (due_date == today) &
        (last_date != today)
    ].copy()
    
    # ==========================================
    # ثانياً: استخراج "الوعود المكسورة"
    # ==========================================
    broken_promises_df = df[
        base_condition &
        (due_date < today)
    ].copy()
    
    broken_due = pd.to_datetime(broken_promises_df['Follow up Due Date'], errors='coerce').dt.normalize()
    broken_last = pd.to_datetime(broken_promises_df['Follow up Last Date'], errors='coerce').dt.normalize()
    
    # طرح Last Date من Due Date والفلترة على الأرقام السالبة فقط (أقل من صفر)
    broken_promises_df['Date_Diff'] = (broken_last - broken_due).dt.days
    broken_promises_df = broken_promises_df[broken_promises_df['Date_Diff'] < 0]
    
    # ==========================================
    # حفظ النتيجة في فولدر output بشيتين منفصلين
    # ==========================================
    output_file = os.path.join(output_folder, f"Daily_Promises_Report_{today}.xlsx")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        active_promises_df.to_excel(writer, sheet_name='الوعود القائمة', index=False)
        broken_promises_df.to_excel(writer, sheet_name='الوعود المكسورة', index=False)
        
    print(f"تم تجهيز الشيت بنجاح وحفظه في: {output_file}")
    print(f"- عدد الوعود القائمة: {len(active_promises_df)}")
    print(f"- عدد الوعود المكسورة: {len(broken_promises_df)}")

if __name__ == '__main__':
    process_daily_sheets()