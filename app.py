                unmapped_notes = [n for n in all_sheet_notes if n not in mapping_dict]
                
                if unmapped_notes:
                    st.warning(f"⚠️ وجدنا ملاحظات جديدة ({len(unmapped_notes)} ملاحظة) غير موجودة في موديل التقييم المرفوع:")
                    
                    if 'custom_note_evals' not in st.session_state:
                        st.session_state.custom_note_evals = {}
                        
                    for un_note in unmapped_notes:
                        eval_choice = st.radio(f"كيف تحسبين الملاحظة الآتية: **({un_note})**؟", options=["ناجحة (1)", "مغطاة (0)"], key=f"eval_radio_{un_note}")
                        st.session_state.custom_note_evals[un_note] = 1 if "ناجحة" in eval_choice else 0
                        
                def classify_with_model(val):
                    v_str = str(val).strip()
                    if v_str in mapping_dict:
                        val_eval = str(mapping_dict[v_str]).strip()
                        return 'ناجحة' if val_eval in ['1', '1.0', '1,0', 'true'] else 'مغطاة'
                    elif v_str in st.session_state.get('custom_note_evals', {}):
                        return 'ناجحة' if st.session_state.custom_note_evals[v_str] == 1 else 'مغطاة'
                    else:
                        return 'غير مصنف'
                        
                act_df['Call_Status'] = act_df[notes_col_name].apply(classify_with_model)
            else:
                act_df['Call_Status'] = 'غير مصنف'
        else:
            target_col = notes_col_name if notes_col_name else find_col_in(act_df, ['التقييم', 'status', 'الحالة', 'sub state', 'result'])
            if target_col:
                available_vals = act_df[target_col].dropna().astype(str).unique().tolist()
                selected_success_vals = st.multiselect(f"حددي القيم التي تعتبر (ناجحة = 1) من عمود [{target_col}]:", options=available_vals, default=[v for v in available_vals if str(v).strip() in ['1', '1.0', 'ناجح', 'وعد', 'سداد', 'تم']])
                act_df['Call_Status'] = act_df[target_col].astype(str).apply(lambda x: 'ناجحة' if x in selected_success_vals else 'مغطاة')
            else:
                act_df['Call_Status'] = 'غير مصنف'

        total_activity = len(act_df)
        success_count = len(act_df[act_df['Call_Status'] == 'ناجحة'])
        covered_count = len(act_df[act_df['Call_Status'] == 'مغطاة'])
        success_rate = (success_count / total_activity * 100) if total_activity > 0 else 0

        collector_col = collector_col_name
        if collector_col and collector_col in act_df.columns and total_activity > 0:
            top_collector_series = act_df[collector_col].value_counts()
            if not top_collector_series.empty:
                top_name = top_collector_series.index[0]
                top_calls = top_collector_series.values[0]
                top_pct = (top_calls / total_activity * 100) if total_activity > 0 else 0
                st.markdown(f"""
                <div style="background-color: {card_color}; border-radius: 12px; padding: 20px 30px; border: 2px solid #586F8C; direction:rtl; text-align: right; margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between;">
                    <div>
                        <h4 style="margin:0; color: #586F8C; font-weight:900;">🏆 بطل الشفت (الأكثر نشاطاً)</h4>
                        <p style="margin: 5px 0 0 0; color: {text_color}; font-size: 18px;">المحصل: <b style="color:#1F4373; font-size: 22px;">{top_name}</b></p>
                    </div>
                    <div style="text-align: left; background-color: {bg_color}; padding: 10px 20px; border-radius: 8px; border: 1px solid #586F8C;">
                        <h3 style="margin:0; color: #586F8C; font-weight:900;">{top_calls:,} مكالمة</h3>
                        <p style="margin: 0; color: #FFFFFF; font-size: 14px; font-weight: bold;">استحوذ على <span style="color:#99AABF;">{top_pct:.1f}%</span> من إجمالي شغل الفترة</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown(f"""
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 20px; direction: rtl;">
            <div class="kpi-card" style="border-right-color: {border_color};">
                <div class="kpi-icon">📞</div>
                <div class="kpi-title">إجمالي المحاولات</div>
                <div class="kpi-value">{total_activity:,}</div>
                <div class="kpi-sub" style="color: {sub_text_color};">بعد إزالة التكرار</div>
            </div>
            <div class="kpi-card" style="border-right-color: #1F4373;">
                <div class="kpi-icon">✅</div>
                <div class="kpi-title">المكالمات الناجحة (1)</div>
                <div class="kpi-value">{success_count:,}</div>
                <div class="kpi-sub" style="color: #1F4373;">تقييم ناجح</div>
            </div>
            <div class="kpi-card" style="border-right-color: {border_color};">
                <div class="kpi-icon">🎧</div>
                <div class="kpi-title">المكالمات المغطاة (0)</div>
                <div class="kpi-value">{covered_count:,}</div>
                <div class="kpi-sub" style="color: {sub_text_color};">تقييم مغطاة</div>
            </div>
            <div class="kpi-card" style="border-right-color: #586F8C;">
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

        st.markdown("<div style='text-align: right; direction: rtl;'>", unsafe_allow_html=True)
        extra_time_cols = [c for c in ['Date', 'Time', 'Call_Status'] if c in act_df.columns]
        display_cols = extra_time_cols + [c for c in act_df.columns if c not in ['Created_DateTime', 'Date', 'Time', 'Call_Status']]
        st.dataframe(act_df[display_cols], use_container_width=True, height=450)
        st.markdown("</div>", unsafe_allow_html=True)

        act_buf = io.BytesIO()
        act_df[display_cols].to_excel(act_buf, index=False, sheet_name='تقرير النشاط والتقييم')
        act_buf.seek(0)
        st.download_button("📥 تحميل تقرير النشاط مفلتر.xlsx", data=act_buf, file_name=f"Activity_Success_Report.xlsx")
