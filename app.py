                        <h3 style="margin:0; color: #586F8C; font-weight:900;">{top_calls:,} مكالمة</h3>
                        <p style="margin: 0; color: #FFFFFF; font-size: 14px; font-weight: bold;">استحوذ على <span style="color:#99AABF;">{top_percentage:.1f}%</span> من إجمالي شغل الفترة</p>
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
            pivot_summary = (
                act_df.groupby([collector_col, "Call_Status"])
                .size()
                .unstack(fill_value=0)
            )

            if "ناجحة" not in pivot_summary.columns:
                pivot_summary["ناجحة"] = 0

            if "مغطاة" not in pivot_summary.columns:
                pivot_summary["مغطاة"] = 0

            pivot_summary["إجمالي المحاولات"] = (
                pivot_summary["ناجحة"] + pivot_summary["مغطاة"]
            )

            pivot_summary["نسبة النجاح %"] = (
                pivot_summary["ناجحة"]
                .div(pivot_summary["إجمالي المحاولات"].replace(0, pd.NA))
                .mul(100)
                .fillna(0)
                .round(1)
                .astype(str)
                + "%"
            )

            pivot_summary = pivot_summary.sort_values(
                by="إجمالي المحاولات",
                ascending=False
            ).reset_index()

            with st.expander(
                "📊 تفصيل أداء المحصلين (ناجحة مقابل مغطاة)",
                expanded=True
            ):
                st.dataframe(
                    pivot_summary,
                    use_container_width=True,
                    hide_index=True
                )

        extra_time_cols = [
            col for col in ["Date", "Time", "Call_Status"]
            if col in act_df.columns
        ]

        display_cols = extra_time_cols + [
            col for col in act_df.columns
            if col not in [
                "Created_DateTime", "Date", "Time", "Call_Status"
            ]
        ]

        st.dataframe(
            act_df[display_cols],
            use_container_width=True,
            height=450
        )

        act_buffer = io.BytesIO()
        act_df[display_cols].to_excel(
            act_buffer,
            index=False,
            sheet_name="تقرير النشاط والتقييم"
        )
        act_buffer.seek(0)

        st.download_button(
            "📥 تحميل تقرير النشاط مفلتر.xlsx",
            data=act_buffer,
            file_name="Activity_Success_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==========================================
# زر مسح ملفات الصفحة الحالية وإعادة الرفع
# ==========================================
if st.session_state.page_files.get(curr_page) is not None:
    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🗑️ مسح ملفات الصفحة الحالية ورفع ملفات جديدة", key=f"clear_{curr_page}"):
        st.session_state.page_files[curr_page] = None

        # إعادة تهيئة مفتاح رفع الملف للصفحة الحالية
        uploader_key = f"uploader_{curr_page}"
        if uploader_key in st.session_state:
            del st.session_state[uploader_key]

        st.rerun()
