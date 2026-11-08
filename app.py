marker_color="#99AABF",
                            text=fl_df['first_time_12h'],
                            textposition='outside',
                            textfont=dict(color="#1F4373", weight='bold', size=12),
                            cliponaxis=False,
                            customdata=list(zip(fl_df['first_time_12h'], first_note_txt)),
                            hovertemplate="<b>%{y}</b><br>أول إفادة: %{customdata[0]}<br>الإفادة: %{customdata[1]}<extra></extra>"
                        ))
                        fig_fl.add_trace(go.Bar(
                            name="آخر إفادة",
                            y=fl_df[collector_col],
                            x=fl_df['last_hour'],
                            orientation='h',
                            marker_color="#1F4373",
                            text=fl_df['last_time_12h'],
                            textposition='outside',
                            textfont=dict(color="#1F4373", weight='bold', size=12),
                            cliponaxis=False,
                            customdata=list(zip(fl_df['last_time_12h'], last_note_txt)),
                            hovertemplate="<b>%{y}</b><br>آخر إفادة: %{customdata[0]}<br>الإفادة: %{customdata[1]}<extra></extra>"
                        ))
                        fig_fl.update_layout(
                            barmode='group',
                            title=dict(text="🕒 ميعاد أول وآخر إفادة", x=0.5, xanchor='center', font=dict(color="#1F4373", size=19, weight='bold')),
                            paper_bgcolor="#FFFFFF",
                            plot_bgcolor="#F2F2F2",
                            font=dict(color="#1F4373", family="Tahoma", size=12),
                            height=max(520, len(fl_df) * 58 + 170),
                            margin=dict(l=20, r=65, t=65, b=45),
                            xaxis=dict(
                                title="الوقت (ساعات ودقائق - نظام 12 ساعة)",
                                range=[0, 27],
                                tickvals=[0, 4, 8, 12, 16, 20, 24],
                                ticktext=["12:00 ص", "4:00 ص", "8:00 ص", "12:00 م", "4:00 م", "8:00 م", "12:00 ص"],
                                gridcolor="#FFFFFF"
                            ),
                            yaxis=dict(title="", automargin=True),
                            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="center", x=0.5)
                        )
                        st.plotly_chart(fig_fl, use_container_width=True, config={'displayModeBar': False})
                    else:
                        st.info("لا يوجد عمود تاريخ/وقت لحساب أول وآخر إفادة.")

            # 3) الشارت الثالثة: ميعاد أول إفادة لكل محصل بعد البريك (بعد 12:30 م) بنظام الـ 12 ساعة
            with ch_col3:
                with st.container(border=True):
                    if 'Created_DateTime' in act_df.columns and act_df['Created_DateTime'].notna().any():
                        after_break_df = act_df[act_df['Created_DateTime'].dt.time > time(12, 30)].dropna(subset=['Created_DateTime', collector_col]).sort_values('Created_DateTime')
                        if not after_break_df.empty:
                            first_ab = after_break_df.groupby(collector_col, as_index=False).first()
                            top_names_list = act_df[collector_col].value_counts().head(10).index.tolist()
                            first_ab = first_ab[first_ab[collector_col].isin(top_names_list)].copy()
                            first_ab['ab_hour'] = first_ab['Created_DateTime'].dt.hour + first_ab['Created_DateTime'].dt.minute / 60.0
                            first_ab['ab_time_12h'] = format_12h(first_ab['Created_DateTime'])
                            first_ab['ab_dt_12h'] = format_dt_12h(first_ab['Created_DateTime'])
                            ab_note_txt = first_ab[notes_col_name].astype(str) if notes_col_name and notes_col_name in first_ab.columns else pd.Series(["-"] * len(first_ab))

                            ab_colors = [CHART_PALETTE[i % 4] for i in range(len(first_ab))]
                            fig_ab = go.Figure(go.Bar(
                                y=first_ab[collector_col],
                                x=first_ab['ab_hour'],
                                orientation='h',
                                marker=dict(color=ab_colors, line=dict(color="#1F4373", width=1)),
                                text=first_ab['ab_time_12h'],
                                textposition='auto',
                                textfont=dict(color="#FFFFFF", weight='bold', size=12),
                                customdata=list(zip(first_ab['ab_dt_12h'], ab_note_txt)),
                                hovertemplate="<b>%{y}</b><br>ميعاد أول إفادة بعد البريك: %{customdata[0]}<br>الإفادة (Notes): %{customdata[1]}<extra></extra>"
                            ))
                            fig_ab.update_layout(
                                title=dict(text="☕ ميعاد أول إفادة بعد البريك (12:30 م)", x=0.5, xanchor='center', font=dict(color="#1F4373", size=15, weight='bold')),
                                paper_bgcolor="#FFFFFF",
                                plot_bgcolor="#F2F2F2",
                                font=dict(color="#1F4373", family="Tahoma"),
                                height=360,
                                margin=dict(l=10, r=10, t=45, b=20),
                                xaxis=dict(
                                    title="ميعاد أول إفادة (نظام 12 ساعة)",
                                    range=[12, 24],
                                    tickvals=[12, 14, 16, 18, 20, 22, 24],
                                    ticktext=["12 م", "2 م", "4 م", "6 م", "8 م", "10 م", "12 ص"],
                                    gridcolor="#FFFFFF"
                                ),
                                yaxis=dict(title="", automargin=True)
                            )
                            st.plotly_chart(fig_ab, use_container_width=True, config={'displayModeBar': False})
                        else:
                            st.info("لا توجد إفادات مسجلة بعد وقت البريك (12:30 م).")
                    else:
                        st.info("لا يوجد عمود تاريخ/وقت لعرض أول إفادة بعد البريك.")

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
