# =====================================================
# ① 出貨類型筆數統計（含銅重量換算噸；「銅重量(噸)」欄位實際單位=kg）
# =====================================================
st.subheader("① 出貨類型筆數統計")
if ship_type_col in data.columns:
    # 筆數
    type_counts = (
        data[ship_type_col]
        .fillna("(空白)")
        .value_counts(dropna=False)
        .rename_axis("出貨類型")
        .reset_index(name="筆數")
    )
    total_rows = int(type_counts["筆數"].sum())

    # 銅重量合計（原欄位為 kg 數值）→ 轉成「噸」顯示
    if copper_ton_col in data.columns:
        data["_copper_kg"] = pd.to_numeric(data[copper_ton_col], errors="coerce")
        copper_sum_kg = (
            data.groupby(ship_type_col)["_copper_kg"]
            .sum(min_count=1)
            .reset_index()
            .rename(columns={"_copper_kg": "銅重量(kg)合計"})
        )
        merged = type_counts.merge(copper_sum_kg, how="left",
                                   left_on="出貨類型", right_on=ship_type_col)
        if ship_type_col in merged.columns:
            merged.drop(columns=[ship_type_col], inplace=True)

        # 換算成「噸」的欄位（四捨五入到小數 2 位）
        merged["銅重量(噸)合計"] = (pd.to_numeric(merged["銅重量(kg)合計"], errors="coerce") / 1000.0).round(2)
        merged.drop(columns=["銅重量(kg)合計"], inplace=True)
    else:
        merged = type_counts.copy()
        merged["銅重量(噸)合計"] = None

    # —— KPI 卡片（就是 st.metric）——
    # KPI 卡片會顯示在這一段區塊上方
    k1, k2 = st.columns(2)
    k1.metric("資料筆數", f"{total_rows:,}")
    total_copper_ton_all = (
        pd.to_numeric(merged["銅重量(噸)合計"], errors="coerce").sum()
        if "銅重量(噸)合計" in merged.columns else 0.0
    )
    k2.metric("總銅重量(噸)", f"{total_copper_ton_all:,.2f}")

    # —— 圖表選擇：長條 / 圓餅 / 環形（都顯示百分比）——
    chart_choice = st.radio("圖表類型", ["長條圖", "圓餅圖", "環形圖"], horizontal=True)
    neutral_colors = ["#7f8c8d", "#95a5a6", "#bdc3c7", "#dfe6e9", "#b2bec3", "#636e72", "#ced6e0"]

    c1, c2 = st.columns([1, 1])
    with c1:
        st.dataframe(merged, use_container_width=True)
        st.download_button(
            "下載出貨類型統計 CSV",
            data=merged.to_csv(index=False).encode("utf-8-sig"),
            file_name="出貨類型_統計.csv",
            mime="text/csv",
        )
    with c2:
        if not type_counts.empty:
            if chart_choice == "長條圖":
                fig = px.bar(type_counts, x="出貨類型", y="筆數")
            else:
                # 圓餅 / 環形：都顯示百分比標籤
                hole_size = 0.0 if chart_choice == "圓餅圖" else 0.5
                fig = px.pie(
                    type_counts,
                    names="出貨類型",
                    values="筆數",
                    color_discrete_sequence=neutral_colors,
                    hole=hole_size,
                )
                fig.update_traces(textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)
else:
    st.warning("找不到『出貨申請類型』欄位，請確認上傳檔案欄名。")
