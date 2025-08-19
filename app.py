import io
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="TMS出貨配送數據分析", page_icon="📊", layout="wide")
st.title("📊 TMS出貨配送數據分析")
st.caption("上傳 Excel/CSV → 過濾與彙整 → 自訂欄位 → 圖表 → 下載結果｜新增：① 出貨類型筆數 ② 達交率（簽收≤指定到貨）")

# ---------- 檔案上傳 ----------
file = st.file_uploader("上傳 Excel 或 CSV 檔", type=["xlsx", "xls", "csv"], help="最多 200 MB；Excel 需使用 openpyxl 解析")

@st.cache_data
def load_data(file):
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file, engine="openpyxl")
    return df

# 自動猜測欄位工具

def _guess_col(cols, keywords):
    for kw in keywords:
        for c in cols:
            if kw in str(c):
                return c
    return None

# 轉時間（自動解析）

def to_dt(series):
    return pd.to_datetime(series, errors="coerce")

if file:
    df = load_data(file)
    st.subheader("原始資料預覽")
    st.dataframe(df, use_container_width=True)

    # ---------- 側欄操作 ----------
    with st.sidebar:
        st.header("⚙️ 操作區")
        cols = list(df.columns)

        # 基本篩選（可略）
        filter_col = st.selectbox("選擇要篩選的欄位（可略）", ["（不篩選）"] + cols)
        filtered = df.copy()
        if filter_col != "（不篩選）":
            unique_vals = filtered[filter_col].dropna().unique().tolist()
            selected_vals = st.multiselect("選取值", unique_vals)
            if selected_vals:
                filtered = filtered[filtered[filter_col].isin(selected_vals)]

        st.markdown("—")
        st.subheader("欄位對應（建議先確認）")
        ship_type_default = _guess_col(cols, ["出貨申請類型", "出貨類型", "配送類型", "類型"]) or (cols[0] if cols else None)
        due_date_default = _guess_col(cols, ["指定到貨日期", "到貨日期", "指定到貨", "到貨日"]) or (cols[0] if cols else None)
        sign_date_default = _guess_col(cols, ["客戶簽收日期", "簽收日期", "簽收日", "客戶簽收日期/時/分"]) or (cols[0] if cols else None)

        ship_type_col = st.selectbox("出貨類型欄位", options=cols, index=(cols.index(ship_type_default) if ship_type_default in cols else 0))
        due_date_col = st.selectbox("指定到貨日期欄位", options=cols, index=(cols.index(due_date_default) if due_date_default in cols else 0))
        sign_date_col = st.selectbox("客戶簽收日期欄位", options=cols, index=(cols.index(sign_date_default) if sign_date_default in cols else 0))

    data = filtered.copy()

    # ---------- 功能①：出貨類型筆數 ----------
    st.subheader("① 出貨類型筆數統計")
    if ship_type_col in data.columns:
        type_counts = (
            data[ship_type_col]
            .fillna("(空白)")
            .value_counts(dropna=False)
            .rename_axis("出貨類型")
            .reset_index(name="筆數")
        )
        c1, c2 = st.columns([1, 1])
        with c1:
            st.dataframe(type_counts, use_container_width=True)
        with c2:
            fig = px.bar(type_counts, x="出貨類型", y="筆數")
            st.plotly_chart(fig, use_container_width=True)

        st.download_button(
            "下載出貨類型筆數 CSV",
            data=type_counts.to_csv(index=False).encode("utf-8-sig"),
            file_name="出貨類型_筆數統計.csv",
            mime="text/csv",
        )
    else:
        st.warning("找不到出貨類型欄位，請在側欄重新選擇。")

    st.markdown("---")

    # ---------- 功能②：達交率（客戶簽收「日期」 ≤ 指定到貨「日期」，忽略時分秒） ----------
st.subheader("② 達交率（僅比對日期，不含時分秒）")
if due_date_col in data.columns and sign_date_col in data.columns:
    due_dt = to_dt(data[due_date_col])
    sign_dt = to_dt(data[sign_date_col])

    # 僅取日期（把時間歸零）
    due_day = due_dt.dt.normalize()
    sign_day = sign_dt.dt.normalize()

    on_time = sign_day <= due_day
    # 兩個日期皆有值時才判定
    valid_mask = due_day.notna() & sign_day.notna()
    on_time_valid = on_time[valid_mask]
    total_valid = int(valid_mask.sum())
    on_time_count = int(on_time_valid.sum())
    rate = (on_time_count / total_valid * 100.0) if total_valid > 0 else 0.0

    k1, k2, k3 = st.columns(3)
    k1.metric("有效筆數（兩日期皆有值）", f"{total_valid:,}")
    k2.metric("準時交付筆數", f"{on_time_count:,}")
    k3.metric("達交率", f"{rate:.2f}%")

    # 依出貨類型拆解達交率
    if ship_type_col in data.columns:
        tmp = pd.DataFrame({
            "出貨類型": data[ship_type_col],
            "due": due_day,
            "sign": sign_day,
        })
        tmp = tmp.dropna(subset=["due", "sign"]).copy()
        tmp["是否準時"] = (tmp["sign"] <= tmp["due"]).astype(int)
        by_type = (
            tmp.groupby("出貨類型")["是否準時"]
            .agg(["count", "sum"])  # count=有效數、sum=準時數
            .rename(columns={"count": "有效筆數", "sum": "準時筆數"})
            .reset_index()
        )
        by_type["達交率(%)"] = (by_type["準時筆數"] / by_type["有效筆數"] * 100).round(2)

        st.write("**各出貨類型達交表現（以日期比較）**")
        st.dataframe(by_type, use_container_width=True)
        if not by_type.empty:
            fig2 = px.bar(by_type, x="出貨類型", y="達交率(%)")
            st.plotly_chart(fig2, use_container_width=True)

        st.download_button(
            "下載達交率（依出貨類型，日期比較）CSV",
            data=by_type.to_csv(index=False).encode("utf-8-sig"),
            file_name="各出貨類型_達交率_僅日期.csv",
            mime="text/csv",
        )
else:
    st.warning("請在側欄選好『指定到貨日期』與『客戶簽收日期』欄位。")

st.markdown("---")

    # ----------（保留原有）自訂欄位 / 聚合 / 圖表 / 下載 ----------
    with st.sidebar:
        st.subheader("🧮 新增自訂欄位")
        st.caption("使用現有欄位做計算，例如：`銷售額 - 成本` 或 `數量 * 單價`")
        new_col_name = st.text_input("新欄位名稱", value="")
        formula = st.text_input("輸入公式（等號右邊）", placeholder="例如：數量 * 單價")

    data_for_calc = data.copy()
    if new_col_name and formula:
        try:
            data_for_calc[new_col_name] = pd.eval(formula, engine="python", local_dict=data_for_calc.to_dict("series"))
            st.success(f"已新增欄位：{new_col_name}")
        except Exception as e:
            st.error(f"公式錯誤：{e}")

    with st.sidebar:
        st.subheader("📦 聚合彙整")
        group_cols = st.multiselect("群組欄位（可多選）", list(data_for_calc.columns))
        numeric_cols = data_for_calc.select_dtypes(include="number").columns.tolist()
        agg_col = st.selectbox("聚合欄位", numeric_cols or ["（無數值欄位）"])
        agg_fn = st.selectbox("聚合函數", ["sum", "mean", "max", "min", "count"])

    agg_df = data_for_calc.copy()
    if group_cols and agg_col in data_for_calc.columns:
        agg_df = (
            agg_df.groupby(group_cols, dropna=False)[agg_col]
            .agg(agg_fn)
            .reset_index()
            .rename(columns={agg_col: f"{agg_fn}_{agg_col}"})
        )

    st.subheader("處理後資料（聚合結果）")
    st.dataframe(agg_df, use_container_width=True)

    st.subheader("📈 視覺化")
    num_cols = agg_df.select_dtypes(include="number").columns.tolist()
    cat_cols = [c for c in agg_df.columns if c not in num_cols]
    if len(agg_df.columns) >= 2 and num_cols:
        x = st.selectbox("X 軸", cat_cols + num_cols)
        y = st.selectbox("Y 軸（數值）", num_cols)
        chart_type = st.selectbox("圖表類型", ["bar", "line", "scatter"])
        if chart_type == "bar":
            fig = px.bar(agg_df, x=x, y=y)
        elif chart_type == "line":
            fig = px.line(agg_df, x=x, y=y)
        else:
            fig = px.scatter(agg_df, x=x, y=y)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("目前沒有足夠欄位可畫圖，請先做聚合或新增數值欄位。")

    # 下載
    st.subheader("⬇️ 下載")
    # CSV（聚合結果）
    st.download_button(
        "下載聚合結果 CSV",
        data=agg_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="聚合結果.csv",
        mime="text/csv",
    )

    # Excel（原始/處理/聚合）
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="篩選後原始")
        agg_df.to_excel(writer, index=False, sheet_name="聚合結果")
    st.download_button(
        "下載 Excel",
        data=output.getvalue(),
        file_name="TMS_出貨配送分析.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("請先在上方上傳 Excel 或 CSV 檔。")
