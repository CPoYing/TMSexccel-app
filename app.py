import io
import re
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="TMS出貨配送數據分析", page_icon="📊", layout="wide")
st.title("📊 TMS出貨配送數據分析")
st.caption("上傳 Excel/CSV → 出貨類型筆數（銅重量換算噸）→ 配送達交率（僅日期，排除SWI-寄庫）→ 配送區域分析 → 配送裝載分析（僅完成）")

# ---------- 檔案上傳 ----------
file = st.file_uploader(
    "上傳 Excel 或 CSV 檔", type=["xlsx", "xls", "csv"],
    help="最多 200 MB；Excel 需使用 openpyxl 解析"
)

@st.cache_data
def load_data(file):
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file, engine="openpyxl")
    return df

# ---------- 工具函式 ----------
def _guess_col(cols, keywords):
    for kw in keywords:
        for c in cols:
            if kw in str(c):
                return c
    return None

def to_dt(series):
    return pd.to_datetime(series, errors="coerce")

def extract_city(addr):
    """從地址字串中擷取台灣縣市（將『臺』正規化為『台』），失敗回傳 None。"""
    if pd.isna(addr):
        return None
    s = str(addr).strip().replace("臺", "台")
    pattern = r"(台北市|新北市|桃園市|台中市|台南市|高雄市|基隆市|新竹市|新竹縣|苗栗縣|彰化縣|南投縣|雲林縣|嘉義市|嘉義縣|屏東縣|宜蘭縣|花蓮縣|台東縣|澎湖縣|金門縣|連江縣)"
    m = re.search(pattern, s)
    return m.group(1) if m else None

# ======================== 主流程 ========================
if file:
    df = load_data(file)

    # ---- 自動欄位偵測（不顯示 UI） ----
    cols = list(df.columns)
    ship_type_col     = _guess_col(cols, ["出貨申請類型", "出貨類型", "配送類型", "類型"])
    due_date_col      = _guess_col(cols, ["指定到貨日期", "到貨日期", "指定到貨", "到貨日"])
    sign_date_col     = _guess_col(cols, ["客戶簽收日期", "簽收日期", "簽收日", "客戶簽收日期/時/分"])
    cust_id_col       = _guess_col(cols, ["客戶編號", "客戶代號", "客編"])
    cust_name_col     = _guess_col(cols, ["客戶名稱", "客名", "客戶"])
    address_col       = _guess_col(cols, ["地址", "收貨地址", "送貨地址", "交貨地址"])
    # 注意：實際數值單位是 kg，但欄名可能為「銅重量(噸)」
    copper_ton_col    = _guess_col(cols, ["銅重量(噸)", "銅重量(噸數)", "銅噸", "銅重量"])
    qty_col           = _guess_col(cols, ["出貨數量", "數量", "出貨量"])
    ship_no_col       = _guess_col(cols, ["出庫單號", "出庫單", "出庫編號"])
    do_col            = _guess_col(cols, ["DO號", "DO", "出貨單號", "交貨單號"])
    item_desc_col     = _guess_col(cols, ["料號說明", "品名", "品名規格", "物料說明"])
    lot_col           = _guess_col(cols, ["批次", "批號", "Lot"])
    fg_net_ton_col    = _guess_col(cols, ["成品淨重(噸)", "成品淨量(噸)", "淨重(噸)"])
    fg_gross_ton_col  = _guess_col(cols, ["成品毛重(噸)", "成品毛量(噸)", "毛重(噸)"])
    status_col        = _guess_col(cols, ["通知單項次狀態", "項次狀態", "狀態"])

    # Tabs：分析 / 原始資料
    tab_analysis, tab_raw = st.tabs(["📊 出貨&達交分析", "📄 原始資料預覽"])

    # -------- 原始資料分頁 --------
    with tab_raw:
        st.subheader("原始資料預覽")
        st.dataframe(df, use_container_width=True)

    # -------- 分析分頁 --------
    with tab_analysis:
        # ---------- 側欄操作（常用篩選） ----------
        with st.sidebar:
            st.header("⚙️ 操作區（常用篩選）")

            # 出貨申請類型（簡單多選）
            if ship_type_col and ship_type_col in df.columns:
                ship_opts = sorted(df[ship_type_col].dropna().astype(str).unique().tolist())
                sel_types = st.multiselect("出貨申請類型（多選）", options=ship_opts, default=[])
            else:
                sel_types = []

            # 客戶名稱（可搜尋多選）
            if cust_name_col and cust_name_col in df.columns:
                name_opts = sorted(df[cust_name_col].dropna().astype(str).unique().tolist())
                sel_names = st.multiselect("客戶名稱（可搜尋）", options=name_opts, default=[])
            else:
                sel_names = []

            # 客戶編號（可搜尋多選）
            if cust_id_col and cust_id_col in df.columns:
                id_opts = sorted(df[cust_id_col].dropna().astype(str).unique().tolist())
                sel_ids = st.multiselect("客戶編號（可搜尋）", options=id_opts, default=[])
            else:
                sel_ids = []

            # 指定到貨日期（日期區間）
            if due_date_col and due_date_col in df.columns:
                due_series = to_dt(df[due_date_col])
                min_date = due_series.min().date() if not due_series.dropna().empty else None
                max_date = due_series.max().date() if not due_series.dropna().empty else None
                if min_date and max_date:
                    date_range = st.date_input("指定到貨日期（區間）", value=(min_date, max_date))
                else:
                    date_range = None
            else:
                date_range = None

            # 料號說明（可搜尋多選）
            if item_desc_col and item_desc_col in df.columns:
                item_opts = sorted(df[item_desc_col].dropna().astype(str).unique().tolist())
                sel_items = st.multiselect("料號說明（可搜尋）", options=item_opts, default=[])
            else:
                sel_items = []

            st.markdown("---")

            # 其它：選擇要篩選的欄位（可略）
            filter_col = st.selectbox("選擇要篩選的欄位（可略）", ["（不篩選）"] + cols)
            extra_selected_vals = []
            if filter_col != "（不篩選）":
                unique_vals = df[filter_col].dropna().unique().tolist()
                extra_selected_vals = st.multiselect("選取值（可多選）", unique_vals)

            # 每客戶 Top N 縣市（供③使用）
            topn_cities = st.slider("每客戶 Top N 縣市（③用）", min_value=1, max_value=10, value=3, step=1)

        # ---- 套用篩選 ----
        data = df.copy()

        if sel_types and (ship_type_col in data.columns):
            data = data[data[ship_type_col].astype(str).isin(sel_types)]

        if sel_names and (cust_name_col in data.columns):
            data = data[data[cust_name_col].astype(str).isin(sel_names)]

        if sel_ids and (cust_id_col in data.columns):
            data = data[data[cust_id_col].astype(str).isin(sel_ids)]

        if date_range and (due_date_col in data.columns):
            due_day_all = to_dt(data[due_date_col]).dt.normalize()
            start_d, end_d = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
            data = data[(due_day_all >= start_d) & (due_day_all <= end_d)]

        if sel_items and (item_desc_col in data.columns):
            data = data[data[item_desc_col].astype(str).isin(sel_items)]

        if filter_col != "（不篩選）" and extra_selected_vals:
            data = data[data[filter_col].isin(extra_selected_vals)]

        # 共用：排除 SWI-寄庫（供 ②/③/④ 使用）
        exclude_swi_mask = pd.Series(True, index=data.index)
        if ship_type_col in data.columns:
            exclude_swi_mask = data[ship_type_col] != "SWI-寄庫"

        # =====================================================
        # ① 出貨類型筆數統計（含銅重量換算噸；「銅重量(噸)」欄位實際單位=kg）
        # =====================================================
        st.subheader("① 出貨類型筆數統計")
        if ship_type_col in data.columns:
            type_counts = (
                data[ship_type_col]
                .fillna("(空白)")
                .value_counts(dropna=False)
                .rename_axis("出貨類型")
                .reset_index(name="筆數")
            )
            total_rows = int(type_counts["筆數"].sum())

            # 銅重量合計（kg）→ 顯示成「噸」
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

                merged["銅重量(噸)合計"] = (
                    pd.to_numeric(merged["銅重量(kg)合計"], errors="coerce") / 1000.0
                ).round(2)
                merged.drop(columns=["銅重量(kg)合計"], inplace=True)
            else:
                merged = type_counts.copy()
                merged["銅重量(噸)合計"] = None

            # —— KPI 卡片 ——（st.metric）
            k1, k2 = st.columns(2)
            k1.metric("資料筆數", f"{total_rows:,}")
            total_copper_ton_all = (
                pd.to_numeric(merged["銅重量(噸)合計"], errors="coerce").sum()
                if "銅重量(噸)合計" in merged.columns else 0.0
            )
            k2.metric("總銅重量(噸)", f"{total_copper_ton_all:,.2f}")

            # —— 圖表選擇：長條 / 圓餅 / 環形（圓餅/環形顯示百分比）——
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

        st.markdown("---")

        # =====================================================
        # ② 配送達交率（僅比對日期，不含時分秒；排除 SWI-寄庫）
        # =====================================================
        st.subheader("② 配送達交率（僅比對日期，不含時分秒）")
        if (due_date_col in data.columns) and (sign_date_col in data.columns):
            due_dt = to_dt(data[due_date_col])
            sign_dt = to_dt(data[sign_date_col])
            due_day = due_dt.dt.normalize()
            sign_day = sign_dt.dt.normalize()

            valid_mask = due_day.notna() & sign_day.notna() & exclude_swi_mask
            on_time = (sign_day <= due_day) & valid_mask

            total_valid = int(valid_mask.sum())
            on_time_count = int(on_time.sum())
            rate = (on_time_count / total_valid * 100.0) if total_valid > 0 else 0.0

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("有效筆數（兩日期皆有值）", f"{total_valid:,}")
            k2.metric("準時交付筆數", f"{on_time_count:,}")
            k3.metric("達交率", f"{rate:.2f}%")

            # 未達標（遲交）清單 + 延遲天數 + 平均延遲天數
            late_mask = valid_mask & (sign_day > due_day)
            delay_days = (sign_day - due_day).dt.days
            late_df = pd.DataFrame({
                "客戶編號": data[cust_id_col] if cust_id_col in data.columns else None,
                "客戶名稱": data[cust_name_col] if cust_name_col in data.columns else None,
                "指定到貨日期": due_day.dt.strftime("%Y-%m-%d"),
                "客戶簽收日期": sign_day.dt.strftime("%Y-%m-%d"),
                "延遲天數": delay_days,
            })[late_mask]

            avg_delay = float(late_df["延遲天數"].mean()) if not late_df.empty else 0.0
            k4.metric("平均延遲天數", f"{avg_delay:.2f}")

            st.write("**未達標清單**（已排除SWI-寄庫；僅列簽收晚於指定到貨者）")
            st.dataframe(late_df, use_container_width=True)
            st.download_button(
                "下載未達標清單 CSV",
                data=late_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="未達標清單.csv",
                mime="text/csv",
            )

            # 依客戶名稱：未達交筆數與比例（僅顯示未達交 > 0）
            if cust_name_col in data.columns:
                tmp = pd.DataFrame({
                    "客戶名稱": data[cust_name_col],
                    "是否有效": valid_mask,
                    "是否遲交": late_mask,
                })
                grp = tmp.groupby("客戶名稱")
                stats = grp["是否有效"].sum().to_frame(name="有效筆數")
                stats["未達交筆數"] = grp["是否遲交"].sum()
                stats = stats[stats["未達交筆數"] > 0]
                stats["未達交比例(%)"] = (stats["未達交筆數"] / stats["有效筆數"] * 100).round(2)
                stats = stats.reset_index().sort_values(["未達交筆數", "未達交比例(%)"], ascending=[False, False])

                st.write("**依客戶名稱統計：未達交筆數與比例（>0）**")
                st.dataframe(stats, use_container_width=True)
                st.download_button(
                    "下載未達交統計（客戶） CSV",
                    data=stats.to_csv(index=False).encode("utf-8-sig"),
                    file_name="未達交統計_依客戶.csv",
                    mime="text/csv",
                )
        else:
            st.warning("找不到『指定到貨日期』或『客戶簽收日期』欄位。")

        st.markdown("---")

        # =====================================================
        # ③ 配送區域分析（排除 SWI-寄庫）
        # =====================================================
        st.subheader("③ 配送區域分析（排除 SWI-寄庫）")
        if address_col in data.columns:
            region_df = data[exclude_swi_mask].copy()
            region_df["縣市"] = region_df[address_col].apply(extract_city)
            region_df["縣市"] = region_df["縣市"].fillna("(未知)")

            # 各縣市配送筆數 + 佔比
            city_counts = (
                region_df["縣市"].value_counts(dropna=False)
                .rename_axis("縣市").reset_index(name="筆數")
            )
            total_city = int(city_counts["筆數"].sum()) if not city_counts.empty else 0
            city_counts["縣市佔比(%)"] = (
                (city_counts["筆數"] / total_city * 100).round(2) if total_city > 0 else 0
            )

            c1, c2 = st.columns([1, 1])
            with c1:
                st.write("**各縣市配送筆數**")
                st.dataframe(city_counts, use_container_width=True)
                st.download_button(
                    "下載各縣市配送筆數 CSV",
                    data=city_counts.to_csv(index=False).encode("utf-8-sig"),
                    file_name="各縣市配送筆數.csv",
                    mime="text/csv",
                )
            with c2:
                if not city_counts.empty:
                    fig_city = px.bar(city_counts, x="縣市", y="筆數")
                    st.plotly_chart(fig_city, use_container_width=True)

            # 各客戶常出貨縣市（每客戶 Top N）
            if cust_name_col in region_df.columns:
                cust_city = (
                    region_df.groupby([cust_name_col, "縣市"]).size()
                    .reset_index(name="筆數")
                )
                cust_city["客戶總筆數"] = cust_city.groupby(cust_name_col)["筆數"].transform("sum")
                cust_city["縣市佔比(%)"] = (cust_city["筆數"] / cust_city["客戶總筆數"] * 100).round(2)
                cust_city = cust_city.sort_values([cust_name_col, "筆數"], ascending=[True, False])
                top_table = cust_city.groupby(cust_name_col).head(topn_cities)

                st.write(f"**各客戶常出貨縣市（每客戶 Top {topn_cities}）**")
                st.dataframe(top_table.rename(columns={cust_name_col: "客戶名稱"}), use_container_width=True)
                st.download_button(
                    "下載各客戶常出貨縣市 CSV",
                    data=top_table.to_csv(index=False).encode("utf-8-sig"),
                    file_name="各客戶常出貨縣市_TopN.csv",
                    mime="text/csv",
                )
        else:
            st.warning("找不到『地址』欄位。")

        st.markdown("---")

        # =====================================================
        # ④ 配送裝載分析（出庫單號前13碼=同車；排除 SWI-寄庫；僅『完成』）
        # =====================================================
        st.subheader("④ 配送裝載分析（出庫單號前13碼=同車；排除 SWI-寄庫；僅完成）")
        if ship_no_col in data.columns:
            # 僅分析「通知單項次狀態＝完成」且排除 SWI-寄庫
            if status_col in data.columns:
                status_mask = data[status_col].astype(str).str.strip() == "完成"
                base_mask = exclude_swi_mask & status_mask
            else:
                st.info("未對應到『通知單項次狀態』欄位，將不套用「完成」過濾。")
                base_mask = exclude_swi_mask

            load_df = data[base_mask].copy()
            load_df["車次代碼"] = load_df[ship_no_col].astype(str).str[:13]  # 分車用，清單仍顯示完整出庫單號

            # 數值轉換
            load_df["_qty_num"] = pd.to_numeric(load_df[qty_col], errors="coerce") if (qty_col in load_df.columns) else pd.Series(dtype="float64", index=load_df.index)
            load_df["_copper_kg"] = pd.to_numeric(load_df[copper_ton_col], errors="coerce") if (copper_ton_col in load_df.columns) else pd.Series(dtype="float64", index=load_df.index)
            load_df["_copper_ton"] = load_df["_copper_kg"] / 1000.0
            load_df["_fg_net_ton"] = pd.to_numeric(load_df[fg_net_ton_col], errors="coerce") if (fg_net_ton_col in load_df.columns) else pd.Series(dtype="float64", index=load_df.index)
            load_df["_fg_gross_ton"] = pd.to_numeric(load_df[fg_gross_ton_col], errors="coerce") if (fg_gross_ton_col in load_df.columns) else pd.Series(dtype="float64", index=load_df.index)

            # 明細清單
            wanted = [
                ("車次代碼", "車次代碼"),
                (ship_no_col, "出庫單號"),
                (do_col, "DO號"),
                (item_desc_col, "料號說明"),
                (lot_col, "批次"),
                ("_qty_num", "出貨數量"),
                ("_copper_ton", "銅重量(噸)"),
                ("_fg_net_ton", "成品淨重(噸)"),
                ("_fg_gross_ton", "成品毛重(噸)"),
            ]
            display_cols, used_names = {}, set()
            for src, nice in wanted:
                if isinstance(src, str) and (src in load_df.columns):
                    name = nice
                    i = 2
                    while name in used_names:
                        name = f"{nice} ({i})"; i += 1
                    display_cols[name] = load_df[src]; used_names.add(name)

            if display_cols:
                display_df = pd.DataFrame(display_cols)
                for c in ["出貨數量", "銅重量(噸)", "成品淨重(噸)", "成品毛重(噸)"]:
                    if c in display_df.columns:
                        display_df[c] = pd.to_numeric(display_df[c], errors="coerce").round(3)
                sort_keys = [c for c in ["車次代碼", "出庫單號"] if c in display_df.columns]
                if sort_keys:
                    display_df = display_df.sort_values(by=sort_keys).reset_index(drop=True)

                st.write("**配送裝載清單（依車次代碼分車；僅完成）**")
                st.dataframe(display_df, use_container_width=True)
                st.download_button(
                    "下載配送裝載清單 CSV",
                    data=display_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name="配送裝載清單.csv",
                    mime="text/csv",
                )
            else:
                st.info("目前無可顯示欄位，請確認原始資料。")

            # 車次彙總（每車小計）
            group = load_df.groupby("車次代碼", dropna=False)
            agg_dict = {
                "出貨數量小計": ("_qty_num", "sum"),
                "銅重量(kg)_小計": ("_copper_kg", "sum"),
                "成品淨重(噸)_小計": ("_fg_net_ton", "sum"),
                "成品毛重(噸)_小計": ("_fg_gross_ton", "sum"),
            }
            summary = group.agg(**agg_dict).reset_index()
            if "銅重量(kg)_小計" in summary.columns:
                summary["銅重量(噸)_小計"] = (pd.to_numeric(summary["銅重量(kg)_小計"], errors="coerce") / 1000.0).round(3)
            for c in ["出貨數量小計", "銅重量(kg)_小計", "成品淨重(噸)_小計", "成品毛重(噸)_小計", "銅重量(噸)_小計"]:
                if c in summary.columns:
                    summary[c] = pd.to_numeric(summary[c], errors="coerce").round(3)

            st.write("**車次彙總（每車小計；僅完成）**")
            st.dataframe(summary, use_container_width=True)
            st.download_button(
                "下載車次彙總 CSV",
                data=summary.to_csv(index=False).encode("utf-8-sig"),
                file_name="車次彙總.csv",
                mime="text/csv",
            )

            # KPI：平均車子載重(噸)（僅計入銅重量>0的車次）
            if "銅重量(噸)_小計" in summary.columns:
                valid = summary[pd.to_numeric(summary["銅重量(噸)_小計"], errors="coerce") > 0]
                avg_load = float(valid["銅重量(噸)_小計"].mean()) if not valid.empty else 0.0
                k1, k2 = st.columns(2)
                k1.metric("納入計算車次", f"{len(valid):,}")
                k2.metric("平均車子載重(噸)", f"{avg_load:.2f}")
        else:
            st.warning("找不到『出庫單號』欄位。")

else:
    st.info("請先在上方上傳 Excel 或 CSV 檔。")
