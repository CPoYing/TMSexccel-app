import io
import re
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="TMS出貨配送數據分析", page_icon="📊", layout="wide")
st.title("📊 TMS出貨配送數據分析")
st.caption("上傳 Excel/CSV → 出貨類型筆數（含銅重量kg）→ 達交率（僅日期，排除SWI-寄庫）→ 配送區域分析 → 配送裝載分析 → 自訂欄位 → 聚合(獨立分頁) → 視覺化/下載")

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

    # Tabs：分析 / 原始資料 / 聚合結果
    tab_analysis, tab_raw, tab_agg = st.tabs(["📊 出貨&達交分析", "📄 原始資料預覽", "📦 處理後資料（聚合結果）"])

    # -------- 原始資料分頁 --------
    with tab_raw:
        st.subheader("原始資料預覽")
        st.dataframe(df, use_container_width=True)

    # -------- 分析分頁 --------
    with tab_analysis:
        # ---------- 側欄操作 ----------
        with st.sidebar:
            st.header("⚙️ 操作區")
            cols = list(df.columns)

            # 篩選器（可略）
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
            cust_id_default = _guess_col(cols, ["客戶編號", "客戶代號", "客編"]) or (cols[0] if cols else None)
            cust_name_default = _guess_col(cols, ["客戶名稱", "客名", "客戶"]) or (cols[0] if cols else None)
            address_default = _guess_col(cols, ["地址", "收貨地址", "送貨地址", "交貨地址"]) or (cols[0] if cols else None)

            # 「銅重量(噸)」欄位實際單位為 kg，這裡仍讓你對到該欄位
            copper_ton_default = _guess_col(cols, ["銅重量(噸)", "銅重量(噸數)", "銅噸", "銅重量"]) or (cols[0] if cols else None)
            qty_default = _guess_col(cols, ["出貨數量", "數量", "出貨量"]) or (cols[0] if cols else None)
            shipno_default = _guess_col(cols, ["出庫單號", "出庫單", "出庫編號"]) or (cols[0] if cols else None)
            do_default = _guess_col(cols, ["DO號", "DO", "出貨單號", "交貨單號"]) or (cols[0] if cols else None)
            item_desc_default = _guess_col(cols, ["料號說明", "品名", "品名規格", "物料說明"]) or (cols[0] if cols else None)
            lot_default = _guess_col(cols, ["批次", "批號", "Lot"]) or (cols[0] if cols else None)
            fg_net_ton_default = _guess_col(cols, ["成品淨重(噸)", "淨重(噸)"]) or (cols[0] if cols else None)
            fg_gross_ton_default = _guess_col(cols, ["成品毛重(噸)", "毛重(噸)"]) or (cols[0] if cols else None)

            ship_type_col = st.selectbox("出貨類型欄位", options=cols, index=(cols.index(ship_type_default) if ship_type_default in cols else 0))
            due_date_col = st.selectbox("指定到貨日期欄位", options=cols, index=(cols.index(due_date_default) if due_date_default in cols else 0))
            sign_date_col = st.selectbox("客戶簽收日期欄位", options=cols, index=(cols.index(sign_date_default) if sign_date_default in cols else 0))
            cust_id_col = st.selectbox("客戶編號欄位", options=cols, index=(cols.index(cust_id_default) if cust_id_default in cols else 0))
            cust_name_col = st.selectbox("客戶名稱欄位", options=cols, index=(cols.index(cust_name_default) if cust_name_default in cols else 0))
            address_col = st.selectbox("地址欄位", options=cols, index=(cols.index(address_default) if address_default in cols else 0))
            copper_ton_col = st.selectbox("銅重量(噸) 欄位（實際單位=kg）", options=cols, index=(cols.index(copper_ton_default) if copper_ton_default in cols else 0))
            qty_col = st.selectbox("出貨數量 欄位", options=cols, index=(cols.index(qty_default) if qty_default in cols else 0))
            ship_no_col = st.selectbox("出庫單號 欄位", options=cols, index=(cols.index(shipno_default) if shipno_default in cols else 0))
            do_col = st.selectbox("DO號 欄位", options=cols, index=(cols.index(do_default) if do_default in cols else 0))
            item_desc_col = st.selectbox("料號說明 欄位", options=cols, index=(cols.index(item_desc_default) if item_desc_default in cols else 0))
            lot_col = st.selectbox("批次 欄位", options=cols, index=(cols.index(lot_default) if lot_default in cols else 0))
            fg_net_ton_col = st.selectbox("成品淨重(噸) 欄位", options=cols, index=(cols.index(fg_net_ton_default) if fg_net_ton_default in cols else 0))
            fg_gross_ton_col = st.selectbox("成品毛重(噸) 欄位", options=cols, index=(cols.index(fg_gross_ton_default) if fg_gross_ton_default in cols else 0))

            # 配送區域 TopN（每客戶顯示前幾名縣市）
            topn = st.slider("每客戶顯示前幾大縣市", min_value=1, max_value=10, value=3, step=1)

        # 套用篩選後資料
        data = filtered.copy()

        # 共用：排除 SWI-寄庫（供達交率/配送區域/配送裝載使用）
        exclude_swi_mask = pd.Series(True, index=data.index)
        if ship_type_col in data.columns:
            exclude_swi_mask = data[ship_type_col] != "SWI-寄庫"

        # =====================================================
        # ① 出貨類型筆數統計（右側新增：銅重量(kg)合計；且將「銅重量(噸)」視為kg數值）
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

            # 將銅重量欄位轉為數值（實際單位=kg）
            if copper_ton_col in data.columns:
                data["_copper_kg"] = pd.to_numeric(data[copper_ton_col], errors="coerce")
                copper_sum = (
                    data.groupby(ship_type_col)["_copper_kg"]
                    .sum(min_count=1)
                    .reset_index()
                    .rename(columns={"_copper_kg": "銅重量(kg)合計"})
                )
                merged = type_counts.merge(copper_sum, how="left", left_on="出貨類型", right_on=ship_type_col)
                if ship_type_col in merged.columns:
                    merged.drop(columns=[ship_type_col], inplace=True)
            else:
                merged = type_counts.copy()
                merged["銅重量(kg)合計"] = None

            chart_choice = st.radio("圖表類型", ["長條圖", "圓餅圖", "折線圖"], horizontal=True)

            c1, c2 = st.columns([1, 1])
            with c1:
                st.write(f"**加總筆數：{total_rows:,}**")
                if "銅重量(kg)合計" in merged.columns:
                    total_copper = pd.to_numeric(merged["銅重量(kg)合計"], errors="coerce").sum()
                    st.write(f"**總銅重量 (kg)：{total_copper:,.2f}**")
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
                    elif chart_choice == "圓餅圖":
                        fig = px.pie(type_counts, names="出貨類型", values="筆數", hole=0)
                    else:  # 折線圖
                        fig = px.line(type_counts, x="出貨類型", y="筆數", markers=True)
                    st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("找不到出貨類型欄位，請在側欄重新選擇。")

        st.markdown("---")

        # =====================================================
        # ② 達交率（僅比對日期，不含時分秒；排除 SWI-寄庫）
        # =====================================================
        st.subheader("② 達交率（僅比對日期，不含時分秒）")
        if due_date_col in data.columns and sign_date_col in data.columns:
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

            st.write("**未達標清單**（已排除出貨類型=SWI-寄庫；僅列簽收晚於指定到貨者）")
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

                st.write("**依客戶名稱統計：未達交筆數與比例（僅顯示未達交>0；已排除SWI-寄庫）**")
                st.dataframe(stats, use_container_width=True)
                st.download_button(
                    "下載未達交統計（客戶） CSV",
                    data=stats.to_csv(index=False).encode("utf-8-sig"),
                    file_name="未達交統計_依客戶.csv",
                    mime="text/csv",
                )
        else:
            st.warning("請在側欄選好『指定到貨日期』與『客戶簽收日期』欄位。")

        st.markdown("---")

        # =====================================================
        # ③ 配送區域分析（排除 SWI-寄庫）
        # =====================================================
        st.subheader("③ 配送區域分析（排除 SWI-寄庫）")
        if address_col in data.columns:
            region_df = data[exclude_swi_mask].copy()
            region_df["縣市"] = region_df[address_col].apply(extract_city)
            region_df["縣市"] = region_df["縣市"].fillna("(未知)")

            # 3-1 各縣市配送筆數 + 縣市佔比
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

            st.markdown("—")

            # 3-2 各客戶常出貨縣市（每客戶 TopN）+ 客戶內部的縣市佔比
            if cust_name_col in region_df.columns:
                cust_city = (
                    region_df.groupby([cust_name_col, "縣市"]).size()
                    .reset_index(name="筆數")
                )
                cust_city["客戶總筆數"] = cust_city.groupby(cust_name_col)["筆數"].transform("sum")
                cust_city["縣市佔比(%)"] = (cust_city["筆數"] / cust_city["客戶總筆數"] * 100).round(2)
                cust_city = cust_city.sort_values([cust_name_col, "筆數"], ascending=[True, False])
                top_table = cust_city.groupby(cust_name_col).head(topn)

                st.write(f"**各客戶常出貨縣市（每客戶 Top {topn}）**")
                st.dataframe(top_table.rename(columns={cust_name_col: "客戶名稱"}), use_container_width=True)
                st.download_button(
                    "下載各客戶常出貨縣市 CSV",
                    data=top_table.to_csv(index=False).encode("utf-8-sig"),
                    file_name="各客戶常出貨縣市_TopN.csv",
                    mime="text/csv",
                )
            else:
                st.info("找不到『客戶名稱』欄位，無法產生各客戶常出貨縣市統計。")
        else:
            st.warning("請在側欄選擇正確的『地址』欄位。")

        st.markdown("---")

# =====================================================
# ④ 配送裝載分析（出庫單號前13碼=同車；排除 SWI-寄庫）
# =====================================================
        st.subheader("④ 配送裝載分析（出庫單號前13碼=同車，排除 SWI-寄庫）")
            if ship_no_col in data.columns:
            load_df = data[exclude_swi_mask].copy()
        
            # 以出庫單號前13碼定義「車次代碼」，但清單仍顯示完整出庫單號
            load_df["車次代碼"] = load_df[ship_no_col].astype(str).str[:13]
        
            # 數值轉換：出貨數量、銅重量(kg)（原欄名為「銅重量(噸)」但實際單位是 kg）、成品淨/毛重(噸)
            load_df["_qty_num"] = (
                pd.to_numeric(load_df[qty_col], errors="coerce")
                if (qty_col in load_df.columns) else pd.Series(dtype="float64", index=load_df.index)
            )
            load_df["_copper_kg"] = (
                pd.to_numeric(load_df[copper_ton_col], errors="coerce")
                if (copper_ton_col in load_df.columns) else pd.Series(dtype="float64", index=load_df.index)
            )
            load_df["_copper_ton"] = load_df["_copper_kg"] / 1000.0
        
            load_df["_fg_net_ton"] = (
                pd.to_numeric(load_df[fg_net_ton_col], errors="coerce")
                if (fg_net_ton_col in load_df.columns) else pd.Series(dtype="float64", index=load_df.index)
            )
            load_df["_fg_gross_ton"] = (
                pd.to_numeric(load_df[fg_gross_ton_col], errors="coerce")
                if (fg_gross_ton_col in load_df.columns) else pd.Series(dtype="float64", index=load_df.index)
            )
        
            # ==== 明細清單 ====
            # 使用安全命名與去重，避免欄位重名造成崩潰；出庫單號顯示完整值
            wanted = [
                ("車次代碼", "車次代碼"),
                (ship_no_col, "出庫單號"),        # 顯示完整值
                (do_col, "DO號"),
                (item_desc_col, "料號說明"),
                (lot_col, "批次"),
                ("_qty_num", "出貨數量"),
                ("_copper_ton", "銅重量(噸)"),   # 以 kg 轉噸顯示
                ("_fg_net_ton", "成品淨重(噸)"),
                ("_fg_gross_ton", "成品毛重(噸)"),
            ]
            display_cols = {}
            used_names = set()
            for src, nice in wanted:
                # src 可能是字串欄位名也可能是計算欄（上面已放到 load_df）
                if isinstance(src, str) and (src in load_df.columns):
                    name = nice
                    i = 2
                    while name in used_names:
                        name = f"{nice} ({i})"
                        i += 1
                    display_cols[name] = load_df[src]
                    used_names.add(name)
        
            if display_cols:
                display_df = pd.DataFrame(display_cols)
        
                # 數值欄四捨五入
                for c in ["出貨數量", "銅重量(噸)", "成品淨重(噸)", "成品毛重(噸)"]:
                    if c in display_df.columns:
                        display_df[c] = pd.to_numeric(display_df[c], errors="coerce").round(3)
        
                # 依車次代碼、出庫單號排序
                sort_keys = [c for c in ["車次代碼", "出庫單號"] if c in display_df.columns]
                if sort_keys:
                    display_df = display_df.sort_values(by=sort_keys).reset_index(drop=True)
        
                st.write("**配送裝載清單（依車次代碼分車）**")
                st.dataframe(display_df, use_container_width=True)
                st.download_button(
                    "下載配送裝載清單 CSV",
                    data=display_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name="配送裝載清單.csv",
                    mime="text/csv",
                )
            else:
                st.info("已排除 SWI-寄庫，但目前無可顯示欄位，請在側欄確認欄位對應。")
        
            # ==== 車次彙總（每車小計） ====
            group = load_df.groupby("車次代碼", dropna=False)
            agg_dict = {
                "出貨數量小計": ("_qty_num", "sum"),
                "銅重量(kg)_小計": ("_copper_kg", "sum"),
                "成品淨重(噸)_小計": ("_fg_net_ton", "sum"),
                "成品毛重(噸)_小計": ("_fg_gross_ton", "sum"),
            }
            summary = group.agg(**agg_dict).reset_index()
        
            # 補上銅重量(噸)_小計（由 kg 轉噸）
            if "銅重量(kg)_小計" in summary.columns:
                summary["銅重量(噸)_小計"] = (
                    pd.to_numeric(summary["銅重量(kg)_小計"], errors="coerce") / 1000.0
                ).round(3)
        
            # 數值欄統一四捨五入
            for c in ["出貨數量小計", "銅重量(kg)_小計", "成品淨重(噸)_小計", "成品毛重(噸)_小計", "銅重量(噸)_小計"]:
                if c in summary.columns:
                    summary[c] = pd.to_numeric(summary[c], errors="coerce").round(3)
        
            st.write("**車次彙總（每車小計）**")
            st.dataframe(summary, use_container_width=True)
            st.download_button(
                "下載車次彙總 CSV",
                data=summary.to_csv(index=False).encode("utf-8-sig"),
                file_name="車次彙總.csv",
                mime="text/csv",
            )
        
            # ==== KPI：平均車子載重(噸) ====
            # 僅計算銅重量(噸)_小計 > 0 的車次
            if "銅重量(噸)_小計" in summary.columns:
                valid = summary[pd.to_numeric(summary["銅重量(噸)_小計"], errors="coerce") > 0]
                avg_load = float(valid["銅重量(噸)_小計"].mean()) if not valid.empty else 0.0
                k1, k2 = st.columns(2)
                k1.metric("納入計算車次", f"{len(valid):,}")
                k2.metric("平均車子載重(噸)", f"{avg_load:.2f}")
        else:
            st.info("請在側欄指定『出庫單號』欄位，以進行裝載分析。")
        

        # =====================================================
        # 自訂欄位 → 聚合（計算在此，但顯示搬到 tab_agg）
        # =====================================================
        with st.sidebar:
            st.subheader("🧮 新增自訂欄位")
            st.caption("使用現有欄位做計算，例如：`銷售額 - 成本` 或 `數量 * 單價`")
            new_col_name = st.text_input("新欄位名稱", value="")
            formula = st.text_input("輸入公式（等號右邊）", placeholder="例如：數量 * 單價")

        data_for_calc = data.copy()
        if new_col_name and formula:
            try:
                data_for_calc[new_col_name] = pd.eval(
                    formula, engine="python", local_dict=data_for_calc.to_dict("series")
                )
                st.success(f"已新增欄位：{new_col_name}")
            except Exception as e:
                st.error(f"公式錯誤：{e}")

        with st.sidebar:
            st.subheader("📦 聚合彙整（結果顯示在『📦 聚合結果』分頁）")
            group_cols = st.multiselect("群組欄位（可多選）", list(data_for_calc.columns))
            numeric_cols = data_for_calc.select_dtypes(include="number").columns.tolist()
            agg_col = st.selectbox("聚合欄位", numeric_cols or ["（無數值欄位）"])
            agg_fn = st.selectbox("聚合函數", ["sum", "mean", "max", "min", "count"])

        # 在分析分頁中先算好，稍後在 tab_agg 顯示
        agg_df = data_for_calc.copy()
        if group_cols and (agg_col in data_for_calc.columns):
            agg_df = (
                agg_df.groupby(group_cols, dropna=False)[agg_col]
                .agg(agg_fn)
                .reset_index()
                .rename(columns={agg_col: f"{agg_fn}_{agg_col}"})
            )

        # ---------- 視覺化（沿用聚合結果） ----------
        st.subheader("📈 視覺化（使用目前聚合結果）")
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

        # 把 agg_df 暴露到外層供 tab_agg 使用
        st.session_state["_agg_df"] = agg_df
        st.session_state["_data_filtered"] = data

    # ============= 聚合結果分頁（顯示用） =============
    with tab_agg:
        st.subheader("處理後資料（聚合結果）")
        agg_df = st.session_state.get("_agg_df", pd.DataFrame())
        data = st.session_state.get("_data_filtered", df)
        st.dataframe(agg_df, use_container_width=True)

        st.subheader("⬇️ 下載")
        st.download_button(
            "下載聚合結果 CSV",
            data=agg_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="聚合結果.csv",
            mime="text/csv",
        )

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


