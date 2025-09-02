import io
import re
import pandas as pd
import streamlit as st
import plotly.express as px
from typing import Dict, Any

# ========================
# 1. 頁面設定與工具函式
# ========================

st.set_page_config(
    page_title="TMS出貨配送數據分析",
    page_icon="📊",
    layout="wide"
)
st.title("📊 TMS出貨配送數據分析")
st.caption("上傳 Excel/CSV → 出貨類型筆數（銅重量換算噸）→ 配送達交率 → 配送區域分析 → 配送裝載分析")

# --- 檔案上傳與快取 ---
@st.cache_data
def load_data(file_buffer: io.BytesIO, file_type: str) -> pd.DataFrame:
    """快取並載入 Excel 或 CSV 檔案。"""
    if file_type == "csv":
        return pd.read_csv(file_buffer)
    elif file_type in ["xlsx", "xls"]:
        return pd.read_excel(file_buffer, engine="openpyxl")
    else:
        st.error("不支援的檔案類型。請上傳 .xlsx, .xls 或 .csv 檔案。")
        return pd.DataFrame()

# --- 欄位自動偵測 ---
def _guess_col(cols: list, keywords: list) -> str | None:
    """從欄位列表中猜測符合關鍵字的欄位名稱。"""
    for kw in keywords:
        for c in cols:
            if kw in str(c):
                return c
    return None

def guess_column_names(df_cols: list) -> Dict[str, str | None]:
    """批量偵測常用的欄位名稱，並回傳字典。"""
    col_map = {
        "ship_type": _guess_col(df_cols, ["出貨申請類型", "出貨類型", "配送類型", "類型"]),
        "due_date": _guess_col(df_cols, ["指定到貨日期", "到貨日期", "指定到貨", "到貨日"]),
        "sign_date": _guess_col(df_cols, ["客戶簽收日期", "簽收日期", "簽收日", "客戶簽收日期/時/分"]),
        "cust_id": _guess_col(df_cols, ["客戶編號", "客戶代號", "客編"]),
        "cust_name": _guess_col(df_cols, ["客戶名稱", "客名", "客戶"]),
        "address": _guess_col(df_cols, ["地址", "收貨地址", "送貨地址", "交貨地址"]),
        "copper_ton": _guess_col(df_cols, ["銅重量(噸)", "銅重量(噸數)", "銅噸", "銅重量"]),
        "qty": _guess_col(df_cols, ["出貨數量", "數量", "出貨量"]),
        "ship_no": _guess_col(df_cols, ["出庫單號", "出庫單", "出庫編號"]),
        "do_no": _guess_col(df_cols, ["DO號", "DO", "出貨單號", "交貨單號"]),
        "item_desc": _guess_col(df_cols, ["料號說明", "品名", "品名規格", "物料說明"]),
        "lot_no": _guess_col(df_cols, ["批次", "批號", "Lot"]),
        "fg_net_ton": _guess_col(df_cols, ["成品淨重(噸)", "成品淨量(噸)", "淨重(噸)"]),
        "fg_gross_ton": _guess_col(df_cols, ["成品毛重(噸)", "成品毛量(噸)", "毛重(噸)"]),
        "status": _guess_col(df_cols, ["通知單項次狀態", "項次狀態", "狀態"]),
    }
    return col_map

# --- 資料處理與正規化 ---
def to_dt(series: pd.Series) -> pd.Series:
    """將 Series 轉換為日期時間格式。"""
    return pd.to_datetime(series, errors="coerce")

def extract_city(addr: str) -> str | None:
    """從地址字串中擷取台灣縣市（將『臺』正規化為『台』）。"""
    if pd.isna(addr):
        return None
    s = str(addr).strip().replace("臺", "台")
    pattern = r"(台北市|新北市|桃園市|台中市|台南市|高雄市|基隆市|新竹市|新竹縣|苗栗縣|彰化縣|南投縣|雲林縣|嘉義市|嘉義縣|屏東縣|宜蘭縣|花蓮縣|台東縣|澎湖縣|金門縣|連江縣)"
    m = re.search(pattern, s)
    return m.group(1) if m else None

# ========================
# 2. 核心分析函式
# ========================

def display_shipment_type_analysis(df: pd.DataFrame, col_map: Dict[str, str | None]):
    """① 顯示出貨類型筆數統計與銅重量。"""
    st.subheader("① 出貨類型筆數與重量統計")
    ship_type_col = col_map.get("ship_type")
    copper_ton_col = col_map.get("copper_ton")

    if not ship_type_col or ship_type_col not in df.columns:
        st.warning("找不到『出貨申請類型』欄位，請確認上傳檔案欄名。")
        return

    type_counts = df[ship_type_col].fillna("(空白)").value_counts(dropna=False).rename_axis("出貨類型").reset_index(name="筆數")
    total_rows = int(type_counts["筆數"].sum())

    merged_df = type_counts.copy()
    if copper_ton_col and copper_ton_col in df.columns:
        df["_copper_kg"] = pd.to_numeric(df[copper_ton_col], errors="coerce")
        copper_sums = df.groupby(ship_type_col)["_copper_kg"].sum(min_count=1).reset_index(name="銅重量(kg)合計")
        merged_df = merged_df.merge(copper_sums, how="left", left_on="出貨類型", right_on=ship_type_col)
        merged_df["銅重量(噸)合計"] = (merged_df["銅重量(kg)合計"] / 1000).round(2)
        merged_df.drop(columns=[ship_type_col, "銅重量(kg)合計"], inplace=True)
    else:
        merged_df["銅重量(噸)合計"] = None

    k1, k2 = st.columns(2)
    k1.metric("總資料筆數", f"{total_rows:,}")
    total_copper_ton = merged_df["銅重量(噸)合計"].sum() if "銅重量(噸)合計" in merged_df.columns else 0.0
    k2.metric("總銅重量(噸)", f"{total_copper_ton:,.2f}")

    c1, c2 = st.columns([1, 1])
    with c1:
        st.dataframe(merged_df, use_container_width=True)
        st.download_button(
            "下載出貨類型統計 CSV",
            data=merged_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="出貨類型_統計.csv",
            mime="text/csv",
        )
    with c2:
        chart_choice = st.radio("圖表類型", ["長條圖", "圓餅圖", "環形圖"], horizontal=True, key="chart1")
        fig = px.bar(type_counts, x="出貨類型", y="筆數") if chart_choice == "長條圖" else px.pie(
            type_counts, names="出貨類型", values="筆數", hole=0.5 if chart_choice == "環形圖" else 0.0
        )
        st.plotly_chart(fig, use_container_width=True)

def display_delivery_performance(df: pd.DataFrame, col_map: Dict[str, str | None], exclude_swi_mask: pd.Series):
    """② 顯示配送達交率分析。"""
    st.subheader("② 配送達交率分析（不含時分秒）")
    due_date_col = col_map.get("due_date")
    sign_date_col = col_map.get("sign_date")
    cust_id_col = col_map.get("cust_id")
    cust_name_col = col_map.get("cust_name")

    if not due_date_col or not sign_date_col or due_date_col not in df.columns or sign_date_col not in df.columns:
        st.warning("找不到『指定到貨日期』或『客戶簽收日期』欄位。")
        return

    due_day = to_dt(df[due_date_col]).dt.normalize()
    sign_day = to_dt(df[sign_date_col]).dt.normalize()
    valid_mask = due_day.notna() & sign_day.notna() & exclude_swi_mask
    on_time_mask = (sign_day <= due_day) & valid_mask

    total_valid = int(valid_mask.sum())
    on_time_count = int(on_time_mask.sum())
    rate = (on_time_count / total_valid * 100) if total_valid > 0 else 0.0

    late_mask = valid_mask & (sign_day > due_day)
    delay_days = (sign_day - due_day).dt.days
    late_df = pd.DataFrame({
        "客戶編號": df[cust_id_col] if cust_id_col else None,
        "客戶名稱": df[cust_name_col] if cust_name_col else None,
        "指定到貨日期": due_day.dt.strftime("%Y-%m-%d"),
        "客戶簽收日期": sign_day.dt.strftime("%Y-%m-%d"),
        "延遲天數": delay_days,
    })[late_mask].dropna(axis=1, how="all")

    avg_delay = late_df["延遲天數"].mean() if not late_df.empty else 0.0
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("有效筆數", f"{total_valid:,}")
    k2.metric("準時交付筆數", f"{on_time_count:,}")
    k3.metric("達交率", f"{rate:.2f}%")
    k4.metric("平均延遲天數", f"{avg_delay:.2f}")

    st.write("**未達標清單**（已排除SWI-寄庫；僅列簽收晚於指定到貨者）")
    st.dataframe(late_df, use_container_width=True)
    st.download_button(
        "下載未達標清單 CSV",
        data=late_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="未達標清單.csv",
        mime="text/csv",
    )
    if cust_name_col and cust_name_col in df.columns:
        customer_perf = pd.DataFrame({
            "客戶名稱": df[cust_name_col],
            "是否有效": valid_mask,
            "是否遲交": late_mask,
        }).groupby("客戶名稱").agg(
            有效筆數=("是否有效", "sum"),
            未達交筆數=("是否遲交", "sum")
        )
        customer_perf = customer_perf[customer_perf["未達交筆數"] > 0].reset_index()
        customer_perf["未達交比例(%)"] = (customer_perf["未達交筆數"] / customer_perf["有效筆數"] * 100).round(2)
        customer_perf = customer_perf.sort_values(["未達交筆數", "未達交比例(%)"], ascending=[False, False])
        st.write("**依客戶名稱統計：未達交筆數與比例（>0）**")
        st.dataframe(customer_perf, use_container_width=True)
        st.download_button(
            "下載未達交統計（客戶） CSV",
            data=customer_perf.to_csv(index=False).encode("utf-8-sig"),
            file_name="未達交統計_依客戶.csv",
            mime="text/csv",
        )

def display_area_analysis(df: pd.DataFrame, col_map: Dict[str, str | None], exclude_swi_mask: pd.Series, topn_cities: int):
    """③ 顯示配送區域分析。"""
    st.subheader("③ 配送區域分析（排除 SWI-寄庫）")
    address_col = col_map.get("address")
    cust_name_col = col_map.get("cust_name")

    if not address_col or address_col not in df.columns:
        st.warning("找不到『地址』欄位。")
        return

    region_df = df[exclude_swi_mask].copy()
    region_df["縣市"] = region_df[address_col].apply(extract_city).fillna("(未知)")
    city_counts = region_df["縣市"].value_counts(dropna=False).rename_axis("縣市").reset_index(name="筆數")
    total_city = int(city_counts["筆數"].sum())
    city_counts["縣市佔比(%)"] = (city_counts["筆數"] / total_city * 100).round(2) if total_city > 0 else 0.0

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
        fig_city = px.bar(city_counts, x="縣市", y="筆數")
        st.plotly_chart(fig_city, use_container_width=True)

    if cust_name_col and cust_name_col in region_df.columns:
        cust_city = region_df.groupby([cust_name_col, "縣市"]).size().reset_index(name="筆數")
        cust_city["客戶總筆數"] = cust_city.groupby(cust_name_col)["筆數"].transform("sum")
        cust_city["縣市佔比(%)"] = (cust_city["筆數"] / cust_city["客戶總筆數"] * 100).round(2)
        top_table = cust_city.sort_values("筆數", ascending=False).groupby(cust_name_col).head(topn_cities)
        st.write(f"**各客戶常出貨縣市（每客戶 Top {topn_cities}）**")
        st.dataframe(top_table.rename(columns={cust_name_col: "客戶名稱"}), use_container_width=True)
        st.download_button(
            "下載各客戶常出貨縣市 CSV",
            data=top_table.to_csv(index=False).encode("utf-8-sig"),
            file_name="各客戶常出貨縣市_TopN.csv",
            mime="text/csv",
        )

def display_loading_analysis(df: pd.DataFrame, col_map: Dict[str, str | None], exclude_swi_mask: pd.Series):
    """④ 顯示配送裝載分析。"""
    st.subheader("④ 配送裝載分析（出庫單號前13碼=同車；僅完成）")
    ship_no_col = col_map.get("ship_no")
    status_col = col_map.get("status")
    
    if not ship_no_col or ship_no_col not in df.columns:
        st.warning("找不到『出庫單號』欄位。")
        return

    base_mask = exclude_swi_mask
    if status_col and status_col in df.columns:
        base_mask &= (df[status_col].astype(str).str.strip() == "完成")
    else:
        st.info("未對應到『通知單項次狀態』欄位，將不套用「完成」過濾。")

    load_df = df[base_mask].copy()
    if load_df.empty:
        st.info("根據篩選條件，無任何『完成』的出庫單資料。")
        return

    load_df["車次代碼"] = load_df[ship_no_col].astype(str).str[:13]
    load_df["_qty_num"] = pd.to_numeric(load_df[col_map.get("qty")], errors="coerce")
    load_df["_copper_ton_calc"] = pd.to_numeric(load_df[col_map.get("copper_ton")], errors="coerce") / 1000.0
    load_df["_fg_net_ton"] = pd.to_numeric(load_df[col_map.get("fg_net_ton")], errors="coerce")
    load_df["_fg_gross_ton"] = pd.to_numeric(load_df[col_map.get("fg_gross_ton")], errors="coerce")

    display_cols = {
        "車次代碼": load_df["車次代碼"],
        "出庫單號": load_df[ship_no_col],
        "DO號": load_df.get(col_map.get("do_no")),
        "料號說明": load_df.get(col_map.get("item_desc")),
        "批次": load_df.get(col_map.get("lot_no")),
        "出貨數量": load_df["_qty_num"],
        "銅重量(噸)": load_df["_copper_ton_calc"],
        "成品淨重(噸)": load_df["_fg_net_ton"],
        "成品毛重(噸)": load_df["_fg_gross_ton"],
    }
    
    display_df = pd.DataFrame({k: v for k, v in display_cols.items() if v is not None})
    for col in ["出貨數量", "銅重量(噸)", "成品淨重(噸)", "成品毛重(噸)"]:
        if col in display_df.columns:
            display_df[col] = pd.to_numeric(display_df[col], errors="coerce").round(3)
    
    sort_keys = [c for c in ["車次代碼", "出庫單號"] if c in display_df.columns]
    if sort_keys:
        display_df = display_df.sort_values(by=sort_keys).reset_index(drop=True)

    st.write("**配送裝載清單**")
    st.dataframe(display_df, use_container_width=True)
    st.download_button(
        "下載配送裝載清單 CSV",
        data=display_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="配送裝載清單.csv",
        mime="text/csv",
    )

    summary = load_df.groupby("車次代碼").agg(
        出貨數量小計=("_qty_num", "sum"),
        銅重量噸小計=("_copper_ton_calc", "sum"),
        成品淨重噸小計=("_fg_net_ton", "sum"),
        成品毛重噸小計=("_fg_gross_ton", "sum")
    ).round(3).reset_index()

    st.write("**車次彙總（每車小計）**")
    st.dataframe(summary, use_container_width=True)
    st.download_button(
        "下載車次彙總 CSV",
        data=summary.to_csv(index=False).encode("utf-8-sig"),
        file_name="車次彙總.csv",
        mime="text/csv",
    )

    if "銅重量噸小計" in summary.columns:
        valid_trips = summary[summary["銅重量噸小計"] > 0]
        avg_load = valid_trips["銅重量噸小計"].mean() if not valid_trips.empty else 0.0
        k1, k2 = st.columns(2)
        k1.metric("納入計算車次", f"{len(valid_trips):,}")
        k2.metric("平均車輛載重(噸)", f"{avg_load:.2f}")

# ========================
# 3. Streamlit 主流程
# ========================

file = st.file_uploader(
    "上傳 Excel 或 CSV 檔",
    type=["xlsx", "xls", "csv"],
    help="最多 200 MB；Excel 需使用 openpyxl 解析"
)

if file:
    file_ext = file.name.split(".")[-1].lower()
    df = load_data(file, file_ext)

    if not df.empty:
        col_map = guess_column_names(df.columns)
        
        tab_analysis, tab_raw = st.tabs(["📊 出貨&達交分析", "📄 原始資料預覽"])

        with tab_raw:
            st.subheader("原始資料預覽")
            st.dataframe(df, use_container_width=True)

        with tab_analysis:
            with st.sidebar:
                st.header("⚙️ 篩選操作區")
                st.info("您可以選擇多個條件來交叉分析數據。")
                
                # 使用 st.expander 將不常用篩選收合
                with st.expander("常用篩選", expanded=True):
                    # 常用篩選
                    sel_types = st.multiselect("出貨申請類型", options=sorted(df[col_map["ship_type"]].dropna().unique()) if col_map["ship_type"] else [], default=[])
                    sel_names = st.multiselect("客戶名稱", options=sorted(df[col_map["cust_name"]].dropna().unique()) if col_map["cust_name"] else [], default=[])
                    date_range = None
                    if col_map["due_date"]:
                        due_series = to_dt(df[col_map["due_date"]])
                        min_date, max_date = due_series.min().date(), due_series.max().date()
                        if min_date and max_date:
                            date_range = st.date_input("指定到貨日期（區間）", value=(min_date, max_date))
                    
                    st.slider_label = "每客戶 Top N 縣市"
                    topn_cities = st.slider(st.slider_label, min_value=1, max_value=10, value=3, step=1)
                
                with st.expander("進階篩選", expanded=False):
                    sel_ids = st.multiselect("客戶編號", options=sorted(df[col_map["cust_id"]].dropna().unique()) if col_map["cust_id"] else [], default=[])
                    sel_items = st.multiselect("料號說明", options=sorted(df[col_map["item_desc"]].dropna().unique()) if col_map["item_desc"] else [], default=[])
                    filter_col = st.selectbox("選擇額外篩選欄位", ["（不篩選）"] + list(df.columns))
                    extra_selected_vals = []
                    if filter_col != "（不篩選）":
                        extra_selected_vals = st.multiselect("選取值", df[filter_col].dropna().unique().tolist())
                
            filtered_df = df.copy()
            if sel_types:
                filtered_df = filtered_df[filtered_df[col_map["ship_type"]].isin(sel_types)]
            if sel_names:
                filtered_df = filtered_df[filtered_df[col_map["cust_name"]].isin(sel_names)]
            if sel_ids:
                filtered_df = filtered_df[filtered_df[col_map["cust_id"]].isin(sel_ids)]
            if sel_items:
                filtered_df = filtered_df[filtered_df[col_map["item_desc"]].isin(sel_items)]
            if date_range and len(date_range) == 2:
                start_d, end_d = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
                due_day = to_dt(filtered_df[col_map["due_date"]]).dt.normalize()
                filtered_df = filtered_df[(due_day >= start_d) & (due_day <= end_d)]
            if filter_col != "（不篩選）" and extra_selected_vals:
                filtered_df = filtered_df[filtered_df[filter_col].isin(extra_selected_vals)]
            
            # 共用：排除 SWI-寄庫
            exclude_swi_mask = (filtered_df[col_map["ship_type"]] != "SWI-寄庫") if col_map.get("ship_type") else pd.Series(True, index=filtered_df.index)
            
            # 顯示各分析區塊
            display_shipment_type_analysis(filtered_df.copy(), col_map)
            st.markdown("---")
            display_delivery_performance(filtered_df.copy(), col_map, exclude_swi_mask)
            st.markdown("---")
            display_area_analysis(filtered_df.copy(), col_map, exclude_swi_mask, topn_cities)
            st.markdown("---")
            display_loading_analysis(filtered_df.copy(), col_map, exclude_swi_mask)
            
    else:
        st.warning("檔案載入失敗或為空。")
else:
    st.info("請先在上方上傳 Excel 或 CSV 檔以開始分析。")
