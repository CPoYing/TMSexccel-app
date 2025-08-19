import io
import pandas as pd
import streamlit as st
import plotly.express as px
import re

st.set_page_config(page_title="TMS出貨配送數據分析", page_icon="📊", layout="wide")
st.title("📊 TMS出貨配送數據分析")
st.caption("上傳 Excel/CSV → 出貨類型筆數 → 達交率（日期比較）→ 配送區域分析 → 自訂欄位 → 聚合 → 圖表 → 下載")

# ---------- 檔案上傳 ----------
file = st.file_uploader("上傳 Excel 或 CSV 檔", type=["xlsx", "xls", "csv"],
                        help="最多 200 MB；Excel 需使用 openpyxl 解析")

@st.cache_data
def load_data(file):
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file, engine="openpyxl")
    return df

def _guess_col(cols, keywords):
    for kw in keywords:
        for c in cols:
            if kw in str(c):
                return c
    return None

def to_dt(series):
    return pd.to_datetime(series, errors="coerce")

def extract_city(addr: object) -> str | None:
    if pd.isna(addr):
        return None
    s = str(addr).strip().replace("臺", "台")
    pattern = r"(台北市|新北市|桃園市|台中市|台南市|高雄市|基隆市|新竹市|新竹縣|苗栗縣|彰化縣|南投縣|雲林縣|嘉義市|嘉義縣|屏東縣|宜蘭縣|花蓮縣|台東縣|澎湖縣|金門縣|連江縣)"
    m = re.search(pattern, s)
    return m.group(1) if m else None

if file:
    df = load_data(file)

    # Tabs
    tab_analysis, tab_raw = st.tabs(["📊 出貨&達交分析", "📄 原始資料預覽"])

    with tab_raw:
        st.subheader("原始資料預覽")
        st.dataframe(df, use_container_width=True)

    with tab_analysis:
        # ---------- 側欄操作 ----------
        with st.sidebar:
            st.header("⚙️ 操作區")
            cols = list(df.columns)

            filter_col = st.selectbox("選擇要篩選的欄位（可略）", ["（不篩選）"] + cols)
            filtered = df.copy()
            if filter_col != "（不篩選）":
                unique_vals = filtered[filter_col].dropna().unique().tolist()
                selected_vals = st.multiselect("選取值", unique_vals)
                if selected_vals:
                    filtered = filtered[filtered[filter_col].isin(selected_vals)]

            st.markdown("—")
            st.subheader("欄位對應")
            ship_type_col = st.selectbox("出貨類型欄位", options=cols)
            due_date_col = st.selectbox("指定到貨日期欄位", options=cols)
            sign_date_col = st.selectbox("客戶簽收日期欄位", options=cols)
            cust_id_col = st.selectbox("客戶編號欄位", options=cols)
            cust_name_col = st.selectbox("客戶名稱欄位", options=cols)
            address_col = st.selectbox("地址欄位", options=cols)

        data = filtered.copy()

        # ---------- 功能① 出貨類型筆數 ----------
        st.subheader("① 出貨類型筆數統計")
        type_counts = (
            data[ship_type_col].fillna("(空白)").value_counts(dropna=False)
            .rename_axis("出貨類型").reset_index(name="筆數")
        )
        total_rows = int(type_counts["筆數"].sum())
        st.write(f"**加總筆數：{total_rows:,}**")
        st.dataframe(type_counts, use_container_width=True)

        # ---------- 功能② 達交率 ----------
        st.subheader("② 達交率（僅比對日期，不含時分秒）")
        due_dt = to_dt(data[due_date_col])
        sign_dt = to_dt(data[sign_date_col])
        due_day = due_dt.dt.normalize()
        sign_day = sign_dt.dt.normalize()

        exclude_mask = data[ship_type_col] != "SWI-寄庫"
        valid_mask = due_day.notna() & sign_day.notna() & exclude_mask
        on_time = (sign_day <= due_day) & valid_mask

        total_valid = int(valid_mask.sum())
        on_time_count = int(on_time.sum())
        rate = (on_time_count / total_valid * 100.0) if total_valid > 0 else 0.0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("有效筆數", f"{total_valid:,}")
        k2.metric("準時交付筆數", f"{on_time_count:,}")
        k3.metric("達交率", f"{rate:.2f}%")

        late_mask = valid_mask & (sign_day > due_day)
        delay_days = (sign_day - due_day).dt.days
        late_df = pd.DataFrame({
            "客戶編號": data[cust_id_col],
            "客戶名稱": data[cust_name_col],
            "指定到貨日期": due_day.dt.strftime("%Y-%m-%d"),
            "客戶簽收日期": sign_day.dt.strftime("%Y-%m-%d"),
            "延遲天數": delay_days,
        })[late_mask]
        avg_delay = float(late_df["延遲天數"].mean()) if not late_df.empty else 0.0
        k4.metric("平均延遲天數", f"{avg_delay:.2f}")
        st.dataframe(late_df, use_container_width=True)

        # ---------- 功能③ 配送區域分析 ----------
        st.subheader("③ 配送區域分析（排除 SWI-寄庫）")
        region_df = data[exclude_mask].copy()
        region_df["縣市"] = region_df[address_col].apply(extract_city).fillna("(未知)")

        # 依縣市統計
        city_counts = region_df["縣市"].value_counts().rename_axis("縣市").reset_index(name="筆數")
        st.write("**各縣市配送筆數**")
        st.dataframe(city_counts, use_container_width=True)

        # 客戶常出的縣市
        cust_city = region_df.groupby([cust_name_col, "縣市"]).size().reset_index(name="筆數")
        cust_city = cust_city.sort_values([cust_name_col, "筆數"], ascending=[True, False])
        top_table = cust_city.groupby(cust_name_col).head(3)
        st.write("**各客戶常出貨縣市（Top 3）**")
        st.dataframe(top_table.rename(columns={cust_name_col: "客戶名稱"}), use_container_width=True)

else:
    st.info("請先在上方上傳 Excel 或 CSV 檔。")
