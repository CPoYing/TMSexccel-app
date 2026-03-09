import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3

st.set_page_config(page_title="華新電線庫存追蹤", layout="wide", page_icon="📊")

# 自訂簡單主題（讓畫面更專業）
st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    .metric-box { background: white; padding: 1rem; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    h1, h2, h3 { color: #1f2937; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 華新電線庫存歷史追蹤系統")
st.caption("每日自動更新 • 支援多品項比較 • 低庫存即時警示")

conn = sqlite3.connect('inventory.db')
df = pd.read_sql_query("SELECT * FROM history", conn)
conn.close()

if df.empty:
    st.warning("資料庫目前無資料，請先執行 scraper.py")
    st.stop()

df["日期"] = pd.to_datetime(df["日期"])

# ── 側邊欄篩選 ──
with st.sidebar:
    st.header("篩選")
    selected_codes = st.multiselect("貨品代號（可多選）", options=sorted(df["貨品代號"].unique()), default=[])
    date_range = st.date_input("日期範圍", [df["日期"].min().date(), df["日期"].max().date()])
    min_stock_alert = st.number_input("低庫存警示門檻（米）", value=1000, min_value=100, step=100)

    filtered = df[
        (df["日期"] >= pd.to_datetime(date_range[0])) &
        (df["日期"] <= pd.to_datetime(date_range[1]))
    ]
    if selected_codes:
        filtered = filtered[filtered["貨品代號"].isin(selected_codes)]

# ── KPI 區 ──
cols = st.columns(4)
latest_date = filtered["日期"].max().strftime("%Y-%m-%d")
total_latest = filtered[filtered["日期"] == filtered["日期"].max()]["現有數量"].sum()
days_tracked = filtered["日期"].nunique()
items_tracked = filtered["貨品代號"].nunique()

with cols[0]:
    st.metric("最新更新", latest_date, delta=None)
with cols[1]:
    st.metric("總庫存量（最新）", f"{total_latest:,} 米")
with cols[2]:
    st.metric("追蹤天數", days_tracked)
with cols[3]:
    st.metric("追蹤品項數", items_tracked)

# ── 分頁 ──
tab1, tab2, tab3, tab4 = st.tabs(["總趨勢", "單品分析", "低庫存警報", "原始資料"])

with tab1:
    st.subheader("全域總庫存趨勢")
    daily_sum = filtered.groupby("日期")["現有數量"].sum().reset_index()
    fig_total = px.line(daily_sum, x="日期", y="現有數量",
                        title="總庫存量變化（可縮放）",
                        markers=True, line_shape="spline")
    fig_total.update_layout(height=500, hovermode="x unified")
    st.plotly_chart(fig_total, use_container_width=True)

with tab2:
    st.subheader("選定品項庫存變化比較")
    if not selected_codes:
        st.info("請在側邊欄選擇至少一個貨品代號")
    else:
        fig_detail = px.line(filtered, x="日期", y="現有數量",
                             color="貨品名稱", symbol="庫位",
                             title="多規格/多庫位疊加趨勢",
                             markers=True)
        fig_detail.update_traces(line=dict(width=2.5))
        fig_detail.update_layout(height=600, hovermode="x unified", legend=dict(orientation="h"))
        st.plotly_chart(fig_detail, use_container_width=True)

with tab3:
    st.subheader(f"低庫存警示（最新日 < {min_stock_alert} 米）")
    latest_df = filtered[filtered["日期"] == filtered["日期"].max()]
    low_stock = latest_df[latest_df["現有數量"] < min_stock_alert].sort_values("現有數量")
    if low_stock.empty:
        st.success("目前無低庫存品項")
    else:
        st.dataframe(
            low_stock.style
               .format({"現有數量": "{:,}"})
               .background_gradient(subset=["現有數量"], cmap="OrRd_r", low=0.8, high=0),
            use_container_width=True, hide_index=True
        )

with tab4:
    st.subheader("完整歷史資料")
    st.dataframe(filtered.sort_values(["日期", "貨品代號"], ascending=[False, True]),
                 use_container_width=True, hide_index=True)
    
    csv = filtered.to_csv(index=False).encode('utf-8-sig')
    st.download_button("下載 CSV（含 BOM）", csv, f"華新庫存_{latest_date}.csv", "text/csv")

st.caption("系統自動更新時間：每日執行 scraper.py 後刷新頁面即可看到最新資料")
