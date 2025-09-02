import io
import re
from typing import Dict, Optional, Tuple
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# 依賴檢查
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# ========================
# 頁面設定與 CSS 樣式
# ========================

st.set_page_config(
    page_title="TMS出貨配送數據分析",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 素雅色彩配置
ELEGANT_COLORS = {
    'primary': '#6c7b7f',      # 優雅灰藍
    'secondary': '#a8b5b8',    # 淺灰藍
    'success': '#8fbc8f',      # 柔和綠
    'warning': '#daa520',      # 古金色
    'danger': '#cd5c5c',       # 印度紅
    'light_gray': '#f5f7fa',   # 淺灰
    'dark_gray': '#2c3e50'     # 深灰
}

# 彩色調色板
COLORFUL_PALETTE = [
    '#DDA0DD',  # 淡紫色
    '#87CEEB',  # 淡藍色
    '#98FB98',  # 淡綠色
    '#F0E68C',  # 淡金色
    '#FFB6C1',  # 淡粉色
    '#DEB887',  # 淡棕色
    '#E0E0E0',  # 淡灰色
    '#FFA07A',  # 淡橙色
    '#20B2AA',  # 淡青色
    '#D8BFD8',  # 薊色
    '#AFEEEE',  # 淡藍綠色
    '#F5DEB3'   # 小麥色
]

def format_number_with_comma(value, decimal_places=0):
    """格式化數字：千分位 + 靠右對齊"""
    if pd.isna(value):
        return ""
    try:
        if decimal_places == 0:
            return f"{int(value):,}"
        else:
            return f"{float(value):,.{decimal_places}f}"
    except:
        return str(value)

# 自定義 CSS 樣式
st.markdown(f"""
<style>
    .main-header {{
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(90deg, {ELEGANT_COLORS['primary']}, {ELEGANT_COLORS['secondary']});
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 1rem;
    }}
    
    .metric-container {{
        background: {ELEGANT_COLORS['light_gray']};
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
        border-left: 4px solid {ELEGANT_COLORS['primary']};
        margin: 0.5rem 0;
    }}
    
    .section-header {{
        color: {ELEGANT_COLORS['dark_gray']};
        border-bottom: 2px solid {ELEGANT_COLORS['secondary']};
        padding-bottom: 0.5rem;
        margin: 1.5rem 0 1rem 0;
    }}
</style>
""", unsafe_allow_html=True)

# 主標題
st.markdown('<h1 class="main-header">📊 TMS出貨配送數據分析</h1>', unsafe_allow_html=True)
st.markdown("""
<div style="text-align: center; color: #7f8c8d; margin-bottom: 2rem;">
    <strong>智能化數據分析平台</strong> | 上傳 Excel/CSV → 出貨類型分析 → 配送達交率 → 區域分析 → 裝載分析
</div>
""", unsafe_allow_html=True)

# ========================
# 工具函式優化
# ========================

@st.cache_data(show_spinner="正在載入檔案...")
def load_data(file_buffer: io.BytesIO, file_type: str) -> pd.DataFrame:
    """高效能檔案載入與快取"""
    try:
        if file_type == "csv":
            encodings = ['utf-8', 'utf-8-sig', 'big5', 'gbk']
            for encoding in encodings:
                try:
                    file_buffer.seek(0)
                    return pd.read_csv(file_buffer, encoding=encoding)
                except UnicodeDecodeError:
                    continue
            file_buffer.seek(0)
            return pd.read_csv(file_buffer, encoding='utf-8', errors='ignore')
        elif file_type in ["xlsx", "xls"]:
            if not EXCEL_AVAILABLE:
                st.error("需要安裝 openpyxl 來處理 Excel 檔案")
                return pd.DataFrame()
            file_buffer.seek(0)
            return pd.read_excel(file_buffer, engine="openpyxl")
    except Exception as e:
        st.error(f"檔案載入失敗: {str(e)}")
        return pd.DataFrame()
    
    st.error("不支援的檔案類型")
    return pd.DataFrame()

def smart_column_detector(df_cols: list) -> Dict[str, Optional[str]]:
    """智能欄位偵測器"""
    def find_best_match(keywords: list, priority_keywords: list = None) -> Optional[str]:
        if priority_keywords:
            for kw in priority_keywords:
                for col in df_cols:
                    if kw.lower() in str(col).lower():
                        return col
        
        for kw in keywords:
            for col in df_cols:
                if kw.lower() in str(col).lower():
                    return col
        return None

    return {
        "ship_type": find_best_match(["出貨申請類型", "出貨類型", "配送類型", "類型"], ["出貨申請類型"]),
        "due_date": find_best_match(["指定到貨日期", "到貨日期", "指定到貨"], ["指定到貨日期"]),
        "sign_date": find_best_match(["客戶簽收日期", "簽收日期", "簽收日"], ["客戶簽收日期"]),
        "cust_id": find_best_match(["客戶編號", "客戶代號", "客編"], ["客戶編號"]),
        "cust_name": find_best_match(["客戶名稱", "客名", "客戶"], ["客戶名稱"]),
        "address": find_best_match(["地址", "收貨地址", "送貨地址"], ["地址"]),
        "copper_ton": find_best_match(["銅重量(噸)", "銅重量", "銅噸"], ["銅重量(噸)"]),
        "qty": find_best_match(["出貨數量", "數量", "出貨量"], ["出貨數量"]),
        "ship_no": find_best_match(["出庫單號", "出庫單", "出庫編號"], ["出庫單號"]),
        "do_no": find_best_match(["DO號", "DO", "出貨單號"], ["DO號"]),
        "item_desc": find_best_match(["料號說明", "品名", "品名規格"], ["料號說明"]),
        "lot_no": find_best_match(["批次", "批號", "Lot"], ["批次"]),
        "fg_net_ton": find_best_match(["成品淨重(噸)", "淨重(噸)", "淨重"], ["成品淨重(噸)"]),
        "fg_gross_ton": find_best_match(["成品毛重(噸)", "毛重(噸)", "毛重"], ["成品毛重(噸)"]),
        "status": find_best_match(["通知單項次狀態", "項次狀態", "狀態"], ["通知單項次狀態"]),
    }

def safe_datetime_convert(series: pd.Series) -> pd.Series:
    """安全的日期時間轉換"""
    return pd.to_datetime(series, errors="coerce")

def enhanced_city_extraction(address: str) -> Optional[str]:
    """增強版台灣縣市擷取"""
    if pd.isna(address) or not address:
        return None
    
    normalized_addr = str(address).strip().replace("臺", "台")
    taiwan_cities = [
        "台北市", "新北市", "桃園市", "台中市", "台南市", "高雄市",
        "基隆市", "新竹市", "新竹縣", "苗栗縣", "彰化縣", "南投縣",
        "雲林縣", "嘉義市", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣",
        "台東縣", "澎湖縣", "金門縣", "連江縣"
    ]
    
    for city in taiwan_cities:
        if city in normalized_addr:
            return city
    return None

def create_enhanced_metric_card(title: str, value: str, delta: str = None, delta_color: str = "normal"):
    """建立增強版指標卡片"""
    delta_html = ""
    if delta:
        color = {"normal": "#666", "positive": ELEGANT_COLORS['success'], "negative": ELEGANT_COLORS['danger']}.get(delta_color, "#666")
        delta_html = f'<div style="color: {color}; font-size: 0.9rem; margin-top: 0.5rem;">{delta}</div>'
    
    st.markdown(f"""
    <div class="metric-container">
        <div style="color: #666; font-size: 0.9rem; margin-bottom: 0.3rem;">{title}</div>
        <div style="font-size: 2rem; font-weight: bold; color: {ELEGANT_COLORS['dark_gray']};">{value}</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)

# ========================
# 核心分析函式
# ========================

def enhanced_shipment_analysis(df: pd.DataFrame, col_map: Dict[str, Optional[str]]) -> None:
    """增強版出貨類型分析"""
    st.markdown('<h2 class="section-header">① 出貨類型筆數與重量統計</h2>', unsafe_allow_html=True)
    
    ship_type_col = col_map.get("ship_type")
    copper_ton_col = col_map.get("copper_ton")
    
    if not ship_type_col or ship_type_col not in df.columns:
        st.error("⚠️ 找不到『出貨申請類型』欄位")
        return
    
    # 處理出貨類型統計
    df_work = df.copy()
    df_work[ship_type_col] = df_work[ship_type_col].fillna("(空白)")
    
    type_stats = df_work[ship_type_col].value_counts().reset_index()
    type_stats.columns = ["出貨類型", "筆數"]
    
    # 計算銅重量統計
    final_stats = type_stats.copy()
    if copper_ton_col and copper_ton_col in df.columns:
        df_work["_copper_numeric"] = pd.to_numeric(df_work[copper_ton_col], errors="coerce")
        
        copper_summary = []
        for ship_type in type_stats["出貨類型"]:
            mask = df_work[ship_type_col] == ship_type
            copper_sum = df_work[mask]["_copper_numeric"].sum()
            copper_summary.append(copper_sum / 1000.0)  # 轉為噸
        
        final_stats["銅重量(噸)合計"] = copper_summary
    else:
        final_stats["銅重量(噸)合計"] = None
    
    # 計算百分比
    total_records = final_stats["筆數"].sum()
    final_stats["佔比(%)"] = (final_stats["筆數"] / total_records * 100).round(2)
    
    # 顯示關鍵指標
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        create_enhanced_metric_card("總資料筆數", f"{total_records:,}")
    with col2:
        unique_types = len(final_stats)
        create_enhanced_metric_card("出貨類型數", f"{unique_types}")
    with col3:
        total_copper = final_stats["銅重量(噸)合計"].sum() if "銅重量(噸)合計" in final_stats.columns and final_stats["銅重量(噸)合計"].notna().any() else 0
        create_enhanced_metric_card("總銅重量(噸)", f"{total_copper:,.3f}")
    with col4:
        avg_per_record = total_copper / total_records if total_records > 0 else 0
        create_enhanced_metric_card("平均每筆重量(噸)", f"{avg_per_record:.3f}")
    
    # 顯示詳細統計表格與圖表
    tab1, tab2 = st.tabs(["📋 詳細統計", "📊 視覺化圖表"])
    
    with tab1:
        st.dataframe(final_stats, use_container_width=True, height=400)
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "📥 下載統計表 (CSV)",
                data=final_stats.to_csv(index=False).encode("utf-8-sig"),
                file_name="出貨類型統計.csv",
                mime="text/csv",
            )
    
    with tab2:
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            chart_type = st.selectbox("圖表類型", ["長條圖", "圓餅圖", "環形圖", "水平長條圖"], key="chart_type_1")
            
            # 素雅色彩配置
            elegant_palette = ['#6c7b7f', '#a8b5b8', '#95a5a6', '#7f8c8d', '#bdc3c7', '#ecf0f1']
            
            if chart_type == "長條圖":
                fig = px.bar(final_stats, x="出貨類型", y="筆數", title="出貨類型筆數分佈")
                fig.update_traces(marker_color=ELEGANT_COLORS['primary'])
                fig.update_layout(xaxis_tickangle=-45)
            elif chart_type == "水平長條圖":
                fig = px.bar(final_stats, x="筆數", y="出貨類型", orientation='h', title="出貨類型筆數分佈")
                fig.update_traces(marker_color=ELEGANT_COLORS['secondary'])
            elif chart_type == "圓餅圖":
                fig = px.pie(final_stats, names="出貨類型", values="筆數", title="出貨類型佔比分佈", color_discrete_sequence=elegant_palette)
                fig.update_traces(textposition='inside', textinfo='percent+label')
            else:  # 環形圖
                fig = px.pie(final_stats, names="出貨類型", values="筆數", title="出貨類型佔比分佈", hole=0.4, color_discrete_sequence=elegant_palette)
                fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)
        
        with chart_col2:
            if "銅重量(噸)合計" in final_stats.columns and final_stats["銅重量(噸)合計"].notna().any():
                fig_copper = px.treemap(
                    final_stats[final_stats["銅重量(噸)合計"].notna()],
                    path=["出貨類型"],
                    values="銅重量(噸)合計",
                    title="銅重量分佈樹狀圖",
                    color_discrete_sequence=COLORFUL_PALETTE
                )
                st.plotly_chart(fig_copper, use_container_width=True)
            else:
                st.info("無銅重量數據可供分析")

def enhanced_delivery_performance(df: pd.DataFrame, col_map: Dict[str, Optional[str]], exclude_swi_mask: pd.Series) -> None:
    """增強版配送達交率分析"""
    st.markdown('<h2 class="section-header">② 配送達交率分析</h2>', unsafe_allow_html=True)
    
    due_date_col = col_map.get("due_date")
    sign_date_col = col_map.get("sign_date")
    cust_name_col = col_map.get("cust_name")
    
    if not all([due_date_col, sign_date_col]) or not all([col in df.columns for col in [due_date_col, sign_date_col]]):
        st.error("⚠️ 找不到必要的日期欄位")
        return
    
    # 日期處理
    due_day = safe_datetime_convert(df[due_date_col]).dt.normalize()
    sign_day = safe_datetime_convert(df[sign_date_col]).dt.normalize()
    
    # 有效性檢查
    valid_mask = due_day.notna() & sign_day.notna() & exclude_swi_mask
    on_time_mask = (sign_day <= due_day) & valid_mask
    late_mask = valid_mask & (sign_day > due_day)
    
    # 統計計算
    total_valid = int(valid_mask.sum())
    on_time_count = int(on_time_mask.sum())
    late_count = int(late_mask.sum())
    
    if total_valid == 0:
        st.warning("無有效的達交分析數據")
        return
    
    delivery_rate = on_time_count / total_valid * 100
    delay_days = (sign_day - due_day).dt.days
    avg_delay = delay_days[late_mask].mean() if late_count > 0 else 0.0
    max_delay = delay_days[late_mask].max() if late_count > 0 else 0
    
    # 關鍵指標展示
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        create_enhanced_metric_card("有效筆數", f"{total_valid:,}")
    with col2:
        create_enhanced_metric_card("準時交付", f"{on_time_count:,}", f"({delivery_rate:.1f}%)", "positive")
    with col3:
        create_enhanced_metric_card("延遲交付", f"{late_count:,}", f"({100-delivery_rate:.1f}%)", "negative")
    with col4:
        status = "優異" if delivery_rate >= 95 else "良好" if delivery_rate >= 90 else "需改善"
        color = "positive" if delivery_rate >= 95 else "normal" if delivery_rate >= 90 else "negative"
        create_enhanced_metric_card("達交率", f"{delivery_rate:.2f}%", status, color)
    with col5:
        create_enhanced_metric_card("平均延遲", f"{avg_delay:.1f}天", f"最大: {max_delay}天" if max_delay > 0 else "")
    
    # 詳細分析
    analysis_tab1, analysis_tab2, analysis_tab3 = st.tabs(["🚚 未達標明細", "📈 達交趨勢", "👥 客戶表現"])
    
    with analysis_tab1:
        if late_count > 0:
            late_indices = df.index[late_mask]
            late_details = pd.DataFrame({
                "指定到貨日期": [due_day.loc[i].strftime("%Y-%m-%d") if pd.notna(due_day.loc[i]) else "" for i in late_indices],
                "客戶簽收日期": [sign_day.loc[i].strftime("%Y-%m-%d") if pd.notna(sign_day.loc[i]) else "" for i in late_indices],
                "延遲天數": [delay_days.loc[i] if pd.notna(delay_days.loc[i]) else 0 for i in late_indices]
            })
            
            if col_map.get("cust_name") and col_map["cust_name"] in df.columns:
                late_details["客戶名稱"] = [df.loc[i, col_map["cust_name"]] for i in late_indices]
            
            # 延遲程度分類
            late_details["延遲程度"] = pd.cut(
                late_details["延遲天數"],
                bins=[-1, 0, 3, 7, 30, 999],
                labels=["準時", "輕微延遲(1-3天)", "中度延遲(4-7天)", "重度延遲(8-30天)", "嚴重延遲(>30天)"]
            )
            
            late_details = late_details.sort_values("延遲天數", ascending=False)
            
            # 格式化延遲天數顯示
            display_late_details = late_details.copy()
            display_late_details["延遲天數"] = display_late_details["延遲天數"].apply(lambda x: format_number_with_comma(x))
            
            st.dataframe(display_late_details, use_container_width=True, height=400)
            
            # 延遲分佈圖 - 彩色版
            delay_dist = late_details["延遲程度"].value_counts()
            fig_delay = px.bar(
                x=delay_dist.index, 
                y=delay_dist.values,
                title="延遲程度分佈",
                labels={"x": "延遲程度", "y": "筆數"},
                color_discrete_sequence=COLORFUL_PALETTE
            )
            st.plotly_chart(fig_delay, use_container_width=True)
            
            st.download_button(
                "📥 下載未達標明細",
                data=late_details.to_csv(index=False).encode("utf-8-sig"),
                file_name="未達標明細.csv",
                mime="text/csv",
            )
        else:
            st.success("🎉 恭喜！所有出貨都準時達交！")
    
    with analysis_tab2:
        if total_valid >= 7:
            # 建立趨勢數據
            valid_indices = df.index[valid_mask]
            trend_data = []
            
            for idx in valid_indices:
                if pd.notna(due_day.loc[idx]):
                    trend_data.append({
                        "到貨日期": due_day.loc[idx].date(),
                        "準時": 1 if on_time_mask.loc[idx] else 0
                    })
            
            if trend_data:
                trend_df = pd.DataFrame(trend_data)
                daily_trend = trend_df.groupby("到貨日期").agg({
                    "準時": ["count", "sum"]
                }).reset_index()
                daily_trend.columns = ["到貨日期", "總筆數", "準時筆數"]
                daily_trend["達交率(%)"] = (daily_trend["準時筆數"] / daily_trend["總筆數"] * 100).round(2)
                
                fig_trend = px.line(
                    daily_trend, 
                    x="到貨日期", 
                    y="達交率(%)",
                    title="每日達交率趨勢",
                    markers=True,
                    color_discrete_sequence=COLORFUL_PALETTE
                )
                fig_trend.update_traces(line_color=COLORFUL_PALETTE[0], marker_color=COLORFUL_PALETTE[1])
                fig_trend.add_hline(y=95, line_dash="dash", line_color=COLORFUL_PALETTE[2], annotation_text="目標線(95%)")
                fig_trend.add_hline(y=90, line_dash="dash", line_color=COLORFUL_PALETTE[3], annotation_text="警戒線(90%)")
                
                st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.info("數據量不足，無法分析趨勢（需至少7筆記錄）")
    
    with analysis_tab3:
        if cust_name_col and cust_name_col in df.columns:
            # 客戶表現分析
            valid_indices = df.index[valid_mask]
            customer_data = []
            
            for idx in valid_indices:
                customer_data.append({
                    "客戶名稱": df.loc[idx, cust_name_col],
                    "準時": 1 if on_time_mask.loc[idx] else 0,
                    "總計": 1
                })
            
            if customer_data:
                customer_df = pd.DataFrame(customer_data)
                customer_performance = customer_df.groupby("客戶名稱").agg({
                    "總計": "sum",
                    "準時": "sum"
                }).reset_index()
                
                customer_performance.columns = ["客戶名稱", "有效筆數", "準時筆數"]
                customer_performance["延遲筆數"] = customer_performance["有效筆數"] - customer_performance["準時筆數"]
                customer_performance["達交率(%)"] = (customer_performance["準時筆數"] / customer_performance["有效筆數"] * 100).round(2)
                customer_performance = customer_performance.sort_values("達交率(%)", ascending=False)
                
                # 格式化客戶表現數據
                display_customer_performance = customer_performance.copy()
                for col in ["有效筆數", "準時筆數", "延遲筆數"]:
                    if col in display_customer_performance.columns:
                        display_customer_performance[col] = display_customer_performance[col].apply(lambda x: format_number_with_comma(x))
                display_customer_performance["達交率(%)"] = display_customer_performance["達交率(%)"].apply(lambda x: f"{x:.2f}%")
                
                st.dataframe(display_customer_performance, use_container_width=True)
                
                st.download_button(
                    "📥 下載客戶表現分析",
                    data=customer_performance.to_csv(index=False).encode("utf-8-sig"),
                    file_name="客戶達交表現.csv",
                    mime="text/csv",
                )

def enhanced_area_analysis(df: pd.DataFrame, col_map: Dict[str, Optional[str]], exclude_swi_mask: pd.Series, topn_cities: int) -> None:
    """增強版配送區域分析"""
    st.markdown('<h2 class="section-header">③ 配送區域分析</h2>', unsafe_allow_html=True)
    
    address_col = col_map.get("address")
    cust_name_col = col_map.get("cust_name")
    
    if not address_col or address_col not in df.columns:
        st.error("⚠️ 找不到『地址』欄位")
        return
    
    # 處理區域資料
    region_df = df[exclude_swi_mask].copy()
    region_df["縣市"] = region_df[address_col].apply(enhanced_city_extraction)
    region_df["縣市"] = region_df["縣市"].fillna("(未知)")
    
    city_stats = region_df["縣市"].value_counts().reset_index()
    city_stats.columns = ["縣市", "筆數"]
    
    total_records = city_stats["筆數"].sum()
    city_stats["佔比(%)"] = (city_stats["筆數"] / total_records * 100).round(2)
    
    # 區域指標
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        create_enhanced_metric_card("總配送筆數", f"{total_records:,}")
    with col2:
        valid_cities = len(city_stats[city_stats["縣市"] != "(未知)"])
        create_enhanced_metric_card("覆蓋縣市數", f"{valid_cities}")
    with col3:
        main_city = city_stats.iloc[0]["縣市"] if not city_stats.empty else "無"
        main_count = city_stats.iloc[0]["筆數"] if not city_stats.empty else 0
        create_enhanced_metric_card("主要配送區", main_city, f"{main_count}筆")
    with col4:
        unknown_count = city_stats[city_stats["縣市"] == "(未知)"]["筆數"].sum()
        unknown_rate = unknown_count / total_records * 100 if total_records > 0 else 0
        create_enhanced_metric_card("地址識別率", f"{100-unknown_rate:.1f}%")
    
    # 詳細分析
    tab1, tab2 = st.tabs(["🗺️ 縣市分佈", "🏢 客戶區域"])
    
    with tab1:
        col1, col2 = st.columns([1, 1])
        with col1:
            # 格式化縣市統計數據
            display_city_stats = city_stats.copy()
            display_city_stats["筆數"] = display_city_stats["筆數"].apply(lambda x: format_number_with_comma(x))
            display_city_stats["佔比(%)"] = display_city_stats["佔比(%)"].apply(lambda x: f"{x:.2f}%")
            
            st.dataframe(display_city_stats, use_container_width=True, height=400)
        with col2:
            fig_map = px.bar(
                city_stats.head(15), 
                x="縣市", 
                y="筆數",
                title="前15縣市配送量",
                color_discrete_sequence=COLORFUL_PALETTE
            )
            fig_map.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_map, use_container_width=True)
        
        st.download_button(
            "📥 下載縣市統計",
            data=city_stats.to_csv(index=False).encode("utf-8-sig"),
            file_name="縣市配送統計.csv",
            mime="text/csv",
        )
    
    with tab2:
        if cust_name_col and cust_name_col in region_df.columns:
            customer_city_analysis = region_df.groupby([cust_name_col, "縣市"]).size().reset_index(name="筆數")
            customer_city_analysis["客戶總筆數"] = customer_city_analysis.groupby(cust_name_col)["筆數"].transform("sum")
            customer_city_analysis["縣市佔比(%)"] = (
                customer_city_analysis["筆數"] / customer_city_analysis["客戶總筆數"] * 100
            ).round(2)
            
            # 取每客戶前N個縣市
            top_customer_cities = (
                customer_city_analysis
                .sort_values("筆數", ascending=False)
                .groupby(cust_name_col)
                .head(topn_cities)
                .sort_values([cust_name_col, "筆數"], ascending=[True, False])
            )
            
            # 格式化客戶區域數據
            display_cities = top_customer_cities.rename(columns={cust_name_col: "客戶名稱"}).copy()
            for col in ["筆數", "客戶總筆數"]:
                if col in display_cities.columns:
                    display_cities[col] = display_cities[col].apply(lambda x: format_number_with_comma(x))
            display_cities["縣市佔比(%)"] = display_cities["縣市佔比(%)"].apply(lambda x: f"{x:.2f}%")
            
            st.dataframe(display_cities, use_container_width=True)
            
            st.download_button(
                "📥 下載客戶區域分析",
                data=top_customer_cities.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"客戶區域分析_Top{topn_cities}.csv",
                mime="text/csv",
            )

def enhanced_loading_analysis(df: pd.DataFrame, col_map: Dict[str, Optional[str]], exclude_swi_mask: pd.Series) -> None:
    """增強版配送裝載分析"""
    st.markdown('<h2 class="section-header">④ 配送裝載分析</h2>', unsafe_allow_html=True)
    
    ship_no_col = col_map.get("ship_no")
    status_col = col_map.get("status")
    
    if not ship_no_col or ship_no_col not in df.columns:
        st.error("⚠️ 找不到『出庫單號』欄位")
        return
    
    # 篩選條件
    base_mask = exclude_swi_mask.copy()
    if status_col and status_col in df.columns:
        completed_mask = df[status_col].astype(str).str.strip() == "完成"
        base_mask &= completed_mask
        st.info("✅ 已套用狀態篩選：僅顯示『完成』狀態的記錄")
    
    loading_df = df[base_mask].copy()
    
    if loading_df.empty:
        st.warning("無符合條件的裝載數據")
        return
    
    # 車次代碼生成
    loading_df["車次代碼"] = loading_df[ship_no_col].astype(str).str[:13]
    
    # 數值欄位處理
    numeric_cols = ["qty", "copper_ton", "fg_net_ton", "fg_gross_ton"]
    for col_key in numeric_cols:
        col_name = col_map.get(col_key)
        if col_name and col_name in loading_df.columns:
            loading_df[f"_{col_key}_numeric"] = pd.to_numeric(loading_df[col_name], errors="coerce")
    
    # 轉換銅重量為噸
    if "_copper_ton_numeric" in loading_df.columns:
        loading_df["_copper_ton_converted"] = loading_df["_copper_ton_numeric"] / 1000.0
    
    # 建立展示資料
    display_data = {"車次代碼": loading_df["車次代碼"], "出庫單號": loading_df[ship_no_col]}
    
    # 可選欄位
    optional_mappings = {
        "DO號": "do_no",
        "料號說明": "item_desc", 
        "批次": "lot_no",
        "客戶名稱": "cust_name"
    }
    
    for display_name, map_key in optional_mappings.items():
        col_name = col_map.get(map_key)
        if col_name and col_name in loading_df.columns:
            display_data[display_name] = loading_df[col_name]
    
    # 數值欄位
    numeric_mappings = {
        "出貨數量": "_qty_numeric",
        "銅重量(噸)": "_copper_ton_converted", 
        "成品淨重(噸)": "_fg_net_ton_numeric",
        "成品毛重(噸)": "_fg_gross_ton_numeric"
    }
    
    for display_name, internal_name in numeric_mappings.items():
        if internal_name in loading_df.columns:
            display_data[display_name] = loading_df[internal_name].round(3)
    
    detailed_loading_df = pd.DataFrame(display_data)
    
    # 排序
    if "車次代碼" in detailed_loading_df.columns:
        detailed_loading_df = detailed_loading_df.sort_values(["車次代碼", "出庫單號"]).reset_index(drop=True)
    
    # 車次彙總
    trip_summary_data = []
    for trip_code in loading_df["車次代碼"].unique():
        trip_data = loading_df[loading_df["車次代碼"] == trip_code]
        summary_row = {"車次代碼": trip_code}
        
        for display_name, internal_name in numeric_mappings.items():
            if internal_name in loading_df.columns:
                summary_row[f"{display_name}小計"] = trip_data[internal_name].sum()
        
        trip_summary_data.append(summary_row)
    
    trip_summary = pd.DataFrame(trip_summary_data)
    for col in trip_summary.columns:
        if "小計" in col and trip_summary[col].dtype in ['float64', 'int64']:
            trip_summary[col] = trip_summary[col].round(3)
    
    # 分析標籤
    load_tab1, load_tab2, load_tab3 = st.tabs(["📦 裝載明細", "🚛 車次彙總", "📊 裝載分析"])
    
    with load_tab1:
        st.dataframe(detailed_loading_df, use_container_width=True, height=500)
        st.download_button(
            "📥 下載裝載明細",
            data=detailed_loading_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="配送裝載明細.csv",
            mime="text/csv",
        )
    
    with load_tab2:
        total_trips = len(trip_summary)
        col1, col2, col3 = st.columns(3)
        with col1:
            create_enhanced_metric_card("總車次數", f"{total_trips:,}")
        with col2:
            if not detailed_loading_df.empty:
                avg_items = detailed_loading_df.groupby("車次代碼").size().mean()
                create_enhanced_metric_card("平均每車項目數", f"{avg_items:.1f}")
        with col3:
            copper_col = "銅重量(噸)小計"
            if copper_col in trip_summary.columns:
                valid_weights = trip_summary[trip_summary[copper_col] > 0][copper_col]
                avg_weight = valid_weights.mean() if not valid_weights.empty else 0
                create_enhanced_metric_card("平均載重(噸)", f"{avg_weight:.2f}")
        
        st.dataframe(trip_summary, use_container_width=True)
        
        st.download_button(
            "📥 下載車次彙總",
            data=trip_summary.to_csv(index=False).encode("utf-8-sig"),
            file_name="車次彙總統計.csv",
            mime="text/csv",
        )
    
    with load_tab3:
        copper_col = "銅重量(噸)小計"
        if copper_col in trip_summary.columns:
            weight_data = trip_summary[trip_summary[copper_col] > 0][copper_col]
            
            if not weight_data.empty and len(weight_data) > 1:
                fig_dist = go.Figure()
                fig_dist.add_trace(go.Histogram(
                    x=weight_data,
                    nbinsx=min(20, len(weight_data)),
                    name="載重分佈",
                    marker_color=COLORFUL_PALETTE[0],
                    opacity=0.8
                ))
                fig_dist.update_layout(
                    title="車輛載重分佈直方圖",
                    xaxis_title="載重(噸)",
                    yaxis_title="車次數",
                    plot_bgcolor='white'
                )
                st.plotly_chart(fig_dist, use_container_width=True)
                
                # 統計摘要
                stats_data = {
                    "統計項目": ["最小載重", "最大載重", "平均載重", "中位數載重", "標準差"],
                    "數值(噸)": [
                        f"{weight_data.min():.3f}",
                        f"{weight_data.max():.3f}",
                        f"{weight_data.mean():.3f}",
                        f"{weight_data.median():.3f}",
                        f"{weight_data.std():.3f}"
                    ]
                }
                st.dataframe(pd.DataFrame(stats_data), use_container_width=True)

# ========================
# 主程式流程
# ========================

def main():
    """主程式入口"""
    
    if not EXCEL_AVAILABLE:
        st.warning("⚠️ 未安裝 openpyxl，Excel 功能將受限")
    
    # 檔案上傳
    st.markdown("### 📁 檔案上傳")
    uploaded_file = st.file_uploader(
        "選擇您的數據檔案",
        type=["xlsx", "xls", "csv"],
        help="支援格式：Excel (.xlsx, .xls) 和 CSV (.csv)",
        accept_multiple_files=False
    )
    
    if not uploaded_file:
        st.markdown("### 📖 使用說明")
        st.info("""
        **功能概述：**
        - 📊 **出貨類型分析**：統計各類型筆數與銅重量
        - 🚚 **達交率分析**：計算準時交付率，識別延遲問題  
        - 🗺️ **區域分析**：分析配送地區分佈與客戶偏好
        - 📦 **裝載分析**：車次裝載統計與效率分析
        """)
        return
    
    # 載入資料
    file_extension = uploaded_file.name.split(".")[-1].lower()
    
    with st.spinner("正在處理檔案..."):
        df = load_data(uploaded_file, file_extension)
    
    if df.empty:
        st.error("檔案載入失敗或檔案為空")
        return
    
    st.success(f"✅ 成功載入 {len(df):,} 筆記錄，{len(df.columns)} 個欄位")
    
    # 自動偵測欄位
    auto_mapping = smart_column_detector(df.columns.tolist())
    
    # 主要頁籤
    main_tab1, main_tab2 = st.tabs(["📊 數據分析", "📄 原始數據"])
    
    with main_tab2:
        st.markdown("### 原始數據預覽")
        st.dataframe(df, use_container_width=True, height=600)
        
        st.markdown("### 📈 基本統計")
        basic_stats = pd.DataFrame({
            "項目": ["總記錄數", "欄位數", "空值總數", "重複記錄數"],
            "數值": [
                f"{len(df):,}",
                f"{len(df.columns)}",
                f"{df.isnull().sum().sum():,}",
                f"{df.duplicated().sum():,}"
            ]
        })
        st.dataframe(basic_stats, use_container_width=True)
    
    with main_tab1:
        # 側邊欄篩選
        with st.sidebar:
            st.markdown("### ⚙️ 分析控制面板")
            
            with st.expander("🔍 通用篩選", expanded=True):
                # 出貨類型篩選
                selected_ship_types = []
                if auto_mapping.get("ship_type"):
                    unique_types = sorted(df[auto_mapping["ship_type"]].dropna().unique())
                    selected_ship_types = st.multiselect(
                        "出貨申請類型", 
                        options=unique_types,
                        default=[]
                    )
                
                # 客戶名稱篩選
                selected_customers = []
                if auto_mapping.get("cust_name"):
                    unique_customers = sorted(df[auto_mapping["cust_name"]].dropna().unique())
                    selected_customers = st.multiselect(
                        "客戶名稱",
                        options=unique_customers[:100],
                        default=[]
                    )
                
                # 客戶編號篩選
                selected_cust_ids = []
                if auto_mapping.get("cust_id"):
                    unique_cust_ids = sorted(df[auto_mapping["cust_id"]].dropna().unique())
                    selected_cust_ids = st.multiselect(
                        "客戶編號",
                        options=unique_cust_ids[:100],
                        default=[]
                    )
                
                # 料號說明篩選
                selected_item_descs = []
                if auto_mapping.get("item_desc"):
                    unique_items = sorted(df[auto_mapping["item_desc"]].dropna().unique())
                    selected_item_descs = st.multiselect(
                        "料號說明",
                        options=unique_items[:100],
                        default=[]
                    )
                
                # 日期篩選
                date_range = None
                if auto_mapping.get("due_date"):
                    due_dates = safe_datetime_convert(df[auto_mapping["due_date"]])
                    if due_dates.notna().any():
                        min_date = due_dates.min().date()
                        max_date = due_dates.max().date()
                        date_range = st.date_input(
                            "指定到貨日期範圍",
                            value=(min_date, max_date),
                            min_value=min_date,
                            max_value=max_date
                        )
            
            with st.expander("⚙️ 進階設定"):
                topn_cities = st.slider("每客戶顯示前N個配送縣市", 1, 10, 3)
        
        # 套用篩選
        filtered_df = df.copy()
        
        if selected_ship_types and auto_mapping.get("ship_type"):
            filtered_df = filtered_df[filtered_df[auto_mapping["ship_type"]].isin(selected_ship_types)]
        
        if selected_customers and auto_mapping.get("cust_name"):
            filtered_df = filtered_df[filtered_df[auto_mapping["cust_name"]].isin(selected_customers)]
        
        if date_range and len(date_range) == 2 and auto_mapping.get("due_date"):
            start_date = pd.to_datetime(date_range[0])
            end_date = pd.to_datetime(date_range[1])
            due_dates = safe_datetime_convert(filtered_df[auto_mapping["due_date"]])
            date_mask = (due_dates >= start_date) & (due_dates <= end_date)
            filtered_df = filtered_df[date_mask]
        
        if len(filtered_df) != len(df):
            st.info(f"🔍 篩選後顯示 {len(filtered_df):,} 筆記錄（原始：{len(df):,} 筆）")
        
        # SWI排除邏輯
        exclude_swi = pd.Series(True, index=filtered_df.index)
        if auto_mapping.get("ship_type"):
            exclude_swi = filtered_df[auto_mapping["ship_type"]] != "SWI-寄庫"
        
        # 執行分析
        if not filtered_df.empty:
            enhanced_shipment_analysis(filtered_df, auto_mapping)
            st.markdown("---")
            enhanced_delivery_performance(filtered_df, auto_mapping, exclude_swi)
            st.markdown("---") 
            enhanced_area_analysis(filtered_df, auto_mapping, exclude_swi, topn_cities)
            st.markdown("---")
            enhanced_loading_analysis(filtered_df, auto_mapping, exclude_swi)
        else:
            st.warning("⚠️ 篩選條件過於嚴格，無數據可顯示")

if __name__ == "__main__":
    main()
