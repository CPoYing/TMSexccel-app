import io
import re
from typing import Dict, Optional, Tuple
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ========================
# 頁面設定與 CSS 樣式
# ========================

st.set_page_config(
    page_title="TMS出貨配送數據分析",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定義 CSS 樣式
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(90deg, #4CAF50, #2196F3);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 1rem;
    }
    
    .metric-container {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #4CAF50;
        margin: 0.5rem 0;
    }
    
    .section-header {
        color: #2c3e50;
        border-bottom: 2px solid #3498db;
        padding-bottom: 0.5rem;
        margin: 1.5rem 0 1rem 0;
    }
    
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .success-box {
        background-color: #d1edff;
        border: 1px solid #74b9ff;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
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
    """
    高效能檔案載入與快取
    
    Args:
        file_buffer: 檔案緩衝區
        file_type: 檔案類型 (csv, xlsx, xls)
    
    Returns:
        DataFrame: 載入的數據
    """
    try:
        if file_type == "csv":
            # 自動偵測編碼
            encodings = ['utf-8', 'utf-8-sig', 'big5', 'gbk']
            for encoding in encodings:
                try:
                    file_buffer.seek(0)
                    df = pd.read_csv(file_buffer, encoding=encoding)
                    return df
                except UnicodeDecodeError:
                    continue
            # 如果都失敗，使用預設編碼
            file_buffer.seek(0)
            return pd.read_csv(file_buffer, encoding='utf-8', errors='ignore')
        elif file_type in ["xlsx", "xls"]:
            return pd.read_excel(file_buffer, engine="openpyxl")
    except Exception as e:
        st.error(f"檔案載入失敗: {str(e)}")
        return pd.DataFrame()
    
    st.error("不支援的檔案類型。請上傳 .xlsx, .xls 或 .csv 檔案。")
    return pd.DataFrame()

def smart_column_detector(df_cols: list) -> Dict[str, Optional[str]]:
    """
    智能欄位偵測器 - 提升偵測準確度
    
    Args:
        df_cols: DataFrame 欄位列表
    
    Returns:
        Dict: 偵測到的欄位對應字典
    """
    def find_best_match(keywords: list, priority_keywords: list = None) -> Optional[str]:
        """尋找最佳匹配欄位，支援優先級搜尋"""
        # 優先搜尋高優先級關鍵字
        if priority_keywords:
            for kw in priority_keywords:
                for col in df_cols:
                    if kw.lower() in str(col).lower():
                        return col
        
        # 一般搜尋
        for kw in keywords:
            for col in df_cols:
                if kw.lower() in str(col).lower():
                    return col
        return None

    column_mapping = {
        "ship_type": find_best_match(
            ["出貨申請類型", "出貨類型", "配送類型", "類型", "申請類型"],
            ["出貨申請類型", "出貨類型"]
        ),
        "due_date": find_best_match(
            ["指定到貨日期", "到貨日期", "指定到貨", "到貨日", "預計到貨"],
            ["指定到貨日期"]
        ),
        "sign_date": find_best_match(
            ["客戶簽收日期", "簽收日期", "簽收日", "簽收時間"],
            ["客戶簽收日期"]
        ),
        "cust_id": find_best_match(
            ["客戶編號", "客戶代號", "客編", "客戶id"],
            ["客戶編號"]
        ),
        "cust_name": find_best_match(
            ["客戶名稱", "客名", "客戶", "客戶簡稱"],
            ["客戶名稱"]
        ),
        "address": find_best_match(
            ["地址", "收貨地址", "送貨地址", "交貨地址", "配送地址"],
            ["地址"]
        ),
        "copper_ton": find_best_match(
            ["銅重量(噸)", "銅重量(噸數)", "銅噸", "銅重量", "銅重量（噸）"],
            ["銅重量(噸)", "銅重量（噸）"]
        ),
        "qty": find_best_match(
            ["出貨數量", "數量", "出貨量", "出庫數量"],
            ["出貨數量"]
        ),
        "ship_no": find_best_match(
            ["出庫單號", "出庫單", "出庫編號", "出貨單號"],
            ["出庫單號"]
        ),
        "do_no": find_best_match(
            ["DO號", "DO", "出貨單號", "交貨單號", "do號"],
            ["DO號"]
        ),
        "item_desc": find_best_match(
            ["料號說明", "品名", "品名規格", "物料說明", "產品名稱"],
            ["料號說明", "品名"]
        ),
        "lot_no": find_best_match(
            ["批次", "批號", "Lot", "lot", "批次號"],
            ["批次", "批號"]
        ),
        "fg_net_ton": find_best_match(
            ["成品淨重(噸)", "成品淨量(噸)", "淨重(噸)", "淨重"],
            ["成品淨重(噸)"]
        ),
        "fg_gross_ton": find_best_match(
            ["成品毛重(噸)", "成品毛量(噸)", "毛重(噸)", "毛重"],
            ["成品毛重(噸)"]
        ),
        "status": find_best_match(
            ["通知單項次狀態", "項次狀態", "狀態", "單據狀態"],
            ["通知單項次狀態"]
        ),
    }
    
    return column_mapping

def safe_datetime_convert(series: pd.Series) -> pd.Series:
    """安全的日期時間轉換，支援多種格式"""
    if series.empty:
        return series
    
    # 嘗試多種日期格式
    formats = [
        "%Y-%m-%d",
        "%Y/%m/%d", 
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y",
        "%d/%m/%Y"
    ]
    
    result = pd.to_datetime(series, errors="coerce")
    
    # 如果轉換失敗率太高，嘗試其他格式
    if result.isna().sum() / len(result) > 0.5:
        for fmt in formats:
            try:
                temp_result = pd.to_datetime(series, format=fmt, errors="coerce")
                if temp_result.isna().sum() < result.isna().sum():
                    result = temp_result
                    break
            except:
                continue
    
    return result

def enhanced_city_extraction(address: str) -> Optional[str]:
    """
    增強版台灣縣市擷取，支援更多格式
    
    Args:
        address: 地址字串
        
    Returns:
        Optional[str]: 擷取到的縣市名稱
    """
    if pd.isna(address) or not address:
        return None
    
    # 正規化地址
    normalized_addr = str(address).strip().replace("臺", "台").replace("　", " ")
    
    # 台灣縣市列表（包含常見變體）
    taiwan_cities = [
        "台北市", "新北市", "桃園市", "台中市", "台南市", "高雄市",
        "基隆市", "新竹市", "新竹縣", "苗栗縣", "彰化縣", "南投縣",
        "雲林縣", "嘉義市", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣",
        "台東縣", "澎湖縣", "金門縣", "連江縣"
    ]
    
    # 優先完整匹配
    for city in taiwan_cities:
        if city in normalized_addr:
            return city
    
    # 次要：模糊匹配（去掉市/縣）
    for city in taiwan_cities:
        city_base = city.replace("市", "").replace("縣", "")
        if city_base in normalized_addr and len(city_base) >= 2:
            return city
            
    return None

def create_enhanced_metric_card(title: str, value: str, delta: str = None, delta_color: str = "normal"):
    """建立增強版指標卡片"""
    delta_html = ""
    if delta:
        color = {"normal": "#666", "positive": "#28a745", "negative": "#dc3545"}.get(delta_color, "#666")
        delta_html = f'<div style="color: {color}; font-size: 0.9rem; margin-top: 0.5rem;">{delta}</div>'
    
    st.markdown(f"""
    <div class="metric-container">
        <div style="color: #666; font-size: 0.9rem; margin-bottom: 0.3rem;">{title}</div>
        <div style="font-size: 2rem; font-weight: bold; color: #2c3e50;">{value}</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)

# ========================
# 核心分析函式優化
# ========================

def enhanced_shipment_analysis(df: pd.DataFrame, col_map: Dict[str, Optional[str]]) -> None:
    """
    增強版出貨類型分析
    
    Args:
        df: 資料框
        col_map: 欄位對應字典
    """
    st.markdown('<h2 class="section-header">① 出貨類型筆數與重量統計</h2>', unsafe_allow_html=True)
    
    ship_type_col = col_map.get("ship_type")
    copper_ton_col = col_map.get("copper_ton")
    
    if not ship_type_col or ship_type_col not in df.columns:
        st.error("⚠️ 找不到『出貨申請類型』欄位，請確認檔案格式是否正確。")
        return
    
    # 處理出貨類型統計
    df_work = df.copy()
    df_work[ship_type_col] = df_work[ship_type_col].fillna("(空白)")
    
    type_stats = df_work[ship_type_col].value_counts().reset_index()
    type_stats.columns = ["出貨類型", "筆數"]
    
    # 計算銅重量統計
    if copper_ton_col and copper_ton_col in df.columns:
        df_work["_copper_numeric"] = pd.to_numeric(df_work[copper_ton_col], errors="coerce")
        copper_stats = df_work.groupby(ship_type_col)["_copper_numeric"].agg([
            ('銅重量_kg_合計', 'sum'),
            ('平均銅重量_kg', 'mean')
        ]).reset_index()
        copper_stats.columns = ["出貨類型", "銅重量(kg)合計", "平均銅重量(kg)"]
        copper_stats["銅重量(噸)合計"] = (copper_stats["銅重量(kg)合計"] / 1000).round(3)
        copper_stats["平均銅重量(噸)"] = (copper_stats["平均銅重量(kg)"] / 1000).round(3)
        
        # 合併統計數據
        final_stats = type_stats.merge(
            copper_stats[["出貨類型", "銅重量(噸)合計", "平均銅重量(噸)"]], 
            on="出貨類型", 
            how="left"
        )
    else:
        final_stats = type_stats.copy()
        final_stats["銅重量(噸)合計"] = None
        final_stats["平均銅重量(噸)"] = None
    
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
        total_copper = final_stats["銅重量(噸)合計"].sum() if "銅重量(噸)合計" in final_stats.columns else 0
        create_enhanced_metric_card("總銅重量(噸)", f"{total_copper:,.3f}")
    with col4:
        avg_per_record = total_copper / total_records if total_records > 0 else 0
        create_enhanced_metric_card("平均每筆重量(噸)", f"{avg_per_record:.3f}")
    
    # 顯示詳細統計表格與圖表
    tab1, tab2 = st.tabs(["📋 詳細統計", "📊 視覺化圖表"])
    
    with tab1:
        st.dataframe(
            final_stats.style.format({
                "筆數": "{:,}",
                "佔比(%)": "{:.2f}%",
                "銅重量(噸)合計": "{:.3f}",
                "平均銅重量(噸)": "{:.3f}"
            }),
            use_container_width=True,
            height=400
        )
        
        # 下載按鈕
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "📥 下載統計表 (CSV)",
                data=final_stats.to_csv(index=False).encode("utf-8-sig"),
                file_name="出貨類型統計.csv",
                mime="text/csv",
            )
        with col2:
            st.download_button(
                "📥 下載統計表 (Excel)",
                data=final_stats.to_excel(index=False).encode() if hasattr(final_stats.to_excel(index=False), 'encode') else final_stats.to_excel(index=False, engine='openpyxl'),
                file_name="出貨類型統計.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    
    with tab2:
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            chart_type = st.selectbox(
                "圖表類型", 
                ["長條圖", "圓餅圖", "環形圖", "水平長條圖"], 
                key="chart_type_1"
            )
            
            # 創建互動式圖表
            if chart_type == "長條圖":
                fig = px.bar(
                    final_stats, 
                    x="出貨類型", 
                    y="筆數",
                    title="出貨類型筆數分佈",
                    color="筆數",
                    color_continuous_scale="viridis"
                )
                fig.update_layout(xaxis_tickangle=-45)
            elif chart_type == "水平長條圖":
                fig = px.bar(
                    final_stats, 
                    x="筆數", 
                    y="出貨類型",
                    orientation='h',
                    title="出貨類型筆數分佈",
                    color="筆數",
                    color_continuous_scale="viridis"
                )
            elif chart_type == "圓餅圖":
                fig = px.pie(
                    final_stats, 
                    names="出貨類型", 
                    values="筆數",
                    title="出貨類型佔比分佈"
                )
            else:  # 環形圖
                fig = px.pie(
                    final_stats, 
                    names="出貨類型", 
                    values="筆數",
                    title="出貨類型佔比分佈",
                    hole=0.4
                )
            
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)
        
        with chart_col2:
            if "銅重量(噸)合計" in final_stats.columns and final_stats["銅重量(噸)合計"].notna().any():
                fig_copper = px.treemap(
                    final_stats[final_stats["銅重量(噸)合計"].notna()],
                    path=["出貨類型"],
                    values="銅重量(噸)合計",
                    title="銅重量分佈樹狀圖",
                    color="銅重量(噸)合計",
                    color_continuous_scale="RdYlBu"
                )
                st.plotly_chart(fig_copper, use_container_width=True)
            else:
                st.info("無銅重量數據可供分析")

def enhanced_delivery_performance(df: pd.DataFrame, col_map: Dict[str, Optional[str]], exclude_swi_mask: pd.Series) -> None:
    """
    增強版配送達交率分析
    
    Args:
        df: 資料框
        col_map: 欄位對應字典
        exclude_swi_mask: 排除SWI遮罩
    """
    st.markdown('<h2 class="section-header">② 配送達交率分析</h2>', unsafe_allow_html=True)
    
    due_date_col = col_map.get("due_date")
    sign_date_col = col_map.get("sign_date")
    cust_id_col = col_map.get("cust_id")
    cust_name_col = col_map.get("cust_name")
    
    if not all([due_date_col, sign_date_col]) or not all([col in df.columns for col in [due_date_col, sign_date_col]]):
        st.error("⚠️ 找不到必要的日期欄位（指定到貨日期、客戶簽收日期）")
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
        create_enhanced_metric_card("達交率", f"{delivery_rate:.2f}%", 
                                   "優異" if delivery_rate >= 95 else "良好" if delivery_rate >= 90 else "需改善",
                                   "positive" if delivery_rate >= 95 else "normal" if delivery_rate >= 90 else "negative")
    with col5:
        create_enhanced_metric_card("平均延遲", f"{avg_delay:.1f}天", f"最大: {max_delay}天" if max_delay > 0 else "")
    
    # 詳細分析
    analysis_tab1, analysis_tab2, analysis_tab3 = st.tabs(["🚚 未達標明細", "📈 達交趨勢", "👥 客戶表現"])
    
    with analysis_tab1:
        if late_count > 0:
            late_details = pd.DataFrame({
                col: df[col] if col else None for col in [cust_id_col, cust_name_col]
            }).assign(
                指定到貨日期=due_day.dt.strftime("%Y-%m-%d"),
                客戶簽收日期=sign_day.dt.strftime("%Y-%m-%d"),
                延遲天數=delay_days
            )[late_mask].dropna(axis=1, how="all")
            
            # 延遲程度分類
            late_details["延遲程度"] = pd.cut(
                late_details["延遲天數"],
                bins=[-float('inf'), 0, 3, 7, 30, float('inf')],
                labels=["準時", "輕微延遲(1-3天)", "中度延遲(4-7天)", "重度延遲(8-30天)", "嚴重延遲(>30天)"]
            )
            
            # 排序：延遲天數降序
            late_details = late_details.sort_values("延遲天數", ascending=False)
            
            st.dataframe(
                late_details.style.format({"延遲天數": "{:.0f}天"}),
                use_container_width=True,
                height=400
            )
            
            # 延遲分佈分析
            delay_dist = late_details["延遲程度"].value_counts()
            fig_delay = px.bar(
                x=delay_dist.index, 
                y=delay_dist.values,
                title="延遲程度分佈",
                labels={"x": "延遲程度", "y": "筆數"}
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
        if total_valid >= 7:  # 至少要有一週的數據
            # 依日期統計達交率趨勢
            df_trend = df[valid_mask].copy()
            df_trend["到貨日期"] = due_day[valid_mask]
            df_trend["準時"] = on_time_mask[valid_mask]
            
            daily_trend = df_trend.groupby(df_trend["到貨日期"].dt.date).agg(
                總筆數=("準時", "count"),
                準時筆數=("準時", "sum")
            )
            daily_trend["達交率(%)"] = (daily_trend["準時筆數"] / daily_trend["總筆數"] * 100).round(2)
            daily_trend = daily_trend.reset_index()
            
            fig_trend = px.line(
                daily_trend, 
                x="到貨日期", 
                y="達交率(%)",
                title="每日達交率趨勢",
                markers=True
            )
            fig_trend.add_hline(y=95, line_dash="dash", line_color="green", annotation_text="目標線(95%)")
            fig_trend.add_hline(y=90, line_dash="dash", line_color="orange", annotation_text="警戒線(90%)")
            
            st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.info("數據量不足，無法分析趨勢（需至少7筆記錄）")
    
    with analysis_tab3:
        if cust_name_col and cust_name_col in df.columns:
            customer_performance = df[valid_mask].groupby(cust_name_col).agg(
                有效筆數=("客戶簽收日期", "count") if "客戶簽收日期" in df.columns else (cust_name_col, "count"),
                準時筆數=(on_time_mask[valid_mask], "sum"),
                延遲筆數=(late_mask[valid_mask], "sum")
            ).reset_index()
            
            customer_performance["達交率(%)"] = (
                customer_performance["準時筆數"] / customer_performance["有效筆數"] * 100
            ).round(2)
            
            customer_performance = customer_performance.sort_values("達交率(%)", ascending=False)
            
            # 客戶表現分級
            customer_performance["表現等級"] = pd.cut(
                customer_performance["達交率(%)"],
                bins=[0, 80, 90, 95, 100],
                labels=["需改善", "一般", "良好", "優秀"],
                include_lowest=True
            )
            
            st.dataframe(
                customer_performance.style.format({
                    "有效筆數": "{:,}",
                    "準時筆數": "{:,}",
                    "延遲筆數": "{:,}",
                    "達交率(%)": "{:.2f}%"
                }).background_gradient(subset=["達交率(%)"], cmap="RdYlGn"),
                use_container_width=True
            )
            
            st.download_button(
                "📥 下載客戶表現分析",
                data=customer_performance.to_csv(index=False).encode("utf-8-sig"),
                file_name="客戶達交表現.csv",
                mime="text/csv",
            )

def enhanced_area_analysis(df: pd.DataFrame, col_map: Dict[str, Optional[str]], exclude_swi_mask: pd.Series, topn_cities: int) -> None:
    """
    增強版配送區域分析
    
    Args:
        df: 資料框
        col_map: 欄位對應字典
        exclude_swi_mask: 排除SWI遮罩
        topn_cities: 每客戶顯示前N個縣市
    """
    st.markdown('<h2 class="section-header">③ 配送區域分析</h2>', unsafe_allow_html=True)
    
    address_col = col_map.get("address")
    cust_name_col = col_map.get("cust_name")
    
    if not address_col or address_col not in df.columns:
        st.error("⚠️ 找不到『地址』欄位")
        return
    
    # 處理區域資料
    region_df = df[exclude_swi_mask].copy()
    region_df["縣市"] = region_df[address_col].apply(enhanced_city_extraction)
    
    # 統計無法識別的地址
    unknown_addresses = region_df[region_df["縣市"].isna()][address_col].value_counts().head(10)
    
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
    
    # 詳細分析標籤
    tab1, tab2, tab3 = st.tabs(["🗺️ 縣市分佈", "🏢 客戶區域", "⚠️ 地址問題"])
    
    with tab1:
        col1, col2 = st.columns([1, 1])
        with col1:
            st.dataframe(
                city_stats.style.format({
                    "筆數": "{:,}",
                    "佔比(%)": "{:.2f}%"
                }),
                use_container_width=True,
                height=400
            )
        with col2:
            # 台灣地圖視覺化 (使用長條圖代替)
            fig_map = px.bar(
                city_stats.head(15), 
                x="縣市", 
                y="筆數",
                title="前15縣市配送量",
                color="筆數",
                color_continuous_scale="viridis"
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
                .sort_values(["筆數"], ascending=False)
                .groupby(cust_name_col)
                .head(topn_cities)
                .sort_values([cust_name_col, "筆數"], ascending=[True, False])
            )
            
            st.dataframe(
                top_customer_cities.rename(columns={cust_name_col: "客戶名稱"}).style.format({
                    "筆數": "{:,}",
                    "客戶總筆數": "{:,}",
                    "縣市佔比(%)": "{:.2f}%"
                }),
                use_container_width=True
            )
            
            st.download_button(
                "📥 下載客戶區域分析",
                data=top_customer_cities.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"客戶區域分析_Top{topn_cities}.csv",
                mime="text/csv",
            )
        else:
            st.info("無客戶名稱欄位，無法進行客戶區域分析")
    
    with tab3:
        if not unknown_addresses.empty:
            st.warning(f"發現 {len(unknown_addresses)} 種無法識別的地址格式")
            unknown_df = unknown_addresses.reset_index()
            unknown_df.columns = ["地址", "出現次數"]
            st.dataframe(unknown_df, use_container_width=True)
        else:
            st.success("✅ 所有地址都已成功識別縣市")

def enhanced_loading_analysis(df: pd.DataFrame, col_map: Dict[str, Optional[str]], exclude_swi_mask: pd.Series) -> None:
    """
    增強版配送裝載分析
    
    Args:
        df: 資料框
        col_map: 欄位對應字典
        exclude_swi_mask: 排除SWI遮罩
    """
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
        st.info(f"✅ 已套用狀態篩選：僅顯示『完成』狀態的記錄")
    else:
        st.warning("⚠️ 未找到『狀態』欄位，將顯示所有記錄")
    
    loading_df = df[base_mask].copy()
    
    if loading_df.empty:
        st.warning("無符合條件的裝載數據")
        return
    
    # 車次代碼生成
    loading_df["車次代碼"] = loading_df[ship_no_col].astype(str).str[:13]
    
    # 數值欄位處理
    numeric_cols = {
        "qty": col_map.get("qty"),
        "copper_ton": col_map.get("copper_ton"),
        "fg_net_ton": col_map.get("fg_net_ton"),
        "fg_gross_ton": col_map.get("fg_gross_ton")
    }
    
    for key, col_name in numeric_cols.items():
        if col_name and col_name in loading_df.columns:
            loading_df[f"_{key}_numeric"] = pd.to_numeric(loading_df[col_name], errors="coerce")
    
    # 轉換銅重量為噸
    if "_copper_ton_numeric" in loading_df.columns:
        loading_df["_copper_ton_converted"] = loading_df["_copper_ton_numeric"] / 1000.0
    
    # 建立展示用資料框
    display_columns = {
        "車次代碼": loading_df["車次代碼"],
        "出庫單號": loading_df[ship_no_col],
    }
    
    # 動態新增可用欄位
    optional_cols = {
        "DO號": col_map.get("do_no"),
        "料號說明": col_map.get("item_desc"),
        "批次": col_map.get("lot_no"),
        "客戶名稱": col_map.get("cust_name")
    }
    
    for display_name, col_name in optional_cols.items():
        if col_name and col_name in loading_df.columns:
            display_columns[display_name] = loading_df[col_name]
    
    # 數值欄位
    numeric_display_cols = {
        "出貨數量": "_qty_numeric",
        "銅重量(噸)": "_copper_ton_converted",
        "成品淨重(噸)": "_fg_net_ton_numeric",
        "成品毛重(噸)": "_fg_gross_ton_numeric"
    }
    
    for display_name, internal_name in numeric_display_cols.items():
        if internal_name in loading_df.columns:
            display_columns[display_name] = loading_df[internal_name].round(3)
    
    detailed_loading_df = pd.DataFrame(display_columns)
    
    # 排序
    sort_columns = [col for col in ["車次代碼", "出庫單號"] if col in detailed_loading_df.columns]
    if sort_columns:
        detailed_loading_df = detailed_loading_df.sort_values(sort_columns).reset_index(drop=True)
    
    # 車次彙總統計
    summary_agg = {}
    for display_name, internal_name in numeric_display_cols.items():
        if internal_name in loading_df.columns:
            summary_agg[display_name + "小計"] = (internal_name, "sum")
            summary_agg[display_name + "項目數"] = (internal_name, "count")
    
    if summary_agg:
        trip_summary = loading_df.groupby("車次代碼").agg(summary_agg).round(3)
        trip_summary.columns = [col[0] for col in trip_summary.columns]
        trip_summary = trip_summary.reset_index()
    else:
        trip_summary = loading_df.groupby("車次代碼").size().reset_index(name="項目數")
    
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
        # 車次統計指標
        total_trips = len(trip_summary)
        col1, col2, col3 = st.columns(3)
        with col1:
            create_enhanced_metric_card("總車次數", f"{total_trips:,}")
        with col2:
            avg_items_per_trip = detailed_loading_df.groupby("車次代碼").size().mean()
            create_enhanced_metric_card("平均每車項目數", f"{avg_items_per_trip:.1f}")
        with col3:
            if "銅重量(噸)小計" in trip_summary.columns:
                valid_trips = trip_summary[trip_summary["銅重量(噸)小計"] > 0]
                avg_weight = valid_trips["銅重量(噸)小計"].mean() if not valid_trips.empty else 0
                create_enhanced_metric_card("平均載重(噸)", f"{avg_weight:.2f}")
        
        st.dataframe(
            trip_summary.style.format({col: "{:.3f}" for col in trip_summary.columns if "小計" in col}),
            use_container_width=True
        )
        
        st.download_button(
            "📥 下載車次彙總",
            data=trip_summary.to_csv(index=False).encode("utf-8-sig"),
            file_name="車次彙總統計.csv",
            mime="text/csv",
        )
    
    with load_tab3:
        if "銅重量(噸)小計" in trip_summary.columns:
            # 載重分佈分析
            weight_data = trip_summary[trip_summary["銅重量(噸)小計"] > 0]["銅重量(噸)小計"]
            
            if not weight_data.empty:
                fig_dist = go.Figure()
                fig_dist.add_trace(go.Histogram(
                    x=weight_data,
                    nbinsx=20,
                    name="載重分佈",
                    marker_color="lightblue"
                ))
                fig_dist.update_layout(
                    title="車輛載重分佈直方圖",
                    xaxis_title="載重(噸)",
                    yaxis_title="車次數"
                )
                st.plotly_chart(fig_dist, use_container_width=True)
                
                # 載重統計摘要
                stats_summary = pd.DataFrame({
                    "統計項目": ["最小載重", "最大載重", "平均載重", "中位數載重", "標準差"],
                    "數值(噸)": [
                        weight_data.min(),
                        weight_data.max(),
                        weight_data.mean(),
                        weight_data.median(),
                        weight_data.std()
                    ]
                }).round(3)
                
                st.dataframe(stats_summary, use_container_width=True)
        else:
            st.info("無銅重量數據可供載重分析")

# ========================
# 智能欄位對應介面
# ========================

def display_column_mapping_interface(df: pd.DataFrame, auto_mapping: Dict[str, Optional[str]]) -> Dict[str, Optional[str]]:
    """
    顯示智能欄位對應介面
    
    Args:
        df: 資料框
        auto_mapping: 自動偵測的欄位對應
        
    Returns:
        Dict: 用戶確認後的欄位對應
    """
    st.markdown('<h2 class="section-header">🔧 欄位對應設定</h2>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("**自動偵測結果**")
        mapping_status = []
        for logical_name, detected_col in auto_mapping.items():
            status = "✅ 已偵測" if detected_col else "❌ 未偵測"
            mapping_status.append({
                "邏輯欄位": logical_name,
                "偵測到的欄位": detected_col or "(無)",
                "狀態": status
            })
        
        st.dataframe(pd.DataFrame(mapping_status), use_container_width=True)
    
    with col2:
        st.markdown("**手動調整**")
        
        # 欄位選項
        column_options = ["(不使用)"] + list(df.columns)
        
        # 關鍵欄位映射
        key_fields = {
            "ship_type": "出貨申請類型",
            "due_date": "指定到貨日期", 
            "sign_date": "客戶簽收日期",
            "address": "地址",
            "cust_name": "客戶名稱"
        }
        
        user_mapping = {}
        for field_key, field_display in key_fields.items():
            default_idx = 0
            if auto_mapping.get(field_key) in column_options:
                default_idx = column_options.index(auto_mapping.get(field_key))
            
            selected = st.selectbox(
                f"{field_display}:",
                column_options,
                index=default_idx,
                key=f"map_{field_key}"
            )
            user_mapping[field_key] = selected if selected != "(不使用)" else None
        
        # 其他欄位摺疊顯示
        with st.expander("其他欄位設定"):
            other_fields = {
                "cust_id": "客戶編號",
                "copper_ton": "銅重量",
                "qty": "出貨數量",
                "ship_no": "出庫單號",
                "do_no": "DO號",
                "item_desc": "料號說明",
                "lot_no": "批次",
                "fg_net_ton": "成品淨重",
                "fg_gross_ton": "成品毛重",
                "status": "狀態"
            }
            
            for field_key, field_display in other_fields.items():
                default_idx = 0
                if auto_mapping.get(field_key) in column_options:
                    default_idx = column_options.index(auto_mapping.get(field_key))
                
                selected = st.selectbox(
                    f"{field_display}:",
                    column_options,
                    index=default_idx,
                    key=f"map_{field_key}_other"
                )
                user_mapping[field_key] = selected if selected != "(不使用)" else None
    
    return user_mapping

# ========================
# 主程式流程
# ========================

def main():
    """主程式入口"""
    
    # 檔案上傳區
    st.markdown("### 📁 檔案上傳")
    uploaded_file = st.file_uploader(
        "選擇您的數據檔案",
        type=["xlsx", "xls", "csv"],
        help="支援格式：Excel (.xlsx, .xls) 和 CSV (.csv)，檔案大小限制 200MB",
        accept_multiple_files=False
    )
    
    if not uploaded_file:
        # 顯示使用說明
        st.markdown("### 📖 使用說明")
        st.info("""
        **功能概述：**
        - 📊 **出貨類型分析**：統計各類型筆數與銅重量
        - 🚚 **達交率分析**：計算準時交付率，識別延遲問題  
        - 🗺️ **區域分析**：分析配送地區分佈與客戶偏好
        - 📦 **裝載分析**：車次裝載統計與效率分析
        
        **開始使用：**
        1. 上傳包含出貨數據的 Excel 或 CSV 檔案
        2. 系統將自動偵測欄位並開始分析
        3. 使用側邊欄進行篩選和客製化分析
        """)
        return
    
    # 載入資料
    file_extension = uploaded_file.name.split(".")[-1].lower()
    
    with st.spinner("正在處理檔案..."):
        df = load_data(uploaded_file, file_extension)
    
    if df.empty:
        st.error("檔案載入失敗或檔案為空，請檢查檔案格式")
        return
    
    # 成功載入提示
    st.success(f"✅ 成功載入 {len(df):,} 筆記錄，{len(df.columns)} 個欄位")
    
    # 自動偵測欄位對應
    auto_detected_mapping = smart_column_detector(df.columns.tolist())
    
    # 建立主要頁籤
    main_tab1, main_tab2, main_tab3 = st.tabs(["📊 數據分析", "⚙️ 欄位設定", "📄 原始數據"])
    
    with main_tab3:
        st.markdown("### 原始數據預覽")
        st.dataframe(df, use_container_width=True, height=600)
        
        # 基本統計資訊
        st.markdown("### 📈 基本統計")
        basic_stats = pd.DataFrame({
            "項目": ["總記錄數", "欄位數", "空值總數", "重複記錄數"],
            "數值": [
                len(df),
                len(df.columns),
                df.isnull().sum().sum(),
                df.duplicated().sum()
            ]
        })
        st.dataframe(basic_stats, use_container_width=True)
    
    with main_tab2:
        final_column_mapping = display_column_mapping_interface(df, auto_detected_mapping)
    
    with main_tab1:
        # 使用最終的欄位對應
        final_column_mapping = auto_detected_mapping  # 簡化版：直接使用自動偵測
        
        # 側邊欄篩選控制
        with st.sidebar:
            st.markdown("### ⚙️ 分析控制面板")
            
            # 通用篩選
            with st.expander("🔍 通用篩選", expanded=True):
                selected_ship_types = []
                if final_column_mapping.get("ship_type"):
                    unique_types = sorted(df[final_column_mapping["ship_type"]].dropna().unique())
                    selected_ship_types = st.multiselect(
                        "出貨申請類型", 
                        options=unique_types,
                        default=[],
                        key="filter_ship_types"
                    )
                
                selected_customers = []
                if final_column_mapping.get("cust_name"):
                    unique_customers = sorted(df[final_column_mapping["cust_name"]].dropna().unique())
                    selected_customers = st.multiselect(
                        "客戶名稱",
                        options=unique_customers[:100],  # 限制顯示數量避免介面過載
                        default=[],
                        key="filter_customers"
                    )
                
                # 日期範圍篩選
                date_filter_range = None
                if final_column_mapping.get("due_date"):
                    due_date_series = safe_datetime_convert(df[final_column_mapping["due_date"]])
                    if due_date_series.notna().any():
                        min_date = due_date_series.min().date()
                        max_date = due_date_series.max().date()
                        date_filter_range = st.date_input(
                            "指定到貨日期範圍",
                            value=(min_date, max_date),
                            min_value=min_date,
                            max_value=max_date,
                            key="date_range_filter"
                        )
            
            # 進階設定
            with st.expander("⚙️ 進階設定", expanded=False):
                topn_cities_setting = st.slider(
                    "每客戶顯示前N個配送縣市", 
                    min_value=1, 
                    max_value=10, 
                    value=3,
                    key="topn_cities_slider"
                )
                
                enable_detailed_analysis = st.checkbox(
                    "啟用詳細分析模式", 
                    value=True,
                    help="包含更多統計圖表和深度分析"
                )
        
        # 套用篩選條件
        filtered_data = df.copy()
        
        if selected_ship_types and final_column_mapping.get("ship_type"):
            filtered_data = filtered_data[filtered_data[final_column_mapping["ship_type"]].isin(selected_ship_types)]
        
        if selected_customers and final_column_mapping.get("cust_name"):
            filtered_data = filtered_data[filtered_data[final_column_mapping["cust_name"]].isin(selected_customers)]
        
        if date_filter_range and len(date_filter_range) == 2 and final_column_mapping.get("due_date"):
            start_date = pd.to_datetime(date_filter_range[0])
            end_date = pd.to_datetime(date_filter_range[1])
            due_dates = safe_datetime_convert(filtered_data[final_column_mapping["due_date"]])
            date_mask = (due_dates >= start_date) & (due_dates <= end_date)
            filtered_data = filtered_data[date_mask]
        
        # 顯示篩選結果
        if len(filtered_data) != len(df):
            st.info(f"🔍 篩選後顯示 {len(filtered_data):,} 筆記錄（原始：{len(df):,} 筆）")
        
        # SWI排除邏輯
        exclude_swi_condition = pd.Series(True, index=filtered_data.index)
        if final_column_mapping.get("ship_type"):
            exclude_swi_condition = filtered_data[final_column_mapping["ship_type"]] != "SWI-寄庫"
        
        # 執行各項分析
        if not filtered_data.empty:
            enhanced_shipment_analysis(filtered_data, final_column_mapping)
            st.markdown("---")
            enhanced_delivery_performance(filtered_data, final_column_mapping, exclude_swi_condition)
            st.markdown("---") 
            enhanced_area_analysis(filtered_data, final_column_mapping, exclude_swi_condition, topn_cities_setting)
            st.markdown("---")
            enhanced_loading_analysis(filtered_data, final_column_mapping, exclude_swi_condition)
        else:
            st.warning("⚠️ 篩選條件過於嚴格，無數據可顯示")

# ========================
# 程式入口
# ========================

if __name__ == "__main__":
    main()
