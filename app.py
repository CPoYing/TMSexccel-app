import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

st.title("華新電線庫存監控")

url = "http://www.wire.com.tw/cooperate.asp"

try:
    response = requests.get(url, timeout=10)
    response.raise_for_status()  # 檢查 HTTP 錯誤
    soup = BeautifulSoup(response.text, 'html.parser')
    
    table = soup.find('table')
    if not table:
        st.error("找不到表格，請檢查網頁結構是否改變")
        st.stop()
    
    rows = table.find_all('tr')
    data = []
    for row in rows[1:]:  # 跳過表頭
        cols = row.find_all('td')
        if len(cols) >= 5:  # 至少有前5欄
            data.append([col.get_text(strip=True) for col in cols[:7]])
    
    if not data:
        st.warning("表格無資料")
    else:
        df = pd.DataFrame(data, columns=['貨品代號', '貨品名稱', '現有數量', '單位', '庫位', '包裝_單位', '包裝_數量'])
        df['現有數量'] = pd.to_numeric(df['現有數量'], errors='coerce')
        
        st.dataframe(df)
        
        low_stock = df[df['現有數量'] < 1000].sort_values('現有數量')
        if not low_stock.empty:
            st.subheader("低庫存警報 (< 1000 米)")
            st.dataframe(low_stock)
        else:
            st.success("目前無低庫存品項")

except Exception as e:
    st.error(f"抓取失敗：{str(e)}")
    st.info("可能原因：網路問題、網站維護、或頁面結構改變")
