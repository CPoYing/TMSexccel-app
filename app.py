import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd

st.title("華新電線庫存監控（Big5 編碼修正）")

url = "http://www.wire.com.tw/cooperate.asp"

try:
    response = requests.get(url, timeout=15)
    response.raise_for_status()
    
    # 關鍵修正：用 Big5 解碼（或 cp950 也行，兩者非常接近）
    # 如果 big5 還有少許亂碼，換成 'cp950' 再試
    html_content = response.content.decode('big5')
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    table = soup.find('table')
    if not table:
        st.error("找不到表格，頁面結構可能改變")
        st.stop()
    
    rows = table.find_all('tr')
    data = []
    headers = None
    
    for i, row in enumerate(rows):
        cols = row.find_all(['th', 'td'])
        if not cols:
            continue
        cols_text = [col.get_text(strip=True) for col in cols]
        
        if i == 0 and len(cols_text) >= 5:  # 第一行通常是表頭
            headers = cols_text
        elif headers and len(cols_text) == len(headers):
            data.append(cols_text)
    
    if data and headers:
        df = pd.DataFrame(data, columns=headers)
        
        # 轉數字（忽略錯誤）
        if '現有數量' in df.columns:
            df['現有數量'] = pd.to_numeric(df['現有數量'], errors='coerce')
        
        # 美化顯示
        st.subheader("完整庫存表格")
        st.dataframe(
            df.style.format({"現有數量": "{:,.0f}" if pd.notna else "{}"}),
            use_container_width=True
        )
        
        # 低庫存警報（自訂 < 1000 米，可改）
        if '現有數量' in df.columns:
            low_stock = df[df['現有數量'] < 1000].sort_values('現有數量')
            if not low_stock.empty:
                st.subheader(f"低庫存警報（< 1000 米，共 {len(low_stock)} 項）")
                st.dataframe(low_stock, use_container_width=True)
            else:
                st.success("目前所有品項庫存充足！")
    else:
        st.warning("無法從表格提取資料，請檢查網頁")
        
except requests.exceptions.RequestException as e:
    st.error(f"網路抓取失敗：{e}")
except UnicodeDecodeError as e:
    st.error(f"編碼錯誤：{e} → 試試把 decode('big5') 改成 decode('cp950') 或 decode('utf-8')")
except Exception as e:
    st.error(f"其他錯誤：{e}")
