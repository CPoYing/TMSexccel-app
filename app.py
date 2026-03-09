import streamlit as st
import pandas as pd
import sqlite3
import os

# ── 資料庫初始化 ──
DB_FILE = 'inventory.db'

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS history (
            日期 TEXT,
            貨品代號 TEXT,
            貨品名稱 TEXT,
            現有數量 INTEGER,
            單位 TEXT,
            庫位 TEXT,
            包裝_單位 TEXT,
            包裝_數量 INTEGER,
            PRIMARY KEY (日期, 貨品代號, 庫位)
        )
    ''')
    conn.commit()
    conn.close()
    st.session_state['db_initialized'] = True

if not os.path.exists(DB_FILE) or 'db_initialized' not in st.session_state:
    with st.spinner("初始化資料庫中..."):
        init_db()
        st.success("資料庫已建立/確認！")

# 現在再讀取
conn = sqlite3.connect(DB_FILE)
try:
    df = pd.read_sql_query("SELECT * FROM history", conn)
except Exception as e:
    st.error(f"讀取資料庫失敗：{str(e)}")
    df = pd.DataFrame()  # 防崩潰
finally:
    conn.close()

if df.empty:
    st.warning("目前資料庫無歷史記錄，請執行 scraper.py 至少一次來累積資料。")
    # 可加提示：如何執行 scraper（或加按鈕手動抓取）
else:
    # 你的原程式碼繼續...
    st.title("華新電線庫存歷史追蹤系統")
    # ... KPI、tabs 等
