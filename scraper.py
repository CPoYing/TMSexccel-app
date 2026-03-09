import pandas as pd
import requests
import sqlite3
from datetime import datetime

url = "http://www.wire.com.tw/cooperate.asp"

try:
    # 先用 requests 抓，明確指定 Big5（實測最穩）
    resp = requests.get(url, timeout=15)
    resp.raise_for_status()
    resp.encoding = 'big5'          # 或試 'cp950' 如果還有亂碼

    # 用 pandas 直接讀 table（通常只有一個 table）
    dfs = pd.read_html(resp.text)
    if not dfs:
        raise ValueError("頁面沒有找到表格")

    df = dfs[0]
    # 強制欄位名稱（防頁面改版）
    df.columns = ["貨品代號", "貨品名稱", "現有數量", "單位", "庫位", "包裝_單位", "包裝_數量"]

    # 清理：數量轉 int，空值填 0
    df["現有數量"] = pd.to_numeric(df["現有數量"], errors='coerce').fillna(0).astype(int)

    # 加日期
    today = datetime.now().strftime("%Y-%m-%d")
    df["日期"] = today

    # 連資料庫，append + 自動去重（primary key 已設）
    conn = sqlite3.connect('inventory.db')
    df.to_sql('history', conn, if_exists='append', index=False)
    conn.close()

    print(f"[{today}] 成功記錄 {len(df)} 筆資料")

except Exception as e:
    print(f"抓取/儲存失敗: {e}")
    # 可加 email 或 Line 通知（未來擴充）
