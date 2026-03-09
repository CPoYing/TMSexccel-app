import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

url = "http://www.wire.com.tw/cooperate.asp"

response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')

table = soup.find('table')  # 找到唯一的 table
rows = table.find_all('tr')

data = []
for row in rows[1:]:  # 跳過表頭
    cols = row.find_all('td')
    if len(cols) == 7:
        data.append([col.get_text(strip=True) for col in cols])

df = pd.DataFrame(data, columns=['貨品代號', '貨品名稱', '現有數量', '單位', '庫位', '包裝_單位', '包裝_數量'])

# 轉數字
df['現有數量'] = pd.to_numeric(df['現有數量'], errors='coerce')

# 例如：篩選低於 1000 米的
low_stock = df[df['現有數量'] < 1000]

# 存檔
today = datetime.now().strftime('%Y-%m-%d')
df.to_csv(f'inventory_{today}.csv', index=False, encoding='utf-8-sig')
low_stock.to_csv(f'low_stock_{today}.csv', index=False, encoding='utf-8-sig')

# 未來可加：寄 email、推 Line Notify、寫入 Google Sheets API
print("更新完成！", len(df), "筆資料")
