import pandas as pd
from pathlib import Path

file = Path.cwd()/'初期値.xlsx'
df = pd.read_excel(file, engine='openpyxl', encoding='cp932', usecols=[0,1])
df = df.set_index('入力項目')
d = df.to_dict(orient='index')
print(d['支払基金']['初期値'])