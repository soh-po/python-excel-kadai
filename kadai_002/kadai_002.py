import pandas as pd

# ファイル名
file_name = "業績.xlsx"

# データフレームの作成
data = {
    '日付':
      ['2023-05-17', '2023-05-18', '2023-05-19', '2023-05-20', '2023-05-21'],
    '社員名': ['山田', '佐藤', '鈴木', '田中', '高橋'],
    '売上': [100, 200, 150, 300, 250],
    '部門': ['メーカー', '代理店', 'メーカー', '商社', '代理店'],
}
df = pd.DataFrame(data)

# "平均売上"列を作成
df["平均売上"] = df["売上"].mean()
average_sales = df["売上"].mean()

# 業績ランク付けを行う関数
def performance(rank):
    result = ""
    if rank >= (average_sales + 50):
        result = "A"
    elif rank >= average_sales:
        result = "B"
    else:
        result = "C"
    return result

# "業績ランク"列を作成
df["業績ランク"] = df["売上"].apply(performance)

# ファイルの作成～クローズ
writer = pd.ExcelWriter(file_name)
df.to_excel(writer, sheet_name="Sheet1", index=False)
writer.close()
