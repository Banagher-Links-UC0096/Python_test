import openpyxl
from datetime import datetime
from tkinter import Tk, filedialog
from collections import defaultdict

# ファイル選択ダイアログを表示
def select_file():
    root = Tk()
    root.withdraw()  # Tkinterウインドウを非表示にする
    file_path = filedialog.askopenfilename(
        title="Excelファイルを選択してください",
        filetypes=[("Excelファイル", "*.xlsx *.xlsm")]
    )
    return file_path

# ファイルを選択
file_path = select_file()
if not file_path:
    print("ファイルが選択されませんでした。")
    exit()

# Excelファイルを読み込む
wb = openpyxl.load_workbook(file_path)
sheet = wb.active

# 時間帯ごとの最終データを格納する辞書
time_ranges = {
    "0時～7時": (0, 7),
    "7時～9時": (7, 9),
    "9時～17時": (9, 17),
    "17時～23時": (17, 23),
    "23時～24時": (23, 24)
}
time_last_values = defaultdict(float)

# データを逆順で処理（行の終わりが先頭データ）
rows = list(sheet.iter_rows(min_row=2))  # ヘッダーをスキップ
for row in reversed(rows):  # 行を逆順に処理
    date_cell = row[0].value  # 1列目の日付データ
    value_cell = row[30].value  # 31列目のデータ

    # date_cell の確認と変換
    if isinstance(date_cell, str):
        try:
            date_cell = datetime.strptime(date_cell, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            print(f"date_cell を datetime に変換できませんでした: {date_cell}")
            continue
    elif not isinstance(date_cell, datetime):
        print(f"date_cell が datetime 型ではありません: {date_cell} (type: {type(date_cell)})")
        continue

    # value_cell の確認と変換
    try:
        if isinstance(value_cell, str):
            value_cell = value_cell.replace(",", "")
            value_cell = float(value_cell)
        else:
            value_cell = float(value_cell)
    except (ValueError, TypeError):
        print(f"value_cell を数値に変換できませんでした: {value_cell}")
        continue

    # 時間部分を抽出
    hour = date_cell.hour

    # 時間帯ごとの最終データを更新
    for label, (start, end) in time_ranges.items():
        if start <= hour < end:
            time_last_values[label] = value_cell

# 時間帯ごとの増加量を計算
time_increments = defaultdict(float)
previous_value = None
for label, last_value in time_last_values.items():
    if previous_value is not None:
        time_increments[label] = last_value - previous_value
    else:
        time_increments[label] = 0  # 初回は増加量を0とする
    previous_value = last_value

# タイムゾーンごとの集計
time_zone_totals = {
    "ナイトタイム": time_increments["0時～7時"] + time_increments["23時～24時"],
    "ホームタイム": time_increments["7時～9時"] + time_increments["17時～23時"],
    "デイタイム": time_increments["9時～17時"]
}

# 各時間帯の増加量を表示
print("各時間帯の増加量:")
for label, increment in time_increments.items():
    print(f"{label}: {increment:.2f} kWh")

# 各タイムゾーンの集計を表示
print("\nタイムゾーンごとの集計:")
total_cost = 0
rates = {"デイタイム": 34.06, "ホームタイム": 26.00, "ナイトタイム": 16.11}

for zone, total in time_zone_totals.items():
    cost = total * rates[zone]
    total_cost += cost
    print(f"{zone}: {total:.2f} kWh, 金額: {cost:.2f} 円")

# 合計金額を表示
print(f"\n合計金額: {total_cost:.2f} 円")