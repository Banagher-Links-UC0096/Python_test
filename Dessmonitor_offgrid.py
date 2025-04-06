import openpyxl
from datetime import datetime
from tkinter import Tk, filedialog
from collections import defaultdict
import os

# ファイル選択ダイアログを表示
def select_file():
    root = Tk()
    root.withdraw()  # Tkinterウインドウを非表示にする
    file_path = filedialog.askopenfilename(
        title="Excelファイルを選択してください",
        filetypes=[("Excelファイル", "*.xlsx *.xlsm")]
    )
    if not file_path:
        print("ファイルが選択されませんでした。")
        exit()
    print(f"選択されたファイル: {file_path}")
    return file_path

# ファイル名の拡張子を除いた右から10文字を取得する関数
def get_file_suffix_without_extension(file_path):
    file_name = os.path.splitext(os.path.basename(file_path))[0]  # 拡張子を除いたファイル名
    return file_name[-10:]  # 右から10文字を取得

# 時間帯ごとのデータを計算する関数
def calculate_time_data(file_path):
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

        # 列数チェック（31列目が存在するか確認）
        if len(row) <= 30:
            print(f"警告: ファイル {file_path} に31列目のデータが存在しません。スキップします。")
            continue

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

    return time_zone_totals

# メイン処理
file_path = select_file()

# 選択したファイルの後半10文字を取得（拡張子を除く）
file_suffix = get_file_suffix_without_extension(file_path)

# 同じディレクトリ内で後半10文字が一致する別のファイルを検索
directory = os.path.dirname(file_path)
matching_files = []
for f in os.listdir(directory):
    if f.endswith(".xlsx") or f.endswith(".xlsm"):
        if f != os.path.basename(file_path):
            if get_file_suffix_without_extension(os.path.join(directory, f)) == file_suffix:
                matching_files.append(os.path.join(directory, f))

# 選択されたファイル名を表示
print(f"選択されたファイル: {file_path}")

# 一致する別のファイル名を表示
if matching_files:
    print("一致するファイル:")
    for matching_file in matching_files:
        print(f"  - {matching_file}")
else:
    print("一致するファイルは見つかりませんでした。")

# 選択したファイルのデータを計算
total_time_zone_totals = defaultdict(float)
time_zone_totals = calculate_time_data(file_path)
for zone, total in time_zone_totals.items():
    total_time_zone_totals[zone] += total

# 一致する別のファイルがあれば計算
for matching_file in matching_files:
    print(f"一致するファイルを処理中: {matching_file}")
    time_zone_totals = calculate_time_data(matching_file)
    for zone, total in time_zone_totals.items():
        total_time_zone_totals[zone] += total

# 各タイムゾーンの集計を表示
print("\nタイムゾーンごとの集計:")
total_cost = 0
rates = {"デイタイム": 34.06, "ホームタイム": 26.00, "ナイトタイム": 16.11}

for zone, total in total_time_zone_totals.items():
    cost = total * rates[zone]
    total_cost += cost
    print(f"{zone}: {total:.2f} kWh, 金額: {cost:.2f} 円")

# 合計金額を表示
print(f"\n合計金額: {total_cost:.2f} 円")
