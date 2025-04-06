import openpyxl
from datetime import datetime, date
from tkinter import Tk, filedialog
from collections import defaultdict
import os
import jpholiday  # 日本の祝日判定ライブラリ

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
    loss_last_values = defaultdict(float)  # 損失電力の最終データ

    # データを逆順で処理（行の終わりが先頭データ）
    rows = list(sheet.iter_rows(min_row=2))  # ヘッダーをスキップ
    for row in reversed(rows):  # 行を逆順に処理
        date_cell = row[0].value  # 1列目の日付データ

        # 列数チェック（31列目と32列目が存在するか確認）
        if len(row) <= 31:
            continue

        value_cell = row[30].value  # 31列目のデータ
        loss_cell = row[31].value  # 32列目のデータ

        # date_cell の確認と変換
        if isinstance(date_cell, str):
            try:
                date_cell = datetime.strptime(date_cell, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue
        elif not isinstance(date_cell, datetime):
            continue

        # value_cell の確認と変換
        try:
            value_cell = float(value_cell)
        except (ValueError, TypeError):
            continue

        # loss_cell の確認と変換
        try:
            if isinstance(loss_cell, str):
                loss_cell = loss_cell.replace(",", "")  # カンマを削除
            loss_cell = float(loss_cell)  # 数値に変換
        except (ValueError, TypeError):
            continue

        # 時間部分を抽出
        hour = date_cell.hour

        # 土日祝の判定
        is_weekend_or_holiday = date_cell.weekday() >= 5 or jpholiday.is_holiday(date_cell.date())

        # 時間帯ごとの最終データを更新
        for label, (start, end) in time_ranges.items():
            if start <= hour < end:
                if label == "9時～17時" and is_weekend_or_holiday:
                    # 土日祝の9時～17時は「ホームタイム」に分類
                    time_last_values["ホームタイム"] = value_cell
                    loss_last_values["ホームタイム"] = loss_cell
                else:
                    time_last_values[label] = value_cell
                    loss_last_values[label] = loss_cell

    # 時間帯ごとの増加量を計算
    time_increments = defaultdict(float)
    loss_increments = defaultdict(float)
    previous_value = None
    previous_loss = None
    for label in time_ranges.keys():
        if previous_value is not None:
            time_increments[label] = time_last_values[label] - previous_value
        else:
            time_increments[label] = time_last_values[label]  # 初回はそのまま
        previous_value = time_last_values[label]

        if previous_loss is not None:
            loss_increments[label] = loss_last_values[label] - previous_loss
        else:
            loss_increments[label] = loss_last_values[label]  # 初回はそのまま
        previous_loss = loss_last_values[label]

    # タイムゾーンごとの集計
    time_zone_totals = {
        "ナイトタイム": time_increments["0時～7時"] + time_increments["23時～24時"],
        "ホームタイム": time_increments["7時～9時"] + time_increments["17時～23時"] + time_increments.get("ホームタイム", 0),
        "デイタイム": time_increments["9時～17時"] if not is_weekend_or_holiday else 0
    }
    loss_zone_totals = {
        "ナイトタイム": loss_increments["0時～7時"] + loss_increments["23時～24時"],
        "ホームタイム": loss_increments["7時～9時"] + loss_increments["17時～23時"] + loss_increments.get("ホームタイム", 0),
        "デイタイム": loss_increments["9時～17時"] if not is_weekend_or_holiday else 0
    }

    return time_zone_totals, loss_zone_totals

# メイン処理
file_path = select_file()

# 選択したファイルのディレクトリを取得
directory = os.path.dirname(file_path)

# 同じディレクトリ内のすべてのExcelファイルを検索
all_files = [
    os.path.join(directory, f) for f in os.listdir(directory)
    if f.endswith(".xlsx") or f.endswith(".xlsm")
]

# 選択されたファイル名を表示
#print(f"選択されたファイル: {file_path}")

# 処理対象のファイルを表示
print("処理対象のファイル:")
for f in all_files:
    print(f"  - {f}")

# 全ファイルのデータを計算
total_time_zone_totals = defaultdict(float)
total_loss_zone_totals = defaultdict(float)
for file in all_files:
    #print(f"ファイルを処理中: {file}")
    time_zone_totals, loss_zone_totals = calculate_time_data(file)
    for zone, total in time_zone_totals.items():
        total_time_zone_totals[zone] += total
    for zone, loss in loss_zone_totals.items():
        total_loss_zone_totals[zone] += loss

# 各タイムゾーンの集計を表示
print("\nタイムゾーンごとの集計:")
total_cost = 0
rates = {"デイタイム": 34.06, "ホームタイム": 26.00, "ナイトタイム": 16.11}
saiene_cost = 3.49  # 再エネ賦課金単価
nenryou_cost = 0.06  # 燃料費調整単価
total_renewable_energy_cost = 0
total_fuel_adjustment_cost = 0

for zone, total in total_time_zone_totals.items():
    cost = total * rates[zone]
    renewable_energy_cost = total * saiene_cost  # 再エネ賦課金単価
    fuel_adjustment_cost = total * nenryou_cost  # 燃料費調整単価

    total_cost += cost
    total_renewable_energy_cost += renewable_energy_cost
    total_fuel_adjustment_cost += fuel_adjustment_cost

    # 計算結果のみを表示
    print(f"{zone}: {total:.2f} kWh, 金額: {cost:.2f} 円, "
          f"再エネ賦課金: {renewable_energy_cost:.2f} 円, 燃料費調整額: {fuel_adjustment_cost:.2f} 円")

# 損失電力の集計を表示
print("\n損失電力の集計:")
total_loss_cost = 0
total_loss_renewable_cost = 0
total_loss_fuel_adjustment_cost = 0

for zone, loss in total_loss_zone_totals.items():
    loss_cost = loss * rates[zone]
    loss_renewable_cost = loss * saiene_cost  # 損失電力分の再エネ賦課金
    loss_fuel_adjustment_cost = loss * nenryou_cost  # 損失電力分の燃料費調整額

    total_loss_cost += loss_cost
    total_loss_renewable_cost += loss_renewable_cost
    total_loss_fuel_adjustment_cost += loss_fuel_adjustment_cost

    # 計算結果のみを表示
    print(f"{zone}: {loss:.2f} kWh, 損失金額: {loss_cost:.2f} 円, "
          f"再エネ賦課金: {loss_renewable_cost:.2f} 円, 燃料費調整額: {loss_fuel_adjustment_cost:.2f} 円")

# 再エネ賦課金単価と燃料費調整単価を計算
total_energy = sum(total_time_zone_totals.values())

print(f"\n再エネ賦課金: {total_renewable_energy_cost:.2f} 円")
print(f"燃料費調整額: {total_fuel_adjustment_cost:.2f} 円")

# 合計金額を表示
total_cost += total_renewable_energy_cost + total_fuel_adjustment_cost
print(f"\n太陽光発電による節約金額: {total_cost:.2f} 円")

# 損失を含む合計金額を計算
total_loss_adjustment = total_loss_cost + total_loss_renewable_cost + total_loss_fuel_adjustment_cost
adjusted_total_cost = total_cost - total_loss_adjustment
print(f"損失電力を含む経済効果: {adjusted_total_cost:.2f} 円")
