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
