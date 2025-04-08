import os
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
from tkinter import Tk
from tkinter.filedialog import askdirectory
from collections import defaultdict

# フォルダー選択ウィンドウを表示
Tk().withdraw()  # Tkinterのルートウィンドウを非表示にする
folder_path = askdirectory(title="フォルダーを選択してください")  # フォルダーを選択

if not folder_path:
    print("フォルダーが選択されませんでした。")
    exit()

# フォルダー内のすべてのxlsxファイルを取得
file_paths = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith(".xlsx")]

# 選択されたファイル名をコンソールに表示
print("フォルダー内のファイル:")
for file_path in file_paths:
    print(os.path.basename(file_path))

# データを格納する辞書 (キー: ファイル名の共通部分, 値: 日付と値のリスト)
data_dict = defaultdict(list)

# フォルダー内のすべてのファイルを処理
for file_path in file_paths:
    try:
        # ファイル名の後半10文字を削除した共通部分を取得
        base_name = os.path.basename(file_path)[:-10]
        
        # Excelファイルを読み込む
        df = pd.read_excel(file_path, header=0)
        if df.shape[1] >= 30:  # 必要な列が存在するか確認
            # 1列目と30列目の2行目のデータを取得
            date_str = df.iloc[1, 0]  # 2行目の1列目
            value_str = df.iloc[1, 29]  # 2行目の30列目
            try:
                # データ変換
                date = pd.to_datetime(date_str).date()  # 自動的に日付を解析
                value = float(value_str)
                data_dict[base_name].append((date, value))
            except Exception as e:
                print(f"データ変換エラー: {file_path}, 値: {date_str}, エラー: {e}")
                continue
    except Exception as e:
        print(f"ファイル読み込みエラー: {file_path}, エラー: {e}")

# 日付ごとに発電量を合算 (積上げ棒グラフ用)
stacked_data = defaultdict(lambda: defaultdict(float))
for base_name, records in data_dict.items():
    for date, value in records:
        stacked_data[date][base_name] += value

# 日付順にソート
sorted_dates = sorted(stacked_data.keys())
base_names = list(data_dict.keys())

# 積上げ棒グラフのデータ準備
stacked_values = {base_name: [] for base_name in base_names}
for date in sorted_dates:
    for base_name in base_names:
        stacked_values[base_name].append(stacked_data[date][base_name])

# 積上げ棒グラフを作成
fig = plt.figure(figsize=(12, 6))  # 図を作成
fig.canvas.manager.set_window_title("Dessmonitor月間チャート")  # ウィンドウのタイトルを設定

#plt.suptitle("Dessmonitor月間チャート", fontsize=16)  # グラフ全体のタイトルを設定

bottom = [0] * len(sorted_dates)
colors = plt.cm.tab20.colors  # カラーマップを使用して色を設定

for i, base_name in enumerate(base_names):
    plt.bar(sorted_dates, stacked_values[base_name], bottom=bottom, label=base_name, color=colors[i % len(colors)])
    bottom = [b + v for b, v in zip(bottom, stacked_values[base_name])]

plt.xlabel("Date")
plt.ylabel("Total Generation (kWh)")
plt.title("Daily Total Power Generation (Stacked by File Group)")
plt.xticks(sorted_dates, rotation=45)  # 日付をすべて表示し、45度回転

# 凡例をグラフの下に配置（縦並びで小さめに表示）
plt.legend(
    title="File Groups",
    loc="upper center",
    bbox_to_anchor=(0.5, -0.3),  # グラフの下に配置
    fontsize="small",  # フォントサイズを小さく設定
    title_fontsize="medium",  # タイトルのフォントサイズ
    ncol=1  # 縦並び（1列）に設定
)

plt.tight_layout()
plt.show()