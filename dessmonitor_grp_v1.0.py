import os
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from tkinter import Tk
from tkinter.filedialog import askdirectory
from collections import defaultdict
from matplotlib import rcParams
import matplotlib.colors as mcolors

# 日本語フォントの設定
rcParams['font.family'] = 'MS Gothic'  # Windows環境では「MS Gothic」を使用

# フォルダー選択ウィンドウを表示
Tk().withdraw()  # Tkinterのルートウィンドウを非表示にする
folder_path = askdirectory(title="フォルダーを選択してください")  # フォルダーを選択

if not folder_path:
    print("フォルダーが選択されませんでした。")
    exit()

# フォルダー内のすべてのxlsxファイルを再帰的に取得
file_paths = []
for root, _, files in os.walk(folder_path):  # os.walkでフォルダー内を再帰的に探索
    for file in files:
        if file.endswith(".xlsx"):
            file_paths.append(os.path.join(root, file))

# データを格納する辞書 (キー: ファイル名の共通部分, 値: 日付と値のリスト)
data_dict = defaultdict(list)
usage_dict = defaultdict(list)  # 電力使用量を格納する辞書

# フォルダー内のすべてのファイルを処理
for file_path in file_paths:
    try:
        base_name = os.path.basename(file_path)[:-10]
        df = pd.read_excel(file_path, header=0)
        if df.shape[1] >= 31:
            date_str = df.iloc[1, 0]
            value_str = df.iloc[1, 29]
            usage_str = df.iloc[1, 30]
            try:
                date = pd.to_datetime(date_str).date()
                value = float(value_str)
                usage = float(usage_str)
                data_dict[base_name].append((date, value))
                usage_dict[base_name].append((date, usage))
            except Exception as e:
                print(f"データ変換エラー: {file_path}, 値: {date_str}, エラー: {e}")
                continue
    except Exception as e:
        print(f"ファイル読み込みエラー: {file_path}, エラー: {e}")

# 日付ごとに発電量と電力使用量を合算
stacked_data = defaultdict(lambda: defaultdict(float))
stacked_usage = defaultdict(lambda: defaultdict(float))

for base_name, records in data_dict.items():
    for date, value in records:
        stacked_data[date][base_name] += value

for base_name, records in usage_dict.items():
    for date, usage in records:
        stacked_usage[date][base_name] += usage

# 日付順にソート
sorted_dates = sorted(stacked_data.keys())
base_names = list(data_dict.keys())

# 積上げ棒グラフのデータ準備
stacked_values = {base_name: [] for base_name in base_names}
stacked_usages = {base_name: [] for base_name in base_names}

for date in sorted_dates:
    for base_name in base_names:
        stacked_values[base_name].append(stacked_data[date][base_name])
        stacked_usages[base_name].append(stacked_usage[date][base_name])

# 発電量と電力使用量をx軸に対して交互に表示
fig, ax = plt.subplots(figsize=(12, 6))  # 図を作成
fig.canvas.manager.set_window_title("Dessmonitor月間チャート - 発電量と電力使用量")  # ウィンドウのタイトルを設定

x = range(len(sorted_dates))  # x軸の位置
width = 0.4  # 棒グラフの幅

# カスタムカラーマップを作成
orange_gradient = mcolors.LinearSegmentedColormap.from_list("orange_gradient", ["#FFA500", "#FF4500"])
blue_gradient = mcolors.LinearSegmentedColormap.from_list("blue_gradient", ["#87CEEB", "#0000FF"])

# 発電量と電力消費量の色リストを初期化
bar_colors_generation = []  # 発電量の色リスト
bar_colors_usage = []       # 電力消費量の色リスト

# 発電量の積上げ棒グラフ
bottom_values = [0] * len(sorted_dates)  # 発電量の積上げ基準値
for i, base_name in enumerate(base_names):
    color = orange_gradient((len(base_names) - i - 1) / len(base_names))  # 下が濃く、上が薄くなるオレンジグラデーション
    bar_colors_generation.append(color)  # 色をリストに追加
    ax.bar(
        [pos - width / 2 for pos in x],  # 発電量を左側に配置
        stacked_values[base_name],
        bottom=bottom_values,  # 前の棒グラフの上に積上げ
        width=width,
        label=f"{base_name} (発電量)",
        color=color
    )
    # 基準値を更新
    bottom_values = [bottom + value for bottom, value in zip(bottom_values, stacked_values[base_name])]

# 電力消費量の積上げ棒グラフ
bottom_usages = [0] * len(sorted_dates)  # 電力消費量の積上げ基準値
for i, base_name in enumerate(base_names):
    color = blue_gradient((len(base_names) - i - 1) / len(base_names))  # 下が濃く、上が薄くなる青グラデーション
    bar_colors_usage.append(color)  # 色をリストに追加
    ax.bar(
        [pos + width / 2 for pos in x],  # 電力消費量を右側に配置
        stacked_usages[base_name],
        bottom=bottom_usages,  # 前の棒グラフの上に積上げ
        width=width,
        label=f"{base_name} (電力使用量)",
        color=color,
        alpha=0.7  # 電力消費量を少し透明にする
    )
    # 基準値を更新
    bottom_usages = [bottom + usage for bottom, usage in zip(bottom_usages, stacked_usages[base_name])]

# 軸ラベルとタイトルの設定
ax.set_xlabel("日付")
ax.set_ylabel("量 (kWh)")
ax.set_title("日別 総発電量と総電力使用量")
ax.set_xticks(x)
ax.set_xticklabels(sorted_dates, rotation=45)  # 日付をすべて表示し、45度回転

# y軸の上限をデータに基づいて設定
max_generation = max([sum(values) for values in zip(*stacked_values.values())])  # 発電量の最大値
max_usage = max([sum(values) for values in zip(*stacked_usages.values())])  # 電力使用量の最大値
y_max = max(max_generation, max_usage)  # 発電量と電力使用量の最大値を取得
y_max = int(y_max) + 1  # 最大値を切り上げて1kWh単位で調整

# y軸の範囲と目盛りを設定
ax.set_yticks(range(0, y_max + 1, 1))  # 1kWh単位で目盛りを設定
ax.set_ylim(0, y_max)  # y軸の範囲を0から設定値に設定

# y軸にメモリラインを追加
ax.yaxis.grid(True, linestyle='--', alpha=0.7)  # 点線でメモリラインを追加

# 凡例の設定
from matplotlib.patches import Patch

# 発電量と電力使用量をグループ化して凡例を作成
legend_elements_generation = [
    Patch(facecolor=bar_colors_generation[i], label=f"{base_name} (発電量)")
    for i, base_name in enumerate(base_names)
]
legend_elements_usage = [
    Patch(facecolor=bar_colors_usage[i], label=f"{base_name} (電力使用量)", alpha=0.7)
    for i, base_name in enumerate(base_names)
]

# 凡例を設定
ax.legend(
    handles=legend_elements_generation + legend_elements_usage,  # 発電量と電力使用量を結合
    title="ファイルグループ",
    loc="upper center",
    bbox_to_anchor=(0.5, -0.3),  # グラフの下に配置
    fontsize="small",  # フォントサイズを小さく設定
    title_fontsize="medium",  # タイトルのフォントサイズ
    ncol=2  # 2列に設定
)

plt.tight_layout()
plt.show()
