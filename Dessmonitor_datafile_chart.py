import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import numpy as np

# 日本語フォント設定
font_path = 'c:/windows/Fonts/meiryo.ttc'
jp_font = fm.FontProperties(fname=font_path)

# ファイル選択関数
def select_file(title):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(title=title, filetypes=[('Excel files', '*.xlsx')])

# データ取得関数
def extract_data(sheet, col, is_string=False):
    data = []
    for row in sheet.iter_rows(min_row=1, min_col=col, max_col=col):
        for cell in row:
            if cell.value is not None:  # 空セルの処理を防ぐ
                if is_string:  # 文字列として処理
                    data.append(str(cell.value))
                else:  # 数値として処理
                    try:
                        data.append(int(cell.value))  # 数値に変換
                    except ValueError:
                        print(f"数値に変換できないデータをスキップしました: {cell.value}")
    return data


# グラフ描画関数
def plot_graph(x, y_data, labels, colors):
    for y, label, color in zip(y_data, labels, colors):
        plt.plot(x, y, label=label, linestyle='-', color=color, linewidth=0.2 if 'Total' not in label else 1)

# データ読み込み
file1_path, file2_path = select_file("データファイル1"), select_file("データファイル2")
sheet1, sheet2 = load_workbook(file1_path).active, load_workbook(file2_path).active

data0 = extract_data(sheet1, 1, is_string=True)  # 1列目
data1, data2 = extract_data(sheet1, 15), extract_data(sheet2, 15)  # 14列目
data4, data5 = extract_data(sheet1, 22), extract_data(sheet2, 22)  # 21列目

# データ調整
min_length = min(map(len, [data0, data1, data2, data4, data5]))
data0, data1, data2, data4, data5 = [d[:min_length] for d in [data0, data1, data2, data4, data5]]
data3 = [a + b for a, b in zip(data1, data2)]
data6 = [a + b for a, b in zip(data4, data5)]

# データ反転（左右逆）
data0 = data0[::-1]
data1 = data1[::-1]
data2 = data2[::-1]
data3 = data3[::-1]
data4 = data4[::-1]
data5 = data5[::-1]
data6 = data6[::-1]

# グラフ描画設定
sns.set(style='darkgrid', palette='winter_r')
plt.figure(figsize=(15, 8))
plt.xticks(rotation=80)
plt.yticks(rotation=-10)
# ウィンドウの位置を指定
manager = plt.get_current_fig_manager()
manager.window.state("zoomed")  # フルスクリーン設定
# プロット
plot_graph(data0, [data1, data2, data3, data4, data5, data6],
           ['PV1_power', 'PV2_power', 'PV_Total_power', 'Inv1_power', 'Inv2_power', 'INV_Total_power'],
           ['red', 'red', 'red', 'blue', 'blue', 'blue'])

# 最大値マーク
for data, label,color in [(data3, "Max_PV_power","red"), (data6, "Max_INV_power","blue")]:
    max_idx = np.argmax(data)
    plt.scatter(data0[max_idx], data[max_idx], color=color, label=f"{label}")
    plt.text(data0[max_idx], data[max_idx], f"({data0[max_idx]}, {data[max_idx]}W)", fontsize=10, color=color, ha="right")

# グラフ設定
plt.title('電力量', fontsize=20, fontproperties=jp_font)
plt.xlabel('日時', fontsize=12, fontproperties=jp_font)
plt.ylabel('電力', fontsize=14, fontproperties=jp_font)
plt.legend(bbox_to_anchor=(1, 1), loc='upper left')
plt.grid(True)
plt.gca().xaxis.set_major_locator(plt.MultipleLocator(10))
plt.gca().yaxis.set_major_locator(plt.MultipleLocator(500))
plt.show()
