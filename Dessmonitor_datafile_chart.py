import csv
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
import numpy as np
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os


# 日本語フォント設定
font_path = 'c:/windows/Fonts/meiryo.ttc'
jp_font = fm.FontProperties(fname=font_path)

# ファイル選択関数
def select_file(title):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(title=title, filetypes=[('CSV or Excel files', '*.csv;*.xlsx')])

# CSVデータ取得関数
def extract_data_csv(filepath, col, is_string=False):
    data = []
    with open(filepath, mode='r', encoding='utf-8') as file:
        reader = csv.reader(file)
        for row in reader:
            try:
                value = row[col - 1]  # 0-indexed
                if is_string:
                    data.append(value)
                else:
                    # 余分な文字を取り除く処理
                    value = value.strip("[]")  # "["や"]"を削除
                    data.append(int(value))  # 数値に変換
            except ValueError:
                print(f"数値に変換できないデータをスキップしました: {value}")
            except IndexError:
                print("指定された列番号が範囲外です")
    return data

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
                        data.append(int(float(cell.value)))  # 数値に変換
                    except ValueError:
                        print(f"数値に変換できないデータをスキップしました: {cell.value}")
    return data

# ファイル検索関数（後半10文字一致かつ異なるファイルを探す）
def find_matching_file(directory, base_filename_tail, file1_path):
    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)  # フルパスを作成
        # ファイル名の後半10文字が一致し、かつ1つ目のファイルと異なるファイルを選択
        if file.endswith(base_filename_tail) and os.path.abspath(file_path) != os.path.abspath(file1_path):
            return file_path
    return None  # 一致するファイルが見つからない場合

# グラフ描画関数
def plot_graph(x, y_data, labels, colors):
    for y, label, color in zip(y_data, labels, colors):
        plt.plot(x, y, label=label, linestyle='-', color=color, linewidth=0.2 if 'Total' not in label else 1)

def toggle_line_visibility(line, button):
    line.set_visible(not line.get_visible())  # ラインの表示・非表示を切り替え
    button.config(text=f"{line.get_label()} {'表示' if line.get_visible() else '非表示'}")
    canvas.draw()

# ファイル読み込み
file1_path = select_file("Dessmonitor Detail report Download File1 or today_logfile.csv")
if file1_path.endswith('.csv'):
    # CSV処理
    data0 = extract_data_csv(file1_path, 1, is_string=True)  # 1列目
    data1 = extract_data_csv(file1_path, 19)  # 18列目
    data2 = extract_data_csv(file1_path, 60)  # 42+18列目
    data4 = extract_data_csv(file1_path, 9)  # 8列目
    data5 = extract_data_csv(file1_path, 50)  # 42+8列目
else:
    if file1_path:  # 2つ目のファイルを自動検索
        directory = os.path.dirname(file1_path)  # 1つ目のファイルのディレクトリを取得
        file1_name_tail = os.path.basename(file1_path)[-10:]  # ファイル名の後半10文字を取得
        file2_path = find_matching_file(directory, file1_name_tail, file1_path)  # 修正済みの関数を使用

    if file2_path:
        print(f"2つ目のファイルが見つかりました: {file2_path}")
    else:
        print("2つ目のファイルが見つかりませんでした")   
    # Excel処理
    from openpyxl import load_workbook
    sheet1 = load_workbook(file1_path).active
    sheet2 = load_workbook(file2_path).active
    data0 = extract_data(sheet1, 1, is_string=True)  # 1列目
    data1 = extract_data(sheet1, 15)  # 14列目
    data2 = extract_data(sheet2, 15)  # 14列目
    data4 = extract_data(sheet1, 22)  # 21列目
    data5 = extract_data(sheet2, 22)  # 21列目
    data7 = extract_data(sheet1, 4)  # 3列目
    data8 = extract_data(sheet2, 4)  # 3列目
    data10= extract_data(sheet1, 5)  # 4列目
    data11= extract_data(sheet2, 5)  # 4列目
    data13= extract_data(sheet1, 9)  # 8列目
    data14= extract_data(sheet2, 9)  # 8列目
    data16= extract_data(sheet1,10)  # 9列目
    data17= extract_data(sheet2,10)  # 9列目

# データ調整
min_length = min(map(len, [data0, data1, data2, data4, data5,data7,data8,data10,data11,data13,data14,data16,data17]))                    
data0, data1, data2, data4, data5,data7,data8,data10,data11,data13,data14,data16,data17 = [d[:min_length] for d in [data0, data1, data2, data4, data5,data7,data8,data10,data11,data13,data14,data16,data17]]
data3 = [a + b for a, b in zip(data1, data2)]
data6 = [a + b for a, b in zip(data4, data5)]
data9 = [a * b for a, b in zip(data7, data10)]
data12 = [a * b for a, b in zip(data8, data11)]
data15 = [a * b for a, b in zip(data13, data16)]
data18 = [a * b for a, b in zip(data14, data17)]
data19 = [a + b for a, b in zip(data9, data12)]
data20 = [a + b for a, b in zip(data15, data18)]
# データ反転（左右逆）
data0 = data0[::-1]
data1 = data1[::-1]
data2 = data2[::-1]
data3 = data3[::-1]
data4 = data4[::-1]
data5 = data5[::-1]
data6 = data6[::-1]
data9 = data9[::-1]
data12= data12[::-1]
data19= data19[::-1]
data15= data15[::-1]
data18= data18[::-1]
data20= data20[::-1]
# グラフ描画設定
sns.set(style='darkgrid', palette='winter_r')
plt.figure(figsize=(15, 8))
plt.xticks(rotation=80)
plt.yticks(rotation=-10)

# プロット
plot_graph(data0, [data1, data2, data3, data4, data5, data6,data9,data12,data19,data15,data18,data20],
           ['PV1_power', 'PV2_power', 'PV_Total_power', 'Inv1_power', 'Inv2_power', 'INV_Total_power','Batt_power1','Batt_power2','Batt_Total_power','Grid_power1','Grid_power2','Grid_Total_power'],
           ['red', 'red', 'red', 'blue', 'blue', 'blue','green','green','green','purple','purple','purple'])

# 最大値マーク
for data, label, color,ha in [(data3, "Max_PV_power", "red","right"), (data6, "Max_INV_power", "blue","left"),(data19, "Max_Batt_power", "green","right"),(data20, "Max_Grid_power", "purple","right")]:
    max_idx = np.argmax(data)
    plt.scatter(data0[max_idx], data[max_idx], color=color, label=f"{label}")
    plt.text(data0[max_idx], data[max_idx], f"({data0[max_idx]}, {data[max_idx]}W)", fontsize=10, color=color, ha=ha)

# Batt_powerデータの最低値マーク
for data, label, color in [(data19, "Min_Batt_Total_power", "green")]:
    min_idx = np.argmin(data)  # 最低値のインデックスを取得
    plt.scatter(data0[min_idx], data[min_idx], color=color, marker='x', label=f"{label}")  # 最低値マーク
    plt.text(data0[min_idx], data[min_idx], f"({data0[min_idx]}, {data[min_idx]}W)", fontsize=10, color=color, ha="left")  # 最低値にテキスト表示

# グラフ設定
plt.title('電力量', fontsize=40, fontproperties=jp_font)
plt.xlabel('日時', fontsize=12, fontproperties=jp_font)
plt.ylabel('電力', fontsize=20, fontproperties=jp_font)
plt.legend(bbox_to_anchor=(1, 1), loc='upper left')
plt.grid(True)
plt.gca().xaxis.set_major_locator(plt.MultipleLocator(10))
plt.gca().yaxis.set_major_locator(plt.MultipleLocator(500))

# Maximize and set window title
manager = plt.get_current_fig_manager()
manager.window.state('zoomed')  # For maximizing (Windows)
manager.set_window_title("PV & Inverter Power Chart")  # Set window title
# Show the plot

plt.show()

