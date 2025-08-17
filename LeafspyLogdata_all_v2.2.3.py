import tkinter as tk
from tkinter import filedialog
import pandas as pd

import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.family'] = 'MS Gothic'
import glob
import os
import numpy as np

def skip_invalid_rows(df):
    # 2列目・3列目が欠損している行を除外
    return df.dropna(subset=[df.columns[1], df.columns[2]])

def get_last_valid_row(df):
    # 2列目・3列目が欠損していない最後の行を取得
    valid_df = skip_invalid_rows(df)
    if not valid_df.empty:
        return valid_df.iloc[-1].tolist()
    return None

def sort_and_format_datetime(df):
    # 1列目を "mm/dd/yyyy hh:mm:ss" から datetime型に変換
    df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], format='%m/%d/%Y %H:%M:%S', errors='coerce')
    # 日時で昇順ソート
    df = df.sort_values(by=df.columns[0])
    # "yyyy/mm/dd hh:mm:ss" 形式に変換
    df[df.columns[0]] = df[df.columns[0]].dt.strftime('%Y/%m/%d %H:%M:%S')
    return df

root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory(title="フォルダを選択してください")

if folder_path:
    csv_files = sorted(
        glob.glob(os.path.join(folder_path, 'log_*.csv')),
        key=lambda x: os.path.getmtime(x)
    )

    merged_df = None
    last_row_before_merge = None

    for i, file in enumerate(csv_files):
        if i == 0:
            df = pd.read_csv(file, encoding='utf-8')
            df = skip_invalid_rows(df)
            df = sort_and_format_datetime(df)
            merged_df = df
        else:
            df = pd.read_csv(file, encoding='utf-8', skiprows=1, header=None)
            df.columns = merged_df.columns
            df = skip_invalid_rows(df)
            df = sort_and_format_datetime(df)
            merged_df = pd.concat([merged_df, df], ignore_index=True)

    # 124列目（0始まりなので123）の値が0の行を削除
    merged_df = merged_df[merged_df.iloc[:, 123] != 0]

    # 結合前の最後の有効行
    if not merged_df.empty:
        last_row_before_merge = merged_df.iloc[-1].tolist()

    out_path = os.path.join(folder_path, 'alllog.csv')

    # alllog.csvの最後の有効行を取得
    if os.path.exists(out_path):
        alllog_df = pd.read_csv(out_path, encoding='utf-8')
        alllog_df = skip_invalid_rows(alllog_df)
        last_row_alllog = get_last_valid_row(alllog_df)
    else:
        last_row_alllog = None

    # 最後の行が同じ場合は結合せずalllog.csvのみでグラフ処理
    if last_row_alllog == last_row_before_merge and last_row_alllog is not None:
        print('alllog.csvの最後の行と結合前の最後の行が同じため、結合せずグラフ処理します。')
        file_path = out_path
    else:
        merged_df.to_csv(out_path, index=False, encoding='utf-8')
        print(f'{out_path} に結合完了')
        file_path = out_path

    if file_path:
        # ヘッダー取得
        with open(file_path, encoding='utf-8') as f:
            header = f.readline().strip().split(',')

        # データ部のみ読み込み
        df = pd.read_csv(file_path, header=None, skiprows=1)
        # 日付の無い行を削除
        df = df[df[0].notnull() & (df[0] != '')]
        # 1列目を日付型に変換（"yyyy/mm/dd hh:mm:ss"形式）
        df[0] = pd.to_datetime(df[0], format='%Y/%m/%d %H:%M:%S', errors='coerce')
        # 日付型変換できなかった行も削除
        df = df[df[0].notnull()]
        # 日付昇順でソート
        df = df.sort_values(by=0).reset_index(drop=True)
        # 日付を文字列に戻す
        df[0] = df[0].dt.strftime('%Y/%m/%d %H:%M:%S')
        # ヘッダー＋データで再保存
        with open(file_path, 'w', encoding='utf-8', newline='') as f:
            f.write(','.join(header) + '\n')
            df.to_csv(f, index=False, header=False)

        # ---以下グラフ処理はそのまま---
        x = df[0].astype(str)
        soc = df[6] / 10000  # SOCは10000で割る
        hx = df[121]  # HX
        soh = df[131] # SOH
        gids = df[5]  # Gids
        odo = df[123]  # ODO
        gids_a = (gids / 750) *100
        gids_c = ((gids * 0.08)  / (62.76384 * (soh/100)))*100 # Gidsx80
        gids_b = gids_c/gids_a
        gids_h = ((gids * 0.08) / (soc/100)  / 0.6276384)+2.2# Gids
        
        voltage = df[8]
        current = df[9]
        cell_max_voltage = df[10] / 1000  # セル最大電圧は1000で割る
        cell_min_voltage = df[11] / 1000  # セル最小電圧は1000で割る
        cell_temp1 = df[16]
        cell_temp2 = df[18]
        cell_temp3 = df[22]
        ambient_temp = df[130]  # 環境温度

        ambient_temp = ambient_temp.combine(cell_temp1, lambda amb, cell: (amb - 32) * 5 / 9 if amb >= cell + 20 else amb)

        # ---グラフ描画処理は既存のまま---
        fig, axs = plt.subplots(2, 2, figsize=(16, 10))
        fig.canvas.manager.set_window_title(file_path)  # ウインドウタイトルをファイル名に変更

        # 左上: SOC, Hx, SOH
        ax1 = axs[0, 0]
        l_soc, = ax1.plot(x, soc, label='SOC', color="#244cff", linewidth=0.8, picker=True)
        l_hx, = ax1.plot(x, hx, label='Hx', color="#ff5e00", linewidth=0.8, picker=True)
        l_soh, = ax1.plot(x, soh, label='SOH', color="#ff0080", linewidth=0.8, picker=True)
        ax1.set_xlabel('日付')
        ax1.set_ylabel('Hx / SOH / SOC (%)')
        ax1.set_xticks(range(0, len(x), 800))
        ax1.set_xticklabels(x[::800], rotation=45)
        ax1.legend(loc='lower left', bbox_to_anchor=(0, 0), ncol=1)
        ax1.set_title('SOC / Hx / SOH')
        ax1.grid(axis='y')

        # 左下: 各温度
        ax2 = axs[1, 0]
        l_temp1, = ax2.plot(x, cell_temp1, label='Cell Temp 1', color='#66bb6a', linewidth=0.5, picker=True)
        l_temp2, = ax2.plot(x, cell_temp2, label='Cell Temp 2', color='#388e3c', linewidth=0.5, picker=True)
        l_temp3, = ax2.plot(x, cell_temp3, label='Cell Temp 3', color='#1b5e20', linewidth=0.5, picker=True)
        l_ambient, = ax2.plot(x, ambient_temp, label='Ambient Temp', color="#eb87ff", linewidth=0.3, picker=True)
        ax2.set_xlabel('日付')
        ax2.set_ylabel('Temp (℃)')
        ax2.set_xticks(range(0, len(x), 800))
        ax2.set_xticklabels(x[::800], rotation=45)
        ax2.legend(loc='lower left', bbox_to_anchor=(0, 0), ncol=1)
        ax2.set_title('各温度')
        ax2.grid(axis='y')

        # 右上: セル最大・最小電圧
        ax3 = axs[0, 1]
        l_max, = ax3.plot(x, cell_max_voltage, color="#8301fd", label='Cell Max Voltage', linewidth=0.5, picker=True)
        l_min, = ax3.plot(x, cell_min_voltage, color="#65e8ff", linestyle='--', label='Cell Min Voltage', linewidth=0.5, picker=True)
        ax3.set_xlabel('日付')
        ax3.set_ylabel('Cell Voltage(V)')
        ax3.set_xticks(range(0, len(x), 800))
        ax3.set_xticklabels(x[::800], rotation=45)
        ax3.set_ylim(3.3, 4.4)
        ax3.legend(loc='lower left', bbox_to_anchor=(0, 0), ncol=1)
        ax3.set_title('セル電圧')
        ax3.grid(axis='y')

        # 右下: 走行距離
        ax4 = axs[1, 1]
        l_odo, = ax4.plot(x, odo, color="#FFBB00", label='Odo', linewidth=0.8, picker=True)
        ax4.set_xlabel('日付')
        ax4.set_ylabel('Odo(km)')
        ax4.set_xticks(range(0, len(x), 800))
        ax4.set_xticklabels(x[::800], rotation=45)
        ax4.set_ylim(bottom=10000)
        ax4.legend(loc='lower left', bbox_to_anchor=(0, 0), ncol=1)
        ax4.set_title('走行距離')
        ax4.grid(axis='y')

        plt.tight_layout()
        plt.suptitle('LeafSpy Logdata', fontname='MS Gothic', fontsize=16, y=1.02)
        plt.subplots_adjust(top=0.92)

        # グラフウインドウを最大化
        mng = plt.get_current_fig_manager()
        mng.window.state('zoomed')  # Windows環境で最大化

        # クリックイベントで別ウインドウ表示
        def on_pick(event):
            artist = event.artist
            label = artist.get_label()
            xdata = np.array(artist.get_xdata())
            ydata = np.array(artist.get_ydata())

            max_idx = np.argmax(ydata)
            min_idx = np.argmin(ydata)

            fig2, ax2 = plt.subplots(figsize=(8, 4))
            ax2.plot(xdata, ydata, color=artist.get_color(), label=label)
            ax2.plot(xdata[max_idx], ydata[max_idx], 'ro', label='最大値')
            ax2.plot(xdata[min_idx], ydata[min_idx], 'bo', label='最小値')
            ax2.set_title(label)
            ax2.grid(axis='y')

            # x軸ラベルを最初・最後のみ表示、角度付き
            xticks = [xdata[0], xdata[-1]]
            xticklabels = [str(xdata[0]), str(xdata[-1])]
            ax2.set_xticks(xticks)
            ax2.set_xticklabels(xticklabels, rotation=0)

            # 最大値・最小値の横にx軸ラベルと値を表示
            ax2.annotate(f'x={xdata[max_idx]}\ny={ydata[max_idx]}',
                         xy=(xdata[max_idx], ydata[max_idx]),
                         xytext=(10, 10),
                         textcoords='offset points',
                         color='red',
                         fontsize=10,
                         arrowprops=dict(arrowstyle='->', color='red'))
            ax2.annotate(f'x={xdata[min_idx]}\ny={ydata[min_idx]}',
                         xy=(xdata[min_idx], ydata[min_idx]),
                         xytext=(10, -20),
                         textcoords='offset points',
                         color='blue',
                         fontsize=10,
                         arrowprops=dict(arrowstyle='->', color='blue'))

            ax2.legend(loc='best')
            fig2.canvas.manager.set_window_title(label)
            fig2.tight_layout()  # はみ出し防止
            fig2.canvas.mpl_connect('button_press_event', lambda e: plt.close(fig2))
            fig2.show()

        fig.canvas.mpl_connect('pick_event', on_pick)

        plt.show()
else:
    print('フォルダが選択されませんでした。')