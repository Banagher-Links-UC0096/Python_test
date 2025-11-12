dessmonitor_data_cache_global = None

# -*- coding: utf-8 -*-
"""
CSVフォーマット想定:
1行目: ヘッダー
2行目以降: 1列目=日付(yyyymmdd), 8列目=日合計(kWh), 10列目=0時-1時, 11列目=1時-2時, ...

スクリプトは月毎に各時間帯(デイ/ホーム/ナイト)の合計を積み上げ棒グラフで出力します。
使い方: python smartmeter_monthly_stack.py path/to/file.csv

出力: ./monthly_stack.png
"""

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

import sys
import glob, os, re
from pathlib import Path
import pandas as pd
import numpy as np

import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.ticker as ticker
from matplotlib.patches import Patch

import tkinter.ttk as ttk
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# 祝日判定用ライブラリ
import jpholiday
import datetime
import time
from collections import defaultdict


# 日本語フォントの設定
plt.rcParams['font.family'] = 'MS Gothic'  # Windows の場合

# ---- グローバル変数 ----
LAST_FIG_OPEN = 0.0

# ---- 共通の定義 ----
# ---- 料金関連の定義 ----
# 時間帯単価（円/kWh）
PRICE_PERIODS = [
    # 適用期間 ~2023年3月以前
    (None, pd.Timestamp('2023-03-31'), {
         'day': 33.97, 
         'home': 25.91, 
         'night': 15.89
         }),
    # 適用期間 2023年4月～2024年3月
    (pd.Timestamp('2023-04-01'), pd.Timestamp('2024-03-31'), {
        'day': 34.21, 
        'home': 26.15, 
        'night': 16.22
        }),
    # 適用期間 2024年4月～2025年4月
    (pd.Timestamp('2025-05-01'), None, {
        'day': 34.06, 
        'home': 26.00, 
        'night': 16.11}),
]

# 再エネ賦課金単価: 毎年5月に更新 -> その年の5月から翌年4月まで有効とする
RENEWABLE_BY_YEAR = {
    2022: 3.45,
    2023: 1.40,
    2024: 3.49,
    2025: 3.98,
}

# 燃料費調整単価: 月次 (円/kWh)
FUEL_ADJ = {
    "2022-12":11.04,

    "2023-01":12.03,
    "2023-02": 5.51,
    "2023-03": 4.28,
    "2023-04": 2.93,
    "2023-05": 1.95,
    "2023-06": 0.60,
    "2023-07":-0.94,
    "2023-08":-2.57,
    "2023-09":-3.74,
    "2023-10":-0.73,
    "2023-11":-0.96,
    "2023-12":-1.10,

    "2024-01":-1.01,
    "2024-02":-0.82,
    "2024-03":-0.31,
    "2024-04":-0.10,
    "2024-05": 0.04,
    "2024-06": 1.51,
    "2024-07": 2.84,
    "2024-08": 2.54,
    "2024-09":-1.55,
    "2024-10":-1.25,
    "2024-11": 0.30,
    "2024-12": 2.59,

    "2025-01": 2.33,
    "2025-02":-0.15,
    "2025-03": 0.06,
    "2025-04": 1.64,
    "2025-05": 2.84,
    "2025-06": 2.63,
    "2025-07": 1.98,
    "2025-08":-0.49,
    "2025-09":-1.21,
    "2025-10":-1.02,
    "2025-11": 0.93,
    "2025-12": 0.86,

    "2026-01": 0.00,
}

# 時間帯定義
TIME_BANDS = {
    'day': {
        'hours': list(range(9, 17)),  # デイタイム9時～17時
        'color': '#FFEB99',
        'color_rgb': (1.0, 0.92, 0.6),
        'label': 'デイタイム(平日:9-17時)'
    },
    'home': {
        'hours': list(range(7, 9)) + list(range(17, 23)),  # ホームタイム7-9時,17-23時
        'color': '#99E699',
        'color_rgb': (0.6, 0.9, 0.6),
        'label': 'ホームタイム(平日:7-9,17-23時\n※土日祝は9-17も含む)'
    },
    'night': {
        'hours': list(range(23, 24)) + list(range(0, 7)),  # ナイトタイム23時-7時
        'color': '#99CCFF',
        'color_rgb': (0.6, 0.8, 1.0),
        'label': 'ナイトタイム(23,0-7時)'
    }
}

# 凡例や軸の共通設定
PLOT_SETTINGS = {
    'cost_color': 'red', # 電気料金プロット色
    'cost_marker': 'o',
    'cost_linewidth': 0.5, # 電気料金プロット線幅
    'grid_color': 'gray', # グリッド線色
    'grid_alpha': 0.7, # グリッド線透明度
    'grid_linestyle': '--',
    'annotation_bgcolor': 'yellow', # アノテーション背景色
    'annotation_alpha': 0.9, # アノテーション背景透明度
    'summary_bgcolor': 'white', # 右側凡例背景色
    'summary_alpha': 0.85 # 右側凡例背景透明度
}

# 列の順序定義
COLUMN_ORDER = ['night', 'home', 'day']  # 積み上げの順序（下から）
DISPLAY_ORDER = ['day', 'home', 'night']  # 表示の順序（凡例など）

#---- 料金計算関連関数 ----
# 時間帯判定関数
def hour_to_band(hour: int, is_holiday_flag: bool) -> str:
    """時間 (0-23) を受け取り、休日フラグに応じて 'day'/'home'/'night' を返す。"""
    try:
        h = int(hour) % 24
    except Exception:
        return 'night'
    if is_holiday_flag:
        # 休日は 'home' の時間帯が拡張される
        if h in TIME_BANDS['home']['hours'] or h in TIME_BANDS['day']['hours']:
            return 'home'
        else:
            return 'night'
    # 平日
    for band, info in TIME_BANDS.items():
        if h in info['hours']:
            return band
    return 'night'

# 土日祝日判定関数
def is_holiday(date):
    """土日祝日判定"""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    # 土日判定
    if date.weekday() >= 5:  # 5=土曜日, 6=日曜日
        return True
    # 祝日判定
    return jpholiday.is_holiday(date)

# 指定日の単価を返す関数
def get_unit_prices_for_date(date):
    """指定日のデイ/ホーム/ナイト単価を返す（円/kWh）。"""
    if pd.isna(date):
        return PRICE_PERIODS[-1][2]
    if isinstance(date, (pd.Timestamp, datetime.date)):
        dt = pd.Timestamp(date)
    else:
        dt = pd.to_datetime(date)
    for start, end, prices in PRICE_PERIODS:
        if (start is None or dt >= start) and (end is None or dt <= end):
            return prices
    return PRICE_PERIODS[-1][2]

# 再エネ賦課金単価を返す関数
def get_renewable_unit_for_date(date):
    """再エネ賦課金単価を返す。毎年5月に切替: 例) 2023年5月～2024年4月は RENEWABLE_BY_YEAR[2023]"""
    if isinstance(date, (pd.Timestamp, datetime.date)):
        dt = pd.Timestamp(date)
    else:
        dt = pd.to_datetime(date)
    year = dt.year
    key_year = year if dt.month >= 5 else year - 1
    return RENEWABLE_BY_YEAR.get(key_year, 0.0)

# 燃料費調整単価を返す関数
def get_fuel_adj_for_date(date):
    if isinstance(date, (pd.Timestamp, datetime.date)):
        dt = pd.Timestamp(date)
    else:
        dt = pd.to_datetime(date)
    key = f"{dt.year:04d}-{dt.month:02d}"
    return FUEL_ADJ.get(key, 0.0)

# 日単位の電気料金詳細計算関数
def compute_cost_breakdown(date, day_kwh, home_kwh, night_kwh):
    """日単位のエネルギー構成から時間帯別の金額、再エネ・燃料調整額、合計を詳細に返す。

    戻り値: dict {
        'prices': {..}, 'renew': float, 'fuel': float,
        'base_costs': {'day':..,'home':..,'night':..}, 'base_total':..,
        'renew_amount':.., 'fuel_amount':.., 'total_kwh':.., 'total_cost':..
    }
    """
    prices = get_unit_prices_for_date(date) # 時間帯単価取得
    renew = get_renewable_unit_for_date(date) # 再エネ賦課金単価取得
    fuel = get_fuel_adj_for_date(date) # 燃料費調整単価取得
    day_kwh = float(day_kwh or 0.0) # デイタイム時間帯ごとのkWhをfloat化
    home_kwh = float(home_kwh or 0.0) # ホームタイム時間帯ごとのkWhをfloat化
    night_kwh = float(night_kwh or 0.0) # ナイトタイム時間帯ごとのkWhをfloat化
    base_costs = {
        'day': day_kwh * prices.get('day', 0.0), # デイタイム金額
        'home': home_kwh * prices.get('home', 0.0), # ホームタイム金額
        'night': night_kwh * prices.get('night', 0.0) # ナイトタイム金額
    }
    base_total = sum(base_costs.values()) # 基本料金合計
    total_kwh = day_kwh + home_kwh + night_kwh # 総使用電力量
    renew_amount = total_kwh * renew # 再エネ賦課金額
    fuel_amount = total_kwh * fuel # 燃料費調整額
    total_cost = base_total + renew_amount + fuel_amount # 総合計金額
    return {
        'prices': prices,
        'renew': renew,
        'fuel': fuel,
        'base_costs': base_costs,
        'base_total': base_total,
        'renew_amount': renew_amount,
        'fuel_amount': fuel_amount,
        'total_kwh': total_kwh,
        'total_cost': total_cost
    }

# --- Dessmonitorデータ一括キャッシュ関数 ---
def build_dessmonitor_data_cache(dessmonitor_folder):
    """指定フォルダ配下の全energy-storage-container-*.xlsxを一括で読み込み、
    {(date, hour): {'energy_sum': [値リスト], 'charge_sum': [値リスト]}} のdictを返す。"""
    
    cache = defaultdict(lambda: {'PV_sum': [], 'energy_sum': [], 'charge_sum': []}) # 全体キャッシュ
    pattern = os.path.join(dessmonitor_folder, '**/energy-storage-container-*.xlsx') # ファイルパターン
    files = glob.glob(pattern, recursive=True) # ファイル一覧取得
    for f in files: # 各ファイルを処理
        try:
            df = pd.read_excel(f) # Excel読み込み
            df_sorted = df.iloc[::-1].copy() # 日時降順にソートし、明示的にコピー
            df_sorted['hour'] = df_sorted.iloc[:,0].apply(lambda x: int(str(x)[11:13]) if isinstance(x,str) and len(x)>=13 else None) # 時間抽出
            df_sorted['date'] = df_sorted.iloc[:,0].apply(lambda x: str(x)[:10] if isinstance(x,str) and len(x)>=10 else None) # 日付抽出
            for (date, hour), group in df_sorted.groupby(['date', 'hour']): # 日付・時間毎にグループ化
                if date is None or hour is None:
                    continue
                # 発電量
                vals_gen = group.iloc[:,29].values if group.shape[1]>29 else []
                if len(vals_gen) >= 2:
                    if vals_gen[0] > vals_gen[-1]:
                        gen = 0.0
                    else:
                        gen = float(vals_gen[-1]) - float(vals_gen[0])
                else:
                    gen = 0.0
                #gen_sum += gen
                # 蓄電
                vals = group.iloc[:,30].values if group.shape[1]>30 else [] # 31列目: 蓄電量
                if len(vals) >= 2: # 蓄電量計算: 最初と最後の差分を合算
                    if vals[0] > vals[-1]: # 値が減少している場合
                        energy = 0.0 # 減少分は0扱い
                    else:
                        energy = float(vals[-1]) - float(vals[0]) # 差分を加算
                else:
                    energy = 0.0 # 蓄電量データ不足時は0扱い
                # 商用充電量
                vals_9 = group.iloc[:,8].values if group.shape[1]>8 else [] # 9列目: Grid電圧
                vals_10 = group.iloc[:,9].values if group.shape[1]>9 else [] # 10列目: Grid電流
                vals_22 = group.iloc[:,21].values if group.shape[1]>21 else [] # 22列目: Inverter出力電力
                charge_sum = 0.0 # 商用充電量合計
                if len(vals_9) >= 2 and len(vals_10) >= 2 and len(vals_22) >= 2: # 必要列が存在するか
                    for i in range(len(group)-1, 0, -1): # 各行を処理
                        try:
                            v9 = float(vals_9[i]) # 9列目
                            v10 = float(vals_10[i]) # 10列目
                            v22 = float(vals_22[i]) # 22列目
                            charge = (v9 * v10 -v22)* (1/12000) # kWh換算(5分毎)
                            if charge < 0: # 負の値は0扱い
                                charge = 0.0 # 負の差分は0扱い
                            charge_sum += charge # 加算
                        except Exception:
                            continue
                cache[(date, hour)]['PV_sum'].append(gen) # 蓄電量リスト初期化
                cache[(date, hour)]['energy_sum'].append(energy) # 蓄電量追加
                cache[(date, hour)]['charge_sum'].append(charge_sum) # 商用充電量追加
        except Exception:
            continue
    return cache

# ---- 共通ユーティリティ関数 ----

# 時間ごとの単価リスト作成関数
def build_hourly_unit(prices: dict, renew: float, fuel: float, is_holiday_flag: bool) -> list:
    """単日 (24要素) の時間ごとの単価リストを返す (円/kWh)。"""
    return [prices[hour_to_band(h, is_holiday_flag)] + renew + fuel for h in range(24)]

# 時間帯ごとの合計取得関数
def band_sums_from_values(values, is_holiday_flag: bool):
    """values: iterable of 24 hourly kWh -> bandごとの合計辞書を返す."""
    sums = {'day': 0.0, 'home': 0.0, 'night': 0.0}
    for h, v in enumerate(values):
        band = hour_to_band(h, is_holiday_flag)
        try:
            sums[band] += float(v)
        except Exception:
            try:
                sums[band] += 0.0
            except Exception:
                pass
    return sums

# アノテーション追加関数
def annotate_in_axes(ax, fig, x, y, msg, bgcolor=None, alpha=None):
    """Axes内に収まるようにアノテーションを追加して返す。"""
    if bgcolor is None:
        bgcolor = PLOT_SETTINGS.get('annotation_bgcolor', 'yellow')
    if alpha is None:
        alpha = PLOT_SETTINGS.get('annotation_alpha', 0.9)
    try:
        ylim = ax.get_ylim()
        y = min(y + (ylim[1] - ylim[0]) * 0.03, ylim[1] * 0.95)
    except Exception:
        pass
    ann = ax.annotate(msg, xy=(x, y), xytext=(0, 0), xycoords='data', textcoords='offset points',
                      ha='center', va='bottom', fontsize=10,
                      bbox=dict(boxstyle='round,pad=0.3', fc=bgcolor, alpha=alpha), zorder=1000)
    ann.set_clip_on(True)
    try:
        fig.canvas.draw_idle()
    except Exception:
        pass
    return ann

# CSV読み込み関数
def load_csv(path: Path) -> pd.DataFrame:
    # try common encodings if utf-8 fails (cp932/shift_jis on Windows)
    encodings = ['utf-8', 'cp932', 'shift_jis', 'latin1']
    last_err = None
    for enc in encodings:
        try:
            df = pd.read_csv(path, header=0, dtype=str, encoding=enc)
            print(f"Loaded CSV with encoding: {enc}")
            return df
        except Exception as e:
            last_err = e
    # if all fail, raise the last exception
    raise last_err

# CSV解析・集計関数
def parse_and_aggregate(df: pd.DataFrame) -> pd.DataFrame:
    # ...existing code...
    if df.shape[1] < 33:
        raise ValueError('CSVに十分な列がありません。24時間分の列が必要です。')

    date_col = df.columns[0]
    df['date'] = pd.to_datetime(df[date_col], format='%Y%m%d', errors='coerce')
    if df['date'].isna().any():
        print('警告: 日付のパースに失敗した行があります。')

    start_hour_col_idx = 9
    hourly_cols = df.columns[start_hour_col_idx:start_hour_col_idx+24]
    df_hourly = df[hourly_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

    df['day'] = 0.0
    df['home'] = 0.0
    df['night'] = 0.0

    for idx in df.index:
        date_val = df.at[idx, 'date']
        try:
            is_hol = is_holiday(date_val)
        except Exception:
            is_hol = False

        band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0}
        for h in range(24):
            band = hour_to_band(h, is_hol)
            val = df_hourly.at[idx, hourly_cols[h]]
            band_sums[band] += float(val)
        df.at[idx, 'day'] = band_sums['day']
        df.at[idx, 'home'] = band_sums['home']
        df.at[idx, 'night'] = band_sums['night']

    df['year_month'] = df['date'].dt.to_period('M')
    monthly = df.groupby('year_month')[['day', 'home', 'night']].sum()
    monthly = monthly.sort_index()
    monthly.index = monthly.index.astype(str)
    return monthly, df

# 時間毎のグラフを表示する関数
def plot_hourly(df: pd.DataFrame, year: int, month: int, day: int, file_path: Path = None, dessmonitor_folder: str = None):
    global dessmonitor_data_cache_global # Dessmonitorデータキャッシュ
    if dessmonitor_data_cache_global is None: # 初回読み込み
        dessmonitor_data_cache_global = build_dessmonitor_data_cache(dessmonitor_folder) # Dessmonitorデータキャッシュを構築
    dessmonitor_data_dict = dessmonitor_data_cache_global # ローカル参照
    
    # --- Dessmonitorデータを時間毎グラフに反映する関数 ---
    def load_dessmonitor_data_hourly(
        values,
        target_date,
        colors,
        prices,
        renew,
        fuel,
        is_hol,
        file_path=None
    ):
        
        # Dessmonitorデータキャッシュを使用
        global dessmonitor_data_cache_global
        if dessmonitor_data_cache_global is None:
            messagebox.showerror('エラー', 'Dessmonitorデータキャッシュがありません')
            return
        date_str = target_date.strftime('%Y-%m-%d')
        pv_values = np.zeros(24)
        dess_values = np.zeros(24)
        commercial_values = np.zeros(24)
        for h in range(24):
            key = (date_str, h)
            vals_pv = dessmonitor_data_cache_global[key]['PV_sum'] if key in dessmonitor_data_cache_global else []
            pv_values[h] += sum(vals_pv)
            vals = dessmonitor_data_cache_global[key]['energy_sum'] if key in dessmonitor_data_cache_global else []
            dess_values[h] += sum(vals)
            vals_c = dessmonitor_data_cache_global[key]['charge_sum'] if key in dessmonitor_data_cache_global else []
            commercial_values[h] += sum(vals_c)
        # プラス側：電力会社データそのまま
        new_values = values
        # マイナス側：Dessmonitorデータそのまま負値で
        pv_values_plot = -pv_values
        dess_values_plot = -dess_values
        # 商用充電はプラス側で幅半分、中心揃え
        bar_width = 0.8
        commercial_bar_width = bar_width / 2
        # --- グラフ表示（plot_hourlyと同じ構成） ---
        fig2, ax2 = plt.subplots(figsize=(10,6))
        win2 = tk.Toplevel(win)
        win2.wm_title(f'{date_str} 時間帯別電力使用量 (Dessmonitor反映)')
        canvas2 = FigureCanvasTkAgg(fig2, master=win2)
        canvas2.draw()
        widget2 = canvas2.get_tk_widget()
        widget2.pack(side='top', fill='both', expand=1)
        
         # --- グラフにデータをプロット ---
        # 棒グラフ描画
        bars2 = ax2.bar(range(24), new_values, color=colors, label='電力会社データ', width=bar_width, align='center')
        bars2_dess = ax2.bar(range(24), dess_values_plot, color='deepskyblue', alpha=0.6, label='蓄電池出力', width=bar_width, align='center')
        # 商用充電はプラス側で幅半分、中心揃え
        bars2_commercial = ax2.bar([x for x in range(24)], commercial_values, color='deepskyblue', alpha=0.6, label='商用充電', width=commercial_bar_width, align='center')
        ax2.plot(range(24), pv_values_plot, color='orange', marker='', linewidth=1.2, label='発電量', zorder=10)
        
        ax2.set_ylabel('消費電力量 (kWh)')
        ax2.set_title(f'{date_str} 時間帯別電力使用量 (Dessmonitor反映)')
        ax2.set_xticks(range(24))
        ax2.set_xticklabels([str(h) for h in range(24)], rotation=0)
        
        legend_patches = [Patch(color=TIME_BANDS[band]['color'], label=TIME_BANDS[band]['label']) for band in DISPLAY_ORDER]
        legend_patches += [Patch(color='deepskyblue', label='蓄電池出力')]
        legend_patches += [Patch(color='deepskyblue', label='商用充電')]
        legend_patches += [Patch(color='orange', label='発電量')]
        ax2.legend(handles=legend_patches, bbox_to_anchor=(1.05, 1.1), loc='upper left')
        ax2.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
        
        ax2.yaxis.set_major_locator(ticker.MultipleLocator(0.5))
        # 右軸に料金をプロット
        hourly_unit = build_hourly_unit(prices, renew, fuel, is_hol)
        hourly_costs = new_values * np.array(hourly_unit)
        ax2_2 = ax2.twinx()
        ax2_2.plot(range(24), hourly_costs, color='red', marker='o', linewidth=0.5, label='電気料金 (円)')
        # 節約金額 = 蓄電池分 - 商用充電分
        saving_costs = (-dess_values + commercial_values) * np.array(hourly_unit)
        ax2_2.plot(range(24), saving_costs, color='blue', marker='o', linewidth=0.5, label='節約料金 (円)')
        ax2_2.set_ylabel('電気料金 (円)')
        ax2_2.yaxis.grid(False)
        ax2_2.legend(loc='upper right')
        # --- 左右y軸の0ラインを揃え、左y軸を右y軸の1/10に設定 ---
        left_ylim = ax2.get_ylim()
        # 右y軸の範囲を取得
        min_ylim = min(0, left_ylim[0])
        max_ylim = left_ylim[1]+0.2  # 少し余裕を持たせる
        ax2.set_ylim(min_ylim, max_ylim)
        ax2_2.set_ylim(min_ylim*40, max_ylim*40)# 倍率
        # --- 右側の凡例下に料金・集計表示を追加 ---
        # 詳細情報テキスト
        band_sums = band_sums_from_values(new_values, is_hol)
        bd = compute_cost_breakdown(target_date, band_sums['day'], band_sums['home'], band_sums['night'])
        base_costs = bd['base_costs']
        base_total = bd['base_total']
        renew_amount = bd['renew_amount']
        fuel_amount = bd['fuel_amount']
        total_cost = bd['total_cost']
        prices = bd['prices']
        renew = bd['renew']
        fuel = bd['fuel']
        total_kwh = bd.get('total_kwh', float(np.nansum(new_values)))
        dess_band_sums = band_sums_from_values(dess_values, is_hol)
        pv_band_sums = band_sums_from_values(pv_values, is_hol)
        dess_bd = compute_cost_breakdown(target_date, dess_band_sums['day'], dess_band_sums['home'], dess_band_sums['night'])
        dess_base_costs = dess_bd['base_costs']
        dess_base_total = dess_bd['base_total']
        dess_renew_amount = dess_bd['renew_amount']
        dess_fuel_amount = dess_bd['fuel_amount']
        dess_total_cost = dess_bd['total_cost']
        dess_total_kwh = dess_bd.get('total_kwh', float(np.nansum(dess_values)))
        charg_sums = band_sums_from_values(commercial_values, is_hol)
        charge_cost = charg_sums['day'] * prices['day'] + charg_sums['home'] * prices['home'] + charg_sums['night'] * prices['night'] + (charg_sums['day'] + charg_sums['home'] + charg_sums['night']) * (renew + fuel)
        saving_cost = dess_total_cost- charge_cost
        buy_total = base_costs['day'] + base_costs['home'] + base_costs['night'] + renew_amount + fuel_amount 
        total_cost2 = buy_total + saving_cost
        
        # 単価表示
        unit_text = (f"適用単価:\n"
                     f"デイタイム単価　 {prices['day']:.2f} 円/kWh\n"
                     f"ホームタイム単価 {prices['home']:.2f} 円/kWh\n"
                     f"ナイトタイム単価 {prices['night']:.2f} 円/kWh\n"
                     f"再エネ賦課金単価 {renew:.2f} 円/kWh\n"
                     f"燃料費調整単価　 {fuel:.2f} 円/kWh")
        
        #  使用電力量合計
        band_lines = (f"使用電力量:\n"
                      f"デイタイム電力　 {band_sums['day']:.2f} kWh\n"
                      f"ホームタイム電力 {band_sums['home']:.2f} kWh\n"
                      f"ナイトタイム電力 {band_sums['night']:.2f} kWh\n"
                      f"買電電力量　　　 {total_kwh:.2f} kWh\n"
                      f"蓄電電力量　　　 {dess_total_kwh:.2f} kWh\n"
                      f"商用充電電力量　 {charg_sums['day']+charg_sums['home']+charg_sums['night']:.2f} kWh\n"
                      f"発電量電力量　　 {pv_band_sums['day']+pv_band_sums['home']+pv_band_sums['night']:.2f} kWh")
        
        # 集計金額（使用電力量合計に対する金額）
        totals = (f"集計金額:\n"
                  f"デイタイム金額　 {base_costs['day']:.0f} 円\n"
                  f"ホームタイム金額 {base_costs['home']:.0f} 円\n"
                  f"ナイトタイム金額 {base_costs['night']:.0f} 円\n"
                  f"再エネ賦課金額　 {renew_amount:.2f} 円\n"
                  f"燃料調整費金額　 {fuel_amount:.2f} 円\n"
                  f"買電金額　　　　 {buy_total:.0f} 円\n"
                  f"蓄電池節約金額　 {dess_total_cost:.0f} 円\n"
                  f"商用充電金額　　 {charge_cost:.0f} 円\n"
                  f"総合節約節約金額 {saving_cost:.0f} 円\n"
                  f"太陽光発電無金額 {total_cost2:.0f} 円")
        
        # 右側凡例下に表示
        summary_text = unit_text + "\n\n" + band_lines + "\n\n" + totals
        ax2.text(1.15, 0.75, summary_text, transform=ax2.transAxes,ha='left', va='top', fontsize=9,
            bbox=dict(boxstyle='round,pad=0.5', fc='white', ec='black', alpha=0.85))
        plt.tight_layout()
        
       
        # マウスホバー注釈（棒グラフ上で各時間帯の電力量表示）
        annotation = None
        fade_job = None
        
        # アノテーション削除関数
        def remove_annotation():
            nonlocal annotation, fade_job
            if annotation is not None:
                try:
                    annotation.remove()
                except Exception:
                    pass
                annotation = None
            if fade_job is not None and widget2 is not None:
                try:
                    widget2.after_cancel(fade_job)
                except Exception:
                    pass
                fade_job = None
            try:
                fig2.canvas.draw_idle()
            except Exception:
                pass
        
        # グラフ上のマウス移動イベントハンドラ
        def on_motion(event):
            nonlocal annotation
            if getattr(event, 'inaxes', None) is None or not hasattr(event.inaxes, 'containers'):
                remove_annotation()
                return
            # 棒グラフ（bars2, bars2_dess）両方に対応
            for cont in ax2.containers:
                for rect in cont:
                    try:
                        contains = rect.contains(event)[0]
                    except Exception:
                        contains = False
                    if contains:
                        xcenter = rect.get_x() + rect.get_width() / 2.0
                        hour = int(xcenter + 0.5)
                        value = rect.get_height()
                        # 棒の色で判別
                        if rect.get_facecolor()[:3] == (0.0, 0.7490196078431373, 1.0):  # deepskyblue
                            label = '蓄電使用量'
                            value = -value  # 負値を正値に変換
                        else:
                            label = '買電使用量'
                        msg = f'{hour}時-{(hour+1)%24}時\n{label}: {value:.2f}kWh'
                        if annotation is not None:
                            try:
                                annotation.remove()
                            except Exception:
                                pass
                        x = rect.get_x() + rect.get_width() / 2
                        y = rect.get_y() + rect.get_height()
                        annotation = annotate_in_axes(ax2, fig2, x, y, msg)
                        return
            remove_annotation()
        
        # マウスがグラフ外に出たときの処理
        def on_leave(event):
            remove_annotation()
        
        # 画像保存の右クリックメニュー追加
        def save_via_dialog2():
            try:
                title2 = getattr(ax2, 'get_title', lambda: "graph")()
                fname2 = filedialog.asksaveasfilename(parent=win2, defaultextension='.png', filetypes=[('PNG image','*.png')], initialfile=title2 if title2 else "graph")
                if fname2:
                    fig2.savefig(fname2, bbox_inches='tight')
                    print(f'Saved figure to {fname2}')
            except Exception as e:
                print('Save failed:', e)
        
        # コンテキストメニュー表示関数
        def show_context_menu2(event):
            try:
                menu2 = tk.Menu(widget2, tearoff=0)
                menu2.add_command(label='グラフ保存', command=save_via_dialog2)
                try:
                    menu2.tk_popup(widget2.winfo_pointerx(), widget2.winfo_pointery())
                finally:
                    menu2.grab_release()
            except Exception:
                pass
        
        # イベント接続
        try:
            canvas2 = getattr(fig2, 'canvas', None)
            if canvas2:
                canvas2.mpl_connect('motion_notify_event', on_motion)
                canvas2.mpl_connect('figure_leave_event', on_leave)
                widget2.bind('<Button-3>', show_context_menu2)
        except Exception:
            pass
        
        win2.mainloop()
    
    # 指定日のデータを抽出
    target_date = pd.Timestamp(year=year, month=month, day=day)
    row = df[df['date'] == target_date]
    if row.empty:
        print(f"{year}-{month}-{day} のデータがありません")
        return
    # 0時～23時のカラム名
    start_hour_col_idx = 9
    hourly_cols = df.columns[start_hour_col_idx:start_hour_col_idx+24]
    values = row.iloc[0][hourly_cols].astype(float).values
    # 時間帯ごとに色分け
    colors = []
    try:
        is_hol = is_holiday(target_date)
    except Exception:
        is_hol = False

    for h in range(24):
        if is_hol:
            # 土日祝日は9-17をホームタイム扱い
            if 7 <= h <= 22:  # 7-22時はすべてホーム
                colors.append(TIME_BANDS['home']['color'])
            else:
                colors.append(TIME_BANDS['night']['color'])
        else:
            if h in TIME_BANDS['day']['hours']:
                colors.append(TIME_BANDS['day']['color'])
            elif h in TIME_BANDS['home']['hours']:
                colors.append(TIME_BANDS['home']['color'])
            else:
                colors.append(TIME_BANDS['night']['color'])
    fig, ax = plt.subplots(figsize=(10,5))
    global LAST_FIG_OPEN
    LAST_FIG_OPEN = time.time()
    # Tkウインドウを作成
    master = tk._default_root if tk._default_root is not None else None
    if master is None:
        win = tk.Tk()
    else:
        win = tk.Toplevel(master)
    # Tkキャンバスに描画
    widget = None
    try:
        canvas = FigureCanvasTkAgg(fig, master=win) # キャンバス作成
        canvas.draw() # 描画
        widget = canvas.get_tk_widget() # キャンバスウィジェット取得
        widget.pack(side='top', fill='both', expand=1) # キャンバスをウインドウに配置
    except Exception:
        pass
    # ウインドウタイトルをファイルパスに設定
    if file_path is not None:
        try:
            win.wm_title(str(file_path)) # ウインドウタイトル設定
        except Exception:
            pass
    bars = ax.bar(range(24), values, color=colors) # 棒グラフ作成
    ax.set_ylabel('消費電力量 (kWh)') # Y軸ラベル
    ax.set_title(f'{year}-{month:02d}-{day:02d} 時間帯別電力使用量') # タイトル
    ax.set_xticks(range(24)) # X軸目盛り
    
    # 凡例
    legend_patches = [
        Patch(color=TIME_BANDS[band]['color'], label=TIME_BANDS[band]['label'])
        for band in DISPLAY_ORDER
    ]
    ax.legend(handles=legend_patches, bbox_to_anchor=(1.05, 1), loc='upper left') # 凡例表示
    
    # 0.5kWh単位でグリッド線
    ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7) # グリッド線
    ax.yaxis.set_major_locator(ticker.MultipleLocator(0.5)) # 0.5kWh単位
    ax.yaxis.set_minor_locator(ticker.AutoMinorLocator(0.1)) # 補助目盛り
    
    # 右軸に料金をプロットする: 時間ごとの単価は日付基準で同じと仮定
    # 各時間のコストは時間ごとのkWh * 該当時間帯単価 + kWh * (再エネ + 燃料)
    # 再エネ/燃料は日次合計に対して加算するため、ここでは時間ごとに按分して表示
    # 日次の合計再エネ+燃料単価
    renew = get_renewable_unit_for_date(target_date) # 再エネ賦課金単価
    fuel = get_fuel_adj_for_date(target_date) # 燃料費調整単価
    prices = get_unit_prices_for_date(target_date) # 時間帯単価取得
    hourly_unit = build_hourly_unit(prices, renew, fuel, is_hol) # 時間ごとの単価リスト作成
    hourly_costs = values * np.array(hourly_unit) # 各時間のコスト計算
    
    ax2 = ax.twinx() # 右側Y軸
    ax2.plot(range(24), hourly_costs, color='red', marker='o', linewidth=0.5, label='電気料金 (円)') # 料金プロット
    ax2.set_ylabel('電気料金 (円)') # 右Y軸ラベル
    ax2.yaxis.grid(False) # 右Y軸グリッド無効化
    
    # 左右y軸の0ラインを完全一致させる
    left_ylim = ax.get_ylim() # 左Y軸の範囲取得
    min_ylim = min(0, left_ylim[0]+0.2)  # 少し余裕を持たせる
    max_ylim = left_ylim[1]+0.2  # 少し余裕を持たせる
    ax.set_ylim(min_ylim, max_ylim) # 左Y軸の範囲設定
    ax2.set_ylim(min_ylim*50, max_ylim*50)  # 倍率
    ax2.legend(loc='upper right') # 右Y軸凡例表示
    
    # --- 右側の凡例下に料金・集計表示を追加 ---
    # 再エネ/燃料と時間帯単価、各時間帯合計と金額、総計を計算（共通化関数を利用）
    band_sums = band_sums_from_values(values, is_hol) # 各時間帯合計取得
    bd = compute_cost_breakdown(target_date, band_sums['day'], band_sums['home'], band_sums['night']) # 料金内訳計算
    base_costs = bd['base_costs'] # 時間帯別金額
    base_total = bd['base_total'] # 基本料金合計
    renew_amount = bd['renew_amount'] # 再エネ賦課金額
    fuel_amount = bd['fuel_amount'] # 燃料費調整額
    total_cost = bd['total_cost'] # 総合計金額
    prices = bd['prices'] # 時間帯単価
    renew = bd['renew'] # 再エネ賦課金単価
    fuel = bd['fuel'] # 燃料費調整単価
    total_kwh = bd.get('total_kwh', float(np.nansum(values))) # 総使用電力量

    # 単価表示
    unit_text = (f"適用単価:\n"
                 f"デイタイム単価　 {prices['day']:.2f} 円/kWh\n"
                 f"ホームタイム単価 {prices['home']:.2f} 円/kWh\n"
                 f"ナイトタイム単価 {prices['night']:.2f} 円/kWh\n"
                 f"再エネ賦課金単価 {renew:.2f} 円/kWh\n"
                 f"燃料費調整単価　 {fuel:.2f} 円/kWh")
    
    #  使用電力量合計
    band_lines = (f"使用電力量:\n"
                  f"デイタイム電力　 {band_sums['day']:.2f} kWh\n"
                  f"ホームタイム電力 {band_sums['home']:.2f} kWh\n"
                  f"ナイトタイム電力 {band_sums['night']:.2f} kWh\n"
                  f"総電力量　　　　 {total_kwh:.2f} kWh")
    
    # 買電金額（使用電力量合計に対する金額）
    totals = (f"集計金額:\n"
              f"デイタイム金額　 {base_costs['day']:.0f} 円\n"
              f"ホームタイム金額 {base_costs['home']:.0f} 円\n"
              f"ナイトタイム金額 {base_costs['night']:.0f} 円\n"
              f"再エネ賦課金額　 {renew_amount:.2f} 円\n"
              f"燃料調整費金額　 {fuel_amount:.2f} 円\n"
              f"合計金額　　　　 {total_cost:.0f} 円"
              )

    # 右側凡例下に表示
    summary_text = unit_text + "\n\n" + band_lines + "\n\n" + totals # まとめて表示
    ax.text(1.1, 0.75, summary_text, transform=ax.transAxes, ha='left', va='top', fontsize=9,
        bbox=dict(boxstyle='round,pad=0.4', fc='white', ec='black', alpha=0.85))
    plt.tight_layout() # レイアウト調整
    
    # --- インタラクティブ機能の追加 ---
    annotation = None # アノテーションオブジェクト
    fade_job = None # アノテーション消去用ジョブID

    # アノテーション削除関数
    def remove_annotation():
        nonlocal annotation, fade_job
        if annotation is not None:
            try:
                annotation.remove()
            except Exception:
                pass
            annotation = None
        if fade_job is not None and widget is not None:
            try:
                widget.after_cancel(fade_job)
            except Exception:
                pass
            fade_job = None
        try:
            fig.canvas.draw_idle()
        except Exception:
            pass

    # グラフ保存ダイアログ
    def save_via_dialog():
        try:
            title = getattr(ax, 'get_title', lambda: "graph")()
            fname = filedialog.asksaveasfilename(parent=win, defaultextension='.png', filetypes=[('PNG image', '*.png')], initialfile=title if title else "graph")
            if fname:
                fig.savefig(fname, bbox_inches='tight')
                print(f'Saved figure to {fname}')
        except Exception as e:
            print('Save failed:', e)
    
    # コンテキストメニュー表示
    def show_context_menu(event):
        try:
            menu = tk.Menu(widget, tearoff=0)
            menu.add_command(label='グラフ保存', command=save_via_dialog)
            menu.add_command(
                label='Dessmonitorデータ反映',
                command=lambda: load_dessmonitor_data_hourly(
                    values, target_date, colors, prices, renew, fuel, is_hol, file_path
                )
            )
            try:
                menu.tk_popup(widget.winfo_pointerx(), widget.winfo_pointery())
            finally:
                menu.grab_release()
        except Exception:
            pass

    # マウス移動イベントハンドラ 
    def on_motion(event): # グラフ上のマウス移動イベントハンドラ
        nonlocal annotation # アノテーションオブジェクト
        if getattr(event, 'inaxes', None) is None or not hasattr(event.inaxes, 'containers'): # グラフ外またはコンテナ無し
            remove_annotation()
            return
        for cont in ax.containers: # 棒グラフコンテナをループ
            for rect in cont: # 各棒グラフ矩形をループ
                try:
                    contains = rect.contains(event)[0] # マウス位置が矩形内か判定
                except Exception:
                    contains = False # 判定失敗時はFalse
                if contains:
                    # このバーの情報を表示
                    try:
                        xcenter = rect.get_x() + rect.get_width() / 2.0 # 棒グラフの中心X座標
                        hour = int(xcenter + 0.5) # 時間帯を計算
                    except Exception:
                        hour = int(rect.get_x()) # 失敗時はX座標を時間帯とする
                    value = rect.get_height() # 棒グラフの高さ（kWh）
                    band = hour_to_band(hour, is_hol) # 時間帯に対応する料金帯を取得
                    unit_price = prices[band] # 該当時間帯の単価取得
                    total_unit = unit_price + renew + fuel # 総単価計算
                    cost = value * total_unit # この時間帯の料金計算
                    msg = (f'{hour}時-{(hour+1)%24}時\n'
                           f'電力量: {value:.2f}kWh\n'
                           f'料金: {cost:.0f}円')
                    if annotation is not None: # 既存アノテーションがある場合
                        try:
                            annotation.remove() # 削除
                        except Exception:
                            pass
                    x = rect.get_x() + rect.get_width() / 2 # アノテーションX座標
                    y = rect.get_y() + rect.get_height() # アノテーションY座標
                    annotation = annotate_in_axes(ax, fig, x, y, msg) # アノテーション作成
                    return
        remove_annotation() # 棒グラフ外ならアノテーション削除

    # マウスボタンイベントハンドラ
    def on_button(event):
        if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1:
            # keep behavior: nothing to open for hourly
            return
        if getattr(event, 'button', None) == 3:
            show_context_menu(event)

    # イベント接続
    try:
        canvas = getattr(fig, 'canvas', None)
        if canvas:
            canvas.mpl_connect('motion_notify_event', on_motion)
            canvas.mpl_connect('figure_leave_event', lambda ev: remove_annotation())
            canvas.mpl_connect('button_press_event', on_button)
    except Exception:
        
        def on_click(event):
            nonlocal annotation
            clicked_on_bar = False
            for cont in ax.containers:
                for rect in cont:
                    try:
                        contains = rect.contains(event)[0]
                    except Exception:
                        contains = False
                    if contains:
                        if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1:
                            return
                        value = rect.get_height()
                        try:
                            xcenter = rect.get_x() + rect.get_width() / 2.0
                            hour = int(xcenter + 0.5)
                        except Exception:
                            hour = int(rect.get_x())
                        band = hour_to_band(hour, is_hol)
                        unit_price = prices[band]
                        total_unit = unit_price + renew + fuel
                        cost = value * total_unit
                        msg = (f'{hour}時-{(hour+1)%24}時\n'
                              f'電力量: {value:.2f}kWh\n'
                              f'料金: {cost:.0f}円')
                        if annotation is not None:
                            try:
                                annotation.remove()
                            except Exception:
                                pass
                        x = rect.get_x() + rect.get_width()/2
                        y = rect.get_y() + rect.get_height()
                        annotation = annotate_in_axes(ax, fig, x, y, msg)
                        clicked_on_bar = True
                        break
                if clicked_on_bar:
                    break
            if not clicked_on_bar and annotation is not None:
                try:
                    annotation.remove()
                except Exception:
                    pass
                annotation = None
                try:
                    fig.canvas.draw_idle()
                except Exception:
                    pass
        fig.canvas.mpl_connect('button_press_event', on_click)
        
    # ウインドウクローズ時の処理
    def _on_close(): # ウインドウクローズ時の処理
        try:
            plt.close(fig) # 図を閉じる
        except Exception:
            pass
        try:
            win.destroy() # ウインドウ破棄
        except Exception:
            pass
        try:
            if not tk._default_root: # メインウインドウの場合は終了
                sys.exit(0) # 終了
        except Exception:
            pass

    try:
        win.protocol('WM_DELETE_WINDOW', _on_close) # クローズイベント設定
    except Exception:
        pass
    # メインループ開始
    if master is None:
        try:
            win.mainloop()
        except Exception:
            pass



# 月間グラフを表示する関数
def plot_daily(df: pd.DataFrame, year_month: str, file_path: Path = None, dessmonitor_folder: str = None):
    global dessmonitor_data_cache_global # Dessmonitorデータキャッシュ
    if dessmonitor_data_cache_global is None: # 初回読み込み
        dessmonitor_data_cache_global = build_dessmonitor_data_cache(dessmonitor_folder) # Dessmonitorデータキャッシュを構築
    dessmonitor_data_dict = dessmonitor_data_cache_global # ローカル参照
    
    # 指定月のデータを抽出
    if isinstance(year_month, (pd.Timestamp, datetime.datetime, datetime.date)):
        ym = pd.Period(year_month, freq='M')
    else:
        ym = pd.Period(str(year_month), freq='M')
    df_month = df[df['date'].dt.to_period('M') == ym]
    if df_month.empty:
        print(f"{year_month} のデータがありません")
        return
    # 日付ごとに合計
    daily = df_month.groupby(df_month['date'].dt.day)[['day', 'home', 'night']].sum()
    # 共通の定義を使用
    colors = {band: TIME_BANDS[band]['color'] for band in TIME_BANDS}
    fig, ax = plt.subplots(figsize=(12,7))
    global LAST_FIG_OPEN
    LAST_FIG_OPEN = time.time()
    if file_path is not None:
        try:
            fig.canvas.manager.set_window_title(str(file_path))
        except Exception:
            pass
    bars = daily[COLUMN_ORDER].plot(kind='bar', stacked=True, color=[colors[col] for col in COLUMN_ORDER], ax=ax, legend=False)
    # hide x-axis label (pandas may set it to 'date')
    ax.set_xlabel('')
    ax.set_ylabel('消費電力量 (kWh)')
    ax.set_title(f'{year_month} 月間（日別）電力使用量')
    handles, _ = ax.get_legend_handles_labels()
    ax.legend(handles, [TIME_BANDS[col]['label'] for col in DISPLAY_ORDER], bbox_to_anchor=(1.05, 1), loc='upper left')
    # 5kWh単位でグリッド線
    ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
    
    ax.yaxis.set_major_locator(ticker.MultipleLocator(5))
    
    # 凡例下に月間集計情報を表示
    # 月の合計を計算 (daily DataFrame の合計を利用)
    total_cost = 0.0
    month_total_series = daily[['day', 'home', 'night']].sum()
    month_total = {k: float(month_total_series.get(k, 0.0)) for k in ['day', 'home', 'night']}
    # 各日のコストは個別に計算して合算（単価別内訳も集計）
    month_costs = {'day': 0.0, 'home': 0.0, 'night': 0.0, 'renew': 0.0, 'fuel': 0.0, 'total': 0.0}
    for day_idx in daily.index:
        date_row = df_month[df_month['date'].dt.day == day_idx]
        if not date_row.empty:
            d = date_row.iloc[0]
            bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night'])
            month_costs['day'] += bd['base_costs'].get('day', 0.0)
            month_costs['home'] += bd['base_costs'].get('home', 0.0)
            month_costs['night'] += bd['base_costs'].get('night', 0.0)
            month_costs['renew'] += bd.get('renew_amount', 0.0)
            month_costs['fuel'] += bd.get('fuel_amount', 0.0)
            month_costs['total'] += bd.get('total_cost', 0.0)
    
    # サンプルとして月の初日の単価を取得
    sample_date = df_month['date'].iloc[0] if not df_month.empty else pd.Timestamp(ym.year, ym.month, 1)
    prices = get_unit_prices_for_date(sample_date)
    renew = get_renewable_unit_for_date(sample_date)
    fuel = get_fuel_adj_for_date(sample_date)
    
    total_kwh = sum(month_total.values())
    
    # 情報テキストを作成
    # 単価表示
    unit_text = (f"適用単価:\n"
                 f"デイタイム単価　 {prices['day']:.2f} 円/kWh\n"
                 f"ホームタイム単価 {prices['home']:.2f} 円/kWh\n"
                 f"ナイトタイム単価 {prices['night']:.2f} 円/kWh\n"
                 f"再エネ賦課金単価 {renew:.2f} 円/kWh\n"
                 f"燃料費調整単価　 {fuel:.2f} 円/kWh")

    #  使用電力量合計
    usage_text = (f"月間使用電力量:\n"
                  f"デイタイム電力　 {month_total['day']:.2f} kWh\n"
                  f"ホームタイム電力 {month_total['home']:.2f} kWh\n"
                  f"ナイトタイム電力 {month_total['night']:.2f} kWh\n"
                  f"総使用電力量　　 {total_kwh:.2f} kWh")
    
    # 集計金額（使用電力量合計に対する金額）
    cost_text = (f"月間金額:\n"
                 f"デイタイム金額　 {month_costs['day']:.0f} 円\n"
                 f"ホームタイム金額 {month_costs['home']:.0f} 円\n"
                 f"ナイトタイム金額 {month_costs['night']:.0f} 円\n"
                 f"再エネ賦課金額　 {month_costs['renew']:.2f} 円\n"
                 f"燃料調整費金額　 {month_costs['fuel']:.2f} 円\n"
                 f"合計金額　　　　 {month_costs['total']:.0f} 円")
    
    # まとめて表示用テキスト
    summary_text = unit_text + "\n\n" + usage_text + "\n\n" + cost_text
    ax.text(1.1, 0.83, summary_text, transform=ax.transAxes,ha='left', va='top', fontsize=9,
        bbox=dict(boxstyle='round,pad=0.5', fc='white', ec='black', alpha=0.85))
    plt.tight_layout()

    # 各棒グラフに日付情報を設定
    days = daily.index.tolist()
    bar_rects = []
    for cont in getattr(ax, 'containers', []):
        for idx, rect in enumerate(cont):
            if idx < len(days):
                setattr(rect, '_day', int(days[idx]))
                bar_rects.append(rect)

    # 土日祝日を赤文字に
    tick_labels = []
    for d in days:
        date_row = df_month[df_month['date'].dt.day == d]
        if not date_row.empty:
            date_val = date_row.iloc[0]['date']
            if is_holiday(date_val):
                tick_labels.append({'label': str(d), 'color': 'red'})
            else:
                tick_labels.append({'label': str(d), 'color': 'black'})
        else:
            tick_labels.append({'label': str(d), 'color': 'black'})
    ax.set_xticks(range(len(days)))
    ax.set_xticklabels([tl['label'] for tl in tick_labels], rotation=0)
    for tick, tl in zip(ax.get_xticklabels(), tick_labels):
        tick.set_color(tl['color'])


    # dessmonitorデータを反映して月間グラフを表示する関数
    def plot_daily_dessmonitor(df: pd.DataFrame, year_month: str, file_path: Path = None, dessmonitor_folder: str = None):
        """日別Dessmonitorウインドウを表示する。グラフは時間別Dessmonitorを参考に描画。"""
        # 指定月のデータ抽出
        if isinstance(year_month, (pd.Timestamp, datetime.datetime, datetime.date)):
            ym = pd.Period(year_month, freq='M')
        else:
            ym = pd.Period(str(year_month), freq='M')
        df_month = df[df['date'].dt.to_period('M') == ym]
        if df_month.empty:
            print(f"{year_month} のデータがありません")
            return
        # 日付ごとに合計
        daily = df_month.groupby(df_month['date'].dt.day)[['day', 'home', 'night']].sum()
        
        # Dessmonitorデータ取得: 各日付ごとにファイル検索・集計
        global dessmonitor_data_cache_global # Dessmonitorデータキャッシュ
        dessmonitor_day_values = {} # 日別の時間帯別蓄電出力値格納用
        dessmonitor_day_pv_values = {} # 日別太陽光発電値格納用
        commercial_output_band = {band: [] for band in COLUMN_ORDER} # 商用充電時間帯別集計格納用
        commercial_bar_width = 0.6 / 2 # 電力会社グラフの半分の幅
        saving_costs = [] # 節約金額格納用
        
        for day_idx in daily.index: # 日付ごとに処理
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0] # サンプル日付取得
            date_str = sample_date.strftime('%Y-%m-%d') # 日付文字列
            pv_values = np.zeros(24) # 太陽光発電値格納用
            dess_values = np.zeros(24) # 蓄電出力値格納用
            commercial_hourly = np.zeros(24) # 商用充電値格納用
            # キャッシュから取得
            for h in range(24):
                key = (date_str, h) # キー作成
                vals_pv = dessmonitor_data_cache_global[key].get('PV_sum', []) if dessmonitor_data_cache_global and key in dessmonitor_data_cache_global else [] # 太陽光発電値取得
                pv_values[h] += sum(vals_pv) # 合計
                vals = dessmonitor_data_cache_global[key]['energy_sum'] if dessmonitor_data_cache_global and key in dessmonitor_data_cache_global else [] # 蓄電出力値取得
                dess_values[h] += sum(vals) # 合計
                vals_c = dessmonitor_data_cache_global[key]['charge_sum'] if dessmonitor_data_cache_global and key in dessmonitor_data_cache_global else [] # 商用充電値取得
                commercial_hourly[h] += sum(vals_c) # 合計
            dessmonitor_day_values[day_idx] = dess_values # 日別蓄電出力値を保存
            dessmonitor_day_pv_values[day_idx] = pv_values # 日別太陽光発電値を保存
            is_hol = is_holiday(sample_date) # 休日判定
            band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0} # 時間帯別合計初期化
            for h in range(24):
                band = hour_to_band(h, is_hol) # 時間帯判定
                band_sums[band] += commercial_hourly[h] # 商用充電値を時間帯別に集計
            for band in COLUMN_ORDER: # 各時間帯ごとに商用充電値を保存
                commercial_output_band[band].append(band_sums[band])
            dess_band_sums = band_sums_from_values(dessmonitor_day_values[day_idx], is_hol) # 蓄電出力値を時間帯別に集計
            pv_band_sums = band_sums_from_values(dessmonitor_day_pv_values[day_idx], is_hol) # 太陽光発電値を時間帯別に集計
            dess_bd = compute_cost_breakdown(sample_date, dess_band_sums['day'], dess_band_sums['home'], dess_band_sums['night']) # コスト内訳計算
            saving_costs.append(dess_bd['total_cost']) # 総コストを保存

        # --- 情報表示用の月間集計 ---
        month_total = daily[['day', 'home', 'night']].sum() # 電力会社データ: 1ヶ月間の時間帯別合計
        month_buy_total = month_total.sum() # 電力会社データ: 1ヶ月間の買電電力量合計
        dess_month_total = {'day': 0.0, 'home': 0.0, 'night': 0.0} # Dessmonitorデータ: 1ヶ月間の蓄電電力量合計
        pv_month_total = {'day': 0.0, 'home': 0.0, 'night': 0.0} # 太陽光発電データ: 1ヶ月間の発電量合計
        commercial_month_total = {'day': 0.0, 'home': 0.0, 'night': 0.0} # 商用充電データ: 1ヶ月間の時間帯別合計
        for i, day_idx in enumerate(daily.index): # 日付ごとに処理
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0] # サンプル日付取得
            is_hol = is_holiday(sample_date) # 休日判定
            dess_values = dessmonitor_day_values.get(day_idx, np.zeros(24)) # 蓄電出力値取得
            pv_values = dessmonitor_day_pv_values.get(day_idx, np.zeros(24)) # 太陽光発電値取得
            for h in range(24): # 蓄電出力値を時間帯別に集計
                band = hour_to_band(h, is_hol) # 時間帯判定
                dess_month_total[band] += dess_values[h]
                pv_month_total[band] += pv_values[h]
            # 商用充電
            for band in COLUMN_ORDER: # 各時間帯ごとに商用充電値を集計
                commercial_month_total[band] += commercial_output_band[band][i] # 商用充電値を集計
        dess_month_total_sum = sum(dess_month_total.values()) # Dessmonitorデータ: 1ヶ月間の蓄電電力量合計
        pv_month_total_sum = sum(pv_month_total.values()) # 太陽光発電データ: 1ヶ月間の発電量合計
        commercial_month_total_sum = sum(commercial_month_total.values()) # 商用充電データ: 1ヶ月間の商用充電電力量合計
        # グラフ描画
        fig, ax = plt.subplots(figsize=(12,7)) # 描画領域作成
        # プラス側: 電力会社データ（積み上げ棒グラフ）
        bar_width = 0.6 # 棒グラフの幅
        x = np.array(list(daily.index)) # x軸位置（日付）
        bars = ax.bar(x, daily['night'], width=bar_width, color=TIME_BANDS['night']['color'], label=TIME_BANDS['night']['label'], align='center') # ナイトタイム
        bars_home = ax.bar(x, daily['home'], width=bar_width, color=TIME_BANDS['home']['color'], bottom=daily['night'], label=TIME_BANDS['home']['label'], align='center') # ホームタイム
        bars_day = ax.bar(x, daily['day'], width=bar_width, color=TIME_BANDS['day']['color'], bottom=daily['night']+daily['home'], label=TIME_BANDS['day']['label'], align='center') # デイタイム
        # Dessmonitorデータ（日別・時間帯別集計の積み上げ棒グラフ、左y軸に統合）
        dess_colors = {
            'day': '#b3e5fc',
            'home': '#81d4fa',
            'night': '#4fc3f7'
        }
        dess_labels = {
            'day': 'デイタイム（蓄電）',
            'home': 'ホームタイム（蓄電）',
            'night': 'ナイトタイム（蓄電）'
        }
        # 商用出力データ（32列目=カラムindex31）を時間帯ごとに集計
        commercial_output_band = {band: [] for band in COLUMN_ORDER} # 商用充電時間帯別集計格納用
        commercial_bar_width = bar_width / 2  # 電力会社の半分の幅
        for day_idx in daily.index: # 日付ごとに処理
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0] # サンプル日付取得
            is_hol = is_holiday(sample_date) # 休日判定
            date_str = sample_date.strftime('%Y-%m-%d') # 日付文字列
            commercial_hourly = np.zeros(24) # 商用充電値格納用
            # キャッシュから取得
            for h in range(24):
                key = (date_str, h) # キー作成
                vals_c = dessmonitor_data_cache_global[key]['charge_sum'] if dessmonitor_data_cache_global and key in dessmonitor_data_cache_global else [] # 商用充電値取得
                commercial_hourly[h] += sum(vals_c) # 合計
            # 時間帯ごとに集計
            band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0} # 時間帯別合計初期化
            for h in range(24): # 24時間分ループ
                band = hour_to_band(h, is_hol) # 時間帯判定
                band_sums[band] += commercial_hourly[h] # 商用充電値を時間帯別に集計
            for band in COLUMN_ORDER: # 各時間帯ごとに商用充電値を保存
                commercial_output_band[band].append(band_sums[band]) # 商用充電値を保存
        # 太陽光発電折れ線グラフ用の日別データを事前に集計
        pv_daily_values = [] # 日別太陽光発電量合計
        for day_idx in daily.index:
            pv_values = dessmonitor_day_pv_values.get(day_idx, np.zeros(24))
            pv_daily_values.append(np.sum(-pv_values))
        #print(f"DEBUG: pv_daily_values = {pv_daily_values}, sum = {sum(pv_daily_values)}")
        
        for i, day_idx in enumerate(daily.index): # 日付ごとに処理
            dess_values = dessmonitor_day_values.get(day_idx, np.zeros(24)) # 蓄電出力値取得
            pv_values = dessmonitor_day_pv_values.get(day_idx, np.zeros(24)) # 太陽光発電値取得
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0] # サンプル日付取得
            is_hol = is_holiday(sample_date) # 休日判定
            # 時間帯ごとに集計
            band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0} # 時間帯別合計初期化
            pv_band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0} # 太陽光発電量時間帯別合計
            for h in range(24): # 24時間分ループ
                band = hour_to_band(h, is_hol) # 時間帯判定
                band_sums[band] += dess_values[h] # 蓄電出力値を時間帯別に集計
                pv_band_sums[band] += pv_values[h] # 太陽光発電値を時間帯別に集計
            # 積み上げ棒グラフとして表示（マイナス側、左y軸）
            bottom = 0.0 # 積み上げ底辺初期化
            for band in COLUMN_ORDER: # 各時間帯ごとに描画
                value = -band_sums[band] # マイナス値に変換
                ax.bar(day_idx, value, width=bar_width, bottom=bottom, color=dess_colors[band], alpha=0.7, label=dess_labels[band] if i==0 else "", align='center') # 棒グラフ描画
                bottom += value # 底辺更新
        
        # 太陽光発電折れ線グラフを描画
        ax.plot(x, pv_daily_values, color='orange', marker='', linewidth=1.2, markersize=4, label='太陽光発電') # 太陽光発電折れ線
            
        # 商用出力データをdeepskyblueで重ねる（時間帯ごとに積み上げ、幅は半分）
        commercial_bottom = np.zeros(len(x)) # 商用充電積み上げ底辺初期化
        for j, band in enumerate(COLUMN_ORDER): # 各時間帯ごとに描画
            values = commercial_output_band[band] # 商用充電値取得
            ax.bar(x, values, width=commercial_bar_width, bottom=commercial_bottom, color='deepskyblue', alpha=0.7, label='商用充電' if j==0 else "", align='center') # 棒グラフ描画
            commercial_bottom += np.array(values) # 底辺更新
        
        # 軸・ラベル・タイトル設定
        ax.set_ylabel('消費電力量 (kWh)') # 左y軸ラベル
        ax.set_title(f'{year_month} 月間（日別）電力使用量（Dessmonitor反映）') # タイトル
        ax.set_xticks(x) # x軸目盛位置設定
        # 土日祝日を赤文字に
        tick_labels = [] # 土日祝日ラベルリスト
        for d in x:
            date_row = df_month[df_month['date'].dt.day == d] # 日付行取得
            if not date_row.empty: # 行が存在する場合
                date_val = date_row.iloc[0]['date'] # 日付値取得
                if is_holiday(date_val): # 休日判定
                    tick_labels.append({'label': str(d), 'color': 'red'}) # 休日は赤文字
                else:
                    tick_labels.append({'label': str(d), 'color': 'black'}) # 平日は黒文字
            else:
                tick_labels.append({'label': str(d), 'color': 'black'}) # データなしは黒文字
        ax.set_xticklabels([tl['label'] for tl in tick_labels], rotation=0) # x軸目盛ラベル設定
        for tick, tl in zip(ax.get_xticklabels(), tick_labels): # 目盛ラベルの色設定
            tick.set_color(tl['color'])
        # グリッド線
        ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7) # y軸グリッド線
        ax.yaxis.set_major_locator(ticker.MultipleLocator(5)) # 5kWh単位で目盛設定
        # 右軸: 電気料金折れ線
        costs = [] # 日別電気料金格納用
        net_saving_costs = [] # 日別節約金額格納用
        for idx, day_idx in enumerate(daily.index): # 日付ごとに処理
            date_row = df_month[df_month['date'].dt.day == day_idx] # 日付行取得
            if not date_row.empty: # 行が存在する場合
                d = date_row.iloc[0] # 行データ取得
                bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night']) # コスト内訳計算
                c = bd.get('total_cost', 0.0) # 総コスト取得
            else:
                c = 0.0 # データなしは0.0
            costs.append(c) # 日別電気料金を保存
            # 商用充電分を差し引いた節約金額
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0] # サンプル日付取得
            is_hol = is_holiday(sample_date) # 休日判定
            dess_band_sums = band_sums_from_values(dessmonitor_day_values[day_idx], is_hol) # 蓄電出力値を時間帯別に集計
            dess_bd = compute_cost_breakdown(sample_date, dess_band_sums['day'], dess_band_sums['home'], dess_band_sums['night']) # コスト内訳計算
            # 商用充電
            commercial_band_sums = {band: commercial_output_band[band][idx] for band in COLUMN_ORDER} # 商用充電時間帯別集計取得
            commercial_bd = compute_cost_breakdown(sample_date, commercial_band_sums['day'], commercial_band_sums['home'], commercial_band_sums['night']) # コスト内訳計算
            net_saving = dess_bd['total_cost'] - commercial_bd['total_cost'] # 節約金額計算
            net_saving_costs.append(net_saving) # 日別節約金額を保存

        # 凡例
        legend_patches = [Patch(color=TIME_BANDS[c]['color'], label=TIME_BANDS[c]['label']) for c in DISPLAY_ORDER] # 電力会社凡例
        legend_patches += [Patch(color=dess_colors[c], label=dess_labels[c]) for c in DISPLAY_ORDER] # Dessmonitor凡例
        legend_patches.append(Patch(color='deepskyblue', label='商用充電')) # 商用充電凡例
        legend_patches.append(Patch(color='orange', label='太陽光発電')) # 太陽光発電凡例
        ax.legend(handles=legend_patches, bbox_to_anchor=(1.05, 1), loc='upper left') # 凡例表示
        # 集計情報（テキストボックス）
        if len(daily.index) > 0: # データが存在する場合
            # 月の初日で単価を取得
            day_idx = daily.index[0] # 最初の日付インデックス
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0] # サンプル日付取得
            prices = get_unit_prices_for_date(sample_date) # 単価取得
            renew = get_renewable_unit_for_date(sample_date) # 再エネ賦課金単価取得
            fuel = get_fuel_adj_for_date(sample_date) # 燃料費調整単価取得
            # 商用充電料金を計算
            commercial_charge_total = 0.0 # 商用充電料金初期化
            month_saving_total = 0.0 # 月間節約金額初期化
            for day_idx2 in daily.index: # 日付ごとに処理
                sample_date2 = df_month[df_month['date'].dt.day == day_idx2]['date'].iloc[0] # サンプル日付取得
                is_hol2 = is_holiday(sample_date2) # 休日判定
                commercial_band_sums2 = {band: commercial_output_band[band][list(daily.index).index(day_idx2)] for band in COLUMN_ORDER} # 商用充電時間帯別集計取得
                commercial_bd2 = compute_cost_breakdown(sample_date2, commercial_band_sums2['day'], commercial_band_sums2['home'], commercial_band_sums2['night']) # コスト内訳計算
                commercial_charge_total += commercial_bd2['total_cost'] # 商用充電料金を加算
            # 1ヶ月分の節約金額合計を計算し、商用充電料金を差し引く
                dess_band_sums2 = band_sums_from_values(dessmonitor_day_values[day_idx2], is_hol2) # 蓄電出力値を時間帯別に集計
                dess_bd2 = compute_cost_breakdown(sample_date2, dess_band_sums2['day'], dess_band_sums2['home'], dess_band_sums2['night']) # コスト内訳計算
                month_saving_total += dess_bd2['total_cost'] # 蓄電池節約金額を加算
            net_saving = month_saving_total - commercial_charge_total # 総合節約金額計算
            
            # 単価表示
            unit_text = (f"適用単価:\n"
                         f"デイタイム単価　 {prices['day']:.2f} 円/kWh\n"
                         f"ホームタイム単価 {prices['home']:.2f} 円/kWh\n"
                         f"ナイトタイム単価 {prices['night']:.2f} 円/kWh\n"
                         f"再エネ賦課金単価 {renew:.2f} 円/kWh\n"
                         f"燃料費調整単価　 {fuel:.2f} 円/kWh")
            
            #  使用電力量合計
            band_lines = (f"使用電力量:\n"
                          f"デイタイム電力　 {month_total['day']:.2f} kWh\n"
                          f"ホームタイム電力 {month_total['home']:.2f} kWh\n"
                          f"ナイトタイム電力 {month_total['night']:.2f} kWh\n"
                          f"買電電力量　　　 {month_buy_total:.2f} kWh\n"
                          f"蓄電電力量　　　 {dess_month_total_sum:.2f} kWh\n"
                          f"商用充電電力量　 {commercial_month_total_sum:.2f} kWh\n"
                          f"太陽光発電量　　 {-sum(pv_daily_values):.2f} kWh")
            
            # 集計金額（使用電力量合計に対する金額）
            totals = (f"集計金額:\n"
                      f"デイタイム金額　 {month_costs['day']:.0f} 円\n"
                      f"ホームタイム金額 {month_costs['home']:.0f} 円\n"
                      f"ナイトタイム金額 {month_costs['night']:.0f} 円\n"
                      f"再エネ賦課金額　 {month_costs['renew']:.2f} 円\n"
                      f"燃料調整費金額　 {month_costs['fuel']:.2f} 円\n"
                      f"買電金額　　　　 {month_costs['total']:.0f} 円\n"
                      f"蓄電池節約金額　 {month_saving_total:.0f} 円\n"
                      f"商用充電料金　　 {commercial_charge_total:.0f} 円\n"
                      f"総合節約金額　　 {net_saving:.0f} 円\n"
                      f"太陽光発電無金額 {month_costs['total'] + net_saving:.0f} 円")
            # まとめて表示用テキスト
            summary_text = unit_text + "\n\n" + band_lines + "\n\n" + totals # まとめテキスト作成
            ax.text(1.15, 0.65, summary_text, transform=ax.transAxes,ha='left', va='top', fontsize=9,
                bbox=dict(boxstyle='round,pad=0.5', fc='white', ec='black', alpha=0.85)) # テキスト表示
        plt.tight_layout() # レイアウト調整
        
        # Tkウインドウ表示
        win = tk.Toplevel() if tk._default_root else tk.Tk() # 新規ウインドウ作成
        win.wm_title(f'{year_month} 日別Dessmonitor') # ウインドウタイトル設定
        canvas = FigureCanvasTkAgg(fig, master=win) # Tkキャンバス作成
        canvas.draw() # 描画
        widget = canvas.get_tk_widget() # ウィジェット取得
        widget.pack(side='top', fill='both', expand=1) # ウィジェット配置
        
        # 画像保存の右クリックメニュー追加
        def save_via_dialog(): # グラフ保存ダイアログ表示関数
            try:
                title = getattr(ax, 'get_title', lambda: "graph")() # グラフタイトルを取得
                fname = filedialog.asksaveasfilename(parent=win, defaultextension='.png', filetypes=[('PNG image','*.png')], initialfile=title if title else "graph")
                if fname:
                    fig.savefig(fname, bbox_inches='tight')
                    print(f'Saved figure to {fname}')
            except Exception as e:
                print('Save failed:', e)

        def show_context_menu(event): # 右クリックメニュー表示関数
            try:
                menu = tk.Menu(widget, tearoff=0) # コンテキストメニュー作成
                menu.add_command(label='グラフ保存', command=save_via_dialog) # グラフ保存メニュー追加
                try:
                    menu.tk_popup(widget.winfo_pointerx(), widget.winfo_pointery()) # メニュー表示
                finally:
                    menu.grab_release()
            except Exception:
                pass
        widget.bind('<Button-3>', show_context_menu) # 右クリックイベントバインド


        # --- ホバー注釈: 各積み上げbarのnight/home/dayごとに対応 ---
        annotation = None
        fade_job = None
        def remove_annotation(): # マウスホバー注釈の削除関数
            nonlocal annotation, fade_job
            if annotation is not None:
                try:
                    annotation.remove()
                except Exception:
                    pass
                annotation = None
            if fade_job is not None and widget is not None:
                try:
                    widget.after_cancel(fade_job)
                except Exception:
                    pass
                fade_job = None
            try:
                fig.canvas.draw_idle()
            except Exception:
                pass

                # マウス移動イベントハンドラ
        def on_motion(event): # マウス移動イベントハンドラ
            nonlocal annotation
            if getattr(event, 'inaxes', None) is None: 
                remove_annotation()
                return
            
            found = False
            for container in ax.containers:
                for rect in container:
                    try:
                        contains, _ = rect.contains(event)
                    except Exception:
                        contains = False
                    if contains: # マウスが乗っている場合
                        day = getattr(rect, '_day', None)
                        if day is None:
                            continue
                        year, month = map(int, year_month.split('-'))
                        row = daily.loc[int(day)] if int(day) in daily.index else None
                        if row is None:
                            continue
                        date_row = df_month[df_month['date'].dt.day == int(day)]
                        if not date_row.empty:
                            d = date_row.iloc[0]
                            bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night'])
                        else:
                            bd = None
                        
                        # 蓄電量と発電量を取得
                        sample_date = df_month[df_month['date'].dt.day == int(day)]['date'].iloc[0] if not date_row.empty else None
                        date_str = sample_date.strftime('%Y-%m-%d') if sample_date else None
                        dess_total = 0.0
                        pv_total = 0.0
                        is_hol = is_holiday(sample_date) if sample_date else False
                        dess_band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0}
                        
                        if date_str and dessmonitor_data_cache_global:
                            for h in range(24):
                                key = (date_str, h)
                                if key in dessmonitor_data_cache_global:
                                    vals_dess = dessmonitor_data_cache_global[key].get('energy_sum', [])
                                    vals_pv = dessmonitor_data_cache_global[key].get('PV_sum', [])
                                    dess_total += sum(vals_dess)
                                    pv_total += sum(vals_pv)
                                    # 蓄電量を時間帯別に集計
                                    band = hour_to_band(h, is_hol)
                                    dess_band_sums[band] += sum(vals_dess)
                        
                        # 蓄電量の料金を計算
                        dess_bd = compute_cost_breakdown(sample_date, dess_band_sums['day'], dess_band_sums['home'], dess_band_sums['night']) if sample_date else None
                        
                        # 商用充電量を取得
                        commercial_total = 0.0
                        commercial_band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0}
                        if date_str and dessmonitor_data_cache_global:
                            for h in range(24):
                                key = (date_str, h)
                                if key in dessmonitor_data_cache_global:
                                    vals_c = dessmonitor_data_cache_global[key].get('charge_sum', [])
                                    commercial_total += sum(vals_c)
                                    # 商用充電量を時間帯別に集計
                                    band = hour_to_band(h, is_hol)
                                    commercial_band_sums[band] += sum(vals_c)
                        
                        # 商用充電の料金を計算
                        commercial_bd = compute_cost_breakdown(sample_date, commercial_band_sums['day'], commercial_band_sums['home'], commercial_band_sums['night']) if sample_date else None
                        
                        # 蓄電の節約金額 = 蓄電料金 - 商用充電料金
                        dess_saving_cost = (dess_bd['total_cost'] - commercial_bd['total_cost']) if (dess_bd and commercial_bd) else 0.0
                        
                        if bd is not None:
                            msg = (
                                f'{year}年{month}月{int(day)}日\n'
                                f'--- 買電 ---\n'
                                f'デイタイム: {row.get("day",0.0):.2f}kWh ({bd["base_costs"]["day"]:.0f}円)\n'
                                f'ホームタイム: {row.get("home",0.0):.2f}kWh ({bd["base_costs"]["home"]:.0f}円)\n'
                                f'ナイトタイム: {row.get("night",0.0):.2f}kWh ({bd["base_costs"]["night"]:.0f}円)\n'
                                f'買電計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: {bd["total_cost"]:.0f}円\n'
                                f'--- 蓄電 ---\n'
                                f'商用充電量: {commercial_total:.2f}kWh  料金: {commercial_bd["total_cost"]:.0f}円\n'
                                f'発電量: {pv_total:.2f}kWh\n'
                                f'蓄電量: {dess_total:.2f}kWh  料金: {dess_bd["total_cost"]:.0f}円\n'
                                f'蓄電節約金額: {dess_saving_cost:.0f}円'                     
                            )
                        else:
                            msg = (
                                f'{year}年{month}月{int(day)}日\n'
                                f'--- 買電 ---\n'
                                f'デイタイム: {row.get("day",0.0):.2f}kWh\n'
                                f'ホームタイム: {row.get("home",0.0):.2f}kWh\n'
                                f'ナイトタイム: {row.get("night",0.0):.2f}kWh\n'
                                f'買電量: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh\n'
                                f'--- 蓄電 ---\n'
                                f'商用充電量: {commercial_total:.2f}kWh  料金: {commercial_bd["total_cost"]:.0f}円\n'
                                f'発電量: {pv_total:.2f}kWh\n'
                                f'蓄電量: {dess_total:.2f}kWh  料金: {dess_bd["total_cost"]:.0f}円\n'
                                f'蓄電節約金額: {dess_saving_cost:.0f}円'                     
                            )
                        if annotation is not None:
                            try:
                                annotation.remove()
                            except Exception:
                                pass
                        x = rect.get_x() + rect.get_width()/2
                        y = rect.get_y() + rect.get_height() -5
                        # アノテーション作成（ha='left'で左寄せ）
                        annotation = ax.annotate(
                                msg, xy=(x, y), xytext=(5, 0), xycoords='data', textcoords='offset points',
                                ha='left', va='bottom', fontsize=9,
                                bbox=dict(boxstyle='round,pad=0.5', fc='yellow', ec='black', alpha=0.85),
                                zorder=1000
                            )
                        annotation.set_clip_on(False)
                        try:
                                fig.canvas.draw_idle()
                        except Exception:
                                pass
                        found = True
                        break
                    if found:
                        break
                if not found:
                    remove_annotation()
        
        def on_leave(event): # マウスがfigureから離れたときの処理
            remove_annotation()
        
        def on_button(event): # マウスクリックイベントハンドラ
            # ダブルクリック -> 時間別グラフ表示
            if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1: # 左ダブルクリック
                for rect in ax.patches:
                    try:
                        contains = rect.contains(event)[0]
                    except Exception:
                        contains = False
                    if contains: # 対象の棒グラフが見つかった場合
                        day = getattr(rect, '_day', None)
                        if day is not None: # 日付情報がある場合
                            year, month = map(int, year_month.split('-'))
                            plot_hourly(df, year, month, int(day), file_path=file_path, dessmonitor_folder=dessmonitor_folder)
                            return
            # 右クリック -> コンテキストメニュー表示
            if getattr(event, 'button', None) == 3: # 右クリック
                show_context_menu(event)

        fig.canvas.mpl_connect('motion_notify_event', on_motion) # マウス移動イベント接続
        fig.canvas.mpl_connect('figure_leave_event', on_leave) # マウス離脱イベント接続
        fig.canvas.mpl_connect('button_press_event', on_button) # マウスクリックイベント接続

        # 各棒グラフに日付情報を設定
        days = daily.index.tolist()
        for cont in [bars, bars_home, bars_day]:
            for idx, rect in enumerate(cont):
                if idx < len(days):
                    setattr(rect, '_day', int(days[idx]))

        # 左y軸と右y軸の0ラインを完全一致させる
        ax2 = ax.twinx() # 右y軸作成
        costs = [] # 日別電気料金格納用
        for day_idx in daily.index: # 日付ごとに処理
            date_row = df_month[df_month['date'].dt.day == day_idx] # 日付行取得
            if not date_row.empty: # 行が存在する場合
                d = date_row.iloc[0] # 行データ取得
                bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night']) # コスト内訳計算
                c = bd.get('total_cost', 0.0) # 総コスト取得
            else:
                c = 0.0 # データなしは0.0
            costs.append(c) # 日別電気料金を保存
        ax2.plot(daily.index, costs, color='red', marker='o', linewidth=0.5, label='電気料金 (円)') # 折れ線グラフ描画
        ax2.plot(daily.index, [-v for v in net_saving_costs], color='blue', marker='o', linewidth=0.5, label='節約金額 (円,反転)') # 折れ線グラフ描画
        ax2.set_ylabel('電気料金 (円)') # 右y軸ラベル
        # 0ラインを完全一致させる
        left_ylim = ax.get_ylim() # 左y軸の表示範囲取得
        min_ylim = min(0, left_ylim[0]) # 最小値設定
        max_ylim = left_ylim[1]+5 # 最大値設定（少し余裕を持たせる）
        ax.set_ylim(min_ylim, max_ylim) # 左y軸の表示範囲設定
        ax2.set_ylim(min_ylim*50, max_ylim*50) # 右y軸の表示範囲設定（倍率調整）
        ax2.legend(loc='upper right') # 右y軸凡例表示

        win.mainloop() # Tkウインドウのメインループ開始

    # 日別の料金を計算して右軸に描画
    costs = [] # 日別電気料金格納用
    for day_idx in daily.index: # 日付ごとに処理
        date_row = df_month[df_month['date'].dt.day == day_idx] # 日付行取得
        if not date_row.empty: # 行が存在する場合
            d = date_row.iloc[0] # 行データ取得
            bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night']) # コスト内訳計算
            c = bd.get('total_cost', 0.0) # 総コスト取得
        else:
            c = 0.0 # データなしは0.0
        costs.append(c) # 日別電気料金を保存
        
    ax2 = ax.twinx() # 右y軸作成
    ax2.plot(np.arange(len(daily.index)), costs, color='red', marker='o', linewidth=0.5, label='電気料金 (円)') # 折れ線グラフ描画
    ax2.set_ylabel('電気料金 (円)') # 右y軸ラベル
    # 左右y軸の0ラインを完全一致させる
    left_ylim = ax.get_ylim() # 左y軸の表示範囲取得
    min_ylim = min(0, left_ylim[0]) # 最小値設定
    max_ylim = left_ylim[1]+5 # 最大値設定（少し余裕を持たせる）
    ax.set_ylim(min_ylim, max_ylim) # 左y軸の表示範囲設定
    ax2.set_ylim(min_ylim*50, max_ylim*50)  # rough scaling factor
    ax2.legend(loc='upper right') # 右y軸凡例表示
    plt.tight_layout() # レイアウト調整
    
    # 画像保存の右クリックメニュー追加
    annotation = None
    fade_job = None

    # マウスホバー注釈の削除関数
    def remove_annotation():
        nonlocal annotation, fade_job
        if annotation is not None:
            try:
                annotation.remove()
            except Exception:
                pass
            annotation = None
        if fade_job is not None and widget is not None:
            try:
                widget.after_cancel(fade_job)
            except Exception:
                pass
            fade_job = None
        try:
            fig.canvas.draw_idle()
        except Exception:
            pass

    # マウス移動イベントハンドラ
    def save_via_dialog():
        try:
            # グラフタイトルを取得
            title = getattr(ax, 'get_title', lambda: "graph")( )
            fname = filedialog.asksaveasfilename(parent=win, defaultextension='.png', filetypes=[('PNG image','*.png')], initialfile=title if title else "graph")
            if fname:
                fig.savefig(fname, bbox_inches='tight')
                print(f'Saved figure to {fname}')
        except Exception as e:
            print('Save failed:', e)

    # コンテキストメニュー表示関数
    def show_context_menu(event):
        try:
            menu = tk.Menu(widget, tearoff=0)
            menu.add_command(label='グラフ保存', command=save_via_dialog)
            menu.add_command(label='Dessmonitorデータ反映', command=lambda: plot_daily_dessmonitor(df, year_month, file_path, dessmonitor_folder))
            try:
                menu.tk_popup(widget.winfo_pointerx(), widget.winfo_pointery())
            finally:
                menu.grab_release()
        except Exception:
            pass
        
    # --- ホバー注釈: 日別棒グラフ全体に対応 ---
    def on_motion(event): # マウス移動イベントハンドラ
        nonlocal annotation # ホバー注釈
        if getattr(event, 'inaxes', None) is None or not hasattr(event.inaxes, 'patches'): # 軸外またはパッチなし
            remove_annotation()
            return
        for rect in getattr(ax, 'patches', []): # 各棒グラフに対して
            try:
                contains = rect.contains(event)[0]
            except Exception:
                contains = False
            if contains: # マウスが乗っている場合
                day = getattr(rect, '_day', None)
                if day is None:
                    continue
                year, month = map(int, year_month.split('-'))
                row = daily.loc[int(day)] if int(day) in daily.index else None
                if row is None:
                    continue
                date_row = df_month[df_month['date'].dt.day == int(day)]
                if not date_row.empty:
                    d = date_row.iloc[0]
                    bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night'])
                else:
                    bd = None
                if bd is not None:
                    msg = (
                        f'{year}年{month}月{int(day)}日\n'
                        f'デイタイム: {row.get("day",0.0):.2f}kWh ({bd["base_costs"]["day"]:.0f}円)\n'
                        f'ホームタイム: {row.get("home",0.0):.2f}kWh ({bd["base_costs"]["home"]:.0f}円)\n'
                        f'ナイトタイム: {row.get("night",0.0):.2f}kWh ({bd["base_costs"]["night"]:.0f}円)\n'
                        f'合計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: {bd["total_cost"]:.0f}円'
                    )
                else:
                    msg = (
                        f'{year}年{month}月{int(day)}日\n'
                        f'デイタイム: {row.get("day",0.0):.2f}kWh\nホームタイム: {row.get("home",0.0):.2f}kWh\nナイトタイム: {row.get("night",0.0):.2f}kWh\n'
                        f'合計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: 0円'
                    )
                if annotation is not None:
                    try:
                        annotation.remove()
                    except Exception:
                        pass
                x = rect.get_x() + rect.get_width()/2
                y = rect.get_y() + rect.get_height()
                annotation = annotate_in_axes(ax, fig, x, y, msg)
                return
        remove_annotation()
    
    def on_button(event): # マウスクリックイベントハンドラ
        if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1: # 左ダブルクリック
            for axes in plt.gcf().axes: # 各軸に対して
                for rect in getattr(axes, 'patches', []): # 各棒グラフに対して
                    try:
                        contains = rect.contains(event)[0]
                    except Exception:
                        contains = False
                    if contains: # マウスが乗っている場合
                        day = getattr(rect, '_day', None)
                        if day is not None: # 日付情報がある場合
                            year, month = map(int, year_month.split('-'))
                            plot_hourly(df, year, month, int(day), file_path=file_path, dessmonitor_folder=dessmonitor_folder)
                            return
        if getattr(event, 'button', None) == 3: # 右クリック
            show_context_menu(event)

    try:
        fig.canvas.mpl_connect('motion_notify_event', on_motion) # マウス移動イベント接続
        fig.canvas.mpl_connect('figure_leave_event', lambda ev: remove_annotation()) # マウス離脱イベント接続
        fig.canvas.mpl_connect('button_press_event', on_button) # マウスクリックイベント接続
    except Exception:
        
        def on_click(event): # マウスクリックイベントハンドラ（簡易版）
            nonlocal annotation
            bar_clicked = False
            for axes in plt.gcf().axes: # 各軸に対して
                for rect in getattr(axes, 'patches', []):
                    try:
                        contains = rect.contains(event)[0]
                    except Exception:
                        contains = False
                    if contains: # マウスが乗っている場合
                        day = getattr(rect, '_day', None)
                        if day is not None: # 日付情報がある場合
                            year, month = map(int, year_month.split('-'))
                            if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1: # 左ダブルクリック
                                plot_hourly(df, year, month, int(day), file_path=file_path)
                            else:
                                row = daily.loc[int(day)] if int(day) in daily.index else None
                                if row is not None: 
                                    date_row = df_month[df_month['date'].dt.day == int(day)]
                                    if not date_row.empty: 
                                        d = date_row.iloc[0]
                                        bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night'])
                                    else:
                                        bd = None
                                    if bd is not None: 
                                        msg = (
                                            f'{year}年{month}月{int(day)}日\n'
                                            f'デイタイム: {row.get("day",0.0):.2f}kWh ({bd["base_costs"]["day"]:.0f}円)\n'
                                            f'ホームタイム: {row.get("home",0.0):.2f}kWh ({bd["base_costs"]["home"]:.0f}円)\n'
                                            f'ナイトタイム: {row.get("night",0.0):.2f}kWh ({bd["base_costs"]["night"]:.0f}円)\n'
                                            f'合計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: {bd["total_cost"]:.0f}円'
                                        )
                                    else:
                                        msg = (
                                            f'{year}年{month}月{int(day)}日\n'
                                            f'デイタイム: {row.get("day",0.0):.2f}kWh\nホームタイム: {row.get("home",0.0):.2f}kWh\nナイトタイム: {row.get("night",0.0):.2f}kWh\n'
                                            f'合計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: 0円'
                                        )
                                    if annotation is not None: 
                                        annotation.remove()
                                    x = rect.get_x() + rect.get_width()/2
                                    y = rect.get_y() + rect.get_height()
                                    try:
                                        ylim = axes.get_ylim()
                                        y = min(y + (ylim[1]-ylim[0]) * 0.03, ylim[1] * 0.95)
                                    except Exception:
                                        pass
                                    annotation = axes.annotate(msg, xy=(x, y), xytext=(0, 0), xycoords='data', textcoords='offset points',
                                        ha='center', va='bottom', fontsize=10, bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.9), zorder=1000)
                                    annotation.set_clip_on(True)
                                    fig.canvas.draw_idle()
                            bar_clicked = True
                        break
                if bar_clicked: 
                    break
            if not bar_clicked and annotation is not None: # 棒グラフ以外をクリックした場合
                try:
                    annotation.remove()
                except Exception:
                    pass
                annotation = None
                try:
                    fig.canvas.draw_idle()
                except Exception:
                    pass
        fig.canvas.mpl_connect('button_press_event', on_click) # マウスクリックイベント接続
        
    # Tkウインドウ表示
    master = tk._default_root if tk._default_root is not None else None # 親ウインドウ取得
    if master is None: 
        win = tk.Tk() # 新規ウインドウ作成
    else:
        win = tk.Toplevel(master) # 親ウインドウの子ウインドウ作成

    # ウインドウタイトルをファイルパスに設定
    if file_path is not None:
        try:
            win.wm_title(str(file_path)) # ウインドウタイトル設定
        except Exception:
            pass

    try:
        canvas = FigureCanvasTkAgg(fig, master=win) # Tkキャンバス作成
        canvas.draw() # 描画
        widget = canvas.get_tk_widget() # ウィジェット取得
        widget.pack(side='top', fill='both', expand=1) # ウィジェット配置
    except Exception:
        try:
            plt.show()
            return
        except Exception:
            pass

    annotation_local = None # ホバー注釈
    
    # ウインドウクローズ時の処理
        # ウインドウクローズ時の処理
    def _on_close(): # ウインドウクローズ時の処理
        try:
            plt.close(fig) # 図を閉じる
        except Exception:
            pass
        try:
            win.destroy() # ウインドウ破棄
        except Exception:
            pass
        try:
            if not tk._default_root: # メインウインドウの場合は終了
                sys.exit(0) # 終了
        except Exception:
            pass
    try:
        win.protocol('WM_DELETE_WINDOW', _on_close) # ウインドウクローズイベント設定
    except Exception:
        pass

    # メインループ開始（親ウインドウがない場合のみ）
    if master is None:
        try:
            win.mainloop() # Tkウインドウのメインループ開始
        except Exception:
            pass

# 月別グラフをウインドウ表示し、クリックで日別グラフを表示
def plot_monthly_interactive(monthly: pd.DataFrame, df: pd.DataFrame, file_path: Path = None, dessmonitor_folder: str = None):

    # --- 年間（月別）Dessmonitor反映ウインドウ ---
    def plot_monthly_dessmonitor(monthly: pd.DataFrame, df: pd.DataFrame, file_path: Path = None, dessmonitor_folder: str = None):
        """月別Dessmonitorウインドウを表示する。グラフは日別Dessmonitorを参考に、月単位で集計・描画。"""
        
        # --- キャッシュ構築 ---
        global dessmonitor_data_cache_global # グローバルキャッシュ変数
        if dessmonitor_data_cache_global is None: # キャッシュが未構築の場合
            dessmonitor_data_cache_global = build_dessmonitor_data_cache(dessmonitor_folder) # キャッシュ構築
        dessmonitor_data_dict = dessmonitor_data_cache_global # キャッシュデータ取得
        
        # 月ごとに集計
        months = sorted(list(set(df['date'].dt.to_period('M')))) # 月リスト取得
        # 月ごとの電力会社データ
        monthly_data = df.copy() # 元データコピー
        monthly_data['month'] = monthly_data['date'].dt.to_period('M') # 月列追加
        month_group = monthly_data.groupby('month')[['day', 'home', 'night']].sum() # 月・時間帯ごとに集計
        # Dessmonitorデータ・商用充電データを月ごとに集計
        # --- キャッシュdictをDataFrame化して高速集計 ---
        # キャッシュdict: {(date, hour): {'energy_sum': [...], 'charge_sum': [...]}}
        cache_rows = [] # キャッシュ行リスト初期化
        for (date_str, hour), vals in dessmonitor_data_dict.items(): 
            # 日付文字列をdatetime化
            try:
                date_dt = pd.to_datetime(date_str) # 日付文字列をdatetime化
            except Exception:
                continue
            month_period = date_dt.to_period('M') # 月Period取得
            is_hol = is_holiday(date_dt) # 休日判定
            band = hour_to_band(hour, is_hol) # 時間帯取得
            PV_sum = sum(vals['PV_sum']) # 太陽光発電量合計
            energy_sum = sum(vals['energy_sum']) # 蓄電電力量合計
            charge_sum = sum(vals['charge_sum']) # 商用充電電力量合計
            cache_rows.append({ 
                'month': month_period,
                'band': band,
                'PV_sum': PV_sum,
                'energy_sum': energy_sum,
                'charge_sum': charge_sum
            })
        cache_df = pd.DataFrame(cache_rows) # キャッシュDataFrame化
        # 月・時間帯ごとに合算
        dess_month_total = {m: {'day':0.0, 'home':0.0, 'night':0.0} for m in months} # 蓄電月別合計初期化
        commercial_month_total = {m: {'day':0.0, 'home':0.0, 'night':0.0} for m in months} # 商用充電月別合計初期化
        pv_month_total = {m: {'day':0.0, 'home':0.0, 'night':0.0} for m in months} # 太陽光発電月別合計初期化
        if not cache_df.empty: # キャッシュDataFrameが空でない場合
            grouped = cache_df.groupby(['month', 'band']).sum() # 月・時間帯ごとに集計
            for m in months: # 各月に対して
                for band in COLUMN_ORDER: # 各時間帯に対して
                    # 発電
                    val_pv = grouped.loc[(m, band), 'PV_sum'] if (m, band) in grouped.index else 0.0 
                    pv_month_total[m][band] = val_pv
                    # 蓄電
                    val = grouped.loc[(m, band), 'energy_sum'] if (m, band) in grouped.index else 0.0 
                    dess_month_total[m][band] = val 
                    # 商用充電
                    val_c = grouped.loc[(m, band), 'charge_sum'] if (m, band) in grouped.index else 0.0 
                    commercial_month_total[m][band] = val_c 
        # グラフ描画
        fig, ax = plt.subplots(figsize=(15,7)) # 図と軸作成
        bar_width = 0.8 # 棒グラフ幅設定
        x = np.arange(len(months)) # x位置設定
        # プラス側: 電力会社データ（積み上げ棒グラフ）
        bars = ax.bar(x, [month_group.loc[m]['night'] if m in month_group.index else 0 for m in months], width=bar_width, color=TIME_BANDS['night']['color'], label=TIME_BANDS['night']['label'], align='center')
        bars_home = ax.bar(x, [month_group.loc[m]['home'] if m in month_group.index else 0 for m in months], width=bar_width, color=TIME_BANDS['home']['color'], bottom=[month_group.loc[m]['night'] if m in month_group.index else 0 for m in months], label=TIME_BANDS['home']['label'], align='center')
        bars_day = ax.bar(x, [month_group.loc[m]['day'] if m in month_group.index else 0 for m in months], width=bar_width, color=TIME_BANDS['day']['color'], bottom=[(month_group.loc[m]['night']+month_group.loc[m]['home']) if m in month_group.index else 0 for m in months], label=TIME_BANDS['day']['label'], align='center')
        # Dessmonitorデータ（月別・時間帯別集計の積み上げ棒グラフ、左y軸に統合、マイナス側）
        dess_colors = {'day': '#b3e5fc', 'home': '#81d4fa', 'night': '#4fc3f7'} # Dessmonitor用色設定
        dess_labels = {'day': 'デイタイム（蓄電）', 'home': 'ホームタイム（蓄電）', 'night': 'ナイトタイム（蓄電）'} # Dessmonitor用ラベル設定
        bottom = np.zeros(len(months)) # Dessmonitor積み上げ底辺初期化
        commercial_bar_width = bar_width / 2 # 商用充電棒グラフ幅設定
        commercial_bottom = np.zeros(len(months)) # 商用充電積み上げ底辺初期化
        for band in COLUMN_ORDER: # 各時間帯に対して
            values = [-dess_month_total[m][band] for m in months] # 蓄電データをマイナス値に変換
            ax.bar(x, values, width=bar_width, bottom=bottom, color=dess_colors[band], alpha=0.7, label=dess_labels[band], align='center') 
            bottom += values
        # 商用充電データをdeepskyblueで重ねる（時間帯ごとに積み上げ、幅は半分）
            values = [commercial_month_total[m][band] for m in months] # 商用充電データ取得
            ax.bar(x, values, width=commercial_bar_width, bottom=commercial_bottom, color='deepskyblue', alpha=0.7, label='商用充電' if band=='day' else "", align='center')
            commercial_bottom += np.array(values) 
        # 太陽光発電折れ線グラフ（各月の太陽光発電合計）
        pv_month_values = [-sum(pv_month_total[m].values()) for m in months] # 各月の太陽光発電合計
        ax.plot(x, pv_month_values, color='orange', marker='', linewidth=1.2, markersize=4, label='太陽光発電') # 太陽光発電折れ線
        # --- 金額グラフ追加 ---
        cost_list = [] # 電気料金リスト
        dess_saving_list = [] # 蓄電節約金額リスト
        commercial_cost_list = [] # 商用充電金額リスト
        for m in months: 
            # 月の初日を代表日とする
            date_for_price = pd.Timestamp(m.start_time) 
            # 電力会社データ
            kwhs = month_group.loc[m] if m in month_group.index else {'day':0,'home':0,'night':0} 
            cost = compute_cost_breakdown(date_for_price, kwhs['day'], kwhs['home'], kwhs['night']) 
            cost_list.append(cost.get('total_cost', 0.0)) 
            # 商用充電金額
            comm_kwhs = commercial_month_total[m] 
            comm_cost = compute_cost_breakdown(date_for_price, comm_kwhs['day'], comm_kwhs['home'], comm_kwhs['night']) 
            commercial_cost_list.append(comm_cost.get('total_cost', 0.0)) 
            # 蓄電節約金額
            dess_kwhs = dess_month_total[m]
            dess_saving = compute_cost_breakdown(date_for_price, dess_kwhs['day'], dess_kwhs['home'], dess_kwhs['night'])
            dess_saving = dess_saving.get('total_cost',0.0) - comm_cost.get('total_cost',0.0)  # 商用充電金額を差し引く
            dess_saving_list.append(dess_saving)
            
        ax2 = ax.twinx() # 右y軸作成
        ax2.plot(x, cost_list, color='red', marker='o', linewidth=0.5, label='電気料金(円)') # 折れ線グラフ描画
        ax2.plot(x, [-v for v in dess_saving_list], color='blue', marker='o', linewidth=0.5, label='蓄電節約金額(円)') # 折れ線グラフ描画

        ax2.set_ylabel('金額 (円)') # 右y軸ラベル
        ax2.legend(loc='upper right') # 右y軸凡例表示
        # 0ラインを完全一致させる
        left_ylim = ax.get_ylim() # 左y軸の表示範囲取得
        min_ylim = min(0, left_ylim[0]) # 最小値設定
        max_ylim = left_ylim[1]+5 # 最大値設定（少し余裕を持たせる）
        ax.set_ylim(min_ylim, max_ylim) # 左y軸の表示範囲設定
        ax2.set_ylim(min_ylim*40, max_ylim*40) # 右y軸の表示範囲設定（倍率調整）
        ax2.legend(loc='upper right') # 右y軸凡例表示

        # 軸・ラベル・タイトル
        ax.set_ylabel('消費電力量 (kWh)') # 左y軸ラベル
        ax.set_title('年間（月別）電力使用量・金額（Dessmonitor反映）') # グラフタイトル
        ax.set_xticks(x + width * (len(years) - 1) / 2 -0.2) # x軸目盛位置設定
        ax.set_xticklabels([f'{m.month}月' for m in months]) # x軸目盛ラベル設定
        ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7) # y軸グリッド表示
        ax.yaxis.set_major_locator(ticker.MultipleLocator(100)) # y軸主目盛間隔設定
        ax.tick_params(axis='x', pad=10)  # パディングを10ポイント設定して左へ移動
         # 年の境界に縦グリッドを追加
        if len(years) > 1: 
            for i in range(1, len(years)): 
                xpos = 12 *( i - 1 )+ 0.5 # x位置計算
                ax.axvline(x=xpos, color='black', linestyle=':', linewidth=1.2, alpha=0.7) # 縦グリッド描画

        # 凡例
        legend_patches = [Patch(color=TIME_BANDS[c]['color'], label=TIME_BANDS[c]['label']) for c in DISPLAY_ORDER] # 時間帯凡例パッチ作成
        legend_patches += [Patch(color=dess_colors[c], label=dess_labels[c]) for c in DISPLAY_ORDER] # Dessmonitor凡例パッチ作成
        legend_patches.append(Patch(color='deepskyblue', label='商用充電')) # 商用充電凡例パッチ追加
        legend_patches.append(plt.Line2D([0], [0], color='orange', lw=1.2, label='太陽光発電')) # 太陽光発電凡例パッチ追加
        ax.legend(handles=legend_patches, bbox_to_anchor=(1.05, 1), loc='upper left') # 凡例表示
        
        # --- 年間集計情報を表示 ---
        # 年の境界を取得
        years_in_data = sorted(list(set(df['date'].dt.year)))
        
        # 年別集計を計算
        annual_summary = {}
        for year in years_in_data:
            year_months = [m for m in months if m.year == year]
            year_data = {
                'day': 0.0, 'home': 0.0, 'night': 0.0,
                'dess_day': 0.0, 'dess_home': 0.0, 'dess_night': 0.0,
                'pv_day': 0.0, 'pv_home': 0.0, 'pv_night': 0.0,
                'commercial_day': 0.0, 'commercial_home': 0.0, 'commercial_night': 0.0,
                'cost': 0.0, 'dess_saving': 0.0, 'commercial_cost': 0.0
            }
            for m in year_months:
                # 買電データ
                if m in month_group.index:
                    year_data['day'] += month_group.loc[m]['day']
                    year_data['home'] += month_group.loc[m]['home']
                    year_data['night'] += month_group.loc[m]['night']
                # 蓄電データ
                year_data['dess_day'] += dess_month_total[m]['day']
                year_data['dess_home'] += dess_month_total[m]['home']
                year_data['dess_night'] += dess_month_total[m]['night']
                # 発電データ
                year_data['pv_day'] += pv_month_total[m]['day']
                year_data['pv_home'] += pv_month_total[m]['home']
                year_data['pv_night'] += pv_month_total[m]['night']
                # 商用充電データ
                year_data['commercial_day'] += commercial_month_total[m]['day']
                year_data['commercial_home'] += commercial_month_total[m]['home']
                year_data['commercial_night'] += commercial_month_total[m]['night']
                # 金額
                month_idx = list(months).index(m) if m in months else -1
                if month_idx >= 0 and month_idx < len(cost_list):
                    year_data['cost'] += cost_list[month_idx]
                    year_data['dess_saving'] += dess_saving_list[month_idx]
                    year_data['commercial_cost'] += commercial_cost_list[month_idx]
            annual_summary[year] = year_data
        
        # 年別集計情報のテキストを作成
        summary_text = "年間集計情報:\n\n"
        for year in years_in_data:
            data = annual_summary[year]
            total_kwh = data['day'] + data['home'] + data['night']
            total_dess_kwh = data['dess_day'] + data['dess_home'] + data['dess_night']
            total_pv_kwh = data['pv_day'] + data['pv_home'] + data['pv_night']
            total_commercial_kwh = data['commercial_day'] + data['commercial_home'] + data['commercial_night']
            net_saving = data['dess_saving'] - data['commercial_cost']
            
            summary_text += (
                f"{year}年:\n"
                f"買電: {total_kwh:.2f}kWh ({data['cost']:.0f}円)\n"
                f"蓄電: {total_dess_kwh:.2f}kWh (- {data['dess_saving']:.0f}円)\n"
                f"発電: {total_pv_kwh:.2f}kWh\n"
                f"商用充電: {total_commercial_kwh:.2f}kWh ({data['commercial_cost']:.0f}円)\n"
                f"総合節約: {net_saving:.0f}円\n\n"
            )
        
        # 凡例下に集計情報を表示
        ax.text(1.1, 0.65, summary_text, transform=ax.transAxes, ha='left', va='top', fontsize=9,
                bbox=dict(boxstyle='round,pad=0.5', fc='white', ec='black', alpha=0.85))
        
        plt.tight_layout() # レイアウト調整
        
        # Tkウインドウ表示
        win = tk.Toplevel() if tk._default_root else tk.Tk() # ウインドウ作成
        win.wm_title('年間（月別）Dessmonitor') # ウインドウタイトル設定
        canvas = FigureCanvasTkAgg(fig, master=win) # Tkキャンバス作成
        canvas.draw()  # 描画
        widget = canvas.get_tk_widget() # ウィジェット取得
        widget.pack(side='top', fill='both', expand=1) # ウィジェット配置
                
        # ホバー注釈用変数
        annotation = None
        fade_job = None
        
        # アノテーション削除関数
        def remove_annotation():
            nonlocal annotation, fade_job
            if annotation is not None:
                try:
                    annotation.remove()
                except Exception:
                    pass
                annotation = None
            if fade_job is not None and widget is not None:
                try:
                    widget.after_cancel(fade_job)
                except Exception:
                    pass
                fade_job = None
            try:
                fig.canvas.draw_idle()
            except Exception:
                pass
        
        # マウス移動イベントハンドラ
                # マウス移動イベントハンドラ
        def on_motion(event):
            nonlocal annotation
            if getattr(event, 'inaxes', None) is None or not hasattr(event.inaxes, 'patches'):
                remove_annotation()
                return
            # 各棒グラフ（bars, bars_home, bars_day）に対して検索
            for cont, band in zip([bars, bars_home, bars_day], ['night', 'home', 'day']):
                for idx, rect in enumerate(cont):
                    try:
                        contains, _ = rect.contains(event)
                    except Exception:
                        contains = False
                    if contains:
                        if idx >= len(months):
                            continue
                        m = months[idx]
                        # 対応する月データを取得
                        sample_date = pd.Timestamp(m.start_time)
                        row = month_group.loc[m] if m in month_group.index else {'day': 0.0, 'home': 0.0, 'night': 0.0}
                        bd = compute_cost_breakdown(sample_date, row['day'], row['home'], row['night'])
                        
                        # 蓄電データ
                        dess_bd = compute_cost_breakdown(sample_date, dess_month_total[m]['day'], 
                                                          dess_month_total[m]['home'], dess_month_total[m]['night'])
                        
                        # 商用充電データ
                        commercial_bd = compute_cost_breakdown(sample_date, commercial_month_total[m]['day'],
                                                                commercial_month_total[m]['home'], commercial_month_total[m]['night'])
                        
                        # 発電データ
                        pv_total = sum(pv_month_total[m].values())
                        
                        # メッセージ作成
                        msg = (
                            f'{m.year}年{m.month}月\n'
                            f'時間帯別: {row["day"]+row["home"]+row["night"]:.2f}kWh\n'
                            f'金額: {bd.get("total_cost", 0.0):.0f}円\n'
                            f'--- {TIME_BANDS[band]["label"].split("(")[0]} ---\n'
                            f'買電量: {row[band]:.2f}kWh\n'
                            f'\n'
                            f'商用充電量: {sum(commercial_month_total[m].values()):.2f}kWh\n'
                            f'発電量　　: {pv_total:.2f}kWh\n'
                            f'蓄電量　　: {sum(dess_month_total[m].values()):.2f}kWh\n'
                            f'節約金額　: {dess_bd.get("total_cost", 0.0):.0f}円'
                        )
                        
                        if annotation is not None:
                            try:
                                annotation.remove()
                            except Exception:
                                pass
                        
                        x = rect.get_x() + rect.get_width() / 2
                        y = rect.get_y() + rect.get_height() - 500

                        # アノテーション作成（ha='left'で左寄せ）
                        annotation = ax.annotate(
                            msg, xy=(x, y), xytext=(5, 0), xycoords='data', textcoords='offset points',
                            ha='left', va='bottom', fontsize=9,
                            bbox=dict(boxstyle='round,pad=0.5', fc='yellow', ec='black', alpha=0.85),
                            zorder=1000
                        )
                        annotation.set_clip_on(False) # クリップオフ
                        try:
                            fig.canvas.draw_idle() # 再描画
                        except Exception:
                            pass
                        return
            remove_annotation()
        
        # マウス離脱イベントハンドラ
        def on_leave(event):
            remove_annotation()

                # マウスクリックイベントハンドラ
        def on_button(event): # マウスクリックイベントハンドラ
            # ダブルクリック -> 日別グラフ表示
            if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1: # 左ダブルクリック
                for i, m in enumerate(months):
                    # 各棒グラフ（bars, bars_home, bars_day）をチェック
                    for cont in [bars, bars_home, bars_day]:
                        if i >= len(cont):
                            continue
                        rect = cont[i]
                        try:
                            contains, _ = rect.contains(event)
                        except Exception:
                            contains = False
                        if contains: # 対象の棒グラフが見つかった場合
                            year = m.year
                            month = m.month
                            year_month = f'{year}-{int(month):02d}'
                            plot_daily(df, year_month, file_path=file_path, dessmonitor_folder=dessmonitor_folder)
                            return
            # 右クリック -> コンテキストメニュー表示
            if getattr(event, 'button', None) == 3: # 右クリック
                show_context_menu(event) # コンテキストメニュー表示
        
        # イベント接続
        fig.canvas.mpl_connect('motion_notify_event', on_motion)
        fig.canvas.mpl_connect('figure_leave_event', on_leave)
        fig.canvas.mpl_connect('button_press_event', on_button)
        
        # 右クリックメニュー（グラフ保存）
        def save_via_dialog():
            try:
                title = getattr(ax, 'get_title', lambda: "graph")()
                fname = filedialog.asksaveasfilename(parent=win, defaultextension='.png', filetypes=[('PNG image','*.png')], initialfile=title if title else "graph")
                if fname:
                    fig.savefig(fname, bbox_inches='tight')
                    print(f'Saved figure to {fname}')
            except Exception as e:
                print('Save failed:', e)
        
        # コンテキストメニュー表示関数
        def show_context_menu(event):
            try:
                menu = tk.Menu(widget, tearoff=0)
                menu.add_command(label='グラフ保存', command=save_via_dialog)
                try:
                    menu.tk_popup(widget.winfo_pointerx(), widget.winfo_pointery())
                finally:
                    menu.grab_release()
            except Exception:
                pass
        
        widget.bind('<Button-3>', show_context_menu)
        win.mainloop()
        
    """月別積み上げ棒グラフ（年比較）を描画し、年別の月次電気料金を色分けの折れ線で表示する。
    バーをクリックすると該当年月の日別グラフを開く。
    """
    base_colors = {band: TIME_BANDS[band]['color_rgb'] for band in TIME_BANDS} # 基本色RGB取得
    plot_data = monthly[COLUMN_ORDER].copy() # プロット用データコピー
    plot_data.index = pd.to_datetime(plot_data.index.astype(str) + '-01') # インデックスを月初日時刻に変換
    
    # 年別の表示/非表示状態を管理する辞書
    year_visibility = {} # 年別表示状態辞書初期化

    years = sorted(plot_data.index.year.unique()) # 年リスト取得
    months = list(range(1, 13)) # 月リスト取得

    fig, ax = plt.subplots(figsize=(15, 7)) # 図と軸作成
    global LAST_FIG_OPEN # 最後に開いた図の時刻
    LAST_FIG_OPEN = time.time() # 最後に開いた図の時刻更新
    if file_path is not None: 
        try:
            fig.canvas.manager.set_window_title(str(file_path)) # ウインドウタイトル設定
        except Exception:
            pass

    x = np.arange(len(months)) # x位置設定
    width = 0.8 / max(1, len(years)) # 棒グラフ幅設定（年数に応じて調整）

    # 描画: 年ごとに並べた積み上げ棒
    for i, year in enumerate(years): 
        year_rows = []
        for m in months:
            ts = pd.Timestamp(f'{year}-{m:02d}-01') 
            if ts in plot_data.index: 
                year_rows.append(plot_data.loc[ts]) 
            else:
                year_rows.append(pd.Series({c: 0.0 for c in COLUMN_ORDER})) 
        year_df = pd.DataFrame(year_rows, columns=COLUMN_ORDER) # 年データDataFrame化

        bottom = np.zeros(len(months)) # 積み上げ底辺初期化
        # 少しずつ色を変えるためalphaを用いる
        alpha = 0.9 + (0.1 * i / (len(years) - 1)) if len(years) > 1 else 1.0 
        for category in COLUMN_ORDER: # 各時間帯に対して
            color = tuple(c * alpha for c in base_colors[category]) 
            bars = ax.bar(x + i * width, year_df[category].values, width, bottom=bottom, color=color) 
            for j, rect in enumerate(bars): # 各棒グラフに年・月情報を付加
                rect._year = year 
                rect._month = months[j]
            bottom += year_df[category].values # 積み上げ底辺更新

    # 軸・ラベル設定
    ax.set_ylabel('消費電力量 (kWh)') # 左y軸ラベル
    ax.set_title('年間（月別）電力使用量 (年比較)') # グラフタイトル
    ax.set_xticks(x + width * (len(years) - 1) / 2) # x軸目盛位置設定
    ax.set_xticklabels([f'{m}月' for m in months]) # x軸目盛ラベル設定
    ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7) # y軸グリッド表示
    
    ax.yaxis.set_major_locator(ticker.MultipleLocator(100)) # y軸主目盛間隔設定

    # 凡例（時間帯） - 軸内に表示して常に全体が見えるようにする
    
    legend_patches = [Patch(color=TIME_BANDS[c]['color'], label=TIME_BANDS[c]['label']) for c in DISPLAY_ORDER] # 時間帯凡例パッチ作成
    ax.legend(handles=legend_patches, loc='lower left') # 凡例表示

    # 右軸: 年別の月次コストを年ごとに色分けしてプロット
    ax2 = ax.twinx() # 右y軸作成
    cmap = plt.get_cmap('tab10') # カラーマップ取得
    for i, year in enumerate(years): 
        year_costs = [] # 年別月次コストリスト初期化
        for m in months: 
            ts = pd.Timestamp(f'{year}-{m:02d}-01') 
            if ts in plot_data.index: 
                row = plot_data.loc[ts] 
                c = compute_cost_breakdown(ts, row.get('day', 0.0), row.get('home', 0.0), row.get('night', 0.0)).get('total_cost', 0.0) 
            else:
                c = 0.0
            year_costs.append(c) 
        color = cmap(i % 10) # 色取得
        line = ax2.plot(x + i * width, year_costs, color=color, marker='o', linewidth=0.5, label=f'{year}年 電気料金')[0] # 折れ線グラフ描画
        line._year = year  # 年情報を線に付加
    ax2.legend(loc='lower right') # 右y軸凡例表示

    ax2.set_ylabel('電気料金 (円)') # 右y軸ラベル
    # 0ラインを完全一致させる
    left_ylim = ax.get_ylim() # 左y軸の表示範囲取得
    min_ylim = min(0, left_ylim[0]) # 最小値設定
    max_ylim = left_ylim[1]+5 # 最大値設定（少し余裕を持たせる）
    ax.set_ylim(min_ylim, max_ylim) # 左y軸の表示範囲設定
    ax2.set_ylim(min_ylim*40, max_ylim*40) # 右y軸表示範囲設定

    # 年別集計情報を表示
    year_totals = {} # 年別集計初期化
    for year in years: 
        year_data = {'day': 0.0, 'home': 0.0, 'night': 0.0, 'cost': 0.0} # 年別データ初期化
        monthly_costs = [] # 月ごとのコストリスト
        for month in months: 
            try:
                ts = pd.Timestamp(f'{year}-{month:02d}-01') 
                if ts in plot_data.index: 
                    row = plot_data.loc[ts] 
                    year_data['day'] += float(row.get('day', 0.0)) 
                    year_data['home'] += float(row.get('home', 0.0)) 
                    year_data['night'] += float(row.get('night', 0.0)) 
                    monthly_cost = compute_cost_breakdown(ts, row.get('day', 0.0), row.get('home', 0.0), row.get('night', 0.0)).get('total_cost', 0.0)
                    monthly_costs.append(monthly_cost)
                else:
                    monthly_costs.append(0.0)
            except Exception:
                monthly_costs.append(0.0)
                continue
        year_data['cost'] = sum(monthly_costs)
        year_totals[year] = year_data 

    # 年別集計情報のテキストを作成
    years_text = "年間集計:\n\n" 
    for year in years: 
        total = year_totals[year] # 年別集計データ取得
        total_kwh = total['day'] + total['home'] + total['night'] # 年間総使用電力量計算
        years_text += (f"{year}年\n"
                      f"デイタイム電力　 {total['day']:.2f} kWh\n"
                      f"ホームタイム電力 {total['home']:.2f} kWh\n"
                      f"ナイトタイム電力 {total['night']:.2f} kWh\n"
                      f"年間総使用電力量 {total_kwh:.2f} kWh\n"
                      f"年間合計金額　　 {total['cost']:.0f} 円\n\n")

    # 年別集計情報をテーブル化するデータを作成
    col_labels = ['年', 'デイ (kWh)', 'ホーム (kWh)', 'ナイト (kWh)', '総計 (kWh)', '年間金額 (円)'] # 列ラベル
    table_rows = [] # テーブル行リスト初期化
    for year in years:
        t = year_totals[year]
        total_kwh = t['day'] + t['home'] + t['night'] # 年間総使用電力量計算
        table_rows.append([
            str(year),
            f"{t['day']:.2f}",
            f"{t['home']:.2f}",
            f"{t['night']:.2f}",
            f"{total_kwh:.2f}",
            f"{t['cost']:.0f}"
        ])

    # プロットをTkウィンドウに埋め込み、下にスクロール可能な表を配置する
    win = tk.Tk()
     # ウインドウタイトルをファイルパスに設定
    if file_path is not None:
        try:
            win.wm_title(str(file_path))
        except Exception:
            pass
    # Canvas for matplotlib figure
    widget = None
    canvas = FigureCanvasTkAgg(fig, master=win)
    canvas.draw()
    widget = canvas.get_tk_widget()
    widget.pack(side='top', fill='both', expand=1)

    # Frame for table with scrollbar
    table_frame = tk.Frame(win)
    table_frame.pack(side='bottom', fill='x')

    tree_frame = tk.Frame(table_frame)
    tree_frame.pack(side='left', fill='both', expand=True)

    vsb = tk.Scrollbar(table_frame, orient='vertical')
    vsb.pack(side='right', fill='y')

    tree = ttk.Treeview(tree_frame, columns=[f'col{i}' for i in range(len(col_labels))], show='headings', yscrollcommand=vsb.set, height=min(max(3, len(table_rows)), 10))
    for i, h in enumerate(col_labels):
        tree.heading(f'col{i}', text=h)
        tree.column(f'col{i}', anchor='center', width=120)

    for row in table_rows:
        tree.insert('', 'end', values=row)

    tree.pack(side='left', fill='both', expand=True)
    vsb.config(command=tree.yview)

    # ウインドウクローズ時の処理
    def on_closing():
        try:
            plt.close('all') # すべてのグラフ閉じる
        except Exception:
            pass
        try:
            win.destroy() # ウインドウ破棄
        except Exception:
            pass
        try:
            sys.exit(0) # 終了
        except Exception:
            pass

    win.protocol('WM_DELETE_WINDOW', on_closing) # ウインドウクローズイベント設定

    annotation = None # ホバー注釈
    fade_job = None # フェードタイマーID

    # アノテーション削除関数
    def remove_annotation():
        nonlocal annotation, fade_job, widget
        # Clean up annotation
        if annotation is not None:
            try:
                annotation.remove()
            except Exception:
                pass
            annotation = None
        # Cancel fade timer if exists
        if fade_job is not None and widget is not None and widget.winfo_exists():
            try:
                widget.after_cancel(fade_job)
            except Exception:
                pass
            fade_job = None
        # Redraw if canvas exists
        try:
            if hasattr(fig, 'canvas') and fig.canvas is not None:
                fig.canvas.draw_idle()
        except Exception:
            pass

    # グラフ保存ダイアログ関数
    def save_via_dialog():
        try:
            title = getattr(ax, 'get_title', lambda: "graph")()
            fname = filedialog.asksaveasfilename(parent=win, defaultextension='.png', filetypes=[('PNG image','*.png')], initialfile=title if title else "graph")
            if fname:
                fig.savefig(fname, bbox_inches='tight')
                print(f'Saved figure to {fname}')
        except Exception as e:
            print('Save failed:', e)

    # コンテキストメニュー表示関数
    def show_context_menu(event):
        try:
            menu = tk.Menu(widget, tearoff=0)
            menu.add_command(label='グラフ保存', command=save_via_dialog)
            # 年間ウインドウにもDessmonitorメニューを追加
            menu.add_command(label='Dessmonitorデータ反映', command=lambda: plot_monthly_dessmonitor(monthly, df, file_path, dessmonitor_folder))
            try:
                menu.tk_popup(widget.winfo_pointerx(), widget.winfo_pointery())
            finally:
                menu.grab_release()
        except Exception:
            pass

    # マウス移動イベントハンドラ
    def on_motion(event):
        nonlocal annotation
        if getattr(event, 'inaxes', None) is None or not hasattr(event.inaxes, 'patches'):
            remove_annotation()
            return
        for rect in getattr(ax, 'patches', []):
            try:
                contains, _ = rect.contains(event)
            except Exception:
                contains = False
            if contains:
                year = getattr(rect, '_year', None)
                month = getattr(rect, '_month', None)
                if year is None or month is None:
                    continue
                try:
                    idx = pd.Timestamp(f'{year}-{int(month):02d}-01')
                    row = monthly.loc[f'{year}-{int(month):02d}']
                except Exception:
                    row = None
                if row is None:
                    continue
                cost = compute_cost_breakdown(idx, row.get('day',0.0), row.get('home',0.0), row.get('night',0.0))
                total_cost = cost.get('total_cost', 0.0) 
                msg = (
                    f'{year}年{int(month)}月\n'
                    f'デイタイム: {row.get("day",0.0):.2f}kWh\nホームタイム: {row.get("home",0.0):.2f}kWh\nナイトタイム: {row.get("night",0.0):.2f}kWh\n'
                    f'合計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: {total_cost:.0f}円'
                )
                if annotation is not None:
                    try:
                        annotation.remove()
                    except Exception:
                        pass
                x = rect.get_x() + rect.get_width()/2
                y = rect.get_y() + rect.get_height()
                annotation = annotate_in_axes(ax, fig, x, y, msg)
                return
        remove_annotation()

    # マウスクリックイベントハンドラ
    def on_button(event): # マウスクリックイベントハンドラ
        # ダブルクリック -> 日別グラフ表示
        if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1: # 左ダブルクリック
            for rect in ax.patches:
                try:
                    contains, _ = rect.contains(event)
                except Exception:
                    contains = False
                if contains: # 対象の棒グラフが見つかった場合
                    year = getattr(rect, '_year', None) # 年取得
                    month = getattr(rect, '_month', None) # 月取得
                    if year is not None and month is not None: # 年月が有効な場合
                        year_month = f'{year}-{int(month):02d}'
                        plot_daily(df, year_month, file_path=file_path, dessmonitor_folder=dessmonitor_folder)
                        return
        # 右クリック -> コンテキストメニュー表示
        if getattr(event, 'button', None) == 3: # 右クリック
            show_context_menu(event) # コンテキストメニュー表示

    try:
        fig.canvas.mpl_connect('motion_notify_event', on_motion)
        fig.canvas.mpl_connect('figure_leave_event', lambda ev: remove_annotation())
        fig.canvas.mpl_connect('button_press_event', on_button)
    except Exception:
        def on_click(event): # マウスクリックイベントハンドラ（代替）
            nonlocal annotation
            bar_clicked = False
            for rect in ax.patches:
                try:
                    contains, _ = rect.contains(event)
                except Exception:
                    contains = False
                if contains:
                    year = getattr(rect, '_year', None)
                    month = getattr(rect, '_month', None)
                    if year is not None and month is not None:
                        year_month = f'{year}-{int(month):02d}'
                        if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1:
                            plot_daily(df, year_month, file_path=file_path)
                        else:
                            try:
                                idx = pd.Timestamp(f'{year}-{int(month):02d}-01')
                                row = monthly.loc[f'{year}-{int(month):02d}']
                            except Exception:
                                row = None
                            if row is not None:
                                cost = compute_cost_breakdown(idx, row.get('day',0.0), row.get('home',0.0), row.get('night',0.0))
                                msg = (
                                    f'{year}年{int(month)}月\n'
                                    f'デイタイム: {row.get("day",0.0):.2f}kWh\nホームタイム: {row.get("home",0.0):.2f}kWh\nナイトタイム: {row.get("night",0.0):.2f}kWh\n'
                                    f'合計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: {cost:.0f}円'
                                )
                                if annotation is not None:
                                    annotation.remove()
                                x = rect.get_x() + rect.get_width()/2
                                y = rect.get_y() + rect.get_height()
                                try:
                                    ylim = ax.get_ylim()
                                    y = min(y + (ylim[1]-ylim[0]) * 0.03, ylim[1] * 0.95)
                                except Exception:
                                    pass
                                annotation = ax.annotate(msg, xy=(x, y), xytext=(0, 0), xycoords='data', textcoords='offset points',
                                    ha='center', va='bottom', fontsize=10, bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.9), zorder=1000)
                                annotation.set_clip_on(True)
                                fig.canvas.draw_idle()
                        bar_clicked = True
                    break
            if not bar_clicked and annotation is not None:
                try:
                    annotation.remove()
                except Exception:
                    pass
                annotation = None
                try:
                    fig.canvas.draw_idle()
                except Exception:
                    pass
        fig.canvas.mpl_connect('button_press_event', on_click)
    # 右クリックメニュー（グラフ保存、Dessmonitor反映）
    global dessmonitor_data_cache_global # Dessmonitorデータキャッシュ
    if dessmonitor_data_cache_global is None: # 初回読み込み
        dessmonitor_data_cache_global = build_dessmonitor_data_cache(dessmonitor_folder) # Dessmonitorデータキャッシュを構築
    dessmonitor_data_dict = dessmonitor_data_cache_global # ローカル参照
    
    try:
        canvas.draw()
    except Exception:
        pass
    win.mainloop()

# ファイル選択ダイアログを表示してCSVファイルを選択
def select_file() -> Path:
    # GUIのメインウィンドウを作成（表示はしない）
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを非表示

    # ファイル選択ダイアログを表示
    file_path = filedialog.askopenfilename(
        title='電力会社からダウンロードしたCSVファイルを選択してください',
        filetypes=[('CSVファイル', '*.csv'), ('すべてのファイル', '*.*')],
        initialdir=str(Path.cwd())  # カレントディレクトリを初期フォルダに
    )
    
    if not file_path:  # キャンセルされた場合
        print('ファイルが選択されませんでした。')
        sys.exit(1) # 終了
    
    return Path(file_path)

# メイン処理
def main():
    # コマンドライン引数があればそれを使用、なければファイル選択ダイアログを表示
    if len(sys.argv) > 1:
        path = Path(sys.argv[1]) 
    else:
        path = select_file()

    if not path.exists():
        print('File not found:', path)
        sys.exit(1) # 終了

    print(f'処理するファイル: {path}')
    # Dessmonitorフォルダー選択
    
    root = tk.Tk()
    root.withdraw() # メインウィンドウを非表示
    dessmonitor_folder = filedialog.askdirectory(parent=root, title='Dessmonitorデータのフォルダーを選択してください')
    if not dessmonitor_folder: # キャンセルされた場合
        print('Dessmonitorフォルダーが選択されませんでした。')
        sys.exit(1) # 終了

    df = load_csv(path) # CSVファイル読み込み
    monthly, df_all = parse_and_aggregate(df) # 月別集計データ作成
    plot_monthly_interactive(monthly, df_all, file_path=path, dessmonitor_folder=dessmonitor_folder) # 月別グラフ描画

if __name__ == '__main__': 
    main() 
