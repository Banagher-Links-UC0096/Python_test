# -*- coding: utf-8 -*-
"""
CSVフォーマット想定:
1行目: ヘッダー
2行目以降: 1列目=日付(yyyymmdd), 8列目=日合計(kWh), 10列目=0時-1時, 11列目=1時-2時, ...

スクリプトは月毎に各時間帯(デイ/ホーム/ナイト)の合計を積み上げ棒グラフで出力します。
使い方: python smartmeter_monthly_stack.py path/to/file.csv

出力: ./monthly_stack.png
"""
import sys
from pathlib import Path
import pandas as pd
import matplotlib
import numpy as np
# matplotlib.use('Agg')  # ← ウインドウ表示のためコメントアウト
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter.ttk as ttk
import tkinter as tk
from tkinter import filedialog
# matplotlib font manager not required explicitly here
import jpholiday
import datetime
import time

# 日本語フォントの設定
plt.rcParams['font.family'] = 'MS Gothic'  # Windows の場合

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
}

# ---- 共通の定義 ----
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
    'cost_color': 'red',
    'cost_marker': 'o',
    'cost_linewidth': 0.5,
    'grid_color': 'gray',
    'grid_alpha': 0.7,
    'grid_linestyle': '--',
    'annotation_bgcolor': 'yellow',
    'annotation_alpha': 0.9,
    'summary_bgcolor': 'white',
    'summary_alpha': 0.85
}

# 列の順序定義
COLUMN_ORDER = ['night', 'home', 'day']  # 積み上げの順序（下から）
DISPLAY_ORDER = ['day', 'home', 'night']  # 表示の順序（凡例など）

# 但し上のrangeは"時間"ではなく、CSVの時間カラムインデックスを計算する関係で調整が必要
# 実際にはCSVの10列目が0時-1時のため、0時カラムの列インデックス = 9


def is_holiday(date):
    """土日祝日判定"""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    # 土日判定
    if date.weekday() >= 5:  # 5=土曜日, 6=日曜日
        return True
    # 祝日判定
    return jpholiday.is_holiday(date)

def get_unit_prices_for_date(date):
    """指定日のデイ/ホーム/ナイト単価を返す（円/kWh）。"""
    if pd.isna(date):
        # default to latest
        return PRICE_PERIODS[-1][2]
    if isinstance(date, (pd.Timestamp, datetime.date)):
        dt = pd.Timestamp(date)
    else:
        dt = pd.to_datetime(date)
    for start, end, prices in PRICE_PERIODS:
        if (start is None or dt >= start) and (end is None or dt <= end):
            return prices
    return PRICE_PERIODS[-1][2]

def get_renewable_unit_for_date(date):
    """再エネ賦課金単価を返す。毎年5月に切替: 例) 2023年5月～2024年4月は RENEWABLE_BY_YEAR[2023]"""
    if isinstance(date, (pd.Timestamp, datetime.date)):
        dt = pd.Timestamp(date)
    else:
        dt = pd.to_datetime(date)
    year = dt.year
    # if month < 5, use previous year's value (e.g., 2024-04 uses 2023 value)
    key_year = year if dt.month >= 5 else year - 1
    return RENEWABLE_BY_YEAR.get(key_year, 0.0)

def get_fuel_adj_for_date(date):
    if isinstance(date, (pd.Timestamp, datetime.date)):
        dt = pd.Timestamp(date)
    else:
        dt = pd.to_datetime(date)
    key = f"{dt.year:04d}-{dt.month:02d}"
    return FUEL_ADJ.get(key, 0.0)

def compute_cost_from_parts(date, day_kwh, home_kwh, night_kwh):
    """日単位のエネルギー構成からその日の電気料金（円）を計算する。"""
    prices = get_unit_prices_for_date(date)
    renew = get_renewable_unit_for_date(date)
    fuel = get_fuel_adj_for_date(date)
    total = float(day_kwh + home_kwh + night_kwh)
    cost = day_kwh * prices['day'] + home_kwh * prices['home'] + night_kwh * prices['night']
    # 再エネと燃料調整は総使用量に対して加算
    cost += total * (renew + fuel)
    return cost


def compute_cost_breakdown(date, day_kwh, home_kwh, night_kwh):
    """日単位のエネルギー構成から時間帯別の金額、再エネ・燃料調整額、合計を詳細に返す。

    戻り値: dict {
        'prices': {..}, 'renew': float, 'fuel': float,
        'base_costs': {'day':..,'home':..,'night':..}, 'base_total':..,
        'renew_amount':.., 'fuel_amount':.., 'total_kwh':.., 'total_cost':..
    }
    """
    prices = get_unit_prices_for_date(date)
    renew = get_renewable_unit_for_date(date)
    fuel = get_fuel_adj_for_date(date)
    day_kwh = float(day_kwh or 0.0)
    home_kwh = float(home_kwh or 0.0)
    night_kwh = float(night_kwh or 0.0)
    base_costs = {
        'day': day_kwh * prices.get('day', 0.0),
        'home': home_kwh * prices.get('home', 0.0),
        'night': night_kwh * prices.get('night', 0.0)
    }
    base_total = sum(base_costs.values())
    total_kwh = day_kwh + home_kwh + night_kwh
    renew_amount = total_kwh * renew
    fuel_amount = total_kwh * fuel
    total_cost = base_total + renew_amount + fuel_amount
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

# Timestamp of last figure creation to avoid cascade save on same click
LAST_FIG_OPEN = 0.0


# ---- 共通ユーティリティ関数 ----
def hour_to_band(hour: int, is_holiday_flag: bool) -> str:
    """時間 (0-23) を受け取り、休日フラグに応じて 'day'/'home'/'night' を返す。"""
    try:
        h = int(hour) % 24
    except Exception:
        return 'night'
    if is_holiday_flag:
        return 'home' if 7 <= h <= 22 else 'night'
    # 平日
    if 9 <= h <= 16:
        return 'day'
    if 7 <= h <= 8 or 17 <= h <= 22:
        return 'home'
    return 'night'


def build_hourly_unit(prices: dict, renew: float, fuel: float, is_holiday_flag: bool) -> list:
    """単日 (24要素) の時間ごとの単価リストを返す (円/kWh)。"""
    return [prices[hour_to_band(h, is_holiday_flag)] + renew + fuel for h in range(24)]


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


def set_ax2_min_zero(ax2):
    """右軸 (ax.twinx()) の下限を 0 に設定する試みを行う。"""
    try:
        ymax = ax2.get_ylim()[1]
        ax2.set_ylim(0, max(ymax, 0.1))
    except Exception:
        try:
            ax2.set_ylim(bottom=0)
        except Exception:
            pass

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


def parse_and_aggregate(df: pd.DataFrame) -> pd.DataFrame:
    # 前提: 1列目が日付(yyyymmdd)
    # 時間帯列は10列目(0時)から33列目(23時)のはず -> pandasカラム index 9..32
    # カラム名がない場合でも位置で取り出す
    if df.shape[1] < 33:
        raise ValueError('CSVに十分な列がありません。24時間分の列が必要です。')

    # 日付列
    date_col = df.columns[0]
    df['date'] = pd.to_datetime(df[date_col], format='%Y%m%d', errors='coerce')
    if df['date'].isna().any():
        print('警告: 日付のパースに失敗した行があります。')

    # 時間帯ごとの集計を作る
    # 0時カラムの位置
    start_hour_col_idx = 9  # 0時の列インデックス(0-origin)

    # build numeric hourly columns
    hourly_cols = df.columns[start_hour_col_idx:start_hour_col_idx+24]
    # 文字列を数値に
    df_hourly = df[hourly_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

    # Map hours to categories
    # column 0 -> 0時, column 1 -> 1時 ...
    hour_to_col = {h: hourly_cols[h] for h in range(24)}

    def sum_for_category(row, hours):
        cols = [hour_to_col[h] for h in hours]
        return row[cols].sum()

    # 時間帯定義 (hours numbers)
    # 通常の平日定義
    default_day_hours = list(range(9, 17))  # 9～17時 inclusive (9..17)
    default_home_hours = list(range(7, 9)) + list(range(17, 23))  # 7-8,17-22
    night_hours = list(range(23, 24)) + list(range(0, 7))  # 23,0-6

    # 各行ごとに日付を見て土日祝なら9-17をホームタイムにする
    # 結果を格納するカラムを初期化
    df['day'] = 0.0
    df['home'] = 0.0
    df['night'] = 0.0

    # iterate rows by index to access date and hourly values together
    # (hour_to_col keys available when needed) 不要な代替変数を削除
    for idx in df.index:
        date_val = df.at[idx, 'date']
        # If date parsing failed, treat as weekday (conservative)
        try:
            is_hol = is_holiday(date_val)
        except Exception:
            is_hol = False

        if is_hol:
            # 土日祝日は9-17をホームタイム扱い -> ホームは7-22、デイは無し
            day_h = []
            home_h = list(range(7, 23))  # 7..22
        else:
            day_h = default_day_hours
            home_h = default_home_hours

        # sum values from df_hourly for this row
        try:
            df.at[idx, 'day'] = df_hourly.loc[idx, [hour_to_col[h] for h in day_h]].sum() if day_h else 0.0
            df.at[idx, 'home'] = df_hourly.loc[idx, [hour_to_col[h] for h in home_h]].sum() if home_h else 0.0
            df.at[idx, 'night'] = df_hourly.loc[idx, [hour_to_col[h] for h in night_hours]].sum()
        except Exception:
            # fallback to zero if something unexpected occurs
            df.at[idx, 'day'] = 0.0
            df.at[idx, 'home'] = 0.0
            try:
                df.at[idx, 'night'] = df_hourly.loc[idx, [hour_to_col[h] for h in night_hours]].sum()
            except Exception:
                df.at[idx, 'night'] = 0.0

    # 月単位で集計
    df['year_month'] = df['date'].dt.to_period('M')
    monthly = df.groupby('year_month')[['day', 'home', 'night']].sum()
    monthly = monthly.sort_index()
    # Convert index to string for plotting
    monthly.index = monthly.index.astype(str)
    # 日ごとの集計も返す
    return monthly, df



# 時間毎のグラフを表示する関数
def plot_hourly(df: pd.DataFrame, year: int, month: int, day: int, file_path: Path = None, dessmonitor_folder: str = None):
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
    # embed into a Tk Toplevel (attach to existing root if present)
    master = tk._default_root if tk._default_root is not None else None
    if master is None:
        win = tk.Tk()
    else:
        win = tk.Toplevel(master)
    # embed matplotlib figure into the Toplevel
    widget = None
    try:
        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        widget = canvas.get_tk_widget()
        widget.pack(side='top', fill='both', expand=1)
    except Exception:
        # fallback: continue and let matplotlib manage the window
        pass
    # ウインドウタイトルをファイルパスに設定
    if file_path is not None:
        try:
            win.wm_title(str(file_path))
        except Exception:
            pass
    bars = ax.bar(range(24), values, color=colors)
    ax.set_ylabel('消費電力量 (kWh)')
    ax.set_title(f'{year}-{month:02d}-{day:02d} 時間帯別電力使用量')
    ax.set_xticks(range(24))
    # 凡例
    from matplotlib.patches import Patch
    legend_patches = [
        Patch(color=TIME_BANDS[band]['color'], label=TIME_BANDS[band]['label'])
        for band in DISPLAY_ORDER
    ]
    ax.legend(handles=legend_patches, bbox_to_anchor=(1.05, 1), loc='upper left')
    # 0.5kWh単位でグリッド線
    ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
    import matplotlib.ticker as ticker
    ax.yaxis.set_major_locator(ticker.MultipleLocator(0.5))
    # 右軸に料金をプロットする: 時間ごとの単価は日付基準で同じと仮定
    # 各時間のコストは時間ごとのkWh * 該当時間帯単価 + kWh * (再エネ + 燃料)
    # 再エネ/燃料は日次合計に対して加算するため、ここでは時間ごとに按分して表示
    # 日次の合計再エネ+燃料単価
    renew = get_renewable_unit_for_date(target_date)
    fuel = get_fuel_adj_for_date(target_date)
    prices = get_unit_prices_for_date(target_date)

    # 時間ごとの単価を作る (共通処理を使用)
    hourly_unit = build_hourly_unit(prices, renew, fuel, is_hol)

    hourly_costs = values * np.array(hourly_unit)
    ax2 = ax.twinx()
    ax2.plot(range(24), hourly_costs, color='red', marker='o', linewidth=0.5, label='電気料金 (円)')
    ax2.set_ylabel('電気料金 (円)')
    ax2.yaxis.grid(False)
    # 左右y軸の0ラインを完全一致させる
    left_ylim = ax.get_ylim()
    min_ylim = min(0, left_ylim[0])
    max_ylim = left_ylim[1]+0.2  # 少し余裕を持たせる
    ax.set_ylim(min_ylim, max_ylim)
    ax2.set_ylim(min_ylim*50, max_ylim*50)  # 倍率
    ax2.legend(loc='upper right')
    # --- 右側の凡例下に料金・集計表示を追加 ---
    # 再エネ/燃料と時間帯単価、各時間帯合計と金額、総計を計算（共通化関数を利用）
    band_sums = band_sums_from_values(values, is_hol)
    bd = compute_cost_breakdown(target_date, band_sums['day'], band_sums['home'], band_sums['night'])
    base_costs = bd['base_costs']
    base_total = bd['base_total']
    renew_amount = bd['renew_amount']
    fuel_amount = bd['fuel_amount']
    total_cost = bd['total_cost']
    prices = bd['prices']
    renew = bd['renew']
    fuel = bd['fuel']
    total_kwh = bd.get('total_kwh', float(np.nansum(values)))

    unit_text = (f"適用単価:\n"
                 f"デイタイム単価　 {prices['day']:.2f} 円/kWh\n"
                 f"ホームタイム単価 {prices['home']:.2f} 円/kWh\n"
                 f"ナイトタイム単価 {prices['night']:.2f} 円/kWh\n"
                 f"再エネ賦課金単価 {renew:.2f} 円/kWh\n"
                 f"燃料費調整単価　 {fuel:.2f} 円/kWh")

    band_lines = (f"使用電力量:\n"
                  f"デイタイム電力　 {band_sums['day']:.2f} kWh\n"
                  f"ホームタイム電力 {band_sums['home']:.2f} kWh\n"
                  f"ナイトタイム電力 {band_sums['night']:.2f} kWh\n"
                  f"総電力量　　　　 {total_kwh:.2f} kWh")
    
    # 買電金額（使用電力量合計に対する金額）
    buy_total = float(np.nansum(values)) * (prices['day'] + prices['home'] + prices['night'] + renew + fuel) / 3  # 平均単価で合計
    totals = (f"集計金額:\n"
              f"デイタイム金額　 {base_costs['day']:.0f} 円\n"
              f"ホームタイム金額 {base_costs['home']:.0f} 円\n"
              f"ナイトタイム金額 {base_costs['night']:.0f} 円\n"
              f"再エネ賦課金額　 {renew_amount:.2f} 円\n"
              f"燃料調整費金額　 {fuel_amount:.2f} 円\n"
              f"合計金額　　　　 {total_cost:.0f} 円"
              )

    summary_text = unit_text + "\n\n" + band_lines + "\n\n" + totals
    # place the summary inside the axes (top-right) so it does not overflow the window
    ax.text(1.4, 0.75, summary_text, transform=ax.transAxes, ha='right', va='top', fontsize=9,
        bbox=dict(boxstyle='round,pad=0.4', fc='white', ec='black', alpha=0.85))
    # background-click save handler removed (deleted to simplify code)
    plt.tight_layout()
    # Hover-based annotation + left-double-click open + right-click context menu (Save)
    annotation = None
    fade_job = None

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

    def schedule_fade(timeout_ms: int = 2500):
        nonlocal fade_job
        if widget is None:
            return
        try:
            if fade_job is not None:
                widget.after_cancel(fade_job)
        except Exception:
            pass
        try:
            fade_job = widget.after(timeout_ms, remove_annotation)
        except Exception:
            fade_job = None

    def save_via_dialog():
        # prompt user for filename
        try:
            title = getattr(ax, 'get_title', lambda: "graph")()
            fname = filedialog.asksaveasfilename(parent=win, defaultextension='.png', filetypes=[('PNG image', '*.png')], initialfile=title if title else "graph")
            if fname:
                fig.savefig(fname, bbox_inches='tight')
                print(f'Saved figure to {fname}')
        except Exception as e:
            print('Save failed:', e)

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
     
    def on_motion(event):
        nonlocal annotation
        # check if we're hovering over any data elements
        if getattr(event, 'inaxes', None) is None or not hasattr(event.inaxes, 'containers'):
            remove_annotation()
            return
        # find the bar under cursor
        for cont in ax.containers:
            for rect in cont:
                try:
                    contains = rect.contains(event)[0]
                except Exception:
                    contains = False
                if contains:
                    # show annotation for this rect (replace existing)
                    try:
                        xcenter = rect.get_x() + rect.get_width() / 2.0
                        hour = int(xcenter + 0.5)
                    except Exception:
                        hour = int(rect.get_x())
                    value = rect.get_height()
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
                    x = rect.get_x() + rect.get_width() / 2
                    y = rect.get_y() + rect.get_height()
                    annotation = annotate_in_axes(ax, fig, x, y, msg)
                    schedule_fade()
                    return
        # not on any bar
        remove_annotation()

    def on_button(event):
        # left double-click opens deeper view? hourly is deepest -> no-op
        if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1:
            # keep behavior: nothing to open for hourly
            return
        # right-click -> show context menu
        if getattr(event, 'button', None) == 3:
            show_context_menu(event)

    # Connect handlers with error checking
    try:
        canvas = getattr(fig, 'canvas', None)
        if canvas:
            canvas.mpl_connect('motion_notify_event', on_motion)
            canvas.mpl_connect('figure_leave_event', lambda ev: remove_annotation())
            canvas.mpl_connect('button_press_event', on_button)
    except Exception:
        # fallback to click-based handler if motion events not available
        def on_click(event):
            # keep previous behavior as fallback
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

    # closing handler for the Toplevel
    def _on_close():
        try:
            plt.close(fig)
        except Exception:
            pass
        try:
            win.destroy()
        except Exception:
            pass
        # if there is no main Tk window left, exit
        try:
            if not tk._default_root:
                sys.exit(0)
        except Exception:
            pass

    try:
        win.protocol('WM_DELETE_WINDOW', _on_close)
    except Exception:
        pass

    # if we created a standalone Tk, start its mainloop; otherwise drawing is enough
    if master is None:
        try:
            win.mainloop()
        except Exception:
            pass


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
        import tkinter as tk
        from tkinter import messagebox
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        import numpy as np
        import glob, os, re
        # フォルダーはdessmonitor_folder引数を使用
        folder = dessmonitor_folder
        if not folder:
            messagebox.showerror('エラー', 'Dessmonitorフォルダーが指定されていません')
            return
        # グラフの日付取得
        date_str = target_date.strftime('%Y-%m-%d')
        pattern = os.path.join(folder, f"**/energy-storage-container-*-{date_str}.xlsx")
        files = glob.glob(pattern, recursive=True)
        if not files:
            messagebox.showinfo('ファイルなし', f'{date_str} のDessmonitorファイルが見つかりません')
            return
        # 複数ID分まとめてデータ読込
        all_data = []
        for f in files:
            try:
                import pandas as pd
                df = pd.read_excel(f)
                all_data.append(df)
            except Exception as e:
                print(f'ファイル読込失敗: {f}', e)
        if not all_data:
            messagebox.showerror('エラー', 'Dessmonitorファイルの読込に失敗しました')
            win.destroy()
            return
        print(f'Dessmonitorファイル読込: {len(all_data)}件')
        # 24時間分の合計値（全ID合算）
        dess_values = np.zeros(24)
        commercial_values = np.zeros(24)
        for df in all_data:
            try:
                times = df.iloc[:,0]
                used = df.iloc[:,30]
            except Exception as e:
                print('Dessmonitor列抽出失敗:', e)
                continue
            df_sorted = df.iloc[::-1]
            df_sorted['hour'] = df_sorted.iloc[:,0].apply(lambda x: int(str(x)[11:13]) if isinstance(x,str) and len(x)>=13 else np.nan)
            for h in range(24):
                group = df_sorted[df_sorted['hour']==h]
                if len(group) == 0:
                    continue
                vals = group.iloc[:,30].values
                vals_c = group.iloc[:,31].values if group.shape[1]>31 else []
                # 蓄電
                if len(vals) >= 2:
                    if vals[0] > vals[-1]:
                        dess_values[h] += 0.0
                    else:
                        dess_values[h] += float(vals[-1]) - float(vals[0])
                # 商用充電
                if len(vals_c) >= 2:
                    if vals_c[0] > vals_c[-1]:
                        commercial_values[h] += 0.0
                    else:
                        commercial_values[h] += float(vals_c[-1]) - float(vals_c[0])
        # プラス側：電力会社データそのまま
        new_values = values
        # マイナス側：Dessmonitorデータそのまま負値で
        dess_values_plot = -dess_values
        # 商用充電はプラス側で幅半分、中心揃え
        bar_width = 0.8
        commercial_bar_width = bar_width / 2
        # --- グラフ表示（plot_hourlyと同じ構成） ---
        fig2, ax2 = plt.subplots(figsize=(10,5))
        win2 = tk.Toplevel(win)
        win2.wm_title(f'{date_str} 時間帯別電力使用量 (Dessmonitor反映)')
        canvas2 = FigureCanvasTkAgg(fig2, master=win2)
        canvas2.draw()
        widget2 = canvas2.get_tk_widget()
        widget2.pack(side='top', fill='both', expand=1)
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

        if widget2 is not None:
            widget2.bind('<Button-3>', show_context_menu2)
            bars2 = ax2.bar(range(24), new_values, color=colors, label='電力会社データ', width=bar_width, align='center')
            bars2_dess = ax2.bar(range(24), dess_values_plot, color='deepskyblue', alpha=0.6, label='蓄電池出力', width=bar_width, align='center')
            # 商用充電はプラス側で幅半分、中心揃え
            bars2_commercial = ax2.bar([x + 0.25 for x in range(24)], commercial_values, color='deepskyblue', alpha=0.6, label='商用充電', width=commercial_bar_width, align='center')

            # マウスホバー注釈（棒グラフ上で各時間帯の電力量表示）
            annotation = None
            fade_job = None
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
            def schedule_fade(timeout_ms: int = 2500):
                nonlocal fade_job
                if widget2 is None:
                    return
                try:
                    if fade_job is not None:
                        widget2.after_cancel(fade_job)
                except Exception:
                    pass
                try:
                    fade_job = widget2.after(timeout_ms, remove_annotation)
                except Exception:
                    fade_job = None
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
                                label = '蓄電/商用充電'
                            else:
                                label = '電力会社データ'
                            msg = f'{hour}時-{(hour+1)%24}時\n{label}: {value:.2f}kWh'
                            if annotation is not None:
                                try:
                                    annotation.remove()
                                except Exception:
                                    pass
                            x = rect.get_x() + rect.get_width() / 2
                            y = rect.get_y() + rect.get_height()
                            annotation = annotate_in_axes(ax2, fig2, x, y, msg)
                            schedule_fade()
                            return
                remove_annotation()
            def on_leave(event):
                remove_annotation()
            fig2.canvas.mpl_connect('motion_notify_event', on_motion)
            fig2.canvas.mpl_connect('figure_leave_event', on_leave)
        ax2.set_ylabel('消費電力量 (kWh)')
        ax2.set_title(f'{date_str} 時間帯別電力使用量 (Dessmonitor反映)')
        ax2.set_xticks(range(24))
        ax2.set_xticklabels([str(h) for h in range(24)], rotation=0)
        from matplotlib.patches import Patch
        legend_patches = [Patch(color=TIME_BANDS[band]['color'], label=TIME_BANDS[band]['label']) for band in DISPLAY_ORDER]
        legend_patches += [Patch(color='deepskyblue', label='蓄電池出力')]
        legend_patches += [Patch(color='deepskyblue', label='商用充電')]
        ax2.legend(handles=legend_patches, bbox_to_anchor=(1.05, 1.1), loc='upper left')
        ax2.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
        import matplotlib.ticker as ticker
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
        dess_bd = compute_cost_breakdown(target_date, dess_band_sums['day'], dess_band_sums['home'], dess_band_sums['night'])
        dess_base_costs = dess_bd['base_costs']
        dess_base_total = dess_bd['base_total']
        dess_renew_amount = dess_bd['renew_amount']
        dess_fuel_amount = dess_bd['fuel_amount']
        dess_total_cost = dess_bd['total_cost']
        dess_total_kwh = dess_bd.get('total_kwh', float(np.nansum(dess_values)))
        saving_cost = dess_total_cost
        unit_text = (f"適用単価:\n"
                     f"デイタイム単価　 {prices['day']:.2f} 円/kWh\n"
                     f"ホームタイム単価 {prices['home']:.2f} 円/kWh\n"
                     f"ナイトタイム単価 {prices['night']:.2f} 円/kWh\n"
                     f"再エネ賦課金単価 {renew:.2f} 円/kWh\n"
                     f"燃料費調整単価　 {fuel:.2f} 円/kWh")
        band_lines = (f"使用電力量:\n"
                      f"デイタイム電力　 {band_sums['day']:.2f} kWh\n"
                      f"ホームタイム電力 {band_sums['home']:.2f} kWh\n"
                      f"ナイトタイム電力 {band_sums['night']:.2f} kWh\n"
                      f"買電電力量　　　 {total_kwh:.2f} kWh\n"
                      f"蓄電電力量　　　 {dess_total_kwh:.2f} kWh")
        buy_total = base_costs['day'] + base_costs['home'] + base_costs['night'] + renew_amount + fuel_amount
        total_cost2 = buy_total + saving_cost
        totals = (f"集計金額:\n"
                  f"デイタイム金額　 {base_costs['day']:.0f} 円\n"
                  f"ホームタイム金額 {base_costs['home']:.0f} 円\n"
                  f"ナイトタイム金額 {base_costs['night']:.0f} 円\n"
                  f"再エネ賦課金額　 {renew_amount:.2f} 円\n"
                  f"燃料調整費金額　 {fuel_amount:.2f} 円\n"
                  f"買電金額　　　　 {buy_total:.0f} 円\n"
                  f"節約金額　　　　 {saving_cost:.0f} 円\n"
                  f"太陽光発電無金額 {total_cost2:.0f} 円")
        summary_text = unit_text + "\n\n" + band_lines + "\n\n" + totals
        ax2.text(1.15, 0.7, summary_text, transform=ax2.transAxes,
            ha='left', va='top', fontsize=9,
            bbox=dict(boxstyle='round,pad=0.5', fc='white', ec='black', alpha=0.85))
        plt.tight_layout()
        win2.mainloop()

# 日ごとのグラフを表示する関数
def plot_daily(df: pd.DataFrame, year_month: str, file_path: Path = None, dessmonitor_folder: str = None):
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
    import matplotlib.ticker as ticker
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
    unit_text = (f"適用単価:\n"
                 f"デイタイム単価　 {prices['day']:.2f} 円/kWh\n"
                 f"ホームタイム単価 {prices['home']:.2f} 円/kWh\n"
                 f"ナイトタイム単価 {prices['night']:.2f} 円/kWh\n"
                 f"再エネ賦課金単価 {renew:.2f} 円/kWh\n"
                 f"燃料費調整単価　 {fuel:.2f} 円/kWh")

    usage_text = (f"月間使用電力量:\n"
                  f"デイタイム電力　 {month_total['day']:.2f} kWh\n"
                  f"ホームタイム電力 {month_total['home']:.2f} kWh\n"
                  f"ナイトタイム電力 {month_total['night']:.2f} kWh\n"
                  f"総使用電力量　　 {total_kwh:.2f} kWh")
    
    cost_text = (f"月間金額:\n"
                 f"デイタイム金額　 {month_costs['day']:.0f} 円\n"
                 f"ホームタイム金額 {month_costs['home']:.0f} 円\n"
                 f"ナイトタイム金額 {month_costs['night']:.0f} 円\n"
                 f"再エネ賦課金額　 {month_costs['renew']:.2f} 円\n"
                 f"燃料調整費金額　 {month_costs['fuel']:.2f} 円\n"
                 f"合計金額　　　　 {month_costs['total']:.0f} 円")
    
    summary_text = unit_text + "\n\n" + usage_text + "\n\n" + cost_text
    
    # 凡例の下に表示 -> グラフ内の右上に移動してウインドウからはみ出さないようにする
    ax.text(1.3, 0.83, summary_text, transform=ax.transAxes,
        ha='right', va='top', fontsize=9,
        bbox=dict(boxstyle='round,pad=0.5', fc='white', ec='black', alpha=0.85))
    
    plt.tight_layout()

    # attach day metadata to each bar rectangle using containers (reliable)
    days = daily.index.tolist()
    bar_rects = []
    for cont in getattr(ax, 'containers', []):
        for idx, rect in enumerate(cont):
            if idx < len(days):
                setattr(rect, '_day', int(days[idx]))
                bar_rects.append(rect)

    # ensure x tick labels show day numbers and are not rotated
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

    # Hover-based annotations + left-double-click to open hourly + right-click context menu
    annotation = None
    fade_job = None

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

    def schedule_fade(timeout_ms: int = 2500):
        nonlocal fade_job
        if widget is None:
            return
        try:
            if fade_job is not None:
                widget.after_cancel(fade_job)
        except Exception:
            pass
        try:
            fade_job = widget.after(timeout_ms, remove_annotation)
        except Exception:
            fade_job = None

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

    def show_context_menu(event):
        try:
            menu = tk.Menu(widget, tearoff=0)
            menu.add_command(label='グラフ保存', command=save_via_dialog)
            menu.add_command(label='Dessmonitor', command=lambda: plot_daily_dessmonitor(df, year_month, file_path, dessmonitor_folder))
            try:
                menu.tk_popup(widget.winfo_pointerx(), widget.winfo_pointery())
            finally:
                menu.grab_release()
        except Exception:
            pass
        
    # dessmonitorデータを反映して日別グラフを表示する関数
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
        import glob, os
        import numpy as np
        # Dessmonitorデータ取得: 各日付ごとにファイル検索・集計
        dessmonitor_day_values = {}
        # 商用充電データ: 1日ごとに時間帯別集計
        commercial_output_band = {band: [] for band in COLUMN_ORDER}
        commercial_bar_width = 0.6 / 2  # bar_widthは後で定義されるが0.6でOK
        for day_idx in daily.index:
            # 日付オブジェクト生成
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0]
            date_str = sample_date.strftime('%Y-%m-%d')
            dess_values = np.zeros(24)
            commercial_hourly = np.zeros(24)
            if dessmonitor_folder:
                pattern = os.path.join(dessmonitor_folder, f"**/energy-storage-container-*-{date_str}.xlsx")
                files = glob.glob(pattern, recursive=True)
                for f in files:
                    try:
                        ddf = pd.read_excel(f)
                        # 最終行から早い時刻で昇順
                        ddf_sorted = ddf.iloc[::-1].sort_values(by=ddf.columns[0])
                        ddf_sorted['hour'] = ddf_sorted.iloc[:,0].apply(lambda x: int(str(x)[11:13]) if isinstance(x,str) and len(x)>=13 else np.nan)
                        for h in range(24):
                            group = ddf_sorted[ddf_sorted['hour']==h]
                            # 蓄電
                            vals = group.iloc[:,30].values if len(group)>0 else []
                            if len(vals) >= 2:
                                early = float(vals[0])
                                late = float(vals[-1])
                                if late >= early:
                                    dess_values[h] += late - early
                                else:
                                    dess_values[h] += 0.0
                            # 商用充電
                            vals_c = group.iloc[:,31].values if len(group)>0 and group.shape[1]>31 else []
                            if len(vals_c) >= 2:
                                early_c = float(vals_c[0])
                                late_c = float(vals_c[-1])
                                if late_c >= early_c:
                                    commercial_hourly[h] += late_c - early_c
                                else:
                                    commercial_hourly[h] += 0.0
                    except Exception:
                        continue
            dessmonitor_day_values[day_idx] = dess_values
            # 時間帯ごとに集計
            is_hol = is_holiday(sample_date)
            band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0}
            for h in range(24):
                band = hour_to_band(h, is_hol)
                band_sums[band] += commercial_hourly[h]
            for band in COLUMN_ORDER:
                commercial_output_band[band].append(band_sums[band])
        # Dessmonitor節約金額グラフ化
        saving_costs = []
        for day_idx in daily.index:
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0]
            is_hol = is_holiday(sample_date)
            dess_band_sums = band_sums_from_values(dessmonitor_day_values[day_idx], is_hol)
            dess_bd = compute_cost_breakdown(sample_date, dess_band_sums['day'], dess_band_sums['home'], dess_band_sums['night'])
            saving_costs.append(dess_bd['total_cost'])

        # --- 情報表示用の月間集計 ---
        # 電力会社データ: 1ヶ月間の時間帯別合計
        month_total = daily[['day', 'home', 'night']].sum()
        # 電力会社データ: 1ヶ月間の買電電力量合計
        month_buy_total = month_total.sum()
        # Dessmonitorデータ: 1ヶ月間の蓄電電力量合計
        dess_month_total = {'day': 0.0, 'home': 0.0, 'night': 0.0}
        # 商用充電データ: 1ヶ月間の時間帯別合計
        commercial_month_total = {'day': 0.0, 'home': 0.0, 'night': 0.0}
        for i, day_idx in enumerate(daily.index):
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0]
            is_hol = is_holiday(sample_date)
            dess_values = dessmonitor_day_values.get(day_idx, np.zeros(24))
            for h in range(24):
                band = hour_to_band(h, is_hol)
                dess_month_total[band] += dess_values[h]
            # 商用充電
            for band in COLUMN_ORDER:
                commercial_month_total[band] += commercial_output_band[band][i]
        dess_month_total_sum = sum(dess_month_total.values())
        commercial_month_total_sum = sum(commercial_month_total.values())
        # グラフ描画
        fig, ax = plt.subplots(figsize=(12,7))
        # プラス側: 電力会社データ（積み上げ棒グラフ）
        bar_width = 0.6
        x = np.array(list(daily.index))
        bars = ax.bar(x, daily['night'], width=bar_width, color=TIME_BANDS['night']['color'], label=TIME_BANDS['night']['label'], align='center')
        bars_home = ax.bar(x, daily['home'], width=bar_width, color=TIME_BANDS['home']['color'], bottom=daily['night'], label=TIME_BANDS['home']['label'], align='center')
        bars_day = ax.bar(x, daily['day'], width=bar_width, color=TIME_BANDS['day']['color'], bottom=daily['night']+daily['home'], label=TIME_BANDS['day']['label'], align='center')
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
        commercial_output_band = {band: [] for band in COLUMN_ORDER}
        commercial_bar_width = bar_width / 2  # 電力会社の半分の幅
        import glob, os
        for day_idx in daily.index:
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0]
            is_hol = is_holiday(sample_date)
            # 商用出力データ（energy-storage-container-*-{date}.xlsx）を探す
            commercial_hourly = np.zeros(24)
            if dessmonitor_folder:
                date_str = sample_date.strftime('%Y-%m-%d')
                pattern = os.path.join(dessmonitor_folder, f"**/energy-storage-container-*-{date_str}.xlsx")
                files = glob.glob(pattern, recursive=True)
                for f in files:
                    try:
                        ddf = pd.read_excel(f)
                        ddf_sorted = ddf.iloc[::-1].sort_values(by=ddf.columns[0])
                        ddf_sorted['hour'] = ddf_sorted.iloc[:,0].apply(lambda x: int(str(x)[11:13]) if isinstance(x,str) and len(x)>=13 else np.nan)
                        for h in range(24):
                            group = ddf_sorted[ddf_sorted['hour']==h]
                            vals = group.iloc[:,31].values if len(group)>0 and group.shape[1]>31 else []
                            if len(vals) >= 2:
                                early = float(vals[0])
                                late = float(vals[-1])
                                if late >= early:
                                    commercial_hourly[h] += late - early
                                else:
                                    commercial_hourly[h] += 0.0
                    except Exception:
                        continue
            # 時間帯ごとに集計
            band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0}
            for h in range(24):
                band = hour_to_band(h, is_hol)
                band_sums[band] += commercial_hourly[h]
            for band in COLUMN_ORDER:
                commercial_output_band[band].append(band_sums[band])
        for i, day_idx in enumerate(daily.index):
            dess_values = dessmonitor_day_values.get(day_idx, np.zeros(24))
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0]
            is_hol = is_holiday(sample_date)
            # 時間帯ごとに集計
            band_sums = {'day': 0.0, 'home': 0.0, 'night': 0.0}
            for h in range(24):
                band = hour_to_band(h, is_hol)
                band_sums[band] += dess_values[h]
            # 積み上げ棒グラフとして表示（マイナス側、左y軸）
            bottom = 0.0
            for band in COLUMN_ORDER:
                value = -band_sums[band]
                ax.bar(day_idx, value, width=bar_width, bottom=bottom, color=dess_colors[band], alpha=0.7, label=dess_labels[band] if i==0 else "", align='center')
                bottom += value
        # 商用出力データをdeepskyblueで重ねる（時間帯ごとに積み上げ、幅は半分）
        commercial_bottom = np.zeros(len(x))
        for j, band in enumerate(COLUMN_ORDER):
            values = commercial_output_band[band]
            ax.bar(x, values, width=commercial_bar_width, bottom=commercial_bottom, color='deepskyblue', alpha=0.7, label='商用充電' if j==0 else "", align='center')
            commercial_bottom += np.array(values)
        # 軸・ラベル・タイトル
        #ax.set_xlabel('日')
        ax.set_ylabel('消費電力量 (kWh)')
        ax.set_title(f'{year_month} 月間（日別）電力使用量（Dessmonitor反映）')
        ax.set_xticks(x)
        # 土日祝日を赤文字に
        tick_labels = []
        for d in x:
            date_row = df_month[df_month['date'].dt.day == d]
            if not date_row.empty:
                date_val = date_row.iloc[0]['date']
                if is_holiday(date_val):
                    tick_labels.append({'label': str(d), 'color': 'red'})
                else:
                    tick_labels.append({'label': str(d), 'color': 'black'})
            else:
                tick_labels.append({'label': str(d), 'color': 'black'})
        ax.set_xticklabels([tl['label'] for tl in tick_labels], rotation=0)
        for tick, tl in zip(ax.get_xticklabels(), tick_labels):
            tick.set_color(tl['color'])
        # グリッド線
        ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
        import matplotlib.ticker as ticker
        ax.yaxis.set_major_locator(ticker.MultipleLocator(5))
        # 右軸: 電気料金折れ線
        costs = []
        net_saving_costs = []
        for idx, day_idx in enumerate(daily.index):
            date_row = df_month[df_month['date'].dt.day == day_idx]
            if not date_row.empty:
                d = date_row.iloc[0]
                bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night'])
                c = bd.get('total_cost', 0.0)
            else:
                c = 0.0
            costs.append(c)
            # 商用充電分を差し引いた節約金額
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0]
            is_hol = is_holiday(sample_date)
            dess_band_sums = band_sums_from_values(dessmonitor_day_values[day_idx], is_hol)
            dess_bd = compute_cost_breakdown(sample_date, dess_band_sums['day'], dess_band_sums['home'], dess_band_sums['night'])
            # 商用充電
            commercial_band_sums = {band: commercial_output_band[band][idx] for band in COLUMN_ORDER}
            commercial_bd = compute_cost_breakdown(sample_date, commercial_band_sums['day'], commercial_band_sums['home'], commercial_band_sums['night'])
            net_saving = dess_bd['total_cost'] - commercial_bd['total_cost']
            net_saving_costs.append(net_saving)

        # 凡例
        from matplotlib.patches import Patch
        legend_patches = [Patch(color=TIME_BANDS[c]['color'], label=TIME_BANDS[c]['label']) for c in DISPLAY_ORDER]
        legend_patches += [Patch(color=dess_colors[c], label=dess_labels[c]) for c in DISPLAY_ORDER]
        legend_patches.append(Patch(color='deepskyblue', label='商用充電'))
        ax.legend(handles=legend_patches, bbox_to_anchor=(1.05, 1), loc='upper left')
        # 集計情報（テキストボックス）
        if len(daily.index) > 0:
            # 月の初日で単価を取得
            day_idx = daily.index[0]
            sample_date = df_month[df_month['date'].dt.day == day_idx]['date'].iloc[0]
            prices = get_unit_prices_for_date(sample_date)
            renew = get_renewable_unit_for_date(sample_date)
            fuel = get_fuel_adj_for_date(sample_date)
            unit_text = (f"適用単価:\n"
                         f"デイタイム単価　 {prices['day']:.2f} 円/kWh\n"
                         f"ホームタイム単価 {prices['home']:.2f} 円/kWh\n"
                         f"ナイトタイム単価 {prices['night']:.2f} 円/kWh\n"
                         f"再エネ賦課金単価 {renew:.2f} 円/kWh\n"
                         f"燃料費調整単価　 {fuel:.2f} 円/kWh")
            band_lines = (f"使用電力量:\n"
                          f"デイタイム電力　 {month_total['day']:.2f} kWh\n"
                          f"ホームタイム電力 {month_total['home']:.2f} kWh\n"
                          f"ナイトタイム電力 {month_total['night']:.2f} kWh\n"
                          f"買電電力量　　　 {month_buy_total:.2f} kWh\n"
                          f"蓄電電力量　　　 {dess_month_total_sum:.2f} kWh\n"
                          f"商用充電電力量　 {commercial_month_total_sum:.2f} kWh")
            # 商用充電料金を計算
            commercial_charge_total = 0.0
            for day_idx2 in daily.index:
                sample_date2 = df_month[df_month['date'].dt.day == day_idx2]['date'].iloc[0]
                is_hol2 = is_holiday(sample_date2)
                commercial_band_sums2 = {band: commercial_output_band[band][list(daily.index).index(day_idx2)] for band in COLUMN_ORDER}
                commercial_bd2 = compute_cost_breakdown(sample_date2, commercial_band_sums2['day'], commercial_band_sums2['home'], commercial_band_sums2['night'])
                commercial_charge_total += commercial_bd2['total_cost']
            # 1ヶ月分の節約金額合計を計算し、商用充電料金を差し引く
            month_saving_total = 0.0
            for day_idx2 in daily.index:
                sample_date2 = df_month[df_month['date'].dt.day == day_idx2]['date'].iloc[0]
                is_hol2 = is_holiday(sample_date2)
                dess_band_sums2 = band_sums_from_values(dessmonitor_day_values[day_idx2], is_hol2)
                dess_bd2 = compute_cost_breakdown(sample_date2, dess_band_sums2['day'], dess_band_sums2['home'], dess_band_sums2['night'])
                month_saving_total += dess_bd2['total_cost']
            net_saving = month_saving_total - commercial_charge_total
            totals = (f"集計金額:\n"
                      f"デイタイム金額　 {month_costs['day']:.0f} 円\n"
                      f"ホームタイム金額 {month_costs['home']:.0f} 円\n"
                      f"ナイトタイム金額 {month_costs['night']:.0f} 円\n"
                      f"再エネ賦課金額　 {month_costs['renew']:.2f} 円\n"
                      f"燃料調整費金額　 {month_costs['fuel']:.2f} 円\n"
                      f"買電金額　　　　 {month_costs['total']:.0f} 円\n"
                      f"商用充電料金　　 {commercial_charge_total:.0f} 円\n"
                      f"節約金額　　　　 {net_saving:.0f} 円\n"
                      f"太陽光発電無金額 {month_costs['total'] + net_saving:.0f} 円")
            summary_text = unit_text + "\n\n" + band_lines + "\n\n" + totals
            ax.text(1.15, 0.7, summary_text, transform=ax.transAxes,
                ha='left', va='top', fontsize=9,
                bbox=dict(boxstyle='round,pad=0.5', fc='white', ec='black', alpha=0.85))
        plt.tight_layout()
        # Tkウインドウ表示
        win = tk.Toplevel() if tk._default_root else tk.Tk()
        win.wm_title(f'{year_month} 日別Dessmonitor')
        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        widget = canvas.get_tk_widget()
        widget.pack(side='top', fill='both', expand=1)

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


        # --- ホバー注釈: 各積み上げbarのnight/home/dayごとに対応 ---
        annotation = None
        fade_job = None
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
        def schedule_fade(timeout_ms: int = 2500):
            nonlocal fade_job
            if widget is None:
                return
            try:
                if fade_job is not None:
                    widget.after_cancel(fade_job)
            except Exception:
                pass
            try:
                fade_job = widget.after(timeout_ms, remove_annotation)
            except Exception:
                fade_job = None
        def on_motion(event):
            nonlocal annotation
            if getattr(event, 'inaxes', None) is None:
                remove_annotation()
                return
            # 各積み上げbarのnight/home/dayごとに対応
            for cont, band in zip([bars, bars_home, bars_day], ['night', 'home', 'day']):
                for idx, rect in enumerate(cont):
                    try:
                        contains = rect.contains(event)[0]
                    except Exception:
                        contains = False
                    if contains:
                        day = int(x[idx]) if idx < len(x) else None
                        if day is None or day not in daily.index:
                            continue
                        value = rect.get_height()
                        # 金額計算
                        sample_date = df_month[df_month['date'].dt.day == day]['date'].iloc[0]
                        is_hol = is_holiday(sample_date)
                        bd = compute_cost_breakdown(sample_date, value if band=='day' else 0, value if band=='home' else 0, value if band=='night' else 0)
                        msg = (
                            f'{sample_date.year}年{sample_date.month}月{sample_date.day}日\n'
                            f'{TIME_BANDS[band]["label"]}: {value:.2f}kWh ({bd["base_costs"][band]:.0f}円)'
                        )
                        if annotation is not None:
                            try:
                                annotation.remove()
                            except Exception:
                                pass
                        x_ = rect.get_x() + rect.get_width()/2
                        y_ = rect.get_y() + rect.get_height()
                        annotation = annotate_in_axes(ax, fig, x_, y_, msg)
                        schedule_fade()
                        return
            remove_annotation()
        def on_leave(event):
            remove_annotation()
        fig.canvas.mpl_connect('motion_notify_event', on_motion)
        fig.canvas.mpl_connect('figure_leave_event', on_leave)

        # 左y軸と右y軸の0ラインを完全一致させる
        ax2 = ax.twinx()
        costs = []
        for day_idx in daily.index:
            date_row = df_month[df_month['date'].dt.day == day_idx]
            if not date_row.empty:
                d = date_row.iloc[0]
                bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night'])
                c = bd.get('total_cost', 0.0)
            else:
                c = 0.0
            costs.append(c)
        ax2.plot(daily.index, costs, color='red', marker='o', linewidth=0.5, label='電気料金 (円)')
        ax2.plot(daily.index, [-v for v in net_saving_costs], color='blue', marker='o', linewidth=0.5, label='節約金額 (円,反転)')
        ax2.set_ylabel('電気料金 (円)')
        # 0ラインを完全一致させる
        left_ylim = ax.get_ylim()
        min_ylim = min(0, left_ylim[0])
        max_ylim = left_ylim[1]+5
        ax.set_ylim(min_ylim, max_ylim)
        ax2.set_ylim(min_ylim*50, max_ylim*50)# rough scaling factor
        ax2.legend(loc='upper right')

        win.mainloop()

    def on_motion(event):
        nonlocal annotation
        # More robust check for valid axes and data elements
        if getattr(event, 'inaxes', None) is None or not hasattr(event.inaxes, 'patches'):
            remove_annotation()
            return
        for rect in getattr(ax, 'patches', []):
            try:
                contains = rect.contains(event)[0]
            except Exception:
                contains = False
            if contains:
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
                schedule_fade()
                return
        remove_annotation()

    def on_button(event):
        # left double-click opens hourly
        if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1:
            # find which rect was clicked
            for axes in plt.gcf().axes:
                for rect in getattr(axes, 'patches', []):
                    try:
                        contains = rect.contains(event)[0]
                    except Exception:
                        contains = False
                    if contains:
                        day = getattr(rect, '_day', None)
                        if day is not None:
                            year, month = map(int, year_month.split('-'))
                            plot_hourly(df, year, month, int(day), file_path=file_path, dessmonitor_folder=dessmonitor_folder)
                            return
        # right click -> context menu
        if getattr(event, 'button', None) == 3:
            show_context_menu(event)

    try:
        fig.canvas.mpl_connect('motion_notify_event', on_motion)
        fig.canvas.mpl_connect('figure_leave_event', lambda ev: remove_annotation())
        fig.canvas.mpl_connect('button_press_event', on_button)
    except Exception:
        # fallback keep old on_click
        def on_click(event):
            nonlocal annotation
            bar_clicked = False
            for axes in plt.gcf().axes:
                for rect in getattr(axes, 'patches', []):
                    try:
                        contains = rect.contains(event)[0]
                    except Exception:
                        contains = False
                    if contains:
                        day = getattr(rect, '_day', None)
                        if day is not None:
                            year, month = map(int, year_month.split('-'))
                            if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1:
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
    # 日別の料金を計算して右軸に描画
    costs = []
    for day_idx in daily.index:
        date_row = df_month[df_month['date'].dt.day == day_idx]
        if not date_row.empty:
            d = date_row.iloc[0]
            bd = compute_cost_breakdown(d['date'], d['day'], d['home'], d['night'])
            c = bd.get('total_cost', 0.0)
        else:
            c = 0.0
        costs.append(c)
    ax2 = ax.twinx()
    ax2.plot(np.arange(len(daily.index)), costs, color='red', marker='o', linewidth=0.5, label='電気料金 (円)')
    ax2.set_ylabel('電気料金 (円)')
    # 左右y軸の0ラインを完全一致させる
    left_ylim = ax.get_ylim()
    min_ylim = min(0, left_ylim[0])
    max_ylim = left_ylim[1]+5
    ax.set_ylim(min_ylim, max_ylim)
    ax2.set_ylim(min_ylim*50, max_ylim*50)  # rough scaling factor
    ax2.legend(loc='upper right')
    # background-click save handler removed (deleted to simplify code)
    plt.tight_layout()

    # embed into Tk Toplevel instead of plt.show() to avoid creating separate matplotlib windows
    master = tk._default_root if tk._default_root is not None else None
    if master is None:
        win = tk.Tk()
    else:
        win = tk.Toplevel(master)

    # ウインドウタイトルをファイルパスに設定
    if file_path is not None:
        try:
            win.wm_title(str(file_path))
        except Exception:
            pass

    try:
        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        widget = canvas.get_tk_widget()
        widget.pack(side='top', fill='both', expand=1)
    except Exception:
        # fallback to interactive show if embedding fails
        try:
            plt.show()
            return
        except Exception:
            pass

    annotation_local = None
    def _on_close():
        try:
            plt.close(fig)
        except Exception:
            pass
        try:
            win.destroy()
        except Exception:
            pass
        # if no default root remaining, exit
        try:
            if not tk._default_root:
                sys.exit(0)
        except Exception:
            pass

    try:
        win.protocol('WM_DELETE_WINDOW', _on_close)
    except Exception:
        pass

    # start mainloop only if we created a standalone root
    if master is None:
        try:
            win.mainloop()
        except Exception:
            pass

# 月別グラフをウインドウ表示し、クリックで日別グラフを表示
def plot_monthly_interactive(monthly: pd.DataFrame, df: pd.DataFrame, file_path: Path = None, dessmonitor_folder: str = None):
    """月別積み上げ棒グラフ（年比較）を描画し、年別の月次電気料金を色分けの折れ線で表示する。
    バーをクリックすると該当年月の日別グラフを開く。
    """
    # use common definitions from TIME_BANDS / COLUMN_ORDER / DISPLAY_ORDER
    # base_colors: use rgb tuples from TIME_BANDS so we can apply alpha shading per year
    base_colors = {band: TIME_BANDS[band]['color_rgb'] for band in TIME_BANDS}
    # copy and normalize index to Timestamp (first day of month)
    plot_data = monthly[COLUMN_ORDER].copy()
    plot_data.index = pd.to_datetime(plot_data.index.astype(str) + '-01')
    
    # 年別の表示/非表示状態を管理する辞書
    year_visibility = {}

    years = sorted(plot_data.index.year.unique())
    months = list(range(1, 13))

    fig, ax = plt.subplots(figsize=(15, 7))
    global LAST_FIG_OPEN
    LAST_FIG_OPEN = time.time()
    if file_path is not None:
        try:
            fig.canvas.manager.set_window_title(str(file_path))
        except Exception:
            pass

    x = np.arange(len(months))
    width = 0.8 / max(1, len(years))

    # 描画: 年ごとに並べた積み上げ棒
    for i, year in enumerate(years):
        year_rows = []
        for m in months:
            ts = pd.Timestamp(f'{year}-{m:02d}-01')
            if ts in plot_data.index:
                year_rows.append(plot_data.loc[ts])
            else:
                year_rows.append(pd.Series({c: 0.0 for c in COLUMN_ORDER}))
        year_df = pd.DataFrame(year_rows, columns=COLUMN_ORDER)

        bottom = np.zeros(len(months))
        # 少しずつ色を変えるためalphaを用いる
        alpha = 0.9 + (0.1 * i / (len(years) - 1)) if len(years) > 1 else 1.0
        for category in COLUMN_ORDER:
            color = tuple(c * alpha for c in base_colors[category])
            bars = ax.bar(x + i * width, year_df[category].values, width, bottom=bottom, color=color)
            # attach metadata for click handling
            for j, rect in enumerate(bars):
                rect._year = year
                rect._month = months[j]
            bottom += year_df[category].values

    # 軸・ラベル設定
    ax.set_ylabel('消費電力量 (kWh)')
    ax.set_title('年間（月別）電力使用量 (年比較)')
    ax.set_xticks(x + width * (len(years) - 1) / 2)
    ax.set_xticklabels([f'{m}月' for m in months])
    ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
    import matplotlib.ticker as ticker
    ax.yaxis.set_major_locator(ticker.MultipleLocator(100))

    # 凡例（時間帯） - 軸内に表示して常に全体が見えるようにする
    from matplotlib.patches import Patch
    legend_patches = [Patch(color=TIME_BANDS[c]['color'], label=TIME_BANDS[c]['label']) for c in DISPLAY_ORDER]
    ax.legend(handles=legend_patches, loc='lower right')

    # 右軸: 年別の月次コストを年ごとに色分けしてプロット
    ax2 = ax.twinx()
    cmap = plt.get_cmap('tab10')
    for i, year in enumerate(years):
        year_costs = []
        for m in months:
            ts = pd.Timestamp(f'{year}-{m:02d}-01')
            if ts in plot_data.index:
                row = plot_data.loc[ts]
                # row contains night/home/day
                c = compute_cost_from_parts(ts, row.get('day', 0.0), row.get('home', 0.0), row.get('night', 0.0))
            else:
                c = 0.0
            year_costs.append(c)
        color = cmap(i % 10)
        line = ax2.plot(x + i * width, year_costs, color=color, marker='o', linewidth=0.5, label=f'{year}年 電気料金')[0]
        line._year = year  # 年情報を線に付加

    ax2.set_ylabel('電気料金 (円)')
    # 固定表示: 右軸の下限を 0 にする
    set_ax2_min_zero(ax2)
    ax2.legend(loc='lower left')

    # 年別集計情報を表示
    year_totals = {}
    for year in years:
        year_data = {'day': 0.0, 'home': 0.0, 'night': 0.0, 'cost': 0.0}
        for month in months:
            try:
                ts = pd.Timestamp(f'{year}-{month:02d}-01')
                if ts in plot_data.index:
                    row = plot_data.loc[ts]
                    year_data['day'] += float(row.get('day', 0.0))
                    year_data['home'] += float(row.get('home', 0.0))
                    year_data['night'] += float(row.get('night', 0.0))
                    year_data['cost'] += compute_cost_from_parts(ts, row.get('day', 0.0), 
                                                               row.get('home', 0.0), 
                                                               row.get('night', 0.0))
            except Exception:
                continue
        year_totals[year] = year_data

    # 年別集計情報のテキストを作成
    years_text = "年間集計:\n\n"
    for year in years:
        total = year_totals[year]
        total_kwh = total['day'] + total['home'] + total['night']
        years_text += (f"{year}年\n"
                      f"デイタイム電力　 {total['day']:.2f} kWh\n"
                      f"ホームタイム電力 {total['home']:.2f} kWh\n"
                      f"ナイトタイム電力 {total['night']:.2f} kWh\n"
                      f"年間総使用電力量 {total_kwh:.2f} kWh\n"
                      f"年間合計金額　　 {total['cost']:.0f} 円\n\n")

    # 年別集計情報をテーブル化するデータを作成
    col_labels = ['年', 'デイ (kWh)', 'ホーム (kWh)', 'ナイト (kWh)', '総計 (kWh)', '年間金額 (円)']
    table_rows = []
    for year in years:
        t = year_totals[year]
        total_kwh = t['day'] + t['home'] + t['night']
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

    # Ensure matplotlib events still work
    def on_closing():
        try:
            plt.close('all')
        except Exception:
            pass
        try:
            win.destroy()
        except Exception:
            pass
        # exit process when main window closed
        try:
            sys.exit(0)
        except Exception:
            pass

    win.protocol('WM_DELETE_WINDOW', on_closing)

    annotation = None
    fade_job = None

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

    def schedule_fade(timeout_ms: int = 3000):
        nonlocal fade_job
        if widget is None:
            return
        try:
            if fade_job is not None:
                widget.after_cancel(fade_job)
        except Exception:
            pass
        try:
            fade_job = widget.after(timeout_ms, remove_annotation)
        except Exception:
            fade_job = None

    def save_via_dialog():
        try:
            title = getattr(ax, 'get_title', lambda: "graph")()
            fname = filedialog.asksaveasfilename(parent=win, defaultextension='.png', filetypes=[('PNG image','*.png')], initialfile=title if title else "graph")
            if fname:
                fig.savefig(fname, bbox_inches='tight')
                print(f'Saved figure to {fname}')
        except Exception as e:
            print('Save failed:', e)

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

    def on_motion(event):
        nonlocal annotation
        # More robust check for valid axes and data elements
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
                cost = compute_cost_from_parts(idx, row.get('day',0.0), row.get('home',0.0), row.get('night',0.0))
                msg = (
                    f'{year}年{int(month)}月\n'
                    f'デイタイム: {row.get("day",0.0):.2f}kWh\nホームタイム: {row.get("home",0.0):.2f}kWh\nナイトタイム: {row.get("night",0.0):.2f}kWh\n'
                    f'合計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: {cost:.0f}円'
                )
                if annotation is not None:
                    try:
                        annotation.remove()
                    except Exception:
                        pass
                x = rect.get_x() + rect.get_width()/2
                y = rect.get_y() + rect.get_height()
                annotation = annotate_in_axes(ax, fig, x, y, msg)
                schedule_fade()
                return
        remove_annotation()

    def on_button(event):
        # left double-click opens daily
        if getattr(event, 'dblclick', False) and getattr(event, 'button', None) == 1:
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
                        plot_daily(df, year_month, file_path=file_path, dessmonitor_folder=dessmonitor_folder)
                        return
        # right click -> context menu
        if getattr(event, 'button', None) == 3:
            show_context_menu(event)

    try:
        fig.canvas.mpl_connect('motion_notify_event', on_motion)
        fig.canvas.mpl_connect('figure_leave_event', lambda ev: remove_annotation())
        fig.canvas.mpl_connect('button_press_event', on_button)
    except Exception:
        # fallback to old click behavior
        def on_click(event):
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
                                cost = compute_cost_from_parts(idx, row.get('day',0.0), row.get('home',0.0), row.get('night',0.0))
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
    # background-click save handler removed (deleted to simplify code)
    # ensure canvas is drawn and start the Tk event loop
    try:
        canvas.draw()
    except Exception:
        pass
    win.mainloop()


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
        sys.exit(1)
    
    return Path(file_path)

def main():
    # コマンドライン引数があればそれを使用、なければファイル選択ダイアログを表示
    if len(sys.argv) > 1:
        path = Path(sys.argv[1])
    else:
        path = select_file()

    if not path.exists():
        print('File not found:', path)
        sys.exit(1)

    print(f'処理するファイル: {path}')
    # Dessmonitorフォルダー選択
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    dessmonitor_folder = filedialog.askdirectory(parent=root, title='Dessmonitorデータのフォルダーを選択してください')
    if not dessmonitor_folder:
        print('Dessmonitorフォルダーが選択されませんでした。')
        sys.exit(1)

    df = load_csv(path)
    monthly, df_all = parse_and_aggregate(df)
    plot_monthly_interactive(monthly, df_all, file_path=path, dessmonitor_folder=dessmonitor_folder)

if __name__ == '__main__':
    main()
