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
import tkinter as tk
from tkinter import filedialog
from matplotlib import font_manager
import jpholiday
import datetime
import time

# 日本語フォントの設定
plt.rcParams['font.family'] = 'MS Gothic'  # Windows の場合

# 時間帯定義
# columns: CSVの10列目が0時-1時 => pandasで読み込むと0-origin index 9が0時
DAY_HOURS = list(range(9, 17))   # 9時～16時のカラムインデックス(10列目=0時なので9->9時)
HOME_HOURS = list(range(7, 9)) + list(range(17, 23))  # 7-8時,17-22時
NIGHT_HOURS = list(range(23, 24)) + list(range(0, 7))  # 23時,0-6時

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


# ---- 料金関連の定義とヘルパー ----
# 時間帯単価（円/kWh）: 適用期間は ~2023年3月以前, ~2024年3月以前, それ以後
PRICE_PERIODS = [
    # (start_inclusive, end_inclusive, prices)
    (None, pd.Timestamp('2023-03-31'), {'day': 33.97, 'home': 25.91, 'night': 15.89}),
    (pd.Timestamp('2023-04-01'), pd.Timestamp('2024-03-31'), {'day': 34.21, 'home': 26.15, 'night': 16.22}),
    (pd.Timestamp('2024-04-01'), None, {'day': 34.06, 'home': 26.00, 'night': 16.11}),
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

# Timestamp of last figure creation to avoid cascade save on same click
LAST_FIG_OPEN = 0.0


def connect_save_on_bg_click(fig, file_path: Path = None, prefix: str = 'plot'):
    """図の背景（プロット要素以外）をクリックしたらPNGで保存するハンドラを接続する。
    保存先は file_path の親フォルダ（file_path が None の場合はカレントディレクトリ）。
    ファイル名: <stem>_<prefix>_YYYYmmdd_HHMMSS.png
    """
    out_dir = Path(file_path).parent if file_path is not None else Path.cwd()
    stem = Path(file_path).stem if file_path is not None else 'plot'

    def on_click_save(event):
        # Only react on double-click to trigger save
        if not getattr(event, 'dblclick', False):
            return

        # if a figure was just opened very recently, ignore this event to avoid cascade saving
        try:
            if time.time() - LAST_FIG_OPEN < 0.35:
                return
        except Exception:
            pass
        # クリックがどの axes でもない（図の余白）なら保存
        if event.inaxes is None:
            ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            fname = out_dir / f"{stem}_{prefix}_{ts}.png"
            try:
                fig.savefig(fname, bbox_inches='tight')
                print(f'Saved figure to {fname}')
            except Exception as e:
                print('Failed to save figure:', e)
            return

        # クリックがすでに何らかのデータArtistに当たっているかどうか調べる
        # (Axesの背景や凡例、スパイン等は無視する)
        from matplotlib.lines import Line2D
        from matplotlib.collections import Collection
        from matplotlib.patches import Patch
        from matplotlib.image import AxesImage
        from matplotlib.legend import Legend

        data_types = (Line2D, Collection, Patch, AxesImage)

        for ax in fig.axes:
            for art in ax.findobj(lambda a: isinstance(a, data_types)):
                # ignore axes background patch
                try:
                    if art is ax.patch:
                        continue
                    # ignore legend artists
                    if isinstance(art, Legend):
                        continue
                    # some patches (like axis background) are caught above; now test contains
                    contains = False
                    try:
                        contains = art.contains(event)[0]
                    except Exception:
                        contains = False
                    if contains:
                        # アーティスト上のクリック -> 保存しない
                        return
                except Exception:
                    continue

        # ここまで来たら要素上ではないクリック --> 保存
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        fname = out_dir / f"{stem}_{prefix}_{ts}.png"
        try:
            fig.savefig(fname, bbox_inches='tight')
            print(f'Saved figure to {fname}')
        except Exception as e:
            print('Failed to save figure:', e)

    fig.canvas.mpl_connect('button_press_event', on_click_save)


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
    months_list = list(hour_to_col.keys())  # 0..23
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



# 日ごとのグラフを表示する関数
def plot_hourly(df: pd.DataFrame, year: int, month: int, day: int, file_path: Path = None):
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
            if 9 <= h <= 17 or (7 <= h <= 8) or (17 < h <= 22):
                colors.append('#99E699')  # ホームタイム (休日は9-17含む)
            elif 23 == h or 0 <= h <= 6:
                colors.append('#99CCFF')  # ナイトタイム
            else:
                colors.append('#99CCFF')
        else:
            if 9 <= h <= 17:
                colors.append('#FFEB99')  # デイタイム
            elif 7 <= h <= 8 or 17 < h <= 22:
                colors.append('#99E699')  # ホームタイム
            else:
                colors.append('#99CCFF')  # ナイトタイム
    fig, ax = plt.subplots(figsize=(10,5))
    global LAST_FIG_OPEN
    LAST_FIG_OPEN = time.time()
    # ウインドウタイトルをファイルパスに設定
    if file_path is not None:
        try:
            fig.canvas.manager.set_window_title(str(file_path))
        except Exception:
            pass
    ax.bar(range(24), values, color=colors)
    #ax.set_xlabel('時')
    ax.set_ylabel('消費電力量 (kWh)')
    ax.set_title(f'{year}-{month:02d}-{day:02d} 時間単位の電力使用量')
    ax.set_xticks(range(24))
    # 凡例
    from matplotlib.patches import Patch
    # 凡例: 注意書きとして休日は9-17がホームに含まれることをコメント
    legend_patches = [
        Patch(color='#FFEB99', label='デイタイム(平日:9-17時)'),
        Patch(color='#99E699', label='ホームタイム(平日:7-8,18-22時。土日祝は9-17も含む)'),
        Patch(color='#99CCFF', label='ナイトタイム(23,0-6時)')
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

    # 時間ごとの単価を作る
    hourly_unit = []
    for h in range(24):
        # 休日は9-17がホーム扱い
        if is_hol:
            if 7 <= h <= 22:
                band = 'home'
            elif 23 == h or 0 <= h <= 6:
                band = 'night'
            else:
                band = 'night'
        else:
            if 9 <= h <= 17:
                band = 'day'
            elif 7 <= h <= 8 or 18 <= h <= 22:
                band = 'home'
            else:
                band = 'night'
        hourly_unit.append(prices[band] + renew + fuel)

    hourly_costs = values * np.array(hourly_unit)
    ax2 = ax.twinx()
    ax2.plot(range(24), hourly_costs, color='red', marker='o', linewidth=0.5, label='電気料金 (円)')
    ax2.set_ylabel('電気料金 (円)')
    ax2.yaxis.grid(False)
    ax2.legend(loc='upper right')
    # connect save-on-background-click
    connect_save_on_bg_click(fig, file_path=file_path, prefix=f'hourly_{year}_{month:02d}_{day:02d}')
    plt.tight_layout()
    plt.show()

def plot_daily(df: pd.DataFrame, year_month: str, file_path: Path = None):
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
    columns_order = ['night', 'home', 'day']
    colors = {
        'day': '#FFEB99',
        'home': '#99E699',
        'night': '#99CCFF'
    }
    labels = {
        'day': 'デイタイム',
        'home': 'ホームタイム',
        'night': 'ナイトタイム'
    }
    fig, ax = plt.subplots(figsize=(12,7))
    global LAST_FIG_OPEN
    LAST_FIG_OPEN = time.time()
    if file_path is not None:
        try:
            fig.canvas.manager.set_window_title(str(file_path))
        except Exception:
            pass
    bars = daily[columns_order].plot(kind='bar', stacked=True, color=[colors[col] for col in columns_order], ax=ax, legend=False)
    #ax.set_xlabel('日')
    ax.set_ylabel('消費電力量 (kWh)')
    ax.set_title(f'{year_month} 日別・時間帯別電力使用量')
    handles, _ = ax.get_legend_handles_labels()
    ax.legend(handles, [labels[col] for col in columns_order], bbox_to_anchor=(1.05, 1), loc='upper left')
    # 5kWh単位でグリッド線
    ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
    import matplotlib.ticker as ticker
    ax.yaxis.set_major_locator(ticker.MultipleLocator(5))
    plt.tight_layout()

    # attach day metadata to each bar rectangle using containers (reliable)
    days = daily.index.tolist()
    bar_rects = []
    for cont in getattr(ax, 'containers', []):
        for idx, rect in enumerate(cont):
            if idx < len(days):
                setattr(rect, '_day', int(days[idx]))
                bar_rects.append(rect)

    annotation = None
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
                        # ダブルクリックならサブウインドウ表示
                        if getattr(event, 'dblclick', False):
                            plot_hourly(df, year, month, int(day), file_path=file_path)
                        else:
                            # シングルクリックなら情報をグラフ上にアノテーション表示
                            row = daily.loc[int(day)] if int(day) in daily.index else None
                            if row is not None:
                                date_row = df_month[df_month['date'].dt.day == int(day)]
                                if not date_row.empty:
                                    d = date_row.iloc[0]
                                    cost = compute_cost_from_parts(d['date'], d['day'], d['home'], d['night'])
                                else:
                                    cost = 0.0
                                msg = (
                                    f'{year}年{month}月{int(day)}日\n'
                                    f'デイ: {row.get("day",0.0):.2f}kWh  ホーム: {row.get("home",0.0):.2f}kWh  ナイト: {row.get("night",0.0):.2f}kWh\n'
                                    f'合計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: {cost:.0f}円'
                                )
                                if annotation is not None:
                                    annotation.remove()
                                x = rect.get_x() + rect.get_width()/2
                                y = rect.get_y() + rect.get_height()
                                # draw annotation on the top-most axes so it appears above twin axes (cost lines)
                                top_ax = fig.axes[-1] if len(fig.axes) > 0 else axes
                                # transform (x,y) from the original axes data coords to top_ax data coords
                                try:
                                    disp = axes.transData.transform((x, y))
                                    xy_top = top_ax.transData.inverted().transform(disp)
                                except Exception:
                                    xy_top = (x, y)
                                annotation = top_ax.annotate(msg, xy=xy_top, xytext=(0, 10), xycoords='data', textcoords='offset points',
                                    ha='center', va='bottom', fontsize=10, bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.7), zorder=1000)
                                annotation.set_clip_on(False)
                                fig.canvas.draw_idle()
                        bar_clicked = True
                    break
            if bar_clicked:
                break
        # バー以外のダブルクリックはPNG保存（connect_save_on_bg_clickが担当）
        # ここでは何もしない

    fig.canvas.mpl_connect('button_press_event', on_click)
    # 日別の料金を計算して右軸に描画
    # daily は day/home/night のkWh合計
    # 各日のコストを compute_cost_from_parts で計算
    costs = []
    for day_idx in daily.index:
        # find corresponding date in df_month
        date_row = df_month[df_month['date'].dt.day == day_idx]
        if not date_row.empty:
            # sum per day already aggregated
            d = date_row.iloc[0]
            c = compute_cost_from_parts(d['date'], d['day'], d['home'], d['night'])
        else:
            c = 0.0
        costs.append(c)

    ax2 = ax.twinx()
    ax2.plot(np.arange(len(daily.index)), costs, color='red', marker='o', linewidth=0.5, label='電気料金 (円)')
    ax2.set_ylabel('電気料金 (円)')
    ax2.legend(loc='upper right')
    # connect save-on-background-click
    connect_save_on_bg_click(fig, file_path=file_path, prefix=f'daily_{ym.year}_{ym.month:02d}')
    plt.tight_layout()
    plt.show()

# 月別グラフをウインドウ表示し、クリックで日別グラフを表示
def plot_monthly_interactive(monthly: pd.DataFrame, df: pd.DataFrame, file_path: Path = None):
    """月別積み上げ棒グラフ（年比較）を描画し、年別の月次電気料金を色分けの折れ線で表示する。
    バーをクリックすると該当年月の日別グラフを開く。
    """
    columns_order = ['night', 'home', 'day']
    labels = {'day': 'デイタイム', 'home': 'ホームタイム', 'night': 'ナイトタイム'}

    # copy and normalize index to Timestamp (first day of month)
    plot_data = monthly[columns_order].copy()
    plot_data.index = pd.to_datetime(plot_data.index.astype(str) + '-01')

    years = sorted(plot_data.index.year.unique())
    months = list(range(1, 13))

    # 色の基本値を設定
    base_colors = {
        'day': (1.0, 0.92, 0.6),    # #FFEB99
        'home': (0.6, 0.9, 0.6),    # #99E699
        'night': (0.6, 0.8, 1.0)    # #99CCFF
    }

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
                year_rows.append(pd.Series({c: 0.0 for c in columns_order}))
        year_df = pd.DataFrame(year_rows, columns=columns_order)

        bottom = np.zeros(len(months))
        # 少しずつ色を変えるためalphaを用いる
        alpha = 0.7 + (0.3 * i / (len(years) - 1)) if len(years) > 1 else 1.0
        for category in columns_order:
            color = tuple(c * alpha for c in base_colors[category])
            bars = ax.bar(x + i * width, year_df[category].values, width, bottom=bottom, color=color)
            # attach metadata for click handling
            for j, rect in enumerate(bars):
                rect._year = year
                rect._month = months[j]
            bottom += year_df[category].values

    # 軸・ラベル設定
    ax.set_ylabel('消費電力量 (kWh)')
    ax.set_title('月別・時間帯別電力使用量 (年比較)')
    ax.set_xticks(x + width * (len(years) - 1) / 2)
    ax.set_xticklabels([f'{m}月' for m in months])
    ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
    import matplotlib.ticker as ticker
    ax.yaxis.set_major_locator(ticker.MultipleLocator(100))

    # 凡例（時間帯）
    from matplotlib.patches import Patch
    legend_patches = [Patch(color=base_colors[c], label=labels[c]) for c in columns_order]
    ax.legend(handles=legend_patches, bbox_to_anchor=(1.05, 1), loc='upper left')

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
        ax2.plot(x + i * width, year_costs, color=color, marker='o', linewidth=0.5, label=f'{year}年 電気料金')

    ax2.set_ylabel('電気料金 (円)')
    ax2.legend(bbox_to_anchor=(1.05, 0.5), loc='upper left')

    annotation = None
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
                    # ダブルクリックならサブウインドウ表示
                    if getattr(event, 'dblclick', False):
                        plot_daily(df, year_month, file_path=file_path)
                    else:
                        # シングルクリックなら情報をグラフ上にアノテーション表示
                        row = None
                        try:
                            idx = pd.Timestamp(f'{year}-{int(month):02d}-01')
                            row = monthly.loc[f'{year}-{int(month):02d}']
                        except Exception:
                            pass
                        if row is not None:
                            cost = compute_cost_from_parts(idx, row.get('day',0.0), row.get('home',0.0), row.get('night',0.0))
                            msg = (
                                f'{year}年{int(month)}月\n'
                                f'デイ: {row.get("day",0.0):.2f}kWh  ホーム: {row.get("home",0.0):.2f}kWh  ナイト: {row.get("night",0.0):.2f}kWh\n'
                                f'合計: {row.get("day",0.0)+row.get("home",0.0)+row.get("night",0.0):.2f}kWh  料金: {cost:.0f}円'
                            )
                            # 既存annotation削除
                            if annotation is not None:
                                annotation.remove()
                            # 棒グラフの中央上に表示
                            x = rect.get_x() + rect.get_width()/2
                            y = rect.get_y() + rect.get_height()
                            # draw annotation on the top-most axes so it appears above twin axes (cost lines)
                            top_ax = fig.axes[-1] if len(fig.axes) > 0 else ax
                            try:
                                disp = ax.transData.transform((x, y))
                                xy_top = top_ax.transData.inverted().transform(disp)
                            except Exception:
                                xy_top = (x, y)
                            annotation = top_ax.annotate(msg, xy=xy_top, xytext=(0, 10), xycoords='data', textcoords='offset points',
                                ha='center', va='bottom', fontsize=10, bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.7), zorder=1000)
                            annotation.set_clip_on(False)
                            fig.canvas.draw_idle()
                    bar_clicked = True
                break
        # バー以外のダブルクリックはPNG保存（connect_save_on_bg_clickが担当）
        # ここでは何もしない

    fig.canvas.mpl_connect('button_press_event', on_click)
    # connect save-on-background-click for the monthly figure
    connect_save_on_bg_click(fig, file_path=file_path, prefix='monthly')
    plt.tight_layout()
    plt.show()


def select_file() -> Path:
    # GUIのメインウィンドウを作成（表示はしない）
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを非表示

    # ファイル選択ダイアログを表示
    file_path = filedialog.askopenfilename(
        title='CSVファイルを選択してください',
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
    df = load_csv(path)
    monthly, df_all = parse_and_aggregate(df)
    plot_monthly_interactive(monthly, df_all, file_path=path)

if __name__ == '__main__':
    main()
