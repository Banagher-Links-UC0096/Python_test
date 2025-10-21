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
    default_day_hours = list(range(9, 18))  # 9～17時 inclusive (9..17)
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
def plot_hourly(df: pd.DataFrame, year: int, month: int, day: int):
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
    plt.tight_layout()
    plt.show()

def plot_daily(df: pd.DataFrame, year_month: str):
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

    # クリックイベントで時間単位グラフを表示
    def on_click(event):
        if event.inaxes != ax:
            return
        for rect in ax.patches:
            if rect.contains(event)[0]:
                idx = int(rect.get_x() + rect.get_width()/2 + 0.5)
                days = daily.index.tolist()
                if 0 <= idx < len(days):
                    day = days[idx]
                    # 年月取得
                    year, month = map(int, year_month.split('-'))
                    plot_hourly(df, year, month, day)
                break
    fig.canvas.mpl_connect('button_press_event', on_click)
    plt.show()

# 月別グラフをウインドウ表示し、クリックで日別グラフを表示
def plot_monthly_interactive(monthly: pd.DataFrame, df: pd.DataFrame):
    columns_order = ['night', 'home', 'day']
    plot_data = monthly[columns_order].copy()
    
    # 年月を年と月に分割
    plot_data.index = pd.to_datetime(plot_data.index + '-01')
    years = plot_data.index.year.unique()
    months = range(1, 13)
    
    # 色の基本値を設定
    base_colors = {
        'day': (1.0, 0.92, 0.6),    # #FFEB99
        'home': (0.6, 0.9, 0.6),    # #99E699
        'night': (0.6, 0.8, 1.0)    # #99CCFF
    }
    
    labels = {
        'day': 'デイタイム',
        'home': 'ホームタイム',
        'night': 'ナイトタイム'
    }
    
    fig, ax = plt.subplots(figsize=(15,7))
    
    # 各月のx軸位置を計算
    x = np.arange(len(months))
    width = 0.8 / len(years)  # バーの幅
    
    # 年ごとにバーを描画（古い年から新しい年へ）
    legend_handles = []
    for i, year in enumerate(sorted(years)):  # 古い年から順に処理
        # 古い年ほど濃い色にする（alphaが小さいほど濃い）
        alpha = 0.7 + (0.3 * i / (len(years) - 1)) if len(years) > 1 else 1.0
        
        year_colors = {
            category: tuple(c * alpha for c in base_colors[category])
            for category in columns_order
        }
        
        # 月ごとのデータを抽出
        year_data = []
        for month in months:
            try:
                idx = pd.Timestamp(f'{year}-{month:02d}-01')
                if idx in plot_data.index:
                    year_data.append(plot_data.loc[idx])
                else:
                    year_data.append(pd.Series({col: 0 for col in columns_order}))
            except:
                year_data.append(pd.Series({col: 0 for col in columns_order}))
        
        year_df = pd.DataFrame(year_data, columns=columns_order)
        bottom = np.zeros(len(months))
        
        for category in columns_order:
            bars = ax.bar(x + width * i, year_df[category], width,
                         bottom=bottom, color=year_colors[category],
                         label=f'{year}年 {labels[category]}' if category == columns_order[0] else None)
            bottom += year_df[category].values
            
            if category == columns_order[0]:
                legend_handles.append(bars)
            # attach year/month metadata to each Rectangle so click handler can find the correct date
            try:
                for j, rect in enumerate(bars):
                    # months is range(1,13) so index j corresponds to months[j]
                    rect._year = year
                    rect._month = months[j]
            except Exception:
                pass
    
    #ax.set_xlabel('月')
    ax.set_ylabel('消費電力量 (kWh)')
    ax.set_title('月別・時間帯別電力使用量 (年比較)')
    
    # X軸の設定
    ax.set_xticks(x + width * (len(years) - 1) / 2)
    ax.set_xticklabels([f'{m}月' for m in months])
    
    # グリッド線の設定
    ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
    import matplotlib.ticker as ticker
    ax.yaxis.set_major_locator(ticker.MultipleLocator(100))
    
    # 凡例の作成：カテゴリと年の組み合わせ
    handles = []
    labels_list = []
    
    # まず年ごとの凡例エントリを作成
    for year in sorted(years):
        for category in columns_order:
            i = list(sorted(years)).index(year)
            alpha = 0.7 + (0.3 * i / (len(years) - 1)) if len(years) > 1 else 1.0
            color = tuple(c * alpha for c in base_colors[category])
            patch = plt.Rectangle((0, 0), 1, 1, fc=color)
            handles.append(patch)
            labels_list.append(f'{year}年 {labels[category]}')
    
    # 凡例を配置
    ax.legend(handles, labels_list, bbox_to_anchor=(1.05, 1), loc='upper left')
    
    plt.tight_layout()

    # クリックイベントの処理を追加（1つのハンドラに統合）
    def on_click(event):
        if event.inaxes != ax:
            return
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
                    # 存在する年月なら表示
                    if any(ym.startswith(year_month) for ym in plot_data.index.strftime('%Y-%m')):
                        plot_daily(df, year_month)
                break

    fig.canvas.mpl_connect('button_press_event', on_click)
    # 100kWh単位でグリッド線
    ax.yaxis.grid(True, which='major', linestyle='--', color='gray', alpha=0.7)
    import matplotlib.ticker as ticker
    ax.yaxis.set_major_locator(ticker.MultipleLocator(100))
    plt.tight_layout()

    # (以前の重複ハンドラは削除済み)
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
    plot_monthly_interactive(monthly, df_all)

if __name__ == '__main__':
    main()
