import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from datetime import datetime
import os

# 日本語フォント設定（必要に応じて）
matplotlib.rcParams['font.family'] = 'Meiryo'  # または 'MS Gothic'

# 銘柄コードと英語ラベル
stocks = {
    "7203.T": "TOYOTA",
    "6460.T": "SEGAsummy",
    "3765.T": "GungHo"
}

# ベース出力ディレクトリ
base_dir = os.path.join(os.path.dirname(__file__), "output")

# 今日の日付フォルダ（例: 2025-10-24）
today_str = datetime.today().strftime('%Y-%m-%d')
target_dir = os.path.join(base_dir, today_str)
os.makedirs(target_dir, exist_ok=True)

# タイムスタンプ付きファイル名（例: stock_data_20251024_0023.xlsx）
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
excel_path = os.path.join(target_dir, f"stock_data_{timestamp}.xlsx")

# Excelファイル作成
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    for code, label in stocks.items():
        df = yf.Ticker(code).history(interval="15m", period="1d")
        if df.empty:
            print(f"⚠️ データなし: {label}（{code}）")
            continue

        df.index = df.index.tz_localize(None)
        df.reset_index(inplace=True)
        df.to_excel(writer, sheet_name=label, index=False)

        # チャート画像保存
        chart_path = os.path.join(target_dir, f"{label}_chart.png")
        plt.figure(figsize=(10, 4))

        # 青線を少し薄くして視認性向上
        plt.plot(df["Datetime"], df["Close"], marker="o", color="blue", alpha=0.6, label="Close")

        # 最大値・最小値のインデックス
        max_idx = df["Close"].idxmax()
        min_idx = df["Close"].idxmin()

        # 最大値・最小値の座標
        max_time = df.loc[max_idx, "Datetime"]
        max_price = df.loc[max_idx, "Close"]
        min_time = df.loc[min_idx, "Datetime"]
        min_price = df.loc[min_idx, "Close"]

        # 最大値（赤）のマーカーと数値表示（上に表示）
        plt.plot(max_time, max_price, marker="o", color="red", markersize=10, label="Max")
        plt.text(max_time, max_price + 5, f"{max_price:.0f}", color="red", fontsize=12, ha="center", va="bottom")

        # 最小値（緑）のマーカーと数値表示（左右にずらして日付と重ならないように）
        plt.plot(min_time, min_price, marker="o", color="green", markersize=10, label="Min")
        min_ha = "left" if min_idx > len(df) // 2 else "right"
        min_offset = 0.02 * (df["Datetime"].max() - df["Datetime"].min())
        min_x = min_time + min_offset if min_ha == "left" else min_time - min_offset
        plt.text(min_x, min_price, f"{min_price:.0f}", color="green", fontsize=12, ha=min_ha, va="center")

        # 各点に価格ラベル表示（濃いグレー＋フォントサイズUP＋位置調整）
        for i in range(len(df)):
            time = df.loc[i, "Datetime"]
            price = df.loc[i, "Close"]
            if i != max_idx and i != min_idx:
                plt.text(time, price + 3, f"{price:.0f}", fontsize=9, ha="center", va="bottom", color="dimgray")

        # タイトル（英語表記）
        plt.title(label, fontsize=14, pad=20)
        plt.xlabel("Time", fontsize=12)
        plt.ylabel("Price", fontsize=12)
        plt.grid(True)
        plt.tight_layout()
        plt.savefig(chart_path)
        plt.close()

print(f"✅ 完了: {excel_path}")
print(f"📁 ファイル保存先: {target_dir}")
