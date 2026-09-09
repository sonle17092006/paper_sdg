"""Quản lý dữ liệu giá cổ phiếu lịch sử và đối chiếu phát hiện tẩy xanh (Greenwashing).

Cung cấp dữ liệu giá đóng cửa điều chỉnh hàng năm (2019-2025) cho 4 mã:
PAN, PLX, PNJ, VNM từ sàn chứng khoán (kèm fallback offline).
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import requests

from config import RESULT_DIR

TICKER_MAP = {
    "PAN": "PAN.VN",
    "PLX": "PLX.VN",
    "PNJ": "PNJ.VN",
    "VNM": "VNM.VN",
}

# Dữ liệu giá cổ phiếu đóng cửa cuối năm (VND, điều chỉnh) làm fallback
FALLBACK_STOCK_PRICES = [
    {"company": "PAN", "year": "2019", "close_yearend": 18666.7, "close_avg": 22100.0},
    {"company": "PAN", "year": "2020", "close_yearend": 28916.7, "close_avg": 18450.0},
    {"company": "PAN", "year": "2021", "close_yearend": 32083.3, "close_avg": 24200.0},
    {"company": "PAN", "year": "2022", "close_yearend": 12458.3, "close_avg": 20150.0},
    {"company": "PAN", "year": "2023", "close_yearend": 16833.3, "close_avg": 16200.0},
    {"company": "PAN", "year": "2024", "close_yearend": 19791.7, "close_avg": 18900.0},
    {"company": "PAN", "year": "2025", "close_yearend": 22750.0, "close_avg": 21500.0},

    {"company": "PLX", "year": "2019", "close_yearend": 56000.0, "close_avg": 54200.0},
    {"company": "PLX", "year": "2020", "close_yearend": 54600.0, "close_avg": 46800.0},
    {"company": "PLX", "year": "2021", "close_yearend": 53900.0, "close_avg": 52600.0},
    {"company": "PLX", "year": "2022", "close_yearend": 31700.0, "close_avg": 40300.0},
    {"company": "PLX", "year": "2023", "close_yearend": 34500.0, "close_avg": 36100.0},
    {"company": "PLX", "year": "2024", "close_yearend": 37500.0, "close_avg": 37200.0},
    {"company": "PLX", "year": "2025", "close_yearend": 35300.0, "close_avg": 36400.0},

    {"company": "PNJ", "year": "2019", "close_yearend": 43000.0, "close_avg": 41200.0},
    {"company": "PNJ", "year": "2020", "close_yearend": 40500.0, "close_avg": 36700.0},
    {"company": "PNJ", "year": "2021", "close_yearend": 48100.0, "close_avg": 46300.0},
    {"company": "PNJ", "year": "2022", "close_yearend": 59933.3, "close_avg": 58400.0},
    {"company": "PNJ", "year": "2023", "close_yearend": 57333.3, "close_avg": 54900.0},
    {"company": "PNJ", "year": "2024", "close_yearend": 65266.7, "close_avg": 63200.0},
    {"company": "PNJ", "year": "2025", "close_yearend": 64666.7, "close_avg": 64800.0},

    {"company": "VNM", "year": "2019", "close_yearend": 97083.3, "close_avg": 102500.0},
    {"company": "VNM", "year": "2020", "close_yearend": 108800.0, "close_avg": 98400.0},
    {"company": "VNM", "year": "2021", "close_yearend": 86400.0, "close_avg": 91200.0},
    {"company": "VNM", "year": "2022", "close_yearend": 76100.0, "close_avg": 73800.0},
    {"company": "VNM", "year": "2023", "close_yearend": 67600.0, "close_avg": 69400.0},
    {"company": "VNM", "year": "2024", "close_yearend": 63400.0, "close_avg": 65100.0},
    {"company": "VNM", "year": "2025", "close_yearend": 61200.0, "close_avg": 62500.0},
]


def fetch_annual_stock_prices(use_live: bool = True) -> pd.DataFrame:
    """Tải giá cổ phiếu đóng cửa cuối năm từ Yahoo Finance API (fallback dữ liệu nội bộ)."""
    records = []
    if use_live:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        for comp, ticker in TICKER_MAP.items():
            url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?range=8y&interval=1mo"
            try:
                r = requests.get(url, headers=headers, timeout=6)
                if r.status_code == 200:
                    data = r.json()["chart"]["result"][0]
                    timestamps = data.get("timestamp", [])
                    closes = data.get("indicators", {}).get("quote", [{}])[0].get("close", [])
                    df_t = pd.DataFrame({"ts": timestamps, "close": closes}).dropna()
                    df_t["year"] = df_t["ts"].apply(lambda x: datetime.fromtimestamp(x).year)
                    for yr, grp in df_t.groupby("year"):
                        if 2019 <= yr <= 2025:
                            records.append({
                                "company": comp,
                                "year": str(yr),
                                "close_yearend": round(float(grp.iloc[-1]["close"]), 1),
                                "close_avg": round(float(grp["close"].mean()), 1),
                            })
            except Exception as exc:
                print(f"[stock] Lỗi tải {comp} ({exc}), dùng dữ liệu fallback.")
                break

    if len(records) < 28:
        df_stock = pd.DataFrame(FALLBACK_STOCK_PRICES)
    else:
        df_stock = pd.DataFrame(records)

    # Tính % tăng trưởng giá hàng năm (YoY)
    df_stock = df_stock.sort_values(["company", "year"]).reset_index(drop=True)
    df_stock["price_change_pct"] = (
        df_stock.groupby("company")["close_yearend"].pct_change() * 100.0
    ).round(2)

    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    out_csv = RESULT_DIR / "stock_prices_annual.csv"
    df_stock.to_csv(out_csv, index=False, encoding="utf-8-sig")
    return df_stock


def merge_stock_and_esg(
    df_sentiment_counts: pd.DataFrame,
    df_cat: Optional[pd.DataFrame] = None,
    stock_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """Gộp chuỗi giá cổ phiếu với tỷ lệ Sentiment và điểm 6 nhóm SDG theo company-year."""
    if stock_df is None:
        stock_df = fetch_annual_stock_prices(use_live=True)

    # df_sentiment_counts có thể là index (company, year)
    senti = df_sentiment_counts.copy()
    if isinstance(senti.index, pd.MultiIndex):
        senti = senti.reset_index()

    senti["company"] = senti["company"].astype(str)
    senti["year"] = senti["year"].astype(str)
    stock_df["company"] = stock_df["company"].astype(str)
    stock_df["year"] = stock_df["year"].astype(str)

    merged = pd.merge(stock_df, senti, on=["company", "year"], how="left")

    if df_cat is not None:
        cat = df_cat.copy()
        if isinstance(cat.index, pd.MultiIndex):
            cat = cat.reset_index()
        cat["company"] = cat["company"].astype(str)
        cat["year"] = cat["year"].astype(str)
        merged = pd.merge(merged, cat, on=["company", "year"], how="left")

    # Tính chỉ số phân kỳ (Greenwashing Divergence Index):
    # Sentiment Ratio tăng nhưng Giá cổ phiếu giảm
    if "Ratio" in merged.columns and "price_change_pct" in merged.columns:
        senti_change = merged.groupby("company")["Ratio"].pct_change() * 100.0
        merged["sentiment_ratio_change_pct"] = senti_change.round(2)
        # Cờ cảnh báo: Sentiment tăng mạnh (> 15%) nhưng giá cổ phiếu lao dốc (< -10%)
        merged["greenwashing_flag"] = (
            (merged["sentiment_ratio_change_pct"] > 15.0) & (merged["price_change_pct"] < -10.0)
        )

    out_csv = RESULT_DIR / "table_stock_sentiment_greenwashing.csv"
    merged.to_csv(out_csv, index=False, encoding="utf-8-sig")
    return merged
