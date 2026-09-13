"""Giá cổ phiếu tháng + đối chiếu Pos/Neg khi giá tụt (tẩy xanh / trung thực / None)."""

from __future__ import annotations

from typing import Optional

import numpy as np
import pandas as pd
import requests
from scipy.stats import pearsonr, spearmanr

from config import RESULT_DIR

TICKER_MAP = {
    "BVH": "BVH.VN",
    "PAN": "PAN.VN",
    "PLX": "PLX.VN",
    "PNJ": "PNJ.VN",
    "SSI": "SSI.VN",
    "VCS": "VCS",
    "VNM": "VNM.VN",
}

YEAR_MIN = 2020
YEAR_MAX = 2025
PRICE_DROP_PCT = -10.0
RATIO_UP_PCT = 15.0

# Năm-cuối (VND, điều chỉnh) — chỉ dùng khi cả VNDirect và Yahoo không truy cập được và chưa có cache tháng
FALLBACK_STOCK_PRICES = [
    {"company": "BVH", "year": "2020", "close_yearend": 57129.0, "close_avg": 43037.1},
    {"company": "BVH", "year": "2021", "close_yearend": 49161.0, "close_avg": 50122.2},
    {"company": "BVH", "year": "2022", "close_yearend": 43618.0, "close_avg": 48711.2},
    {"company": "BVH", "year": "2023", "close_yearend": 37870.0, "close_avg": 41703.6},
    {"company": "BVH", "year": "2024", "close_yearend": 49747.0, "close_avg": 42496.0},
    {"company": "BVH", "year": "2025", "close_yearend": 56800.0, "close_avg": 51969.8},
    {"company": "PAN", "year": "2020", "close_yearend": 24839.0, "close_avg": 15491.6},
    {"company": "PAN", "year": "2021", "close_yearend": 28136.0, "close_avg": 21441.4},
    {"company": "PAN", "year": "2022", "close_yearend": 10925.0, "close_avg": 17176.7},
    {"company": "PAN", "year": "2023", "close_yearend": 14762.0, "close_avg": 13921.7},
    {"company": "PAN", "year": "2024", "close_yearend": 17723.0, "close_avg": 17168.9},
    {"company": "PAN", "year": "2025", "close_yearend": 20674.0, "close_avg": 21453.5},
    {"company": "PLX", "year": "2020", "close_yearend": 45324.0, "close_avg": 38692.1},
    {"company": "PLX", "year": "2021", "close_yearend": 45680.0, "close_avg": 44579.4},
    {"company": "PLX", "year": "2022", "close_yearend": 28005.0, "close_avg": 36328.8},
    {"company": "PLX", "year": "2023", "close_yearend": 31015.0, "close_avg": 32786.7},
    {"company": "PLX", "year": "2024", "close_yearend": 35030.0, "close_avg": 37174.6},
    {"company": "PLX", "year": "2025", "close_yearend": 34079.0, "close_avg": 34783.6},
    {"company": "PNJ", "year": "2020", "close_yearend": 36207.0, "close_avg": 29103.8},
    {"company": "PNJ", "year": "2021", "close_yearend": 43690.0, "close_avg": 42466.7},
    {"company": "PNJ", "year": "2022", "close_yearend": 55443.0, "close_avg": 51265.5},
    {"company": "PNJ", "year": "2023", "close_yearend": 54414.0, "close_avg": 49448.0},
    {"company": "PNJ", "year": "2024", "close_yearend": 63208.0, "close_avg": 61098.2},
    {"company": "PNJ", "year": "2025", "close_yearend": 64049.0, "close_avg": 56987.8},
    {"company": "SSI", "year": "2020", "close_yearend": 10719.0, "close_avg": 5168.8},
    {"company": "SSI", "year": "2021", "close_yearend": 24957.0, "close_avg": 16904.7},
    {"company": "SSI", "year": "2022", "close_yearend": 9440.0, "close_avg": 13964.8},
    {"company": "SSI", "year": "2023", "close_yearend": 18182.0, "close_avg": 14436.8},
    {"company": "SSI", "year": "2024", "close_yearend": 18496.0, "close_avg": 19118.1},
    {"company": "SSI", "year": "2025", "close_yearend": 24179.0, "close_avg": 21741.1},
    {"company": "VCS", "year": "2020", "close_yearend": 59488.0, "close_avg": 44823.0},
    {"company": "VCS", "year": "2021", "close_yearend": 82728.0, "close_avg": 74277.2},
    {"company": "VCS", "year": "2022", "close_yearend": 43463.0, "close_avg": 58410.8},
    {"company": "VCS", "year": "2023", "close_yearend": 45535.0, "close_avg": 44954.5},
    {"company": "VCS", "year": "2024", "close_yearend": 55571.0, "close_avg": 55614.8},
    {"company": "VCS", "year": "2025", "close_yearend": 40579.0, "close_avg": 45248.0},
    {"company": "VNM", "year": "2020", "close_yearend": 79366.0, "close_avg": 68862.1},
    {"company": "VNM", "year": "2021", "close_yearend": 65510.0, "close_avg": 68837.7},
    {"company": "VNM", "year": "2022", "close_yearend": 61739.0, "close_avg": 60091.5},
    {"company": "VNM", "year": "2023", "close_yearend": 57090.0, "close_avg": 59455.8},
    {"company": "VNM", "year": "2024", "close_yearend": 56564.0, "close_avg": 58466.2},
    {"company": "VNM", "year": "2025", "close_yearend": 59260.0, "close_avg": 55478.5},
]


def _vndirect_monthly(symbol: str) -> pd.DataFrame:
    """Lấy dữ liệu giá tháng từ VNDirect DChart API (hỗ trợ cả HOSE và HNX như VCS)."""
    url = f"https://dchart-api.vndirect.com.vn/dchart/history?resolution=D&symbol={symbol}&from=1577836800&to=1767225600"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    r = requests.get(url, headers=headers, timeout=12)
    r.raise_for_status()
    data = r.json()
    if not data or "t" not in data or "c" not in data or not data["t"]:
        raise ValueError(f"Empty data from VNDirect for {symbol}")
    df = pd.DataFrame({"ts": data["t"], "close": data["c"]})
    df["date"] = pd.to_datetime(df["ts"], unit="s").dt.tz_localize(None)
    df["date"] = df["date"].dt.to_period("M").dt.to_timestamp()
    df["close"] = df["close"] * 1000.0  # VNDirect trả giá nghìn VND -> đổi sang VND
    m = df.groupby("date")["close"].last().reset_index()
    return m


def _yahoo_monthly(ticker: str) -> pd.DataFrame:
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?range=8y&interval=1mo"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    r = requests.get(url, headers=headers, timeout=12)
    r.raise_for_status()
    result = r.json()["chart"]["result"][0]
    ts = result.get("timestamp") or []
    ind = result.get("indicators") or {}
    adj = (ind.get("adjclose") or [{}])[0].get("adjclose")
    raw = (ind.get("quote") or [{}])[0].get("close") or []
    closes = adj if adj else raw
    df = pd.DataFrame({"ts": ts, "close": closes}).dropna()
    df["date"] = pd.to_datetime(df["ts"], unit="s").dt.tz_localize(None)
    df["date"] = df["date"].dt.to_period("M").dt.to_timestamp()
    df = df.drop(columns=["ts"]).drop_duplicates("date")
    return df


def _monthly_from_annual_fallback() -> pd.DataFrame:
    annual = pd.DataFrame(FALLBACK_STOCK_PRICES)
    rows: list[dict] = []
    for comp, g in annual.groupby("company"):
        g = g.sort_values("year")
        prices = {int(y): float(p) for y, p in zip(g["year"], g["close_yearend"])}
        years = sorted(prices)
        prev = prices[years[0]]
        for y in years:
            end = prices[y]
            for m in range(1, 13):
                close = prev + (end - prev) * (m / 12.0)
                rows.append(
                    {
                        "company": comp,
                        "date": pd.Timestamp(year=y, month=m, day=1),
                        "year": str(y),
                        "month": m,
                        "close": round(close, 1),
                    }
                )
            prev = end
    return pd.DataFrame(rows)


def _drop_wild_last_bar(df: pd.DataFrame) -> pd.DataFrame:
    """Yahoo đôi khi trả thanh tháng cuối lệch (vd. PLX 12/2025 +67%). Bỏ nếu |MoM| > 35% và gấp 4 lần biến động gần đó."""
    parts = []
    for _, g in df.groupby("company", sort=False):
        g = g.sort_values("date").copy()
        if len(g) >= 8:
            mom = g["close"].pct_change()
            last = mom.iloc[-1]
            typical = mom.iloc[-7:-1].abs().median()
            if pd.notna(last) and abs(last) > 0.35 and abs(last) > 4 * max(float(typical or 0), 0.01):
                print(f"[stock] drop last bar {g['company'].iloc[-1]} {g['date'].iloc[-1].date()} MoM={last:.0%}")
                g = g.iloc[:-1]
        parts.append(g)
    return pd.concat(parts, ignore_index=True)


def fetch_monthly_stock_prices(use_live: bool = True) -> pd.DataFrame:
    """Giá đóng cửa từng tháng 2019–2025. Cache CSV; ưu tiên VNDirect -> Yahoo -> cache -> fallback."""
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    cache = RESULT_DIR / "stock_prices_monthly.csv"
    frames: list[pd.DataFrame] = []
    if use_live:
        for comp, ticker in TICKER_MAP.items():
            df_t = None
            # 1. Thử VNDirect trước (hỗ trợ cả HNX và HOSE)
            try:
                df_t = _vndirect_monthly(comp)
            except Exception as e_vnd:
                # 2. Dự phòng sang Yahoo Finance
                try:
                    df_t = _yahoo_monthly(ticker)
                except Exception as e_yah:
                    print(f"[stock] fetch failed {comp} (VNDirect: {e_vnd}; Yahoo: {e_yah})")
            if df_t is not None and not df_t.empty:
                df_t["company"] = comp
                frames.append(df_t)
            else:
                frames = []
                break

    if frames and len(frames) == len(TICKER_MAP):
        df = pd.concat(frames, ignore_index=True)
        df["year"] = df["date"].dt.year.astype(str)
        df["month"] = df["date"].dt.month
        df = df[(df["date"].dt.year >= YEAR_MIN) & (df["date"].dt.year <= YEAR_MAX)]
        df = df.sort_values(["company", "date"]).reset_index(drop=True)
        df = _drop_wild_last_bar(df)
        df[["company", "date", "year", "month", "close"]].to_csv(
            cache, index=False, encoding="utf-8-sig"
        )
        return df[["company", "date", "year", "month", "close"]]

    if cache.exists():
        print("[stock] using monthly cache")
        df = pd.read_csv(cache, parse_dates=["date"])
        df["year"] = df["year"].astype(str)
        df = df[(df["date"].dt.year >= YEAR_MIN) & (df["date"].dt.year <= YEAR_MAX)]
        return df

    print("[stock] interpolate monthly from year-end fallback")
    df = _monthly_from_annual_fallback()
    df.to_csv(cache, index=False, encoding="utf-8-sig")
    return df


def annual_from_monthly(monthly: pd.DataFrame) -> pd.DataFrame:
    g = monthly.sort_values("date").groupby(["company", "year"], sort=False)
    out = g.agg(close_yearend=("close", "last"), close_avg=("close", "mean")).reset_index()
    out["close_yearend"] = out["close_yearend"].round(1)
    out["close_avg"] = out["close_avg"].round(1)
    out = out.sort_values(["company", "year"]).reset_index(drop=True)
    out["price_change_pct"] = (
        out.groupby("company")["close_yearend"].pct_change() * 100.0
    ).round(2)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    out.to_csv(RESULT_DIR / "stock_prices_annual.csv", index=False, encoding="utf-8-sig")
    return out


def fetch_annual_stock_prices(use_live: bool = True) -> pd.DataFrame:
    """Giữ API cũ: năm được gộp từ chuỗi tháng."""
    return annual_from_monthly(fetch_monthly_stock_prices(use_live=use_live))


def _as_senti_table(df_sentiment_counts: pd.DataFrame) -> pd.DataFrame:
    senti = df_sentiment_counts.copy()
    if isinstance(senti.index, pd.MultiIndex):
        senti = senti.reset_index()
    senti["company"] = senti["company"].astype(str)
    senti["year"] = senti["year"].astype(str)
    return senti


def classify_when_price_drops(df: pd.DataFrame) -> pd.DataFrame:
    """Khi giá năm giảm >10%: Tẩy xanh (Pos/Neg tăng >15%), Trung thực (Pos/Neg giảm), else None."""
    out = df.copy()
    price = pd.to_numeric(out.get("price_change_pct"), errors="coerce")
    ratio_ch = pd.to_numeric(out.get("sentiment_ratio_change_pct"), errors="coerce")
    down = price < PRICE_DROP_PCT
    out["gia_tut"] = down.fillna(False)
    label = np.full(len(out), "None", dtype=object)
    label[down.fillna(False) & (ratio_ch > RATIO_UP_PCT)] = "Tẩy xanh"
    label[down.fillna(False) & (ratio_ch < 0)] = "Trung thực"
    out["phan_ung"] = label
    out["greenwashing_flag"] = out["phan_ung"] == "Tẩy xanh"
    return out


def _corr_row(name: str, x: pd.Series, y: pd.Series) -> dict:
    mask = x.notna() & y.notna()
    xx, yy = x[mask].astype(float), y[mask].astype(float)
    n = int(len(xx))
    if n < 3:
        return {
            "company": name,
            "n": n,
            "r_pearson": np.nan,
            "p_pearson": np.nan,
            "r_spearman": np.nan,
            "p_spearman": np.nan,
            "doc": "None",
        }
    rp, pp = pearsonr(xx, yy)
    rs, ps = spearmanr(xx, yy)
    if n < 4 or not np.isfinite(pp) or pp >= 0.10:
        doc = "None"
    elif rp <= -0.3:
        doc = "Thiên tẩy xanh"
    elif rp >= 0.3:
        doc = "Thiên trung thực"
    else:
        doc = "None"
    return {
        "company": name,
        "n": n,
        "r_pearson": round(float(rp), 3),
        "p_pearson": round(float(pp), 3),
        "r_spearman": round(float(rs), 3),
        "p_spearman": round(float(ps), 3),
        "doc": doc,
    }


def correlation_by_company(annual: pd.DataFrame) -> pd.DataFrame:
    """Pearson/Spearman giữa % đổi giá năm và % đổi Pos/Neg. n nhỏ → p thường không đủ tin."""
    xcol, ycol = "price_change_pct", "sentiment_ratio_change_pct"
    rows = []
    for comp, g in annual.groupby("company"):
        rows.append(_corr_row(str(comp), g[xcol], g[ycol]))
    rows.append(_corr_row("Tổng (pool)", annual[xcol], annual[ycol]))
    out = pd.DataFrame(rows)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    out.to_csv(RESULT_DIR / "table_stock_sentiment_correlation.csv", index=False, encoding="utf-8-sig")
    return out


def _indexed_monthly(monthly: pd.DataFrame, annual: pd.DataFrame) -> pd.DataFrame:
    """Gắn Pos/Neg (bậc thang theo năm) và kéo cả hai về index 100 = tháng đầu có dữ liệu."""
    keep = ["company", "year", "Ratio", "phan_ung", "price_change_pct", "sentiment_ratio_change_pct"]
    keep = [c for c in keep if c in annual.columns]
    m = monthly.merge(annual[keep], on=["company", "year"], how="left")
    m = m.sort_values(["company", "date"]).reset_index(drop=True)

    def _index_one(g: pd.DataFrame) -> pd.DataFrame:
        g = g.copy()
        base_p = g["close"].dropna()
        g["price_idx"] = np.nan
        if not base_p.empty and base_p.iloc[0] != 0:
            g["price_idx"] = g["close"] / base_p.iloc[0] * 100.0
        base_r = g["Ratio"].dropna()
        g["ratio_idx"] = np.nan
        if not base_r.empty and base_r.iloc[0] != 0:
            g["ratio_idx"] = g["Ratio"] / base_r.iloc[0] * 100.0
        return g

    parts = [_index_one(g) for _, g in m.groupby("company", sort=False)]
    return pd.concat(parts, ignore_index=True)


def merge_stock_and_esg(
    df_sentiment_counts: pd.DataFrame,
    df_cat: Optional[pd.DataFrame] = None,
    stock_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """Gộp giá năm với Pos/Neg, gắn nhãn khi giá tụt. df_cat không dùng (bỏ SDG môi trường)."""
    del df_cat
    if stock_df is None:
        stock_df = fetch_annual_stock_prices(use_live=True)

    senti = _as_senti_table(df_sentiment_counts)
    stock_df = stock_df.copy()
    stock_df["company"] = stock_df["company"].astype(str)
    stock_df["year"] = stock_df["year"].astype(str)
    merged = pd.merge(stock_df, senti, on=["company", "year"], how="left")

    if "Ratio" in merged.columns:
        merged["sentiment_ratio_change_pct"] = (
            merged.groupby("company")["Ratio"].pct_change() * 100.0
        ).round(2)
    merged = classify_when_price_drops(merged)

    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    merged.to_csv(RESULT_DIR / "table_stock_sentiment_greenwashing.csv", index=False, encoding="utf-8-sig")
    return merged


def build_stock_sentiment_view(
    df_sentiment_counts: pd.DataFrame,
    use_live: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """monthly (để vẽ), annual (nhãn năm), corr (Pearson/Spearman)."""
    monthly = fetch_monthly_stock_prices(use_live=use_live)
    annual_px = annual_from_monthly(monthly)
    annual = merge_stock_and_esg(df_sentiment_counts, df_cat=None, stock_df=annual_px)
    corr = correlation_by_company(annual)
    monthly_plot = _indexed_monthly(monthly, annual)
    monthly_plot.to_csv(RESULT_DIR / "stock_prices_monthly_with_ratio.csv", index=False, encoding="utf-8-sig")
    return monthly_plot, annual, corr
