"""Vẽ phân phối, xu hướng 6 nhóm SDG, tỷ lệ sentiment — bám notebook 04–05."""

from __future__ import annotations

from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from matplotlib.patches import Patch

from config import CATEGORY_ORDER, FIGURE_DIR, GOAL_COLS


def _save(fig, name: str) -> Path:
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    path = FIGURE_DIR / name
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def plot_sentiment_hist(df: pd.DataFrame, name: str = "sentiment_hist.png") -> Path:
    fig, ax = plt.subplots(figsize=(12, 5))
    sns.set_style("whitegrid")
    sns.histplot(df["score"], color="darkblue", alpha=0.3, edgecolor="darkslateblue", ax=ax)
    ax.set_title("Phân phối polarity sentiment (0 = tiêu cực, 1 = tích cực)")
    ax.set_xlabel("score")
    return _save(fig, name)


def plot_similarity_hist(df: pd.DataFrame, name: str = "similarity_hist.png") -> Path:
    stacked = pd.concat([df[c] for c in GOAL_COLS if c in df.columns], ignore_index=True)
    fig, ax = plt.subplots(figsize=(12, 5))
    sns.set_style("whitegrid")
    sns.histplot(stacked, color="indianred", alpha=0.3, kde=True, stat="density", binwidth=2, ax=ax)
    ax.set_xlim(0, 100)
    ax.set_title("Phân phối điểm tương đồng SDG (đã kéo 0–100)")
    ax.set_xlabel("similarity")
    return _save(fig, name)


def plot_category_trends(
    df_cat: pd.DataFrame,
    company: str,
    years: list[str],
    name: str | None = None,
) -> Path:
    if (company not in df_cat.index.get_level_values(0)) and company not in getattr(df_cat.index, "levels", [[]])[0]:
        raise KeyError(company)
    sub = df_cat.loc[company]
    fig, ax = plt.subplots(figsize=(12, 6))
    for col in CATEGORY_ORDER:
        if col in sub.columns:
            ax.plot(sub.index.astype(str), sub[col], marker="o", linewidth=2, label=col)
    ax.set_title(company)
    ax.set_xlabel("")
    ax.grid(True)
    ax.legend(loc="best", fontsize=8)
    fname = name or f"trend_{company}.png"
    return _save(fig, fname)


def plot_sentiment_ratio(ratio_by_year: pd.DataFrame, name: str = "sentiment_ratio.png") -> Path:
    fig, ax = plt.subplots(figsize=(12, 6))
    ratio_by_year.plot(ax=ax, marker="o", linewidth=2)
    ax.set_title("Tỷ lệ câu Positive / Negative theo năm")
    ax.set_xlabel("")
    ax.grid(True)
    ax.set_ylim(bottom=0)
    return _save(fig, name)


def plot_sentiment_by_company(df_counts: pd.DataFrame, name: str = "sentiment_by_company.png") -> Path:
    """Vẽ cơ cấu sentiment (Positive, Neutral, Negative) riêng biệt cho từng công ty."""
    df = df_counts.reset_index() if isinstance(df_counts.index, pd.MultiIndex) else df_counts.copy()
    companies = sorted(df["company"].unique())
    n = len(companies)
    fig, axes = plt.subplots(n, 1, figsize=(12, 4.5 * max(n, 1)), squeeze=False)

    colors = {"Positive": "#2ca02c", "Neutral": "#7f7f7f", "Negative": "#d62728"}

    for ax, comp in zip(axes.ravel(), companies):
        sub = df[df["company"] == comp].sort_values("year").copy()
        years = sub["year"].astype(str).tolist()
        x = np.arange(len(years))
        width = 0.55

        tot = sub[["Positive", "Neutral", "Negative"]].sum(axis=1).replace(0, 1)
        pos_pct = (sub["Positive"] / tot) * 100
        neu_pct = (sub["Neutral"] / tot) * 100
        neg_pct = (sub["Negative"] / tot) * 100

        ax.bar(x, pos_pct, width, label="Positive %", color=colors["Positive"], alpha=0.8)
        ax.bar(x, neu_pct, width, bottom=pos_pct, label="Neutral %", color=colors["Neutral"], alpha=0.7)
        ax.bar(x, neg_pct, width, bottom=pos_pct + neu_pct, label="Negative %", color=colors["Negative"], alpha=0.8)

        ax2 = ax.twinx()
        ax2.plot(x, sub["Ratio"], color="black", marker="o", linewidth=2.5, label="Ratio (Pos/Neg)")
        ax2.set_ylabel("Pos / Neg Ratio", fontsize=10, fontweight="bold")
        ax2.grid(False)

        ax.set_title(f"Cơ cấu Sentiment từng năm - Công ty {comp}", fontsize=13, fontweight="bold")
        ax.set_xticks(x)
        ax.set_xticklabels(years)
        ax.set_ylabel("Cơ cấu câu (%)", fontsize=10)
        ax.set_ylim(0, 100)

        h1, l1 = ax.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        ax.legend(h1 + h2, l1 + l2, loc="upper left", fontsize=8.5, ncol=4)

    fig.tight_layout()
    return _save(fig, name)

def _yearly_sentiment_peaks(sub: pd.DataFrame) -> pd.DataFrame:
    """Một đỉnh Pos/Neg giữa mỗi năm — nối lại thành ziczac/ngọn núi."""
    s = sub.dropna(subset=["ratio_idx"]).copy()
    if s.empty:
        return s
    years = s.groupby(s["date"].dt.year, sort=True)["ratio_idx"].first()
    return pd.DataFrame(
        {
            "date": [pd.Timestamp(year=int(y), month=7, day=1) for y in years.index],
            "ratio_idx": years.to_numpy(),
        }
    )


def plot_stock_vs_sentiment_greenwashing(
    df_monthly: pd.DataFrame,
    corr_df: pd.DataFrame | None = None,
    name: str = "stock_vs_sentiment_greenwashing.png",
) -> Path:
    """Giá tháng (đường) + Pos/Neg (ziczac đỉnh năm, tô như núi). Cùng index 100."""
    companies = sorted(df_monthly["company"].unique())
    n = len(companies)
    fig, axes = plt.subplots(n, 1, figsize=(14, 4.4 * max(n, 1)), squeeze=False)
    corr_map = {}
    if corr_df is not None and "company" in corr_df.columns:
        corr_map = corr_df.set_index("company").to_dict(orient="index")

    color_stock = "#1f77b4"
    color_senti = "#d95f02"
    band = {"Tẩy xanh": "#d62728", "Trung thực": "#2ca02c"}

    for ax, comp in zip(axes.ravel(), companies):
        sub = df_monthly[df_monthly["company"] == comp].sort_values("date").copy()
        sub["date"] = pd.to_datetime(sub["date"])
        peaks = _yearly_sentiment_peaks(sub)

        if "phan_ung" in sub.columns:
            for year, g in sub.groupby("year"):
                lab = str(g["phan_ung"].dropna().iloc[0]) if g["phan_ung"].notna().any() else "None"
                if lab in band:
                    ax.axvspan(
                        g["date"].min(),
                        g["date"].max() + pd.offsets.MonthBegin(1),
                        color=band[lab],
                        alpha=0.13,
                        zorder=0,
                    )

        both_parts = [sub["price_idx"]]
        if not peaks.empty:
            both_parts.append(peaks["ratio_idx"])
        both = pd.concat(both_parts, ignore_index=True).dropna()
        if not both.empty:
            lo, hi = float(both.min()), float(both.max())
            pad = max((hi - lo) * 0.12, 8.0)
            ax.set_ylim(lo - pad, hi + pad)
        ylo = ax.get_ylim()[0]

        if not peaks.empty:
            ax.fill_between(
                peaks["date"],
                ylo,
                peaks["ratio_idx"],
                color=color_senti,
                alpha=0.28,
                zorder=1,
            )
            ax.plot(
                peaks["date"],
                peaks["ratio_idx"],
                color=color_senti,
                marker="^",
                markersize=8,
                linewidth=2.4,
                label="Pos/Neg (đỉnh năm, index 100)",
                zorder=4,
            )

        ax.plot(
            sub["date"],
            sub["price_idx"],
            color=color_stock,
            linewidth=2.0,
            label="Giá tháng (index 100)",
            zorder=5,
        )
        ax.axhline(100, color="0.5", linewidth=0.8, linestyle=":")

        ax.set_ylabel("Chỉ số (tháng đầu = 100)", fontsize=10, fontweight="bold")
        ax.grid(True, linestyle="--", alpha=0.4)
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        ax.xaxis.set_minor_locator(mdates.MonthLocator(bymonth=[7]))

        info = corr_map.get(comp, {})
        r = info.get("r_pearson")
        p = info.get("p_pearson")
        nn = info.get("n")
        doc = info.get("doc", "")
        rtxt = f"r(Δgiá, ΔPos/Neg) = {r:.2f}" if isinstance(r, (int, float)) and np.isfinite(r) else "r = n/a"
        ptxt = f", p = {p:.2f}" if isinstance(p, (int, float)) and np.isfinite(p) else ""
        ntxt = f", n = {nn}" if nn is not None else ""
        ax.set_title(f"{comp}:  {rtxt}{ptxt}{ntxt}   → {doc or 'None'}", fontsize=12, fontweight="bold")

        handles = [
            plt.Line2D([0], [0], color=color_stock, lw=2, label="Giá tháng (index 100)"),
            plt.Line2D([0], [0], color=color_senti, marker="^", lw=2.4, label="Pos/Neg (ziczac đỉnh năm)"),
            Patch(facecolor=color_senti, alpha=0.28, label="Nền Pos/Neg (ngọn núi)"),
            Patch(facecolor=band["Tẩy xanh"], alpha=0.25, label="Năm giá tụt: Tẩy xanh"),
            Patch(facecolor=band["Trung thực"], alpha=0.25, label="Năm giá tụt: Trung thực"),
        ]
        ax.legend(handles=handles, loc="upper left", fontsize=8, ncol=2)

    fig.tight_layout()
    return _save(fig, name)
