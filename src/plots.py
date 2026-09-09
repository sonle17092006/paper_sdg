"""Vẽ phân phối, xu hướng 6 nhóm SDG, tỷ lệ sentiment — bám notebook 04–05."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

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


def plot_stock_vs_sentiment_greenwashing(
    df_merged: pd.DataFrame, name: str = "stock_vs_sentiment_greenwashing.png"
) -> Path:
    """Lồng ghép Giá cổ phiếu (VND) và Sentiment Ratio / SDG Score để đối chiếu dấu hiệu Tẩy xanh (Greenwashing)."""
    companies = sorted(df_merged["company"].unique())
    n = len(companies)
    fig, axes = plt.subplots(n, 1, figsize=(13, 5 * max(n, 1)), squeeze=False)

    for ax, comp in zip(axes.ravel(), companies):
        sub = df_merged[df_merged["company"] == comp].sort_values("year").copy()
        years = sub["year"].astype(str).tolist()
        x = np.arange(len(years))

        # Trục Y1 (Trái): Giá cổ phiếu
        color_stock = "#1f77b4"
        line1 = ax.plot(
            x, sub["close_yearend"], color=color_stock, marker="s", linewidth=2.5, label="Giá cổ phiếu cuối năm (VND)"
        )
        ax.set_ylabel("Giá cổ phiếu (VND)", color=color_stock, fontsize=11, fontweight="bold")
        ax.tick_params(axis="y", labelcolor=color_stock)
        ax.set_xticks(x)
        ax.set_xticklabels(years)
        ax.grid(True, linestyle="--", alpha=0.5)

        # Trục Y2 (Phải): Sentiment Ratio & Environments SDG
        ax2 = ax.twinx()
        color_senti = "#d95f02"
        color_env = "#2ca02c"
        line2 = ax2.plot(
            x, sub["Ratio"], color=color_senti, marker="o", linestyle="--", linewidth=2.0, label="Tỷ lệ Sentiment (Pos/Neg)"
        )

        lines = line1 + line2
        if "Environments" in sub.columns:
            line3 = ax2.plot(
                x,
                sub["Environments"],
                color=color_env,
                marker="^",
                linestyle=":",
                linewidth=2.0,
                label="Điểm SDG Môi trường (0-100)",
            )
            lines += line3

        ax2.set_ylabel("Sentiment Ratio / Điểm SDG", color=color_senti, fontsize=11, fontweight="bold")
        ax2.tick_params(axis="y", labelcolor=color_senti)
        ax2.grid(False)

        # Đánh dấu cờ cảnh báo Greenwashing nếu có
        if "greenwashing_flag" in sub.columns and sub["greenwashing_flag"].any():
            for idx_f in sub[sub["greenwashing_flag"] == True].index:
                pos_x = list(sub.index).index(idx_f)
                ax.axvspan(pos_x - 0.25, pos_x + 0.25, color="red", alpha=0.15)

        labels = [l.get_label() for l in lines]
        ax.legend(lines, labels, loc="upper left", fontsize=8.5)
        ax.set_title(
            f"Đối chiếu Cổ phiếu & Báo cáo PTBV - {comp} (Dấu hiệu Phân kỳ / Tẩy xanh)",
            fontsize=13,
            fontweight="bold",
        )

    fig.tight_layout()
    return _save(fig, name)
