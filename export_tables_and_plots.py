"""Xuất bảng CSV + PNG từ embedding đã lưu — không load model.

Dùng khi chưa mở Jupyter: python export_tables_and_plots.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from config import CATEGORY_ORDER, EMBED_DIR, FIGURE_DIR, GOAL_COLS, GOAL_TITLE_VI, RESULT_DIR
from src.aggregate import category_means, company_year_means, minmax_scale_goals, sentiment_ratio
from src.encode import cosine_from_normalized, mean_similarity_by_goal
from src.plots import plot_sentiment_by_company, plot_stock_vs_sentiment_greenwashing
from src.stock import build_stock_sentiment_view


def load_parts() -> pd.DataFrame:
    parts = []
    for lang in ("vi", "en"):
        r = EMBED_DIR / f"report_{lang}.npy"
        s = EMBED_DIR / f"sdg_{lang}.npy"
        if not (r.exists() and s.exists()):
            continue
        report_emb = np.load(r)
        sdg_emb = np.load(s)
        report_idx = pd.read_pickle(EMBED_DIR / f"report_index_{lang}.pkl")
        sdg_idx = pd.read_pickle(EMBED_DIR / f"sdg_index_{lang}.pkl")
        sim = cosine_from_normalized(report_emb, sdg_emb)
        goals = mean_similarity_by_goal(sim, sdg_idx["goalnum"], GOAL_COLS)
        parts.append(pd.concat([report_idx.reset_index(drop=True), goals.reset_index(drop=True)], axis=1))
        print(f"[{lang}] sim {sim.shape} min={sim.min():.3f} max={sim.max():.3f}")
    if not parts:
        raise SystemExit("Chưa có embedding. Chạy python run_pipeline.py --encode-only")
    return pd.concat(parts, ignore_index=True)


def main() -> None:
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    sns.set_style("whitegrid")

    df = load_parts()
    df_s = minmax_scale_goals(df)
    df_comp = company_year_means(df_s)
    tbl = df_comp[GOAL_COLS].round(2)
    tbl.to_csv(RESULT_DIR / "table_company_year_17sdg.csv", encoding="utf-8-sig")
    df_cat = category_means(df_comp)
    df_cat.round(2).to_csv(RESULT_DIR / "table_company_year_6cat.csv", encoding="utf-8-sig")

    companies = sorted(df_s["company"].dropna().unique())
    n = max(len(companies), 1)
    fig, axes = plt.subplots(n, 1, figsize=(12, 4 * n), squeeze=False)
    for ax, company in zip(axes.ravel(), companies):
        sub = df_cat.loc[company].sort_index()
        for col in CATEGORY_ORDER:
            if col in sub.columns:
                ax.plot(sub.index.astype(str), sub[col], marker="o", linewidth=2, label=col)
        ax.set_title(company)
        ax.grid(True)
        ax.legend(loc="best", fontsize=8, ncol=3)
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "trends_6categories.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    stacked = pd.concat([df_s[c] for c in GOAL_COLS], ignore_index=True)
    fig, ax = plt.subplots(figsize=(12, 5))
    sns.histplot(stacked, color="indianred", alpha=0.35, kde=True, stat="density", binwidth=2, ax=ax)
    ax.set_xlim(0, 100)
    ax.set_title("Phân phối điểm tương đồng SDG (0–100)")
    fig.savefig(FIGURE_DIR / "similarity_hist.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    heat = df_cat.copy()
    heat.index = [f"{c}-{y}" for c, y in heat.index]
    fig, ax = plt.subplots(figsize=(10, max(4, 0.35 * len(heat))))
    sns.heatmap(heat[CATEGORY_ORDER], annot=True, fmt=".1f", cmap="YlOrRd", ax=ax)
    ax.set_title("Điểm 6 nhóm SDG (0–100)")
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "heatmap_6cat.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    both = RESULT_DIR / "result_both.pkl"
    if both.exists():
        df_both = pd.read_pickle(both)
        fig, ax = plt.subplots(figsize=(12, 5))
        sns.histplot(df_both["score"].dropna(), color="darkblue", alpha=0.35, ax=ax)
        ax.set_title("Phân phối polarity sentiment")
        fig.savefig(FIGURE_DIR / "sentiment_hist.png", dpi=150, bbox_inches="tight")
        plt.close(fig)

        counts = sentiment_ratio(df_both.dropna(subset=["label"]))
        counts.to_csv(RESULT_DIR / "table_sentiment_counts.csv", encoding="utf-8-sig")

        ratio = counts["Ratio"].unstack(level=0)
        fig, ax = plt.subplots(figsize=(12, 6))
        ratio.sort_index().plot(ax=ax, marker="o", linewidth=2)
        ax.set_title("Positive / Negative theo năm")
        ax.grid(True)
        fig.savefig(FIGURE_DIR / "sentiment_ratio.png", dpi=150, bbox_inches="tight")
        plt.close(fig)

        # 1. Vẽ cơ cấu Sentiment chi tiết từng công ty
        plot_sentiment_by_company(counts)

        monthly, _annual, corr = build_stock_sentiment_view(counts, use_live=True)
        plot_stock_vs_sentiment_greenwashing(monthly, corr)

    print("tables ->", RESULT_DIR)
    print("figures ->", FIGURE_DIR)


if __name__ == "__main__":
    main()
