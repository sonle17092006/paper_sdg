"""Gộp điểm similarity / sentiment theo công ty–năm, 17 SDG, 6 nhóm — như notebook 04–05."""

from __future__ import annotations

import pandas as pd

from config import CATEGORY_MAP, CATEGORY_ORDER, GOAL_COLS


def minmax_scale_goals(df: pd.DataFrame, goal_cols: list[str] | None = None) -> pd.DataFrame:
    goal_cols = goal_cols or GOAL_COLS
    block = df[goal_cols]
    vmin = float(block.min().min())
    vmax = float(block.max().max())
    span = vmax - vmin if vmax != vmin else 1.0
    scaled = (block - vmin) * 100.0 / span
    out = df.copy()
    out[goal_cols] = scaled
    out.attrs["sim_min"] = vmin
    out.attrs["sim_max"] = vmax
    return out


def company_year_means(df: pd.DataFrame) -> pd.DataFrame:
    num_cols = ["score", *GOAL_COLS]
    have = [c for c in num_cols if c in df.columns]
    g = df.groupby(["company", "year"], dropna=False)[have].mean(numeric_only=True)
    return g


def category_means(df_comp: pd.DataFrame) -> pd.DataFrame:
    """df_comp index (company, year), columns goal01..goal17 (đã scale)."""
    goals = [c for c in GOAL_COLS if c in df_comp.columns]
    tmp = df_comp[goals].T.copy()
    tmp["category"] = tmp.index.map(CATEGORY_MAP)
    cat = tmp.groupby("category").mean(numeric_only=True).reindex(CATEGORY_ORDER).T
    return cat


def sentiment_ratio(df: pd.DataFrame) -> pd.DataFrame:
    """Positive / Negative theo công ty–năm, giống paper. Neutral tách cột riêng."""
    counts = (
        df.groupby(["company", "year"])["label"]
        .value_counts()
        .unstack(fill_value=0)
    )
    for col in ("Positive", "Negative", "Neutral"):
        if col not in counts.columns:
            counts[col] = 0
    counts["Ratio"] = (counts["Positive"] / counts["Negative"].replace(0, pd.NA)).astype(float)
    return counts
