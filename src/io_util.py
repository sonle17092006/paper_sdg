"""Lưu/đọc bảng: parquet nếu có engine, không thì pickle + csv (như paper)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def save_table(df: pd.DataFrame, parquet_path: Path) -> list[Path]:
    parquet_path.parent.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []
    pkl = parquet_path.with_suffix(".pkl")
    csv = parquet_path.with_suffix(".csv")
    try:
        df.to_parquet(parquet_path, index=False)
        saved.append(parquet_path)
    except Exception as exc:
        print(f"[io] parquet skip ({exc.__class__.__name__}), dùng pickle")
    df.to_pickle(pkl)
    saved.append(pkl)
    df.to_csv(csv, index=False)
    saved.append(csv)
    return saved


def load_table(parquet_path: Path) -> pd.DataFrame:
    candidates = [
        parquet_path,
        parquet_path.with_suffix(".pkl"),
        parquet_path.with_suffix(".csv"),
    ]
    last_err: Exception | None = None
    for path in candidates:
        if not path.exists():
            continue
        try:
            if path.suffix == ".parquet":
                return pd.read_parquet(path)
            if path.suffix == ".pkl":
                return pd.read_pickle(path)
            return pd.read_csv(path)
        except Exception as exc:
            last_err = exc
            continue
    raise FileNotFoundError(f"Không đọc được bảng {parquet_path} ({last_err})")
