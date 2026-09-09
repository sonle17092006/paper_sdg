"""Mã hóa câu thành vector và lưu .npy đã L2-normalize — notebook chỉ việc nhân ma trận."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
from sentence_transformers import SentenceTransformer

from config import ENCODE_BATCH, MAX_SEQ_LEN


def l2_normalize(mat: np.ndarray, eps: float = 1e-12) -> np.ndarray:
    norms = np.linalg.norm(mat, axis=1, keepdims=True)
    return mat / np.clip(norms, eps, None)


def encode_sentences(
    model: SentenceTransformer,
    sentences: list[str],
    batch_size: int = ENCODE_BATCH,
    show_progress: bool = True,
) -> np.ndarray:
    if not sentences:
        dim = model.get_sentence_embedding_dimension()
        return np.zeros((0, dim), dtype=np.float32)
    emb = model.encode(
        sentences,
        batch_size=batch_size,
        convert_to_numpy=True,
        show_progress_bar=show_progress,
        normalize_embeddings=True,
    )
    return np.asarray(emb, dtype=np.float32)


def save_embeddings(path: Path, emb: np.ndarray) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    np.save(path, emb)


def load_embeddings(path: Path) -> np.ndarray:
    return np.load(path)


def cosine_from_normalized(a: np.ndarray, b: np.ndarray) -> np.ndarray:
    """a: (n, d) đã L2-normalize, b: (m, d) đã L2-normalize → (n, m) cosine."""
    return a @ b.T


def mean_similarity_by_goal(
    sim: np.ndarray,
    sdg_goals: pd.Series,
    goal_cols: list[str],
) -> pd.DataFrame:
    """sim[i, j] = cosine(report_i, sdg_j). Trung bình theo goal như paper."""
    out = {}
    for goal in goal_cols:
        mask = (sdg_goals.to_numpy() == goal)
        if not mask.any():
            out[goal] = np.zeros(sim.shape[0], dtype=np.float32)
        else:
            out[goal] = sim[:, mask].mean(axis=1)
    return pd.DataFrame(out)
