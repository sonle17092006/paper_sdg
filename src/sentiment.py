"""Sentiment theo paper: DistilBERT SST-2 cho EN; VisoBERT cho VI.

Paper đảo điểm NEGATIVE: score := 1 - score để trục 0–1 hướng dương.
Notebook gốc chỉ append label khi POSITIVE (lỗi lệch độ dài) — ở đây giữ mọi nhãn.
Tiếng Việt có thêm Neutral; Neutral gán polarity 0.5.
"""

from __future__ import annotations

from typing import Any

import torch
from transformers import AutoModelForSequenceClassification, AutoTokenizer, pipeline

from config import MAX_SEQ_LEN, SENTIMENT_BATCH

POS_LABELS = {"POSITIVE", "POS", "LABEL_1", "5 STARS", "4 STARS"}
NEG_LABELS = {"NEGATIVE", "NEG", "LABEL_0", "1 STAR", "2 STARS"}
NEU_LABELS = {"NEUTRAL", "NEU", "LABEL_2", "3 STARS"}


def load_sentiment_pipeline(model_dir: str, max_length: int = 256):
    tokenizer = AutoTokenizer.from_pretrained(model_dir, use_fast=False)
    model = AutoModelForSequenceClassification.from_pretrained(model_dir)
    device = 0 if torch.cuda.is_available() else -1
    return pipeline(
        "sentiment-analysis",
        model=model,
        tokenizer=tokenizer,
        device=device,
        truncation=True,
        max_length=max_length,
    )


def _canon_label(raw: str) -> str:
    u = raw.upper().replace("-", "_").strip()
    if u in POS_LABELS or "POS" in u:
        return "Positive"
    if u in NEG_LABELS or "NEG" in u:
        return "Negative"
    if u in NEU_LABELS or "NEU" in u:
        return "Neutral"
    return raw.title()


def score_to_polarity(label: str, score: float) -> float:
    """Cùng quy ước paper: càng gần 1 càng tích cực."""
    if label == "Negative":
        return 1.0 - float(score)
    if label == "Neutral":
        return 0.5
    return float(score)


def _segment_vi(text: str) -> str:
    """PhoBERT cần tách tiếng. pyvi nếu có; không thì giữ nguyên."""
    try:
        from pyvi.ViTokenizer import tokenize

        return tokenize(text)
    except Exception:
        return text


def analyze_sentiment(
    classifier,
    sentences: list[str],
    batch_size: int = SENTIMENT_BATCH,
    max_length: int = 256,
    segment_vi: bool = False,
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    n = len(sentences)
    for i in range(0, n, batch_size):
        chunk = sentences[i : i + batch_size]
        if segment_vi:
            chunk = [_segment_vi(s) for s in chunk]
        raw = classifier(chunk, truncation=True, max_length=max_length, batch_size=min(batch_size, len(chunk)))
        if isinstance(raw, dict):
            raw = [raw]
        for item in raw:
            label = _canon_label(str(item.get("label", "")))
            conf = float(item.get("score", 0.0))
            results.append(
                {
                    "label": label,
                    "confidence": round(conf, 4),
                    "score": round(score_to_polarity(label, conf), 4),
                }
            )
        print(f"  sentiment {min(i + batch_size, n)}/{n}", flush=True)
    return results
