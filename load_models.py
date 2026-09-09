"""Tải model về thư mục models/ rồi load từ đĩa.

Lần sau chạy lại không cần tải HuggingFace.

    python load_models.py              # tải nếu thiếu, kiểm tra load
    python load_models.py --force      # tải lại
    python load_models.py --list       # chỉ liệt kê trạng thái

EN  (paper): all-MiniLM-L6-v2 + DistilBERT SST-2
VI  (thích ứng): keepitreal/vietnamese-sbert + 5CD-AI/Vietnamese-Sentiment-visobert
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config import HUB_MODELS, LOCAL_MODELS, MODEL_DIR
from src.progress import mark_step


def _ok(path: Path) -> bool:
    if not path.exists():
        return False
    has_st = (path / "modules.json").exists() or (path / "config_sentence_transformers.json").exists()
    has_hf = (path / "config.json").exists()
    return has_st or has_hf


def status() -> dict[str, dict]:
    rows = {}
    for name, hub in HUB_MODELS.items():
        local = LOCAL_MODELS[name]
        rows[name] = {
            "hub": hub,
            "local": str(local),
            "ready": _ok(local),
        }
    return rows


def download_one(name: str, force: bool = False) -> Path:
    dest = LOCAL_MODELS[name]
    dest.parent.mkdir(parents=True, exist_ok=True)
    if _ok(dest) and not force:
        print(f"[skip] {name} đã có tại {dest}")
        return dest

    hub = HUB_MODELS[name]
    print(f"[download] {name}  <-  {hub}")
    dest.mkdir(parents=True, exist_ok=True)

    if name.endswith("_sbert"):
        from sentence_transformers import SentenceTransformer

        model = SentenceTransformer(hub)
        model.save(str(dest))
        del model
    else:
        from transformers import AutoModelForSequenceClassification, AutoTokenizer

        tok = AutoTokenizer.from_pretrained(hub, use_fast=False)
        mdl = AutoModelForSequenceClassification.from_pretrained(hub)
        tok.save_pretrained(str(dest))
        mdl.save_pretrained(str(dest))
        del tok, mdl

    print(f"[saved] {dest}")
    return dest


def load_sbert(lang: str):
    """Load sentence-transformer từ đĩa. lang in {'en','vi'}."""
    from sentence_transformers import SentenceTransformer

    key = f"{lang}_sbert"
    path = LOCAL_MODELS[key]
    if not _ok(path):
        download_one(key)
    print(f"[load] {key} from {path}")
    return SentenceTransformer(str(path))


def load_sentiment(lang: str):
    """Load pipeline sentiment từ đĩa. lang in {'en','vi'}."""
    from src.sentiment import load_sentiment_pipeline
    from config import MAX_SEQ_LEN

    key = f"{lang}_sentiment"
    path = LOCAL_MODELS[key]
    if not _ok(path):
        download_one(key)
    print(f"[load] {key} from {path}")
    return load_sentiment_pipeline(str(path), max_length=MAX_SEQ_LEN[key])


def main() -> None:
    parser = argparse.ArgumentParser(description="Tải / kiểm tra model local")
    parser.add_argument("--force", action="store_true", help="tải lại dù đã có")
    parser.add_argument("--list", action="store_true", help="chỉ in trạng thái")
    parser.add_argument(
        "--only",
        nargs="*",
        default=None,
        help="chỉ các key: en_sbert en_sentiment vi_sbert vi_sentiment",
    )
    args = parser.parse_args()

    MODEL_DIR.mkdir(parents=True, exist_ok=True)
    st = status()
    print(json.dumps(st, ensure_ascii=False, indent=2))
    if args.list:
        return

    names = args.only or list(HUB_MODELS)
    for name in names:
        if name not in HUB_MODELS:
            raise SystemExit(f"unknown model key: {name}")
        download_one(name, force=args.force)

    mark_step(
        "load_models",
        status="done",
        models={k: v["ready"] or True for k, v in status().items()},
        note="Models cached under models/. Subsequent runs load from disk.",
    )
    print("==== models ready ====")


if __name__ == "__main__":
    main()
