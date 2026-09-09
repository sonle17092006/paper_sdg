"""Encode + cosine similarity + sentiment. Lưu vector để notebook không load model.

    python run_pipeline.py                 # full
    python run_pipeline.py --encode-only   # chỉ embedding + similarity
    python run_pipeline.py --sentiment-only
    python run_pipeline.py --lang vi       # chỉ nhánh tiếng Việt

Phụ thuộc: load_models.py đã tải model, process_data.py đã extract câu.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config import (
    EMBED_DIR,
    GOAL_COLS,
    RESULT_DIR,
    SDG_CLEAN_PARQUET,
    SENTENCES_PARQUET,
    SENTIMENT_BATCH,
)
from load_models import load_sbert, load_sentiment
from src.aggregate import company_year_means, minmax_scale_goals, sentiment_ratio
from src.encode import cosine_from_normalized, encode_sentences, mean_similarity_by_goal, save_embeddings
from src.io_util import load_table, save_table
from src.progress import mark_step
from src.sentiment import analyze_sentiment


def _paths(lang: str) -> dict[str, Path]:
    EMBED_DIR.mkdir(parents=True, exist_ok=True)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    return {
        "sdg_emb": EMBED_DIR / f"sdg_{lang}.npy",
        "rep_emb": EMBED_DIR / f"report_{lang}.npy",
        "sdg_idx": EMBED_DIR / f"sdg_index_{lang}.pkl",
        "rep_idx": EMBED_DIR / f"report_index_{lang}.pkl",
        "sim": EMBED_DIR / f"sim_{lang}.npy",
        "meta": EMBED_DIR / f"meta_{lang}.json",
    }


def encode_lang(df_rep: pd.DataFrame, df_sdg: pd.DataFrame, lang: str, reuse: bool = True) -> None:
    p = _paths(lang)
    sdg = df_sdg[df_sdg["lang"] == lang].reset_index(drop=True)
    rep = df_rep[df_rep["doc_lang"] == lang].reset_index(drop=True)
    if sdg.empty:
        print(f"[{lang}] không có câu SDG — bỏ qua")
        return
    if rep.empty:
        print(f"[{lang}] không có câu báo cáo — bỏ qua (PDF tiếng này chưa có)")
        return

    if reuse and p["sdg_emb"].exists() and p["rep_emb"].exists() and p["sdg_idx"].exists() and p["rep_idx"].exists():
        old_rep = pd.read_pickle(p["rep_idx"])
        if len(old_rep) == len(rep) and list(old_rep["sentence"]) == list(rep["sentence"]):
            print(f"[{lang}] embeddings đã khớp, skip encode")
            return

    print(f"[{lang}] load SBERT ...")
    model = load_sbert(lang)
    print(f"[{lang}] encode SDG n={len(sdg)}")
    sdg_emb = encode_sentences(model, sdg["sentence"].astype(str).tolist())
    print(f"[{lang}] encode reports n={len(rep)}")
    rep_emb = encode_sentences(model, rep["sentence"].astype(str).tolist())
    del model

    save_embeddings(p["sdg_emb"], sdg_emb)
    save_embeddings(p["rep_emb"], rep_emb)
    sdg[["goalnum", "gpnum", "gpname", "sentence"]].to_pickle(p["sdg_idx"])
    cols = [c for c in ["doc_id", "fname", "company", "year", "doc_lang", "sent_lang", "sent_idx", "sentence"] if c in rep.columns]
    rep[cols].to_pickle(p["rep_idx"])

    sim = cosine_from_normalized(rep_emb, sdg_emb)
    np.save(p["sim"], sim.astype(np.float32))
    meta = {
        "lang": lang,
        "n_sdg": int(len(sdg)),
        "n_report": int(len(rep)),
        "dim": int(rep_emb.shape[1]),
        "sdg_by_goal": sdg["goalnum"].value_counts().sort_index().astype(int).to_dict(),
    }
    p["meta"].write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[{lang}] sim shape={sim.shape}  saved under {EMBED_DIR}")


def similarity_table(df_rep: pd.DataFrame, langs: list[str]) -> pd.DataFrame:
    parts = []
    for lang in langs:
        p = _paths(lang)
        if not p["rep_idx"].exists() or not (p["sim"].exists() or (p["rep_emb"].exists() and p["sdg_emb"].exists())):
            continue
        idx = pd.read_pickle(p["rep_idx"])
        sdg_idx = pd.read_pickle(p["sdg_idx"])
        if p["sim"].exists():
            sim = np.load(p["sim"])
        else:
            sim = cosine_from_normalized(np.load(p["rep_emb"]), np.load(p["sdg_emb"]))
        goals = mean_similarity_by_goal(sim, sdg_idx["goalnum"], GOAL_COLS)
        part = pd.concat([idx.reset_index(drop=True), goals.reset_index(drop=True)], axis=1)
        parts.append(part)
    if not parts:
        return pd.DataFrame()
    out = pd.concat(parts, ignore_index=True)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    out.to_pickle(RESULT_DIR / "result_similarity.pkl")
    save_table(out, RESULT_DIR / "result_similarity.parquet")
    return out


def run_sentiment(df_sim: pd.DataFrame, langs: list[str], reuse: bool = True) -> pd.DataFrame:
    ckpt = RESULT_DIR / "sentiment_checkpoint.pkl"
    done = pd.DataFrame()
    if reuse and ckpt.exists():
        done = pd.read_pickle(ckpt)
        print(f"sentiment checkpoint: {len(done)} câu")

    frames = [done] if len(done) else []
    for lang in langs:
        sub = df_sim[df_sim["doc_lang"] == lang].copy()
        if sub.empty:
            continue
        if len(done):
            already = set(zip(done["fname"], done["sent_idx"])) if "sent_idx" in done.columns else set()
            if "sent_idx" in sub.columns:
                mask = [ (r.fname, r.sent_idx) not in already for r in sub.itertuples() ]
                sub = sub.loc[mask]
        if sub.empty:
            print(f"[{lang}] sentiment đã đủ")
            continue
        print(f"[{lang}] load sentiment model, n={len(sub)}")
        clf = load_sentiment(lang)
        chunk_n = 500
        for start in range(0, len(sub), chunk_n):
            piece = sub.iloc[start : start + chunk_n]
            res = analyze_sentiment(
                clf,
                piece["sentence"].astype(str).tolist(),
                batch_size=SENTIMENT_BATCH,
                segment_vi=(lang == "vi"),
            )
            add = piece[["fname", "sent_idx"]].reset_index(drop=True)
            add["label"] = [r["label"] for r in res]
            add["confidence"] = [r["confidence"] for r in res]
            add["score"] = [r["score"] for r in res]
            add["sent_lang_model"] = lang
            frames.append(add)
            merged = pd.concat(frames, ignore_index=True)
            merged.to_pickle(ckpt)
            frames = [merged]
            print(f"[{lang}] checkpoint {len(merged)}", flush=True)
        del clf

    senti = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if senti.empty:
        return df_sim
    keys = ["fname", "sent_idx"]
    out = df_sim.merge(senti, on=keys, how="left")
    out.to_pickle(RESULT_DIR / "result_both.pkl")
    save_table(out, RESULT_DIR / "result_both.parquet")
    return out


def write_aggregates(df: pd.DataFrame) -> None:
    scaled = minmax_scale_goals(df)
    save_table(scaled, RESULT_DIR / "result_scaled.parquet")
    scaled.to_pickle(RESULT_DIR / "result_scaled.pkl")

    comp = company_year_means(scaled)
    comp.to_pickle(RESULT_DIR / "company_year_means.pkl")
    comp.to_csv(RESULT_DIR / "company_year_means.csv")

    from src.aggregate import category_means

    cat = category_means(comp)
    cat.to_pickle(RESULT_DIR / "category_means.pkl")
    cat.to_csv(RESULT_DIR / "category_means.csv")

    if "label" in df.columns:
        ratio = sentiment_ratio(df)
        ratio.to_pickle(RESULT_DIR / "sentiment_counts.pkl")
        ratio.to_csv(RESULT_DIR / "sentiment_counts.csv")
        print("sentiment counts:\n", ratio)

    print("company-year means (goal sample):\n", comp[GOAL_COLS[:5]].round(2).head(20) if len(comp) else comp)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--encode-only", action="store_true")
    parser.add_argument("--sentiment-only", action="store_true")
    parser.add_argument("--lang", choices=["vi", "en", "both"], default="both")
    parser.add_argument("--no-reuse", action="store_true")
    args = parser.parse_args()

    langs = ["vi", "en"] if args.lang == "both" else [args.lang]
    reuse = not args.no_reuse

    df_rep = load_table(SENTENCES_PARQUET)
    df_sdg = load_table(SDG_CLEAN_PARQUET)
    print(f"reports={len(df_rep)}  sdg={len(df_sdg)}")

    if not args.sentiment_only:
        for lang in langs:
            encode_lang(df_rep, df_sdg, lang, reuse=reuse)
        df_sim = similarity_table(df_rep, langs)
        mark_step(
            "encode_similarity",
            status="done",
            n_rows=int(len(df_sim)),
            langs=langs,
            embed_dir=str(EMBED_DIR),
        )
    else:
        df_sim = pd.read_pickle(RESULT_DIR / "result_similarity.pkl")

    if args.encode_only:
        write_aggregates(df_sim)
        print("==== encode-only done ====")
        return

    df_both = run_sentiment(df_sim, langs, reuse=reuse)
    write_aggregates(df_both)
    mark_step(
        "sentiment",
        status="done",
        n_rows=int(len(df_both)),
        labels=df_both["label"].value_counts().astype(int).to_dict() if "label" in df_both.columns else {},
    )
    print("==== pipeline done ====")


if __name__ == "__main__":
    main()
