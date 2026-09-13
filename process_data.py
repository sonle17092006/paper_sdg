"""Xử lý PDF + corpus SDG, lưu dữ liệu đã làm sạch để lần sau không extract lại.

    python process_data.py                 # incremental: PDF mới / đổi mtime mới chạy
    python process_data.py --rebuild       # làm lại toàn bộ
    python process_data.py --sdg-only      # chỉ xuất sdg_en / sdg_vi đã clean

Output:
    data/processed/sentences.parquet|.csv
    data/processed/sdg_clean.parquet
    data/processed/manifest.json
    data/sdg_vi.xlsx
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from config import (
    GOAL_TO_GP,
    GP_NAME,
    MANIFEST_JSON,
    PDF_DIR,
    PROCESSED_DIR,
    SDG_CLEAN_PARQUET,
    SDG_EN_XLSX,
    SDG_VI_XLSX,
    SENTENCES_PARQUET,
    EXCLUDE_FILES,
)
from src.clean import clean_sentence, filter_sentences
from src.extract import iter_pdfs, parse_fname, sentences_from_pdf
from src.io_util import load_table, save_table
from src.language import detect_doc_lang, detect_sentence_lang
from src.progress import mark_step
from src.sdg_vi_corpus import build_sdg_vi_frame
from src.tokenize_sent import count_words


def _file_id(path: Path) -> dict:
    st = path.stat()
    raw = f"{path.resolve()}|{st.st_mtime_ns}|{st.st_size}"
    return {
        "path": str(path.resolve()),
        "rel": str(path.relative_to(ROOT)) if path.is_relative_to(ROOT) else str(path),
        "mtime_ns": st.st_mtime_ns,
        "size": st.st_size,
        "sha1_key": hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16],
    }


def load_manifest() -> dict:
    if not MANIFEST_JSON.exists():
        return {"files": {}}
    with MANIFEST_JSON.open(encoding="utf-8") as f:
        return json.load(f)


def save_manifest(man: dict) -> None:
    MANIFEST_JSON.parent.mkdir(parents=True, exist_ok=True)
    man["updated_at"] = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    with MANIFEST_JSON.open("w", encoding="utf-8") as f:
        json.dump(man, f, ensure_ascii=False, indent=2)


def write_sdg_tables() -> pd.DataFrame:
    """Xuất SDG EN (copy paper) + SDG VI, gộp bảng clean có cột lang."""
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    df_en = pd.read_excel(SDG_EN_XLSX)
    df_en = df_en.rename(columns={c: c.strip() for c in df_en.columns})
    df_en["lang"] = "en"
    df_en["sentence"] = df_en["sentence"].map(lambda s: clean_sentence(str(s), "en"))
    df_en = df_en[df_en["sentence"].str.len() > 0].copy()
    df_en["source"] = "paper_sdg.xlsx"

    df_vi = build_sdg_vi_frame()
    SDG_VI_XLSX.parent.mkdir(parents=True, exist_ok=True)
    df_vi.to_excel(SDG_VI_XLSX, index=False)
    df_vi = df_vi.copy()
    df_vi["lang"] = "vi"
    df_vi["sentence"] = df_vi["sentence"].map(lambda s: clean_sentence(str(s), "vi"))
    df_vi["source"] = "sdg_vi_corpus"

    df = pd.concat([df_en, df_vi], ignore_index=True)
    df["gpnum"] = df["goalnum"].map(GOAL_TO_GP).fillna(df.get("gpnum"))
    df["gpname"] = df["gpnum"].map(GP_NAME).fillna(df.get("gpname"))
    save_table(df, SDG_CLEAN_PARQUET)
    print(f"SDG clean: {len(df)} rows  (en={int((df.lang=='en').sum())}, vi={int((df.lang=='vi').sum())})")
    print(f"  saved {SDG_CLEAN_PARQUET.with_suffix('.pkl')} / .csv")
    print(f"  saved {SDG_VI_XLSX}")
    return df


def extract_one(path: Path, doc_id: int, ocr: bool = False) -> pd.DataFrame:
    company, year = parse_fname(path.name)
    sents = filter_sentences(sentences_from_pdf(path, ocr=ocr))
    if not sents:
        return pd.DataFrame()
    sent_lang = [detect_sentence_lang(s) for s in sents]
    doc_lang = detect_doc_lang(sents)
    rows = []
    for i, (s, sl) in enumerate(zip(sents, sent_lang)):
        rows.append(
            {
                "doc_id": doc_id,
                "fname": path.name,
                "relpath": str(path.relative_to(PDF_DIR)) if path.is_relative_to(PDF_DIR) else path.name,
                "company": company,
                "year": year,
                "doc_lang": doc_lang,
                "sent_lang": sl,
                "sent_idx": i,
                "n_words": count_words(s),
                "sentence_raw": s,
                "sentence": clean_sentence(s, sl),
            }
        )
    return pd.DataFrame(rows)


def _rel_key(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(PDF_DIR.resolve())).replace("\\", "/")
    except ValueError:
        return path.name


def write_quality_report(df: pd.DataFrame) -> Path:
    """Ghi chú file mỏng / bất thường sau extract."""
    path = PROCESSED_DIR / "extraction_notes.md"
    lines = [
        "# Ghi chú chất lượng extract",
        "",
        f"Tổng: **{len(df)} câu / {df['fname'].nunique() if len(df) else 0} file**.",
        "",
        "## Số câu theo file",
        "",
        "| File | Câu | Năm | Ghi chú |",
        "|---|---:|---|---|",
    ]
    if len(df):
        g = df.groupby(["fname", "company", "year"]).size().reset_index(name="n")
        for r in g.itertuples():
            note = ""
            if r.n < 80:
                note = "mỏng — kiểm tra scan / landing PDF / OCR"
            lines.append(f"| {r.fname} | {r.n} | {r.year} | {note} |")
    lines += [
        "",
        "## Cách xử lý",
        "",
        "- PDF scan (không text layer): `python process_data.py --ocr --rebuild`",
        "  cần Tesseract + `models/tessdata/vie.traineddata` (tự tải).",
        "- PDF landing/web (PLX 2021, 2024): không phải full report — nên thay file nguồn.",
        "- Câu CID rác / HOME-MENU / mục lục dán 1 dòng: đã lọc trong `src/clean.py`.",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")
    print(f"  saved {path}")
    return path


def process_pdfs(rebuild: bool = False, ocr: bool = False) -> pd.DataFrame:
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    pdfs = iter_pdfs(PDF_DIR)
    if not pdfs:
        print(f"Không thấy PDF trong {PDF_DIR}")
        return pd.DataFrame()

    man = {"files": {}} if rebuild else load_manifest()
    old = pd.DataFrame()
    if not rebuild:
        try:
            old = load_table(SENTENCES_PARQUET)
        except FileNotFoundError:
            old = pd.DataFrame()

    parts: list[pd.DataFrame] = []
    for i, path in enumerate(pdfs):
        if path.name in EXCLUDE_FILES:
            print(f"[skip] {path.name} (bỏ qua theo EXCLUDE_FILES)")
            continue
        meta = _file_id(path)
        key = _rel_key(path)
        prev = man.get("files", {}).get(key)
        thin = bool(prev) and int(prev.get("n_sentences") or 0) < 80
        can_reuse = (
            not rebuild
            and not (ocr and thin)
            and bool(prev)
            and prev.get("sha1_key") == meta["sha1_key"]
            and len(old)
            and bool((old["fname"] == path.name).any())
        )
        if can_reuse:
            chunk = old[old["fname"] == path.name].copy()
            parts.append(chunk)
            man.setdefault("files", {})[key] = {**prev, **meta, "n_sentences": int(len(chunk))}
            print(f"[skip] {path.name}  ({len(chunk)} câu)")
            continue

        print(f"[extract] {path.name} ...", flush=True)
        part = extract_one(path, doc_id=i, ocr=ocr)
        n = len(part)
        lang = part["doc_lang"].iloc[0] if n else "?"
        print(f"         {n} câu, doc_lang={lang}")
        company, year = parse_fname(path.name)
        meta.update({"n_sentences": n, "doc_lang": lang, "company": company, "year": year})
        man.setdefault("files", {})[key] = meta
        if n:
            parts.append(part)

    df = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
    if len(df):
        df = df[df["sentence"].astype(str).str.len() > 0].copy()
        df["doc_id"] = pd.factorize(df["fname"])[0]
        save_table(df, SENTENCES_PARQUET)
    save_manifest(man)
    n_files = int(df["fname"].nunique()) if len(df) else 0
    print(f"sentences: {len(df)}  files={n_files}")
    print(f"  saved {SENTENCES_PARQUET.with_suffix('.pkl')} / .csv")
    return df


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract + clean + lưu dữ liệu")
    parser.add_argument("--rebuild", action="store_true")
    parser.add_argument("--sdg-only", action="store_true")
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="OCR trang không có text / CID rác (Tesseract vie+eng). Chậm.",
    )
    args = parser.parse_args()

    df_sdg = write_sdg_tables()
    if args.sdg_only:
        mark_step("process_data_sdg", status="done", n_sdg=int(len(df_sdg)))
        return

    df = process_pdfs(rebuild=args.rebuild, ocr=args.ocr)
    if len(df):
        write_quality_report(df)
    payload = {
        "status": "done",
        "n_sentences": int(len(df)),
        "n_files": int(df["fname"].nunique()) if len(df) else 0,
        "by_company": df.groupby("company").size().astype(int).to_dict() if len(df) else {},
        "by_lang": df["doc_lang"].value_counts().astype(int).to_dict() if len(df) else {},
        "n_sdg": int(len(df_sdg)),
    }
    mark_step("process_data", **payload)
    print("==== process_data done ====")


if __name__ == "__main__":
    main()
