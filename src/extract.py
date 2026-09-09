"""Trích câu từ PDF — bám 01_extract_text.ipynb của paper, bổ sung tiếng Việt."""

from __future__ import annotations

import re
from pathlib import Path

import fitz

from config import MIN_WORD_CNT, SKIP_PAGES
from src.clean import is_garbled
from src.tokenize_sent import count_words, normalize_block, split_sentences

FNAME_RE = re.compile(
    r"^(?P<company>[A-Za-z0-9]+)[_\-](?:SR|SDR|ESG|AR)?[_\-]?(?P<year>20\d{2})",
    re.IGNORECASE,
)


def parse_fname(fname: str) -> tuple[str, str]:
    stem = Path(fname).stem
    m = FNAME_RE.match(stem)
    if m:
        return m.group("company").upper(), m.group("year")
    parts = re.split(r"[_\-\.\(\)]", stem)
    company = parts[0].upper() if parts else stem
    year = next((p for p in parts if re.fullmatch(r"20\d{2}", p)), "")
    return company, year


def is_header_or_footer(block, page_h: float) -> bool:
    """Loại bỏ khối văn bản nằm hoàn toàn ở mép trên (top 4.5%) hoặc mép dưới (bottom 4.5%)."""
    y0, y1 = block[1], block[3]
    if y1 <= page_h * 0.045:
        return True
    if y0 >= page_h * 0.955:
        return True
    return False


def get_text_from_blocks(block_lst, min_word_cnt: int = MIN_WORD_CNT) -> str:
    """Giữ lại cho tương thích ngược nếu cần."""
    text_lst = []
    for block in block_lst:
        if block[6] != 0:
            continue
        text = normalize_block(block[4])
        if count_words(text) < min_word_cnt:
            continue
        text_lst.append(text)
    return "\n".join(text_lst)


def _page_text_raw(page) -> str:
    try:
        raw = page.get_text("text")
    except Exception:
        raw = page.getText("text")
    return normalize_block(raw or "")


def _sentences_from_text(text: str, min_word_cnt: int) -> list[str]:
    out = []
    for sentence in split_sentences(text):
        r_sent = " ".join(sentence.split())
        if count_words(r_sent) < min_word_cnt:
            continue
        out.append(r_sent)
    return out


def _need_ocr(raw: str, n_sents: int) -> bool:
    if n_sents == 0 and len(raw.strip()) < 80:
        return True
    if is_garbled(raw) and len(raw) > 40:
        return True
    return False


def sentences_from_pdf(
    path: Path,
    skip_pages: tuple[int, ...] = SKIP_PAGES,
    min_word_cnt: int = MIN_WORD_CNT,
    ocr: bool = False,
) -> list[str]:
    """skip_pages là số trang 1-based, mặc định bỏ trang bìa như paper.

    Trích xuất từng khối văn bản theo bố cục thực tế, bỏ header/footer lặp mép trang.
    Trang 0 câu → fallback full-page text.
    Nếu vẫn rỗng/CID rác và ocr=True → Tesseract vie+eng.
    """
    doc = fitz.open(path)
    sent_lst: list[str] = []
    ocr_pages = 0
    try:
        for page_no, page in enumerate(doc):
            if page_no + 1 in skip_pages:
                continue
            h = page.rect.height
            try:
                blocks = page.get_text("blocks")
            except Exception:
                blocks = page.getText("blocks")

            page_sents: list[str] = []
            for b in blocks:
                if b[6] != 0:  # Không phải khối văn bản
                    continue
                if is_header_or_footer(b, h):
                    continue
                block_sents = _sentences_from_text(b[4], min_word_cnt)
                page_sents.extend(block_sents)

            raw = _page_text_raw(page)
            if not page_sents:
                page_sents = _sentences_from_text(raw, min_word_cnt)

            if ocr and _need_ocr(raw, len(page_sents)):
                from src.ocr import ocr_page

                ocr_text = ocr_page(page)
                ocr_sents = _sentences_from_text(ocr_text, min_word_cnt)
                if len(ocr_sents) > len(page_sents):
                    page_sents = ocr_sents
                    ocr_pages += 1
            sent_lst.extend(page_sents)
    finally:
        doc.close()
    if ocr_pages:
        print(f"         OCR {ocr_pages} trang")
    return sent_lst


def iter_pdfs(root: Path) -> list[Path]:
    return sorted(p for p in Path(root).rglob("*.pdf") if p.is_file())
