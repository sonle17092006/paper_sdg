"""Tách câu không phụ thuộc NLTK (Python 3.14 chặn optparse của NLTK)."""

from __future__ import annotations

import re

# Viết tắt thường gặp trong báo cáo PTBV / GRI / tiếng Việt
_ABBREV = (
    r"(?:TP|Q|P|PGS|GS|TS|ThS|CN|TNHH|CP|CTCP|UBND|HĐQT|HĐTV|BTC|NXB|"
    r"TT|ĐT|KHCN|NN|NT|XH|MT|PTBV|BCT|BPTBV|KSV|KSNB|TGD|TGĐ|PTGĐ|"
    r"GRI|SDG|ESG|IFRS|ISO|UN|UNDP|OECD|ILO|WHO|USD|VND|EUR|VAT|"
    r"Mr|Mrs|Ms|Dr|Prof|Inc|Ltd|Co|vs|etc|e\.g|i\.e|No|Vol|Fig|"
    r"tr|tỷ|triệu)"
)

_PROTECT = re.compile(rf"\b({_ABBREV})\.", flags=re.IGNORECASE)
_SPLIT = re.compile(r"(?<=[\.!?…;])\s+(?=[A-ZÀ-Ỹ0-9\"“\(\–\—•\-])")
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
BULLET_RE = re.compile(r"^[•\-\+\*▪▫►–—\d+\.][\s\t]*")


def count_words(text: str) -> int:
    """Đếm token có ít nhất một chữ cái — tương đương get_cnt() của paper, hỗ trợ Unicode."""
    n = 0
    for tok in text.split():
        if any(ch.isalpha() for ch in tok):
            n += 1
    return n


def normalize_block(text: str) -> str:
    text = text.replace("fi ", "fi")  # ligature bug PyMuPDF, giữ như paper
    text = text.replace("-\n", "")
    text = text.replace("\u00ad", "")
    text = text.replace("\xa0", " ")
    text = _CTRL.sub(" ", text)
    text = text.replace("\r", "\n")
    return text


def split_sentences(text: str) -> list[str]:
    text = normalize_block(text)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    if not lines:
        return []

    chunks = []
    curr = []
    for line in lines:
        is_bullet = bool(BULLET_RE.match(line))
        if curr and (is_bullet or curr[-1].endswith((".", ":", "!", "?", ";"))):
            chunks.append(" ".join(curr))
            curr = [line]
        else:
            curr.append(line)
    if curr:
        chunks.append(" ".join(curr))

    out = []
    for chunk in chunks:
        protected = _PROTECT.sub(r"\1<ABBR>", chunk)
        parts = _SPLIT.split(protected)
        for p in parts:
            s = " ".join(p.replace("<ABBR>", ".").split()).strip()
            s = BULLET_RE.sub("", s).strip()
            if s:
                out.append(s)
    return out
