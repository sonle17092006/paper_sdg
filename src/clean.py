"""Làm sạch câu SDG / báo cáo.

Paper thay mọi ký tự không phải ASCII keyboard bằng khoảng trắng.
Cách đó phá tiếng Việt, nên chỉ áp dụng cho câu tiếng Anh.
"""

from __future__ import annotations

import re

# Giữ nguyên regex paper (02_sentence_similarity.ipynb)
EN_NON_KEYBOARD = re.compile(r"[^!\"#$%&'()*+,-./:;<=>?@\[\]^_`{|}~\\0-9a-zA-Z]")
VI_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\u200b\u200c\u200d\ufeff]")
MULTI_SPACE = re.compile(r"\s+")


def clean_en(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = EN_NON_KEYBOARD.sub(" ", text)
    return MULTI_SPACE.sub(" ", text).strip()


def clean_vi(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = VI_CTRL.sub(" ", text)
    text = text.replace("\xa0", " ").replace("\u00ad", "")
    return MULTI_SPACE.sub(" ", text).strip()


def clean_sentence(text: str, lang: str) -> str:
    return clean_vi(text) if lang == "vi" else clean_en(text)


NAV_RE = re.compile(
    r"^(HOME|MENU|NEXT|PREV|Trang chủ|click vào)(\s+(HOME|MENU|NEXT|PREV))*$",
    re.I,
)
GARBLE_RE = re.compile(r"[^\w\s.,;:!?%/()\[\]\"'`“”–\-0-9A-Za-zÀ-ỹ]")


def is_garbled(text: str) -> bool:
    """Text layer CID hỏng (PLX 2023) — gần như không đọc được."""
    if not isinstance(text, str) or len(text) < 20:
        return False
    letters = sum(ch.isalpha() for ch in text)
    weird = len(GARBLE_RE.findall(text))
    if weird / len(text) > 0.15:
        return True
    vn = len(
        re.findall(
            r"[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡ"
            r"ùúụủũưừứựửữỳýỵỷỹđÀ-Ỹ]",
            text,
        )
    )
    if letters > 40 and vn < 2 and weird > 8:
        return True
    return False


TOC_RE = re.compile(r"(\.{3,}|(?:\.\s+){3,})\s*\d+")


def is_junk_sentence(text: str) -> bool:
    if not text or len(text.strip()) < 15:
        return True
    s = text.strip()
    if NAV_RE.match(s):
        return True
    if s.lower() in {"home", "menu", "next", "prev"}:
        return True
    if is_garbled(s):
        return True
    # Mục lục có chuỗi chấm lửng liên tiếp dẫn tới số trang
    if TOC_RE.search(s):
        return True
    return False


def filter_sentences(sentences: list[str]) -> list[str]:
    out = []
    seen: set[str] = set()
    for s in sentences:
        s = s.replace("\ufffd", " ").replace("�", " ")
        s = MULTI_SPACE.sub(" ", s).strip()
        if is_junk_sentence(s):
            continue
        key = s.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(s)
    return out
