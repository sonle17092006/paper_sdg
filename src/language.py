"""Phát hiện ngôn ngữ câu/báo cáo: tiếng Việt vs tiếng Anh."""

from __future__ import annotations

import re

VN_CHARS = re.compile(
    r"[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡ"
    r"ùúụủũưừứựửữỳýỵỷỹđ"
    r"ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ"
    r"ÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ]"
)
LATIN_WORD = re.compile(r"\b[a-zA-Z]{3,}\b")


def detect_sentence_lang(text: str) -> str:
    """Trả 'vi' hoặc 'en'. Ngưỡng dựa trên mật độ dấu thanh tiếng Việt."""
    if not text or not text.strip():
        return "en"
    vn = len(VN_CHARS.findall(text))
    en = len(LATIN_WORD.findall(text))
    if vn >= 3:
        return "vi"
    if vn > 0 and vn >= en * 0.08:
        return "vi"
    return "en"


def detect_doc_lang(sentences: list[str]) -> str:
    """Ngôn ngữ báo cáo = đa số câu. Hòa = vi vì corpus hiện tại là báo cáo VN."""
    if not sentences:
        return "vi"
    n_vi = sum(1 for s in sentences if detect_sentence_lang(s) == "vi")
    return "vi" if n_vi >= len(sentences) / 2 else "en"
