"""OCR trang PDF khi không có text layer (scan) hoặc text CID rác.

Ưu tiên Tesseract `vie+eng` (có dấu). RapidOCR là dự phòng (thường mất dấu).
"""

from __future__ import annotations

import os
import shutil
import urllib.request
from pathlib import Path

from config import ROOT

TESSERACT_CANDIDATES = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
]
TESSDATA_DIR = ROOT / "models" / "tessdata"
TESSDATA_URL = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/{name}"


def tesseract_cmd() -> Path | None:
    w = shutil.which("tesseract")
    if w:
        return Path(w)
    for p in TESSERACT_CANDIDATES:
        if p.exists():
            return p
    return None


def ensure_tessdata() -> Path:
    TESSDATA_DIR.mkdir(parents=True, exist_ok=True)
    for name in ("vie.traineddata", "eng.traineddata"):
        dest = TESSDATA_DIR / name
        if dest.exists() and dest.stat().st_size > 10_000:
            continue
        url = TESSDATA_URL.format(name=name)
        print(f"[ocr] download {name}")
        urllib.request.urlretrieve(url, dest)
    return TESSDATA_DIR


def _ocr_tesseract_image(path: Path) -> str:
    import pytesseract
    from PIL import Image

    cmd = tesseract_cmd()
    if cmd is None:
        return ""
    os.environ["TESSDATA_PREFIX"] = str(ensure_tessdata())
    pytesseract.pytesseract.tesseract_cmd = str(cmd)
    img = Image.open(path)
    w, h = img.size
    parts = [img]
    if w / max(h, 1) > 1.55:
        parts = [img.crop((0, 0, w // 2, h)), img.crop((w // 2, 0, w, h))]
    chunks = []
    for part in parts:
        chunks.append(
            pytesseract.image_to_string(part, lang="vie+eng", config="--psm 4")
        )
    return "\n".join(chunks)


def _ocr_rapid_image(path: Path) -> str:
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        return ""
    engine = RapidOCR()
    result, _ = engine(str(path))
    if not result:
        return ""
    return "\n".join(item[1] for item in result if item and len(item) > 1)


def ocr_page(page, zoom: float = 2.0, tmp_dir: Path | None = None) -> str:
    """Render page → OCR. Trả '' nếu không có engine."""
    import fitz

    tmp_dir = tmp_dir or (ROOT / "data" / "processed" / "_ocr_tmp")
    tmp_dir.mkdir(parents=True, exist_ok=True)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    img_path = tmp_dir / f"p{page.number}.png"
    pix.save(str(img_path))
    text = ""
    if tesseract_cmd() is not None:
        try:
            text = _ocr_tesseract_image(img_path)
        except Exception as exc:
            print(f"[ocr] tesseract fail page {page.number+1}: {exc}")
    if len(text.strip()) < 40:
        alt = _ocr_rapid_image(img_path)
        if len(alt.strip()) > len(text.strip()):
            text = alt
    try:
        img_path.unlink()
    except OSError:
        pass
    return text
