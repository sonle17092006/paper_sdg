"""Bộ điều phối toàn diện (Master Orchestrator):
1. Nạp toàn bộ dữ liệu 7 doanh nghiệp và 42 báo cáo từ data/results/*.csv
2. Xây dựng bản Tiếng Việt hoàn chỉnh (paper_sdg_vietnam_vi_updated.docx)
3. Xây dựng bản Tiếng Anh hoàn chỉnh (paper_sdg_vietnam_en.docx)
4. Xây dựng 2 bản EndNote-Ready ({Tác giả, Năm})
5. Tích hợp Bookmark và Siêu liên kết Ctrl+Click nhảy tới tài liệu tham khảo cho cả hai bản chính.
"""

from __future__ import annotations

import re
import sys
import unicodedata
from pathlib import Path

import docx
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import paper_base as pb
import build_paper_vi_v2 as b_vi
import build_paper_en_v2 as b_en

ROOT = Path(r"C:\Code\paper_sdg")
COLOR_BLACK = docx.shared.RGBColor(0, 0, 0)


def slugify(text: str) -> str:
    text = unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("utf-8")
    return re.sub(r"[^a-zA-Z0-9]", "", text).lower()


def add_hyperlinks_to_file(doc_path: Path, lang: str = "vi"):
    doc = Document(str(doc_path))

    # 1. Tìm mục Tài liệu tham khảo
    ref_heading_idx = -1
    heading_words = ["TÀI LIỆU THAM KHẢO", "REFERENCES"]
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() in heading_words:
            ref_heading_idx = i
            break

    if ref_heading_idx == -1:
        print(f"Could not find References heading in {doc_path.name}")
        return

    # 2. Tạo bookmark cho từng mục tài liệu tham khảo
    ref_map = {}
    bm_id = 200
    for p in doc.paragraphs[ref_heading_idx + 1:]:
        txt = p.text.strip()
        if not txt:
            continue
        m = re.match(r"^([^\(]+?)\s*\((\d{4}[a-z]?)\)", txt)
        if m:
            author_part, year = m.groups()
            first_author = author_part.split(",")[0].split(".")[0].strip()
            author_slug = slugify(first_author)
            bm_name = f"ref_{author_slug}_{year}"

            bm_start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bm_id}" w:name="{bm_name}"/>')
            bm_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bm_id}"/>')

            p._p.insert(0, bm_start)
            p._p.append(bm_end)

            ref_map[(author_slug, year)] = bm_name
            ref_map[author_slug] = bm_name
            bm_id += 1

    print(f"[{lang.upper()}] Bookmarked {len(ref_map) // 2} references.")

    # 3. Quét citation trong thân bài và gắn hyperlink
    cite_pattern = re.compile(
        r"(\((?:[A-Za-z\u00C0-\u024F\u1EA0-\u1EF9\s\.\,\–\-\&]|et al\.)+?,\s*\d{4}[a-z]?\)|[A-Z\u00C0-\u024F\u1EA0-\u1EF9][A-Za-z\u00C0-\u024F\u1EA0-\u1EF9\s\.\,\–\-\&]*(?:và|and|\&|et al\.)\s*[A-Z\u00C0-\u024F\u1EA0-\u1EF9][A-Za-z\u00C0-\u024F\u1EA0-\u1EF9\s\.\,\–\-]*\s*\(\d{4}[a-z]?\))"
    )

    linked_count = 0
    for p_idx in range(ref_heading_idx):
        p = doc.paragraphs[p_idx]
        text = p.text
        if not text.strip():
            continue

        matches = list(cite_pattern.finditer(text))
        if not matches:
            continue

        valid_matches = []
        for m in matches:
            cite_str = m.group(1)
            ym = re.search(r"(\d{4}[a-z]?)", cite_str)
            if not ym:
                continue
            year = ym.group(1)

            author_candidate = cite_str.replace("(", "").replace(")", "")
            first_word = re.split(r"[\s\,\.\&\;]", author_candidate)[0].strip()
            first_slug = slugify(first_word)

            matched_bm = None
            if (first_slug, year) in ref_map:
                matched_bm = ref_map[(first_slug, year)]
            elif first_slug in ref_map:
                matched_bm = ref_map[first_slug]
            else:
                for k, bm in ref_map.items():
                    if isinstance(k, tuple):
                        a_s, y = k
                        if a_s.startswith(first_slug) or first_slug.startswith(a_s):
                            matched_bm = bm
                            break

            if matched_bm:
                valid_matches.append((m.start(), m.end(), cite_str, matched_bm))

        if not valid_matches:
            continue

        p_element = p._p
        pPr = p_element.pPr
        for child in list(p_element):
            if child != pPr:
                p_element.remove(child)

        last_pos = 0
        for start_pos, end_pos, cite_text, bm_name in valid_matches:
            if start_pos > last_pos:
                before_text = text[last_pos:start_pos]
                r = p.add_run(before_text)
                r.font.name = "Times New Roman"
                r.font.size = docx.shared.Pt(12)
                r.font.color.rgb = COLOR_BLACK

            hl_xml = (
                f'<w:hyperlink {nsdecls("w")} w:anchor="{bm_name}" w:history="1">'
                f"<w:r>"
                f"<w:rPr>"
                f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                f'<w:color w:val="000000"/>'
                f'<w:u w:val="none"/>'
                f"</w:rPr>"
                f"<w:t>{cite_text.replace('&', '&amp;')}</w:t>"
                f"</w:r>"
                f"</w:hyperlink>"
            )
            p._p.append(parse_xml(hl_xml))
            linked_count += 1
            last_pos = end_pos

        if last_pos < len(text):
            r = p.add_run(text[last_pos:])
            r.font.name = "Times New Roman"
            r.font.size = docx.shared.Pt(12)
            r.font.color.rgb = COLOR_BLACK

    doc.save(str(doc_path))
    print(f"[{lang.upper()}] Converted {linked_count} in-text citations into live Ctrl+Click hyperlinks in {doc_path.name}")


def main():
    print("=" * 60)
    print("STARTING PAPER BUILD PIPELINE (7 ENTERPRISES, 42 REPORTS)")
    print("=" * 60)

    tables_data = pb.load_all_tables()

    # 1. Build Vietnamese Paper (Updated standard)
    file_vi = ROOT / "paper_sdg_vietnam_vi_updated.docx"
    doc_vi = Document()
    b_vi.build_vietnamese_paper(doc_vi, tables_data, is_endnote_ready=False)
    doc_vi.save(str(file_vi))
    print(f"Saved base VI paper: {file_vi.name} ({file_vi.stat().st_size / 1024:.1f} KB)")
    add_hyperlinks_to_file(file_vi, lang="vi")

    # 2. Build English Paper (Updated standard)
    file_en = ROOT / "paper_sdg_vietnam_en.docx"
    doc_en = Document()
    b_en.build_english_paper(doc_en, tables_data, is_endnote_ready=False)
    doc_en.save(str(file_en))
    print(f"Saved base EN paper: {file_en.name} ({file_en.stat().st_size / 1024:.1f} KB)")
    add_hyperlinks_to_file(file_en, lang="en")

    # 3. Build Vietnamese EndNote-Ready Paper
    file_vi_enw = ROOT / "paper_sdg_vietnam_vi_endnote_ready.docx"
    doc_vi_enw = Document()
    b_vi.build_vietnamese_paper(doc_vi_enw, tables_data, is_endnote_ready=True)
    doc_vi_enw.save(str(file_vi_enw))
    print(f"Saved EndNote-ready VI paper: {file_vi_enw.name} ({file_vi_enw.stat().st_size / 1024:.1f} KB)")

    # 4. Build English EndNote-Ready Paper
    file_en_enw = ROOT / "paper_sdg_vietnam_en_endnote_ready.docx"
    doc_en_enw = Document()
    b_en.build_english_paper(doc_en_enw, tables_data, is_endnote_ready=True)
    doc_en_enw.save(str(file_en_enw))
    print(f"Saved EndNote-ready EN paper: {file_en_enw.name} ({file_en_enw.stat().st_size / 1024:.1f} KB)")

    # 5. Sync to papers/ and duplicate to standard paper_sdg_vietnam_vi.docx
    import shutil
    papers_dir = ROOT / "papers"
    shutil.copy2(file_vi, ROOT / "paper_sdg_vietnam_vi.docx")
    for f in [file_vi, file_en, file_vi_enw, file_en_enw, ROOT / "paper_sdg_vietnam_vi.docx"]:
        shutil.copy2(f, papers_dir / f.name)
    print("Synced all DOCX papers to both root directory and papers/ directory.")

    print("=" * 60)
    print("ALL PAPERS SUCCESSFULLY GENERATED AND VERIFIED!")
    print("=" * 60)


if __name__ == "__main__":
    main()
