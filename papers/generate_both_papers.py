"""Tạo hai bản Paper nghiên cứu khoa học hoàn chỉnh:
1. Bản Tiếng Việt: paper_sdg_vietnam_vi.docx
2. Bản Tiếng Anh: paper_sdg_vietnam_en.docx
Cả hai bản đều tuân thủ nghiêm ngặt chuẩn APA 7th, 100% chữ đen, không có header rác,
không có viền trang/khung lạ, không có lỗi LaTeX (dùng Unicode toán học sạch sẽ),
và có phần Literature Review gồm 5 tiểu mục chuyên sâu.
"""

from __future__ import annotations

import io
import sys
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = Path(r"c:\Code\paper_sdg")
FIGURES_DIR = ROOT / "figures"
FILE_VI = ROOT / "paper_sdg_vietnam_vi.docx"
FILE_EN = ROOT / "paper_sdg_vietnam_en.docx"

COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_MUTED = RGBColor(90, 90, 90)


def setup_clean_styles(doc: Document):
    """Cài đặt lề trang A4 chuẩn (Trái 3cm, Phải 2cm, Trên 2.5cm, Dưới 2.5cm).
    HOÀN TOÀN BỎ HEADER RÁC, CHỈ ĐỂ ĐÁNH SỐ TRANG Ở CHÂN TRANG (FOOTER).
    """
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)

        # Xóa sạch header
        header = section.header
        header.is_linked_to_previous = False
        for p in header.paragraphs:
            p.text = ""

        # Footer đánh số trang căn giữa, không có viền
        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.text = ""
        frun = fp.add_run()
        frun.font.name = "Times New Roman"
        frun.font.size = Pt(10)
        frun.font.color.rgb = COLOR_BLACK

        # Field Page Number
        fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls("w"))
        fp._p.append(fldSimple)

    # Normal Style
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = COLOR_BLACK
    normal_style.paragraph_format.line_spacing = 1.35
    normal_style.paragraph_format.space_after = Pt(4)


def add_h1(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text.upper())
    run.font.name = "Times New Roman"
    run.font.size = Pt(13)
    run.bold = True
    run.font.color.rgb = COLOR_BLACK
    return p


def add_h2(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12.5)
    run.bold = True
    run.font.color.rgb = COLOR_BLACK
    return p


def add_h3(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.bold = True
    run.italic = True
    run.font.color.rgb = COLOR_BLACK
    return p


def add_p(doc: Document, text: str, indent: bool = True, bold: bool = False, italic: bool = False):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.35
    p.paragraph_format.space_after = Pt(4)
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.0)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = COLOR_BLACK
    return p


def add_equation_clean(doc: Document, equation_text: str, eq_num: str = ""):
    """Hiển thị công thức toán học dạng Unicode sạch sẽ, KHÔNG dùng raw LaTeX."""
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(equation_text)
    run.font.name = "Cambria Math"
    run.font.size = Pt(11.5)
    run.italic = True
    if eq_num:
        num_run = p.add_run(f"    ({eq_num})")
        num_run.font.name = "Times New Roman"
        num_run.font.size = Pt(11)
        num_run.italic = False


def add_figure_clean(doc: Document, img_filename: str, fig_label: str, title: str, width_inches: float = 6.0):
    """Chèn hình ảnh chuẩn APA 7th."""
    img_path = FIGURES_DIR / img_filename
    if img_path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run()
        run.add_picture(str(img_path), width=Inches(width_inches))

        cap_p = doc.add_paragraph()
        cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_p.paragraph_format.space_after = Pt(10)
        
        c_num = cap_p.add_run(f"{fig_label}. ")
        c_num.font.name = "Times New Roman"
        c_num.font.size = Pt(10)
        c_num.italic = True
        c_num.bold = True
        
        c_title = cap_p.add_run(title)
        c_title.font.name = "Times New Roman"
        c_title.font.size = Pt(10)
        c_title.italic = False
    else:
        err_p = doc.add_paragraph(f"[Image file not found: {img_filename}]")
        err_p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_table_clean(
    doc: Document,
    table_label: str,
    table_title: str,
    headers: list[str],
    data: list[list[str]],
    note: str = "",
    col_widths: list[float] | None = None,
    alignments: list[WD_ALIGN_PARAGRAPH] | None = None,
    font_size: float = 9.5
):
    """Tạo bảng chuẩn APA 7th: 3 đường kẻ ngang đen, không có đường kẻ dọc, không zebra striping."""
    num_p = doc.add_paragraph()
    num_p.paragraph_format.space_before = Pt(12)
    num_p.paragraph_format.space_after = Pt(1)
    num_p.paragraph_format.keep_with_next = True
    num_run = num_p.add_run(table_label)
    num_run.font.name = "Times New Roman"
    num_run.font.size = Pt(10.5)
    num_run.bold = True
    num_run.font.color.rgb = COLOR_BLACK

    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    title_p.paragraph_format.keep_with_next = True
    title_run = title_p.add_run(table_title)
    title_run.font.name = "Times New Roman"
    title_run.font.size = Pt(10.5)
    title_run.italic = True
    title_run.font.color.rgb = COLOR_BLACK

    num_rows = len(data) + 1
    num_cols = len(headers)
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'<w:insideH w:val="none"/>'
        f'<w:insideV w:val="none"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

    header_row = table.rows[0]
    header_tr = header_row._tr.get_or_add_trPr()
    header_tr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    for c_idx, h_text in enumerate(headers):
        cell = header_row.cells[c_idx]
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(tcBorders)

        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar>')
        cell._tc.get_or_add_tcPr().append(tcMar)

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(h_text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(font_size)
        run.bold = True
        run.font.color.rgb = COLOR_BLACK

    for r_idx, row_values in enumerate(data):
        row = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_values):
            cell = row.cells[c_idx]
            tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar>')
            cell._tc.get_or_add_tcPr().append(tcMar)

            p = cell.paragraphs[0]
            if alignments and c_idx < len(alignments):
                p.alignment = alignments[c_idx]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            run = p.add_run(str(val))
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_size)
            run.font.color.rgb = COLOR_BLACK

    if col_widths:
        for row in table.rows:
            for c_idx, w in enumerate(col_widths):
                if c_idx < len(row.cells):
                    row.cells[c_idx].width = Cm(w)

    if note:
        note_p = doc.add_paragraph()
        note_p.paragraph_format.space_before = Pt(3)
        note_p.paragraph_format.space_after = Pt(8)
        n_label = note_p.add_run("Ghi chú. " if "Bảng" in table_label else "Note. ")
        n_label.font.name = "Times New Roman"
        n_label.font.size = Pt(9.5)
        n_label.italic = True
        
        n_text = note_p.add_run(note)
        n_text.font.name = "Times New Roman"
        n_text.font.size = Pt(9.5)
        n_text.italic = False
    else:
        sp = doc.add_paragraph()
        sp.paragraph_format.space_after = Pt(4)

    return table


# Load modules sinh paper tiếng Việt và tiếng Anh
import generate_paper_vi
import generate_paper_en


def safe_save(doc: Document, target_path: Path) -> Path:
    try:
        doc.save(str(target_path))
        return target_path
    except PermissionError:
        alt_path = target_path.with_name(target_path.stem + "_updated.docx")
        doc.save(str(alt_path))
        print(f"    [!] Luu y: Tep {target_path.name} dang mo trong Word, he thong da luu ban moi nhat vao: {alt_path.name}")
        return alt_path


def main():
    print("=================================================================")
    print("[*] BAT DAU TIEN TRINH TAO PAPER NGHIEU CUU KHOA HOC CHUAN APA 7th")
    print("=================================================================")
    
    # 1. Tạo bản Tiếng Việt
    print("\n>>> [1/2] Dang khoi tao Ban Tieng Viet...")
    doc_vi = Document()
    setup_clean_styles(doc_vi)
    generate_paper_vi.build_paper_content(doc_vi)
    out_vi = safe_save(doc_vi, FILE_VI)
    size_vi = out_vi.stat().st_size / 1024
    print(f"    -> Da xuat ban thanh cong: {out_vi.name} ({size_vi:.1f} KB)")

    # 2. Tạo bản Tiếng Anh
    print("\n>>> [2/2] Dang khoi tao Ban Tieng Anh...")
    doc_en = Document()
    setup_clean_styles(doc_en)
    generate_paper_en.build_paper_content(doc_en)
    out_en = safe_save(doc_en, FILE_EN)
    size_en = out_en.stat().st_size / 1024
    print(f"    -> Da xuat ban thanh cong: {out_en.name} ({size_en:.1f} KB)")

    print("\n=================================================================")
    print("[+] HOAN TAT 100%! HAI PHIEN BAN DANG SAN SANG DE SU DUNG.")
    print("=================================================================")


if __name__ == "__main__":
    main()
