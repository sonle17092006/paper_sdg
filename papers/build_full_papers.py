"""Chương trình tạo hai bản Paper nghiên cứu khoa học hoàn chỉnh (1 bản Tiếng Việt, 1 bản Tiếng Anh)
Chuẩn mực định dạng: APA 7th edition, 100% chữ đen, không có header rác, không có viền lạ,
không có bất kỳ lỗi cú pháp LaTeX nào (toàn bộ công thức dùng ký hiệu Unicode toán học sạch sẽ).
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
import latex2mathml.converter
import lxml.etree as ET

# Load Microsoft Office MathML to OMML stylesheet
XSL_PATH = Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL")
if XSL_PATH.exists():
    try:
        XSLT_DOC = ET.parse(str(XSL_PATH))
        TRANSFORM_MML = ET.XSLT(XSLT_DOC)
    except Exception:
        TRANSFORM_MML = None
else:
    TRANSFORM_MML = None

if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

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

        # Xóa sạch header, không để bất kỳ chữ nào hoặc viền nào ở header
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


def latex_to_omml(latex_str: str):
    if TRANSFORM_MML is None:
        return None
    try:
        mathml = latex2mathml.converter.convert(latex_str)
        mathml_doc = ET.fromstring(mathml.encode("utf-8"))
        omml = TRANSFORM_MML(mathml_doc)
        return parse_xml(ET.tostring(omml).decode("utf-8"))
    except Exception as e:
        return None


def add_equation_clean(doc: Document, latex_eq: str, eq_num: str = ""):
    """Hiển thị công thức toán học dưới dạng đối tượng Office Math (OMML) chuẩn Microsoft Word.
    Tự động biên dịch từ mã LaTeX sang MathML và OMML, hỗ trợ phân số thật có gạch ngang,
    dấu tổng lớn, hàm từng đoạn cases có ngoặc nhọn, chỉ số trên/dưới.
    """
    omml_element = latex_to_omml(latex_eq)
    if omml_element is not None:
        # Bảng 1 dòng 2 cột không viền: Cột 1 căn giữa công thức, Cột 2 căn phải số thứ tự (X)
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        tblPr = table._tbl.tblPr
        borders_xml = (
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/><w:right w:val="none"/>'
            f'<w:insideH w:val="none"/><w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(parse_xml(borders_xml))

        c0 = table.cell(0, 0)
        c0.width = Cm(14.5)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(4)
        p0.paragraph_format.space_after = Pt(4)
        p0._p.append(omml_element)

        c1 = table.cell(0, 1)
        c1.width = Cm(1.5)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_before = Pt(4)
        p1.paragraph_format.space_after = Pt(4)
        if eq_num:
            r1 = p1.add_run(f"({eq_num})")
            r1.font.name = "Times New Roman"
            r1.font.size = Pt(11)
            r1.font.color.rgb = COLOR_BLACK
    else:
        # Fallback nếu không có OMML transform
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(latex_eq)
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


# =========================================================================
# DỮ LIỆU CÁC BẢNG DÙNG CHUNG
# =========================================================================

HEADERS_T1_VI = ["Mã CK", "Tên doanh nghiệp", "Ngành hoạt động cốt lõi", "Sàn", "Giai đoạn", "Chuẩn mực báo cáo áp dụng"]
DATA_T1_VI = [
    ["PAN", "CTCP Tập đoàn PAN (The PAN Group)", "Nông nghiệp và thực phẩm đóng gói", "HOSE", "2019–2025", "GRI Standards, Ma trận trọng yếu"],
    ["PLX", "Tập đoàn Xăng dầu Việt Nam (Petrolimex)", "Năng lượng, xăng dầu và hóa dầu", "HOSE", "2018–2025", "GRI Standards, ISO 14064-1"],
    ["PNJ", "CTCP Vàng bạc Đá quý Phú Nhuận", "Chế tác và bán lẻ trang sức", "HOSE", "2019–2025", "GRI Standards, Trụ cột DE&I"],
    ["VNM", "CTCP Sữa Việt Nam (Vinamilk)", "Chế biến sữa và chăn nuôi bò sữa", "HOSE", "2019–2025", "GRI Standards, PAS 2060, CDP"]
]

HEADERS_T1_EN = ["Ticker", "Company Name", "Core Industry Sector", "Exchange", "Period", "Reporting Standards Applied"]
DATA_T1_EN = [
    ["PAN", "The PAN Group JSC", "Agriculture & Packaged Food", "HOSE", "2019–2025", "GRI Standards, Materiality Matrix"],
    ["PLX", "Vietnam National Petroleum Group (Petrolimex)", "Energy, Petroleum & Petrochemicals", "HOSE", "2018–2025", "GRI Standards, ISO 14064-1"],
    ["PNJ", "Phu Nhuan Jewelry JSC", "Jewelry Manufacturing & Retail", "HOSE", "2019–2025", "GRI Standards, DE&I Pillars"],
    ["VNM", "Vietnam Dairy Products JSC (Vinamilk)", "Dairy Processing & High-tech Farming", "HOSE", "2019–2025", "GRI Standards, PAS 2060, CDP"]
]

HEADERS_T2_VI = ["Công ty", "Năm", "Tên tệp báo cáo PDF", "Số câu", "Số trang", "Mật độ (câu/trang)", "Ghi chú tài liệu số hóa"]
DATA_T2_VI = [
    ["PAN", "2019", "PAN_SR_2019.pdf", "618", "112", "5.52", "Văn bản điện tử chuẩn"],
    ["PAN", "2020", "PAN_SR_2020.pdf", "752", "83", "9.06", "Văn bản điện tử chuẩn"],
    ["PAN", "2021", "PAN_SR_2021.pdf", "834", "76", "10.97", "Văn bản điện tử chuẩn"],
    ["PAN", "2022", "PAN_SR_2022.pdf", "812", "67", "12.12", "Văn bản điện tử chuẩn"],
    ["PAN", "2023", "PAN_SR_2023.pdf", "411", "61", "6.74", "Văn bản điện tử chuẩn"],
    ["PAN", "2024", "PAN_SR_2024.pdf", "788", "87", "9.06", "Văn bản điện tử chuẩn"],
    ["PAN", "2025", "PAN_SR_2025.pdf", "730", "89", "8.20", "Văn bản điện tử chuẩn"],
    ["PLX", "2018", "PLX_SR_2018.pdf", "413", "59", "7.00", "Văn bản điện tử chuẩn"],
    ["PLX", "2019", "PLX_SR_2019.pdf", "747", "89", "8.39", "Văn bản điện tử chuẩn"],
    ["PLX", "2020", "PLX_SR_2020.pdf", "842", "85", "9.91", "Văn bản điện tử chuẩn"],
    ["PLX", "2021", "PLX_SR_2021.pdf", "48", "31", "1.55", "Bản tóm tắt trực tuyến (mẫu mỏng)"],
    ["PLX", "2022", "PLX_SR_2022.pdf", "1.015", "107", "9.49", "Văn bản điện tử chuẩn"],
    ["PLX", "2023", "PLX_SR_2023.pdf", "379", "76", "4.99", "Văn bản điện tử chuẩn"],
    ["PLX", "2024", "PLX_SR_2024.pdf", "40", "23", "1.74", "Bản tóm tắt trực tuyến (mẫu mỏng)"],
    ["PLX", "2025", "PLX_SR_2025.pdf", "605", "81", "7.47", "Văn bản điện tử chuẩn"],
    ["PNJ", "2019", "PNJ_SR_2019.pdf", "587", "61", "9.62", "Văn bản điện tử chuẩn"],
    ["PNJ", "2020", "PNJ_SR_2020.pdf", "541", "59", "9.17", "Văn bản điện tử chuẩn"],
    ["PNJ", "2021", "PNJ_SR_2021.pdf", "345", "50", "6.90", "Văn bản điện tử chuẩn"],
    ["PNJ", "2022", "PNJ_SR_2022.pdf", "338", "57", "5.93", "Tài liệu ảnh scan (nhiễu OCR)"],
    ["PNJ", "2023", "PNJ_SR_2023.pdf", "510", "56", "9.11", "Văn bản điện tử chuẩn"],
    ["PNJ", "2024", "PNJ_SR_2024.pdf", "567", "69", "8.22", "Văn bản điện tử chuẩn"],
    ["PNJ", "2025", "PNJ_SR_2025.pdf", "712", "82", "8.68", "Văn bản điện tử chuẩn"],
    ["VNM", "2019", "VNM_SR_2019.pdf", "569", "97", "5.87", "Văn bản điện tử chuẩn"],
    ["VNM", "2020", "VNM_SR_2020.pdf", "553", "107", "5.17", "Văn bản điện tử chuẩn"],
    ["VNM", "2021", "VNM_SR_2021.pdf", "585", "109", "5.37", "Văn bản điện tử chuẩn"],
    ["VNM", "2022", "VNM_SR_2022.pdf", "418", "93", "4.49", "Văn bản điện tử chuẩn"],
    ["VNM", "2023", "VNM_SR_2023.pdf", "643", "120", "5.36", "Văn bản điện tử chuẩn"],
    ["VNM", "2024", "VNM_SR_2024.pdf", "691", "112", "6.17", "Văn bản điện tử chuẩn"],
    ["VNM", "2025", "VNM_SR_2025.pdf", "954", "152", "6.28", "Văn bản điện tử chuẩn"]
]

HEADERS_T2_EN = ["Company", "Year", "Report PDF Filename", "Sentences", "Pages", "Density (sent/page)", "Digitization Quality Notes"]
DATA_T2_EN = [
    ["PAN", "2019", "PAN_SR_2019.pdf", "618", "112", "5.52", "Standard digital PDF"],
    ["PAN", "2020", "PAN_SR_2020.pdf", "752", "83", "9.06", "Standard digital PDF"],
    ["PAN", "2021", "PAN_SR_2021.pdf", "834", "76", "10.97", "Standard digital PDF"],
    ["PAN", "2022", "PAN_SR_2022.pdf", "812", "67", "12.12", "Standard digital PDF"],
    ["PAN", "2023", "PAN_SR_2023.pdf", "411", "61", "6.74", "Standard digital PDF"],
    ["PAN", "2024", "PAN_SR_2024.pdf", "788", "87", "9.06", "Standard digital PDF"],
    ["PAN", "2025", "PAN_SR_2025.pdf", "730", "89", "8.20", "Standard digital PDF"],
    ["PLX", "2018", "PLX_SR_2018.pdf", "413", "59", "7.00", "Standard digital PDF"],
    ["PLX", "2019", "PLX_SR_2019.pdf", "747", "89", "8.39", "Standard digital PDF"],
    ["PLX", "2020", "PLX_SR_2020.pdf", "842", "85", "9.91", "Standard digital PDF"],
    ["PLX", "2021", "PLX_SR_2021.pdf", "48", "31", "1.55", "Online summary landing page (thin sample)"],
    ["PLX", "2022", "PLX_SR_2022.pdf", "1,015", "107", "9.49", "Standard digital PDF"],
    ["PLX", "2023", "PLX_SR_2023.pdf", "379", "76", "4.99", "Standard digital PDF"],
    ["PLX", "2024", "PLX_SR_2024.pdf", "40", "23", "1.74", "Online summary landing page (thin sample)"],
    ["PLX", "2025", "PLX_SR_2025.pdf", "605", "81", "7.47", "Standard digital PDF"],
    ["PNJ", "2019", "PNJ_SR_2019.pdf", "587", "61", "9.62", "Standard digital PDF"],
    ["PNJ", "2020", "PNJ_SR_2020.pdf", "541", "59", "9.17", "Standard digital PDF"],
    ["PNJ", "2021", "PNJ_SR_2021.pdf", "345", "50", "6.90", "Standard digital PDF"],
    ["PNJ", "2022", "PNJ_SR_2022.pdf", "338", "57", "5.93", "Scanned image PDF (OCR noise)"],
    ["PNJ", "2023", "PNJ_SR_2023.pdf", "510", "56", "9.11", "Standard digital PDF"],
    ["PNJ", "2024", "PNJ_SR_2024.pdf", "567", "69", "8.22", "Standard digital PDF"],
    ["PNJ", "2025", "PNJ_SR_2025.pdf", "712", "82", "8.68", "Standard digital PDF"],
    ["VNM", "2019", "VNM_SR_2019.pdf", "569", "97", "5.87", "Standard digital PDF"],
    ["VNM", "2020", "VNM_SR_2020.pdf", "553", "107", "5.17", "Standard digital PDF"],
    ["VNM", "2021", "VNM_SR_2021.pdf", "585", "109", "5.37", "Standard digital PDF"],
    ["VNM", "2022", "VNM_SR_2022.pdf", "418", "93", "4.49", "Standard digital PDF"],
    ["VNM", "2023", "VNM_SR_2023.pdf", "643", "120", "5.36", "Standard digital PDF"],
    ["VNM", "2024", "VNM_SR_2024.pdf", "691", "112", "6.17", "Standard digital PDF"],
    ["VNM", "2025", "VNM_SR_2025.pdf", "954", "152", "6.28", "Standard digital PDF"]
]

HEADERS_T3_VI = ["Nhóm danh mục", "Mục tiêu SDG cấu thành", "Nội dung nhu cầu con người", "Số câu chuẩn"]
DATA_T3_VI = [
    ["Đời sống (Life)", "SDG 1, SDG 2, SDG 3", "Xóa nghèo, an ninh lương thực, sức khỏe và phúc lợi", "70 câu"],
    ["Kinh tế (Economic)", "SDG 8, SDG 9", "Việc làm bền vững, tăng trưởng kinh tế, đổi mới và hạ tầng", "43 câu"],
    ["Công bằng (Equity)", "SDG 4, SDG 5, SDG 10", "Giáo dục chất lượng, bình đẳng giới, giảm bất bình đẳng", "69 câu"],
    ["Xã hội (Social)", "SDG 11, SDG 16, SDG 17", "Đô thị bền vững, thể chế minh bạch, đối tác toàn cầu", "72 câu"],
    ["Tài nguyên (Resources)", "SDG 6, SDG 7, SDG 12, SDG 14", "Nước sạch, năng lượng sạch, sản xuất trách nhiệm, đại dương", "89 câu"],
    ["Môi trường (Environments)", "SDG 13, SDG 15", "Hành động khí hậu, bảo tồn hệ sinh thái đất liền và rừng", "48 câu"]
]

HEADERS_T3_EN = ["Category", "Constituent SDGs", "Human Needs Dimension", "Benchmark Sentences"]
DATA_T3_EN = [
    ["Life", "SDG 1, SDG 2, SDG 3", "Poverty alleviation, zero hunger, health & well-being", "70 sentences"],
    ["Economic", "SDG 8, SDG 9", "Decent work, economic growth, industry & infrastructure", "43 sentences"],
    ["Equity", "SDG 4, SDG 5, SDG 10", "Quality education, gender equality, reduced inequalities", "69 sentences"],
    ["Social", "SDG 11, SDG 16, SDG 17", "Sustainable cities, peace & justice, global partnerships", "72 sentences"],
    ["Resources", "SDG 6, SDG 7, SDG 12, SDG 14", "Clean water, clean energy, responsible production, life below water", "89 sentences"],
    ["Environments", "SDG 13, SDG 15", "Climate action, life on land & forest ecosystems", "48 sentences"]
]

HEADERS_T4_VI = ["Mục tiêu SDG", "Số câu", "Trung bình", "Độ lệch chuẩn", "Tối thiểu", "Phân vị 25%", "Trung vị", "Phân vị 75%", "Tối đa"]
DATA_T4_VI = [
    ["SDG 01 (Xóa nghèo)", "17.047", "48.35", "12.48", "6.13", "39.45", "48.85", "57.55", "85.41"],
    ["SDG 02 (Không còn nạn đói)", "17.047", "48.28", "12.46", "7.64", "39.56", "48.76", "57.08", "90.48"],
    ["SDG 03 (Sức khỏe và phúc lợi)", "17.047", "42.42", "10.69", "8.63", "34.99", "42.83", "50.29", "76.30"],
    ["SDG 04 (Giáo dục chất lượng)", "17.047", "46.76", "12.95", "7.34", "37.36", "46.84", "55.66", "92.11"],
    ["SDG 05 (Bình đẳng giới)", "17.047", "38.70", "11.51", "0.00", "30.60", "38.87", "46.82", "86.98"],
    ["SDG 06 (Nước sạch và vệ sinh)", "17.047", "44.56", "12.32", "0.89", "36.43", "43.94", "51.73", "96.79"],
    ["SDG 07 (Năng lượng sạch)", "17.047", "47.94", "12.40", "7.08", "39.00", "48.36", "56.86", "84.94"],
    ["SDG 08 (Việc làm và tăng trưởng)", "17.047", "48.18", "13.72", "5.26", "38.23", "48.37", "58.00", "91.60"],
    ["SDG 09 (Công nghiệp và hạ tầng)", "17.047", "50.11", "14.11", "4.06", "40.11", "50.25", "60.27", "92.85"],
    ["SDG 10 (Giảm bất bình đẳng)", "17.047", "44.70", "10.76", "9.19", "37.01", "45.21", "52.72", "74.65"],
    ["SDG 11 (Đô thị bền vững)", "17.047", "44.12", "11.43", "6.18", "36.28", "44.62", "52.32", "77.97"],
    ["SDG 12 (Tiêu dùng trách nhiệm)", "17.047", "49.71", "12.11", "6.88", "41.30", "50.24", "58.51", "84.69"],
    ["SDG 13 (Hành động khí hậu)", "17.047", "44.28", "11.88", "1.64", "36.13", "44.36", "52.52", "87.19"],
    ["SDG 14 (Tài nguyên biển)", "17.047", "45.97", "11.16", "7.79", "38.18", "46.33", "53.96", "84.24"],
    ["SDG 15 (Tài nguyên đất liền)", "17.047", "49.01", "12.70", "5.77", "40.21", "49.41", "58.06", "94.37"],
    ["SDG 16 (Hòa bình và công lý)", "17.047", "45.56", "11.73", "5.60", "37.34", "46.21", "54.26", "77.66"],
    ["SDG 17 (Đối tác vì mục tiêu)", "17.047", "53.74", "15.25", "7.82", "42.86", "53.78", "64.73", "100.00"]
]

HEADERS_T4_EN = ["Target Goal", "Count", "Mean", "Std. Dev.", "Min", "25%", "Median (50%)", "75%", "Max"]
DATA_T4_EN = [
    ["SDG 01 (No Poverty)", "17,047", "48.35", "12.48", "6.13", "39.45", "48.85", "57.55", "85.41"],
    ["SDG 02 (Zero Hunger)", "17,047", "48.28", "12.46", "7.64", "39.56", "48.76", "57.08", "90.48"],
    ["SDG 03 (Good Health)", "17,047", "42.42", "10.69", "8.63", "34.99", "42.83", "50.29", "76.30"],
    ["SDG 04 (Quality Education)", "17,047", "46.76", "12.95", "7.34", "37.36", "46.84", "55.66", "92.11"],
    ["SDG 05 (Gender Equality)", "17,047", "38.70", "11.51", "0.00", "30.60", "38.87", "46.82", "86.98"],
    ["SDG 06 (Clean Water)", "17,047", "44.56", "12.32", "0.89", "36.43", "43.94", "51.73", "96.79"],
    ["SDG 07 (Affordable Energy)", "17,047", "47.94", "12.40", "7.08", "39.00", "48.36", "56.86", "84.94"],
    ["SDG 08 (Decent Work)", "17,047", "48.18", "13.72", "5.26", "38.23", "48.37", "58.00", "91.60"],
    ["SDG 09 (Industry & Infra)", "17,047", "50.11", "14.11", "4.06", "40.11", "50.25", "60.27", "92.85"],
    ["SDG 10 (Reduced Inequalities)", "17,047", "44.70", "10.76", "9.19", "37.01", "45.21", "52.72", "74.65"],
    ["SDG 11 (Sustainable Cities)", "17,047", "44.12", "11.43", "6.18", "36.28", "44.62", "52.32", "77.97"],
    ["SDG 12 (Responsible Consumption)", "17,047", "49.71", "12.11", "6.88", "41.30", "50.24", "58.51", "84.69"],
    ["SDG 13 (Climate Action)", "17,047", "44.28", "11.88", "1.64", "36.13", "44.36", "52.52", "87.19"],
    ["SDG 14 (Life Below Water)", "17,047", "45.97", "11.16", "7.79", "38.18", "46.33", "53.96", "84.24"],
    ["SDG 15 (Life on Land)", "17,047", "49.01", "12.70", "5.77", "40.21", "49.41", "58.06", "94.37"],
    ["SDG 16 (Peace & Justice)", "17,047", "45.56", "11.73", "5.60", "37.34", "46.21", "54.26", "77.66"],
    ["SDG 17 (Partnerships)", "17,047", "53.74", "15.25", "7.82", "42.86", "53.78", "64.73", "100.00"]
]

HEADERS_T5_VI = ["Công ty", "Năm", "Đời sống", "Kinh tế", "Công bằng", "Xã hội", "Tài nguyên", "Môi trường", "Trung bình"]
DATA_T5_VI = [
    ["PAN", "2019", "46.95", "48.18", "42.61", "46.24", "45.98", "45.24", "45.87"],
    ["PAN", "2020", "44.49", "45.98", "40.21", "44.14", "43.72", "42.96", "43.58"],
    ["PAN", "2021", "47.35", "49.02", "42.94", "47.04", "45.66", "45.51", "46.25"],
    ["PAN", "2022", "47.63", "49.83", "43.80", "47.80", "46.82", "46.68", "47.09"],
    ["PAN", "2023", "44.93", "46.83", "42.16", "45.10", "43.48", "42.96", "44.24"],
    ["PAN", "2024", "46.59", "48.12", "42.34", "46.47", "46.01", "46.10", "45.94"],
    ["PAN", "2025", "49.97", "52.99", "45.80", "50.50", "49.70", "49.68", "49.77"],
    ["PLX", "2018", "45.56", "47.74", "42.17", "46.50", "45.60", "45.27", "45.47"],
    ["PLX", "2019", "45.05", "47.42", "41.95", "45.99", "45.78", "44.96", "45.19"],
    ["PLX", "2020", "45.78", "48.39", "42.51", "47.08", "46.54", "46.08", "46.06"],
    ["PLX", "2021", "47.05", "50.46", "41.74", "48.87", "47.55", "47.90", "47.26*"],
    ["PLX", "2022", "44.66", "47.80", "42.05", "46.55", "46.28", "45.89", "45.54"],
    ["PLX", "2023", "43.92", "46.85", "41.26", "46.45", "45.29", "45.64", "44.90"],
    ["PLX", "2024", "44.09", "50.20", "39.92", "48.82", "48.22", "49.91", "46.86*"],
    ["PLX", "2025", "45.51", "48.48", "42.29", "47.71", "46.88", "47.15", "46.34"],
    ["PNJ", "2019", "46.87", "52.28", "45.90", "50.64", "49.84", "48.38", "48.99"],
    ["PNJ", "2020", "48.50", "53.90", "46.92", "51.75", "50.77", "49.50", "50.22"],
    ["PNJ", "2021", "46.94", "50.96", "44.29", "49.46", "50.11", "48.40", "48.36"],
    ["PNJ", "2022", "49.91", "54.84", "47.38", "52.99", "50.21", "51.17", "51.08"],
    ["PNJ", "2023", "44.54", "48.58", "43.23", "47.33", "47.23", "45.36", "46.05"],
    ["PNJ", "2024", "47.32", "52.07", "45.98", "50.69", "50.51", "48.94", "49.25"],
    ["PNJ", "2025", "48.61", "53.26", "47.40", "51.75", "50.61", "49.54", "50.20"],
    ["VNM", "2019", "42.95", "44.61", "40.80", "44.24", "43.49", "42.60", "43.12"],
    ["VNM", "2020", "44.33", "47.42", "41.71", "47.51", "44.90", "44.75", "45.10"],
    ["VNM", "2021", "45.53", "47.88", "42.39", "47.68", "45.13", "44.75", "45.56"],
    ["VNM", "2022", "43.15", "44.97", "40.58", "44.59", "44.04", "44.45", "43.63"],
    ["VNM", "2023", "45.99", "48.31", "43.59", "47.46", "47.16", "47.18", "46.62"],
    ["VNM", "2024", "45.35", "47.43", "42.83", "46.57", "46.31", "46.30", "45.80"],
    ["VNM", "2025", "50.06", "52.28", "46.48", "51.16", "51.13", "51.95", "50.51"]
]

HEADERS_T5_EN = ["Company", "Year", "Life", "Economic", "Equity", "Social", "Resources", "Environments", "Mean Score"]
DATA_T5_EN = DATA_T5_VI

HEADERS_T6_VI = ["Công ty", "Năm", "Tiêu cực (NEG)", "Trung tính (NEU)", "Tích cực (POS)", "Tổng số câu", "Tỷ số Pos/Neg Ratio"]
DATA_T6_VI = [
    ["PAN", "2019", "55", "160", "403", "618", "7.33"],
    ["PAN", "2020", "66", "172", "514", "752", "7.79"],
    ["PAN", "2021", "88", "139", "607", "834", "6.90"],
    ["PAN", "2022", "89", "149", "574", "812", "6.45"],
    ["PAN", "2023", "55", "59", "297", "411", "5.40"],
    ["PAN", "2024", "60", "115", "613", "788", "10.22"],
    ["PAN", "2025", "69", "127", "534", "730", "7.74"],
    ["PLX", "2018", "55", "96", "262", "413", "4.76"],
    ["PLX", "2019", "118", "216", "413", "747", "3.50"],
    ["PLX", "2020", "121", "274", "447", "842", "3.69"],
    ["PLX", "2021", "10", "8", "30", "48", "3.00*"],
    ["PLX", "2022", "182", "246", "587", "1.015", "3.23"],
    ["PLX", "2023", "55", "78", "246", "379", "4.47"],
    ["PLX", "2024", "3", "9", "28", "40", "9.33*"],
    ["PLX", "2025", "77", "126", "402", "605", "5.22"],
    ["PNJ", "2019", "54", "128", "405", "587", "7.50"],
    ["PNJ", "2020", "43", "97", "401", "541", "9.33"],
    ["PNJ", "2021", "34", "46", "265", "345", "7.79"],
    ["PNJ", "2022", "137", "19", "182", "338", "1.33**"],
    ["PNJ", "2023", "72", "57", "381", "510", "5.29"],
    ["PNJ", "2024", "94", "60", "413", "567", "4.39"],
    ["PNJ", "2025", "106", "107", "499", "712", "4.71"],
    ["VNM", "2019", "61", "102", "406", "569", "6.66"],
    ["VNM", "2020", "38", "60", "455", "553", "11.97"],
    ["VNM", "2021", "64", "63", "458", "585", "7.16"],
    ["VNM", "2022", "38", "57", "323", "418", "8.50"],
    ["VNM", "2023", "57", "107", "479", "643", "8.40"],
    ["VNM", "2024", "51", "86", "554", "691", "10.86"],
    ["VNM", "2025", "78", "163", "713", "954", "9.14"]
]

HEADERS_T6_EN = ["Company", "Year", "Negative (NEG)", "Neutral (NEU)", "Positive (POS)", "Total Sentences", "Pos/Neg Ratio"]
DATA_T6_EN = DATA_T6_VI

HEADERS_T7_VI = ["Công ty", "Năm", "Số câu", "Δ Giá CP", "Tỷ số Ratio", "Δ Ratio", "Phản ứng thị trường", "Điểm GW", "Mức độ rủi ro"]
DATA_T7_VI = [
    ["PAN", "2020", "752", "+14.35%", "7.79", "+6.29%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PAN", "2021", "834", "+16.10%", "6.90", "-11.43%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PAN", "2022", "812", "-40.00%", "6.45", "-6.50%", "Trung thực", "0.00", "Trung thực"],
    ["PAN", "2023", "411", "+21.55%", "5.40", "-16.27%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PAN", "2024", "788", "+21.20%", "10.22", "+89.20%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PAN", "2025", "730", "+20.61%", "7.74", "-24.25%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PLX", "2020", "842", "+3.24%", "3.69", "+5.55%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PLX", "2021", "48", "+16.88%", "3.00", "-18.79%", "Không phân loại", "0.00", "Không rủi ro (mẫu mỏng)"],
    ["PLX", "2022", "1.015", "-31.86%", "3.23", "+7.51%", "Phân kỳ dương", "2.39", "Trung bình"],
    ["PLX", "2023", "379", "-4.32%", "4.47", "+38.68%", "Phân kỳ dương", "1.67", "Nhẹ"],
    ["PLX", "2024", "40", "+18.63%", "9.33", "+108.67%", "Không phân loại", "0.00", "Không rủi ro (mẫu mỏng)"],
    ["PLX", "2025", "605", "-7.64%", "5.22", "-44.06%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PNJ", "2020", "541", "+1.50%", "9.33", "+24.34%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PNJ", "2021", "345", "+27.15%", "7.79", "-16.42%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PNJ", "2022", "338", "+18.61%", "1.33", "-82.96%", "Giá tăng", "0.00", "Không rủi ro (nhiễu OCR)"],
    ["PNJ", "2023", "510", "+3.38%", "5.29", "+298.33%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PNJ", "2024", "567", "+8.85%", "4.39", "-16.97%", "Không phân loại", "0.00", "Không rủi ro"],
    ["PNJ", "2025", "712", "+34.34%", "4.71", "+7.15%", "Không phân loại", "0.00", "Không rủi ro"],
    ["VNM", "2020", "553", "+18.09%", "11.97", "+79.90%", "Không phân loại", "0.00", "Không rủi ro"],
    ["VNM", "2021", "585", "-15.35%", "7.16", "-40.23%", "Trung thực", "0.00", "Trung thực"],
    ["VNM", "2022", "418", "-2.29%", "8.50", "+18.78%", "Phân kỳ dương", "0.43", "Nhẹ"],
    ["VNM", "2023", "643", "-9.54%", "8.40", "-1.14%", "Không phân loại", "0.00", "Không rủi ro"],
    ["VNM", "2024", "691", "-1.93%", "10.86", "+29.26%", "Phân kỳ dương", "0.56", "Nhẹ"],
    ["VNM", "2025", "954", "+23.04%", "9.14", "-15.85%", "Không phân loại", "0.00", "Không rủi ro"]
]

HEADERS_T7_EN = ["Company", "Year", "Sentences", "Δ Stock Price", "Ratio", "Δ Ratio", "Market Classification", "GW Score", "Risk Severity"]
DATA_T7_EN = [
    ["PAN", "2020", "752", "+14.35%", "7.79", "+6.29%", "Unclassified", "0.00", "No Risk"],
    ["PAN", "2021", "834", "+16.10%", "6.90", "-11.43%", "Unclassified", "0.00", "No Risk"],
    ["PAN", "2022", "812", "-40.00%", "6.45", "-6.50%", "Honest Reporting", "0.00", "Honest"],
    ["PAN", "2023", "411", "+21.55%", "5.40", "-16.27%", "Unclassified", "0.00", "No Risk"],
    ["PAN", "2024", "788", "+21.20%", "10.22", "+89.20%", "Unclassified", "0.00", "No Risk"],
    ["PAN", "2025", "730", "+20.61%", "7.74", "-24.25%", "Unclassified", "0.00", "No Risk"],
    ["PLX", "2020", "842", "+3.24%", "3.69", "+5.55%", "Unclassified", "0.00", "No Risk"],
    ["PLX", "2021", "48", "+16.88%", "3.00", "-18.79%", "Unclassified", "0.00", "No Risk (thin sample)"],
    ["PLX", "2022", "1,015", "-31.86%", "3.23", "+7.51%", "Positive Divergence", "2.39", "Moderate Risk"],
    ["PLX", "2023", "379", "-4.32%", "4.47", "+38.68%", "Positive Divergence", "1.67", "Mild Risk"],
    ["PLX", "2024", "40", "+18.63%", "9.33", "+108.67%", "Unclassified", "0.00", "No Risk (thin sample)"],
    ["PLX", "2025", "605", "-7.64%", "5.22", "-44.06%", "Unclassified", "0.00", "No Risk"],
    ["PNJ", "2020", "541", "+1.50%", "9.33", "+24.34%", "Unclassified", "0.00", "No Risk"],
    ["PNJ", "2021", "345", "+27.15%", "7.79", "-16.42%", "Unclassified", "0.00", "No Risk"],
    ["PNJ", "2022", "338", "+18.61%", "1.33", "-82.96%", "Price Growth", "0.00", "No Risk (OCR noise)"],
    ["PNJ", "2023", "510", "+3.38%", "5.29", "+298.33%", "Unclassified", "0.00", "No Risk"],
    ["PNJ", "2024", "567", "+8.85%", "4.39", "-16.97%", "Unclassified", "0.00", "No Risk"],
    ["PNJ", "2025", "712", "+34.34%", "4.71", "+7.15%", "Unclassified", "0.00", "No Risk"],
    ["VNM", "2020", "553", "+18.09%", "11.97", "+79.90%", "Unclassified", "0.00", "No Risk"],
    ["VNM", "2021", "585", "-15.35%", "7.16", "-40.23%", "Honest Reporting", "0.00", "Honest"],
    ["VNM", "2022", "418", "-2.29%", "8.50", "+18.78%", "Positive Divergence", "0.43", "Mild Risk"],
    ["VNM", "2023", "643", "-9.54%", "8.40", "-1.14%", "Unclassified", "0.00", "No Risk"],
    ["VNM", "2024", "691", "-1.93%", "10.86", "+29.26%", "Positive Divergence", "0.56", "Mild Risk"],
    ["VNM", "2025", "954", "+23.04%", "9.14", "-15.85%", "Unclassified", "0.00", "No Risk"]
]

HEADERS_T8_VI = ["Hạng", "Doanh nghiệp", "Số năm tính toán (n)", "Số năm giá giảm", "Số năm GW > 0", "GW trung bình", "GW cực đại", "Năm đạt GW Max", "Số năm Trung thực"]
DATA_T8_VI = [
    ["1", "Petrolimex (PLX)", "6 năm", "3 năm", "2 năm", "0.68", "2.39", "Năm 2022", "0 năm"],
    ["2", "Vinamilk (VNM)", "6 năm", "4 năm", "2 năm", "0.17", "0.56", "Năm 2024", "1 năm (2021)"],
    ["3", "The PAN Group (PAN)", "6 năm", "1 năm", "0 năm", "0.00", "0.00", "—", "1 năm (2022)"],
    ["4", "PNJ (PNJ)", "6 năm", "0 năm", "0 năm", "0.00", "0.00", "—", "0 năm"]
]

HEADERS_T8_EN = ["Rank", "Company", "Years Analyzed (n)", "Years Price Dropped", "Years GW > 0", "Mean GW Score", "Peak GW Score", "Peak Year", "Honest Years"]
DATA_T8_EN = [
    ["1", "Petrolimex (PLX)", "6 years", "3 years", "2 years", "0.68", "2.39", "Year 2022", "0 years"],
    ["2", "Vinamilk (VNM)", "6 years", "4 years", "2 years", "0.17", "0.56", "Year 2024", "1 year (2021)"],
    ["3", "The PAN Group (PAN)", "6 years", "1 year", "0 years", "0.00", "0.00", "—", "1 year (2022)"],
    ["4", "PNJ (PNJ)", "6 years", "0 years", "0 years", "0.00", "0.00", "—", "0 years"]
]

HEADERS_T9_VI = ["Doanh nghiệp", "Cỡ mẫu (n)", "Hệ số Pearson r", "p-value (Pearson)", "Hệ số Spearman ρ", "p-value (Spearman)", "Kết luận thống kê"]
DATA_T9_VI = [
    ["The PAN Group (PAN)", "6", "+0.173", "0.744", "-0.257", "0.623", "Không tương quan (p > 0.10)"],
    ["Petrolimex (PLX)", "6", "+0.352", "0.494", "+0.314", "0.544", "Không tương quan (p > 0.10)"],
    ["PNJ (PNJ)", "6", "-0.470", "0.347", "-0.429", "0.397", "Không tương quan (p > 0.10)"],
    ["Vinamilk (VNM)", "6", "+0.462", "0.356", "+0.429", "0.397", "Không tương quan (p > 0.10)"],
    ["Toàn bộ mẫu gộp (Pool)", "24", "+0.030", "0.890", "-0.047", "0.828", "Hoàn toàn độc lập (p = 0.89)"]
]

HEADERS_T9_EN = ["Company", "Sample (n)", "Pearson r", "p-value (Pearson)", "Spearman ρ", "p-value (Spearman)", "Statistical Conclusion"]
DATA_T9_EN = [
    ["The PAN Group (PAN)", "6", "+0.173", "0.744", "-0.257", "0.623", "No significant correlation (p > 0.10)"],
    ["Petrolimex (PLX)", "6", "+0.352", "0.494", "+0.314", "0.544", "No significant correlation (p > 0.10)"],
    ["PNJ (PNJ)", "6", "-0.470", "0.347", "-0.429", "0.397", "No significant correlation (p > 0.10)"],
    ["Vinamilk (VNM)", "6", "+0.462", "0.356", "+0.429", "0.397", "No significant correlation (p > 0.10)"],
    ["Pooled Sample (Total)", "24", "+0.030", "0.890", "-0.047", "0.828", "Statistically independent (p = 0.89)"]
]

APA_REFS_VI = [
    "Arvidsson, S., & Dumay, J. (2022). Corporate ESG reporting quantity, quality and performance: Where to now for environmental policy and practice? Business Strategy and the Environment, 31(3), 1091–1110. https://doi.org/10.1002/bse.2937",
    "Bingler, J. A., Kraus, M., Leippold, M., & Webersinke, N. (2022). Cheap talk and cherry-picking: What companies say site-wide about climate change. Finance Research Letters, 47, Article 102760. https://doi.org/10.1016/j.frl.2022.102760",
    "Blei, D. M., Ng, A. Y., & Jordan, M. I. (2003). Latent Dirichlet allocation. Journal of Machine Learning Research, 3, 993–1022.",
    "Bộ Kế hoạch và Đầu tư. (2023). Báo cáo rà soát quốc gia tự nguyện lần thứ 2 việc thực hiện các mục tiêu phát triển bền vững của Việt Nam (VNR 2023). Nhà xuất bản Thống kê.",
    "Bộ Tài chính. (2020). Thông tư số 96/2020/TT-BTC ngày 16/11/2020 hướng dẫn công bố thông tin trên thị trường chứng khoán.",
    "Cer, D., Yang, Y., Kong, S. Y., Hua, N., Limtiaco, N., St. John, R., Constant, N., Guajardo-Céspedes, M., Yuan, S., Tar, C., Strope, B., & Kurzweil, R. (2018). Universal sentence encoder for English. In Proceedings of EMNLP 2018: System Demonstrations (pp. 169–174). Association for Computational Linguistics. https://doi.org/10.18653/v1/D18-2029",
    "Chính phủ Việt Nam. (2020). Nghị định số 155/2020/NĐ-CP ngày 31/12/2020 quy định chi tiết thi hành một số điều của Luật Chứng khoán.",
    "Connelly, B. L., Certo, S. T., Ireland, R. D., & Reutzel, C. R. (2011). Signaling theory: A review and assessment. Journal of Management, 37(1), 39–67. https://doi.org/10.1177/0149206310388419",
    "Deegan, C. (2002). Introduction: The legitimising effect of social and environmental disclosures—A theoretical foundation. Accounting, Auditing & Accountability Journal, 15(3), 282–311. https://doi.org/10.1108/09513570210435852",
    "Devlin, J., Chang, M. W., Lee, K., & Toutanova, K. (2019). BERT: Pre-training of deep bidirectional transformers for language understanding. In Proceedings of NAACL-HLT 2019 (pp. 4171–4186). Association for Computational Linguistics.",
    "Du, S., & Yu, K. (2017). The business case for sustainability reporting: Evidence from stock market reactions. Journal of Public Policy & Marketing, 36(2), 313–330. https://doi.org/10.1509/jppm.16.112",
    "El-Haj, M., Rayson, P., Walker, M., Young, S., & Simaki, V. (2020). In search of 'sunlight'? Measuring transparency, disclosure quality and tone in corporate annual reports. Accounting and Business Research, 50(5), 450–476. https://doi.org/10.1080/00014788.2020.1771960",
    "Freeman, R. E. (1984). Strategic management: A stakeholder approach. Pitman.",
    "Friede, G., Busch, T., & Bassen, A. (2015). ESG and financial performance: Aggregated evidence from more than 2000 empirical studies. Journal of Sustainable Finance & Investment, 5(4), 210–233. https://doi.org/10.1080/20430795.2015.1118917",
    "Gerged, A. M., Beddewela, E., & Cowton, C. J. (2021). Is corporate environmental disclosure associated with firm value? A multicountry study of emerging Asian markets. Business Strategy and the Environment, 30(4), 1853–1867. https://doi.org/10.1002/bse.2720",
    "Global Reporting Initiative. (2021). GRI universal standards 2021. Global Sustainability Standards Board.",
    "Heras-Saizarbitoria, I., Urbieta, L., & Boiral, O. (2022). Organizations' engagement with SDGs: From cherry-picking to SDG-washing? Corporate Social Responsibility and Environmental Management, 29(2), 316–328. https://doi.org/10.1002/csr.2202",
    "Hoang, T. C., Abeysekera, I., & Ma, S. (2019). Sustainable reporting in Southeast Asia: A comparative study. Journal of Cleaner Production, 211, 1475–1491. https://doi.org/10.1016/j.jclepro.2018.11.246",
    "Hội đồng Doanh nghiệp vì sự Phát triển Bền vững Việt Nam. (2024). Báo cáo chỉ số doanh nghiệp bền vững (CSI 2024). Liên đoàn Thương mại và Công nghiệp Việt Nam.",
    "Huang, A. H., Wang, H., & Yang, Y. (2023). FinBERT: A large language model for extracting information from financial text. Contemporary Accounting Research, 40(2), 806–841. https://doi.org/10.1111/1911-3846.12832",
    "Hummel, K., & Schlick, C. (2016). The relationship between sustainability performance and sustainability disclosure—Reconciling voluntary disclosure theory and legitimacy theory. Journal of Accounting and Public Policy, 35(5), 455–476. https://doi.org/10.1016/j.jaccpubpol.2016.06.001",
    "Kang, H., & Kim, J. (2022). Analyzing and visualizing text information in corporate sustainability reports using natural language processing methods. Applied Sciences, 12(11), Article 5614. https://doi.org/10.3390/app12115614",
    "Loughran, T., & McDonald, B. (2011). When is a liability not a liability? Textual analysis, dictionaries, and 10-Ks. The Journal of Finance, 66(1), 35–65. https://doi.org/10.1111/j.1540-6261.2010.01625.x",
    "Luccioni, S. A., Baylor, E., & Duchene, N. (2020). Analyzing sustainability reports using natural language processing. In NeurIPS 2020 Workshop on Tackling Climate Change with Machine Learning. arXiv:2011.08073.",
    "Lyon, T. P., & Montgomery, A. W. (2015). The means and end of greenwash. Organization & Environment, 28(2), 223–249. https://doi.org/10.1177/1086026615575332",
    "Matsui, T., Suzuki, M., & Seki, K. (2022). Automated classification of corporate statements for the United Nations sustainable development goals. Sustainability, 14(16), Article 10149. https://doi.org/10.3390/su141610149",
    "Mercereau, B., & Melin, L. (2020). ESG analysis: An NLP approach to assessing corporate sustainability disclosures. The Journal of Investing, 29(7), 50–63. https://doi.org/10.3905/joi.2020.1.157",
    "Merkl-Davies, N. O., & Brennan, N. M. (2007). Discretionary disclosure strategies in corporate narratives: Incremental information or impression management? Journal of Accounting Literature, 26, 116–196.",
    "Mikolov, T., Chen, K., Corrado, G., & Dean, J. (2013). Efficient estimation of word representations in vector space. In Proceedings of ICLR 2013. arXiv:1301.3781.",
    "Muñoz-Torres, M. J., Fernández-Izquierdo, M. Á., Rivera-Lirio, J. M., & Escrig-Olmedo, E. (2019). Can modern sustainability reports track the SDGs? An assessment framework. Sustainability, 11(5), Article 1421. https://doi.org/10.3390/su11051421",
    "Nguyen, D. Q., & Nguyen, A. T. (2020). PhoBERT: Pre-trained language models for Vietnamese. In Findings of EMNLP 2020 (pp. 1037–1042). Association for Computational Linguistics. https://doi.org/10.18653/v1/2020.findings-emnlp.92",
    "Pennington, J., Socher, R., & Manning, C. D. (2014). GloVe: Global vectors for word representation. In Proceedings of EMNLP 2014 (pp. 1532–1543). Association for Computational Linguistics. https://doi.org/10.3115/v1/D14-1162",
    "Pizzi, S., Caputo, A., Corvino, A., & Ficco, A. (2020). Management research and the UN sustainable development goals (SDGs): A bibliometric investigation and systematic review. Journal of Cleaner Production, 276, Article 124033. https://doi.org/10.1016/j.jclepro.2020.124033",
    "PwC Vietnam. (2022). Báo cáo khảo sát mức độ sẵn sàng thực hành ESG tại Việt Nam năm 2022: Từ tham vọng đến hành động. PwC Việt Nam.",
    "Reimers, N., & Gurevych, I. (2019). Sentence-BERT: Sentence embeddings using Siamese BERT-networks. In Proceedings of EMNLP-IJCNLP 2019 (pp. 3982–3992). Association for Computational Linguistics. https://doi.org/10.18653/v1/D19-1410",
    "Sanh, V., Debut, L., Chaumond, J., & Wolf, T. (2019). DistilBERT, a distilled version of BERT: smaller, faster, cheaper and lighter. arXiv preprint arXiv:1910.01108.",
    "Seele, P., & Gatti, L. (2017). Greenwashing revisited: In search of a typology and accusation-based definition. Business Strategy and the Environment, 26(2), 239–252. https://doi.org/10.1002/bse.1912",
    "Spence, M. (1973). Job market signaling. The Quarterly Journal of Economics, 87(3), 355–374. https://doi.org/10.2307/1882010",
    "Stacchezzini, R., Melloni, G., & Lai, A. (2016). Sustainability management and reporting: The role of impression management for corporate social responsibility disclosure. Journal of Cleaner Production, 136, 102–110. https://doi.org/10.1016/j.jclepro.2016.04.095",
    "Suchman, M. C. (1995). Managing legitimacy: Strategic and institutional approaches. Academy of Management Review, 20(3), 571–610. https://doi.org/10.5465/amr.1995.9508080331",
    "Tran, M., & Beddewela, E. (2020). Evaluating the quality of CSR disclosure in Vietnam: An empirical examination of listed firms. Journal of Business Ethics, 166(3), 569–589. https://doi.org/10.1007/s10551-019-04135-2",
    "United Nations. (2015). Transforming our world: The 2030 agenda for sustainable development (Resolution A/RES/70/1). United Nations General Assembly.",
    "Ủy ban Chứng khoán Nhà nước. (2024). Sổ tay hướng dẫn thực hành và công bố thông tin môi trường, xã hội và quản trị (ESG) cho doanh nghiệp niêm yết. Nhà xuất bản Tài chính.",
    "Veenstra, E. M., & Ellemers, N. (2020). CSR does not equal investment in CSR: A linguistic analysis of corporate social responsibility reports. Journal of Business Ethics, 161(2), 347–363. https://doi.org/10.1007/s10551-018-3904-7",
    "Wang, W., Wei, F., Dong, L., Bao, H., Yang, N., & Zhou, M. (2020). MiniLM: Deep self-attention distillation for task-agnostic compression of pre-trained transformers. In Advances in Neural Information Processing Systems (Vol. 33, pp. 5776–5788). Curran Associates, Inc.",
    "World Commission on Environment and Development. (1987). Our common future. Oxford University Press."
]

APA_REFS_EN = [
    "Arvidsson, S., & Dumay, J. (2022). Corporate ESG reporting quantity, quality and performance: Where to now for environmental policy and practice? Business Strategy and the Environment, 31(3), 1091–1110. https://doi.org/10.1002/bse.2937",
    "Bingler, J. A., Kraus, M., Leippold, M., & Webersinke, N. (2022). Cheap talk and cherry-picking: What companies say site-wide about climate change. Finance Research Letters, 47, Article 102760. https://doi.org/10.1016/j.frl.2022.102760",
    "Blei, D. M., Ng, A. Y., & Jordan, M. I. (2003). Latent Dirichlet allocation. Journal of Machine Learning Research, 3, 993–1022.",
    "Cer, D., Yang, Y., Kong, S. Y., Hua, N., Limtiaco, N., St. John, R., Constant, N., Guajardo-Céspedes, M., Yuan, S., Tar, C., Strope, B., & Kurzweil, R. (2018). Universal sentence encoder for English. In Proceedings of EMNLP 2018: System Demonstrations (pp. 169–174). Association for Computational Linguistics. https://doi.org/10.18653/v1/D18-2029",
    "Connelly, B. L., Certo, S. T., Ireland, R. D., & Reutzel, C. R. (2011). Signaling theory: A review and assessment. Journal of Management, 37(1), 39–67. https://doi.org/10.1177/0149206310388419",
    "Deegan, C. (2002). Introduction: The legitimising effect of social and environmental disclosures—A theoretical foundation. Accounting, Auditing & Accountability Journal, 15(3), 282–311. https://doi.org/10.1108/09513570210435852",
    "Devlin, J., Chang, M. W., Lee, K., & Toutanova, K. (2019). BERT: Pre-training of deep bidirectional transformers for language understanding. In Proceedings of NAACL-HLT 2019 (pp. 4171–4186). Association for Computational Linguistics.",
    "Du, S., & Yu, K. (2017). The business case for sustainability reporting: Evidence from stock market reactions. Journal of Public Policy & Marketing, 36(2), 313–330. https://doi.org/10.1509/jppm.16.112",
    "El-Haj, M., Rayson, P., Walker, M., Young, S., & Simaki, V. (2020). In search of 'sunlight'? Measuring transparency, disclosure quality and tone in corporate annual reports. Accounting and Business Research, 50(5), 450–476. https://doi.org/10.1080/00014788.2020.1771960",
    "Freeman, R. E. (1984). Strategic management: A stakeholder approach. Pitman.",
    "Friede, G., Busch, T., & Bassen, A. (2015). ESG and financial performance: Aggregated evidence from more than 2000 empirical studies. Journal of Sustainable Finance & Investment, 5(4), 210–233. https://doi.org/10.1080/20430795.2015.1118917",
    "Gerged, A. M., Beddewela, E., & Cowton, C. J. (2021). Is corporate environmental disclosure associated with firm value? A multicountry study of emerging Asian markets. Business Strategy and the Environment, 30(4), 1853–1867. https://doi.org/10.1002/bse.2720",
    "Global Reporting Initiative. (2021). GRI universal standards 2021. Global Sustainability Standards Board.",
    "Government of Vietnam. (2020). Decree No. 155/2020/ND-CP detailing the implementation of certain articles of the Law on Securities.",
    "Heras-Saizarbitoria, I., Urbieta, L., & Boiral, O. (2022). Organizations' engagement with SDGs: From cherry-picking to SDG-washing? Corporate Social Responsibility and Environmental Management, 29(2), 316–328. https://doi.org/10.1002/csr.2202",
    "Hoang, T. C., Abeysekera, I., & Ma, S. (2019). Sustainable reporting in Southeast Asia: A comparative study. Journal of Cleaner Production, 211, 1475–1491. https://doi.org/10.1016/j.jclepro.2018.11.246",
    "Huang, A. H., Wang, H., & Yang, Y. (2023). FinBERT: A large language model for extracting information from financial text. Contemporary Accounting Research, 40(2), 806–841. https://doi.org/10.1111/1911-3846.12832",
    "Hummel, K., & Schlick, C. (2016). The relationship between sustainability performance and sustainability disclosure—Reconciling voluntary disclosure theory and legitimacy theory. Journal of Accounting and Public Policy, 35(5), 455–476. https://doi.org/10.1016/j.jaccpubpol.2016.06.001",
    "Kang, H., & Kim, J. (2022). Analyzing and visualizing text information in corporate sustainability reports using natural language processing methods. Applied Sciences, 12(11), Article 5614. https://doi.org/10.3390/app12115614",
    "Loughran, T., & McDonald, B. (2011). When is a liability not a liability? Textual analysis, dictionaries, and 10-Ks. The Journal of Finance, 66(1), 35–65. https://doi.org/10.1111/j.1540-6261.2010.01625.x",
    "Luccioni, S. A., Baylor, E., & Duchene, N. (2020). Analyzing sustainability reports using natural language processing. In NeurIPS 2020 Workshop on Tackling Climate Change with Machine Learning. arXiv:2011.08073.",
    "Lyon, T. P., & Montgomery, A. W. (2015). The means and end of greenwash. Organization & Environment, 28(2), 223–249. https://doi.org/10.1177/1086026615575332",
    "Matsui, T., Suzuki, M., & Seki, K. (2022). Automated classification of corporate statements for the United Nations sustainable development goals. Sustainability, 14(16), Article 10149. https://doi.org/10.3390/su141610149",
    "Mercereau, B., & Melin, L. (2020). ESG analysis: An NLP approach to assessing corporate sustainability disclosures. The Journal of Investing, 29(7), 50–63. https://doi.org/10.3905/joi.2020.1.157",
    "Merkl-Davies, N. O., & Brennan, N. M. (2007). Discretionary disclosure strategies in corporate narratives: Incremental information or impression management? Journal of Accounting Literature, 26, 116–196.",
    "Mikolov, T., Chen, K., Corrado, G., & Dean, J. (2013). Efficient estimation of word representations in vector space. In Proceedings of ICLR 2013. arXiv:1301.3781.",
    "Ministry of Finance of Vietnam. (2020). Circular No. 96/2020/TT-BTC guiding information disclosure on the securities market.",
    "Ministry of Planning and Investment of Vietnam. (2023). The 2nd voluntary national review on the implementation of the sustainable development goals (VNR 2023). Statistical Publishing House.",
    "Muñoz-Torres, M. J., Fernández-Izquierdo, M. Á., Rivera-Lirio, J. M., & Escrig-Olmedo, E. (2019). Can modern sustainability reports track the SDGs? An assessment framework. Sustainability, 11(5), Article 1421. https://doi.org/10.3390/su11051421",
    "Nguyen, D. Q., & Nguyen, A. T. (2020). PhoBERT: Pre-trained language models for Vietnamese. In Findings of EMNLP 2020 (pp. 1037–1042). Association for Computational Linguistics. https://doi.org/10.18653/v1/2020.findings-emnlp.92",
    "Pennington, J., Socher, R., & Manning, C. D. (2014). GloVe: Global vectors for word representation. In Proceedings of EMNLP 2014 (pp. 1532–1543). Association for Computational Linguistics. https://doi.org/10.3115/v1/D14-1162",
    "Pizzi, S., Caputo, A., Corvino, A., & Ficco, A. (2020). Management research and the UN sustainable development goals (SDGs): A bibliometric investigation and systematic review. Journal of Cleaner Production, 276, Article 124033. https://doi.org/10.1016/j.jclepro.2020.124033",
    "PwC Vietnam. (2022). ESG readiness report in Vietnam 2022: From ambition to impact. PwC Vietnam.",
    "Reimers, N., & Gurevych, I. (2019). Sentence-BERT: Sentence embeddings using Siamese BERT-networks. In Proceedings of EMNLP-IJCNLP 2019 (pp. 3982–3992). Association for Computational Linguistics. https://doi.org/10.18653/v1/D19-1410",
    "Sanh, V., Debut, L., Chaumond, J., & Wolf, T. (2019). DistilBERT, a distilled version of BERT: smaller, faster, cheaper and lighter. arXiv preprint arXiv:1910.01108.",
    "Seele, P., & Gatti, L. (2017). Greenwashing revisited: In search of a typology and accusation-based definition. Business Strategy and the Environment, 26(2), 239–252. https://doi.org/10.1002/bse.1912",
    "Spence, M. (1973). Job market signaling. The Quarterly Journal of Economics, 87(3), 355–374. https://doi.org/10.2307/1882010",
    "Stacchezzini, R., Melloni, G., & Lai, A. (2016). Sustainability management and reporting: The role of impression management for corporate social responsibility disclosure. Journal of Cleaner Production, 136, 102–110. https://doi.org/10.1016/j.jclepro.2016.04.095",
    "State Securities Commission of Vietnam. (2024). Handbook on ESG implementation and disclosure for listed companies. Finance Publishing House.",
    "Suchman, M. C. (1995). Managing legitimacy: Strategic and institutional approaches. Academy of Management Review, 20(3), 571–610. https://doi.org/10.5465/amr.1995.9508080331",
    "Tran, M., & Beddewela, E. (2020). Evaluating the quality of CSR disclosure in Vietnam: An empirical examination of listed firms. Journal of Business Ethics, 166(3), 569–589. https://doi.org/10.1007/s10551-019-04135-2",
    "United Nations. (2015). Transforming our world: The 2030 agenda for sustainable development (Resolution A/RES/70/1). United Nations General Assembly.",
    "VBCSD. (2024). Corporate sustainability index report (CSI 2024). Vietnam Chamber of Commerce and Industry.",
    "Veenstra, E. M., & Ellemers, N. (2020). CSR does not equal investment in CSR: A linguistic analysis of corporate social responsibility reports. Journal of Business Ethics, 161(2), 347–363. https://doi.org/10.1007/s10551-018-3904-7",
    "Wang, W., Wei, F., Dong, L., Bao, H., Yang, N., & Zhou, M. (2020). MiniLM: Deep self-attention distillation for task-agnostic compression of pre-trained transformers. In Advances in Neural Information Processing Systems (Vol. 33, pp. 5776–5788). Curran Associates, Inc.",
    "World Commission on Environment and Development. (1987). Our common future. Oxford University Press."
]

print("Data tables and APA references loaded.")
