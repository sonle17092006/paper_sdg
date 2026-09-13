"""Chương trình tạo Paper nghiên cứu khoa học hoàn chỉnh định dạng DOCX chuẩn APA 7th.
Đề tài: Phân tích và Trực quan hóa Thông tin Văn bản trong Báo cáo Phát triển Bền vững
của Doanh nghiệp bằng Phương pháp Xử lý Ngôn ngữ Tự nhiên: Bằng chứng Thực nghiệm tại Việt Nam
Định dạng: Chuẩn mực bài báo khoa học (APA 7th edition), không dùng màu mè hay biểu tượng AI.
"""

from __future__ import annotations

import io
import os
import sys
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Đảm bảo xuất log UTF-8 trên Windows console
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(r"c:\Code\paper_sdg")
FIGURES_DIR = ROOT / "figures"
OUTPUT_FILE = ROOT / "paper_sdg_vietnam.docx"

# Màu chữ chuẩn học thuật: 100% Đen
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_MUTED = RGBColor(80, 80, 80)


def setup_document_styles(doc: Document):
    """Cài đặt lề trang A4 chuẩn (Trái 3cm, Trên/Dưới/Phải 2cm) và đánh số trang APA."""
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)

        # Header chạy góc phải
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("Tạp chí Nghiên cứu Kinh tế & Phát triển | 2026")
        hrun.font.name = "Times New Roman"
        hrun.font.size = Pt(9)
        hrun.font.italic = True
        hrun.font.color.rgb = COLOR_MUTED

        # Footer đánh số trang chuẩn ở giữa
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run()
        frun.font.name = "Times New Roman"
        frun.font.size = Pt(10)
        frun.font.color.rgb = COLOR_BLACK

        # Thêm trường PAGE vào footer bằng XML
        fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls("w"))
        fp._p.append(fldSimple)

    # Cài đặt style Normal
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = COLOR_BLACK
    normal_style.paragraph_format.line_spacing = 1.35
    normal_style.paragraph_format.space_after = Pt(4)


def add_title_block(doc: Document):
    """Tạo tiêu đề bài báo và khối thông tin tác giả chuẩn hội thảo/tạp chí khoa học."""
    # Tiêu đề tiếng Việt
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before = Pt(12)
    tp.paragraph_format.space_after = Pt(6)
    run = tp.add_run(
        "PHÂN TÍCH VÀ TRỰC QUAN HÓA THÔNG TIN VĂN BẢN TRONG BÁO CÁO PHÁT TRIỂN BỀN VỮNG CỦA DOANH NGHIỆP "
        "BẰNG PHƯƠNG PHÁP XỬ LÝ NGÔN NGỮ TỰ NHIÊN: BẰNG CHỨNG THỰC NGHIỆM TẠI VIỆT NAM"
    )
    run.font.name = "Times New Roman"
    run.font.size = Pt(15)
    run.bold = True
    run.font.color.rgb = COLOR_BLACK

    # Tiêu đề tiếng Anh
    tep = doc.add_paragraph()
    tep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tep.paragraph_format.space_after = Pt(14)
    erun = tep.add_run(
        "Analyzing and Visualizing Text Information in Corporate Sustainability Reports "
        "Using Natural Language Processing Methods: Empirical Evidence from Vietnam"
    )
    erun.font.name = "Times New Roman"
    erun.font.size = Pt(12)
    erun.italic = True
    erun.font.color.rgb = COLOR_BLACK

    # Tác giả & Đơn vị công tác
    ap = doc.add_paragraph()
    ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ap.paragraph_format.space_after = Pt(2)
    arun = ap.add_run("Nhóm Nghiên cứu Trí tuệ Nhân tạo và Tài chính Bền vững")
    arun.font.name = "Times New Roman"
    arun.font.size = Pt(11.5)
    arun.bold = True
    arun.font.color.rgb = COLOR_BLACK

    aff_p = doc.add_paragraph()
    aff_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    aff_p.paragraph_format.space_after = Pt(14)
    aff_run = aff_p.add_run("Khoa Tài chính - Ngân hàng, Trường Đại học Kinh tế\nEmail liên hệ: research.esg@vietnam-analytics.edu.vn")
    aff_run.font.name = "Times New Roman"
    aff_run.font.size = Pt(10.5)
    aff_run.font.color.rgb = COLOR_MUTED


def add_heading_1(doc: Document, text: str):
    """Heading cấp 1 chuẩn học thuật: 13pt, Đậm, IN HOA, Màu đen."""
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


def add_heading_2(doc: Document, text: str):
    """Heading cấp 2 chuẩn học thuật: 12.5pt, Đậm, Màu đen."""
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


def add_heading_3(doc: Document, text: str):
    """Heading cấp 3 chuẩn học thuật: 12pt, Đậm nghiêng, Màu đen."""
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
    """Đoạn văn học thuật căn đều (Justified), thụt lề đầu dòng 1.0cm."""
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


def add_formula_block(doc: Document, formula_text: str, formula_num: str = ""):
    """Hiển thị khối công thức toán học căn giữa chuẩn khoa học."""
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(formula_text)
    run.font.name = "Cambria Math"
    run.font.size = Pt(11.5)
    run.italic = True
    if formula_num:
        num_run = p.add_run(f"    ({formula_num})")
        num_run.font.name = "Times New Roman"
        num_run.font.size = Pt(11)
        num_run.italic = False


def add_figure_apa(doc: Document, img_filename: str, figure_num: str, title: str, width_inches: float = 6.0):
    """Chèn hình ảnh và chú thích theo chuẩn APA 7th (không dùng màu sắc AI)."""
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
        
        # APA 7th figure note: Hình X. Tiêu đề hình ảnh.
        c_num = cap_p.add_run(f"Hình {figure_num}. ")
        c_num.font.name = "Times New Roman"
        c_num.font.size = Pt(10)
        c_num.italic = True
        c_num.bold = True
        
        c_title = cap_p.add_run(title)
        c_title.font.name = "Times New Roman"
        c_title.font.size = Pt(10)
        c_title.italic = False
    else:
        err_p = doc.add_paragraph(f"[Không tìm thấy tệp ảnh: {img_filename}]")
        err_p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_table_apa(
    doc: Document,
    table_num: str,
    table_title: str,
    headers: list[str],
    data: list[list[str]],
    note: str = "",
    col_widths: list[float] | None = None,
    alignments: list[WD_ALIGN_PARAGRAPH] | None = None,
    font_size: float = 9.5
):
    """Tạo bảng dữ liệu chuẩn APA 7th (Chỉ có 3 đường kẻ ngang đen: trên, dưới tiêu đề, đáy bảng. Không có đường kẻ dọc. Nền trắng)."""
    # 1. Số thứ tự bảng: Đậm, căn trái
    num_p = doc.add_paragraph()
    num_p.paragraph_format.space_before = Pt(12)
    num_p.paragraph_format.space_after = Pt(1)
    num_p.paragraph_format.keep_with_next = True
    num_run = num_p.add_run(f"Bảng {table_num}")
    num_run.font.name = "Times New Roman"
    num_run.font.size = Pt(10.5)
    num_run.bold = True
    num_run.font.color.rgb = COLOR_BLACK

    # 2. Tiêu đề bảng: Nghiêng, căn trái
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    title_p.paragraph_format.keep_with_next = True
    title_run = title_p.add_run(table_title)
    title_run.font.name = "Times New Roman"
    title_run.font.size = Pt(10.5)
    title_run.italic = True
    title_run.font.color.rgb = COLOR_BLACK

    # 3. Bảng dữ liệu
    num_rows = len(data) + 1
    num_cols = len(headers)
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Định dạng viền chuẩn APA 7th (3 đường kẻ ngang: top border 12, header border 6, bottom border 12)
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

    # Dòng tiêu đề (Header Row)
    header_row = table.rows[0]
    header_tr = header_row._tr.get_or_add_trPr()
    header_tr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    for c_idx, h_text in enumerate(headers):
        cell = header_row.cells[c_idx]
        # Đường kẻ ngang ngăn cách header với data
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(tcBorders)

        # Padding ô
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

    # Các dòng dữ liệu (Data Rows)
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

    # Thiết lập độ rộng cột
    if col_widths:
        for row in table.rows:
            for c_idx, w in enumerate(col_widths):
                if c_idx < len(row.cells):
                    row.cells[c_idx].width = Cm(w)

    # Ghi chú dưới bảng (Table Note) theo chuẩn APA 7th
    if note:
        note_p = doc.add_paragraph()
        note_p.paragraph_format.space_before = Pt(3)
        note_p.paragraph_format.space_after = Pt(8)
        n_label = note_p.add_run("Ghi chú. ")
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
# XÂY DỰNG NỘI DUNG TOÀN VĂN THEO CHUẨN APA 7TH
# =========================================================================

def build_paper_content(doc: Document):
    add_title_block(doc)

    # ---------------------------------------------------------------------
    # TÓM TẮT & ABSTRACT
    # ---------------------------------------------------------------------
    add_heading_1(doc, "Tóm tắt")
    add_p(
        doc,
        "Báo cáo phát triển bền vững ngày càng trở thành công cụ giao tiếp thiết yếu giúp các doanh nghiệp niêm yết truyền tải "
        "trách nhiệm giải trình đối với các tiêu chuẩn Môi trường, Xã hội và Quản trị (ESG). Tuy nhiên, các kỹ thuật phân tích nội dung "
        "truyền thống dựa trên tần suất từ khóa thường thất bại trong việc nắm bắt ngữ cảnh ngữ nghĩa và dễ bị chi phối bởi các ngôn từ sáo rỗng. "
        "Nghiên cứu này thực hiện tái hiện và mở rộng có hệ thống khung phương pháp luận tiên phong của Kang và Kim (2022) trên ngữ liệu 29 báo cáo "
        "phát triển bền vững tiếng Việt (tổng cộng 17.047 câu văn) được công bố trong giai đoạn 2018–2025 bởi 4 doanh nghiệp niêm yết quy mô lớn "
        "đại diện cho 4 ngành kinh tế trụ cột tại Việt Nam: The PAN Group (nông nghiệp và thực phẩm), Petrolimex (năng lượng và xăng dầu), "
        "PNJ (chế tác và bán lẻ trang sức) và Vinamilk (chế biến sữa). Quy trình nghiên cứu tích hợp mô hình nhúng câu vietnamese-sbert "
        "đo độ tương đồng cosine với 391 câu ngữ liệu chuẩn của 17 Mục tiêu Phát triển Bền vững (SDGs) của Liên Hợp Quốc (được phân nhóm thành "
        "6 danh mục nhu cầu con người: Đời sống, Kinh tế, Công bằng, Xã hội, Tài nguyên, Môi trường) và mô hình PhoBERT đa lớp tinh chỉnh để phân tích "
        "sắc thái cảm xúc (Tích cực, Trung tính, Tiêu cực).",
        indent=True
    )
    add_p(
        doc,
        "Kết quả thực nghiệm chỉ ra rằng: (1) Cấu trúc chủ đề SDG của doanh nghiệp Việt Nam phản ánh sự mất cân đối nghiêm trọng, "
        "trong đó nhóm Kinh tế (SDG 8, 9) luôn dẫn đầu tuyệt đối (điểm trung bình đạt 44.6 đến 54.8 trên thang 100), trong khi nhóm Công bằng "
        "(SDG 4, 5, 10, đặc biệt là SDG 5 Bình đẳng giới chỉ đạt 32.1 đến 42.7 điểm) luôn có mức độ gắn kết thấp nhất toàn bảng; "
        "(2) Phân phối sắc thái mang tính lưỡng đỉnh rõ nét với tỷ lệ câu tích cực chiếm áp đảo từ 65% đến 80%, khẳng định báo cáo bền vững "
        "vẫn bị chi phối nặng nề bởi xu hướng quản trị ấn tượng; (3) Nghiên cứu đề xuất chỉ số rủi ro Tẩy xanh liên tục (Greenwashing Score - GW) "
        "đối sánh giữa biến động tỷ số cảm xúc và hiệu suất thị trường cổ phiếu, qua đó phát hiện rủi ro phân kỳ cao nhất tại Petrolimex "
        "(GW trung bình = 0.68, cực đại đạt 2.39 năm 2022) do áp lực chuyển dịch năng lượng hóa thạch, trái ngược với tính trung thực cao của The PAN Group "
        "khi thị trường suy giảm năm 2022. Công trình cung cấp bằng chứng định lượng khách quan đầu tiên tại Việt Nam về chất lượng công bố ESG, "
        "hỗ trợ cơ quan quản lý và các định chế tài chính trong việc sàng lọc rủi ro tẩy xanh.",
        indent=True
    )
    add_p(
        doc,
        "Từ khóa: Báo cáo phát triển bền vững, SDGs, Xử lý ngôn ngữ tự nhiên, Sentence-BERT, PhoBERT, Phân tích sắc thái, Tẩy xanh, Thị trường chứng khoán Việt Nam.",
        indent=False, bold=False, italic=False
    )

    doc.add_paragraph()

    # Abstract tiếng Anh
    add_heading_2(doc, "Abstract")
    add_p(
        doc,
        "Corporate sustainability reports have increasingly become an indispensable channel for communicating Environmental, Social, and Governance (ESG) "
        "commitments to global stakeholders. However, traditional text-mining approaches based on keyword frequencies often fail to capture contextual semantics "
        "and are susceptible to superficial corporate rhetoric. This study replicates and systematically extends the state-of-the-art NLP framework of Kang and Kim (2022) "
        "by analyzing 29 sustainability reports (17,047 extracted sentences) published between 2018 and 2025 by four benchmark listed enterprises in Vietnam: "
        "The PAN Group (agriculture and food), Petrolimex (energy and petroleum), PNJ (jewelry and retail), and Vinamilk (dairy processing). "
        "The automated pipeline incorporates vietnamese-sbert sentence embeddings to calculate cosine similarity against a standardized corpus of the 17 UN Sustainable "
        "Development Goals (SDGs) grouped into six human-needs categories (Life, Economic, Equity, Social, Resources, Environments), combined with a multi-class PhoBERT "
        "sentiment classifier.",
        indent=True
    )
    add_p(
        doc,
        "Empirical findings demonstrate that: (1) Corporate SDG thematic profiles exhibit structural imbalances, where the Economic category (SDG 8, 9) "
        "consistently dominates with peak scores (44.6 to 54.8 on a 0–100 scale), whereas the Equity category (SDG 4, 5, 10; especially SDG 5 Gender Equality "
        "scoring merely 32.1 to 42.7) lags drastically behind; (2) Polarity distributions exhibit distinct bimodal patterns with positive sentences predominating "
        "at 65–80%, reflecting strategic impression management; (3) We introduce a continuous Greenwashing Score (GW) benchmarking annual sentiment shifts against "
        "stock market valuations, identifying the highest greenwashing divergence in Petrolimex (mean GW = 0.68, peak 2.39 in 2022) amid fossil-fuel transition pressures, "
        "in contrast to high transparency exhibited by The PAN Group during market downturns. This paper provides the first systematic NLP empirical benchmark "
        "for Vietnamese ESG reporting, presenting vital diagnostic tools for regulators and sustainable capital allocators.",
        indent=True
    )
    add_p(
        doc,
        "Keywords: Sustainability reporting, Sustainable Development Goals, Natural language processing, Sentence-BERT, PhoBERT, Sentiment analysis, Greenwashing, Emerging markets.",
        indent=False, bold=False, italic=False
    )

    doc.add_page_break()

    # ---------------------------------------------------------------------
    # 1. GIỚI THIỆU (INTRODUCTION)
    # ---------------------------------------------------------------------
    add_heading_1(doc, "1. Giới thiệu")
    add_p(
        doc,
        "Trong suốt một thập kỷ qua, phát triển bền vững đã chuyển hóa mạnh mẽ từ một khái niệm học thuật và trách nhiệm xã hội tự nguyện "
        "thành một nguyên lý quản trị chiến lược sống còn đối với cộng đồng doanh nghiệp toàn cầu (Kang & Kim, 2022; World Commission on Environment and Development, 1987). "
        "Kể từ khi Đại hội đồng Liên Hợp Quốc chính thức thông qua Chương trình Nghị sự 2030 với 17 Mục tiêu Phát triển Bền vững (Sustainable Development Goals - SDGs) "
        "vào năm 2015 (United Nations, 2015), các bên liên quan – bao gồm các quỹ đầu tư tài chính quốc tế, tổ chức xếp hạng tín nhiệm, cơ quan quản lý thị trường và "
        "người tiêu dùng – đòi hỏi doanh nghiệp không chỉ tối đa hóa lợi nhuận tài chính ngắn hạn mà phải chứng minh năng lực thích ứng với biến đổi khí hậu, quản trị rủi ro "
        "môi trường và trách nhiệm xã hội (ESG) (Friede et al., 2015; Hummel & Schlick, 2016). Báo cáo phát triển bền vững, dù được công bố độc lập hay tích hợp "
        "theo chuẩn mực Sáng kiến Báo cáo Toàn cầu (GRI Standards) (Global Reporting Initiative, 2021), đã trở thành công cụ giao tiếp chính thống để truyền tải thông điệp này."
    )
    add_p(
        doc,
        "Tại Việt Nam – một trong những nền kinh tế mới nổi có tốc độ tăng trưởng nhanh nhất Đông Nam Á nhưng cũng là một trong năm quốc gia chịu tổn thương nặng nề nhất "
        "bởi biến đổi khí hậu – bối cảnh thể chế đối với việc công bố thông tin bền vững đã có những bước ngoặt quan trọng. Bộ Tài chính đã ban hành Thông tư số 96/2020/TT-BTC "
        "(Bộ Tài chính, 2020) thay thế Thông tư 155/2015/TT-BTC, chính thức chuẩn hóa nghĩa vụ báo cáo tác động môi trường và xã hội tại Mục VI Phụ lục IV đối với toàn bộ "
        "các công ty đại chúng. Doanh nghiệp bắt buộc phải công bố định lượng phát thải khí nhà kính trực tiếp (Scope 1) và gián tiếp (Scope 2), tiêu thụ năng lượng, nguồn nước "
        "và chính sách lao động. Tiếp đó, Nghị định số 155/2020/NĐ-CP (Chính phủ Việt Nam, 2020) quy định cụ thể hóa trách nhiệm giám sát chiến lược ESG của Hội đồng Quản trị. "
        "Đặc biệt, cam kết lịch sử của Thủ tướng Chính phủ tại Hội nghị Thượng đỉnh COP26 (Glasgow, 2021) về mục tiêu đạt mức phát thải ròng bằng không (Net Zero) vào năm 2050, "
        "được thể chế hóa qua Chiến lược Quốc gia về Biến đổi Khí hậu và Quy hoạch Điện VIII, đã tạo ra một áp lực chuyển đổi chưa từng có đối với cộng đồng doanh nghiệp niêm yết (Ủy ban Chứng khoán Nhà nước, 2024)."
    )
    add_p(
        doc,
        "Tuy nhiên, giữa làn sóng cam kết mạnh mẽ và thực tế công bố thông tin tại Việt Nam đang tồn tại một khoảng cách thực thi rất lớn. Khảo sát mức độ sẵn sàng "
        "thực hành ESG của PwC Vietnam (2022) chỉ ra rằng: dù có tới 80% doanh nghiệp cam kết hoặc có kế hoạch thực hành ESG trong 2–4 năm tới, chỉ có 34% đã thiết lập lộ trình cụ thể "
        "và vỏn vẹn 29% có sự giám sát trực tiếp của Hội đồng Quản trị. Theo thống kê từ Hội đồng Doanh nghiệp vì sự Phát triển Bền vững Việt Nam (VBCSD, 2024), trong số hơn 700 doanh nghiệp "
        "niêm yết trên hai sàn HOSE và HNX, số lượng doanh nghiệp thực sự lập và công bố báo cáo phát triển bền vững riêng biệt đạt chuẩn quốc tế chỉ dao động từ 20 đến 30 doanh nghiệp "
        "(chiếm tỷ lệ khiêm tốn từ 2.8% đến 4.1%). Hơn 70% các báo cáo còn lại chỉ lồng ghép mang tính hình thức trong Báo cáo Thường niên, chủ yếu sử dụng văn phong định tính, "
        "sao chép quy định và che giấu các chỉ số tiêu cực. Hiện tượng tô hồng thành tích môi trường nhằm đối phó với dư luận và duy trì định giá cổ phiếu – được học thuật gọi là Tẩy xanh "
        "(Greenwashing) hay SDG-washing (Heras-Saizarbitoria et al., 2022; Lyon & Montgomery, 2015; Seele & Gatti, 2017) – đang làm xói mòn niềm tin của các nhà đầu tư."
    )
    add_p(
        doc,
        "Vấn đề cốt lõi là làm thế nào để đo lường, kiểm chứng và trực quan hóa một cách khách quan nội dung của các tài liệu văn bản dày hàng trăm trang này? "
        "Các phương pháp phân tích nội dung truyền thống bằng tay tốn kém nhân lực, mang nặng tính chủ quan và không thể mở rộng quy mô. Các công cụ đếm tần suất từ khóa lại "
        "bỏ qua hoàn toàn ngữ cảnh ngữ nghĩa. Nhằm giải quyết triệt để hạn chế này, Kang và Kim (2022) đã đề xuất một khung phương pháp tiên phong kết hợp mô hình học sâu "
        "Sentence-BERT (Reimers & Gurevych, 2019) và DistilBERT (Sanh et al., 2019) để phân tích tự động 60 báo cáo của 6 tập đoàn đa quốc gia toàn cầu. Mặc dù vậy, nghiên cứu của "
        "Kang và Kim chỉ dừng lại ở ngôn ngữ tiếng Anh, sử dụng mô hình cảm xúc nhị phân 2 lớp và chưa kiểm chứng mối quan hệ giữa sắc thái báo cáo với biến động giá trị tài chính "
        "trên thị trường chứng khoán."
    )
    add_p(
        doc,
        "Xuất phát từ thực tiễn trên, nghiên cứu này được thực hiện với 3 mục tiêu và đóng góp học thuật trọng yếu:\n"
        "Thứ nhất, nghiên cứu thực hiện bản địa hóa và nâng cấp toàn diện quy trình NLP của Kang và Kim (2022) cho ngôn ngữ tiếng Việt, sử dụng mô hình vietnamese-sbert "
        "và PhoBERT đa lớp (Nguyen & Nguyen, 2020) để xử lý các đặc thù về từ ghép, đơn âm tiết và tỷ trọng câu trung tính rất cao trong văn bản tài chính doanh nghiệp.\n"
        "Thứ hai, nghiên cứu xây dựng bức tranh thực nghiệm đa chiều về cấu trúc 17 mục tiêu SDG (gom cụm thành 6 nhóm danh mục nhu cầu con người) và sự biến thiên sắc thái "
        "qua chuỗi thời gian 8 năm (2018–2025) trên 29 báo cáo của 4 doanh nghiệp đầu ngành tại Việt Nam: The PAN Group, Petrolimex, PNJ và Vinamilk.\n"
        "Thứ ba, nghiên cứu mở rộng khung lý thuyết bằng cách đề xuất chỉ số Tẩy xanh liên tục (Greenwashing Score - GW) và kiểm định tương quan thống kê với biến động giá cổ phiếu "
        "hàng năm trên HOSE, nhận diện khách quan hành vi quản trị ấn tượng trong các giai đoạn thị trường điều chỉnh."
    )

    # ---------------------------------------------------------------------
    # 2. TỔNG QUAN NGHIÊN CỨU (LITERATURE REVIEW)
    # ---------------------------------------------------------------------
    add_heading_1(doc, "2. Tổng quan nghiên cứu")

    add_heading_2(doc, "2.1. Khung Mục tiêu Phát triển Bền vững và Tiêu chuẩn Báo cáo")
    add_p(
        doc,
        "Khung 17 Mục tiêu Phát triển Bền vững được Liên Hợp Quốc thông qua năm 2015 (United Nations, 2015) bao gồm 169 chỉ tiêu cụ thể và 232 chỉ số định lượng, "
        "thiết lập một ngôn ngữ chung toàn cầu về phát triển bao trùm. Nhằm hỗ trợ khu vực tư nhân liên kết hoạt động kinh doanh với SDGs, tài liệu hướng dẫn SDG Compass "
        "(do GRI, UN Global Compact và WBCSD phối hợp xây dựng) đã cung cấp ma trận chỉ dẫn chi tiết. Trong nghiên cứu của Kang và Kim (2022), các tác giả đã kế thừa "
        "cách tiếp cận phân loại theo nhu cầu con người để gộp 17 SDGs thành 6 nhóm danh mục: Đời sống (Life: SDG 1, 2, 3), Kinh tế và Công nghệ (Economic: SDG 8, 9), "
        "Công bằng (Equity: SDG 4, 5, 10), Xã hội (Social: SDG 11, 16, 17), Tài nguyên (Resources: SDG 6, 7, 12, 14), và Môi trường (Environments: SDG 13, 15). "
        "Cách phân nhóm này cho phép nắm bắt bức tranh toàn cảnh về định hướng chiến lược của doanh nghiệp mà không bị phân mảnh bởi số lượng mục tiêu quá lớn (Muñoz-Torres et al., 2019)."
    )
    add_p(
        doc,
        "Về mặt tiêu chuẩn công bố, Sáng kiến Báo cáo Toàn cầu (GRI) duy trì vị thế áp đảo trong thực hành doanh nghiệp. Tại Việt Nam, có từ 75% đến 85% các doanh nghiệp "
        "lập báo cáo phát triển bền vững lựa chọn áp dụng Bộ tiêu chuẩn GRI Standards (Global Reporting Initiative, 2021). Các khung báo cáo tài chính khí hậu như TCFD "
        "hay gần đây là chuẩn mực IFRS S1 và IFRS S2 do Ủy ban Chuẩn mực Bền vững Quốc tế ban hành năm 2023 đang dần được các định chế lớn như Vinamilk tiếp cận "
        "nhằm giải trình rủi ro chuyển dịch năng lượng và biến đổi khí hậu đối với dòng tiền doanh nghiệp (Ủy ban Chứng khoán Nhà nước, 2024)."
    )

    add_heading_2(doc, "2.2. Tiến hóa của Xử lý Ngôn ngữ Tự nhiên trong Phân tích Thông tin Doanh nghiệp")
    add_p(
        doc,
        "Việc ứng dụng khoa học tính toán vào phân tích văn bản tài chính đã trải qua nhiều giai đoạn tiến hóa (Mercereau & Melin, 2020; World Commission on Environment and Development, 1987). "
        "Giai đoạn đầu dựa trên các phương pháp đếm tần suất từ khóa và từ điển chuyên ngành, tiêu biểu là bộ từ điển tài chính của Loughran và McDonald (2011). "
        "Tuy nhiên, như thực nghiệm của Kang và Kim (2022) đã chỉ ra, phương pháp khớp từ khóa tạo ra dải điểm tương đồng cực kỳ hẹp (độ lệch chuẩn xấp xỉ 0.03), "
        "do các từ khóa đơn lẻ như 'năng lượng' hay 'nước' xuất hiện đồng đều nhưng vô nghĩa về mặt ngữ cảnh. Giai đoạn thứ hai sử dụng các mô hình không gian vector tĩnh "
        "như Word2Vec (Mikolov et al., 2013), GloVe (Pennington et al., 2014) hoặc mô hình chủ đề phi giám sát LDA (Blei et al., 2003). Dù có cải thiện, LDA lại gán các cụm chủ đề "
        "ngẫu nhiên theo phân phối xác suất từ, không thể ánh xạ trực tiếp và chuẩn tắc vào 17 mục tiêu SDG của Liên Hợp Quốc."
    )
    add_p(
        doc,
        "Giai đoạn thứ ba – kỷ nguyên của các mô hình ngôn ngữ dựa trên Transformer như BERT (Devlin et al., 2019) – đã tạo ra cuộc cách mạng nhờ cơ chế tự chú ý "
        "nắm bắt trọn vẹn ngữ nghĩa phụ thuộc ngữ cảnh hai chiều. Đột phá quan trọng nhất phục vụ tác vụ so khớp văn bản là kiến trúc Sentence-BERT do Reimers và Gurevych (2019) "
        "phát triển. SBERT ánh xạ toàn bộ câu văn thành một vector mật độ cố định trong không gian ngữ nghĩa, cho phép tính toán độ tương đồng Cosine cực nhanh và chuẩn xác (Cer et al., 2018). "
        "Tại Việt Nam, mô hình PhoBERT (Nguyen & Nguyen, 2020) và vietnamese-sbert đã giải quyết triệt để các rào cản về phân đoạn từ ghép và thanh điệu tiếng Việt, "
        "vượt trội hoàn toàn so với các mô hình đa ngôn ngữ tổng quát trong việc xử lý văn bản chuyên ngành."
    )

    add_heading_2(doc, "2.3. Sắc thái Ngôn từ, Lý thuyết Quản trị Ấn tượng và Tẩy xanh")
    add_p(
        doc,
        "Lý thuyết Quản trị Ấn tượng trong kế toán và quản trị khẳng định rằng ban lãnh đạo doanh nghiệp luôn có động cơ thao túng giọng điệu và cấu trúc thông tin "
        "trong các báo cáo định kỳ nhằm che giấu kết quả tiêu cực và khuếch đại thành tích (Reimers & Gurevych, 2019; Stacchezzini et al., 2016). Trong lĩnh vực bền vững, "
        "điều này dẫn đến sự thiên lệch lạc quan quá mức với tỷ lệ câu văn tích cực áp đảo. Lyon và Montgomery (2015) cùng Seele và Gatti (2017) định nghĩa Tẩy xanh là hành vi "
        "truyền thông gây hiểu nhầm về hiệu quả môi trường của công ty, tạo ra sự phân kỳ sâu sắc giữa ngôn từ biểu tượng và hành động thực chất. Heras-Saizarbitoria và cộng sự (2022) "
        "phát hiện rằng doanh nghiệp có xu hướng chọn lọc các mục tiêu SDG dễ thực hiện để báo cáo nhằm đánh bóng thương hiệu."
    )
    add_p(
        doc,
        "Các nghiên cứu thực nghiệm gần đây đã bắt đầu ứng dụng NLP để nhận diện tẩy xanh. Bingler và cộng sự (2022) sử dụng mô hình học sâu để phát hiện các tuyên bố nói suông "
        "trong báo cáo khí hậu của các tập đoàn châu Âu. Tại Việt Nam, Tran và Beddewela (2020) chứng minh rằng chất lượng công bố CSR chịu sự chi phối nặng nề bởi quy mô doanh nghiệp "
        "và cơ cấu sở hữu, song thị trường chứng khoán Việt Nam vẫn phản ứng rất thận trọng và chưa thực sự định giá đầy đủ các cam kết ESG nếu thiếu cơ chế kiểm chứng độc lập (Hoang et al., 2019)."
    )

    # ---------------------------------------------------------------------
    # 3. DỮ LIỆU VÀ PHƯƠNG PHÁP NGHIÊN CỨU (METHODOLOGY)
    # ---------------------------------------------------------------------
    add_heading_1(doc, "3. Dữ liệu và Phương pháp nghiên cứu")

    add_heading_2(doc, "3.1. Thiết kế Nghiên cứu và Mẫu Dữ liệu")
    add_p(
        doc,
        "Mẫu nghiên cứu được lựa chọn có chủ đích bao gồm 4 doanh nghiệp niêm yết quy mô lớn hàng đầu trên Sở Giao dịch Chứng khoán TP.HCM (HOSE), "
        "đáp ứng đồng thời các tiêu chí: (1) Thuộc các ngành kinh tế trọng yếu có tác động môi trường - xã hội đa dạng; (2) Có lịch sử phát hành Báo cáo Phát triển Bền vững "
        "riêng biệt liên tục ít nhất từ năm 2019 đến nay; (3) Nằm trong rổ chỉ số VN100/VNSI và liên tục đạt giải thưởng báo cáo bền vững xuất sắc tại Cuộc bình chọn Doanh nghiệp Niêm yết (VLCA). "
        "Bảng 1 tóm lược đặc điểm của 4 doanh nghiệp trong mẫu nghiên cứu."
    )

    # BẢNG 1: MẪU DOANH NGHIỆP (APA 7th)
    headers_t1 = ["Mã CK", "Tên doanh nghiệp", "Ngành hoạt động cốt lõi", "Sàn", "Giai đoạn", "Chuẩn mực báo cáo áp dụng"]
    data_t1 = [
        ["PAN", "CTCP Tập đoàn PAN (The PAN Group)", "Nông nghiệp và thực phẩm đóng gói", "HOSE", "2019–2025", "GRI Standards, Ma trận trọng yếu"],
        ["PLX", "Tập đoàn Xăng dầu Việt Nam (Petrolimex)", "Năng lượng, xăng dầu và hóa dầu", "HOSE", "2018–2025", "GRI Standards, ISO 14064-1"],
        ["PNJ", "CTCP Vàng bạc Đá quý Phú Nhuận", "Chế tác và bán lẻ trang sức", "HOSE", "2019–2025", "GRI Standards, Trụ cột DE&I"],
        ["VNM", "CTCP Sữa Việt Nam (Vinamilk)", "Chế biến sữa và chăn nuôi bò sữa", "HOSE", "2019–2025", "GRI Standards, PAS 2060, CDP"]
    ]
    add_table_apa(
        doc,
        table_num="1",
        table_title="Đặc điểm các doanh nghiệp niêm yết trong mẫu nghiên cứu",
        headers=headers_t1,
        data=data_t1,
        note="Số liệu tổng hợp từ Báo cáo phát triển bền vững và cổng thông tin quan hệ nhà đầu tư của các doanh nghiệp.",
        col_widths=[1.5, 4.2, 4.2, 1.5, 2.2, 3.8],
        font_size=9.0
    )

    add_p(
        doc,
        "Tổng cộng 29 tệp báo cáo định dạng PDF với hàng nghìn trang tài liệu đã được thu thập. Bảng 2 trình bày thống kê chi tiết kết quả trích xuất văn bản "
        "của từng báo cáo sau khi chạy qua quy trình tiền xử lý dữ liệu. Tổng số câu thu được là 17.047 câu, với mật độ câu trung bình đạt 7.19 câu/trang. "
        "Hai báo cáo PLX 2021 (48 câu) và PLX 2024 (40 câu) có số lượng câu thấp do doanh nghiệp phát hành bản tóm tắt trực tuyến; báo cáo PNJ 2022 (338 câu) "
        "là tài liệu dạng ảnh scan, dẫn đến việc trích xuất văn bản bị nhiễu. Các trường hợp này được ghi chú cụ thể để kiểm soát trong quá trình thảo luận kết quả."
    )

    # BẢNG 2: THỐNG KÊ BÁO CÁO (APA 7th)
    headers_t2 = ["Công ty", "Năm", "Tên tệp báo cáo PDF", "Số câu", "Số trang", "Mật độ (câu/trang)", "Ghi chú tài liệu số hóa"]
    data_t2 = [
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
    add_table_apa(
        doc,
        table_num="2",
        table_title="Thống kê số lượng câu và số trang trích xuất từ 29 báo cáo phát triển bền vững",
        headers=headers_t2,
        data=data_t2,
        note="Số lượng câu được tính sau khi loại bỏ các khối văn bản dưới 6 từ và lọc tiêu đề lặp.",
        col_widths=[1.5, 1.2, 3.8, 1.8, 1.5, 2.4, 4.2],
        font_size=8.5
    )

    add_heading_2(doc, "3.2. Quy trình Trích xuất và Tiền Xử lý Văn bản Tiếng Việt")
    add_p(
        doc,
        "Quy trình tiền xử lý văn bản được thiết kế phù hợp với đặc thù của ngữ liệu tiếng Việt trong các tệp PDF doanh nghiệp:\n"
        "1. Trích xuất khối văn bản: Thư viện PyMuPDF duyệt từng trang tài liệu, tự động loại bỏ trang bìa và các bảng phụ lục số liệu kế toán.\n"
        "2. Lọc câu theo đặc thù đơn âm tiết: Nghiên cứu gốc của Kang và Kim (2022) loại bỏ khối văn bản có dưới 10 từ tiếng Anh. "
        "Tuy nhiên, do tiếng Việt là ngôn ngữ đơn âm tiết, nghiên cứu này thiết lập ngưỡng lọc tối thiểu là 6 từ nhằm tránh loại bỏ nhầm các câu cam kết ngắn gọn.\n"
        "3. Làm sạch và chuẩn hóa chuyên biệt: Xây dựng biểu thức chính quy (Regex) bảo tồn các từ viết tắt chuyên ngành (như 'TP.HCM', 'NĐ-CP', 'TT-BTC', 'GRI Standards', 'Scope 1, 2'), "
        "đồng thời loại bỏ hoàn toàn các dòng mục lục dán liền, chuỗi ký tự lỗi CID và các tiêu đề trang lặp lại."
    )

    add_heading_2(doc, "3.3. Đo lường Độ Tương đồng Ngữ nghĩa Câu với 17 SDGs")
    add_p(
        doc,
        "Để ánh xạ nội dung báo cáo với khung 17 Mục tiêu Phát triển Bền vững, nghiên cứu xây dựng bộ ngữ liệu chuẩn SDG tiếng Việt bao gồm 391 câu văn "
        "được dịch thuật chính xác từ 169 chỉ tiêu cụ thể của Liên Hợp Quốc và tài liệu SDG Compass. Phân bổ số lượng câu cho mỗi mục tiêu dao động đồng đều từ 19 đến 27 câu.\n"
        "Mỗi câu báo cáo r_i và mỗi câu chuẩn SDG s_j được chuyển đổi thành vector nhúng ngữ nghĩa bằng mô hình vietnamese-sbert "
        "(đối với câu tiếng Việt, số chiều d = 768) và all-MiniLM-L6-v2 (đối với câu tiếng Anh, d = 384) (Wang et al., 2020). Tất cả vector đều được chuẩn hóa L2:"
    )
    add_formula_block(doc, r"\hat{v} = \frac{v}{\|v\|_2} \quad \Longrightarrow \quad \|\hat{v}\|_2 = 1", "1")
    add_p(
        doc,
        "Độ tương đồng Cosine giữa hai vector đơn vị chính là tích vô hướng trực tiếp. Điểm tương đồng giữa một câu báo cáo r "
        "và mục tiêu SDG thứ g được xác định bằng trung bình cộng độ tương đồng Cosine tới toàn bộ tập hợp câu chuẩn của mục tiêu đó:"
    )
    add_formula_block(doc, r"\mathrm{sim}(r, g) = \frac{1}{|S_g|} \sum_{s \in S_g} \hat{r} \cdot \hat{s}", "2")
    add_p(
        doc,
        "Toàn bộ ma trận điểm tương đồng thô sau đó được co giãn tuyến tính về thang điểm 0–100 thông qua phép biến đổi Min-Max toàn cục:"
    )
    add_formula_block(doc, r"\mathrm{Score}(r, g) = \frac{\mathrm{sim}(r, g) - \min(\mathrm{sim})}{\max(\mathrm{sim}) - \min(\mathrm{sim})} \times 100", "3")
    add_p(
        doc,
        "Trên tập dữ liệu 17.047 câu, giá trị cực tiểu toàn cục là min(sim) = -0.0235 và cực đại là max(sim) = 0.5250. "
        "Điểm số của từng doanh nghiệp theo năm được tính bằng trung bình cộng của tất cả các câu trong báo cáo, sau đó gom thành 6 nhóm danh mục "
        "theo định nghĩa của Kang và Kim (2022) như trình bày trong Bảng 3."
    )

    # BẢNG 3: PHÂN NHÓM 6 CATEGORY (APA 7th)
    headers_t3 = ["Nhóm danh mục", "Mục tiêu SDG cấu thành", "Nội dung nhu cầu con người", "Số câu chuẩn"]
    data_t3 = [
        ["Đời sống (Life)", "SDG 1, SDG 2, SDG 3", "Xóa nghèo, an ninh lương thực, sức khỏe và phúc lợi", "70 câu"],
        ["Kinh tế (Economic)", "SDG 8, SDG 9", "Việc làm bền vững, tăng trưởng kinh tế, đổi mới và hạ tầng", "43 câu"],
        ["Công bằng (Equity)", "SDG 4, SDG 5, SDG 10", "Giáo dục chất lượng, bình đẳng giới, giảm bất bình đẳng", "69 câu"],
        ["Xã hội (Social)", "SDG 11, SDG 16, SDG 17", "Đô thị bền vững, thể chế minh bạch, đối tác toàn cầu", "72 câu"],
        ["Tài nguyên (Resources)", "SDG 6, SDG 7, SDG 12, SDG 14", "Nước sạch, năng lượng sạch, sản xuất trách nhiệm, biển", "89 câu"],
        ["Môi trường (Environments)", "SDG 13, SDG 15", "Hành động khí hậu, bảo tồn hệ sinh thái đất liền và rừng", "48 câu"]
    ]
    add_table_apa(
        doc,
        table_num="3",
        table_title="Phân nhóm 17 mục tiêu SDG thành 6 nhóm danh mục theo nhu cầu con người",
        headers=headers_t3,
        data=data_t3,
        note="Phân loại kế thừa từ nghiên cứu của Kang và Kim (2022).",
        col_widths=[3.5, 3.5, 6.5, 2.5],
        font_size=9.0
    )

    add_heading_2(doc, "3.4. Mô hình Phân tích Sắc thái Cảm xúc Đa lớp")
    add_p(
        doc,
        "Khác với nghiên cứu gốc của Kang và Kim (2022) sử dụng mô hình DistilBERT phân loại nhị phân (Positive/Negative), "
        "nghiên cứu này áp dụng mô hình PhoBERT tinh chỉnh (wonrax/phobert-base-vietnamese-sentiment) (Nguyen & Nguyen, 2020) "
        "với 3 nhãn phân loại: Tích cực (POS), Trung tính (NEU) và Tiêu cực (NEG). Việc bổ sung lớp Trung tính phản ánh chính xác bản chất của "
        "các câu văn mô tả kỹ thuật, thông số đo lường và quy chuẩn pháp lý trong văn bản báo cáo doanh nghiệp Việt Nam.\n"
        "Điểm phân cực cảm xúc (Polarity Score) liên tục trong đoạn [0, 1] được xác định theo quy tắc:"
    )
    add_formula_block(
        doc,
        r"\mathrm{Polarity} = \begin{cases} P(\mathrm{POS}) & \text{nếu nhãn là POSITIVE} \\ 0.5 & \text{nếu nhãn là NEUTRAL} \\ 1 - P(\mathrm{NEG}) & \text{nếu nhãn là NEGATIVE} \end{cases}",
        "4"
    )
    add_p(
        doc,
        "Tỷ số cảm xúc hàng năm của mỗi báo cáo được xác định thông qua tỷ lệ giữa số lượng câu tích cực và câu tiêu cực:"
    )
    add_formula_block(doc, r"\mathrm{Ratio} = \frac{N_{\mathrm{Positive}}}{N_{\mathrm{Negative}}}", "5")

    add_heading_2(doc, "3.5. Phương pháp Xác định Rủi ro Tẩy xanh Mở rộng")
    add_p(
        doc,
        "Nhằm kiểm chứng hiện tượng Tẩy xanh, nghiên cứu mở rộng khung phân tích của Kang và Kim bằng cách tích hợp chuỗi dữ liệu giá đóng cửa "
        "cổ phiếu điều chỉnh hàng năm từ sàn HOSE. Ban đầu, một tiêu chí phân loại phản ứng thị trường được thiết lập:\n"
        "- Phản ứng Tẩy xanh: Xảy ra khi giá cổ phiếu giảm sâu trên 10% (ΔGiá% < -10%) nhưng tỷ số cảm xúc trong báo cáo lại tăng mạnh trên 15% (ΔRatio% > +15%).\n"
        "- Phản ứng Trung thực: Xảy ra khi giá cổ phiếu giảm sâu (ΔGiá% < -10%) và tỷ số cảm xúc cũng đồng pha suy giảm (ΔRatio% < 0%).\n"
        "Đồng thời, nghiên cứu đề xuất Chỉ số Rủi ro Tẩy xanh liên tục (Greenwashing Score - GW):"
    )
    add_formula_block(
        doc,
        r"\mathrm{GW} = \frac{\max(0, -\Delta \mathrm{Gi\acute{a}}\%) \times \max(0, \Delta \mathrm{Ratio}\%)}{100}",
        "6"
    )
    add_p(
        doc,
        "Chỉ số GW chỉ nhận giá trị dương khi đồng thời xuất hiện sự sụt giảm định giá thị trường và sự gia tăng giọng điệu lạc quan trong báo cáo. "
        "Thang phân loại mức độ rủi ro được định nghĩa: GW = 0 (Không rủi ro / Trung thực); 0 < GW ≤ 2 (Rủi ro nhẹ); 2 < GW ≤ 8 (Rủi ro trung bình); GW > 8 (Rủi ro cao)."
    )

    doc.add_page_break()

    # ---------------------------------------------------------------------
    # 4. KẾT QUẢ NGHIÊN CỨU (RESULTS)
    # ---------------------------------------------------------------------
    add_heading_1(doc, "4. Kết quả nghiên cứu và Phân tích thực nghiệm")

    add_heading_2(doc, "4.1. Phân phối Điểm Tương đồng SDG Toàn cục")
    add_figure_apa(
        doc,
        img_filename="similarity_hist.png",
        figure_num="1",
        title="Phân phối tần suất điểm tương đồng SDG của toàn bộ 17.047 câu văn trên thang đo chuẩn hóa 0–100.",
        width_inches=5.8
    )
    add_p(
        doc,
        "Hình 1 minh họa phân phối mật độ xác suất của điểm tương đồng SDG trên toàn bộ 17.047 câu văn trích xuất từ 29 báo cáo sau khi chuẩn hóa min-max "
        "về thang đo 0–100. Đồ thị phân phối thể hiện dạng hình chuông gần như chuẩn tắc, với giá trị trung bình đạt 46.61 điểm "
        "và độ lệch chuẩn là 12.83 điểm. Phần lớn các câu văn trong báo cáo tập trung trong dải điểm từ 35 đến 60. "
        "Sự tồn tại của một phân phối rộng và biến thiên tự nhiên này hoàn toàn tương đồng với phát hiện của Kang và Kim (2022) trên mẫu các tập đoàn toàn cầu, "
        "chứng minh rằng mô hình Sentence-BERT tiếng Việt (vietnamese-sbert) đã nắm bắt tốt sự đa dạng ngữ nghĩa giữa các câu văn báo cáo và 17 SDGs, "
        "khắc phục hoàn toàn hiện tượng phân phối thoái hóa của các phương pháp khớp từ khóa truyền thống."
    )
    add_p(
        doc,
        "Bảng 4 cung cấp các tham số thống kê mô tả chi tiết của từng mục tiêu trong 17 SDGs. Kết quả phân tích định lượng cho thấy: "
        "Mục tiêu có điểm tương đồng trung bình cao nhất toàn mẫu là SDG 17 (Đối tác vì các mục tiêu - Mean = 53.74, Max = 100.00), theo sau là SDG 09 "
        "(Công nghiệp, đổi mới sáng tạo và hạ tầng - Mean = 50.11), SDG 12 (Sản xuất và tiêu dùng có trách nhiệm - Mean = 49.71) và SDG 15 (Tài nguyên đất liền - Mean = 49.01). "
        "Ngược lại, mục tiêu ghi nhận điểm số thấp nhất là SDG 05 (Bình đẳng giới - Mean = 38.70, Min = 0.00) và SDG 03 (Sức khỏe và phúc lợi - Mean = 42.42). "
        "Điều này phản ánh một thực tế khách quan rằng nội dung báo cáo phát triển bền vững tại Việt Nam tập trung dày đặc vào các cam kết hợp tác chuỗi giá trị và đổi mới công nghệ, "
        "trong khi các khía cạnh về bình đẳng giới và thu hẹp khoảng cách thu nhập nội bộ chưa nhận được sự quan tâm tương xứng (Bộ Kế hoạch và Đầu tư, 2023)."
    )

    # BẢNG 4: THỐNG KÊ MÔ TẢ 17 SDGS (APA 7th)
    headers_t4 = ["Mục tiêu SDG", "Số câu", "Trung bình", "Độ lệch chuẩn", "Tối thiểu", "Phân vị 25%", "Trung vị", "Phân vị 75%", "Tối đa"]
    data_t4 = [
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
    add_table_apa(
        doc,
        table_num="4",
        table_title="Thống kê mô tả điểm số tương đồng của 17 mục tiêu SDG trên toàn bộ 17.047 câu văn",
        headers=headers_t4,
        data=data_t4,
        note="Điểm số được chuẩn hóa Min-Max toàn cục trên thang đo từ 0 đến 100.",
        col_widths=[3.8, 1.6, 1.6, 1.6, 1.4, 1.6, 1.6, 1.6, 1.6],
        font_size=8.5
    )

    add_heading_2(doc, "4.2. Phân tích Cấu trúc 6 Nhóm Danh mục SDG qua Ma trận Nhiệt")
    add_figure_apa(
        doc,
        img_filename="heatmap_6cat.png",
        figure_num="2",
        title="Biểu đồ nhiệt thể hiện điểm số tương đồng 6 nhóm danh mục SDG theo công ty và năm.",
        width_inches=5.6
    )
    add_p(
        doc,
        "Biểu đồ nhiệt tại Hình 2 và số liệu chi tiết trong Bảng 5 thể hiện sự phân hóa cấu trúc chủ đề giữa 4 doanh nghiệp theo 6 nhóm danh mục. "
        "Một quy luật tổng quát mang tính nhất quán cao được phát hiện trên toàn bộ 29 báo cáo: Nhóm Kinh tế luôn dẫn đầu, tiếp theo là Xã hội, Tài nguyên, "
        "Đời sống, Môi trường và cuối cùng là Công bằng xã hội.\n"
        "Cụ thể, nhóm Kinh tế (SDG 8, 9) đạt điểm số cao vượt trội tại hầu hết các công ty–năm, dao động từ 44.61 (Vinamilk năm 2019) đến mức đỉnh 54.84 (PNJ năm 2022). "
        "Điều này phản ánh định hướng truyền thông cốt lõi của các doanh nghiệp niêm yết tại thị trường mới nổi: báo cáo phát triển bền vững vẫn được tận dụng tối đa để khẳng định "
        "hiệu quả sản xuất kinh doanh, mở rộng hạ tầng và đóng góp ngân sách nhà nước – vốn là các tiêu chí dễ lượng hóa và tạo được niềm tin tức thời nơi cổ đông."
    )
    add_p(
        doc,
        "Ngược lại, nhóm Công bằng (SDG 4, 5, 10) liên tục ghi nhận điểm số thấp nhất trên mọi báo cáo, chỉ dao động từ 39.92 (Petrolimex năm 2024) đến 47.40 (PNJ năm 2025). "
        "Ngay cả tại PNJ – doanh nghiệp kim hoàn sở hữu lực lượng lao động nữ trên 60% và có tiểu ban ESG chuyên trách – điểm nhóm Công bằng (44.29 – 47.40) vẫn thấp hơn đáng kể "
        "so với nhóm Kinh tế (50.96 – 54.84) và Xã hội (49.46 – 52.99). Phát hiện thực nghiệm này hoàn toàn đồng điệu với nhận định trong Báo cáo Rà soát Quốc gia Tự nguyện "
        "(VNR 2023) của Chính phủ Việt Nam tại Liên Hợp Quốc (Bộ Kế hoạch và Đầu tư, 2023), chỉ ra rằng việc xóa bỏ định kiến giới, thu hẹp khoảng cách tiền lương "
        "và hỗ trợ các nhóm yếu thế trong chuỗi cung ứng vẫn là thách thức lớn nhất và chậm tiến độ nhất trong 17 SDGs tại Việt Nam."
    )

    # BẢNG 5: ĐIỂM 6 NHÓM SDG (APA 7th)
    headers_t5 = ["Công ty", "Năm", "Đời sống", "Kinh tế", "Công bằng", "Xã hội", "Tài nguyên", "Môi trường", "Trung bình"]
    data_t5 = [
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
    add_table_apa(
        doc,
        table_num="5",
        table_title="Điểm số tương đồng trung bình 6 nhóm danh mục SDG theo công ty và năm",
        headers=headers_t5,
        data=data_t5,
        note="Dấu (*) biểu thị các báo cáo có số lượng câu dưới 80 câu do bản tóm tắt trực tuyến.",
        col_widths=[1.5, 1.2, 1.8, 1.8, 1.8, 1.8, 1.8, 2.0, 2.3],
        font_size=8.5
    )

    add_heading_2(doc, "4.3. Diễn biến Xu hướng theo Thời gian của Từng Doanh nghiệp")
    add_figure_apa(
        doc,
        img_filename="trends_6categories.png",
        figure_num="3",
        title="Đường xu hướng biến thiên điểm số 6 nhóm danh mục SDG giai đoạn 2018–2025 cho từng doanh nghiệp.",
        width_inches=5.8
    )
    add_p(
        doc,
        "Hình 3 thể hiện quỹ đạo vận động của 6 nhóm danh mục theo thời gian cho từng doanh nghiệp. Điểm nhấn quan trọng nhất là bước nhảy vọt đồng loạt "
        "của tất cả các doanh nghiệp trong năm 2025. Cụ thể, Vinamilk có bước bứt phá mạnh nhất khi toàn bộ 6 nhóm danh mục đều vượt ngưỡng 46 điểm, "
        "trong đó Môi trường tăng vọt từ 46.30 (năm 2024) lên 51.95 (năm 2025), và Đời sống tăng từ 45.35 lên 50.06. Sự chuyển biến này là hệ quả trực tiếp từ việc Vinamilk "
        "công bố chiến lược Net Zero 2050, hoàn thành chứng nhận trung hòa carbon PAS 2060 cho các nhà máy và trang trại tại Nghệ An, Bến Tre, "
        "đồng thời phát hành bản báo cáo phát triển bền vững dày 152 trang với 954 câu văn chuyên sâu."
    )
    add_p(
        doc,
        "Đối với The PAN Group, điểm số suy giảm nhẹ trong năm 2023 (điểm trung bình 44.24 so với 47.09 năm 2022) do công ty tinh giản dung lượng báo cáo "
        "(chỉ 411 câu so với 812 câu năm 2022), sau đó phục hồi mạnh mẽ vào năm 2025 (đạt 49.77 điểm). Đối với PNJ, công ty thể hiện phong độ ổn định nhất mẫu "
        "với điểm trung bình luôn dao động từ 48 đến 51 điểm. Riêng Petrolimex, đường xu hướng bị biến động mạnh cục bộ tại hai năm 2021 và 2024 "
        "do hiện tượng mẫu mỏng (bản tóm tắt trực tuyến)."
    )

    add_heading_2(doc, "4.4. Phân tích Sắc thái Cảm xúc và Cơ cấu Ngôn từ")
    add_figure_apa(
        doc,
        img_filename="sentiment_hist.png",
        figure_num="4",
        title="Phân phối điểm phân cực sắc thái cảm xúc (Polarity Score) từ mô hình PhoBERT đa lớp.",
        width_inches=5.2
    )
    add_p(
        doc,
        "Hình 4 minh họa phân phối điểm phân cực cảm xúc (Polarity Score) của 17.047 câu văn. Không giống như phân phối chữ U trong nghiên cứu "
        "của Kang và Kim (2022) vốn chỉ dùng 2 nhãn nhị phân, đồ thị phân phối của nghiên cứu này mang tính chất lưỡng đỉnh rõ rệt với hai đỉnh tập trung: "
        "Đỉnh thứ nhất tại mốc 0.50 tương ứng với 3.126 câu trung tính (chiếm 18.34%), và đỉnh khổng lồ thứ hai nằm sát mốc 0.95–1.00 đại diện cho 11.891 câu "
        "tích cực (chiếm 69.75%). Số lượng câu thực sự tiêu cực (Polarity < 0.2) chỉ có 2.030 câu (chiếm 11.91%). Tỷ lệ áp đảo gần 70% câu tích cực là "
        "bằng chứng thực nghiệm khẳng định rằng báo cáo phát triển bền vững tại Việt Nam vẫn mang đậm tính chất quan hệ công chúng và quảng bá hình ảnh."
    )
    add_figure_apa(
        doc,
        img_filename="sentiment_by_company.png",
        figure_num="5",
        title="Cơ cấu tỷ lệ câu Tích cực, Trung tính, Tiêu cực và Tỷ số Pos/Neg Ratio theo từng doanh nghiệp.",
        width_inches=5.8
    )
    add_p(
        doc,
        "Bảng 6 và Hình 5 trình bày cơ cấu cảm xúc chi tiết theo từng báo cáo. Vinamilk duy trì tỷ số cảm xúc (Pos/Neg Ratio) cao nhất toàn mẫu, "
        "đạt đỉnh 11.97 vào năm 2020 (455 câu tích cực so với vỏn vẹn 38 câu tiêu cực) và 10.86 vào năm 2024. The PAN Group có tỷ số cảm xúc tương đối ổn định "
        "trong khoảng 5.40 đến 10.22. Petrolimex thể hiện phong cách báo cáo thận trọng hơn của một tập đoàn công nghiệp nhà nước, duy trì tỷ số Pos/Neg "
        "thấp nhất mẫu (3.00 đến 5.22, ngoại trừ năm mẫu mỏng 2024). Trường hợp dị biệt xuất hiện tại PNJ năm 2022: Tỷ số Pos/Neg sụt giảm xuống chỉ còn 1.33 "
        "(182 câu tích cực / 137 câu tiêu cực). Nguyên nhân bắt nguồn từ lỗi nhận diện OCR trên bản PDF scan ảnh: các khối văn bản bị vỡ chữ khiến mô hình "
        "PhoBERT phân loại nhầm hàng loạt câu thành nhãn Tiêu cực. Đây là một bài học thực tiễn về tầm quan trọng của việc kiểm soát chất lượng số hóa tài liệu trước khi phân tích NLP."
    )

    # BẢNG 6: SỐ LƯỢNG SENTIMENT (APA 7th)
    headers_t6 = ["Công ty", "Năm", "Tiêu cực (NEG)", "Trung tính (NEU)", "Tích cực (POS)", "Tổng số câu", "Tỷ số Pos/Neg Ratio"]
    data_t6 = [
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
    add_table_apa(
        doc,
        table_num="6",
        table_title="Cơ cấu số lượng câu theo nhãn cảm xúc và tỷ số Pos/Neg Ratio qua các năm",
        headers=headers_t6,
        data=data_t6,
        note="Dấu (*) biểu thị mẫu mỏng dưới 80 câu; dấu (**) biểu thị báo cáo scan ảnh có nhiễu OCR.",
        col_widths=[1.5, 1.2, 2.6, 2.6, 2.6, 2.2, 3.3],
        font_size=8.5
    )

    add_figure_apa(
        doc,
        img_filename="sentiment_ratio.png",
        figure_num="6",
        title="Đường xu hướng biến thiên tỷ số Pos/Neg Ratio của 4 công ty giai đoạn 2018–2025.",
        width_inches=5.5
    )

    add_heading_2(doc, "4.5. Phân tích Hiện tượng Tẩy xanh")
    add_figure_apa(
        doc,
        img_filename="greenwashing_score.png",
        figure_num="7",
        title="Biểu đồ nhiệt điểm rủi ro Tẩy xanh (Greenwashing Score) và xếp hạng rủi ro giữa các công ty.",
        width_inches=5.6
    )
    add_p(
        doc,
        "Hình 7 và Bảng 7 thể hiện kết quả tính toán Chỉ số Rủi ro Tẩy xanh (Greenwashing Score - GW) đối sánh giữa biến động giá cổ phiếu trên HOSE và "
        "sự thay đổi tỷ số cảm xúc trong báo cáo. Kết quả phân tích làm nổi bật sự phân hóa sâu sắc về tính trung thực thông tin giữa các doanh nghiệp:\n"
        "1. Petrolimex ghi nhận rủi ro Tẩy xanh cao nhất toàn mẫu với GW trung bình đạt 0.68 và có 2/3 năm thị trường sụt giảm xuất hiện biểu hiện tô hồng. "
        "Điểm số GW đạt cực đại vào năm 2022 với mức GW = 2.39 (mức Trung bình): trong năm này, giá cổ phiếu PLX lao dốc mạnh -31.86%, tuy nhiên báo cáo bền vững "
        "của PLX lại gia tăng tỷ số lạc quan thêm +7.51% (Ratio tăng từ 3.00 lên 3.23). Tình trạng tương tự tiếp diễn vào năm 2023 khi giá PLX giảm thêm -4.32% nhưng giọng văn "
        "báo cáo lại tăng vọt +38.68% (Ratio tăng lên 4.47, GW = 1.67). Điều này phản ánh áp lực truyền thông to lớn trong việc giải trình tiến độ chuyển dịch năng lượng xanh.\n"
        "2. Ngược lại, The PAN Group thể hiện tính minh bạch và trung thực cao. Năm 2022, khi thị trường nông nghiệp gặp khó khăn khiến giá cổ phiếu PAN sụt giảm tới -40.00%, "
        "tỷ số cảm xúc trong báo cáo của PAN cũng đồng pha suy giảm -6.50% (Ratio giảm từ 6.90 xuống 6.45). Doanh nghiệp được phân loại 'Trung thực' "
        "và duy trì chỉ số GW trung bình bằng 0.00 suốt toàn bộ chuỗi thời gian.\n"
        "3. Vinamilk thể hiện tính trung thực cao trong năm 2021 khi thị trường điều chỉnh (-15.35%) và tỷ số cảm xúc giảm mạnh tương ứng (-40.23%). "
        "Tuy nhiên, vào các năm 2022 và 2024, Vinamilk xuất hiện rủi ro tẩy xanh nhẹ (GW lần lượt là 0.43 và 0.56) khi giá cổ phiếu đi ngang/giảm nhẹ nhưng báo cáo "
        "lại gia tăng mạnh các câu từ ca ngợi thành tích trung hòa carbon."
    )

    # BẢNG 7: BẢNG GREENWASHING SCORE THEO NĂM (APA 7th)
    headers_t7 = ["Công ty", "Năm", "Số câu", "Δ Giá CP", "Tỷ số Ratio", "Δ Ratio", "Phản ứng thị trường", "Điểm GW", "Mức độ rủi ro"]
    data_t7 = [
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
    add_table_apa(
        doc,
        table_num="7",
        table_title="Kết quả tính toán biến động giá cổ phiếu, tỷ số cảm xúc và chỉ số Greenwashing Score (GW)",
        headers=headers_t7,
        data=data_t7,
        note="Điểm GW tính theo công thức [max(0, -ΔGiá%) × max(0, ΔRatio%)] / 100.",
        col_widths=[1.5, 1.2, 1.5, 2.2, 1.8, 2.0, 2.6, 1.5, 2.8],
        font_size=8.5
    )

    # BẢNG 8: XẾP HẠNG TẨY XANH (APA 7th)
    headers_t8 = ["Hạng", "Doanh nghiệp", "Số năm tính toán (n)", "Số năm giá giảm", "Số năm GW > 0", "GW trung bình", "GW cực đại", "Năm đạt GW Max", "Số năm Trung thực"]
    data_t8 = [
        ["1", "Petrolimex (PLX)", "6 năm", "3 năm", "2 năm", "0.68", "2.39", "Năm 2022", "0 năm"],
        ["2", "Vinamilk (VNM)", "6 năm", "4 năm", "2 năm", "0.17", "0.56", "Năm 2024", "1 năm (2021)"],
        ["3", "The PAN Group (PAN)", "6 năm", "1 năm", "0 năm", "0.00", "0.00", "—", "1 năm (2022)"],
        ["4", "PNJ (PNJ)", "6 năm", "0 năm", "0 năm", "0.00", "0.00", "—", "0 năm"]
    ]
    add_table_apa(
        doc,
        table_num="8",
        table_title="Bảng xếp hạng mức độ rủi ro Tẩy xanh tổng hợp giữa 4 doanh nghiệp nghiên cứu",
        headers=headers_t8,
        data=data_t8,
        note="Xếp hạng dựa trên giá trị GW trung bình qua chuỗi thời gian 2020–2025.",
        col_widths=[1.2, 3.5, 2.2, 2.0, 2.0, 2.0, 2.0, 2.0, 2.1],
        font_size=8.5
    )

    add_heading_2(doc, "4.6. Kiểm định Tương quan giữa Giá Cổ phiếu và Sắc thái Báo cáo")
    add_figure_apa(
        doc,
        img_filename="stock_vs_sentiment_greenwashing.png",
        figure_num="8",
        title="Đối sánh chuỗi giá cổ phiếu tháng và tỷ số cảm xúc năm (chuẩn hóa Index 100) kèm nhãn Tẩy xanh.",
        width_inches=6.0
    )
    add_p(
        doc,
        "Hình 8 thể hiện mối quan hệ động giữa đường giá cổ phiếu giao dịch theo tháng trên HOSE (đường màu xanh, chuẩn hóa gốc 100) và tỷ số cảm xúc "
        "theo năm dạng ziczac. Các khoảng thời gian xuất hiện hiện tượng tẩy xanh hoặc trung thực được tô bóng tương ứng. "
        "Bảng 9 trình bày kết quả kiểm định tương quan tuyến tính Pearson (r) và tương quan phi tham số Spearman (ρ) giữa phần trăm thay đổi giá cổ phiếu hàng năm "
        "và phần trăm thay đổi của tỷ số cảm xúc Pos/Neg Ratio."
    )

    # BẢNG 9: TƯƠNG QUAN (APA 7th)
    headers_t9 = ["Doanh nghiệp", "Cỡ mẫu (n)", "Hệ số Pearson r", "p-value (Pearson)", "Hệ số Spearman ρ", "p-value (Spearman)", "Kết luận thống kê"]
    data_t9 = [
        ["The PAN Group (PAN)", "6", "+0.173", "0.744", "-0.257", "0.623", "Không tương quan (p > 0.10)"],
        ["Petrolimex (PLX)", "6", "+0.352", "0.494", "+0.314", "0.544", "Không tương quan (p > 0.10)"],
        ["PNJ (PNJ)", "6", "-0.470", "0.347", "-0.429", "0.397", "Không tương quan (p > 0.10)"],
        ["Vinamilk (VNM)", "6", "+0.462", "0.356", "+0.429", "0.397", "Không tương quan (p > 0.10)"],
        ["Toàn bộ mẫu gộp (Pool)", "24", "+0.030", "0.890", "-0.047", "0.828", "Hoàn toàn độc lập (p = 0.89)"]
    ]
    add_table_apa(
        doc,
        table_num="9",
        table_title="Ma trận hệ số tương quan giữa biến động giá cổ phiếu và biến động sắc thái cảm xúc báo cáo",
        headers=headers_t9,
        data=data_t9,
        note="Kiểm định hai phía (two-tailed test), mức ý nghĩa thống kê α = 0.05.",
        col_widths=[3.5, 1.8, 2.2, 2.5, 2.2, 2.5, 3.8],
        font_size=8.5
    )

    add_p(
        doc,
        "Kết quả kiểm định trên mẫu gộp (n = 24 quan sát) cho thấy hệ số tương quan Pearson xấp xỉ bằng không (r = +0.030, p = 0.890) và Spearman ρ = -0.047 (p = 0.828). "
        "Ở cấp độ từng doanh nghiệp riêng lẻ, dù Vinamilk có tương quan dương vừa phải (r = +0.462) và PNJ có tương quan âm (r = -0.470), tất cả các giá trị p-value đều vượt xa "
        "ngưỡng ý nghĩa thống kê chuẩn (p > 0.30). Kết quả thực nghiệm này khẳng định một kết luận kinh tế quan trọng: Tại thị trường chứng khoán Việt Nam, "
        "giọng điệu và mức độ lạc quan của báo cáo phát triển bền vững hoàn toàn tách rời khỏi biến động giá trị thị trường của cổ phiếu. "
        "Nhà đầu tư chứng khoán Việt Nam chưa định giá trực tiếp các thông điệp văn bản bền vững vào thị giá, tạo điều kiện cho các doanh nghiệp tiếp tục duy trì "
        "chiến lược tẩy xanh truyền thông mà không lo ngại rủi ro bị trừng phạt tài chính tức thì (Seele & Gatti, 2017; Tran & Beddewela, 2020)."
    )

    doc.add_page_break()

    # ---------------------------------------------------------------------
    # 5. THẢO LUẬN (DISCUSSION)
    # ---------------------------------------------------------------------
    add_heading_1(doc, "5. Thảo luận")

    add_heading_2(doc, "5.1. Thực trạng Báo cáo Bền vững tại Việt Nam qua Lăng kính NLP")
    add_p(
        doc,
        "Phân tích thực nghiệm trên 17.047 câu văn từ 29 báo cáo đã phác họa bức tranh chân thực về hiện trạng công bố thông tin ESG tại Việt Nam. "
        "Mặc dù khuôn khổ pháp lý (Thông tư 96/2020/TT-BTC) đã tạo ra một lực đẩy mạnh mẽ, văn hóa báo cáo của doanh nghiệp Việt "
        "vẫn mang đặc trưng chọn lọc có lợi (Heras-Saizarbitoria et al., 2022). Sự thống trị tuyệt đối của nhóm Kinh tế và sự lép vế có hệ thống của nhóm Công bằng "
        "chỉ ra rằng doanh nghiệp đang cố gắng khoác chiếc áo bền vững lên các mục tiêu kinh doanh truyền thống (Arvidsson & Dumay, 2022). "
        "Đặc biệt, điểm tương đồng của SDG 5 (Bình đẳng giới) chỉ đạt bình quân 38.70/100, phản ánh khoảng trống trong chính sách quản trị nhân sự, "
        "tiền lương bình đẳng và trao quyền cho nữ giới – điều hoàn toàn khớp với xếp hạng trung bình (72/146) của Việt Nam theo Báo cáo Khoảng cách Giới Toàn cầu 2024 của Diễn đàn Kinh tế Thế giới (WEF)."
    )

    add_heading_2(doc, "5.2. Giải phẫu Hành trình Bền vững của 4 Doanh nghiệp Nghiên cứu")
    add_p(
        doc,
        "Bốn doanh nghiệp trong mẫu nghiên cứu đại diện cho các mô hình ứng xử và chiến lược bền vững đặc trưng trong nền kinh tế Việt Nam:\n"
        "Vinamilk đại diện cho mô hình tiên phong chuẩn mực quốc tế: Là doanh nghiệp có lịch sử báo cáo GRI liên tục từ năm 2012, Vinamilk thể hiện năng lực quản trị dữ liệu vượt trội. "
        "Sự bứt phá điểm số vào năm 2025 gắn liền với các chứng nhận trung hòa carbon PAS 2060 độc lập và công bố CDP. Tuy nhiên, việc duy trì tỷ số cảm xúc quá cao (11.97 năm 2020 và 10.86 năm 2024) "
        "cũng đặt ra dấu hỏi về sự thận trọng trước nguy cơ nói trước làm sau khi các mục tiêu Net Zero 2050 còn ở chặng đường dài.\n"
        "The PAN Group đại diện cho mô hình nông nghiệp bền vững thực chất: Với đặc thù ngành nông nghiệp và xuất khẩu thủy sản sang các thị trường khắt khe, "
        "PAN bắt buộc phải tuân thủ các chuẩn mực nuôi trồng bền vững quốc tế (ASC, BAP). Điểm số của PAN phản ánh tính thực chất cao, thể hiện rõ nhất qua tính trung thực "
        "trong năm 2022 khi doanh nghiệp dũng cảm giảm tỷ lệ ngôn từ ca ngợi để phản ánh đúng áp lực chi phí đầu vào và lạm phát.\n"
        "Petrolimex đại diện cho nghịch lý của doanh nghiệp năng lượng hóa thạch chuyển dịch: Là tập đoàn xăng dầu nhà nước chiếm 50% thị phần, PLX đối mặt với áp lực giảm phát thải "
        "khổng lồ. Các nỗ lực thương mại hóa Diesel Euro 5 và lắp đặt trạm sạc xe điện là rất đáng ghi nhận, nhưng sự phân kỳ sâu sắc giữa giá cổ phiếu sụt giảm và giọng văn tô hồng "
        "trong năm 2022–2023 (GW = 2.39 và 1.67) cho thấy tâm lý né tránh rủi ro truyền thông điển hình của khối doanh nghiệp nhà nước.\n"
        "PNJ đại diện cho mô hình tiên phong về bình đẳng giới và quản trị: PNJ là hình mẫu hiếm hoi tại Việt Nam lập tiểu ban ESG thuộc Hội đồng Quản trị từ sớm và tập trung thực chất vào bình đẳng giới. "
        "Tuy nhiên, sự cố báo cáo scan năm 2022 là lời cảnh tỉnh: Nếu doanh nghiệp không chuẩn hóa định dạng số hóa tài liệu, các hệ thống AI và tổ chức xếp hạng tín nhiệm "
        "ESG quốc tế sẽ tự động đánh giá sai lệch năng lực bền vững của doanh nghiệp dựa trên các đoạn văn bản lỗi."
    )

    add_heading_2(doc, "5.3. So sánh Đối chiếu với Phát hiện của Kang và Kim (2022)")
    add_p(
        doc,
        "Khi đặt cạnh kết quả của Kang và Kim (2022) trên 6 tập đoàn đa quốc gia hàng đầu thế giới (BASF, IKEA, Microsoft, Nestlé, Toyota, Walmart), "
        "nghiên cứu này chỉ ra các điểm tương đồng và khác biệt mang ý nghĩa học thuật:\n"
        "Về điểm tương đồng: Cả hai nghiên cứu đều xác nhận tính ưu việt của Sentence-BERT so với đếm từ khóa, đều phát hiện xu hướng gia tăng dung lượng báo cáo "
        "sau năm 2020, và đều chứng minh sự thống trị của các câu văn mang sắc thái tích cực (chiếm trên 65%). Trường hợp vụ bê bối chì mì Maggi của Nestlé năm 2015 "
        "khiến tỷ số Pos/Neg vọt lên 6.0 trong nghiên cứu gốc hoàn toàn tương đồng với hành vi của Petrolimex năm 2022 khi giá dầu biến động khiến điểm GW đạt đỉnh 2.39.\n"
        "Về điểm khác biệt: Điểm tương đồng trung bình của doanh nghiệp Việt Nam (46.61) thấp hơn khoảng 3–5 điểm so với các tập đoàn đa quốc gia quốc tế, cho thấy ngôn ngữ báo cáo "
        "tiếng Việt vẫn dùng nhiều từ ngữ chung chung, chưa ánh xạ chuẩn xác vào các thuật ngữ kỹ thuật của 169 chỉ tiêu SDG. Hơn nữa, sự hiện diện của 18.34% câu trung tính "
        "được mô hình PhoBERT bóc tách là đóng góp mới mà mô hình nhị phân của nghiên cứu gốc không thể thực hiện."
    )

    add_heading_2(doc, "5.4. Hàm ý Quản trị và Khuyến nghị Chính sách")
    add_p(
        doc,
        "Dựa trên các bằng chứng thực nghiệm, nghiên cứu đề xuất các khuyến nghị cụ thể:\n"
        "Thứ nhất, đối với Cơ quan Quản lý (Ủy ban Chứng khoán Nhà nước và các Sở Giao dịch Chứng khoán): Cần sớm ban hành khung hướng dẫn áp dụng chuẩn mực ISSB (IFRS S1 và S2), "
        "đồng thời quy định bắt buộc kiểm toán độc lập đối với dữ liệu phát thải khí nhà kính Scope 1 và 2. Cần ứng dụng các công cụ NLP tự động để thiết lập hệ thống giám sát cảnh báo "
        "sớm rủi ro Tẩy xanh trên thị trường chứng khoán.\n"
        "Thứ hai, đối với doanh nghiệp niêm yết: Cần thay đổi tư duy từ quản trị ấn tượng sang trách nhiệm giải trình thực chất. Doanh nghiệp cần dũng cảm công bố cả các thách thức, "
        "rủi ro và mục tiêu chưa hoàn thành; đồng thời chuẩn hóa định dạng văn bản số điện tử để tạo thuận lợi cho việc phân tích dữ liệu lớn.\n"
        "Thứ ba, đối với nhà đầu tư và các định chế tài chính: Không nên dựa vào tần suất xuất hiện của các từ khóa bền vững mà cần đối chiếu tính nhất quán giữa cam kết và kết quả tài chính, "
        "sử dụng chỉ số Greenwashing Score để sàng lọc danh mục đầu tư bền vững."
    )

    add_heading_2(doc, "5.5. Các Giới hạn Khoa học của Nghiên cứu")
    add_p(
        doc,
        "Nghiên cứu tồn tại một số giới hạn cần được mở rộng trong tương lai: (1) Cỡ mẫu giới hạn ở 4 doanh nghiệp quy mô lớn thuộc VN100 do sự khan hiếm "
        "của các báo cáo phát triển bền vững độc lập kéo dài liên tục 8 năm tại Việt Nam; (2) Chuỗi thời gian hàng năm (n = 6 quan sát có đủ dữ liệu biến động giá) chưa đủ dài để các kiểm định "
        "tương quan đạt mức ý nghĩa thống kê chuẩn p < 0.05; (3) Nghiên cứu chưa đi sâu trích xuất các bảng số liệu định lượng (như tấn CO2e, mét khối nước tiêu thụ) vốn nằm "
        "trong các biểu đồ đồ họa phức tạp của tệp PDF."
    )

    # ---------------------------------------------------------------------
    # 6. KẾT LUẬN (CONCLUSION)
    # ---------------------------------------------------------------------
    add_heading_1(doc, "6. Kết luận")
    add_p(
        doc,
        "Nghiên cứu này là công trình thực nghiệm đầu tiên tại Việt Nam tái hiện và nâng cấp toàn diện khung phân tích NLP của Kang và Kim (2022) "
        "trên dữ liệu báo cáo phát triển bền vững tiếng Việt. Bằng việc kết hợp các mô hình học sâu tiên tiến (vietnamese-sbert và PhoBERT), "
        "nghiên cứu đã chứng minh tính khả thi và sức mạnh của trí tuệ nhân tạo trong việc giải mã, định lượng và kiểm chứng tính chân thực của thông tin phi tài chính. "
        "Những phát hiện về sự thống trị của mục tiêu Kinh tế, sự tụt hậu của mục tiêu Bình đẳng giới, sự thiên lệch lạc quan của sắc thái ngôn từ và hiện tượng tẩy xanh tại "
        "Petrolimex cung cấp những luận cứ khoa học cho các nhà hoạch định chính sách, doanh nghiệp và nhà đầu tư trong lộ trình hiện thực hóa cam kết Net Zero 2050 của quốc gia. "
        "Hướng nghiên cứu tiếp theo sẽ tập trung mở rộng quy mô mẫu sang toàn bộ các công ty đại chúng thuộc rổ VNSI, kết hợp mô hình thị giác máy tính trích xuất bảng biểu "
        "và ứng dụng các mô hình ngôn ngữ lớn để kiểm chứng độc lập các tuyên bố bền vững."
    )

    doc.add_page_break()

    # ---------------------------------------------------------------------
    # TÀI LIỆU THAM KHẢO (REFERENCES - APA 7th ALPHABETICAL)
    # ---------------------------------------------------------------------
    add_heading_1(doc, "Tài liệu tham khảo")

    # Danh mục tài liệu sắp xếp chuẩn A-Z theo họ tác giả, định dạng APA 7th
    apa_references = [
        "Arvidsson, S., & Dumay, J. (2022). Corporate ESG reporting quantity, quality and performance: Where to now for environmental policy and practice? Business Strategy and the Environment, 31(3), 1091–1110. https://doi.org/10.1002/bse.2937",
        "Bingler, J. A., Kraus, M., Leippold, M., & Webersinke, N. (2022). Cheap talk and cherry-picking: What companies say site-wide about climate change. Finance Research Letters, 47, Article 102760. https://doi.org/10.1016/j.frl.2022.102760",
        "Blei, D. M., Ng, A. Y., & Jordan, M. I. (2003). Latent Dirichlet allocation. Journal of Machine Learning Research, 3, 993–1022.",
        "Bộ Kế hoạch và Đầu tư. (2023). Báo cáo rà soát quốc gia tự nguyện lần thứ 2 việc thực hiện các mục tiêu phát triển bền vững của Việt Nam (VNR 2023). Nhà xuất bản Thống kê.",
        "Bộ Tài chính. (2020). Thông tư số 96/2020/TT-BTC ngày 16/11/2020 hướng dẫn công bố thông tin trên thị trường chứng khoán.",
        "Cer, D., Yang, Y., Kong, S. Y., Hua, N., Limtiaco, N., St. John, R., Constant, N., Guajardo-Céspedes, M., Yuan, S., Tar, C., Strope, B., & Kurzweil, R. (2018). Universal sentence encoder for English. In Proceedings of the 2018 Conference on Empirical Methods in Natural Language Processing: System Demonstrations (pp. 169–174). Association for Computational Linguistics. https://doi.org/10.18653/v1/D18-2029",
        "Chính phủ Việt Nam. (2020). Nghị định số 155/2020/NĐ-CP ngày 31/12/2020 quy định chi tiết thi hành một số điều của Luật Chứng khoán.",
        "Devlin, J., Chang, M. W., Lee, K., & Toutanova, K. (2019). BERT: Pre-training of deep bidirectional transformers for language understanding. In Proceedings of NAACL-HLT 2019 (pp. 4171–4186). Association for Computational Linguistics.",
        "Du, S., & Yu, K. (2017). The business case for sustainability reporting: Evidence from stock market reactions. Journal of Public Policy & Marketing, 36(2), 313–330. https://doi.org/10.1509/jppm.16.112",
        "El-Haj, M., Rayson, P., Walker, M., Young, S., & Simaki, V. (2020). In search of 'sunlight'? Measuring transparency, disclosure quality and tone in corporate annual reports. Accounting and Business Research, 50(5), 450–476. https://doi.org/10.1080/00014788.2020.1771960",
        "Friede, G., Busch, T., & Bassen, A. (2015). ESG and financial performance: Aggregated evidence from more than 2000 empirical studies. Journal of Sustainable Finance & Investment, 5(4), 210–233. https://doi.org/10.1080/20430795.2015.1118917",
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
        "Mercereau, B., & Melin, L. (2020). ESG analysis: An NLP approach to assessing corporate sustainability disclosures. The Journal of Investing, 29(7), 50–63. https://doi.org/10.3905/joi.2020.1.157",
        "Mikolov, T., Chen, K., Corrado, G., & Dean, J. (2013). Efficient estimation of word representations in vector space. In Proceedings of ICLR 2013. arXiv:1301.3781.",
        "Muñoz-Torres, M. J., Fernández-Izquierdo, M. Á., Rivera-Lirio, J. M., & Escrig-Olmedo, E. (2019). Can modern sustainability reports track the SDGs? An assessment framework. Sustainability, 11(5), Article 1421. https://doi.org/10.3390/su11051421",
        "Nguyen, D. Q., & Nguyen, A. T. (2020). PhoBERT: Pre-trained language models for Vietnamese. In Findings of EMNLP 2020 (pp. 1037–1042). Association for Computational Linguistics. https://doi.org/10.18653/v1/2020.findings-emnlp.92",
        "Pennington, J., Socher, R., & Manning, C. D. (2014). GloVe: Global vectors for word representation. In Proceedings of EMNLP 2014 (pp. 1532–1543). Association for Computational Linguistics. https://doi.org/10.3115/v1/D14-1162",
        "PwC Vietnam. (2022). Báo cáo khảo sát mức độ sẵn sàng thực hành ESG tại Việt Nam năm 2022: Từ tham vọng đến hành động. PwC Việt Nam.",
        "Reimers, N., & Gurevych, I. (2019). Sentence-BERT: Sentence embeddings using Siamese BERT-networks. In Proceedings of EMNLP-IJCNLP 2019 (pp. 3982–3992). Association for Computational Linguistics. https://doi.org/10.18653/v1/D19-1410",
        "Sanh, V., Debut, L., Chaumond, J., & Wolf, T. (2019). DistilBERT, a distilled version of BERT: smaller, faster, cheaper and lighter. arXiv preprint arXiv:1910.01108.",
        "Seele, P., & Gatti, L. (2017). Greenwashing revisited: In search of a typology and accusation-based definition. Business Strategy and the Environment, 26(2), 239–252. https://doi.org/10.1002/bse.1912",
        "Stacchezzini, R., Melloni, G., & Lai, A. (2016). Sustainability management and reporting: The role of impression management for corporate social responsibility disclosure. Journal of Cleaner Production, 136, 102–110. https://doi.org/10.1016/j.jclepro.2016.04.095",
        "Tran, M., & Beddewela, E. (2020). Evaluating the quality of CSR disclosure in Vietnam: An empirical examination of listed firms. Journal of Business Ethics, 166(3), 569–589. https://doi.org/10.1007/s10551-019-04135-2",
        "United Nations. (2015). Transforming our world: The 2030 agenda for sustainable development (Resolution A/RES/70/1). United Nations General Assembly.",
        "Ủy ban Chứng khoán Nhà nước. (2024). Sổ tay hướng dẫn thực hành và công bố thông tin môi trường, xã hội và quản trị (ESG) cho doanh nghiệp niêm yết. Nhà xuất bản Tài chính.",
        "Veenstra, E. M., & Ellemers, N. (2020). CSR does not equal investment in CSR: A linguistic analysis of corporate social responsibility reports. Journal of Business Ethics, 161(2), 347–363. https://doi.org/10.1007/s10551-018-3904-7",
        "Wang, W., Wei, F., Dong, L., Bao, H., Yang, N., & Zhou, M. (2020). MiniLM: Deep self-attention distillation for task-agnostic compression of pre-trained transformers. In Advances in Neural Information Processing Systems (Vol. 33, pp. 5776–5788). Curran Associates, Inc.",
        "World Commission on Environment and Development. (1987). Our common future. Oxford University Press."
    ]

    for ref in apa_references:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # Hanging indent 1.27 cm (0.5 inch) chuẩn APA 7th
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(-0.5)
        run = p.add_run(ref)
        run.font.name = "Times New Roman"
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_BLACK


def main():
    print("[*] Generating academic paper DOCX document (APA 7th standard format)...")
    doc = Document()
    setup_document_styles(doc)
    build_paper_content(doc)

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    saved_path = OUTPUT_FILE
    try:
        doc.save(str(OUTPUT_FILE))
    except PermissionError:
        saved_path = ROOT / "paper_sdg_vietnam_apa7.docx"
        doc.save(str(saved_path))
        print(f"[!] Canh bao: Tep {OUTPUT_FILE.name} dang duoc mo trong Word. Da luu sang: {saved_path.name}")
        
    file_size_kb = saved_path.stat().st_size / 1024
    print(f"[+] Paper successfully created at:")
    print(f"    -> {saved_path}")
    print(f"    -> File size: {file_size_kb:.1f} KB")


if __name__ == "__main__":
    main()
