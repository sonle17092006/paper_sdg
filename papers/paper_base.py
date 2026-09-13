"""Nền tảng định dạng văn bản APA 7th, chuyển đổi công thức Office Math (OMML),
và nạp dữ liệu chuẩn xác từ các file CSV trong data/results/.
"""

from __future__ import annotations

import re
import sys
import unicodedata
from pathlib import Path

import docx
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Cm, Inches, Pt, RGBColor
import latex2mathml.converter
import lxml.etree as ET
import pandas as pd

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Code\paper_sdg")
FIGURES_DIR = ROOT / "figures"
DATA_DIR = ROOT / "data" / "results"

COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_MUTED = RGBColor(80, 80, 80)

# Load Microsoft MathML to OMML stylesheet
XSL_PATH = Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL")
if XSL_PATH.exists():
    try:
        XSLT_DOC = ET.parse(str(XSL_PATH))
        TRANSFORM_MML = ET.XSLT(XSLT_DOC)
    except Exception:
        TRANSFORM_MML = None
else:
    TRANSFORM_MML = None


def setup_clean_styles(doc: Document):
    """Cài đặt lề trang A4 chuẩn APA 7th (Trái 3cm, Phải 2cm, Trên 2.5cm, Dưới 2.5cm).
    Không header thừa, đánh số trang căn giữa ở chân trang (footer).
    """
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)

        header = section.header
        header.is_linked_to_previous = False
        for p in header.paragraphs:
            p.text = ""

        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.text = ""
        frun = fp.add_run()
        frun.font.name = "Times New Roman"
        frun.font.size = Pt(10)
        frun.font.color.rgb = COLOR_BLACK

        fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls("w"))
        fp._p.append(fldSimple)

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
    except Exception:
        return None


def add_equation_clean(doc: Document, latex_eq: str, eq_num: str = ""):
    omml_element = latex_to_omml(latex_eq)
    if omml_element is not None:
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
    font_size: float = 8.5,
):
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

        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>'
        )
        tcPr.append(tcMar)

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
            tcMar = parse_xml(
                f'<w:tcMar {nsdecls("w")}><w:top w:w="50" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>'
            )
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
        n_label.font.size = Pt(9.0)
        n_label.italic = True

        n_text = note_p.add_run(note)
        n_text.font.name = "Times New Roman"
        n_text.font.size = Pt(9.0)
        n_text.italic = False
    else:
        sp = doc.add_paragraph()
        sp.paragraph_format.space_after = Pt(4)

    return table


# =========================================================================
# LOAD TOÀN BỘ DỮ LIỆU TỪ FILE CSV
# =========================================================================


def load_all_tables():
    # Table 1: Sample Overview (7 companies)
    t1_headers_vi = [
        "Mã CK",
        "Tên doanh nghiệp",
        "Ngành hoạt động cốt lõi",
        "Sàn",
        "Giai đoạn",
        "Chuẩn mực công bố áp dụng",
    ]
    t1_data_vi = [
        [
            "BVH",
            "Tập đoàn Bảo Việt (Bao Viet Holdings)",
            "Tài chính, Bảo hiểm nhân thọ/phi nhân thọ, Quản lý quỹ",
            "HOSE",
            "2020–2025",
            "Báo cáo tích hợp quốc tế (<IIRC>), GRI Standards",
        ],
        [
            "PAN",
            "CTCP Tập đoàn PAN (The PAN Group)",
            "Nông nghiệp công nghệ cao, Thủy sản, Thực phẩm",
            "HOSE",
            "2020–2025",
            "GRI Standards, Ma trận tính trọng yếu chuỗi giá trị",
        ],
        [
            "PLX",
            "Tập đoàn Xăng dầu Việt Nam (Petrolimex)",
            "Năng lượng, Kinh doanh xăng dầu, Hóa dầu",
            "HOSE",
            "2020–2025",
            "GRI Standards, ISO 14064-1 Kiểm kê KNK",
        ],
        [
            "PNJ",
            "CTCP Vàng bạc Đá quý Phú Nhuận",
            "Chế tác kim hoàn, Bán lẻ trang sức thời trang",
            "HOSE",
            "2020–2025",
            "GRI Standards, Trụ cột DE&I, Phát triển nguồn nhân lực",
        ],
        [
            "SSI",
            "CTCP Chứng khoán SSI",
            "Dịch vụ tài chính, Môi giới chứng khoán, Ngân hàng đầu tư",
            "HOSE",
            "2020–2025",
            "GRI Standards, Khung tài chính xanh & Đầu tư có trách nhiệm",
        ],
        [
            "VCS",
            "CTCP Vicostone (Tập đoàn Phenikaa)",
            "Sản xuất vật liệu xây dựng cao cấp, Đá thạch anh",
            "HNX",
            "2020–2025",
            "GRI Standards, Kinh tế tuần hoàn & An toàn hóa chất",
        ],
        [
            "VNM",
            "CTCP Sữa Việt Nam (Vinamilk)",
            "Chăn nuôi bò sữa, Chế biến sữa và đồ uống dinh dưỡng",
            "HOSE",
            "2020–2025",
            "GRI Standards, PAS 2060 (Trung hòa Carbon), CDP",
        ],
    ]

    t1_headers_en = [
        "Ticker",
        "Company Name",
        "Core Industry Sector",
        "Exchange",
        "Period",
        "Reporting Framework Applied",
    ]
    t1_data_en = [
        [
            "BVH",
            "Bao Viet Holdings",
            "Finance, Life & Non-Life Insurance, Fund Management",
            "HOSE",
            "2020–2025",
            "International Integrated Reporting (<IIRC>), GRI Standards",
        ],
        [
            "PAN",
            "The PAN Group JSC",
            "Hi-Tech Agriculture, Aquaculture, Packaged Food",
            "HOSE",
            "2020–2025",
            "GRI Standards, Value-chain Materiality Matrix",
        ],
        [
            "PLX",
            "Vietnam National Petroleum Group (Petrolimex)",
            "Energy, Downstream Petroleum Distribution, Petrochemicals",
            "HOSE",
            "2020–2025",
            "GRI Standards, ISO 14064-1 GHG Inventory",
        ],
        [
            "PNJ",
            "Phu Nhuan Jewelry JSC",
            "Jewelry Manufacturing, Retail Fashion",
            "HOSE",
            "2020–2025",
            "GRI Standards, DE&I Pillars, Human Capital",
        ],
        [
            "SSI",
            "SSI Securities Corporation",
            "Financial Services, Securities Brokerage, Investment Banking",
            "HOSE",
            "2020–2025",
            "GRI Standards, Sustainable Finance & Responsible Investment",
        ],
        [
            "VCS",
            "Vicostone JSC (Phenikaa Group)",
            "Engineered Quartz Stone Surfaces, Manufacturing",
            "HNX",
            "2020–2025",
            "GRI Standards, Circular Economy & Chemical Safety",
        ],
        [
            "VNM",
            "Vietnam Dairy Products JSC (Vinamilk)",
            "Dairy Farming, Nutritional Beverage Manufacturing",
            "HOSE",
            "2020–2025",
            "GRI Standards, PAS 2060 (Carbon Neutrality), CDP",
        ],
    ]

    # Table 2: 42 reports statistics
    df_stats = pd.read_csv(DATA_DIR / "table_report_sentence_stats.csv")
    t2_headers_vi = ["Công ty", "Năm", "Tên tệp báo cáo PDF", "Số câu", "Số trang", "Mật độ (câu/trang)"]
    t2_data_vi = []
    for _, r in df_stats.iterrows():
        t2_data_vi.append([
            str(r["Công ty"]),
            str(r["Năm"]),
            str(r["Tên file"]),
            f"{int(r['Số câu']):,}",
            str(r["Số trang"]),
            f"{float(r['Số câu / trang']):.2f}",
        ])

    t2_headers_en = ["Company", "Year", "Report PDF Filename", "Sentences", "Pages", "Density (sent/page)"]
    t2_data_en = []
    for _, r in df_stats.iterrows():
        t2_data_en.append([
            str(r["Công ty"]),
            str(r["Năm"]),
            str(r["Tên file"]),
            f"{int(r['Số câu']):,}",
            str(r["Số trang"]),
            f"{float(r['Số câu / trang']):.2f}",
        ])

    # Table 3: 6 Category Means
    df_cat = pd.read_csv(DATA_DIR / "table_company_year_6cat.csv")
    t3_headers_vi = ["Công ty", "Năm", "Đời sống", "Kinh tế", "Công bằng", "Xã hội", "Tài nguyên", "Môi trường"]
    t3_data_vi = []
    for _, r in df_cat.iterrows():
        t3_data_vi.append([
            str(r["company"]),
            str(r["year"]),
            f"{float(r['Life']):.2f}",
            f"{float(r['Economic']):.2f}",
            f"{float(r['Equity']):.2f}",
            f"{float(r['Social']):.2f}",
            f"{float(r['Resources']):.2f}",
            f"{float(r['Environments']):.2f}",
        ])

    t3_headers_en = ["Company", "Year", "Life", "Economic", "Equity", "Social", "Resources", "Environments"]
    t3_data_en = []
    for _, r in df_cat.iterrows():
        t3_data_en.append([
            str(r["company"]),
            str(r["year"]),
            f"{float(r['Life']):.2f}",
            f"{float(r['Economic']):.2f}",
            f"{float(r['Equity']):.2f}",
            f"{float(r['Social']):.2f}",
            f"{float(r['Resources']):.2f}",
            f"{float(r['Environments']):.2f}",
        ])

    # Table 4: Sentiment Counts
    df_senti = pd.read_csv(DATA_DIR / "table_sentiment_counts.csv")
    t4_headers_vi = ["Công ty", "Năm", "Tiêu cực (Neg)", "Trung tính (Neu)", "Tích cực (Pos)", "Tỷ số Pos/Neg"]
    t4_data_vi = []
    for _, r in df_senti.iterrows():
        t4_data_vi.append([
            str(r["company"]),
            str(r["year"]),
            f"{int(r['Negative']):,}",
            f"{int(r['Neutral']):,}",
            f"{int(r['Positive']):,}",
            f"{float(r['Ratio']):.2f}",
        ])

    t4_headers_en = ["Company", "Year", "Negative (Neg)", "Neutral (Neu)", "Positive (Pos)", "Pos/Neg Ratio"]
    t4_data_en = []
    for _, r in df_senti.iterrows():
        t4_data_en.append([
            str(r["company"]),
            str(r["year"]),
            f"{int(r['Negative']):,}",
            f"{int(r['Neutral']):,}",
            f"{int(r['Positive']):,}",
            f"{float(r['Ratio']):.2f}",
        ])

    # Table 5: Correlation
    df_corr = pd.read_csv(DATA_DIR / "table_stock_sentiment_correlation.csv")
    t5_headers_vi = [
        "Doanh nghiệp",
        "Cỡ mẫu (n)",
        "Pearson r",
        "p-value (Pearson)",
        "Spearman ρ",
        "p-value (Spearman)",
        "Kết luận thống kê",
    ]
    t5_data_vi = []
    for _, r in df_corr.iterrows():
        comp_name = str(r["company"])
        if comp_name == "Tổng (pool)":
            c_label = "Mẫu gộp toàn bộ (Pooled Sample)"
            interp = "Không có tương quan thống kê (p = 0,79)"
        else:
            c_label = comp_name
            p_val = float(r["p_pearson"])
            interp = "Mẫu nhỏ (n=5), p ≥ 0,10 (Không suy diễn)" if p_val >= 0.10 else "Tương quan có ý nghĩa (p < 0,10)"

        t5_data_vi.append([
            c_label,
            str(r["n"]),
            f"{float(r['r_pearson']):+.3f}" if pd.notna(r["r_pearson"]) else "N/A",
            f"{float(r['p_pearson']):.3f}" if pd.notna(r["p_pearson"]) else "N/A",
            f"{float(r['r_spearman']):+.3f}" if pd.notna(r["r_spearman"]) else "N/A",
            f"{float(r['p_spearman']):.3f}" if pd.notna(r["p_spearman"]) else "N/A",
            interp,
        ])

    t5_headers_en = [
        "Company",
        "Sample (n)",
        "Pearson r",
        "p-value (Pearson)",
        "Spearman ρ",
        "p-value (Spearman)",
        "Statistical Interpretation",
    ]
    t5_data_en = []
    for _, r in df_corr.iterrows():
        comp_name = str(r["company"])
        if comp_name == "Tổng (pool)":
            c_label = "Pooled Sample (All Firms)"
            interp = "Statistically independent (p = 0.79)"
        else:
            c_label = comp_name
            p_val = float(r["p_pearson"])
            interp = (
                "Small sample (n=5), p ≥ 0.10 (Inconclusive)"
                if p_val >= 0.10
                else "Statistically significant (p < 0.10)"
            )

        t5_data_en.append([
            c_label,
            str(r["n"]),
            f"{float(r['r_pearson']):+.3f}" if pd.notna(r["r_pearson"]) else "N/A",
            f"{float(r['p_pearson']):.3f}" if pd.notna(r["p_pearson"]) else "N/A",
            f"{float(r['r_spearman']):+.3f}" if pd.notna(r["r_spearman"]) else "N/A",
            f"{float(r['p_spearman']):.3f}" if pd.notna(r["p_spearman"]) else "N/A",
            interp,
        ])

    # Table 6: Greenwashing Rank
    df_rank = pd.read_csv(DATA_DIR / "table_greenwashing_rank.csv")
    t6_headers_vi = [
        "Hạng",
        "Doanh nghiệp",
        "Số năm khảo sát",
        "Số năm giá giảm",
        "Số năm GW > 0",
        "Điểm GW trung bình",
        "Điểm GW cực đại",
        "Năm GW cực đại",
        "Số năm trung thực",
    ]
    t6_data_vi = []
    for idx, r in df_rank.iterrows():
        t6_data_vi.append([
            str(idx + 1),
            str(r["company"]),
            str(r["n_nam"]),
            str(r["n_nam_gia_tut"]),
            str(r["n_nam_GW>0"]),
            f"{float(r['GW_trung_binh']):.2f}",
            f"{float(r['GW_max']):.2f}",
            str(r["nam_GW_max"]),
            str(r["n_nam_trung_thuc"]),
        ])

    t6_headers_en = [
        "Rank",
        "Company",
        "Years Examined",
        "Price Drop Years",
        "GW > 0 Years",
        "Mean GW Score",
        "Peak GW Score",
        "Peak GW Year",
        "Honest Tone Years",
    ]
    t6_data_en = []
    for idx, r in df_rank.iterrows():
        t6_data_en.append([
            str(idx + 1),
            str(r["company"]),
            str(r["n_nam"]),
            str(r["n_nam_gia_tut"]),
            str(r["n_nam_GW>0"]),
            f"{float(r['GW_trung_binh']):.2f}",
            f"{float(r['GW_max']):.2f}",
            str(r["nam_GW_max"]),
            str(r["n_nam_trung_thuc"]),
        ])

    return {
        "t1_vi": (t1_headers_vi, t1_data_vi),
        "t1_en": (t1_headers_en, t1_data_en),
        "t2_vi": (t2_headers_vi, t2_data_vi),
        "t2_en": (t2_headers_en, t2_data_en),
        "t3_vi": (t3_headers_vi, t3_data_vi),
        "t3_en": (t3_headers_en, t3_data_en),
        "t4_vi": (t4_headers_vi, t4_data_vi),
        "t4_en": (t4_headers_en, t4_data_en),
        "t5_vi": (t5_headers_vi, t5_data_vi),
        "t5_en": (t5_headers_en, t5_data_en),
        "t6_vi": (t6_headers_vi, t6_data_vi),
        "t6_en": (t6_headers_en, t6_data_en),
    }
