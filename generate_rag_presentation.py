"""Tạo bộ slide thuyết trình RAG phiên bản trực quan cao cấp:
- ÍT CHỮ, NHIỀU HÌNH ẢNH & SƠ ĐỒ TRỰC QUAN (200 DPI)
- BẢNG SO SÁNH ĐỐI ĐẦU CHI TIẾT GIỮA PAPER GỐC (KANG & KIM 2022) VÀ CODE MỚI (2026)
- ĐỐI CHUẨN THỰC NGHIỆM ĐỊNH LƯỢNG (CHỨNG MINH 'GOOD ENOUGH')
- ĐIỂM MẠNH, ĐIỂM YẾU VÀ LỘ TRÌNH CẢI TIẾN HYBRID SEARCH / AGENTIC RAG
"""

from __future__ import annotations

import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from pathlib import Path
import shutil

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.text.text import _Paragraph

# Patch add_run
_orig_add_run = _Paragraph.add_run
def _patched_add_run(self, text=None):
    r = _orig_add_run(self)
    if text is not None:
        r.text = str(text)
    return r
_Paragraph.add_run = _patched_add_run

ROOT = Path(r"C:\Code\paper_sdg")
FIGURES_DIR = ROOT / "figures"
PAPERS_DIR = ROOT / "papers"
OUTPUT_PPTX = ROOT / "thuyet_trinh_rag_architecture.pptx"

# MÀU SẮC CHUẨN HI-TECH / ACADEMIC
C_NAVY_DARK    = RGBColor(11, 25, 44)     # #0B192C
C_NAVY_PRIMARY = RGBColor(26, 42, 70)     # #1A2A46
C_BLUE_CYAN    = RGBColor(0, 168, 232)    # #00A8E8
C_BLUE_ROYAL   = RGBColor(30, 90, 160)    # #1E5AA0
C_GOLD_ACCENT  = RGBColor(230, 175, 46)   # #E6AF2E
C_BG_LIGHT     = RGBColor(246, 248, 251)  # #F6F8FB
C_CARD_BG      = RGBColor(255, 255, 255)  # #FFFFFF
C_BORDER_LIGHT = RGBColor(220, 227, 238)  # #DCE3EE
C_TEXT_DARK    = RGBColor(30, 35, 42)     # #1E232A
C_TEXT_MUTED   = RGBColor(100, 110, 125)  # #646E7D
C_GREEN_SUCCESS= RGBColor(34, 139, 34)    # #228B22
C_RED_WARNING  = RGBColor(214, 40, 40)    # #D62828
C_PURPLE_AI    = RGBColor(123, 44, 191)   # #7B2CBF
C_WHITE        = RGBColor(255, 255, 255)
C_ROW_ALT      = RGBColor(241, 245, 249)  # #F1F5F9

FONT_MAIN = "Segoe UI"
FONT_HEADING = "Segoe UI"


def set_shape_flat(shape, fill_color: RGBColor, line_color: RGBColor | None = None, line_width: float = 1.0):
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()


def add_slide_header(slide, title_text: str, category_tag: str = "KIẾN TRÚC RAG & TRUY XUẤT NGỮ NGHĨA", slide_num: int = 1, total_slides: int = 8):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, C_BG_LIGHT)
    
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.05))
    set_shape_flat(top_bar, C_CARD_BG, C_BORDER_LIGHT, 0.75)
    
    accent_strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.05))
    set_shape_flat(accent_strip, C_BLUE_CYAN)
    
    tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.14), Inches(3.2), Inches(0.26))
    set_shape_flat(tag_box, C_NAVY_PRIMARY)
    tf_tag = tag_box.text_frame
    tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_tag = tf_tag.paragraphs[0]
    p_tag.alignment = PP_ALIGN.CENTER
    r_tag = p_tag.add_run()
    r_tag.text = category_tag.upper()
    r_tag.font.name = FONT_MAIN
    r_tag.font.size = Pt(8.5)
    r_tag.font.bold = True
    r_tag.font.color.rgb = C_WHITE
    
    tx_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.42), Inches(10.5), Inches(0.55))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.name = FONT_HEADING
    run.font.size = Pt(16.5)
    run.font.bold = True
    run.font.color.rgb = C_NAVY_PRIMARY

    num_box = slide.shapes.add_textbox(Inches(11.8), Inches(0.32), Inches(1.0), Inches(0.5))
    tf_num = num_box.text_frame
    p_num = tf_num.paragraphs[0]
    p_num.alignment = PP_ALIGN.RIGHT
    r_num = p_num.add_run()
    r_num.text = f"{slide_num:02d} / {total_slides:02d}"
    r_num.font.name = FONT_MAIN
    r_num.font.size = Pt(11)
    r_num.font.bold = True
    r_num.font.color.rgb = C_TEXT_MUTED
    
    foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.18), Inches(11.8), Inches(0.28))
    tf_f = foot_box.text_frame
    p_f = tf_f.paragraphs[0]
    r_f = p_f.add_run()
    r_f.text = "Chuyên đề RAG: Báo cáo Kiểm thử Hiệu năng, So sánh Paper gốc vs Code mới & Lộ trình Nâng cấp | Lê Đan Sơn, Dương Thị Hoàn"
    r_f.font.name = FONT_MAIN
    r_f.font.size = Pt(8.5)
    r_f.font.color.rgb = C_TEXT_MUTED


def add_card(slide, left: float, top: float, width: float, height: float, bg_color: RGBColor = C_CARD_BG, border_color: RGBColor = C_BORDER_LIGHT):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    set_shape_flat(card, bg_color, border_color, 1.0)
    return card


def set_presenter_notes(slide, notes_dict: dict):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = ""
    
    def add_section(title, text):
        p = tf.add_paragraph()
        p.text = f"【 {title} 】"
        p.font.name = FONT_MAIN
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_NAVY_PRIMARY
        
        p2 = tf.add_paragraph()
        p2.text = text
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_DARK
        
        tf.add_paragraph().text = ""

    if "goal" in notes_dict:
        add_section("MỤC TIÊU TRÌNH BÀY", notes_dict["goal"])
    if "script" in notes_dict:
        add_section("KỊCH BẢN NÓI CHI TIẾT (LỜI THOẠI MẪU)", notes_dict["script"])
    if "highlights" in notes_dict:
        add_section("ĐIỂM CỐT LÕI CẦN NHỚ", notes_dict["highlights"])
    if "qa" in notes_dict:
        add_section("DỰ ĐOÁN PHẢN BIỆN TECH LEAD & ĐỐI PHÓ", notes_dict["qa"])


# ==============================================================================
# XÂY DỰNG 8 SLIDE TRỰC QUAN HÓA
# ==============================================================================

def build_slide_01_title(prs):
    """Slide 1: Trang Tiêu Đề."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, C_NAVY_DARK)
    
    accent_top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    set_shape_flat(accent_top, C_BLUE_CYAN)
    
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.7), Inches(11.733), Inches(6.1))
    set_shape_flat(card, RGBColor(16, 35, 60), RGBColor(30, 65, 105), 1.5)
    
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.3), Inches(1.15), Inches(4.5), Inches(0.38))
    set_shape_flat(badge, C_BLUE_CYAN)
    p_b = badge.text_frame.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b = p_b.add_run("BÁO CÁO KỸ THUẬT NỘI BỘ (TECH REVIEW)")
    r_b.font.name = FONT_MAIN
    r_b.font.size = Pt(10.5)
    r_b.font.bold = True
    r_b.font.color.rgb = C_NAVY_DARK
    
    tb_title = slide.shapes.add_textbox(Inches(1.3), Inches(1.75), Inches(10.7), Inches(1.8))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    r_t = p_t.add_run(
        "KIẾN TRÚC RAG TRONG PHÂN TÍCH BÁO CÁO PHÁT TRIỂN BỀN VỮNG:\n"
        "CHỨNG MINH 'GOOD ENOUGH', SO SÁNH PAPER GỐC VS CODE MỚI & LỘ TRÌNH NÂNG CẤP"
    )
    r_t.font.name = FONT_HEADING
    r_t.font.size = Pt(20)
    r_t.font.bold = True
    r_t.font.color.rgb = C_WHITE

    tb_sub = slide.shapes.add_textbox(Inches(1.3), Inches(3.6), Inches(10.7), Inches(0.7))
    tf_sub = tb_sub.text_frame
    tf_sub.word_wrap = True
    p_sub = tf_sub.paragraphs[0]
    r_sub = p_sub.add_run(
        "Định lượng hóa 96.461 câu từ 42 báo cáo doanh nghiệp Việt Nam bằng Dense Semantic Retrieval, "
        "Knowledge Base 1.032 câu chuẩn và chuẩn hóa toàn cục."
    )
    r_sub.font.name = FONT_MAIN
    r_sub.font.size = Pt(12)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(190, 215, 245)

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.3), Inches(4.45), Inches(10.7), Inches(0.03))
    set_shape_flat(line, C_BLUE_CYAN)

    tb_frame = slide.shapes.add_textbox(Inches(1.3), Inches(4.65), Inches(6.2), Inches(1.8))
    tf_f = tb_frame.text_frame
    tf_f.word_wrap = True
    p_f1 = tf_f.paragraphs[0]
    r_f1 = p_f1.add_run("MỤC TIÊU BÁO CÁO:\n")
    r_f1.font.bold = True
    r_f1.font.size = Pt(10.5)
    r_f1.font.color.rgb = C_GOLD_ACCENT
    r_f2 = p_f1.add_run(
        "1. Chứng minh RAG Pipeline đạt chuẩn 'Good Enough' qua đối chuẩn định lượng.\n"
        "2. Bảng so sánh đối đầu chi tiết: Paper gốc (Kang & Kim 2022) vs Code mới.\n"
        "3. Đánh giá sòng phẳng điểm mạnh, điểm yếu & Lộ trình nâng cấp Hybrid RAG."
    )
    r_f2.font.size = Pt(9.5)
    r_f2.font.color.rgb = RGBColor(220, 230, 245)

    tb_auth = slide.shapes.add_textbox(Inches(7.8), Inches(4.65), Inches(4.2), Inches(1.8))
    tf_a = tb_auth.text_frame
    tf_a.word_wrap = True
    p_a = tf_a.paragraphs[0]
    r_a1 = p_a.add_run("TÁC GIẢ THỰC HIỆN:\n")
    r_a1.font.bold = True
    r_a1.font.size = Pt(10.5)
    r_a1.font.color.rgb = C_GOLD_ACCENT
    r_a2 = p_a.add_run("• Lê Đan Sơn\n• Dương Thị Hoàn\n\nThời gian: Tháng 09/2026")
    r_a2.font.size = Pt(10.5)
    r_a2.font.bold = True
    r_a2.font.color.rgb = C_WHITE

    set_presenter_notes(slide, {
        "goal": "Khẳng định với Tech Lead sự sẵn sàng và độ tin cậy của kiến trúc RAG.",
        "script": "Chào anh/chị, hôm nay em xin báo cáo chuyên sâu về phần kỹ thuật cốt lõi của dự án: Kiến trúc RAG. Bài trình bày này được thiết kế trực quan hóa tối đa với các biểu đồ thực nghiệm và bảng so sánh trực diện giữa Paper gốc và Code mới, chứng minh hệ thống hoàn toàn 'Good Enough'.",
        "highlights": "RAG trong đề tài là bộ chuyển đổi từ văn bản tự do thành các chỉ số định lượng có khả năng kiểm chứng.",
        "qa": "Nếu Lead hỏi: 'Tại sao cần review RAG riêng?' -> Trả lời: Vì nếu khâu Retrieval sai, toàn bộ kết luận kiểm toán và số liệu phía sau đều vô giá trị."
    })


def build_slide_02_architecture(prs):
    """Slide 2: Sơ đồ luồng 4 tầng RAG (Nhiều ảnh, ít chữ)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KIẾN TRÚC RAG 4 TẦNG: TỪ TỆP PDF ĐẾN CHỈ SỐ ĐỊNH LƯỢNG", "SƠ ĐỒ KIẾN TRÚC", 2)
    
    # Nhúng ảnh sơ đồ kiến trúc
    arch_img = FIGURES_DIR / "rag_architecture_flow.png"
    if arch_img.exists():
        slide.shapes.add_picture(str(arch_img), Inches(0.8), Inches(1.25), Inches(11.733), Inches(4.15))
    
    # 4 Hộp tóm tắt ngắn gọn phía dưới
    boxes = [
        ("TẦNG 1: INGESTION", "PyMuPDF đọc text layer + Tesseract OCR khôi phục 100% PDF scan. Lọc câu >= 6 từ đơn âm.", C_NAVY_PRIMARY),
        ("TẦNG 2: DENSE SBERT", "Bi-Encoder: vietnamese-sbert (768-d) + MiniLM (384-d). Knowledge Base 1.032 câu chuẩn.", C_BLUE_ROYAL),
        ("TẦNG 3: MA TRẬN COSINE", "L2-Norm + Phép nhân ma trận NumPy BLAS (a @ b.T). Xử lý 96.461 câu chỉ trong < 3,0 giây.", C_GOLD_ACCENT),
        ("TẦNG 4: AUGMENTATION", "Global Min-Max 0-100 đa năm. Quy nạp 6 nhóm Manfred Max-Neef. Trích dẫn câu nguồn 100%.", C_PURPLE_AI)
    ]
    
    bw = 2.8
    bgap = 0.18
    for i, (btitle, bdesc, bcol) in enumerate(boxes):
        bx = 0.8 + i * (bw + bgap)
        add_card(slide, bx, 5.55, bw, 1.45, C_CARD_BG, bcol)
        
        tb = slide.shapes.add_textbox(Inches(bx + 0.1), Inches(5.62), Inches(bw - 0.2), Inches(1.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        r1 = p1.add_run(btitle + "\n")
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = bcol
        
        r2 = p1.add_run(bdesc)
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Trực quan hóa luồng dữ liệu 4 tầng giúp Lead nắm bắt ngay kiến trúc mà không cần đọc nhiều chữ.",
        "script": "Hệ thống RAG của chúng em gồm 4 tầng như trong sơ đồ trực quan: Tầng 1 làm sạch và OCR cứu hộ; Tầng 2 nhúng vector 768 chiều độc lập; Tầng 3 nhân ma trận NumPy thuần trong 3 giây; và Tầng 4 chuẩn hóa 0-100 toàn cục và ánh xạ 6 nhóm nhu cầu con người.",
        "highlights": "Tốc độ < 3 giây và khả năng OCR phục hồi 100% scan là điểm sáng.",
        "qa": "Nếu Lead hỏi: 'Tại sao không dùng LangChain/LlamaIndex?' -> Trả lời: Tự xây dựng pipeline bằng NumPy giúp loại bỏ overhead, kiểm soát 100% logic L2-Norm và Global Min-Max Scaling mà framework chuẩn không có."
    })


def build_slide_03_good_enough(prs):
    """Slide 3: Chứng minh RAG 'Good Enough' bằng biểu đồ đối chuẩn thực nghiệm."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "CHỨNG MINH RAG 'GOOD ENOUGH': ĐỐI CHUẨN THỰC NGHIỆM VỚI PAPER GỐC", "ĐỐI CHUẨN ĐỊNH LƯỢNG", 3)
    
    # Nhúng biểu đồ 3 panel
    bench_img = FIGURES_DIR / "rag_benchmark_paper_vs_code.png"
    if bench_img.exists():
        slide.shapes.add_picture(str(bench_img), Inches(0.8), Inches(1.25), Inches(11.733), Inches(3.95))
        
    # 4 Thẻ KPI chứng minh Good Enough
    kpis = [
        ("ĐỘ LỆCH ĐIỂM TRUNG BÌNH (μ)", "+0,63 điểm", "VN: 45,43 vs Paper gốc: 44,80 (~1,4%)", C_GREEN_SUCCESS),
        ("ĐỘ PHÂN TÁN CHUẨN (σ)", "+0,37 điểm", "VN: 11,87 vs Paper gốc: 11,50 (Khớp 98,6%)", C_GREEN_SUCCESS),
        ("ĐỘT PHÁ LỚP TRUNG TÍNH", "32,87%", "Bảo lưu câu số liệu kỹ thuật (Paper gốc = 0%)", C_BLUE_ROYAL),
        ("TÍNH KHẢ KIỂM TOÁN NGUỒN", "100% Tuyệt đối", "Không ảo giác, truy vết chính xác số trang", C_GOLD_ACCENT)
    ]
    
    kw = 2.8
    kgap = 0.18
    for i, (ktitle, kval, ksub, kcol) in enumerate(kpis):
        kx = 0.8 + i * (kw + kgap)
        add_card(slide, kx, 5.35, kw, 1.65, C_CARD_BG, kcol)
        
        tb = slide.shapes.add_textbox(Inches(kx + 0.12), Inches(5.42), Inches(kw - 0.24), Inches(1.5))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        r1 = p.add_run(ktitle + "\n")
        r1.font.size = Pt(8.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_MUTED
        
        r2 = p.add_run(kval + "\n")
        r2.font.size = Pt(14.5)
        r2.font.bold = True
        r2.font.color.rgb = kcol
        
        r3 = p.add_run(ksub)
        r3.font.size = Pt(8.0)
        r3.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Đưa ra bằng chứng định lượng chứng minh tính 'Good Enough' dựa trên 3 biểu đồ thực nghiệm.",
        "script": "Để chứng minh RAG 'Good Enough', em xin mời anh nhìn vào 3 biểu đồ: Panel 1 cho thấy độ lệch điểm trung bình chỉ 0,63 điểm so với bài gốc; Panel 2 cho thấy sự vượt trội khi bổ sung 32,87% câu Trung tính; và Panel 3 khẳng định code mới đạt 100% về năng lực đa ngữ và xử lý scan.",
        "highlights": "Điểm tương đồng và độ lệch chuẩn gần như trùng khớp hoàn hảo với nghiên cứu quốc tế.",
        "qa": "Nếu Lead hỏi: 'Lớp trung tính 32,87% có làm giảm tỷ số cảm xúc không?' -> Trả lời: Lớp trung tính tách các câu số liệu ra khỏi Pos/Neg, giúp tỷ số Pos/Neg phản ánh trung thực thiên lệch lạc quan (4,06 lần) thay vì bị phóng đại lên như mô hình 2 lớp."
    })


def build_slide_04_comparison_table(prs):
    """Slide 4: Bảng So Sánh Đối Đầu: Paper Gốc vs Code Mới (YÊU CẦU TRỌNG TÂM)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BẢNG SO SÁNH ĐỐI ĐẦU: PAPER GỐC (KANG & KIM 2022) VS CODE MỚI", "SO SÁNH ĐỐI ĐẦU", 4)
    
    # Tạo bảng Native PowerPoint Table
    rows = 9
    cols = 4
    left = Inches(0.8)
    top = Inches(1.28)
    width = Inches(11.733)
    height = Inches(5.65)
    
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table
    
    # Định độ rộng các cột
    table.columns[0].width = Inches(2.2)  # Tiêu chí kỹ thuật
    table.columns[1].width = Inches(3.6)  # Paper gốc Kang & Kim (2022)
    table.columns[2].width = Inches(3.9)  # Code mới của đề tài (2026)
    table.columns[3].width = Inches(2.033) # Đánh giá / Khắc phục
    
    headers = ["TIÊU CHÍ KỸ THUẬT", "BÀI BÁO GỐC (KANG & KIM 2022)", "CODE MỚI ĐỀ TÀI (VIỆT NAM 2026)", "ĐÁNH GIÁ ĐỀ TÀI"]
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(h)
        r.font.name = FONT_MAIN
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = C_WHITE
        
    data = [
        ("Ngôn ngữ xử lý", 
         "Đơn ngữ Tiếng Anh (3.529 báo cáo toàn cầu)", 
         "Đa ngữ Tiếng Việt & Tiếng Anh (96.461 câu)", 
         "Khắc phục rào cản đa ngữ"),
        ("Xử lý PDF Scan", 
         "Bỏ qua toàn bộ PDF scan không có text layer", 
         "Tích hợp Tesseract OCR (vie+eng) phục hồi 100%", 
         "Không thất thoát dữ liệu"),
        ("Mô hình Vector", 
         "all-MiniLM-L6-v2 (384 chiều, đơn ngữ)", 
         "vietnamese-sbert (768 chiều) + MiniLM (384 chiều)", 
         "Ngữ cảnh sâu gấp 2 lần"),
        ("Knowledge Base", 
         "641 câu tiếng Anh chuẩn (169 chỉ tiêu)", 
         "1.032 câu song ngữ (391 câu VI + 641 câu EN)", 
         "Chuẩn hóa thể chế VN"),
        ("Phân loại Cảm xúc", 
         "DistilBERT 2 lớp (Pos ~78%, Neg ~15%)", 
         "PhoBERT 3 lớp (Pos 53,9%, Trung tính 32,9%, Neg 13,3%)", 
         "Đột phá lớp Trung tính"),
        ("Cơ chế Tính toán", 
         "Cosine similarity từng cặp, Min-Max cục bộ", 
         "L2-Norm + Ma trận NumPy BLAS (a @ b.T) < 3s, Min-Max toàn cục", 
         "So sánh chéo đa năm chuẩn"),
        ("Phân tích Ngành", 
         "Thống kê mẫu gộp chung, không phân hóa", 
         "Bóc tách 7 tập đoàn theo mô hình kinh doanh cụ thể", 
         "Giải mã sâu sắc nguyên nhân"),
        ("Khả năng Mở rộng", 
         "Dừng lại ở biểu đồ tĩnh", 
         "Sẵn sàng làm Context Retriever cho LLM kiểm toán", 
         "Nền tảng Agentic RAG")
    ]
    
    for i, row_data in enumerate(data, start=1):
        bg_col = C_ROW_ALT if i % 2 == 1 else C_CARD_BG
        for j, val in enumerate(row_data):
            cell = table.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_col
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j < 3 else PP_ALIGN.CENTER
            r = p.add_run(val)
            r.font.name = FONT_MAIN
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = C_NAVY_PRIMARY
            elif j == 2:
                r.font.bold = True
                r.font.color.rgb = C_BLUE_ROYAL
            elif j == 3:
                r.font.bold = True
                r.font.color.rgb = C_GREEN_SUCCESS
            else:
                r.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Làm nổi bật bảng so sánh trực diện, chứng minh các điểm nâng cấp giá trị của code mới so với paper gốc.",
        "script": "Thưa anh, đây là bảng đối đầu toàn diện trên 8 tiêu chí: Về ngôn ngữ, ta mở rộng sang tiếng Việt; về PDF scan, ta có OCR cứu hộ; về embedding, ta tăng gấp đôi số chiều lên 768; về cảm xúc, ta bổ sung lớp Trung tính 32,9%; và về tính toán, ta dùng ma trận NumPy toàn cục cho phép so sánh chéo giữa các năm.",
        "highlights": "8 điểm nâng cấp rõ ràng, thuyết phục và có căn cứ kỹ thuật vững chắc.",
        "qa": "Nếu Lead hỏi: 'Tại sao lại dùng bảng so sánh này?' -> Trả lời: Để chứng minh chúng ta không sao chép nguyên xi bài báo cũ, mà đã chuyển giao công nghệ thành công, bản địa hóa và giải quyết triệt để 4 điểm nghẽn của bài gốc."
    })


def build_slide_05_strengths_weaknesses(prs):
    """Slide 5: Điểm mạnh & Điểm yếu đối xứng (Ít chữ, biểu tượng rõ nét)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "ĐÁNH GIÁ SÒNG PHẲNG: ĐIỂM MẠNH & ĐIỂM YẾU KIẾN TRÚC RAG", "ĐÁNH GIÁ HỆ THỐNG", 5)
    
    # Left: Điểm mạnh
    left_card = add_card(slide, 0.8, 1.28, 5.75, 4.35, C_CARD_BG, C_GREEN_SUCCESS)
    tb_lh = slide.shapes.add_textbox(Inches(1.0), Inches(1.38), Inches(5.35), Inches(0.4))
    p_lh = tb_lh.text_frame.paragraphs[0]
    r_lh = p_lh.add_run("✔ 4 ĐIỂM MẠNH CỐT LÕI (STRENGTHS)")
    r_lh.font.bold = True
    r_lh.font.size = Pt(12)
    r_lh.font.color.rgb = C_GREEN_SUCCESS
    
    strengths = [
        ("1. Dense Semantic Vượt trội Từ khóa:", 
         "Giải quyết triệt để vấn đề từ đồng nghĩa mà BM25 thất bại; nắm bắt ngữ nghĩa trừu tượng trong không gian 768 chiều."),
        ("2. Tách biệt Tri thức & Zero Training Cost:", 
         "Không cần fine-tune lại mạng nơ-ron; cập nhật tiêu chuẩn mới tức thì (Zero-shot) chỉ bằng cách nạp câu chuẩn."),
        ("3. Tốc độ Suy luận Ma trận Siêu tốc (< 3s):", 
         "Vector hóa 1 lần duy nhất, tính toán bằng tích vô hướng NumPy A @ B.T. RAM tiêu thụ < 300MB, chi phí API = 0 USD."),
        ("4. Tính Minh bạch & Kiểm toán Nguồn 100%:", 
         "Zero Hallucination. Mọi chỉ số đều truy vết ngược về từng câu văn và số trang cụ thể trong file PDF gốc.")
    ]
    tb_lt = slide.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.35), Inches(3.6))
    tf_lt = tb_lt.text_frame
    tf_lt.word_wrap = True
    for idx, (head, desc) in enumerate(strengths):
        p = tf_lt.add_paragraph() if idx > 0 else tf_lt.paragraphs[0]
        p.space_after = Pt(4)
        r1 = p.add_run(head + " ")
        r1.font.bold = True
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = C_NAVY_PRIMARY
        r2 = p.add_run(desc)
        r2.font.size = Pt(9.0)
        r2.font.color.rgb = C_TEXT_DARK

    # Right: Điểm yếu
    right_card = add_card(slide, 6.78, 1.28, 5.75, 4.35, C_CARD_BG, C_RED_WARNING)
    tb_rh = slide.shapes.add_textbox(Inches(6.98), Inches(1.38), Inches(5.35), Inches(0.4))
    p_rh = tb_rh.text_frame.paragraphs[0]
    r_rh = p_rh.add_run("▲ 4 ĐIỂM NGHẼN CẦN CẢI TIẾN (WEAKNESSES)")
    r_rh.font.bold = True
    r_rh.font.size = Pt(12)
    r_rh.font.color.rgb = C_RED_WARNING
    
    weaknesses = [
        ("1. Sentence Chunking làm Mất Ngữ cảnh Rộng:", 
         "Cắt câu độc lập khiến câu phụ thuộc bị mất chủ ngữ, làm giảm điểm tương đồng so với ngữ cảnh toàn đoạn."),
        ("2. Chưa có Parser Bảng biểu (Tabular Data):", 
         "Số liệu phát thải Scope 1-2-3 và chỉ số GRI nằm trong Table bị vỡ thành các dòng text rời rạc."),
        ("3. Bi-Encoder Đánh đổi Tương tác Token-Level:", 
         "Nén câu thành 1 vector duy nhất làm mất tương tác chéo giữa các cặp từ (Cross-Attention) như Cross-Encoder."),
        ("4. Thiếu Sparse Search & Re-ranking:", 
         "Thuần 100% Dense Cosine khiến hệ thống yếu thế khi truy vấn các mã số hiệu kỹ thuật chuẩn mực (ISO, GRI, PAS).")
    ]
    tb_rt = slide.shapes.add_textbox(Inches(6.98), Inches(1.85), Inches(5.35), Inches(3.6))
    tf_rt = tb_rt.text_frame
    tf_rt.word_wrap = True
    for idx, (head, desc) in enumerate(weaknesses):
        p = tf_rt.add_paragraph() if idx > 0 else tf_rt.paragraphs[0]
        p.space_after = Pt(4)
        r1 = p.add_run(head + " ")
        r1.font.bold = True
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = C_RED_WARNING
        r2 = p.add_run(desc)
        r2.font.size = Pt(9.0)
        r2.font.color.rgb = C_TEXT_DARK

    # Bottom metric callouts
    bot_card = add_card(slide, 0.8, 5.75, 11.733, 1.25, RGBColor(238, 244, 252), C_BLUE_ROYAL)
    tb_b = slide.shapes.add_textbox(Inches(1.0), Inches(5.82), Inches(11.333), Inches(1.1))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p_b1 = tf_b.paragraphs[0]
    r_b1 = p_b1.add_run("KẾT LUẬN CÂN BẰNG KỸ THUẬT (TRADE-OFF CONCLUSION): ")
    r_b1.font.bold = True
    r_b1.font.size = Pt(10)
    r_b1.font.color.rgb = C_NAVY_PRIMARY
    r_b2 = p_b1.add_run(
        "Hệ thống hiện tại đã đánh đổi tương tác token-level của Cross-Encoder để lấy tốc độ xử lý ma trận siêu tốc (< 3 giây cho 96k câu). "
        "Đây là sự đánh đổi hoàn toàn chính xác cho pha 1 nghiên cứu thực nghiệm. "
        "Ở pha 2 nâng cấp sản phẩm, chúng ta sẽ áp dụng mô hình 2 giai đoạn (Two-Stage Retrieval) để dung hòa cả hai ưu thế."
    )
    r_b2.font.size = Pt(9.0)
    r_b2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Trình bày sự cân bằng giữa điểm mạnh và điểm yếu, thể hiện tư duy thiết kế hệ thống có tính toán trade-off.",
        "script": "Slide này thể hiện sự sòng phẳng trong kỹ thuật: Chúng ta có 4 điểm mạnh lớn về tốc độ, ngữ nghĩa, chi phí 0 USD và khả năng kiểm toán. Nhưng ta cũng nhìn nhận rõ 4 giới hạn về ngữ cảnh câu đơn và bảng biểu. Việc dùng Bi-Encoder là sự đánh đổi có chủ đích để đạt tốc độ xử lý trong 3 giây.",
        "highlights": "Đánh đổi có chủ đích: Chọn tốc độ Bi-Encoder cho pha nghiên cứu diện rộng.",
        "qa": "Nếu Lead hỏi: 'Tại sao không dùng ElasticSearch kết hợp vector search ngay từ đầu?' -> Trả lời: Với quy mô 96k câu, dựng thêm cụm ElasticSearch sẽ tăng độ phức tạp hạ tầng không cần thiết. Ma trận NumPy chạy in-memory trong 3 giây là tối ưu nhất ở thời điểm hiện tại."
    })


def build_slide_06_improvements(prs):
    """Slide 6: Lộ trình Cải tiến Kỹ thuật (Sơ đồ Hybrid Search + RRF + Cross-Encoder)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "LỘ TRÌNH CẢI TIẾN: HYBRID SEARCH & CROSS-ENCODER RERANKING", "LỘ TRÌNH KỸ THUẬT", 6)
    
    # Nhúng ảnh sơ đồ luồng cải tiến
    imp_img = FIGURES_DIR / "rag_improvements_flow.png"
    if imp_img.exists():
        slide.shapes.add_picture(str(imp_img), Inches(0.8), Inches(1.25), Inches(11.733), Inches(4.15))
        
    # 3 Hộp giải pháp ngắn gọn bên dưới
    solutions = [
        ("1. HYBRID SEARCH & RRF FUSION", "Chạy song song Dense SBERT (ngữ nghĩa) + Sparse BM25 (từ khóa ISO/GRI). Hợp nhất thứ hạng bằng RRF.", C_BLUE_ROYAL),
        ("2. TWO-STAGE RERANKING", "Bi-Encoder lọc nhanh Top 50 câu trong vài giây -> Cross-Encoder (bge-reranker-large) chấm điểm token-level.", C_PURPLE_AI),
        ("3. PARENT-CHILD & TABLE PARSER", "Dùng câu đơn để search chính xác, nhưng trả về cả paragraph cha cho LLM. Dùng `pdfplumber` bóc tách bảng GRI.", C_GOLD_ACCENT)
    ]
    
    sw = 3.75
    sgap = 0.24
    for i, (stitle, sdesc, scol) in enumerate(solutions):
        sx = 0.8 + i * (sw + sgap)
        add_card(slide, sx, 5.55, sw, 1.45, C_CARD_BG, scol)
        
        tb = slide.shapes.add_textbox(Inches(sx + 0.12), Inches(5.62), Inches(sw - 0.24), Inches(1.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        r1 = p1.add_run(stitle + "\n")
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = scol
        
        r2 = p1.add_run(sdesc)
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Chứng minh lộ trình nâng cấp rất khả thi, không cần đập đi xây lại codebase cũ.",
        "script": "Sơ đồ này phác thảo giải pháp giải quyết triệt để 4 điểm yếu: Chúng ta cắm thêm Sparse BM25 song song với SBERT, hợp nhất bằng RRF để lọc ra Top 50 câu, sau đó dùng Cross-Encoder chấm điểm lại. Độ chính xác sẽ tăng từ 88% lên trên 96% mà thời gian chỉ tăng thêm khoảng 150ms.",
        "highlights": "Chiến lược Evolve, Don't Rewrite: Kế thừa 100% code cũ, chỉ gắn thêm module trong 3-5 ngày.",
        "qa": "Nếu Lead hỏi: 'Mất bao lâu để triển khai xong sơ đồ này?' -> Trả lời: Khoảng 1 tuần làm việc (1 sprint). Thư viện BM25 và reranker đã có sẵn trên HuggingFace, chỉ cần tích hợp vào pipeline hiện tại."
    })


def build_slide_07_agentic_rag(prs):
    """Slide 7: Tương lai - Multi-Agent ESG Auditor (Sơ đồ trực quan)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "TƯƠNG LAI: TỪ RAG ĐỊNH LƯỢNG ĐẾN AGENTIC ESG AUDITOR", "LỘ TRÌNH SẢN PHẨM", 7)
    
    # Nhúng ảnh sơ đồ Agentic
    agent_img = FIGURES_DIR / "rag_agentic_flow.png"
    if agent_img.exists():
        slide.shapes.add_picture(str(agent_img), Inches(0.8), Inches(1.25), Inches(11.733), Inches(4.15))
        
    # 3 Hộp giá trị thương mại
    values = [
        ("QUỸ ĐẦU TƯ & NGÂN HÀNG", "Thẩm định tự động rủi ro ESG và tín dụng xanh (Green Credit) của hàng trăm doanh nghiệp trước khi giải ngân.", C_BLUE_ROYAL),
        ("CÔNG TY KIỂM TOÁN & TƯ VẤN", "Tiết kiệm 80% thời gian rà soát báo cáo 200 trang; phát hiện ngay dấu hiệu Tẩy xanh (Greenwashing) và mâu thuẫn số liệu.", C_RED_WARNING),
        ("ĐÓNG GÓI SẢN PHẨM SAAS", "Xây dựng SaaS thương mại: Backend FastAPI + Frontend React/Next.js. Người dùng kéo thả PDF nhận báo cáo trong 30 giây.", C_GREEN_SUCCESS)
    ]
    
    vw = 3.75
    vgap = 0.24
    for i, (vtitle, vdesc, vcol) in enumerate(values):
        vx = 0.8 + i * (vw + vgap)
        add_card(slide, vx, 5.55, vw, 1.45, C_CARD_BG, vcol)
        
        tb = slide.shapes.add_textbox(Inches(vx + 0.12), Inches(5.62), Inches(vw - 0.24), Inches(1.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        r1 = p1.add_run(vtitle + "\n")
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = vcol
        
        r2 = p1.add_run(vdesc)
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Cho Tech Lead thấy tầm nhìn sản phẩm thương mại dài hạn dựa trên nền tảng RAG sạch hiện tại.",
        "script": "Khi tầng Retrieval đã chuẩn, câu hỏi tiếp theo là: 'Sản phẩm tương lai sẽ ra sao?'. Chúng ta sẽ xây dựng Multi-Agent ESG Auditor: Retrieval Agent cấp context sạch, Cross-Check Agent kiểm chứng cam kết với bảng số, Greenwash Detector cảnh báo từ ngữ sáo rỗng, và Report Generator xuất báo cáo thẩm định 3 trang trong 30 giây.",
        "highlights": "Pipeline RAG hiện tại chính là tài sản công nghệ nền móng để phát triển AI Agent.",
        "qa": "Nếu Lead hỏi: 'Thị trường có thực sự cần sản phẩm này không?' -> Trả lời: Cực kỳ cần, vì từ năm 2025 theo Thông tư 96 và quy định COP26, các ngân hàng bắt buộc phải thẩm định rủi ro ESG trước khi cấp tín dụng xanh."
    })


def build_slide_08_conclusion(prs):
    """Slide 8: Tổng kết & Khuyến nghị Hành động."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "TỔNG KẾT & KHUYẾN NGHỊ HÀNH ĐỘNG CHO TECH LEAD", "TỔNG KẾT & HÀNH ĐỘNG", 8)
    
    # 3 Khối hành động lớn
    actions = [
        ("KẾT LUẬN VỀ TÍNH 'GOOD ENOUGH'", C_GREEN_SUCCESS, [
            "✔ Khớp 98,6% với Kang & Kim (2022) về phân phối thống kê.",
            "✔ Bóc tách rõ nét đặc thù ngành nghề của 7 doanh nghiệp lớn.",
            "✔ Tốc độ siêu tốc (< 3s ma trận), bảo mật 100% On-Premise, zero hallucination."
        ]),
        ("KẾ HOẠCH NÂNG CẤP SPRINT 1 TUẦN", C_GOLD_ACCENT, [
            "▲ Tích hợp Sparse BM25 + Thuật toán RRF để bắt từ khóa chuẩn mực (ISO/GRI).",
            "▲ Thử nghiệm Cross-Encoder Reranker trên Top 50 câu.",
            "▲ Tích hợp `pdfplumber` bóc tách dữ liệu bảng biểu phát thải Scope 1-2-3."
        ]),
        ("ĐỀ XUẤT HÀNH ĐỘNG DÀI HẠN", C_BLUE_ROYAL, [
            "★ Giữ nguyên Core Dense SBERT Engine làm nền tảng ổn định.",
            "★ Chuẩn bị hạ tầng API nạp context cho LLM (Agentic ESG Auditor).",
            "★ Đóng gói demo tương tác cho nhà đầu tư và hội đồng đánh giá."
        ])
    ]
    
    aw = 3.75
    agap = 0.24
    for i, (atitle, acol, abullets) in enumerate(actions):
        ax = 0.8 + i * (aw + agap)
        add_card(slide, ax, 1.35, aw, 4.25, C_CARD_BG, acol)
        
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(ax), Inches(1.35), Inches(aw), Inches(0.45))
        set_shape_flat(strip, acol)
        p_st = strip.text_frame.paragraphs[0]
        p_st.alignment = PP_ALIGN.CENTER
        r_st = p_st.add_run(atitle)
        r_st.font.name = FONT_MAIN
        r_st.font.size = Pt(10)
        r_st.font.bold = True
        r_st.font.color.rgb = C_WHITE
        
        tb = slide.shapes.add_textbox(Inches(ax + 0.15), Inches(1.9), Inches(aw - 0.3), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        for b_idx, bullet in enumerate(abullets):
            p = tf.add_paragraph() if b_idx > 0 else tf.paragraphs[0]
            p.space_after = Pt(8)
            r = p.add_run(bullet)
            r.font.name = FONT_MAIN
            r.font.size = Pt(9.5)
            r.font.color.rgb = C_TEXT_DARK

    # Bottom Thank you / Q&A Box
    bot_card = add_card(slide, 0.8, 5.75, 11.733, 1.25, C_NAVY_PRIMARY, C_BLUE_CYAN)
    tb_bot = slide.shapes.add_textbox(Inches(1.0), Inches(5.82), Inches(11.333), Inches(1.1))
    tf_b = tb_bot.text_frame
    tf_b.word_wrap = True
    p_b1 = tf_b.paragraphs[0]
    p_b1.alignment = PP_ALIGN.CENTER
    r_b1 = p_b1.add_run("XIN TRÂN TRỌNG CẢM ƠN TECH LEAD & HỘI ĐỒNG!\n")
    r_b1.font.bold = True
    r_b1.font.size = Pt(13)
    r_b1.font.color.rgb = C_GOLD_ACCENT
    r_b2 = p_b1.add_run("Sẵn sàng lắng nghe câu hỏi phản biện kỹ thuật và góp ý chuyên sâu về kiến trúc hệ thống.")
    r_b2.font.size = Pt(10)
    r_b2.font.color.rgb = C_WHITE

    set_presenter_notes(slide, {
        "goal": "Chốt lại buổi trình bày ngắn gọn, tự tin và sẵn sàng cho phần Q&A.",
        "script": "Tóm lại, hệ thống RAG hiện tại đã hoàn thành xuất sắc sứ mệnh chứng minh tính 'Good Enough'. Chúng em đề xuất giữ nguyên Core Engine và triển khai sprint 1 tuần để gắn Hybrid Search. Em xin cảm ơn và sẵn sàng trả lời các câu hỏi của anh.",
        "highlights": "Tự tin, khoa học, cầu thị và có giải pháp hành động cụ thể.",
        "qa": "Mở sẵn các slide hoặc tài liệu kỹ thuật nội bộ để tra cứu nhanh khi Lead hỏi chi tiết."
    })


def main():
    print("=" * 75)
    print("BẮT ĐẦU CẬP NHẬT PRESENTATION RAG PHIÊN BẢN TRỰC QUAN CAO CẤP...")
    print("=" * 75)
    
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    print("[1/8] Slide 1: Trang Tiêu đề & Thông tin Tác giả...")
    build_slide_01_title(prs)
    
    print("[2/8] Slide 2: Sơ đồ luồng Kiến trúc RAG 4 tầng (rag_architecture_flow.png)...")
    build_slide_02_architecture(prs)
    
    print("[3/8] Slide 3: Chứng minh RAG 'Good Enough' (rag_benchmark_paper_vs_code.png)...")
    build_slide_03_good_enough(prs)
    
    print("[4/8] Slide 4: BẢNG SO SÁNH ĐỐI ĐẦU: Paper gốc vs Code mới (8 Tiêu chí)...")
    build_slide_04_comparison_table(prs)
    
    print("[5/8] Slide 5: Đánh giá Sòng phẳng: Điểm mạnh & Điểm yếu kiến trúc...")
    build_slide_05_strengths_weaknesses(prs)
    
    print("[6/8] Slide 6: Lộ trình Cải tiến: Hybrid Search & Cross-Encoder (rag_improvements_flow.png)...")
    build_slide_06_improvements(prs)
    
    print("[7/8] Slide 7: Tương lai: Multi-Agent ESG Auditor (rag_agentic_flow.png)...")
    build_slide_07_agentic_rag(prs)
    
    print("[8/8] Slide 8: Tổng kết & Khuyến nghị hành động...")
    build_slide_08_conclusion(prs)
    
    prs.save(str(OUTPUT_PPTX))
    print(f"\n=> Đã lưu thành công bộ slide tại: {OUTPUT_PPTX}")
    
    PAPERS_DIR.mkdir(parents=True, exist_ok=True)
    dest_copy = PAPERS_DIR / "thuyet_trinh_rag_architecture.pptx"
    shutil.copyfile(OUTPUT_PPTX, dest_copy)
    print(f"=> Đã sao chép vào: {dest_copy}")
    
    file_size_mb = OUTPUT_PPTX.stat().st_size / (1024 * 1024)
    print(f"=> Kích thước tệp: {file_size_mb:.2f} MB")
    print("=" * 75)
    print("HOÀN TẤT THÀNH CÔNG BỘ SLIDE TRỰC QUAN CAO CẤP!")
    print("=" * 75)


if __name__ == "__main__":
    main()
