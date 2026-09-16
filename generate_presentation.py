"""Tạo bộ slide thuyết trình PowerPoint (.pptx) chuẩn học thuật 16:9 báo cáo Hội đồng.
Chủ đề: Xử lý ngôn ngữ tự nhiên đa ngữ trong phân tích SDG và Cảm xúc Báo cáo Phát triển Bền vững của Doanh nghiệp Việt Nam.
Tác giả: Lê Đan Sơn, Dương Thị Hoàn (2026).
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.text.text import _Paragraph
_orig_add_run = _Paragraph.add_run
def _patched_add_run(self, text=None):
    r = _orig_add_run(self)
    if text is not None:
        r.text = str(text)
    return r
_Paragraph.add_run = _patched_add_run


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Code\paper_sdg")
FIGURES_DIR = ROOT / "figures"
PAPERS_DIR = ROOT / "papers"
OUTPUT_PPTX = ROOT / "bao_cao_nghien_cuu_sdg_vietnam.pptx"

# ==============================================================================
# BẢNG MÀU CHUẨN HỌC THUẬT (ACADEMIC SLATE & NAVY PALETTE)
# ==============================================================================
C_NAVY_DARK    = RGBColor(12, 30, 54)     # #0C1E36 - Nền tiêu đề, banner chính
C_NAVY_PRIMARY = RGBColor(27, 54, 93)     # #1B365D - Tiêu đề chính, khung chính
C_BLUE_ACCENT  = RGBColor(31, 119, 180)   # #1F77B4 - Màu nhấn thứ cấp
C_GOLD_ACCENT  = RGBColor(212, 175, 55)   # #D4AF37 - Điểm nhấn sang trọng
C_BG_LIGHT     = RGBColor(245, 247, 250)  # #F5F7FA - Nền slide nội dung
C_CARD_BG      = RGBColor(255, 255, 255)  # #FFFFFF - Nền card trắng
C_BORDER_LIGHT = RGBColor(218, 224, 233)  # #DAE0E9 - Viền card
C_TEXT_DARK    = RGBColor(33, 37, 41)     # #212529 - Chữ chính
C_TEXT_MUTED   = RGBColor(108, 117, 125)  # #6C757D - Chữ phụ / chú thích
C_GREEN_EMERALD= RGBColor(40, 140, 60)    # #288C3C - Màu tích cực / sinh thái
C_RED_ACCENT   = RGBColor(214, 39, 40)    # #D62728 - Màu cảnh báo / tiêu cực
C_WHITE        = RGBColor(255, 255, 255)

FONT_MAIN = "Segoe UI"
FONT_HEADING = "Segoe UI"


def set_shape_flat(shape, fill_color: RGBColor, line_color: RGBColor | None = None, line_width: float = 1.0):
    """Cài đặt màu nền và đường viền phẳng cho shape."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()


def add_slide_header(slide, title_text: str, category_tag: str = "BÁO CÁO KHOA HỌC", slide_num: int = 1):
    """Tạo thanh tiêu đề chuẩn mực ở đầu mỗi slide nội dung."""
    # Nền slide màu sáng
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, C_BG_LIGHT)
    
    # Thanh top banner màu trắng sang trọng
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.15))
    set_shape_flat(top_bar, C_CARD_BG, C_BORDER_LIGHT, 0.75)
    
    # Dải màu trang trí phía trên cùng
    accent_strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.06))
    set_shape_flat(accent_strip, C_GOLD_ACCENT)
    
    # Tag danh mục (Pill badge)
    tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.18), Inches(2.6), Inches(0.28))
    set_shape_flat(tag_box, C_NAVY_PRIMARY)
    tf_tag = tag_box.text_frame
    tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_tag = tf_tag.paragraphs[0]
    p_tag.alignment = PP_ALIGN.CENTER
    run_tag = p_tag.add_run()
    run_tag.text = category_tag.upper()
    run_tag.font.name = FONT_MAIN
    run_tag.font.size = Pt(9.5)
    run_tag.font.bold = True
    run_tag.font.color.rgb = C_WHITE
    
    # Tiêu đề slide
    tx_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.48), Inches(10.5), Inches(0.6))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.name = FONT_HEADING
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = C_NAVY_PRIMARY

    # Số trang ở góc phải
    num_box = slide.shapes.add_textbox(Inches(11.8), Inches(0.35), Inches(1.0), Inches(0.5))
    tf_num = num_box.text_frame
    p_num = tf_num.paragraphs[0]
    p_num.alignment = PP_ALIGN.RIGHT
    r_num = p_num.add_run()
    r_num.text = f"{slide_num:02d} / 18"
    r_num.font.name = FONT_MAIN
    r_num.font.size = Pt(11)
    r_num.font.bold = True
    r_num.font.color.rgb = C_TEXT_MUTED
    
    # Footer trang trọng ở chân slide
    foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.15), Inches(11.8), Inches(0.3))
    tf_f = foot_box.text_frame
    p_f = tf_f.paragraphs[0]
    r_f = p_f.add_run()
    r_f.text = "Đề tài: NLP Đa ngữ trong Phân tích SDG & Cảm xúc Báo cáo PTBV tại Việt Nam | Tác giả: Lê Đan Sơn, Dương Thị Hoàn"
    r_f.font.name = FONT_MAIN
    r_f.font.size = Pt(9)
    r_f.font.color.rgb = C_TEXT_MUTED


def add_card(slide, left: float, top: float, width: float, height: float, bg_color: RGBColor = C_CARD_BG, border_color: RGBColor = C_BORDER_LIGHT):
    """Tạo khối Card nổi bật để chứa nội dung hoặc hình ảnh."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    set_shape_flat(card, bg_color, border_color, 1.0)
    return card


def set_presenter_notes(slide, notes_dict: dict):
    """Ghi chú thuyết minh chi tiết cho người trình bày (Presenter Notes)."""
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
        add_section("ĐIỂM NHẤN CẦN NHỚ", notes_dict["highlights"])
    if "qa" in notes_dict:
        add_section("DỰ ĐOÁN CÂU HỎI HỘI ĐỒNG & CÁCH TRẢ LỜI", notes_dict["qa"])


# ==============================================================================
# HÀM XÂY DỰNG TỪNG SLIDE TRONG TỔNG SỐ 18 SLIDE
# ==============================================================================

def build_slide_01_title(prs):
    """Slide 1: Trang Tiêu Đề Báo Cáo."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, C_NAVY_DARK)
    
    accent_top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    set_shape_flat(accent_top, C_GOLD_ACCENT)
    
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.7), Inches(11.733), Inches(6.1))
    set_shape_flat(card, RGBColor(18, 42, 74), RGBColor(40, 75, 120), 1.5)
    
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.3), Inches(1.1), Inches(4.5), Inches(0.38))
    set_shape_flat(badge, C_GOLD_ACCENT)
    p_b = badge.text_frame.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b = p_b.add_run()
    r_b.text = "BÁO CÁO KẾT QUẢ NGHIÊN CỨU KHOA HỌC"
    r_b.font.name = FONT_MAIN
    r_b.font.size = Pt(11)
    r_b.font.bold = True
    r_b.font.color.rgb = C_NAVY_DARK
    
    tb_title = slide.shapes.add_textbox(Inches(1.3), Inches(1.6), Inches(10.7), Inches(1.8))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = (
        "XỬ LÝ NGÔN NGỮ TỰ NHIÊN ĐA NGỮ TRONG PHÂN TÍCH SDG VÀ CẢM XÚC "
        "BÁO CÁO PHÁT TRIỂN BỀN VỮNG CỦA DOANH NGHIỆP:\n"
        "BẰNG CHỨNG THỰC NGHIỆM TỪ CÁC DOANH NGHIỆP VIỆT NAM"
    )
    r_t.font.name = FONT_HEADING
    r_t.font.size = Pt(21)
    r_t.font.bold = True
    r_t.font.color.rgb = C_WHITE

    tb_en = slide.shapes.add_textbox(Inches(1.3), Inches(3.45), Inches(10.7), Inches(0.8))
    tf_en = tb_en.text_frame
    tf_en.word_wrap = True
    p_en = tf_en.paragraphs[0]
    r_en = p_en.add_run()
    r_en.text = (
        "Multilingual NLP for SDG and Sentiment Analysis of Corporate Sustainability Reports: "
        "Evidence from Vietnamese Enterprises"
    )
    r_en.font.name = FONT_MAIN
    r_en.font.size = Pt(12.5)
    r_en.font.italic = True
    r_en.font.color.rgb = RGBColor(190, 210, 235)

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.3), Inches(4.35), Inches(10.7), Inches(0.03))
    set_shape_flat(line, C_GOLD_ACCENT)

    tb_frame = slide.shapes.add_textbox(Inches(1.3), Inches(4.5), Inches(6.0), Inches(1.8))
    tf_f = tb_frame.text_frame
    tf_f.word_wrap = True
    p_f1 = tf_f.paragraphs[0]
    r_f1 = p_f1.add_run()
    r_f1.text = "KHUNG PHƯƠNG PHÁP LUẬN KẾ THỪA:\n"
    r_f1.font.bold = True
    r_f1.font.size = Pt(10)
    r_f1.font.color.rgb = C_GOLD_ACCENT
    r_f2 = p_f1.add_run()
    r_f2.text = (
        "Mô hình Trí tuệ Nhân tạo & NLP của Kang & Kim (2022),\n"
        "Tạp chí Applied Sciences (MDPI), 12(11), 5614.\n"
        "Ứng dụng chuyển giao & mở rộng thực nghiệm trên thị trường Việt Nam."
    )
    r_f2.font.size = Pt(10)
    r_f2.font.color.rgb = RGBColor(220, 230, 245)

    tb_auth = slide.shapes.add_textbox(Inches(7.6), Inches(4.5), Inches(4.4), Inches(1.8))
    tf_a = tb_auth.text_frame
    tf_a.word_wrap = True
    p_a = tf_a.paragraphs[0]
    r_a1 = p_a.add_run()
    r_a1.text = "TÁC GIẢ THỰC HIỆN ĐỀ TÀI:\n"
    r_a1.font.bold = True
    r_a1.font.size = Pt(10)
    r_a1.font.color.rgb = C_GOLD_ACCENT
    r_a2 = p_a.add_run()
    r_a2.text = "• Lê Đan Sơn\n• Dương Thị Hoàn\n\nNăm thực hiện: 2026"
    r_a2.font.size = Pt(11.5)
    r_a2.font.bold = True
    r_a2.font.color.rgb = C_WHITE

    set_presenter_notes(slide, {
        "goal": "Giới thiệu đề tài trang trọng, nêu bật tính cấp thiết, định vị rõ đây là một nghiên cứu thực nghiệm kế thừa bài báo quốc tế của Kang & Kim (2022) và mở rộng cho Việt Nam.",
        "script": "Kính thưa Quý Thầy Cô trong Hội đồng và toàn thể quý vị, hôm nay nhóm nghiên cứu gồm hai tác giả Lê Đan Sơn và Dương Thị Hoàn xin trân trọng báo cáo kết quả đề tài: 'Xử lý ngôn ngữ tự nhiên đa ngữ trong phân tích SDG và Cảm xúc Báo cáo Phát triển Bền vững của Doanh nghiệp: Bằng chứng thực nghiệm từ các doanh nghiệp Việt Nam'. Đề tài của chúng em bắt nguồn từ việc kế thừa trực tiếp khung phương pháp luận xử lý văn bản tiên phong của Kang và Kim (2022) trên tạp chí Applied Sciences, từ đó tinh chỉnh và chuyển giao vào bối cảnh thực tế của các doanh nghiệp niêm yết hàng đầu tại thị trường Việt Nam giai đoạn 2020–2025.",
        "highlights": "Nhớ nhấn mạnh tên 2 tác giả (Lê Đan Sơn, Dương Thị Hoàn) và nguồn gốc học thuật bài báo gốc Kang & Kim (2022).",
        "qa": "Nếu thầy cô hỏi 'Tại sao các em chọn đề tài này?': Trả lời: Báo cáo phát triển bền vững đang bùng nổ tại Việt Nam sau cam kết Net Zero COP26 và Thông tư 96 của Bộ Tài chính, nhưng việc đọc và giám sát thủ công hàng trăm trang báo cáo gặp nhiều hạn chế. Việc áp dụng công nghệ AI/NLP giúp tự động hóa, định lượng khách quan và mở ra hướng tiếp cận kiểm toán phi tài chính hiện đại."
    })


def build_slide_02_context(prs):
    """Slide 2: Bối Cảnh Nghiên Cứu & Động Lực Thể Chế."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BỐI CẢNH NGHIÊN CỨU & ĐỘNG LỰC THỰC TIỄN", "TỔNG QUAN VẤN ĐỀ", 2)
    
    w = 3.65
    h = 5.4
    top = 1.45
    
    add_card(slide, 0.8, top, w, h)
    tb1 = slide.shapes.add_textbox(Inches(0.95), Inches(top + 0.15), Inches(w - 0.3), Inches(h - 0.3))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.add_run("1. ĐỘNG LỰC THỂ CHẾ HÓA ESG\n\n").font.bold = True
    p1.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p1.runs[0].font.size = Pt(13)
    
    body1 = (
        "• Cam kết COP26 & Net Zero 2050: Thủ tướng Chính phủ cam kết đưa phát thải ròng về 0 vào năm 2050, thúc đẩy tái cơ cấu nền kinh tế xanh.\n\n"
        "• Thông tư 96/2020/TT-BTC: Bộ Tài chính chính thức bắt buộc các công ty niêm yết phải công bố báo cáo tác động môi trường và xã hội trong Báo cáo Thường niên hoặc Báo cáo Bền vững độc lập.\n\n"
        "• Khung hướng dẫn CSI & UBCKNN: Bộ Chỉ số Doanh nghiệp Bền vững (CSI) và Sổ tay thực hành ESG định hình tiêu chuẩn công bố."
    )
    p_b1 = tf1.add_paragraph()
    p_b1.text = body1
    p_b1.font.size = Pt(10.5)
    p_b1.font.color.rgb = C_TEXT_DARK

    add_card(slide, 4.84, top, w, h)
    tb2 = slide.shapes.add_textbox(Inches(5.0), Inches(top + 0.15), Inches(w - 0.3), Inches(h - 0.3))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.add_run("2. NGHỊCH LÝ & THÁCH THỨC\n\n").font.bold = True
    p2.runs[0].font.color.rgb = C_RED_ACCENT
    p2.runs[0].font.size = Pt(13)
    
    body2 = (
        "• Bùng nổ dung lượng văn bản: Báo cáo ngày càng dày (trung bình 100–200 trang/báo cáo), chứa đầy ngôn ngữ tự do phi cấu trúc.\n\n"
        "• Quá tải giám sát thủ công: Chi phí rà soát bằng con người quá lớn; kiểm toán viên và cơ quan quản lý không đủ nguồn lực kiểm tra định tính từng câu chữ.\n\n"
        "• Nguy cơ Quản trị Ấn tượng (Impression Management): Doanh nghiệp có xu hướng tô hồng thành tích, che giấu rủi ro và các sự cố vi phạm."
    )
    p_b2 = tf2.add_paragraph()
    p_b2.text = body2
    p_b2.font.size = Pt(10.5)
    p_b2.font.color.rgb = C_TEXT_DARK

    add_card(slide, 8.88, top, w, h)
    tb3 = slide.shapes.add_textbox(Inches(9.05), Inches(top + 0.15), Inches(w - 0.3), Inches(h - 0.3))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.add_run("3. ĐỘT PHÁ CÔNG NGHỆ NLP\n\n").font.bold = True
    p3.runs[0].font.color.rgb = C_GREEN_EMERALD
    p3.runs[0].font.size = Pt(13)
    
    body3 = (
        "• Định lượng hóa tự động: Khả năng đọc hiểu hàng trăm nghìn câu văn bản trong vài phút mà không phụ thuộc vào cảm tính con người.\n\n"
        "• Đo lường chuẩn hóa 17 SDGs: Ánh xạ ngữ nghĩa từng câu báo cáo vào 17 Mục tiêu Phát triển Bền vững của LHQ.\n\n"
        "• Nhận diện thiên lệch cảm xúc: Đo lường sắc thái (Pos/Neg) để phát hiện sớm các tín hiệu làm đẹp báo cáo và quản trị ấn tượng.\n\n"
        "• Khả năng mở rộng (Scalability): Tạo tiền đề cho hệ thống kiểm toán phi tài chính tự động trong tương lai."
    )
    p_b3 = tf3.add_paragraph()
    p_b3.text = body3
    p_b3.font.size = Pt(10.5)
    p_b3.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Làm rõ 3 trụ cột: (1) Thể chế bắt buộc, (2) Nghịch lý thông tin dày nhưng khó kiểm chứng, (3) Nhu cầu cấp bách của công cụ NLP.",
        "script": "Kính thưa Hội đồng, trước tiên em xin trình bày bối cảnh nghiên cứu. Tại Việt Nam, sau Hội nghị COP26 với cam kết Net Zero 2050 và đặc biệt là Thông tư 96/2020 của Bộ Tài chính, việc công bố thông tin môi trường - xã hội không còn là việc thiện nguyện tự chọn mà đã trở thành nghĩa vụ bắt buộc. Tuy nhiên, một nghịch lý lớn nảy sinh: các báo cáo ngày càng dày hàng trăm trang với ngôn từ tự do, gây quá tải cho các chuyên viên phân tích và cơ quan quản lý. Doanh nghiệp dễ dàng lồng ghép các chiến lược quản trị ấn tượng, chỉ nói điều hay mà giấu đi rủi ro. Vì vậy, việc áp dụng công nghệ Xử lý ngôn ngữ tự nhiên (NLP) chính là chìa khóa để định lượng hóa, tự động hóa và minh bạch hóa thông tin phi tài chính.",
        "highlights": "Nhớ 2 mốc pháp lý then chốt: Cam kết Net Zero COP26 và Thông tư số 96/2020/TT-BTC.",
        "qa": "Thầy cô có thể hỏi: 'Thông tư 96 có bắt buộc kiểm toán báo cáo phát triển bền vững không?': Trả lời: Hiện Thông tư 96 mới chỉ bắt buộc công bố thông tin chứ chưa bắt buộc kiểm toán độc lập đối với báo cáo phi tài chính. Chính khoảng trống kiểm toán này khiến nguy cơ tô hồng số liệu gia tăng, và nghiên cứu của chúng em nhằm cung cấp công cụ lấp đầy khoảng trống đó."
    })


def build_slide_03_original_paper(prs):
    """Slide 3: Bài Báo Gốc Kang & Kim (2022) & Khoảng Trống Nghiên Cứu."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BÀI BÁO GỐC KANG & KIM (2022) & SỰ THÍCH ỨNG TẠI VIỆT NAM", "TỔNG QUAN HỌC THUẬT", 3)
    
    add_card(slide, 0.8, 1.45, 5.7, 5.4)
    tb_left = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(5.3), Inches(5.0))
    tf_l = tb_left.text_frame
    tf_l.word_wrap = True
    p_l = tf_l.paragraphs[0]
    p_l.add_run("BÀI BÁO GỐC: KANG & KIM (2022)\n").font.bold = True
    p_l.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p_l.runs[0].font.size = Pt(13)
    
    p_sub = tf_l.add_paragraph()
    p_sub.text = "Applied Sciences (MDPI), 12(11), 5614 | SCIE / Scopus Q2"
    p_sub.font.size = Pt(10)
    p_sub.font.italic = True
    p_sub.font.color.rgb = C_TEXT_MUTED
    
    tf_l.add_paragraph().text = ""
    p_body_l = tf_l.add_paragraph()
    p_body_l.text = (
        "• Phạm vi nghiên cứu gốc:\n"
        "  - Thu thập báo cáo phát triển bền vững của các tập đoàn đa quốc gia toàn cầu.\n"
        "  - Toàn bộ ngữ liệu sử dụng ngôn ngữ tiếng Anh.\n\n"
        "• Khung phương pháp luận tiên phong:\n"
        "  - Sử dụng Sentence-BERT (all-MiniLM-L6-v2) để nhúng câu văn bản.\n"
        "  - Tính độ tương đồng Cosine giữa câu báo cáo với 17 mục tiêu SDG của Liên Hợp Quốc.\n"
        "  - Gom 17 SDGs thành 6 nhóm danh mục nhu cầu con người (Max-Neef, 1991).\n"
        "  - Dùng mô hình DistilBERT phân tích cảm xúc nhị phân 2 lớp (Positive / Negative).\n\n"
        "• Hạn chế của bài gốc:\n"
        "  - Chưa từng được kiểm nghiệm trên các ngôn ngữ thứ hai ngoài tiếng Anh.\n"
        "  - Phân loại cảm xúc 2 lớp bỏ qua các câu trung tính kỹ thuật."
    )
    p_body_l.font.size = Pt(10.2)
    p_body_l.font.color.rgb = C_TEXT_DARK

    add_card(slide, 6.83, 1.45, 5.7, 5.4)
    tb_right = slide.shapes.add_textbox(Inches(7.05), Inches(1.6), Inches(5.3), Inches(5.0))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    p_r = tf_r.paragraphs[0]
    p_r.add_run("SỰ THÍCH ỨNG & MỞ RỘNG TẠI VIỆT NAM\n").font.bold = True
    p_r.runs[0].font.color.rgb = C_GREEN_EMERALD
    p_r.runs[0].font.size = Pt(13)
    
    p_sub2 = tf_r.add_paragraph()
    p_sub2.text = "Nghiên cứu của Lê Đan Sơn & Dương Thị Hoàn (2026)"
    p_sub2.font.size = Pt(10)
    p_sub2.font.italic = True
    p_sub2.font.color.rgb = C_TEXT_MUTED

    tf_r.add_paragraph().text = ""
    p_body_r = tf_r.add_paragraph()
    p_body_r.text = (
        "• Chuyển giao mô hình sang ngôn ngữ tiếng Việt:\n"
        "  - Sử dụng vietnamese-sbert (768 chiều) đặc thù cho cấu trúc đơn âm tiết tiếng Việt.\n"
        "  - Xây dựng bộ ngữ liệu đối sánh chuẩn song ngữ gồm 169 mục tiêu cụ thể của LHQ (~400 câu chuẩn).\n\n"
        "• Bổ sung tầng nhận dạng quang học OCR:\n"
        "  - Phục hồi dữ liệu từ các báo cáo dạng scan hình ảnh (như PNJ 2022) vốn làm tê liệt các bộ đọc PDF thông thường.\n\n"
        "• Nâng cấp mô hình Cảm xúc 3 Lớp (PhoBERT):\n"
        "  - Thay vì phân loại nhị phân gượng ép, nghiên cứu sử dụng PhoBERT 3 lớp (Tích cực, Trung tính, Tiêu cực).\n"
        "  - Giữ nguyên các câu số liệu kỹ thuật vào lớp Trung tính để phản ánh chân thực văn phong báo cáo.\n\n"
        "• Phân tích sâu sắc bản chất ngành tại Việt Nam:\n"
        "  - Giải mã nghịch lý: Doanh nghiệp 'Nói nhiều về gì' vs 'Ít nói về gì'."
    )
    p_body_r.font.size = Pt(10.2)
    p_body_r.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Làm nổi bật giá trị kế thừa bài báo gốc và khẳng định rõ những đóng góp cải tiến mới khi đưa về thị trường Việt Nam.",
        "script": "Kính thưa các thầy cô, về mặt học thuật, đề tài của chúng em đứng trên vai người khổng lồ là công trình của Kang và Kim (2022) trên tạp chí Applied Sciences. Họ là những người đầu tiên kết hợp Sentence-BERT và DistilBERT để đo lường SDG và cảm xúc báo cáo bền vững. Tuy nhiên, mô hình gốc chỉ chạy trên tiếng Anh và các công ty toàn cầu. Khi đưa về Việt Nam, nhóm nghiên cứu đã thực hiện 4 sự nâng cấp quan trọng: Thứ nhất, chuyển sang kiến trúc SBERT tiếng Việt đa ngữ; thứ hai, dịch và chuẩn hóa bộ ngữ liệu 169 mục tiêu SDG của Liên Hợp Quốc sang tiếng Việt; thứ ba, bổ sung lớp OCR để đọc các file scan phức tạp; và thứ tư, nâng cấp mô hình cảm xúc từ 2 lớp nhị phân lên 3 lớp bằng PhoBERT để nhận diện chính xác các câu trung tính kỹ thuật.",
        "highlights": "Điểm khác biệt lớn nhất: PhoBERT 3 lớp (có lớp Trung tính) thay vì DistilBERT 2 lớp của bài báo gốc.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao việc thêm lớp Trung tính lại là một bước nâng cấp quan trọng?': Trả lời: Báo cáo phát triển bền vững chứa rất nhiều câu mô tả thông số kỹ thuật thuần túy (ví dụ: 'Năm qua công ty tiêu thụ 1,2 triệu kWh điện'). Nếu ép vào mô hình 2 lớp như bài gốc, máy tính buộc phải gán nhãn tích cực hoặc tiêu cực một cách gượng ép, làm méo mó bản chất văn bản."
    })


def build_slide_04_pipeline(prs):
    """Slide 4: Khung Phương Pháp Luận 5 Bước (Methodological Pipeline)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "QUY TRÌNH PHƯƠNG PHÁP LUẬN 5 BƯỚC KHÉP KÍN", "PHƯƠNG PHÁP NGHIÊN CỨU", 4)
    
    steps = [
        ("BƯỚC 1", "Thu Thập Dữ Liệu\n& Quản Lý PDF", 
         "• 42 Báo cáo độc lập & tích hợp (2020–2025).\n• 7 Tập đoàn niêm yết lớn trên HOSE & HNX.\n• Tải trực tiếp từ cổng công bố thông tin.", C_NAVY_PRIMARY),
        ("BƯỚC 2", "Tiền Xử Lý\n& Tesseract OCR", 
         "• PyMuPDF trích xuất văn bản từng trang.\n• Tách câu ngữ pháp tiếng Việt.\n• Lọc câu rác, độ dài < 6 từ.\n• OCR phục hồi các trang scan.", C_BLUE_ACCENT),
        ("BƯỚC 3", "Nhúng Câu Đa Ngữ\n& Đo Tương Đồng SDG", 
         "• vietnamese-sbert (768 chiều) cho tiếng Việt.\n• all-MiniLM-L6-v2 cho tiếng Anh.\n• Cosine similarity với 17 SDGs (~400 câu chuẩn LHQ).", C_NAVY_PRIMARY),
        ("BƯỚC 4", "Chuẩn Hóa Min-Max\n& Gom 6 Nhóm Nhu Cầu", 
         "• Chuẩn hóa toàn cục thang điểm 0–100:\n  Score = (sim - min)/(max - min) * 100.\n• Gom 17 SDGs về 6 nhóm nhu cầu (Max-Neef, 1991).", C_BLUE_ACCENT),
        ("BƯỚC 5", "Phân Tích Cảm Xúc\n& Tỷ Số Pos/Neg Ratio", 
         "• PhoBERT tinh chỉnh phân loại 3 lớp (Tích cực, Trung tính, Tiêu cực).\n• Đo lường tỷ số thiên lệch:\n  Pos/Neg Ratio = N_pos / N_neg.", C_GREEN_EMERALD)
    ]
    
    w = 2.22
    gap = 0.16
    top = 1.45
    h = 5.4
    
    for i, (b_name, b_title, b_desc, col) in enumerate(steps):
        x = 0.8 + i * (w + gap)
        card = add_card(slide, x, top, w, h)
        header_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(top), Inches(w), Inches(1.15))
        set_shape_flat(header_bar, col)
        
        tf_h = header_bar.text_frame
        tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_h1 = tf_h.paragraphs[0]
        p_h1.alignment = PP_ALIGN.CENTER
        r_h1 = p_h1.add_run()
        r_h1.text = b_name + "\n"
        r_h1.font.name = FONT_MAIN
        r_h1.font.size = Pt(10)
        r_h1.font.bold = True
        r_h1.font.color.rgb = C_GOLD_ACCENT
        
        r_h2 = p_h1.add_run()
        r_h2.text = b_title
        r_h2.font.name = FONT_HEADING
        r_h2.font.size = Pt(11)
        r_h2.font.bold = True
        r_h2.font.color.rgb = C_WHITE
        
        tb_body = slide.shapes.add_textbox(Inches(x + 0.08), Inches(top + 1.25), Inches(w - 0.16), Inches(h - 1.35))
        tf_b = tb_body.text_frame
        tf_b.word_wrap = True
        p_b = tf_b.paragraphs[0]
        p_b.text = b_desc
        p_b.font.name = FONT_MAIN
        p_b.font.size = Pt(9.8)
        p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Trình bày mạch lạc, khúc chiết quy trình kỹ thuật 5 bước từ lúc thu thập file PDF thô cho đến khi ra được các chỉ số định lượng cuối cùng.",
        "script": "Kính thưa Hội đồng, đây là sơ đồ kiến trúc tổng thể gồm 5 bước khép kín của đề tài: Bước 1 là thu thập 42 báo cáo chính thức. Bước 2 là bóc tách văn bản bằng PyMuPDF và áp dụng OCR để khôi phục các trang scan hình ảnh, phân tách câu và lọc câu rác. Bước 3 là đưa câu văn vào mô hình Transformer đa ngữ: dùng vietnamese-sbert để chuyển câu thành vector 768 chiều, so sánh độ tương đồng Cosine với tập ngữ liệu 17 SDGs của Liên Hợp Quốc. Bước 4 là chuẩn hóa toàn cục thang điểm 0–100 theo đúng công thức của Kang và Kim, sau đó quy nạp 17 SDGs thành 6 nhóm nhu cầu con người. Bước 5 là phân tích cảm xúc 3 lớp bằng PhoBERT và tính toán Tỷ số Pos/Neg để nhận diện mức độ thiên lệch lạc quan.",
        "highlights": "Nhớ giải thích rõ: Bước 4 áp dụng chuẩn hóa Min-Max TOÀN CỤC (global scaling) để bảo đảm điểm số giữa các doanh nghiệp và các năm có thể so sánh trực tiếp với nhau.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao lại dùng ngưỡng lọc 6 từ mà không phải 10 từ như bài báo gốc?': Trả lời: Trong bài báo gốc, tiếng Anh là ngôn ngữ đa âm tiết nên dùng ngưỡng 10 từ (words). Tuy nhiên, tiếng Việt là ngôn ngữ đơn âm tiết, một câu có 6 từ đơn âm đã có thể tạo thành một mệnh đề hoàn chỉnh (ví dụ: 'Công ty tiết kiệm điện năng'). Nếu áp ngưỡng 10 từ tiếng Anh, chúng ta sẽ vô tình loại bỏ rất nhiều câu có nghĩa trong tiếng Việt."
    })


def build_slide_05_sample(prs):
    """Slide 5: Mẫu Dữ Liệu Thực Nghiệm (Empirical Sample Overview)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "MẪU DỮ LIỆU THỰC NGHIỆM: 7 TẬP ĐOÀN NIÊM YẾT (2020–2025)", "DỮ LIỆU THỰC NGHIỆM", 5)
    
    kpis = [
        ("7", "DOANH NGHIỆP TRỤ CỘT", "Đại diện 7 ngành kinh tế then chốt", C_NAVY_PRIMARY),
        ("6 NĂM", "CHUỖI THỜI GIAN", "Giai đoạn bản lề 2020 – 2025", C_BLUE_ACCENT),
        ("42 BÁO CÁO", "TỔNG QUY MÔ DỮ LIỆU", "4.997 trang tài liệu PDF", C_NAVY_PRIMARY),
        ("96.461", "CÂU VĂN BẢN HỢP LỆ", "Mật độ bình quân 20,35 câu/trang", C_GREEN_EMERALD)
    ]
    
    w_kpi = 2.8
    gap_kpi = 0.17
    top_kpi = 1.45
    h_kpi = 1.15
    
    for i, (val, label, sub, col) in enumerate(kpis):
        x = 0.8 + i * (w_kpi + gap_kpi)
        card = add_card(slide, x, top_kpi, w_kpi, h_kpi)
        tb = slide.shapes.add_textbox(Inches(x), Inches(top_kpi + 0.08), Inches(w_kpi), Inches(h_kpi - 0.16))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = val + "\n"
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(20)
        r1.font.bold = True
        r1.font.color.rgb = col
        
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = label + "\n"
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(9.5)
        r2.font.bold = True
        r2.font.color.rgb = C_TEXT_DARK
        
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        r3 = p3.add_run()
        r3.text = sub
        r3.font.name = FONT_MAIN
        r3.font.size = Pt(8.5)
        r3.font.color.rgb = C_TEXT_MUTED

    table_card = add_card(slide, 0.8, 2.75, 11.733, 4.25)
    
    rows = 8
    cols = 6
    tbl_shape = slide.shapes.add_table(rows, cols, Inches(0.9), Inches(2.85), Inches(11.533), Inches(4.0))
    table = tbl_shape.table
    
    col_widths = [1.0, 3.2, 2.6, 1.0, 1.8, 1.933]
    for i, w in enumerate(col_widths):
        table.columns[i].width = Inches(w)
        
    headers = ["Mã CK", "Tên Tập đoàn / Doanh nghiệp", "Ngành hoạt động cốt lõi", "Sàn", "Số câu trích xuất", "Chuẩn mực áp dụng"]
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = h
        run.font.name = FONT_MAIN
        run.font.size = Pt(9.5)
        run.font.bold = True
        run.font.color.rgb = C_WHITE

    sample_data = [
        ["BVH", "Tập đoàn Bảo Việt (Bao Viet Holdings)", "Tài chính - Bảo hiểm nhân thọ & quỹ", "HOSE", "18.067 câu (845 trang)", "Báo cáo tích hợp <IIRC>, GRI"],
        ["PAN", "CTCP Tập đoàn PAN (The PAN Group)", "Nông nghiệp công nghệ cao, thủy sản", "HOSE", "14.154 câu (963 trang)", "GRI Standards chuỗi giá trị"],
        ["PLX", "Tập đoàn Xăng dầu Việt Nam (Petrolimex)", "Năng lượng, xăng dầu hạ nguồn", "HOSE", "3.398 câu (376 trang)", "GRI Standards, ISO 14064-1"],
        ["PNJ", "CTCP Vàng bạc Đá quý Phú Nhuận", "Chế tác kim hoàn & Bán lẻ trang sức", "HOSE", "9.774 câu (669 trang)", "GRI Standards, Trụ cột DE&I"],
        ["SSI", "CTCP Chứng khoán SSI", "Dịch vụ tài chính, ngân hàng đầu tư", "HOSE", "18.257 câu (889 trang)", "GRI Standards, Tài chính xanh"],
        ["VCS", "CTCP Vicostone (Tập đoàn Phenikaa)", "Sản xuất vật liệu đá thạch anh nhân tạo", "HNX", "13.918 câu (536 trang)", "GRI Standards, Kinh tế tuần hoàn"],
        ["VNM", "CTCP Sữa Việt Nam (Vinamilk)", "Chăn nuôi bò sữa & Chế biến thực phẩm", "HOSE", "18.893 câu (719 trang)", "GRI, PAS 2060 (Net Zero), CDP"]
    ]

    for i, row in enumerate(sample_data):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_CARD_BG if i % 2 == 0 else RGBColor(240, 244, 250)
            p = cell.text_frame.paragraphs[0]
            if j in [0, 3, 4]:
                p.alignment = PP_ALIGN.CENTER
            else:
                p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = val
            run.font.name = FONT_MAIN
            run.font.size = Pt(9.0)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = C_NAVY_PRIMARY
            else:
                run.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Chứng minh quy mô dữ liệu đồ sộ và tính đại diện toàn diện của mẫu nghiên cứu gồm 7 doanh nghiệp đầu ngành.",
        "script": "Kính thưa Hội đồng, mẫu nghiên cứu của đề tài gồm 7 tập đoàn niêm yết lớn trên HOSE và HNX trong giai đoạn 6 năm liên tục từ 2020 đến 2025, với tổng cộng 42 báo cáo chính thức. Tổng số trang tài liệu số hóa là gần 5.000 trang và chúng em đã bóc tách thành công 96.461 câu văn bản hợp lệ. Mẫu đại diện trọn vẹn cho 7 lĩnh vực then chốt: từ sản xuất nông nghiệp (PAN), chăn nuôi chế biến (Vinamilk), vật liệu công nghiệp (Vicostone), năng lượng (Petrolimex) cho đến bán lẻ (PNJ) và các định chế tài chính (Bảo Việt, SSI). Nhờ vậy, các kết luận rút ra mang tính đại diện rất cao cho bức tranh báo cáo phát triển bền vững tại Việt Nam.",
        "highlights": "96.461 câu văn bản - quy mô ngữ liệu thuộc hàng lớn nhất từng được xử lý trong các nghiên cứu NLP về ESG tại Việt Nam.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao chỉ chọn 7 công ty mà không phải toàn bộ VN30?': Trả lời: Thực tế tại Việt Nam, số lượng doanh nghiệp kiên trì công bố Báo cáo Phát triển Bền vững độc lập hoặc Báo cáo Tích hợp liên tục suốt 6 năm liên tiếp là rất hiếm. Đa số các doanh nghiệp khác chỉ công bố ngắt quãng hoặc chỉ viết 2-3 trang lồng ghép sơ sài. Do đó, 7 doanh nghiệp này là những đại diện tiên phong và chuẩn mực nhất cho chuỗi thời gian hoàn chỉnh."
    })


def build_slide_06_result1_similarity(prs):
    """Slide 6: Kết quả 1 - Phân phối Điểm Tương đồng SDG Toàn cục."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 1: PHÂN PHỐI ĐIỂM TƯƠNG ĐỒNG SDG TOÀN CỤC", "KẾT QUẢ THỰC NGHIỆM", 6)
    
    add_card(slide, 0.8, 1.45, 5.2, 5.4)
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(4.8), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p0 = tf.paragraphs[0]
    p0.add_run("ĐẶC TRƯNG PHÂN PHỐI GAUSSIAN\n").font.bold = True
    p0.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p0.runs[0].font.size = Pt(13)
    
    body = (
        "• Hình thái chuông đối xứng chuẩn mực:\n"
        "  - Điểm trung bình toàn cục: μ = 45,43 điểm.\n"
        "  - Độ lệch chuẩn: σ = 11,87 điểm.\n"
        "  - Điểm số trải dài từ 10 đến 90 trên thang đo 0–100.\n\n"
        "• Phân khúc bối cảnh chung (35–55 điểm):\n"
        "  - Chiếm hơn 75% toàn bộ ngữ liệu báo cáo.\n"
        "  - Đại diện cho các câu văn mô tả bối cảnh điều hành, quy chế tổ chức và các tuyên bố định hướng chung.\n\n"
        "• Đuôi phải chuyên sâu (>65 điểm - 8,5% câu):\n"
        "  - Nhóm các câu văn mang hàm lượng kỹ thuật cao.\n"
        "  - Mô tả trực tiếp các chỉ tiêu hành động cụ thể: công nghệ xử lý nước thải tuần hoàn, chứng nhận quốc tế, đầu tư dây chuyền giảm phát thải.\n\n"
        "• Ý nghĩa phương pháp luận:\n"
        "  - Khẳng định Sentence-BERT hoạt động cực kỳ ổn định trên tiếng Việt, tái lập hoàn hảo quy luật thống kê của Kang & Kim (2022)."
    )
    p_b = tf.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(10.2)
    p_b.font.color.rgb = C_TEXT_DARK

    add_card(slide, 6.2, 1.45, 6.333, 5.4)
    fig_path = FIGURES_DIR / "similarity_hist.png"
    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(6.35), Inches(1.65), width=Inches(6.033))
        
    tb_cap = slide.shapes.add_textbox(Inches(6.35), Inches(6.25), Inches(6.033), Inches(0.5))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run()
    r_cap.text = "Hình 1: Phân phối tần suất điểm tương đồng SDG của 96.461 câu (thang đo Min-Max 0–100)"
    r_cap.font.size = Pt(9.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    set_presenter_notes(slide, {
        "goal": "Chứng minh mô hình SBERT hoạt động ổn định và chính xác trên tiếng Việt, tạo ra phân phối chuẩn lý thuyết tương tự nghiên cứu gốc.",
        "script": "Kính thưa các thầy cô, tại Hình 1, chúng em biểu diễn phân phối điểm tương đồng SDG của toàn bộ 96.461 câu văn bản sau khi đã chuẩn hóa Min-Max toàn cục về thang điểm 0–100. Đồ thị cho thấy một phân phối dạng chuông đối xứng rất đẹp với điểm trung bình đạt 45,43 và độ lệch chuẩn 11,87. Phần lớn các câu văn nằm ở khoảng giữa từ 35 đến 55 điểm, phản ánh ngôn ngữ hành chính mô tả bối cảnh chung. Đặc biệt, có 8,5% số câu văn nằm ở đuôi bên phải đạt trên 65 điểm, đây chính là những câu văn chất lượng cao mô tả cụ thể các sáng kiến xanh, tiết kiệm năng lượng và giảm phát thải. Hình thái phân phối này hoàn toàn trùng khớp với kết quả mà Kang và Kim ghi nhận trên các công ty quốc tế, chứng minh Sentence-BERT tiếng Việt hoạt động cực kỳ đáng tin cậy.",
        "highlights": "Điểm trung bình μ = 45,43; Độ lệch chuẩn σ = 11,87. 8,5% câu chuyên biệt >65 điểm.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao điểm trung bình lại chỉ quanh mức 45 mà không phải 70 hay 80?': Trả lời: Bởi vì đây là điểm tương đồng ngữ nghĩa giữa từng câu báo cáo với 17 mục tiêu SDG của LHQ. Một báo cáo dù xuất sắc đến đâu thì vẫn phải có các câu chào mừng, giới thiệu cơ cấu tổ chức, lời mở đầu... nên mức trung bình 45 là hoàn toàn tự nhiên và phản ánh đúng thực tế khách quan."
    })


def build_slide_07_result2_heatmap(prs):
    """Slide 7: Kết quả 2 - Cấu trúc 6 Nhóm SDG Qua Heatmap."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 2: CẤU TRÚC 6 NHÓM SDG QUA BIỂU ĐỒ NHIỆT (HEATMAP)", "KẾT QUẢ THỰC NGHIỆM", 7)
    
    add_card(slide, 0.8, 1.45, 4.4, 5.4)
    fig_path = FIGURES_DIR / "heatmap_6cat.png"
    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(0.95), Inches(1.6), width=Inches(4.1), height=Inches(4.7))
        
    tb_cap = slide.shapes.add_textbox(Inches(0.95), Inches(6.35), Inches(4.1), Inches(0.45))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run()
    r_cap.text = "Hình 2: Heatmap 6 nhóm SDG của 7 DN (2020–2025)"
    r_cap.font.size = Pt(9.0)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    add_card(slide, 5.4, 1.45, 7.133, 5.4)
    tb_right = slide.shapes.add_textbox(Inches(5.6), Inches(1.6), Inches(6.7), Inches(5.1))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    p0.add_run("THỨ BẬC ƯU TIÊN TOÀN CỤC CỦA DOANH NGHIỆP VIỆT NAM\n").font.bold = True
    p0.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p0.runs[0].font.size = Pt(12.5)
    
    body = (
        "Từ biểu đồ nhiệt Heatmap, cấu trúc cam kết thể hiện rõ trật tự thứ bậc phân tầng:\n\n"
        "1. Kinh tế (Economic - SDG 8, 9) [48,0 – 52,5 điểm]:\n"
        "   - Sắc đỏ sẫm nhất xuyên suốt 42 báo cáo. Doanh nghiệp đặt ưu tiên hàng đầu vào tăng trưởng doanh thu, đổi mới hạ tầng công nghệ và đảm bảo việc làm.\n\n"
        "2. Xã hội (Social - SDG 11, 16, 17) [46,2 – 50,8 điểm]:\n"
        "   - Sắc đỏ cam đậm. Tập trung vào trách nhiệm cộng đồng, quản trị minh bạch và quan hệ đối tác phát triển bền vững.\n\n"
        "3. Tài nguyên (Resources - SDG 6, 7, 12, 14) [44,1 – 48,9 điểm]:\n"
        "   - Nổi trội ở khối sản xuất: năng lượng sạch, tiết kiệm nước và sản xuất có trách nhiệm.\n\n"
        "4. Đời sống (Life - SDG 1, 2, 3) [43,0 – 47,5 điểm]:\n"
        "   - An sinh xã hội, phúc lợi y tế và dinh dưỡng (đặc biệt cao tại Vinamilk và PAN).\n\n"
        "5. Môi trường (Environments - SDG 13, 15) [41,2 – 47,3 điểm]:\n"
        "   - Có sự bứt phá mạnh mẽ ở giai đoạn 2023–2025 nhờ chiến lược Net Zero.\n\n"
        "6. Công bằng (Equity - SDG 4, 5, 10) [39,0 – 44,2 điểm]:\n"
        "   - Luôn mang sắc vàng nhạt nhất (vùng trũng lớn nhất về bình đẳng giới lãnh đạo & phân tầng thu nhập)."
    )
    p_b = tf_r.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.8)
    p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Phân tích trật tự ưu tiên 6 nhóm SDG từ Heatmap và lý giải logic kinh tế đằng sau thứ bậc này.",
        "script": "Kính thưa Hội đồng, Hình 2 là biểu đồ nhiệt Heatmap tổng hợp điểm số 6 nhóm nhu cầu con người qua 42 báo cáo. Gam màu chuyển từ vàng nhạt (thấp) sang đỏ đậm (cao). Nhìn vào Heatmap, chúng ta thấy ngay một quy luật phân tầng rất nhất quán: Cột Kinh tế (Economic) luôn đỏ đậm nhất, theo sau là Xã hội (Social) và Tài nguyên (Resources). Ngược lại, cột Công bằng (Equity) luôn nhạt màu nhất ở tất cả các doanh nghiệp. Điều này phản ánh tư duy thực tế của các doanh nghiệp Việt Nam: Trọng tâm sống còn vẫn là tăng trưởng kinh doanh và việc làm, còn các vấn đề bình đẳng giới cấp cao hay khoảng cách giàu nghèo vẫn chưa nhận được sự quan tâm thỏa đáng.",
        "highlights": "Trật tự ưu tiên: Kinh tế > Xã hội > Tài nguyên > Đời sống > Môi trường > Công bằng.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao nhóm Công bằng (Equity) lại thấp nhất?': Trả lời: Nhóm Equity bao gồm SDG 4 (Giáo dục), SDG 5 (Bình đẳng giới) và SDG 10 (Giảm bất bình đẳng). Tại Việt Nam, doanh nghiệp thường chỉ báo cáo chung chung về số lượng lao động nữ mà rất ít khi đề cập đến bình đẳng lương bổng thực chất hay tỷ lệ lãnh đạo nữ cấp cao, dẫn đến độ tương đồng với SDG này bị thấp."
    })


def build_slide_08_result3_companies_p1(prs):
    """Slide 8: Đặc thù Ngành Chuyên Sâu (Khối Sản Xuất & Năng Lượng)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "ĐẶC THÙ NGÀNH: KHỐI SẢN XUẤT, NÔNG NGHIỆP & NĂNG LƯỢNG", "ĐẶC THÙ NGÀNH DOANH NGHIỆP", 8)
    
    w = 5.7
    h = 2.55
    top1 = 1.45
    top2 = 4.25
    left1 = 0.8
    left2 = 6.83
    
    add_card(slide, left1, top1, w, h)
    tb1 = slide.shapes.add_textbox(Inches(left1 + 0.15), Inches(top1 + 0.1), Inches(w - 0.3), Inches(h - 0.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.add_run("VINAMILK (VNM) — TIÊN PHONG NET ZERO\n").font.bold = True
    p1.runs[0].font.color.rgb = C_GREEN_EMERALD
    p1.runs[0].font.size = Pt(11.5)
    body1 = (
        "• Bứt phá nhóm Môi trường: Tăng mạnh nhất toàn mẫu (+6,06 điểm, từ 41,23 lên 47,29 điểm).\n"
        "• Đặc thù: 15 cụm trang trại bò sữa và 13 nhà máy đòi hỏi xử lý chất thải chăn nuôi và khí nhà kính.\n"
        "• Điểm sáng: Đạt chứng nhận trung hòa Carbon quốc tế PAS 2060 cho nhà máy Nghệ An; phát triển trang trại Green Farm tuần hoàn 100% nước, điện mặt trời áp mái."
    )
    p_b1 = tf1.add_paragraph()
    p_b1.text = body1
    p_b1.font.size = Pt(9.2)
    p_b1.font.color.rgb = C_TEXT_DARK

    add_card(slide, left2, top1, w, h)
    tb2 = slide.shapes.add_textbox(Inches(left2 + 0.15), Inches(top1 + 0.1), Inches(w - 0.3), Inches(h - 0.2))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.add_run("VICOSTONE (VCS) — KINH TẾ TUẦN HOÀN VẬT LIỆU\n").font.bold = True
    p2.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p2.runs[0].font.size = Pt(11.5)
    body2 = (
        "• Dẫn đầu Kinh tế (51,90) & Tài nguyên (50,03) cao nhất toàn mẫu.\n"
        "• Đặc thù: Chế tác đá thạch anh nhân tạo xuất khẩu đòi hỏi tiêu hao khoáng sản và hóa chất kết dính.\n"
        "• Điểm sáng: Công nghệ rung ép chân không Breton (Ý), tái chế 100% bùn thải đá thành phụ gia xi măng, hệ thống nước khép kín, chứng chỉ an toàn hóa chất Greenguard Gold."
    )
    p_b2 = tf2.add_paragraph()
    p_b2.text = body2
    p_b2.font.size = Pt(9.2)
    p_b2.font.color.rgb = C_TEXT_DARK

    add_card(slide, left1, top2, w, h)
    tb3 = slide.shapes.add_textbox(Inches(left1 + 0.15), Inches(top2 + 0.1), Inches(w - 0.3), Inches(h - 0.2))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.add_run("THE PAN GROUP (PAN) — NÔNG NGHIỆP XANH & LƯƠNG THỰC\n").font.bold = True
    p3.runs[0].font.color.rgb = C_BLUE_ACCENT
    p3.runs[0].font.size = Pt(11.5)
    body3 = (
        "• Trọng tâm Đời sống (Life: SDG 2, 3) & Tài nguyên (SDG 12) ổn định 45–47 điểm.\n"
        "• Đặc thù: Chuỗi giá trị nông nghiệp - thủy sản khép kín từ hạt giống (Vinaseed) đến tôm xuất khẩu (Fimex VN).\n"
        "• Điểm sáng: Mô hình cánh đồng lúa giảm phát thải carbon, nuôi tôm an toàn sinh học không kháng sinh, truy xuất nguồn gốc nông sản phục vụ xuất khẩu EU."
    )
    p_b3 = tf3.add_paragraph()
    p_b3.text = body3
    p_b3.font.size = Pt(9.2)
    p_b3.font.color.rgb = C_TEXT_DARK

    add_card(slide, left2, top2, w, h)
    tb4 = slide.shapes.add_textbox(Inches(left2 + 0.15), Inches(top2 + 0.1), Inches(w - 0.3), Inches(h - 0.2))
    tf4 = tb4.text_frame
    tf4.word_wrap = True
    p4 = tf4.paragraphs[0]
    p4.add_run("PETROLIMEX (PLX) — CHUYỂN DỊCH NĂNG LƯỢNG HẠ NGUỒN\n").font.bold = True
    p4.runs[0].font.color.rgb = C_RED_ACCENT
    p4.runs[0].font.size = Pt(11.5)
    body4 = (
        "• Trọng tâm Năng lượng sạch (SDG 7: 46,86) & Khí hậu (SDG 13: 46,41 điểm năm 2025).\n"
        "• Đặc thù: Đơn vị hạ nguồn năng lượng hóa thạch lớn nhất VN (>50% thị phần), chịu áp lực chuyển dịch xanh.\n"
        "• Điểm sáng: Phân phối nhiên liệu Euro 5 (DO 0,001S-V), phát triển mạng lưới xăng sinh học E5, điện mặt trời áp mái trạm xăng, kiểm kê khí nhà kính ISO 14064-1."
    )
    p_b4 = tf4.add_paragraph()
    p_b4.text = body4
    p_b4.font.size = Pt(9.2)
    p_b4.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Làm nổi bật phát hiện cốt lõi: Điểm số NLP không phải là con số ngẫu nhiên mà phản ánh rất chân thực mô hình kinh doanh kỹ thuật của từng công ty sản xuất.",
        "script": "Kính thưa các thầy cô, một trong những đóng góp thực nghiệm quan trọng nhất của đề tài là chứng minh mô hình NLP phản ánh cực kỳ nhạy bén bản chất kinh doanh cốt lõi của từng đơn vị: Vinamilk do sở hữu 15 trang trại bò sữa nên chịu áp lực xử lý phát thải nông nghiệp, do đó điểm Môi trường của họ tăng vọt hơn 6 điểm khi thực hiện cam kết Net Zero PAS 2060. Vicostone sản xuất đá thạch anh nhân tạo nên điểm Kinh tế và Tài nguyên cao nhất toàn mẫu, gắn liền với công nghệ tái chế 100% bùn thải đá. PAN Group làm nông nghiệp công nghệ cao nên nhóm Đời sống và Lương thực SDG 2 luôn dẫn đầu. Còn Petrolimex, dù là công ty phân phối xăng dầu, điểm số lại tập trung cao vào chuyển dịch nhiên liệu sạch Euro 5 và kiểm kê khí nhà kính.",
        "highlights": "Mỗi doanh nghiệp có một 'dấu chân SDG' riêng biệt hoàn toàn phù hợp với ngành nghề hoạt động.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao Petrolimex là công ty xăng dầu mà điểm môi trường lại không thấp?': Trả lời: Bởi vì báo cáo của Petrolimex dành phần lớn dung lượng để giải trình về các giải pháp giảm phát thải, cung cấp nhiên liệu sạch tiêu chuẩn Euro 5 và lộ trình năng lượng xanh để đáp ứng quy định nhà nước, do đó độ tương đồng với SDG 7 và SDG 13 là rất cao."
    })


def build_slide_09_result3_companies_p2(prs):
    """Slide 9: Đặc thù Ngành Chuyên Sâu (Khối Tài Chính & Bán Lẻ)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "ĐẶC THÙ NGÀNH: KHỐI TÀI CHÍNH, CHỨNG KHOÁN & BÁN LẺ", "ĐẶC THÙ NGÀNH DOANH NGHIỆP", 9)
    
    w = 3.65
    h = 5.4
    top = 1.45
    
    add_card(slide, 0.8, top, w, h)
    tb1 = slide.shapes.add_textbox(Inches(0.95), Inches(top + 0.15), Inches(w - 0.3), Inches(h - 0.3))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.add_run("TẬP ĐOÀN BẢO VIỆT (BVH)\nTIÊN PHONG BÁO CÁO TÍCH HỢP\n\n").font.bold = True
    p1.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p1.runs[0].font.size = Pt(12)
    body1 = (
        "• Chuẩn mực quốc tế <IIRC>:\n"
        "  - Tiên phong áp dụng khung Báo cáo Tích hợp kết nối 6 nguồn vốn (Tài chính, Sản xuất, Trí tuệ, Con người, Xã hội, Tự nhiên).\n\n"
        "• Trọng tâm Kinh tế & Xã hội:\n"
        "  - Điểm số nhóm Kinh tế và Xã hội luôn duy trì trên 49–50 điểm.\n"
        "  - Báo cáo tập trung vào quản trị rủi ro minh bạch và tuân thủ thể chế tài chính.\n\n"
        "• Đóng góp An sinh Xã hội:\n"
        "  - Mở rộng các gói bảo hiểm vi mô bảo vệ tài chính cho người thu nhập thấp và nông dân trước rủi ro thiên tai."
    )
    p_b1 = tf1.add_paragraph()
    p_b1.text = body1
    p_b1.font.size = Pt(9.8)
    p_b1.font.color.rgb = C_TEXT_DARK

    add_card(slide, 4.84, top, w, h)
    tb2 = slide.shapes.add_textbox(Inches(5.0), Inches(top + 0.15), Inches(w - 0.3), Inches(h - 0.3))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.add_run("CHỨNG KHOÁN SSI (SSI)\nKIẾN TẠO TÀI CHÍNH XANH\n\n").font.bold = True
    p2.runs[0].font.color.rgb = C_BLUE_ACCENT
    p2.runs[0].font.size = Pt(12)
    body2 = (
        "• Dẫn đầu nhóm Kinh tế & Quản trị:\n"
        "  - Điểm nhóm Kinh tế (SDG 8, 9) và Xã hội/Đối tác (SDG 16, 17) luôn đạt trên 50 điểm.\n\n"
        "• Vai trò Điều phối Dòng vốn:\n"
        "  - Do không có nhà máy sản xuất vật lý, phát thải trực tiếp rất nhỏ. Điểm nhấn là vai trò định chế tài chính xanh.\n\n"
        "• Sáng kiến Nổi bật:\n"
        "  - Tư vấn phát hành Trái phiếu Xanh (Green Bonds).\n"
        "  - Xây dựng khung thẩm định ESG trong đầu tư và tài trợ vốn bền vững cho thị trường chứng khoán VN."
    )
    p_b2 = tf2.add_paragraph()
    p_b2.text = body2
    p_b2.font.size = Pt(9.8)
    p_b2.font.color.rgb = C_TEXT_DARK

    add_card(slide, 8.88, top, w, h)
    tb3 = slide.shapes.add_textbox(Inches(9.05), Inches(top + 0.15), Inches(w - 0.3), Inches(h - 0.3))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.add_run("PHÚ NHUẬN JEWELRY (PNJ)\nĐIỂM SÁNG BÌNH ĐẲNG GIỚI DE&I\n\n").font.bold = True
    p3.runs[0].font.color.rgb = C_GOLD_ACCENT
    p3.runs[0].font.size = Pt(12)
    body3 = (
        "• Thế mạnh Xã hội & Đời sống:\n"
        "  - Điểm số nhóm Xã hội và Đời sống vượt trội nhờ đặc thù mạng lưới bán lẻ trang sức thời trang.\n\n"
        "• Tiên phong Trụ cột DE&I:\n"
        "  - Lực lượng lao động nữ chiếm tỷ trọng áp đảo (>60%).\n"
        "  - Chính sách đào tạo nghệ nhân kim hoàn nữ và phát triển nghề truyền thống.\n\n"
        "• Chăm sóc Nguồn nhân lực:\n"
        "  - Đẩy mạnh SDG 3 (Sức khỏe người lao động) và SDG 5 (Bình đẳng giới) thông qua các chương trình phúc lợi nội bộ toàn diện."
    )
    p_b3 = tf3.add_paragraph()
    p_b3.text = body3
    p_b3.font.size = Pt(9.8)
    p_b3.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Giải thích sự khác biệt giữa khối Tài chính - Bán lẻ và khối Sản xuất nặng.",
        "script": "Kính thưa các thầy cô, đối với nhóm ngành dịch vụ tài chính và bán lẻ, mô hình kinh doanh không vận hành các nhà máy khói bụi, do đó họ không tập trung vào xử lý nước thải hay chất thải rắn. Thay vào đó, Bảo Việt và SSI dẫn đầu tuyệt đối ở nhóm Kinh tế và Xã hội thông qua vai trò dẫn dắt dòng vốn xanh, phát hành trái phiếu xanh và bảo hiểm vi mô. Trong khi đó, PNJ là doanh nghiệp thời trang kim hoàn có hơn 60% lao động là nữ, nên điểm số của PNJ tỏa sáng ở trụ cột Bình đẳng giới, Đa dạng và Hòa nhập (DE&I) cùng việc bảo tồn văn hóa nghệ nhân.",
        "highlights": "Ngành tài chính đóng góp ESG bằng 'Dòng vốn xanh' (Green Finance), còn bán lẻ thời trang đóng góp bằng 'Bình đẳng giới & Nhân lực' (DE&I).",
        "qa": "Thầy cô có thể hỏi: 'Tại sao điểm môi trường của SSI hay BVH lại thấp hơn Vinamilk?': Trả lời: Đây chính là bằng chứng xác thực về tính khách quan của NLP. Các định chế tài chính phát thải Scope 1 và 2 rất thấp và họ không có dây chuyền xử lý rác thải công nghiệp, nên nếu điểm môi trường của họ mà cao bằng Vinamilk thì mô hình NLP mới là có vấn đề. Điểm số thấp ở môi trường là hoàn toàn chính xác với mô hình văn phòng tài chính."
    })


def build_slide_10_result4_trends(prs):
    """Slide 10: Kết quả 4 - Xu Hướng Dịch Chuyển Chuỗi Thời Gian 2020–2025."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 4: XU HƯỚNG DỊCH CHUYỂN CHUỖI THỜI GIAN (2020–2025)", "KẾT QUẢ THỰC NGHIỆM", 10)
    
    add_card(slide, 0.8, 1.45, 6.8, 5.4)
    fig_path = FIGURES_DIR / "slide_trends_grid.png"
    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(0.95), Inches(1.6), width=Inches(6.5), height=Inches(4.8))
        
    tb_cap = slide.shapes.add_textbox(Inches(0.95), Inches(6.45), Inches(6.5), Inches(0.35))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run()
    r_cap.text = "Hình 3: Xu hướng điểm số 6 nhóm SDG qua các năm (Trung bình toàn mẫu & 7 doanh nghiệp)"
    r_cap.font.size = Pt(8.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    add_card(slide, 7.8, 1.45, 4.733, 5.4)
    tb_r = slide.shapes.add_textbox(Inches(8.0), Inches(1.6), Inches(4.333), Inches(5.1))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    p0.add_run("BƯỚC NGOẶT CHÍNH SÁCH 2020–2025\n").font.bold = True
    p0.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p0.runs[0].font.size = Pt(12)
    
    body = (
        "Chuỗi thời gian ghi nhận 3 giai đoạn chuyển biến rõ rệt:\n\n"
        "• Giai đoạn 1: 2020–2021 (Giai đoạn Sơ khởi):\n"
        "  - Báo cáo mỏng, nội dung lồng ghép tối thiểu trong Báo cáo thường niên.\n"
        "  - Điểm số dao động phân tán, mức độ cam kết môi trường còn thấp (~41 điểm).\n\n"
        "• Giai đoạn 2: 2022–2023 (Bước ngoặt Thể chế):\n"
        "  - Thông tư 96/2020/TT-BTC có hiệu lực đầy đủ kết hợp cam kết Net Zero COP26.\n"
        "  - Số lượng câu văn tăng gấp 2–3 lần; độ bao phủ các mục tiêu tăng vọt.\n\n"
        "• Giai đoạn 3: 2024–2025 (Chuyên nghiệp hóa):\n"
        "  - Điểm số nhóm Môi trường và Tài nguyên tăng trưởng rõ rệt (+3 đến +6 điểm ở VNM, VCS, BVH).\n"
        "  - Doanh nghiệp đầu tư kiểm kê khí nhà kính độc lập và áp dụng các khung GRI/CDP chuẩn mực."
    )
    p_b = tf_r.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.6)
    p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Làm rõ tính năng động qua thời gian và chứng minh tác động tích cực của Thông tư 96 đối với chất lượng báo cáo doanh nghiệp.",
        "script": "Kính thưa Hội đồng, Hình 3 biểu diễn xu hướng biến động điểm số 6 nhóm SDG từ 2020 đến 2025 cho toàn mẫu và từng doanh nghiệp. Điểm cốt lõi rút ra là một xu thế đi lên rất rõ nét, chia làm 3 giai đoạn: Giai đoạn 2020–2021 là giai đoạn chuẩn bị sơ khởi. Bước ngoặt thực sự diễn ra vào năm 2022–2023 khi Thông tư 96 có hiệu lực bắt buộc, kéo theo dung lượng và chiều sâu báo cáo tăng vọt. Đến giai đoạn 2024–2025, các tập đoàn lớn như Vinamilk, Vicostone và Bảo Việt có bước tiến vượt bậc ở nhóm Môi trường và Tài nguyên, phản ánh việc doanh nghiệp đã chuyển từ nhận thức chung chung sang các cam kết định lượng có kiểm toán.",
        "highlights": "Thông tư 96/2020/TT-BTC là 'cú hích thể chế' làm thay đổi căn bản chất lượng công bố ESG tại Việt Nam.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao năm 2025 điểm số lại tăng đồng loạt ở hầu hết các nhóm?': Trả lời: Do năm 2025 các doanh nghiệp mở rộng quy mô báo cáo theo chuẩn quốc tế GRI mới, tích hợp thêm nhiều số liệu định lượng chi tiết về năng lượng tái tạo, kiểm kê khí nhà kính và kinh tế tuần hoàn, giúp mô hình NLP nhận diện được nhiều câu có điểm tương đồng cao hơn."
    })


def build_slide_11_result5_sentiment(prs):
    """Slide 11: Kết quả 5 - Phân Tích Cảm Xúc & Thiên Lệch Lạc Quan."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 5: SẮC THÁI CẢM XÚC & THIÊN LỆCH LẠC QUAN (POSITIVITY BIAS)", "KẾT QUẢ THỰC NGHIỆM", 11)
    
    add_card(slide, 0.8, 1.45, 6.2, 5.4)
    fig1 = FIGURES_DIR / "sentiment_hist.png"
    fig2 = FIGURES_DIR / "slide_sentiment_summary.png"
    
    if fig1.exists():
        slide.shapes.add_picture(str(fig1), Inches(0.95), Inches(1.55), width=Inches(5.9), height=Inches(2.2))
    if fig2.exists():
        slide.shapes.add_picture(str(fig2), Inches(0.95), Inches(3.9), width=Inches(5.9), height=Inches(2.4))

    tb_cap = slide.shapes.add_textbox(Inches(0.95), Inches(6.35), Inches(5.9), Inches(0.4))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run()
    r_cap.text = "Hình 4 & 5: Phân phối phân cực cảm xúc PhoBERT & Cơ cấu theo doanh nghiệp"
    r_cap.font.size = Pt(8.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    add_card(slide, 7.2, 1.45, 5.333, 5.4)
    tb_r = slide.shapes.add_textbox(Inches(7.4), Inches(1.6), Inches(4.933), Inches(5.1))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    p0.add_run("CƠ CẤU CẢM XÚC TOÀN MẪU (96.461 CÂU)\n").font.bold = True
    p0.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p0.runs[0].font.size = Pt(12)
    
    body = (
        "• Tỷ lệ phân bổ 3 lớp cảm xúc PhoBERT:\n"
        "  - Tích cực (Positive): 53,87% (51.966 câu) — Chiếm đa số áp đảo.\n"
        "  - Trung tính (Neutral): 32,87% (31.706 câu) — Các câu mô tả số liệu.\n"
        "  - Tiêu cực (Negative): 13,26% (12.789 câu) — Tỷ lệ rất khiêm tốn.\n\n"
        "• Đặc trưng phân phối 2 đỉnh (Bimodal Distribution):\n"
        "  - Đỉnh 1 (~0,50): Câu trung tính trần thuật số liệu vận hành.\n"
        "  - Đỉnh 2 (~0,95–1,00): Đỉnh rất lớn chứa các câu ca ngợi thành tựu.\n\n"
        "• Luận giải Lý thuyết Quản trị Ấn tượng (Impression Management):\n"
        "  - 'Hiệu ứng Pollyanna': Doanh nghiệp chủ động sử dụng văn phong lạc quan để xây dựng hình ảnh tính chính danh (legitimacy) trước cổ đông.\n"
        "  - Các rủi ro, sự cố thường bị né tránh hoặc làm mờ bằng cách diễn đạt giảm nhẹ."
    )
    p_b = tf_r.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.6)
    p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Trình bày các con số định lượng về sắc thái cảm xúc và kết nối với Lý thuyết Quản trị Ấn tượng (Impression Management).",
        "script": "Kính thưa các thầy cô, tại Hình 4 và 5, mô hình PhoBERT đã bóc tách sắc thái cảm xúc của toàn bộ 96.461 câu văn bản. Kết quả cho thấy một sự thiên lệch lạc quan mang tính cấu trúc: Câu Tích cực chiếm tới 53,87%, câu Trung tính chiếm 32,87%, trong khi câu Tiêu cực chỉ chiếm vỏn vẹn 13,26%. Phân phối điểm phân cực có dạng hai đỉnh (bimodal): một đỉnh trung tính mô tả số liệu và một đỉnh rất lớn ở vùng tích cực tuyệt đối (0,95–1,0). Dưới góc độ học thuật, đây là minh chứng thực nghiệm điển hình của Lý thuyết Quản trị Ấn tượng (Impression Management) và Hiệu ứng Pollyanna: doanh nghiệp dùng báo cáo bền vững như một công cụ tiếp thị hình ảnh, tập trung nói về thành tích và né tránh rủi ro.",
        "highlights": "Tỷ lệ: 53,87% Tích cực, 32,87% Trung tính, 13,26% Tiêu cực. Phân phối bimodal 2 đỉnh.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao câu Tiêu cực lại có tới 13,26%? Tiêu cực trong báo cáo thường là gì?': Trả lời: Các câu 'tiêu cực' trong báo cáo không phải là tự nhận mình xấu, mà là các câu mô tả bối cảnh khó khăn khách quan, ví dụ như: 'Đại dịch Covid-19 làm đứt gãy chuỗi cung ứng', 'Biến đổi khí hậu gây hạn mặn tại Đồng bằng Sông Cửu Long', hoặc 'Chi phí nguyên vật liệu đầu vào tăng cao'."
    })


def build_slide_12_result6_sentiment_ratio(prs):
    """Slide 12: Kết quả 6 - Diễn Biến Tỷ Số Cảm Xúc Pos/Neg Qua Thời Gian."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 6: TỶ SỐ CẢM XÚC POS/NEG RATIO & ĐỘ NHẠY BỐI CẢNH", "KẾT QUẢ THỰC NGHIỆM", 12)
    
    add_card(slide, 0.8, 1.45, 6.0, 5.4)
    fig_path = FIGURES_DIR / "sentiment_ratio.png"
    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(0.95), Inches(1.6), width=Inches(5.7), height=Inches(4.8))

    tb_cap = slide.shapes.add_textbox(Inches(0.95), Inches(6.45), Inches(5.7), Inches(0.35))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run()
    r_cap.text = "Hình 6: Diễn biến Tỷ số Pos/Neg Ratio theo năm giữa 7 doanh nghiệp"
    r_cap.font.size = Pt(8.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    add_card(slide, 7.0, 1.45, 5.533, 5.4)
    tb_r = slide.shapes.add_textbox(Inches(7.2), Inches(1.6), Inches(5.133), Inches(5.1))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    p0.add_run("TỶ SỐ POS/NEG BÌNH QUÂN ĐẠT 4,06 LẦN\n").font.bold = True
    p0.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p0.runs[0].font.size = Pt(12)
    
    body = (
        "• Ý nghĩa chỉ báo:\n"
        "  - Cứ mỗi câu đề cập đến rủi ro hoặc thách thức, doanh nghiệp sử dụng bình quân 4,06 câu để ca ngợi thành tích.\n\n"
        "• Vinamilk (VNM) — Ổn định ở mức cao (5,38 lần):\n"
        "  - Tỷ số Pos/Neg duy trì từ 4,5 đến 6,8 lần suốt 6 năm.\n"
        "  - Phản ánh phong cách truyền thông phát triển bền vững chuyên nghiệp, chau chuốt và định hướng thành tựu cao.\n\n"
        "• PNJ — Độ nhạy cảm xúc trước cú sốc Covid-19 (2022):\n"
        "  - Năm 2022: Tỷ số tụt xuống mức kỷ lục 1,21 lần (317 câu tích cực vs 261 câu tiêu cực), phản ánh chân thực khó khăn đóng cửa mạng lưới bán lẻ tại TP.HCM.\n"
        "  - Năm 2023: Tỷ số phục hồi mạnh mẽ lên 4,23 lần khi kinh doanh khởi sắc trở lại.\n\n"
        "• Kết luận:\n"
        "  - Tỷ số Pos/Neg là một 'nhiệt kế' nhạy bén đo lường mức độ trung thực và tính chân thực của văn phong báo cáo."
    )
    p_b = tf_r.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.6)
    p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Làm rõ ý nghĩa của chỉ số Tỷ số Pos/Neg và phân tích case study PNJ 2022 như một bằng chứng về độ nhạy của mô hình NLP.",
        "script": "Kính thưa Hội đồng, Hình 6 thể hiện diễn biến Tỷ số Pos/Neg, tức tỷ lệ giữa số câu tích cực chia cho số câu tiêu cực. Toàn mẫu đạt trung bình 4,06 lần, nghĩa là cứ 1 câu nhắc đến khó khăn thì có hơn 4 câu ca ngợi thành công. Vinamilk là doanh nghiệp có tỷ số ổn định và cao nhất với bình quân 5,38 lần. Tuy nhiên, case study thú vị nhất là PNJ năm 2022: Tỷ số Pos/Neg của PNJ tụt dốc xuống chỉ còn 1,21 lần do báo cáo dành rất nhiều câu mô tả tác động tiêu cực của các đợt giãn cách xã hội tại TP.HCM. Sang năm 2023, khi thị trường phục hồi, tỷ số này lập tức bật tăng trở lại 4,23 lần. Điều này chứng minh chỉ báo NLP của chúng em phản ánh cực kỳ trung thực và nhạy cảm với các biến cố thực tế của nền kinh tế.",
        "highlights": "Tỷ số toàn mẫu: 4,06 lần. PNJ năm 2022 tụt xuống 1,21 lần (Covid-19) và bật tăng lên 4,23 lần năm 2023.",
        "qa": "Thầy cô có thể hỏi: 'Tỷ số Pos/Neg cao có phải luôn luôn là tốt không?': Trả lời: Dạ không. Tỷ số Pos/Neg quá cao (ví dụ trên 8 hay 10 lần) thường là dấu hiệu cảnh báo của việc 'tô hồng' báo cáo quá đà (Pollyanna effect), thiếu tính giải trình trách nhiệm đối với các rủi ro thực tế. Một báo cáo chất lượng cao nên có sự cân bằng lành mạnh."
    })


def build_slide_13_discussion_talk_heavy(prs):
    """Slide 13: Thảo luận Chuyên sâu - Doanh nghiệp Việt Nam 'NÓI NHIỀU VỀ GÌ'?"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "THẢO LUẬN: DOANH NGHIỆP VIỆT NAM \"NÓI NHIỀU VỀ GÌ\"?", "THẢO LUẬN CHUYÊN SÂU", 13)
    
    w = 11.733
    h = 1.65
    top = 1.45
    gap = 0.22
    
    topics = [
        ("1. TĂNG TRƯỞNG KINH TẾ, VIỆC LÀM & ĐÓNG GÓP NGÂN SÁCH (SDG 8, 9)",
         "• Doanh nghiệp dành dung lượng trang lớn nhất để tường thuật chi tiết về doanh thu, lợi nhuận, số tiền nộp thuế nhà nước và mở rộng thị phần.\n"
         "• Tập trung nhấn mạnh việc đầu tư máy móc tự động hóa, mở rộng nhà máy và chính sách lương, thưởng, bảo hiểm cho người lao động.\n"
         "• Nguyên nhân: Đây là các chỉ tiêu cốt lõi phản ánh trực tiếp kết quả tài chính, dễ đo lường và thỏa mãn kỳ vọng ngắn hạn của cổ đông.",
         C_NAVY_PRIMARY),
        ("2. AN SINH XÃ HỘI & THIỆN NGUYỆN CSR BỀ NỔI (SDG 1, 17)",
         "• Các báo cáo dành hàng chục trang hình ảnh để kể về các chương trình tài trợ: xây cầu nông thôn, trao nhà tình thương, cấp học bổng học sinh nghèo, cứu trợ bão lũ.\n"
         "• Hoạt động trách nhiệm xã hội (CSR) tại Việt Nam vẫn mang đậm tính thiện nguyện truyền thống, bề nổi nhằm tạo dựng thiện cảm công chúng.\n"
         "• Rất hiếm khi doanh nghiệp tích hợp hoạt động xã hội vào chuỗi giá trị cốt lõi để tạo ra giá trị chia sẻ chung (Creating Shared Value - CSV).",
         C_BLUE_ACCENT),
        ("3. SÁNG KIẾN TIẾT KIỆM CHI PHÍ NỘI BỘ 'LỢI ÍCH KÉP' (SDG 6, 7, 12)",
         "• Doanh nghiệp rất hào hứng công bố các số liệu về: tiết kiệm điện chiếu sáng văn phòng, giảm sử dụng giấy in, thay bóng đèn LED, tái tuần hoàn nước làm mát.\n"
         "• Lý do lựa chọn: Đây là các sáng kiến môi trường mang lại 'lợi ích kép' tức thì — vừa giúp doanh nghiệp khoác lên chiếc áo 'xanh' trước công chúng, vừa giúp cắt giảm chi phí vận hành ngay lập tức.",
         C_GREEN_EMERALD)
    ]
    
    for i, (t_title, t_desc, col) in enumerate(topics):
        y = top + i * (h + gap)
        card = add_card(slide, 0.8, y, w, h)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(y), Inches(0.18), Inches(h))
        set_shape_flat(bar, col)
        
        tb = slide.shapes.add_textbox(Inches(1.15), Inches(y + 0.08), Inches(w - 0.45), Inches(h - 0.16))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p0 = tf.paragraphs[0]
        p0.add_run(t_title + "\n").font.bold = True
        p0.runs[0].font.color.rgb = col
        p0.runs[0].font.size = Pt(11)
        
        p1 = tf.add_paragraph()
        p1.text = t_desc
        p1.font.size = Pt(9.3)
        p1.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Làm rõ 3 chủ đề mà doanh nghiệp Việt Nam có xu hướng 'nói rất nhiều' và lý giải động cơ đằng sau hành vi này.",
        "script": "Kính thưa Hội đồng, từ việc phân tích chuyên sâu 96.461 câu văn, chúng em phát hiện ra một bức tranh tương phản rất rõ rệt về những gì doanh nghiệp 'nói nhiều' và 'ít nói'. Thứ nhất, doanh nghiệp nói rất nhiều về Tăng trưởng kinh tế, doanh thu và nộp ngân sách (SDG 8, 9). Thứ hai, doanh nghiệp dành hàng chục trang ảnh cho các hoạt động Từ thiện, xây cầu, tặng quà (SDG 1, 17) mang tính chất CSR bề nổi để tạo thiện cảm. Và thứ ba, doanh nghiệp rất tích cực khoe các sáng kiến tiết kiệm điện, nước, giấy in văn phòng (SDG 6, 7, 12). Động cơ ở đây rất rõ ràng: Đây là những chủ đề dễ làm, mang lại 'lợi ích kép' vừa cắt giảm chi phí trực tiếp, vừa giúp đánh bóng thương hiệu.",
        "highlights": "3 chủ đề nói nhiều: Kinh tế doanh thu (SDG 8,9), Từ thiện bề nổi (SDG 1,17), Tiết kiệm chi phí nội bộ (SDG 6,7,12).",
        "qa": "Thầy cô có thể hỏi: 'Tại sao từ thiện xã hội lại bị coi là CSR bề nổi?': Trả lời: Bởi vì các hoạt động thiện nguyện như tặng quà Tết hay trao học bổng tuy rất nhân văn nhưng chưa giải quyết được bài toán phát triển bền vững từ gốc rễ chuỗi cung ứng, chưa tạo ra giá trị chia sẻ chung (Creating Shared Value) gắn liền với mô hình kinh doanh."
    })


def build_slide_14_discussion_rarely_talk(prs):
    """Slide 14: Thảo luận Chuyên sâu - Doanh nghiệp Việt Nam 'ÍT NÓI VỀ GÌ'?"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "THẢO LUẬN: DOANH NGHIỆP VIỆT NAM \"ÍT NÓI VỀ GÌ\" (NÉ TRÁNH)?", "THẢO LUẬN CHUYÊN SÂU", 14)
    
    w = 5.7
    h = 2.55
    top1 = 1.45
    top2 = 4.25
    left1 = 0.8
    left2 = 6.83
    
    voids = [
        ("1. ĐA DẠNG SINH HỌC & HỆ SINH THÁI (SDG 14, 15)",
         "• 'Vùng trũng' lớn nhất toàn bộ 42 báo cáo khảo sát.\n"
         "• Hoàn toàn vắng bóng các số liệu đo lường cụ thể về tác động của khai thác tài nguyên và chuỗi cung ứng lên hệ sinh thái rừng, đất ngập nước hay biển.\n"
         "• Nội dung chỉ dừng ở mức các khẩu hiệu chung chung về bảo vệ môi trường hoặc phát động phong trào trồng vài cây xanh cảnh quan.",
         left1, top1, C_RED_ACCENT),
        ("2. BÌNH ĐẲNG GIỚI LÃNH ĐẠO & CHÊNH LỆCH LƯƠNG (SDG 5, 10)",
         "• Doanh nghiệp chỉ công bố tỷ lệ lao động nữ nói chung (công nhân trực tiếp), rất hiếm khi minh bạch tỷ lệ nữ trong HĐQT hay Ban Giám đốc.\n"
         "• Tuyệt đối né tránh công bố tỷ lệ chênh lệch thu nhập giữa ban điều hành và người lao động bình thường (CEO-to-worker pay ratio).\n"
         "• Bỏ qua các thống kê về khoảng cách tiền lương theo giới tính ở cùng cấp bậc.",
         left2, top1, C_RED_ACCENT),
        ("3. PHÁT THẢI CHUỖI CUNG ỨNG PHẠM VI 3 (SCOPE 3 GHG)",
         "• Doanh nghiệp mới chỉ bước đầu kiểm kê phát thải trực tiếp tại nhà máy (Scope 1) và tiêu thụ điện lưới (Scope 2).\n"
         "• Phát thải gián tiếp từ toàn bộ chuỗi cung ứng đầu vào và quá trình tiêu dùng đầu ra (Scope 3 — chiếm 70–80% tổng phát thải thực tế) gần như bị bỏ ngỏ.\n"
         "• Nguyên nhân do chuỗi cung ứng rời rạc và thiếu công cụ đo lường chuyên sâu.",
         left1, top2, C_RED_ACCENT),
        ("4. SỰ CỐ TIÊU CỰC, TRANH CHẤP & XỬ PHẠT (SDG 16)",
         "• Hiện tượng 'Gạn đục khơi trong' điển hình:\n"
         "  - Hầu như không có báo cáo nào ghi nhận các tai nạn lao động nghiêm trọng, khiếu nại của khách hàng, hay các sự cố rò rỉ xả thải.\n"
         "  - Các quyết định xử phạt vi phạm hành chính về thuế hoặc môi trường bị che giấu hoàn toàn.\n"
         "• Báo cáo trở thành tài liệu PR quảng bá thay vì công cụ giải trình rủi ro.",
         left2, top2, C_RED_ACCENT)
    ]
    
    for v_title, v_desc, x, y, col in voids:
        card = add_card(slide, x, y, w, h)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.15), Inches(h))
        set_shape_flat(bar, col)
        
        tb = slide.shapes.add_textbox(Inches(x + 0.25), Inches(y + 0.1), Inches(w - 0.4), Inches(h - 0.2))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p0 = tf.paragraphs[0]
        p0.add_run(v_title + "\n").font.bold = True
        p0.runs[0].font.color.rgb = col
        p0.runs[0].font.size = Pt(10.5)
        
        p1 = tf.add_paragraph()
        p1.text = v_desc
        p1.font.size = Pt(9.0)
        p1.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Chỉ ra 4 'vùng trũng' thông tin bị né tránh, khẳng định giá trị phản biện học thuật sắc bén của đề tài.",
        "script": "Kính thưa Quý Thầy Cô, mặt đối lập của bức tranh là 4 'vùng trũng' thông tin mà các doanh nghiệp Việt Nam gần như né tránh hoặc chỉ nói rất hời hợt: Thứ nhất là Đa dạng sinh học (SDG 14, 15) hoàn toàn không có số liệu định lượng. Thứ hai là Bình đẳng giới thực chất cấp lãnh đạo và Khoảng cách thu nhập (SDG 5, 10), doanh nghiệp tuyệt đối không công bố chênh lệch lương CEO với công nhân. Thứ ba là Phát thải Scope 3 chuỗi cung ứng - dù chiếm tới 80% dấu chân carbon thực tế nhưng đều bị bỏ qua. Và thứ tư là các sự cố tiêu cực hay xử phạt vi phạm hành chính bị giấu kín hoàn toàn. Đây chính là phát hiện có giá trị thực tiễn cao nhất của đề tài, gióng lên hồi chuông cảnh báo cho các cơ quan quản lý về tính minh bạch của báo cáo ESG.",
        "highlights": "4 vùng né tránh: Đa dạng sinh học, Bình đẳng giới lãnh đạo & chênh lệch lương, Scope 3 chuỗi cung ứng, Sự cố tiêu cực & vi phạm xử phạt.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao việc phát hiện các nội dung né tránh này lại quan trọng?': Trả lời: Bởi vì kiểm toán và giám sát không chỉ là kiểm tra xem doanh nghiệp nói những gì, mà quan trọng hơn là phải biết doanh nghiệp ĐANG NÉ TRÁNH điều gì. AI/NLP giúp chúng ta phát hiện ra những khoảng trống vô hình mà mắt thường khi đọc lướt khó nhận ra."
    })


def build_slide_15_comparison(prs):
    """Slide 15: Đóng Góp Học Thuật & So Sánh Đối Chuẩn Với Kang & Kim (2022)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "ĐÓNG GÓP HỌC THUẬT: SO SÁNH ĐỐI CHUẨN VỚI KANG & KIM (2022)", "ĐÓNG GÓP HỌC THUẬT", 15)
    
    table_card = add_card(slide, 0.8, 1.45, 11.733, 5.4)
    
    rows = 6
    cols = 3
    tbl_shape = slide.shapes.add_table(rows, cols, Inches(0.95), Inches(1.6), Inches(11.433), Inches(5.1))
    table = tbl_shape.table
    
    table.columns[0].width = Inches(2.3)
    table.columns[1].width = Inches(4.5)
    table.columns[2].width = Inches(4.633)
    
    headers = ["TIÊU CHÍ ĐỐI CHUẨN", "NGHIÊN CỨU GỐC: KANG & KIM (2022)", "NGHIÊN CỨU HIỆN TẠI TẠI VIỆT NAM (2026)"]
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = h
        run.font.name = FONT_MAIN
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = C_WHITE

    matrix = [
        ["1. Phạm vi ngôn ngữ & Đối tượng",
         "• Chỉ áp dụng trên văn bản tiếng Anh.\n• Khảo sát các tập đoàn đa quốc gia phát triển toàn cầu.",
         "• Tích hợp NLP đa ngữ (tiếng Việt & tiếng Anh).\n• Khảo sát chuyên sâu 7 tập đoàn niêm yết lớn tại Việt Nam."],
        ["2. Ngữ liệu mục tiêu SDG",
         "• Ngữ liệu tiếng Anh trích từ báo cáo Liên Hợp Quốc.",
         "• Xây dựng bộ ngữ liệu đối sánh chuẩn song ngữ 169 mục tiêu cụ thể của LHQ (~400 câu chuẩn tiếng Việt)."],
        ["3. Tiền xử lý & Khôi phục dữ liệu",
         "• Đọc PDF thông thường, bỏ qua các tài liệu lỗi/scan.",
         "• Tích hợp OCR (Tesseract vie+eng) phục hồi văn bản từ các trang PDF scan hình ảnh phức tạp (PNJ 2022)."],
        ["4. Mô hình Phân tích Cảm xúc",
         "• DistilBERT phân loại nhị phân 2 lớp (Pos / Neg).\n• Ép các câu số liệu kỹ thuật vào nhãn tích cực hoặc tiêu cực.",
         "• Nâng cấp mô hình PhoBERT 3 lớp (Tích cực, Trung tính, Tiêu cực).\n• Giữ nguyên 32,87% câu Trung tính kỹ thuật."],
        ["5. Chiều sâu phân tích đặc thù",
         "• Dừng lại ở phân phối thống kê chung toàn cầu.",
         "• Phân tích sâu sắc mối quan hệ giữa mô hình kinh doanh và hiện tượng 'Nói nhiều về gì' vs 'Ít nói về gì'."]
    ]

    for i, row in enumerate(matrix):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_CARD_BG if i % 2 == 0 else RGBColor(240, 244, 250)
            p = cell.text_frame.paragraphs[0]
            if j == 0:
                p.alignment = PP_ALIGN.CENTER
            else:
                p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = val
            run.font.name = FONT_MAIN
            run.font.size = Pt(9.0)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = C_NAVY_PRIMARY
            else:
                run.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Tóm kết các đóng góp học thuật và khẳng định vị thế nghiên cứu so với bài báo quốc tế.",
        "script": "Kính thưa các thầy cô, Bảng trên slide tóm tắt toàn bộ sự đối chuẩn giữa nghiên cứu của chúng em và công trình gốc của Kang và Kim (2022). Chúng em không chỉ đơn thuần là tái lập (replicate) mà đã mở rộng và giải quyết được 5 vấn đề cốt tử: Chuyển giao thành công sang tiếng Việt, dịch và chuẩn hóa ngữ liệu SDG song ngữ, giải quyết bài toán tệp scan bằng OCR, nâng cấp mô hình cảm xúc 3 lớp PhoBERT để bảo toàn câu số liệu trung tính, và cuối cùng là giải mã được đặc thù ngành sâu sắc tại Việt Nam. Đây là cơ sở dữ liệu thực nghiệm đầu tiên về văn bản SDG của doanh nghiệp niêm yết được phân tích tự động tại nước ta.",
        "highlights": "5 đóng góp cốt lõi: Đa ngữ tiếng Việt, Bộ ngữ liệu song ngữ, OCR tệp scan, PhoBERT 3 lớp, Giải mã đặc thù ngành VN.",
        "qa": "Thầy cô có thể hỏi: 'Đóng góp lớn nhất của các em là gì?': Trả lời: Dạ, đóng góp lớn nhất là chứng minh được tính khả thi và độ chính xác của quy trình NLP trong việc giám sát phi tài chính tự động tại một thị trường mới nổi như Việt Nam, mở ra phương pháp luận mới thay thế cho việc đọc thủ công truyền thống."
    })


def build_slide_16_implications(prs):
    """Slide 16: Hàm Ý Quản Trị & Đề Xuất Chính Sách."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "HÀM Ý THỰC TIỄN & ĐỀ XUẤT CHÍNH SÁCH CHO CÁC BÊN", "HÀM Ý QUẢN TRỊ", 16)
    
    w = 3.65
    h = 5.4
    top = 1.45
    
    add_card(slide, 0.8, top, w, h)
    tb1 = slide.shapes.add_textbox(Inches(0.95), Inches(top + 0.15), Inches(w - 0.3), Inches(h - 0.3))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.add_run("CƠ QUAN QUẢN LÝ\n(UBCKNN & SỞ GDCK)\n\n").font.bold = True
    p1.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p1.runs[0].font.size = Pt(12)
    body1 = (
        "• Hoàn thiện Khung pháp lý ESG:\n"
        "  - Bổ sung quy định hướng dẫn kiểm kê Scope 3 chuỗi cung ứng và yêu cầu công bố chênh lệch thu nhập lãnh đạo.\n\n"
        "• Chuẩn hóa Định dạng số (Digital Reporting):\n"
        "  - Yêu cầu doanh nghiệp nộp báo cáo dưới dạng văn bản số mở (chuẩn XBRL hoặc text PDF chuẩn).\n"
        "  - Hạn chế các tệp scan hình ảnh gây cản trở công tác giám sát tự động.\n\n"
        "• Thiết lập Cổng giám sát AI:\n"
        "  - Ứng dụng các thuật toán NLP để quét tự động toàn thị trường, phát hiện sớm các dấu hiệu tô hồng báo cáo."
    )
    p_b1 = tf1.add_paragraph()
    p_b1.text = body1
    p_b1.font.size = Pt(9.6)
    p_b1.font.color.rgb = C_TEXT_DARK

    add_card(slide, 4.84, top, w, h)
    tb2 = slide.shapes.add_textbox(Inches(5.0), Inches(top + 0.15), Inches(w - 0.3), Inches(h - 0.3))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.add_run("DOANH NGHIỆP NIÊM YẾT\n(BAN ĐIỀU HÀNH & HĐQT)\n\n").font.bold = True
    p2.runs[0].font.color.rgb = C_BLUE_ACCENT
    p2.runs[0].font.size = Pt(12)
    body2 = (
        "• Chuyển đổi Tư duy Công bố:\n"
        "  - Từ tư duy 'tiếp thị thành tích' sang tư duy 'giải trình trách nhiệm và quản trị rủi ro'.\n\n"
        "• Minh bạch hóa các Thách thức:\n"
        "  - Chủ động chia sẻ các rủi ro khí hậu, sự cố kỹ thuật và lộ trình khắc phục thay vì che giấu.\n"
        "  - Giúp nâng cao tính chính danh và độ tin cậy trong mắt các tổ chức xếp hạng tín nhiệm quốc tế.\n\n"
        "• Tích hợp ESG vào Chiến lược:\n"
        "  - Chuyển từ CSR từ thiện bề nổi sang CSV (tạo giá trị chia sẻ chung) trong chuỗi giá trị."
    )
    p_b2 = tf2.add_paragraph()
    p_b2.text = body2
    p_b2.font.size = Pt(9.6)
    p_b2.font.color.rgb = C_TEXT_DARK

    add_card(slide, 8.88, top, w, h)
    tb3 = slide.shapes.add_textbox(Inches(9.05), Inches(top + 0.15), Inches(w - 0.3), Inches(h - 0.3))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.add_run("NHÀ ĐẦU TƯ & KIỂM TOÁN\n(QUẢN LÝ QUỸ & ĐỊNH CHẾ)\n\n").font.bold = True
    p3.runs[0].font.color.rgb = C_GREEN_EMERALD
    p3.runs[0].font.size = Pt(12)
    body3 = (
        "• Công cụ Thẩm định Tự động:\n"
        "  - Tận dụng các mô hình NLP mã nguồn mở để sàng lọc nhanh hàng trăm báo cáo phi tài chính, tiết kiệm 90% thời gian rà soát sơ bộ.\n\n"
        "• Cảnh báo Tẩy xanh (Greenwashing Alert):\n"
        "  - Theo dõi Tỷ số Pos/Neg Ratio: nếu tỷ số quá cao bất thường mà thiếu số liệu định lượng thì cần đưa vào diện kiểm tra trọng điểm.\n\n"
        "• Định hướng Dòng vốn Bền vững:\n"
        "  - Lựa chọn danh mục đầu tư dựa trên bằng chứng cam kết SDG thực chất thay vì các lời hứa sáo rỗng."
    )
    p_b3 = tf3.add_paragraph()
    p_b3.text = body3
    p_b3.font.size = Pt(9.6)
    p_b3.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Đưa ra các hàm ý thực tiễn cụ thể, khả thi cho 3 đối tượng chính: Nhà nước, Doanh nghiệp và Nhà đầu tư/Kiểm toán.",
        "script": "Kính thưa các thầy cô, nghiên cứu của chúng em đem lại các hàm ý quản trị thiết thực: Đối với UBCKNN và Sở Giao dịch Chứng khoán, cần thúc đẩy chuẩn hóa báo cáo sang định dạng số mở (XBRL) và đưa AI vào giám sát tự động. Đối với các doanh nghiệp, đã đến lúc phải thay đổi tư duy: công bố thông tin trung thực, dám nói về khó khăn rủi ro sẽ giúp nâng cao xếp hạng tín nhiệm quốc tế tốt hơn là chỉ tô hồng thành tích. Đối với các nhà đầu tư và kiểm toán viên, công cụ NLP của đề tài mang lại khả năng sàng lọc tự động, giúp phát hiện sớm các nguy cơ tẩy xanh (greenwashing) và tiết kiệm tối đa thời gian rà soát.",
        "highlights": "Khuyến nghị 3 bên: Chuẩn hóa định dạng XBRL cho Nhà nước; Chuyển từ PR sang Giải trình trách nhiệm cho Doanh nghiệp; Dùng NLP thẩm định nhanh cho Nhà đầu tư.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao cần chuẩn hóa định dạng XBRL?': Trả lời: Định dạng XBRL gắn thẻ dữ liệu cho từng chỉ tiêu tài chính và phi tài chính, giúp máy tính có thể đọc và phân tích tự động 100% dữ liệu một cách chuẩn xác mà không bao giờ bị lỗi font hay lỗi trang scan như tệp PDF."
    })


def build_slide_17_limitations_future(prs):
    """Slide 17: Hạn Chế Đề Tài & Hướng Phát Triển Tương Lai."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "HẠN CHẾ CỦA ĐỀ TÀI & HƯỚNG PHÁT TRIỂN TƯƠNG LAI", "HƯỚNG PHÁT TRIỂN", 17)
    
    add_card(slide, 0.8, 1.45, 5.7, 5.4)
    tb_left = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(5.3), Inches(5.0))
    tf_l = tb_left.text_frame
    tf_l.word_wrap = True
    p_l = tf_l.paragraphs[0]
    p_l.add_run("HẠN CHẾ HIỆN TẠI CỦA NGHIÊN CỨU\n\n").font.bold = True
    p_l.runs[0].font.color.rgb = C_NAVY_PRIMARY
    p_l.runs[0].font.size = Pt(13)
    body_l = (
        "• Quy mô mẫu nghiên cứu:\n"
        "  - Tập trung vào 7 tập đoàn niêm yết lớn hàng đầu có báo cáo liên tục; chưa bao phủ toàn diện nhóm doanh nghiệp vừa và nhỏ (DNNVV) hoặc toàn bộ rổ chỉ số VN30.\n\n"
        "• Bản chất đo lường ngữ nghĩa:\n"
        "  - Sentence-BERT đo lường mức độ tương đồng về mặt từ ngữ và ngữ cảnh (semantic similarity), nhưng chưa thể tự động kiểm chứng tính xác thực (Fact-checking) của các con số định lượng (ví dụ: công ty báo giảm 20% phát thải nhưng thực tế có giảm đúng như vậy không).\n\n"
        "• Giới hạn của mô hình phân tích cảm xúc:\n"
        "  - PhoBERT được tinh chỉnh trên ngữ liệu chung, đôi khi chưa bao quát hết một số thuật ngữ kế toán kiểm toán chuyên biệt sâu."
    )
    p_bl = tf_l.add_paragraph()
    p_bl.text = body_l
    p_bl.font.size = Pt(10.2)
    p_bl.font.color.rgb = C_TEXT_DARK

    add_card(slide, 6.83, 1.45, 5.7, 5.4)
    tb_right = slide.shapes.add_textbox(Inches(7.05), Inches(1.6), Inches(5.3), Inches(5.0))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    p_r = tf_r.paragraphs[0]
    p_r.add_run("HƯỚNG PHÁT TRIỂN TIẾP THEO (GENAI & AGENTS)\n\n").font.bold = True
    p_r.runs[0].font.color.rgb = C_GREEN_EMERALD
    p_r.runs[0].font.size = Pt(13)
    body_r = (
        "• Mở rộng quy mô mẫu toàn thị trường:\n"
        "  - Nâng cấp hệ thống thu thập tự động để quét toàn bộ 300+ doanh nghiệp niêm yết trên HOSE và HNX có báo cáo lồng ghép.\n\n"
        "• Ứng dụng Mô hình Ngôn ngữ Lớn (LLMs & RAG):\n"
        "  - Tích hợp các LLMs tiên tiến (GPT-4, Gemini) kết hợp kỹ thuật Retrieval-Augmented Generation (RAG) để trích xuất số liệu phát thải và đối chiếu chéo với Báo cáo Tài chính.\n\n"
        "• Xây dựng Hệ thống Trợ lý Kiểm toán AI (Agentic ESG Auditor):\n"
        "  - Xây dựng hệ thống nhiều Agent tự động: Agent đọc bảng biểu, Agent kiểm tra tuân thủ Thông tư 96, Agent phát hiện bất thường.\n"
        "  - Cung cấp công cụ hỗ trợ đắc lực cho các kiểm toán viên trong kỷ nguyên số."
    )
    p_br = tf_r.add_paragraph()
    p_br.text = body_r
    p_br.font.size = Pt(10.2)
    p_br.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Thể hiện thái độ nghiên cứu nghiêm túc, thẳng thắn nhìn nhận hạn chế và vạch ra tầm nhìn ứng dụng công nghệ GenAI / LLMs trong tương lai.",
        "script": "Kính thưa Hội đồng, nhóm nghiên cứu cũng thẳng thắn nhìn nhận các hạn chế của đề tài: Thứ nhất là mẫu mới gồm 7 doanh nghiệp lớn; thứ hai là mô hình SBERT đo lường sự tương đồng về ngữ nghĩa chứ chưa thể tự động kiểm tra xem các con số định lượng có bị khai khống hay không. Trong giai đoạn tiếp theo, chúng em định hướng mở rộng mẫu lên toàn bộ rổ VN30 và đặc biệt là kết hợp các Mô hình Ngôn ngữ Lớn (LLMs) như Gemini hay GPT-4 cùng kỹ thuật RAG để xây dựng một Hệ thống Trợ lý Kiểm toán ESG Tự động (Agentic ESG Auditor), có khả năng đọc hiểu số liệu bảng biểu và đối chiếu chéo giữa báo cáo tài chính với báo cáo bền vững.",
        "highlights": "Thẳng thắn nhìn nhận hạn chế về Fact-checking; mở ra hướng đi tương lai với GenAI, LLMs và Agentic ESG Auditor.",
        "qa": "Thầy cô có thể hỏi: 'LLM có thay thế hoàn toàn SBERT trong đề tài này được không?': Trả lời: SBERT có ưu thế vượt trội về tốc độ tính toán vector nhanh gấp hàng trăm lần và chi phí cực thấp khi xử lý 100.000 câu văn bản. Mô hình kết hợp tối ưu là dùng SBERT để quét nhanh toàn bộ ngữ liệu, sau đó dùng LLMs để phân tích sâu các đoạn văn bản có nghi vấn tẩy xanh."
    })


def build_slide_18_conclusion(prs):
    """Slide 18: Kết Luận & Phiên Hỏi Đáp (Q&A)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, C_NAVY_DARK)
    
    accent_top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    set_shape_flat(accent_top, C_GOLD_ACCENT)

    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.733), Inches(0.8))
    p_t = tb_title.text_frame.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = "TỔNG KẾT 4 THÔNG ĐIỆP CỐT LÕI CỦA ĐỀ TÀI"
    r_t.font.name = FONT_HEADING
    r_t.font.size = Pt(20)
    r_t.font.bold = True
    r_t.font.color.rgb = C_WHITE

    conclusions = [
        ("1. TÍNH CHUYỂN GIAO THÀNH CÔNG",
         "Quy trình NLP của Kang & Kim (2022) hoàn toàn có khả năng thích ứng xuất sắc với ngôn ngữ tiếng Việt và bối cảnh doanh nghiệp tại thị trường mới nổi."),
        ("2. PHẢN ÁNH CHÂN THỰC ĐẶC THÙ NGÀNH",
         "Điểm số SDG phản ánh sát sao bản chất kinh doanh: Khối sản xuất dẫn đầu về Tài nguyên/Môi trường; Khối tài chính áp đảo về Kinh tế/Quản trị; Khối bán lẻ nổi bật về DE&I."),
        ("3. BỘC LỘ BỨC TRANH CÔNG BỐ TƯƠNG PHẢN",
         "Phát hiện sự phân hóa sâu sắc: Doanh nghiệp 'nói nhiều' về tăng trưởng và từ thiện CSR bề nổi, nhưng 'né tránh' đa dạng sinh học, Scope 3 và chênh lệch thu nhập."),
        ("4. XÁC NHẬN THIÊN LỆCH LẠC QUAN CẤU TRÚC",
         "Tỷ số Pos/Neg trung bình đạt 4,06 lần khẳng định sự tồn tại của chiến lược Quản trị Ấn tượng, đòi hỏi các công cụ AI hỗ trợ kiểm toán phi tài chính khách quan.")
    ]

    w_c = 5.7
    h_c = 1.95
    top1 = 1.4
    top2 = 3.55
    l1 = 0.8
    l2 = 6.83
    
    coords = [(l1, top1), (l2, top1), (l1, top2), (l2, top2)]
    
    for i, ((c_head, c_body), (cx, cy)) in enumerate(zip(conclusions, coords)):
        c_card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(cy), Inches(w_c), Inches(h_c))
        set_shape_flat(c_card, RGBColor(20, 45, 80), RGBColor(50, 90, 140), 1.2)
        
        c_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(cy), Inches(0.15), Inches(h_c))
        set_shape_flat(c_bar, C_GOLD_ACCENT)
        
        tb = slide.shapes.add_textbox(Inches(cx + 0.25), Inches(cy + 0.1), Inches(w_c - 0.35), Inches(h_c - 0.2))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p0 = tf.paragraphs[0]
        p0.add_run(c_head + "\n").font.bold = True
        p0.runs[0].font.color.rgb = C_GOLD_ACCENT
        p0.runs[0].font.size = Pt(11)
        
        p1 = tf.add_paragraph()
        p1.text = c_body
        p1.font.size = Pt(9.8)
        p1.font.color.rgb = RGBColor(225, 235, 250)

    qa_card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(5.7), Inches(11.733), Inches(1.3))
    set_shape_flat(qa_card, RGBColor(15, 35, 65), C_GOLD_ACCENT, 1.5)
    
    tb_qa = slide.shapes.add_textbox(Inches(1.0), Inches(5.75), Inches(11.333), Inches(1.2))
    tf_qa = tb_qa.text_frame
    tf_qa.word_wrap = True
    
    p_qa1 = tf_qa.paragraphs[0]
    p_qa1.alignment = PP_ALIGN.CENTER
    r_qa1 = p_qa1.add_run()
    r_qa1.text = "TRÂN TRỌNG CẢM ƠN QUÝ THẦY CÔ TRONG HỘI ĐỒNG KHOA HỌC!\n"
    r_qa1.font.name = FONT_HEADING
    r_qa1.font.size = Pt(15)
    r_qa1.font.bold = True
    r_qa1.font.color.rgb = C_GOLD_ACCENT
    
    p_qa2 = tf_qa.add_paragraph()
    p_qa2.alignment = PP_ALIGN.CENTER
    r_qa2 = p_qa2.add_run()
    r_qa2.text = "Nhóm nghiên cứu rất mong nhận được các câu hỏi và ý kiến đóng góp quý báu từ Quý Thầy Cô.\n(Tác giả: Lê Đan Sơn, Dương Thị Hoàn — 2026)"
    r_qa2.font.name = FONT_MAIN
    r_qa2.font.size = Pt(11)
    r_qa2.font.italic = True
    r_qa2.font.color.rgb = C_WHITE

    set_presenter_notes(slide, {
        "goal": "Tóm kết đĩnh đạc 4 thông điệp cốt lõi, gửi lời cảm ơn trang trọng đến Hội đồng và tự tin mở phiên Hỏi đáp (Q&A).",
        "script": "Kính thưa Quý Thầy Cô trong Hội đồng, để kết lại bài thuyết trình hôm nay, nhóm nghiên cứu xin gửi gắm 4 thông điệp cốt lõi: Thứ nhất, mô hình NLP của Kang và Kim hoàn toàn khả thi và thích ứng tuyệt vời với ngôn ngữ tiếng Việt. Thứ hai, điểm số SDG phản ánh rất chân thực mô hình kinh doanh của từng doanh nghiệp. Thứ ba, kết quả nghiên cứu đã bóc trần sự đối lập giữa những gì doanh nghiệp nói nhiều và những chủ đề đang bị né tránh. Và thứ tư, sự thiên lệch lạc quan (Pos/Neg đạt 4,06 lần) khẳng định rằng thị trường rất cần các công cụ AI hỗ trợ kiểm toán phi tài chính. Nhóm tác giả Lê Đan Sơn và Dương Thị Hoàn xin trân trọng cảm ơn sự lắng nghe của Quý Thầy Cô và rất mong nhận được những nhận xét, chỉ bảo quý báu từ Hội đồng!",
        "highlights": "Cúi đầu chào và mời các thầy cô đặt câu hỏi. Giữ phong thái tự tin, khiêm tốn và lắng nghe.",
        "qa": "Sẵn sàng mở lại các slide số liệu tương ứng khi thầy cô yêu cầu giải trình sâu hơn."
    })


def main():
    print("=" * 80)
    print("BẮT ĐẦU TẠO SLIDES THUYẾT TRÌNH BÁO CÁO HỘI ĐỒNG (16:9 WIDESCREEN)...")
    print("=" * 80)
    
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    print("[1/18] Tạo Slide 1: Trang Tiêu đề & Thông tin Tác giả...")
    build_slide_01_title(prs)
    
    print("[2/18] Tạo Slide 2: Đặt vấn đề & Bối cảnh Thể chế Việt Nam...")
    build_slide_02_context(prs)
    
    print("[3/18] Tạo Slide 3: Bài báo gốc Kang & Kim (2022) & Khoảng trống Nghiên cứu...")
    build_slide_03_original_paper(prs)
    
    print("[4/18] Tạo Slide 4: Khung Phương pháp luận 5 bước (Pipeline)...")
    build_slide_04_pipeline(prs)
    
    print("[5/18] Tạo Slide 5: Mẫu Dữ liệu Thực nghiệm (7 Doanh nghiệp, 42 Báo cáo)...")
    build_slide_05_sample(prs)
    
    print("[6/18] Tạo Slide 6: Kết quả 1 - Phân phối Điểm Tương đồng SDG Toàn cục...")
    build_slide_06_result1_similarity(prs)
    
    print("[7/18] Tạo Slide 7: Kết quả 2 - Cấu trúc 6 Nhóm SDG qua Heatmap...")
    build_slide_07_result2_heatmap(prs)
    
    print("[8/18] Tạo Slide 8: Đặc thù Ngành (Khối Sản xuất & Năng lượng: VNM, VCS, PAN, PLX)...")
    build_slide_08_result3_companies_p1(prs)
    
    print("[9/18] Tạo Slide 9: Đặc thù Ngành (Khối Tài chính & Bán lẻ: BVH, SSI, PNJ)...")
    build_slide_09_result3_companies_p2(prs)
    
    print("[10/18] Tạo Slide 10: Kết quả 4 - Xu hướng Dịch chuyển Chuỗi Thời gian (2020–2025)...")
    build_slide_10_result4_trends(prs)
    
    print("[11/18] Tạo Slide 11: Kết quả 5 - Sắc thái Cảm xúc & Thiên lệch Lạc quan...")
    build_slide_11_result5_sentiment(prs)
    
    print("[12/18] Tạo Slide 12: Kết quả 6 - Tỷ số Cảm xúc Pos/Neg Ratio & Độ nhạy bối cảnh...")
    build_slide_12_result6_sentiment_ratio(prs)
    
    print("[13/18] Tạo Slide 13: Thảo luận Chuyên sâu - Doanh nghiệp 'Nói nhiều về gì'...")
    build_slide_13_discussion_talk_heavy(prs)
    
    print("[14/18] Tạo Slide 14: Thảo luận Chuyên sâu - Doanh nghiệp 'Ít nói về gì' (Né tránh)...")
    build_slide_14_discussion_rarely_talk(prs)
    
    print("[15/18] Tạo Slide 15: Đóng góp Học thuật & So sánh Đối chuẩn Kang & Kim (2022)...")
    build_slide_15_comparison(prs)
    
    print("[16/18] Tạo Slide 16: Hàm ý Thực tiễn & Đề xuất Chính sách...")
    build_slide_16_implications(prs)
    
    print("[17/18] Tạo Slide 17: Hạn chế của Đề tài & Hướng Phát triển Tương lai (GenAI)...")
    build_slide_17_limitations_future(prs)
    
    print("[18/18] Tạo Slide 18: Tổng kết 4 Thông điệp Cốt lõi & Phiên Hỏi đáp (Q&A)...")
    build_slide_18_conclusion(prs)
    
    # Lưu file PPTX
    prs.save(str(OUTPUT_PPTX))
    print(f"\n=> Đã lưu thành công bộ slide tại: {OUTPUT_PPTX}")
    
    # Sao chép sang thư mục papers/
    PAPERS_DIR.mkdir(parents=True, exist_ok=True)
    dest_copy = PAPERS_DIR / "bao_cao_nghien_cuu_sdg_vietnam.pptx"
    shutil.copyfile(OUTPUT_PPTX, dest_copy)
    print(f"=> Đã sao chép một bản vào: {dest_copy}")
    
    file_size_mb = OUTPUT_PPTX.stat().st_size / (1024 * 1024)
    print(f"=> Kích thước tệp: {file_size_mb:.2f} MB")
    print("=" * 80)
    print("HOÀN TẤT THÀNH CÔNG!")
    print("=" * 80)


if __name__ == "__main__":
    main()
