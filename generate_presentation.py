"""Tạo bộ slide thuyết trình PowerPoint (.pptx) chuẩn học thuật 16:9 báo cáo Hội đồng.
Phiên bản nâng cấp v2:
- Tóm tắt kết quả của bài báo gốc Kang & Kim (2022) và nêu rõ 4 vấn đề đề tài khắc phục
- Biểu đồ phân tích SDG cốt lõi cho từng doanh nghiệp ví dụ (VCS, VNM, PAN, PLX, PNJ, BVH, SSI)
- Nêu rõ doanh nghiệp nào có goal nào có similarity cao vượt trội và giải thích nguyên nhân
- So sánh định lượng trực tiếp từng kết quả (trung bình, độ lệch chuẩn, 6 nhóm, tỷ số Pos/Neg) với paper gốc
- Rút ra kết luận học thuật sâu sắc về báo cáo doanh nghiệp tiếng Việt
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

# Monkey patch để hỗ trợ add_run(text)
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

# BẢNG MÀU CHUẨN HỌC THUẬT (ACADEMIC SLATE & NAVY PALETTE)
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
C_PURPLE_ACCENT= RGBColor(112, 48, 160)   # #7030A0 - Màu bình đẳng giới DE&I
C_WHITE        = RGBColor(255, 255, 255)

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


def add_slide_header(slide, title_text: str, category_tag: str = "BÁO CÁO KHOA HỌC", slide_num: int = 1):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, C_BG_LIGHT)
    
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.15))
    set_shape_flat(top_bar, C_CARD_BG, C_BORDER_LIGHT, 0.75)
    
    accent_strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.06))
    set_shape_flat(accent_strip, C_GOLD_ACCENT)
    
    tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.18), Inches(2.8), Inches(0.28))
    set_shape_flat(tag_box, C_NAVY_PRIMARY)
    tf_tag = tag_box.text_frame
    tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_tag = tf_tag.paragraphs[0]
    p_tag.alignment = PP_ALIGN.CENTER
    r_tag = p_tag.add_run()
    r_tag.text = category_tag.upper()
    r_tag.font.name = FONT_MAIN
    r_tag.font.size = Pt(9.5)
    r_tag.font.bold = True
    r_tag.font.color.rgb = C_WHITE
    
    tx_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.48), Inches(10.5), Inches(0.6))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.name = FONT_HEADING
    run.font.size = Pt(17.5)
    run.font.bold = True
    run.font.color.rgb = C_NAVY_PRIMARY

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
    
    foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.15), Inches(11.8), Inches(0.3))
    tf_f = foot_box.text_frame
    p_f = tf_f.paragraphs[0]
    r_f = p_f.add_run()
    r_f.text = "Đề tài: NLP Đa ngữ trong Phân tích SDG & Cảm xúc Báo cáo PTBV tại Việt Nam | Tác giả: Lê Đan Sơn, Dương Thị Hoàn"
    r_f.font.name = FONT_MAIN
    r_f.font.size = Pt(9)
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
        add_section("ĐIỂM NHẤN CẦN NHỚ", notes_dict["highlights"])
    if "qa" in notes_dict:
        add_section("DỰ ĐOÁN CÂU HỎI HỘI ĐỒNG & CÁCH TRẢ LỜI", notes_dict["qa"])


# ==============================================================================
# XÂY DỰNG 18 SLIDE
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
    r_b = p_b.add_run("BÁO CÁO KẾT QUẢ NGHIÊN CỨU KHOA HỌC")
    r_b.font.name = FONT_MAIN
    r_b.font.size = Pt(11)
    r_b.font.bold = True
    r_b.font.color.rgb = C_NAVY_DARK
    
    tb_title = slide.shapes.add_textbox(Inches(1.3), Inches(1.6), Inches(10.7), Inches(1.8))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    r_t = p_t.add_run(
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
    r_en = p_en.add_run(
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
    r_f1 = p_f1.add_run("KHUNG PHƯƠNG PHÁP LUẬN KẾ THỪA:\n")
    r_f1.font.bold = True
    r_f1.font.size = Pt(10)
    r_f1.font.color.rgb = C_GOLD_ACCENT
    r_f2 = p_f1.add_run(
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
    r_a1 = p_a.add_run("TÁC GIẢ THỰC HIỆN ĐỀ TÀI:\n")
    r_a1.font.bold = True
    r_a1.font.size = Pt(10)
    r_a1.font.color.rgb = C_GOLD_ACCENT
    r_a2 = p_a.add_run("• Lê Đan Sơn\n• Dương Thị Hoàn\n\nNăm thực hiện: 2026")
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
    r1 = p1.add_run("1. ĐỘNG LỰC THỂ CHẾ HÓA ESG\n\n")
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY_PRIMARY
    r1.font.size = Pt(13)
    
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
    r2 = p2.add_run("2. NGHỊCH LÝ & THÁCH THỨC\n\n")
    r2.font.bold = True
    r2.font.color.rgb = C_RED_ACCENT
    r2.font.size = Pt(13)
    
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
    r3 = p3.add_run("3. ĐỘT PHÁ CÔNG NGHỆ NLP\n\n")
    r3.font.bold = True
    r3.font.color.rgb = C_GREEN_EMERALD
    r3.font.size = Pt(13)
    
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
    """Slide 3: Bài Báo Gốc Kang & Kim (2022) & Sự Thích Ứng (Đã bổ sung số liệu cụ thể và 4 hạn chế khắc phục)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BÀI BÁO GỐC KANG & KIM (2022) & ĐỀ TÀI NÀY KHẮC PHỤC ĐIỀU GÌ?", "TỔNG QUAN HỌC THUẬT", 3)
    
    # Cột trái: Kết quả cụ thể của Paper gốc
    add_card(slide, 0.8, 1.45, 5.7, 5.4)
    tb_left = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(5.3), Inches(5.0))
    tf_l = tb_left.text_frame
    tf_l.word_wrap = True
    p_l = tf_l.paragraphs[0]
    r_l = p_l.add_run("KẾT QUẢ CỐT LÕI BÀI BÁO GỐC: KANG & KIM (2022)\n")
    r_l.font.bold = True
    r_l.font.color.rgb = C_NAVY_PRIMARY
    r_l.font.size = Pt(12)
    
    p_sub = tf_l.add_paragraph()
    p_sub.text = "Applied Sciences (MDPI), 12(11), 5614 | SCIE / Scopus Q2"
    p_sub.font.size = Pt(9.5)
    p_sub.font.italic = True
    p_sub.font.color.rgb = C_TEXT_MUTED
    
    p_body_l = tf_l.add_paragraph()
    p_body_l.text = (
        "• Dữ liệu thực nghiệm toàn cầu:\n"
        "  - Thu thập 3.529 báo cáo PTBV quốc tế (2011–2020) từ Corporate Register.\n"
        "  - 100% ngữ liệu sử dụng ngôn ngữ tiếng Anh.\n\n"
        "• Kết quả định lượng điểm tương đồng SBERT:\n"
        "  - Phân phối chuẩn Gaussian hình chuông: μ ≈ 44,80 điểm, σ ≈ 11,50 điểm.\n"
        "  - Cấu trúc 6 nhóm: Nhóm Kinh tế (Economic) luôn cao nhất (~48–54đ), nhóm Công bằng (Equity) luôn thấp nhất (~38–44đ).\n\n"
        "• Kết quả cảm xúc DistilBERT 2 lớp (nhị phân):\n"
        "  - Câu Tích cực áp đảo (~78%), câu Tiêu cực rất thấp (~15%).\n"
        "  - Tỷ số Cảm xúc Pos/Neg bình quân đạt ~5,2 lần (xác nhận Hiệu ứng Pollyanna trên toàn cầu)."
    )
    p_body_l.font.size = Pt(9.8)
    p_body_l.font.color.rgb = C_TEXT_DARK

    # Cột phải: Đề tài này khắc phục vấn đề nào?
    add_card(slide, 6.83, 1.45, 5.7, 5.4)
    tb_right = slide.shapes.add_textbox(Inches(7.05), Inches(1.6), Inches(5.3), Inches(5.0))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    p_r = tf_r.paragraphs[0]
    r_r = p_r.add_run("ĐỀ TÀI NÀY KHẮC PHỤC 4 VẤN ĐỀ CỐT TỬ CỦA PAPER GỐC\n")
    r_r.font.bold = True
    r_r.font.color.rgb = C_GREEN_EMERALD
    r_r.font.size = Pt(12)
    
    p_sub2 = tf_r.add_paragraph()
    p_sub2.text = "Nghiên cứu của Lê Đan Sơn & Dương Thị Hoàn (2026)"
    p_sub2.font.size = Pt(9.5)
    p_sub2.font.italic = True
    p_sub2.font.color.rgb = C_TEXT_MUTED

    p_body_r = tf_r.add_paragraph()
    p_body_r.text = (
        "1. Khắc phục rào cản đơn ngữ tiếng Anh:\n"
        "   - Xây dựng pipeline NLP đa ngữ: vietnamese-sbert (768 chiều) đặc thù cho từ đơn âm tiếng Việt kết hợp MiniLM cho tiếng Anh.\n"
        "   - Chuẩn hóa bộ ngữ liệu 1.032 câu song ngữ đối sánh 169 mục tiêu SDG.\n\n"
        "2. Khắc phục lỗi tê liệt trước tệp scan hình ảnh:\n"
        "   - Bài gốc bỏ qua các file scan lỗi font -> Đề tài tích hợp Tesseract OCR khôi phục 100% dữ liệu từ các trang scan phức tạp (như PNJ 2022).\n\n"
        "3. Khắc phục phân loại nhị phân gượng ép của mô hình cảm xúc:\n"
        "   - Bài gốc ép các câu số liệu kỹ thuật vào Tích cực/Tiêu cực -> Đề tài nâng cấp lên PhoBERT 3 lớp, bổ sung lớp Trung tính (32,87%) bảo toàn thông số vận hành.\n\n"
        "4. Từ khảo sát bề mặt toàn cầu đến giải mã chiều sâu đặc thù ngành:\n"
        "   - Phân tích chi tiết từng mô hình kinh doanh ngành và hiện tượng 'Nói nhiều về gì vs Né tránh điều gì' tại thị trường mới nổi Việt Nam."
    )
    p_body_r.font.size = Pt(9.6)
    p_body_r.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Nêu rõ kết quả định lượng của bài báo gốc (3.529 báo cáo, μ=44.8, σ=11.5, Pos/Neg ~ 5.2x) và chứng minh 4 vấn đề kỹ thuật mà đề tài này đã giải quyết triệt để.",
        "script": "Kính thưa Quý Thầy Cô, để khẳng định tính học thuật, đề tài kế thừa từ công trình gốc của Kang và Kim (2022) trên tạp chí Applied Sciences. Họ phân tích 3.529 báo cáo toàn cầu bằng tiếng Anh, cho ra phân phối chuẩn μ=44,80, Economic cao nhất, Equity thấp nhất, và Tỷ số Pos/Neg đạt khoảng 5,2 lần. Tuy nhiên, bài gốc có 4 điểm nghẽn lớn mà đề tài của chúng em đã khắc phục trọn vẹn: Thứ nhất, bài gốc chỉ chạy tiếng Anh, chúng em xây dựng pipeline đa ngữ cho tiếng Việt; thứ hai, bài gốc bỏ qua tệp scan, chúng em dùng OCR khôi phục thành công; thứ ba, bài gốc ép cảm xúc nhị phân 2 lớp, chúng em nâng cấp lên PhoBERT 3 lớp với 32,87% câu Trung tính kỹ thuật; và thứ tư, chúng em đi sâu giải mã bản chất từng ngành kinh tế tại Việt Nam thay vì chỉ thống kê bề mặt.",
        "highlights": "4 điểm khắc phục: Đa ngữ tiếng Việt, OCR tệp scan, PhoBERT 3 lớp (Trung tính), Giải mã đặc thù ngành.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao việc bổ sung lớp Trung tính lại là một đóng góp học thuật?': Trả lời: Trong văn bản phi tài chính, một câu như 'Nhà máy tiêu thụ 15 triệu kWh điện' là dữ liệu kỹ thuật thuần túy. Nếu ép nhị phân như Kang & Kim, máy buộc phải gán nhãn tích cực hoặc tiêu cực sai lệch. Giữ lại 32,87% câu trung tính giúp việc đo lường Tỷ số Pos/Neg sau đó phản ánh chân thực mức độ thiên lệch quản trị ấn tượng."
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
        r_h1 = p_h1.add_run(b_name + "\n")
        r_h1.font.name = FONT_MAIN
        r_h1.font.size = Pt(10)
        r_h1.font.bold = True
        r_h1.font.color.rgb = C_GOLD_ACCENT
        
        r_h2 = p_h1.add_run(b_title)
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
        r1 = p1.add_run(val + "\n")
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(20)
        r1.font.bold = True
        r1.font.color.rgb = col
        
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run(label + "\n")
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(9.5)
        r2.font.bold = True
        r2.font.color.rgb = C_TEXT_DARK
        
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        r3 = p3.add_run(sub)
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
        run = p.add_run(h)
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
            run = p.add_run(val)
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
    """Slide 6: Kết quả 1 - Phân phối Điểm Tương đồng SDG & ĐỐI CHUẨN ĐỊNH LƯỢNG VỚI PAPER GỐC."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 1: PHÂN PHỐI ĐIỂM TƯƠNG ĐỒNG SDG & ĐỐI CHUẨN PAPER GỐC", "KẾT QUẢ THỰC NGHIỆM", 6)
    
    # Cột trái: Luận điểm & Bảng đối chuẩn định lượng
    add_card(slide, 0.8, 1.45, 5.2, 5.4)
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(4.8), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p0 = tf.paragraphs[0]
    r0 = p0.add_run("ĐẶC TRƯNG PHÂN PHỐI & ĐỐI CHUẨN QUỐC TẾ\n")
    r0.font.bold = True
    r0.font.color.rgb = C_NAVY_PRIMARY
    r0.font.size = Pt(12)
    
    body = (
        "• BẢNG ĐỐI CHUẨN ĐỊNH LƯỢNG TRỰC TIẾP:\n"
        "  ┌────────────────────────┬─────────────┬─────────────┐\n"
        "  │ Chỉ số thống kê        │ Việt Nam    │ Kang & Kim  │\n"
        "  ├────────────────────────┼─────────────┼─────────────┤\n"
        "  │ Điểm trung bình (μ)    │ 45,43 điểm  │ 44,80 điểm  │\n"
        "  │ Độ lệch chuẩn (σ)      │ 11,87 điểm  │ 11,50 điểm  │\n"
        "  │ Hình thái phân phối    │ Chuông Gauss│ Chuông Gauss│\n"
        "  │ Đuôi phải (>65 điểm)   │ 8,50% câu   │ ~8,00% câu  │\n"
        "  └────────────────────────┴─────────────┴─────────────┘\n\n"
        "• Nhận định học thuật quan trọng:\n"
        "  - Độ chênh lệch giữa điểm trung bình chỉ 0,63 điểm; độ phân tán lệch chuẩn gần như trùng khớp hoàn toàn!\n"
        "  - Chứng minh vietnamese-sbert biểu diễn ngữ nghĩa tiếng Việt đạt độ chuẩn xác và ổn định tương đương tuyệt đối với SBERT tiếng Anh trên ngữ liệu toàn cầu.\n\n"
        "• Phân khúc đuôi phải chuyên sâu (>65 điểm):\n"
        "  - Đại diện cho các cam kết kỹ thuật cao: xử lý tuần hoàn bùn thải, kiểm kê khí nhà kính ISO 14064, năng lượng tái tạo."
    )
    p_b = tf.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.5)
    p_b.font.color.rgb = C_TEXT_DARK

    # Cột phải: Hình ảnh similarity_hist.png
    add_card(slide, 6.2, 1.45, 6.333, 5.4)
    fig_path = FIGURES_DIR / "similarity_hist.png"
    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(6.35), Inches(1.65), width=Inches(6.033))
        
    tb_cap = slide.shapes.add_textbox(Inches(6.35), Inches(6.25), Inches(6.033), Inches(0.5))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run("Hình 1: Phân phối tần suất điểm tương đồng SDG của 96.461 câu (thang đo Min-Max 0–100)")
    r_cap.font.size = Pt(9.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    set_presenter_notes(slide, {
        "goal": "So sánh trực tiếp các tham số định lượng (μ, σ, hình thái chuông) giữa Việt Nam và nghiên cứu gốc của Kang & Kim (2022).",
        "script": "Kính thưa các thầy cô, tại Hình 1, chúng em đối chiếu trực tiếp phân phối điểm tương đồng SDG của 96.461 câu văn bản Việt Nam với kết quả của Kang và Kim (2022). Kết quả đối chuẩn cho thấy sự tương đồng đáng kinh ngạc: Điểm trung bình của Việt Nam là 45,43 so với 44,80 của bài gốc (chỉ lệch 0,63 điểm); độ lệch chuẩn là 11,87 so với 11,50 của bài gốc. Cả hai phân phối đều có hình chuông Gaussian đối xứng hoàn hảo và có khoảng 8,5% số câu văn nằm ở đuôi bên phải (>65 điểm). Sự trùng khớp này là bằng chứng thực nghiệm vững chắc khẳng định Sentence-BERT tiếng Việt đạt độ tin cậy và sự ổn định ngữ nghĩa tương đương 100% với các mô hình tiếng Anh quốc tế.",
        "highlights": "So sánh: μ = 45,43 vs 44,80; σ = 11,87 vs 11,50. Khẳng định SBERT tiếng Việt hoạt động cực kỳ ổn định.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao điểm trung bình ở cả Việt Nam và quốc tế đều quanh mức 45 điểm?': Trả lời: Bởi vì cấu trúc ngôn ngữ của một báo cáo bền vững dù ở đâu trên thế giới cũng luôn có khoảng 70-75% câu văn mang tính giới thiệu bối cảnh, mô tả quy trình chung (similarity vừa phải), và chỉ có khoảng 8-10% câu văn mang cam kết kỹ thuật định lượng sâu (similarity >65 điểm). Đây là quy luật văn phong tự nhiên."
    })


def build_slide_07_result2_heatmap(prs):
    """Slide 7: Kết quả 2 - Cấu trúc 6 Nhóm SDG Qua Heatmap & QUY LUẬT TOÀN CẦU."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 2: CẤU TRÚC 6 NHÓM SDG & QUY LUẬT TÂM LÝ TOÀN CẦU", "KẾT QUẢ THỰC NGHIỆM", 7)
    
    add_card(slide, 0.8, 1.45, 4.4, 5.4)
    fig_path = FIGURES_DIR / "heatmap_6cat.png"
    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(0.95), Inches(1.6), width=Inches(4.1), height=Inches(4.7))
        
    tb_cap = slide.shapes.add_textbox(Inches(0.95), Inches(6.35), Inches(4.1), Inches(0.45))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run("Hình 2: Heatmap 6 nhóm SDG của 7 DN (2020–2025)")
    r_cap.font.size = Pt(9.0)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    add_card(slide, 5.4, 1.45, 7.133, 5.4)
    tb_right = slide.shapes.add_textbox(Inches(5.6), Inches(1.6), Inches(6.7), Inches(5.1))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    r0 = p0.add_run("THỨ BẬC ƯU TIÊN & SỰ TƯƠNG ĐỒNG VỚI PAPER GỐC\n")
    r0.font.bold = True
    r0.font.color.rgb = C_NAVY_PRIMARY
    r0.font.size = Pt(12)
    
    body = (
        "• THỨ BẬC PHÂN TẦNG TẠI VIỆT NAM:\n"
        "  1. Kinh tế (Economic): 48,0 – 52,5 điểm (Sắc đỏ sẫm nhất - Doanh thu, đổi mới công nghệ).\n"
        "  2. Xã hội (Social): 46,2 – 50,8 điểm (Quản trị minh bạch, đối tác phát triển bền vững).\n"
        "  3. Tài nguyên (Resources): 44,1 – 48,9 điểm (Tiết kiệm năng lượng, tuần hoàn nước thải).\n"
        "  4. Đời sống (Life): 43,0 – 47,5 điểm (An sinh xã hội, phúc lợi y tế người lao động).\n"
        "  5. Môi trường (Environments): 41,2 – 47,3 điểm (Bứt phá mạnh ở VNM, VCS giai đoạn 2024–2025).\n"
        "  6. Công bằng (Equity): 39,0 – 44,2 điểm (Sắc vàng nhạt nhất toàn mẫu - Vùng trũng lớn nhất).\n\n"
        "• SO SÁNH VỚI QUY LUẬT TOÀN CẦU CỦA KANG & KIM (2022):\n"
        "  - Trong nghiên cứu gốc trên 3.529 báo cáo thế giới: Nhóm Kinh tế luôn đạt điểm cao nhất (~48-54đ) và nhóm Công bằng luôn đạt điểm thấp nhất (~38-44đ).\n"
        "  - Kết quả tại Việt Nam hoàn toàn tái lập quy luật này! Điều này khẳng định: Tâm lý ưu tiên tăng trưởng kinh doanh trước khi chú trọng đến bình đẳng giới hay giảm bất bình đẳng thu nhập là một quy luật phổ quát toàn cầu, không chỉ riêng tại thị trường mới nổi."
    )
    p_b = tf_r.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.6)
    p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Phân tích trật tự ưu tiên 6 nhóm SDG từ Heatmap và so sánh quy luật này với nghiên cứu quốc tế của Kang & Kim (2022).",
        "script": "Kính thưa Hội đồng, Hình 2 là biểu đồ nhiệt Heatmap thể hiện mức độ gắn kết với 6 nhóm nhu cầu con người. Điều thú vị là khi đối chiếu với nghiên cứu của Kang và Kim trên 3.529 tập đoàn toàn cầu, chúng em ghi nhận sự trùng hợp hoàn toàn về trật tự thứ bậc: Cột Kinh tế (Economic) luôn giữ sắc đỏ đậm nhất, và cột Công bằng (Equity) luôn mang sắc vàng nhạt nhất. Điều này chứng minh rằng tâm lý doanh nghiệp ưu tiên các mục tiêu tạo ra lợi nhuận và việc làm trực tiếp trước khi quan tâm tới bình đẳng giới cấp cao hay giảm chênh lệch giàu nghèo là một quy luật toàn cầu, phản ánh áp lực tối đa hóa giá trị cổ đông ngắn hạn.",
        "highlights": "Quy luật toàn cầu: Kinh tế luôn cao nhất, Công bằng luôn thấp nhất ở cả Việt Nam và thế giới.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao nhóm Môi trường ở Việt Nam lại có sự bứt phá ở các năm cuối?': Trả lời: Do tác động trực tiếp của cam kết COP26 Net Zero 2050 và Thông tư 96/2020/TT-BTC, các doanh nghiệp đầu ngành như Vinamilk và Vicostone bắt đầu đầu tư mạnh cho kiểm kê khí nhà kính và kinh tế tuần hoàn, kéo điểm nhóm Môi trường tăng vọt hơn 6 điểm."
    })


def build_slide_08_result3_companies_p1(prs):
    """Slide 8: Đặc thù Ngành (Khối Sản Xuất & Năng Lượng) KÈM BIỂU ĐỒ MINH HỌA TỪNG DOANH NGHIỆP."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "ĐẶC THÙ NGÀNH: KHỐI SẢN XUẤT & NĂNG LƯỢNG (VNM, VCS, PAN, PLX)", "ĐẶC THÙ NGÀNH DOANH NGHIỆP", 8)
    
    # Cột trái: Biểu đồ slide_manuf_sdgs.png
    add_card(slide, 0.8, 1.45, 5.8, 5.4)
    fig_path = FIGURES_DIR / "slide_manuf_sdgs.png"
    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(0.95), Inches(1.6), width=Inches(5.5), height=Inches(4.8))
        
    tb_cap = slide.shapes.add_textbox(Inches(0.95), Inches(6.45), Inches(5.5), Inches(0.35))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run("Hình: Điểm số các mục tiêu SDG cốt lõi năm 2025 của khối Sản xuất & Năng lượng")
    r_cap.font.size = Pt(8.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    # Cột phải: Phân tích chi tiết Goal nào cao và tại sao
    add_card(slide, 6.8, 1.45, 5.733, 5.4)
    tb_right = slide.shapes.add_textbox(Inches(7.0), Inches(1.6), Inches(5.333), Inches(5.1))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    r0 = p0.add_run("MÔ HÌNH KINH DOANH ĐỊNH HÌNH SDG NÀO CAO VƯỢT TRỘI?\n")
    r0.font.bold = True
    r0.font.color.rgb = C_NAVY_PRIMARY
    r0.font.size = Pt(11.5)
    
    body = (
        "Từ biểu đồ phân tích, từng doanh nghiệp bộc lộ rõ các mục tiêu SDG dẫn đầu:\n\n"
        "• Vinamilk (VNM) — Bứt phá SDG 13 (Khí hậu) & SDG 12 (Tuần hoàn):\n"
        "  - SDG 13 đạt 45,36đ, SDG 12 đạt 49,62đ; nhóm Environments bứt phá mạnh nhất toàn mẫu (+6,06 điểm từ 41,23 lên 47,29đ).\n"
        "  - Nguyên nhân: 15 trang trại và 13 nhà máy đối diện áp lực giảm phát thải; tiên phong đạt chứng nhận trung hòa Carbon PAS 2060, Green Farm tuần hoàn 100% nước.\n\n"
        "• Vicostone (VCS) — Đỉnh cao SDG 9 (Đổi mới) & SDG 12 (Tuần hoàn):\n"
        "  - Dẫn đầu toàn mẫu: SDG 9 đạt 52,96đ, SDG 12 đạt 53,18đ, SDG 7 đạt 52,80đ.\n"
        "  - Nguyên nhân: Chế tác đá thạch anh nhân tạo Breton (Ý); tái chế 100% bùn thải đá thành phụ gia xi măng; chứng chỉ an toàn hóa chất Greenguard Gold.\n\n"
        "• The PAN Group (PAN) — Dẫn đầu SDG 2 (An ninh Lương thực: 47,47đ):\n"
        "  - Chuỗi giá trị nông nghiệp thực phẩm khép kín; lúa gạo giảm phát thải carbon (Vinaseed), nuôi tôm sinh thái không kháng sinh (Fimex VN).\n\n"
        "• Petrolimex (PLX) — Trọng tâm SDG 7 (Năng lượng: 48,04đ) & SDG 13 (46,19đ):\n"
        "  - Doanh nghiệp xăng dầu hạ nguồn chuyển dịch xanh: Phân phối nhiên liệu Euro 5, điện mặt trời trạm xăng, kiểm kê khí nhà kính ISO 14064-1."
    )
    p_b = tf_r.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.2)
    p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Chỉ trực tiếp vào biểu đồ minh họa và giải thích rõ với từng doanh nghiệp sản xuất thì goal nào cao vượt trội và lý do kinh doanh đằng sau.",
        "script": "Kính thưa các thầy cô, nhìn vào biểu đồ bên trái, chúng ta thấy mô hình NLP phản ánh cực kỳ sát thực trạng kỹ thuật của từng công ty: Vicostone dẫn đầu tuyệt đối ở SDG 9 Đổi mới hạ tầng (52,96 điểm) và SDG 12 Sản xuất tuần hoàn (53,18 điểm) nhờ công nghệ rung ép thạch anh Breton và tái chế 100% bùn thải đá. Vinamilk có bước nhảy vọt ở SDG 13 Hành động khí hậu (tăng lên 45,36 điểm) và nhóm Môi trường tăng hơn 6 điểm nhờ tiên phong đạt chứng nhận Net Zero PAS 2060. PAN Group dẫn đầu khối sản xuất về SDG 2 Nông nghiệp và an ninh lương thực (47,47 điểm) nhờ giống lúa phát thải thấp. Còn Petrolimex tập trung cao nhất vào SDG 7 Năng lượng sạch và SDG 13 nhờ phân phối nhiên liệu Euro 5.",
        "highlights": "VCS cao nhất SDG 9 & 12; VNM bứt phá SDG 13 (PAS 2060); PAN dẫn đầu SDG 2; PLX trọng tâm SDG 7 & 13.",
        "qa": "Thầy cô có thể hỏi: 'Tại sao điểm SDG 12 của Vicostone lại cao hơn Vinamilk?': Trả lời: Bởi vì Vicostone là ngành chế tạo công nghiệp nặng tiêu hao khoáng sản, nên họ tập trung giải trình sâu sắc về tỷ lệ tái sinh bùn thải đá thành phụ gia xi măng và kiểm soát dư lượng hóa chất hữu cơ bay hơi (VOC), giúp câu văn tương đồng rất mạnh với tiêu chí tái chế của SDG 12."
    })


def build_slide_09_result3_companies_p2(prs):
    """Slide 9: Đặc thù Ngành (Khối Tài Chính & Bán Lẻ) KÈM BIỂU ĐỒ MINH HỌA TỪNG DOANH NGHIỆP."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "ĐẶC THÙ NGÀNH: KHỐI TÀI CHÍNH & BÁN LẺ (PNJ, BVH, SSI)", "ĐẶC THÙ NGÀNH DOANH NGHIỆP", 9)
    
    # Cột trái: Biểu đồ slide_finance_sdgs.png
    add_card(slide, 0.8, 1.45, 5.8, 5.4)
    fig_path = FIGURES_DIR / "slide_finance_sdgs.png"
    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(0.95), Inches(1.6), width=Inches(5.5), height=Inches(4.8))
        
    tb_cap = slide.shapes.add_textbox(Inches(0.95), Inches(6.45), Inches(5.5), Inches(0.35))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run("Hình: Điểm số các mục tiêu SDG cốt lõi năm 2025 của PNJ, Bảo Việt (BVH) và SSI")
    r_cap.font.size = Pt(8.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    # Cột phải: Phân tích chi tiết Goal nào cao và tại sao
    add_card(slide, 6.8, 1.45, 5.733, 5.4)
    tb_right = slide.shapes.add_textbox(Inches(7.0), Inches(1.6), Inches(5.333), Inches(5.1))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    r0 = p0.add_run("ĐIỂM SÁNG BÌNH ĐẲNG GIỚI & DÒNG VỐN TÀI CHÍNH XANH\n")
    r0.font.bold = True
    r0.font.color.rgb = C_NAVY_PRIMARY
    r0.font.size = Pt(11.5)
    
    body = (
        "Khối dịch vụ tài chính và bán lẻ thể hiện bản đồ SDG hoàn toàn khác biệt:\n\n"
        "• PNJ — DẪN ĐẦU TOÀN MẪU SDG 5 (BÌNH ĐẲNG GIỚI: 40,70 ĐIỂM):\n"
        "  - Điểm SDG 5 của PNJ cao nhất trong toàn bộ 7 doanh nghiệp, vượt trội so với các ngành sản xuất nặng (chỉ 35–39 điểm).\n"
        "  - Nguyên nhân: Đặc thù mạng lưới bán lẻ kim hoàn với lực lượng lao động nữ chiếm >60%; tiên phong áp dụng trụ cột Đa dạng, Bình đẳng & Hòa nhập (DE&I), chính sách phát triển nữ nghệ nhân kim hoàn.\n\n"
        "• Tập đoàn Bảo Việt (BVH) — DẪN ĐẦU TOÀN MẪU SDG 17 (ĐỐI TÁC: 53,32 ĐIỂM):\n"
        "  - Điểm SDG 17 đạt đỉnh 53,32đ; SDG 9 đạt 51,46đ; SDG 8 đạt 50,62đ.\n"
        "  - Nguyên nhân: Áp dụng chuẩn Báo cáo Tích hợp quốc tế <IIRC> kết nối 6 nguồn vốn; quản trị minh bạch; phát triển bảo hiểm vi mô an sinh xã hội cho người nghèo.\n\n"
        "• Chứng khoán SSI — Trọng tâm SDG 17 (47,17đ) & SDG 9 (46,18đ):\n"
        "  - Định chế dẫn dắt dòng vốn xanh: Tư vấn phát hành Trái phiếu Xanh (Green Bonds), xây dựng khung thẩm định ESG trong đầu tư và tài trợ vốn bền vững."
    )
    p_b = tf_r.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.2)
    p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Chỉ vào biểu đồ minh họa và phân tích điểm sáng bình đẳng giới SDG 5 của PNJ cùng vai trò đối tác tài chính xanh SDG 17 của Bảo Việt và SSI.",
        "script": "Kính thưa các thầy cô, biểu đồ bên trái phác họa bức tranh rất đặc sắc của khối dịch vụ và bán lẻ: PNJ là điểm sáng duy nhất trong toàn mẫu dẫn đầu tuyệt đối ở SDG 5 Bình đẳng giới với 40,70 điểm, vượt trội hoàn toàn so với khối sản xuất nặng chỉ đạt 35-39 điểm. Lý do là PNJ có hơn 60% lao động là nữ và họ đầu tư rất bài bản cho trụ cột DE&I. Trong khi đó, Bảo Việt dẫn đầu toàn mẫu ở SDG 17 Quan hệ đối tác bền vững (53,32 điểm) nhờ áp dụng chuẩn Báo cáo Tích hợp quốc tế IIRC và bảo hiểm vi mô. SSI cũng tập trung cao vào SDG 17 và SDG 9 thông qua việc tư vấn phát hành trái phiếu xanh và thẩm định đầu tư ESG.",
        "highlights": "PNJ cao nhất SDG 5 (40,70đ - dẫn đầu toàn mẫu); BVH dẫn đầu SDG 17 (53,32đ) và SDG 9 (51,46đ).",
        "qa": "Thầy cô có thể hỏi: 'Tại sao điểm môi trường của SSI và BVH lại không cao?': Trả lời: Do mô hình văn phòng tài chính không có nhà máy phát thải trực tiếp, việc điểm môi trường ở mức vừa phải và điểm kinh tế/đối tác ở mức đỉnh cao là hoàn toàn chính xác với bản chất ngành."
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
    r_cap = p_cap.add_run("Hình 3: Xu hướng điểm số 6 nhóm SDG qua các năm (Trung bình toàn mẫu & 7 doanh nghiệp)")
    r_cap.font.size = Pt(8.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    add_card(slide, 7.8, 1.45, 4.733, 5.4)
    tb_r = slide.shapes.add_textbox(Inches(8.0), Inches(1.6), Inches(4.333), Inches(5.1))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    r0 = p0.add_run("BƯỚC NGOẶT CHÍNH SÁCH 2020–2025\n")
    r0.font.bold = True
    r0.font.color.rgb = C_NAVY_PRIMARY
    r0.font.size = Pt(12)
    
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
    """Slide 11: Kết quả 5 - Sắc thái Cảm xúc & SO SÁNH TRỰC TIẾP VỚI PAPER GỐC."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 5: SẮC THÁI CẢM XÚC & ĐỐI CHUẨN PAPER GỐC", "KẾT QUẢ THỰC NGHIỆM", 11)
    
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
    r_cap = p_cap.add_run("Hình 4 & 5: Phân phối phân cực cảm xúc PhoBERT & Cơ cấu theo doanh nghiệp")
    r_cap.font.size = Pt(8.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    add_card(slide, 7.2, 1.45, 5.333, 5.4)
    tb_r = slide.shapes.add_textbox(Inches(7.4), Inches(1.6), Inches(4.933), Inches(5.1))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    r0 = p0.add_run("SO SÁNH CẢM XÚC VỚI KANG & KIM (2022)\n")
    r0.font.bold = True
    r0.font.color.rgb = C_NAVY_PRIMARY
    r0.font.size = Pt(11.5)
    
    body = (
        "• BẢNG SO SÁNH CƠ CẤU CẢM XÚC:\n"
        "  - Paper gốc (DistilBERT 2 lớp): Tích cực ~78%, Tiêu cực ~15%.\n"
        "  - Nghiên cứu Việt Nam (PhoBERT 3 lớp):\n"
        "    + Tích cực (Positive): 53,87% (51.966 câu)\n"
        "    + Trung tính (Neutral): 32,87% (31.706 câu)\n"
        "    + Tiêu cực (Negative): 13,26% (12.789 câu)\n\n"
        "• KẾT LUẬN VỀ BÁO CÁO DOANH NGHIỆP TIẾNG VIỆT:\n"
        "  1. Tính thận trọng & Kỹ thuật: Sự xuất hiện của 32,87% câu Trung tính phản ánh các doanh nghiệp Việt Nam trình bày rất nhiều câu số liệu đo lường kỹ thuật khách quan (kWh điện, m3 nước, tấn bùn thải).\n"
        "  2. Thiên lệch Lạc quan có hệ thống: Dù đã tách lớp trung tính, tỷ lệ tích cực vẫn gấp 4,06 lần tỷ lệ tiêu cực, khẳng định báo cáo tại Việt Nam cũng chịu sự chi phối mạnh mẽ của Lý thuyết Quản trị Ấn tượng (Impression Management) và Hiệu ứng Pollyanna."
    )
    p_b = tf_r.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.3)
    p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "So sánh cơ cấu cảm xúc giữa DistilBERT 2 lớp của bài gốc và PhoBERT 3 lớp của đề tài, rút ra kết luận về tính trung tính và thiên lệch lạc quan của báo cáo tiếng Việt.",
        "script": "Kính thưa các thầy cô, tại Slide này chúng em so sánh trực tiếp cơ cấu cảm xúc: Bài báo gốc của Kang và Kim dùng mô hình 2 lớp nhị phân cho ra tỷ lệ tích cực tới 78%. Trong khi đó, mô hình PhoBERT 3 lớp của chúng em tại Việt Nam bóc tách được 32,87% câu Trung tính (mô tả số liệu kỹ thuật khách quan). Tuy nhiên, câu Tích cực vẫn chiếm tới 53,87% so với chỉ 13,26% câu Tiêu cực. Điều này dẫn đến hai kết luận khoa học quan trọng: Thứ nhất, báo cáo tiếng Việt có tính kỹ thuật rất cao nên việc bổ sung lớp trung tính là hoàn toàn chính xác. Thứ hai, hiện tượng thiên lệch lạc quan và quản trị ấn tượng vẫn tồn tại sâu sắc tại Việt Nam, doanh nghiệp vẫn chuộng dùng ngôn từ tích cực để làm đẹp hình ảnh.",
        "highlights": "So sánh cảm xúc: Bài gốc 2 lớp (Pos ~78%) vs Việt Nam 3 lớp (Pos 53,87%, Neu 32,87%, Neg 13,26%).",
        "qa": "Thầy cô có thể hỏi: 'Tại sao câu trung tính lại chiếm tới gần 1/3?': Trả lời: Báo cáo phát triển bền vững theo chuẩn GRI đòi hỏi rất nhiều bảng biểu và câu mô tả định lượng (ví dụ: 'Năm qua công ty tiêu thụ 12 triệu kWh điện'). Những câu này hoàn toàn không có cảm xúc khen hay chê, nên tỷ lệ 32,87% trung tính là hoàn toàn phản ánh trung thực bản chất dữ liệu."
    })


def build_slide_12_result6_sentiment_ratio(prs):
    """Slide 12: Kết quả 6 - Tỷ số Cảm xúc Pos/Neg Ratio & ĐỐI CHUẨN TỶ LỆ."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 6: TỶ SỐ CẢM XÚC POS/NEG RATIO & ĐỐI CHUẨN TỶ LỆ", "KẾT QUẢ THỰC NGHIỆM", 12)
    
    add_card(slide, 0.8, 1.45, 6.0, 5.4)
    fig_path = FIGURES_DIR / "sentiment_ratio.png"
    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(0.95), Inches(1.6), width=Inches(5.7), height=Inches(4.8))

    tb_cap = slide.shapes.add_textbox(Inches(0.95), Inches(6.45), Inches(5.7), Inches(0.35))
    p_cap = tb_cap.text_frame.paragraphs[0]
    p_cap.alignment = PP_ALIGN.CENTER
    r_cap = p_cap.add_run("Hình 6: Diễn biến Tỷ số Pos/Neg Ratio theo năm giữa 7 doanh nghiệp")
    r_cap.font.size = Pt(8.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = C_TEXT_MUTED

    add_card(slide, 7.0, 1.45, 5.533, 5.4)
    tb_r = slide.shapes.add_textbox(Inches(7.2), Inches(1.6), Inches(5.133), Inches(5.1))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    r0 = p0.add_run("TỶ SỐ POS/NEG: VIỆT NAM (4,06x) VS PAPER GỐC (5,20x)\n")
    r0.font.bold = True
    r0.font.color.rgb = C_NAVY_PRIMARY
    r0.font.size = Pt(11.5)
    
    body = (
        "• Đối chuẩn Tỷ số Pos/Neg với nghiên cứu gốc:\n"
        "  - Paper gốc Kang & Kim (2022): Tỷ số Pos/Neg bình quân toàn cầu đạt ~5,20 lần.\n"
        "  - Doanh nghiệp Việt Nam: Tỷ số Pos/Neg bình quân đạt 4,06 lần (VNM đạt cao nhất 5,38 lần).\n"
        "  - Nhận xét: Doanh nghiệp Việt Nam có mức độ thiên lệch lạc quan tương đương nhưng thận trọng hơn một chút so với các tập đoàn đa quốc gia phương Tây.\n\n"
        "• Vinamilk (VNM) — Ổn định ở mức cao (5,38 lần):\n"
        "  - Duy trì từ 4,5 đến 6,8 lần suốt 6 năm; phong cách truyền thông phát triển bền vững chuyên nghiệp, chau chuốt và định hướng thành tựu cao.\n\n"
        "• PNJ — Bằng chứng về độ nhạy cảm xúc trước cú sốc Covid-19 (2022):\n"
        "  - Năm 2022: Tỷ số tụt xuống mức kỷ lục 1,21 lần (317 câu tích cực vs 261 câu tiêu cực), phản ánh khó khăn đóng cửa mạng lưới bán lẻ tại TP.HCM.\n"
        "  - Năm 2023: Tỷ số phục hồi mạnh mẽ lên 4,23 lần khi kinh doanh khởi sắc.\n\n"
        "• Kết luận: Tỷ số Pos/Neg là 'nhiệt kế' nhạy bén đo lường mức độ trung thực của báo cáo trước biến cố thực tế."
    )
    p_b = tf_r.add_paragraph()
    p_b.text = body
    p_b.font.size = Pt(9.3)
    p_b.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Đối chuẩn tỷ số Pos/Neg (4.06x vs 5.20x) và phân tích độ nhạy của mô hình qua case study PNJ 2022.",
        "script": "Kính thưa Hội đồng, Hình 6 thể hiện Tỷ số Cảm xúc Pos/Neg. Khi so sánh với bài báo gốc của Kang và Kim (đạt 5,2 lần), các doanh nghiệp Việt Nam đạt bình quân 4,06 lần. Điều này cho thấy văn phong báo cáo tại Việt Nam có mức độ thiên lệch tương tự nhưng có phần thận trọng hơn. Vinamilk là doanh nghiệp có tỷ số ổn định và cao nhất với 5,38 lần. Tuy nhiên, minh chứng rõ nhất cho độ nhạy của PhoBERT là trường hợp PNJ năm 2022: Tỷ số tụt xuống 1,21 lần do báo cáo mô tả chân thực các khó khăn của dịch Covid-19 tại TP.HCM, sau đó bật tăng lên 4,23 lần vào năm 2023 khi phục hồi kinh doanh.",
        "highlights": "Đối chuẩn Tỷ số Pos/Neg: Việt Nam đạt 4,06 lần vs Paper gốc đạt ~5,20 lần. Case study PNJ 2022 (1,21x -> 4,23x).",
        "qa": "Thầy cô có thể hỏi: 'Tại sao tỷ số Pos/Neg ở Việt Nam lại thấp hơn bài báo gốc (4.06 vs 5.20)?': Trả lời: Do mô hình PhoBERT của chúng ta có lớp Trung tính giúp tách riêng các câu số liệu kỹ thuật, trong khi mô hình nhị phân của Kang & Kim có xu hướng đẩy một phần các câu trung tính vào nhóm tích cực, làm tỷ số của họ bị đội lên cao hơn."
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
        r0 = p0.add_run(t_title + "\n")
        r0.font.bold = True
        r0.font.color.rgb = col
        r0.font.size = Pt(11)
        
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
        r0 = p0.add_run(v_title + "\n")
        r0.font.bold = True
        r0.font.color.rgb = col
        r0.font.size = Pt(10.5)
        
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
        run = p.add_run(h)
        run.font.name = FONT_MAIN
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = C_WHITE

    matrix = [
        ["1. Phạm vi ngôn ngữ & Đối tượng",
         "• Chỉ áp dụng trên văn bản tiếng Anh.\n• Khảo sát 3.529 báo cáo toàn cầu từ Corporate Register.",
         "• Tích hợp NLP đa ngữ (tiếng Việt & tiếng Anh).\n• Khảo sát chuyên sâu 42 báo cáo (96.461 câu) của 7 tập đoàn VN."],
        ["2. Ngữ liệu mục tiêu SDG",
         "• Ngữ liệu tiếng Anh trích từ báo cáo Liên Hợp Quốc.",
         "• Xây dựng bộ ngữ liệu đối sánh chuẩn song ngữ 169 mục tiêu cụ thể của LHQ (1.032 câu chuẩn EN/VI)."],
        ["3. Tiền xử lý & Khôi phục dữ liệu",
         "• Đọc PDF thông thường, bỏ qua các tài liệu lỗi/scan.",
         "• Tích hợp OCR (Tesseract vie+eng) phục hồi văn bản từ các trang PDF scan hình ảnh phức tạp (PNJ 2022)."],
        ["4. Mô hình Phân tích Cảm xúc",
         "• DistilBERT phân loại nhị phân 2 lớp (Pos ~78%, Neg ~15%).\n• Tỷ số Pos/Neg ~ 5,20 lần (ép câu số liệu vào nhãn cảm xúc).",
         "• Nâng cấp PhoBERT 3 lớp (Pos 53,87%, Trung tính 32,87%, Neg 13,26%).\n• Tỷ số Pos/Neg = 4,06 lần (bảo toàn câu số liệu kỹ thuật)."],
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
            run = p.add_run(val)
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
    r1 = p1.add_run("CƠ QUAN QUẢN LÝ\n(UBCKNN & SỞ GDCK)\n\n")
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY_PRIMARY
    r1.font.size = Pt(12)
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
    r2 = p2.add_run("DOANH NGHIỆP NIÊM YẾT\n(BAN ĐIỀU HÀNH & HĐQT)\n\n")
    r2.font.bold = True
    r2.font.color.rgb = C_BLUE_ACCENT
    r2.font.size = Pt(12)
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
    r3 = p3.add_run("NHÀ ĐẦU TƯ & KIỂM TOÁN\n(QUẢN LÝ QUỸ & ĐỊNH CHẾ)\n\n")
    r3.font.bold = True
    r3.font.color.rgb = C_GREEN_EMERALD
    r3.font.size = Pt(12)
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
    r_l = p_l.add_run("HẠN CHẾ HIỆN TẠI CỦA NGHIÊN CỨU\n\n")
    r_l.font.bold = True
    r_l.font.color.rgb = C_NAVY_PRIMARY
    r_l.font.size = Pt(13)
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
    r_r = p_r.add_run("HƯỚNG PHÁT TRIỂN TIẾP THEO (GENAI & AGENTS)\n\n")
    r_r.font.bold = True
    r_r.font.color.rgb = C_GREEN_EMERALD
    r_r.font.size = Pt(13)
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
    r_t = p_t.add_run("TỔNG KẾT 4 THÔNG ĐIỆP CỐT LÕI CỦA ĐỀ TÀI")
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
        r0 = p0.add_run(c_head + "\n")
        r0.font.bold = True
        r0.font.color.rgb = C_GOLD_ACCENT
        r0.font.size = Pt(11)
        
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
    r_qa1 = p_qa1.add_run("TRÂN TRỌNG CẢM ƠN QUÝ THẦY CÔ TRONG HỘI ĐỒNG KHOA HỌC!\n")
    r_qa1.font.name = FONT_HEADING
    r_qa1.font.size = Pt(15)
    r_qa1.font.bold = True
    r_qa1.font.color.rgb = C_GOLD_ACCENT
    
    p_qa2 = tf_qa.add_paragraph()
    p_qa2.alignment = PP_ALIGN.CENTER
    r_qa2 = p_qa2.add_run("Nhóm nghiên cứu rất mong nhận được các câu hỏi và ý kiến đóng góp quý báu từ Quý Thầy Cô.\n(Tác giả: Lê Đan Sơn, Dương Thị Hoàn — 2026)")
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
    print("BẮT ĐẦU CẬP NHẬT SLIDES BÁO CÁO HỘI ĐỒNG (PHIÊN BẢN NÂNG CẤP V2)...")
    print("=" * 80)
    
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    print("[1/18] Slide 1: Trang Tiêu đề & Thông tin Tác giả...")
    build_slide_01_title(prs)
    
    print("[2/18] Slide 2: Đặt vấn đề & Bối cảnh Thể chế Việt Nam...")
    build_slide_02_context(prs)
    
    print("[3/18] Slide 3: Kết quả Paper gốc Kang & Kim (2022) & 4 Vấn đề Đề tài Khắc phục...")
    build_slide_03_original_paper(prs)
    
    print("[4/18] Slide 4: Khung Phương pháp luận 5 bước (Pipeline)...")
    build_slide_04_pipeline(prs)
    
    print("[5/18] Slide 5: Mẫu Dữ liệu Thực nghiệm (7 Doanh nghiệp, 42 Báo cáo)...")
    build_slide_05_sample(prs)
    
    print("[6/18] Slide 6: Kết quả 1 - Phân phối Tương đồng SDG & Đối chuẩn Định lượng Paper Gốc...")
    build_slide_06_result1_similarity(prs)
    
    print("[7/18] Slide 7: Kết quả 2 - Cấu trúc 6 Nhóm SDG qua Heatmap & Quy luật Toàn cầu...")
    build_slide_07_result2_heatmap(prs)
    
    print("[8/18] Slide 8: Đặc thù Ngành Sản xuất & Năng lượng kèm BIỂU ĐỒ MINH HỌA (VNM, VCS, PAN, PLX)...")
    build_slide_08_result3_companies_p1(prs)
    
    print("[9/18] Slide 9: Đặc thù Ngành Tài chính & Bán lẻ kèm BIỂU ĐỒ MINH HỌA (PNJ, BVH, SSI)...")
    build_slide_09_result3_companies_p2(prs)
    
    print("[10/18] Slide 10: Kết quả 4 - Xu hướng Dịch chuyển Chuỗi Thời gian (2020–2025)...")
    build_slide_10_result4_trends(prs)
    
    print("[11/18] Slide 11: Kết quả 5 - Sắc thái Cảm xúc & Đối chuẩn Paper Gốc (PhoBERT 3 lớp vs DistilBERT 2 lớp)...")
    build_slide_11_result5_sentiment(prs)
    
    print("[12/18] Slide 12: Kết quả 6 - Tỷ số Cảm xúc Pos/Neg Ratio: Việt Nam (4,06x) vs Paper Gốc (5,20x)...")
    build_slide_12_result6_sentiment_ratio(prs)
    
    print("[13/18] Slide 13: Thảo luận Chuyên sâu - Doanh nghiệp 'Nói nhiều về gì'...")
    build_slide_13_discussion_talk_heavy(prs)
    
    print("[14/18] Slide 14: Thảo luận Chuyên sâu - Doanh nghiệp 'Ít nói về gì' (Né tránh)...")
    build_slide_14_discussion_rarely_talk(prs)
    
    print("[15/18] Slide 15: Đóng góp Học thuật & So sánh Đối chuẩn Chi tiết...")
    build_slide_15_comparison(prs)
    
    print("[16/18] Slide 16: Hàm ý Thực tiễn & Đề xuất Chính sách...")
    build_slide_16_implications(prs)
    
    print("[17/18] Slide 17: Hạn chế của Đề tài & Hướng Phát triển Tương lai (GenAI)...")
    build_slide_17_limitations_future(prs)
    
    print("[18/18] Slide 18: Tổng kết 4 Thông điệp Cốt lõi & Phiên Hỏi đáp (Q&A)...")
    build_slide_18_conclusion(prs)
    
    prs.save(str(OUTPUT_PPTX))
    print(f"\n=> Đã lưu thành công bộ slide nâng cấp tại: {OUTPUT_PPTX}")
    
    PAPERS_DIR.mkdir(parents=True, exist_ok=True)
    dest_copy = PAPERS_DIR / "bao_cao_nghien_cuu_sdg_vietnam.pptx"
    shutil.copyfile(OUTPUT_PPTX, dest_copy)
    print(f"=> Đã sao chép vào: {dest_copy}")
    
    file_size_mb = OUTPUT_PPTX.stat().st_size / (1024 * 1024)
    print(f"=> Kích thước tệp: {file_size_mb:.2f} MB")
    print("=" * 80)
    print("HOÀN TẤT THÀNH CÔNG V2!")
    print("=" * 80)


if __name__ == "__main__":
    main()
