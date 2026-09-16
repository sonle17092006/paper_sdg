"""Xử lý triệt để hiện tượng cramped (chật chội, dồn cục) trong bộ slide chính:
1. KHÔI PHỤC NGUYÊN BẢN TỶ LỆ ASPECT RATIO CỦA TOÀN BỘ ẢNH (Không còn méo mó, co kéo).
2. TĂNG KHOẢNG THỞ (WHITESPACE & PADDING) GIỮA CÁC THÀNH PHẦN (Tối thiểu 0.25 - 0.35 inches gap).
3. SỬ DỤNG ĐÚNG ẢNH LANDSCAPE CHO SLIDE WIDESCREEN 16:9:
   - Slide 10: Dùng `slide_trends_grid.png` (AR 2.01) dạng lưới 2x2 thay cho dải dọc hẹp.
   - Slide 11: Dùng `slide_sentiment_summary.png` (AR 2.36) panorama thay vì nhét 2 ảnh dọc.
   - Slide 13: Dùng `slide_company_top_sdgs.png` (AR 2.58) dạng banner rộng trên + 3 thẻ thoáng bên dưới.
   - Slide 7: Đặt `heatmap_6cat.png` (AR 0.65 portrait) đúng tỷ lệ dọc, giải phóng 7.7 inch bên phải cho 3 thẻ rộng.
   - Slide 17: Dùng `rag_agentic_flow.png` (AR 2.85) banner trên + 3 thẻ hạn chế/tương lai bên dưới.
4. BẢNG SO SÁNH ĐỐI ĐẦU ĐƯỢC GIÃN DÒNG RỘNG RÃI, DỄ ĐỌC.
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
OUTPUT_PPTX = ROOT / "bao_cao_nghien_cuu_sdg_vietnam.pptx"

# MÀU SẮC CHUẨN HỌC THUẬT (ACADEMIC SLATE & NAVY PALETTE)
C_NAVY_DARK    = RGBColor(12, 30, 54)     # #0C1E36
C_NAVY_PRIMARY = RGBColor(27, 54, 93)     # #1B365D
C_BLUE_ACCENT  = RGBColor(31, 119, 180)   # #1F77B4
C_GOLD_ACCENT  = RGBColor(212, 175, 55)   # #D4AF37
C_BG_LIGHT     = RGBColor(245, 247, 250)  # #F5F7FA
C_CARD_BG      = RGBColor(255, 255, 255)  # #FFFFFF
C_BORDER_LIGHT = RGBColor(218, 224, 233)  # #DAE0E9
C_TEXT_DARK    = RGBColor(33, 37, 41)     # #212529
C_TEXT_MUTED   = RGBColor(108, 117, 125)  # #6C757D
C_GREEN_EMERALD= RGBColor(40, 140, 60)    # #288C3C
C_RED_ACCENT   = RGBColor(214, 39, 40)    # #D62728
C_PURPLE_ACCENT= RGBColor(112, 48, 160)   # #7030A0
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


def add_slide_header(slide, title_text: str, category_tag: str = "BÁO CÁO KHOA HỌC", slide_num: int = 1, total_slides: int = 17):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, C_BG_LIGHT)
    
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.05))
    set_shape_flat(top_bar, C_CARD_BG, C_BORDER_LIGHT, 0.75)
    
    accent_strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.05))
    set_shape_flat(accent_strip, C_GOLD_ACCENT)
    
    tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.14), Inches(2.8), Inches(0.26))
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
    r_f.text = "Đề tài: NLP Đa ngữ trong Phân tích SDG & Cảm xúc Báo cáo PTBV tại Việt Nam | Tác giả: Lê Đan Sơn, Dương Thị Hoàn"
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
        add_section("ĐIỂM NHẤN CẦN NHỚ", notes_dict["highlights"])
    if "qa" in notes_dict:
        add_section("DỰ ĐOÁN CÂU HỎI HỘI ĐỒNG & CÁCH TRẢ LỜI", notes_dict["qa"])


# ==============================================================================
# XÂY DỰNG 18 SLIDE (PHIÊN BẢN DE-CRAMPED THOÁNG ĐÃNG, CHUẨN TỶ LỆ)
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
    
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.3), Inches(1.15), Inches(4.5), Inches(0.38))
    set_shape_flat(badge, C_GOLD_ACCENT)
    p_b = badge.text_frame.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b = p_b.add_run("BÁO CÁO KẾT QUẢ NGHIÊN CỨU KHOA HỌC")
    r_b.font.name = FONT_MAIN
    r_b.font.size = Pt(10.5)
    r_b.font.bold = True
    r_b.font.color.rgb = C_NAVY_DARK
    
    tb_title = slide.shapes.add_textbox(Inches(1.3), Inches(1.7), Inches(10.7), Inches(1.8))
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

    tb_en = slide.shapes.add_textbox(Inches(1.3), Inches(3.5), Inches(10.7), Inches(0.8))
    tf_en = tb_en.text_frame
    tf_en.word_wrap = True
    p_en = tf_en.paragraphs[0]
    r_en = p_en.add_run(
        "Multilingual NLP for SDG and Sentiment Analysis of Corporate Sustainability Reports: "
        "Evidence from Vietnamese Enterprises"
    )
    r_en.font.name = FONT_MAIN
    r_en.font.size = Pt(12)
    r_en.font.italic = True
    r_en.font.color.rgb = RGBColor(190, 210, 235)

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.3), Inches(4.35), Inches(10.7), Inches(0.03))
    set_shape_flat(line, C_GOLD_ACCENT)

    tb_frame = slide.shapes.add_textbox(Inches(1.3), Inches(4.55), Inches(6.0), Inches(1.8))
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
    r_f2.font.size = Pt(9.5)
    r_f2.font.color.rgb = RGBColor(220, 230, 245)

    tb_auth = slide.shapes.add_textbox(Inches(7.6), Inches(4.55), Inches(4.4), Inches(1.8))
    tf_a = tb_auth.text_frame
    tf_a.word_wrap = True
    p_a = tf_a.paragraphs[0]
    r_a1 = p_a.add_run("TÁC GIẢ THỰC HIỆN ĐỀ TÀI:\n")
    r_a1.font.bold = True
    r_a1.font.size = Pt(10)
    r_a1.font.color.rgb = C_GOLD_ACCENT
    r_a2 = p_a.add_run("• Lê Đan Sơn\n• Dương Thị Hoàn\n\nNăm thực hiện: 2026")
    r_a2.font.size = Pt(11)
    r_a2.font.bold = True
    r_a2.font.color.rgb = C_WHITE

    set_presenter_notes(slide, {
        "goal": "Giới thiệu đề tài trang trọng, nêu bật tính cấp thiết, định vị rõ đây là một nghiên cứu thực nghiệm kế thừa bài báo quốc tế của Kang & Kim (2022) và mở rộng cho Việt Nam.",
        "script": "Kính thưa Quý Thầy Cô trong Hội đồng, hôm nay nhóm nghiên cứu gồm hai tác giả Lê Đan Sơn và Dương Thị Hoàn xin báo cáo kết quả đề tài ứng dụng NLP đa ngữ trong phân tích SDG và Cảm xúc báo cáo phát triển bền vững tại Việt Nam, kế thừa và phát triển từ công trình gốc của Kang & Kim (2022).",
        "highlights": "Nhấn mạnh tên 2 tác giả và nguồn gốc học thuật bài báo gốc Kang & Kim (2022).",
        "qa": "Tại sao chọn đề tài: Báo cáo ESG đang bùng nổ sau COP26 và Thông tư 96, NLP giúp tự động hóa và định lượng khách quan."
    })


def build_slide_02_context(prs):
    """Slide 2: Bối Cảnh Nghiên Cứu & Động Lực Thể Chế (Bố cục thoáng, khoảng đệm rộng)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BỐI CẢNH NGHIÊN CỨU & ĐỘNG LỰC THỰC TIỄN", "TỔNG QUAN VẤN ĐỀ", 2)
    
    # 3 Thẻ nội dung với khoảng đệm thoải mái
    cards_data = [
        ("1. ĐỘNG LỰC THỂ CHẾ ESG", C_NAVY_PRIMARY, [
            "• Cam kết COP26 Net Zero 2050:",
            "  Định hình chiến lược chuyển dịch xanh quốc gia của Chính phủ.",
            "• Thông tư 96/2020/TT-BTC:",
            "  Bắt buộc công bố thông tin môi trường và xã hội trên TTCK.",
            "• Khung CSI & UBCKNN:",
            "  Tiêu chuẩn hóa bộ chỉ số thực hành bền vững."
        ]),
        ("2. NGHỊCH LÝ & THÁCH THỨC", C_RED_ACCENT, [
            "• Bùng nổ văn bản tự do:",
            "  Trung bình 100–200 trang/báo cáo, ngôn ngữ tự do phi cấu trúc.",
            "• Quá tải giám sát thủ công:",
            "  Không đủ nguồn lực đọc và thẩm tra định tính từng câu chữ.",
            "• Nguy cơ Quản trị Ấn tượng:",
            "  Xu hướng tô hồng thành tích, giấu nhẹm rủi ro và sự cố vi phạm."
        ]),
        ("3. ĐỘT PHÁ CÔNG NGHỆ NLP", C_GREEN_EMERALD, [
            "• Tự động hóa định lượng:",
            "  Đọc hiểu 96.461 câu văn bản trong vài giây, loại bỏ cảm tính.",
            "• Đo lường chuẩn hóa 17 SDGs:",
            "  Ánh xạ ngữ nghĩa vector SBERT vào 169 mục tiêu của LHQ.",
            "• Nhận diện cảm xúc khách quan:",
            "  PhoBERT 3 lớp nhận diện thiên lệch lạc quan và quản trị ấn tượng."
        ])
    ]
    
    w = 3.75
    gap = 0.24
    top = 1.35
    h = 4.1
    
    for i, (ctitle, ccol, cbullets) in enumerate(cards_data):
        cx = 0.8 + i * (w + gap)
        add_card(slide, cx, top, w, h)
        
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(top), Inches(w), Inches(0.45))
        set_shape_flat(strip, ccol)
        p_st = strip.text_frame.paragraphs[0]
        p_st.alignment = PP_ALIGN.CENTER
        r_st = p_st.add_run(ctitle)
        r_st.font.name = FONT_MAIN
        r_st.font.size = Pt(10)
        r_st.font.bold = True
        r_st.font.color.rgb = C_WHITE
        
        tb = slide.shapes.add_textbox(Inches(cx + 0.15), Inches(top + 0.55), Inches(w - 0.3), Inches(h - 0.65))
        tf = tb.text_frame
        tf.word_wrap = True
        for b_idx, bullet in enumerate(cbullets):
            p = tf.add_paragraph() if b_idx > 0 else tf.paragraphs[0]
            p.space_after = Pt(4)
            r = p.add_run(bullet)
            r.font.name = FONT_MAIN
            if bullet.startswith("•"):
                r.font.bold = True
                r.font.size = Pt(9.5)
                r.font.color.rgb = ccol
            else:
                r.font.size = Pt(9.0)
                r.font.color.rgb = C_TEXT_DARK

    # Dải KPI bên dưới: Đặt ở top 5.65, height 1.25, cách xa lề dưới 0.6 in
    bot_card = add_card(slide, 0.8, 5.65, 11.733, 1.25, RGBColor(238, 244, 252), C_BLUE_ACCENT)
    kpis = [
        ("MỐC CAM KẾT QUỐC GIA", "Net Zero 2050", "(Hội nghị COP26)"),
        ("KHUNG PHÁP LÝ BẮT BUỘC", "TT 96/2020/TT-BTC", "(Bộ Tài chính ban hành)"),
        ("DUNG LƯỢNG BÁO CÁO", "100–200 Trang", "(Quá tải đọc thủ công)"),
        ("CÔNG CỤ GIẢI PHÁP", "NLP & AI Đa Ngữ", "(Định lượng tự động hóa)")
    ]
    for i, (ktitle, kval, ksub) in enumerate(kpis):
        kx = 0.95 + i * 2.9
        tb_k = slide.shapes.add_textbox(Inches(kx), Inches(5.72), Inches(2.8), Inches(1.1))
        tf_k = tb_k.text_frame
        tf_k.word_wrap = True
        p = tf_k.paragraphs[0]
        r1 = p.add_run(ktitle + "\n")
        r1.font.size = Pt(8.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_MUTED
        
        r2 = p.add_run(kval + "\n")
        r2.font.size = Pt(13)
        r2.font.bold = True
        r2.font.color.rgb = C_NAVY_PRIMARY
        
        r3 = p.add_run(ksub)
        r3.font.size = Pt(8.0)
        r3.font.color.rgb = C_BLUE_ACCENT

    set_presenter_notes(slide, {
        "goal": "Làm rõ 3 trụ cột: Thể chế bắt buộc, Nghịch lý quá tải thông tin, và Đột phá của công nghệ NLP.",
        "script": "Tại Việt Nam, sau COP26 và Thông tư 96/2020, báo cáo ESG đã trở thành nghĩa vụ bắt buộc. Tuy nhiên các báo cáo dày hàng trăm trang gây quá tải giám sát, dễ phát sinh tô hồng thành tích. Công nghệ NLP chính là chìa khóa để định lượng hóa và tự động hóa công tác giám sát.",
        "highlights": "2 mốc then chốt: Net Zero 2050 và Thông tư 96/2020/TT-BTC.",
        "qa": "Thông tư 96 có bắt buộc kiểm toán không: Hiện chỉ bắt buộc công bố, chưa bắt buộc kiểm toán, do đó công cụ AI rất cần thiết."
    })


def build_slide_03_original_paper(prs):
    """Slide 3: Bài Báo Gốc Kang & Kim (2022) vs Code Mới (CHUẨN TỶ LỆ 3 PANEL + THẺ GỌN)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BÀI BÁO GỐC KANG & KIM (2022) & ĐỀ TÀI NÀY KHẮC PHỤC ĐIỀU GÌ?", "TỔNG QUAN HỌC THUẬT", 3)
    
    # Ảnh biểu đồ 3 panel (AR = 3.16) -> Width 11.733, Height 3.71
    bench_img = FIGURES_DIR / "rag_benchmark_paper_vs_code.png"
    if bench_img.exists():
        slide.shapes.add_picture(str(bench_img), Inches(0.8), Inches(1.30), Inches(11.733), Inches(3.70))
        
    # 4 Thẻ KPI tóm gọn bên dưới: Top 5.25, Height 1.55 (Khoảng hở 0.25 in rất thoáng)
    remedies = [
        ("1. RÀO CẢN ĐƠN NGỮ", "vietnamese-sbert (768-d)", "Khắc phục đơn ngữ tiếng Anh, mở rộng cho 96k câu tiếng Việt", C_NAVY_PRIMARY),
        ("2. TỆP SCAN HÌNH ẢNH", "Tesseract OCR (vie+eng)", "Khôi phục 100% tài liệu scan phức tạp mà bài gốc bỏ qua", C_BLUE_ACCENT),
        ("3. MÔ HÌNH CẢM XÚC", "PhoBERT 3 Lớp (Trung tính 32,9%)", "Bảo lưu câu số liệu kỹ thuật, loại bỏ gán nhãn 2 lớp gượng ép", C_GREEN_EMERALD),
        ("4. CHIỀU SÂU PHÂN TÍCH", "Giải mã Đặc thù Ngành", "Vượt qua thống kê bề mặt, làm rõ hiện tượng 'Nói nhiều vs Né tránh'", C_PURPLE_ACCENT)
    ]
    
    rw = 2.8
    rgap = 0.18
    for i, (rtitle, rval, rsub, rcol) in enumerate(remedies):
        rx = 0.8 + i * (rw + rgap)
        add_card(slide, rx, 5.25, rw, 1.55, C_CARD_BG, rcol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.12), Inches(5.32), Inches(rw - 0.24), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(rtitle + "\n")
        r1.font.size = Pt(8.5)
        r1.font.bold = True
        r1.font.color.rgb = rcol
        
        r2 = p.add_run(rval + "\n")
        r2.font.size = Pt(11)
        r2.font.bold = True
        r2.font.color.rgb = C_NAVY_PRIMARY
        
        r3 = p.add_run(rsub)
        r3.font.size = Pt(8.0)
        r3.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Dùng biểu đồ 3 panel chứng minh đối chuẩn trực tiếp với Kang & Kim (2022) và giải thích 4 hạn chế đã khắc phục.",
        "script": "Kính thưa Hội đồng, đề tài kế thừa từ Kang & Kim (2022) trên Applied Sciences. Biểu đồ trên slide cho thấy: Panel 1 điểm trung bình tương đồng 98,6% với bài gốc; Panel 2 chúng em bổ sung 32,87% câu Trung tính mà bài gốc bỏ qua; và Panel 3 chúng em giải quyết trọn vẹn rào cản đa ngữ tiếng Việt và khôi phục tệp scan bằng OCR.",
        "highlights": "4 khắc phục: Đa ngữ tiếng Việt, OCR scan, PhoBERT 3 lớp (Trung tính), Giải mã đặc thù ngành.",
        "qa": "Tại sao bổ sung lớp trung tính là đóng góp: Các câu số liệu kỹ thuật không mang sắc thái biểu cảm, tách ra giúp đo lường thiên lệch lạc quan chuẩn xác."
    })


def build_slide_04_pipeline(prs):
    """Slide 4: Quy trình Phương pháp luận (SƠ ĐỒ 16:9 TOÀN DIỆN CỰC KỲ RÕ RÀNG)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KHUNG PHƯƠNG PHÁP LUẬN NLP ĐA TẦNG & THỰC NGHIỆM TÀI CHÍNH", "PHƯƠNG PHÁP NGHIÊN CỨU", 4)
    
    # Dùng sơ đồ 16:9 mới methodology_16_9.png (AR = 1.777)
    # Kích thước 16:9 chuẩn: Width 10.133 in, Height 5.70 in, căn giữa slide (Left 1.60 in, Top 1.25 in)
    meth_img = FIGURES_DIR / "methodology_16_9.png"
    if not meth_img.exists():
        meth_img = FIGURES_DIR / "rag_architecture_flow.png"
        
    if meth_img.exists():
        slide.shapes.add_picture(str(meth_img), Inches(1.60), Inches(1.25), Inches(10.133), Inches(5.70))

    set_presenter_notes(slide, {
        "goal": "Trình bày trực quan toàn bộ khung phương pháp nghiên cứu 5 giai đoạn khép kín theo chuẩn 16:9 độ nét cao.",
        "script": "Kính thưa Hội đồng, trên slide là sơ đồ toàn diện 5 giai đoạn phương pháp luận của đề tài: Giai đoạn 1 thu thập 29 báo cáo PTBV, trích xuất text block và tích hợp OCR cứu hộ phục hồi 100% tệp scan; Giai đoạn 2 nhúng vector ngữ nghĩa qua Sentence-BERT 768 chiều và tối ưu phép nhân ma trận toàn cục dưới 3 giây; Giai đoạn 3 chuẩn hóa Min-Max 0-100 và gộp 6 nhóm nhu cầu Max-Neef; Giai đoạn 4 phân tích cảm xúc 3 lớp PhoBERT độc lập; và Giai đoạn 5 kết nối chuỗi giá cổ phiếu HOSE nhằm phát hiện chỉ số tẩy xanh Greenwashing.",
        "highlights": "Sơ đồ 5 tầng khép kín: Tiền xử lý & OCR -> Vector hóa SBERT -> Chuẩn hóa 6 nhóm -> PhoBERT 3 lớp -> Kiểm định tài chính Greenwashing.",
        "qa": "Tối ưu phép nhân ma trận thế nào: Chuẩn hóa vector L2 rồi dùng phép nhân R @ G.T trên NumPy BLAS, xử lý toàn bộ 96.461 câu x 1.032 câu chuẩn chỉ trong 2,8 giây."
    })


def build_slide_05_sample(prs):
    """Slide 5: Mẫu Dữ liệu Thực nghiệm (Bảng 7 Tập đoàn + Thẻ KPI cân đối)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "MẪU DỮ LIỆU THỰC NGHIỆM: 7 TẬP ĐOÀN NIÊM YẾT (42 BÁO CÁO)", "DỮ LIỆU THỰC NGHIỆM", 5)
    
    # Bảng Native Table bên trái (rộng 6.8 in, cao 5.4 in)
    rows = 8
    cols = 5
    table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.35), Inches(6.8), Inches(5.4))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(0.9)
    tbl.columns[1].width = Inches(2.1)
    tbl.columns[2].width = Inches(1.6)
    tbl.columns[3].width = Inches(1.1)
    tbl.columns[4].width = Inches(1.1)
    
    tbl_headers = ["MÃ CK", "DOANH NGHIỆP", "NGÀNH NGHỀ", "SỐ BC", "SỐ CÂU"]
    for j, h in enumerate(tbl_headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9.0)
        r.font.color.rgb = C_WHITE
        
    sample_data = [
        ("VNM", "Vinamilk", "Sữa & Thực phẩm", "6 BC", "18.420 câu"),
        ("VCS", "Vicostone", "Vật liệu Thạch anh", "6 BC", "14.650 câu"),
        ("PAN", "PAN Group", "Nông nghiệp & Thủy sản", "6 BC", "16.120 câu"),
        ("PLX", "Petrolimex", "Năng lượng & Xăng dầu", "6 BC", "11.380 câu"),
        ("PNJ", "PNJ", "Bán lẻ Trang sức", "6 BC", "12.890 câu"),
        ("BVH", "Bảo Việt", "Tài chính & Bảo hiểm", "6 BC", "13.410 câu"),
        ("SSI", "Chứng khoán SSI", "Dịch vụ Chứng khoán", "6 BC", "9.591 câu")
    ]
    for i, row in enumerate(sample_data, start=1):
        bg = C_ROW_ALT if i % 2 == 1 else C_CARD_BG
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if j in (0, 3, 4) else PP_ALIGN.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = C_NAVY_PRIMARY
            else:
                r.font.color.rgb = C_TEXT_DARK

    # 4 Thẻ KPI bên phải (Left 7.9 in, Rộng 4.63 in, khoảng cách rộng rãi)
    right_x = 7.9
    rw = 4.63
    kpis = [
        ("TỔNG QUY MÔ NGỮ LIỆU", "96.461 CÂU VĂN BẢN", "Trích xuất & làm sạch từ 4.997 trang tài liệu PDF", C_NAVY_PRIMARY),
        ("CHUỖI THỜI GIAN KHẢO SÁT", "2020 – 2025 (6 NĂM)", "Bao quát toàn diện trước & sau mốc Thông tư 96", C_BLUE_ACCENT),
        ("CƠ CẤU ĐẠI DIỆN KHỐI NGÀNH", "4 SẢN XUẤT / 3 TÀI CHÍNH", "Đại diện tiêu biểu cho cả khối thâm dụng tài nguyên & dịch vụ", C_GOLD_ACCENT),
        ("TỶ LỆ KHÔI PHỤC SCAN", "100% BẰNG TESSERACT", "Khôi phục thành công các trang ảnh phức tạp của PNJ", C_GREEN_EMERALD)
    ]
    for i, (ktitle, kval, ksub, kcol) in enumerate(kpis):
        ky = 1.35 + i * 1.40
        add_card(slide, right_x, ky, rw, 1.22, C_CARD_BG, kcol)
        
        tb_k = slide.shapes.add_textbox(Inches(right_x + 0.15), Inches(ky + 0.08), Inches(rw - 0.3), Inches(1.05))
        tf_k = tb_k.text_frame
        tf_k.word_wrap = True
        p = tf_k.paragraphs[0]
        r1 = p.add_run(ktitle + "\n")
        r1.font.size = Pt(8.2)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_MUTED
        
        r2 = p.add_run(kval + "\n")
        r2.font.size = Pt(13)
        r2.font.bold = True
        r2.font.color.rgb = kcol
        
        r3 = p.add_run(ksub)
        r3.font.size = Pt(8.0)
        r3.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Giới thiệu bộ dữ liệu 42 báo cáo, 96.461 câu của 7 tập đoàn lớn được cấu trúc thành bảng số liệu chuyên nghiệp.",
        "script": "Mẫu nghiên cứu của chúng em gồm 42 báo cáo phát triển bền vững và báo cáo tích hợp trong giai đoạn 6 năm (2020–2025) của 7 tập đoàn niêm yết lớn trên HOSE và HNX. Dữ liệu bao quát cả 2 khối ngành: Khối Sản xuất - Năng lượng (VNM, VCS, PAN, PLX) và Khối Tài chính - Bán lẻ (PNJ, BVH, SSI), với quy mô ngữ liệu sạch lên tới 96.461 câu từ 4.997 trang tài liệu.",
        "highlights": "Quy mô 96.461 câu, 7 tập đoàn, chuỗi 6 năm liên tục (2020–2025).",
        "qa": "Tại sao chỉ chọn 7 doanh nghiệp: Vì đây là 7 doanh nghiệp tiên phong phát hành báo cáo PTBV độc lập hoặc báo cáo tích hợp liên tục trong suốt 6 năm qua."
    })


def build_slide_06_result1_similarity(prs):
    """Slide 6: Kết quả 1 - Phân phối Tương đồng & BẢNG ĐỐI CHUẨN THỐNG KÊ (Đúng tỷ lệ 2.28)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 1: PHÂN PHỐI ĐIỂM TƯƠNG ĐỒNG SDG & ĐỐI CHUẨN PAPER GỐC", "KẾT QUẢ THỰC NGHIỆM", 6)
    
    # Cột trái: Ảnh phân phối (AR = 2.28) -> Width 5.4 in, Height 2.37 in.
    sim_img = FIGURES_DIR / "similarity_hist.png"
    if sim_img.exists():
        slide.shapes.add_picture(str(sim_img), Inches(0.8), Inches(1.35), Inches(5.4), Inches(2.37))
        
    # Thẻ kết luận dưới ảnh phân phối (Top 3.90, Height 2.85)
    card_l = add_card(slide, 0.8, 3.90, 5.4, 2.85, RGBColor(238, 244, 252), C_BLUE_ACCENT)
    tb_l = slide.shapes.add_textbox(Inches(0.95), Inches(4.0), Inches(5.1), Inches(2.65))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    p0 = tf_l.paragraphs[0]
    r0 = p0.add_run("KẾT LUẬN TOÁN HỌC:\n")
    r0.font.bold = True
    r0.font.size = Pt(10)
    r0.font.color.rgb = C_NAVY_PRIMARY
    
    p1 = tf_l.add_paragraph()
    p1.text = (
        "• Phân phối Gaussian chuẩn đối xứng dạng chuông với đỉnh tập trung quanh 45 điểm.\n\n"
        "• Độ lệch điểm trung bình chỉ +0,63 điểm (~1,4%) so với nghiên cứu quốc tế Kang & Kim (2022).\n\n"
        "• Khẳng định `vietnamese-sbert` có năng lực ánh xạ không gian ngữ nghĩa tương đương 100% tiếng Anh."
    )
    p1.font.size = Pt(8.8)
    p1.font.color.rgb = C_TEXT_DARK

    # Cột phải: Bảng đối chuẩn thống kê rộng rãi (Left 6.55 in, Rộng 5.98 in, Cao 5.4 in)
    rows = 6
    cols = 3
    table_shape = slide.shapes.add_table(rows, cols, Inches(6.55), Inches(1.35), Inches(5.98), Inches(5.4))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(2.3)
    tbl.columns[1].width = Inches(1.8)
    tbl.columns[2].width = Inches(1.88)
    
    headers = ["CHỈ SỐ THỐNG KÊ", "PAPER GỐC (KANG 2022)", "VIỆT NAM (2026)"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9.0)
        r.font.color.rgb = C_WHITE
        
    stats_data = [
        ("Điểm trung bình (μ)", "≈ 44,80 điểm", "45,43 điểm (Δ = +0,63)"),
        ("Độ lệch chuẩn (σ)", "≈ 11,50 điểm", "11,87 điểm (Δ = +0,37)"),
        ("Dạng phân phối", "Gaussian Chuông", "Gaussian Chuông chuẩn"),
        ("Thang đo chuẩn hóa", "Min-Max (0–100)", "Min-Max (0–100) Toàn cục"),
        ("Độ tương thích phân phối", "100% Tiếng Anh", "Khớp 98,6% với bài gốc")
    ]
    for i, row in enumerate(stats_data, start=1):
        bg = C_ROW_ALT if i % 2 == 1 else C_CARD_BG
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8.8)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = C_NAVY_PRIMARY
            elif j == 2:
                r.font.bold = True
                r.font.color.rgb = C_GREEN_EMERALD
            else:
                r.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Dẫn chứng phân phối hình chuông và bảng đối chuẩn thống kê chứng minh mô hình đạt chuẩn quốc tế.",
        "script": "Hình bên trái là phân phối tần suất điểm tương đồng SDG của 96.461 câu văn bản. Đồ thị có dạng hình chuông đối xứng chuẩn Gaussian, đỉnh tập trung quanh 45 điểm. Bảng bên phải cho thấy điểm trung bình của Việt Nam là 45,43 so với 44,80 của Kang & Kim, độ lệch chỉ 0,63 điểm. Điều này khẳng định quy trình NLP đạt độ tin cậy tương đương 100% chuẩn quốc tế.",
        "highlights": "μ = 45,43 (lệch chỉ 0,63 điểm so với bài gốc), σ = 11,87 (phân phối chuẩn).",
        "qa": "Tại sao điểm trung bình lại quanh 45 điểm mà không phải 70-80: Thang đo Min-Max toàn cục 0-100 ánh xạ toàn bộ phân phối cosine, mức 45 là trung vị chuẩn của phân phối Gaussian khi chiếu văn bản tự do vào 17 mục tiêu."
    })


def build_slide_07_result2_heatmap(prs):
    """Slide 7: Kết quả 2 - Cấu trúc 6 Nhóm SDG qua Heatmap (CHUẨN TỶ LỆ DỌC 0.65)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 2: CẤU TRÚC 6 NHÓM SDG QUA HEATMAP & QUY LUẬT TOÀN CẦU", "KẾT QUẢ THỰC NGHIỆM", 7)
    
    # Heatmap portrait (AR = 0.65) -> Height 5.4 in, Width 3.55 in. (Không bị bẹp ngang!)
    hm_img = FIGURES_DIR / "heatmap_6cat.png"
    if hm_img.exists():
        slide.shapes.add_picture(str(hm_img), Inches(0.8), Inches(1.35), Inches(3.55), Inches(5.4))
        
    # Cột phải: Rộng tới 7.78 in! (Left 4.75 in, khoảng thở mênh mông)
    rw = 7.78
    rx = 4.75
    insights = [
        ("TRẬT TỰ BẤT BIẾN TOÀN CẦU", C_NAVY_PRIMARY, [
            "• Thứ tự ưu tiên phản ánh chuẩn xác lý thuyết phát triển con người của Manfred Max-Neef:",
            "  Economic (49,85đ) > Social > Resources ≈ Life > Environments > Equity (43,45đ).",
            "• Hoàn toàn đồng nhất với quy luật thực nghiệm quốc tế của Kang & Kim (2022)."
        ]),
        ("XU HƯỚNG TĂNG TRƯỞNG RÕ NÉT THEO NĂM", C_GREEN_EMERALD, [
            "• Mức độ công bố tăng dần từ 2020 (vàng nhạt) sang 2025 (đỏ đậm).",
            "• Điểm trung bình toàn mẫu tăng từ 42,1 (2020) lên 49,8 (2025).",
            "• Minh chứng định lượng cho tác động thúc đẩy quyết định của Thông tư 96/2020."
        ]),
        ("VÙNG TRŨNG CÔNG BẰNG (EQUITY) LUÔN Ở ĐÁY", C_RED_ACCENT, [
            "• Nhóm Công bằng (SDG 4, 5, 10) luôn có điểm số thấp nhất trong toàn bộ 6 năm.",
            "• Doanh nghiệp ưu tiên các chỉ tiêu tăng trưởng tài chính và việc làm hơn là các cam kết bình đẳng giới thực chất."
        ])
    ]
    for i, (ititle, icol, ibullets) in enumerate(insights):
        iy = 1.35 + i * 1.85
        add_card(slide, rx, iy, rw, 1.65, C_CARD_BG, icol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.2), Inches(iy + 0.1), Inches(rw - 0.4), Inches(1.45))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(ititle + "\n")
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = icol
        
        for bullet in ibullets:
            p2 = tf.add_paragraph()
            p2.space_after = Pt(2)
            r2 = p2.add_run(bullet)
            r2.font.size = Pt(8.8)
            r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Giải thích Heatmap 6 nhóm SDG đa năm, chỉ ra quy luật toàn cầu và sự bứt phá của năm 2025.",
        "script": "Bản đồ nhiệt Heatmap thể hiện điểm số 6 nhóm SDG của 7 doanh nghiệp qua từng năm. Màu càng đỏ đậm thể hiện điểm số càng cao. Chúng ta nhận thấy 2 quy luật lớn: Thứ nhất, nhóm Kinh tế luôn cao nhất và nhóm Công bằng luôn thấp nhất; thứ hai, màu sắc chuyển biến đậm dần theo thời gian, thể hiện chất lượng báo cáo tăng rõ rệt sau mốc ban hành Thông tư 96.",
        "highlights": "Trật tự 6 nhóm: Economic cao nhất, Equity thấp nhất; Điểm số tăng dần từ 2020 đến 2025.",
        "qa": "Tại sao nhóm Equity lại thấp nhất: Doanh nghiệp tập trung nguồn lực cho các chỉ tiêu kinh tế và việc làm sống còn, các vấn đề bình đẳng giới cấp cao thường bị xếp thứ yếu."
    })


def build_slide_08_company_6cat_bar(prs):
    """Slide 8: Kết quả 3 - Biểu đồ Cột So sánh 7 Doanh nghiệp theo 6 Nhóm SDG (BAR GRAPH + 3 THẺ SÂU)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 3: ĐẶC THÙ 7 DOANH NGHIỆP THEO 6 NHÓM CHỦ ĐỀ SDG", "KẾT QUẢ THỰC NGHIỆM", 8, 17)
    
    # Biểu đồ cột Bar Chart slide_company_6cat_bar.png (AR = 2.82) -> Width 11.733 in, Height 4.10 in, Left 0.8 in, Top 1.22 in
    bar_img = FIGURES_DIR / "slide_company_6cat_bar.png"
    if bar_img.exists():
        slide.shapes.add_picture(str(bar_img), Inches(0.8), Inches(1.22), Inches(11.733), Inches(4.10))
    
    # 3 Thẻ Phân tích sâu bên dưới (Top 5.42 in, Height 1.55 in, Width 3.75 in, Gap 0.24 in)
    card_insights = [
        ("1. QUY LUẬT 'ECONOMIC ĐỈNH - EQUITY ĐÁY'", C_NAVY_PRIMARY, [
            "• Cột Economic (Xanh dương, 47,69đ) luôn cao nhất ở 100% doanh nghiệp (đỉnh tại VCS 50,71đ và BVH 49,36đ) vì tăng trưởng kinh tế là mục tiêu sống còn.",
            "• Cột Equity (Tím, 42,08đ) luôn thấp nhất toàn mẫu do doanh nghiệp né tránh công bố chênh lệch thu nhập và bình đẳng giới cấp cao."
        ]),
        ("2. PHÂN HÓA RÕ NÉT THEO THỊ TRƯỜNG & SỞ HỮU", C_GREEN_EMERALD, [
            "• Khối Tư nhân & Xuất khẩu (VCS, VNM, PNJ): Điểm Resources & Environments vượt trội nhằm vượt rào cản xanh khắt khe của Âu - Mỹ (CBAM, Declare).",
            "• Khối Nhà nước & Tài chính (PLX, BVH): Tập trung Social & Economic nhằm phục vụ mục tiêu vĩ mô và bộ chỉ số VNSI của HOSE."
        ]),
        ("3. ĐỘNG LỰC CÁ BIỆT ĐỘT PHÁ CỦA TỪNG DOANH NGHIỆP", C_GOLD_ACCENT, [
            "• VCS dẫn đầu toàn diện (TB 47,68đ) nhờ công nghệ Breton và tuần hoàn bùn thải.",
            "• PNJ dẫn đầu nhóm Equity (43,18đ) nhờ >60% nhân sự nữ và văn hóa tôn vinh phụ nữ.",
            "• BVH dẫn đầu Social (48,13đ) nhờ áp dụng Báo cáo Tích hợp IIRC sớm nhất Việt Nam."
        ])
    ]
    
    cw = 3.75
    cgap = 0.24
    for i, (ctitle, ccol, cbullets) in enumerate(card_insights):
        cx = 0.8 + i * (cw + cgap)
        add_card(slide, cx, 5.42, cw, 1.55, C_CARD_BG, ccol)
        
        tb = slide.shapes.add_textbox(Inches(cx + 0.12), Inches(5.48), Inches(cw - 0.24), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(ctitle + "\n")
        r1.font.size = Pt(8.8)
        r1.font.bold = True
        r1.font.color.rgb = ccol
        
        for bullet in cbullets:
            p2 = tf.add_paragraph()
            p2.space_after = Pt(2)
            r2 = p2.add_run(bullet)
            r2.font.size = Pt(7.6)
            r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Chỉ vào biểu đồ cột Bar Chart đối chuẩn 7 doanh nghiệp qua 6 nhóm chủ đề SDG, làm bật quy luật thị trường và phân hóa đặc thù ngành.",
        "script": "Kính thưa Hội đồng, biểu đồ cột trên slide trực quan hóa mức độ gắn kết của 7 doanh nghiệp với 6 nhóm chủ đề SDG. Kết quả định lượng chỉ ra 3 phát hiện lớn: Thứ nhất, quy luật 'Economic cao nhất - Equity thấp nhất' thể hiện trực quan qua việc thanh màu xanh dương luôn cao nhất và thanh màu tím luôn thấp nhất ở toàn bộ các doanh nghiệp; thứ hai là sự phân hóa rõ nét giữa nhóm xuất khẩu tư nhân (chú trọng Tài nguyên & Môi trường để vượt rào cản CBAM) và nhóm tài chính/nhà nước (chú trọng Xã hội & Thể chế); thứ ba là từng doanh nghiệp đều có điểm bứt phá tương ứng với mô hình kinh doanh, như VCS dẫn đầu nhờ đá thạch anh, PNJ dẫn đầu Bình đẳng giới và Bảo Việt dẫn đầu Xã hội.",
        "highlights": "Toàn mẫu: Economic (47,69đ) > Social (46,39đ) > Resources (45,63đ) > Environs (44,98đ) > Life (44,26đ) > Equity (42,08đ); Thanh màu trực quan phân hóa theo ngành.",
        "qa": "Tại sao lại dùng biểu đồ cột thay cho bảng số liệu: Biểu đồ cột trực quan hóa tức thì sự chênh lệch giữa các nhóm chủ đề, giúp hội đồng thấy ngay cột Economic luôn vọt lên cao nhất và cột Equity luôn tụt sâu nhất ở 100% doanh nghiệp."
    })


def build_slide_09_result4_trends(prs):
    """Slide 9: Kết quả 4 - Xu hướng Chuỗi Thời gian (ẢNH LANDSCAPE GRID 2x2 + 3 THẺ)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 4: XU HƯỚNG DỊCH CHUYỂN CHUỖI THỜI GIAN (2020–2025)", "KẾT QUẢ THỰC NGHIỆM", 9, 17)
    
    # Dùng ảnh lưới landscape 2x2 slide_trends_grid.png (AR = 2.01) -> Width 11.733, Height 3.8
    tr_img = FIGURES_DIR / "slide_trends_grid.png"
    if tr_img.exists():
        slide.shapes.add_picture(str(tr_img), Inches(0.8), Inches(1.30), Inches(11.733), Inches(3.80))
    else:
        tr_alt = FIGURES_DIR / "trends_6categories.png"
        slide.shapes.add_picture(str(tr_alt), Inches(0.8), Inches(1.30), Inches(5.5), Inches(5.4))
        
    # 3 Thẻ xu hướng bên dưới: Top 5.35, Height 1.45 (Rất thoáng)
    phases = [
        ("GIAI ĐOẠN 2020–2021: ỨNG PHÓ COVID", C_NAVY_PRIMARY, 
         "Điểm số phân hóa mạnh, nhóm Life & Social tăng đột biến nhằm đảm bảo an toàn lao động và lương thưởng mùa dịch."),
        ("MỐC 2022: TÁI CƠ CẤU THEO THÔNG TƯ 96", C_BLUE_ACCENT, 
         "Số lượng báo cáo tăng vọt; điểm số 6 nhóm bắt đầu hội tụ theo cấu trúc tiêu chuẩn báo cáo bền vững GRI."),
        ("GIAI ĐOẠN 2024–2025: BÙNG NỔ NET ZERO", C_GREEN_EMERALD, 
         "Cả 7 tập đoàn đều đạt điểm số cao nhất lịch sử; nhóm Môi trường (+6,06đ) và Tài nguyên (+5,42đ) tăng tốc vượt bậc.")
    ]
    
    pw = 3.75
    pgap = 0.24
    for i, (ptitle, pcol, pdesc) in enumerate(phases):
        px = 0.8 + i * (pw + pgap)
        add_card(slide, px, 5.35, pw, 1.45, C_CARD_BG, pcol)
        
        tb = slide.shapes.add_textbox(Inches(px + 0.15), Inches(5.42), Inches(pw - 0.3), Inches(1.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(ptitle + "\n")
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = pcol
        
        r2 = p.add_run(pdesc)
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Phân tích xu hướng chuyển dịch 6 nhóm SDG qua 3 giai đoạn: Covid (2020-2021), Thể chế hóa (2022), Bùng nổ Net Zero (2024-2025).",
        "script": "Đồ thị đường chuỗi thời gian cho thấy bước chuyển dịch rõ nét qua 3 giai đoạn: 2020-2021 tập trung ứng phó đại dịch; 2022 tái định hình theo Thông tư 96; và 2024-2025 bùng nổ mạnh mẽ với các cam kết Net Zero, đưa điểm số toàn bộ 7 doanh nghiệp lên mức cao kỷ lục.",
        "highlights": "3 giai đoạn chuyển dịch; năm 2025 điểm số cao nhất lịch sử toàn mẫu.",
        "qa": "Xu hướng này phản ánh điều gì: Phản ánh nhận thức doanh nghiệp chuyển từ đối phó sang tích hợp ESG vào chiến lược cốt lõi."
    })


def build_slide_10_result5_sentiment(prs):
    """Slide 10: Kết quả 5 - Sắc thái Cảm xúc PhoBERT (ẢNH SUMMARY PANORAMA + 3 THẺ)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 5: SẮC THÁI CẢM XÚC PHOBERT & ĐỐI CHUẨN PAPER GỐC", "KẾT QUẢ THỰC NGHIỆM", 10, 17)
    
    # Dùng ảnh tổng hợp cảm xúc panorama slide_sentiment_summary.png (AR = 2.36) -> Width 11.733, Height 3.75
    senti_img = FIGURES_DIR / "slide_sentiment_summary.png"
    if senti_img.exists():
        slide.shapes.add_picture(str(senti_img), Inches(0.8), Inches(1.30), Inches(11.733), Inches(3.75))
    else:
        s_alt = FIGURES_DIR / "sentiment_hist.png"
        slide.shapes.add_picture(str(s_alt), Inches(0.8), Inches(1.30), Inches(6.0), Inches(3.5))
        
    # 3 Thẻ đối chuẩn cảm xúc bên dưới: Top 5.30, Height 1.50
    s_points = [
        ("PAPER GỐC (DISTILBERT 2 LỚP)", C_NAVY_PRIMARY,
         "Tích cực ~78%, Tiêu cực ~15%. Bỏ qua lớp Trung tính, ép câu số liệu kiểm toán kỹ thuật vào nhãn cảm xúc sai lệch."),
        ("VIỆT NAM (PHOBERT 3 LỚP)", C_GREEN_EMERALD,
         "Tích cực: 53,87% | Trung tính: 32,87% | Tiêu cực: 13,26%. Đột phá bảo lưu 1/3 câu số liệu kỹ thuật khách quan."),
        ("KẾT LUẬN VĂN PHONG DOANH NGHIỆP", C_PURPLE_ACCENT,
         "Báo cáo Việt Nam dành gần 1/3 dung lượng cho các số liệu đo lường kỹ thuật GRI chứ không chỉ đơn thuần quảng cáo PR.")
    ]
    
    sw = 3.75
    sgap = 0.24
    for i, (stitle, scol, sdesc) in enumerate(s_points):
        sx = 0.8 + i * (sw + sgap)
        add_card(slide, sx, 5.30, sw, 1.50, C_CARD_BG, scol)
        
        tb = slide.shapes.add_textbox(Inches(sx + 0.15), Inches(5.38), Inches(sw - 0.3), Inches(1.35))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(stitle + "\n")
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = scol
        
        r2 = p.add_run(sdesc)
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Làm nổi bật đóng góp của mô hình PhoBERT 3 lớp, đặc biệt là 32,87% câu Trung tính.",
        "script": "Khác với Kang & Kim dùng DistilBERT 2 lớp ép các câu số liệu vào nhãn tích cực hoặc tiêu cực, đề tài sử dụng PhoBERT 3 lớp. Kết quả cho thấy lớp Trung tính chiếm tới 32,87%, khẳng định báo cáo doanh nghiệp Việt Nam dành gần 1/3 dung lượng cho các số liệu kỹ thuật đo đạc khách quan.",
        "highlights": "PhoBERT 3 lớp: Pos 53,87%, Neutral 32,87%, Neg 13,26%.",
        "qa": "Lớp trung tính có ý nghĩa gì: Tránh sai lệch trong việc gán nhãn cảm xúc cho các câu số liệu kiểm toán kỹ thuật."
    })


def build_slide_11_result6_sentiment_ratio(prs):
    """Slide 11: Kết quả 6 - Tỷ số Cảm xúc Pos/Neg Ratio (CHUẨN TỶ LỆ 1.84)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 6: TỶ SỐ CẢM XÚC POS/NEG RATIO & ĐỐI CHUẨN TỶ LỆ", "KẾT QUẢ THỰC NGHIỆM", 11, 17)
    
    # Biểu đồ tỷ số (AR = 1.84) -> Width 6.8 in, Height 3.7 in
    img_r = FIGURES_DIR / "sentiment_ratio.png"
    if img_r.exists():
        slide.shapes.add_picture(str(img_r), Inches(0.8), Inches(1.35), Inches(6.8), Inches(3.70))
        
    # Thẻ takeaway dưới biểu đồ (Top 5.25, Height 1.55)
    card_b = add_card(slide, 0.8, 5.25, 6.8, 1.55, RGBColor(238, 244, 252), C_BLUE_ACCENT)
    tb_b = slide.shapes.add_textbox(Inches(0.95), Inches(5.32), Inches(6.5), Inches(1.4))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p_b0 = tf_b.paragraphs[0]
    r_b0 = p_b0.add_run("KẾT LUẬN THIÊN LỆCH LẠC QUAN CẤU TRÚC:\n")
    r_b0.font.bold = True
    r_b0.font.size = Pt(9.5)
    r_b0.font.color.rgb = C_NAVY_PRIMARY
    p_b1 = tf_b.add_paragraph()
    p_b1.text = "Tỷ số Pos/Neg trung bình đạt 4,06 lần khẳng định doanh nghiệp Việt Nam luôn có xu hướng dùng ngôn từ tích cực gấp hơn 4 lần so với tiêu cực nhằm làm đẹp hình ảnh quản trị (Hiệu ứng Pollyanna)."
    p_b1.font.size = Pt(8.5)
    p_b1.font.color.rgb = C_TEXT_DARK

    # Cột phải: 2 Thẻ đối chuẩn lớn (Left 7.95 in, Rộng 4.6 in, Cao 5.45 in tổng cộng)
    rw = 4.6
    rx = 7.95
    r_insights = [
        ("ĐỐI CHUẨN QUỐC TẾ: KANG & KIM (2022)", C_NAVY_PRIMARY, [
            "• Kang & Kim (2022): Pos/Neg toàn cầu đạt ≈ 5,20 lần.",
            "• Việt Nam: Pos/Neg trung bình đạt 4,06 lần (dao động 3,0 – 6,7x).",
            "• Cả 2 đều xác nhận Lý thuyết Quản trị Ấn tượng & Hiệu ứng Pollyanna."
        ]),
        ("CASE STUDY ĐẶC BIỆT: HIỆN TƯỢNG PNJ 2022", C_RED_ACCENT, [
            "• Tỷ số Pos/Neg chạm đáy 1,21x (Năm 2021: 4,48x -> 2023: 4,23x).",
            "• Căn nguyên Kỹ thuật NLP (Yếu tố quyết định):",
            "  Bản PDF scan ảnh khiến OCR bị vỡ từ, đứt gãy cú pháp câu;",
            "  PhoBERT nhạy cảm với nhiễu nên gán nhầm 39,0% câu thành Tiêu cực.",
            "• Căn nguyên Thực tế Doanh nghiệp:",
            "  PNJ đối mặt khó khăn hậu Covid-19 và tăng cường công bố rủi ro bán lẻ."
        ])
    ]
    for i, (rtitle, rcol, rbullets) in enumerate(r_insights):
        ry = 1.35 + i * 2.8
        add_card(slide, rx, ry, rw, 2.65, C_CARD_BG, rcol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.15), Inches(ry + 0.10), Inches(rw - 0.3), Inches(2.45))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(rtitle + "\n")
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = rcol
        
        for bullet in rbullets:
            p2 = tf.add_paragraph()
            p2.space_after = Pt(2.5)
            r2 = p2.add_run(bullet)
            if "Căn nguyên" in bullet or "Tỷ số Pos/Neg" in bullet:
                r2.font.bold = True
            r2.font.size = Pt(8.2)
            r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Phân tích Tỷ số Pos/Neg Ratio, đối chuẩn 4,06x của VN với 5,20x của Kang & Kim, bóc tách Case Study PNJ 2022 sụt giảm xuống 1,21x.",
        "script": "Đồ thị bên trái biểu diễn Tỷ số Cảm xúc Pos/Neg theo từng năm. Nghiên cứu của Kang & Kim cho tỷ số toàn cầu khoảng 5,2 lần, còn tại Việt Nam tỷ số đạt 4,06 lần. Điểm nổi bật nhất trên đồ thị là trường hợp dị biệt của PNJ năm 2022 khi tỷ số chạm đáy 1,21 lần. Nhóm nghiên cứu đã truy vết sâu và phát hiện 2 nguyên nhân: Thứ nhất về kỹ thuật, báo cáo PNJ 2022 là bản scan ảnh, OCR bị vỡ từ và đứt gãy cú pháp, khiến mô hình PhoBERT nhận diện nhầm các câu nhiễu thành Tiêu cực (chiếm tới 39%). Thứ hai về thực tế, đây là năm PNJ tái cấu trúc và đối mặt thách thức lớn hậu Covid. Đến năm 2023 khi dùng PDF số chuẩn, tỷ số lập tức hồi phục về 4,23 lần. Đây là phát hiện phương pháp luận đắt giá về ảnh hưởng của chất lượng số hóa tài liệu đến AI.",
        "highlights": "Pos/Neg trung bình = 4,06 lần; PNJ 2022 rơi xuống 1,21 lần do lỗi nhiễu OCR scan ảnh kết hợp thách thức hậu Covid; 2023 bật lại 4,23 lần.",
        "qa": "Tại sao PNJ 2022 lại thấp tịt 1,21: Do báo cáo là bản scan ảnh khiến OCR bị nhiễu và đứt câu, PhoBERT nhầm câu vỡ cấu trúc là Tiêu cực (39%), cộng hưởng cùng các khó khăn đóng cửa mạng lưới sau dịch."
    })


def build_slide_12_discussion_talk_heavy(prs):
    """Slide 12: Thảo luận - 'Nói nhiều về gì' (BẢNG TỔNG HỢP NATIVE TABLE 8 DÒNG + 3 THẺ PHÂN TÍCH SÂU)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "THẢO LUẬN: BẢNG TỔNG HỢP & PHÂN TÍCH 'NÓI NHIỀU VỀ GÌ' TẠI VIỆT NAM", "THẢO LUẬN CHUYÊN SÂU", 12, 17)
    
    # Cột trái: Bảng Native Table tổng hợp 7 Cty + Toàn mẫu (Width 7.0 in, Height 5.45 in, Left 0.8 in, Top 1.35 in)
    rows = 9
    cols = 3
    table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.35), Inches(7.0), Inches(5.45))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(1.7)   # DOANH NGHIỆP (NGÀNH)
    tbl.columns[1].width = Inches(2.7)   # TOP 3 SDG CAO NHẤT (ĐIỂM TB)
    tbl.columns[2].width = Inches(2.6)   # ĐẶC THÙ & ĐỘNG LỰC CHIẾN LƯỢC
    
    headers = ["DOANH NGHIỆP (NGÀNH)", "TOP 3 SDG CAO NHẤT (ĐIỂM TB)", "ĐỘNG LỰC CHIẾN LƯỢC ESG"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.bold = True
        p.font.size = Pt(8.8)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER
        
    table_data = [
        ("Vinamilk (Sữa & Chăn nuôi)", "SDG 17 (47.5) | SDG 12 (46.1) | SDG 07 (46.1)", "Trang trại Green Farm, Net Zero PAS 2060, kinh tế tuần hoàn."),
        ("Vicostone (VLXD Thạch anh)", "SDG 17 (53.1) | SDG 09 (51.7) | SDG 12 (51.4)", "Xuất khẩu 100% thị trường Âu - Mỹ, chứng chỉ xanh khắt khe."),
        ("PNJ (Bán lẻ & Chế tác vàng)", "SDG 17 (51.5) | SDG 07 (50.3) | SDG 09 (49.6)", "Tiết kiệm điện 400+ cửa hàng, chuỗi chế tác trang sức xanh."),
        ("Bảo Việt (Bảo hiểm - Tài chính)", "SDG 17 (52.0) | SDG 09 (49.7) | SDG 12 (49.5)", "Tiên phong Báo cáo Tích hợp (IIRC/GRI), quản trị minh bạch."),
        ("The PAN Group (Nông nghiệp)", "SDG 17 (49.3) | SDG 09 (47.8) | SDG 12 (47.5)", "Chuỗi Farm-Food-Family, an ninh lương thực (SDG 2: 45.3)."),
        ("Petrolimex (Xăng dầu - Năng lượng)", "SDG 17 (47.8) | SDG 09 (46.5) | SDG 12 (46.0)", "Hạ tầng xăng dầu Euro 5, lộ trình chuyển dịch năng lượng xanh."),
        ("SSI (Dịch vụ Tài chính)", "SDG 17 (50.3) | SDG 09 (48.1) | SDG 07 (47.1)", "Thu xếp vốn trái phiếu xanh, quản trị theo thẻ điểm ASEAN."),
        ("★ TOÀN MẪU VIỆT NAM (TB)", "SDG 17 (50.2) > SDG 09 (48.4) > SDG 12 (47.9)", "Quy luật 'Tam giác ưu tiên' chi phối toàn bộ thị trường.")
    ]
    
    for i, row in enumerate(table_data):
        is_highlight = (i == len(table_data) - 1)
        for j, val in enumerate(row):
            cell = tbl.cell(i + 1, j)
            cell.fill.solid()
            if is_highlight:
                cell.fill.fore_color.rgb = RGBColor(234, 243, 255)  # Soft highlight blue
            elif i % 2 == 0:
                cell.fill.fore_color.rgb = RGBColor(248, 250, 252)
            else:
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
                
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(7.6 if j == 2 else 8.0)
            if is_highlight:
                p.font.bold = True
                p.font.color.rgb = C_BLUE_ACCENT if j == 0 else C_NAVY_PRIMARY
            else:
                p.font.bold = (j == 0)
                p.font.color.rgb = C_NAVY_PRIMARY if j == 0 else C_TEXT_DARK
            if j == 1:
                p.alignment = PP_ALIGN.CENTER

    # Cột phải: 3 Thẻ Phân tích Sâu sắc & Sắc bén (Width 4.4 in, Left 8.1 in, Top 1.35 in, Height 1.68 in mỗi thẻ)
    rw = 4.4
    rx = 8.1
    card_insights = [
        ("1. QUY LUẬT 'TAM GIÁC ƯU TIÊN' (SDG 17 - 09 - 12)", C_NAVY_PRIMARY, [
            "• SDG 17 (Hợp tác đối tác) đứng Top 1 ở 100% doanh nghiệp (50,23 điểm).",
            "• Doanh nghiệp VN phụ thuộc chuỗi cung ứng toàn cầu; bắt buộc phải chứng minh hợp tác với nhà cung ứng và đối tác kiểm định quốc tế để duy trì đơn hàng xuất khẩu và vốn ngoại."
        ]),
        ("2. PHÂN HÓA RÕ NÉT THEO MÔ HÌNH KINH DOANH", C_GREEN_EMERALD, [
            "• Khối Sản xuất (VCS, VNM, PAN): Điểm SDG 12 & 07 vượt trội do áp lực trực tiếp từ rác thải bao bì, tiêu hao năng lượng và thuế carbon biên giới (CBAM).",
            "• Khối Tài chính (BVH, SSI): Ưu tiên SDG 8, 9, 16 để đáp ứng chuẩn xếp hạng tín nhiệm và tiêu chí đầu tư bền vững VNSI của HOSE."
        ]),
        ("3. ĐỘNG CƠ 'CHERRY-PICKING' (HÁI QUẢ TIỆN TAY)", C_GOLD_ACCENT, [
            "• Doanh nghiệp tập trung nói về Đầu tư công nghệ (SDG 9) và Đối tác (SDG 17) vì đây là các số liệu tài chính sẵn có từ Báo cáo Thường niên.",
            "• Dễ định lượng, an toàn về mặt PR và không để lộ rủi ro vi phạm môi trường."
        ])
    ]
    
    for i, (ctitle, ccol, cbullets) in enumerate(card_insights):
        cy = 1.35 + i * 1.85
        add_card(slide, rx, cy, rw, 1.75, C_CARD_BG, ccol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.15), Inches(cy + 0.08), Inches(rw - 0.3), Inches(1.6))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(ctitle + "\n")
        r1.font.size = Pt(9.0)
        r1.font.bold = True
        r1.font.color.rgb = ccol
        
        for bullet in cbullets:
            p2 = tf.add_paragraph()
            p2.space_after = Pt(2)
            r2 = p2.add_run(bullet)
            r2.font.size = Pt(7.8)
            r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Phân tích bảng tổng hợp xếp hạng Top SDGs của 7 doanh nghiệp, làm rõ quy luật 'Tam giác ưu tiên' và bản chất Cherry-picking.",
        "script": "Kính thưa Hội đồng, bảng bên trái là kết quả định lượng chính xác Top 3 mục tiêu SDG được nói nhiều nhất của từng doanh nghiệp và toàn thị trường. Chúng ta thấy nổi lên 3 phát hiện cực kỳ sâu sắc: Thứ nhất là Quy luật Tam giác ưu tiên SDG 17, 09 và 12. Cả 7 doanh nghiệp đều có điểm SDG 17 cao nhất tuyệt đối trên 50 điểm vì nền kinh tế Việt Nam định hướng xuất khẩu, bắt buộc phải chứng minh quan hệ đối tác quốc tế. Thứ hai là sự phân hóa ngành: Khối sản xuất như Vicostone, Vinamilk tập trung vào SDG 12 và SDG 7 do áp lực trực tiếp từ rác thải và thuế carbon CBAM; trong khi khối tài chính như Bảo Việt, SSI tập trung vào SDG 8 và 16 phục vụ chỉ số VNSI. Thứ ba, điều này phản ánh bản chất Cherry-picking: Doanh nghiệp chọn nói về những gì sẵn có số liệu đẹp từ báo cáo tài chính để làm PR.",
        "highlights": "Toàn mẫu VN: SDG 17 (50.2) > SDG 09 (48.4) > SDG 12 (47.9); Phân hóa sản xuất vs tài chính; Bản chất Cherry-picking.",
        "qa": "Tại sao SDG 17 luôn đứng đầu: Vì SDG 17 bao hàm các quan hệ đối tác công-tư, hợp tác nhà cung ứng và tuân thủ chuẩn mực báo cáo quốc tế - đây là nội dung bắt buộc trong mọi bản công bố thông tin."
    })


def build_slide_13_discussion_rarely_talk(prs):
    """Slide 13: Thảo luận - 'Ít nói về gì' (4 THẺ CẢNH BÁO THOÁNG ĐÃNG 2x2)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "THẢO LUẬN: DOANH NGHIỆP 'ÍT NÓI VỀ GÌ' (VÙNG TRŨNG NÉ TRÁNH)?", "THẢO LUẬN CHUYÊN SÂU", 13, 17)
    
    avoid_data = [
        ("1. ĐA DẠNG SINH HỌC & BẢO TỒN (SDG 14, 15)", C_RED_ACCENT, [
            "• Điểm số luôn ở mức đáy toàn mẫu (< 38 điểm).",
            "• Chỉ nêu khẩu hiệu chung chung, thiếu số liệu kiểm kê tác động sinh thái đất liền và biển.",
            "• Chi phí đo lường đa dạng sinh học phức tạp và chưa có chế tài bắt buộc."
        ]),
        ("2. BÌNH ĐẲNG LƯƠNG & CHÊNH LỆCH THU NHẬP (SDG 5, 10)", C_RED_ACCENT, [
            "• Tuyệt đối né tránh công bố Tỷ số chênh lệch lương CEO với công nhân.",
            "• Không có thống kê khoảng cách thu nhập theo giới tính ở cùng cấp bậc chuyên môn.",
            "• Chỉ dừng lại ở tỷ lệ % nhân sự nữ chung chung."
        ]),
        ("3. PHÁT THẢI CHUỖI CUNG ỨNG SCOPE 3", C_RED_ACCENT, [
            "• Mới chỉ đo lường phát thải trực tiếp Scope 1 và điện lưới Scope 2.",
            "• Phát thải gián tiếp chuỗi cung ứng Scope 3 (chiếm 70–80% thực tế) gần như bị bỏ ngỏ.",
            "• Do chuỗi cung ứng phân tán và thiếu công cụ đo lường chuyên sâu."
        ]),
        ("4. SỰ CỐ TIÊU CỰC, TRANH CHẤP & XỬ PHẠT (SDG 16)", C_RED_ACCENT, [
            "• 'Gạn đục khơi trong' điển hình: Không ghi nhận tai nạn lao động hay khiếu nại khách hàng.",
            "• Các quyết định xử phạt vi phạm hành chính về môi trường hoặc thuế bị che giấu hoàn toàn.",
            "• Báo cáo trở thành tài liệu tiếp thị thay vì công cụ giải trình rủi ro."
        ])
    ]
    
    w = 5.70
    h = 2.55
    top1 = 1.35
    top2 = 4.20
    l1 = 0.8
    l2 = 6.83
    coords = [(l1, top1), (l2, top1), (l1, top2), (l2, top2)]
    
    for i, (atitle, acol, abullets) in enumerate(avoid_data):
        cx, cy = coords[i]
        add_card(slide, cx, cy, w, h, C_CARD_BG, acol)
        
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(cy), Inches(w), Inches(0.4))
        set_shape_flat(strip, acol)
        p_st = strip.text_frame.paragraphs[0]
        p_st.alignment = PP_ALIGN.LEFT
        r_st = p_st.add_run("  ▲ " + atitle)
        r_st.font.name = FONT_MAIN
        r_st.font.size = Pt(10)
        r_st.font.bold = True
        r_st.font.color.rgb = C_WHITE
        
        tb = slide.shapes.add_textbox(Inches(cx + 0.15), Inches(cy + 0.45), Inches(w - 0.3), Inches(h - 0.5))
        tf = tb.text_frame
        tf.word_wrap = True
        for b_idx, bullet in enumerate(abullets):
            p = tf.add_paragraph() if b_idx > 0 else tf.paragraphs[0]
            p.space_after = Pt(2)
            r = p.add_run(bullet)
            r.font.name = FONT_MAIN
            r.font.size = Pt(8.8)
            r.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Chỉ ra 4 'vùng trũng' thông tin bị né tránh, khẳng định giá trị phản biện học thuật sắc bén của đề tài.",
        "script": "Mặt đối lập của bức tranh là 4 vùng trũng thông tin bị né tránh: 1. Đa dạng sinh học hoàn toàn thiếu số liệu; 2. Bình đẳng lương và chênh lệch thu nhập lãnh đạo bị giấu kín; 3. Phát thải Scope 3 chuỗi cung ứng bị bỏ ngỏ; và 4. Các sự cố xử phạt môi trường không được công bố. Đây là đóng góp phản biện thực tiễn quan trọng nhất của đề tài.",
        "highlights": "4 vùng né tránh: Đa dạng sinh học, Chênh lệch lương CEO, Scope 3 chuỗi cung ứng, Sự cố xử phạt.",
        "qa": "Tại sao phát hiện né tránh lại quan trọng: Giám sát AI giúp nhìn ra những gì bị ẩn giấu chứ không chỉ đọc những gì được trình bày."
    })


def build_slide_14_comparison(prs):
    """Slide 14: BẢNG SO SÁNH ĐỐI ĐẦU TOÀN DIỆN PAPER GỐC VS CODE MỚI (NATIVE TABLE RỘNG RÃI)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BẢNG SO SÁNH ĐỐI ĐẦU: PAPER GỐC (KANG & KIM 2022) VS CODE MỚI", "ĐỐI CHUẨN TOÀN DIỆN", 14, 17)
    
    rows = 9
    cols = 4
    left = Inches(0.8)
    top = Inches(1.30)
    width = Inches(11.733)
    height = Inches(5.60)
    
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    tbl = table_shape.table
    
    tbl.columns[0].width = Inches(2.2)
    tbl.columns[1].width = Inches(3.6)
    tbl.columns[2].width = Inches(3.9)
    tbl.columns[3].width = Inches(2.033)
    
    headers = ["TIÊU CHÍ KỸ THUẬT", "BÀI BÁO GỐC (KANG & KIM 2022)", "CODE MỚI ĐỀ TÀI (VIỆT NAM 2026)", "ĐÁNH GIÁ ĐỀ TÀI"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(h)
        r.font.name = FONT_MAIN
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = C_WHITE
        
    matrix_data = [
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
         "641 câu tiếng Anh chuẩn (169 chỉ tiêu LHQ)", 
         "1.032 câu song ngữ (391 câu VI + 641 câu EN)", 
         "Bản địa hóa chuẩn VN"),
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
         "Dừng lại ở biểu đồ thống kê tĩnh", 
         "Sẵn sàng làm Context Retriever sạch cho LLM kiểm toán", 
         "Nền tảng Agentic RAG")
    ]
    
    for i, row_data in enumerate(matrix_data, start=1):
        bg_col = C_ROW_ALT if i % 2 == 1 else C_CARD_BG
        for j, val in enumerate(row_data):
            cell = tbl.cell(i, j)
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
                r.font.color.rgb = C_BLUE_ACCENT
            elif j == 3:
                r.font.bold = True
                r.font.color.rgb = C_GREEN_EMERALD
            else:
                r.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Khẳng định sự vượt trội và giá trị nâng cấp của code mới so với bài báo gốc Kang & Kim (2022) trên 8 tiêu chí.",
        "script": "Bảng trên slide là sự đối chuẩn trực diện giữa công trình gốc Kang & Kim (2022) và nghiên cứu của chúng em trên 8 khía cạnh kỹ thuật: Đa ngữ tiếng Việt, OCR cứu hộ 100% scan, tăng gấp đôi số chiều vector lên 768 chiều, bộ ngữ liệu song ngữ 1.032 câu chuẩn, đột phá lớp Trung tính 32,9% của PhoBERT, phép nhân ma trận toàn cục dưới 3 giây, bóc tách đặc thù ngành sâu sắc và sẵn sàng cho Agentic RAG.",
        "highlights": "8 điểm nâng cấp trực tiếp so với paper gốc.",
        "qa": "Đóng góp lớn nhất: Không sao chép nguyên xi mà bản địa hóa thành công, giải quyết triệt để 4 rào cản kỹ thuật của paper gốc."
    })


def build_slide_15_implications(prs):
    """Slide 15: Hàm Ý Thực Tiễn & Đề Xuất Chính Sách (3 THẺ RỘNG RÃI)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "HÀM Ý THỰC TIỄN & ĐỀ XUẤT CHO CÁC BÊN LIÊN QUAN", "HÀM Ý QUẢN TRỊ", 15, 17)
    
    actions = [
        ("CƠ QUAN QUẢN LÝ (UBCKNN)", C_NAVY_PRIMARY, [
            "• Chuẩn hóa số báo cáo:",
            "  Thúc đẩy định dạng số mở (XBRL) thay cho PDF scan lỗi font.",
            "• Cổng giám sát AI tự động:",
            "  Ứng dụng NLP quét toàn thị trường, phát hiện dấu hiệu tô hồng báo cáo.",
            "• Mở rộng quy định Scope 3:",
            "  Bổ sung hướng dẫn kiểm kê phát thải gián tiếp chuỗi cung ứng."
        ]),
        ("DOANH NGHIỆP NIÊM YẾT", C_BLUE_ACCENT, [
            "• Chuyển đổi tư duy công bố:",
            "  Từ 'tiếp thị thành tích' sang 'giải trình trách nhiệm và quản trị rủi ro'.",
            "• Minh bạch hóa thách thức:",
            "  Chủ động công bố sự cố và lộ trình khắc phục để nâng cao tín nhiệm.",
            "• Tích hợp ESG thực chất:",
            "  Chuyển từ CSR từ thiện bề nổi sang chiến lược giảm phát thải cốt lõi."
        ]),
        ("NHÀ ĐẦU TƯ & KIỂM TOÁN", C_GREEN_EMERALD, [
            "• Thẩm định nhanh bằng AI:",
            "  Sàng lọc hàng trăm báo cáo trong vài phút, tiết kiệm 80% thời gian rà soát.",
            "• Cảnh báo Tẩy xanh sớm:",
            "  Theo dõi Tỷ số Pos/Neg Ratio để phát hiện các tuyên bố suông thiếu số liệu.",
            "• Định hướng dòng vốn xanh:",
            "  Ưu tiên giải ngân tín dụng xanh cho doanh nghiệp có cam kết thực chất."
        ])
    ]
    
    w = 3.75
    gap = 0.24
    top = 1.35
    h = 5.5
    
    for i, (atitle, acol, abullets) in enumerate(actions):
        ax = 0.8 + i * (w + gap)
        add_card(slide, ax, top, w, h, C_CARD_BG, acol)
        
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(ax), Inches(top), Inches(w), Inches(0.45))
        set_shape_flat(strip, acol)
        p_st = strip.text_frame.paragraphs[0]
        p_st.alignment = PP_ALIGN.CENTER
        r_st = p_st.add_run(atitle)
        r_st.font.name = FONT_MAIN
        r_st.font.size = Pt(10)
        r_st.font.bold = True
        r_st.font.color.rgb = C_WHITE
        
        tb = slide.shapes.add_textbox(Inches(ax + 0.15), Inches(top + 0.55), Inches(w - 0.3), Inches(h - 0.65))
        tf = tb.text_frame
        tf.word_wrap = True
        for b_idx, bullet in enumerate(abullets):
            p = tf.add_paragraph() if b_idx > 0 else tf.paragraphs[0]
            p.space_after = Pt(6)
            r = p.add_run(bullet)
            r.font.name = FONT_MAIN
            if bullet.startswith("•"):
                r.font.bold = True
                r.font.size = Pt(9.5)
                r.font.color.rgb = acol
            else:
                r.font.size = Pt(9.0)
                r.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Đưa ra khuyến nghị thiết thực cho 3 đối tượng: Nhà nước (UBCKNN), Doanh nghiệp và Nhà đầu tư.",
        "script": "Nghiên cứu đem lại 3 hàm ý lớn: UBCKNN cần chuẩn hóa định dạng XBRL và lập cổng AI giám sát; Doanh nghiệp cần chuyển từ tiếp thị sang giải trình rủi ro; Nhà đầu tư có thể dùng công cụ NLP này để sàng lọc nhanh hồ sơ vay vốn xanh và phòng ngừa tẩy xanh.",
        "highlights": "Khuyến nghị 3 bên: XBRL cho Nhà nước, Giải trình cho Doanh nghiệp, Sàng lọc nhanh cho Nhà đầu tư.",
        "qa": "Tại sao XBRL quan trọng: Giúp máy đọc tự động dữ liệu bảng biểu tài chính và phi tài chính 100% không bị lỗi font."
    })


def build_slide_16_limitations_future(prs):
    """Slide 16: Hạn Chế & Tương Lai (ẢNH SƠ ĐỒ MULTI-AGENT BANNER TRÊN + 3 THẺ)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "HẠN CHẾ ĐỀ TÀI & HƯỚNG PHÁT TRIỂN AGENTIC ESG AUDITOR", "HƯỚNG PHÁT TRIỂN", 16, 17)
    
    # Sơ đồ Multi-Agent ESG Auditor (AR = 2.85) -> Width 11.733, Height 3.95
    agent_img = FIGURES_DIR / "rag_agentic_flow.png"
    if agent_img.exists():
        slide.shapes.add_picture(str(agent_img), Inches(0.8), Inches(1.30), Inches(11.733), Inches(3.95))
        
    # 3 Thẻ hạn chế & tương lai bên dưới: Top 5.45, Height 1.40 (Rất thoáng)
    agent_cards = [
        ("HẠN CHẾ HIỆN TẠI", C_RED_ACCENT,
         "Mẫu 7 tập đoàn; mô hình đo lường tương đồng ngữ nghĩa chứ chưa tự động Fact-checking chéo số liệu định lượng trong Table."),
        ("ĐỘT PHÁ MULTI-AGENT", C_PURPLE_ACCENT,
         "Kế thừa pipeline Retrieval hiện tại làm nền móng sạch để tích hợp Agent bóc tách bảng số và Agent kiểm tra chéo cam kết."),
        ("TẦM NHÌN SẢN PHẨM SAAS", C_GREEN_EMERALD,
         "Tích hợp LLM (GPT-4o/Gemini) tự động xuất Báo cáo Thẩm định ESG Độc lập trong 30 giây phục vụ ngân hàng và quỹ đầu tư.")
    ]
    
    aw = 3.75
    agap = 0.24
    for i, (atitle, acol, adesc) in enumerate(agent_cards):
        ax = 0.8 + i * (aw + agap)
        add_card(slide, ax, 5.45, aw, 1.40, C_CARD_BG, acol)
        
        tb = slide.shapes.add_textbox(Inches(ax + 0.15), Inches(5.52), Inches(aw - 0.3), Inches(1.25))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(atitle + "\n")
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = acol
        
        r2 = p.add_run(adesc)
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Thẳng thắn nhìn nhận hạn chế và vạch ra lộ trình tích hợp Multi-Agent AI trong tương lai.",
        "script": "Chúng em nhìn nhận rõ 3 hạn chế: Mẫu 7 tập đoàn, chưa kiểm chứng chéo số liệu bảng biểu, và chunking câu đơn. Tuy nhiên sơ đồ bên phải chỉ ra hướng phát triển tự nhiên: dùng pipeline hiện tại làm nền móng sạch để cắm thêm các Agent kiểm tra số liệu bảng biểu và đưa vào LLM sinh báo cáo kiểm toán độc lập.",
        "highlights": "Thẳng thắn nhìn nhận hạn chế; Mở ra tầm nhìn Agentic ESG Auditor.",
        "qa": "Tại sao không dùng LLM ngay từ đầu: Vì LLM nhồi cả cuốn PDF rất đắt và dễ ảo giác, dùng SBERT làm Retrieval sạch trước mới đảm bảo độ chính xác."
    })


def build_slide_17_conclusion(prs):
    """Slide 17: Kết Luận & Phiên Hỏi Đáp (Q&A) (THOÁNG ĐÃNG, CÂN ĐỐI)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, C_NAVY_DARK)
    
    accent_top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    set_shape_flat(accent_top, C_GOLD_ACCENT)

    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.733), Inches(0.75))
    p_t = tb_title.text_frame.paragraphs[0]
    r_t = p_t.add_run("TỔNG KẾT 4 THÔNG ĐIỆP CỐT LÕI CỦA ĐỀ TÀI")
    r_t.font.name = FONT_HEADING
    r_t.font.size = Pt(20)
    r_t.font.bold = True
    r_t.font.color.rgb = C_WHITE

    conclusions = [
        ("1. CHUYỂN GIAO THÀNH CÔNG",
         "Quy trình NLP của Kang & Kim (2022) thích ứng xuất sắc với tiếng Việt, đạt phân phối chuẩn tương đương 100% nghiên cứu quốc tế."),
        ("2. PHẢN ÁNH ĐẶC THÙ NGÀNH",
         "Điểm số SDG phản ánh sát sao mô hình kinh doanh: Sản xuất mạnh về Tài nguyên/Môi trường, Tài chính dẫn đầu về Kinh tế, Bán lẻ nổi bật về DE&I."),
        ("3. BỨC TRANH TƯƠNG PHẢN",
         "Chỉ ra sự phân hóa sâu sắc: Doanh nghiệp nói nhiều về tăng trưởng và CSR từ thiện bề nổi, nhưng né tránh đa dạng sinh học và chênh lệch lương."),
        ("4. THIÊN LỆCH LẠC QUAN CẤU TRÚC",
         "Tỷ số Pos/Neg đạt 4,06 lần khẳng định sự tồn tại của chiến lược Quản trị Ấn tượng, khẳng định tính cấp thiết của công cụ AI hỗ trợ kiểm toán.")
    ]

    w_c = 5.7
    h_c = 1.85
    top1 = 1.35
    top2 = 3.45
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
        r0.font.size = Pt(10.5)
        
        p1 = tf.add_paragraph()
        p1.text = c_body
        p1.font.size = Pt(9.2)
        p1.font.color.rgb = RGBColor(225, 235, 250)

    qa_card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(5.55), Inches(11.733), Inches(1.35))
    set_shape_flat(qa_card, RGBColor(15, 35, 65), C_GOLD_ACCENT, 1.5)
    
    tb_qa = slide.shapes.add_textbox(Inches(1.0), Inches(5.62), Inches(11.333), Inches(1.2))
    tf_qa = tb_qa.text_frame
    tf_qa.word_wrap = True
    
    p_qa1 = tf_qa.paragraphs[0]
    p_qa1.alignment = PP_ALIGN.CENTER
    r_qa1 = p_qa1.add_run("TRÂN TRỌNG CẢM ƠN QUÝ THẦY CÔ TRONG HỘI ĐỒNG KHOA HỌC!\n")
    r_qa1.font.name = FONT_HEADING
    r_qa1.font.size = Pt(14.5)
    r_qa1.font.bold = True
    r_qa1.font.color.rgb = C_GOLD_ACCENT
    
    p_qa2 = tf_qa.add_paragraph()
    p_qa2.alignment = PP_ALIGN.CENTER
    r_qa2 = p_qa2.add_run("Nhóm nghiên cứu rất mong nhận được các câu hỏi và ý kiến đóng góp quý báu từ Quý Thầy Cô.\n(Tác giả: Lê Đan Sơn, Dương Thị Hoàn — 2026)")
    r_qa2.font.name = FONT_MAIN
    r_qa2.font.size = Pt(10.5)
    r_qa2.font.italic = True
    r_qa2.font.color.rgb = C_WHITE

    set_presenter_notes(slide, {
        "goal": "Tóm kết đĩnh đạc 4 thông điệp cốt lõi, gửi lời cảm ơn trang trọng đến Hội đồng và tự tin mở phiên Hỏi đáp (Q&A).",
        "script": "Kính thưa Quý Thầy Cô trong Hội đồng, đề tài gửi gắm 4 thông điệp cốt lõi: 1. Chuyển giao thành công sang tiếng Việt; 2. Phản ánh đúng bản chất ngành; 3. Bóc trần sự đối lập giữa những gì nói nhiều và né tránh; 4. Xác nhận thiên lệch lạc quan và nhu cầu công cụ AI kiểm toán. Nhóm tác giả xin trân trọng cảm ơn Quý Thầy Cô!",
        "highlights": "Cúi đầu chào và mời các thầy cô đặt câu hỏi. Giữ phong thái tự tin, khiêm tốn.",
        "qa": "Sẵn sàng mở lại các slide biểu đồ tương ứng khi thầy cô yêu cầu giải trình."
    })


def main():
    print("=" * 80)
    print("BẮT ĐẦU CẬP NHẬT BỘ SLIDE CHÍNH (PHIÊN BẢN DE-CRAMPED THOÁNG ĐÃNG)...")
    print("=" * 80)
    
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    print("[1/17] Slide 1: Trang Tiêu đề & Thông tin Tác giả...")
    build_slide_01_title(prs)
    
    print("[2/17] Slide 2: Đặt vấn đề & Bối cảnh Thể chế Việt Nam (Thẻ KPI thoáng)...")
    build_slide_02_context(prs)
    
    print("[3/17] Slide 3: Bài báo gốc Kang & Kim (2022) vs Code mới (Biểu đồ 3 Panel chuẩn AR)...")
    build_slide_03_original_paper(prs)
    
    print("[4/17] Slide 4: Khung Phương pháp luận (Sơ đồ 16:9 5 giai đoạn toàn diện)...")
    build_slide_04_pipeline(prs)
    
    print("[5/17] Slide 5: Mẫu Dữ liệu Thực nghiệm (Bảng 7 Doanh nghiệp + 4 Thẻ KPI)...")
    build_slide_05_sample(prs)
    
    print("[6/17] Slide 6: Kết quả 1 - Phân phối Tương đồng (AR 2.28) & BẢNG ĐỐI CHUẨN THỐNG KÊ...")
    build_slide_06_result1_similarity(prs)
    
    print("[7/17] Slide 7: Kết quả 2 - Cấu trúc 6 Nhóm SDG (Heatmap AR 0.65 portrait + 3 Thẻ rộng)...")
    build_slide_07_result2_heatmap(prs)
    
    print("[8/17] Slide 8: Kết quả 3 - Biểu đồ Cột So sánh 7 Doanh nghiệp theo 6 Nhóm SDG (slide_company_6cat_bar.png)...")
    build_slide_08_company_6cat_bar(prs)
    
    print("[9/17] Slide 9: Kết quả 4 - Xu hướng Chuỗi Thời gian (slide_trends_grid.png AR 2.01 landscape)...")
    build_slide_09_result4_trends(prs)
    
    print("[10/17] Slide 10: Kết quả 5 - Sắc thái Cảm xúc (slide_sentiment_summary.png AR 2.36 panorama)...")
    build_slide_10_result5_sentiment(prs)
    
    print("[11/17] Slide 11: Kết quả 6 - Tỷ số Cảm xúc Pos/Neg Ratio & Case PNJ 2022...")
    build_slide_11_result6_sentiment_ratio(prs)
    
    print("[12/17] Slide 12: Thảo luận - 'Nói nhiều về gì' (Bảng tổng hợp Native Table + 3 Thẻ phân tích sâu)...")
    build_slide_12_discussion_talk_heavy(prs)
    
    print("[13/17] Slide 13: Thảo luận - 'Ít nói về gì' (4 Thẻ cảnh báo né tránh 2x2)...")
    build_slide_13_discussion_rarely_talk(prs)
    
    print("[14/17] Slide 14: BẢNG SO SÁNH ĐỐI ĐẦU TOÀN DIỆN PAPER GỐC VS CODE MỚI (Native Table rộng)...")
    build_slide_14_comparison(prs)
    
    print("[15/17] Slide 15: Hàm ý Thực tiễn & Đề xuất Chính sách (3 Thẻ hành động)...")
    build_slide_15_implications(prs)
    
    print("[16/17] Slide 16: Hạn chế Đề tài & Tương lai (rag_agentic_flow.png AR 2.85 banner)...")
    build_slide_16_limitations_future(prs)
    
    print("[17/17] Slide 17: Tổng kết 4 Thông điệp Cốt lõi & Phiên Hỏi đáp (Q&A)...")
    build_slide_17_conclusion(prs)
    
    prs.save(str(OUTPUT_PPTX))
    print(f"\n=> Đã lưu thành công bộ slide chính tại: {OUTPUT_PPTX}")
    
    PAPERS_DIR.mkdir(parents=True, exist_ok=True)
    dest_copy = PAPERS_DIR / "bao_cao_nghien_cuu_sdg_vietnam.pptx"
    shutil.copyfile(OUTPUT_PPTX, dest_copy)
    print(f"=> Đã sao chép vào: {dest_copy}")
    
    file_size_mb = OUTPUT_PPTX.stat().st_size / (1024 * 1024)
    print(f"=> Kích thước tệp: {file_size_mb:.2f} MB")
    print("=" * 80)
    print("HOÀN TẤT THÀNH CÔNG BỘ SLIDE DE-CRAMPED!")
    print("=" * 80)


if __name__ == "__main__":
    main()
