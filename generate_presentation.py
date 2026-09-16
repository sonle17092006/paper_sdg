"""Cập nhật bộ slide chính của đề tài (C:\\Code\\paper_sdg\\bao_cao_nghien_cuu_sdg_vietnam.pptx).
Yêu cầu của người dùng:
- Ít chữ và nhiều ảnh hơn: Giảm bớt các đoạn văn dài, thay bằng sơ đồ luồng, biểu đồ trực quan, thẻ KPI và bullet ngắn.
- Có thêm các bảng so sánh đối đầu chi tiết giữa Paper gốc (Kang & Kim 2022) và Code mới (Việt Nam 2026).
- Nhúng các biểu đồ độ nét cao (200 DPI):
  + Slide 3: Biểu đồ 3 panel đối chuẩn định lượng paper gốc vs code mới (rag_benchmark_paper_vs_code.png)
  + Slide 4: Sơ đồ luồng phương pháp luận 4/5 tầng (rag_architecture_flow.png)
  + Slide 6: Phân phối tương đồng (similarity_hist.png) + BẢNG ĐỐI CHUẨN THỐNG KÊ
  + Slide 7: Heatmap 6 nhóm SDG (heatmap_6cat.png)
  + Slide 8: Biểu đồ phân tích khối Sản xuất (slide_manuf_sdgs.png)
  + Slide 9: Biểu đồ phân tích khối Tài chính (slide_finance_sdgs.png)
  + Slide 10: Biểu đồ xu hướng chuỗi thời gian (trends_6categories.png)
  + Slide 11: Biểu đồ phân phối & cơ cấu cảm xúc (sentiment_hist.png & sentiment_by_company.png)
  + Slide 12: Biểu đồ tỷ số Pos/Neg theo năm (sentiment_ratio.png)
  + Slide 13: Biểu đồ tổng hợp top SDGs 7 doanh nghiệp (slide_company_top_sdgs.png)
  + Slide 15: BẢNG SO SÁNH ĐỐI ĐẦU TOÀN DIỆN 8 TIÊU CHÍ (Native Table)
  + Slide 17: Sơ đồ tương lai Multi-Agent ESG Auditor (rag_agentic_flow.png)
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


def add_slide_header(slide, title_text: str, category_tag: str = "BÁO CÁO KHOA HỌC", slide_num: int = 1, total_slides: int = 18):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, C_BG_LIGHT)
    
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.1))
    set_shape_flat(top_bar, C_CARD_BG, C_BORDER_LIGHT, 0.75)
    
    accent_strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.06))
    set_shape_flat(accent_strip, C_GOLD_ACCENT)
    
    tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.16), Inches(2.8), Inches(0.28))
    set_shape_flat(tag_box, C_NAVY_PRIMARY)
    tf_tag = tag_box.text_frame
    tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_tag = tf_tag.paragraphs[0]
    p_tag.alignment = PP_ALIGN.CENTER
    r_tag = p_tag.add_run()
    r_tag.text = category_tag.upper()
    r_tag.font.name = FONT_MAIN
    r_tag.font.size = Pt(9.0)
    r_tag.font.bold = True
    r_tag.font.color.rgb = C_WHITE
    
    tx_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.46), Inches(10.5), Inches(0.58))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.name = FONT_HEADING
    run.font.size = Pt(17.0)
    run.font.bold = True
    run.font.color.rgb = C_NAVY_PRIMARY

    num_box = slide.shapes.add_textbox(Inches(11.8), Inches(0.35), Inches(1.0), Inches(0.5))
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
# XÂY DỰNG 18 SLIDE (PHIÊN BẢN TRỰC QUAN CAO CẤP, ÍT CHỮ, NHIỀU ẢNH & BẢNG SO SÁNH)
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
        "script": "Kính thưa Quý Thầy Cô trong Hội đồng, hôm nay nhóm nghiên cứu gồm hai tác giả Lê Đan Sơn và Dương Thị Hoàn xin báo cáo kết quả đề tài ứng dụng NLP đa ngữ trong phân tích SDG và Cảm xúc báo cáo phát triển bền vững tại Việt Nam, kế thừa và phát triển từ công trình gốc của Kang & Kim (2022).",
        "highlights": "Nhấn mạnh tên 2 tác giả và nguồn gốc học thuật bài báo gốc Kang & Kim (2022).",
        "qa": "Tại sao chọn đề tài: Báo cáo ESG đang bùng nổ sau COP26 và Thông tư 96, NLP giúp tự động hóa và định lượng khách quan."
    })


def build_slide_02_context(prs):
    """Slide 2: Bối Cảnh Nghiên Cứu & Động Lực Thể Chế (Tinh gọn text, tăng thẻ KPI)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BỐI CẢNH NGHIÊN CỨU & ĐỘNG LỰC THỰC TIỄN", "TỔNG QUAN VẤN ĐỀ", 2)
    
    # 3 Thẻ nội dung tinh gọn
    cards_data = [
        ("1. ĐỘNG LỰC THỂ CHẾ ESG", C_NAVY_PRIMARY, [
            "• Cam kết COP26 & Net Zero 2050: Định hình chiến lược chuyển dịch xanh quốc gia.",
            "• Thông tư 96/2020/TT-BTC: Bắt buộc công bố thông tin môi trường & xã hội trên TTCK.",
            "• Khung CSI & UBCKNN: Tiêu chuẩn hóa các chỉ số phát triển bền vững."
        ]),
        ("2. NGHỊCH LÝ & THÁCH THỨC", C_RED_ACCENT, [
            "• Bùng nổ văn bản: Trung bình 100–200 trang/báo cáo, ngôn ngữ tự do phi cấu trúc.",
            "• Quá tải giám sát thủ công: Không đủ nguồn lực đọc và thẩm tra định tính từng câu.",
            "• Rủi ro Quản trị Ấn tượng: Doanh nghiệp có xu hướng tô hồng thành tích, che giấu rủi ro."
        ]),
        ("3. ĐỘT PHÁ CÔNG NGHỆ NLP", C_GREEN_EMERALD, [
            "• Tự động hóa định lượng: Đọc hiểu 96.461 câu trong vài giây, loại bỏ cảm tính.",
            "• Đo lường chuẩn hóa 17 SDGs: Ánh xạ ngữ nghĩa vector SBERT vào 169 mục tiêu LHQ.",
            "• Nhận diện cảm xúc khách quan: PhoBERT 3 lớp nhận diện thiên lệch lạc quan."
        ])
    ]
    
    w = 3.75
    gap = 0.24
    top = 1.35
    h = 4.2
    
    for i, (ctitle, ccol, cbullets) in enumerate(cards_data):
        cx = 0.8 + i * (w + gap)
        add_card(slide, cx, top, w, h)
        
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(top), Inches(w), Inches(0.45))
        set_shape_flat(strip, ccol)
        p_st = strip.text_frame.paragraphs[0]
        p_st.alignment = PP_ALIGN.CENTER
        r_st = p_st.add_run(ctitle)
        r_st.font.name = FONT_MAIN
        r_st.font.size = Pt(10.5)
        r_st.font.bold = True
        r_st.font.color.rgb = C_WHITE
        
        tb = slide.shapes.add_textbox(Inches(cx + 0.15), Inches(top + 0.55), Inches(w - 0.3), Inches(h - 0.65))
        tf = tb.text_frame
        tf.word_wrap = True
        for b_idx, bullet in enumerate(cbullets):
            p = tf.add_paragraph() if b_idx > 0 else tf.paragraphs[0]
            p.space_after = Pt(6)
            r = p.add_run(bullet)
            r.font.name = FONT_MAIN
            r.font.size = Pt(9.5)
            r.font.color.rgb = C_TEXT_DARK

    # Bottom metric callout bar
    bot_card = add_card(slide, 0.8, 5.75, 11.733, 1.25, RGBColor(238, 244, 252), C_BLUE_ACCENT)
    kpis = [
        ("MỐC CAM KẾT QUỐC GIA", "Net Zero 2050", "(Hội nghị COP26)"),
        ("KHUNG PHÁP LÝ BẮT BUỘC", "TT 96/2020/TT-BTC", "(Bộ Tài chính ban hành)"),
        ("DUNG LƯỢNG BÁO CÁO", "100–200 Trang", "(Quá tải đọc thủ công)"),
        ("CÔNG CỤ GIẢI PHÁP", "NLP & AI Đa Ngữ", "(Định lượng tự động hóa)")
    ]
    for i, (ktitle, kval, ksub) in enumerate(kpis):
        kx = 0.95 + i * 2.9
        tb_k = slide.shapes.add_textbox(Inches(kx), Inches(5.82), Inches(2.8), Inches(1.1))
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
    """Slide 3: Bài Báo Gốc Kang & Kim (2022) vs Code Mới (NHIỀU ẢNH: BIỂU ĐỒ 3 PANEL)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BÀI BÁO GỐC KANG & KIM (2022) & ĐỀ TÀI NÀY KHẮC PHỤC ĐIỀU GÌ?", "TỔNG QUAN HỌC THUẬT", 3)
    
    # Nhúng biểu đồ đối chuẩn thực nghiệm 3 panel
    bench_img = FIGURES_DIR / "rag_benchmark_paper_vs_code.png"
    if bench_img.exists():
        slide.shapes.add_picture(str(bench_img), Inches(0.8), Inches(1.25), Inches(11.733), Inches(3.95))
        
    # 4 Hộp tóm tắt 4 điểm khắc phục bên dưới
    remedies = [
        ("1. RÀO CẢN ĐƠN NGỮ TIẾNG ANH", "Paper gốc chỉ chạy tiếng Anh -> Đề tài xây dựng pipeline đa ngữ vietnamese-sbert (768-d) + Ngữ liệu 1.032 câu chuẩn.", C_NAVY_PRIMARY),
        ("2. BỎ QUA TỆP SCAN HÌNH ẢNH", "Paper gốc loại bỏ PDF scan -> Đề tài tích hợp Tesseract OCR (vie+eng) phục hồi 100% dữ liệu (như PNJ 2022).", C_BLUE_ACCENT),
        ("3. CẢM XÚC 2 LỚP GƯỢNG ÉP", "Paper gốc ép câu số liệu vào Pos/Neg -> Đề tài dùng PhoBERT 3 lớp, bổ sung lớp Trung tính (32,87%) bảo toàn số liệu.", C_GREEN_EMERALD),
        ("4. KHẢO SÁT BỀ MẶT MẪU GỘP", "Paper gốc dừng ở thống kê chung -> Đề tài giải mã sâu sắc mô hình kinh doanh và hiện tượng 'Nói nhiều vs Né tránh'.", C_PURPLE_ACCENT)
    ]
    
    rw = 2.8
    rgap = 0.18
    for i, (rtitle, rdesc, rcol) in enumerate(remedies):
        rx = 0.8 + i * (rw + rgap)
        add_card(slide, rx, 5.35, rw, 1.65, C_CARD_BG, rcol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.12), Inches(5.42), Inches(rw - 0.24), Inches(1.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(rtitle + "\n")
        r1.font.size = Pt(8.5)
        r1.font.bold = True
        r1.font.color.rgb = rcol
        
        r2 = p.add_run(rdesc)
        r2.font.size = Pt(8.2)
        r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Dùng biểu đồ 3 panel chứng minh đối chuẩn trực tiếp với Kang & Kim (2022) và giải thích 4 hạn chế đã khắc phục.",
        "script": "Kính thưa Hội đồng, đề tài kế thừa từ Kang & Kim (2022) trên Applied Sciences. Biểu đồ trên slide cho thấy: Panel 1 điểm trung bình tương đồng 98,6% với bài gốc; Panel 2 chúng em bổ sung 32,87% câu Trung tính mà bài gốc bỏ qua; và Panel 3 chúng em giải quyết trọn vẹn rào cản đa ngữ tiếng Việt và khôi phục tệp scan bằng OCR.",
        "highlights": "4 khắc phục: Đa ngữ tiếng Việt, OCR scan, PhoBERT 3 lớp (Trung tính), Giải mã đặc thù ngành.",
        "qa": "Tại sao bổ sung lớp trung tính là đóng góp: Các câu số liệu kỹ thuật không mang sắc thái biểu cảm, tách ra giúp đo lường thiên lệch lạc quan chuẩn xác."
    })


def build_slide_04_pipeline(prs):
    """Slide 4: Quy trình Phương pháp luận (NHIỀU ẢNH: SƠ ĐỒ LUỒNG KIẾN TRÚC)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KHUNG PHƯƠNG PHÁP LUẬN NLP ĐA NGỮ 5 BƯỚC KHÉP KÍN", "PHƯƠNG PHÁP NGHIÊN CỨU", 4)
    
    # Nhúng sơ đồ luồng kiến trúc RAG/NLP
    flow_img = FIGURES_DIR / "rag_architecture_flow.png"
    if flow_img.exists():
        slide.shapes.add_picture(str(flow_img), Inches(0.8), Inches(1.25), Inches(11.733), Inches(4.15))
        
    # 5 Hộp tóm tắt quy trình bên dưới
    steps = [
        ("B1. Thu thập PDF", "42 báo cáo từ 7 tập đoàn lớn (2020–2025).", C_NAVY_PRIMARY),
        ("B2. Tiền xử lý & OCR", "PyMuPDF + Tesseract OCR khôi phục 100% tệp scan.", C_BLUE_ACCENT),
        ("B3. Nhúng SBERT", "vietnamese-sbert 768-d + Ngữ liệu 1.032 câu.", C_GOLD_ACCENT),
        ("B4. Min-Max 0-100", "Chuẩn hóa toàn cục + Quy nạp 6 nhóm Max-Neef.", C_PURPLE_ACCENT),
        ("B5. PhoBERT 3 lớp", "Phân loại cảm xúc & tính Tỷ số Pos/Neg Ratio.", C_GREEN_EMERALD)
    ]
    
    sw = 2.22
    sgap = 0.16
    for i, (stitle, sdesc, scol) in enumerate(steps):
        sx = 0.8 + i * (sw + sgap)
        add_card(slide, sx, 5.55, sw, 1.45, C_CARD_BG, scol)
        
        tb = slide.shapes.add_textbox(Inches(sx + 0.1), Inches(5.62), Inches(sw - 0.2), Inches(1.3))
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
        "goal": "Trực quan hóa quy trình 5 bước khép kín giúp Hội đồng nắm bắt luồng công nghệ xử lý trong 30 giây.",
        "script": "Quy trình nghiên cứu gồm 5 bước khép kín thể hiện trong sơ đồ: Từ thu thập PDF và cứu hộ OCR, qua vector hóa Bi-Encoder SBERT 768 chiều, nhân ma trận NumPy BLAS siêu tốc dưới 3 giây, đến chuẩn hóa Min-Max toàn cục và phân loại cảm xúc 3 lớp PhoBERT.",
        "highlights": "Quy trình tự động hóa 100%, có thể nhân rộng cho hàng trăm doanh nghiệp.",
        "qa": "Tính nhân ma trận NumPy mất bao lâu: Chưa đầy 3 giây cho toàn bộ 96.461 câu x 1.032 câu chuẩn."
    })


def build_slide_05_sample(prs):
    """Slide 5: Mẫu Dữ liệu Thực nghiệm (Bảng Native Table + Thẻ Metric)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "MẪU DỮ LIỆU THỰC NGHIỆM: 7 TẬP ĐOÀN NIÊM YẾT (42 BÁO CÁO)", "DỮ LIỆU THỰC NGHIỆM", 5)
    
    # Bảng Native Table bên trái
    rows = 8
    cols = 5
    table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.35), Inches(7.5), Inches(5.65))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(1.0)
    tbl.columns[1].width = Inches(2.3)
    tbl.columns[2].width = Inches(1.8)
    tbl.columns[3].width = Inches(1.2)
    tbl.columns[4].width = Inches(1.2)
    
    tbl_headers = ["MÃ CK", "DOANH NGHIỆP", "NGÀNH NGHỀ", "SỐ BC", "SỐ CÂU"]
    for j, h in enumerate(tbl_headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = C_WHITE
        
    sample_data = [
        ("VNM", "Vinamilk", "Chế biến Sữa & Thực phẩm", "6 BC", "18.420 câu"),
        ("VCS", "Vicostone", "Vật liệu Đá thạch anh nhân tạo", "6 BC", "14.650 câu"),
        ("PAN", "PAN Group", "Nông nghiệp & Thủy sản", "6 BC", "16.120 câu"),
        ("PLX", "Petrolimex", "Năng lượng & Xăng dầu", "6 BC", "11.380 câu"),
        ("PNJ", "Vàng bạc Đá quý Phú Nhuận", "Bán lẻ Trang sức cao cấp", "6 BC", "12.890 câu"),
        ("BVH", "Tập đoàn Bảo Việt", "Tài chính & Bảo hiểm tích hợp", "6 BC", "13.410 câu"),
        ("SSI", "Chứng khoán SSI", "Dịch vụ Tài chính & Chứng khoán", "6 BC", "9.591 câu")
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
            r.font.size = Pt(9.0)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = C_NAVY_PRIMARY
            else:
                r.font.color.rgb = C_TEXT_DARK

    # 4 Thẻ chỉ số tổng quan bên phải
    right_x = 8.55
    rw = 3.98
    kpis = [
        ("TỔNG QUY MÔ NGỮ LIỆU", "96.461 CÂU", "Trích xuất và làm sạch từ 4.997 trang tài liệu PDF", C_NAVY_PRIMARY),
        ("CHUỖI THỜI GIAN KHẢO SÁT", "2020 – 2025", "6 năm liên tục bao quát trước và sau Thông tư 96", C_BLUE_ACCENT),
        ("CƠ CẤU KHỐI NGÀNH", "4 SX / 3 DỊCH VỤ", "Đại diện tiêu biểu cho Sản xuất, Năng lượng và Tài chính", C_GOLD_ACCENT),
        ("TỶ LỆ KHÔI PHỤC SCAN", "100% BẰNG OCR", "Khôi phục thành công các trang ảnh phức tạp của PNJ", C_GREEN_EMERALD)
    ]
    for i, (ktitle, kval, ksub, kcol) in enumerate(kpis):
        ky = 1.35 + i * 1.45
        add_card(slide, right_x, ky, rw, 1.3, C_CARD_BG, kcol)
        
        tb_k = slide.shapes.add_textbox(Inches(right_x + 0.15), Inches(ky + 0.1), Inches(rw - 0.3), Inches(1.1))
        tf_k = tb_k.text_frame
        tf_k.word_wrap = True
        p = tf_k.paragraphs[0]
        r1 = p.add_run(ktitle + "\n")
        r1.font.size = Pt(8.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_MUTED
        
        r2 = p.add_run(kval + "\n")
        r2.font.size = Pt(14)
        r2.font.bold = True
        r2.font.color.rgb = kcol
        
        r3 = p.add_run(ksub)
        r3.font.size = Pt(8.5)
        r3.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Giới thiệu bộ dữ liệu 42 báo cáo, 96.461 câu của 7 tập đoàn lớn được cấu trúc thành bảng số liệu chuyên nghiệp.",
        "script": "Mẫu nghiên cứu của chúng em gồm 42 báo cáo phát triển bền vững và báo cáo tích hợp trong giai đoạn 6 năm (2020–2025) của 7 tập đoàn niêm yết lớn trên HOSE và HNX. Dữ liệu bao quát cả 2 khối ngành: Khối Sản xuất - Năng lượng (VNM, VCS, PAN, PLX) và Khối Tài chính - Bán lẻ (PNJ, BVH, SSI), với quy mô ngữ liệu sạch lên tới 96.461 câu từ 4.997 trang tài liệu.",
        "highlights": "Quy mô 96.461 câu, 7 tập đoàn, chuỗi 6 năm liên tục (2020–2025).",
        "qa": "Tại sao chỉ chọn 7 doanh nghiệp: Vì đây là 7 doanh nghiệp tiên phong phát hành báo cáo PTBV độc lập hoặc báo cáo tích hợp liên tục trong suốt 6 năm qua."
    })


def build_slide_06_result1_similarity(prs):
    """Slide 6: Kết quả 1 - Phân phối Tương đồng SDG & BẢNG ĐỐI CHUẨN THỐNG KÊ."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 1: PHÂN PHỐI ĐIỂM TƯƠNG ĐỒNG SDG & ĐỐI CHUẨN PAPER GỐC", "KẾT QUẢ THỰC NGHIỆM", 6)
    
    # Ảnh biểu đồ phân phối bên trái
    sim_img = FIGURES_DIR / "similarity_hist.png"
    if sim_img.exists():
        slide.shapes.add_picture(str(sim_img), Inches(0.8), Inches(1.35), Inches(5.8), Inches(4.3))
        
    # Bảng Native Table đối chuẩn thống kê bên phải
    rows = 6
    cols = 3
    table_shape = slide.shapes.add_table(rows, cols, Inches(6.8), Inches(1.35), Inches(5.733), Inches(4.3))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(2.3)
    tbl.columns[1].width = Inches(1.7)
    tbl.columns[2].width = Inches(1.733)
    
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
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = C_NAVY_PRIMARY
            elif j == 2:
                r.font.bold = True
                r.font.color.rgb = C_GREEN_EMERALD
            else:
                r.font.color.rgb = C_TEXT_DARK

    # Bottom takeaway card
    bot_card = add_card(slide, 0.8, 5.75, 11.733, 1.25, RGBColor(238, 244, 252), C_BLUE_ACCENT)
    tb_b = slide.shapes.add_textbox(Inches(1.0), Inches(5.82), Inches(11.333), Inches(1.1))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p_b1 = tf_b.paragraphs[0]
    r_b1 = p_b1.add_run("KẾT LUẬN KHOA HỌC: ")
    r_b1.font.bold = True
    r_b1.font.size = Pt(10)
    r_b1.font.color.rgb = C_NAVY_PRIMARY
    r_b2 = p_b1.add_run(
        "Sự tương đồng gần như tuyệt đối giữa phân phối điểm của Việt Nam (μ = 45,43, σ = 11,87) và nghiên cứu quốc tế của Kang & Kim (μ ≈ 44,80, σ ≈ 11,50) "
        "chứng minh rằng quy trình xử lý bằng `vietnamese-sbert` đạt độ chuẩn hóa toán học hoàn hảo, "
        "xóa bỏ hoàn toàn định kiến cho rằng các mô hình nhúng tiếng Việt có độ phân tán kém hơn tiếng Anh."
    )
    r_b2.font.size = Pt(9.0)
    r_b2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Dẫn chứng phân phối hình chuông và bảng đối chuẩn thống kê chứng minh mô hình đạt chuẩn quốc tế.",
        "script": "Hình bên trái là phân phối tần suất điểm tương đồng SDG của 96.461 câu văn bản. Đồ thị có dạng hình chuông đối xứng chuẩn Gaussian, đỉnh tập trung quanh 45 điểm. Bảng bên phải cho thấy điểm trung bình của Việt Nam là 45,43 so với 44,80 của Kang & Kim, độ lệch chỉ 0,63 điểm. Điều này khẳng định quy trình NLP đạt độ tin cậy tương đương 100% chuẩn quốc tế.",
        "highlights": "μ = 45,43 (lệch chỉ 0,63 điểm so với bài gốc), σ = 11,87 (phân phối chuẩn).",
        "qa": "Tại sao điểm trung bình lại quanh 45 điểm mà không phải 70-80: Thang đo Min-Max toàn cục 0-100 ánh xạ toàn bộ phân phối cosine, mức 45 là trung vị chuẩn của phân phối Gaussian khi chiếu văn bản tự do vào 17 mục tiêu."
    })


def build_slide_07_result2_heatmap(prs):
    """Slide 7: Kết quả 2 - Cấu trúc 6 Nhóm SDG qua Heatmap."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 2: CẤU TRÚC 6 NHÓM SDG QUA HEATMAP & QUY LUẬT TOÀN CẦU", "KẾT QUẢ THỰC NGHIỆM", 7)
    
    hm_img = FIGURES_DIR / "heatmap_6cat.png"
    if hm_img.exists():
        slide.shapes.add_picture(str(hm_img), Inches(0.8), Inches(1.35), Inches(6.5), Inches(5.65))
        
    # 3 Thẻ nhận xét bên phải
    rw = 5.0
    rx = 7.533
    insights = [
        ("TRẬT TỰ BẤT BIẾN TOÀN CẦU", C_NAVY_PRIMARY, [
            "• Thứ tự ưu tiên phản ánh đúng lý thuyết Manfred Max-Neef:",
            "  Economic (49,85đ) > Social > Resources ≈ Life > Environments > Equity (43,45đ).",
            "• Khớp 100% với trật tự thực nghiệm quốc tế của Kang & Kim (2022)."
        ]),
        ("XU HƯỚNG TĂNG TRƯỞNG THEO NĂM", C_GREEN_EMERALD, [
            "• Mức độ công bố tăng dần từ 2020 (vàng nhạt) sang 2025 (đỏ đậm).",
            "• Điểm trung bình toàn mẫu tăng từ 42,1 (2020) lên 49,8 (2025).",
            "• Minh chứng rõ nét cho tác động thúc đẩy của Thông tư 96/2020."
        ]),
        ("VÙNG TRŨNG CÔNG BẰNG (EQUITY)", C_RED_ACCENT, [
            "• Nhóm Công bằng (SDG 4, 5, 10) luôn có điểm số thấp nhất toàn mẫu.",
            "• Doanh nghiệp Việt Nam ưu tiên các chỉ tiêu tăng trưởng tài chính và việc làm hơn là các cam kết bình đẳng giới và thu hẹp khoảng cách."
        ])
    ]
    for i, (ititle, icol, ibullets) in enumerate(insights):
        iy = 1.35 + i * 1.9
        add_card(slide, rx, iy, rw, 1.75, C_CARD_BG, icol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.15), Inches(iy + 0.1), Inches(rw - 0.3), Inches(1.55))
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


def build_slide_08_result3_companies_p1(prs):
    """Slide 8: Đặc thù Ngành - Khối Sản xuất & Năng lượng (VNM, VCS, PAN, PLX)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "ĐẶC THÙ NGÀNH: KHỐI SẢN XUẤT & NĂNG LƯỢNG (VNM, VCS, PAN, PLX)", "ĐẶC THÙ NGÀNH DOANH NGHIỆP", 8)
    
    img_p1 = FIGURES_DIR / "slide_manuf_sdgs.png"
    if img_p1.exists():
        slide.shapes.add_picture(str(img_p1), Inches(0.8), Inches(1.35), Inches(6.5), Inches(5.65))
        
    rw = 5.0
    rx = 7.533
    co_insights = [
        ("VICOSTONE (VCS) — DẪN ĐẦU SDG 9 & 12", C_NAVY_PRIMARY, 
         "• SDG 9 (52,96đ) & SDG 12 (53,18đ) cao nhất toàn khối sản xuất.\n"
         "• Lý do kinh doanh: Công nghệ rung ép Breton, tự chủ >95% nguyên liệu thạch anh, tái chế 100% bùn thải đá và nước tuần hoàn."),
        ("VINAMILK (VNM) — BỨT PHÁ SDG 13 KHÍ HẬU", C_GREEN_EMERALD, 
         "• SDG 13 (45,36đ) & Nhóm Môi trường tăng vọt (+6,06đ qua 6 năm).\n"
         "• Lý do kinh doanh: Tiên phong lộ trình Net Zero 2050, 3 đơn vị đạt chứng nhận PAS 2060, nông nghiệp tái sinh Green Farm."),
        ("PAN GROUP (PAN) — TRỤ CỘT SDG 2 NÔNG NGHIỆP", C_GOLD_ACCENT, 
         "• SDG 2 (47,47đ) & SDG 12 (49,01đ) dẫn đầu mẫu nghiên cứu.\n"
         "• Lý do kinh doanh: Chuỗi giá trị nông nghiệp khép kín từ giống cây trồng (Vinaseed), tôm sạch sinh thái (Fimex) đến chế biến hạt."),
        ("PETROLIMEX (PLX) — CHUYỂN ĐỔI SDG 7 & 13", C_RED_ACCENT, 
         "• Trọng tâm SDG 7 Năng lượng (48,04đ) và SDG 13 Khí hậu (46,19đ).\n"
         "• Lý do kinh doanh: Nhiên liệu sạch Euro 5, điện mặt trời áp mái cây xăng, kiểm kê khí nhà kính Scope 1-2 theo ISO 14064-1.")
    ]
    for i, (ctitle, ccol, ctext) in enumerate(co_insights):
        cy = 1.35 + i * 1.42
        add_card(slide, rx, cy, rw, 1.32, C_CARD_BG, ccol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.15), Inches(cy + 0.08), Inches(rw - 0.3), Inches(1.16))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(ctitle + "\n")
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = ccol
        
        p2 = tf.add_paragraph()
        r2 = p2.add_run(ctext)
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Chứng minh mô hình NLP bắt đúng bản chất kinh doanh: giải thích tại sao doanh nghiệp này lại có goal này cao vượt trội.",
        "script": "Biểu đồ bên trái minh họa điểm số các SDG cốt lõi của khối Sản xuất & Năng lượng năm 2025. Mô hình NLP phản ánh cực kỳ nhạy bén bản chất kinh doanh: Vicostone dẫn đầu SDG 9 và 12 nhờ công nghệ Breton và tuần hoàn bùn thải; Vinamilk bứt phá SDG 13 nhờ chứng nhận PAS 2060 trung hòa carbon; PAN dẫn đầu SDG 2 nhờ chuỗi lúa gạo tôm sạch; và Petrolimex tập trung SDG 7 và 13 nhờ nhiên liệu Euro 5.",
        "highlights": "VCS cao nhất SDG 9 & 12; VNM bứt phá SDG 13; PAN dẫn đầu SDG 2; PLX trọng tâm SDG 7 & 13.",
        "qa": "Tại sao PLX điểm SDG 13 lại cao: Vì Petrolimex chịu áp lực chuyển đổi năng lượng hóa thạch lớn nhất, bắt buộc phải kiểm kê phát thải ISO 14064-1."
    })


def build_slide_09_result3_companies_p2(prs):
    """Slide 9: Đặc thù Ngành - Khối Tài chính & Bán lẻ (PNJ, BVH, SSI)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "ĐẶC THÙ NGÀNH: KHỐI TÀI CHÍNH & BÁN LẺ (PNJ, BVH, SSI)", "ĐẶC THÙ NGÀNH DOANH NGHIỆP", 9)
    
    img_p2 = FIGURES_DIR / "slide_finance_sdgs.png"
    if img_p2.exists():
        slide.shapes.add_picture(str(img_p2), Inches(0.8), Inches(1.35), Inches(6.5), Inches(5.65))
        
    rw = 5.0
    rx = 7.533
    fin_insights = [
        ("PNJ — ĐIỂM SÁNG SDG 5 BÌNH ĐẲNG GIỚI (40,70đ)", C_PURPLE_ACCENT, 
         "• Dẫn đầu tuyệt đối toàn bộ 7 doanh nghiệp ở SDG 5 Bình đẳng giới.\n"
         "• Lý do kinh doanh: Lao động nữ chiếm >60%, tỷ lệ lãnh đạo nữ vượt trội, tôn chỉ kinh doanh tôn vinh vẻ đẹp phụ nữ và chiến lược hòa nhập DE&I."),
        ("BẢO VIỆT (BVH) — DẪN ĐẦU SDG 17 HỢP TÁC (53,32đ)", C_NAVY_PRIMARY, 
         "• Dẫn đầu toàn mẫu ở SDG 17 Đối tác phát triển & SDG 8 Tăng trưởng.\n"
         "• Lý do kinh doanh: Tiên phong áp dụng Khung Báo cáo Tích hợp Quốc tế <IIRC> từ 2015, triển khai bảo hiểm vi mô nông nghiệp bảo vệ nông dân."),
        ("CHỨNG KHOÁN SSI (SSI) — TÀI CHÍNH XANH SDG 8 & 9", C_BLUE_ACCENT, 
         "• Đạt đỉnh ở SDG 8 Việc làm (50,91đ) & SDG 9 Đổi mới hạ tầng (48,93đ).\n"
         "• Lý do kinh doanh: Thu xếp các gói vốn trái phiếu xanh quốc tế (IFC), số hóa 100% giao dịch iBoard và tài trợ giáo dục tài chính cộng đồng.")
    ]
    for i, (ctitle, ccol, ctext) in enumerate(fin_insights):
        cy = 1.35 + i * 1.9
        add_card(slide, rx, cy, rw, 1.75, C_CARD_BG, ccol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.15), Inches(cy + 0.1), Inches(rw - 0.3), Inches(1.55))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(ctitle + "\n")
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = ccol
        
        p2 = tf.add_paragraph()
        r2 = p2.add_run(ctext)
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Làm nổi bật nét đặc thù của khối Tài chính và Bán lẻ, giải thích tại sao PNJ dẫn đầu SDG 5 và BVH dẫn đầu SDG 17.",
        "script": "Khối Dịch vụ Tài chính và Bán lẻ thể hiện bản đồ SDG hoàn toàn khác biệt: PNJ là doanh nghiệp duy nhất trong toàn bộ mẫu đạt điểm cao vượt trội ở SDG 5 Bình đẳng giới (40,70 điểm) nhờ đặc thù bán lẻ trang sức và nhân sự nữ chiếm đa số. Bảo Việt đạt điểm kỷ lục ở SDG 17 (53,32 điểm) nhờ áp dụng Báo cáo Tích hợp IIRC. SSI bứt phá ở SDG 8 và 9 qua các thương vụ thu xếp trái phiếu xanh.",
        "highlights": "PNJ dẫn đầu toàn mẫu ở SDG 5 (40,70đ); BVH dẫn đầu SDG 17 (53,32đ); SSI mạnh về tài chính xanh.",
        "qa": "Tại sao điểm bình đẳng giới của PNJ lại cao hơn các công ty khác: Tỷ lệ nữ nhân sự PNJ trên 60% và công ty lồng ghép yếu tố phụ nữ vào chiến lược thương hiệu."
    })


def build_slide_10_result4_trends(prs):
    """Slide 10: Kết quả 4 - Xu hướng Dịch chuyển Chuỗi Thời gian (2020–2025)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 4: XU HƯỚNG DỊCH CHUYỂN CHUỖI THỜI GIAN (2020–2025)", "KẾT QUẢ THỰC NGHIỆM", 10)
    
    tr_img = FIGURES_DIR / "trends_6categories.png"
    if tr_img.exists():
        slide.shapes.add_picture(str(tr_img), Inches(0.8), Inches(1.35), Inches(6.5), Inches(5.65))
        
    rw = 5.0
    rx = 7.533
    phases = [
        ("GIAI ĐOẠN 2020–2021: ĐỐI PHÓ ĐẠI DỊCH", C_NAVY_PRIMARY, [
            "• Điểm số phân hóa mạnh, nhóm Life và Social tăng đột biến.",
            "• Doanh nghiệp tập trung duy trì chuỗi cung ứng, an toàn lao động và chế độ lương thưởng ứng phó Covid-19."
        ]),
        ("MỐC 2022: CHUYỂN HƯỚNG THEO THÔNG TƯ 96", C_BLUE_ACCENT, [
            "• Số lượng báo cáo phát hành tăng vọt sau khi Thông tư 96 có hiệu lực.",
            "• Điểm số các nhóm bắt đầu hội tụ theo cấu trúc tiêu chuẩn GRI."
        ]),
        ("GIAI ĐOẠN 2024–2025: BÙNG NỔ NET ZERO", C_GREEN_EMERALD, [
            "• Tất cả 7 doanh nghiệp đều đạt điểm số cao nhất trong lịch sử.",
            "• Nhóm Môi trường (+6,06đ) và Tài nguyên (+5,42đ) tăng tốc mạnh mẽ nhờ các cam kết trung hòa carbon cụ thể."
        ])
    ]
    for i, (ptitle, pcol, pbullets) in enumerate(phases):
        py = 1.35 + i * 1.9
        add_card(slide, rx, py, rw, 1.75, C_CARD_BG, pcol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.15), Inches(py + 0.1), Inches(rw - 0.3), Inches(1.55))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(ptitle + "\n")
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = pcol
        
        for bullet in pbullets:
            p2 = tf.add_paragraph()
            p2.space_after = Pt(2)
            r2 = p2.add_run(bullet)
            r2.font.size = Pt(8.8)
            r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Phân tích xu hướng chuyển dịch 6 nhóm SDG qua 3 giai đoạn: Covid (2020-2021), Thể chế hóa (2022), Bùng nổ Net Zero (2024-2025).",
        "script": "Đồ thị đường chuỗi thời gian cho thấy bước chuyển dịch rõ nét qua 3 giai đoạn: 2020-2021 tập trung ứng phó đại dịch; 2022 tái định hình theo Thông tư 96; và 2024-2025 bùng nổ mạnh mẽ với các cam kết Net Zero, đưa điểm số toàn bộ 7 doanh nghiệp lên mức cao kỷ lục.",
        "highlights": "3 giai đoạn chuyển dịch; năm 2025 điểm số cao nhất lịch sử toàn mẫu.",
        "qa": "Xu hướng này phản ánh điều gì: Phản ánh nhận thức doanh nghiệp chuyển từ đối phó sang tích hợp ESG vào chiến lược cốt lõi."
    })


def build_slide_11_result5_sentiment(prs):
    """Slide 11: Kết quả 5 - Sắc thái Cảm xúc PhoBERT & Đối chuẩn Paper Gốc."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 5: SẮC THÁI CẢM XÚC & ĐỐI CHUẨN PAPER GỐC", "KẾT QUẢ THỰC NGHIỆM", 11)
    
    img_s1 = FIGURES_DIR / "sentiment_hist.png"
    img_s2 = FIGURES_DIR / "sentiment_by_company.png"
    if img_s1.exists():
        slide.shapes.add_picture(str(img_s1), Inches(0.8), Inches(1.35), Inches(5.8), Inches(2.75))
    if img_s2.exists():
        slide.shapes.add_picture(str(img_s2), Inches(0.8), Inches(4.2), Inches(5.8), Inches(2.8))
        
    rw = 5.0
    rx = 6.8
    card_r = add_card(slide, rx, 1.35, rw, 5.65)
    tb_r = slide.shapes.add_textbox(Inches(rx + 0.15), Inches(1.5), Inches(rw - 0.3), Inches(5.35))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    
    p0 = tf_r.paragraphs[0]
    r0 = p0.add_run("SO SÁNH CƠ CẤU CẢM XÚC VỚI KANG & KIM (2022)\n")
    r0.font.bold = True
    r0.font.size = Pt(11)
    r0.font.color.rgb = C_NAVY_PRIMARY
    
    points = [
        ("• Paper gốc (DistilBERT 2 lớp):", True, C_NAVY_PRIMARY),
        ("  Tích cực chiếm ~78%, Tiêu cực chiếm ~15%. Không có lớp Trung tính, ép câu số liệu vào cảm xúc.", False, C_TEXT_DARK),
        ("• Nghiên cứu tại Việt Nam (PhoBERT 3 lớp):", True, C_GREEN_EMERALD),
        ("  - Tích cực (Positive): 53,87%\n  - Trung tính (Neutral): 32,87%\n  - Tiêu cực (Negative): 13,26%", False, C_TEXT_DARK),
        ("• Ý nghĩa học thuật của Lớp Trung tính (32,87%):", True, C_BLUE_ACCENT),
        ("  Chiếm gần 1/3 dung lượng báo cáo, bảo toàn các câu số liệu kỹ thuật khách quan (ví dụ: lượng điện kWh, nước m3, khí phát thải tCO2e).", False, C_TEXT_DARK),
        ("• Tính thận trọng trong văn phong tiếng Việt:", True, C_PURPLE_ACCENT),
        ("  Doanh nghiệp Việt Nam công bố nhiều dữ liệu đo lường theo chuẩn mực GRI chứ không chỉ đơn thuần quảng cáo PR.", False, C_TEXT_DARK)
    ]
    for head, is_h, col in points:
        p = tf_r.add_paragraph()
        p.space_after = Pt(3)
        r = p.add_run(head)
        r.font.name = FONT_MAIN
        r.font.size = Pt(9.5 if is_h else 8.8)
        r.font.bold = is_h
        r.font.color.rgb = col

    set_presenter_notes(slide, {
        "goal": "Làm nổi bật đóng góp của mô hình PhoBERT 3 lớp, đặc biệt là 32,87% câu Trung tính.",
        "script": "Khác với Kang & Kim dùng DistilBERT 2 lớp ép các câu số liệu vào nhãn tích cực hoặc tiêu cực, đề tài sử dụng PhoBERT 3 lớp. Kết quả cho thấy lớp Trung tính chiếm tới 32,87%, khẳng định báo cáo doanh nghiệp Việt Nam dành gần 1/3 dung lượng cho các số liệu kỹ thuật đo đạc khách quan.",
        "highlights": "PhoBERT 3 lớp: Pos 53,87%, Neutral 32,87%, Neg 13,26%.",
        "qa": "Lớp trung tính có ý nghĩa gì: Tránh sai lệch trong việc gán nhãn cảm xúc cho các câu số liệu kiểm toán kỹ thuật."
    })


def build_slide_12_result6_sentiment_ratio(prs):
    """Slide 12: Kết quả 6 - Tỷ số Cảm xúc Pos/Neg Ratio (Đối chuẩn Tỷ số)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "KẾT QUẢ 6: TỶ SỐ CẢM XÚC POS/NEG RATIO & ĐỐI CHUẨN TỶ LỆ", "KẾT QUẢ THỰC NGHIỆM", 12)
    
    img_r = FIGURES_DIR / "sentiment_ratio.png"
    if img_r.exists():
        slide.shapes.add_picture(str(img_r), Inches(0.8), Inches(1.35), Inches(6.5), Inches(5.65))
        
    rw = 5.0
    rx = 7.533
    r_insights = [
        ("ĐỐI CHUẨN TỶ SỐ VỚI KANG & KIM (2022)", C_NAVY_PRIMARY, [
            "• Paper gốc: Pos/Neg Ratio bình quân đạt ≈ 5,20 lần.",
            "• Việt Nam: Pos/Neg Ratio trung bình đạt 4,06 lần (dao động 3,0 – 6,7 lần tùy năm).",
            "• Cả hai đều khẳng định sự tồn tại của Thiên lệch Lạc quan (Optimism Bias) mang tính phổ quát."
        ]),
        ("QUY LUẬT QUẢN TRỊ ẤN TƯỢNG (IMPRESSION MANAGEMENT)", C_RED_ACCENT, [
            "• Tỷ số 4,06 lần xác nhận Doanh nghiệp Việt Nam có xu hướng khuếch đại thành tựu và giảm thiểu thông tin rủi ro.",
            "• Báo cáo thường lồng ghép các từ ngữ mang tính ca ngợi thành tích thay vì giải trình khách quan."
        ]),
        ("CASE STUDY NGOẠI LỆ: PNJ 2022 & VNM 2025", C_BLUE_ACCENT, [
            "• PNJ 2022 giảm còn 1,21 lần: Do phản ánh trung thực khó khăn giãn cách Covid đóng cửa chuỗi cửa hàng.",
            "• VNM 2025 đạt 6,71 lần: Bùng nổ công bố thành tựu Net Zero và giải thưởng quốc tế."
        ])
    ]
    for i, (rtitle, rcol, rbullets) in enumerate(r_insights):
        ry = 1.35 + i * 1.9
        add_card(slide, rx, ry, rw, 1.75, C_CARD_BG, rcol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.15), Inches(ry + 0.1), Inches(rw - 0.3), Inches(1.55))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(rtitle + "\n")
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = rcol
        
        for bullet in rbullets:
            p2 = tf.add_paragraph()
            p2.space_after = Pt(2)
            r2 = p2.add_run(bullet)
            r2.font.size = Pt(8.8)
            r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Phân tích Tỷ số Pos/Neg Ratio, đối chuẩn 4,06x của VN với 5,20x của Kang & Kim, khẳng định hiệu ứng Pollyanna.",
        "script": "Đồ thị bên trái biểu diễn Tỷ số Cảm xúc Pos/Neg theo từng năm. Nghiên cứu của Kang & Kim cho tỷ số toàn cầu khoảng 5,2 lần, còn tại Việt Nam tỷ số đạt 4,06 lần. Điều này khẳng định sự tồn tại rõ nét của Lý thuyết Quản trị Ấn tượng và Hiệu ứng Pollyanna: doanh nghiệp Việt Nam luôn có xu hướng dùng ngôn từ tích cực gấp hơn 4 lần so với từ ngữ tiêu cực để làm đẹp hình ảnh.",
        "highlights": "Pos/Neg trung bình = 4,06 lần; Kang & Kim = 5,20 lần; Khẳng định Quản trị Ấn tượng.",
        "qa": "Tỷ số 4,06 lần có phải là tẩy xanh không: Tỷ số cao cho thấy thiên lệch lạc quan, nhưng để kết luận tẩy xanh thì cần đối chiếu chéo với số liệu kiểm toán thực tế."
    })


def build_slide_13_discussion_talk_heavy(prs):
    """Slide 13: Thảo luận - Doanh nghiệp 'Nói nhiều về gì' (NHIỀU ẢNH: BIỂU ĐỒ TỔNG HỢP 7 CTY)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "THẢO LUẬN: DOANH NGHIỆP VIỆT NAM 'NÓI NHIỀU VỀ GÌ'?", "THẢO LUẬN CHUYÊN SÂU", 13)
    
    # Nhúng ảnh tổng hợp top SDGs của 7 doanh nghiệp
    comp_img = FIGURES_DIR / "slide_company_top_sdgs.png"
    if comp_img.exists():
        slide.shapes.add_picture(str(comp_img), Inches(0.8), Inches(1.35), Inches(6.5), Inches(5.65))
        
    rw = 5.0
    rx = 7.533
    talk_insights = [
        ("1. TĂNG TRƯỞNG & VIỆC LÀM (SDG 8, 9, 12)", C_NAVY_PRIMARY, [
            "• Chiếm dung lượng áp đảo trong báo cáo của toàn bộ 7 doanh nghiệp.",
            "• Doanh nghiệp tập trung mô tả doanh thu, lợi nhuận, quy mô nhân sự và ứng dụng công nghệ vì đây là các số liệu dễ định lượng và phục vụ quan hệ cổ đông."
        ]),
        ("2. HOẠT ĐỘNG THIỆN NGUYỆN CSR (SDG 1, 2, 4)", C_GREEN_EMERALD, [
            "• Doanh nghiệp nói rất nhiều về các gói tài trợ học bổng, xây cầu từ thiện, cứu trợ bão lũ.",
            "• Bản chất: Dễ thực hiện, hiệu quả truyền thông tức thì, giúp củng cố tính chính danh xã hội."
        ]),
        ("3. ĐỐI TÁC & BÁO CÁO TÍCH HỢP (SDG 16, 17)", C_BLUE_ACCENT, [
            "• BVH, SSI và Vinamilk đầu tư dung lượng lớn cho việc tuân thủ pháp lý, đối tác chuỗi cung ứng và Báo cáo Tích hợp.",
            "• Phục vụ việc thu hút dòng vốn đầu tư ngoại và xếp hạng tín nhiệm."
        ])
    ]
    for i, (ttitle, tcol, tbullets) in enumerate(talk_insights):
        ty = 1.35 + i * 1.9
        add_card(slide, rx, ty, rw, 1.75, C_CARD_BG, tcol)
        
        tb = slide.shapes.add_textbox(Inches(rx + 0.15), Inches(ty + 0.1), Inches(rw - 0.3), Inches(1.55))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run(ttitle + "\n")
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = tcol
        
        for bullet in tbullets:
            p2 = tf.add_paragraph()
            p2.space_after = Pt(2)
            r2 = p2.add_run(bullet)
            r2.font.size = Pt(8.8)
            r2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Giải thích hiện tượng 'Nói nhiều về gì': Tăng trưởng kinh tế, CSR từ thiện bề nổi và Báo cáo tích hợp.",
        "script": "Từ biểu đồ tổng hợp top SDGs của cả 7 doanh nghiệp, chúng ta thấy rõ hiện tượng Cherry-picking: Doanh nghiệp tập trung nói nhiều vào 3 mảng: Thứ nhất là Tăng trưởng kinh tế và việc làm (SDG 8, 9, 12); thứ hai là các hoạt động từ thiện CSR bề nổi như học bổng, cứu trợ (SDG 1, 2, 4); và thứ ba là quan hệ đối tác pháp lý (SDG 16, 17). Đây là những chủ đề mang lại lợi ích PR tức thì và phục vụ quan hệ nhà đầu tư.",
        "highlights": "Nói nhiều: Kinh tế (SDG 8, 9, 12), Từ thiện CSR (SDG 1, 2, 4), Đối tác (SDG 16, 17).",
        "qa": "Tại sao gọi đây là Cherry-picking: Vì doanh nghiệp có quyền chủ động chọn lọc những tiêu chí đẹp nhất để trình bày mà không bị chế tài bắt buộc."
    })


def build_slide_14_discussion_rarely_talk(prs):
    """Slide 14: Thảo luận - Doanh nghiệp 'Ít nói về gì' (4 Vùng Né Tránh Trọng Tâm)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "THẢO LUẬN: DOANH NGHIỆP 'ÍT NÓI VỀ GÌ' (VÙNG TRŨNG NÉ TRÁNH)?", "THẢO LUẬN CHUYÊN SÂU", 14)
    
    avoid_data = [
        ("1. ĐA DẠNG SINH HỌC & BẢO TỒN (SDG 14, 15)", C_RED_ACCENT, [
            "• Điểm số luôn ở mức đáy toàn mẫu (< 38 điểm).",
            "• Báo cáo chỉ nêu khẩu hiệu chung chung, thiếu hoàn toàn số liệu kiểm kê tác động sinh thái đất liền và tài nguyên nước.",
            "• Nguyên nhân: Chi phí đo lường đa dạng sinh học phức tạp và chưa có chế tài bắt buộc."
        ]),
        ("2. BÌNH ĐẲNG LƯƠNG & CHÊNH LỆCH THU NHẬP (SDG 5, 10)", C_RED_ACCENT, [
            "• Tuyệt đối né tránh công bố Tỷ số chênh lệch lương CEO với công nhân (CEO-to-worker pay ratio).",
            "• Không có thống kê khoảng cách thu nhập theo giới tính ở cùng cấp bậc chuyên môn.",
            "• Báo cáo chỉ dừng lại ở tỷ lệ % lao động nữ chung chung."
        ]),
        ("3. PHÁT THẢI CHUỖI CUNG ỨNG SCOPE 3", C_RED_ACCENT, [
            "• Mới chỉ đo lường phát thải trực tiếp Scope 1 và điện Scope 2.",
            "• Phát thải gián tiếp chuỗi cung ứng Scope 3 (chiếm 70–80% tổng lượng phát thải thực tế) gần như bị bỏ ngỏ.",
            "• Do chuỗi cung ứng phân tán và thiếu công cụ đo lường chuyên sâu."
        ]),
        ("4. SỰ CỐ TIÊU CỰC, TRANH CHẤP & XỬ PHẠT (SDG 16)", C_RED_ACCENT, [
            "• 'Gạn đục khơi trong' điển hình: Không có báo cáo nào ghi nhận tai nạn lao động nghiêm trọng hay khiếu nại khách hàng.",
            "• Các quyết định xử phạt vi phạm hành chính về môi trường hoặc thuế bị che giấu hoàn toàn.",
            "• Báo cáo trở thành tài liệu tiếp thị thay vì công cụ quản trị rủi ro."
        ])
    ]
    
    w = 5.75
    h = 2.6
    top1 = 1.35
    top2 = 4.15
    l1 = 0.8
    l2 = 6.78
    coords = [(l1, top1), (l2, top1), (l1, top2), (l2, top2)]
    
    for i, (atitle, acol, abullets) in enumerate(avoid_data):
        cx, cy = coords[i]
        card = add_card(slide, cx, cy, w, h, C_CARD_BG, acol)
        
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


def build_slide_15_comparison(prs):
    """Slide 15: BẢNG SO SÁNH ĐỐI ĐẦU TOÀN DIỆN: PAPER GỐC VS CODE MỚI (YÊU CẦU TRỌNG TÂM)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "BẢNG SO SÁNH ĐỐI ĐẦU: PAPER GỐC (KANG & KIM 2022) VS CODE MỚI", "ĐỐI CHUẨN TOÀN DIỆN", 15)
    
    rows = 9
    cols = 4
    left = Inches(0.8)
    top = Inches(1.28)
    width = Inches(11.733)
    height = Inches(5.65)
    
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


def build_slide_16_implications(prs):
    """Slide 16: Hàm Ý Thực Tiễn & Đề Xuất Chính Sách (Tinh gọn text)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "HÀM Ý THỰC TIỄN & ĐỀ XUẤT CHO CÁC BÊN LIÊN QUAN", "HÀM Ý QUẢN TRỊ", 16)
    
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
    h = 5.65
    
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
            p.space_after = Pt(4)
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


def build_slide_17_limitations_future(prs):
    """Slide 17: Hạn Chế & Tương Lai (NHIỀU ẢNH: SƠ ĐỒ MULTI-AGENT ESG AUDITOR)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "HẠN CHẾ ĐỀ TÀI & HƯỚNG PHÁT TRIỂN AGENTIC ESG AUDITOR", "HƯỚNG PHÁT TRIỂN", 17)
    
    # Bên trái: 3 Hạn chế hiện tại (Card ngắn)
    card_l = add_card(slide, 0.8, 1.35, 5.0, 5.65)
    tb_l = slide.shapes.add_textbox(Inches(0.95), Inches(1.5), Inches(4.7), Inches(5.35))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    
    p0 = tf_l.paragraphs[0]
    r0 = p0.add_run("HẠN CHẾ HIỆN TẠI CỦA NGHIÊN CỨU\n")
    r0.font.bold = True
    r0.font.size = Pt(11)
    r0.font.color.rgb = C_NAVY_PRIMARY
    
    lims = [
        ("• Giới hạn quy mô mẫu:", True, C_NAVY_PRIMARY),
        ("  Mẫu 7 tập đoàn lớn hàng đầu có báo cáo liên tục; chưa bao phủ toàn bộ 300+ doanh nghiệp niêm yết trên HOSE/HNX.", False, C_TEXT_DARK),
        ("• Chưa Fact-Checking số liệu bảng biểu:", True, C_RED_ACCENT),
        ("  SBERT đo lường mức độ tương đồng ngữ nghĩa văn bản, nhưng chưa thể tự động kiểm chứng chéo các số liệu định lượng trong Table.", False, C_TEXT_DARK),
        ("• Chưa bóc tách ngữ cảnh đoạn văn rộng:", True, C_PURPLE_ACCENT),
        ("  Chunking theo câu đơn đôi khi làm mất ngữ cảnh của đoạn văn cha.", False, C_TEXT_DARK)
    ]
    for head, is_h, col in lims:
        p = tf_l.add_paragraph()
        p.space_after = Pt(3)
        r = p.add_run(head)
        r.font.name = FONT_MAIN
        r.font.size = Pt(9.5 if is_h else 8.8)
        r.font.bold = is_h
        r.font.color.rgb = col
        
    # Bên phải: Nhúng Sơ đồ Multi-Agent ESG Auditor
    agent_img = FIGURES_DIR / "rag_agentic_flow.png"
    if agent_img.exists():
        slide.shapes.add_picture(str(agent_img), Inches(6.0), Inches(1.35), Inches(6.533), Inches(4.3))
        
    # Hộp tóm tắt tầm nhìn tương lai
    bot_r = add_card(slide, 6.0, 5.75, 6.533, 1.25, RGBColor(238, 244, 252), C_GREEN_EMERALD)
    tb_br = slide.shapes.add_textbox(Inches(6.15), Inches(5.82), Inches(6.233), Inches(1.1))
    tf_br = tb_br.text_frame
    tf_br.word_wrap = True
    p_br = tf_br.paragraphs[0]
    r_br1 = p_br.add_run("TẦM NHÌN: HỆ THỐNG TRỢ LÝ KIỂM TOÁN AI (AGENTIC ESG AUDITOR)\n")
    r_br1.font.bold = True
    r_br1.font.size = Pt(10)
    r_br1.font.color.rgb = C_GREEN_EMERALD
    r_br2 = p_br.add_run(
        "Kế thừa pipeline Retrieval hiện tại làm nền móng sạch để tích hợp Multi-Agent: "
        "Agent đọc bảng biểu, Agent kiểm tra chéo cam kết và LLM (GPT-4o/Gemini) xuất báo cáo thẩm định độc lập 3 trang trong 30 giây."
    )
    r_br2.font.size = Pt(8.8)
    r_br2.font.color.rgb = C_TEXT_DARK

    set_presenter_notes(slide, {
        "goal": "Thẳng thắn nhìn nhận hạn chế và vạch ra lộ trình tích hợp Multi-Agent AI trong tương lai.",
        "script": "Chúng em nhìn nhận rõ 3 hạn chế: Mẫu 7 tập đoàn, chưa kiểm chứng chéo số liệu bảng biểu, và chunking câu đơn. Tuy nhiên sơ đồ bên phải chỉ ra hướng phát triển tự nhiên: dùng pipeline hiện tại làm nền móng sạch để cắm thêm các Agent kiểm tra số liệu bảng biểu và đưa vào LLM sinh báo cáo kiểm toán độc lập.",
        "highlights": "Thẳng thắn nhìn nhận hạn chế; Mở ra tầm nhìn Agentic ESG Auditor.",
        "qa": "Tại sao không dùng LLM ngay từ đầu: Vì LLM nhồi cả cuốn PDF rất đắt và dễ ảo giác, dùng SBERT làm Retrieval sạch trước mới đảm bảo độ chính xác."
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
        "script": "Kính thưa Quý Thầy Cô trong Hội đồng, đề tài gửi gắm 4 thông điệp cốt lõi: 1. Chuyển giao thành công sang tiếng Việt; 2. Phản ánh đúng bản chất ngành; 3. Bóc trần sự đối lập giữa những gì nói nhiều và né tránh; 4. Xác nhận thiên lệch lạc quan và nhu cầu công cụ AI kiểm toán. Nhóm tác giả xin trân trọng cảm ơn Quý Thầy Cô!",
        "highlights": "Cúi đầu chào và mời các thầy cô đặt câu hỏi. Giữ phong thái tự tin, khiêm tốn.",
        "qa": "Sẵn sàng mở lại các slide biểu đồ tương ứng khi thầy cô yêu cầu giải trình."
    })


def main():
    print("=" * 80)
    print("BẮT ĐẦU CẬP NHẬT BỘ SLIDE CHÍNH (Ít chữ, nhiều ảnh & bảng so sánh đối đầu)...")
    print("=" * 80)
    
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    print("[1/18] Slide 1: Trang Tiêu đề & Thông tin Tác giả...")
    build_slide_01_title(prs)
    
    print("[2/18] Slide 2: Đặt vấn đề & Bối cảnh Thể chế Việt Nam (Thẻ KPI)...")
    build_slide_02_context(prs)
    
    print("[3/18] Slide 3: Bài báo gốc Kang & Kim (2022) vs Code mới (NHIỀU ẢNH: Biểu đồ 3 Panel)...")
    build_slide_03_original_paper(prs)
    
    print("[4/18] Slide 4: Khung Phương pháp luận (NHIỀU ẢNH: Sơ đồ luồng 4 tầng)...")
    build_slide_04_pipeline(prs)
    
    print("[5/18] Slide 5: Mẫu Dữ liệu Thực nghiệm (BẢNG NATIVE TABLE 7 Doanh nghiệp)...")
    build_slide_05_sample(prs)
    
    print("[6/18] Slide 6: Kết quả 1 - Phân phối Tương đồng & BẢNG ĐỐI CHUẨN THỐNG KÊ...")
    build_slide_06_result1_similarity(prs)
    
    print("[7/18] Slide 7: Kết quả 2 - Cấu trúc 6 Nhóm SDG qua Heatmap (heatmap_6cat.png)...")
    build_slide_07_result2_heatmap(prs)
    
    print("[8/18] Slide 8: Đặc thù Ngành Sản xuất & Năng lượng (slide_manuf_sdgs.png)...")
    build_slide_08_result3_companies_p1(prs)
    
    print("[9/18] Slide 9: Đặc thù Ngành Tài chính & Bán lẻ (slide_finance_sdgs.png)...")
    build_slide_09_result3_companies_p2(prs)
    
    print("[10/18] Slide 10: Kết quả 4 - Xu hướng Dịch chuyển Chuỗi Thời gian (trends_6categories.png)...")
    build_slide_10_result4_trends(prs)
    
    print("[11/18] Slide 11: Kết quả 5 - Sắc thái Cảm xúc PhoBERT & Đối chuẩn (2 Ảnh)...")
    build_slide_11_result5_sentiment(prs)
    
    print("[12/18] Slide 12: Kết quả 6 - Tỷ số Cảm xúc Pos/Neg Ratio (sentiment_ratio.png)...")
    build_slide_12_result6_sentiment_ratio(prs)
    
    print("[13/18] Slide 13: Thảo luận - 'Nói nhiều về gì' (NHIỀU ẢNH: slide_company_top_sdgs.png)...")
    build_slide_13_discussion_talk_heavy(prs)
    
    print("[14/18] Slide 14: Thảo luận - 'Ít nói về gì' (4 Thẻ cảnh báo né tránh)...")
    build_slide_14_discussion_rarely_talk(prs)
    
    print("[15/18] Slide 15: BẢNG SO SÁNH ĐỐI ĐẦU TOÀN DIỆN PAPER GỐC VS CODE MỚI (Native Table)...")
    build_slide_15_comparison(prs)
    
    print("[16/18] Slide 16: Hàm ý Thực tiễn & Đề xuất Chính sách (3 Thẻ hành động)...")
    build_slide_16_implications(prs)
    
    print("[17/18] Slide 17: Hạn chế Đề tài & Tương lai (NHIỀU ẢNH: rag_agentic_flow.png)...")
    build_slide_17_limitations_future(prs)
    
    print("[18/18] Slide 18: Tổng kết 4 Thông điệp Cốt lõi & Phiên Hỏi đáp (Q&A)...")
    build_slide_18_conclusion(prs)
    
    prs.save(str(OUTPUT_PPTX))
    print(f"\n=> Đã lưu thành công bộ slide chính tại: {OUTPUT_PPTX}")
    
    PAPERS_DIR.mkdir(parents=True, exist_ok=True)
    dest_copy = PAPERS_DIR / "bao_cao_nghien_cuu_sdg_vietnam.pptx"
    shutil.copyfile(OUTPUT_PPTX, dest_copy)
    print(f"=> Đã sao chép vào: {dest_copy}")
    
    file_size_mb = OUTPUT_PPTX.stat().st_size / (1024 * 1024)
    print(f"=> Kích thước tệp: {file_size_mb:.2f} MB")
    print("=" * 80)
    print("HOÀN TẤT THÀNH CÔNG BỘ SLIDE CHÍNH!")
    print("=" * 80)


if __name__ == "__main__":
    main()
