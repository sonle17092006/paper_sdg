"""Mô-đun xây dựng toàn văn Bài báo tiếng Việt hoàn chỉnh chuẩn APA 7th.
Bao gồm:
- Tiêu đề, Tác giả, Abstract song ngữ
- 6 Phần nội dung đầy đủ (Giới thiệu, Tổng quan, Phương pháp, Kết quả, Thảo luận, Kết luận)
- 8 Hình ảnh (Figures)
- 6 Bảng số liệu (Tables)
- 6 Công thức toán học OMML (LaTeX to OMML)
- Hơn 40 tài liệu tham khảo APA 7th
"""

from __future__ import annotations

import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm, RGBColor

import paper_base as pb


def build_vietnamese_paper(doc: Document, tables_data: dict, is_endnote_ready: bool = False):
    pb.setup_clean_styles(doc)

    # Helper citation formatter
    def cite(author: str, year: str, paren: bool = True, alt_text: str = "") -> str:
        if is_endnote_ready:
            return f"{{{author}, {year}}}"
        if alt_text:
            return alt_text
        if paren:
            return f"({author}, {year})"
        return f"{author} ({year})"

    # =========================================================================
    # TIÊU ĐỀ BÀI BÁO VÀ THÔNG TIN TÁC GIẢ
    # =========================================================================
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before = Pt(14)
    tp.paragraph_format.space_after = Pt(6)
    run = tp.add_run(
        "XỬ LÝ NGÔN NGỮ TỰ NHIÊN ĐA NGỮ TRONG PHÂN TÍCH SDG VÀ CẢM XÚC "
        "BÁO CÁO PHÁT TRIỂN BỀN VỮNG CỦA DOANH NGHIỆP: "
        "BẰNG CHỨNG THỰC NGHIỆM TỪ CÁC DOANH NGHIỆP VIỆT NAM"
    )
    run.font.name = "Times New Roman"
    run.font.size = Pt(15)
    run.bold = True
    run.font.color.rgb = pb.COLOR_BLACK

    tep = doc.add_paragraph()
    tep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tep.paragraph_format.space_after = Pt(12)
    erun = tep.add_run(
        "Multilingual NLP for SDG and Sentiment Analysis of Corporate Sustainability Reports: "
        "Evidence from Vietnamese Enterprises"
    )
    erun.font.name = "Times New Roman"
    erun.font.size = Pt(12)
    erun.italic = True
    erun.font.color.rgb = pb.COLOR_BLACK

    ap = doc.add_paragraph()
    ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ap.paragraph_format.space_after = Pt(16)
    arun = ap.add_run("Lê Đan Sơn, Dương Thị Hoàn")
    arun.font.name = "Times New Roman"
    arun.font.size = Pt(12)
    arun.bold = True
    arun.font.color.rgb = pb.COLOR_BLACK

    # =========================================================================
    # TÓM TẮT & ABSTRACT
    # =========================================================================
    pb.add_h1(doc, "Tóm tắt")
    c_kang = cite("Kang", "2022", alt_text="Kang và Kim (2022)", paren=False)
    pb.add_p(
        doc,
        f"Báo cáo phát triển bền vững đóng vai trò kênh truyền thông huyết mạch giúp các doanh nghiệp niêm yết giải trình trách nhiệm "
        f"về môi trường, xã hội và quản trị (ESG) trước các bên liên quan và thu hút dòng vốn đầu tư có trách nhiệm. Kế thừa và mở rộng "
        f"khung phương pháp luận của {c_kang}, nghiên cứu này thực hiện phân tích định lượng tự động và trực quan hóa dữ liệu văn bản từ "
        f"42 báo cáo phát triển bền vững và báo cáo tích hợp của 7 tập đoàn niêm yết quy mô lớn hàng đầu tại Việt Nam (Bảo Việt, PAN Group, "
        f"Petrolimex, PNJ, SSI, Vicostone, Vinamilk) trong giai đoạn 6 năm liên tục (2020–2025) với tổng quy mô 96.461 câu văn bản số hóa (đã khôi phục hoàn chỉnh dữ liệu scan qua OCR). "
        f"Bằng việc sử dụng các mô hình học sâu biến đổi ngôn ngữ tiên tiến nhất (vietnamese-sbert và PhoBERT), nghiên cứu định lượng mức độ tương đồng "
        f"với 17 Mục tiêu Phát triển Bền vững (SDGs) của Liên Hợp Quốc, quy nạp vào 6 nhóm nhu cầu con người (Đời sống, Kinh tế, Công bằng, "
        f"Xã hội, Tài nguyên, Môi trường), phân tích cơ cấu cảm xúc đa lớp (Tích cực, Trung tính, Tiêu cực) và định lượng chỉ số tẩy xanh liên tục (GW Index) "
        f"thông qua đối chiếu với biến động giá cổ phiếu trên thị trường chứng khoán. Kết quả thực nghiệm chỉ ra rằng: "
        f"(1) Các doanh nghiệp Việt Nam thể hiện sự thiên lệch áp đảo vào nhóm Kinh tế (SDG 8, 9) và Đối tác xã hội (SDG 17), trong khi nhóm Công bằng (SDG 4, 5, 10) "
        f"liên tục nhận điểm số thấp nhất; (2) Tồn tại hiện tượng thiên lệch lạc quan (optimism bias) mang tính cấu trúc với 53,87% câu tích cực, "
        f"32,87% câu trung tính và chỉ 13,26% câu tiêu cực; (3) Kiểm định chuỗi thời gian cho thấy sự phân kỳ và tách rời hoàn toàn giữa giọng văn báo cáo "
        f"với biến động giá cổ phiếu thực tế (hệ số tương quan gộp r = -0,043, p = 0,806), phản ánh thị trường Việt Nam vẫn đang trong giai đoạn đầu định giá ESG; "
        f"(4) Chỉ số Greenwashing định lượng giúp nhận diện các giai đoạn thị trường suy giảm nhưng văn phong báo cáo lại đột ngột tô hồng, đồng thời cung cấp bài học "
        f"phương pháp luận sâu sắc về rủi ro diễn giải sai lệch do chất lượng số hóa tài liệu và hiệu quả giải cứu dữ liệu của công nghệ OCR. Nghiên cứu đề xuất các hàm ý chính sách then chốt "
        f"nhằm nâng cao tính minh bạch và chất lượng số hóa thông tin phi tài chính tại Việt Nam."
    )
    pb.add_p(
        doc,
        "Từ khóa: Báo cáo phát triển bền vững; Mục tiêu phát triển bền vững (SDGs); Xử lý ngôn ngữ tự nhiên (NLP); "
        "Sentence-BERT; PhoBERT; Phân tích cảm xúc; Tẩy xanh (Greenwashing); Nhận dạng ký tự quang học (OCR); Thị trường chứng khoán Việt Nam.",
        bold=True
    )

    pb.add_h1(doc, "Abstract")
    c_kang_en = cite("Kang", "2022", alt_text="Kang and Kim (2022)", paren=False)
    pb.add_p(
        doc,
        f"Corporate sustainability reports serve as pivotal communication conduits for disclosing environmental, social, and governance (ESG) "
        f"commitments to global investors and regulatory bodies. Adapting and extending the computational framework established by {c_kang_en}, "
        f"this study performs automated textual mining and longitudinal visualization of 42 standalone sustainability and integrated reports published "
        f"by seven prominent Vietnamese listed corporations (Bao Viet Holdings, The PAN Group, Petrolimex, PNJ, SSI Securities, Vicostone, Vinamilk) "
        f"spanning a balanced six-year horizon (2020–2025), encompassing 96,461 digitally extracted sentences (including full OCR text recovery for scanned reports). Employing state-of-the-art domain-specific "
        f"Transformer architectures (vietnamese-sbert and PhoBERT), we quantify enterprise alignment with the 17 UN Sustainable Development Goals (SDGs), "
        f"aggregate scores into six essential human-needs categories (Life, Economic, Equity, Social, Resources, Environments), evaluate contextual sentiment distributions "
        f"(Positive, Neutral, Negative), and formulate a continuous Greenwashing Index (GW) benchmarked against monthly equity market fluctuations. "
        f"Empirical findings demonstrate that: (1) Vietnamese corporate disclosures exhibit an overwhelming economic dominance (SDGs 8, 9) and partnership focus (SDG 17), "
        f"while systemic gaps persist in Equity disclosures (SDGs 4, 5, 10); (2) Disclosures are characterized by pronounced structural optimism bias "
        f"(53.87% positive, 32.87% neutral, and merely 13.26% negative sentences); (3) Longitudinal market tests reveal complete decoupling between narrative corporate tone "
        f"and annual stock performance (pooled Pearson r = -0.043, p = 0.806), evidencing impression management in an emerging market where non-financial data is not yet fully priced; "
        f"and (4) The continuous GW index effectively isolates market penalty periods juxtaposed against corporate tone inflation, while underscoring critical methodological lessons "
        f"regarding PDF digitization hygiene and OCR remediation of scanned reporting artifacts. Crucial policy recommendations are proposed to standardize digital sustainability reporting infrastructure in Vietnam."
    )
    pb.add_p(
        doc,
        "Keywords: Corporate sustainability reports; Sustainable Development Goals (SDGs); Natural language processing (NLP); "
        "Sentence-BERT; PhoBERT; Sentiment analysis; Greenwashing; Vietnam stock market.",
        bold=True
    )

    # =========================================================================
    # 1. GIỚI THIỆU (INTRODUCTION)
    # =========================================================================
    pb.add_h1(doc, "1. Giới thiệu (Introduction)")
    c_un = cite("United Nations", "2015")
    c_gri = cite("Global Reporting Initiative", "2021")
    c_btc = cite("Bộ Tài chính", "2020")
    c_pwc = cite("PwC Vietnam", "2022")
    c_ssc = cite("Ủy ban Chứng khoán Nhà nước", "2024")
    c_vbcsd = cite("Hội đồng Doanh nghiệp vì sự Phát triển Bền vững Việt Nam", "2024")

    pb.add_p(
        doc,
        f"Trong bối cảnh biến đổi khí hậu toàn cầu và sự chuyển dịch mạnh mẽ hướng tới mô hình kinh tế tuần hoàn, việc công bố thông tin "
        f"Môi trường, Xã hội và Quản trị (ESG) cùng các Mục tiêu Phát triển Bền vững {c_un} không còn là hoạt động mang tính thiện nguyện "
        f"đơn thuần mà đã trở thành yêu cầu pháp lý và tiêu chuẩn cạnh tranh sống còn đối với các doanh nghiệp niêm yết trên toàn cầu {c_gri}. "
        f"Tại Việt Nam, cam kết mạnh mẽ của Chính phủ tại Hội nghị COP26 về việc đạt mức phát thải ròng bằng 0 (Net-Zero) vào năm 2050 cùng sự ra đời "
        f"của Thông tư số 96/2020/TT-BTC {c_btc} đã chính thức bắt buộc các công ty đại chúng phải công bố báo cáo tác động môi trường và xã hội. "
        f"Thêm vào đó, việc phát triển Bộ Chỉ số Doanh nghiệp Bền vững (CSI) bởi Liên đoàn Thương mại và Công nghiệp Việt Nam {c_vbcsd} và Sổ tay thực hành ESG "
        f"do Ủy ban Chứng khoán Nhà nước ban hành {c_ssc} đã thúc đẩy các tập đoàn đầu ngành chuyển từ việc lồng ghép sơ lược trong báo cáo thường niên sang "
        f"phát hành các Báo cáo Phát triển Bền vững độc lập và Báo cáo Tích hợp toàn diện theo chuẩn mực quốc tế của Sáng kiến Báo cáo Toàn cầu (GRI)."
    )

    c_luccioni = cite("Luccioni et al.", "2020")
    c_arvidsson = cite("Arvidsson & Dumay", "2022")
    c_merkl = cite("Merkl-Davies & Brennan", "2007")
    c_lyon = cite("Lyon & Montgomery", "2015")
    pb.add_p(
        doc,
        f"Tuy nhiên, sự bùng nổ về số lượng và độ dài của các tài liệu báo cáo phi tài chính—thường lên tới hàng trăm trang với định dạng văn bản "
        f"tự do và đồ họa phức tạp—đang đặt ra thách thức chưa từng có đối với các cơ quan quản lý, chuyên viên phân tích tài chính và công chúng đầu tư {c_arvidsson}. "
        f"Phương pháp kiểm tra thủ công truyền thống (content analysis) bằng cách đọc và mã hóa thủ công bộc lộ những hạn chế cố hữu: tốn kém nguồn lực, "
        f"không thể mở rộng quy mô (unscalable), và dễ bị chi phối bởi định kiến chủ quan của người đánh giá. Nguy hiểm hơn, bản chất tự nguyện và thiếu cơ chế "
        f"kiểm toán độc lập nghiêm ngặt đối với thông tin định tính thường biến các báo cáo này thành công cụ phục vụ chiến lược quản trị ấn tượng "
        f"(impression management) {c_merkl}. Nhiều doanh nghiệp có xu hướng lựa chọn cherry-picking các thành tựu nổi bật, sử dụng ngôn từ mỹ miều, tô hồng "
        f"thực trạng hoạt động để che đậy các yếu kém về môi trường, dẫn tới hiện tượng tẩy xanh (greenwashing) gây méo mó thị trường và xói mòn niềm tin "
        f"của nhà đầu tư {c_lyon}."
    )

    c_reimers = cite("Reimers & Gurevych", "2019")
    c_nguyen = cite("Nguyen & Nguyen", "2020")
    pb.add_p(
        doc,
        f"Để giải quyết triệt để bài toán trên, việc ứng dụng Trí tuệ Nhân tạo (AI), đặc biệt là Xử lý Ngôn ngữ Tự nhiên (Natural Language Processing - NLP) "
        f"và mô hình học sâu biến đổi ngôn ngữ (Transformers), đã mở ra kỷ nguyên mới cho kiểm toán phi tài chính tự động {c_luccioni}. "
        f"Tiêu biểu cho hướng tiếp cận này là công trình nền tảng của {c_kang}, trong đó các tác giả đã sử dụng Sentence-BERT {c_reimers} và mô hình phân tích "
        f"cảm xúc để trực quan hóa mức độ gắn kết với 17 mục tiêu SDG trên 24 báo cáo của 6 tập đoàn tài phiệt lớn tại Hàn Quốc. Mặc dù công trình của Kang và Kim "
        f"đã đặt nền móng phương pháp luận xuất sắc, việc chuyển giao và ứng dụng mô hình này vào thị trường mới nổi như Việt Nam gặp phải những thách thức kỹ thuật "
        f"đặc thù: rào cản xử lý ngôn ngữ tiếng Việt (vốn có cấu trúc từ ghép đơn âm tiết, phụ thuộc chặt chẽ vào ngữ cảnh tách từ) và sự non trẻ của hệ sinh thái "
        f"công bố thông tin tài chính trong nước {c_nguyen}."
    )

    pb.add_p(
        doc,
        f"Xuất phát từ thực tiễn trên, nghiên cứu này được thực hiện nhằm tái hiện, chuẩn hóa và mở rộng toàn diện khung phân tích của {c_kang} "
        f"cho thị trường chứng khoán Việt Nam. Điểm đột phá của công trình này là việc mở rộng quy mô nghiên cứu lên 7 tập đoàn niêm yết lớn thuộc các ngành "
        f"kinh tế trụ cột (Bảo hiểm - Tài chính, Nông nghiệp - Thực phẩm, Năng lượng - Xăng dầu, Bán lẻ - Vàng bạc trang sức, Chứng khoán - Ngân hàng đầu tư, "
        f"Vật liệu xây dựng - Công nghiệp chế tạo, Hàng tiêu dùng nhanh - Sữa) trong chuỗi thời gian 6 năm liên tục từ 2020 đến 2025. Bằng việc phân tích định lượng "
        f"96.461 câu văn bản số hóa (kết hợp bóc tách PDF điện tử và phục hồi văn bản scan qua OCR), bài báo không chỉ trực quan hóa cơ cấu cam kết phát triển bền vững và sắc thái cảm xúc doanh nghiệp, mà còn thiết kế "
        f"chỉ số tẩy xanh liên tục (GW Index) đối chiếu với diễn biến thị trường chứng khoán thực tế, đồng thời rút ra bài học sâu sắc về chất lượng dữ liệu số hóa."
    )

    # =========================================================================
    # 2. TỔNG QUAN NGHIÊN CỨU (LITERATURE REVIEW)
    # =========================================================================
    pb.add_h1(doc, "2. Tổng quan nghiên cứu (Literature Review)")

    pb.add_h2(doc, "2.1 Các Mục tiêu Phát triển Bền vững (SDGs) và Chuẩn mực Báo cáo Phi Tài chính")
    c_pizzi = cite("Pizzi et al.", "2020")
    c_munoz = cite("Muñoz-Torres et al.", "2019")
    c_freeman = cite("Freeman", "1984")
    c_deegan = cite("Deegan", "2002")
    pb.add_p(
        doc,
        f"Năm 2015, Đại hội đồng Liên Hợp Quốc đã chính thức thông qua Chương trình Nghị sự 2030 với 17 Mục tiêu Phát triển Bền vững (SDGs) "
        f"và 169 mục tiêu cụ thể, thiết lập một khuôn khổ hành động toàn diện nhằm giải quyết các thách thức khẩn cấp về nghèo đói, bất bình đẳng, "
        f"suy thoái môi trường và biến đổi khí hậu {c_un}. Trong lý thuyết các bên liên quan (Stakeholder Theory) {c_freeman} và lý thuyết tính chính danh "
        f"(Legitimacy Theory) {c_deegan}, doanh nghiệp không chỉ có nghĩa vụ tối đa hóa lợi nhuận cho cổ đông mà còn phải chứng minh tính hợp pháp "
        f"và sự đóng góp tích cực vào phúc lợi xã hội thông qua các cam kết phát triển bền vững minh bạch."
    )
    pb.add_p(
        doc,
        f"Nhằm hỗ trợ doanh nghiệp cụ thể hóa các mục tiêu SDG vào chiến lược kinh doanh, Tổ chức Sáng kiến Báo cáo Toàn cầu đã phát triển Bộ tiêu chuẩn GRI "
        f"{c_gri}, cung cấp hệ thống chỉ số định lượng và định tính chuẩn hóa về kinh tế, môi trường và xã hội. Tuy nhiên, các nghiên cứu thực nghiệm quốc tế "
        f"chỉ ra rằng các doanh nghiệp thường có xu hướng 'chọn việc dễ, bỏ việc khó' (cherry-picking), tập trung công bố sâu vào các mục tiêu mang lại lợi ích "
        f"kinh tế trực tiếp mà phớt lờ các mục tiêu xã hội nhạy cảm hoặc đòi hỏi chi phí đầu tư lớn {c_munoz, c_pizzi}. Điều này đòi hỏi những công cụ đo lường "
        f"khách quan để giám sát toàn diện mức độ gắn kết thực chất của doanh nghiệp đối với toàn bộ 17 mục tiêu SDG."
    )

    pb.add_h2(doc, "2.2 Ứng dụng Xử lý Ngôn ngữ Tự nhiên (NLP) trong Đánh giá Báo cáo Doanh nghiệp")
    c_mikolov = cite("Mikolov et al.", "2013")
    c_devlin = cite("Devlin et al.", "2019")
    c_mercereau = cite("Mercereau & Melin", "2020")
    pb.add_p(
        doc,
        f"Sự tiến hóa của các phương pháp phân tích văn bản trong lĩnh vực tài chính - kế toán đã trải qua nhiều giai đoạn đột phá. Các phương pháp ban đầu "
        f"dựa trên tần suất từ vựng (Bag-of-Words, TF-IDF) hoàn toàn bỏ qua trật tự từ và cấu trúc ngữ pháp. Tiếp đó, các mô hình nhúng từ (word embeddings) "
        f"như Word2Vec {c_mikolov} và GloVe đã ghi nhận được mối quan hệ không gian giữa các từ nhưng gặp trở ngại nghiêm trọng trước hiện tượng đa nghĩa của từ "
        f"trong các ngữ cảnh chuyên ngành tài chính khác nhau."
    )
    pb.add_p(
        doc,
        f"Bước ngoặt xuất hiện với sự ra đời của kiến trúc Transformer và mô hình BERT {c_devlin}, cho phép máy tính hiểu sâu sắc ngữ cảnh hai chiều của câu văn. "
        f"Đặc biệt, mô hình Sentence-BERT (SBERT) của {c_reimers} đã khắc phục nhược điểm về tốc độ tính toán của BERT nguyên bản bằng cách sử dụng mạng Siamese "
        f"để tạo ra các vector nhúng câu có ý nghĩa ngữ nghĩa (semantically meaningful sentence embeddings), cho phép so sánh độ tương đồng cosine giữa các đoạn văn "
        f"với tốc độ vượt trội. Trong nghiên cứu về phát triển bền vững, {c_kang} đã chứng minh SBERT và USE vượt trội hoàn toàn so với TF-IDF trong việc nắm bắt "
        f"nội hàm mục tiêu SDG, mở đường cho việc tự động hóa đánh giá báo cáo doanh nghiệp ở quy mô lớn {c_mercereau}."
    )

    pb.add_h2(doc, "2.3 Phân tích Cảm xúc và Xu hướng Thiên lệch Lạc quan (Optimism Bias)")
    c_loughran = cite("Loughran & McDonald", "2011")
    c_huang = cite("Huang et al.", "2023")
    c_veenstra = cite("Veenstra & Ellemers", "2020")
    pb.add_p(
        doc,
        f"Phân tích cảm xúc (Sentiment Analysis) trong văn bản tài chính có lịch sử nghiên cứu phong phú bắt đầu từ từ điển chuẩn mực của {c_loughran}, "
        f"chỉ ra rằng các từ điển tâm lý thông thường thường phân loại sai các thuật ngữ tài chính tiêu chuẩn (như 'cost', 'liability', 'tax'). Gần đây, "
        f"sự xuất hiện của các mô hình chuyên biệt như FinBERT {c_huang} và PhoBERT {c_nguyen} đã nâng độ chính xác nhận diện sắc thái cảm xúc lên tầm cao mới. "
        f"Trong báo cáo phát triển bền vững, nhiều nghiên cứu quốc tế {c_veenstra} chỉ ra hiện tượng 'thiên lệch lạc quan' (optimism bias) mang tính hệ thống: "
        f"doanh nghiệp chủ động sử dụng ngôn từ tích cực áp đảo để xây dựng hình ảnh trách nhiệm xã hội, trong khi các sự cố môi trường, vi phạm lao động "
        f"hoặc chỉ tiêu phát thải không đạt thường bị làm mờ bằng cách dùng thể bị động hoặc câu từ trung tính né tránh trách nhiệm."
    )

    pb.add_h2(doc, "2.4 Hiện tượng Tẩy xanh (Greenwashing) và Sự Phân kỳ Thị trường")
    c_seele = cite("Seele & Gatti", "2017")
    c_heras = cite("Heras-Saizarbitoria et al.", "2022")
    c_spence = cite("Spence", "1973")
    c_connelly = cite("Connelly et al.", "2011")
    pb.add_p(
        doc,
        f"Theo định nghĩa kinh điển của {c_lyon}, tẩy xanh (greenwashing) là hành vi cố tình đánh lừa công chúng và người tiêu dùng về hiệu quả môi trường "
        f"của một công ty hoặc lợi ích môi trường của sản phẩm/dịch vụ. {c_seele} phân loại tẩy xanh thành hai cấp độ: tẩy xanh tuyên bố (claim greenwashing) "
        f"và tẩy xanh điều hành (executive greenwashing). Dưới góc độ Lý thuyết Tín hiệu (Signaling Theory) {c_spence, c_connelly}, báo cáo phát triển bền vững "
        f"đáng lẽ phải là tín hiệu tốn kém (costly signal) để các doanh nghiệp thực thi ESG xuất sắc phân biệt mình với các đối thủ yếu kém. "
        f"Tuy nhiên, khi chi phí phát hành báo cáo bằng lời nói rẻ hơn rất nhiều so với chi phí đầu tư công nghệ xanh (cheap talk), hiện tượng tẩy xanh sẽ bùng phát {c_heras}."
    )
    pb.add_p(
        doc,
        f"Để nhận diện hiện tượng này, một số nghiên cứu tiên phong đã đối chiếu sự biến động của giọng văn báo cáo với hiệu quả hoạt động thực tế trên thị trường "
        f"chứng khoán. Nếu giá cổ phiếu sụt giảm nghiêm trọng (phản ánh thị trường trừng phạt kết quả kinh doanh hoặc rủi ro pháp lý) nhưng giọng văn trong báo cáo "
        f"phát triển bền vững lại đột ngột gia tăng mức độ lạc quan và tích cực, đó là dấu hiệu cảnh báo mạnh mẽ về sự phân kỳ mang tính tẩy xanh nhằm xoa dịu áp lực cổ đông."
    )

    pb.add_h2(doc, "2.5 Khung Pháp lý và Thực trạng Công bố Thông tin Bền vững tại Việt Nam")
    c_tran = cite("Tran & Beddewela", "2020")
    c_hoang = cite("Hoang et al.", "2019")
    c_gerged = cite("Gerged et al.", "2021")
    pb.add_p(
        doc,
        f"Tại Việt Nam, tiến trình công bố thông tin phi tài chính đã có những bước tiến dài về mặt thể chế. Trước năm 2016, việc báo cáo ESG diễn ra phân tán "
        f"và chủ yếu mang hình thức các hoạt động từ thiện cộng đồng (CSR). Kể từ khi Bộ Tài chính ban hành Thông tư số 155/2015/TT-BTC và sau đó thay thế bằng "
        f"Thông tư số 96/2020/TT-BTC {c_btc}, các công ty niêm yết bắt buộc phải công bố thông tin về phát thải khí nhà kính, tiêu thụ năng lượng, chính sách lao động "
        f"và trách nhiệm cộng đồng trong Báo cáo Thường niên. Mặc dù khung khổ pháp lý ngày càng hoàn thiện, các nghiên cứu thực nghiệm của {c_tran, c_hoang} "
        f"chỉ ra rằng chất lượng báo cáo tại Việt Nam vẫn tồn tại khoảng cách rất lớn giữa nhóm các tập đoàn đầu ngành (VN30) và phần còn lại của thị trường. "
        f"Đa phần các doanh nghiệp vừa và nhỏ chỉ dừng lại ở mức tuân thủ hình thức (tick-box compliance), trong khi việc lập báo cáo chuyên biệt và có bảo đảm "
        f"độc lập vẫn là điều hiếm hoi trên thị trường chứng khoán Việt Nam {c_pwc, c_gerged}."
    )

    # =========================================================================
    # 3. PHƯƠNG PHÁP NGHIÊN CỨU (METHODOLOGY)
    # =========================================================================
    pb.add_h1(doc, "3. Phương pháp nghiên cứu (Methodology)")

    pb.add_h2(doc, "3.1 Dữ liệu và Mẫu Nghiên cứu (Sample Selection)")
    pb.add_p(
        doc,
        f"Để đảm bảo tính đại diện cao cho nền kinh tế Việt Nam và xây dựng một bộ dữ liệu bảng cân bằng (balanced panel data) phục vụ phân tích chuỗi thời gian, "
        f"nghiên cứu tiến hành thu thập toàn bộ các Báo cáo Phát triển Bền vững độc lập và Báo cáo Tích hợp công bố công khai của 7 doanh nghiệp niêm yết lớn "
        f"trong giai đoạn 6 năm liên tục từ 2020 đến 2025 (tổng cộng 42 báo cáo). 7 doanh nghiệp này đại diện cho các ngành kinh tế xương sống: "
        f"Bảo hiểm - Tài chính (BVH), Nông nghiệp - Thực phẩm chế biến (PAN), Năng lượng - Xăng dầu (PLX), Bán lẻ kim hoàn (PNJ), Dịch vụ tài chính - Chứng khoán (SSI), "
        f"Vật liệu xây dựng chế tạo (VCS), và Hàng tiêu dùng nhanh - Sữa (VNM). Thông tin khái quát về các doanh nghiệp được trình bày tại Bảng 1."
    )

    # Table 1: Sample Overview
    h1, d1 = tables_data["t1_vi"]
    pb.add_table_clean(
        doc,
        table_label="Bảng 1",
        table_title="Khái quát mẫu nghiên cứu 7 doanh nghiệp niêm yết giai đoạn 2020–2025",
        headers=h1,
        data=d1,
        note="Dữ liệu tổng hợp từ các Báo cáo Phát triển Bền vững và Báo cáo Tích hợp công bố chính thức trên cổng thông tin HOSE, HNX và website doanh nghiệp.",
        col_widths=[1.5, 3.8, 3.8, 1.2, 1.8, 3.9],
        alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
        ],
        font_size=8.5,
    )

    pb.add_h2(doc, "3.2 Quy trình Trích xuất và Tiền Xử lý Dữ liệu Văn bản (Text Preprocessing & OCR Pipeline)")
    pb.add_p(
        doc,
        f"Các báo cáo phát triển bền vững thường được phát hành dưới định dạng PDF với bố cục đồ họa đa cột phức tạp, chèn xen kẽ hình ảnh, biểu đồ và bảng biểu. "
        f"Đặc biệt, trong bối cảnh thực tiễn tại Việt Nam, một số báo cáo phát hành dưới dạng tệp scan ảnh thuần túy (raster image scans) không chứa lớp văn bản điện tử (text layer). "
        f"Để đảm bảo tính toàn vẹn và khách quan của dữ liệu, quy trình trích xuất văn bản được thiết kế theo kiến trúc hai nhánh chuyên biệt: "
        f"(1) Đối với các tệp PDF điện tử tiêu chuẩn: Sử dụng thư viện PyMuPDF (fitz) bóc tách cấu trúc từng trang, tự động loại bỏ trang bìa, mục lục tra cứu, lời tựa thủ tục, "
        f"tách câu theo quy chuẩn kết thúc câu tiếng Việt và áp dụng bộ lọc độ dài loại bỏ các chuỗi ký tự gãy khúc hoặc dưới 6 từ; "
        f"(2) Đối với các tệp scan dạng ảnh (điển hình là Báo cáo PTBV năm 2022 của PNJ với 52 trang scan hình ảnh): Xây dựng quy trình nhận dạng ký tự quang học (OCR) "
        f"chuyên dụng sử dụng công cụ Tesseract OCR 5.4.0 với bộ dữ liệu huấn luyện song ngữ tiếng Việt - tiếng Anh ('vie+eng'), độ phân giải trích xuất 300 DPI kèm tiền xử lý nhị phân hóa (binarization) "
        f"và khử nhiễu đồ họa, giúp khôi phục hoàn chỉnh 669 câu văn bản thực chất thay vì bị bỏ sót dữ liệu. "
        f"Sau khi tiền xử lý và hoàn tất khâu OCR, toàn bộ 42 báo cáo đã kết xuất được 96.461 câu văn bản hợp lệ với tổng số 4.997 trang tài liệu, đạt mật độ bình quân 20,35 câu/trang. "
        f"Chi tiết thống kê số câu và số trang từng báo cáo được ghi nhận tại Bảng 2."
    )

    # Table 2: Text Digitization Stats
    h2, d2 = tables_data["t2_vi"]
    pb.add_table_clean(
        doc,
        table_label="Bảng 2",
        table_title="Thống kê mô tả quy mô số hóa dữ liệu văn bản từ 42 báo cáo (2020–2025)",
        headers=h2,
        data=d2,
        note="Tổng quy mô: 42 báo cáo, 4.997 trang tài liệu, 96.461 câu văn bản hợp lệ (đã khôi phục 669 câu OCR cho PNJ 2022). Mật độ bình quân đạt 20,35 câu/trang.",
        col_widths=[2.0, 1.5, 4.5, 2.5, 2.5, 3.0],
        alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.0,
    )

    pb.add_h2(doc, "3.3 Đo lường Độ Tương đồng Ngữ nghĩa SDG bằng Sentence-BERT")
    pb.add_p(
        doc,
        f"Kế thừa phương pháp luận của {c_kang}, nghiên cứu xây dựng tập văn bản ngữ liệu mục tiêu chuẩn (SDG reference corpus) gồm 169 mục tiêu cụ thể "
        f"thuộc 17 Mục tiêu Phát triển Bền vững của Liên Hợp Quốc được chuyển ngữ chính xác sang tiếng Việt và bổ sung các thuật ngữ phát triển bền vững "
        f"chuẩn mực doanh nghiệp (~400 câu mẫu đại diện). Ký hiệu S_g là tập hợp các câu chuẩn thuộc mục tiêu SDG thứ g (với g từ 1 đến 17)."
    )
    pb.add_p(
        doc,
        f"Để mã hóa ngữ nghĩa, nghiên cứu sử dụng mô hình học sâu biến đổi vietnamese-sbert {c_reimers}, được huấn luyện theo kiến trúc mạng Siamese "
        f"trên ngữ liệu tiếng Việt quy mô lớn. Mỗi câu báo cáo r và mỗi câu chuẩn s trong tập mục tiêu S_g được ánh xạ thành các vector nhúng d-chiều "
        f"r_hat và s_hat (với d = 768). Điểm tương đồng giữa câu báo cáo r và mục tiêu SDG g là giá trị trung bình độ tương đồng cosine giữa r_hat "
        f"với tất cả các câu chuẩn s_hat thuộc mục tiêu đó:"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{sim}(r, g) = \frac{1}{|S_g|} \sum_{s \in S_g} \cos(\hat{\mathbf{r}}, \hat{\mathbf{s}}) = \frac{1}{|S_g|} \sum_{s \in S_g} \frac{\hat{\mathbf{r}} \cdot \hat{\mathbf{s}}}{\|\hat{\mathbf{r}}\| \|\hat{\mathbf{s}}\|}",
        eq_num="1"
    )

    pb.add_h2(doc, "3.4 Phân tích Cảm xúc Ngôn ngữ Ngữ cảnh bằng PhoBERT")
    pb.add_p(
        doc,
        f"Trong bài báo gốc của {c_kang}, các tác giả áp dụng mô hình phân loại cảm xúc nhị phân DistilBERT trên tiếng Anh với thang điểm phân cực từ 0 đến 1. "
        f"Nhằm tối ưu hóa cho ngữ cảnh tiếng Việt và nắm bắt bản chất của văn phong hành chính doanh nghiệp (vốn chứa nhiều câu báo cáo thông số trung tính), "
        f"nghiên cứu này nâng cấp lên mô hình PhoBERT-base {c_nguyen} được tinh chỉnh chuyên sâu cho bài toán phân tích cảm xúc tiếng Việt (phobert-base-vietnamese-sentiment). "
        f"Mô hình phân loại mỗi câu văn r vào một trong ba lớp xác suất: Tích cực (Positive), Trung tính (Neutral) và Tiêu cực (Negative). "
        f"Để định lượng sắc thái lạc quan tổng thể của từng báo cáo theo năm t, tỷ số cảm xúc (Sentiment Ratio) được tính bằng tỷ lệ giữa số câu tích cực "
        f"và số câu tiêu cực:"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{Ratio}_t = \frac{N_{\mathrm{pos}, t}}{N_{\mathrm{neg}, t}}",
        eq_num="2"
    )

    pb.add_p(
        doc,
        "Các câu mang sắc thái Trung tính (Neutral) phản ánh thông tin số liệu kỹ thuật thuần túy nên không đưa vào tử số hoặc mẫu số nhằm đảm bảo "
        "tính nhạy cảm và độ tập trung của tỷ số đối với sự đối lập giữa thành tựu và rủi ro/khó khăn."
    )

    pb.add_h2(doc, "3.5 Chuẩn hóa Min-Max và Gom nhóm 6 Danh mục Nhu cầu Con người")
    pb.add_p(
        doc,
        f"Do giá trị độ tương đồng cosine ban đầu thường nằm trong khoảng hẹp (từ -0,071 đến 0,531), nghiên cứu tuân thủ chặt chẽ quy trình của {c_kang} "
        f"bằng cách áp dụng phép chuẩn hóa tuyến tính Min-Max toàn cục trên toàn bộ 17 cột mục tiêu để đưa điểm số về thang đo trực quan từ 0 đến 100:"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{Score}(r, g) = \frac{\mathrm{sim}(r, g) - \min_{r', g'}(\mathrm{sim})}{\max_{r', g'}(\mathrm{sim}) - \min_{r', g'}(\mathrm{sim})} \times 100",
        eq_num="3"
    )

    pb.add_p(
        doc,
        f"Tiếp theo, kế thừa cấu trúc phân loại của {c_kang}, 17 mục tiêu SDG được tổng hợp vào 6 nhóm danh mục nhu cầu thiết yếu của con người: "
        f"(1) Đời sống (Life: SDG 1, 2, 3); (2) Kinh tế (Economic: SDG 8, 9); (3) Công bằng (Equity: SDG 4, 5, 10); (4) Xã hội (Social: SDG 11, 16, 17); "
        f"(5) Tài nguyên (Resources: SDG 6, 7, 12, 14); và (6) Môi trường (Environments: SDG 13, 15). Điểm số của nhóm danh mục k cho câu văn r "
        f"được xác định bằng trung bình cộng các SDG thành phần:"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{Cat}_k(r) = \frac{1}{|G_k|} \sum_{g \in G_k} \mathrm{Score}(r, g)",
        eq_num="4"
    )

    pb.add_h2(doc, "3.6 Phương pháp Đối chiếu Thị trường và Định lượng Chỉ số Tẩy xanh (Greenwashing Index)")
    pb.add_p(
        doc,
        f"Một đóng góp phương pháp luận vượt trội của nghiên cứu này so với công trình của {c_kang} là việc thiết kế mô hình đối chiếu động giữa diễn biến "
        f"giá cổ phiếu thực tế và giọng văn báo cáo phát triển bền vững. Nghiên cứu thu thập dữ liệu giá đóng cửa hàng tháng của 7 cổ phiếu niêm yết trên HOSE và HNX, "
        f"đồng thời tính toán mức tăng giảm giá thường niên (\\Delta\\text{{giá}}\\%) và mức biến động của tỷ số cảm xúc (\\Delta\\text{{Pos/Neg}}\\%)."
    )
    pb.add_p(
        doc,
        f"Khi thị trường chứng khoán suy giảm (\\Delta\\text{{giá}}\\% < -10\\%), hành vi phản ứng của doanh nghiệp được phân loại như sau: "
        f"(1) Tẩy xanh (Greenwashing): Tỷ số cảm xúc tăng mạnh (\\Delta\\text{{Pos/Neg}}\\% > 15\\%), phản ánh việc doanh nghiệp cố tình dùng ngôn từ tô hồng "
        f"để xoa dịu nhà đầu tư khi kết quả kinh doanh kém cỏi; (2) Trung thực (Honest): Tỷ số cảm xúc cũng sụt giảm đồng thuận theo thực tế thị trường; "
        f"và (3) Bình thường (None): Không có dấu hiệu đột biến rõ rệt. Nhằm định lượng hóa mức độ tẩy xanh một cách liên tục thay vì chỉ dùng nhãn nhị phân, "
        f"nghiên cứu đề xuất Chỉ số Tẩy xanh (Greenwashing Score - GW) chỉ kích hoạt khi giá cổ phiếu suy giảm:"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{GW} = \frac{\max(0, -\Delta \text{giá}\%) \times \max(0, \Delta \text{Pos/Neg}\%)}{100}",
        eq_num="5"
    )

    pb.add_p(
        doc,
        "Chỉ số GW được phân thành 4 bậc thang mức độ: GW = 0 (Không tẩy xanh / Trung thực); 0 < GW <= 2 (Nhẹ); 2 < GW <= 8 (Trung bình); và GW > 8 (Mạnh). "
        "Đồng thời, để kiểm định mối liên hệ tổng thể giữa thị trường và ngôn ngữ báo cáo, hệ số tương quan tuyến tính Pearson (r) và tương quan thứ hạng Spearman (\\rho) "
        "được ước lượng trên mẫu gộp và từng doanh nghiệp:"
    )

    pb.add_equation_clean(
        doc,
        r"r = \frac{\sum_{i=1}^n (x_i - \bar{x})(y_i - \bar{y})}{\sqrt{\sum_{i=1}^n (x_i - \bar{x})^2 \sum_{i=1}^n (y_i - \bar{y})^2}}",
        eq_num="6"
    )

    pb.add_h2(doc, "3.7 Điểm Kế thừa và Các Cải tiến Vượt trội so với Kang & Kim (2022)")
    pb.add_p(
        doc,
        f"Để độc giả theo dõi thuận tiện mạch phát triển phương pháp luận, nghiên cứu làm rõ các điểm tương đồng và khác biệt then chốt so với {c_kang}: "
        f"(1) Về ngôn ngữ và mô hình nhúng: Kang & Kim xử lý báo cáo tiếng Anh bằng USE và GloVe; nghiên cứu này phát triển đường ống chuyên biệt cho tiếng Việt "
        f"với vietnamese-sbert và PhoBERT; (2) Về quy mô và thời gian: Nghiên cứu gốc khảo sát 6 công ty Hàn Quốc trong 4 năm (2017–2020, 24 báo cáo); nghiên cứu này "
        f"khảo sát 7 tập đoàn Việt Nam trong 6 năm liên tục (2020–2025, 42 báo cáo, 96.461 câu); (3) Về cơ chế cảm xúc: Kang & Kim chỉ phân loại nhị phân 0–1; "
        f"nghiên cứu này phân loại 3 lớp (Positive, Neutral, Negative) nhằm bảo toàn tính trung thực của các câu văn báo cáo số liệu kỹ thuật; "
        f"(4) Về mở rộng thị trường: Nghiên cứu này bổ sung hoàn toàn mô hình chuỗi thời gian giá cổ phiếu hàng tháng và chỉ số định lượng tẩy xanh liên tục (GW Index), "
        f"một nội dung không có trong bài báo gốc; (5) Về tiền xử lý dữ liệu thực chiến: Tích hợp module OCR phục hồi dữ liệu scan chuyên sâu, ngăn chặn hiện tượng méo mó dữ liệu."
    )

    # =========================================================================
    # 4. KẾT QUẢ NGHIÊN CỨU THỰC NGHIỆM (EMPIRICAL RESULTS)
    # =========================================================================
    pb.add_h1(doc, "4. Kết quả nghiên cứu thực nghiệm (Empirical Results)")

    pb.add_h2(doc, "4.1 Phân phối Điểm Tương đồng SDG Toàn cục")
    pb.add_figure_clean(
        doc,
        "similarity_hist.png",
        "Hình 1",
        "Phân phối tần suất điểm tương đồng mục tiêu phát triển bền vững (SDGs) của 96.461 câu sau chuẩn hóa Min-Max 0–100",
        width_inches=5.8
    )
    pb.add_p(
        doc,
        "Hình 1 minh họa phân phối tần suất của điểm tương đồng SDG trên toàn bộ 96.461 câu văn bản số hóa sau khi được chuẩn hóa Min-Max toàn cục về thang điểm 0–100. "
        "Đồ thị thể hiện hình dạng phân phối chuẩn dạng chuông đối xứng tuyệt đẹp với điểm trung bình đạt 45,43 điểm và độ lệch chuẩn 11,87 điểm. "
        "Khoảng điểm tập trung dày đặc nhất nằm trong khoảng từ 35 đến 55 điểm, phản ánh bản chất của ngôn ngữ báo cáo phát triển bền vững: phần lớn các câu văn "
        "trong báo cáo mang tính giới thiệu quy trình, cấu trúc tổ chức hoặc mô tả hoạt động chung với mức độ gắn kết vừa phải đối với các mục tiêu SDG cụ thể."
    )
    pb.add_p(
        doc,
        "Đáng chú ý, phần đuôi bên phải của đồ thị (với điểm số từ 65 đến 90 điểm) chiếm tỷ lệ khoảng 8,5% tổng số câu văn. Đây chính là các câu mang hàm lượng thông tin "
        "chuyên biệt sâu sắc, mô tả trực diện các sáng kiến hành động cụ thể của doanh nghiệp gắn kết chặt chẽ với các chỉ tiêu con của Liên Hợp Quốc (như đầu tư "
        "công nghệ giảm phát thải CO2, chương trình an toàn sinh học vùng nuôi tôm, hay hệ thống xử lý nước thải tuần hoàn). Kết quả phân phối này hoàn toàn tương thích "
        f"với phát hiện của {c_kang} trên thị trường Hàn Quốc, khẳng định tính ổn định và độ tin cậy của thuật toán SBERT khi ánh xạ văn bản tiếng Việt sang không gian SDG."
    )

    pb.add_h2(doc, "4.2 Cấu trúc Cam kết 6 Nhóm SDG và Ma trận Nhiệt (Heatmap)")
    pb.add_figure_clean(
        doc,
        "heatmap_6cat.png",
        "Hình 2",
        "Bản đồ nhiệt (Heatmap) mức độ cam kết 6 nhóm danh mục SDG của 7 doanh nghiệp qua 42 báo cáo giai đoạn 2020–2025",
        width_inches=5.6
    )
    pb.add_p(
        doc,
        "Hình 2 trình bày bản đồ nhiệt mức độ gắn kết với 6 nhóm danh mục nhu cầu con người của 7 doanh nghiệp qua 42 báo cáo từ năm 2020 đến 2025. "
        "Các gam màu chuyển dần từ vàng nhạt (mức độ gắn kết thấp, ~38 điểm) sang đỏ sẫm (mức độ gắn kết cao, >51 điểm). Chi tiết các giá trị định lượng trung bình "
        "từng năm của 7 doanh nghiệp được trình bày chi tiết tại Bảng 3. Quan sát bản đồ nhiệt cho thấy một trật tự ưu tiên mang tính cấu trúc vững chắc: "
        "Nhóm Kinh tế (Economic) luôn áp đảo với sắc đỏ đậm nhất trên toàn bộ các công ty và năm khảo sát, tiếp theo là nhóm Xã hội (Social) và Tài nguyên (Resources), "
        "trong khi nhóm Công bằng (Equity) liên tục hiển thị màu vàng nhạt thấp nhất."
    )

    # Table 3: 6 Category Means
    h3, d3 = tables_data["t3_vi"]
    pb.add_table_clean(
        doc,
        table_label="Bảng 3",
        table_title="Điểm trung bình 6 nhóm danh mục SDG theo doanh nghiệp và năm (thang điểm 0–100)",
        headers=h3,
        data=d3,
        note="Life: SDG 1, 2, 3; Economic: SDG 8, 9; Equity: SDG 4, 5, 10; Social: SDG 11, 16, 17; Resources: SDG 6, 7, 12, 14; Environments: SDG 13, 15.",
        col_widths=[1.5, 1.2, 1.8, 1.8, 1.8, 1.8, 1.8, 1.8],
        alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.0,
    )

    pb.add_p(
        doc,
        "Dữ liệu Bảng 3 chỉ ra rằng nhóm Kinh tế luôn giữ điểm số cao nhất trong mọi báo cáo, dao động từ 43,19 điểm (VNM 2022) đến đỉnh cao 51,90 điểm (VCS 2025) "
        "và 51,04 điểm (BVH 2025). Điều này phản ánh rõ nét định hướng thực dụng của các doanh nghiệp Việt Nam: tăng trưởng kinh tế, việc làm bền vững (SDG 8) "
        "và đổi mới sáng tạo, phát triển hạ tầng (SDG 9) là những trụ cột mang tính sống còn gắn liền với lợi ích cổ đông. Ngược lại, nhóm Công bằng (Equity: giáo dục chất lượng, "
        "bình đẳng giới, giảm bất bình đẳng) liên tục ghi nhận mức điểm thấp kỷ lục (chỉ đạt từ 38,36 điểm ở SSI 2025 đến 44,94 điểm ở VCS 2025). "
        "Khoảng cách chênh lệch giữa nhóm cao nhất (Kinh tế) và nhóm thấp nhất (Công bằng) luôn duy trì ổn định ở mức 6 đến 8 điểm trên toàn bộ các công ty, "
        "bộc lộ một khoảng trống cam kết xã hội mang tính hệ thống cần có sự can thiệp điều chỉnh từ chính sách."
    )

    pb.add_h2(doc, "4.3 Xu hướng Biến động Cam kết Phát triển Bền vững 2020–2025")
    pb.add_figure_clean(
        doc,
        "trends_6categories.png",
        "Hình 3",
        "Xu hướng biến động điểm 6 nhóm danh mục SDG theo thời gian (2020–2025) của 7 doanh nghiệp niêm yết",
        width_inches=5.8
    )
    pb.add_p(
        doc,
        "Hình 3 minh họa quỹ đạo biến động điểm số của 6 nhóm danh mục SDG qua 6 năm cho từng doanh nghiệp. Điểm nhấn nổi bật nhất là xu hướng tăng trưởng đồng loạt "
        "của hầu hết các doanh nghiệp trong giai đoạn 2023–2025. Đặc biệt, Tập đoàn Bảo Việt (BVH), Vinamilk (VNM) và Vicostone (VCS) cho thấy sự bứt phá mạnh mẽ "
        "ở tất cả các nhóm chỉ tiêu trong hai năm 2024 và 2025. Tại Vinamilk, nhóm Môi trường (Environments) đã tăng vọt từ mức 41,23 điểm năm 2020 lên 47,29 điểm năm 2025, "
        "nhờ vào việc doanh nghiệp này chính thức công bố lộ trình trung hòa carbon Net Zero 2050 theo tiêu chuẩn PAS 2060 và triển khai hàng loạt trang trại sinh thái chuẩn quốc tế."
    )
    pb.add_p(
        doc,
        "Tương tự, Vicostone (VCS) ghi nhận mức mở rộng quy mô báo cáo ngoạn mục vào năm 2025 với 5.974 câu văn bản (gấp hơn 2,5 lần các năm trước), đưa điểm nhóm Kinh tế "
        "lên 51,90 điểm và nhóm Xã hội lên 50,51 điểm, phản ánh sự chuyển dịch từ báo cáo tuân thủ sang tích hợp chiến lược ESG sâu rộng trong chuỗi cung ứng đá nhân tạo. "
        "Đối với các doanh nghiệp tài chính như SSI và BVH, nhóm Xã hội (đối tác vì các mục tiêu và quản trị minh bạch) luôn duy trì vị trí thứ hai sau kinh tế, "
        "chứng minh vai trò dẫn dắt dòng vốn xanh và đầu tư có trách nhiệm."
    )

    pb.add_h2(doc, "4.4 Phân tích Cảm xúc Ngôn ngữ Báo cáo")
    pb.add_figure_clean(
        doc,
        "sentiment_hist.png",
        "Hình 4",
        "Phân phối xác suất độ phân cực cảm xúc của các câu văn bản (mô hình PhoBERT)",
        width_inches=5.2
    )
    pb.add_p(
        doc,
        "Hình 4 thể hiện phân phối tần suất của điểm xác suất phân cực cảm xúc dự báo bởi PhoBERT. Đồ thị mang đặc trưng phân phối hai đỉnh (bimodal distribution) "
        "rõ rệt: một đỉnh tập trung tại vùng trung tính (xung quanh giá trị 0,45 - 0,55) và một đỉnh cực đại nhô cao tại vùng giá trị tích cực tuyệt đối (0,90 - 1,00). "
        "Số lượng câu văn rơi vào vùng tiêu cực (< 0,20) là cực kỳ khiêm tốn. Điều này trực quan hóa bằng chứng sống động về sự áp đảo của các thông điệp truyền thông tích cực."
    )

    pb.add_figure_clean(
        doc,
        "sentiment_by_company.png",
        "Hình 5",
        "Cơ cấu tỷ lệ các câu Tích cực / Trung tính / Tiêu cực và Tỷ số Pos/Neg theo doanh nghiệp qua các năm",
        width_inches=5.8
    )
    pb.add_figure_clean(
        doc,
        "sentiment_ratio.png",
        "Hình 6",
        "Diễn biến Tỷ số Cảm xúc (Pos/Neg Ratio) hàng năm giữa 7 doanh nghiệp giai đoạn 2020–2025",
        width_inches=5.6
    )

    # Table 4: Sentiment Counts
    h4, d4 = tables_data["t4_vi"]
    pb.add_table_clean(
        doc,
        table_label="Bảng 4",
        table_title="Thống kê cơ cấu và tỷ lệ cảm xúc văn bản trong 42 báo cáo phát triển bền vững",
        headers=h4,
        data=d4,
        note="Phân loại cảm xúc tự động bằng mô hình PhoBERT. Tỷ số Pos/Neg = Số câu Tích cực / Số câu Tiêu cực.",
        col_widths=[1.8, 1.2, 2.2, 2.2, 2.2, 2.2],
        alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.0,
    )

    pb.add_p(
        doc,
        "Hình 5 và Bảng 4 tổng hợp bức tranh cơ cấu cảm xúc trên toàn bộ 96.461 câu văn bản. Kết quả thống kê chỉ ra rằng: số câu mang sắc thái Tích cực (Positive) "
        "chiếm tỷ lệ đa số tuyệt đối với 51.966 câu (tương đương 53,87%), câu Trung tính (Neutral) chiếm 31.706 câu (32,87%), trong khi câu mang sắc thái Tiêu cực (Negative) "
        "chỉ chiếm 12.789 câu (13,26%). Nhìn chung, cứ mỗi câu văn đề cập đến rủi ro hoặc khó khăn, các doanh nghiệp Việt Nam sử dụng tới hơn 4 câu văn "
        "để ca ngợi thành tựu và kết quả tích cực (tỷ số Pos/Neg bình quân toàn mẫu đạt 4,06 lần)."
    )
    pb.add_p(
        doc,
        "Hình 6 khắc họa sự phân hóa sâu sắc về tỷ số Pos/Neg giữa các doanh nghiệp. Vinamilk (VNM) duy trì phong độ ổn định ở mức cao với tỷ số bình quân đạt 5,38 lần (dao động từ 4,35 đến 6,06), "
        "thể hiện chiến lược truyền thông thương hiệu bền vững bậc thầy gắn liền với lộ trình Net Zero. Ngược lại, Bảo Việt (BVH) có tỷ lệ câu trung tính rất cao (trên 40%) trong các năm 2020–2022 "
        "do đặc thù báo cáo tích hợp phải giải trình nhiều bảng cân đối tài chính và quản trị rủi ro kỹ thuật, trước khi bật tăng mạnh trên 6,0 lần trong giai đoạn 2023–2025. "
        "Đặc biệt, đối với PNJ, việc áp dụng công nghệ OCR trên 52 trang scan năm 2022 đã khôi phục thành công 669 câu văn bản, xác định tỷ số Pos/Neg thực tế đạt 1,21 lần "
        "(317 câu tích cực và 261 câu tiêu cực). Điểm trũng này phản ánh khách quan các khó khăn thực tế mà PNJ phải đương đầu sau giai đoạn phong tỏa nghiêm ngặt do đại dịch COVID-19 "
        "cuối năm 2021 tại TP.HCM (khiến hơn 80% cửa hàng phải tạm đóng cửa) cùng áp lực biến động giá vàng nguyên liệu, thay vì con số 0,33 lần giả tạo khi chỉ bóc tách được 12 câu pháp lý như trước khi OCR. "
        "Bước sang năm 2023, tỷ số Pos/Neg của PNJ phục hồi về mức 4,23 lần khi doanh nghiệp số hóa hoàn chỉnh báo cáo."
    )

    pb.add_h2(doc, "4.5 Đối chiếu Giá Cổ phiếu Hàng tháng và Tỷ lệ Cảm xúc Thường niên")
    pb.add_figure_clean(
        doc,
        "stock_vs_sentiment_greenwashing.png",
        "Hình 7",
        "Đối chiếu chuỗi giá cổ phiếu hàng tháng (đường xanh) và Pos/Neg ziczac đỉnh năm (nền núi cam) quy đổi về index 100",
        width_inches=6.0
    )

    # Table 5: Correlation
    h5, d5 = tables_data["t5_vi"]
    pb.add_table_clean(
        doc,
        table_label="Bảng 5",
        table_title="Kết quả kiểm định tương quan Pearson và Spearman giữa biến động giá cổ phiếu và tỷ lệ cảm xúc",
        headers=h5,
        data=d5,
        note="Kiểm định trên chuỗi biến động hàng năm (% đổi giá cuối năm vs % đổi tỷ số Pos/Neg). Cỡ mẫu n = 5 cặp quan sát/công ty; Mẫu gộp N = 35 quan sát (giai đoạn 2020–2025).",
        col_widths=[3.5, 1.5, 1.8, 1.8, 1.8, 1.8, 3.8],
        alignments=[
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.LEFT,
        ],
        font_size=8.0,
    )

    pb.add_p(
        doc,
        "Hình 7 mô tả trực quan mối tương quan động giữa chuỗi giá cổ phiếu hàng tháng (đường màu xanh nước biển) và diễn biến tỷ lệ Pos/Neg thường niên (đỉnh ziczac "
        "kèm diện tích bóng đổ màu cam), với cả hai chuỗi số liệu được quy đổi về cùng một gốc chỉ số index 100 tại tháng đầu tiên khảo sát (tháng 1/2020). "
        "Các dải màu nền đỏ nhạt thể hiện các giai đoạn cảnh báo tẩy xanh (giá cổ phiếu năm tụt dốc > 10% nhưng tỷ lệ Pos/Neg lại tăng vọt > 15%), "
        "trong khi các dải màu xanh lá cây thể hiện thái độ trung thực (giá giảm và giọng văn báo cáo cũng giảm mức độ lạc quan tương ứng)."
    )
    pb.add_p(
        doc,
        "Bảng 5 trình bày kết quả ước lượng thống kê tương quan Pearson và Spearman. Trên toàn bộ mẫu gộp (Pooled Sample, N = 35 cặp quan sát), "
        "hệ số tương quan tuyến tính Pearson đạt r = -0,043 (p = 0,806) và tương quan thứ hạng Spearman đạt rho = -0,098 (p = 0,575). Với giá trị p-value "
        "lớn hơn rất nhiều so với ngưỡng ý nghĩa thống kê 0,10, nghiên cứu kết luận rằng biến động thị trường cổ phiếu và sắc thái cảm xúc trong báo cáo PTBV "
        "tại Việt Nam hoàn toàn độc lập và tách rời nhau. Khi xét riêng từng công ty với cỡ mẫu nhỏ (n = 5), các hệ số tương quan biến thiên từ âm sang dương "
        "nhưng đều không đạt độ tin cậy thống kê (p >= 0,120), khẳng định giọng văn báo cáo không bị chi phối trực tiếp bởi các cú sốc định giá ngắn hạn của thị trường chứng khoán."
    )

    pb.add_h2(doc, "4.6 Định lượng và Xếp hạng Mức độ Tẩy xanh Liên tục (Greenwashing Score)")
    pb.add_figure_clean(
        doc,
        "greenwashing_score.png",
        "Hình 8",
        "Bản đồ nhiệt điểm tẩy xanh (GW) từng năm và biểu đồ cột xếp hạng mức độ tẩy xanh trung bình của 7 doanh nghiệp",
        width_inches=5.6
    )

    # Table 6: Greenwashing Rank
    h6, d6 = tables_data["t6_vi"]
    pb.add_table_clean(
        doc,
        table_label="Bảng 6",
        table_title="Bảng xếp hạng mức độ tẩy xanh định lượng (Greenwashing Index) của 7 doanh nghiệp niêm yết",
        headers=h6,
        data=d6,
        note="Điểm GW tính theo Công thức (5) chỉ kích hoạt khi giá cổ phiếu suy giảm. Phân cấp: GW = 0 (Không/Trung thực); 0 < GW <= 2 (Nhẹ); 2 < GW <= 8 (Trung bình); GW > 8 (Mạnh).",
        col_widths=[1.0, 1.8, 1.8, 1.8, 1.8, 2.0, 2.0, 1.8, 1.8],
        alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
        ],
        font_size=8.0,
    )

    pb.add_p(
        doc,
        "Hình 8 và Bảng 6 tổng hợp kết quả tính toán chỉ số tẩy xanh định lượng liên tục (GW Index). Sau khi dữ liệu scan của PNJ năm 2022 được chuẩn hóa bằng OCR, "
        "Bảng xếp hạng ghi nhận Tập đoàn Bảo Việt (BVH) đứng đầu với GW trung bình 3,02 điểm và GW cực đại đạt 13,96 điểm vào năm 2023. Hiện tượng này phản ánh bối cảnh "
        "năm 2023 ngành bảo hiểm nhân thọ đối mặt khủng hoảng niềm tin nghiêm trọng về kênh bancassurance và các đợt thanh tra gắt gao của cơ quan quản lý khiến giá cổ phiếu BVH sụt giảm -13,18%, "
        "nhưng báo cáo tích hợp lại đẩy mạnh chiến lược truyền thông phòng thủ với tỷ số Pos/Neg tăng vọt +105,89%. "
        "Đứng thứ hai là Vicostone (VCS) với GW trung bình 1,23 điểm (cực đại 5,16 điểm năm 2022, khi thị trường bất động sản Bắc Mỹ đóng băng do lãi suất Fed tăng cao làm giá cổ phiếu VCS giảm -47,46% "
        "nhưng báo cáo vẫn duy trì giọng văn ca ngợi với Pos/Neg tăng +10,88%). "
        "PNJ lùi xuống vị trí thứ ba với GW trung bình 0,92 điểm và cực đại 4,62 điểm năm 2023 (thay vì mức cực đại sai lệch 21,75 điểm trước khi có OCR). "
        "Vinamilk (VNM) đứng thứ tư với GW trung bình 0,35 điểm (mức tẩy xanh nhẹ mang tính tín hiệu chiến lược khi doanh nghiệp liên tục mở rộng diễn ngôn Net Zero). "
        "Ngược lại, cả ba doanh nghiệp The PAN Group (PAN), Petrolimex (PLX) và Chứng khoán SSI đều đạt điểm GW trung bình bằng 0,00 điểm. "
        "Trong năm 2022 đầy biến động khi thị trường chứng khoán giảm điểm sâu (giá cổ phiếu giảm từ -38,69% đến -62,17%), cả PAN, PLX và SSI đều thẳng thắn hạ thấp tỷ số cảm xúc Pos/Neg "
        "(giảm từ -4,22% đến -21,62%), thể hiện tính trung thực cao trong công bố thông tin thay vì tìm cách che đậy thực trạng khó khăn."
    )
    pb.add_p(
        doc,
        "Sự thay đổi căn bản trong bảng xếp hạng sau khi phục hồi dữ liệu OCR của PNJ là một bằng chứng thực nghiệm đắt giá về tầm quan trọng của tiền xử lý dữ liệu. "
        "Phần thảo luận dưới đây sẽ phân tích sâu sắc các phát hiện kinh tế - xã hội đằng sau những con số này."
    )

    # =========================================================================
    # 5. THẢO LUẬN CHUYÊN SÂU VÀ HÀM Ý CHÍNH SÁCH (DISCUSSION)
    # =========================================================================
    pb.add_h1(doc, "5. Thảo luận chuyên sâu và Hàm ý chính sách (Discussion)")

    pb.add_h2(doc, "5.1 Thực trạng Công bố ESG tại Việt Nam qua Lăng kính Định lượng AI")
    pb.add_p(
        doc,
        "Kết quả phân tích định lượng từ 96.461 câu văn bản trong 42 báo cáo đã phác họa một bức tranh toàn cảnh, khách quan và chân thực về hiện trạng công bố thông tin "
        "phát triển bền vững tại Việt Nam qua ba đặc trưng cốt lõi: "
        "(1) Sự thiên lệch áp đảo của tư duy kinh tế (Economic Dominance): Nhóm Kinh tế (SDG 8, 9) đạt điểm trung bình cao nhất (47,69 điểm), kết hợp với Đối tác phát triển (SDG 17) luôn giữ vị trí độc tôn. Điều này phản ánh thực tế "
        "rằng đối với các tập đoàn niêm yết tại một quốc gia đang phát triển như Việt Nam, mục tiêu tăng trưởng doanh thu, bảo toàn vốn và đầu tư hạ tầng kỹ thuật "
        "vẫn là ưu tiên số một để thỏa mãn kỳ vọng cổ đông ngắn hạn; "
        "(2) Khoảng trống cam kết về bình đẳng và quyền con người (Equity Gap): Nhóm Bình đẳng (SDG 4, 5, 10) liên tục xếp cuối bảng với điểm số trung bình chỉ đạt 42,08 điểm. Các nội dung "
        "về bình đẳng giới ở cấp quản trị thượng tầng, hòa nhập người yếu thế và giảm bất bình đẳng thu nhập nội bộ vẫn chỉ được đề cập mờ nhạt mang tính hình thức; "
        "(3) Tính lạc quan có hệ thống (Systemic Optimism Bias): Tỷ lệ câu tích cực chiếm tới 53,87% (51.966 câu) và câu trung tính chiếm 32,87% (31.706 câu), trong khi câu tiêu cực chỉ chiếm 13,26% (12.789 câu). "
        "Báo cáo phát triển bền vững tại Việt Nam hiện nay phần lớn vẫn đang được định vị như một tài liệu quan hệ công chúng (PR) cao cấp nhằm quảng bá hình ảnh "
        "thay vì là một công cụ quản trị rủi ro toàn diện."
    )

    pb.add_h2(doc, "5.2 Phân tích Chuyên sâu Sắc thái Cảm xúc và Bối cảnh Doanh nghiệp: Tẩy xanh hay Trung thực?")
    pb.add_p(
        doc,
        "Đi sâu vào diễn biến tỷ lệ Pos/Neg thường niên đối chiếu với các sự kiện kinh tế vĩ mô và bối cảnh ngành, nghiên cứu ghi nhận những phát hiện (findings) thực nghiệm đặc biệt giá trị: "
    )
    pb.add_p(
        doc,
        "(1) Khủng hoảng niềm tin ngành bảo hiểm và hiện tượng quản trị ấn tượng tại Tập đoàn Bảo Việt (BVH): "
        "Năm 2023, thị trường chứng khoán Việt Nam chứng kiến cơn địa chấn của ngành bảo hiểm nhân thọ xoay quanh các vụ việc tranh chấp hợp đồng bảo hiểm liên kết đầu tư "
        "và hoạt động bán chéo qua kênh ngân hàng (bancassurance). Làn sóng khiếu nại của khách hàng cùng các đợt thanh tra gắt gao từ Bộ Tài chính đã khiến tâm lý thị trường "
        "hoang mang, kéo giá cổ phiếu BVH sụt giảm -13,18%. Tuy nhiên, trong Báo cáo tích hợp 2023, BVH lại bất ngờ gia tăng đột biến tần suất thông điệp tích cực: "
        "tỷ số Pos/Neg nhảy vọt từ 3,02 lên 6,22 lần (+105,89%). Việc sử dụng ngôn từ lạc quan dày đặc để khẳng định vị thế thương hiệu và cam kết xã hội ngay giữa tâm bão khủng hoảng "
        "đã khiến chỉ số Greenwashing của BVH chạm đỉnh 13,96 điểm (mức Mạnh). Đây là minh chứng điển hình của chiến lược 'quản trị ấn tượng phòng thủ' (defensive impression management) "
        "nhằm xoa dịu áp lực từ các bên liên quan."
    )
    pb.add_p(
        doc,
        "(2) Tác động của suy thoái thị trường xuất khẩu bất động sản Bắc Mỹ đối với Vicostone (VCS): "
        "Năm 2022, chu kỳ thắt chặt tiền tệ quyết liệt của Cục Dự trữ Liên bang Mỹ (Fed) với các đợt tăng lãi suất dồn dập đã đẩy lãi suất vay mua nhà tại Mỹ lên đỉnh nhiều thập kỷ, "
        "khiến thị trường xây dựng và bất động sản nhà ở tại Bắc Mỹ (thị trường tiêu thụ đóng góp trên 60% doanh thu của VCS) rơi vào đình trệ nghiêm trọng. "
        "Hậu quả là kết quả kinh doanh và thị giá cổ phiếu VCS lao dốc -47,46%. Mặc dù vậy, báo cáo PTBV năm 2022 của VCS vẫn duy trì tỷ số Pos/Neg tăng +10,88% (từ 3,31 lên 3,67 lần), "
        "tập trung ca ngợi dây chuyền công nghệ đá ốp lát thạch anh Breton và các chứng chỉ xanh quốc tế, kích hoạt điểm GW = 5,16 điểm (mức Trung bình)."
    )
    pb.add_p(
        doc,
        "(3) Khủng hoảng thị trường vốn năm 2022 và thái độ công bố 'Trung thực' mẫu mực của PAN, PLX và SSI: "
        "Năm 2022 là năm thử thách khắc nghiệt nhất đối với thị trường tài chính Việt Nam khi chỉ số VN-Index sụt giảm hơn 32% do khủng hoảng thanh khoản trái phiếu doanh nghiệp, "
        "siết chặt tín dụng bất động sản và các vụ xử lý sai phạm thị trường vốn. Cổ phiếu SSI giảm tới -62,17%, The PAN Group giảm -61,17% và Petrolimex giảm -38,69%. "
        "Tuy nhiên, cả ba doanh nghiệp này đều ghi nhận điểm Greenwashing tuyệt đối bằng 0,00 nhờ thái độ công bố thông tin thẳng thắn và minh bạch: "
        "- SSI chủ động phân tích sự co hẹp thanh khoản thị trường và rủi ro sụt giảm doanh thu môi giới/cho vay ký quỹ, làm tỷ số Pos/Neg giảm -4,22%; "
        "- The PAN Group trình bày chi tiết các áp lực lạm phát chi phí đầu vào nông nghiệp (giá phân bón thế giới tăng phi mã, thức ăn chăn nuôi leo thang) và đứt gãy chuỗi cung ứng xuất khẩu thủy sản, làm tỷ số Pos/Neg giảm -7,92%; "
        "- Petrolimex đối mặt với cú sốc kép từ biến động giá dầu dị biệt do xung đột Nga - Ukraine và quy định giá cơ sở xăng dầu trong nước khiến khâu bán lẻ chịu lỗ định mức, dẫn đến tỷ số Pos/Neg giảm mạnh -21,62%. "
        "Hành vi đồng pha giữa thực tế khó khăn kinh doanh và giọng văn báo cáo đã khẳng định tính trung thực cao độ của nhóm doanh nghiệp này."
    )
    pb.add_p(
        doc,
        "(4) Chiến lược định vị vị thế dẫn dắt ESG và tín hiệu Net Zero của Vinamilk (VNM): "
        "VNM là doanh nghiệp duy nhất duy trì tỷ số Pos/Neg trên 4,3 lần trong suốt giai đoạn khảo sát, chạm đỉnh 6,06 lần vào năm 2024 (trung bình 5,38 lần). "
        "Dù cổ phiếu có những nhịp điều chỉnh nhẹ do thị trường sữa nội địa bước vào giai đoạn bão hòa, VNM liên tục công bố các cột mốc thực chất: "
        "Nhà máy sữa Nghệ An và Trang trại bò sữa Vinamilk Nghệ An đạt chứng nhận trung hòa carbon theo chuẩn quốc tế PAS 2060, trang trại sinh thái Vinamilk Green Farm, "
        "và lộ trình Net Zero 2050 bài bản. Mức điểm GW dao động từ 0,11 đến 1,13 của VNM không mang tính chất lừa dối cổ đông mà là hệ quả tự nhiên của việc doanh nghiệp "
        "đang thực sự dẫn dắt xu hướng chuyển đổi xanh tại Việt Nam."
    )

    pb.add_h2(doc, "5.3 Đối chuẩn So sánh Thực nghiệm với Nghiên cứu của Kang & Kim (2022)")
    pb.add_p(
        doc,
        f"So sánh đối chuẩn với kết quả nghiên cứu của {c_kang} trên các chaebol Hàn Quốc đem lại những phát hiện học thuật vô cùng thú vị. "
        f"Điểm tương đồng lớn nhất là cả hai thị trường đều ghi nhận trật tự phân phối điểm SDG dạng chuông đối xứng và xu hướng áp đảo của nhóm Kinh tế "
        f"so với nhóm Công bằng xã hội. Tuy nhiên, có hai điểm khác biệt mang tính bản chất: "
        f"(1) Tại Hàn Quốc, các tập đoàn chaebol có sự phân hóa sâu sắc hơn về điểm số nhóm Môi trường (Environments) tùy thuộc vào ngành thâm dụng phát thải "
        f"(như thép POSCO, hóa chất LG Chem) so với viễn thông (SK Telecom); tại Việt Nam, điểm số nhóm môi trường giữa các doanh nghiệp đa ngành lại có độ phân tán hẹp hơn, "
        f"cho thấy mức độ chuyên biệt hóa số liệu phát thải ở Việt Nam chưa cao; "
        f"(2) Tỷ lệ câu tích cực tại Việt Nam (53,87%) cao hơn so với kết quả tại Hàn Quốc, phản ánh văn hóa truyền thông doanh nghiệp Á Đông tại các thị trường mới nổi "
        f"thường có tâm lý e ngại phơi bày các số liệu tiêu cực trước công chúng."
    )

    pb.add_h2(doc, "5.4 Bài học Phương pháp luận: Phục hồi Dữ liệu Scan qua OCR và Rủi ro Diễn giải Chỉ số AI")
    pb.add_p(
        doc,
        "Một trong những đóng góp học thuật và thực tiễn cốt lõi nhất của công trình này là việc đưa ra cảnh báo phương pháp luận sâu sắc về 'bẫy dữ liệu số hóa' "
        "(Data Quality Artifact) khi áp dụng Trí tuệ Nhân tạo và Xử lý Ngôn ngữ Tự nhiên trong phân tích tài chính - kế toán. "
        "Điển hình là trường hợp Báo cáo PTBV năm 2022 của PNJ: tài liệu gốc được phát hành dưới dạng tệp scan hình ảnh (raster PDF) không chứa lớp văn bản điện tử. "
        "Nếu chỉ áp dụng bộ trích xuất văn bản tiêu chuẩn (PyMuPDF không tích hợp OCR), hệ thống chỉ bóc tách được vỏn vẹn 12 câu văn bản (chủ yếu là các lưu ý bản quyền "
        "và tuyên bố miễn trừ trách nhiệm pháp lý mang từ ngữ cảnh báo tiêu cực), làm tỷ số Pos/Neg năm 2022 sụt giảm nhân tạo xuống đáy 0,33 lần. "
        "Đến năm 2023, khi PNJ phát hành tài liệu số hóa chuẩn mực với 1.050 câu và tỷ số Pos/Neg phục hồi về 4,23 lần, thuật toán toán học ghi nhận mức tăng cơ học "
        "lên tới +1.169%, kết hợp với mức giảm giá nhẹ của cổ phiếu (-1,86%) tạo nên điểm số GW cực đại hoàn toàn giả tạo lên tới 21,75 điểm—sai lệch nghiêm trọng thực tế khách quan!"
    )
    pb.add_p(
        doc,
        "Để giải quyết triệt để rủi ro phương pháp luận này, nghiên cứu đã xây dựng đường ống OCR quang học chuyên dụng (Tesseract 5.4.0, gói vie+eng, 300 DPI, "
        "nhị phân hóa thích nghi và lọc nhiễu đồ họa). Kết quả đã khôi phục trọn vẹn 669 câu văn bản thực chất từ 52 trang scan, xác định chính xác tỷ số Pos/Neg "
        "năm 2022 của PNJ là 1,21 lần (317 câu tích cực và 261 câu tiêu cực). Sắc thái thận trọng này phản ánh đúng bối cảnh PNJ vừa trải qua đợt phong tỏa kéo dài "
        "do dịch COVID-19 cuối năm 2021 tại TP.HCM (tạm dừng hoạt động hơn 80% cửa hàng) và phải đối mặt với áp lực tái cơ cấu chi phí vận hành, an toàn lao động và giá vàng thế giới. "
        "Nhờ dữ liệu OCR thực chất, bước nhảy tỷ số sang năm 2023 chỉ còn +248,3%, giúp kéo giảm điểm GW năm 2023 của PNJ từ đỉnh sai lệch 21,75 điểm xuống còn 4,62 điểm (mức Trung bình), "
        "và đưa vị trí xếp hạng của PNJ từ 'tẩy xanh nghiêm trọng nhất' xuống vị trí thứ 3 toàn đoàn."
    )
    pb.add_p(
        doc,
        "Bài học phương pháp luận rút ra là: các hệ thống AI phân tích tài chính tự động không bao giờ được phép hoạt động như một 'hộp đen' (black-box). "
        "Người nghiên cứu và chuyên viên phân tích bắt buộc phải tích hợp các quy trình kiểm soát vệ sinh dữ liệu (data hygiene validation), "
        "đặc biệt là công nghệ OCR chuyên dụng cho tài liệu scan, trước khi đưa ra bất kỳ kết luận hành vi nào đối với doanh nghiệp niêm yết."
    )

    pb.add_h2(doc, "5.5 Khuyến nghị Chính sách và Giải pháp Thực tiễn")
    pb.add_p(
        doc,
        "Từ các phát hiện thực nghiệm, nghiên cứu đề xuất 3 nhóm khuyến nghị chính sách trọng tâm: "
        "(1) Đối với Cơ quan Quản lý (UBCKNN và các Sở GDCK): Cần sớm ban hành quy chuẩn số hóa bắt buộc đối với báo cáo phát triển bền vững (như áp dụng chuẩn báo cáo "
        "điện tử XBRL hoặc PDF searchable có cấu trúc thẻ tag chuẩn), chấm dứt tình trạng phát hành các bản scan ảnh chất lượng thấp làm cản trở quá trình giám sát tự động; "
        "(2) Đối với Doanh nghiệp Niêm yết: Cần chuyển đổi tư duy từ báo cáo hình thức sang báo cáo thực chất, gia tăng tỷ trọng công bố các mục tiêu nhạy cảm "
        "như bình đẳng giới, đào tạo nhân lực và đa dạng sinh học, đồng thời chủ động công bố các khó khăn, thách thức để gia tăng độ tin cậy đối với các quỹ đầu tư quốc tế; "
        "(3) Đối với Nhà đầu tư và Công chúng: Cần sử dụng các công cụ phân tích văn bản NLP như một bộ lọc sàng lọc sơ bộ hiệu quả để phát hiện các dấu hiệu bất thường, "
        "kết hợp chặt chẽ với kiểm tra số liệu kiểm toán độc lập để đưa ra quyết định đầu tư chính xác."
    )

    pb.add_h2(doc, "5.6 Hạn chế của Nghiên cứu và Định hướng Phát triển")
    pb.add_p(
        doc,
        "Mặc dù đã đạt được những kết quả toàn diện, nghiên cứu vẫn tồn tại một số hạn chế nhất định mở ra hướng phát triển trong tương lai: "
        "(1) Cỡ mẫu 7 doanh nghiệp tuy mang tính đại diện đầu ngành nhưng vẫn chưa bao quát toàn bộ hơn 700 doanh nghiệp niêm yết trên hai sàn HOSE và HNX; "
        "(2) Mô hình phân tích hiện tại chỉ tập trung vào dữ liệu văn bản chữ, chưa bóc tách và kiểm chứng chéo với dữ liệu số trong các bảng biểu và phụ lục GRI; "
        "(3) Hướng nghiên cứu tiếp theo có thể ứng dụng các Mô hình Ngôn ngữ Lớn (LLMs) thế hệ mới để thực hiện trích xuất sự kiện và đối chiếu cam kết ESG "
        "với thông tin phản ánh từ báo chí truyền thông độc lập nhằm tăng cường khả năng phát hiện tẩy xanh thời gian thực."
    )

    # =========================================================================
    # 6. KẾT LUẬN (CONCLUSION)
    # =========================================================================
    pb.add_h1(doc, "6. Kết luận (Conclusion)")
    pb.add_p(
        doc,
        f"Nghiên cứu này đã hoàn thành việc tái hiện, chuẩn hóa và mở rộng toàn diện khung phân tích văn bản của {c_kang} cho thị trường chứng khoán Việt Nam. "
        f"Thông qua việc xử lý định lượng 96.461 câu văn bản từ 42 báo cáo phát triển bền vững của 7 doanh nghiệp niêm yết lớn trong giai đoạn 2020–2025 "
        f"(bao gồm khôi phục hoàn chỉnh dữ liệu scan bằng công nghệ OCR), "
        f"công trình đã cung cấp những bằng chứng thực nghiệm vững chắc về cơ cấu cam kết 17 mục tiêu SDG, sắc thái cảm xúc ngôn ngữ và mối liên hệ giữa báo cáo phi tài chính "
        f"với thị trường cổ phiếu. Các phát hiện về sự áp đảo của nhóm Kinh tế, khoảng trống nhóm Bình đẳng, tính lạc quan mang tính cấu trúc và sự phân kỳ hoàn toàn "
        f"giữa giọng văn báo cáo với giá thị trường (r = -0,043, p = 0,806) đã làm sáng tỏ bức tranh thực tế về công bố ESG tại một thị trường mới nổi. "
        f"Đặc biệt, bài học sâu sắc về bẫy diễn giải chỉ số tẩy xanh do chất lượng số hóa tài liệu cùng giải pháp can thiệp kỹ thuật OCR thành công đã đóng góp một kinh nghiệm "
        f"phương pháp luận vô giá cho cộng đồng nghiên cứu kế toán - tài chính ứng dụng trí tuệ nhân tạo."
    )

    # =========================================================================
    # TÀI LIỆU THAM KHẢO (REFERENCES)
    # =========================================================================
    pb.add_h1(doc, "Tài liệu tham khảo")
    
    apa_refs = [
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

    for ref in apa_refs:
        rp = doc.add_paragraph()
        rp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        rp.paragraph_format.left_indent = Inches(0.5)
        rp.paragraph_format.first_line_indent = Inches(-0.5)
        rp.paragraph_format.space_after = Pt(4)
        run = rp.add_run(ref)
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)
        run.font.color.rgb = pb.COLOR_BLACK

    print("[VI] Built complete Vietnamese paper structure.")
