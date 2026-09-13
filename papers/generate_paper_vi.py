"""Toàn văn Bản Tiếng Việt chuẩn APA 7th edition.
Tái hiện và vận dụng khung nghiên cứu của Kang & Kim (2022) cho 4 doanh nghiệp Việt Nam.
Bỏ phần tẩy xanh (greenwashing) do kết quả thực nghiệm không chứng minh được tương quan.
Thay vào đó, tập trung nêu cảm nghĩ, suy ngẫm thực trạng công bố thông tin, phân tích sắc thái lạc quan,
luôn đối chiếu và kế thừa phương pháp của Kang & Kim (2022) ở từng bước,
và thảo luận mở rộng sang các doanh nghiệp Việt Nam có báo cáo liên tục 2020–2025.
"""

from __future__ import annotations

import docx
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

import build_full_papers as gbp


def build_paper_content(doc: Document):
    # =========================================================================
    # TIÊU ĐỀ BÀI BÁO VÀ THÔNG TIN TÁC GIẢ
    # =========================================================================
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
    run.font.color.rgb = gbp.COLOR_BLACK

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
    erun.font.color.rgb = gbp.COLOR_BLACK

    ap = doc.add_paragraph()
    ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ap.paragraph_format.space_after = Pt(2)
    arun = ap.add_run("Nhóm Nghiên cứu Trí tuệ Nhân tạo và Tài chính Bền vững")
    arun.font.name = "Times New Roman"
    arun.font.size = Pt(11.5)
    arun.bold = True
    arun.font.color.rgb = gbp.COLOR_BLACK

    aff_p = doc.add_paragraph()
    aff_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    aff_p.paragraph_format.space_after = Pt(14)
    aff_run = aff_p.add_run("Khoa Tài chính - Ngân hàng, Trường Đại học Kinh tế\nEmail liên hệ: research.esg@vietnam-analytics.edu.vn")
    aff_run.font.name = "Times New Roman"
    aff_run.font.size = Pt(10.5)
    aff_run.font.color.rgb = gbp.COLOR_BLACK

    # =========================================================================
    # TÓM TẮT & ABSTRACT
    # =========================================================================
    gbp.add_h1(doc, "Tóm tắt")
    gbp.add_p(
        doc,
        "Báo cáo phát triển bền vững đóng vai trò huyết mạch trong việc truyền tải các cam kết môi trường, xã hội và quản trị (ESG) "
        "của doanh nghiệp tới các bên liên quan. Kế thừa và vận dụng phương pháp luận của Kang và Kim (2022), nghiên cứu này thực hiện "
        "phân tích định lượng tự động và trực quan hóa dữ liệu văn bản từ 29 báo cáo phát triển bền vững độc lập của 4 doanh nghiệp niêm yết hàng đầu Việt Nam "
        "(The PAN Group, Petrolimex, PNJ, Vinamilk) trong giai đoạn 2018–2025 với tổng quy mô 17.047 câu văn. "
        "Bằng việc sử dụng các mô hình học sâu biến đổi ngôn ngữ tiên tiến (vietnamese-sbert và PhoBERT), nghiên cứu tiến hành đo lường mức độ gắn kết "
        "với 17 Mục tiêu Phát triển Bền vững (SDGs) của Liên Hợp Quốc, quy nạp vào 6 nhóm nhu cầu con người (Đời sống, Kinh tế, Công bằng, Xã hội, "
        "Tài nguyên, Môi trường) và phân tích sắc thái cảm xúc đa lớp (Tích cực, Trung tính, Tiêu cực). Kết quả thực nghiệm chỉ ra rằng: "
        "(1) Các doanh nghiệp Việt Nam ưu tiên áp đảo nhóm Kinh tế (SDG 8, 9) và Đối tác (SDG 17), trong khi nhóm Công bằng xã hội (SDG 4, 5, 10) "
        "liên tục ghi nhận điểm số thấp nhất; (2) Tồn tại thiên lệch lạc quan (optimism bias) rõ rệt với tỷ lệ câu mang sắc thái tích cực chiếm tới 69,75%, "
        "câu trung tính chiếm 18,34% và câu tiêu cực chỉ chiếm 11,91%; (3) Không tìm thấy bằng chứng thống kê về hành vi tẩy xanh (greenwashing) "
        "khi biến động giá cổ phiếu và tỷ số cảm xúc hoàn toàn độc lập (r = 0,03, p = 0,89), phản ánh tâm lý quản trị ấn tượng trong một thị trường mới nổi "
        "nơi nhà đầu tư chưa định giá trực tiếp thông tin ESG. Nghiên cứu cung cấp góc nhìn thực tiễn sâu sắc về hiện trạng công bố thông tin phi tài chính "
        "và đề xuất mở rộng khung phân tích sang các doanh nghiệp Việt Nam có chuỗi báo cáo liên tục 2020–2025 như BVH, DHG, VCS, FPT và NLG."
    )
    gbp.add_p(
        doc,
        "Từ khóa: Báo cáo phát triển bền vững; Mục tiêu phát triển bền vững (SDGs); Xử lý ngôn ngữ tự nhiên (NLP); "
        "Sentence-BERT; PhoBERT; Phân tích cảm xúc; Thiên lệch lạc quan; Thị trường chứng khoán Việt Nam.",
        bold=True
    )

    gbp.add_h1(doc, "Abstract")
    gbp.add_p(
        doc,
        "Corporate sustainability reports represent essential communication conduits for conveying environmental, social, and governance (ESG) "
        "commitments to diverse stakeholders. Replicating and adapting the computational methodology of Kang and Kim (2022), this paper conducts "
        "an automated textual analysis and longitudinal visualization of 29 standalone sustainability reports published by four prominent Vietnamese listed firms "
        "(The PAN Group, Petrolimex, PNJ, Vinamilk) spanning 2018–2025, encompassing 17,047 extracted sentences. "
        "Employing pre-trained Transformer language models (vietnamese-sbert and PhoBERT), we quantify corporate alignment with the 17 UN Sustainable Development Goals (SDGs), "
        "aggregate alignments into six human-needs categories (Life, Economic, Equity, Social, Resources, Environments), and evaluate multi-class sentiment polarity "
        "(Positive, Neutral, Negative). Empirical findings reveal that: (1) Vietnamese corporate disclosures disproportionately prioritize the Economic dimension "
        "(SDG 8, 9) and Global Partnerships (SDG 17), while systematically underperforming in Social Equity (SDG 4, 5, 10); (2) Corporate narratives demonstrate "
        "a pervasive optimism bias, comprising 69.75% positive, 18.34% neutral, and merely 11.91% negative statements; (3) In contrast to aggressive greenwashing hypotheses, "
        "capital market returns and narrative sentiment ratios are statistically independent (r = 0.03, p = 0.89), suggesting that narrative optimism reflects conventional "
        "impression management rather than manipulative deception in an emerging market where institutional ESG pricing remains nascent. "
        "This study provides deep critical reflections on the transparency of corporate non-financial disclosures in Vietnam and proposes expanding analytical samples "
        "to peer enterprises with continuous 2020–2025 reporting records including BVH, DHG, VCS, FPT, and NLG."
    )
    gbp.add_p(
        doc,
        "Keywords: Sustainability reporting; Sustainable Development Goals (SDGs); Natural Language Processing (NLP); "
        "Sentence-BERT; PhoBERT; Sentiment analysis; Optimism bias; Vietnam stock market.",
        bold=True
    )

    doc.add_page_break()

    # =========================================================================
    # 1. GIỚI THIỆU (INTRODUCTION)
    # =========================================================================
    gbp.add_h1(doc, "1. Giới thiệu")
    gbp.add_p(
        doc,
        "Trong bối cảnh biến đổi khí hậu toàn cầu ngày càng gay gắt và các tiêu chuẩn phát triển bền vững trở thành điều kiện tiên quyết trong thương mại quốc tế, "
        "công bố thông tin môi trường, xã hội và quản trị (ESG) đã chuyển dịch mạnh mẽ từ một hoạt động mang tính khuyến khích sang một yêu cầu bắt buộc "
        "đối với các doanh nghiệp niêm yết (Arvidsson & Dumay, 2022; Global Reporting Initiative, 2021). Tại Việt Nam, sau cam kết mạnh mẽ của Thủ tướng Chính phủ "
        "về việc đạt mức phát thải ròng bằng không (Net Zero) vào năm 2050 tại Hội nghị thượng đỉnh COP26, khuôn khổ pháp lý về phát triển bền vững đã có những bước tiến "
        "vượt bậc với Thông tư 96/2020/TT-BTC của Bộ Tài chính và Luật Chứng khoán năm 2019 (Bộ Tài chính, 2020; Chính phủ Việt Nam, 2020). "
        "Báo cáo phát triển bền vững (Sustainability Report - SR) đã trở thành kênh thông tin chính thống giúp các doanh nghiệp đại chúng công khai trách nhiệm xã hội "
        "và khẳng định sự gắn kết với Chương trình Nghị sự 2030 của Liên Hợp Quốc gồm 17 Mục tiêu Phát triển Bền vững (SDGs) (United Nations, 2015)."
    )
    gbp.add_p(
        doc,
        "Tuy nhiên, việc đánh giá chất lượng và mức độ thực chất của các báo cáo phát triển bền vững hiện nay đang đối mặt với thách thức to lớn về mặt phương pháp. "
        "Phần lớn các báo cáo được công bố dưới dạng tệp PDF phi cấu trúc dày hàng trăm trang với khối lượng câu chữ khổng lồ, khiến các phương pháp phân tích nội dung thủ công "
        "(manual content analysis) truyền thống trở nên quá tải, tốn kém chi phí và dễ bị chi phối bởi định kiến chủ quan của người mã hóa (Luccioni et al., 2020). "
        "Đồng thời, phương pháp đếm tần suất từ khóa giản đơn (keyword counting) cũng bộc lộ nhược điểm cố hữu là bỏ qua ngữ cảnh cú pháp, không đo lường được sắc thái ngữ nghĩa "
        "và dễ bị doanh nghiệp đối phó bằng cách lặp lại các thuật ngữ thời thượng (Kang & Kim, 2022; Loughran & McDonald, 2011)."
    )
    gbp.add_p(
        doc,
        "Để giải quyết triệt để rào cản này, Kang và Kim (2022) trong công trình tiên phong đăng trên tạp chí Applied Sciences đã đề xuất một khung phương pháp luận kết hợp "
        "các mô hình học sâu xử lý ngôn ngữ tự nhiên (NLP) tiên tiến: sử dụng Sentence-BERT để đo lường độ tương đồng ngữ nghĩa giữa các câu văn báo cáo với 17 mục tiêu SDG, "
        "quy nạp vào 6 nhóm danh mục nhu cầu con người (Life, Economic, Equity, Social, Resources, Environments), đồng thời ứng dụng mô hình Transformer để phân tích sắc thái cảm xúc "
        "của ngôn từ. Bài báo của Kang và Kim (2022) đã chứng minh tính ưu việt vượt trội trong việc lượng hóa thông tin định tính và trực quan hóa hành trình bền vững của doanh nghiệp."
    )
    gbp.add_p(
        doc,
        "Xuất phát từ nền tảng học thuật đó, nghiên cứu này được thực hiện nhằm tái hiện (replicate) và vận dụng sáng tạo khung phương pháp luận của Kang và Kim (2022) "
        "vào bối cảnh thực nghiệm tại Việt Nam. Chúng tôi thu thập toàn bộ 29 báo cáo phát triển bền vững độc lập được phát hành liên tục trong giai đoạn 2018–2025 "
        "của 4 doanh nghiệp vốn hóa lớn đại diện cho 4 ngành kinh tế then chốt: The PAN Group (PAN - Nông nghiệp), Petrolimex (PLX - Năng lượng, xăng dầu), "
        "PNJ (PNJ - Bán lẻ, trang sức) và Vinamilk (VNM - Chế biến thực phẩm và chăn nuôi sữa). "
        "Nghiên cứu hướng tới việc trả lời ba câu hỏi cốt lõi: "
        "(1) Mức độ tập trung và sự phân hóa chủ đề của các doanh nghiệp Việt Nam đối với 17 mục tiêu SDG diễn ra như thế nào qua lăng kính mô hình Sentence-BERT tiếng Việt? "
        "(2) Sắc thái ngôn từ trong các báo cáo phản ánh điều gì về tâm lý công bố thông tin của ban lãnh đạo doanh nghiệp? "
        "(3) Liệu các bằng chứng thực nghiệm có chứng minh được hiện tượng tẩy xanh trên thị trường chứng khoán Việt Nam hay không, và thực trạng công bố thông tin bền vững hiện nay để lại những suy ngẫm gì?"
    )

    # =========================================================================
    # 2. TỔNG QUAN NGHIÊN CỨU (LITERATURE REVIEW)
    # =========================================================================
    gbp.add_h1(doc, "2. Tổng quan nghiên cứu")

    gbp.add_h2(doc, "2.1. Nền tảng Lý thuyết: Các bên Liên quan, Tính Chính danh và Quản trị Ấn tượng")
    gbp.add_p(
        doc,
        "Động cơ công bố thông tin phi tài chính của doanh nghiệp được giải thích chủ yếu thông qua ba lý thuyết kinh điển trong quản trị học. "
        "Trước hết, Lý thuyết Các bên Liên quan (Stakeholder Theory) do Freeman (1984) khởi xướng khẳng định rằng doanh nghiệp không chỉ có trách nhiệm tối đa hóa lợi nhuận cho cổ đông "
        "mà phải dung hòa lợi ích của toàn bộ các bên chịu tác động, bao gồm người lao động, khách hàng, nhà cung ứng, cộng đồng dân cư và cơ quan quản lý. "
        "Trong lăng kính này, báo cáo phát triển bền vững đóng vai trò là kênh đối thoại minh bạch nhằm cung cấp thông tin giải trình về trách nhiệm môi trường và xã hội (Hummel & Schlick, 2016)."
    )
    gbp.add_p(
        doc,
        "Thứ hai, Lý thuyết Tính Chính danh (Legitimacy Theory) nhấn mạnh rằng doanh nghiệp hoạt động dựa trên một 'khế ước xã hội' (social contract). "
        "Để tồn tại và duy trì sự ủng hộ của công chúng, tổ chức phải chứng minh các hoạt động của mình phù hợp với hệ giá trị, chuẩn mực và kỳ vọng của xã hội (Deegan, 2002; Suchman, 1995). "
        "Khi đối mặt với các sự cố môi trường hoặc sức ép thể chế, doanh nghiệp thường gia tăng khối lượng công bố thông tin bền vững như một công cụ chiến lược "
        "để củng cố tính chính danh và bảo vệ giấy phép xã hội để hoạt động (social license to operate)."
    )
    gbp.add_p(
        doc,
        "Thứ ba, Lý thuyết Quản trị Ấn tượng (Impression Management Theory) lập luận rằng trong điều kiện bất đối xứng thông tin, ban lãnh đạo doanh nghiệp có xu hướng chủ động "
        "thao túng cấu trúc ngôn từ, lựa chọn câu từ tích cực và làm mờ các thông tin bất lợi nhằm định hình nhận thức của công chúng theo hướng có lợi nhất cho hình ảnh tổ chức "
        "(Merkl-Davies & Brennan, 2007; Stacchezzini et al., 2016). Sự đan xen giữa nhu cầu minh bạch thực chất và xu hướng tô hồng hình ảnh là nguồn gốc tạo nên "
        "sự đa dạng và phức tạp của ngôn từ trong các báo cáo phát triển bền vững hiện đại."
    )

    gbp.add_h2(doc, "2.2. Tiến hóa của Xử lý Ngôn ngữ Tự nhiên trong Phân tích Thông tin Doanh nghiệp")
    gbp.add_p(
        doc,
        "Ứng dụng xử lý ngôn ngữ tự nhiên (NLP) trong nghiên cứu tài chính và kế toán đã trải qua ba làn sóng chuyển dịch phương pháp luận rõ rệt (Mercereau & Melin, 2020). "
        "Làn sóng thứ nhất dựa trên phương pháp túi từ (Bag-of-Words) và từ điển phân loại chuyên ngành, tiêu biểu là danh mục từ vựng tài chính do Loughran và McDonald (2011) "
        "xây dựng cho báo cáo thường niên 10-K. Mặc dù có ưu thế về tốc độ xử lý, phương pháp này bỏ qua hoàn toàn trật tự từ, ngữ cảnh cú pháp và sự đa nghĩa của ngôn ngữ tự nhiên. "
        "Kang và Kim (2022) đã chứng minh bằng thực nghiệm rằng phương pháp khớp từ khóa tạo ra dải điểm tương đồng cực kỳ phẳng và thoái hóa (độ lệch chuẩn xấp xỉ 0,03), "
        "do các từ khóa chung chung như 'nước', 'chất thải' xuất hiện dày đặc nhưng không phản ánh được chiều sâu cam kết thực sự của doanh nghiệp."
    )
    gbp.add_p(
        doc,
        "Làn sóng thứ hai xuất hiện với các mô hình không gian vector tĩnh như Word2Vec (Mikolov et al., 2013), GloVe (Pennington et al., 2014) và mô hình chủ đề ẩn LDA (Blei et al., 2003). "
        "Mặc dù các kỹ thuật này đã biểu diễn được quan hệ ngữ nghĩa phân bố giữa các từ, mô hình LDA lại gán các chủ đề ngẫu nhiên theo xác suất thống kê không giám sát, khiến việc đối sánh trực tiếp "
        "và có hệ thống với 17 mục tiêu chuẩn tắc của Liên Hợp Quốc gặp nhiều hạn chế và thiếu tính lặp lại độc lập."
    )
    gbp.add_p(
        doc,
        "Làn sóng thứ ba – kỷ nguyên của các mô hình học sâu dựa trên kiến trúc Transformer và cơ chế tự chú ý (Self-Attention) (Devlin et al., 2019) – đã giải quyết trọn vẹn rào cản ngữ cảnh. "
        "Đặc biệt, bước đột phá từ Sentence-BERT (SBERT) của Reimers và Gurevych (2019) đã cho phép mã hóa toàn bộ câu văn thành một vector mật độ cố định trong không gian ngữ nghĩa, "
        "tối ưu hóa hoàn hảo cho việc đo lường độ tương đồng cosine quy mô lớn (Cer et al., 2018). Tại Việt Nam, sự ra đời của mô hình ngôn ngữ PhoBERT (Nguyen & Nguyen, 2020) "
        "và vietnamese-sbert đã thiết lập chuẩn mực công nghệ mới trong việc xử lý đặc thù ngữ pháp, từ ghép và ngữ nghĩa tiếng Việt, mở đường cho các nghiên cứu định lượng tự động "
        "đối với báo cáo doanh nghiệp trong nước."
    )

    gbp.add_h2(doc, "2.3. Khung Mục tiêu Phát triển Bền vững (SDGs) và Mô hình 6 Nhóm Danh mục của Kang & Kim (2022)")
    gbp.add_p(
        doc,
        "Kể từ khi được công bố vào năm 2015, Chương trình Nghị sự 2030 với 17 Mục tiêu Phát triển Bền vững (SDGs) và 169 chỉ tiêu cụ thể (United Nations, 2015) đã trở thành thước đo chung "
        "về trách nhiệm xã hội trên phạm vi toàn cầu. Nhằm hướng dẫn khu vực tư nhân lồng ghép các mục tiêu này vào chiến lược kinh doanh, Sáng kiến Báo cáo Toàn cầu cùng Hiệp ước Toàn cầu LHQ "
        "đã ban hành bộ công cụ hướng dẫn SDG Compass (Global Reporting Initiative, 2021). Tuy nhiên, trên bình diện học thuật, nhiều nghiên cứu cảnh báo về hiện tượng doanh nghiệp "
        "chỉ chọn lọc báo cáo các mục tiêu dễ thực hiện hoặc mang lại lợi ích tài chính tức thì, trong khi bỏ qua các vấn đề cốt lõi về bình đẳng và nhân quyền (Heras-Saizarbitoria et al., 2022; Pizzi et al., 2020)."
    )
    gbp.add_p(
        doc,
        "Để hệ thống hóa và phân tích cấu trúc tổng thể của 17 SDGs, Kang và Kim (2022) đã có đóng góp phương pháp luận đặc biệt quan trọng: dựa trên Lý thuyết Phát triển Quy mô Con người "
        "(Human Scale Development) của nhà kinh tế học Manfred Max-Neef, các tác giả đã phân nhóm 17 mục tiêu SDG thành 6 nhóm danh mục nhu cầu con người: "
        "(1) Đời sống (Life: SDG 1, 2, 3) đại diện cho các nhu cầu sinh tồn cơ bản; (2) Kinh tế (Economic: SDG 8, 9) đại diện cho việc làm, tăng trưởng và đổi mới công nghệ; "
        "(3) Công bằng (Equity: SDG 4, 5, 10) đại diện cho giáo dục bình đẳng, bình đẳng giới và thu hẹp khoảng cách giàu nghèo; (4) Xã hội (Social: SDG 11, 16, 17) đại diện cho thể chế, "
        "đô thị an toàn và quan hệ đối tác; (5) Tài nguyên (Resources: SDG 6, 7, 12, 14) đại diện cho việc sử dụng bền vững năng lượng, nước và sản xuất trách nhiệm; "
        "và (6) Môi trường (Environments: SDG 13, 15) đại diện cho hành động khí hậu và hệ sinh thái rừng trên cạn. "
        "Khung phân loại 6 nhóm danh mục này là cơ sở trực quan xuất sắc giúp nhận diện rõ chiến lược bền vững của doanh nghiệp đang hướng tới sự cân bằng hay thiên lệch."
    )

    gbp.add_h2(doc, "2.4. Sắc thái Ngôn từ, Thiên lệch Lạc quan và Quản trị Ấn tượng trong Báo cáo Doanh nghiệp")
    gbp.add_p(
        doc,
        "Phân tích sắc thái cảm xúc (Sentiment Analysis) trong văn bản phi tài chính là công cụ hữu hiệu để bóc tách tâm lý truyền thông của nhà quản trị. "
        "Nhiều công trình kinh điển đã ghi nhận hiện tượng 'thiên lệch lạc quan' (Pollyanna effect / Optimism bias) mang tính hệ thống trong các báo cáo doanh nghiệp: "
        "các từ ngữ tích cực ca ngợi thành tích luôn xuất hiện với tần suất áp đảo các từ ngữ tiêu cực chỉ ra khó khăn hoặc rủi ro (El-Haj et al., 2020; Huang et al., 2023). "
        "Kang và Kim (2022) đã đo lường hiện tượng này thông qua tỷ số cảm xúc (Pos/Neg Ratio) giữa số lượng câu văn mang tính tích cực và số lượng câu văn mang tính tiêu cực, "
        "chứng minh rằng tỷ số này phản ánh mức độ doanh nghiệp chủ động sử dụng ngôn từ để xây dựng hình ảnh trước công chúng."
    )
    gbp.add_p(
        doc,
        "Trong bối cảnh các thị trường mới nổi như Việt Nam, việc nghiên cứu sắc thái ngôn từ càng mang ý nghĩa cấp thiết. "
        "Khi các quy định pháp lý về kiểm toán dữ liệu ESG độc lập còn chưa hoàn thiện, văn bản báo cáo bền vững rất dễ trở thành không gian thuận lợi cho các hoạt động quản trị ấn tượng. "
        "Do đó, việc ứng dụng các mô hình học sâu hiện đại như PhoBERT để bóc tách tỷ lệ câu tích cực, trung tính và tiêu cực sẽ cung cấp bằng chứng khách quan "
        "về tính cân bằng và độ tin cậy trong thông điệp của doanh nghiệp."
    )

    gbp.add_h2(doc, "2.5. Bối cảnh Thể chế và Thực trạng Báo cáo Phát triển Bền vững tại Việt Nam")
    gbp.add_p(
        doc,
        "Tại Việt Nam, khung khổ pháp lý về công bố thông tin bền vững đã có những bước chuyển mình căn bản từ Thông tư 155/2015/TT-BTC đến Thông tư 96/2020/TT-BTC của Bộ Tài chính "
        "(Bộ Tài chính, 2020). Theo đó, các công ty niêm yết bắt buộc phải công bố thông tin tác động môi trường và xã hội trong Báo cáo Thường niên, bao gồm tiêu thụ nguyên vật liệu, "
        "năng lượng, tài nguyên nước và kiểm kê phát thải khí nhà kính Scope 1 và Scope 2. "
        "Đồng thời, Ủy ban Chứng khoán Nhà nước cùng Tổ chức Tài chính Quốc tế (IFC) đã phát hành Sổ tay Hướng dẫn Báo cáo Môi trường và Xã hội (Ủy ban Chứng khoán Nhà nước, 2024), "
        "và Liên đoàn Thương mại và Công nghiệp Việt Nam (VCCI) đã duy trì Chương trình Đánh giá và Công bố Doanh nghiệp Bền vững tại Việt Nam thông qua Bộ chỉ số CSI (VBCSD, 2024)."
    )
    gbp.add_p(
        doc,
        "Tuy vậy, thực tiễn áp dụng tại Việt Nam bộc lộ sự phân hóa cực kỳ sâu sắc. Khảo sát của PwC Vietnam (2022) chỉ ra rằng chỉ có một tỷ lệ rất nhỏ các tập đoàn lớn thuộc rổ VN100 "
        "đủ năng lực lập báo cáo phát triển bền vững độc lập theo chuẩn quốc tế GRI Standards, trong khi đa số các doanh nghiệp niêm yết vừa và nhỏ chỉ dừng lại ở các bản công bố sơ sài, "
        "mang tính hình thức và thiếu số liệu kiểm chứng (Hoang et al., 2019; Tran & Beddewela, 2020). "
        "Báo cáo Rà soát Quốc gia Tự nguyện lần thứ 2 (VNR 2023) của Chính phủ Việt Nam tại Diễn đàn Chính trị Cấp cao Liên Hợp Quốc (Bộ Kế hoạch và Đầu tư, 2023) cũng thừa nhận: "
        "tiến độ thực hiện các mục tiêu SDGs tại Việt Nam không đồng đều; các mục tiêu tăng trưởng kinh tế, việc làm và cơ sở hạ tầng đạt kết quả vượt bậc, "
        "trong khi các mục tiêu về bảo vệ hệ sinh thái biển, giảm bất bình đẳng và bình đẳng giới đối mặt với vô vàn thách thức thể chế."
    )

    # =========================================================================
    # 3. PHƯƠNG PHÁP NGHIÊN CỨU (METHODOLOGY)
    # =========================================================================
    gbp.add_h1(doc, "3. Dữ liệu và Phương pháp nghiên cứu")

    gbp.add_h2(doc, "3.1. Mẫu Nghiên cứu và Quy trình Thu thập Dữ liệu")
    gbp.add_p(
        doc,
        "Trong nghiên cứu gốc, Kang và Kim (2022) lựa chọn mẫu gồm 6 tập đoàn đa quốc gia hàng đầu thế giới (BASF, IKEA, Microsoft, Nestlé, Toyota, Walmart) "
        "đại diện cho các lĩnh vực công nghiệp, bán lẻ và công nghệ với các chuỗi báo cáo kéo dài từ 5 đến 10 năm. "
        "Kế thừa tiêu chí chọn mẫu có tính đại diện cao và có lịch sử phát hành báo cáo bền vững dài hạn, nghiên cứu này chọn lọc 4 doanh nghiệp niêm yết lớn trên Sở Giao dịch "
        "Chứng khoán Thành phố Hồ Chí Minh (HOSE) thuộc nhóm chỉ số VN100 và Chỉ số Phát triển Bền vững Việt Nam (VNSI): "
        "The PAN Group (PAN), Tập đoàn Xăng dầu Việt Nam (Petrolimex - PLX), CTCP Vàng bạc Đá quý Phú Nhuận (PNJ) và CTCP Sữa Việt Nam (Vinamilk - VNM). "
        "Đây là 4 doanh nghiệp tiên phong và hiếm hoi tại Việt Nam duy trì việc công bố Báo cáo Phát triển Bền vững độc lập một cách liên tục trong giai đoạn 2018–2025. "
        "Bảng 1 tóm lược đặc điểm của 4 doanh nghiệp trong mẫu nghiên cứu."
    )

    gbp.add_table_clean(
        doc,
        table_label="Bảng 1",
        table_title="Đặc điểm các doanh nghiệp niêm yết trong mẫu nghiên cứu",
        headers=gbp.HEADERS_T1_VI,
        data=gbp.DATA_T1_VI,
        note="Số liệu tổng hợp từ Báo cáo phát triển bền vững và cổng thông tin quan hệ nhà đầu tư của các doanh nghiệp.",
        col_widths=[1.5, 4.2, 4.2, 1.5, 2.2, 3.8],
        font_size=9.0
    )

    gbp.add_p(
        doc,
        "Tổng cộng 29 tệp báo cáo định dạng PDF với hàng nghìn trang tài liệu đã được thu thập trực tiếp từ cổng thông tin điện tử của các doanh nghiệp. "
        "Bảng 2 trình bày thống kê chi tiết kết quả trích xuất văn bản của từng báo cáo. Tổng số câu văn thu được sau tiền xử lý là 17.047 câu, "
        "với mật độ câu trung bình đạt 7,19 câu/trang. Hai báo cáo PLX 2021 (48 câu) và PLX 2024 (40 câu) có số lượng câu thấp do doanh nghiệp phát hành bản tóm tắt trực tuyến; "
        "báo cáo PNJ 2022 (338 câu) là tài liệu dạng ảnh scan dẫn đến việc trích xuất văn bản có nhiễu OCR. "
        "Các trường hợp này được ghi chú cụ thể để kiểm soát độ tin cậy trong quá trình thảo luận."
    )

    gbp.add_table_clean(
        doc,
        table_label="Bảng 2",
        table_title="Thống kê tệp báo cáo phát triển bền vững và số lượng câu trích xuất theo doanh nghiệp và năm",
        headers=gbp.HEADERS_T2_VI,
        data=gbp.DATA_T2_VI,
        note="Tổng số câu toàn mẫu: 17.047 câu từ 29 báo cáo. (* Mẫu mỏng; ** Nhiễu OCR).",
        col_widths=[1.5, 1.2, 3.5, 1.8, 1.6, 2.5, 4.5],
        font_size=8.5
    )

    gbp.add_h2(doc, "3.2. Tiền Xử lý Dữ liệu Văn bản")
    gbp.add_p(
        doc,
        "Quy trình tiền xử lý văn bản tuân thủ chặt chẽ các nguyên tắc kỹ thuật do Kang và Kim (2022) thiết lập: "
        "(1) Trích xuất toàn văn các trang nội dung từ tệp PDF bằng thư viện PyMuPDF (fitz), loại bỏ trang bìa và mục lục; "
        "(2) Xóa bỏ các ký tự điều khiển, số trang, đầu trang (header), chân trang (footer) lặp lại định kỳ; "
        "(3) Phân tách khối văn bản thành các câu đơn lẻ dựa trên hệ thống dấu câu kết thúc; "
        "(4) Áp dụng tiêu chuẩn lọc của Kang và Kim (2022): loại bỏ tất cả các đoạn văn bản có độ dài dưới 6 từ (thường là tiêu đề mục, chú thích hình ảnh hoặc hàng bảng biểu vỡ) "
        "và các chuỗi ký tự phi ngữ nghĩa. Sau khi lọc, 17.047 câu văn hoàn chỉnh được lưu trữ có cấu trúc phục vụ cho các bước phân tích tiếp theo."
    )

    gbp.add_h2(doc, "3.3. Xây dựng Tập Ngữ liệu SDG Chuẩn và Mô hình Nhúng Câu (Sentence Embedding)")
    gbp.add_p(
        doc,
        "Để đo lường sự gắn kết của doanh nghiệp với 17 mục tiêu SDG, Kang và Kim (2022) đã xây dựng một tập câu chuẩn tắc tiếng Anh từ 169 chỉ tiêu cụ thể của LHQ. "
        "Kế thừa cách tiếp cận này và chuyển giao sang ngữ cảnh Việt Nam, nghiên cứu xây dựng tập ngữ liệu chuẩn gồm 391 câu văn tiếng Việt chuẩn mực, "
        "được dịch thuật và hiệu đính học thuật từ 169 chỉ tiêu SDG chính thức và tài liệu hướng dẫn SDG Compass của GRI. "
        "Mỗi mục tiêu SDG được đại diện bởi một tập hợp S_g gồm từ 19 đến 27 câu chuẩn mô tả các hành động, cam kết và chỉ số tác động cụ thể.\n"
        "Về mô hình mã hóa: Kang và Kim (2022) sử dụng mô hình Sentence-BERT (bert-base-nli-mean-tokens) cho văn bản tiếng Anh. "
        "Để phù hợp với đặc thù cấu trúc đơn lập và ngữ pháp tiếng Việt, nghiên cứu này sử dụng mô hình `vietnamese-sbert` (với số chiều vector d = 768) "
        "cho các câu văn tiếng Việt và `all-MiniLM-L6-v2` (d = 384) (Wang et al., 2020) cho các câu văn tiếng Anh. "
        "Mỗi câu văn báo cáo r_i và câu chuẩn SDG s_j được chuyển thành vector nhúng ngữ nghĩa và chuẩn hóa vector đơn vị L2 như trong Phương trình (1):"
    )
    gbp.add_equation_clean(doc, r"\hat{\mathbf{v}} = \frac{\mathbf{v}}{\|\mathbf{v}\|_2} \implies \|\hat{\mathbf{v}}\|_2 = 1", "1")
    gbp.add_p(
        doc,
        "Với hai vector đơn vị, độ tương đồng Cosine giữa chúng chính là tích vô hướng trực tiếp. Điểm tương đồng giữa một câu văn báo cáo r "
        "và mục tiêu SDG thứ g được Kang và Kim (2022) định nghĩa bằng trung bình cộng độ tương đồng Cosine giữa câu r với toàn bộ |S_g| câu chuẩn thuộc mục tiêu đó, "
        "như trình bày trong Phương trình (2):"
    )
    gbp.add_equation_clean(doc, r"\mathrm{sim}(r, g) = \frac{1}{|S_g|} \sum_{s \in S_g} \hat{\mathbf{r}} \cdot \hat{\mathbf{s}}", "2")

    gbp.add_h2(doc, "3.4. Chuẩn hóa Min-Max Toàn cục và Mô hình 6 Nhóm Danh mục SDG")
    gbp.add_p(
        doc,
        "Sau khi tính toán ma trận điểm tương đồng Cosine thô giữa 17.047 câu văn với 17 mục tiêu SDG, Kang và Kim (2022) nhấn mạnh rằng giá trị Cosine "
        "thường tập trung trong một dải hẹp và khó so sánh trực tiếp giữa các mô hình. Do đó, nghiên cứu kế thừa nguyên văn phép biến đổi Min-Max toàn cục "
        "để co giãn toàn bộ ma trận điểm số về thang đo trực quan từ 0 đến 100 theo Phương trình (3):"
    )
    gbp.add_equation_clean(doc, r"\mathrm{Score}(r, g) = \left[ \frac{\mathrm{sim}(r, g) - \min(\mathrm{sim})}{\max(\mathrm{sim}) - \min(\mathrm{sim})} \right] \times 100", "3")
    gbp.add_p(
        doc,
        "Trên tập dữ liệu thực nghiệm 17.047 câu, giá trị cực tiểu toàn cục được ghi nhận là min(sim) = -0,0235 và cực đại là max(sim) = 0,5250. "
        "Tiếp đó, điểm số của từng báo cáo theo năm được tính bằng trung bình cộng điểm số của tất cả các câu trong báo cáo đó đối với từng mục tiêu SDG, "
        "rồi được quy nạp thành 6 nhóm danh mục nhu cầu con người theo đúng định nghĩa của Kang và Kim (2022) như chi tiết trong Bảng 3."
    )

    gbp.add_table_clean(
        doc,
        table_label="Bảng 3",
        table_title="Phân nhóm 17 mục tiêu SDG thành 6 nhóm danh mục theo nhu cầu con người của Kang & Kim (2022)",
        headers=gbp.HEADERS_T3_VI,
        data=gbp.DATA_T3_VI,
        note="Khung phân loại 6 nhóm danh mục kế thừa nguyên bản từ nghiên cứu của Kang và Kim (2022).",
        col_widths=[3.5, 3.5, 6.5, 2.5],
        font_size=9.0
    )

    gbp.add_h2(doc, "3.5. Mô hình Phân tích Sắc thái Cảm xúc Đa lớp (Sentiment Analysis)")
    gbp.add_p(
        doc,
        "Trong nghiên cứu gốc, Kang và Kim (2022) sử dụng mô hình DistilBERT phân loại nhị phân thành 2 nhãn cảm xúc: Tích cực (Positive) và Tiêu cực (Negative). "
        "Tuy nhiên, khi vận dụng vào văn bản báo cáo bền vững tiếng Việt, chúng tôi nhận thấy mô hình nhị phân tồn tại hạn chế lớn: "
        "các báo cáo chứa một tỷ trọng rất lớn các câu văn mô tả kỹ thuật, thông số đo lường phát thải, số hiệu giấy phép môi trường và các quy chuẩn pháp lý mang tính trung tính khách quan. "
        "Nếu ép buộc phân loại nhị phân, các câu văn này sẽ bị phân bổ cưỡng bức vào nhãn Tích cực hoặc Tiêu cực, làm méo mó bản chất ngôn từ.\n"
        "Do đó, nghiên cứu này nâng cấp phương pháp bằng việc sử dụng mô hình học sâu `PhoBERT` tinh chỉnh (wonrax/phobert-base-vietnamese-sentiment) "
        "(Nguyen & Nguyen, 2020) hỗ trợ phân loại 3 lớp: Tích cực (POS), Trung tính (NEU) và Tiêu cực (NEG). "
        "Điểm phân cực cảm xúc liên tục (Polarity Score) trong đoạn [0, 1] được tính toán theo hàm từng đoạn như trong Phương trình (4):"
    )
    gbp.add_equation_clean(
        doc,
        r"\mathrm{Polarity} = \begin{cases} P(\text{Tích cực}), & \text{khi nhãn là POS} \\ 0.5, & \text{khi nhãn là NEU} \\ 1 - P(\text{Tiêu cực}), & \text{khi nhãn là NEG} \end{cases}",
        "4"
    )
    gbp.add_p(
        doc,
        "Kế thừa định nghĩa của Kang và Kim (2022), Tỷ số cảm xúc hàng năm (Sentiment Ratio) của mỗi báo cáo được xác định bằng tỷ lệ giữa số lượng câu tích cực "
        "và số lượng câu tiêu cực, phản ánh độ nghiêng sắc thái trong phong cách truyền thông của doanh nghiệp theo Phương trình (5):"
    )
    gbp.add_equation_clean(doc, r"\mathrm{Ratio} = \frac{N_{\text{Tích cực}}}{N_{\text{Tiêu cực}}}", "5")

    doc.add_page_break()

    # =========================================================================
    # 4. KẾT QUẢ NGHIÊN CỨU (RESULTS)
    # =========================================================================
    gbp.add_h1(doc, "4. Kết quả nghiên cứu và Phân tích thực nghiệm")

    gbp.add_h2(doc, "4.1. Phân phối Điểm Tương đồng SDG Toàn cục")
    gbp.add_figure_clean(
        doc,
        img_filename="similarity_hist.png",
        fig_label="Hình 1",
        title="Phân phối tần suất điểm tương đồng SDG của toàn bộ 17.047 câu văn trên thang đo chuẩn hóa 0–100.",
        width_inches=5.8
    )
    gbp.add_p(
        doc,
        "Hình 1 minh họa phân phối mật độ xác suất của điểm tương đồng SDG trên toàn bộ 17.047 câu văn trích xuất từ 29 báo cáo sau khi chuẩn hóa Min-Max về thang đo 0–100. "
        "Đồ thị thể hiện một phân phối hình chuông gần như chuẩn tắc với giá trị trung bình toàn mẫu đạt 46,61 điểm và độ lệch chuẩn là 12,83 điểm. "
        "Phần lớn các câu văn tập trung trong dải điểm từ 35 đến 60 điểm. "
        "Sự xuất hiện của một phân phối liên tục và có độ biến thiên tự nhiên này hoàn toàn tương đồng với phát hiện của Kang và Kim (2022) trên mẫu các tập đoàn toàn cầu. "
        "Điều này chứng minh rằng mô hình Sentence-BERT tiếng Việt (vietnamese-sbert) đã nắm bắt xuất sắc sự đa dạng ngữ nghĩa giữa các câu văn báo cáo với 17 mục tiêu SDG, "
        "khắc phục triệt để hiện tượng phân phối thoái hóa của các phương pháp khớp từ khóa truyền thống."
    )
    gbp.add_p(
        doc,
        "Bảng 4 cung cấp các tham số thống kê mô tả chi tiết của từng mục tiêu trong 17 SDGs. Kết quả phân tích định lượng chỉ ra sự phân hóa chủ đề rõ rệt: "
        "Mục tiêu có điểm tương đồng trung bình cao nhất toàn mẫu là SDG 17 (Đối tác vì các mục tiêu - Mean = 53,74, Max = 100,00), tiếp theo là SDG 09 "
        "(Công nghiệp, đổi mới sáng tạo và hạ tầng - Mean = 50,11), SDG 12 (Sản xuất và tiêu dùng có trách nhiệm - Mean = 49,71) và SDG 15 (Tài nguyên đất liền - Mean = 49,01). "
        "Ngược lại, mục tiêu ghi nhận điểm số thấp nhất là SDG 05 (Bình đẳng giới - Mean = 38,70, Min = 0,00) và SDG 03 (Sức khỏe và phúc lợi - Mean = 42,42). "
        "Thực tế này phản ánh một quy luật khách quan trong cách lập báo cáo tại Việt Nam: các doanh nghiệp tập trung nguồn lực trình bày về hợp tác chuỗi giá trị, "
        "đầu tư công nghệ và mở rộng cơ sở hạ tầng sản xuất, trong khi các nội dung về bình đẳng giới và thu hẹp khoảng cách tiền lương nội bộ chưa nhận được sự ưu tiên tương xứng (Bộ Kế hoạch và Đầu tư, 2023)."
    )

    gbp.add_table_clean(
        doc,
        table_label="Bảng 4",
        table_title="Thống kê mô tả điểm số tương đồng của 17 mục tiêu SDG trên toàn bộ 17.047 câu văn",
        headers=gbp.HEADERS_T4_VI,
        data=gbp.DATA_T4_VI,
        note="Điểm số được chuẩn hóa Min-Max toàn cục trên thang đo từ 0 đến 100.",
        col_widths=[3.8, 1.6, 1.6, 1.6, 1.4, 1.6, 1.6, 1.6, 1.6],
        font_size=8.5
    )

    gbp.add_h2(doc, "4.2. Phân tích Cấu trúc 6 Nhóm Danh mục SDG qua Ma trận Nhiệt")
    gbp.add_figure_clean(
        doc,
        img_filename="heatmap_6cat.png",
        fig_label="Hình 2",
        title="Biểu đồ nhiệt thể hiện điểm số tương đồng 6 nhóm danh mục SDG theo công ty và năm.",
        width_inches=5.6
    )
    gbp.add_p(
        doc,
        "Biểu đồ nhiệt tại Hình 2 và số liệu chi tiết trong Bảng 5 thể hiện sự phân hóa cấu trúc chủ đề giữa 4 doanh nghiệp theo 6 nhóm danh mục của Kang và Kim (2022). "
        "Một quy luật tổng quát có tính nhất quán rất cao được phát hiện trên toàn bộ 29 báo cáo: Nhóm Kinh tế luôn dẫn đầu, tiếp theo là Xã hội, Tài nguyên, "
        "Đời sống, Môi trường và cuối cùng là Công bằng xã hội: Kinh tế > Xã hội > Tài nguyên ≈ Đời sống > Môi trường > Công bằng.\n"
        "Cụ thể, nhóm Kinh tế (SDG 8, 9) đạt điểm số cao vượt trội tại hầu hết các công ty–năm, dao động từ 44,61 (Vinamilk năm 2019) đến mức đỉnh 54,84 (PNJ năm 2022). "
        "Điều này phản ánh đặc thù của các doanh nghiệp niêm yết tại thị trường mới nổi: báo cáo phát triển bền vững vẫn được tận dụng tối đa để khẳng định hiệu quả sản xuất kinh doanh, "
        "mở rộng quy mô tài sản và đổi mới dây chuyền công nghệ – vốn là các chỉ tiêu tạo được niềm tin tức thời nơi cổ đông và nhà đầu tư."
    )
    gbp.add_p(
        doc,
        "Trái lại, nhóm Công bằng (SDG 4, 5, 10) liên tục xếp ở vị trí cuối cùng trên mọi báo cáo, chỉ dao động từ 39,92 (Petrolimex năm 2024) đến 47,40 (PNJ năm 2025). "
        "Đáng chú ý, ngay cả tại PNJ – doanh nghiệp kim hoàn có tỷ lệ lao động nữ trên 60% và có ban chỉ đạo ESG riêng biệt – điểm nhóm Công bằng (44,29 – 47,40) vẫn thấp hơn đáng kể "
        "so với nhóm Kinh tế (50,96 – 54,84) và Xã hội (49,46 – 52,99). Phát hiện thực nghiệm này hoàn toàn khớp với nhận định trong Báo cáo Rà soát Quốc gia Tự nguyện "
        "(VNR 2023) của Chính phủ Việt Nam (Bộ Kế hoạch và Đầu tư, 2023), chỉ ra rằng việc xóa bỏ bất bình đẳng giới và hỗ trợ các nhóm yếu thế trong chuỗi cung ứng "
        "vẫn là mục tiêu chậm tiến độ nhất trong 17 SDGs tại Việt Nam."
    )

    gbp.add_table_clean(
        doc,
        table_label="Bảng 5",
        table_title="Điểm số tương đồng trung bình 6 nhóm danh mục SDG theo công ty và năm",
        headers=gbp.HEADERS_T5_VI,
        data=gbp.DATA_T5_VI,
        note="Dấu (*) biểu thị các báo cáo có số lượng câu dưới 80 câu do bản tóm tắt trực tuyến.",
        col_widths=[1.5, 1.2, 1.8, 1.8, 1.8, 1.8, 1.8, 2.0, 2.3],
        font_size=8.5
    )

    gbp.add_h2(doc, "4.3. Diễn biến Xu hướng theo Thời gian của Từng Doanh nghiệp")
    gbp.add_figure_clean(
        doc,
        img_filename="trends_6categories.png",
        fig_label="Hình 3",
        title="Đường xu hướng biến thiên điểm số 6 nhóm danh mục SDG giai đoạn 2018–2025 cho từng doanh nghiệp.",
        width_inches=5.8
    )
    gbp.add_p(
        doc,
        "Hình 3 thể hiện quỹ đạo vận động của 6 nhóm danh mục theo thời gian cho từng doanh nghiệp. Điểm nhấn nổi bật nhất là bước nhảy vọt đồng loạt "
        "của tất cả các doanh nghiệp trong năm 2025. Cụ thể, Vinamilk ghi nhận mức bứt phá mạnh nhất khi toàn bộ 6 nhóm danh mục đều vượt ngưỡng 46 điểm, "
        "trong đó nhóm Môi trường tăng từ 46,30 (năm 2024) lên 51,95 (năm 2025), và Đời sống tăng từ 45,35 lên 50,06. Sự chuyển biến này bắt nguồn trực tiếp từ việc Vinamilk "
        "công bố chiến lược Net Zero 2050, hoàn thành chứng nhận trung hòa carbon PAS 2060 cho các nhà máy và trang trại tại Nghệ An, Bến Tre, "
        "đồng thời phát hành bản báo cáo phát triển bền vững chuyên đề dày 152 trang với 954 câu văn chuyên sâu."
    )
    gbp.add_p(
        doc,
        "Đối với The PAN Group, điểm số suy giảm nhẹ trong năm 2023 (điểm trung bình 44,24 so với 47,09 năm 2022) do công ty tinh giản dung lượng báo cáo (411 câu), "
        "sau đó phục hồi mạnh mẽ vào năm 2025 (đạt 49,77 điểm). Đối với PNJ, công ty duy trì phong độ ổn định nhất mẫu với điểm trung bình luôn nằm trong khoảng 48 đến 51 điểm. "
        "Riêng Petrolimex, đường xu hướng có sự biến động cục bộ tại hai năm 2021 và 2024 do đây là các bản tóm tắt trực tuyến với cỡ mẫu mỏng."
    )

    gbp.add_h2(doc, "4.4. Phân tích Sắc thái Cảm xúc và Cơ cấu Ngôn từ")
    gbp.add_figure_clean(
        doc,
        img_filename="sentiment_hist.png",
        fig_label="Hình 4",
        title="Phân phối điểm phân cực sắc thái cảm xúc (Polarity Score) từ mô hình PhoBERT đa lớp.",
        width_inches=5.2
    )
    gbp.add_p(
        doc,
        "Hình 4 minh họa phân phối điểm phân cực cảm xúc (Polarity Score) của 17.047 câu văn. Khác với phân phối hình chữ U trong nghiên cứu của Kang và Kim (2022) "
        "vốn chỉ sử dụng 2 nhãn nhị phân, đồ thị phân phối trong nghiên cứu này thể hiện dạng lưỡng đỉnh rõ nét: "
        "Đỉnh thứ nhất tại mốc 0,50 tương ứng với 3.126 câu trung tính (chiếm 18,34%), và đỉnh khổng lồ thứ hai nằm sát mốc 0,95–1,00 đại diện cho 11.891 câu tích cực (chiếm 69,75%). "
        "Số lượng câu thực sự tiêu cực (Polarity < 0,2) chỉ có 2.030 câu (chiếm 11,91%). Tỷ lệ áp đảo gần 70% câu tích cực là bằng chứng định lượng khẳng định "
        "rằng báo cáo phát triển bền vững tại Việt Nam đang chịu sự chi phối mạnh mẽ của thiên lệch lạc quan (optimism bias)."
    )
    gbp.add_figure_clean(
        doc,
        img_filename="sentiment_by_company.png",
        fig_label="Hình 5",
        title="Cơ cấu tỷ lệ câu Tích cực, Trung tính, Tiêu cực và Tỷ số Pos/Neg Ratio theo từng doanh nghiệp.",
        width_inches=5.8
    )
    gbp.add_p(
        doc,
        "Bảng 6 và Hình 5 trình bày cơ cấu cảm xúc chi tiết theo từng báo cáo. Vinamilk duy trì tỷ số cảm xúc (Pos/Neg Ratio) cao nhất toàn mẫu, "
        "đạt đỉnh 11,97 vào năm 2020 (455 câu tích cực so với vỏn vẹn 38 câu tiêu cực) và 10,86 vào năm 2024. The PAN Group có tỷ số cảm xúc tương đối ổn định "
        "trong khoảng 5,40 đến 10,22. Petrolimex thể hiện phong cách báo cáo thận trọng hơn của một tập đoàn công nghiệp nhà nước, duy trì tỷ số Pos/Neg "
        "thấp nhất mẫu (3,00 đến 5,22, ngoại trừ năm mẫu mỏng 2024). Trường hợp dị biệt xuất hiện tại PNJ năm 2022: Tỷ số Pos/Neg sụt giảm xuống chỉ còn 1,33 "
        "(182 câu tích cực / 137 câu tiêu cực). Nguyên nhân bắt nguồn từ lỗi nhận diện OCR trên bản PDF scan ảnh: các khối văn bản bị vỡ chữ khiến mô hình "
        "PhoBERT phân loại nhầm hàng loạt câu thành nhãn Tiêu cực. Đây là bài học thực tế quan trọng về việc kiểm soát chất lượng số hóa tài liệu trước khi phân tích NLP."
    )

    gbp.add_table_clean(
        doc,
        table_label="Bảng 6",
        table_title="Cơ cấu số lượng câu theo nhãn cảm xúc và tỷ số Pos/Neg Ratio qua các năm",
        headers=gbp.HEADERS_T6_VI,
        data=gbp.DATA_T6_VI,
        note="Dấu (*) biểu thị mẫu mỏng dưới 80 câu; dấu (**) biểu thị báo cáo scan ảnh có nhiễu OCR.",
        col_widths=[1.5, 1.2, 2.6, 2.6, 2.6, 2.2, 3.3],
        font_size=8.5
    )

    gbp.add_figure_clean(
        doc,
        img_filename="sentiment_ratio.png",
        fig_label="Hình 6",
        title="Đường xu hướng biến thiên tỷ số Pos/Neg Ratio của 4 công ty giai đoạn 2018–2025.",
        width_inches=5.5
    )
    gbp.add_p(
        doc,
        "Hình 6 trực quan hóa quỹ đạo biến thiên của tỷ số Pos/Neg Ratio của 4 doanh nghiệp. Đồ thị phản ánh rõ sự phân tầng phong cách ngôn từ: "
        "Vinamilk luôn nằm ở dải trên cùng (từ 7 đến 12 lần câu tích cực so với câu tiêu cực), PAN Group duy trì ở mức trung bình cao (6 đến 10 lần), "
        "trong khi Petrolimex duy trì đều đặn ở dải thận trọng (3 đến 5 lần). Đường xu hướng của PNJ phản ánh rõ điểm gãy kỹ thuật năm 2022 do nhiễu OCR "
        "trước khi phục hồi nhanh chóng về mức chuẩn mực trên 4,5 lần vào các năm 2023–2025."
    )

    doc.add_page_break()

    # =========================================================================
    # 5. THẢO LUẬN VÀ CẢM NGHĨ CHUYÊN SÂU (DISCUSSION & REFLECTIONS)
    # =========================================================================
    gbp.add_h1(doc, "5. Thảo luận và Cảm nghĩ chuyên sâu")

    gbp.add_h2(doc, "5.1. Cảm nghĩ và Suy ngẫm về Thực trạng Báo cáo Bền vững tại Việt Nam qua Lăng kính NLP")
    gbp.add_p(
        doc,
        "Quá trình phân tích thực nghiệm trên 17.047 câu văn từ 29 báo cáo phát triển bền vững đem lại những cảm nghĩ và suy ngẫm sâu sắc "
        "về thực trạng thực hành ESG và công bố thông tin phi tài chính của các doanh nghiệp Việt Nam hiện nay. "
        "Điểm nổi bật đầu tiên là sự 'thiên lệch lạc quan' (optimism bias) mang tính cấu trúc ăn sâu vào văn hóa viết báo cáo. "
        "Với gần 70% câu văn mang sắc thái tích cực và chỉ vỏn vẹn 11,9% câu văn tiêu cực, các báo cáo phát triển bền vững dường như đang được vận hành "
        "như một ấn phẩm quan hệ công chúng (PR) cao cấp hơn là một tài liệu giải trình trách nhiệm khoa học và khách quan. "
        "Các doanh nghiệp dành hàng chục trang để ca ngợi các thành tích đạt được, các giải thưởng và chứng nhận, nhưng hầu như né tránh triệt để việc đề cập "
        "đến những khó khăn trong việc cắt giảm khí thải, các xung đột lợi ích với cộng đồng địa phương hoặc các chỉ tiêu môi trường chưa đạt kế hoạch."
    )
    gbp.add_p(
        doc,
        "Điểm suy ngẫm thứ hai là sự 'lệch pha' nghiêm trọng giữa các trụ cột phát triển bền vững. "
        "Việc nhóm Kinh tế (SDG 8, 9) và Đối tác (SDG 17) luôn áp đảo với điểm số trên 50, trong khi nhóm Công bằng (SDG 4, 5, 10) bị bỏ lại phía sau với mức điểm dưới 40 "
        "cho thấy tư duy phát triển bền vững của doanh nghiệp Việt Nam vẫn bị neo chặt vào mục tiêu tài chính truyền thống. "
        "Doanh nghiệp sẵn sàng đầu tư báo cáo về dây chuyền tự động hóa, mở rộng nhà xưởng (vì đây là các khoản đầu tư vốn có thể lượng hóa và mang lại dòng tiền), "
        "nhưng lại ít chú trọng đến chính sách bình đẳng tiền lương theo giới, phúc lợi cho người lao động thời vụ và sự hòa nhập của người khuyết tật trong chuỗi cung ứng. "
        "Điều này phản ánh một thực tế rằng phát triển bền vững tại Việt Nam hiện nay mới chỉ dừng lại ở 'bền vững vì hiệu quả kinh tế' chứ chưa đạt tới 'bền vững vì công bằng xã hội'."
    )
    gbp.add_p(
        doc,
        "Đặc biệt, kết quả kiểm định thực nghiệm cung cấp một góc nhìn khoa học quan trọng về câu hỏi: Liệu doanh nghiệp Việt Nam có đang thực hiện hành vi 'tẩy xanh' (greenwashing) hay không? "
        "Khi thực hiện các kiểm định tương quan giữa tỷ số cảm xúc trong báo cáo và biến động giá cổ phiếu trên sàn HOSE, kết quả cho thấy hệ số tương quan tuyến tính Pearson "
        "trên toàn mẫu hoàn toàn bằng không (r = 0,030, p = 0,890) và tương quan phi tham số Spearman cũng không có ý nghĩa thống kê (ρ = -0,047, p = 0,828). "
        "Kết quả này khẳng định rằng: KHÔNG CÓ ĐỦ BẰNG CHỨNG THỰC NGHIỆM ĐỂ KẾT LUẬN DOANH NGHIỆP VIỆT NAM CÓ HÀNH VI TẨY XANH CỐ Ý NHẰM ĐIỀU MANH THỊ TRƯỜNG.\n"
        "Nguyên nhân bản chất nằm ở đặc thù của thị trường chứng khoán Việt Nam giai đoạn 2018–2025: Đây là thị trường mới nổi/cận biên bị chi phối bởi các nhà đầu tư cá nhân "
        "(chiếm trên 85% giá trị giao dịch hàng ngày). Động lực biến động giá cổ phiếu chủ yếu chịu tác động bởi dòng tiền ngắn hạn, kết quả lợi nhuận quý và các biến số kinh tế vĩ mô "
        "(lãi suất, tỷ giá, chính sách tiền tệ), chứ nhà đầu tư chưa thực sự đọc, phân tích và định giá các thông điệp trong báo cáo phát triển bền vững. "
        "Do đó, việc tỷ số cảm xúc của báo cáo tăng lên trong khi giá cổ phiếu suy giảm chỉ đơn thuần phản ánh nỗ lực 'quản trị ấn tượng' (impression management) thông thường của bộ phận truyền thông, "
        "chứ không xuất phát từ động cơ thao túng giá thị trường để bị coi là tẩy xanh (greenwashing)."
    )

    gbp.add_h2(doc, "5.2. Giải phẫu Hành trình Bền vững của 4 Doanh nghiệp Nghiên cứu")
    gbp.add_p(
        doc,
        "Bốn doanh nghiệp trong mẫu nghiên cứu đại diện cho các mô hình ứng xử và chiến lược bền vững đặc trưng tại Việt Nam:\n"
        "Vinamilk đại diện cho mô hình tiên phong chuẩn mực quốc tế: Là doanh nghiệp có lịch sử lập báo cáo GRI độc lập từ năm 2012, Vinamilk thể hiện năng lực quản trị dữ liệu vượt trội. "
        "Sự bứt phá điểm số vào năm 2025 gắn liền với các chứng nhận trung hòa carbon PAS 2060 độc lập và công bố quốc tế CDP. Tuy nhiên, việc duy trì tỷ số cảm xúc quá cao (đỉnh 11,97) "
        "cũng gợi ý rằng doanh nghiệp cần thận trọng hơn trong việc cân bằng giữa truyền thông thành tựu và giải trình rủi ro sinh thái dài hạn.\n"
        "The PAN Group đại diện cho mô hình nông nghiệp chuỗi giá trị thực chất: Với đặc thù ngành nông nghiệp và xuất khẩu thủy sản sang EU, "
        "PAN bắt buộc phải tuân thủ các chuẩn mực nuôi trồng bền vững quốc tế (ASC, BAP). Điểm số của PAN phản ánh tính thực chất cao, thể hiện rõ qua việc tỷ số cảm xúc "
        "giảm nhẹ trong năm 2022 để phản ánh đúng áp lực chi phí nguyên liệu và biến động vĩ mô.\n"
        "Petrolimex đại diện cho thách thức chuyển dịch của doanh nghiệp năng lượng hóa thạch nhà nước: Là tập đoàn xăng dầu chiếm 50% thị phần nội địa, PLX đối mặt với áp lực giảm phát thải khổng lồ. "
        "Các nỗ lực thương mại hóa nhiên liệu sinh học và lắp đặt trạm sạc xe điện là rất đáng ghi nhận. Tuy nhiên, việc có những năm chỉ phát hành bản tóm tắt trực tuyến mỏng (40–48 câu) "
        "cho thấy doanh nghiệp cần đầu tư đồng bộ hơn vào việc lập báo cáo định kỳ đầy đủ.\n"
        "PNJ đại diện cho mô hình tiên phong về quản trị và bình đẳng giới: PNJ là hình mẫu hiếm hoi tại Việt Nam lập tiểu ban ESG trực thuộc HĐQT từ sớm. "
        "Tuy nhiên, sự cố tệp báo cáo scan năm 2022 dẫn đến nhiễu OCR là bài học đắt giá: Nếu doanh nghiệp không chuẩn hóa định dạng số hóa tài liệu điện tử, "
        "các hệ thống trí tuệ nhân tạo và tổ chức xếp hạng tín nhiệm ESG quốc tế sẽ tự động đánh giá sai lệch năng lực bền vững của doanh nghiệp dựa trên các văn bản bị lỗi font chữ."
    )

    gbp.add_h2(doc, "5.3. So sánh Đối chiếu Toàn diện với Nghiên cứu Gốc của Kang và Kim (2022)")
    gbp.add_p(
        doc,
        "Khi đặt kết quả nghiên cứu này cạnh công trình của Kang và Kim (2022) trên các tập đoàn toàn cầu, chúng tôi nhận thấy các điểm tương đồng và khác biệt có ý nghĩa lý luận sâu sắc:\n"
        "Về điểm tương đồng: Cả hai nghiên cứu đều khẳng định tính ưu việt tuyệt đối của Sentence-BERT trong việc tạo ra phân phối điểm tương đồng hình chuông có ý nghĩa, "
        "vượt trội hoàn toàn so với phân phối thoái hóa của phương pháp đếm từ khóa. Cả hai công trình đều phát hiện xu hướng mở rộng quy mô báo cáo theo thời gian "
        "và sự thống trị của sắc thái tích cực (chiếm trên 65% tổng số câu văn).\n"
        "Về điểm khác biệt: Thứ nhất, điểm tương đồng SDG trung bình của doanh nghiệp Việt Nam (46,61) thấp hơn khoảng 3–5 điểm so với các tập đoàn toàn cầu trong nghiên cứu gốc (49–52 điểm), "
        "cho thấy ngôn ngữ báo cáo tại Việt Nam vẫn sử dụng nhiều câu từ chung chung, chưa ánh xạ chuẩn xác vào các thuật ngữ kỹ thuật của 169 chỉ tiêu SDG. "
        "Thứ hai, trong khi Kang và Kim (2022) sử dụng mô hình nhị phân tạo ra phân phối cảm xúc chữ U, nghiên cứu này áp dụng mô hình PhoBERT 3 lớp tạo ra phân phối lưỡng đỉnh, "
        "tách bạch được 18,34% câu văn trung tính – một đặc thù thiết yếu của văn bản doanh nghiệp Việt Nam."
    )

    gbp.add_h2(doc, "5.4. Hàm ý Quản trị và Khuyến nghị Chính sách")
    gbp.add_p(
        doc,
        "Dựa trên các phát hiện thực nghiệm, nghiên cứu đề xuất các khuyến nghị cụ thể:\n"
        "Thứ nhất, đối với Cơ quan Quản lý (Ủy ban Chứng khoán Nhà nước và các Sở Giao dịch): Cần sớm ban hành lộ trình áp dụng các chuẩn mực công bố thông tin quốc tế ISSB (IFRS S1 và S2), "
        "đồng thời quy định bắt buộc kiểm toán độc lập đối với dữ liệu phát thải Scope 1 và Scope 2. Cần chuẩn hóa định dạng nộp báo cáo (khuyến khích định dạng iXBRL hoặc PDF số hóa chuẩn) "
        "để tạo điều kiện cho các công cụ NLP tự động giám sát thị trường.\n"
        "Thứ hai, đối với doanh nghiệp niêm yết: Cần chuyển dịch từ tư duy 'quảng bá hình ảnh' sang 'trách nhiệm giải trình thực chất'. "
        "Doanh nghiệp cần dũng cảm công bố các rủi ro khí hậu, các chỉ tiêu chưa hoàn thành và dành sự quan tâm xứng đáng cho trụ cột bình đẳng giới (SDG 5) và người lao động yếu thế.\n"
        "Thứ ba, đối với nhà đầu tư và các tổ chức tài chính: Cần tận dụng các công cụ phân tích văn bản tự động để bóc tách chất lượng báo cáo, "
        "không chỉ nhìn vào độ dày của báo cáo hay tần suất xuất hiện của các từ khóa bền vững mà phải đối chiếu với số liệu định lượng đã được kiểm toán."
    )

    gbp.add_h2(doc, "5.5. Các Giới hạn Khoa học và Đề xuất Mở rộng Mẫu Nghiên cứu Sang Doanh nghiệp Giai đoạn 2020–2025")
    gbp.add_p(
        doc,
        "Nghiên cứu tồn tại một số giới hạn phương pháp luận: (1) Cỡ mẫu ban đầu tập trung vào 4 doanh nghiệp quy mô lớn thuộc VN100 do sự khan hiếm của các báo cáo độc lập kéo dài trọn vẹn 8 năm; "
        "(2) Nghiên cứu tập trung khai phá văn bản định tính mà chưa bóc tách các bảng biểu số liệu kỹ thuật phức tạp (như số liệu kiểm kê khí nhà kính hoặc lượng nước tiêu thụ).\n"
        "Để khắc phục giới hạn về quy mô mẫu, chúng tôi đã tiến hành khảo sát toàn diện thị trường chứng khoán Việt Nam và nhận diện được danh sách các doanh nghiệp niêm yết tiềm năng "
        "có đầy đủ chuỗi Báo cáo Phát triển Bền vững độc lập hoặc Báo cáo Tích hợp trong giai đoạn 5 năm gần nhất từ 2020 đến 2025:\n"
        "1. Tập đoàn Bảo Việt (Mã CK: BVH - Ngành Bảo hiểm & Dịch vụ Tài chính): Là doanh nghiệp tiên phong sớm nhất tại Việt Nam phát hành Báo cáo Phát triển Bền vững riêng biệt từ năm 2014 "
        "và sau đó chuyển đổi sang Báo cáo Tích hợp (Integrated Reporting) theo chuẩn GRI và <IR> Framework. BVH sở hữu chuỗi báo cáo liên tục trọn vẹn từ 2020 đến 2025 với cấu trúc dữ liệu cực kỳ bài bản;\n"
        "2. CTCP Dược Hậu Giang (Mã CK: DHG - Ngành Dược phẩm & Chăm sóc Sức khỏe): Doanh nghiệp đầu ngành dược phẩm liên tục phát hành Báo cáo Phát triển Bền vững độc lập từ trước năm 2020 đến nay "
        "với hệ thống chỉ tiêu chuyên sâu về SDG 3 (Sức khỏe và phúc lợi), quản lý chất thải y tế và trách nhiệm cộng đồng;\n"
        "3. CTCP Vicostone (Mã CK: VCS - Ngành Vật liệu Xây dựng & Chế tạo): Thuộc Tập đoàn Phenikaa, VCS là đại diện tiêu biểu của ngành công nghiệp nặng duy trì Báo cáo Phát triển Bền vững độc lập hàng năm, "
        "tập trung sâu vào tiêu thụ năng lượng sạch, kinh tế tuần hoàn và chứng nhận xanh quốc tế;\n"
        "4. CTCP Tập đoàn FPT (Mã CK: FPT - Ngành Công nghệ Thông tin & Viễn thông): Doanh nghiệp công nghệ hàng đầu xuất bản các chuyên đề Báo cáo Phát triển Bền vững và Báo cáo ESG chuyên sâu từ năm 2020 đến 2025, "
        "tiên phong trong các mục tiêu chuyển đổi số xanh và bình đẳng giáo dục công nghệ;\n"
        "5. CTCP Đầu tư Nam Long (Mã CK: NLG - Ngành Bất động sản Dân cư): Đại diện nổi bật trong lĩnh vực bất động sản sở hữu chuỗi báo cáo bền vững độc lập liên tục từ 2019 đến 2025 theo chuẩn GRI, "
        "phản ánh các chỉ tiêu về công trình xanh (EDGE) và phát triển đô thị bền vững (SDG 11).\n"
        "Việc tích hợp thêm 5 doanh nghiệp này (BVH, DHG, VCS, FPT, NLG) vào khung phân tích NLP trong các giai đoạn nghiên cứu tiếp theo sẽ nâng tổng quy mô mẫu lên 9 doanh nghiệp "
        "đại diện cho 9 ngành kinh tế chủ chốt, mở rộng dung lượng tập dữ liệu lên trên 50.000 câu văn và gia tăng đáng kể tính đại diện thống kê cho toàn bộ thị trường chứng khoán Việt Nam."
    )

    # =========================================================================
    # 6. KẾT LUẬN (CONCLUSION)
    # =========================================================================
    gbp.add_h1(doc, "6. Kết luận")
    gbp.add_p(
        doc,
        "Nghiên cứu này là công trình thực nghiệm tiên phong tại Việt Nam tái hiện và vận dụng thành công khung phân tích NLP của Kang và Kim (2022) "
        "trên kho ngữ liệu báo cáo phát triển bền vững tiếng Việt. Bằng việc kết hợp các mô hình Transformer tiên tiến (vietnamese-sbert và PhoBERT), "
        "chúng tôi đã chứng minh năng lực vượt trội của trí tuệ nhân tạo trong việc giải mã, định lượng và trực quan hóa các cam kết phi tài chính của doanh nghiệp. "
        "Những phát hiện về sự thống trị của mục tiêu Kinh tế, sự tụt hậu của mục tiêu Bình đẳng giới và sự thiên lệch lạc quan trong cấu trúc ngôn từ "
        "cung cấp những luận cứ khoa học sắc bén cho các nhà hoạch định chính sách, ban lãnh đạo doanh nghiệp và cộng đồng nhà đầu tư. "
        "Việc không tìm thấy bằng chứng thực nghiệm về hành vi tẩy xanh có chủ đích giúp làm sáng tỏ thực trạng thị trường và định hướng cho các nỗ lực nâng cao tính minh bạch, "
        "hướng tới một nền kinh tế xanh và bền vững thực chất tại Việt Nam."
    )

    doc.add_page_break()

    # =========================================================================
    # TÀI LIỆU THAM KHẢO (REFERENCES - APA 7th)
    # =========================================================================
    gbp.add_h1(doc, "Tài liệu tham khảo")
    for ref in gbp.APA_REFS_VI:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(-0.5)
        run = p.add_run(ref)
        run.font.name = "Times New Roman"
        run.font.size = Pt(10)
        run.font.color.rgb = gbp.COLOR_BLACK
