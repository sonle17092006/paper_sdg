"""Mô-đun xây dựng toàn văn Bài báo tiếng Anh hoàn chỉnh chuẩn APA 7th.
Bao gồm:
- Tiêu đề, Tác giả, Abstract chuyên sâu
- 6 Phần nội dung học thuật tiếng Anh đầy đủ (Introduction, Literature Review, Methodology, Results, Discussion, Conclusion)
- 8 Hình ảnh (Figures)
- 6 Bảng số liệu (Tables) tiếng Anh
- 6 Công thức toán học OMML
- Hơn 40 tài liệu tham khảo APA 7th
"""

from __future__ import annotations

import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm, RGBColor

import paper_base as pb


def build_english_paper(doc: Document, tables_data: dict, is_endnote_ready: bool = False):
    pb.setup_clean_styles(doc)

    def cite(author: str, year: str, paren: bool = True, alt_text: str = "") -> str:
        if is_endnote_ready:
            return f"{{{author}, {year}}}"
        if alt_text:
            return alt_text
        if paren:
            return f"({author}, {year})"
        return f"{author} ({year})"

    # =========================================================================
    # TITLE & AUTHOR INFORMATION
    # =========================================================================
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before = Pt(14)
    tp.paragraph_format.space_after = Pt(12)
    run = tp.add_run(
        "MULTILINGUAL NLP FOR SDG AND SENTIMENT ANALYSIS OF CORPORATE SUSTAINABILITY REPORTS: "
        "EVIDENCE FROM VIETNAMESE ENTERPRISES"
    )
    run.font.name = "Times New Roman"
    run.font.size = Pt(15)
    run.bold = True
    run.font.color.rgb = pb.COLOR_BLACK

    ap = doc.add_paragraph()
    ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ap.paragraph_format.space_after = Pt(18)
    arun = ap.add_run("Le Dan Son, Duong Thi Hoan")
    arun.font.name = "Times New Roman"
    arun.font.size = Pt(12)
    arun.bold = True
    arun.font.color.rgb = pb.COLOR_BLACK

    # =========================================================================
    # ABSTRACT & KEYWORDS
    # =========================================================================
    pb.add_h1(doc, "Abstract")
    c_kang = cite("Kang", "2022", alt_text="Kang and Kim (2022)", paren=False)
    pb.add_p(
        doc,
        f"Corporate sustainability reports serve as fundamental communication channels for conveying environmental, social, and governance (ESG) "
        f"commitments to global investors, stakeholders, and regulatory authorities. Replicating and substantially extending the computational "
        f"framework pioneered by {c_kang}, this study conducts an automated textual analysis and longitudinal visualization of 42 standalone sustainability "
        f"and integrated reports published by seven prominent Vietnamese listed corporations (Bao Viet Holdings, The PAN Group, Petrolimex, "
        f"Phu Nhuan Jewelry, SSI Securities, Vicostone, and Vinamilk) spanning a balanced six-year observation period (2020–2025), encompassing "
        f"96,461 digitally extracted sentences (augmented with dedicated optical character recognition OCR recovery for scanned reports). Employing domain-adapted Transformer architectures (vietnamese-sbert and PhoBERT), we quantify "
        f"corporate alignment with the 17 UN Sustainable Development Goals (SDGs), aggregate alignments into six essential human-needs categories "
        f"(Life, Economic, Equity, Social, Resources, Environments), evaluate multi-class sentiment distributions (Positive, Neutral, Negative), "
        f"and formulate a continuous Greenwashing Index (GW) benchmarked against monthly equity market dynamics. "
        f"Empirical findings demonstrate that: (1) Disclosures are characterized by an overwhelming economic dominance (SDGs 8, 9) and partnership focus (SDG 17), "
        f"while systemic neglect persists in Equity disclosures (SDGs 4, 5, 10); (2) Corporate texts exhibit pronounced structural optimism bias "
        f"(53.87% positive, 32.87% neutral, and only 13.26% negative sentences); (3) Longitudinal market tests reveal complete decoupling between narrative corporate tone "
        f"and annual stock returns (pooled Pearson r = -0.043, p = 0.806), reflecting pervasive impression management in an emerging market where non-financial data "
        f"is not yet efficiently priced; and (4) The continuous GW index effectively isolates market penalty periods juxtaposed against corporate tone inflation, "
        f"while underscoring critical methodological lessons regarding PDF digitization hygiene and OCR remediation of scanned reporting artifacts. Crucial policy recommendations are proposed "
        f"to standardize digital sustainability reporting infrastructure in Vietnam."
    )
    pb.add_p(
        doc,
        "Keywords: Corporate sustainability reports; Sustainable Development Goals (SDGs); Natural language processing (NLP); "
        "Sentence-BERT; PhoBERT; Sentiment analysis; Greenwashing; Vietnam stock market.",
        bold=True
    )

    # =========================================================================
    # 1. INTRODUCTION
    # =========================================================================
    pb.add_h1(doc, "1. Introduction")
    c_un = cite("United Nations", "2015")
    c_gri = cite("Global Reporting Initiative", "2021")
    c_btc = cite("Ministry of Finance of Vietnam", "2020")
    c_ssc = cite("State Securities Commission of Vietnam", "2024")
    c_vbcsd = cite("VBCSD", "2024")

    pb.add_p(
        doc,
        f"Against the backdrop of intensifying global climate challenges and the accelerating transition toward circular economic models, "
        f"transparent disclosure of Environmental, Social, and Governance (ESG) performance and alignment with the United Nations Sustainable Development Goals "
        f"{c_un} have evolved from voluntary philanthropic gestures into mandatory regulatory imperatives and strategic competitive benchmarks {c_gri}. "
        f"In Vietnam, the government's ambitious commitment at COP26 to achieve net-zero carbon emissions by 2050, coupled with Circular No. 96/2020/TT-BTC "
        f"promulgated by the Ministry of Finance {c_btc}, has made sustainability reporting legally binding for public companies. Furthermore, the deployment "
        f"of the Corporate Sustainability Index (CSI) by the Vietnam Chamber of Commerce and Industry {c_vbcsd} and ESG implementation guidelines issued by "
        f"the State Securities Commission {c_ssc} have motivated leading listed enterprises to transition from superficial annual report sections toward comprehensive, "
        f"standalone sustainability reports and integrated reports adhering to the Global Reporting Initiative (GRI) Universal Standards."
    )

    c_luccioni = cite("Luccioni et al.", "2020")
    c_arvidsson = cite("Arvidsson & Dumay", "2022")
    c_merkl = cite("Merkl-Davies & Brennan", "2007")
    c_lyon = cite("Lyon & Montgomery", "2015")
    pb.add_p(
        doc,
        f"Nevertheless, the dramatic surge in the volume, length, and visual complexity of unstructured corporate disclosures—frequently exceeding hundreds "
        f"of pages per publication—presents an unprecedented cognitive burden for institutional investors, rating agencies, and market regulators {c_arvidsson}. "
        f"Traditional manual content analysis suffers from severe structural bottlenecks: it is prohibitively labor-intensive, unscalable across multi-year cross-sectional "
        f"datasets, and inherently prone to subjective coder biases. More critically, the predominantly self-laudatory nature of unverified corporate disclosures "
        f"frequently degenerates into strategic impression management {c_merkl}. Corporations are incentivized to cherry-pick accomplishments, deploy hyperbolic rhetoric, "
        f"and obscure environmental violations or unmet reduction targets, thereby engaging in greenwashing that misallocates sustainable capital and erodes investor confidence {c_lyon}."
    )

    c_reimers = cite("Reimers & Gurevych", "2019")
    c_nguyen = cite("Nguyen & Nguyen", "2020")
    pb.add_p(
        doc,
        f"To resolve this analytical impasse, the deployment of Artificial Intelligence (AI) and Natural Language Processing (NLP) powered by Transformer "
        f"deep learning models offers a transformative auditing paradigm {c_luccioni}. A seminal contribution in this domain is the work of {c_kang}, "
        f"who developed a Sentence-BERT {c_reimers} and sentiment analysis framework to map 24 corporate sustainability reports from six major South Korean chaebols "
        f"onto the 17 UN SDGs. While Kang and Kim established a rigorous methodological precedent, adapting their computational pipeline to an emerging Southeast Asian economy "
        f"like Vietnam encounters substantial domain-specific hurdles: the complex tonal morphology of the Vietnamese language (characterized by monosyllabic compound words "
        f"and context-dependent tokenization) and the nascent stage of corporate non-financial disclosure infrastructure {c_nguyen}."
    )

    pb.add_p(
        doc,
        f"In response to these imperatives, this paper replicates, refines, and extensively expands the analytical framework of {c_kang} for the Vietnamese equity market. "
        f"We examine a comprehensive panel of seven premier listed conglomerates representing key economic pillars (Banking & Insurance, Agriculture & Food, Energy & Petroleum, "
        f"Jewelry Retail, Securities & Financial Services, Engineered Materials Manufacturing, and Dairy Consumer Goods) across six continuous reporting years (2020–2025). "
        f"By mining 96,461 digitally extracted sentences across 42 publications (integrating high-accuracy OCR for scanned reports), we not only uncover structural commitment patterns and narrative sentiment distributions, "
        f"but also introduce a dynamic continuous Greenwashing Index (GW) juxtaposed against monthly equity market fluctuations, providing critical methodological lessons "
        f"for automated non-financial auditing in emerging economies."
    )

    # =========================================================================
    # 2. LITERATURE REVIEW
    # =========================================================================
    pb.add_h1(doc, "2. Literature Review")

    pb.add_h2(doc, "2.1 SDGs and Corporate Sustainability Reporting Frameworks")
    c_pizzi = cite("Pizzi et al.", "2020")
    c_munoz = cite("Muñoz-Torres et al.", "2019")
    c_freeman = cite("Freeman", "1984")
    c_deegan = cite("Deegan", "2002")
    pb.add_p(
        doc,
        f"Adopted in 2015 by the United Nations General Assembly, the 2030 Agenda for Sustainable Development articulates 17 universal goals (SDGs) and 169 targets, "
        f"providing an overarching blueprint for global environmental preservation, social equity, and economic prosperity {c_un}. Grounded in Stakeholder Theory {c_freeman} "
        f"and Legitimacy Theory {c_deegan}, corporations operate under an implicit social contract requiring them to demonstrate societal contributions and mitigate negative "
        f"externalities through rigorous non-financial disclosures. The Global Reporting Initiative (GRI) Universal Standards {c_gri} have emerged as the dominant global benchmark, "
        f"standardizing quantitative and qualitative indicators across economic, environmental, and social dimensions. However, international empirical studies consistently reveal "
        f"systemic cherry-picking, whereby corporations prioritize operationally convenient goals while omitting targets that require costly capital expenditures or fundamental business model restructuring {c_munoz, c_pizzi}."
    )

    pb.add_h2(doc, "2.2 NLP and Deep Learning in Corporate Financial and Non-Financial Disclosures")
    c_mikolov = cite("Mikolov et al.", "2013")
    c_devlin = cite("Devlin et al.", "2019")
    c_mercereau = cite("Mercereau & Melin", "2020")
    pb.add_p(
        doc,
        f"The computational processing of corporate narratives has evolved from basic word count techniques (Bag-of-Words, TF-IDF) toward dense semantic vector spaces. "
        f"While static word embeddings such as Word2Vec {c_mikolov} and GloVe represented significant advances, they failed to account for polysemy across distinct financial contexts. "
        f"The advent of bidirectional Transformer architectures, epitomized by BERT {c_devlin}, revolutionized contextual representation. Crucially, Sentence-BERT (SBERT) "
        f"developed by {c_reimers} utilized Siamese and triplet network structures to generate semantically meaningful sentence embeddings that enable rapid, accurate cosine similarity "
        f"computations. In the ESG reporting domain, {c_kang} demonstrated that SBERT and Universal Sentence Encoder substantially outperform legacy lexical methods in capturing nuanced "
        f"sustainability disclosures, providing the computational foundation for automated large-scale corporate auditing {c_mercereau}."
    )

    pb.add_h2(doc, "2.3 Narrative Sentiment and Impression Management in Sustainability Disclosures")
    c_loughran = cite("Loughran & McDonald", "2011")
    c_huang = cite("Huang et al.", "2023")
    c_veenstra = cite("Veenstra & Ellemers", "2020")
    pb.add_p(
        doc,
        f"Textual sentiment analysis in finance originated with specialized accounting dictionaries developed by {c_loughran}, demonstrating that generic psychological lexicons "
        f"systematically misclassify standard corporate terms. Subsequent developments produced context-aware deep models such as FinBERT {c_huang} and PhoBERT {c_nguyen}. "
        f"Within sustainability narratives, empirical scholars {c_veenstra} document pervasive structural optimism bias: corporate communicators overwhelmingly deploy positive language "
        f"to celebrate corporate stewardship, while hazardous incidents, emissions non-compliance, and regulatory fines are routinely obfuscated through passive voice and ambiguous framing."
    )

    pb.add_h2(doc, "2.4 Greenwashing Typologies and Capital Market Decoupling")
    c_seele = cite("Seele & Gatti", "2017")
    c_heras = cite("Heras-Saizarbitoria et al.", "2022")
    c_spence = cite("Spence", "1973")
    c_connelly = cite("Connelly et al.", "2011")
    pb.add_p(
        doc,
        f"Greenwashing is broadly defined by {c_lyon} as the selective disclosure of positive environmental information combined with the retention of negative performance data. "
        f"{c_seele} distinguishes between claim greenwashing (misleading rhetoric) and executive greenwashing (decoupled organizational actions). Under Signaling Theory {c_spence, c_connelly}, "
        f"sustainability disclosures should function as costly, credible signals distinguishing genuine ESG leaders from laggards. However, when the cost of publishing polished digital reports "
        f"is negligible compared to genuine environmental remediation, cheap talk proliferates {c_heras}. Consequently, tracking whether corporate tone diverges from actual market performance "
        f"provides a critical empirical lens for identifying opportunistic impression management."
    )

    pb.add_h2(doc, "2.5 Regulatory Evolution of ESG Reporting in Vietnam")
    c_tran = cite("Tran & Beddewela", "2020")
    c_hoang = cite("Hoang et al.", "2019")
    c_gerged = cite("Gerged et al.", "2021")
    pb.add_p(
        doc,
        f"In Vietnam, institutional mandates governing non-financial disclosures have matured rapidly. Mandatory corporate social responsibility reporting was initiated under Circular 155/2015/TT-BTC "
        f"and significantly upgraded under Circular 96/2020/TT-BTC {c_btc}, mandating listed firms to report greenhouse gas emissions, energy usage, and workplace demographics. "
        f"Nonetheless, empirical assessments by {c_tran, c_hoang} emphasize a pronounced governance divide: while top-tier conglomerates (VN30 index) demonstrate increasing sophistication "
        f"in GRI-aligned reporting, the broader market remains mired in superficial compliance without independent external assurance {c_gerged}."
    )

    # =========================================================================
    # 3. METHODOLOGY
    # =========================================================================
    pb.add_h1(doc, "3. Methodology")

    pb.add_h2(doc, "3.1 Sample and Data Collection")
    pb.add_p(
        doc,
        f"To construct a robust, balanced panel capturing Vietnam's dynamic multi-sector economy, this study collects all officially published standalone sustainability "
        f"and integrated reports from seven leading listed corporations across six continuous fiscal years (2020–2025, N = 42 reports). The sample encompasses: "
        f"Bao Viet Holdings (BVH; Insurance & Financial Services), The PAN Group (PAN; Agriculture & Packaged Food), Petrolimex (PLX; Petroleum Distribution & Energy), "
        f"Phu Nhuan Jewelry (PNJ; Luxury Jewelry Manufacturing & Retail), SSI Securities (SSI; Investment Banking & Brokerage), Vicostone (VCS; Engineered Quartz Stone Manufacturing), "
        f"and Vinamilk (VNM; Dairy Processing & Agribusiness). Table 1 outlines the core enterprise profiles."
    )

    # Table 1: Sample Overview (EN)
    h1, d1 = tables_data["t1_en"]
    pb.add_table_clean(
        doc,
        table_label="Table 1",
        table_title="Overview of the Seven Target Vietnamese Listed Corporations (2020–2025)",
        headers=h1,
        data=d1,
        note="Data compiled from official sustainability and integrated reports published via HOSE, HNX, and corporate portals.",
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

    pb.add_h2(doc, "3.2 Text Extraction and OCR Preprocessing Pipeline")
    pb.add_p(
        doc,
        f"Sustainability reports were ingested under rigorous data hygiene protocols designed for emerging-market reporting formats: "
        f"(1) Native electronic PDFs were parsed using the PyMuPDF engine, extracting structured text streams while systematically stripping covers, "
        f"tables of contents, and administrative boilerplate; (2) Scanned image-raster PDFs lacking electronic text layers (specifically PNJ's 2022 52-page scanned report) "
        f"were routed through a customized optical character recognition (OCR) pipeline employing Tesseract 5.4.0 with dual Vietnamese and English language dictionaries "
        f"('vie+eng') at 300 DPI with adaptive binarization, successfully recovering 669 authentic sentences; "
        f"(3) Sentence boundary disambiguation was executed using Vietnamese orthographic conventions; and (4) Strict length thresholding discarded fragments under six words. "
        f"The refined corpus encompasses 96,461 valid sentences across 4,997 document pages, achieving a mean density of 20.35 sentences per page. "
        f"Detailed extraction statistics are presented in Table 2."
    )

    # Table 2: Text Extraction Stats (EN)
    h2, d2 = tables_data["t2_en"]
    pb.add_table_clean(
        doc,
        table_label="Table 2",
        table_title="Descriptive Statistics of Digitally Extracted Sustainability Reports (2020–2025)",
        headers=h2,
        data=d2,
        note="Sample total: 42 reports, 4,997 pages, 96,461 valid sentences (augmented with 669 OCR-recovered sentences for PNJ 2022). Mean sentence density: 20.35 sentences per page.",
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

    pb.add_h2(doc, "3.3 SDG Similarity Quantification via Vietnamese Sentence-BERT")
    pb.add_p(
        doc,
        f"Following the conceptual formulation of {c_kang}, a reference corpus representing the 17 UN SDGs was constructed by professionally translating "
        f"all 169 official UN targets into Vietnamese and augmenting them with verified corporate sustainability statements (~400 reference sentences). "
        f"Let S_g denote the set of reference sentences associated with SDG g (g in {{1, ..., 17}})."
    )
    pb.add_p(
        doc,
        f"Sentence embeddings were generated using vietnamese-sbert {c_reimers}, mapping each report sentence r and reference sentence s into dense 768-dimensional "
        f"vectors r_hat and s_hat. The semantic similarity score between report sentence r and SDG g is computed as the mean cosine similarity:"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{sim}(r, g) = \frac{1}{|S_g|} \sum_{s \in S_g} \cos(\hat{\mathbf{r}}, \hat{\mathbf{s}}) = \frac{1}{|S_g|} \sum_{s \in S_g} \frac{\hat{\mathbf{r}} \cdot \hat{\mathbf{s}}}{\|\hat{\mathbf{r}}\| \|\hat{\mathbf{s}}\|}",
        eq_num="1"
    )

    pb.add_h2(doc, "3.4 Contextual Sentiment Analysis via PhoBERT")
    pb.add_p(
        doc,
        f"While {c_kang} utilized a binary DistilBERT classifier, we deploy PhoBERT-base {c_nguyen} fine-tuned on financial and administrative Vietnamese corpora "
        f"(phobert-base-vietnamese-sentiment). Sentences are classified into Positive, Neutral, and Negative categories. "
        f"The net narrative tone of report t is quantified via the Pos/Neg Sentiment Ratio:"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{Ratio}_t = \frac{N_{\mathrm{pos}, t}}{N_{\mathrm{neg}, t}}",
        eq_num="2"
    )

    pb.add_p(
        doc,
        "Neutral sentences representing purely objective operational statistics are intentionally excluded from the ratio numerator and denominator "
        "to ensure responsiveness to the dialectic between achievements and risk disclosures."
    )

    pb.add_h2(doc, "3.5 Min-Max Scaling and Six Human-Needs Aggregation")
    pb.add_p(
        doc,
        f"Because raw cosine similarity scores typically occupy a narrow interval ([-0.071, 0.531]), we execute global Min-Max scaling across all 17 goals "
        f"to project scores onto an intuitive 0–100 index {c_kang}:"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{Score}(r, g) = \frac{\mathrm{sim}(r, g) - \min_{r', g'}(\mathrm{sim})}{\max_{r', g'}(\mathrm{sim}) - \min_{r', g'}(\mathrm{sim})} \times 100",
        eq_num="3"
    )

    pb.add_p(
        doc,
        f"Following the thematic taxonomy established by {c_kang}, the 17 SDGs are aggregated into six core human-needs categories: "
        f"(1) Life (SDGs 1, 2, 3); (2) Economic (SDGs 8, 9); (3) Equity (SDGs 4, 5, 10); (4) Social (SDGs 11, 16, 17); "
        f"(5) Resources (SDGs 6, 7, 12, 14); and (6) Environments (SDGs 13, 15). The category score for sentence r is computed as:"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{Cat}_k(r) = \frac{1}{|G_k|} \sum_{g \in G_k} \mathrm{Score}(r, g)",
        eq_num="4"
    )

    pb.add_h2(doc, "3.6 Capital Market Benchmarking and Continuous Greenwashing Index (GW)")
    pb.add_p(
        doc,
        f"Extending beyond {c_kang}, we evaluate the correspondence between sustainability narratives and capital market returns. Monthly closing prices "
        f"were collected for all seven equities and converted to a base index of 100 at the initial month. Annual stock price percentage change (\\Delta\\text{{price}}\\%) "
        f"and annual sentiment ratio change (\\Delta\\text{{Pos/Neg}}\\%) were calculated."
    )
    pb.add_p(
        doc,
        f"When a firm experiences substantial annual market devaluation (\\Delta\\text{{price}}\\% < -10\\%), corporate behavioral responses are classified as: "
        f"(1) Greenwashing (Ratio inflation: \\Delta\\text{{Pos/Neg}}\\% > 15\\%); (2) Honest / Realistic (Tone contraction: \\Delta\\text{{Pos/Neg}}\\% < 0\\%); "
        f"and (3) Inconclusive (None). To capture nuanced variations rather than relying strictly on discrete binary flags, we formulate a continuous Greenwashing Score (GW):"
    )

    pb.add_equation_clean(
        doc,
        r"\mathrm{GW} = \frac{\max(0, -\Delta \text{price}\%) \times \max(0, \Delta \text{Pos/Neg}\%)}{100}",
        eq_num="5"
    )

    pb.add_p(
        doc,
        "The GW score is stratified into four severity tiers: GW = 0 (None / Honest); 0 < GW <= 2 (Mild); 2 < GW <= 8 (Moderate); and GW > 8 (Severe). "
        "Overall market-narrative coupling is formally evaluated via Pearson (r) and Spearman (\\rho) correlation coefficients:"
    )

    pb.add_equation_clean(
        doc,
        r"r = \frac{\sum_{i=1}^n (x_i - \bar{x})(y_i - \bar{y})}{\sqrt{\sum_{i=1}^n (x_i - \bar{x})^2 \sum_{i=1}^n (y_i - \bar{y})^2}}",
        eq_num="6"
    )

    # =========================================================================
    # 4. EMPIRICAL RESULTS
    # =========================================================================
    pb.add_h1(doc, "4. Empirical Results")

    pb.add_h2(doc, "4.1 Global SDG Similarity Distribution")
    pb.add_figure_clean(
        doc,
        "similarity_hist.png",
        "Figure 1",
        "Global distribution of SDG similarity scores across 96,461 sentences following 0–100 Min-Max scaling",
        width_inches=5.8
    )
    pb.add_p(
        doc,
        "Figure 1 illustrates the frequency distribution of SDG similarity scores across the entire 96,461-sentence dataset following global Min-Max normalization. "
        "The distribution exhibits a symmetric, bell-shaped Gaussian morphology centered at a mean of 45.43 with a standard deviation of 11.87. "
        "The central mass between 35 and 55 points captures conventional corporate administrative prose describing general operational frameworks and governance routines."
    )
    pb.add_p(
        doc,
        f"The upper right tail (> 65 points), representing approximately 8.5% of all analyzed sentences, encapsulates highly specialized operational disclosures "
        f"directly addressing specific UN targets (such as closed-loop wastewater recycling, carbon capture investments, and sustainable aquaculture protocols). "
        f"This distribution mirrors the empirical patterns observed by {c_kang} in Korean conglomerates, validating the cross-linguistic stability of Sentence-BERT."
    )

    pb.add_h2(doc, "4.2 Six-Category SDG Heatmap Analysis")
    pb.add_figure_clean(
        doc,
        "heatmap_6cat.png",
        "Figure 2",
        "Heatmap of mean corporate alignment across the six human-needs SDG categories for 42 reports (2020–2025)",
        width_inches=5.6
    )

    # Table 3: 6 Category Means (EN)
    h3, d3 = tables_data["t3_en"]
    pb.add_table_clean(
        doc,
        table_label="Table 3",
        table_title="Mean Annual Alignment Scores across Six SDG Categories (0–100 Scale)",
        headers=h3,
        data=d3,
        note="Life: SDGs 1, 2, 3; Economic: SDGs 8, 9; Equity: SDGs 4, 5, 10; Social: SDGs 11, 16, 17; Resources: SDGs 6, 7, 12, 14; Environments: SDGs 13, 15.",
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
        "Figure 2 and Table 3 provide an exhaustive cross-sectional view of corporate alignment across the six categories. A rigid structural hierarchy emerges: "
        "the Economic category (SDGs 8, 9) consistently achieves the highest scores across all firms and years, fluctuating between 43.19 (VNM 2022) and peaks of "
        "51.90 (VCS 2025) and 51.04 (BVH 2025). This reflects the pragmatic orientation of Vietnamese enterprises, where employment generation, capital accumulation, "
        "and industrial infrastructure remain core operational priorities."
    )
    pb.add_p(
        doc,
        "Conversely, the Equity category (SDGs 4, 5, 10) records the lowest scores across the entire sample, stagnating between 38.36 (SSI 2025) and 44.94 (VCS 2025). "
        "The persistent 6-to-8 point gap between the Economic and Equity dimensions underscores an institutional disparity: diversity, inclusion, and internal wage equity "
        "remain secondary considerations relative to core commercial expansion."
    )

    pb.add_h2(doc, "4.3 Longitudinal Category Trends (2020–2025)")
    pb.add_figure_clean(
        doc,
        "trends_6categories.png",
        "Figure 3",
        "Longitudinal evolution of six SDG categories across the seven target corporations (2020–2025)",
        width_inches=5.8
    )
    pb.add_p(
        doc,
        "Figure 3 traces the multi-year trajectory of the six categories. Most corporations exhibit accelerated scores during 2023–2025, driven by heightened regulatory "
        "mandates under Circular 96 and growing foreign institutional investor scrutiny. Vinamilk demonstrates steady upward momentum in Environments (rising from 41.23 in 2020 "
        "to 47.29 in 2025) following its formal commitment to PAS 2060 carbon neutrality."
    )
    pb.add_p(
        doc,
        "Vicostone (VCS) demonstrates an extraordinary expansion in 2025, publishing a 5,974-sentence comprehensive report that propelled its Economic and Social scores "
        "to 51.90 and 50.51 respectively. Financial sector leaders (BVH and SSI) consistently maintain the Social category as their second highest pillar, "
        "reflecting extensive reporting on governance transparency and sustainable finance underwriting."
    )

    pb.add_h2(doc, "4.4 Contextual Narrative Sentiment Dynamics")
    pb.add_figure_clean(
        doc,
        "sentiment_hist.png",
        "Figure 4",
        "Distribution of narrative polarity prediction probabilities from the PhoBERT model",
        width_inches=5.2
    )
    pb.add_figure_clean(
        doc,
        "sentiment_by_company.png",
        "Figure 5",
        "Proportional sentiment composition (Positive, Neutral, Negative) and Pos/Neg Ratio trends across firms",
        width_inches=5.8
    )
    pb.add_figure_clean(
        doc,
        "sentiment_ratio.png",
        "Figure 6",
        "Annual trajectory of the Sentiment Pos/Neg Ratio across the seven target firms (2020–2025)",
        width_inches=5.6
    )

    # Table 4: Sentiment Counts (EN)
    h4, d4 = tables_data["t4_en"]
    pb.add_table_clean(
        doc,
        table_label="Table 4",
        table_title="Corporate Narrative Sentiment Classification across 42 Sustainability Reports",
        headers=h4,
        data=d4,
        note="Classified via PhoBERT-base. Pos/Neg Ratio = Total Positive Sentences / Total Negative Sentences.",
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
        "Figures 4–6 and Table 4 delineate the narrative sentiment profile across the 96,461 sentences. Disclosures are characterized by structural optimism bias: "
        "Positive sentences constitute an absolute majority of 51,966 sentences (53.87%), Neutral statements represent 31,706 sentences (32.87%), while Negative sentences "
        "account for only 12,789 sentences (13.26%). On average, Vietnamese corporations deploy over four positive assertions for every single risk or operational deficiency acknowledged "
        "(sample-wide mean Pos/Neg ratio of 4.06)."
    )
    pb.add_p(
        doc,
        "Significant sectoral divergence is observable: Vinamilk maintains an elevated ratio (averaging 5.38, spanning 4.35 to 6.06), driven by consistent Net Zero branding. "
        "Bao Viet Holdings exhibits substantial neutral volume (exceeding 40%) in 2020–2022 due to technical actuarial and capital adequacy disclosures, before surging above 6.0 in 2023–2025. "
        "Crucially, for PNJ, deploying optical character recognition (OCR) across 52 scanned pages from 2022 successfully recovered 669 valid sentences, establishing an authentic "
        "2022 Pos/Neg ratio of 1.21 (317 positive and 261 negative sentences). This realistic trough reflects operational headwinds, supply chain disruptions, and health safety measures "
        "following the severe Q3/2021 COVID-19 lockdowns in Ho Chi Minh City, completely replacing the artificial 0.33 artifact produced when standard parsers only captured 12 legal caveat lines. "
        "By 2023, PNJ's sentiment ratio rebounded to 4.23 as operations normalized."
    )

    pb.add_h2(doc, "4.5 Stock Price vs. Narrative Tone Alignment")
    pb.add_figure_clean(
        doc,
        "stock_vs_sentiment_greenwashing.png",
        "Figure 7",
        "Juxtaposition of monthly stock prices (blue curve) and annual Pos/Neg ratios (mountain peaks, index 100)",
        width_inches=6.0
    )

    # Table 5: Correlation (EN)
    h5, d5 = tables_data["t5_en"]
    pb.add_table_clean(
        doc,
        table_label="Table 5",
        table_title="Pearson and Spearman Correlation Tests between Annual Stock Price Changes and Sentiment Shifts",
        headers=h5,
        data=d5,
        note="Estimated on annual percentage shifts (% price change vs. % Pos/Neg ratio change). Firm sample n = 5 pairs; Pooled sample N = 35 observations (2020–2025).",
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
        "Figure 7 illustrates the dynamic interplay between monthly stock price performance and annual narrative sentiment ratios indexed to 100 at base month (January 2020). "
        "Across the pooled panel (Table 5, N = 35), the Pearson correlation coefficient is r = -0.043 (p = 0.806) and the Spearman rank correlation is rho = -0.098 (p = 0.575). "
        "With p-values vastly exceeding conventional significance thresholds (p >= 0.120 across all individual corporate series), the empirical results formally establish that "
        "corporate sustainability narratives in Vietnam are entirely decoupled from capital market valuation shocks."
    )

    pb.add_h2(doc, "4.6 Continuous Greenwashing Index (GW) and Firm Rankings")
    pb.add_figure_clean(
        doc,
        "greenwashing_score.png",
        "Figure 8",
        "Heatmap of annual Greenwashing Scores (GW) and bar chart of mean corporate severity rankings",
        width_inches=5.6
    )

    # Table 6: Greenwashing Rank (EN)
    h6, d6 = tables_data["t6_en"]
    pb.add_table_clean(
        doc,
        table_label="Table 6",
        table_title="Quantitative Greenwashing Index (GW) Severity Rankings across the Target Panel",
        headers=h6,
        data=d6,
        note="GW formulated via Eq. (5), activated solely during annual equity declines. Tiers: GW = 0 (None/Honest); 0 < GW <= 2 (Mild); 2 < GW <= 8 (Moderate); GW > 8 (Severe).",
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
        "Figure 8 and Table 6 present the continuous Greenwashing (GW) rankings following full OCR integration. Bao Viet Holdings (BVH) ranks first with a mean GW of 3.02 "
        "and a peak score of 13.96 in 2023. This severe divergence mirrors the systemic bancassurance and investment-linked life insurance crisis in Vietnam during 2023, "
        "where widespread policyholder complaints and Ministry of Finance regulatory inspections depressed BVH stock by -13.18%, yet the corporate integrated report "
        "countered with aggressive reputation-defense cheerleading (Pos/Neg ratio jumping +105.89%). "
        "Vicostone (VCS) ranks second with a mean GW of 1.23 (peak 5.16 in 2022), triggered when aggressive US Federal Reserve interest rate hikes froze North American "
        "residential construction markets (causing a -47.46% stock collapse), while VCS maintained unyielding promotional rhetoric regarding Breton technology (+10.88% ratio increase). "
        "PNJ drops to third position with a normalized mean GW of 0.92 and a peak of 4.62 in 2023 (down 79% from the distorted 21.75 peak prior to OCR correction). "
        "Vinamilk ranks fourth with a mild mean GW of 0.35 (peak 1.13 in 2023), reflecting strategic Net Zero signaling rather than deceptive concealment. "
        "Conversely, The PAN Group, Petrolimex, and SSI Securities achieve flawless mean GW scores of 0.00. During the catastrophic 2022 capital market downturn "
        "(where equities collapsed by -38.69% to -62.17%), all three firms exhibited textbook honest reporting behavior by candidly reducing their narrative sentiment ratios "
        "(-4.22% to -21.62%) to reflect acute macroeconomic realities."
    )

    # =========================================================================
    # 5. IN-DEPTH DISCUSSION & POLICY IMPLICATIONS
    # =========================================================================
    pb.add_h1(doc, "5. In-depth Discussion & Policy Implications")

    pb.add_h2(doc, "5.1 Emerging Market Realities: Economic Dominance and the Equity Blindspot")
    pb.add_p(
        doc,
        "The quantitative extraction of 96,461 sentences across 42 reports illuminates core structural features of corporate disclosure in an emerging economy: "
        "(1) Economic hegemony: Commercial growth, physical asset accumulation, and partnership frameworks (SDGs 8, 9, 17) dominate narrative attention (Economic mean: 47.69 points); "
        "(2) Equity blindspots: Human rights, workplace diversity, and income equality (SDGs 4, 5, 10) remain persistently under-addressed (Equity mean: 42.08 points); and "
        "(3) Pervasive narrative optimism: Disproportionately high positive sentiment (53.87%) combined with low negative acknowledgement (13.26%) confirms that Vietnamese sustainability "
        "reports still primarily function as promotional branding vehicles rather than comprehensive risk-management tools."
    )

    pb.add_h2(doc, "5.2 Sectoral Divergence and Narrative Strategies: Impression Management vs. Honest Adaptation")
    pb.add_p(
        doc,
        "Triangulating temporal sentiment trajectories with sector-specific macroeconomic developments yields valuable empirical findings: "
        "(1) Defensive impression management at Bao Viet Holdings (BVH): When external institutional crises threaten corporate legitimacy—as occurred during the 2023 bancassurance scrutiny—"
        "corporations actively escalate narrative cheerleading to protect brand equity, generating substantial greenwashing scores (GW = 13.96); "
        "(2) Export demand shock at Vicostone (VCS): Cyclical downturns in primary overseas markets (North American housing slowdown in 2022) exposed a reluctance to downgrade promotional tone, "
        "yielding moderate greenwashing (GW = 5.16); "
        "(3) Honest down-market alignment at PAN, PLX, and SSI: Facing acute systemic shocks in 2022 (agricultural input inflation for PAN, oil retail price caps for PLX, and liquidity contraction for SSI), "
        "these corporations demonstrated exemplary transparency by candidly acknowledging operational headwinds, resulting in zero greenwashing penalties; and "
        "(4) Systematic ESG leadership at Vinamilk (VNM): Sustained high sentiment ratios reflect genuine, capital-intensive sustainability milestones, including PAS 2060 carbon-neutrality certifications "
        "for dairy factories and automated Green Farm complexes."
    )

    pb.add_h2(doc, "5.3 Methodological Lessons: The Crucial Role of OCR and the Danger of Black-Box NLP Artifacts")
    pb.add_p(
        doc,
        "A paramount methodological contribution of this research is the empirical identification and successful resolution of the 'digitization artifact trap' in computational finance. "
        "Prior to OCR intervention, PNJ's 2022 sustainability report—issued as a pure raster scanned PDF—yielded only 12 boilerplate legal sentences under standard parsing, "
        "collapsing its Pos/Neg ratio to an artificial nadir of 0.33. When followed by a digitally native 1,050-sentence report in 2023 with a normalized ratio of 4.23, "
        "the algorithmic change metric surged by +1,169%, producing a fabricated peak GW score of 21.75 and falsely categorizing PNJ as the most severe greenwasher in the sample. "
        "By implementing a customized high-resolution OCR pipeline (Tesseract 5.4.0, 'vie+eng', 300 DPI, adaptive binarization), we extracted 669 authentic sentences, restoring the true 2022 ratio "
        "to 1.21. Consequently, the 2023 mathematical shift dropped to +248.3%, lowering PNJ's peak GW score by 79% (to 4.62) and readjusting its ranking to third place. "
        "This case underscores that automated AI auditing tools must never operate as unverified black boxes; comprehensive OCR preprocessing and data hygiene controls are imperative."
    )

    pb.add_h2(doc, "5.4 Policy Recommendations for Regulators and Listed Firms")
    pb.add_p(
        doc,
        "Three actionable policy directives are proposed: (1) Regulators (SSC, HOSE, HNX) must mandate standardized digital reporting standards (such as structured XBRL or "
        "searchable PDF guidelines), eliminating unreadable scanned PDFs; (2) Listed firms must transition toward substantive disclosure by balancing self-congratulatory "
        "rhetoric with candid evaluations of environmental risks; and (3) Institutional investors should utilize NLP tools as screening mechanisms while cross-verifying claims "
        "against audited quantitative metrics."
    )

    pb.add_h2(doc, "5.5 Limitations and Future Research")
    pb.add_p(
        doc,
        "This study is constrained by its 7-firm sample and focus on unstructured textual prose rather than financial appendix tables. Future research should deploy "
        "Large Language Models (LLMs) to perform multimodal auditing and real-time cross-verification against independent investigative news sources."
    )

    # =========================================================================
    # 6. CONCLUSION
    # =========================================================================
    pb.add_h1(doc, "6. Conclusion")
    pb.add_p(
        doc,
        f"This investigation has successfully replicated, validated, and extended the pioneering framework of {c_kang} within the Vietnamese equity market. "
        f"By mining 96,461 sentences across 42 reports from seven premier corporations over 2020–2025 (integrating dedicated OCR recovery for scanned documentation), "
        f"we provide definitive empirical evidence regarding SDG alignment, narrative sentiment bias, and complete market decoupling (pooled r = -0.043, p = 0.806). "
        f"Crucially, the empirical demonstration of OCR data remediation provides a profound methodological blueprint for the international computational accounting and ESG research community."
    )

    # =========================================================================
    # REFERENCES
    # =========================================================================
    pb.add_h1(doc, "References")
    
    apa_refs = [
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

    print("[EN] Built complete English paper structure.")
