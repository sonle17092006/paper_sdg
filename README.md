# Định lượng Cam kết Phát triển Bền vững (SDG) và Phân tích Hiện tượng Tẩy xanh (Greenwashing) tại Doanh nghiệp Tiêu biểu Việt Nam (2020–2025)

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Methodology: Kang & Kim (2022)](https://img.shields.io/badge/Methodology-Kang%20%26%20Kim%20(2022)-brightgreen.svg)](https://doi.org/10.3390/app12115614)
[![NLP: Vietnamese-SBERT & PhoBERT](https://img.shields.io/badge/NLP-Vietnamese--SBERT%20%7C%20PhoBERT-orange.svg)](https://huggingface.co/keepitreal/vietnamese-sbert)

Kho lưu trữ mã nguồn mở và dữ liệu thực nghiệm tái hiện và mở rộng phương pháp luận của nghiên cứu gốc **Kang & Kim (2022)** (*Applied Sciences, MDPI*), ứng dụng công nghệ Xử lý Ngôn ngữ Tự nhiên (NLP) đa tầng và chuỗi giá cổ phiếu thực tế nhằm đo lường mức độ cam kết 17 Mục tiêu Phát triển Bền vững (SDG) và phát hiện dấu hiệu tẩy xanh (greenwashing) trên thị trường chứng khoán Việt Nam.

---

## 1. Tóm tắt Nghiên cứu (Abstract)

Báo cáo Phát triển Bền vững (Báo cáo PTBV độc lập - Standalone Sustainability Report) là kênh công bố thông tin phi tài chính quan trọng của doanh nghiệp. Nghiên cứu này phân tích toàn diện **42 báo cáo PTBV độc lập** của **7 doanh nghiệp đầu ngành** niêm yết tại Việt Nam trong giai đoạn **2020 – 2025** (tương ứng **95.804 câu văn bản** tiếng Việt). 

Quy trình nghiên cứu áp dụng:
1. **Trích xuất & làm sạch văn bản thông minh**: Tách câu theo khối văn bản thực tế, khử header/footer lặp trang bằng PyMuPDF.
2. **Biểu diễn ngữ nghĩa & đo lường SDG**: Mã hóa vector ngữ nghĩa câu văn bằng mô hình **Vietnamese-SBERT** (`keepitreal/vietnamese-sbert`), tính độ tương đồng Cosine chuẩn hóa với 17 mục tiêu SDG của Liên Hợp Quốc và quy đổi về thang 0–100.
3. **Phân nhóm chuyên sâu theo Kang & Kim (2022)**: Gộp 17 SDG thành 6 nhóm chủ đề chiến lược (*Life, Economic, Equity, Social, Resources, Environments*).
4. **Đo lường sắc thái biểu đạt**: Ứng dụng mô hình **PhoBERT Sentiment** (`wonrax/phobert-base-vietnamese-sentiment`) phân loại câu văn thành Tích cực (Positive), Trung tính (Neutral), Tiêu cực (Negative), xác định tỷ lệ $Pos/Neg\ Ratio$.
5. **Mô hình kiểm định Tẩy xanh (Greenwashing Divergence)**: Tích hợp chuỗi giá cổ phiếu đóng cửa hàng tháng (84 tháng: 2019–2025 từ VNDirect DChart API), đối chiếu sự phân kỳ giữa đà giảm giá thị trường ($>10\%$) và sự gia tăng giọng văn tích cực ($>15\%$) để lượng hóa điểm tẩy xanh ($GW\_Score$).

---

## 2. Mẫu Dữ liệu Thực nghiệm (Dataset & Scope)

Nghiên cứu tập trung vào **7 doanh nghiệp quy mô lớn** đại diện cho các ngành kinh tế trọng điểm tại Việt Nam, sở hữu chuỗi báo cáo PTBV độc lập liên tục 6 năm (2020–2025):

| Mã CP | Tên Doanh nghiệp | Ngành kinh tế | Sàn GD | Số báo cáo | Tổng số câu (vi) |
|:---:|---|---|:---:|:---:|:---:|
| **BVH** | Tập đoàn Bảo Việt | Tài chính / Bảo hiểm | HOSE | 6 (2020–2025) | 30.645 |
| **PAN** | Tập đoàn PAN | Nông nghiệp / Chế biến thực phẩm | HOSE | 6 (2020–2025) | 9.336 |
| **PLX** | Tập đoàn Xăng dầu Việt Nam (Petrolimex) | Năng lượng / Phân phối dầu khí | HOSE | 6 (2020–2025) | 6.627 |
| **PNJ** | Vàng bạc Đá quý Phú Nhuận | Bán lẻ / Trang sức | HOSE | 6 (2020–2025) | 6.870 |
| **SSI** | Công ty Cổ phần Chứng khoán SSI | Tài chính / Dịch vụ chứng khoán | HOSE | 6 (2020–2025) | 8.443 |
| **VCS** | Công ty Cổ phần Vicostone | Sản xuất / Đá thạch anh nhân tạo | HNX | 6 (2020–2025) | 17.697 |
| **VNM** | Công ty Cổ phần Sữa Việt Nam (Vinamilk) | Hàng tiêu dùng nhanh / Sữa & đồ uống | HOSE | 6 (2020–2025) | 16.033 |
| **Tổng** | **7 Doanh nghiệp** | **Đa ngành tiêu biểu** | **HOSE & HNX** | **42 Báo cáo** | **95.804 câu** |

*Toàn bộ 42 tài liệu đều là Báo cáo PTBV độc lập (Standalone Sustainability Reports), không sử dụng Báo cáo Thường niên tổng hợp.*

---

## 3. Cấu trúc Thư mục Kho lưu trữ (Repository Structure)

```text
paper_sdg/
│
├── config.py                   # Cấu hình đường dẫn, hằng số ngưỡng, mapping 17 SDG -> 6 nhóm
├── METHOD.md                   # Tài liệu chi tiết phương pháp luận toán học & NLP
├── docs.md                     # Hướng dẫn kỹ thuật và so sánh chi tiết với paper gốc
├── requirements.txt            # Danh sách thư viện Python phụ thuộc
├── .gitignore                  # Cấu hình loại trừ file nhị phân nặng chuẩn nghiên cứu
│
├── run_all.bat                 # Script chạy 1-click toàn bộ pipeline (Windows)
├── run_all.py                  # Script điều phối Python Master Script
├── run_remaining.bat           # Script chạy nhanh các bước mô hình còn lại
│
├── process_data.py             # Bóc tách và làm sạch câu từ 42 file PDF -> data/processed/
├── run_pipeline.py             # Mã hóa SBERT, Cosine Similarity & PhoBERT Sentiment
├── export_tables_and_plots.py  # Lấy giá cổ phiếu, tính Greenwashing, xuất 8 biểu đồ & bảng CSV
├── load_models.py              # Kiểm tra và tải trước mô hình HuggingFace về đĩa
│
├── notebook.ipynb              # Sổ tay Jupyter trực quan hóa tương tác 30 cells
│
├── src/                        # Thư viện module chức năng cốt lõi
│   ├── extract.py              # Trích xuất văn bản từ cấu trúc khối PDF (PyMuPDF)
│   ├── clean.py                # Khử nhiễu, loại bỏ mục lục, mã CID rác và câu cụt
│   ├── tokenize_sent.py        # Phân tách câu văn bản tiếng Việt
│   ├── language.py             # Nhận diện ngôn ngữ tài liệu và từng câu
│   ├── sdg_vi_corpus.py        # Khung văn bản tham chiếu 17 mục tiêu SDG tiếng Việt
│   ├── encode.py               # Vector embedding SBERT, chuẩn hóa L2 và Cosine Similarity
│   ├── sentiment.py            # Phân tích độ phân cực cảm xúc PhoBERT
│   ├── aggregate.py            # Tổng hợp điểm trung bình, min-max scaling và gộp nhóm
│   ├── stock.py                # Tích hợp API VNDirect DChart, lấy chuỗi giá và tính correlation
│   ├── plots.py                # Module đồ họa Matplotlib/Seaborn xuất bản phẩm chất lượng cao
│   ├── io_util.py              # Quản lý đọc/ghi tối ưu (Parquet / Pickle / CSV)
│   └── progress.py             # Ghi nhận trạng thái tiến độ pipeline
│
├── data/                       # Dữ liệu phục vụ nghiên cứu
│   ├── sdg_en.xlsx             # Corpus 17 SDG tiếng Anh (từ paper gốc)
│   ├── sdg_vi.xlsx             # Corpus 17 SDG tiếng Việt đối chiếu
│   ├── processed/              # Bảng thống kê trích xuất văn bản (manifest.json, extraction_notes.md)
│   └── results/                # Bảng kết quả định lượng phục vụ xuất bản bài báo (CSV)
│       ├── table_company_year_17sdg.csv              # Điểm số 17 mục tiêu SDG từng công ty - năm
│       ├── table_company_year_6cat.csv               # Điểm số 6 nhóm chủ đề SDG theo Kang & Kim
│       ├── table_sentiment_counts.csv                # Tỷ lệ câu Pos/Neu/Neg và Pos/Neg Ratio
│       ├── table_report_sentence_stats.csv           # Thống kê số câu và mật độ câu/trang
│       ├── table_stock_sentiment_correlation.csv     # Hệ số tương quan Pearson & Spearman
│       ├── table_stock_sentiment_greenwashing.csv    # Đối chiếu biến động giá và cảm xúc
│       ├── table_greenwashing_score.csv              # Bảng điểm tẩy xanh chi tiết công ty - năm
│       ├── table_greenwashing_rank.csv               # Bảng xếp hạng mức độ tẩy xanh tổng hợp
│       ├── stock_prices_monthly.csv                  # Chuỗi giá đóng cửa 84 tháng (2019-2025)
│       └── stock_prices_annual.csv                   # Giá cổ phiếu tổng hợp theo năm
│
├── figures/                    # Toàn bộ 8 biểu đồ khoa học định dạng PNG xuất bản
│   ├── trends_6categories.png                        # Xu hướng điểm số 6 nhóm SDG theo thời gian
│   ├── heatmap_6cat.png                              # Bản đồ nhiệt điểm SDG 7 doanh nghiệp qua các năm
│   ├── similarity_hist.png                           # Phân phối điểm tương đồng SDG (0-100)
│   ├── sentiment_by_company.png                      # Cơ cấu tỷ lệ cảm xúc từng năm của 7 doanh nghiệp
│   ├── sentiment_ratio.png                           # Biến động chỉ số Pos/Neg Ratio theo năm
│   ├── sentiment_hist.png                            # Phân phối điểm polarity cảm xúc
│   ├── stock_vs_sentiment_greenwashing.png           # Biểu đồ ngọn núi đối chiếu giá cổ phiếu & Pos/Neg
│   └── greenwashing_score.png                        # Heatmap & Barplot xếp hạng điểm Greenwashing
│
└── papers/                     # Bản thảo bài báo khoa học và thư viện trích dẫn
    ├── paper_sdg_vietnam_vi.docx                     # Bản thảo toàn văn tiếng Việt
    ├── paper_sdg_vietnam_en.docx                     # Bản thảo toàn văn tiếng Anh
    ├── references_sdg_vietnam.enw                    # Thư viện trích dẫn EndNote (.enw)
    ├── references_sdg_vietnam.ris                    # Thư viện trích dẫn chuẩn (.ris)
    └── ...                                           # Các công cụ tạo bản thảo tự động
```

---

## 4. Hướng dẫn Cài đặt & Tái hiện Thực nghiệm (Reproducibility)

### 4.1. Cài đặt Môi trường
Yêu cầu Python $\ge$ 3.10. Khuyến nghị tạo môi trường ảo:
```powershell
# Khởi tạo môi trường ảo
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Cài đặt toàn bộ thư viện cần thiết
pip install -r requirements.txt
```

### 4.2. Chạy Tự động 1-Click (Khuyến nghị)
Nhấp đúp chuột vào file:
```text
run_all.bat
```
Hoặc thực thi qua dòng lệnh:
```powershell
python run_all.py
```

### 4.3. Chạy Tách rời từng Giai đoạn (Modular Execution)
1. **Bước 1 — Trích xuất câu từ PDF**:
   ```powershell
   python process_data.py --rebuild
   ```
2. **Bước 2 — Mã hóa SBERT & Tính Cosine Similarity 17 SDG**:
   ```powershell
   python run_pipeline.py --lang vi --encode-only --no-reuse
   ```
3. **Bước 3 — Phân tích Cảm xúc PhoBERT**:
   ```powershell
   python run_pipeline.py --lang vi --sentiment-only
   ```
4. **Bước 4 — Xuất Bảng số liệu & 8 Biểu đồ Tẩy xanh**:
   ```powershell
   python export_tables_and_plots.py
   ```

### 4.4. Khởi chạy Jupyter Notebook Tương tác
Khởi động sổ tay nghiên cứu để tương tác và tùy biến đồ họa:
```powershell
jupyter notebook notebook.ipynb
```

---

## 5. Phương pháp Định lượng Tẩy xanh (Greenwashing Metric)

Nghiên cứu định nghĩa và lượng hóa hiện tượng **Tẩy xanh (Greenwashing)** dựa trên nguyên lý phân kỳ ngược chiều giữa hiệu quả thị trường và thái độ công bố thông tin:

1. **Điều kiện phát sinh tín hiệu Tẩy xanh**:
   Xảy ra khi giá cổ phiếu cuối năm của doanh nghiệp sụt giảm đáng kể ($\Delta P < -10\%$), phản ánh thị trường tài chính ghi nhận khó khăn hoặc rủi ro tiêu cực, nhưng báo cáo PTBV lại gia tăng đột biến sắc thái biểu đạt tích cực ($\Delta (Pos/Neg) > +15\%$).

2. **Chỉ số Điểm Tẩy xanh ($GW\_Score$)**:
   $$\text{GW\_Score} = \frac{\max(0, -\Delta P) \times \max(0, \Delta \text{Ratio})}{100}$$
   * Trong đó: $\Delta P$ là phần trăm thay đổi giá cổ phiếu so với cùng kỳ; $\Delta \text{Ratio}$ là phần trăm thay đổi tỷ số câu Tích cực / Tiêu cực.
   * Doanh nghiệp có giá giảm mạnh nhưng vẫn gia tăng giọng điệu tô hồng sẽ có $GW\_Score$ cao.
   * Ngược lại, nếu giá cổ phiếu giảm và doanh nghiệp thừa nhận khó khăn (giảm $Pos/Neg\ Ratio$), trường hợp này được phân loại là **Trung thực (Honest)**.

---

## 6. Trích dẫn Nghiên cứu (Citation)

Nếu bạn sử dụng mã nguồn, dữ liệu hoặc kết quả từ nghiên cứu này, vui lòng trích dẫn theo định dạng:

```bibtex
@article{kang2022task,
  title={Task-Adaptive Pre-Trained Language Models for Measuring Corporate Sustainability},
  author={Kang, Min-Seok and Kim, Young-Min},
  journal={Applied Sciences},
  volume={12},
  number={11},
  pages={5614},
  year={2022},
  publisher={MDPI},
  doi={10.3390/app12115614}
}
```

---

## 7. Giấy phép (License)
Dự án được phân phối dưới giấy phép mã nguồn mở **MIT License**. Chi tiết xem tại file [LICENSE](LICENSE).
