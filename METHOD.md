# Bám bài Kang & Kim (2022) và chỗ thích ứng

Paper: *Analyzing and Visualizing Text Information in Corporate Sustainability Reports using Natural Language Processing Methods*, Appl. Sci. 2022, 12(11), 5614.  
Code gốc: `C:\Code\paper_ssm01` (DOI https://doi.org/10.3390/app12115614).

## Pipeline gốc (5 notebook)

1. **Extract** — PyMuPDF `blocks`, bỏ block không phải text, bỏ block < 10 từ chữ cái, bỏ trang bìa, `nltk.sent_tokenize`, lưu `sentences.csv`.
2. **Sentence similarity** — SentenceTransformer `all-MiniLM-L6-v2`, cosine SDG × câu báo cáo, **trung bình theo từng goal** (17 cột `goal01`…`goal17`).
3. **Sentiment** — `distilbert-base-uncased-finetuned-sst-2-english`; nếu NEGATIVE thì `score = 1 - score` để trục 0–1 hướng dương.
4. **Score distribution** — histogram polarity và similarity (sau min–max 0–100).
5. **Visualization** — trung bình công ty–năm, 6 nhóm SDG, tỷ lệ Positive/Negative.

Nhóm 6 category (giữ nguyên paper):

| Category | SDG |
|---|---|
| Life | 1, 2, 3 |
| Economic | 8, 9 |
| Equity | 4, 5, 10 |
| Social | 11, 16, 17 |
| Resources | 6, 7, 12, 14 |
| Environments | 13, 15 |

## Thích ứng bắt buộc vì báo cáo tiếng Việt

| Hạng mục | Paper (EN) | Bản này |
|---|---|---|
| Corpus SDG | `sdg.xlsx` 641 câu UNGC/SDG Compass | Giữ nguyên `data/sdg_en.xlsx` + corpus VN `data/sdg_vi.xlsx` (tiêu đề + 169 target LHQ + câu doanh nghiệp dịch) |
| Embedding | `all-MiniLM-L6-v2` | EN: giống paper. VI: `keepitreal/vietnamese-sbert` |
| Sentiment | DistilBERT SST-2 (2 lớp) | EN: giống paper. VI: `wonrax/phobert-base-vietnamese-sentiment` (POS/NEG/NEU). VisoBERT lỗi tokenizer trên transformers 5.x. |
| Clean regex ASCII | Xóa mọi ký tự không phải bàn phím EN | Chỉ áp dụng cho câu EN; câu VI giữ Unicode |
| Tách câu | NLTK | Tách riêng (NLTK lỗi trên Python 3.14); viết tắt GRI/VN |
| Ngôn ngữ | Toàn EN | `doc_lang` theo đa số câu; model theo ngôn ngữ báo cáo |

Neutral (VI) gán polarity 0.5. Tỷ lệ paper vẫn là Positive/Negative; Neutral để cột riêng.

## Lỗi trong notebook gốc đã sửa

`03_sentiment_analysis.ipynb` chỉ `label.append` khi không phải NEGATIVE → lệch độ dài với `score`. Bản này ghi nhãn mọi câu.

## File để lần sau không chạy lại phần nặng

- `load_models.py` → weight nằm trong `models/`
- `process_data.py` → câu đã tách trong `data/processed/`
- `run_pipeline.py` → vector L2-normalize `data/embeddings/*.npy`
- Notebook chỉ **nhân vector** (`report @ sdg.T`) + vẽ bảng, **không load model**
