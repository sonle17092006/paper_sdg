# Tài liệu phương pháp

Repo này **tái hiện pipeline NLP** của Kang & Kim (2022) trên **báo cáo phát triển bền vững (PTBV) công ty Việt Nam**. Mục tiêu: đo mức độ mỗi báo cáo “nói về” 17 SDG, rồi xem sắc thái tích cực/tiêu cực theo năm.

Cách chạy: `README.md`. Chi tiết kỹ thuật: `METHOD.md`.

Mẫu hiện tại: **PAN, PLX, PNJ, VNM** (đã bỏ BVH).

---

## Quy trình code (lần lượt)

Không phải một file chạy hết mọi thứ. Mỗi bước một script; bước sau đọc file bước trước đã lưu.

```
PDF trong report pdfs/<MÃ CK>/
        │
        ▼
①  load_models.py          tải 4 model vào models/  (một lần)
        │
        ▼
②  process_data.py         PDF → câu, lọc rác, OCR nếu --ocr
        │                  lưu data/processed/sentences.pkl
        ▼
③  run_pipeline.py --encode-only
        │                  câu → vector (SBERT) → cosine với 17 SDG
        │                  lưu data/embeddings/*.npy
        ▼
④  run_pipeline.py --sentiment-only
        │                  từng câu → Positive/Negative/(Neutral)
        │                  lưu data/results/result_both.pkl
        ▼
⑤  notebook.ipynb
   hoặc export_tables_and_plots.py
                           nhân vector (không load model) → bảng + PNG
```

Chi tiết từng bước:

1. **Tải model** — EN: MiniLM + DistilBERT (giống paper). VI: vietnamese-sbert + PhoBERT sentiment. Lần sau load đĩa, không lên HuggingFace.

2. **Extract** — PyMuPDF lấy block chữ, bỏ trang bìa và block < 10 từ (như paper). Trang trống → lấy full text; vẫn trống/`--ocr` → Tesseract tiếng Việt. Lọc HOME/MENU, CID rác, mục lục dán 1 dòng.

3. **Similarity** — mỗi câu báo cáo so với mọi câu corpus SDG cùng ngôn ngữ. Điểm goal *g* = trung bình cosine. Vector L2-normalize nên cosine = phép nhân `report @ sdg.T`.

4. **Sentiment** — mặc định `run_pipeline.py` **không có cờ** sẽ encode (nếu chưa có) **rồi chạy sentiment luôn**. Muốn dừng sau cosine: `--encode-only`.

5. **Bảng/plot** — min–max 0–100, trung bình công ty–năm, gộp 6 nhóm SDG, histogram, heatmap.

---

## Paper gốc

**Kang, H.; Kim, J.** (2022). *Analyzing and Visualizing Text Information in Corporate Sustainability Reports Using Natural Language Processing Methods.* **Applied Sciences**, 12(11), 5614.

- DOI: <https://doi.org/10.3390/app12115614>
- Open access MDPI
- Code notebook gốc: `C:\Code\paper_ssm01` (5 notebook: extract → similarity → sentiment → phân phối → vẽ)

Mẫu paper: 6 tập đoàn quốc tế (BASF, IKEA, Microsoft, Nestlé, Toyota, Walmart), báo cáo **tiếng Anh**, 2011–2020.

Ý tưởng cốt lõi (giữ nguyên):

1. PDF → câu (không dùng tần suất từ).
2. So từng câu báo cáo với corpus 17 SDG bằng **cosine similarity** trên sentence embedding.
3. Điểm mỗi SDG = **trung bình** cosine tới mọi câu SDG của goal đó.
4. Sentiment transformer; điểm NEGATIVE đảo `1 − score` để trục 0–1 hướng dương.
5. Min–max 0–100; gộp công ty–năm; gộp 17 SDG thành **6 nhóm**.

Nhóm 6 category (giống paper):

| Nhóm | SDG |
|---|---|
| Life | 1, 2, 3 |
| Economic | 8, 9 |
| Equity | 4, 5, 10 |
| Social | 11, 16, 17 |
| Resources | 6, 7, 12, 14 |
| Environments | 13, 15 |

---

## Code này làm gì

Áp đúng các bước trên cho PDF trong `report pdfs/` (PAN, PLX, PNJ, VNM).

| Bước paper | File |
|---|---|
| Extract câu | `process_data.py` + `src/extract.py` |
| Embedding + cosine | `run_pipeline.py` (`--encode-only`) |
| Sentiment | `run_pipeline.py` (`--sentiment-only`) |
| Bảng + plot | `notebook.ipynb` hoặc `export_tables_and_plots.py` |

Model tải một lần vào `models/` (`load_models.py`). Vector đã L2-normalize lưu `data/embeddings/*.npy` — notebook **chỉ nhân** `report @ sdg.T`, không load model.

Công thức similarity (như paper):

\[
\mathrm{sim}(r, g) = \frac{1}{|S_g|} \sum_{s \in S_g} \hat{r}\cdot\hat{s}
\]

\(S_g\) = các câu corpus SDG của mục tiêu \(g\).

---

## Khác paper gốc chỗ nào

Báo cáo VN không chạy nguyên model tiếng Anh được. Chỉ đổi chỗ bắt buộc; **công thức và 6 nhóm giữ nguyên**.

| Hạng mục | Paper (2022) | Bản này | Lý do |
|---|---|---|---|
| Dữ liệu | 6 MNCs, EN, 2011–2020 | PAN, PLX, PNJ, VNM, chủ yếu **tiếng Việt** | Đối tượng nghiên cứu khác |
| Corpus SDG | `sdg.xlsx` 641 câu UNGC/SDG Compass (EN) | Giữ `data/sdg_en.xlsx` **và** `data/sdg_vi.xlsx` (tiêu đề + 169 target LHQ + câu DN dịch) | VN-SBERT cần câu SDG tiếng Việt |
| Embedding | `all-MiniLM-L6-v2` | EN: giống paper. VI: `keepitreal/vietnamese-sbert` | MiniLM kém trên tiếng Việt |
| Sentiment | DistilBERT SST-2 (POS/NEG) | EN: giống paper. VI: `wonrax/phobert-base-vietnamese-sentiment` (POS/NEG/**NEU**) | Cần model tiếng Việt; VisoBERT lỗi tokenizer trên transformers 5.x |
| Neutral | Không có | Polarity Neutral = 0.5; tỷ lệ paper vẫn **Positive / Negative** | Lớp thêm của PhoBERT |
| Làm sạch câu | Regex xóa mọi ký tự không ASCII | Chỉ áp dụng câu EN; câu VI **giữ dấu** | Regex paper phá tiếng Việt |
| Tách câu | NLTK `sent_tokenize` | Tách riêng + viết tắt GRI/VN | NLTK lỗi Python 3.14 trên máy này |
| Ngôn ngữ | Một model EN | `doc_lang` theo đa số câu → chọn model VI hoặc EN | Báo cáo có thể song ngữ |
| Extract PDF | Chỉ `blocks`, bỏ block &lt; 10 từ | Giống paper; **fallback** full-page text nếu trang 0 câu | PDF VN nhiều layout/scan |
| Bug notebook gốc | Chỉ `append` label khi không phải NEGATIVE | Ghi nhãn **mọi** câu, rồi đảo điểm NEG | Lệch độ dài label/score |
| Tiện nghiên cứu | 5 notebook Colab | Script + cache local + notebook không load model | Chạy lại không tải HuggingFace |

**Không đổi:** ngưỡng 10 từ, bỏ trang bìa, cosine + trung bình theo goal, min–max 0–100, 6 nhóm SDG, trục sentiment 0–1.

---

## Giới hạn (khi đọc kết quả)

- `PNJ_SR_2022`: PDF scan (trang ảnh, không text layer). OCR: `python process_data.py --ocr` (Tesseract `vie+eng`). Chất lượng thấp hơn PDF có chữ.
- `PLX_SR_2021` / `PLX_SR_2024`: file landing/web, không phải full report — nên thay PDF nguồn.
- Sentiment VN 3 lớp (có Neutral), không so tuyệt đối với SST-2 2 lớp của paper.
- Embedding VI và EN khác không gian vector — không trộn điểm raw giữa hai ngôn ngữ trước khi scale riêng.
- Chưa OCR; PDF ảnh cần bổ sung nếu muốn đủ mẫu.

Khi PDF mới: `python process_data.py` rồi `python run_pipeline.py --lang vi` (hoặc `en`).
