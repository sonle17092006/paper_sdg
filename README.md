# Tái hiện Kang & Kim (2022) trên báo cáo PTBV Việt Nam

Tách câu từ PDF → embedding → cosine với 17 SDG → sentiment → bảng/biểu đồ theo công ty–năm.

**Quy trình (không chạy một phát hết):** `load_models.py` → `process_data.py` → `run_pipeline.py --encode-only` → (tuỳ chọn) `--sentiment-only` → `notebook.ipynb` / `export_tables_and_plots.py`. Chi tiết trong `docs.md`. `run_pipeline.py` không cờ = encode **rồi sentiment luôn**.

Paper: [Appl. Sci. 2022, 12(11), 5614](https://doi.org/10.3390/app12115614).  
Tài liệu: **`docs.md`** (paper nào, code làm gì, khác gốc chỗ nào). Chi tiết kỹ thuật: `METHOD.md`.

## Cấu trúc

```
load_models.py              # tải model về models/, lần sau load local
process_data.py             # tách PDF + làm sạch → data/processed/
run_pipeline.py             # encode + cosine + sentiment → embeddings/results
export_tables_and_plots.py  # nhân vector + xuất bảng/PNG, không load model
notebook.ipynb              # follow từng bước, không load model
config.py                   # đường dẫn, tên model, 6 nhóm SDG
src/                        # extract, encode, sentiment, aggregate
report pdfs/<MÃ CK>/*.pdf   # input (không commit)
models/                     # weight local (không commit)
data/sdg_en.xlsx            # corpus SDG tiếng Anh (paper)
data/sdg_vi.xlsx            # corpus SDG tiếng Việt
data/processed/             # câu đã tách
data/embeddings/            # vector L2-normalize + sim.npy
data/results/               # bảng 17 SDG / 6 nhóm
figures/                    # heatmap, trend, histogram
```

## Git không chứa gì

File nặng **gitignore** — clone về rồi tạo lại bằng lệnh bên dưới:

| Bỏ git | Lý do | Tạo lại |
|---|---|---|
| `models/` | ~1.4 GB weight | `python load_models.py` |
| `report pdfs/` | PDF gốc | copy vào `report pdfs/<MÃ CK>/` |
| `data/embeddings/*.npy` | vector + cosine | `python run_pipeline.py --encode-only` |
| `data/processed/sentences.*` | ~17 MB câu | `python process_data.py` |
| `data/results/result_*.pkl/csv` | điểm từng câu | `run_pipeline.py` |
| `_tmp_scrape/` | HTML/JSON tạm | không cần |

**Vẫn commit:** code, `requirements.txt`, corpus SDG (`sdg_en.xlsx`, `sdg_vi.xlsx`), bảng tóm tắt `data/results/table_*.csv`, `figures/*.png`, `notebook.ipynb`.

## Cài môi trường

```powershell
cd C:\Code\paper_sdg
python -m pip install -r requirements.txt
```

## Chạy

### Máy này (đã có model + PDF + embedding)

Sentiment lần trước bị cắt, chạy tiếp:

```powershell
cd C:\Code\paper_sdg
python run_pipeline.py --lang vi --sentiment-only
```

Xem bảng/plot (không load model):

```powershell
python export_tables_and_plots.py
```

Hoặc mở `notebook.ipynb` → Run All.

### Clone máy mới / làm lại từ đầu

```powershell
cd C:\Code\paper_sdg
python load_models.py
# bỏ PDF vào report pdfs\PAN\, PLX\, PNJ\, VNM\
python process_data.py
python run_pipeline.py --lang vi
python export_tables_and_plots.py
```

`load_models.py` chỉ tải nếu `models/` chưa có. Lần sau load đĩa, không lên HuggingFace.

### PDF mới / PDF tiếng Anh

```powershell
python process_data.py                 # incremental theo mtime
python run_pipeline.py --lang vi       # hoặc --lang en / both
```

### Lệnh tách

| Việc | Lệnh |
|---|---|
| Kiểm tra model local | `python load_models.py --list` |
| Tải lại model | `python load_models.py --force` |
| Chỉ extract PDF | `python process_data.py` |
| Extract lại hết | `python process_data.py --rebuild` |
| OCR PDF scan (PNJ 2022…) | `python process_data.py --ocr` |
| Chỉ encode + cosine | `python run_pipeline.py --lang vi --encode-only` |
| Chỉ sentiment | `python run_pipeline.py --lang vi --sentiment-only` |
| Full encode + sentiment | `python run_pipeline.py --lang vi` |
| Xuất bảng + PNG | `python export_tables_and_plots.py` |

CPU, không GPU: encode ~1 giờ, sentiment ~2 giờ (có checkpoint mỗi 500 câu).

## Model

| Việc | Tiếng Anh (paper) | Tiếng Việt |
|---|---|---|
| Similarity | `all-MiniLM-L6-v2` | `keepitreal/vietnamese-sbert` |
| Sentiment | DistilBERT SST-2 | `wonrax/phobert-base-vietnamese-sentiment` |

## Kết quả tóm tắt (PDF đã có)

PAN, PLX, PNJ, VNM (đã bỏ BVH). Bảng 6 nhóm: `data/results/table_company_year_6cat.csv`. Hình: `figures/heatmap_6cat.png`, `trends_6categories.png`, `similarity_hist.png`.

Lưu ý extract: `PNJ_SR_2022` gần như scan (4 câu). `PLX_SR_2021` là landing 7 trang, không phải full report. Chi tiết: `data/processed/extraction_notes.md`.
