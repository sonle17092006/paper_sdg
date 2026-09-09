# Ghi chú chất lượng extract

Tổng: **46912 câu / 29 file**.

## Số câu theo file

| File | Câu | Năm | Ghi chú |
|---|---:|---|---|
| PAN_SR_2019.pdf | 1227 | 2019 |  |
| PAN_SR_2020.pdf | 1464 | 2020 |  |
| PAN_SR_2021.pdf | 1572 | 2021 |  |
| PAN_SR_2022.pdf | 1630 | 2022 |  |
| PAN_SR_2023.pdf | 883 | 2023 |  |
| PAN_SR_2024.pdf | 1758 | 2024 |  |
| PAN_SR_2025.pdf | 2044 | 2025 |  |
| PLX_SR_2018.pdf | 1149 | 2018 |  |
| PLX_SR_2019.pdf | 2127 | 2019 |  |
| PLX_SR_2020.pdf | 2038 | 2020 |  |
| PLX_SR_2021.pdf | 206 | 2021 |  |
| PLX_SR_2022.pdf | 1948 | 2022 |  |
| PLX_SR_2023.pdf | 969 | 2023 |  |
| PLX_SR_2024.pdf | 112 | 2024 |  |
| PLX_SR_2025.pdf | 1377 | 2025 |  |
| PNJ_SR_2019.pdf | 1512 | 2019 |  |
| PNJ_SR_2020.pdf | 1839 | 2020 |  |
| PNJ_SR_2021.pdf | 883 | 2021 |  |
| PNJ_SR_2022.pdf | 12 | 2022 | mỏng — kiểm tra scan / landing PDF / OCR |
| PNJ_SR_2023.pdf | 1050 | 2023 |  |
| PNJ_SR_2024.pdf | 1489 | 2024 |  |
| PNJ_SR_2025.pdf | 1618 | 2025 |  |
| VNM_SR_2019.pdf | 1941 | 2019 |  |
| VNM_SR_2020.pdf | 3031 | 2020 |  |
| VNM_SR_2021.pdf | 3509 | 2021 |  |
| VNM_SR_2022.pdf | 2886 | 2022 |  |
| VNM_SR_2023.pdf | 1907 | 2023 |  |
| VNM_SR_2024.pdf | 2221 | 2024 |  |
| VNM_SR_2025.pdf | 2510 | 2025 |  |

## Cách xử lý

- PDF scan (không text layer): `python process_data.py --ocr --rebuild`
  cần Tesseract + `models/tessdata/vie.traineddata` (tự tải).
- PDF landing/web (PLX 2021, 2024): không phải full report — nên thay file nguồn.
- Câu CID rác / HOME-MENU / mục lục dán 1 dòng: đã lọc trong `src/clean.py`.
