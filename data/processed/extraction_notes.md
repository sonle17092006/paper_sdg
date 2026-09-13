# Ghi chú chất lượng extract

Tổng: **96461 câu / 42 file**.

## Số câu theo file

| File | Câu | Năm | Ghi chú |
|---|---:|---|---|
| BVH_SR_2020.pdf | 5627 | 2020 |  |
| BVH_SR_2021.pdf | 6388 | 2021 |  |
| BVH_SR_2022.pdf | 6419 | 2022 |  |
| BVH_SR_2023.pdf | 3691 | 2023 |  |
| BVH_SR_2024.pdf | 4037 | 2024 |  |
| BVH_SR_2025.pdf | 4515 | 2025 |  |
| PAN_SR_2020.pdf | 1464 | 2020 |  |
| PAN_SR_2021.pdf | 1572 | 2021 |  |
| PAN_SR_2022.pdf | 1630 | 2022 |  |
| PAN_SR_2023.pdf | 883 | 2023 |  |
| PAN_SR_2024.pdf | 1758 | 2024 |  |
| PAN_SR_2025.pdf | 2044 | 2025 |  |
| PLX_SR_2020.pdf | 2038 | 2020 |  |
| PLX_SR_2021.pdf | 206 | 2021 |  |
| PLX_SR_2022.pdf | 1948 | 2022 |  |
| PLX_SR_2023.pdf | 969 | 2023 |  |
| PLX_SR_2024.pdf | 112 | 2024 |  |
| PLX_SR_2025.pdf | 1377 | 2025 |  |
| PNJ_SR_2020.pdf | 1839 | 2020 |  |
| PNJ_SR_2021.pdf | 883 | 2021 |  |
| PNJ_SR_2022.pdf | 669 | 2022 |  |
| PNJ_SR_2023.pdf | 1050 | 2023 |  |
| PNJ_SR_2024.pdf | 1489 | 2024 |  |
| PNJ_SR_2025.pdf | 1618 | 2025 |  |
| SSI_SR_2020.pdf | 1246 | 2020 |  |
| SSI_SR_2021.pdf | 1213 | 2021 |  |
| SSI_SR_2022.pdf | 1190 | 2022 |  |
| SSI_SR_2023.pdf | 1203 | 2023 |  |
| SSI_SR_2024.pdf | 1669 | 2024 |  |
| SSI_SR_2025.pdf | 1938 | 2025 |  |
| VCS_SR_2020.pdf | 2174 | 2020 |  |
| VCS_SR_2021.pdf | 2139 | 2021 |  |
| VCS_SR_2022.pdf | 2455 | 2022 |  |
| VCS_SR_2023.pdf | 2454 | 2023 |  |
| VCS_SR_2024.pdf | 2516 | 2024 |  |
| VCS_SR_2025.pdf | 5974 | 2025 |  |
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
