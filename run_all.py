"""Kịch bản chạy tự động toàn bộ Pipeline (End-to-End Master Script).

Quy trình:
1. load_models.py           -> Kiểm tra/tải các model vào models/ (nếu chưa có)
2. process_data.py          -> Bóc tách câu từ 42 báo cáo PDF độc lập (7 công ty: BVH, PAN, PLX, PNJ, SSI, VCS, VNM giai đoạn 2020-2025)
3. run_pipeline.py (encode) -> Mã hóa vector SBERT & tính Cosine Similarity với 17 SDG
4. run_pipeline.py (senti)  -> Phân tích cảm xúc PhoBERT Sentiment
5. export_tables_and_plots  -> Tải giá cổ phiếu, xuất bảng và các biểu đồ phân tích Greenwashing
6. notebook.ipynb           -> Chạy toàn bộ notebook và lưu lại output/hình vẽ inline

Cách dùng:
    python run_all.py                  # Chạy toàn bộ từ đầu đến cuối
    python run_all.py --skip-extract   # Bỏ qua bước trích xuất PDF nếu đã có sentences.parquet
    python run_all.py --skip-sentiment # Bỏ qua bước sentiment nếu chỉ muốn xem SDG similarity
    python run_all.py --notebook-only  # Chỉ chạy lại notebook và xuất bảng/biểu đồ
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Đảm bảo in tiếng Việt trên console Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def run_step(step_name: str, cmd: list[str]) -> None:
    print(f"\n{'='*70}")
    print(f"  BƯỚC: {step_name}")
    print(f"  Lệnh: {' '.join(cmd)}")
    print(f"{'='*70}\n", flush=True)
    t0 = time.time()
    res = subprocess.run(cmd, cwd=str(ROOT))
    if res.returncode != 0:
        print(f"\n[LỖI] Bước '{step_name}' thất bại với mã lỗi {res.returncode}.")
        sys.exit(res.returncode)
    dt = time.time() - t0
    print(f"\n[XONG] '{step_name}' hoàn tất trong {dt:.1f} giây ({dt/60:.2f} phút).\n")


def run_notebook_and_save() -> None:
    print(f"\n{'='*70}")
    print("  BƯỚC: Chạy và lưu notebook.ipynb")
    print(f"{'='*70}\n", flush=True)
    t0 = time.time()
    
    import jupyter_client

    nb_path = ROOT / "notebook.ipynb"
    with nb_path.open(encoding="utf-8") as f:
        nb = json.load(f)

    km = jupyter_client.KernelManager(kernel_name="python3")
    km.start_kernel(cwd=str(ROOT))
    kc = km.client()
    kc.start_channels()
    kc.wait_for_ready(timeout=30)
    print("Kernel Jupyter đã sẵn sàng.")

    # Khởi tạo inline matplotlib
    msg_id = kc.execute("%matplotlib inline")
    while True:
        msg = kc.get_iopub_msg(timeout=10)
        if msg["parent_header"].get("msg_id") == msg_id:
            if msg["msg_type"] == "status" and msg["content"]["execution_state"] == "idle":
                break

    def to_lines(val):
        if isinstance(val, str):
            lines = val.splitlines(keepends=True)
            return lines if lines else [""]
        if isinstance(val, list):
            return val
        return [str(val)]

    def clean_data(data):
        out = {}
        for mime, val in data.items():
            if mime in ("text/plain", "text/html"):
                out[mime] = to_lines(val)
            else:
                out[mime] = val
        return out

    exec_counter = 1
    total_code = sum(1 for c in nb["cells"] if c.get("cell_type") == "code")
    curr_code = 0

    for i, cell in enumerate(nb["cells"]):
        if cell.get("cell_type") != "code":
            continue
        curr_code += 1
        src = "".join(cell.get("source", []))
        cell["execution_count"] = exec_counter
        outputs = []
        msg_id = kc.execute(src)
        print(f"  -> Đang chạy cell {curr_code}/{total_code} (Cell index {i})...", flush=True)

        while True:
            try:
                msg = kc.get_iopub_msg(timeout=180)
            except Exception as exc:
                print(f"Timeout cell {i}: {exc}")
                break

            if msg["parent_header"].get("msg_id") != msg_id:
                continue

            mtype = msg["msg_type"]
            content = msg["content"]

            if mtype == "status" and content["execution_state"] == "idle":
                break
            elif mtype == "stream":
                sname = content["name"]
                stext = to_lines(content["text"])
                if outputs and outputs[-1].get("output_type") == "stream" and outputs[-1].get("name") == sname:
                    outputs[-1]["text"].extend(stext)
                else:
                    outputs.append({"output_type": "stream", "name": sname, "text": stext})
            elif mtype in ("execute_result", "display_data"):
                item = {
                    "output_type": mtype,
                    "data": clean_data(content["data"]),
                    "metadata": content.get("metadata", {}),
                }
                if mtype == "execute_result":
                    item["execution_count"] = exec_counter
                outputs.append(item)
            elif mtype == "error":
                print(f"  [LỖI Cell {i}] {content['ename']}: {content['evalue']}")
                outputs.append({
                    "output_type": "error",
                    "ename": content["ename"],
                    "evalue": content["evalue"],
                    "traceback": content["traceback"],
                })

        cell["outputs"] = outputs
        exec_counter += 1

    kc.stop_channels()
    km.shutdown_kernel(now=True)

    with nb_path.open("w", encoding="utf-8") as f:
        json.dump(nb, f, ensure_ascii=False, indent=1)

    dt = time.time() - t0
    print(f"\n[XONG] notebook.ipynb đã được thực thi và lưu đầy đủ kết quả ({dt:.1f}s).\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Chạy toàn bộ quy trình phân tích SDG & Sentiment")
    parser.add_argument("--skip-models", action="store_true", help="Bỏ qua bước kiểm tra model")
    parser.add_argument("--skip-extract", action="store_true", help="Bỏ qua trích xuất PDF")
    parser.add_argument("--skip-encode", action="store_true", help="Bỏ qua bước mã hóa SBERT")
    parser.add_argument("--skip-sentiment", action="store_true", help="Bỏ qua bước phân tích cảm xúc")
    parser.add_argument("--notebook-only", action="store_true", help="Chỉ chạy lại notebook và xuất biểu đồ")
    args = parser.parse_args()

    py = sys.executable

    print("\n" + "#"*70)
    print("  KHỞI ĐỘNG PIPELINE TÁI HIỆN KANG & KIM (2022) TRÊN BÁO CÁO PTBV VIỆT NAM")
    print("#"*70)

    if args.notebook_only:
        run_step("Xuất bảng, tải giá cổ phiếu & vẽ biểu đồ", [py, "export_tables_and_plots.py"])
        run_notebook_and_save()
        print("\n==== TOÀN BỘ CÔNG VIỆC ĐÃ HOÀN TẤT ====\n")
        return

    # 1. Models
    if not args.skip_models:
        run_step("Kiểm tra và tải mô hình NLP", [py, "load_models.py"])

    # 2. Extract PDF
    if not args.skip_extract:
        run_step("Trích xuất câu sạch từ 42 báo cáo PTBV (7 công ty: BVH, PAN, PLX, PNJ, SSI, VCS, VNM giai đoạn 2020-2025)", [py, "process_data.py", "--rebuild"])

    # 3. Encode SBERT
    if not args.skip_encode:
        run_step("Mã hóa vector SBERT & tính Cosine Similarity 17 SDG", [py, "run_pipeline.py", "--lang", "vi", "--encode-only", "--no-reuse"])

    # 4. Sentiment PhoBERT
    if not args.skip_sentiment:
        run_step("Phân tích cảm xúc PhoBERT Sentiment", [py, "run_pipeline.py", "--lang", "vi", "--sentiment-only"])

    # 5. Export Tables & Plots & Greenwashing Stock Analysis
    run_step("Xuất bảng, tải giá cổ phiếu & vẽ biểu đồ phân tích Greenwashing", [py, "export_tables_and_plots.py"])

    # 6. Execute & Save Notebook
    run_notebook_and_save()

    print("\n" + "#"*70)
    print("  CHÚC MỪNG! TOÀN BỘ PIPELINE ĐÃ CHẠY THÀNH CÔNG RỰC RỠ!")
    print("  - Dữ liệu kết quả: data/results/")
    print("  - Hình vẽ phân tích: figures/")
    print("  - Notebook trực quan hóa: notebook.ipynb")
    print("#"*70 + "\n")


if __name__ == "__main__":
    main()
