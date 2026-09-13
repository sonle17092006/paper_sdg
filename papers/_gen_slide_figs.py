"""Generate replacement figures for ESG slides from the 4-company notebook results."""
from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from matplotlib.patches import Patch
from PIL import Image, ImageDraw, ImageFont

from config import CATEGORY_ORDER, FIGURE_DIR, GOAL_COLS, RESULT_DIR

OUT = Path(r"C:\Users\admin\Downloads\_esg_new_figs")
OUT.mkdir(parents=True, exist_ok=True)

sns.set_style("whitegrid")
plt.rcParams.update(
    {
        "font.size": 10,
        "axes.titlesize": 12,
        "axes.titleweight": "bold",
        "figure.facecolor": "white",
        "axes.facecolor": "white",
        "savefig.facecolor": "white",
        "savefig.bbox": "tight",
        "savefig.dpi": 160,
    }
)

CAT_COLORS = {
    "Life": "#2ca02c",
    "Economic": "#1f77b4",
    "Equity": "#9467bd",
    "Social": "#ff7f0e",
    "Resources": "#8c564b",
    "Environments": "#17becf",
}

COMPANY_NAME = {
    "PAN": "PAN — The PAN Group",
    "PLX": "PLX — Petrolimex",
    "PNJ": "PNJ — Phú Nhuận Jewelry",
    "VNM": "VNM — Vinamilk",
}


def save(fig, name: str, size=None) -> Path:
    path = OUT / name
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    if size is not None:
        im = Image.open(path).convert("RGB")
        im = im.resize(size, Image.Resampling.LANCZOS)
        im.save(path)
    return path


def load_tables():
    cat = pd.read_csv(RESULT_DIR / "table_company_year_6cat.csv")
    cat["year"] = cat["year"].astype(int)
    sent = pd.read_csv(RESULT_DIR / "table_sentiment_counts.csv")
    sent["year"] = sent["year"].astype(int)
    stats = pd.read_csv(RESULT_DIR / "table_report_sentence_stats.csv")
    gw = pd.read_csv(RESULT_DIR / "table_greenwashing_score.csv")
    both = pd.read_pickle(RESULT_DIR / "result_both.pkl")
    cy = pd.read_csv(RESULT_DIR / "company_year_means.csv")
    return cat, sent, stats, gw, both, cy


def plot_company_trends(cat: pd.DataFrame):
    companies = ["PAN", "PLX", "PNJ", "VNM"]
    files = ["image17.png", "image19.png", "image18.png", "image20.png"]
    for comp, fname in zip(companies, files):
        sub = cat[cat["company"] == comp].sort_values("year")
        fig, ax = plt.subplots(figsize=(6.4, 2.85))
        for col in CATEGORY_ORDER:
            ax.plot(
                sub["year"],
                sub[col],
                marker="o",
                linewidth=2,
                markersize=5,
                label=col,
                color=CAT_COLORS[col],
            )
        ax.set_title(COMPANY_NAME[comp])
        ax.set_ylabel("Scaled similarity (0–100)")
        ax.set_xlabel("")
        ax.grid(True, alpha=0.4)
        ax.legend(loc="best", fontsize=7, ncol=3, frameon=True)
        ax.set_xticks(sub["year"].tolist())
        save(fig, fname, size=(1089, 490))
        print("wrote", fname)


def plot_year_sdg(cat: pd.DataFrame, cy: pd.DataFrame):
    yearly = cat.groupby("year")[CATEGORY_ORDER].mean().sort_index()
    fig, axes = plt.subplots(1, 2, figsize=(13.4, 4.2), gridspec_kw={"width_ratios": [1.05, 1.15]})

    ax = axes[0]
    for col in CATEGORY_ORDER:
        ax.plot(
            yearly.index,
            yearly[col],
            marker="o",
            linewidth=2.2,
            markersize=6,
            label=col,
            color=CAT_COLORS[col],
        )
    ax.set_title("6-category scores over time (4 DN)")
    ax.set_ylabel("Scaled similarity (0–100)")
    ax.set_xlabel("year")
    ax.legend(loc="lower right", fontsize=8, ncol=2)
    ax.set_xticks(yearly.index.tolist())
    ax.grid(True, alpha=0.4)

    ax = axes[1]
    goal_cols = [c for c in cy.columns if str(c).startswith("goal")]
    if not goal_cols:
        # heatmap of 6 cat by year instead
        heat = yearly.T
        sns.heatmap(heat, ax=ax, cmap="YlOrRd", annot=True, fmt=".1f", cbar_kws={"shrink": 0.85})
        ax.set_title("6 nhóm SDG × năm")
    else:
        heat = cy.copy()
        heat["year"] = heat["year"].astype(int)
        g = heat.groupby("year")[goal_cols].mean().T
        g.index = [f"{i+1:02d}" for i in range(len(g))]
        sns.heatmap(g, ax=ax, cmap="YlOrRd", cbar_kws={"shrink": 0.85})
        ax.set_ylabel("SDG")
        ax.set_title("17 SDGs × year")
        ax.set_yticklabels(g.index, rotation=0)
    fig.tight_layout()
    save(fig, "image21.png", size=(1600, 512))
    print("wrote image21.png")


def plot_sentiment_pie_hist(sent: pd.DataFrame, both: pd.DataFrame):
    tot = sent[["Positive", "Neutral", "Negative"]].sum()
    fig, axes = plt.subplots(1, 2, figsize=(12.4, 4.8), gridspec_kw={"width_ratios": [0.85, 1.25]})

    ax = axes[0]
    sizes = [tot["Positive"], tot["Neutral"], tot["Negative"]]
    labels = ["Positive", "Neutral", "Negative"]
    colors = ["#2ca02c", "#d9d9d9", "#d62728"]
    explode = (0.02, 0.02, 0.02)
    wedges, texts, autotexts = ax.pie(
        sizes,
        labels=None,
        colors=colors,
        explode=explode,
        autopct=lambda p: f"{p:.1f}%",
        startangle=90,
        pctdistance=0.62,
        textprops={"fontsize": 11, "fontweight": "bold"},
    )
    ax.legend(wedges, labels, loc="upper right", fontsize=9)
    ax.set_title("Overall sentiment")

    ax = axes[1]
    scores = both["score"] if "score" in both.columns else None
    if scores is None:
        raise SystemExit(f"no score col: {list(both.columns)}")
    ax.hist(scores.dropna(), bins=40, color="#6c7ae0", edgecolor="white", alpha=0.9)
    ax.set_title("Sentiment score distribution")
    ax.set_xlabel("Score (Negative ≈ 0 → Positive ≈ 1)")
    ax.set_ylabel("Count")
    fig.tight_layout()
    save(fig, "image22.png", size=(1255, 490))
    print("wrote image22.png")


def plot_pn_bar(sent: pd.DataFrame):
    g = sent.groupby("company")[["Positive", "Negative"]].sum()
    g["PN"] = g["Positive"] / g["Negative"]
    g = g.reindex(["VNM", "PAN", "PNJ", "PLX"])
    colors = ["#2ca02c", "#1f77b4", "#9467bd", "#8c564b"]
    fig, ax = plt.subplots(figsize=(6.6, 5.3))
    bars = ax.barh(g.index[::-1], g["PN"].values[::-1], color=colors[::-1], height=0.55)
    ax.set_xlabel("Positive / Negative ratio")
    ax.set_title("P/N ratio — 4 companies (paper metric)")
    for bar, val in zip(bars, g["PN"].values[::-1]):
        ax.text(val + 0.08, bar.get_y() + bar.get_height() / 2, f"{val:.2f}", va="center", fontsize=10)
    ax.set_xlim(0, max(g["PN"]) * 1.18)
    ax.grid(True, axis="x", alpha=0.4)
    fig.tight_layout()
    save(fig, "image23.png", size=(900, 716))
    print("wrote image23.png")


def plot_sentence_table(stats: pd.DataFrame):
    df = stats.copy()
    # columns: Công ty, Năm, Tên file, Số câu, Số trang, Số câu / trang
    wide = df.pivot_table(index="Công ty", columns="Năm", values="Số câu", aggfunc="sum")
    wide = wide.reindex(["PAN", "PLX", "PNJ", "VNM"])
    years = [c for c in wide.columns]
    fig, ax = plt.subplots(figsize=(7.2, 5.5))
    ax.axis("off")
    ax.set_title("Số câu trích được theo công ty–năm", pad=8, fontsize=14, fontweight="bold")
    cell = wide.copy()
    cell_text = []
    for _, row in cell.iterrows():
        cell_text.append(["" if pd.isna(v) else str(int(v)) for v in row])
    col_labels = [str(c) for c in cell.columns]
    row_labels = list(cell.index)
    table = ax.table(
        cellText=cell_text,
        rowLabels=row_labels,
        colLabels=col_labels,
        loc="center",
        cellLoc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1.2, 2.15)
    for (r, c), cell_obj in table.get_celld().items():
        cell_obj.set_edgecolor("#cccccc")
        if r == 0:
            cell_obj.set_facecolor("#1f4e79")
            cell_obj.set_text_props(color="white", fontweight="bold")
        elif c == -1:
            cell_obj.set_facecolor("#1f4e79")
            cell_obj.set_text_props(color="white", fontweight="bold")
        else:
            val = cell.values[r - 1][c]
            if pd.notna(val) and val < 80:
                cell_obj.set_facecolor("#f4c7c3")
            elif r % 2 == 0:
                cell_obj.set_facecolor("#f4f7fb")
            else:
                cell_obj.set_facecolor("white")
    note = "Ô đỏ: báo cáo quá mỏng (PLX 2021 = 48, PLX 2024 = 40). Tổng 17.047 câu / 29 báo cáo."
    ax.text(0.0, -0.02, note, transform=ax.transAxes, fontsize=9, color="#444444")
    fig.tight_layout()
    path = OUT / "image28.png"
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white", pad_inches=0.15)
    plt.close(fig)
    print("wrote image28.png")


def plot_banner():
    w, h = 1024, 121
    im = Image.new("RGB", (w, h), "#111111")
    draw = ImageDraw.Draw(im)
    items = [
        ("PAN", "The PAN Group", "#2e7d32"),
        ("PLX", "Petrolimex", "#ef6c00"),
        ("PNJ", "PNJ Jewelry", "#6a1b9a"),
        ("VNM", "Vinamilk", "#1565c0"),
    ]
    slot = w // 4
    try:
        font_b = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 28)
        font_s = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 14)
        font_t = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 16)
    except Exception:
        font_b = font_s = font_t = ImageFont.load_default()
    draw.rectangle([0, 0, w, 22], fill="#b71c1c")
    draw.text((12, 3), "MẪU 4 DOANH NGHIỆP — BÁO CÁO PTBV", fill="white", font=font_t)
    for i, (tk, name, color) in enumerate(items):
        x0 = i * slot
        draw.rectangle([x0 + 8, 32, x0 + slot - 8, h - 8], outline=color, width=3)
        draw.rectangle([x0 + 8, 32, x0 + 18, h - 8], fill=color)
        draw.text((x0 + 28, 42), tk, fill="white", font=font_b)
        draw.text((x0 + 28, 78), name, fill="#dddddd", font=font_s)
    path = OUT / "image16.jpg"
    im.save(path, quality=92)
    print("wrote image16.jpg")


def crop_panels(src: Path, n: int, names: list[str], target: tuple[int, int]):
    im = Image.open(src).convert("RGB")
    w, h = im.size
    # skip a little top padding if any
    panel_h = h // n
    for i, name in enumerate(names):
        y0 = i * panel_h
        y1 = (i + 1) * panel_h if i < n - 1 else h
        crop = im.crop((0, y0, w, y1))
        crop = crop.resize(target, Image.Resampling.LANCZOS)
        crop.save(OUT / name)
        print("cropped", name, crop.size)


def main():
    cat, sent, stats, gw, both, cy = load_tables()
    print("both cols", list(both.columns)[:20], "n", len(both))
    print("cy cols", list(cy.columns)[:12])
    plot_company_trends(cat)
    plot_year_sdg(cat, cy)
    plot_sentiment_pie_hist(sent, both)
    plot_pn_bar(sent)
    plot_sentence_table(stats)
    plot_banner()

    # stock 2x2 from notebook figure
    crop_panels(
        FIGURE_DIR / "stock_vs_sentiment_greenwashing.png",
        4,
        ["image24.png", "image26.png", "image25.png", "image27.png"],
        (1040, 458),
    )


if __name__ == "__main__":
    main()
