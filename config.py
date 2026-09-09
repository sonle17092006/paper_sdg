"""Đường dẫn, tên model và hằng số phương pháp Kang & Kim (2022)."""

from pathlib import Path

ROOT = Path(__file__).resolve().parent

PDF_DIR = ROOT / "report pdfs"
DATA_DIR = ROOT / "data"
PROCESSED_DIR = DATA_DIR / "processed"
EMBED_DIR = DATA_DIR / "embeddings"
RESULT_DIR = DATA_DIR / "results"
MODEL_DIR = ROOT / "models"
FIGURE_DIR = ROOT / "figures"

SDG_EN_XLSX = DATA_DIR / "sdg_en.xlsx"
SDG_VI_XLSX = DATA_DIR / "sdg_vi.xlsx"

SENTENCES_CSV = PROCESSED_DIR / "sentences.csv"
SENTENCES_PARQUET = PROCESSED_DIR / "sentences.parquet"
SENTENCES_PKL = PROCESSED_DIR / "sentences.pkl"
SDG_CLEAN_PARQUET = PROCESSED_DIR / "sdg_clean.parquet"
SDG_CLEAN_PKL = PROCESSED_DIR / "sdg_clean.pkl"
SDG_CLEAN_CSV = PROCESSED_DIR / "sdg_clean.csv"
MANIFEST_JSON = PROCESSED_DIR / "manifest.json"
PROGRESS_JSON = DATA_DIR / "progress.json"

# HuggingFace hub ids — lần đầu tải về, lần sau load từ MODEL_DIR
HUB_MODELS = {
    "en_sbert": "sentence-transformers/all-MiniLM-L6-v2",
    "en_sentiment": "distilbert-base-uncased-finetuned-sst-2-english",
    "vi_sbert": "keepitreal/vietnamese-sbert",
    "vi_sentiment": "wonrax/phobert-base-vietnamese-sentiment",
}

LOCAL_MODELS = {name: MODEL_DIR / name for name in HUB_MODELS}

# Paper: bỏ block/câu quá ngắn; tiếng Việt từ đơn âm nên dùng ngưỡng 6 từ (thay vì 10 từ tiếng Anh)
MIN_WORD_CNT = 6
SKIP_PAGES = (1,)
EXCLUDE_FILES = {"PLX_SR_2018.pdf"}

# Paper dùng all-MiniLM max 256; DistilBERT 512. PhoBERT/SBERT VN thường 256.
MAX_SEQ_LEN = {
    "en_sbert": 256,
    "vi_sbert": 256,
    "en_sentiment": 512,
    "vi_sentiment": 256,
}

ENCODE_BATCH = 128
SENTIMENT_BATCH = 64

GOAL_COLS = [f"goal{i:02d}" for i in range(1, 18)]

# Kang & Kim (2022), Table grouping 17 SDGs → 6 categories
CATEGORY_MAP = {
    "goal01": "Life",
    "goal02": "Life",
    "goal03": "Life",
    "goal04": "Equity",
    "goal05": "Equity",
    "goal10": "Equity",
    "goal06": "Resources",
    "goal07": "Resources",
    "goal12": "Resources",
    "goal14": "Resources",
    "goal08": "Economic",
    "goal09": "Economic",
    "goal11": "Social",
    "goal16": "Social",
    "goal17": "Social",
    "goal13": "Environments",
    "goal15": "Environments",
}

CATEGORY_ORDER = [
    "Life",
    "Economic",
    "Equity",
    "Social",
    "Resources",
    "Environments",
]

GP_NAME = {
    "gp01": "Life",
    "gp02": "Economic and Technological Development",
    "gp03": "Equity",
    "gp04": "Social Development",
    "gp05": "Resources",
    "gp06": "Environments",
}

GOAL_TO_GP = {
    "goal01": "gp01",
    "goal02": "gp01",
    "goal03": "gp01",
    "goal08": "gp02",
    "goal09": "gp02",
    "goal04": "gp03",
    "goal05": "gp03",
    "goal10": "gp03",
    "goal11": "gp04",
    "goal16": "gp04",
    "goal17": "gp04",
    "goal06": "gp05",
    "goal07": "gp05",
    "goal12": "gp05",
    "goal14": "gp05",
    "goal13": "gp06",
    "goal15": "gp06",
}

GOAL_TITLE_EN = {
    "goal01": "No Poverty",
    "goal02": "Zero Hunger",
    "goal03": "Good Health and Well-being",
    "goal04": "Quality Education",
    "goal05": "Gender Equality",
    "goal06": "Clean Water and Sanitation",
    "goal07": "Affordable and Clean Energy",
    "goal08": "Decent Work and Economic Growth",
    "goal09": "Industry, Innovation and Infrastructure",
    "goal10": "Reduced Inequalities",
    "goal11": "Sustainable Cities and Communities",
    "goal12": "Responsible Consumption and Production",
    "goal13": "Climate Action",
    "goal14": "Life Below Water",
    "goal15": "Life on Land",
    "goal16": "Peace, Justice and Strong Institutions",
    "goal17": "Partnerships for the Goals",
}

GOAL_TITLE_VI = {
    "goal01": "Chấm dứt nghèo",
    "goal02": "Không còn nạn đói",
    "goal03": "Sức khỏe và phúc lợi",
    "goal04": "Giáo dục chất lượng",
    "goal05": "Bình đẳng giới",
    "goal06": "Nước sạch và vệ sinh",
    "goal07": "Năng lượng sạch với giá phải chăng",
    "goal08": "Việc làm thỏa đáng và tăng trưởng kinh tế",
    "goal09": "Công nghiệp, đổi mới và hạ tầng",
    "goal10": "Giảm bất bình đẳng",
    "goal11": "Thành phố và cộng đồng bền vững",
    "goal12": "Sản xuất và tiêu dùng có trách nhiệm",
    "goal13": "Hành động vì khí hậu",
    "goal14": "Tài nguyên biển",
    "goal15": "Tài nguyên đất liền",
    "goal16": "Hòa bình, công lý và thể chế vững mạnh",
    "goal17": "Đối tác vì các mục tiêu",
}
