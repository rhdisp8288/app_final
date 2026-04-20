import sys
# [핵심 패치 1] 사내 DRM/백신이 파이썬 임시 파일(.pyc)을 감시하며 
# 시스템을 멈추게 하는 현상을 OS 레벨에서 원천 금지합니다.
sys.dont_write_bytecode = True 

import time
import io
import os
import hashlib
import importlib.util
import difflib
import re
import socket
import tempfile
import urllib.request
import streamlit as st
import pandas as pd
import xlsxwriter
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import pypdfium2 as pdfium
import json
from typing import List, Dict, Optional, Any
from rapidocr_onnxruntime import RapidOCR
from app.analysis.quote_adequacy import (
    build_llm_context,
    evaluate_quote_adequacy,
    generate_html_report,
)
from app.analysis.industry_financial_db import (
    build_estimated_industry_baseline,
    ensure_workbook,
    get_baseline_items,
    ingest_industry_statements,
    load_industry_db,
)

INDUSTRY_GROUP_LABELS = {
    "C": "제조업",
    "D": "전기, 가스, 증기 및 공기조절 공급업",
    "E": "수도, 하수 및 폐기물 처리, 원료 재생업",
    "F": "건설업",
    "G": "도매 및 소매업",
    "J": "정보통신업",
    "M": "전문, 과학 및 기술 서비스업",
    "N": "사업시설 관리, 사업 지원 및 임대 서비스업",
    "O": "공공행정, 국방 및 사회보장 행정",
    "P": "교육 서비스업",
    "S": "협회 및 단체, 수리 및 기타 개인 서비스업",
}
BASIS_TYPE_LABELS = {
    "observed": "실측 기반",
    "repaired": "복원 기반",
    "proxy": "업종 프록시",
    "estimated": "보수 추정",
}
CONFIDENCE_LABELS = {
    "A": "높음 (A)",
    "B": "보통 (B)",
    "C": "낮음 (C)",
    "D": "매우 낮음 (D)",
}
ADEQUACY_JUDGEMENT_LABELS = {
    "ok": "적정",
    "watch": "주의",
    "alert": "경고",
    "critical": "심각",
    "OK": "적정",
    "WATCH": "주의",
    "ALERT": "경고",
    "CRITICAL": "심각",
}
DEFAULT_STANDARD_CATEGORIES = [
    "1.일반용역",
    "2.SW용역",
    "3-1.종합건축",
    "3-2.전문/인테리어",
    "3-3.전기/통신",
    "4.마케팅",
    "5.연구용역",
]
STANDARD_CATEGORY_GROUPS = [
    {"id": "general", "label": "일반용역"},
    {"id": "sw", "label": "SW용역"},
    {"id": "construction", "label": "공사"},
    {"id": "marketing", "label": "마케팅"},
    {"id": "research", "label": "연구용역"},
]
STANDARD_CATEGORY_GROUP_LABELS = {
    row["id"]: row["label"]
    for row in STANDARD_CATEGORY_GROUPS
}

# ==============================================================================
# [NEW 핵심 패치 7] 8GB RAM 환경을 위한 극한의 메모리 최적화 (Additive)
# ==============================================================================
# PyTorch 메모리 할당자가 메모리 파편화를 방지하고 부족한 RAM을 효율적으로 쓰도록 강제 설정
os.environ["PYTORCH_CUDA_ALLOC_CONF"] = "expandable_segments:True"

# ==============================================================================
# [핵심 패치 2] 사내망 DNS 타임아웃 방지 (Streamlit 서버 IP 확인 중 무한 멈춤 해결)
# ==============================================================================
_original_gethostbyname = socket.gethostbyname
def _patched_gethostbyname(hostname):
    if hostname == socket.gethostname(): return "127.0.0.1"
    try: return _original_gethostbyname(hostname)
    except: return "127.0.0.1"
socket.gethostbyname = _patched_gethostbyname

# ==============================================================================
# [핵심 패치 3] 네트워크 병목, DRM 충돌, 스레드 폭주 원천 차단
# ==============================================================================
os.environ["TRANSFORMERS_OFFLINE"] = "1"
os.environ["HF_DATASETS_OFFLINE"] = "1"
os.environ["HF_HUB_OFFLINE"] = "1"
os.environ["http_proxy"] = ""
os.environ["https_proxy"] = ""
os.environ["HTTP_PROXY"] = ""
os.environ["HTTPS_PROXY"] = ""
os.environ["NO_PROXY"] = "localhost,127.0.0.1,::1,*"

os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
os.environ["STREAMLIT_GATHER_USAGE_STATS"] = "false"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["WANDB_DISABLED"] = "true"
os.environ["DISABLE_TELEMETRY"] = "1"

# Streamlit 사내망 충돌 방지 및 웹소켓 보호
os.environ["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"
os.environ["STREAMLIT_SERVER_ENABLE_CORS"] = "false"
os.environ["STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION"] = "false"
os.environ["STREAMLIT_SERVER_ENABLE_WEBSOCKET_COMPRESSION"] = "false"
os.environ["TOKENIZERS_PARALLELISM"] = "false" 
os.environ["PLOTLY_OFFLINE"] = "True"
os.environ["STREAMLIT_SERVER_MAX_UPLOAD_SIZE"] = "10" 

os.environ["STREAMLIT_SERVER_WEBSOCKET_PING_INTERVAL"] = "5"    
os.environ["STREAMLIT_SERVER_WEBSOCKET_PING_TIMEOUT"] = "86400" 
os.environ["STREAMLIT_SERVER_MESSAGE_TIMEOUT"] = "86400"        

os.environ["PYTORCH_JIT"] = "0"
os.environ["TORCH_INDUCTOR_DISABLE"] = "1"
# ==============================================================================

urllib.request.install_opener(urllib.request.build_opener(urllib.request.ProxyHandler({})))

# --------------------------------------------------------------------------------
# [0] 시스템 설정 및 전역 스타일 (MVC: View - Global)
# --------------------------------------------------------------------------------
st.set_page_config(page_title="Pro_Buying_Agent v6.1: EXAONE Edition", page_icon="🛡️", layout="wide")

# 라이브러리 유무 체크
try:
    from pydantic import BaseModel, Field, ValidationError
    HAS_PYDANTIC = True
except ImportError:
    HAS_PYDANTIC = False

try:
    from rank_bm25 import BM25Okapi
    from rapidfuzz import process, fuzz
    HAS_NLP_LIB = True
except ImportError:
    HAS_NLP_LIB = False
    fuzz = None 

def apply_custom_css():
    st.markdown("""
    <style>
        html, body, .stApp, [class*="css"] { font-family: 'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', Dotum, sans-serif !important; color: #333333; }
        .stApp { background: linear-gradient(180deg, #f6f7f9 0%, #eef1f4 100%); }
        .block-container { padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1520px; }
        [data-testid="stSidebar"] { background: linear-gradient(180deg, #ffffff 0%, #f6f7fa 100%); border-right: 1px solid #E5E7EB; }
        .header-box { background: linear-gradient(180deg, #ffffff 0%, #fbfcfd 100%); padding: 26px 30px 24px; border-radius: 0; margin-bottom: 30px; border-top: 8px solid #E53935; border-bottom: 1px solid #E5E7EB; box-shadow: 0 10px 28px rgba(15, 23, 42, 0.06); }
        .header-title { font-size: 34px; font-weight: 900; color: #17243C; letter-spacing: -0.7px; margin-bottom: 8px; }
        .header-subtitle { font-size: 17px; color: #6B7280; font-weight: 600; letter-spacing: -0.2px; }
        .section-header-l1 { font-size: 24px; font-weight: 900; color: #17243C; padding: 0 0 10px; border-left: none; border-bottom: 3px solid #E53935; background: transparent; margin-top: 42px; margin-bottom: 22px; border-radius: 0; letter-spacing: -0.5px; }
        .section-header-l2 { font-size: 18px; font-weight: 800; color: #374151; margin-top: 8px; margin-bottom: 14px; padding-left: 0; border-left: none; border-bottom: 1px solid #F0B24A; display: inline-block; padding-bottom: 6px; }
        .metric-container { background-color: #FFFFFF; padding: 24px 16px 22px; border: 1px solid #E5E7EB; border-top: 4px solid #E53935; text-align: center; border-radius: 0; box-shadow: 0 6px 16px rgba(15, 23, 42, 0.03); }
        .metric-label { font-size: 12px; color: #6B7280; margin-bottom: 10px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; }
        .metric-value { font-size: 26px; font-weight: 900; color: #17243C; letter-spacing: -0.6px; }
        .stTabs [data-baseweb="tab-list"] { gap: 8px; border-bottom: none; }
        .stTabs [data-baseweb="tab"] { height: 48px; white-space: pre-wrap; background-color: #FFFFFF; border-radius: 0; color: #4B5563; font-weight: 700; border: 1px solid #DEE2E6; padding: 0 18px; }
        .stTabs [aria-selected="true"] { background-color: #FFF7F5 !important; color: #17243C !important; border: 1px solid #E53935 !important; font-weight: 900; box-shadow: none; }
        .component-title { font-size: 15px; font-weight: 700; color: #666666; margin-bottom: 10px; display: block; }
        .insight-card { background-color: #FFFFFF; padding: 16px 20px; border: 1px solid #E5E7EB; margin-bottom: 10px; border-radius: 0; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03); }
        .insight-title { font-weight: 800; color: #17243C; font-size: 15px; margin-bottom: 6px; }
        div.stButton > button { border-radius: 0; font-weight: 800; border: 1px solid #17243C; background-color: #17243C; color: white; padding: 0.54rem 1.2rem; transition: all 0.2s; }
        div.stButton > button:hover { background-color: #0F172A; color: white; border-color: #0F172A; }
        .stDownloadButton > button { border-radius: 0 !important; font-weight: 800 !important; border: 1px solid #E53935 !important; background-color: #E53935 !important; color: #FFFFFF !important; }
        .stDownloadButton > button:hover { background-color: #C62828 !important; border-color: #C62828 !important; color: #FFFFFF !important; }
        [data-testid="stExpander"] { background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 0; box-shadow: 0 4px 10px rgba(15, 23, 42, 0.03); }
        [data-testid="stDataFrame"] { border: 1px solid #E5E7EB; }
    </style>
    """, unsafe_allow_html=True)

# --------------------------------------------------------------------------------
# [0-1] 데이터 로딩 캐싱
# --------------------------------------------------------------------------------
@st.cache_data(ttl=86400, max_entries=5)
def load_db_data(filepath, sheet_name=0):
    if not os.path.exists(filepath): return pd.DataFrame()
    try:
        try: df = pd.read_excel(filepath, sheet_name=sheet_name)
        except: df = pd.read_excel(filepath, sheet_name=0)

        new_cols = []
        for c in df.columns:
            clean_c = c.strip().replace(" ", "_")
            new_cols.append(clean_c)
        df.columns = new_cols
        return df
    except Exception: return pd.DataFrame()


ESTIMATE_COLUMN_ALIASES = {
    "카테고리": ["카테고리", "category", "분류", "용역구분", "공종"],
    "비목": ["비목", "품목", "항목", "비용항목", "item"],
    "품명(상세내역)": ["품명(상세내역)", "내역(상세내역)", "상세내역", "세부내역", "내역", "품명", "description", "item_name"],
    "등급": ["등급", "직급", "레벨", "grade"],
    "단위": ["단위", "unit"],
    "수량": ["수량", "qty", "quantity", "물량"],
    "단가": ["단가", "price", "unitprice", "금액단가"],
    "금액": ["금액(원)", "금액", "공급가액", "합계", "amount", "total"],
}
ESTIMATE_REQUIRED_COLUMNS = {"비목", "단가", "수량"}


def _normalize_estimate_token(value):
    text = str(value or "").strip().lower()
    text = re.sub(r"[\s_\-]+", "", text)
    text = text.replace("(원)", "").replace("원", "")
    return text


def _match_estimate_column_name(raw_name):
    token = _normalize_estimate_token(raw_name)
    if not token:
        return None
    for canonical, aliases in ESTIMATE_COLUMN_ALIASES.items():
        for alias in aliases:
            alias_token = _normalize_estimate_token(alias)
            if not alias_token:
                continue
            if token == alias_token:
                return canonical
            if alias_token in token and len(alias_token) >= 2:
                return canonical
    return None


def _make_unique_headers(headers):
    counts = {}
    unique_headers = []
    for idx, raw_header in enumerate(headers):
        text = str(raw_header or "").strip()
        header = text if text else f"col_{idx+1}"
        counts[header] = counts.get(header, 0) + 1
        if counts[header] > 1:
            header = f"{header}_{counts[header]}"
        unique_headers.append(header)
    return unique_headers


def _coerce_numeric_series(series):
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("원", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
    )
    cleaned = cleaned.replace({"": np.nan, "-": np.nan, "nan": np.nan, "None": np.nan})
    return pd.to_numeric(cleaned, errors="coerce")


def _score_estimate_header_row(row_values):
    found = set()
    score = 0
    for raw_value in row_values:
        canonical = _match_estimate_column_name(raw_value)
        if not canonical or canonical in found:
            continue
        found.add(canonical)
        score += 4 if canonical in ESTIMATE_REQUIRED_COLUMNS else 1
    return score, found


@st.cache_data(show_spinner=False)
def load_uploaded_estimate_data(file_bytes):
    workbook = io.BytesIO(file_bytes)
    raw_df = pd.read_excel(workbook, header=None)
    raw_df = raw_df.dropna(how="all").dropna(axis=1, how="all")
    if raw_df.empty:
        raise ValueError("엑셀 파일에서 읽을 수 있는 데이터가 없습니다.")

    best_header_idx = 0
    best_score = -1
    best_found = set()
    max_scan_rows = min(len(raw_df), 20)
    for idx in range(max_scan_rows):
        score, found = _score_estimate_header_row(raw_df.iloc[idx].tolist())
        if score > best_score:
            best_score = score
            best_found = found
            best_header_idx = idx

    if not ESTIMATE_REQUIRED_COLUMNS.issubset(best_found):
        best_header_idx = 0

    headers = _make_unique_headers(raw_df.iloc[best_header_idx].tolist())
    df = raw_df.iloc[best_header_idx + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all").dropna(axis=1, how="all")

    rename_map = {}
    used_canonical = set()
    for column in df.columns:
        canonical = _match_estimate_column_name(column)
        if canonical and canonical not in used_canonical:
            rename_map[column] = canonical
            used_canonical.add(canonical)
    df = df.rename(columns=rename_map)

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()

    for numeric_col in ["수량", "단가", "금액"]:
        if numeric_col in df.columns:
            df[numeric_col] = _coerce_numeric_series(df[numeric_col])

    if "단가" not in df.columns and {"금액", "수량"}.issubset(df.columns):
        qty_nonzero = df["수량"].replace(0, np.nan)
        df["단가"] = df["금액"] / qty_nonzero
    if "수량" not in df.columns and {"금액", "단가"}.issubset(df.columns):
        price_nonzero = df["단가"].replace(0, np.nan)
        df["수량"] = df["금액"] / price_nonzero
    if "금액" not in df.columns and {"단가", "수량"}.issubset(df.columns):
        df["금액"] = df["단가"].fillna(0) * df["수량"].fillna(0)

    if "비목" in df.columns:
        item_series = df["비목"].fillna("").astype(str).str.strip()
        summary_mask = item_series.str.contains("합계|총계|소계", case=False, regex=True)
        df = df[~summary_mask].copy()

    if "금액" in df.columns:
        df = df[df["금액"].fillna(0) != 0].copy()

    df = df.reset_index(drop=True)

    recognized_columns = [col for col in df.columns if col in ESTIMATE_COLUMN_ALIASES]
    notes = [
        f"헤더 자동 인식: {best_header_idx + 1}행",
        f"인식 컬럼: {', '.join(recognized_columns) if recognized_columns else '없음'}",
    ]
    missing_required = [col for col in ESTIMATE_REQUIRED_COLUMNS if col not in df.columns]
    if missing_required:
        notes.append(f"필수 컬럼 보완 필요: {', '.join(missing_required)}")

    return df, {
        "header_row": best_header_idx + 1,
        "recognized_columns": recognized_columns,
        "missing_required": missing_required,
        "row_count": len(df),
        "notes": notes,
        "summary": " / ".join(notes + [f"분석 대상 행 수: {len(df)}"]),
    }


FINANCIAL_UPLOAD_LABELS = {
    "BS": "재무상태표(BS)",
    "IS": "손익계산서(IS)",
    "CF": "현금흐름표(CF)",
    "COST": "제조원가명세서(COST, 선택)",
}
FINANCIAL_REQUIRED_UPLOADS = ("BS", "IS", "CF")
FINANCIAL_UPLOAD_TYPES = ["pdf", "xlsx", "xlsm", "xls"]
FINANCIAL_STATEMENT_DETECTION_RULES = {
    "BS": ["재무상태표", "대차대조표", "balancesheet", "statementoffinancialposition"],
    "IS": ["손익계산서", "포괄손익계산서", "incomestatement", "profitandloss", "statementofincome"],
    "CF": ["현금흐름표", "cashflow", "statementofcashflows"],
    "COST": ["제조원가명세서", "원가명세서", "costofgoodsmanufactured", "cogm"],
}
BUSINESS_CERT_UPLOAD_TYPES = ["pdf", "png", "jpg", "jpeg", "bmp", "webp"]
BUSINESS_CERT_STOP_LABELS = [
    "상호",
    "법인명",
    "단체명",
    "사업자등록번호",
    "성명",
    "대표자",
    "개업연월일",
    "법인등록번호",
    "사업장소재지",
    "본점소재지",
    "업태",
    "종목",
    "교부사유",
    "발급사유",
]
INDUSTRY_GROUP_KEYWORDS = [
    ("F", ["전기공사", "정보통신공사", "실내건축", "건설", "토목", "시공", "인테리어", "설비공사", "시설물유지", "배관공사"]),
    ("J", ["소프트웨어", "정보통신", "시스템개발", "프로그램개발", "컴퓨터프로그래밍", "데이터", "플랫폼", "호스팅", "통신서비스", "it서비스", "sw"]),
    ("C", ["제조업", "제조", "가공", "조립", "생산", "부품", "기계", "금속", "전자", "식품제조", "화장품제조"]),
    ("G", ["도소매", "도매", "소매", "유통", "판매", "무역", "전자상거래", "상품중개"]),
    ("M", ["엔지니어링", "컨설팅", "연구개발", "기술서비스", "설계", "시험분석", "광고대행", "디자인", "회계", "법무", "연구용역", "전문서비스"]),
    ("N", ["사업지원", "시설관리", "경비", "청소", "인력공급", "근로자파견", "임대업", "렌탈", "콜센터", "여행사"]),
    ("P", ["교육", "학원", "훈련", "연수", "평생교육"]),
    ("S", ["수리업", "세탁", "미용", "협회", "단체", "개인서비스", "장례"]),
    ("E", ["폐기물", "재활용", "하수", "수도", "정수", "원료재생", "환경복원"]),
    ("D", ["전력", "발전업", "가스", "증기", "열공급", "에너지공급"]),
    ("O", ["공공행정", "국방", "사회보장", "행정기관"]),
]


@st.cache_resource
def get_ocr_engine():
    return RapidOCR()


def _normalize_ocr_token(text):
    return re.sub(r"[^0-9A-Za-z가-힣]+", "", str(text or "")).strip().lower()


def _format_business_id(value):
    digits = re.sub(r"\D+", "", str(value or ""))
    if len(digits) != 10:
        return str(value or "").strip()
    return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"


def _trim_field_value(value):
    text = re.sub(r"\s+", " ", str(value or "")).strip(" :,-")
    text = re.sub(r"^(법인명|단체명|상호|대표자|사업장소재지|본점소재지|업태|종목)\s*", "", text).strip()
    return text


def _looks_like_label_line(line):
    normalized = _normalize_ocr_token(line)
    return any(_normalize_ocr_token(label) in normalized for label in BUSINESS_CERT_STOP_LABELS)


def _render_pdf_bytes_to_images(file_bytes, page_limit=2):
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_tmp_pdf_uploads")
    os.makedirs(temp_dir, exist_ok=True)
    tmp_path = None
    images = []
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf", dir=temp_dir) as handle:
            handle.write(file_bytes)
            tmp_path = handle.name
        pdf = pdfium.PdfDocument(tmp_path)
        try:
            for page_idx in range(min(len(pdf), page_limit)):
                pil_image = pdf[page_idx].render(scale=2.5).to_pil()
                images.append(np.array(pil_image.convert("RGB")))
        finally:
            pdf.close()
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    return images


def _load_business_certificate_images(uploaded_file):
    filename = str(getattr(uploaded_file, "name", "") or "").lower()
    file_bytes = uploaded_file.getvalue()
    if filename.endswith(".pdf"):
        return _render_pdf_bytes_to_images(file_bytes)
    image = Image.open(io.BytesIO(file_bytes)).convert("RGB")
    return [np.array(image)]


def _extract_pdf_text_lines(file_bytes, page_limit=2):
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_tmp_pdf_uploads")
    os.makedirs(temp_dir, exist_ok=True)
    tmp_path = None
    collected = []
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf", dir=temp_dir) as handle:
            handle.write(file_bytes)
            tmp_path = handle.name
        pdf = pdfium.PdfDocument(tmp_path)
        try:
            for page_idx in range(min(len(pdf), page_limit)):
                page = pdf[page_idx]
                text = (page.get_textpage().get_text_range() or "").replace("\x00", "")
                if text.strip():
                    collected.extend(line.strip() for line in text.splitlines() if line.strip())
        finally:
            pdf.close()
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    return collected


def _ocr_lines_from_image_array(image_array, y_tolerance=22):
    result, _ = get_ocr_engine()(image_array)
    tokens = []
    for box, text, _score in result or []:
        ys = [point[1] for point in box]
        xs = [point[0] for point in box]
        tokens.append((sum(ys) / len(ys), sum(xs) / len(xs), str(text)))
    tokens.sort(key=lambda row: (row[0], row[1]))

    grouped = []
    for cy, cx, text in tokens:
        if not grouped or abs(grouped[-1][0] - cy) > y_tolerance:
            grouped.append([cy, [(cx, text)]])
        else:
            grouped[-1][1].append((cx, text))

    lines = []
    for _, items in grouped:
        items.sort(key=lambda row: row[0])
        line_text = " ".join(text for _, text in items).strip()
        if line_text:
            lines.append(line_text)
    return lines


def _extract_field_from_lines(lines, labels, *, join_following=False, max_follow_lines=2):
    normalized_labels = [_normalize_ocr_token(label) for label in labels]
    for idx, line in enumerate(lines):
        normalized_line = _normalize_ocr_token(line)
        if not any(label in normalized_line for label in normalized_labels):
            continue

        candidate = line
        for label in labels:
            candidate = re.sub(rf"{re.escape(label)}\s*[:：]?\s*", "", candidate, flags=re.IGNORECASE)
        candidate = _trim_field_value(candidate)

        extra_lines = []
        next_idx = idx + 1
        while next_idx < len(lines) and len(extra_lines) < max_follow_lines:
            next_line = lines[next_idx].strip()
            if not next_line or _looks_like_label_line(next_line):
                break
            extra_lines.append(next_line)
            if not join_following:
                break
            next_idx += 1

        if extra_lines:
            if candidate:
                candidate = f"{candidate} {' '.join(extra_lines)}"
            else:
                candidate = " ".join(extra_lines)

        candidate = _trim_field_value(candidate)
        if candidate:
            return candidate
    return ""


def _extract_field_from_text(full_text, pattern, *, multiline=False):
    flags = re.IGNORECASE | (re.DOTALL if multiline else 0)
    match = re.search(pattern, full_text, flags)
    if not match:
        return ""
    return _trim_field_value(re.sub(r"\s+", " ", match.group(1)).strip())


def _is_noise_certificate_line(line):
    normalized = _normalize_ocr_token(line)
    if not normalized:
        return True
    noise_tokens = [
        "ntsgokr",
        "사업자등록증",
        "간이과세자",
        "세무서장",
        "등록번호",
        "발급사유",
        "공동사업자",
        "전자세금계산서",
        "사업자단위과세",
    ]
    if any(token in normalized for token in noise_tokens):
        return True
    if re.search(r"\d{4}\s*년|\d{3}-?\d{2}-?\d{5}", line):
        return True
    return False


def _infer_industry_group_from_text(*texts):
    joined_text = " ".join(str(text or "") for text in texts)
    normalized_text = _normalize_ocr_token(joined_text)
    best_group = "M"
    best_score = 0
    best_keywords = []

    for group_code, keywords in INDUSTRY_GROUP_KEYWORDS:
        score = 0
        matched = []
        for keyword in keywords:
            key = _normalize_ocr_token(keyword)
            if key and key in normalized_text:
                score += max(1, len(key) // 3)
                matched.append(keyword)
        if score > best_score:
            best_group = group_code
            best_score = score
            best_keywords = matched

    if best_keywords:
        reason = f"사업자등록증 OCR 키워드 매칭: {', '.join(best_keywords[:4])}"
        confidence = "high" if best_score >= 4 else "medium"
    else:
        reason = "명확한 업종 키워드가 없어 전문서비스(M) 기본값을 사용"
        confidence = "low"

    return {
        "industry_group": best_group,
        "industry_label": INDUSTRY_GROUP_LABELS.get(best_group, best_group),
        "reason": reason,
        "confidence": confidence,
    }


def extract_business_certificate_info(uploaded_file):
    filename = str(getattr(uploaded_file, "name", "") or "").lower()
    lines = []
    if filename.endswith(".pdf"):
        lines.extend(_extract_pdf_text_lines(uploaded_file.getvalue()))

    if not lines:
        images = _load_business_certificate_images(uploaded_file)
        for image_array in images:
            lines.extend(_ocr_lines_from_image_array(image_array))

    full_text = "\n".join(lines)
    business_id_match = re.search(r"\d{3}-?\d{2}-?\d{5}", full_text)
    business_id = _format_business_id(business_id_match.group(0)) if business_id_match else ""
    company_name = _extract_field_from_text(
        full_text,
        r"상\s*호\s*[:：]\s*([^\n]+)",
    ) or _extract_field_from_lines(lines, ["상호", "법인명", "단체명", "상호명"])
    location = _extract_field_from_text(
        full_text,
        r"(?:사\s*업\s*장\s*소\s*재\s*지|본\s*점\s*소\s*재\s*지)\s*[:：]\s*(.+?)(?=\n\s*(?:사\s*업\s*의\s*종\s*류|발\s*급\s*사\s*유|공\s*동\s*사\s*업\s*자|사업자\s*단위\s*과세|전자세금계산서)|$)",
        multiline=True,
    ) or _extract_field_from_lines(lines, ["사업장 소재지", "사업장소재지", "본점 소재지", "본점소재지", "소재지"], join_following=True)
    business_type = _extract_field_from_text(
        full_text,
        r"업\s*태\s*([^\n]+)",
    ) or _extract_field_from_lines(lines, ["업태"], join_following=False)
    business_item = _extract_field_from_text(
        full_text,
        r"종\s*목\s*(.+?)(?=\n\s*(?:발\s*급\s*사\s*유|공\s*동\s*사\s*업\s*자|사업자\s*단위\s*과세|전자세금계산서)|$)",
        multiline=True,
    ) or _extract_field_from_lines(lines, ["종목"], join_following=True)

    if not company_name:
        line_candidates = [
            line for line in lines
            if 2 <= len(line.strip()) <= 40 and not _is_noise_certificate_line(line)
        ]
        company_name = line_candidates[0].strip() if line_candidates else ""

    inference = _infer_industry_group_from_text(company_name, location, business_type, business_item, full_text)
    summary_parts = []
    if business_type:
        summary_parts.append(f"업태: {business_type}")
    if business_item:
        summary_parts.append(f"종목: {business_item}")
    if inference["industry_group"]:
        summary_parts.append(f"추정 산업군: {inference['industry_label']} ({inference['industry_group']})")

    return {
        "company_name": company_name,
        "business_id": business_id,
        "location": location,
        "business_type": business_type,
        "business_item": business_item,
        "industry_group": inference["industry_group"],
        "industry_label": inference["industry_label"],
        "industry_reason": inference["reason"],
        "industry_confidence": inference["confidence"],
        "summary_text": " / ".join(summary_parts),
        "ocr_text": full_text,
        "ocr_lines": lines,
    }


def _safe_upload_token(value):
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "").strip())
    return text.strip("._") or "upload"


def _extract_excel_preview_text(file_bytes, max_rows=20):
    snippets = []
    try:
        workbook = pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception:
        return ""

    for sheet_name in workbook.sheet_names[:4]:
        try:
            df = pd.read_excel(workbook, sheet_name=sheet_name, header=None, nrows=max_rows)
        except Exception:
            continue
        flat_values = []
        for value in df.fillna("").astype(str).values.flatten().tolist():
            value = str(value).strip()
            if value:
                flat_values.append(value)
        if flat_values:
            snippets.append(" ".join(flat_values[:120]))
    return "\n".join(snippets)


@st.cache_data(show_spinner=False)
def _extract_financial_preview_text_cached(filename: str, file_bytes: bytes):
    filename = str(filename or "").lower()
    file_bytes = file_bytes or b""
    if filename.endswith(".pdf"):
        lines = _extract_pdf_text_lines(file_bytes, page_limit=2)
        if lines:
            return "\n".join(lines[:80])
        images = _render_pdf_bytes_to_images(file_bytes, page_limit=1)
        ocr_lines = []
        for image_array in images:
            ocr_lines.extend(_ocr_lines_from_image_array(image_array))
        return "\n".join(ocr_lines[:80])
    return _extract_excel_preview_text(file_bytes)


def _extract_financial_preview_text(uploaded_file):
    filename = str(getattr(uploaded_file, "name", "") or "").lower()
    file_bytes = uploaded_file.getvalue()
    return _extract_financial_preview_text_cached(filename, file_bytes)


def _normalize_statement_detection_text(text):
    return re.sub(r"[^0-9A-Za-z가-힣]+", "", str(text or "")).lower()


def _detect_financial_statement_type(uploaded_file):
    filename = str(getattr(uploaded_file, "name", "") or "")
    preview_text = _extract_financial_preview_text(uploaded_file)
    combined_name = _normalize_statement_detection_text(filename)
    combined_text = _normalize_statement_detection_text(preview_text)

    best_type = None
    best_score = 0
    matched_keywords = []
    for statement_type, keywords in FINANCIAL_STATEMENT_DETECTION_RULES.items():
        score = 0
        matched = []
        for keyword in keywords:
            token = _normalize_statement_detection_text(keyword)
            if token and token in combined_name:
                score += 5
                matched.append(f"name:{keyword}")
            if token and token in combined_text:
                score += 3
                matched.append(f"text:{keyword}")
        if score > best_score:
            best_type = statement_type
            best_score = score
            matched_keywords = matched

    current_year = datetime.utcnow().year + 1
    years = sorted(
        {
            str(year)
            for year in re.findall(r"(19\d{2}|20\d{2}|21\d{2})", preview_text)
            if 2000 <= int(year) <= current_year
        },
        reverse=True,
    )

    return {
        "statement_type": best_type if best_score > 0 else None,
        "score": best_score,
        "matched_keywords": matched_keywords,
        "preview_text": preview_text,
        "years": years,
        "filename": filename,
    }


def classify_financial_statement_uploads(uploaded_files):
    uploaded_files = uploaded_files or []
    selected_files = {}
    details = []
    duplicates = []
    unclassified = []
    detected_years = set()

    for uploaded_file in uploaded_files:
        detail = _detect_financial_statement_type(uploaded_file)
        details.append(detail)
        detected_years.update(detail.get("years", []))
        statement_type = detail.get("statement_type")
        if not statement_type:
            unclassified.append(detail["filename"])
            continue
        current = selected_files.get(statement_type)
        if current is None or detail["score"] > current["score"]:
            if current is not None:
                duplicates.append(current["filename"])
            selected_files[statement_type] = {**detail, "uploaded_file": uploaded_file}
        else:
            duplicates.append(detail["filename"])

    return {
        "statement_files": {
            statement_type: detail["uploaded_file"]
            for statement_type, detail in selected_files.items()
        },
        "detected_map": {
            statement_type: detail["filename"]
            for statement_type, detail in selected_files.items()
        },
        "details": details,
        "duplicates": duplicates,
        "unclassified": unclassified,
        "detected_years": sorted(detected_years, reverse=True),
    }


def _sanitize_classification_meta(classification_meta):
    meta = dict(classification_meta or {})
    return {
        "detected_map": dict(meta.get("detected_map") or {}),
        "details": list(meta.get("details") or []),
        "duplicates": list(meta.get("duplicates") or []),
        "unclassified": list(meta.get("unclassified") or []),
        "detected_years": list(meta.get("detected_years") or []),
    }


def _json_default(value):
    if isinstance(value, np.generic):
        return value.item()
    return str(value)


def _safe_float_value(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except Exception:
        return None


class BufferedUpload:
    def __init__(self, name: str, data: bytes):
        self.name = str(name or "")
        self._data = data or b""

    def getvalue(self) -> bytes:
        return self._data


def _fast_upload_signature(upload_obj) -> str:
    payload = upload_obj.getvalue() if upload_obj is not None else b""
    payload = payload or b""
    if len(payload) <= 8192:
        sample = payload
    else:
        sample = payload[:4096] + payload[-4096:]
    digest = hashlib.sha1(sample).hexdigest()
    return f"{getattr(upload_obj, 'name', '')}:{len(payload)}:{digest}"


def _latest_metric_ratio(
    metric_df: pd.DataFrame,
    industry_group: str,
    period_key: str,
    item_code: str,
) -> Optional[float]:
    if metric_df is None or metric_df.empty:
        return None
    filtered = metric_df[
        (metric_df["industry_group"].astype(str) == str(industry_group))
        & (metric_df["period_key"].astype(str) == str(period_key))
        & (metric_df["item_code"].astype(str) == str(item_code))
    ].copy()
    if filtered.empty:
        return None
    filtered["created_at_sort"] = pd.to_datetime(filtered["created_at"], errors="coerce")
    filtered = filtered.sort_values("created_at_sort")
    return _safe_float_value(filtered.iloc[-1].get("value_pct"))


def _latest_statement_amount(
    statement_df: pd.DataFrame,
    industry_group: str,
    period_key: str,
    account_code: str,
    *,
    absolute: bool = True,
) -> Optional[float]:
    if statement_df is None or statement_df.empty:
        return None
    filtered = statement_df[
        (statement_df["industry_group"].astype(str) == str(industry_group))
        & (statement_df["period_key"].astype(str) == str(period_key))
        & (statement_df["normalized_account_code"].astype(str) == str(account_code))
    ].copy()
    if filtered.empty:
        return None
    filtered["ingested_at_sort"] = pd.to_datetime(filtered["ingested_at"], errors="coerce")
    filtered = filtered.sort_values("ingested_at_sort")
    value = _safe_float_value(filtered.iloc[-1].get("amount"))
    if value is None:
        return None
    return abs(value) if absolute else value


def _estimate_utilization(asset_turnover: Optional[float], cogs_ratio: Optional[float]) -> Optional[float]:
    turnover = _safe_float_value(asset_turnover)
    cogs = _safe_float_value(cogs_ratio)
    if turnover is None and cogs is None:
        return None

    turnover_score = 0.72
    if turnover is not None:
        turnover_score = 0.55 + max(0.0, min(turnover, 2.0)) * 0.20

    cogs_score = 0.72
    if cogs is not None:
        cogs_score = 0.60 + max(0.0, min(cogs, 0.95)) * 0.30

    return max(0.55, min((turnover_score * 0.6) + (cogs_score * 0.4), 0.95))


def _extract_financial_ratios_from_ingest(
    workbook_path: str,
    industry_group: str,
    period_keys: List[str],
    selected_period: Optional[str] = None,
) -> Dict[str, Any]:
    if not workbook_path or not os.path.exists(workbook_path):
        return {}

    sheets = load_industry_db(workbook_path)
    statement_df = sheets.get("statement_rows", pd.DataFrame())
    metric_df = sheets.get("metric_observations", pd.DataFrame())
    normalized_periods = [str(period) for period in period_keys if str(period).strip()]
    if not normalized_periods:
        return {}

    current_period = str(selected_period).strip() if str(selected_period or "").strip() in normalized_periods else next((period for period in normalized_periods if str(period).isdigit()), normalized_periods[0])
    previous_period = None
    for candidate in normalized_periods:
        if candidate != current_period:
            previous_period = candidate
            break

    ratios: Dict[str, Any] = {
        "selected_period": current_period,
        "previous_period": previous_period,
    }

    metric_ratio_map = {
        "observed_material_ratio": "material",
        "observed_labor_ratio": "labor",
        "observed_expense_ratio": "expense",
        "observed_gna_ratio": "gna",
        "observed_profit_ratio": "profit",
        "observed_rnd_ratio": "rnd",
    }
    for target_key, item_code in metric_ratio_map.items():
        metric_value = _latest_metric_ratio(metric_df, industry_group, current_period, item_code)
        if metric_value is not None:
            ratios[target_key] = metric_value / 100.0

    sales_revenue = _latest_statement_amount(statement_df, industry_group, current_period, "is_revenue")
    operating_profit = _latest_statement_amount(statement_df, industry_group, current_period, "is_operating_profit", absolute=False)
    total_assets = _latest_statement_amount(statement_df, industry_group, current_period, "bs_total_assets")
    cogs_amount = _latest_statement_amount(statement_df, industry_group, current_period, "is_cogs")
    employee_benefits = _latest_statement_amount(statement_df, industry_group, current_period, "is_employee_benefits")
    direct_labor_amount = _latest_statement_amount(statement_df, industry_group, current_period, "cost_labor")

    if sales_revenue:
        ratios["sales_revenue"] = float(sales_revenue)
    if sales_revenue and operating_profit is not None:
        ratios["operating_margin"] = float(operating_profit) / float(sales_revenue)
    if sales_revenue and total_assets:
        ratios["asset_turnover"] = float(sales_revenue) / float(total_assets)

    cogs_ratio = None
    if sales_revenue and cogs_amount is not None:
        cogs_ratio = float(cogs_amount) / float(sales_revenue)
    else:
        direct_sum = sum(
            float(ratios.get(key, 0.0) or 0.0)
            for key in ("observed_material_ratio", "observed_labor_ratio", "observed_expense_ratio")
        )
        if direct_sum > 0:
            cogs_ratio = direct_sum
    if cogs_ratio is not None:
        ratios["cogs_ratio"] = cogs_ratio

    if employee_benefits and direct_labor_amount and direct_labor_amount > 0:
        burden_candidate = (float(employee_benefits) - float(direct_labor_amount)) / float(direct_labor_amount)
        ratios["labor_burden_rate"] = max(0.05, min(burden_candidate, 1.20))
    elif ratios.get("observed_labor_ratio") is not None:
        ratios["labor_burden_rate"] = 0.30

    estimated_utilization = _estimate_utilization(ratios.get("asset_turnover"), ratios.get("cogs_ratio"))
    if estimated_utilization is not None:
        ratios["estimated_utilization"] = estimated_utilization

    previous_revenue = None
    previous_operating_profit = None
    if previous_period:
        previous_revenue = _latest_statement_amount(statement_df, industry_group, previous_period, "is_revenue")
        previous_operating_profit = _latest_statement_amount(statement_df, industry_group, previous_period, "is_operating_profit", absolute=False)

    if sales_revenue and previous_revenue:
        ratios["sales_growth_yoy"] = ((float(sales_revenue) / float(previous_revenue)) - 1.0) * 100.0
    if sales_revenue and previous_revenue and operating_profit is not None and previous_operating_profit is not None:
        current_margin = float(operating_profit) / float(sales_revenue)
        previous_margin = float(previous_operating_profit) / float(previous_revenue)
        ratios["op_margin_delta_yoy"] = (current_margin - previous_margin) * 100.0

    return ratios


def _prepare_uploaded_supplier_reference(
    company_name,
    industry_group,
    uploaded_statement_files,
    business_info=None,
    statement_period_candidates=None,
    classification_meta=None,
):
    if not industry_group:
        raise ValueError("사업자등록증 OCR로 산업군을 추정하지 못했습니다. 등록증 이미지를 다시 확인해 주세요.")
    if not uploaded_statement_files:
        raise ValueError("재무제표 파일을 1개 이상 업로드해 주세요.")

    base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_tmp_uploaded_financial_refs")
    os.makedirs(base_dir, exist_ok=True)
    sanitized_classification_meta = _sanitize_classification_meta(classification_meta)
    cache_payload = {
        "company_name": str(company_name or "").strip(),
        "industry_group": str(industry_group or "").strip().upper(),
        "statement_period_candidates": sorted(str(year) for year in (statement_period_candidates or [])),
        "files": [
            {
                "statement_type": statement_type,
                "signature": _fast_upload_signature(uploaded_file),
            }
            for statement_type, uploaded_file in sorted((uploaded_statement_files or {}).items())
            if uploaded_file is not None
        ],
    }
    run_token = hashlib.sha1(
        json.dumps(cache_payload, ensure_ascii=False, sort_keys=True, default=_json_default).encode("utf-8")
    ).hexdigest()[:12]
    work_dir = os.path.join(base_dir, run_token)
    os.makedirs(work_dir, exist_ok=True)
    cache_meta_path = os.path.join(work_dir, "reference_ctx.json")
    workbook_path = os.path.join(work_dir, "uploaded_supplier_financial_db.xlsx")

    if os.path.exists(cache_meta_path) and os.path.exists(workbook_path):
        try:
            with open(cache_meta_path, "r", encoding="utf-8") as handle:
                cached_ctx = json.load(handle)
            cached_ctx["industry_reason"] = (business_info or {}).get("industry_reason", cached_ctx.get("industry_reason", ""))
            cached_ctx["industry_confidence"] = (business_info or {}).get("industry_confidence", cached_ctx.get("industry_confidence", ""))
            cached_ctx["classification_meta"] = sanitized_classification_meta
            cached_ctx["workbook_path"] = workbook_path
            return cached_ctx
        except Exception:
            pass

    saved_paths = {}
    source_names = {}
    for statement_type, uploaded_file in uploaded_statement_files.items():
        if uploaded_file is None:
            continue
        original_name = os.path.basename(getattr(uploaded_file, "name", "") or f"{statement_type.lower()}.pdf")
        ext = os.path.splitext(original_name)[1].lower() or ".pdf"
        safe_name = f"{statement_type.lower()}_{_safe_upload_token(os.path.splitext(original_name)[0])}{ext}"
        target_path = os.path.join(work_dir, safe_name)
        with open(target_path, "wb") as handle:
            handle.write(uploaded_file.getvalue())
        saved_paths[statement_type] = target_path
        source_names[statement_type] = original_name

    ensure_workbook(workbook_path)

    rows_written = 0
    restatement_detected = False
    mode = "estimated_partial"
    fallback_note = ""
    period_keys = []

    if all(statement_type in saved_paths for statement_type in FINANCIAL_REQUIRED_UPLOADS):
        ingest_result = ingest_industry_statements(
            industry_group,
            f"manual_upload_{datetime.utcnow().strftime('%Y%m%d%H%M%S')}",
            saved_paths,
            workbook_path=workbook_path,
        )
        period_keys = sorted(
            {str(period_key) for period_key in ingest_result.get("period_keys", [])},
            reverse=True,
        )
        rows_written = int(ingest_result.get("rows_written", 0) or 0)
        restatement_detected = bool(ingest_result.get("restatement_detected"))
        mode = "uploaded_observed"
        fallback_note = "재무상태표/손익계산서/현금흐름표를 모두 인식해 실측 재무기준으로 분석합니다."
    else:
        current_year = datetime.utcnow().year + 1
        candidate_years = sorted(
            {
                str(year)
                for year in (statement_period_candidates or [])
                if str(year).isdigit() and 2000 <= int(str(year)) <= current_year
            },
            reverse=True,
        )
        selected_year = next((year for year in candidate_years if year.isdigit()), str(datetime.utcnow().year - 1))
        build_estimated_industry_baseline(
            industry_group,
            selected_year,
            workbook_path=workbook_path,
        )
        period_keys = [selected_year]
        fallback_note = (
            f"부분 재무제표 업로드({', '.join(sorted(saved_paths.keys()))})로 인해 "
            "업종 추정 기반의 보수적 기준을 사용합니다."
        )

    if not period_keys:
        raise ValueError("업로드한 재무자료에서 분석 기준 연도를 찾지 못했습니다.")

    selected_period = next((key for key in period_keys if key.isdigit()), period_keys[0])
    preview_items = get_baseline_items(
        industry_group,
        selected_period,
        workbook_path=workbook_path,
    )
    financial_ratios = _extract_financial_ratios_from_ingest(
        workbook_path,
        str(industry_group).strip().upper(),
        period_keys,
        selected_period=selected_period,
    )

    reference_ctx = {
        "company_name": str(company_name or "").strip(),
        "industry_group": str(industry_group).strip().upper(),
        "industry_label": INDUSTRY_GROUP_LABELS.get(str(industry_group).strip().upper(), str(industry_group).strip().upper()),
        "industry_reason": (business_info or {}).get("industry_reason", ""),
        "industry_confidence": (business_info or {}).get("industry_confidence", ""),
        "workbook_path": workbook_path,
        "period_keys": period_keys,
        "selected_period": selected_period,
        "preview_items": preview_items,
        "uploaded_files": source_names,
        "uploaded_summary": ", ".join(f"{key}={value}" for key, value in source_names.items()),
        "rows_written": rows_written,
        "restatement_detected": restatement_detected,
        "available_statement_types": sorted(saved_paths.keys()),
        "mode": mode,
        "fallback_note": fallback_note,
        "financial_ratios": financial_ratios,
        "classification_meta": sanitized_classification_meta,
    }
    try:
        with open(cache_meta_path, "w", encoding="utf-8") as handle:
            json.dump(reference_ctx, handle, ensure_ascii=False, indent=2, default=_json_default)
    except Exception:
        pass
    return reference_ctx


def build_financial_forensic_context(
    supplier_reference_ctx: Optional[Dict[str, Any]],
    should_cost_result: Optional[Dict[str, Any]] = None,
) -> str:
    if not supplier_reference_ctx:
        return ""

    financial_ratios = supplier_reference_ctx.get("financial_ratios", {}) or {}
    mode = supplier_reference_ctx.get("mode", "estimated_partial")
    company = supplier_reference_ctx.get("company_name") or "협력사"
    industry = supplier_reference_ctx.get("industry_label", supplier_reference_ctx.get("industry_group", ""))
    period = supplier_reference_ctx.get("selected_period", "")
    quality_label = "실측 재무기준" if mode == "uploaded_observed" else "업종 추정기준(프록시)"

    lines = [
        f"[협력사 재무 분석 컨텍스트 — {company} / {industry} / {period}]",
        f"  • 데이터 품질: {quality_label}",
    ]

    ratio_lines = [
        ("observed_labor_ratio", "재무제표 관측 인건비 비율"),
        ("observed_gna_ratio", "재무제표 관측 일반관리비 비율"),
        ("observed_profit_ratio", "재무제표 관측 영업이익률"),
        ("cogs_ratio", "재무제표 관측 매출원가율"),
    ]
    for key, label in ratio_lines:
        value = _safe_float_value(financial_ratios.get(key))
        if value is not None:
            lines.append(f"  • {label}: {value * 100:.1f}%")

    sales_revenue = _safe_float_value(financial_ratios.get("sales_revenue"))
    if sales_revenue is not None:
        lines.append(f"  • 최근 매출액: {sales_revenue:,.0f}원")

    sales_growth_yoy = _safe_float_value(financial_ratios.get("sales_growth_yoy"))
    if sales_growth_yoy is not None:
        sign = "+" if sales_growth_yoy >= 0 else ""
        lines.append(f"  • 전년 대비 매출 성장률(YoY): {sign}{sales_growth_yoy:.1f}%")

    op_margin_delta_yoy = _safe_float_value(financial_ratios.get("op_margin_delta_yoy"))
    if op_margin_delta_yoy is not None:
        sign = "+" if op_margin_delta_yoy >= 0 else ""
        lines.append(f"  • 영업이익률 변화(YoY): {sign}{op_margin_delta_yoy:.1f}%p")

    asset_turnover = _safe_float_value(financial_ratios.get("asset_turnover"))
    if asset_turnover is not None:
        lines.append(f"  • 자산회전율: {asset_turnover:.2f}x")

    estimated_utilization = _safe_float_value(financial_ratios.get("estimated_utilization"))
    if estimated_utilization is not None:
        lines.append(f"  • 추정 가동률: {estimated_utilization * 100:.1f}%")

    if should_cost_result:
        utilization_gap = _safe_float_value(should_cost_result.get("utilization_gap"))
        if utilization_gap is not None and utilization_gap > 0:
            current_util = 0.85 - utilization_gap
            unit_save = float(should_cost_result.get("unit_cost_saving_pct", 0.0) or 0.0) * 100.0
            lines.append(
                f"  • 가동률 분석: 추정 가동률 {current_util * 100:.0f}%, 85% Sweet-Spot 달성 시 단위원가 {unit_save:.1f}% 절감 여력"
            )

    strategy_hints: List[str] = []
    if float(financial_ratios.get("observed_profit_ratio", 0.0) or 0.0) > 0.15:
        strategy_hints.append("영업이익률이 15%를 초과해 과도한 이윤 마진 협상 근거를 확보할 수 있습니다.")
    if float(financial_ratios.get("observed_labor_ratio", 0.0) or 0.0) > 0.65:
        strategy_hints.append("노무비 비중이 65%를 초과해 자동화 수준과 효율화 여지를 질의할 수 있습니다.")
    if _safe_float_value(financial_ratios.get("sales_growth_yoy")) is not None and float(financial_ratios.get("sales_growth_yoy", 0.0)) < -5.0:
        strategy_hints.append("매출 역성장이 뚜렷해 당사 주문 의존도 기반 가격 협상력을 확보할 수 있습니다.")
    if strategy_hints:
        lines.append("  • 재무 기반 협상 레버리지:")
        for hint in strategy_hints:
            lines.append(f"    - {hint}")

    return "\n".join(lines)


def build_quick_industry_reference_ctx(
    company_name: str,
    business_info: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    business_info = business_info or {}
    industry_group = str(business_info.get("industry_group") or "").strip().upper()
    if not industry_group:
        raise ValueError("빠른 분석을 하려면 먼저 사업자등록증을 업로드해 산업군을 자동 추정해 주세요.")

    base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_tmp_uploaded_financial_refs")
    os.makedirs(base_dir, exist_ok=True)
    work_dir = os.path.join(base_dir, "quick_analysis")
    os.makedirs(work_dir, exist_ok=True)
    workbook_path = os.path.join(work_dir, "quick_industry_baseline.xlsx")
    ensure_workbook(workbook_path)

    selected_year = str(datetime.utcnow().year - 1)
    build_estimated_industry_baseline(industry_group, selected_year, workbook_path=workbook_path)
    preview_items = get_baseline_items(industry_group, selected_year, workbook_path=workbook_path)
    financial_ratios = _extract_financial_ratios_from_ingest(
        workbook_path,
        industry_group,
        [selected_year],
        selected_period=selected_year,
    )

    return {
        "company_name": str(company_name or "").strip(),
        "industry_group": industry_group,
        "industry_label": INDUSTRY_GROUP_LABELS.get(industry_group, industry_group),
        "industry_reason": business_info.get("industry_reason", ""),
        "industry_confidence": business_info.get("industry_confidence", ""),
        "workbook_path": workbook_path,
        "period_keys": [selected_year],
        "selected_period": selected_year,
        "preview_items": preview_items,
        "uploaded_files": {},
        "uploaded_summary": "업종 기본 기준(빠른 분석)",
        "rows_written": 0,
        "restatement_detected": False,
        "available_statement_types": [],
        "mode": "quick_industry_proxy",
        "fallback_note": "재무제표 기준 없이, 사업자등록증 기반 업종 기본 기준으로 빠른 분석을 수행합니다.",
        "financial_ratios": financial_ratios,
        "classification_meta": {"quick_analysis": True},
    }


# ==============================================================================
# [NEW 핵심 패치 7] PC 계정명 불일치 및 OneDrive 경로 문제 지능형 탐색 패치 (Additive)
# ==============================================================================
_home_dir = os.path.expanduser("~")
_candidate_paths = [
    r"C:\Users\limsanghyeon.s\Desktop\EXAONE-4.0-1.2B",
    r"C:\Users\Jin Seo\OneDrive\바탕 화면\EXAONE-4.0-1.2B",
    r"C:\Users\Jin Seo\Desktop\EXAONE-4.0-1.2B",
    r"C:\Users\Jin Seo\OneDrive\바탕 화면\rnd_nego\EXAONE-4.0-1.2B",  
    os.path.join(_home_dir, "OneDrive", "바탕 화면", "EXAONE-4.0-1.2B"),
    os.path.join(_home_dir, "Desktop", "EXAONE-4.0-1.2B"),
    os.path.abspath("./EXAONE-4.0-1.2B"),
    os.path.abspath("../EXAONE-4.0-1.2B")
]

# 기본 경로를 설정하지만, 아래 main() 함수에서 UI를 통해 덮어쓸 수 있도록 합니다.
DEFAULT_MODEL_PATH = r"C:\Users\limsanghyeon.s\Desktop\EXAONE-4.0-1.2B" 
for _path in _candidate_paths:
    if os.path.exists(_path):
        DEFAULT_MODEL_PATH = _path
        break
# ==============================================================================

def get_exaone_runtime_error() -> str:
    missing_modules = [
        module_name
        for module_name in ("torch", "transformers")
        if importlib.util.find_spec(module_name) is None
    ]
    if not missing_modules:
        return ""

    module_list = ", ".join(missing_modules)
    return (
        f"EXAONE 로컬 보고서 생성에 필요한 패키지가 설치되어 있지 않습니다: {module_list}\n"
        f"설치 예시: pip install {' '.join(missing_modules)}"
    )


def safe_rerun():
    st.rerun()

@st.cache_resource
def load_exaone_local(model_path):
    """웹 페이지 로딩이 끝난 뒤, 사용자가 'AI 버튼'을 누르는 순간에만 최초 1회 로딩됩니다.
    경로를 매개변수로 받아 유연성을 극대화합니다."""
    try:
        import os
        dependency_error = get_exaone_runtime_error()
        if dependency_error:
            return None, None, dependency_error

        # 윈도우 경로(\)를 HuggingFace가 Repo ID로 오인하는 버그 원천 차단
        safe_model_path = os.path.abspath(model_path).replace('\\', '/')
        
        if not os.path.exists(safe_model_path):
            err_msg = (
                f"물리적인 AI 모델 폴더를 찾을 수 없습니다.\n\n"
                f"현재 PC에 아래 경로가 실제로 존재하는지 탐색기에서 확인해주세요:\n"
                f"👉 {safe_model_path}\n\n"
                f"(팁: EXAONE-4.0-1.2B 모델 파일들이 해당 폴더 안에 압축 해제되어 있어야 합니다.)"
            )
            return None, None, err_msg

        import torch
        from transformers import AutoTokenizer, AutoModelForCausalLM
        
        device_str = "cuda" if torch.cuda.is_available() else "cpu"
        model_dtype = torch.float16 if device_str == "cuda" else torch.float32

        tokenizer = AutoTokenizer.from_pretrained(
            safe_model_path, 
            local_files_only=True,
            trust_remote_code=True
        )
        
        model = AutoModelForCausalLM.from_pretrained(
            safe_model_path,
            local_files_only=True,
            device_map="auto" if device_str == "cuda" else None,
            torch_dtype=model_dtype,   
            low_cpu_mem_usage=True,    
            trust_remote_code=True
        )
        if device_str == "cpu":
            model.to("cpu")
            
        model.eval() 
        return tokenizer, model, device_str
    except Exception as e:
        import traceback
        return None, None, f"오류 발생: {str(e)}"

HAS_LLM = True 
device_info = "On-Demand (지연 로딩 및 실시간 스트리밍 대기 중)"
EXAONE_SYSTEM_PROMPT = """당신은 구매 재무 분석 전문가(Procurement Forensic Analyst)입니다.
## 핵심 분석 스킬
1. 재무제표 역산(Financial Reverse Engineering): BS/IS/CF 간 유기적 연관성을 이용해
   협력사의 실제 원가 구조(노무비/재료비/제조경비)를 역산합니다.
2. Should-Cost vs Will-Cost 분석: 상향식(Bottom-up) 원가 추정과 협력사 제출가의
   Gap을 계량적으로 제시합니다.
3. 가동률 Sweet-Spot(85%) 기준 What-if 시뮬레이션: 가동률 개선 시 단위원가 절감
   여력을 수치로 제시합니다.
4. TCO(총소유비용) 관점: 단순 단가가 아닌 물류/품질/결제조건을 포괄합니다.
## 출력 구조 (반드시 준수)
### 1. 비목 적정성 판정 요약
| 비목 | 협력사 제출 비중 | 재무 기준 | 판정 | 핵심 근거 |
|------|----------------|----------|------|---------|
### 2. 재무 역산 기반 Should-Cost 분석
- 협력사 재무제표에서 역산된 실제 원가 구조를 서술
- 가동률/노동부담률/COGM 역산 결과 포함
### 3. 비목별 협상 전략 (3가지 시나리오)
#### 3-1. 공격형 (Gap 최대화)
#### 3-2. 방어형 (리스크 최소화)
#### 3-3. 상생형 (JCR/장기 파트너십)
### 4. 협상 실행 스크립트 (말하기 예시)
- 구체적인 발화 예시는 반드시 큰따옴표로 제시
### 5. 비가격적 가치 확보 방안
- 결제 조건, Price Freeze, 무상 검사 등
수치는 반드시 원 단위까지 구체적으로 제시하고, 추상적 서술을 지양합니다."""

# ==============================================================================
# 저사양 PC를 위한 실시간 스트리밍 제너레이터 구현 (방화벽 타임아웃 100% 방지)
# ==============================================================================
def ask_exaone_stream(prompt_text, user_defined_model_path):
    import gc
    dependency_error = get_exaone_runtime_error()
    if dependency_error:
        yield f"⚠️ {dependency_error}"
        return

    import torch
    
    # [메모리 다이어트 1] AI 구동 직전, 메모리에 쌓인 불필요한 찌꺼기를 강제로 모두 비워냅니다 (8GB RAM 방어)
    gc.collect()
    if torch.cuda.is_available():
        torch.cuda.empty_cache()

    tokenizer, ai_model, device_info_real = load_exaone_local(user_defined_model_path)
    
    if ai_model is None:
        yield f"⚠️ AI 모델 로딩에 실패했습니다.\n\n(사유: {device_info_real})"
        return

    try:
        from transformers import TextIteratorStreamer
        import threading

        # [하드웨어 스로틀링] 저사양 CPU 환경에서 PC 전체가 다운되는 현상 방지
        is_cpu_generation = ai_model.device.type == 'cpu'
        if is_cpu_generation:
            torch.set_num_threads(2)

        messages = [
            {"role": "system", "content": EXAONE_SYSTEM_PROMPT},
            {"role": "user", "content": prompt_text}
        ]
        
        formatted_prompt = tokenizer.apply_chat_template(
            messages, 
            tokenize=False, 
            add_generation_prompt=True
        )
        
        inputs = tokenizer(formatted_prompt, return_tensors="pt").to(ai_model.device)
        
        # 스트리머 생성: 단어가 하나 생성될 때마다 바로바로 뱉어냄(Yield)
        streamer = TextIteratorStreamer(tokenizer, skip_prompt=True, skip_special_tokens=True)

        generation_kwargs = dict(
            input_ids=inputs["input_ids"], 
            attention_mask=inputs.get("attention_mask", None),
            max_new_tokens=1024 if is_cpu_generation else 1536,
            eos_token_id=tokenizer.eos_token_id,
            pad_token_id=tokenizer.eos_token_id,
            do_sample=not is_cpu_generation,
            use_cache=True,
            repetition_penalty=1.03,
            streamer=streamer # 스트리머 연결
        )
        if not is_cpu_generation:
            generation_kwargs["temperature"] = 0.7
            generation_kwargs["top_p"] = 0.9
        
        # [메모리 다이어트 2] torch.inference_mode()를 통해 AI 생성 시 메모리를 잡아먹는 백그라운드 추적 기능을 완벽히 끕니다.
        def _generate_with_mem_clear():
            with torch.inference_mode(): 
                ai_model.generate(**generation_kwargs)
            # 연산 완료 후 다시 한 번 메모리 강제 반환
            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

        thread = threading.Thread(target=_generate_with_mem_clear)
        thread.start()
        
        # 생성되는 단어를 즉시 UI로 Yield (방화벽이 계속 데이터가 오가는 것으로 인식하여 절대로 안 끊음)
        for new_text in streamer:
            yield new_text
            
    except Exception as e:
        import traceback
        yield f"\n\n⚠️ 답변 생성 중 오류가 발생했습니다:\n{traceback.format_exc()}"

# --------------------------------------------------------------------------------
# [1] 데이터 무결성 검증 엔진
# --------------------------------------------------------------------------------
class DataValidator:
    @staticmethod
    def validate_estimate_data(df: pd.DataFrame) -> List[str]:
        errors = []
        col_item = next((c for c in df.columns if '비목' in c or '품목' in c), None)
        col_price = next((c for c in df.columns if '단가' in c or 'Price' in c), None)
        col_qty = next((c for c in df.columns if '수량' in c or 'Qty' in c), None)

        if not (col_item and col_price and col_qty):
            return ["❌ 엑셀 파일에 [비목, 단가, 수량] 컬럼이 포함되어야 합니다."]

        df.rename(columns={col_item:'비목', col_price:'단가', col_qty:'수량'}, inplace=True)
        
        if not pd.api.types.is_numeric_dtype(df['단가']): errors.append("⚠️ '단가' 컬럼에 숫자가 아닌 값이 있습니다.")
        if not pd.api.types.is_numeric_dtype(df['수량']): errors.append("⚠️ '수량' 컬럼에 숫자가 아닌 값이 있습니다.")

        return errors

# --------------------------------------------------------------------------------
# [2] 검색 엔진
# --------------------------------------------------------------------------------
class LocalHybridSearchEngine:
    def __init__(self, df_knowledge_base: pd.DataFrame):
        self.df = df_knowledge_base.copy()
        self.df = self.df.fillna("")
        self._recent_fingerprints: List[str] = []
        
    def _normalize(self, text):
        if not isinstance(text, str):
            return ""
        text = text.lower()
        text = re.sub(r"[^0-9a-zA-Z\u3131-\u318E\uAC00-\uD7A3\s]", " ", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    def _tokenize(self, text):
        norm = self._normalize(text)
        if not norm:
            return []
        return [tok for tok in norm.split(" ") if len(tok) >= 2]

    def _strategy_fingerprint(self, row: Dict[str, Any], keys: List[str]) -> str:
        # Stable key for the strategy triplet to suppress repeated recommendations.
        parts = [self._normalize(str(row.get(k, ""))) for k in keys]
        return "||".join(parts)

    def _diversity_rerank(self, candidates: List[Dict], top_k: int, per_fingerprint_limit: int = 1) -> List[Dict]:
        if not candidates:
            return []
        primary = []
        overflow = []
        seen = {}
        for c in candidates:
            fp = c.get("fingerprint", "")
            used = seen.get(fp, 0)
            if used < per_fingerprint_limit:
                primary.append(c)
                seen[fp] = used + 1
            else:
                overflow.append(c)
        return (primary + overflow)[:top_k]

    def _novelty_rerank(self, candidates: List[Dict]) -> List[Dict]:
        if not candidates or not self._recent_fingerprints:
            return candidates
        recent = set(self._recent_fingerprints)
        fresh = [c for c in candidates if c.get("fingerprint", "") not in recent]
        stale = [c for c in candidates if c.get("fingerprint", "") in recent]
        return fresh + stale

    def clear_fingerprints(self):
        self._recent_fingerprints = []

    def _remember_fingerprint(self, fingerprint: str, max_recent: int = 50):
        if not fingerprint:
            return
        if fingerprint in self._recent_fingerprints:
            self._recent_fingerprints.remove(fingerprint)
        self._recent_fingerprints.insert(0, fingerprint)
        if len(self._recent_fingerprints) > max_recent:
            self._recent_fingerprints = self._recent_fingerprints[:max_recent]

    def search_strategy(self, query: str, category_filter: str = None, top_k: int = 1) -> List[Dict]:
        if not query or self.df.empty:
            return []

        norm_query = self._normalize(query)
        query_tokens = set(self._tokenize(query))
        if not query_tokens and norm_query:
            query_tokens = set(norm_query.split())

        candidates = []
        t_key = next((c for c in self.df.columns if 'Trigger' in c), 'Trigger_Keywords')
        r_key = next((c for c in self.df.columns if 'Risk' in c), 'Risk_Type')
        c_key = next((c for c in self.df.columns if 'Category' in c), 'Category')
        a_key = next((c for c in self.df.columns if 'Aggressive' in c), 'Mode_Aggressive')
        d_key = next((c for c in self.df.columns if 'Defensive' in c), 'Mode_Defensive')
        w_key = next((c for c in self.df.columns if 'WinWin' in c), 'Mode_WinWin')

        for idx, row in self.df.iterrows():
            score = 0

            row_category = str(row.get(c_key, ''))
            if category_filter and category_filter.lower() in row_category.lower():
                score += 50

            trigger_text = str(row.get(t_key, ''))
            risk_text = str(row.get(r_key, ''))
            mode_agg = str(row.get(a_key, ''))
            mode_def = str(row.get(d_key, ''))
            mode_win = str(row.get(w_key, ''))

            db_keywords = [self._normalize(k) for k in re.split(r'[,|\n]', trigger_text) if str(k).strip()]
            search_blob = " ".join([trigger_text, risk_text, mode_agg, mode_def, mode_win])
            norm_blob = self._normalize(search_blob)
            blob_tokens = set(self._tokenize(search_blob))

            matched_cnt = 0
            for kw in db_keywords:
                if not kw:
                    continue
                if kw in query_tokens:
                    score += 40
                    matched_cnt += 1
                elif kw in norm_query:
                    score += 30
                    matched_cnt += 1

            overlap_tokens = query_tokens.intersection(blob_tokens)
            if overlap_tokens:
                # Context matching across risk/script text to improve query-specific ranking.
                score += len(overlap_tokens) * 12
                matched_cnt += len(overlap_tokens)

            norm_risk = self._normalize(risk_text)
            if norm_risk and norm_risk in norm_query:
                score += 25
                matched_cnt += 1

            if HAS_NLP_LIB and fuzz:
                similarity = fuzz.token_set_ratio(norm_query, norm_blob)
                if similarity >= 80:
                    score += 20
                    matched_cnt += 1

            if score >= 30:
                row_dict = row.to_dict()
                fingerprint = self._strategy_fingerprint(row_dict, [a_key, d_key, w_key])
                candidates.append({
                    'index': idx,
                    'score': score,
                    'data': row_dict,
                    'match_count': matched_cnt,
                    'risk_type': row.get(r_key, ''),
                    'fingerprint': fingerprint
                })

        candidates.sort(key=lambda x: (x['score'], x['match_count']), reverse=True)
        novelty_candidates = self._novelty_rerank(candidates)
        ranked_candidates = self._diversity_rerank(novelty_candidates, top_k=top_k, per_fingerprint_limit=1)
        if ranked_candidates:
            self._remember_fingerprint(ranked_candidates[0].get("fingerprint", ""))
        results = [c['data'] for c in ranked_candidates]
        return results

# --------------------------------------------------------------------------------
# [3] 분석 엔진
# --------------------------------------------------------------------------------
DB_FILENAME = "standard_db.xlsx"

class CategoryAnalysisEngine:
    def __init__(self, db_path):
        self.df_price_db = load_db_data(db_path, 'UNIT_PRICE')
        self.df_ratio_db = load_db_data(db_path, 'INDIRECT_RATIO')
        self.standard_categories = self._collect_standard_categories()

    def _safe_number(self, value, default=0.0):
        try:
            if pd.isna(value):
                return default
            if isinstance(value, str):
                cleaned = value.replace('%', '').replace(',', '').strip()
                if cleaned == "":
                    return default
                return float(cleaned)
            return float(value)
        except Exception:
            return default

    def _normalize_text(self, value):
        return re.sub(r'\s+', '', str(value or '').strip().lower())

    def _standard_category_key(self, value):
        text = self._normalize_text(value)
        if not text:
            return ""
        return re.sub(r'^\d+(?:-\d+)?\.?', '', text)

    def _find_col(self, df, candidates):
        normalized = {self._normalize_text(col): col for col in df.columns}
        for candidate in candidates:
            key = self._normalize_text(candidate)
            if key in normalized:
                return normalized[key]
        for candidate in candidates:
            key = self._normalize_text(candidate)
            for norm_col, raw_col in normalized.items():
                if key in norm_col:
                    return raw_col
        return None

    def _fallback_standard_category_legacy(self, ui_category):
        if str(ui_category).startswith("1."):
            return "1.일반용역"
        if str(ui_category).startswith("2."):
            return "2.SW용역"
        if str(ui_category).startswith("3."):
            return "3.공사"
        return str(ui_category or "").strip()

    def get_standard_categories(self):
        return [row["label"] for row in STANDARD_CATEGORY_GROUPS]

    def _category_group_id(self, category_value):
        category = self._get_preferred_category(category_value)
        text = str(category or category_value or "").strip()
        if text.startswith("3-") or text.startswith("3."):
            return "construction"
        if text.startswith("2."):
            return "sw"
        if text.startswith("4."):
            return "marketing"
        if text.startswith("5."):
            return "research"
        return "general"

    def _category_group_label(self, category_value):
        return STANDARD_CATEGORY_GROUP_LABELS.get(self._category_group_id(category_value), "일반용역")

    def get_standard_category_groups(self):
        groups = []
        for row in STANDARD_CATEGORY_GROUPS:
            if self.get_standard_categories_for_group(row["id"]):
                groups.append(row["id"])
        return groups

    def get_standard_categories_for_group(self, group_id):
        return [category for category in self.standard_categories if self._category_group_id(category) == group_id]

    def get_default_standard_category_for_group(self, group_id):
        categories = self.get_standard_categories_for_group(group_id)
        if categories:
            return categories[0]
        return self._fallback_standard_category(None)

    def _category_sort_key(self, value):
        text = str(value or "").strip()
        match = re.match(r"(\d+)(?:-(\d+))?", text)
        if match:
            return (int(match.group(1)), int(match.group(2) or 0), text)
        return (999, 999, text)

    def _collect_standard_categories(self):
        categories = []
        for df in (self.df_price_db, self.df_ratio_db):
            if df.empty:
                continue
            col_category = self._find_col(df, ['Category', '카테고리'])
            if not col_category:
                continue
            for raw in df[col_category].dropna().astype(str):
                value = raw.strip()
                if value:
                    categories.append(value)

        categories = list(dict.fromkeys(categories or DEFAULT_STANDARD_CATEGORIES))
        for default_category in DEFAULT_STANDARD_CATEGORIES:
            if default_category not in categories:
                categories.append(default_category)
        return sorted(categories, key=self._category_sort_key)

    def _get_first_category_with_prefix(self, prefix):
        for category in self.standard_categories:
            if str(category).startswith(prefix):
                return category
        return ""

    def _match_known_standard_category(self, raw_value):
        raw = str(raw_value or "").strip()
        if not raw:
            return ""

        raw_norm = self._normalize_text(raw)
        if not raw_norm:
            return ""

        normalized_map = {
            self._normalize_text(category): category
            for category in self.standard_categories
        }
        if raw_norm in normalized_map:
            return normalized_map[raw_norm]

        for category in self.standard_categories:
            category_norm = self._normalize_text(category)
            if raw_norm == category_norm.split('(')[0]:
                return category
            if raw_norm in category_norm or category_norm in raw_norm:
                return category
        return ""

    def _get_preferred_category(self, target):
        exact = self._match_known_standard_category(target)
        if exact:
            return exact

        text = self._normalize_text(target)
        if text.startswith("3-3") or "전기" in text or "통신" in text or "약전" in text:
            return self._get_first_category_with_prefix("3-3.") or self._get_first_category_with_prefix("3-3") or self._get_first_category_with_prefix("3-")
        if text.startswith("3-2") or "인테리어" in text or "전문" in text:
            return self._get_first_category_with_prefix("3-2.") or self._get_first_category_with_prefix("3-2") or self._get_first_category_with_prefix("3-")
        if text.startswith("3") or "공사" in text or "건축" in text or "시공" in text:
            return self._get_first_category_with_prefix("3-1.") or self._get_first_category_with_prefix("3-1") or self._get_first_category_with_prefix("3-")
        if text.startswith("2") or "sw" in text or "개발" in text or "소프트웨어" in text:
            return self._get_first_category_with_prefix("2.") or "2.SW용역"
        if text.startswith("4") or "마케팅" in text or "광고" in text:
            return self._get_first_category_with_prefix("4.") or "4.마케팅"
        if text.startswith("5") or "연구" in text or "r&d" in text or "rnd" in text:
            return self._get_first_category_with_prefix("5.") or "5.연구용역"
        if text.startswith("1") or "일반" in text or "용역" in text:
            return self._get_first_category_with_prefix("1.") or "1.일반용역"
        return self._get_first_category_with_prefix("1.") or DEFAULT_STANDARD_CATEGORIES[0]

    def _fallback_standard_category(self, ui_category):
        return self._get_preferred_category(ui_category or DEFAULT_STANDARD_CATEGORIES[0])

    def _strategy_category(self, category_value):
        category = self._get_preferred_category(category_value)
        if str(category).startswith("3-") or str(category).startswith("3."):
            return "3.공사"
        if str(category).startswith("2."):
            return "2.SW용역"
        if str(category).startswith("4."):
            return "4.마케팅"
        if str(category).startswith("5."):
            return "5.연구용역"
        return "1.일반용역"

    def _infer_standard_category_from_text(self, text):
        normalized = self._normalize_text(text)
        rules = [
            ("3-3.전기/통신", ["전기", "통신", "약전", "배선", "cctv", "조명", "분전반", "네트워크케이블"]),
            ("3-2.전문/인테리어", ["인테리어", "철거", "도장", "방수", "타일", "목공", "창호", "도배"]),
            ("3-1.종합건축", ["건축", "공사", "시공", "토목", "배관", "설비", "조경", "골조"]),
            ("2.SW용역", ["sw", "소프트웨어", "개발", "프로그램", "앱", "웹", "시스템", "서버", "데이터", "db", "네트워크", "유지보수"]),
            ("4.마케팅", ["마케팅", "광고", "홍보", "브랜딩", "프로모션", "캠페인", "콘텐츠", "sns"]),
            ("5.연구용역", ["연구", "r&d", "rnd", "조사", "실험", "검증", "타당성", "컨설팅"]),
        ]
        for target_category, keywords in rules:
            if any(keyword in normalized for keyword in keywords):
                return self._get_preferred_category(target_category)
        return self._fallback_standard_category(None)

    def _detect_standard_category_context(self, df_input, category_override=None):
        df = df_input.copy()
        if '금액' not in df.columns:
            df['금액'] = pd.to_numeric(df.get('단가', 0), errors='coerce').fillna(0) * pd.to_numeric(df.get('수량', 0), errors='coerce').fillna(0)

        col_category = next((col for col in df.columns if str(col) == '카테고리'), None)
        col_item = next((col for col in df.columns if str(col) == '비목'), None)
        col_detail = next((col for col in df.columns if str(col) in ['품명(상세내역)', '내역(상세내역)', '상세내역', '내역', '품명']), None)
        col_grade = next((col for col in df.columns if str(col) == '등급'), None)
        col_unit = next((col for col in df.columns if str(col) == '단위'), None)

        category_stats = {}
        row_categories = {}
        row_reasons = {}
        explicit_rows = 0
        inferred_rows = 0
        mode = "수동 지정" if category_override else "자동 인식"

        for idx, row in df.iterrows():
            actual_amount = self._safe_number(row.get('금액'), 0.0)
            raw_category = str(row.get(col_category, '') if col_category else '').strip()
            if category_override:
                resolved_category = self._get_preferred_category(category_override)
                reason = "수동 지정"
            else:
                resolved_category = self._match_known_standard_category(raw_category)
                if resolved_category:
                    reason = "행 카테고리"
                    explicit_rows += 1
                else:
                    hint_text = " ".join(
                        str(row.get(col, '') or '').strip()
                        for col in [col_category, col_item, col_detail, col_grade, col_unit]
                        if col
                    )
                    resolved_category = self._infer_standard_category_from_text(hint_text)
                    reason = "자동 추정"
                    inferred_rows += 1

            bucket = category_stats.setdefault(
                resolved_category,
                {"amount": 0.0, "row_count": 0, "reason_counts": {}},
            )
            bucket["amount"] += actual_amount
            bucket["row_count"] += 1
            bucket["reason_counts"][reason] = bucket["reason_counts"].get(reason, 0) + 1
            row_categories[idx] = resolved_category
            row_reasons[idx] = reason

        if not category_stats:
            default_category = self._fallback_standard_category(category_override)
            category_stats[default_category] = {"amount": 0.0, "row_count": 0, "reason_counts": {mode: 1}}

        sorted_mix = sorted(
            category_stats.items(),
            key=lambda item: (item[1]["amount"], item[1]["row_count"], item[0]),
            reverse=True,
        )
        total_amount = sum(info["amount"] for _, info in sorted_mix)
        total_rows = sum(info["row_count"] for _, info in sorted_mix) or 1

        mix_rows = []
        for category, info in sorted_mix:
            share_pct = (info["amount"] / total_amount * 100.0) if total_amount > 0 else (info["row_count"] / total_rows * 100.0)
            mix_rows.append(
                {
                    "category": category,
                    "group_id": self._category_group_id(category),
                    "group_label": self._category_group_label(category),
                    "amount": float(info["amount"]),
                    "row_count": int(info["row_count"]),
                    "share_pct": float(share_pct),
                    "reason": max(info["reason_counts"], key=info["reason_counts"].get),
                }
            )

        primary_category = mix_rows[0]["category"]
        primary_group_label = self._category_group_label(primary_category)
        group_mix_map = {}
        for row in mix_rows:
            bucket = group_mix_map.setdefault(
                row["group_id"],
                {
                    "group_id": row["group_id"],
                    "group_label": row["group_label"],
                    "amount": 0.0,
                    "row_count": 0,
                },
            )
            bucket["amount"] += float(row["amount"])
            bucket["row_count"] += int(row["row_count"])
        group_mix_rows = sorted(
            [
                {
                    **value,
                    "share_pct": (value["amount"] / total_amount * 100.0) if total_amount > 0 else (value["row_count"] / total_rows * 100.0),
                }
                for value in group_mix_map.values()
            ],
            key=lambda row: (row["amount"], row["row_count"], row["group_label"]),
            reverse=True,
        )
        mix_summary = ", ".join(
            f"{row['group_label']} {row['share_pct']:.1f}%"
            for row in group_mix_rows[:3]
        )
        summary = f"{mode} 원가 체계: {mix_summary or primary_category}"
        if not category_override and (explicit_rows or inferred_rows):
            summary += f" (행 카테고리 {explicit_rows}건, 자동 추정 {inferred_rows}건)"

        summary = f"{mode} 원가 체계: {mix_summary or primary_group_label}"
        if not category_override and (explicit_rows or inferred_rows):
            summary += f" (직접 카테고리 {explicit_rows}건, 자동 추정 {inferred_rows}건)"

        summary = f"{mode} 원가 체계: {mix_summary or primary_group_label}"
        if not category_override and (explicit_rows or inferred_rows):
            summary += f" (직접 카테고리 {explicit_rows}건, 자동 추정 {inferred_rows}건)"

        return {
            "mode": mode,
            "primary_category": primary_category,
            "primary_group": self._category_group_id(primary_category),
            "primary_group_label": primary_group_label,
            "strategy_category": self._strategy_category(primary_category),
            "summary": summary,
            "mix_rows": mix_rows,
            "group_mix_rows": group_mix_rows,
            "row_categories": row_categories,
            "row_reasons": row_reasons,
        }

    def _canonical_unit(self, value):
        text = self._normalize_text(value)
        aliases = {
            "시간(hr)": "hr",
            "시간": "hr",
            "hr": "hr",
            "hour": "hr",
            "hours": "hr",
            "h": "hr",
            "일(day)": "day",
            "일": "day",
            "day": "day",
            "days": "day",
            "월(month)": "month",
            "월": "month",
            "month": "month",
            "months": "month",
            "연(year)": "year",
            "연": "year",
            "년": "year",
            "year": "year",
            "years": "year",
            "ea": "ea",
            "pcs": "pcs",
            "pc": "pcs",
            "lot": "lot",
            "set": "set",
            "box": "box",
            "roll": "roll",
            "can": "can",
            "drum": "drum",
            "kg": "kg",
            "m": "m",
            "m2": "m2",
            "㎡": "m2",
            "m3": "m3",
            "㎥": "m3",
        }
        return aliases.get(text, text)

    def _normalize_item_name(self, value):
        text = self._normalize_text(value)
        replacements = {
            "소프트웨어": "sw",
            "프로젝트매니저": "pm",
            "projectmanager": "pm",
            "프로그램매니저": "pm",
            "엔지니어링": "엔지니어",
            "개발자": "개발",
            "프로그래머": "개발",
            "디자이너": "디자인",
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub(r'[^0-9a-z가-힣]', '', text)

    def _name_similarity_score(self, quote_name, standard_name):
        quote_norm = self._normalize_item_name(quote_name)
        standard_norm = self._normalize_item_name(standard_name)
        if not quote_norm or not standard_norm:
            return 0.0
        if quote_norm == standard_norm:
            return 1.0
        if quote_norm in standard_norm or standard_norm in quote_norm:
            return 0.93
        score = difflib.SequenceMatcher(None, quote_norm, standard_norm).ratio()
        if HAS_NLP_LIB and fuzz:
            try:
                score = max(
                    score,
                    float(fuzz.token_set_ratio(quote_norm, standard_norm)) / 100.0,
                    float(fuzz.partial_ratio(quote_norm, standard_norm)) / 100.0,
                )
            except Exception:
                pass
        return score

    def _convert_quantity_to_standard_unit(self, qty, quote_unit, standard_unit):
        quote_canonical = self._canonical_unit(quote_unit)
        standard_canonical = self._canonical_unit(standard_unit)
        if qty is None:
            return None, "수량 없음"
        if not quote_canonical or not standard_canonical:
            return None, "단위 정보 부족"
        if quote_canonical == standard_canonical:
            return float(qty), f"단위 일치({standard_canonical})"

        time_factors = {
            "hr": 1.0,
            "day": 8.0,
            "month": 176.0,
            "year": 2112.0,
        }
        if quote_canonical in time_factors and standard_canonical in time_factors:
            converted = float(qty) * time_factors[quote_canonical] / time_factors[standard_canonical]
            return converted, f"단위 환산({quote_canonical}->{standard_canonical})"
        return None, f"단위 미매칭({quote_unit}->{standard_unit})"

    def _filter_standard_category(self, df, category_value):
        if df.empty:
            return df
        col_category = self._find_col(df, ['Category', '카테고리'])
        if not col_category:
            return df
        category_norm = self._standard_category_key(category_value)
        if not category_norm:
            return df
        category_series = df[col_category].fillna("").astype(str)
        category_keys = category_series.map(self._standard_category_key)
        exact = df[category_keys == category_norm]
        if not exact.empty:
            return exact
        contains = df[category_keys.str.contains(re.escape(category_norm), na=False)]
        if not contains.empty:
            return contains
        # Do not widen to the whole DB when category matching fails.
        return df.iloc[0:0].copy()

    def _pick_unit_price(self, df_group, item_kind, item_name, grade, unit, qty):
        if df_group.empty:
            return None, None, "표준단가 DB 매칭 실패"

        col_item_name = self._find_col(df_group, ['Item_Name', '품명', '상세내역'])
        col_grade = self._find_col(df_group, ['Grade', '등급'])
        col_unit = self._find_col(df_group, ['Unit', '단위'])
        col_price = self._find_col(df_group, ['Price', '단가'])

        if not col_price:
            return None, None, "표준단가 컬럼 없음"

        working = df_group.copy()
        match_steps = []
        item_norm = self._normalize_item_name(item_name)
        grade_norm = self._normalize_text(grade)
        unit_norm = self._canonical_unit(unit)

        if col_grade and grade_norm:
            exact = working[working[col_grade].fillna("").astype(str).map(self._normalize_text) == grade_norm]
            if exact.empty:
                return None, None, "등급 미매칭"
            working = exact
            match_steps.append("등급")

        if working.empty:
            return None, None, "표준단가 후보 없음"

        converted_candidates = []
        if col_unit:
            if not unit_norm:
                return None, None, "단위 미입력"
            for _, candidate in working.iterrows():
                standard_unit = str(candidate.get(col_unit, "") or "").strip()
                converted_qty, conversion_note = self._convert_quantity_to_standard_unit(qty, unit, standard_unit)
                if converted_qty is None:
                    continue
                converted_candidates.append(
                    {
                        "row": candidate,
                        "converted_qty": converted_qty,
                        "conversion_note": conversion_note,
                        "standard_unit": standard_unit,
                    }
                )
            if not converted_candidates:
                return None, None, "단위 미매칭"

            exact_candidates = [
                candidate
                for candidate in converted_candidates
                if self._canonical_unit(candidate["standard_unit"]) == unit_norm
            ]
            if exact_candidates:
                converted_candidates = exact_candidates
                match_steps.append("단위 일치")
            else:
                match_steps.append("단위 환산")

            if col_item_name and item_norm:
                for candidate in converted_candidates:
                    candidate["name_score"] = self._name_similarity_score(item_name, candidate["row"].get(col_item_name, ""))
                best_score = max(candidate["name_score"] for candidate in converted_candidates)
                best_candidates = [candidate for candidate in converted_candidates if candidate["name_score"] == best_score]

                if best_score >= 0.93:
                    converted_candidates = best_candidates
                    match_steps.append(f"항목명 유사도 {best_score:.2f}")
                elif item_kind == 'labor' and best_score >= 0.72:
                    converted_candidates = best_candidates
                    match_steps.append(f"항목명 근사매칭 {best_score:.2f}")
                elif item_kind == 'material':
                    return None, None, f"품명/상세내역 미매칭({best_score:.2f})"

            canonical_units = {
                self._canonical_unit(candidate["standard_unit"])
                for candidate in converted_candidates
            }

            if len(converted_candidates) > 1:
                if item_kind == 'labor' and len(canonical_units) == 1:
                    price_series = pd.to_numeric(
                        pd.Series([candidate["row"].get(col_price) for candidate in converted_candidates]),
                        errors='coerce',
                    ).dropna()
                    if price_series.empty:
                        return None, None, "대표단가 계산 실패"
                    representative_price = float(price_series.median())
                    representative_qty = float(converted_candidates[0]["converted_qty"])
                    note = ", ".join(match_steps + ["대표단가(중앙값)"]) if match_steps else "대표단가(중앙값)"
                    return representative_price, representative_qty, note
                price_values = pd.to_numeric(
                    pd.Series([candidate["row"].get(col_price) for candidate in converted_candidates]),
                    errors='coerce',
                ).dropna().unique()
                if len(price_values) != 1:
                    return None, None, "복수 표준단가 후보"

            selected = converted_candidates[0]
            price = self._safe_number(selected["row"].get(col_price), None)
            if price is None:
                return None, None, "표준단가 숫자 없음"
            note = ", ".join(match_steps + [selected["conversion_note"]]) if match_steps else selected["conversion_note"]
            return float(price), float(selected["converted_qty"]), note

        prices = pd.to_numeric(working[col_price], errors='coerce').dropna().unique()
        if len(prices) == 1:
            note = ", ".join(match_steps) if match_steps else "카테고리 fallback"
            return float(prices[0]), float(qty), note
        if item_kind == 'labor' and len(prices) > 1:
            representative_price = float(pd.Series(prices).median())
            note = ", ".join(match_steps + ["대표단가(중앙값)"]) if match_steps else "대표단가(중앙값)"
            return representative_price, float(qty), note
        return None, None, "복수 표준단가 후보"

    def _pick_ratio_row(self, category_value, cost_item):
        if self.df_ratio_db.empty:
            return None
        working = self._filter_standard_category(self.df_ratio_db, category_value)
        if working.empty:
            return None
        col_cost_item = self._find_col(working, ['Cost_item', '비목', '항목'])
        if not col_cost_item:
            return None
        col_min = self._find_col(working, ['Min_Rate', '최소'])
        col_max = self._find_col(working, ['Max_Rate', '최대'])
        col_rate_target = self._find_col(working, ['Target_Rate', 'Base_Rate', '기준비율', '목표비율'])
        col_base_target = self._find_col(working, ['Base_Target', '기준', '목표'])

        aliases = {
            '경비': ['경비', '제경비'],
            '일반관리비': ['일반관리비'],
            '이윤': ['이윤', '기술료'],
        }
        candidates = aliases.get(cost_item, [cost_item])
        match = pd.DataFrame()
        cost_series = working[col_cost_item].fillna("").astype(str)
        cost_norm = cost_series.map(self._normalize_text)
        for candidate in candidates:
            key = self._normalize_text(candidate)
            exact = working[cost_norm == key]
            if not exact.empty:
                match = exact
                break
        if match.empty:
            for candidate in candidates:
                key = self._normalize_text(candidate)
                contains = working[cost_norm.str.contains(key, na=False)]
                if not contains.empty:
                    match = contains
                    break
        if match.empty:
            return None
        row = match.iloc[0]
        rate = self._safe_number(row.get(col_rate_target), None) if col_rate_target else None
        min_rate = self._safe_number(row.get(col_min), None) if col_min else None
        max_rate = self._safe_number(row.get(col_max), None) if col_max else None
        if rate is None and min_rate is not None and max_rate is not None:
            rate = (float(min_rate) + float(max_rate)) / 2.0
        if rate is None:
            rate = max_rate
        if rate is None:
            rate = min_rate if min_rate is not None else 0.0
        return {
            'rate': float(rate),
            'base_target': str(row.get(col_base_target, '') or ''),
            'cost_item': str(row.get(col_cost_item, cost_item) or cost_item),
        }

    def _resolve_ratio_base(self, cost_item, base_target_text, amounts):
        text = self._normalize_text(base_target_text)
        total = 0.0
        if any(keyword in text for keyword in ['재료', '자재', '원재료']):
            total += amounts.get('material', 0.0)
        if any(keyword in text for keyword in ['노무', '인건', '직접노무', '직접인건']):
            total += amounts.get('labor', 0.0)
        if any(keyword in text for keyword in ['경비', '제경비']):
            total += amounts.get('expense', 0.0)
        if any(keyword in text for keyword in ['일반관리', '관리비']):
            total += amounts.get('gna', 0.0)

        if total > 0:
            return total

        if cost_item == '경비':
            return amounts.get('labor', 0.0)
        if cost_item == '일반관리비':
            return amounts.get('material', 0.0) + amounts.get('labor', 0.0) + amounts.get('expense', 0.0)
        if cost_item == '이윤':
            return amounts.get('material', 0.0) + amounts.get('labor', 0.0) + amounts.get('expense', 0.0) + amounts.get('gna', 0.0)
        return 0.0

    def _build_standard_reference_legacy(self, df_input, ui_category):
        df = df_input.copy()
        if '금액' not in df.columns:
            df['금액'] = pd.to_numeric(df['단가'], errors='coerce').fillna(0) * pd.to_numeric(df['수량'], errors='coerce').fillna(0)

        col_category = next((col for col in df.columns if str(col) == '카테고리'), None)
        col_item = next((col for col in df.columns if str(col) == '비목'), None)
        col_detail = next((col for col in df.columns if str(col) in ['품명(상세내역)', '내역(상세내역)', '상세내역', '내역', '품명']), None)
        col_grade = next((col for col in df.columns if str(col) == '등급'), None)
        col_unit = next((col for col in df.columns if str(col) == '단위'), None)

        category_buckets = {}
        notes = []
        matched_rows = 0
        fallback_rows = 0
        excluded_rows = 0

        for _, row in df.iterrows():
            item_text = str(row.get(col_item, '') if col_item else '').strip()
            if not item_text:
                continue

            item_kind = None
            if re.search(r'노무|인건|급여|인력', item_text):
                item_kind = 'labor'
            elif re.search(r'재료|자재|원재료|부품', item_text):
                item_kind = 'material'
            if not item_kind:
                continue

            category_value = str(row.get(col_category, '') if col_category else '').strip() or self._fallback_standard_category(ui_category)
            detail_value = str(row.get(col_detail, '') if col_detail else '').strip()
            grade_value = str(row.get(col_grade, '') if col_grade else '').strip()
            unit_value = str(row.get(col_unit, '') if col_unit else '').strip()
            qty = self._safe_number(row.get('수량'), 0.0)
            actual_amount = self._safe_number(row.get('금액'), 0.0)

            price_group = self._filter_standard_category(self.df_price_db, category_value)
            std_price, converted_qty, match_note = self._pick_unit_price(price_group, item_kind, detail_value, grade_value, unit_value, qty)
            if std_price is None:
                if item_kind == 'material' and actual_amount > 0:
                    ref_amount = actual_amount
                    fallback_rows += 1
                    if len(notes) < 8:
                        notes.append(f"{item_text}: {match_note}, 재료비는 제출 금액으로 반영")
                else:
                    ref_amount = 0.0
                    excluded_rows += 1
                    if len(notes) < 8:
                        notes.append(f"{item_text}: {match_note}, 공시기준가 산정에서 제외")
            else:
                ref_amount = std_price * converted_qty
                matched_rows += 1
                if len(notes) < 6:
                    notes.append(f"{item_text}: 표준단가 {std_price:,.0f} × 환산수량 {converted_qty:,.2f} ({match_note})")

            bucket_key = category_value or self._fallback_standard_category(ui_category)
            if bucket_key not in category_buckets:
                category_buckets[bucket_key] = {'labor': 0.0, 'material': 0.0}
            category_buckets[bucket_key][item_kind] += ref_amount

        totals = {'labor': 0.0, 'material': 0.0, 'expense': 0.0, 'gna': 0.0, 'profit': 0.0, 'rnd': 0.0}
        for category_value, direct_amounts in category_buckets.items():
            subtotal = {
                'labor': direct_amounts.get('labor', 0.0),
                'material': direct_amounts.get('material', 0.0),
                'expense': 0.0,
                'gna': 0.0,
                'profit': 0.0,
                'rnd': 0.0,
            }
            for cost_item, code in [('경비', 'expense'), ('일반관리비', 'gna'), ('이윤', 'profit')]:
                ratio_row = self._pick_ratio_row(category_value, cost_item)
                if not ratio_row:
                    continue
                base_amount = self._resolve_ratio_base(cost_item, ratio_row['base_target'], subtotal)
                subtotal[code] = base_amount * (ratio_row['rate'] / 100.0)
            for code in totals:
                totals[code] += subtotal.get(code, 0.0)

        totals['total'] = sum(totals.values())
        coverage = f"표준단가 매칭 {matched_rows}건 / fallback {fallback_rows}건 / 제외 {excluded_rows}건"
        return {
            'amounts': totals,
            'coverage': coverage,
            'notes': notes,
        }
        
    def _build_standard_reference(self, df_input, category_context=None, category_override=None):
        df = df_input.copy()
        if '금액' not in df.columns:
            df['금액'] = pd.to_numeric(df['단가'], errors='coerce').fillna(0) * pd.to_numeric(df['수량'], errors='coerce').fillna(0)

        col_item = next((col for col in df.columns if str(col) == '비목'), None)
        col_detail = next((col for col in df.columns if str(col) in ['품명(상세내역)', '내역(상세내역)', '상세내역', '내역', '품명']), None)
        col_grade = next((col for col in df.columns if str(col) == '등급'), None)
        col_unit = next((col for col in df.columns if str(col) == '단위'), None)

        category_context = category_context or self._detect_standard_category_context(
            df,
            category_override=category_override,
        )
        row_categories = category_context.get('row_categories', {})
        default_category = category_context.get('primary_category') or self._fallback_standard_category(category_override)

        category_buckets = {}
        notes = []
        matched_rows = 0
        fallback_rows = 0
        excluded_rows = 0

        for row_idx, row in df.iterrows():
            item_text = str(row.get(col_item, '') if col_item else '').strip()
            if not item_text:
                continue

            item_kind = None
            if re.search(r'노무|인건|급여|인력', item_text):
                item_kind = 'labor'
            elif re.search(r'재료|자재|원재료|부품', item_text):
                item_kind = 'material'
            if not item_kind:
                continue

            category_value = row_categories.get(row_idx) or default_category
            detail_value = str(row.get(col_detail, '') if col_detail else '').strip()
            grade_value = str(row.get(col_grade, '') if col_grade else '').strip()
            unit_value = str(row.get(col_unit, '') if col_unit else '').strip()
            qty = self._safe_number(row.get('수량'), 0.0)
            actual_amount = self._safe_number(row.get('금액'), 0.0)

            price_group = self._filter_standard_category(self.df_price_db, category_value)
            std_price, converted_qty, match_note = self._pick_unit_price(
                price_group,
                item_kind,
                detail_value,
                grade_value,
                unit_value,
                qty,
            )
            if std_price is None:
                if item_kind == 'material' and actual_amount > 0:
                    ref_amount = actual_amount
                    fallback_rows += 1
                    if len(notes) < 8:
                        notes.append(f"{item_text}: {match_note}, 재료비는 제출 금액으로 반영")
                else:
                    ref_amount = 0.0
                    excluded_rows += 1
                    if len(notes) < 8:
                        notes.append(f"{item_text}: {match_note}, 공시기준가 산정에서 제외")
            else:
                ref_amount = std_price * converted_qty
                matched_rows += 1
                if len(notes) < 6:
                    notes.append(f"{item_text}: 표준단가 {std_price:,.0f} x 환산수량 {converted_qty:,.2f} ({match_note})")

            bucket_key = category_value or default_category
            if bucket_key not in category_buckets:
                category_buckets[bucket_key] = {'labor': 0.0, 'material': 0.0}
            category_buckets[bucket_key][item_kind] += ref_amount

        totals = {'labor': 0.0, 'material': 0.0, 'expense': 0.0, 'gna': 0.0, 'profit': 0.0, 'rnd': 0.0}
        for category_value, direct_amounts in category_buckets.items():
            subtotal = {
                'labor': direct_amounts.get('labor', 0.0),
                'material': direct_amounts.get('material', 0.0),
                'expense': 0.0,
                'gna': 0.0,
                'profit': 0.0,
                'rnd': 0.0,
            }
            for cost_item, code in [('경비', 'expense'), ('일반관리비', 'gna'), ('이윤', 'profit')]:
                ratio_row = self._pick_ratio_row(category_value, cost_item)
                if not ratio_row:
                    continue
                base_amount = self._resolve_ratio_base(cost_item, ratio_row['base_target'], subtotal)
                subtotal[code] = base_amount * (ratio_row['rate'] / 100.0)
            for code in totals:
                totals[code] += subtotal.get(code, 0.0)

        totals['total'] = sum(totals.values())
        coverage = f"표준단가 매칭 {matched_rows}건 / fallback {fallback_rows}건 / 제외 {excluded_rows}건"
        return {
            'amounts': totals,
            'coverage': coverage,
            'notes': notes,
            'category_context': category_context,
        }

    def _extract_strategic_insights(self, context):
        insights = []
        if not context: return ["특이사항 없음 (표준 단가 가이드 준용)"]
        
        if "단종" in context: insights.append("[공급망 리스크] 단종(EOL) 예정 품목 식별 - Last Time Buy 재고 확보 필요")
        if "비상주" in context: insights.append("[비용 최적화] 원격/비상주 인력에 대한 체재비 및 교통비 전액 삭감 요구")
        if "장기" in context: insights.append("[파트너십] 장기 거래 실적 기반 단가 동결 또는 Loyalty Discount(3~5%) 적용 요구")
        if "독점" in context: insights.append("[공급자 우위] 가격 인하 한계 - SLA(서비스수준협약) 강화 및 기술지원 기간 연장 주력")
        if "신규" in context: insights.append("[신규 진입] 레퍼런스 확보 조건부 공격적 진입 가격(Target -10%) 요구")
        if "환율" in context: insights.append("[변동성 관리] 환율 하락 시 단가 인하 연동 조항(De-escalation Clause) 계약 반영")
        
        return insights if insights else ["특이사항 없음 (표준 단가 가이드 준용)"]

    def detect_outliers(self, df_input: pd.DataFrame) -> List[str]:
        outliers = []
        try:
            prices = pd.to_numeric(df_input['단가'], errors='coerce').dropna()
            if len(prices) > 5:
                Q1, Q3 = prices.quantile(0.25), prices.quantile(0.75)
                threshold = Q3 + 1.5 * (Q3 - Q1)
                high_items = df_input[pd.to_numeric(df_input['단가'], errors='coerce') > threshold]
                for _, row in high_items.iterrows():
                    outliers.append(f"🚨 **이상 고가**: '{row['비목']}' ({float(row['단가']):,.0f}원) > 상한선 {threshold:,.0f}원")
        except: pass
        return outliers

    def _baseline_target_amount(self, adequacy_rows, item_code, total_amount):
        row = adequacy_rows.get(item_code, {})
        base_pct = float(row.get('base', 0.0) or 0.0)
        if total_amount <= 0 or base_pct <= 0:
            return 0.0
        return total_amount * (base_pct / 100.0)

    def _baseline_bound_amount(self, adequacy_rows, item_code, total_amount, bound_key):
        row = adequacy_rows.get(item_code, {})
        pct = float(row.get(bound_key, 0.0) or 0.0)
        if total_amount <= 0 or pct <= 0:
            return 0.0
        return total_amount * (pct / 100.0)

    def _apply_total_cap(self, amounts, total_cap):
        adjusted = {key: float(value or 0.0) for key, value in amounts.items() if key != 'total'}
        if total_cap <= 0:
            return adjusted

        current_total = sum(adjusted.values())
        excess = current_total - total_cap
        if excess <= 0:
            return adjusted

        # Trim negotiated overhead first before touching direct costs.
        reduction_order = ['profit', 'gna', 'expense', 'rnd', 'labor', 'material']
        for item_code in reduction_order:
            available = float(adjusted.get(item_code, 0.0) or 0.0)
            if available <= 0 or excess <= 0:
                continue
            reduction = min(available, excess)
            adjusted[item_code] = available - reduction
            excess -= reduction

        return adjusted

    def _cap_reference_amounts(self, quote_amounts, reference_amounts, adequacy_rows):
        vendor_total = float(quote_amounts.get('total', 0.0) or 0.0)
        capped = {}
        for item_code in ['labor', 'material', 'expense', 'gna', 'profit', 'rnd']:
            candidate = float(reference_amounts.get(item_code, 0.0) or 0.0)
            actual_amount = float(quote_amounts.get(item_code, 0.0) or 0.0)
            if candidate <= 0:
                capped[item_code] = 0.0
                continue

            limit = candidate
            if actual_amount > 0:
                limit = min(limit, actual_amount)
            if item_code in {'expense', 'gna', 'profit', 'rnd'}:
                high_amount = self._baseline_bound_amount(adequacy_rows, item_code, vendor_total, 'high')
                if high_amount > 0:
                    limit = min(limit, high_amount)
            capped[item_code] = max(limit, 0.0)

        return self._apply_total_cap(capped, vendor_total)

    def _build_financial_ratio_reference_amounts(self, quote_amounts, adequacy_rows):
        vendor_total = float(quote_amounts.get('total', 0.0) or 0.0)
        reference_amounts = {}
        for item_code in ['labor', 'material', 'expense', 'gna', 'profit', 'rnd']:
            reference_amounts[item_code] = self._baseline_target_amount(adequacy_rows, item_code, vendor_total)
        reference_amounts['total'] = sum(reference_amounts.values())
        return reference_amounts

    def _cap_internal_amount(self, item_code, candidate_amount, quote_amounts, adequacy_rows):
        candidate = float(candidate_amount or 0.0)
        if candidate <= 0:
            return 0.0

        vendor_total = float(quote_amounts.get('total', 0.0) or 0.0)
        actual_amount = float(quote_amounts.get(item_code, 0.0) or 0.0)
        capped = candidate

        if actual_amount > 0:
            capped = min(capped, actual_amount)
        if item_code in {'expense', 'gna', 'profit', 'rnd'}:
            high_amount = self._baseline_bound_amount(adequacy_rows, item_code, vendor_total, 'high')
            if high_amount > 0:
                capped = min(capped, high_amount)
        return max(capped, 0.0)

    def _derive_internal_target_amounts(self, quote_amounts, reference_amounts, adequacy_rows):
        vendor_total = float(quote_amounts.get('total', 0.0))
        labor_candidate = float(reference_amounts.get('labor', 0.0)) or float(quote_amounts.get('labor', 0.0))
        material_candidate = float(reference_amounts.get('material', 0.0)) or float(quote_amounts.get('material', 0.0))
        expense_candidate = float(reference_amounts.get('expense', 0.0)) or self._baseline_target_amount(adequacy_rows, 'expense', vendor_total)
        gna_candidate = float(reference_amounts.get('gna', 0.0)) or self._baseline_target_amount(adequacy_rows, 'gna', vendor_total)
        profit_candidate = float(reference_amounts.get('profit', 0.0)) or self._baseline_target_amount(adequacy_rows, 'profit', vendor_total)
        rnd_candidate = min(
            float(quote_amounts.get('rnd', 0.0)),
            self._baseline_target_amount(adequacy_rows, 'rnd', vendor_total) or float(quote_amounts.get('rnd', 0.0)),
        )
        internal = {
            'labor': self._cap_internal_amount('labor', labor_candidate, quote_amounts, adequacy_rows),
            'material': self._cap_internal_amount('material', material_candidate, quote_amounts, adequacy_rows),
            'expense': self._cap_internal_amount('expense', expense_candidate, quote_amounts, adequacy_rows),
            'gna': self._cap_internal_amount('gna', gna_candidate, quote_amounts, adequacy_rows),
            'profit': self._cap_internal_amount('profit', profit_candidate, quote_amounts, adequacy_rows),
            'rnd': self._cap_internal_amount('rnd', rnd_candidate, quote_amounts, adequacy_rows),
        }
        return self._apply_total_cap(internal, vendor_total)

    def _derive_settlement_amounts(self, quote_amounts, internal_amounts):
        settlement = {}
        for item_code, target_amount in internal_amounts.items():
            actual_amount = float(quote_amounts.get(item_code, 0.0))
            target_amount = float(target_amount or 0.0)
            if actual_amount > target_amount > 0:
                settlement[item_code] = (actual_amount + target_amount) / 2.0
            elif actual_amount > 0 and target_amount <= 0:
                settlement[item_code] = actual_amount
            elif actual_amount <= target_amount and actual_amount > 0:
                settlement[item_code] = actual_amount
            else:
                settlement[item_code] = target_amount
        return settlement

    def _build_target_policy_notes(self, category_context, internal_amounts, reference_amounts, reference_mode=""):
        primary_category = category_context.get('primary_category', '-')
        if reference_mode == "uploaded_observed":
            basis_note = "당사 목표는 협력사 업로드 재무제표에서 산출한 비목 Base 비율을 1차 기준점으로 사용했습니다."
        elif reference_mode == "estimated_partial":
            basis_note = "당사 목표는 일부 업로드 재무자료를 우선 반영하고, 부족한 항목만 업종 추정 기준으로 보완했습니다."
        elif reference_mode == "quick_industry_proxy":
            basis_note = "당사 목표는 빠른 분석용 업종 프록시 기준을 사용했습니다."
        else:
            basis_note = "당사 목표는 standard_db 직접비와 업로드 재무기준 비목 비율을 우선 사용하되, 협력사 제출액과 기준 상단 범위를 넘지 않도록 자동 상한을 적용했습니다."
        return [
            f"원가 체계는 {primary_category} 기준으로 자동 인식했습니다.",
            basis_note,
            f"표준 노무비 기준: {float(reference_amounts.get('labor', 0.0)):,.0f}원" if reference_amounts.get('labor') else "표준 노무비가 없어 협력사 제출 노무비를 기준으로 사용했습니다.",
            f"자동 산정 내부 목표 합계: {sum(float(v) for v in internal_amounts.values()):,.0f}원",
        ]

    def compute_should_cost(self, financial_ratios, quote_amounts, context_text=""):
        ratios = financial_ratios or {}
        v_total = float(quote_amounts.get("total", 0.0) or 0.0)
        if v_total <= 0:
            return {}

        result: Dict[str, Any] = {}
        observed_labor_ratio = _safe_float_value(ratios.get("observed_labor_ratio"))
        if observed_labor_ratio is not None and observed_labor_ratio > 0:
            labor_total_should = v_total * observed_labor_ratio
            burden_rate = _safe_float_value(ratios.get("labor_burden_rate"))
            if burden_rate is None or burden_rate < 0:
                burden_rate = 0.30
            burden_rate = max(0.0, min(burden_rate, 1.50))
            direct_labor_should = labor_total_should / (1.0 + burden_rate)
            indirect_labor_should = labor_total_should - direct_labor_should
            result["direct_labor_should"] = direct_labor_should
            result["indirect_labor_should"] = indirect_labor_should
            result["labor_burden_rate"] = burden_rate

        utilization = _safe_float_value(ratios.get("estimated_utilization"))
        sweet_spot = 0.85
        if utilization is not None:
            utilization = max(0.0, min(utilization, 1.0))
            if utilization < sweet_spot:
                fixed_cost_reduction = 1.0 - (utilization / sweet_spot)
                result["utilization_gap"] = sweet_spot - utilization
                result["fixed_cost_reduction_pct"] = fixed_cost_reduction
                result["unit_cost_saving_pct"] = fixed_cost_reduction * 0.60
            else:
                result["utilization_gap"] = 0.0
                result["fixed_cost_reduction_pct"] = 0.0
                result["unit_cost_saving_pct"] = 0.0

        cogs_ratio = _safe_float_value(ratios.get("cogs_ratio"))
        if cogs_ratio is not None and cogs_ratio > 0:
            cogs_ratio = max(0.0, min(cogs_ratio, 1.0))
            result["estimated_cogm"] = v_total * cogs_ratio
            result["estimated_gross_margin"] = v_total * (1.0 - cogs_ratio)
            result["cogs_ratio"] = cogs_ratio

        if ("비상주" in context_text or "원격" in context_text) and result.get("direct_labor_should"):
            travel_factor = 0.10
            result["travel_expense_add"] = result["direct_labor_should"] * travel_factor

        return result

    def build_should_cost_summary(self, should_cost_result, quote_amounts):
        sc = should_cost_result or {}
        if not sc:
            return ""

        lines = ["[Should-Cost 역산 분석 (재무제표 기반)]"]
        direct_labor_should = _safe_float_value(sc.get("direct_labor_should"))
        indirect_labor_should = _safe_float_value(sc.get("indirect_labor_should"))
        burden_rate = _safe_float_value(sc.get("labor_burden_rate"))
        if direct_labor_should is not None:
            lines.append(f"  • 재무제표 역산 적정 직접노무비: {direct_labor_should:,.0f}원")
        if indirect_labor_should is not None and burden_rate is not None:
            lines.append(f"  • 적정 간접노무비(부담률 {burden_rate * 100:.0f}%): {indirect_labor_should:,.0f}원")

        utilization_gap = _safe_float_value(sc.get("utilization_gap"))
        unit_cost_saving_pct = _safe_float_value(sc.get("unit_cost_saving_pct"))
        if utilization_gap is not None and utilization_gap > 0:
            current_utilization = 0.85 - utilization_gap
            lines.append(
                f"  • 추정 가동률: {current_utilization * 100:.0f}% (Sweet-Spot 85% 대비 {utilization_gap * 100:.1f}%p 부족)"
            )
        if unit_cost_saving_pct is not None and unit_cost_saving_pct > 0:
            lines.append(f"  • 가동률 개선 시 단위원가 절감 여력: {unit_cost_saving_pct * 100:.1f}%")

        estimated_cogm = _safe_float_value(sc.get("estimated_cogm"))
        cogs_ratio = _safe_float_value(sc.get("cogs_ratio"))
        if estimated_cogm is not None:
            cogs_ratio_text = f" ({cogs_ratio * 100:.1f}%)" if cogs_ratio is not None else ""
            lines.append(f"  • 추정 COGM: {estimated_cogm:,.0f}원{cogs_ratio_text}")

        travel_expense_add = _safe_float_value(sc.get("travel_expense_add"))
        if travel_expense_add is not None and travel_expense_add > 0:
            lines.append(f"  • 비상주/원격 보정 출장비 가산 추정: {travel_expense_add:,.0f}원")

        return "\n".join(lines)

    def _analyze_legacy(self, df_input, ui_category, user_targets, context_text="", analysis_meta=None):
        validation_errors = DataValidator.validate_estimate_data(df_input)
        
        df_input['금액'] = df_input['단가'] * df_input['수량']
        adequacy_report = evaluate_quote_adequacy(
            df_input,
            ui_category,
            meta=analysis_meta or {},
        )

        quote_amounts = adequacy_report.get('quote_amounts', {})
        standard_reference = self._build_standard_reference(df_input, ui_category)
        reference_amounts = standard_reference.get('amounts', {})
        reference_labor = float(reference_amounts.get('labor', 0.0))
        v_labor = float(quote_amounts.get('labor', 0.0))
        v_material = float(quote_amounts.get('material', 0.0))
        v_expense = float(quote_amounts.get('expense', 0.0))
        v_admin = float(quote_amounts.get('gna', 0.0))
        v_profit = float(quote_amounts.get('profit', 0.0))
        v_rnd = float(quote_amounts.get('rnd', 0.0))
        v_total = float(quote_amounts.get('total', df_input['금액'].sum()))

        # 2. 당사 목표가 산출
        labor_target_base = reference_labor if reference_labor > 0 else v_labor
        t_labor = labor_target_base * user_targets['labor_coeff']
        t_expense = t_labor * (user_targets['exp'] / 100)
        if "비상주" in context_text: t_expense *= 0.3
        t_admin = (t_labor + t_expense) * (user_targets['adm'] / 100)
        t_profit = (t_labor + t_expense + t_admin) * (user_targets['prof'] / 100)
        t_rnd = v_rnd
        t_total = t_labor + v_material + t_expense + t_admin + t_profit + t_rnd
        
        # 3. 현실적 타결가 산출 로직 (비목별 차등 타결법)
        r_labor = (v_labor + t_labor) / 2
        r_expense = t_expense 
        r_admin = t_admin     
        r_profit = t_profit   
        r_rnd = v_rnd
        r_total = r_labor + v_material + r_expense + r_admin + r_profit + r_rnd

        def format_amount_with_ratio(val, ratio=None):
            s_val = f"{val:,.0f}"
            if ratio is None:
                return s_val
            return f"{s_val} ({ratio:.1f}%)"

        def with_note(formatted_text, note_text):
            if not note_text:
                return formatted_text
            return f"{formatted_text} [{note_text}]"

        item_order = ['labor', 'material', 'expense', 'gna', 'profit']
        adequacy_rows = {row['item']: row for row in adequacy_report.get('item_rows', [])}
        if 'rnd' in adequacy_rows:
            item_order.append('rnd')

        item_labels = {
            'labor': '노무비',
            'material': '재료비',
            'expense': '경비',
            'gna': '일반관리비',
            'profit': '이윤',
            'rnd': 'R&D',
        }
        internal_amounts = {
            'labor': t_labor,
            'material': v_material,
            'expense': t_expense,
            'gna': t_admin,
            'profit': t_profit,
            'rnd': t_rnd,
        }
        settlement_amounts = {
            'labor': r_labor,
            'material': v_material,
            'expense': r_expense,
            'gna': r_admin,
            'profit': r_profit,
            'rnd': r_rnd,
        }

        summary_rows = []
        exception_notes = []
        adjusted_reference_total = 0.0
        for item_code in item_order:
            item_row = adequacy_rows.get(item_code, {})
            actual_amount = float(quote_amounts.get(item_code, 0.0))
            reference_amount = float(reference_amounts.get(item_code, 0.0))
            adjustment_note = ""
            row_note = ""
            if item_row:
                actual_ratio = float(item_row.get('actual_ratio', 0.0))
                base_ratio = float(item_row.get('base', 0.0))
                gap_pctp = actual_ratio - base_ratio
                candidate_average = (actual_amount + internal_amounts.get(item_code, 0.0)) / 2.0
                if abs(gap_pctp) >= 15.0 and candidate_average > 0:
                    reference_amount = candidate_average
                    adjustment_note = "예외보정"
                    row_note = f"Base gap {gap_pctp:+.1f}%p로 협력사 제출/당사 목표 평균 적용"
                    exception_notes.append(f"{item_labels.get(item_code, item_code)}: {row_note}")
            adjusted_reference_total += reference_amount
            reference_ratio = (reference_amount / v_total * 100.0) if v_total > 0 and reference_amount > 0 else (0.0 if reference_amount == 0 else None)
            target_range = "-"
            if item_row:
                target_range = f"{float(item_row.get('low', 0.0)):.1f}% ~ {float(item_row.get('high', 0.0)):.1f}%"

            summary_rows.append(
                {
                    '구분': item_labels.get(item_code, item_code),
                    '1. 협력사 제출': format_amount_with_ratio(actual_amount, item_row.get('actual_ratio') if item_row else None),
                    '2. 공시기준가': with_note(format_amount_with_ratio(reference_amount, reference_ratio), adjustment_note),
                    'Target Range': target_range,
                    '판정': ADEQUACY_JUDGEMENT_LABELS.get(str(item_row.get('judgement', '')).lower(), '-') if item_row else '-',
                    '기준 유형': BASIS_TYPE_LABELS.get(str(item_row.get('basis_type')), str(item_row.get('basis_type', '-'))) if item_row else '-',
                    '신뢰도': CONFIDENCE_LABELS.get(str(item_row.get('confidence_grade')), str(item_row.get('confidence_grade', '-'))) if item_row else '-',
                    '3. 당사 목표': format_amount_with_ratio(internal_amounts.get(item_code, 0.0), (internal_amounts.get(item_code, 0.0) / t_total * 100.0) if t_total > 0 else None),
                    '4. 현실적 타결': format_amount_with_ratio(settlement_amounts.get(item_code, 0.0), (settlement_amounts.get(item_code, 0.0) / r_total * 100.0) if r_total > 0 else None),
                    '주석': row_note,
                }
            )

        summary_rows.append(
            {
                '구분': '합계',
                '1. 협력사 제출': format_amount_with_ratio(v_total, 100.0 if v_total > 0 else None),
                '2. 공시기준가': format_amount_with_ratio(adjusted_reference_total, (adjusted_reference_total / v_total * 100.0) if v_total > 0 and adjusted_reference_total > 0 else None),
                'Target Range': '-',
                '판정': ADEQUACY_JUDGEMENT_LABELS.get(str(adequacy_report.get('overall_judgement', '')).lower(), str(adequacy_report.get('overall_label', '-'))),
                '기준 유형': '-',
                '신뢰도': CONFIDENCE_LABELS.get(str(adequacy_report.get('base_confidence_bucket')), str(adequacy_report.get('base_confidence_bucket', '-'))),
                '3. 당사 목표': format_amount_with_ratio(t_total, 100.0 if t_total > 0 else None),
                '4. 현실적 타결': format_amount_with_ratio(r_total, 100.0 if r_total > 0 else None),
                '주석': f"예외보정 {len(exception_notes)}건 적용" if exception_notes else "",
            }
        )

        summary = pd.DataFrame(summary_rows)
        
        summary.index = range(1, len(summary) + 1)

        strat_notes = self._extract_strategic_insights(context_text)
        outliers = self.detect_outliers(df_input)
        
        basic_draft = f"""1. 분석 대상: {ui_category}
2. 주요 전략 포인트:
{chr(10).join(['- ' + n for n in strat_notes])}

3. 종합 의견:
상기 전략적 배경을 바탕으로 원가 구조를 재분석한 결과, 협력사 제출가 대비 약 {((v_total-r_total)/v_total*100):.1f}%의 절감 여력이 존재하는 것으로 판단됨. 이에 목표 타결가를 설정하고 협상을 진행하고자 함."""

        draft_text = basic_draft  

        return {
            'summary': summary, 
            'v_total': v_total, 't_total': t_total, 'r_total': r_total, 
            'save_amt': v_total-r_total, 'bottom_line': f"{v_total-r_total:,.0f}원 절감 가능", 
            'draft_text': draft_text,
            'strat_notes': strat_notes,
            'outliers': outliers,
            'validation_errors': validation_errors,
            'adequacy_report': adequacy_report,
            'cost_reference_note': "공시기준가 열은 standard_db 기준으로 직접노무비(표준단가×수량), 직접재료비(존재 시 반영), 간접비(INDIRECT_RATIO 비율) 순으로 산출한 참조 금액입니다. 표준단가는 정확 매칭 → 유사명칭 매칭 → 카테고리/등급/단위 대표단가(중앙값) 순으로 탐색합니다. 시간 단위는 1일=8시간, 1월=176시간, 1년=2112시간 가정으로 환산합니다. 다만 산업 baseline Base와 협력사 제출 실제 비중의 차이가 15%p 이상이면 예외보정으로 협력사 제출/당사 목표 평균값을 사용합니다.",
            'cost_reference_coverage': standard_reference.get('coverage', ''),
            'cost_reference_notes': standard_reference.get('notes', []) + exception_notes,
            'target_labor_basis_note': "당사 목표 노무비는 표준DB 기준 노무비가 있으면 이를 기준으로, 없으면 협력사 제출 노무비를 기준으로 산정합니다.",
        }, None

# --------------------------------------------------------------------------------
# [4] 챗봇 및 DB 인터페이스
# --------------------------------------------------------------------------------
    def analyze(self, df_input, user_targets=None, context_text="", analysis_meta=None, category_override=None):
        validation_errors = DataValidator.validate_estimate_data(df_input)

        df_input = df_input.copy()
        df_input['금액'] = pd.to_numeric(df_input['단가'], errors='coerce').fillna(0) * pd.to_numeric(df_input['수량'], errors='coerce').fillna(0)

        category_context = self._detect_standard_category_context(
            df_input,
            category_override=category_override,
        )
        effective_category = category_context.get('primary_category', self._fallback_standard_category(category_override))
        primary_group_label = category_context.get('primary_group_label', self._category_group_label(effective_category))
        category_display = primary_group_label if primary_group_label == effective_category else f"{primary_group_label} ({effective_category})"
        strategy_category = category_context.get('strategy_category', self._strategy_category(effective_category))

        adequacy_report = evaluate_quote_adequacy(
            df_input,
            effective_category,
            meta=analysis_meta or {},
        )

        quote_amounts = adequacy_report.get('quote_amounts', {})
        financial_ratios = dict((analysis_meta or {}).get("financial_ratios") or {})
        reference_mode = str((analysis_meta or {}).get("reference_mode") or "")
        reference_source_label = str((analysis_meta or {}).get("reference_source_label") or "").strip()
        use_uploaded_financial_reference = reference_mode == "uploaded_observed"
        should_cost_result = self.compute_should_cost(financial_ratios, quote_amounts, context_text)
        should_cost_summary = self.build_should_cost_summary(should_cost_result, quote_amounts)
        standard_reference = self._build_standard_reference(
            df_input,
            category_context=category_context,
            category_override=category_override,
        )
        v_total = float(quote_amounts.get('total', df_input['금액'].sum()))
        adequacy_rows = {row['item']: row for row in adequacy_report.get('item_rows', [])}
        if use_uploaded_financial_reference:
            reference_amounts = self._cap_reference_amounts(
                quote_amounts,
                self._build_financial_ratio_reference_amounts(quote_amounts, adequacy_rows),
                adequacy_rows,
            )
        else:
            reference_amounts = self._cap_reference_amounts(
                quote_amounts,
                standard_reference.get('amounts', {}),
                adequacy_rows,
            )

        internal_amounts = self._derive_internal_target_amounts(
            quote_amounts,
            reference_amounts,
            adequacy_rows,
        )
        settlement_amounts = self._derive_settlement_amounts(quote_amounts, internal_amounts)
        t_total = sum(float(v) for v in internal_amounts.values())
        r_total = sum(float(v) for v in settlement_amounts.values())
        target_policy_notes = self._build_target_policy_notes(
            category_context,
            internal_amounts,
            reference_amounts,
            reference_mode=reference_mode,
        )
        if use_uploaded_financial_reference and reference_source_label:
            target_policy_notes.insert(1, f"기준 출처: {reference_source_label}")

        if use_uploaded_financial_reference:
            cost_reference_note = (
                "공시기준가 열은 협력사 업로드 재무제표에서 산출한 비목 Base 비율을 견적 총액에 적용해 만든 기준 금액입니다. "
                "노무비·재료비·간접비 모두 협력사 실측 재무기준을 1차 기준으로 사용하고, 협력사 제출액과 업로드 재무기준 상단 범위를 넘지 않도록 자동 상한을 적용합니다. "
                "Base와 협력사 제출 실제 비중의 차이가 15%p 이상이면 예외보정으로 협력사 제출/당사 목표 평균값을 사용합니다."
            )
            cost_reference_coverage = (
                f"{reference_source_label} 기준 Base 비율 × 견적 총액"
                if reference_source_label
                else "협력사 업로드 재무제표 기준 Base 비율 × 견적 총액"
            )
            cost_reference_notes = []
            if reference_source_label:
                cost_reference_notes.append(f"기준 출처: {reference_source_label}")
            cost_reference_notes.append("비목 기준점은 업종 평균이 아니라 협력사 재무제표에서 복원한 Base 비율입니다.")
        elif reference_mode == "estimated_partial":
            cost_reference_note = (
                "공시기준가 열은 일부 업로드 재무자료를 우선 반영하되, 부족한 항목은 업종 추정 기준과 standard_db 직접비를 보조적으로 사용해 산출한 참조 금액입니다. "
                "시간 단위는 1일=8시간, 1월=176시간, 1년=2112시간 가정으로 환산하며, 간접비·이윤은 협력사 제출액과 업로드 재무기준 상단 범위를 넘지 않도록 자동 상한을 적용합니다. "
                "Base와 협력사 제출 실제 비중의 차이가 15%p 이상이면 예외보정으로 협력사 제출/당사 목표 평균값을 사용합니다."
            )
            cost_reference_coverage = standard_reference.get('coverage', '')
            cost_reference_notes = standard_reference.get('notes', [])
            if reference_source_label:
                cost_reference_notes = [f"기준 출처: {reference_source_label}"] + cost_reference_notes
        elif reference_mode == "quick_industry_proxy":
            cost_reference_note = (
                "공시기준가 열은 사업자등록증 기반 업종 프록시 기준과 standard_db 직접비를 이용해 빠른 분석용으로 산출한 참조 금액입니다. "
                "정밀 판정은 협력사 재무제표 업로드 후 생성한 재무기준을 우선 사용해야 합니다."
            )
            cost_reference_coverage = standard_reference.get('coverage', '')
            cost_reference_notes = standard_reference.get('notes', [])
            if reference_source_label:
                cost_reference_notes = [f"기준 출처: {reference_source_label}"] + cost_reference_notes
        else:
            cost_reference_note = (
                "공시기준가 열은 standard_db 기준으로 직접노무비(표준단가×수량), 직접재료비(존재 시 반영), 간접비(INDIRECT_RATIO 중간비율) 순으로 산출한 참조 금액입니다. "
                "카테고리는 접두사 차이를 정규화해 매칭하고, 카테고리 매칭 실패 시 다른 업종 DB로 확장하지 않습니다. "
                "시간 단위는 1일=8시간, 1월=176시간, 1년=2112시간 가정으로 환산하며, 간접비·이윤은 협력사 제출액과 업로드 재무기준 상단 범위를 넘지 않도록 자동 상한을 적용합니다. "
                "다만 업로드 재무기준 Base와 협력사 제출 실제 비중의 차이가 15%p 이상이면 예외보정으로 협력사 제출/당사 목표 평균값을 사용합니다."
            )
            cost_reference_coverage = standard_reference.get('coverage', '')
            cost_reference_notes = standard_reference.get('notes', [])

        def format_amount_with_ratio(val, ratio=None):
            s_val = f"{val:,.0f}"
            if ratio is None:
                return s_val
            return f"{s_val} ({ratio:.1f}%)"

        def with_note(formatted_text, note_text):
            if not note_text:
                return formatted_text
            return f"{formatted_text} [{note_text}]"

        item_order = ['labor', 'material', 'expense', 'gna', 'profit']
        if 'rnd' in adequacy_rows:
            item_order.append('rnd')

        item_labels = {
            'labor': '노무비',
            'material': '재료비',
            'expense': '경비',
            'gna': '일반관리비',
            'profit': '이윤',
            'rnd': 'R&D',
        }
        summary_rows = []
        exception_notes = []
        adjusted_reference_total = 0.0
        for item_code in item_order:
            item_row = adequacy_rows.get(item_code, {})
            actual_amount = float(quote_amounts.get(item_code, 0.0))
            reference_amount = float(reference_amounts.get(item_code, 0.0))
            adjustment_note = ""
            row_note = ""
            if item_row:
                actual_ratio = float(item_row.get('actual_ratio', 0.0))
                base_ratio = float(item_row.get('base', 0.0))
                gap_pctp = actual_ratio - base_ratio
                candidate_average = (actual_amount + internal_amounts.get(item_code, 0.0)) / 2.0
                if abs(gap_pctp) >= 15.0 and candidate_average > 0:
                    reference_amount = candidate_average
                    adjustment_note = "예외보정"
                    row_note = f"Base gap {gap_pctp:+.1f}%p로 협력사 제출/당사 목표 평균 적용"
                    exception_notes.append(f"{item_labels.get(item_code, item_code)}: {row_note}")
            adjusted_reference_total += reference_amount
            reference_ratio = (reference_amount / v_total * 100.0) if v_total > 0 and reference_amount > 0 else (0.0 if reference_amount == 0 else None)
            target_range = "-"
            if item_row:
                target_range = f"{float(item_row.get('low', 0.0)):.1f}% ~ {float(item_row.get('high', 0.0)):.1f}%"

            summary_rows.append(
                {
                    '구분': item_labels.get(item_code, item_code),
                    '1. 협력사 제출': format_amount_with_ratio(actual_amount, item_row.get('actual_ratio') if item_row else None),
                    '2. 공시기준가': with_note(format_amount_with_ratio(reference_amount, reference_ratio), adjustment_note),
                    'Target Range': target_range,
                    '판정': ADEQUACY_JUDGEMENT_LABELS.get(str(item_row.get('judgement', '')).lower(), '-') if item_row else '-',
                    '기준 유형': BASIS_TYPE_LABELS.get(str(item_row.get('basis_type')), str(item_row.get('basis_type', '-'))) if item_row else '-',
                    '신뢰도': CONFIDENCE_LABELS.get(str(item_row.get('confidence_grade')), str(item_row.get('confidence_grade', '-'))) if item_row else '-',
                    '3. 당사 목표': format_amount_with_ratio(internal_amounts.get(item_code, 0.0), (internal_amounts.get(item_code, 0.0) / t_total * 100.0) if t_total > 0 else None),
                    '4. 현실적 타결': format_amount_with_ratio(settlement_amounts.get(item_code, 0.0), (settlement_amounts.get(item_code, 0.0) / r_total * 100.0) if r_total > 0 else None),
                    '주석': row_note,
                }
            )

        summary_rows.append(
            {
                '구분': '합계',
                '1. 협력사 제출': format_amount_with_ratio(v_total, 100.0 if v_total > 0 else None),
                '2. 공시기준가': format_amount_with_ratio(adjusted_reference_total, (adjusted_reference_total / v_total * 100.0) if v_total > 0 and adjusted_reference_total > 0 else None),
                'Target Range': '-',
                '판정': ADEQUACY_JUDGEMENT_LABELS.get(str(adequacy_report.get('overall_judgement', '')).lower(), str(adequacy_report.get('overall_label', '-'))),
                '기준 유형': '-',
                '신뢰도': CONFIDENCE_LABELS.get(str(adequacy_report.get('base_confidence_bucket')), str(adequacy_report.get('base_confidence_bucket', '-'))),
                '3. 당사 목표': format_amount_with_ratio(t_total, 100.0 if t_total > 0 else None),
                '4. 현실적 타결': format_amount_with_ratio(r_total, 100.0 if r_total > 0 else None),
                '주석': f"예외보정 {len(exception_notes)}건 적용" if exception_notes else "",
            }
        )

        summary = pd.DataFrame(summary_rows)
        summary.index = range(1, len(summary) + 1)

        strat_notes = self._extract_strategic_insights(context_text)
        outliers = self.detect_outliers(df_input)
        savings_pct = ((v_total - r_total) / v_total * 100.0) if v_total > 0 else 0.0

        basic_draft = f"""1. 분석 대상 원가 체계: {effective_category}
1-1. 원가 체계 요약: {category_context.get('summary', '-')}
2. 주요 전략 시사점:
{chr(10).join(['- ' + n for n in strat_notes])}

3. 종합 의견:
상기 전략적 배경을 바탕으로 원가 구조를 재분석한 결과, 협력사 제출가 대비 약 {savings_pct:.1f}%의 절감 여력이 존재하는 것으로 판단됩니다. 이에 목표 타결가를 설정하고 협상을 진행할 것을 권고합니다."""

        return {
            'summary': summary,
            'v_total': v_total,
            't_total': t_total,
            'r_total': r_total,
            'save_amt': v_total - r_total,
            'bottom_line': f"{v_total-r_total:,.0f}원 절감 가능",
            'draft_text': basic_draft,
            'strat_notes': strat_notes,
            'outliers': outliers,
            'validation_errors': validation_errors,
            'adequacy_report': adequacy_report,
            'cost_reference_note': cost_reference_note,
            'cost_reference_coverage': cost_reference_coverage,
            'cost_reference_notes': cost_reference_notes + exception_notes,
            'target_labor_basis_note': (
                "당사 목표는 협력사 업로드 재무제표 기준 Base 비율을 우선 사용해 자동 산정되며, 현실적 타결가는 협력사 제출가와 내부 목표 사이에서 비목별로 자동 조정됩니다."
                if use_uploaded_financial_reference
                else "당사 목표는 표준DB 직접비와 업로드 재무기준 비율을 이용해 자동 산정되며, 현실적 타결가는 협력사 제출가와 내부 목표 사이에서 비목별로 자동 조정됩니다."
            ),
            'target_policy_notes': target_policy_notes,
            'analysis_category': strategy_category,
            'analysis_category_display': category_display,
            'analysis_category_summary': category_context.get('summary', ''),
            'analysis_category_mix': category_context.get('mix_rows', []),
            'analysis_category_group_mix': category_context.get('group_mix_rows', []),
            'analysis_category_mode': category_context.get('mode', '자동 인식'),
            'should_cost': should_cost_result,
            'should_cost_summary': should_cost_summary,
            'financial_ratios': financial_ratios,
            'reference_mode': reference_mode,
        }, None

class SecureHardcodedChatbot:
    def __init__(self, analysis_result, scenario_db_path="buying_scenario_db.xlsx"):
        self.res = analysis_result
        current_dir = os.path.dirname(os.path.abspath(__file__))
        db_full_path = os.path.join(current_dir, scenario_db_path)
        self.strat_db = load_db_data(db_full_path, 0)
        self.nlp_engine = LocalHybridSearchEngine(self.strat_db)

    def explain_calculation_logic(self, category, cost_item):
        res = {
            "name": f"{category} > {cost_item}",
            "formula": "시장 최저가 적용 (Market Low Price)",
            "definition": "해당 비목에 대한 통상적인 시장 거래 가격입니다.",
            "reason": "명확한 법적/통계적 산출 기준이 부재하므로, 3개사 이상 비교 견적을 통한 최저가 적용을 원칙으로 합니다.",
            "description": "세부 내역서(Breakdown)를 요구하여 불필요한 마진이 포함되지 않았는지 점검하십시오."
        }

        if "SW" in category or "개발" in category:
            if "노무비" in cost_item or "인건" in cost_item:
                res = {
                    "name": "SW 엔지니어링 직접노무비",
                    "formula": "∑ (KOSA 등급별 평균임금 × 투입 M/M × 보정계수)",
                    "definition": "SW사업 대가산정 가이드에 따라 소프트웨어 기술자 등급별(초/중/고/특급)로 지급되는 노임 단가입니다.",
                    "reason": "소프트웨어 진흥법 제46조에 의거, 매년 한국소프트웨어산업협회(KOSA)가 공표하는 통계치를 기준으로 합니다. 단, 단순 유지보수 사업은 통상 80~90% 할인율(낙찰률 고려)을 적용하는 것이 시장 관례입니다.",
                    "description": "투입 인력의 등급이 실제 경력 증명서와 일치하는지 반드시 대조 확인이 필요합니다."
                }
            elif "경비" in cost_item or "제경비" in cost_item:
                res = {
                    "name": "SW 제경비 (Overhead Costs)",
                    "formula": "직접인건비 × 110% ~ 120%",
                    "definition": "직접적인 개발 업무 외에 발생하는 기업의 간접 비용(사무실 임대료, 공과금, 복리후생비, 감가상각비 등)입니다.",
                    "reason": "엔지니어링산업 진흥법 및 SW사업 대가산정 가이드상 제경비 상한선은 110~120%입니다. 그러나 '비상주(원격) 개발'의 경우 공급사의 임대료/관리비 부담이 감소하므로 100% 이하로 삭감 요구가 타당합니다.",
                    "description": "비상주 인력 비율이 높다면 제경비율을 하향 조정(110% → 100% 이하)하도록 협상하십시오."
                }
            elif "이윤" in cost_item or "기술료" in cost_item:
                res = {
                    "name": "SW 기술료 및 이윤",
                    "formula": "(직접인건비 + 제경비) × 20% 이내",
                    "definition": "공급사의 영업이익 및 기술 축적/재투자를 위한 비용 항목입니다.",
                    "reason": "관례상 원가의 20%를 상한으로 인정하나, R&D 요소가 적은 단순 코딩/유지보수 용역은 10% 내외로 조정하는 것이 합리적입니다.",
                    "description": "단순 인력 공급 성격의 사업이라면 이윤율 10% 적용을 목표로 하십시오."
                }

        elif "공사" in category or "시설" in category:
            if "노무비" in cost_item:
                res = {
                    "name": "공사 직접노무비",
                    "formula": "표준품셈(투입 공수) × 시중노임단가(대한건설협회)",
                    "definition": "공사 목적물을 완성하기 위해 현장에서 직접 근로하는 인력에 대한 대가입니다.",
                    "reason": "통계법 제27조에 따라 대한건설협회가 조사/공표하는 '시중노임단가'를 적용해야 합니다. 자의적인 단가 적용은 불가하며, 품셈(투입 인원수)의 적정성이 주요 검토 대상입니다.",
                    "description": "야간/휴일 할증(50%)이 불필요하게 포함되었는지 작업 계획표와 대조하십시오."
                }
            elif "재료비" in cost_item or "자재" in cost_item:
                res = {
                    "name": "직접재료비",
                    "formula": "소요량 × (조달청 단가 or 물가정보지 최저가)",
                    "definition": "공사에 직접적으로 소모되는 자재, 부품, 재료의 비용입니다.",
                    "reason": "국가계약법 시행규칙 제74조에 의거, 거래실례가격(조달청 나라장터, 물가자료, 물가정보, 유통물가) 중 '최저가'를 적용하는 것이 원칙입니다.",
                    "description": "자재 손실률(Loss)이 표준(3~5%)을 초과하여 계상되었는지 확인하십시오."
                }
            elif "경비" in cost_item:
                res = {
                    "name": "공사 법정경비",
                    "formula": "재료비/노무비 × 법정 요율 (산재, 고용, 안전관리비 등)",
                    "definition": "법령에 의해 의무적으로 가입/납부해야 하는 보험료 및 안전관리 비용입니다.",
                    "reason": "산업안전보건법 및 보험료징수법에 따라 정해진 요율을 반드시 적용해야 합니다. 이는 법적 의무 사항이므로 임의 삭감이 불가합니다.",
                    "description": "법정 요율 미만으로 계상 시 법 위반 소지가 있으므로 삭감 대상에서 제외하십시오."
                }

        else:
            if "관리" in cost_item: 
                res = {
                    "name": "일반관리비 (General Admin)",
                    "formula": "총 원가 × 5% ~ 6%",
                    "definition": "기업의 일반적인 관리 및 유지 활동에 소요되는 제비용(본사 운영비, 임원 급여 등)입니다.",
                    "reason": "국가를 당사자로 하는 계약에 관한 법률 시행규칙 제8조에 따라, 일반 용역의 경우 일반관리비율은 6%(경비용역 등은 5%)를 초과할 수 없습니다.",
                    "description": "법적 상한선(6%)을 초과하여 청구된 경우 즉시 삭감 요구하십시오."
                }
            elif "이윤" in cost_item:
                res = {
                    "name": "적정 이윤",
                    "formula": "(총 원가 + 일반관리비) × 10% 이내",
                    "definition": "영업이익으로서 기업의 존속을 위한 이익금입니다.",
                    "reason": "기획재정부 계약예규 '예정가격 작성기준'에 따라 이윤은 원가+관리비의 10%를 초과할 수 없습니다.",
                    "description": "비영리법인과의 계약인 경우 이윤 항목은 인정되지 않습니다 (0원 처리)."
                }

        return res

    def get_advice(self, keyword, ui_category_selection=None):
        filter_key = ""
        if ui_category_selection:
            if "SW" in ui_category_selection: filter_key = "IT"
            elif "공사" in ui_category_selection: filter_key = "시설"
            elif "일반" in ui_category_selection: filter_key = ""
        
        results = self.nlp_engine.search_strategy(keyword, category_filter=filter_key)
        if results: return results[0]
        return None

    def get_keyword_options(self, ui_category_selection=None, max_items=300):
        if self.strat_db is None or self.strat_db.empty:
            return []

        filter_key = ""
        if ui_category_selection:
            if "SW" in ui_category_selection:
                filter_key = "IT"
            elif "공사" in ui_category_selection:
                filter_key = "시설"
            elif "일반" in ui_category_selection:
                filter_key = ""

        c_key = next((c for c in self.strat_db.columns if "Category" in c), "Category")
        t_key = next((c for c in self.strat_db.columns if "Trigger" in c), "Trigger_Keywords")

        keywords = []
        for _, row in self.strat_db.iterrows():
            category = str(row.get(c_key, ""))
            if filter_key and filter_key.lower() not in category.lower():
                continue
            raw = str(row.get(t_key, ""))
            if not raw:
                continue
            for token in re.split(r'[,|\n]', raw):
                token = token.strip()
                if len(token) >= 2:
                    keywords.append(token)

        deduped = sorted(set(keywords))
        return deduped[:max_items]

    def get_keyword_suggestions(self, partial_keyword, ui_category_selection=None, max_items=20):
        query = str(partial_keyword or "").strip().lower()
        if not query:
            return []

        options = self.get_keyword_options(ui_category_selection=ui_category_selection, max_items=1000)
        if not options:
            return []

        exact = [k for k in options if k.lower() == query]
        contains = [k for k in options if query in k.lower() and k.lower() != query]
        return (exact + contains)[:max_items]

# --------------------------------------------------------------------------------
# [5] 시각화 및 리포트
# --------------------------------------------------------------------------------
class WaterfallChartBuilder:
    @staticmethod
    def draw(v_total, r_total, summary_df):
        savings = v_total - r_total
        val_v = v_total / 10000; val_s = savings / 10000;
        val_r = r_total / 10000
        
        fig = go.Figure(go.Waterfall(
            name = "Cost Breakdown", orientation = "v",
            measure = ["relative", "relative", "total"],
            x = ["협력사 제출", "협상 절감", "목표 타결"],
            text = [f"{val_v:,.0f}만", f"-{val_s:,.0f}만", f"{val_r:,.0f}만"],
            y = [v_total, -savings, r_total],
            connector = {"line":{"color":"#ADB5BD", "dash":"dot"}},
            decreasing = {"marker":{"color":"#198754"}}, 
            increasing = {"marker":{"color":"#DC3545"}}, 
            totals = {"marker":{"color":"#0D6EFD"}}
        ))
        fig.update_layout(
            title=dict(text="비용 절감 구조 분석", font=dict(family="Malgun Gothic", size=16, color="#333333")),
            height=350, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(family="Malgun Gothic")
        )
        return fig

def wrap_text_pixel(text, font, max_width):
    lines = []
    if not text: return lines
    current_line = ""
    for char in text:
        test_line = current_line + char
        try: width = font.getlength(test_line)
        except: width = len(test_line) * 15 
        if width <= max_width: current_line = test_line
        else: 
            lines.append(current_line)
            current_line = char
    if current_line: lines.append(current_line)
    return lines

@st.cache_resource 
def get_cached_font(size):
    candidates = ["malgunbd.ttf", "malgun.ttf", "AppleGothic", "gulim.ttc", "arial.ttf"]
    for f in candidates:
        try: return ImageFont.truetype(f, size)
        except: continue
    return ImageFont.load_default()


def get_cached_emoji_font(size):
    for font_name in ["seguiemj.ttf", "Segoe UI Emoji.ttf"]:
        try:
            return ImageFont.truetype(font_name, size)
        except Exception:
            continue
    return None


def _normalize_strategy_report_text(text: str, emoji_font_available: bool) -> str:
    return (
        str(text or "")
        .replace("🚨", "★ 1순위")
        .replace("✅", "● 우수")
        .replace("⚠️", "▲ 주의")
        .replace("⚠", "▲ 주의")
    )

def create_strategy_report_image(category, v_total, r_total, summary_df, adopted_strategies, ai_strategy=""):
    # [NEW 핵심 패치] AI 텍스트 및 마크다운 표 출력을 위해 폭(W)과 높이(H)를 대폭 확장 (잘림 방지)
    W, H = 1200, 15000 
    MARGIN_X = 50    
    img = Image.new('RGB', (W, H), (255, 255, 255))
    d = ImageDraw.Draw(img)

    font_title = get_cached_font(32)
    font_subtitle = get_cached_font(16) 
    font_header = get_cached_font(18) 
    font_body = get_cached_font(14)
    font_emoji = get_cached_emoji_font(14)
    emoji_font_available = font_emoji is not None
    
    d.rectangle([(0, 0), (W, 90)], fill=(30, 58, 138)) 
    d.text((MARGIN_X, 28), "AI Buying Strategy Report", font=font_title, fill='white')
    d.text((W-200, 38), datetime.now().strftime('%Y-%m-%d'), font=font_subtitle, fill='#E2E8F0')

    current_y = 120

    d.rectangle([(MARGIN_X, current_y), (W-MARGIN_X, current_y+30)], fill='#E9ECEF')
    d.text((MARGIN_X+10, current_y+5), "01. Executive Summary (비용 분석 요약)", font=font_header, fill='#333333')
    current_y += 50

    d.text((MARGIN_X, current_y), f"Target Category: {category}", font=font_subtitle, fill='#333333')
    current_y += 30

    savings = v_total - r_total
 
    box_w = (W - (MARGIN_X*2) - 20) / 2
    
    d.rectangle([(MARGIN_X, current_y), (MARGIN_X+box_w, current_y+80)], outline='#1E3A8A', width=2)
    d.text((MARGIN_X+20, current_y+15), "Estimated Savings (절감액)", font=font_subtitle, fill='#666666')
    d.text((MARGIN_X+20, current_y+40), f"{savings:,.0f} KRW", font=get_cached_font(24), fill='#1E3A8A')
    
    d.rectangle([(MARGIN_X+box_w+20, current_y), (W-MARGIN_X, current_y+80)], outline='#CCCCCC', width=1)
    d.text((MARGIN_X+box_w+40, current_y+15), "Target Price (목표가)", font=font_subtitle, fill='#666666')
    d.text((MARGIN_X+box_w+40, current_y+40), f"{r_total:,.0f} KRW", font=get_cached_font(24), fill='#333333')

    current_y += 110

    d.rectangle([(MARGIN_X, current_y), (W-MARGIN_X, current_y+30)], fill='#E9ECEF')
    d.text((MARGIN_X+10, current_y+5), "02. Negotiation Strategy Roadmap (협상 전략)", font=font_header, fill='#333333')
    current_y += 50
    
    if adopted_strategies:
        for idx, item in enumerate(adopted_strategies):
            script = item.get('script', '')
            wrapped = wrap_text_pixel(script, font_body, W - (MARGIN_X*2) - 40)
            box_h = 35 + (len(wrapped)*20) + 20
            
            d.rectangle([(MARGIN_X, current_y), (W-MARGIN_X, current_y+box_h)], outline='#CCCCCC', width=1)
            d.rectangle([(MARGIN_X, current_y), (W-MARGIN_X, current_y+30)], fill='#F8F9FA')
            
            d.text((MARGIN_X+15, current_y+7), f"Strategy #{idx+1} | {item.get('category','-')}", font=font_subtitle, fill='#1E3A8A')
            
            text_y = current_y + 40
          
            for line in wrapped:
                d.text((MARGIN_X+20, text_y), line, font=font_body, fill='#333333')
                text_y += 20
            current_y += box_h + 20
            
            if current_y > H - 500: break
           
    current_y += 20
    if ai_strategy and not ai_strategy.startswith("⚠️"):
        d.rectangle([(MARGIN_X, current_y), (W-MARGIN_X, current_y+30)], fill='#E9ECEF')
        d.text((MARGIN_X+10, current_y+5), "03. EXAONE 종합 협상 전략 리포트", font=font_header, fill='#333333')
        current_y += 50
        
        # [NEW 핵심 패치] 마크다운 테이블 완벽 렌더링 및 이모지 그래픽 깨짐 방지 로직 (Additive)
        clean_ai_strategy = ai_strategy.replace('**', '').replace('### ', '').replace('## ', '').replace('# ', '')
        
        for paragraph in clean_ai_strategy.split('\n'):
            if current_y > H - 150:
                d.text((MARGIN_X, current_y), "... (이하 생략) ...", font=font_body, fill='#999999')
                current_y += 30
                break
                
            if paragraph.strip() == "":
                current_y += 15
                continue
                
            # 마크다운 표(|)를 엑셀처럼 깔끔한 그리드(Grid) 형태로 이미지에 그립니다.
            if "|" in paragraph:
                if "---" in paragraph.replace(' ', ''):
                    d.line([(MARGIN_X, current_y), (W-MARGIN_X, current_y)], fill='#1E3A8A', width=2)
                    current_y += 10
                    continue
                    
                cells = [c.strip() for c in paragraph.split('|')]
                if paragraph.strip().startswith('|'): cells = cells[1:]
                if paragraph.strip().endswith('|'): cells = cells[:-1]
                
                if not cells: continue
                
                cell_w = (W - (MARGIN_X*2)) / max(1, len(cells))
                for i, cell in enumerate(cells):
                    safe_cell = _normalize_strategy_report_text(cell, emoji_font_available)
                    d.text((MARGIN_X + (i * cell_w) + 10, current_y), safe_cell, font=font_body, fill='#1E3A8A')
                current_y += 25
            else:
                safe_paragraph = _normalize_strategy_report_text(paragraph, emoji_font_available)
                wrapped = wrap_text_pixel(safe_paragraph, font_body, W - (MARGIN_X*2))
                for w_line in wrapped:
                    if current_y > H - 150: break
                    d.text((MARGIN_X, current_y), w_line, font=font_body, fill='#333333')
                    current_y += 22

    current_y += 50
    d.line([(MARGIN_X, current_y), (W-MARGIN_X, current_y)], fill='#CCCCCC')
    d.text((MARGIN_X, current_y + 10), "Confidential Document | Generated by Pro_Buying_Agent & EXAONE", font=font_body, fill='#999999')
    
    # 텍스트가 다 그려진 후 실제 사용된 높이(current_y)까지만 이미지를 깔끔하게 자름
    final_h = min(current_y + 50, H)
    img = img.crop((0, 0, W, int(final_h)))

    return img

def create_excel_report(summary_df, category, bottom_line, draft_text, strategies, ai_strategy=""):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#E9ECEF', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_cell = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    
    ws1 = workbook.add_worksheet('종합 비용 분석')
    ws1.write(0, 0, f"Category: {category}")
    ws1.write(1, 0, f"Result: {bottom_line}")
    for i, col in enumerate(summary_df.columns): ws1.write(3, i, col, fmt_header)
    
    for r, row in summary_df.iterrows():
        for c, val in enumerate(row): ws1.write(4+r, c, str(val), fmt_cell)

    ws2 = workbook.add_worksheet('협상 전략 로드맵')
    ws2.set_column('A:A', 50);
    ws2.set_column('B:B', 80)
    ws2.write(0, 0, 'Category (Issue)', fmt_header)
    ws2.write(0, 1, 'Action Script', fmt_header)
    for i, item in enumerate(strategies):
        ws2.write(i+1, 0, item.get('category', '-'), fmt_cell)
        ws2.write(i+1, 1, item.get('script', '-'), fmt_cell)

    ws3 = workbook.add_worksheet('품의서')
    ws3.set_column('A:A', 2);
    ws3.set_column('B:B', 12)
    ws3.set_column('C:E', 10); ws3.set_column('F:F', 12); ws3.set_column('G:I', 10)
    
    f_title = workbook.add_format({'bold': True, 'font_size': 22, 'align': 'center', 'valign': 'vcenter', 'underline': 1})
    f_label = workbook.add_format({'bg_color': '#F8F9FA', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 10})
    f_val = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter', 'font_size': 10})
    f_val_center = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 10})
    f_box = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9})
    f_content = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 11, 'text_wrap': True})
  
    ws3.merge_range('B2:E4', '품  의  서', f_title)
    ws3.merge_range('F2:F4', '결\n\n재', f_label)
    ws3.write('G2', '담 당', f_box);
    ws3.write('H2', '팀 장', f_box); ws3.write('I2', '부문장', f_box)
    ws3.merge_range('G3:G4', '', f_box); ws3.merge_range('H3:H4', '', f_box);
    ws3.merge_range('I3:I4', '', f_box)

    ws3.write('B6', '문서번호', f_label); ws3.merge_range('C6:E6', '', f_val) 
    ws3.write('F6', '보존연한', f_label);
    ws3.merge_range('G6:I6', '5년', f_val_center)

    ws3.write('B7', '부  서', f_label); ws3.merge_range('C7:E7', '구매팀', f_val_center)
    ws3.write('F7', '기안일자', f_label);
    ws3.merge_range('G7:I7', datetime.now().strftime('%Y-%m-%d'), f_val_center)

    ws3.write('B8', '직  위', f_label);
    ws3.merge_range('C8:E8', '담당', f_val_center)
    ws3.write('F8', '성  명', f_label);
    ws3.merge_range('G8:I8', 'AI Agent', f_val_center)

    ws3.set_row(8, 5)
    ws3.write('B10', '제  목', f_label);
    ws3.merge_range('C10:I10', f" [{category}] 구매 협상 전략 및 원가 분석 보고의 건", f_val)
    ws3.write('B11', '내  용', f_label)
    ws3.merge_range('B12:I40', draft_text, f_content)
    ws3.merge_range('B42:I42', '위와 같이 품의하오니 재가 바랍니다.', workbook.add_format({'align': 'center', 'bold': True, 'font_size': 12}))

    if ai_strategy and not ai_strategy.startswith("⚠️"):
        ws4 = workbook.add_worksheet('AI 전략 리포트')
        ws4.set_column('A:A', 2)
        ws4.set_column('B:I', 12)
        
        f_ai_title = workbook.add_format({'bold': True, 'font_size': 16, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        f_ai_content = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 11, 'text_wrap': True})
        
        ws4.merge_range('B2:I3', '🎯 EXAONE 종합 협상 전략 리포트', f_ai_title)
        
        clean_ai_text = ai_strategy.replace('**', '').replace('### ', '')
        ws4.merge_range('B5:I100', clean_ai_text, f_ai_content)

    workbook.close()
    return output.getvalue()


def parse_amount_cell(value):
    text = str(value or "")
    text = text.split("[")[0]
    text = text.split("(")[0]
    text = text.replace(",", "").replace("원", "").strip()
    try:
        return float(text) if text else 0.0
    except Exception:
        return 0.0


def build_auto_negotiation_points(res, max_items=3):
    summary_df = res.get('summary')
    if summary_df is None or summary_df.empty:
        return []

    candidates = []
    for _, row in summary_df.iterrows():
        item_name = str(row.get('구분', '')).strip()
        if item_name in {'', '합계'}:
            continue
        submit_amount = parse_amount_cell(row.get('1. 협력사 제출', 0))
        public_amount = parse_amount_cell(row.get('2. 공시기준가', 0))
        target_amount = parse_amount_cell(row.get('4. 현실적 타결', 0))
        gap_amount = max(submit_amount - target_amount, 0.0)
        judgement = str(row.get('판정', '-')).strip()
        basis_type = str(row.get('기준 유형', '-')).strip()
        confidence = str(row.get('신뢰도', '-')).strip()
        target_range = str(row.get('Target Range', '-')).strip()
        note = str(row.get('주석', '')).strip()

        severity_score = {'심각': 4, '경고': 3, '주의': 2, '적정': 1}.get(judgement, 0)
        priority_score = gap_amount + (severity_score * 1000000)
        if gap_amount <= 0 and severity_score <= 1:
            continue

        if gap_amount > 0:
            action = f"협력사 제출 {submit_amount:,.0f}원에서 현실적 타결 {target_amount:,.0f}원까지 조정 여지가 있습니다."
        else:
            action = f"금액 차이는 크지 않지만 비목 구조상 판정은 {judgement}입니다."

        reason_parts = [f"목표 범위 {target_range}", f"기준 {basis_type}", f"신뢰도 {confidence}"]
        if public_amount > 0:
            reason_parts.append(f"공시기준가 {public_amount:,.0f}원")
        if note:
            reason_parts.append(note)

        candidates.append(
            {
                "item": item_name,
                "title": f"{item_name} 조정 검토",
                "body": f"{action} {' / '.join(reason_parts)}",
                "priority": priority_score,
            }
        )

    candidates.sort(key=lambda row: row["priority"], reverse=True)
    return candidates[:max_items]

def generate_standard_estimate_template():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = workbook.add_worksheet('견적서')
    
    fmt_title = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter', 'underline': True})
    fmt_label = workbook.add_format({'font_size': 11, 'align': 'left', 'valign': 'vcenter'})
    fmt_label_r = workbook.add_format({'font_size': 11, 'align': 'right', 'valign': 'vcenter'})
    fmt_header = workbook.add_format({'bold': True, 'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E2E8F0', 'border': 1})
    fmt_center = workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_text = workbook.add_format({'font_size': 10, 'align': 'left', 'valign': 'vcenter', 'border': 1})
    fmt_num = workbook.add_format({'font_size': 10, 'align': 'right', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'})
    fmt_total = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'right', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0', 'bg_color': '#FFF7ED'})
    fmt_total_label = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFF7ED'})

    ws.set_column('A:A', 5); ws.set_column('B:B', 13); ws.set_column('C:C', 10);
    ws.set_column('D:D', 35)
    ws.set_column('E:E', 10); ws.set_column('F:F', 8); ws.set_column('G:G', 8); ws.set_column('H:H', 12)
    ws.set_column('I:I', 15);
    ws.set_column('J:J', 15)

    ws.merge_range('A1:J2', '견  적  서', fmt_title)
    ws.write('A3', f" 일  자 : {datetime.now().strftime('%Y 년  %m 월  %d 일')}", fmt_label)
    ws.write('A4', " 수  신 : 귀하", fmt_label)
    ws.write('A5', " 참  조 : ", fmt_label)

    ws.write('G3', "등록번호", fmt_label_r);
    ws.write('H3', "", fmt_text)
    ws.write('G4', "상  호", fmt_label_r);
    ws.write('H4', "", fmt_text)
    ws.write('G5', "대  표", fmt_label_r);
    ws.write('H5', "", fmt_text)

    ws.merge_range('A7:F7', "합 계 금 액 :   일금                        원정 (VAT 별도)", fmt_label)
    ws.write_formula('G7', '=I60', fmt_total)

    headers = ['No', '카테고리', '비목', '품명(상세내역)', '등급', '단위', '수량', '단가', '금액(원)', '비고']
    for col, h in enumerate(headers): ws.write(8, col, h, fmt_header)

    for i in range(50):
        row = 9 + i
       
        ws.write(row, 0, i + 1, fmt_center)
        ws.write_blank(row, 1, None, fmt_center);
        ws.write_blank(row, 2, None, fmt_center)
        ws.write_blank(row, 3, None, fmt_text);
        ws.write_blank(row, 4, None, fmt_center)
        ws.write_blank(row, 5, None, fmt_center)
        ws.write(row, 6, 0, fmt_center);
        ws.write(row, 7, 0, fmt_num)
        formula = f'=IF(AND(ISNUMBER(G{row+1}), ISNUMBER(H{row+1})), G{row+1}*H{row+1}, 0)'
        ws.write_formula(row, 8, formula, fmt_num)
        ws.write_blank(row, 9, None, fmt_text)

    last_row = 59
    ws.merge_range(last_row, 0, last_row, 7, "총     계", fmt_total_label)
    ws.write_formula(last_row, 8, f'=SUM(I10:I59)', fmt_total)
    ws.write_blank(last_row, 9, None, fmt_total)

    ref_sheet = workbook.add_worksheet('Reference');
    ref_sheet.hide()
    categories = ['1.일반용역', '2.SW용역', '3-1.종합건축', '3-2.전문/인테리어', '3-3.전기/통신', '4.마케팅', '5.연구용역']
    cost_items = ['노무비', '재료비', '경비', '제경비', '일반관리비', '이윤', '산업안전보건관리비', '기술료']
    grades = ['특급', '고급', '중급', '초급', '반장', '보통', '해당없음']
    units = ['시간(Hr)', '일(day)', '월(Month)', '연(Year)', 'EA', 'Lot', 'Pcs', 'kg', 'm', 'set', 'Box', 'Roll', 'Can', 'Drum', 'm2', 'm3']

    ref_sheet.write_column('A1', categories);
    ref_sheet.write_column('B1', cost_items)
    ref_sheet.write_column('C1', grades); ref_sheet.write_column('D1', units)

    ws.data_validation('B10:B59', {'validate': 'list', 'source': '=Reference!$A$1:$A$7'})
    ws.data_validation('C10:C59', {'validate': 'list', 'source': '=Reference!$B$1:$B$8'})
    ws.data_validation('E10:E59', {'validate': 'list', 'source': '=Reference!$C$1:$C$7'})
    ws.data_validation('F10:F59', {'validate': 'list', 'source': '=Reference!$D$1:$D$16'})

    start_ref_row = 61
    ref_data_visual = [
        ['1.일반용역', '노무비', '특급', '시간(Hr)', 'EA'],
        ['2.SW용역', '재료비', '고급', '일(day)', 'Lot'],
        ['3-1.종합건축', '경비', '중급', '월(Month)', 'Pcs'],
        ['3-2.전문/인테리어', '제경비', '초급', '연(Year)', 'kg'],
        ['3-3.전기/통신', '일반관리비', '반장', '', 'm'],
        ['4.마케팅', '이윤', '보통', '', 'set'],
        ['5.연구용역', '산업안전보건관리비', '해당없음', '', 'Box'],
        ['', '기술료', '', '', 'Roll'],
        ['', '', '', '', 'Can'],
        ['', '', '', '', 'Drum'],
        ['', '', '', '', 'm2'],
        ['', '', '', '', 'm3']
    ]
    fmt_gray = workbook.add_format({'font_color': '#808080', 'font_size': 9})
    for i, row_data in enumerate(ref_data_visual):
        ws.write_row(start_ref_row + i, 1, row_data, fmt_gray)

    workbook.close()
    return output.getvalue()


# --------------------------------------------------------------------------------
# [6] 메인 실행부
# --------------------------------------------------------------------------------
@st.cache_resource
def get_analysis_engine():
    return CategoryAnalysisEngine(DB_FILENAME)

def main():
    apply_custom_css()

    if 'analysis_done' not in st.session_state: st.session_state['analysis_done'] = False
    if 'uploader_id' not in st.session_state: st.session_state['uploader_id'] = 0
    if 'ai_generated_strategy' not in st.session_state: st.session_state['ai_generated_strategy'] = None
    if 'exaone_response_cache' not in st.session_state: st.session_state['exaone_response_cache'] = {}
    if 'financial_upload_signature' not in st.session_state: st.session_state['financial_upload_signature'] = ""
    if 'financial_upload_cache' not in st.session_state: st.session_state['financial_upload_cache'] = []
    if 'financial_classification_signature' not in st.session_state: st.session_state['financial_classification_signature'] = ""
    if 'financial_classification_cache' not in st.session_state: st.session_state['financial_classification_cache'] = None
    if 'estimate_parse_signature' not in st.session_state: st.session_state['estimate_parse_signature'] = ""
    if 'estimate_parse_cache' not in st.session_state: st.session_state['estimate_parse_cache'] = None
    if 'estimate_parse_meta_cache' not in st.session_state: st.session_state['estimate_parse_meta_cache'] = None

    st.markdown("""
        <div class="header-box">
            <div class="header-title">구매 협상 AI 에이전트</div>
            <div class="header-subtitle">Corporate Edition (EXAONE AI 연동 완료)</div>
        </div>
    """, unsafe_allow_html=True)

    st.caption("🚀 AI 구동 환경: **저사양 PC 스로틀링 제한 및 실시간 스트리밍 출력 (방화벽 끊김 완전 차단)**")

    engine = get_analysis_engine()
    is_locked = False
    has_result = st.session_state.get('res') is not None
    supplier_reference_ctx = st.session_state.get('supplier_reference_ctx')
    selected_baseline_period = ""
    business_certificate_info = st.session_state.get("business_certificate_info")
    supplier_name = st.session_state.get("supplier_name_input", "")
    supplier_business_id = st.session_state.get("supplier_business_id_input", "")
    supplier_location = st.session_state.get("supplier_location_input", "")
    supplier_listing_status = st.session_state.get("supplier_listing_status_input", "비상장")
    supplier_doc_summary = st.session_state.get("supplier_doc_summary_input", "")

    with st.sidebar:
        st.header("설정 (Settings)")
        user_defined_model_path = DEFAULT_MODEL_PATH
        if os.path.exists(DEFAULT_MODEL_PATH):
            st.caption(f"AI 모델 경로 자동 인식 완료: {DEFAULT_MODEL_PATH}")
        else:
            st.warning("AI 모델 기본 경로를 찾지 못했습니다. 고급 설정에서 경로를 지정해 주세요.")

        manual_cost_category = ""
        with st.expander("고급 설정", expanded=False):
            st.markdown("##### 🤖 AI 모델 경로 설정")
            user_defined_model_path = st.text_input(
                "모델 폴더 절대 경로",
                value=DEFAULT_MODEL_PATH,
                help="EXAONE 모델 경로는 AI 협상 전략 생성 기능에서만 사용됩니다.",
            )
            if user_defined_model_path and not os.path.exists(user_defined_model_path):
                st.error("⚠️ 지정한 모델 경로를 찾을 수 없습니다.")
            elif user_defined_model_path:
                st.success("✅ 모델 경로 인식 완료")

            st.markdown("##### 원가 체계 설정")
            use_manual_cost_category = st.checkbox(
                "표준원가 체계 수동 지정",
                value=False,
                disabled=is_locked,
                help="기본은 업로드한 견적서 행 정보를 바탕으로 자동 인식합니다.",
            )
            if use_manual_cost_category:
                manual_cost_category = st.selectbox(
                    "표준원가 체계",
                    engine.get_standard_categories(),
                    key="manual_cost_category_override",
                    disabled=is_locked,
                )
                st.caption("자동 인식이 잘못될 때만 수동 지정값을 사용합니다.")
            else:
                st.caption("카테고리/비목/품명/등급/단위를 조합해 원가 체계를 자동 인식합니다.")

        st.markdown("---")
        st.markdown("##### 수동 재무 기준 업로드")
        st.caption("API 대신 직접 내려받은 재무자료를 업로드해 비목 적정성 기준을 만듭니다.")
        business_cert_file = st.file_uploader(
            "사업자등록증",
            type=BUSINESS_CERT_UPLOAD_TYPES,
            key=f"business_cert_uploader_{st.session_state['uploader_id']}",
            help="사업자등록증을 업로드하면 OCR로 협력사 기본 정보를 자동 입력합니다.",
        )
        if business_cert_file is not None:
            cert_bytes = business_cert_file.getvalue()
            cert_digest = hashlib.sha1(cert_bytes).hexdigest()
            if cert_digest != st.session_state.get("business_certificate_digest"):
                try:
                    business_certificate_info = extract_business_certificate_info(business_cert_file)
                    st.session_state["business_certificate_digest"] = cert_digest
                    st.session_state["business_certificate_info"] = business_certificate_info
                    if business_certificate_info.get("company_name"):
                        st.session_state["supplier_name_input"] = business_certificate_info["company_name"]
                        supplier_name = business_certificate_info["company_name"]
                    if business_certificate_info.get("business_id"):
                        st.session_state["supplier_business_id_input"] = business_certificate_info["business_id"]
                        supplier_business_id = business_certificate_info["business_id"]
                    if business_certificate_info.get("location"):
                        st.session_state["supplier_location_input"] = business_certificate_info["location"]
                        supplier_location = business_certificate_info["location"]
                    summary_lines = []
                    if business_certificate_info.get("business_type"):
                        summary_lines.append(f"업태: {business_certificate_info['business_type']}")
                    if business_certificate_info.get("business_item"):
                        summary_lines.append(f"종목: {business_certificate_info['business_item']}")
                    if business_certificate_info.get("industry_group"):
                        summary_lines.append(
                            f"추정 산업군: {business_certificate_info['industry_label']} ({business_certificate_info['industry_group']})"
                        )
                    st.session_state["supplier_doc_summary_input"] = " / ".join(summary_lines)
                    supplier_doc_summary = st.session_state["supplier_doc_summary_input"]
                except Exception as business_cert_error:
                    st.error(f"사업자등록증 OCR 실패: {business_cert_error}")

        business_certificate_info = st.session_state.get("business_certificate_info")
        if business_certificate_info:
            st.caption(
                f"OCR 인식: {business_certificate_info.get('summary_text') or '주요 정보 추출 완료'}"
            )
            st.caption(
                f"산업군 자동 추정 근거: {business_certificate_info.get('industry_reason', '-')}"
            )

        supplier_name = st.text_input(
            "협력사명",
            value=supplier_name,
            key="supplier_name_input",
            disabled=is_locked,
        )
        supplier_business_id = st.text_input(
            "사업자등록번호",
            value=supplier_business_id,
            key="supplier_business_id_input",
            disabled=is_locked,
        )
        supplier_location = st.text_input(
            "공장 소재지",
            value=supplier_location,
            key="supplier_location_input",
            disabled=is_locked,
        )
        supplier_listing_status = st.selectbox(
            "상장 상태",
            ["비상장", "상장"],
            index=0 if supplier_listing_status != "상장" else 1,
            key="supplier_listing_status_input",
            disabled=is_locked,
        )
        supplier_doc_summary = st.text_area(
            "협력사 설명/보충 메모",
            value=supplier_doc_summary,
            key="supplier_doc_summary_input",
            height=100,
            disabled=is_locked,
            help="외주 구조, 현장 설치, 프로젝트형 수행 등 정성 정보를 적어둘 수 있습니다.",
        )

        widget_financial_files = st.file_uploader(
            "재무제표",
            type=FINANCIAL_UPLOAD_TYPES,
            key=f"financial_uploader_bundle_{st.session_state['uploader_id']}",
            accept_multiple_files=True,
            help="재무상태표, 손익계산서, 현금흐름표, 제조원가명세서를 한 곳에 업로드해 두면 `분석 실행` 시 파일명/본문을 보고 자동 분류합니다.",
        )
        using_cached_financial_files = False
        if widget_financial_files:
            cached_financial_payloads = []
            for uploaded_file in widget_financial_files:
                cached_financial_payloads.append(
                    {"name": uploaded_file.name, "data": uploaded_file.getvalue()}
                )
            st.session_state["financial_upload_cache"] = cached_financial_payloads
            uploaded_financial_files = [
                BufferedUpload(cached["name"], cached["data"])
                for cached in cached_financial_payloads
            ]
        elif st.session_state.get("financial_upload_cache"):
            uploaded_financial_files = [
                BufferedUpload(cached["name"], cached["data"])
                for cached in st.session_state.get("financial_upload_cache", [])
            ]
            using_cached_financial_files = True
        else:
            uploaded_financial_files = []

        current_financial_signature = "|".join(
            sorted(_fast_upload_signature(uploaded_file) for uploaded_file in (uploaded_financial_files or []))
        )
        if current_financial_signature != st.session_state.get('financial_upload_signature', ""):
            st.session_state['financial_upload_signature'] = current_financial_signature
            st.session_state.pop('supplier_reference_ctx', None)
            st.session_state.pop('sidebar_selected_baseline_period_rev', None)
            st.session_state['financial_classification_signature'] = ""
            st.session_state['financial_classification_cache'] = None
            supplier_reference_ctx = None
        uploaded_financial_count = len(uploaded_financial_files or [])
        reference_ready = bool(supplier_reference_ctx)
        recognized_statement_count = len((supplier_reference_ctx or {}).get("available_statement_types", []))

        latest_detected_map = dict(((supplier_reference_ctx or {}).get("classification_meta") or {}).get("detected_map") or {})
        if latest_detected_map:
            recognized_text = ", ".join(
                f"{FINANCIAL_UPLOAD_LABELS.get(statement_type, statement_type)}={filename}"
                for statement_type, filename in latest_detected_map.items()
            )
            st.caption(f"자동 분류(최근 분석 기준): {recognized_text}")
        if using_cached_financial_files:
            st.caption("업로드 파일이 세션 캐시에서 복원되었습니다. 파일 목록이 비어 보이면 다시 업로드하거나 아래 초기화 버튼으로 캐시를 비워 주세요.")
        if uploaded_financial_count and not reference_ready:
            st.caption("재무제표는 업로드만 완료했고, 자동 분류와 재무 기준 생성은 `분석 실행` 시 한 번에 수행합니다.")

        st.markdown("##### 재무 기준 준비 상태")
        status_col1, status_col2, status_col3 = st.columns(3)
        with status_col1:
            st.metric("업로드", f"{uploaded_financial_count}건")
        with status_col2:
            st.metric("자동 분류", f"{recognized_statement_count}종" if reference_ready else "대기")
        with status_col3:
            st.metric("재무 기준", "완료" if reference_ready else "분석 시 생성")

        if not uploaded_financial_count:
            st.caption("1. 재무제표 업로드")
        else:
            st.caption("1. 업로드 완료 · 2. 견적서 첨부 후 `분석 실행` 시 자동 분류/재무 기준 생성")

        def _get_financial_classification():
            cached_signature = st.session_state.get("financial_classification_signature", "")
            cached_classification = st.session_state.get("financial_classification_cache")
            if (
                current_financial_signature
                and current_financial_signature == cached_signature
                and cached_classification is not None
            ):
                return cached_classification

            classification_meta = classify_financial_statement_uploads(uploaded_financial_files)
            st.session_state["financial_classification_signature"] = current_financial_signature
            st.session_state["financial_classification_cache"] = classification_meta
            return classification_meta

        def _build_supplier_reference_ctx():
            financial_classification = _get_financial_classification()
            if not financial_classification["statement_files"]:
                raise ValueError("업로드한 재무제표를 자동 분류하지 못했습니다. 파일명 또는 본문 형식을 확인해 주세요.")
            return _prepare_uploaded_supplier_reference(
                supplier_name,
                (business_certificate_info or {}).get("industry_group"),
                financial_classification["statement_files"],
                business_info=business_certificate_info,
                statement_period_candidates=financial_classification["detected_years"],
                classification_meta=financial_classification,
            )

        if st.button(
            "기준 초기화",
            use_container_width=True,
            disabled=not reference_ready,
        ):
            st.session_state.pop('supplier_reference_ctx', None)
            st.session_state.pop('sidebar_selected_baseline_period_rev', None)
            safe_rerun()
        if using_cached_financial_files and st.button("업로드 캐시 비우기", use_container_width=True):
            st.session_state["financial_upload_cache"] = []
            st.session_state["financial_upload_signature"] = ""
            st.session_state["financial_classification_signature"] = ""
            st.session_state["financial_classification_cache"] = None
            st.session_state.pop('supplier_reference_ctx', None)
            st.session_state.pop('sidebar_selected_baseline_period_rev', None)
            safe_rerun()

        if supplier_reference_ctx:
            period_options = supplier_reference_ctx.get("period_keys", [])
            default_period = supplier_reference_ctx.get("selected_period")
            default_index = period_options.index(default_period) if default_period in period_options else 0
            selected_baseline_period = st.selectbox(
                "분석 기준 연도",
                period_options,
                index=default_index,
                key="sidebar_selected_baseline_period_rev",
                disabled=is_locked,
            )
            if selected_baseline_period != supplier_reference_ctx.get("selected_period"):
                supplier_reference_ctx["selected_period"] = selected_baseline_period
                supplier_reference_ctx["financial_ratios"] = _extract_financial_ratios_from_ingest(
                    supplier_reference_ctx["workbook_path"],
                    supplier_reference_ctx["industry_group"],
                    period_options,
                    selected_period=selected_baseline_period,
                )
                st.session_state["supplier_reference_ctx"] = supplier_reference_ctx
            st.caption(f"기준 파일: {supplier_reference_ctx.get('uploaded_summary', '-')}")
            st.caption(
                f"rows_written={int(supplier_reference_ctx.get('rows_written', 0))} / "
                f"restatement={bool(supplier_reference_ctx.get('restatement_detected'))}"
            )
            if supplier_reference_ctx.get("fallback_note"):
                st.caption(f"기준 생성 방식: {supplier_reference_ctx.get('fallback_note')}")
        elif uploaded_financial_count:
            st.caption("분석 실행 시 파일명/본문을 기준으로 자동 분류하고 재무 기준을 생성합니다.")
        
        st.markdown("---")
        if has_result and st.button("분석 결과 초기화", use_container_width=True):
            chatbot = st.session_state.get('chatbot')
            if chatbot is not None and getattr(chatbot, 'nlp_engine', None) is not None:
                chatbot.nlp_engine.clear_fingerprints()
            st.session_state['analysis_done'] = False
            st.session_state['res'] = None
            st.session_state['ai_generated_strategy'] = None
            st.session_state['estimate_parse_signature'] = ""
            st.session_state['estimate_parse_cache'] = None
            st.session_state['estimate_parse_meta_cache'] = None
            
            st.session_state['uploader_id'] += 1
            safe_rerun()

    def _compose_analysis_meta(active_reference_ctx, active_period):
        current_fin_ratios = dict((active_reference_ctx or {}).get("financial_ratios") or {})
        current_industry_group = active_reference_ctx["industry_group"] if active_reference_ctx else (business_certificate_info or {}).get("industry_group", "")
        reference_mode = active_reference_ctx.get("mode", "") if active_reference_ctx else ""
        reference_source_label = ""
        if active_reference_ctx:
            reference_source_label = (
                active_reference_ctx.get("uploaded_summary")
                or active_reference_ctx.get("fallback_note")
                or active_reference_ctx.get("industry_label", "")
            )
        return {
            "company_name": supplier_name,
            "business_id": supplier_business_id,
            "location": supplier_location,
            "industry_group": current_industry_group,
            "industry_group_name": (
                f"{supplier_name} 업로드 재무기준"
                if supplier_name
                else (business_certificate_info or {}).get("industry_label", "")
            ),
            "baseline_period": active_period,
            "industry_code": current_industry_group,
            "industry_db_path": active_reference_ctx["workbook_path"] if active_reference_ctx else None,
            "period_basis": active_period,
            "scope": supplier_listing_status,
            "quote_subtype": "자동",
            "company_doc_summary": supplier_doc_summary,
            "manual_tags": [],
            "reference_mode": reference_mode,
            "reference_source_label": reference_source_label,
            "financial_ratios": current_fin_ratios,
            "observed_material_ratio": current_fin_ratios.get("observed_material_ratio"),
            "observed_labor_ratio": current_fin_ratios.get("observed_labor_ratio"),
            "observed_expense_ratio": current_fin_ratios.get("observed_expense_ratio"),
            "observed_gna_ratio": current_fin_ratios.get("observed_gna_ratio"),
            "observed_profit_ratio": current_fin_ratios.get("observed_profit_ratio"),
            "observed_rnd_ratio": current_fin_ratios.get("observed_rnd_ratio"),
            "sales_growth_yoy": current_fin_ratios.get("sales_growth_yoy"),
            "op_margin_delta_yoy": current_fin_ratios.get("op_margin_delta_yoy"),
            "sales_growth_qoq": current_fin_ratios.get("sales_growth_qoq"),
            "op_margin_delta_qoq": current_fin_ratios.get("op_margin_delta_qoq"),
        }

    context_text = ""
    template_col, upload_col = st.columns([1, 1.3])
    with template_col:
        if 'excel_template_bytes' not in st.session_state:
            st.session_state['excel_template_bytes'] = generate_standard_estimate_template()
            
        st.download_button(
            label="표준 양식 다운로드 (Template)",
            data=st.session_state['excel_template_bytes'],
            file_name="Standard_Estimate_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="분석에 최적화된 표준 엑셀 양식을 다운로드합니다."
        )

    with upload_col:
        unique_key = f"excel_uploader_{st.session_state['uploader_id']}"
        uploaded_file = st.file_uploader("견적서 업로드 (Excel)", type=['xlsx'], key=unique_key)
        if uploaded_file is not None:
            st.caption(f"견적서 업로드 완료: {uploaded_file.name} · 분석 실행 시 자동 인식합니다.")

    def _load_prepared_estimate():
        if uploaded_file is None:
            raise ValueError("견적서를 먼저 업로드해 주세요.")
        file_bytes = uploaded_file.getvalue()
        estimate_signature = _fast_upload_signature(uploaded_file)
        cached_signature = st.session_state.get("estimate_parse_signature", "")
        cached_df = st.session_state.get("estimate_parse_cache")
        cached_meta = st.session_state.get("estimate_parse_meta_cache")
        if estimate_signature == cached_signature and cached_df is not None:
            return cached_df.copy(), dict(cached_meta or {})

        prepared_df, parse_meta = load_uploaded_estimate_data(file_bytes)
        st.session_state["estimate_parse_signature"] = estimate_signature
        st.session_state["estimate_parse_cache"] = prepared_df.copy()
        st.session_state["estimate_parse_meta_cache"] = dict(parse_meta or {})
        return prepared_df.copy(), parse_meta

    def _run_quote_analysis(active_reference_ctx, active_selected_period):
        active_df_input, active_upload_parse_meta = _load_prepared_estimate()
        analysis_meta = _compose_analysis_meta(active_reference_ctx, active_selected_period)
        res, err = engine.analyze(
            active_df_input.copy(),
            context_text=context_text,
            analysis_meta=analysis_meta,
            category_override=manual_cost_category or None,
        )
        if active_upload_parse_meta:
            res['upload_parse_summary'] = active_upload_parse_meta.get('summary', '')
            res['upload_parse_notes'] = active_upload_parse_meta.get('notes', [])
        if active_reference_ctx:
            res['uploaded_reference_summary'] = active_reference_ctx.get('uploaded_summary', '')
            res['uploaded_reference_company'] = active_reference_ctx.get('company_name', '')
            res['uploaded_reference_period'] = active_selected_period
        st.session_state['res'] = res
        st.session_state['chatbot'] = SecureHardcodedChatbot(res)
        st.session_state['analysis_done'] = True
        return res

    can_run_analysis = bool(uploaded_file) and (bool(uploaded_financial_count) or bool(supplier_reference_ctx))
    can_run_quick = bool(uploaded_file) and not bool(supplier_reference_ctx) and bool((business_certificate_info or {}).get("industry_group"))
    if uploaded_file and uploaded_financial_count and not supplier_reference_ctx:
        st.caption("`분석 실행`을 누르면 견적서 자동 인식과 재무 기준 생성을 한 번에 수행합니다.")
    elif uploaded_file and not uploaded_financial_count and not supplier_reference_ctx:
        st.caption("재무제표를 업로드한 뒤 `분석 실행`을 누르거나, 사업자등록증 기반 `빠른 분석`을 사용해 주세요.")
        if not (business_certificate_info or {}).get("industry_group"):
            st.caption("빠른 분석을 쓰려면 사업자등록증을 업로드해 산업군 자동 추정을 먼저 완료해 주세요.")
    elif supplier_reference_ctx and not uploaded_file:
        st.caption("견적서를 업로드하면 분석 버튼이 활성화됩니다.")

    analysis_btn_col1, analysis_btn_col2 = st.columns(2)
    with analysis_btn_col1:
        run_full_analysis = st.button("분석 실행 (Analyze)", use_container_width=True, disabled=not can_run_analysis)
    with analysis_btn_col2:
        run_quick_analysis = st.button("⚡ 빠른 분석 (업종 기본 기준)", use_container_width=True, disabled=not can_run_quick)

    if run_full_analysis:
        try:
            if not uploaded_financial_count and not supplier_reference_ctx:
                st.error("좌측 설정 바의 재무제표 업로드에 파일을 1개 이상 추가해 주세요.")
                st.stop()
            with st.spinner("견적서와 재무제표를 함께 준비하고 분석하는 중입니다..."):
                active_reference_ctx = supplier_reference_ctx or _build_supplier_reference_ctx()
                active_selected_period = (
                    st.session_state.get('sidebar_selected_baseline_period_rev')
                    or active_reference_ctx.get('selected_period')
                )
                if active_selected_period not in (active_reference_ctx.get("period_keys") or []):
                    active_selected_period = active_reference_ctx.get("selected_period")
                if active_selected_period != active_reference_ctx.get("selected_period"):
                    active_reference_ctx["selected_period"] = active_selected_period
                    active_reference_ctx["financial_ratios"] = _extract_financial_ratios_from_ingest(
                        active_reference_ctx["workbook_path"],
                        active_reference_ctx["industry_group"],
                        active_reference_ctx.get("period_keys", []),
                        selected_period=active_selected_period,
                    )
                st.session_state['supplier_reference_ctx'] = active_reference_ctx
                st.session_state['sidebar_selected_baseline_period_rev'] = active_selected_period
                _run_quote_analysis(active_reference_ctx, active_selected_period)
            safe_rerun()
        except Exception as e: st.error(f"오류 발생: {e}")

    if run_quick_analysis:
        try:
            with st.spinner("사업자등록증 기반 빠른 분석을 준비하는 중입니다..."):
                quick_reference_ctx = build_quick_industry_reference_ctx(
                    supplier_name,
                    business_certificate_info,
                )
                st.session_state['supplier_reference_ctx'] = quick_reference_ctx
                st.session_state['sidebar_selected_baseline_period_rev'] = quick_reference_ctx['selected_period']
                _run_quote_analysis(quick_reference_ctx, quick_reference_ctx['selected_period'])
            safe_rerun()
        except Exception as e:
            st.error(f"빠른 분석 실행 실패: {e}")

    if 'res' in st.session_state and st.session_state['res']:
        res = st.session_state['res']
        if res['validation_errors']: st.error(res['validation_errors'][0])

        adequacy_report = res.get('adequacy_report')
        def render_adequacy_reference_section():
            if not adequacy_report:
                return

            sum_check = next((c for c in adequacy_report['structure_checks'] if c['code'].startswith('SUM_RATIO')), None)
            overall_label_kr = ADEQUACY_JUDGEMENT_LABELS.get(
                str(adequacy_report.get('overall_judgement', adequacy_report.get('overall_label', ''))).lower(),
                str(adequacy_report.get('overall_label', '-')),
            )
            kpi_a, kpi_b, kpi_c, kpi_d = st.columns(4)
            with kpi_a:
                st.markdown(
                    f"<div class='metric-container'><div class='metric-label'>종합 판정</div><div class='metric-value' style='color:#C2410C'>{overall_label_kr}</div></div>",
                    unsafe_allow_html=True,
                )
            with kpi_b:
                st.markdown(
                    f"<div class='metric-container'><div class='metric-label'>비목 합계</div><div class='metric-value'>{adequacy_report['total_ratio']:.2f}%</div></div>",
                    unsafe_allow_html=True,
                )
            with kpi_c:
                st.markdown(
                    f"<div class='metric-container'><div class='metric-label'>비교 기준</div><div class='metric-value' style='font-size:18px'>{adequacy_report['profile_label']}</div></div>",
                    unsafe_allow_html=True,
                )
            with kpi_d:
                st.markdown(
                    f"<div class='metric-container'><div class='metric-label'>신뢰도</div><div class='metric-value'>{CONFIDENCE_LABELS.get(str(adequacy_report['base_confidence_bucket']), str(adequacy_report['base_confidence_bucket']))}</div></div>",
                    unsafe_allow_html=True,
                )

            col_meta, col_findings = st.columns([1, 1.3], gap="large")
            with col_meta:
                st.markdown('<span class="component-title">핵심 해석</span>', unsafe_allow_html=True)
                st.info(adequacy_report['overall_summary'])
                if adequacy_report['tag_rows']:
                    tag_text = " / ".join(f"{row['label']}({row['confidence']:.2f})" for row in adequacy_report['tag_rows'])
                    st.caption(f"태그: {tag_text}")
                if sum_check:
                    level = sum_check['severity']
                    if level == 'ok':
                        st.success(sum_check['message'])
                    elif level == 'watch':
                        st.warning(sum_check['message'])
                    else:
                        st.error(sum_check['message'])
                st.caption(res.get('cost_reference_note', ''))
                if res.get('cost_reference_coverage'):
                    st.caption(res['cost_reference_coverage'])
            with col_findings:
                st.markdown('<span class="component-title">주요 발견사항</span>', unsafe_allow_html=True)
                if adequacy_report['findings']:
                    for finding in adequacy_report['findings']:
                        severity_label = ADEQUACY_JUDGEMENT_LABELS.get(str(finding.get('severity', '')).lower(), str(finding.get('severity', '-')))
                        st.markdown(
                            f"<div class='insight-card'><div class='insight-title'>[{severity_label}] {finding['title']}</div>{finding['description']}</div>",
                            unsafe_allow_html=True,
                        )
                else:
                    st.success("기준 범위를 벗어난 주요 비목이 없습니다.")
                with st.expander("분석 한계 / 구조 룰", expanded=False):
                    for check in adequacy_report['structure_checks']:
                        level_label = ADEQUACY_JUDGEMENT_LABELS.get(str(check['severity']).lower(), str(check['severity']))
                        st.write(f"[{level_label}] {check['code']} - {check['message']}")
                    for note in adequacy_report['limitations']:
                        st.write(f"- {note}")
                    if res.get('cost_reference_notes'):
                        st.write("- standard_db 참조 산출 메모")
                        for note in res['cost_reference_notes']:
                            st.write(f"  · {note}")

        if res.get('upload_parse_summary'):
            st.caption(f"업로드 자동 인식: {res['upload_parse_summary']}")
        if res.get('uploaded_reference_summary'):
            ref_company = res.get('uploaded_reference_company') or "업로드 협력사"
            ref_period = res.get('uploaded_reference_period') or "-"
            st.caption(f"수동 재무 기준: {ref_company} / {ref_period} / {res['uploaded_reference_summary']}")
        if res.get('analysis_category_summary'):
            st.info(res['analysis_category_summary'])
            group_mix_rows = res.get('analysis_category_group_mix') or []
            if group_mix_rows:
                group_summary = ", ".join(
                    f"{row.get('group_label', row.get('group_id', '-'))} {float(row.get('share_pct', 0.0)):.1f}%"
                    for row in group_mix_rows[:3]
                )
                st.caption(f"상위 그룹 기준: {group_summary}")
            mix_rows = res.get('analysis_category_mix') or []
            if mix_rows:
                with st.expander("원가 체계 인식 상세", expanded=False):
                    mix_df = pd.DataFrame(
                        [
                            {
                                "원가 체계": row["category"],
                                "비중": f"{float(row['share_pct']):.1f}%",
                                "금액 기준": f"{float(row['amount']):,.0f}",
                                "행 수": int(row["row_count"]),
                                "판별 근거": row["reason"],
                            }
                            for row in mix_rows
                        ]
                    )
                    mix_df.insert(0, "상위 그룹", [row.get("group_label", row.get("category", "-")) for row in mix_rows])
                    st.dataframe(mix_df, use_container_width=True, hide_index=True)
        if res.get('target_policy_notes'):
            with st.expander("자동 목표 산정 기준", expanded=False):
                st.write(f"- 현재 적용 표준원가체계: {res.get('analysis_category_display', '-')}")
                for note in res['target_policy_notes']:
                    st.write(f"- {note}")
        
        # [Section 1] Cost Analysis
        st.markdown('<div class="section-header-l1">01. 비용 구조 분석 (Cost Breakdown)</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-header-l2">1. 종합 원가 지표 (Total Cost Overview)</div>', unsafe_allow_html=True)
        if res.get('target_labor_basis_note'):
            st.caption(res['target_labor_basis_note'])
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        
        with col_m1: 
            st.markdown(f"<div class='metric-container'><div class='metric-label'>협력사 제출 (Quote)</div><div class='metric-value'>{res['v_total']:,.0f}</div></div>", unsafe_allow_html=True)
        with col_m2: 
            st.markdown(f"<div class='metric-container'><div class='metric-label'>당사 목표 (Internal)</div><div class='metric-value' style='color:#555555'>{res['t_total']:,.0f}</div></div>", unsafe_allow_html=True)
        with col_m3: 
            st.markdown(f"<div class='metric-container'><div class='metric-label'>현실적 타결 (Target)</div><div class='metric-value' style='color:#1E3A8A'>{res['r_total']:,.0f}</div></div>", unsafe_allow_html=True)
        with col_m4: 
            st.markdown(f"<div class='metric-container'><div class='metric-label'>절감 예상 (Savings)</div><div class='metric-value' style='color:#28A745'>{((res['v_total']-res['r_total'])/res['v_total']*100):.1f}%</div></div>", unsafe_allow_html=True)

        st.write("")
        st.write("") 

        st.markdown('<div class="section-header-l2">2. 비목별 상세 분석표 (Detailed Breakdown)</div>', unsafe_allow_html=True)
        st.caption(res.get('cost_reference_note', ''))
        if res.get('cost_reference_coverage'):
            st.caption(res['cost_reference_coverage'])
        c1, c2 = st.columns([1.2, 1.8])
        with c1: 
            st.markdown('<span class="component-title">비용 절감 구조 분석 (Chart)</span>', unsafe_allow_html=True)
            st.plotly_chart(WaterfallChartBuilder.draw(res['v_total'], res['r_total'], res['summary']), use_container_width=True)
        with c2: 
            st.markdown('<span class="component-title">상세 견적 비교 (Detailed Breakdown)</span>', unsafe_allow_html=True)
            st.dataframe(res['summary'], use_container_width=True, height=300)

        auto_points = build_auto_negotiation_points(res)
        st.markdown('<div class="section-header-l2">3. 자동 협상 포인트</div>', unsafe_allow_html=True)
        if auto_points:
            for point in auto_points:
                st.markdown(
                    f"""
                    <div class="insight-card">
                        <div class="insight-title">{point['title']}</div>
                        {point['body']}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        else:
            st.success("자동 추출된 우선 협상 포인트가 없습니다.")

        if res.get('should_cost_summary'):
            st.markdown('<div class="section-header-l2">4. 재무 역산 기반 Should-Cost</div>', unsafe_allow_html=True)
            st.info(res['should_cost_summary'])

        category_display = res.get('analysis_category_display') or res.get('analysis_category', '1.일반용역')

        st.markdown('<div class="section-header-l1">02. AI 협상 전략 수립 (EXAONE LLM)</div>', unsafe_allow_html=True)
        default_role = (
            "당신은 구매/재무 분석 AI입니다.\n"
            "업종 기준선(aggregate)과 기업 문서 태그를 사용해 견적 비목의 적정성을 판정하고, "
            "observed / repaired / proxy / estimated basis_type, confidence, reason_code를 명시하는 검토자입니다.\n"
            "raw 분포가 없으면 median/Q1/Q3/IQR을 만들지 말고, QoQ/YoY 변화, 합계 정합성, 구조 룰로 보강 판정해야 합니다."
        )
        default_request = (
            "아래 제공된 [협력사 재무 역산 컨텍스트], [구조화 판정 데이터], [비용 구조 분석 결과], "
            "[Should-Cost 역산 결과], [자동 협상 포인트], [전략 참고 정보]를 종합하여 "
            "5개 섹션 구조의 협상 리포트를 작성해 주세요."
        )

        ai_role = default_role
        ai_request = default_request

        exaone_dependency_error = get_exaone_runtime_error()
        if exaone_dependency_error:
            st.warning(exaone_dependency_error)

        st.info("현재 분석 결과와 자동 협상 포인트를 바탕으로 EXAONE 보고서를 생성합니다.")

        if st.button("🤖 EXAONE 보고서 작성", use_container_width=True):
            if exaone_dependency_error:
                st.error(exaone_dependency_error)
                st.stop()

            cost_summary_text = "구분 | 협력사 제출 | 공시기준가(업로드 재무기준) | 당사 목표 | 현실적 타결 | Gap 금액 | 협상 우선순위\n---|---|---|---|---|---|---\n"
            if st.session_state.get('res') is not None:
                summary_df = st.session_state['res']['summary']
                max_gap = 0.0
                for _, row in summary_df.iterrows():
                    if row.get('구분', '') == '합계':
                        continue
                    gap_val = parse_amount_cell(row.get('1. 협력사 제출', '0')) - parse_amount_cell(row.get('4. 현실적 타결', '0'))
                    max_gap = max(max_gap, gap_val)

                for _, row in summary_df.iterrows():
                    gubun = row.get('구분', '')
                    v_submit = row.get('1. 협력사 제출', '0')
                    v_public = row.get('2. 공시기준가', '0')
                    v_internal = row.get('3. 당사 목표', '0')
                    v_target = row.get('4. 현실적 타결', '0')
                    gap_val = parse_amount_cell(v_submit) - parse_amount_cell(v_target)
                    gap_str = f"{gap_val:,.0f}원" if gap_val > 0 else "-"
                    if gubun == '합계':
                        priority_str = ""
                    elif gap_val == max_gap and max_gap > 0:
                        priority_str = "🚨 1순위 타겟"
                    elif gap_val > 0:
                        priority_str = "⚠️ 단가 인하 요망"
                    else:
                        priority_str = "✅ 경쟁력 우수"
                    cost_summary_text += f"{gubun} | {v_submit} | {v_public} | {v_internal} | {v_target} | **{gap_str}** | {priority_str}\n"
                cost_summary_text += f"\n* 총 절감 예상액(Total Gap): {st.session_state['res']['save_amt']:,.0f}원"

            auto_points_text = ""
            for idx, point in enumerate(auto_points, start=1):
                auto_points_text += f"{idx}. {point['title']}\n   - {point['body']}\n"
            if not auto_points_text:
                auto_points_text = "자동 추출된 우선 협상 포인트 없음"

            strategy_reference_text = "전략 코치 입력 없이 현재 분석 결과와 자동 협상 포인트를 바탕으로 협상 전략을 직접 구성해 주세요."

            adequacy_context_text = ""
            if st.session_state['res'].get('adequacy_report'):
                adequacy_context_text = build_llm_context(st.session_state['res']['adequacy_report'])

            financial_forensic_text = ""
            should_cost_text = ""
            supplier_ref = st.session_state.get("supplier_reference_ctx")
            if supplier_ref:
                should_cost = st.session_state["res"].get("should_cost", {})
                financial_forensic_text = build_financial_forensic_context(supplier_ref, should_cost)
                should_cost_text = st.session_state["res"].get("should_cost_summary", "")
            financial_forensic_text = financial_forensic_text or "재무 역산 컨텍스트 없음"
            should_cost_text = should_cost_text or "재무 실측 비율이 부족해 Should-Cost 역산 결과를 생성하지 못했습니다."

            output_format_instruction = """
[출력 형식 지시]
- 반드시 위의 5개 섹션 구조(### 1~5)를 순서대로 작성하세요.
- 비목별 비교는 마크다운 표(|)로 작성하세요.
- 협상 발화 예시는 "큰따옴표"로 감싸 명확히 구분하세요.
- 수치는 원 단위 또는 %p 단위로 구체적으로 제시하세요.
- 불확실한 추정치는 반드시 "추정:" 접두어를 붙이세요.
""".strip()

            final_prompt = (
                f"{ai_role}\n\n[요청 사항]\n{ai_request}\n\n"
                f"[협력사 재무 역산 컨텍스트]\n{financial_forensic_text}\n\n"
                f"[구조화 판정 데이터]\n{adequacy_context_text}\n\n"
                f"[비용 구조 분석 결과]\n{cost_summary_text}\n\n"
                f"[Should-Cost 역산 결과]\n{should_cost_text}\n\n"
                f"[자동 협상 포인트]\n{auto_points_text}\n\n"
                f"[전략 참고 정보]\n{strategy_reference_text}\n\n"
                f"{output_format_instruction}"
            )

            prompt_signature = hashlib.sha1(
                f"{user_defined_model_path}||{final_prompt}".encode("utf-8")
            ).hexdigest()
            cached_exaone_response = st.session_state.get("exaone_response_cache", {}).get(prompt_signature)
            if cached_exaone_response:
                st.session_state['ai_generated_strategy'] = cached_exaone_response
                st.markdown(cached_exaone_response)
                st.caption("동일 입력 기준 EXAONE 캐시 결과를 재사용했습니다.")
            else:
                output_container = st.empty()
                full_generated_text = ""
                for chunk in ask_exaone_stream(final_prompt, user_defined_model_path):
                    full_generated_text += chunk
                    output_container.markdown(full_generated_text + "▌")

                output_container.markdown(full_generated_text)
                st.session_state.setdefault("exaone_response_cache", {})[prompt_signature] = full_generated_text
                st.session_state['ai_generated_strategy'] = full_generated_text

        elif st.session_state.get('ai_generated_strategy'):
            st.markdown(st.session_state['ai_generated_strategy'])

        # ==============================================================================
        # [Section 4] Report
        # ==============================================================================
        if adequacy_report:
            st.markdown('<div class="section-header-l1">03. 비목 적정성 리포트 (Report Generation)</div>', unsafe_allow_html=True)

            with st.expander("업종/비목 적정성 참고 판정", expanded=False):
                render_adequacy_reference_section()
            if res.get('should_cost_summary'):
                with st.expander("재무 역산 기반 Should-Cost 참고", expanded=False):
                    st.write(res['should_cost_summary'])

            html_report = generate_html_report(
                adequacy_report,
                st.session_state.get('ai_generated_strategy', ""),
            )

            with st.expander("HTML 리포트 미리보기", expanded=False):
                st.components.v1.html(html_report, height=1200, scrolling=True)

            c_d1, c_d2, c_d3 = st.columns(3)
            with c_d1:
                st.download_button(
                    "HTML 리포트 다운로드 (.html)",
                    html_report.encode("utf-8"),
                    "Quote_Adequacy_Report.html",
                    "text/html",
                    use_container_width=True,
                )

            with c_d2:
                if st.button("이미지 리포트 생성 (Preview)", use_container_width=True):
                    img = create_strategy_report_image(
                        category_display,
                        res['v_total'],
                        res['r_total'],
                        res['summary'],
                        [],
                        st.session_state.get('ai_generated_strategy', "")
                    )
                    st.image(img, caption="Generated Report Preview")

                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    st.download_button("이미지 다운로드 (.png)", buf.getvalue(), "Strategy_Report.png", "image/png", use_container_width=True)

            with c_d3:
                excel_data = create_excel_report(
                    res['summary'],
                    category_display,
                    res['bottom_line'],
                    res['draft_text'],
                    [],
                    st.session_state.get('ai_generated_strategy', "")
                )
                st.download_button("엑셀 리포트 다운로드 (.xlsx)", excel_data, "Buying_Report_Expert.xlsx", use_container_width=True)

if __name__ == "__main__":
    main()
