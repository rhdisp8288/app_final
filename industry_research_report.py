from __future__ import annotations

from datetime import datetime
from html import escape
from typing import Any, Dict, List, Optional

import math
import pandas as pd

from industry_financial_db import get_baseline_items, load_industry_db, resolve_workbook_path


ITEM_ORDER = ["material", "labor", "expense", "gna", "profit", "rnd"]
ITEM_LABELS = {
    "material": "재료비",
    "labor": "노무비",
    "expense": "경비",
    "gna": "일반관리비",
    "profit": "이윤",
    "rnd": "R&D",
}
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
COLOR_MAP = {
    "material": "#0f69ff",
    "labor": "#0f8f83",
    "expense": "#f39a1e",
    "gna": "#7b4fd1",
    "profit": "#d93c45",
    "rnd": "#0d766e",
}
REFERENCE_SOURCES = [
    {
        "org": "한국은행",
        "title": "2024년 연간 기업경영분석",
        "url": "https://www.bok.or.kr/portal/bbs/B0000559/view.do?menuNo=200690&nttId=10092387",
        "note": "업종별 수익성 및 비용구조 공식 앵커 참고",
    },
    {
        "org": "한국은행",
        "title": "2023년 기업경영분석 통계편",
        "url": "https://www.bok.or.kr/fileSrc/portal/bf513c6fcd2e44a7b3e49795602bf796/1/8ba0be5b2a5e4287b575e6dbfb130126.pdf",
        "note": "업종별 손익 및 비목 비율 cross-check 참고",
    },
    {
        "org": "KOSIS",
        "title": "광업제조업조사",
        "url": "https://kosis.kr/index/index.do",
        "note": "제조업 원재료비 및 급여 계열 보정 참고",
    },
    {
        "org": "KOSIS",
        "title": "서비스업조사 / 기업활동조사",
        "url": "https://kosis.kr/index/index.do",
        "note": "서비스업 인건비 및 영업비용 구조 보정 참고",
    },
]
JUDGEMENT_GUIDE = [
    {"code": "ok", "label": "적정", "description": "실제 비중이 목표 범위 안에 있습니다."},
    {"code": "watch", "label": "주의", "description": "범위를 소폭 벗어나 추가 확인이 필요한 상태입니다."},
    {"code": "alert", "label": "경고", "description": "의미 있게 벗어나 재검토가 필요한 상태입니다."},
    {"code": "critical", "label": "심각", "description": "크게 벗어나 구조적 오류 또는 과대계상 가능성이 높습니다."},
]
BASIS_GUIDE = [
    {"label": "실측 기반", "description": "재무제표 또는 제조원가명세서에서 직접 관측한 값입니다."},
    {"label": "업종 프록시", "description": "직접 근거가 약해 업종 평균 믹스를 활용한 값입니다."},
    {"label": "보수 추정", "description": "원시 추출값에 오류 가능성이 있어 회계식 또는 공식기준으로 보정한 값입니다."},
]
CALC_RULE_LABELS = {
    "IS_COGS_TO_REVENUE": "손익계산서 매출원가 직접 사용",
    "IS_SGA_TO_REVENUE": "손익계산서 판관비 직접 사용",
    "IS_OPERATING_PROFIT_TO_REVENUE": "손익계산서 영업이익 직접 사용",
    "IS_RND_TO_REVENUE": "손익계산서 R&D 직접 사용",
    "COST_DIRECT_TO_REVENUE": "제조원가명세서 직접 항목 사용",
    "COST_SCALED_TO_REPAIRED_COGS": "직접비 합계를 보정 매출원가에 비례 배분",
    "OFFICIAL_OR_PROXY_MIX_TO_REPAIRED_COGS": "공식기준/프록시 믹스를 보정 매출원가에 배분",
    "PROXY_MIX_FROM_COGS": "업종 프록시 믹스를 매출원가에 배분",
    "REPAIRED_COGS_FROM_GROSS_PROFIT": "매출총이익으로 매출원가 복원",
    "REPAIRED_COGS_FROM_GNA_AND_PROFIT": "판관비 및 영업이익으로 매출원가 복원",
    "REPAIRED_COGS_FROM_COST_SUM": "직접비 합계로 매출원가 복원",
    "REPAIRED_COGS_FROM_COST_AND_INVENTORY": "직접비와 재고 증감으로 매출원가 복원",
    "OFFICIAL_REFERENCE_COGS": "공식기준 매출원가율 사용",
    "REPAIRED_MATERIAL_FROM_NATURE_NOTE": "비용 성격 주석의 재료비로 재료비 보정",
    "LABOR_SUPPORT_FROM_EMPLOYEE_COST": "종업원급여 기준으로 노무비 보정",
    "DEPRECIATION_FLOOR_SUPPORT": "감가상각비 하한을 반영해 경비 보정",
    "PPE_CAPEX_SUPPORT": "유형자산·CAPEX 신호를 반영해 경비 보정",
    "REPAIRED_GNA_FROM_GROSS_PROFIT_AND_PROFIT": "매출총이익 및 영업이익으로 판관비 복원",
    "REPAIRED_GNA_FROM_COGS_AND_PROFIT": "매출원가 및 영업이익으로 판관비 복원",
    "OFFICIAL_REFERENCE_GNA": "공식기준 판관비율 사용",
    "REPAIRED_PROFIT_FROM_GROSS_PROFIT_AND_GNA": "매출총이익 및 판관비로 영업이익 복원",
    "REPAIRED_PROFIT_FROM_COGS_AND_GNA": "매출원가 및 판관비로 영업이익 복원",
    "OFFICIAL_REFERENCE_PROFIT": "공식기준 영업이익률 사용",
    "OUTLIER_SHRINK_TO_OFFICIAL": "공식기준 대비 아웃라이어 축소 보정",
    "CFO_QUALITY_ADJUSTED": "영업현금흐름 기반 품질 조정",
    "ESTIMATED_FROM_3Y_GROWTH_BLEND": "최근 3개년 공식기준 추세로 추정",
}


def _safe_float(value: Any, default: float = 0.0) -> float:
    try:
        number = float(value)
    except Exception:
        return default
    if math.isnan(number) or math.isinf(number):
        return default
    return number


def _format_pct(value: Any, digits: int = 2) -> str:
    return f"{_safe_float(value):.{digits}f}%"


def _format_amount(value: Any) -> str:
    return f"{_safe_float(value):,.0f}"


def _basis_label(value: Any) -> str:
    key = str(value or "").strip()
    return BASIS_TYPE_LABELS.get(key, key or "-")


def _confidence_label(value: Any) -> str:
    key = str(value or "").strip()
    return CONFIDENCE_LABELS.get(key, key or "-")


def _calc_rule_label(value: Any) -> str:
    key = str(value or "").strip()
    if not key:
        return "-"
    parts = [part.strip() for part in key.split("|") if part.strip()]
    if len(parts) > 1:
        return " + ".join(CALC_RULE_LABELS.get(part, part.replace("_", " ").title()) for part in parts)
    return CALC_RULE_LABELS.get(key, key.replace("_", " ").title())


def _period_sort_key(period_key: Any) -> tuple[int, str]:
    text = str(period_key)
    return (0, text.zfill(8)) if text.isdigit() else (1, text)


def _ratio(amount: float, revenue: float) -> float:
    return amount / revenue * 100.0 if revenue > 0 else 0.0


def _latest_statement_rows(statement_rows: pd.DataFrame, industry_group: str, period_key: str) -> pd.DataFrame:
    filtered = statement_rows[
        (statement_rows["industry_group"].astype(str) == str(industry_group))
        & (statement_rows["period_key"].astype(str) == str(period_key))
    ].copy()
    if filtered.empty:
        return filtered
    filtered["ingested_at_sort"] = pd.to_datetime(filtered["ingested_at"], errors="coerce")
    filtered = filtered.sort_values(["statement_type", "normalized_account_code", "ingested_at_sort"])
    return filtered.drop_duplicates(["statement_type", "normalized_account_code"], keep="last")


def _latest_metric_rows(metric_rows: pd.DataFrame, industry_group: str, period_key: str) -> pd.DataFrame:
    filtered = metric_rows[
        (metric_rows["industry_group"].astype(str) == str(industry_group))
        & (metric_rows["period_key"].astype(str) == str(period_key))
    ].copy()
    if filtered.empty:
        return filtered
    filtered["created_at_sort"] = pd.to_datetime(filtered["created_at"], errors="coerce")
    filtered = filtered.sort_values(["metric_code", "created_at_sort"])
    return filtered.drop_duplicates(["metric_code"], keep="last")


def _latest_baseline_rows(baseline_df: pd.DataFrame, industry_group: str) -> pd.DataFrame:
    filtered = baseline_df[baseline_df["industry_group"].astype(str) == str(industry_group)].copy()
    if filtered.empty:
        return filtered
    filtered["created_at_sort"] = pd.to_datetime(filtered["created_at"], errors="coerce")
    filtered = filtered.sort_values(["period_key", "item_code", "created_at_sort"])
    return filtered.drop_duplicates(["period_key", "item_code"], keep="last").drop(columns=["created_at_sort"])


def _latest_reference_rows(reference_df: pd.DataFrame, industry_group: str) -> pd.DataFrame:
    filtered = reference_df[reference_df["industry_group"].astype(str) == str(industry_group)].copy()
    if filtered.empty:
        return filtered
    filtered = filtered.assign(period_sort=pd.to_numeric(filtered["period_key"], errors="coerce"))
    filtered = filtered.sort_values(["period_sort", "period_key"], na_position="last")
    return filtered.drop_duplicates(["period_key"], keep="last").drop(columns=["period_sort"])


def _item_history_rows(baseline_df: pd.DataFrame, periods: List[str]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for item_code in ITEM_ORDER:
        item_rows = baseline_df[baseline_df["item_code"].astype(str) == item_code]
        if item_rows.empty:
            continue
        values = {}
        last_basis = ""
        last_conf = ""
        for period in periods:
            matched = item_rows[item_rows["period_key"].astype(str) == period]
            if matched.empty:
                values[period] = None
                continue
            row = matched.iloc[-1]
            values[period] = _safe_float(row.get("base_pct"))
            last_basis = row.get("basis_type", last_basis)
            last_conf = row.get("confidence_grade", last_conf)
        rows.append(
            {
                "item_code": item_code,
                "label": ITEM_LABELS.get(item_code, item_code),
                "values": values,
                "basis_type": _basis_label(last_basis),
                "confidence_grade": _confidence_label(last_conf),
            }
        )
    return rows


def _current_reference_map(reference_row: Dict[str, Any]) -> Dict[str, float]:
    cogs_pct = _safe_float(reference_row.get("cogs_pct"))
    return {
        "material": cogs_pct * _safe_float(reference_row.get("material_share")) / 100.0,
        "labor": cogs_pct * _safe_float(reference_row.get("labor_share")) / 100.0,
        "expense": cogs_pct * _safe_float(reference_row.get("expense_share")) / 100.0,
        "gna": _safe_float(reference_row.get("gna_pct")),
        "profit": _safe_float(reference_row.get("profit_pct")),
        "rnd": _safe_float(reference_row.get("rnd_pct")),
    }


def _build_rating(
    current_items: List[Dict[str, Any]],
    metric_rows: pd.DataFrame,
    statement_rows: pd.DataFrame,
    current_reference_row: Dict[str, Any],
) -> tuple[str, str, str]:
    basis_values = [str(row.get("basis_type", "")) for row in current_items]
    estimated_count = sum(value == "estimated" for value in basis_values)
    proxy_count = sum(value == "proxy" for value in basis_values)
    repaired_count = 0
    for _, row in metric_rows.iterrows():
        calc_rule = str(row.get("calc_rule_code", ""))
        if "REPAIRED_" in calc_rule or "OFFICIAL_REFERENCE" in calc_rule or "ESTIMATED_FROM_3Y" in calc_rule:
            repaired_count += 1

    if statement_rows.empty and str(current_reference_row.get("source", "")).startswith("estimated_from_3y_growth"):
        return ("추정 활용", "선택 연도는 최근 3개년 기준 추세를 바탕으로 추정된 baseline입니다.", "warn")
    if repaired_count >= 3 or estimated_count >= 3:
        return ("보수 활용", "원시 추출값을 회계식 또는 공식기준으로 복원한 비중이 높습니다.", "warn")
    if proxy_count >= 2 or estimated_count >= 1:
        return ("주의 활용", "일부 직접비 항목이 업종 프록시 또는 보정치에 의존합니다.", "warn")
    return ("기준 활용", "직접 관측값 비중이 높아 비교 기준으로 활용하기 적절합니다.", "good")


def _build_summary_points(
    current_items: List[Dict[str, Any]],
    diagnostics: Dict[str, Any],
    current_reference_row: Dict[str, Any],
    estimate_row: Optional[Dict[str, Any]],
    adjustment_events: List[Dict[str, Any]],
) -> List[str]:
    current_sorted = sorted(current_items, key=lambda row: _safe_float(row.get("base_pct")), reverse=True)
    top_items = ", ".join(f"{row['label']} {_format_pct(row['base_pct'])}" for row in current_sorted[:3])
    points = [f"현재 baseline 구성의 상위 비목은 {top_items} 입니다."]

    if diagnostics["statement_available"]:
        points.append(
            f"원시 추출 기준 매출원가율은 {_format_pct(diagnostics['raw_cogs_pct'])}, "
            f"직접비 합계는 {_format_pct(diagnostics['cost_sum_pct'])}로 관측되었습니다."
        )
        if diagnostics["raw_vs_repaired_gap"] >= 10:
            points.append(
                f"원시 추출과 복원 기준 간 괴리가 {_format_pct(diagnostics['raw_vs_repaired_gap'])}p 수준으로 커 "
                "공식기준 및 회계식 보정의 영향이 큽니다."
            )
    else:
        points.append("선택 연도는 직접 재무추출 데이터가 없어 공식기준 시계열 기반 추정 baseline으로 구성되었습니다.")

    points.append(
        f"현재 연도 공식기준 source는 `{current_reference_row.get('source', '-')}`이며, "
        f"매출원가 {_format_pct(current_reference_row.get('cogs_pct'))}, "
        f"판관비 {_format_pct(current_reference_row.get('gna_pct'))}, "
        f"영업이익 {_format_pct(current_reference_row.get('profit_pct'))}를 앵커로 활용합니다."
    )
    if estimate_row:
        points.append(
            f"2025 추정치는 최근 3개년 공식기준 추세를 반영해 "
            f"매출원가 {_format_pct(estimate_row.get('cogs_pct'))}, "
            f"판관비 {_format_pct(estimate_row.get('gna_pct'))}, "
            f"영업이익 {_format_pct(estimate_row.get('profit_pct'))}로 산출했습니다."
        )
    top_adjustments = [row for row in adjustment_events if "보정" in row["calc_rule_label"] or "공식기준" in row["calc_rule_label"]]
    if top_adjustments:
        points.append("보정 영향이 큰 항목은 " + ", ".join(row["label"] for row in top_adjustments[:3]) + " 입니다.")
    return points[:6]


def _build_method_notes(
    period_key: str,
    statement_rows: pd.DataFrame,
    diagnostics: Dict[str, Any],
    current_reference_row: Dict[str, Any],
    estimate_row: Optional[Dict[str, Any]],
) -> List[str]:
    notes = [
        "본 리포트의 `base_pct`는 선택 비목을 100으로 정규화한 구성비이며, `metric value_pct`는 매출액 대비 비율입니다.",
        "config_official_reference는 국가통계 앵커와 workbook 복원 이력을 결합한 연도별 표준행입니다.",
        "source가 `generated_official_standard_blend`인 경우, 국가통계 앵커와 복원 history를 혼합한 표준안입니다.",
        "source가 `estimated_from_3y_growth`인 경우, 최근 3개년 공식기준 시계열을 바탕으로 2025를 추정한 값입니다.",
        "직접비(material/labor/expense)는 COST 직접값, 보정 매출원가 배분, 공식기준/프록시 믹스 중 가장 신뢰 가능한 경로로 채택됩니다.",
    ]
    if diagnostics["statement_available"]:
        notes.append(
            f"{period_key} 원시 추출 연도는 statement_rows 기준으로 직접 관측되었으며, "
            f"매출원가-직접비 괴리는 {_format_pct(diagnostics['cost_gap_pct'])}p 입니다."
        )
    else:
        notes.append(f"{period_key}는 statement_rows가 없어 추정 baseline으로만 구성됩니다.")
    if current_reference_row.get("note"):
        notes.append(f"현재 연도 공식기준 note: {current_reference_row['note']}")
    if estimate_row and estimate_row.get("note"):
        notes.append(f"2025 추정 note: {estimate_row['note']}")
    notes.append("본 리포트는 NH투자증권 리포트 레이아웃을 차용한 내부 참고용 형식입니다.")
    return notes


def _build_house_view(
    profile_label: str,
    current_items: List[Dict[str, Any]],
    diagnostics: Dict[str, Any],
    current_reference_row: Dict[str, Any],
    estimate_row: Optional[Dict[str, Any]],
) -> Dict[str, str]:
    current_sorted = sorted(current_items, key=lambda row: _safe_float(row.get("base_pct")), reverse=True)
    top_row = current_sorted[0]
    top_label = top_row["label"]
    top_pct = _format_pct(top_row.get("base_pct"))
    material_pct = _safe_float(next((row["base_pct"] for row in current_items if row["item_code"] == "material"), 0.0))
    labor_pct = _safe_float(next((row["base_pct"] for row in current_items if row["item_code"] == "labor"), 0.0))
    gna_pct = _safe_float(next((row["base_pct"] for row in current_items if row["item_code"] == "gna"), 0.0))

    if material_pct >= 40:
        structure = f"{profile_label}은 재료비 중심 구조가 강한 업종으로 해석됩니다."
    elif labor_pct >= 32:
        structure = f"{profile_label}은 노무비 비중이 높은 인력집약형 구조가 두드러집니다."
    elif gna_pct >= 32:
        structure = f"{profile_label}은 직접비보다 일반관리비 비중이 상대적으로 큰 서비스형 구조에 가깝습니다."
    else:
        structure = f"{profile_label}은 특정 한 비목보다 복합형 비용구조가 관찰됩니다."

    if diagnostics["statement_available"] and diagnostics["raw_vs_repaired_gap"] >= 10:
        repair = (
            "이번 baseline은 원시 추출값을 그대로 채택하지 않고, "
            "공식기준 앵커와 회계식 복원을 반영해 보수적으로 재구성한 값입니다."
        )
    elif diagnostics["statement_available"]:
        repair = "원시 추출값과 복원값의 괴리가 제한적이어서, 현재 기준선은 비교적 안정적으로 읽을 수 있습니다."
    else:
        repair = "선택 연도는 직접 추출치가 없어 공식기준 시계열과 최근 3개년 추세를 통해 추정했습니다."

    outlook = (
        f"현재 연도 기준 최상위 비목은 {top_label}({top_pct})이며, "
        f"공식기준 source는 {current_reference_row.get('source', '-')} 입니다."
    )
    if estimate_row:
        outlook += (
            f" 2025 추정치는 매출원가 {_format_pct(estimate_row.get('cogs_pct'))}, "
            f"판관비 {_format_pct(estimate_row.get('gna_pct'))}, 영업이익 {_format_pct(estimate_row.get('profit_pct'))} 수준을 가정합니다."
        )

    return {
        "lede": f"{structure} 다만 {repair}",
        "house_view": outlook,
    }


def _build_section_captions(period_key: str, diagnostics: Dict[str, Any]) -> Dict[str, str]:
    raw_state = "직접 추출값이 존재하는 연도" if diagnostics["statement_available"] else "추정 baseline 연도"
    return {
        "composition": f"{period_key} 기준 정규화된 비목 구성입니다. 어떤 비용이 baseline을 주도하는지 한눈에 보여줍니다.",
        "comparison": f"{raw_state}에 대해 원시 추출, 최종 복원, 공식기준이 얼마나 다른지 비교합니다.",
        "raw_accounts": "원시 재무 계정값과 매출 대비 비율을 함께 제시해 OCR/매핑 오류 가능성을 점검합니다.",
        "trend": "같은 업종의 연도별 비목 구성비 변화를 통해 구조적 방향성을 읽기 위한 추이 차트입니다.",
        "official_reference": "config_official_reference 시트에 적재된 공식기준 표준안과 2025 추정치를 그대로 보여줍니다.",
        "adjustment": "비목별로 어떤 로직이 적용되었는지, 그리고 어떤 근거 계정이 사용되었는지 기록합니다.",
        "methods": "리포트 해석 시 꼭 알아야 할 normalization, 보정, 추정 규칙을 정리했습니다.",
        "sources": "국가기관 및 통계 포털 기준 참고 자료입니다. 현재 workbook은 이 자료들을 앵커로 삼아 표준안을 생성합니다.",
    }


def _series_to_polyline(points: List[tuple[float, float]], width: int, height: int, pad_left: int, pad_right: int, pad_top: int, pad_bottom: int) -> str:
    if not points:
        return ""
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(min(ys), 0.0), max(ys)
    if math.isclose(max_x, min_x):
        max_x += 1.0
    if math.isclose(max_y, min_y):
        max_y += 1.0

    def project_x(value: float) -> float:
        return pad_left + ((value - min_x) / (max_x - min_x)) * (width - pad_left - pad_right)

    def project_y(value: float) -> float:
        return height - pad_bottom - ((value - min_y) / (max_y - min_y)) * (height - pad_top - pad_bottom)

    return " ".join(f"{project_x(x):.1f},{project_y(y):.1f}" for x, y in points)


def _stacked_composition_svg(current_items: List[Dict[str, Any]]) -> str:
    width, height = 860, 94
    x = 0.0
    rects: List[str] = []
    labels: List[str] = []
    for row in current_items:
        item_code = row["item_code"]
        pct = max(_safe_float(row.get("base_pct")), 0.0)
        rect_width = width * pct / 100.0
        rects.append(f"<rect x='{x:.2f}' y='18' width='{rect_width:.2f}' height='32' fill='{COLOR_MAP.get(item_code, '#8892a0')}' />")
        if rect_width >= 62:
            labels.append(
                f"<text x='{x + rect_width / 2:.2f}' y='38' text-anchor='middle' fill='white' font-size='12' font-weight='700'>{escape(ITEM_LABELS.get(item_code, item_code))} {_format_pct(pct, 1)}</text>"
            )
        x += rect_width
    return (
        f"<svg viewBox='0 0 {width} {height}' class='chart-svg' role='img' aria-label='현재 비목 구성'>"
        f"<rect x='0' y='18' width='{width}' height='32' fill='#e9edf2' />"
        f"{''.join(rects)}{''.join(labels)}"
        "<text x='0' y='74' font-size='12' fill='#667085'>기준 비목 100 정규화 stacked composition</text></svg>"
    )


def _trend_svg(history_rows: List[Dict[str, Any]], periods: List[str]) -> str:
    width, height = 860, 260
    pad_left, pad_right, pad_top, pad_bottom = 50, 26, 18, 40
    usable_w = width - pad_left - pad_right
    usable_h = height - pad_top - pad_bottom
    all_values = [
        _safe_float(row["values"].get(period))
        for row in history_rows
        for period in periods
        if row["values"].get(period) is not None
    ]
    max_y = max(all_values) if all_values else 100.0
    max_y = max(max_y, 10.0)
    grid: List[str] = []
    for tick in range(5):
        y_value = max_y * tick / 4
        y = height - pad_bottom - usable_h * tick / 4
        grid.append(f"<line x1='{pad_left}' y1='{y:.1f}' x2='{width-pad_right}' y2='{y:.1f}' stroke='#e7eaef' stroke-width='1' />")
        grid.append(f"<text x='{pad_left-8}' y='{y+4:.1f}' text-anchor='end' font-size='11' fill='#69707d'>{y_value:.0f}%</text>")

    x_labels: List[str] = []
    for idx, period in enumerate(periods):
        x = pad_left + usable_w * idx / max(len(periods) - 1, 1)
        x_labels.append(f"<text x='{x:.1f}' y='{height-14}' text-anchor='middle' font-size='11' fill='#69707d'>{escape(period)}</text>")

    lines: List[str] = []
    legend: List[str] = []
    for row in history_rows:
        points = []
        for idx, period in enumerate(periods):
            value = row["values"].get(period)
            if value is None:
                continue
            points.append((float(idx), _safe_float(value)))
        if not points:
            continue
        polyline = _series_to_polyline(points, width, height, pad_left, pad_right, pad_top, pad_bottom)
        color = COLOR_MAP.get(row["item_code"], "#68707f")
        lines.append(f"<polyline fill='none' stroke='{color}' stroke-width='2.4' points='{polyline}' />")
        for idx, value in points:
            x = pad_left + usable_w * idx / max(len(periods) - 1, 1)
            y = height - pad_bottom - (value / max_y) * usable_h
            lines.append(f"<circle cx='{x:.1f}' cy='{y:.1f}' r='3.4' fill='{color}' />")
        legend.append(f"<div class='legend-item'><span class='legend-dot' style='background:{color}'></span>{escape(row['label'])}</div>")

    return (
        "<div class='chart-wrap'>"
        f"<svg viewBox='0 0 {width} {height}' class='chart-svg' role='img' aria-label='3개년 비목 추이'>{''.join(grid)}"
        f"<line x1='{pad_left}' y1='{height-pad_bottom}' x2='{width-pad_right}' y2='{height-pad_bottom}' stroke='#98a2b3' stroke-width='1.2' />"
        f"{''.join(lines)}{''.join(x_labels)}</svg>"
        f"<div class='legend-row'>{''.join(legend)}</div></div>"
    )


def _grouped_bar_svg(crosscheck_rows: List[Dict[str, Any]]) -> str:
    width, height = 860, 280
    pad_left, pad_right, pad_top, pad_bottom = 50, 30, 18, 56
    usable_w = width - pad_left - pad_right
    usable_h = height - pad_top - pad_bottom
    max_y = max(
        [_safe_float(row.get("raw_pct")) for row in crosscheck_rows]
        + [_safe_float(row.get("repaired_pct")) for row in crosscheck_rows]
        + [_safe_float(row.get("official_pct")) for row in crosscheck_rows]
        + [10.0]
    )
    group_w = usable_w / max(len(crosscheck_rows), 1)
    bar_w = min(24.0, group_w / 4.5)
    grid: List[str] = []
    for tick in range(5):
        y_value = max_y * tick / 4
        y = height - pad_bottom - usable_h * tick / 4
        grid.append(f"<line x1='{pad_left}' y1='{y:.1f}' x2='{width-pad_right}' y2='{y:.1f}' stroke='#e7eaef' stroke-width='1' />")
        grid.append(f"<text x='{pad_left-8}' y='{y+4:.1f}' text-anchor='end' font-size='11' fill='#69707d'>{y_value:.0f}%</text>")

    series_colors = {"raw": "#b4bcc9", "repaired": "#e4362b", "official": "#f39a1e"}
    bars: List[str] = []
    labels: List[str] = []
    for idx, row in enumerate(crosscheck_rows):
        base_x = pad_left + idx * group_w + group_w * 0.18
        for offset, (key, color) in enumerate(series_colors.items()):
            value = max(_safe_float(row.get(f"{key}_pct")), 0.0)
            x = base_x + offset * (bar_w + 6)
            bar_h = usable_h * value / max_y if max_y > 0 else 0.0
            y = height - pad_bottom - bar_h
            bars.append(f"<rect x='{x:.1f}' y='{y:.1f}' width='{bar_w:.1f}' height='{bar_h:.1f}' fill='{color}' rx='2' />")
        labels.append(f"<text x='{base_x + (bar_w + 6):.1f}' y='{height-20}' text-anchor='middle' font-size='11' fill='#69707d'>{escape(row['label'])}</text>")
    legend = "".join(
        f"<div class='legend-item'><span class='legend-dot' style='background:{color}'></span>{label}</div>"
        for label, color in [("원시 추출", "#b4bcc9"), ("최종 복원", "#e4362b"), ("공식 기준", "#f39a1e")]
    )
    return (
        "<div class='chart-wrap'>"
        f"<svg viewBox='0 0 {width} {height}' class='chart-svg' role='img' aria-label='원시 추출 대비 복원 및 공식 기준 비교'>{''.join(grid)}"
        f"<line x1='{pad_left}' y1='{height-pad_bottom}' x2='{width-pad_right}' y2='{height-pad_bottom}' stroke='#98a2b3' stroke-width='1.2' />"
        f"{''.join(bars)}{''.join(labels)}</svg>"
        f"<div class='legend-row'>{legend}</div></div>"
    )


def build_industry_research_payload(
    industry_group: str,
    period_key: str,
    workbook_path: Optional[str] = None,
    industry_group_name: Optional[str] = None,
) -> Dict[str, Any]:
    workbook = resolve_workbook_path(workbook_path)
    sheets = load_industry_db(workbook)
    current_items = get_baseline_items(industry_group, str(period_key), workbook_path=workbook)
    current_items = sorted(
        current_items,
        key=lambda row: ITEM_ORDER.index(row["item_code"]) if row["item_code"] in ITEM_ORDER else len(ITEM_ORDER),
    )
    if not current_items:
        raise ValueError(f"{industry_group} / {period_key} baseline 데이터가 없습니다.")

    baseline_df = _latest_baseline_rows(sheets["industry_baselines"], industry_group)
    metric_df = _latest_metric_rows(sheets["metric_observations"], industry_group, str(period_key))
    reference_df = _latest_reference_rows(sheets["config_official_reference"], industry_group)
    statement_rows = _latest_statement_rows(sheets["statement_rows"], industry_group, str(period_key))

    history_periods = sorted(baseline_df["period_key"].astype(str).unique().tolist(), key=_period_sort_key)
    history_rows = _item_history_rows(baseline_df, history_periods)

    reference_periods = sorted(
        [period for period in reference_df["period_key"].astype(str).tolist() if period.isdigit()],
        key=_period_sort_key,
    )
    current_reference_match = reference_df[reference_df["period_key"].astype(str) == str(period_key)]
    if current_reference_match.empty:
        current_reference_match = reference_df[reference_df["period_key"].astype(str) == "DEFAULT"]
    current_reference_row = current_reference_match.iloc[-1].to_dict() if not current_reference_match.empty else {}
    estimate_match = reference_df[reference_df["period_key"].astype(str) == "2025"]
    estimate_row = estimate_match.iloc[-1].to_dict() if not estimate_match.empty else None

    values = {
        str(row["normalized_account_code"]): _safe_float(row["amount"])
        for _, row in statement_rows.iterrows()
        if pd.notna(row.get("normalized_account_code"))
    }
    revenue = abs(values.get("is_revenue", 0.0))
    raw_map = {
        "material": _ratio(abs(values.get("cost_material", 0.0)), revenue),
        "labor": _ratio(abs(values.get("cost_labor", 0.0)), revenue),
        "expense": _ratio(abs(values.get("cost_expense", 0.0)), revenue),
        "gna": _ratio(abs(values.get("is_gna", 0.0)), revenue),
        "profit": _ratio(values.get("is_operating_profit", 0.0), revenue),
        "rnd": _ratio(abs(values.get("is_rnd_research", 0.0)) + abs(values.get("is_rnd_development", 0.0)), revenue),
    }
    raw_cogs_pct = _ratio(abs(values.get("is_cogs", 0.0)), revenue)
    cost_sum_pct = raw_map["material"] + raw_map["labor"] + raw_map["expense"]

    metric_map = {
        str(row["item_code"]): {
            "value_pct": _safe_float(row.get("value_pct")),
            "basis_type": str(row.get("basis_type", "")),
            "confidence_grade": str(row.get("confidence_grade", "")),
            "calc_rule_code": str(row.get("calc_rule_code", "")),
            "supporting_accounts": str(row.get("supporting_accounts", "")),
        }
        for _, row in metric_df.iterrows()
        if pd.notna(row.get("item_code"))
    }
    repaired_cogs_pct = sum(_safe_float(metric_map.get(item, {}).get("value_pct")) for item in ("material", "labor", "expense"))
    official_metric_map = _current_reference_map(current_reference_row)

    current_items_enriched: List[Dict[str, Any]] = []
    adjustment_events: List[Dict[str, Any]] = []
    crosscheck_rows: List[Dict[str, Any]] = []
    for row in current_items:
        item_code = row["item_code"]
        metric = metric_map.get(item_code, {})
        official_pct = official_metric_map.get(item_code, 0.0)
        enriched = {
            **row,
            "basis_type_label": _basis_label(row.get("basis_type")),
            "confidence_label": _confidence_label(row.get("confidence_grade")),
            "metric_value_pct": _safe_float(metric.get("value_pct")),
            "calc_rule_code": metric.get("calc_rule_code", row.get("calc_rule_code", "")),
            "calc_rule_label": _calc_rule_label(metric.get("calc_rule_code", row.get("calc_rule_code", ""))),
            "supporting_accounts": metric.get("supporting_accounts", row.get("supporting_accounts", "")),
            "official_pct": official_pct,
        }
        current_items_enriched.append(enriched)
        adjustment_events.append(
            {
                "label": enriched["label"],
                "basis_type_label": enriched["basis_type_label"],
                "confidence_label": enriched["confidence_label"],
                "calc_rule_label": enriched["calc_rule_label"],
                "supporting_accounts": enriched["supporting_accounts"],
                "metric_value_pct": enriched["metric_value_pct"],
            }
        )
        crosscheck_rows.append(
            {
                "item_code": item_code,
                "label": enriched["label"],
                "raw_pct": raw_map.get(item_code) if not statement_rows.empty else None,
                "repaired_pct": enriched["metric_value_pct"],
                "official_pct": official_pct,
                "base_pct": _safe_float(enriched.get("base_pct")),
                "note": enriched["calc_rule_label"],
            }
        )

    diagnostics = {
        "statement_available": not statement_rows.empty,
        "revenue": revenue,
        "raw_cogs_pct": raw_cogs_pct,
        "cost_sum_pct": cost_sum_pct,
        "raw_vs_repaired_gap": abs(raw_cogs_pct - repaired_cogs_pct),
        "cost_gap_pct": abs(cost_sum_pct - raw_cogs_pct),
        "repaired_cogs_pct": repaired_cogs_pct,
        "current_reference_source": str(current_reference_row.get("source", "")),
    }

    raw_account_rows = []
    for code, label in [
        ("is_revenue", "매출액"),
        ("is_cogs", "매출원가"),
        ("is_gross_profit", "매출총이익"),
        ("is_gna", "판관비"),
        ("is_operating_profit", "영업이익"),
        ("cost_material", "재료비"),
        ("cost_labor", "노무비"),
        ("cost_expense", "경비"),
        ("is_rnd_research", "연구비"),
        ("is_rnd_development", "경상개발비"),
    ]:
        amount = values.get(code)
        if amount or code in {"is_revenue", "is_cogs", "is_gna", "is_operating_profit"}:
            raw_account_rows.append(
                {
                    "label": label,
                    "amount": amount,
                    "ratio_pct": _ratio(amount, revenue) if revenue > 0 and amount is not None else None,
                }
            )

    official_reference_rows: List[Dict[str, Any]] = []
    for period in reference_periods:
        row = reference_df[reference_df["period_key"].astype(str) == str(period)]
        if row.empty:
            continue
        selected = row.iloc[-1]
        official_reference_rows.append(
            {
                "period_key": str(period),
                "cogs_pct": _safe_float(selected.get("cogs_pct")),
                "gna_pct": _safe_float(selected.get("gna_pct")),
                "profit_pct": _safe_float(selected.get("profit_pct")),
                "rnd_pct": _safe_float(selected.get("rnd_pct")),
                "material_share": _safe_float(selected.get("material_share")),
                "labor_share": _safe_float(selected.get("labor_share")),
                "expense_share": _safe_float(selected.get("expense_share")),
                "source": str(selected.get("source", "")),
                "note": str(selected.get("note", "")),
            }
        )

    rating, rating_note, quality_badge_class = _build_rating(current_items_enriched, metric_df, statement_rows, current_reference_row)
    summary_points = _build_summary_points(current_items_enriched, diagnostics, current_reference_row, estimate_row, adjustment_events)
    method_notes = _build_method_notes(str(period_key), statement_rows, diagnostics, current_reference_row, estimate_row)
    house_view = _build_house_view(profile_label=industry_group_name or INDUSTRY_GROUP_LABELS.get(industry_group, industry_group), current_items=current_items_enriched, diagnostics=diagnostics, current_reference_row=current_reference_row, estimate_row=estimate_row)
    section_captions = _build_section_captions(str(period_key), diagnostics)

    profile_label = industry_group_name or INDUSTRY_GROUP_LABELS.get(industry_group, industry_group)
    subtitle = f"{profile_label} ({industry_group}) | Baseline {period_key}"
    if estimate_row:
        subtitle += " | 2025 Estimate Included"

    return {
        "title": f"{profile_label} ({industry_group})",
        "subtitle": subtitle,
        "profile_label": profile_label,
        "industry_group": str(industry_group),
        "period_key": str(period_key),
        "generated_at": datetime.now().strftime("%Y-%m-%d"),
        "rating": rating,
        "rating_note": rating_note,
        "quality_badge_class": quality_badge_class,
        "summary_points": summary_points,
        "method_notes": method_notes,
        "house_lede": house_view["lede"],
        "house_view": house_view["house_view"],
        "section_captions": section_captions,
        "source_note": "국가통계 공식자료와 workbook 복원값을 함께 사용한 내부 산업 baseline 연구 리포트입니다.",
        "current_items": current_items_enriched,
        "history_periods": history_periods,
        "history_rows": history_rows,
        "official_reference_rows": official_reference_rows,
        "current_reference_row": current_reference_row,
        "estimate_row": estimate_row,
        "diagnostics": diagnostics,
        "crosscheck_rows": crosscheck_rows,
        "raw_account_rows": raw_account_rows,
        "adjustment_events": adjustment_events,
        "judgement_guide": JUDGEMENT_GUIDE,
        "basis_guide": BASIS_GUIDE,
        "reference_sources": REFERENCE_SOURCES,
        "stacked_composition_svg": _stacked_composition_svg(current_items_enriched),
        "trend_svg": _trend_svg(history_rows, history_periods),
        "crosscheck_svg": _grouped_bar_svg(crosscheck_rows),
    }


def build_industry_research_markdown(payload: Dict[str, Any]) -> str:
    lines = [
        f"# {payload['title']}",
        payload["subtitle"],
        "",
        f"- 기준 기간: {payload['period_key']}",
        f"- 활용 판단: {payload['rating']} ({payload['rating_note']})",
        f"- House View: {payload['house_view']}",
        "",
        "## 핵심 요약",
    ]
    lines.extend([f"- {point}" for point in payload["summary_points"]])
    lines.append("")
    lines.append("## 방법론")
    lines.extend([f"- {note}" for note in payload["method_notes"]])
    lines.append("")
    lines.append("## 참고 출처")
    lines.extend([f"- {row['org']} | {row['title']} | {row['url']}" for row in payload["reference_sources"]])
    return "\n".join(lines)


def generate_industry_research_html(payload: Dict[str, Any]) -> str:
    current_rows_html = "".join(
        f"<tr><td>{escape(row['label'])}</td><td>{escape(_format_pct(row['base_pct']))}</td><td>{escape(_format_pct(row['metric_value_pct']))}</td><td>{escape(_format_pct(row['official_pct']))}</td><td>{escape(row['basis_type_label'])}</td><td>{escape(row['confidence_label'])}</td></tr>"
        for row in payload["current_items"]
    )
    history_header = "".join(f"<th>{escape(period)}</th>" for period in payload["history_periods"])
    history_rows_html = ""
    for row in payload["history_rows"]:
        value_cells = "".join(
            f"<td>{escape(_format_pct(row['values'][period])) if row['values'].get(period) is not None else '-'}</td>"
            for period in payload["history_periods"]
        )
        history_rows_html += f"<tr><td>{escape(row['label'])}</td>{value_cells}<td>{escape(row['basis_type'])}</td><td>{escape(row['confidence_grade'])}</td></tr>"

    official_rows_html = "".join(
        f"<tr><td>{escape(row['period_key'])}</td><td>{escape(_format_pct(row['cogs_pct']))}</td><td>{escape(_format_pct(row['gna_pct']))}</td><td>{escape(_format_pct(row['profit_pct']))}</td><td>{escape(_format_pct(row['rnd_pct']))}</td><td>{escape(_format_pct(row['material_share']))} / {escape(_format_pct(row['labor_share']))} / {escape(_format_pct(row['expense_share']))}</td><td>{escape(row['source'])}</td></tr>"
        for row in payload["official_reference_rows"]
    )
    adjustment_rows_html = "".join(
        f"<tr><td>{escape(row['label'])}</td><td>{escape(_format_pct(row['metric_value_pct']))}</td><td>{escape(row['basis_type_label'])}</td><td>{escape(row['confidence_label'])}</td><td>{escape(row['calc_rule_label'])}</td><td>{escape(row['supporting_accounts'] or '-')}</td></tr>"
        for row in payload["adjustment_events"]
    )
    raw_rows_html = "".join(
        f"<tr><td>{escape(row['label'])}</td><td>{escape(_format_amount(row['amount']))}</td><td>{escape(_format_pct(row['ratio_pct'])) if row['ratio_pct'] is not None else '-'}</td></tr>"
        for row in payload["raw_account_rows"]
    )
    crosscheck_rows_html = "".join(
        f"<tr><td>{escape(row['label'])}</td><td>{escape(_format_pct(row['raw_pct'])) if row['raw_pct'] is not None else '-'}</td><td>{escape(_format_pct(row['repaired_pct']))}</td><td>{escape(_format_pct(row['official_pct']))}</td><td>{escape(_format_pct(row['base_pct']))}</td><td>{escape(row['note'])}</td></tr>"
        for row in payload["crosscheck_rows"]
    )
    summary_html = "".join(f"<li>{escape(point)}</li>" for point in payload["summary_points"])
    notes_html = "".join(f"<li>{escape(note)}</li>" for note in payload["method_notes"])
    basis_guide_html = "".join(
        f"<li><strong>{escape(row['label'])}</strong><br>{escape(row['description'])}</li>"
        for row in payload["basis_guide"]
    )
    judgement_guide_html = "".join(
        f"<li><strong>{escape(row['label'])}</strong><br>{escape(row['description'])}</li>"
        for row in payload["judgement_guide"]
    )
    source_rows_html = "".join(
        f"<li><strong>{escape(row['org'])}</strong> | <a href='{escape(row['url'])}' target='_blank'>{escape(row['title'])}</a><br>{escape(row['note'])}</li>"
        for row in payload["reference_sources"]
    )

    current_ref = payload["current_reference_row"]
    diagnostics = payload["diagnostics"]
    captions = payload["section_captions"]
    estimate_box = ""
    if payload["estimate_row"]:
        estimate = payload["estimate_row"]
        estimate_box = (
            "<div class='info-card'><div class='card-title'>2025 추정치</div>"
            f"<div class='big-metric'>{escape(_format_pct(estimate.get('cogs_pct')))}</div>"
            "<div class='small-note'>매출원가율 추정</div>"
            f"<div class='metric-mini-grid'><div>판관비 {escape(_format_pct(estimate.get('gna_pct')))}</div>"
            f"<div>영업이익 {escape(_format_pct(estimate.get('profit_pct')))}</div>"
            f"<div>R&D {escape(_format_pct(estimate.get('rnd_pct')))}</div></div></div>"
        )

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{escape(payload['title'])} Research Report</title>
  <style>
    :root {{
      --page:#f4f5f7; --paper:#fff; --ink:#20242c; --muted:#667085; --line:#d9dde4; --soft:#f7f8fb;
      --accent-main:#e5392d; --accent-sub:#f59a1e; --good:#0f7a5d; --warn:#9a5a18; --blue:#2757c6;
    }}
    * {{ box-sizing:border-box; }}
    @page {{ size:A4 portrait; margin:0; }}
    body {{ margin:0; padding:8mm 0; background:var(--page); color:var(--ink); font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif; }}
    .page {{ width:210mm; max-width:210mm; min-height:297mm; margin:0 auto; background:var(--paper); padding:10mm 11mm 12mm; box-shadow:0 12px 28px rgba(15,23,42,.08); overflow:hidden; }}
    .meta-line {{ display:flex; justify-content:space-between; font-size:12px; color:var(--muted); margin-bottom:10px; }}
    .hero {{ display:flex; justify-content:space-between; align-items:flex-end; gap:20px; }}
    h1 {{ margin:0; font-size:36px; line-height:1.02; letter-spacing:-.06em; font-weight:900; }}
    .subtitle {{ margin-top:8px; font-size:18px; font-weight:800; letter-spacing:-.03em; color:#343946; max-width:118mm; overflow-wrap:anywhere; }}
    .desk {{ text-align:right; font-size:12px; font-weight:800; color:var(--accent-main); line-height:1.35; min-width:34mm; max-width:34mm; overflow-wrap:anywhere; }}
    .desk strong {{ display:block; font-size:15px; letter-spacing:-.03em; }}
    .ribbon {{ display:flex; height:12px; margin:14px 0 14px; border-top:1px solid #efefef; }}
    .ribbon .main {{ flex:1.08; background:var(--accent-main); }}
    .ribbon .sub {{ flex:.38; background:var(--accent-sub); }}
    .layout {{ display:grid; grid-template-columns:44mm minmax(0,1fr); gap:7mm; align-items:start; }}
    .layout > * {{ min-width:0; }}
    .sidebar {{ border-right:1px solid var(--line); padding-right:5mm; min-width:0; overflow:hidden; }}
    .content {{ min-width:0; }}
    .side-card,.info-card {{ background:var(--soft); padding:9px 10px; margin-bottom:10px; border:1px solid #e8ebf0; min-width:0; overflow:hidden; }}
    .side-card h3,.info-card .card-title {{ margin:0 0 7px; font-size:12px; letter-spacing:-.02em; }}
    .rating .label {{ font-size:18px; font-weight:900; letter-spacing:-.04em; }}
    .rating .note {{ margin-top:4px; font-size:10.5px; color:var(--muted); line-height:1.5; }}
    .badge {{ display:inline-block; margin-top:8px; padding:3px 8px; border-radius:999px; font-size:10px; font-weight:800; }}
    .badge.good {{ background:rgba(15,122,93,.14); color:var(--good); }}
    .badge.warn {{ background:rgba(154,90,24,.16); color:var(--warn); }}
    .mini-table {{ width:100%; border-collapse:collapse; table-layout:fixed; font-size:10.5px; }}
    .mini-table td {{ border-top:1px solid var(--line); padding:5px 0; vertical-align:top; overflow-wrap:anywhere; word-break:break-word; }}
    .mini-table td:first-child {{ width:36%; padding-right:6px; color:var(--muted); font-weight:700; }}
    .mini-table td:last-child {{ text-align:left; font-weight:800; padding-left:6px; }}
    .content section {{ margin-bottom:18px; }}
    h2 {{ margin:0 0 8px; font-size:17px; letter-spacing:-.04em; }}
    .lede {{ font-size:12.5px; line-height:1.62; font-weight:800; color:#2d3541; margin-bottom:8px; }}
    .house-view {{ margin:0 0 10px; padding:8px 10px; background:#fbf4f3; border-left:3px solid var(--accent-main); line-height:1.6; color:#2b3340; font-size:11px; }}
    .caption {{ margin:5px 0 8px; font-size:10.5px; color:var(--muted); line-height:1.5; }}
    .summary-list,.notes-list,.guide-list,.source-list {{ margin:0; padding-left:16px; line-height:1.58; color:#2f3642; font-size:11px; }}
    .summary-list li,.notes-list li,.guide-list li,.source-list li {{ margin-bottom:4px; }}
    .two-col {{ display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:12px; }}
    .panel {{ border-top:2px solid var(--accent-main); padding-top:8px; min-width:0; overflow:hidden; }}
    .chart-wrap {{ border:1px solid #e8ebf0; background:linear-gradient(180deg,#fff,#fafbfc); padding:8px 8px 6px; min-width:0; overflow:hidden; }}
    .chart-svg {{ width:100%; height:auto; display:block; max-width:100%; }}
    .legend-row {{ display:flex; flex-wrap:wrap; gap:12px 18px; padding:6px 2px 0; font-size:12px; color:#616977; }}
    .legend-item {{ display:inline-flex; align-items:center; gap:7px; }}
    .legend-dot {{ width:10px; height:10px; border-radius:50%; display:inline-block; }}
    table {{ width:100%; border-collapse:collapse; table-layout:fixed; }}
    th,td {{ border-bottom:1px solid var(--line); padding:5px 5px; vertical-align:top; text-align:left; font-size:10.5px; overflow-wrap:anywhere; word-break:break-word; }}
    th {{ background:#f7f8fa; font-size:10px; color:#596171; text-transform:uppercase; letter-spacing:.01em; }}
    .big-metric {{ font-size:22px; font-weight:900; letter-spacing:-.05em; color:var(--blue); }}
    .small-note {{ margin-top:4px; font-size:10px; color:var(--muted); }}
    .metric-mini-grid {{ display:grid; gap:3px; color:#4e5664; font-size:10px; }}
    a {{ color:var(--blue); text-decoration:none; }}
    .footnote {{ margin-top:16px; padding-top:10px; border-top:1px solid var(--line); font-size:10px; color:var(--muted); line-height:1.55; }}
    @media print {{
      body {{ padding:0; background:#fff; }}
      .page {{ margin:0; box-shadow:none; width:210mm; max-width:210mm; min-height:297mm; overflow:hidden; }}
    }}
    @media (max-width:980px) {{
      .page {{ width:auto; max-width:none; min-height:auto; margin:0; padding:16px; overflow:visible; }} .layout,.two-col {{ grid-template-columns:1fr; }} .sidebar {{ border-right:none; padding-right:0; overflow:visible; }}
      .hero {{ flex-direction:column; align-items:flex-start; gap:10px; }} h1 {{ font-size:32px; }} .subtitle {{ font-size:20px; max-width:none; }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <div class="meta-line"><div>Industry Baseline Research Report</div><div>{escape(payload['generated_at'])}</div></div>
    <div class="hero"><div><h1>{escape(payload['title'])}</h1><div class="subtitle">{escape(payload['subtitle'])}</div></div><div class="desk"><strong>Industry Research</strong>Internal Review Edition</div></div>
    <div class="ribbon"><div class="main"></div><div class="sub"></div></div>
    <div class="layout">
      <aside class="sidebar">
        <div class="side-card rating"><div class="label">{escape(payload['rating'])}</div><div class="note">{escape(payload['rating_note'])}</div><div class="badge {escape(payload['quality_badge_class'])}">{escape(payload['quality_badge_class'] == 'good' and '활용 양호' or '보수 활용')}</div></div>
        <div class="side-card"><h3>Baseline Snapshot</h3><table class="mini-table">
          <tr><td>기준 기간</td><td>{escape(payload['period_key'])}</td></tr>
          <tr><td>산업군</td><td>{escape(payload['profile_label'])}</td></tr>
          <tr><td>공식기준 Source</td><td>{escape(str(current_ref.get('source', '-')))}</td></tr>
          <tr><td>원시 추출 여부</td><td>{'있음' if diagnostics['statement_available'] else '없음'}</td></tr>
          <tr><td>원시 매출원가율</td><td>{escape(_format_pct(diagnostics['raw_cogs_pct']))}</td></tr>
          <tr><td>복원 매출원가율</td><td>{escape(_format_pct(diagnostics['repaired_cogs_pct']))}</td></tr>
        </table></div>
        <div class="side-card"><h3>현재 연도 공식 기준</h3><table class="mini-table">
          <tr><td>매출원가</td><td>{escape(_format_pct(current_ref.get('cogs_pct')))}</td></tr>
          <tr><td>판관비</td><td>{escape(_format_pct(current_ref.get('gna_pct')))}</td></tr>
          <tr><td>영업이익</td><td>{escape(_format_pct(current_ref.get('profit_pct')))}</td></tr>
          <tr><td>R&D</td><td>{escape(_format_pct(current_ref.get('rnd_pct')))}</td></tr>
        </table></div>
        {estimate_box}
        <div class="side-card"><h3>기준 유형 해설</h3><ul class="guide-list">{basis_guide_html}</ul></div>
        <div class="side-card"><h3>판정 등급 해설</h3><ul class="guide-list">{judgement_guide_html}</ul></div>
      </aside>
      <main class="content">
        <section><div class="lede">{escape(payload['house_lede'])}</div><div class="house-view">{escape(payload['house_view'])}</div><ul class="summary-list">{summary_html}</ul></section>
        <section><h2>1. 현재 비목 구조</h2><div class="caption">{escape(captions['composition'])}</div><div class="panel">{payload['stacked_composition_svg']}</div><div class="panel" style="margin-top:16px;"><table><thead><tr><th>비목</th><th>구성비 Base</th><th>매출 대비 복원값</th><th>공식 기준</th><th>기준 유형</th><th>신뢰도</th></tr></thead><tbody>{current_rows_html}</tbody></table></div></section>
        <section class="two-col"><div class="panel"><h2>2. 원시 추출 vs 복원 vs 공식 기준</h2><div class="caption">{escape(captions['comparison'])}</div>{payload['crosscheck_svg']}<table style="margin-top:12px;"><thead><tr><th>비목</th><th>원시 추출</th><th>최종 복원</th><th>공식 기준</th><th>구성비 Base</th><th>적용 로직</th></tr></thead><tbody>{crosscheck_rows_html}</tbody></table></div><div class="panel"><h2>3. 원시 재무 계정 진단</h2><div class="caption">{escape(captions['raw_accounts'])}</div><table><thead><tr><th>계정</th><th>금액</th><th>매출 대비</th></tr></thead><tbody>{raw_rows_html}</tbody></table></div></section>
        <section><h2>4. 3개년 비목 추이</h2><div class="caption">{escape(captions['trend'])}</div>{payload['trend_svg']}<table style="margin-top:12px;"><thead><tr><th>비목</th>{history_header}<th>기준 유형</th><th>신뢰도</th></tr></thead><tbody>{history_rows_html}</tbody></table></section>
        <section class="two-col"><div class="panel"><h2>5. 공식 기준 시트 이력</h2><div class="caption">{escape(captions['official_reference'])}</div><table><thead><tr><th>연도</th><th>매출원가</th><th>판관비</th><th>영업이익</th><th>R&D</th><th>직접비 믹스(M/L/E)</th><th>source</th></tr></thead><tbody>{official_rows_html}</tbody></table></div><div class="panel"><h2>6. 보정 로그</h2><div class="caption">{escape(captions['adjustment'])}</div><table><thead><tr><th>비목</th><th>매출 대비</th><th>기준 유형</th><th>신뢰도</th><th>적용 로직</th><th>근거 계정</th></tr></thead><tbody>{adjustment_rows_html}</tbody></table></div></section>
        <section class="two-col"><div><h2>7. 방법론 및 해석</h2><div class="caption">{escape(captions['methods'])}</div><ul class="notes-list">{notes_html}</ul></div><div><h2>8. 국가통계 및 참고 출처</h2><div class="caption">{escape(captions['sources'])}</div><ul class="source-list">{source_rows_html}</ul></div></section>
      </main>
    </div>
    <div class="footnote">{escape(payload['source_note'])} 현재 workbook의 연도별 공식 기준 행은 국가통계 exact raw import가 아니라, 국가통계 앵커와 workbook 복원 이력을 혼합한 standard row입니다.</div>
  </div>
</body>
</html>
"""


def _legacy_generate_industry_research_html_old(payload: Dict[str, Any]) -> str:
    current_rows_html = "".join(
        f"<tr><td>{escape(row['label'])}</td><td>{escape(_format_pct(row['base_pct']))}</td><td>{escape(_format_pct(row['metric_value_pct']))}</td><td>{escape(_format_pct(row['official_pct']))}</td><td>{escape(row['basis_type_label'])}</td><td>{escape(row['confidence_label'])}</td></tr>"
        for row in payload["current_items"]
    )
    history_header = "".join(f"<th>{escape(period)}</th>" for period in payload["history_periods"])
    history_rows_html = ""
    for row in payload["history_rows"]:
        value_cells = "".join(
            f"<td>{escape(_format_pct(row['values'][period])) if row['values'].get(period) is not None else '-'}</td>"
            for period in payload["history_periods"]
        )
        history_rows_html += f"<tr><td>{escape(row['label'])}</td>{value_cells}<td>{escape(row['basis_type'])}</td><td>{escape(row['confidence_grade'])}</td></tr>"

    official_rows_html = "".join(
        f"<tr><td>{escape(row['period_key'])}</td><td>{escape(_format_pct(row['cogs_pct']))}</td><td>{escape(_format_pct(row['gna_pct']))}</td><td>{escape(_format_pct(row['profit_pct']))}</td><td>{escape(_format_pct(row['rnd_pct']))}</td><td>{escape(_format_pct(row['material_share']))} / {escape(_format_pct(row['labor_share']))} / {escape(_format_pct(row['expense_share']))}</td><td>{escape(row['source'])}</td></tr>"
        for row in payload["official_reference_rows"]
    )
    adjustment_rows_html = "".join(
        f"<tr><td>{escape(row['label'])}</td><td>{escape(_format_pct(row['metric_value_pct']))}</td><td>{escape(row['basis_type_label'])}</td><td>{escape(row['confidence_label'])}</td><td>{escape(row['calc_rule_label'])}</td><td>{escape(row['supporting_accounts'] or '-')}</td></tr>"
        for row in payload["adjustment_events"]
    )
    raw_rows_html = "".join(
        f"<tr><td>{escape(row['label'])}</td><td>{escape(_format_amount(row['amount']))}</td><td>{escape(_format_pct(row['ratio_pct'])) if row['ratio_pct'] is not None else '-'}</td></tr>"
        for row in payload["raw_account_rows"]
    )
    crosscheck_rows_html = "".join(
        f"<tr><td>{escape(row['label'])}</td><td>{escape(_format_pct(row['raw_pct'])) if row['raw_pct'] is not None else '-'}</td><td>{escape(_format_pct(row['repaired_pct']))}</td><td>{escape(_format_pct(row['official_pct']))}</td><td>{escape(_format_pct(row['base_pct']))}</td><td>{escape(row['note'])}</td></tr>"
        for row in payload["crosscheck_rows"]
    )
    summary_html = "".join(f"<li>{escape(point)}</li>" for point in payload["summary_points"])
    notes_html = "".join(f"<li>{escape(note)}</li>" for note in payload["method_notes"])
    basis_guide_html = "".join(
        f"<li><strong>{escape(row['label'])}</strong><br>{escape(row['description'])}</li>"
        for row in payload["basis_guide"]
    )
    judgement_guide_html = "".join(
        f"<li><strong>{escape(row['label'])}</strong><br>{escape(row['description'])}</li>"
        for row in payload["judgement_guide"]
    )
    source_rows_html = "".join(
        f"<li><strong>{escape(row['org'])}</strong> | <a href='{escape(row['url'])}' target='_blank'>{escape(row['title'])}</a><br>{escape(row['note'])}</li>"
        for row in payload["reference_sources"]
    )

    current_ref = payload["current_reference_row"]
    diagnostics = payload["diagnostics"]
    captions = payload["section_captions"]
    estimate_box = ""
    if payload["estimate_row"]:
        estimate = payload["estimate_row"]
        estimate_box = (
            "<div class='info-card'><div class='card-title'>2025 추정치</div>"
            f"<div class='big-metric'>{escape(_format_pct(estimate.get('cogs_pct')))}</div>"
            "<div class='small-note'>매출원가율 추정</div>"
            f"<div class='metric-mini-grid'><div>판관비 {escape(_format_pct(estimate.get('gna_pct')))}</div>"
            f"<div>영업이익 {escape(_format_pct(estimate.get('profit_pct')))}</div>"
            f"<div>R&D {escape(_format_pct(estimate.get('rnd_pct')))}</div></div></div>"
        )

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{escape(payload['title'])} Research Report</title>
  <style>
    :root {{
      --page:#f4f5f7; --paper:#fff; --ink:#20242c; --muted:#667085; --line:#d9dde4; --soft:#f7f8fb;
      --accent-main:#e5392d; --accent-sub:#f59a1e; --good:#0f7a5d; --warn:#9a5a18; --blue:#2757c6;
    }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; background:var(--page); color:var(--ink); font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif; }}
    .page {{ max-width:1180px; margin:18px auto; background:var(--paper); padding:24px 28px 30px; box-shadow:0 18px 42px rgba(15,23,42,.08); }}
    .meta-line {{ display:flex; justify-content:space-between; font-size:12px; color:var(--muted); margin-bottom:10px; }}
    .hero {{ display:flex; justify-content:space-between; align-items:flex-end; gap:20px; }}
    h1 {{ margin:0; font-size:52px; line-height:.95; letter-spacing:-.06em; font-weight:900; }}
    .subtitle {{ margin-top:10px; font-size:30px; font-weight:800; letter-spacing:-.04em; color:#343946; }}
    .desk {{ text-align:right; font-size:18px; font-weight:800; color:var(--accent-main); line-height:1.3; }}
    .desk strong {{ display:block; font-size:24px; letter-spacing:-.04em; }}
    .ribbon {{ display:flex; height:30px; margin:22px 0 20px; border-top:1px solid #efefef; }}
    .ribbon .main {{ flex:1.08; background:var(--accent-main); }}
    .ribbon .sub {{ flex:.38; background:var(--accent-sub); }}
    .layout {{ display:grid; grid-template-columns:300px 1fr; gap:26px; }}
    .sidebar {{ border-right:1px solid var(--line); padding-right:20px; }}
    .side-card,.info-card {{ background:var(--soft); padding:14px 16px; margin-bottom:14px; border:1px solid #e8ebf0; }}
    .side-card h3,.info-card .card-title {{ margin:0 0 10px; font-size:16px; letter-spacing:-.03em; }}
    .rating .label {{ font-size:30px; font-weight:900; letter-spacing:-.05em; }}
    .rating .note {{ margin-top:6px; font-size:13px; color:var(--muted); line-height:1.65; }}
    .badge {{ display:inline-block; margin-top:10px; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:800; }}
    .badge.good {{ background:rgba(15,122,93,.14); color:var(--good); }}
    .badge.warn {{ background:rgba(154,90,24,.16); color:var(--warn); }}
    .mini-table {{ width:100%; border-collapse:collapse; font-size:12px; }}
    .mini-table td {{ border-top:1px solid var(--line); padding:7px 0; vertical-align:top; }}
    .mini-table td:last-child {{ text-align:right; font-weight:800; }}
    .content section {{ margin-bottom:28px; }}
    h2 {{ margin:0 0 12px; font-size:28px; letter-spacing:-.05em; }}
    .lede {{ font-size:17px; line-height:1.74; font-weight:800; color:#2d3541; margin-bottom:12px; }}
    .house-view {{ margin:0 0 14px; padding:14px 16px; background:#fbf4f3; border-left:4px solid var(--accent-main); line-height:1.75; color:#2b3340; }}
    .caption {{ margin:8px 0 12px; font-size:12px; color:var(--muted); line-height:1.65; }}
    .summary-list,.notes-list,.guide-list,.source-list {{ margin:0; padding-left:18px; line-height:1.72; color:#2f3642; }}
    .summary-list li,.notes-list li,.guide-list li,.source-list li {{ margin-bottom:6px; }}
    .two-col {{ display:grid; grid-template-columns:1fr 1fr; gap:22px; }}
    .panel {{ border-top:3px solid var(--accent-main); padding-top:12px; }}
    .chart-wrap {{ border:1px solid #e8ebf0; background:linear-gradient(180deg,#fff,#fafbfc); padding:12px 12px 8px; }}
    .chart-svg {{ width:100%; height:auto; display:block; }}
    .legend-row {{ display:flex; flex-wrap:wrap; gap:12px 18px; padding:6px 2px 0; font-size:12px; color:#616977; }}
    .legend-item {{ display:inline-flex; align-items:center; gap:7px; }}
    .legend-dot {{ width:10px; height:10px; border-radius:50%; display:inline-block; }}
    table {{ width:100%; border-collapse:collapse; }}
    th,td {{ border-bottom:1px solid var(--line); padding:9px 8px; vertical-align:top; text-align:left; font-size:13px; }}
    th {{ background:#f7f8fa; font-size:12px; color:#596171; text-transform:uppercase; letter-spacing:.02em; }}
    .big-metric {{ font-size:30px; font-weight:900; letter-spacing:-.05em; color:var(--blue); }}
    .small-note {{ margin-top:6px; font-size:12px; color:var(--muted); }}
    .metric-mini-grid {{ display:grid; gap:4px; color:#4e5664; font-size:12px; }}
    a {{ color:var(--blue); text-decoration:none; }}
    .footnote {{ margin-top:26px; padding-top:14px; border-top:1px solid var(--line); font-size:12px; color:var(--muted); line-height:1.7; }}
    @media (max-width:980px) {{
      .page {{ margin:0; padding:16px; }} .layout,.two-col {{ grid-template-columns:1fr; }} .sidebar {{ border-right:none; padding-right:0; }}
      .hero {{ flex-direction:column; align-items:flex-start; gap:10px; }} h1 {{ font-size:40px; }} .subtitle {{ font-size:24px; }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <div class="meta-line"><div>Industry Baseline Research Report</div><div>{escape(payload['generated_at'])}</div></div>
    <div class="hero"><div><h1>{escape(payload['title'])}</h1><div class="subtitle">{escape(payload['subtitle'])}</div></div><div class="desk"><strong>Industry Research</strong>Internal Review Edition</div></div>
    <div class="ribbon"><div class="main"></div><div class="sub"></div></div>
    <div class="layout">
      <aside class="sidebar">
        <div class="side-card rating"><div class="label">{escape(payload['rating'])}</div><div class="note">{escape(payload['rating_note'])}</div><div class="badge {escape(payload['quality_badge_class'])}">{escape(payload['quality_badge_class'] == 'good' and '활용 양호' or '보수 활용')}</div></div>
        <div class="side-card"><h3>Baseline Snapshot</h3><table class="mini-table">
          <tr><td>기준 기간</td><td>{escape(payload['period_key'])}</td></tr>
          <tr><td>산업군</td><td>{escape(payload['profile_label'])}</td></tr>
          <tr><td>공식기준 Source</td><td>{escape(str(current_ref.get('source', '-')))}</td></tr>
          <tr><td>원시 추출 여부</td><td>{'있음' if diagnostics['statement_available'] else '없음'}</td></tr>
          <tr><td>원시 매출원가율</td><td>{escape(_format_pct(diagnostics['raw_cogs_pct']))}</td></tr>
          <tr><td>복원 매출원가율</td><td>{escape(_format_pct(diagnostics['repaired_cogs_pct']))}</td></tr>
        </table></div>
        <div class="side-card"><h3>현재 연도 공식 기준</h3><table class="mini-table">
          <tr><td>매출원가</td><td>{escape(_format_pct(current_ref.get('cogs_pct')))}</td></tr>
          <tr><td>판관비</td><td>{escape(_format_pct(current_ref.get('gna_pct')))}</td></tr>
          <tr><td>영업이익</td><td>{escape(_format_pct(current_ref.get('profit_pct')))}</td></tr>
          <tr><td>R&D</td><td>{escape(_format_pct(current_ref.get('rnd_pct')))}</td></tr>
        </table></div>
        {estimate_box}
        <div class="side-card"><h3>기준 유형 해설</h3><ul class="guide-list">{basis_guide_html}</ul></div>
        <div class="side-card"><h3>판정 등급 해설</h3><ul class="guide-list">{judgement_guide_html}</ul></div>
      </aside>
      <main class="content">
        <section><div class="lede">본 리포트는 재무 4표 OCR, 복원 로직, 공식기준 시트, 2025 추정치까지 포함해 산업별 baseline의 현재 상태와 활용상 주의점을 한 번에 검토하기 위한 내부 리서치 문서입니다.</div><ul class="summary-list">{summary_html}</ul></section>
        <section><h2>1. 현재 비목 구조</h2><div class="panel">{payload['stacked_composition_svg']}</div><div class="panel" style="margin-top:16px;"><table><thead><tr><th>비목</th><th>구성비 Base</th><th>매출 대비 복원값</th><th>공식 기준</th><th>기준 유형</th><th>신뢰도</th></tr></thead><tbody>{current_rows_html}</tbody></table></div></section>
        <section class="two-col"><div class="panel"><h2>2. 원시 추출 vs 복원 vs 공식 기준</h2>{payload['crosscheck_svg']}<table style="margin-top:12px;"><thead><tr><th>비목</th><th>원시 추출</th><th>최종 복원</th><th>공식 기준</th><th>구성비 Base</th><th>적용 로직</th></tr></thead><tbody>{crosscheck_rows_html}</tbody></table></div><div class="panel"><h2>3. 원시 재무 계정 진단</h2><table><thead><tr><th>계정</th><th>금액</th><th>매출 대비</th></tr></thead><tbody>{raw_rows_html}</tbody></table></div></section>
        <section><h2>4. 3개년 비목 추이</h2>{payload['trend_svg']}<table style="margin-top:12px;"><thead><tr><th>비목</th>{history_header}<th>기준 유형</th><th>신뢰도</th></tr></thead><tbody>{history_rows_html}</tbody></table></section>
        <section class="two-col"><div class="panel"><h2>5. 공식 기준 시트 이력</h2><table><thead><tr><th>연도</th><th>매출원가</th><th>판관비</th><th>영업이익</th><th>R&D</th><th>직접비 믹스(M/L/E)</th><th>source</th></tr></thead><tbody>{official_rows_html}</tbody></table></div><div class="panel"><h2>6. 보정 로그</h2><table><thead><tr><th>비목</th><th>매출 대비</th><th>기준 유형</th><th>신뢰도</th><th>적용 로직</th><th>근거 계정</th></tr></thead><tbody>{adjustment_rows_html}</tbody></table></div></section>
        <section class="two-col"><div><h2>7. 방법론 및 해석</h2><ul class="notes-list">{notes_html}</ul></div><div><h2>8. 국가통계 및 참고 출처</h2><ul class="source-list">{source_rows_html}</ul></div></section>
      </main>
    </div>
    <div class="footnote">{escape(payload['source_note'])} 현재 workbook의 연도별 공식 기준 행은 국가통계 exact raw import가 아니라, 국가통계 앵커와 workbook 복원 이력을 혼합한 standard row입니다.</div>
  </div>
</body>
</html>
"""
