from __future__ import annotations

import hashlib
import json
from collections import defaultdict
from html import escape
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd

from app.analysis.industry_financial_db import (
    append_company_tag_evidence,
    append_evaluation_run,
    extract_company_tags,
    get_baseline_items,
    load_industry_db,
    resolve_workbook_path,
)


ITEM_ORDER = ["material", "labor", "expense", "gna", "profit", "rnd"]
ITEM_META: Dict[str, Dict[str, str]] = {
    "material": {"label": "재료비", "color": "#3b82f6"},
    "labor": {"label": "노무비", "color": "#14b8a6"},
    "expense": {"label": "경비", "color": "#f59e0b"},
    "gna": {"label": "일반관리비", "color": "#8b5cf6"},
    "profit": {"label": "이윤", "color": "#ef4444"},
    "rnd": {"label": "R&D", "color": "#0f766e"},
}
SEVERITY_RANK = {"ok": 0, "watch": 1, "alert": 2, "critical": 3}
SEVERITY_LABEL = {"ok": "OK", "watch": "WATCH", "alert": "ALERT", "critical": "CRITICAL"}
CONFIDENCE_RANK = {"A": 0, "B": 1, "C": 2, "D": 3}
CONFIDENCE_PENALTY_WIDTH = {"A": 0.0, "B": 0.5, "C": 1.0, "D": 1.5}
BASIS_TYPE_LABEL = {
    "observed": "실측 기반",
    "repaired": "복원 기반",
    "proxy": "업종 프록시",
    "estimated": "보수 추정",
}
CONFIDENCE_GRADE_LABEL = {
    "A": "높음",
    "B": "보통",
    "C": "낮음",
    "D": "매우 낮음",
}

LEGACY_TAG_MAP = {
    "유지보수": "maintenance_service",
    "현장시공": "onsite_installation",
    "프로젝트": "project_based",
    "외주높음": "outsourcing_level:high",
    "자산집약": "asset_heavy",
    "노무집약": "labor_intensive",
    "R&D집약": "rnd_intensity:high",
}
QUOTE_SUBTYPE_TAGS = {
    "유지보수형": ["maintenance_service"],
    "프로젝트형": ["project_based"],
    "현장시공형": ["onsite_installation"],
}


def _clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def _safe_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except Exception:
        return None


def _format_pct(value: Optional[float]) -> str:
    if value is None:
        return "-"
    return f"{value:.2f}%"


def _format_amount(value: float) -> str:
    return f"{value:,.0f}원"


def _basis_type_label(value: Any) -> str:
    key = str(value or "").strip()
    return BASIS_TYPE_LABEL.get(key, key or "-")


def _confidence_grade_label(value: Any) -> str:
    key = str(value or "").strip()
    label = CONFIDENCE_GRADE_LABEL.get(key, key or "-")
    return f"{label} ({key})" if key else label


def amount_to_short(value: float) -> str:
    if value >= 100000000:
        return f"{value / 100000000:.2f}억원"
    if value >= 10000:
        return f"{value / 10000:.0f}만원"
    return _format_amount(value)


def _normalize_ratios(ratios: Dict[str, float], active_items: Iterable[str]) -> Dict[str, float]:
    active = list(active_items)
    sanitized = {item: max(ratios.get(item, 0.0), 0.0) for item in active}
    total = sum(sanitized.values())
    if total <= 0:
        even = 100.0 / max(1, len(active))
        return {item: even for item in active}
    return {item: sanitized[item] / total * 100.0 for item in active}


def _shift_ratio(ratios: Dict[str, float], target_item: str, delta: float, active_items: Iterable[str]) -> Dict[str, float]:
    active = [item for item in active_items if item in ratios]
    if target_item not in active or abs(delta) < 1e-9:
        return dict(ratios)

    updated = dict(ratios)
    others = [item for item in active if item != target_item and updated.get(item, 0.0) > 0]
    if not others:
        updated[target_item] = max(0.0, updated.get(target_item, 0.0) + delta)
        return _normalize_ratios(updated, active)

    if delta > 0:
        available = sum(updated[item] for item in others)
        delta = min(delta, available)
        for item in others:
            share = updated[item] / available if available else 0.0
            updated[item] = max(0.0, updated[item] - delta * share)
        updated[target_item] = updated.get(target_item, 0.0) + delta
    else:
        reducible = min(abs(delta), updated.get(target_item, 0.0))
        updated[target_item] = max(0.0, updated.get(target_item, 0.0) - reducible)
        available = sum(updated[item] for item in others)
        for item in others:
            share = updated[item] / available if available else 1.0 / len(others)
            updated[item] += reducible * share
    return _normalize_ratios(updated, active)


def _build_quote_amounts(df_input: pd.DataFrame) -> Dict[str, float]:
    df = df_input.copy()
    df["금액"] = pd.to_numeric(df["단가"], errors="coerce").fillna(0) * pd.to_numeric(df["수량"], errors="coerce").fillna(0)
    text_series = df["비목"].astype(str)
    amounts = {
        "material": float(df[text_series.str.contains("재료|자재|원재료|부품", na=False)]["금액"].sum()),
        "labor": float(df[text_series.str.contains("노무|인건|급여|인력", na=False)]["금액"].sum()),
        "expense": float(df[text_series.str.contains("경비|제경비|운영비|간접비", na=False)]["금액"].sum()),
        "gna": float(df[text_series.str.contains("관리|일반관리", na=False)]["금액"].sum()),
        "profit": float(df[text_series.str.contains("이윤|마진|이익|기술료", na=False)]["금액"].sum()),
        "rnd": float(df[text_series.str.contains("r&d|연구개발|개발비", na=False, case=False)]["금액"].sum()),
    }
    amounts["total"] = float(df["금액"].sum())
    return amounts


def _build_quote_ratios(amounts: Dict[str, float], active_items: Iterable[str]) -> Dict[str, float]:
    total = amounts.get("total", 0.0)
    if total <= 0:
        return {item: 0.0 for item in active_items}
    return {item: amounts.get(item, 0.0) / total * 100.0 for item in active_items}


def _build_line_item_rows(df_input: pd.DataFrame) -> List[Dict[str, Any]]:
    df = df_input.copy()
    desc_col = next((col for col in df.columns if str(col) in ["품명(상세내역)", "내역(상세내역)", "상세내역", "내역", "품명"]), None)
    desc_values = df[desc_col].fillna("-").astype(str) if desc_col else pd.Series(["-"] * len(df))
    df["금액"] = pd.to_numeric(df["단가"], errors="coerce").fillna(0) * pd.to_numeric(df["수량"], errors="coerce").fillna(0)
    rows = []
    for _, row in df.sort_values("금액", ascending=False).head(20).iterrows():
        rows.append(
            {
                "item": str(row.get("비목", "-")),
                "detail": desc_values.loc[row.name],
                "qty": _safe_float(row.get("수량")) or 0.0,
                "unit_price": _safe_float(row.get("단가")) or 0.0,
                "amount": _safe_float(row.get("금액")) or 0.0,
            }
        )
    return rows


def _stable_digest(payload: Any) -> str:
    encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _manual_tags_from_meta(meta: Dict[str, Any]) -> List[str]:
    manual = []
    for tag in meta.get("manual_tags") or []:
        tag = LEGACY_TAG_MAP.get(str(tag), str(tag))
        if tag not in manual:
            manual.append(tag)
    for tag in QUOTE_SUBTYPE_TAGS.get(str(meta.get("quote_subtype", "")), []):
        if tag not in manual:
            manual.append(tag)
    return manual


def _tag_cap(range_df: pd.DataFrame, item_code: str, basis_type: str) -> float:
    match = range_df[
        (range_df["item_code"].astype(str) == item_code)
        & (range_df["basis_type"].astype(str) == basis_type)
    ]
    if match.empty:
        return 5.0
    return float(match.iloc[0]["tag_cap_pctp"])


def _range_width(range_df: pd.DataFrame, item_code: str, basis_type: str, confidence_grade: str) -> float:
    match = range_df[
        (range_df["item_code"].astype(str) == item_code)
        & (range_df["basis_type"].astype(str) == basis_type)
    ]
    base_width = float(match.iloc[0]["base_width_pctp"]) if not match.empty else 5.0
    return base_width + CONFIDENCE_PENALTY_WIDTH.get(confidence_grade, 1.5)


def _severity_thresholds(severity_df: pd.DataFrame, confidence_grade: str) -> Tuple[float, float]:
    match = severity_df[severity_df["confidence_grade"].astype(str) == confidence_grade]
    if match.empty:
        return 4.0, 8.0
    row = match.iloc[0]
    return float(row["alert_gap_pctp"]), float(row["critical_gap_pctp"])


def _classify_ratio(actual: float, low: float, high: float, confidence_grade: str, severity_df: pd.DataFrame) -> Tuple[str, float]:
    if low <= actual <= high:
        return "ok", 0.0
    gap = low - actual if actual < low else actual - high
    alert_threshold, critical_threshold = _severity_thresholds(severity_df, confidence_grade)
    if gap >= critical_threshold:
        return "critical", gap
    if gap >= alert_threshold:
        return "alert", gap
    return "watch", gap


def _apply_tag_adjustments(
    base_ratios: Dict[str, float],
    active_items: List[str],
    tag_rows: List[Dict[str, Any]],
    tag_rule_df: pd.DataFrame,
    range_df: pd.DataFrame,
    basis_map: Dict[str, str],
) -> Tuple[Dict[str, float], Dict[str, List[str]], List[str]]:
    updated = dict(base_ratios)
    reason_map: Dict[str, List[str]] = defaultdict(list)
    notes: List[str] = []
    cumulative: Dict[str, float] = defaultdict(float)
    rule_lookup = {
        (str(row["tag_code"]), str(row["tag_value"])): row
        for _, row in tag_rule_df.iterrows()
        if str(row.get("enabled", "Y")).strip().upper() == "Y"
    }
    for tag_row in tag_rows:
        rule = rule_lookup.get((tag_row["tag_code"], tag_row["tag_value"]))
        if rule is None and tag_row["tag_value"] != "true":
            rule = rule_lookup.get((tag_row["tag_code"], "true"))
        if rule is None:
            continue
        notes.append(f"{tag_row['label']} 반영")
        for item in active_items:
            delta = float(rule.get(f"{item}_delta", 0.0) or 0.0)
            if abs(delta) < 1e-9:
                continue
            cap = _tag_cap(range_df, item, basis_map.get(item, "estimated"))
            remaining = max(0.0, cap - abs(cumulative[item]))
            applied = max(-remaining, min(delta, remaining))
            if abs(applied) < 1e-9:
                continue
            updated = _shift_ratio(updated, item, applied, active_items)
            cumulative[item] += applied
            reason_map[item].append(f"TAG_{tag_row['tag_code'].upper()}_{tag_row['tag_value'].upper()}")
    return _normalize_ratios(updated, active_items), reason_map, notes


def _overall_confidence(item_rows: List[Dict[str, Any]]) -> Tuple[float, str]:
    if not item_rows:
        return 0.55, "D"
    avg = sum(float(row["confidence_score"]) for row in item_rows) / len(item_rows)
    worst = max(item_rows, key=lambda row: CONFIDENCE_RANK.get(row["confidence_grade"], 99))["confidence_grade"]
    return avg, worst


def _build_structure_checks(quote_ratios: Dict[str, float], item_rows: List[Dict[str, Any]], tag_rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    checks: List[Dict[str, Any]] = []
    total_ratio = sum(quote_ratios.values())
    total_gap = abs(total_ratio - 100.0)
    checks.append(
        {
            "code": "SUM_RATIO_OK" if total_gap <= 1.0 else "SUM_RATIO_MISMATCH",
            "severity": "ok" if total_gap <= 1.0 else ("alert" if total_gap <= 3.0 else "critical"),
            "message": (
                f"비목 합계가 {total_ratio:.2f}%로 정합성이 양호합니다."
                if total_gap <= 1.0
                else f"비목 합계가 {total_ratio:.2f}%로 100%와 차이가 있습니다."
            ),
        }
    )
    by_item = {row["item"]: row for row in item_rows}
    if "gna" in by_item and quote_ratios.get("gna", 0.0) > max(by_item["gna"]["high"], 12.0):
        checks.append({"code": "STRUCT_GNA_HIGH", "severity": "alert", "message": "일반관리비가 기준 범위를 상회합니다."})
    if "profit" in by_item and quote_ratios.get("profit", 0.0) > max(by_item["profit"]["high"], 15.0):
        checks.append({"code": "STRUCT_PROFIT_HIGH", "severity": "watch", "message": "이윤이 높은 편이라 재협상 여지가 있습니다."})
    if "profit" in by_item and quote_ratios.get("profit", 0.0) < min(by_item["profit"]["low"], 3.0):
        checks.append({"code": "STRUCT_PROFIT_LOW", "severity": "watch", "message": "이윤이 낮아 다른 비목으로 마진이 숨겨졌는지 확인이 필요합니다."})
    tag_index = {f"{row['tag_code']}:{row['tag_value']}" for row in tag_rows}
    if "outsourcing_level:high" in tag_index and quote_ratios.get("labor", 0.0) > quote_ratios.get("expense", 0.0) + 5.0:
        checks.append({"code": "STRUCT_OUTSOURCE_MISMATCH", "severity": "watch", "message": "외주높음 태그 대비 노무비 비중이 높습니다."})
    return checks


def _build_consistency_checks(quote_ratios: Dict[str, float], meta: Dict[str, Any]) -> List[Dict[str, Any]]:
    checks: List[Dict[str, Any]] = []
    observed_map = {
        "material": meta.get("observed_material_ratio"),
        "labor": meta.get("observed_labor_ratio"),
        "expense": meta.get("observed_expense_ratio"),
        "gna": meta.get("observed_gna_ratio"),
        "profit": meta.get("observed_profit_ratio"),
        "rnd": meta.get("observed_rnd_ratio"),
    }
    for item, observed in observed_map.items():
        observed_value = _safe_float(observed)
        if observed_value is None or item not in quote_ratios:
            continue
        if abs(observed_value) <= 1.0:
            observed_value *= 100.0
        gap = abs(quote_ratios[item] - observed_value)
        if gap > 5.0:
            severity = "alert"
        elif gap > 2.5:
            severity = "watch"
        else:
            continue
        checks.append(
            {
                "code": f"CONSISTENCY_{item.upper()}_MISMATCH",
                "severity": severity,
                "message": f"{ITEM_META[item]['label']}이 최근 회사 관측치와 {gap:.2f}%p 차이납니다.",
            }
        )
    profit_ratio = quote_ratios.get("profit", 0.0)
    if _safe_float(meta.get("sales_growth_yoy")) is not None and float(meta["sales_growth_yoy"]) < 0 and profit_ratio > 12.0:
        checks.append({"code": "CONSISTENCY_NEG_YOY_HIGH_PROFIT", "severity": "watch", "message": "매출 역성장 대비 이윤이 높은 편입니다."})
    if _safe_float(meta.get("op_margin_delta_yoy")) is not None and float(meta["op_margin_delta_yoy"]) < 0 and profit_ratio > 12.0:
        checks.append({"code": "CONSISTENCY_MARGIN_DOWN_HIGH_PROFIT", "severity": "watch", "message": "마진 하락 추세 대비 이윤이 높은 편입니다."})
    return checks


def _overall_judgement(item_rows: List[Dict[str, Any]], structure_checks: List[Dict[str, Any]], consistency_checks: List[Dict[str, Any]]) -> str:
    severities = [row["judgement"] for row in item_rows]
    severities.extend(check["severity"] for check in structure_checks)
    severities.extend(check["severity"] for check in consistency_checks)
    if not severities:
        return "ok"
    return max(severities, key=lambda severity: SEVERITY_RANK.get(severity, 0))


def _top_findings(item_rows: List[Dict[str, Any]], checks: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    findings = [
        {
            "title": row["finding_title"],
            "description": f"실제 {row['actual_ratio']:.2f}% vs target {row['low']:.1f}~{row['high']:.1f}%",
            "severity": row["judgement"],
        }
        for row in item_rows
        if row["judgement"] != "ok"
    ]
    findings.extend({"title": check["code"], "description": check["message"], "severity": check["severity"]} for check in checks if check["severity"] != "ok")
    findings.sort(key=lambda finding: SEVERITY_RANK.get(finding["severity"], 0), reverse=True)
    return findings[:6]


def evaluate_quote_adequacy(df_input: pd.DataFrame, ui_category: str, meta: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    meta = meta or {}
    workbook_path = resolve_workbook_path(meta.get("industry_db_path"))
    industry_group = str(meta.get("industry_group") or "").strip().upper()
    industry_group_name = str(meta.get("industry_group_name") or "").strip()
    baseline_period = str(meta.get("baseline_period") or "").strip()
    reference_mode = str(meta.get("reference_mode") or "").strip()
    reference_source_label = str(meta.get("reference_source_label") or "").strip()
    supplier_financial_mode = reference_mode == "uploaded_observed"
    estimated_partial_mode = reference_mode == "estimated_partial"
    quick_proxy_mode = reference_mode == "quick_industry_proxy"
    if supplier_financial_mode:
        reference_profile_label = reference_source_label or "협력사 업로드 재무제표"
        quantiles_note_default = "분위수 불가: 협력사 업로드 재무제표 기준이라 분포 통계 대신 실측 Base를 사용합니다."
        limitation_note = "적정 범위의 base는 협력사 업로드 재무제표에서 관측한 비목 비율과 태그 보정으로 계산했습니다."
        distribution_reason_code = "SUPPLIER_FINANCIAL_BASELINE"
    elif estimated_partial_mode:
        reference_profile_label = reference_source_label or "일부 업로드 + 업종 추정 기준"
        quantiles_note_default = "분위수 불가: 일부 업로드 재무자료만 있어 부족한 항목은 업종 aggregate 기준으로 보완했습니다."
        limitation_note = "적정 범위의 base는 일부 업로드 재무자료를 반영했지만, 부족한 항목은 업종 baseline과 태그 보정으로 계산했습니다."
        distribution_reason_code = "AGGREGATE_ONLY_NO_DISTRIBUTION"
    elif quick_proxy_mode:
        reference_profile_label = reference_source_label or "사업자등록증 기반 업종 프록시"
        quantiles_note_default = "분위수 불가: 사업자등록증 기반 업종 aggregate 기준만 있습니다."
        limitation_note = "적정 범위의 base는 사업자등록증 기반 업종 프록시와 태그 보정으로 계산했습니다."
        distribution_reason_code = "AGGREGATE_ONLY_NO_DISTRIBUTION"
    else:
        reference_profile_label = reference_source_label or (industry_group_name or industry_group)
        quantiles_note_default = "분위수 불가: 업종 aggregate 기준만 있습니다."
        limitation_note = "적정 범위의 base는 기업값이 아니라 업종 baseline과 태그 보정만으로 계산했습니다."
        distribution_reason_code = "AGGREGATE_ONLY_NO_DISTRIBUTION"
    if not industry_group or not baseline_period:
        raise ValueError("industry_group와 baseline_period를 선택해야 비목 적정성 판정을 수행할 수 있습니다.")

    sheets = load_industry_db(workbook_path)
    range_df = sheets["config_range"]
    severity_df = sheets["config_severity"]
    tag_rule_df = sheets["config_tag_rules"]
    baseline_items = get_baseline_items(industry_group, baseline_period, workbook_path=workbook_path)
    baseline_lookup = {row["item_code"]: row for row in baseline_items}

    amounts = _build_quote_amounts(df_input)
    manual_tags = _manual_tags_from_meta(meta)
    tag_rows = extract_company_tags(meta.get("company_doc_summary", ""), manual_tags, workbook_path=workbook_path)
    for row in tag_rows:
        row["tag"] = row["tag_code"] if row["tag_value"] == "true" else f"{row['tag_code']}:{row['tag_value']}"
        row["confidence"] = row["confidence_score"]

    include_rnd = (
        amounts.get("rnd", 0.0) > 0
        or "rnd" in baseline_lookup
        or any(row["tag_code"] == "rnd_intensity" and row["tag_value"] in {"high", "med"} for row in tag_rows)
        or _safe_float(meta.get("observed_rnd_ratio")) not in (None, 0.0)
    )
    active_items = [item for item in ITEM_ORDER if item != "rnd" or include_rnd]
    if "rnd" not in baseline_lookup and "rnd" in active_items:
        baseline_lookup["rnd"] = {
            "item_code": "rnd",
            "label": "R&D",
            "base_pct": 0.0,
            "basis_type": "estimated",
            "confidence_score": 0.55,
            "confidence_grade": "D",
            "supporting_accounts": "",
            "calc_rule_code": "TAG_ONLY_RND_SLOT",
            "inputs_digest": baseline_items[0]["inputs_digest"] if baseline_items else "",
            "restatement_flag": False,
            "quantiles_note": quantiles_note_default,
        }

    base_ratios = {item: float(baseline_lookup[item]["base_pct"]) for item in active_items if item in baseline_lookup}
    basis_map = {item: str(baseline_lookup[item]["basis_type"]) for item in active_items if item in baseline_lookup}
    adjusted_base, tag_reason_map, tag_notes = _apply_tag_adjustments(base_ratios, active_items, tag_rows, tag_rule_df, range_df, basis_map)
    quote_ratios = _build_quote_ratios(amounts, active_items)

    item_rows: List[Dict[str, Any]] = []
    for item in active_items:
        baseline_row = baseline_lookup[item]
        confidence_grade = str(baseline_row["confidence_grade"])
        width = _range_width(range_df, item, str(baseline_row["basis_type"]), confidence_grade)
        base_value = adjusted_base.get(item, 0.0)
        low = _clamp(base_value - width, 0.0, 100.0)
        high = _clamp(base_value + width, 0.0, 100.0)
        actual = quote_ratios.get(item, 0.0)
        judgement, gap = _classify_ratio(actual, low, high, confidence_grade, severity_df)

        direction = "within"
        if actual < low:
            direction = "low"
        elif actual > high:
            direction = "high"

        reason_codes = []
        if direction == "high":
            reason_codes.append(f"{item.upper()}_HIGH")
        elif direction == "low":
            reason_codes.append(f"{item.upper()}_LOW")
        else:
            reason_codes.append("WITHIN_TARGET_RANGE")
        reason_codes.append(distribution_reason_code)
        reason_codes.extend(tag_reason_map.get(item, []))

        notes = [f"기준 출처={reference_profile_label}", quantiles_note_default, f"기준 유형={_basis_type_label(baseline_row['basis_type'])}"]
        if tag_notes:
            notes.extend(tag_notes[:2])

        item_rows.append(
            {
                "item": item,
                "label": ITEM_META[item]["label"],
                "amount": amounts.get(item, 0.0),
                "actual_ratio": actual,
                "base": base_value,
                "low": low,
                "high": high,
                "judgement": judgement,
                "gap": gap,
                "basis_type": baseline_row["basis_type"],
                "confidence_score": float(baseline_row["confidence_score"]),
                "confidence_grade": confidence_grade,
                "reason_codes": sorted(set(reason_codes)),
                "reason_summary": " | ".join([", ".join(sorted(set(reason_codes)))] + notes[:3]),
                "notes": notes,
                "finding_title": f"{ITEM_META[item]['label']} {'적정' if judgement == 'ok' else '주의'}",
                "supporting_accounts": baseline_row.get("supporting_accounts", ""),
                "applied_rules": sorted(set(reason_codes)),
                "calc_rule_code": baseline_row.get("calc_rule_code", ""),
                "tag_evidence": [row["evidence_text"] for row in tag_rows if any(code.startswith("TAG_") for code in reason_codes)],
                "color": ITEM_META[item]["color"],
            }
        )

    structure_checks = _build_structure_checks(quote_ratios, item_rows, tag_rows)
    consistency_checks = _build_consistency_checks(quote_ratios, meta)
    overall = _overall_judgement(item_rows, structure_checks, consistency_checks)
    findings = _top_findings(item_rows, structure_checks + consistency_checks)
    base_confidence, overall_conf_grade = _overall_confidence(item_rows)

    summary_text = "비목 구조가 기준 범위에 대체로 부합합니다."
    if overall == "critical":
        summary_text = "복수 비목이 기준 범위를 크게 벗어나며 재분류 또는 과다 계상 가능성이 큽니다."
    elif overall == "alert":
        summary_text = "일부 비목이 기준 범위를 유의미하게 벗어나 추가 근거 확인이 필요합니다."
    elif overall == "watch":
        summary_text = "대체로 합리적이지만 경계 구간의 비목이 있습니다."

    quantiles_note = quantiles_note_default if supplier_financial_mode else (baseline_items[0]["quantiles_note"] if baseline_items else quantiles_note_default)
    limitations = [
        quantiles_note,
        limitation_note,
        "제조원가명세서가 불완전하면 재료/노무/경비는 proxy/estimated로 낮춰 해석합니다.",
    ]

    quote_digest = _stable_digest(df_input.fillna("").astype(str).to_dict(orient="records"))
    report_text_digest = _stable_digest(meta.get("company_doc_summary", "") or "")
    evaluation_id = append_evaluation_run(
        company_name=str(meta.get("company_name", "")).strip(),
        industry_group=industry_group,
        period_key=baseline_period,
        quote_digest=quote_digest,
        baseline_inputs_digest=str(baseline_items[0]["inputs_digest"]) if baseline_items else "",
        report_text_digest=report_text_digest,
        overall_judgement=overall,
        overall_confidence_grade=overall_conf_grade,
        restatement_seen=bool(any(str(row.get("restatement_flag", False)).lower() == "true" for row in baseline_items)),
        workbook_path=workbook_path,
    )
    append_company_tag_evidence(
        evaluation_id=evaluation_id,
        company_name=str(meta.get("company_name", "")).strip(),
        industry_group=industry_group,
        period_key=baseline_period,
        report_text_digest=report_text_digest,
        tag_rows=tag_rows,
        workbook_path=workbook_path,
    )

    return {
        "profile_key": f"{industry_group}:{baseline_period}",
        "profile_label": f"{reference_profile_label} {baseline_period} baseline",
        "profile_family": ui_category,
        "quote_subtype": meta.get("quote_subtype", "자동"),
        "meta": meta,
        "quote_amounts": amounts,
        "quote_ratios": quote_ratios,
        "total_ratio": sum(quote_ratios.values()),
        "tag_rows": tag_rows,
        "item_rows": item_rows,
        "structure_checks": structure_checks,
        "consistency_checks": consistency_checks,
        "overall_judgement": overall,
        "overall_label": SEVERITY_LABEL[overall],
        "overall_summary": summary_text,
        "findings": findings,
        "limitations": limitations,
        "base_confidence": base_confidence,
        "base_confidence_bucket": overall_conf_grade,
        "line_items": _build_line_item_rows(df_input),
        "financial_notes": [check["message"] for check in consistency_checks],
        "baseline_inputs_digest": str(baseline_items[0]["inputs_digest"]) if baseline_items else "",
        "quantiles_note": quantiles_note,
        "evaluation_id": evaluation_id,
        "reference_mode": reference_mode,
        "reference_source_label": reference_profile_label,
    }


def build_evaluation_dataframe(report: Dict[str, Any]) -> pd.DataFrame:
    rows = []
    for item_row in report["item_rows"]:
        rows.append(
            {
                "비목": item_row["label"],
                "금액": _format_amount(item_row["amount"]),
                "실제 비중": _format_pct(item_row["actual_ratio"]),
                "Target Range": f"{item_row['low']:.1f}% ~ {item_row['high']:.1f}%",
                "Base": _format_pct(item_row["base"]),
                "판정": SEVERITY_LABEL[item_row["judgement"]],
                "기준 유형": _basis_type_label(item_row["basis_type"]),
                "신뢰도": _confidence_grade_label(item_row["confidence_grade"]),
                "근거 계정": item_row["supporting_accounts"],
                "reason_code / 해석": item_row["reason_summary"],
            }
        )
    return pd.DataFrame(rows)


def build_llm_context(report: Dict[str, Any]) -> str:
    meta = report["meta"]
    display_group = str(meta.get("industry_group_name") or meta.get("industry_group") or "-")
    raw_group = str(meta.get("industry_group") or "-")
    lines = [
        "[입력 메타데이터]",
        f"- 공급사/기업명: {meta.get('company_name', '-')}",
        f"- industry_group: {display_group} ({raw_group})" if raw_group != "-" else f"- industry_group: {display_group}",
        f"- baseline_period: {meta.get('baseline_period', '-')}",
        f"- industry_code(KSIC): {meta.get('industry_code', '-')}",
        f"- scope: {meta.get('scope', '-')}",
        f"- quantiles: {report.get('quantiles_note', '-')}",
        "",
        "[태그 근거]",
    ]
    if report["tag_rows"]:
        for tag_row in report["tag_rows"]:
            lines.append(
                f"- {tag_row['label']} ({tag_row['tag_code']}={tag_row['tag_value']}): {tag_row['evidence_text']} | confidence={tag_row['confidence_score']:.2f}"
            )
    else:
        lines.append("- 태그 없음 또는 근거 부족")

    lines.extend(["", "[비목 판정]"])
    for item_row in report["item_rows"]:
        lines.append(
            f"- {item_row['label']}: actual={item_row['actual_ratio']:.2f}% | target={item_row['low']:.1f}~{item_row['high']:.1f}% | "
            f"base={item_row['base']:.2f}% | judgement={SEVERITY_LABEL[item_row['judgement']]} | "
            f"basis_type={_basis_type_label(item_row['basis_type'])} | confidence={_confidence_grade_label(item_row['confidence_grade'])} | "
            f"accounts={item_row['supporting_accounts']} | rules={','.join(item_row['applied_rules'])}"
        )
    lines.extend(["", "[구조/정합성 룰]"])
    for check in report["structure_checks"] + report.get("consistency_checks", []):
        lines.append(f"- [{check['severity'].upper()}] {check['code']} - {check['message']}")
    lines.extend(["", "[한계]", *[f"- {note}" for note in report["limitations"]]])
    return "\n".join(lines)


def generate_html_report(report: Dict[str, Any], ai_commentary: str = "") -> str:
    item_rows = "".join(
        "<tr>"
        f"<td>{escape(row['label'])}</td>"
        f"<td>{_format_amount(row['amount'])}</td>"
        f"<td>{row['actual_ratio']:.2f}%</td>"
        f"<td>{row['low']:.1f}% ~ {row['high']:.1f}%<br/><span style='color:#64748b'>Base {row['base']:.2f}%</span></td>"
        f"<td><strong>{SEVERITY_LABEL[row['judgement']]}</strong></td>"
        f"<td>{escape(_basis_type_label(row['basis_type']))}</td>"
        f"<td>{escape(_confidence_grade_label(row['confidence_grade']))}</td>"
        f"<td>{escape(str(row['supporting_accounts']))}</td>"
        f"<td>{escape(row['reason_summary'])}</td>"
        "</tr>"
        for row in report["item_rows"]
    )
    tag_rows = "".join(
        f"<li><strong>{escape(row['label'])}</strong> ({escape(row['tag_code'])}={escape(row['tag_value'])}) - {escape(row['evidence_text'])} · {row['confidence_score']:.2f}</li>"
        for row in report["tag_rows"]
    ) or "<li>태그 근거 없음</li>"
    finding_rows = "".join(
        f"<li><strong>{escape(finding['title'])}</strong> - {escape(finding['description'])} ({finding['severity']})</li>"
        for finding in report["findings"]
    ) or "<li>주요 이슈 없음</li>"
    line_rows = "".join(
        "<tr>"
        f"<td>{escape(str(row['item']))}</td>"
        f"<td>{escape(str(row['detail']))}</td>"
        f"<td>{row['qty']:.2f}</td>"
        f"<td>{_format_amount(row['unit_price'])}</td>"
        f"<td>{_format_amount(row['amount'])}</td>"
        "</tr>"
        for row in report["line_items"]
    )
    limitations = "".join(f"<li>{escape(note)}</li>" for note in report["limitations"])
    checks = "".join(
        f"<li>[{escape(check['severity'].upper())}] {escape(check['code'])} - {escape(check['message'])}</li>"
        for check in report["structure_checks"] + report.get("consistency_checks", [])
    )

    return f"""
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title>Quote Adequacy Report</title>
  <style>
    body {{ font-family: 'Malgun Gothic', sans-serif; margin: 24px; color: #1f2937; }}
    h1, h2 {{ color: #0f172a; }}
    .kpis {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 24px; }}
    .kpi {{ border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px; background: #f8fafc; }}
    .label {{ color: #64748b; font-size: 12px; }}
    .value {{ font-size: 22px; font-weight: 700; margin-top: 6px; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 12px; }}
    th, td {{ border: 1px solid #e2e8f0; padding: 10px; font-size: 13px; vertical-align: top; }}
    th {{ background: #f1f5f9; text-align: left; }}
    section {{ margin-top: 24px; }}
    ul {{ padding-left: 20px; }}
    .note {{ color: #475569; }}
    .ai {{ white-space: pre-wrap; background: #eff6ff; border-left: 4px solid #2563eb; padding: 12px; border-radius: 8px; }}
  </style>
</head>
<body>
  <h1>비목 적정성 판정 리포트</h1>
  <div class="note">{escape(report['profile_label'])} · {escape(report.get('quantiles_note', '-'))}</div>
  <div class="kpis">
    <div class="kpi"><div class="label">종합 판정</div><div class="value">{escape(report['overall_label'])}</div></div>
    <div class="kpi"><div class="label">견적 총액</div><div class="value">{amount_to_short(report['quote_amounts']['total'])}</div></div>
    <div class="kpi"><div class="label">신뢰도</div><div class="value">{escape(_confidence_grade_label(report['base_confidence_bucket']))}</div></div>
    <div class="kpi"><div class="label">baseline digest</div><div class="value" style="font-size:12px">{escape(report['baseline_inputs_digest'])}</div></div>
  </div>

  <section>
    <h2>핵심 요약</h2>
    <p>{escape(report['overall_summary'])}</p>
    <ul>{finding_rows}</ul>
  </section>

  <section>
    <h2>태그 근거</h2>
    <ul>{tag_rows}</ul>
  </section>

  <section>
    <h2>비목별 판정</h2>
    <table>
      <thead>
        <tr><th>비목</th><th>금액</th><th>실제 비중</th><th>Target Range</th><th>판정</th><th>기준 유형</th><th>신뢰도</th><th>근거 계정</th><th>해석</th></tr>
      </thead>
      <tbody>{item_rows}</tbody>
    </table>
  </section>

  <section>
    <h2>구조/정합성 룰</h2>
    <ul>{checks}</ul>
  </section>

  <section>
    <h2>견적 상세</h2>
    <table>
      <thead>
        <tr><th>비목</th><th>상세</th><th>수량</th><th>단가</th><th>금액</th></tr>
      </thead>
      <tbody>{line_rows}</tbody>
    </table>
  </section>

  <section>
    <h2>한계</h2>
    <ul>{limitations}</ul>
  </section>

  <section>
    <h2>AI 메모</h2>
    <div class="ai">{escape(ai_commentary or '-')}</div>
  </section>
</body>
</html>
"""
