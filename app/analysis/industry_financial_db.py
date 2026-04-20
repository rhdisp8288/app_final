from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import pypdfium2 as pdfium
from rapidocr_onnxruntime import RapidOCR


DEFAULT_WORKBOOK_FILENAME = "industry_financial_db.xlsx"
DEFAULT_RAW_DIRNAME = "financial_raw"
DEFAULT_BATCH_PERIOD_LABEL = "financial_raw batch import"
STATEMENT_TYPES = ("BS", "IS", "CF", "COST")
EXCEL_STATEMENT_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
QUOTE_ITEMS = ["material", "labor", "expense", "gna", "profit", "rnd"]
RAW_FILENAME_RE = re.compile(
    r"^(?P<industry_label>.+?)\((?P<industry_group>[A-Z])\)_(?P<statement_label>.+?)\.(?P<extension>pdf|xlsx|xlsm|xls)$",
    re.IGNORECASE,
)
STATEMENT_FILE_ALIASES = {
    "BS": ("재무상태표", "재무제표"),
    "IS": ("손익계산서", "손인계산서"),
    "CF": ("현금흐름표",),
    "COST": ("제조원가명세서",),
}
STATEMENT_LABEL_TO_TYPE = {
    alias.replace(" ", ""): statement_type
    for statement_type, aliases in STATEMENT_FILE_ALIASES.items()
    for alias in aliases
}

ITEM_LABELS = {
    "material": "재료비",
    "labor": "노무비",
    "expense": "경비",
    "gna": "일반관리비",
    "profit": "이윤",
    "rnd": "R&D",
}

EXCEL_ACCOUNT_ALIASES: Dict[str, Dict[str, Tuple[str, ...]]] = {
    "BS": {
        "bs_total_assets": ("자산총계", "자산"),
        "bs_current_assets": ("유동자산",),
        "bs_inventory": ("재고자산",),
        "bs_ppe": ("유형자산", "유형리스자산"),
    },
    "IS": {
        "is_revenue": ("매출액",),
        "is_cogs": ("매출원가",),
        "is_gross_profit": ("매출총이익손실", "매출총이익"),
        "is_gna": ("판매비와관리비",),
        "is_operating_profit": ("영업이익손실", "영업이익"),
        "is_rnd_research": ("연구비",),
        "is_rnd_development": ("경상개발비",),
        "is_material_nature": ("원재료및소모품사용액과상품의매입액", "원재료및소모품사용액", "상품의매입액"),
        "is_employee_benefits": ("종업원급여", "급여"),
        "is_depreciation": ("감가상각비", "유형임대주택자산감가상각비"),
    },
    "CF": {
        "cf_operating_cash_flow": ("영업활동으로인한현금흐름",),
        "cf_capex": ("유형자산의취득", "유형리스자산의취득", "유형리스자산취득", "유형자산취득"),
    },
    "COST": {
        "cost_material": ("원재료비",),
        "cost_labor": ("노동관계비용", "노무비"),
        "cost_expense": ("경비",),
    },
}

SHEET_SCHEMAS: Dict[str, List[str]] = {
    "statement_rows": [
        "industry_group",
        "statement_type",
        "period_key",
        "page_number",
        "row_index",
        "raw_account_name",
        "normalized_account_code",
        "normalized_account_name",
        "item_code",
        "amount",
        "source_file",
        "source_digest",
        "inputs_digest",
        "restatement_flag",
        "ingested_at",
    ],
    "metric_observations": [
        "industry_group",
        "period_key",
        "metric_code",
        "item_code",
        "value_pct",
        "basis_type",
        "confidence_score",
        "confidence_grade",
        "confidence_note",
        "source_statement",
        "supporting_accounts",
        "cost_presence_flag",
        "inputs_digest",
        "restatement_flag",
        "calc_rule_code",
        "created_at",
    ],
    "industry_baselines": [
        "industry_group",
        "period_key",
        "item_code",
        "label",
        "base_pct",
        "low_pct",
        "high_pct",
        "basis_type",
        "confidence_score",
        "confidence_grade",
        "confidence_note",
        "quantiles_available",
        "quantiles_note",
        "supporting_accounts",
        "calc_rule_code",
        "cost_presence_flag",
        "inputs_digest",
        "restatement_flag",
        "created_at",
    ],
    "company_tag_evidence": [
        "evaluation_id",
        "company_name",
        "industry_group",
        "period_key",
        "report_text_digest",
        "tag_code",
        "tag_value",
        "label",
        "evidence_text",
        "extractor_rule_id",
        "confidence_score",
        "confidence_grade",
        "source",
        "note",
        "created_at",
    ],
    "evaluation_runs": [
        "evaluation_id",
        "company_name",
        "industry_group",
        "period_key",
        "quote_digest",
        "baseline_inputs_digest",
        "report_text_digest",
        "overall_judgement",
        "overall_confidence_grade",
        "restatement_seen",
        "created_at",
    ],
    "config_account_map": [
        "statement_type",
        "page_number",
        "row_index",
        "raw_account_name",
        "normalized_account_code",
        "normalized_account_name",
        "item_code",
        "required_flag",
        "enabled",
    ],
    "config_proxy_mix": [
        "industry_group",
        "basis_type",
        "material_share",
        "labor_share",
        "expense_share",
        "note",
    ],
    "config_official_reference": [
        "industry_group",
        "period_key",
        "cogs_pct",
        "gna_pct",
        "profit_pct",
        "rnd_pct",
        "material_share",
        "labor_share",
        "expense_share",
        "source",
        "note",
    ],
    "config_tag_rules": [
        "tag_code",
        "tag_value",
        "label",
        "keywords",
        "negative_keywords",
        "confidence_base",
        "material_delta",
        "labor_delta",
        "expense_delta",
        "gna_delta",
        "profit_delta",
        "rnd_delta",
        "enabled",
    ],
    "config_range": [
        "item_code",
        "basis_type",
        "base_width_pctp",
        "tag_cap_pctp",
    ],
    "config_severity": [
        "confidence_grade",
        "watch_gap_pctp",
        "alert_gap_pctp",
        "critical_gap_pctp",
    ],
}


class IndustryFinancialError(ValueError):
    pass


DEFAULT_ACCOUNT_MAP_ROWS = [
    {
        "statement_type": "IS",
        "page_number": 1,
        "row_index": 2,
        "raw_account_name": "매출액",
        "normalized_account_code": "is_revenue",
        "normalized_account_name": "매출액",
        "item_code": "",
        "required_flag": "Y",
        "enabled": "Y",
    },
    {
        "statement_type": "IS",
        "page_number": 1,
        "row_index": 14,
        "raw_account_name": "매출원가",
        "normalized_account_code": "is_cogs",
        "normalized_account_name": "매출원가",
        "item_code": "",
        "required_flag": "Y",
        "enabled": "Y",
    },
    {
        "statement_type": "IS",
        "page_number": 1,
        "row_index": 25,
        "raw_account_name": "매출총이익(손실)",
        "normalized_account_code": "is_gross_profit",
        "normalized_account_name": "매출총이익(손실)",
        "item_code": "",
        "required_flag": "N",
        "enabled": "Y",
    },
    {
        "statement_type": "IS",
        "page_number": 1,
        "row_index": 26,
        "raw_account_name": "판매비와관리비",
        "normalized_account_code": "is_gna",
        "normalized_account_name": "판매비와관리비",
        "item_code": "gna",
        "required_flag": "Y",
        "enabled": "Y",
    },
    {
        "statement_type": "IS",
        "page_number": 2,
        "row_index": 16,
        "raw_account_name": "원재료 및 소모품 사용액과 상품의 매입액",
        "normalized_account_code": "is_material_nature",
        "normalized_account_name": "원재료 및 소모품 사용액과 상품의 매입액",
        "item_code": "material",
        "required_flag": "N",
        "enabled": "N",
    },
    {
        "statement_type": "IS",
        "page_number": 2,
        "row_index": 17,
        "raw_account_name": "종업원급여",
        "normalized_account_code": "is_employee_benefits",
        "normalized_account_name": "종업원급여",
        "item_code": "labor",
        "required_flag": "N",
        "enabled": "N",
    },
    {
        "statement_type": "IS",
        "page_number": 2,
        "row_index": 18,
        "raw_account_name": "감가상각비",
        "normalized_account_code": "is_depreciation",
        "normalized_account_name": "감가상각비",
        "item_code": "expense",
        "required_flag": "N",
        "enabled": "N",
    },
    {
        "statement_type": "IS",
        "page_number": 2,
        "row_index": 19,
        "raw_account_name": "연구비",
        "normalized_account_code": "is_rnd_research",
        "normalized_account_name": "연구비",
        "item_code": "rnd",
        "required_flag": "N",
        "enabled": "Y",
    },
    {
        "statement_type": "IS",
        "page_number": 2,
        "row_index": 20,
        "raw_account_name": "경상개발비",
        "normalized_account_code": "is_rnd_development",
        "normalized_account_name": "경상개발비",
        "item_code": "rnd",
        "required_flag": "N",
        "enabled": "Y",
    },
    {
        "statement_type": "IS",
        "page_number": 2,
        "row_index": 21,
        "raw_account_name": "영업이익(손실)",
        "normalized_account_code": "is_operating_profit",
        "normalized_account_name": "영업이익(손실)",
        "item_code": "profit",
        "required_flag": "Y",
        "enabled": "Y",
    },
    {
        "statement_type": "COST",
        "page_number": 1,
        "row_index": 2,
        "raw_account_name": "원재료비",
        "normalized_account_code": "cost_material",
        "normalized_account_name": "원재료비",
        "item_code": "material",
        "required_flag": "N",
        "enabled": "Y",
    },
    {
        "statement_type": "COST",
        "page_number": 1,
        "row_index": 6,
        "raw_account_name": "노무비",
        "normalized_account_code": "cost_labor",
        "normalized_account_name": "노무비",
        "item_code": "labor",
        "required_flag": "N",
        "enabled": "Y",
    },
    {
        "statement_type": "COST",
        "page_number": 1,
        "row_index": 10,
        "raw_account_name": "경비",
        "normalized_account_code": "cost_expense",
        "normalized_account_name": "경비",
        "item_code": "expense",
        "required_flag": "N",
        "enabled": "Y",
    },
    {
        "statement_type": "BS",
        "page_number": 1,
        "row_index": 2,
        "raw_account_name": "자산",
        "normalized_account_code": "bs_total_assets",
        "normalized_account_name": "자산",
        "item_code": "",
        "required_flag": "N",
        "enabled": "Y",
    },
    {
        "statement_type": "BS",
        "page_number": 1,
        "row_index": 3,
        "raw_account_name": "유동자산",
        "normalized_account_code": "bs_current_assets",
        "normalized_account_name": "유동자산",
        "item_code": "",
        "required_flag": "N",
        "enabled": "Y",
    },
    {
        "statement_type": "BS",
        "page_number": 1,
        "row_index": 4,
        "raw_account_name": "재고자산",
        "normalized_account_code": "bs_inventory",
        "normalized_account_name": "재고자산",
        "item_code": "",
        "required_flag": "N",
        "enabled": "N",
    },
    {
        "statement_type": "BS",
        "page_number": 1,
        "row_index": 5,
        "raw_account_name": "유형자산",
        "normalized_account_code": "bs_ppe",
        "normalized_account_name": "유형자산",
        "item_code": "",
        "required_flag": "N",
        "enabled": "N",
    },
    {
        "statement_type": "CF",
        "page_number": 1,
        "row_index": 2,
        "raw_account_name": "영업활동으로 인한 현금흐름",
        "normalized_account_code": "cf_operating_cash_flow",
        "normalized_account_name": "영업활동으로 인한 현금흐름",
        "item_code": "",
        "required_flag": "N",
        "enabled": "Y",
    },
    {
        "statement_type": "CF",
        "page_number": 2,
        "row_index": 4,
        "raw_account_name": "유형자산의 취득",
        "normalized_account_code": "cf_capex",
        "normalized_account_name": "유형자산의 취득",
        "item_code": "",
        "required_flag": "N",
        "enabled": "N",
    },
]

DEFAULT_PROXY_MIX_ROWS = [
    {"industry_group": "DEFAULT", "basis_type": "proxy", "material_share": 10.0, "labor_share": 45.0, "expense_share": 45.0, "note": "generic service fallback"},
    {"industry_group": "C", "basis_type": "proxy", "material_share": 72.1, "labor_share": 14.0, "expense_share": 13.9, "note": "manufacturing/material-heavy"},
    {"industry_group": "D", "basis_type": "proxy", "material_share": 15.0, "labor_share": 25.0, "expense_share": 60.0, "note": "utility/energy fallback"},
    {"industry_group": "E", "basis_type": "proxy", "material_share": 18.0, "labor_share": 28.0, "expense_share": 54.0, "note": "environment/service fallback"},
    {"industry_group": "F", "basis_type": "proxy", "material_share": 55.8, "labor_share": 23.3, "expense_share": 20.9, "note": "construction fallback"},
    {"industry_group": "G", "basis_type": "proxy", "material_share": 60.0, "labor_share": 15.0, "expense_share": 25.0, "note": "retail/wholesale fallback"},
    {"industry_group": "J", "basis_type": "proxy", "material_share": 5.5, "labor_share": 35.6, "expense_share": 58.9, "note": "information service fallback"},
    {"industry_group": "M", "basis_type": "proxy", "material_share": 7.3, "labor_share": 56.1, "expense_share": 36.6, "note": "professional service fallback"},
    {"industry_group": "N", "basis_type": "proxy", "material_share": 7.3, "labor_share": 56.1, "expense_share": 36.6, "note": "business support fallback"},
    {"industry_group": "O", "basis_type": "proxy", "material_share": 3.0, "labor_share": 67.0, "expense_share": 30.0, "note": "public administration fallback"},
    {"industry_group": "P", "basis_type": "proxy", "material_share": 2.0, "labor_share": 68.0, "expense_share": 30.0, "note": "education service fallback"},
    {"industry_group": "S", "basis_type": "proxy", "material_share": 8.0, "labor_share": 50.0, "expense_share": 42.0, "note": "personal service fallback"},
]

DEFAULT_OFFICIAL_REFERENCE_ROWS = [
    {"industry_group": "DEFAULT", "period_key": "DEFAULT", "cogs_pct": 60.0, "gna_pct": 28.0, "profit_pct": 5.0, "rnd_pct": 1.0, "material_share": 10.0, "labor_share": 45.0, "expense_share": 45.0, "source": "default_sector_anchor", "note": "editable fallback anchor"},
    {"industry_group": "C", "period_key": "DEFAULT", "cogs_pct": 84.0, "gna_pct": 11.0, "profit_pct": 4.0, "rnd_pct": 1.0, "material_share": 72.1, "labor_share": 14.0, "expense_share": 13.9, "source": "default_sector_anchor", "note": "manufacturing anchor"},
    {"industry_group": "D", "period_key": "DEFAULT", "cogs_pct": 78.0, "gna_pct": 9.0, "profit_pct": 10.0, "rnd_pct": 0.2, "material_share": 15.0, "labor_share": 25.0, "expense_share": 60.0, "source": "default_sector_anchor", "note": "utility anchor"},
    {"industry_group": "E", "period_key": "DEFAULT", "cogs_pct": 72.0, "gna_pct": 20.0, "profit_pct": 6.0, "rnd_pct": 0.3, "material_share": 18.0, "labor_share": 28.0, "expense_share": 54.0, "source": "default_sector_anchor", "note": "environment anchor"},
    {"industry_group": "F", "period_key": "DEFAULT", "cogs_pct": 87.0, "gna_pct": 9.0, "profit_pct": 3.0, "rnd_pct": 0.1, "material_share": 55.8, "labor_share": 23.3, "expense_share": 20.9, "source": "default_sector_anchor", "note": "construction anchor"},
    {"industry_group": "G", "period_key": "DEFAULT", "cogs_pct": 80.0, "gna_pct": 15.0, "profit_pct": 4.0, "rnd_pct": 0.3, "material_share": 60.0, "labor_share": 15.0, "expense_share": 25.0, "source": "default_sector_anchor", "note": "retail anchor"},
    {"industry_group": "J", "period_key": "DEFAULT", "cogs_pct": 38.0, "gna_pct": 47.0, "profit_pct": 8.0, "rnd_pct": 5.0, "material_share": 5.5, "labor_share": 35.6, "expense_share": 58.9, "source": "default_sector_anchor", "note": "information anchor"},
    {"industry_group": "M", "period_key": "DEFAULT", "cogs_pct": 45.0, "gna_pct": 45.0, "profit_pct": 6.0, "rnd_pct": 2.0, "material_share": 7.3, "labor_share": 56.1, "expense_share": 36.6, "source": "default_sector_anchor", "note": "professional service anchor"},
    {"industry_group": "N", "period_key": "DEFAULT", "cogs_pct": 56.0, "gna_pct": 34.0, "profit_pct": 5.0, "rnd_pct": 0.2, "material_share": 7.3, "labor_share": 56.1, "expense_share": 36.6, "source": "default_sector_anchor", "note": "business support anchor"},
    {"industry_group": "O", "period_key": "DEFAULT", "cogs_pct": 67.0, "gna_pct": 29.0, "profit_pct": 3.0, "rnd_pct": 0.2, "material_share": 3.0, "labor_share": 67.0, "expense_share": 30.0, "source": "default_sector_anchor", "note": "public administration anchor"},
    {"industry_group": "P", "period_key": "DEFAULT", "cogs_pct": 64.0, "gna_pct": 31.0, "profit_pct": 4.0, "rnd_pct": 0.5, "material_share": 2.0, "labor_share": 68.0, "expense_share": 30.0, "source": "default_sector_anchor", "note": "education service anchor"},
    {"industry_group": "S", "period_key": "DEFAULT", "cogs_pct": 58.0, "gna_pct": 30.0, "profit_pct": 6.0, "rnd_pct": 0.5, "material_share": 8.0, "labor_share": 50.0, "expense_share": 42.0, "source": "default_sector_anchor", "note": "personal service anchor"},
]

GENERATED_OFFICIAL_REFERENCE_SOURCE = "generated_official_standard_blend"
ESTIMATED_OFFICIAL_REFERENCE_SOURCE = "estimated_from_3y_growth"
OFFICIAL_REFERENCE_SERIES_SOURCE_NOTE = "Bank of Korea annual Financial Statement Analysis anchor blended with repaired workbook history and research-based triangulation."
OFFICIAL_REFERENCE_FIELD_LIMITS: Dict[str, Tuple[float, float]] = {
    "cogs_pct": (5.0, 95.0),
    "gna_pct": (0.0, 75.0),
    "profit_pct": (-10.0, 25.0),
    "rnd_pct": (0.0, 15.0),
    "material_share": (0.0, 100.0),
    "labor_share": (0.0, 100.0),
    "expense_share": (0.0, 100.0),
}
TRIANGULATION_RULE_PRIORITIES: Dict[str, float] = {
    "IS_COGS_TO_REVENUE": 0.00,
    "IS_SGA_TO_REVENUE": 0.00,
    "IS_OPERATING_PROFIT_TO_REVENUE": 0.00,
    "IS_RND_TO_REVENUE": 0.00,
    "COST_DIRECT_TO_REVENUE": 0.05,
    "REPAIRED_COGS_FROM_GROSS_PROFIT": 0.20,
    "REPAIRED_GNA_FROM_GROSS_PROFIT_AND_PROFIT": 0.20,
    "REPAIRED_PROFIT_FROM_GROSS_PROFIT_AND_GNA": 0.20,
    "REPAIRED_COGS_FROM_GNA_AND_PROFIT": 0.28,
    "REPAIRED_GNA_FROM_COGS_AND_PROFIT": 0.28,
    "REPAIRED_PROFIT_FROM_COGS_AND_GNA": 0.28,
    "REPAIRED_COGS_FROM_COST_SUM": 0.40,
    "COST_SCALED_TO_REPAIRED_COGS": 0.45,
    "OFFICIAL_REFERENCE_COGS": 0.72,
    "OFFICIAL_REFERENCE_GNA": 0.72,
    "OFFICIAL_REFERENCE_PROFIT": 0.72,
    "OFFICIAL_OR_PROXY_MIX_TO_REPAIRED_COGS": 0.82,
    "PROXY_MIX_FROM_COGS": 0.90,
    "ESTIMATED_FROM_3Y_GROWTH_BLEND": 0.95,
}
OUTLIER_SHRINK_THRESHOLDS: Dict[str, float] = {
    "material": 8.0,
    "labor": 7.0,
    "expense": 8.0,
    "gna": 8.0,
    "profit": 4.0,
    "rnd": 3.0,
}
OUTLIER_SHRINK_WEIGHTS: Dict[str, float] = {
    "observed": 0.90,
    "repaired": 0.72,
    "proxy": 0.58,
    "estimated": 0.48,
}
ACCOUNT_MAP_SUGGESTION_RULES: Dict[str, Dict[str, Any]] = {
    "is_material_nature": {
        "statement_type": "IS",
        "scan_pages": (2,),
        "keywords": ("원재료", "소모품", "상품의매입", "상품의 매입"),
    },
    "is_employee_benefits": {
        "statement_type": "IS",
        "scan_pages": (2,),
        "keywords": ("종업원급여", "급여", "인건비"),
    },
    "is_depreciation": {
        "statement_type": "IS",
        "scan_pages": (2,),
        "keywords": ("감가상각비",),
    },
    "bs_inventory": {
        "statement_type": "BS",
        "scan_pages": (1,),
        "keywords": ("재고자산",),
    },
    "bs_ppe": {
        "statement_type": "BS",
        "scan_pages": (1,),
        "keywords": ("유형자산",),
    },
    "cf_capex": {
        "statement_type": "CF",
        "scan_pages": (2, 1),
        "keywords": ("유형자산의취득", "유형자산취득"),
    },
}

DEFAULT_TAG_RULE_ROWS = [
    {
        "tag_code": "project_based",
        "tag_value": "true",
        "label": "프로젝트형",
        "keywords": "프로젝트|구축|고도화|수주|납품|turnkey|epc",
        "negative_keywords": "",
        "confidence_base": 0.74,
        "material_delta": -1.0,
        "labor_delta": -2.0,
        "expense_delta": 3.0,
        "gna_delta": -1.0,
        "profit_delta": 1.0,
        "rnd_delta": 0.0,
        "enabled": "Y",
    },
    {
        "tag_code": "onsite_installation",
        "tag_value": "true",
        "label": "현장설치형",
        "keywords": "현장|설치|시공|배관|전기공사|셋업|commissioning",
        "negative_keywords": "원격",
        "confidence_base": 0.76,
        "material_delta": 5.0,
        "labor_delta": 4.0,
        "expense_delta": 3.0,
        "gna_delta": -4.0,
        "profit_delta": -4.0,
        "rnd_delta": 0.0,
        "enabled": "Y",
    },
    {
        "tag_code": "maintenance_service",
        "tag_value": "true",
        "label": "유지보수형",
        "keywords": "유지보수|보수|유지관리|운영지원|상주운영|o&m|maintenance",
        "negative_keywords": "",
        "confidence_base": 0.76,
        "material_delta": -1.0,
        "labor_delta": 2.0,
        "expense_delta": 2.0,
        "gna_delta": -1.0,
        "profit_delta": -2.0,
        "rnd_delta": 0.0,
        "enabled": "Y",
    },
    {
        "tag_code": "outsourcing_level",
        "tag_value": "high",
        "label": "외주높음",
        "keywords": "외주|하도급|재위탁|협력사|파트너사|subcontract",
        "negative_keywords": "",
        "confidence_base": 0.78,
        "material_delta": 0.0,
        "labor_delta": -5.0,
        "expense_delta": 6.0,
        "gna_delta": -1.0,
        "profit_delta": -1.0,
        "rnd_delta": 0.0,
        "enabled": "Y",
    },
    {
        "tag_code": "outsourcing_level",
        "tag_value": "med",
        "label": "외주중간",
        "keywords": "외주 활용|협력사 수행|부분 위탁",
        "negative_keywords": "",
        "confidence_base": 0.70,
        "material_delta": 0.0,
        "labor_delta": -2.5,
        "expense_delta": 3.0,
        "gna_delta": -0.5,
        "profit_delta": 0.0,
        "rnd_delta": 0.0,
        "enabled": "Y",
    },
    {
        "tag_code": "outsourcing_level",
        "tag_value": "low",
        "label": "외주낮음",
        "keywords": "직접 수행|내재화|자체 인력|내부 인력",
        "negative_keywords": "",
        "confidence_base": 0.68,
        "material_delta": 0.0,
        "labor_delta": 3.0,
        "expense_delta": -3.0,
        "gna_delta": 0.0,
        "profit_delta": 0.0,
        "rnd_delta": 0.0,
        "enabled": "Y",
    },
    {
        "tag_code": "labor_intensive",
        "tag_value": "true",
        "label": "노무집약형",
        "keywords": "인력|상주|엔지니어|인건비|노무|기술지원|pm",
        "negative_keywords": "",
        "confidence_base": 0.75,
        "material_delta": -2.0,
        "labor_delta": 8.0,
        "expense_delta": -5.0,
        "gna_delta": -1.0,
        "profit_delta": 0.0,
        "rnd_delta": 0.0,
        "enabled": "Y",
    },
    {
        "tag_code": "asset_heavy",
        "tag_value": "true",
        "label": "자산집약형",
        "keywords": "설비|장비|발전설비|공장|유형자산|감가상각|capex",
        "negative_keywords": "",
        "confidence_base": 0.74,
        "material_delta": 6.0,
        "labor_delta": -2.0,
        "expense_delta": -3.0,
        "gna_delta": -1.0,
        "profit_delta": 0.0,
        "rnd_delta": 0.0,
        "enabled": "Y",
    },
    {
        "tag_code": "rnd_intensity",
        "tag_value": "high",
        "label": "R&D높음",
        "keywords": "연구개발|r&d|기술개발|실증|prototype|프로토타입|특허",
        "negative_keywords": "",
        "confidence_base": 0.80,
        "material_delta": 0.0,
        "labor_delta": -1.5,
        "expense_delta": -1.0,
        "gna_delta": -2.5,
        "profit_delta": -1.0,
        "rnd_delta": 6.0,
        "enabled": "Y",
    },
    {
        "tag_code": "rnd_intensity",
        "tag_value": "med",
        "label": "R&D중간",
        "keywords": "개발비|연구비|테스트베드|기술검증|poc",
        "negative_keywords": "",
        "confidence_base": 0.72,
        "material_delta": 0.0,
        "labor_delta": -1.0,
        "expense_delta": -0.5,
        "gna_delta": -1.5,
        "profit_delta": -0.5,
        "rnd_delta": 3.0,
        "enabled": "Y",
    },
    {
        "tag_code": "rnd_intensity",
        "tag_value": "low",
        "label": "R&D낮음",
        "keywords": "연구개발 비중 낮|단순 유지|반복 운영",
        "negative_keywords": "",
        "confidence_base": 0.66,
        "material_delta": 0.0,
        "labor_delta": 0.0,
        "expense_delta": 0.0,
        "gna_delta": 0.0,
        "profit_delta": 0.0,
        "rnd_delta": 0.0,
        "enabled": "Y",
    },
]

DEFAULT_RANGE_ROWS = [
    {"item_code": "material", "basis_type": "observed", "base_width_pctp": 3.0, "tag_cap_pctp": 6.0},
    {"item_code": "material", "basis_type": "repaired", "base_width_pctp": 3.8, "tag_cap_pctp": 6.0},
    {"item_code": "material", "basis_type": "proxy", "base_width_pctp": 4.0, "tag_cap_pctp": 6.0},
    {"item_code": "material", "basis_type": "estimated", "base_width_pctp": 5.0, "tag_cap_pctp": 6.0},
    {"item_code": "labor", "basis_type": "observed", "base_width_pctp": 5.0, "tag_cap_pctp": 8.0},
    {"item_code": "labor", "basis_type": "repaired", "base_width_pctp": 5.4, "tag_cap_pctp": 8.0},
    {"item_code": "labor", "basis_type": "proxy", "base_width_pctp": 5.8, "tag_cap_pctp": 8.0},
    {"item_code": "labor", "basis_type": "estimated", "base_width_pctp": 6.5, "tag_cap_pctp": 8.0},
    {"item_code": "expense", "basis_type": "observed", "base_width_pctp": 5.0, "tag_cap_pctp": 8.0},
    {"item_code": "expense", "basis_type": "repaired", "base_width_pctp": 5.4, "tag_cap_pctp": 8.0},
    {"item_code": "expense", "basis_type": "proxy", "base_width_pctp": 5.8, "tag_cap_pctp": 8.0},
    {"item_code": "expense", "basis_type": "estimated", "base_width_pctp": 6.5, "tag_cap_pctp": 8.0},
    {"item_code": "gna", "basis_type": "observed", "base_width_pctp": 1.8, "tag_cap_pctp": 3.0},
    {"item_code": "gna", "basis_type": "repaired", "base_width_pctp": 2.2, "tag_cap_pctp": 3.0},
    {"item_code": "gna", "basis_type": "proxy", "base_width_pctp": 2.4, "tag_cap_pctp": 3.0},
    {"item_code": "gna", "basis_type": "estimated", "base_width_pctp": 3.0, "tag_cap_pctp": 3.0},
    {"item_code": "profit", "basis_type": "observed", "base_width_pctp": 2.0, "tag_cap_pctp": 4.0},
    {"item_code": "profit", "basis_type": "repaired", "base_width_pctp": 2.4, "tag_cap_pctp": 4.0},
    {"item_code": "profit", "basis_type": "proxy", "base_width_pctp": 2.6, "tag_cap_pctp": 4.0},
    {"item_code": "profit", "basis_type": "estimated", "base_width_pctp": 3.2, "tag_cap_pctp": 4.0},
    {"item_code": "rnd", "basis_type": "observed", "base_width_pctp": 2.0, "tag_cap_pctp": 4.0},
    {"item_code": "rnd", "basis_type": "repaired", "base_width_pctp": 2.4, "tag_cap_pctp": 4.0},
    {"item_code": "rnd", "basis_type": "proxy", "base_width_pctp": 2.8, "tag_cap_pctp": 4.0},
    {"item_code": "rnd", "basis_type": "estimated", "base_width_pctp": 3.5, "tag_cap_pctp": 4.0},
]

DEFAULT_SEVERITY_ROWS = [
    {"confidence_grade": "A", "watch_gap_pctp": 0.0, "alert_gap_pctp": 1.5, "critical_gap_pctp": 4.0},
    {"confidence_grade": "B", "watch_gap_pctp": 0.0, "alert_gap_pctp": 2.0, "critical_gap_pctp": 5.0},
    {"confidence_grade": "C", "watch_gap_pctp": 0.0, "alert_gap_pctp": 3.0, "critical_gap_pctp": 6.5},
    {"confidence_grade": "D", "watch_gap_pctp": 0.0, "alert_gap_pctp": 4.0, "critical_gap_pctp": 8.0},
]

CONFIDENCE_GRADE_THRESHOLDS = (("A", 0.85), ("B", 0.70), ("C", 0.55))
CONFIDENCE_GRADE_LABELS = {
    "A": "높음",
    "B": "보통",
    "C": "낮음",
    "D": "매우 낮음",
}
CONFIDENCE_PENALTY_WIDTH = {"A": 0.0, "B": 0.5, "C": 1.0, "D": 1.5}
YEAR_RE = re.compile(r"20\d{2}")
NUMBER_RE = re.compile(r"-?\d[\d,]*")
_OCR_ENGINE: Optional[RapidOCR] = None


def resolve_workbook_path(workbook_path: Optional[str] = None) -> str:
    if workbook_path:
        return workbook_path
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_WORKBOOK_FILENAME)


def resolve_raw_dir_path(raw_dir: Optional[str] = None) -> str:
    if raw_dir:
        return raw_dir
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_RAW_DIRNAME)


def _utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _empty_sheet(sheet_name: str) -> pd.DataFrame:
    return pd.DataFrame(columns=SHEET_SCHEMAS[sheet_name])


def _clean_df(sheet_name: str, df: pd.DataFrame) -> pd.DataFrame:
    clean = df.copy()
    for col in SHEET_SCHEMAS[sheet_name]:
        if col not in clean.columns:
            clean[col] = pd.NA
    return clean[SHEET_SCHEMAS[sheet_name]]


def _merge_config_account_map_defaults(df: pd.DataFrame) -> pd.DataFrame:
    existing = _clean_df("config_account_map", df)
    default_df = pd.DataFrame(DEFAULT_ACCOUNT_MAP_ROWS, columns=SHEET_SCHEMAS["config_account_map"])
    existing_keys = {
        (str(row["statement_type"]), str(row["normalized_account_code"]))
        for _, row in existing.iterrows()
        if str(row.get("normalized_account_code", "")).strip()
    }
    missing_rows = [
        row
        for row in DEFAULT_ACCOUNT_MAP_ROWS
        if (str(row["statement_type"]), str(row["normalized_account_code"])) not in existing_keys
    ]
    if not missing_rows:
        return existing
    return _clean_df("config_account_map", pd.concat([existing, pd.DataFrame(missing_rows, columns=default_df.columns)], ignore_index=True))


def _default_sheet_data() -> Dict[str, pd.DataFrame]:
    return {
        "statement_rows": _empty_sheet("statement_rows"),
        "metric_observations": _empty_sheet("metric_observations"),
        "industry_baselines": _empty_sheet("industry_baselines"),
        "company_tag_evidence": _empty_sheet("company_tag_evidence"),
        "evaluation_runs": _empty_sheet("evaluation_runs"),
        "config_account_map": pd.DataFrame(DEFAULT_ACCOUNT_MAP_ROWS, columns=SHEET_SCHEMAS["config_account_map"]),
        "config_proxy_mix": pd.DataFrame(DEFAULT_PROXY_MIX_ROWS, columns=SHEET_SCHEMAS["config_proxy_mix"]),
        "config_official_reference": pd.DataFrame(DEFAULT_OFFICIAL_REFERENCE_ROWS, columns=SHEET_SCHEMAS["config_official_reference"]),
        "config_tag_rules": pd.DataFrame(DEFAULT_TAG_RULE_ROWS, columns=SHEET_SCHEMAS["config_tag_rules"]),
        "config_range": pd.DataFrame(DEFAULT_RANGE_ROWS, columns=SHEET_SCHEMAS["config_range"]),
        "config_severity": pd.DataFrame(DEFAULT_SEVERITY_ROWS, columns=SHEET_SCHEMAS["config_severity"]),
    }


def _write_all_sheets(workbook_path: str, sheet_map: Dict[str, pd.DataFrame]) -> None:
    root, ext = os.path.splitext(workbook_path)
    tmp_path = f"{root}.tmp{ext or '.xlsx'}"
    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
        for sheet_name in SHEET_SCHEMAS:
            df = _clean_df(sheet_name, sheet_map.get(sheet_name, _empty_sheet(sheet_name)))
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    os.replace(tmp_path, workbook_path)


def ensure_workbook(workbook_path: Optional[str] = None) -> str:
    resolved = resolve_workbook_path(workbook_path)
    if not os.path.exists(resolved):
        _write_all_sheets(resolved, _default_sheet_data())
        return resolved

    with pd.ExcelFile(resolved) as xls:
        missing = [sheet for sheet in SHEET_SCHEMAS if sheet not in xls.sheet_names]
    defaults = _default_sheet_data()
    rewrite_needed = bool(missing)
    sheets: Dict[str, pd.DataFrame] = {}
    with pd.ExcelFile(resolved) as xls:
        for sheet_name in SHEET_SCHEMAS:
            if sheet_name in xls.sheet_names:
                sheets[sheet_name] = _clean_df(sheet_name, pd.read_excel(resolved, sheet_name=sheet_name))
            else:
                sheets[sheet_name] = defaults[sheet_name]
    merged_account_map = _merge_config_account_map_defaults(sheets["config_account_map"])
    if len(merged_account_map) != len(sheets["config_account_map"]):
        sheets["config_account_map"] = merged_account_map
        rewrite_needed = True
    if rewrite_needed:
        _write_all_sheets(resolved, sheets)
    return resolved


def load_industry_db(workbook_path: Optional[str] = None) -> Dict[str, pd.DataFrame]:
    resolved = ensure_workbook(workbook_path)
    defaults = _default_sheet_data()
    sheets: Dict[str, pd.DataFrame] = {}
    with pd.ExcelFile(resolved) as xls:
        for sheet_name in SHEET_SCHEMAS:
            if sheet_name in xls.sheet_names:
                sheets[sheet_name] = _clean_df(sheet_name, pd.read_excel(resolved, sheet_name=sheet_name))
            else:
                sheets[sheet_name] = defaults[sheet_name]
    return sheets


def _hash_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def _hash_payload(payload: Any) -> str:
    encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _file_digest(file_path: str) -> str:
    with open(file_path, "rb") as handle:
        return _hash_bytes(handle.read())


def _parse_number(token: str) -> Optional[float]:
    if token is None:
        return None
    text = str(token).strip()
    if not NUMBER_RE.fullmatch(text):
        return None
    return float(text.replace(",", ""))


def _normalize_account_text(text: Any) -> str:
    cleaned = re.sub(r"\([^)]*\)", "", str(text or ""))
    cleaned = re.sub(r"[^0-9A-Za-z가-힣]+", "", cleaned)
    return cleaned.strip().lower()


def _coerce_excel_year(value: Any) -> Optional[str]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, (int, np.integer)):
        text = str(int(value))
    elif isinstance(value, (float, np.floating)) and float(value).is_integer():
        text = str(int(value))
    else:
        text = str(value).strip()
        if text.endswith(".0") and text[:-2].isdigit():
            text = text[:-2]
    return text if YEAR_RE.fullmatch(text) else None


def _coerce_excel_amount(value: Any) -> Optional[float]:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    text = str(value).strip().replace("%", "")
    if not text:
        return None
    if text.startswith("(") and text.endswith(")"):
        text = f"-{text[1:-1]}"
    return _parse_number(text)


def _detect_excel_statement_header(df: pd.DataFrame) -> Tuple[int, int, Dict[str, int]]:
    for row_idx in range(len(df)):
        row = df.iloc[row_idx].tolist()
        account_col_idx = None
        year_cols: Dict[str, int] = {}
        for col_idx, value in enumerate(row):
            normalized = _normalize_account_text(value)
            if normalized == "계정명":
                account_col_idx = col_idx
            year = _coerce_excel_year(value)
            if year:
                year_cols[year] = col_idx
        if account_col_idx is not None and len(year_cols) >= 2:
            ordered_year_cols = {year: year_cols[year] for year in sorted(year_cols)}
            return row_idx, account_col_idx, ordered_year_cols
    raise IndustryFinancialError("엑셀 재무제표에서 계정명/연도 헤더를 찾지 못했습니다.")


def _load_excel_statement_table(excel_path: str) -> Dict[str, Any]:
    df = pd.read_excel(excel_path, header=None)
    header_row_idx, account_col_idx, year_cols = _detect_excel_statement_header(df)
    rows: List[Dict[str, Any]] = []
    for row_idx in range(header_row_idx + 1, len(df)):
        row = df.iloc[row_idx]
        raw_account_name = row.iloc[account_col_idx] if account_col_idx < len(row) else None
        normalized_label = _normalize_account_text(raw_account_name)
        if not normalized_label or normalized_label == "계정명":
            continue
        values: Dict[str, float] = {}
        for year, col_idx in year_cols.items():
            if col_idx >= len(row):
                continue
            amount = _coerce_excel_amount(row.iloc[col_idx])
            if amount is not None:
                values[year] = amount
        if values:
            rows.append(
                {
                    "row_index": int(row_idx + 1),
                    "raw_account_name": str(raw_account_name),
                    "normalized_label": normalized_label,
                    "values": values,
                }
            )
    return {
        "sheet_name": "Sheet1",
        "header_row_index": int(header_row_idx + 1),
        "year_cols": year_cols,
        "rows": rows,
    }


def _account_metadata_by_statement(account_map: pd.DataFrame, statement_type: str) -> Dict[str, Dict[str, Any]]:
    filtered = account_map[account_map["statement_type"].astype(str) == str(statement_type)].copy()
    meta: Dict[str, Dict[str, Any]] = {}
    if not filtered.empty:
        filtered = filtered.drop_duplicates("normalized_account_code", keep="last")
        for _, row in filtered.iterrows():
            code = str(row["normalized_account_code"])
            meta[code] = row.to_dict()
    for default_row in DEFAULT_ACCOUNT_MAP_ROWS:
        if default_row["statement_type"] != statement_type:
            continue
        code = str(default_row["normalized_account_code"])
        meta.setdefault(code, default_row.copy())
    return meta


def _match_excel_account_row(
    rows: List[Dict[str, Any]],
    aliases: Tuple[str, ...],
) -> Optional[Dict[str, Any]]:
    best_row: Optional[Dict[str, Any]] = None
    best_score = -1.0
    normalized_aliases = [_normalize_account_text(alias) for alias in aliases if _normalize_account_text(alias)]
    for row in rows:
        label = str(row["normalized_label"])
        for alias in normalized_aliases:
            score = -1.0
            if label == alias:
                score = 100.0
            elif label.startswith(alias):
                score = 85.0 - max(len(label) - len(alias), 0) * 0.1
            elif alias in label:
                score = 70.0 - max(len(label) - len(alias), 0) * 0.1
            if score > best_score:
                best_score = score
                best_row = row
    return best_row if best_score >= 0 else None


def _map_excel_statement_rows(
    statement_type: str,
    excel_path: str,
    account_map: pd.DataFrame,
) -> List[Dict[str, Any]]:
    table = _load_excel_statement_table(excel_path)
    aliases = EXCEL_ACCOUNT_ALIASES.get(statement_type, {})
    metadata = _account_metadata_by_statement(account_map, statement_type)
    mapped_rows: List[Dict[str, Any]] = []
    for code, code_aliases in aliases.items():
        matched_row = _match_excel_account_row(table["rows"], code_aliases)
        if not matched_row:
            continue
        meta = metadata.get(code)
        if not meta:
            continue
        for year, amount in matched_row["values"].items():
            mapped_rows.append(
                {
                    "statement_type": statement_type,
                    "period_key": str(year),
                    "page_number": 1,
                    "row_index": int(matched_row["row_index"]),
                    "raw_account_name": matched_row["raw_account_name"],
                    "normalized_account_code": code,
                    "normalized_account_name": meta.get("normalized_account_name", code),
                    "item_code": meta.get("item_code", ""),
                    "amount": float(amount),
                }
            )
    return mapped_rows


def _ocr_engine() -> RapidOCR:
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        _OCR_ENGINE = RapidOCR()
    return _OCR_ENGINE


def _group_ocr_rows(
    image: np.ndarray,
    year_centers: Optional[Dict[str, float]] = None,
) -> Tuple[List[str], Dict[str, float], Dict[int, Dict[str, Optional[float]]]]:
    result, _ = _ocr_engine()(image)
    tokens: List[Tuple[float, float, str]] = []
    for box, text, _score in result or []:
        ys = [point[1] for point in box]
        xs = [point[0] for point in box]
        tokens.append((sum(ys) / len(ys), sum(xs) / len(xs), str(text)))
    tokens.sort(key=lambda row: (row[0], row[1]))

    grouped: List[List[Any]] = []
    for cy, cx, text in tokens:
        if not grouped or abs(grouped[-1][0] - cy) > 18:
            grouped.append([cy, [(cx, text)]])
        else:
            grouped[-1][1].append((cx, text))

    local_years: List[str] = []
    header_row_idx = -1
    if year_centers is None:
        for idx, (_cy, items) in enumerate(grouped):
            candidates = []
            for cx, text in items:
                match = YEAR_RE.search(text)
                if match:
                    candidates.append((match.group(0), cx))
            if len(candidates) >= 2:
                local_years = [year for year, _cx in candidates]
                year_centers = {year: cx for year, cx in candidates}
                header_row_idx = idx
                break
        if not year_centers:
            raise IndustryFinancialError("연도 헤더를 인식하지 못했습니다.")
    else:
        local_years = list(year_centers.keys())

    data_groups = grouped[header_row_idx + 1 :] if header_row_idx >= 0 else grouped
    row_values: Dict[int, Dict[str, Optional[float]]] = {}
    row_index = 0
    for _cy, items in data_groups:
        values = {year: None for year in local_years}
        for cx, text in items:
            number = _parse_number(text)
            if number is None:
                continue
            nearest_year = min(local_years, key=lambda year: abs(cx - year_centers[year]))
            if abs(cx - year_centers[nearest_year]) <= 260:
                values[nearest_year] = number
        if any(value is not None for value in values.values()):
            row_index += 1
            row_values[row_index] = values
    return local_years, year_centers, row_values


def _render_pdf_page(pdf_path: str, page_number: int) -> np.ndarray:
    pdf = pdfium.PdfDocument(pdf_path)
    try:
        page = pdf[page_number - 1]
    except IndexError as exc:
        raise IndustryFinancialError(f"{pdf_path}의 {page_number}페이지를 찾지 못했습니다.") from exc
    return np.array(page.render(scale=2.0).to_pil())


def _normalize_ocr_text(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).strip()


def _ocr_group_lines(image: np.ndarray, *, y_tolerance: int = 20) -> List[Dict[str, Any]]:
    result, _ = _ocr_engine()(image)
    tokens: List[Tuple[float, float, str]] = []
    for box, text, _score in result or []:
        ys = [point[1] for point in box]
        xs = [point[0] for point in box]
        tokens.append((sum(ys) / len(ys), sum(xs) / len(xs), str(text)))
    tokens.sort(key=lambda row: (row[0], row[1]))

    grouped: List[List[Any]] = []
    for cy, cx, text in tokens:
        if not grouped or abs(grouped[-1][0] - cy) > y_tolerance:
            grouped.append([cy, [(cx, text)]])
        else:
            grouped[-1][1].append((cx, text))

    lines: List[Dict[str, Any]] = []
    numeric_row_index = 0
    for line_index, (_cy, items) in enumerate(grouped, start=1):
        text = "".join(text for _cx, text in items)
        normalized = _normalize_ocr_text(text)
        has_number = any(_parse_number(token_text) is not None for _cx, token_text in items)
        if has_number:
            numeric_row_index += 1
        lines.append(
            {
                "line_index": line_index,
                "numeric_row_index": numeric_row_index if has_number else None,
                "text": text,
                "normalized_text": normalized,
            }
        )
    return lines


def suggest_account_map_candidates(
    raw_dir: Optional[str] = None,
    *,
    target_codes: Optional[List[str]] = None,
    max_samples_per_code: int = 24,
) -> Dict[str, Any]:
    resolved = resolve_raw_dir_path(raw_dir)
    if not os.path.isdir(resolved):
        raise IndustryFinancialError(f"raw PDF 폴더를 찾을 수 없습니다: {resolved}")

    selected_codes = set(target_codes or ACCOUNT_MAP_SUGGESTION_RULES.keys())
    rule_subset = {code: rule for code, rule in ACCOUNT_MAP_SUGGESTION_RULES.items() if code in selected_codes}
    if not rule_subset:
        return {"raw_dir": resolved, "targets": {}, "skipped": True}

    hits_map: Dict[str, List[Dict[str, Any]]] = {code: [] for code in rule_subset}
    for item in discover_financial_raw_sets(resolved):
        pdf_paths = item.get("pdf_paths", {})
        for code, rule in rule_subset.items():
            statement_type = str(rule["statement_type"])
            pdf_path = pdf_paths.get(statement_type)
            if not pdf_path or not os.path.exists(pdf_path):
                continue
            pdf = pdfium.PdfDocument(pdf_path)
            try:
                for page_number in rule["scan_pages"]:
                    if page_number < 1 or page_number > len(pdf):
                        continue
                    image = np.array(pdf[page_number - 1].render(scale=3.0).to_pil())
                    crops = [("left60", image[:, : max(1, int(image.shape[1] * 0.60))])]
                    for crop_name, crop_image in crops:
                        lines = _ocr_group_lines(crop_image, y_tolerance=20)
                        for line in lines:
                            normalized = line["normalized_text"]
                            if not normalized:
                                continue
                            matched_keywords = [keyword for keyword in rule["keywords"] if _normalize_ocr_text(keyword) in normalized]
                            if not matched_keywords:
                                continue
                            nearest_numeric = line["numeric_row_index"]
                            if nearest_numeric is None:
                                for later in lines[line["line_index"] : min(line["line_index"] + 3, len(lines))]:
                                    if later["numeric_row_index"] is not None:
                                        nearest_numeric = later["numeric_row_index"]
                                        break
                            score = len(matched_keywords) * 2.0
                            if crop_name != "full":
                                score += 0.25
                            if nearest_numeric is not None:
                                score += 0.75
                            hits_map[code].append(
                                {
                                    "industry_group": item["industry_group"],
                                    "source_file": os.path.basename(pdf_path),
                                    "statement_type": statement_type,
                                    "page_number": int(page_number),
                                    "line_index": int(line["line_index"]),
                                    "row_index": int(nearest_numeric) if nearest_numeric is not None else None,
                                    "crop": crop_name,
                                    "matched_keywords": matched_keywords,
                                    "score": round(score, 2),
                                    "text": line["text"][:200],
                                }
                            )
            finally:
                pdf.close()

    targets: Dict[str, Any] = {}
    for code, rule in rule_subset.items():
        hits = hits_map.get(code, [])
        scored = sorted(hits, key=lambda row: (-float(row["score"]), row["page_number"], row["row_index"] or 9999, row["source_file"]))
        counter = {}
        for row in scored:
            key = (row["page_number"], row["row_index"])
            if row["row_index"] is None:
                continue
            counter[key] = counter.get(key, 0) + 1
        top_candidates = [
            {"page_number": page_number, "row_index": row_index, "count": count}
            for (page_number, row_index), count in sorted(counter.items(), key=lambda item: (-item[1], item[0][0], item[0][1]))
        ]
        seed_row = next((row for row in DEFAULT_ACCOUNT_MAP_ROWS if row["normalized_account_code"] == code), None)
        targets[code] = {
            "statement_type": rule["statement_type"],
            "keywords": list(rule["keywords"]),
            "top_candidates": top_candidates[:8],
            "samples": scored[:max_samples_per_code],
            "default_seed": {
                "page_number": int(seed_row["page_number"]) if seed_row else None,
                "row_index": int(seed_row["row_index"]) if seed_row else None,
            },
            "hit_count": len(hits),
        }

    return {"raw_dir": resolved, "targets": targets, "skipped": False}


def _latest_statement_rows(statement_rows: pd.DataFrame, industry_group: str, period_key: str) -> pd.DataFrame:
    filtered = statement_rows[
        (statement_rows["industry_group"].astype(str) == str(industry_group))
        & (statement_rows["period_key"].astype(str) == str(period_key))
    ].copy()
    if filtered.empty:
        return filtered
    filtered["ingested_at_sort"] = pd.to_datetime(filtered["ingested_at"], errors="coerce")
    filtered = filtered.sort_values(["statement_type", "normalized_account_code", "ingested_at_sort"])
    deduped = filtered.drop_duplicates(["statement_type", "normalized_account_code"], keep="last")
    return deduped.drop(columns=["ingested_at_sort"])


def _grade_confidence(score: float) -> str:
    for grade, threshold in CONFIDENCE_GRADE_THRESHOLDS:
        if score >= threshold:
            return grade
    return "D"


def _clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def _normalize_ratios(values: Dict[str, float], active_items: Iterable[str]) -> Dict[str, float]:
    active = list(active_items)
    total = sum(max(values.get(item, 0.0), 0.0) for item in active)
    if total <= 0:
        even = 100.0 / max(len(active), 1)
        return {item: even for item in active}
    return {item: max(values.get(item, 0.0), 0.0) / total * 100.0 for item in active}


def _is_suspicious_cogs(cogs: float, revenue: float, cost_sum: float) -> bool:
    if revenue <= 0 or cost_sum <= 0:
        return False
    cogs_ratio_pct = abs(cogs) / revenue * 100.0
    cost_ratio_pct = abs(cost_sum) / revenue * 100.0
    return bool(cogs_ratio_pct < 1.0 and cost_ratio_pct >= 2.0 and abs(cost_sum) > max(abs(cogs) * 20.0, 1.0))


def _confidence_evaluation(
    basis_type: str,
    supporting_accounts: List[str],
    cost_absent: bool,
    restatement_seen: bool,
    *,
    official_aligned: bool = False,
    formula_repaired: bool = False,
    direct_cost_quality: str = "none",
    generated_reference: bool = False,
    triangulated: bool = False,
    cashflow_quality: str = "neutral",
) -> Dict[str, Any]:
    base_score = {"observed": 0.85, "repaired": 0.72, "proxy": 0.70, "estimated": 0.55}.get(basis_type, 0.55)
    score = base_score
    note_parts = [f"기본 {base_score:.2f}({basis_type})"]
    valid_supports = [str(account).strip() for account in supporting_accounts if str(account).strip()]

    if len(valid_supports) >= 2:
        score += 0.03
        note_parts.append(f"+ 근거계정 0.03({len(valid_supports)}개)")
    elif len(valid_supports) == 1:
        score += 0.02
        note_parts.append("+ 근거계정 0.02(1개)")

    if official_aligned and basis_type in {"repaired", "proxy", "estimated"}:
        score += 0.03
        note_parts.append("+ 공식정합 0.03")

    if formula_repaired and basis_type == "repaired":
        score += 0.03
        note_parts.append("+ 복원공식 0.03")

    if triangulated and basis_type in {"repaired", "proxy", "estimated"}:
        score += 0.02
        note_parts.append("+ 삼각검증 0.02")

    if direct_cost_quality == "usable" and basis_type in {"observed", "repaired"}:
        score += 0.03
        note_parts.append("+ 직접비정합 0.03")
    elif direct_cost_quality == "present" and basis_type in {"repaired", "proxy", "estimated"}:
        score += 0.01
        note_parts.append("+ 원가명세 0.01")

    if generated_reference and basis_type == "estimated":
        score += 0.02
        note_parts.append("+ 추정참조 0.02")

    if cashflow_quality == "strong" and basis_type in {"observed", "repaired"}:
        score += 0.01
        note_parts.append("+ 현금흐름정합 0.01")
    elif cashflow_quality == "weak":
        score -= 0.04
        note_parts.append("- 현금흐름괴리 0.04")

    if cost_absent and basis_type != "observed":
        score -= 0.10
        note_parts.append("- 직접원가근거 0.10")

    if restatement_seen:
        score -= 0.05
        note_parts.append("- 재작성표시 0.05")

    score = _clamp(score, 0.35, 0.95)
    grade = _grade_confidence(score)
    note = f"{' '.join(note_parts)} = {score:.2f} -> {CONFIDENCE_GRADE_LABELS.get(grade, grade)}({grade})"
    return {"score": score, "grade": grade, "note": note}


def _confidence_score(basis_type: str, supporting_accounts: List[str], cost_absent: bool, restatement_seen: bool) -> float:
    return float(
        _confidence_evaluation(
            basis_type=basis_type,
            supporting_accounts=supporting_accounts,
            cost_absent=cost_absent,
            restatement_seen=restatement_seen,
        )["score"]
    )


def _range_width(range_df: pd.DataFrame, item_code: str, basis_type: str, confidence_grade: str) -> float:
    match = range_df[
        (range_df["item_code"].astype(str) == item_code)
        & (range_df["basis_type"].astype(str) == basis_type)
    ]
    width = float(match.iloc[0]["base_width_pctp"]) if not match.empty else 5.0
    return width + CONFIDENCE_PENALTY_WIDTH.get(confidence_grade, 1.5)


def _basis_type_for_proxy(industry_group: str, proxy_df: pd.DataFrame) -> str:
    match = proxy_df[proxy_df["industry_group"].astype(str) == industry_group]
    if not match.empty:
        return str(match.iloc[0]["basis_type"])
    default_match = proxy_df[proxy_df["industry_group"].astype(str) == "DEFAULT"]
    if not default_match.empty:
        return str(default_match.iloc[0]["basis_type"])
    return "estimated"


def _proxy_mix(industry_group: str, proxy_df: pd.DataFrame) -> Dict[str, float]:
    match = proxy_df[proxy_df["industry_group"].astype(str) == industry_group]
    if match.empty:
        match = proxy_df[proxy_df["industry_group"].astype(str) == "DEFAULT"]
    if match.empty:
        return {"material": 10.0, "labor": 45.0, "expense": 45.0}
    row = match.iloc[0]
    return {
        "material": float(row["material_share"]),
        "labor": float(row["labor_share"]),
        "expense": float(row["expense_share"]),
    }


def _official_reference(industry_group: str, period_key: str, official_df: pd.DataFrame) -> Dict[str, float]:
    if official_df.empty:
        return {}

    normalized = official_df.copy()
    normalized["industry_group"] = normalized["industry_group"].astype(str)
    normalized["period_key"] = normalized["period_key"].astype(str)

    def _pick(match_df: pd.DataFrame) -> Optional[pd.Series]:
        if match_df.empty:
            return None
        working = match_df.copy()
        working["period_sort"] = pd.to_numeric(working["period_key"], errors="coerce")
        working = working.sort_values(["period_sort", "period_key"], na_position="last")
        return working.iloc[-1]

    by_group = normalized[normalized["industry_group"] == str(industry_group)]
    selected = _pick(by_group[by_group["period_key"] == str(period_key)])
    if selected is None:
        selected = _pick(by_group[by_group["period_key"].str.upper() == "DEFAULT"])
    if selected is None:
        default_group = normalized[normalized["industry_group"] == "DEFAULT"]
        selected = _pick(default_group[default_group["period_key"] == str(period_key)])
    if selected is None:
        default_group = normalized[normalized["industry_group"] == "DEFAULT"]
        selected = _pick(default_group[default_group["period_key"].str.upper() == "DEFAULT"])
    if selected is None:
        return {}

    def _to_float(name: str, default: float = 0.0) -> float:
        value = pd.to_numeric(pd.Series([selected.get(name)]), errors="coerce").iloc[0]
        if pd.isna(value):
            return default
        return float(value)

    return {
        "cogs_pct": _to_float("cogs_pct", 0.0),
        "gna_pct": _to_float("gna_pct", 0.0),
        "profit_pct": _to_float("profit_pct", 0.0),
        "rnd_pct": _to_float("rnd_pct", 0.0),
        "material_share": _to_float("material_share", 10.0),
        "labor_share": _to_float("labor_share", 45.0),
        "expense_share": _to_float("expense_share", 45.0),
        "source": str(selected.get("source", "") or ""),
        "note": str(selected.get("note", "") or ""),
    }


def _ratio_pct(amount: float, revenue: float, *, absolute: bool = True) -> float:
    if revenue <= 0:
        return 0.0
    value = abs(amount) if absolute else amount
    return float(value) / float(revenue) * 100.0


def _ratio_band(anchor: float, generic_low: float, generic_high: float, tolerance: float) -> Tuple[float, float]:
    if anchor <= 0:
        return generic_low, generic_high
    return max(generic_low, anchor - tolerance), min(generic_high, anchor + tolerance)


def _rule_priority(rule: str) -> float:
    cleaned = str(rule or "").strip()
    if not cleaned:
        return 1.0
    primary = cleaned.split("|", 1)[0].replace("_CLAMPED", "")
    return TRIANGULATION_RULE_PRIORITIES.get(primary, 0.85)


def _official_source_penalty(source: str) -> float:
    key = str(source or "").strip().lower()
    if not key:
        return 0.0
    if key.startswith("default_sector_anchor"):
        return 0.18
    if key.startswith(ESTIMATED_OFFICIAL_REFERENCE_SOURCE.lower()):
        return 0.10
    if key.startswith(GENERATED_OFFICIAL_REFERENCE_SOURCE.lower()):
        return 0.04
    return 0.0


def _choose_ratio_candidate(
    candidates: List[Tuple[Optional[float], str, List[str]]],
    low: float,
    high: float,
    *,
    allow_negative: bool = False,
    anchor: Optional[float] = None,
    official_source: str = "",
) -> Tuple[Optional[float], str, List[str]]:
    cleaned: List[Tuple[float, str, List[str]]] = []
    for value, rule, supporting_accounts in candidates:
        if value is None:
            continue
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            continue
        if not math.isfinite(numeric):
            continue
        if not allow_negative and numeric < 0:
            continue
        cleaned.append((numeric, rule, supporting_accounts))

    if not cleaned:
        return None, "", []

    anchor_penalty_weight = 0.08 if anchor is not None and math.isfinite(float(anchor)) else 0.0
    source_penalty = _official_source_penalty(official_source)

    def _score(row: Tuple[float, str, List[str]]) -> float:
        numeric, rule, supporting_accounts = row
        band_gap = 0.0 if low <= numeric <= high else min(abs(numeric - low), abs(numeric - high))
        anchor_gap = abs(numeric - float(anchor)) * anchor_penalty_weight if anchor_penalty_weight else 0.0
        support_bonus = min(len([x for x in supporting_accounts if str(x).strip()]), 3) * 0.08
        score = band_gap + anchor_gap + _rule_priority(rule) - support_bonus
        if str(rule).startswith("OFFICIAL_REFERENCE_"):
            score += source_penalty
        return score

    numeric, rule, supporting_accounts = min(cleaned, key=_score)
    clamped = _clamp(numeric, low, high)
    if math.isclose(clamped, numeric, rel_tol=0.0, abs_tol=1e-6):
        return numeric, rule, supporting_accounts
    return clamped, f"{rule}_CLAMPED", supporting_accounts


def _cashflow_quality_signal(
    revenue: float,
    cf_operating_cash_flow: float,
    profit_ratio: float,
    official_profit_ratio: Optional[float],
) -> Dict[str, Any]:
    if revenue <= 0 or not math.isfinite(float(cf_operating_cash_flow)):
        return {"status": "none", "note": "", "cfo_margin": None, "gap_pctp": None, "adjusted_profit_ratio": profit_ratio}

    cfo_margin = _ratio_pct(cf_operating_cash_flow, revenue, absolute=False)
    gap = abs(cfo_margin - profit_ratio)
    status = "neutral"
    if (profit_ratio > 0 and cfo_margin < 0) or gap >= 10.0:
        status = "weak"
    elif gap <= 3.0:
        status = "strong"

    adjusted_profit_ratio = float(profit_ratio)
    if status == "weak":
        benchmark = official_profit_ratio if official_profit_ratio is not None else None
        if benchmark is None or not math.isfinite(float(benchmark)):
            benchmark = max(min(cfo_margin, profit_ratio), 0.0)
        blend_weight = 0.75 if gap < 14.0 else 0.62
        adjusted_profit_ratio = _clamp((profit_ratio * blend_weight) + (float(benchmark) * (1.0 - blend_weight)), -15.0, 25.0)

    note = f"영업현금흐름 마진 {cfo_margin:.2f}%, 영업이익률 격차 {gap:.2f}%p"
    return {
        "status": status,
        "note": note,
        "cfo_margin": cfo_margin,
        "gap_pctp": gap,
        "adjusted_profit_ratio": adjusted_profit_ratio,
    }


def _apply_outlier_shrink(
    item_code: str,
    observed_pct: float,
    official_pct: Optional[float],
    basis_type: str,
    support_count: int,
    *,
    cost_present: bool,
    cost_usable: bool,
    restatement_seen: bool,
    official_source: str = "",
) -> Dict[str, Any]:
    if official_pct is None or not math.isfinite(float(official_pct)):
        return {"value_pct": observed_pct, "applied": False, "rule_suffix": "", "note": ""}

    threshold = OUTLIER_SHRINK_THRESHOLDS.get(item_code, 6.0)
    if item_code in {"material", "labor", "expense"} and cost_usable:
        threshold += 2.5
    gap = observed_pct - float(official_pct)
    if abs(gap) <= threshold:
        return {"value_pct": observed_pct, "applied": False, "rule_suffix": "", "note": ""}

    observed_weight = OUTLIER_SHRINK_WEIGHTS.get(basis_type, 0.55)
    if support_count >= 2:
        observed_weight += 0.03
    elif support_count == 0:
        observed_weight -= 0.04
    if item_code in {"material", "labor", "expense"}:
        if cost_usable:
            observed_weight += 0.04
        elif not cost_present:
            observed_weight -= 0.08
    if restatement_seen:
        observed_weight -= 0.05
    observed_weight += _official_source_penalty(official_source)
    observed_weight = _clamp(observed_weight, 0.38, 0.94)
    value_pct = (observed_pct * observed_weight) + (float(official_pct) * (1.0 - observed_weight))
    note = (
        f"공식기준 {official_pct:.2f}% 대비 괴리 {gap:+.2f}%p로 "
        f"{(1.0 - observed_weight) * 100.0:.0f}% 비중만큼 축소 보정"
    )
    return {
        "value_pct": value_pct,
        "applied": True,
        "rule_suffix": "OUTLIER_SHRINK_TO_OFFICIAL",
        "note": note,
    }


def _scaled_direct_cost_ratios(
    revenue: float,
    target_cogs_ratio: float,
    material_amount: float,
    labor_amount: float,
    expense_amount: float,
    mix: Dict[str, float],
) -> Dict[str, float]:
    if revenue <= 0 or target_cogs_ratio <= 0:
        return {"material": 0.0, "labor": 0.0, "expense": 0.0}

    direct_sum = max(material_amount, 0.0) + max(labor_amount, 0.0) + max(expense_amount, 0.0)
    if direct_sum > 0:
        material_share = max(material_amount, 0.0) / direct_sum * 100.0
        labor_share = max(labor_amount, 0.0) / direct_sum * 100.0
        expense_share = max(expense_amount, 0.0) / direct_sum * 100.0
    else:
        material_share = float(mix.get("material", 10.0))
        labor_share = float(mix.get("labor", 45.0))
        expense_share = float(mix.get("expense", 45.0))

    return {
        "material": target_cogs_ratio * material_share / 100.0,
        "labor": target_cogs_ratio * labor_share / 100.0,
        "expense": target_cogs_ratio * expense_share / 100.0,
    }


def _latest_metric_rows(metric_rows: pd.DataFrame, industry_group: str, period_key: str) -> pd.DataFrame:
    filtered = metric_rows[
        (metric_rows["industry_group"].astype(str) == str(industry_group))
        & (metric_rows["period_key"].astype(str) == str(period_key))
    ].copy()
    if filtered.empty:
        return filtered
    filtered["created_at_sort"] = pd.to_datetime(filtered["created_at"], errors="coerce")
    filtered = filtered.sort_values(["metric_code", "created_at_sort"]).drop_duplicates(["metric_code"], keep="last")
    return filtered.drop(columns=["created_at_sort"])


def _safe_float(value: Any, default: float = 0.0) -> float:
    numeric = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.isna(numeric):
        return default
    return float(numeric)


def _previous_numeric_period_key(period_key: str) -> Optional[str]:
    text = str(period_key).strip()
    if not YEAR_RE.fullmatch(text):
        return None
    return str(int(text) - 1)


def _latest_statement_value(
    statement_rows: pd.DataFrame,
    industry_group: str,
    period_key: str,
    normalized_account_code: str,
) -> Optional[float]:
    filtered = _latest_statement_rows(statement_rows, industry_group, str(period_key))
    if filtered.empty:
        return None
    match = filtered[filtered["normalized_account_code"].astype(str) == str(normalized_account_code)]
    if match.empty:
        return None
    numeric = pd.to_numeric(pd.Series([match.iloc[-1]["amount"]]), errors="coerce").iloc[0]
    if pd.isna(numeric):
        return None
    return float(numeric)


def _normalize_share_triplet(material_share: float, labor_share: float, expense_share: float) -> Dict[str, float]:
    total = max(material_share, 0.0) + max(labor_share, 0.0) + max(expense_share, 0.0)
    if total <= 0:
        return {"material_share": 10.0, "labor_share": 45.0, "expense_share": 45.0}
    return {
        "material_share": max(material_share, 0.0) / total * 100.0,
        "labor_share": max(labor_share, 0.0) / total * 100.0,
        "expense_share": max(expense_share, 0.0) / total * 100.0,
    }


def _rebalance_direct_ratios(
    direct_values: Dict[str, float],
    target_total: float,
    *,
    locked_items: Optional[Iterable[str]] = None,
) -> Dict[str, float]:
    working = {item: max(float(value), 0.0) for item, value in direct_values.items()}
    total = sum(working.values())
    if total <= 0 or target_total <= 0:
        return working
    if total <= target_total + 1e-6:
        return working

    locked = {str(item) for item in (locked_items or []) if str(item)}
    locked_total = sum(working.get(item, 0.0) for item in locked)
    if locked_total >= target_total:
        scale = target_total / max(locked_total, 1e-9)
        return {item: working.get(item, 0.0) * scale if item in locked else 0.0 for item in working}

    unlocked = [item for item in working if item not in locked]
    unlocked_total = sum(working[item] for item in unlocked)
    if unlocked_total <= 0:
        scale = target_total / total
        return {item: value * scale for item, value in working.items()}

    scale = (target_total - locked_total) / unlocked_total
    for item in unlocked:
        working[item] = working[item] * scale
    return working


def _metric_value(metric_rows: pd.DataFrame, item_code: str) -> Optional[float]:
    match = metric_rows[metric_rows["item_code"].astype(str) == item_code]
    if match.empty:
        return None
    return _safe_float(match.iloc[-1]["value_pct"], default=0.0)


def _blend_reference_value(anchor: float, observed: Optional[float], history_weight: float, field_name: str) -> float:
    low, high = OFFICIAL_REFERENCE_FIELD_LIMITS[field_name]
    if observed is None or not math.isfinite(float(observed)):
        return _clamp(anchor, low, high)
    return _clamp((float(observed) * history_weight) + (float(anchor) * (1.0 - history_weight)), low, high)


def _project_next_value(series: List[float], field_name: str) -> float:
    low, high = OFFICIAL_REFERENCE_FIELD_LIMITS[field_name]
    valid = [float(value) for value in series if value is not None and math.isfinite(float(value))]
    if not valid:
        return low
    if len(valid) == 1:
        return _clamp(valid[-1], low, high)

    if len(valid) >= 3 and min(valid) > 0.1:
        base = valid[0]
        last = valid[-1]
        growth_factor = (last / base) ** (1.0 / max(len(valid) - 1, 1)) if base > 0 else 1.0
        growth_factor = _clamp(growth_factor, 0.85, 1.15)
        projected = last * growth_factor
    else:
        diffs = [curr - prev for prev, curr in zip(valid, valid[1:])]
        avg_diff = sum(diffs) / len(diffs) if diffs else 0.0
        max_step = max((high - low) * 0.12, 1.0)
        projected = valid[-1] + _clamp(avg_diff, -max_step, max_step)

    return _clamp(projected, low, high)


def _estimated_reference_row(
    industry_group: str,
    period_key: str,
    historical_rows: List[Dict[str, Any]],
) -> Dict[str, Any]:
    fields = [
        "cogs_pct",
        "gna_pct",
        "profit_pct",
        "rnd_pct",
        "material_share",
        "labor_share",
        "expense_share",
    ]
    projected = {
        field_name: _project_next_value([_safe_float(row.get(field_name), 0.0) for row in historical_rows], field_name)
        for field_name in fields
    }
    normalized_shares = _normalize_share_triplet(
        projected["material_share"],
        projected["labor_share"],
        projected["expense_share"],
    )
    projected.update(normalized_shares)
    return {
        "industry_group": str(industry_group),
        "period_key": str(period_key),
        "cogs_pct": round(projected["cogs_pct"], 6),
        "gna_pct": round(projected["gna_pct"], 6),
        "profit_pct": round(projected["profit_pct"], 6),
        "rnd_pct": round(projected["rnd_pct"], 6),
        "material_share": round(projected["material_share"], 6),
        "labor_share": round(projected["labor_share"], 6),
        "expense_share": round(projected["expense_share"], 6),
        "source": ESTIMATED_OFFICIAL_REFERENCE_SOURCE,
        "note": "Projected from recent 3-year official-reference standard series using CAGR/average-delta logic.",
    }


def _overwrite_period_rows(df: pd.DataFrame, industry_group: str, period_key: str) -> pd.DataFrame:
    return df[
        ~(
            (df["industry_group"].astype(str) == str(industry_group))
            & (df["period_key"].astype(str) == str(period_key))
        )
    ].copy()


def _tag_label(tag_code: str, tag_value: str) -> str:
    return tag_code if tag_value in {"", "true"} else f"{tag_code}:{tag_value}"


def _append_rows(existing: pd.DataFrame, rows: List[Dict[str, Any]], sheet_name: str) -> pd.DataFrame:
    if not rows:
        return _clean_df(sheet_name, existing)
    incoming = pd.DataFrame(rows, columns=SHEET_SCHEMAS[sheet_name])
    if existing.empty:
        return _clean_df(sheet_name, incoming)
    return _clean_df(sheet_name, pd.concat([existing, incoming], ignore_index=True))


def reset_industry_db(workbook_path: Optional[str] = None) -> str:
    resolved = resolve_workbook_path(workbook_path)
    _write_all_sheets(resolved, _default_sheet_data())
    return resolved


def _statement_file_priority(path: str) -> int:
    ext = os.path.splitext(str(path))[1].lower()
    if ext in EXCEL_STATEMENT_EXTENSIONS:
        return 2
    if ext == ".pdf":
        return 1
    return 0


def _normalize_extension_filter(file_mode: str) -> Optional[set[str]]:
    mode = str(file_mode or "auto").strip().lower()
    if mode == "auto":
        return None
    if mode == "xlsx":
        return set(EXCEL_STATEMENT_EXTENSIONS)
    if mode == "pdf":
        return {".pdf"}
    raise IndustryFinancialError(f"지원하지 않는 file_mode입니다: {file_mode}")


def discover_financial_raw_sets(raw_dir: Optional[str] = None, *, file_mode: str = "auto") -> List[Dict[str, Any]]:
    resolved = resolve_raw_dir_path(raw_dir)
    if not os.path.isdir(resolved):
        raise IndustryFinancialError(f"raw 재무제표 폴더를 찾을 수 없습니다: {resolved}")
    allowed_extensions = _normalize_extension_filter(file_mode)

    grouped: Dict[str, Dict[str, Any]] = {}
    unmatched_files: List[str] = []
    duplicate_files: List[str] = []

    for entry in sorted(os.scandir(resolved), key=lambda item: item.name):
        if not entry.is_file():
            continue
        file_ext = os.path.splitext(entry.name)[1].lower()
        if file_ext != ".pdf" and file_ext not in EXCEL_STATEMENT_EXTENSIONS:
            continue
        if allowed_extensions is not None and file_ext not in allowed_extensions:
            continue
        match = RAW_FILENAME_RE.match(entry.name)
        if not match:
            unmatched_files.append(entry.name)
            continue

        industry_group = match.group("industry_group").upper()
        industry_label = match.group("industry_label").strip()
        statement_label = match.group("statement_label").strip().replace(" ", "")
        statement_type = STATEMENT_LABEL_TO_TYPE.get(statement_label)
        if statement_type is None:
            unmatched_files.append(entry.name)
            continue

        bucket = grouped.setdefault(
            industry_group,
            {
                "industry_group": industry_group,
                "industry_label": industry_label,
                "pdf_paths": {},
            },
        )
        if bucket["industry_label"] != industry_label:
            raise IndustryFinancialError(
                f"동일 업종 코드({industry_group})에 서로 다른 업종명이 섞여 있습니다: "
                f"{bucket['industry_label']} / {industry_label}"
            )
        previous_path = bucket["pdf_paths"].get(statement_type)
        if previous_path:
            previous_priority = _statement_file_priority(previous_path)
            current_priority = _statement_file_priority(entry.path)
            if current_priority > previous_priority:
                duplicate_files.append(os.path.basename(previous_path))
                bucket["pdf_paths"][statement_type] = entry.path
            else:
                duplicate_files.append(entry.name)
            continue
        bucket["pdf_paths"][statement_type] = entry.path

    discovered = [grouped[key] for key in sorted(grouped)]
    for item in discovered:
        item["missing_required"] = [
            statement_type
            for statement_type in ("BS", "IS", "CF")
            if statement_type not in item["pdf_paths"]
        ]
    if unmatched_files:
        for item in discovered:
            item.setdefault("unmatched_files", unmatched_files)
    if duplicate_files:
        for item in discovered:
            item.setdefault("duplicate_files", duplicate_files)
    return discovered


def ingest_financial_raw_directory(
    period_label: str = DEFAULT_BATCH_PERIOD_LABEL,
    raw_dir: Optional[str] = None,
    workbook_path: Optional[str] = None,
    reset_workbook: bool = False,
    estimate_period: str = "2025",
    file_mode: str = "auto",
) -> Dict[str, Any]:
    resolved_workbook = reset_industry_db(workbook_path) if reset_workbook else ensure_workbook(workbook_path)
    discovered = discover_financial_raw_sets(raw_dir, file_mode=file_mode)
    results: List[Dict[str, Any]] = []
    errors: List[Dict[str, str]] = []
    skipped: List[Dict[str, Any]] = []

    for item in discovered:
        missing_required = item.get("missing_required", [])
        if missing_required:
            skipped.append(
                {
                    "industry_group": item["industry_group"],
                    "industry_label": item["industry_label"],
                    "reason": f"missing required statements: {', '.join(missing_required)}",
                }
            )
            continue

        try:
            ingest_result = ingest_industry_statements(
                item["industry_group"],
                period_label,
                item["pdf_paths"],
                workbook_path=resolved_workbook,
            )
            results.append(
                {
                    "industry_group": item["industry_group"],
                    "industry_label": item["industry_label"],
                    "period_keys": ingest_result["period_keys"],
                    "rows_written": ingest_result["rows_written"],
                    "noop": bool(ingest_result.get("noop")),
                    "restatement_detected": bool(ingest_result.get("restatement_detected")),
                }
            )
        except Exception as exc:
            errors.append(
                {
                    "industry_group": item["industry_group"],
                    "industry_label": item["industry_label"],
                    "error": str(exc),
                }
            )

    reference_refresh = refresh_official_reference_sheet(resolved_workbook, estimate_period=str(estimate_period))
    estimated_baselines = refresh_estimated_baselines(resolved_workbook, target_period=str(estimate_period))
    available_baselines = list_available_baselines(resolved_workbook)
    baseline_groups = (
        sorted(available_baselines["industry_group"].astype(str).unique().tolist())
        if not available_baselines.empty
        else []
    )

    return {
        "raw_dir": resolve_raw_dir_path(raw_dir),
        "workbook_path": resolved_workbook,
        "period_label": period_label,
        "reset_workbook": bool(reset_workbook),
        "discovered_groups": [item["industry_group"] for item in discovered],
        "unmatched_files": sorted({name for item in discovered for name in item.get("unmatched_files", [])}),
        "duplicate_files": sorted({name for item in discovered for name in item.get("duplicate_files", [])}),
        "imported_groups": [item["industry_group"] for item in results],
        "baseline_groups": baseline_groups,
        "estimate_period": str(estimate_period),
        "file_mode": str(file_mode),
        "reference_refresh": reference_refresh,
        "estimated_baselines": estimated_baselines,
        "results": results,
        "skipped": skipped,
        "errors": errors,
    }


def ingest_industry_statements(
    industry_group: str,
    period_label: str,
    pdf_paths: Dict[str, str],
    workbook_path: Optional[str] = None,
) -> Dict[str, Any]:
    resolved = ensure_workbook(workbook_path)
    sheets = load_industry_db(resolved)
    account_map = sheets["config_account_map"]

    required = {"BS", "IS", "CF"}
    missing_required = sorted(statement for statement in required if statement not in pdf_paths or not pdf_paths[statement])
    if missing_required:
        raise IndustryFinancialError(f"필수 PDF가 누락되었습니다: {', '.join(missing_required)}")

    source_digests = {statement: _file_digest(path) for statement, path in pdf_paths.items() if path}
    ingestion_digest = _hash_payload(
        {
            "industry_group": industry_group,
            "period_label": period_label,
            "source_digests": source_digests,
        }
    )

    existing = sheets["statement_rows"]
    same_digest = existing[
        (existing["industry_group"].astype(str) == str(industry_group))
        & (existing["inputs_digest"].astype(str) == ingestion_digest)
    ]
    if not same_digest.empty:
        return {
            "industry_group": industry_group,
            "inputs_digest": ingestion_digest,
            "period_keys": sorted(same_digest["period_key"].astype(str).unique().tolist()),
            "rows_written": 0,
            "noop": True,
            "restatement_detected": bool(same_digest["restatement_flag"].astype(str).str.lower().eq("true").any()),
        }

    mapped_rows: List[Dict[str, Any]] = []
    required_found = {"BS": False, "IS": False, "CF": False}
    now = _utc_now()

    for statement_type, pdf_path in pdf_paths.items():
        if not pdf_path:
            continue
        file_ext = os.path.splitext(str(pdf_path))[1].lower()
        statement_config = account_map[
            account_map["statement_type"].astype(str).eq(statement_type)
            & account_map["enabled"].astype(str).str.upper().eq("Y")
        ]
        if file_ext in EXCEL_STATEMENT_EXTENSIONS:
            excel_rows = _map_excel_statement_rows(statement_type, pdf_path, account_map)
            if excel_rows:
                required_found[statement_type] = True
            for excel_row in excel_rows:
                previous = existing[
                    (existing["industry_group"].astype(str) == str(industry_group))
                    & (existing["statement_type"].astype(str) == statement_type)
                    & (existing["period_key"].astype(str) == str(excel_row["period_key"]))
                    & (existing["normalized_account_code"].astype(str) == str(excel_row["normalized_account_code"]))
                ].copy()
                previous["ingested_at_sort"] = pd.to_datetime(previous["ingested_at"], errors="coerce")
                previous = previous.sort_values("ingested_at_sort")
                prev_amount = float(previous.iloc[-1]["amount"]) if not previous.empty else None
                restatement = prev_amount is not None and not math.isclose(prev_amount, excel_row["amount"], rel_tol=0.0, abs_tol=0.5)
                mapped_rows.append(
                    {
                        "industry_group": industry_group,
                        "statement_type": statement_type,
                        "period_key": str(excel_row["period_key"]),
                        "page_number": int(excel_row["page_number"]),
                        "row_index": int(excel_row["row_index"]),
                        "raw_account_name": excel_row["raw_account_name"],
                        "normalized_account_code": excel_row["normalized_account_code"],
                        "normalized_account_name": excel_row["normalized_account_name"],
                        "item_code": excel_row["item_code"],
                        "amount": float(excel_row["amount"]),
                        "source_file": os.path.basename(pdf_path),
                        "source_digest": source_digests[statement_type],
                        "inputs_digest": ingestion_digest,
                        "restatement_flag": bool(restatement),
                        "ingested_at": now,
                    }
                )
            continue
        if statement_config.empty:
            continue

        years: List[str] = []
        year_centers: Optional[Dict[str, float]] = None
        row_cache: Dict[int, Dict[int, Dict[str, Optional[float]]]] = {}
        for page_number in sorted(statement_config["page_number"].astype(int).unique()):
            image = _render_pdf_page(pdf_path, page_number)
            years, year_centers, page_rows = _group_ocr_rows(image, year_centers)
            row_cache[page_number] = page_rows

        for _, config_row in statement_config.iterrows():
            page_number = int(config_row["page_number"])
            row_index = int(config_row["row_index"])
            values = row_cache.get(page_number, {}).get(row_index)
            if not values:
                continue
            if values and any(value is not None for value in values.values()):
                required_found[statement_type] = True or required_found.get(statement_type, False)
            for year in years:
                amount = values.get(year)
                if amount is None:
                    continue
                previous = existing[
                    (existing["industry_group"].astype(str) == str(industry_group))
                    & (existing["statement_type"].astype(str) == statement_type)
                    & (existing["period_key"].astype(str) == str(year))
                    & (existing["normalized_account_code"].astype(str) == str(config_row["normalized_account_code"]))
                ].copy()
                previous["ingested_at_sort"] = pd.to_datetime(previous["ingested_at"], errors="coerce")
                previous = previous.sort_values("ingested_at_sort")
                prev_amount = float(previous.iloc[-1]["amount"]) if not previous.empty else None
                restatement = prev_amount is not None and not math.isclose(prev_amount, amount, rel_tol=0.0, abs_tol=0.5)
                mapped_rows.append(
                    {
                        "industry_group": industry_group,
                        "statement_type": statement_type,
                        "period_key": str(year),
                        "page_number": page_number,
                        "row_index": row_index,
                        "raw_account_name": config_row["raw_account_name"],
                        "normalized_account_code": config_row["normalized_account_code"],
                        "normalized_account_name": config_row["normalized_account_name"],
                        "item_code": config_row["item_code"],
                        "amount": amount,
                        "source_file": os.path.basename(pdf_path),
                        "source_digest": source_digests[statement_type],
                        "inputs_digest": ingestion_digest,
                        "restatement_flag": bool(restatement),
                        "ingested_at": now,
                    }
                )

    if not required_found["IS"]:
        raise IndustryFinancialError("손익계산서 핵심 행을 추출하지 못했습니다.")
    if not required_found["BS"]:
        raise IndustryFinancialError("재무상태표 핵심 행을 추출하지 못했습니다.")
    if not required_found["CF"]:
        raise IndustryFinancialError("현금흐름표 핵심 행을 추출하지 못했습니다.")
    if not mapped_rows:
        raise IndustryFinancialError("산업 4표에서 추출된 행이 없습니다.")

    sheets["statement_rows"] = _append_rows(existing, mapped_rows, "statement_rows")
    _write_all_sheets(resolved, sheets)

    period_keys = sorted({row["period_key"] for row in mapped_rows})
    restatement_detected = any(row["restatement_flag"] for row in mapped_rows)
    for key in period_keys:
        build_industry_baseline(industry_group, key, workbook_path=resolved)

    return {
        "industry_group": industry_group,
        "inputs_digest": ingestion_digest,
        "period_keys": period_keys,
        "rows_written": len(mapped_rows),
        "noop": False,
        "restatement_detected": restatement_detected,
        "period_label": period_label,
    }


def refresh_official_reference_sheet(
    workbook_path: Optional[str] = None,
    estimate_period: str = "2025",
) -> Dict[str, Any]:
    resolved = ensure_workbook(workbook_path)
    sheets = load_industry_db(resolved)
    statement_df = sheets["statement_rows"].copy()
    official_df = sheets["config_official_reference"].copy()
    proxy_df = sheets["config_proxy_mix"].copy()

    numeric_periods = sorted(
        {
            str(value)
            for value in statement_df["period_key"].astype(str).tolist()
            if YEAR_RE.fullmatch(str(value))
        }
    )
    history_periods = numeric_periods[-3:]
    industry_groups = sorted(
        {
            str(value)
            for value in statement_df["industry_group"].astype(str).tolist()
            if str(value).strip()
        }
        | {
            str(value)
            for value in proxy_df["industry_group"].astype(str).tolist()
            if str(value).strip() and str(value).upper() != "DEFAULT"
        }
    )

    if not history_periods or not industry_groups:
        return {
            "workbook_path": resolved,
            "history_periods": history_periods,
            "estimate_period": str(estimate_period),
            "refreshed_rows": 0,
            "estimated_rows": 0,
            "skipped": True,
        }

    preserved = official_df[
        ~(
            official_df["period_key"].astype(str).isin(history_periods + [str(estimate_period)])
            & official_df["industry_group"].astype(str).isin(industry_groups)
        )
    ].copy()

    refreshed_rows: List[Dict[str, Any]] = []
    estimated_rows: List[Dict[str, Any]] = []
    build_notes: List[Dict[str, Any]] = []

    for industry_group in industry_groups:
        anchor = _official_reference(industry_group, "DEFAULT", official_df)
        group_history: List[Dict[str, Any]] = []
        for period_key in history_periods:
            try:
                baseline = build_industry_baseline(industry_group, period_key, workbook_path=resolved)
            except Exception as exc:
                build_notes.append(
                    {
                        "industry_group": industry_group,
                        "period_key": str(period_key),
                        "status": "skipped",
                        "reason": str(exc),
                    }
                )
                continue

            metric_rows = pd.DataFrame(baseline.get("metric_rows", []))
            observed_cogs = sum(
                max(_metric_value(metric_rows, item_code) or 0.0, 0.0)
                for item_code in ("material", "labor", "expense")
            )
            observed_gna = _metric_value(metric_rows, "gna")
            observed_profit = _metric_value(metric_rows, "profit")
            observed_rnd = _metric_value(metric_rows, "rnd") or 0.0

            direct_total = sum(
                max(_metric_value(metric_rows, item_code) or 0.0, 0.0)
                for item_code in ("material", "labor", "expense")
            )
            if direct_total > 0:
                observed_material_share = max(_metric_value(metric_rows, "material") or 0.0, 0.0) / direct_total * 100.0
                observed_labor_share = max(_metric_value(metric_rows, "labor") or 0.0, 0.0) / direct_total * 100.0
                observed_expense_share = max(_metric_value(metric_rows, "expense") or 0.0, 0.0) / direct_total * 100.0
            else:
                observed_material_share = None
                observed_labor_share = None
                observed_expense_share = None

            suspicious_count = sum(
                [
                    bool(baseline.get("suspicious_cogs")),
                    bool(baseline.get("suspicious_gna")),
                    bool(baseline.get("suspicious_profit")),
                ]
            )
            history_weight = 0.65 if suspicious_count == 0 else 0.45 if suspicious_count == 1 else 0.30
            if not baseline.get("cost_present"):
                history_weight = min(history_weight, 0.40)
            direct_weight = 0.75 if baseline.get("cost_usable") else 0.55 if baseline.get("cost_present") else 0.35

            blended_cogs = _blend_reference_value(_safe_float(anchor.get("cogs_pct"), 60.0), observed_cogs, history_weight, "cogs_pct")
            blended_gna = _blend_reference_value(_safe_float(anchor.get("gna_pct"), 28.0), observed_gna, history_weight, "gna_pct")
            blended_profit = _blend_reference_value(_safe_float(anchor.get("profit_pct"), 5.0), observed_profit, history_weight, "profit_pct")
            blended_rnd = _blend_reference_value(_safe_float(anchor.get("rnd_pct"), 1.0), observed_rnd, history_weight, "rnd_pct")

            blended_material_share = _blend_reference_value(
                _safe_float(anchor.get("material_share"), 10.0),
                observed_material_share,
                direct_weight,
                "material_share",
            )
            blended_labor_share = _blend_reference_value(
                _safe_float(anchor.get("labor_share"), 45.0),
                observed_labor_share,
                direct_weight,
                "labor_share",
            )
            blended_expense_share = _blend_reference_value(
                _safe_float(anchor.get("expense_share"), 45.0),
                observed_expense_share,
                direct_weight,
                "expense_share",
            )
            normalized_shares = _normalize_share_triplet(
                blended_material_share,
                blended_labor_share,
                blended_expense_share,
            )

            row = {
                "industry_group": str(industry_group),
                "period_key": str(period_key),
                "cogs_pct": round(blended_cogs, 6),
                "gna_pct": round(blended_gna, 6),
                "profit_pct": round(blended_profit, 6),
                "rnd_pct": round(blended_rnd, 6),
                "material_share": round(normalized_shares["material_share"], 6),
                "labor_share": round(normalized_shares["labor_share"], 6),
                "expense_share": round(normalized_shares["expense_share"], 6),
                "source": GENERATED_OFFICIAL_REFERENCE_SOURCE,
                "note": (
                    f"{OFFICIAL_REFERENCE_SERIES_SOURCE_NOTE} "
                    f"history_weight={history_weight:.2f}; direct_weight={direct_weight:.2f}; suspicious={suspicious_count}."
                ),
            }
            refreshed_rows.append(row)
            group_history.append(row)
            build_notes.append(
                {
                    "industry_group": industry_group,
                    "period_key": str(period_key),
                    "status": "refreshed",
                    "suspicious_count": suspicious_count,
                    "history_weight": round(history_weight, 2),
                    "direct_weight": round(direct_weight, 2),
                }
            )

        if group_history:
            estimated_rows.append(_estimated_reference_row(industry_group, str(estimate_period), group_history))

    latest_sheets = load_industry_db(resolved)
    latest_official_df = latest_sheets["config_official_reference"].copy()
    latest_preserved = latest_official_df[
        ~(
            latest_official_df["period_key"].astype(str).isin(history_periods + [str(estimate_period)])
            & latest_official_df["industry_group"].astype(str).isin(industry_groups)
        )
    ].copy()
    combined_rows = latest_preserved
    if refreshed_rows or estimated_rows:
        combined_rows = _append_rows(combined_rows, refreshed_rows + estimated_rows, "config_official_reference")
    latest_sheets["config_official_reference"] = combined_rows
    _write_all_sheets(resolved, latest_sheets)

    return {
        "workbook_path": resolved,
        "history_periods": history_periods,
        "estimate_period": str(estimate_period),
        "refreshed_rows": len(refreshed_rows),
        "estimated_rows": len(estimated_rows),
        "details": build_notes,
    }


def build_estimated_industry_baseline(
    industry_group: str,
    period_key: str,
    workbook_path: Optional[str] = None,
) -> Dict[str, Any]:
    resolved = ensure_workbook(workbook_path)
    sheets = load_industry_db(resolved)
    official_ref = _official_reference(industry_group, str(period_key), sheets["config_official_reference"])
    if not official_ref:
        refresh_official_reference_sheet(resolved, estimate_period=str(period_key))
        sheets = load_industry_db(resolved)
        official_ref = _official_reference(industry_group, str(period_key), sheets["config_official_reference"])
    if not official_ref:
        raise IndustryFinancialError(f"{industry_group} / {period_key} official reference row was not found.")

    cogs_ratio = _safe_float(official_ref.get("cogs_pct"), 0.0)
    gna_ratio = _safe_float(official_ref.get("gna_pct"), 0.0)
    profit_ratio = _safe_float(official_ref.get("profit_pct"), 0.0)
    rnd_ratio = _safe_float(official_ref.get("rnd_pct"), 0.0)
    normalized_shares = _normalize_share_triplet(
        _safe_float(official_ref.get("material_share"), 10.0),
        _safe_float(official_ref.get("labor_share"), 45.0),
        _safe_float(official_ref.get("expense_share"), 45.0),
    )

    observed_ratios_raw: Dict[str, float] = {
        "material": cogs_ratio * normalized_shares["material_share"] / 100.0,
        "labor": cogs_ratio * normalized_shares["labor_share"] / 100.0,
        "expense": cogs_ratio * normalized_shares["expense_share"] / 100.0,
        "gna": max(gna_ratio, 0.0),
        "profit": profit_ratio,
    }
    if rnd_ratio > 0:
        observed_ratios_raw["rnd"] = rnd_ratio

    active_items = ["material", "labor", "expense", "gna", "profit"]
    if rnd_ratio > 0:
        active_items.append("rnd")

    normalized_base = _normalize_ratios(
        {item: max(observed_ratios_raw.get(item, 0.0), 0.0) for item in active_items},
        active_items,
    )

    range_df = sheets["config_range"]
    now = _utc_now()
    baseline_inputs_digest = _hash_payload(
        {
            "industry_group": str(industry_group),
            "period_key": str(period_key),
            "official_reference": official_ref,
        }
    )
    metric_rows: List[Dict[str, Any]] = []
    baseline_rows: List[Dict[str, Any]] = []

    for item in active_items:
        basis_type = "estimated"
        supporting_accounts = [str(official_ref.get("source", "config_official_reference")), f"period={period_key}"]
        confidence = _confidence_evaluation(
            basis_type=basis_type,
            supporting_accounts=supporting_accounts,
            cost_absent=False,
            restatement_seen=False,
            official_aligned=True,
            generated_reference=str(official_ref.get("source", "")).startswith(GENERATED_OFFICIAL_REFERENCE_SOURCE),
        )
        score = float(confidence["score"])
        grade = str(confidence["grade"])
        confidence_note = str(confidence["note"])
        base_pct = normalized_base[item]
        width = _range_width(range_df, item, basis_type, grade)
        low = _clamp(base_pct - width, 0.0, 100.0)
        high = _clamp(base_pct + width, 0.0, 100.0)
        calc_rule_code = "ESTIMATED_FROM_3Y_GROWTH_BLEND"
        metric_rows.append(
            {
                "industry_group": str(industry_group),
                "period_key": str(period_key),
                "metric_code": f"quote_ratio_{item}",
                "item_code": item,
                "value_pct": round(observed_ratios_raw.get(item, 0.0), 6),
                "basis_type": basis_type,
                "confidence_score": round(score, 6),
                "confidence_grade": grade,
                "confidence_note": confidence_note,
                "source_statement": "OFFICIAL_REF",
                "supporting_accounts": ", ".join(supporting_accounts),
                "cost_presence_flag": False,
                "inputs_digest": baseline_inputs_digest,
                "restatement_flag": False,
                "calc_rule_code": calc_rule_code,
                "created_at": now,
            }
        )
        baseline_rows.append(
            {
                "industry_group": str(industry_group),
                "period_key": str(period_key),
                "item_code": item,
                "label": ITEM_LABELS[item],
                "base_pct": round(base_pct, 6),
                "low_pct": round(low, 6),
                "high_pct": round(high, 6),
                "basis_type": basis_type,
                "confidence_score": round(score, 6),
                "confidence_grade": grade,
                "confidence_note": confidence_note,
                "quantiles_available": False,
                "quantiles_note": "Estimated baseline from official reference series and 3-year trend projection.",
                "supporting_accounts": ", ".join(supporting_accounts),
                "calc_rule_code": calc_rule_code,
                "cost_presence_flag": False,
                "inputs_digest": baseline_inputs_digest,
                "restatement_flag": False,
                "created_at": now,
            }
        )

    metric_df = _overwrite_period_rows(sheets["metric_observations"], industry_group, str(period_key))
    sheets["metric_observations"] = _append_rows(metric_df, metric_rows, "metric_observations")
    baseline_df = _overwrite_period_rows(sheets["industry_baselines"], industry_group, str(period_key))
    sheets["industry_baselines"] = _append_rows(baseline_df, baseline_rows, "industry_baselines")
    _write_all_sheets(resolved, sheets)

    return {
        "industry_group": str(industry_group),
        "period_key": str(period_key),
        "inputs_digest": baseline_inputs_digest,
        "cost_usable": False,
        "cost_present": False,
        "suspicious_cogs": False,
        "suspicious_gna": False,
        "suspicious_profit": False,
        "restatement_seen": False,
        "official_reference_source": str(official_ref.get("source", "")),
        "items": baseline_rows,
        "metric_rows": metric_rows,
        "quantiles_note": "Estimated baseline from official reference series and 3-year trend projection.",
    }


def refresh_estimated_baselines(
    workbook_path: Optional[str] = None,
    target_period: str = "2025",
) -> Dict[str, Any]:
    resolved = ensure_workbook(workbook_path)
    sheets = load_industry_db(resolved)
    official_df = sheets["config_official_reference"]
    target_rows = official_df[official_df["period_key"].astype(str) == str(target_period)].copy()
    if target_rows.empty:
        refresh_official_reference_sheet(resolved, estimate_period=str(target_period))
        sheets = load_industry_db(resolved)
        official_df = sheets["config_official_reference"]
        target_rows = official_df[official_df["period_key"].astype(str) == str(target_period)].copy()

    results: List[Dict[str, Any]] = []
    errors: List[Dict[str, Any]] = []
    for industry_group in sorted(target_rows["industry_group"].astype(str).unique().tolist()):
        try:
            baseline = build_estimated_industry_baseline(industry_group, str(target_period), workbook_path=resolved)
            results.append(
                {
                    "industry_group": str(industry_group),
                    "period_key": str(target_period),
                    "item_count": len(baseline.get("items", [])),
                    "official_reference_source": baseline.get("official_reference_source", ""),
                }
            )
        except Exception as exc:
            errors.append(
                {
                    "industry_group": str(industry_group),
                    "period_key": str(target_period),
                    "error": str(exc),
                }
            )
    return {
        "workbook_path": resolved,
        "target_period": str(target_period),
        "results": results,
        "errors": errors,
    }


def build_industry_baseline(industry_group: str, period_key: str, workbook_path: Optional[str] = None) -> Dict[str, Any]:
    resolved = ensure_workbook(workbook_path)
    sheets = load_industry_db(resolved)
    statement_rows = _latest_statement_rows(sheets["statement_rows"], industry_group, str(period_key))
    if statement_rows.empty:
        if YEAR_RE.fullmatch(str(period_key)):
            return build_estimated_industry_baseline(industry_group, str(period_key), workbook_path=resolved)
        raise IndustryFinancialError(f"{industry_group} / {period_key} baseline을 위한 관측값이 없습니다.")

    values = {
        row["normalized_account_code"]: float(row["amount"])
        for _, row in statement_rows.iterrows()
        if pd.notna(row["normalized_account_code"]) and pd.notna(row["amount"])
    }
    supporting = {
        row["normalized_account_code"]: row["raw_account_name"]
        for _, row in statement_rows.iterrows()
        if pd.notna(row["normalized_account_code"]) and pd.notna(row["raw_account_name"])
    }

    revenue = abs(values.get("is_revenue", 0.0))
    cogs = abs(values.get("is_cogs", 0.0))
    gna = abs(values.get("is_gna", 0.0))
    operating_profit_raw = float(values.get("is_operating_profit", 0.0))
    if revenue <= 0:
        raise IndustryFinancialError(f"{industry_group} / {period_key} 매출액을 계산하지 못했습니다.")

    material_amount = abs(values.get("cost_material", 0.0))
    labor_amount = abs(values.get("cost_labor", 0.0))
    expense_amount = abs(values.get("cost_expense", 0.0))
    inventory_amount = abs(values.get("bs_inventory", 0.0))
    ppe_amount = abs(values.get("bs_ppe", 0.0))
    depreciation_amount = abs(values.get("is_depreciation", 0.0))
    material_nature_amount = abs(values.get("is_material_nature", 0.0))
    employee_benefits_amount = abs(values.get("is_employee_benefits", 0.0))
    capex_amount = abs(values.get("cf_capex", 0.0))
    inventory_available = "bs_inventory" in values
    ppe_available = "bs_ppe" in values
    depreciation_available = "is_depreciation" in values
    material_nature_available = "is_material_nature" in values
    employee_benefits_available = "is_employee_benefits" in values
    capex_available = "cf_capex" in values
    previous_inventory_amount: Optional[float] = None
    previous_period_key = _previous_numeric_period_key(str(period_key))
    if previous_period_key and inventory_available:
        previous_inventory_amount = _latest_statement_value(sheets["statement_rows"], industry_group, previous_period_key, "bs_inventory")
        if previous_inventory_amount is not None:
            previous_inventory_amount = abs(previous_inventory_amount)
    cost_sum = material_amount + labor_amount + expense_amount
    cost_present = bool(cost_sum > 0)
    suspicious_cogs = _is_suspicious_cogs(cogs, revenue, cost_sum)
    cost_usable = bool(cost_sum > 0 and abs(cost_sum - cogs) / revenue <= 0.03)
    cost_direct_mode = bool(cost_usable or suspicious_cogs)
    restatement_seen = statement_rows["restatement_flag"].astype(str).str.lower().eq("true").any()
    baseline_inputs_digest = _hash_payload(
        {
            "industry_group": industry_group,
            "period_key": period_key,
            "statement_inputs": sorted(statement_rows["inputs_digest"].astype(str).unique().tolist()),
            "cost_usable": cost_usable,
            "suspicious_cogs": suspicious_cogs,
        }
    )

    proxy_df = sheets["config_proxy_mix"]
    range_df = sheets["config_range"]
    now = _utc_now()

    observed_ratios_raw: Dict[str, float] = {
        "gna": gna / revenue * 100.0,
        "profit": operating_profit_raw / revenue * 100.0,
    }
    basis_map = {"gna": "observed", "profit": "observed"}
    calc_rules = {"gna": "IS_SGA_TO_REVENUE", "profit": "IS_OPERATING_PROFIT_TO_REVENUE"}
    support_map = {
        "gna": [supporting.get("is_gna", "판매비와관리비"), supporting.get("is_revenue", "매출액")],
        "profit": [supporting.get("is_operating_profit", "영업이익(손실)"), supporting.get("is_revenue", "매출액")],
    }

    if cost_usable:
        observed_ratios_raw.update(
            {
                "material": material_amount / revenue * 100.0,
                "labor": labor_amount / revenue * 100.0,
                "expense": expense_amount / revenue * 100.0,
            }
        )
        basis_map.update({"material": "observed", "labor": "observed", "expense": "observed"})
        calc_rules.update(
            {
                "material": "COST_DIRECT_TO_REVENUE",
                "labor": "COST_DIRECT_TO_REVENUE",
                "expense": "COST_DIRECT_TO_REVENUE",
            }
        )
        support_map.update(
            {
                "material": [supporting.get("cost_material", "원재료비"), supporting.get("is_revenue", "매출액")],
                "labor": [supporting.get("cost_labor", "노무비"), supporting.get("is_revenue", "매출액")],
                "expense": [supporting.get("cost_expense", "경비"), supporting.get("is_revenue", "매출액")],
            }
        )
    elif suspicious_cogs:
        observed_ratios_raw.update(
            {
                "material": material_amount / revenue * 100.0,
                "labor": labor_amount / revenue * 100.0,
                "expense": expense_amount / revenue * 100.0,
            }
        )
        basis_map.update({"material": "estimated", "labor": "estimated", "expense": "estimated"})
        calc_rules.update(
            {
                "material": "COST_PARTIAL_TO_REVENUE",
                "labor": "COST_PARTIAL_TO_REVENUE",
                "expense": "COST_PARTIAL_TO_REVENUE",
            }
        )
        support_map.update(
            {
                "material": [supporting.get("cost_material", "원재료비"), supporting.get("is_revenue", "매출액")],
                "labor": [supporting.get("cost_labor", "노무비"), supporting.get("is_revenue", "매출액")],
                "expense": [supporting.get("cost_expense", "경비"), supporting.get("is_revenue", "매출액")],
            }
        )
    else:
        cogs_ratio = cogs / revenue * 100.0
        mix = _proxy_mix(industry_group, proxy_df)
        proxy_basis = _basis_type_for_proxy(industry_group, proxy_df)
        observed_ratios_raw.update(
            {
                "material": cogs_ratio * mix["material"] / 100.0,
                "labor": cogs_ratio * mix["labor"] / 100.0,
                "expense": cogs_ratio * mix["expense"] / 100.0,
            }
        )
        basis_map.update({"material": proxy_basis, "labor": proxy_basis, "expense": proxy_basis})
        calc_rules.update(
            {
                "material": "PROXY_MIX_FROM_COGS",
                "labor": "PROXY_MIX_FROM_COGS",
                "expense": "PROXY_MIX_FROM_COGS",
            }
        )
        support_map.update(
            {
                "material": [supporting.get("is_cogs", "매출원가"), supporting.get("is_revenue", "매출액")],
                "labor": [supporting.get("is_cogs", "매출원가"), supporting.get("is_revenue", "매출액")],
                "expense": [supporting.get("is_cogs", "매출원가"), supporting.get("is_revenue", "매출액")],
            }
        )

    rnd_observed = abs(values.get("is_rnd_research", 0.0)) + abs(values.get("is_rnd_development", 0.0))
    if rnd_observed > 0:
        rnd_ratio = rnd_observed / revenue * 100.0
        observed_ratios_raw["rnd"] = rnd_ratio
        basis_map["rnd"] = "observed"
        calc_rules["rnd"] = "IS_RND_TO_REVENUE"
        support_map["rnd"] = [
            name
            for name in [
                supporting.get("is_rnd_research", ""),
                supporting.get("is_rnd_development", ""),
                supporting.get("is_revenue", "매출액"),
            ]
            if name
        ]
        observed_ratios_raw["gna"] = max(observed_ratios_raw["gna"] - rnd_ratio, 0.0)

    active_items = ["material", "labor", "expense", "gna", "profit"]
    if observed_ratios_raw.get("rnd", 0.0) > 0:
        active_items.append("rnd")

    normalized_base = _normalize_ratios({item: max(observed_ratios_raw.get(item, 0.0), 0.0) for item in active_items}, active_items)

    metric_rows: List[Dict[str, Any]] = []
    baseline_rows: List[Dict[str, Any]] = []
    for item in active_items:
        basis_type = basis_map[item]
        supporting_accounts = [name for name in support_map.get(item, []) if name]
        score = _confidence_score(
            basis_type=basis_type,
            supporting_accounts=supporting_accounts,
            cost_absent=(not cost_present and item in {"material", "labor", "expense"}),
            restatement_seen=restatement_seen,
        )
        grade = _grade_confidence(score)
        base_pct = normalized_base[item]
        width = _range_width(range_df, item, basis_type, grade)
        low = _clamp(base_pct - width, 0.0, 100.0)
        high = _clamp(base_pct + width, 0.0, 100.0)
        metric_rows.append(
            {
                "industry_group": industry_group,
                "period_key": str(period_key),
                "metric_code": f"quote_ratio_{item}",
                "item_code": item,
                "value_pct": round(observed_ratios_raw.get(item, 0.0), 6),
                "basis_type": basis_type,
                "confidence_score": round(score, 6),
                "confidence_grade": grade,
                "source_statement": "COST" if item in {"material", "labor", "expense"} and cost_direct_mode else "IS",
                "supporting_accounts": ", ".join(supporting_accounts),
                "cost_presence_flag": bool(cost_present),
                "inputs_digest": baseline_inputs_digest,
                "restatement_flag": bool(restatement_seen),
                "calc_rule_code": calc_rules[item],
                "created_at": now,
            }
        )
        baseline_rows.append(
            {
                "industry_group": industry_group,
                "period_key": str(period_key),
                "item_code": item,
                "label": ITEM_LABELS[item],
                "base_pct": round(base_pct, 6),
                "low_pct": round(low, 6),
                "high_pct": round(high, 6),
                "basis_type": basis_type,
                "confidence_score": round(score, 6),
                "confidence_grade": grade,
                "confidence_note": confidence_note,
                "quantiles_available": False,
                "quantiles_note": "분위수 불가: 업종 aggregate 기준만 있습니다.",
                "supporting_accounts": ", ".join(supporting_accounts),
                "calc_rule_code": calc_rules[item],
                "cost_presence_flag": bool(cost_present),
                "inputs_digest": baseline_inputs_digest,
                "restatement_flag": bool(restatement_seen),
                "created_at": now,
            }
        )

    metric_df = sheets["metric_observations"]
    metric_df = metric_df[
        ~(
            (metric_df["industry_group"].astype(str) == str(industry_group))
            & (metric_df["period_key"].astype(str) == str(period_key))
        )
    ]
    sheets["metric_observations"] = _append_rows(metric_df, metric_rows, "metric_observations")

    baseline_df = sheets["industry_baselines"]
    baseline_df = baseline_df[
        ~(
            (baseline_df["industry_group"].astype(str) == str(industry_group))
            & (baseline_df["period_key"].astype(str) == str(period_key))
        )
    ]
    sheets["industry_baselines"] = _append_rows(baseline_df, baseline_rows, "industry_baselines")
    _write_all_sheets(resolved, sheets)

    return {
        "industry_group": industry_group,
        "period_key": str(period_key),
        "inputs_digest": baseline_inputs_digest,
        "cost_usable": bool(cost_usable),
        "cost_present": bool(cost_present),
        "suspicious_cogs": bool(suspicious_cogs),
        "restatement_seen": bool(restatement_seen),
        "items": baseline_rows,
        "metric_rows": metric_rows,
        "quantiles_note": "분위수 불가: 업종 aggregate 기준만 있습니다.",
    }


def build_industry_baseline(industry_group: str, period_key: str, workbook_path: Optional[str] = None) -> Dict[str, Any]:
    resolved = ensure_workbook(workbook_path)
    sheets = load_industry_db(resolved)
    statement_rows = _latest_statement_rows(sheets["statement_rows"], industry_group, str(period_key))
    if statement_rows.empty:
        raise IndustryFinancialError(f"{industry_group} / {period_key} baseline 관측값이 없습니다.")

    values = {
        row["normalized_account_code"]: float(row["amount"])
        for _, row in statement_rows.iterrows()
        if pd.notna(row["normalized_account_code"]) and pd.notna(row["amount"])
    }
    supporting = {
        row["normalized_account_code"]: row["raw_account_name"]
        for _, row in statement_rows.iterrows()
        if pd.notna(row["normalized_account_code"]) and pd.notna(row["raw_account_name"])
    }

    revenue = abs(values.get("is_revenue", 0.0))
    cogs = abs(values.get("is_cogs", 0.0))
    gna = abs(values.get("is_gna", 0.0))
    gross_profit_raw = float(values.get("is_gross_profit", 0.0))
    operating_profit_raw = float(values.get("is_operating_profit", 0.0))
    cf_operating_cash_flow = float(values.get("cf_operating_cash_flow", 0.0))
    if revenue <= 0:
        raise IndustryFinancialError(f"{industry_group} / {period_key} 매출액을 계산하지 못했습니다.")

    material_amount = abs(values.get("cost_material", 0.0))
    labor_amount = abs(values.get("cost_labor", 0.0))
    expense_amount = abs(values.get("cost_expense", 0.0))
    inventory_amount = abs(values.get("bs_inventory", 0.0))
    ppe_amount = abs(values.get("bs_ppe", 0.0))
    depreciation_amount = abs(values.get("is_depreciation", 0.0))
    material_nature_amount = abs(values.get("is_material_nature", 0.0))
    employee_benefits_amount = abs(values.get("is_employee_benefits", 0.0))
    capex_amount = abs(values.get("cf_capex", 0.0))
    inventory_available = "bs_inventory" in values
    ppe_available = "bs_ppe" in values
    depreciation_available = "is_depreciation" in values
    material_nature_available = "is_material_nature" in values
    employee_benefits_available = "is_employee_benefits" in values
    capex_available = "cf_capex" in values
    previous_inventory_amount: Optional[float] = None
    previous_period_key = _previous_numeric_period_key(str(period_key))
    if previous_period_key and inventory_available:
        previous_inventory_amount = _latest_statement_value(
            sheets["statement_rows"], industry_group, previous_period_key, "bs_inventory"
        )
        if previous_inventory_amount is not None:
            previous_inventory_amount = abs(previous_inventory_amount)
    cost_sum = material_amount + labor_amount + expense_amount
    cost_present = bool(cost_sum > 0)
    suspicious_cogs_raw = _is_suspicious_cogs(cogs, revenue, cost_sum)
    restatement_seen = statement_rows["restatement_flag"].astype(str).str.lower().eq("true").any()

    proxy_df = sheets["config_proxy_mix"]
    official_df = sheets["config_official_reference"]
    range_df = sheets["config_range"]
    now = _utc_now()

    mix = _proxy_mix(industry_group, proxy_df)
    proxy_basis = _basis_type_for_proxy(industry_group, proxy_df)
    official_ref = _official_reference(industry_group, str(period_key), official_df)
    baseline_inputs_digest = _hash_payload(
        {
            "industry_group": industry_group,
            "period_key": str(period_key),
            "statement_inputs": sorted(statement_rows["inputs_digest"].astype(str).unique().tolist()),
            "official_reference": official_ref,
            "cost_present": bool(cost_present),
            "suspicious_cogs_raw": bool(suspicious_cogs_raw),
        }
    )

    raw_cogs_ratio = _ratio_pct(cogs, revenue)
    raw_gna_ratio = _ratio_pct(gna, revenue)
    raw_profit_ratio = _ratio_pct(operating_profit_raw, revenue, absolute=False)
    cost_ratio_pct = _ratio_pct(cost_sum, revenue)

    derived_cogs_from_gp = _ratio_pct(revenue - gross_profit_raw, revenue, absolute=False) if gross_profit_raw else None
    derived_cogs_from_gna_profit = _ratio_pct(revenue - gna - operating_profit_raw, revenue, absolute=False) if (gna or operating_profit_raw) else None
    derived_gna_from_gp_op = _ratio_pct(gross_profit_raw - operating_profit_raw, revenue, absolute=False) if gross_profit_raw else None
    derived_profit_from_gp_gna = _ratio_pct(gross_profit_raw - gna, revenue, absolute=False) if gross_profit_raw or gna else None

    suspicious_cogs = bool(
        suspicious_cogs_raw
        or (derived_cogs_from_gp is not None and abs(raw_cogs_ratio - derived_cogs_from_gp) > 8.0)
        or (raw_cogs_ratio < 1.0 and (gross_profit_raw > 0 or cost_sum > 0))
    )
    suspicious_gna = bool(
        gross_profit_raw > 0
        and raw_gna_ratio > 0
        and abs(raw_gna_ratio - max(derived_gna_from_gp_op or 0.0, 0.0)) > 6.0
    )
    suspicious_profit = bool(
        gross_profit_raw > 0
        and abs(raw_profit_ratio - (derived_profit_from_gp_gna or 0.0)) > 3.0
    )

    cogs_low, cogs_high = _ratio_band(float(official_ref.get("cogs_pct", 0.0)), 15.0, 98.0, 18.0)
    cogs_candidates: List[Tuple[Optional[float], str, List[str]]] = []
    cogs_from_cost_inventory = None
    if cost_sum > 0 and inventory_available and previous_inventory_amount is not None:
        cogs_from_cost_inventory = _ratio_pct(cost_sum + previous_inventory_amount - inventory_amount, revenue, absolute=False)
    if suspicious_cogs:
        cogs_candidates.extend(
            [
                (
                    derived_cogs_from_gp,
                    "REPAIRED_COGS_FROM_GROSS_PROFIT",
                    [supporting.get("is_gross_profit", "매출총이익(손실)"), supporting.get("is_revenue", "매출액")],
                ),
                (
                    derived_cogs_from_gna_profit,
                    "REPAIRED_COGS_FROM_GNA_AND_PROFIT",
                    [
                        supporting.get("is_gna", "판매비와관리비"),
                        supporting.get("is_operating_profit", "영업이익(손실)"),
                        supporting.get("is_revenue", "매출액"),
                    ],
                ),
                (
                    _ratio_pct(cost_sum, revenue),
                    "REPAIRED_COGS_FROM_COST_SUM",
                    [
                        supporting.get("cost_material", "원재료비"),
                        supporting.get("cost_labor", "노무비"),
                        supporting.get("cost_expense", "경비"),
                        supporting.get("is_revenue", "매출액"),
                    ],
                ),
                (
                    cogs_from_cost_inventory,
                    "REPAIRED_COGS_FROM_COST_AND_INVENTORY",
                    [
                        supporting.get("cost_material", "원재료비"),
                        supporting.get("cost_labor", "노무비"),
                        supporting.get("cost_expense", "경비"),
                        supporting.get("bs_inventory", "재고자산"),
                        f"prior_bs_inventory:{previous_period_key}" if previous_inventory_amount is not None and previous_period_key else "",
                        supporting.get("is_revenue", "매출액"),
                    ],
                ),
                (
                    float(official_ref.get("cogs_pct", 0.0)) if official_ref else None,
                    "OFFICIAL_REFERENCE_COGS",
                    [str(official_ref.get("source", "official_reference"))] if official_ref else [],
                ),
            ]
        )
    cogs_candidates.append(
        (
            raw_cogs_ratio,
            "IS_COGS_TO_REVENUE",
            [supporting.get("is_cogs", "매출원가"), supporting.get("is_revenue", "매출액")],
        )
    )
    repaired_cogs_ratio, repaired_cogs_rule, repaired_cogs_support = _choose_ratio_candidate(
        cogs_candidates,
        cogs_low,
        cogs_high,
        anchor=float(official_ref.get("cogs_pct", 0.0)) if official_ref else None,
        official_source=str(official_ref.get("source", "")) if official_ref else "",
    )
    repaired_cogs_ratio = float(repaired_cogs_ratio or 0.0)
    repaired_cogs_amount = revenue * repaired_cogs_ratio / 100.0

    profit_low, profit_high = _ratio_band(float(official_ref.get("profit_pct", 0.0)), -15.0, 25.0, 10.0)
    profit_candidates: List[Tuple[Optional[float], str, List[str]]] = []
    if suspicious_profit:
        profit_candidates.extend(
            [
                (
                    derived_profit_from_gp_gna,
                    "REPAIRED_PROFIT_FROM_GROSS_PROFIT_AND_GNA",
                    [supporting.get("is_gross_profit", "매출총이익(손실)"), supporting.get("is_gna", "판매비와관리비")],
                ),
                (
                    _ratio_pct(revenue - repaired_cogs_amount - gna, revenue, absolute=False),
                    "REPAIRED_PROFIT_FROM_COGS_AND_GNA",
                    [
                        supporting.get("is_cogs", "매출원가"),
                        supporting.get("is_gna", "판매비와관리비"),
                        supporting.get("is_revenue", "매출액"),
                    ],
                ),
                (
                    float(official_ref.get("profit_pct", 0.0)) if official_ref else None,
                    "OFFICIAL_REFERENCE_PROFIT",
                    [str(official_ref.get("source", "official_reference"))] if official_ref else [],
                ),
            ]
        )
    profit_candidates.append(
        (
            raw_profit_ratio,
            "IS_OPERATING_PROFIT_TO_REVENUE",
            [supporting.get("is_operating_profit", "영업이익(손실)"), supporting.get("is_revenue", "매출액")],
        )
    )
    repaired_profit_ratio, repaired_profit_rule, repaired_profit_support = _choose_ratio_candidate(
        profit_candidates,
        profit_low,
        profit_high,
        allow_negative=True,
        anchor=float(official_ref.get("profit_pct", 0.0)) if official_ref else None,
        official_source=str(official_ref.get("source", "")) if official_ref else "",
    )
    repaired_profit_ratio = float(repaired_profit_ratio or 0.0)
    cashflow_quality = _cashflow_quality_signal(
        revenue,
        cf_operating_cash_flow,
        repaired_profit_ratio,
        float(official_ref.get("profit_pct", 0.0)) if official_ref else None,
    )
    if cashflow_quality["status"] == "weak" and not math.isclose(
        repaired_profit_ratio,
        float(cashflow_quality["adjusted_profit_ratio"]),
        rel_tol=0.0,
        abs_tol=1e-6,
    ):
        repaired_profit_ratio = float(cashflow_quality["adjusted_profit_ratio"])
        repaired_profit_rule = f"{repaired_profit_rule}|CFO_QUALITY_ADJUSTED"
        repaired_profit_support = repaired_profit_support + [supporting.get("cf_operating_cash_flow", "영업활동으로 인한 현금흐름")]

    gna_low, gna_high = _ratio_band(float(official_ref.get("gna_pct", 0.0)), 0.0, 80.0, 18.0)
    gna_candidates: List[Tuple[Optional[float], str, List[str]]] = []
    if suspicious_gna or suspicious_profit:
        gna_candidates.extend(
            [
                (
                    derived_gna_from_gp_op,
                    "REPAIRED_GNA_FROM_GROSS_PROFIT_AND_PROFIT",
                    [supporting.get("is_gross_profit", "매출총이익(손실)"), supporting.get("is_operating_profit", "영업이익(손실)")],
                ),
                (
                    _ratio_pct(revenue - repaired_cogs_amount - (repaired_profit_ratio / 100.0 * revenue), revenue, absolute=False),
                    "REPAIRED_GNA_FROM_COGS_AND_PROFIT",
                    [
                        supporting.get("is_cogs", "매출원가"),
                        supporting.get("is_operating_profit", "영업이익(손실)"),
                        supporting.get("is_revenue", "매출액"),
                    ],
                ),
                (
                    float(official_ref.get("gna_pct", 0.0)) if official_ref else None,
                    "OFFICIAL_REFERENCE_GNA",
                    [str(official_ref.get("source", "official_reference"))] if official_ref else [],
                ),
            ]
        )
    gna_candidates.append(
        (
            raw_gna_ratio,
            "IS_SGA_TO_REVENUE",
            [supporting.get("is_gna", "판매비와관리비"), supporting.get("is_revenue", "매출액")],
        )
    )
    repaired_gna_ratio, repaired_gna_rule, repaired_gna_support = _choose_ratio_candidate(
        gna_candidates,
        gna_low,
        gna_high,
        anchor=float(official_ref.get("gna_pct", 0.0)) if official_ref else None,
        official_source=str(official_ref.get("source", "")) if official_ref else "",
    )
    repaired_gna_ratio = float(repaired_gna_ratio or 0.0)

    observed_ratios_raw: Dict[str, float] = {
        "gna": repaired_gna_ratio,
        "profit": repaired_profit_ratio,
    }
    basis_map = {
        "gna": "observed"
        if repaired_gna_rule == "IS_SGA_TO_REVENUE"
        else ("repaired" if str(repaired_gna_rule).startswith("REPAIRED_") else "proxy"),
        "profit": "observed"
        if repaired_profit_rule == "IS_OPERATING_PROFIT_TO_REVENUE"
        else ("repaired" if str(repaired_profit_rule).startswith("REPAIRED_") else "proxy"),
    }
    calc_rules = {
        "gna": repaired_gna_rule or "IS_SGA_TO_REVENUE",
        "profit": repaired_profit_rule or "IS_OPERATING_PROFIT_TO_REVENUE",
    }
    support_map = {
        "gna": repaired_gna_support,
        "profit": repaired_profit_support,
    }

    cost_usable = bool(cost_sum > 0 and abs(cost_ratio_pct - repaired_cogs_ratio) <= 3.0)
    cost_direct_mode = bool(cost_present)
    if cost_usable:
        observed_ratios_raw.update(
            {
                "material": _ratio_pct(material_amount, revenue),
                "labor": _ratio_pct(labor_amount, revenue),
                "expense": _ratio_pct(expense_amount, revenue),
            }
        )
        basis_map.update({"material": "observed", "labor": "observed", "expense": "observed"})
        calc_rules.update(
            {
                "material": "COST_DIRECT_TO_REVENUE",
                "labor": "COST_DIRECT_TO_REVENUE",
                "expense": "COST_DIRECT_TO_REVENUE",
            }
        )
        support_map.update(
            {
                "material": [supporting.get("cost_material", "원재료비"), supporting.get("is_revenue", "매출액")],
                "labor": [supporting.get("cost_labor", "노무비"), supporting.get("is_revenue", "매출액")],
                "expense": [supporting.get("cost_expense", "경비"), supporting.get("is_revenue", "매출액")],
            }
        )
    elif cost_present and repaired_cogs_ratio > 0:
        scaled_direct = _scaled_direct_cost_ratios(revenue, repaired_cogs_ratio, material_amount, labor_amount, expense_amount, mix)
        observed_ratios_raw.update(scaled_direct)
        basis_map.update({"material": "repaired", "labor": "repaired", "expense": "repaired"})
        calc_rules.update(
            {
                "material": "COST_SCALED_TO_REPAIRED_COGS",
                "labor": "COST_SCALED_TO_REPAIRED_COGS",
                "expense": "COST_SCALED_TO_REPAIRED_COGS",
            }
        )
        support_map.update(
            {
                "material": [supporting.get("cost_material", "원재료비"), repaired_cogs_rule],
                "labor": [supporting.get("cost_labor", "노무비"), repaired_cogs_rule],
                "expense": [supporting.get("cost_expense", "경비"), repaired_cogs_rule],
            }
        )
    else:
        reference_mix = {
            "material": float(official_ref.get("material_share", mix["material"])) if official_ref else mix["material"],
            "labor": float(official_ref.get("labor_share", mix["labor"])) if official_ref else mix["labor"],
            "expense": float(official_ref.get("expense_share", mix["expense"])) if official_ref else mix["expense"],
        }
        scaled_direct = _scaled_direct_cost_ratios(revenue, repaired_cogs_ratio, 0.0, 0.0, 0.0, reference_mix)
        observed_ratios_raw.update(scaled_direct)
        basis_map.update({"material": proxy_basis, "labor": proxy_basis, "expense": proxy_basis})
        calc_rules.update(
            {
                "material": "OFFICIAL_OR_PROXY_MIX_TO_REPAIRED_COGS",
                "labor": "OFFICIAL_OR_PROXY_MIX_TO_REPAIRED_COGS",
                "expense": "OFFICIAL_OR_PROXY_MIX_TO_REPAIRED_COGS",
            }
        )
        mix_support = [str(official_ref.get("source", "config_proxy_mix"))] if official_ref else ["config_proxy_mix"]
        support_map.update({"material": mix_support, "labor": mix_support, "expense": mix_support})

    direct_target_ratio = repaired_cogs_ratio if repaired_cogs_ratio > 0 else (
        max(observed_ratios_raw.get("material", 0.0), 0.0)
        + max(observed_ratios_raw.get("labor", 0.0), 0.0)
        + max(observed_ratios_raw.get("expense", 0.0), 0.0)
    )
    if direct_target_ratio > 0 and material_nature_available:
        material_nature_ratio = _ratio_pct(material_nature_amount, revenue)
        if material_nature_ratio > 0 and basis_map.get("material") != "observed":
            current_material_ratio = max(observed_ratios_raw.get("material", 0.0), 0.0)
            adjusted_material_ratio = min(direct_target_ratio, (current_material_ratio * 0.40) + (material_nature_ratio * 0.60))
            direct_values = {
                "material": adjusted_material_ratio,
                "labor": max(observed_ratios_raw.get("labor", 0.0), 0.0),
                "expense": max(observed_ratios_raw.get("expense", 0.0), 0.0),
            }
            rebalanced = _rebalance_direct_ratios(direct_values, direct_target_ratio, locked_items={"material"})
            observed_ratios_raw.update(rebalanced)
            basis_map["material"] = "repaired"
            calc_rules["material"] = "REPAIRED_MATERIAL_FROM_NATURE_NOTE"
            support_map["material"] = [
                supporting.get("is_material_nature", "원재료 및 소모품 사용액과 상품의 매입액"),
                supporting.get("is_revenue", "매출액"),
            ]

    if direct_target_ratio > 0 and employee_benefits_available and basis_map.get("labor") != "observed":
        employee_cost_ratio = _ratio_pct(employee_benefits_amount, revenue)
        if employee_cost_ratio > 0:
            service_capture = 0.55 if industry_group in {"J", "M", "N", "S"} else 0.40
            labor_hint_ratio = min(direct_target_ratio, employee_cost_ratio * service_capture)
            current_labor_ratio = max(observed_ratios_raw.get("labor", 0.0), 0.0)
            adjusted_labor_ratio = min(direct_target_ratio, (current_labor_ratio * 0.60) + (labor_hint_ratio * 0.40))
            direct_values = {
                "material": max(observed_ratios_raw.get("material", 0.0), 0.0),
                "labor": adjusted_labor_ratio,
                "expense": max(observed_ratios_raw.get("expense", 0.0), 0.0),
            }
            rebalanced = _rebalance_direct_ratios(direct_values, direct_target_ratio, locked_items={"labor"})
            observed_ratios_raw.update(rebalanced)
            basis_map["labor"] = "repaired"
            calc_rules["labor"] = f"{calc_rules['labor']}|LABOR_SUPPORT_FROM_EMPLOYEE_COST"
            support_map["labor"] = support_map.get("labor", []) + [supporting.get("is_employee_benefits", "종업원급여")]

    if direct_target_ratio > 0 and depreciation_available:
        depreciation_ratio = _ratio_pct(depreciation_amount, revenue)
        current_expense_ratio = max(observed_ratios_raw.get("expense", 0.0), 0.0)
        if depreciation_ratio > 0 and current_expense_ratio + 0.2 < depreciation_ratio:
            adjusted_expense_ratio = min(direct_target_ratio, (current_expense_ratio * 0.35) + (depreciation_ratio * 0.65))
            direct_values = {
                "material": max(observed_ratios_raw.get("material", 0.0), 0.0),
                "labor": max(observed_ratios_raw.get("labor", 0.0), 0.0),
                "expense": adjusted_expense_ratio,
            }
            rebalanced = _rebalance_direct_ratios(direct_values, direct_target_ratio, locked_items={"expense"})
            observed_ratios_raw.update(rebalanced)
            basis_map["expense"] = "repaired" if basis_map.get("expense") != "observed" else "observed"
            calc_rules["expense"] = f"{calc_rules['expense']}|DEPRECIATION_FLOOR_SUPPORT"
            support_map["expense"] = support_map.get("expense", []) + [supporting.get("is_depreciation", "감가상각비")]

    if direct_target_ratio > 0 and basis_map.get("expense") != "observed" and (ppe_available or capex_available):
        ppe_ratio = _ratio_pct(ppe_amount, revenue)
        capex_ratio = _ratio_pct(capex_amount, revenue)
        asset_hint_ratio = min(direct_target_ratio, (ppe_ratio * 0.02) + (capex_ratio * 0.20))
        current_expense_ratio = max(observed_ratios_raw.get("expense", 0.0), 0.0)
        if asset_hint_ratio > current_expense_ratio * 1.20 and asset_hint_ratio > 0:
            adjusted_expense_ratio = min(direct_target_ratio, (current_expense_ratio * 0.80) + (asset_hint_ratio * 0.20))
            direct_values = {
                "material": max(observed_ratios_raw.get("material", 0.0), 0.0),
                "labor": max(observed_ratios_raw.get("labor", 0.0), 0.0),
                "expense": adjusted_expense_ratio,
            }
            rebalanced = _rebalance_direct_ratios(direct_values, direct_target_ratio, locked_items={"expense"})
            observed_ratios_raw.update(rebalanced)
            basis_map["expense"] = "repaired" if basis_map.get("expense") != "observed" else "observed"
            calc_rules["expense"] = f"{calc_rules['expense']}|PPE_CAPEX_SUPPORT"
            if ppe_available:
                support_map["expense"] = support_map.get("expense", []) + [supporting.get("bs_ppe", "유형자산")]
            if capex_available:
                support_map["expense"] = support_map.get("expense", []) + [supporting.get("cf_capex", "유형자산의 취득")]

    rnd_observed = abs(values.get("is_rnd_research", 0.0)) + abs(values.get("is_rnd_development", 0.0))
    if rnd_observed > 0:
        rnd_ratio = _ratio_pct(rnd_observed, revenue)
        observed_ratios_raw["rnd"] = rnd_ratio
        basis_map["rnd"] = "observed"
        calc_rules["rnd"] = "IS_RND_TO_REVENUE"
        support_map["rnd"] = [
            name
            for name in [
                supporting.get("is_rnd_research", ""),
                supporting.get("is_rnd_development", ""),
                supporting.get("is_revenue", "매출액"),
            ]
            if name
        ]
        observed_ratios_raw["gna"] = max(observed_ratios_raw["gna"] - rnd_ratio, 0.0)

    active_items = ["material", "labor", "expense", "gna", "profit"]
    if observed_ratios_raw.get("rnd", 0.0) > 0:
        active_items.append("rnd")

    official_metric_map: Dict[str, float] = {}
    adjustment_note_map: Dict[str, str] = {}
    if official_ref:
        official_cogs_ratio = float(official_ref.get("cogs_pct", 0.0))
        official_shares = _normalize_share_triplet(
            float(official_ref.get("material_share", mix["material"])),
            float(official_ref.get("labor_share", mix["labor"])),
            float(official_ref.get("expense_share", mix["expense"])),
        )
        official_metric_map = {
            "material": official_cogs_ratio * official_shares["material_share"] / 100.0,
            "labor": official_cogs_ratio * official_shares["labor_share"] / 100.0,
            "expense": official_cogs_ratio * official_shares["expense_share"] / 100.0,
            "gna": float(official_ref.get("gna_pct", 0.0)),
            "profit": float(official_ref.get("profit_pct", 0.0)),
            "rnd": float(official_ref.get("rnd_pct", 0.0)),
        }

    for item_code in ["material", "labor", "expense", "gna", "profit", "rnd"]:
        if item_code not in observed_ratios_raw:
            continue
        shrink = _apply_outlier_shrink(
            item_code,
            float(observed_ratios_raw.get(item_code, 0.0)),
            official_metric_map.get(item_code),
            str(basis_map.get(item_code, "estimated")),
            len([name for name in support_map.get(item_code, []) if str(name).strip()]),
            cost_present=cost_present,
            cost_usable=cost_usable,
            restatement_seen=restatement_seen,
            official_source=str(official_ref.get("source", "")) if official_ref else "",
        )
        if shrink["applied"]:
            observed_ratios_raw[item_code] = float(shrink["value_pct"])
            calc_rules[item_code] = f"{calc_rules[item_code]}|{shrink['rule_suffix']}"
            support_map[item_code] = support_map.get(item_code, []) + [str(official_ref.get("source", "official_reference"))]
            adjustment_note_map[item_code] = str(shrink["note"])

    normalized_base = _normalize_ratios(
        {item: max(observed_ratios_raw.get(item, 0.0), 0.0) for item in active_items},
        active_items,
    )

    metric_rows: List[Dict[str, Any]] = []
    baseline_rows: List[Dict[str, Any]] = []
    for item in active_items:
        basis_type = basis_map[item]
        supporting_accounts = [name for name in support_map.get(item, []) if name]
        calc_rule_code = calc_rules[item]
        official_pct = official_metric_map.get(item)
        tolerance = 4.0 if item in {"material", "labor", "expense"} else 3.0
        official_aligned = official_pct is not None and abs(observed_ratios_raw.get(item, 0.0) - official_pct) <= tolerance
        triangulated = official_aligned and len(supporting_accounts) >= 2
        direct_cost_quality = "none"
        if item in {"material", "labor", "expense"}:
            if cost_usable:
                direct_cost_quality = "usable"
            elif cost_present:
                direct_cost_quality = "present"
        confidence = _confidence_evaluation(
            basis_type=basis_type,
            supporting_accounts=supporting_accounts,
            cost_absent=(not cost_present and item in {"material", "labor", "expense"}),
            restatement_seen=restatement_seen,
            official_aligned=official_aligned,
            formula_repaired=str(calc_rule_code).startswith("REPAIRED_"),
            direct_cost_quality=direct_cost_quality,
            triangulated=triangulated,
            cashflow_quality=cashflow_quality["status"] if item in {"gna", "profit"} else "neutral",
        )
        score = float(confidence["score"])
        grade = str(confidence["grade"])
        confidence_note = str(confidence["note"])
        extra_notes: List[str] = []
        if item in {"gna", "profit"} and cashflow_quality["status"] in {"weak", "strong"}:
            extra_notes.append(str(cashflow_quality["note"]))
        if adjustment_note_map.get(item):
            extra_notes.append(str(adjustment_note_map[item]))
        if extra_notes:
            confidence_note = f"{confidence_note} | {' / '.join(extra_notes)}"
        base_pct = normalized_base[item]
        width = _range_width(range_df, item, basis_type, grade)
        low = _clamp(base_pct - width, 0.0, 100.0)
        high = _clamp(base_pct + width, 0.0, 100.0)
        metric_rows.append(
            {
                "industry_group": industry_group,
                "period_key": str(period_key),
                "metric_code": f"quote_ratio_{item}",
                "item_code": item,
                "value_pct": round(observed_ratios_raw.get(item, 0.0), 6),
                "basis_type": basis_type,
                "confidence_score": round(score, 6),
                "confidence_grade": grade,
                "confidence_note": confidence_note,
                "source_statement": "COST" if item in {"material", "labor", "expense"} and cost_direct_mode else "IS",
                "supporting_accounts": ", ".join(supporting_accounts),
                "cost_presence_flag": bool(cost_present),
                "inputs_digest": baseline_inputs_digest,
                "restatement_flag": bool(restatement_seen),
                "calc_rule_code": calc_rule_code,
                "created_at": now,
            }
        )
        baseline_rows.append(
            {
                "industry_group": industry_group,
                "period_key": str(period_key),
                "item_code": item,
                "label": ITEM_LABELS[item],
                "base_pct": round(base_pct, 6),
                "low_pct": round(low, 6),
                "high_pct": round(high, 6),
                "basis_type": basis_type,
                "confidence_score": round(score, 6),
                "confidence_grade": grade,
                "confidence_note": confidence_note,
                "quantiles_available": False,
                "quantiles_note": "분위수 불가: 업종 aggregate 기준만 있습니다.",
                "supporting_accounts": ", ".join(supporting_accounts),
                "calc_rule_code": calc_rule_code,
                "cost_presence_flag": bool(cost_present),
                "inputs_digest": baseline_inputs_digest,
                "restatement_flag": bool(restatement_seen),
                "created_at": now,
            }
        )

    metric_df = sheets["metric_observations"]
    metric_df = metric_df[
        ~(
            (metric_df["industry_group"].astype(str) == str(industry_group))
            & (metric_df["period_key"].astype(str) == str(period_key))
        )
    ]
    sheets["metric_observations"] = _append_rows(metric_df, metric_rows, "metric_observations")

    baseline_df = sheets["industry_baselines"]
    baseline_df = baseline_df[
        ~(
            (baseline_df["industry_group"].astype(str) == str(industry_group))
            & (baseline_df["period_key"].astype(str) == str(period_key))
        )
    ]
    sheets["industry_baselines"] = _append_rows(baseline_df, baseline_rows, "industry_baselines")
    _write_all_sheets(resolved, sheets)

    return {
        "industry_group": industry_group,
        "period_key": str(period_key),
        "inputs_digest": baseline_inputs_digest,
        "cost_usable": bool(cost_usable),
        "cost_present": bool(cost_present),
        "suspicious_cogs": bool(suspicious_cogs),
        "suspicious_gna": bool(suspicious_gna),
        "suspicious_profit": bool(suspicious_profit),
        "restatement_seen": bool(restatement_seen),
        "official_reference_source": str(official_ref.get("source", "")) if official_ref else "",
        "items": baseline_rows,
        "metric_rows": metric_rows,
        "quantiles_note": "분위수 불가: 업종 aggregate 기준만 있습니다.",
    }


def list_available_baselines(workbook_path: Optional[str] = None) -> pd.DataFrame:
    sheets = load_industry_db(workbook_path)
    baseline_df = sheets["industry_baselines"].copy()
    if baseline_df.empty:
        return pd.DataFrame(columns=["industry_group", "period_key", "item_count", "inputs_digest", "cost_presence_flag"])
    grouped = (
        baseline_df.groupby(["industry_group", "period_key"], dropna=False)
        .agg(
            item_count=("item_code", "count"),
            inputs_digest=("inputs_digest", "last"),
            cost_presence_flag=("cost_presence_flag", "max"),
        )
        .reset_index()
        .sort_values(["industry_group", "period_key"])
    )
    return grouped


def get_baseline_items(industry_group: str, period_key: str, workbook_path: Optional[str] = None) -> List[Dict[str, Any]]:
    resolved = ensure_workbook(workbook_path)
    sheets = load_industry_db(resolved)

    def _rebuild_items() -> List[Dict[str, Any]]:
        statement_rows = _latest_statement_rows(sheets["statement_rows"], industry_group, str(period_key))
        if statement_rows.empty:
            baseline = build_estimated_industry_baseline(industry_group, str(period_key), workbook_path=resolved)
        else:
            baseline = build_industry_baseline(industry_group, period_key, workbook_path=resolved)
        return baseline["items"]

    baseline_df = sheets["industry_baselines"]
    filtered = baseline_df[
        (baseline_df["industry_group"].astype(str) == str(industry_group))
        & (baseline_df["period_key"].astype(str) == str(period_key))
    ].copy()
    if filtered.empty:
        return _rebuild_items()
    filtered["created_at_sort"] = pd.to_datetime(filtered["created_at"], errors="coerce")
    filtered = filtered.sort_values(["item_code", "created_at_sort"]).drop_duplicates(["item_code"], keep="last")
    filtered = filtered.drop(columns=["created_at_sort"])
    if "confidence_note" not in filtered.columns or filtered["confidence_note"].fillna("").astype(str).str.strip().eq("").all():
        return _rebuild_items()
    return filtered.to_dict("records")


def extract_company_tags(
    report_text: str,
    manual_overrides: Optional[List[str]] = None,
    workbook_path: Optional[str] = None,
) -> List[Dict[str, Any]]:
    sheets = load_industry_db(workbook_path)
    tag_rules = sheets["config_tag_rules"]
    text = (report_text or "").strip()
    sentences = [segment.strip() for segment in re.split(r"[.\n\r]+", text) if segment.strip()]
    candidates: List[Dict[str, Any]] = []

    for override in manual_overrides or []:
        raw = str(override).strip()
        if not raw:
            continue
        if ":" in raw:
            tag_code, tag_value = raw.split(":", 1)
        else:
            tag_code, tag_value = raw, "true"
        candidates.append(
            {
                "tag_code": tag_code,
                "tag_value": tag_value,
                "label": _tag_label(tag_code, tag_value),
                "evidence_text": "수동 태그 지정",
                "extractor_rule_id": "manual",
                "confidence_score": 0.95,
                "confidence_grade": "A",
                "source": "manual",
                "note": "",
            }
        )

    if text:
        for _, row in tag_rules.iterrows():
            if str(row.get("enabled", "Y")).strip().upper() == "N":
                continue
            keywords = [keyword.strip().lower() for keyword in str(row["keywords"]).split("|") if keyword.strip()]
            negatives = [keyword.strip().lower() for keyword in str(row["negative_keywords"]).split("|") if keyword.strip()]
            base = float(row["confidence_base"])
            for sentence in sentences:
                lowered = sentence.lower()
                hits = [keyword for keyword in keywords if keyword in lowered]
                if not hits:
                    continue
                score = base + min(0.12, 0.04 * max(len(hits) - 1, 0))
                if any(keyword in lowered for keyword in negatives):
                    score -= 0.08
                bounded = _clamp(score, 0.35, 0.95)
                candidates.append(
                    {
                        "tag_code": str(row["tag_code"]),
                        "tag_value": str(row["tag_value"]),
                        "label": str(row["label"]),
                        "evidence_text": sentence,
                        "extractor_rule_id": f"{row['tag_code']}:{row['tag_value']}",
                        "confidence_score": bounded,
                        "confidence_grade": _grade_confidence(bounded),
                        "source": "text",
                        "note": "",
                    }
                )

    winners: List[Dict[str, Any]] = []
    by_tag: Dict[str, List[Dict[str, Any]]] = {}
    for candidate in candidates:
        by_tag.setdefault(candidate["tag_code"], []).append(candidate)

    for tag_code, rows in by_tag.items():
        rows.sort(key=lambda row: row["confidence_score"], reverse=True)
        values = {row["tag_value"] for row in rows}
        if len(values) == 1:
            winners.append(rows[0])
            continue
        top = rows[0].copy()
        second = rows[1] if len(rows) > 1 else None
        if second and top["tag_value"] != second["tag_value"] and (top["confidence_score"] - second["confidence_score"]) <= 0.05:
            top["confidence_score"] = _clamp(top["confidence_score"] - 0.10, 0.35, 0.95)
            top["confidence_grade"] = _grade_confidence(top["confidence_score"])
            top["note"] = f"conflict:{second['tag_value']}"
        winners.append(top)

    winners.sort(key=lambda row: (row["tag_code"], row["tag_value"]))
    return winners


def append_company_tag_evidence(
    evaluation_id: str,
    company_name: str,
    industry_group: str,
    period_key: str,
    report_text_digest: str,
    tag_rows: List[Dict[str, Any]],
    workbook_path: Optional[str] = None,
) -> None:
    if not tag_rows:
        return
    resolved = ensure_workbook(workbook_path)
    sheets = load_industry_db(resolved)
    now = _utc_now()
    evidence_rows = []
    for row in tag_rows:
        evidence_rows.append(
            {
                "evaluation_id": evaluation_id,
                "company_name": company_name,
                "industry_group": industry_group,
                "period_key": period_key,
                "report_text_digest": report_text_digest,
                "tag_code": row["tag_code"],
                "tag_value": row["tag_value"],
                "label": row["label"],
                "evidence_text": row["evidence_text"],
                "extractor_rule_id": row["extractor_rule_id"],
                "confidence_score": row["confidence_score"],
                "confidence_grade": row["confidence_grade"],
                "source": row["source"],
                "note": row.get("note", ""),
                "created_at": now,
            }
        )
    sheets["company_tag_evidence"] = _append_rows(sheets["company_tag_evidence"], evidence_rows, "company_tag_evidence")
    _write_all_sheets(resolved, sheets)


def append_evaluation_run(
    company_name: str,
    industry_group: str,
    period_key: str,
    quote_digest: str,
    baseline_inputs_digest: str,
    report_text_digest: str,
    overall_judgement: str,
    overall_confidence_grade: str,
    restatement_seen: bool,
    workbook_path: Optional[str] = None,
) -> str:
    resolved = ensure_workbook(workbook_path)
    sheets = load_industry_db(resolved)
    evaluation_id = uuid.uuid4().hex
    row = {
        "evaluation_id": evaluation_id,
        "company_name": company_name,
        "industry_group": industry_group,
        "period_key": period_key,
        "quote_digest": quote_digest,
        "baseline_inputs_digest": baseline_inputs_digest,
        "report_text_digest": report_text_digest,
        "overall_judgement": overall_judgement,
        "overall_confidence_grade": overall_confidence_grade,
        "restatement_seen": bool(restatement_seen),
        "created_at": _utc_now(),
    }
    sheets["evaluation_runs"] = _append_rows(sheets["evaluation_runs"], [row], "evaluation_runs")
    _write_all_sheets(resolved, sheets)
    return evaluation_id


def _build_cli_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build or inspect industry financial baselines from raw statement files.")
    parser.add_argument("--raw-dir", default=None, help="폴더 전체 재무제표 파일(pdf/xlsx)을 스캔할 raw 경로. 기본값은 script 폴더의 financial_raw")
    parser.add_argument("--workbook", default=None, help="출력 workbook 경로. 기본값은 script 폴더의 industry_financial_db.xlsx")
    parser.add_argument("--period-label", default=DEFAULT_BATCH_PERIOD_LABEL, help="적재 입력 해시에 사용할 설명 문자열")
    parser.add_argument("--reset-workbook", action="store_true", help="기존 workbook 내용을 초기화하고 다시 적재")
    parser.add_argument("--discover-only", action="store_true", help="raw 폴더 파일 매칭만 출력하고 적재는 하지 않음")
    parser.add_argument("--file-mode", default="auto", choices=["auto", "xlsx", "pdf"], help="raw 폴더에서 읽을 파일 형식 선택. auto는 xlsx 우선, pdf 보조")
    parser.add_argument("--refresh-reference-only", action="store_true", help="Refresh official reference rows and estimated baselines only.")
    parser.add_argument("--suggest-account-map", action="store_true", help="Scan raw statement PDFs and print page/row candidates for extended account-map rows.")
    parser.add_argument("--suggest-account-map-codes", default="", help="Comma-separated normalized_account_code list for suggestion scan.")
    parser.add_argument("--estimate-period", default="2025", help="Target period for estimated baseline generation.")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = _build_cli_parser()
    args = parser.parse_args(argv)

    if args.discover_only:
        payload = discover_financial_raw_sets(args.raw_dir, file_mode=str(args.file_mode))
    elif args.suggest_account_map:
        target_codes = [code.strip() for code in str(args.suggest_account_map_codes or "").split(",") if code.strip()]
        payload = suggest_account_map_candidates(args.raw_dir, target_codes=target_codes or None)
    elif args.refresh_reference_only:
        resolved = ensure_workbook(args.workbook)
        payload = {
            "reference_refresh": refresh_official_reference_sheet(resolved, estimate_period=str(args.estimate_period)),
            "estimated_baselines": refresh_estimated_baselines(resolved, target_period=str(args.estimate_period)),
        }
    else:
        payload = ingest_financial_raw_directory(
            period_label=args.period_label,
            raw_dir=args.raw_dir,
            workbook_path=args.workbook,
            reset_workbook=bool(args.reset_workbook),
            estimate_period=str(args.estimate_period),
            file_mode=str(args.file_mode),
        )

    print(json.dumps(payload, ensure_ascii=False, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
