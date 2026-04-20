from pathlib import Path

from app.analysis.industry_financial_db import (
    ensure_workbook,
    get_baseline_items,
    ingest_industry_statements,
    load_industry_db,
)


BASE_DIR = Path(__file__).resolve().parent
INDUSTRY_GROUP = "C"
PDF_PREFIX = "\uc81c\uc870\uc5c5(C)"
PERIOD_LABEL = "2024 \uc0ac\uc5c5\ubcf4\uace0\uc11c"
PERIOD_KEY = "2024"
ITEM_ORDER = ["material", "labor", "expense", "gna", "profit", "rnd"]
KEY_ACCOUNT_CODES = [
    "is_revenue",
    "is_cogs",
    "is_gna",
    "is_rnd_research",
    "is_rnd_development",
    "is_operating_profit",
    "cost_material",
    "cost_labor",
    "cost_expense",
    "bs_total_assets",
    "bs_current_assets",
    "cf_operating_cash_flow",
]


def _find_pdf(*filenames: str) -> str | None:
    for filename in filenames:
        path = BASE_DIR / filename
        if path.exists():
            return str(path)
    return None


def _build_pdf_paths() -> dict[str, str]:
    pdf_paths = {
        "BS": _find_pdf(
            f"{PDF_PREFIX}_\uc7ac\ubb34\uc0c1\ud0dc\ud45c.pdf",
            f"{PDF_PREFIX}_\uc7ac\ubb34\uc81c\ud45c.pdf",
        ),
        "IS": _find_pdf(
            f"{PDF_PREFIX}_\uc190\uc775\uacc4\uc0b0\uc11c.pdf",
            f"{PDF_PREFIX}_\uc190\uc778\uacc4\uc0b0\uc11c.pdf",
        ),
        "CF": _find_pdf(f"{PDF_PREFIX}_\ud604\uae08\ud750\ub984\ud45c.pdf"),
        "COST": _find_pdf(f"{PDF_PREFIX}_\uc81c\uc870\uc6d0\uac00\uba85\uc138\uc11c.pdf"),
    }

    missing = [key for key in ("BS", "IS", "CF") if not pdf_paths[key]]
    if missing:
        raise FileNotFoundError(f"Required PDF not found: {', '.join(missing)}")

    return {key: value for key, value in pdf_paths.items() if value}


def _format_amount(value: object) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "-"
    return f"{number:,.0f}"


def _latest_statement_rows(workbook_path: str):
    sheets = load_industry_db(workbook_path)
    statement_rows = sheets["statement_rows"]
    filtered = statement_rows[
        (statement_rows["industry_group"].astype(str) == INDUSTRY_GROUP)
        & (statement_rows["period_key"].astype(str) == PERIOD_KEY)
    ].copy()
    if filtered.empty:
        return filtered
    filtered = filtered.sort_values(["statement_type", "normalized_account_code", "ingested_at"])
    filtered = filtered.drop_duplicates(["statement_type", "normalized_account_code"], keep="last")
    return filtered.sort_values(["statement_type", "page_number", "row_index"])


def _print_ingest_result(result: dict[str, object], pdf_paths: dict[str, str]) -> None:
    print("[INGEST RESULT]")
    print(f"industry_group: {result.get('industry_group')}")
    print(f"period_keys: {result.get('period_keys')}")
    print(f"rows_written: {result.get('rows_written')}")
    print(f"noop: {result.get('noop')}")
    print(f"restatement_detected: {result.get('restatement_detected')}")
    print("pdf_paths:")
    for key in ("BS", "IS", "CF", "COST"):
        print(f"  {key}: {pdf_paths.get(key, '-')}")


def _print_statement_rows(workbook_path: str) -> None:
    latest_rows = _latest_statement_rows(workbook_path)
    print("\n[LATEST STATEMENT ROWS]")
    if latest_rows.empty:
        print("No statement rows found.")
        return

    latest_rows = latest_rows[
        latest_rows["normalized_account_code"].astype(str).isin(KEY_ACCOUNT_CODES)
    ].copy()
    if latest_rows.empty:
        print("No key account rows found.")
        return

    for _, row in latest_rows.iterrows():
        print(
            f"{row['statement_type']:<4} "
            f"{str(row['normalized_account_code']):<24} "
            f"amount={_format_amount(row['amount']):>12} "
            f"page={row['page_number']} "
            f"row={row['row_index']} "
            f"raw={row['raw_account_name']}"
        )


def _print_baseline_items(workbook_path: str) -> None:
    baseline_items = get_baseline_items(INDUSTRY_GROUP, PERIOD_KEY, workbook_path)
    baseline_items.sort(
        key=lambda row: ITEM_ORDER.index(row["item_code"]) if row["item_code"] in ITEM_ORDER else len(ITEM_ORDER)
    )

    print("\n[BASELINE ITEMS]")
    for row in baseline_items:
        print(
            f"{row['item_code']:<8} "
            f"base={float(row['base_pct']):7.3f}% "
            f"range=({float(row['low_pct']):7.3f}, {float(row['high_pct']):7.3f}) "
            f"basis={row['basis_type']:<8} "
            f"conf={row['confidence_grade']}({float(row['confidence_score']):.2f}) "
            f"cost={'Y' if row['cost_presence_flag'] else 'N'}"
        )

    proxy_items = [row["item_code"] for row in baseline_items if str(row.get("basis_type")) != "observed"]
    if proxy_items:
        print("\n[CHECK]")
        print(f"Proxy or estimated basis used for: {', '.join(proxy_items)}")


workbook = ensure_workbook()
pdf_paths = _build_pdf_paths()

result = ingest_industry_statements(
    industry_group=INDUSTRY_GROUP,
    period_label=PERIOD_LABEL,
    pdf_paths=pdf_paths,
    workbook_path=workbook,
)

_print_ingest_result(result, pdf_paths)
_print_statement_rows(workbook)
_print_baseline_items(workbook)
