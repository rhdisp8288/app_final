from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parents[1]

MOVES = {
    "app_download.py": "app/main.py",
    "app_download_rev.py": "app/main_rev.py",
    "mecro.py": "scripts/mecro.py",

    "industry_financial_db.py": "app/analysis/industry_financial_db.py",
    "industry_research_report.py": "app/analysis/industry_research_report.py",
    "quote_adequacy.py": "app/analysis/quote_adequacy.py",

    "run_industry_db_test.py": "tests/run_industry_db_test.py",

    "buying_scenario_db.xlsx": "data/sample/buying_scenario_db.xlsx",
    "industry_financial_db.xlsx": "data/sample/industry_financial_db.xlsx",
    "standard_db.xlsx": "data/sample/standard_db.xlsx",
    "시뮬레이션용.xlsx": "data/sample/시뮬레이션용.xlsx",

    "financial_raw": "data/input/financial_raw",

    "_tmp_pdf_uploads": "data/temp/_tmp_pdf_uploads",
    "_tmp_uploaded_financial_refs": "data/temp/_tmp_uploaded_financial_refs",
    "tmp_pdf_pages": "data/temp/tmp_pdf_pages",

    "Pro_Buying_Agent_Context_Guide.pdf": "docs/guide/Pro_Buying_Agent_Context_Guide.pdf",
    "Pro_Buying_Agent_정밀진단_수정명세.docx": "docs/guide/Pro_Buying_Agent_정밀진단_수정명세.docx",
    "재무제표 기반 원가 분석 방법론 연구.docx": "docs/reference/재무제표 기반 원가 분석 방법론 연구.docx",
    "SCENARIO_CHECKLIST_app_download_rev.md": "docs/checklist/SCENARIO_CHECKLIST_app_download_rev.md",

    "_nh_report_page1.png": "assets/images/_nh_report_page1.png",
    "tmp_d_is_page1.png": "assets/images/tmp_d_is_page1.png",
}


def move_item(src_rel: str, dst_rel: str) -> None:
    src = ROOT / src_rel
    dst = ROOT / dst_rel

    if not src.exists():
        print(f"[SKIP] 없음: {src_rel}")
        return

    dst.parent.mkdir(parents=True, exist_ok=True)

    if dst.exists():
        print(f"[SKIP] 대상 이미 존재: {dst_rel}")
        return

    shutil.move(str(src), str(dst))
    print(f"[OK] {src_rel} -> {dst_rel}")


def main() -> None:
    # 기본 폴더 생성
    for folder in [
        "app",
        "app/analysis",
        "app/ui",
        "app/services",
        "app/utils",
        "data/input",
        "data/sample",
        "data/output",
        "data/temp",
        "docs/guide",
        "docs/reference",
        "docs/checklist",
        "assets/images",
        "tests",
        "scripts",
        ".codex",
    ]:
        (ROOT / folder).mkdir(parents=True, exist_ok=True)

    for src_rel, dst_rel in MOVES.items():
        move_item(src_rel, dst_rel)

    print("\n정리 완료. 이제 import 경로와 실행 경로를 점검하세요.")


if __name__ == "__main__":
    main()