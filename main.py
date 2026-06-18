"""Data Sheet 비교 분석 AI 에이전트 - CLI 진입점."""
import argparse
import sys
from pathlib import Path

from parsers.pdf_parser import parse_pdf
from parsers.excel_parser import parse_excel
from parsers.csv_parser import parse_csv
from agent.analyzer import DataSheetAnalyzer
from utils.reporter import print_report, save_report


SUPPORTED_EXTENSIONS = {
    ".pdf": parse_pdf,
    ".xlsx": parse_excel,
    ".xls": parse_excel,
    ".csv": parse_csv,
}


def parse_file(file_path: str) -> dict:
    """파일 확장자에 따라 적절한 파서를 선택합니다."""
    ext = Path(file_path).suffix.lower()
    parser = SUPPORTED_EXTENSIONS.get(ext)
    if parser is None:
        supported = ", ".join(SUPPORTED_EXTENSIONS.keys())
        raise ValueError(f"지원하지 않는 파일 형식: {ext}\n지원 형식: {supported}")
    return parser(file_path)


def main():
    parser = argparse.ArgumentParser(
        description="Data Sheet 비교 분석 AI 에이전트",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "files",
        nargs="+",
        help="비교할 Data Sheet 파일 경로 (2개 이상)\n예) main.py product_a.pdf product_b.xlsx",
    )
    parser.add_argument(
        "--save",
        action="store_true",
        help="분석 결과를 Markdown 파일로 저장",
    )
    parser.add_argument(
        "--output-dir",
        default=".",
        help="보고서 저장 디렉토리 (기본값: 현재 디렉토리)",
    )
    args = parser.parse_args()

    if len(args.files) < 2:
        print("오류: 비교하려면 최소 2개 이상의 파일이 필요합니다.", file=sys.stderr)
        sys.exit(1)

    # 파일 파싱
    parsed_files = []
    print("\n파일 파싱 중...")
    for file_path in args.files:
        try:
            print(f"  ✓ {file_path}")
            data = parse_file(file_path)
            parsed_files.append(data)
        except (FileNotFoundError, ValueError) as e:
            print(f"  ✗ {file_path}: {e}", file=sys.stderr)
            sys.exit(1)

    # AI 분석
    print("\nAI 분석 중...")
    try:
        analyzer = DataSheetAnalyzer()
        result = analyzer.analyze(parsed_files)
    except ValueError as e:
        print(f"\n오류: {e}", file=sys.stderr)
        sys.exit(1)

    # 결과 출력
    print_report(result)

    # 파일 저장
    if args.save:
        saved_path = save_report(result, args.output_dir)
        print(f"\n보고서 저장 완료: {saved_path}")


if __name__ == "__main__":
    main()
