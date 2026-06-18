"""Data Sheet 비교 분석 AI 에이전트 - 대화형 실행 진입점."""
import sys
import os
from pathlib import Path

import questionary
from colorama import init, Fore, Style

init(autoreset=True)

# 프로젝트 루트를 sys.path에 추가
sys.path.insert(0, str(Path(__file__).parent))

from parsers.pdf_parser import parse_pdf
from parsers.excel_parser import parse_excel
from parsers.csv_parser import parse_csv
from agent.analyzer import DataSheetAnalyzer
from utils.reporter import print_report, save_report

SUPPORTED_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".csv"}
DATA_DIR = Path(__file__).parent / "data"


def find_datasheet_files() -> list[Path]:
    """data/ 폴더 및 현재 폴더에서 지원 파일을 탐색합니다."""
    found = []
    search_dirs = [DATA_DIR, Path(__file__).parent]
    for d in search_dirs:
        if d.exists():
            for f in sorted(d.iterdir()):
                if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS:
                    if f not in found:
                        found.append(f)
    return found


def parse_file(path: Path) -> dict:
    ext = path.suffix.lower()
    parsers = {
        ".pdf": parse_pdf,
        ".xlsx": parse_excel,
        ".xls": parse_excel,
        ".csv": parse_csv,
    }
    return parsers[ext](str(path))


def print_banner():
    print(Fore.CYAN + Style.BRIGHT + """
╔══════════════════════════════════════════════════╗
║     📊 Data Sheet 비교 분석 AI 에이전트          ║
║        Powered by AI                            ║
╚══════════════════════════════════════════════════╝
""")


def select_files(available: list[Path]) -> list[Path]:
    """파일 선택 메뉴를 표시합니다."""
    if not available:
        print(Fore.RED + f"\n  ✗ '{DATA_DIR}' 폴더에 분석 가능한 파일이 없습니다.")
        print(Fore.YELLOW + f"  → PDF, Excel, CSV 파일을 '{DATA_DIR}' 폴더에 넣어주세요.\n")
        return []

    choices = [
        questionary.Choice(
            title=f"  {f.name}  ({f.suffix.upper()[1:]})",
            value=f
        )
        for f in available
    ]

    print(Fore.WHITE + Style.BRIGHT + "  비교할 파일을 선택하세요 (스페이스바로 선택, 엔터로 확인):\n")
    selected = questionary.checkbox(
        "",
        choices=choices,
        style=questionary.Style([
            ("checkbox-selected", "fg:#00ff88 bold"),
            ("highlighted", "fg:#00aaff bold"),
            ("pointer", "fg:#00aaff bold"),
        ]),
        instruction="  (최소 2개 이상 선택)",
    ).ask()

    return selected or []


def ask_save() -> bool:
    return questionary.confirm(
        "  보고서를 Markdown 파일로 저장할까요?",
        default=True,
        style=questionary.Style([("answer", "fg:#00ff88 bold")])
    ).ask()


def main():
    print_banner()

    # 파일 탐색
    available = find_datasheet_files()
    selected = select_files(available)

    if len(selected) < 2:
        print(Fore.RED + "\n  ✗ 최소 2개 이상의 파일을 선택해야 합니다.\n")
        sys.exit(1)

    print()
    print(Fore.GREEN + Style.BRIGHT + "  선택된 파일:")
    for f in selected:
        print(Fore.GREEN + f"    ✓ {f.name}")

    # 파일 파싱
    print(Fore.CYAN + "\n  파일 파싱 중...")
    parsed_files = []
    for path in selected:
        try:
            data = parse_file(path)
            parsed_files.append(data)
            print(Fore.GREEN + f"    ✓ {path.name}")
        except Exception as e:
            print(Fore.RED + f"    ✗ {path.name}: {e}")
            sys.exit(1)

    # 저장 여부 확인
    print()
    do_save = ask_save()

    # AI 분석
    print(Fore.CYAN + Style.BRIGHT + "\n  AI 분석 중... (잠시 기다려주세요)")
    try:
        analyzer = DataSheetAnalyzer()
        result = analyzer.analyze(parsed_files)
    except ValueError as e:
        print(Fore.RED + f"\n  ✗ {e}\n")
        sys.exit(1)

    # 결과 출력
    print()
    print_report(result)

    # 저장
    if do_save:
        output_dir = str(Path(__file__).parent)
        saved_path = save_report(result, output_dir)
        print(Fore.GREEN + Style.BRIGHT + f"\n  ✓ 보고서 저장 완료: {saved_path}\n")

    # 다시 실행 여부
    print()
    again = questionary.confirm("  다른 파일을 추가로 분석할까요?", default=False).ask()
    if again:
        main()
    else:
        print(Fore.CYAN + "\n  분석을 종료합니다. 감사합니다!\n")


if __name__ == "__main__":
    main()
