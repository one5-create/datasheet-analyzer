"""분석 결과를 Markdown 및 콘솔 리포트로 출력하는 유틸리티."""
from datetime import datetime
from pathlib import Path
from colorama import Fore, Back, Style, init

init(autoreset=True)


def print_report(result: dict) -> None:
    """분석 결과를 콘솔에 출력합니다."""
    if "raw_response" in result:
        print(Fore.WHITE + "\n[분석 결과]\n")
        print(result["raw_response"])
        return

    sep = "═" * 62
    print(Fore.CYAN + Style.BRIGHT + f"\n{sep}")
    print(Fore.CYAN + Style.BRIGHT + "  DATA SHEET 비교 분석 보고서")
    print(Fore.CYAN + Style.BRIGHT + f"{sep}\n")

    # 요약
    print(Fore.YELLOW + Style.BRIGHT + "▌ 요약")
    print(Fore.WHITE + f"  {result.get('summary', '-')}\n")

    # 제품별 정보
    colors = [Fore.GREEN, Fore.MAGENTA, Fore.BLUE, Fore.RED]
    products = result.get("products", [])
    for i, prod in enumerate(products):
        color = colors[i % len(colors)]
        print(color + Style.BRIGHT + f"┌{'─' * 50}┐")
        print(color + Style.BRIGHT + f"  📄 {prod.get('file_name', '-')}")
        print(color + Style.BRIGHT + f"  🏷  {prod.get('product_name', '-')}")

        specs = prod.get("key_specs", {})
        if specs:
            print(Fore.WHITE + Style.BRIGHT + "\n  【주요 사양】")
            for k, v in specs.items():
                print(Fore.WHITE + f"    • {k}: " + Fore.CYAN + f"{v}")

        perf = prod.get("performance", {})
        if perf:
            print(Fore.WHITE + Style.BRIGHT + "\n  【성능 지표】")
            for k, v in perf.items():
                print(Fore.WHITE + f"    • {k}: " + Fore.CYAN + f"{v}")

        highlights = prod.get("highlights", [])
        if highlights:
            print(Fore.WHITE + Style.BRIGHT + "\n  【핵심 특징】")
            for h in highlights:
                print(Fore.GREEN + f"    ✓ {h}")
        print(color + f"└{'─' * 50}┘\n")

    # 비교 테이블
    table = result.get("comparison_table", [])
    if table:
        print(Fore.YELLOW + Style.BRIGHT + "▌ 상세 비교 테이블\n")
        file_names = list(table[0]["values"].keys()) if table else []
        col_w = 22
        header = f"  {'항목':<{col_w}}" + "".join(f"{fn[:col_w-2]:<{col_w}}" for fn in file_names)
        print(Fore.WHITE + Style.BRIGHT + header)
        print(Fore.WHITE + "  " + "─" * (col_w * (len(file_names) + 1)))
        for row in table:
            cat = row["category"]
            vals = [str(row["values"].get(fn, "-")) for fn in file_names]
            line = Fore.WHITE + f"  {cat:<{col_w}}"
            for v in vals:
                line += Fore.CYAN + f"{v:<{col_w}}"
            print(line)
        print()

    # 종합 의견
    if result.get("recommendation"):
        print(Fore.YELLOW + Style.BRIGHT + "▌ 종합 의견")
        print(Fore.WHITE + f"  {result['recommendation']}\n")

    print(Fore.CYAN + Style.BRIGHT + sep)


def save_report(result: dict, output_dir: str = ".") -> str:
    """분석 결과를 Markdown 파일로 저장합니다."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = Path(output_dir) / f"report_{timestamp}.md"

    lines = ["# Data Sheet 비교 분석 보고서\n"]
    lines.append(f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

    if "raw_response" in result:
        lines.append(result["raw_response"])
    else:
        lines.append(f"## 요약\n{result.get('summary', '-')}\n\n")

        for prod in result.get("products", []):
            lines.append(f"## {prod.get('product_name', prod.get('file_name', '-'))}\n")
            lines.append(f"**파일**: {prod.get('file_name', '-')}\n\n")

            if prod.get("key_specs"):
                lines.append("### 주요 사양\n")
                for k, v in prod["key_specs"].items():
                    lines.append(f"- **{k}**: {v}\n")
                lines.append("\n")

            if prod.get("performance"):
                lines.append("### 성능 지표\n")
                for k, v in prod["performance"].items():
                    lines.append(f"- **{k}**: {v}\n")
                lines.append("\n")

            if prod.get("highlights"):
                lines.append("### 핵심 특징\n")
                for h in prod["highlights"]:
                    lines.append(f"- {h}\n")
                lines.append("\n")

        table = result.get("comparison_table", [])
        if table:
            file_names = list(table[0]["values"].keys())
            lines.append("## 비교 테이블\n\n")
            lines.append("| 항목 | " + " | ".join(file_names) + " |\n")
            lines.append("|---" * (len(file_names) + 1) + "|\n")
            for row in table:
                vals = " | ".join(str(row["values"].get(fn, "-")) for fn in file_names)
                lines.append(f"| {row['category']} | {vals} |\n")
            lines.append("\n")

        if result.get("recommendation"):
            lines.append(f"## 종합 의견\n{result['recommendation']}\n")

    with open(output_path, "w", encoding="utf-8") as f:
        f.writelines(lines)

    return str(output_path)
