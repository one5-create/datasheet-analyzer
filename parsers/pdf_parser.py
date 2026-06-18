"""PDF 파일에서 텍스트를 추출하는 파서."""
import pdfplumber
from pathlib import Path


def parse_pdf(file_path: str) -> dict:
    """PDF 파일을 파싱하여 텍스트와 테이블 데이터를 반환합니다."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    result = {
        "file_name": path.name,
        "file_type": "pdf",
        "pages": [],
        "full_text": "",
        "tables": [],
    }

    with pdfplumber.open(file_path) as pdf:
        all_text = []
        for i, page in enumerate(pdf.pages):
            page_text = page.extract_text() or ""
            all_text.append(page_text)
            result["pages"].append({
                "page_num": i + 1,
                "text": page_text,
            })

            # 테이블 추출
            tables = page.extract_tables()
            for table in tables:
                if table:
                    result["tables"].append({
                        "page_num": i + 1,
                        "data": table,
                    })

        result["full_text"] = "\n\n".join(all_text)

    return result
