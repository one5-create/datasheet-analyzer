"""Excel 파일(.xlsx, .xls)을 파싱하는 파서."""
import pandas as pd
from pathlib import Path


def parse_excel(file_path: str) -> dict:
    """Excel 파일을 파싱하여 시트별 데이터를 반환합니다."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    result = {
        "file_name": path.name,
        "file_type": "excel",
        "sheets": {},
        "full_text": "",
    }

    xl = pd.ExcelFile(file_path)
    text_parts = []

    for sheet_name in xl.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df = df.dropna(how="all").fillna("")

        result["sheets"][sheet_name] = {
            "columns": list(df.columns),
            "data": df.to_dict(orient="records"),
            "shape": df.shape,
        }

        # 텍스트로 변환
        sheet_text = f"[시트: {sheet_name}]\n"
        sheet_text += df.to_string(index=False)
        text_parts.append(sheet_text)

    result["full_text"] = "\n\n".join(text_parts)
    return result
