"""CSV 파일을 파싱하는 파서."""
import pandas as pd
from pathlib import Path


def parse_csv(file_path: str) -> dict:
    """CSV 파일을 파싱하여 데이터를 반환합니다."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    result = {
        "file_name": path.name,
        "file_type": "csv",
        "full_text": "",
    }

    # 인코딩 자동 감지 (UTF-8 → CP949 순서로 시도)
    for encoding in ["utf-8", "cp949", "euc-kr", "utf-8-sig"]:
        try:
            df = pd.read_csv(file_path, encoding=encoding)
            break
        except (UnicodeDecodeError, Exception):
            continue
    else:
        raise ValueError(f"파일 인코딩을 감지할 수 없습니다: {file_path}")

    df = df.dropna(how="all").fillna("")
    result["columns"] = list(df.columns)
    result["data"] = df.to_dict(orient="records")
    result["shape"] = df.shape
    result["full_text"] = df.to_string(index=False)

    return result
