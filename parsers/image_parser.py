"""이미지 파일(회로도 등)을 읽는 파서."""
from pathlib import Path

MIME_MAP = {
    ".png":  "image/png",
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".gif":  "image/gif",
    ".pdf":  "application/pdf",
}


def parse_image(file_path: str) -> dict:
    """이미지 파일을 읽어 바이트와 메타데이터를 반환합니다."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    ext = path.suffix.lower()
    mime = MIME_MAP.get(ext, "image/png")

    with open(file_path, "rb") as f:
        data = f.read()

    return {
        "file_name": path.name,
        "file_type": "image",
        "mime_type": mime,
        "image_bytes": data,
    }
