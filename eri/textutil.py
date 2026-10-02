"""공통 유틸리티."""
from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

ENCODINGS = ("utf-8-sig", "utf-8", "cp949", "euc-kr", "utf-16")


def read_text(path: str | Path) -> str:
    """인코딩을 자동으로 판별해 텍스트 파일을 읽는다 (UTF-8, CP949/EUC-KR 메모장 파일 등)."""
    raw = Path(path).read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    for enc in ENCODINGS:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def round_half_up(value: float, digits: int = 1) -> float:
    """소수점 (digits+1)째 자리에서 반올림 (특허 [0209]). 파이썬 round()의 짝수 반올림을 피한다."""
    q = Decimal(1).scaleb(-digits)
    return float(Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP))


def normalize_header(name) -> str:
    """엑셀 열 이름 비교용: 공백 제거."""
    return "".join(str(name).split()) if name is not None else ""
