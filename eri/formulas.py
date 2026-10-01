"""정량 지수 회귀식과 ERI → 학년·단계 대응 (특허 식1, 식2, 도10)."""
from __future__ import annotations

from .textutil import round_half_up


def quantitative_elementary(Y: float, X1: float, coef: dict) -> float:
    """식1 (초등): 7.438 - 0.391×Y + 0.032×X1 + 0.016×(X1×Y)"""
    return coef["intercept"] + coef["Y"] * Y + coef["X1"] * X1 + coef["X1Y"] * X1 * Y


def quantitative_middle(X2: float, Z: float, K: float, coef: dict) -> float:
    """식2 (중등): -0.060×X2 + 0.145×Z + 0.110×K + 0.024×(K×Z) + 9.075"""
    return coef["intercept"] + coef["X2"] * X2 + coef["Z"] * Z + coef["K"] * K + coef["KZ"] * K * Z


def stage_table(grade_min: int = 3, grade_max: int = 9):
    """도10 '학습 단계별 텍스트 ERI 지수 범위'.

    초등(1~6학년): 기초 [g-0.5, g+0.5), 기본 [g, g+1), 심화 [g+0.5, g+1.5)
    중등(7학년~) : 기본 [g, g+1), 심화 [g+0.5, g+1.5)
    """
    rows = []
    for g in range(grade_min, grade_max + 1):
        if g <= 6:
            rows.append((g, "기초", g - 0.5, g + 0.5))
        rows.append((g, "기본", float(g), g + 1.0))
        rows.append((g, "심화", g + 0.5, g + 1.5))
    return rows


def stages_for(eri: float, grade_min: int = 3, grade_max: int = 9) -> str:
    table = stage_table(grade_min, grade_max)
    hits = [f"{g}학년 {s}" for g, s, lo, hi in table if lo <= eri < hi]
    if hits:
        return " / ".join(hits)
    if eri < table[0][2]:
        return f"{table[0][0]}학년 {table[0][1]} 미만"
    return f"{table[-1][0]}학년 {table[-1][1]} 초과"


def final_eri(quant: float, qual_mean: float | None) -> float:
    """ERI = 정량 지수(소수 첫째 자리) + 정성 지수 평균 [0213], [0230]."""
    q = round_half_up(quant, 1)
    return round_half_up(q + (qual_mean or 0.0), 1)
