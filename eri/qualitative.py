"""정성(질적) 평가 – 특허 [0082]~[0111], 청구항 1.

- 평가자(기본 3명)가 정량 지수에 대해 -3 ~ +3 범위의 보정값을 준다.
- 정성 지수 = 평가자 점수의 평균.
- 사전 평가(파일럿): 3·6·9학년 텍스트에 대해 평가자 2명씩 짝지어 일치도
  (같은 값을 준 대상 수 / 전체 대상 수)를 구하고, 그 평균이 0.9 이상이어야 본 평가를 한다.
"""
from __future__ import annotations

from itertools import combinations

CRITERIA = ("화제 친숙성", "구조 복잡성", "개념 수", "추론 수준", "문체의 친숙성", "텍스트 길이", "시각 자료와 편집 요소")


def parse_score(value, lo=-3.0, hi=3.0):
    """빈 값이면 None. 범위를 벗어나거나 숫자가 아니면 ValueError."""
    if value is None:
        return None
    if isinstance(value, float) and value != value:  # NaN
        return None
    s = str(value).strip().replace("＋", "+").replace("−", "-")
    if not s:
        return None
    v = float(s)
    if not lo <= v <= hi:
        raise ValueError(f"{v}은(는) {lo:+g} ~ {hi:+g} 범위를 벗어납니다.")
    return v


def mean_score(scores):
    vals = [s for s in scores if s is not None]
    return sum(vals) / len(vals) if vals else None


def pairwise_agreement(ratings_by_rater, tolerance: float = 0.0):
    """평가자별 점수 목록들 → (평균 일치도, {(i, j): 일치도}). 둘 다 점수를 준 대상만 비교한다."""
    pairs = {}
    for (i, a), (j, b) in combinations(enumerate(ratings_by_rater), 2):
        both = [(x, y) for x, y in zip(a, b) if x is not None and y is not None]
        if both:
            pairs[(i + 1, j + 1)] = sum(abs(x - y) <= tolerance for x, y in both) / len(both)
    if not pairs:
        return None, pairs
    return sum(pairs.values()) / len(pairs), pairs
