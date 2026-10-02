"""문장 분리와 100어절 표본 추출 (특허 [0092], [0201], [0219]).

- 완전한 문장 단위로 뽑는다(문장 중간을 자르지 않는다).
- 도입부·중간부·끝부분 근처에서 각각 약 1/3씩 뽑는다.
- 표제(제목)가 있으면 첫 문장으로 포함한다.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Sentence:
    text: str
    tokens: list = field(default_factory=list)   # kiwi Token 목록 (문장 기준 위치)
    is_title: bool = False

    @property
    def eojeol(self) -> int:
        return len(self.text.split())


def split_sentences(kiwi, text: str) -> list[Sentence]:
    """kiwi 문장 분리기 사용. 줄바꿈도 문장 경계로 본다(마침표 없는 소제목 등)."""
    out = []
    for line in text.replace("\r", "\n").split("\n"):
        line = line.strip()
        if not line:
            continue
        for s in kiwi.split_into_sents(line):
            t = s.text.strip()
            if t:
                out.append(Sentence(t))
    return out


def select_sample(sentences: list[Sentence], target: int = 100, tolerance: float = 0.15,
                  title_eojeol: int = 0) -> list[Sentence]:
    """도입·중간·끝에서 문장 단위로 약 target 어절을 고른다 (제목 어절 수는 도입부에 포함)."""
    total = sum(s.eojeol for s in sentences) + title_eojeol
    if total <= target * (1 + tolerance) or len(sentences) <= 3:
        return list(sentences)
    quota = target / 3
    n = len(sentences)

    def take(indices, start_count=0):
        """할당량에 더 가까워질 때만 문장을 더한다 (첫 문장은 항상 포함)."""
        chosen, count = [], start_count
        for k in indices:
            size = sentences[k].eojeol
            if chosen and abs(count + size - quota) >= abs(count - quota):
                break
            chosen.append(k)
            count += size
            if count >= quota:
                break
        return chosen

    begin = take(range(n), title_eojeol)
    after_begin = begin[-1] + 1
    end = sorted(take(range(n - 1, after_begin - 1, -1)))
    lo, hi = after_begin, (end[0] if end else n)
    middle = []
    if lo < hi:
        # 도입부와 끝부분 사이 구간의 가운데에서 시작
        span = sum(s.eojeol for s in sentences[lo:hi])
        cum, k = 0, lo
        while k < hi - 1 and cum + sentences[k].eojeol <= (span - quota) / 2:
            cum += sentences[k].eojeol
            k += 1
        middle = take(range(k, hi))
    return [sentences[x] for x in begin + middle + end]
