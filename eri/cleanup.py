"""지문 정리: PDF·한글(HWP)에서 복사할 때 생긴 줄바꿈과 문제용 기호를 바로잡는다.

PDF에서 본문을 복사하면 화면의 줄 끝마다 줄바꿈이 들어가고, 단어 중간에서 끊기기도 한다
(예: '어\\n둡고', '가난\\n하고'). 프로그램은 줄바꿈을 문장 경계로 보므로 그대로 두면
문장 수(Y)가 늘고 문장 복잡도(K)가 줄며, '둡' 같은 가짜 단어가 생긴다.

규칙
- 줄이 마침표·물음표·느낌표(닫는 따옴표 포함)로 끝나면 진짜 줄바꿈으로 둔다.
- 다른 줄보다 눈에 띄게 짧은 줄(소제목, '-1885년 4월' 같은 출처 표시)도 그대로 둔다.
- 나머지는 다음 줄과 잇는다. 이을 때 띄어 쓸지 붙여 쓸지는 형태소 분석기의
  띄어쓰기 교정으로 이음매 부분만 판단한다 ('어'+'둡고' → '어둡고', '얻은'+'양식을' → '얻은 양식을').
- ㉠ ⓐ ① 같은 문제용 기호는 지운다.
"""
from __future__ import annotations

import re
from statistics import median

SENTENCE_END = re.compile(r"[.?!。…]['\"’”」』)\]]*$")
SOFT_END = re.compile(r"[,;:·」』’”)\]]$")
CIRCLED = re.compile(r"[①-⓿❶-➓㉑-㉟㉠-㉿㊱-㊿]")
SHORT_LINE_RATIO = 0.6


def remove_markers(text: str) -> str:
    text = CIRCLED.sub("", text)
    return re.sub(r"[ \t 　]{2,}", " ", text)


def _needs_space(kiwi, left: str, right: str, prev: str = "", nxt: str = "") -> bool:
    """left와 right를 이을 때 띄어 써야 하는지 형태소 분석기로 판단한다."""
    if SOFT_END.search(left):
        return True
    if kiwi is None:
        return True
    head = (prev + " " if prev else "") + left
    snippet = head + right + (" " + nxt if nxt else "")
    spaced = kiwi.space(snippet, reset_whitespace=False)
    # 이음매(공백을 뺀 문자 수 기준) 바로 앞에 공백이 생겼는지 확인
    target = len(head.replace(" ", ""))
    count = 0
    for i, ch in enumerate(spaced):
        if ch == " ":
            continue
        if count == target:
            return i > 0 and spaced[i - 1] == " "
        count += 1
    return True


def join_wrapped_lines(text: str, kiwi=None) -> str:
    """빈 줄로 나뉜 문단마다 PDF식 줄바꿈을 복원한다."""
    out_blocks = []
    for block in re.split(r"\n\s*\n", text.replace("\r\n", "\n").replace("\r", "\n")):
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if len(lines) < 2:
            out_blocks.append("\n".join(lines))
            continue
        typical = median(len(l) for l in lines)
        merged = [lines[0]]
        for i in range(1, len(lines)):
            cur, nxt_line = merged[-1], lines[i]
            last_line = lines[i - 1]
            if SENTENCE_END.search(cur) or len(last_line) < typical * SHORT_LINE_RATIO:
                merged.append(nxt_line)
                continue
            left_words, right_words = cur.split(), nxt_line.split()
            left, right = left_words[-1], right_words[0]
            prev = left_words[-2] if len(left_words) > 1 else ""
            after = right_words[1] if len(right_words) > 1 else ""
            sep = " " if _needs_space(kiwi, left, right, prev, after) else ""
            merged[-1] = cur + sep + nxt_line
        out_blocks.append("\n".join(merged))
    return "\n\n".join(b for b in out_blocks if b)


def clean_text(text: str, kiwi=None, join_lines: bool = True, strip_markers: bool = True) -> str:
    if strip_markers:
        text = remove_markers(text)
    if join_lines:
        text = join_wrapped_lines(text, kiwi)
    return text
