"""ERI 계산 전 맞춤법·띄어쓰기 자동 교정 (인터넷 연결 없이 동작).

1) 띄어쓰기 – 붙여 써야 할 곳을 띄운 경우
   - 조사·어미·접미사 앞을 띄운 경우 (예: '가난 하고', '학교 에서', '먹었 다', '사람 들의')
   - 줄 맞춤 때문에 단어 한가운데에 들어간 공백 (예: '가 난하고', '투 박한', '명 암을')
   두 어절을 붙였을 때 형태소 분석기의 점수가 띄어 쓴 경우보다 확실히 좋을 때만 붙인다.
   '이 그림', '할 수', '큰 위로', '큰 일이'처럼 올바른 띄어쓰기는 건드리지 않는다.
2) 띄어쓰기 – 띄어 써야 할 곳을 붙인 경우 (예: '할수있다' → '할 수 있다', '그것을알' → '그것을 알')
   관형사형 어미 뒤 의존 명사, 조사·종결 어미 뒤 새 단어처럼 확실한 경우만 띄운다.
   '국어교육'(합성어)이나 '먹고있다'(보조 용언)처럼 붙여 써도 허용되는 경우는 그대로 둔다.
3) 맞춤법 – 자주 틀리는 표기 목록(됬다→됐다, 몇일→며칠, 오랫만에→오랜만에 …)에 있는 것만 고친다.

모든 교정은 Correction으로 기록되어 결과 엑셀 '교정 내역' 시트와 창에 표시된다.
일반 맞춤법 검사기(부산대·네이버 등)처럼 모든 오류를 잡지는 못하며, 고친 내용은 꼭 확인할 것.
"""
from __future__ import annotations

import re

from .cleanup import Correction, context
from .lexical import base_tag

DEPENDENT = ("J", "E", "XSN", "XSV", "XSA", "VCP")
MERGE_MARGIN = 2.0          # 조사·어미 앞 공백: 붙였을 때 점수가 이만큼 이상 좋아야 붙인다
SPLIT_WORD_MARGIN = 4.0     # 단어 한가운데 공백: 더 엄격하게 (올바른 띄어쓰기는 0~1.2 정도)
HANGUL = re.compile(r"[가-힣]")

# (틀린 표기 정규식, 바른 표기, 설명). 앞에 (?<![가-힣])가 붙은 것은 어절 첫머리에서만 고친다.
WORD_START = r"(?<![가-힣])"
SPELLING_RULES = [
    (r"됬", "됐", "'되었'의 준말은 '됐'"),
    (r"됀", "된", ""),
    (r"되요(?![가-힣])", "돼요", "'되어요'의 준말은 '돼요'"),
    (r"되서(?![가-힣])", "돼서", "'되어서'의 준말은 '돼서'"),
    (r"되야(?![가-힣])", "돼야", "'되어야'의 준말은 '돼야'"),
    (r"뵈요", "봬요", ""),
    (WORD_START + r"몇일", "며칠", ""),
    (r"오랫만", "오랜만", ""),
    (WORD_START + r"어떻해", "어떡해", ""),
    (WORD_START + r"왠일", "웬일", ""),
    (WORD_START + r"웬지", "왠지", ""),
    (WORD_START + r"금새(?![가-힣])", "금세", ""),
    (r"설레임", "설렘", ""),
    (r"역활", "역할", ""),
    (r"희안", "희한", ""),
    (WORD_START + r"일일히", "일일이", ""),
    (WORD_START + r"깨끗히", "깨끗이", ""),
    (WORD_START + r"곰곰히", "곰곰이", ""),
    (WORD_START + r"틈틈히", "틈틈이", ""),
    (WORD_START + r"번번히", "번번이", ""),
    (WORD_START + r"일찌기", "일찍이", ""),
    (WORD_START + r"가까히", "가까이", ""),
    (r"어의없", "어이없", ""),
    (WORD_START + r"구지(?![가-힣])", "굳이", ""),
    (WORD_START + r"갯수", "개수", ""),
    (WORD_START + r"촛점", "초점", ""),
    (WORD_START + r"도데체", "도대체", ""),
    (WORD_START + r"내노라", "내로라", ""),
    (WORD_START + r"짜집기", "짜깁기", ""),
    (WORD_START + r"설겆이", "설거지", ""),
    (WORD_START + r"메세지", "메시지", ""),
    (r"아니예요", "아니에요", ""),
    (r"(있|없|했|었|았|겠)슴(?![가-힣])", r"\1음", "명사형 어미는 '-음'"),
]
_COMPILED = [(re.compile(p), r, n) for p, r, n in SPELLING_RULES]


def _has_final_rieul(ch: str) -> bool:
    code = ord(ch) - 0xAC00
    return 0 <= code < 11172 and code % 28 == 8   # 받침 ㄹ


def fix_spelling(text: str, log=None) -> str:
    for rx, repl, note in _COMPILED:
        def sub(m):
            new = m.expand(repl)
            if log is not None:
                log.append(Correction("맞춤법", *context(text, m.start(), m.end(), new), note))
            return new
        text = rx.sub(sub, text)
    # '-ㄹ께' → '-ㄹ게' (할께, 갈께요)
    def sub_ge(m):
        new = m.group(1) + "게"
        if log is not None:
            log.append(Correction("맞춤법", *context(text, m.start(), m.end(), new), "'-ㄹ게'로 적음"))
        return new
    text = re.sub(r"([가-힣])께(?=요|(?![가-힣]))",
                  lambda m: sub_ge(m) if _has_final_rieul(m.group(1)) else m.group(0), text)
    return text


# ----------------------------------------------------------------------
# 띄어쓰기
# ----------------------------------------------------------------------
def _score(kiwi, s):
    return kiwi.analyze(s, top_n=1)[0][1]


def _should_merge(kiwi, a, b, prev, nxt) -> bool:
    if not (HANGUL.match(a[-1]) and HANGUL.match(b[0])):
        return False
    joined = a + b
    toks = kiwi.tokenize(joined)
    spanning = [t for t in toks if t.start < len(a) < t.end]
    if spanning:
        # 붙이면 한 형태소가 되는 경우 ('가 난하고', '채 색'). 합성 명사('오징어 게임', '프렉탈 기하학')는
        # 띄어 써도 맞으므로, 명사일 때는 끊긴 조각이 한 글자인 경우만 붙인다.
        t = spanning[0]
        pieces = (len(a) - t.start, t.end - len(a))
        if base_tag(t.tag).startswith("NN") and min(pieces) > 1:
            return False
        margin = SPLIT_WORD_MARGIN
    else:
        first = next((t for t in toks if t.start >= len(a)), None)
        if first is None or not base_tag(first.tag).startswith(DEPENDENT):
            return False
        margin = MERGE_MARGIN
    pre = prev + " " if prev else ""
    post = " " + nxt if nxt else ""
    return _score(kiwi, pre + joined + post) - _score(kiwi, pre + a + " " + b + post) >= margin


def _insert_points(kiwi, word):
    """한 어절 안에서 띄어 써야 하는 위치(문자 위치) 목록."""
    if len(word) < 3 or not HANGUL.search(word):
        return []
    toks = kiwi.tokenize(word)
    points = []
    for t, u in zip(toks, toks[1:]):
        lt, rt = base_tag(t.tag), base_tag(u.tag)
        if u.start <= 0 or u.start >= len(word) or u.start < t.end:
            continue
        ok = ((lt == "ETM" and rt == "NNB")
              or (lt == "NNB" and rt in ("VV", "VA") and t.form in ("수", "것", "줄", "리", "법", "적"))
              or (lt.startswith("J") and lt != "JKG" and rt[:2] in ("NN", "NP", "NR", "VV", "VA", "MA", "MM"))
              or (lt == "EF" and rt[:2] in ("NN", "NP", "VV", "VA", "MA", "MM", "IC")))
        if ok:
            points.append(u.start)
    if not points:
        return []
    spaced = kiwi.space(word, reset_whitespace=False)
    confirmed, count = set(), 0
    for i, ch in enumerate(spaced):
        if ch == " ":
            confirmed.add(count)
            continue
        count += 1
    return [p for p in points if p in confirmed]


def fix_spacing(text: str, kiwi, log=None) -> str:
    out_lines = []
    for line in text.split("\n"):
        words = line.split()
        if not words:
            out_lines.append("")
            continue
        # 붙이기
        merged = [words[0]]
        for i in range(1, len(words)):
            a, b = merged[-1], words[i]
            prev = merged[-2] if len(merged) > 1 else ""
            nxt = words[i + 1] if i + 1 < len(words) else ""
            if _should_merge(kiwi, a, b, prev, nxt):
                if log is not None:
                    ctx = (prev + " " if prev else "") + "{}" + (" " + nxt if nxt else "")
                    log.append(Correction("띄어쓰기(붙임)", ctx.format(a + " " + b), ctx.format(a + b),
                                          f"'{b}' 앞 띄어쓰기 삭제"))
                merged[-1] = a + b
            else:
                merged.append(b)
        # 띄우기
        final = []
        for i, w in enumerate(merged):
            pts = _insert_points(kiwi, w)
            if pts:
                new = "".join(w[s:e] + (" " if e < len(w) else "")
                              for s, e in zip([0] + pts, pts + [len(w)]))
                if log is not None:
                    prev = merged[i - 1] + " " if i else ""
                    nxt = " " + merged[i + 1] if i + 1 < len(merged) else ""
                    log.append(Correction("띄어쓰기(띄움)", prev + w + nxt, prev + new + nxt, "띄어 쓸 곳 추가"))
                w = new
            final.append(w)
        out_lines.append(" ".join(final))
    return "\n".join(out_lines)


def correct_text(text: str, kiwi, log=None, spelling=True, spacing=True) -> str:
    if spelling:
        text = fix_spelling(text, log)
    if spacing:
        text = fix_spacing(text, kiwi, log)
    return text
