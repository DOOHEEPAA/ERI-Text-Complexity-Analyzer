"""문장 복잡도(K) 산정 – 특허 [0151]~[0167], 도9.

문장 하나의 점수 = Σ(이어진 문장의 각 절: 기본 형식 + 첨가조건①) + 첨가조건②

기본 형식      주술 1 / 주목술·주보술 2 / 주목보술 3
               보어는 '되다·아니다' 앞의 '이/가'와 필수적 부사어까지 포함 [0158]
               (예: '수요량은 증가하게 된다' → 주보술 2점)
첨가조건①     절마다 관형어·부사어·독립어 개수: 1~3개 +1, 4~6개 +2, 7개 이상 +4
               관형사형 어미(-ㄴ/-ㄹ)·부사형 어미(-게) 활용형은 제외 [0160]
               명사가 나란히 놓여 꾸미는 경우(예: '대중 문화의')는 관형어 [0161]
첨가조건②     내포절(명사절·관형절·부사절·인용절·서술절) 1개당 +3,
               한 단어 관형절 +2, 내포절 안의 내포절 +1, 6개 이상이면 18점 [0164]~[0166]
               부사절 어미는 '-도록, -(아)서, -게'로 한정, 나머지 연결어미는 이어진 문장 [0167]

형태소 분석 결과만으로 문장 성분을 판정하는 규칙 기반 근사치이므로,
결과 엑셀의 '문장 복잡도 상세' 시트에서 절 나누기와 점수를 확인할 수 있게 했다.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .lexical import base_tag

PRED_START = {"VV", "VA", "VCP", "VCN", "XSV", "XSA", "VX"}
NOUNS = {"NNG", "NNP", "NNB", "NR", "NP", "XR", "SN", "SL", "SH"}
ADVERBIAL_EC = {"도록", "어서", "아서", "여서", "서", "게"}
QUOTE_EC = {"다고", "라고", "자고", "냐고", "느냐고", "으냐고", "ᆫ다고", "는다고", "으라고", "이라고", "ᆫ다며", "다며", "라며"}
SERIAL_EC = {"어", "아", "여"}
COPULA_LIKE = {"되", "하"}           # '-게 되다/하다' → 필수적 부사어(보어)
COMPARE_PRED = {"같", "다르", "비슷하", "닮", "똑같"}
SUBJECT_JX = {"은", "는"}
# 굳어진 표현: 절이 아니라 부사어/관형어 하나로 센다 (어떻게, 이렇게 / 이러한, 그런 …)
LEXICAL_ADV_STEMS = {"어떻", "이렇", "그렇", "저렇", "아무렇"}
LEXICAL_DET_STEMS = {"이러하", "그러하", "저러하", "어떠하", "이렇", "그렇", "저렇", "어떻", "아무러하"}


@dataclass
class Group:
    pred: int              # 서술어(용언·서술격 조사·접미사) 토큰 위치
    head: int              # 서술어가 들어 있는 어절의 첫 토큰 위치 (예: '증가하게'의 '증가')
    end: int               # 마지막 토큰 위치 (어미 포함)
    arg_start: int         # 이 서술어에 딸린 성분이 시작하는 위치
    kind: str = "final"    # final / connective / serial / 관형절 / 명사절 / 부사절 / 인용절 / lexical
    ending: str = ""
    gehada: bool = False   # '-게 되다/하다'


@dataclass
class UnitScore:
    text: str
    base: int
    base_label: str
    modifiers: int
    add1: int

    @property
    def score(self):
        return self.base + self.add1


@dataclass
class SentenceScore:
    text: str
    units: list = field(default_factory=list)
    embedded: list = field(default_factory=list)   # (종류, 서술어, 점수)
    add2: int = 0

    @property
    def score(self):
        return sum(u.score for u in self.units) + self.add2

    def describe(self):
        u = " + ".join(f"[{x.text}] {x.base_label}{x.base}+수식{x.modifiers}개→{x.add1}" for x in self.units)
        e = ", ".join(f"{k}({p})+{s}" for k, p, s in self.embedded)
        return u, e


def _next_content(toks, i):
    while i < len(toks) and base_tag(toks[i].tag) in ("SF", "SP", "SS", "SSO", "SSC", "SE", "SO", "SW"):
        i += 1
    return i


def find_groups(toks) -> list[Group]:
    tags = [base_tag(t.tag) for t in toks]
    n = len(toks)
    groups, i, arg_start = [], 0, 0
    while i < n:
        if tags[i] not in PRED_START:
            i += 1
            continue
        head = i
        while head - 1 >= arg_start and toks[head - 1].word_position == toks[i].word_position:
            head -= 1
        g = Group(pred=i, head=head, end=i, arg_start=arg_start)
        j = i + 1
        while j < n:
            tg = tags[j]
            if tg in ("EP",) or tg in PRED_START and tags[j - 1] not in ("EC",):
                j += 1
                continue
            if tg == "EC":
                nxt = j + 1
                if nxt < n and tags[nxt] == "VX":            # 보조 용언 (-지 않다, -어 있다, -고 싶다)
                    j = nxt + 1
                    continue
                if toks[j].form == "게" and nxt < n and tags[nxt] in ("VV", "VX") and toks[nxt].form in COPULA_LIKE:
                    g.gehada = True
                    j = nxt + 1
                    continue
                g.ending = toks[j].form
                if g.ending in QUOTE_EC or (g.ending in ("라", "이라") and tags[i] == "VCP"):  # '위헌이라 판단하다'
                    g.kind = "인용절"
                elif g.ending in ADVERBIAL_EC:
                    g.kind = "부사절"
                else:
                    g.kind = "connective"
                g.end = j
                break
            if tg in ("ETM", "ETN", "EF"):
                g.ending = toks[j].form
                g.kind = {"ETM": "관형절", "ETN": "명사절"}.get(tg, "final")
                if tg == "EF":
                    k = _next_content(toks, j + 1)
                    if k < n and tags[k] == "JKQ":
                        g.kind = "인용절"
                g.end = j
                break
            g.end = j - 1
            break
        else:
            g.end = n - 1
        stem = toks[g.pred].form
        if g.head == g.pred and ((g.kind == "부사절" and g.ending == "게" and stem in LEXICAL_ADV_STEMS)
                                 or (g.kind == "관형절" and stem in LEXICAL_DET_STEMS)):
            g.kind = "lexical"
        groups.append(g)
        i = g.end + 1
        arg_start = i
    # 연속 동사(모여 거래하는, 깨어나 뛰어다닐): '-어/아' 뒤에 바로 다음 서술어가 오면 나누지 않는다.
    for a, b in zip(groups, groups[1:]):
        if a.kind == "connective" and a.ending in SERIAL_EC and b.arg_start == b.head:
            a.kind = "serial"
    return groups


def score_sentence(toks, text: str, cfg) -> SentenceScore:
    tags = [base_tag(t.tag) for t in toks]
    n = len(toks)
    groups = find_groups(toks)
    result = SentenceScore(text)
    if n == 0:
        return result

    # --- 성분 귀속: 목적어·보어는 바로 뒤 서술어에 딸린다. 형용사는 목적어를 갖지 않으므로 다음으로 넘긴다.
    comp_tokens = set()
    has_obj, has_comp = {}, {}
    carry_obj = False
    for gi, g in enumerate(groups):
        span = range(g.arg_start, g.head)
        obj = carry_obj or any(tags[k] == "JKO" for k in span)
        stem = tags[g.pred]
        # 형용사는 목적어를 갖지 않는다. 'A를 B라고 하다'의 목적어는 바깥 서술어 '하다'의 것이다.
        if obj and ((stem in ("VA", "XSA") and g.kind not in ("final", "connective"))
                    or (g.kind == "인용절" and stem == "VCP")):
            carry_obj, obj = True, False
        else:
            carry_obj = False
        comp = g.gehada or any(tags[k] == "JKC" for k in span)
        last = g.head - 1
        # '물이 얼음으로 되다', '사과와 같다'
        if last >= g.arg_start and tags[last] == "JKB":
            f = toks[last].form
            pf = toks[g.pred].form
            if (f in ("로", "으로") and pf == "되") or (f in ("와", "과", "하고") and pf in COMPARE_PRED):
                comp = True
                comp_tokens.add(last)
        if last >= g.arg_start and tags[last] == "JC" and toks[last].form in ("와", "과") and toks[g.pred].form in COMPARE_PRED:
            comp = True
            comp_tokens.add(last)
        has_obj[gi], has_comp[gi] = obj, comp

    # --- 이어진 문장의 절(단위)로 나누기
    units, start, last_gi = [], 0, []
    for gi, g in enumerate(groups):
        last_gi.append(gi)
        if g.kind == "connective":
            units.append((start, g.end, last_gi)); start, last_gi = g.end + 1, []
    if start < n:
        if units and all(tags[k].startswith("S") for k in range(start, n)):
            s0, _, g0 = units[-1]          # 연결어미 뒤에 문장부호만 남은 경우 앞 절에 붙인다
            units[-1] = (s0, n - 1, g0 + last_gi)
        else:
            units.append((start, n - 1, last_gi))

    for (s, e, gis) in units:
        main = [gi for gi in gis if groups[gi].kind in ("final", "connective", "serial")]
        if main:
            gi = main[-1]
            base = 1 + int(has_obj[gi]) + int(has_comp[gi])
            label = {(0, 0): "주술", (1, 0): "주목술", (0, 1): "주보술", (1, 1): "주목보술"}[(int(has_obj[gi]), int(has_comp[gi]))]
        else:
            base, label = 0, "서술어없음"
        mods = _count_modifiers(toks, tags, s, e, comp_tokens)
        mods += sum(1 for gi in gis if groups[gi].kind == "lexical")
        add1 = 0 if mods == 0 else 1 if mods <= 3 else 2 if mods <= 6 else 4
        unit_text = " ".join(_surface_words(text, toks[s:e + 1]))
        result.units.append(UnitScore(unit_text, base, label, mods, add1))

    # --- 첨가조건② 내포절
    emb = []
    for gi, g in enumerate(groups):
        if g.kind in ("관형절", "명사절", "부사절", "인용절"):
            args = [k for k in range(g.arg_start, g.head) if tags[k] not in ("SP", "SS", "SSO", "SSC", "SE", "SO", "SW")]
            score = cfg.embedded_clause_score
            if g.kind == "관형절" and not args:
                score = cfg.one_word_relative_clause_score
            # 내포절 안의 내포절: 앞 관형절·명사절이 꾸미는 말이 이 절의 성분인 경우
            # (쉼표나 접속 조사 '와/과'로 나열된 경우는 제외)
            prev = groups[gi - 1] if gi else None
            if (prev is not None and prev.kind in ("관형절", "명사절") and args and g.arg_start == prev.end + 1
                    and not any(tags[k] in ("SP", "JC") for k in range(g.arg_start, g.head))):
                score += cfg.nested_clause_bonus
            emb.append((g.kind, text[toks[g.head].start: toks[g.end].end], score))
        elif g.kind in ("final", "connective") and tags[g.pred] in ("VA",) and not has_obj[gi]:
            # 서술절: '코끼리는 코가 길다' – 주어 두 개, 서술어 바로 앞 어절이 '이/가'
            span = range(g.arg_start, g.head)
            subj = [k for k in span if tags[k] == "JKS" or (tags[k] == "JX" and toks[k].form in SUBJECT_JX)]
            pred_word = toks[g.pred].word_position
            if len(subj) >= 2 and tags[subj[-1]] == "JKS" and toks[subj[-1]].word_position == pred_word - 1:
                emb.append(("서술절", toks[subj[-1] - 1].form + toks[subj[-1]].form + " " + toks[g.pred].form,
                            cfg.embedded_clause_score))
    total = sum(s for _, _, s in emb)
    if len(emb) >= cfg.embedded_clause_cap_count:
        total = cfg.embedded_clause_cap_score
    result.embedded = emb
    result.add2 = min(total, cfg.embedded_clause_cap_score)
    return result


def _count_modifiers(toks, tags, s, e, comp_tokens) -> int:
    """관형어·부사어·독립어 개수."""
    count = 0
    for k in range(s, e + 1):
        tg = tags[k]
        if tg in ("MM", "IC", "JKV", "JKG", "MAJ"):
            count += 1
        elif tg == "MAG":
            if not (k + 1 <= e and tags[k + 1] in ("XSA", "XSV")):
                count += 1
        elif tg == "JKB" and k not in comp_tokens:
            count += 1
    # 명사 어절이 조사 없이 다음 명사 어절을 꾸미는 경우 (개구리 몸은, 대중 문화의)
    words = {}
    for k in range(s, e + 1):
        words.setdefault(toks[k].word_position, []).append(k)
    order = sorted(words)
    for a, b in zip(order, order[1:]):
        last, first = words[a][-1], words[b][0]
        if b == a + 1 and tags[last] in NOUNS and tags[first] in NOUNS and tags[last] != "SN":
            count += 1
    return count


def _surface_words(text, toks):
    if not toks:
        return []
    return text[toks[0].start: toks[-1].end].split()
