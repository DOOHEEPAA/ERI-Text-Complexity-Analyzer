"""표본에서 '서로 다른 단어'를 뽑아 어휘 등급을 매긴다 (X1, X2, Z).

형태소 분석기는 합성어·파생어를 잘게 나누는 경우가 많다
(예: 겨울옷 → 겨울+옷, 끈적끈적한 → 끈적끈적+하+ㄴ, 추워지면 → 춥+어+지+면).
특허 예시([0204], [0221])처럼 사전 표제어 단위로 세기 위해 한 어절 안에서
사전에 있는 가장 긴 표제어를 먼저 찾고, 없을 때만 형태소 단위로 나눈다.
"""
from __future__ import annotations

from dataclasses import dataclass

NOUNISH = {"NNG", "NNP", "NNB", "NR", "NP", "XR"}
AFFIX = {"XPN", "XSN"}
VERBAL = {"VV", "VA", "VX", "VCN"}
DERIV = {"XSV", "XSA"}
SINGLE = {"MAG", "MAJ", "MM", "IC"}
LINKING_EC = {"어", "아", "여", "어다", "아다", "고"}


def base_tag(tag: str) -> str:
    return tag.split("-")[0]


@dataclass
class Word:
    lemma: str
    tag: str          # 대표 kiwi 품사
    grade: int | None = None


def _span(text, toks):
    return text[toks[0].start: toks[-1].end]


def _eojeols(tokens):
    groups, cur, last = [], [], None
    for t in tokens:
        if cur and t.word_position != last:
            groups.append(cur); cur = []
        cur.append(t); last = t.word_position
    if cur:
        groups.append(cur)
    return groups


def extract_words(kiwi_tokens, text: str, vocab, count_proper_nouns=True) -> list[Word]:
    """문장 하나의 kiwi 토큰에서 단어(표제어) 목록을 뽑는다. 순서 유지, 중복 포함."""
    words = []
    for ej in _eojeols(kiwi_tokens):
        tags = [base_tag(t.tag) for t in ej]
        i, n = 0, len(ej)
        while i < n:
            tg = tags[i]
            # 1) 명사·어근 덩어리 (+접사) → 최장 일치, 뒤에 하다/되다 등이 붙으면 파생어
            if tg in NOUNISH or tg in AFFIX or (tg in ("MAG",) and i + 1 < n and tags[i + 1] in DERIV):
                j = i
                while j < n and (tags[j] in NOUNISH or tags[j] in AFFIX or (j == i and tg == "MAG")):
                    j += 1
                run = ej[i:j]
                if j < n and tags[j] in DERIV:
                    cand = _span(text, run) + ej[j].form + "다"
                    if cand in vocab:
                        words.append(Word(cand, "VA" if tags[j] == "XSA" else "VV"))
                        i = j + 1
                        continue
                words += _longest_match(run, [tags[k] for k in range(i, j)], text, vocab, count_proper_nouns)
                i = j
                if i < n and tags[i] in DERIV:
                    i += 1
                continue
            # 2) 용언: 사전에 있으면 '용언+어/아+용언' 합성어(추워지다, 갈아입다)로 묶는다
            if tg in VERBAL:
                lemma, end = ej[i].form + "다", i + 1
                k = i + 1
                while k + 1 < n and tags[k] == "EC" and ej[k].form in LINKING_EC and tags[k + 1] in VERBAL:
                    cand = _span(text, ej[i:k + 2]) + "다"
                    if cand in vocab:
                        lemma, end = cand, k + 2
                    k += 2
                words.append(Word(lemma, tg))
                i = end
                continue
            if tg in SINGLE:
                words.append(Word(ej[i].form, tg))
            i += 1
    return words


def _longest_match(run, tags, text, vocab, count_proper_nouns):
    """명사 덩어리를 사전 최장 일치로 나눈다. 접사 단독(들, 적 …)은 단어로 세지 않는다."""
    out, i, n = [], 0, len(run)
    whole = _span(text, run)
    if n > 1 and whole in vocab:
        return [Word(whole, "NNG")]
    while i < n:
        hit = None
        for j in range(n, i + 1, -1):
            cand = _span(text, run[i:j])
            if cand in vocab:
                hit = j; break
        if hit:
            out.append(Word(_span(text, run[i:hit]), "NNG")); i = hit
            continue
        t, tg = run[i], tags[i]
        if tg in AFFIX:
            pass
        elif tg == "NNP" and not count_proper_nouns:
            pass
        else:
            out.append(Word(t.form, tg))
        i += 1
    # 사전에 없는 합성명사(예: 수요량)는 조각이 아니라 하나의 단어로 센다.
    if n > 1 and len(out) > 1 and all(w.lemma not in vocab for w in out):
        core = run[:-1] if tags[-1] == "XSN" and run[-1].form == "들" else run
        if all(tags[k] in ("NNG", "NNP", "XPN", "XSN") for k in range(len(core))):
            return [Word(_span(text, core), "NNG")]
    return out


@dataclass
class VocabStats:
    words: dict            # lemma -> grade(None=사전에 없음)
    a_words: list
    not_a_words: list
    not_c_words: list

    @property
    def X1(self): return len(self.not_a_words)

    @property
    def X2(self): return len(self.a_words)

    @property
    def Z(self): return len(self.not_c_words)


def vocab_stats(words: list[Word], vocab, cfg) -> VocabStats:
    graded = {}
    for w in words:
        g = vocab.grade(w.lemma, w.tag, cfg.use_pos_for_grade)
        if w.lemma not in graded or (g is not None and (graded[w.lemma] is None or g < graded[w.lemma])):
            graded[w.lemma] = g
    a = sorted(w for w, g in graded.items() if g is not None and g <= cfg.a_max_grade)
    not_a = sorted(w for w, g in graded.items() if g is None or g > cfg.a_max_grade)
    not_c = sorted(w for w, g in graded.items() if g is None or g > cfg.c_max_grade)
    return VocabStats(graded, a, not_a, not_c)
