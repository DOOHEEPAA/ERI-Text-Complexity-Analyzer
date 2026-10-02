import pytest

from eri.analyzer import ERIAnalyzer
from eri.complexity import score_sentence
from eri.config import ERIConfig
from eri.lexical import extract_words
from eri.passages import Passage
from eri.sampling import Sentence, select_sample, split_sentences
from tests.data.patent_examples import FROG, FROG_TITLE, MARKET

cfg = ERIConfig()


def k_of(kiwi, text):
    return score_sentence(kiwi.tokenize(text), text, cfg)


def test_patent_sentence_complexity_example(kiwi):
    # 특허 [0223]: 주술1 + 부사어1 = 2, 주보술('증가하게 된다') 2 → 4
    sc = k_of(kiwi, "반대로 가격이 내리면 수요량은 증가하게 된다.")
    assert sc.score == 4
    assert [u.base_label for u in sc.units] == ["주술", "주보술"]


def test_basic_forms(kiwi):
    assert k_of(kiwi, "개구리들은 겨울에 잠을 자.").units[0].base_label == "주목술"
    assert k_of(kiwi, "그는 학생이 아니다.").units[0].base_label == "주보술"
    assert k_of(kiwi, "얼음이 물이 되었다.").units[0].base_label == "주보술"


def test_embedded_clauses(kiwi):
    assert [e[0] for e in k_of(kiwi, "코끼리는 코가 길다.").embedded] == ["서술절"]
    assert [e[0] for e in k_of(kiwi, "다행히 괜찮다고 해.").embedded] == ["인용절"]
    assert k_of(kiwi, "어떻게 지내고 있을까?").embedded == []       # 굳어진 부사 '어떻게'
    assert k_of(kiwi, "이러한 관계를 수요 법칙이라고 한다.").units[0].base_label == "주목술"


def test_embedded_cap(kiwi):
    text = ("내가 어제 읽은 책에 나온 사람이 만든 기계가 움직이는 원리를 설명한 글을 쓴 학생이 상을 받았다.")
    sc = k_of(kiwi, text)
    assert len(sc.embedded) >= 6
    assert sc.add2 == 18


def test_market_text_matches_patent(kiwi, mini_vocab):
    r = ERIAnalyzer(mini_vocab, cfg, kiwi).analyze(Passage("", MARKET, "중등"))
    assert r.K == pytest.approx(9.1, abs=0.3)      # 특허 [0223]: 평균 9.1
    assert set(r.vocab.not_c_words) >= {"수요자", "공급자", "수요량"}
    assert r.formula == "식2(중등)"


def test_frog_text_uses_elementary_formula_and_counts_title(kiwi, mini_vocab):
    r = ERIAnalyzer(mini_vocab, cfg, kiwi).analyze(Passage(FROG_TITLE, FROG, "초등"))
    assert r.Y == 17                                # 특허 [0206]: 제목 포함 17문장
    assert r.formula == "식1(초등)"
    assert set(r.vocab.not_a_words) >= {"겨울옷", "끈적끈적하다", "산토끼"}


def test_compound_and_derived_words(kiwi, mini_vocab):
    text = "땅이 꽁꽁 얼어붙으면 끈적끈적한 겨울옷 수요량은 증가하게 된다. 날씨가 추워지면 개구리들은 갈아입어요."
    words = set()
    for s in split_sentences(kiwi, text):
        words |= {w.lemma for w in extract_words(kiwi.tokenize(s.text), s.text, mini_vocab)}
    assert {"얼어붙다", "끈적끈적하다", "겨울옷", "수요량", "증가하다", "추워지다", "개구리", "갈아입다"} <= words
    assert not {"끈적끈적", "증가", "옷", "들"} & words


def test_sample_uses_whole_sentences_near_100_eojeol():
    sents = [Sentence(" ".join(["단어"] * n) + ".") for n in [12, 9, 15, 11, 8, 14, 10, 13, 9, 12, 16, 7, 10, 11, 9]]
    sample = select_sample(sents, 100, 0.15)
    total = sum(s.eojeol for s in sample)
    assert 85 <= total <= 115
    assert sample[0] is sents[0] and sample[-1] is sents[-1]
    assert all(s in sents for s in sample)


def test_short_text_uses_everything():
    sents = [Sentence("짧은 문장이다."), Sentence("또 있다.")]
    assert select_sample(sents) == sents


def test_pdf_line_wraps_are_restored(kiwi, mini_vocab):
    from pathlib import Path
    from eri.cleanup import clean_text
    raw = (Path(__file__).parent / "data" / "pdf_wrapped.txt").read_text(encoding="utf-8")
    text = clean_text(raw, kiwi)
    for fixed in ("어둡고 칙칙해", "얻은 양식을", "가난하고 고독했다", "미처 방 안을", "그 얼굴은", "분명히 보여",
                  "노동으로 인해", "가운데 있는 주황색", "생생하게 표현했다", "어떤 주제나"):
        assert fixed in text, fixed
    assert "㉠" not in text and "ⓐ" not in text
    assert "\n-1885년 4월\n" in text           # 짧은 줄(출처 표시)은 그대로 둔다
    # 줄바꿈을 복원하면 문장 수가 PDF 줄 수가 아니라 실제 문장 수가 된다
    r = ERIAnalyzer(mini_vocab, cfg, kiwi).analyze(Passage("감자 먹는 사람들", raw, "중등"))
    assert r.Y <= 13
    assert "둡" not in r.vocab.words and "감자 먹는 사람들" not in r.vocab.words


def test_normal_text_unchanged(kiwi):
    from eri.cleanup import clean_text
    text = "첫 문단은 한 줄로 쓴 문단이다. 두 번째 문장도 있다.\n둘째 문단이다."
    assert clean_text(text, kiwi) == text
