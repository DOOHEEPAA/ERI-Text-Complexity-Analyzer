"""지문 하나의 정량 지수 계산 (특허 도5·도6의 S110~S140 / S210~S240)."""
from __future__ import annotations

from dataclasses import dataclass, field

from . import formulas
from .cleanup import clean_text
from .complexity import score_sentence
from .config import ERIConfig
from .lexical import extract_words, vocab_stats
from .passages import Passage
from .sampling import Sentence, select_sample, split_sentences
from .textutil import round_half_up


@dataclass
class PassageResult:
    title: str
    level: str
    sample: list                 # Sentence 목록
    eojeol: int
    Y: int                       # 문장 수 (초등 공식)
    X1: int
    X2: int
    Z: int
    K: float                     # 문장 복잡도 평균 (중등 공식)
    quant: float                 # 정량 지수 (반올림 전)
    vocab: object                # VocabStats
    sentence_scores: list
    ratings: list = field(default_factory=list)
    qual_mean: float | None = None
    eri: float | None = None
    stage: str = ""
    note: str = ""

    @property
    def quant_rounded(self):
        return round_half_up(self.quant, 1)

    @property
    def formula(self):
        return "식1(초등)" if self.level == "초등" else "식2(중등)"

    @property
    def sample_text(self):
        return " ".join(s.text for s in self.sample)


class ERIAnalyzer:
    def __init__(self, vocab, cfg: ERIConfig | None = None, kiwi=None):
        self.cfg = cfg or ERIConfig()
        self.vocab = vocab
        if kiwi is None:
            from kiwipiepy import Kiwi
            kiwi = Kiwi()
        self.kiwi = kiwi

    def analyze(self, passage: Passage) -> PassageResult:
        cfg = self.cfg
        level = passage.level or cfg.default_level
        if level not in ("초등", "중등"):
            raise ValueError(f"학교급은 '초등' 또는 '중등'이어야 합니다: {level}")

        # S110/S210: 100어절 표본 (문장 단위, 도입·중간·끝, 제목 포함)
        text = clean_text(passage.text, self.kiwi, cfg.join_wrapped_lines, cfg.remove_question_markers)
        sents = split_sentences(self.kiwi, text)
        if not sents:
            raise ValueError("본문이 비어 있습니다.")
        use_title = cfg.include_title_in_sample and passage.title and not passage.title.startswith("지문 ")
        title_eojeol = len(passage.title.split()) if use_title else 0
        sample = select_sample(sents, cfg.sample_eojeol, cfg.sample_tolerance, title_eojeol)
        if use_title:
            sample = [Sentence(passage.title, is_title=True)] + sample
        for s in sample:
            s.tokens = self.kiwi.tokenize(s.text)
        eojeol = sum(s.eojeol for s in sample)

        # S120/S220: 어휘 (서로 다른 단어 수)
        words = []
        for s in sample:
            words += extract_words(s.tokens, s.text, self.vocab, cfg.count_proper_nouns, self.kiwi)
        stats = vocab_stats(words, self.vocab, cfg)

        # S130: 문장 수, S230: 문장 복잡도 평균
        Y = len(sample)
        scores = [score_sentence(s.tokens, s.text, cfg) for s in sample]
        K = sum(sc.score for sc in scores) / len(scores)

        if level == "초등":
            y = Y * 100 / eojeol if cfg.normalize_sentence_count and eojeol else Y
            quant = formulas.quantitative_elementary(y, stats.X1, cfg.elem_coef)
        else:
            quant = formulas.quantitative_middle(stats.X2, stats.Z, K, cfg.mid_coef)

        r = PassageResult(passage.title, level, sample, eojeol, Y, stats.X1, stats.X2, stats.Z,
                          K, quant, stats, scores)
        self.apply_qualitative(r, [])
        return r

    def apply_qualitative(self, r: PassageResult, ratings):
        """S170/S180: 정성 지수(평가자 평균)를 더해 최종 ERI와 학년·단계를 구한다."""
        from .qualitative import mean_score
        r.ratings = list(ratings)
        r.qual_mean = mean_score(r.ratings)
        r.eri = formulas.final_eri(r.quant, r.qual_mean)
        r.stage = formulas.stages_for(r.eri, self.cfg.stage_grade_min, self.cfg.stage_grade_max)
        given = sum(x is not None for x in r.ratings)
        if given == 0:
            r.note = "질적평가 미입력 (정량 지수만 반영)"
        elif given < self.cfg.rater_count:
            r.note = f"평가자 {given}명만 입력 (특허 권장 {self.cfg.rater_count}명)"
        else:
            r.note = ""
        return r
