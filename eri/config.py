"""ERI 계산 설정.

모든 값에는 기본값이 있으므로 설정 파일 없이도 동작한다.
값을 바꾸고 싶으면 프로그램 폴더에 `eri_config.json`을 두거나
`python run_eri.py --init-config`로 기본 설정 파일을 만든 뒤 수정한다.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path

CONFIG_FILE_NAME = "eri_config.json"


@dataclass
class ERIConfig:
    # --- 회귀식 계수 (특허 10-2309633 식1, 식2) ---
    # 초등: 7.438 - 0.391×Y + 0.032×X1 + 0.016×(X1×Y)
    elem_coef: dict = field(default_factory=lambda: {
        "intercept": 7.438, "Y": -0.391, "X1": 0.032, "X1Y": 0.016})
    # 중등: -0.060×X2 + 0.145×Z + 0.110×K + 0.024×(K×Z) + 9.075
    mid_coef: dict = field(default_factory=lambda: {
        "intercept": 9.075, "X2": -0.060, "Z": 0.145, "K": 0.110, "KZ": 0.024})

    # --- 어휘 등급 → 특허의 A/B/C 등급 대응 (누적, A ⊂ B ⊂ C) ---
    # 국립국어원 어휘 등급(1~5등급)을 사용할 때 "몇 등급 이하를 A로 볼 것인가".
    # 특허 A등급 ≈ 3천 어, C등급 ≈ 1만 6천 어. 특허 예시 지문(도11·도12)으로
    # 검증한 결과 A=1등급, C=1~4등급일 때 예시값에 가장 가깝다.
    a_max_grade: int = 1
    b_max_grade: int = 2
    c_max_grade: int = 4
    # 고유명사를 어휘 수(X1, Z)에 포함할지
    count_proper_nouns: bool = True
    # 사전의 품사 정보를 이용해 동형어(예: 가정/가정) 등급을 고를지
    use_pos_for_grade: bool = True

    # --- 표본 추출 ---
    sample_eojeol: int = 100          # 목표 어절 수
    sample_tolerance: float = 0.15    # 전체가 목표의 (1+tolerance)배 이하이면 전체 사용
    include_title_in_sample: bool = True   # 특허 [0201]: 표제도 표본에 포함
    # PDF·HWP에서 복사할 때 생긴 줄바꿈(단어 중간 끊김 포함)을 복원할지
    join_wrapped_lines: bool = True
    # ㉠ ⓐ ① 같은 문제용 기호를 지울지
    remove_question_markers: bool = True
    # 분석 전 맞춤법·띄어쓰기 자동 교정 (가난 하고 → 가난하고, 됬다 → 됐다). 교정 내역은 결과에 표시
    auto_correct: bool = True
    # 초등 공식의 Y를 '100어절당 문장 수'로 환산할지 (특허 예시는 환산하지 않음)
    normalize_sentence_count: bool = False

    # --- 문장 복잡도(K) ---
    embedded_clause_score: int = 3        # 내포절 1개당 점수
    one_word_relative_clause_score: int = 2   # 한 단어 관형절(예: 빨간 사과) 점수 [0164]
    nested_clause_bonus: int = 1          # 내포절 안에 내포절이 있을 때 추가 [0165]
    embedded_clause_cap_count: int = 6    # 이 개수 이상이면
    embedded_clause_cap_score: int = 18   # 이 점수로 고정 [0166]

    # --- 질적 평가 ---
    qual_min: float = -3.0
    qual_max: float = 3.0
    rater_count: int = 3
    agreement_threshold: float = 0.9      # 평가자 간 일치도 기준 [0110]

    # --- 기타 ---
    default_level: str = "중등"           # 지문에 학교급 표시가 없을 때 적용
    # 도10 학년·단계 표의 범위. 도10은 3~9학년이지만 특허는 중등을 7~12학년으로 보므로([0192])
    # 같은 규칙(기본 [g, g+1), 심화 [g+0.5, g+1.5))으로 12학년까지 늘려 둔다.
    stage_grade_min: int = 3
    stage_grade_max: int = 12

    # ------------------------------------------------------------------
    @classmethod
    def load(cls, path: str | Path | None = None, search_dirs=()) -> "ERIConfig":
        """설정 파일이 있으면 읽어 기본값 위에 덮어쓴다."""
        candidates = [Path(path)] if path else [Path(d) / CONFIG_FILE_NAME for d in search_dirs]
        for p in candidates:
            if p.is_file():
                with open(p, encoding="utf-8-sig") as f:
                    data = json.load(f)
                cfg = cls()
                known = {f.name for f in fields(cls)}
                for k, v in data.items():
                    if k.startswith("_"):
                        continue
                    if k not in known:
                        print(f"[설정] 알 수 없는 항목 '{k}'은(는) 무시합니다.")
                        continue
                    if isinstance(getattr(cfg, k), dict):
                        getattr(cfg, k).update(v)
                    else:
                        setattr(cfg, k, v)
                cfg.source = str(p)
                return cfg
            if path:
                raise FileNotFoundError(f"설정 파일을 찾을 수 없습니다: {p}")
        cfg = cls()
        cfg.source = None
        return cfg

    def save(self, path: str | Path) -> None:
        data = {"_설명": "ERI 계산기 설정 파일. 필요한 항목만 남겨도 됩니다."}
        data.update(asdict(self))
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def as_rows(self):
        return [(k, json.dumps(v, ensure_ascii=False) if isinstance(v, dict) else v)
                for k, v in asdict(self).items()]
