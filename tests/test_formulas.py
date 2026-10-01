from eri import formulas
from eri.config import ERIConfig
from eri.textutil import round_half_up

cfg = ERIConfig()


def test_elementary_formula_matches_patent_example():
    # 특허 [0207]: Y=17, X1=8 → 3.2
    assert round_half_up(formulas.quantitative_elementary(17, 8, cfg.elem_coef)) == 3.2


def test_middle_formula_matches_patent_example():
    # 특허 [0224]: X2=36, Z=4, K=9.1 → 9.4
    assert round_half_up(formulas.quantitative_middle(36, 4, 9.1, cfg.mid_coef)) == 9.4


def test_final_eri_examples():
    assert formulas.final_eri(3.2, 0.0) == 3.2          # [0213]
    assert formulas.final_eri(9.4, 0.6) == 10.0         # [0230]
    assert formulas.final_eri(9.4, None) == 9.4


def test_round_half_up_not_bankers():
    assert round_half_up(2.25) == 2.3
    assert round_half_up(2.35) == 2.4


def test_stage_table_matches_figure_10():
    table = {(g, s): (lo, hi) for g, s, lo, hi in formulas.stage_table(3, 9)}
    assert table[(3, "기초")] == (2.5, 3.5)
    assert table[(3, "심화")] == (3.5, 4.5)
    assert table[(6, "심화")] == (6.5, 7.5)
    assert table[(7, "기본")] == (7.0, 8.0)
    assert (7, "기초") not in table
    assert table[(9, "심화")] == (9.5, 10.5)


def test_stage_lookup():
    # [0215] ERI 3.2 → 3학년 기초 또는 기본
    assert formulas.stages_for(3.2, 3, 9) == "3학년 기초 / 3학년 기본"
    assert formulas.stages_for(2.0, 3, 9) == "3학년 기초 미만"
    assert formulas.stages_for(11.0, 3, 9) == "9학년 심화 초과"
