"""결과 엑셀 저장과 질적 점수 재계산."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import formulas
from .qualitative import mean_score, pairwise_agreement, parse_score
from .textutil import normalize_header

RESULT_SHEET = "ERI 결과"
HEAD_FILL = PatternFill("solid", fgColor="DCE6F8")
ERI_FONT = Font(bold=True, color="C00000")


def rater_header(i):
    return f"평가자{i}"


def result_columns(n_raters):
    return (["지문명", "학교급", "적용 공식", "표본 어절 수", "문장 수(Y)",
             "X1(A등급 외 단어 수)", "X2(A등급 단어 수)", "Z(C등급 외 단어 수)", "K(문장 복잡도 평균)",
             "정량 지수"]
            + [rater_header(i + 1) for i in range(n_raters)]
            + ["정성 지수(평균)", "ERI", "학년·단계", "비고", "표본 텍스트"])


def _grade_class(g, cfg):
    if g is None:
        return "C등급 외 (사전에 없음)"
    if g <= cfg.a_max_grade:
        return "A"
    if g <= cfg.b_max_grade:
        return "B"
    if g <= cfg.c_max_grade:
        return "C"
    return "C등급 외"


def _style_header(ws, widths):
    for c, w in enumerate(widths, 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "B2"


def write_report(results, cfg, path, vocab_source="", agreement=None):
    wb = Workbook()
    ws = wb.active
    ws.title = RESULT_SHEET
    cols = result_columns(cfg.rater_count)
    ws.append(cols)
    for r in results:
        ratings = list(r.ratings) + [None] * (cfg.rater_count - len(r.ratings))
        ws.append([r.title, r.level, r.formula, r.eojeol, r.Y, r.X1, r.X2, r.Z, round(r.K, 2),
                   r.quant_rounded] + ratings[:cfg.rater_count]
                  + [None if r.qual_mean is None else round(r.qual_mean, 2), r.eri, r.stage, r.note,
                     r.sample_text])
    widths = [26, 7, 10, 8, 8, 10, 10, 10, 11, 9] + [8] * cfg.rater_count + [10, 8, 22, 26, 70]
    _style_header(ws, widths)
    eri_col = cols.index("ERI") + 1
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            wrap = cell.column == len(cols)
            cell.alignment = Alignment(horizontal="left" if cell.column in (1, len(cols), len(cols) - 1) else "center",
                                       vertical="center", wrap_text=wrap)
        row[eri_col - 1].font = ERI_FONT

    # 어휘 상세
    wv = wb.create_sheet("어휘 상세")
    wv.append(["지문명", "단어", "사전 등급", "구분"])
    for r in results:
        for w, g in sorted(r.vocab.words.items(), key=lambda x: (x[1] is None, x[1] or 0, x[0])):
            wv.append([r.title, w, g, _grade_class(g, cfg)])
    _style_header(wv, [26, 16, 10, 22])

    # 문장 복잡도 상세
    wk = wb.create_sheet("문장 복잡도 상세")
    wk.append(["지문명", "번호", "문장", "절별 점수 (기본 형식 + 첨가조건①)", "내포절 (첨가조건②)", "문장 점수"])
    for r in results:
        for i, (s, sc) in enumerate(zip(r.sample, r.sentence_scores), 1):
            units, emb = sc.describe()
            wk.append([r.title, ("제목" if s.is_title else i), s.text, units, emb, sc.score])
    _style_header(wk, [26, 6, 60, 70, 40, 9])
    for row in wk.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=cell.column in (3, 4, 5))

    # 평가자 일치도
    if agreement is None:
        agreement = pairwise_agreement([[(r.ratings + [None] * cfg.rater_count)[i] for r in results]
                                        for i in range(cfg.rater_count)])
    _write_agreement(wb, agreement, cfg)

    # 설정
    wc = wb.create_sheet("설정")
    wc.append(["항목", "값"])
    wc.append(["생성 일시", datetime.now().strftime("%Y-%m-%d %H:%M")])
    wc.append(["어휘 등급 파일", vocab_source])
    wc.append(["설정 파일", getattr(cfg, "source", None) or "(기본값)"])
    for k, v in cfg.as_rows():
        wc.append([k, v])
    _style_header(wc, [32, 80])

    wb.save(path)
    return Path(path)


def _write_agreement(wb, agreement, cfg):
    if "평가자 일치도" in wb.sheetnames:
        del wb["평가자 일치도"]
    wa = wb.create_sheet("평가자 일치도", 1 if len(wb.sheetnames) > 1 else None)
    mean, pairs = agreement
    wa.append(["평가자 쌍", "일치도"])
    for (i, j), v in pairs.items():
        wa.append([f"평가자{i}-평가자{j}", round(v, 3)])
    wa.append([])
    if mean is None:
        wa.append(["평균", "점수가 2명 이상 입력되지 않아 계산하지 않음"])
    else:
        ok = mean >= cfg.agreement_threshold
        wa.append(["평균", round(mean, 3)])
        wa.append(["판정", f"{'충족' if ok else '미달'} (기준 {cfg.agreement_threshold} 이상, 특허 [0110])"])
    wa.append(["참고", "특허는 3·6·9학년 텍스트로 사전 평가를 반복해 평균 일치도 0.9 이상을 확보한 뒤 본 평가를 하도록 한다."])
    _style_header(wa, [22, 60])


def recalc_report(path, cfg, out_path=None):
    """결과 엑셀의 평가자 점수 열을 직접 채운 뒤 다시 계산한다 (재분석 없이)."""
    wb = load_workbook(path)
    if RESULT_SHEET not in wb.sheetnames:
        raise ValueError(f"'{RESULT_SHEET}' 시트가 없습니다. 이 프로그램이 만든 결과 파일인지 확인하세요.")
    ws = wb[RESULT_SHEET]
    header = [normalize_header(c.value) for c in ws[1]]
    def col(name):
        name = normalize_header(name)
        return header.index(name) + 1 if name in header else None
    rater_cols = [i + 1 for i, h in enumerate(header) if h.startswith("평가자")]
    qc, qm, ec, sc, nc, tc = (col("정량 지수"), col("정성 지수(평균)"), col("ERI"), col("학년·단계"),
                              col("비고"), col("지문명"))
    all_ratings = [[] for _ in rater_cols]
    count = 0
    for row in range(2, ws.max_row + 1):
        quant = ws.cell(row, qc).value
        if quant is None:
            continue
        ratings = []
        for k, c in enumerate(rater_cols):
            try:
                v = parse_score(ws.cell(row, c).value, cfg.qual_min, cfg.qual_max)
            except ValueError as e:
                raise ValueError(f"{ws.cell(row, tc).value} / {header[c - 1]}: {e}") from None
            ratings.append(v)
            all_ratings[k].append(v)
        m = mean_score(ratings)
        eri = formulas.final_eri(float(quant), m)
        ws.cell(row, qm).value = None if m is None else round(m, 2)
        ws.cell(row, ec).value = eri
        ws.cell(row, sc).value = formulas.stages_for(eri, cfg.stage_grade_min, cfg.stage_grade_max)
        given = sum(v is not None for v in ratings)
        ws.cell(row, nc).value = ("질적평가 미입력 (정량 지수만 반영)" if given == 0 else
                                  f"평가자 {given}명만 입력 (특허 권장 {cfg.rater_count}명)" if given < cfg.rater_count else "")
        count += 1
    _write_agreement(wb, pairwise_agreement(all_ratings), cfg)
    out = out_path or path
    wb.save(out)
    return Path(out), count
