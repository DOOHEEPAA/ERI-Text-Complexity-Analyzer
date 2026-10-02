"""명령줄 실행.

    python run_eri.py                         # 창(GUI)으로 실행 (파일을 자동으로 찾아 채워 둠)
    python run_eri.py 지문.txt                 # 바로 분석 (어휘 파일은 자동 탐색)
    python run_eri.py 지문폴더 --level 초등 --vocab 어휘.xlsx --out 결과.xlsx
    python run_eri.py --recalc 결과.xlsx       # 엑셀에 평가자 점수를 적은 뒤 ERI 다시 계산
    python run_eri.py --pilot 사전평가.xlsx    # 평가자 간 일치도(사전 평가) 확인
    python run_eri.py --calibrate 학년지문.xlsx --save-config   # 계수 재보정
    python run_eri.py --init-config            # 설정 파일(eri_config.json) 만들기
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from .config import CONFIG_FILE_NAME, ERIConfig

APP_DIR = Path(sys.argv[0]).resolve().parent if sys.argv and sys.argv[0] else Path.cwd()


def search_dirs():
    dirs = []
    for base in (Path.cwd(), APP_DIR, Path(__file__).resolve().parent.parent):
        for d in (base, base / "data"):
            if d.is_dir() and d not in dirs:
                dirs.append(d)
    return dirs


def default_output_path(first_input=None) -> Path:
    if first_input:
        p = Path(first_input).resolve()
        base = p if p.is_dir() else p.parent
    else:
        base = Path.cwd()
    return base / f"ERI_분석결과_{datetime.now():%Y%m%d_%H%M}.xlsx"


def resolve_vocab(path, log=print):
    from .vocab import find_vocab_file
    if path:
        return Path(path)
    found = find_vocab_file(search_dirs())
    if not found:
        raise FileNotFoundError(
            "어휘 등급 엑셀 파일을 찾지 못했습니다. 프로그램 폴더(또는 data 폴더)에 넣거나 --vocab 으로 지정하세요.")
    log(f"어휘 등급 파일 자동 선택: {found.name}")
    return found


def run_analysis(inputs, vocab_path, cfg, log=print, progress=None):
    """지문 파일 목록 → (분석 결과 목록, 어휘 사전). GUI와 CLI가 함께 쓴다."""
    from .analyzer import ERIAnalyzer
    from .passages import load_passages
    from .vocab import load_vocabulary

    log("어휘 등급 사전 불러오는 중…")
    vocab = load_vocabulary(vocab_path)
    log(f"  {len(vocab):,}개 표제어 ({Path(vocab_path).name})")
    passages = load_passages(inputs)
    if not passages:
        raise ValueError("분석할 지문이 없습니다. 지문 파일 형식을 확인하세요.")
    log(f"지문 {len(passages)}개를 찾았습니다. 형태소 분석기 준비 중…")
    analyzer = ERIAnalyzer(vocab, cfg)
    results, failed = [], []
    for i, p in enumerate(passages, 1):
        try:
            r = analyzer.analyze(p)
            results.append(r)
            fixed = f", 교정 {len(r.corrections)}건" if r.corrections else ""
            log(f"  [{i}/{len(passages)}] {r.title} ({r.level}, {r.formula}) 정량 지수 {r.quant_rounded}{fixed}")
        except Exception as e:  # 한 지문의 오류로 전체가 멈추지 않도록
            failed.append((p.title, str(e)))
            log(f"  [{i}/{len(passages)}] {p.title}: 오류 – {e}")
        if progress:
            progress(i, len(passages))
    return results, vocab, analyzer, passages, failed


def build_parser():
    ap = argparse.ArgumentParser(prog="run_eri.py", description="ERI(EBS Reading Index) 텍스트 복잡도 계산기",
                                 formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    ap.add_argument("inputs", nargs="*", help="지문 파일 또는 폴더 (.txt .md .xlsx .csv .docx)")
    ap.add_argument("--vocab", help="어휘 등급 엑셀/CSV (생략하면 자동 탐색)")
    ap.add_argument("--level", choices=["초등", "중등"], help="학교급 표시가 없는 지문에 적용할 학교급")
    ap.add_argument("--out", help="결과 엑셀 경로")
    ap.add_argument("--config", help=f"설정 파일 (기본: 프로그램 폴더의 {CONFIG_FILE_NAME})")
    ap.add_argument("--gui", action="store_true", help="창으로 실행")
    ap.add_argument("--no-gui", action="store_true", help="창을 띄우지 않음")
    ap.add_argument("--no-correct", action="store_true", help="분석 전 맞춤법·띄어쓰기 자동 교정을 하지 않음")
    ap.add_argument("--check-only", action="store_true",
                    help="ERI는 계산하지 않고 맞춤법·띄어쓰기 교정 결과만 엑셀로 저장")
    ap.add_argument("--recalc", metavar="결과.xlsx", help="결과 엑셀의 평가자 점수로 ERI 다시 계산")
    ap.add_argument("--pilot", metavar="사전평가.xlsx", help="평가자 간 일치도 계산 ('평가자1', '평가자2'… 열)")
    ap.add_argument("--calibrate", metavar="파일", help="학년이 알려진 지문으로 회귀 계수 재추정")
    ap.add_argument("--save-config", action="store_true", help="--calibrate 결과를 설정 파일에 저장")
    ap.add_argument("--init-config", action="store_true", help="기본 설정 파일 만들기")
    return ap


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        return _main(args)
    except (FileNotFoundError, ValueError, PermissionError, ImportError) as e:
        print(f"오류: {e}")
        return 1


def _main(args):
    creating = args.init_config or args.save_config
    if args.config and creating and not Path(args.config).is_file():
        cfg = ERIConfig.load(None, ())          # 새로 만들 설정 파일: 기본값에서 시작
    else:
        cfg = ERIConfig.load(args.config, search_dirs())
    if args.level:
        cfg.default_level = args.level
    if args.no_correct:
        cfg.auto_correct = False

    if args.init_config:
        path = Path(args.config or APP_DIR / CONFIG_FILE_NAME)
        cfg.save(path)
        print(f"설정 파일을 만들었습니다: {path}")
        return 0

    if args.recalc:
        from .report import recalc_report
        out, n = recalc_report(args.recalc, cfg, args.out)
        print(f"{n}개 지문의 ERI를 다시 계산했습니다 → {out}")
        return 0

    if args.pilot:
        return _pilot(args.pilot, cfg)

    if args.calibrate:
        return _calibrate(args, cfg)

    want_gui = args.gui or (not args.inputs and not args.no_gui)
    if want_gui:
        try:
            from .gui import launch
        except ImportError as e:
            print(f"창(GUI)을 열 수 없습니다 ({e}). 명령줄 모드로 실행합니다.")
        else:
            return launch(cfg, args)

    inputs = args.inputs
    if not inputs:
        from .passages import find_passage_files
        inputs = find_passage_files(search_dirs())
        if not inputs:
            print("지문 파일을 찾지 못했습니다. `python run_eri.py 지문.txt` 처럼 지정하세요.")
            return 1
        print("지문 파일 자동 선택: " + ", ".join(p.name for p in inputs))
    missing = [str(p) for p in inputs if not Path(p).exists()]
    if missing:
        raise FileNotFoundError("지문 파일을 찾을 수 없습니다: " + ", ".join(missing))
    if args.check_only:
        return _check_only(inputs, cfg, args.out)
    vocab_path = resolve_vocab(args.vocab)
    results, vocab, _, _, failed = run_analysis(inputs, vocab_path, cfg)
    if not results:
        return 1
    from .report import write_report
    out = Path(args.out) if args.out else default_output_path(inputs[0])
    write_report(results, cfg, out, str(vocab_path))
    print(f"\n결과 저장: {out}")
    print("질적 평가: 결과 엑셀의 '평가자1~3' 열에 -3~+3 점수를 적은 뒤 "
          f"`python run_eri.py --recalc \"{out.name}\"` 를 실행하면 ERI가 다시 계산됩니다.")
    if failed:
        print(f"분석하지 못한 지문 {len(failed)}개: " + ", ".join(t for t, _ in failed))
    return 0


def _read_rater_table(path):
    from openpyxl import load_workbook
    from .qualitative import parse_score
    from .textutil import normalize_header
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    if "ERI 결과" in wb.sheetnames:
        ws = wb["ERI 결과"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header = [normalize_header(h) for h in rows[0]]
    cols = [i for i, h in enumerate(header) if h.startswith("평가자")]
    if len(cols) < 2:
        raise ValueError("'평가자1', '평가자2' … 열이 2개 이상 필요합니다.")
    data = [[parse_score(r[c]) if c < len(r) else None for r in rows[1:]] for c in cols]
    return [header[c] for c in cols], data


def _pilot(path, cfg):
    from .qualitative import pairwise_agreement
    names, data = _read_rater_table(path)
    mean, pairs = pairwise_agreement(data)
    print("평가자 간 일치도 (같은 점수를 준 비율)")
    for (i, j), v in pairs.items():
        print(f"  {names[i - 1]} - {names[j - 1]}: {v:.3f}")
    if mean is None:
        print("비교할 점수가 없습니다.")
        return 1
    ok = mean >= cfg.agreement_threshold
    print(f"평균 일치도 {mean:.3f} → {'기준 충족, 본 평가를 진행하세요.' if ok else '기준 미달, 평가 기준을 맞춘 뒤 사전 평가를 반복하세요.'}"
          f" (기준 {cfg.agreement_threshold})")
    return 0 if ok else 2


def _calibrate(args, cfg):
    from .calibrate import fit
    vocab_path = resolve_vocab(args.vocab)
    results, _, _, passages, _ = run_analysis([args.calibrate], vocab_path, cfg)
    by_title = {p.title: p.grade for p in passages}
    grades = [by_title.get(r.title) for r in results]
    if not any(g is not None for g in grades):
        print("학년 정보가 없습니다. 엑셀이면 '학년' 열을, 텍스트면 제목에 [5학년] 같은 표시를 넣으세요.")
        return 1
    fitted = fit(results, grades)
    for level, (coef, r2, n) in fitted.items():
        if coef is None:
            print(f"{level}: 학년이 표시된 지문이 {n}개뿐이라 추정하지 않았습니다 (최소 {6 if level == '초등' else 7}개).")
            continue
        print(f"{level} (n={n}, R²={r2:.3f}): {coef}")
        if n < 30:
            print(f"  ※ 지문이 {n}개뿐이라 계수가 불안정할 수 있습니다. 학년별로 고르게 30개 이상을 권장합니다.")
        if level == "초등":
            cfg.elem_coef.update(coef)
        else:
            cfg.mid_coef.update(coef)
    if args.save_config:
        path = Path(args.config or APP_DIR / CONFIG_FILE_NAME)
        cfg.save(path)
        print(f"설정 파일에 저장했습니다: {path}")
    return 0


def run_check(inputs, cfg, log=print):
    """맞춤법·띄어쓰기 교정만 실행 → [(지문명, 교정 목록, 원문, 교정 후)]."""
    from kiwipiepy import Kiwi

    from .analyzer import ERIAnalyzer
    from .passages import load_passages
    passages = load_passages(inputs)
    if not passages:
        raise ValueError("검사할 지문이 없습니다.")
    cfg.auto_correct = True
    analyzer = ERIAnalyzer(None, cfg, Kiwi())
    items = []
    for p in passages:
        log_items = []
        fixed = analyzer.preprocess(p.text, log_items)
        items.append((p.title, log_items, p.text, fixed))
        log(f"  {p.title}: 교정 {len(log_items)}건")
        for c in log_items:
            log(f"      [{c.kind}] {c.before}  →  {c.after}")
    return items


def _check_only(inputs, cfg, out):
    from openpyxl import Workbook

    from .report import write_corrections
    items = run_check(inputs, cfg)
    out = Path(out) if out else default_output_path(inputs[0]).with_name(
        f"ERI_교정결과_{datetime.now():%Y%m%d_%H%M}.xlsx")
    wb = Workbook()
    wb.remove(wb.active)
    write_corrections(wb, items)
    wb.save(out)
    print(f"\n교정 결과 저장: {out}")
    return 0
