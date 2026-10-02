"""지문 불러오기.

지원 형식
- 텍스트(.txt, .md): 인코딩 자동 판별(UTF-8 / 메모장 ANSI(CP949) 등). 다음 형식을 자동 인식한다.
    1) `제목: 본문`        (기존 지문모음.txt 형식, 본문이 여러 줄이어도 됨)
    2) `# 제목` 또는 `[제목]` 머리줄 다음에 본문
    3) `---` / `===` 또는 빈 줄 2개 이상으로 구분된 덩어리 (첫 줄이 짧고 마침표가 없으면 제목)
    4) 위 표시가 하나도 없으면 파일 전체를 지문 하나로 보고 파일 이름을 제목으로 쓴다.
- 엑셀/CSV(.xlsx, .csv): '제목/지문명' + '본문/지문/내용' 열 (선택: '학교급' 열)
- 워드(.docx): 견본(견본/지문_입력_견본.docx)의 '제목·학교급·내용' 표 형식, 또는 제목 스타일 문단
- 폴더: 폴더 안의 지원 파일을 모두 읽는다.

학교급(초등/중등)은 제목 앞뒤의 `[초등]`, `(중등)`, `[5학년]` 같은 표시나 '학교급' 열로 지정한다.
표시가 없으면 설정의 기본 학교급을 쓴다.
"""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path

from .textutil import normalize_header, read_text

PASSAGE_EXTS = (".txt", ".md", ".xlsx", ".csv", ".docx")
TITLE_HEADERS = ("지문명", "제목", "title", "이름")
BODY_HEADERS = ("본문", "지문", "내용", "텍스트", "text", "body")
LEVEL_HEADERS = ("학교급", "수준", "level")
GRADE_HEADERS = ("학년", "목표학년", "grade")

LEVEL_TAG = re.compile(r"[\[\(<【]\s*(초등학교|중학교|고등학교|초등|중등|고등|\d{1,2}\s*학년)\s*[\]\)>】]")
SEPARATOR = re.compile(r"^\s*(-{3,}|={3,}|\*{3,}|#{3,}\s*$)\s*$")
HEADING = re.compile(r"^\s*(#{1,2})\s+(.+?)\s*$|^\s*\[([^\[\]]{1,60})\]\s*$")
# '제목: 본문' – 제목은 60자 이하, 문장부호로 끝나지 않고 콜론 앞에 공백/숫자만 있는 경우는 제외
TITLE_COLON = re.compile(r"^\s*([^:：\n]{1,60}?)\s*[:：]\s*(.*)$")


@dataclass
class Passage:
    title: str
    text: str
    level: str | None = None     # '초등' / '중등' / None(기본값 사용)
    source: str = ""
    grade: float | None = None   # 알려진 학년 (계수 재보정용, 예: [5학년])


def parse_level(value) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    m = re.search(r"(\d{1,2})\s*학년", s) or re.fullmatch(r"(\d{1,2})", s)
    if m:
        return "초등" if int(m.group(1)) <= 6 else "중등"
    if s.startswith("초"):
        return "초등"
    if s.startswith(("중", "고")):
        return "중등"
    return None


def parse_grade_number(value):
    if value is None:
        return None
    m = re.search(r"\d+(\.\d+)?", str(value))
    return float(m.group()) if m else None


def split_level_tag(title: str):
    """제목에서 [초등]/[중등]/[5학년] 표시를 떼어 (제목, 학교급, 학년)을 돌려준다."""
    level = grade = None
    while True:
        m = LEVEL_TAG.search(title)
        if not m:
            break
        tag = m.group(1)
        level = parse_level(tag) or level
        if "학년" in tag:
            grade = parse_grade_number(tag)
        title = title[:m.start()] + title[m.end():]
    return title.strip(" -:\t"), level, grade


NOT_TITLES = {"예", "예시", "보기", "참고", "주", "출처", "단", "답", "정답", "해설", "질문", "문제"}


def _looks_like_title(s: str) -> bool:
    s = s.strip()
    return (0 < len(s) <= 60 and s not in NOT_TITLES
            and not re.search(r"[.?!。…\"'”’]$", s))


def _make(title, lines, source, idx):
    body = "\n".join(l.strip() for l in lines if l.strip())
    title, level, grade = split_level_tag(title or f"지문 {idx}")
    return Passage(title or f"지문 {idx}", body, level, source, grade)


def parse_text(content: str, source: str = "", default_title: str = "지문") -> list[Passage]:
    lines = content.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    # 1) '# 제목' / '[제목]' 머리줄
    if any(HEADING.match(l) for l in lines):
        out, title, buf = [], None, []
        for l in lines:
            m = HEADING.match(l)
            if m:
                if title is not None or any(x.strip() for x in buf):
                    out.append(_make(title, buf, source, len(out) + 1))
                title, buf = (m.group(2) or m.group(3)), []
            elif not SEPARATOR.match(l):
                buf.append(l)
        out.append(_make(title, buf, source, len(out) + 1))
        return [p for p in out if p.text]

    # 2) '제목: 본문' (줄 맨 앞, 이전 줄이 비어 있거나 파일 첫 줄이거나 앞 지문이 끝난 경우)
    colon_lines = [i for i, l in enumerate(lines) if l.strip() and TITLE_COLON.match(l)
                   and _looks_like_title(TITLE_COLON.match(l).group(1))]
    if colon_lines and colon_lines[0] == next(i for i, l in enumerate(lines) if l.strip()):
        out, title, buf = [], None, []
        for i, l in enumerate(lines):
            m = TITLE_COLON.match(l) if i in colon_lines else None
            # 본문 중간의 '예: ...' 같은 줄을 제목으로 착각하지 않도록 제목 후보는
            # 지문 첫 줄이거나 바로 앞 줄이 빈 줄/구분선이어야 한다.
            prev_blank = i == 0 or not lines[i - 1].strip() or SEPARATOR.match(lines[i - 1])
            if m and (title is None or prev_blank or _starts_new_passage(lines, i)):
                if title is not None:
                    out.append(_make(title, buf, source, len(out) + 1))
                title, buf = m.group(1), [m.group(2)]
            elif not SEPARATOR.match(l):
                buf.append(l)
        out.append(_make(title, buf, source, len(out) + 1))
        return [p for p in out if p.text]

    # 3) 구분선 또는 빈 줄 2개 이상으로 나뉜 덩어리
    blocks, buf, blank = [], [], 0
    for l in lines:
        if SEPARATOR.match(l):
            blocks.append(buf); buf, blank = [], 0
            continue
        if not l.strip():
            blank += 1
            if blank >= 2 and buf:
                blocks.append(buf); buf = []
            continue
        blank = 0
        buf.append(l)
    blocks.append(buf)
    blocks = [b for b in blocks if any(x.strip() for x in b)]
    out = []
    for b in blocks:
        b = [x for x in b if x.strip()]
        if len(b) > 1 and _looks_like_title(b[0]):
            out.append(_make(b[0], b[1:], source, len(out) + 1))
        else:
            name = default_title if len(blocks) == 1 else f"{default_title} {len(out) + 1}"
            out.append(_make(name, b, source, len(out) + 1))
    return [p for p in out if p.text]


def _starts_new_passage(lines, i):
    """기존 지문모음.txt처럼 지문 사이에 빈 줄이 없을 때: 앞 줄이 문장 끝으로 끝나면 새 지문으로 본다."""
    prev = lines[i - 1].strip() if i else ""
    return bool(re.search(r"[.?!。…\"”’)]$", prev))


def _read_table(path: Path) -> list[Passage]:
    if path.suffix.lower() == ".csv":
        rows = list(csv.reader(read_text(path).splitlines()))
        tables = [(path.stem, rows)]
    else:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        tables = [(ws.title, list(ws.iter_rows(values_only=True))) for ws in wb.worksheets]
        wb.close()
    out = []
    for _, rows in tables:
        for hi, header in enumerate(rows[:10]):
            norm = [normalize_header(h).lower() for h in header]
            def col(names):
                for n in names:
                    for i, h in enumerate(norm):
                        if h == n.lower() or (n.lower() in h and len(h) <= 8):
                            return i
                return None
            bi, ti, li, gi = col(BODY_HEADERS), col(TITLE_HEADERS), col(LEVEL_HEADERS), col(GRADE_HEADERS)
            if bi is None or bi == ti:
                continue
            for r in rows[hi + 1:]:
                if not r or bi >= len(r) or not r[bi] or not str(r[bi]).strip():
                    continue
                title = str(r[ti]).strip() if ti is not None and ti < len(r) and r[ti] else f"지문 {len(out) + 1}"
                title, tag_level, tag_grade = split_level_tag(title)
                level = parse_level(r[li]) if li is not None and li < len(r) else None
                grade = parse_grade_number(r[gi]) if gi is not None and gi < len(r) else None
                grade = grade or tag_grade
                if not (level or tag_level) and grade:
                    level = "초등" if grade <= 6 else "중등"
                out.append(Passage(title, str(r[bi]).strip(), level or tag_level, str(path), grade))
            break
    return out


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DOCX_LABELS = (("title", ("제목", "지문명", "title")), ("level", ("학교급", "level")),
               ("grade", ("학년", "grade")), ("body", ("내용", "본문", "지문", "text", "body")))


def _docx_label(cell_text: str):
    """표 왼쪽 칸 이름('제목', '학교급(초등/중등)', '내용:' 등) → 필드 이름."""
    first = next((l for l in cell_text.split("\n") if l.strip()), "")   # 첫 줄만 (둘째 줄은 도움말)
    label = re.sub(r"[\(\[].*?[\)\]]", "", normalize_header(first)).strip(":：*").lower()
    for field_name, names in DOCX_LABELS:
        if label in names:
            return field_name
    return None


def _docx_paragraph_text(p) -> str:
    parts = []
    for node in p.iter():
        if node.tag == W_NS + "t" and node.text:
            parts.append(node.text)
        elif node.tag in (W_NS + "tab",):
            parts.append(" ")
        elif node.tag in (W_NS + "br", W_NS + "cr"):
            parts.append("\n")
    return "".join(parts)


def _read_docx(path: Path) -> list[Passage]:
    """워드 파일 읽기 (python-docx 없이 표준 라이브러리만 사용).

    1) 견본 형식: '제목 | …', '학교급 | …', '내용 | …' 행으로 된 표 → 표(또는 '제목' 행) 하나가 지문 하나.
       표 밖의 안내문은 무시한다.
    2) 표가 없으면 문단을 읽는다. 제목 스타일 문단이나 '제목: …' 문단이 새 지문의 시작이다.
    """
    import zipfile
    from xml.etree import ElementTree as ET

    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
        style_names = {}
        if "word/styles.xml" in z.namelist():
            for st in ET.fromstring(z.read("word/styles.xml")).iter(W_NS + "style"):
                name = st.find(W_NS + "name")
                style_names[st.get(W_NS + "styleId")] = (name.get(W_NS + "val") if name is not None else "").lower()
    body = root.find(W_NS + "body")

    # 1) 견본 표
    out, cur = [], None

    def flush():
        if cur and cur.get("body", "").strip():
            title, tag_level, tag_grade = split_level_tag(cur.get("title", "").strip() or f"지문 {len(out) + 1}")
            level = parse_level(cur.get("level")) or tag_level
            grade = parse_grade_number(cur.get("grade")) or tag_grade
            if not level and grade:
                level = "초등" if grade <= 6 else "중등"
            out.append(Passage(title or f"지문 {len(out) + 1}", cur["body"].strip(), level, str(path), grade))

    for tbl in body.iter(W_NS + "tbl"):
        for tr in tbl.findall(W_NS + "tr"):
            cells = ["\n".join(_docx_paragraph_text(p) for p in tc.iter(W_NS + "p")).strip()
                     for tc in tr.findall(W_NS + "tc")]
            if len(cells) < 2:
                continue
            field_name = _docx_label(cells[0])
            if not field_name:
                continue
            value = "\n".join(c for c in cells[1:] if c)
            if field_name == "title" or cur is None or field_name in cur:
                flush()
                cur = {}
            cur[field_name] = value
        flush()
        cur = None
    if out:
        return out

    # 2) 문단 형식
    lines = []
    for p in body.iter(W_NS + "p"):
        text = _docx_paragraph_text(p).strip()
        sid = p.find(f"{W_NS}pPr/{W_NS}pStyle")
        style = style_names.get(sid.get(W_NS + "val"), "") if sid is not None else ""
        m = re.match(r"^(제목|내용|본문|학교급)\s*[:：]\s*(.*)$", text)
        if m:
            label, value = m.groups()
            if label == "제목":
                lines.append(f"# {value}")
            elif label == "학교급" and lines and lines[-1].startswith("# "):
                lines[-1] += f" [{value}]" if value else ""
            else:
                lines.append(value)
        elif text and ("heading" in style or style.startswith("제목")):
            lines.append(f"# {text}")
        else:
            lines.append(text)
    return parse_text("\n".join(lines), str(path), path.stem)


def load_passages(paths) -> list[Passage]:
    """파일/폴더 경로 목록에서 지문을 모두 읽는다."""
    if isinstance(paths, (str, Path)):
        paths = [paths]
    files = []
    for p in map(Path, paths):
        if p.is_dir():
            files += sorted(x for x in p.iterdir() if x.suffix.lower() in PASSAGE_EXTS
                            and not x.name.startswith(("~$", "ERI_", "."))
                            and x.name.lower() not in ("readme.md", "requirements.txt"))
        elif p.is_file():
            files.append(p)
        else:
            raise FileNotFoundError(f"지문 파일을 찾을 수 없습니다: {p}")
    out = []
    for f in files:
        ext = f.suffix.lower()
        if ext in (".xlsx", ".csv"):
            out += _read_table(f)
        elif ext == ".docx":
            out += _read_docx(f)
        else:
            out += parse_text(read_text(f), str(f), f.stem)
    return out


def find_passage_files(search_dirs) -> list[Path]:
    """앞 폴더부터 지문 파일 후보를 찾는다.
    이름에 '지문'이 들어간 파일(.txt .md .docx .xlsx .csv)을 우선하고, 없으면 모든 .txt를 쓴다.
    처음으로 후보가 있는 폴더의 파일만 돌려준다 (같은 파일이 여러 폴더에 있어도 한 번만 분석)."""
    for d in map(Path, search_dirs):
        if not d.is_dir():
            continue
        files = [p for p in d.iterdir() if p.is_file() and not p.name.startswith(("~$", "ERI_", "."))]
        named = sorted(p for p in files if "지문" in p.name and p.suffix.lower() in PASSAGE_EXTS)
        txts = sorted(p for p in files if p.suffix.lower() == ".txt"
                      and p.name.lower() not in ("requirements.txt", "license.txt"))
        if named or txts:
            return named or txts
    return []
