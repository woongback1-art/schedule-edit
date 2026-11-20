# -*- coding: utf-8 -*-
"""스케줄 수정 규칙 (구글 시트/엑셀 공용).

Grid: 날짜 탭 하나. get(r, c) -> 표시 글자, set(r, c, v) 로 바꾸면 writes 에 쌓임.
"""
import re

HEADER_ROW = 3     # 치료사 이름 줄
TIME_COL = 3       # C열 = 시간
FIRST_ROW = 4
MAIN_ROWS = 3      # 시간칸 위 3줄 = 그 시간 치료, 아래 = 다음 시간 앞 F 줄
AUX = re.compile(r"^\s*(└|┕|ㄴ|-)")
DATE_TAB = re.compile(r"^\d{2,4}\.\d{1,2}\.\d{1,2}$")

COLS = ["날짜", "작업", "환자", "코드", "치료사", "시간", "받는 치료사", "받는 시간", "넣을/바꿀 글자", "요청한 쌤", "결과 (자동)"]
ACTS = ["이동", "추가", "삭제", "수정"]
CODES = ["M", "N", "G", "F", "NB"]
TIMES = ["8:30", "9:00", "9:35", "10:05", "10:40", "11:10", "11:45",
         "1:15", "1:45", "2:20", "2:50", "3:25", "3:55", "4:30", "5:00"]


class MemoError(Exception):
    pass


class Grid:
    """표시값 2차원 배열 + 병합정보(C열) 를 가진 날짜 탭."""

    def __init__(self, title, values, merges_c=None):
        self.title = title
        self.v = [list(map(lambda x: "" if x is None else str(x), row)) for row in values]
        self.merges = merges_c or {}        # 시작행 -> 끝행 (C열)
        self.writes = []                    # [(r, c, 값)]
        self._slots = None

    @property
    def max_row(self):
        return len(self.v)

    def get(self, r, c):
        if r < 1 or r > len(self.v):
            return ""
        row = self.v[r - 1]
        return row[c - 1] if c - 1 < len(row) else ""

    def set(self, r, c, val):
        val = "" if val is None else str(val)
        while len(self.v) < r:
            self.v.append([])
        row = self.v[r - 1]
        while len(row) < c:
            row.append("")
        row[c - 1] = val
        self.writes.append((r, c, val))


# ---------------- 날짜 탭 ----------------
def tab_date(title):
    p = [int(x) for x in re.findall(r"\d+", title)]
    y = p[0] + 2000 if p[0] < 100 else p[0]
    return (y, p[1], p[2])


def date_tabs(titles):
    return sorted([t for t in titles if DATE_TAB.match(t.strip())], key=tab_date, reverse=True)


def pick_tab(titles, key):
    key = str(key).strip()
    if key in titles:
        return key
    nums = [int(x) for x in re.findall(r"\d+", key)]
    if not nums:
        raise MemoError(f"날짜 '{key}'를 못 읽었어요")
    if len(nums) >= 3:
        mon, day = nums[-2], nums[-1]
    elif len(nums) == 2:
        mon, day = nums
    else:
        mon, day = None, nums[0]
    hits = [t for t in date_tabs(titles) if tab_date(t)[2] == day and (mon is None or tab_date(t)[1] == mon)]
    if hits:
        return hits[0]  # 같은 일자가 여러 달이면 가장 최근 탭
    raise MemoError(f"'{key}' 날짜 탭이 없어요")


# ---------------- 구조 ----------------
def therapists(g):
    out = {}
    row = g.v[HEADER_ROW - 1] if len(g.v) >= HEADER_ROW else []
    for c in range(TIME_COL + 1, len(row) + 1):
        name = re.sub(r"[\s\d]+$", "", row[c - 1].strip())
        if re.fullmatch(r"[가-힣]{2,4}", name):
            out[c] = name
    return out


def therapist_col(g, who):
    who = str(who).strip().replace("쌤", "").replace("선생님", "")
    th = therapists(g)
    hits = [c for c, n in th.items() if n == who] or [c for c, n in th.items() if n.endswith(who) or who in n]
    if len(hits) == 1:
        return hits[0]
    if hits:
        raise MemoError(f"치료사 '{who}'가 여러 명이에요: {', '.join(th[h] for h in hits)}")
    raise MemoError(f"{g.title} 탭에 치료사 '{who}'가 없어요")


def _to_min(h, m):
    return ((h + 12) if 1 <= h <= 7 else h) * 60 + m


def slot_minutes(v):
    m = re.search(r"(오전|오후)?\s*(\d{1,2})\s*[:;시]\s*(\d{1,2})?", str(v))
    if not m:
        return None
    h, mm = int(m.group(2)), int(m.group(3) or 0)
    if m.group(1) == "오후" and h < 12:
        h += 12
    return h * 60 + mm if h > 7 else _to_min(h, mm)


def parse_time(s):
    m = re.search(r"(\d{1,2})\s*(?:[:;]|시)?\s*(\d{1,2})?", str(s))
    if not m:
        raise MemoError(f"'{s}'를 시간으로 못 읽었어요")
    return _to_min(int(m.group(1)), int(m.group(2) or 0))


def slots(g):
    if g._slots is not None:
        return g._slots
    starts = []
    for r in range(FIRST_ROW, g.max_row + 1):
        mins = slot_minutes(g.get(r, TIME_COL))
        if mins is not None:
            starts.append((r, mins))
    out = []
    for i, (r, mins) in enumerate(starts):
        end = starts[i + 1][0] - 1 if i + 1 < len(starts) else r + 6
        out.append((r, end, mins, f"{(mins // 60 - 1) % 12 + 1}:{mins % 60:02d}"))
    g._slots = out
    return out


def slot_for(g, t):
    target = parse_time(t)
    sl = slots(g)
    if not sl:
        raise MemoError(f"{g.title} 탭에서 시간칸(C열)을 못 찾았어요")
    best = min(sl, key=lambda s: abs(s[2] - target))
    if abs(best[2] - target) > 20:
        raise MemoError(f"'{t}'에 맞는 시간칸이 없어요")
    return best


def slot_label(g, row):
    sl = slots(g)
    for i, (r0, r1, _, lab) in enumerate(sl):
        if r0 <= row <= r1:
            if row >= r0 + MAIN_ROWS and i + 1 < len(sl):
                return f"{sl[i + 1][3]} 앞 F줄"
            return lab
    return f"{row}행"


def search_rows(g, t):
    sl = slots(g)
    s = slot_for(g, t)
    i = sl.index(s)
    start = sl[i - 1][0] + MAIN_ROWS if i > 0 else s[0]
    return range(start, min(s[1], s[0] + MAIN_ROWS - 1) + 1)


def code_of(text, name):
    rest = text.split(name, 1)[1] if name in text else text
    m = re.match(r"\s*([A-Za-z]+)", rest)
    return m.group(1).upper() if m else ""


def matches(val, patient, code=None):
    v = str(val or "").replace("★", "").replace("▶", "").strip().upper()
    patient = patient.strip().upper()
    if not v or not v.startswith(patient):
        return False
    tail = v[len(patient):]
    if tail and re.match(r"[가-힣]", tail):
        return False
    return not code or code_of(v, patient) == code.upper()


def find_in_col(g, col, patient, code=None, rows=None):
    rows = rows if rows is not None else range(FIRST_ROW, g.max_row + 1)
    return [r for r in rows if matches(g.get(r, col), patient, code)]


def col_name(n):
    s = ""
    while n > 0:
        n, m = divmod(n - 1, 26)
        s = chr(65 + m) + s
    return s


def cell_name(c, r):
    return f"{col_name(c)}{r}"


# ---------------- 동작 ----------------
def locate(g, therapist, patient, code, t):
    col = therapist_col(g, therapist)
    th = therapists(g)[col]
    note, hits = "", []
    if t:
        hits = find_in_col(g, col, patient, code, search_rows(g, t))
    if not hits:
        hits = find_in_col(g, col, patient, code)
        if t and len(hits) == 1:
            note = f"  ※ 입력은 {t}인데 실제로는 {slot_label(g, hits[0])} 칸에 있었어요"
    if not hits:
        raise MemoError(f"{th}쌤 칸에서 '{patient} {code or ''}'를 못 찾았어요")
    if len(hits) > 1:
        where = ", ".join(slot_label(g, r) for r in hits)
        raise MemoError(f"{th}쌤 칸에 '{patient} {code or ''}'가 여러 개예요({where}) → 시간을 적어주세요")
    return col, hits[0], note


def free_row(g, col, t):
    r0, r1, _, lab = slot_for(g, t)
    end = min(max(g.merges.get(r0, r0), r0 + 1), r0 + MAIN_ROWS - 1)
    for r in range(r0, end + 1):
        if not g.get(r, col).strip():
            return r, lab
    raise MemoError(f"{therapists(g)[col]}쌤 {lab} 칸이 이미 차 있어요")


def aux_rows(g, col, row):
    out, r = [], row + 1
    while len(out) < 3 and AUX.match(g.get(r, col)):
        out.append(r)
        r += 1
    return out


def do_move(g, p, code, src, t1, dst, t2):
    fcol, frow, note = locate(g, src, p, code, t1)
    text = g.get(frow, fcol)
    tcol = therapist_col(g, dst)
    trow, lab = free_row(g, tcol, t2 or slot_label(g, frow))
    g.set(trow, tcol, text)
    g.set(frow, fcol, "")
    th = therapists(g)
    return (f"이동 '{text}'  {th[fcol]} {slot_label(g, frow)}({cell_name(fcol, frow)}) → "
            f"{th[tcol]} {lab}({cell_name(tcol, trow)}){note}")


def do_add(g, therapist, t, text):
    col = therapist_col(g, therapist)
    row, lab = free_row(g, col, t)
    g.set(row, col, text)
    return f"추가 '{text}' → {therapists(g)[col]} {lab}({cell_name(col, row)})"


def do_delete(g, therapist, t, p, code):
    col, row, note = locate(g, therapist, p, code, t)
    old = g.get(row, col)
    extra = []
    for r in aux_rows(g, col, row):
        extra.append(f"'{g.get(r, col)}'")
        g.set(r, col, "")
    g.set(row, col, "")
    more = f" + 아래 보조줄 {', '.join(extra)}" if extra else ""
    return f"삭제 '{old}'{more}  {therapists(g)[col]} {slot_label(g, row)}({cell_name(col, row)}){note}"


def do_edit(g, therapist, t, p, code, text):
    col, row, note = locate(g, therapist, p, code, t)
    old = g.get(row, col)
    g.set(row, col, text)
    return f"수정 '{old}' → '{text}'  {therapists(g)[col]} {slot_label(g, row)}({cell_name(col, row)}){note}"


# ---------------- 입력표 줄 실행 ----------------
def run_rows(rows, titles, load_grid):
    """rows: [{'key':..., 'date','act','pat','code','th','t','th2','t2','text','who','prev'}]
    titles: 탭 이름 목록, load_grid(title) -> Grid (캐시는 여기서).
    반환: (결과 [(row, ok, msg)], {title: Grid})"""
    cache, results, last_date = {}, [], ""
    for x in rows:
        if x.get("date"):
            last_date = x["date"]
        act = x.get("act", "")
        if not act or str(x.get("prev", "")).startswith("✔"):
            continue
        try:
            if not last_date:
                raise MemoError("날짜를 적어주세요")
            title = pick_tab(titles, last_date)
            if title not in cache:
                cache[title] = load_grid(title)
            g = cache[title]
            snap = len(g.writes)
            code = (x.get("code") or "").rstrip("/").upper() or None
            pat, th, t, th2, t2, text = (x.get(k, "") for k in ("pat", "th", "t", "th2", "t2", "text"))

            def need(**kw):
                miss = [k for k, v in kw.items() if not v]
                if miss:
                    raise MemoError(f"{', '.join(miss)} 칸을 채워주세요")
            try:
                if act == "이동":
                    need(환자=pat, 치료사=th, 받는치료사=th2)
                    msg = do_move(g, pat, code, th, t or None, th2, t2 or None)
                elif act == "추가":
                    need(치료사=th, 시간=t, 넣을글자=text or pat)
                    msg = do_add(g, th, t, text or f"{pat} {code or ''}".strip())
                elif act == "삭제":
                    need(환자=pat, 치료사=th)
                    msg = do_delete(g, th, t or None, pat, code)
                elif act == "수정":
                    need(환자=pat, 치료사=th, 바꿀글자=text)
                    msg = do_edit(g, th, t or None, pat, code, text)
                else:
                    raise MemoError("작업은 이동/추가/삭제/수정 중 하나로 골라주세요")
            except MemoError:
                del g.writes[snap:]
                raise
            results.append((x, True, f"{title} {msg}"))
        except MemoError as e:
            results.append((x, False, str(e)))
    return results, cache


def patient_names(g):
    th = set(therapists(g).values())
    names = set()
    for r in range(FIRST_ROW, g.max_row + 1):
        for c in range(TIME_COL + 1, len(g.v[r - 1]) + 1):
            v = re.sub(r"^[★▶*\s]+", "", g.get(r, c))
            m = re.match(r"([가-힣]{2,4}[A-Za-z]?)(?=$|[\s(\-,/‡]|[A-Za-z])", v)
            if m and m.group(1) not in th:
                n = m.group(1)
                names.add(n[:-1] + n[-1].upper() if n[-1].isalpha() and n[-1].isascii() else n)
    return sorted(n for n in names if not (n[-1] in "NMGF" and n[:-1] in names))
