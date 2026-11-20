# -*- coding: utf-8 -*-
"""구글 시트 연결 + 입력표(엑셀) 읽고 쓰기."""
import json
import os
import re
from datetime import datetime, time as dtime

from core import Grid, MemoError, COLS, date_tabs

FORM_SHEET = "수정입력"
FORM_HEAD = 4


# ======================= 구글 시트 =======================
class GoogleBook:
    """서비스 계정 키(JSON)로 구글 시트 하나를 연다."""

    def __init__(self, key_path, url):
        import gspread  # exe 안에 들어 있음
        self._gspread = gspread
        try:
            self.client = gspread.service_account(filename=key_path)
            self.sh = self.client.open_by_url(url) if "/" in url else self.client.open_by_key(url)
        except Exception as e:  # noqa
            raise ConnectionError(friendly_error(e, key_path)) from e
        self.title = self.sh.title
        self._meta = None

    def titles(self):
        return [ws.title for ws in self.sh.worksheets()]

    def _merges(self):
        if self._meta is None:
            meta = self.sh.fetch_sheet_metadata({"fields": "sheets(properties(sheetId,title),merges)"})
            self._meta = {}
            for s in meta.get("sheets", []):
                m = {}
                for mg in s.get("merges", []):
                    if mg.get("startColumnIndex", 0) <= 2 < mg.get("endColumnIndex", 0):
                        m[mg["startRowIndex"] + 1] = mg["endRowIndex"]  # 1부터, 끝행 포함
                self._meta[s["properties"]["title"]] = m
        return self._meta

    def load(self, title):
        ws = self.sh.worksheet(title)
        values = ws.get_all_values()  # 화면에 보이는 글자 그대로
        return Grid(title, values, self._merges().get(title, {}))

    def write(self, grids):
        data = []
        for g in grids:
            last = {}
            for r, c, v in g.writes:  # 같은 칸을 여러 번 바꿨으면 마지막 값만
                last[(r, c)] = v
            for (r, c), v in last.items():
                data.append({"range": f"'{g.title}'!{_a1(c, r)}", "values": [[v]]})
        if data:
            # RAW = 글자 그대로 (날짜·숫자로 바뀌지 않음), 칸 서식은 그대로
            self.sh.values_batch_update({"valueInputOption": "RAW", "data": data})
        return len(data)


def _a1(c, r):
    s = ""
    while c > 0:
        c, m = divmod(c - 1, 26)
        s = chr(65 + m) + s
    return f"{s}{r}"


def service_email(key_path):
    try:
        with open(key_path, encoding="utf-8") as f:
            d = json.load(f)
        if d.get("type") != "service_account":
            return None
        return d.get("client_email")
    except Exception:  # noqa
        return None


def friendly_error(e, key_path=""):
    s = f"{type(e).__name__}: {e}"
    email = service_email(key_path) or "(서비스 계정 이메일)"
    if "SpreadsheetNotFound" in s or "404" in s:
        return f"시트를 못 찾았어요.\n시트 주소가 맞는지, 그리고 시트 [공유]에\n{email}\n을(를) '편집자'로 추가했는지 확인하세요."
    if "403" in s or "PERMISSION" in s.upper():
        if "has not been used" in s or "disabled" in s:
            return "구글 클라우드에서 'Google Sheets API' 사용 설정이 안 되어 있어요. 사용법.txt 2번을 확인하세요."
        return f"권한이 없어요. 시트 [공유]에\n{email}\n을(를) '편집자'로 추가하세요."
    if "invalid_grant" in s or "JWT" in s or "private_key" in s:
        return "키 파일(JSON)이 올바르지 않거나 삭제된 키예요. 새 키를 받아서 다시 지정하세요."
    if "ConnectionError" in s or "Max retries" in s or "NameResolution" in s or "timed out" in s:
        return "인터넷 연결을 확인하세요. (병원 방화벽이 구글을 막는지도 확인)"
    return "구글 시트 연결 실패: " + s[:300]


# ======================= 입력표(엑셀) =======================
def _s(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return f"{v.month}.{v.day}"
    if isinstance(v, dtime):
        return f"{v.hour}:{v.minute:02d}"
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def read_form(path):
    """입력표 엑셀 → 줄 목록."""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    ws = wb[FORM_SHEET] if FORM_SHEET in wb.sheetnames else wb.worksheets[0]
    rows = []
    for r in range(FORM_HEAD + 1, ws.max_row + 1):
        v = [_s(ws.cell(r, c).value) for c in range(1, len(COLS) + 1)]
        if not any(v[:10]):
            continue
        rows.append({"file": path, "row": r, "date": v[0], "act": v[1], "pat": v[2], "code": v[3],
                     "th": v[4], "t": v[5], "th2": v[6], "t2": v[7], "text": v[8], "who": v[9], "prev": v[10]})
    return rows


def write_results(path, results):
    """입력표 '결과' 칸에 ✔/✘ 적기. 파일이 열려 있으면 False."""
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.load_workbook(path)
    ws = wb[FORM_SHEET] if FORM_SHEET in wb.sheetnames else wb.worksheets[0]
    col = len(COLS)
    for x, ok, msg in results:
        if x["file"] != path:
            continue
        c = ws.cell(x["row"], col, ("✔ " if ok else "✘ ") + msg)
        c.font = Font(size=9, color="1E7B34" if ok else "C00000")
    try:
        wb.save(path)
        return True
    except PermissionError:
        return False
