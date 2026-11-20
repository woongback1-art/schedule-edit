# -*- coding: utf-8 -*-
"""스케줄 수정 프로그램 (구글 시트 '스케줄 수정 최종본' 직접 수정)."""
import json
import os
import queue
import shutil
import sys
import threading
import traceback

APP = "스케줄수정"
VERSION = "1.0"


def resource(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def conf_dir():
    d = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), APP)
    os.makedirs(d, exist_ok=True)
    return d


CONF = os.path.join(conf_dir(), "config.json")
KEY = os.path.join(conf_dir(), "google_key.json")


def load_conf():
    try:
        with open(CONF, encoding="utf-8") as f:
            return json.load(f)
    except Exception:  # noqa
        return {}


def save_conf(d):
    with open(CONF, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)


# ============================== 화면 ==============================
def gui(uitest_out=None):
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    import backend
    import core

    conf = load_conf()
    state = {"book": None}
    forms = [f for f in conf.get("forms", []) if os.path.exists(f)]
    jobs = queue.Queue()

    root = tk.Tk()
    root.title(f"스케줄 수정 v{VERSION}")
    root.geometry("860x640")
    root.minsize(700, 520)
    F = ("Malgun Gothic", 10)
    FB = ("Malgun Gothic", 10, "bold")

    # ---------- 백그라운드 작업 ----------
    def bg(work, done, busy_msg):
        status.set(busy_msg)
        root.config(cursor="watch")
        for b in action_buttons:
            b.config(state="disabled")

        def run():
            try:
                jobs.put((done, work(), None))
            except Exception as e:  # noqa
                jobs.put((done, None, e))
        threading.Thread(target=run, daemon=True).start()

    def poll():
        try:
            while True:
                done, res, err = jobs.get_nowait()
                root.config(cursor="")
                for b in action_buttons:
                    b.config(state="normal")
                done(res, err)
        except queue.Empty:
            pass
        root.after(100, poll)

    # ---------- ① 구글 연결 ----------
    top = tk.Frame(root, padx=12, pady=10)
    top.pack(fill="x")
    conn = tk.StringVar(value="구글 시트 연결 전")
    conn_lbl = tk.Label(top, textvariable=conn, font=FB, anchor="w")
    conn_lbl.pack(side="left", fill="x", expand=True)
    btn_settings = tk.Button(top, text="구글 연결 설정", font=F, command=lambda: open_settings())
    btn_settings.pack(side="right")

    def connect(silent=False):
        c = load_conf()
        if not (os.path.exists(KEY) and c.get("url")):
            conn.set("⚠ 구글 연결 설정이 필요해요 → 오른쪽 [구글 연결 설정]")
            conn_lbl.config(fg="#c00000")
            return

        def work():
            b = backend.GoogleBook(KEY, c["url"])
            return b, core.date_tabs(b.titles())

        def done(res, err):
            status.set("")
            if err:
                state["book"] = None
                conn.set("⚠ 연결 실패 — [구글 연결 설정]에서 확인하세요")
                conn_lbl.config(fg="#c00000")
                if not silent:
                    messagebox.showerror("연결 실패", str(err))
                return
            b, tabs = res
            state["book"] = b
            conn.set(f"✔ 연결됨: {b.title}  (날짜 탭 {len(tabs)}개, 최신 {tabs[0] if tabs else '-'})")
            conn_lbl.config(fg="#1e7b34")
        bg(work, done, "구글 시트 연결 중…")

    # ---------- ② 입력표 ----------
    mid = tk.LabelFrame(root, text=" 쌤들 입력표 (엑셀) ", font=FB, padx=10, pady=6)
    mid.pack(fill="x", padx=12)
    lb = tk.Listbox(mid, height=6, font=F, selectmode="extended")
    lb.pack(side="left", fill="both", expand=True)
    side = tk.Frame(mid)
    side.pack(side="left", padx=(8, 0), anchor="n")

    def refresh():
        lb.delete(0, "end")
        for f in forms:
            lb.insert("end", f"{os.path.basename(f)}    ({os.path.dirname(f)})")
        c = load_conf()
        c["forms"] = forms
        save_conf(c)

    def add_files():
        for p in filedialog.askopenfilenames(title="입력표 선택 (여러 개 가능)", filetypes=[("엑셀", "*.xlsx")]):
            p = os.path.normpath(p)
            if p not in forms:
                forms.append(p)
        refresh()

    def add_folder():
        d = filedialog.askdirectory(title="입력표를 모아둔 폴더")
        if not d:
            return
        n = 0
        for name in sorted(os.listdir(d)):
            p = os.path.normpath(os.path.join(d, name))
            if name.lower().endswith(".xlsx") and not name.startswith("~$") and p not in forms:
                forms.append(p)
                n += 1
        refresh()
        if not n:
            messagebox.showinfo("폴더", "그 폴더에 새 엑셀 파일이 없어요")

    def remove_sel():
        for i in reversed(lb.curselection()):
            forms.pop(i)
        refresh()

    for txt, cmd in (("파일 추가", add_files), ("폴더 통째로", add_folder), ("선택 빼기", remove_sel),
                     ("모두 비우기", lambda: (forms.clear(), refresh()))):
        tk.Button(side, text=txt, font=F, width=12, command=cmd).pack(pady=1)

    # ---------- ③ 실행 ----------
    act = tk.Frame(root, padx=12, pady=8)
    act.pack(fill="x")
    btn_run = tk.Button(act, text="미리보기 → 구글 시트에 반영", font=("Malgun Gothic", 12, "bold"),
                        bg="#2f6fdf", fg="white", activebackground="#2459bd", activeforeground="white",
                        command=lambda: run_preview())
    btn_run.pack(side="left", fill="x", expand=True, ipady=6)
    btn_html = tk.Button(act, text="쌤들 입력 화면\n만들기", font=F, command=lambda: make_input_html())
    btn_html.pack(side="left", padx=(8, 0), ipady=2)

    status = tk.StringVar()
    tk.Label(root, textvariable=status, font=("Malgun Gothic", 9), fg="#2f6fdf", anchor="w").pack(fill="x", padx=14)
    log = scrolledtext.ScrolledText(root, font=("Malgun Gothic", 10), bg="#f7f7f7", height=14)
    log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    log.tag_config("ok", foreground="#1e7b34")
    log.tag_config("bad", foreground="#c00000")
    log.tag_config("h", font=("Malgun Gothic", 10, "bold"))
    action_buttons = [btn_run, btn_html, btn_settings]

    def put(text, tag=None):
        log.insert("end", text + "\n", tag)
        log.see("end")

    def need_book():
        if state["book"] is None:
            messagebox.showwarning("연결 필요", "먼저 구글 시트에 연결하세요 (위쪽 [구글 연결 설정]).")
            return False
        return True

    def gather_rows():
        rows, errs = [], []
        for f in forms:
            try:
                rows += backend.read_form(f)
            except Exception as e:  # noqa
                errs.append(f"{os.path.basename(f)}: 읽기 실패 ({e})")
        return rows, errs

    def compute():
        rows, errs = gather_rows()
        b = state["book"]
        titles = b.titles()
        results, grids = core.run_rows(rows, titles, b.load)
        return rows, errs, results, grids

    def show_results(results, errs, title):
        log.delete("1.0", "end")
        put(title, "h")
        for e in errs:
            put("✘ " + e, "bad")
        cur = None
        for x, ok, msg in results:
            if x["file"] != cur:
                cur = x["file"]
                put(f"\n■ {os.path.basename(cur)}", "h")
            who = f"[{x['who']}] " if x.get("who") else ""
            put(f"  {x['row']}행 {'✔' if ok else '✘'} {who}{msg}", "ok" if ok else "bad")

    def run_preview():
        if not need_book():
            return
        if not forms:
            messagebox.showwarning("입력표", "입력표를 추가하세요")
            return

        def done(res, err):
            status.set("")
            if err:
                messagebox.showerror("오류", backend.friendly_error(err, KEY) if "gspread" in repr(type(err)) else str(err))
                put(traceback.format_exception_only(type(err), err)[-1], "bad")
                return
            rows, errs, results, grids = res
            okn = sum(1 for r in results if r[1])
            badn = len(results) - okn
            skipped = sum(1 for x in rows if x["act"] and x["prev"].startswith("✔"))
            show_results(results, errs, f"미리보기 — 바뀔 것 {okn}건 / 못 하는 것 {badn}건"
                                        + (f" / 이미 반영된 줄 {skipped}개는 건너뜀" if skipped else ""))
            if not okn:
                messagebox.showinfo("미리보기", "반영할 수 있는 줄이 없어요." +
                                    (" (빨간 줄 이유를 확인하세요)" if badn else ""))
                return
            if messagebox.askyesno("구글 시트에 반영할까요?",
                                   f"'{state['book'].title}'에\n성공 {okn}건을 반영합니다." +
                                   (f"\n(실패 {badn}건은 건너뜀)" if badn else "") +
                                   "\n\n아래 결과 창에서 내용을 확인했으면 [예]"):
                apply_now()
        bg(compute, done, "입력표 읽고 구글 시트 확인 중…")

    def apply_now():
        def work():
            rows, errs, results, grids = compute()  # 최신 시트로 다시 계산
            n = state["book"].write(grids.values())
            saved = {}
            for f in {x["file"] for x, _, _ in results}:
                saved[f] = backend.write_results(f, results)
            return results, errs, n, saved

        def done(res, err):
            status.set("")
            if err:
                messagebox.showerror("반영 실패", backend.friendly_error(err, KEY))
                return
            results, errs, n, saved = res
            okn = sum(1 for r in results if r[1])
            show_results(results, errs, f"반영 완료 — 성공 {okn}건 / 실패 {len(results) - okn}건 (구글 시트 {n}칸 변경)")
            locked = [os.path.basename(f) for f, s in saved.items() if not s]
            if locked:
                put("\n※ 엑셀에 열려 있어서 결과 칸을 못 적은 입력표: " + ", ".join(locked)
                    + "  → 다음에 같은 파일을 다시 넣으면 같은 줄이 또 반영되니 주의!", "bad")
            put("\n잘못 반영됐으면 구글 시트 [파일 → 버전 기록]에서 되돌릴 수 있어요.")
            messagebox.showinfo("완료", f"구글 시트에 반영했어요. (성공 {okn}건)\n입력표 결과 칸에 ✔가 적혀서 다음엔 다시 반영되지 않아요.")
        bg(work, done, "구글 시트에 쓰는 중…")

    # ---------- 쌤들 입력 화면 (html) ----------
    def make_input_html():
        if not need_book():
            return

        def work():
            b = state["book"]
            tabs = core.date_tabs(b.titles())
            if not tabs:
                raise RuntimeError("날짜 탭을 못 찾았어요")
            g = b.load(tabs[0])
            th = [n for n in core.therapists(g).values() if n != "대기"]
            return tabs[0], th, core.patient_names(g)

        def done(res, err):
            status.set("")
            if err:
                messagebox.showerror("오류", str(err))
                return
            tab, th, pts = res
            p = filedialog.asksaveasfilename(title="쌤들께 나눠줄 입력 화면 저장", initialfile="수정요청_입력.html",
                                             defaultextension=".html", filetypes=[("웹 페이지", "*.html")])
            if not p:
                return
            with open(resource("input_template.html"), encoding="utf-8") as f:
                html = f.read()
            html = html.replace("__THERAPISTS__", json.dumps(th, ensure_ascii=False)) \
                       .replace("__PATIENTS__", json.dumps(pts, ensure_ascii=False))
            with open(p, "w", encoding="utf-8") as f:
                f.write(html)
            messagebox.showinfo("완료", f"입력 화면을 만들었어요 ({tab} 탭 기준: 치료사 {len(th)}명, 환자 {len(pts)}명).\n\n"
                                      "쌤들께 이 html 파일을 나눠주세요. 크롬으로 열어서 적고 [엑셀로 저장] →\n"
                                      "그 엑셀을 정웅쌤께 보내면 이 프로그램에 넣으면 돼요.")
        bg(work, done, "최신 탭에서 명단 읽는 중…")

    # ---------- 설정 창 ----------
    def open_settings():
        w = tk.Toplevel(root)
        w.title("구글 연결 설정")
        w.geometry("640x430")
        w.transient(root)
        w.grab_set()
        c = load_conf()
        pad = {"padx": 14, "pady": 4, "anchor": "w"}

        tk.Label(w, text="① 구글 키 파일 (서비스 계정 JSON)", font=FB).pack(**pad)
        kf = tk.Frame(w)
        kf.pack(fill="x", padx=14)
        key_lbl = tk.StringVar(value="✔ 키 파일 있음" if os.path.exists(KEY) else "아직 없음")
        tk.Label(kf, textvariable=key_lbl, font=F).pack(side="left")

        email = tk.StringVar(value=backend.service_email(KEY) or "")

        def pick_key():
            p = filedialog.askopenfilename(parent=w, title="구글에서 받은 키 파일(.json)", filetypes=[("JSON", "*.json")])
            if not p:
                return
            em = backend.service_email(p)
            if not em:
                messagebox.showerror("키 파일", "서비스 계정 키(JSON)가 아니에요. 사용법.txt 4번을 확인하세요.", parent=w)
                return
            shutil.copyfile(p, KEY)
            email.set(em)
            key_lbl.set("✔ 키 파일 있음 (프로그램 안에 복사해 둠 — 원래 파일은 지워도 돼요)")
        tk.Button(kf, text="키 파일 선택", font=F, command=pick_key).pack(side="right")

        tk.Label(w, text="② 이 이메일을 구글 시트 [공유]에 '편집자'로 추가하세요", font=FB).pack(**pad)
        ef = tk.Frame(w)
        ef.pack(fill="x", padx=14)
        ent = tk.Entry(ef, textvariable=email, font=F, state="readonly")
        ent.pack(side="left", fill="x", expand=True)

        def copy_email():
            root.clipboard_clear()
            root.clipboard_append(email.get())
            messagebox.showinfo("복사", "이메일을 복사했어요. 시트 [공유] 칸에 붙여넣으세요.", parent=w)
        tk.Button(ef, text="복사", font=F, command=copy_email).pack(side="left", padx=(6, 0))

        tk.Label(w, text="③ 구글 시트 주소 ('스케줄 수정 최종본' 주소창 그대로 복사)", font=FB).pack(**pad)
        url = tk.StringVar(value=c.get("url", ""))
        tk.Entry(w, textvariable=url, font=F).pack(fill="x", padx=14)

        msg = tk.StringVar()
        tk.Label(w, textvariable=msg, font=F, fg="#2f6fdf", wraplength=600, justify="left").pack(**pad)

        def save_and_test():
            u = url.get().strip()
            if not os.path.exists(KEY):
                messagebox.showwarning("설정", "① 키 파일을 먼저 선택하세요", parent=w)
                return
            if "docs.google.com/spreadsheets" not in u:
                messagebox.showwarning("설정", "③ 구글 시트 주소를 붙여넣으세요 (https://docs.google.com/spreadsheets/... )", parent=w)
                return
            c2 = load_conf()
            c2["url"] = u
            save_conf(c2)
            w.destroy()
            connect()

        bf = tk.Frame(w)
        bf.pack(fill="x", padx=14, pady=10)
        tk.Button(bf, text="저장하고 연결 시험", font=FB, bg="#2f6fdf", fg="white", command=save_and_test).pack(side="right")
        tk.Button(bf, text="사용법 열기", font=F, command=lambda: open_manual()).pack(side="left")
        tk.Label(w, text="처음 한 번만 하면 돼요. 키 파일·주소는 이 컴퓨터에만 저장돼요.",
                 font=("Malgun Gothic", 9), fg="#777").pack(**pad)

    def open_manual():
        p = os.path.join(os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__), "사용법.txt")
        if not os.path.exists(p):
            p = resource("사용법.txt")
        try:
            os.startfile(p)
        except Exception:  # noqa
            messagebox.showinfo("사용법", p)

    refresh()
    root.after(100, poll)
    root.after(300, lambda: connect(silent=True))

    if uitest_out:  # CI 화면 시험
        def finish():
            with open(uitest_out, "w", encoding="utf-8") as f:
                f.write("UI OK: 창·버튼 생성됨 " + ", ".join(b.cget("text").replace("\n", " ") for b in action_buttons))
            root.destroy()
        root.after(1500, finish)
    root.mainloop()


# ============================== 자체 시험 ==============================
def selftest(out):
    """가짜 스케줄로 규칙 시험 (CI에서 exe 안에서 실행)."""
    import core
    lines, ok = [], True

    def grid():
        v = [["" for _ in range(8)] for _ in range(40)]
        v[2] = ["", "", "", "안연수", "박승현", "김진수", "", ""]
        for r, t in ((4, "9:00"), (10, "10:05"), (16, "10:40"), (22, "5:00")):
            v[r - 1][2] = t
        v[9][3] = "김형근 N-서식"; v[8][3] = "김형근 F,"; v[10][3] = "└서식"
        v[15][4] = "고윤상 M"
        return core.Grid("26.11.17", v, {4: 9, 10: 15, 16: 21, 22: 27})
    g = grid()
    rows = [
        {"file": "t", "row": 5, "date": "17", "act": "이동", "pat": "김형근", "code": "n", "th": "연수", "t": "10:40", "th2": "승현", "t2": "5:00", "text": "", "who": "", "prev": ""},
        {"file": "t", "row": 6, "date": "", "act": "삭제", "pat": "김형근", "code": "F", "th": "안연수", "t": "10:05", "th2": "", "t2": "", "text": "", "who": "", "prev": ""},
        {"file": "t", "row": 7, "date": "", "act": "수정", "pat": "고윤상", "code": "M", "th": "박승현", "t": "10:40", "th2": "", "t2": "", "text": "고윤상 G", "who": "", "prev": ""},
        {"file": "t", "row": 8, "date": "", "act": "추가", "pat": "", "code": "", "th": "김진수", "t": "9:00", "th2": "", "t2": "", "text": "홍길동 N", "who": "", "prev": ""},
        {"file": "t", "row": 9, "date": "", "act": "이동", "pat": "없는사람", "code": "N", "th": "연수", "t": "", "th2": "승현", "t2": "", "text": "", "who": "", "prev": ""},
        {"file": "t", "row": 10, "date": "", "act": "삭제", "pat": "김형근", "code": "N", "th": "연수", "t": "", "th2": "", "t2": "", "text": "", "who": "", "prev": "✔ 이미"},
    ]
    res, grids = core.run_rows(rows, ["26.11.17", "26.11.16"], lambda t: g)
    expect = [True, True, True, True, False]
    got = [r[1] for r in res]
    lines.append(f"결과 {got} (기대 {expect})")
    ok &= got == expect
    want = {(22, 5): "김형근 N-서식", (10, 4): "", (11, 4): "└서식", (9, 4): "", (16, 5): "고윤상 G", (4, 6): "홍길동 N"}
    for (r, c), v in want.items():
        good = g.get(r, c) == v
        ok &= good
        lines.append(f"{'OK ' if good else 'BAD'} {core.cell_name(c, r)} = '{g.get(r, c)}' (기대 '{v}')")
    for x, s, m in res:
        lines.append(f"  {x['row']} {'✔' if s else '✘'} {m}")
    try:
        import gspread  # noqa
        import openpyxl  # noqa
        lines.append("라이브러리 OK (gspread, openpyxl)")
    except Exception as e:  # noqa
        ok = False
        lines.append(f"라이브러리 없음: {e}")
    lines.append("SELFTEST " + ("PASS" if ok else "FAIL"))
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return 0 if ok else 1


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--selftest":
        sys.exit(selftest(sys.argv[2]))
    if len(sys.argv) >= 3 and sys.argv[1] == "--uitest":
        gui(uitest_out=sys.argv[2])
        sys.exit(0)
    gui()
