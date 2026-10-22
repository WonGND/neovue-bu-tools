"""BU 트렌드 리포트 v0.1 — 이벤트 관리 · 과거 Daily data 업로드 · 트렌드 리포트 (별도 exe, D4 A안)"""
from __future__ import annotations

import json
import os
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

_HERE = Path(__file__).resolve().parent
for _p in (_HERE.parent / "common", _HERE / "common", _HERE):
    if _p.exists() and str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import pandas as pd  # noqa: E402

import history_db as db  # noqa: E402
import trend_report as tr  # noqa: E402

APP_TITLE = f"BU_TREND_리포트_{tr.VERSION}"
AUTHOR = "(made by Kwakwonyong)"
SETTINGS_NAME = "BU_TREND_settings.json"
T = {"app": "#f3f4f6", "card": "#e9eef5", "panel": "#ffffff", "text": "#111827", "sub": "#6b7280",
     "accent": "#2563eb", "header": "#f8fafc", "warn": "#fef3c7", "warn_fg": "#92400e", "done": "#e5e7eb"}
METRICS = ("BU·WU", "White 휘도·CR", "White x/y")


def app_base_dir() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else _HERE


def parse_day(text: str) -> date | None:
    try:
        return datetime.strptime(text.strip(), "%Y-%m-%d").date()
    except (AttributeError, ValueError):
        return None


class EventDialog(tk.Toplevel):
    def __init__(self, master, title, values=None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        values = values or {}
        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)
        self.vars = {}
        for i, (key, hint) in enumerate((("이벤트명", ""), ("모델", "비우면 모든 모델 · 예: WINE"),
                                         ("시작일", "YYYY-MM-DD"), ("종료일", "YYYY-MM-DD"), ("메모", ""))):
            ttk.Label(frm, text=key).grid(row=i, column=0, sticky="w", pady=4)
            v = tk.StringVar(value=str(values.get(key, "") or ""))
            ttk.Entry(frm, textvariable=v, width=34).grid(row=i, column=1, sticky="ew", padx=8)
            ttk.Label(frm, text=hint, foreground=T["sub"]).grid(row=i, column=2, sticky="w")
            self.vars[key] = v
        btn = ttk.Frame(frm)
        btn.grid(row=6, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(btn, text="취소", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btn, text="저장", command=self._ok).pack(side="right")
        self.transient(master)
        self.grab_set()

    def _ok(self):
        name = self.vars["이벤트명"].get().strip()
        s, e = parse_day(self.vars["시작일"].get()), parse_day(self.vars["종료일"].get())
        if not name:
            messagebox.showwarning("입력 확인", "이벤트명을 입력하세요.", parent=self)
            return
        if not s or not e or e < s:
            messagebox.showwarning("입력 확인", "시작일·종료일을 YYYY-MM-DD 형식으로, 종료일이 시작일 이후가 되게 입력하세요.", parent=self)
            return
        self.result = {"이벤트명": name, "모델": self.vars["모델"].get().strip(), "시작일": s, "종료일": e,
                       "메모": self.vars["메모"].get().strip()}
        self.destroy()


class TrendApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(f"{APP_TITLE} {AUTHOR}")
        root.geometry("1120x940")
        root.configure(bg=T["app"])
        self.font = "Malgun Gothic" if "Malgun Gothic" in tkfont.families(root) else "TkDefaultFont"
        self.settings_path = app_base_dir() / SETTINGS_NAME
        self.settings = self._load_settings()
        self.uploads: list[dict] = []
        self.event_vars: dict[str, tk.BooleanVar] = {}
        self.metric_vars = {m: tk.BooleanVar(value=True) for m in METRICS}
        self.status = tk.StringVar(value="대기 중")
        self._style()
        self._build()
        self.refresh_events()

    # ---------------------------------------------------------------- 설정
    def _load_settings(self):
        s = {"db_path": str(app_base_dir() / "BU_History_DB.xlsx"), "report_dir": str(app_base_dir()),
             "report": dict(tr.DEFAULT_SETTINGS)}
        try:
            loaded = json.loads(self.settings_path.read_text(encoding="utf-8"))
            s.update({k: v for k, v in loaded.items() if k != "report"})
            s["report"].update(loaded.get("report", {}))
        except (OSError, ValueError):
            pass
        return s

    def _save_settings(self):
        try:
            self.settings_path.write_text(json.dumps(self.settings, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    @property
    def db_path(self):
        return self.settings["db_path"]

    # ---------------------------------------------------------------- 화면
    def _style(self):
        f = self.font
        st = ttk.Style()
        st.theme_use("clam")
        st.configure("TFrame", background=T["app"])
        st.configure("Card.TFrame", background=T["card"])
        st.configure("Header.TFrame", background=T["header"])
        st.configure("Title.TLabel", background=T["header"], foreground=T["text"], font=(f, 20, "bold"))
        st.configure("Author.TLabel", background=T["header"], foreground=T["sub"], font=(f, 10, "bold"))
        st.configure("HeadInfo.TLabel", background=T["header"], foreground=T["sub"], font=(f, 9))
        st.configure("Sec.TLabel", background=T["card"], foreground=T["text"], font=(f, 14, "bold"))
        st.configure("Cap.TLabel", background=T["card"], foreground=T["sub"], font=(f, 10, "bold"))
        st.configure("Note.TLabel", background=T["card"], foreground=T["sub"], font=(f, 9))
        st.configure("TLabel", font=(f, 10))
        st.configure("TButton", font=(f, 9, "bold"), padding=6)
        st.configure("Big.TButton", font=(f, 12, "bold"), padding=10, background=T["accent"], foreground="white")
        st.map("Big.TButton", background=[("active", "#1d4ed8"), ("disabled", "#94a3b8")])
        st.configure("TCheckbutton", background=T["card"], font=(f, 10))
        st.configure("Treeview", font=(f, 9), rowheight=26, background=T["panel"], fieldbackground=T["panel"])
        st.configure("Treeview.Heading", font=(f, 9, "bold"))
        st.configure("Status.TLabel", background=T["header"], foreground=T["text"], font=(f, 9))

    def _card(self, title, note=None):
        c = ttk.Frame(self.body, padding=12, style="Card.TFrame")
        c.pack(fill="x", pady=(0, 10))
        top = ttk.Frame(c, style="Card.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text=title, style="Sec.TLabel").pack(side="left")
        if note:
            ttk.Label(top, text=note, style="Note.TLabel").pack(side="left", padx=(10, 0), pady=(4, 0))
        return c, top

    def _build(self):
        hdr = ttk.Frame(self.root, padding=(18, 12), style="Header.TFrame")
        hdr.pack(fill="x")
        tk.Label(hdr, text="▣", font=(self.font, 16, "bold"), fg=T["accent"], bg=T["header"]).pack(side="left")
        ttk.Label(hdr, text=APP_TITLE, style="Title.TLabel").pack(side="left", padx=(10, 0))
        ttk.Label(hdr, text=AUTHOR, style="Author.TLabel").pack(side="left", padx=(5, 0), pady=(8, 0))
        ttk.Button(hdr, text="DB 변경", command=self.change_db).pack(side="right")
        self.db_label = ttk.Label(hdr, text="", style="HeadInfo.TLabel")
        self.db_label.pack(side="right", padx=(0, 10))

        self.body = ttk.Frame(self.root, padding=(14, 8))
        self.body.pack(fill="both", expand=True)

        # ① 이벤트 관리 표
        c1, t1 = self._card("① 이벤트 관리 표", "기간만 등록하면 측정일로 자동 소속 · 언제든 수정 가능")
        for text, cmd in (("삭제", self.delete_event), ("수정", self.edit_event), ("+ 이벤트 추가", self.add_event)):
            ttk.Button(t1, text=text, command=cmd).pack(side="right", padx=(6, 0))
        cols = ("이벤트명", "모델", "시작일", "종료일", "데이터 수", "메모")
        self.tv_ev = ttk.Treeview(c1, columns=cols, show="headings", height=4)
        for col, w in zip(cols, (180, 110, 110, 110, 90, 420)):
            self.tv_ev.heading(col, text=col)
            self.tv_ev.column(col, width=w, anchor="w" if col in ("이벤트명", "메모") else "center")
        self.tv_ev.pack(fill="x", pady=(10, 0))
        self.tv_ev.bind("<Double-1>", lambda _e: self.edit_event())

        # ② 업로드 미리보기
        c2, t2 = self._card("② Daily data 업로드 미리보기", "측정일: 파일명 날짜 → 다르면 직접 지정 · 경고 행은 적재하지 않음")
        ttk.Button(t2, text="폴더 선택", command=self.choose_folder).pack(side="right")
        self.folder_var = tk.StringVar()
        ttk.Entry(c2, textvariable=self.folder_var).pack(fill="x", pady=(10, 6))
        cols2 = ("파일", "측정일", "건수", "매핑 이벤트", "상태")
        self.tv_up = ttk.Treeview(c2, columns=cols2, show="headings", height=6, selectmode="extended")
        for col, w in zip(cols2, (430, 120, 70, 160, 260)):
            self.tv_up.heading(col, text=col)
            self.tv_up.column(col, width=w, anchor="w" if col in ("파일", "상태") else "center")
        self.tv_up.tag_configure("warn", background=T["warn"], foreground=T["warn_fg"])
        self.tv_up.tag_configure("done", background=T["done"], foreground=T["sub"])
        self.tv_up.pack(fill="x")
        b2 = ttk.Frame(c2, style="Card.TFrame")
        b2.pack(fill="x", pady=(8, 0))
        self.up_summary = ttk.Label(b2, text="폴더를 선택하세요", style="Cap.TLabel")
        self.up_summary.pack(side="left")
        self.btn_append = ttk.Button(b2, text="DB에 적재", command=self.append_uploads)
        self.btn_append.pack(side="right", padx=(6, 0))
        ttk.Button(b2, text="이벤트 변경", command=self.set_upload_event).pack(side="right", padx=(6, 0))
        ttk.Button(b2, text="측정일 직접 지정", command=self.set_upload_date).pack(side="right")

        # ③ 트렌드 리포트
        c3, _ = self._card("③ 트렌드 리포트", "거시: 이벤트 간 비교 · 미시: 이벤트 안 측정일별 비교")
        g = ttk.Frame(c3, style="Card.TFrame")
        g.pack(fill="x", pady=(10, 0))
        ttk.Label(g, text="비교할 이벤트", style="Cap.TLabel").grid(row=0, column=0, sticky="nw", pady=2)
        self.ev_box = ttk.Frame(g, style="Card.TFrame")
        self.ev_box.grid(row=0, column=1, sticky="w", padx=(16, 0))
        ttk.Label(g, text="지표", style="Cap.TLabel").grid(row=1, column=0, sticky="w", pady=(8, 0))
        mb = ttk.Frame(g, style="Card.TFrame")
        mb.grid(row=1, column=1, sticky="w", padx=(16, 0), pady=(8, 0))
        for i, m in enumerate(METRICS):
            ttk.Checkbutton(mb, text=m, variable=self.metric_vars[m]).grid(row=0, column=i, sticky="w", padx=(0, 18))
        self.btn_report = ttk.Button(c3, text="▶  트렌드 리포트 생성", style="Big.TButton", command=self.make_report)
        self.btn_report.pack(fill="x", pady=(12, 0))

        bar = ttk.Frame(self.root, padding=(18, 6), style="Header.TFrame")
        bar.pack(fill="x", side="bottom")
        ttk.Label(bar, textvariable=self.status, style="Status.TLabel").pack(side="left")

    def log(self, msg):
        self.root.after(0, self.status.set, msg)

    # ---------------------------------------------------------------- 이벤트
    def refresh_events(self):
        try:
            data, events = db.load(self.db_path)
        except db.DBLockedError as exc:
            messagebox.showerror("DB 오류", str(exc))
            return
        self.events = events
        counts = {}
        if len(data):
            m = db.assign_events(data, events)
            counts = m[m["이벤트"] != ""].groupby("이벤트").size().to_dict()
        n_db = len(data)
        self.db_label.config(text=f"DB: {self.db_path} (내 PC · {n_db:,}행)")
        self.tv_ev.delete(*self.tv_ev.get_children())
        for _, e in events.sort_values("시작일").iterrows():
            self.tv_ev.insert("", "end", values=(e["이벤트명"], e["모델"], e["시작일"], e["종료일"],
                                                 f"{counts.get(str(e['이벤트명']), 0):,}", e["메모"]))
        for w in self.ev_box.winfo_children():
            w.destroy()
        old = {k: v.get() for k, v in self.event_vars.items()}
        self.event_vars = {}
        for i, name in enumerate(events.sort_values("시작일")["이벤트명"].astype(str)):
            v = tk.BooleanVar(value=old.get(name, True))
            self.event_vars[name] = v
            ttk.Checkbutton(self.ev_box, text=name, variable=v).grid(row=i // 6, column=i % 6, sticky="w", padx=(0, 18))
        if not len(events):
            ttk.Label(self.ev_box, text="먼저 ①에서 이벤트를 등록하세요", style="Note.TLabel").grid(row=0, column=0)
        self.refresh_upload_status()

    def _selected_event_name(self):
        sel = self.tv_ev.selection()
        return str(self.tv_ev.item(sel[0], "values")[0]) if sel else None

    def _write_events(self, events):
        try:
            db.save_events(self.db_path, events)
        except db.DBLockedError as exc:
            messagebox.showerror("DB 오류", str(exc))
            return False
        self.refresh_events()
        return True

    def add_event(self):
        d = EventDialog(self.root, "이벤트 추가")
        self.root.wait_window(d)
        if d.result:
            if d.result["이벤트명"] in set(self.events["이벤트명"].astype(str)):
                messagebox.showwarning("중복", "같은 이름의 이벤트가 이미 있습니다.")
                return
            self._write_events(pd.concat([self.events, pd.DataFrame([d.result])], ignore_index=True))

    def edit_event(self):
        name = self._selected_event_name()
        if not name:
            return
        row = self.events[self.events["이벤트명"].astype(str) == name].iloc[0].to_dict()
        d = EventDialog(self.root, "이벤트 수정", row)
        self.root.wait_window(d)
        if d.result:
            ev = self.events.copy()
            idx = ev.index[ev["이벤트명"].astype(str) == name][0]
            for k, v in d.result.items():
                ev.at[idx, k] = v
            self._write_events(ev)

    def delete_event(self):
        name = self._selected_event_name()
        if name and messagebox.askyesno("삭제", f"'{name}' 이벤트를 삭제할까요?\n(데이터는 지워지지 않고 기간 연결만 사라집니다)"):
            self._write_events(self.events[self.events["이벤트명"].astype(str) != name])

    # ---------------------------------------------------------------- 업로드
    def choose_folder(self):
        p = filedialog.askdirectory(title="과거 Daily data 통합정리 파일 폴더")
        if not p:
            return
        self.folder_var.set(p)
        files = sorted(Path(p).glob("*.xlsx"))
        files = [f for f in files if not f.name.startswith("~$") and "트렌드리포트" not in f.name
                 and f.name != Path(self.db_path).name]
        self.up_summary.config(text=f"{len(files)}개 파일 읽는 중...")
        threading.Thread(target=self._parse_files, args=(files,), daemon=True).start()

    def _parse_files(self, files):
        ups = []
        for f in files:
            try:
                parsed = tr.parse_daily_file(f)
            except Exception as exc:          # 양식이 다른 파일은 건너뜀
                self.log(f"건너뜀: {f.name} ({exc})")
                continue
            if len(parsed["rows"]):
                ups.append({"path": str(f), "parsed": parsed, "date": parsed["date"], "manual": ""})
        self.root.after(0, self._set_uploads, ups)

    def _set_uploads(self, ups):
        self.uploads = ups
        self.refresh_upload_status()

    def refresh_upload_status(self):
        if not hasattr(self, "tv_up"):
            return
        try:
            data, _ = db.load(self.db_path)
            loaded = set(data["원본파일"].dropna().astype(str))
        except db.DBLockedError:
            loaded = set()
        self.tv_up.delete(*self.tv_up.get_children())
        n_rows = n_warn = 0
        for i, u in enumerate(self.uploads):
            p = u["parsed"]
            hits = db.events_for(u["date"], [p["model"], p["code"]], self.events)
            if Path(u["path"]).name in loaded:
                u["status"], tag, ev = "이미 적재됨", "done", ""
            elif u["manual"]:
                u["status"], tag, ev = "수동 지정", "", u["manual"]
            elif len(hits) == 1:
                u["status"], tag, ev = "정상", "", hits[0]
            elif not hits:
                u["status"], tag, ev = "⚠ 기간 밖 · 이벤트 선택 필요", "warn", "—"
            else:
                u["status"], tag, ev = "⚠ 겹침: " + ", ".join(hits), "warn", "—"
            n_warn += tag == "warn"
            n_rows += len(p["rows"])
            label = Path(u["path"]).name + ("  [LD]" if p["ld"] else "")
            self.tv_up.insert("", "end", iid=str(i), values=(label, u["date"], len(p["rows"]), ev, u["status"]), tags=(tag,))
        ready = sum(1 for u in self.uploads if u.get("status") in ("정상", "수동 지정"))
        if self.uploads:
            self.up_summary.config(text=f"{len(self.uploads)}개 파일 · {n_rows:,}건 · 경고 {n_warn}건 · 적재 대상 {ready}개")
        self.btn_append.config(text=f"DB에 적재 ({ready}개 파일)")

    def _selected_uploads(self):
        sel = self.tv_up.selection()
        if not sel:
            messagebox.showinfo("선택", "표에서 파일을 먼저 선택하세요 (Ctrl·Shift로 여러 개).")
        return [self.uploads[int(i)] for i in sel]

    def set_upload_date(self):
        ups = self._selected_uploads()
        if not ups:
            return
        text = simpledialog.askstring("측정일 지정", "측정일 (YYYY-MM-DD)", initialvalue=str(ups[0]["date"]), parent=self.root)
        d = parse_day(text or "")
        if text and not d:
            messagebox.showwarning("형식", "YYYY-MM-DD 형식으로 입력하세요.")
            return
        if d:
            for u in ups:
                u["date"] = d
            self.refresh_upload_status()

    def set_upload_event(self):
        ups = self._selected_uploads()
        if not ups:
            return
        names = ["(자동 매핑으로 되돌리기)"] + list(self.events["이벤트명"].astype(str))
        top = tk.Toplevel(self.root)
        top.title("이벤트 지정")
        frm = ttk.Frame(top, padding=14)
        frm.pack()
        ttk.Label(frm, text=f"선택한 {len(ups)}개 파일의 이벤트").pack(anchor="w")
        cb = ttk.Combobox(frm, values=names, state="readonly", width=32)
        cb.current(0)
        cb.pack(pady=8)

        def ok():
            v = cb.get()
            for u in ups:
                u["manual"] = "" if v == names[0] else v
            top.destroy()
            self.refresh_upload_status()
        ttk.Button(frm, text="적용", command=ok).pack(anchor="e")
        top.transient(self.root)
        top.grab_set()

    def append_uploads(self):
        targets = [u for u in self.uploads if u.get("status") in ("정상", "수동 지정")]
        if not targets:
            messagebox.showinfo("적재", "적재할 파일이 없습니다. 경고 행은 측정일이나 이벤트를 지정하세요.")
            return
        total = skipped = 0
        try:
            for u in targets:
                a, s = db.append(self.db_path, tr.rows_for_db(u["parsed"], u["date"], u["path"], u["manual"]))
                total, skipped = total + a, skipped + s
        except db.DBLockedError as exc:
            messagebox.showerror("DB 오류", str(exc))
            return
        self.log(f"DB 적재: {total:,}행 추가" + (f", 중복 {skipped:,}행 건너뜀" if skipped else ""))
        self.refresh_events()
        messagebox.showinfo("적재 완료", f"{len(targets)}개 파일 · {total:,}행 추가")

    # ---------------------------------------------------------------- 리포트
    def make_report(self):
        evs = [k for k, v in self.event_vars.items() if v.get()]
        mets = [m for m, v in self.metric_vars.items() if v.get()]
        if not evs:
            messagebox.showinfo("리포트", "비교할 이벤트를 하나 이상 선택하세요.")
            return
        out = filedialog.asksaveasfilename(title="리포트 저장", defaultextension=".xlsx",
                                           initialdir=self.settings.get("report_dir"),
                                           initialfile=f"트렌드리포트_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                                           filetypes=[("Excel files", "*.xlsx")])
        if not out:
            return
        self.settings["report_dir"] = str(Path(out).parent)
        self._save_settings()
        self.btn_report.config(state="disabled", text="리포트 만드는 중...")

        def work():
            try:
                path = tr.generate(self.db_path, evs, out, mets, self.settings.get("report"), self.log)
                self.root.after(0, self._report_done, path, None)
            except Exception as exc:
                self.root.after(0, self._report_done, None, exc)
        threading.Thread(target=work, daemon=True).start()

    def _report_done(self, path, exc):
        self.btn_report.config(state="normal", text="▶  트렌드 리포트 생성")
        if exc:
            messagebox.showerror("리포트 오류", str(exc))
            return
        if messagebox.askyesno("완료", "트렌드 리포트를 만들었습니다. 지금 열까요?") and hasattr(os, "startfile"):
            os.startfile(path)

    def change_db(self):
        p = filedialog.asksaveasfilename(title="누적 DB 파일 (Tool A와 같은 파일)", defaultextension=".xlsx",
                                         initialfile=Path(self.db_path).name, confirmoverwrite=False,
                                         filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.settings["db_path"] = p
            self._save_settings()
            self.refresh_events()


def main():
    root = tk.Tk()
    TrendApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
