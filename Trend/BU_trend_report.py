"""NEOVUE BU 트렌드 리포트 v0.6 — 측정일·이벤트 한 표 · 파일 끌어놓기/선택 → 이벤트 지정 · 트렌드 리포트

v0.3: NEOVUE 이름 · 리포트 '엑셀 열기'·'폴더 열기' 버튼 · 변화 감지 설명 보강 · 공용 화면 스타일
v0.4: 리포트 시트 배치 개선 (요약 핵심 그림 · 미시 접기)
v0.5: WU -100 제외·알람 · 색좌표 확대 그림 · 측정일별 색좌표 이동 · 변화 감지 문구 정리
v0.6: 측정일을 끌어서 이벤트로 옮기기 · 되돌리기 · 이벤트 수정 즉시 반영(메모리, 저장은 뒤에서) · 리포트 진행률

v0.2 변경 (사용자 피드백)
- 이벤트 관리 표와 업로드 표를 '측정일 · 이벤트' 한 표로 합침. 기간·날짜 입력 없음
- 엑셀 파일(또는 폴더)을 끌어다 놓거나 '파일 추가'로 고르면 → 이벤트만 입력하고 바로 적재
- Tool A가 쌓은 날짜는 '미지정'에 모이고, 날짜를 골라 '이벤트로 묶기'
- 누적 DB 기본 위치: 문서\\BU_Tools (Tool A와 같은 파일)
"""
from __future__ import annotations

import json
import os
import sys
import threading
import tkinter as tk
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
import ui_style as ui  # noqa: E402

try:                                            # 파일 끌어다 놓기 (없으면 '파일 추가' 버튼만)
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:                               # pragma: no cover - 설치 안 된 PC
    DND_FILES = TkinterDnD = None

APP_TITLE = f"NEOVUE_BU_TREND_리포트_{tr.VERSION}"
AUTHOR = ui.AUTHOR
SETTINGS_NAME = "BU_TREND_settings.json"
T = ui.T
METRICS = ("BU·WU", "White 휘도·CR", "White x/y")
UNASSIGNED = "미지정"


def app_base_dir() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else _HERE


def parse_day(text: str) -> date | None:
    try:
        return datetime.strptime(text.strip(), "%Y-%m-%d").date()
    except (AttributeError, ValueError):
        return None


def source_text(srcs) -> str:
    kinds = []
    s = [str(x or "") for x in srcs]
    if any(x.startswith("Tool A") for x in s):
        kinds.append("Tool A 자동")
    if any(not x.startswith("Tool A") for x in s):
        kinds.append("과거 파일")
    return " · ".join(kinds)


def files_text(names) -> str:
    u = [n for n in dict.fromkeys(str(x) for x in names if str(x or "").strip() and str(x) != "nan")]
    if not u:
        return ""
    return u[0] + (f"  외 {len(u) - 1}개" if len(u) > 1 else "")


def value_keys(df: pd.DataFrame) -> pd.Series:
    d = df.copy()
    d["LotID"] = d["LotID"].map(db.normalize_lot_id)
    return db._value_key(d)


def expand_paths(paths, db_name: str) -> list[Path]:
    """끌어다 놓은 파일·폴더 → 엑셀 파일 목록 (폴더는 바로 아래 .xlsx)"""
    out = []
    for p in map(Path, paths):
        cands = sorted(p.glob("*.xlsx")) if p.is_dir() else [p]
        for f in cands:
            n = f.name
            if f.suffix.lower() != ".xlsx" or n.startswith("~$") or "트렌드리포트" in n or n == db_name \
                    or n.endswith(".backup.xlsx") or n.endswith(".tmp.xlsx"):
                continue
            if f not in out:
                out.append(f)
    return out


# ====================================================================== 대화상자
class EventChoice(ttk.Frame):
    """새 이벤트 / 기존 이벤트 / (선택) 미지정으로 두기"""

    def __init__(self, master, events: list[str], allow_none: bool, default: str = "new", default_event: str = ""):
        super().__init__(master)
        self.mode = tk.StringVar(value=default if (default != "old" or events) else "new")
        self.new_name = tk.StringVar()
        self.memo = tk.StringVar()
        ttk.Radiobutton(self, text="새 이벤트", variable=self.mode, value="new", command=self._sync).grid(row=0, column=0, sticky="w")
        self.ent_new = ttk.Entry(self, textvariable=self.new_name, width=32)
        self.ent_new.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=3)
        ttk.Radiobutton(self, text="기존 이벤트에 추가", variable=self.mode, value="old", command=self._sync).grid(row=1, column=0, sticky="w")
        self.cb_old = ttk.Combobox(self, values=events, width=30, state="readonly")
        if events:
            self.cb_old.set(default_event if default_event in events else events[0])
        self.cb_old.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=3)
        r = 2
        if allow_none:
            ttk.Radiobutton(self, text="미지정으로 두기 (나중에 묶기)", variable=self.mode, value="none",
                            command=self._sync).grid(row=r, column=0, columnspan=2, sticky="w", pady=(2, 0))
            r += 1
        ttk.Label(self, text="메모 (새 이벤트)").grid(row=r, column=0, sticky="w", pady=(8, 0))
        self.ent_memo = ttk.Entry(self, textvariable=self.memo, width=32)
        self.ent_memo.grid(row=r, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        if not events:
            self.cb_old.configure(state="disabled")
        self._sync()

    def _sync(self):
        m = self.mode.get()
        self.ent_new.configure(state="normal" if m == "new" else "disabled")
        self.ent_memo.configure(state="normal" if m == "new" else "disabled")
        if self.cb_old.cget("values"):
            self.cb_old.configure(state="readonly" if m == "old" else "disabled")
        if m == "new":
            self.ent_new.focus_set()

    def value(self) -> tuple[str, str | None]:
        """(이벤트 이름 또는 '', 메모 또는 None). 이름이 비면 ValueError"""
        m = self.mode.get()
        if m == "none":
            return "", None
        if m == "old":
            name = self.cb_old.get().strip()
            if not name:
                raise ValueError("기존 이벤트를 고르세요.")
            return name, None
        name = self.new_name.get().strip()
        if not name:
            raise ValueError("새 이벤트 이름을 입력하세요.")
        return name, self.memo.get().strip()


class _Dialog(tk.Toplevel):
    def __init__(self, master, title):
        super().__init__(master)
        self.title(title)
        self.configure(bg=T["app"])
        self.resizable(False, False)
        self.result = None
        self.body = ttk.Frame(self, padding=16)
        self.body.pack(fill="both", expand=True)

    def buttons(self, ok_text):
        b = ttk.Frame(self.body)
        b.pack(anchor="e", pady=(14, 0))
        self.btn_ok = ttk.Button(b, text=ok_text, command=self._ok)
        self.btn_ok.pack(side="left")
        ttk.Button(b, text="취소", command=self.destroy).pack(side="left", padx=(6, 0))
        self.bind("<Escape>", lambda _e: self.destroy())
        self.transient(self.master)
        self.grab_set()

    def _ok(self):
        raise NotImplementedError


class AddFilesDialog(_Dialog):
    """파일 추가: 읽은 통합정리 파일 목록 + 이벤트 선택"""

    def __init__(self, master, items: list[dict], skipped: list[str], events: list[str]):
        super().__init__(master, "파일 추가 — 이벤트 지정")
        self.items = items
        head = f"통합정리 파일 {len(items)}개 · {sum(len(i['parsed']['rows']) for i in items):,}건"
        if skipped:
            head += f"   (양식이 아닌 파일 {len(skipped)}개는 제외)"
        ttk.Label(self.body, text=head, font=(master.app_font, 10, "bold")).pack(anchor="w")
        cols = ("측정일", "건수", "모델", "DB 상태")
        self.tv = ttk.Treeview(self.body, columns=cols, show="tree headings", height=min(8, max(3, len(items))))
        self.tv.heading("#0", text="파일")
        self.tv.column("#0", width=360, anchor="w")
        for c, w in zip(cols, (110, 70, 100, 230)):
            self.tv.heading(c, text=c)
            self.tv.column(c, width=w, anchor="w" if c == "DB 상태" else "center")
        self.tv.tag_configure("done", foreground=T["sub"])
        self.tv.pack(fill="x", pady=(10, 4))
        self._fill()
        self.tv.bind("<Double-1>", self._edit_date)
        ttk.Label(self.body, text="측정일은 파일명 날짜로 채웠습니다 · 다르면 행을 더블클릭해서 바꿀 수 있어요 (선택 사항)",
                  style="Hint.TLabel").pack(anchor="w")
        ttk.Separator(self.body).pack(fill="x", pady=10)
        ttk.Label(self.body, text="이 파일들의 이벤트", font=(master.app_font, 10, "bold")).pack(anchor="w", pady=(0, 4))
        dup_labels = {lab for i in items for lab in i["dup_labels"] if lab}
        default_event = next(iter(dup_labels)) if len(dup_labels) == 1 else ""
        self.choice = EventChoice(self.body, events, allow_none=True,
                                  default="old" if default_event else "new", default_event=default_event)
        self.choice.pack(anchor="w")
        ttk.Label(self.body, text="이미 DB에 있는 데이터는 다시 넣지 않고, 미지정이면 고른 이벤트로 묶습니다",
                  style="Hint.TLabel").pack(anchor="w", pady=(8, 0))
        self.buttons(f"{len(items)}개 파일 추가")

    def _fill(self):
        self.tv.delete(*self.tv.get_children())
        for k, it in enumerate(self.items):
            p = it["parsed"]
            n, dup = len(p["rows"]), it["dup"]
            if dup == 0:
                st = f"새로 추가 {n:,}건"
            elif dup >= n:
                labs = ", ".join(sorted({x or UNASSIGNED for x in it["dup_labels"]}))
                st = f"이미 있음 ({labs})"
            else:
                st = f"새로 {n - dup:,}건 · 이미 있음 {dup:,}건"
            model = p["model"] or p["code"]
            self.tv.insert("", "end", iid=str(k), text=Path(it["path"]).name + ("  [LD]" if p["ld"] else ""),
                           values=(str(it["date"]) + (" ✎" if it.get("edited") else ""), f"{n:,}", model, st),
                           tags=("done",) if dup >= n else ())

    def _edit_date(self, _e):
        sel = self.tv.selection()
        if not sel:
            return
        it = self.items[int(sel[0])]
        text = simpledialog.askstring("측정일", "측정일 (YYYY-MM-DD)", initialvalue=str(it["date"]), parent=self)
        d = parse_day(text or "")
        if text and not d:
            messagebox.showwarning("형식", "YYYY-MM-DD 형식으로 입력하세요.", parent=self)
            return
        if d:
            it["date"], it["edited"] = d, True
            self._fill()

    def _ok(self):
        try:
            label, memo = self.choice.value()
        except ValueError as exc:
            messagebox.showwarning("이벤트", str(exc), parent=self)
            return
        self.result = {"items": self.items, "label": label, "memo": memo}
        self.destroy()


class GroupDialog(_Dialog):
    def __init__(self, master, n_days: int, n_rows: int, days_text: str, events: list[str]):
        super().__init__(master, "이벤트로 묶기")
        ttk.Label(self.body, text=f"선택한 측정일 {n_days}일 · {n_rows:,}건", font=(master.app_font, 10, "bold")).pack(anchor="w")
        ttk.Label(self.body, text=days_text, style="Hint.TLabel").pack(anchor="w", pady=(0, 10))
        self.choice = EventChoice(self.body, events, allow_none=False)
        self.choice.pack(anchor="w")
        ttk.Label(self.body, text="기간은 묶인 날짜로 자동 계산 (날짜 입력 없음)", style="Hint.TLabel").pack(anchor="w", pady=(10, 0))
        self.buttons("묶기")

    def _ok(self):
        try:
            self.result = self.choice.value()
        except ValueError as exc:
            messagebox.showwarning("이벤트", str(exc), parent=self)
            return
        self.destroy()


class RenameDialog(_Dialog):
    def __init__(self, master, name: str, memo: str):
        super().__init__(master, "이벤트 이름·메모")
        f = ttk.Frame(self.body)
        f.pack(anchor="w")
        self.name, self.memo = tk.StringVar(value=name), tk.StringVar(value=memo)
        for i, (k, v) in enumerate((("이벤트명", self.name), ("메모", self.memo))):
            ttk.Label(f, text=k).grid(row=i, column=0, sticky="w", pady=4)
            ttk.Entry(f, textvariable=v, width=36).grid(row=i, column=1, sticky="w", padx=8)
        self.buttons("저장")

    def _ok(self):
        if not self.name.get().strip():
            messagebox.showwarning("이벤트", "이벤트명을 입력하세요.", parent=self)
            return
        self.result = (self.name.get().strip(), self.memo.get().strip())
        self.destroy()


# ====================================================================== 뒤에서 저장 (v0.6)
class DBSaver:
    """이벤트 수정은 화면(메모리)에 바로 반영하고, 엑셀 DB 저장은 이 스레드가 뒤에서 한다.
    여러 번 고치면 마지막 상태만 저장. on_done(오류 또는 None, 저장 후 파일 시각)은 저장 스레드에서 불림"""

    def __init__(self, on_done):
        self.on_done = on_done
        self._cv = threading.Condition()
        self._job = None
        self._busy = False
        threading.Thread(target=self._run, daemon=True).start()

    def save(self, path, data, events, backup=False):
        with self._cv:
            if self._job is not None and self._job[3]:
                backup = True                     # 아직 못 한 백업 요청은 유지
            self._job = (str(path), data.copy(), events.copy(), backup)
            self._cv.notify_all()

    def pending(self) -> bool:
        with self._cv:
            return self._job is not None or self._busy

    def flush(self, timeout=None) -> bool:
        with self._cv:
            return self._cv.wait_for(lambda: self._job is None and not self._busy, timeout)

    def _run(self):
        while True:
            with self._cv:
                self._cv.wait_for(lambda: self._job is not None)
                path, data, events, backup = self._job
                self._job, self._busy = None, True
            err, mtime = None, None
            try:
                if backup:
                    db.backup(path)
                db.save(path, data, events)
                mtime = os.path.getmtime(path)
            except Exception as exc:              # 엑셀에서 DB가 열려 있는 경우 등
                err = exc
            with self._cv:
                self._busy = False
                self._cv.notify_all()
            self.on_done(err, mtime)


class ProgressButton(tk.Canvas):
    """누르면 리포트 생성, 만드는 동안 버튼이 왼쪽부터 채워지며 % · 단계 표시 (v0.6)"""

    IDLE = "▶  트렌드 리포트 생성"

    def __init__(self, master, command, font):
        super().__init__(master, height=48, highlightthickness=0, bd=0, bg=T["card"], cursor="hand2")
        self.command, self.font = command, font
        self.state_ = "idle"
        self.pct, self.shown, self.step = 0.0, 0.0, ""
        self.hover = False
        self.bind("<Configure>", lambda _e: self._draw())
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda _e: self._hover(True))
        self.bind("<Leave>", lambda _e: self._hover(False))

    def _hover(self, on):
        self.hover = on
        self._draw()

    def _click(self, _e):
        if self.state_ == "idle":
            self.command()

    def _box(self, x0, x1, color):
        h = int(self["height"])
        self.create_rectangle(x0, 2, x1, h - 2, fill=color, outline=color)

    def _draw(self):
        self.delete("all")
        w, h = max(self.winfo_width(), 10), int(self["height"])
        if self.state_ == "idle":
            self._box(0, w, "#1d4ed8" if self.hover else T["accent"])
            self.create_text(w / 2, h / 2, text=self.IDLE, fill="white", font=(self.font, 12, "bold"))
        elif self.state_ == "run":
            self._box(0, w, "#dbeafe")
            self._box(0, w * self.shown / 100, "#60a5fa")
            self.create_text(w / 2, h / 2, text=f"리포트 만드는 중  {self.shown:.0f}%   ·   {self.step}",
                             fill="#0b1220", font=(self.font, 11, "bold"))
        else:
            self._box(0, w, "#16a34a")
            self.create_text(w / 2, h / 2, text="✓  리포트 완성", fill="white", font=(self.font, 12, "bold"))

    def start(self):
        self.state_, self.pct, self.shown, self.step = "run", 0.0, 0.0, "준비 중"
        self.configure(cursor="watch")
        self._tick()

    def progress(self, pct, step):
        self.pct, self.step = max(self.pct, float(pct)), step
        self._draw()

    def _tick(self):
        if self.state_ != "run":
            return
        if self.shown < self.pct:                       # 단계 사이를 부드럽게 채움
            self.shown = min(self.pct, self.shown + max(0.6, (self.pct - self.shown) * 0.25))
        elif self.shown < min(self.pct + 6, 99):        # 오래 걸리는 단계에서도 멈춰 보이지 않게 조금씩
            self.shown += 0.08
        self._draw()
        self.after(60, self._tick)

    def finish(self, ok=True):
        if ok:
            self.state_, self.shown = "done", 100
            self._draw()
            self.after(1800, self.reset)
        else:
            self.reset()

    def reset(self):
        self.state_ = "idle"
        self.configure(cursor="hand2")
        self._draw()


# ====================================================================== 메인 화면
class TrendApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(f"{APP_TITLE} {AUTHOR}")
        root.geometry("1120x900")
        ui.apply_icon(root)
        self.font = ui.font_family(root)
        root.app_font = self.font
        self.settings_path = app_base_dir() / SETTINGS_NAME
        self.settings = self._load_settings()
        self.event_vars: dict[str, tk.BooleanVar] = {}
        self.metric_vars = {m: tk.BooleanVar(value=True) for m in METRICS}
        self.status = tk.StringVar(value="대기 중")
        self.busy = False
        self.data = pd.DataFrame(columns=db.DATA_COLS)
        self.events = pd.DataFrame(columns=db.EVENT_COLS)
        self._disk_mtime = None            # 마지막으로 읽거나 저장한 DB 파일 시각 (Tool A가 바꿨는지 확인용)
        self._unsaved = False              # 저장 실패해 메모리에만 있는 변경
        self._undo: list[tuple] = []
        self._drag = None
        self.saver = DBSaver(lambda err, mtime: root.after(0, self._saved, err, mtime))
        self._style()
        self._build()
        self._enable_drop()
        self.refresh(reload=True)
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        if self._migrate_msg:
            root.after(300, lambda: messagebox.showinfo("누적 DB 위치", self._migrate_msg))

    # ---------------------------------------------------------------- 설정
    def _load_settings(self):
        s = {"db_path": "", "db_merged": [], "report_dir": str(app_base_dir()), "report": dict(tr.DEFAULT_SETTINGS)}
        try:
            loaded = json.loads(self.settings_path.read_text(encoding="utf-8"))
            s.update({k: v for k, v in loaded.items() if k != "report"})
            s["report"].update(loaded.get("report", {}))
        except (OSError, ValueError):
            pass
        self._migrate_msg = ""
        try:
            s["db_path"], self._migrate_msg = db.adopt_shared_db(s.get("db_path"), app_base_dir() / db.DB_NAME,
                                                                 s.setdefault("db_merged", []))
        except Exception as exc:                       # 합치기 실패해도 툴은 열리게
            s["db_path"] = s.get("db_path") or str(db.default_db_path())
            self._migrate_msg = f"이전 DB를 합치지 못했습니다: {exc}"
        self.settings = s
        self._save_settings()
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
        st, _ = ui.setup_style(self.root)          # 공용 스타일 (Tool A · Tool B와 같은 색 · ✓ 체크박스)
        st.configure("Recent.TLabel", background=T["card"], foreground=T["text"], font=(self.font, 9, "bold"))

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
        hdr = ui.header(self.root, APP_TITLE, self.font)
        ttk.Button(hdr, text="DB 변경", command=self.change_db).pack(side="right")
        self.db_label = ttk.Label(hdr, text="", style="HeadInfo.TLabel")
        self.db_label.pack(side="right", padx=(0, 10))

        self.body = ttk.Frame(self.root, padding=(14, 8))
        self.body.pack(fill="both", expand=True)

        # ① 측정일 · 이벤트
        self.c1, t1 = self._card("① 측정일 · 이벤트")
        for text, cmd in (("삭제", self.delete), ("묶음 풀기", self.ungroup), ("이름·메모", self.rename),
                          ("이벤트로 묶기", self.group), ("+ 파일 추가", self.add_files)):
            ttk.Button(t1, text=text, command=cmd).pack(side="right", padx=(6, 0))
        self.btn_undo = ttk.Button(t1, text="↶ 되돌리기", command=self.undo, state="disabled")
        self.btn_undo.pack(side="right", padx=(6, 12))
        cols = ("모델", "건수", "출처", "파일 · 메모")
        wrap = ttk.Frame(self.c1, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, pady=(10, 0))
        self.tv = ttk.Treeview(wrap, columns=cols, show="tree headings", height=13, selectmode="extended")
        self.tv.heading("#0", text="이벤트 / 측정일")
        self.tv.column("#0", width=330, anchor="w")
        for c, w in zip(cols, (110, 80, 150, 390)):
            self.tv.heading(c, text=c)
            self.tv.column(c, width=w, anchor="w" if c == "파일 · 메모" else "center")
        self.tv.tag_configure("grp", background=T["group"], font=(self.font, 9, "bold"))
        self.tv.tag_configure("todo", background=T["warn"], foreground=T["warn_fg"], font=(self.font, 9, "bold"))
        self.tv.tag_configure("todo_row", background=T["todo_row"])
        self.tv.tag_configure("empty", foreground=T["sub"])
        self.tv.tag_configure("drop", background="#bfdbfe", foreground="#1e3a8a", font=(self.font, 9, "bold"))
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tv.bind("<Double-1>", self._on_double)
        self.tv.bind("<ButtonPress-1>", self._drag_press)          # v0.6: 측정일을 끌어서 이벤트로 옮기기
        self.tv.bind("<B1-Motion>", self._drag_motion)
        self.tv.bind("<ButtonRelease-1>", self._drag_release)
        self.root.bind("<Control-z>", lambda _e: self.undo())
        self.hint = ttk.Label(self.c1, text="", style="Note.TLabel")
        self.hint.pack(anchor="w", pady=(6, 0))

        # ② 트렌드 리포트
        c3, _ = self._card("② 트렌드 리포트", "거시: 이벤트 간 비교 · 미시: 이벤트 안 측정일별 비교")
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
        self.btn_report = ProgressButton(c3, self.make_report, self.font)
        self.btn_report.pack(fill="x", pady=(12, 0))
        # v0.3: 만든 리포트 바로 열기
        rr = ttk.Frame(c3, style="Card.TFrame")
        rr.pack(fill="x", pady=(10, 0))
        self.recent_var = tk.StringVar()
        ttk.Label(rr, textvariable=self.recent_var, style="Recent.TLabel").pack(side="left")
        self.btn_open_folder = ttk.Button(rr, text="폴더 열기", command=self.open_report_folder)
        self.btn_open_folder.pack(side="right", padx=(6, 0))
        self.btn_open_excel = ttk.Button(rr, text="엑셀 열기", command=self.open_report)
        self.btn_open_excel.pack(side="right")
        self._update_recent()

        bar = ttk.Frame(self.root, padding=(18, 6), style="Header.TFrame")
        bar.pack(fill="x", side="bottom")
        ttk.Label(bar, textvariable=self.status, style="Status.TLabel").pack(side="left")

    def _hint_text(self):
        tail = " · 측정일을 끌어서 이벤트(또는 미지정)에 놓으면 옮겨짐 · 실수하면 ‘되돌리기’(Ctrl+Z)"
        if DND_FILES:
            return "통합정리 엑셀 파일(또는 폴더)을 이 표로 끌어다 놓거나 ‘+ 파일 추가’ → 이벤트만 입력" + tail
        return "‘+ 파일 추가’로 통합정리 파일을 고르면 이벤트만 입력" + tail

    def _enable_drop(self):
        self.hint.configure(text=self._hint_text())
        if not DND_FILES:
            return
        try:
            for w in (self.tv, self.c1, self.hint):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self.on_drop)
                w.dnd_bind("<<DropEnter>>", self._drop_enter)
                w.dnd_bind("<<DropLeave>>", self._drop_leave)
        except Exception:                              # tkdnd를 못 불러오면 버튼만
            pass

    def _drop_enter(self, event):
        self.hint.configure(text="여기에 놓으면 파일을 읽고 이벤트를 물어봅니다", style="Drop.TLabel")
        return event.action

    def _drop_leave(self, event):
        self.hint.configure(text=self._hint_text(), style="Note.TLabel")
        return event.action

    def log(self, msg):
        self.root.after(0, self.status.set, msg)

    # ---------------------------------------------------------------- 표
    # ---------------------------------------------------------------- DB (v0.6: 처음 한 번 읽고 이후는 메모리)
    def _file_mtime(self):
        try:
            return os.path.getmtime(self.db_path)
        except OSError:
            return None

    def _external_change(self) -> bool:
        """다른 프로그램(Tool A 등)이 DB 파일을 바꿨는지 — 내가 저장 중이면 비교하지 않음"""
        return not self.saver.pending() and self._file_mtime() != self._disk_mtime

    def _load_disk(self) -> bool:
        try:
            data, events = db.load(self.db_path)
            if len(events) and (events["구분"] != "묶음").any():          # v0.1 기간 이벤트 → 묶음으로 1회 전환
                data, events, n = db.migrate_period_events(data, events)
                db.save(self.db_path, data, events)
                self.log(f"v0.1 기간 이벤트를 날짜 묶음으로 전환했습니다 ({n:,}행)")
        except db.DBLockedError as exc:
            messagebox.showerror("DB 오류", str(exc))
            return False
        self.data, self.events = data, db.ensure_events(data, events)
        self._disk_mtime = self._file_mtime()
        self._unsaved = False
        self._undo.clear()
        self.btn_undo.config(state="disabled")
        return True

    def refresh(self, reload=False):
        """reload=True 또는 DB 파일이 밖에서 바뀌었을 때만 파일을 다시 읽고, 아니면 메모리로 표를 그림"""
        if (reload or (self._external_change() and not self._unsaved)) and not self._load_disk():
            return
        self._render()

    def _apply(self, fn, done_msg, backup=False):
        """fn(data, events) → (data, events): 메모리에서 바로 바꾸고 표를 다시 그린 뒤, 저장은 뒤에서"""
        if self._external_change() and not self._unsaved:   # Tool A가 그사이 DB에 넣었으면 먼저 읽기
            self._load_disk()
        try:
            data, events = fn(self.data, self.events)
        except ValueError as exc:
            messagebox.showwarning("확인", str(exc))
            return
        self._undo.append((self.data, self.events))
        del self._undo[:-20]
        self.btn_undo.config(state="normal")
        self.data, self.events = data, db.ensure_events(data, events)
        self._render()
        self.saver.save(self.db_path, self.data, self.events, backup=backup)
        self.log(done_msg + "  · 저장 중…")

    def _saved(self, err, mtime):
        if err:
            self._unsaved = True
            self.log("⚠ DB 저장 실패 — 화면 변경은 남아 있음 (DB를 엑셀에서 닫으면 다음 수정 때 함께 저장)")
            messagebox.showerror("DB 저장 실패", f"{err}\n\n화면의 변경은 그대로 있습니다. DB 파일을 닫은 뒤 아무 수정이나 "
                                 "하면(또는 되돌리기 후 다시) 함께 저장됩니다.")
            return
        if not self.saver.pending():
            self._disk_mtime, self._unsaved = mtime, False
            cur = self.status.get()
            if cur.endswith("저장 중…"):
                self.status.set(cur.replace("저장 중…", "저장됨"))

    def undo(self):
        if not self._undo or self.busy:
            return
        self.data, self.events = self._undo.pop()
        self.btn_undo.config(state="normal" if self._undo else "disabled")
        self._render()
        self.saver.save(self.db_path, self.data, self.events)
        self.log("되돌렸습니다  · 저장 중…")

    def _on_close(self):
        if self.saver.pending():
            self.status.set("DB 저장 마무리 중…")
            self.root.update_idletasks()
            self.saver.flush(timeout=30)
        if self._unsaved and not messagebox.askyesno(
                "저장 안 된 변경", "DB 저장에 실패한 변경이 있습니다 (DB 파일이 열려 있었을 수 있음).\n그래도 닫을까요?"):
            return
        self.root.destroy()

    # ---------------------------------------------------------------- 표
    def _render(self):
        data = self.data
        lab = data[db.LABEL].fillna("").astype(str).str.strip() if len(data) else pd.Series(dtype=str)
        n_un = int((lab == "").sum())
        self.db_label.config(text=f"DB: {self.db_path} (내 PC · {len(data):,}행 · 미지정 {n_un:,}행)")
        prev = set(self.tv.get_children())
        opened = {iid for iid in prev if self.tv.item(iid, "open")}
        first = not hasattr(self, "_shown")
        self._shown = True
        self.tv.delete(*self.tv.get_children())
        self.day_rows = {}                     # iid → (이벤트, key, 행 수)
        if len(data):
            d = data.assign(_lab=lab, _key=db.day_key(data))
            un = d[d._lab == ""]
            if len(un):
                days = un.groupby("_key")
                gid = "U"
                self.tv.insert("", "end", iid=gid, text=f"⚠ {UNASSIGNED}  ·  {days.ngroups}일 · {len(un):,}건  (묶어야 리포트에 들어감)",
                               open=True, tags=("todo",))
                for key, g in sorted(days, key=lambda kv: kv[0], reverse=True):
                    self._day_row(gid, "", key, g, ("todo_row",))
        for _, e in self.events.sort_values("시작일", ascending=False, na_position="last").iterrows():
            name = str(e["이벤트명"])
            g_all = data[lab == name] if len(data) else data
            gid = f"E:{name}"
            if len(g_all):
                s, t = e["시작일"], e["종료일"]
                period = str(s) if s == t else f"{s} ~ {str(t)[5:] if str(s)[:4] == str(t)[:4] else t}"
                n_days = db.day_key(g_all).nunique()
                text = f"{name}   {period}  ·  {n_days}일 · {len(g_all):,}건"
                tags = ("grp",)
            else:
                text, tags = f"{name}   (데이터 없음)", ("grp", "empty")
            self.tv.insert("", "end", iid=gid, text=text, values=(e["모델"], "", "", e["메모"]),
                           open=(gid in opened) or first or gid not in prev, tags=tags)
            if len(g_all):
                for key, g in sorted(g_all.assign(_key=db.day_key(g_all)).groupby("_key"), key=lambda kv: kv[0]):
                    self._day_row(gid, name, key, g, ())
        if not self.tv.get_children():
            self.tv.insert("", "end", iid="EMPTY", text="데이터 없음 — Tool A 실행 또는 파일 추가",
                           tags=("empty",))
        self._refresh_report_events()

    def _day_row(self, parent, label, key, g, tags):
        day, model = key.split("|", 1)
        iid = f"D:{label}:{key}"
        self.tv.insert(parent, "end", iid=iid, text=day if day != "날짜없음" else "날짜 없음",
                       values=(model, f"{len(g):,}", source_text(g["출처"]), files_text(g["원본파일"])), tags=tags)
        self.day_rows[iid] = (label, key, len(g))

    def _refresh_report_events(self):
        for w in self.ev_box.winfo_children():
            w.destroy()
        old = {k: v.get() for k, v in self.event_vars.items()}
        self.event_vars = {}
        names = self._events_with_data(ascending=True)
        for i, name in enumerate(names):
            v = tk.BooleanVar(value=old.get(name, True))
            self.event_vars[name] = v
            ttk.Checkbutton(self.ev_box, text=name, variable=v).grid(row=i // 6, column=i % 6, sticky="w", padx=(0, 18))
        if not names:
            ttk.Label(self.ev_box, text="먼저 ①에서 날짜를 이벤트로 묶거나, 파일을 추가할 때 이벤트를 입력하세요",
                      style="Note.TLabel").grid(row=0, column=0)

    def _events_with_data(self, ascending=True) -> list[str]:
        if not len(self.data):
            return []
        lab = self.data[db.LABEL].fillna("").astype(str).str.strip()
        has = set(lab) - {""}
        ev = self.events[self.events["이벤트명"].isin(has)].sort_values("시작일", ascending=ascending, na_position="last")
        return ev["이벤트명"].astype(str).tolist()

    def _event_names(self) -> list[str]:
        return self.events.sort_values("시작일", ascending=False, na_position="last")["이벤트명"].astype(str).tolist()

    def _selected_pairs(self, items=None):
        """선택한 날짜 행 → [(이벤트, key)], 그룹 행을 고르면 그 안의 날짜 전부"""
        pairs, rows = [], 0
        for iid in (self.tv.selection() if items is None else items):
            kids = [iid] if iid.startswith("D:") else list(self.tv.get_children(iid))
            for k in kids:
                if k in self.day_rows and (self.day_rows[k][0], self.day_rows[k][1]) not in pairs:
                    lab, key, n = self.day_rows[k]
                    pairs.append((lab, key))
                    rows += n
        return pairs, rows

    def _selected_event(self) -> str | None:
        for iid in self.tv.selection():
            if iid.startswith("E:"):
                return iid[2:]
            if iid.startswith("D:") and self.day_rows.get(iid, ("",))[0]:
                return self.day_rows[iid][0]
        return None

    # ---------------------------------------------------------------- 끌어서 옮기기 (v0.6)
    def _drop_label(self, iid):
        """놓을 곳 행 → 이벤트 이름 ('' = 미지정), 놓을 수 없으면 None"""
        if not iid:
            return None
        if iid == "U":
            return ""
        if iid.startswith("E:"):
            return iid[2:]
        if iid in self.day_rows:
            return self.day_rows[iid][0]
        return None

    def _drag_press(self, e):
        self._drag = None
        iid = self.tv.identify_row(e.y)
        if not iid or self.tv.identify_region(e.x, e.y) not in ("tree", "cell") or self.busy:
            return None
        if e.state & 0x0005:                           # Shift·Ctrl = 여러 개 고르기 (기본 동작)
            return None
        if iid in self.tv.selection():                 # 이미 고른 행을 잡으면 선택을 유지한 채 끌기 준비
            self._drag = {"x": e.x, "y": e.y, "items": self.tv.selection(), "moved": False, "click": iid}
            return "break"
        self._drag = {"x": e.x, "y": e.y, "items": None, "moved": False, "click": None}
        return None

    def _drag_motion(self, e):
        d = self._drag
        if not d:
            return None
        if not d["moved"]:
            if abs(e.x - d["x"]) + abs(e.y - d["y"]) < 8:
                return "break"
            d["items"] = d["items"] or self.tv.selection()
            d["pairs"], d["rows"] = self._selected_pairs(d["items"])
            if not d["pairs"]:
                self._drag = None
                return None
            d["moved"] = True
            self.tv.configure(cursor="fleur")
        h = self.tv.winfo_height()                     # 위·아래 끝에 가면 저절로 스크롤
        if e.y < 18:
            self.tv.yview_scroll(-1, "units")
        elif e.y > h - 18:
            self.tv.yview_scroll(1, "units")
        target = self._drop_label(self.tv.identify_row(e.y))
        self._mark_drop(target)
        if target is None:
            self.status.set(f"{len(d['pairs'])}일 · {d['rows']:,}건 — 이벤트 또는 미지정 위에 놓으세요")
        else:
            where = f"'{target}'" if target else UNASSIGNED
            self.status.set(f"{len(d['pairs'])}일 · {d['rows']:,}건 → {where}(으)로 옮기기")
        d["target"] = target
        return "break"

    def _mark_drop(self, target):
        """놓을 이벤트 행을 파란색 + '⬇ 여기에 놓기'로 표시 (원래 모양은 기억했다가 되돌림)"""
        gid = None if target is None else ("U" if target == "" else f"E:{target}")
        old = getattr(self, "_marked", None)
        if old and old[0] != gid and self.tv.exists(old[0]):
            self.tv.item(old[0], tags=old[1], text=old[2])
        if gid and (not old or old[0] != gid) and self.tv.exists(gid):
            tags, text = self.tv.item(gid, "tags"), self.tv.item(gid, "text")
            self.tv.item(gid, tags=("drop",), text="⬇ 여기에 놓기  ·  " + text)
            self._marked = (gid, tags, text)
        elif not gid:
            self._marked = None

    def _drag_release(self, e):
        d, self._drag = self._drag, None
        if not d:
            return None
        if not d["moved"]:
            if d["click"]:                             # 고른 행을 그냥 클릭 = 그 행만 선택
                self.tv.selection_set(d["click"])
                self.tv.focus(d["click"])
            return None
        self.tv.configure(cursor="")
        self._mark_drop(None)
        target = d.get("target")
        if target is None:
            self.status.set("옮기기 취소")
            return "break"
        pairs = [p for p in d["pairs"] if p[0] != target]
        if not pairs:
            self.status.set("이미 그 이벤트에 있습니다")
            return "break"
        where = f"'{target}'" if target else UNASSIGNED
        self._apply(lambda data, ev: db.set_label_df(data, ev, pairs, target)[:2], f"{len(pairs)}일을 {where}(으)로 옮겼습니다")
        return "break"

    def _on_double(self, _e):
        if self._selected_event() and any(i.startswith("E:") for i in self.tv.selection()):
            self.rename()
        elif any(i.startswith("D:") for i in self.tv.selection()):
            self.group()

    # ---------------------------------------------------------------- 파일 추가 (끌어놓기 · 선택)
    def on_drop(self, event):
        self._drop_leave(event)
        paths = self.root.tk.splitlist(event.data)
        self.root.after(10, self.handle_paths, paths)
        return event.action

    def add_files(self):
        paths = filedialog.askopenfilenames(title="Daily data 통합정리 파일 (여러 개 선택 가능)",
                                            filetypes=[("Excel files", "*.xlsx")])
        if paths:
            self.handle_paths(paths)

    def handle_paths(self, paths):
        if self.busy:
            messagebox.showinfo("잠시만요", "앞의 작업이 끝난 뒤 다시 시도하세요.")
            return
        files = expand_paths(paths, Path(self.db_path).name)
        if not files:
            messagebox.showinfo("파일 추가", "추가할 엑셀(.xlsx) 파일이 없습니다.")
            return
        self.busy = True
        self.log(f"{len(files)}개 파일 읽는 중...")
        threading.Thread(target=self._read_files, args=(files,), daemon=True).start()

    def _read_files(self, files):
        items, skipped = [], []
        vk_old = {}
        if len(self.data):
            for k, lab in zip(value_keys(self.data), self.data[db.LABEL].fillna("").astype(str).str.strip()):
                if k:
                    vk_old[k] = lab
        for i, f in enumerate(files, 1):
            self.log(f"파일 읽는 중 {i}/{len(files)}: {f.name}")
            try:
                parsed = tr.parse_daily_file(f)
            except Exception:                 # 양식이 다른 파일(CA410/LMK6 cleaned 등)은 제외
                skipped.append(f.name)
                continue
            if not len(parsed["rows"]):
                skipped.append(f.name)
                continue
            keys = value_keys(parsed["rows"])
            dup_labels = [vk_old[k] for k in keys if k and k in vk_old]
            items.append({"path": str(f), "parsed": parsed, "date": parsed["date"], "dup": len(dup_labels),
                          "dup_labels": sorted(set(dup_labels))})
        self.root.after(0, self._files_read, items, skipped)

    def _files_read(self, items, skipped):
        self.busy = False
        if not items:
            self.log("추가할 통합정리 파일이 없습니다")
            messagebox.showinfo("파일 추가", "Daily data 통합정리 양식 파일이 없습니다."
                                + (f"\n(양식이 아닌 파일 {len(skipped)}개 제외: {', '.join(skipped[:3])}"
                                   + (" ..." if len(skipped) > 3 else "") + ")" if skipped else ""))
            return
        self.log(f"통합정리 파일 {len(items)}개" + (f" · 양식 아님 {len(skipped)}개 제외" if skipped else ""))
        d = AddFilesDialog(self.root, items, skipped, self._event_names())
        self.root.wait_window(d)
        if d.result:
            self._append(d.result)

    def _append(self, res):
        label, memo = res["label"], res["memo"]
        self.busy = True
        self.log("DB에 적재 중...")

        if self._external_change() and not self._unsaved:
            self._load_disk()
        base_data, base_events = self.data, self.events

        def work():
            try:
                data, events = base_data, base_events
                total = {"added": 0, "skipped": 0, "relabeled": 0, "upgraded": 0}
                for it in res["items"]:
                    rows = tr.rows_for_db(it["parsed"], it["date"], it["path"], label)
                    data, st = db.merge_rows(data, rows, label_dups=bool(label))
                    for k in total:
                        total[k] += st[k]
                events = db.ensure_events(data, events, {label: memo} if (label and memo) else None)
                self.root.after(0, self._appended, len(res["items"]), label, total, None, (data, events))
            except Exception as exc:
                self.root.after(0, self._appended, 0, label, None, exc, None)
        threading.Thread(target=work, daemon=True).start()

    def _appended(self, n_files, label, total, exc, new=None):
        self.busy = False
        if exc:
            messagebox.showerror("적재 오류", str(exc))
            self.log("적재 실패")
            return
        self._undo.append((self.data, self.events))
        self.btn_undo.config(state="normal")
        self.data, self.events = new
        self._render()
        self.saver.save(self.db_path, self.data, self.events)
        where = f"이벤트 '{label}'" if label else UNASSIGNED
        msg = f"{n_files}개 파일 → {where}\n새로 추가 {total['added']:,}건"
        if total["skipped"]:
            msg += f"\n이미 DB에 있던 {total['skipped']:,}건은 다시 넣지 않음"
            if total["relabeled"]:
                msg += f" (그중 미지정 {total['relabeled']:,}건을 '{label}'로 묶음)"
        self.log(msg.replace("\n", " · "))
        messagebox.showinfo("적재 완료", msg)

    # ---------------------------------------------------------------- 묶기 · 풀기 · 이름 · 삭제
    def _guard(self):
        if self.busy:
            messagebox.showinfo("잠시만요", "앞의 작업이 끝난 뒤 다시 시도하세요.")
            return False
        return True

    def group(self):
        if not self._guard():
            return
        pairs, rows = self._selected_pairs()
        if not pairs:
            messagebox.showinfo("이벤트로 묶기", "표에서 측정일 행을 먼저 고르세요 (Ctrl·Shift로 여러 개).")
            return
        days = sorted({k.split("|")[0] for _, k in pairs})
        text = ", ".join(days[:4]) + (f" 외 {len(days) - 4}일" if len(days) > 4 else "")
        d = GroupDialog(self.root, len(pairs), rows, text, self._event_names())
        self.root.wait_window(d)
        if d.result:
            label, memo = d.result
            self._apply(lambda d_, e_: db.set_label_df(d_, e_, pairs, label, memo)[:2], f"{len(pairs)}일을 '{label}'로 묶었습니다")

    def ungroup(self):
        if not self._guard():
            return
        pairs, rows = self._selected_pairs()
        pairs = [p for p in pairs if p[0]]
        if not pairs:
            messagebox.showinfo("묶음 풀기", "이벤트 안의 측정일(또는 이벤트)을 고르세요.")
            return
        if messagebox.askyesno("묶음 풀기", f"{len(pairs)}일 · {rows:,}건을 미지정으로 되돌릴까요? (데이터는 그대로)"):
            self._apply(lambda d_, e_: db.set_label_df(d_, e_, pairs, "")[:2], "미지정으로 되돌렸습니다")

    def rename(self):
        if not self._guard():
            return
        name = self._selected_event()
        if not name:
            messagebox.showinfo("이름·메모", "이벤트(또는 그 안의 측정일)를 고르세요.")
            return
        row = self.events[self.events["이벤트명"] == name]
        d = RenameDialog(self.root, name, row["메모"].iloc[0] if len(row) else "")
        self.root.wait_window(d)
        if d.result:
            new, memo = d.result
            self._apply(lambda d_, e_: db.rename_event_df(d_, e_, name, new, memo), f"이벤트 '{new}' 저장")

    def delete(self):
        if not self._guard():
            return
        groups = [i[2:] for i in self.tv.selection() if i.startswith("E:")]
        days = [i for i in self.tv.selection() if i.startswith("D:")]
        if len(groups) == 1 and not days:
            name = groups[0]
            pairs, rows = self._selected_pairs()
            if not rows:
                if messagebox.askyesno("이벤트 삭제", f"데이터가 없는 이벤트 '{name}'를 목록에서 지울까요?"):
                    self._apply(lambda d_, e_: db.delete_event_df(d_, e_, name), f"이벤트 '{name}' 삭제")
                return
            ans = messagebox.askyesnocancel(
                "이벤트 삭제", f"이벤트 '{name}' ({rows:,}건)\n\n예: 데이터까지 DB에서 삭제\n아니오: 이벤트만 지우고 데이터는 미지정으로")
            if ans is None:
                return
            if ans:
                self._apply(lambda d_, e_: db.delete_event_df(*db.delete_rows_df(d_, e_, pairs)[:2], name),
                            f"'{name}' 데이터 {rows:,}건 삭제 (백업: .backup.xlsx)", backup=True)
            else:
                self._apply(lambda d_, e_: db.delete_event_df(d_, e_, name), f"이벤트 '{name}' 삭제 · 데이터는 미지정으로")
            return
        pairs, rows = self._selected_pairs()
        if not pairs:
            messagebox.showinfo("삭제", "지울 측정일 행(또는 이벤트 하나)을 고르세요.")
            return
        if messagebox.askyesno("데이터 삭제", f"선택한 {len(pairs)}일 · {rows:,}건을 DB에서 삭제할까요?\n"
                               "(지우기 전 DB 옆에 .backup.xlsx 복사본을 남깁니다)", icon="warning"):
            self._apply(lambda d_, e_: db.delete_rows_df(d_, e_, pairs)[:2], f"{rows:,}건 삭제 (백업: .backup.xlsx)",
                        backup=True)

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
        self.btn_report.start()

        def prog(pct, text):
            self.root.after(0, self.btn_report.progress, pct, text)

        def work():
            try:
                if self.saver.pending():               # 방금 고친 이벤트가 파일에 저장된 뒤에 리포트
                    prog(1, "이벤트 변경 저장 마무리 중")
                    self.saver.flush(timeout=60)
                path = tr.generate(self.db_path, evs, out, mets, self.settings.get("report"), self.log, progress=prog)
                self.root.after(0, self._report_done, path, None)
            except Exception as exc:
                self.root.after(0, self._report_done, None, exc)
        threading.Thread(target=work, daemon=True).start()

    def _report_done(self, path, exc):
        self.btn_report.finish(ok=exc is None)
        if exc:
            messagebox.showerror("리포트 오류", str(exc))
            return
        self.log(f"리포트 저장: {Path(path).name}")
        self.settings["last_report"] = str(path)
        self._save_settings()
        self._update_recent()
        if messagebox.askyesno("완료", "트렌드 리포트를 만들었습니다. 지금 열까요?\n(나중에 '엑셀 열기' 버튼으로도 열 수 있어요)"):
            self.open_report()

    # ---------------------------------------------------------------- 리포트 열기 (v0.3)
    def _last_report(self) -> Path | None:
        p = self.settings.get("last_report")
        return Path(p) if p and Path(p).exists() else None

    def _update_recent(self):
        p = self._last_report()
        self.recent_var.set(f"최근 리포트: {p.name}" if p else "최근 리포트: 아직 없음")
        self.btn_open_excel.config(state="normal" if p else "disabled")
        folder = p.parent if p else Path(self.settings.get("report_dir") or "")
        self.btn_open_folder.config(state="normal" if str(folder) not in ("", ".") and folder.exists() else "disabled")

    def open_report(self):
        p = self._last_report()
        if not p:
            messagebox.showinfo("엑셀 열기", "최근 리포트 파일을 찾을 수 없습니다 (옮겼거나 지웠을 수 있어요).")
            self._update_recent()
            return
        try:
            ui.open_path(p)
        except OSError as exc:
            messagebox.showerror("엑셀 열기", str(exc))

    def open_report_folder(self):
        p = self._last_report()
        folder = p.parent if p else Path(self.settings.get("report_dir") or "")
        if folder.exists():
            ui.open_path(folder)

    def change_db(self):
        p = filedialog.asksaveasfilename(title="누적 DB 파일 (Tool A와 같은 파일)", defaultextension=".xlsx",
                                         initialdir=str(Path(self.db_path).parent), initialfile=Path(self.db_path).name,
                                         confirmoverwrite=False, filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.saver.flush(timeout=30)
            self.settings["db_path"] = p
            self._save_settings()
            self.refresh(reload=True)


def main():
    root = None
    if TkinterDnD is not None:
        try:
            root = TkinterDnD.Tk()
        except Exception:                              # tkdnd 라이브러리를 못 불러오면 일반 창
            root = None
    if root is None:
        globals()["DND_FILES"] = None
        root = tk.Tk()
    TrendApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
