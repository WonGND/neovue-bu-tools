"""BU 트렌드 리포트 v0.2 — 측정일·이벤트 한 표 · 파일 끌어놓기/선택 → 이벤트 지정 · 트렌드 리포트

v0.2 변경 (사용자 피드백)
- 이벤트 관리 표와 업로드 표를 '측정일 · 이벤트' 한 표로 합침. 기간·날짜 입력 없음
- 엑셀 파일(또는 폴더)을 끌어다 놓거나 '파일 추가'로 고르면 → 이벤트만 입력하고 바로 적재
- Tool A가 쌓은 날짜는 '미지정'에 모이고, 날짜를 골라 '이벤트로 묶기'
- 누적 DB 기본 위치: 문서\\BU_Tools (Tool A와 같은 파일)
"""
from __future__ import annotations

import base64
import io
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

try:                                            # 파일 끌어다 놓기 (없으면 '파일 추가' 버튼만)
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:                               # pragma: no cover - 설치 안 된 PC
    DND_FILES = TkinterDnD = None

APP_TITLE = f"BU_TREND_리포트_{tr.VERSION}"
AUTHOR = "(made by Kwakwonyong)"
SETTINGS_NAME = "BU_TREND_settings.json"
T = {"app": "#f3f4f6", "card": "#e9eef5", "panel": "#ffffff", "text": "#111827", "sub": "#6b7280",
     "accent": "#2563eb", "header": "#f8fafc", "warn": "#fef3c7", "warn_fg": "#92400e", "done": "#e5e7eb",
     "group": "#dbe4f0", "todo_row": "#fffbeb", "drop": "#dbeafe"}
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


# ====================================================================== 메인 화면
class TrendApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(f"{APP_TITLE} {AUTHOR}")
        root.geometry("1120x900")
        root.configure(bg=T["app"])
        self.font = "Malgun Gothic" if "Malgun Gothic" in tkfont.families(root) else "TkDefaultFont"
        root.app_font = self.font
        self.settings_path = app_base_dir() / SETTINGS_NAME
        self.settings = self._load_settings()
        self.event_vars: dict[str, tk.BooleanVar] = {}
        self.metric_vars = {m: tk.BooleanVar(value=True) for m in METRICS}
        self.status = tk.StringVar(value="대기 중")
        self.busy = False
        self.data = pd.DataFrame(columns=db.DATA_COLS)
        self.events = pd.DataFrame(columns=db.EVENT_COLS)
        self._style()
        self._build()
        self._enable_drop()
        self.refresh()
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
        st.configure("Drop.TLabel", background=T["drop"], foreground=T["accent"], font=(f, 10, "bold"))
        # 기본 라벨·라디오 배경을 창 배경과 같게 (clam 기본 회색 배경 제거)
        st.configure("TLabel", font=(f, 10), background=T["app"])
        st.configure("Hint.TLabel", font=(f, 9), background=T["app"], foreground=T["sub"])
        st.configure("TRadiobutton", font=(f, 10), background=T["app"])
        st.map("TRadiobutton", background=[("active", T["app"])])
        st.map("TCombobox", fieldbackground=[("readonly", "white")], selectbackground=[("readonly", "white")],
               selectforeground=[("readonly", T["text"])])
        st.configure("TButton", font=(f, 9, "bold"), padding=6)
        st.configure("Big.TButton", font=(f, 12, "bold"), padding=10, background=T["accent"], foreground="white")
        st.map("Big.TButton", background=[("active", "#1d4ed8"), ("disabled", "#94a3b8")])
        self._check_style(st)
        st.configure("TCheckbutton", background=T["card"], font=(f, 10))
        st.map("TCheckbutton", background=[("active", T["card"])])
        st.configure("Treeview", font=(f, 9), rowheight=26, background=T["panel"], fieldbackground=T["panel"])
        st.configure("Treeview.Heading", font=(f, 9, "bold"))
        st.configure("Status.TLabel", background=T["header"], foreground=T["text"], font=(f, 9))

    def _check_style(self, st):
        """체크박스 표시를 X 대신 ✓ 로: 화면 배율에 맞춘 그림으로 표시 칸을 바꿔 끼운다."""
        if hasattr(self, "_chk_imgs"):
            return
        try:
            from PIL import Image, ImageDraw
        except ImportError:                           # Pillow가 없으면 기본 표시 유지
            return
        scale = float(self.root.tk.call("tk", "scaling")) / (96 / 72)   # 100% 화면 = 1.0
        s = max(14, round(15 * scale))

        def png(checked):
            big = s * 4
            im = Image.new("RGBA", (big, big), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            box = [4, 4, big - 5, big - 5]
            if checked:
                d.rounded_rectangle(box, radius=big // 6, fill=T["accent"])
                w = max(5, big // 9)
                d.line([(big * .26, big * .52), (big * .43, big * .69), (big * .75, big * .33)],
                       fill="white", width=w, joint="curve")
            else:
                d.rounded_rectangle(box, radius=big // 6, fill="white", outline="#9ca3af", width=max(4, big // 16))
            buf = io.BytesIO()
            im.resize((s, s), Image.LANCZOS).save(buf, "PNG")
            return tk.PhotoImage(master=self.root, data=base64.b64encode(buf.getvalue()))

        self._chk_imgs = (png(False), png(True))
        st.element_create("Tick.indicator", "image", self._chk_imgs[0], ("selected", self._chk_imgs[1]),
                          width=s + round(6 * scale), sticky="w")
        st.layout("TCheckbutton", [("Checkbutton.padding", {"sticky": "nswe", "children": [
            ("Tick.indicator", {"side": "left", "sticky": ""}),
            ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [
                ("Checkbutton.label", {"sticky": "nswe"})]})]})])

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

        # ① 측정일 · 이벤트
        self.c1, t1 = self._card("① 측정일 · 이벤트")
        for text, cmd in (("삭제", self.delete), ("묶음 풀기", self.ungroup), ("이름·메모", self.rename),
                          ("이벤트로 묶기", self.group), ("+ 파일 추가", self.add_files)):
            ttk.Button(t1, text=text, command=cmd).pack(side="right", padx=(6, 0))
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
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tv.bind("<Double-1>", self._on_double)
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
        self.btn_report = ttk.Button(c3, text="▶  트렌드 리포트 생성", style="Big.TButton", command=self.make_report)
        self.btn_report.pack(fill="x", pady=(12, 0))

        bar = ttk.Frame(self.root, padding=(18, 6), style="Header.TFrame")
        bar.pack(fill="x", side="bottom")
        ttk.Label(bar, textvariable=self.status, style="Status.TLabel").pack(side="left")

    def _hint_text(self):
        tail = " · Tool A 실행분은 자동으로 ‘미지정’에 → 날짜를 골라(Ctrl·Shift) ‘이벤트로 묶기’"
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
    def refresh(self):
        try:
            data, events = db.load(self.db_path)
            if len(events) and (events["구분"] != "묶음").any():          # v0.1 기간 이벤트 → 묶음으로 1회 전환
                data, events, n = db.migrate_period_events(data, events)
                db.save(self.db_path, data, events)
                self.log(f"v0.1 기간 이벤트를 날짜 묶음으로 전환했습니다 ({n:,}행)")
        except db.DBLockedError as exc:
            messagebox.showerror("DB 오류", str(exc))
            return
        self.data, self.events = data, db.ensure_events(data, events)
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

    def _selected_pairs(self):
        """선택한 날짜 행 → [(이벤트, key)], 그룹 행을 고르면 그 안의 날짜 전부"""
        pairs, rows = [], 0
        for iid in self.tv.selection():
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

        def work():
            try:
                data, events = db.load(self.db_path)
                total = {"added": 0, "skipped": 0, "relabeled": 0, "upgraded": 0}
                for it in res["items"]:
                    rows = tr.rows_for_db(it["parsed"], it["date"], it["path"], label)
                    data, st = db.merge_rows(data, rows, label_dups=bool(label))
                    for k in total:
                        total[k] += st[k]
                events = db.ensure_events(data, events, {label: memo} if (label and memo) else None)
                db.save(self.db_path, data, events)
                self.root.after(0, self._appended, len(res["items"]), label, total, None)
            except Exception as exc:
                self.root.after(0, self._appended, 0, label, None, exc)
        threading.Thread(target=work, daemon=True).start()

    def _appended(self, n_files, label, total, exc):
        self.busy = False
        if exc:
            messagebox.showerror("적재 오류", str(exc))
            self.log("적재 실패")
            return
        self.refresh()
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
            self._db_call(lambda: db.set_label(self.db_path, pairs, label, memo), f"{len(pairs)}일을 '{label}'로 묶었습니다")

    def ungroup(self):
        if not self._guard():
            return
        pairs, rows = self._selected_pairs()
        pairs = [p for p in pairs if p[0]]
        if not pairs:
            messagebox.showinfo("묶음 풀기", "이벤트 안의 측정일(또는 이벤트)을 고르세요.")
            return
        if messagebox.askyesno("묶음 풀기", f"{len(pairs)}일 · {rows:,}건을 미지정으로 되돌릴까요? (데이터는 그대로)"):
            self._db_call(lambda: db.set_label(self.db_path, pairs, ""), "미지정으로 되돌렸습니다")

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
            self._db_call(lambda: db.rename_event(self.db_path, name, new, memo), f"이벤트 '{new}' 저장")

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
                    self._db_call(lambda: db.delete_event(self.db_path, name), f"이벤트 '{name}' 삭제")
                return
            ans = messagebox.askyesnocancel(
                "이벤트 삭제", f"이벤트 '{name}' ({rows:,}건)\n\n예: 데이터까지 DB에서 삭제\n아니오: 이벤트만 지우고 데이터는 미지정으로")
            if ans is None:
                return
            if ans:
                self._db_call(lambda: (db.delete_rows(self.db_path, pairs), db.delete_event(self.db_path, name)),
                              f"'{name}' 데이터 {rows:,}건 삭제 (백업: .backup.xlsx)")
            else:
                self._db_call(lambda: db.delete_event(self.db_path, name), f"이벤트 '{name}' 삭제 · 데이터는 미지정으로")
            return
        pairs, rows = self._selected_pairs()
        if not pairs:
            messagebox.showinfo("삭제", "지울 측정일 행(또는 이벤트 하나)을 고르세요.")
            return
        if messagebox.askyesno("데이터 삭제", f"선택한 {len(pairs)}일 · {rows:,}건을 DB에서 삭제할까요?\n"
                               "(지우기 전 DB 옆에 .backup.xlsx 복사본을 남깁니다)", icon="warning"):
            self._db_call(lambda: db.delete_rows(self.db_path, pairs), f"{rows:,}건 삭제 (백업: .backup.xlsx)")

    def _db_call(self, fn, done_msg):
        try:
            fn()
        except db.DBLockedError as exc:
            messagebox.showerror("DB 오류", str(exc))
            return
        except ValueError as exc:
            messagebox.showwarning("확인", str(exc))
            return
        self.refresh()
        self.log(done_msg)

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
        self.log(f"리포트 저장: {Path(path).name}")
        if messagebox.askyesno("완료", "트렌드 리포트를 만들었습니다. 지금 열까요?") and hasattr(os, "startfile"):
            os.startfile(path)

    def change_db(self):
        p = filedialog.asksaveasfilename(title="누적 DB 파일 (Tool A와 같은 파일)", defaultextension=".xlsx",
                                         initialdir=str(Path(self.db_path).parent), initialfile=Path(self.db_path).name,
                                         confirmoverwrite=False, filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.settings["db_path"] = p
            self._save_settings()
            self.refresh()


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
