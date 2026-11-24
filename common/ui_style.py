"""NEOVUE BU Tools 공용 화면 스타일 — Tool A(OneClick) · 트렌드 리포트가 같은 모양을 쓰도록.

Tool B(BU_DATA 정리)의 밝은 테마와 같은 색을 쓴다: 연회색 바탕 · 흰 머리글 · 파란 강조 · 카드 묶음.
"""
from __future__ import annotations

import base64
import io
import os
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import ttk

BRAND = "NEOVUE"
AUTHOR = "(made by Kwakwonyong)"
ICON_NAME = "neovue_bu_data.ico"

T = {"app": "#f3f4f6", "card": "#e9eef5", "panel": "#ffffff", "text": "#111827", "sub": "#6b7280",
     "accent": "#2563eb", "accent_dark": "#1d4ed8", "header": "#f8fafc", "warn": "#fef3c7", "warn_fg": "#92400e",
     "done": "#e5e7eb", "group": "#dbe4f0", "todo_row": "#fffbeb", "drop": "#dbeafe",
     "log_bg": "#0b1220", "log_fg": "#dbeafe", "ok": "#15803d", "err": "#b91c1c"}


def font_family(root) -> str:
    fams = set(tkfont.families(root))
    for f in ("Malgun Gothic", "맑은 고딕", "Noto Sans KR", "Noto Sans CJK KR"):
        if f in fams:
            return f
    return "TkDefaultFont"


def icon_path() -> Path | None:
    """exe 안(_MEIPASS) → exe 옆 → 소스 폴더 순으로 아이콘을 찾는다"""
    here = Path(__file__).resolve().parent
    cands = [Path(getattr(sys, "_MEIPASS", "")) / ICON_NAME if getattr(sys, "frozen", False) else None,
             Path(sys.executable).resolve().parent / ICON_NAME if getattr(sys, "frozen", False) else None,
             here.parent / "BU_LotID" / ICON_NAME, here / ICON_NAME]
    for p in cands:
        if p and p.exists():
            return p
    return None


def apply_icon(root) -> None:
    p = icon_path()
    if p and sys.platform.startswith("win"):
        try:
            root.iconbitmap(str(p))
        except tk.TclError:
            pass


def open_path(path) -> None:
    """파일·폴더를 기본 프로그램으로 열기 (Windows: 탐색기·엑셀)"""
    path = str(path)
    if hasattr(os, "startfile"):
        os.startfile(path)                       # noqa: S606 - 사용자 PC의 결과 파일
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def setup_style(root) -> tuple[ttk.Style, str]:
    """공용 ttk 스타일. 반환: (style, 글꼴 이름)"""
    f = font_family(root)
    root.configure(bg=T["app"])
    st = ttk.Style(root)
    st.theme_use("clam")
    st.configure("TFrame", background=T["app"])
    st.configure("Card.TFrame", background=T["card"])
    st.configure("Panel.TFrame", background=T["panel"])
    st.configure("Header.TFrame", background=T["header"])
    st.configure("Title.TLabel", background=T["header"], foreground=T["text"], font=(f, 20, "bold"))
    st.configure("Author.TLabel", background=T["header"], foreground=T["sub"], font=(f, 10, "bold"))
    st.configure("HeadInfo.TLabel", background=T["header"], foreground=T["sub"], font=(f, 9))
    st.configure("HeadStatus.TLabel", background=T["header"], foreground=T["accent"], font=(f, 11, "bold"))
    st.configure("Sec.TLabel", background=T["card"], foreground=T["text"], font=(f, 14, "bold"))
    st.configure("Cap.TLabel", background=T["card"], foreground=T["sub"], font=(f, 10, "bold"))
    st.configure("Note.TLabel", background=T["card"], foreground=T["sub"], font=(f, 9))
    st.configure("File.TLabel", background=T["card"], foreground=T["text"], font=(f, 10, "bold"))
    st.configure("Drop.TLabel", background=T["drop"], foreground=T["accent"], font=(f, 10, "bold"))
    st.configure("TLabel", font=(f, 10), background=T["app"])
    st.configure("Hint.TLabel", font=(f, 9), background=T["app"], foreground=T["sub"])
    st.configure("TRadiobutton", font=(f, 10), background=T["app"])
    st.map("TRadiobutton", background=[("active", T["app"])])
    st.map("TCombobox", fieldbackground=[("readonly", "white")], selectbackground=[("readonly", "white")],
           selectforeground=[("readonly", T["text"])])
    st.configure("TEntry", padding=5)
    st.configure("TButton", font=(f, 9, "bold"), padding=6)
    st.configure("Big.TButton", font=(f, 12, "bold"), padding=10, background=T["accent"], foreground="white")
    st.map("Big.TButton", background=[("active", T["accent_dark"]), ("disabled", "#94a3b8")],
           foreground=[("disabled", "white")])
    _check_style(root, st)
    st.configure("TCheckbutton", background=T["card"], font=(f, 10))
    st.map("TCheckbutton", background=[("active", T["card"])])
    st.configure("Treeview", font=(f, 9), rowheight=26, background=T["panel"], fieldbackground=T["panel"])
    st.configure("Treeview.Heading", font=(f, 9, "bold"))
    st.configure("Status.TLabel", background=T["header"], foreground=T["text"], font=(f, 9))
    return st, f


def _check_style(root, st):
    """체크박스 표시를 X 대신 ✓ 로: 화면 배율에 맞춘 그림으로 표시 칸을 바꿔 끼운다."""
    if getattr(root, "_chk_imgs", None):
        return
    try:
        from PIL import Image, ImageDraw
    except ImportError:                           # Pillow가 없으면 기본 표시 유지
        return
    scale = float(root.tk.call("tk", "scaling")) / (96 / 72)   # 100% 화면 = 1.0
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
        return tk.PhotoImage(master=root, data=base64.b64encode(buf.getvalue()))

    root._chk_imgs = (png(False), png(True))
    st.element_create("Tick.indicator", "image", root._chk_imgs[0], ("selected", root._chk_imgs[1]),
                      width=s + round(6 * scale), sticky="w")
    st.layout("TCheckbutton", [("Checkbutton.padding", {"sticky": "nswe", "children": [
        ("Tick.indicator", {"side": "left", "sticky": ""}),
        ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [
            ("Checkbutton.label", {"sticky": "nswe"})]})]})])


def header(root, title: str, font: str) -> ttk.Frame:
    """▣ 제목 (made by …) 머리글. 오른쪽에 붙일 위젯은 반환된 프레임에 side='right'로 pack"""
    hdr = ttk.Frame(root, padding=(18, 12), style="Header.TFrame")
    hdr.pack(fill="x")
    tk.Label(hdr, text="▣", font=(font, 16, "bold"), fg=T["accent"], bg=T["header"]).pack(side="left")
    ttk.Label(hdr, text=title, style="Title.TLabel").pack(side="left", padx=(10, 0))
    ttk.Label(hdr, text=AUTHOR, style="Author.TLabel").pack(side="left", padx=(5, 0), pady=(8, 0))
    return hdr


def card(parent, title: str, note: str | None = None, **pack):
    """제목이 있는 카드. 반환: (카드, 제목 줄)"""
    c = ttk.Frame(parent, padding=12, style="Card.TFrame")
    c.pack(fill=pack.pop("fill", "x"), pady=pack.pop("pady", (0, 10)), **pack)
    top = ttk.Frame(c, style="Card.TFrame")
    top.pack(fill="x")
    ttk.Label(top, text=title, style="Sec.TLabel").pack(side="left")
    if note:
        ttk.Label(top, text=note, style="Note.TLabel").pack(side="left", padx=(10, 0), pady=(4, 0))
    return c, top


# ------------------------------------------------------------------ 노트북 화면 대응 (v1.6 / v0.7 / v0.7)
def fit_window(root, width: int, height: int, min_w: int = 720, min_h: int = 560) -> None:
    """원하는 크기로 열되 화면(작업 표시줄 제외)보다 크면 화면에 맞추고 가운데 배치.
    1366×768 노트북·125% 배율에서도 창이 화면 밖으로 나가지 않게"""
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    w, h = min(width, sw - 40), min(height, sh - 90)
    x, y = max(0, (sw - w) // 2), max(0, (sh - h) // 2 - 20)
    root.geometry(f"{w}x{h}+{x}+{y}")
    root.minsize(min(min_w, w), min(min_h, h))


class ScrollFrame(ttk.Frame):
    """세로 스크롤 몸통. 창이 넉넉하면 안쪽이 창 높이만큼 늘어나고(표·로그가 커짐),
    작으면 스크롤 막대가 생긴다. 안쪽 위젯은 self.body 에 붙인다."""

    def __init__(self, parent, bg=None, **kw):
        super().__init__(parent, **kw)
        bg = bg or T["app"]
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, bg=bg)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.body = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window(0, 0, window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda _e: self._fit())
        self.canvas.bind("<Configure>", lambda _e: self._fit())
        self.canvas.bind("<Enter>", lambda _e: self.canvas.bind_all("<MouseWheel>", self._wheel))
        self.canvas.bind("<Leave>", lambda _e: self.canvas.unbind_all("<MouseWheel>"))
        self._bar_on = False

    def _fit(self):
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        need = self.body.winfo_reqheight()
        self.canvas.itemconfigure(self._win, width=cw, height=max(ch, need))
        self.canvas.configure(scrollregion=(0, 0, cw, max(ch, need)))
        show = need > ch + 2
        if show != self._bar_on:
            self._bar_on = show
            if show:
                self.vbar.pack(side="right", fill="y", before=self.canvas)
            else:
                self.vbar.pack_forget()
                self.canvas.yview_moveto(0)

    def _wheel(self, e):
        w = e.widget
        try:                                     # 표·로그 위에서는 그 위젯이 스크롤
            if str(w.winfo_class()) in ("Treeview", "Text", "TCombobox", "Listbox"):
                return None
        except (AttributeError, tk.TclError):
            pass
        if self._bar_on:
            self.canvas.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1), "units")
        return None


# ------------------------------------------------------------------ 오류를 쉬운 말로
def _chain(exc):
    seen = set()
    while exc is not None and id(exc) not in seen:
        seen.add(id(exc))
        yield exc
        exc = exc.__cause__ or exc.__context__


def friendly_error(exc) -> str:
    """오류 → '무엇이 문제고 어떻게 하면 되는지' 문장. 원문은 마지막 줄에 작게 남김"""
    raw = f"{type(exc).__name__}: {exc}"
    for e in _chain(exc):
        name = getattr(e, "filename", None)
        if isinstance(e, PermissionError) or getattr(e, "errno", None) == 13:
            f = f"\n파일: {name}" if name else ""
            return (f"파일이 엑셀 등 다른 프로그램에서 열려 있어 저장할 수 없습니다.{f}\n\n"
                    f"→ 엑셀에서 그 파일을 닫고 다시 실행하세요.\n\n(원문: {raw})")
        if type(e).__name__ == "DBLockedError":
            return f"{e}\n\n→ 누적 DB 파일을 엑셀에서 닫고 다시 실행하세요."
        if isinstance(e, FileNotFoundError):
            f = f"\n경로: {name}" if name else f"\n{e}"
            return f"파일이나 폴더를 찾을 수 없습니다.{f}\n\n→ 옮기거나 이름을 바꾸지 않았는지 확인하고 다시 고르세요."
        if isinstance(e, UnicodeDecodeError):
            return ("CSV 파일의 글자 형식을 읽지 못했습니다.\n\n→ 설비에서 나온 원본 CSV인지 확인하세요 "
                    f"(엑셀에서 다른 형식으로 다시 저장한 파일이면 원본으로 바꿔 주세요).\n\n(원문: {raw})")
        if isinstance(e, KeyError):
            return (f"필요한 열 {e} 이(가) 파일에 없습니다.\n\n→ CA410·LMK6 칸에 맞는 설비 원본 파일을 넣었는지 확인하세요."
                    f"\n\n(원문: {raw})")
    return f"{exc}\n\n(자세한 내용은 처리 로그를 확인하세요)"


# ------------------------------------------------------------------ 사용법
def find_manual(base_dir) -> Path | None:
    """exe 옆의 사용 Manual (이름에 Manual·매뉴얼·사용법이 들어간 pptx/pdf/docx)"""
    try:
        for p in sorted(Path(base_dir).iterdir()):
            if p.suffix.lower() in (".pptx", ".pdf", ".docx") and any(
                    k in p.stem.lower() for k in ("manual", "매뉴얼", "사용법")):
                return p
    except OSError:
        pass
    return None


def show_help(root, title: str, sections, base_dir=None) -> None:
    """사용법 창: sections = [(소제목, [줄, …]), …]. exe 옆에 Manual 파일이 있으면 '사용 Manual 열기' 버튼"""
    font = font_family(root)
    win = tk.Toplevel(root)
    win.title(f"{title} — 사용법")
    win.configure(bg=T["app"])
    win.transient(root)
    fit_window(win, 640, 620, 480, 360)
    bar = ttk.Frame(win, padding=(14, 10), style="Header.TFrame")
    bar.pack(fill="x", side="bottom")
    manual = find_manual(base_dir) if base_dir else None
    if manual:
        ttk.Button(bar, text="사용 Manual 열기", command=lambda: open_path(manual)).pack(side="left")
    else:
        ttk.Label(bar, text="exe 옆에 이름에 'Manual'이 들어간 파일을 두면 여기서 바로 열 수 있습니다",
                  style="HeadInfo.TLabel").pack(side="left")
    ttk.Button(bar, text="닫기", command=win.destroy).pack(side="right")
    txt = tk.Text(win, wrap="word", relief="flat", padx=18, pady=14, bg="white", font=(font, 10), spacing1=2, spacing3=2)
    sb = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    txt.pack(fill="both", expand=True)
    txt.tag_configure("h", font=(font, 12, "bold"), foreground=T["accent"], spacing1=10, spacing3=4)
    for head, lines in sections:
        txt.insert("end", head + "\n", "h")
        for ln in lines:
            txt.insert("end", f"  {ln}\n")
    txt.configure(state="disabled")


def help_button(header_frame, command) -> ttk.Button:
    b = ttk.Button(header_frame, text="? 사용법", command=command)
    b.pack(side="right", padx=(8, 0))
    return b


class ProgressButton(tk.Canvas):
    """큰 실행 버튼. 누르면 command, 실행 중에는 버튼이 왼쪽부터 채워지며 % · 단계 표시, 끝나면 초록 완료.
    (트렌드 v0.6에서 시작 → 세 툴 공용)"""

    def __init__(self, master, command, font, idle="▶  실행", running="처리 중", done="✓  완료", bg=None, height=48):
        super().__init__(master, height=height, highlightthickness=0, bd=0, bg=bg or T["card"], cursor="hand2")
        self.command, self.font = command, font
        self.IDLE, self.RUNNING, self.DONE = idle, running, done
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
            self.create_text(w / 2, h / 2, text=f"{self.RUNNING}  {self.shown:.0f}%   ·   {self.step}",
                             fill="#0b1220", font=(self.font, 11, "bold"))
        else:
            self._box(0, w, "#16a34a")
            self.create_text(w / 2, h / 2, text=self.DONE, fill="white", font=(self.font, 12, "bold"))

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




def tooltip(widget, text_fn) -> None:
    """마우스를 올리면 작은 설명 창 (text_fn: 보여 줄 글자를 돌려주는 함수)"""
    state = {"win": None, "job": None}

    def show():
        text = text_fn()
        if not text:
            return
        win = tk.Toplevel(widget)
        win.wm_overrideredirect(True)
        x, y = widget.winfo_rootx(), widget.winfo_rooty() + widget.winfo_height() + 4
        win.wm_geometry(f"+{max(0, x - 200)}+{y}")
        tk.Label(win, text=text, bg="#111827", fg="white", padx=8, pady=5, justify="left", wraplength=620,
                 font=(font_family(widget), 9)).pack()
        state["win"] = win

    def enter(_e):
        state["job"] = widget.after(350, show)

    def leave(_e):
        if state["job"]:
            widget.after_cancel(state["job"])
            state["job"] = None
        if state["win"]:
            state["win"].destroy()
            state["win"] = None

    widget.bind("<Enter>", enter, add="+")
    widget.bind("<Leave>", leave, add="+")
