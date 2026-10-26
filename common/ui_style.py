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
