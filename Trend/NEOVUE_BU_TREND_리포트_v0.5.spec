# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import _tkinter
import sysconfig
from PyInstaller.utils.hooks import collect_data_files

import importlib.util

# 실행 위치(현재 폴더)와 상관없이 spec 파일이 있는 폴더를 기준으로 경로를 잡는다
project_dir = Path(SPECPATH).resolve()
common_dir = project_dir.parent / "common"   # v1.5: 공용 모듈(capability, charts, history_db)

# 빌드 전 점검: 빠진 것이 있으면 exe를 만들지 않고 바로 이유를 알려준다
_missing_files = [m + ".py" for m in ('capability', 'charts', 'history_db', 'ui_style') if not (common_dir / (m + ".py")).exists()]
if _missing_files:
    raise SystemExit("[빌드 중단] common 폴더에 공용 모듈이 없습니다: " + ", ".join(_missing_files)
                     + "\n  압축을 푼 neovue-bu-tools 폴더 구조(OneClick·Trend·common)를 그대로 두세요.")
_missing_pkgs = [m for m in ("pandas", "openpyxl", "numpy", "matplotlib", "tkinterdnd2")   # v0.2: 파일 끌어다 놓기
                 if importlib.util.find_spec(m) is None]
if _missing_pkgs:
    raise SystemExit("[빌드 중단] 이 파이썬에 설치되지 않은 패키지: " + ", ".join(_missing_pkgs)
                     + "\n  설치: python -m pip install " + " ".join(_missing_pkgs)
                     + "\n  빌드: python -m PyInstaller --noconfirm " + Path(SPEC).name)
python_base = Path(sysconfig.get_config_var("base"))
runtime_base = Path(_tkinter.__file__).resolve().parents[1]
tcl_root = python_base / "tcl"
dll_root = python_base / "DLLs"

if not (tcl_root / "tcl8.6").exists():
    tcl_root = runtime_base / "tcl"
if not (dll_root / "tcl86t.dll").exists():
    dll_root = runtime_base / "DLLs"

binaries = []
for dll_name in ("tcl86t.dll", "tk86t.dll"):
    dll_path = dll_root / dll_name
    if dll_path.exists():
        binaries.append((str(dll_path), "."))

icon_file = project_dir.parent / "BU_LotID" / "neovue_bu_data.ico"
datas = [(str(icon_file), ".")] if icon_file.exists() else []
datas.extend(collect_data_files("tkinter"))
# tkdnd 라이브러리: 빌드하는 PC 것만 (예: win-x64). PyInstaller 기본 hook이 없을 때를 대비해 직접 넣음
import platform as _pf
import sys as _sys
_os = "win" if _sys.platform.startswith("win") else ("osx" if _sys.platform == "darwin" else "linux")
_arch = {"amd64": "x64", "x86_64": "x64", "arm64": "arm64", "aarch64": "arm64", "x86": "x86", "i386": "x86",
         "i686": "x86"}.get(_pf.machine().lower(), "x64")
datas.extend(d for d in collect_data_files("tkinterdnd2") if Path(d[1]).name in (f"{_os}-{_arch}", "tkdnd"))

for source, target in (
    (tcl_root / "tcl8.6", "_tcl_data"),
    (tcl_root / "tk8.6", "_tk_data"),
):
    if source.exists():
        datas.append((str(source), target))

a = Analysis(
    ["BU_trend_report.py"],
    pathex=[str(project_dir), str(common_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=["tkinter", "_tkinter", "pandas", "openpyxl", "numpy",
                   "matplotlib", "matplotlib.backends.backend_agg",
                   "capability", "charts", "history_db", "trend_report", "tkinterdnd2", "PIL.Image", "PIL.ImageDraw", "ui_style"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib.backends.backend_qtagg", "PyQt5", "PySide6", "scipy"],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="NEOVUE_BU_TREND_리포트_v0.5",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(icon_file)] if icon_file.exists() else None,
)
