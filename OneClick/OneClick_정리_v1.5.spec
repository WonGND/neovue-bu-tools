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
_missing_files = [m + ".py" for m in ('capability', 'capability_sheet', 'charts', 'history_db') if not (common_dir / (m + ".py")).exists()]
if _missing_files:
    raise SystemExit("[빌드 중단] common 폴더에 공용 모듈이 없습니다: " + ", ".join(_missing_files)
                     + "\n  압축을 푼 tovis-bu-tools 폴더 구조(OneClick·Trend·common)를 그대로 두세요.")
_missing_pkgs = [m for m in ("pandas", "openpyxl", "numpy", "matplotlib") if importlib.util.find_spec(m) is None]
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

datas = []
datas.extend(collect_data_files("tkinter"))

for source, target in (
    (tcl_root / "tcl8.6", "_tcl_data"),
    (tcl_root / "tk8.6", "_tk_data"),
):
    if source.exists():
        datas.append((str(source), target))

a = Analysis(
    ["One_click_정리.py"],
    pathex=[str(project_dir), str(common_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=["tkinter", "_tkinter", "pandas", "openpyxl", "numpy",
                   "matplotlib", "matplotlib.backends.backend_agg",
                   "capability", "capability_sheet", "charts", "history_db"],
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
    name="OneClick_정리_v1.5",
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
)
