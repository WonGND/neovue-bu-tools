import json
import os
import re
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd
from openpyxl import load_workbook

# 공용 모듈(common/) 경로: 소스 실행 시 ../common, exe는 spec에서 함께 묶음
_HERE = Path(__file__).resolve().parent
for _p in (_HERE.parent / "common", _HERE / "common"):
    if _p.exists() and str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import capability_sheet  # noqa: E402
import history_db  # noqa: E402
import ui_style as ui  # noqa: E402
from capability import is_constant, summary_line  # noqa: E402

VERSION = "v1.6.2"
APP_TITLE = f"NEOVUE_OneClick_정리_{VERSION}"

# [설정] 시간 포맷 및 기본 양식 경로
TIME_FMT = "%Y.%m.%d %H:%M:%S"
TEMPLATE_NAME = "양식_Daily data_BU설비_광학데이터_날짜.xlsx"   # v1.6: exe에 같이 묶고, exe 옆에 두면 그 파일을 우선
DEFAULT_TEMPLATE_PATH = r"../../업무/00. 곽원용/01. PJT/2025/3. 12.9 inch_VDA129BLAC01/0. 양산 대응 폴더/Daily data 정리/양식_Daily data_BU설비_광학데이터_날짜.xlsx"
SUPPORTED_TIME_FORMATS = (
    "%Y.%m.%d %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
)
SETTINGS_NAME = "OneClick_settings.json"
DB_DEFAULT_NAME = "BU_History_DB.xlsx"

# 시트별 선별 컬럼 정의
CA_SELECTED_COLS = [
    "Time", "Panel", "Panel_ID",
    "Red_X", "Red_Y",
    "Green_X", "Green_Y",
    "Blue_X", "Blue_Y",
    "White_SX", "White_SY", "White_LV",
    "Black_LV", "EEPROM_59"
]

LMK_SELECTED_COLS = [
    "Time", "Panel", "Panel_ID", "Judge",
    "Black_Uniformity", "White_Uniformity"
]

TEMPLATE_START_ROW = 20
TEMPLATE_MAPPING = {
    'Red_X': 6, 'Red_Y': 7, 'Green_X': 8, 'Green_Y': 9, 'Blue_X': 10, 'Blue_Y': 11,
    'White_SX': 13, 'White_SY': 14, 'White_LV': 15, 'Black_LV': 16,
    'White_Uniformity': 19, 'Black_Uniformity': 20, 'EEPROM_59': 27
}


def app_base_dir() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else _HERE


# ---------------------------------------------------------------- 데이터 처리 (화면과 분리)
# v1.6.1: EEPROM 값은 16진수 글자('06', '1B')라 숫자로 읽으면 앞 0이 사라짐 → 글자로 읽음
_TEXT_COLS = {f"EEPROM_{i}": str for i in range(100)} | {"EEPROM_Verify": str}


def _strip_text(df):
    for c in df.columns:
        if c in _TEXT_COLS:
            df[c] = df[c].astype("string").str.strip().astype(object).where(df[c].notna(), None)
    return df


def read_df(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        last_error = None
        for encoding in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
            try:
                return _strip_text(pd.read_csv(path, encoding=encoding, dtype=_TEXT_COLS))
            except UnicodeDecodeError as exc:
                last_error = exc
        raise last_error
    return _strip_text(pd.read_excel(path, dtype=_TEXT_COLS))


def clean_df(df):
    df = df.copy()
    # 시간 처리
    if "Time" in df.columns:
        df["Time_dt"] = pd.to_datetime(df["Time"], format=TIME_FMT, errors="coerce")
        for fmt in SUPPORTED_TIME_FORMATS[1:]:
            missing = df["Time_dt"].isna()
            if not missing.any():
                break
            parsed = pd.to_datetime(df.loc[missing, "Time"], format=fmt, errors="coerce")
            df.loc[missing, "Time_dt"] = parsed
        df = df[df["Time_dt"].notna()]
        df = df.sort_values("Time_dt")

    # Panel_ID 기준 최신화
    if "Panel_ID" in df.columns:
        df["Panel_ID"] = df["Panel_ID"].astype(str).str.strip().str.replace(" ", "", regex=False).str.replace("-", "", regex=False).str.replace("_", "", regex=False)
        df["Panel_ID"] = df["Panel_ID"].str.replace(".0", "", regex=False).str.upper()
        df = df.drop_duplicates(subset=["Panel_ID"], keep="last")
        df = df.sort_values("Panel_ID")

    return df


def sniff_kind(path) -> str | None:
    """파일 열 이름으로 설비 종류 추정: 'CA410'(색좌표) / 'LMK6'(BU·WU) / None (v1.6: 칸을 바꿔 넣었는지 바로 알림)"""
    try:
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            cols = None
            for enc in ("utf-8-sig", "cp949"):
                try:
                    cols = pd.read_csv(path, nrows=0, encoding=enc).columns
                    break
                except UnicodeDecodeError:
                    continue
        else:
            cols = pd.read_excel(path, nrows=0).columns
    except Exception:
        return None
    cols = {str(c).strip() for c in (cols if cols is not None else [])}
    if cols & {"White_Uniformity", "Black_Uniformity"}:
        return "LMK6"
    if cols & {"Red_X", "White_SX", "White_LV"}:
        return "CA410"
    return None


# v1.6.1: 색좌표 파일 이름은 호기마다 다름 — 4호기 CA410DataLog, 3호기 LVDataLog
NAME_KEYS = {"CA410": ("ca410", "lvdatalog"), "LMK6": ("lmk6",)}


def find_pair(path, want: str) -> str | None:
    """색좌표(CA410/LV) 파일을 고르면 같은 폴더에서 짝이 되는 LMK6 파일(또는 반대)을 찾음.
    파일 이름의 날짜·접미사가 같은 것을 우선, 없으면 그 폴더에 하나뿐인 것"""
    folder, name = os.path.split(path)
    stem = os.path.splitext(name)[0]
    key = re.sub(r"^(CA410|LMK6|LV)[_\- ]*(DataLog)?[_\- ]*", "", stem, flags=re.I)
    keys = NAME_KEYS.get(want, (want.lower(),))
    try:
        cands = [os.path.join(folder, f) for f in sorted(os.listdir(folder))
                 if f.lower().endswith((".csv", ".xlsx", ".xls")) and any(k in f.lower() for k in keys)]
    except OSError:
        return None
    if key:
        same = [c for c in cands if key.lower() in os.path.basename(c).lower()]
        if len(same) == 1:
            return same[0]
    return cands[0] if len(cands) == 1 else None


def detect_ld(lmk_path) -> bool:
    """LMK6 DataLog의 BU(Black_Uniformity)가 전부 같은 값이면 LD 모델로 판단 (D3 B안)."""
    if not lmk_path or not os.path.exists(lmk_path):
        return False
    try:
        df = read_df(lmk_path)
    except Exception:
        return False
    return "Black_Uniformity" in df.columns and is_constant(df["Black_Uniformity"])


def save_individual_file(df, mode, save_dir):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    out_path = os.path.join(save_dir, f"{mode}_cleaned_{timestamp}.xlsx")
    selected_cols = CA_SELECTED_COLS if mode == "CA410" else LMK_SELECTED_COLS
    actual_cols = [c for c in selected_cols if c in df.columns]   # 실제 데이터에 있는 컬럼만 (에러 방지)
    df_selected = df[actual_cols].copy()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1_All", index=False)
        df_selected.to_excel(writer, sheet_name="Sheet2_Selected", index=False)
    return out_path


def _template_meta(ws):
    cut = ws["R2"].value if ws["R2"].value else ""
    code = ws["R3"].value if ws["R3"].value else ""
    return str(cut).strip(), str(code).strip()


def _cap_frame(df):
    """공정능력 계산용 표: LotID, WU, BU, time(LMK 측정 시각 우선)"""
    t = df["LMK_Time"] if "LMK_Time" in df else df.get("Time_dt")
    if "Time_dt" in df and t is not None:
        t = t.fillna(df["Time_dt"])
    return pd.DataFrame({"LotID": df["Panel_ID"].values,
                         "WU": pd.to_numeric(df.get("White_Uniformity"), errors="coerce").values,
                         "BU": pd.to_numeric(df.get("Black_Uniformity"), errors="coerce").values,
                         "time": pd.to_datetime(t).values})


def _db_rows(df, cut_model, ld, ca_path, lmk_path):
    t = df["LMK_Time"] if "LMK_Time" in df else df.get("Time_dt")
    if "Time_dt" in df and t is not None:
        t = t.fillna(df["Time_dt"])
    num = lambda c: pd.to_numeric(df[c], errors="coerce") if c in df else None
    rows = pd.DataFrame({
        "LotID": df["Panel_ID"].values,
        "모델명": cut_model,
        "측정일시": pd.to_datetime(t).values,
        "채널": df["Panel"].values if "Panel" in df else None,
        "Red_X": num("Red_X"), "Red_Y": num("Red_Y"), "Green_X": num("Green_X"), "Green_Y": num("Green_Y"),
        "Blue_X": num("Blue_X"), "Blue_Y": num("Blue_Y"),
        "White_x": num("White_SX"), "White_y": num("White_SY"),
        "White_LV": num("White_LV"), "Black_LV": num("Black_LV"),
        "WU": num("White_Uniformity"), "BU": num("Black_Uniformity"),
        "판정": df["Judge"].values if "Judge" in df else None,
        "EEPROM_59": df["EEPROM_59"].astype(str).values if "EEPROM_59" in df else None,
    })
    rows["CR"] = pd.to_numeric(rows["White_LV"], errors="coerce") / pd.to_numeric(rows["Black_LV"], errors="coerce")
    rows["모델코드"] = rows["LotID"].map(history_db.model_code)
    rows["생산일"] = rows["LotID"].map(history_db.prod_date)
    rows["LD"] = "Y" if ld else ""
    rows["출처"] = f"Tool A {VERSION}"
    rows["원본파일"] = " / ".join(os.path.basename(p) for p in (ca_path, lmk_path) if p)
    rows["Tool"] = f"OneClick {VERSION}"
    return rows


def process(opts: dict, log=print, progress=None) -> dict:
    """opts: ca_path, lmk_path, template_path, save_dir, make_individual, make_cpk, ld, append_db, db_path
    progress(퍼센트, 단계 글자) — 화면 진행 표시용 (v1.6)"""
    def step(pct, text):
        if progress:
            progress(pct, text)
    ca_path, lmk_path = opts.get("ca_path") or "", opts.get("lmk_path") or ""
    save_dir, temp_path = opts["save_dir"], opts.get("template_path") or ""
    ld = bool(opts.get("ld"))
    result = {"out_path": None, "caps": None, "db": None}
    df_ca = df_lmk = None
    if ca_path:
        step(5, "색좌표 정리")
        log("색좌표(CA410/LV) 데이터 정리 중...")
        raw = read_df(ca_path)
        df_ca = clean_df(raw)
        log(f"   색좌표 {len(raw)}건 → {len(df_ca)}건 (중복 {len(raw) - len(df_ca)}건 제거)")
    if lmk_path:
        step(18, "LMK6 정리")
        log("LMK6 데이터 정리 중...")
        raw = read_df(lmk_path)
        df_lmk = clean_df(raw)
        log(f"   LMK6 {len(raw)}건 → {len(df_lmk)}건 (중복 {len(raw) - len(df_lmk)}건 제거)")

    lmk_file = None
    if opts.get("make_individual"):
        step(30, "개별 정리 파일")
        if df_ca is not None:
            log("CA410 개별 정리 파일(2개 시트) 생성 중...")
            res = save_individual_file(df_ca, "CA410", save_dir)
            if not (ca_path and lmk_path):
                result["out_path"] = res
        if df_lmk is not None:
            log("LMK6 개별 정리 파일(2개 시트) 생성 중...")
            lmk_file = save_individual_file(df_lmk, "LMK6", save_dir)
            if not (ca_path and lmk_path):
                result["out_path"] = lmk_file

    merged, wb, cut_model, code = None, None, "", ""
    if ca_path and lmk_path:
        if not os.path.exists(temp_path):
            raise FileNotFoundError(f"양식 파일을 찾을 수 없습니다: {temp_path}")
        if "Panel_ID" not in df_ca.columns or "Panel_ID" not in df_lmk.columns:
            raise ValueError("CA410/LMK6 양쪽 모두 Panel_ID 컬럼이 필요합니다.")
        step(45, "병합 · 양식 기입")
        log("데이터 병합 중...")
        lmk_subset = df_lmk[[c for c in ["Panel_ID", "Black_Uniformity", "White_Uniformity", "Judge", "Time_dt"]
                             if c in df_lmk.columns]].rename(columns={"Time_dt": "LMK_Time"})
        merged = pd.merge(df_ca, lmk_subset, on="Panel_ID", how="outer").sort_values("Panel_ID")
        log("양식 기입 중...")
        wb = load_workbook(temp_path)
        ws = wb.active
        cut_model, code = _template_meta(ws)
        for i, row in enumerate(merged.itertuples(), start=TEMPLATE_START_ROW):
            ws.cell(row=i, column=2, value=i - TEMPLATE_START_ROW + 1)
            ws.cell(row=i, column=3, value=row.Panel_ID)
            for attr, col in TEMPLATE_MAPPING.items():
                if hasattr(row, attr):
                    ws.cell(row=i, column=col, value=getattr(row, attr))

    # ---- v1.5: 공정능력 시트 (BU/WU, LMK6 데이터가 있을 때)
    base_df = merged if merged is not None else df_lmk
    if opts.get("make_cpk") and base_df is not None and "White_Uniformity" in base_df.columns:
        step(60, "공정능력(Cp/Cpk) 계산")
        log(f"공정능력 계산 중... (LotID별 최신 1건, n = {len(base_df)}){' · LD 모델: BU 제외' if ld else ''}")
        model = " ".join(x for x in (cut_model, code) if x) or history_db.model_code(base_df["Panel_ID"].iloc[0])
        meta = {"model": model, "version": f"Tool A {VERSION}",
                "source": " / ".join(os.path.basename(p) for p in (ca_path, lmk_path) if p)}
        target_wb = wb
        if target_wb is None:
            if lmk_file:
                target_wb = load_workbook(lmk_file)
            else:
                from openpyxl import Workbook
                target_wb = Workbook()
                target_wb.remove(target_wb.active)
        caps = capability_sheet.write_sheet(target_wb, _cap_frame(base_df), ld, meta)
        result["caps"] = caps
        for c in caps.values():
            log("   " + summary_line(c))
            for w in c.warnings:
                if not c.ld:
                    log(f"   ⚠ {c.name}: {w}")
        if wb is None:
            if lmk_file:
                target_wb.save(lmk_file)
            else:
                result["out_path"] = os.path.join(save_dir, f"공정능력_BU_WU_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
                target_wb.save(result["out_path"])

    if wb is not None:
        step(82, "엑셀 저장")
        out_path = os.path.join(save_dir, f"Daily data 통합정리_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        wb.save(out_path)
        result["out_path"] = out_path
        log(f"✅ 통합 완료: {os.path.basename(out_path)}")
        if opts.get("make_cpk") and result["caps"]:
            log(f"   └ 시트 추가: {capability_sheet.SHEET_NAME}")

    # ---- v1.5: 누적 DB 적재 (트렌드 리포트용, 측정 시각 포함)
    if opts.get("append_db") and base_df is not None:
        step(92, "누적 DB 적재")
        try:
            rows = _db_rows(base_df, cut_model, ld, ca_path, lmk_path)
            added, skipped = history_db.append(opts["db_path"], rows)
            result["db"] = (added, skipped)
            log(f"누적 DB 적재: {added}행 추가" + (f", 이미 있는 {skipped}행 건너뜀" if skipped else ""))
        except history_db.DBLockedError as exc:
            log(f"⚠ 누적 DB 적재 실패: {exc}")
            result["db_error"] = str(exc)
    step(100, "완료")
    return result


# ---------------------------------------------------------------- 화면 (v1.6: 노트북 화면 대응 · 경로 기억 · 진행률)
T = ui.T
HELP = [
    ("이 툴이 하는 일", ["색좌표(4호기 CA410DataLog · 3호기 LVDataLog)·LMK6(BU/WU) DataLog를 정리해 Daily data 양식에 채우고,",
                     "공정능력(Cp/Cpk) 시트와 트렌드 리포트용 누적 DB까지 한 번에 만듭니다."]),
    ("순서", ["① 색좌표 칸 '찾기'로 파일 선택 → 같은 폴더에 짝이 되는 LMK6 파일이 있으면 자동으로 채워짐",
            "② 저장 폴더 확인 (처음 한 번 고르면 다음부터 기억)",
            "③ 파란 'One-Click 정리 시작' → 버튼이 채워지며 진행률 표시 → '결과 파일 열기'"]),
    ("칸 설명", ["양식 파일: exe 옆에 양식 엑셀을 두면 그 파일을 쓰고, 없으면 exe에 들어 있는 기본 양식을 씀",
              "LD 모델: BU 값이 전부 같으면 자동 체크 (BU Cp/Cpk 제외)",
              "누적 DB: 문서\\BU_Tools\\BU_History_DB.xlsx — 트렌드 리포트가 같은 파일을 씀"]),
    ("자주 나는 문제", ["'파일이 열려 있어 저장할 수 없습니다' → 결과 엑셀이나 누적 DB를 엑셀에서 닫고 다시 실행",
                   "'LMK6 파일로 보입니다' 안내 → CA410·LMK6 칸이 바뀌지 않았는지 확인",
                   "WU -100(측정 이상값) 행은 Cp/Cpk에서 자동 제외되고 시트 4행에 표시"]),
]


class OneClickApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} {ui.AUTHOR}")
        ui.fit_window(root, 900, 900, 720, 560)
        ui.apply_icon(root)
        self.style, self.font = ui.setup_style(root)

        self.last_saved_file = None
        self.last_saved_folder = None
        self._main_thread = threading.current_thread()
        self.settings_path = app_base_dir() / SETTINGS_NAME
        self.settings = self._load_settings()
        self.template_path = self.find_template()
        self.status = tk.StringVar(value="대기 중")

        self.setup_ui()
        self.log("프로그램이 시작되었습니다.")
        for line in self._db_msg.splitlines():
            self.log(line)
        self._save_settings()
        self.check_template()

    # ---- 설정 저장 (각자 PC) — v1.6: 마지막 폴더·양식·저장 폴더도 기억
    def _load_settings(self):
        s = {"db_path": "", "db_merged": [], "make_cpk": True, "append_db": True,
             "make_individual": True, "last_input_dir": "", "save_dir": "", "template_path": ""}
        try:
            s.update(json.loads(self.settings_path.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            pass
        # v1.5.1: 누적 DB 기본 위치 = 문서\BU_Tools (트렌드 리포트와 같은 파일). exe 옆 옛 DB는 처음 한 번 합침
        self._db_msg = ""
        try:
            s["db_path"], self._db_msg = history_db.adopt_shared_db(
                s.get("db_path"), app_base_dir() / DB_DEFAULT_NAME, s.setdefault("db_merged", []))
        except Exception as exc:
            s["db_path"] = s.get("db_path") or str(history_db.default_db_path())
            self._db_msg = f"⚠ 이전 DB를 합치지 못했습니다: {exc}"
        return s

    def _save_settings(self):
        self.settings.update({"db_path": self.ent_db.get().strip(), "make_cpk": self.var_cpk.get(),
                              "append_db": self.var_db.get(), "make_individual": self.var_make_individual.get(),
                              "save_dir": self.ent_save.get().strip(), "template_path": self.ent_temp.get().strip()})
        try:
            self.settings_path.write_text(json.dumps(self.settings, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _path_row(self, parent, row, label, button_text, command):
        ttk.Label(parent, text=label, style="Cap.TLabel").grid(row=row, column=0, sticky="w", pady=4)
        ent = ttk.Entry(parent)
        ent.grid(row=row, column=1, sticky="ew", padx=(10, 6), pady=4)
        ttk.Button(parent, text=button_text, command=command, width=7).grid(row=row, column=2, pady=4)
        return ent

    def setup_ui(self):
        hdr = ui.header(self.root, APP_TITLE, self.font)
        ui.help_button(hdr, lambda: ui.show_help(self.root, APP_TITLE, HELP, app_base_dir()))
        ttk.Label(hdr, textvariable=self.status, style="HeadStatus.TLabel").pack(side="right", padx=(0, 6))

        scroll = ui.ScrollFrame(self.root)                 # v1.6: 화면이 작으면 스크롤
        scroll.pack(fill="both", expand=True)
        body = ttk.Frame(scroll.body, padding=(14, 8))
        body.pack(fill="both", expand=True)

        # ① 입력 파일
        c1, _ = ui.card(body, "① 입력 파일", "색좌표(CA410·LV)를 고르면 같은 폴더의 LMK6를 자동으로 찾아 채움")
        g = ttk.Frame(c1, style="Card.TFrame")
        g.pack(fill="x", pady=(8, 0))
        self.ent_ca = self._path_row(g, 0, "색좌표 (CA410/LV)", "찾기", lambda: self.browse_file(self.ent_ca, "CA410"))
        self.ent_lmk = self._path_row(g, 1, "LMK6 (BU/WU)", "찾기", lambda: self.browse_file(self.ent_lmk, "LMK6"))
        self.ent_lmk.bind("<FocusOut>", lambda _e: self.update_ld_detection())
        self.ent_temp = self._path_row(g, 2, "양식 파일", "변경", self.browse_template)
        self.ent_temp.insert(0, self.template_path or "")
        self.ent_save = self._path_row(g, 3, "저장 폴더", "찾기", self.browse_folder)
        self.ent_save.insert(0, self.settings.get("save_dir", ""))
        g.columnconfigure(1, weight=1)
        self.lbl_kind = tk.Label(c1, text="", fg=T["warn_fg"], bg=T["warn"], anchor="w", padx=8, pady=3,
                                 font=(self.font, 9, "bold"))
        self.var_make_individual = tk.BooleanVar(value=bool(self.settings.get("make_individual", True)))
        self.chk_individual = ttk.Checkbutton(c1, text="개별 정리 파일(Cleaned) 생성 (2개 시트 구성)",
                                              variable=self.var_make_individual)
        self.chk_individual.pack(anchor="w", pady=(8, 0))

        # ② 분석 옵션
        c2, _ = ui.card(body, "② 분석 옵션")
        opt = ttk.Frame(c2, style="Card.TFrame")
        opt.pack(fill="x", pady=(8, 0))
        self.var_cpk = tk.BooleanVar(value=bool(self.settings.get("make_cpk", True)))
        ttk.Checkbutton(opt, text="공정능력(Cp/Cpk) 시트 생성  —  WU 80~100 · BU 50~100",
                        variable=self.var_cpk).grid(row=0, column=0, columnspan=3, sticky="w", pady=2)
        self.var_ld = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="LD 모델 (BU 미측정)  —  BU Cp/Cpk 제외, 'N/A (LD 모델)' 표기", variable=self.var_ld,
                        command=self._hide_ld_note).grid(row=1, column=0, columnspan=3, sticky="w", pady=2)
        self.lbl_ld = tk.Label(opt, text="", fg=T["warn_fg"], bg=T["warn"], anchor="w", padx=8, pady=3,
                               font=(self.font, 9, "bold"))
        self.var_db = tk.BooleanVar(value=bool(self.settings.get("append_db", True)))
        ttk.Checkbutton(opt, text="누적 DB에 적재 (트렌드 리포트용, 측정 시각 함께 저장)",
                        variable=self.var_db).grid(row=3, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Label(opt, text="DB 위치", style="Cap.TLabel").grid(row=4, column=0, sticky="w", pady=(4, 0))
        self.ent_db = ttk.Entry(opt)
        self.ent_db.insert(0, self.settings.get("db_path", ""))
        self.ent_db.grid(row=4, column=1, sticky="ew", padx=(10, 6), pady=(4, 0))
        ttk.Button(opt, text="변경", command=self.browse_db, width=7).grid(row=4, column=2, pady=(4, 0))
        opt.columnconfigure(1, weight=1)

        # 실행 버튼 (v1.6: 진행률이 버튼에 채워짐)
        self.btn_run = ui.ProgressButton(body, self.start_process, self.font, idle="▶  One-Click 정리 시작",
                                         running="정리 중", done="✓  정리 완료", bg=T["app"])
        self.btn_run.pack(fill="x", pady=(0, 10))

        # ③ 결과 · 처리 로그
        c3, t3 = ui.card(body, "③ 결과 · 처리 로그", fill="both", expand=True)
        self.btn_open_d = ttk.Button(t3, text="폴더 열기", state="disabled", command=self.open_last_folder)
        self.btn_open_d.pack(side="right", padx=(6, 0))
        self.btn_open_f = ttk.Button(t3, text="결과 파일 열기", state="disabled", command=self.open_last_file)
        self.btn_open_f.pack(side="right")
        self.result_var = tk.StringVar(value="결과 파일: 아직 생성되지 않음")
        ttk.Label(c3, textvariable=self.result_var, style="File.TLabel").pack(anchor="w", pady=(8, 6))
        wrap = ttk.Frame(c3, style="Panel.TFrame", padding=1)
        wrap.pack(fill="both", expand=True)
        self.txt_log = scrolledtext.ScrolledText(wrap, height=8, state="disabled", font=("Consolas", 10),
                                                 bg=T["log_bg"], fg=T["log_fg"], insertbackground=T["log_fg"],
                                                 relief="flat", borderwidth=8)
        self.txt_log.pack(fill="both", expand=True)

    # ---- LD 자동 감지 (D3 B안: 데이터 검사로 미리 체크, 사람은 확인)
    def update_ld_detection(self):
        path = self.ent_lmk.get().strip()
        if detect_ld(path):
            self.var_ld.set(True)
            self.lbl_ld.config(text="ⓘ BU 값이 전부 같은 값입니다 (σ = 0) → LD 모델로 자동 선택했습니다. 아니면 체크를 해제하세요.")
            self.lbl_ld.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(2, 4))
            self.log("데이터 검사: BU 값이 모두 같음 → LD 모델로 자동 선택")
        else:
            self._hide_ld_note()

    def _hide_ld_note(self):
        self.lbl_ld.grid_remove()

    def log(self, msg):
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n"
        if threading.current_thread() is self._main_thread:
            self._append_log(line)
        else:
            self.root.after(0, self._append_log, line)

    def _append_log(self, line):
        self.txt_log.config(state="normal")
        self.txt_log.insert(tk.END, line)
        self.txt_log.see(tk.END)
        self.txt_log.config(state="disabled")

    def _set_result_buttons(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self.btn_open_f.config(state=state)
        self.btn_open_d.config(state=state)

    # ---- 양식 파일 (v1.6: 기억한 경로 → exe 옆 → 예전 기본 경로 → exe에 들어 있는 기본 양식)
    def find_template(self):
        base = app_base_dir()
        cands = [self.settings.get("template_path") or "", str(base / TEMPLATE_NAME),
                 os.path.normpath(os.path.join(str(base), DEFAULT_TEMPLATE_PATH)),
                 os.path.normpath(os.path.join(os.path.dirname(str(base)), DEFAULT_TEMPLATE_PATH))]
        if getattr(sys, "frozen", False):
            cands.append(str(Path(getattr(sys, "_MEIPASS", "")) / TEMPLATE_NAME))
        for c in cands:
            if c and os.path.exists(c):
                return c
        return ""

    def check_template(self):
        path = self.ent_temp.get()
        if path and os.path.exists(path):
            inside = getattr(sys, "frozen", False) and str(Path(getattr(sys, "_MEIPASS", ""))) in path
            self.log(f"양식 파일: {os.path.basename(path)}" + (" (exe에 들어 있는 기본 양식)" if inside else ""))
        else:
            self.log("⚠ 양식 파일을 찾지 못했습니다. '변경'으로 양식 엑셀을 골라 주세요 (한 번 고르면 기억).")

    def _initial_dir(self):
        d = self.settings.get("last_input_dir") or ""
        return d if d and os.path.isdir(d) else None

    def _show_kind(self, text):
        if text:
            self.lbl_kind.config(text=text)
            self.lbl_kind.pack(fill="x", pady=(6, 0), before=self.chk_individual)
        else:
            self.lbl_kind.pack_forget()

    def browse_file(self, entry, name):
        p = filedialog.askopenfilename(title=f"{name} 파일 선택", initialdir=self._initial_dir(),
                                       filetypes=[("Data files", "*.csv *.xlsx *.xls")])
        if not p:
            return
        self._set(entry, p)
        self.settings["last_input_dir"] = os.path.dirname(p)
        if not self.ent_save.get():
            self._set(self.ent_save, os.path.dirname(p))
        # v1.6: 칸을 바꿔 넣었는지 바로 확인
        kind = sniff_kind(p)
        notes = []
        if kind and kind != name:
            lbl = {"CA410": "색좌표(CA410/LV)", "LMK6": "LMK6(BU/WU)"}
            notes.append(f"ⓘ 고른 파일은 {lbl[kind]} 파일로 보입니다 ({lbl[name]} 칸). 칸이 바뀌지 않았는지 확인하세요.")
        # v1.6: 짝 파일 자동 채우기
        other, other_entry = ("LMK6", self.ent_lmk) if name == "CA410" else ("CA410", self.ent_ca)
        if not other_entry.get().strip():
            pair = find_pair(p, other)
            if pair and sniff_kind(pair) in (other, None):
                self._set(other_entry, pair)
                self.log(f"같은 폴더에서 {other} 파일을 찾아 채웠습니다: {os.path.basename(pair)}")
        self._show_kind(" ".join(notes))
        self._save_settings()
        self.update_ld_detection()

    @staticmethod
    def _set(entry, value):
        entry.delete(0, tk.END)
        entry.insert(0, value)

    def browse_template(self):
        cur = self.ent_temp.get().strip()
        p = filedialog.askopenfilename(title="양식 파일 선택", initialdir=os.path.dirname(cur) if cur else None,
                                       filetypes=[("Excel files", "*.xlsx")])
        if p:
            self._set(self.ent_temp, p)
            self._save_settings()
            self.log(f"양식 파일 변경: {os.path.basename(p)} (다음 실행에도 기억)")

    def browse_folder(self):
        p = filedialog.askdirectory(title="저장 폴더 선택", initialdir=self.ent_save.get().strip() or None)
        if p:
            self._set(self.ent_save, p)
            self._save_settings()

    def browse_db(self):
        p = filedialog.asksaveasfilename(title="누적 DB 파일 위치", defaultextension=".xlsx",
                                         initialfile=DB_DEFAULT_NAME, filetypes=[("Excel files", "*.xlsx")])
        if p:
            self._set(self.ent_db, p)
            self._save_settings()

    def start_process(self):
        # v1.5: 입력 검사를 화면 스레드에서 먼저 (v1.4.1의 '처리 중...' 버튼 고착 수정)
        ca_path, lmk_path = self.ent_ca.get().strip(), self.ent_lmk.get().strip()
        if not any([ca_path, lmk_path]):
            messagebox.showwarning("알림", "정리할 파일을 하나 이상 선택해 주세요.")
            return
        for label, path in (("CA410", ca_path), ("LMK6", lmk_path)):
            if path and not os.path.exists(path):
                messagebox.showwarning("알림", f"{label} 파일을 찾을 수 없습니다.\n{path}\n\n→ '찾기'로 다시 골라 주세요.")
                return
        save_dir = self.ent_save.get().strip()
        if not save_dir:
            messagebox.showwarning("알림", "저장 폴더를 지정해 주세요.")
            return
        if not os.path.isdir(save_dir):
            messagebox.showwarning("알림", f"저장 폴더가 없습니다.\n{save_dir}\n\n→ '찾기'로 다시 골라 주세요.")
            return
        if ca_path and lmk_path and not os.path.exists(self.ent_temp.get().strip()):
            messagebox.showwarning("알림", "양식 파일을 찾을 수 없습니다.\n→ '변경'으로 양식 엑셀을 골라 주세요.")
            return
        if self.var_cpk.get() and lmk_path and not self.var_ld.get() and detect_ld(lmk_path):
            if messagebox.askyesno("LD 모델 확인", "BU 값이 모두 같습니다 (σ = 0).\nLD 모델(BU 미측정)로 보입니다.\n\nBU Cp/Cpk를 계산에서 제외할까요?"):
                self.var_ld.set(True)
        self._save_settings()
        opts = {"ca_path": ca_path, "lmk_path": lmk_path, "template_path": self.ent_temp.get().strip(),
                "save_dir": save_dir, "make_individual": self.var_make_individual.get(),
                "make_cpk": self.var_cpk.get(), "ld": self.var_ld.get(),
                "append_db": self.var_db.get(), "db_path": self.ent_db.get().strip()}
        self.btn_run.start()
        self.status.set("처리 중")
        threading.Thread(target=self.run, args=(opts,), daemon=True).start()

    def run(self, opts):
        try:
            result = process(opts, self.log, progress=lambda p, t: self.root.after(0, self.btn_run.progress, p, t))
            out_path = result.get("out_path")
            if out_path:
                self.last_saved_file = out_path
                self.last_saved_folder = opts["save_dir"]
                self.root.after(0, self._done, out_path, result.get("db_error"))
            else:
                self.log("⚠ 작업은 완료되었으나 생성된 파일이 없습니다 (옵션 확인).")
                self.root.after(0, self._finish, False, "완료 (생성 파일 없음)")
        except Exception as e:
            self.log(f"❌ 에러: {e}")
            self.root.after(0, self._finish, False, "오류")
            self.root.after(0, messagebox.showerror, "오류", ui.friendly_error(e))

    def _finish(self, ok, status):
        self.btn_run.finish(ok)
        self.status.set(status)

    def _done(self, out_path, db_error=None):
        # v1.6: 완료 팝업 대신 버튼 '✓ 정리 완료' + 결과 파일 이름 표시 (DB 적재 실패만 알림)
        self._set_result_buttons(True)
        self.result_var.set(f"결과 파일: {os.path.basename(out_path)}")
        self._finish(True, "완료")
        if db_error:
            messagebox.showwarning("누적 DB", "정리 파일은 만들었지만 누적 DB에는 넣지 못했습니다.\n\n" + db_error
                                   + "\n\n→ DB 파일을 엑셀에서 닫고 같은 파일로 다시 실행하면 DB에만 추가됩니다 (중복은 자동 제외).")

    def open_last_file(self):
        if self.last_saved_file:
            self._open(self.last_saved_file)

    def open_last_folder(self):
        if self.last_saved_folder:
            self._open(self.last_saved_folder)

    def _open(self, path):
        try:
            ui.open_path(path)
        except OSError as exc:
            messagebox.showerror("열기 실패", ui.friendly_error(exc))


if __name__ == "__main__":
    root = tk.Tk()
    app = OneClickApp(root)
    root.mainloop()
