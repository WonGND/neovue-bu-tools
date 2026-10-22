"""BU 트렌드 리포트 — 과거 파일 읽기 + 리포트 생성 (화면과 분리).

그래프는 `CHARTS` 모듈을 통해서만 그린다. 엑셀 기본 차트(B안)로 바꿀 때는
같은 함수 이름을 가진 모듈로 CHARTS만 교체하면 된다.
"""
from __future__ import annotations

import io
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

_HERE = Path(__file__).resolve().parent
for _p in (_HERE.parent / "common", _HERE / "common"):
    if _p.exists() and str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import charts as CHARTS  # noqa: E402
import history_db as db  # noqa: E402
from capability import SPECS, capability, grade, is_constant  # noqa: E402

VERSION = "v0.1"
FONT = "맑은 고딕"
DEFAULT_SETTINGS = {
    "bu_sigma_increase_pct": 20.0,      # 직전 이벤트 대비 BU σ 증가 경고(%)
    "wu_mean_shift": 1.0,               # 평균 이동 경고(%p)
    "bu_mean_shift": 3.0,
    "ng_rate_increase": 2.0,            # NG율 증가 경고(%p)
    "daily_mean_dev_sigma": 0.5,        # 이벤트 안 일자 평균이 이벤트 평균에서 몇 σ 벗어나면 표시
    "white_typ": {"WINE": [0.293, 0.320], "VDA129BLAC01": [0.293, 0.320], "VMA129BLAC01": [0.293, 0.320]},
    "gcs": 0.005,
}
TEMPLATE_COLS = {"LotID": 3, "Red_X": 6, "Red_Y": 7, "Green_X": 8, "Green_Y": 9, "Blue_X": 10, "Blue_Y": 11,
                 "White_x": 13, "White_y": 14, "White_LV": 15, "Black_LV": 16, "WU": 19, "BU": 20, "EEPROM_59": 27}
_thin = Side(style="thin", color="CBD5E1")
_BOX = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_HEAD = PatternFill("solid", fgColor="1E3A5F")


# ------------------------------------------------------------------ 과거 Daily data 파일 읽기
def date_from_filename(path) -> date | None:
    m = re.search(r"(20\d{6})[_-]?\d{0,6}", Path(path).stem)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y%m%d").date()
        except ValueError:
            return None
    return None


def parse_daily_file(path) -> dict:
    """Tool A 'Daily data 통합정리' 결과 파일 → {'rows': DataFrame, 'model': .., 'code': .., 'date': ..}"""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    cut = ws["R2"].value or ""
    code = ws["R3"].value or ""
    recs = []
    for row in ws.iter_rows(min_row=20, max_col=27, values_only=True):
        lot = row[2]
        if lot is None or str(lot).strip() == "":
            if recs:
                break
            continue
        recs.append({k: row[c - 1] for k, c in TEMPLATE_COLS.items()})
    wb.close()
    df = pd.DataFrame(recs, columns=list(TEMPLATE_COLS))
    for c in df.columns:
        if c not in ("LotID", "EEPROM_59"):
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df["CR"] = df["White_LV"] / df["Black_LV"]
    d = date_from_filename(path) or datetime.fromtimestamp(os.path.getmtime(path)).date()
    return {"rows": df, "model": str(cut).strip(), "code": str(code).strip(), "date": d,
            "ld": bool(len(df) > 1 and is_constant(df["BU"]))}


def rows_for_db(parsed: dict, meas_date: date, path, manual_event: str = "") -> pd.DataFrame:
    df = parsed["rows"].copy()
    df["모델명"] = parsed["model"]
    df["모델코드"] = df["LotID"].map(db.model_code).replace("", parsed["code"])
    df["측정일시"] = pd.Timestamp(meas_date)
    df["생산일"] = df["LotID"].map(db.prod_date)
    df["판정"] = np.where((df["WU"] < SPECS["WU"][0]) | ((df["BU"] < SPECS["BU"][0]) & (not parsed["ld"])), "NG", "OK")
    df["LD"] = "Y" if parsed["ld"] else ""
    df["이벤트(수동)"] = manual_event or ""
    df["출처"] = "Tool C 업로드(과거 파일)"
    df["원본파일"] = Path(path).name
    df["Tool"] = f"BU_TREND {VERSION}"
    df["EEPROM_59"] = df["EEPROM_59"].astype(str)
    return df


# ------------------------------------------------------------------ 엑셀 도우미
def _title(ws, cell, text, sub=None):
    ws[cell] = text
    ws[cell].font = Font(name=FONT, size=16, bold=True, color="1E3A5F")
    if sub:
        c = ws[cell]
        ws.cell(c.row + 1, c.column, sub).font = Font(name=FONT, size=9, color="6B7280")


def _table(ws, r0, c0, header, rows, fills=None, widths=None):
    for j, h in enumerate(header):
        c = ws.cell(r0, c0 + j, h)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = _HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _BOX
    ws.row_dimensions[r0].height = 30
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            c = ws.cell(r0 + i, c0 + j, v)
            c.font = Font(name=FONT, size=9, bold=(j == 0))
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = _BOX
            if fills and (i - 1, j) in fills:
                c.fill = PatternFill("solid", fgColor=fills[(i - 1, j)])
    return r0 + len(rows) + 1


def _img(ws, png: bytes, anchor: str, width_px: int) -> int:
    img = XLImage(io.BytesIO(png))
    ratio = width_px / img.width
    img.width, img.height = width_px, int(img.height * ratio)
    ws.add_image(img, anchor)
    return int(img.height * 0.75 / 15) + 1


def _r(v, nd=2):
    return round(float(v), nd) if v is not None and np.isfinite(v) else "N/A"


# ------------------------------------------------------------------ 리포트
def _event_stats(name, g, ev_row):
    g = g.sort_values("측정일시")
    ld = (g["LD"].astype(str) == "Y").any() or (len(g) > 1 and is_constant(g["BU"]))
    cw = capability(g["WU"], "WU")
    cb = capability(g["BU"], "BU", ld=ld)
    ng = ((g["WU"] < SPECS["WU"][0]) | ((g["BU"] < SPECS["BU"][0]) & (not ld))).mean() * 100
    period = ""
    if ev_row is not None:
        period = f"{ev_row['시작일']} ~ {ev_row['종료일']}"
    return {"name": name, "g": g, "ld": ld, "cw": cw, "cb": cb, "ng": ng, "period": period,
            "model": (ev_row["모델"] if ev_row is not None and ev_row["모델"] else ", ".join(sorted(set(g["모델명"].dropna().astype(str))))),
            "memo": ev_row["메모"] if ev_row is not None else "",
            "lv": g["White_LV"].mean(), "lv_sd": g["White_LV"].std(), "cr": g["CR"].mean(), "cr_sd": g["CR"].std(),
            "wx": g["White_x"].mean(), "wy": g["White_y"].mean()}


def generate(db_path, selected_events: list[str], out_path, metrics=("BU·WU", "White 휘도·CR", "White x/y"),
             settings: dict | None = None, log=print) -> str:
    st = {**DEFAULT_SETTINGS, **(settings or {})}
    data, events = db.load(db_path)
    if data.empty:
        raise ValueError("누적 DB에 데이터가 없습니다.")
    data = db.assign_events(data, events)
    ev_order = (events.sort_values("시작일")["이벤트명"].astype(str).tolist())
    selected = [e for e in ev_order if e in selected_events] + [e for e in selected_events if e not in ev_order]
    data = data[data["이벤트"].isin(selected)].copy()
    if data.empty:
        raise ValueError("선택한 이벤트에 해당하는 데이터가 없습니다. 이벤트 기간을 확인하세요.")
    for c in ("WU", "BU", "White_LV", "Black_LV", "CR", "White_x", "White_y"):
        data[c] = pd.to_numeric(data[c], errors="coerce")
    # 같은 이벤트 안 동일 LotID는 최신 측정 1건
    data = data.sort_values("측정일시").drop_duplicates(subset=["이벤트", "LotID"], keep="last")
    ev_rows = {str(r["이벤트명"]): r for _, r in events.iterrows()}
    stats = [_event_stats(e, data[data["이벤트"] == e], ev_rows.get(e)) for e in selected if (data["이벤트"] == e).any()]
    labels = [s["name"] for s in stats]
    log(f"리포트 대상: {', '.join(labels)} (총 {len(data)}행)")

    wb = Workbook()
    # ---- 1. 이벤트 요약 + 변화 감지
    ws = wb.active
    ws.title = "이벤트 요약"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJKLMNOPQRS", (2, 14, 22, 12, 7, 9, 8, 8, 9, 8, 8, 8, 10, 9, 9, 9, 8, 24, 2)):
        ws.column_dimensions[col].width = w
    _title(ws, "B2", "트렌드 리포트 — 이벤트 요약",
           f"BU_TREND {VERSION} · {datetime.now():%Y-%m-%d %H:%M} · 이벤트 안 동일 LotID는 최신 1건 · Cp/Cpk 같은 기준 색상")
    hdr = ["이벤트", "기간", "모델", "n", "WU 평균", "WU Cp", "WU Cpk", "BU 평균", "BU σ", "BU Cp", "BU Cpk",
           "W 휘도 평균", "CR 평균", "W x", "W y", "NG율", "메모"]
    rows, fills = [], {}
    for i, s in enumerate(stats):
        cw, cb = s["cw"], s["cb"]
        rows.append([s["name"], s["period"], s["model"], len(s["g"]), _r(cw.mean), _r(cw.cp), _r(cw.cpk),
                     "N/A" if s["ld"] else _r(cb.mean), "N/A" if s["ld"] else _r(cb.sigma_overall),
                     "N/A (LD)" if s["ld"] else _r(cb.cp), "N/A (LD)" if s["ld"] else _r(cb.cpk),
                     _r(s["lv"], 1), _r(s["cr"], 0), _r(s["wx"], 4), _r(s["wy"], 4), f"{s['ng']:.1f}%", s["memo"]])
        for j, v in ((5, cw.cp), (6, cw.cpk)) + (() if s["ld"] else ((9, cb.cp), (10, cb.cpk))):
            fills[(i, j)] = grade(v)[1]
    r = _table(ws, 5, 2, hdr, rows, fills)
    ws.auto_filter.ref = f"B5:R{r - 1}"
    r += 1
    ws.cell(r, 2, "변화 감지 (직전 이벤트 대비 · 기준값은 설정 파일)").font = Font(name=FONT, size=11, bold=True, color="1E3A5F")
    flags, ff = [], {}
    order = ["부적합", "주의", "적합", "양호"]
    for prev, cur in zip(stats, stats[1:]):
        checks = []
        if str(prev["model"]).strip().upper() != str(cur["model"]).strip().upper():
            flags.append([cur["name"], "전체", f"{prev['model']} → {cur['model']}", "", "모델이 달라 비교 생략"])
            continue
        if not (prev["ld"] or cur["ld"]):
            d_sig = (cur["cb"].sigma_overall / prev["cb"].sigma_overall - 1) * 100
            checks.append(("BU σ", f"{prev['cb'].sigma_overall:.2f} → {cur['cb'].sigma_overall:.2f}", f"{d_sig:+.0f}%",
                           d_sig >= st["bu_sigma_increase_pct"], f"σ {st['bu_sigma_increase_pct']:.0f}% 이상 증가"))
            dm = cur["cb"].mean - prev["cb"].mean
            checks.append(("BU 평균", f"{prev['cb'].mean:.2f} → {cur['cb'].mean:.2f}", f"{dm:+.2f}",
                           abs(dm) >= st["bu_mean_shift"], "평균 이동"))
            g0, g1 = grade(prev["cb"].cpk)[0], grade(cur["cb"].cpk)[0]
            checks.append(("BU Cpk 등급", f"{g0} → {g1}", "", order.index(g1) < order.index(g0), "등급 하락"))
        dm = cur["cw"].mean - prev["cw"].mean
        checks.append(("WU 평균", f"{prev['cw'].mean:.2f} → {cur['cw'].mean:.2f}", f"{dm:+.2f}",
                       abs(dm) >= st["wu_mean_shift"], "평균 이동"))
        g0, g1 = grade(prev["cw"].cpk)[0], grade(cur["cw"].cpk)[0]
        checks.append(("WU Cpk 등급", f"{g0} → {g1}", "", order.index(g1) < order.index(g0), "등급 하락"))
        dng = cur["ng"] - prev["ng"]
        checks.append(("NG율", f"{prev['ng']:.1f}% → {cur['ng']:.1f}%", f"{dng:+.1f}%p",
                       dng >= st["ng_rate_increase"], "NG율 증가"))
        for item, chg, amt, hit, why in checks:
            flags.append([cur["name"], item, chg, amt, f"⚠ {why}" if hit else "정상"])
            if hit:
                ff[(len(flags) - 1, 4)] = "FEE2E2"
    if flags:
        _table(ws, r + 1, 2, ["이벤트", "항목", "변화", "변화량", "판정"], flags, ff)
    else:
        ws.cell(r + 1, 2, "비교할 이벤트가 2개 이상일 때 표시됩니다.").font = Font(name=FONT, size=9, color="6B7280")

    # ---- 2. 트렌드(거시)
    ws2 = wb.create_sheet("트렌드(거시)")
    ws2.sheet_view.showGridLines = False
    _title(ws2, "B2", "트렌드 그래프 — 이벤트 간 비교 (거시)", "박스플롯: 상자=사분위, 선=1.5×IQR, 점=이상치 · 최신 이벤트 파란색")
    row = 5
    if "BU·WU" in metrics:
        bu_groups = {s["name"]: s["g"]["BU"].dropna() for s in stats if not s["ld"]}
        if bu_groups:
            _img(ws2, CHARTS.boxplot_groups(bu_groups, "BU", SPECS["BU"][0], "BU 분포 (이벤트별)"), f"B{row}", 600)
        _img(ws2, CHARTS.boxplot_groups({s["name"]: s["g"]["WU"].dropna() for s in stats}, "WU", SPECS["WU"][0],
                                        "WU 분포 (이벤트별)"), f"L{row}", 600)
        row += 24
        series = {"WU Cp": [s["cw"].cp for s in stats], "WU Cpk": [s["cw"].cpk for s in stats],
                  "BU Cp": [np.nan if s["ld"] else s["cb"].cp for s in stats],
                  "BU Cpk": [np.nan if s["ld"] else s["cb"].cpk for s in stats]}
        _img(ws2, CHARTS.cp_cpk_bars(labels, series, "Cp · Cpk 추이 (이벤트별)"), f"B{row}", 600)
        _img(ws2, CHARTS.ng_rate_line(labels, [s["ng"] for s in stats], "NG율 추이 (WU<80 또는 BU<50)"), f"L{row}", 600)
        row += 24
    if "White x/y" in metrics:
        models = [str(s["model"]) for s in stats]
        model = max(set(models), key=models.count) if models else ""
        typ = None
        for key, v in st["white_typ"].items():
            if key and key.upper() in model.upper():
                typ = v
        title = "White 색좌표 평균 이동 (GCS 원 안이면 OK)" if typ else f"White 색좌표 평균 이동 (모델 '{model}' Typ 미설정 → 원 생략)"
        _img(ws2, CHARTS.white_xy(labels, [s["wx"] for s in stats], [s["wy"] for s in stats],
                                  tuple(typ) if typ else None, st["gcs"], title), f"B{row}", 480)
    if "White 휘도·CR" in metrics:
        _img(ws2, CHARTS.mean_band(labels, [s["lv"] for s in stats], [s["lv_sd"] for s in stats],
                                   "White 휘도 평균 ± 3σ", "cd/m²"), f"L{row}", 600)
        row += 22
        _img(ws2, CHARTS.mean_band(labels, [s["cr"] for s in stats], [s["cr_sd"] for s in stats],
                                   "CR 평균 ± 3σ", "CR"), f"L{row}", 600)

    # ---- 3. 측정일별(미시)
    ws3 = wb.create_sheet("측정일별(미시)")
    ws3.sheet_view.showGridLines = False
    _title(ws3, "B2", "측정일별 비교 — 이벤트 안 (미시)",
           f"일자 평균이 이벤트 평균에서 {st['daily_mean_dev_sigma']}σ 이상 벗어나면 표시 · 당일 컨디션 편차 확인용")
    row = 5
    for s in stats:
        g = s["g"].copy()
        g["d"] = pd.to_datetime(g["측정일"]).dt.strftime("%m-%d")
        ws3.cell(row, 2, f"■ {s['name']}  ({s['period']})" + (f"  메모: {s['memo']}" if s["memo"] else "")).font = \
            Font(name=FONT, size=11, bold=True, color="1E3A5F")
        row += 1
        days = sorted(g["d"].unique())
        if not s["ld"]:
            _img(ws3, CHARTS.boxplot_groups({d: g.loc[g.d == d, "BU"].dropna() for d in days}, "BU", SPECS["BU"][0],
                                            "BU 측정일별", width=5.0, highlight_last=False), f"B{row}", 430)
        _img(ws3, CHARTS.boxplot_groups({d: g.loc[g.d == d, "WU"].dropna() for d in days}, "WU", SPECS["WU"][0],
                                        "WU 측정일별", width=5.0, highlight_last=False), f"I{row}", 430)
        daily = []
        for d in days:
            gd = g[g.d == d].sort_values("측정일시")
            daily.append((d, len(gd), capability(gd["WU"], "WU"), capability(gd["BU"], "BU", ld=s["ld"])))
        _img(ws3, CHARTS.daily_cpk_line(days, [x[2].cpk for x in daily], [np.nan if s["ld"] else x[3].cpk for x in daily],
                                        s["memo"], "측정일별 Cp · Cpk", wu_cp=[x[2].cp for x in daily],
                                        bu_cp=None if s["ld"] else [x[3].cp for x in daily]), f"P{row}", 460)
        row += 18
        k = st["daily_mean_dev_sigma"]
        trows, tf = [], {}
        for i, (d, n, cw, cb) in enumerate(daily):
            dev_w = abs(cw.mean - s["cw"].mean) >= k * s["cw"].sigma_overall
            dev_b = (not s["ld"]) and abs(cb.mean - s["cb"].mean) >= k * s["cb"].sigma_overall
            trows.append([d, n, _r(cw.mean), _r(cw.cp), _r(cw.cpk), "N/A" if s["ld"] else _r(cb.mean),
                          "N/A" if s["ld"] else _r(cb.cp), "N/A" if s["ld"] else _r(cb.cpk),
                          "⚠ 일자 평균 이탈" if (dev_w or dev_b) else "정상"])
            for j, v in ((3, cw.cp), (4, cw.cpk)) + (() if s["ld"] else ((6, cb.cp), (7, cb.cpk))):
                tf[(i, j)] = grade(v)[1]
            if dev_w or dev_b:
                tf[(i, 8)] = "FEF3C7"
        row = _table(ws3, row, 2, ["측정일", "n", "WU 평균", "WU Cp", "WU Cpk", "BU 평균", "BU Cp", "BU Cpk", "판정"], trows, tf) + 2

    # ---- 4. NG 이력
    ws4 = wb.create_sheet("NG 이력")
    ng_mask = (data["WU"] < SPECS["WU"][0]) | ((data["BU"] < SPECS["BU"][0]) & (data["LD"].astype(str) != "Y")) | \
              (data["판정"].astype(str).str.upper() == "NG")
    ng = data[ng_mask][["이벤트", "측정일", "LotID", "WU", "BU", "판정", "원본파일"]].sort_values(["이벤트", "측정일"])
    ws4.append(list(ng.columns))
    for rec in ng.itertuples(index=False):
        ws4.append([str(v) if isinstance(v, date) else v for v in rec])
    ws4.auto_filter.ref = f"A1:G{len(ng) + 1}"
    for col, w in zip("ABCDEFG", (14, 12, 48, 9, 9, 8, 46)):
        ws4.column_dimensions[col].width = w

    # ---- 5. Raw
    ws5 = wb.create_sheet("Raw")
    cols = ["이벤트", "매핑상태"] + [c for c in db.DATA_COLS if c in data.columns]
    ws5.append(cols)
    for rec in data[cols].itertuples(index=False):
        ws5.append([None if (isinstance(v, float) and np.isnan(v)) else (str(v) if isinstance(v, date) else v) for v in rec])
    ws5.auto_filter.ref = f"A1:{ws5.cell(1, len(cols)).column_letter}{len(data) + 1}"

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    log(f"✅ 리포트 저장: {Path(out_path).name}")
    return str(out_path)
