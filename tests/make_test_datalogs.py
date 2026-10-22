"""시험용 CA410 / LMK6 DataLog 생성기.

첨부 Daily data(WINE 775ea)의 실측값을 되돌려 DataLog 형식으로 만든다.
- 재측정 중복 109건을 넣는다(더 이른 시각, 값 약간 다름) → Tool이 최신 1건만 남기는지 확인
- 측정일은 생산일(LotID 날짜) + 1일로 가정한 가상 값
- 과거 이벤트 B·C용 가상 데이터와 LD 모델(G10, BU 전부 50) 데이터도 만든다
사용: python make_test_datalogs.py <Daily data xlsx> <출력 폴더>
"""
import sys
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

SRC, OUT = Path(sys.argv[1]), Path(sys.argv[2])
OUT.mkdir(parents=True, exist_ok=True)
rng = np.random.default_rng(42)

ws = openpyxl.load_workbook(SRC, data_only=True).worksheets[0]
cols = {"Panel_ID": 3, "Red_X": 6, "Red_Y": 7, "Green_X": 8, "Green_Y": 9, "Blue_X": 10, "Blue_Y": 11,
        "White_SX": 13, "White_SY": 14, "White_LV": 15, "Black_LV": 16, "White_Uniformity": 19,
        "Black_Uniformity": 20, "EEPROM_59": 27}
base = pd.DataFrame([{k: ws.cell(r, c).value for k, c in cols.items()} for r in range(20, 795)])
base["prod"] = base.Panel_ID.str[32:38]
base["ser"] = base.Panel_ID.str[-5:].astype(int)
base = base.sort_values(["prod", "ser"]).reset_index(drop=True)
CA_COLS = ["Time", "Panel", "Panel_ID", "Red_X", "Red_Y", "Green_X", "Green_Y", "Blue_X", "Blue_Y",
           "White_SX", "White_SY", "White_LV", "Black_LV", "EEPROM_59"]
LMK_COLS = ["Time", "Panel", "Panel_ID", "Judge", "Black_Uniformity", "White_Uniformity"]


def write_day(df, day: datetime, tag: str, n_dup: int, ld=False):
    """하루치 DataLog 2종 저장. df는 그날 측정한 패널(최종값)."""
    df = df.copy().reset_index(drop=True)
    t0 = day.replace(hour=8, minute=30)
    df["Time"] = [t0 + timedelta(seconds=40 * i) for i in range(len(df))]
    df["Panel"] = np.where(np.arange(len(df)) % 2 == 0, "L", "R")
    dup = df.sample(n=min(n_dup, len(df)), random_state=int(rng.integers(1e6))).copy()
    dup["Time"] = dup["Time"] - timedelta(minutes=20)          # 재측정 전 값(더 이른 시각)
    for c in ("White_LV", "Black_LV"):
        dup[c] = dup[c] * rng.normal(1, 0.01, len(dup))
    if not ld:
        dup["Black_Uniformity"] = (dup["Black_Uniformity"] - rng.uniform(3, 12, len(dup))).round(3)
    all_rows = pd.concat([df, dup]).sort_values("Time")
    all_rows["Judge"] = np.where((all_rows.White_Uniformity >= 80) & (all_rows.Black_Uniformity >= 50), "OK", "NG")
    ca = all_rows[CA_COLS].copy()
    ca["Time"] = ca["Time"].dt.strftime("%Y.%m.%d %H:%M:%S")
    lmk = all_rows[LMK_COLS].copy()
    lmk["Time"] = lmk["Time"].dt.strftime("%Y-%m-%d %H:%M:%S")
    ca.to_csv(OUT / f"CA410_DataLog_{tag}.csv", index=False, encoding="utf-8-sig")
    lmk.to_csv(OUT / f"LMK6DataLog_{tag}.csv", index=False, encoding="cp949")
    return len(all_rows), len(df)


def shifted(df, n, wu, bu, wx, wy, lv):
    s = df.sample(n=n, replace=True, random_state=int(rng.integers(1e6))).reset_index(drop=True)
    for col, (m, k) in {"White_Uniformity": wu, "Black_Uniformity": bu}.items():
        s[col] = ((s[col] - df[col].mean()) * k + m).clip(upper=99.9).round(3)
    s["White_SX"] = (s["White_SX"] - df["White_SX"].mean() + wx).round(4)
    s["White_SY"] = (s["White_SY"] - df["White_SY"].mean() + wy).round(4)
    s["White_LV"] = (s["White_LV"] - df["White_LV"].mean() + lv).round(4)
    s["Panel_ID"] = [p[:-5] + f"{90000 + i:05d}" for i, p in enumerate(s.Panel_ID)]   # 다른 Lot
    return s


log = []
# 이벤트 A: 실데이터 3일치 (생산일 +1일에 측정했다고 가정)
for prod, n_dup in (("260415", 8), ("260417", 25), ("260420", 76)):
    day = datetime.strptime("20" + prod, "%Y%m%d") + timedelta(days=1)
    d = base[base["prod"] == prod]
    log.append((f"A_{day:%m%d}",) + write_day(d, day, f"A_{day:%m%d}", n_dup))
# 이벤트 B·C: 가상(과거 파일 업로드 시험용)
for tag, day, args in (("B_0610", datetime(2026, 6, 10), (640, (86.2, 0.85), (66.0, 0.80), 0.2941, 0.3197, 915.0)),
                       ("C_0826", datetime(2026, 8, 26), (812, (87.0, 0.70), (64.0, 1.15), 0.2947, 0.3182, 918.5))):
    log.append((tag,) + write_day(shifted(base, *args), day, tag, 20))
# LD 모델(G10): BU 전부 50
g = base.sample(n=30, random_state=3).copy()
g["Panel_ID"] = [f"GJ961110001032601120VDA099AYAC01261002A{i:05d}" for i in range(30)]
g["Black_Uniformity"] = 50.0
g["White_Uniformity"] = rng.normal(86.0, 1.5, 30).round(3)
log.append(("G10_LD",) + write_day(g, datetime(2026, 10, 2), "G10_LD", 3, ld=True))
for row in log:
    print("%-8s 입력 %4d행 → 고유 %4d" % row)
