"""make_test_datalogs.py 입력용 가상 Daily data(xlsx) 생성기 — 실측 파일이 없을 때 사용.
실측 파일과 같은 자리(20~794행, C·F~K·M~P·S·T·AA열)에 가상 값을 넣는다.
생산일별 개수(260415 57 · 260417 153 · 260420 565)는 run_tool_a.py의 기대 행 수(805)와 맞춘다.
사용: python make_fake_daily.py <출력 xlsx>
"""
import sys
from pathlib import Path

import numpy as np
from openpyxl import Workbook

OUT = Path(sys.argv[1])
OUT.parent.mkdir(parents=True, exist_ok=True)
rng = np.random.default_rng(7)

wb = Workbook()
ws = wb.active
ws.title = "Daily_BU설비데이터"
r = 20
for prod, n in (("260415", 57), ("260417", 153), ("260420", 565)):
    for i in range(n):
        ws.cell(r, 3, f"GJ961110001032601120VDA129BLAC01{prod}A{i + 1:05d}")
        for c, (m, s) in zip(range(6, 12), ((0.640, 0.004), (0.330, 0.004), (0.300, 0.004),
                                            (0.600, 0.004), (0.150, 0.003), (0.060, 0.003))):
            ws.cell(r, c, round(float(rng.normal(m, s)), 4))
        ws.cell(r, 13, round(float(rng.normal(0.2945, 0.002)), 4))
        ws.cell(r, 14, round(float(rng.normal(0.3190, 0.002)), 4))
        ws.cell(r, 15, round(float(rng.normal(916.0, 12.0)), 2))
        ws.cell(r, 16, round(float(rng.normal(0.45, 0.03)), 4))
        ws.cell(r, 19, round(float(min(99.9, rng.normal(86.5, 1.6))), 3))
        ws.cell(r, 20, round(float(rng.normal(65.0, 4.5)), 3))
        ws.cell(r, 27, f"{int(rng.integers(0, 256)):02X}")
        r += 1
wb.save(OUT)
print(f"가상 Daily data {r - 20}행 → {OUT}")
