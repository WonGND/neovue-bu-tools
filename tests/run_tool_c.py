"""Tool C(트렌드 리포트) 로직 시험: 과거 파일 업로드 → 이벤트 등록 → 리포트 생성.
사용: python run_tool_c.py <Tool A 시험 출력 폴더(DB 포함)> <리포트 출력 폴더>"""
import glob
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "Trend"))
sys.path.insert(0, str(ROOT / "common"))
import pandas as pd  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

import history_db as db  # noqa: E402
import trend_report as tr  # noqa: E402

A_OUT, OUT = Path(sys.argv[1]), Path(sys.argv[2])
OUT.mkdir(parents=True, exist_ok=True)
DB = A_OUT / "BU_History_DB.xlsx"

# 1) 과거 파일 업로드 (B, C) — 파일명 날짜는 생성일이라 측정일을 직접 지정
for tag, d in (("B_0610", date(2026, 6, 10)), ("C_0826", date(2026, 8, 26))):
    f = sorted(glob.glob(str(A_OUT / tag / "Daily data 통합정리_*.xlsx")))[-1]
    p = tr.parse_daily_file(f)
    print(f"{tag}: {len(p['rows'])}행, 모델 {p['model']} {p['code']}, 파일명 날짜 {p['date']}, LD {p['ld']}")
    added, skipped = db.append(DB, tr.rows_for_db(p, d, f))
    print(f"   적재 {added}행 (건너뜀 {skipped})")
    added, skipped = db.append(DB, tr.rows_for_db(p, d, f))
    assert added == 0, "같은 파일 재업로드는 0행이어야 함"

# 2) 이벤트 등록
events = pd.DataFrame({
    "이벤트명": ["이벤트 A", "이벤트 B", "이벤트 C", "G10 확인"],
    "모델": ["WINE", "WINE", "WINE", "VDA099"],
    "시작일": [date(2026, 4, 16), date(2026, 6, 8), date(2026, 8, 24), date(2026, 10, 1)],
    "종료일": [date(2026, 4, 22), date(2026, 6, 12), date(2026, 8, 28), date(2026, 10, 3)],
    "메모": ["", "", "오후 측정분 재측정(예시)", "LD 모델"]})
db.save_events(DB, events)
data, ev = db.load(DB)
m = db.assign_events(data, ev)
print(m.groupby(["이벤트", "매핑상태"]).size())
assert set(m["매핑상태"]) == {"정상"}, "모든 행이 이벤트에 매핑돼야 함"

# 3) 리포트
out = tr.generate(DB, ["이벤트 A", "이벤트 B", "이벤트 C", "G10 확인"], OUT / "트렌드리포트_test.xlsx", log=print)
wb = load_workbook(out)
print("시트:", wb.sheetnames)
ws = wb["이벤트 요약"]
for r in range(6, 10):
    print([ws.cell(r, c).value for c in range(2, 13)])
out2 = tr.generate(DB, ["이벤트 A", "이벤트 B", "이벤트 C"], OUT / "트렌드리포트_WINE.xlsx", log=print)
print("Tool C 시험 통과")
