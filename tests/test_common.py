"""공통 모듈 검증. 사용: python test_common.py <Daily data xlsx> <임시 폴더>
1) 공정능력: 첨부 WINE 775ea로 이전 계산값·양식 19행 수식(Ppk)과 일치
2) LD 감지 3) 누적 DB 중복 방지·이벤트 매핑"""
import sys
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "common"))
import numpy as np  # noqa: E402
import openpyxl  # noqa: E402
import pandas as pd  # noqa: E402

import history_db as db  # noqa: E402
from capability import capability, grade, is_constant  # noqa: E402

SRC, TMP = Path(sys.argv[1]), Path(sys.argv[2])
TMP.mkdir(parents=True, exist_ok=True)
ok = 0


def check(cond, msg):
    global ok
    assert cond, msg
    ok += 1
    print("  OK", msg)


ws = openpyxl.load_workbook(SRC, data_only=True).worksheets[0]
df = pd.DataFrame([[ws.cell(r, 3).value, ws.cell(r, 19).value, ws.cell(r, 20).value] for r in range(20, 795)],
                  columns=["LotID", "WU", "BU"])
df["k"] = df.LotID.str[32:38] + df.LotID.str[-5:]
df = df.sort_values("k")
print("[1] 공정능력 (WINE 775ea)")
wu, bu = capability(df.WU, "WU"), capability(df.BU, "BU")
for c, exp in ((wu, (2.05, 1.04, 0.88)), (bu, (1.07, 0.53, 0.45))):
    got = (round(c.cp, 2), round(c.cpk, 2), round(c.ppk, 2))
    check(got == exp, f"{c.name} Cp/Cpk/Ppk = {got} (기대 {exp})")
for name, lsl in (("WU", 80), ("BU", 50)):
    x = df[name].astype(float)
    template = (x.mean() - lsl) / (3 * x.std(ddof=1))        # 양식 19행: (AVERAGE-LSL)/(3*STDEV)
    c = capability(x, name)
    check(abs(template - c.ppk) < 1e-9, f"{name} Ppk = 양식 19행 수식 ({template:.4f})")
check(wu.ng == 10 and bu.ng == 44, f"실측 NG WU {wu.ng} / BU {bu.ng} (기대 10 / 44)")
check(grade(1.04)[0] == "주의" and grade(0.53)[0] == "부적합" and grade(2.05)[0] == "양호", "판정 등급")
check(any("비정규" in w for w in wu.warnings), "WU 비정규 경고")

print("[2] LD 감지")
check(is_constant([50.0] * 30), "BU 전부 50 → LD")
check(not is_constant(df.BU), "WINE BU → LD 아님")
ld = capability([50.0] * 30, "BU", ld=True)
check(ld.verdict == "N/A (LD 모델)" and ld.n == 0, "LD 모델 BU는 계산 제외")

print("[3] 누적 DB")
p = TMP / "test_db.xlsx"
p.unlink(missing_ok=True)
rows = pd.DataFrame({"LotID": ["GJ96111000-1032601120-VMA129BLAC01260415A00014", "X2"],
                     "측정일시": [datetime(2026, 4, 16, 8, 30), datetime(2026, 4, 16, 9, 0)],
                     "WU": [86.3, 85.1], "BU": [61.2, 55.0], "모델명": ["WINE", "WINE"]})
check(db.append(p, rows) == (2, 0), "첫 적재 2행")
check(db.append(p, rows) == (0, 2), "같은 파일 다시 적재 → 0행 (중복 건너뜀)")
data, events = db.load(p)
check(data.LotID.iloc[0] == "GJ961110001032601120VMA129BLAC01260415A00014", "LotID 정규화(하이픈 제거)")
check(db.prod_date(data.LotID.iloc[0]) == date(2026, 4, 15), "LotID 생산일 추출")
events = pd.DataFrame({"이벤트명": ["A", "B"], "모델": ["WINE", ""], "시작일": [date(2026, 4, 15), date(2026, 4, 16)],
                       "종료일": [date(2026, 4, 20), date(2026, 4, 16)], "메모": ["", ""]})
db.save_events(p, events)
data, events = db.load(p)
m = db.assign_events(data, events)
check(m["매핑상태"].iloc[0].startswith("겹침"), "두 이벤트 기간이 겹치면 '겹침'")
events.loc[1, "시작일"] = date(2026, 5, 1)
events.loc[1, "종료일"] = date(2026, 5, 2)
m = db.assign_events(data, events)
check(list(m["이벤트"]) == ["A", "A"], "기간 안 → 이벤트 A")
events.loc[0, "모델"] = "IND"
check(set(db.assign_events(data, events)["매핑상태"]) == {"기간 밖"}, "모델 불일치 → 기간 밖")
print(f"전체 {ok}건 통과")
