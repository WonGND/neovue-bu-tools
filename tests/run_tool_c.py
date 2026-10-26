"""트렌드 리포트 v0.4 로직 시험 (화면 없이)
사용: python run_tool_c.py <Tool A 시험 출력 폴더(DB 포함)> <리포트 출력 폴더> [Tool B 시험 출력 폴더]
1) 과거 파일(B, C) 추가 + 이벤트 지정 → 다시 추가해도 0건
2) Tool A가 이미 쌓은 날짜의 통합정리 파일을 추가하며 이벤트 지정 → 새로 넣지 않고 미지정 행만 묶음
3) 남은 미지정 날짜를 이벤트로 묶기 (G10은 LD 모델)
4) 리포트 생성 (기간은 데이터에서 자동) · 5) v0.3 변화 감지 시트(설명·기준·의미)"""
import glob
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "Trend"))
sys.path.insert(0, str(ROOT / "common"))
from openpyxl import load_workbook  # noqa: E402

import history_db as db  # noqa: E402
import trend_report as tr  # noqa: E402

A_OUT, OUT = Path(sys.argv[1]), Path(sys.argv[2])
OUT.mkdir(parents=True, exist_ok=True)
DB = A_OUT / "BU_History_DB.xlsx"


def add_files(files, label, dates=None, memo=None):
    """화면의 '파일 추가 → 이벤트 지정'과 같은 순서 (BU_trend_report.TrendApp._append)"""
    data, events = db.load(DB)
    total = {"added": 0, "skipped": 0, "relabeled": 0, "upgraded": 0}
    for i, f in enumerate(files):
        p = tr.parse_daily_file(f)
        d = (dates or {}).get(i, p["date"])
        data, st = db.merge_rows(data, tr.rows_for_db(p, d, f, label), label_dups=bool(label))
        for k in total:
            total[k] += st[k]
    db.save(DB, data, db.ensure_events(data, events, {label: memo} if (label and memo) else None))
    return total


def last(tag):
    return sorted(glob.glob(str(A_OUT / tag / "Daily data 통합정리_*.xlsx")))[-1]


d0, _ = db.load(DB)
n0 = len(d0)
print(f"시작: DB {n0}행 (Tool A 자동 적재분, 전부 미지정 {int((d0[db.LABEL] == '').sum())}행)")

# 1) 과거 파일 추가 + 이벤트 지정 (파일명 날짜는 생성일이라 측정일은 직접 지정)
for tag, d, ev in (("B_0610", date(2026, 6, 10), "이벤트 B"), ("C_0826", date(2026, 8, 26), "이벤트 C")):
    st = add_files([last(tag)], ev, {0: d}, memo="오후 측정분 재측정(예시)" if ev == "이벤트 C" else None)
    print(f"{tag} → {ev}: {st}")
    assert st["added"] > 0
    st = add_files([last(tag)], ev, {0: d})
    assert st["added"] == 0 and st["skipped"] > 0, "같은 파일 다시 추가는 0건"

# 2) Tool A가 이미 쌓은 날(04-21)의 통합정리 파일을 '이벤트 A'로 추가
st = add_files([last("A_0421")], "이벤트 A")
print("A_0421 통합정리 → 이벤트 A:", st)
assert st["added"] == 0 and st["relabeled"] == st["skipped"] > 0, "Tool A 적재분은 새로 넣지 않고 이벤트만 붙어야 함"

# 3) 남은 미지정 날짜 묶기
data, events = db.load(DB)
un = data[data[db.LABEL] == ""]
keys = sorted(set(db.day_key(un)))
print("미지정 날짜:", keys)
a_keys = [k for k in keys if k.startswith("2026-04")]          # 10-02는 G10(LD) 시험 데이터
g_keys = [k for k in keys if k not in a_keys]
db.set_label(DB, [("", k) for k in a_keys], "이벤트 A")
db.set_label(DB, [("", k) for k in g_keys], "G10 확인", "LD 모델")
data, events = db.load(DB)
assert (data[db.LABEL] != "").all(), "모든 행이 이벤트에 속해야 함"
print(events[["이벤트명", "모델", "시작일", "종료일", "메모", "구분"]].to_string(index=False))
print(data.groupby(db.LABEL).size().to_string())

# 4) 리포트
out = tr.generate(DB, ["이벤트 A", "이벤트 B", "이벤트 C", "G10 확인"], OUT / "트렌드리포트_test.xlsx", log=print)
wb = load_workbook(out)
print("시트:", wb.sheetnames)
ws = wb["이벤트 요약"]
h = next(r for r in range(1, 80) if ws.cell(r, 2).value == "이벤트" and ws.cell(r, 3).value == "기간")   # v0.4: 그림 아래로 내려감
for r in range(h + 1, h + 5):
    print([ws.cell(r, c).value for c in range(2, 13)])
assert ws.cell(h + 1, 3).value == "2026-04-16 ~ 2026-04-21", "기간은 데이터에서 자동 계산"
assert ws.cell(h + 4, 2).value == "G10 확인" and ws.cell(h + 4, 12).value == "N/A (LD)", "G10은 LD로 BU 제외"
# 5) v0.3 변화 감지 시트
assert wb.sheetnames[:2] == ["이벤트 요약", "변화 감지"], wb.sheetnames
wc = wb["변화 감지"]
head = [wc.cell(10, c).value for c in range(2, 10)]
assert head[-1] == "의미 · 확인할 것" and head[5] == "경고 기준", head
rows = [[wc.cell(r, c).value for c in range(2, 10)] for r in range(11, 40) if wc.cell(r, 2).value]
hits = [r for r in rows if str(r[6]).startswith("⚠")]
assert hits and all(r[7] for r in hits), "⚠ 행에는 의미·확인할 것이 있어야 함"
assert all(r[1] for r in rows[:len(hits)]), "비교 대상(직전 이벤트) 표시"
link = [ws.cell(r, 4) for r in range(h, h + 15) if ws.cell(r, 4).hyperlink]
assert link and "확인 필요" in link[0].value and str(len(hits)) in link[0].value, "요약 시트에 건수 + 바로가기"
print(f"변화 감지 시트 OK — 확인 필요 {len(hits)}건 / {len(rows)}행 · 요약: {link[0].value}")
# 6) v0.4 배치: 요약 핵심 그림 2개 + 결론 메모 칸 · 미시 이벤트 접기 · NG 이력 이미지 링크
assert len(ws._images) == 2, "요약 시트 핵심 그림 A·B"
assert any("결론 메모" == ws.cell(r, 2).value for r in range(1, 60)), "결론 메모 칸"
wm = wb["측정일별(미시)"]
levels = [r for r, d in wm.row_dimensions.items() if d.outlineLevel == 1]
assert levels and all(type(im.anchor).__name__ == "TwoCellAnchor" for im in wm._images), "행 그룹 + 그림이 행과 같이 접힘"
links = [c for c in wm[5] if c.hyperlink]
assert len(links) == 4, "미시 바로가기 4개"
B_OUT = Path(sys.argv[3]) if len(sys.argv) > 3 else None
if B_OUT:
    out2 = tr.generate(DB, ["이벤트 A", "이벤트 B", "이벤트 C", "G10 확인"], OUT / "트렌드리포트_링크.xlsx", log=print,
                       image_dir=B_OUT)
    wn = load_workbook(out2)["NG 이력"]
    assert [wn.cell(1, c).value for c in (8, 9)] == ["BU Image", "WU Image"]
    linked = [r for r in range(2, wn.max_row + 1) if wn.cell(r, 8).hyperlink]
    assert linked, "Tool B 결과에 있는 Lot ID는 링크"
    tgt = Path(wn.cell(linked[0], 8).hyperlink.target)
    assert tgt.exists() and tgt.name.startswith(str(wn.cell(linked[0], 3).value)[:10]), tgt
    s1, s2 = (OUT / "트렌드리포트_test.xlsx").stat().st_size, Path(out2).stat().st_size
    print(f"NG 이미지 링크 {len(linked)}행 · 파일 크기 {s1/1024:.0f}KB → {s2/1024:.0f}KB")
    assert s2 - s1 < 50 * 1024, "링크만 넣으므로 크기 차이 작음"
tr.generate(DB, ["이벤트 A", "이벤트 B", "이벤트 C"], OUT / "트렌드리포트_WINE.xlsx", log=print)
print("트렌드 리포트 v0.4 시험 통과")
