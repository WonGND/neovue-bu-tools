"""Tool A v1.5 처리 로직을 화면 없이 끝까지 실행해 검증.
사용: python run_tool_a.py <testdata 폴더> <출력 폴더>
- 이벤트 A(3일치): 공정능력 시트 + 누적 DB 적재
- 이벤트 B·C: 통합정리 파일만 생성(Tool C 과거 파일 업로드 시험용)
- G10: LD 자동 감지 → BU 제외
"""
import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("oneclick", ROOT / "OneClick" / "One_click_정리.py")
oc = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oc)
from openpyxl import load_workbook  # noqa: E402

TD, OUT = Path(sys.argv[1]), Path(sys.argv[2])
OUT.mkdir(parents=True, exist_ok=True)
TEMPLATE = ROOT / "OneClick" / "양식_Daily data_BU설비_광학데이터_날짜.xlsx"
DB = OUT / "BU_History_DB.xlsx"
DB.unlink(missing_ok=True)
logs = []


def run(tag, append_db, make_cpk=True):
    lmk = str(TD / f"LMK6DataLog_{tag}.csv")
    ld = oc.detect_ld(lmk)
    lines = []
    res = oc.process({"ca_path": str(TD / f"CA410_DataLog_{tag}.csv"), "lmk_path": lmk, "template_path": str(TEMPLATE),
                      "save_dir": str(OUT / tag), "make_individual": False, "make_cpk": make_cpk, "ld": ld,
                      "append_db": append_db, "db_path": str(DB)}, lines.append)
    print(f"--- {tag} (LD 자동 감지: {ld})")
    for line in lines:
        print("   ", line)
    return res, ld


for tag in ("A_0416", "A_0418", "A_0421", "B_0610", "C_0826", "G10_LD"):
    (OUT / tag).mkdir(exist_ok=True)
    res, ld = run(tag, append_db=tag.startswith(("A_", "G10")))
    wb = load_workbook(res["out_path"])
    assert "공정능력_BU_WU" in wb.sheetnames, "공정능력 시트 없음"
    ws = wb["공정능력_BU_WU"]
    assert ws["B7"].value == "BU" and (ws["J7"].value == "N/A") == ld, "LD 표기 확인"
# 같은 날 파일을 다시 넣어도 DB 행이 늘지 않아야 함
res, _ = run("A_0418", append_db=True)
assert res["db"][0] == 0, "중복 적재 방지 실패"
import pandas as pd  # noqa: E402
d = pd.read_excel(DB, sheet_name="data")
print("DB 행 수:", len(d), "(기대 57+153+565+30 = 805)")
assert len(d) == 805
print("Tool A v1.5 시험 통과")
