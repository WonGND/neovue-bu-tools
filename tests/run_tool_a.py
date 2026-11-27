"""Tool A v1.6.2 처리 로직을 화면 없이 끝까지 실행해 검증.
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
# v1.6: 파일 종류 확인 · 짝 파일 찾기 · 진행률 · 쉬운 오류 문구
import os  # noqa: E402
import ui_style as ui  # noqa: E402
ca, lmk = TD / "CA410_DataLog_A_0418.csv", TD / "LMK6DataLog_A_0418.csv"
assert oc.sniff_kind(str(ca)) == "CA410" and oc.sniff_kind(str(lmk)) == "LMK6", "열 이름으로 설비 종류 구분"
assert os.path.basename(oc.find_pair(str(ca), "LMK6")) == lmk.name, oc.find_pair(str(ca), "LMK6")
assert os.path.basename(oc.find_pair(str(lmk), "CA410")) == ca.name, "반대 방향도"
prog = []
oc.process({"ca_path": str(ca), "lmk_path": str(lmk), "template_path": str(TEMPLATE), "save_dir": str(OUT),
            "make_individual": False, "make_cpk": True, "ld": False, "append_db": False, "db_path": ""},
           log=lambda *_: None, progress=lambda p_, t_: prog.append(p_))
assert prog == sorted(prog) and prog[-1] == 100 and len(prog) >= 5, prog
e = PermissionError(13, "Permission denied", "C:/x/결과.xlsx")
assert "엑셀에서 그 파일을 닫고" in ui.friendly_error(e) and "결과.xlsx" in ui.friendly_error(e)
assert "찾을 수 없습니다" in ui.friendly_error(FileNotFoundError(2, "No such file", "a.csv"))
try:
    try:
        raise PermissionError(13, "denied", "db.xlsx")
    except PermissionError as inner:
        raise RuntimeError("저장 실패") from inner
except RuntimeError as outer:
    assert "닫고" in ui.friendly_error(outer), "원인 오류까지 따라가서 안내"
print(f"v1.6 확인 OK — 짝 파일 · 진행률 {len(prog)}단계 · 쉬운 오류 문구")
# v1.6.1: 호기별 파일 형상 (3호기 LVDataLog / 4호기 CA410DataLog — 열 이름은 실제 파일 머리글, 값·ID는 가상)
import tempfile  # noqa: E402
H3 = ("Time,Panel,Panel_ID,Judge,Stage_No,White_SX,White_SY,White_LV,Red_X,Red_Y,Green_X,Green_Y,Blue_X,Blue_Y,"
      "Black_LV,CR,EEPROM_36,EEPROM_59,EEPROM_Verify")
H4 = ("Time,Panel,Panel_ID,Judge,Stage_No,WPC_White_SX,WPC_White_SY,WPC_White_LV,WPC_GCS,WPC_ASIL_LED,White_SX,White_SY,"
      "White_LV,Red_X,Red_Y,Green_X,Green_Y,Blue_X,Blue_Y,Black_LV,CR,EEPROM_36,EEPROM_59,EEPROM_Verify")
HL = "Time,Panel,Panel_ID,Judge,Stage_No,Black_Uniformity,White_Uniformity,9P_White_Uniformity"
for hogi, col_name, head in (("3호기", "LVDataLog.csv", H3), ("4호기", "CA410DataLog.csv", H4)):
    d = Path(tempfile.mkdtemp())
    ids = ([f"XX00000000-0000000000-VMA129BLAC01261110A{i:05d}" for i in range(1, 7)] if hogi == "3호기"
           else [f"XX000000000VMA129BLAC01260202A{i:05d}" for i in range(1, 7)])
    ca, lm = [head], [HL]
    for i, pid in enumerate(ids + ids[:1]):          # 마지막 줄 = 첫 패널 재측정
        t = f"2026.11.10 09:{i:02d}:00"
        eep = "06" if pid == ids[0] else "1B"
        if hogi == "3호기":
            ca.append(f"{t},Panel1,{pid},OK,1,0.2948,0.3224,906.4,0.6434,0.3513,0.2990,0.6518,0.1470,0.0612,0.65,1389.7,09,{eep},OK")
        else:
            ca.append(f"{t},Panel1,{pid},OK,4,0.2931,0.3180,1434.7,0.0008,OK, 0.2937,0.3197,1016.6,0.6437,0.3510,"
                      f"0.2992,0.6516,0.1471,0.0630,0.7362,1380.9,16,{eep},OK")
        lm.append(f"{t},Panel1,{pid},OK,1,{60 + i},{88 - i * 0.5},90")
    (d / col_name).write_text("\n".join(ca) + "\n")
    (d / "LMK6DataLog.csv").write_text("\n".join(lm) + "\n")
    cpath, lpath = str(d / col_name), str(d / "LMK6DataLog.csv")
    assert oc.sniff_kind(cpath) == "CA410" and oc.sniff_kind(lpath) == "LMK6", hogi
    assert oc.find_pair(cpath, "LMK6") == lpath and oc.find_pair(lpath, "CA410") == cpath, f"{hogi} 짝 파일"
    r = oc.process({"ca_path": cpath, "lmk_path": lpath, "template_path": str(TEMPLATE), "save_dir": str(d),
                    "make_individual": False, "make_cpk": True, "ld": False, "append_db": True,
                    "db_path": str(d / "DB.xlsx")}, log=lambda *_: None)
    ws = load_workbook(r["out_path"]).worksheets[0]
    got = {ws.cell(row, 3).value: ws.cell(row, 27).value for row in range(20, 26)}
    assert len(got) == 6 and got[ids[0].replace("-", "")] == "06", (hogi, got)   # 재측정은 최신 1건, EEPROM 앞 0 유지
    assert all(v == "1B" for k, v in got.items() if k != ids[0].replace("-", "")), got
    assert r["db"] == (6, 0), r["db"]
    print(f"v1.6.1 {hogi} ({col_name}) 형상 OK — 짝 파일 · 중복 제거 · EEPROM '06' 유지 · DB 6행")
print("Tool A v1.6.2 시험 통과")
