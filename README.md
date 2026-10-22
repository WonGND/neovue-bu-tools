# tovis-bu-tools

TOVIS(NEOVUE) BU 설비 데이터 정리·분석 도구 모음.

## 포함 도구

### OneClick (`OneClick/`) — Tool A v1.5
BU 설비 광학 데이터(CA410, LMK6)를 양식 Excel에 자동 정리하는 Tkinter GUI 도구.
v1.5부터 BU/WU 공정능력(Cp/Cpk) 시트, LD 모델 자동 체크, 누적 DB 적재를 지원.

- `One_click_정리.py` - 메인 Python 소스
- `OneClick_정리_v1.5.spec` - PyInstaller spec (v1.5, 공용 모듈 포함)
- `OneClick_정리_v1.4.spec`, `OneClick_정리_v1.4.1.spec` - 이전 버전 spec
- `양식_Daily data_BU설비_광학데이터_날짜.xlsx` - 기본 양식 파일

### BU_LotID (`BU_LotID/`) — Tool B v0.5
BU/WU 이미지 파일을 LotID 단위로 묶어 Excel로 정리하는 Tkinter GUI 도구 (변경 없음).

### Trend (`Trend/`) — BU 트렌드 리포트 v0.1 (신규)
이벤트 기간 관리, 과거 Daily data 업로드, 이벤트별 트렌드 리포트(거시·미시) 생성.

- `BU_trend_report_v01.py` - GUI
- `trend_report.py` - 과거 파일 읽기·리포트 생성 로직
- `BU_TREND_리포트_v0.1.spec` - PyInstaller spec

### 공용 모듈 (`common/`)
`capability.py`(계산) · `capability_sheet.py`(공정능력 시트) · `charts.py`(그림) · `history_db.py`(누적 DB)

### 시험 (`tests/`)
- `make_test_datalogs.py` - 실측 Daily data로 시험용 CA410/LMK6 DataLog 생성(재측정 중복·LD 모델 포함)
- `test_common.py` - 공정능력 값(양식 19행 수식 포함)·LD 감지·DB 중복 방지 검증
- `run_tool_a.py`, `run_tool_c.py` - 화면 없이 Tool A·트렌드 리포트를 끝까지 실행

## 빌드

각 도구 폴더에서 실행 (공용 모듈은 spec이 `../common`에서 함께 묶음).
`python -m PyInstaller`로 실행하면 패키지를 설치한 그 파이썬으로 빌드됩니다.

```powershell
cd OneClick;  python -m PyInstaller --noconfirm OneClick_정리_v1.5.spec
cd ..\Trend;  python -m PyInstaller --noconfirm BU_TREND_리포트_v0.1.spec
cd ..\BU_LotID; python -m PyInstaller --noconfirm TOVIS_BU_DATA_정리_v0.5.spec
```

- 성공하면 마지막 줄에 `Build complete! The results are available in: <경로>`가 나오고, exe는 그 경로(명령을 실행한 폴더의 `dist`)에 생깁니다.
- `build`·`dist`가 비어 있으면 빌드가 중간에 멈춘 것입니다. 창에 나온 `[빌드 중단]` 또는 `ERROR`/`Traceback` 줄이 원인입니다.
- v1.5·트렌드 spec은 빌드 전에 공용 모듈(`common/`)과 필요한 패키지(pandas·openpyxl·numpy·matplotlib)를 점검하고, 빠진 것이 있으면 설치 명령과 함께 멈춥니다.
- `Build complete!`까지 나왔는데 `dist`가 비어 있으면 백신(V3 등)이 exe를 검역한 경우입니다.

Tool A와 트렌드 리포트 exe를 **같은 폴더**에 두면 누적 DB(`BU_History_DB.xlsx`)를 자동으로 같이 씁니다.

## 요구사항

- Python 3.11+
- pandas, openpyxl, numpy, Pillow, matplotlib (v1.5부터)
- PyInstaller (빌드 시)
