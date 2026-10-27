# neovue-bu-tools

NEOVUE BU 설비 데이터 정리·분석 도구 모음. NEOVUE 표시는 창 제목·exe 이름에만 (결과 엑셀에는 넣지 않음). 세 툴 모두 같은 화면 디자인(머리글 · 카드 · 파란 실행 버튼)과 `(made by Kwakwonyong)` 표시.

## 포함 도구

### OneClick (`OneClick/`) — Tool A v1.5.3 (`NEOVUE_OneClick_정리`)
BU 설비 광학 데이터(CA410, LMK6)를 양식 Excel에 자동 정리하는 Tkinter GUI 도구.
v1.5부터 BU/WU 공정능력(Cp/Cpk) 시트, LD 모델 자동 체크, 누적 DB 적재를 지원.
v1.5.1부터 누적 DB 기본 위치가 `문서\BU_Tools\BU_History_DB.xlsx` (트렌드 리포트와 같은 파일).
v1.5.2부터 화면 디자인을 Tool B·트렌드 리포트와 통일.

- `One_click_정리.py` - 메인 Python 소스
- `NEOVUE_OneClick_정리_v1.5.3.spec` - PyInstaller spec (이전 버전 spec은 git 기록에)
- `양식_Daily data_BU설비_광학데이터_날짜.xlsx` - 기본 양식 파일

### BU_LotID (`BU_LotID/`) — Tool B v0.6.2 (`NEOVUE_BU_DATA_정리`)
BU/WU 이미지 파일을 LotID 단위로 묶어 Excel로 정리하는 Tkinter GUI 도구.
v0.6부터 `crop_report.xlsx`에 공정능력 시트 추가 (판정 = 그 실행에서 고른 측정 CSV 전체, 이미지 정리본 LotID는 참고로 표시).
설정 화면에서 끌 수 있음. 결과 시트는 v0.5와 같음.

- `BU_organize_gui_v05.py`, `BU_organize_one_click_v05.py` - 소스 (파일 이름은 v0.5 그대로)
- `NEOVUE_BU_DATA_정리_v0.6.2.spec` - PyInstaller spec
- `neovue_bu_data.ico` - 세 툴 공용 아이콘 (exe·창 아이콘)

### Trend (`Trend/`) — BU 트렌드 리포트 v0.5 (`NEOVUE_BU_TREND_리포트`)
측정일·이벤트 한 표, 통합정리 파일 끌어다 놓기/선택 → 이벤트만 입력, 이벤트별 트렌드 리포트(거시·미시).
v0.3: 리포트 '엑셀 열기'·'폴더 열기' 버튼, '변화 감지' 시트(읽는 법·경고 기준·의미/확인할 것).
v0.4: 이벤트 요약 핵심 그림(한눈 보드 · Cpk+NG율) · 결론 메모 칸, 미시 시트 이벤트별 접기.
v0.5: WU -100(측정 이상값) 제외·알람, 색좌표 확대 그림, 측정일별 색좌표 이동, 변화 감지 '내용' 열.

- `BU_trend_report.py` - GUI
- `trend_report.py` - 통합정리 파일 읽기·리포트 생성 로직
- `NEOVUE_BU_TREND_리포트_v0.5.spec` - PyInstaller spec

### 공용 모듈 (`common/`)
`capability.py`(계산) · `capability_sheet.py`(공정능력 시트) · `charts.py`(그림) · `history_db.py`(누적 DB) · `ui_style.py`(공용 화면 스타일 · 아이콘 · 파일 열기)

### 시험 (`tests/`)
- `make_test_datalogs.py` - 실측 Daily data로 시험용 CA410/LMK6 DataLog 생성(재측정 중복·LD 모델 포함)
- `test_common.py` - 공정능력 값(양식 19행 수식 포함)·LD 감지·누적 DB(중복 방지·날짜 묶기·공용 위치) 검증
- `run_tool_a.py`, `run_tool_b.py`, `run_tool_c.py` - 화면 없이 Tool A·Tool B·트렌드 리포트를 끝까지 실행

## 빌드

처음 한 번 패키지 설치 (v1.5부터 matplotlib, 트렌드 v0.2부터 tkinterdnd2):

```powershell
python -m pip install pandas openpyxl numpy pillow matplotlib tkinterdnd2 pyinstaller
```

각 도구 폴더에서 실행 (공용 모듈은 spec이 `../common`에서 함께 묶음).
`python -m PyInstaller`로 실행하면 패키지를 설치한 그 파이썬으로 빌드됩니다.

```powershell
cd OneClick;  python -m PyInstaller --noconfirm NEOVUE_OneClick_정리_v1.5.3.spec
cd ..\Trend;  python -m PyInstaller --noconfirm NEOVUE_BU_TREND_리포트_v0.5.spec
cd ..\BU_LotID; python -m PyInstaller --noconfirm NEOVUE_BU_DATA_정리_v0.6.2.spec
```

- 성공하면 마지막 줄에 `Build complete! The results are available in: <경로>`가 나오고, exe는 그 경로(명령을 실행한 폴더의 `dist`)에 생깁니다.
- `build`·`dist`가 비어 있으면 빌드가 중간에 멈춘 것입니다. 창에 나온 `[빌드 중단]` 또는 `ERROR`/`Traceback` 줄이 원인입니다.
- spec은 빌드 전에 공용 모듈(`common/`)과 필요한 패키지를 점검하고, 빠진 것이 있으면 설치 명령과 함께 멈춥니다.
- `Build complete!`까지 나왔는데 `dist`가 비어 있으면 백신(V3 등)이 exe를 검역한 경우입니다.

## 배포 zip

`python scripts/make_release_zip.py <출력.zip>` — 한글 파일 이름(spec·양식·체크리스트)이 Windows에서 깨지지 않게 묶음

## WU -100 (측정 이상값)

LMK6가 측정에 실패하면 WU에 -100을 남깁니다. 세 툴 모두 이 행을 Cp/Cpk·그림 계산에서 빼고 알림을 표시합니다.
(Tool A·B 공정능력 시트 4행, 트렌드 리포트 이벤트 요약·변화 감지 시트, 트렌드 Raw 시트 '분석 제외' 열)

## 누적 DB

- 위치: `문서\BU_Tools\BU_History_DB.xlsx` — Tool A와 트렌드 리포트가 exe 위치와 상관없이 같은 파일을 씀
- 이전 버전처럼 exe 옆에 DB가 있으면 처음 실행 때 공용 DB로 합침 (원본 파일은 그대로 둠). 다른 위치를 쓰려면 각 툴의 'DB 변경'
- 같은 측정(LotID+측정 시각, 또는 LotID+WU+BU 값)은 다시 넣지 않음 → Tool A 결과 파일을 트렌드에 다시 넣어도 중복 없음
- 데이터 삭제 전에는 DB 옆에 `BU_History_DB.backup.xlsx` 복사본을 남김

## 요구사항

- Python 3.11+
- pandas, openpyxl, numpy, Pillow, matplotlib (v1.5부터), tkinterdnd2 (트렌드 v0.2부터)
- PyInstaller (빌드 시)
