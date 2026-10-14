# tovis-bu-tools

TOVIS BU 설비 데이터 정리 도구 모음.

## 포함 도구

### OneClick (`OneClick/`)
BU 설비 광학 데이터(CA, LMK)를 양식 Excel에 자동 정리하는 Tkinter GUI 도구.

- `One_click_정리.py` - 메인 Python 소스
- `OneClick_정리_v1.4.spec` - PyInstaller spec (v1.4)
- `OneClick_정리_v1.4.1.spec` - PyInstaller spec (v1.4.1)
- `양식_Daily data_BU설비_광학데이터_날짜.xlsx` - 기본 양식 파일

### BU_LotID (`BU_LotID/`)
BU/WU 이미지 파일을 LotID 단위로 묶어 Excel로 정리하는 Tkinter GUI 도구 (v0.5).

- `BU_organize_one_click_v05.py` - 파이프라인 로직
- `BU_organize_gui_v05.py` - Tkinter GUI
- `TOVIS_BU_DATA_정리_v0.5.spec` - PyInstaller spec
- `tovis_bu_data_settings_v05.json` - 기본 설정값
- `tovis_bu_data.ico` - 앱 아이콘

## 빌드

각 도구 폴더에서 실행:

```powershell
pyinstaller OneClick_정리_v1.4.1.spec
# or
pyinstaller TOVIS_BU_DATA_정리_v0.5.spec
```

## 요구사항

- Python 3.11+
- pandas, openpyxl, numpy, Pillow
- PyInstaller (빌드 시)
