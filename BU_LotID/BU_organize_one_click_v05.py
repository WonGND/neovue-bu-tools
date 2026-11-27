import csv
import io
import json
import os
import re
import shutil
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

import sys

import numpy as np
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import PatternFill
from PIL import Image

# v0.6: 공용 모듈(common/) 경로 — 소스 실행 시 ../common, exe는 spec에서 함께 묶음
_HERE = Path(__file__).resolve().parent
for _p in (_HERE.parent / "common", _HERE / "common"):
    if _p.exists() and str(_p) not in sys.path:
        sys.path.insert(0, str(_p))
TOOL_VERSION = "v0.8"

# 처리 대상 이미지 확장자
ALLOWED_EXTENSIONS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")
# 파일명에서 LotID/종류(BU, WU)를 뽑기 위한 패턴
LOT_PATTERN = re.compile(r"^(?P<lotid>.+)_(?P<kind>BU|WU)_\d+$", re.IGNORECASE)
DATA_FILE_PATTERN = "LMK6DataLog.csv"
BU_SPEC_MIN = 50.0
WU_SPEC_MIN = 80.0
BU_GRID_COLS = 48
BU_GRID_ROWS = 27
MODEL_NAME_CANDIDATES = (
    "Model_Name", "ModelName", "Model", "RecipeName", "Recipe", "Product", "Product_Name",
)
LOT_ID_CANDIDATES = ("Panel_ID", "LotID", "Lot_ID", "PanelID", "Panel Id")
# v0.8: 이 툴이 만드는 결과 폴더 이름 끝 — 측정 CSV·폴더 미리보기에서 훑지 않음 (이미지 수천 장)
OUTPUT_DIR_SUFFIXES = ("_LotID_latest_v1", "_LotID_latest_v1_cropped_v1")
CROP_CACHE_NAME = ".crop_cache.json"      # v0.8: 이전 실행의 자르기 결과 기억 (같은 원본·같은 설정이면 다시 안 자름)
CROP_CACHE_VERSION = 1
THUMB_WIDTH_PX = 240                       # 엑셀 결과 시트에 넣는 이미지 폭


def is_output_dir(name: str) -> bool:
    return name.endswith(OUTPUT_DIR_SUFFIXES)


def walk_files(root: Path, skip_output_dirs: bool = False):
    """os.walk 한 번으로 모든 파일 경로를 돌려줌 (rglob을 확장자마다 반복하지 않도록)"""
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        if skip_output_dirs:
            dirnames[:] = [d for d in dirnames if not is_output_dir(d)]
        base = Path(dirpath)
        for name in sorted(filenames):
            yield base / name

class PipelineCancelled(Exception):
    pass

def resized_png_bytes(image_path: Path, max_width_px: int) -> bytes | None:
    """엑셀에 넣을 축소 이미지(PNG) — v0.8부터 자르기 작업 스레드에서 미리 만들어 둠"""
    if not image_path.exists():
        return None
    try:
        with Image.open(image_path) as img:
            w, h = img.size
            if w > max_width_px:
                ratio = max_width_px / w
                new_w, new_h = int(w * ratio), int(h * ratio)
                img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format="PNG")
            return img_byte_arr.getvalue()
    except Exception as e:
        print(f"Error resizing image {image_path}: {e}")
        return None


def xl_image_from_bytes(data: bytes | None) -> XLImage | None:
    return XLImage(io.BytesIO(data)) if data else None


def get_resized_xl_image(image_path: Path, max_width_px: int) -> XLImage | None:
    return xl_image_from_bytes(resized_png_bytes(image_path, max_width_px))

def print_progress(label: str, current: int, total: int, done: bool = False) -> None:
    if total <= 0: return
    percent = (current / total) * 100
    print(f"{label}: {current}/{total} ({percent:5.1f}%)", flush=True)

def ensure_not_cancelled(cancel_check=None) -> None:
    if cancel_check and cancel_check():
        raise PipelineCancelled("사용자 요청으로 작업이 중지되었습니다.")

def parse_lot_kind(filename: str) -> tuple[str, str]:
    match = LOT_PATTERN.search(filename)
    if match:
        return match.group("lotid"), match.group("kind").upper()
    return filename, "UNKNOWN"

def normalize_lot_id(value) -> str:
    text = "" if value is None else str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return re.sub(r"[\s_-]+", "", text).upper()

def first_row_value(row: dict, candidates: tuple[str, ...]) -> str:
    for key in candidates:
        value = row.get(key)
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""

def parse_lmk_time(value, fallback_path: Path) -> datetime:
    raw = "" if value is None else str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y.%m.%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            pass
    return datetime.fromtimestamp(fallback_path.stat().st_mtime)

def iter_csv_dict_rows(csv_path: Path):
    encodings = ("utf-8-sig", "cp949", "euc-kr", "utf-8")
    last_error = None
    for encoding in encodings:
        try:
            with csv_path.open("r", encoding=encoding, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError as exc:
            last_error = exc
    raise UnicodeDecodeError(
        getattr(last_error, "encoding", "unknown"),
        getattr(last_error, "object", b""),
        getattr(last_error, "start", 0),
        getattr(last_error, "end", 1),
        f"{csv_path} 파일 인코딩을 읽지 못함",
    )

def collect_latest_lotid_folders(integrated_root: Path, cancel_check=None) -> tuple[dict[str, Path], list[dict]]:
    lotid_folders = [d for d in integrated_root.iterdir() if d.is_dir()]
    total = len(lotid_folders)
    print(f"\n[1/7] LotID 폴더 스캔 시작 (대상 폴더: {total}개)")

    latest_by_lotid: dict[str, Path] = {}
    all_rows = []

    for idx, folder in enumerate(lotid_folders, start=1):
        ensure_not_cancelled(cancel_check)
        lot_id = folder.name
        created_time = folder.stat().st_ctime
        
        row = {
            "lot_id": lot_id,
            "folder_path": str(folder),
            "created_time": datetime.fromtimestamp(created_time).strftime("%Y-%m-%d %H:%M:%S"),
        }

        if lot_id not in latest_by_lotid or created_time > latest_by_lotid[lot_id].stat().st_ctime:
            latest_by_lotid[lot_id] = folder

        all_rows.append(row)
        if idx == 1 or idx % 50 == 0 or idx == total:
            print_progress("  스캔 진행", idx, total, done=(idx == total))

    return latest_by_lotid, all_rows

def _same_file(src: Path, dst: Path) -> bool:
    """copy2는 수정 시각을 그대로 옮기므로 크기·수정 시각이 같으면 이전 실행에서 복사한 같은 파일"""
    try:
        a, b = src.stat(), dst.stat()
    except OSError:
        return False
    return a.st_size == b.st_size and abs(a.st_mtime - b.st_mtime) < 2.0   # FAT/네트워크 드라이브 2초 단위 대비


def _remove(path: Path) -> None:
    if path.is_dir() and not path.is_symlink():
        shutil.rmtree(path)
    else:
        path.unlink()


def sync_tree(src: Path, dst: Path, cancel_check=None) -> tuple[int, int]:
    """dst를 src와 같게 맞춤 — 바뀐 파일만 복사, src에 없는 것은 지움. (복사 수, 건너뜀 수)"""
    copied = skipped = 0
    for dirpath, dirnames, filenames in os.walk(src):
        ensure_not_cancelled(cancel_check)
        dirnames.sort()
        rel = Path(dirpath).relative_to(src)
        out_dir = dst / rel
        if out_dir.exists() and not out_dir.is_dir():
            out_dir.unlink()
        out_dir.mkdir(parents=True, exist_ok=True)
        keep = set(dirnames) | set(filenames)
        for extra in list(out_dir.iterdir()):
            if extra.name not in keep:
                _remove(extra)
        for name in filenames:
            s, d = Path(dirpath) / name, out_dir / name
            if d.is_dir():
                shutil.rmtree(d)
            if _same_file(s, d):
                skipped += 1
            else:
                shutil.copy2(s, d)
                copied += 1
    return copied, skipped


def copy_latest_folders(latest_by_lotid: dict[str, Path], dst_root: Path, cancel_check=None) -> None:
    """v0.8: 결과 폴더를 통째로 지우고 다시 복사하지 않고, 바뀐 파일만 복사 (결과 폴더 내용은 같음)"""
    dst_root.mkdir(parents=True, exist_ok=True)

    total = len(latest_by_lotid)
    print(f"\n[2/7] 최신 LotID 폴더 복사 시작 (대상: {total}개)")

    for stale in list(dst_root.iterdir()):          # 이번 대상이 아닌 LotID 폴더(또는 파일)는 정리
        if stale.name not in latest_by_lotid:
            _remove(stale)

    n_copy = n_skip = 0
    for idx, (lot_id, src_path) in enumerate(latest_by_lotid.items(), start=1):
        ensure_not_cancelled(cancel_check)
        c, k = sync_tree(src_path, dst_root / lot_id, cancel_check)
        n_copy += c
        n_skip += k
        if idx == 1 or idx % 20 == 0 or idx == total:
            print_progress("  복사 진행", idx, total, done=(idx == total))
    print(f"  새로 복사 {n_copy:,}개 · 이미 있어 건너뜀 {n_skip:,}개")

def collect_latest_measurements(data_root: Path, cancel_check=None) -> tuple[dict[str, dict], list[dict]]:
    target = DATA_FILE_PATTERN.lower()
    csv_files = [f for f in walk_files(data_root, skip_output_dirs=True) if f.name.lower() == target]
    total = len(csv_files)
    print(f"\n[4/7] 측정 CSV 스캔 시작 (대상 파일: {total}개)")

    latest_measurements: dict[str, dict] = {}
    all_measurement_rows = []
    if total == 0:
        print("  측정 CSV를 찾지 못해서 판정/BU/WU data는 공란으로 둡니다.")
        return latest_measurements, all_measurement_rows

    for idx, csv_path in enumerate(csv_files, start=1):
        ensure_not_cancelled(cancel_check)
        try:
            for row_data in iter_csv_dict_rows(csv_path):
                lot_id = first_row_value(row_data, LOT_ID_CANDIDATES)
                if not lot_id:
                    continue

                time_str = str(row_data.get("Time", "")).strip()
                time_val = parse_lmk_time(time_str, csv_path)

                m_row = {
                    "lot_id": lot_id,
                    "judge": str(row_data.get("Judge", "")).strip(),
                    "black_uniformity": str(row_data.get("Black_Uniformity", "")).strip(),
                    "white_uniformity": str(row_data.get("White_Uniformity", "")).strip(),
                    "time_raw": time_val,
                    "time_str": time_str,
                    "model_name": first_row_value(row_data, MODEL_NAME_CANDIDATES),
                    "source_file": str(csv_path),
                }

                normalized_lot_id = normalize_lot_id(lot_id)
                if (
                    normalized_lot_id not in latest_measurements
                    or time_val > latest_measurements[normalized_lot_id]["time_raw"]
                ):
                    latest_measurements[normalized_lot_id] = m_row
                all_measurement_rows.append(m_row)
        except Exception as exc:
            print(f"  CSV 읽기 실패: {csv_path} ({exc})")

        if idx == 1 or idx % 50 == 0 or idx == total:
            print_progress("  CSV 스캔 진행", idx, total, done=(idx == total))

    return latest_measurements, all_measurement_rows

def find_measurement(lot_id: str, latest_measurements: dict[str, dict]) -> dict:
    normalized = normalize_lot_id(lot_id)
    if normalized in latest_measurements:
        return latest_measurements[normalized]
    for key, measurement in latest_measurements.items():
        if normalized and (normalized in key or key in normalized):
            return measurement
    return {}

def excel_measurement_value(value):
    text = "" if value is None else str(value).strip()
    if not text:
        return ""
    try:
        return float(text)
    except ValueError:
        return text

def find_non_black_bbox(img: Image.Image, threshold: int = 12) -> tuple[int, int, int, int] | None:
    """임계값보다 밝은 픽셀의 경계 (x_min, y_min, x_max, y_max, 끝 포함).
    v0.8: numpy 좌표 배열 대신 PIL 내장 getbbox — 같은 결과, 큰 이미지에서 수십 배 빠름"""
    lut = [255 if v > threshold else 0 for v in range(256)]
    box = img.convert("L").point(lut).getbbox()
    if box is None:
        return None
    return (box[0], box[1], box[2] - 1, box[3] - 1)

def _load_crop_cache(cropped_root: Path, threshold: int, padding: int) -> dict:
    path = cropped_root / CROP_CACHE_NAME
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    if data.get("v") != CROP_CACHE_VERSION or data.get("threshold") != threshold or data.get("padding") != padding:
        return {}                                 # 설정이 바뀌면 전부 다시 자름
    return data.get("files", {})


def _save_crop_cache(cropped_root: Path, threshold: int, padding: int, files: dict) -> None:
    try:
        (cropped_root / CROP_CACHE_NAME).write_text(
            json.dumps({"v": CROP_CACHE_VERSION, "threshold": threshold, "padding": padding, "files": files}),
            encoding="utf-8")
    except OSError as exc:
        print(f"  (자르기 기록 저장 실패 — 다음 실행 때 전부 다시 자름: {exc})")


def _list_images(merged_root: Path) -> list[Path]:
    """확장자 순서(png → jpg → …)대로 — v0.7의 확장자별 rglob과 같은 순서"""
    order = {ext: i for i, ext in enumerate(ALLOWED_EXTENSIONS)}
    files = [f for f in walk_files(merged_root) if f.suffix.lower() in order]
    return sorted(files, key=lambda f: order[f.suffix.lower()])


def crop_images(merged_root: Path, cropped_root: Path, threshold: int, padding: int, cancel_check=None,
                thumb_width_px: int | None = THUMB_WIDTH_PX) -> list[dict]:
    """v0.8: 결과 폴더를 지우지 않고, 원본(크기·수정 시각)과 설정이 같은 이미지는 이전 결과를 그대로 씀.
    엑셀용 축소 이미지도 여기서(여러 스레드로) 미리 만듦."""
    cropped_root.mkdir(parents=True, exist_ok=True)
    cache = _load_crop_cache(cropped_root, threshold, padding)

    image_files = _list_images(merged_root)
    total = len(image_files)
    print(f"\n[5/7] 이미지 크롭 시작 (대상: {total}개, 멀티스레드 활성화)")

    # 이번 대상이 아닌 이전 결과 이미지는 정리 (crop_report.xlsx·기록 파일은 둠)
    wanted = {str(f.relative_to(merged_root)) for f in image_files}
    for f in list(walk_files(cropped_root)):
        if f.suffix.lower() in ALLOWED_EXTENSIONS and str(f.relative_to(cropped_root)) not in wanted:
            f.unlink()
    for dirpath, dirnames, filenames in os.walk(cropped_root, topdown=False):   # 비게 된 폴더 정리
        if Path(dirpath) != cropped_root and not os.listdir(dirpath):
            os.rmdir(dirpath)

    results: list[dict | None] = [None] * total
    new_cache: dict = {}
    lock = threading.Lock()
    processed_count = 0
    reused = 0

    def process_one(i: int, src: Path):
        nonlocal processed_count, reused
        if cancel_check and cancel_check(): return

        rel = str(src.relative_to(merged_root))
        dst = cropped_root / rel
        lot_id, kind = parse_lot_kind(src.stem)
        rec = None
        try:
            st = src.stat()
            sig = [st.st_size, st.st_mtime_ns]
            old = cache.get(rel)
            if old and old.get("sig") == sig and old.get("dst_size") == (dst.stat().st_size if dst.exists() else -1):
                rec = {"lot_id": lot_id, "kind": kind, "src": src, "dst": dst,
                       "bbox": tuple(old["bbox"]), "status": old["status"]}
                with lock:
                    reused += 1
            else:
                dst.parent.mkdir(parents=True, exist_ok=True)
                with Image.open(src) as img:
                    bbox = find_non_black_bbox(img, threshold=threshold)
                    if bbox is None:
                        status = "NO_OBJECT_DETECTED"
                        used_bbox = (0, 0, img.width, img.height)
                        cropped = img.copy()
                    else:
                        status = "OK"
                        x_min, y_min, x_max, y_max = bbox
                        used_bbox = (max(0, x_min-padding), max(0, y_min-padding), min(img.width, x_max+padding), min(img.height, y_max+padding))
                        cropped = img.crop(used_bbox)

                    if dst.suffix.lower() in (".jpg", ".jpeg") and cropped.mode not in ("RGB", "L"):
                        cropped = cropped.convert("RGB")
                    cropped.save(dst)
                rec = {"lot_id": lot_id, "kind": kind, "src": src, "dst": dst, "bbox": used_bbox, "status": status}
            if thumb_width_px and kind in ("BU", "WU"):
                rec["thumb"] = resized_png_bytes(dst, thumb_width_px)
            with lock:
                new_cache[rel] = {"sig": sig, "bbox": list(rec["bbox"]), "status": rec["status"],
                                  "dst_size": dst.stat().st_size}
        except Exception as exc:
            rec = {"lot_id": lot_id, "kind": kind, "src": src, "dst": None, "bbox": None, "status": f"ERROR: {exc}"}
        results[i] = rec

        with lock:
            processed_count += 1
            if processed_count == 1 or processed_count % 20 == 0 or processed_count == total:
                print_progress("  크롭 진행", processed_count, total, done=(processed_count == total))

    max_workers = min(32, (os.cpu_count() or 4) + 4)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        list(executor.map(process_one, range(total), image_files))

    _save_crop_cache(cropped_root, threshold, padding, new_cache)
    if reused:
        print(f"  이전 결과 재사용 {reused:,}개 · 새로 자름 {total - reused:,}개")
    return [r for r in results if r is not None]

def write_excel(records, excel_path: Path, latest_measurements=None, image_width_px: int = THUMB_WIDTH_PX, cancel_check=None,
                extra=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "결과"
    ws.append(["LotID", "판정", "BU data 수치화", "BU Image", "WU data", "WU Image"])

    grouped = {}
    thumbs = {str(rec["dst"]): rec.get("thumb") for rec in records if rec.get("dst") is not None}
    for rec in records:
        lot_id, kind = rec["lot_id"], rec["kind"]
        grouped.setdefault(lot_id, {"BU": None, "WU": None})
        if kind in ("BU", "WU") and rec["dst"] and not grouped[lot_id][kind]:
            grouped[lot_id][kind] = rec["dst"]

    row_idx = 2
    for lot_id in sorted(grouped.keys()):
        ensure_not_cancelled(cancel_check)
        ws.cell(row=row_idx, column=1, value=lot_id)
        m = find_measurement(lot_id, latest_measurements or {})
        ws.cell(row=row_idx, column=2, value=m.get("judge", ""))
        ws.cell(row=row_idx, column=3, value=excel_measurement_value(m.get("black_uniformity", "")))
        ws.cell(row=row_idx, column=5, value=excel_measurement_value(m.get("white_uniformity", "")))

        max_h = 0
        for col, kind in [("D", "BU"), ("F", "WU")]:
            img_path = grouped[lot_id][kind]
            if img_path and Path(img_path).exists():
                pre = thumbs.get(str(img_path)) if image_width_px == THUMB_WIDTH_PX else None
                xl_img = xl_image_from_bytes(pre) if pre else get_resized_xl_image(Path(img_path), image_width_px)
                if xl_img:
                    ws.add_image(xl_img, f"{col}{row_idx}")
                    max_h = max(max_h, xl_img.height)

        ws.row_dimensions[row_idx].height = max(25, int(max_h * 0.75))
        row_idx += 1

    # 조건부 서식
    last_row = max(2, row_idx - 1)
    green_fill = PatternFill("solid", fgColor="D1FAE5")
    red_fill = PatternFill("solid", fgColor="FEE2E2")
    yellow_fill = PatternFill("solid", fgColor="FEF3C7")
    
    ws.conditional_formatting.add(f"B2:B{last_row}", FormulaRule(formula=['B2="OK"'], fill=green_fill))
    ws.conditional_formatting.add(f"B2:B{last_row}", FormulaRule(formula=['B2="NG"'], fill=red_fill))
    ws.conditional_formatting.add(f"C2:C{last_row}", CellIsRule(operator="lessThan", formula=[str(BU_SPEC_MIN)], fill=yellow_fill))
    ws.conditional_formatting.add(f"E2:E{last_row}", CellIsRule(operator="lessThan", formula=[str(WU_SPEC_MIN)], fill=yellow_fill))

    # 컬럼 너비
    for col, width in zip("ABCDEF", [26, 12, 14, 36, 14, 36]):
        ws.column_dimensions[col].width = width

    if extra is not None:                      # v0.6: 공정능력 시트 (실패해도 결과 시트는 저장)
        try:
            extra(wb)
        except Exception as exc:
            print(f"  공정능력 시트를 만들지 못했습니다: {exc}")

    print("\n메인 엑셀 저장")
    wb.save(excel_path)

def capability_sheet_writer(latest_measurements: dict, records: list):
    """(4번 ②안) 판정 = 이번 실행에서 고른 측정 CSV 전체(LotID별 최신), 이미지 정리본 LotID는 참고로 표시"""
    def add(wb):
        try:                                    # 무거운 모듈은 시트를 만들 때만 불러옴 (실행 시작 속도 유지)
            import pandas as pd
            import capability_sheet
            from capability import is_constant
        except ImportError as exc:
            print(f"  공정능력 시트 생략 (필요한 모듈 없음: {exc})")
            return
        if not latest_measurements:
            print("  공정능력 시트 생략 (측정 CSV 없음)")
            return
        rows = []
        for key, m in latest_measurements.items():
            bu, wu = excel_measurement_value(m.get("black_uniformity")), excel_measurement_value(m.get("white_uniformity"))
            rows.append({"key": key, "LotID": m.get("lot_id", key), "time": m.get("time_raw"), "model": m.get("model_name", ""),
                         "BU": bu if isinstance(bu, float) else np.nan, "WU": wu if isinstance(wu, float) else np.nan})
        full = pd.DataFrame(rows)
        image_lots = {rec["lot_id"] for rec in records if rec.get("dst")}
        sub_keys = set()
        for lot in image_lots:
            m = find_measurement(lot, latest_measurements)
            if m:
                sub_keys.add(normalize_lot_id(m["lot_id"]))
        sub_mask = full["key"].isin(sub_keys)
        bu_vals = full["BU"].dropna()
        ld = bool(len(bu_vals) > 1 and is_constant(bu_vals))
        models = [str(x) for x in full["model"] if str(x or "").strip()]
        model = max(set(models), key=models.count) if models else ""
        if not model:
            codes = [re.search(r"V[A-Z]{2}\d{3}[A-Z]{4}\d{2}", str(x)) for x in full["LotID"]]
            codes = [c.group(0) for c in codes if c]
            model = max(set(codes), key=codes.count) if codes else ""
        files = {m.get("source_file") for m in latest_measurements.values()}
        caps = capability_sheet.write_compare_sheet(
            wb, full[["LotID", "WU", "BU", "time"]], sub_mask, ld,
            {"model": model, "csv_files": len(files), "version": f"Tool B {TOOL_VERSION}"})
        print(f"\n[공정능력] 측정 전체 n={len(full)} · 이미지 정리본 n={int(sub_mask.sum())}" + (" · LD 모델(BU 제외)" if ld else ""))
        for name in ("WU", "BU"):
            a, b = caps[name]
            if a.ld:
                print(f"  {name}: N/A (LD 모델)")
            elif a.n >= 2:
                txt = f"  {name}: 전체 Cp {a.cp:.2f} / Cpk {a.cpk:.2f} [{a.verdict}]"
                if b.n >= 2:
                    txt += f" · 정리본 Cp {b.cp:.2f} / Cpk {b.cpk:.2f}"
                print(txt)
    return add


def run_pipeline(integrated_root: Path, data_root: Path, threshold: int, padding: int, cancel_check=None,
                 make_capability: bool = True) -> dict:
    cropped_root = integrated_root.parent / f"{integrated_root.name}_LotID_latest_v1_cropped_v1"
    excel_path = cropped_root / "crop_report.xlsx"

    ensure_not_cancelled(cancel_check)
    latest_folders, _ = collect_latest_lotid_folders(integrated_root, cancel_check)
    if not latest_folders: return {}

    merged_root = integrated_root.parent / f"{integrated_root.name}_LotID_latest_v1"
    copy_latest_folders(latest_folders, merged_root, cancel_check)
    
    latest_m, _ = collect_latest_measurements(data_root, cancel_check)
    crop_records = crop_images(merged_root, cropped_root, threshold, padding, cancel_check)
    
    extra = capability_sheet_writer(latest_m, crop_records) if make_capability else None
    write_excel(crop_records, excel_path, latest_m, cancel_check=cancel_check, extra=extra)

    print(f"\n--- 최종 결과 ({TOOL_VERSION}) ---")
    print(f"완료! (성공: {sum(1 for r in crop_records if r['status']=='OK')})")

    return {"excel_path": excel_path}

if __name__ == "__main__":
    print(f"\n--- BU Organize One Click {TOOL_VERSION} ---")
    ir = Path(input("1) 이미지 통합 폴더: ").strip())
    dr = Path(input("2) 측정 데이터 폴더: ").strip())
    th = int(input("3) 임계값 [12]: ") or 12)
    pd = int(input("4) 패딩 [20]: ") or 20)
    run_pipeline(ir, dr, th, pd)
