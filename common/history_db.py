"""누적 DB (각자 PC의 엑셀 파일 1개, 서버 없음).

시트
- data   : 1행 = LotID × 측정일시. '이벤트(수동)' 열 = 그 행이 속한 이벤트 (날짜 묶기·파일 추가 때 지정, 비면 미지정)
- events : 이벤트 목록 (이벤트명·모델·시작일·종료일·메모·구분)
           구분 '묶음' = 행에 붙은 이벤트로 관리(v0.2~, 기간은 데이터에서 자동 계산)
           구분 '기간' = v0.1에서 기간으로 등록한 이벤트 → 트렌드 툴을 열 때 '묶음'으로 자동 전환
기본 위치: 문서\\BU_Tools\\BU_History_DB.xlsx (Tool A·트렌드 리포트가 같이 씀)
"""
from __future__ import annotations

import os
import re
import shutil
from datetime import date, datetime
from pathlib import Path

import pandas as pd

DATA_COLS = ["LotID", "모델명", "모델코드", "측정일시", "측정일", "생산일", "채널",
             "Red_X", "Red_Y", "Green_X", "Green_Y", "Blue_X", "Blue_Y",
             "White_x", "White_y", "White_LV", "Black_LV", "CR", "WU", "BU", "판정", "EEPROM_59",
             "LD", "이벤트(수동)", "출처", "원본파일", "Tool", "적재일시"]
EVENT_COLS = ["이벤트명", "모델", "시작일", "종료일", "메모", "구분"]
LABEL = "이벤트(수동)"            # 이벤트 이름이 들어가는 열 (열 이름은 v1.5 DB와 호환되게 유지)
UPLOAD_SOURCE = "과거 파일 업로드"
DB_NAME = "BU_History_DB.xlsx"
MODEL_RE = re.compile(r"V[A-Z]{2}\d{3}[A-Z]{4}\d{2}")


def documents_dir() -> Path:
    """Windows '문서' 폴더 (OneDrive로 옮긴 경우 포함). 그 밖의 OS는 ~/Documents"""
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            if ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf) == 0 and buf.value:   # 5 = CSIDL_PERSONAL
                return Path(buf.value)
        except Exception:
            pass
    return Path.home() / "Documents"


def default_db_path() -> Path:
    return documents_dir() / "BU_Tools" / DB_NAME


class DBLockedError(RuntimeError):
    pass


def normalize_lot_id(value) -> str:
    text = "" if value is None else str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return re.sub(r"[\s_-]+", "", text).upper()


def model_code(lot_id: str) -> str:
    m = MODEL_RE.search(str(lot_id))
    return m.group(0) if m else ""


def prod_date(lot_id: str):
    """LotID 안 모델코드 뒤 6자리(YYMMDD) = 생산·투입일. 없거나 날짜가 아니면 None"""
    s = str(lot_id)
    m = MODEL_RE.search(s)
    if not m:
        return None
    tail = s[m.end():m.end() + 6]
    try:
        return datetime.strptime(tail, "%y%m%d").date()
    except ValueError:
        return None


def _empty_data() -> pd.DataFrame:
    return pd.DataFrame(columns=DATA_COLS)


def _empty_events() -> pd.DataFrame:
    return pd.DataFrame(columns=EVENT_COLS)


def load(path) -> tuple[pd.DataFrame, pd.DataFrame]:
    p = Path(path)
    if not p.exists():
        return _empty_data(), _empty_events()
    try:
        sheets = pd.read_excel(p, sheet_name=None, dtype={"LotID": str, "EEPROM_59": str})
    except PermissionError as exc:
        raise DBLockedError(f"DB 파일을 열 수 없습니다(다른 프로그램에서 사용 중): {p}") from exc
    data = sheets.get("data", _empty_data())
    events = sheets.get("events", _empty_events())
    old_events = "구분" not in events.columns
    for c in DATA_COLS:
        if c not in data:
            data[c] = None
    for c in EVENT_COLS:
        if c not in events:
            events[c] = None
    data["측정일시"] = pd.to_datetime(data["측정일시"], errors="coerce")
    data["측정일"] = pd.to_datetime(data["측정일"], errors="coerce").dt.date
    for c in ("시작일", "종료일"):
        events[c] = pd.to_datetime(events[c], errors="coerce").dt.date
    events["이벤트명"] = events["이벤트명"].fillna("").astype(str)
    events["모델"] = events["모델"].fillna("").astype(str)
    events["메모"] = events["메모"].fillna("").astype(str)
    # 구분이 없거나 빈 이벤트 = v0.1 기간 이벤트
    events["구분"] = "기간" if old_events else events["구분"].fillna("").astype(str).replace("", "기간")
    events = events[events["이벤트명"].str.strip() != ""]
    data[LABEL] = data[LABEL].fillna("").astype(str)
    return data[DATA_COLS], events[EVENT_COLS].reset_index(drop=True)


def save(path, data: pd.DataFrame, events: pd.DataFrame) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".tmp.xlsx")
    out = data.copy()
    out["측정일시"] = pd.to_datetime(out["측정일시"], errors="coerce")
    ev = events.copy()
    for c in EVENT_COLS:
        if c not in ev:
            ev[c] = ""
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as w:
            out[DATA_COLS].to_excel(w, sheet_name="data", index=False)
            ev[EVENT_COLS].to_excel(w, sheet_name="events", index=False)
            pd.DataFrame({"항목": ["설명", "갱신"], "값": [
                "BU/WU 누적 DB — data: LotID×측정일시('이벤트(수동)' = 소속 이벤트), events: 이벤트 목록. "
                "직접 수정 시 열 이름은 바꾸지 마세요.",
                f"{datetime.now():%Y-%m-%d %H:%M:%S}"]}).to_excel(w, sheet_name="info", index=False)
        tmp.replace(p)
    except PermissionError as exc:
        if tmp.exists():
            tmp.unlink(missing_ok=True)
        raise DBLockedError(f"DB 파일이 열려 있어 저장하지 못했습니다. 엑셀에서 닫고 다시 시도하세요: {p}") from exc


# ------------------------------------------------------------------ 적재 · 중복 방지
def _time_key(df: pd.DataFrame) -> pd.Series:
    return df["LotID"].astype(str) + "|" + pd.to_datetime(df["측정일시"], errors="coerce").dt.strftime("%Y%m%d%H%M%S").fillna("")


def _value_key(df: pd.DataFrame) -> pd.Series:
    """같은 측정인지: LotID + WU + BU (소수 셋째 자리). WU·BU가 비면 키 없음(빈 문자열)"""
    wu = pd.to_numeric(df["WU"], errors="coerce").round(3)
    bu = pd.to_numeric(df["BU"], errors="coerce").round(3)
    key = df["LotID"].astype(str) + "|" + wu.map("{:.3f}".format) + "|" + bu.map("{:.3f}".format)
    return key.where(wu.notna() & bu.notna(), "")


def _is_tool_a(src) -> bool:
    return str(src or "").startswith("Tool A")


def _prepare(rows: pd.DataFrame) -> pd.DataFrame:
    rows = rows.copy()
    for c in DATA_COLS:
        if c not in rows:
            rows[c] = None
    rows["LotID"] = rows["LotID"].map(normalize_lot_id)
    rows["측정일시"] = pd.to_datetime(rows["측정일시"], errors="coerce")
    rows["측정일"] = rows["측정일시"].dt.date
    rows[LABEL] = rows[LABEL].fillna("").astype(str)
    rows["적재일시"] = rows["적재일시"].where(rows["적재일시"].notna(), datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    return rows.drop_duplicates(subset=["LotID", "측정일시"], keep="last")[DATA_COLS]


def merge_rows(data: pd.DataFrame, rows: pd.DataFrame, label_dups: bool = False) -> tuple[pd.DataFrame, dict]:
    """rows를 data에 합친다 (저장은 하지 않음).
    같은 측정 = (LotID, 측정일시) 같거나, (LotID, WU, BU) 같음 → 새로 넣지 않음.
    - 기존 행이 과거 파일 업로드분이고 새 행이 Tool A(측정 시각 있음)면 기존 행의 시각·출처를 Tool A 것으로 바꿈
    - label_dups: 새 행에 이벤트가 있고 기존 행이 미지정이면 기존 행에 그 이벤트를 붙임"""
    rows = _prepare(rows)
    data = data.copy()
    stats = {"added": 0, "skipped": 0, "relabeled": 0, "upgraded": 0}
    if not len(rows):
        return data, stats
    tkey_old = set(_time_key(data)) if len(data) else set()
    vkey_old = {}
    if len(data):
        for i, k in zip(data.index, _value_key(data)):
            if k:
                vkey_old.setdefault(k, i)
    tkeys, vkeys = _time_key(rows), _value_key(rows)
    new_idx = []
    for (i, r), tk_, vk in zip(rows.iterrows(), tkeys, vkeys):
        hit = None
        if vk and vk in vkey_old:
            hit = vkey_old[vk]
        elif tk_ in tkey_old:
            hit = -1
        if hit is None:
            new_idx.append(i)
            if vk:
                vkey_old[vk] = -2            # 같은 묶음 안 중복도 막음 (-2 = 이번에 새로 넣는 행)
            tkey_old.add(tk_)
            continue
        stats["skipped"] += 1
        if hit < 0 or hit not in data.index:
            continue
        old = data.loc[hit]
        if _is_tool_a(r["출처"]) and not _is_tool_a(old["출처"]) and pd.notna(r["측정일시"]):
            for c in ("측정일시", "측정일", "채널", "출처", "원본파일", "Tool"):
                data.at[hit, c] = r[c]
            stats["upgraded"] += 1
        if label_dups and r[LABEL] and not str(old[LABEL]).strip():
            data.at[hit, LABEL] = r[LABEL]
            stats["relabeled"] += 1
    new = rows.loc[new_idx]
    stats["added"] = len(new)
    if len(new):
        data = pd.concat([data, new], ignore_index=True) if len(data) else new.reset_index(drop=True)
    return data, stats


def append_rows(path, rows: pd.DataFrame, label_dups: bool = False) -> dict:
    """rows를 DB에 추가하고 저장. 반환 {'added', 'skipped', 'relabeled', 'upgraded'}"""
    data, events = load(path)
    data, stats = merge_rows(data, rows, label_dups)
    if stats["added"] or stats["relabeled"] or stats["upgraded"]:
        events = ensure_events(data, events)
        save(path, data, events)
    return stats


def append(path, rows: pd.DataFrame) -> tuple[int, int]:
    """(Tool A용) rows를 DB에 추가. 같은 측정은 건너뜀. 반환 (추가, 중복 건너뜀)"""
    s = append_rows(path, rows)
    return s["added"], s["skipped"]


# ------------------------------------------------------------------ 공용 위치 · 옛 DB 합치기
def merge_db(src, dst) -> dict:
    """옛 DB(src)를 공용 DB(dst)로 합친다. src 파일은 그대로 둔다."""
    s_data, s_events = load(src)
    d_data, d_events = load(dst)
    order = s_data["출처"].map(lambda x: 0 if _is_tool_a(x) else 1)        # Tool A(시각 있음) 먼저
    data, stats = merge_rows(d_data, s_data.assign(_o=order).sort_values("_o").drop(columns="_o"), label_dups=True)
    names = set(d_events["이벤트명"])
    add = s_events[~s_events["이벤트명"].isin(names)]
    events = pd.concat([d_events, add], ignore_index=True) if len(add) else d_events
    stats["events"] = len(add)
    if len(s_data) or len(add):
        save(dst, data, ensure_events(data, events))
    return stats


def adopt_shared_db(saved_path: str | None, legacy_path, merged: list) -> tuple[str, str]:
    """설정에 저장된 DB 경로를 정한다. 반환 (쓸 경로, 안내 문구)
    - 설정이 비었거나 옛 기본값(exe 옆)이면 → 공용 경로(문서\\BU_Tools)
    - exe 옆 옛 DB가 있고 아직 안 합쳤으면 공용 DB로 합침 (merged 목록에 기록, 원본 파일은 그대로)
    - 사용자가 직접 고른 다른 경로면 그대로"""
    shared = default_db_path()
    legacy = Path(legacy_path)
    saved = Path(saved_path) if saved_path else None
    msg = ""
    use_shared = saved is None or saved.resolve() == legacy.resolve()
    if legacy.exists() and str(legacy.resolve()) not in merged and legacy.resolve() != shared.resolve():
        if use_shared:
            st = merge_db(legacy, shared)
            merged.append(str(legacy.resolve()))
            msg = (f"누적 DB를 공용 위치로 옮겼습니다: {shared}\n(이전 파일 {legacy.name}의 {st['added']:,}행 합침, "
                   f"이미 있던 {st['skipped']:,}행 건너뜀 · 원본 파일은 그대로 둠)")
    return (str(shared) if use_shared else str(saved)), msg


# ------------------------------------------------------------------ 이벤트
def save_events(path, events: pd.DataFrame) -> None:
    data, _ = load(path)
    save(path, data, events)


def _model_match(event_model: str, names: list[str]) -> bool:
    em = (event_model or "").strip().upper()
    if not em:
        return True
    return any(em in str(n).upper() for n in names if n)


def events_for(d: date, models: list[str], events: pd.DataFrame) -> list[str]:
    """(v0.1 '기간' 이벤트용) 측정일 d가 기간 안에 들고 모델이 맞는 이벤트 이름 목록"""
    if d is None or pd.isna(d):
        return []
    out = []
    for _, e in events.iterrows():
        if str(e.get("구분", "기간")) == "묶음":
            continue
        if pd.isna(e["시작일"]) or pd.isna(e["종료일"]):
            continue
        if e["시작일"] <= d <= e["종료일"] and _model_match(e["모델"], models):
            out.append(str(e["이벤트명"]))
    return out


def assign_events(data: pd.DataFrame, events: pd.DataFrame) -> pd.DataFrame:
    """'이벤트'·'매핑상태' 열 추가. 행에 붙은 이벤트가 우선, 없으면 (전환 전) 기간 이벤트로."""
    df = data.copy()
    ev, st = [], []
    if len(events) and "구분" in events:
        period = events[events["구분"] != "묶음"]
    else:
        period = events
    for _, r in df.iterrows():
        label = str(r.get(LABEL) or "").strip()
        if label:
            ev.append(label)
            st.append("지정")
            continue
        hits = events_for(r["측정일"], [r.get("모델명"), r.get("모델코드")], period) if len(period) else []
        if len(hits) == 1:
            ev.append(hits[0])
            st.append("기간")
        else:
            ev.append("")
            st.append("미지정" if not hits else "겹침: " + ", ".join(hits))
    df["이벤트"] = ev
    df["매핑상태"] = st
    return df


def migrate_period_events(data: pd.DataFrame, events: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    """v0.1 '기간' 이벤트 → 행에 이벤트를 붙이는 '묶음' 방식으로 전환. 반환 (data, events, 붙인 행 수)"""
    if not len(events) or (events["구분"] == "묶음").all():
        return data, events, 0
    data = data.copy()
    m = assign_events(data, events)
    hit = (data[LABEL].str.strip() == "") & (m["매핑상태"] == "기간")
    data.loc[hit, LABEL] = m.loc[hit, "이벤트"]
    events = events.copy()
    events["구분"] = "묶음"
    return data, ensure_events(data, events), int(hit.sum())


def ensure_events(data: pd.DataFrame, events: pd.DataFrame, memo: dict | None = None) -> pd.DataFrame:
    """행에 붙은 이벤트가 목록에 다 있게 하고, 기간·모델은 데이터에서 다시 계산"""
    events = events.copy() if len(events) else _empty_events()
    labels = data[LABEL].fillna("").astype(str).str.strip() if len(data) else pd.Series(dtype=str)
    for name in sorted(set(labels) - {""} - set(events["이벤트명"].astype(str))):
        events = pd.concat([events, pd.DataFrame([{"이벤트명": name, "모델": "", "메모": "", "구분": "묶음"}])],
                           ignore_index=True)
    for i, e in events.iterrows():
        name = str(e["이벤트명"])
        if memo and name in memo:
            events.at[i, "메모"] = memo[name]
        if str(e.get("구분") or "") != "묶음":
            continue
        g = data[labels == name] if len(data) else data
        if len(g):
            d = pd.to_datetime(g["측정일"], errors="coerce").dropna()
            events.at[i, "시작일"] = d.min().date() if len(d) else None
            events.at[i, "종료일"] = d.max().date() if len(d) else None
            events.at[i, "모델"] = ", ".join(sorted(set(g["모델명"].dropna().astype(str)) - {""}))
    events["구분"] = events["구분"].fillna("").astype(str).replace("", "기간")
    return events[EVENT_COLS].reset_index(drop=True)


def day_key(df: pd.DataFrame) -> pd.Series:
    """표의 '날짜 행' 키: 측정일|모델명"""
    return pd.to_datetime(df["측정일"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("날짜없음") + "|" + \
        df["모델명"].fillna("").astype(str)


def _pair_mask(data: pd.DataFrame, pairs) -> pd.Series:
    """pairs: [(현재 이벤트('' = 미지정), '측정일|모델명'), ...] 에 해당하는 행"""
    want = {f"{lab}||{key}" for lab, key in pairs}
    return (data[LABEL].fillna("").astype(str).str.strip() + "||" + day_key(data)).isin(want)


def set_label(path, pairs, label: str, memo: str | None = None) -> int:
    """고른 날짜 행들의 이벤트를 label로 바꿈 (빈 문자열 = 묶음 풀기 → 미지정). 반환 바뀐 행 수"""
    data, events = load(path)
    hit = _pair_mask(data, pairs)
    data.loc[hit, LABEL] = label
    events = ensure_events(data, events, {label: memo} if (label and memo) else None)
    save(path, data, events)
    return int(hit.sum())


def rename_event(path, old: str, new: str, memo: str) -> None:
    data, events = load(path)
    if new != old and new in set(events["이벤트명"]):
        raise ValueError(f"'{new}' 이벤트가 이미 있습니다.")
    data.loc[data[LABEL].str.strip() == old, LABEL] = new
    events.loc[events["이벤트명"] == old, ["이벤트명", "메모"]] = [new, memo]
    save(path, data, ensure_events(data, events))


def delete_event(path, name: str) -> None:
    """이벤트를 목록에서 지우고 그 행들은 미지정으로 (데이터는 남김)"""
    data, events = load(path)
    data.loc[data[LABEL].str.strip() == name, LABEL] = ""
    save(path, data, events[events["이벤트명"] != name])


def delete_rows(path, pairs) -> int:
    """고른 날짜 행의 데이터를 DB에서 삭제 (지우기 전 같은 폴더에 .backup.xlsx 복사본을 남김)"""
    p = Path(path)
    data, events = load(p)
    hit = _pair_mask(data, pairs)
    if hit.any():
        shutil.copy2(p, p.with_name(p.stem + ".backup.xlsx"))
        data = data[~hit].reset_index(drop=True)
        save(p, data, ensure_events(data, events))
    return int(hit.sum())
