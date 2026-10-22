"""누적 DB (각자 PC의 엑셀 파일 1개, 서버 없음).

시트
- data   : 1행 = LotID × 측정일시. 이벤트는 저장하지 않고(수동 지정만 저장) 리포트 때 기간 표로 붙인다.
- events : 이벤트 관리 표 (이벤트명·모델·시작일·종료일·메모)
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

DATA_COLS = ["LotID", "모델명", "모델코드", "측정일시", "측정일", "생산일", "채널",
             "Red_X", "Red_Y", "Green_X", "Green_Y", "Blue_X", "Blue_Y",
             "White_x", "White_y", "White_LV", "Black_LV", "CR", "WU", "BU", "판정", "EEPROM_59",
             "LD", "이벤트(수동)", "출처", "원본파일", "Tool", "적재일시"]
EVENT_COLS = ["이벤트명", "모델", "시작일", "종료일", "메모"]
MODEL_RE = re.compile(r"V[A-Z]{2}\d{3}[A-Z]{4}\d{2}")


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
    events["모델"] = events["모델"].fillna("").astype(str)
    events["메모"] = events["메모"].fillna("").astype(str)
    data["이벤트(수동)"] = data["이벤트(수동)"].fillna("").astype(str)
    return data[DATA_COLS], events[EVENT_COLS]


def save(path, data: pd.DataFrame, events: pd.DataFrame) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".tmp.xlsx")
    out = data.copy()
    out["측정일시"] = pd.to_datetime(out["측정일시"], errors="coerce")
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as w:
            out[DATA_COLS].to_excel(w, sheet_name="data", index=False)
            events[EVENT_COLS].to_excel(w, sheet_name="events", index=False)
            pd.DataFrame({"항목": ["설명", "갱신"], "값": [
                "BU/WU 누적 DB — data: LotID×측정일시, events: 이벤트 기간 표. 직접 수정 시 열 이름은 바꾸지 마세요.",
                f"{datetime.now():%Y-%m-%d %H:%M:%S}"]}).to_excel(w, sheet_name="info", index=False)
        tmp.replace(p)
    except PermissionError as exc:
        if tmp.exists():
            tmp.unlink(missing_ok=True)
        raise DBLockedError(f"DB 파일이 열려 있어 저장하지 못했습니다. 엑셀에서 닫고 다시 시도하세요: {p}") from exc


def append(path, rows: pd.DataFrame) -> tuple[int, int]:
    """rows를 DB에 추가. 같은 (LotID, 측정일시)는 건너뜀. 반환 (추가, 중복 건너뜀)"""
    data, events = load(path)
    rows = rows.copy()
    for c in DATA_COLS:
        if c not in rows:
            rows[c] = None
    rows["LotID"] = rows["LotID"].map(normalize_lot_id)
    rows["측정일시"] = pd.to_datetime(rows["측정일시"], errors="coerce")
    rows["측정일"] = rows["측정일시"].dt.date
    rows["이벤트(수동)"] = rows["이벤트(수동)"].fillna("").astype(str)
    rows["적재일시"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    key = lambda df: df["LotID"].astype(str) + "|" + pd.to_datetime(df["측정일시"]).dt.strftime("%Y%m%d%H%M%S").fillna("")
    existing = set(key(data)) if len(data) else set()
    rows = rows.drop_duplicates(subset=["LotID", "측정일시"], keep="last")
    mask = ~key(rows).isin(existing)
    new = rows[mask][DATA_COLS]
    if len(new):
        data = pd.concat([data, new], ignore_index=True) if len(data) else new.reset_index(drop=True)
        save(path, data, events)
    return int(mask.sum()), int((~mask).sum())


def save_events(path, events: pd.DataFrame) -> None:
    data, _ = load(path)
    save(path, data, events)


def _model_match(event_model: str, names: list[str]) -> bool:
    em = (event_model or "").strip().upper()
    if not em:
        return True
    return any(em in str(n).upper() for n in names if n)


def events_for(d: date, models: list[str], events: pd.DataFrame) -> list[str]:
    """측정일 d가 기간 안에 들고 모델이 맞는 이벤트 이름 목록"""
    if d is None or pd.isna(d):
        return []
    out = []
    for _, e in events.iterrows():
        if pd.isna(e["시작일"]) or pd.isna(e["종료일"]):
            continue
        if e["시작일"] <= d <= e["종료일"] and _model_match(e["모델"], models):
            out.append(str(e["이벤트명"]))
    return out


def assign_events(data: pd.DataFrame, events: pd.DataFrame) -> pd.DataFrame:
    """'이벤트'·'매핑상태' 열 추가. 수동 지정이 있으면 우선."""
    df = data.copy()
    ev, st = [], []
    for _, r in df.iterrows():
        manual = str(r.get("이벤트(수동)") or "").strip()
        if manual:
            ev.append(manual)
            st.append("수동 지정")
            continue
        hits = events_for(r["측정일"], [r.get("모델명"), r.get("모델코드")], events)
        if len(hits) == 1:
            ev.append(hits[0])
            st.append("정상")
        elif not hits:
            ev.append("")
            st.append("기간 밖")
        else:
            ev.append("")
            st.append("겹침: " + ", ".join(hits))
    df["이벤트"] = ev
    df["매핑상태"] = st
    return df
