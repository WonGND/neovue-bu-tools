"""BU/WU 공정능력 계산 (Tool A·Tool C 공용).

- 규격: WU 80~100, BU 50~100 (상한 100은 균일도의 물리적 상한)
- Cp·Cpk: 군내 σ = 이동범위 평균 / 1.128 (I-MR, 측정 순서 기준)
- Ppk: 전체 σ (표본 표준편차, 양식 19행 CPK 수식과 같은 값)
- LD 모델: BU 값이 전부 같거나 σ≈0이면 LD로 보고 BU를 계산하지 않는다
scipy 없이 numpy·math만 사용 (exe 용량 절약).
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

import numpy as np

SPECS = {"WU": (80.0, 100.0), "BU": (50.0, 100.0)}
GRADES = ((1.67, "양호", "D1FAE5"), (1.33, "적합", "DBEAFE"), (1.00, "주의", "FEF3C7"), (-math.inf, "부적합", "FEE2E2"))
D2 = 1.128          # 이동범위(n=2) 관리도 상수
MIN_N_WARN = 30
SKEW_WARN = 0.5
INVALID_WU = -100.0  # LMK6가 측정 실패 때 남기는 값 → 분석에서 제외


def is_invalid_wu(values):
    """WU = -100(측정 이상값)인 행 표시. pandas Series를 넣으면 같은 인덱스의 bool Series"""
    try:
        import pandas as pd
        if isinstance(values, pd.Series):
            return pd.to_numeric(values, errors="coerce").sub(INVALID_WU).abs().lt(1e-9)
    except ImportError:                     # pragma: no cover
        pass
    x = np.asarray([float(v) if str(v).strip() not in ("", "None", "nan") else np.nan for v in values], dtype=float)
    return np.abs(x - INVALID_WU) < 1e-9


def grade(v: float | None) -> tuple[str, str]:
    """(판정, 엑셀 채우기 색) — Cp·Cpk 공통 기준."""
    if v is None or not np.isfinite(v):
        return "N/A", "E5E7EB"
    for th, name, color in GRADES:
        if v >= th:
            return name, color
    return "부적합", "FEE2E2"


def to_numeric(values) -> np.ndarray:
    out = []
    for v in values:
        try:
            f = float(str(v).strip())
        except (TypeError, ValueError):
            continue
        if np.isfinite(f):
            out.append(f)
    return np.asarray(out, dtype=float)


def is_constant(values, tol: float = 1e-6) -> bool:
    """LD 모델 감지: 값이 전부 같거나 σ≈0."""
    x = to_numeric(values)
    if len(x) < 2:
        return False
    return bool(np.ptp(x) <= tol or np.std(x) <= tol)


def _norm_cdf(z: float) -> float:
    return 0.5 * (1 + math.erf(z / math.sqrt(2)))


@dataclass
class Capability:
    name: str
    lsl: float
    usl: float
    n: int = 0
    n_excluded: int = 0
    mean: float = float("nan")
    sigma_within: float = float("nan")
    sigma_overall: float = float("nan")
    cp: float = float("nan")
    cpk: float = float("nan")
    pp: float = float("nan")
    ppk: float = float("nan")
    ng: int = 0
    ng_pct: float = float("nan")
    expected_ng_pct: float = float("nan")
    skew: float = float("nan")
    ld: bool = False
    warnings: list[str] = field(default_factory=list)

    @property
    def normality(self) -> str:
        if not np.isfinite(self.skew):
            return "-"
        return "정규" if abs(self.skew) < SKEW_WARN else f"주의(왜도 {self.skew:+.2f})"

    @property
    def verdict(self) -> str:
        return "N/A (LD 모델)" if self.ld else grade(self.cpk)[0]


def capability(values, name: str, lsl: float | None = None, usl: float | None = None, ld: bool = False,
               raw_count: int | None = None) -> Capability:
    """values는 측정 순서대로 정렬돼 있어야 한다(군내 σ가 순서에 의존)."""
    d_lsl, d_usl = SPECS.get(name, (lsl, usl))
    lsl = d_lsl if lsl is None else lsl
    usl = d_usl if usl is None else usl
    x = to_numeric(values)
    total = raw_count if raw_count is not None else len(list(values))
    c = Capability(name=name, lsl=lsl, usl=usl, n=len(x), n_excluded=max(0, total - len(x)), ld=ld)
    if ld:
        c.n_excluded, c.n = total, 0
        c.warnings.append("LD 모델 — BU 미측정으로 계산 제외")
        return c
    if len(x) < 2:
        c.warnings.append("데이터 부족 (n<2)")
        return c
    m = float(x.mean())
    so = float(x.std(ddof=1))
    mr = np.abs(np.diff(x))
    sw = float(mr.mean() / D2) if len(mr) else float("nan")
    c.mean, c.sigma_overall, c.sigma_within = m, so, sw
    if sw > 0:
        c.cp = (usl - lsl) / (6 * sw)
        c.cpk = min((m - lsl) / (3 * sw), (usl - m) / (3 * sw))
    if so > 0:
        c.pp = (usl - lsl) / (6 * so)
        c.ppk = min((m - lsl) / (3 * so), (usl - m) / (3 * so))
        c.expected_ng_pct = 100 * (_norm_cdf((lsl - m) / so) + 1 - _norm_cdf((usl - m) / so))
        c.skew = float(np.mean(((x - m) / x.std(ddof=0)) ** 3)) if x.std(ddof=0) > 0 else 0.0
    c.ng = int(((x < lsl) | (x > usl)).sum())
    c.ng_pct = 100 * c.ng / len(x)
    if len(x) < MIN_N_WARN:
        c.warnings.append(f"n<{MIN_N_WARN} 참고용")
    if np.isfinite(c.skew) and abs(c.skew) >= SKEW_WARN:
        c.warnings.append(f"비정규 분포(왜도 {c.skew:+.2f}) → 실측 NG율 함께 확인")
    if so == 0:
        c.warnings.append("값이 모두 같음(σ=0)")
    return c


def summary_line(c: Capability) -> str:
    if c.ld:
        return f"{c.name}   N/A (LD 모델 — 계산 제외)"
    if not np.isfinite(c.cpk):
        return f"{c.name}   계산 불가 ({', '.join(c.warnings)})"
    return (f"{c.name}   Cp {c.cp:.2f} / Cpk {c.cpk:.2f}   [{c.verdict}]   "
            f"실측 NG {c.ng}ea ({c.ng_pct:.1f}%)")
