"""그래프 모듈 (그림 방식, D2 A안).

모든 함수는 PNG 바이트를 돌려준다. 엑셀 기본 차트(B안)로 바꿀 때는 이 모듈과 같은
이름의 함수를 가진 다른 모듈로 갈아 끼우면 된다.
"""
from __future__ import annotations

import io
import math

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402

plt.rcParams.update({
    "font.family": ["Malgun Gothic", "Noto Sans CJK KR", "Noto Sans CJK JP", "AppleGothic", "DejaVu Sans"],
    "axes.unicode_minus": False, "font.size": 9,
    "axes.spines.top": False, "axes.spines.right": False,
})
INK, GREY, BLUE, RED, AMBER, SLATE = "#111827", "#9ca3af", "#2563eb", "#dc2626", "#d97706", "#475569"
ORANGE = "#f59e0b"


def _png(fig) -> bytes:
    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=130)
    plt.close(fig)
    return buf.getvalue()


def _norm_pdf(x, m, s):
    return np.exp(-0.5 * ((x - m) / s) ** 2) / (s * math.sqrt(2 * math.pi))


def hist_with_spec(x, name, lsl, usl, cap, width=5.6, height=3.0) -> bytes:
    """분포 + LSL/USL + 평균 + 정규곡선. cap = capability.Capability"""
    x = np.asarray(x, dtype=float)
    fig, ax = plt.subplots(figsize=(width, height))
    lo = min(x.min(), lsl) - 3
    bins = np.linspace(lo, usl, 46)
    ax.hist(x, bins=bins, color=GREY, edgecolor="white", linewidth=0.4)
    if cap.sigma_overall and cap.sigma_overall > 0:
        xs = np.linspace(lo, usl, 300)
        ax.plot(xs, _norm_pdf(xs, cap.mean, cap.sigma_overall) * len(x) * (bins[1] - bins[0]), color=INK, lw=1)
    top = ax.get_ylim()[1]
    for v, lab, ha in ((lsl, f"LSL {lsl:g}", "left"), (usl, f"USL {usl:g}", "right")):
        ax.axvline(v, color=RED, ls="--", lw=1.2)
        ax.text(v, top * 0.97, lab, color=RED, ha=ha, va="top", fontsize=8)
    ax.axvline(cap.mean, color=BLUE, lw=1.2)
    ax.text(cap.mean, top * 0.70, f" 평균 {cap.mean:.2f}", color=BLUE, fontsize=8,
            bbox=dict(facecolor="white", edgecolor="none", pad=1))
    ax.set_title(f"{name} 분포  ·  Cp {cap.cp:.2f}  /  Cpk {cap.cpk:.2f}", loc="left", fontsize=11, fontweight="bold")
    ax.set_xlabel(f"{name} (%)")
    ax.set_ylabel("수량 (ea)")
    return _png(fig)


def ichart(values, name, lsl, group_labels=None, width=11.4, height=2.4, highlight=None, highlight_label="") -> bytes:
    """측정 순서 개별값 + 평균 + LSL. group_labels: 각 점의 그룹(측정일) — 구분선과 라벨 표시.
    highlight: 강조할 점(True/False 목록, 예: 이미지 정리본 LotID) — 주황 점으로 표시"""
    x = np.asarray(values, dtype=float)
    idx = np.arange(len(x))
    fig, ax = plt.subplots(figsize=(width, height))
    ax.plot(idx, x, color=GREY, lw=0.6)
    bad = x < lsl
    ax.scatter(idx[~bad], x[~bad], s=4, color="#64748b")
    if highlight is not None and len(highlight) == len(x):
        hl = np.asarray(highlight, dtype=bool)
        ax.scatter(idx[hl], x[hl], s=12, color=ORANGE, zorder=2, label=highlight_label or f"강조 {int(hl.sum())}ea")
    ax.scatter(idx[bad], x[bad], s=10, color=RED, zorder=3, label=f"LSL 미달 {int(bad.sum())}ea")
    ax.axhline(lsl, color=RED, ls="--", lw=1)
    ax.axhline(x.mean(), color=BLUE, lw=1)
    if group_labels is not None and len(group_labels) == len(x):
        g = list(group_labels)
        starts = [i for i in range(len(g)) if i == 0 or g[i] != g[i - 1]]
        ends = starts[1:] + [len(g)]
        for s, e in zip(starts, ends):
            if s:
                ax.axvline(s - 0.5, color="#cbd5e1", lw=1)
            ax.text((s + e - 1) / 2, ax.get_ylim()[1], str(g[s]), ha="center", va="bottom", fontsize=8, color=SLATE)
    ax.set_title(f"{name} 개별값 추이 (측정 순서, 측정일 구분)", loc="left", fontsize=10, fontweight="bold", pad=14)
    ax.set_ylabel(f"{name} (%)")
    if bad.any() or highlight is not None:
        ax.legend(loc="lower right", fontsize=8, frameon=False, ncol=2)
    return _png(fig)


def hist_overlay(x_all, x_sub, name, lsl, usl, cap_all, cap_sub, width=5.6, height=3.0,
                 all_label="측정 CSV 전체", sub_label="이미지 정리본") -> bytes:
    """전체 분포(회색) 위에 일부(주황)를 같은 구간으로 겹쳐 그림 + 전체 정규곡선 + LSL/USL"""
    x = np.asarray(x_all, dtype=float)
    x = x[~np.isnan(x)]
    sx = np.asarray(x_sub, dtype=float)
    sx = sx[~np.isnan(sx)]
    fig, ax = plt.subplots(figsize=(width, height))
    lo = min(x.min(), lsl) - 3
    bins = np.linspace(lo, usl, 46)
    ax.hist(x, bins=bins, color="#cbd5e1", edgecolor="white", linewidth=0.4, label=f"{all_label} n={len(x)}")
    if len(sx):
        ax.hist(sx, bins=bins, color=ORANGE, edgecolor="white", linewidth=0.4, label=f"{sub_label} n={len(sx)}")
    if cap_all.sigma_overall and cap_all.sigma_overall > 0:
        xs = np.linspace(lo, usl, 300)
        ax.plot(xs, _norm_pdf(xs, cap_all.mean, cap_all.sigma_overall) * len(x) * (bins[1] - bins[0]), color=INK, lw=1)
    top = ax.get_ylim()[1]
    for v, lab, ha in ((lsl, f"LSL {lsl:g}", "left"), (usl, f"USL {usl:g}", "right")):
        ax.axvline(v, color=RED, ls="--", lw=1.2)
        ax.text(v, top * 0.97, lab, color=RED, ha=ha, va="top", fontsize=8)
    t = f"{name}  ·  전체 Cp {cap_all.cp:.2f} / Cpk {cap_all.cpk:.2f}"
    if cap_sub is not None and len(sx) >= 2:
        t += f"   ·   정리본 Cp {cap_sub.cp:.2f} / Cpk {cap_sub.cpk:.2f}"
    ax.set_title(t, loc="left", fontsize=9.5, fontweight="bold")
    ax.set_xlabel(f"{name} (%)")
    ax.set_ylabel("수량 (ea)")
    ax.legend(fontsize=7.5, frameon=False, loc="upper left", bbox_to_anchor=(0.0, 0.9))
    return _png(fig)


def boxplot_groups(groups: dict, name, lsl, title, width=6.0, height=3.4, highlight_last=True) -> bytes:
    """그룹(이벤트·측정일)별 박스플롯 + LSL. groups = {라벨: 값 배열}"""
    labels = list(groups)
    data = [np.asarray(groups[k], dtype=float) for k in labels]
    fig, ax = plt.subplots(figsize=(width, height))
    bp = ax.boxplot(data, widths=0.5, patch_artist=True,
                    flierprops=dict(markersize=2, markerfacecolor=GREY, markeredgecolor=GREY))
    for i, b in enumerate(bp["boxes"]):
        last = highlight_last and i == len(data) - 1 and len(data) > 1
        b.set(facecolor="#dbeafe" if last else "#e2e8f0", edgecolor=BLUE if last else SLATE)
    for m in bp["medians"]:
        m.set(color=INK, lw=1.4)
    ax.axhline(lsl, color=RED, ls="--", lw=1)
    ax.text(len(data) + 0.45, lsl, f"LSL {lsl:g}", color=RED, va="bottom", ha="right", fontsize=8)
    ax.set_xticks(range(1, len(labels) + 1), [f"{k}\nn={len(v)}" for k, v in zip(labels, data)])
    ax.set_ylabel(f"{name} (%)")
    ax.set_title(title, loc="left", fontsize=10, fontweight="bold")
    return _png(fig)


def cp_cpk_bars(labels, series: dict, title="Cp · Cpk 추이", width=6.0, height=3.4) -> bytes:
    """series = {'WU Cp': [..], 'WU Cpk': [..], 'BU Cp': [..], 'BU Cpk': [..]} (NaN은 N/A)"""
    colors = {"WU Cp": "#93c5fd", "WU Cpk": BLUE, "BU Cp": "#d1d5db", "BU Cpk": "#4b5563"}
    fig, ax = plt.subplots(figsize=(width, height))
    xpos = np.arange(len(labels))
    keys = list(series)
    w = 0.8 / max(1, len(keys))
    ymax = 1.5
    for i, k in enumerate(keys):
        vals = np.asarray(series[k], dtype=float)
        off = (i - (len(keys) - 1) / 2) * w
        ax.bar(xpos + off, np.nan_to_num(vals), width=w, color=colors.get(k, GREY), label=k)
        for xx, vv in zip(xpos + off, vals):
            ax.text(xx, (0 if not np.isfinite(vv) else vv) + 0.03, "N/A" if not np.isfinite(vv) else f"{vv:.2f}",
                    ha="center", fontsize=7)
        if np.isfinite(vals).any():
            ymax = max(ymax, np.nanmax(vals))
    ax.axhline(1.33, color=RED, ls="--", lw=1)
    ax.text(len(labels) - 0.5, 1.36, "1.33", color=RED, fontsize=8, ha="right")
    ax.set_xticks(xpos, labels)
    ax.set_ylim(0, ymax * 1.25)
    ax.legend(ncol=len(keys), fontsize=8, frameon=False, loc="upper left")
    ax.set_title(title, loc="left", fontsize=10, fontweight="bold")
    return _png(fig)


def ng_rate_line(labels, rates, title="NG율 추이", width=6.0, height=3.0) -> bytes:
    fig, ax = plt.subplots(figsize=(width, height))
    ax.plot(range(len(labels)), rates, marker="o", color=SLATE)
    for i, r in enumerate(rates):
        ax.text(i, r, f" {r:.1f}%", fontsize=8, va="bottom")
    ax.set_xticks(range(len(labels)), labels)
    ax.set_ylim(0, max(max(rates) * 1.4, 1))
    ax.set_ylabel("NG율 (%)")
    ax.set_title(title, loc="left", fontsize=10, fontweight="bold")
    return _png(fig)


def white_xy(labels, xs, ys, typ=(0.293, 0.320), gcs=0.005, title="White 색좌표 평균 이동", width=5.0, height=3.8) -> bytes:
    """typ=None이면 Typ·GCS 원 없이 이동 궤적만 그린다(설정에 Typ가 없는 모델)."""
    fig, ax = plt.subplots(figsize=(width, height))
    if typ is not None:
        th = np.linspace(0, 2 * np.pi, 200)
        ax.plot(typ[0] + gcs * np.cos(th), typ[1] + gcs * np.sin(th), color=RED, ls="--", lw=1)
        ax.scatter([typ[0]], [typ[1]], marker="+", color=RED, s=60)
        ax.text(typ[0], typ[1] + gcs * 1.1, f"Typ ({typ[0]}, {typ[1]}) · GCS {gcs}", color=RED, ha="center", fontsize=8)
        cx, cy, pad = typ[0], typ[1], gcs * 1.16
    else:
        cx, cy = float(np.nanmean(xs)), float(np.nanmean(ys))
        pad = max(0.002, float(np.nanmax(np.abs(np.r_[np.asarray(xs) - cx, np.asarray(ys) - cy]))) * 2.5)
    ax.plot(xs, ys, color="#94a3b8", lw=1)
    offs = [(-0.0004, 0, "right"), (0, 0.0006, "center"), (0.0004, -0.0003, "left")]
    for k, (a, b) in enumerate(zip(xs, ys)):
        ax.scatter([a], [b], color=BLUE if k == len(xs) - 1 else "#64748b", zorder=3, s=30)
        dx, dy, ha = offs[k % 3]
        ax.text(a + dx, b + dy, labels[k], fontsize=8, va="center", ha=ha)
    ax.set_aspect("equal")
    ax.set_xlim(cx - pad, cx + pad)
    ax.set_ylim(cy - pad, cy + pad)
    ax.set_xlabel("White x")
    ax.set_ylabel("White y")
    ax.set_title(title, loc="left", fontsize=10, fontweight="bold")
    return _png(fig)


def mean_band(labels, means, sds, title, ylabel, lsl=None, usl=None, width=6.0, height=3.0) -> bytes:
    """평균 ± 3σ 밴드 (휘도·CR 등)"""
    means, sds = np.asarray(means, float), np.asarray(sds, float)
    fig, ax = plt.subplots(figsize=(width, height))
    xs = np.arange(len(labels))
    ax.fill_between(xs, means - 3 * sds, means + 3 * sds, color="#e2e8f0")
    ax.plot(xs, means, marker="o", color=SLATE)
    for v, lab in ((lsl, "Min"), (usl, "Max")):
        if v is not None:
            ax.axhline(v, color=RED, ls="--", lw=1)
            ax.text(len(labels) - 0.5, v, f"{lab} {v:g}", color=RED, fontsize=8, ha="right", va="bottom")
    ax.set_xticks(xs, labels)
    ax.set_ylabel(ylabel)
    ax.set_title(title, loc="left", fontsize=10, fontweight="bold")
    return _png(fig)


def daily_cpk_line(dates, wu_cpk, bu_cpk, memo=None, title="측정일별 Cp · Cpk", wu_cp=None, bu_cp=None,
                   width=5.4, height=3.4) -> bytes:
    fig, ax = plt.subplots(figsize=(width, height))
    xs = range(len(dates))
    lines = [("WU Cpk", wu_cpk, BLUE, "-"), ("BU Cpk", bu_cpk, "#4b5563", "-")]
    if wu_cp is not None:
        lines.append(("WU Cp", wu_cp, "#93c5fd", ":"))
    if bu_cp is not None:
        lines.append(("BU Cp", bu_cp, "#9ca3af", ":"))
    top = 2.0
    for lab, vals, col, ls in lines:
        vals = np.asarray(vals, float)
        if not np.isfinite(vals).any():
            continue
        ax.plot(xs, vals, marker="o", color=col, ls=ls, label=lab)
        top = max(top, np.nanmax(vals))
        if ls == "-":
            for k, v in enumerate(vals):
                if np.isfinite(v):
                    ax.text(k, v + 0.05, f"{v:.2f}", ha="center", fontsize=8, color=col)
    ax.axhline(1.33, color=RED, ls="--", lw=1)
    ax.set_xticks(list(xs), dates)
    ax.set_ylim(0, top * 1.2)
    ax.legend(fontsize=8, frameon=False, loc="upper right", ncol=2)
    if memo:
        ax.text(0.01, 0.02, f"메모: {memo}", transform=ax.transAxes, fontsize=8, color=AMBER)
    ax.set_title(title, loc="left", fontsize=10, fontweight="bold")
    return _png(fig)
