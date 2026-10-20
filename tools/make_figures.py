# -*- coding: utf-8 -*-
"""보고서용 상가공실 지도 모음 → docs/maps/*.png + docs/상가공실_지도모음_{1,2,3}.docx

  python tools/make_figures.py

모든 지도는 같은 틀: 그림 크기·지도 영역·오른쪽 범례 패널 위치·글꼴·색 구간·축척/방위 위치를 고정한다.
  범위: 행복도시 전체, 생활권(1·2·3·4·6), 동(법정동), 집합상가 밀집지(DBSCAN 군집, 이름은 조사 '동 세부')
  대상: 2023.06·2025.12 × 집합상가·아파트 단지 내 상가
  핫스팟: Gi*(공실률), Gi*(공실 면적), 국지 모란(LISA), 반경 민감도, 2023→2025 변화
"""
from __future__ import annotations

import collections
import json
import math
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
from matplotlib.collections import LineCollection, PolyCollection  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import FancyBboxPatch, Rectangle  # noqa: E402
import matplotlib.patheffects as pe  # noqa: E402
from PIL import Image  # noqa: E402
from sklearn.cluster import DBSCAN  # noqa: E402

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, "tools"))
from make_report import Document, Cm, Pt, WD_ALIGN_PARAGRAPH, para, add_runs, h1, sq, ci, box, MUTED, ACCENT  # noqa: E402

OUTDIR = os.path.join(HERE, "docs", "maps")
FONT = "Malgun Gothic"
plt.rcParams["font.family"] = FONT
plt.rcParams["axes.unicode_minus"] = False

# ── 고정 틀 ─────────────────────────────────────────────────────────
FIG_W, FIG_H, DPI = 6.4, 4.6, 300
MAP_BOX = [0.008, 0.012, 0.735, 0.976]      # 그림 좌표(0~1)
LEG_BOX = [0.752, 0.012, 0.240, 0.976]
MAP_AR = (FIG_W * MAP_BOX[2]) / (FIG_H * MAP_BOX[3])
INK, SUB, LINE = "#15232A", "#4B5B62", "#9AA5AA"
FS_T, FS, FS_S = 8.0, 6.8, 5.8              # 범례 제목·본문·작은 글자

RATE_BREAKS = [10, 20, 30, 40]
RATE_COLORS = ["#2E7D4F", "#8DB86B", "#E5B94A", "#E07B39", "#B03A3A"]
RATE_LABELS = ["10% 미만", "10~20%", "20~30%", "30~40%", "40% 이상"]
GI_CLS = [("hot99", "핫스팟 99%", "#B2182B"), ("hot95", "핫스팟 95%", "#EF8A62"), ("hot90", "핫스팟 90%", "#FDDBC7"),
          ("ns", "유의하지 않음", "#D3D7D9"),
          ("cold90", "콜드스팟 90%", "#D1E5F0"), ("cold95", "콜드스팟 95%", "#67A9CF"), ("cold99", "콜드스팟 99%", "#2166AC")]
LISA_CLS = [("HH", "높음-높음(HH)", "#D7301F"), ("HL", "높음-낮음(HL)", "#FDAE61"), ("ns", "유의하지 않음", "#D3D7D9"),
            ("LH", "낮음-높음(LH)", "#ABD9E9"), ("LL", "낮음-낮음(LL)", "#2C7BB6")]
CHG_BREAKS = [-20, -10, -3, 3, 10, 20]
CHG_COLORS = ["#1A7A43", "#58AE6A", "#B7DDA9", "#E3E3E3", "#F6B48E", "#E2683C", "#A8231E"]
CHG_LABELS = ["-20%p 이하", "-20~-10", "-10~-3", "-3~+3(변화 작음)", "+3~+10", "+10~+20", "+20%p 이상"]
SIZE_CLS = [(2000, 14, "2천㎡ 미만"), (10000, 32, "2천~1만㎡"), (1e12, 62, "1만㎡ 이상")]
KIND_MK = {"mall": "o", "apt": "D"}
YRL = {"2023": "2023년 6월", "2025": "2025년 12월"}
KINDL = {"mall": "집합상가", "apt": "아파트 단지 내 상가"}
ZNAME = {"1": "1생활권", "2": "2생활권", "3": "3생활권", "4": "4생활권", "6": "6생활권"}

KX, KY = 111320 * math.cos(math.radians(36.5)), 110940


def P(lon, lat):
    return (lon - 127.25) * KX, (lat - 36.5) * KY


def jl(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def rate_col(v):
    for b, c in zip(RATE_BREAKS, RATE_COLORS):
        if v < b:
            return c
    return RATE_COLORS[-1]


def chg_col(v):
    for b, c in zip(CHG_BREAKS, CHG_COLORS):
        if v < b:
            return c
    return CHG_COLORS[-1]


def size_of(area):
    for lim, s, _ in SIZE_CLS:
        if area < lim:
            return s
    return SIZE_CLS[-1][1]


# ── 자료 ────────────────────────────────────────────────────────────
class Base:
    def __init__(self):
        fp = jl(os.path.join(HERE, "data", "footprints.json"))["features"]
        self.fp = []
        for f in fp:
            g = f["geometry"]
            polys = g["coordinates"] if g["type"] == "MultiPolygon" else [g["coordinates"]]
            rings = [np.array([P(x, y) for x, y in p[0]]) for p in polys]
            allp = np.vstack(rings)
            self.fp.append((rings, allp.min(0), allp.max(0)))
        self.units = []
        for u in jl(os.path.join(HERE, "assets", "sejong_units.json"))["units"]:
            rings = [np.array([P(x, y) for x, y in r]) for r in u["rings"]]
            self.units.append((u["name"], rings))
        self.zones = [(z["id"], [np.array([P(x, y) for x, y in r]) for r in z["rings"]])
                      for z in jl(os.path.join(HERE, "assets", "sejong_zones.json"))["zones"]]
        self.roads = collections.defaultdict(list)
        for r in jl(os.path.join(HERE, "assets", "sejong_roads.json"))["roads"]:
            for l in r["lines"]:
                self.roads[r["cls"]].append(np.array([P(x, y) for x, y in l]))
        self.S = jl(os.path.join(HERE, "data", "survey_all.json"))
        for r in self.S:
            r["x"], r["y"] = P(r["lon"], r["lat"])


def label_xy(rings):
    r = max(rings, key=lambda a: len(a))
    x, y = r[:, 0], r[:, 1]
    a = x * np.roll(y, -1) - np.roll(x, -1) * y
    A = a.sum() / 2
    if abs(A) < 1:
        return x.mean(), y.mean()
    return ((x + np.roll(x, -1)) * a).sum() / (6 * A), ((y + np.roll(y, -1)) * a).sum() / (6 * A)


def fit_extent(xs, ys, pad=0.10, min_w=900):
    x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
    w, h = max(x1 - x0, 1), max(y1 - y0, 1)
    w, h = w * (1 + 2 * pad), h * (1 + 2 * pad)
    w = max(w, min_w, h * MAP_AR)
    h = w / MAP_AR
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    return cx - w / 2, cx + w / 2, cy - h / 2, cy + h / 2


# ── 그리기 ──────────────────────────────────────────────────────────
HALO = [pe.withStroke(linewidth=2.2, foreground="white")]


def new_fig():
    fig = plt.figure(figsize=(FIG_W, FIG_H), dpi=DPI)
    ax = fig.add_axes(MAP_BOX)
    lg = fig.add_axes(LEG_BOX)
    for a in (ax, lg):
        a.set_xticks([])
        a.set_yticks([])
    lg.set_xlim(0, 1)
    lg.set_ylim(0, 1)
    lg.axis("off")
    return fig, ax, lg


def draw_base(ax, B, ext, hi_fp=()):
    x0, x1, y0, y1 = ext
    w = x1 - x0
    ax.set_xlim(x0, x1)
    ax.set_ylim(y0, y1)
    ax.set_aspect("equal")
    ax.set_facecolor("#FFFFFF")
    for s in ax.spines.values():
        s.set_color(LINE)
        s.set_linewidth(.6)
    vis = lambda mn, mx: not (mx[0] < x0 or mn[0] > x1 or mx[1] < y0 or mn[1] > y1)
    # 행정동 면(아주 옅게) + 경계
    polys = [r for _, rings in B.units for r in rings]
    ax.add_collection(PolyCollection(polys, facecolors="#F7F8F8", edgecolors="none", zorder=0))
    # 도로
    detail = w < 4500
    for cls, col, lw in (("collector", "#E6E9EA", .45), ("sub", "#E6E9EA", .45), ("arterial", "#D3D8DA", 1.0),
                         ("expressway", "#C3CACD", 1.5), ("brt", "#9C86BE", 1.1)):
        if cls in ("collector", "sub") and not detail:
            continue
        ax.add_collection(LineCollection(B.roads[cls], colors=col, linewidths=lw, zorder=1, capstyle="round"))
    # 일반 건물
    gp = [r for i, (rings, mn, mx) in enumerate(B.fp) if i not in hi_fp and vis(mn, mx) for r in rings]
    ax.add_collection(PolyCollection(gp, facecolors="#DCE0E2", edgecolors="none", zorder=2))
    ax.add_collection(LineCollection(polys, colors=LINE, linewidths=.6, zorder=3))
    for zid, rings in B.zones:
        ax.add_collection(LineCollection(rings, colors="#1D5C7C", linewidths=.9, linestyles=(0, (4, 2.5)), zorder=3, alpha=.8))
    # 행정동 이름(지도 안쪽만)
    for name, rings in B.units:
        lx, ly = label_xy(rings)
        if x0 + w * .04 < lx < x1 - w * .04 and y0 + (y1 - y0) * .06 < ly < y1 - (y1 - y0) * .08:
            ax.text(lx, ly, name, fontsize=6.4 if detail else 5.6, color="#6B7A80", ha="center", va="center", zorder=6.5,
                    path_effects=HALO)
    # 생활권 이름(전체 지도에서만)
    if w > 9000:
        for zid, rings in B.zones:
            if zid in ZNAME:
                lx, ly = label_xy(rings)
                ax.text(lx, ly + 450, ZNAME[zid], fontsize=7.2, color="#1D5C7C", fontweight="bold", ha="center", zorder=6.6,
                        path_effects=HALO)
    scalebar(ax, ext)
    north(ax)


def scalebar(ax, ext):
    x0, x1, y0, y1 = ext
    w = x1 - x0
    L = min([100, 200, 250, 500, 1000, 2000, 3000, 5000], key=lambda v: abs(v - w / 5))
    bx, by = x0 + w * .035, y0 + (y1 - y0) * .035
    ax.add_patch(Rectangle((bx - w * .012, by - (y1 - y0) * .02), L + w * .05, (y1 - y0) * .075, facecolor="white",
                           edgecolor="none", alpha=.85, zorder=8))
    ax.plot([bx, bx + L / 2], [by, by], color=INK, lw=2.2, solid_capstyle="butt", zorder=9)
    ax.plot([bx + L / 2, bx + L], [by, by], color="white", lw=2.2, solid_capstyle="butt", zorder=9)
    ax.plot([bx, bx + L], [by, by], color=INK, lw=.4, zorder=9)
    ax.plot([bx, bx + L], [by - (y1 - y0) * .006] * 2, color=INK, lw=.4, zorder=9)
    ax.text(bx + L / 2, by + (y1 - y0) * .016, f"{L / 1000:g} km" if L >= 1000 else f"{L} m", fontsize=5.8, ha="center",
            color=INK, zorder=9)


def north(ax):
    ax.annotate("", xy=(.955, .965), xytext=(.955, .905), xycoords="axes fraction",
                arrowprops=dict(arrowstyle="-|>", color=INK, lw=.9, mutation_scale=7), zorder=9)
    ax.text(.955, .87, "N", transform=ax.transAxes, fontsize=6.2, ha="center", va="center", color=INK, fontweight="bold", zorder=9)


def tag(ax, text):
    ax.text(.018, .978, text, transform=ax.transAxes, fontsize=7.2, fontweight="bold", color="white", va="top", ha="left", zorder=10,
            bbox=dict(boxstyle="round,pad=.35,rounding_size=.25", facecolor="#1D5C7C", edgecolor="none"))


def draw_recs(ax, B, recs, colf, ext, labels=False, fp=True):
    """조사 대상: 바닥 외곽선(색) + 기호(집합 ●, 아파트 ◆, 크기 = 상가 면적)."""
    x0, x1, y0, y1 = ext
    w = x1 - x0
    if fp:
        polys, cols = [], []
        for r in recs:
            for i in r.get("fp") or []:
                for ring in B.fp[i][0]:
                    polys.append(ring)
                    cols.append(colf(r))
        ax.add_collection(PolyCollection(polys, facecolors=cols, edgecolors="#3A3A3A", linewidths=.25, zorder=5))
    for kind in ("mall", "apt"):
        rs = sorted([r for r in recs if r["kind"] == kind], key=lambda r: -r["area"])
        if not rs:
            continue
        ax.scatter([r["x"] for r in rs], [r["y"] for r in rs], s=[size_of(r["area"]) for r in rs], c=[colf(r) for r in rs],
                   marker=KIND_MK[kind], edgecolors="#202020", linewidths=.45, zorder=6)
    if labels:
        for r in recs:
            if x0 < r["x"] < x1 and y0 < r["y"] < y1:
                ax.text(r["x"] + w * .012, r["y"] + w * .006, f"{r['rate']:.0f}", fontsize=4.8, color=INK, zorder=7,
                        path_effects=[pe.withStroke(linewidth=1.4, foreground="white")])


# ── 범례(위치 고정) ─────────────────────────────────────────────────
# 범례 칸 위치(그림 좌표 0~1, 모든 그림 공통 — 구간 수가 달라도 아래 칸은 움직이지 않는다)
Y_CLS, Y_SYM, Y_SIZE, Y_BASE, STEP = .975, .712, .598, .452, .030


def row_patch(lg, y, col, edge="#555"):
    lg.add_patch(Rectangle((0.03, y - .013), .085, .026, facecolor=col, edgecolor=edge, linewidth=.3))


def legend(lg, cls_title, cls_rows, summary, extra_note=None):
    lg.add_patch(FancyBboxPatch((0.0, 0.0), 1.0, 1.0, boxstyle="square,pad=0", facecolor="#FAFBFB", edgecolor=LINE, linewidth=.6))
    T = lambda y, t: lg.text(0.02, y, t, fontsize=FS_T, fontweight="bold", color=INK, va="top")
    L = lambda y, t: lg.text(0.15, y, t, fontsize=FS, color=INK, va="center")
    T(Y_CLS, cls_title)
    y = Y_CLS - .050
    for lab, col in cls_rows:
        row_patch(lg, y, col)
        L(y, lab)
        y -= STEP
    T(Y_SYM, "조사 대상(기호)")
    y = Y_SYM - .050
    for kind in ("mall", "apt"):
        lg.scatter([0.072], [y], s=26, marker=KIND_MK[kind], facecolor="white", edgecolor="#202020", linewidths=.5)
        L(y, KINDL[kind])
        y -= STEP
    T(Y_SIZE, "기호 크기(상가 면적)")
    y = Y_SIZE - .050
    for lim, sz, lab in SIZE_CLS:
        lg.scatter([0.072], [y], s=sz, marker="o", facecolor="white", edgecolor="#202020", linewidths=.5)
        L(y, lab)
        y -= STEP
    T(Y_BASE, "바탕")
    y = Y_BASE - .050
    row_patch(lg, y, "#E07B39", "#3A3A3A")
    L(y, "조사 건물 바닥(같은 색)")
    y -= STEP
    row_patch(lg, y, "#DCE0E2", "none")
    L(y, "그 밖의 건물")
    y -= STEP
    for lab, kw in (("행정동 경계", dict(color=LINE, lw=.8)), ("생활권 경계", dict(color="#1D5C7C", lw=.9, ls=(0, (4, 2.5)))),
                    ("간선도로", dict(color="#C3CACD", lw=1.6)), ("BRT", dict(color="#9C86BE", lw=1.3))):
        lg.add_line(Line2D([0.03, 0.115], [y] * 2, **kw))
        L(y, lab)
        y -= STEP
    # 요약(아래 고정)
    lg.add_patch(Rectangle((0.03, 0.018), .94, .118, facecolor="white", edgecolor=LINE, linewidth=.5))
    for i, ln in enumerate(summary[:3]):
        lg.text(0.06, 0.112 - i * .034, ln, fontsize=FS if i else FS + .2, fontweight="bold" if i == 0 else "normal", color=INK, va="center")
    if extra_note:
        lg.text(0.02, 0.148, extra_note, fontsize=FS_S, color=SUB, va="bottom", wrap=True)


def rate_legend(lg, summary, note=None):
    legend(lg, "공실률(면적 기준)", list(zip(RATE_LABELS, RATE_COLORS)), summary, note)


def agg(recs):
    a = sum(r["area"] for r in recs)
    v = sum(r["vac"] for r in recs)
    return (v / a * 100 if a else 0), a, v


def save(fig, name):
    os.makedirs(OUTDIR, exist_ok=True)
    p = os.path.join(OUTDIR, name + ".png")
    fig.savefig(p, dpi=DPI, facecolor="white")
    plt.close(fig)
    im = Image.open(p).convert("RGB").quantize(colors=256, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE)
    im.save(p, optimize=True)
    return p


# ── 공간 통계 ───────────────────────────────────────────────────────
def dmat(recs):
    X = np.array([[r["x"], r["y"]] for r in recs])
    return np.sqrt(((X[:, None, :] - X[None, :, :]) ** 2).sum(-1))


def gistar(recs, val, band):
    x = np.array([val(r) for r in recs], float)
    n = len(x)
    W = (dmat(recs) <= band).astype(float)          # 자기 자신 포함(Gi*)
    xb, s = x.mean(), x.std()
    sw, sw2 = W.sum(1), (W ** 2).sum(1)
    num = W @ x - xb * sw
    den = s * np.sqrt((n * sw2 - sw ** 2) / (n - 1))
    z = np.where(den > 0, num / np.where(den > 0, den, 1), 0)
    out = []
    for zi in z:
        a = abs(zi)
        lvl = "99" if a >= 2.576 else "95" if a >= 1.96 else "90" if a >= 1.645 else None
        out.append("ns" if not lvl else ("hot" if zi > 0 else "cold") + lvl)
    return out, z


def lisa(recs, val, band, perm=999, seed=7):
    x = np.array([val(r) for r in recs], float)
    z = (x - x.mean()) / (x.std() or 1)
    W = ((dmat(recs) <= band) & (dmat(recs) > 0)).astype(float)
    rs = W.sum(1)
    Wr = np.divide(W, rs[:, None], out=np.zeros_like(W), where=rs[:, None] > 0)
    lag = Wr @ z
    I = z * lag
    rng = np.random.default_rng(seed)
    out = []
    n = len(z)
    for i in range(n):
        k = int(rs[i])
        if k == 0:
            out.append("ns")
            continue
        others = np.delete(z, i)
        sims = np.array([z[i] * rng.choice(others, k, replace=False).mean() for _ in range(perm)])
        p = (np.sum(np.abs(sims) >= abs(I[i])) + 1) / (perm + 1)
        if p >= .05:
            out.append("ns")
        else:
            out.append(("H" if z[i] > 0 else "L") + ("H" if lag[i] > 0 else "L"))
    return out


# ── 그림 목록 만들기 ────────────────────────────────────────────────
class Book:
    def __init__(self, title, sub):
        self.title, self.sub, self.items = title, sub, []

    def sec(self, t, note=None):
        self.items.append(("sec", t, note))

    def fig(self, path, title, desc):
        self.items.append(("fig", path, title, desc))


def sel(B, yr, kind, pred=lambda r: True):
    return [r for r in B.S if r["yr"] == yr and r["kind"] == kind and pred(r)]


def short(nm):
    return nm.split(" · ")[0] + (" 일대" if " · " in nm else "")


def in_ext(r, ext):
    return ext[0] <= r["x"] <= ext[1] and ext[2] <= r["y"] <= ext[3]


def rate_desc(recs, area_name, yr, kind, group_key="sub"):
    rate, a, v = agg(recs)
    unit = "동" if kind == "mall" else "곳"
    hi = sum(1 for r in recs if r["rate"] >= 40)
    zero = sum(1 for r in recs if r["rate"] == 0)
    g = collections.defaultdict(list)
    for r in recs:
        g[r[group_key]].append(r)
    gs = sorted(((agg(v_)[0], k, len(v_)) for k, v_ in g.items() if len(v_) >= 2), reverse=True)
    s = (f"{area_name}의 {YRL[yr]} 조사 {KINDL[kind]} {len(recs)}{unit}, 상가 면적 {a:,.0f}㎡ 중 공실 {v:,.0f}㎡로 공실률(면적) {rate:.1f}%. "
         f"공실률 40% 이상 {hi}{unit}, 공실 0% {zero}{unit}.")
    if len(gs) >= 2:
        s += f" 동 세부 기준 가장 높은 곳은 {gs[0][1]}({gs[0][0]:.1f}%, {gs[0][2]}{unit}), 가장 낮은 곳은 {gs[-1][1]}({gs[-1][0]:.1f}%)."
    return s


def rate_map(B, recs_all, ext, tagtxt, name, labels=False, summary_name=""):
    recs = [r for r in recs_all if in_ext(r, ext)]
    fig, ax, lg = new_fig()
    hi = {i for r in recs for i in (r.get("fp") or [])}
    draw_base(ax, B, ext, hi)
    draw_recs(ax, B, recs, lambda r: rate_col(r["rate"]), ext, labels=labels)
    tag(ax, tagtxt)
    rate, a, v = agg(recs_all)
    unit = "동" if recs_all and recs_all[0]["kind"] == "mall" else "곳"
    rate_legend(lg, [summary_name or "선택 범위", f"대상 {len(recs_all)}{unit} · 공실률 {rate:.1f}%", f"상가 {a / 1e4:,.1f}만㎡ · 공실 {v / 1e4:,.1f}만㎡"],
                "숫자 = 건물 공실률(%)" if labels else None)
    return save(fig, name)


def class_map(B, recs, cls, ext, tagtxt, name, rows, title, summary, note=None):
    col = dict((k, c) for k, _, c in rows)
    cmap = {id(r): col[c] for r, c in zip(recs, cls)}
    fig, ax, lg = new_fig()
    hi = {i for r in recs for i in (r.get("fp") or [])}
    draw_base(ax, B, ext, hi)
    order = sorted(range(len(recs)), key=lambda i: cls[i] != "ns")     # 유의한 것이 위로
    rr = [recs[i] for i in order]
    draw_recs(ax, B, rr, lambda r: cmap[id(r)], ext)
    tag(ax, tagtxt)
    legend(lg, title, [(lab, c) for _, lab, c in rows], summary, note)
    return save(fig, name)


def main():
    B = Base()
    allx = [r["x"] for r in B.S]
    ally = [r["y"] for r in B.S]
    EXT_ALL = fit_extent(allx, ally, pad=.04)
    combos = [("2023", "mall"), ("2025", "mall"), ("2023", "apt"), ("2025", "apt")]
    tagf = lambda yr, kind, where: f"{where} · {YRL[yr]} · {KINDL[kind]}"
    n = collections.Counter()

    # ━━ 책 1: 전체·생활권·핫스팟 ━━
    b1 = Book("행복도시 상가공실 지도 모음 ①", "전체 · 생활권 · 핫스팟 · 2023→2025 변화")
    b1.sec("1. 행복도시 전체", "네 장 모두 같은 범위·같은 색 구간. 기호 = 조사 단위(집합상가 건물 ●, 아파트 단지 상가 ◆), 색 = 공실률(면적), 크기 = 상가 면적.")
    for yr, kind in combos:
        recs = sel(B, yr, kind)
        p = rate_map(B, recs, EXT_ALL, tagf(yr, kind, "행복도시 전체"), f"A_all_{yr}_{kind}", summary_name="행복도시 전체")
        b1.fig(p, f"행복도시 전체 — {YRL[yr]} {KINDL[kind]} 공실률", rate_desc(recs, "행복도시 전체", yr, kind, "dong"))
    b1.sec("2. 생활권별", "생활권 경계(점선) 안 조사 대상. 5생활권은 조사 대상이 없다.")
    for z in ("1", "2", "3", "4", "6"):
        zrec = [r for r in B.S if r["zone"] == z]
        ext = fit_extent([r["x"] for r in zrec], [r["y"] for r in zrec], pad=.08, min_w=2500)
        for yr, kind in combos:
            recs = sel(B, yr, kind, lambda r: r["zone"] == z)
            if not recs:
                continue
            p = rate_map(B, recs, ext, tagf(yr, kind, ZNAME[z]), f"B_zone{z}_{yr}_{kind}", summary_name=ZNAME[z])
            b1.fig(p, f"{ZNAME[z]} — {YRL[yr]} {KINDL[kind]} 공실률", rate_desc(recs, ZNAME[z], yr, kind))

    # 핫스팟
    b1.sec("3. 핫스팟 ① Gi* — 공실률(면적) 기준, 반경 500m",
           "Getis-Ord Gi*: 반경 500m 이웃(자기 포함)의 공실률 합이 전체 평균에서 기대되는 것보다 얼마나 큰지의 z값. "
           "|z| ≥ 1.645·1.96·2.576 → 90·95·99% 신뢰수준. 핫스팟 = 공실률 높은 건물이 모인 곳, 콜드스팟 = 낮은 건물이 모인 곳.")
    for yr, kind in combos:
        recs = sel(B, yr, kind)
        cls, z = gistar(recs, lambda r: r["rate"], 500)
        c = collections.Counter(cls)
        p = class_map(B, recs, cls, EXT_ALL, f"Gi* 공실률 · {YRL[yr]} · {KINDL[kind]}", f"C_gi_rate_{yr}_{kind}", GI_CLS,
                      "Gi* 핫스팟(공실률)", ["반경 500m · 공실률(면적)", f"핫 {c['hot99'] + c['hot95'] + c['hot90']} · 콜드 {c['cold99'] + c['cold95'] + c['cold90']}",
                                          f"전체 {len(recs)} · 유의 안 함 {c['ns']}"])
        b1.fig(p, f"Gi* 핫스팟(공실률) — {YRL[yr]} {KINDL[kind]}", hot_desc(recs, cls, kind))
    b1.sec("4. 핫스팟 ② Gi* — 공실 면적(㎡) 기준, 반경 500m",
           "공실'률'이 아니라 비어 있는 면적의 절대량으로 본 핫스팟. 큰 상가가 크게 비어 실제로 빈 공간이 쌓인 곳을 찾는다(정책 물량 산정에 유리).")
    for yr, kind in combos:
        recs = sel(B, yr, kind)
        cls, z = gistar(recs, lambda r: r["vac"], 500)
        c = collections.Counter(cls)
        p = class_map(B, recs, cls, EXT_ALL, f"Gi* 공실 면적 · {YRL[yr]} · {KINDL[kind]}", f"D_gi_vac_{yr}_{kind}", GI_CLS,
                      "Gi* 핫스팟(공실 면적)", ["반경 500m · 공실 면적(㎡)", f"핫 {c['hot99'] + c['hot95'] + c['hot90']} · 콜드 {c['cold99'] + c['cold95'] + c['cold90']}",
                                           f"전체 {len(recs)} · 유의 안 함 {c['ns']}"])
        b1.fig(p, f"Gi* 핫스팟(공실 면적) — {YRL[yr]} {KINDL[kind]}", hot_desc(recs, cls, kind, "공실 면적"))
    b1.sec("5. 핫스팟 ③ 국지 모란 지수(LISA) — 공실률, 반경 500m",
           "자기 공실률과 이웃 평균(자기 제외, 행 표준화)의 관계. HH = 높은 곳이 높은 곳에 둘러싸임(공실 집적), LL = 낮은 곳끼리, "
           "HL = 주변은 낮은데 혼자 높음(개별 건물 문제), LH = 주변은 높은데 혼자 낮음. 조건부 무작위 순열 999회, p<0.05.")
    for yr, kind in combos:
        recs = sel(B, yr, kind)
        cls = lisa(recs, lambda r: r["rate"], 500)
        c = collections.Counter(cls)
        p = class_map(B, recs, cls, EXT_ALL, f"LISA 공실률 · {YRL[yr]} · {KINDL[kind]}", f"E_lisa_{yr}_{kind}", LISA_CLS,
                      "국지 모란(LISA)", ["반경 500m · p<0.05", f"HH {c['HH']} · LL {c['LL']}", f"HL {c['HL']} · LH {c['LH']} · 유의 안 함 {c['ns']}"])
        b1.fig(p, f"LISA 군집 — {YRL[yr]} {KINDL[kind]}", lisa_desc(recs, cls, kind))
    b1.sec("6. 핫스팟 ④ 반경 민감도 — 2025년 12월 집합상가 Gi*(공실률)",
           "이웃 반경을 300m·500m·1,000m로 바꿔 같은 Gi*를 계산했다. 반경이 작으면 블록 단위, 크면 생활권 단위의 집적이 드러난다.")
    for band in (300, 500, 1000):
        recs = sel(B, "2025", "mall")
        cls, z = gistar(recs, lambda r: r["rate"], band)
        c = collections.Counter(cls)
        p = class_map(B, recs, cls, EXT_ALL, f"Gi* 공실률 · 반경 {band:,}m · 2025년 12월 · 집합상가", f"F_gi_band{band}", GI_CLS,
                      "Gi* 핫스팟(공실률)", [f"반경 {band:,}m · 공실률(면적)", f"핫 {c['hot99'] + c['hot95'] + c['hot90']} · 콜드 {c['cold99'] + c['cold95'] + c['cold90']}",
                                          f"전체 {len(recs)} · 유의 안 함 {c['ns']}"])
        b1.fig(p, f"반경 {band:,}m Gi* 핫스팟 — 2025년 12월 집합상가", hot_desc(recs, cls, "mall"))

    b1.sec("7. 2023 → 2025 변화(같은 건물·단지)",
           "두 조사에 모두 있는 같은 코드의 건물(단지)만 비교. 색 = 공실률 변화(2025 − 2023, %p): 초록 = 공실 감소, 빨강 = 증가. "
           "위치는 2025년 조사 좌표.")
    for kind in ("mall", "apt"):
        pairs = change_pairs(B, kind)
        if len(pairs) < 5:
            continue
        p = change_map(B, pairs, EXT_ALL, f"공실률 변화 · 2023→2025 · {KINDL[kind]}", f"G_chg_{kind}", "행복도시 전체")
        b1.fig(p, f"공실률 변화(2023.06 → 2025.12) — {KINDL[kind]}", chg_desc(pairs, kind))
        if len(pairs) >= 30:
            cls, z = gistar([q for q, _ in pairs], lambda q: q["d"], 500)
            c = collections.Counter(cls)
            p = class_map(B, [q for q, _ in pairs], cls, EXT_ALL, f"Gi* 변화 · 2023→2025 · {KINDL[kind]}", f"G_chg_gi_{kind}",
                          [(k, l.replace("핫스팟", "증가 집적").replace("콜드스팟", "감소 집적"), cc) for k, l, cc in GI_CLS],
                          "Gi*(공실률 변화)", ["반경 500m · 변화량(%p)", f"증가 집적 {c['hot99'] + c['hot95'] + c['hot90']} · 감소 집적 {c['cold99'] + c['cold95'] + c['cold90']}",
                                             f"비교 {len(pairs)} · 유의 안 함 {c['ns']}"])
            b1.fig(p, f"공실률 변화의 Gi* — {KINDL[kind]}",
                   f"공실률이 함께 늘어난(증가 집적) 곳과 함께 줄어든(감소 집적) 곳. 증가 집적 {c['hot99'] + c['hot95'] + c['hot90']}{'동' if kind == 'mall' else '곳'}"
                   f"{where_txt([q for q, cl in zip([q for q, _ in pairs], cls) if cl.startswith('hot')])}, 감소 집적 "
                   f"{c['cold99'] + c['cold95'] + c['cold90']}{'동' if kind == 'mall' else '곳'}"
                   f"{where_txt([q for q, cl in zip([q for q, _ in pairs], cls) if cl.startswith('cold')])}.")

    # ━━ 책 2: 동별 ━━
    b2 = Book("행복도시 상가공실 지도 모음 ②", "동(법정동)별 — 2023·2025 × 집합상가·아파트 상가")
    dongs = sorted({r["dong"] for r in B.S}, key=lambda d: (min(r["zone"] for r in B.S if r["dong"] == d), d))
    for d in dongs:
        drec = [r for r in B.S if r["dong"] == d]
        ext = fit_extent([r["x"] for r in drec], [r["y"] for r in drec], pad=.12, min_w=1500)
        b2.sec(f"{d} ({ZNAME.get(drec[0]['zone'], drec[0]['zone'])})",
               "네 장 모두 같은 범위. 숫자 = 건물(단지) 공실률(%). 해당 조사에 대상이 없으면 그림을 넣지 않았다.")
        for yr, kind in combos:
            recs = sel(B, yr, kind, lambda r: r["dong"] == d)
            if not recs:
                continue
            p = rate_map(B, recs, ext, tagf(yr, kind, d), f"H_{d}_{yr}_{kind}", labels=True, summary_name=d)
            b2.fig(p, f"{d} — {YRL[yr]} {KINDL[kind]} 공실률", rate_desc(recs, d, yr, kind))

    # ━━ 책 3: 밀집지 ━━
    b3 = Book("행복도시 상가공실 지도 모음 ③", "집합상가 밀집지별 — 2023·2025 공실률과 변화")
    malls = [r for r in B.S if r["kind"] == "mall"]
    X = np.array([[r["x"], r["y"]] for r in malls])
    lab = DBSCAN(eps=250, min_samples=8).fit(X).labels_
    groups = collections.defaultdict(list)
    for r, l in zip(malls, lab):
        if l >= 0:
            groups[l].append(r)
    named = []
    for l, rs in groups.items():
        top = [k for k, _ in collections.Counter(r["sub"] for r in rs).most_common(3)]
        nm = " · ".join(top[:2]) + (" 일대" if len(top) > 2 else "")
        named.append((collections.Counter(r["zone"] for r in rs).most_common(1)[0][0], nm, rs))
    named.sort(key=lambda t: (t[0], t[1]))
    b3.sec("밀집지 정의",
           "2023·2025 집합상가 건물 위치를 DBSCAN(이웃 반경 250m, 최소 8곳)으로 묶은 군집 "
           f"{len(named)}곳. 이름은 군집 안에 많은 조사 '동 세부' 1~2개. 군집에 들지 않은 외딴 건물 {int((lab < 0).sum())}건은 제외. "
           "각 밀집지마다 2023 공실률 · 2025 공실률 · 같은 건물의 변화 세 장. 범위 안의 아파트 단지 상가도 같은 범례로 함께 표시(2025).")
    for zi, nm, rs in named:
        ext = fit_extent([r["x"] for r in rs], [r["y"] for r in rs], pad=.14, min_w=800)
        b3.sec(f"{nm} ({ZNAME.get(zi, zi)})")
        codes = {r["code"] for r in rs}
        for yr in ("2023", "2025"):
            recs = [r for r in rs if r["yr"] == yr]
            if not recs:
                continue
            apts = [r for r in sel(B, yr, "apt") if in_ext(r, ext)] if yr == "2025" else []
            p = rate_map(B, recs + apts, ext, f"{nm} · {YRL[yr]}", f"I_{zi}_{nm.replace(' · ', '_').replace(' ', '')}_{yr}",
                         labels=True, summary_name=short(nm))
            d = rate_desc(recs, nm, yr, "mall")
            if apts:
                ra = agg(apts)[0]
                d += f" 범위 안 아파트 단지 상가 {len(apts)}곳(공실률 {ra:.1f}%)도 ◆로 함께 표시."
            b3.fig(p, f"{nm} — {YRL[yr]} 집합상가 공실률", d)
        pairs = [pq for pq in change_pairs(B, "mall") if pq[0]["code"] in codes]
        if len(pairs) >= 3:
            p = change_map(B, pairs, ext, f"{nm} · 공실률 변화 2023→2025", f"J_{zi}_{nm.replace(' · ', '_').replace(' ', '')}_chg", short(nm),
                           labels=True)
            b3.fig(p, f"{nm} — 공실률 변화(2023.06 → 2025.12)", chg_desc(pairs, "mall", nm))

    for i, b in enumerate((b1, b2, b3), start=1):
        write_book(b, os.path.join(HERE, "docs", f"상가공실_지도모음_{i}.docx"))


# ── 변화 ────────────────────────────────────────────────────────────
def change_pairs(B, kind):
    """같은 건물(집합상가 = 조사 코드) · 같은 단지(아파트 = 두 조사의 코드 체계가 달라 주소, 없으면 단지명)."""
    import re
    nz = lambda t: re.sub(r"\s|\(.*?\)", "", t or "")
    old = [r for r in B.S if r["yr"] == "2023" and r["kind"] == kind]
    if kind == "mall":
        look = lambda r: a.get(r["code"])
        a = {r["code"]: r for r in old}
    else:
        byaddr = {nz(r["addr"]): r for r in old}
        byname = {nz(r["name"].split("|")[0]): r for r in old}
        look = lambda r: byaddr.get(nz(r["addr"])) or byname.get(nz(r["name"]))
    out = []
    for r in B.S:
        if r["yr"] == "2025" and r["kind"] == kind:
            p = look(r)
            if p:
                q = dict(r)
                q["d"] = r["rate"] - p["rate"]
                q["r23"] = p["rate"]
                out.append((q, p))
    return out


def change_map(B, pairs, ext, tagtxt, name, sname, labels=False):
    recs = [q for q, _ in pairs]
    fig, ax, lg = new_fig()
    hi = {i for r in recs for i in (r.get("fp") or [])}
    draw_base(ax, B, ext, hi)
    draw_recs(ax, B, recs, lambda r: chg_col(r["d"]), ext)
    if labels:
        w = ext[1] - ext[0]
        for r in recs:
            if in_ext(r, ext):
                ax.text(r["x"] + w * .012, r["y"] + w * .006, f"{r['d']:+.0f}", fontsize=4.8, color=INK, zorder=7,
                        path_effects=[pe.withStroke(linewidth=1.4, foreground="white")])
    tag(ax, tagtxt)
    a23 = agg([p for _, p in pairs])[0]
    a25 = agg(recs)[0]
    up = sum(1 for r in recs if r["d"] >= 3)
    dn = sum(1 for r in recs if r["d"] <= -3)
    legend(lg, "공실률 변화(%p)", list(zip(CHG_LABELS, CHG_COLORS)),
           [sname, f"비교 {len(recs)} · {a23:.1f}% → {a25:.1f}%", f"증가 {up} · 감소 {dn} · 비슷 {len(recs) - up - dn}"],
           "숫자 = 변화(%p)" if labels else None)
    return save(fig, name)


def chg_desc(pairs, kind, where="행복도시 전체"):
    recs = [q for q, _ in pairs]
    unit = "동" if kind == "mall" else "곳"
    a23, a25 = agg([p for _, p in pairs])[0], agg(recs)[0]
    up = sorted(recs, key=lambda r: -r["d"])
    s = (f"{where}에서 두 조사에 모두 있는 {KINDL[kind]} {len(recs)}{unit}의 공실률(면적)은 {a23:.1f}% → {a25:.1f}%({a25 - a23:+.1f}%p). "
         f"3%p 이상 증가 {sum(1 for r in recs if r['d'] >= 3)}{unit}, 3%p 이상 감소 {sum(1 for r in recs if r['d'] <= -3)}{unit}.")
    g = collections.defaultdict(list)
    for q, p in pairs:
        g[q["sub"]].append((q, p))
    gs = sorted(((agg([q for q, _ in v])[0] - agg([p for _, p in v])[0], k) for k, v in g.items() if len(v) >= 2))
    if len(gs) >= 2:
        s += f" 동 세부 기준 가장 크게 줄어든 곳 {gs[0][1]}({gs[0][0]:+.1f}%p), 가장 크게 늘어난 곳 {gs[-1][1]}({gs[-1][0]:+.1f}%p)."
    return s


def where_txt(recs, k=3):
    if not recs:
        return ""
    c = collections.Counter(r["sub"] for r in recs).most_common(k)
    return "(" + ", ".join(f"{s} {n}" for s, n in c) + ")"


def hot_desc(recs, cls, kind, what="공실률"):
    unit = "동" if kind == "mall" else "곳"
    hot = [r for r, c in zip(recs, cls) if c.startswith("hot")]
    cold = [r for r, c in zip(recs, cls) if c.startswith("cold")]
    s = (f"전체 {len(recs)}{unit} 중 {what} 핫스팟 {len(hot)}{unit}{where_txt(hot)}, 콜드스팟 {len(cold)}{unit}{where_txt(cold)}. ")
    if hot:
        s += f"핫스팟의 공실률(면적) {agg(hot)[0]:.1f}%, "
    if cold:
        s += f"콜드스팟 {agg(cold)[0]:.1f}%."
    if len(recs) < 40:
        s += f" 대상이 {len(recs)}{unit}뿐이라 이웃 수가 적어 결과가 불안정하다 — 참고용."
    return s


def lisa_desc(recs, cls, kind):
    unit = "동" if kind == "mall" else "곳"
    g = {k: [r for r, c in zip(recs, cls) if c == k] for k in ("HH", "LL", "HL", "LH")}
    s = (f"HH(공실 집적) {len(g['HH'])}{unit}{where_txt(g['HH'])}, LL(낮은 곳 집적) {len(g['LL'])}{unit}{where_txt(g['LL'])}, "
         f"HL(주변보다 혼자 높음) {len(g['HL'])}{unit}{where_txt(g['HL'], 2)}, LH {len(g['LH'])}{unit}. "
         "HH는 구역 단위 대응, HL은 개별 건물 사정(관리·임대 조건) 점검 대상으로 읽을 수 있다.")
    if len(recs) < 40:
        s += f" 대상이 {len(recs)}{unit}뿐이라 참고용."
    return s


# ── Word ────────────────────────────────────────────────────────────
def write_book(b, path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    sec.left_margin = sec.right_margin = Cm(2.0)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_runs(fp, f"세종연구원 연구모임 · {b.title}", 7.5, color=MUTED)
    para(doc, "세종연구원 · 연구모임 검토자료", 9, color=MUTED, after=2)
    para(doc, b.title, 17, True, ACCENT, after=2)
    para(doc, b.sub, 10.5, color=MUTED, after=8)
    box(doc, [
        "**자료**: 세종연구원 행정중심복합도시 상가 전수조사(2023년 6월·2025년 12월), 집합상가(건물 단위)와 아파트 단지 내 상가(단지 단위).",
        "**공통 틀**: 모든 그림은 같은 크기(300dpi)·같은 범례 위치(오른쪽)·같은 글꼴·같은 색 구간을 쓴다. 공실률 = 공실 면적 ÷ 상가 면적.",
        "**기호**: ● 집합상가 건물, ◆ 아파트 단지 내 상가(단지 대표 주소 위치). 크기 = 상가 면적 3단계. 조사 건물의 바닥 외곽선도 같은 색으로 칠했다"
        "(아파트 상가는 단지 필지의 낮은 근생동을 찾은 경우만).",
        "**왼쪽 위 꼬리표** = 범위 · 조사 시점 · 대상, **오른쪽 아래 상자** = 그 그림의 대상 수와 합계 공실률.",
        "그림 파일 원본은 docs/maps 폴더에 같은 이름의 PNG로 있다.",
    ], title="읽는 법")
    k = 0
    for it in b.items:
        if it[0] == "sec":
            h1(doc, it[1])
            if it[2]:
                ci(doc, it[2])
            continue
        _, p, title, desc = it
        k += 1
        pp = doc.add_paragraph()
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.space_before = Pt(6)
        pp.paragraph_format.space_after = Pt(1)
        pp.paragraph_format.keep_with_next = True
        pp.add_run().add_picture(p, width=Cm(16.5))
        c = para(doc, f"그림 {k}. {title}", 9.5, True, ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, before=1, after=1)
        c.paragraph_format.keep_with_next = True
        para(doc, desc, 8.8, color=MUTED, after=8, line=1.1)
    try:
        doc.save(path)
    except PermissionError:                 # Word 에서 열려 있으면 다른 이름으로
        path = path.replace(".docx", "_v2.docx")
        doc.save(path)
    print("저장:", path, f"그림 {k}장", f"{os.path.getsize(path) / 1e6:.1f}MB")


if __name__ == "__main__":
    main()
