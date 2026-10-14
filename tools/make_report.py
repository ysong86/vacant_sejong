# -*- coding: utf-8 -*-
"""보고자료: 행복도시 상가공실 — 조사 자료(엑셀)·분석 요인·분석 과정과 결과 (.docx)

  python tools/make_report.py

- 수치는 모두 data/buildings.json(prepare.py)·data/analysis.json(analyze.py)에서 다시 읽고,
  회귀·분해는 analyze.py 의 함수로 다시 계산해 analysis.json 과 맞는지 검산한다.
- 서식은 생활권 상황판 보고서(sejong_livingzone/tools/make_report.py)와 같다(맑은 고딕, A4).
"""
from __future__ import annotations

import json
import math
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import openpyxl  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)
import analyze as AN  # noqa: E402

OUT = os.path.join(HERE, "docs", "상가공실_분석_보고서.docx")
FIG = os.path.join(HERE, "docs", "figures")
FONT = "맑은 고딕"
INK = RGBColor(0x15, 0x23, 0x2A)
ACCENT = RGBColor(0x1D, 0x5C, 0x7C)
MUTED = RGBColor(0x4B, 0x5B, 0x62)
HEAD_FILL = "E1ECF1"
BOX_FILL = "F2F6F7"
RATE_BREAKS = [10, 20, 30, 40]
RATE_COLORS = ["#2E7D4F", "#8DB86B", "#E5B94A", "#E07B39", "#B03A3A"]
rate_color = lambda v: RATE_COLORS[sum(v >= b for b in RATE_BREAKS)]
plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False


# ── 서식 도우미(생활권 보고서와 같음) ─────────────────────────────────────
def set_font(run, size=10, bold=False, color=INK):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_runs(p, text, size=10, bold=False, color=INK):
    for i, part in enumerate(text.split("**")):
        if part:
            set_font(p.add_run(part), size, bold or i % 2 == 1, color)


def para(d, text="", size=10, bold=False, color=INK, align=None, before=0, after=3, indent=0.0, hanging=0.0, line=1.12):
    p = d.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line
    if indent or hanging:
        pf.left_indent = Cm(indent + hanging)
        pf.first_line_indent = Cm(-hanging)
    if align is not None:
        p.alignment = align
    if text:
        add_runs(p, text, size, bold, color)
    return p


def shade(cell, fill):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    for k, v in (("val", "clear"), ("color", "auto"), ("fill", fill)):
        shd.set(qn(f"w:{k}"), v)
    tcPr.append(shd)


def cell_margins(t, top=50, bottom=50, left=90, right=90):
    mar = OxmlElement("w:tblCellMar")
    for side, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    t._element.tblPr.append(mar)


def borders(t, color="B8C4C8", size=4):
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        for k, v in (("val", "single"), ("sz", str(size)), ("color", color)):
            el.set(qn(f"w:{k}"), v)
        b.append(el)
    t._element.tblPr.append(b)


def table(d, header, rows, widths, size=8.5, align_center=(), align_right=(), bold_rows=()):
    t = d.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    borders(t)
    cell_margins(t, 30, 30, 80, 80)
    for r_i, row in enumerate([header] + rows):
        for c_i, text in enumerate(row):
            c = t.cell(r_i, c_i)
            c.width = Cm(widths[c_i])
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            if r_i == 0 or c_i in align_center:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif c_i in align_right:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            add_runs(p, str(text), size, r_i == 0 or (r_i - 1) in bold_rows)
            if r_i == 0:
                shade(c, HEAD_FILL)
    for row in t.rows:
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    return t


def caption(d, text):
    p = para(d, text, size=8.5, bold=True, color=MUTED, before=7, after=3)
    p.paragraph_format.keep_with_next = True


def note(d, text):
    para(d, text, size=8, color=MUTED, before=2, after=6)


def h1(d, text):
    p = para(d, text, size=12.5, bold=True, color=ACCENT, before=10, after=4)
    bdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "6"), ("space", "2"), ("color", "1D5C7C")):
        bot.set(qn(f"w:{k}"), v)
    bdr.append(bot)
    p._element.get_or_add_pPr().append(bdr)
    p.paragraph_format.keep_with_next = True


def sq(d, text):
    p = para(d, "□ " + text, size=10, bold=True, before=5, after=1.5, hanging=0.5)
    p.paragraph_format.keep_with_next = True


def ci(d, text):
    return para(d, "○ " + text, size=9.5, after=1.5, indent=0.4, hanging=0.45)


def da(d, text):
    return para(d, "- " + text, size=9.5, after=1.5, indent=0.95, hanging=0.35)


def formula(d, text):
    p = para(d, text, size=9.5, color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, before=3, after=4)
    p.runs[0].font.name = "Cambria Math"


def box(d, lines, title=None):
    t = d.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    borders(t, color="1D5C7C", size=8)
    cell_margins(t, 110, 110, 160, 160)
    c = t.cell(0, 0)
    c.width = Cm(17)
    shade(c, BOX_FILL)
    c.paragraphs[0].paragraph_format.space_after = Pt(0)
    if title:
        add_runs(c.paragraphs[0], title, 10.5, True, ACCENT)
    else:
        c._element.remove(c.paragraphs[0]._element)
    for ln in lines:
        para(c, "○ " + ln, size=9.5, after=2, hanging=0.45)


def picture(d, path, width=16.0):
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(path, width=Cm(width))


f1 = lambda v: f"{v:,.1f}"
f0 = lambda v: f"{v:,.0f}"
pstar = lambda p: "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else ""


# ── 자료·재계산 ──────────────────────────────────────────────────────
def excel_profile():
    wb = openpyxl.load_workbook(os.path.join(HERE, "data", "raw", "survey.xlsx"), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = [r for r in ws.iter_rows(values_only=True)]
    head = [h for h in rows[0] if h is not None]
    body = [r for r in rows[1:] if r[0] is not None]
    # 엑셀의 공실률 칸과 다시 계산한 값의 차이(검산)
    diff_a = max(abs((r[10] / r[9]) - r[11]) for r in body if r[9] and r[11] is not None)
    diff_u = max(abs((r[13] / r[12]) - r[14]) for r in body if r[12] and r[14] is not None)
    return {"sheet": ws.title, "n": len(body), "head": head, "diff_a": diff_a, "diff_u": diff_u}


def recompute(A):
    B = AN.load()
    cols = [c for c, _, _ in AN.FACTORS]
    D = B.dropna(subset=cols + ["y"]).copy()
    X = np.column_stack([np.ones(len(D))] + [D[c].values for c in cols])
    w = D["shop_area"].values
    beta, se, r2, n, k = AN.wls_hc1(X, D["y"].values, w)
    # 검산: analysis.json 의 계수와 같은가
    for i, c in enumerate(A["ols"]["coefs"], start=1):
        assert abs(beta[i] * AN.SCALE[c["key"]] - c["b"]) < 0.01, c["key"]
    # 다중공선성(VIF): 각 요인을 나머지 요인으로 회귀한 R² → 1/(1−R²)
    Z = D[cols].values
    vif = {}
    for j, c in enumerate(cols):
        others = np.column_stack([np.ones(len(Z)), np.delete(Z, j, axis=1)])
        b, *_ = np.linalg.lstsq(others, Z[:, j], rcond=None)
        res = Z[:, j] - others @ b
        r2j = 1 - res.var() / Z[:, j].var()
        vif[c] = 1 / (1 - r2j) if r2j < 1 else float("inf")
    # 종속변수(공실률) 분포
    y = D["y"].values
    ystat = {"mean": y.mean(), "wmean": np.average(y, weights=w), "sd": y.std(ddof=1), "med": np.median(y),
             "zero": int((y == 0).sum()), "ge50": int((y >= 50).sum()), "max": y.max()}
    # 요인별 기술통계
    desc = {}
    for c in cols:
        v = D[c].astype(float).values
        desc[c] = (v.mean(), v.std(ddof=1), v.min(), v.max())
    overall = {c: float(np.average(D[c], weights=w)) for c in cols}
    return {"B": B, "D": D, "beta": beta, "se": se, "r2": r2, "n": n, "k": k, "vif": vif, "y": ystat,
            "desc": desc, "overall": overall, "cols": cols}


# ── 그림 ────────────────────────────────────────────────────────────
def figures(A, R):
    os.makedirs(FIG, exist_ok=True)
    out = {}
    # 1 동별 공실률
    dd = sorted(A["by_dong"], key=lambda d: d["rate"])
    fig, ax = plt.subplots(figsize=(7.2, 3.7), dpi=200)
    ax.barh([f"{d['dong']} ({d['n']}동)" for d in dd], [d["rate"] for d in dd], color=[rate_color(d["rate"]) for d in dd], height=.65)
    ax.axvline(A["rate_area"], color="#15232A", lw=1, ls="--")
    ax.text(A["rate_area"] + .6, -0.9, f"행복도시 전체 {A['rate_area']}%", fontsize=7.5, va="bottom")
    for i, d in enumerate(dd):
        ax.text(d["rate"] + .6, i, f"{d['rate']:.1f}%", va="center", fontsize=7.5)
    ax.set_xlabel("공실률(면적 기준, %)", fontsize=8)
    ax.tick_params(labelsize=7.5)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_xlim(0, 58)
    fig.tight_layout()
    out["dong"] = os.path.join(FIG, "f1_dong.png")
    fig.savefig(out["dong"])
    plt.close(fig)

    # 2 구간별(경과연수·용도·상가비중·정류장)
    panels = [("age", "준공 경과연수"), ("use", "건물 주용도"), ("shop_share", "연면적 중 상가 비중"), ("d_stop", "버스정류장까지 거리")]
    fig, axs = plt.subplots(1, 4, figsize=(7.4, 2.5), dpi=200, sharey=True)
    for ax, (k, t) in zip(axs, panels):
        rows = A["bins"][k]
        ax.bar(range(len(rows)), [r["rate"] for r in rows], color=[rate_color(r["rate"]) for r in rows], width=.7)
        ax.axhline(A["rate_area"], color="#15232A", lw=.8, ls="--")
        ax.set_xticks(range(len(rows)))
        ax.set_xticklabels([r["label"].replace("제1종 근린생활", "1종\n근생").replace("제2종 근린생활", "2종\n근생").replace("주상복합", "주상\n복합").replace("업무시설", "업무") for r in rows],
                           fontsize=5.8, rotation=0)
        for i, r in enumerate(rows):
            ax.text(i, r["rate"] + 1, f"{r['rate']:.0f}\n({r['n']})", ha="center", fontsize=5.6, linespacing=.95)
        ax.set_title(t, fontsize=7.8)
        ax.spines[["top", "right"]].set_visible(False)
        ax.tick_params(axis="y", labelsize=6.5)
    axs[0].set_ylabel("공실률(%)", fontsize=7)
    axs[0].set_ylim(0, 62)
    fig.tight_layout(w_pad=.6)
    out["bins"] = os.path.join(FIG, "f2_bins.png")
    fig.savefig(out["bins"])
    plt.close(fig)

    # 3 회귀계수(95% 신뢰구간)
    co = A["ols"]["coefs"][::-1]
    fig, ax = plt.subplots(figsize=(7.2, 3.9), dpi=200)
    for i, c in enumerate(co):
        sig = c["p"] < 0.05
        colr = ("#B03A3A" if c["b"] > 0 else "#2E7D4F") if sig else "#9AA5A9"
        ax.plot([c["lo"], c["hi"]], [i, i], color=colr, lw=1.6)
        ax.plot(c["b"], i, "o", color=colr, ms=5)
        ax.text(31.5, i, f"{c['b']:+.1f}{pstar(c['p'])}", va="center", fontsize=7.2, fontweight="bold" if sig else "normal")
    ax.axvline(0, color="#15232A", lw=.8)
    ax.set_yticks(range(len(co)))
    ax.set_yticklabels([f"{c['name']} (/{c['unit']})" for c in co], fontsize=7)
    ax.set_xlim(-27, 36)
    ax.set_xlabel("공실률 변화(%p) — 다른 요인을 같게 놓았을 때", fontsize=8)
    ax.tick_params(axis="x", labelsize=7)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    out["ols"] = os.path.join(FIG, "f3_ols.png")
    fig.savefig(out["ols"])
    plt.close(fig)

    # 4 랜덤포레스트 중요도
    im = A["rf"]["importance"][::-1]
    fig, ax = plt.subplots(figsize=(7.0, 3.3), dpi=200)
    ax.barh([x["name"] for x in im], [x["imp"] for x in im], xerr=[x["sd"] for x in im], color="#1D5C7C", height=.6,
            error_kw={"lw": .8, "capsize": 2})
    ax.set_xlabel("순열 중요도(섞었을 때 설명력 R² 감소폭)", fontsize=8)
    ax.tick_params(labelsize=7)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    out["rf"] = os.path.join(FIG, "f4_rf.png")
    fig.savefig(out["rf"])
    plt.close(fig)

    # 5 핫스팟 지도(점) — 행정동 경계 위
    units = json.load(open(os.path.join(HERE, "assets", "sejong_units.json"), encoding="utf-8"))["units"]
    P = A["points"]
    fig, ax = plt.subplots(figsize=(7.0, 4.6), dpi=200)
    kx = math.cos(math.radians(36.5))
    for u in units:
        if u["kind"] != "dong":
            continue
        for r in u["rings"]:
            xs, ys = zip(*r)
            ax.fill([x * kx for x in xs], ys, color="#EEF2F1", ec="#B8C4C8", lw=.5)
        lx, ly = u.get("label") or (np.mean([p[0] for p in u["rings"][0]]), np.mean([p[1] for p in u["rings"][0]]))
        ax.text(lx * kx, ly, u["name"], fontsize=5.5, color="#7C8A8F", ha="center", zorder=2)
    for p in P:
        ax.scatter(p["lon"] * kx, p["lat"], s=6 + 40 * math.sqrt(p["area"] / max(q["area"] for q in P)),
                   c=rate_color(p["rate"]), ec="white", lw=.3, zorder=3)
        if abs(p["gi"]) > 1.96:
            ax.scatter(p["lon"] * kx, p["lat"], s=70, facecolors="none", ec="#B03A3A" if p["gi"] > 0 else "#2E7D4F",
                       lw=.7, zorder=4)
    xs = [p["lon"] * kx for p in P]
    ys = [p["lat"] for p in P]
    ax.set_xlim(min(xs) - .004, max(xs) + .004)
    ax.set_ylim(min(ys) - .006, max(ys) + .006)
    ax.set_aspect("equal")
    ax.axis("off")
    from matplotlib.lines import Line2D
    h = [Line2D([], [], marker="o", ls="", color=c, ms=5) for c in RATE_COLORS] + \
        [Line2D([], [], marker="o", ls="", mfc="none", mec="#B03A3A", ms=8), Line2D([], [], marker="o", ls="", mfc="none", mec="#2E7D4F", ms=8)]
    ax.legend(h, ["10% 미만", "10~20%", "20~30%", "30~40%", "40% 이상", "핫스팟(Gi*>1.96)", "콜드스팟(Gi*<-1.96)"],
              fontsize=6, loc="upper right", frameon=True, title="건물 공실률(원 크기=상가 면적)", title_fontsize=6.5)
    fig.tight_layout()
    out["hot"] = os.path.join(FIG, "f5_hot.png")
    fig.savefig(out["hot"])
    plt.close(fig)
    return out


# ── 본문 ────────────────────────────────────────────────────────────
FACTOR_DEF = {   # 요인: (자료 출처, 만드는 법)
    "age": ("엑셀 사용승인일자", "조사연도 6월 30일 − 사용승인일(년)"),
    "new5": ("엑셀 사용승인일자", "경과연수 ≤ 5년이면 1"),
    "type2": ("엑셀 건물 주용도", "'제2종근린생활시설'이면 1"),
    "office": ("엑셀 건물 주용도", "'업무시설'이 들어 있으면 1"),
    "mixed_res": ("엑셀 건물 주용도", "'공동주택'·'도시형생활주택'이 들어 있으면 1"),
    "shop_share": ("엑셀 연면적·상가면적", "상가면적 ÷ 연면적 × 100"),
    "log_area": ("엑셀 상가면적", "ln(상가면적)"),
    "d_stop": ("국토부 TAGO 버스정류소", "가장 가까운 정류장까지 직선거리(m)"),
    "d_brt": ("TAGO + BRT 노선(OSM)", "BRT 노선 40m 안 정류장까지 직선거리(m)"),
    "d_center": ("좌표 계산", "나성동 중심상업지(127.2600, 36.4870)까지 거리(m)"),
    "log_stores": ("소상공인 상가(상권)정보", "ln(1 + 반경 200m 안 점포 수)"),
    "u_pop_k": ("행안부 주민등록인구", "건물이 있는 행정동 인구(만 명)"),
    "u_transit": ("KTDB 여객 OD(2024)", "행정동 출발 통행 중 대중교통 비율(%)"),
}


# ── 멀티레벨(multilevel.py → data/multilevel.json) ─────────────────────
ML = None
LV_COL = {"L3": "#7A4FB5", "L2": "#1D5C7C", "L1": "#9AA5A9"}


def ml_step(ML, s):
    return next(x for x in ML["steps"] if x["step"] == s)


def ml_summary(ML):
    m0, m1, m3, m4 = (ml_step(ML, s) for s in ("M0", "M1", "M3", "M4"))
    f4 = {c["key"]: c for c in next(x for x in ML["full"] if x["step"] == "M4")["coefs"]}
    return [f"**멀티레벨 분석**(건물 ⊂ 소권역 ⊂ 행정동): 공실 차이의 {m0['icc']['L2'] * 100:.0f}%가 소권역(상권 블록) 간, "
            f"{m0['icc']['L3'] * 100:.0f}%만 행정동 간 차이다. 소권역 간 차이는 건물 요인이 {m1['pcv']['L2'] * 100:.0f}%, "
            f"배후 수요·공급까지 {m3['pcv']['L2'] * 100:.0f}%, 행정동 여건까지 {m4['pcv']['L2'] * 100:.0f}%를 설명한다. "
            f"**반경 500m 세대당 상가면적**이 2배면 공실률 +{f4['log2_sph']['b']:.1f}%p(287동 모형, p={f4['log2_sph']['p']:.3f})."]


def ml_figure(ML):
    """단계별: 남은 분산(빈 모형 = 100)을 수준별로 쌓아 그린다."""
    steps = ML["steps"]
    tot0 = sum(steps[0]["var"].values())
    fig, ax = plt.subplots(figsize=(7.2, 2.9), dpi=200)
    names = [f"{s['step']}\n{s['name'].replace('+ ', '+')}" for s in steps]
    left = np.zeros(len(steps))
    for k, lab in (("L3", "행정동 간"), ("L2", "소권역(상권 블록) 간"), ("L1", "같은 소권역 안 건물 간")):
        v = np.array([s["var"][k] / tot0 * 100 for s in steps])
        ax.barh(range(len(steps)), v, left=left, color=LV_COL[k], height=.62, label=lab)
        for i, (l, w) in enumerate(zip(left, v)):
            if w >= 5:
                ax.text(l + w / 2, i, f"{w:.0f}", ha="center", va="center", fontsize=7, color="white", fontweight="bold")
        left += v
    for i, s in enumerate(steps):
        ax.text(left[i] + 1, i, f"남은 분산 {left[i]:.0f}  ·  R²m {s['r2m']:.2f}", va="center", fontsize=6.8, color="#4B5B62")
    ax.set_yticks(range(len(steps)))
    ax.set_yticklabels(names, fontsize=6.6)
    ax.invert_yaxis()
    ax.set_xlim(0, 135)
    ax.set_xlabel("빈 모형(M0)의 전체 분산 = 100", fontsize=7.5)
    ax.tick_params(axis="x", labelsize=7)
    ax.legend(fontsize=6.5, loc="lower right", frameon=False, ncol=3, bbox_to_anchor=(1, 1.0))
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    p = os.path.join(FIG, "f6_multilevel.png")
    fig.savefig(p)
    plt.close(fig)
    return p


def ml_chapter(doc, ML, F):
    st = {s["step"]: s for s in ML["steps"]}
    m0, m4 = st["M0"], st["M4"]
    f4 = {c["key"]: c for c in next(x for x in ML["full"] if x["step"] == "M4")["coefs"]}
    g = {c["key"]: c for c in ML["glmm"]["coefs"]}
    h1(doc, "10. 멀티레벨 분석: 건물·상권 블록·행정동")
    sq(doc, "왜 멀티레벨인가")
    ci(doc, "건물은 같은 소권역(동 세부 상권 블록)·같은 행정동 안에서 서로 닮는다. 이를 무시한 회귀는 표본이 독립이라고 보고 유의성을 부풀리고, "
            "공실 차이가 **어느 수준에서** 생기는지(건물 탓인가 동네 탓인가)를 말해 주지 못한다.")
    ci(doc, f"구조: 건물(수준 1, {ML['n']}동) ⊂ 소권역(수준 2, {ML['n_L2']}곳) ⊂ 행정동(수준 3, {ML['n_L3']}곳 — 도담·어진동은 인구 통계가 합쳐져 한 곳).")
    sq(doc, "추가한 변수(이미 받아 둔 공공데이터만 사용)")
    caption(doc, "표 10. 추가 변수와 기술통계(분석 표본)")
    src = {"floors": "건축물대장 표제부(조사 건물 필지)", "park_per_100": "건축물대장 주차대수 ÷ 상가면적",
           "sh_food": "소상공인 상가정보(도로명 주소로 건물 연결)", "sh_edu": "〃", "sh_med": "〃", "diversity": "〃 업종 대분류 엔트로피(0~1)",
           "franchise": "〃 지점명 있는 점포", "sh_3up": "〃 층 정보 있는 점포 중", "shop_per_hh": "조사 상가면적 ÷ 공동주택 세대(필지 좌표화)",
           "hh500": "건축물대장 공동주택 세대수", "apt_ppm": "아파트 실거래(최근 24개월, 세대 가중, 반경 1km)",
           "u_chg": "주민등록 인구(생활권 상황판)", "u_old": "〃", "u_emp": "SGIS 종사자 ÷ 주민등록 인구"}
    table(doc, ["변수", "자료·만드는 법", "n", "평균(표준편차)", "범위"],
          [[d["name"], src.get(d["key"], ""), f"{d['n']}", f"{d['mean']:,.2f} ({d['sd']:,.2f})", f"{d['min']:,.1f} ~ {d['max']:,.1f}"]
           for d in ML["raw_desc"]], [4.2, 6.0, 1.0, 3.0, 2.8], size=7.6, align_right=(2, 3, 4))
    note(doc, "점포 '개수'는 공실이 많으면 저절로 줄어드는 공실의 거울이라 쓰지 않고 업종·층 '구성 비율'만 썼다. 상가정보는 현재 시점이고 공실 조사는 "
              "2023·2025년이라 시점이 다르다. 주차·층 정보가 없는 건물은 중앙값으로 채웠다.")
    sq(doc, "분석 과정")
    ci(doc, f"**표본**: 입점 점포가 하나도 연결되지 않은 {ML['excluded']['n']}동은 공실률이 {ML['excluded']['rate']}%로, 점포가 없는 것 자체가 공실의 결과다. "
            f"업종 구성을 쓰는 주 모형은 {ML['n']}동으로 하고, {ML['n_full']}동 전체로 같은 단계(업종 묶음 제외)를 다시 추정해 결론을 대조했다.")
    ci(doc, "**모형**: 선형 혼합모형. 공실률(면적, %)을 고정효과(요인)와 행정동·소권역 무작위절편, 건물 오차로 나눈다. REML 로 추정하고, 단계 간 비교(AIC)는 최대우도로 다시 맞췄다.")
    formula(doc, "공실률ᵢⱼₖ = β₀ + Σ βₚ·xₚ + uₖ(행정동) + vⱼₖ(소권역) + εᵢⱼₖ,   uₖ ~ N(0, σ²₃), vⱼₖ ~ N(0, σ²₂), ε ~ N(0, σ²₁)")
    ci(doc, "**급내상관(ICC)**: 전체 분산 중 각 수준이 차지하는 몫. 빈 모형(M0)에서 구하면 '공실 차이가 어느 수준에서 나는가'가 된다.")
    formula(doc, "ICC(행정동) = σ²₃ / (σ²₃ + σ²₂ + σ²₁),   ICC(소권역) = σ²₂ / (σ²₃ + σ²₂ + σ²₁)")
    ci(doc, "**설명된 분산 비율(PCV)**: 요인을 넣었을 때 각 수준의 분산이 빈 모형보다 몇 % 줄었는지. 어느 수준의 차이를 어떤 요인 묶음이 설명하는지 보여 준다.")
    formula(doc, "PCV(수준) = 1 − σ²(모형) / σ²(빈 모형)")
    ci(doc, "**단계**: M0 빈 모형 → M1 기존 건물 요인 8개 → M2 건물 내부(층·주차·업종 구성) 8개 → M3 반경 수요·공급 2개 → M4 행정동 여건 3개.")
    ci(doc, "**설명력**: Nakagawa R² — 고정효과만의 몫(R²m)과 무작위효과까지 합한 몫(R²c).")
    ci(doc, f"**교차 확인**: 점포 한 칸(호)을 공실 1/0으로 놓은 이항 혼합모형({ML['glmm']['n_units']:,}칸, 건물·소권역·행정동 무작위절편, 변분 베이즈). "
            "큰 건물이 더 큰 무게를 갖는다. 변분 베이즈는 불확실성을 작게 잡는 경향이 있어 **방향 일치 여부**만 본다.")
    ci(doc, f"**공간 점검**: M4 잔차의 Moran's I(반경 500m) = {ML['moran']['I']} (p = {ML['moran']['p']}) → 남은 공간 자기상관이 없다.")
    sq(doc, "결과 ① 공실 차이는 어느 수준에서 나는가")
    caption(doc, f"표 11. 단계별 수준별 분산·설명력(n = {ML['n']})")
    rows = []
    for s in ML["steps"]:
        p = s["pcv"]
        rows.append([f"{s['step']} {s['name']}", f"{s['k']}", f"{s['aic']:,.1f}", f"{s['var']['L3']:.1f}", f"{s['var']['L2']:.1f}", f"{s['var']['L1']:.1f}",
                     "—" if s["step"] == "M0" else f"{(p['L2'] or 0) * 100:.0f}%", "—" if s["step"] == "M0" else f"{(p['L1'] or 0) * 100:.0f}%",
                     f"{s['r2m']:.2f} / {s['r2c']:.2f}"])
    table(doc, ["단계", "요인", "AIC", "σ²행정동", "σ²소권역", "σ²건물", "소권역 PCV", "건물 PCV", "R²m / R²c"], rows,
          [4.3, 1.0, 1.6, 1.6, 1.6, 1.6, 1.7, 1.6, 2.0], size=7.6, align_right=(1, 2, 3, 4, 5, 6, 7, 8))
    note(doc, "σ² 는 공실률(%p)의 분산. 행정동 분산은 0에 가까워 PCV 가 불안정하므로 표에서 뺐다. AIC 는 작을수록 좋다(최대우도 기준).")
    caption(doc, "그림 6. 단계별 남은 분산(빈 모형 = 100)")
    picture(doc, F["ml"], 15.5)
    ci(doc, f"빈 모형에서 공실 차이의 **{m0['icc']['L2'] * 100:.0f}%가 소권역 간**, {m0['icc']['L3'] * 100:.0f}%가 행정동 간, 나머지 "
            f"{(1 - m0['icc']['L2'] - m0['icc']['L3']) * 100:.0f}%가 같은 소권역 안 건물 간 차이다. **공실은 행정동보다 상권 블록 단위에서 갈린다.**")
    ci(doc, f"소권역 간 차이는 건물 연식·용도(M1)로 {st['M1']['pcv']['L2'] * 100:.0f}%, 반경 수요·공급(M3)까지 {st['M3']['pcv']['L2'] * 100:.0f}%, "
            f"행정동 여건(M4)까지 {m4['pcv']['L2'] * 100:.0f}%가 설명된다. 블록 간 차이는 대부분 '어떤 건물이 얼마나 공급됐나'로 설명된다.")
    ci(doc, f"반면 같은 블록 안 건물 간 차이는 {m4['pcv']['L1'] * 100:.0f}%만 줄었다 — 임대료·관리·건물주 전략 등 자료에 없는 건물 고유 요인이 크다.")
    sq(doc, "결과 ② 요인별 효과(M4, 모든 수준을 함께)")
    caption(doc, "표 12. M4 계수(공실률 %p)와 대조")
    rows, bold = [], []
    for i, c in enumerate(m4["coefs"]):
        f = f4.get(c["key"])
        gg = g.get(c["key"])
        if c["p"] < 0.05:
            bold.append(i)
        rows.append([c["level"], c["name"], c["unit"], f"{c['b']:+.2f}", f"[{c['lo']:+.1f}, {c['hi']:+.1f}]", f"{c['p']:.3f}{pstar(c['p'])}",
                     (f"{f['b']:+.1f}{pstar(f['p'])}" if f else "—"), (f"{gg['or']:.2f}" if gg else "—")])
    table(doc, ["수준", "요인", "단위", "계수", "95% 신뢰구간", "p", f"{ML['n_full']}동", "이항 OR"], rows,
          [1.3, 4.4, 1.6, 1.6, 2.8, 1.5, 1.6, 1.6], size=7.6, align_right=(3, 5, 6, 7), align_center=(0, 2, 4), bold_rows=tuple(bold))
    note(doc, f"굵은 줄 = 주 모형에서 유의(p<0.05). '{ML['n_full']}동' = 업종 묶음을 뺀 전체 표본 M4 계수(별표 = 유의). "
              "'이항 OR' = 점포 한 칸 단위 이항 혼합모형의 승산비(1보다 크면 공실 쪽).")
    agree = sum(1 for c in m4["coefs"] if c["key"] in g and (c["b"] > 0) == (g[c["key"]]["or"] > 1))
    sph = next(c for c in m4["coefs"] if c["key"] == "log2_sph")
    ci(doc, f"**연식·용도는 그대로 핵심**: 준공 5년 이하 {next(c for c in m4['coefs'] if c['key'] == 'new5')['b']:+.1f}%p, "
            f"제2종 근생 {next(c for c in m4['coefs'] if c['key'] == 'type2')['b']:+.1f}%p로, 동네 차이를 통제해도 유지된다.")
    ci(doc, f"**배후 세대 대비 상가 총량**: 반경 500m 세대당 상가면적이 2배면 {sph['b']:+.1f}%p(p={sph['p']:.2f}), 전체 {ML['n_full']}동에서는 "
            f"{f4['log2_sph']['b']:+.1f}%p(p={f4['log2_sph']['p']:.3f})로 유의하다. '상가 총량 과잉' 가설을 직접 뒷받침한다.")
    div = next(c for c in m4["coefs"] if c["key"] == "div01")
    ci(doc, f"**업종 다양도**: 다양도가 0.1 높으면 {div['b']:+.1f}%p(p={div['p']:.3f}). 한두 업종에 치우친 건물이 더 비어 있다 — 업종 구성이 다양한 건물일수록 공실이 낮다(상관이며, 공실이 업종 구성을 바꾸는 역방향도 가능).")
    old_, emp_ = (next(c for c in m4["coefs"] if c["key"] == k) for k in ("old1", "emp01"))
    ci(doc, f"**행정동 여건**: 고령 비중 1%p당 {old_['b']:+.1f}%p, 직주비 0.1당 {emp_['b']:+.1f}%p. 낮 인구(종사자)가 많은 업무형 동은 주말·저녁 수요가 비어 "
            "근린상가가 버티기 어려운 것으로 보인다.")
    ci(doc, f"이항 혼합모형과 비교하면 {len(m4['coefs'])}개 요인 중 {agree}개의 방향이 같다.")
    sq(doc, "한계")
    ci(doc, f"행정동이 {ML['n_L3']}곳뿐이라 행정동 분산 추정이 불안정하다(대부분 0에 가깝게 추정). 행정동 효과는 고정효과(인구 증감·고령·직주비)로 읽는 것이 안전하다.")
    ci(doc, "입점 업종(현재)과 공실(2023·2025)의 시점이 다르고, 업종 구성은 공실과 서로 영향을 주고받을 수 있다(역인과).")
    ci(doc, "카드 매출·생활인구·임대료는 쓰지 않았다(별도 확보 필요). 이 자료가 붙으면 같은 틀에서 수요 쪽 설명이 강해진다.")


def build(A, R, X, F):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.left_margin = sec.right_margin = Cm(2.0)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_runs(fp, "세종연구원 연구모임 · 행복도시 상가공실 분석(검토용 초안)", 7.5, color=MUTED)

    co = {c["key"]: c for c in A["ols"]["coefs"]}
    age = {r["label"]: r for r in A["bins"]["age"]}
    use = {r["label"]: r for r in A["bins"]["use"]}
    zones = {z["zone"]: z for z in A["by_zone"]}
    dmax = max(A["by_dong"], key=lambda d: d["rate"])
    years = A["survey_years"]
    n23 = sum(1 for p in A["points"] if p["survey"] == years[0])

    para(doc, "세종연구원 · 연구모임 검토자료", 9, color=MUTED, after=2)
    para(doc, "행복도시 상가공실 실태와 원인 분석", 18, True, ACCENT, after=2)
    para(doc, "집합상가 전수조사 자료로 본 생활권·동·건물별 공실과 공실 요인", 11, color=MUTED, after=8)
    box(doc, [
        f"행복도시 집합상가 **{A['n_buildings']}동, {A['units']:,}호**를 조사한 결과 상가 공실률은 면적 기준 **{A['rate_area']}%**, "
        f"호수 기준 **{A['rate_units']}%**다(공실 {A['vac_area']:,}㎡ / 상가 {A['shop_area']:,}㎡).",
        f"생활권별로는 3생활권 {zones['3']['rate']}%·4생활권 {zones['4']['rate']}%가 높고 2생활권 {zones['2']['rate']}%가 가장 낮다. "
        f"동별로는 **{dmax['dong']} {dmax['rate']}%**가 가장 높다.",
        f"13개 요인을 함께 넣은 회귀분석(상가면적 가중, R² {A['ols']['r2']})에서 공실을 유의하게 높이는 요인은 "
        f"**준공 5년 이하(+{co['new5']['b']:.1f}%p)**와 **제2종 근린생활시설(+{co['type2']['b']:.1f}%p)**, 낮추는 요인은 "
        f"**상가 비중(10%p당 {co['shop_share']['b']:+.1f}%p)**이다. 정류장·중심지 거리 등 입지 요인은 유의하지 않다.",
        f"머신러닝(랜덤포레스트)으로 교차 확인해도 **건물 연식**이 가장 중요한 요인이다. 공실은 입지보다 **공급 구조(연식·용도·상가 총량)**에서 갈린다.",
        f"공실이 모여 있는 핫스팟은 {A['hot']['hot']}동으로, 가로·블록 단위의 집중 대응 대상 구역을 고르는 근거가 된다.",
    ] + (ml_summary(ML) if ML else []), title="요 약")

    # 1. 개요
    h1(doc, "1. 분석 개요")
    sq(doc, "목적")
    ci(doc, "행복도시 집합상가 전수조사 결과를 건물 단위로 정리해 공실이 **어디에** 얼마나 있는지 보이고, **왜** 생기는지 건물·입지 요인으로 나눠 본다.")
    ci(doc, "결과는 '세종 행복도시 상가공실 상황판'(ysong86.github.io/vacant_sejong, 비밀번호)에 2D·3D 지도와 함께 실었다. 이 보고서는 그 분석의 자료·방법·결과를 정리한 것이다.")
    sq(doc, "분석 흐름")
    table(doc, ["단계", "내용", "결과물"], [
        ["① 자료 정리", "조사 엑셀을 읽어 공실률 재계산·경과연수 산출, 주소 → 좌표(브이월드)", f"건물 {A['n_buildings']}동 좌표 100%"],
        ["② 요인 결합", "건물 좌표에 정류장·점포·인구·통행 등 공공데이터 결합", "건물별 13개 요인"],
        ["③ 현황 집계", "생활권·동·건물 공실률(면적 합계 기준)", "표 3·4, 그림 1"],
        ["④ 단순 비교", "요인 구간별 공실률", "표 5, 그림 2"],
        ["⑤ 회귀분석", "상가면적 가중 최소제곱 + 강건 표준오차", "표 6, 그림 3"],
        ["⑥ 교차 확인", "랜덤포레스트 + 순열 중요도", "표 7, 그림 4"],
        ["⑦ 동별 분해", "동 공실률이 평균과 다른 이유를 요인별로 분해", "표 8·9"],
        ["⑧ 공간 군집", "Getis-Ord Gi* 핫스팟(반경 500m)", "그림 5"],
    ] + ([["⑨ 멀티레벨", "건물 ⊂ 소권역 ⊂ 행정동 3수준 혼합모형, 건물 내부·배후 수요·행정동 여건 추가", "표 10~12, 그림 6"]] if ML else []),
        [2.4, 9.6, 5.0])

    # 2. 조사 자료
    h1(doc, "2. 조사 자료(엑셀)와 정리 과정")
    sq(doc, "자료 개요")
    ci(doc, f"파일: 집합상가 전수조사 엑셀(시트 '{X['sheet']}'), 건물 {X['n']}동 × {len(X['head'])}개 항목. "
            f"조사 시점은 {years[0]}년 {n23}동, {years[-1]}년 {X['n'] - n23}동으로 섞여 있다.")
    caption(doc, "표 1. 엑셀 항목과 분석에서의 쓰임")
    table(doc, ["엑셀 항목", "내용", "분석에서의 쓰임"], [
        ["NO · 건물명", "일련번호 · 건물 이름", "건물 식별, 지도 표시"],
        ["동 · 동 세부", "법정동 · 동 안 소권역(예: 고운동 북측)", "동별 집계, 생활권 배정"],
        ["동 코드 · 건물세부코드", "조사 관리 코드", "사용 안 함"],
        ["주소", "도로명 주소", "좌표 변환(브이월드) → 입지 요인 계산"],
        ["건물 주용도", "제1종·제2종 근린생활시설, 업무시설 등", "용도 요인(제2종·업무·주상복합)"],
        ["연면적", "건물 전체 바닥면적(㎡)", "상가 비중 = 상가면적 ÷ 연면적"],
        ["상가면적 · 공실_면적", "상가로 쓰는 면적 · 그중 빈 면적(㎡)", "**공실률(면적) = 공실면적 ÷ 상가면적** — 종속변수, 회귀 가중치"],
        ["호수 · 공실_호수", "점포 수 · 빈 점포 수", "공실률(호수) — 보조 지표"],
        ["공실률_면적 · 공실률_호수", "엑셀에서 계산된 값", "재계산 값과 대조(검산)"],
        ["비고(1)", "조사연도(2023/2025)", "경과연수 기준일, 시점 구분"],
        ["사용승인일자", "준공(사용승인) 날짜", "경과연수, 준공 5년 이하"],
        ["비고(2)", "경과연수 구간(예: 10년초과)", "참고(직접 다시 계산해 사용)"],
    ], [3.6, 5.4, 8.0])
    sq(doc, "정리 과정")
    ci(doc, f"**공실률 검산**: 엑셀의 공실률 칸을 공실면적 ÷ 상가면적으로 다시 계산해 대조했다. 최대 차이 {X['diff_a']:.1e}(면적), "
            f"{X['diff_u']:.1e}(호수)로 엑셀 값과 일치한다.")
    ci(doc, "**경과연수**: 사용승인일자부터 조사연도 6월 30일까지의 햇수로 다시 계산했다. 2025년 조사분은 2025년 6월 30일 기준이다.")
    ci(doc, f"**좌표 변환**: 도로명 주소를 브이월드 주소 API로 좌표화했다(실패 시 '동 + 건물명'으로 재시도). {A['n_buildings']}동 모두 좌표를 얻었다.")
    ci(doc, "**생활권 배정**: 동 → 생활권(1: 고운·아름·종촌·도담·어진 / 2: 다정·새롬·한솔·나성 / 3: 대평·보람·소담 / 4: 반곡·집현 / 6: 해밀·산울). 5생활권은 조사 건물이 없다.")
    ci(doc, "**집계 원칙**: 생활권·동 공실률은 건물 공실률의 단순 평균이 아니라 **공실면적 합계 ÷ 상가면적 합계**로 계산했다(큰 건물의 무게를 반영).")

    # 3. 요인
    h1(doc, "3. 분석 요인")
    ci(doc, "공실에 영향을 줄 수 있는 요인을 건물 특성(엑셀에서 직접)과 입지·수요 특성(공공데이터 결합)으로 나눠 13개를 만들었다. "
            "입지 요인은 생활권 재구조화 상황판이 받아 둔 공공데이터를 그대로 썼다.")
    caption(doc, "표 2. 분석 요인의 정의와 기술통계(조사 건물 287동, 단순 통계)")
    rows = []
    for c, name, unit in AN.FACTORS:
        src, how = FACTOR_DEF[c]
        m, s, lo, hi = R["desc"][c]
        dummy = c in ("new5", "type2", "office", "mixed_res")
        stat = f"{m * 100:.0f}% 해당" if dummy else f"{m:,.1f} ({s:,.1f})"
        rng = "0/1" if dummy else f"{lo:,.1f} ~ {hi:,.1f}"
        rows.append([name, src, how, stat, rng])
    table(doc, ["요인", "자료", "만드는 법", "평균(표준편차)", "범위"], rows, [3.4, 3.3, 5.3, 2.6, 2.4], size=7.8, align_right=(3, 4))
    y = R["y"]
    note(doc, f"종속변수 공실률(면적): 건물 평균 {y['mean']:.1f}%, 상가면적 가중 평균 {y['wmean']:.1f}%, 중앙값 {y['med']:.1f}%, 표준편차 {y['sd']:.1f}%p. "
              f"공실 0% 건물 {y['zero']}동, 50% 이상 {y['ge50']}동, 최대 {y['max']:.1f}%. 로그 요인은 '2배' 늘 때의 효과로 바꿔 읽는다.")

    # 4. 현황
    h1(doc, "4. 공실 현황")
    caption(doc, "표 3. 생활권별 공실률")
    zrow = []
    for z in sorted(A["by_zone"], key=lambda z: z["zone"]):
        zrow.append([f"{z['zone']}생활권", f"{z['n']}", f0(z["area"]), f0(z["vac_area"]), f1(z["rate"]), f1(z["rate_units"])])
    zrow.append(["행복도시 전체", f"{A['n_buildings']}", f0(A["shop_area"]), f0(A["vac_area"]), f1(A["rate_area"]), f1(A["rate_units"])])
    table(doc, ["생활권", "건물(동)", "상가면적(㎡)", "공실면적(㎡)", "공실률 면적(%)", "공실률 호수(%)"], zrow,
          [3.0, 2.0, 3.2, 3.2, 2.8, 2.8], align_right=(1, 2, 3, 4, 5), bold_rows=(len(zrow) - 1,))
    note(doc, "6생활권은 조사 건물이 2동뿐이라 생활권 값으로 읽기 어렵다. 공실률(면적) = Σ공실면적 ÷ Σ상가면적 × 100.")
    caption(doc, "그림 1. 동별 공실률(면적 기준) — 점선은 행복도시 전체")
    picture(doc, F["dong"], 15.5)
    caption(doc, "표 4. 동별 공실률과 건물 특성")
    drows = [[d["dong"], f"{d['zone']}", f"{d['n']}", f"{d['units']:,}", f1(d["rate"]), f1(d["rate_units"]), f1(d["age"]), f"{d['new5_share']}"]
             for d in sorted(A["by_dong"], key=lambda d: -d["rate"])]
    table(doc, ["동", "생활권", "건물", "호수", "공실률 면적(%)", "공실률 호수(%)", "평균 경과연수", "5년 이하 비중(%)"], drows,
          [2.2, 1.5, 1.4, 1.9, 2.4, 2.4, 2.4, 2.8], size=8, align_right=(2, 3, 4, 5, 6, 7), align_center=(1,))
    note(doc, "평균 경과연수와 5년 이하 비중은 상가면적 가중. 공실률이 높은 대평·반곡·소담동은 5년 이하 신축 비중이 51~98%로 높다.")

    # 5. 구간 비교
    h1(doc, "5. 요인별 단순 비교")
    ci(doc, "회귀분석에 앞서 요인을 몇 구간으로 나눠 구간마다 공실률(면적 합계 기준)을 비교했다. 다른 요인을 통제하지 않은 값이라 방향을 보는 용도다.")
    caption(doc, "그림 2. 요인 구간별 공실률(%) — 괄호는 건물 수, 점선은 전체")
    picture(doc, F["bins"], 16.5)
    caption(doc, "표 5. 요인 구간별 공실률(면적 기준, %)")
    labels = {"age": "준공 경과연수", "use": "건물 주용도", "shop_share": "상가 비중", "d_stop": "버스정류장 거리",
              "d_brt": "BRT 정류장 거리", "d_center": "나성동 중심 거리", "stores": "주변 200m 점포"}
    brow = [[labels[k], " · ".join(f"{r['label']} **{r['rate']:.1f}**({r['n']})" for r in A["bins"][k])] for k in labels]
    table(doc, ["요인", "구간별 공실률(건물 수)"], brow, [3.2, 13.8], size=8)
    ci(doc, f"준공 3년 이하 {age['3년 이하']['rate']}% → 5~7년 {age['5~7년']['rate']}% → 9년 초과 {age['9년 초과']['rate']}%로 연식이 쌓일수록 공실이 빠르게 준다.")
    ci(doc, f"제2종 근린생활시설 {use['제2종 근린생활']['rate']}%로 제1종({use['제1종 근린생활']['rate']}%)의 두 배다. 상가 비중이 낮은(30% 이하) 건물도 높다.")
    ci(doc, "정류장에서 200m를 넘으면 공실이 높지만 해당 건물이 16동뿐이고, BRT·중심지 거리는 방향이 일정하지 않다.")

    # 6. 회귀
    h1(doc, "6. 회귀분석: 다른 요인을 같게 놓았을 때의 순효과")
    sq(doc, "분석 과정")
    ci(doc, "**모형**: 건물 i의 공실률(%)을 13개 요인의 선형 결합으로 설명한다.")
    formula(doc, "공실률ᵢ = β₀ + β₁·경과연수ᵢ + β₂·5년이하ᵢ + … + β₁₃·대중교통분담률ᵢ + εᵢ")
    ci(doc, "**가중 최소제곱(WLS)**: 상가 5,000㎡ 건물과 300㎡ 건물이 같은 무게를 가지면 전체 공실면적을 잘 설명하지 못하므로, "
            "**상가면적을 가중치**로 두고 Σ wᵢ(yᵢ − ŷᵢ)²을 최소화했다. 이렇게 하면 회귀가 표 3의 면적 합계 공실률과 같은 기준이 된다.")
    formula(doc, "β̂ = (XᵀWX)⁻¹ XᵀWy,   W = diag(상가면적ᵢ)")
    ci(doc, "**강건 표준오차(HC1)**: 공실률의 흩어짐이 건물 규모·유형마다 달라(이분산) 일반 표준오차는 유의성을 과대평가할 수 있다. "
            "White(HC1) 방식으로 표준오차를 구하고, t분포로 95% 신뢰구간과 p값을 냈다.")
    formula(doc, f"Var(β̂) = (XᵀWX)⁻¹ [Σ wᵢ êᵢ² xᵢxᵢᵀ] (XᵀWX)⁻¹ × n/(n−k),   n = {R['n']}, k = {R['k']}")
    ci(doc, "**단위 환산**: 계수를 읽기 쉬운 단위로 바꿨다 — 거리 100m(중심지는 1km), 상가 비중 10%p, 로그 요인은 ×ln2(= 2배 늘 때), 인구 1만 명.")
    vmax = max(R["vif"].items(), key=lambda kv: kv[1])
    vname = {c: n for c, n, _ in AN.FACTORS}
    ci(doc, f"**다중공선성 점검(VIF)**: 요인끼리 너무 겹치면 계수가 불안정해진다. 각 요인을 나머지로 회귀해 VIF = 1/(1−R²)를 구했다. "
            f"가장 큰 값은 {vname[vmax[0]]} {vmax[1]:.1f}이고, 경과연수({R['vif']['age']:.1f})와 5년 이하({R['vif']['new5']:.1f})가 서로 겹친다. "
            "10을 넘는 요인은 " + ("없다." if vmax[1] < 10 else "있어 해석에 주의가 필요하다."))
    sq(doc, "결과")
    caption(doc, f"표 6. 가중 회귀 결과(n = {A['ols']['n']}, R² = {A['ols']['r2']})")
    orow = []
    for c, name, unit in AN.FACTORS:
        x = co[c]
        orow.append([name, unit, f"{x['b']:+.2f}", f"[{x['lo']:+.1f}, {x['hi']:+.1f}]", f"{x['p']:.3f}{pstar(x['p'])}", f"{R['vif'][c]:.1f}"])
    sig_rows = tuple(i for i, (c, _, _) in enumerate(AN.FACTORS) if co[c]["p"] < 0.05)
    table(doc, ["요인", "단위", "계수(%p)", "95% 신뢰구간", "p값", "VIF"], orow, [5.0, 1.6, 2.2, 3.4, 2.2, 1.6],
          size=8, align_right=(2, 4, 5), align_center=(1, 3), bold_rows=sig_rows)
    note(doc, "굵은 줄 = 5% 수준에서 유의. * p<0.05, ** p<0.01, *** p<0.001. 계수는 해당 단위만큼 요인이 커질 때 공실률(%p) 변화.")
    caption(doc, "그림 3. 회귀계수와 95% 신뢰구간(빨강 = 공실을 높임, 초록 = 낮춤, 회색 = 유의하지 않음)")
    picture(doc, F["ols"], 15.5)
    sq(doc, "해석")
    ci(doc, f"**준공 5년 이하**: 다른 조건이 같아도 공실률이 **{co['new5']['b']:+.1f}%p** 높다(p={co['new5']['p']:.3f}). "
            f"경과연수 자체(1년당 {co['age']['b']:+.2f}%p)는 유의하지 않아, 연식 효과는 '서서히'가 아니라 **준공 초기 몇 년에 몰려** 있다.")
    ci(doc, f"**제2종 근린생활시설**: **{co['type2']['b']:+.1f}%p**(p={co['type2']['p']:.4f}). 상층부·대형 점포 위주의 공급이 수요와 어긋난 것으로 보인다.")
    ci(doc, f"**상가 비중**: 연면적 중 상가 비중이 10%p 높을수록 **{co['shop_share']['b']:+.1f}%p**(p={co['shop_share']['p']:.3f}). 업무·주거가 주된 건물의 부속 상가가 더 빈다.")
    ci(doc, "**입지 요인**(버스·BRT 정류장 거리, 나성동 중심 거리, 주변 점포, 행정동 인구·대중교통)은 방향은 예상과 같지만 모두 유의하지 않다. "
            "조사 건물 대부분이 정류장 200m 안에 있어 입지 차이 자체가 작다.")
    ci(doc, f"모형 설명력 R² {A['ols']['r2']}: 건물 간 공실률 차이의 약 43%를 13개 요인이 설명한다. 나머지는 임대료·관리·업종 구성 등 자료에 없는 요인이다.")

    # 7. RF
    h1(doc, "7. 머신러닝 교차 확인(랜덤포레스트)")
    sq(doc, "분석 과정")
    ci(doc, "회귀는 요인 효과를 직선·더하기로 가정한다. 비선형(예: 준공 3년까지만 높음)과 요인 간 상호작용을 허용하는 랜덤포레스트로 같은 13개 요인을 다시 분석했다.")
    da(doc, "의사결정나무 600개, 잎당 최소 4개 건물, 상가면적 가중, 난수 고정(random_state=7)으로 재현 가능하게 했다.")
    da(doc, f"**OOB R²**: 각 나무가 학습에 쓰지 않은 건물(out-of-bag)로 예측력을 쟀다 → **{A['rf']['oob_r2']}**. 처음 보는 건물의 공실률을 약 21% 설명한다.")
    da(doc, "**순열 중요도**: 한 요인의 값을 건물끼리 무작위로 섞었을 때 예측력(R²)이 얼마나 떨어지는지를 30회 반복해 평균·표준편차를 냈다. 많이 떨어질수록 중요한 요인이다.")
    caption(doc, "표 7. 랜덤포레스트 순열 중요도(상위 8개)")
    table(doc, ["순위", "요인", "중요도(R² 감소)", "표준편차", "회귀 결과"], [
        [f"{i}", x["name"], f"{x['imp']:.2f}", f"{x['sd']:.2f}",
         (f"{co[x['key']]['b']:+.1f}%p" + (" (유의)" if co[x["key"]]["p"] < 0.05 else ""))]
        for i, x in enumerate(A["rf"]["importance"][:8], start=1)], [1.4, 5.2, 3.2, 2.4, 3.8], size=8.5, align_center=(0,), align_right=(2, 3))
    caption(doc, "그림 4. 요인 중요도(오차막대 = 30회 반복 표준편차)")
    picture(doc, F["rf"], 14.5)
    ci(doc, "두 방법 모두 **건물 연식**을 가장 중요한 요인으로 꼽는다. 회귀에서 유의하지 않던 경과연수가 1위인 것은 효과가 비선형(초기 몇 년에 집중)이기 때문이고, 회귀의 '5년 이하' 더미와 같은 이야기다.")
    ci(doc, "상가 규모·상가 비중이 그다음이고, 정류장 거리는 비선형으로 일부 기여한다. 제2종 근생은 회귀에선 강하지만 여기선 낮은데, 연식·규모와 겹쳐 나무가 다른 요인으로 대신 나눴기 때문으로 보인다.")
    ci(doc, f"OOB R²({A['rf']['oob_r2']})가 회귀 R²({A['ols']['r2']}, 학습 자료 기준)보다 낮은 것은 정상이다. 건물 287동은 머신러닝에 작은 표본이라 결과는 **보조 확인**으로 쓴다.")

    # 8. 분해
    h1(doc, "8. 동별 격차 분해: 그 동은 왜 평균과 다른가")
    sq(doc, "분석 과정")
    ci(doc, "동 공실률과 전체 평균의 차이를 회귀계수로 요인별로 나눴다. 동의 요인 평균이 전체 평균보다 얼마나 다른지에 계수를 곱하면, 그 요인 때문에 생긴 공실률 차이가 된다.")
    formula(doc, "요인 k의 기여(%p) = β̂ₖ × (동의 가중평균 x̄ₖ − 전체 가중평균 x̄ₖ)")
    formula(doc, "격차 = 동 공실률 − 전체 공실률 = Σ 요인 기여(설명된 부분) + 모형 밖 부분")
    D = R["D"]
    g = D[D["dong"] == dmax["dong"]]
    ex_rows, tot = [], 0.0
    for i, (c, name, _) in enumerate(AN.FACTORS, start=1):
        dm = float(np.average(g[c], weights=g["shop_area"]))
        diff = dm - R["overall"][c]
        pp = R["beta"][i] * diff
        tot += pp
        ex_rows.append((abs(pp), [name, f"{R['beta'][i]:+.4f}", f"{dm:,.3f}", f"{R['overall'][c]:,.3f}", f"{diff:+,.3f}", f"{pp:+.1f}"]))
    ex_rows.sort(key=lambda r: -r[0])
    dec = A["decomp"][dmax["dong"]]
    assert abs(tot - dec["explained"]) < 0.15, (tot, dec)
    caption(doc, f"표 8. 계산 예시 — {dmax['dong']}(공실률 {dmax['rate']}%, 전체보다 {dec['gap']:+.1f}%p)")
    table(doc, ["요인", "계수 β̂(원단위)", f"{dmax['dong']} 평균", "전체 평균", "차이", "기여(%p)"],
          [r for _, r in ex_rows] + [["합계(설명된 부분)", "", "", "", "", f"{tot:+.1f}"]],
          [4.6, 2.6, 2.4, 2.4, 2.4, 2.4], size=8, align_right=(1, 2, 3, 4, 5), bold_rows=(len(ex_rows),))
    note(doc, f"계수는 원래 단위(1년, 0/1, %, ln㎡, m, ln점포, 만 명, %)의 값. 평균은 상가면적 가중. "
              f"격차 {dec['gap']:+.1f}%p = 설명된 부분 {dec['explained']:+.1f}%p + 모형 밖 {dec['other']:+.1f}%p.")
    ci(doc, f"{dmax['dong']}은 상가 면적의 대부분이 5년 이하 신축이고(5년 이하 더미 기여 가장 큼), 제2종 근생 비중도 높아 공실이 평균보다 높다. "
            f"나머지 {dec['other']:+.1f}%p는 연식·용도로 설명되지 않는 이 동만의 사정(상권 형성 지연 등)이다.")
    caption(doc, "표 9. 동별 격차 분해 요약(주요 기여 요인 상위 2개)")
    rws = []
    for d in sorted(A["by_dong"], key=lambda d: -A["decomp"][d["dong"]]["gap"]):
        v = A["decomp"][d["dong"]]
        top = " · ".join(f"{p['name']} {p['pp']:+.1f}" for p in v["parts"][:2])
        rws.append([d["dong"], f"{d['rate']:.1f}", f"{v['gap']:+.1f}", f"{v['explained']:+.1f}", f"{v['other']:+.1f}", top])
    table(doc, ["동", "공실률(%)", "격차(%p)", "설명(%p)", "모형 밖(%p)", "주요 기여 요인(%p)"], rws,
          [1.9, 1.8, 1.8, 1.8, 2.0, 7.7], size=7.8, align_right=(1, 2, 3, 4))
    note(doc, "해밀동(2동)·새롬동(11동)·다정동(10동)처럼 건물이 적은 동은 분해 값이 불안정하다. 해밀동은 신축·제2종인데도 공실이 3%로, 모형 밖 부분이 크다.")

    # 9. 핫스팟
    h1(doc, "9. 공실의 공간 군집(핫스팟)")
    sq(doc, "분석 과정")
    ci(doc, "개별 건물이 아니라 '공실이 높은 건물이 모여 있는 곳'을 찾기 위해 Getis-Ord Gi* 통계를 썼다.")
    da(doc, "건물 좌표를 평면 좌표(m)로 바꾸고, 반경 **500m** 안의 건물을 이웃(가중치 1, 자기 자신 포함)으로 두었다.")
    da(doc, "이웃 공실률 합이 전체 평균으로 기대되는 값보다 얼마나 큰지를 표준화한 값이 Gi*(z점수)다.")
    formula(doc, "Gi* = [Σⱼ wᵢⱼ xⱼ − x̄ Σⱼ wᵢⱼ] ÷ { s · √[(n Σⱼ wᵢⱼ² − (Σⱼ wᵢⱼ)²) / (n − 1)] }")
    da(doc, "Gi* > 1.96 → 핫스팟(공실 높은 건물 밀집, 95% 수준), Gi* < −1.96 → 콜드스팟.")
    hot_by = {}
    for p in A["points"]:
        if p["gi"] > 1.96:
            hot_by[p["dong"]] = hot_by.get(p["dong"], 0) + 1
    cold_by = {}
    for p in A["points"]:
        if p["gi"] < -1.96:
            cold_by[p["dong"]] = cold_by.get(p["dong"], 0) + 1
    sq(doc, "결과")
    ci(doc, f"핫스팟 **{A['hot']['hot']}동**: " + ", ".join(f"{k} {v}" for k, v in sorted(hot_by.items(), key=lambda kv: -kv[1])) + ".")
    ci(doc, f"콜드스팟 **{A['hot']['cold']}동**: " + ", ".join(f"{k} {v}" for k, v in sorted(cold_by.items(), key=lambda kv: -kv[1])) + ".")
    caption(doc, "그림 5. 건물별 공실률과 핫스팟")
    picture(doc, F["hot"], 15.0)

    # 10. 멀티레벨
    if ML:
        ml_chapter(doc, ML, F)

    # 11. 시사점
    h1(doc, ("11" if ML else "10") + ". 종합과 시사점")
    box(doc, [
        "**공급 시기 관리**: 공실은 준공 초기에 몰린다. 생활권별 상가 준공 시기를 나눠 공급하고, 준공 초기 임대료 감면·용도 유연화로 흡수 기간을 줄인다.",
        "**제2종 근생·상층부 용도 전환**: 제2종 근린생활시설 건물의 공실이 크다. 공유업무·돌봄·교육·생활SOC 등으로 전환을 허용하도록 지구단위계획 지침 완화를 검토한다.",
        "**부속 상가 총량 조정**: 업무·주거가 주된 건물의 부속 상가가 더 빈다. 이런 건물의 상가 의무 비율을 낮추고 1층 가로변 위주로 계획한다.",
        "**구역 단위 대응**: 핫스팟 구역을 집중 지원·임시 활용(팝업·공공 임대) 대상으로 삼는다. 3D 상황판에서 해당 블록을 확인할 수 있다.",
        "**입지 투자보다 공급 구조**: 정류장·중심지 접근성 개선만으로는 공실이 크게 줄지 않는다는 결과다. 교통 투자와 상가 공급 관리를 함께 봐야 한다.",
    ] + ([
        "**대응 단위는 행정동이 아니라 상권 블록**: 멀티레벨 분석에서 행정동 고유 차이는 거의 없고 공실 차이는 소권역(블록)에서 갈린다. "
        "지원·용도 전환·임시 활용을 블록 단위로 설계한다.",
        "**배후 세대 대비 상가 총량 관리**: 반경 500m 세대당 상가면적이 2배면 공실률이 높아진다(287동 모형에서 유의). "
        "신규 상업용지 공급과 상가 비율 지침을 배후 세대 수에 연동한다.",
    ] if ML else []), title="시사점")
    sq(doc, "한계")
    ci(doc, f"조사 시점이 섞여 있다({years[0]}년 {n23}동, {years[-1]}년 {A['n_buildings'] - n23}동). 시점 효과는 통제하지 않았다.")
    ci(doc, "건물 단위 자료라 층별(1층·상층) 차이와 업종, 임대료, 관리 상태는 보지 못했다.")
    ci(doc, "회귀·머신러닝 결과는 상관관계이지 인과가 아니다. 정류장·중심지 거리는 직선거리다.")
    ci(doc, "건물이 적은 동(해밀 2동, 다정 10동 등)의 값과 분해 결과는 불안정하다. 5생활권은 조사 건물이 없다.")
    para(doc, "")
    note(doc, "재현: sejong_vacant 폴더에서 prepare.py → analyze.py → tools/extra_features.py → multilevel.py → tools/make_report.py. 모든 수치는 같은 원자료에서 다시 계산되며, "
              "보고서 생성 시 회귀계수·분해 합계를 analysis.json 과 대조해 검산한다.")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    try:
        doc.save(OUT)
        return OUT
    except PermissionError:          # Word 에서 열려 있으면 덮어쓸 수 없다 → 새 이름으로
        alt = OUT.replace(".docx", "_v2.docx")
        doc.save(alt)
        return alt


def main():
    global ML
    mp = os.path.join(HERE, "data", "multilevel.json")
    ML = json.load(open(mp, encoding="utf-8")) if os.path.exists(mp) else None
    A = json.load(open(os.path.join(HERE, "data", "analysis.json"), encoding="utf-8"))
    X = excel_profile()
    R = recompute(A)
    F = figures(A, R)
    if ML:
        F["ml"] = ml_figure(ML)
    out = build(A, R, X, F)
    print("저장:", out)
    print("VIF", {k: round(v, 1) for k, v in R["vif"].items()})


if __name__ == "__main__":
    main()
