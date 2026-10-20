# -*- coding: utf-8 -*-
"""멀티레벨 분석 기술 부록(연구진 검토용) → docs/상가공실_멀티레벨_기술부록.docx

  python tools/make_ml_appendix.py      (multilevel.py 를 먼저 돌려 data/multilevel.json 이 있어야 한다)

상황판 '모든 수준을 함께 넣은 모형(M4)' 그림의 변수 정의·코딩 기준(더미·업종 다양도)·표본·자료 출처·추정식·
단계별 결과·민감도를 표와 그림으로 풀어 쓴다. 모형은 multilevel.py 의 prep()/fit_lmm() 을 그대로 다시 적합해
화면 숫자와 같다(SE·z·LR 검정·VIF 등 화면에 없는 값만 여기서 더 계산).
"""
from __future__ import annotations

import collections
import json
import math
import os
import sys

import numpy as np
import pandas as pd
from scipy import stats
from statsmodels.stats.outliers_influence import variance_inflation_factor

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "tools"))
import make_report as MR  # noqa: E402  (문서 서식·그림 글꼴 공유)
import multilevel as ML  # noqa: E402
import extra_features as XF  # noqa: E402
from make_report import plt, Document, Cm, WD_ALIGN_PARAGRAPH, para, table, caption, note, h1, sq, ci, da, formula, box, picture, add_runs  # noqa: E402,E501

OUT = os.path.join(HERE, "docs", "상가공실_멀티레벨_기술부록.docx")
FIG = os.path.join(HERE, "docs", "figures")
ZRAW = os.path.join(os.path.dirname(HERE), "sejong_livingzone", "data", "raw")
pstar = lambda p: "∗∗∗" if p < 0.001 else "∗∗" if p < 0.01 else "∗" if p < 0.05 else ""   # '**' 는 add_runs 의 굵게 표시와 겹친다
LEVEL_KO = {"건물": "건물(L1)", "반경": "반경(L1 주변)", "행정동": "행정동(L3)"}

# 변수 정의서: key → (원 변수, 정의·산식, 원자료·기준 시점, 코딩 기준 / 기준범주, 결측 처리)
DEF = {
    "new5": ("age", "준공 후 경과연수 ≤ 5년이면 1", "엑셀 '사용승인일자' → 2025-12-31까지 경과연수",
             "1 = 사용승인 2020-12-31 이후 / 0 = 그 이전(기준범주)", "없음"),
    "type2": ("use", "엑셀 주용도 문자열에 '제2종'이 들어 있으면 1", "엑셀 '주용도'",
              "1 = 제2종 근린생활시설 / 0 = 그 밖의 모든 주용도(제1종 근생 115동이 다수)", "없음"),
    "office": ("use", "주용도에 '업무'가 들어 있으면 1", "엑셀 '주용도'", "1 = 업무시설(오피스텔 포함) / 0 = 그 밖", "없음"),
    "mixed_res": ("use", "주용도에 '공동주택' 또는 '도시형생활주택'이 들어 있으면 1", "엑셀 '주용도'",
                  "1 = 주상복합(공동주택 하부 상가) / 0 = 그 밖. 주용도 '아파트'(1동)는 0으로 코딩됨", "없음"),
    "shop_share10": ("shop_share", "상가 면적 ÷ 연면적 × 100, 계수는 10%p 단위", "엑셀 '상가 면적'·'연면적'", "연속", "없음"),
    "log2_area": ("shop_area", "log₂(상가 면적㎡) — 계수 = 상가 면적이 2배일 때", "엑셀 '상가 면적'", "연속(로그)", "없음"),
    "d_stop100": ("d_stop", "가장 가까운 버스정류장까지 직선거리(m) ÷ 100", "TAGO 정류장 좌표(2026)", "연속", "없음(3km 상한)"),
    "d_center1k": ("d_center", "나성동 중심(127.2600E, 36.4870N)까지 직선거리(m) ÷ 1000", "브이월드 좌표", "연속", "없음"),
    "floors5": ("floors", "필지 내 건물 중 최고 지상층수 ÷ 5", "건축물대장 표제부(필지 PNU)", "연속", "대장 미매칭 시 브이월드 층수, 그래도 없으면 중앙값"),
    "log2_park": ("park_per_100", "log₂(상가 100㎡당 주차대수 + 0.5)", "표제부 옥내·옥외 자주식+기계식 합 ÷ 상가 면적 × 100", "연속(로그)", "중앙값 대체"),
    "food10": ("sh_food", "입점 점포 중 대분류 '음식' 비율(%) ÷ 10", "소상공인 상가(상권)정보 2026.06", "구성 비율", "점포 0곳 건물은 표본 제외"),
    "edu10": ("sh_edu", "대분류 '교육' 비율(%) ÷ 10", "소상공인 상가정보", "구성 비율", "상동"),
    "med10": ("sh_med", "대분류 '보건의료' 비율(%) ÷ 10", "소상공인 상가정보", "구성 비율", "상동"),
    "div01": ("diversity", "대분류 10종 정규화 섀넌 엔트로피(0~1) × 10 — 계수 = 0.1 증가", "소상공인 상가정보", "0 = 한 업종, 1 = 10종 균등", "상동"),
    "fran10": ("franchise", "'지점명'이 적힌 점포 비율(%) ÷ 10 — 체인 점포 대리", "소상공인 상가정보", "구성 비율", "상동"),
    "up3_10": ("sh_3up", "층 정보가 있는 점포 중 3층 이상 비율(%) ÷ 10", "소상공인 상가정보 '층'", "층 정보 3곳 이상일 때만 계산", "중앙값 대체"),
    "log2_sph": ("shop_per_hh", "log₂(반경 500m 조사 상가면적 합 ÷ 반경 500m 공동주택 세대)", "엑셀 상가면적 + 건축물대장 세대수(필지 좌표화)",
                 "연속(로그) — 세대당 상가 공급", "중앙값 대체"),
    "ppm100": ("apt_ppm", "반경 1km 아파트 ㎡당 실거래가(필지별 중앙값의 세대 가중평균, 만원) ÷ 100", "국토부 실거래가 2025.09~2026.08(해제 거래 제외)",
               "연속, 계수 = 100만원/㎡", "없음"),
    "chg1": ("u_chg", "행정동 주민등록인구 1년 증감률(%)", "행정안전부 주민등록 2025.08→2026.08", "행정동 값(같은 동 건물 공통)", "분동 전 묶음은 합산값"),
    "old1": ("u_old", "행정동 60세 이상 인구 비중(%)", "행정안전부 연령별 주민등록 2026.08", "행정동 값 — 화면 '고령'은 60세 이상", "없음"),
    "emp01": ("u_emp", "행정동 종사자 수 ÷ 주민등록인구, 계수는 0.1 단위", "SGIS 사업체 종사자 2024 / 주민등록 2026.08", "행정동 값", "분동 전 묶음은 합산값"),
}


def save(fig, name):
    os.makedirs(FIG, exist_ok=True)
    p = os.path.join(FIG, name)
    fig.savefig(p, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return p


def fit_all(D):
    out = {}
    for step in ML.STEPS:
        r = ML.fit_lmm(D, step)
        rml = ML.fit_lmm(D, step, reml=False)
        out[step] = (r, rml)
    return out


def store_examples(D):
    """업종 다양도 계산 예 — 낮음·중간·높음 건물 각 1곳의 대분류별 점포 수."""
    S = json.load(open(os.path.join(ZRAW, "sbiz.json"), encoding="utf-8"))["items"]
    idx = collections.defaultdict(list)
    for s in S:
        try:
            idx[XF.road_key(s["rdnm"], s["bldMnno"], s["bldSlno"])].append(s)
        except (ValueError, TypeError):
            pass
    tot = collections.Counter()
    per = {}
    for _, r in D.iterrows():
        st = idx.get(str(r["addr"]).strip(), [])
        c = collections.Counter(s["indsLclsNm"] for s in st)
        tot.update(c)
        per[r["no"]] = c
    q = D.sort_values("diversity")
    picks = [q.iloc[int(len(q) * f)] for f in (0.1, 0.5, 0.9)]
    return tot, [(p, per[p["no"]]) for p in picks]


def main():
    A = json.load(open(os.path.join(HERE, "data", "multilevel.json"), encoding="utf-8"))
    B = ML.prep()
    D = B[B["stores"] > 0].reset_index(drop=True)
    X21 = ML.cols_upto("M4")
    keys = [c for c, *_ in X21]
    name = {c: n for c, n, *_ in X21}
    unit = {c: u for c, n, u, _ in X21}
    lev = {c: l for c, n, u, l in X21}
    fits = fit_all(D)
    r4 = fits["M4"][0]
    m4j = {c["key"]: c for c in A["steps"][-1]["coefs"]}
    full4 = {c["key"]: c for c in A["full"][-1]["coefs"]}
    gl = {c["key"]: c for c in A["glmm"]["coefs"]}

    # ── 그림 ──────────────────────────────────────────────────────────
    F = {}
    # 1. 종속변수 분포 + 행정동별
    fig, ax = plt.subplots(1, 2, figsize=(10, 3.4), gridspec_kw={"width_ratios": [1, 1.5]})
    ax[0].hist(D["y"], bins=20, color="#1D5C7C", edgecolor="white")
    ax[0].axvline(D["y"].mean(), color="#B03A3A", ls="--", lw=1)
    ax[0].text(D["y"].mean() + 1, ax[0].get_ylim()[1] * .9, f"평균 {D['y'].mean():.1f}%", color="#B03A3A", fontsize=8)
    ax[0].set_xlabel("건물 공실률(면적, %)", fontsize=8)
    ax[0].set_ylabel("건물 수", fontsize=8)
    ax[0].set_title("(a) 건물 공실률 분포", fontsize=9)
    order = D.groupby("L3")["y"].median().sort_values().index.tolist()
    ax[1].boxplot([D.loc[D["L3"] == g, "y"] for g in order], vert=False, widths=.6,
                  medianprops={"color": "#B03A3A"}, flierprops={"markersize": 3})
    ax[1].set_yticks(range(1, len(order) + 1), [f"{g} ({(D['L3'] == g).sum()})" for g in order], fontsize=7.5)
    ax[1].set_xlabel("건물 공실률(면적, %)", fontsize=8)
    ax[1].set_title("(b) 행정동(L3)별 — 괄호는 건물 수", fontsize=9)
    for a in ax:
        a.tick_params(labelsize=7.5)
    F["y"] = save(fig, "ml_y.png")

    # 2. 연속 설명변수 분포(원 단위)
    cont = [("shop_share", "상가 비중(%)"), ("shop_area", "상가 면적(㎡, 로그축)"), ("d_stop", "정류장 거리(m)"), ("d_center", "나성동 중심 거리(m)"),
            ("floors", "지상 층수"), ("park_per_100", "상가 100㎡당 주차(로그축)"), ("sh_food", "음식점 비중(%)"), ("sh_edu", "학원·교육 비중(%)"),
            ("sh_med", "병의원 비중(%)"), ("diversity", "업종 다양도(0~1)"), ("franchise", "프랜차이즈 비중(%)"), ("sh_3up", "3층 이상 점포 비중(%)"),
            ("shop_per_hh", "세대당 상가면적(㎡, 로그축)"), ("apt_ppm", "아파트 ㎡당 가격(만원)"), ("age", "경과연수(년)")]
    fig, axs = plt.subplots(3, 5, figsize=(12, 6.6))
    for a, (c, t) in zip(axs.flat, cont):
        v = D[c].dropna()
        if "로그" in t:
            a.hist(np.log10(v.clip(lower=.1)), bins=18, color="#5E8FA8", edgecolor="white")
            a.set_xticks(a.get_xticks(), [f"{10 ** x:,.0f}" for x in a.get_xticks()])
        else:
            a.hist(v, bins=18, color="#5E8FA8", edgecolor="white")
        a.set_title(f"{t}\nn={len(v)} 평균 {v.mean():,.1f}", fontsize=7.5)
        a.tick_params(labelsize=6.5)
    fig.tight_layout()
    F["x"] = save(fig, "ml_x_hist.png")

    # 3. 상관 행렬(모형 단위 21개)
    C = D[keys].corr()
    fig, ax = plt.subplots(figsize=(9, 7.6))
    im = ax.imshow(C.values, cmap="RdBu_r", vmin=-1, vmax=1)
    ax.set_xticks(range(len(keys)), [name[k] for k in keys], rotation=75, fontsize=6.5, ha="right")
    ax.set_yticks(range(len(keys)), [name[k] for k in keys], fontsize=6.5)
    for i in range(len(keys)):
        for j in range(len(keys)):
            if i != j and abs(C.values[i, j]) >= .3:
                ax.text(j, i, f"{C.values[i, j]:.2f}", ha="center", va="center", fontsize=5,
                        color="white" if abs(C.values[i, j]) > .6 else "black")
    fig.colorbar(im, shrink=.7)
    F["corr"] = save(fig, "ml_corr.png")

    # 4. 업종 대분류 구성 + 다양도 분포
    tot, ex = store_examples(D)
    cats = XF.LCLS
    fig, ax = plt.subplots(1, 2, figsize=(10, 3.2))
    vals = [tot.get(c, 0) for c in cats]
    ax[0].barh(range(len(cats))[::-1], vals, color="#1D5C7C")
    ax[0].set_yticks(range(len(cats))[::-1], cats, fontsize=7.5)
    for i, v in zip(range(len(cats))[::-1], vals):
        ax[0].text(v, i, f" {v:,} ({v / sum(vals) * 100:.0f}%)", va="center", fontsize=7)
    ax[0].set_title(f"(a) 표본 건물 입점 점포의 대분류 구성(n={sum(vals):,})", fontsize=9)
    ax[0].set_xlim(0, max(vals) * 1.35)
    ax[1].hist(D["diversity"], bins=20, color="#5E8FA8", edgecolor="white")
    for (p, _), lab in zip(ex, ("예시 A", "예시 B", "예시 C")):
        ax[1].axvline(p["diversity"], color="#B03A3A", lw=1, ls=":")
        ax[1].text(p["diversity"], ax[1].get_ylim()[1] * .92, lab, fontsize=7, color="#B03A3A", ha="center")
    ax[1].set_xlabel("업종 다양도 H*(0~1)", fontsize=8)
    ax[1].set_title("(b) 건물별 업종 다양도 분포", fontsize=9)
    for a in ax:
        a.tick_params(labelsize=7.5)
    F["div"] = save(fig, "ml_diversity.png")

    # 5. 단계별 분산 분해
    fig, ax = plt.subplots(figsize=(8, 3.2))
    labs = [f"{s['step']}\n{s['name']}" for s in A["steps"]]
    l1 = [s["var"]["L1"] for s in A["steps"]]
    l2 = [s["var"]["L2"] for s in A["steps"]]
    l3 = [s["var"]["L3"] for s in A["steps"]]
    ax.bar(labs, l1, color="#B8C4C8", label="건물 내(잔차) σ²ₑ")
    ax.bar(labs, l2, bottom=l1, color="#5E8FA8", label="소권역 σ²ᵥ")
    ax.bar(labs, l3, bottom=np.add(l1, l2), color="#1D5C7C", label="행정동 σ²ᵤ")
    for i, s in enumerate(A["steps"]):
        ax.text(i, l1[i] + l2[i] + l3[i] + 8, f"{l1[i] + l2[i] + l3[i]:.0f}", ha="center", fontsize=7.5)
    ax.set_ylabel("분산(%p²)", fontsize=8)
    ax.legend(fontsize=7.5, loc="upper right")
    ax.tick_params(labelsize=7)
    F["var"] = save(fig, "ml_var.png")

    # 6. M4 계수 + 표준화 효과
    sd_mod = D[keys].std()
    std_eff = {k: float(r4.fe_params[k] * sd_mod[k]) for k in keys}
    order_k = sorted(keys, key=lambda k: std_eff[k])
    fig, ax = plt.subplots(figsize=(7.5, 6.4))
    for i, k in enumerate(order_k):
        b, se = r4.fe_params[k] * sd_mod[k], r4.bse_fe[k] * sd_mod[k]
        col = "#B03A3A" if (r4.pvalues[k] < .05 and b > 0) else "#2E7D4F" if r4.pvalues[k] < .05 else "#8A979C"
        ax.plot([b - 1.96 * se, b + 1.96 * se], [i, i], color=col, lw=1.6)
        ax.plot(b, i, "o", color=col, ms=4.5)
    ax.axvline(0, color="#555", lw=.8)
    ax.set_yticks(range(len(order_k)), [f"{name[k]}" for k in order_k], fontsize=7.5)
    ax.set_xlabel("1 표준편차 변화 시 공실률 변화(%p, 95% 신뢰구간)", fontsize=8)
    ax.tick_params(labelsize=7.5)
    F["std"] = save(fig, "ml_std_effect.png")

    # 7. 계수 안정성 M1~M4
    show = [k for k, *_ in ML.BLOCKS["M1"]] + ["div01", "up3_10", "food10", "log2_sph"]
    fig, ax = plt.subplots(figsize=(11, 4.4))
    mk = {"M1": "o", "M2": "s", "M3": "^", "M4": "D"}
    for j, step in enumerate(("M1", "M2", "M3", "M4")):
        cs = {c["key"]: c for c in next(s for s in A["steps"] if s["step"] == step)["coefs"]}
        for i, k in enumerate(show):
            if k in cs:
                ax.errorbar(i + (j - 1.5) * .17, cs[k]["b"], yerr=[[cs[k]["b"] - cs[k]["lo"]], [cs[k]["hi"] - cs[k]["b"]]],
                            fmt=mk[step], ms=4, lw=.9, capsize=0, color=["#8DB86B", "#E5B94A", "#5E8FA8", "#1D5C7C"][j],
                            label=step if i == 0 or (step == "M2" and k == "div01") or (step == "M3" and k == "log2_sph") else None)
    ax.axhline(0, color="#555", lw=.8)
    wrap = lambda t: "\n".join(t[i:i + 6] for i in range(0, len(t), 6))
    ax.set_xticks(range(len(show)), [f"{wrap(name[k])}\n/{unit[k]}" for k in show], fontsize=6.5, rotation=0)
    ax.set_ylabel("계수(%p, 95% CI)", fontsize=8)
    h, l = ax.get_legend_handles_labels()
    seen = {}
    for hh, ll in zip(h, l):
        seen.setdefault(ll, hh)
    ax.legend(seen.values(), seen.keys(), fontsize=7.5, ncol=4)
    ax.tick_params(labelsize=7)
    F["stab"] = save(fig, "ml_stability.png")

    # 8. 잔차 진단
    fv = r4.fittedvalues
    rs = r4.resid
    fig, ax = plt.subplots(1, 3, figsize=(11, 3.1))
    ax[0].scatter(fv, rs, s=8, color="#1D5C7C", alpha=.7)
    ax[0].axhline(0, color="#555", lw=.8)
    ax[0].set_xlabel("적합값(%)", fontsize=8)
    ax[0].set_ylabel("잔차(%p)", fontsize=8)
    ax[0].set_title("(a) 적합값 대 잔차", fontsize=9)
    stats.probplot(rs, dist="norm", plot=ax[1])
    ax[1].get_lines()[0].set_markersize(3)
    ax[1].set_title("(b) 잔차 정규 Q-Q", fontsize=9)
    ax[1].set_xlabel("이론 분위수", fontsize=8)
    ax[1].set_ylabel("표본 분위수", fontsize=8)
    re = sorted(A["dong_re"], key=lambda x: x["re"])
    ax[2].barh([f"{x['unit']} ({x['n']})" for x in re], [x["re"] for x in re],
               color=["#B03A3A" if x["re"] > 0 else "#2E7D4F" for x in re])
    ax[2].axvline(0, color="#555", lw=.8)
    ax[2].set_title("(c) 행정동 무작위절편 û (BLUP, %p)", fontsize=9)
    for a in ax:
        a.tick_params(labelsize=7)
    fig.tight_layout()
    F["diag"] = save(fig, "ml_diag.png")

    # ── 문서 ──────────────────────────────────────────────────────────
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.left_margin = sec.right_margin = Cm(2.0)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_runs(fp, "세종연구원 연구모임 · 행복도시 상가공실 멀티레벨 분석 기술 부록(검토용)", 7.5, color=MR.MUTED)

    para(doc, "세종연구원 · 연구모임 검토자료 — 기술 부록", 9, color=MR.MUTED, after=2)
    para(doc, "행복도시 상가공실 멀티레벨 분석: 변수·표본·추정 상세", 17, True, MR.ACCENT, after=2)
    para(doc, "상황판 '모든 수준을 함께 넣은 모형(M4)' 그림의 지표 정의, 더미·업종 다양도 기준, 자료 구성, 추정식과 수치 해설", 10.5,
         color=MR.MUTED, after=8)
    st4 = A["steps"][-1]
    box(doc, [
        f"**표본**: 2025년 12월 집합상가 전수조사 {A['n_full']}동 중 입점 점포가 확인된 **{A['n']}동**(L1) ⊂ 소권역 **{A['n_L2']}곳**(L2) ⊂ 행정동 **{A['n_L3']}곳**(L3). "
        f"점포가 확인되지 않은 {A['excluded']['n']}동(공실률 {A['excluded']['rate']}%)은 업종 구성 변수를 만들 수 없어 제외하고, 별도 민감도 분석(10장)으로 확인.",
        "**모형**: 선형 혼합모형(REML), 종속변수 = 건물 공실률(공실 면적 ÷ 상가 면적 × 100, %). 행정동·소권역 무작위절편. "
        "계수는 '설명변수가 표시 단위만큼 늘 때 공실률이 몇 %p 변하는가'.",
        f"**M4 적합**: 고정효과 21개, 주변 R²(고정효과만) {st4['r2m']}, 조건부 R²(무작위효과 포함) {st4['r2c']}, "
        f"잔차 Moran's I {A['moran']['I']}(p={A['moran']['p']}).",
        "**그림 읽기**: 굵은 이름 = p<0.05. '287동 유의' = 점포 정보가 없는 21동까지 넣은 모형에서도 p<0.05(업종 구성 변수는 그 모형에 없음).",
    ], title="한눈에")

    # 1. 표본
    h1(doc, "1. 자료와 표본 구성")
    sq(doc, "자료 출처")
    table(doc, ["자료", "제공", "기준 시점", "공간 단위", "건물과 잇는 방법", "만든 변수"], [
        ["상가 전수조사(엑셀) 2025.12 집합상가 시트", "세종연구원", "2025년 12월", "건물", "—", "공실률, 상가·연면적, 주용도, 사용승인일"],
        ["브이월드 지오코딩·건물 외곽선", "국토교통부", "2026", "건물·필지", "주소 → 좌표 → 외곽선 안 점 → PNU", "좌표, 필지 번호"],
        ["건축물대장 표제부", "건축HUB", "2026.10 조회", "필지 내 동", "PNU", "층수, 주차대수"],
        ["소상공인 상가(상권)정보", "소상공인시장진흥공단", "2026.06", "점포", "도로명+건물본번·부번 일치", "업종 구성·다양도·프랜차이즈·층"],
        ["버스정류장(TAGO)", "국토교통부", "2026", "점", "최근접 직선거리", "정류장 거리"],
        ["건축물대장 공동주택 세대수", "건축HUB", "2026", "필지", "지번 → 좌표, 반경 500m", "세대당 상가면적"],
        ["아파트 매매 실거래가", "국토교통부", "2025.09~2026.08", "필지", "지번 → 좌표, 반경 1km", "㎡당 가격"],
        ["주민등록 인구·연령", "행정안전부", "2025.08, 2026.08", "행정동", "행정동 코드", "인구 증감률, 60세 이상 비중"],
        ["사업체 종사자", "통계청 SGIS", "2024", "행정동", "행정동 이름", "직주비"],
    ], [3.4, 2.1, 2.0, 1.6, 3.6, 4.3], size=7.5)
    note(doc, "점포·실거래·인구는 상황판 수집 시점(2026년) 값으로, 공실 조사(2025.12)보다 반년 남짓 늦다. 원인-결과 시차가 없는 단면 상관으로 읽어야 한다.")

    sq(doc, "표본 흐름과 수준 구조")
    ci(doc, f"엑셀 2025.12 집합상가 시트 {A['n_full']}동 → 좌표화 287동(실패 0) → 입점 점포 1곳 이상 매칭 **{A['n']}동**(분석 표본).")
    ci(doc, "L3 행정동: 도담동·어진동은 분동 이력·자료 체계 때문에 '도담·어진동'으로 합쳐 13곳. "
            "L2 소권역: 엑셀 '동 세부' 열(예: 고운동 북측·남측)로 행정동 안을 나눈 상권 블록. 한 소권역은 한 행정동에만 속한다(완전 내포).")
    rows = []
    for g, sub in D.groupby("L3"):
        rows.append([g, f"{len(sub)}", f"{sub['L2'].nunique()}", f"{sub['shop_area'].sum():,.0f}",
                     f"{sub['vac_area'].sum() / sub['shop_area'].sum() * 100:.1f}", f"{sub['y'].mean():.1f}",
                     f"{(B['L3'] == g).sum() - len(sub)}"])
    rows.sort(key=lambda r: -int(r[1]))
    rows.append(["합계", f"{len(D)}", f"{D['L2'].nunique()}", f"{D['shop_area'].sum():,.0f}",
                 f"{D['vac_area'].sum() / D['shop_area'].sum() * 100:.1f}", f"{D['y'].mean():.1f}", f"{len(B) - len(D)}"])
    caption(doc, "표 1. 행정동(L3)별 표본 구성")
    table(doc, ["행정동(L3)", "건물(L1)", "소권역(L2)", "상가 면적(㎡)", "공실률 면적합(%)", "건물 공실률 평균(%)", "제외(점포 0)"], rows,
          [3.0, 1.8, 1.9, 2.8, 2.6, 2.8, 2.1], size=8, align_right=(1, 2, 3, 4, 5, 6), bold_rows=(len(rows) - 1,))
    note(doc, "'공실률 면적합' = Σ공실면적 ÷ Σ상가면적(큰 건물 비중 큼). '건물 공실률 평균' = 건물별 공실률의 단순평균 — 혼합모형의 종속변수 눈금은 이쪽이다.")
    l2 = D.groupby("L2").agg(n=("y", "size"), y=("y", "mean")).reset_index()
    l2[["L3", "sub"]] = l2["L2"].str.split("|", expand=True)
    caption(doc, "표 2. 소권역(L2) 목록 — 건물 수와 건물 공실률 평균")
    half = (len(l2) + 1) // 2
    rr = []
    for i in range(half):
        a = l2.iloc[i]
        b_ = l2.iloc[i + half] if i + half < len(l2) else None
        rr.append([a["sub"], f"{a['n']}", f"{a['y']:.1f}"] + ([b_["sub"], f"{b_['n']}", f"{b_['y']:.1f}"] if b_ is not None else ["", "", ""]))
    table(doc, ["소권역", "건물", "평균(%)", "소권역", "건물", "평균(%)"], rr, [4.2, 1.4, 1.7, 4.2, 1.4, 1.7], size=7.5,
          align_right=(1, 2, 4, 5))
    note(doc, f"소권역당 건물 수 중앙값 {int(l2['n'].median())}동(최소 {l2['n'].min()}, 최대 {l2['n'].max()}). 건물이 1~2동인 소권역은 무작위효과가 0 쪽으로 강하게 수축된다.")

    # 2. 종속변수
    h1(doc, "2. 종속변수: 건물 공실률")
    formula(doc, "y_i = (공실 면적_i ÷ 상가 면적_i) × 100   [%]")
    ci(doc, "단위는 건물. 혼합모형은 건물마다 같은 무게(가중치 없음)로 적합한다. 상가 면적 가중 회귀(본 보고서의 OLS)와 달리 "
            "작은 건물도 큰 건물과 같은 한 표본이며, 큰 건물에 무게를 두는 확인은 점포(호) 단위 이항 혼합모형(9장)이 맡는다.")
    q = D["y"].describe(percentiles=[.25, .5, .75])
    table(doc, ["n", "평균", "표준편차", "최소", "25%", "중앙값", "75%", "최대", "공실 0% 건물", "공실 50% 이상"],
          [[f"{int(q['count'])}", f"{q['mean']:.1f}", f"{q['std']:.1f}", f"{q['min']:.1f}", f"{q['25%']:.1f}", f"{q['50%']:.1f}",
            f"{q['75%']:.1f}", f"{q['max']:.1f}", f"{(D['y'] == 0).sum()}", f"{(D['y'] >= 50).sum()}"]],
          [1.2, 1.4, 1.6, 1.3, 1.3, 1.5, 1.3, 1.3, 2.4, 2.6], size=8, align_right=tuple(range(10)))
    caption(doc, "그림 1. 종속변수 분포와 행정동별 상자그림")
    picture(doc, F["y"], 16.5)
    note(doc, "0%에 몰리고 오른쪽으로 긴 분포다. 선형모형의 계수(평균 변화)는 그대로 해석 가능하지만, 잔차 정규성이 약해 p값은 근사치로 본다(8장 잔차 진단, 9장 이항모형으로 교차 확인).")

    # 3. 변수 정의
    h1(doc, "3. 설명변수 정의서(21개)")
    ci(doc, "모든 연속변수는 그림의 '/단위'에 맞게 미리 나누거나 로그를 취해, 계수가 곧 '표시 단위 1만큼 변할 때 공실률 %p 변화'가 되도록 했다.")
    ci(doc, "**로그 변수(/2배)**: log₂ x 를 쓰면 x 가 2배가 될 때 log₂ x 가 정확히 1 늘어난다. 그래서 계수 = 'x 가 2배일 때 %p'.")
    caption(doc, "표 3. 변수 정의·원자료·코딩 기준")
    rows = []
    for k in keys:
        raw, df, src, code, miss = DEF[k]
        rows.append([f"{name[k]}\n({k})", LEVEL_KO[lev[k]], f"/{unit[k]}", df, src, code, miss])
    table(doc, ["변수(코드)", "수준", "계수 단위", "정의·산식", "원자료·시점", "코딩 기준 / 기준범주", "결측 처리"], rows,
          [2.6, 1.6, 1.3, 3.6, 3.0, 3.2, 1.9], size=6.8)

    sq(doc, "더미 변수의 기준")
    ci(doc, "주용도 더미 3개(제2종 근생·업무시설·주상복합)는 서로 겹치지 않는다. 셋 다 0인 건물이 **기준범주**이며, 계수는 '기준범주보다 몇 %p 높은가/낮은가'다.")
    use_rows = []
    for u, n in D["use"].value_counts().items():
        grp = "제2종 근생(type2=1)" if "제2종" in u else "업무시설(office=1)" if "업무" in u else \
            "주상복합(mixed_res=1)" if ("공동주택" in u or "도시형생활주택" in u) else "기준범주(셋 다 0)"
        use_rows.append([u, grp, f"{n}", f"{D.loc[D['use'] == u, 'y'].mean():.1f}"])
    caption(doc, "표 4. 엑셀 주용도 값과 더미 배정")
    table(doc, ["엑셀 주용도", "배정", "건물", "공실률 평균(%)"], use_rows, [5.0, 5.0, 1.8, 2.8], size=8, align_right=(2, 3))
    note(doc, "기준범주는 제1종 근린생활시설이 대부분이고 자동차관련시설(주차빌딩 상가)·판매·교육연구 등 소수 용도가 섞여 있다. 주용도 '아파트' 1동은 문자열 규칙상 주상복합 더미에서 빠졌다(영향은 미미).")
    drows = []
    for k in ("new5", "type2", "office", "mixed_res"):
        one, zero = D[D[k] == 1], D[D[k] == 0]
        t = stats.ttest_ind(one["y"], zero["y"], equal_var=False)
        drows.append([name[k], DEF[k][3], f"{len(one)} / {len(zero)}", f"{one['y'].mean():.1f}", f"{zero['y'].mean():.1f}",
                      f"{one['y'].mean() - zero['y'].mean():+.1f} (p={t.pvalue:.3f})", f"{m4j[k]['b']:+.1f} (p={m4j[k]['p']:.3f})"])
    caption(doc, "표 5. 더미별 단순 비교와 M4 계수(다른 변수 통제 후)")
    table(doc, ["더미", "1 / 0 기준", "n(1/0)", "평균(1)", "평균(0)", "단순 차이(Welch)", "M4 계수"], drows,
          [2.6, 4.6, 1.6, 1.4, 1.4, 2.6, 2.6], size=7.5, align_right=(2, 3, 4, 5, 6))
    mr = drows[3]
    note(doc, "단순 차이와 M4 계수가 다른 것은 다른 요인(연식·상가 비중·입지·업종 구성 등)을 통제했기 때문이다. "
              f"예: 주상복합은 단순 비교 차이가 {mr[5].split(' ')[0]}%p로 거의 없지만, 상가 비중(주상복합은 낮음)·입지 등을 통제하면 {mr[6].split(' ')[0]}%p가 된다. "
              f"준공 5년 이하는 {drows[0][2].split(' / ')[0]}동뿐이라 단순 차이·계수 모두 신뢰구간이 매우 넓다.")

    sq(doc, "업종 다양도(업종 대분류 기준)")
    ci(doc, "소상공인 상가(상권)정보의 **표준산업분류 기반 상권업종 '대분류' 10종**을 쓴다(중·소분류는 쓰지 않음 — 건물당 점포 수가 적어 세분류를 쓰면 대부분 1~2곳씩이 되어 지표가 불안정).")
    formula(doc, "H*_i = − Σ_k p_ik · ln p_ik  ÷  ln 10,      p_ik = 건물 i 의 대분류 k 점포 수 ÷ 건물 i 의 점포 수")
    ci(doc, "정규화 섀넌 엔트로피: 모든 점포가 한 업종이면 0, 10개 대분류에 똑같이 나뉘면 1. 화면 계수는 0.1 단위(div01 = H* × 10).")
    ci(doc, "점포 '개수'는 공실의 거울(빈 칸이 많으면 점포가 적다)이라 쓰지 않고, 구성 비율만 쓴다.")
    caption(doc, "그림 2. 대분류 구성과 업종 다양도 분포")
    picture(doc, F["div"], 16.5)
    ex_rows = []
    for (p, c), lab in zip(ex, ("A(하위 10%)", "B(중앙)", "C(상위 10%)")):
        n = sum(c.values())
        ex_rows.append([lab, f"{p['L3']}", f"{n}"] + [f"{c.get(k, 0)}" for k in cats] + [f"{p['diversity']:.3f}"])
    caption(doc, "표 6. 업종 다양도 계산 예 — 건물별 대분류 점포 수")
    table(doc, ["예시", "행정동", "점포"] + cats + ["H*"], ex_rows,
          [1.9, 1.7, 0.9] + [1.05] * len(cats) + [1.1], size=6.5, align_right=tuple(range(2, 4 + len(cats))))
    pB, cB = ex[1]
    nB = sum(cB.values())
    terms = " + ".join(f"{v}/{nB}·ln({v}/{nB})" for v in cB.values() if v)
    note(doc, f"예시 B: H* = −[{terms}] ÷ ln10 = {pB['diversity']:.3f}. "
              f"M4 계수 {m4j['div01']['b']:+.2f}%p/0.1 → 예시 A({ex[0][0]['diversity']:.2f})에서 C({ex[2][0]['diversity']:.2f})로 가면 "
              f"공실률이 약 {m4j['div01']['b'] * (ex[2][0]['diversity'] - ex[0][0]['diversity']) * 10:+.1f}%p 달라진다(다른 조건 같을 때).")

    sq(doc, "그 밖의 점포 구성 변수")
    da(doc, "음식점·학원교육·병의원 비중 = 각 대분류('음식', '교육', '보건의료') 점포 ÷ 건물 점포 × 100. 기준은 나머지 대분류 전체.")
    da(doc, "프랜차이즈 비중 = 상가정보 '지점명(brchNm)'이 비어 있지 않은 점포 비율. 공식 가맹 여부가 아니라 '○○점' 표기를 체인의 대리로 쓴 것.")
    da(doc, "3층 이상 점포 비중 = 층 정보가 숫자로 있는 점포 중 3층 이상 비율. 층 정보가 3곳 미만이면 계산하지 않고 중앙값으로 채웠다"
            f"({int(D['sh_3up'].isna().sum())}동).")

    sq(doc, "반경·행정동 변수")
    da(doc, "세대당 상가면적 = 반경 500m 안 조사 건물 상가면적 합(자기 건물 포함) ÷ 반경 500m 안 공동주택 세대 합. '배후 세대 대비 상가 공급 과잉' 지표. "
            f"세대가 0이라 정의되지 않는 {int(D['shop_per_hh'].isna().sum())}동은 중앙값으로 채움.")
    da(doc, "배후 아파트 ㎡당 가격 = 반경 1km 공동주택 필지별 전용 ㎡당 매매가 중앙값을 세대 수로 가중평균(만원). 배후 구매력의 대리.")
    da(doc, "행정동 변수 3개는 같은 행정동 건물이 모두 같은 값을 갖는 L3 수준 변수다. 행정동이 13곳뿐이라 이 계수의 유효 표본은 13이며 신뢰구간이 넓다.")

    # 4. 기술통계
    h1(doc, "4. 기술통계와 공선성")
    raw_rows = []
    raw_map = [("age", "경과연수(년)")] + [(DEF[k][0], name[k]) for k in keys if DEF[k][0] not in ("age", "use")]
    for c, n in raw_map:
        v = D[c]
        raw_rows.append([n, f"{int(v.notna().sum())}", f"{v.mean():,.2f}", f"{v.std():,.2f}", f"{v.min():,.2f}", f"{v.median():,.2f}", f"{v.max():,.2f}"])
    caption(doc, f"표 7. 설명변수 원 단위 기술통계(n={len(D)}, 대체 전)")
    table(doc, ["변수", "n", "평균", "표준편차", "최소", "중앙값", "최대"], raw_rows, [5.0, 1.2, 2.0, 2.0, 1.9, 2.0, 2.2], size=7.5,
          align_right=(1, 2, 3, 4, 5, 6))
    over = int((D["shop_share"] > 100).sum())
    note(doc, f"상가 비중이 100%를 넘는 건물 {over}동은 엑셀의 상가 면적(전용+공용 등)이 연면적보다 크게 적힌 경우로, 원자료 그대로 두었다(검토 대상). "
              "주차는 오른쪽 꼬리가 매우 길어(최대 120대/100㎡ — 주차빌딩 상가) 모형에서는 log₂(x+0.5)로 넣었다.")
    caption(doc, "그림 3. 연속 설명변수 분포(원 단위)")
    picture(doc, F["x"], 16.5)
    X = D[keys].astype(float)
    Xc = np.column_stack([np.ones(len(X)), X.values])
    vif = {k: variance_inflation_factor(Xc, i + 1) for i, k in enumerate(keys)}
    C = X.corr()
    pairs = sorted(((abs(C.loc[a, b]), a, b) for i, a in enumerate(keys) for b in keys[i + 1:]), reverse=True)[:8]
    caption(doc, "그림 4. 설명변수 상관행렬(모형 단위, |r| ≥ 0.3만 숫자 표시)")
    picture(doc, F["corr"], 14.5)
    caption(doc, "표 8. 상관이 큰 변수 쌍과 분산팽창계수(VIF)")
    vr = [[f"{name[a]} – {name[b]}", f"{C.loc[a, b]:+.2f}"] for _, a, b in pairs]
    vv = sorted(vif.items(), key=lambda x: -x[1])[:8]
    table(doc, ["상관 큰 쌍", "r", "VIF 큰 변수", "VIF"],
          [vr[i] + [name[vv[i][0]], f"{vv[i][1]:.2f}"] for i in range(8)], [6.4, 1.4, 5.0, 1.6], size=7.5, align_right=(1, 3))
    note(doc, f"VIF 최대 {max(vif.values()):.1f} — 통상 기준(5~10)보다 작아 심각한 공선성은 아니다. 행정동 변수끼리는 13개 값만 갖고 있어 상관이 큰 편이다.")

    # 5. 모형
    h1(doc, "5. 모형과 추정")
    formula(doc, "y_ijk = β₀ + Σ_p β_p x_p,ijk + u_k + v_jk + ε_ijk")
    formula(doc, "u_k ~ N(0, σ²_u),   v_jk ~ N(0, σ²_v),   ε_ijk ~ N(0, σ²_ε)   (서로 독립)")
    ci(doc, "i = 건물, j = 소권역, k = 행정동. u_k = 행정동 무작위절편, v_jk = 소권역 무작위절편(행정동 안에 내포), ε = 건물 고유 오차.")
    ci(doc, "추정: statsmodels MixedLM, 제한최대우도(REML)로 분산성분·고정효과. 단계 비교용 AIC·우도비 검정은 같은 고정효과를 최대우도(ML)로 다시 적합해 계산"
            "(REML 우도는 고정효과가 다른 모형끼리 비교할 수 없기 때문).")
    ci(doc, "고정효과 검정: Wald z = β̂ ÷ SE(β̂), 양측 정규 p값. 95% 신뢰구간 = β̂ ± 1.96·SE.")
    sq(doc, "적합도 지표")
    formula(doc, "ICC_L3 = σ²_u ÷ (σ²_u + σ²_v + σ²_ε),     ICC_L2 = σ²_v ÷ (σ²_u + σ²_v + σ²_ε)")
    formula(doc, "PCV_level(Mk) = 1 − σ²_level(Mk) ÷ σ²_level(M0)")
    formula(doc, "R²_m = σ²_f ÷ (σ²_f + σ²_u + σ²_v + σ²_ε),   R²_c = (σ²_f + σ²_u + σ²_v) ÷ (σ²_f + σ²_u + σ²_v + σ²_ε),   σ²_f = Var(Xβ̂)")
    ci(doc, "ICC = 전체 분산 중 그 수준 '사이'의 몫. PCV = 빈 모형 대비 그 수준 분산이 줄어든 비율(그 수준의 차이를 변수들이 얼마나 설명했나). "
            "R²는 Nakagawa–Schielzeth 방식(주변 = 고정효과만, 조건부 = 무작위효과 포함).")
    ci(doc, "주의: 빈 모형의 행정동 분산이 거의 0이면 PCV 분모가 0에 가까워 음수·극단값이 나온다(표 9의 L3 PCV). 이때는 PCV 대신 분산 자체를 본다.")

    # 6. 단계별
    h1(doc, "6. 단계별 결과(M0 → M4)")
    lr = []
    prev = None
    for s in ML.STEPS:
        rml = fits[s][1]
        k = len(ML.cols_upto(s))
        if prev is None:
            lr.append("—")
        else:
            dev = 2 * (rml.llf - prev[0])
            df_ = k - prev[1]
            lr.append(f"χ²({df_})={dev:.1f}, p={stats.chi2.sf(dev, df_):.3f}")
        prev = (rml.llf, k)
    srows = []
    for s, l in zip(A["steps"], lr):
        pc = lambda v: "—" if v is None else f"{v:.2f}"
        srows.append([f"{s['step']} {s['name']}", f"{s['k']}", f"{s['aic']:.1f}", l, f"{s['var']['L3']:.1f}", f"{s['var']['L2']:.1f}",
                      f"{s['var']['L1']:.1f}", f"{s['icc']['L3']:.3f}", f"{s['icc']['L2']:.3f}",
                      f"{pc(s['pcv']['L2'])} / {pc(s['pcv']['L1'])}", f"{s['r2m']:.3f}", f"{s['r2c']:.3f}"])
    caption(doc, "표 9. 단계별 분산성분·적합도")
    table(doc, ["단계", "변수", "AIC(ML)", "우도비(직전 대비)", "σ²_u L3", "σ²_v L2", "σ²_ε", "ICC L3", "ICC L2", "PCV L2/L1", "R²m", "R²c"], srows,
          [3.2, 0.9, 1.3, 2.7, 1.2, 1.2, 1.2, 1.1, 1.1, 1.5, 1.0, 1.0], size=6.8, align_right=tuple(range(1, 12)))
    m0 = A["steps"][0]
    note(doc, f"빈 모형(M0): 건물 간 공실률 분산의 {m0['icc']['L2'] * 100:.0f}%가 소권역 사이, {m0['icc']['L3'] * 100:.1f}%가 행정동 사이에 있다 — "
              "같은 동이어도 어느 블록이냐가 훨씬 중요하다. 단계마다 L2 분산이 줄어드는 것은 블록 간 차이 일부를 건물·반경 요인이 설명했다는 뜻이다.")
    caption(doc, "그림 5. 단계별 분산 분해(막대 위 = 총분산)")
    picture(doc, F["var"], 14.5)

    # 7. M4 계수
    h1(doc, "7. M4 계수 상세(상황판 그림의 수치)")
    ci(doc, "상황판 그림의 오른쪽 숫자 = 아래 표의 β̂, 가로선 = 95% 신뢰구간. 색: 빨강 = 유의하게 공실을 높임, 초록 = 유의하게 낮춤, 회색 = 유의하지 않음.")
    crow = []
    sig_rows = []
    for i, k in enumerate(keys):
        b, se, p = float(r4.fe_params[k]), float(r4.bse_fe[k]), float(r4.pvalues[k])
        if p < .05:
            sig_rows.append(i)
        fk = full4.get(k)
        g = gl[k]
        crow.append([f"{name[k]}", LEVEL_KO[lev[k]].split("(")[0], f"/{unit[k]}", f"{b:+.2f}", f"{se:.2f}", f"{b / se:+.2f}",
                     f"{p:.3f}{pstar(p)}", f"[{b - 1.96 * se:+.1f}, {b + 1.96 * se:+.1f}]", f"{sd_mod[k]:.2f}", f"{std_eff[k]:+.1f}",
                     (f"{fk['b']:+.1f}{pstar(fk['p'])}" if fk else "—"), f"{g['or']:.2f}{'†' if g['sig'] else ''}"])
    caption(doc, f"표 10. M4 고정효과(n={A['n']}, REML) — 절편 {r4.fe_params['Intercept']:.1f}")
    table(doc, ["변수", "수준", "단위", "β̂(%p)", "SE", "z", "p", "95% CI", "SD(단위)", "1SD 효과", "287동 M4", "이항 OR"], crow,
          [3.3, 1.1, 1.3, 1.3, 1.0, 1.1, 1.4, 2.0, 1.3, 1.3, 1.4, 1.2], size=6.8, align_right=tuple(range(3, 12)), bold_rows=tuple(sig_rows))
    note(doc, "굵은 줄 = p<0.05. ∗ p<0.05, ∗∗ p<0.01, ∗∗∗ p<0.001. 'SD(단위)' = 그 변수의 표본 표준편차를 계수 단위로 잰 값, '1SD 효과' = β̂ × SD — 변수끼리 크기를 견줄 때 쓴다. "
              "'287동 M4' = 점포 구성 변수를 뺀 287동 모형 계수(10장). '이항 OR' = 점포 단위 이항 혼합모형의 승산비, † = 95% 구간이 1을 포함하지 않음.")
    caption(doc, "그림 6. 1 표준편차 변화 시 효과(변수 간 크기 비교)")
    picture(doc, F["std"], 13.5)

    sq(doc, "숫자 읽는 법 — 예시")
    ex_ = lambda k: m4j[k]["b"]
    box(doc, [
        f"**제2종 근린생활시설 {ex_('type2'):+.1f}%p**: 연식·상가 비중·입지·업종 구성·행정동 여건이 같은 두 건물 중 주용도가 제2종 근생인 건물의 공실률이 기준범주(주로 제1종 근생)보다 평균 {ex_('type2'):.1f}%p 높다.",
        f"**나성동 중심 거리 {ex_('d_center1k'):+.1f}%p/1km**: 중심에서 1km 멀어질 때마다 공실률 {ex_('d_center1k'):+.1f}%p. 표본 거리 범위(약 {D['d_center'].min() / 1000:.1f}~{D['d_center'].max() / 1000:.1f}km)를 다 걸으면 약 {ex_('d_center1k') * (D['d_center'].max() - D['d_center'].min()) / 1000:+.0f}%p.",
        f"**상가 규모 {ex_('log2_area'):+.1f}%p/2배**: 상가 면적 1,000㎡ → 2,000㎡ 또는 5,000㎡ → 10,000㎡ 모두 같은 {ex_('log2_area'):+.1f}%p(로그 눈금이라 '배율'에 대한 효과).",
        f"**업종 다양도 {ex_('div01'):+.1f}%p/0.1**: H*가 0.4 → 0.6(0.2 증가)이면 {ex_('div01') * 2:+.1f}%p.",
        f"**3층 이상 점포 비중 {ex_('up3_10'):+.1f}%p/10%p**: 입점 점포 중 3층 이상이 20% → 50%이면 {ex_('up3_10') * 3:+.1f}%p. "
        "상층 점포가 '들어와 있다'는 것 자체가 상층까지 수요가 있다는 신호로 읽힌다(상층 공실이 많다는 뜻이 아님).",
        f"**행정동 인구 증감률 {ex_('chg1'):+.1f}%p/1%p**: 신뢰구간이 넓다 — 행정동 13곳뿐인 L3 변수의 한계.",
    ], title="계수 → 공실률 %p 환산")

    sq(doc, "계수 안정성")
    ci(doc, "변수 묶음을 더해 가도 주요 계수의 부호·크기가 크게 바뀌지 않는지 확인했다(그림 7).")
    caption(doc, "그림 7. 단계별 계수(M1→M4)")
    picture(doc, F["stab"], 16.0)
    trows = []
    for k in keys:
        r_ = [name[k]]
        for s in ("M1", "M2", "M3", "M4"):
            cs = {c["key"]: c for c in next(x for x in A["steps"] if x["step"] == s)["coefs"]}
            r_.append(f"{cs[k]['b']:+.2f}{pstar(cs[k]['p'])}" if k in cs else "")
        trows.append(r_)
    caption(doc, "표 11. 단계별 계수(%p)")
    table(doc, ["변수", "M1", "M2", "M3", "M4"], trows, [5.0, 2.4, 2.4, 2.4, 2.4], size=7.5, align_right=(1, 2, 3, 4))

    # 8. 진단
    h1(doc, "8. 잔차 진단과 공간 자기상관")
    ci(doc, f"잔차 Moran's I(반경 500m 이웃, 이진 가중) = {A['moran']['I']}, z = {A['moran']['z']}, p = {A['moran']['p']}. "
            "모형이 소권역 무작위절편으로 이웃 간 닮음을 이미 흡수해, 남은 공간 자기상관은 유의하지 않다(공간오차모형까지는 불필요).")
    formula(doc, "I = (n ÷ S₀) · (zᵀ W z) ÷ (zᵀ z),   z = 잔차 − 평균,   W_ij = 1(0 < d_ij ≤ 500m),   S₀ = Σ W_ij")
    sk, ku = stats.skew(rs), stats.kurtosis(rs)
    ci(doc, f"잔차 왜도 {sk:.2f}, 첨도(초과) {ku:.2f}. 0% 공실 건물이 많아 아래쪽이 잘린 분포다. 표본이 266으로 계수 추정은 일관적이나 p값은 근사치로 보고, "
            "결론은 이항 모형과 287동 모형이 같은 방향인지로 판단했다.")
    caption(doc, "그림 8. 잔차 진단과 행정동 무작위효과")
    picture(doc, F["diag"], 16.5)
    caption(doc, "표 12. 행정동 무작위절편(BLUP) — 모형 변수로 설명하고 남은 동 고유 효과")
    table(doc, ["행정동", "û_k(%p)", "건물", "공실률 면적합(%)"],
          [[x["unit"], f"{x['re']:+.1f}", f"{x['n']}", f"{x['rate']:.1f}"] for x in A["dong_re"]], [4, 2.5, 2, 3.4], size=8, align_right=(1, 2, 3))
    note(doc, "M4에서 σ²_u 가 작아 BLUP 이 ±1%p 안으로 강하게 수축된다 — 변수를 통제하면 '어느 동이냐' 자체의 효과는 거의 남지 않는다.")

    # 9. GLMM
    h1(doc, "9. 교차 확인: 점포(호) 단위 이항 혼합모형")
    formula(doc, "logit Pr(공실_hijk = 1) = γ₀ + Σ_p γ_p x_p,ijk + b_i + v_jk + u_k")
    ci(doc, f"점포 한 칸(호) h 를 공실 1 / 영업 0 으로 펼친 {A['glmm']['n_units']:,}호. 건물(b_i)·소권역·행정동 무작위절편. "
            "변분 베이즈(statsmodels BinomialBayesMixedGLM). 칸이 많은 큰 건물이 자연히 큰 무게를 갖는다.")
    ci(doc, "승산비 OR = exp(γ). OR 1.10 = 공실 승산 10% 증가. 설명변수는 추정 안정을 위해 표준화했다가 원 단위로 되돌렸다.")
    gi = A["glmm"]["icc"]
    ci(doc, f"잠재척도 ICC(로지스틱 잔차분산 π²/3 기준): 건물 {gi['bld']}, 소권역 {gi['L2']}, 행정동 {gi['L3']}.")
    grows = [[name[k], f"/{unit[k]}", f"{gl[k]['or']:.3f}", f"[{gl[k]['lo']:.3f}, {gl[k]['hi']:.3f}]", "†" if gl[k]["sig"] else "",
              f"{m4j[k]['b']:+.2f}", "같음" if (gl[k]["or"] > 1) == (m4j[k]["b"] > 0) else "반대"] for k in keys]
    caption(doc, "표 13. 이항 혼합모형 승산비와 선형 M4 방향 비교")
    table(doc, ["변수", "단위", "OR", "95% 구간", "유의", "선형 β̂", "방향"], grows, [4.4, 1.6, 1.5, 3.0, 1.0, 1.6, 1.4], size=7.5,
          align_right=(2, 3, 5), align_center=(4, 6))
    agree = sum(1 for r_ in grows if r_[-1] == "같음")
    note(doc, f"21개 중 {agree}개가 같은 방향. 유의성은 이항 모형이 더 쉽게 나온다(관측치가 호 단위로 많기 때문) — 구간 폭보다 방향 일치를 본다.")

    # 10. 민감도
    h1(doc, "10. 민감도: 점포 정보가 없는 21동 포함(287동)")
    ci(doc, f"제외된 {A['excluded']['n']}동은 공실률 {A['excluded']['rate']}%로 매우 높다(점포가 없다는 것 자체가 공실의 결과일 수 있음). "
            "이들을 빼면 표본이 '덜 빈' 쪽으로 기울 수 있어, 업종 구성 묶음(M2)만 빼고 287동으로 다시 적합했다.")
    frows = []
    for s in A["full"]:
        cs = {c["key"]: c for c in s["coefs"]}
        frows.append([s["step"], f"{s['icc']['L3']:.3f}", f"{s['icc']['L2']:.3f}",
                      ", ".join(f"{name[k]} {c['b']:+.1f}{pstar(c['p'])}" for k, c in cs.items() if c["p"] < .05) or "—"])
    caption(doc, "표 14. 287동 모형 — 유의한 계수(p<0.05)")
    table(doc, ["단계", "ICC L3", "ICC L2", "유의한 변수(%p)"], frows, [1.4, 1.6, 1.6, 12.4], size=7.5, align_right=(1, 2))
    note(doc, "상황판 그림의 '287동 유의' 표시는 이 표의 M4에서 p<0.05인 변수다.")

    # 11. 한계
    h1(doc, "11. 해석상 주의와 한계")
    for t in [
        "단면 상관이다. 공실과 점포 구성은 동시에 결정되므로(빈 건물엔 특정 업종만 남는다) 업종 구성 계수는 원인보다 '공실 건물의 모습'에 가깝다.",
        "자료 시점이 다르다: 공실 2025.12, 점포 2026.06, 인구 2026.08, 종사자 2024. 상권 변화가 빠른 신도시라 수개월 차이도 영향이 있을 수 있다.",
        "종속변수는 0에 몰리고 0~100으로 묶인 비율이다. 선형모형은 평균 효과 해석이 쉬워 주 모형으로 쓰고, 분포 가정은 이항 모형으로 보완했다. "
        "분수 로짓·토빗·베타 회귀는 추가 검토 대상.",
        "L3 변수(행정동 3개)는 13개 값으로 추정해 검정력이 낮다. 유의하지 않음 ≠ 효과 없음.",
        "점포 매칭은 도로명 주소 완전 일치라, 주소 표기가 다른 점포는 빠질 수 있다(매칭 0곳 21동 중 일부는 이 이유일 가능성).",
        "프랜차이즈 = 지점명 기재 여부, 고령 = 60세 이상, 거리 = 직선거리 등 대리 지표의 한계가 있다.",
    ]:
        ci(doc, t)
    para(doc, "")
    note(doc, "재현: python prepare.py → python tools/extra_features.py → python multilevel.py → python tools/make_ml_appendix.py "
              "(sejong_vacant 폴더). 숫자는 data/multilevel.json 과 같다.")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("저장:", OUT)


if __name__ == "__main__":
    main()
