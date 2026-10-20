# -*- coding: utf-8 -*-
"""상가공실 멀티레벨(3수준) 분석 → data/multilevel.json

  python tools/extra_features.py   (추가 변수, 처음 한 번)
  python multilevel.py

구조: 건물(수준 1, 266동) ⊂ 소권역(수준 2, 동 세부 26곳) ⊂ 행정동(수준 3, 13곳 — 도담·어진 합침)

1) 주 모형 — 선형 혼합모형(LMM, REML): 공실률(면적, %) = 고정효과 + 행정동 무작위절편 + 소권역 무작위절편 + 오차
   계수를 '공실률 %p'로 읽을 수 있어 기존 회귀(analyze.py)와 같은 눈금이다.
   단계: M0(빈 모형) → M1(기존 건물 요인) → M2(+건물 내부: 층·주차·업종) → M3(+반경 수요·공급) → M4(+행정동 여건)
   각 단계에서 수준별 분산이 얼마나 줄었는지(PCV)로 '어느 수준의 차이를 무엇이 설명하는가'를 본다.
2) 교차 확인 — 이항 혼합모형(GLMM, 변분 베이즈): 점포 한 칸(호)을 공실/아님(0/1)으로 보고
   건물·소권역·행정동 무작위절편. 큰 건물이 더 큰 무게를 갖는다(면적 가중 회귀와 같은 취지). 결과는 승산비.
3) 잔차 공간 자기상관: Moran's I(반경 500m)
4) 민감도: 점포 정보가 없는 21동(공실 55.9% — 점포가 없는 것 자체가 공실의 결과)까지 287동으로 M1 재추정

점포 '개수'는 공실의 거울이라 쓰지 않고 업종 '구성 비율'만 쓴다. 상가정보는 현재 시점, 공실 조사는 2025년 12월이다.
"""
from __future__ import annotations

import json
import math
import os
import warnings

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from scipy import stats
from statsmodels.genmod.bayes_mixed_glm import BinomialBayesMixedGLM

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))

# (열, 화면 이름, 단위 설명, 수준) — 계수는 이 단위로 바꿔 보고한다(원 변수는 아래 prep 에서 이 단위로 만든다)
BLOCKS = {
    "M1": [("new5", "준공 5년 이하", "해당", "건물"), ("type2", "제2종 근린생활시설", "해당", "건물"),
           ("office", "업무시설(오피스텔 등)", "해당", "건물"), ("mixed_res", "주상복합", "해당", "건물"),
           ("shop_share10", "연면적 중 상가 비중", "10%p", "건물"), ("log2_area", "상가 규모", "2배", "건물"),
           ("d_stop100", "버스정류장 거리", "100m", "건물"), ("d_center1k", "나성동 중심 거리", "1km", "건물")],
    "M2": [("floors5", "지상 층수", "5층", "건물"), ("log2_park", "상가 100㎡당 주차", "2배", "건물"),
           ("food10", "음식점 비중", "10%p", "건물"), ("edu10", "학원·교육 비중", "10%p", "건물"),
           ("med10", "병의원 비중", "10%p", "건물"), ("div01", "업종 다양도", "0.1", "건물"),
           ("fran10", "프랜차이즈 비중", "10%p", "건물"), ("up3_10", "3층 이상 점포 비중", "10%p", "건물")],
    "M3": [("log2_sph", "반경 500m 세대당 상가면적", "2배", "반경"), ("ppm100", "배후 아파트 ㎡당 가격", "100만원", "반경")],
    "M4": [("chg1", "행정동 인구 증감률", "1%p", "행정동"), ("old1", "행정동 고령 비중", "1%p", "행정동"),
           ("emp01", "행정동 직주비(종사자÷인구)", "0.1", "행정동")],
}
STEPS = ["M0", "M1", "M2", "M3", "M4"]
STEP_NAME = {"M0": "빈 모형", "M1": "+ 기존 건물 요인", "M2": "+ 건물 내부(층·주차·업종)", "M3": "+ 반경 수요·공급",
             "M4": "+ 행정동 여건"}


def prep():
    B = pd.DataFrame(json.load(open(os.path.join(HERE, "data", "buildings_plus.json"), encoding="utf-8")))
    B["y"] = B["vac_area"] / B["shop_area"] * 100
    B["new5"] = (B["age"] <= 5).astype(float)
    B["type2"] = B["use"].str.contains("제2종").astype(float)
    B["office"] = B["use"].str.contains("업무").astype(float)
    B["mixed_res"] = B["mixed_res"].astype(float)
    B["shop_share10"] = B["shop_share"] / 10
    B["log2_area"] = np.log2(B["shop_area"])
    B["d_stop100"] = B["d_stop"] / 100
    B["d_center1k"] = B["d_center"] / 1000
    med = lambda c: B[c].fillna(B[c].median())
    B["floors5"] = med("floors") / 5
    B["log2_park"] = np.log2(med("park_per_100") + 0.5)
    B["food10"] = B["sh_food"] / 10
    B["edu10"] = B["sh_edu"] / 10
    B["med10"] = B["sh_med"] / 10
    B["div01"] = B["diversity"] * 10
    B["fran10"] = B["franchise"] / 10
    B["up3_10"] = med("sh_3up") / 10
    B["log2_sph"] = np.log2(B["shop_per_hh"].fillna(B["shop_per_hh"].median()))
    B["ppm100"] = B["apt_ppm"] / 100
    B["chg1"] = B["u_chg"]
    B["old1"] = B["u_old"]
    B["emp01"] = B["u_emp"] * 10
    B["L3"] = B["unit"]
    B["L2"] = B["unit"] + "|" + B["sub"]
    return B


def cols_upto(step):
    out = []
    for s in STEPS[1:STEPS.index(step) + 1]:
        out += BLOCKS[s]
    return out


def fit_lmm(D, step, reml=True):
    cols = [c for c, *_ in cols_upto(step)]
    f = "y ~ " + (" + ".join(cols) if cols else "1")
    m = smf.mixedlm(f, D, groups="L3", re_formula="1", vc_formula={"L2": "0 + C(L2)"})
    return m.fit(reml=reml, method=["lbfgs", "powell"])


def var_parts(r):
    return {"L3": float(r.cov_re.iloc[0, 0]), "L2": float(r.vcomp[0]), "L1": float(r.scale)}


def r2_nakagawa(r, D):
    vf = float(np.var(r.model.exog @ r.fe_params, ddof=0))
    v = var_parts(r)
    tot = vf + v["L3"] + v["L2"] + v["L1"]
    return vf / tot, (vf + v["L3"] + v["L2"]) / tot


def morans_i(D, resid, band=500.0):
    lat0 = math.radians(36.5)
    x = D["lon"].values * 111320 * math.cos(lat0)
    y = D["lat"].values * 110940
    d = np.sqrt((x[:, None] - x[None, :]) ** 2 + (y[:, None] - y[None, :]) ** 2)
    W = ((d <= band) & (d > 0)).astype(float)
    z = resid - resid.mean()
    n, s0 = len(z), W.sum()
    I = n / s0 * (z @ W @ z) / (z @ z)
    # 정규 근사 검정
    EI = -1 / (n - 1)
    s1 = 0.5 * ((W + W.T) ** 2).sum()
    s2 = ((W.sum(0) + W.sum(1)) ** 2).sum()
    VI = (n * n * s1 - n * s2 + 3 * s0 * s0) / ((n * n - 1) * s0 * s0) - EI ** 2
    zI = (I - EI) / math.sqrt(VI)
    return float(I), float(zI), float(2 * (1 - stats.norm.cdf(abs(zI))))


def glmm(D):
    """호(점포 칸) 단위 이항 혼합모형 — 건물·소권역·행정동 무작위절편(변분 베이즈)."""
    cols = [c for c, *_ in cols_upto("M4")]
    rows = []
    for _, r in D.iterrows():
        base = {c: r[c] for c in cols}
        base.update({"bld": str(r["no"]), "L2": r["L2"], "L3": r["L3"]})
        rows += [dict(base, v=1)] * int(r["vac_units"]) + [dict(base, v=0)] * int(r["units"] - r["vac_units"])
    U = pd.DataFrame(rows)
    Z = U[cols]
    mu, sd = Z.mean(), Z.std(ddof=0)
    U[cols] = (Z - mu) / sd        # 변분 추정 안정을 위해 표준화한 뒤 원 단위로 되돌린다
    f = "v ~ " + " + ".join(cols)
    m = BinomialBayesMixedGLM.from_formula(f, {"bld": "0 + C(bld)", "L2": "0 + C(L2)", "L3": "0 + C(L3)"}, U)
    r = m.fit_vb()
    names = m.exog_names
    out = []
    for c, name, unit, lev in cols_upto("M4"):
        i = names.index(c)
        b, s = r.fe_mean[i] / sd[c], r.fe_sd[i] / sd[c]
        out.append({"key": c, "name": name, "unit": unit, "level": lev, "or": round(math.exp(b), 3),
                    "lo": round(math.exp(b - 1.96 * s), 3), "hi": round(math.exp(b + 1.96 * s), 3),
                    "sig": bool(abs(b / s) > 1.96)})
    vc = {n: float(np.exp(2 * m_)) for n, m_ in zip(m.vcp_names, r.vcp_mean)}
    lat = math.pi ** 2 / 3
    tot = sum(vc.values()) + lat
    return {"n_units": int(len(U)), "coefs": out, "var": {k: round(v, 3) for k, v in vc.items()},
            "icc": {k: round(v / tot, 3) for k, v in vc.items()}}


def main():
    B = prep()
    full = B.copy()
    D = B[B["stores"] > 0].reset_index(drop=True)
    res = {"n": int(len(D)), "n_full": int(len(full)), "n_L3": int(D["L3"].nunique()), "n_L2": int(D["L2"].nunique()),
           "excluded": {"n": int((B["stores"] == 0).sum()),
                        "rate": round(float(B.loc[B["stores"] == 0, "vac_area"].sum() / B.loc[B["stores"] == 0, "shop_area"].sum() * 100), 1)},
           "steps": [], "blocks": {k: [{"key": c, "name": n, "unit": u, "level": l} for c, n, u, l in v] for k, v in BLOCKS.items()}}
    # 기술통계(새 변수)
    res["desc"] = {}
    for c, name, unit, lev in [x for s in ("M2", "M3", "M4") for x in BLOCKS[s]]:
        v = D[c]
        res["desc"][c] = {"mean": round(float(v.mean()), 3), "sd": round(float(v.std()), 3), "min": round(float(v.min()), 3),
                          "max": round(float(v.max()), 3)}
    raw_desc = {"floors": "지상 층수", "park_per_100": "상가 100㎡당 주차(대)", "sh_food": "음식점 비중(%)", "sh_edu": "학원·교육 비중(%)",
                "sh_med": "병의원 비중(%)", "diversity": "업종 다양도(0~1)", "franchise": "프랜차이즈 비중(%)", "sh_3up": "3층 이상 점포 비중(%)",
                "shop_per_hh": "반경 500m 세대당 상가면적(㎡)", "hh500": "반경 500m 공동주택 세대", "apt_ppm": "배후 아파트 ㎡당 가격(만원)",
                "u_chg": "행정동 인구 증감률(%)", "u_old": "행정동 고령 비중(%)", "u_emp": "행정동 직주비"}
    res["raw_desc"] = [{"key": k, "name": n, "n": int(D[k].notna().sum()), "mean": round(float(D[k].mean()), 2),
                        "sd": round(float(D[k].std()), 2), "min": round(float(D[k].min()), 2), "max": round(float(D[k].max()), 2)}
                       for k, n in raw_desc.items()]

    base = None
    for step in STEPS:
        r = fit_lmm(D, step)
        rml = fit_lmm(D, step, reml=False)
        v = var_parts(r)
        if base is None:
            base = v
        tot = sum(v.values())
        r2m, r2c = r2_nakagawa(r, D) if step != "M0" else (0.0, (v["L3"] + v["L2"]) / tot)
        coefs = []
        for c, name, unit, lev in cols_upto(step):
            b, se = float(r.fe_params[c]), float(r.bse_fe[c])
            p = float(r.pvalues[c])
            coefs.append({"key": c, "name": name, "unit": unit, "level": lev, "b": round(b, 2), "lo": round(b - 1.96 * se, 2),
                          "hi": round(b + 1.96 * se, 2), "p": round(p, 4)})
        res["steps"].append({
            "step": step, "name": STEP_NAME[step], "k": len(coefs), "aic": round(float(rml.aic), 1),
            "var": {k: round(x, 2) for k, x in v.items()},
            "icc": {"L3": round(v["L3"] / tot, 3), "L2": round(v["L2"] / tot, 3), "L23": round((v["L3"] + v["L2"]) / tot, 3)},
            "pcv": {k: (round(1 - v[k] / base[k], 3) if base[k] > 1e-6 else None) for k in v},
            "r2m": round(r2m, 3), "r2c": round(r2c, 3), "coefs": coefs})
        if step == "M4":
            D["resid"] = r.resid
            D["re3"] = D["L3"].map({g: float(e.iloc[0]) for g, e in r.random_effects.items()})
            res["moran"] = dict(zip(("I", "z", "p"), [round(x, 4) for x in morans_i(D, D["resid"].values)]))
            # 행정동 무작위효과(조건부) — 모형의 요인으로 설명하고 남은 동 고유 효과
            res["dong_re"] = sorted([{"unit": g, "re": round(float(e.iloc[0]), 1), "n": int((D["L3"] == g).sum()),
                                      "rate": round(float(D.loc[D["L3"] == g, "vac_area"].sum() / D.loc[D["L3"] == g, "shop_area"].sum() * 100), 1)}
                                     for g, e in r.random_effects.items()], key=lambda x: -x["re"])
    # 민감도: 287동 전체 — 점포 구성(M2 묶음)만 빼고 같은 단계로(점포 없는 고공실 21동을 넣어도 결론이 같은가)
    res["full"] = []
    fb = None
    for step in ("M0", "M1", "M3", "M4"):
        cols = [c for c, *_ in cols_upto(step) if c not in [x[0] for x in BLOCKS["M2"]]]
        m = smf.mixedlm("y ~ " + (" + ".join(cols) if cols else "1"), full, groups="L3", re_formula="1",
                        vc_formula={"L2": "0 + C(L2)"}).fit(reml=True, method=["lbfgs", "powell"])
        v = var_parts(m)
        fb = fb or v
        tot = sum(v.values())
        res["full"].append({"step": step, "icc": {"L3": round(v["L3"] / tot, 3), "L2": round(v["L2"] / tot, 3)},
                            "pcv": {k: (round(1 - v[k] / fb[k], 3) if fb[k] > 1e-6 else None) for k in v},
                            "coefs": [{"key": c, "b": round(float(m.fe_params[c]), 2), "p": round(float(m.pvalues[c]), 4)} for c in cols]})
    # 이항 GLMM 교차 확인
    res["glmm"] = glmm(D)
    json.dump(res, open(os.path.join(HERE, "data", "multilevel.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print(f"n={res['n']} (L2 {res['n_L2']}, L3 {res['n_L3']})  제외 {res['excluded']}")
    for st in res["steps"]:
        print(f"{st['step']} {st['name']:18s} AIC {st['aic']:8.1f}  ICC 동 {st['icc']['L3']:.3f} 소권역 {st['icc']['L2']:.3f}  "
              f"PCV {st['pcv']}  R2m {st['r2m']} R2c {st['r2c']}")
    for c in res["steps"][-1]["coefs"]:
        print(f"   {c['level']:3s} {c['name']:22s} {c['b']:+7.2f}%p/{c['unit']:6s} [{c['lo']:+.1f},{c['hi']:+.1f}] p={c['p']:.3f}")
    print("Moran", res["moran"])
    print("GLMM ICC", res["glmm"]["icc"])
    for c in res["glmm"]["coefs"]:
        if c["sig"]:
            print(f"   GLMM {c['name']:22s} OR {c['or']} [{c['lo']},{c['hi']}]")
    print("동 효과", res["dong_re"])
    for st in res["full"]:
        print("287동", st["step"], st["icc"], st["pcv"], [(c["key"], c["b"], c["p"]) for c in st["coefs"] if c["p"] < 0.1])


if __name__ == "__main__":
    main()
