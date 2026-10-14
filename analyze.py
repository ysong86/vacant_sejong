# -*- coding: utf-8 -*-
"""건물별 공실률의 원인 분석 → data/analysis.json

  python analyze.py

1) 집계: 전체·생활권·동별 공실률(면적·호수)
2) 구간별 공실률: 정류장·BRT·중심상업지 거리, 경과연수, 주변 점포 밀도, 상가 비중, 용도
3) 가중 회귀(상가면적 가중, HC1 강건 표준오차): 요인별 순효과(%p)
4) 랜덤포레스트 + 순열 중요도: 비선형·상호작용을 포함한 교차 확인
5) 동별 '평균보다 높은(낮은) 이유' = 회귀계수 × (동 평균 − 전체 평균)
6) 공실 핫스팟: Getis-Ord Gi*(500m)
"""
from __future__ import annotations

import json
import math
import os

import numpy as np
import pandas as pd
from scipy import stats
from sklearn.ensemble import RandomForestRegressor
from sklearn.inspection import permutation_importance

HERE = os.path.dirname(os.path.abspath(__file__))

# 회귀에 넣는 요인 — (열, 화면 이름, 단위 설명)
FACTORS = [
    ("age", "준공 경과연수", "1년"),
    ("new5", "준공 5년 이하", "해당"),
    ("type2", "제2종 근린생활시설", "해당"),
    ("office", "업무시설(오피스텔 등)", "해당"),
    ("mixed_res", "주상복합(공동주택 하부 상가)", "해당"),
    ("shop_share", "연면적 중 상가 비중", "10%p"),
    ("log_area", "상가 규모(면적, 로그)", "2배"),
    ("d_stop", "버스정류장까지 거리", "100m"),
    ("d_brt", "BRT 정류장까지 거리", "100m"),
    ("d_center", "나성동 중심상업지까지 거리", "1km"),
    ("log_stores", "주변 200m 점포 수(로그)", "2배"),
    ("u_pop_k", "행정동 인구", "1만 명"),
    ("u_transit", "행정동 대중교통 분담률", "1%p"),
]
SCALE = {"age": 1, "new5": 1, "type2": 1, "office": 1, "mixed_res": 1, "shop_share": 10, "log_area": math.log(2),
         "d_stop": 100, "d_brt": 100, "d_center": 1000, "log_stores": math.log(2), "u_pop_k": 1, "u_transit": 1}


def wrate(g, kind="area"):
    if kind == "area":
        s = g["shop_area"].sum()
        return round(g["vac_area"].sum() / s * 100, 1) if s else None
    s = g["units"].sum()
    return round(g["vac_units"].sum() / s * 100, 1) if s else None


def load():
    B = pd.DataFrame(json.load(open(os.path.join(HERE, "data", "buildings.json"), encoding="utf-8")))
    B = B[B["shop_area"].fillna(0) > 0].copy()
    B["y"] = B["vac_area"] / B["shop_area"] * 100
    B["new5"] = (B["age"] <= 5).astype(float)
    B["type2"] = B["use"].str.contains("제2종").astype(float)
    B["office"] = B["use"].str.contains("업무").astype(float)
    B["mixed_res"] = B["mixed_res"].astype(float)
    B["log_area"] = np.log(B["shop_area"])
    B["log_stores"] = np.log1p(B["stores_200"])
    B["u_pop_k"] = B["u_pop"] / 10000
    return B


def wls_hc1(X, y, w):
    """가중 최소제곱 + HC1 강건 표준오차."""
    W = np.sqrt(w)[:, None]
    Xw, yw = X * W, y * W[:, 0]
    XtX_inv = np.linalg.pinv(Xw.T @ Xw)
    beta = XtX_inv @ Xw.T @ yw
    e = yw - Xw @ beta
    n, k = Xw.shape
    meat = (Xw * e[:, None]).T @ (Xw * e[:, None])
    cov = XtX_inv @ meat @ XtX_inv * n / (n - k)
    se = np.sqrt(np.diag(cov))
    yhat = X @ beta
    ybar = np.average(y, weights=w)
    r2 = 1 - np.sum(w * (y - yhat) ** 2) / np.sum(w * (y - ybar) ** 2)
    return beta, se, r2, n, k


def bins_table(B, col, edges, labels):
    c = pd.cut(B[col], edges, labels=labels, include_lowest=True)
    out = []
    for lab, g in B.groupby(c, observed=True):
        out.append({"label": str(lab), "rate": wrate(g), "n": int(len(g)), "area": round(g["shop_area"].sum())})
    return out


def gi_star(B, band=500.0):
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(HERE), "sejong_livingzone"))
    import collect as ZC
    xy = np.array([ZC._xy(lo, la) for lo, la in zip(B["lon"], B["lat"])])
    x = B["y"].values
    n, xb, s = len(x), x.mean(), x.std()
    D = np.sqrt(((xy[:, None, :] - xy[None, :, :]) ** 2).sum(-1))
    Wm = (D <= band).astype(float)
    sw = Wm.sum(1)
    num = Wm @ x - xb * sw
    den = s * np.sqrt((n * (Wm ** 2).sum(1) - sw ** 2) / (n - 1))
    return np.where(den > 0, num / den, 0)


def main():
    B = load()
    res = {"n_buildings": int(len(B)), "units": int(B["units"].sum()), "vac_units": int(B["vac_units"].sum()),
           "shop_area": round(B["shop_area"].sum()), "vac_area": round(B["vac_area"].sum()),
           "rate_area": wrate(B), "rate_units": wrate(B, "units"),
           "survey_years": sorted(B["survey"].unique().tolist())}
    res["by_zone"] = [{"zone": z, "rate": wrate(g), "rate_units": wrate(g, "units"), "n": int(len(g)),
                       "area": round(g["shop_area"].sum()), "vac_area": round(g["vac_area"].sum())}
                      for z, g in B.groupby("zone")]
    res["by_dong"] = [{"dong": d, "unit": g["unit"].iloc[0], "zone": g["zone"].iloc[0], "rate": wrate(g),
                       "rate_units": wrate(g, "units"), "n": int(len(g)), "units": int(g["units"].sum()),
                       "vac_units": int(g["vac_units"].sum()), "area": round(g["shop_area"].sum()),
                       "age": round(float(np.average(g["age"], weights=g["shop_area"])), 1),
                       "new5_share": round(float((g["new5"] * g["shop_area"]).sum() / g["shop_area"].sum() * 100))}
                      for d, g in B.groupby("dong")]
    res["by_sub"] = [{"sub": s.strip(), "dong": g["dong"].iloc[0], "rate": wrate(g), "n": int(len(g))}
                     for s, g in B.groupby("sub")]
    # 구간별
    res["bins"] = {
        "age": bins_table(B, "age", [0, 3, 5, 7, 9, 20], ["3년 이하", "3~5년", "5~7년", "7~9년", "9년 초과"]),
        "d_stop": bins_table(B, "d_stop", [0, 100, 200, 3000], ["100m 이내", "100~200m", "200m 초과"]),
        "d_brt": bins_table(B, "d_brt", [0, 200, 400, 600, 8000], ["200m 이내", "200~400m", "400~600m", "600m 초과"]),
        "d_center": bins_table(B, "d_center", [0, 1000, 2000, 3000, 9000], ["1km 이내", "1~2km", "2~3km", "3km 초과"]),
        "stores": bins_table(B, "stores_200", [0, 60, 100, 200, 5000], ["60곳 이하", "61~100곳", "101~200곳", "200곳 초과"]),
        "shop_share": bins_table(B, "shop_share", [0, 30, 50, 70, 200], ["30% 이하", "30~50%", "50~70%", "70% 초과"]),
    }
    ug = B["use"].where(B["type2"] == 1, B["use"].where(B["use"].str.contains("제1종"), "기타"))
    ug = ug.where(~B["office"].astype(bool), "업무시설").where(~B["mixed_res"].astype(bool), "주상복합")
    ug = ug.str.replace("제2종근린생활시설", "제2종 근린생활").str.replace("제1종근린생활시설", "제1종 근린생활")
    res["bins"]["use"] = [{"label": k, "rate": wrate(g), "n": int(len(g)), "area": round(g["shop_area"].sum())}
                          for k, g in B.groupby(ug)]

    # 회귀
    cols = [c for c, _, _ in FACTORS]
    D = B.dropna(subset=cols + ["y"]).copy()
    X = np.column_stack([np.ones(len(D))] + [D[c].values for c in cols])
    beta, se, r2, n, k = wls_hc1(X, D["y"].values, D["shop_area"].values)
    tcrit = stats.t.ppf(0.975, n - k)
    coefs = []
    for i, (c, name, unit) in enumerate(FACTORS, start=1):
        sc = SCALE[c]
        b, s = beta[i] * sc, se[i] * sc
        p = 2 * (1 - stats.t.cdf(abs(beta[i] / se[i]), n - k))
        coefs.append({"key": c, "name": name, "unit": unit, "b": round(b, 2), "lo": round(b - tcrit * s, 2),
                      "hi": round(b + tcrit * s, 2), "p": round(float(p), 4)})
    res["ols"] = {"n": int(n), "r2": round(float(r2), 3), "coefs": coefs}

    # 랜덤포레스트 교차 확인
    rf = RandomForestRegressor(n_estimators=600, min_samples_leaf=4, random_state=7, n_jobs=-1)
    rf.fit(D[cols].values, D["y"].values, sample_weight=D["shop_area"].values)
    pi = permutation_importance(rf, D[cols].values, D["y"].values, n_repeats=30, random_state=7,
                                sample_weight=D["shop_area"].values)
    oob = RandomForestRegressor(n_estimators=600, min_samples_leaf=4, random_state=7, oob_score=True, n_jobs=-1)
    oob.fit(D[cols].values, D["y"].values)
    res["rf"] = {"oob_r2": round(float(oob.oob_score_), 3),
                 "importance": sorted([{"key": c, "name": name, "imp": round(float(m), 2), "sd": round(float(s), 2)}
                                       for (c, name, _), m, s in zip(FACTORS, pi.importances_mean, pi.importances_std)],
                                      key=lambda x: -x["imp"])}

    # 동별 분해: 계수 × (동 가중평균 − 전체 가중평균)
    w = D["shop_area"]
    overall = {c: float(np.average(D[c], weights=w)) for c in cols}
    decomp = {}
    for d, g in D.groupby("dong"):
        parts = []
        for i, (c, name, _) in enumerate(FACTORS, start=1):
            diff = float(np.average(g[c], weights=g["shop_area"])) - overall[c]
            parts.append({"key": c, "name": name, "pp": round(beta[i] * diff, 1)})
        explained = sum(x["pp"] for x in parts)
        parts.sort(key=lambda x: -abs(x["pp"]))
        gap = round(float(wrate(g) - res["rate_area"]), 1)
        decomp[d] = {"gap": gap, "explained": round(explained, 1), "other": round(gap - explained, 1),
                     "parts": [dict(x, pp=float(x["pp"])) for x in parts[:4]]}
    res["decomp"] = decomp

    # 핫스팟
    B["gi"] = gi_star(B)
    res["hot"] = {"hot": int((B["gi"] > 1.96).sum()), "cold": int((B["gi"] < -1.96).sum())}

    pts = []
    for _, r in B.iterrows():
        pts.append({"name": r["name"], "dong": r["dong"], "sub": r["sub"], "zone": r["zone"], "lon": round(r["lon"], 6),
                    "lat": round(r["lat"], 6), "use": r["use"], "area": round(r["shop_area"]), "rate": round(r["y"], 1),
                    "units": int(r["units"]), "vac_units": int(r["vac_units"]), "age": r["age"],
                    "d_stop": int(r["d_stop"]), "d_brt": int(r["d_brt"]), "gi": round(float(r["gi"]), 2),
                    "survey": r["survey"]})
    res["points"] = pts
    json.dump(res, open(os.path.join(HERE, "data", "analysis.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("저장: data/analysis.json")
    print("전체", res["rate_area"], res["rate_units"], "R2", res["ols"]["r2"], "RF oob", res["rf"]["oob_r2"])
    for c in coefs:
        print(f"  {c['name']:22s} {c['b']:7.2f}%p / {c['unit']}  [{c['lo']}, {c['hi']}] p={c['p']}")
    print("RF", [(x["name"], x["imp"]) for x in res["rf"]["importance"][:6]])
    print("hot/cold", res["hot"])
    for d in ("대평동", "해밀동", "반곡동", "새롬동"):
        print(d, decomp.get(d))


if __name__ == "__main__":
    main()
