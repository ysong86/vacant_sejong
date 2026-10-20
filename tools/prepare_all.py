# -*- coding: utf-8 -*-
"""조사 엑셀 네 시트(2023.06·2025.12 × 집합상가·아파트 상가) → data/survey_all.json

  python tools/prepare_all.py        (원본: data/raw/survey_2325.xlsx = 「상황판 데이터_23,25_최종.xlsx」)

기록 하나 = 한 조사 시점의 한 건물(집합상가) 또는 한 단지 상가(아파트 상가).
- yr: "2023"(2023년 6월) / "2025"(2025년 12월),  kind: "mall"(집합상가) / "apt"(아파트 상가)
- 좌표: 브이월드 주소→좌표(캐시 data/raw/geo_cache.json).
- 3D 건물 연결(fp: data/footprints.json 의 순번 목록)
  · 집합상가: 좌표가 들어 있는 외곽선(없으면 40m 안 가장 가까운 것).
  · 아파트 상가: 단지 안 상가동은 주소 점이 가리키지 않는다(주소 점은 단지 대표 위치).
    주소 점이 든 필지(PNU)를 찾고, 건축물대장 표제부로 그 필지의 근린생활시설 동 수(n)를 센 뒤,
    같은 필지의 낮은 건물(15m 이하) 가운데 바닥면적이 큰 n 개를 상가동으로 본다. 못 찾으면 fp 는 빈칸(지도에서 표식).
- 2025.12 시트의 '비고'(2023/2025)는 조사연도가 아니라 그 건물이 처음 조사 목록에 든 해로 보인다
  (같은 건물의 2023.06 시트 값과 2025.12 시트 값이 다르다) → first 로만 남긴다.
"""
from __future__ import annotations

import datetime as dt
import json
import math
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "tools"))
import prepare as PR          # noqa: E402  DONG_ZONE·DONG_UNIT·ZC(생활권 수집 도구)
import extra_features as EF   # noqa: E402  inside·dist·titles(표제부 API, 캐시)

ZC = PR.ZC
SRC = os.path.join(HERE, "data", "raw", "survey_2325.xlsx")
REF = {"2023": dt.date(2023, 6, 30), "2025": dt.date(2025, 12, 31)}


def num(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def rows_of(ws):
    rows = [r for r in ws.iter_rows(values_only=True) if any(v not in (None, "") for v in r)]
    head = [str(h).strip() if h is not None else "" for h in rows[0]]
    out = []
    for r in rows[1:]:
        d = {}
        for i, h in enumerate(head):
            if not h:
                continue
            key = h if h not in d else h + "2"          # '비고' 가 두 번 나오는 시트
            d[key] = r[i] if i < len(r) else None
        if d.get("NO") in (None, ""):
            continue
        out.append(d)
    return out


def age_of(apr, yr):
    s = str(apr or "").strip()[:8]
    try:
        d = dt.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
        return round((REF[yr] - d).days / 365.25, 1)
    except (ValueError, IndexError):
        return None


def load():
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    sheets = {ws.title: rows_of(ws) for ws in wb.worksheets}
    recs = []
    for title, rows in sheets.items():
        yr = "2023" if title.startswith("2023") else "2025"
        kind = "apt" if "아파트" in title else "mall"
        for r in rows:
            if kind == "mall":
                name, sub, code, use = r.get("건물명"), r.get("동 세부"), r.get("건물세부코드"), r.get("건물 주용도")
                area, vac, units, vu = num(r.get("상가면적")), num(r.get("공실_면적")), num(r.get("호수")), num(r.get("공실_호수"))
            else:
                apt, danji = (r.get("아파트명") or "").strip(), (r.get("단지명") or "").strip()
                name = f"{danji} {apt}".strip()
                sub = (r.get("상가동명") or "단지 내 상가")
                code = r.get("단지세부코드") or f"{danji}|{(r.get('주소') or '').strip()}"
                use = "아파트 단지 내 상가"
                area = num(r.get("상가면적"))
                vac = num(r.get("공실_면적") if r.get("공실_면적") is not None else r.get("공실면적"))
                units = num(r.get("호수"))
                vu = num(r.get("공실_호수") if r.get("공실_호수") is not None else r.get("공실호수"))
            if not area:
                continue
            dong = (r.get("동") or "").strip()
            first = r.get("비고") if yr == "2025" and isinstance(r.get("비고"), (int, float)) else None
            recs.append({
                "yr": yr, "kind": kind, "code": str(code).strip(), "name": str(name).replace("\xa0", " ").strip(),
                "sub": str(sub).strip(), "dong": dong, "zone": PR.DONG_ZONE.get(dong), "unit": PR.DONG_UNIT.get(dong, dong),
                "addr": str(r.get("주소") or "").strip(), "use": str(use or "").replace("\xa0", " ").strip(),
                "gfa": num(r.get("연면적")), "area": round(area, 1), "vac": round(vac or 0, 1),
                "units": int(units or 0), "vu": int(vu or 0),
                "rate": round((vac or 0) / area * 100, 1), "ru": round(vu / units * 100, 1) if units else None,
                "apr": str(r.get("사용승인일자") or "").strip()[:8], "age": age_of(r.get("사용승인일자"), yr),
                "first": int(first) if first else None,
            })
    return recs


def geocode(recs):
    with open(os.path.join(PR.ZONE, "config.json"), encoding="utf-8") as f:
        key = json.load(f)["vworld_key"]
    cpath = os.path.join(HERE, "data", "raw", "geo_cache.json")
    cache = json.load(open(cpath, encoding="utf-8")) if os.path.exists(cpath) else {}
    fail = []
    for i, r in enumerate(recs):
        xy = ZC._geocode("세종특별자치시 " + r["addr"], key, cache) if r["addr"] else None
        if not xy:
            xy = ZC._geocode("세종특별자치시 " + r["dong"] + " " + r["name"].split()[0], key, cache)
        r["lon"], r["lat"] = (round(xy[0], 6), round(xy[1], 6)) if xy else (None, None)
        if not xy:
            fail.append(f"{r['yr']}{r['kind']} {r['name']}")
        if i % 40 == 0:
            json.dump(cache, open(cpath, "w", encoding="utf-8"), ensure_ascii=False)
    json.dump(cache, open(cpath, "w", encoding="utf-8"), ensure_ascii=False)
    return fail


def ring_area(ring):
    k = 111320 * math.cos(math.radians(36.5))
    s = 0.0
    for (x0, y0), (x1, y1) in zip(ring, ring[1:] + ring[:1]):
        s += (x0 * k) * (y1 * 110940) - (x1 * k) * (y0 * 110940)
    return abs(s) / 2


def link_footprints(recs):
    feats = json.load(open(os.path.join(HERE, "data", "footprints.json"), encoding="utf-8"))["features"]
    cen, grid, by_pnu = [], {}, {}
    for i, f in enumerate(feats):
        ring = f["geometry"]["coordinates"][0][0]
        c = (sum(p[0] for p in ring) / len(ring), sum(p[1] for p in ring) / len(ring))
        cen.append(c)
        grid.setdefault((int(c[0] / .002), int(c[1] / .002)), []).append(i)
        by_pnu.setdefault(f["properties"].get("pnu") or "", []).append(i)
    near = lambda x, y: [i for dx in (-1, 0, 1) for dy in (-1, 0, 1) for i in grid.get((int(x / .002) + dx, int(y / .002) + dy), [])]

    def hit(x, y, maxd=40):
        cand = near(x, y)
        h = [i for i in cand if any(EF.inside(x, y, poly[0]) for poly in feats[i]["geometry"]["coordinates"])]
        if h:
            return h[0]
        d = sorted((EF.dist((x, y), cen[i]), i) for i in cand)
        return d[0][1] if d and d[0][0] <= maxd else None

    # 집합상가: 점이 든 외곽선
    for r in recs:
        r["fp"] = []
        if r["lon"] is None:
            continue
        if r["kind"] == "mall":
            i = hit(r["lon"], r["lat"])
            if i is not None:
                r["fp"] = [i]
    mall_fp = {i for r in recs if r["kind"] == "mall" for i in r["fp"]}
    # 아파트 상가: 단지 필지 → 표제부 근생 동 수 → 같은 필지의 낮은 건물 중 큰 것부터
    apt = [r for r in recs if r["kind"] == "apt" and r["lon"] is not None]
    pnus = {}
    for r in apt:
        i = hit(r["lon"], r["lat"], 80)
        pnus[id(r)] = feats[i]["properties"].get("pnu") if i is not None else None
    T = EF.titles([p for p in pnus.values() if p])
    stat = {"apt": len(apt), "linked": 0}
    for r in apt:
        pnu = pnus[id(r)]
        if not pnu:
            continue
        rows = T.get(pnu) or []
        shop = [x for x in rows if "근린생활" in (x.get("mainPurpsCdNm") or "") or "판매" in (x.get("mainPurpsCdNm") or "")]
        n = max(1, len(shop)) if rows else 1
        cand = [i for i in by_pnu.get(pnu, []) if i not in mall_fp and feats[i]["properties"]["h"] <= 15]
        cand.sort(key=lambda i: -ring_area(feats[i]["geometry"]["coordinates"][0][0]))
        r["fp"] = cand[:n]
        r["fp_note"] = f"필지 {pnu} 근생 {len(shop)}동/대장 {len(rows)}동"
        if r["fp"]:
            stat["linked"] += 1
    return stat


def main():
    recs = load()
    fail = geocode(recs)
    stat = link_footprints(recs)
    # 2025.12 집합상가에는 원인 분석(analysis.json)의 핫스팟 값 등을 붙인다(같은 건물·같은 값)
    an = json.load(open(os.path.join(HERE, "data", "analysis.json"), encoding="utf-8"))
    gi = {(p["name"], p["dong"]): p for p in an["points"]}
    for r in recs:
        if r["yr"] == "2025" and r["kind"] == "mall":
            p = gi.get((r["name"], r["dong"]))
            if p:
                r["gi"] = p.get("gi")
    for i, r in enumerate(recs, start=1):
        r["id"] = i
    json.dump(recs, open(os.path.join(HERE, "data", "survey_all.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    import collections
    c = collections.Counter((r["yr"], r["kind"]) for r in recs)
    print("기록", dict(c), "좌표 실패", len(fail), fail[:8])
    print("아파트 상가 3D 연결", stat, "집합상가 3D 연결", sum(1 for r in recs if r["kind"] == "mall" and r["fp"]),
          "/", sum(1 for r in recs if r["kind"] == "mall"))
    for yr in ("2023", "2025"):
        for k in ("mall", "apt"):
            g = [r for r in recs if r["yr"] == yr and r["kind"] == k]
            a = sum(r["area"] for r in g)
            print(yr, k, len(g), "공실률(면적)", round(sum(r["vac"] for r in g) / a * 100, 1) if a else None,
                  "(호수)", round(sum(r["vu"] for r in g) / max(1, sum(r["units"] for r in g)) * 100, 1))


if __name__ == "__main__":
    main()
