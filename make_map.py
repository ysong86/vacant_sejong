# -*- coding: utf-8 -*-
"""행복도시 건물 3D 지도 + 상가공실 → map3d.html

  python tools/fetch_buildings.py   (건물 외곽선, 처음 한 번)
  python prepare.py                 (조사 엑셀 → data/buildings.json)
  python make_map.py

조사 건물은 좌표가 들어 있는 건물 외곽선에 붙인다(없으면 40m 안의 가장 가까운 외곽선).
"""
from __future__ import annotations

import json
import math
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ZONE = os.path.join(os.path.dirname(HERE), "sejong_livingzone")
ZNAME = {"1": "1생활권", "2": "2생활권", "3": "3생활권", "4": "4생활권", "5": "5생활권", "6": "6생활권", "S": "세종동"}


def inside(x, y, ring):
    c = False
    j = len(ring) - 1
    for i in range(len(ring)):
        xi, yi = ring[i]
        xj, yj = ring[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi) + xi:
            c = not c
        j = i
    return c


def m_dist(a, b):
    dx = (a[0] - b[0]) * 111320 * math.cos(math.radians(36.5))
    dy = (a[1] - b[1]) * 110940
    return math.hypot(dx, dy)


def label_pt(rings):
    """가장 큰 고리의 면적 중심(라벨 자리)."""
    best = None
    for r in rings:
        a = cx = cy = 0.0
        for (x0, y0), (x1, y1) in zip(r, r[1:] + r[:1]):
            c = x0 * y1 - x1 * y0
            a += c
            cx += (x0 + x1) * c
            cy += (y0 + y1) * c
        if a and (best is None or abs(a) > best[0]):
            best = (abs(a), [round(cx / (3 * a), 5), round(cy / (3 * a), 5)])
    return best[1]


def compact(data):
    """상황판(build.py)에 끼울 가벼운 자료 → data/map3d.json.
    건물 = [높이(dm), 조사번호(0=없음), 고리...], 고리 = 좌표 정수 평탄화((경도-127)·1e5, (위도-36)·1e5)."""
    B = []
    for f in data["footprints"]["features"]:
        p = f["properties"]
        rings = []
        for poly in f["geometry"]["coordinates"]:
            for r in poly[:1]:          # 안뜰 구멍은 3D 에서 거의 안 보여 바깥 고리만
                flat = []
                for x, y in r:
                    flat += [round((x - 127) * 1e5), round((y - 36) * 1e5)]
                rings.append(flat)
        B.append([round(p["h"] * 10), p.get("sv") or 0] + rings)
    S = [{k: s[k] for k in ("no", "name", "dong", "sub", "zone", "use", "addr", "area", "vac", "units", "vu", "rate", "ru",
                             "survey", "age", "lon", "lat")} for s in data["survey"]]
    out = {"b": B, "s": S, "z": data["zones"]}
    path = os.path.join(HERE, "data", "map3d.json")
    json.dump(out, open(path, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    print("상황판용:", path, round(os.path.getsize(path) / 1e6, 2), "MB")


def main():
    fp = json.load(open(os.path.join(HERE, "data", "footprints.json"), encoding="utf-8"))
    B = json.load(open(os.path.join(HERE, "data", "buildings.json"), encoding="utf-8"))
    Z = [z for z in json.load(open(os.path.join(ZONE, "assets", "sejong_zones.json"), encoding="utf-8"))["zones"]
         if z["id"] != "R"]

    # 외곽선 격자 색인(약 200m)
    feats = fp["features"]
    cen, grid = [], {}
    for i, f in enumerate(feats):
        ring = f["geometry"]["coordinates"][0][0]
        c = (sum(p[0] for p in ring) / len(ring), sum(p[1] for p in ring) / len(ring))
        cen.append(c)
        grid.setdefault((int(c[0] / 0.002), int(c[1] / 0.002)), []).append(i)

    def near(lon, lat):
        gx, gy = int(lon / 0.002), int(lat / 0.002)
        return [i for dx in (-1, 0, 1) for dy in (-1, 0, 1) for i in grid.get((gx + dx, gy + dy), [])]

    survey, matched = [], 0
    for r in B:
        if r.get("lon") is None or not r.get("shop_area"):
            continue
        lon, lat = r["lon"], r["lat"]
        cand = near(lon, lat)
        hit = [i for i in cand if any(inside(lon, lat, poly[0]) for poly in feats[i]["geometry"]["coordinates"])]
        if not hit:
            d = sorted((m_dist((lon, lat), cen[i]), i) for i in cand)
            hit = [d[0][1]] if d and d[0][0] <= 40 else []
        rate = round(r["vac_area"] / r["shop_area"] * 100, 1)
        rec = {"no": r["no"], "name": r["name"], "dong": r["dong"], "sub": r["sub"], "zone": r.get("zone"),
               "use": r["use"], "addr": r["addr"], "area": round(r["shop_area"]), "vac": round(r["vac_area"]),
               "units": r["units"], "vu": r["vac_units"], "rate": rate,
               "ru": round(r["vac_units"] / r["units"] * 100, 1) if r["units"] else None,
               "survey": r["survey"], "age": r.get("age"), "lon": round(lon, 6), "lat": round(lat, 6)}
        if hit:
            matched += 1
            i = hit[0]
            p = feats[i]["properties"]
            if "sv" in p and feats[i]["properties"]["sv_area"] >= rec["area"]:
                pass   # 한 외곽선에 조사 건물 둘 → 상가가 큰 쪽을 남긴다
            else:
                p["sv"] = r["no"]
                p["sv_area"] = rec["area"]
            rec["h"] = p["h"]
        survey.append(rec)
    for f in feats:
        f["properties"].pop("sv_area", None)
        for k in ("pnu", "fl"):
            f["properties"].pop(k, None)
    print("조사 건물", len(survey), "외곽선 연결", matched)

    data = {"footprints": fp, "survey": survey,
            "zones": [{"id": z["id"], "name": ZNAME.get(z["id"], z["id"]), "rings": z["rings"], "lab": label_pt(z["rings"])}
                      for z in Z]}
    compact(data)
    tpl = open(os.path.join(HERE, "template_map.html"), encoding="utf-8").read()
    html = tpl.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    out = os.path.join(HERE, "map3d.html")
    open(out, "w", encoding="utf-8").write(html)
    print("저장:", out, round(len(html) / 1e6, 1), "MB")


if __name__ == "__main__":
    main()
