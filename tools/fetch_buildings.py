# -*- coding: utf-8 -*-
"""행복도시(생활권 1~6·S) 건물 외곽선 + 높이 → data/raw/bldg_tiles/*.json, data/footprints.json

  python tools/fetch_buildings.py

- 브이월드 WFS lt_c_bldginfo(GIS건물통합정보). 키는 생활권 상황판 config.json 의 vworld_key.
- WFS 는 키에 등록된 주소(ysong86.github.io)를 domain·Referer 로 보내야 한다. 안 맞으면 INCORRECT_KEY.
- 높이: height(m) → 없으면 지상층수×3.3m → 둘 다 없으면 4m.
"""
from __future__ import annotations

import json
import os
import sys
import time
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ZONE = os.path.join(os.path.dirname(HERE), "sejong_livingzone")
TILES = os.path.join(HERE, "data", "raw", "bldg_tiles")
STEP = 0.006          # 타일 한 변(도) ≈ 550~650m
PAGE = 1000
REF = "https://ysong86.github.io/"   # 키에 등록된 서비스 주소 — domain 인자와 Referer 둘 다 맞아야 한다


def key():
    with open(os.path.join(ZONE, "config.json"), encoding="utf-8") as f:
        return json.load(f)["vworld_key"]


def zones():
    d = json.load(open(os.path.join(ZONE, "assets", "sejong_zones.json"), encoding="utf-8"))
    return [z for z in d["zones"] if z["id"] != "R"]


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


def zone_of(x, y, Z):
    for z in Z:
        if any(inside(x, y, r) for r in z["rings"]):
            return z["id"]
    return None


def wfs(k, bbox, start):
    q = dict(service="WFS", version="1.1.0", request="GetFeature", typename="lt_c_bldginfo",
             bbox=f"{bbox[1]},{bbox[0]},{bbox[3]},{bbox[2]},EPSG:4326", srsname="EPSG:4326",
             output="application/json", maxfeatures=str(PAGE), startindex=str(start), key=k, domain=REF)
    url = "https://api.vworld.kr/req/wfs?" + urllib.parse.urlencode(q)
    for a in range(8):
        try:
            t = urllib.request.urlopen(urllib.request.Request(url, headers={"Referer": REF, "User-Agent": "Mozilla/5.0"}), timeout=90).read()
            if t[:1] == b"{":
                return json.loads(t)["features"]
        except Exception:
            pass
        time.sleep(1.5 + a)
    raise RuntimeError("WFS 실패 " + str(bbox))


def fetch():
    k = key()
    Z = zones()
    xs = [p[0] for z in Z for r in z["rings"] for p in r]
    ys = [p[1] for z in Z for r in z["rings"] for p in r]
    os.makedirs(TILES, exist_ok=True)
    x0, y0 = min(xs), min(ys)
    nx, ny = int((max(xs) - x0) / STEP) + 1, int((max(ys) - y0) / STEP) + 1
    for i in range(nx):
        for j in range(ny):
            p = os.path.join(TILES, f"{i}_{j}.json")
            if os.path.exists(p):
                continue
            bb = (x0 + i * STEP, y0 + j * STEP, x0 + (i + 1) * STEP, y0 + (j + 1) * STEP)
            # 타일이 생활권과 겹치지 않으면 건너뜀(모서리·중심 검사)
            pts = [(bb[0], bb[1]), (bb[2], bb[1]), (bb[0], bb[3]), (bb[2], bb[3]), ((bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2)]
            if not any(zone_of(x, y, Z) for x, y in pts) and not any(bb[0] <= x <= bb[2] and bb[1] <= y <= bb[3] for x, y in zip(xs, ys)):
                json.dump([], open(p, "w"))
                continue
            feats, s = [], 0
            while True:
                f = wfs(k, bb, s)
                feats += f
                if len(f) < PAGE:
                    break
                s += PAGE
            json.dump(feats, open(p, "w", encoding="utf-8"), ensure_ascii=False)
            print(f"{i}_{j}", len(feats), flush=True)


def build():
    Z = zones()
    seen, out = set(), []
    for fn in os.listdir(TILES):
        for f in json.load(open(os.path.join(TILES, fn), encoding="utf-8")):
            pr = f["properties"]
            uid = f.get("id") or pr.get("ufid")   # ufid 는 "NN" 이 수천 건이라 못 쓴다
            if uid in seen:
                continue
            seen.add(uid)
            g = f["geometry"]
            polys = g["coordinates"] if g["type"] == "MultiPolygon" else [g["coordinates"]]
            ring = polys[0][0]
            cx = sum(p[0] for p in ring) / len(ring)
            cy = sum(p[1] for p in ring) / len(ring)
            z = zone_of(cx, cy, Z)
            if not z:
                continue
            h = float(pr.get("height") or 0)
            fl = int(pr.get("grnd_flr") or 0)
            if h <= 0 or h > 250:
                h = fl * 3.3 if fl > 0 else 4.0
            polys = [[[[round(x, 6), round(y, 6)] for x, y in r] for r in poly] for poly in polys]
            out.append({"type": "Feature", "geometry": {"type": "MultiPolygon", "coordinates": polys},
                        "properties": {"h": round(h, 1), "fl": fl, "z": z, "nm": pr.get("bld_nm") or "",
                                       "use": pr.get("usability") or "", "apr": pr.get("useapr_day") or "",
                                       "pnu": pr.get("pnu") or "", "ta": round(float(pr.get("totalarea") or 0))}})
    json.dump({"type": "FeatureCollection", "features": out},
              open(os.path.join(HERE, "data", "footprints.json"), "w", encoding="utf-8"), ensure_ascii=False,
              separators=(",", ":"))
    hs = sum(1 for f in out if f["properties"]["h"] != 4.0)
    print("건물", len(out), "높이·층수 있음", hs)


if __name__ == "__main__":
    if "--build" not in sys.argv:
        fetch()
    build()
