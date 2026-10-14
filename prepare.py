# -*- coding: utf-8 -*-
"""상가공실 전수조사(엑셀) → 건물별 좌표·여건 변수 → data/buildings.json

  python prepare.py

- 좌표: 브이월드 주소→좌표(도로명 → 지번). data/raw/geo_cache.json 에 남긴다.
- 여건 변수는 생활권 상황판(sejong_livingzone)이 받아 둔 원자료를 그대로 쓴다.
"""
from __future__ import annotations

import datetime as dt
import json
import math
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ZONE = os.path.join(os.path.dirname(HERE), "sejong_livingzone")
sys.path.insert(0, ZONE)
import collect as ZC          # noqa: E402  (생활권 상황판의 수집·기하 도구)
import livingzone as LZ       # noqa: E402

RAW = os.path.join(HERE, "data", "raw")
ZRAW = os.path.join(ZONE, "data", "raw")

# 조사 동 → 생활권 (어진은 1, 가람 없음)
DONG_ZONE = {"고운동": "1", "아름동": "1", "종촌동": "1", "도담동": "1", "어진동": "1",
             "다정동": "2", "새롬동": "2", "한솔동": "2", "나성동": "2",
             "대평동": "3", "보람동": "3", "소담동": "3", "반곡동": "4", "집현동": "4",
             "합강동": "5", "해밀동": "6", "산울동": "6"}
# 지도 단위(행정동 면) — 도담·어진은 한 면
DONG_UNIT = {"도담동": "도담·어진동", "어진동": "도담·어진동"}
NASUNG = (127.2600, 36.4870)          # 나성동 중심상업(어반아트리움 일대)


def load_survey():
    wb = openpyxl.load_workbook(os.path.join(RAW, "survey.xlsx"), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    out = []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or r[0] is None:
            continue
        num = lambda v: float(v) if isinstance(v, (int, float)) else (float(str(v).replace(",", "")) if v not in (None, "") else None)
        apr = str(r[16] or "").strip()[:8]
        out.append({
            "no": int(r[0]), "name": (r[1] or "").strip(), "dong": (r[2] or "").strip(), "sub": (r[5] or "").strip(),
            "addr": (r[6] or "").strip(), "use": (r[7] or "").replace("\xa0", " ").strip(),
            "gfa": num(r[8]), "shop_area": num(r[9]), "vac_area": num(r[10]),
            "units": int(r[12] or 0), "vac_units": int(r[13] or 0),
            "survey": str(r[15] or "").strip()[:4], "apr": apr, "age_class": (r[17] or "").strip(),
        })
    return out


def geocode_all(rows):
    key = ZC._cfg("vworld_key") if hasattr(ZC, "_cfg") else ""
    if not key:
        with open(os.path.join(ZONE, "config.json"), encoding="utf-8") as f:
            key = json.load(f)["vworld_key"]
    cpath = os.path.join(RAW, "geo_cache.json")
    cache = json.load(open(cpath, encoding="utf-8")) if os.path.exists(cpath) else {}
    fail = []
    for r in rows:
        addr = "세종특별자치시 " + r["addr"]
        xy = ZC._geocode(addr, key, cache)
        if not xy:
            xy = ZC._geocode("세종특별자치시 " + r["dong"] + " " + r["name"], key, cache)
        r["lon"], r["lat"] = (xy or (None, None))
        if not xy:
            fail.append(r["name"])
    json.dump(cache, open(cpath, "w", encoding="utf-8"), ensure_ascii=False)
    return fail


def dist_m(a, b):
    ax, ay = ZC._xy(*a)
    bx, by = ZC._xy(*b)
    return math.hypot(ax - bx, ay - by)


def features(rows):
    raw = lambda n: json.load(open(os.path.join(ZRAW, n + ".json"), encoding="utf-8"))
    stops = [(float(s["gpslong"]), float(s["gpslati"]), s) for s in raw("tago")["items"]
             if s.get("gpslong") and s.get("gpslati")]
    roads = json.load(open(os.path.join(ZONE, "assets", "sejong_roads.json"), encoding="utf-8"))["roads"]
    brt_xy = [[ZC._xy(*p) for p in l] for r in roads if r["cls"] == "brt" for l in r["lines"]]

    def on_brt(lon, lat):
        x, y = ZC._xy(lon, lat)
        return any(ZC._dist_to_seg(x, y, *l[i], *l[i + 1]) <= 40 for l in brt_xy for i in range(len(l) - 1))
    brt = [p for p in stops if "BRT" in (p[2].get("nodenm") or "").upper() or on_brt(p[0], p[1])]
    stores = [(float(s["lon"]), float(s["lat"]), s) for s in raw("sbiz")["items"] if s.get("lon")]
    schools = [(s["lon"], s["lat"], s) for s in raw("neis")["items"] if s.get("lon") and s["kind"] == "초등학교"]
    g_stop, g_brt, g_store, g_sch = ZC.Grid(stops), ZC.Grid(brt, 1000.0), ZC.Grid(stores, 300.0), ZC.Grid(schools, 1000.0)
    latest = json.load(open(os.path.join(ZONE, "data", "latest.json"), encoding="utf-8"))
    U = latest["units"]
    today = dt.date.today()
    for r in rows:
        unit = DONG_UNIT.get(r["dong"], r["dong"])
        r["zone"] = DONG_ZONE.get(r["dong"])
        r["unit"] = unit
        r["vac_rate_area"] = round(r["vac_area"] / r["shop_area"] * 100, 1) if r["shop_area"] else None
        r["vac_rate_units"] = round(r["vac_units"] / r["units"] * 100, 1) if r["units"] else None
        try:
            d = dt.date(int(r["apr"][:4]), int(r["apr"][4:6]), int(r["apr"][6:8]))
            ref = dt.date(int(r["survey"]), 6, 30) if r["survey"].isdigit() else today
            r["age"] = round((ref - d).days / 365.25, 1)
        except (ValueError, IndexError):
            r["age"] = None
        r["shop_share"] = round(r["shop_area"] / r["gfa"] * 100, 1) if r["gfa"] and r["shop_area"] else None
        r["mixed_res"] = "공동주택" in r["use"] or "도시형생활주택" in r["use"]
        if r["lon"] is None:
            continue
        lon, lat = r["lon"], r["lat"]
        r["d_stop"] = round(g_stop.nearest(lon, lat, 3000) or 3000)
        r["d_brt"] = round(g_brt.nearest(lon, lat, 8000) or 8000)
        r["d_center"] = round(dist_m((lon, lat), NASUNG))
        r["stores_200"] = sum(1 for _ in g_store.near(lon, lat, 200))
        r["d_school"] = round(g_sch.nearest(lon, lat, 3000) or 3000)
        u = U.get(unit, {})
        for k in ("pop", "hh", "kid", "emp", "dens", "inter", "transit"):
            r["u_" + k] = u.get(k)
    return rows


def main():
    rows = load_survey()
    fail = geocode_all(rows)
    print("건물", len(rows), "좌표 실패", len(fail), fail[:10])
    rows = features(rows)
    os.makedirs(os.path.join(HERE, "data"), exist_ok=True)
    json.dump(rows, open(os.path.join(HERE, "data", "buildings.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("저장: data/buildings.json")


if __name__ == "__main__":
    main()
