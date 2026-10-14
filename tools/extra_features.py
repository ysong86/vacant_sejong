# -*- coding: utf-8 -*-
"""멀티레벨 분석용 추가 변수 — 이미 받아 둔 자료로만 → data/buildings_plus.json

  python tools/extra_features.py

건물(수준 1)
  - 건축물대장 표제부(조사 건물 필지): 지상층수·지하층수·주차대수·승강기   ← 필지 PNU 는 브이월드 건물 외곽선에서
  - 소상공인 상가정보(입점 점포): 업종 구성·다양도·프랜차이즈·층 구성    ← 도로명 주소로 건물에 붙인다
    (점포 '개수'는 공실의 거울이라 쓰지 않고 '구성 비율'만 쓴다)
반경(수요·공급)
  - 반경 500m 공동주택 세대 수(건축물대장 세대수, 필지 좌표화), 반경 500m 조사 상가면적 → 세대당 상가면적
  - 반경 1km 아파트 실거래 ㎡당 가격(최근 24개월, 세대 가중) — 배후 구매력 대리 변수
행정동(수준 3)
  - 생활권 상황판 latest.json: 인구 증감·고령·유소년·직주비 등
"""
from __future__ import annotations

import collections
import datetime as dt
import json
import math
import os
import re
import statistics
import sys
import urllib.parse

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ZONE = os.path.join(os.path.dirname(HERE), "sejong_livingzone")
sys.path.insert(0, ZONE)
import collect as ZC  # noqa: E402

RAW = os.path.join(HERE, "data", "raw")
ZRAW = os.path.join(ZONE, "data", "raw")
TITLE_URL = ("https://apis.data.go.kr/1613000/BldRgstHubService/getBrTitleInfo?serviceKey={k}&sigunguCd={sg}"
             "&bjdongCd={bj}&platGbCd={pg}&bun={bun}&ji={ji}&numOfRows=100&pageNo=1&_type=json")
LCLS = ["음식", "소매", "수리·개인", "교육", "과학·기술", "부동산", "예술·스포츠", "시설관리·임대", "보건의료", "숙박"]


def jload(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def jsave(p, o):
    with open(p, "w", encoding="utf-8") as f:
        json.dump(o, f, ensure_ascii=False, indent=1)


def inside(x, y, ring):
    c, j = False, len(ring) - 1
    for i in range(len(ring)):
        xi, yi = ring[i]
        xj, yj = ring[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi) + xi:
            c = not c
        j = i
    return c


def dist(a, b):
    return math.hypot((a[0] - b[0]) * 111320 * math.cos(math.radians(36.5)), (a[1] - b[1]) * 110940)


# ── 1. 조사 건물 → 필지 PNU(브이월드 건물 외곽선) ────────────────────────
def pnu_of(B):
    tiles = os.path.join(RAW, "bldg_tiles")
    feats, seen = [], set()
    for fn in os.listdir(tiles):
        for f in jload(os.path.join(tiles, fn)):
            if f.get("id") in seen:
                continue
            seen.add(f.get("id"))
            g = f["geometry"]
            polys = g["coordinates"] if g["type"] == "MultiPolygon" else [g["coordinates"]]
            r = polys[0][0]
            feats.append((polys, (sum(p[0] for p in r) / len(r), sum(p[1] for p in r) / len(r)), f["properties"]))
    grid = collections.defaultdict(list)
    for i, (_, c, _) in enumerate(feats):
        grid[(int(c[0] / .002), int(c[1] / .002))].append(i)
    out = {}
    for r in B:
        x, y = r["lon"], r["lat"]
        cand = [i for dx in (-1, 0, 1) for dy in (-1, 0, 1) for i in grid[(int(x / .002) + dx, int(y / .002) + dy)]]
        hit = [i for i in cand if any(inside(x, y, p[0]) for p in feats[i][0])]
        if not hit:
            d = sorted((dist((x, y), feats[i][1]), i) for i in cand)
            hit = [d[0][1]] if d and d[0][0] <= 40 else []
        if hit:
            p = feats[hit[0]][2]
            out[r["no"]] = {"pnu": p.get("pnu"), "grnd_flr_vw": p.get("grnd_flr") or None}
    return out


# ── 2. 건축물대장 표제부(필지 단위) ─────────────────────────────────────
def titles(pnus):
    cpath = os.path.join(RAW, "bld_title.json")
    cache = jload(cpath) if os.path.exists(cpath) else {}
    key = ZC._q(ZC.load_key()) if hasattr(ZC, "_q") else urllib.parse.quote(ZC.load_key(), safe="")
    for pnu in sorted(set(pnus)):
        if not pnu or pnu in cache or len(pnu) != 19:
            continue
        sg, bj, pg, bun, ji = pnu[:5], pnu[5:10], str(int(pnu[10]) - 1), pnu[11:15], pnu[15:19]
        try:
            d = ZC._json(TITLE_URL.format(k=key, sg=sg, bj=bj, pg=pg, bun=bun, ji=ji))
            b = d["response"]["body"]
            it = (b["items"].get("item") or []) if isinstance(b.get("items"), dict) else []
            it = [it] if isinstance(it, dict) else it
        except RuntimeError as e:
            print("표제부 실패", pnu, e)
            continue
        cache[pnu] = [{k: x.get(k) for k in ("bldNm", "dongNm", "mainAtchGbCd", "mainPurpsCdNm", "totArea", "grndFlrCnt",
                                               "ugrndFlrCnt", "rideUseElvtCnt", "emgenUseElvtCnt", "indrAutoUtcnt",
                                               "oudrAutoUtcnt", "indrMechUtcnt", "oudrMechUtcnt", "hhldCnt")} for x in it]
        jsave(cpath, cache)
    return cache


def title_feats(rows):
    """한 필지에 동이 여럿이면 층수는 최대, 주차·승강기는 합."""
    if not rows:
        return {}
    n = lambda v: float(v) if v not in (None, "") else 0.0
    park = sum(n(x["indrAutoUtcnt"]) + n(x["oudrAutoUtcnt"]) + n(x["indrMechUtcnt"]) + n(x["oudrMechUtcnt"]) for x in rows)
    return {"floors": max(n(x["grndFlrCnt"]) for x in rows) or None,
            "base_floors": max(n(x["ugrndFlrCnt"]) for x in rows),
            "parking": park, "elev": sum(n(x["rideUseElvtCnt"]) + n(x["emgenUseElvtCnt"]) for x in rows),
            "reg_area": sum(n(x["totArea"]) for x in rows)}


# ── 3. 입점 점포 구성(소상공인 상가정보) ─────────────────────────────────
def road_key(rd, mn, sl):
    rd = (rd or "").split()[-1] if rd else ""
    sl = str(sl or "").strip()
    return f"{rd} {int(mn)}" + (f"-{int(sl)}" if sl and sl != "0" else "")


def store_feats(stores):
    if not stores:
        return {"stores": 0}
    n = len(stores)
    c = collections.Counter(s["indsLclsNm"] for s in stores)
    ent = -sum(v / n * math.log(v / n) for v in c.values())
    fl = [s["flrNo"] for s in stores if str(s["flrNo"]).strip().lstrip("-").isdigit()]
    fl = [int(x) for x in fl]
    out = {"stores": n,
           "sh_food": c["음식"] / n * 100, "sh_retail": c["소매"] / n * 100, "sh_edu": c["교육"] / n * 100,
           "sh_med": c["보건의료"] / n * 100, "sh_estate": c["부동산"] / n * 100,
           "diversity": ent / math.log(len(LCLS)),                       # 0~1 정규화 엔트로피
           "franchise": sum(1 for s in stores if (s.get("brchNm") or "").strip()) / n * 100,
           "flr_known": len(fl)}
    if len(fl) >= 3:
        out["sh_1f"] = sum(1 for f in fl if f == 1) / len(fl) * 100
        out["sh_3up"] = sum(1 for f in fl if f >= 3) / len(fl) * 100
    return out


# ── 4. 배후 수요: 공동주택 세대·아파트 가격 ──────────────────────────────
def apt_lots():
    b = jload(os.path.join(ZRAW, "bld.json"))["items"]
    lots = collections.defaultdict(int)
    for x in b:
        if x["mainPurpsCdNm"] == "공동주택" and x.get("hhldCnt"):
            lots[x["platPlc"].replace("번지", "").strip()] += int(x["hhldCnt"])
    with open(os.path.join(ZONE, "config.json"), encoding="utf-8") as f:
        vkey = json.load(f)["vworld_key"]
    cpath = os.path.join(RAW, "geo_apt.json")
    cache = jload(cpath) if os.path.exists(cpath) else {}
    pts = []
    for i, (addr, hh) in enumerate(sorted(lots.items())):
        xy = ZC._geocode(addr, vkey, cache)
        if i % 50 == 0:
            jsave(cpath, cache)
        if xy:
            pts.append({"addr": addr, "hh": hh, "lon": xy[0], "lat": xy[1]})
    jsave(cpath, cache)
    # 아파트 실거래: 필지(동+지번)별 최근 24개월 ㎡당 가격 중앙값(만원)
    r = jload(os.path.join(ZRAW, "rtms.json"))
    items = r["items"] if isinstance(r, dict) else r
    cut = (dt.date.today().year - 2) * 100 + dt.date.today().month
    price = collections.defaultdict(list)
    for t in items:
        try:
            ym = int(t["dealYear"]) * 100 + int(t["dealMonth"])
            if ym < cut or str(t.get("cdealType") or "").strip():
                continue
            v = float(str(t["dealAmount"]).replace(",", "")) / float(t["excluUseAr"])
        except (ValueError, KeyError, ZeroDivisionError):
            continue
        price[f"세종특별자치시 {t['umdNm'].strip()} {str(t['jibun']).strip()}"].append(v)
    for p in pts:
        v = price.get(p["addr"])
        p["ppm"] = statistics.median(v) if v else None
    return pts


def main():
    B = jload(os.path.join(HERE, "data", "buildings.json"))
    B = [r for r in B if r.get("lon") is not None and r.get("shop_area")]
    # 1·2 건축물대장
    pn = pnu_of(B)
    T = titles([v["pnu"] for v in pn.values()])
    # 3 점포
    S = jload(os.path.join(ZRAW, "sbiz.json"))["items"]
    idx = collections.defaultdict(list)
    for s in S:
        try:
            idx[road_key(s["rdnm"], s["bldMnno"], s["bldSlno"])].append(s)
        except (ValueError, TypeError):
            pass
    # 4 배후 수요
    A = apt_lots()
    U = jload(os.path.join(ZONE, "data", "latest.json"))["units"]
    miss = collections.Counter()
    for r in B:
        p = pn.get(r["no"], {})
        r["pnu"] = p.get("pnu")
        tf = title_feats(T.get(r["pnu"]) or [])
        if not tf:
            miss["대장"] += 1
        r.update(tf)
        if not r.get("floors") and p.get("grnd_flr_vw"):
            r["floors"] = float(p["grnd_flr_vw"])
        r["park_per_100"] = r["parking"] / r["shop_area"] * 100 if tf else None     # 상가 100㎡당 주차
        st = store_feats(idx.get(r["addr"].replace(" ", " ").strip(), []))
        if not st["stores"]:
            miss["점포"] += 1
        r.update(st)
        pt = (r["lon"], r["lat"])
        near = [a for a in A if dist(pt, (a["lon"], a["lat"])) <= 500]
        r["hh500"] = sum(a["hh"] for a in near)
        r["supply500"] = sum(q["shop_area"] for q in B if dist(pt, (q["lon"], q["lat"])) <= 500)
        r["shop_per_hh"] = r["supply500"] / r["hh500"] if r["hh500"] else None   # 세대당 상가면적(㎡)
        pr = [(a["ppm"], a["hh"]) for a in A if a["ppm"] and dist(pt, (a["lon"], a["lat"])) <= 1000]
        r["apt_ppm"] = sum(v * w for v, w in pr) / sum(w for _, w in pr) if pr else None
        u = U.get(r["unit"], {})
        for k in ("chg", "old", "kid", "emp", "pop", "hh"):
            r["u_" + k] = u.get(k)
    jsave(os.path.join(HERE, "data", "buildings_plus.json"), B)
    print("건물", len(B), "누락", dict(miss), "공동주택 필지 좌표", len(A), "가격 있는 필지", sum(1 for a in A if a["ppm"]))
    for k in ("floors", "parking", "park_per_100", "stores", "sh_food", "sh_edu", "diversity", "franchise", "sh_1f", "sh_3up",
              "hh500", "shop_per_hh", "apt_ppm", "u_chg", "u_old", "u_emp"):
        v = [r[k] for r in B if r.get(k) is not None]
        print(f"  {k:12s} n={len(v):3d}  평균 {statistics.mean(v):9.2f}  중앙 {statistics.median(v):9.2f}" if v else f"  {k} 없음")


if __name__ == "__main__":
    main()
