# -*- coding: utf-8 -*-
"""세종 행복도시 상가공실 상황판 — dashboard.html(평문) 생성.

  python prepare.py     (엑셀 → 건물 좌표·여건 변수, 처음 한 번 또는 자료가 바뀔 때)
  python analyze.py     (원인 분석)
  python build.py       (화면)
  python ../sejong_livingzone/tools/encrypt_page.py dashboard.html site/index.html "세종 행복도시 상가공실 상황판"   (공개용 암호화)

화면 머리(글꼴·색·지도 스타일)는 생활권 상황판 template.html 의 <style> 을 그대로 가져온다 — '화면은 똑같이'.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ZONE = os.path.join(os.path.dirname(HERE), "sejong_livingzone")
sys.path.insert(0, ZONE)
import livingzone as LZ  # noqa: E402


def load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def head():
    t = open(os.path.join(ZONE, "template.html"), encoding="utf-8").read()
    h = t[: t.index("</style>") + len("</style>")]
    h = re.sub(r"<title>.*?</title>", "<title>세종 상가공실 상황판</title>", h)
    h = re.sub(r'<meta name="description"[^>]*>',
               '<meta name="description" content="행정중심복합도시 집합상가 전수조사로 본 생활권·동·건물별 상가 공실과 원인 분석">', h)
    return h


def finds(a):
    c = {x["key"]: x for x in a["ols"]["coefs"]}
    age = {r["label"]: r for r in a["bins"]["age"]}
    use = {r["label"]: r for r in a["bins"]["use"]}
    dong = {d["dong"]: d for d in a["by_dong"]}
    hi = max(a["by_dong"], key=lambda d: d["rate"])
    lo = min(a["by_dong"], key=lambda d: d["rate"])
    dh = a["decomp"][hi["dong"]]
    sig = lambda k: c[k]["p"] < 0.05
    pv = lambda k: f"p={c[k]['p']:.3f}"
    out = [
        {"t": "연식이 쌓일수록 공실이 빠르게 준다",
         "b": f"다른 조건이 같아도 준공 후 1년마다 공실률이 <b>{c['age']['b']:+.1f}%p</b>({pv('age')}) 달라집니다. "
              f"준공 3년 이하 {age['3년 이하']['rate']}% → 5~7년 {age['5~7년']['rate']}% → 9년 초과 {age['9년 초과']['rate']}%로, "
              "새 상가는 몇 해에 걸쳐 흡수되는 구조입니다(2025년 12월 조사 기준 경과연수). 준공 시기를 나눠 공급하고, 준공 초기 임대료 감면·용도 유연화로 흡수 기간을 줄이는 방안을 검토할 수 있습니다."},
        {"t": "제2종 근린생활시설 건물의 공실이 크다",
         "b": f"제2종 근린생활시설 건물은 공실률이 <b>{c['type2']['b']:+.1f}%p</b> 높습니다({pv('type2')}; 단순 비교 "
              f"{use.get('제2종 근린생활', {}).get('rate', '—')}% 대 제1종 {use.get('제1종 근린생활', {}).get('rate', '—')}%). "
              "상층부·대형 점포 위주의 공급이 수요와 맞지 않는 것으로 보이며, 공유업무·돌봄·교육·생활SOC 등으로의 용도 전환 허용(지구단위계획 지침 완화)이 대상이 됩니다."},
        {"t": "상가가 '곁다리'인 건물일수록 빈다",
         "b": f"연면적 중 상가 비중이 10%p 높을수록 공실률이 <b>{c['shop_share']['b']:+.1f}%p</b>({pv('shop_share')}). "
              f"주상복합(공동주택 하부 상가)은 <b>{c['mixed_res']['b']:+.1f}%p</b>({pv('mixed_res')})로 오히려 덜 빕니다 — 위층 세대가 곧 배후 수요입니다. "
              "업무시설 등에 딸린 부속 상가는 총량을 줄이고 1층 가로변 위주로 계획하는 것이 바람직합니다."},
        {"t": "정류장 거리보다 상권 밀도·배후 인구",
         "b": f"버스정류장({c['d_stop']['b']:+.1f}%p/100m)·BRT({c['d_brt']['b']:+.1f}%p/100m)·나성동 중심({c['d_center']['b']:+.1f}%p/1km) 거리는 "
              f"통계적으로 결정적이지 않습니다{'' if not any(sig(k) for k in ('d_stop', 'd_brt', 'd_center')) else '(일부 유의)'}. "
              f"반면 주변 200m 점포가 2배면 <b>{c['log_stores']['b']:+.1f}%p</b>({pv('log_stores')}), 행정동 인구 1만 명당 <b>{c['u_pop_k']['b']:+.1f}%p</b>({pv('u_pop_k')}) — "
              "상권이 모여 있고 사람이 많은 곳일수록 덜 빕니다."},
        {"t": f"동 간 격차: {hi['dong']} {hi['rate']}% · {lo['dong']} {lo['rate']}%",
         "b": f"{hi['dong']}은 평균보다 {dh['gap']:+.1f}%p 높으며, 이 가운데 연식·용도·입지로 {dh['explained']:+.1f}%p가 설명되고 "
              f"{dh['other']:+.1f}%p는 모형 밖 요인입니다. {lo['dong']}처럼 공실이 낮은 곳은 상가 총량이 작아 수요와 맞은 사례로, "
              "생활권별 상가 총량 관리의 근거가 됩니다."},
        {"t": f"공실이 몰린 곳(핫스팟) {a['hot']['hot']}동",
         "b": "반경 500m 안에서 공실이 높은 건물이 모여 있는 곳(Getis-Ord Gi*)을 지도 '핫스팟' 겹으로 표시했습니다. 개별 건물이 아니라 "
              "가로·블록 단위로 비어 있는 곳이라, 집중 지원·임시 활용(팝업·공공 임대) 대상 구역을 고르는 데 쓸 수 있습니다."},
    ]
    return out


def method(a):
    ys = a["survey_years"]
    return (f"<p><b>자료</b> — 세종연구원 행정중심복합도시 상가 전수조사. 2023년 6월·2025년 12월 두 차례, 집합상가와 아파트 단지 내 상가. "
            f"원인 분석은 2025년 12월 집합상가 {a['n_buildings']}동, "
            f"{a['units']:,}호, 상가 면적 {a['shop_area']:,}㎡. 건물마다 상가 면적·공실 면적·호수·공실 호수·주용도·사용승인일.</p>"
            "<p><b>공실률</b> — 면적 기준(공실 면적 ÷ 상가 면적)을 기본으로, 호수 기준을 함께 표시. 생활권·동 공실률은 건물 합계로 계산(단순 평균 아님).</p>"
            "<p><b>여건 변수</b> — 건물 주소를 브이월드로 좌표화(287동 모두 성공)한 뒤, 생활권 상황판의 공공데이터로 버스·BRT 정류장 거리(TAGO), "
            "나성동 중심상업지까지 거리, 주변 200m 점포 수(소상공인 상가정보), 행정동 인구(주민등록), 행정동 대중교통 분담률(KTDB 2024)을 붙였습니다.</p>"
            f"<p><b>분석</b> — ① 상가 면적 가중 최소제곱 회귀(HC1 강건 표준오차, R² {a['ols']['r2']}) ② 랜덤포레스트 순열 중요도로 교차 확인"
            f"(OOB R² {a['rf']['oob_r2']}) ③ 동별 분해: 계수 × (동 평균 − 전체 평균) ④ Getis-Ord Gi* 핫스팟(반경 500m).</p>"
            "<p><b>아파트 단지 내 상가</b> — 주소가 단지 대표 위치라 2D 지도에서는 단지 주소 자리에 마름모로 표시합니다. 3D 지도에서는 주소가 든 필지(PNU)의 "
            "건축물대장 근린생활시설 동 수만큼, 같은 필지의 낮은 건물(15m 이하) 가운데 바닥이 큰 것을 상가동으로 칠하고, 찾지 못한 단지는 작은 기둥으로 표시합니다.</p>"
            "<p><b>한계</b> — 건물 단위 자료라 층별(1층·상층) 차이는 보지 못했습니다. 2023년과 2025년은 조사 대상 건물 수가 달라(새로 지은 건물 편입) 전체 공실률 차이에는 대상 변화가 섞여 있습니다. "
            "회귀 결과는 상관이지 인과가 아니며, 건물이 적은 동(해밀동 2동 등)의 분해 값은 불안정합니다. 정류장 거리는 직선거리입니다.</p>")


def main():
    a = load(os.path.join(HERE, "data", "analysis.json"))
    latest = load(os.path.join(ZONE, "data", "latest.json"))
    reb = ((latest.get("vacancy") or {}).get("집합") or {}).get("세종")
    roads = [{"cls": r["cls"], "lines": [[[round(x, 5), round(y, 5)] for x, y in l] for l in r["lines"]]}
             for r in load(os.path.join(HERE, "assets", "sejong_roads.json"))["roads"] if r["cls"] in ("expressway", "brt", "arterial")]
    data = {
        "built": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "a": a, "reb": reb[1] if reb else None,
        "zones": LZ.ZONES,
        "shapes": load(os.path.join(HERE, "assets", "sejong_units.json"))["units"],
        "zoneShapes": load(os.path.join(HERE, "assets", "sejong_zones.json"))["zones"],
        "roads": roads,
        "finds": finds(a), "method": method(a),
        # 조사 시점·대상별 기록(2023.06·2025.12 × 집합상가·아파트 상가) — tools/prepare_all.py
        "surv": [{k: r.get(k) for k in ("id", "yr", "kind", "code", "name", "sub", "dong", "zone", "unit", "addr", "use", "area", "vac",
                                         "units", "vu", "rate", "ru", "age", "lon", "lat", "gi")}
                 for r in load(os.path.join(HERE, "data", "survey_all.json"))],
    }
    js = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    body = open(os.path.join(HERE, "body.html"), encoding="utf-8").read().replace("/*__DATA__*/null", js)
    # 3D 지도(2D/3D 전환) — make_map.py 가 만든 data/map3d.json 이 있을 때만 끼운다
    m3 = os.path.join(HERE, "data", "map3d.json")
    if os.path.exists(m3):
        part = open(os.path.join(HERE, "map3d_part.html"), encoding="utf-8").read().replace(
            "/*__DATA3D__*/null", open(m3, encoding="utf-8").read().replace("</", "<\\/"))
        anchor = '<script type="application/json" id="data">'
        body = body.replace(anchor, part + "\n" + anchor, 1)
    # 멀티레벨 분석 구역 — multilevel.py 가 만든 data/multilevel.json 이 있을 때만
    ml = os.path.join(HERE, "data", "multilevel.json")
    if os.path.exists(ml):
        part = open(os.path.join(HERE, "multilevel_part.html"), encoding="utf-8").read().replace(
            "/*__DATA_ML__*/null", open(ml, encoding="utf-8").read().replace("</", "<\\/"))
        anchor = '<script type="application/json" id="data">'
        body = body.replace(anchor, part + "\n" + anchor, 1)
    out = os.path.join(HERE, "dashboard.html")
    open(out, "w", encoding="utf-8").write(head() + "\n" + body)
    print("생성:", out, f"{os.path.getsize(out) / 1024:.0f} KB")


if __name__ == "__main__":
    main()
