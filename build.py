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
    out = [
        {"t": "새로 지은 상가가 채워지지 못한다",
         "b": f"준공 5년 이하 건물은 다른 조건이 같아도 공실률이 <b>{c['new5']['b']:+.1f}%p</b> 높습니다(p={c['new5']['p']:.3f}). "
              f"준공 3년 이하 {age['3년 이하']['rate']}% → 5~7년 {age['5~7년']['rate']}% → 9년 초과 {age['9년 초과']['rate']}%로, "
              "시간이 지나며 흡수되는 구조입니다. 준공 시기를 나눠 공급하고, 준공 초기 임대료 감면·용도 유연화로 흡수 기간을 줄이는 방안을 검토할 수 있습니다."},
        {"t": "제2종 근린생활시설 건물의 공실이 크다",
         "b": f"제2종 근린생활시설 건물은 공실률이 <b>{c['type2']['b']:+.1f}%p</b> 높습니다(p={c['type2']['p']:.3f}; 단순 비교 "
              f"{use.get('제2종 근린생활', {}).get('rate', '—')}% 대 제1종 {use.get('제1종 근린생활', {}).get('rate', '—')}%). "
              "상층부·대형 점포 위주의 공급이 수요와 맞지 않는 것으로 보이며, 공유업무·돌봄·교육·생활SOC 등으로의 용도 전환 허용(지구단위계획 지침 완화)이 대상이 됩니다."},
        {"t": "상가가 '곁다리'인 건물일수록 빈다",
         "b": f"연면적 중 상가 비중이 10%p 높을수록 공실률이 <b>{c['shop_share']['b']:+.1f}%p</b>(p={c['shop_share']['p']:.3f}). "
              "업무·주거가 주된 건물에 딸린 상가가 더 비어 있다는 뜻으로, 이런 건물의 부속 상가 총량을 줄이고 1층 가로변 위주로 계획하는 것이 바람직합니다."},
        {"t": "정류장·중심지 거리는 결정적이지 않다",
         "b": f"버스정류장({c['d_stop']['b']:+.1f}%p/100m), BRT({c['d_brt']['b']:+.1f}%p/100m), 나성동 중심({c['d_center']['b']:+.1f}%p/1km) 거리는 "
              f"방향은 예상대로지만 통계적으로 유의하지 않습니다{'' if not any(sig(k) for k in ('d_stop', 'd_brt', 'd_center')) else '(일부 유의)'}. "
              "조사 건물 대부분이 정류장 200m 안에 있어 접근성 차이가 작고, 공실은 입지보다 <b>공급 구조(연식·용도·총량)</b>에서 갈립니다."},
        {"t": f"동 간 격차: {hi['dong']} {hi['rate']}% · {lo['dong']} {lo['rate']}%",
         "b": f"{hi['dong']}은 평균보다 {dh['gap']:+.1f}%p 높으며, 이 가운데 연식·용도·입지로 {dh['explained']:+.1f}%p가 설명되고 "
              f"{dh['other']:+.1f}%p는 모형 밖 요인입니다. {lo['dong']}처럼 신축인데도 공실이 낮은 곳은 상가 총량이 작아 수요와 맞은 사례로, "
              "생활권별 상가 총량 관리의 근거가 됩니다."},
        {"t": f"공실이 몰린 곳(핫스팟) {a['hot']['hot']}동",
         "b": "반경 500m 안에서 공실이 높은 건물이 모여 있는 곳(Getis-Ord Gi*)을 지도 '핫스팟' 겹으로 표시했습니다. 개별 건물이 아니라 "
              "가로·블록 단위로 비어 있는 곳이라, 집중 지원·임시 활용(팝업·공공 임대) 대상 구역을 고르는 데 쓸 수 있습니다."},
    ]
    return out


def method(a):
    ys = a["survey_years"]
    return (f"<p><b>자료</b> — 세종연구원 행정중심복합도시 집합상가 전수조사({'·'.join(ys)}년). 건물 {a['n_buildings']}동, "
            f"{a['units']:,}호, 상가 면적 {a['shop_area']:,}㎡. 건물마다 상가 면적·공실 면적·호수·공실 호수·주용도·사용승인일.</p>"
            "<p><b>공실률</b> — 면적 기준(공실 면적 ÷ 상가 면적)을 기본으로, 호수 기준을 함께 표시. 생활권·동 공실률은 건물 합계로 계산(단순 평균 아님).</p>"
            "<p><b>여건 변수</b> — 건물 주소를 브이월드로 좌표화(287동 모두 성공)한 뒤, 생활권 상황판의 공공데이터로 버스·BRT 정류장 거리(TAGO), "
            "나성동 중심상업지까지 거리, 주변 200m 점포 수(소상공인 상가정보), 행정동 인구(주민등록), 행정동 대중교통 분담률(KTDB 2024)을 붙였습니다.</p>"
            f"<p><b>분석</b> — ① 상가 면적 가중 최소제곱 회귀(HC1 강건 표준오차, R² {a['ols']['r2']}) ② 랜덤포레스트 순열 중요도로 교차 확인"
            f"(OOB R² {a['rf']['oob_r2']}) ③ 동별 분해: 계수 × (동 평균 − 전체 평균) ④ Getis-Ord Gi* 핫스팟(반경 500m).</p>"
            "<p><b>한계</b> — 조사 시점이 섞여 있습니다(대부분 2023년, 일부 2025년). 건물 단위 자료라 층별(1층·상층) 차이는 보지 못했습니다. "
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
    out = os.path.join(HERE, "dashboard.html")
    open(out, "w", encoding="utf-8").write(head() + "\n" + body)
    print("생성:", out, f"{os.path.getsize(out) / 1024:.0f} KB")


if __name__ == "__main__":
    main()
