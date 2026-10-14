# 세종 행복도시 상가공실 상황판

행정중심복합도시 집합상가 전수조사(세종연구원)로 본 생활권·동·건물별 상가 공실과 원인 분석.
공실 지도는 2D(건물별 원)와 3D(건물을 실제 높이로 세우고 조사 건물만 공실률 색)를 버튼으로 바꿔 봅니다.

공개 화면: https://ysong86.github.io/vacant_sejong/ (연구모임 내부 검토용, 비밀번호)

## 만들기

```
python tools/fetch_buildings.py   # 행복도시 건물 외곽선·높이(브이월드 GIS건물통합정보), 처음 한 번
python prepare.py                 # 조사 엑셀 → 건물 좌표·여건 변수
python analyze.py                 # 원인 분석(회귀·랜덤포레스트·핫스팟)
python make_map.py                # 3D 지도 자료(data/map3d.json)
python build.py                   # dashboard.html
python ../sejong_livingzone/tools/encrypt_page.py dashboard.html site/index.html "세종 행복도시 상가공실 상황판"
```

조사 원자료(`data/`)와 평문 화면(`dashboard.html`, `map3d.html`)은 저장소에 올리지 않습니다.
`gh-pages` 브랜치에는 암호화한 `index.html` 과 `.nojekyll` 만 둡니다.

제작: 세종연구원 연구모임(김성표, 안용준, 김흥주, 남영식, 이재민, 송양호, 이자은)
