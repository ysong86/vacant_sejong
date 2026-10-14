# -*- coding: utf-8 -*-
"""상가공실 상황판 기능 소개 영상(약 30초, MP4) — 로그인 문에서 '미리보기'로 공개한다.

  python tools/make_video.py
    입력: dashboard.html(평문, 로컬)   출력: docs/상가공실상황판_소개.mp4 + site/preview.mp4

공개 영상이라 공실률 숫자·건물명·분석 수치는 흐리게 처리한다(기능·화면 구성만 보이게).
녹화 방식은 생활권 상황판 tools/make_video.py 와 같다(설치된 Chrome, 가짜 포인터·클릭 물결·자막).
"""
from __future__ import annotations

import glob
import os
import shutil
import subprocess
import time

import imageio_ffmpeg
from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = "file:///" + os.path.join(HERE, "dashboard.html").replace("\\", "/")
TMP = os.path.join(HERE, "data", "video_tmp")
OUT = os.path.join(HERE, "docs", "상가공실상황판_소개.mp4")
SITE = os.path.join(HERE, "site", "preview.mp4")
W, H = 1440, 810
LENGTH = 31.0

# 수치·이름 흐리게 — 화면 구성과 색은 그대로 보이게
BLUR = r"""
(() => {
  const css = document.createElement('style');
  css.textContent = `
  .kpi .v,.kpi .s,#notice,.br .bv,#ols-note,#rf-note,.fr .fv,.dc .rate,.dc .meta,.dp .v,#detail,.tip,
  .maplibregl-popup-content,#method,.vblur,#v3 .note,#ml-sec .ml-step .v,#ml-sec .ml-stack span,#ml-sec .ml-leg,
  #ml-sec .ml-find,#ml-sec p,#ml-sec .sec-h .note{filter:blur(5px)!important}`;
  document.head.appendChild(css);
  const mark = () => document.querySelectorAll('.find h3,.find p,.dc .dh b ~ *,.sm .note').forEach(e => {
    if (/\d/.test(e.textContent)) e.classList.add('vblur'); });
  mark(); new MutationObserver(mark).observe(document.body, {childList: true, subtree: true});
})();
"""

OVERLAY = r"""
(() => {
  const css = document.createElement('style');
  css.textContent = `
  #vcur{position:fixed;left:0;top:0;width:26px;height:26px;z-index:2147483647;pointer-events:none;
        transform:translate(-3px,-2px);filter:drop-shadow(0 2px 3px rgba(0,0,0,.35))}
  .vring{position:fixed;width:46px;height:46px;margin:-23px 0 0 -23px;border-radius:50%;border:3px solid #1D5C7C;
         z-index:2147483646;pointer-events:none;animation:vr .6s ease-out forwards}
  @keyframes vr{from{transform:scale(.3);opacity:.9}to{transform:scale(1.4);opacity:0}}
  #vcap{position:fixed;left:50%;bottom:26px;transform:translateX(-50%);z-index:2147483645;pointer-events:none;
        background:rgba(21,35,42,.9);color:#fff;font:700 21px/1.3 "IBM Plex Sans KR","Malgun Gothic",sans-serif;
        padding:10px 22px;border-radius:999px;letter-spacing:-.01em;opacity:0;transition:opacity .35s;white-space:nowrap}
  #vcap small{font-weight:400;font-size:15px;opacity:.85;margin-left:10px}
  html{scroll-behavior:auto!important}`;
  document.head.appendChild(css);
  const c = document.createElement('div'); c.id = 'vcur';
  c.innerHTML = '<svg viewBox="0 0 24 24" width="26" height="26"><path d="M4 2l15 9.5-6.6 1.6 3.9 7.4-3.1 1.6-3.9-7.5L4 19.5z" fill="#fff" stroke="#111" stroke-width="1.4" stroke-linejoin="round"/></svg>';
  document.body.appendChild(c);
  const cap = document.createElement('div'); cap.id = 'vcap'; document.body.appendChild(cap);
  document.addEventListener('mousemove', e => { c.style.left = e.clientX + 'px'; c.style.top = e.clientY + 'px'; }, true);
  document.addEventListener('mousedown', e => {
    const r = document.createElement('div'); r.className = 'vring';
    r.style.left = e.clientX + 'px'; r.style.top = e.clientY + 'px';
    document.body.appendChild(r); setTimeout(() => r.remove(), 700);
  }, true);
  window.__cap = (t, s) => { cap.innerHTML = t + (s ? '<small>' + s + '</small>' : ''); cap.style.opacity = t ? 1 : 0; };
  window.__scroll = (y, ms) => new Promise(res => {
    const y0 = scrollY, t0 = performance.now(), ease = t => t < .5 ? 2*t*t : 1 - Math.pow(-2*t + 2, 2) / 2;
    (function step(now){ const k = Math.min(1, (now - t0) / ms); scrollTo(0, y0 + (y - y0) * ease(k));
      k < 1 ? requestAnimationFrame(step) : res(); })(t0);
  });
  window.__top = sel => { const el = document.querySelector(sel); return el ? el.getBoundingClientRect().top + scrollY : 0; };
})();
"""


def main():
    os.makedirs(TMP, exist_ok=True)
    for f in glob.glob(os.path.join(TMP, "*.webm")):
        os.remove(f)
    with sync_playwright() as p:
        b = p.chromium.launch(channel="chrome", headless=True, args=["--enable-unsafe-swiftshader", "--ignore-gpu-blocklist"])
        ctx = b.new_context(viewport={"width": W, "height": H}, device_scale_factor=1, color_scheme="light",
                            record_video_dir=TMP, record_video_size={"width": W, "height": H})
        ctx.route("**/hits.sh/**", lambda r: r.abort())          # 녹화가 실제 조회수를 올리지 않게
        ctx.add_init_script("try{localStorage.clear()}catch(e){}")
        ctx.add_init_script(f"document.addEventListener('DOMContentLoaded', () => {{ {BLUR} }});")
        t_open = time.time()
        pg = ctx.new_page()
        pg.goto(SRC, wait_until="networkidle")
        pg.evaluate(OVERLAY)
        pg.mouse.move(W * 0.55, H * 0.45)
        pg.wait_for_timeout(900)
        start = time.time() - t_open                       # 이 시점부터 영상으로 쓴다

        def cap(t, s=""):
            pg.evaluate("([t,s]) => __cap(t,s)", [t, s])

        def scroll_to(sel, ms, off=70):
            y = pg.evaluate(f"__top('{sel}')") - off
            pg.evaluate(f"__scroll({max(0, y)}, {ms})")

        def move(sel, steps=18, click=True, dx=0.5, dy=0.5):
            box = pg.locator(sel).first.bounding_box()
            x, y = box["x"] + box["width"] * dx, box["y"] + box["height"] * dy
            pg.mouse.move(x, y, steps=steps)
            if click:
                pg.mouse.down()
                pg.mouse.up()
            return x, y

        # 0~3초 — 제목
        cap("세종 행복도시 상가공실 상황판", "집합상가 전수조사 · 생활권·동·건물 단위")
        pg.wait_for_timeout(1300)
        move(".kpi:nth-child(1)", click=False)
        pg.wait_for_timeout(500)
        move(".kpi:nth-child(3)", steps=16, click=False)
        pg.wait_for_timeout(500)

        # 3~7초 — 생활권·동별
        cap("어디가 비어 있나", "생활권·동별 공실률 — 누르면 지도로 이동")
        scroll_to("#zbars", 900, 120)
        pg.wait_for_timeout(300)
        move("#zbars .br:nth-child(1)", click=False)
        pg.wait_for_timeout(500)
        move("#dbars .br:nth-child(2)", steps=20)
        pg.wait_for_timeout(700)

        # 7~13초 — 2D 지도
        cap("공실 지도(2D)", "건물별 공실률 · 조건 거르기 · 핫스팟")
        scroll_to("#map-sec", 900, 10)
        pg.wait_for_timeout(400)
        move('#rail [data-f="new"]', steps=16)
        pg.wait_for_timeout(900)
        move('#rail [data-l="hot"]', steps=14)
        pg.wait_for_timeout(900)
        move('#rail [data-f="all"]', steps=12)
        pg.wait_for_timeout(500)
        move("#map", steps=16, click=False, dx=.55, dy=.5)
        pg.mouse.wheel(0, -400)
        pg.wait_for_timeout(800)

        # 13~24초 — 3D 지도
        cap("3D 지도", "건물을 실제 높이로 — 조사 건물은 공실률 색")
        move('.v3-toggle button[data-v="3d"]', steps=16)
        pg.wait_for_function("!!document.querySelector('#v3-map canvas') && !document.getElementById('v3-msg')", timeout=30000)
        pg.wait_for_timeout(1200)
        # 행복도시 전체는 3D 로 보기엔 멀다 — 동 막대를 누른 것처럼 나성동으로 날아간다(3D 조각이 막대 클릭을 받는다)
        pg.evaluate("""() => { const y = scrollY; document.querySelector('#dbars .br[data-k="나성동"]').click(); scrollTo(0, y); }""")
        pg.wait_for_timeout(100)
        scroll_to("#map-sec", 300, 10)
        pg.wait_for_timeout(1700)
        cap("보는 방향 바꾸기", "돌리기 · 기울이기 · 한 바퀴 돌아보기")
        move(".v3-pad .rt", steps=14, click=False)
        pg.mouse.down()
        pg.wait_for_timeout(1500)
        pg.mouse.up()
        move(".v3-pad .up", steps=10)
        pg.wait_for_timeout(600)
        move('.v3-pad [data-p="spin"]', steps=12)
        pg.wait_for_timeout(2200)
        cap("공실률 강조", "공실이 클수록 건물이 높이 솟습니다")
        move('#v3-rail [data-g="hmode"][data-v="vac"]', steps=16)
        pg.wait_for_timeout(2300)

        # 24~31초 — 분석
        cap("왜 비는가", "회귀분석·머신러닝으로 본 공실 요인")
        pg.mouse.move(W * 0.92, H * 0.4, steps=8)
        scroll_to("#forest", 1000, 140)
        pg.wait_for_timeout(1200)
        cap("멀티레벨 분석", "건물·상권 블록·행정동 세 수준으로 나눠 본 공실")
        scroll_to("#ml-sec", 1000, 20)
        pg.wait_for_timeout(1900)
        cap("연구모임 내부 검토용", "비밀번호 문의: 송양호(ysong@sri.re.kr)")
        pg.wait_for_timeout(1500)

        end = time.time() - t_open
        video = pg.video.path()
        ctx.close()
        b.close()

    ff = imageio_ffmpeg.get_ffmpeg_exe()
    dur = min(LENGTH, end - start)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    cmd = [ff, "-y", "-ss", f"{start:.2f}", "-i", video, "-t", f"{dur:.2f}",
           "-vf", "fps=30,format=yuv420p", "-c:v", "libx264", "-preset", "slow", "-crf", "23",
           "-movflags", "+faststart", "-an", OUT]
    subprocess.run(cmd, check=True, capture_output=True)
    os.makedirs(os.path.dirname(SITE), exist_ok=True)
    shutil.copyfile(OUT, SITE)
    print(f"영상: {OUT}  ({dur:.1f}초, 녹화 {end - start:.1f}초, {os.path.getsize(OUT) / 1e6:.1f}MB)")


if __name__ == "__main__":
    main()
