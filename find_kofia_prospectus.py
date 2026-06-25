"""
금융투자협회(dis.kofia.or.kr) 투자설명서 메뉴 자동 탐색
- 왼쪽 메뉴를 자동으로 클릭하면서 투자설명서 섹션 찾기
- 발견 시 API 패턴 캡처
"""
import asyncio
import json
from pathlib import Path
from playwright.async_api import async_playwright

KOFIA_URL = "https://dis.kofia.or.kr/websquare/index.jsp"
OUT_DIR = Path("output/prospectus")
OUT_DIR.mkdir(parents=True, exist_ok=True)

captured_api = []


async def main():
    print("=" * 60)
    print("금융투자협회 투자설명서 탐색")
    print("=" * 60)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            accept_downloads=True,
            viewport={"width": 1400, "height": 900},
        )
        page = await context.new_page()

        # 모든 POST 요청 캡처
        async def on_request(req):
            if "kofia.or.kr" in req.url and req.method == "POST":
                body = req.post_data or ""
                captured_api.append({"url": req.url, "body": body[:1000]})
                print(f"  [API] POST {req.url}")
                if body: print(f"        {body[:200]}")

        async def on_response(resp):
            url = resp.url
            ct = resp.headers.get("content-type", "")
            if "kofia.or.kr" in url and ("pdf" in ct or "octet" in ct):
                print(f"  [PDF!] {url}")
                captured_api.append({"url": url, "type": "pdf", "content_type": ct})

        page.on("request", on_request)
        page.on("response", on_response)

        # ── 1. KOFIA DIS 메인 로드 ──
        print("\n[1] KOFIA DIS 로딩...")
        await page.goto(KOFIA_URL, timeout=30000)
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)
        print("  ✅ 로드 완료")

        # ── 2. 메뉴 구조 전체 파악 ──
        print("\n[2] 메뉴 구조 파악...")
        menu_items = await page.evaluate('''() => {
            const items = [];
            // 모든 텍스트 노드 + 클릭 가능 요소
            const els = document.querySelectorAll(
                "li, a, span, div.menu, div.gnb, div.lnb, td[onclick], [role='menuitem']"
            );
            els.forEach(el => {
                const text = el.textContent?.trim().replace(/\\s+/g, " ");
                if (text && text.length >= 2 && text.length <= 30) {
                    const onclick = el.getAttribute("onclick") || "";
                    const href = el.getAttribute("href") || "";
                    const id = el.id || "";
                    const cls = el.className || "";
                    if (onclick || href || id.toLowerCase().includes("menu")) {
                        items.push({tag: el.tagName, text, onclick: onclick.substring(0,200), href, id, class: cls.substring(0,50)});
                    }
                }
            });
            return items;
        }''')

        print(f"  메뉴 요소 {len(menu_items)}개 발견:")
        for m in menu_items:
            print(f"    [{m['tag']}] '{m['text']}' | id={m['id']} | onclick={m['onclick'][:80]}")

        # ── 3. '투자설명서' 텍스트 클릭 시도 ──
        print("\n[3] '투자설명서' 메뉴 클릭 시도...")

        keywords = ["투자설명서", "간이투자", "핵심상품", "설명서"]
        clicked = False

        for kw in keywords:
            try:
                # 텍스트로 요소 찾기
                el = page.locator(f"text={kw}").first
                count = await el.count()
                if count > 0:
                    print(f"  '{kw}' 요소 발견! 클릭 시도...")
                    await el.click(timeout=5000)
                    await asyncio.sleep(2)
                    print(f"  ✅ '{kw}' 클릭 성공")
                    clicked = True

                    # 현재 URL 확인
                    current_url = page.url
                    print(f"  현재 URL: {current_url}")

                    # 페이지 내 검색창 찾기
                    inputs = await page.evaluate('''() => {
                        return Array.from(document.querySelectorAll("input[type='text'], input[type='search']"))
                            .map(el => ({id: el.id, name: el.name, placeholder: el.placeholder, value: el.value}));
                    }''')
                    print(f"  입력창: {inputs}")
                    break
            except Exception as e:
                print(f"  '{kw}' 클릭 실패: {e}")

        if not clicked:
            print("  ⚠️  자동 클릭 실패")

        # ── 4. 현재 페이지에서 w2xPath 파라미터 탐색 ──
        print("\n[4] w2xPath 탐색...")
        paths = await page.evaluate('''() => {
            const src = document.documentElement.innerHTML;
            const found = new Set();
            const r = /w2xPath[='"]+([^'"&\\s]+)/g;
            let m;
            while ((m = r.exec(src)) !== null) found.add(m[1]);
            // iframe src도 확인
            document.querySelectorAll("iframe").forEach(f => {
                const s = f.src || "";
                if (s) found.add("iframe: " + s.substring(0,200));
            });
            return [...found];
        }''')
        print(f"  발견된 w2xPath:")
        for p_val in paths:
            print(f"    {p_val}")

        # ── 5. 페이지 스크린샷 ──
        ss_path = OUT_DIR / "kofia_current.png"
        await page.screenshot(path=str(ss_path), full_page=False)
        print(f"\n  스크린샷 저장: {ss_path}")

        # ── 6. iframe 내부 탐색 ──
        print("\n[5] iframe 탐색...")
        frames = page.frames
        print(f"  프레임 수: {len(frames)}")
        for i, frame in enumerate(frames):
            try:
                frame_url = frame.url
                frame_text = await frame.evaluate("() => document.body?.innerText?.substring(0, 100)")
                print(f"  [Frame {i}] url={frame_url} | text={frame_text}")

                # 투자설명서 관련 링크
                links = await frame.evaluate('''() => {
                    return Array.from(document.querySelectorAll("a, li, [onclick]"))
                        .filter(el => el.textContent?.includes("설명서") || el.textContent?.includes("공시"))
                        .map(el => ({
                            text: el.textContent?.trim().substring(0,40),
                            href: el.getAttribute("href") || "",
                            onclick: (el.getAttribute("onclick") || "").substring(0,200),
                        }));
                }''')
                if links:
                    print(f"    관련 링크 {len(links)}개:")
                    for l in links:
                        print(f"      {l['text']} | {l['href']} | {l['onclick']}")
            except Exception as e:
                print(f"  [Frame {i}] 오류: {e}")

        # 결과 저장
        result_file = OUT_DIR / "kofia_explore.json"
        result_file.write_text(json.dumps({
            "menu_items": menu_items,
            "w2x_paths": paths,
            "captured_api": captured_api,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n결과 저장: {result_file}")

        print("\n10초 후 종료...")
        await asyncio.sleep(10)
        await browser.close()


asyncio.run(main())
