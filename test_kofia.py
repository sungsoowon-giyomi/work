"""
dis.kofia.or.kr에서 투자설명서 검색 및 다운로드 테스트
"""
import asyncio
import json
import requests
from pathlib import Path
from playwright.async_api import async_playwright

OUTPUT_DIR = Path(r"C:\Users\jswon\Desktop\업무\funetf_scraper\output")
PDF_DIR = OUTPUT_DIR / "pdfs"
PDF_DIR.mkdir(exist_ok=True)

TEST_FUNDS = [
    "미래에셋미국나스닥100인덱스증권자투자신탁",
    "NH-Amundi하나로단기채",
    "신한골드증권투자신탁",
]

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        # ── 메인 페이지 로드 ──
        print("dis.kofia.or.kr 접속 중...")
        api_calls = []
        async def on_resp(response):
            url = response.url
            ct = response.headers.get("content-type","")
            if "kofia" in url and ("json" in ct or "pdf" in ct or "xml" in ct):
                api_calls.append({"url": url[:100], "status": response.status, "ct": ct[:40]})

        page.on("response", on_resp)
        await page.goto("https://dis.kofia.or.kr/")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        print(f"페이지 로드 완료. API 호출 {len(api_calls)}개")

        # ── 체크박스 레이블 확인 ──
        labels = await page.evaluate('''() => {
            const results = [];
            document.querySelectorAll("input[type=checkbox]").forEach(el => {
                const label = document.querySelector(`label[for="${el.id}"]`);
                results.push({id: el.id, name: el.name, label: label ? label.innerText.trim() : ""});
            });
            return results;
        }''')
        print(f"\n체크박스 {len(labels)}개:")
        for lb in labels[:20]:
            print(f"  {lb['id']}: [{lb['label']}]")

        # ── 투자설명서 관련 버튼/탭 확인 ──
        buttons = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll("button, a.tab, .tab-btn, [class*=tab]")).map(el => ({
                tag: el.tagName, text: (el.innerText||"").trim().slice(0,30), id: el.id||""
            })).filter(e => e.text).slice(0,30);
        }''')
        print(f"\n버튼/탭:")
        for b in buttons:
            print(f"  {b['tag']} [{b['text']}] id={b['id']}")

        # ── 투자설명서 검색 시도 ──
        for fund_name in TEST_FUNDS[:2]:
            print(f"\n{'='*50}")
            print(f"검색: {fund_name}")
            api_calls.clear()

            try:
                # 1. 펀드명 입력
                await page.fill('#dSearchFndNm', fund_name)
                await asyncio.sleep(0.5)

                # 2. 검색 버튼 클릭
                # 검색 버튼 찾기
                search_btn = await page.evaluate('''() => {
                    const btns = Array.from(document.querySelectorAll("button, a, input[type=button], input[type=submit]"));
                    return btns.map(b => ({text: (b.innerText||b.value||"").trim(), id: b.id, class: b.className}))
                              .filter(b => b.text.includes("검색") || b.text.includes("조회")).slice(0,5);
                }''')
                print(f"  검색 버튼: {search_btn}")

                # 검색 버튼 클릭
                try:
                    await page.click('button:has-text("검색")', timeout=3000)
                except:
                    try:
                        await page.click('[id*="search"], [id*="Search"], [id*="조회"], [id*="srch"]', timeout=3000)
                    except:
                        # Enter 키로 검색
                        await page.press('#dSearchFndNm', 'Enter')

                await asyncio.sleep(3)
                await page.wait_for_load_state("networkidle", timeout=15000)
                await asyncio.sleep(2)

                # 결과 확인
                result_text = await page.inner_text("body")
                print(f"  결과 (앞 500자): {result_text[:500]}")

                # API 응답
                print(f"\n  API 호출 {len(api_calls)}개:")
                for c in api_calls[:5]:
                    print(f"    {c['status']} {c['url']}")

                # 결과 목록에서 PDF 링크 찾기
                pdf_links = await page.evaluate('''() => {
                    const links = [];
                    document.querySelectorAll("a, button, [onclick]").forEach(el => {
                        const href = el.getAttribute("href") || el.getAttribute("onclick") || "";
                        const text = (el.innerText||"").trim();
                        if (href.includes(".pdf") || href.includes("down") || text.includes("다운")) {
                            links.push({href: href.slice(0,100), text: text.slice(0,30)});
                        }
                    });
                    return links.slice(0,10);
                }''')
                print(f"\n  PDF 링크: {pdf_links}")

            except Exception as e:
                print(f"  오류: {e}")

        # ── 투자설명서 탭/메뉴 찾기 (메인 화면에서) ──
        print(f"\n{'='*50}")
        print("투자설명서 메뉴 탐색")

        # 메인 화면으로 돌아가서 메뉴 구조 확인
        await page.goto("https://dis.kofia.or.kr/")
        await page.wait_for_load_state("networkidle", timeout=20000)
        await asyncio.sleep(3)

        # 페이지의 주요 메뉴/링크
        nav_links = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll("nav a, .nav a, .menu a, .gnb a, li a")).map(el => ({
                text: (el.innerText||"").trim().slice(0,40),
                href: el.getAttribute("href")||"",
                onclick: (el.getAttribute("onclick")||"").slice(0,80)
            })).filter(e => e.text).slice(0,30);
        }''')
        print(f"네비게이션 링크 {len(nav_links)}개:")
        for n in nav_links:
            print(f"  [{n['text']}] href={n['href']} onclick={n['onclick'][:50]}")

        # 투자설명서 메뉴 찾아 클릭
        for menu_text in ["투자설명서", "공시", "서류"]:
            try:
                el = page.locator(f"text={menu_text}").first
                if await el.count() > 0:
                    await el.click(timeout=2000)
                    await asyncio.sleep(2)
                    url_after = page.url
                    print(f"\n  [{menu_text}] 클릭 후 URL: {url_after}")
                    break
            except:
                pass

        # 현재 페이지 URL 및 API 호출
        print(f"\n현재 URL: {page.url}")
        api_summary = [c for c in api_calls if "json" in c.get("ct","")][:5]
        print(f"JSON API 호출: {api_summary}")

        await browser.close()

asyncio.run(main())
