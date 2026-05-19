"""Playwright 직접 로그인 후 ETF 페이지 데이터 추출 + 내부 API 엔드포인트 탐지"""
import asyncio
import json
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )

        # 내부 API 응답 수집
        api_data = []
        async def on_response(response):
            url = response.url
            ct = response.headers.get("content-type","")
            if "funetf.co.kr" in url and "json" in ct:
                try:
                    body = await response.json()
                    api_data.append({"url": url, "data": body})
                except:
                    pass
        page.on("response", on_response)

        # 로그인
        await page.goto("https://www.funetf.co.kr/auth")
        await page.wait_for_load_state("domcontentloaded")
        await page.fill('input[name="username"]', "sungsoowon45@gmail.com")
        await page.fill('input[name="password"]', "!sungsoo0405")
        await page.locator('button:has-text("로그인")').last.click()
        await page.wait_for_url("https://www.funetf.co.kr/", timeout=15000)
        print("✅ 로그인 성공")

        # ── ETF 페이지 테스트 (KODEX 미국S&P500) ──
        api_data.clear()
        await page.goto("https://www.funetf.co.kr/product/etf/view/379800")
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(3)

        print(f"\n📡 ETF 페이지 내부 JSON API 호출 {len(api_data)}개:")
        for item in api_data:
            print(f"  URL: {item['url']}")
            print(f"  데이터: {json.dumps(item['data'], ensure_ascii=False)[:300]}\n")

        # 렌더링된 DOM에서 데이터 추출
        body_text = await page.inner_text("body")
        # 핵심 정보 라인들만
        lines = [l.strip() for l in body_text.split('\n') if l.strip()]
        print(f"\n📄 ETF 페이지 렌더링된 텍스트 (앞 50줄):")
        for line in lines[:50]:
            print(f"  {line}")

        # ── 펀드 페이지 테스트 (삼성미국S&P500) ──
        api_data.clear()
        await page.goto("https://www.funetf.co.kr/search?schVal=삼성미국S%26P500인덱스증권자투자신탁")
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(3)

        print(f"\n📡 펀드 검색 JSON API 호출 {len(api_data)}개:")
        for item in api_data:
            print(f"  URL: {item['url']}")
            print(f"  데이터: {json.dumps(item['data'], ensure_ascii=False)[:300]}\n")

        await browser.close()
        print("\n✅ 완료!")

asyncio.run(main())
