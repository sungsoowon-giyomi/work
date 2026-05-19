"""
FunETF 내부 API 호출 감지 테스트
- Playwright로 ETF/펀드 페이지 로드 시 어떤 데이터 요청이 발생하는지 확인
"""
import asyncio
import json
from playwright.async_api import async_playwright

EMAIL = "sungsoowon45@gmail.com"
PASSWORD = "!sungsoo0405"

async def intercept_test():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        # 모든 네트워크 요청 수집
        api_calls = []

        async def on_request(request):
            url = request.url
            # funetf 내부 호출만 수집 (광고/분석 제외)
            if "funetf.co.kr" in url and not any(x in url for x in ["static", "favicon", "google", "facebook"]):
                api_calls.append({
                    "method": request.method,
                    "url": url,
                    "post_data": request.post_data
                })

        page.on("request", on_request)

        # 1. 로그인
        print("=== 로그인 중 ===")
        await page.goto("https://www.funetf.co.kr/auth")
        await page.fill('input[name="username"]', EMAIL)
        await page.fill('input[name="password"]', PASSWORD)
        # 로그인 버튼 클릭 (form action="/auth/login")
        await page.locator('form[action="/auth/login"] button').first.click()
        await page.wait_for_url("https://www.funetf.co.kr/", timeout=15000)
        print("로그인 성공!")

        # 2. ETF 페이지 테스트 (KODEX 미국S&P500)
        print("\n=== ETF 페이지 로드 중 (379800) ===")
        api_calls.clear()
        await page.goto("https://www.funetf.co.kr/product/etf/view/379800")
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(2)

        print(f"ETF 페이지 API 호출 ({len(api_calls)}개):")
        for call in api_calls:
            print(f"  [{call['method']}] {call['url']}")
            if call['post_data']:
                print(f"    BODY: {call['post_data'][:200]}")

        # 3. 펀드 페이지 테스트
        print("\n=== 펀드 검색 테스트 ===")
        api_calls.clear()
        await page.goto("https://www.funetf.co.kr/search?schVal=삼성미국S%26P500인덱스")
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(2)

        print(f"검색 페이지 API 호출 ({len(api_calls)}개):")
        for call in api_calls:
            print(f"  [{call['method']}] {call['url']}")

        await browser.close()

asyncio.run(intercept_test())
