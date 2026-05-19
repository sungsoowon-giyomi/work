"""
FunETF 내부 API 엔드포인트 완전 탐색
- 검색 API로 fundCd 찾기
- 올바른 상세 페이지 URL 찾기
- 상세 페이지의 모든 JSON API 호출 캡쳐
"""
import asyncio
import json
import requests
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

EMAIL = "sungsoowon45@gmail.com"
PASSWORD = "!sungsoo0405"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )

        # 로그인
        await page.goto("https://www.funetf.co.kr/auth")
        await page.wait_for_load_state("domcontentloaded")
        await page.fill('input[name="username"]', EMAIL)
        await page.fill('input[name="password"]', PASSWORD)
        await page.locator('button:has-text("로그인")').last.click()
        await page.wait_for_url("https://www.funetf.co.kr/", timeout=15000)
        print("✅ 로그인 성공")

        # ── 1. 검색 API로 KODEX S&P500 fundCd 확인 ──
        search_url = "https://www.funetf.co.kr/api/public/main/search/all?schVal=KODEX+미국S%26P500&reSchVal=&schKeyword="
        await page.goto(search_url)
        await asyncio.sleep(1)
        search_data = json.loads(await page.inner_text("body"))
        etf_list = search_data.get("etfList", {}).get("content", [])
        if etf_list:
            etf = etf_list[0]
            print(f"\nETF 검색 결과:")
            print(json.dumps(etf, ensure_ascii=False, indent=2))
            fund_cd = etf.get("fundCd","")
            item_id = etf.get("itemId","")
            print(f"\nfundCd: {fund_cd}, itemId: {item_id}")

        # ── 2. ETF 상세 페이지 URL 후보들 테스트 ──
        api_calls = []
        async def on_response(response):
            url = response.url
            ct = response.headers.get("content-type","")
            if "funetf.co.kr/api" in url and "json" in ct:
                try:
                    body = await response.json()
                    api_calls.append({"url": url, "data": body})
                except:
                    pass
        page.on("response", on_response)

        # 후보 URL들 테스트
        test_urls = [
            f"https://www.funetf.co.kr/product/etf/view/{fund_cd}",
            f"https://www.funetf.co.kr/product/etf/view/379800",
        ]

        for test_url in test_urls:
            api_calls.clear()
            await page.goto(test_url)
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(3)
            body_text = await page.inner_text("body")
            success = "서비스 이용에 불편" not in body_text
            print(f"\n{'✅' if success else '❌'} URL: {test_url}")
            if success:
                lines = [l.strip() for l in body_text.split('\n') if l.strip()]
                print(f"  텍스트 (앞 20줄): {lines[:20]}")
                print(f"\n  📡 API 호출 {len(api_calls)}개:")
                for c in api_calls:
                    print(f"    {c['url']}")
                    print(f"    {json.dumps(c['data'], ensure_ascii=False)[:200]}")
                break

        # ── 3. 펀드 상세 페이지 API 탐색 ──
        search_url2 = "https://www.funetf.co.kr/api/public/main/search/all?schVal=삼성미국S%26P500인덱스증권자투자신탁UH&reSchVal=&schKeyword="
        await page.goto(search_url2)
        await asyncio.sleep(1)
        search_data2 = json.loads(await page.inner_text("body"))
        fund_list = search_data2.get("fundList", {}).get("content", [])
        if fund_list:
            fund = fund_list[0]
            fund_cd2 = fund.get("fundCd","")
            print(f"\n펀드 fundCd: {fund_cd2}")

            api_calls.clear()
            await page.goto(f"https://www.funetf.co.kr/product/fund/view/{fund_cd2}")
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(3)
            body_text2 = await page.inner_text("body")
            success2 = "서비스 이용에 불편" not in body_text2
            print(f"{'✅' if success2 else '❌'} 펀드 상세 페이지")
            if success2:
                print(f"\n  📡 펀드 API 호출 {len(api_calls)}개:")
                for c in api_calls:
                    print(f"    {c['url']}")
                    print(f"    {json.dumps(c['data'], ensure_ascii=False)[:300]}\n")

        await browser.close()

asyncio.run(main())
