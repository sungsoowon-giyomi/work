"""
하이브리드 방식 테스트:
1. requests로 로그인 → 쿠키 획득
2. 그 쿠키를 Playwright에 주입 → JS 렌더링으로 데이터 추출
3. 내부 API 호출 경로 확인
"""
import asyncio
import json
import requests
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

EMAIL = "sungsoowon45@gmail.com"
PASSWORD = "!sungsoo0405"

def get_session_cookies():
    """requests로 로그인하고 쿠키 반환"""
    s = requests.Session()
    s.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'})
    r = s.get('https://www.funetf.co.kr/auth')
    csrf = BeautifulSoup(r.text, 'html.parser').find('input', {'name': '_csrf'})['value']
    s.post('https://www.funetf.co.kr/auth/login', data={
        '_csrf': csrf, 'username': EMAIL, 'password': PASSWORD
    })
    cookies = [{'name': c.name, 'value': c.value, 'domain': '.funetf.co.kr', 'path': '/'}
               for c in s.cookies]
    print(f"쿠키 {len(cookies)}개 획득: {[c['name'] for c in cookies]}")
    return cookies

async def test_playwright(cookies):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        # 쿠키 주입
        await context.add_cookies(cookies)

        page = await context.new_page()

        # 내부 API 호출 수집
        api_calls = []
        async def on_response(response):
            url = response.url
            if "funetf.co.kr" in url and not any(x in url for x in ["static", "favicon", ".css", ".ico", ".png", ".jpg"]):
                try:
                    body = await response.body()
                    api_calls.append({
                        "url": url,
                        "status": response.status,
                        "content_type": response.headers.get("content-type",""),
                        "body_preview": body[:300].decode('utf-8', errors='replace')
                    })
                except:
                    pass

        page.on("response", on_response)

        # ETF 페이지 로드
        print("\n=== ETF 페이지 로드 (KODEX 미국S&P500) ===")
        await page.goto("https://www.funetf.co.kr/product/etf/view/379800")
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(3)

        print(f"\n내부 응답 {len(api_calls)}개:")
        for call in api_calls:
            print(f"\n  URL: {call['url']}")
            print(f"  Status: {call['status']}, Type: {call['content_type'][:50]}")
            print(f"  Body: {call['body_preview']}")

        # 페이지에서 텍스트 추출 (JS 렌더링 후)
        print("\n=== 페이지 렌더링 후 텍스트 ===")
        title = await page.title()
        print(f"Title: {title}")

        # 핵심 데이터 추출 시도
        for selector, label in [
            ('.etf-title, h1, .product-title', '제목'),
            ('.total-fee, .fee, [class*="fee"]', '보수'),
            ('[class*="date"], [class*="Date"]', '날짜'),
        ]:
            try:
                el = page.locator(selector).first
                text = await el.inner_text(timeout=2000)
                print(f"{label}: {text[:100]}")
            except:
                pass

        await browser.close()

cookies = get_session_cookies()
asyncio.run(test_playwright(cookies))
