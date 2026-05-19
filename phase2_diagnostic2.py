"""
Phase 2 진단2 스크립트:
- 펀드 항목(비삼성)의 FunETF 페이지에서 실제 R2 URL 확인
- ETF 펀드 페이지 실패 내용 확인
- fundsonline.co.kr 대안 URL 테스트
"""
import asyncio
import json
import re
import requests
import urllib.parse
from pathlib import Path
from playwright.async_api import async_playwright

EMAIL = "sungsoowon45@gmail.com"
PASSWORD = "!sungsoo0405"
OUTPUT_DIR = Path(r"C:\Users\jswon\Desktop\업무\funetf_scraper\output")
PDF_DIR = OUTPUT_DIR / "pdfs"

def try_download_head(url, session):
    try:
        r = session.head(url, timeout=10, allow_redirects=True)
        size = r.headers.get('content-length', '?')
        ct = r.headers.get('content-type', '')
        return r.status_code, size, ct
    except Exception as e:
        return -1, 0, str(e)[:50]

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        # 로그인
        await page.goto("https://www.funetf.co.kr/auth")
        await page.wait_for_load_state("domcontentloaded")
        await page.fill('input[name="username"]', EMAIL)
        await page.fill('input[name="password"]', PASSWORD)
        await page.locator('button:has-text("로그인")').last.click()
        await page.wait_for_url("https://www.funetf.co.kr/", timeout=15000)
        print("✅ 로그인 성공")

        cookies = await context.cookies()
        sess = requests.Session()
        sess.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'})
        for c in cookies:
            sess.cookies.set(c['name'], c['value'])

        # ── 1. ETF 펀드 페이지 실패 원인 확인 (TIME 글로벌AI) ──
        print("\n=== ETF 펀드 페이지 실패 내용 ===")
        time_fund_cd = "K553J1E23876"
        api_calls_etf = {}

        async def on_resp_etf(response):
            url = response.url
            ct = response.headers.get("content-type", "")
            if "funetf.co.kr" in url and "json" in ct:
                try:
                    body = await response.json()
                    api_calls_etf[url] = body
                except:
                    pass

        page.on("response", on_resp_etf)
        await page.goto(f"https://www.funetf.co.kr/product/fund/view/{time_fund_cd}")
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)
        body_text = await page.inner_text("body")
        page.remove_listener("response", on_resp_etf)
        print(f"  Body (앞 300자): {body_text[:300]}")
        print(f"  API 호출: {list(api_calls_etf.keys())[:5]}")

        # ── 2. 비삼성 펀드 페이지 테스트 (미래에셋 나스닥100) ──
        print("\n=== 비삼성 펀드 페이지 테스트 (미래에셋 나스닥100) ===")
        # 먼저 검색
        encoded = urllib.parse.quote("미래에셋미국나스닥100인덱스증권자투자신탁")
        await page.goto(f"https://www.funetf.co.kr/api/public/main/search/all?schVal={encoded}&reSchVal=&schKeyword=")
        await asyncio.sleep(0.5)
        data = json.loads(await page.inner_text("body"))
        fund_list = data.get("fundList", {}).get("content", [])

        fund_cd = None
        for f in fund_list:
            if "PAM078" in str(f.get("shortCd","")).upper():
                fund_cd = f.get("fundCd")
                print(f"  매칭: fundCd={fund_cd}, repFId={f.get('repFId')}")
                print(f"  shortCd={f.get('shortCd')}, fundFnm={f.get('fundFnm')}")
                break
        if not fund_cd and fund_list:
            f = fund_list[0]
            fund_cd = f.get("fundCd")
            print(f"  첫번째 결과: fundCd={fund_cd}, shortCd={f.get('shortCd')}, fundFnm={f.get('fundFnm')}")
            print(f"  repFId={f.get('repFId')}, fid={f.get('fid')}")

        if fund_cd:
            api_calls2 = {}
            async def on_resp2(response):
                url = response.url
                ct = response.headers.get("content-type", "")
                if "funetf.co.kr" in url and "json" in ct:
                    try:
                        body = await response.json()
                        api_calls2[url] = body
                    except:
                        pass

            page.on("response", on_resp2)
            await page.goto(f"https://www.funetf.co.kr/product/fund/view/{fund_cd}")
            await page.wait_for_load_state("networkidle", timeout=20000)
            await asyncio.sleep(3)
            page.remove_listener("response", on_resp2)

            body_text2 = await page.inner_text("body")
            page_ok = "서비스 이용에 불편" not in body_text2
            print(f"  페이지 로드: {'✅' if page_ok else '❌'}")
            if not page_ok:
                print(f"  Body: {body_text2[:200]}")

            # ALL data-fileurl 수집
            dom_docs = await page.evaluate('''() => {
                const results = [];
                document.querySelectorAll("[data-fileurl]").forEach(el => {
                    results.push({attr: "data-fileurl", val: el.getAttribute("data-fileurl"), tag: el.tagName, text: (el.innerText || "").trim()});
                });
                document.querySelectorAll("a[href*='.pdf'], a[href*='/upload/'], button[onclick*='upload']").forEach(el => {
                    results.push({attr: el.getAttribute("href") ? "href" : "onclick",
                                  val: el.getAttribute("href") || el.getAttribute("onclick"),
                                  tag: el.tagName, text: (el.innerText || "").trim()});
                });
                // 모든 버튼/링크 텍스트 포함
                document.querySelectorAll("button, .file-btn, .doc-btn, [class*='invest'], [class*='prosp']").forEach(el => {
                    const txt = (el.innerText || "").trim();
                    if (txt && txt.length < 30) results.push({attr: "button", val: "", tag: el.tagName, text: txt});
                });
                return results;
            }''')

            print(f"\n  DOM 문서 링크 {len(dom_docs)}개:")
            for doc in dom_docs[:20]:
                print(f"    {doc.get('attr')}: {(doc.get('val') or '')[:80]} [{doc.get('text','')}]")

            # API 응답 중 관련된 것
            print(f"\n  API 응답 {len(api_calls2)}개:")
            for url, body in api_calls2.items():
                if "funetf.co.kr/api" in url:
                    ep = url.split("/api/")[-1].split("?")[0]
                    print(f"    /api/{ep}")
                    print(f"    {json.dumps(body, ensure_ascii=False)[:400]}")

        # ── 3. 비삼성 펀드 - 다른 fund 코드 시도 (NH-Amundi 하나로단기채) ──
        print("\n=== NH-Amundi 하나로단기채 테스트 ===")
        encoded2 = urllib.parse.quote("NH-Amundi 하나로단기채증권투자신탁")
        await page.goto(f"https://www.funetf.co.kr/api/public/main/search/all?schVal={encoded2}&reSchVal=&schKeyword=")
        await asyncio.sleep(0.5)
        data2 = json.loads(await page.inner_text("body"))
        fund_list2 = data2.get("fundList", {}).get("content", [])
        print(f"  검색 결과 {len(fund_list2)}개:")
        for f in fund_list2[:5]:
            print(f"    shortCd={f.get('shortCd')}, fundCd={f.get('fundCd')}, repFId={f.get('repFId')}, fundFnm={f.get('fundFnm','')[:40]}")

        # PDNC74 코드 찾기
        nh_fund_cd = None
        for f in fund_list2:
            if "PDNC74" in str(f.get("shortCd","")).upper():
                nh_fund_cd = f.get("fundCd")
                print(f"  매칭: fundCd={nh_fund_cd}")
                break

        if nh_fund_cd:
            api_calls3 = {}
            async def on_resp3(response):
                url = response.url
                ct = response.headers.get("content-type", "")
                if "funetf.co.kr/api" in url and "json" in ct:
                    try:
                        body = await response.json()
                        api_calls3[url] = body
                    except:
                        pass

            page.on("response", on_resp3)
            await page.goto(f"https://www.funetf.co.kr/product/fund/view/{nh_fund_cd}")
            await page.wait_for_load_state("networkidle", timeout=20000)
            await asyncio.sleep(3)
            page.remove_listener("response", on_resp3)

            page_ok3 = "서비스 이용에 불편" not in await page.inner_text("body")
            print(f"  페이지 로드: {'✅' if page_ok3 else '❌'}")

            dom_docs3 = await page.evaluate('''() => {
                const results = [];
                document.querySelectorAll("[data-fileurl]").forEach(el => {
                    results.push({val: el.getAttribute("data-fileurl"), text: (el.innerText || "").trim()});
                });
                return results;
            }''')
            print(f"  data-fileurl {len(dom_docs3)}개:")
            for doc in dom_docs3:
                val = doc.get("val","")
                print(f"    {val}")
                # 여러 도메인에서 다운로드 테스트
                if val.startswith("/"):
                    for base in ["https://www.funetf.co.kr", "https://www.fundsonline.co.kr", "https://dis.kofia.or.kr"]:
                        full_url = base + val
                        status, size, ct = try_download_head(full_url, sess)
                        print(f"      {base}: {status} {size}B {ct[:25]}")

            # API 응답
            for url, body in api_calls3.items():
                if "funetf.co.kr/api" in url:
                    ep = url.split("/api/")[-1].split("?")[0]
                    print(f"  /api/{ep}")
                    print(f"  {json.dumps(body, ensure_ascii=False)[:400]}")

        await browser.close()

asyncio.run(main())
