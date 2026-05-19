"""
dis.kofia.or.kr WebSquare 서비스 호출 인터셉트:
- 펀드별 보수비용비교에서 실제 POST 요청 캡쳐
- 투자설명서 다운로드 URL 패턴 탐색
- Playwright locator 방식으로 셀 클릭 시도
"""
import asyncio
import json
import re
from pathlib import Path
from playwright.async_api import async_playwright

OUTPUT_DIR = Path(r"C:\Users\jswon\Desktop\업무\funetf_scraper\output")
PDF_DIR = OUTPUT_DIR / "pdfs"
PDF_DIR.mkdir(exist_ok=True)


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        # 모든 요청(POST 포함) 인터셉트
        captured_requests = []
        async def on_request(request):
            if "kofia" in request.url:
                post_data = request.post_data or ""
                captured_requests.append({
                    "method": request.method,
                    "url": request.url[:100],
                    "data": post_data[:300]
                })

        captured_responses = []
        async def on_response(response):
            url = response.url
            ct = response.headers.get("content-type","")
            st = response.status
            if "kofia" in url and st == 200:
                try:
                    body = await response.body()
                    if len(body) > 100 and (b"<" in body[:5] or b"{" in body[:5]):
                        captured_responses.append({
                            "url": url[:100], "ct": ct[:30], "size": len(body),
                            "body": body[:800].decode('utf-8', errors='replace')
                        })
                except: pass

        page.on("request", on_request)
        page.on("response", on_response)

        # ════════════════════════════════════════
        # A. DISFundFeeCMS - 검색 후 요청 캡쳐
        # ════════════════════════════════════════
        print("=" * 60)
        print("A. DISFundFeeCMS 검색 - 모든 요청/응답 캡쳐")
        print("=" * 60)

        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        # 미래에셋 선택
        try:
            await page.select_option('#compCd_input_0', label='미래에셋자산운용')
        except:
            try:
                options = await page.evaluate('''() => Array.from(document.querySelector("#compCd_input_0").options).map(o => ({v:o.value, t:o.text})).filter(o => o.t.includes("미래에셋"))''')
                if options:
                    await page.select_option('#compCd_input_0', options[0]['v'])
                    print(f"  회사 선택: {options[0]['t']}")
            except: pass

        await page.fill('#fundNm', '미래에셋미국나스닥100인덱스')
        await asyncio.sleep(0.5)

        # 조회 버튼 클릭 (Playwright locator)
        captured_requests.clear()
        captured_responses.clear()

        try:
            btn = page.locator('button, input[type="button"]').filter(has_text=re.compile(r'조회|검색|확인'))
            cnt = await btn.count()
            if cnt > 0:
                await btn.first.click()
            else:
                await page.keyboard.press('Enter')
        except:
            await page.keyboard.press('Enter')

        await asyncio.sleep(4)
        await page.wait_for_load_state("networkidle", timeout=15000)

        # POST 요청 출력
        post_reqs = [r for r in captured_requests if r["method"] == "POST"]
        print(f"\n  POST 요청 {len(post_reqs)}개:")
        for r in post_reqs:
            print(f"    URL: {r['url']}")
            print(f"    Data: {r['data'][:200]}")

        # 데이터 응답 출력
        data_resps = [r for r in captured_responses if "json" in r.get("ct","") or "xml" in r.get("ct","")]
        print(f"\n  데이터 응답 {len(data_resps)}개:")
        for r in data_resps[:5]:
            print(f"    {r['url']}: {r['body'][:300]}")

        # 페이지 결과
        body = await page.inner_text("body")
        print(f"\n  결과 (앞 800자):\n{body[:800]}")

        # ════════════════════════════════════════
        # B. DISFundMgrFundWebSrch에서 DP747 Playwright 클릭
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("B. DISFundMgrFundWebSrch DP747 Playwright 클릭")
        print("=" * 60)

        captured_requests.clear()
        captured_responses.clear()

        # 현재 DISFundMgrFundWebSrch 페이지로 이동 (검색 결과)
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundAnnSrch.xml&divisionId=MDIS01001000000000&serviceId=SDIS01001000000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(2)
        await page.fill('#dSearchFndNm', '미래에셋미국나스닥100인덱스')
        await page.keyboard.press('Enter')
        await asyncio.sleep(3)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)

        # DP747 텍스트를 포함한 행을 Playwright locator로 찾아 클릭
        print("  DP747 셀 Playwright 클릭...")
        captured_requests.clear()
        captured_responses.clear()

        try:
            dp747_cell = page.locator('text="DP747"').first
            if await dp747_cell.count() > 0:
                await dp747_cell.click()
                await asyncio.sleep(3)
                await page.wait_for_load_state("networkidle", timeout=10000)
                print(f"  클릭 성공! URL: {page.url}")
                body2 = await page.inner_text("body")
                print(f"  결과 (앞 1000자):\n{body2[:1000]}")
            else:
                print("  DP747 텍스트 없음")
        except Exception as e:
            print(f"  클릭 오류: {e}")

        # POST 요청
        print(f"\n  POST 요청 {len([r for r in captured_requests if r['method']=='POST'])}개:")
        for r in [r for r in captured_requests if r['method']=='POST']:
            print(f"    {r['url']}: {r['data'][:200]}")

        # ════════════════════════════════════════
        # C. WebSquare 서비스 직접 POST 호출 테스트
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("C. WebSquare 서비스 직접 POST (인터셉트한 요청 재현)")
        print("=" * 60)

        # DISFundFeeCMS 다시 로드 후 조회
        captured_requests.clear()
        captured_responses.clear()

        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(2)

        # 운용사 선택 (Playwright)
        try:
            select_el = page.locator('#compCd_input_0')
            options = await select_el.evaluate('''el => Array.from(el.options).map(o => ({v:o.value, t:o.text}))''')
            miraeasset = next((o for o in options if '미래에셋자산' in o['t']), None)
            if miraeasset:
                await select_el.select_option(miraeasset['v'])
                print(f"  운용사 선택: {miraeasset['t']}")
        except Exception as e:
            print(f"  운용사 선택 오류: {e}")

        # 펀드명 입력
        await page.fill('#fundNm', '미래에셋미국나스닥100인덱스')
        await asyncio.sleep(0.5)

        # Enter로 조회
        await page.keyboard.press('Enter')
        await asyncio.sleep(4)
        await page.wait_for_load_state("networkidle", timeout=15000)

        # 모든 캡쳐된 요청 출력 (POST 포함)
        print(f"\n  캡쳐된 요청 {len(captured_requests)}개:")
        for r in captured_requests:
            if r['method'] in ['POST', 'GET'] and '/service' in r['url'] or 'fund' in r['url'].lower():
                print(f"    [{r['method']}] {r['url']}")
                if r['data']:
                    print(f"    Data: {r['data'][:200]}")

        # 데이터 응답
        print(f"\n  데이터 응답 {len(captured_responses)}개:")
        for r in captured_responses:
            print(f"  {r['url']}: {r['body'][:400]}")

        # 결과 페이지 내용
        body3 = await page.inner_text("body")
        print(f"\n  페이지 (앞 1000자):\n{body3[:1000]}")

        # ════════════════════════════════════════
        # D. 직접 서비스 POST 시도
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("D. WebSquare 데이터 서비스 직접 호출")
        print("=" * 60)

        import requests as req_lib
        kofia_cookies = await context.cookies()
        sess = req_lib.Session()
        sess.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            'Referer': 'https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml',
            'Origin': 'https://dis.kofia.or.kr',
            'Accept': 'application/xml, text/xml, */*',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        })
        for c in kofia_cookies:
            if 'kofia' in c.get('domain',''):
                sess.cookies.set(c['name'], c['value'])

        # 일반적인 WebSquare 서비스 엔드포인트 패턴 시도
        endpoints = [
            "https://dis.kofia.or.kr/service/DataServices.do",
            "https://dis.kofia.or.kr/websquare/service.do",
            "https://dis.kofia.or.kr/wq/fundann/FundFeeCMSService.do",
        ]

        for ep in endpoints:
            try:
                r = sess.post(ep, data={"fundNm": "미래에셋미국나스닥100인덱스", "compCd": ""}, timeout=10)
                print(f"  {r.status_code} {ep}: {r.text[:200]}")
            except Exception as e:
                print(f"  오류: {ep}: {e}")

        await browser.close()

asyncio.run(main())
