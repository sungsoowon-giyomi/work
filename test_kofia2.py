"""
dis.kofia.or.kr 심화 탐색:
1. 펀드공시검색 (DISFundAnnSrch) → 투자설명서 PDF 다운로드
2. 펀드 보수 및 비용 (DISFundFeeCMS) → 수치 데이터 추출
"""
import asyncio
import json
import requests
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

        # ════════════════════════════════════════
        # A. 펀드 보수 및 비용 검색 → 수치 데이터
        # ════════════════════════════════════════
        print("=" * 60)
        print("A. 펀드 보수 및 비용 (DISFundFeeCMS)")
        print("=" * 60)

        api_fee = []
        async def on_fee(response):
            url = response.url
            ct = response.headers.get("content-type","")
            if "kofia" in url and ("json" in ct or "xml" in ct):
                try:
                    body_text = await response.text()
                    api_fee.append({"url": url[:100], "status": response.status, "body": body_text[:500]})
                except:
                    api_fee.append({"url": url[:100], "status": response.status, "body": ""})

        page.on("response", on_fee)
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)
        page.remove_listener("response", on_fee)

        print(f"API 호출 {len(api_fee)}개:")
        for c in api_fee[:5]:
            print(f"  {c['status']} {c['url']}")

        # 폼 요소
        fee_form = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll("input[type=text], select")).map(el => ({
                tag: el.tagName, id: el.id, name: el.name||"", placeholder: el.placeholder||""
            })).slice(0,15);
        }''')
        print(f"\n폼 요소:")
        for el in fee_form:
            print(f"  {el}")

        # 펀드명 검색
        try:
            fund_input = await page.query_selector('input[type="text"]')
            if fund_input:
                await fund_input.fill("미래에셋미국나스닥100")
                await asyncio.sleep(0.5)
                await fund_input.press("Enter")
                await asyncio.sleep(3)
                await page.wait_for_load_state("networkidle", timeout=10000)

                body_text = await page.inner_text("body")
                print(f"\n검색 결과 (앞 1000자):\n{body_text[:1000]}")
        except Exception as e:
            print(f"보수/비용 검색 오류: {e}")

        # ════════════════════════════════════════
        # B. 펀드공시검색 (DISFundAnnSrch) → 투자설명서
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("B. 펀드공시검색 (DISFundAnnSrch)")
        print("=" * 60)

        api_ann = []
        async def on_ann(response):
            url = response.url
            ct = response.headers.get("content-type","")
            status = response.status
            if "kofia" in url and status == 200:
                try:
                    if "json" in ct or "xml" in ct:
                        body_text = await response.text()
                        api_ann.append({"url": url[:100], "status": status, "ct": ct[:30], "body": body_text[:600]})
                    else:
                        api_ann.append({"url": url[:100], "status": status, "ct": ct[:30], "body": ""})
                except:
                    pass

        page.on("response", on_ann)
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundAnnSrch.xml&divisionId=MDIS01001000000000&serviceId=SDIS01001000000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        # 폼 요소 확인
        ann_form = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll("input, select, button")).map(el => ({
                tag: el.tagName, id: el.id, name: el.name||"",
                type: el.type||"", placeholder: el.placeholder||"",
                text: (el.innerText||el.value||"").trim().slice(0,20)
            })).filter(e => e.id || e.name || e.text).slice(0,25);
        }''')
        print(f"\n폼 요소:")
        for el in ann_form:
            print(f"  {el}")

        # 라디오/체크박스 레이블 (WebSquare 방식)
        ws_labels = await page.evaluate('''() => {
            const results = [];
            // span/div 레이블
            document.querySelectorAll(".w2label, span.label, td span, .group_title").forEach(el => {
                const txt = (el.innerText||"").trim();
                if (txt && txt.length < 30) results.push(txt);
            });
            return results.slice(0, 30);
        }''')
        print(f"\nWebSquare 레이블: {ws_labels}")

        # 페이지 텍스트 (일부)
        body_preview = (await page.inner_text("body"))[:800]
        print(f"\n페이지 텍스트:\n{body_preview}")

        # 펀드명 검색 시도
        print(f"\n--- 미래에셋 나스닥100 검색 시도 ---")
        api_ann.clear()
        try:
            # 텍스트 입력란 찾기
            input_el = await page.query_selector('#schFundNm, #fundNm, #searchFundNm, input[placeholder*="펀드"]')
            if not input_el:
                inputs = await page.query_selector_all('input[type="text"]')
                if inputs:
                    input_el = inputs[0]

            if input_el:
                await input_el.fill("미래에셋미국나스닥100인덱스")
                await asyncio.sleep(0.5)

                # 조회 버튼 찾기
                search_btn = await page.query_selector('button:has-text("조회"), button:has-text("검색"), input[value="조회"]')
                if search_btn:
                    await search_btn.click()
                else:
                    await input_el.press("Enter")

                await asyncio.sleep(3)
                await page.wait_for_load_state("networkidle", timeout=15000)
                await asyncio.sleep(2)

                # API 응답 확인
                json_calls = [c for c in api_ann if "json" in c.get("ct","") or "xml" in c.get("ct","")]
                print(f"  API 응답 {len(json_calls)}개:")
                for c in json_calls[:5]:
                    print(f"    {c['status']} {c['url']}")
                    print(f"    {c['body'][:300]}")

                # 페이지 결과
                result_text = (await page.inner_text("body"))
                print(f"\n  결과 텍스트 (앞 1000자):\n{result_text[:1000]}")

                # 결과 행에서 PDF 링크 찾기
                result_links = await page.evaluate('''() => {
                    const links = [];
                    document.querySelectorAll("tr, .list-item, .result-row").forEach(row => {
                        const text = (row.innerText||"").trim();
                        if (text.includes("투자설명서") || text.includes("설명서")) {
                            const anchors = row.querySelectorAll("a, button");
                            anchors.forEach(a => {
                                links.push({
                                    rowText: text.slice(0,100),
                                    href: a.getAttribute("href")||"",
                                    onclick: (a.getAttribute("onclick")||"").slice(0,100),
                                    text: (a.innerText||"").trim()
                                });
                            });
                        }
                    });
                    return links.slice(0,10);
                }''')
                print(f"\n  투자설명서 관련 링크: {result_links}")

        except Exception as e:
            print(f"  오류: {e}")
            import traceback
            traceback.print_exc()

        page.remove_listener("response", on_ann)

        # ════════════════════════════════════════
        # C. 핵심정보설명서 직접 다운로드 테스트
        # (R3은 funetf.co.kr에서 일부 200 확인)
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("C. FunETF R3 (핵심정보설명서) Playwright 다운로드 테스트")
        print("=" * 60)

        # FunETF 로그인
        await page.goto("https://www.funetf.co.kr/auth")
        await page.wait_for_load_state("domcontentloaded")
        await page.fill('input[name="username"]', "sungsoowon45@gmail.com")
        await page.fill('input[name="password"]', "!sungsoo0405")
        await page.locator('button:has-text("로그인")').last.click()
        await page.wait_for_url("https://www.funetf.co.kr/", timeout=15000)
        print("  FunETF 로그인 완료")

        # 미래에셋 펀드 페이지 로드 후 R3 다운로드
        r3_url = "https://www.funetf.co.kr/upload/FOK/gongsi/R3_K55301DP7470_20240913.pdf"
        r1_url = "https://www.funetf.co.kr/upload/FOK/gongsi/R1_K55301DP7470_20240913.pdf"

        # Playwright로 직접 GET 요청
        for url, label in [(r3_url, "R3 핵심정보"), (r1_url, "R1 집합투자규약")]:
            response = await page.request.get(url)
            status = response.status
            ct = response.headers.get("content-type","")
            size = len(await response.body()) if status == 200 else 0
            print(f"\n  {label}: {status} {ct[:30]} {size}B")
            if status == 200 and size > 1000:
                filename = url.split("/")[-1]
                (PDF_DIR / filename).write_bytes(await response.body())
                print(f"    ✅ 저장 완료: {filename}")

        await browser.close()

asyncio.run(main())
