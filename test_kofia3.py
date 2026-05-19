"""
dis.kofia.or.kr에서 투자설명서 다운로드:
- KOFIA 펀드코드(DP747 등)로 검색
- 펀드공시검색에서 투자설명서 찾기
- 다운로드 URL 추출
"""
import asyncio
import json
import re
import requests
from pathlib import Path
from playwright.async_api import async_playwright

OUTPUT_DIR = Path(r"C:\Users\jswon\Desktop\업무\funetf_scraper\output")
PDF_DIR = OUTPUT_DIR / "pdfs"
PDF_DIR.mkdir(exist_ok=True)

async def search_kofia_ann(page, fund_name_or_code, result_holder):
    """dis.kofia.or.kr 펀드공시검색 실행"""
    api_calls = []
    pdf_urls = []

    async def on_resp(response):
        url = response.url
        ct = response.headers.get("content-type","")
        st = response.status
        if "kofia" in url and st == 200:
            if "json" in ct:
                try:
                    body = await response.json()
                    api_calls.append({"url": url[:100], "body": body})
                except: pass
            elif "pdf" in ct:
                pdf_urls.append(url)

    page.on("response", on_resp)

    # 1. 펀드공시검색 페이지 이동
    await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundAnnSrch.xml&divisionId=MDIS01001000000000&serviceId=SDIS01001000000")
    await page.wait_for_load_state("networkidle", timeout=30000)
    await asyncio.sleep(3)

    # 2. 펀드 검색 입력
    await page.fill('#dSearchFndNm', fund_name_or_code)
    await asyncio.sleep(0.5)

    # 3. 날짜 범위를 최대로 설정 (최근 3년)
    try:
        await page.select_option('#selectSrchDate_input_0', '3년')
    except:
        pass

    # 4. 검색 버튼 클릭
    await page.keyboard.press('Enter')
    await asyncio.sleep(3)
    await page.wait_for_load_state("networkidle", timeout=15000)

    current_url = page.url
    print(f"    현재 URL: {current_url}")

    # 5. 펀드 검색 결과 페이지에서 펀드 선택
    fund_rows = await page.evaluate('''() => {
        const rows = [];
        document.querySelectorAll("tr, .data-row, [id*='grid']").forEach(el => {
            const text = (el.innerText||"").trim();
            if (text.length > 10 && text.length < 200) rows.push(text);
        });
        return rows.slice(0, 20);
    }''')
    print(f"    펀드 목록 행: {fund_rows[:5]}")

    # 첫 번째 결과 클릭 시도
    try:
        first_link = page.locator("a, button").filter(has_text=re.compile(r'미래에셋|NH-Amundi|신한|KB|하나|KCGI|미래에셋')).first
        if await first_link.count() > 0:
            await first_link.click(timeout=3000)
            await asyncio.sleep(2)
            await page.wait_for_load_state("networkidle", timeout=10000)
            print(f"    클릭 후 URL: {page.url}")
    except:
        pass

    # 결과 텍스트
    result_text = (await page.inner_text("body"))[:1000]
    result_holder["text"] = result_text
    result_holder["api_calls"] = api_calls
    result_holder["pdf_urls"] = pdf_urls

    page.remove_listener("response", on_resp)
    return api_calls, pdf_urls


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        # ════════════════════════════════════════
        # A. 펀드 검색결과 → 펀드 클릭 → 문서 목록
        # ════════════════════════════════════════
        print("=" * 60)
        print("A. dis.kofia.or.kr DISFundFeeCMS 상세 보수 데이터")
        print("=" * 60)

        # 네트워크 응답 캡쳐
        api_all = []
        async def on_all(response):
            url = response.url
            ct = response.headers.get("content-type","")
            st = response.status
            if "kofia" in url and st == 200 and ("json" in ct or "xml" in ct):
                try:
                    body_txt = await response.text()
                    if len(body_txt) > 50:
                        api_all.append({"url": url[:100], "body": body_txt[:800]})
                except: pass

        page.on("response", on_all)

        # DISFundFeeCMS 접속
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)
        page.remove_listener("response", on_all)

        # 펀드명 검색 폼 구조 파악
        form_state = await page.evaluate('''() => {
            const result = {};
            // 모든 input의 값
            document.querySelectorAll("input[type=text]").forEach(el => {
                result[el.id||el.name||"unknown"] = el.value;
            });
            // 모든 select
            document.querySelectorAll("select").forEach(el => {
                result["select_"+el.id] = Array.from(el.options).map(o => o.text).join("|");
            });
            return result;
        }''')
        print(f"폼 초기값: {json.dumps(form_state, ensure_ascii=False, indent=2)[:500]}")

        # ════════════════════════════════════════
        # B. DISFundFeeCMS 검색 (미래에셋 나스닥100 C-P2)
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("B. DISFundFeeCMS 검색 → 보수 데이터")
        print("=" * 60)

        api_fee = []
        async def on_fee2(response):
            url = response.url
            ct = response.headers.get("content-type","")
            if "kofia" in url and response.status == 200:
                if "json" in ct or "xml" in ct:
                    try:
                        body_txt = await response.text()
                        if len(body_txt) > 100:
                            api_fee.append({"url": url[:120], "body": body_txt[:600]})
                    except: pass

        page.on("response", on_fee2)

        # 펀드명 입력 (dSearchFndNm)
        await page.fill('#dSearchFndNm', '미래에셋미국나스닥100인덱스')
        await asyncio.sleep(0.5)
        await page.keyboard.press('Enter')
        await asyncio.sleep(3)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)

        page.remove_listener("response", on_fee2)

        # 결과 페이지 분석
        body_text = await page.inner_text("body")
        print(f"결과 (앞 1500자):\n{body_text[:1500]}")

        print(f"\nAPI 응답 {len(api_fee)}개:")
        for c in api_fee[:5]:
            print(f"  {c['url']}")
            print(f"  {c['body'][:300]}")

        # 결과 행 클릭 시도
        rows_text = await page.evaluate('''() => {
            const rows = [];
            document.querySelectorAll("tr").forEach(tr => {
                const t = (tr.innerText||"").trim();
                if (t.includes("C-P2") || t.includes("C-P")) rows.push(t.slice(0,100));
            });
            return rows.slice(0,5);
        }''')
        print(f"\nC-P2 관련 행: {rows_text}")

        # ════════════════════════════════════════
        # C. 펀드공시검색 → 투자설명서 찾기
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("C. 펀드공시검색 → 투자설명서 다운로드")
        print("=" * 60)

        api_ann = []
        pdf_found = []
        async def on_ann2(response):
            url = response.url
            ct = response.headers.get("content-type","")
            st = response.status
            if "kofia" in url:
                if "json" in ct and st == 200:
                    try:
                        body_txt = await response.text()
                        if len(body_txt) > 100:
                            api_ann.append({"url": url[:100], "body": body_txt[:600]})
                    except: pass
                elif "pdf" in ct and st == 200:
                    pdf_found.append(url)
                    print(f"  ✅ PDF 발견: {url}")

        page.on("response", on_ann2)

        # 1. 메인 펀드공시 검색 페이지
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundAnnSrch.xml&divisionId=MDIS01001000000000&serviceId=SDIS01001000000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        # 2. 회사 선택 (미래에셋)
        try:
            # 회사 드롭다운에서 미래에셋 선택
            company_select = await page.query_selector('#compCd_input_0, select[id*="comp"]')
            if company_select:
                options = await page.evaluate('''(sel) => Array.from(sel.options).map(o => ({v:o.value, t:o.text}))''', company_select)
                miraeasset_opt = next((o for o in options if '미래에셋' in o['t']), None)
                if miraeasset_opt:
                    await page.select_option(f'#{company_select.get_attribute("id")}', miraeasset_opt['v'])
                    print(f"  회사 선택: {miraeasset_opt['t']}")
                    await asyncio.sleep(1)
        except Exception as e:
            print(f"  회사 선택 오류: {e}")

        # 3. 조회기간 최대화
        try:
            await page.fill('#srchStartDt_input', '2020-01-01')
            await page.fill('#srchEndDt_input', '2026-05-20')
        except: pass

        # 4. 펀드 검색
        await page.fill('#dSearchFndNm', '미래에셋미국나스닥100인덱스')
        await asyncio.sleep(0.5)

        # 5. 조회 버튼 클릭
        try:
            search_btn = await page.query_selector('button:has-text("조회")')
            if search_btn:
                await search_btn.click()
            else:
                await page.keyboard.press('Enter')
        except:
            await page.keyboard.press('Enter')

        await asyncio.sleep(4)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)

        # 결과 확인
        body3 = await page.inner_text("body")
        print(f"검색 결과 (앞 1500자):\n{body3[:1500]}")

        # 투자설명서 관련 링크 찾기
        invest_links = await page.evaluate('''() => {
            const links = [];
            document.querySelectorAll("tr, .list-item").forEach(row => {
                const txt = (row.innerText||"").trim();
                if (txt.includes("투자설명서") || txt.includes("설명서") || txt.includes("공시")) {
                    const anchors = row.querySelectorAll("a, button, span.btn");
                    anchors.forEach(a => {
                        links.push({
                            rowText: txt.slice(0,80),
                            href: a.getAttribute("href")||"",
                            onclick: (a.getAttribute("onclick")||"").slice(0,100),
                            text: (a.innerText||"").trim().slice(0,30)
                        });
                    });
                }
            });
            return links.slice(0,15);
        }''')
        print(f"\n투자설명서 관련 링크 {len(invest_links)}개:")
        for lnk in invest_links:
            print(f"  {lnk}")

        # API 응답
        print(f"\nAPI 응답 {len(api_ann)}개:")
        for c in api_ann[:3]:
            print(f"  {c['url']}")
            print(f"  {c['body'][:200]}")

        page.remove_listener("response", on_ann2)

        # ════════════════════════════════════════
        # D. DISFundFeeCMS API 직접 호출 테스트
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("D. dis.kofia.or.kr 서비스 API 직접 호출")
        print("=" * 60)

        # dis.kofia.or.kr DataService 직접 POST
        kofia_cookies = await context.cookies()
        sess = requests.Session()
        sess.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            'Referer': 'https://dis.kofia.or.kr/',
            'Origin': 'https://dis.kofia.or.kr',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'X-Requested-With': 'XMLHttpRequest',
        })
        for c in kofia_cookies:
            if 'kofia' in c.get('domain',''):
                sess.cookies.set(c['name'], c['value'])

        # 펀드 보수비용 조회 API 시도
        test_apis = [
            ("POST", "https://dis.kofia.or.kr/service/DataServices.do",
             {"serviceId": "SDIS01005001000", "fundNm": "미래에셋미국나스닥100인덱스", "fundTyp": ""}),
            ("POST", "https://dis.kofia.or.kr/service/DataServices.do",
             {"serviceId": "SDIS01001000000", "fundNm": "미래에셋미국나스닥100인덱스", "srchStartDt": "20200101", "srchEndDt": "20260520"}),
        ]

        for method, url, data in test_apis:
            try:
                r = sess.post(url, data=data, timeout=10)
                print(f"\n  POST {url}")
                print(f"  Status: {r.status_code}, CT: {r.headers.get('content-type','')[:40]}")
                print(f"  Body: {r.text[:300]}")
            except Exception as e:
                print(f"  오류: {e}")

        await browser.close()

asyncio.run(main())
