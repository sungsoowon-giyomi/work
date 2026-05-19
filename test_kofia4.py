"""
dis.kofia.or.kr 최종 탐색:
- 펀드 검색결과에서 DP747 행 클릭
- 펀드별 보수비용비교에서 보수 데이터 추출
- 투자설명서 다운로드 경로 확인
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


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)  # 눈에 보이게
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        # 네트워크 전체 캡쳐
        all_responses = []
        async def on_all(response):
            url = response.url
            ct = response.headers.get("content-type","")
            st = response.status
            if "kofia" in url and st == 200:
                try:
                    body = await response.body()
                    all_responses.append({"url": url, "ct": ct, "size": len(body), "body": body[:500]})
                except: pass

        page.on("response", on_all)

        # ════════════════════════════════════════
        # A. DISFundFeeCMS 펀드 검색 → DP747 클릭
        # ════════════════════════════════════════
        print("=" * 60)
        print("A. 펀드별 보수비용비교 - DP747 클릭")
        print("=" * 60)

        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        # 검색
        await page.fill('#dSearchFndNm', '미래에셋미국나스닥100인덱스')
        await page.keyboard.press('Enter')
        await asyncio.sleep(3)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)

        # DP747 행 클릭 시도
        print("  DP747 행 찾기...")
        clicked = await page.evaluate('''() => {
            const rows = document.querySelectorAll("tr");
            for (const row of rows) {
                if ((row.innerText||"").includes("DP747")) {
                    // 행의 링크 또는 클릭가능한 요소 찾기
                    const clickable = row.querySelector("a, button, span.link, td:first-child");
                    if (clickable) {
                        clickable.click();
                        return "clicked: " + (clickable.innerText||clickable.tagName||"element");
                    }
                    row.click();
                    return "row_clicked: DP747";
                }
            }
            return "not_found";
        }''')
        print(f"  클릭 결과: {clicked}")
        await asyncio.sleep(3)
        await page.wait_for_load_state("networkidle", timeout=15000)

        body_after = await page.inner_text("body")
        print(f"  클릭 후 페이지 (앞 1000자):\n{body_after[:1000]}")

        # JSON API 응답 분석
        json_responses = [r for r in all_responses if "json" in r.get("ct","")]
        print(f"\n  JSON 응답 {len(json_responses)}개:")
        for r in json_responses[-5:]:
            try:
                data = json.loads(r["body"])
                print(f"    {r['url'][:80]}")
                print(f"    {json.dumps(data, ensure_ascii=False)[:300]}")
            except:
                print(f"    {r['url'][:80]}: {r['body'][:100]}")

        # ════════════════════════════════════════
        # B. DISFundMgrFundWebSrch에서 DP747 클릭
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("B. 펀드공시검색 - DP747 선택 후 공시 조회")
        print("=" * 60)

        all_responses.clear()

        # DISFundAnnSrch 이동
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundAnnSrch.xml&divisionId=MDIS01001000000000&serviceId=SDIS01001000000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        # 펀드 검색
        await page.fill('#dSearchFndNm', '미래에셋미국나스닥100인덱스')
        await page.keyboard.press('Enter')
        await asyncio.sleep(3)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)

        # DP747 행에서 "펀드정보" 열 클릭
        print("  DP747 행에서 펀드정보 클릭...")
        clicked2 = await page.evaluate('''() => {
            const rows = document.querySelectorAll("tr");
            for (const row of rows) {
                const txt = (row.innerText||"");
                if (txt.includes("DP747") && txt.includes("자투자신탁")) {
                    const tds = row.querySelectorAll("td");
                    // 마지막 td (펀드정보)
                    const lastTd = tds[tds.length - 1];
                    const link = lastTd ? lastTd.querySelector("a, button, span") : null;
                    if (link) {
                        link.click();
                        return "info_clicked: " + (link.innerText||"").trim();
                    }
                    // 펀드명 td 클릭
                    const nameTd = tds[1];
                    if (nameTd) {
                        nameTd.click();
                        return "name_td_clicked";
                    }
                    return "no_clickable";
                }
            }
            // 없으면 아무 행이나 클릭
            const firstRow = document.querySelector("tr + tr");
            if (firstRow) { firstRow.click(); return "first_row_clicked"; }
            return "no_rows";
        }''')
        print(f"  클릭 결과: {clicked2}")
        await asyncio.sleep(3)
        await page.wait_for_load_state("networkidle", timeout=15000)

        # 현재 URL 확인
        print(f"  현재 URL: {page.url}")
        body2 = await page.inner_text("body")
        print(f"  페이지 (앞 1500자):\n{body2[:1500]}")

        # PDF 응답이 있는지
        pdf_responses = [r for r in all_responses if "pdf" in r.get("ct","")]
        if pdf_responses:
            print(f"  PDF 응답 발견: {[r['url'] for r in pdf_responses]}")

        # ════════════════════════════════════════
        # C. 직접 URL 패턴 시도
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("C. dis.kofia.or.kr 직접 URL 패턴 시도")
        print("=" * 60)

        # 펀드 상세/공시 URL 패턴들
        test_urls = [
            f"https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fund/DISFundDetail.xml&fundCd=DP747",
            f"https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundAnnSrch.xml&fundCd=DP747",
            f"https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundAnnDtl.xml&fundCd=DP747",
        ]

        kofia_cookies = await context.cookies()
        sess = requests.Session()
        sess.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            'Referer': 'https://dis.kofia.or.kr/'
        })
        for c in kofia_cookies:
            if 'kofia' in c.get('domain',''):
                sess.cookies.set(c['name'], c['value'])

        for url in test_urls:
            try:
                r = sess.get(url, timeout=10)
                print(f"  {r.status_code} {url.split('=')[-1]}")
            except Exception as e:
                print(f"  오류: {e}")

        # ════════════════════════════════════════
        # D. 데이터 서비스 API 탐색 (WebSquare service call)
        # ════════════════════════════════════════
        print("\n" + "=" * 60)
        print("D. dis.kofia.or.kr WebSquare 서비스 API 탐색")
        print("=" * 60)

        # 펀드별 보수비용 조회 서비스 직접 호출
        service_tests = [
            ("https://dis.kofia.or.kr/wq/fundann/DISFundFeeCMS.xml", "GET"),
            ("https://dis.kofia.or.kr/service/fund/FundFeeCMSService.do", "POST"),
            ("https://dis.kofia.or.kr/service/FundAnnService.do", "POST"),
        ]

        all_responses.clear()

        # DISFundFeeCMS 다시 로드하면서 API 캡쳐
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(2)

        # 검색
        await page.fill('#fundNm', '미래에셋미국나스닥100인덱스')
        await page.keyboard.press('Enter')
        await asyncio.sleep(3)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)

        # 조회 버튼 찾아 클릭
        btns = await page.evaluate('''() => Array.from(document.querySelectorAll("button, input[type=button], a")).map(b => ({
            text: (b.innerText||b.value||"").trim(), id: b.id||""
        })).filter(b => b.text && b.text.length < 20)''')
        print(f"  버튼 목록: {btns[:10]}")

        # 현재 페이지 API 응답 중 JSON 데이터 찾기
        json_data = [r for r in all_responses if "json" in r.get("ct","") and len(r.get("body", b"")) > 100]
        print(f"\n  JSON 데이터 응답 {len(json_data)}개:")
        for r in json_data:
            try:
                data = json.loads(r["body"])
                print(f"    {r['url'][:80]}")
                print(f"    {json.dumps(data, ensure_ascii=False)[:400]}")
            except:
                print(f"    {r['url'][:80]}: {r['body'][:100]}")

        # 결과 텍스트
        body3 = await page.inner_text("body")
        print(f"\n  결과 (앞 800자):\n{body3[:800]}")

        await asyncio.sleep(5)
        await browser.close()
        print("\n✅ 완료")

asyncio.run(main())
