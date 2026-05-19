"""
dis.kofia.or.kr proframeWeb/XMLSERVICES/ API 완전 캡쳐 및 직접 호출:
- 전체 POST body 캡쳐 (truncation 없음)
- DISFundFeeCmsSO 직접 호출로 합성총보수/증권거래비용 추출
- 투자설명서 관련 서비스명 탐색
"""
import asyncio
import json
import re
import requests
from pathlib import Path
from playwright.async_api import async_playwright

OUTPUT_DIR = Path(r"C:\Users\jswon\Desktop\업무\funetf_scraper\output")
OUTPUT_DIR.mkdir(exist_ok=True)


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        # ════════════════════════════════════════
        # FULL POST body 캡쳐 (route intercept)
        # ════════════════════════════════════════
        print("=" * 70)
        print("1. DISFundFeeCMS 검색 - 전체 POST body 캡쳐")
        print("=" * 70)

        full_posts = []  # (url, post_data, response_text)

        async def handle_route(route):
            req = route.request
            if "proframeWeb/XMLSERVICES" in req.url:
                post_data = req.post_data or ""
                # 계속 진행 (fetch)
                response = await route.fetch()
                resp_text = await response.text()
                full_posts.append({
                    "url": req.url,
                    "post_data": post_data,
                    "status": response.status,
                    "resp_text": resp_text
                })
                await route.fulfill(response=response)
            else:
                await route.continue_()

        await page.route("**/*", handle_route)

        # DISFundFeeCMS 이동
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        # 운용사 선택 및 펀드명 입력
        try:
            select_el = page.locator('#compCd_input_0')
            cnt = await select_el.count()
            if cnt > 0:
                options = await select_el.evaluate('el => Array.from(el.options).map(o => ({v:o.value, t:o.text}))')
                miraeasset = next((o for o in options if '미래에셋자산' in o['t']), None)
                if miraeasset:
                    await select_el.select_option(miraeasset['v'])
                    print(f"  운용사 선택: {miraeasset['t']}")
        except Exception as e:
            print(f"  운용사 선택 오류: {e}")

        # 펀드명 입력 (#fundNm 또는 #dSearchFndNm)
        for fld in ['#fundNm', '#dSearchFndNm', 'input[type="text"]']:
            try:
                el = page.locator(fld).first
                if await el.count() > 0:
                    await el.fill('미래에셋미국나스닥100인덱스')
                    print(f"  입력 필드: {fld}")
                    break
            except:
                pass

        full_posts.clear()
        await page.keyboard.press('Enter')
        await asyncio.sleep(5)
        await page.wait_for_load_state("networkidle", timeout=20000)
        await asyncio.sleep(2)

        # 결과 출력
        print(f"\n  캡쳐된 XMLSERVICES POST {len(full_posts)}개:")
        for i, p_data in enumerate(full_posts):
            svc_match = re.search(r'<pfmSvcName>(\w+)</pfmSvcName>', p_data['post_data'])
            fn_match = re.search(r'<pfmFnName>(\w+)</pfmFnName>', p_data['post_data'])
            svc = svc_match.group(1) if svc_match else "?"
            fn = fn_match.group(1) if fn_match else "?"
            print(f"\n  [{i+1}] Service: {svc}, Function: {fn}")
            print(f"  POST body ({len(p_data['post_data'])}자):")
            print(p_data['post_data'])
            print(f"\n  Response ({p_data['status']}, {len(p_data['resp_text'])}자):")
            print(p_data['resp_text'][:2000])

        # 검색 결과 POST 중 DISFundFeeCmsSO 찾기
        fee_post = next((p for p in full_posts if 'DISFundFeeCmsSO' in p['post_data'] and 'select' in p['post_data']), None)
        if fee_post:
            print("\n  ✅ DISFundFeeCmsSO select POST body 확보!")
            (OUTPUT_DIR / "fee_cms_post_body.xml").write_text(fee_post['post_data'], encoding='utf-8')
            (OUTPUT_DIR / "fee_cms_response.xml").write_text(fee_post['resp_text'], encoding='utf-8')
            print(f"  저장: fee_cms_post_body.xml, fee_cms_response.xml")

        # 결과 페이지 텍스트
        body_txt = await page.inner_text("body")
        print(f"\n  페이지 결과 (앞 1000자):\n{body_txt[:1000]}")

        # ════════════════════════════════════════
        # 2. DISFundAnnSrch 탐색 - 투자설명서 서비스 탐색
        # ════════════════════════════════════════
        print("\n" + "=" * 70)
        print("2. DISFundAnnSrch - 투자설명서 관련 서비스 탐색")
        print("=" * 70)

        full_posts.clear()

        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundAnnSrch.xml&divisionId=MDIS01001000000000&serviceId=SDIS01001000000")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        # 펀드 검색
        for fld in ['#dSearchFndNm', '#fundNm', 'input[type="text"]']:
            try:
                el = page.locator(fld).first
                if await el.count() > 0:
                    await el.fill('미래에셋미국나스닥100인덱스')
                    break
            except:
                pass

        full_posts.clear()
        await page.keyboard.press('Enter')
        await asyncio.sleep(5)
        await page.wait_for_load_state("networkidle", timeout=20000)
        await asyncio.sleep(2)

        print(f"\n  DISFundAnnSrch POST {len(full_posts)}개:")
        for i, p_data in enumerate(full_posts):
            svc_match = re.search(r'<pfmSvcName>(\w+)</pfmSvcName>', p_data['post_data'])
            fn_match = re.search(r'<pfmFnName>(\w+)</pfmFnName>', p_data['post_data'])
            svc = svc_match.group(1) if svc_match else "?"
            fn = fn_match.group(1) if fn_match else "?"
            print(f"  [{i+1}] {svc}.{fn}")
            print(f"  POST: {p_data['post_data'][:500]}")
            print(f"  RESP: {p_data['resp_text'][:500]}")

        # ════════════════════════════════════════
        # 3. 검색 결과에서 DP747 행 클릭 → 투자설명서 POST 캡쳐
        # ════════════════════════════════════════
        print("\n" + "=" * 70)
        print("3. DP747 행 클릭 → 문서 목록 POST 캡쳐")
        print("=" * 70)

        full_posts.clear()

        # JS로 DP747 행 클릭 (WebSquare 방식)
        await page.evaluate('''() => {
            const cells = document.querySelectorAll("td, span, div");
            for (const cell of cells) {
                if ((cell.innerText||"").trim() === "DP747") {
                    // 부모 행의 첫 td 또는 a 태그
                    const row = cell.closest("tr") || cell.parentElement;
                    const clickable = row ? row.querySelector("a[onclick], span[onclick], td[onclick]") : null;
                    if (clickable) {
                        clickable.click();
                        return "clicked_onclick: " + clickable.tagName;
                    }
                    if (row) {
                        row.click();
                        return "clicked_row";
                    }
                    cell.click();
                    return "clicked_cell";
                }
            }
            // 모든 onclick 속성 중 펀드 관련 찾기
            const onclickEls = document.querySelectorAll("[onclick]");
            for (const el of onclickEls) {
                const oc = el.getAttribute("onclick") || "";
                if (oc.includes("DP747") || oc.includes("fund") || oc.includes("Fund")) {
                    el.click();
                    return "onclick_element: " + oc.slice(0, 80);
                }
            }
            return "not_found";
        }''')

        await asyncio.sleep(3)
        await page.wait_for_load_state("networkidle", timeout=10000)
        await asyncio.sleep(2)

        if full_posts:
            print(f"\n  클릭 후 POST {len(full_posts)}개:")
            for p_data in full_posts:
                svc_match = re.search(r'<pfmSvcName>(\w+)</pfmSvcName>', p_data['post_data'])
                svc = svc_match.group(1) if svc_match else "?"
                print(f"  Service: {svc}")
                print(f"  POST: {p_data['post_data'][:600]}")
                print(f"  RESP: {p_data['resp_text'][:600]}")
        else:
            print("  클릭 후 POST 없음")

        # ════════════════════════════════════════
        # 4. dis.kofia.or.kr 직접 XML POST 호출
        # ════════════════════════════════════════
        print("\n" + "=" * 70)
        print("4. proframeWeb/XMLSERVICES/ 직접 호출 테스트")
        print("=" * 70)

        kofia_cookies = await context.cookies()
        sess = requests.Session()
        sess.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            'Referer': 'https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml',
            'Origin': 'https://dis.kofia.or.kr',
            'Content-Type': 'application/xml; charset=UTF-8',
            'Accept': 'application/xml, text/xml, */*',
            'X-Requested-With': 'XMLHttpRequest',
        })
        for c in kofia_cookies:
            if 'kofia' in c.get('domain', ''):
                sess.cookies.set(c['name'], c['value'])

        # DISFundFeeCmsSO 직접 호출 (fund code로)
        xml_body_fee = '''<?xml version="1.0" encoding="utf-8"?>
<message>
  <proframeHeader>
    <pfmAppName>FS-DIS2</pfmAppName>
    <pfmSvcName>DISFundFeeCmsSO</pfmSvcName>
    <pfmFnName>select</pfmFnName>
  </proframeHeader>
  <Parameter>
    <DISFundFeeCmsInput>
      <fundCd>DP747</fundCd>
      <fundNm></fundNm>
      <compCd></compCd>
      <fundTypCd></fundTypCd>
      <pageReqCount>10</pageReqCount>
      <currentPage>1</currentPage>
    </DISFundFeeCmsInput>
  </Parameter>
</message>'''

        try:
            r = sess.post("https://dis.kofia.or.kr/proframeWeb/XMLSERVICES/", data=xml_body_fee.encode('utf-8'), timeout=15)
            print(f"\n  DISFundFeeCmsSO (fundCd=DP747): {r.status_code}")
            print(f"  Response: {r.text[:2000]}")
            if r.status_code == 200:
                (OUTPUT_DIR / "direct_fee_resp.xml").write_text(r.text, encoding='utf-8')
        except Exception as e:
            print(f"  오류: {e}")

        # fundNm으로 시도
        xml_body_fee2 = '''<?xml version="1.0" encoding="utf-8"?>
<message>
  <proframeHeader>
    <pfmAppName>FS-DIS2</pfmAppName>
    <pfmSvcName>DISFundFeeCmsSO</pfmSvcName>
    <pfmFnName>select</pfmFnName>
  </proframeHeader>
  <Parameter>
    <DISFundFeeCmsInput>
      <fundCd></fundCd>
      <fundNm>미래에셋미국나스닥100인덱스</fundNm>
      <compCd></compCd>
      <fundTypCd></fundTypCd>
      <pageReqCount>10</pageReqCount>
      <currentPage>1</currentPage>
    </DISFundFeeCmsInput>
  </Parameter>
</message>'''

        try:
            r2 = sess.post("https://dis.kofia.or.kr/proframeWeb/XMLSERVICES/", data=xml_body_fee2.encode('utf-8'), timeout=15)
            print(f"\n  DISFundFeeCmsSO (fundNm=미래에셋미국나스닥100인덱스): {r2.status_code}")
            print(f"  Response: {r2.text[:2000]}")
            if r2.status_code == 200:
                (OUTPUT_DIR / "direct_fee_resp2.xml").write_text(r2.text, encoding='utf-8')
        except Exception as e:
            print(f"  오류: {e}")

        # DISFundAnnInqSO - 투자설명서 조회 서비스 시도
        xml_body_ann = '''<?xml version="1.0" encoding="utf-8"?>
<message>
  <proframeHeader>
    <pfmAppName>FS-DIS2</pfmAppName>
    <pfmSvcName>DISFundAnnInqSO</pfmSvcName>
    <pfmFnName>select</pfmFnName>
  </proframeHeader>
  <Parameter>
    <DISFundAnnInqInput>
      <fundCd>DP747</fundCd>
      <annTypCd>01</annTypCd>
    </DISFundAnnInqInput>
  </Parameter>
</message>'''

        try:
            r3 = sess.post("https://dis.kofia.or.kr/proframeWeb/XMLSERVICES/", data=xml_body_ann.encode('utf-8'), timeout=15)
            print(f"\n  DISFundAnnInqSO (fundCd=DP747): {r3.status_code}")
            print(f"  Response: {r3.text[:1000]}")
        except Exception as e:
            print(f"  오류: {e}")

        await browser.close()
        print("\n✅ 완료")


asyncio.run(main())
