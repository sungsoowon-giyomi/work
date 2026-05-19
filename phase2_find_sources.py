"""
Phase 2 소스 탐색:
1. 비삼성 펀드의 FunETF DOM R2 URL 수집 후 fundsonline.co.kr 테스트
2. dis.kofia.or.kr에서 투자설명서 검색 시도
3. DART 웹 검색 시도
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
PDF_DIR.mkdir(exist_ok=True)

# 비삼성 펀드 샘플 (각기 다른 운용사)
TEST_FUNDS = [
    {"name": "미래에셋미국나스닥100인덱스", "code": "PAM078", "search": "미래에셋미국나스닥100인덱스증권자투자신탁"},
    {"name": "NH-Amundi하나로단기채", "code": "PDNC74", "search": "NH-Amundi 하나로단기채증권투자신탁"},
    {"name": "신한골드증권투자신탁제1호", "code": "PAJA24", "search": "신한골드증권투자신탁"},
    {"name": "KB온국민TDF2040", "code": "PGZA34", "search": "KB온국민적격TDF2040"},
    {"name": "하나글로벌리츠", "code": "PGD862", "search": "하나글로벌리츠부동산자투자신탁"},
]

# 비삼성 ETF 샘플
TEST_ETFS = [
    {"name": "TIME 글로벌AI", "sotCd": "456600", "search": "TIME 글로벌AI인공지능액티브"},
    {"name": "ACE 글로벌반도체", "sotCd": "446770", "search": "ACE 글로벌반도체TOP4 Plus"},
    {"name": "TIGER 글로벌멀티에셋", "sotCd": "440340", "search": "TIGER 글로벌멀티에셋TIF"},
]

def head_test(url, session, label=""):
    try:
        r = session.head(url, timeout=10, allow_redirects=True)
        ct = r.headers.get('content-type', '')
        size = r.headers.get('content-length', '?')
        return r.status_code, ct[:30], size
    except Exception as e:
        return -1, str(e)[:30], 0

def get_test(url, session):
    try:
        r = session.get(url, timeout=15, allow_redirects=True)
        ct = r.headers.get('content-type', '')
        return r.status_code, ct[:40], len(r.content), r.content[:100]
    except Exception as e:
        return -1, str(e)[:40], 0, b""


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
        print("✅ FunETF 로그인 성공")

        cookies = await context.cookies()
        sess = requests.Session()
        sess.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'})
        for c in cookies:
            sess.cookies.set(c['name'], c['value'])

        all_r2_urls = {}  # name -> list of doc URLs

        # ══════════════════════════════════════════
        # 1. FunETF 비삼성 펀드 페이지 → R2 URL 수집
        # ══════════════════════════════════════════
        print("\n" + "="*60)
        print("1. FunETF 비삼성 펀드 R2 URL 수집")
        print("="*60)

        for fund in TEST_FUNDS:
            name = fund["name"]
            code = fund["code"]
            search = fund["search"]
            print(f"\n  📂 {name}")

            # 검색
            encoded = urllib.parse.quote(search)
            await page.goto(f"https://www.funetf.co.kr/api/public/main/search/all?schVal={encoded}&reSchVal=&schKeyword=")
            await asyncio.sleep(0.5)
            data = json.loads(await page.inner_text("body"))
            fund_list = data.get("fundList", {}).get("content", [])

            fund_cd = None
            rep_fid = None
            for f in fund_list:
                if str(f.get("shortCd","")).upper() == code.upper():
                    fund_cd = f.get("fundCd")
                    rep_fid = f.get("repFId")
                    break
            if not fund_cd and fund_list:
                fund_cd = fund_list[0].get("fundCd")
                rep_fid = fund_list[0].get("repFId")

            if not fund_cd:
                print(f"    ❌ fundCd 없음")
                continue

            print(f"    fundCd={fund_cd}, repFId={rep_fid}")

            # 펀드 상세 페이지 로드 → data-fileurl 수집
            api_calls = {}
            async def on_resp(response):
                url = response.url
                ct = response.headers.get("content-type", "")
                if "funetf.co.kr/api" in url and "json" in ct:
                    ep = url.split("/api/")[-1].split("?")[0]
                    try:
                        body = await response.json()
                        api_calls[ep] = body
                    except: pass

            page.on("response", on_resp)
            await page.goto(f"https://www.funetf.co.kr/product/fund/view/{fund_cd}")
            await page.wait_for_load_state("networkidle", timeout=20000)
            await asyncio.sleep(2)
            page.remove_listener("response", on_resp)

            page_ok = "서비스 이용에 불편" not in (await page.inner_text("body"))
            print(f"    페이지: {'✅' if page_ok else '❌'}")

            dom_docs = await page.evaluate('''() => {
                const r = [];
                document.querySelectorAll("[data-fileurl]").forEach(el => {
                    r.push({val: el.getAttribute("data-fileurl"), text: (el.innerText||"").trim()});
                });
                return r;
            }''')

            # etfdoc API 확인
            etfdoc_key = [k for k in api_calls if "etfdoc" in k]
            if etfdoc_key:
                etfdoc = api_calls[etfdoc_key[0]]
                content = etfdoc.get("content",[]) if isinstance(etfdoc,dict) else etfdoc
                if content:
                    print(f"    etfdoc 첫번째: {json.dumps(content[0], ensure_ascii=False)[:200]}")

            print(f"    data-fileurl {len(dom_docs)}개:")
            doc_urls = []
            for doc in dom_docs:
                val = doc.get("val","")
                print(f"      {val} [{doc.get('text','')}]")
                if val:
                    doc_urls.append(val)

            all_r2_urls[name] = {"fund_cd": fund_cd, "rep_fid": rep_fid, "doc_urls": doc_urls}

        # ══════════════════════════════════════════
        # 2. 수집된 URL을 여러 도메인에서 테스트
        # ══════════════════════════════════════════
        print("\n" + "="*60)
        print("2. 대체 도메인 다운로드 테스트")
        print("="*60)

        alt_domains = [
            "https://www.funetf.co.kr",
            "https://www.fundsonline.co.kr",
            "https://dis.fundsonline.co.kr",
            "https://file.fundsonline.co.kr",
        ]

        for name, info in all_r2_urls.items():
            print(f"\n  📂 {name}")
            for doc_url in info["doc_urls"]:
                print(f"  URL: {doc_url}")
                for domain in alt_domains:
                    full = domain + doc_url if doc_url.startswith("/") else doc_url
                    status, ct, size = head_test(full, sess)
                    marker = "✅" if status == 200 else "❌"
                    print(f"    {marker} {domain}: {status} {ct} {size}B")

        # ══════════════════════════════════════════
        # 3. dis.kofia.or.kr 탐색
        # ══════════════════════════════════════════
        print("\n" + "="*60)
        print("3. dis.kofia.or.kr 탐색")
        print("="*60)

        # 네트워크 요청 캡쳐
        kofia_api_calls = []
        async def on_kofia(response):
            url = response.url
            ct = response.headers.get("content-type","")
            if "kofia" in url and url != "https://dis.kofia.or.kr/":
                kofia_api_calls.append({"url": url, "status": response.status, "ct": ct})

        page.on("response", on_kofia)
        await page.goto("https://dis.kofia.or.kr/")
        await page.wait_for_load_state("networkidle", timeout=20000)
        await asyncio.sleep(3)
        page.remove_listener("response", on_kofia)

        title = await page.title()
        print(f"  Title: {title}")
        print(f"  Network calls {len(kofia_api_calls)}개:")
        for c in kofia_api_calls[:10]:
            print(f"    {c['status']} {c['url'][:80]}")

        # 검색 폼 찾기
        form_elements = await page.evaluate('''() => {
            const inputs = Array.from(document.querySelectorAll("input, select, button")).map(el => ({
                tag: el.tagName, type: el.type||"", name: el.name||"", id: el.id||"",
                placeholder: el.placeholder||"", text: (el.innerText||"").trim().slice(0,20)
            }));
            return inputs.filter(e => e.name || e.id || e.text).slice(0,20);
        }''')
        print(f"\n  폼 요소:")
        for el in form_elements:
            print(f"    {el}")

        # ══════════════════════════════════════════
        # 4. 비KODEX ETF → 다른 FunETF URL 패턴 시도
        # ══════════════════════════════════════════
        print("\n" + "="*60)
        print("4. 비KODEX ETF 대안 URL 패턴 탐색")
        print("="*60)

        for etf in TEST_ETFS:
            name = etf["name"]
            sot_cd = etf["sotCd"]
            print(f"\n  📊 {name} (sotCd={sot_cd})")

            # 검색으로 shortCd 확인
            encoded = urllib.parse.quote(etf["search"])
            await page.goto(f"https://www.funetf.co.kr/api/public/main/search/all?schVal={encoded}&reSchVal=&schKeyword=")
            await asyncio.sleep(0.5)
            data = json.loads(await page.inner_text("body"))
            etf_list = data.get("etfList", {}).get("content", [])

            short_cd = None
            fund_cd = None
            for e in etf_list:
                if str(e.get("sotCd","")) == sot_cd:
                    short_cd = e.get("shortCd","")
                    fund_cd = e.get("fundCd","")
                    print(f"    shortCd={short_cd}, fundCd={fund_cd}, repFId={e.get('repFId')}, fid={e.get('fid')}")
                    break

            if short_cd:
                # shortCd 기반 URL 패턴 시도
                test_urls = [
                    f"https://www.funetf.co.kr/upload/invest/{short_cd}-A.pdf",
                    f"https://www.funetf.co.kr/upload/invest/{short_cd}.pdf",
                    f"https://www.fundsonline.co.kr/upload/invest/{short_cd}-A.pdf",
                ]
                for test_url in test_urls:
                    status, ct, size = head_test(test_url, sess)
                    marker = "✅" if status == 200 else "❌"
                    print(f"    {marker} {status} {test_url.split('/')[-1]}")

        # ══════════════════════════════════════════
        # 5. DART 웹 검색 테스트
        # ══════════════════════════════════════════
        print("\n" + "="*60)
        print("5. DART 투자설명서 검색")
        print("="*60)

        dart_api_calls = []
        async def on_dart(response):
            url = response.url
            ct = response.headers.get("content-type","")
            if "dart.fss.or.kr" in url:
                dart_api_calls.append({"url": url[:80], "status": response.status, "ct": ct[:30]})

        page.on("response", on_dart)
        await page.goto("https://dart.fss.or.kr/dsab002/main.do")
        await page.wait_for_load_state("networkidle", timeout=20000)
        await asyncio.sleep(3)
        page.remove_listener("response", on_dart)

        print(f"  DART 네트워크 호출 {len(dart_api_calls)}개:")
        for c in dart_api_calls[:10]:
            print(f"    {c['status']} {c['url']}")

        # DART에서 미래에셋 나스닥100 검색 시도
        try:
            # 검색 입력란 찾기
            dart_elements = await page.evaluate('''() => {
                return Array.from(document.querySelectorAll("input, select")).map(el => ({
                    tag: el.tagName, type: el.type, name: el.name, id: el.id, placeholder: el.placeholder||""
                })).slice(0,15);
            }''')
            print(f"  DART 폼 요소:")
            for el in dart_elements:
                print(f"    {el}")
        except:
            pass

        await browser.close()
        print("\n✅ 완료")

asyncio.run(main())
