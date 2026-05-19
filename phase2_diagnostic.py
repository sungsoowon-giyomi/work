"""
Phase 2 진단 스크립트:
- 나머지 24개 항목에 대한 문서 URL 탐색
- FunETF fund 페이지 (ETF도 fund 페이지 시도)
- DOM에서 모든 data-fileurl 및 href 수집
- 각 URL 다운로드 가능 여부 테스트
- dis.kofia.or.kr 검색 시도
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

# 아직 이미지 없는 항목들
MISSING_ITEMS = [
    {"name": "TIME 글로벌AI인공지능액티브",                   "type": "etf",  "sotCd": "456600", "search": "TIME 글로벌AI인공지능액티브"},
    {"name": "ACE 글로벌반도체TOP4 Plus",                   "type": "etf",  "sotCd": "446770", "search": "ACE 글로벌반도체TOP4 Plus"},
    {"name": "RISE 대형고배당10TR",                         "type": "etf",  "sotCd": "315960", "search": "RISE 대형고배당10TR"},
    {"name": "RISE 미국단기투자등급회사채액티브",                 "type": "etf",  "sotCd": "437350", "search": "RISE 미국단기투자등급회사채액티브"},
    {"name": "ACE KRX금현물",                              "type": "etf",  "sotCd": "411060", "search": "ACE KRX금현물"},
    {"name": "TIGER 글로벌멀티에셋TIF액티브",                   "type": "etf",  "sotCd": "440340", "search": "TIGER 글로벌멀티에셋TIF"},
    {"name": "미래에셋미국나스닥100인덱스증권자투자신탁(주식)(UH)C-P2",    "type": "fund", "code": "PAM078", "search": "미래에셋미국나스닥100인덱스증권자투자신탁"},
    {"name": "KCGI코리아퇴직연금증권자투자신탁[주식]C",              "type": "fund", "code": "PAMJ03", "search": "KCGI코리아퇴직연금증권자투자신탁"},
    {"name": "NH-Amundi 하나로단기채증권투자신탁[채권]C-P2(퇴직연금)",  "type": "fund", "code": "PDNC74", "search": "NH-Amundi 하나로단기채증권투자신탁"},
    {"name": "신한MAN글로벌투자등급채권증권투자신탁(H)[채권-재간접형]C-r",  "type": "fund", "code": "PGJA46", "search": "신한MAN글로벌투자등급채권증권투자신탁"},
    {"name": "신한골드증권투자신탁제1호[주식]C-r",                  "type": "fund", "code": "PAJA24", "search": "신한골드증권투자신탁"},
    {"name": "하나글로벌리츠부동산자투자신탁[재간접형]C-P2",            "type": "fund", "code": "PGD862", "search": "하나글로벌리츠부동산자투자신탁"},
    {"name": "NH-Amundi하나로적격TDF2040증권투자[주식혼합-재간접형]C-P2","type": "fund", "code": "PGNC86", "search": "NH-Amundi하나로적격TDF2040"},
    {"name": "KB온국민적격TDF2040증권자투자[주식혼합-재간접형](H)C-퇴직", "type": "fund", "code": "PGZA34", "search": "KB온국민적격TDF2040"},
    {"name": "미래에셋전략배분적격TDF2040혼합자산자투자신탁C-P2",        "type": "fund", "code": "PBM573", "search": "미래에셋전략배분적격TDF2040"},
]

def try_download(url, session, label=""):
    """URL 다운로드 가능 여부 확인 (HEAD 요청)"""
    try:
        r = session.head(url, timeout=10, allow_redirects=True)
        size = r.headers.get('content-length', '?')
        ct = r.headers.get('content-type', '')
        return r.status_code, size, ct
    except Exception as e:
        return -1, 0, str(e)[:50]


async def search_funetf(page, search_term):
    """FunETF 검색 API"""
    encoded = urllib.parse.quote(search_term)
    url = f"https://www.funetf.co.kr/api/public/main/search/all?schVal={encoded}&reSchVal=&schKeyword="
    await page.goto(url)
    await asyncio.sleep(0.5)
    try:
        data = json.loads(await page.inner_text("body"))
        return data.get("etfList", {}).get("content", []), data.get("fundList", {}).get("content", [])
    except:
        return [], []


async def get_fund_page_docs(page, fund_cd):
    """FunETF fund 상세 페이지에서 모든 문서 URL + API 응답 수집"""
    api_calls = {}
    all_doc_urls = []

    async def on_response(response):
        url = response.url
        ct = response.headers.get("content-type", "")
        if "funetf.co.kr/api" in url and "json" in ct:
            endpoint = url.split("/api/public/product/view/")[-1].split("?")[0] if "/api/public/product/view/" in url else url.split("/api/")[-1].split("?")[0]
            try:
                body = await response.json()
                api_calls[endpoint] = body
            except:
                pass

    page.on("response", on_response)
    try:
        url = f"https://www.funetf.co.kr/product/fund/view/{fund_cd}"
        await page.goto(url)
        await page.wait_for_load_state("networkidle", timeout=20000)
        await asyncio.sleep(2)

        # 페이지 텍스트 체크
        body_text = await page.inner_text("body")
        page_ok = "서비스 이용에 불편" not in body_text

        # ALL data-fileurl, href, data-* 속성 수집
        dom_docs = await page.evaluate('''() => {
            const results = [];
            // data-fileurl 속성
            document.querySelectorAll("[data-fileurl]").forEach(el => {
                const href = el.getAttribute("data-fileurl") || "";
                if (href) results.push({attr: "data-fileurl", val: href, tag: el.tagName});
            });
            // href 속성 (PDF 링크)
            document.querySelectorAll("a[href*='.pdf'], a[href*='/upload/']").forEach(el => {
                results.push({attr: "href", val: el.getAttribute("href"), tag: el.tagName, text: (el.innerText || "").trim()});
            });
            // onclick 속성
            document.querySelectorAll("[onclick*='upload'], [onclick*='.pdf']").forEach(el => {
                results.push({attr: "onclick", val: el.getAttribute("onclick"), tag: el.tagName, text: (el.innerText || "").trim()});
            });
            // data-src, data-url 등
            document.querySelectorAll("[data-src*='upload'], [data-url*='upload']").forEach(el => {
                results.push({attr: "data-src/url", val: el.getAttribute("data-src") || el.getAttribute("data-url"), tag: el.tagName});
            });
            return results;
        }''')

        return page_ok, api_calls, dom_docs
    except Exception as e:
        return False, api_calls, []
    finally:
        page.remove_listener("response", on_response)


async def check_kofia_dis(page, fund_name):
    """dis.kofia.or.kr에서 펀드 투자설명서 검색"""
    try:
        # 간이 방식: 검색 URL 시도
        encoded = urllib.parse.quote(fund_name)
        test_url = f"https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundProspectusSearch.xml"
        await page.goto(test_url)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)
        title = await page.title()
        body_preview = (await page.inner_text("body"))[:200]
        return title, body_preview
    except Exception as e:
        return "", str(e)[:100]


async def main():
    results = {}

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

        # requests 세션 (쿠키 포함)
        cookies = await context.cookies()
        sess = requests.Session()
        sess.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'})
        for c in cookies:
            sess.cookies.set(c['name'], c['value'])

        # ── 처음 2-3개 항목만 진단 (시간 절약) ──
        test_items = MISSING_ITEMS[:3]  # TIME ETF, ACE ETF, RISE ETF

        for item in test_items:
            name = item["name"]
            item_type = item["type"]
            print(f"\n{'='*60}")
            print(f"진단: {name}")

            # 1. FunETF 검색 → fundCd
            etf_list, fund_list = await search_funetf(page, item["search"])
            fund_cd = None
            rep_fid = None
            if item_type == "etf" and etf_list:
                sot_cd = item["sotCd"]
                for e in etf_list:
                    if str(e.get("sotCd","")) == sot_cd:
                        fund_cd = e.get("fundCd")
                        rep_fid = e.get("repFId") or e.get("fid")
                        print(f"  ETF 검색 결과: fundCd={fund_cd}, repFId={rep_fid}")
                        print(f"  전체 키: {list(e.keys())}")
                        print(f"  데이터: {json.dumps(e, ensure_ascii=False, indent=2)[:500]}")
                        break
            elif item_type == "fund" and fund_list:
                code = item.get("code","")
                for f in fund_list:
                    if str(f.get("shortCd","")).upper() == code.upper():
                        fund_cd = f.get("fundCd")
                        rep_fid = f.get("repFId") or f.get("fid")
                        print(f"  Fund 검색 결과: fundCd={fund_cd}, repFId={rep_fid}")
                        print(f"  데이터: {json.dumps(f, ensure_ascii=False, indent=2)[:500]}")
                        break
                if not fund_cd and fund_list:
                    fund_cd = fund_list[0].get("fundCd")
                    print(f"  첫 번째 결과 사용: fundCd={fund_cd}")
                    print(f"  데이터: {json.dumps(fund_list[0], ensure_ascii=False, indent=2)[:500]}")

            if not fund_cd:
                print(f"  ❌ fundCd 없음")
                continue

            # 2. Fund 상세 페이지 로드
            print(f"\n  📄 Fund 페이지 로드: /product/fund/view/{fund_cd}")
            page_ok, api_calls, dom_docs = await get_fund_page_docs(page, fund_cd)
            print(f"  페이지 로드: {'✅' if page_ok else '❌'}")

            # API 응답 출력
            print(f"\n  📡 API 응답 {len(api_calls)}개:")
            for endpoint, data in api_calls.items():
                print(f"    /{endpoint}")
                print(f"    {json.dumps(data, ensure_ascii=False)[:400]}")

            # DOM 문서 URL 출력
            print(f"\n  🔗 DOM 문서 링크 {len(dom_docs)}개:")
            for doc in dom_docs:
                val = doc.get("val","")
                print(f"    {doc.get('attr')}: {val[:100]}")

            # 3. URL 다운로드 테스트
            print(f"\n  📥 다운로드 테스트:")
            tested_urls = set()
            for doc in dom_docs:
                val = doc.get("val","")
                if not val or val in tested_urls:
                    continue
                tested_urls.add(val)

                # funetf.co.kr
                full_url = f"https://www.funetf.co.kr{val}" if val.startswith("/") else val
                status, size, ct = try_download(full_url, sess)
                print(f"    funetf: {status} {size}B {ct[:30]} → {val[:60]}")

                # fundsonline.co.kr 시도
                if val.startswith("/upload/FOK/"):
                    alt_url = f"https://www.fundsonline.co.kr{val}"
                    status2, size2, ct2 = try_download(alt_url, sess)
                    print(f"    FOK:    {status2} {size2}B {ct2[:30]} → {alt_url[:60]}")

            results[name] = {
                "fund_cd": fund_cd,
                "rep_fid": rep_fid,
                "page_ok": page_ok,
                "api_keys": list(api_calls.keys()),
                "dom_docs": dom_docs[:10],
            }

        # ── dis.kofia.or.kr 테스트 ──
        print(f"\n{'='*60}")
        print("dis.kofia.or.kr 테스트")
        title, body_preview = await check_kofia_dis(page, "미래에셋미국나스닥100")
        print(f"  Title: {title}")
        print(f"  Body: {body_preview}")

        await browser.close()

    # JSON 저장
    out_path = OUTPUT_DIR / "phase2_diagnostic.json"
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n✅ 진단 결과 저장: {out_path}")


if __name__ == "__main__":
    asyncio.run(main())
