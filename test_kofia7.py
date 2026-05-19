"""
나머지 4개 항목 특수 케이스 처리:
1. KODEX 미국S&P500 → & XML 이스케이프
2. RISE 대형고배당10TR → 다른 검색어
3. 삼성미국S&P500인덱스 C-P → & 이스케이프
4. 삼성미국S&P500인덱스 C-Pe → & 이스케이프
"""
import asyncio
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils
from playwright.async_api import async_playwright

API_URL = "/proframeWeb/XMLSERVICES/"
KOFIA_URL = "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000"
QUERY_DATE = "20260430"


def make_xml_body(company_code: str, fund_name: str) -> str:
    # XML 이스케이프 적용 (& → &amp;, < → &lt; 등)
    safe_fund_name = saxutils.escape(fund_name)
    return f"""<?xml version="1.0" encoding="utf-8"?>
<message>
  <proframeHeader>
    <pfmAppName>FS-DIS2</pfmAppName>
    <pfmSvcName>DISFundFeeCmsSO</pfmSvcName>
    <pfmFnName>select</pfmFnName>
  </proframeHeader>
  <systemHeader></systemHeader>
    <DISCondFuncDTO>
    <tmpV30>{QUERY_DATE}</tmpV30>
    <tmpV11>{company_code}</tmpV11>
    <tmpV12>{safe_fund_name}</tmpV12>
    <tmpV3></tmpV3>
    <tmpV5></tmpV5>
    <tmpV4></tmpV4>
</DISCondFuncDTO>
</message>"""


def parse_fee_response(xml_text: str) -> list:
    rows = []
    try:
        root = ET.fromstring(xml_text)
        for meta in root.iter("selectMeta"):
            def v(tag):
                el = meta.find(tag)
                return (el.text or "").strip() if el is not None else ""
            rows.append({
                "fund_class_name": v("tmpV2"),
                "total_fee": v("tmpV9"),
                "trading_cost": v("tmpV11"),
                "synthetic_fee": v("tmpV12"),
                "std_code": v("tmpV15"),
            })
    except Exception as e:
        print(f"  XML 파싱 오류: {e}")
    return rows


async def api_call(page, search_term: str) -> list:
    xml_body = make_xml_body("", search_term)
    result = await page.evaluate('''
        async ([url, body]) => {
            const r = await fetch(url, {
                method: "POST",
                headers: {"Content-Type": "application/xml; charset=UTF-8", "X-Requested-With": "XMLHttpRequest"},
                body: body,
                credentials: "include"
            });
            return {status: r.status, text: await r.text()};
        }
    ''', [API_URL, xml_body])
    if result.get("status") == 200 and result.get("text"):
        return parse_fee_response(result["text"])
    return []


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        # KOFIA 로드
        await page.goto(KOFIA_URL)
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        # ── 1. KODEX 미국S&P500 (& 이스케이프) ──
        print("=" * 60)
        print("1. KODEX 미국S&P500 (& → &amp; XML 이스케이프)")
        print("=" * 60)

        tests = [
            "KODEX 미국S&P500",     # with escaping
            "KODEX 미국SP500",      # without &
            "KODEX S&P500",
            "KODEX미국S&P500",
            "379800",               # KRX 코드로 시도
        ]
        for term in tests:
            rows = await api_call(page, term)
            if rows:
                print(f"✅ '{term}' → {len(rows)}행: {rows[0]['fund_class_name']}")
                print(f"   합성총보수={rows[0]['synthetic_fee']}%, 거래비용={rows[0]['trading_cost']}%")
            else:
                print(f"❌ '{term}' → 0행")

        # ── 2. 삼성미국S&P500인덱스 (& 이스케이프) ──
        print("\n" + "=" * 60)
        print("2. 삼성미국S&P500인덱스 (& 이스케이프)")
        print("=" * 60)

        tests2 = [
            "삼성미국S&P500인덱스",
            "삼성미국S&P500",
            "삼성미국SP500",
            "삼성미국S P500",
            "삼성S&P500인덱스",
        ]
        for term in tests2:
            rows = await api_call(page, term)
            if rows:
                print(f"✅ '{term}' → {len(rows)}행: {rows[0]['fund_class_name']}")
                for r in rows[:3]:
                    print(f"   {r['fund_class_name']} → {r['synthetic_fee']}%")
            else:
                print(f"❌ '{term}' → 0행")

        # ── 3. RISE 대형고배당10TR ──
        print("\n" + "=" * 60)
        print("3. RISE 대형고배당10TR")
        print("=" * 60)

        tests3 = [
            "RISE 대형고배당10TR",
            "대형고배당10TR",
            "RISE 대형고배당",
            "대형고배당10",
            "KB RISE 대형고배당",
            "고배당10TR",
        ]
        for term in tests3:
            rows = await api_call(page, term)
            if rows:
                print(f"✅ '{term}' → {len(rows)}행: {rows[0]['fund_class_name']}")
                print(f"   합성총보수={rows[0]['synthetic_fee']}%, 거래비용={rows[0]['trading_cost']}%")
            else:
                print(f"❌ '{term}' → 0행")

        await browser.close()


asyncio.run(main())
