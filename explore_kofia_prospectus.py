"""
KOFIA 투자설명서 다운로드 방식 자동 탐색
- 투자설명서 메뉴 URL 찾기
- API endpoint 패턴 파악
- 다운로드 방식 확인
"""
import asyncio
import json
from pathlib import Path
from playwright.async_api import async_playwright

# KOFIA DIS 알려진 URL 패턴들
CANDIDATE_URLS = [
    # 보수비용 (이미 알고 있는 것)
    "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000",
    # 투자설명서 후보들
    "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundProspect.xml",
    "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/funddis/DISFundProspect.xml",
    "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundProsp.xml",
    "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundpub/DISFundProspect.xml",
]

# 테스트 펀드
TEST_FUND = "KODEX 200"

XML_SERVICES_URL = "/proframeWeb/XMLSERVICES/"

captured_requests = []

async def test_url(page, url):
    """URL이 유효한지 테스트 (페이지가 로드되는지)"""
    try:
        await page.goto(url, timeout=15000)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)
        title = await page.title()
        print(f"  ✅ {url[-80:]} → title='{title}'")
        return True
    except Exception as e:
        print(f"  ❌ {url[-80:]} → {e}")
        return False


async def call_xml_service(page, svc_name, fn_name="select", params: dict = None) -> dict:
    """KOFIA XMLSERVICES API 직접 호출"""
    params = params or {}
    param_xml = "\n".join(f"    <{k}>{v}</{k}>" for k, v in params.items())
    body = f"""<?xml version="1.0" encoding="utf-8"?>
<message>
  <proframeHeader>
    <pfmAppName>FS-DIS2</pfmAppName>
    <pfmSvcName>{svc_name}</pfmSvcName>
    <pfmFnName>{fn_name}</pfmFnName>
  </proframeHeader>
  <systemHeader></systemHeader>
  <DISCondFuncDTO>
{param_xml}
  </DISCondFuncDTO>
</message>"""

    result = await page.evaluate('''
        async ([url, body]) => {
            try {
                const r = await fetch(url, {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/xml; charset=UTF-8",
                        "X-Requested-With": "XMLHttpRequest"
                    },
                    body: body,
                    credentials: "include"
                });
                return {status: r.status, text: await r.text()};
            } catch(e) {
                return {status: 0, text: String(e)};
            }
        }
    ''', [XML_SERVICES_URL, body])
    return result


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        # ── 1. KOFIA 메인 로드 (쿠키 설정) ──
        print("=" * 60)
        print("1단계: KOFIA 로드 (쿠키 확보)")
        print("=" * 60)
        await page.goto(CANDIDATE_URLS[0])
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)
        print("  ✅ KOFIA 로드 완료")

        # ── 2. 투자설명서 관련 서비스명 추측 테스트 ──
        print("\n" + "=" * 60)
        print("2단계: XMLSERVICES 서비스명 탐색")
        print("=" * 60)

        # 투자설명서 관련 서비스명 후보들
        svc_candidates = [
            ("DISFundProspSO", {"tmpV12": TEST_FUND}),
            ("DISFundProspectSO", {"tmpV12": TEST_FUND}),
            ("DISFundDocSO", {"tmpV12": TEST_FUND}),
            ("DISFundPublicSO", {"tmpV12": TEST_FUND}),
            ("DISFundFileSO", {"tmpV12": TEST_FUND}),
            ("DISFundAnnSO", {"tmpV12": TEST_FUND}),
            ("DISFundDisclosureSO", {"tmpV12": TEST_FUND}),
            ("DISFundProspCmsSO", {"tmpV12": TEST_FUND}),
            ("DISFundCmsSO", {"tmpV12": TEST_FUND}),
        ]

        for svc, params in svc_candidates:
            result = await call_xml_service(page, svc, params=params)
            status = result.get("status")
            text = result.get("text", "")[:200]
            if status == 200 and "<" in text and "error" not in text.lower()[:50]:
                print(f"  ✅ {svc} → {status} | {text[:100]}")
            else:
                print(f"  ❌ {svc} → {status} | {text[:80]}")

        # ── 3. 실제 페이지 네트워크 캡처 ──
        print("\n" + "=" * 60)
        print("3단계: 실제 페이지에서 네트워크 캡처")
        print("=" * 60)

        # 투자설명서 URL 후보 탐색
        print("\n투자설명서 URL 후보 탐색:")
        for url in CANDIDATE_URLS[1:]:
            await test_url(page, url)

        # ── 4. KOFIA 사이트 메뉴 구조 파악 ──
        print("\n" + "=" * 60)
        print("4단계: KOFIA 메인 페이지 메뉴 구조")
        print("=" * 60)
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp", timeout=15000)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(2)

        # 전체 링크/onclick 캡처
        links = await page.evaluate('''() => {
            const items = [];
            document.querySelectorAll("a, li, [onclick]").forEach(el => {
                const text = el.textContent?.trim().replace(/\\s+/g, " ").substring(0, 50);
                const href = el.getAttribute("href") || "";
                const onclick = el.getAttribute("onclick") || "";
                if (text && text.length > 1 && (
                    href.includes("kofia") || href.includes("jsp") ||
                    onclick.includes("go") || onclick.includes("menu") ||
                    onclick.includes("service") ||
                    text.includes("설명서") || text.includes("공시") || text.includes("펀드")
                )) {
                    items.push({text, href: href.substring(0, 200), onclick: onclick.substring(0, 200)});
                }
            });
            return items.slice(0, 50);
        }''')

        print("\n메뉴/링크 목록:")
        for link in links:
            print(f"  [{link['text']}]")
            if link['href']: print(f"    href: {link['href']}")
            if link['onclick']: print(f"    onclick: {link['onclick']}")

        # ── 5. w2xPath 목록 탐색 ──
        print("\n" + "=" * 60)
        print("5단계: w2xPath 패턴 탐색")
        print("=" * 60)

        # JavaScript로 WebSquare 내부 라우팅 정보 추출
        ws_info = await page.evaluate('''() => {
            // WebSquare 설정에서 메뉴 정보 추출 시도
            const result = {};

            // window 전역 객체에서 유용한 정보 찾기
            for (const key of Object.keys(window)) {
                if (key.toLowerCase().includes("menu") || key.toLowerCase().includes("nav")) {
                    try {
                        result[key] = JSON.stringify(window[key]).substring(0, 500);
                    } catch(e) {}
                }
            }

            // 페이지 소스에서 w2xPath 패턴 찾기
            const source = document.documentElement.innerHTML;
            const paths = [];
            const regex = /w2xPath[=\\/\\s"']+([^"'&\\s]+\\.xml)/g;
            let m;
            while ((m = regex.exec(source)) !== null) {
                paths.push(m[1]);
            }
            result._w2xPaths = [...new Set(paths)];

            return result;
        }''')

        print("\nw2xPath 발견:")
        for path in ws_info.get("_w2xPaths", []):
            print(f"  {path}")

        # 결과 저장
        out = Path("output/explore_prospectus.json")
        out.parent.mkdir(exist_ok=True)
        out.write_text(json.dumps({
            "links": links,
            "ws_info": ws_info,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n결과 저장: {out}")

        await browser.close()
        print("\n탐색 완료!")


asyncio.run(main())
