"""
KOFIA 투자설명서 다운로드 스크립트 (3개 테스트)
1단계: 투자설명서 메뉴/API 탐색
2단계: PDF 다운로드
"""
import asyncio
import json
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils
from pathlib import Path
from playwright.async_api import async_playwright

# ── 설정 ──────────────────────────────────────────────────────────────────
TEST_ETFS = [
    {"name": "KODEX 200",           "search": "KODEX 200",            "amc": "삼성"},
    {"name": "TIGER 미국S&P500",    "search": "TIGER 미국S&P500",     "amc": "미래에셋"},
    {"name": "KODEX 미국나스닥100TR", "search": "KODEX 미국나스닥100", "amc": "삼성"},
]

OUT_DIR = Path("output/prospectus")
OUT_DIR.mkdir(parents=True, exist_ok=True)

KOFIA_MAIN = "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000"
XML_SVC    = "/proframeWeb/XMLSERVICES/"

# ── XML 요청 헬퍼 ─────────────────────────────────────────────────────────
def make_xml(svc_name, fn_name, params: dict) -> str:
    param_xml = "\n".join(f"    <{k}>{saxutils.escape(str(v))}</{k}>" for k, v in params.items())
    return f"""<?xml version="1.0" encoding="utf-8"?>
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


async def post_xml(page, svc_name, fn_name, params) -> dict:
    body = make_xml(svc_name, fn_name, params)
    return await page.evaluate('''
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
    ''', [XML_SVC, body])


def parse_xml_rows(xml_text: str, row_tag="selectMeta") -> list[dict]:
    rows = []
    try:
        root = ET.fromstring(xml_text)
        for el in root.iter(row_tag):
            row = {}
            for child in el:
                row[child.tag] = (child.text or "").strip()
            rows.append(row)
    except Exception as e:
        print(f"  XML 파싱 오류: {e}")
    return rows


# ── 투자설명서 서비스 탐색 ────────────────────────────────────────────────
PROSP_SERVICES = [
    # (서비스명, fn명, 파라미터)
    ("DISFundProspSO",      "select", {"tmpV12": "KODEX 200"}),
    ("DISFundProspCmsSO",   "select", {"tmpV12": "KODEX 200"}),
    ("DISFundPubSO",        "select", {"tmpV12": "KODEX 200"}),
    ("DISFundDocSO",        "select", {"tmpV12": "KODEX 200"}),
    ("DISFundAnnSO",        "select", {"tmpV12": "KODEX 200"}),
    ("DISFundFileSO",       "select", {"tmpV12": "KODEX 200"}),
    ("DISFundDiscSO",       "select", {"tmpV12": "KODEX 200"}),
    ("DISFundInvPSO",       "select", {"tmpV12": "KODEX 200"}),
    ("DISFundProspDetailSO","select", {"tmpV12": "KODEX 200"}),
    ("DISFundPublicSO",     "select", {"tmpV12": "KODEX 200"}),
]


async def find_prospectus_service(page):
    """투자설명서 XMLSERVICES API 서비스명 탐색"""
    print("\n[탐색] 투자설명서 API 서비스명 찾는 중...")
    for svc, fn, params in PROSP_SERVICES:
        result = await post_xml(page, svc, fn, params)
        status = result.get("status", 0)
        text = result.get("text", "")
        # 성공 판정: 200이고 XML 구조가 있고 에러가 아님
        if status == 200 and len(text) > 100 and "</" in text:
            err_tags = ["pfmSvcName", "error", "exception", "fail"]
            if not any(t in text.lower() for t in err_tags):
                print(f"  ✅ 발견! 서비스: {svc}")
                print(f"     응답 미리보기: {text[:300]}")
                return svc, text
        print(f"  ❌ {svc} → status={status}, len={len(text)}")
    return None, None


# ── 네트워크 캡처로 투자설명서 URL 탐색 ──────────────────────────────────
async def capture_prospectus_via_ui(page, etf_name: str) -> list[str]:
    """
    KOFIA 투자설명서 페이지를 직접 조작해서 다운로드 URL 캡처
    실제 투자설명서 페이지 URL 후보들을 순차적으로 시도
    """
    pdf_urls = []

    candidate_paths = [
        "/wq/fundann/DISFundProspect.xml",
        "/wq/fundann/DISFundProsp.xml",
        "/wq/fundpub/DISFundProspect.xml",
        "/wq/funddis/DISFundProspect.xml",
        "/wq/fundann/DISFundInvP.xml",
    ]

    for path in candidate_paths:
        url = f"https://dis.kofia.or.kr/websquare/index.jsp?w2xPath={path}"
        captured = []

        async def handle_response(response):
            r_url = response.url
            # PDF 또는 파일 다운로드 응답 캡처
            ct = response.headers.get("content-type", "")
            if "pdf" in ct or "octet" in ct or "download" in r_url.lower():
                captured.append(r_url)
                print(f"    [PDF URL 캡처] {r_url}")

        page.on("response", handle_response)

        try:
            resp = await page.goto(url, timeout=10000)
            await page.wait_for_load_state("domcontentloaded", timeout=10000)
            await asyncio.sleep(1)
            page_text = await page.evaluate("() => document.body?.innerText?.substring(0, 200)")
            if page_text and len(page_text.strip()) > 10 and "blocked" not in page_text.lower():
                print(f"  ✅ 유효한 페이지: {path}")
                print(f"     내용: {page_text[:100]}")
                pdf_urls.extend(captured)
                page.remove_listener("response", handle_response)
                return pdf_urls, path
        except Exception as e:
            pass

        page.remove_listener("response", handle_response)

    return pdf_urls, None


async def main():
    print("=" * 60)
    print("KOFIA 투자설명서 다운로드 테스트 (3개)")
    print("=" * 60)

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,  # 눈으로 확인
            downloads_path=str(OUT_DIR.resolve()),
        )
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            accept_downloads=True,
        )
        page = await context.new_page()

        # ── Step 1: KOFIA 로드 (쿠키/세션 확보) ──
        print("\n[1/3] KOFIA 초기 로드...")
        await page.goto(KOFIA_MAIN)
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)
        print("  ✅ KOFIA 로드 완료")

        # ── Step 2: XMLSERVICES API 서비스명 탐색 ──
        print("\n[2/3] 투자설명서 API 탐색...")
        found_svc, sample_response = await find_prospectus_service(page)

        if found_svc:
            print(f"\n  ✅ API 서비스 발견: {found_svc}")
            # 샘플 파싱
            rows = parse_xml_rows(sample_response)
            print(f"  행 수: {len(rows)}")
            if rows:
                print(f"  첫 행 keys: {list(rows[0].keys())}")
                print(f"  첫 행: {rows[0]}")
        else:
            print("\n  ⚠️  XMLSERVICES에서 투자설명서 서비스 못 찾음")
            print("  → UI 네비게이션 방식으로 시도...")

        # ── Step 3: 투자설명서 페이지 URL 탐색 ──
        print("\n[3/3] 투자설명서 페이지 탐색...")
        pdf_urls, found_path = await capture_prospectus_via_ui(page, "KODEX 200")
        if found_path:
            print(f"  ✅ 투자설명서 페이지: {found_path}")
        else:
            print("  ⚠️  후보 URL 모두 실패")

        # ── Step 4: 전체 페이지 링크에서 '투자설명서' 찾기 ──
        print("\n[추가] 현재 페이지에서 투자설명서 관련 링크 탐색...")
        await page.goto("https://dis.kofia.or.kr/websquare/index.jsp")
        await page.wait_for_load_state("networkidle", timeout=20000)
        await asyncio.sleep(3)

        all_links = await page.evaluate('''() => {
            const result = [];
            // 모든 클릭 가능한 요소 확인
            document.querySelectorAll("a, li, span, div").forEach(el => {
                const text = el.textContent?.trim().replace(/\\s+/g, " ");
                if (text && (
                    text.includes("투자설명서") ||
                    text.includes("간이투자") ||
                    text.includes("핵심상품") ||
                    text.includes("Prospect")
                ) && text.length < 30) {
                    const href = el.getAttribute("href") || "";
                    const onclick = el.getAttribute("onclick") || "";
                    result.push({tag: el.tagName, text, href, onclick: onclick.substring(0, 300)});
                }
            });
            return result;
        }''')

        print(f"\n  '투자설명서' 관련 요소 {len(all_links)}개:")
        for link in all_links:
            print(f"  [{link['tag']}] {link['text']}")
            if link['href']: print(f"    href: {link['href']}")
            if link['onclick']: print(f"    onclick: {link['onclick']}")

        # ── 결과 저장 ──
        result = {
            "found_service": found_svc,
            "found_path": found_path,
            "pdf_urls": pdf_urls,
            "all_links": all_links,
        }
        out_file = OUT_DIR / "explore_result.json"
        out_file.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n결과 저장: {out_file}")

        print("\n" + "=" * 60)
        print("탐색 완료! 브라우저 5초 후 종료...")
        await asyncio.sleep(5)
        await browser.close()


asyncio.run(main())
