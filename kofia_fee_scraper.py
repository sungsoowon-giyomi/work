"""
KOFIA DISFundFeeCMS API 스크래퍼
- dis.kofia.or.kr proframeWeb/XMLSERVICES/ 직접 호출
- 31개 펀드/ETF 합성총보수 & 증권거래비용 추출
- 전체 클래스 보이는 테이블 스크린샷 캡쳐
"""
import asyncio
import json
import re
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils
from pathlib import Path
from datetime import datetime
from playwright.async_api import async_playwright

OUTPUT_DIR = Path(r"C:\Users\jswon\Desktop\업무\funetf_scraper\output")
IMG_DIR = OUTPUT_DIR / "images"
OUTPUT_DIR.mkdir(exist_ok=True)
IMG_DIR.mkdir(exist_ok=True)

# ── 31개 항목 목록 (AMC 검색어 및 클래스 정보 추가) ──
ITEMS = [
    # ETF (company code = "" for all - use fund name search only)
    {"name": "KODEX 미국S&P500",              "type":"etf","sotCd":"379800","search":"KODEX 미국S&P500",        "amc":"삼성자산운용","class":None},
    {"name": "TIME 글로벌AI인공지능액티브",      "type":"etf","sotCd":"456600","search":"TIME 글로벌AI인공지능",    "amc":"파인아시아자산운용","class":None},
    {"name": "KODEX AI전력핵심설비",            "type":"etf","sotCd":"487240","search":"KODEX AI전력핵심설비",     "amc":"삼성자산운용","class":None},
    {"name": "ACE 글로벌반도체TOP4 Plus",       "type":"etf","sotCd":"446770","search":"ACE 글로벌반도체TOP4",    "amc":"한국투자신탁운용","class":None},
    {"name": "RISE 대형고배당10TR",             "type":"etf","sotCd":"315960","search":"RISE 대형고배당",        "amc":"KB자산운용","class":None},
    {"name": "KODEX 머니마켓액티브",             "type":"etf","sotCd":"488770","search":"KODEX 머니마켓액티브",    "amc":"삼성자산운용","class":None},
    {"name": "RISE 미국단기투자등급회사채액티브",  "type":"etf","sotCd":"437350","search":"RISE 미국단기투자등급회사채","amc":"KB자산운용","class":None},
    {"name": "ACE KRX금현물",                  "type":"etf","sotCd":"411060","search":"ACE KRX금현물",          "amc":"한국투자신탁운용","class":None},
    {"name": "KODEX 미국부동산리츠(H)",          "type":"etf","sotCd":"352560","search":"KODEX 미국부동산리츠",    "amc":"삼성자산운용","class":None},
    {"name": "TIGER 글로벌멀티에셋TIF액티브",    "type":"etf","sotCd":"440340","search":"TIGER 글로벌멀티에셋TIF", "amc":"미래에셋자산운용","class":None},
    {"name": "KODEX TRF3070",                  "type":"etf","sotCd":"329650","search":"KODEX TRF3070",          "amc":"삼성자산운용","class":None},
    # 퇴직연금 펀드 (company code = "" - search by fund name only)
    {"name": "삼성미국S&P500인덱스증권자투자신탁UH[주식]Cp(퇴직연금)",   "type":"fund","code":"PAS025","search":"삼성미국S&P500인덱스증권자투자신탁","amc":"삼성자산운용","class":"C-P"},
    {"name": "미래에셋미국나스닥100인덱스증권자투자신탁(주식)(UH)C-P2",   "type":"fund","code":"PAM078","search":"미래에셋미국나스닥100인덱스","amc":"미래에셋자산운용","class":"C-P2"},
    {"name": "KCGI코리아퇴직연금증권자투자신탁[주식]C",                  "type":"fund","code":"PAMJ03","search":"KCGI코리아퇴직연금증권자투자신탁","amc":"KCGI자산운용","class":"C"},
    {"name": "NH-Amundi 하나로단기채증권투자신탁[채권]C-P2(퇴직연금)",    "type":"fund","code":"PDNC74","search":"하나로단기채증권투자신탁","amc":"NH-Amundi자산운용","class":"C-P2"},
    {"name": "신한MAN글로벌투자등급채권증권투자신탁(H)[채권-재간접형]C-r",  "type":"fund","code":"PGJA46","search":"신한MAN글로벌투자등급채권","amc":"신한자산운용","class":"C-r"},
    {"name": "신한골드증권투자신탁제1호[주식]C-r",                       "type":"fund","code":"PAJA24","search":"신한골드증권투자신탁","amc":"신한자산운용","class":"C-r"},
    {"name": "하나글로벌리츠부동산자투자신탁[재간접형]C-P2",               "type":"fund","code":"PGD862","search":"하나글로벌리츠부동산자투자신탁","amc":"하나자산운용","class":"C-P2"},
    {"name": "NH-Amundi하나로적격TDF2040증권투자[주식혼합-재간접형]C-P2",  "type":"fund","code":"PGNC86","search":"하나로적격TDF2040증권투자신탁","amc":"NH-Amundi자산운용","class":"C-P2"},
    {"name": "KB온국민적격TDF2040증권자투자[주식혼합-재간접형](H)C-퇴직",  "type":"fund","code":"PGZA34","search":"KB온국민적격TDF2040","amc":"KB자산운용","class":"C-퇴직"},
    {"name": "미래에셋전략배분적격TDF2040혼합자산자투자신탁C-P2",          "type":"fund","code":"PBM573","search":"미래에셋전략배분적격TDF2040","amc":"미래에셋자산운용","class":"C-P2"},
    # 연금저축 펀드
    {"name": "삼성미국S&P500인덱스증권자투자신탁UH[주식]C-Pe",            "type":"fund","code":"NAQP03","search":"삼성미국S&P500인덱스증권자투자신탁","amc":"삼성자산운용","class":"C-Pe"},
    {"name": "미래에셋미국나스닥100인덱스증권자투자신탁(주식)(UH)C-Pe",     "type":"fund","code":"NAM079","search":"미래에셋미국나스닥100인덱스","amc":"미래에셋자산운용","class":"C-Pe"},
    {"name": "KCGI코리아증권투자신탁1호[주식]C-Pe",                      "type":"fund","code":"NAMP94","search":"KCGI코리아증권투자신탁","amc":"KCGI자산운용","class":"C-Pe"},
    {"name": "NH-Amundi 하나로단기채증권투자신탁[채권]C-P1e(연금저축)",    "type":"fund","code":"NDNP88","search":"하나로단기채증권투자신탁","amc":"NH-Amundi자산운용","class":"C-P1e"},
    {"name": "신한MAN글로벌투자등급채권증권투자신탁(H)[채권-재간접형]C-pe",  "type":"fund","code":"NGJA45","search":"신한MAN글로벌투자등급채권","amc":"신한자산운용","class":"C-pe"},
    {"name": "신한골드증권투자신탁제1호[주식]C-pe",                       "type":"fund","code":"NASJ83","search":"신한골드증권투자신탁","amc":"신한자산운용","class":"C-pe"},
    {"name": "하나글로벌리츠부동산자투자신탁[재간접형]C-PE",                "type":"fund","code":"NGDP20","search":"하나글로벌리츠부동산자투자신탁","amc":"하나자산운용","class":"C-PE"},
    {"name": "NH-Amundi하나로적격TDF2040증권투자[주식혼합-재간접형]C-P1e", "type":"fund","code":"NGUP02","search":"하나로적격TDF2040증권투자신탁","amc":"NH-Amundi자산운용","class":"C-P1e"},
    {"name": "KB온국민적격TDF2040증권자투자[주식혼합-재간접형](H)C-Pe",    "type":"fund","code":"NGZA18","search":"KB온국민적격TDF2040","amc":"KB자산운용","class":"C-Pe"},
    {"name": "미래에셋전략배분적격TDF2040혼합자산자투자신탁C-Pe",           "type":"fund","code":"NBMP57","search":"미래에셋전략배분적격TDF2040","amc":"미래에셋자산운용","class":"C-Pe"},
]

KOFIA_URL = "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000"
API_URL = "/proframeWeb/XMLSERVICES/"
QUERY_DATE = "20260430"  # 기준일 (최근 월말)


def make_xml_body(company_code: str, fund_name: str) -> str:
    # XML 특수문자 이스케이프 필수 (&, <, >, ", ')
    safe_name = saxutils.escape(fund_name)
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
    <tmpV12>{safe_name}</tmpV12>
    <tmpV3></tmpV3>
    <tmpV5></tmpV5>
    <tmpV4></tmpV4>
</DISCondFuncDTO>
</message>"""


def parse_fee_response(xml_text: str) -> list:
    """XML 응답에서 모든 펀드 클래스 데이터 파싱"""
    rows = []
    try:
        root = ET.fromstring(xml_text)
        for meta in root.iter("selectMeta"):
            def v(tag):
                el = meta.find(tag)
                return (el.text or "").strip() if el is not None else ""
            rows.append({
                "amc": v("tmpV1"),
                "fund_class_name": v("tmpV2"),
                "fund_type": v("tmpV3"),
                "set_date": v("tmpV4"),
                "mgmt_fee": v("tmpV5"),
                "dist_fee": v("tmpV6"),
                "trust_fee": v("tmpV7"),
                "admin_fee": v("tmpV8"),
                "total_fee": v("tmpV9"),
                "ter": v("tmpV10"),
                "trading_cost": v("tmpV11"),
                "synthetic_fee": v("tmpV12"),
                "sales_commission": v("tmpV13"),
                "std_code": v("tmpV15"),
                "actual_trading_cost": v("tmpV16"),
            })
    except Exception as e:
        print(f"    XML 파싱 오류: {e}")
    return rows


def find_matching_class(rows: list, class_suffix: str, display_name: str = "") -> dict | None:
    """클래스 접미사로 매칭 행 찾기"""
    if not rows:
        return None
    if class_suffix is None:
        # ETF: 여러 행 중 display_name과 가장 유사한 것 (이름 가장 짧은 것)
        if len(rows) == 1:
            return rows[0]
        # 이름 길이 순 정렬 (짧을수록 순수 ETF, 긴 건 섹터별 변형)
        sorted_rows = sorted(rows, key=lambda r: len(r["fund_class_name"]))
        return sorted_rows[0]

    # 정규화: 소문자, 공백/하이픈 유지
    target = class_suffix.strip().lower()

    # 1차: 정확히 "종류{class_suffix}" 포함
    for row in rows:
        name = row["fund_class_name"].lower()
        if f"종류{target}" in name:
            return row

    # 2차: 이름 끝에 class_suffix 포함 (대소문자 무관)
    for row in rows:
        name = row["fund_class_name"].lower()
        # 하이픈 정규화
        name_norm = name.replace("-", "").replace(" ", "")
        target_norm = target.replace("-", "").replace(" ", "")
        if name_norm.endswith(target_norm):
            return row

    # 3차: 단순 포함
    for row in rows:
        name = row["fund_class_name"].lower()
        if target.replace("-", "") in name.replace("-", ""):
            return row

    print(f"    ⚠️ 클래스 '{class_suffix}' 미매칭. 첫 행 반환")
    return rows[0]


async def call_kofia_api(page, company_code: str, fund_search: str) -> list:
    """page.evaluate로 브라우저 컨텍스트에서 API 호출"""
    xml_body = make_xml_body(company_code, fund_search)
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
                const text = await r.text();
                return {status: r.status, text: text};
            } catch(e) {
                return {status: 0, text: "", error: e.toString()};
            }
        }
    ''', [API_URL, xml_body])

    if result.get("status") != 200 or not result.get("text"):
        print(f"    API 오류: {result}")
        return []

    rows = parse_fee_response(result["text"])
    return rows


async def screenshot_kofia_table(page, fund_search: str, company_code: str, out_path: Path) -> bool:
    """KOFIA 검색 결과 테이블 스크린샷"""
    try:
        # 이미 KOFIA 페이지에 있으므로, 검색만 다시 실행
        # 회사 선택
        if company_code:
            try:
                sel = page.locator('#compCd_input_0')
                if await sel.count() > 0:
                    await sel.select_option(company_code)
            except:
                pass

        # 펀드명 입력
        for fld in ['#fundNm', '#dSearchFndNm']:
            try:
                el = page.locator(fld).first
                if await el.count() > 0:
                    await el.fill(fund_search)
                    break
            except:
                pass

        await page.keyboard.press('Enter')
        await asyncio.sleep(3)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await asyncio.sleep(1)

        # 테이블/그리드 요소 찾아 스크린샷
        # WebSquare 그리드 - 보통 div[id*="grid"] 또는 table
        grid_selectors = [
            '.w2grid',
            'div[id*="vGrid"]',
            'div[id*="Grid"]',
            'table.w2grid_table',
            '#vGridDiv1',
            'div.w2gridContainer',
        ]

        for sel in grid_selectors:
            try:
                el = page.locator(sel).first
                if await el.count() > 0:
                    bbox = await el.bounding_box()
                    if bbox and bbox['height'] > 50:
                        await el.screenshot(path=str(out_path))
                        print(f"    ✅ 스크린샷: {out_path.name} ({sel})")
                        return True
            except:
                pass

        # 전체 페이지 스크린샷 (fallback)
        await page.screenshot(path=str(out_path), full_page=False)
        print(f"    ⚠️ 전체 페이지 스크린샷: {out_path.name}")
        return True

    except Exception as e:
        print(f"    스크린샷 오류: {e}")
        return False


async def main():
    # 기존 결과 로드
    results_files = sorted(OUTPUT_DIR.glob("results_*.json"), reverse=True)
    existing_results = {}
    if results_files:
        with open(results_files[0], "r", encoding="utf-8") as f:
            data = json.load(f)
            for item in data:
                existing_results[item.get("name", "")] = item
        print(f"✅ 기존 결과 로드: {results_files[0].name} ({len(existing_results)}개)")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        # ── KOFIA 초기 로드 및 회사 코드 수집 ──
        print("=" * 60)
        print("KOFIA DISFundFeeCMS 초기 로드 & 회사 코드 수집")
        print("=" * 60)

        await page.goto(KOFIA_URL)
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(3)

        # 회사 드롭다운에서 코드 추출
        company_options = {}
        try:
            opts = await page.evaluate('''
                () => Array.from(document.querySelector('#compCd_input_0').options)
                    .map(o => ({value: o.value, text: o.text.trim()}))
                    .filter(o => o.value)
            ''')
            for opt in opts:
                company_options[opt['text']] = opt['value']
            print(f"  회사 {len(company_options)}개 로드")
        except Exception as e:
            print(f"  회사 드롭다운 오류: {e}")

        # 회사 코드: WebSquare 드롭다운 value는 텍스트 이름으로 서버가 인식 못 함
        # → 모든 검색에서 회사 코드 빈 값(전체 검색)으로 통일
        print(f"  → 모든 검색은 회사 코드 없이 펀드명으로만 검색")

        # ── 각 항목 처리 ──
        results = []

        for idx, item in enumerate(ITEMS):
            name = item["name"]
            fund_search = item["search"]
            amc = item["amc"]
            class_suffix = item["class"]
            company_code = ""  # 빈 값: 전체 회사 검색

            print(f"\n[{idx+1:02d}/{len(ITEMS)}] {name[:50]}")
            print(f"  검색: '{fund_search}' | 회사코드: '{company_code}' | 클래스: '{class_suffix}'")

            # 기존 결과에서 기본 데이터 가져오기
            existing = existing_results.get(name, {})

            result = {
                "name": name,
                "type": item["type"],
                "listing_date": existing.get("listing_date"),
                "amc": existing.get("amc") or amc,
                "risk_grade": existing.get("risk_grade"),
                "total_fee": existing.get("total_fee"),
                "trading_cost": None,
                "synthetic_fee": None,
                "fee_image_path": existing.get("fee_image_path"),
                "kofia_rows": [],
            }

            # API 호출
            rows = await call_kofia_api(page, company_code, fund_search)
            print(f"  API 결과: {len(rows)}행")

            if rows:
                result["kofia_rows"] = [r["fund_class_name"] for r in rows]

                # 대상 클래스 찾기
                matched = find_matching_class(rows, class_suffix, name)

                if matched:
                    result["synthetic_fee"] = matched["synthetic_fee"]
                    result["trading_cost"] = matched["trading_cost"]
                    result["total_fee"] = matched["total_fee"]
                    result["set_date_kofia"] = matched["set_date"]
                    result["std_code"] = matched["std_code"]
                    print(f"  ✅ 매칭: {matched['fund_class_name']}")
                    print(f"     합성총보수={matched['synthetic_fee']}%, 증권거래비용={matched['trading_cost']}%, 총보수={matched['total_fee']}%")
                else:
                    print(f"  ❌ 클래스 매칭 실패")
            else:
                print(f"  ❌ API 결과 없음")

            # 스크린샷 (KOFIA 테이블)
            safe_name = re.sub(r'[\\/*?:"<>|]', '_', name[:40])
            img_path = IMG_DIR / f"kofia_{safe_name}.png"

            if rows:  # 검색 결과가 있을 때만 스크린샷
                await screenshot_kofia_table(page, fund_search, company_code, img_path)
                if img_path.exists():
                    result["fee_image_path"] = str(img_path)

            # KOFIA 페이지 재로드 (다음 검색을 위해)
            await page.goto(KOFIA_URL)
            await page.wait_for_load_state("networkidle", timeout=20000)
            await asyncio.sleep(2)

            results.append(result)

        await browser.close()

    # ── 결과 저장 ──
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    out_file = OUTPUT_DIR / f"kofia_results_{ts}.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\n✅ 결과 저장: {out_file}")

    # ── 결과 요약 ──
    print("\n" + "=" * 60)
    print("결과 요약")
    print("=" * 60)
    has_fee = sum(1 for r in results if r.get("synthetic_fee"))
    has_trading = sum(1 for r in results if r.get("trading_cost"))
    has_img = sum(1 for r in results if r.get("fee_image_path"))
    print(f"  합성총보수 확보: {has_fee}/{len(results)}")
    print(f"  증권거래비용 확보: {has_trading}/{len(results)}")
    print(f"  이미지 확보: {has_img}/{len(results)}")
    print()
    for r in results:
        fee = r.get("synthetic_fee", "-")
        tc = r.get("trading_cost", "-")
        print(f"  {r['name'][:45]:<45} 합성총보수={fee}% 거래비용={tc}%")


if __name__ == "__main__":
    asyncio.run(main())
