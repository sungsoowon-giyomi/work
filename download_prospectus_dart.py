"""
DART에서 ETF 투자설명서 PDF 다운로드 + 합성총보수/증권거래비용 추출
"""
import asyncio
import json
import re
from pathlib import Path
import pdfplumber
from playwright.async_api import async_playwright

# ── 테스트 ETF 3개 ──────────────────────────────────────────────────────────
TEST_ETFS = [
    {"name": "KODEX 200",             "amc": "삼성자산운용", "keyword": "KODEX200"},
    {"name": "TIGER 미국S&P500",      "amc": "미래에셋자산운용", "keyword": "TIGER미국S&P500"},
    {"name": "KODEX 미국나스닥100TR", "amc": "삼성자산운용", "keyword": "KODEX미국나스닥100TR"},
]

OUT_DIR = Path("output/prospectus")
OUT_DIR.mkdir(parents=True, exist_ok=True)


# ── DART 검색 ───────────────────────────────────────────────────────────────
async def search_dart(page, amc: str, etf_name: str, keyword: str) -> list[dict]:
    """DART 회사별 검색으로 투자설명서 목록 가져오기"""
    print(f"\n  DART 검색: {amc} / {etf_name}")

    # DART 메인 페이지 로드 후 폼 제출 (search.ax는 AJAX 전용이라 직접 접근 불가)
    await page.goto("https://dart.fss.or.kr/dsab001/main.do", timeout=20000)
    await page.wait_for_load_state("networkidle", timeout=15000)
    await asyncio.sleep(1)

    # 회사명 입력 (메인 폼의 textCrpNm - visible 요소)
    await page.fill("#textCrpNm", amc)
    await asyncio.sleep(0.5)

    # 폼 제출 (검색 버튼 클릭 또는 submit)
    submitted = False
    for selector in ["#btnSearch", "button.btnSearch", "input[type=submit]", ".btn_search"]:
        try:
            await page.click(selector, timeout=2000)
            submitted = True
            break
        except:
            pass
    if not submitted:
        await page.evaluate("document.forms[0] && document.forms[0].submit()")

    await page.wait_for_load_state("networkidle", timeout=15000)
    await asyncio.sleep(2)

    # 결과 파싱 - 모든 링크에서 투자설명서 찾기
    rows = await page.evaluate('''() => {
        const items = [];
        document.querySelectorAll("a").forEach(a => {
            const text = a.textContent.trim().replace(/\\s+/g, " ");
            const href = a.href || "";
            if (!href.includes("dsaf001") && !href.includes("rcpNo")) return;
            if (!text) return;
            const row = a.closest("tr");
            const date = row ? (row.querySelector("td:last-child")?.textContent.trim() || "") : "";
            items.push({title: text, href, date});
        });
        return items;
    }''')

    # 투자설명서만 필터
    rows = [r for r in rows if "투자설명서" in r["title"] or "설명서" in r["title"]]
    print(f"  전체 결과: {len(rows)}개")

    # ETF 이름 매칭 (공백 제거해서 비교)
    kw_clean = keyword.replace(" ", "").lower()
    matched = []
    for r in rows:
        title_clean = r["title"].replace(" ", "").lower()
        if kw_clean in title_clean:
            matched.append(r)

    if matched:
        print(f"  ✅ '{etf_name}' 매칭: {len(matched)}개")
        for m in matched[:3]:
            print(f"    [{m['date']}] {m['title'][:70]}")
    else:
        print(f"  ⚠️  '{etf_name}' 매칭 없음, 전체 목록 상위 5개:")
        for r in rows[:5]:
            print(f"    [{r['date']}] {r['title'][:70]}")

    return matched if matched else rows[:3]


# ── DART 뷰어 → PDF URL 추출 ────────────────────────────────────────────────
async def get_pdf_url_from_viewer(page, rcp_no: str) -> str | None:
    """DART 뷰어 페이지에서 PDF 다운로드 URL 추출"""
    viewer_url = f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcp_no}"
    print(f"  뷰어 로드: {viewer_url}")

    await page.goto(viewer_url, timeout=20000)
    await page.wait_for_load_state("networkidle", timeout=15000)
    await asyncio.sleep(2)

    # 첨부파일 목록에서 PDF 찾기
    attach_info = await page.evaluate('''() => {
        const items = [];
        // DART 첨부파일 테이블
        document.querySelectorAll("#listContents tr, table tr").forEach(row => {
            const link = row.querySelector("a");
            if (!link) return;
            const text = link.textContent.trim();
            const href = link.href || link.getAttribute("href") || "";
            const onclick = link.getAttribute("onclick") || "";
            items.push({text, href: href.substring(0,300), onclick: onclick.substring(0,300)});
        });
        return items;
    }''')

    print(f"  첨부파일 목록 ({len(attach_info)}개):")
    for a in attach_info[:10]:
        print(f"    [{a['text'][:50]}] href={a['href'][:80]}")
        if a['onclick']:
            print(f"              onclick={a['onclick'][:80]}")

    # PDF URL 패턴 찾기
    for item in attach_info:
        href = item["href"]
        onclick = item["onclick"]
        text = item["text"]

        # PDF 다운로드 URL
        if "pdf" in href.lower() or "download" in href.lower():
            return href

        # onclick에서 URL 추출
        if "pdf" in onclick.lower() or "download" in onclick.lower():
            url_match = re.search(r"['\"]([^'\"]*(?:pdf|download)[^'\"]*)['\"]", onclick, re.I)
            if url_match:
                url = url_match.group(1)
                if url.startswith("/"):
                    url = "https://dart.fss.or.kr" + url
                return url

    # 페이지 전체 소스에서 PDF URL 패턴 찾기
    page_source = await page.content()
    pdf_patterns = [
        r'https://dart\.fss\.or\.kr[^"\'<>\s]*\.pdf',
        r'/pdf/download[^"\'<>\s]*',
        r'downloadPdf[^"\'<>\s]*',
    ]
    for pattern in pdf_patterns:
        matches = re.findall(pattern, page_source, re.I)
        if matches:
            url = matches[0]
            if url.startswith("/"):
                url = "https://dart.fss.or.kr" + url
            print(f"  소스에서 PDF URL 발견: {url[:100]}")
            return url

    # iframe 내부 확인
    for frame in page.frames:
        try:
            furl = frame.url
            if furl and furl != viewer_url:
                print(f"  iframe: {furl[:100]}")
                if "pdf" in furl.lower():
                    return furl
        except:
            pass

    return None


# ── PDF 직접 다운로드 ────────────────────────────────────────────────────────
async def download_pdf(page, context, rcp_no: str, save_path: Path) -> bool:
    """DART에서 PDF 다운로드"""

    # 방법 1: 뷰어에서 PDF URL 추출
    pdf_url = await get_pdf_url_from_viewer(page, rcp_no)

    if pdf_url:
        try:
            print(f"  다운로드: {pdf_url[:80]}")
            response = await page.request.get(pdf_url)
            if response.ok:
                content = await response.body()
                save_path.write_bytes(content)
                print(f"  ✅ 저장: {save_path.name} ({len(content):,} bytes)")
                return True
        except Exception as e:
            print(f"  직접 다운로드 실패: {e}")

    # 방법 2: 클릭으로 다운로드
    viewer_url = f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcp_no}"
    await page.goto(viewer_url, timeout=20000)
    await page.wait_for_load_state("networkidle", timeout=15000)
    await asyncio.sleep(2)

    try:
        # PDF 또는 다운로드 버튼 클릭
        async with context.expect_download(timeout=20000) as dl_info:
            # 첨부파일 중 PDF 클릭
            clicked = await page.evaluate('''() => {
                const links = document.querySelectorAll("a");
                for (const a of links) {
                    const text = a.textContent.trim().toLowerCase();
                    const href = (a.href || "").toLowerCase();
                    if (text.includes("pdf") || href.includes("pdf") ||
                        text.includes("투자설명서") || text.includes("다운")) {
                        a.click();
                        return true;
                    }
                }
                return false;
            }''')
        if clicked:
            dl = await dl_info.value
            await dl.save_as(str(save_path))
            print(f"  ✅ 클릭 다운로드 성공: {save_path.name}")
            return True
    except Exception as e:
        print(f"  클릭 다운로드 실패: {e}")

    # 방법 3: DART 뷰어 인쇄/PDF 저장 URL 패턴 시도
    # rcpNo 로 dcmNo 찾기
    try:
        dcm_info = await page.evaluate('''() => {
            const src = document.documentElement.innerHTML;
            const dcms = [];
            const r = /dcmNo[='"\\s]+(\\d+)/g;
            let m;
            while ((m = r.exec(src)) !== null) dcms.push(m[1]);
            return [...new Set(dcms)];
        }''')
        print(f"  dcmNo 목록: {dcm_info}")

        for dcm_no in dcm_info[:3]:
            pdf_url = f"https://dart.fss.or.kr/pdf/download/main.do?rcp_no={rcp_no}&dcm_no={dcm_no}&lang=ko"
            print(f"  PDF URL 시도: {pdf_url}")
            resp = await page.request.get(pdf_url)
            if resp.ok and len(await resp.body()) > 1000:
                content = await resp.body()
                save_path.write_bytes(content)
                print(f"  ✅ 저장: {save_path.name} ({len(content):,} bytes)")
                return True
    except Exception as e:
        print(f"  dcmNo 방식 실패: {e}")

    return False


# ── PDF에서 보수 추출 ────────────────────────────────────────────────────────
def extract_fees_from_pdf(pdf_path: Path) -> dict:
    """PDF에서 합성총보수/증권거래비용 추출"""
    result = {"synthetic_fee": None, "trading_cost": None, "raw": []}
    if not pdf_path.exists():
        return result

    print(f"\n  PDF 분석: {pdf_path.name} ({pdf_path.stat().st_size:,} bytes)")
    fee_pattern = re.compile(r'(\d{1,2}\.\d{2,4})')

    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"  총 {len(pdf.pages)}페이지")
            for pg_num, pg in enumerate(pdf.pages, 1):
                text = pg.extract_text() or ""
                if not any(kw in text for kw in ["합성총보수", "증권거래비용", "합성 총보수"]):
                    continue

                print(f"  [p.{pg_num}] 보수 정보 발견!")
                for line in text.split("\n"):
                    line = line.strip()
                    if not line:
                        continue
                    if "합성총보수" in line or "합성 총보수" in line:
                        nums = fee_pattern.findall(line)
                        print(f"    합성총보수 줄: {line[:100]}")
                        print(f"    숫자: {nums}")
                        result["raw"].append({"type": "합성총보수", "line": line, "nums": nums})
                        if nums and result["synthetic_fee"] is None:
                            result["synthetic_fee"] = nums[-1]
                    if "증권거래비용" in line:
                        nums = fee_pattern.findall(line)
                        print(f"    증권거래비용 줄: {line[:100]}")
                        print(f"    숫자: {nums}")
                        result["raw"].append({"type": "증권거래비용", "line": line, "nums": nums})
                        if nums and result["trading_cost"] is None:
                            result["trading_cost"] = nums[-1]

                # 테이블도 확인
                for table in pg.extract_tables():
                    for row in (table or []):
                        row_str = " | ".join(str(c or "").strip() for c in row)
                        if "합성총보수" in row_str or "증권거래비용" in row_str:
                            print(f"    테이블 행: {row_str[:120]}")
                            result["raw"].append({"type": "table", "row": row_str})

    except Exception as e:
        print(f"  파싱 오류: {e}")

    print(f"\n  ▶ 합성총보수:   {result['synthetic_fee']}%")
    print(f"  ▶ 증권거래비용: {result['trading_cost']}%")
    return result


# ── 메인 ────────────────────────────────────────────────────────────────────
async def main():
    print("=" * 60)
    print("DART 투자설명서 다운로드 + 보수 추출 (3개 테스트)")
    print("=" * 60)

    all_results = {}

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            accept_downloads=True,
        )
        page = await context.new_page()

        for etf in TEST_ETFS:
            name = etf["name"]
            safe = name.replace(" ", "_")
            pdf_path = OUT_DIR / f"{safe}_투자설명서.pdf"

            print(f"\n{'='*50}")
            print(f"▶ {name}")
            print(f"{'='*50}")

            # 1. DART 검색
            docs = await search_dart(page, etf["amc"], name, etf["keyword"])
            if not docs:
                all_results[name] = {"error": "검색 결과 없음"}
                continue

            # 2. PDF 다운로드 (첫 번째 매칭 문서)
            rcp_match = re.search(r'rcpNo=(\d+)', docs[0]["href"])
            if not rcp_match:
                all_results[name] = {"error": "rcpNo 없음"}
                continue

            rcp_no = rcp_match.group(1)
            print(f"\n  rcpNo: {rcp_no}")
            print(f"  문서: {docs[0]['title'][:60]}")

            success = await download_pdf(page, context, rcp_no, pdf_path)
            if not success:
                all_results[name] = {"error": "PDF 다운로드 실패"}
                continue

            # 3. 보수 추출
            fees = extract_fees_from_pdf(pdf_path)
            all_results[name] = fees

        await browser.close()

    # 최종 결과
    print("\n" + "=" * 60)
    print("최종 결과")
    print("=" * 60)
    for name, r in all_results.items():
        print(f"\n{name}:")
        if "error" in r:
            print(f"  ❌ {r['error']}")
        else:
            print(f"  합성총보수:    {r.get('synthetic_fee', 'N/A')}%")
            print(f"  증권거래비용:  {r.get('trading_cost', 'N/A')}%")

    out = OUT_DIR / "dart_fee_results.json"
    out.write_text(json.dumps(all_results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n저장: {out}")


asyncio.run(main())
