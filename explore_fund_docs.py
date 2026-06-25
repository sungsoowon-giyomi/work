"""
ETF 투자설명서 소재 탐색
1. DART 검색 (제대로 된 방식으로)
2. KOFIA 펀드공시 (fundinfo.kofia.or.kr)
3. dis.kofia.or.kr 투자설명서 섹션
"""
import asyncio
import json
from pathlib import Path
from playwright.async_api import async_playwright

OUT_DIR = Path("output/prospectus")
OUT_DIR.mkdir(parents=True, exist_ok=True)

TEST_ETF = "KODEX 200"
TEST_AMC = "삼성자산운용"


async def try_dart(page):
    print("\n" + "=" * 50)
    print("1. DART (dart.fss.or.kr)")
    print("=" * 50)

    # DART 메인 페이지 로드
    await page.goto("https://dart.fss.or.kr/dsab001/main.do", timeout=20000)
    await page.wait_for_load_state("networkidle", timeout=15000)
    await asyncio.sleep(2)

    # 검색 폼 요소 파악
    inputs = await page.evaluate('''() => {
        return Array.from(document.querySelectorAll("input, select, textarea")).map(el => ({
            tag: el.tagName, type: el.type, id: el.id,
            name: el.name, placeholder: el.placeholder,
            value: el.value
        }));
    }''')
    print("입력 요소:")
    for inp in inputs:
        print(f"  [{inp['tag']}] id={inp['id']} name={inp['name']} placeholder={inp['placeholder']}")

    # 회사명 입력
    try:
        await page.fill("#textCrpNm", TEST_AMC)
        await asyncio.sleep(0.3)
        # 검색 버튼 클릭
        await page.click("#btnSearch, button[type='submit'], .btnSearch")
        await page.wait_for_load_state("networkidle", timeout=10000)
        await asyncio.sleep(2)
    except Exception as e:
        print(f"검색 오류: {e}")

    # 결과 파싱
    text = await page.evaluate("() => document.body.innerText")
    print(f"검색 결과 미리보기:\n{text[:500]}")

    # 링크 파싱
    links = await page.evaluate('''() => {
        return Array.from(document.querySelectorAll("a")).map(a => ({
            text: a.textContent.trim().substring(0, 60),
            href: a.href
        })).filter(l => l.text.includes("투자설명서") || l.text.includes("KODEX"));
    }''')
    print(f"\n투자설명서 관련 링크: {len(links)}개")
    for l in links[:10]:
        print(f"  [{l['text']}] {l['href'][:100]}")

    # 스크린샷
    await page.screenshot(path=str(OUT_DIR / "dart_search.png"))
    print(f"스크린샷: output/prospectus/dart_search.png")


async def try_fundinfo_kofia(page):
    print("\n" + "=" * 50)
    print("2. KOFIA 펀드다모아 (fundinfo.kofia.or.kr)")
    print("=" * 50)

    await page.goto("https://fundinfo.kofia.or.kr", timeout=20000)
    await page.wait_for_load_state("networkidle", timeout=15000)
    await asyncio.sleep(2)

    title = await page.title()
    text = await page.evaluate("() => document.body.innerText")
    print(f"제목: {title}")
    print(f"내용 미리보기:\n{text[:400]}")

    # 투자설명서 관련 링크 찾기
    links = await page.evaluate('''() => {
        return Array.from(document.querySelectorAll("a, [onclick]")).map(el => ({
            text: el.textContent?.trim().substring(0, 60),
            href: el.getAttribute("href") || "",
            onclick: (el.getAttribute("onclick") || "").substring(0, 100)
        })).filter(l => l.text && (
            l.text.includes("투자설명서") || l.text.includes("공시") ||
            l.text.includes("펀드정보") || l.text.includes("서류")
        ));
    }''')
    print(f"\n관련 링크 {len(links)}개:")
    for l in links[:15]:
        print(f"  [{l['text']}] href={l['href'][:60]} onclick={l['onclick'][:60]}")

    await page.screenshot(path=str(OUT_DIR / "fundinfo_main.png"))
    print(f"스크린샷: output/prospectus/fundinfo_main.png")


async def try_dis_kofia_menu(page):
    print("\n" + "=" * 50)
    print("3. KOFIA DIS (dis.kofia.or.kr) 전체 메뉴 탐색")
    print("=" * 50)

    await page.goto("https://dis.kofia.or.kr/websquare/index.jsp", timeout=20000)
    await page.wait_for_load_state("networkidle", timeout=30000)
    await asyncio.sleep(5)  # 충분히 대기

    # 전체 HTML에서 w2xPath 패턴 추출
    all_paths = await page.evaluate('''() => {
        const src = document.documentElement.innerHTML;
        const found = new Set();
        // w2xPath 패턴
        const r1 = /w2xPath[=\\/'"]+([^\\/'"&\\s,)]+\\.xml)/g;
        let m;
        while ((m = r1.exec(src)) !== null) found.add(m[1]);
        // serviceId 패턴
        const r2 = /serviceId[='"]+([A-Z0-9]+)/g;
        while ((m = r2.exec(src)) !== null) found.add("SVC:" + m[1]);
        // divisionId 패턴
        const r3 = /divisionId[='"]+([A-Z0-9]+)/g;
        while ((m = r3.exec(src)) !== null) found.add("DIV:" + m[1]);
        return [...found];
    }''')
    print(f"\nHTML에서 발견된 경로/ID ({len(all_paths)}개):")
    for p in all_paths:
        print(f"  {p}")

    # 모든 iframe 확인
    frames = page.frames
    print(f"\niframe 수: {len(frames)}")
    for i, frame in enumerate(frames):
        try:
            furl = frame.url
            ftitle = await frame.title()
            print(f"  [Frame {i}] url={furl} title={ftitle}")

            # 투자설명서 관련 텍스트 찾기
            links = await frame.evaluate('''() => {
                return Array.from(document.querySelectorAll("a, li, span, [onclick]"))
                    .filter(el => {
                        const t = el.textContent?.trim();
                        return t && (t.includes("설명서") || t.includes("투자") || t.includes("공시")) && t.length < 20;
                    })
                    .map(el => ({
                        text: el.textContent.trim(),
                        href: el.getAttribute("href") || "",
                        onclick: (el.getAttribute("onclick") || "").substring(0, 150)
                    }));
            }''')
            if links:
                print(f"    관련 요소 {len(links)}개:")
                for l in links[:10]:
                    print(f"      [{l['text']}] onclick={l['onclick'][:80]}")
        except Exception as e:
            print(f"  [Frame {i}] 오류: {e}")

    await page.screenshot(path=str(OUT_DIR / "dis_kofia_main.png"))
    print(f"\n스크린샷: output/prospectus/dis_kofia_main.png")


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            viewport={"width": 1400, "height": 900},
        )
        page = await context.new_page()

        await try_dart(page)
        await try_fundinfo_kofia(page)
        await try_dis_kofia_menu(page)

        await browser.close()
        print("\n\n탐색 완료!")


asyncio.run(main())
