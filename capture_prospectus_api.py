"""
KOFIA 투자설명서 API 캡처
- 브라우저가 열리면 직접 '투자설명서' 메뉴를 클릭해주세요
- 모든 API 요청을 자동으로 기록합니다
"""
import asyncio
import json
from pathlib import Path
from playwright.async_api import async_playwright

KOFIA_MAIN = "https://dis.kofia.or.kr/websquare/index.jsp"
OUT_FILE = Path("output/prospectus/captured_api.json")
OUT_FILE.parent.mkdir(parents=True, exist_ok=True)

captured = []

async def main():
    print("=" * 60)
    print("KOFIA 투자설명서 API 캡처")
    print("=" * 60)
    print()
    print("브라우저가 열리면:")
    print("  1) 왼쪽 메뉴에서 '투자설명서' 클릭")
    print("  2) 펀드명 검색창에 'KODEX 200' 입력 후 검색")
    print("  3) 결과 중 하나 클릭")
    print("  4) 다운로드 버튼이 있으면 클릭")
    print()
    print("그 동안 이 창에 API 요청이 실시간으로 출력됩니다.")
    print("완료 후 Enter를 누르면 종료됩니다.")
    print("=" * 60)

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--start-maximized"],
        )
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            accept_downloads=True,
            viewport={"width": 1400, "height": 900},
        )
        page = await context.new_page()

        # ── 모든 요청/응답 캡처 ──
        async def on_request(req):
            url = req.url
            method = req.method
            # kofia.or.kr 요청만 필터
            if "kofia.or.kr" in url and method == "POST":
                body = req.post_data or ""
                entry = {
                    "type": "REQUEST",
                    "url": url,
                    "method": method,
                    "body_preview": body[:500],
                }
                captured.append(entry)
                print(f"\n[POST →] {url}")
                if body:
                    print(f"  BODY: {body[:300]}")

        async def on_response(resp):
            url = resp.url
            status = resp.status
            ct = resp.headers.get("content-type", "")

            # PDF 또는 XMLSERVICES 응답만
            if "kofia.or.kr" in url and (
                "XMLSERVICES" in url or
                "pdf" in ct.lower() or
                "octet" in ct.lower() or
                "download" in url.lower()
            ):
                try:
                    body = await resp.text()
                except:
                    body = "(binary)"

                entry = {
                    "type": "RESPONSE",
                    "url": url,
                    "status": status,
                    "content_type": ct,
                    "body_preview": body[:800],
                }
                captured.append(entry)
                print(f"\n[← {status}] {url}")
                print(f"  Content-Type: {ct}")
                print(f"  Body: {body[:400]}")

        page.on("request", on_request)
        page.on("response", on_response)

        # KOFIA 메인 로드
        print("\nKOFIA 로딩 중...")
        await page.goto(KOFIA_MAIN)
        await page.wait_for_load_state("networkidle", timeout=30000)
        print("✅ 로드 완료! 이제 브라우저에서 투자설명서 메뉴를 클릭해주세요.\n")

        # Enter 입력 대기
        await asyncio.get_event_loop().run_in_executor(None, input, "\n>>> 완료 후 여기서 Enter 누르기: ")

        # 결과 저장
        OUT_FILE.write_text(json.dumps(captured, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n캡처된 요청: {len(captured)}개")
        print(f"결과 저장: {OUT_FILE}")

        # 캡처 요약 출력
        print("\n=== 캡처 요약 ===")
        for i, c in enumerate(captured):
            if c["type"] == "REQUEST":
                print(f"[{i}] POST {c['url']}")
                if "XMLSERVICES" in c['url']:
                    print(f"     {c['body_preview'][:200]}")
            else:
                print(f"[{i}] {c['status']} {c['url']} ({c['content_type']})")

        await browser.close()

asyncio.run(main())
