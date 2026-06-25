@echo off
chcp 65001 > nul
echo ========================================
echo  funetf_scraper 집 노트북 세팅
echo ========================================
echo.

:: Python 확인
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo https://www.python.org 에서 Python 3.11 이상 설치 후 다시 실행하세요.
    pause
    exit /b 1
)
echo [OK] Python 확인됨
python --version

:: Node.js 확인
node --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Node.js가 설치되어 있지 않습니다.
    echo https://nodejs.org 에서 Node.js LTS 설치 후 다시 실행하세요.
    pause
    exit /b 1
)
echo [OK] Node.js 확인됨
node --version

echo.
echo [1/3] Python 패키지 설치 중...
pip install playwright pdfplumber pymupdf pillow python-pptx requests
if errorlevel 1 (
    echo [오류] pip install 실패
    pause
    exit /b 1
)

echo.
echo [2/3] Playwright 브라우저 설치 중...
playwright install chromium
if errorlevel 1 (
    echo [오류] playwright install 실패
    pause
    exit /b 1
)

echo.
echo [3/3] Claude Code 설치 중...
npm install -g @anthropic/claude-code
if errorlevel 1 (
    echo [오류] Claude Code 설치 실패
    pause
    exit /b 1
)

echo.
echo ========================================
echo  세팅 완료!
echo ========================================
echo.
echo 이제 이 폴더에서 아래 명령어로 시작하세요:
echo.
echo   claude
echo.
pause
