@echo off
cd /d "%~dp0"
echo [1/3] 패키지 설치 중...
python -m pip install --prefer-binary flask openpyxl thefuzz python-Levenshtein playwright pandas
echo [2/3] Playwright 브라우저 설치 중...
python -m playwright install chromium
echo [3/3] 서버 시작...
echo 브라우저에서 http://localhost:5000 접속하세요
python app.py
pause
