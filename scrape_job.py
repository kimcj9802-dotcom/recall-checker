"""
GitHub Actions에서 실행되는 MFDS 회수목록 수집 스크립트.
ubuntu-latest 환경에서 google-chrome으로 Selenium 실행.
"""
import os, sys

# GitHub Actions용 Chrome 경로 설정
os.environ.setdefault("CHROME_BIN", "/usr/bin/google-chrome")

# scraper 임포트 후 강제 갱신
from scraper import fetch_recall_data, get_cache_info

print("=== MFDS 회수목록 수집 시작 ===")
try:
    recalls = fetch_recall_data(force_refresh=True)
    info = get_cache_info()
    print(f"✅ 수집 완료: {info['count']}건 ({info['cached_at']})")
except Exception as e:
    print(f"❌ 수집 실패: {e}")
    sys.exit(1)
