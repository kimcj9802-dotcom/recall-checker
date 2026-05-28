"""
GitHub Actions에서 실행되는 MFDS 회수목록 수집 스크립트.
ubuntu-latest 환경에서 google-chrome으로 Selenium 직접 실행.

NOTE: fetch_recall_data()는 requests → Selenium 순서로 시도하므로
      requests 결과가 부분적(소수 건)이어도 그냥 반환해 버린다.
      여기서는 Selenium을 직접 호출해 전체 데이터를 확실히 수집한다.
"""
import os
import sys

# GitHub Actions용 Chrome 경로 설정
os.environ.setdefault("CHROME_BIN", "/usr/bin/google-chrome")

from scraper import _try_selenium, _save_cache, get_cache_info

print("=== MFDS 회수목록 수집 시작 (Selenium) ===")
try:
    recalls = _try_selenium()
    if not recalls:
        raise RuntimeError("Selenium 스크래핑 결과가 비어있습니다.")
    _save_cache(recalls)
    info = get_cache_info()
    print(f"✅ 수집 완료: {info['count']}건 ({info['cached_at']})")
except Exception as e:
    print(f"❌ 수집 실패: {e}")
    sys.exit(1)
