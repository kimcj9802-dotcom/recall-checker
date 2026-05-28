"""
MFDS 의료기기안심책방 회수목록 스크래퍼
- 1차: requests로 직접 Excel 다운로드 시도
- 2차: Selenium (Chrome 헤드리스) 으로 JS 렌더링 후 파싱
- 6시간 캐시
"""
import io
import json
import os
import time
from datetime import datetime, timedelta

import requests

CACHE_FILE = os.path.join(os.path.dirname(__file__), "cache", "recalls.json")
CACHE_HOURS = 6
BASE_URL = "https://emedi.mfds.go.kr"
RECALL_URL = f"{BASE_URL}/recall/MNU20265"

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9",
    "Referer": RECALL_URL,
}


# ── 캐시 ────────────────────────────────────────────────────
def _load_cache(ignore_expiry=False):
    if not os.path.exists(CACHE_FILE):
        return None
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not ignore_expiry:
            age = datetime.now() - datetime.fromisoformat(data.get("cached_at", "2000-01-01"))
            if age > timedelta(hours=CACHE_HOURS):
                return None
        return data.get("recalls", [])
    except Exception:
        return None


def _save_cache(recalls: list, scraped_at: str = ""):
    """
    scraped_at: 실제 MFDS 스크래핑 시각 (GitHub Actions가 기록한 시각).
    미전달 시 현재 시각으로 기록.
    """
    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "cached_at":  scraped_at or datetime.now().isoformat(),
            "fetched_at": datetime.now().isoformat(),   # Render가 GitHub에서 받은 시각
            "recalls":    recalls,
        }, f, ensure_ascii=False, indent=2)


def get_cache_info() -> dict:
    if not os.path.exists(CACHE_FILE):
        return {"exists": False, "count": 0}
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        cached_at  = datetime.fromisoformat(data.get("cached_at",  "2000-01-01"))
        fetched_at = data.get("fetched_at", "")
        count = len(data.get("recalls", []))
        info = {
            "exists":    True,
            "cached_at": cached_at.strftime("%Y-%m-%d %H:%M"),   # 실제 수집(스크래핑) 시각
            "count":     count,
            "expired":   datetime.now() - cached_at > timedelta(hours=CACHE_HOURS),
        }
        if fetched_at:
            try:
                info["fetched_at"] = datetime.fromisoformat(fetched_at).strftime("%Y-%m-%d %H:%M")
            except Exception:
                pass
        return info
    except Exception:
        return {"exists": False, "count": 0}


# ── 1차: requests로 Excel 직접 다운로드 시도 ─────────────────
_EXCEL_ENDPOINTS = [
    "/recall/selectRcllExcelDown.do",
    "/recall/rcllExcelDown.do",
    "/recall/excelDownload.do",
    "/recall/rcllMngExcelDown.do",
    "/recall/selectRcllMngExcelDown.do",
]

_SEARCH_ENDPOINTS = [
    "/recall/selectRcllMngList.do",
    "/recall/getRcllMngList.do",
    "/recall/rcllMngList.do",
    "/recall/selectRcllList.do",
]

def _date_params() -> dict:
    """최근 1개월 날짜 범위 파라미터"""
    end   = datetime.now().strftime("%Y-%m-%d")
    start = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    return {
        # MFDS 실제 파라미터 (detail URL에서 확인)
        "startPlanSbmsnDt": start,
        "endPlanSbmsnDt":   end,
        "searchYn": "true",
        "mid": "MNU20265",
        # 공통 후보
        "pageIndex": "1", "pageNum": "1", "pageUnit": "1000",
        "searchGbn": "", "searchWrd": "",
        "rcllPrgYn": "", "govCorpGbn": "",
        # 날짜 필드명 후보 (호환)
        "startDt": start, "endDt": end,
        "schStartDt": start, "schEndDt": end,
    }

_SEARCH_PARAMS = _date_params


def _try_requests() -> list:
    """requests로 JSON/Excel/HTML 응답을 순서대로 시도"""
    sess = requests.Session()
    sess.headers.update(_HEADERS)
    try:
        sess.get(RECALL_URL, timeout=10)  # 쿠키 획득
    except Exception:
        pass

    params = _SEARCH_PARAMS()

    # Excel 다운로드 시도
    for ep in _EXCEL_ENDPOINTS:
        try:
            r = sess.post(f"{BASE_URL}{ep}", data=params, timeout=15)
            ct = r.headers.get("content-type", "")
            if "excel" in ct or "spreadsheet" in ct or r.content[:4] == b"PK\x03\x04":
                print(f"[scraper] Excel 다운로드 성공: {ep}")
                import pandas as pd
                df = pd.read_excel(io.BytesIO(r.content))
                return _normalize_items([{str(k): v for k, v in row.items()} for row in df.to_dict("records")])
        except Exception:
            pass

    # JSON/HTML 검색 API 시도 (POST + GET 모두 시도)
    for ep in _SEARCH_ENDPOINTS:
        for method in ("post", "get"):
            try:
                fn = getattr(sess, method)
                r = fn(f"{BASE_URL}{ep}", data=params if method == "post" else None,
                       params=params if method == "get" else None, timeout=15)
                if r.status_code != 200:
                    continue
                ct = r.headers.get("content-type", "")
                if "json" in ct:
                    body = r.json()
                    items = _extract_list(body)
                    if items:
                        print(f"[scraper] JSON API 성공 ({method.upper()}): {ep}")
                        return _normalize_items(items)
                elif "html" in ct:
                    rows = _parse_html(r.text)
                    _JUNK = ("없습니다", "클릭하세요", "오류", "error", "404", "검출유형")
                    real_rows = [row for row in rows
                                 if not any(any(j in str(v).lower() for j in _JUNK)
                                            for v in row.values())]
                    if real_rows:
                        print(f"[scraper] HTML 파싱 성공 ({method.upper()}): {ep}")
                        return real_rows
            except Exception:
                pass

    return []


def _parse_html(html: str) -> list:
    """BeautifulSoup 없이 간단한 HTML 테이블 파싱"""
    try:
        from html.parser import HTMLParser

        class TableParser(HTMLParser):
            def __init__(self):
                super().__init__()
                self.in_table = self.in_tr = self.in_th = self.in_td = False
                self.headers = []
                self.rows = []
                self.current_row = []
                self.current_cell = ""
                self.header_done = False

            def handle_starttag(self, tag, attrs):
                if tag == "table": self.in_table = True
                elif tag == "tr" and self.in_table:
                    self.in_tr = True; self.current_row = []
                elif tag == "th" and self.in_tr:
                    self.in_th = True; self.current_cell = ""
                elif tag == "td" and self.in_tr:
                    self.in_td = True; self.current_cell = ""

            def handle_endtag(self, tag):
                if tag == "table": self.in_table = False
                elif tag == "tr" and self.in_tr:
                    self.in_tr = False
                    if self.current_row:
                        if not self.header_done and not self.headers:
                            pass
                        elif self.headers:
                            obj = {self.headers[i]: v for i, v in enumerate(self.current_row) if i < len(self.headers)}
                            self.rows.append(obj)
                elif tag == "th" and self.in_th:
                    self.in_th = False
                    self.headers.append(self.current_cell.strip())
                    self.header_done = True
                elif tag == "td" and self.in_td:
                    self.in_td = False
                    self.current_row.append(self.current_cell.strip())

            def handle_data(self, data):
                if self.in_th or self.in_td:
                    self.current_cell += data

        parser = TableParser()
        parser.feed(html)
        return _normalize_items(parser.rows) if parser.rows else []
    except Exception:
        return []


# ── 2차: Selenium 헤드리스 ───────────────────────────────────
def _try_selenium() -> list:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1280,900")
    opts.add_argument(f"user-agent={_HEADERS['User-Agent']}")
    # 네트워크 요청 로깅 활성화 (CDP)
    opts.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    driver = _create_driver(opts)
    try:
        wait = WebDriverWait(driver, 20)
        # 30일을 10일씩 나눠 구간별 폼 제출 → 날짜 필터 정확 적용
        end_dt     = datetime.now()
        start_dt   = end_dt - timedelta(days=30)
        chunk_days = 10

        seen_dnos = set()
        recalls   = []
        rescued   = {}  # 날짜 범위 밖이지만 아직 미수집 항목 {key: item}

        chunk_start = start_dt
        chunk_num   = 1
        while chunk_start < end_dt:
            chunk_end = min(chunk_start + timedelta(days=chunk_days), end_dt)
            cs = chunk_start.strftime("%Y-%m-%d")
            ce = chunk_end.strftime("%Y-%m-%d")
            print(f"[scraper] 구간{chunk_num}: {cs} ~ {ce}")

            # 매 구간마다 MFDS 페이지 새로 로드 후 날짜 입력 → 검색 버튼 클릭
            driver.get(RECALL_URL)
            time.sleep(2)

            # 날짜 필드 설정 (v1.0 동일 로직)
            set_result = driver.execute_script(f"""
                const startSels = ['[name*="start" i]','[name*="bgn" i]','[name*="strt" i]','[id*="start" i]','[id*="bgn" i]'];
                const endSels   = ['[name*="end" i]',  '[name*="cls" i]', '[name*="fnsh" i]','[id*="end" i]',  '[id*="cls" i]'];
                function setVal(sels, val) {{
                    for (const s of sels) {{
                        for (const el of document.querySelectorAll(s)) {{
                            if (el.type === 'text' || el.type === 'date') {{
                                el.value = val;
                                el.dispatchEvent(new Event('change', {{bubbles:true}}));
                                return el.name || el.id || s;
                            }}
                        }}
                    }}
                    return null;
                }}
                const r1 = setVal(startSels, '{cs}');
                const r2 = setVal(endSels,   '{ce}');
                return r1 + '|' + r2;
            """)
            print(f"[scraper] 날짜 설정: {set_result}")

            # 검색 버튼 클릭
            search_btn = None
            for selector in [
                (By.XPATH, "//button[contains(., '검색')]"),
                (By.XPATH, "//input[@type='button' and @value='검색']"),
                (By.XPATH, "//a[contains(., '검색') and not(contains(., '재검색'))]"),
                (By.CSS_SELECTOR, "button.btn-search, button.search-btn, .btn_search"),
            ]:
                try:
                    search_btn = wait.until(EC.element_to_be_clickable(selector))
                    break
                except Exception:
                    continue

            if search_btn:
                search_btn.click()
            else:
                try:
                    driver.execute_script("document.querySelector('form').submit()")
                except Exception:
                    pass

            time.sleep(3)

            items = _parse_driver_table(driver)
            print(f"[scraper] 구간{chunk_num} 원본: {len(items)}건 파싱")

            new_cnt = 0
            for item in items:
                dno = item.get("deptReceiptNo", "")
                key = dno if dno else f"{item.get('품목명','')}-{item.get('업체명','')}-{item.get('보고일','')}"
                if not key:
                    continue

                # 날짜 파싱 → 구간 범위 내 여부 확인
                date_str = item.get("보고일", "").strip()
                in_range = True  # 날짜 없으면 범위 안으로 처리
                if date_str and date_str not in ("-", "None", "nan"):
                    parsed_d = None
                    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
                        try:
                            parsed_d = datetime.strptime(date_str[:10], fmt)
                            break
                        except ValueError:
                            continue
                    if parsed_d:
                        in_range = chunk_start <= parsed_d <= chunk_end

                if in_range:
                    # 구간 범위 내: 즉시 수집, rescued에서 제거
                    if key not in seen_dnos:
                        seen_dnos.add(key)
                        recalls.append(item)
                        new_cnt += 1
                        rescued.pop(key, None)
                else:
                    # 구간 범위 외: 아직 미수집이면 rescued에 보관
                    if key not in seen_dnos:
                        rescued[key] = item
                        print(f"[scraper]   범위 외 보관: {date_str} ({item.get('품목명','')})")

            print(f"[scraper] 구간{chunk_num} 수집: {new_cnt}건 (신규)")
            chunk_start = chunk_end + timedelta(days=1)
            chunk_num  += 1

        # 어느 구간에서도 날짜 범위 안으로 처리 안 된 항목 추가 (누락 방지)
        rescued_cnt = 0
        for key, item in rescued.items():
            if key not in seen_dnos:
                seen_dnos.add(key)
                recalls.append(item)
                rescued_cnt += 1
        if rescued_cnt:
            print(f"[scraper] 구조 항목 추가: {rescued_cnt}건")

        print(f"[scraper] 전체 수집: {len(recalls)}건")
        return recalls
    finally:
        driver.quit()


EDGE_BINARY   = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
CHROME_BINARY = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# 환경변수로 Chrome 경로 오버라이드 가능 (GitHub Actions: /usr/bin/google-chrome)
_ENV_CHROME = os.environ.get("CHROME_BIN", "")


def _create_driver(chrome_opts):
    """Edge / Chrome / 환경변수 Chrome으로 WebDriver 생성 (자동 드라이버 설치)"""
    from selenium import webdriver

    # ── 환경변수 Chrome 우선 (GitHub Actions / Linux) ──
    if _ENV_CHROME and os.path.exists(_ENV_CHROME):
        chrome_opts.binary_location = _ENV_CHROME
        # 1순위: selenium 내장 selenium-manager (4.6+, 버전 자동 매칭)
        try:
            driver = webdriver.Chrome(options=chrome_opts)
            print(f"[scraper] Chrome 사용 - selenium-manager ({_ENV_CHROME})")
            return driver
        except Exception as e:
            print(f"[scraper] selenium-manager 실패: {e}")
        # 2순위: webdriver_manager
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            from selenium.webdriver.chrome.service import Service
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_opts)
            print(f"[scraper] Chrome 사용 - webdriver_manager ({_ENV_CHROME})")
            return driver
        except Exception as e:
            print(f"[scraper] webdriver_manager 실패: {e}")

    # ── Edge 시도 (Windows 기본 내장) ──
    if os.path.exists(EDGE_BINARY):
        try:
            from selenium.webdriver.edge.options import Options as EdgeOptions
            from selenium.webdriver.edge.service import Service as EdgeService
            from webdriver_manager.microsoft import EdgeChromiumDriverManager

            edge_opts = EdgeOptions()
            for arg in chrome_opts.arguments:
                edge_opts.add_argument(arg)
            edge_opts.binary_location = EDGE_BINARY

            service = EdgeService(EdgeChromiumDriverManager().install())
            driver = webdriver.Edge(service=service, options=edge_opts)
            print("[scraper] Microsoft Edge 사용")
            return driver
        except Exception as e:
            print(f"[scraper] Edge 실패: {e}")

    # ── Chrome 시도 ──
    chrome_path = CHROME_BINARY if os.path.exists(CHROME_BINARY) else None
    # 사용자 로컬 Chrome
    local_chrome = os.path.join(os.environ.get("LOCALAPPDATA", ""),
                                "Google", "Chrome", "Application", "chrome.exe")
    if not chrome_path and os.path.exists(local_chrome):
        chrome_path = local_chrome

    if chrome_path:
        chrome_opts.binary_location = chrome_path

    try:
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.chrome.service import Service
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_opts)
        print("[scraper] Google Chrome 사용")
        return driver
    except Exception as e:
        print(f"[scraper] Chrome 실패: {e}")

    # ── 시스템 기본값 시도 ──
    try:
        driver = webdriver.Chrome(options=chrome_opts)
        return driver
    except Exception as e:
        raise RuntimeError(
            f"브라우저를 찾을 수 없습니다: {e}\n"
            "Microsoft Edge 또는 Google Chrome이 설치되어 있어야 합니다."
        )


def _find_list_api(driver) -> tuple:
    """네트워크 로그에서 MFDS 목록 API의 URL·메서드·원본 파라미터·쿠키를 반환"""
    import json as _json
    from urllib.parse import parse_qs, urlencode, urlparse
    cookies = {c["name"]: c["value"] for c in driver.get_cookies()}
    try:
        logs = driver.get_log("performance")
        for entry in logs:
            msg = _json.loads(entry["message"])["message"]
            if msg.get("method") != "Network.requestWillBeSent":
                continue
            req = msg["params"].get("request", {})
            url = req.get("url", "")
            if "mfds.go.kr" not in url:
                continue
            if "/list/" not in url.lower() and "List" not in url:
                continue
            method    = req.get("method", "GET")
            post_data = req.get("postData", "")
            base_url  = url.split("?")[0]
            # GET 파라미터도 수집
            qs = urlparse(url).query
            print(f"[scraper] API 탐지: {method} {base_url}")
            print(f"[scraper] POST body: {post_data[:200]}")
            return base_url, method, post_data or qs, cookies
    except Exception as e:
        print(f"[scraper] API 탐지 실패: {e}")
    return None, None, None, cookies


def _capture_from_network_logs(driver) -> list:
    """Chrome Performance 로그에서 JSON API 응답 추출"""
    try:
        import json as _json
        logs = driver.get_log("performance")
        for entry in reversed(logs):
            msg = _json.loads(entry["message"])["message"]
            if msg.get("method") == "Network.responseReceived":
                resp = msg["params"].get("response", {})
                url = resp.get("url", "")
                mime = resp.get("mimeType", "")
                if "json" in mime and any(k in url for k in ["rcll", "recall", "Rcll"]):
                    try:
                        req_id = msg["params"]["requestId"]
                        body = driver.execute_cdp_cmd(
                            "Network.getResponseBody", {"requestId": req_id}
                        )
                        data = _json.loads(body.get("body", "{}"))
                        items = _extract_list(data)
                        if items:
                            print(f"[scraper] CDP에서 API 응답 캡처: {url}")
                            return _normalize_items(items)
                    except Exception:
                        pass
    except Exception:
        pass
    return []


def _extract_list(data) -> list:
    if isinstance(data, list):
        return data
    for key in ["list", "items", "data", "result", "rows", "content", "resultList"]:
        if isinstance(data, dict) and isinstance(data.get(key), list):
            return data[key]
    return []


def _parse_driver_table(driver) -> list:
    """Selenium driver에서 테이블 파싱"""
    try:
        result = driver.execute_script("""
            const rows = [];
            for (const table of document.querySelectorAll('table')) {
                const ths = [...table.querySelectorAll('thead th, tr:first-child th')];
                const headers = ths.map(th => th.innerText.trim()).filter(Boolean);
                if (!headers.length) continue;
                for (const tr of table.querySelectorAll('tbody tr')) {
                    const cells = [...tr.querySelectorAll('td')].map(td => td.innerText.trim());
                    if (!cells.length || cells.every(c => !c)) continue;
                    const obj = {};
                    headers.forEach((h, i) => { obj[h] = cells[i] ?? ''; });
                    const a = tr.querySelector('a[href*="deptReceiptNo"]');
                    if (a) {
                        obj['상세링크'] = a.href;
                        const m = a.href.match(/deptReceiptNo=([^&]+)/);
                        if (m) obj['deptReceiptNo'] = m[1];
                    }
                    rows.push(obj);
                }
            }
            return rows;
        """)
        return _normalize_items(result or [])
    except Exception:
        return []


# ── 필드 정규화 ──────────────────────────────────────────────
_FIELD_MAP = {
    "품목명": "품목명", "제품명": "품목명",
    "업체명": "업체명", "제조사": "업체명",
    "모델명": "모델명",
    "품목허가(인증‧신고)번호": "허가번호", "품목허가번호": "허가번호", "허가번호": "허가번호",
    "정부/영업자 회수 구분": "회수구분", "회수구분": "회수구분",
    "회수 진행 여부": "회수진행여부", "회수진행여부": "회수진행여부",
    "보고일": "보고일", "공고일": "보고일", "보고일자": "보고일",
    "상세링크": "상세링크", "deptReceiptNo": "deptReceiptNo",
    "순번": "순번",
}


def _normalize_items(items: list) -> list:
    result = []
    for item in items:
        n = {}
        for k, v in item.items():
            mapped = _FIELD_MAP.get(str(k).strip(), str(k).strip())
            n[mapped] = str(v).strip() if v is not None else ""
        # 품목명(또는 제품명)이 있으면 유효한 레코드로 처리
        name = (n.get("품목명", "") or n.get("제품명", "")).strip()
        if name and name not in ("-", "None", "nan"):
            result.append(n)
    return result


# ── GitHub 캐시 로드 (클라우드 배포용) ───────────────────────────
def _fetch_from_github() -> list:
    """
    환경변수 GITHUB_CACHE_URL이 설정된 경우 GitHub raw URL에서
    recalls.json을 다운로드해 로컬 캐시에 저장 후 반환.
    Render 등 클라우드에서 Selenium 없이 최신 데이터 사용 가능.
    """
    url = os.environ.get("GITHUB_CACHE_URL", "")
    if not url:
        return []
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        data = r.json()
        recalls = data.get("recalls", [])
        # GitHub Actions가 기록한 실제 스크래핑 시각을 그대로 보존
        scraped_at = data.get("cached_at", "")
        if recalls:
            _save_cache(recalls, scraped_at=scraped_at)
            print(f"[scraper] GitHub 캐시 로드 완료: {len(recalls)}건 (수집: {scraped_at})")
        return recalls
    except Exception as e:
        print(f"[scraper] GitHub 캐시 로드 실패: {e}")
        return []


def get_github_cached_at() -> str:
    """GitHub 원본 파일의 cached_at(실제 수집 시각) 반환. 비교용."""
    url = os.environ.get("GITHUB_CACHE_URL", "")
    if not url:
        return ""
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return r.json().get("cached_at", "")
    except Exception:
        return ""


def trigger_workflow_refresh() -> tuple:
    """
    GitHub Actions 워크플로우를 workflow_dispatch 이벤트로 트리거.
    Render 환경에서 "즉시 갱신"을 가능하게 함.

    필요한 환경변수:
      GITHUB_TOKEN : Personal Access Token (Actions: write 권한)
      GITHUB_REPO  : "owner/repo" 형식 (기본: kimcj9802-dotcom/recall-checker)
    """
    token    = os.environ.get("RECALL_GITHUB_TOKEN", "")
    repo     = os.environ.get("GITHUB_REPO", "kimcj9802-dotcom/recall-checker")
    workflow = "scrape_recalls.yml"

    if not token:
        return False, "RECALL_GITHUB_TOKEN 환경변수 미설정"

    api_url = f"https://api.github.com/repos/{repo}/actions/workflows/{workflow}/dispatches"
    try:
        resp = requests.post(
            api_url,
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/vnd.github+json",
                "X-GitHub-Api-Version": "2022-11-28",
            },
            json={"ref": "main"},
            timeout=15,
        )
        if resp.status_code == 204:
            print(f"[scraper] GitHub Actions 워크플로우 트리거 성공 ({repo})")
            return True, "MFDS 스크래핑 워크플로우를 시작했습니다. 약 2~3분 소요됩니다."
        body = resp.text[:300]
        print(f"[scraper] GitHub Actions 트리거 실패: {resp.status_code} {body}")
        return False, f"GitHub API {resp.status_code}: {body}"
    except Exception as e:
        print(f"[scraper] GitHub Actions 트리거 예외: {e}")
        return False, str(e)


# ── 공개 API ─────────────────────────────────────────────────
def fetch_recall_data(force_refresh: bool = False) -> list:
    github_url = os.environ.get("GITHUB_CACHE_URL", "")

    # ── 클라우드 환경 (GITHUB_CACHE_URL 설정 시): Selenium 미사용 ──
    if github_url:
        if not force_refresh:
            cached = _load_cache()
            if cached is not None:
                return cached
        # force_refresh이거나 로컬 캐시 없으면 GitHub에서 가져옴
        github_data = _fetch_from_github()
        if github_data:
            return github_data
        # GitHub도 실패하면 만료 캐시라도
        stale = _load_cache(ignore_expiry=True)
        if stale:
            print("[scraper] 만료 캐시 반환 (GitHub 캐시 로드 실패)")
            return stale
        raise RuntimeError(
            "GitHub 캐시를 불러올 수 없습니다. "
            "GitHub Actions 워크플로우(MFDS 회수목록 자동 수집)를 수동 실행해 주세요."
        )

    # ── 로컬 환경: requests → Selenium 순서로 시도 ──
    if not force_refresh:
        cached = _load_cache()
        if cached is not None:
            return cached

    errors = []

    # 1차: requests
    try:
        recalls = _try_requests()
        if recalls:
            _save_cache(recalls)
            return recalls
    except Exception as e:
        errors.append(f"requests: {e}")

    # 2차: Selenium
    try:
        recalls = _try_selenium()
        if recalls:
            _save_cache(recalls)
            return recalls
    except Exception as e:
        errors.append(f"selenium: {e}")

    # 만료 캐시라도 반환
    stale = _load_cache(ignore_expiry=True)
    if stale:
        print("[scraper] 만료 캐시 반환")
        return stale

    raise RuntimeError("회수목록 조회 실패: " + " / ".join(errors))


def fetch_recall_detail(dept_receipt_no: str) -> list:
    """
    MFDS 상세 페이지에서 제조번호(로트) 목록 조회

    근본 원인별 대응:
    1) 세션/쿠키: 목록 페이지 방문 후 상세 접근 → 일부 품목 HTML 정상 반환
    2) recallSeq 패턴: 따옴표 유무·속성명(recallSeq/recallItemSeq) 다양하게 처리
    3) 폴백 강화: seq 미탐지 시 1~10 순서 시도, 3연속 빈 응답이면 조기 종료

    반환: [{"typeName": 모델명, "makeNo": 제조번호, "makeDate": 제조일자, "recallTargetQty": 수량}, ...]
    """
    import re as _re

    detail_url = (
        f"{BASE_URL}/recall/view/MNU20265"
        f"?mid=MNU20265&pageNum=1&deptReceiptNo={dept_receipt_no}"
    )

    sess = requests.Session()
    sess.headers.update(_HEADERS)

    # ── 0단계: 목록 페이지 방문 → 세션/쿠키 확립 ─────────────────
    # 세션 없이 바로 상세 접근 시 일부 품목에서 HTML 불완전 반환 방지
    try:
        sess.get(RECALL_URL, timeout=10)
    except Exception:
        pass

    # ── 1단계: 상세 페이지 로드 → recallSeq 목록 수집 ─────────────
    # MFDS HTML의 recallSeq 속성명·따옴표 유무가 품목마다 다를 수 있음
    seqs = []
    try:
        resp = sess.get(detail_url, timeout=15)
        html = resp.text
        for pat in [
            r'recallItemSeq\s*=\s*["\'](\d+)["\']',  # recallItemSeq="1"
            r'recallSeq\s*=\s*["\'](\d+)["\']',       # recallSeq="1"
            r'recallItemSeq\s*=\s*(\d+)',              # recallItemSeq=1 (따옴표 없음)
            r'recallSeq\s*=\s*(\d+)',                  # recallSeq=1 (따옴표 없음)
            r'"recallItemSeq"\s*:\s*(\d+)',             # JSON 형태
            r'"recallSeq"\s*:\s*(\d+)',                 # JSON 형태
        ]:
            found = list(dict.fromkeys(_re.findall(pat, html)))
            if found:
                seqs = found
                print(f"[scraper] seq탐지({dept_receipt_no}): {found[:5]}")
                break
    except Exception as e:
        print(f"[scraper] 상세페이지 오류({dept_receipt_no}): {e}")

    # seq 미탐지 시 1~10 순서로 폴백 (3연속 빈 응답 시 조기 종료)
    use_fallback = not seqs
    if use_fallback:
        seqs = [str(i) for i in range(1, 11)]
        print(f"[scraper] seq미탐지({dept_receipt_no}) → 1~10 폴백 시도")

    # ── 2단계: recallSeq별 제조번호 API 호출 ──────────────────────
    models = []
    seen_keys: set = set()
    model_api = f"{BASE_URL}/recall/view/model"
    consecutive_empty = 0

    for seq in seqs:
        try:
            r = sess.get(
                model_api,
                params={"deptReceiptNo": dept_receipt_no, "recallItemSeq": seq},
                headers={
                    "Accept": "application/json, text/javascript, */*; q=0.01",
                    "X-Requested-With": "XMLHttpRequest",
                    "Referer": detail_url,
                },
                timeout=15,
            )
            if r.status_code != 200 or "json" not in r.headers.get("content-type", ""):
                if use_fallback:
                    consecutive_empty += 1
                    if consecutive_empty >= 3:
                        break
                continue

            data = r.json()
            # 응답 키가 품목마다 다를 수 있으므로 여러 키 시도
            items = (
                data.get("recallModel")
                or data.get("list")
                or data.get("items")
                or data.get("result")
                or []
            )

            if not items:
                if use_fallback:
                    consecutive_empty += 1
                    if consecutive_empty >= 3:
                        break
                continue

            consecutive_empty = 0
            for item in items:
                key = f"{item.get('makeNo', '')}-{item.get('typeName', '')}"
                if key not in seen_keys:
                    seen_keys.add(key)
                    models.append(item)

        except Exception as e:
            print(f"[scraper] model API 오류(seq={seq}): {e}")
            if use_fallback:
                consecutive_empty += 1

    if models:
        print(f"[scraper] 제조번호 {len(models)}건 ({dept_receipt_no})")
    else:
        print(f"[scraper] 제조번호 없음 ({dept_receipt_no}) — seqs={seqs[:5]}")

    return models


def load_recalls_from_file(file_bytes: bytes, filename: str) -> list:
    import pandas as pd
    if filename.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8-sig")
    else:
        df = pd.read_excel(io.BytesIO(file_bytes))
    items = [{str(k): v for k, v in r.items()} for r in df.to_dict("records")]
    return _normalize_items(items)
