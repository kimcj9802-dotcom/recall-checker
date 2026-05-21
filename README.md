# 의료기기 회수대상 자산 확인 앱

병원 자산 Excel을 업로드하면 식약처 의료기기안심책방 회수목록과 자동 비교합니다.
Playwright 헤드리스 브라우저로 MFDS 사이트를 실제로 열어 데이터를 수집합니다.

## 설치 및 실행

### Python이 없는 경우
https://www.python.org/downloads 에서 Python 3.10 이상 설치 (PATH 추가 체크)

### 최초 1회 설치

```bat
pip install -r requirements.txt
playwright install chromium
```

### 서버 실행

```bat
python app.py
```
또는 `run.bat` 더블클릭 → 브라우저에서 http://localhost:5000 접속

## 사용 방법

1. **[회수목록 갱신]** 클릭 → MFDS 사이트 자동 조회 (약 30초~1분, 이후 6시간 캐시)
2. 병원 자산 Excel 업로드
3. **[회수대상 확인]** 클릭 → 매칭 결과 확인
4. Excel 내보내기

### 자동조회 실패 시 대안
https://emedi.mfds.go.kr/recall/MNU20265 에서 수동 다운로드 후
"회수목록 직접 업로드" 칸에 업로드

## 매칭 기준

| 등급 | 점수 | 의미 |
|------|------|------|
| 높은 가능성 (빨간) | 85점 이상 | 즉시 조치 필요 |
| 검토 필요 (주황) | 65~84점 | 수동 확인 필요 |

점수 = 제품명(품목명) 50% + 업체명 30% + 모델명 20%

## 프로젝트 구조

```
recall_checker/
├── app.py          Flask 웹 서버
├── scraper.py      Playwright 기반 MFDS 스크래퍼
├── matcher.py      퍼지 매칭 알고리즘
├── requirements.txt
├── run.bat         윈도우 실행 스크립트
├── templates/
│   └── index.html  웹 UI
└── cache/          회수목록 캐시 (자동 생성)
```
