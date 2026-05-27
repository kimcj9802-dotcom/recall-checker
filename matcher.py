import re
import unicodedata
import pandas as pd
from thefuzz import fuzz

# 매칭 임계값 (0~100)
THRESHOLD_HIGH = 85   # 확실한 매칭
THRESHOLD_LOW = 70    # 가능성 있는 매칭 (검토 필요) ← 65→70 상향
THRESHOLD_NAME_MIN = 70  # 제품명 최소 유사도 (이 미만이면 종합점수 무관하게 제외)


def _clean(text: str) -> str:
    """비교용 문자열 정규화"""
    if not text:
        return ""
    text = str(text).strip()
    # 괄호 안 내용 제거, 특수문자 정규화
    text = re.sub(r"\s+", " ", text)
    text = text.upper()
    return text


def _dominant_script(s: str) -> str:
    """문자열의 주요 문자 체계 반환 ('korean' | 'ascii' | 'mixed')"""
    ko = sum(1 for c in s if '가' <= c <= '힣' or 'ᄀ' <= c <= 'ᇿ')
    en = sum(1 for c in s if c.isascii() and c.isalpha())
    if ko > en * 2:
        return 'korean'
    if en > ko * 2:
        return 'ascii'
    return 'mixed'


def _match_score(a: str, b: str) -> int:
    """두 문자열의 유사도 점수 (0-100)"""
    a, b = _clean(a), _clean(b)
    if not a or not b:
        return 0
    # 완전 포함 관계
    if a in b or b in a:
        return 95
    return max(
        fuzz.token_sort_ratio(a, b),
        fuzz.partial_ratio(a, b),
        fuzz.token_set_ratio(a, b),
    )


def _maker_score(asset_maker: str, recall_maker: str) -> int:
    """
    제조사/업체명 유사도 계산.
    영문↔한글 처럼 문자 체계가 다를 경우 중립값(50) 반환
    (예: 'INTUITIVE SURGICAL' vs '인튜이티브서지컬코리아인크')
    """
    a, b = _clean(asset_maker), _clean(recall_maker)
    if not a or not b:
        return 50  # 어느 한 쪽이 없으면 중립
    sa, sb = _dominant_script(a), _dominant_script(b)
    if sa != sb and sa != 'mixed' and sb != 'mixed':
        # 다른 언어 체계 → 비교 불가, 중립 처리
        return 50
    return _match_score(asset_maker, recall_maker)


def _lot_match(asset_lot: str, recall_lot: str) -> bool:
    """로트번호 매칭: 범위 표현(예: '001-050') 지원"""
    a = _clean(asset_lot)
    r = _clean(recall_lot)
    if not a or not r:
        return True  # 로트번호 없으면 무시
    if a == r or a in r or r in a:
        return True
    # 범위 표현 처리 (예: "001-050")
    range_match = re.match(r"^(\d+)\s*[-~]\s*(\d+)$", r)
    if range_match and re.match(r"^\d+$", a):
        lo, hi = int(range_match.group(1)), int(range_match.group(2))
        try:
            return lo <= int(a) <= hi
        except ValueError:
            pass
    return False


def _detect_columns(df: pd.DataFrame) -> dict:
    """
    업로드된 엑셀의 컬럼명을 자동 감지합니다.
    다양한 컬럼명 패턴을 처리합니다.
    """
    col_map = {"name": None, "maker": None, "model": None, "lot": None, "asset_id": None}
    patterns = {
        "name":     ["한글명칭", "명칭", "기기명칭", "자산명", "장비명칭",
                     "제품명", "품목명", "기기명", "장비명", "의료기기명", "기기 명", "제품 명"],
        "maker":    ["업체명", "제조사", "제조업체", "회사명", "브랜드", "메이커", "제조원"],
        "model":    ["모델명", "모델번호", "모델", "형식", "형명", "제품모델", "model"],
        "lot":      ["제조번호", "제조 번호", "로트번호", "로트", "lot", "lot번호",
                     "배치번호", "시리얼", "시리얼번호", "s/n", "sn"],
        "asset_id": ["자산번호", "자산코드", "관리번호", "코드", "번호", "id"],
    }
    for col in df.columns:
        col_lower = str(col).strip().lower()
        for field, kws in patterns.items():
            if col_map[field] is None:
                for kw in kws:
                    if kw in col_lower:
                        col_map[field] = col
                        break
    return col_map


def find_matches(asset_df: pd.DataFrame, recalls: list) -> list:
    """
    병원 자산 목록과 회수 목록을 비교하여 매칭 결과를 반환합니다.

    Returns:
        list of dict with keys:
          - asset_row: 원본 자산 데이터
          - recall: 매칭된 회수 항목
          - score: 매칭 점수
          - level: 'HIGH' | 'LOW'
          - reasons: 매칭 근거 목록
    """
    col_map = _detect_columns(asset_df)
    results = []

    for _, asset_row in asset_df.iterrows():
        asset = {
            "name": str(asset_row.get(col_map["name"], "") or ""),
            "maker": str(asset_row.get(col_map["maker"], "") or ""),
            "model": str(asset_row.get(col_map["model"], "") or ""),
            "lot": str(asset_row.get(col_map["lot"], "") or ""),
            "asset_id": str(asset_row.get(col_map["asset_id"], "") or ""),
        }

        # 이름이 없으면 매칭 불가
        if not asset["name"].strip():
            continue

        for recall in recalls:
            r_name = recall.get("제품명", "") or recall.get("품목명", "")
            r_maker = recall.get("업체명", "")
            r_model = recall.get("모델명", "")
            r_lot = recall.get("로트번호", "")

            name_score  = _match_score(asset["name"], r_name)
            maker_score = _maker_score(asset["maker"], r_maker)
            model_score = _match_score(asset["model"], r_model) if asset["model"] and r_model else 50

            # 제품명 유사도가 너무 낮으면 제외 (어미 우연 일치 등 오매칭 방지)
            if name_score < THRESHOLD_NAME_MIN:
                continue

            # 가중 점수: 제품명 50% + 업체명 30% + 모델명 20%
            combined = name_score * 0.5 + maker_score * 0.3 + model_score * 0.2

            if combined < THRESHOLD_LOW:
                continue

            # 로트번호가 있을 때 불일치하면 제외
            if asset["lot"] and r_lot and not _lot_match(asset["lot"], r_lot):
                continue

            reasons = []
            if name_score >= THRESHOLD_LOW:
                reasons.append(f"제품명 유사도 {name_score}%")
            if maker_score >= THRESHOLD_LOW and asset["maker"] and r_maker:
                reasons.append(f"업체명 유사도 {maker_score}%")
            if model_score >= THRESHOLD_LOW and asset["model"] and r_model:
                reasons.append(f"모델명 유사도 {model_score}%")
            if asset["lot"] and r_lot:
                reasons.append("로트번호 일치")

            asset_dict = {str(k): (str(v) if not pd.isna(v) else "") for k, v in asset_row.items()}

            results.append({
                "asset_row": asset_dict,
                "asset_display": {
                    "자산번호": asset["asset_id"],
                    "제품명": asset["name"],
                    "업체명": asset["maker"],
                    "모델명": asset["model"],
                    "로트번호": asset["lot"],
                },
                "recall": {
                    "품목명":     r_name,
                    "제품명":     r_name,
                    "업체명":     r_maker,
                    "모델명":     r_model,
                    "로트번호":   r_lot,
                    "회수구분":   recall.get("회수구분", ""),
                    "회수진행여부": recall.get("회수진행여부", ""),
                    "보고일":     recall.get("보고일") or recall.get("보고일자", ""),
                    "허가번호":   recall.get("허가번호", ""),
                    "상세링크":   recall.get("상세링크", ""),
                    "deptReceiptNo": recall.get("deptReceiptNo", ""),
                },
                "score": round(combined),
                "name_score": name_score,
                "level": "HIGH" if combined >= THRESHOLD_HIGH else "LOW",
                "reasons": reasons,
            })

    # 점수 높은 순 정렬
    results.sort(key=lambda x: x["score"], reverse=True)
    return results
