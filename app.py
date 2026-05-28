import io
import os
import threading
import time
from datetime import datetime, timedelta

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file

from matcher import find_matches
from scraper import fetch_recall_data, fetch_recall_detail, get_cache_info, load_recalls_from_file, trigger_workflow_refresh, get_github_cached_at

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB


def _filter_active_recalls(recalls: list, days: int = 30) -> list:
    """최근 N일 이내 보고일자 기준 회수 목록 반환 (회수진행여부 전체, 빈 레코드 제외)"""
    cutoff = datetime.now() - timedelta(days=days)
    result = []
    for r in recalls:
        # 0) 품목명이 없는 빈 레코드 제외
        name = (r.get("품목명") or r.get("제품명", "")).strip()
        if not name or name in ("-", "None", "nan"):
            continue

        # 1) 보고일: 최근 N일 이내 (날짜가 없으면 통과)
        date_str = (r.get("보고일") or r.get("보고일자", "")).strip()
        if date_str and date_str != "-":
            parsed = None
            for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
                try:
                    parsed = datetime.strptime(date_str[:10], fmt)
                    break
                except ValueError:
                    continue
            if parsed and parsed < cutoff:
                continue

        result.append(r)
    return result


def _upgrade_by_lot(matches: list) -> list:
    """
    매칭 결과에서 자산의 제조번호/로트번호가 MFDS 회수 상세의
    제조번호(makeNo)와 일치하면 HIGH로 승급.
    - 회수별로 한 번만 상세 API 호출 (캐시 재사용)
    - 실패해도 기존 매칭 결과 유지
    """
    if not matches:
        return matches

    # deptReceiptNo별 제조번호 목록 사전 수집 (중복 방지)
    dno_set = {
        m["recall"].get("deptReceiptNo", "")
        for m in matches
        if m["recall"].get("deptReceiptNo", "")
    }
    lot_map: dict[str, set[str]] = {}  # {deptReceiptNo: {makeNo, ...}}
    for dno in dno_set:
        try:
            models = fetch_recall_detail(dno)
            lot_map[dno] = {
                str(m.get("makeNo", "")).strip().upper()
                for m in models
                if m.get("makeNo", "")
            }
        except Exception:
            lot_map[dno] = set()

    # 각 매칭 항목에 대해 제조번호 비교
    for m in matches:
        asset_lot = str(m["asset_display"].get("로트번호", "")).strip().upper()
        if not asset_lot:
            continue
        dno = m["recall"].get("deptReceiptNo", "")
        recall_lots = lot_map.get(dno, set())
        if not recall_lots:
            continue

        # 완전 일치 또는 포함 관계
        matched_lot = None
        for rl in recall_lots:
            if asset_lot == rl or asset_lot in rl or rl in asset_lot:
                matched_lot = rl
                break

        if matched_lot:
            m["level"] = "HIGH"
            m["score"] = max(m["score"], 95)
            reasons = m.get("reasons", [])
            lot_reason = f"제조번호 일치 ({asset_lot})"
            if lot_reason not in reasons:
                reasons.insert(0, lot_reason)
            m["reasons"] = reasons
            m["lot_matched"] = True

    return matches


def _read_excel_smart(file_bytes: bytes) -> pd.DataFrame:
    """
    첫 번째 시트에서 자산 관련 컬럼(한글명칭, 모델명, 제조번호 등)이
    가장 많이 감지되는 헤더 행(0~5)을 자동 선택.
    """
    ASSET_KEYWORDS = [
        "한글명칭", "명칭", "기기명", "장비명", "제품명", "품목명",
        "모델명", "모델", "제조번호", "시리얼", "자산번호", "관리번호",
        "제조사", "업체명", "제조업체",
    ]
    best_df = None
    best_score = -1

    for header_row in range(6):  # 헤더가 최대 5번째 행까지 있을 수 있음
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
            if df.empty:
                continue
            score = sum(
                1 for col in df.columns
                if any(kw in str(col) for kw in ASSET_KEYWORDS)
            )
            if score > best_score:
                best_score = score
                best_df = df
        except Exception:
            continue

    if best_df is not None and best_score > 0:
        return best_df

    # 폴백: 기존 방식
    for header_row in range(3):
        df = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
        if not all("Unnamed" in str(c) for c in df.columns):
            return df
    return pd.read_excel(io.BytesIO(file_bytes))

# 백그라운드 갱신 상태
_refresh_state = {
    "running":            False,
    "error":              None,
    "workflow_triggered": False,   # GitHub Actions 트리거 성공 여부
    "workflow_message":   "",      # 트리거 결과 메시지
    "workflow_waiting":   False,   # 워크플로우 완료 대기 중 여부
}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/debug-env")
def debug_env():
    """환경변수 진단용 (토큰 값은 노출하지 않음)"""
    token = os.environ.get("RECALL_GITHUB_TOKEN", "")
    github_cache = os.environ.get("GITHUB_CACHE_URL", "")
    # GITHUB/TOKEN/RECALL 포함 키 목록 (값 제외)
    related_keys = [k for k in os.environ if any(x in k.upper() for x in ["GITHUB", "TOKEN", "RECALL"])]
    return jsonify({
        "RECALL_GITHUB_TOKEN_set": bool(token),
        "RECALL_GITHUB_TOKEN_prefix": token[:6] + "..." if token else "(없음)",
        "GITHUB_CACHE_URL_set": bool(github_cache),
        "related_env_keys": sorted(related_keys),
    })


@app.route("/api/cache-info")
def cache_info():
    # fetch_recall_data()를 먼저 호출해 캐시 파일을 확보한 뒤 get_cache_info()로 읽어야
    # "캐시 없음" 오표시를 방지할 수 있음 (순서가 중요)
    filtered_count = 0
    try:
        all_recalls = fetch_recall_data()
        filtered_count = len(_filter_active_recalls(all_recalls))
    except Exception:
        pass
    info = get_cache_info()   # fetch 이후에 읽어야 정확한 상태 반영
    info["filtered_count"] = filtered_count
    return jsonify({**info, "refresh_running": _refresh_state["running"]})


@app.route("/api/refresh-recalls", methods=["POST"])
def refresh_recalls():
    """회수목록 새로고침 (백그라운드 실행)"""
    if _refresh_state["running"]:
        return jsonify({"success": False, "error": "이미 갱신 중입니다."}), 409

    # 스레드 시작 전에 상태 초기화 (폴링 타이밍 경합 방지)
    # running=True를 스레드 내부에서 설정하면 JS 첫 폴링이 running=False를
    # 읽고 즉시 fallback 메시지를 띄우는 race condition 발생
    _refresh_state["running"]            = True
    _refresh_state["error"]              = None
    _refresh_state["workflow_triggered"] = False
    _refresh_state["workflow_message"]   = ""
    _refresh_state["workflow_waiting"]   = False

    def _do_refresh():
        try:
            # 1) GitHub Actions 워크플로우 트리거 (RECALL_GITHUB_TOKEN 설정 시)
            ok, msg = trigger_workflow_refresh()
            _refresh_state["workflow_triggered"] = ok
            _refresh_state["workflow_message"]   = msg

            if ok:
                # 2) 트리거 전 수집 시각 기록 → 새 데이터 감지용
                prev_cached_at = get_github_cached_at()

                # 3) GitHub가 새 데이터를 커밋할 때까지 30초 간격으로 폴링 (최대 5분)
                _refresh_state["workflow_waiting"] = True
                deadline = time.time() + 300
                while time.time() < deadline:
                    time.sleep(30)
                    new_cached_at = get_github_cached_at()
                    if new_cached_at and new_cached_at != prev_cached_at:
                        fetch_recall_data(force_refresh=True)
                        _refresh_state["workflow_message"] = "스크래핑 완료! 최신 데이터로 갱신되었습니다."
                        break
                else:
                    _refresh_state["workflow_message"] = "5분 내 완료 미확인 — 잠시 후 다시 확인해 주세요."
                _refresh_state["workflow_waiting"] = False
            else:
                # 워크플로우 트리거 불가 → 현재 GitHub 캐시 동기화
                fetch_recall_data(force_refresh=True)
        except Exception as e:
            _refresh_state["error"] = str(e)
        finally:
            _refresh_state["running"] = False

    threading.Thread(target=_do_refresh, daemon=True).start()
    return jsonify({"success": True, "message": "갱신을 시작했습니다. 잠시 후 상태를 확인하세요."})


@app.route("/api/refresh-status")
def refresh_status():
    return jsonify({
        "running":            _refresh_state["running"],
        "error":              _refresh_state["error"],
        "workflow_triggered": _refresh_state["workflow_triggered"],
        "workflow_message":   _refresh_state["workflow_message"],
        "workflow_waiting":   _refresh_state["workflow_waiting"],
        **get_cache_info(),
    })


@app.route("/api/recalls")
def get_recalls():
    """캐시된 회수목록 반환 (진행중 + 최근 1개월 필터)"""
    try:
        recalls = fetch_recall_data()
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    recalls = _filter_active_recalls(recalls)

    items = []
    for r in recalls:
        items.append({
            "품목명": r.get("품목명") or r.get("제품명", ""),
            "업체명": r.get("업체명", ""),
            "회수구분": r.get("회수구분", ""),
            "보고일": r.get("보고일") or r.get("보고일자", ""),
            "회수진행여부": r.get("회수진행여부", ""),
            "허가번호": r.get("허가번호", ""),
            "상세링크": r.get("상세링크", ""),
            "deptReceiptNo": r.get("deptReceiptNo", ""),
        })
    return jsonify({"success": True, "count": len(items), "recalls": items})


@app.route("/api/recall-detail")
def recall_detail():
    """MFDS 회수 상세: 제조번호(로트) 목록 반환"""
    dept_receipt_no = request.args.get("deptReceiptNo", "").strip()
    if not dept_receipt_no:
        return jsonify({"error": "deptReceiptNo 파라미터가 필요합니다."}), 400
    try:
        models = fetch_recall_detail(dept_receipt_no)
        return jsonify({"success": True, "count": len(models), "models": models})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/check", methods=["POST"])
def check_assets():
    # 1. 자산 파일 확인
    if "asset_file" not in request.files or not request.files["asset_file"].filename:
        return jsonify({"error": "자산 파일이 없습니다."}), 400

    asset_file = request.files["asset_file"]
    fname = asset_file.filename.lower()
    if not fname.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Excel 파일(.xlsx, .xls)만 업로드 가능합니다."}), 400

    # 2. 자산 데이터 읽기
    try:
        raw_bytes = asset_file.read()
        asset_df = _read_excel_smart(raw_bytes)
        if asset_df.empty:
            return jsonify({"error": "자산 파일이 비어 있습니다."}), 400
    except Exception as e:
        return jsonify({"error": f"자산 파일 읽기 오류: {e}"}), 400

    # 3. 회수 목록 취득
    recall_source = "auto"
    recalls = []

    if "recall_file" in request.files and request.files["recall_file"].filename:
        try:
            rf = request.files["recall_file"]
            recalls = load_recalls_from_file(rf.read(), rf.filename)
            recall_source = "manual"
        except Exception as e:
            return jsonify({"error": f"회수목록 파일 오류: {e}"}), 400
    else:
        try:
            recalls = fetch_recall_data()
            recalls = _filter_active_recalls(recalls)
        except Exception as e:
            return jsonify({
                "error": f"자동 조회 실패: {e}",
                "hint": "MFDS 사이트 수동 다운로드 후 '회수목록 파일' 칸에 업로드하거나, playwright install 명령 실행 여부를 확인하세요."
            }), 500

    if not recalls:
        return jsonify({"error": "회수목록이 비어 있습니다. 갱신 후 다시 시도해주세요."}), 500

    # 4. 매칭 (이름/업체/모델 기반 1차)
    matches = find_matches(asset_df, recalls)

    # 5. 제조번호(로트번호) 교차 확인 → HIGH 승급
    matches = _upgrade_by_lot(matches)

    high = [m for m in matches if m["level"] == "HIGH"]
    low  = [m for m in matches if m["level"] == "LOW"]
    cache = get_cache_info()

    return jsonify({
        "success": True,
        "total_assets": len(asset_df),
        "total_recalls": len(recalls),
        "recall_source": recall_source,
        "recall_cached_at": cache.get("cached_at", ""),
        "match_count": len(matches),
        "high_count": len(high),
        "low_count": len(low),
        "matches": matches,
        "columns": list(asset_df.columns),
    })


@app.route("/api/export", methods=["POST"])
def export_results():
    """매칭 결과를 Excel로 내보내기"""
    data = request.get_json()
    matches = data.get("matches", [])

    rows = []
    for m in matches:
        ad = m.get("asset_display", {})
        rc = m.get("recall", {})
        rows.append({
            "자산번호":      ad.get("자산번호", ""),
            "자산_제품명":   ad.get("제품명", ""),
            "자산_업체명":   ad.get("업체명", ""),
            "자산_모델명":   ad.get("모델명", ""),
            "자산_로트번호": ad.get("로트번호", ""),
            "매칭등급":      m.get("level", ""),
            "매칭점수":      m.get("score", ""),
            "매칭근거":      ", ".join(m.get("reasons", [])),
            "회수_품목명":   rc.get("품목명", "") or rc.get("제품명", ""),
            "회수_업체명":   rc.get("업체명", ""),
            "회수_모델명":   rc.get("모델명", ""),
            "회수_허가번호": rc.get("허가번호", ""),
            "회수구분":      rc.get("회수구분", ""),
            "회수진행여부":  rc.get("회수진행여부", ""),
            "보고일":        rc.get("보고일", ""),
            "상세링크":      rc.get("상세링크", ""),
        })

    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="회수대상자산")
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="회수대상자산_결과.xlsx",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("DEBUG", "false").lower() == "true"
    app.run(host="0.0.0.0", port=port, debug=debug)
