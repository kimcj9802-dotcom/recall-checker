import io
import os
import threading
from datetime import datetime, timedelta

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file

from matcher import find_matches
from scraper import fetch_recall_data, fetch_recall_detail, get_cache_info, load_recalls_from_file

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB


def _filter_active_recalls(recalls: list, months: int = 2) -> list:
    """진행중이고 최근 N개월 이내인 회수 목록만 반환 (빈 레코드 제외)"""
    cutoff = datetime.now() - timedelta(days=months * 30)
    result = []
    for r in recalls:
        # 0) 품목명 + 업체명 둘 다 없으면 빈 레코드로 제외
        name = (r.get("품목명") or r.get("제품명", "")).strip()
        maker = r.get("업체명", "").strip()
        if (not name or name in ("-", "None", "nan")
                or not maker or maker in ("-", "None", "nan")):
            continue

        # 1) 회수진행여부: '진행' 포함 항목만 (값이 없으면 통과)
        status = r.get("회수진행여부", "").strip()
        if status and status != "-" and "진행" not in status:
            continue

        # 2) 보고일: 최근 N개월 이내 (날짜가 없으면 통과)
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
    첫 행이 비어 있는 Excel을 자동으로 처리.
    header=0 → 컬럼이 모두 Unnamed이면 header=1, 2 순서로 재시도.
    """
    for header_row in range(3):
        df = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
        # 의미 있는 컬럼명이 하나라도 있으면 사용
        if not all("Unnamed" in str(c) for c in df.columns):
            return df
    # 그래도 안 되면 기본값 반환
    return pd.read_excel(io.BytesIO(file_bytes))

# 백그라운드 갱신 상태
_refresh_state = {"running": False, "error": None}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/cache-info")
def cache_info():
    return jsonify({**get_cache_info(), "refresh_running": _refresh_state["running"]})


@app.route("/api/refresh-recalls", methods=["POST"])
def refresh_recalls():
    """회수목록 새로고침 (백그라운드 실행)"""
    if _refresh_state["running"]:
        return jsonify({"success": False, "error": "이미 갱신 중입니다."}), 409

    def _do_refresh():
        _refresh_state["running"] = True
        _refresh_state["error"] = None
        try:
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
        "running": _refresh_state["running"],
        "error": _refresh_state["error"],
        **get_cache_info(),
    })


@app.route("/api/recalls")
def get_recalls():
    """캐시된 회수목록 반환 (진행중 + 최근 2개월 필터)"""
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
