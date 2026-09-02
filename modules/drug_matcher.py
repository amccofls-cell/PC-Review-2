# -*- coding: utf-8 -*-
"""
제품 식별/매칭 모듈 — 개발 명세서 5장 규칙 그대로 구현(추측 금지).

매칭 규칙:
1. MFDS BAR_CODE 의 4~11번째 8자리 숫자 == HIRA mdsCd 의 앞 8자리 숫자  (가장 신뢰도 높음)
   Python 슬라이싱: barcode_digits[3:11] == mdscd_digits[:8] (0-index, 총 13자리 바코드 가정)
2. 품목명 완전 일치 (공백/괄호/특수문자 정규화 후)
3. 품목명 부분 일치 — 반드시 "🟠 매칭 확인 필요" 상태로 표시하고 자동 확정하지 않는다.
일치하는 항목이 없으면 절대 임의로 "그나마 비슷한 것"을 자동 선택하지 않는다.
"""
import re

STATUS_BARCODE = "✅ 매칭 확정 (바코드 8자리 일치)"
STATUS_NAME_EXACT = "✅ 매칭 확정 (품목명 완전 일치)"
STATUS_NAME_PARTIAL = "🟠 매칭 확인 필요 (품목명 부분 일치)"
STATUS_FAIL = "❌ 매칭 실패"


def _norm_name(name):
    """품목명 정규화: (주)/주식회사 등 법인 표기, 공백, 괄호, 특수문자 제거 후 소문자."""
    s = str(name or "").strip().lower()
    s = re.sub(r"\(주\)|（주）|주식회사|㈜|\[주\]", "", s)
    s = re.sub(r"[^0-9a-z가-힣]", "", s)
    return s


def barcode_8digits(barcode):
    """바코드(표준코드)에서 4~11번째 8자리 숫자를 추출. 13자리 미만이면 None."""
    digits = re.sub(r"\D", "", str(barcode or ""))
    if len(digits) < 11:
        return None
    return digits[3:11]


def match_hira(detail, hira_rows):
    """
    명세서 5장 매칭 우선순위에 따라 최적 행 1건을 선택한다.
    반환: {"status": ..., "row": dict|None, "method": ...}
    """
    b8 = barcode_8digits(detail.get("바코드(표준코드)"))
    product_name = detail.get("제품명") or ""

    # 1순위: 바코드 8자리 == mdsCd 앞 8자리
    if b8:
        for row in hira_rows:
            md = re.sub(r"\D", "", str(row.get("mdsCd") or ""))
            if md[:8] == b8:
                return {"status": STATUS_BARCODE, "row": row, "method": "barcode"}
    # 2순위: 품목명 완전 일치 (정규화 후)
    pn = _norm_name(product_name)
    if pn:
        for row in hira_rows:
            if _norm_name(row.get("itmNm")) == pn:
                return {"status": STATUS_NAME_EXACT, "row": row, "method": "name_exact"}
    # 3순위: 품목명 부분 일치 — 자동 확정 금지, 상태만 표시
    if pn:
        for row in hira_rows:
            rn = _norm_name(row.get("itmNm"))
            if rn and (pn in rn or rn in pn):
                return {"status": STATUS_NAME_PARTIAL, "row": row, "method": "name_partial"}
    return {"status": STATUS_FAIL, "row": None, "method": None}


def match_result_summary(match):
    """매칭 상태 요약 문구."""
    if not match:
        return STATUS_FAIL
    return match.get("status") or STATUS_FAIL
