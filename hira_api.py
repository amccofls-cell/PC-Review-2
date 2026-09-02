# -*- coding: utf-8 -*-
"""
HIRA(건강보험심사평가원) 약가정보 조회 모듈.

엔드포인트와 필드명은 개발 명세서 4-2장을 그대로 사용한다(추측 금지).
- 약가 조회: https://apis.data.go.kr/B551182/dgamtCrtrInfoService1.2/getDgamtList
- 파라미터: serviceKey, type=json, 그리고 mdsCd(품목코드) 또는 itmNm(품목명)
- 응답 필드: mdsCd, itmNm, mnfEntpNm, mxCprc(상한금액), payTpNm("삭제" 행 제외), meftDivNo(약효분류)
"""
import re
import requests

HIRA_URL = "https://apis.data.go.kr/B551182/dgamtCrtrInfoService1.2/getDgamtList"


class HiraApiError(Exception):
    """HIRA API 호출/응답 관련 오류 (사용자에게 그대로 노출)"""


def _get_body(data):
    body = data.get("body") if isinstance(data, dict) else None
    if body is None and isinstance(data, dict) and isinstance(data.get("response"), dict):
        body = data["response"].get("body")
    return body or {}


def _extract_items(body):
    if not body:
        return []
    items = body.get("items")
    if items is None:
        return []
    if isinstance(items, list):
        return items
    if isinstance(items, dict):
        item = items.get("item") or items.get("items") or []
        return item if isinstance(item, list) else [item]
    return []


def _get_json(params, timeout=30):
    try:
        resp = requests.get(HIRA_URL, params=params, timeout=timeout)
    except requests.RequestException as exc:
        raise HiraApiError(f"HIRA 서버에 연결하지 못했습니다: {exc}")
    try:
        data = resp.json()
    except ValueError:
        m = re.search(r"<resultMsg>([^<]+)</resultMsg>", resp.text)
        msg = m.group(1) if m else resp.text[:200]
        raise HiraApiError(f"HIRA 응답을 해석하지 못했습니다(인증키 오류 또는 서비스 점검 가능): {msg}")
    if not isinstance(data, dict):
        raise HiraApiError("HIRA 응답 형식이 올바르지 않습니다.")
    header = data.get("header") or {}
    code = header.get("resultCode")
    if str(code) not in ("00", "0", "200"):
        msg = header.get("resultMsg") or f"resultCode={code}"
        raise HiraApiError(f"HIRA API 오류: {msg}")
    return data


def _filter_rows(data):
    """응답 items 에서 payTpNm == '삭제' 인 행을 제외하고 명세서 4-2장 필드만 남긴다."""
    body = _get_body(data)
    rows = []
    for it in _extract_items(body):
        pay = str(it.get("payTpNm") or "").strip()
        if pay == "삭제":
            continue
        rows.append({
            "mdsCd": str(it.get("mdsCd") or "").strip(),
            "itmNm": str(it.get("itmNm") or "").strip(),
            "mnfEntpNm": str(it.get("mnfEntpNm") or "").strip(),
            "mxCprc": str(it.get("mxCprc") or "").strip(),
            "payTpNm": pay,
            "meftDivNo": str(it.get("meftDivNo") or "").strip(),
        })
    return rows


def search_by_mdscd(mdscd, service_key):
    """품목코드(mdsCd)로 약가 조회."""
    if not service_key:
        raise HiraApiError("HIRA 인증키가 설정되지 않았습니다. Streamlit Secrets에 HIRA_API_KEY(또는 DATA_GO_KOR_API_KEY)를 설정하세요.")
    data = _get_json({
        "serviceKey": service_key,
        "type": "json",
        "mdsCd": str(mdscd).strip(),
        "numOfRows": 100,
        "pageNo": 1,
    })
    return _filter_rows(data)


def search_by_name(name, service_key):
    """품목명(itmNm)으로 약가 조회."""
    if not service_key:
        raise HiraApiError("HIRA 인증키가 설정되지 않았습니다. Streamlit Secrets에 HIRA_API_KEY(또는 DATA_GO_KOR_API_KEY)를 설정하세요.")
    data = _get_json({
        "serviceKey": service_key,
        "type": "json",
        "itmNm": str(name).strip(),
        "numOfRows": 100,
        "pageNo": 1,
    })
    return _filter_rows(data)


def search_for_product(detail, service_key):
    """
    명세서 5장 매칭 규칙에 따른 조회:
    1) MFDS 바코드의 4~11번째 8자리로 우선 조회 시도
    2) 결과가 없으면 품목명 검색으로 폴백
    두 경로의 결과를 합쳐 반환하면, drug_matcher 가 8자리 규칙으로 최종 매칭한다.
    """
    rows = []
    barcode = str(detail.get("바코드(표준코드)") or "")
    digits = re.sub(r"\D", "", barcode)
    if len(digits) >= 11:
        b8 = digits[3:11]
        if b8:
            try:
                rows.extend(search_by_mdscd(b8, service_key))
            except HiraApiError:
                rows = []  # mdsCd 검색 실패 시 품목명으로 폴백
    if not rows:
        name = (detail.get("제품명") or "").strip()
        if name:
            rows.extend(search_by_name(name, service_key))
    # 중복 제거 (같은 mdsCd 유지)
    seen, uniq = set(), []
    for r in rows:
        if r["mdsCd"] in seen:
            continue
        seen.add(r["mdsCd"])
        uniq.append(r)
    return uniq


def get_price(row):
    """상한금액(mxCprc)을 float 로 변환. 없으면 None."""
    try:
        return float(str(row.get("mxCprc") or "").replace(",", ""))
    except (TypeError, ValueError):
        return None
