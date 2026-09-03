# -*- coding: utf-8 -*-
"""
HIRA(건강보험심사평가원) 약가정보 조회 모듈.

엔드포인트(HTTPS 고정 — 사용자 제공 · 명세서 4-2장):
- 약가 조회: https://apis.data.go.kr/B551182/dgamtCrtrInfoService1.2/getDgamtList
- 파라미터: serviceKey, type=json, 그리고 mdsCd(품목코드) 또는 itmNm(품목명)
- 응답 필드: mdsCd, itmNm, mnfEntpNm, mxCprc(상한금액), payTpNm("삭제" 행 제외), meftDivNo(약효분류)

v4.1 변경점:
- 전체 약가 목록을 1회 적재(fetch_all_drug_prices)해 data/cache/hira_prices.json 저장.
  이후 매칭은 캐시 로컬 검색(8자리 바코드 규칙), 캐시가 없을 때만 제품당 1회 API 폴백.
- serviceKey unquote 1회 정규화(이중 인코딩 방지) + 연결 재시도 3회.
"""
import re
import time
import urllib.parse

import requests

HIRA_URL = "https://apis.data.go.kr/B551182/dgamtCrtrInfoService1.2/getDgamtList"
LIST_ROWS = 500
MAX_FETCH_PAGES = 400    # 전체 적재 상한 (500 × 400 = 200,000건)


class HiraApiError(Exception):
    """HIRA API 호출/응답 관련 오류 (사용자에게 그대로 노출)"""


def prep_service_key(service_key):
    """data.go.kr 인증키 이중 인코딩 방지 — URL 인코딩된 키를 1회 unquote."""
    key = str(service_key or "").strip()
    if not key:
        return key
    return urllib.parse.unquote(key)


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


def _get_json(params, timeout=30, retries=3):
    params = dict(params)
    params["serviceKey"] = prep_service_key(params.get("serviceKey", ""))
    last_exc = None
    resp = None
    for attempt in range(max(1, retries)):
        try:
            resp = requests.get(HIRA_URL, params=params, timeout=timeout)
            break
        except requests.RequestException as exc:
            last_exc = exc
            if attempt < retries - 1:
                time.sleep(1.5 * (attempt + 1))
    if resp is None:
        raise HiraApiError(f"HIRA 서버에 연결하지 못했습니다 (HTTPS, {retries}회 시도): {last_exc}")
    try:
        data = resp.json()
    except ValueError:
        m = re.search(r"<resultMsg>([^<]+)</resultMsg>", resp.text)
        msg = m.group(1) if m else resp.text[:200]
        raise HiraApiError(f"HIRA 응답을 해석하지 못했습니다(인증키 오류 또는 서비스 점검 가능): {msg}")
    if not isinstance(data, dict):
        raise HiraApiError("HIRA 응답 형식이 올바르지 않습니다.")
    # data.go.kr 오류 봉투: {"OpenAPI_ServiceResponse":{"cmmMsgHeader":{...}}}
    # (키 미등록·만료 시 HTTP 403 과 함께 이 형식으로 온다)
    svc = data.get("OpenAPI_ServiceResponse")
    if isinstance(svc, dict):
        hdr = svc.get("cmmMsgHeader") or {}
        err = hdr.get("errMsg") or ""
        auth = hdr.get("returnAuthMsg") or ""
        if err or auth:
            raise HiraApiError(f"HIRA API 오류: {auth or err} ({err})")
    header = data.get("header") or {}
    code = header.get("resultCode")
    if str(code) not in ("00", "0", "200"):
        msg = header.get("resultMsg") or f"resultCode={code}"
        raise HiraApiError(f"HIRA API 오류: {msg}")
    return data


def _map_row(it):
    """명세서 4-2장 필드만 남기고 payTpNm == '삭제' 행 제외."""
    pay = str(it.get("payTpNm") or "").strip()
    if pay == "삭제":
        return None
    return {
        "mdsCd": str(it.get("mdsCd") or "").strip(),
        "itmNm": str(it.get("itmNm") or "").strip(),
        "mnfEntpNm": str(it.get("mnfEntpNm") or "").strip(),
        "mxCprc": str(it.get("mxCprc") or "").strip(),
        "payTpNm": pay,
        "meftDivNo": str(it.get("meftDivNo") or "").strip(),
    }


def fetch_all_drug_prices(service_key, progress_cb=None, max_pages=MAX_FETCH_PAGES):
    """
    전체 약가 목록을 페이지네이션 끝까지 수집한다('삭제' 행 제외).
    progress_cb(진행률 0~1, 누적건수) — UI 진행률 표시용 콜백.
    이후 매칭은 이 결과를 JSON 캐시로 저장한 뒤 로컬에서 수행한다.
    """
    if not service_key:
        raise HiraApiError("HIRA 인증키가 필요합니다. 사이드바에 입력하거나 Secrets에 설정하세요.")
    collected, page, total = [], 1, None
    while page <= max_pages:
        data = _get_json({
            "serviceKey": service_key,
            "pageNo": page,
            "numOfRows": LIST_ROWS,
            "type": "json",
        })
        body = _get_body(data)
        items = _extract_items(body)
        for it in items:
            mapped = _map_row(it)
            if mapped:
                collected.append(mapped)
        try:
            total = int(body.get("totalCount") or 0)
        except (TypeError, ValueError):
            total = None
        if progress_cb:
            frac = min(1.0, len(collected) / total) if total else min(1.0, page / max_pages)
            progress_cb(frac, len(collected))
        if not items:
            break
        if len(items) < LIST_ROWS:
            break
        if total and page * LIST_ROWS >= total:
            break
        page += 1
    return collected


def search_by_mdscd(mdscd, service_key):
    """품목코드(mdsCd)로 약가 조회 (캐시 미적재 시 폴백용)."""
    if not service_key:
        raise HiraApiError("HIRA 인증키가 필요합니다. 사이드바에 입력하거나 Secrets에 설정하세요.")
    data = _get_json({
        "serviceKey": service_key,
        "type": "json",
        "mdsCd": str(mdscd).strip(),
        "numOfRows": 100,
        "pageNo": 1,
    })
    return [r for r in (_map_row(it) for it in _extract_items(_get_body(data))) if r]


def search_by_name(name, service_key):
    """품목명(itmNm)으로 약가 조회 (캐시 미적재 시 폴백용)."""
    if not service_key:
        raise HiraApiError("HIRA 인증키가 필요합니다. 사이드바에 입력하거나 Secrets에 설정하세요.")
    data = _get_json({
        "serviceKey": service_key,
        "type": "json",
        "itmNm": str(name).strip(),
        "numOfRows": 100,
        "pageNo": 1,
    })
    return [r for r in (_map_row(it) for it in _extract_items(_get_body(data))) if r]


def search_for_product(detail, service_key):
    """
    (캐시 미적재 시 폴백) 명세서 5장 규칙:
    1) MFDS 바코드의 4~11번째 8자리로 우선 조회
    2) 결과가 없으면 품목명 검색으로 폴백
    두 경로 결과를 합쳐 반환 — drug_matcher 가 8자리 규칙으로 최종 매칭한다.
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
                rows = []
    if not rows:
        name = (detail.get("제품명") or "").strip()
        if name:
            rows.extend(search_by_name(name, service_key))
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
