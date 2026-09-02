# -*- coding: utf-8 -*-
"""
MFDS(식품의약품안전처) 의약품 허가사항 조회 모듈.

엔드포인트와 필드명은 개발 명세서 4-1장을 그대로 사용한다(추측 금지).
- 목록 조회: http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07/getDrugPrdtPrmsnInq07
- 상세 조회: https://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07/getDrugPrdtPrmsnDtlInq06

EE_DOC_DATA / UD_DOC_DATA / NB_DOC_DATA 는 JSON 안에 들어 있는 중첩 XML 문자열이다.
ElementTree 로 파싱하고, NB_DOC_DATA 는 섹션 타이틀 키워드로 분리한다(명세서 4-1장 키워드표).
"""
import re
import requests
import xml.etree.ElementTree as ET

LIST_URL = "http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07/getDrugPrdtPrmsnInq07"
DETAIL_URL = "https://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07/getDrugPrdtPrmsnDtlInq06"
LIST_ROWS = 500          # 명세서 권장값
MAX_PAGES = 10           # 최대 5,000건까지 전체 페이지를 돌며 키워드 필터링

# 명세서 4-1장 상세 응답 필드 매핑 (내부 필드명 → MFDS 원본 필드)
FIELD_MAP = [
    ("성분명",                 "MAIN_ITEM_INGR"),
    ("영문제품명",             "ITEM_ENG_NAME"),
    ("주성분영문명",           "MAIN_INGR_ENG"),
    ("전문일반구분",           "ETC_OTC_CODE"),
    ("ATC코드",                "ATC_CODE"),
    ("원료약품및분량(함량)",   "MATERIAL_NAME"),
    ("효능효과",               "EE_DOC_DATA"),
    ("용법용량",               "UD_DOC_DATA"),
    ("사용상주의사항 전체",    "NB_DOC_DATA"),
    ("포장단위",               "PACK_UNIT"),
    ("유효기간",               "VALID_TERM"),
    ("보관정보",               "STORAGE_METHOD"),
    ("성상",                   "CHART"),
    ("바코드(표준코드)",       "BAR_CODE"),
    ("EDI코드",                "EDI_CODE"),
]

# 명세서 4-1장 NB_DOC_DATA 섹션 분리 키워드표 (그대로 사용)
NB_SECTION_KEYWORDS = {
    "소아_고령자투여":      ["소아에 대한 투여", "소아투여", "고령자에 대한 투여", "고령자투여"],
    "임부_수유부투여":      ["임부 및 수유부에 대한 투여", "가임여성 및 남성에 대한 투여", "임부에 대한 투여", "수유부에 대한 투여"],
    "금기사항":            ["다음 환자에게는 투여하지 말 것", "투여하지 말 것", "금기"],
    "신중투여":            ["다음 환자에는 신중히 투여할 것", "신중히 투여"],
    "상호작용":            ["상호작용"],
    "이상반응":            ["이상반응", "이상 반응"],
    "과량투여처치":        ["과량투여시의 처치", "과량투여", "과량 투여"],
    "적용상의주의사항":    ["적용상의 주의", "적용상 주의"],
    "보관_취급주의사항":   ["보관 및 취급상의 주의사항", "보관 및 취급상의 주의"],
}

NB_SECTION_HUMAN = {
    "소아_고령자투여": "소아·고령자 투여",
    "임부_수유부투여": "임부·수유부 투여",
    "금기사항": "금기사항",
    "신중투여": "신중투여",
    "상호작용": "상호작용",
    "이상반응": "이상반응",
    "과량투여처치": "과량투여 시 처치",
    "적용상의주의사항": "적용상의 주의사항",
    "보관_취급주의사항": "보관 및 취급상의 주의사항",
    "기타": "기타 주의사항",
}


class MfdsApiError(Exception):
    """MFDS API 호출/응답 관련 오류 (사용자에게 그대로 노출)"""


def _get_body(data):
    """data.go.kr JSON 봉투에서 body 추출 (header/body 또는 response/header/body 모두 지원)."""
    body = data.get("body") if isinstance(data, dict) else None
    if body is None and isinstance(data, dict) and isinstance(data.get("response"), dict):
        body = data["response"].get("body")
    return body or {}


def _extract_items(body):
    """body에서 items 리스트 추출 (items 가 dict{item:[...]} 인 변형에도 대응)."""
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


def _get_json(url, params, timeout=30):
    """data.go.kr 호출 공통. 인증/서버 오류는 예외로 변환."""
    try:
        resp = requests.get(url, params=params, timeout=timeout)
    except requests.RequestException as exc:
        raise MfdsApiError(f"MFDS 서버에 연결하지 못했습니다: {exc}")
    try:
        data = resp.json()
    except ValueError:
        m = re.search(r"<resultMsg>([^<]+)</resultMsg>", resp.text)
        msg = m.group(1) if m else resp.text[:200]
        raise MfdsApiError(f"MFDS 응답을 해석하지 못했습니다(인증키 오류 또는 서비스 점검 가능): {msg}")
    if not isinstance(data, dict):
        raise MfdsApiError("MFDS 응답 형식이 올바르지 않습니다.")
    header = data.get("header") or {}
    code = header.get("resultCode")
    if str(code) not in ("00", "0", "200"):
        msg = header.get("resultMsg") or f"resultCode={code}"
        raise MfdsApiError(f"MFDS API 오류: {msg}")
    return data


def _normalize_title(text):
    """섹션 제목 정규화: 앞 번호 '1.'/'1)', 괄호, 공백, 중점 제거."""
    t = str(text or "").strip()
    t = re.sub(r"^[0-9]+[\.\)]?\s*", "", t)                 # "1." "2)"
    t = re.sub(r"^[\(（][^)）]*[\)）]\s*", "", t)            # "(공고번호...)"
    t = re.sub(r"[\s·•]+", "", t)                            # 공백/중점 제거
    return t


def _elem_text(elem):
    """요소의 모든 하위 텍스트를 이어붙인다 (중첩 태그 대응)."""
    return "".join(elem.itertext()).strip()


def parse_doc_xml(xml_str):
    """
    EE_DOC_DATA / UD_DOC_DATA / NB_DOC_DATA(중첩 XML 문자열)을
    [{"title": ..., "text": ...}, ...] 형태로 파싱한다.
    XML로 해석되지 않으면 문자열 전체를 하나의 (title='', text=원문) 항목으로 돌려준다.
    """
    raw = str(xml_str or "").strip()
    if not raw:
        return []
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return [{"title": "", "text": raw}]
    pairs = []
    cur_title, cur_texts = "", []
    for elem in root.iter():
        tag = (elem.tag or "").upper()
        text = _elem_text(elem)
        if "TITLE" in tag and text:
            if cur_texts:
                pairs.append({"title": cur_title, "text": " ".join(cur_texts)})
                cur_texts = []
            cur_title = text
        elif "TEXT" in tag and text:
            cur_texts.append(text)
    if cur_texts or cur_title:
        pairs.append({"title": cur_title, "text": " ".join(cur_texts)})
    if not pairs:
        entire = re.sub(r"[ \t]+", " ", ET.tostring(root, encoding="unicode", method="text")).strip()
        if entire:
            pairs = [{"title": "", "text": entire}]
    return pairs


def docs_to_text(pairs):
    """파싱된 (title, text) 목록을 줄바꿈으로 이어붙인다. 원문을 요약하지 않는다."""
    out = []
    for p in pairs:
        line = (p.get("title") or "").strip()
        txt = (p.get("text") or "").strip()
        if line and txt:
            out.append(f"{line} {txt}")
        elif txt:
            out.append(txt)
        elif line:
            out.append(line)
    return "\n".join(out)


def split_nb_sections(pairs):
    """
    NB_DOC_DATA 파싱 결과를 섹션 키워드표에 따라 분리한다.
    반환: {섹션키: 원문텍스트(요약 아님)} — 키워드에 걸리지 않으면 '기타'로.
    """
    sections = {k: [] for k in NB_SECTION_KEYWORDS}
    sections["기타"] = []
    for p in pairs:
        norm = _normalize_title(p.get("title", ""))
        matched_key = None
        for key, keywords in NB_SECTION_KEYWORDS.items():
            for kw in keywords:
                if _normalize_title(kw) in norm:
                    matched_key = key
                    break
            if matched_key:
                break
        sections[matched_key or "기타"].append(p)
    return {k: docs_to_text(v) for k, v in sections.items()}


def search_products(keyword, service_key):
    """
    품목명/제조사명 부분 일치 검색.
    CANCEL_NAME 이 '정상'인 품목만 반환한다(명세서 4-1장).
    """
    if not service_key:
        raise MfdsApiError("MFDS 인증키가 설정되지 않았습니다. Streamlit Secrets에 MFDS_API_KEY(또는 DATA_GO_KOR_API_KEY)를 설정하세요.")
    collected, page = [], 1
    while page <= MAX_PAGES:
        data = _get_json(LIST_URL, {
            "serviceKey": service_key,
            "pageNo": page,
            "numOfRows": LIST_ROWS,
            "type": "json",
        })
        body = _get_body(data)
        items = _extract_items(body)
        collected.extend(items)
        try:
            total = int(body.get("totalCount") or 0)
        except (TypeError, ValueError):
            total = 0
        if not items or page * LIST_ROWS >= total:
            break
        page += 1
    kw = (keyword or "").strip().lower()
    rows = []
    for it in collected:
        cancel = str(it.get("CANCEL_NAME") or "").strip()
        if cancel not in ("", "정상"):
            continue
        name = str(it.get("ITEM_NAME") or "").strip()
        entp = str(it.get("ENTP_NAME") or "").strip()
        if kw and kw not in name.lower() and kw not in entp.lower():
            continue
        rows.append({
            "ITEM_SEQ": str(it.get("ITEM_SEQ") or "").strip(),
            "ITEM_NAME": name,
            "ENTP_NAME": entp,
            "ITEM_PERMIT_DATE": str(it.get("ITEM_PERMIT_DATE") or "").strip(),
        })
    rows.sort(key=lambda r: r["ITEM_NAME"])
    return rows


def get_product_detail(item_seq, service_key):
    """
    품목기준코드(ITEM_SEQ)로 상세 허가사항 조회.
    반환 dict 는 명세서 4-1장 매핑표의 내부 필드명을 키로 사용한다.
    """
    if not service_key:
        raise MfdsApiError("MFDS 인증키가 설정되지 않았습니다. Streamlit Secrets에 MFDS_API_KEY(또는 DATA_GO_KOR_API_KEY)를 설정하세요.")
    data = _get_json(DETAIL_URL, {
        "serviceKey": service_key,
        "item_seq": str(item_seq).strip(),
        "type": "json",
    })
    items = _extract_items(_get_body(data))
    if not items:
        raise MfdsApiError(f"상세 허가사항이 없습니다(품목기준코드: {item_seq}).")
    raw = items[0]
    out = {}
    for internal, mfds_field in FIELD_MAP:
        out[internal] = str(raw.get(mfds_field) or "").strip()
    # 성분명의 대괄호 코드([A12345]) 제거 (명세서 4-1장)
    out["성분명"] = re.sub(r"\[[^\]]*\]", "", out["성분명"]).strip()
    out["제품명"] = str(raw.get("ITEM_NAME") or "").strip()
    out["제조판매사"] = str(raw.get("ENTP_NAME") or "").strip()
    out["품목기준코드"] = str(raw.get("ITEM_SEQ") or "").strip()
    # 제형: 상세 응답에 FORM_CODE 가 있는 경우에만 사용(없으면 빈 문자열 → 검증 시 '확인불가')
    out["제형"] = str(raw.get("FORM_CODE") or "").strip()
    # 중첩 XML 파싱 (원문 유지, 요약 금지)
    out["효능효과"] = docs_to_text(parse_doc_xml(out["효능효과"]))
    out["용법용량"] = docs_to_text(parse_doc_xml(out["용법용량"]))
    out["사용상주의사항"] = split_nb_sections(parse_doc_xml(out["사용상주의사항 전체"]))
    return out


def detail_to_raw_text(detail):
    """상세 조회 결과를 원문 그대로의 텍스트로 변환(표시/프롬프트용). 요약하지 않는다."""
    if not detail:
        return "(데이터 없음)"
    if detail.get("error"):
        return f"MFDS 조회 실패: {detail['error']}"
    lines = []
    for internal, _ in FIELD_MAP:
        if internal == "사용상주의사항 전체":
            continue
        val = detail.get(internal)
        if val:
            lines.append(f"{internal}: {val}")
    nb = detail.get("사용상주의사항") or {}
    for key, human in NB_SECTION_HUMAN.items():
        if nb.get(key):
            lines.append(f"사용상주의사항 - {human}:\n{nb[key]}")
    return "\n\n".join(lines)
