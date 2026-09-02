# -*- coding: utf-8 -*-
"""
Python 1차 규칙 검증 모듈 — 명세서 7장.

기계적으로 100% 확실한 것만 판정한다:
1. 기본정보 일치 (제품명/성분명/제조판매사/제형 — 정규화 후 정확 비교)
2. 숫자·단위 (정규식 토큰 추출 후 동일 단위의 숫자 불일치 즉시 수정필요)
3. 약가 차이율 (신청의약품 - 비교의약품 최저가) / 비교의약품 최저가 × 100
그 외 서술형 항목은 절대 자동 판정하지 않고 "🟠 Claude 확인 필요" 로 넘긴다.
"""
import re

STATUS_OK = "✅ 일치"
STATUS_FIX = "⚠ 수정필요"
STATUS_CLAUDE = "🟠 Claude 확인 필요"
STATUS_UNKNOWN = "❌ 확인불가"

BASIC_FIELDS = ("제품명", "성분명", "제조판매사", "제형")
DESCRIPTIVE_FIELDS = (
    "효능효과", "용법용량", "이상반응", "금기사항", "신중투여", "상호작용",
    "소아_고령자투여", "임부_수유부투여", "과량투여처치", "보관_취급주의사항",
    "사용상주의사항", "포장단위", "유효기간", "기타",
)

_UNIT_TOKEN = re.compile(r"(\d+(?:[,.]\d+)?)\s*([a-zA-Z가-힣%]+)")
_SKIP_UNITS = {"형", "상", "번"}  # '제1형' 등 노이즈 유닛


def _norm_compare(text):
    s = str(text or "").strip()
    s = re.sub(r"\(주\)|（주）|주식회사|㈜|\[주\]", "", s)
    s = re.sub(r"[^0-9a-zA-Z가-힣]", "", s).lower()
    return s


def check_basic_info(value_a, value_b):
    """
    기본정보 정규화 비교. (주) 등 법인 표기는 무시.
    반환: (status, reason)
    """
    if value_b is None or not str(value_b).strip():
        return STATUS_UNKNOWN, "원문(MFDS)에 해당 정보가 없어 비교할 수 없습니다."
    a, b = _norm_compare(value_a), _norm_compare(value_b)
    if a and a == b:
        return STATUS_OK, f"표기 일치 (MFDS 원문: {str(value_b).strip()[:60]})"
    if not a:
        return STATUS_UNKNOWN, "비교표에 값이 비어 있습니다."
    return STATUS_FIX, f"표기 불일치 — 비교표: 「{str(value_a).strip()[:60]}」 / 원문: 「{str(value_b).strip()[:60]}」"


def extract_number_units(text):
    """
    '숫자+단위' 토큰 추출: [("500","mg"), ("1","정"), ...] 형태.
    단위가 없거나 노이즈 단위(형/상/번)인 순수 숫자는 제외.
    """
    tokens = []
    for m in _UNIT_TOKEN.finditer(str(text or "")):
        num_raw, unit = m.group(1), m.group(2).lower().strip()
        unit = re.sub(r"[^a-z가-힣%]", "", unit)
        if not unit or unit in _SKIP_UNITS:
            continue
        try:
            num = float(num_raw.replace(",", ""))
        except ValueError:
            continue
        tokens.append((num, unit))
    return tokens


def compare_number_units(value_a, value_b):
    """
    양쪽 '숫자+단위' 토큰을 단위별로 비교.
    같은 단위가 양쪽에 있고 숫자가 다른 토큰이 있으면 수정필요 사유 리스트 반환.
    """
    if value_b is None or not str(value_b).strip():
        return None  # 원문 없음 → 참조 불가 (호출부에서 확인불가 처리)
    ta = extract_number_units(value_a)
    tb = extract_number_units(value_b)
    if not ta and not tb:
        return []
    issues = []
    for num_a, unit_a in ta:
        same_unit = [(n, u) for n, u in tb if u == unit_a]
        if not same_unit:
            issues.append(f"비교표의 '{_fmt(num_a, unit_a)}'가 원문에 없음(누락/추가 의심)")
            continue
        if all(n != num_a for n, _ in same_unit):
            issues.append(f"단위 {unit_a}: 비교표 {_fmt(num_a, unit_a)} vs 원문 {', '.join(_fmt(n, u) for n, u in same_unit)}")
    return issues


def _fmt(num, unit):
    return f"{num:g}{unit}"


def check_numeric_field(value_a, value_b):
    """
    서술형 항목의 숫자·단위만 기계 검사.
    반환: (status, reason) — 숫자 불일치 없으면 Claude 확인 필요로 넘긴다(명세서 7장).
    """
    if value_b is None or not str(value_b).strip():
        return STATUS_UNKNOWN, "원문(MFDS)이 없어 비교할 수 없습니다."
    issues = compare_number_units(value_a, value_b)
    if issues:
        return STATUS_FIX, "숫자·단위 불일치: " + " / ".join(issues[:3])
    return STATUS_CLAUDE, "숫자·단위는 일치. 의미 비교는 Claude 웹에서 검증하세요."


def parse_price(value):
    """비교표 약가 셀에서 숫자 금액 추출. 없으면 None."""
    if value is None:
        return None
    m = re.search(r"(\d[\d,]*(?:\.\d+)?)", str(value).replace(",", ""))
    try:
        return float(m.group(1).replace(",", ""))
    except (AttributeError, ValueError):
        return None


def check_price(value_a, hira_price):
    """
    비교표 약가 vs HIRA 상한금액 비교.
    반환: (status, reason, diff_pct_or_None)
    """
    if hira_price is None:
        return STATUS_UNKNOWN, "HIRA 약가정보가 없어 비교할 수 없습니다.", None
    pa = parse_price(value_a)
    if pa is None:
        return STATUS_UNKNOWN, "비교표의 약가 셀에서 금액을 해석하지 못했습니다.", None
    diff = (pa - hira_price) / hira_price * 100 if hira_price else None
    if abs(diff) < 0.005:
        return STATUS_OK, f"약가 일치 (비교표 {pa:,.0f}원 = HIRA {hira_price:,.0f}원)", diff
    return STATUS_FIX, f"약가 불일치 — 비교표 {pa:,.0f}원 vs HIRA {hira_price:,.0f}원 (차이 {diff:+.1f}%)", diff


def price_diff_percent(applicant_price, comparator_min_price):
    """
    명세서 7장 약가 차이율:
    (신청의약품가격 - 비교의약품최저가) / 비교의약품최저가 × 100
    """
    if applicant_price is None or comparator_min_price in (None, 0):
        return None
    return (applicant_price - comparator_min_price) / comparator_min_price * 100


def evaluate_pair(pair, reference_text, hira_price=None):
    """
    공통 셀 1건에 대한 1차 규칙 검증.
    pair: 명세서 8장 스키마 dict
    reference_text: 해당 제품·항목의 MFDS 원문 텍스트 (없으면 None)
    hira_price: 해당 제품의 HIRA 상한금액 (약가 항목만 사용)
    반환: (status, reason)
    """
    field = pair.get("field", "")
    value = pair.get("value", "")
    if field in BASIC_FIELDS:
        return check_basic_info(value, reference_text)
    if field == "약가":
        status, reason, _ = check_price(value, hira_price)
        return status, reason
    if field in DESCRIPTIVE_FIELDS or field:
        return check_numeric_field(value, reference_text)
    return STATUS_CLAUDE, "의미 비교는 Claude 웹에서 검증하세요."
