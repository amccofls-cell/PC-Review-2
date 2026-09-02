# -*- coding: utf-8 -*-
"""
비교표 공통 스키마 변환 모듈.

PPTX / XLSX / 복사·붙여넣기 세 입력 경로를 명세서 8장의 공통 셀 스키마로 정규화한다.
{
  "slide": 3, "table_index": 1, "row": 4, "column": 2,
  "field": "효능·효과", "product": "신청의약품", "value": "..."
}
"""
import re

# 필드 키워드표 (우선순위 순 — 먼저 등장하는 키워드가 우선 매칭)
FIELD_ORDER = [
    ("효능효과",              ["효능효과", "효능·효과", "효능 및 효과", "적응증", "효능", "효과"]),
    ("용법용량",              ["용법용량", "용법·용량", "용법 및 용량", "투여용량", "용법", "용량"]),
    ("이상반응",              ["이상반응", "이상 반응", "부작용"]),
    ("금기사항",              ["금기사항", "금기", "투여하지 말 것"]),
    ("신중투여",              ["신중투여", "신중히 투여"]),
    ("상호작용",              ["상호작용", "약물상호작용"]),
    ("임부_수유부투여",       ["임부 및 수유부", "수유부", "임부", "임신", "가임여성"]),
    ("소아_고령자투여",       ["소아", "고령자"]),
    ("과량투여처치",          ["과량투여", "과다투여", "과량"]),
    ("보관_취급주의사항",     ["보관 및 취급", "보관", "저장방법"]),
    ("사용상주의사항",        ["사용상주의사항", "사용상 주의사항", "주의사항"]),
    ("제품명",                ["제품명", "품목명", "약품명"]),
    ("성분명",                ["성분명", "주성분", "성분"]),
    ("제조판매사",            ["제조판매사", "제조사", "판매사", "위탁제조", "제조·판매"]),
    ("제형",                  ["제형"]),
    ("함량",                  ["함량", "분량", "원료약품의 분량", "원료약품"]),
    ("포장단위",              ["포장단위", "포장"]),
    ("유효기간",              ["유효기간"]),
    ("약가",                  ["약가", "상한금액", "보험약가", "급여기준"]),
    ("ATC코드",               ["atc"]),
]

# 헤더로 오인하지 않을 일반 라벨
GENERIC_LABELS = {
    "", "구분", "항목", "항목명", "내용", "비고", "제품", "제품명", "약품", "품목",
    "품명", "단위", "수량", "순번", "번호", "no", "no.", "remarks", "구성",
}

ORIENT_ROWS_ARE_ITEMS = "rows_are_items"   # 행 = 항목(필드), 열 = 제품
ORIENT_COLS_ARE_ITEMS = "cols_are_items"   # 열 = 항목(필드), 행 = 제품


def normalize_cell(text):
    """셀 텍스트 정규화: 줄바꿈·공백·전각 공백 제거 후 단일 공백 연결."""
    s = str(text or "").strip()
    s = s.replace("\u3000", " ").replace("\xa0", " ")
    return " ".join(s.split())


def _norm_compact(text):
    """비교용 압축 정규화: 공백/괄호/특수문자/법인 표기 제거."""
    s = normalize_cell(text)
    s = re.sub(r"\(주\)|（주）|주식회사|㈜|\[주\]", "", s)
    s = re.sub(r"[^0-9a-zA-Z가-힣]", "", s).lower()
    return s


def lookup_field(label):
    """라벨이 알려진 필드면 정규 필드명, 아니면 None."""
    n = _norm_compact(label)
    if not n:
        return None
    for key, keywords in FIELD_ORDER:
        for kw in keywords:
            if _norm_compact(kw) in n:
                return key
    return None


def _is_generic(label):
    return _norm_compact(label) in {_norm_compact(g) for g in GENERIC_LABELS}


def guess_orientation(rows):
    """
    표 방향 자동 추정.
    - 첫 행(헤더)에 필드명이 많으면 → cols_are_items (열 = 항목)
    - 첫 열에 필드명이 많으면 → rows_are_items (행 = 항목)
    애매하면 rows_are_items 를 기본값으로 돌려준다(화면에서 사용자가 교체 가능).
    """
    if not rows or len(rows) < 2 or len(rows[0]) < 2:
        return ORIENT_ROWS_ARE_ITEMS
    hdr_hits = sum(1 for c in rows[0] if lookup_field(c.get("text", "")))
    col_hits = sum(1 for r in rows[1:] if lookup_field(r[0].get("text", "")))
    hdr_products = sum(
        1 for c in rows[0]
        if not lookup_field(c.get("text", "")) and not _is_generic(c.get("text", ""))
    )
    if col_hits > hdr_hits:
        return ORIENT_ROWS_ARE_ITEMS
    if hdr_hits > col_hits:
        return ORIENT_COLS_ARE_ITEMS
    return ORIENT_ROWS_ARE_ITEMS if hdr_products >= 2 else ORIENT_COLS_ARE_ITEMS


def _product_candidates(cells):
    """셀 목록에서 제품명 후보 (일반 라벨·필드명 제외) 추출."""
    out = []
    for idx, cell in enumerate(cells):
        label = normalize_cell(cell.get("text", ""))
        if not label or _is_generic(label) or lookup_field(label):
            continue
        out.append((idx, label))
    return out


def _cell_value_at(rows, r, c):
    if 0 <= r < len(rows) and 0 <= c < len(rows[r]):
        return rows[r][c].get("text", "")
    return ""


def to_field_product_pairs(table, orientation):
    """
    파서 출력(rows)을 공통 셀 스키마(명세서 8장) 리스트로 변환한다.
    """
    rows = table.get("rows", [])
    slide = table.get("slide", 0)
    tindex = table.get("table_index", 0)
    pairs = []
    if orientation == ORIENT_ROWS_ARE_ITEMS:
        if not rows:
            return pairs
        products = _product_candidates(rows[0])
        for ri in range(1, len(rows)):
            field_label = normalize_cell(rows[ri][0].get("text", "")) if rows[ri] else ""
            if not field_label:
                continue
            field = lookup_field(field_label) or field_label
            for ci, plabel in products:
                value = _cell_value_at(rows, ri, ci)
                pairs.append({
                    "slide": slide, "table_index": tindex,
                    "row": ri + 1, "column": ci + 1,
                    "field": field, "product": plabel, "value": value,
                })
    else:  # cols_are_items
        if not rows:
            return pairs
        products = [(ri, normalize_cell(rows[ri][0].get("text", "")))
                    for ri in range(1, len(rows))
                    if normalize_cell(rows[ri][0].get("text", ""))
                    and not _is_generic(rows[ri][0].get("text", ""))]
        for ci in range(1, len(rows[0])):
            field_label = normalize_cell(rows[0][ci].get("text", ""))
            if not field_label:
                continue
            field = lookup_field(field_label) or field_label
            for ri, plabel in products:
                value = _cell_value_at(rows, ri, ci)
                pairs.append({
                    "slide": slide, "table_index": tindex,
                    "row": ri + 1, "column": ci + 1,
                    "field": field, "product": plabel, "value": value,
                })
    return pairs
