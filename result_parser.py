# -*- coding: utf-8 -*-
"""
Claude 웹 검증 결과 재붙여넣기 파싱 모듈 — 명세서 10장.

지원 형식:
1. 마크다운 표  | 제품 | 항목 | 판단 | 이유 | 원문 근거 |
2. JSON        {"results": [{"제품": ..., "항목": ..., "판단": ..., "이유": ..., "원문근거": ...}]}
파싱 실패 시 추측으로 채우지 않고 명확한 오류를 돌려준다.
"""
import json
import re

JUDGMENTS = ("일치", "수정필요", "확인필요", "표현차이", "확인불가")

COLUMN_ALIASES = {
    "제품": ["제품", "품목명", "제품명", "약품"],
    "항목": ["항목", "구분", "항목명", "필드"],
    "판단": ["판단", "판정", "결과"],
    "이유": ["이유", "사유", "설명"],
    "원문근거": ["원문근거", "원문 근거", "근거", "원문"],
}


def _find_judgment(cell):
    for j in JUDGMENTS:
        if j in str(cell):
            return j
    return ""


def _norm_row(item):
    def pick(keys):
        for k in keys:
            if k in item and item[k] not in (None, ""):
                return str(item[k]).strip()
        return ""

    row = {
        "제품": pick(COLUMN_ALIASES["제품"]),
        "항목": pick(COLUMN_ALIASES["항목"]),
        "판단": _find_judgment(pick(COLUMN_ALIASES["판단"])) or pick(COLUMN_ALIASES["판단"]),
        "이유": pick(COLUMN_ALIASES["이유"]),
        "원문근거": pick(COLUMN_ALIASES["원문근거"]),
    }
    return row


def parse_result(text):
    """
    반환: {"ok": True, "rows": [...], "format": "json"|"markdown"}
         또는 {"ok": False, "error": "..."}
    """
    t = (text or "").strip()
    if not t:
        return {"ok": False, "error": "붙여넣은 내용이 비어 있습니다."}

    # 1) JSON 시도
    try:
        data = json.loads(t)
        items = None
        if isinstance(data, list):
            items = data
        elif isinstance(data, dict):
            items = data.get("results") or data.get("결과") or data.get("rows")
        if isinstance(items, list) and items:
            rows = [_norm_row(it) for it in items if isinstance(it, dict)]
            if rows:
                return {"ok": True, "rows": rows, "format": "json"}
    except (ValueError, AttributeError):
        pass

    # 2) 마크다운 표 시도
    lines = [ln for ln in t.splitlines() if "|" in ln]
    parsed_rows = []
    header_cols = None
    for ln in lines:
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        # 구분선(|---|---|) 건너뛰기
        if cells and all(re.fullmatch(r":?-{2,}:?", c or "-") for c in cells):
            continue
        joined = " ".join(cells)
        # 헤더 행 감지
        if header_cols is None and any(a in joined for a in ("제품", "품목")) and any(
            a in joined for a in ("판단", "판정")
        ):
            header_cols = {}
            for i, c in enumerate(cells):
                for key, aliases in COLUMN_ALIASES.items():
                    if any(a in c for a in aliases):
                        header_cols[key] = i
                        break
            continue
        if header_cols is not None:
            judged = _find_judgment(joined)
            if not judged and not any(any(a in c for a in COLUMN_ALIASES["제품"]) for c in cells):
                continue
            def get_cell(key):
                idx = header_cols.get(key)
                return cells[idx] if idx is not None and idx < len(cells) else ""
            parsed_rows.append({
                "제품": get_cell("제품"),
                "항목": get_cell("항목"),
                "판단": _find_judgment(get_cell("판단")) or get_cell("판단"),
                "이유": get_cell("이유"),
                "원문근거": get_cell("원문근거"),
            })
    if parsed_rows:
        return {"ok": True, "rows": parsed_rows, "format": "markdown"}

    # 3) 헤더 없이 판정값만 있는 단순 행 (고정 5열 가정)
    simple = []
    for ln in lines:
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if len(cells) >= 3 and _find_judgment(" ".join(cells)):
            simple.append({
                "제품": cells[0] if len(cells) > 0 else "",
                "항목": cells[1] if len(cells) > 1 else "",
                "판단": _find_judgment(" ".join(cells)),
                "이유": cells[3] if len(cells) > 3 else "",
                "원문근거": cells[4] if len(cells) > 4 else "",
            })
    if simple:
        return {"ok": True, "rows": simple, "format": "markdown"}

    return {"ok": False, "error": "형식을 인식하지 못했습니다. 마크다운 표 또는 JSON 형식으로 붙여넣어 주세요."}
