# -*- coding: utf-8 -*-
"""
복사·붙여넣기 표 추출 모듈 — PPT/Excel 에서 복사한 표를 텍스트 영역에 붙여넣으면
탭 구분으로 행/열 구조를 파싱한다(명세서 6장 STEP 3).

반환 형식:
{
  "source": "paste",
  "slide": 0,                 # 복사·붙여넣기 입력은 슬라이드 없음 → 0
  "table_index": 0,
  "rows": [ [ {"text","rowspan","colspan"}, ... ], ... ]
}
"""


class ClipboardParseError(Exception):
    """붙여넣기 표 파싱 오류 (사용자에게 그대로 노출)"""


def parse_clipboard(text):
    """
    탭 구분 텍스트를 표로 파싱한다.
    빈 줄은 무시하고, 탭이 하나도 없으면 단일 열 표로 취급한다.
    """
    if text is None or not str(text).strip():
        raise ClipboardParseError("붙여넣을 내용이 비어 있습니다. PPT/Excel에서 복사한 표를 붙여넣어 주세요.")
    lines = [ln.rstrip("\r") for ln in str(text).split("\n")]
    lines = [ln for ln in lines if ln.strip()]
    if not lines:
        raise ClipboardParseError("붙여넣을 내용이 비어 있습니다.")

    rows = []
    has_tab = any("\t" in ln for ln in lines)
    for ln in lines:
        if has_tab:
            cells = [c.strip() for c in ln.split("\t")]
        else:
            cells = [ln.strip()]
        rows.append([{"text": c, "rowspan": 1, "colspan": 1} for c in cells])
    return [{
        "source": "paste",
        "slide": 0,
        "table_index": 0,
        "rows": rows,
    }]
