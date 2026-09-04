# -*- coding: utf-8 -*-
"""
의약품 그룹(신청의약품/비교의약품N) 관리 및 함량 파싱 유틸리티.

- 그룹 구조: group_id -> {label, seqs:[ITEM_SEQ,...]}
- 각 품목의 함량은 제품명 문자열에서 기계적으로 파싱한다.
- 파싱 실패 시 빈 문자열을 반환하며, 호출부(UI)에서 사용자가 편집 가능해야 한다.
"""
import re

DEFAULT_GROUPS = {
    "applicant": {"label": "신청의약품", "seqs": []},
    "comp1": {"label": "비교의약품1", "seqs": []},
}

UNIT_MAP = {
    "마이크로그램": "mcg",
    "μg": "mcg",
    "ug": "mcg",
    "mcg": "mcg",
    "밀리그램": "mg",
    "mg": "mg",
    "그램": "g",
    "g": "g",
    "밀리리터": "mL",
    "ml": "mL",
    "mL": "mL",
    "iu": "IU",
    "IU": "IU",
    "%": "%",
    "meq": "mEq",
    "mEq": "mEq",
}

STRENGTH_PATTERN = re.compile(
    r"(\d+(?:\.\d+)?)\s*(마이크로그램|μg|ug|mcg|밀리그램|mg|그램|g|밀리리터|mL|ml|IU|iu|%|mEq|meq)"
)


def migrate_legacy(selection=None, applicant_seq=None, comparator_seqs=None):
    """구 버전 단일 신청약/다중 비교약 상태를 그룹 구조로 변환한다."""
    groups = {gid: {"label": meta["label"], "seqs": list(meta["seqs"])} for gid, meta in DEFAULT_GROUPS.items()}
    selection = [str(x) for x in (selection or []) if str(x)]
    applicant_seq = str(applicant_seq) if applicant_seq else None
    comparator_seqs = [str(x) for x in (comparator_seqs or []) if str(x)]
    if applicant_seq:
        groups["applicant"]["seqs"] = [applicant_seq]
    if comparator_seqs:
        groups["comp1"]["seqs"] = comparator_seqs[:]
    known = set(flatten_group_seqs(groups))
    leftover = [seq for seq in selection if seq not in known]
    if leftover:
        groups["comp1"]["seqs"].extend(leftover)
    return groups


def comparator_group_ids(groups):
    ids = [gid for gid in groups if gid.startswith("comp")]
    ids.sort(key=lambda x: int(re.sub(r"\D", "", x) or 0))
    return ids


def next_comparator_group_id(groups):
    nums = [int(re.sub(r"\D", "", gid) or 0) for gid in comparator_group_ids(groups)]
    return f"comp{(max(nums) if nums else 0) + 1}"


def ensure_minimum_groups(groups=None):
    """기본 신청/비교1 그룹을 보장한다."""
    base = {} if not isinstance(groups, dict) else {gid: {"label": meta.get("label", gid), "seqs": list(meta.get("seqs", []))} for gid, meta in groups.items()}
    for gid, meta in DEFAULT_GROUPS.items():
        if gid not in base:
            base[gid] = {"label": meta["label"], "seqs": []}
    return base


def ordered_group_ids(groups):
    groups = ensure_minimum_groups(groups)
    return ["applicant"] + comparator_group_ids(groups)


def flatten_group_seqs(groups):
    out = []
    for gid in ordered_group_ids(groups):
        for seq in groups.get(gid, {}).get("seqs", []):
            s = str(seq)
            if s and s not in out:
                out.append(s)
    return out


def assign_seqs_to_group(groups, group_id, seqs):
    """선택 품목들을 지정 그룹으로 이동한다(중복 제거)."""
    groups = ensure_minimum_groups(groups)
    seqs = [str(x) for x in seqs if str(x)]
    for gid in groups:
        groups[gid]["seqs"] = [str(s) for s in groups[gid].get("seqs", []) if str(s) not in seqs]
    groups.setdefault(group_id, {"label": group_label(group_id), "seqs": []})
    for seq in seqs:
        if seq not in groups[group_id]["seqs"]:
            groups[group_id]["seqs"].append(seq)
    return groups


def remove_seqs(groups, seqs):
    groups = ensure_minimum_groups(groups)
    remove = {str(x) for x in seqs if str(x)}
    for gid in groups:
        groups[gid]["seqs"] = [str(s) for s in groups[gid].get("seqs", []) if str(s) not in remove]
    return groups


def group_label(group_id):
    if group_id == "applicant":
        return "신청의약품"
    if str(group_id).startswith("comp"):
        n = re.sub(r"\D", "", str(group_id)) or "1"
        return f"비교의약품{n}"
    return str(group_id)


def normalize_strength_unit(unit):
    return UNIT_MAP.get(str(unit).strip(), str(unit).strip())


def parse_strength(text):
    """
    제품명 문자열에서 첫 번째 '숫자+단위' 함량 토큰을 파싱한다.
    예: '나르코설하정100마이크로그램' -> '100mcg'
    실패 시 '' 반환.
    """
    t = str(text or "")
    m = STRENGTH_PATTERN.search(t)
    if not m:
        return ""
    num = m.group(1)
    try:
        f = float(num)
        num = str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        pass
    return f"{num}{normalize_strength_unit(m.group(2))}"


def strength_sort_key(text):
    """함량 문자열을 정렬 가능한 수치 키로 변환. 실패 시 매우 큰 값."""
    s = parse_strength(text) if not re.search(r"\d", str(text or "")) else str(text or "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*(mcg|mg|g|mL|IU|%|mEq)", s)
    if not m:
        return (999999999.0, s)
    value = float(m.group(1))
    unit = m.group(2)
    # 가능한 경우 mcg 기준 환산
    if unit == "g":
        value *= 1_000_000
    elif unit == "mg":
        value *= 1_000
    elif unit == "mcg":
        value *= 1
    return (value, unit)


def build_group_items(groups, by_seq, strength_overrides=None):
    """그룹 구조와 MFDS 목록 row dict(by_seq)로부터 그룹별 품목 메타를 생성한다."""
    strength_overrides = strength_overrides or {}
    groups = ensure_minimum_groups(groups)
    out = []
    for gid in ordered_group_ids(groups):
        meta = groups.get(gid, {})
        label = meta.get("label") or group_label(gid)
        items = []
        for seq in meta.get("seqs", []):
            row = by_seq.get(str(seq))
            if not row:
                continue
            item_name = str(row.get("ITEM_NAME") or "")
            parsed = parse_strength(item_name)
            strength = str(strength_overrides.get(str(seq)) or parsed or "")
            items.append({
                "group_id": gid,
                "group_label": label,
                "seq": str(seq),
                "item_name": item_name,
                "entp_name": str(row.get("ENTP_NAME") or ""),
                "strength": strength,
                "parsed_strength": parsed,
                "sort_key": strength_sort_key(strength or parsed or item_name),
            })
        items.sort(key=lambda x: x["sort_key"])
        out.extend(items)
    return out
