# -*- coding: utf-8 -*-
"""
JSON 캐시 저장/로드 + 로컬 검색 모듈 (v4.1 신규).

- 사이드바에서 전체 목록을 1회 적재 → data/cache/*.json 저장
- 이후 검색은 이 캐시를 메모리(세션)로 올려 로컬 필터링만 수행 (API 호출 없음)
- Streamlit Cloud 는 콜드스타트 시 인스턴스 디스크가 초기화되므로 재적재 필요(세션 內 유지).
"""
import datetime
import json
import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_DIR = os.path.join(PROJECT_ROOT, "data", "cache")
MFDS_CACHE_FILE = os.path.join(CACHE_DIR, "mfds_products.json")
HIRA_CACHE_FILE = os.path.join(CACHE_DIR, "hira_prices.json")


def _ensure_dir():
    os.makedirs(CACHE_DIR, exist_ok=True)


def save_cache(path, items, meta=None):
    """
    items 를 {loaded_at, count, items, ...} 봉투로 JSON 저장.
    반환: 저장된 payload (UI 에 적재시각/건수 표시용).
    """
    _ensure_dir()
    payload = {
        "loaded_at": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
        "count": len(items or []),
        "items": items or [],
    }
    if meta:
        payload.update(meta)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    return payload


def load_cache(path):
    """캐시 JSON 로드. 없거나 손상됐으면 None (추측·생성 금지)."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError, TypeError):
        return None


def filter_products(items, keyword):
    """
    MFDS 캐시 항목 로컬 필터링: 제품명(ITEM_NAME) 또는 제조사명(ENTP_NAME) 부분 일치.
    키워드가 비어 있으면 전체 반환. API 호출 없음.
    """
    kw = str(keyword or "").strip().lower()
    out = []
    for it in items or []:
        name = str(it.get("ITEM_NAME") or "").lower()
        entp = str(it.get("ENTP_NAME") or "").lower()
        if kw and kw not in name and kw not in entp:
            continue
        out.append(it)
    return out
