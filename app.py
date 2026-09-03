# -*- coding: utf-8 -*-
"""
의약품 심의자료 진위·오탈자 검증기 v4.1 — Streamlit 웹앱 엔트리포인트

v4.1 변경점 (사용자 피드백 반영):
1. HTTPS 엔드포인트 하드코딩 + serviceKey 이중 인코딩 방지 + 연결 재시도(3회)
   → 이전 'Connection to apis.data.go.kr:80 timed out' 원인 제거.
2. 검색할 때마다 API를 호출하지 않음:
   사이드바에서 인증키 입력 → 「전체 데이터 불러오기」를 1회 실행하면
   MFDS 품목목록 → data/cache/mfds_products.json, HIRA 약가목록 → data/cache/hira_prices.json 으로 저장.
   이후 의약품 검색·목록·HIRA 매칭은 모두 세션 메모리에 올린 캐시 기준 로컬 필터링(API 호출 없음).
   MFDS 상세 허가사항만 선택 제품에 대해 제품당 1회 조회한다.
3. 캐시가 없으면 「먼저 데이터 불러오기를 실행하세요」 안내 — 가짜 데이터를 만들지 않는다.

고정 원칙(명세서 2장·12장 유지):
- Claude API 호출 코드 없음, API Key 입력창(설정창) 없음 — 앱은 Claude 웹용 검증 자료 생성·복사만.
- CSV 필수 중간 파일 없음, 서버 영구 저장 없음(세션 + 로컬 캐시 파일만), API 키 하드코딩 없음.
"""
import json
import re

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from modules import cache_store, claude_prompt_builder, clipboard_parser, drug_matcher
from modules import hira_api, mfds_api, pptx_parser, result_parser, rule_validator
from modules import table_normalizer as normalizer
from modules import xlsx_parser

st.set_page_config(page_title="의약품 심의자료 검증기 v4.1", page_icon="💊", layout="wide")

_SS = st.session_state

# ---------------- 상태 키 초기화 ----------------
for _k, _v in {
    "search_rows": [], "products": [], "table": None, "pairs": [],
    "validation_df": None, "orientation": normalizer.ORIENT_ROWS_ARE_ITEMS,
}.items():
    _SS.setdefault(_k, _v)

# 시작 시 로컬 캐시(JSON) 자동 로드 — 있으면 API 호출 없이 그대로 사용
if "mfds_items" not in _SS:
    _p = cache_store.load_cache(cache_store.MFDS_CACHE_FILE)
    _SS["mfds_items"] = _p["items"] if _p else None
    _SS["mfds_meta"] = _p
if "hira_items" not in _SS:
    _p = cache_store.load_cache(cache_store.HIRA_CACHE_FILE)
    _SS["hira_items"] = _p["items"] if _p else None
    _SS["hira_meta"] = _p


def secrets_value(*keys, default=""):
    """Streamlit Secrets 에서 우선순위대로 값을 읽는다. 미설정이어도 크래시하지 않는다."""
    try:
        for k in keys:
            v = st.secrets.get(k)
            if v:
                return v
    except Exception:
        pass
    return default


# ---------------- 헬퍼 ----------------
def label_of(row):
    return f"{row['ITEM_NAME']} | {row['ENTP_NAME']} | {row['ITEM_SEQ']}"


def copy_button(text, label="📋 복사", height=52):
    """JS clipboard API 로 텍스트를 클립보드에 복사하는 버튼 (pyperclip 비사용)."""
    payload = json.dumps(text, ensure_ascii=False)
    html = f"""
    <script>
    const t = {payload};
    function doCopy(){{
      if (navigator.clipboard && window.isSecureContext) {{
        navigator.clipboard.writeText(t).then(function(){{ alert('✅ 클립보드에 복사되었습니다.'); }})
        .catch(function(){{ fallbackCopy(); }});
      }} else {{ fallbackCopy(); }}
    }}
    function fallbackCopy(){{
      const ta = document.createElement('textarea');
      ta.value = t; document.body.appendChild(ta); ta.select();
      try {{ document.execCommand('copy'); alert('✅ 클립보드에 복사되었습니다.'); }}
      catch(e) {{ alert('❌ 복사 실패 — 아래 코드 블록에서 직접 복사하세요.'); }}
      document.body.removeChild(ta);
    }}
    </script>
    <button onclick="doCopy()" style="padding:8px 18px;font-size:15px;border-radius:6px;
      border:1px solid #bbb;cursor:pointer;background:#f0f2f6;color:#111;">{label}</button>
    """
    components.html(html, height=height)


def norm_label(text):
    return normalizer._norm_compact(text)


def resolve_product(pair, products):
    """공통 셀의 product 라벨 → products 레코드 연결 (정규화 + 신청/비교 롤 매칭)."""
    pname = str(pair.get("product") or "")
    n = norm_label(pname)
    if not n:
        return None
    for p in products:
        if norm_label(p["label"]) == n:
            return p
    if "신청" in n:
        return next((p for p in products if "신청" in p["label"]), None)
    if "비교" in n:
        comps = [p for p in products if "비교" in p["label"]]
        m = re.search(r"(\d+)", pname)
        if m and comps:
            idx = int(m.group(1)) - 1
            if 0 <= idx < len(comps):
                return comps[idx]
        return comps[0] if comps else None
    return None


def mfds_reference_text(detail, field):
    """항목별 MFDS 원문 텍스트 추출 (요약하지 않음)."""
    if not detail or detail.get("error"):
        return None
    if field == "제품명":
        return detail.get("제품명")
    if field == "성분명":
        return detail.get("성분명")
    if field == "제조판매사":
        return detail.get("제조판매사")
    if field == "제형":
        return detail.get("제형")
    if field == "효능효과":
        return detail.get("효능효과")
    if field == "용법용량":
        return detail.get("용법용량")
    nb_map = {
        "이상반응": "이상반응", "금기사항": "금기사항", "신중투여": "신중투여",
        "상호작용": "상호작용", "소아_고령자투여": "소아_고령자투여",
        "임부_수유부투여": "임부_수유부투여", "과량투여처치": "과량투여처치",
        "보관_취급주의사항": "보관_취급주의사항",
    }
    if field in nb_map:
        nb = detail.get("사용상주의사항") or {}
        return nb.get(nb_map[field])
    return None


def status_color_rule(val):
    s = str(val)
    if "일치" in s:
        return "background-color: #d4edda; color: #155724"
    if "수정필요" in s:
        return "background-color: #f8d7da; color: #721c24"
    if "Claude" in s or "확인필요" in s:
        return "background-color: #fff3cd; color: #856404"
    if "확인불가" in s:
        return "background-color: #e2e3e5; color: #383d41"
    return ""


def style_df(df, status_col="1차 판정"):
    try:
        return df.style.map(status_color_rule, subset=[status_col])
    except AttributeError:
        return df.style.applymap(status_color_rule, subset=[status_col])


def pairs_for_product(product, pairs, products):
    return [p for p in pairs if resolve_product(p, products) is product]


# ---------------- 사이드바 ----------------
with st.sidebar:
    st.markdown("## 💊 검증기 v4.1")
    st.caption("MFDS 허가사항 · HIRA 약가정보 대조 검증")

    mfds_key_in = st.text_input(
        "MFDS 인증키", type="password",
        value=secrets_value("MFDS_API_KEY", "DATA_GO_KOR_API_KEY"))
    hira_key_in = st.text_input(
        "HIRA 인증키", type="password",
        value=secrets_value("HIRA_API_KEY", "DATA_GO_KOR_API_KEY"))
    st.caption("키는 세션에서만 사용하며 저장하지 않습니다.")

    if st.button("📥 전체 데이터 불러오기 (1회)", type="primary"):
        mfds_key = mfds_key_in.strip()
        hira_key = hira_key_in.strip()
        if not mfds_key:
            st.error("MFDS 인증키를 입력해 주세요.")
        else:
            bar = st.progress(0.0)
            status_txt = st.empty()
            try:
                items = mfds_api.fetch_all_products(
                    mfds_key,
                    progress_cb=lambda f, n: (bar.progress(f), status_txt.caption(f"MFDS 품목목록: {n:,}건 수집중...")))
                meta = cache_store.save_cache(cache_store.MFDS_CACHE_FILE, items, {"service": "MFDS 품목허가"})
                _SS["mfds_items"], _SS["mfds_meta"] = items, meta
                st.success(f"MFDS {len(items):,}건 적재 완료 — {meta['loaded_at']}")
            except mfds_api.MfdsApiError as e:
                st.error(f"MFDS 적재 실패: {e}")
            if hira_key:
                bar2 = st.progress(0.0)
                status_txt2 = st.empty()
                try:
                    items2 = hira_api.fetch_all_drug_prices(
                        hira_key,
                        progress_cb=lambda f, n: (bar2.progress(f), status_txt2.caption(f"HIRA 약가: {n:,}건 수집중...")))
                    meta2 = cache_store.save_cache(cache_store.HIRA_CACHE_FILE, items2, {"service": "HIRA 약가"})
                    _SS["hira_items"], _SS["hira_meta"] = items2, meta2
                    st.success(f"HIRA {len(items2):,}건 적재 완료 — {meta2['loaded_at']}")
                except hira_api.HiraApiError as e:
                    st.error(f"HIRA 적재 실패: {e}")
            else:
                st.caption("HIRA 키 미입력 — 약가 목록은 불러오지 않습니다.")
    st.divider()
    st.markdown("**캐시 상태 (data/cache/*.json)**")
    if _SS.get("mfds_items"):
        st.info(f"✅ MFDS {len(_SS['mfds_items']):,}건\n적재시각: {(_SS.get('mfds_meta') or {}).get('loaded_at', '?')}")
    else:
        st.warning("❌ MFDS 데이터 없음 — 「전체 데이터 불러오기」 실행")
    if _SS.get("hira_items"):
        st.info(f"✅ HIRA {len(_SS['hira_items']):,}건\n적재시각: {(_SS.get('hira_meta') or {}).get('loaded_at', '?')}")
    else:
        st.warning("❌ HIRA 데이터 없음 — 키 입력 후 불러오기 실행")
    if st.button("🔄 세션에 캐시 다시 로드"):
        _p = cache_store.load_cache(cache_store.MFDS_CACHE_FILE)
        if _p:
            _SS["mfds_items"], _SS["mfds_meta"] = _p["items"], _p
            st.success(f"MFDS 캐시 {len(_p['items']):,}건 로드")
        else:
            st.warning("MFDS 캐시 파일이 없습니다.")
        _p2 = cache_store.load_cache(cache_store.HIRA_CACHE_FILE)
        if _p2:
            _SS["hira_items"], _SS["hira_meta"] = _p2["items"], _p2
            st.success(f"HIRA 캐시 {len(_p2['items']):,}건 로드")
        else:
            st.warning("HIRA 캐시 파일이 없습니다.")
    st.divider()
    st.caption("사용 흐름: ① 데이터 불러오기(1회) → ② 검색·선택 → ③ 자동조회 → ④ 비교표 입력 → ⑤ 구조 확인 → ⑥ 1차 검증 → ⑦ Claude 자료 생성")

# ---------------- 본문 ----------------
st.title("의약품 심의자료 진위·오탈자 검증기 v4.1")
st.caption(
    "기계적으로 확실한 항목(기본정보·숫자단위·약가)은 Python이 자동 판정하고, 의미 비교는 Claude 웹용 자료를 생성해 드립니다. "
    "(Claude API를 호출하지 않습니다.) 의약품 검색은 불러온 캐시(JSON) 기준 로컬 필터링 — 검색마다 API를 호출하지 않습니다."
)

# ============ ① 검색·선택 (캐시 기반, API 호출 없음) ============
st.markdown("## ① 의약품 검색·선택 (캐시 기반)")
with st.form("search_form"):
    keyword = st.text_input("제품명 또는 제조사명 부분 입력", placeholder="예: 아타칸정, 한미약품")
    submitted = st.form_submit_button("🔍 검색")
if submitted:
    if not _SS.get("mfds_items"):
        st.error("MFDS 데이터가 없습니다. 먼저 사이드바에서 「📥 전체 데이터 불러오기」를 실행해 주세요.")
        _SS["search_rows"] = []
    else:
        with st.spinner("로컬 캐시에서 검색 중..."):
            _SS["search_rows"] = cache_store.filter_products(_SS["mfds_items"], keyword)
        if not _SS["search_rows"]:
            st.warning("검색 결과가 없습니다. (캐시 기준 — 취하·말소 품목 제외)")
        else:
            st.success(f"{len(_SS['search_rows'])}건 — 로컬 캐시 검색(API 호출 없음)")

if _SS["search_rows"]:
    options = [label_of(r) for r in _SS["search_rows"]]
    c1, c2 = st.columns(2)
    c1.selectbox("신청의약품 (1개)", ["-- 선택 --"] + options, key="applicant_sel")
    c2.multiselect("비교의약품 (여러 개)", options, key="comparator_sel")

# ============ ② 자동 조회 (MFDS 상세만 1회, HIRA는 캐시) ============
st.markdown("## ② 허가사항·약가 자동 조회")
if _SS["search_rows"] and st.button("🚀 조회 시작 (MFDS 상세 1회/제품 · HIRA는 캐시에서 매칭)", type="primary", key="fetch_btn"):
    sel = []
    if st.session_state.get("applicant_sel", "-- 선택 --") != "-- 선택 --":
        sel.append(("신청의약품", st.session_state.applicant_sel))
    for i, lb in enumerate(st.session_state.get("comparator_sel", []), start=1):
        sel.append((f"비교의약품{i}", lb))
    mfds_key = mfds_key_in.strip() or secrets_value("MFDS_API_KEY", "DATA_GO_KOR_API_KEY")
    hira_key = hira_key_in.strip() or secrets_value("HIRA_API_KEY", "DATA_GO_KOR_API_KEY")
    if not sel:
        st.warning("조회할 제품을 선택해 주세요 (신청의약품 1개 + 비교의약품).")
    elif not mfds_key:
        st.error("MFDS 인증키가 없습니다. 사이드바에 입력해 주세요.")
    else:
        bar = st.progress(0.0)
        status_txt = st.empty()
        products = []
        for i, (role, lb) in enumerate(sel):
            seq = lb.split(" | ")[-1].strip()
            status_txt.info(f"({i+1}/{len(sel)}) {lb} → MFDS 상세 조회 중...")
            try:
                detail = mfds_api.get_product_detail(seq, mfds_key)
                detail_ok = not detail.get("error")
            except mfds_api.MfdsApiError as e:
                detail = {"error": str(e)}
                detail_ok = False
            hira_rows, match = [], {"status": drug_matcher.STATUS_FAIL, "row": None, "method": None}
            if detail_ok:
                if _SS.get("hira_items"):
                    match = drug_matcher.match_hira(detail, _SS["hira_items"])
                    hira_rows = [match["row"]] if match.get("row") else []
                elif hira_key:
                    try:
                        hira_rows = hira_api.search_for_product(detail, hira_key)
                        match = drug_matcher.match_hira(detail, hira_rows)
                    except hira_api.HiraApiError as e:
                        match = {"status": f"⚠ HIRA 조회 실패: {e}", "row": None, "method": None}
                else:
                    match = {"status": "⚠ HIRA 캐시·키 없음 — 사이드바에서 데이터 불러오기 실행", "row": None, "method": None}
            price = None
            if match.get("row"):
                price = hira_api.get_price(match["row"])
            products.append({
                "label": lb, "role": role, "seq": seq,
                "detail": detail, "hira_rows": hira_rows, "match": match, "price": price,
            })
            bar.progress((i + 1) / len(sel))
        status_txt.empty()
        _SS["products"] = products
        st.success("조회 완료")

if _SS["products"]:
    st.markdown("#### 조회 요약")
    rows = []
    for p in _SS["products"]:
        d, m = p["detail"], p["match"]
        rows.append({
            "구분": p["role"],
            "제품": p["label"].split(" | ")[0],
            "MFDS": "✅" if not d.get("error") else "❌",
            "HIRA": "✅" if p["hira_rows"] else ("⚠" if d.get("error") else "❌"),
            "제품 매칭": m.get("status", drug_matcher.STATUS_FAIL),
            "약가(원)": f"{p['price']:,.0f}" if p["price"] is not None else "—",
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True)
    for p in _SS["products"]:
        with st.expander(f"📄 원문 보기 — {p['role']} / {p['label'].split(' | ')[0]} (원문, 요약 아님)"):
            st.markdown("**MFDS 허가사항 원문**")
            st.text(mfds_api.detail_to_raw_text(p["detail"]))
            st.markdown("**HIRA 약가정보 (매칭 결과)**")
            if p["hira_rows"]:
                for r in p["hira_rows"]:
                    st.text(
                        f"품목명: {r.get('itmNm')} / 제조사: {r.get('mnfEntpNm')} / 코드: {r.get('mdsCd')} / "
                        f"상한금액: {r.get('mxCprc')}원 / 급여구분: {r.get('payTpNm')} / 약효분류: {r.get('meftDivNo')}"
                    )
            else:
                st.caption("(HIRA 매칭 결과 없음 — 캐시 재적재 또는 키 확인 필요)")

# ============ ③ 비교표 입력 ============
st.markdown("## ③ 심의자료 비교표 입력")
tab1, tab2, tab3 = st.tabs(["📊 PPTX 업로드", "📈 XLSX 업로드", "📋 복사·붙여넣기"])

with tab1:
    f_pptx = st.file_uploader("PPT 파일 업로드 (.pptx)", type=["pptx"], key="pptx_up")
    if f_pptx:
        try:
            tables = pptx_parser.parse_pptx(f_pptx.getvalue())
            _SS["table"] = tables[0]
            st.success(f"인식 완료: {len(tables)}개 표 중 1번째 표를 사용합니다 (슬라이드 {tables[0]['slide']}, {len(tables[0]['rows'])}행).")
        except pptx_parser.PptxParseError as e:
            st.error(str(e))

with tab2:
    f_xlsx = st.file_uploader("Excel 파일 업로드 (.xlsx)", type=["xlsx"], key="xlsx_up")
    if f_xlsx:
        try:
            tables = xlsx_parser.parse_xlsx(f_xlsx.getvalue())
            _SS["table"] = tables[0]
            st.success(f"인식 완료: 시트 '{tables[0]['slide']}' ({len(tables[0]['rows'])}행 × {len(tables[0]['rows'][0]) if tables[0]['rows'] else 0}열, 병합 셀 보존).")
        except xlsx_parser.XlsxParseError as e:
            st.error(str(e))

with tab3:
    pasted = st.text_area("PPT/Excel에서 복사한 표를 붙여넣기 (탭 구분)", height=200, key="paste_area")
    if st.button("붙여넣기 표 인식", key="paste_btn"):
        try:
            tables = clipboard_parser.parse_clipboard(pasted)
            _SS["table"] = tables[0]
            st.success(f"인식 완료: {len(tables[0]['rows'])}행 × {len(tables[0]['rows'][0])}열")
        except clipboard_parser.ClipboardParseError as e:
            st.error(str(e))

# ============ ④ 구조 확인 ============
st.markdown("## ④ 비교표 구조 확인")
if _SS["table"] is not None:
    tbl = _SS["table"]
    st.caption(f"출처: {tbl['source']} / 슬라이드·시트: {tbl['slide']} / 표 순번: {tbl['table_index']}")
    auto = normalizer.guess_orientation(tbl["rows"])
    auto_label = "행 = 항목, 열 = 제품" if auto == normalizer.ORIENT_ROWS_ARE_ITEMS else "열 = 항목, 행 = 제품"
    st.write(f"자동 추정: **{auto_label}**")
    orient = st.radio(
        "표 방향 (자동 추정이 틀리면 직접 선택)",
        [normalizer.ORIENT_ROWS_ARE_ITEMS, normalizer.ORIENT_COLS_ARE_ITEMS],
        format_func=lambda x: "행 = 항목(필드), 열 = 제품" if x == normalizer.ORIENT_ROWS_ARE_ITEMS else "열 = 항목(필드), 행 = 제품",
        horizontal=True,
        index=0 if auto == normalizer.ORIENT_ROWS_ARE_ITEMS else 1,
    )
    _SS["orientation"] = orient
    pairs = normalizer.to_field_product_pairs(tbl, orient)
    _SS["pairs"] = pairs
    preview = [[c.get("text", "") for c in r] for r in tbl["rows"]]
    st.markdown("**원본 표 그리드** (병합은 시작셀에 colspan/rowspan 정보 유지)")
    maxcols = max((len(r) for r in preview), default=1)
    st.dataframe(pd.DataFrame(preview, columns=[f"C{i+1}" for i in range(maxcols)]), use_container_width=True)
    st.markdown(f"**공통 스키마 변환 결과: {len(pairs)}개 셀** (명세서 8장 형식)")
    if pairs:
        st.dataframe(pd.DataFrame(pairs), use_container_width=True)

# ============ ⑤ 1차 자동 검증 ============
st.markdown("## ⑤ Python 1차 자동 검증")
if _SS.get("products") and _SS.get("pairs"):
    if st.button("🔍 1차 규칙 검증 실행 (기본정보·숫자단위·약가만)", key="validate_btn"):
        products = _SS["products"]
        validation = []
        for pair in _SS["pairs"]:
            product = resolve_product(pair, products)
            ref = mfds_reference_text(product["detail"] if product else None, pair.get("field", ""))
            hira_price = product["price"] if product else None
            status, reason = rule_validator.evaluate_pair(pair, ref, hira_price)
            validation.append({
                "제품(비교표)": pair.get("product", ""),
                "연결 제품": product["label"] if product else "❌ 원문 매칭 안 됨",
                "항목": pair.get("field", ""),
                "1차 판정": status,
                "사유": reason,
            })
        _SS["validation_df"] = pd.DataFrame(validation)
    if _SS.get("validation_df") is not None:
        vdf = _SS["validation_df"]
        st.dataframe(style_df(vdf, "1차 판정"), use_container_width=True)
        st.caption("🟠 Claude 확인 필요 항목은 Python이 판정하지 않습니다. ⑥ 단계 자료를 Claude 웹에 붙여넣어 의미 검증하세요.")

        products = _SS["products"]
        applicant = next((p for p in products if "신청" in p["label"]), None)
        comps = [p for p in products if p is not applicant]
        if applicant and comps:
            app_price = applicant.get("price")
            comp_prices = [c.get("price") for c in comps if c.get("price") is not None]
            if app_price is not None and comp_prices:
                base = min(comp_prices)
                diff = rule_validator.price_diff_percent(app_price, base)
                if diff is not None:
                    color = "red" if diff > 0 else "blue"
                    arrow = "▲" if diff > 0 else ("▼" if diff < 0 else "▶")
                    sign = "+" if diff > 0 else ""
                    st.markdown(
                        f"#### 💰 약가 비교 (신청의약품 vs 비교의약품 최저가)\n"
                        f"신청의약품 {app_price:,.0f}원 — 비교 최저가 {base:,.0f}원 → "
                        f"<span style='color:{color};font-weight:bold;'>{arrow} 차이율 {sign}{diff:.1f}%</span> "
                        f"(양수=상승 빨강 / 음수=하락 파랑, 명세서 7장 공식)",
                        unsafe_allow_html=True,
                    )
else:
    st.caption("② 단계에서 제품을 조회하고 ③~④ 단계에서 비교표를 인식해야 검증할 수 있습니다.")

# ============ ⑥ Claude 검증 자료 생성 ============
st.markdown("## ⑥ Claude 검증 자료 생성 (웹에서 의미 검증)")
if _SS.get("products") and _SS.get("pairs"):
    products = _SS["products"]
    prompt_products = []
    for p in products:
        prompt_products.append({
            "label": p["label"],
            "detail": p["detail"],
            "hira_rows": p["hira_rows"],
            "pairs": pairs_for_product(p, _SS["pairs"], products),
        })
    full_prompt = claude_prompt_builder.build_full_prompt(prompt_products)

    st.markdown("### 전체 자료 (전 제품 × MFDS 원문 + HIRA 약가정보 + 비교표)")
    copy_button(full_prompt, "📋 [Claude용 전체 자료 복사]")
    with st.expander("전체 자료 미리보기"):
        st.code(full_prompt)

    st.markdown("### 제품별 자료")
    for p in prompt_products:
        with st.expander(f"제품별 — {p['label']}"):
            prod_prompt = claude_prompt_builder.build_product_prompt(p)
            copy_button(prod_prompt, f"📋 [{p['label']} 복사]")
            st.code(prod_prompt)

    st.markdown("### 항목별 자료")
    field_sel = st.selectbox(
        "항목 선택",
        ["효능효과", "용법용량", "이상반응", "금기사항", "상호작용", "기본정보"],
        key="field_sel",
    )
    field_prompt = claude_prompt_builder.build_field_prompt(field_sel, prompt_products)
    copy_button(field_prompt, f"📋 [{field_sel}만 복사]")
    with st.expander(f"{field_sel} 항목 자료 미리보기"):
        st.code(field_prompt)

    st.divider()
    st.markdown("### Claude 검증 결과 재붙여넣기")
    st.caption("Claude 웹에서 받은 마크다운 표(`| 제품 | 항목 | 판단 | 이유 | 원문 근거 |`) 또는 JSON(`{\"results\":[...]}`)을 붙여넣으면 결과 표로 렌더링합니다.")
    claude_out = st.text_area("Claude 결과 붙여넣기", height=200, key="claude_result")
    if st.button("결과 표로 렌더링", key="parse_result_btn"):
        parsed = result_parser.parse_result(claude_out)
        if parsed["ok"]:
            rdf = pd.DataFrame(parsed["rows"])
            styler = None
            try:
                styler = rdf.style.map(status_color_rule, subset=["판단"])
            except AttributeError:
                styler = rdf.style.applymap(status_color_rule, subset=["판단"])
            st.dataframe(styler, use_container_width=True)
        else:
            st.error(parsed["error"])
            st.markdown("**입력 원문 (그대로)**")
            st.code(claude_out)
else:
    st.caption("②~④ 단계를 먼저 완료하면 이곳에서 Claude용 검증 자료를 생성할 수 있습니다.")

st.divider()
st.caption(
    "⚠️ 본 앱은 기계적으로 판별 가능한 항목(기본정보·숫자단위·약가)만 자동 판정합니다. "
    "의미 비교는 Claude 웹에서 수행하며, 본 앱은 그 자료를 생성·복사하는 역할만 합니다. "
    "API 오류·캐시 부재 시 데이터를 임의로 생성하지 않습니다."
)
