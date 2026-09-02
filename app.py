# -*- coding: utf-8 -*-
"""
의약품 심의자료 진위·오탈자 검증기 v4.0 — Streamlit 웹앱 엔트리포인트 (명세서 6장 STEP 1~6)

원칙(명세서 2장):
- Claude API 를 호출하는 코드는 일절 없다. API Key 입력창도 없다.
- 앱은 "Claude 웹에 붙여넣을 검증 자료"를 생성·복사해 주는 역할만 한다.
- API 키는 Streamlit Secrets 에서만 읽는다 (하드코딩 금지).
- 서버 영구 저장 없음 — 모든 상태는 st.session_state(세션 범위)에서만 유지한다.
"""
import re
import json

import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

from modules import mfds_api, hira_api, drug_matcher
from modules import pptx_parser, xlsx_parser, clipboard_parser
from modules import table_normalizer as normalizer
from modules import rule_validator
from modules import claude_prompt_builder, result_parser

st.set_page_config(page_title="의약품 심의자료 검증기 v4.0", page_icon="💊", layout="wide")

# ---------------- 상태 초기화 ----------------
_SS = st.session_state
_SS.setdefault("search_rows", [])
_SS.setdefault("products", [])
_SS.setdefault("table", None)
_SS.setdefault("pairs", [])
_SS.setdefault("validation_df", None)
_SS.setdefault("orientation", normalizer.ORIENT_ROWS_ARE_ITEMS)


# ---------------- 헬퍼 ----------------
def get_secrets():
    """Streamlit Secrets 에서 인증키 로드. 없으면 빈 문자열(크래시 방지)."""
    try:
        mfds_key = st.secrets.get("MFDS_API_KEY", "") or st.secrets.get("DATA_GO_KOR_API_KEY", "")
        hira_key = st.secrets.get("HIRA_API_KEY", "") or st.secrets.get("DATA_GO_KOR_API_KEY", "")
        return mfds_key, hira_key
    except Exception:
        return "", ""


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
    if field in ("제품명",):
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
    # 그 외 서술형: 전체 사용상주의사항 중 해당 키워드 섹션 또는 전체 텍스트
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
    st.markdown("## 💊 검증기 v4.0")
    st.caption("MFDS 허가사항 · HIRA 약가정보와 심의자료 비교표를 대조합니다.")
    mfds_key, hira_key = get_secrets()
    if not mfds_key and not hira_key:
        st.warning("API 인증키가 설정되지 않았습니다. Streamlit Secrets 에 MFDS_API_KEY / HIRA_API_KEY 를 설정해 주세요.")
    elif not mfds_key:
        st.warning("MFDS_API_KEY 가 없습니다. MFDS 조회가 불가능합니다.")
    elif not hira_key:
        st.warning("HIRA_API_KEY 가 없습니다. 약가조회가 불가능합니다.(MFDS는 정상 동작)")
    st.divider()
    st.caption("사용 흐름: ① 검색·선택 → ② 자동조회 → ③ 비교표 입력 → ④ 구조 확인 → ⑤ 1차 검증 → ⑥ Claude 검증 자료 생성")

# ---------------- 본문 ----------------
st.title("의약품 심의자료 진위·오탈자 검증기 v4.0")
st.caption(
    "기계적으로 확실한 항목(기본정보·숫자단위·약가)은 Python이 자동 판정하고, "
    "의미 비교가 필요한 항목은 Claude 웹 검증용 자료를 자동 생성해 드립니다. "
    "(이 앱은 Claude API를 호출하지 않습니다.)"
)

# ============ STEP 1. 검색 / 선택 ============
st.markdown("## ① STEP 1. 의약품 검색·선택")
with st.form("search_form", clear_on_submit=False):
    keyword = st.text_input("제품명 또는 제조사명 부분 입력", placeholder="예: 아타칸정, 한미약품")
    submitted = st.form_submit_button("🔍 MFDS에서 검색")
if submitted:
    _SS["search_rows"] = []
    if not keyword.strip():
        st.error("검색어를 입력해 주세요.")
    elif not mfds_key:
        st.error("MFDS 인증키가 없습니다. Streamlit Secrets에 MFDS_API_KEY를 설정해 주세요.")
    else:
        try:
            with st.spinner("MFDS 목록에서 검색 중... (최대 5,000건 스캔)"):
                rows = mfds_api.search_products(keyword.strip(), mfds_key)
            if not rows:
                st.warning("검색 결과가 없습니다. 제품명/제조사명 표기를 확인해 주세요. (취하·말소 품목은 제외됩니다)")
            else:
                _SS["search_rows"] = rows
                st.success(f"{len(rows)}건 발견됨")
        except mfds_api.MfdsApiError as e:
            st.error(f"MFDS 검색 실패: {e}")

# ============ STEP 2. 자동 조회 ============
st.markdown("## ② STEP 2. 허가사항·약가 자동 조회")
if _SS["search_rows"]:
    options = [label_of(r) for r in _SS["search_rows"]]
    c1, c2 = st.columns(2)
    applicant = c1.selectbox("신청의약품 (1개)", ["-- 선택 --"] + options, key="applicant_sel")
    comparators = c2.multiselect("비교의약품 (여러 개)", options, key="comparator_sel")

    if st.button("🚀 허가사항·약가 조회 시작", type="primary", key="fetch_btn"):
        sel = []
        if st.session_state.applicant_sel != "-- 선택 --":
            sel.append(("신청의약품", st.session_state.applicant_sel))
        for i, lb in enumerate(st.session_state.comparator_sel, start=1):
            sel.append((f"비교의약품{i}", lb))
        if not sel:
            st.warning("조회할 제품을 선택해 주세요 (신청의약품 1개 + 비교의약품).")
        else:
            bar = st.progress(0.0)
            status_txt = st.empty()
            products = []
            for i, (role, lb) in enumerate(sel):
                seq = lb.split(" | ")[-1].strip()
                status_txt.info(f"({i+1}/{len(sel)}) {lb} → MFDS 상세 + HIRA 약가 조회 중...")
                try:
                    detail = mfds_api.get_product_detail(seq, mfds_key)
                    detail_ok = not detail.get("error")
                except mfds_api.MfdsApiError as e:
                    detail = {"error": str(e)}
                    detail_ok = False
                hira_rows, match = [], {"status": drug_matcher.STATUS_FAIL, "row": None, "method": None}
                if detail_ok:
                    try:
                        hira_rows = hira_api.search_for_product(detail, hira_key)
                        match = drug_matcher.match_hira(detail, hira_rows)
                    except hira_api.HiraApiError as e:
                        match = {"status": f"⚠ HIRA 조회 실패: {e}", "row": None, "method": None}
                price = None
                if match.get("row"):
                    try:
                        price = float(str(match["row"].get("mxCprc") or "").replace(",", ""))
                    except (TypeError, ValueError):
                        price = None
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
        d = p["detail"]
        m = p["match"]
        rows.append({
            "구분": p["role"],
            "제품": p["label"].split(" | ")[0],
            "MFDS": "✅ 조회됨" if not d.get("error") else "❌ 실패",
            "HIRA": f"✅ {len(p['hira_rows'])}건" if p["hira_rows"] else ("❌ 없음/실패" if not d.get("error") else "—"),
            "제품 매칭": m.get("status", drug_matcher.STATUS_FAIL),
            "약가(원)": f"{p['price']:,.0f}" if p["price"] is not None else "—",
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True)
    for p in _SS["products"]:
        with st.expander(f"📄 원문 보기 — {p['role']} / {p['label'].split(' | ')[0]} (MFDS 허가사항·HIRA 약가정보 원문, 요약 아님)"):
            st.markdown("**MFDS 허가사항 원문**")
            st.text(mfds_api.detail_to_raw_text(p["detail"]))
            st.markdown("**HIRA 약가정보 원문**")
            if p["hira_rows"]:
                for r in p["hira_rows"]:
                    st.text(
                        f"품목명: {r.get('itmNm')} / 제조사: {r.get('mnfEntpNm')} / 코드: {r.get('mdsCd')} / "
                        f"상한금액: {r.get('mxCprc')}원 / 급여구분: {r.get('payTpNm')} / 약효분류: {r.get('meftDivNo')}"
                    )
            else:
                st.caption("(HIRA 약가정보 없음 — 품목명 재검색 또는 분업예외 등 사유 확인 필요)")

# ============ STEP 3. 비교표 입력 ============
st.markdown("## ③ STEP 3. 심의자료 비교표 입력")
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

# ============ STEP 4. 구조 확인 ============
st.markdown("## ④ STEP 4. 비교표 구조 확인")
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
    preview = []
    for r in tbl["rows"]:
        preview.append([c.get("text", "") for c in r])
    st.markdown("**원본 표 그리드** (병합은 시작셀에 colspan/rowspan 정보 유지)")
    maxcols = max((len(r) for r in preview), default=1)
    preview_df = pd.DataFrame(preview, columns=[f"C{i+1}" for i in range(maxcols)])
    st.dataframe(preview_df, use_container_width=True)
    st.markdown(f"**공통 스키마 변환 결과: {len(pairs)}개 셀** (명세서 8장 형식)")
    if pairs:
        st.dataframe(pd.DataFrame(pairs), use_container_width=True)

# ============ STEP 5. 1차 자동 검증 ============
st.markdown("## ⑤ STEP 5. Python 1차 자동 검증")
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
        st.caption("🟠 Claude 확인 필요 항목은 Python이 판정하지 않습니다. ⑥ 단계에서 생성되는 자료를 Claude 웹에 붙여넣어 의미 검증하세요.")

        # ── 약가 차이율 카드 (명세서 7장) ──
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

# ============ STEP 6. Claude 검증 자료 생성 ============
st.markdown("## ⑥ STEP 6. Claude 검증 자료 생성 (웹에서 의미 검증)")
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
            try:
                st.dataframe(rdf.style.map(status_color_rule, subset=["판단"]), use_container_width=True)
            except AttributeError:
                st.dataframe(rdf.style.applymap(status_color_rule, subset=["판단"]), use_container_width=True)
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
    "API 오류 시 데이터를 임의로 생성하지 않습니다."
)
