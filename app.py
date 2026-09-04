# -*- coding: utf-8 -*-
"""
의약품 심의자료 진위·오탈자 검증기 v4.3 — Streamlit 웹앱 엔트리포인트

v4.3 변경점:
1. 신청의약품/비교의약품을 단일 품목이 아닌 "그룹"으로 관리
   - 신청의약품: N개 품목 허용
   - 비교의약품1~N: 각 그룹마다 여러 함량 품목 허용
2. 의약품 검색·선택 UI/UX 개선
   - 검색 결과 다중 선택 → 원하는 그룹에 추가
   - 비교약 그룹 동적 추가
   - 그룹별 표에서 함량 자동 파싱값 확인/수정 및 행 제거
3. 검색은 계속 캐시(JSON) 기준 로컬 필터링
   - 검색마다 MFDS/HIRA 목록 API를 호출하지 않음
   - MFDS 상세만 선택 제품에 대해 제품당 1회 조회
   - HIRA는 캐시 우선 매칭, 캐시가 없으면 제품당 1회 폴백 조회

고정 원칙:
- Claude API 호출 코드 없음
- CSV 필수 중간 파일 없음
- API 키 하드코딩 없음
- 서버 영구 저장 없음(세션 + 로컬 캐시 파일만)
"""
import html as html_lib
import json
import re

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from modules import cache_store, claude_prompt_builder, clipboard_parser, drug_matcher, grouping
from modules import hira_api, mfds_api, pptx_parser, result_parser, rule_validator
from modules import table_normalizer as normalizer
from modules import xlsx_parser

st.set_page_config(page_title="의약품 심의자료 검증기 v4.3", page_icon="💊", layout="wide")

_SS = st.session_state

# ---------------- 상태 키 초기화 ----------------
for _k, _v in {
    "search_rows": [],
    "products": [],
    "table": None,
    "pairs": [],
    "validation_df": None,
    "orientation": normalizer.ORIENT_ROWS_ARE_ITEMS,
    "selection": [],
    "fetched_df": None,
    "groups": None,
    "strength_overrides": {},
}.items():
    _SS.setdefault(_k, _v)


# ---------------- 캐시 자동 로드 ----------------
if "mfds_items" not in _SS:
    _p = cache_store.load_cache(cache_store.MFDS_CACHE_FILE)
    _SS["mfds_items"] = _p["items"] if _p else None
    _SS["mfds_meta"] = _p
if "hira_items" not in _SS:
    _p = cache_store.load_cache(cache_store.HIRA_CACHE_FILE)
    _SS["hira_items"] = _p["items"] if _p else None
    _SS["hira_meta"] = _p


# ---------------- 조회 UI 설정 상수 ----------------
EXTRA_FIELD_SPECS = {
    "영문제품명": {"label": "영문 제품명"},
    "주성분영문명": {"label": "주성분 영문명"},
    "전문일반구분": {"label": "전문·일반의약품 구분"},
    "ATC코드": {"label": "ATC 코드"},
    "원료약품및분량(함량)": {"label": "원료약품 및 분량(함량)"},
    "소아_고령자투여": {"label": "소아·고령자 투여"},
    "임부_수유부투여": {"label": "임부·수유부 투여"},
    "금기사항": {"label": "금기사항"},
    "신중투여": {"label": "신중투여"},
    "상호작용": {"label": "상호작용"},
    "이상반응": {"label": "이상반응"},
    "과량투여처치": {"label": "과량투여 시 처치"},
    "적용상의주의사항": {"label": "적용상의 주의사항"},
    "포장단위": {"label": "포장단위"},
    "유효기간": {"label": "유효기간"},
    "보관정보": {"label": "보관정보"},
    "보관_취급주의사항": {"label": "보관 및 취급상의 주의사항"},
    "성상": {"label": "성상"},
}
EXTRA_FIELD_ORDER = list(EXTRA_FIELD_SPECS.keys())
EXTRA_FIELD_LABELS = {key: EXTRA_FIELD_SPECS[key]["label"] for key in EXTRA_FIELD_ORDER}
NEW_DRUG_PRESET_KEYS = [
    "영문제품명", "ATC코드", "포장단위", "주성분영문명", "원료약품및분량(함량)",
    "유효기간", "성상", "전문일반구분", "보관정보", "보관_취급주의사항",
]
RESULT_COLUMN_ORDER = [
    "구분", "함량", "제형", "함량/제형", "허가제품명", "제약사한글명", "약가", "약효분류",
    "성분명", "효능효과", "용법용량", "영문제품명", "주성분영문명", "전문일반구분",
    "ATC코드", "원료약품및분량(함량)", "포장단위", "유효기간", "성상", "보관정보",
    "소아_고령자투여", "임부_수유부투여", "금기사항", "신중투여", "상호작용", "이상반응",
    "과량투여처치", "적용상의주의사항", "보관_취급주의사항",
]


# ---------------- 상태/그룹 헬퍼 ----------------
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


def init_group_state():
    """구버전 상태를 그룹 구조로 마이그레이션하고 기본 그룹을 보장한다."""
    if not isinstance(_SS.get("groups"), dict):
        _SS["groups"] = grouping.migrate_legacy(
            selection=_SS.get("selection", []),
            applicant_seq=_SS.get("applicant_seq"),
            comparator_seqs=_SS.get("comparator_seqs", []),
        )
    _SS["groups"] = grouping.ensure_minimum_groups(_SS.get("groups"))
    _SS["selection"] = grouping.flatten_group_seqs(_SS["groups"])
    if not isinstance(_SS.get("strength_overrides"), dict):
        _SS["strength_overrides"] = {}


def label_of(row):
    return f"{row['ITEM_NAME']} | {row['ENTP_NAME']} | {row['ITEM_SEQ']}"


def format_price(value):
    if value is None or value == "":
        return ""
    try:
        return f"{float(value):,.0f}원"
    except (TypeError, ValueError):
        return str(value)


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


def strength_text(text):
    return grouping.parse_strength(text)


def group_items(by_seq):
    return grouping.build_group_items(_SS.get("groups"), by_seq, _SS.get("strength_overrides"))


def sync_selection_from_groups():
    _SS["selection"] = grouping.flatten_group_seqs(_SS.get("groups"))


def group_target_options():
    groups = grouping.ensure_minimum_groups(_SS.get("groups"))
    return [(gid, groups[gid].get("label") or grouping.group_label(gid)) for gid in grouping.ordered_group_ids(groups)]


def source_value_from_detail(detail, field_key):
    if not detail or detail.get("error"):
        return ""
    if field_key in (
        "제품명", "성분명", "효능효과", "용법용량", "영문제품명", "주성분영문명",
        "전문일반구분", "ATC코드", "원료약품및분량(함량)", "포장단위", "유효기간",
        "보관정보", "성상", "제형",
    ):
        return detail.get(field_key, "")
    nb = detail.get("사용상주의사항") or {}
    if field_key in nb:
        return nb.get(field_key, "")
    return detail.get(field_key, "")


def order_result_columns(df):
    ordered = [column for column in RESULT_COLUMN_ORDER if column in df.columns]
    ordered += [column for column in df.columns if column not in ordered]
    return df[ordered]


def build_source_result_df(products, selected_extras):
    rows = []
    for p in products:
        detail = p.get("detail") or {}
        matched = p.get("match", {}).get("row") or {}
        strength = p.get("strength", "")
        form = detail.get("제형", "")
        strength_form = " / ".join([x for x in [strength, form] if x])
        row = {
            "구분": p.get("group_label") or p.get("role", ""),
            "함량": strength,
            "제형": form,
            "함량/제형": strength_form,
            "허가제품명": detail.get("제품명") or p.get("item_name") or p.get("label", "").split(" | ")[0],
            "제약사한글명": detail.get("제조판매사") or p.get("entp_name") or (p.get("label", "").split(" | ")[1] if " | " in p.get("label", "") else ""),
            "약가": format_price(p.get("price")),
            "약효분류": matched.get("meftDivNo", "") if matched else "",
            "성분명": detail.get("성분명", ""),
            "효능효과": detail.get("효능효과", ""),
            "용법용량": detail.get("용법용량", ""),
        }
        for key in selected_extras:
            row[key] = source_value_from_detail(detail, key)
        rows.append(row)
    if not rows:
        return None
    return order_result_columns(pd.DataFrame(rows))


def make_display_df(result_df, transpose_view=False):
    display_df = order_result_columns(result_df.copy())
    if transpose_view:
        display_df = display_df.T
        display_df.index.name = "항목"
    return display_df


def make_comparison_df(result_df):
    """행에는 조회 항목, 열에는 그룹-함량별 의약품을 배치한 비교표를 생성합니다."""
    comparison = result_df.copy()
    comparison_names = []
    seen = {}
    for _, row in comparison.iterrows():
        base = f"{row.get('구분', '')} {row.get('함량', '')}".strip()
        if not base:
            base = str(row.get("허가제품명", "품목"))
        else:
            base = f"{base} | {row.get('허가제품명', '품목')}"
        seen[base] = seen.get(base, 0) + 1
        comparison_names.append(base if seen[base] == 1 else f"{base} ({seen[base]})")
    comparison["_비교용약품명"] = comparison_names
    comparison = comparison.set_index("_비교용약품명").T
    comparison.index.name = "조회 항목"
    return comparison


def render_resizable_wrapped_table(display_df, show_index=False, height=720, table_key="result"):
    """외부 라이브러리 없이 컬럼 드래그 리사이즈와 셀 줄바꿈을 함께 제공합니다."""
    table_df = display_df.reset_index() if show_index else display_df.reset_index(drop=True)
    if show_index:
        table_df = table_df.rename(columns={table_df.columns[0]: str(display_df.index.name or "항목")})
    headers = [str(column) for column in table_df.columns]
    long_columns = {"효능효과", "용법용량", "성분명", "이상반응", "상호작용", "금기사항", "원료약품및분량(함량)"}
    narrow_columns = {"약가", "제약사한글명", "약효분류", "구분", "함량", "제형"}

    def cell(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            value = ""
        return html_lib.escape(str(value)).replace("\n", "<br>")

    if show_index:
        index_width_px = 180
        drug_column_count = max(len(headers) - 1, 1)
        col_widths = [f"{index_width_px}px"] + [
            f"calc((100% - {index_width_px}px) / {drug_column_count})"
        ] * drug_column_count
    else:
        def width_for(header):
            if header in narrow_columns:
                return 180
            if header == "허가제품명":
                return 420
            if header in long_columns:
                return 720
            return 320
        col_widths = [f"{width_for(header)}px" for header in headers]

    colgroup = "".join(
        f'<col data-column="{index}" style="width:{width}">' for index, width in enumerate(col_widths)
    )
    header_html = "".join(
        f'<th data-column="{index}">{cell(header)}<span class="resize-handle" data-column="{index}"></span></th>'
        for index, header in enumerate(headers)
    )
    body_html = []
    for row in table_df.itertuples(index=False, name=None):
        body_html.append("<tr>" + "".join(f"<td>{cell(value)}</td>" for value in row) + "</tr>")
    table_width_css = "width:100%;" if show_index else "width:max-content; min-width:100%;"
    markup = f"""
    <style>
      html, body {{ margin:0; padding:0; background:#fff; font-family:Arial, sans-serif; }}
      .table-wrap {{ width:100%; height:calc(100vh - 24px); overflow:auto; border:1px solid #d9dee7; }}
      table {{ border-collapse:collapse; table-layout:fixed; {table_width_css} font-size:13px; }}
      th, td {{ border:1px solid #d9dee7; padding:8px; vertical-align:top; white-space:normal; overflow-wrap:anywhere; word-break:break-word; line-height:1.45; }}
      th {{ position:sticky; top:0; z-index:2; background:#f3f6fa; font-weight:700; text-align:left; user-select:none; }}
      .resize-handle {{ position:absolute; top:0; right:-4px; width:8px; height:100%; cursor:col-resize; z-index:3; }}
      .resize-handle:hover, .resizing {{ background:#5b8def; opacity:.55; }}
      body.resizing {{ cursor:col-resize; user-select:none; }}
    </style>
    <div class="table-wrap" id="wrap-{table_key}">
      <table id="table-{table_key}"><colgroup>{colgroup}</colgroup><thead><tr>{header_html}</tr></thead><tbody>{''.join(body_html)}</tbody></table>
    </div>
    <script>
      (() => {{
        const table = document.getElementById('table-{table_key}');
        const cols = table.querySelectorAll('col');
        table.querySelectorAll('.resize-handle').forEach(handle => {{
          handle.addEventListener('mousedown', event => {{
            event.preventDefault();
            const index = Number(handle.dataset.column);
            const startX = event.clientX;
            const startWidth = cols[index].getBoundingClientRect().width;
            document.body.classList.add('resizing');
            handle.classList.add('resizing');
            const move = moveEvent => {{
              const nextWidth = Math.max(120, startWidth + moveEvent.clientX - startX);
              cols[index].style.width = nextWidth + 'px';
            }};
            const stop = () => {{
              document.body.classList.remove('resizing');
              handle.classList.remove('resizing');
              document.removeEventListener('mousemove', move);
              document.removeEventListener('mouseup', stop);
            }};
            document.addEventListener('mousemove', move);
            document.addEventListener('mouseup', stop);
          }});
        }});
      }})();
    </script>
    """
    components.html(markup, height=height, scrolling=False)


def selected_extra_keys():
    keys = []
    for key in EXTRA_FIELD_ORDER:
        if st.session_state.get(f"extra_{key}", False):
            keys.append(key)
    return keys


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
    if field == "함량":
        return detail.get("원료약품및분량(함량)")
    if field == "효능효과":
        return detail.get("효능효과")
    if field == "용법용량":
        return detail.get("용법용량")
    nb_map = {
        "이상반응": "이상반응",
        "금기사항": "금기사항",
        "신중투여": "신중투여",
        "상호작용": "상호작용",
        "소아_고령자투여": "소아_고령자투여",
        "임부_수유부투여": "임부_수유부투여",
        "과량투여처치": "과량투여처치",
        "보관_취급주의사항": "보관_취급주의사항",
        "적용상의주의사항": "적용상의주의사항",
    }
    if field in nb_map:
        nb = detail.get("사용상주의사항") or {}
        return nb.get(nb_map[field])
    return detail.get(field)


def _match_strength_hint(text):
    t = strength_text(text)
    return t.lower() if t else ""


def resolve_product(pair, products):
    """공통 셀의 product 라벨을 실제 fetched product 한 건에 연결한다."""
    pname = str(pair.get("product") or "")
    n = norm_label(pname)
    if not n:
        return None

    exact = []
    for p in products:
        candidates = {
            norm_label(p.get("label", "")),
            norm_label(p.get("item_name", "")),
            norm_label((p.get("detail") or {}).get("제품명", "")),
            norm_label(p.get("group_label", "")),
            norm_label(p.get("role", "")),
        }
        if n in candidates:
            exact.append(p)
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        hint = _match_strength_hint(pair.get("value", ""))
        if hint:
            hinted = [p for p in exact if str(p.get("strength", "")).lower() == hint]
            if len(hinted) == 1:
                return hinted[0]
        return None

    role_matches = [p for p in products if n in {norm_label(p.get("group_label", "")), norm_label(p.get("role", ""))}]
    if len(role_matches) == 1:
        return role_matches[0]
    if len(role_matches) > 1:
        hint = _match_strength_hint(pair.get("value", ""))
        if hint:
            hinted = [p for p in role_matches if str(p.get("strength", "")).lower() == hint]
            if len(hinted) == 1:
                return hinted[0]
        return None
    return None


def pairs_for_product(product, pairs, products):
    out = []
    for p in pairs:
        resolved = resolve_product(p, products)
        if resolved is product:
            out.append(p)
            continue
        pname = norm_label(p.get("product", ""))
        if pname and pname in {norm_label(product.get("group_label", "")), norm_label(product.get("role", ""))}:
            out.append(p)
    return out


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


def build_group_price_summary(products):
    """신청의약품 각 함량을 동일 함량 비교군 최저가와 비교한다."""
    applicants = [p for p in products if p.get("group_id") == "applicant"]
    comparators = [p for p in products if str(p.get("group_id", "")).startswith("comp")]
    rows = []
    for app in applicants:
        a_strength = str(app.get("strength") or "")
        same_strength = [
            c for c in comparators
            if str(c.get("strength") or "").lower() == a_strength.lower() and c.get("price") is not None
        ]
        if app.get("price") is None:
            rows.append({
                "함량": a_strength or "(미지정)",
                "신청의약품": app.get("item_name", ""),
                "신청약가": "—",
                "비교최저가": "—",
                "기준 비교군": "—",
                "차이율": "계산 불가",
            })
            continue
        if not same_strength:
            rows.append({
                "함량": a_strength or "(미지정)",
                "신청의약품": app.get("item_name", ""),
                "신청약가": f"{app.get('price'):,.0f}원",
                "비교최저가": "—",
                "기준 비교군": "동일 함량 비교약 없음",
                "차이율": "계산 불가",
            })
            continue
        base = min(same_strength, key=lambda x: x.get("price"))
        diff = rule_validator.price_diff_percent(app.get("price"), base.get("price"))
        rows.append({
            "함량": a_strength or "(미지정)",
            "신청의약품": app.get("item_name", ""),
            "신청약가": f"{app.get('price'):,.0f}원",
            "비교최저가": f"{base.get('price'):,.0f}원",
            "기준 비교군": f"{base.get('group_label')} / {base.get('item_name')}",
            "차이율": diff,
        })
    return pd.DataFrame(rows) if rows else None


def style_price_summary(df):
    def fmt(val):
        if isinstance(val, (int, float)):
            sign = "+" if val > 0 else ""
            return f"{sign}{val:.1f}%"
        return val

    def color(val):
        if isinstance(val, (int, float)):
            if val > 0:
                return "color:red;font-weight:bold"
            if val < 0:
                return "color:blue;font-weight:bold"
        return ""

    styler = df.copy()
    try:
        return styler.style.format({"차이율": fmt}).map(color, subset=["차이율"])
    except AttributeError:
        return styler.style.format({"차이율": fmt}).applymap(color, subset=["차이율"])


init_group_state()

# ---------------- 사이드바 ----------------
with st.sidebar:
    st.markdown("## 💊 검증기 v4.3")
    st.caption("MFDS 허가사항 · HIRA 약가정보 대조 검증")

    mfds_key_in = st.text_input(
        "MFDS 인증키",
        type="password",
        value=secrets_value("MFDS_API_KEY", "DATA_GO_KOR_API_KEY"),
    )
    hira_key_in = st.text_input(
        "HIRA 인증키",
        type="password",
        value=secrets_value("HIRA_API_KEY", "DATA_GO_KOR_API_KEY"),
    )
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
                    progress_cb=lambda f, n: (bar.progress(f), status_txt.caption(f"MFDS 품목목록: {n:,}건 수집중...")),
                )
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
                        progress_cb=lambda f, n: (bar2.progress(f), status_txt2.caption(f"HIRA 약가: {n:,}건 수집중...")),
                    )
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
    st.caption("사용 흐름: ① 데이터 불러오기(1회) → ② 검색·선택/그룹지정/항목선택 → ③ 자동조회 → ④ 비교표 입력 → ⑤ 구조 확인 → ⑥ 1차 검증 → ⑦ Claude 자료 생성")

# ---------------- 본문 ----------------
st.title("의약품 심의자료 진위·오탈자 검증기 v4.3")
st.caption(
    "의약품 검색은 불러온 캐시(JSON) 기준 로컬 필터링으로 동작합니다. "
    "신청의약품과 비교의약품은 각각 여러 함량 품목을 가진 그룹으로 관리할 수 있습니다."
)

# ============ ① 검색·선택/그룹지정/항목선택 ============
st.markdown("## ① 의약품 검색·선택 및 조회 설정")
if not _SS.get("mfds_items"):
    st.warning("먼저 사이드바에서 「📥 전체 데이터 불러오기 (1회)」를 실행해 주세요. 이후 검색은 캐시 기준으로만 동작합니다.")

st.subheader("1. 의약품 검색")
query = st.text_input("의약품명 또는 제약사명", placeholder="예: 나르코설하정, 앱스트랄, 인스타닐", key="query_text")
matches = []
if query.strip() and _SS.get("mfds_items"):
    matches = cache_store.filter_products(_SS["mfds_items"], query)[:200]
    _SS["search_rows"] = matches
    st.caption(f"검색결과 {len(matches)}건 표시 (최대 200건)")
else:
    _SS["search_rows"] = []

if matches:
    search_df = pd.DataFrame([
        {"의약품명": row.get("ITEM_NAME", ""), "제약사": row.get("ENTP_NAME", ""), "품목코드": row.get("ITEM_SEQ", "")}
        for row in matches
    ])
    search_event = st.dataframe(
        search_df,
        use_container_width=True,
        hide_index=True,
        height=min(520, 36 + len(search_df) * 35),
        selection_mode="multi-row",
        on_select="rerun",
        key=f"search_results_table_{query.strip().casefold()}",
    )
    search_selected_seqs = [
        str(search_df.iloc[index]["품목코드"])
        for index in search_event.selection.rows
        if 0 <= index < len(search_df)
    ]
    add_col1, add_col2 = st.columns([3, 1])
    target_opts = group_target_options()
    target_gid = add_col1.selectbox(
        "선택 품목을 넣을 그룹",
        options=[gid for gid, _ in target_opts],
        format_func=lambda gid: dict(target_opts).get(gid, gid),
        key="target_group_id",
    )
    if add_col2.button("비교약 그룹 추가", key="add_comp_group"):
        gid = grouping.next_comparator_group_id(_SS["groups"])
        _SS["groups"][gid] = {"label": grouping.group_label(gid), "seqs": []}
        sync_selection_from_groups()
        st.rerun()
    st.caption(f"검색 결과에서 선택한 품목: **{len(search_selected_seqs)}건**")
    if st.button("선택한 검색 결과를 그룹에 추가", disabled=not search_selected_seqs, key="add_search_selection"):
        _SS["groups"] = grouping.assign_seqs_to_group(_SS["groups"], target_gid, search_selected_seqs)
        sync_selection_from_groups()
        st.rerun()

st.subheader("2. 그룹별 조회 품목")
by_seq = {str(row.get("ITEM_SEQ", "")): row for row in (_SS.get("mfds_items") or [])}
grouped_items = group_items(by_seq)
if grouped_items:
    for gid in grouping.ordered_group_ids(_SS["groups"]):
        meta = _SS["groups"].get(gid, {})
        label = meta.get("label") or grouping.group_label(gid)
        items = [x for x in grouped_items if x["group_id"] == gid]
        with st.expander(f"{label} ({len(items)}건)", expanded=True if gid == "applicant" else False):
            if not items:
                st.caption("아직 추가된 품목이 없습니다.")
                continue
            editor_df = pd.DataFrame([
                {
                    "의약품명": item["item_name"],
                    "제약사": item["entp_name"],
                    "품목코드": item["seq"],
                    "함량": item["strength"],
                    "자동파싱": item["parsed_strength"] or "(파싱 실패)",
                    "제거": False,
                }
                for item in items
            ])
            edited = st.data_editor(
                editor_df,
                key=f"group_editor_{gid}",
                use_container_width=True,
                hide_index=True,
                disabled=["의약품명", "제약사", "품목코드", "자동파싱"],
                column_config={
                    "함량": st.column_config.TextColumn(
                        "함량",
                        help="제품명에서 자동 파싱한 값입니다. 파싱 실패 또는 보정이 필요하면 직접 수정하세요.",
                    ),
                    "제거": st.column_config.CheckboxColumn("제거"),
                },
            )
            for row in edited.to_dict("records"):
                _SS["strength_overrides"][str(row["품목코드"])] = str(row.get("함량") or "").strip()
            remove_checked = [str(row["품목코드"]) for row in edited.to_dict("records") if row.get("제거")]
            if st.button(f"{label}에서 체크한 품목 제거", disabled=not remove_checked, key=f"remove_group_rows_{gid}"):
                _SS["groups"] = grouping.remove_seqs(_SS["groups"], remove_checked)
                for seq in remove_checked:
                    _SS["strength_overrides"].pop(str(seq), None)
                sync_selection_from_groups()
                st.rerun()
else:
    st.info("검색 결과 표에서 품목을 선택한 뒤 원하는 그룹으로 추가하세요.")

st.write(f"현재 선택된 전체 품목 수: **{len(grouping.flatten_group_seqs(_SS['groups']))}건**")
app_count = len(_SS["groups"].get("applicant", {}).get("seqs", []))
comp_counts = {gid: len(_SS["groups"].get(gid, {}).get("seqs", [])) for gid in grouping.comparator_group_ids(_SS["groups"])}
st.caption(
    " / ".join(
        [f"신청의약품 {app_count}건"] + [f"{_SS['groups'][gid]['label']} {count}건" for gid, count in comp_counts.items()]
    )
)

st.subheader("3. 추가 조회 항목")
if "preset_new_drug_intro" not in _SS:
    _SS["preset_new_drug_intro"] = all(_SS.get(f"extra_{key}", False) for key in NEW_DRUG_PRESET_KEYS)
if "preset_new_drug_intro_prev" not in _SS:
    _SS["preset_new_drug_intro_prev"] = _SS["preset_new_drug_intro"]

st.checkbox(
    "[신약 도입 준비]",
    key="preset_new_drug_intro",
    help="영문 제품명, ATC 코드, 포장단위, 주성분 영문명, 원료약품 및 분량(함량), 유효기간, 성상, 전문·일반의약품 구분, 보관정보, 보관 및 취급상의 주의사항을 한 번에 선택/해제합니다.",
)
if _SS["preset_new_drug_intro"] != _SS["preset_new_drug_intro_prev"]:
    for preset_key in NEW_DRUG_PRESET_KEYS:
        _SS[f"extra_{preset_key}"] = _SS["preset_new_drug_intro"]
    _SS["preset_new_drug_intro_prev"] = _SS["preset_new_drug_intro"]
    st.rerun()

extra_columns = st.columns(3)
for index, key in enumerate(EXTRA_FIELD_ORDER):
    with extra_columns[index % 3]:
        st.checkbox(EXTRA_FIELD_LABELS[key], key=f"extra_{key}")

# ============ ② 자동 조회 ============
st.markdown("## ② 허가사항·약가 자동 조회")
mfds_key_effective = mfds_key_in.strip() or secrets_value("MFDS_API_KEY", "DATA_GO_KOR_API_KEY")
hira_key_effective = hira_key_in.strip() or secrets_value("HIRA_API_KEY", "DATA_GO_KOR_API_KEY")
fetch_disabled = not grouping.flatten_group_seqs(_SS["groups"]) or not mfds_key_effective
if st.button("선택한 그룹 품목 조회", type="primary", disabled=fetch_disabled, key="fetch_btn"):
    grouped_items = group_items(by_seq)
    if not grouped_items:
        st.warning("조회할 품목을 먼저 그룹에 추가해 주세요.")
    elif not _SS["groups"].get("applicant", {}).get("seqs"):
        st.error("신청의약품 그룹에 최소 1개 품목이 필요합니다.")
    elif not mfds_key_effective:
        st.error("MFDS 인증키가 없습니다. 사이드바에 입력해 주세요.")
    else:
        selected_extras = selected_extra_keys()
        bar = st.progress(0.0)
        status_txt = st.empty()
        products = []
        for i, item in enumerate(grouped_items, start=1):
            row = by_seq.get(item["seq"])
            if not row:
                continue
            lb = label_of(row)
            seq = item["seq"]
            status_txt.info(f"({i}/{len(grouped_items)}) {item['group_label']} / {item['item_name']} → MFDS 상세 조회 중...")
            try:
                detail = mfds_api.get_product_detail(seq, mfds_key_effective)
                detail_ok = not detail.get("error")
            except mfds_api.MfdsApiError as e:
                detail = {"error": str(e)}
                detail_ok = False

            hira_rows = []
            match = {"status": drug_matcher.STATUS_FAIL, "row": None, "method": None}
            if detail_ok:
                if _SS.get("hira_items"):
                    match = drug_matcher.match_hira(detail, _SS["hira_items"])
                    hira_rows = [match["row"]] if match.get("row") else []
                elif hira_key_effective:
                    try:
                        hira_rows = hira_api.search_for_product(detail, hira_key_effective)
                        match = drug_matcher.match_hira(detail, hira_rows)
                    except hira_api.HiraApiError as e:
                        match = {"status": f"⚠ HIRA 조회 실패: {e}", "row": None, "method": None}
                else:
                    match = {"status": "⚠ HIRA 캐시·키 없음 — 사이드바에서 데이터 불러오기 실행", "row": None, "method": None}

            price = hira_api.get_price(match["row"]) if match.get("row") else None
            products.append({
                "label": lb,
                "role": item["group_label"],
                "group_id": item["group_id"],
                "group_label": item["group_label"],
                "item_name": item["item_name"],
                "entp_name": item["entp_name"],
                "strength": item["strength"],
                "parsed_strength": item["parsed_strength"],
                "seq": seq,
                "detail": detail,
                "hira_rows": hira_rows,
                "match": match,
                "price": price,
            })
            bar.progress(i / len(grouped_items))

        status_txt.empty()
        _SS["products"] = products
        _SS["fetched_df"] = build_source_result_df(products, selected_extras)
        _SS["validation_df"] = None
        st.success("조회 완료")

if _SS.get("products"):
    st.markdown("#### 조회 요약")
    rows = []
    for p in _SS["products"]:
        d = p["detail"]
        m = p["match"]
        rows.append({
            "구분": p.get("group_label", p["role"]),
            "함량": p.get("strength", ""),
            "제품": p.get("item_name") or p["label"].split(" | ")[0],
            "MFDS": "✅ 조회됨" if not d.get("error") else "❌ 실패",
            "HIRA": f"✅ {len(p['hira_rows'])}건" if p["hira_rows"] else ("⚠ 없음/실패" if not d.get("error") else "—"),
            "제품 매칭": m.get("status", drug_matcher.STATUS_FAIL),
            "약가(원)": f"{p['price']:,.0f}" if p["price"] is not None else "—",
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    result_df = _SS.get("fetched_df")
    if result_df is not None and not result_df.empty:
        result_tab, comparison_tab = st.tabs(["상세 결과", "여러 약품 비교표"])
        with result_tab:
            transpose_view = st.checkbox(
                "행/열 전환",
                key="source_result_transpose",
                value=True,
                help="켜면 약 하나가 세로줄 하나가 되고, 화면 폭에 맞춰 각 약의 너비가 자동으로 조정됩니다.",
            )
            displayed_df = make_display_df(result_df, transpose_view=transpose_view)
            render_resizable_wrapped_table(displayed_df, show_index=transpose_view, height=720, table_key="detail")
        with comparison_tab:
            if len(result_df) < 2:
                st.info("두 품목 이상 조회하면 비교표가 표시됩니다.")
            else:
                comparison_df = make_comparison_df(result_df)
                st.caption("행은 조회 항목, 열은 그룹-함량별 의약품입니다. 헤더 경계를 드래그해 약품별 컬럼 너비를 조절할 수 있습니다.")
                render_resizable_wrapped_table(comparison_df, show_index=True, height=760, table_key="comparison")

    price_summary = build_group_price_summary(_SS["products"])
    if price_summary is not None and not price_summary.empty:
        st.markdown("#### 💰 함량별 약가 비교 (신청의약품 vs 동일 함량 비교약 최저가)")
        st.dataframe(style_price_summary(price_summary), use_container_width=True)
        st.caption("동일 함량 비교약이 있는 경우에만 차이율을 계산합니다. 양수는 상승(빨강), 음수는 하락(파랑)입니다.")

    for p in _SS["products"]:
        with st.expander(f"📄 원문 보기 — {p['group_label']} / {p['strength'] or '(함량미지정)'} / {p['item_name']} (원문, 요약 아님)"):
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
                "연결 제품": (f"{product.get('group_label')} / {product.get('strength')} / {product.get('item_name')}" if product else "❌ 원문 매칭 안 됨"),
                "항목": pair.get("field", ""),
                "1차 판정": status,
                "사유": reason,
            })
        _SS["validation_df"] = pd.DataFrame(validation)
    if _SS.get("validation_df") is not None:
        vdf = _SS["validation_df"]
        st.dataframe(style_df(vdf, "1차 판정"), use_container_width=True)
        st.caption("🟠 Claude 확인 필요 항목은 Python이 판정하지 않습니다. ⑥ 단계 자료를 Claude 웹에 붙여넣어 의미 검증하세요.")
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
            "prompt_label": f"{p['group_label']} / {p['strength'] or '(함량미지정)'} / {p['item_name']}",
            "role": p["group_label"],
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
        with st.expander(f"제품별 — {p['prompt_label']}"):
            prod_prompt = claude_prompt_builder.build_product_prompt(p)
            copy_button(prod_prompt, f"📋 [{p['prompt_label']} 복사]")
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
    "API 오류·캐시 부재 시 데이터를 임의로 생성하지 않습니다."
)
