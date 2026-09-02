# 의약품 심의자료 진위·오탈자 검증기 v4.0

병원 약제부가 신규의약품 심의자료(비교표)를 작성할 때, 그 내용을 **식품의약품안전처(MFDS) 허가사항**과
**건강보험심사평가원(HIRA) 약가정보** 원문과 대조하여 잘못된 내용·누락·과장 표현을 빠르게 찾아내는
Streamlit 웹앱입니다.

- **기계적으로 확실한 것만 Python이 자동 판정**: 기본정보(제품명/성분명/제조판매사/제형), 숫자·단위 오탈자, 약가 차이율
- **의미 비교는 Claude 웹에서 수행**: 이 앱은 Claude 웹(claude.ai)에 붙여넣을 검증 자료를 자동 생성하고
  클립보드로 복사하는 역할만 합니다. **Claude API 호출 코드는 없습니다.**
- 로컬 Python 설치 불필요 — **GitHub 저장소 → Streamlit Community Cloud 배포** 전제

## 핵심 원칙

| 항목 | 결정 |
|---|---|
| AI 호출 | Claude API 호출 코드·API Key 입력창 **없음** (판정은 사용자가 Claude 웹에서 수행) |
| CSV | 필수 중간 파일로 사용하지 않음 — 앱 안에서 검색→선택→즉시 API 조회 |
| 데이터 저장 | 서버 영구 저장 없음 (Streamlit 세션 범위만 유지) |
| API 키 | 소스코드에 하드코딩하지 않고 **Streamlit Secrets** 사용 |
| 제품 매칭 | MFDS 바코드 8자리(`BAR_CODE[3:11]`) == HIRA `mdsCd` 앞 8자리 → 품목명 완전 일치 → 부분 일치(🟠 사용자 확인) |
| 오류 처리 | API 오류 시 가짜 데이터를 절대 생성하지 않고 원인 메시지 표시 |

## 프로젝트 구조

```
drug-review-validator/
├── app.py                       # Streamlit 엔트리포인트 (STEP 1~6)
├── modules/
│   ├── mfds_api.py              # MFDS 목록/상세 조회 (중첩 XML 파싱, NB 섹션 분리)
│   ├── hira_api.py              # HIRA 약가 조회 (payTpNm='삭제' 제외)
│   ├── drug_matcher.py          # 8자리 바코드 규칙 기반 매칭
│   ├── pptx_parser.py           # python-pptx 표 추출 (병합 정보 보존)
│   ├── xlsx_parser.py           # openpyxl 표 추출 (merged_cells 보존)
│   ├── clipboard_parser.py      # 복사·붙여넣기 표 추출 (탭 구분)
│   ├── table_normalizer.py      # 3가지 입력 → 공통 스키마 변환 + 방향 추정
│   ├── rule_validator.py        # Python 1차 규칙 검증 (기계적 판정만)
│   ├── claude_prompt_builder.py # Claude용 검증 자료 생성 (API 호출 없음)
│   └── result_parser.py         # Claude 웹 결과(마크다운/JSON) 재붙여넣기 파싱
├── requirements.txt
└── README.md
```

## 외부 API (엔드포인트 고정 — 추측 금지)

| API | URL |
|---|---|
| MFDS 목록 | `http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07/getDrugPrdtPrmsnInq07` |
| MFDS 상세 | `https://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07/getDrugPrdtPrmsnDtlInq06` |
| HIRA 약가 | `https://apis.data.go.kr/B551182/dgamtCrtrInfoService1.2/getDgamtList` |

데이터는 공공데이터포털(data.go.kr)에서 **서비스 키를 발급**받아야 합니다:
- MFDS: 「의약품 제품 허가정보」서비스 (1471000)
- HIRA: 「약가정보」서비스 (B551182)
- 하나의 data.go.kr 키로 두 서비스를 모두 활용 가능(활용신청 필요)

## 사용 흐름 (STEP 1~6)

1. **제품 검색·선택** — 제품명/제조사명 부분 검색 → 신청의약품 1개 + 비교의약품 N개 지정
2. **자동 조회** — MFDS 상세 허가사항 + HIRA 약가 + 제품 매칭(8자리 바코드 규칙), 원문은 펼쳐서 확인
3. **비교표 입력** — PPTX 업로드 / XLSX 업로드 / 복사·붙여넣기 (3-way)
4. **구조 확인** — 행=항목/열=제품 방향 자동 추정, 틀리면 수동 지정, 공통 스키마로 변환
5. **1차 자동 검증** — 기본정보·숫자단위·약가만 기계 판정, 나머지는 🟠 Claude 확인 필요
6. **Claude 검증 자료 생성** — 전체/제품별/항목별 복사 버튼(JS 클립보드) → Claude 웹에 붙여넣기 →
   결과(마크다운/JSON)를 다시 붙여넣으면 판정 표로 렌더링

## Streamlit Cloud 배포

### 1) GitHub 저장소에 업로드

```bash
git init
git add .
git commit -m "drug review validator v4.0"
git branch -M main
git remote add origin https://github.com/<사용자명>/drug-review-validator.git
git push -u origin main
```

> `.streamlit/secrets.toml`과 `secrets.toml.example`의 실제 키는 커밋하지 마세요 — `.gitignore`에 포함됩니다.

### 2) Streamlit Community Cloud 연결

1. https://share.streamlit.io 접속 → GitHub 계정으로 로그인
2. **New app** → 리포지토리 `drug-review-validator` / 브랜치 `main` / 메인 파일 `app.py`
3. **Deploy** 클릭 → 빌드 완료 후 앱이 열립니다

### 3) Secrets 설정 (API 키)

앱 화면 우측 상단 메뉴 **⋮ → Settings → Secrets** 에 아래 내용을 붙여넣고 저장:

```toml
# Streamlit Secrets
MFDS_API_KEY = "여기에 MFDS(의약품 제품 허가정보) 데이터활용신청 API 인증키"
HIRA_API_KEY = "여기에 HIRA(약가정보) 데이터활용신청 API 인증키"

# 두 서비스를 같은 data.go.kr 키로 신청했다면 아래 하나로 통일 사용 가능:
# DATA_GO_KOR_API_KEY = "여기에 data.go.kr API 인증키"
```

저장 후 앱이 자동으로 재시작됩니다. 키가 없으면 앱은 안내 문구만 표시하고 API 호출을 하지 않습니다.

### 4) (선택) 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 개발 명세 준수 확인 (v4.0 명세서)

- [x] Claude API 호출 코드 없음 (전체 파이썬 소스 grep 검사 0건)
- [x] API Key 하드코딩 없음 — `st.secrets.get(...)` 으로만 讀取
- [x] CSV 필수 중간 파일 없음 — 앱 내 검색→선택→API 즉시 조회
- [x] MFDS/HIRA 엔드포인트·필드명·8자리 매칭 규칙 명세 그대로 사용
- [x] PPTX/XLSX 병합 셀 정보 보존, 복사·붙여넣기 표 인식
- [x] 명세서 8장 공통 셀 스키마로 정규화
- [x] 서술형 항목은 자동 판정 금지 → Claude 확인 필요로 전환
- [x] 약가 차이율 공식 `(신청-최저)/최저×100`, 양수=빨강/음수=파랑
- [x] 오류 시 가짜 데이터 미생성 — 원인 메시지 표시

## 라이선스·면책

- MFDS/HIRA 공공데이터를 사용하며, 원문은 요약 없이 그대로 표시·전달합니다.
- 본 앱의 자동 판정은 **기계적(문자열/숫자) 비교** 한정입니다. 최종 심의 판단은 담당 약사의 검토를 거쳐야 합니다.
