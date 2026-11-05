# GBS India 외환·무역송금 코어뱅킹 설계

노이다지점(AD Category-I)의 당발·타발송금, LRS, 비거주자 송금, FEMA 수출입(IDPMS/EDPMS), FDI(반환·INR 확정·에스크로), 비계좌 중계(미지급금·NEFT/RTGS), 환전, INR Vostro, SWIFT 메시지, RA 대지급·본지점 정산, RBI CIMS 보고를 위한 요건정의·데이터 설계·화면·서식 시안.

## 산출물

| 파일 | 내용 |
|---|---|
| `docs/01_요건정의_설계서.docx` | 요건정의 및 시스템 설계서 (규제, 프로세스 13종, 화면, FDI 화면 가이드, 데이터 규칙, 기한관리, 인터페이스·SWIFT 카탈로그, GST 기준, 회계, 검토의견, 이행계획) |
| `docs/02_데이터사전_업무규칙.xlsx` | 화면목록, 테이블/컬럼 정의(80 테이블), 업무규칙(207), 기한관리, 규정 파라미터, 계산기(기한·LRS·LEI·GST·FDI 수식), 코드, 인터페이스, 회계분개 요약, 요청사항 추적, 서식필드매핑, FDI 화면 가이드, SWIFT 카탈로그, RBI 기준 검토, 규정 근거 색인, **통합 표준 CoA(인도지점 신규=주황·변경·비활성·정정권고, 메모)**, **업무별 예시 분개(통화별 대차 검증 수식)** |
| `docs/03_고객서식_명세서.docx` | 고객 서식 4종 명세 (FIRC, 당발·타발·환전 확인서): 발급 원칙, 항목·DB 소스, 문구 |
| `docs/04_화면흐름도.pptx` | 화면 연결 다이어그램 13장 (전체 지도, 당발·타발·FDI 초과반환·무역금융·LC/BG·본지점/RA·Vostro/SWIFT/환전, 서식·보고 허브, 화면 목록) |
| `docs/05_CoA검토_RBI보고속성.xlsx` | **통합 표준 CoA(0930) 검토의견**: GL vs 속성 판정원칙·Third Schedule 대응, 예금명세서 GBS 8자리 코드 충돌 자동대조, 계정 신규·변경·속성전환·비활성(이전 반영안 정정 포함), 예금상품→계정 매핑, RBI 보고용 속성 57개, 보고서별 요구, PSL·INR Nostro·FN/NRO 상세, **현행 인도 GBS(2026-07-07 BS/PL) → 11자리 이행표(대차 검증)**, **ECL·RWA 개발명세 계정 매핑**, 결정사항, **LCR 분류·파라미터(TB_LCR_*)**, **적금 L3·TDS/TCS/GST 11자리(신 소득세법·GST FRS)·고정자산(Schedule II, FA코드 연결)·영업비용 Sch.16 매핑(EP코드)**, 시트별 '본점 원안 / 변경 / 이유 / FY25-26 결산 증거' 블록과 결산근거 시트, 모든 표의 「인도 용어 정의」 열(시트 17 용어집) |
| `db/ddl_oracle.sql` | Oracle DDL (테이블·코멘트·인덱스·FK·파라미터 초기적재·보고용 뷰 예시) |
| `screens/index.html` | HTML 화면 시안 26종 (브라우저로 바로 열기). `screens/app.html`은 Artifact 게시용 원본 |
| `screens/forms.html` | 고객 서식 A4 인쇄 시안 4종·14개 변형 (DB 매핑 표시, DUPLICATE 보기) |

## 재생성

모든 문서는 `tools/model.py`(데이터 모델·규칙·화면), `tools/tf_model.py`(무역금융), `tools/forms_model.py`(서식), `tools/coa_model.py`(통합 CoA·인도지점 계정·업무별 분개, 원본 `tools/data/coa_base.json`) 정의에서 생성됩니다. 내용을 바꿀 때는 이 파일들만 수정한 뒤 다시 생성하십시오. `python3 coa_model.py`는 분개의 통화별 대차와 계정 유효성(비활성·직접기장 금지)을 검증합니다.

```bash
cd tools
python3 model.py /tmp/model.json
python3 build_xlsx.py "../docs/02_데이터사전_업무규칙.xlsx"
python3 build_ddl.py ../db/ddl_oracle.sql
python3 build_coa_review.py   # CoA 검토의견 (원천 coa_review_model.py, 이행·ECL은 coa_migration_model.py + data/gbs_live_20260707.json, 세무·고정자산·비용은 tax_fa_model.py + data/fa_code_list.json·ep_code_list.json, LCR은 lcr_model.py, 시트별 근거는 sheet_doc_model.py, 용어 정의는 glossary_model.py)
node build_docx.js /tmp/model.json "../docs/01_요건정의_설계서.docx"
python3 build_forms.py
python3 build_html.py
mkdir -p /tmp/fimg && node shot_forms.js /tmp/fimg
node build_forms_docx.js /tmp/model.json /tmp/fimg "../docs/03_고객서식_명세서.docx"
# pptxgenjs 필요 (npm install pptxgenjs 후 NODE_PATH 지정), PPTX_SKILL_DIR 미지정 시 테마 적용 생략
node build_pptx.js /tmp/model.json "../docs/04_화면흐름도.pptx"
```

`VERIFY=Y` 표시 항목(규정 수치·Purpose Code·RBI 파일 레이아웃·SWIFT MX 일정·GST 공급장소)은 운영 반영 전 컴플라이언스·세무 확인이 필요합니다.
