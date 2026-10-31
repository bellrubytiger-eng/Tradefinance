# GBS India 외환·무역송금 코어뱅킹 설계

노이다지점(AD Category-I)의 당발·타발송금, LRS, FEMA 수출입(IDPMS/EDPMS), FDI 외환송금, INR Vostro(MT950), RBI CIMS 보고를 위한 요건정의·데이터 설계·화면 시안.

## 산출물

| 파일 | 내용 |
|---|---|
| `docs/01_요건정의_설계서.docx` | 요건정의 및 시스템 설계서 (규제요약, TO-BE 프로세스, 화면, 데이터 관리규칙, 기한관리, 인터페이스, 회계, 기존 요청 반영·검토의견, 이행계획) |
| `docs/02_데이터사전_업무규칙.xlsx` | 화면목록, 테이블/컬럼 정의(32 테이블), 업무규칙, 기한관리 매트릭스, 규정 파라미터, 기한계산기(수식), 코드, 인터페이스, 회계분개, 요청사항 추적표 |
| `db/ddl_oracle.sql` | Oracle DDL (테이블·코멘트·인덱스·FK·파라미터 초기적재·보고용 뷰 예시) |
| `screens/index.html` | HTML 화면 시안 14종 (브라우저로 바로 열기). `screens/app.html`은 Artifact 게시용 원본 |

## 재생성

모든 문서는 `tools/model.py` 한 곳의 정의에서 생성됩니다. 테이블·규칙·파라미터를 바꿀 때는 `model.py`만 수정한 뒤 다시 생성하십시오.

```bash
cd tools
python3 model.py /tmp/model.json
python3 build_xlsx.py "../docs/02_데이터사전_업무규칙.xlsx"
python3 build_ddl.py ../db/ddl_oracle.sql
node build_docx.js /tmp/model.json "../docs/01_요건정의_설계서.docx"
python3 build_html.py
```

`VERIFY=Y` 표시 항목(규정 수치·Purpose Code·RBI 파일 레이아웃)은 운영 반영 전 컴플라이언스 확인이 필요합니다.
