# -*- coding: utf-8 -*-
"""
무역금융(여신) 모듈 통합 정의
원천: 지점 전산개발요청서(어음할인·오픈어카운트·D/A·D/P·TReDS, 2026.8.13), CBS 화면설계서 v6,
      SFMS Interface Requirements (Draft 0.1, 2026.10.6), SFMS 화면 와이어프레임
추가: 신용장(LC) · 보증(BG) — 업로드 자료 범위 밖이어서 RBI 최신 기준으로 신규 설계
model.py 가 import 해서 TABLES/RULES/PARAMS/... 에 병합한다.
"""

AUD = [
    ("REG_USER_ID", "등록자(Maker)", "VARCHAR2(10)", "NN", "", ""),
    ("REG_TS", "등록일시", "TIMESTAMP", "NN", "", ""),
    ("UPD_USER_ID", "수정자", "VARCHAR2(10)", "", "", ""),
    ("UPD_TS", "수정일시", "TIMESTAMP", "", "", ""),
    ("APPR_USER_ID", "승인자(Checker)", "VARCHAR2(10)", "", "Maker와 동일인 불가", ""),
    ("AUTH_USER_ID", "Authorizer", "VARCHAR2(10)", "", "ALCO 기준금액 초과 시", ""),
]

TABLES = []


def table(name, kor, area, desc, cols, idx=None):
    TABLES.append({"name": name, "kor": kor, "area": area, "desc": desc, "cols": cols, "idx": idx or []})


# ---------------------------------------------------------------------------
# 1. 거래 원장 (BILL_MASTER)
# ---------------------------------------------------------------------------
table("TB_TF_BILL", "무역금융 거래(BILL_MASTER)", "무역금융(TF11)",
      "어음할인·매입·LC 매입·D/A·D/P 추심·오픈어카운트·TReDS 통합 원장. 거래유형별 하위 테이블 1:1.", [
    ("BILL_REF", "어음참조번호", "VARCHAR2(20)", "PK", "BR-지점-YYMMDD-일련 / TReDS는 TR-플랫폼-YYMMDD-일련. SFMS E2E ID로 사용(35자 이내)", "SFMS"),
    ("BILL_CATEGORY", "거래유형", "VARCHAR2(14)", "NN", "CODE:BILL_CATEGORY", ""),
    ("BILL_TYPE", "어음유형", "VARCHAR2(14)", "NN", "EXPORT_USANCE/EXPORT_SIGHT/IMPORT_USANCE/IMPORT_SIGHT/INLAND_DEMAND/INLAND_USANCE/CLEAN", ""),
    ("DIRECTION", "무역방향", "VARCHAR2(6)", "NN", "EXPORT/IMPORT/INLAND", ""),
    ("CUST_NO", "고객번호", "VARCHAR2(8)", "FK,NN", "한도·여신 심사 대상(수출자/수입자)", ""),
    ("DRAWER_NAME", "발행인", "VARCHAR2(140)", "NN", "", ""),
    ("DRAWEE_NAME", "지급인", "VARCHAR2(140)", "NN", "", ""),
    ("DRAWEE_CTRY", "지급인 국가", "VARCHAR2(2)", "", "국가한도·국가 프리미엄", ""),
    ("COUNTERPARTY_BANK_BIC", "추심·개설은행", "VARCHAR2(11)", "", "LC 매입은 개설은행(신용 프리미엄 평가 대상)", ""),
    ("LC_REF", "신용장 참조", "VARCHAR2(16)", "FK", "LC 기반 매입(NEGOTIATION)·인수", ""),
    ("CCY", "통화", "VARCHAR2(3)", "NN", "", "EDPMS"),
    ("BILL_AMT", "어음금액", "NUMBER(18,2)", "NN", "", "EDPMS"),
    ("FX_RATE", "적용환율", "NUMBER(15,6)", "", "FEDAI 고시 대비 편차 검증", ""),
    ("INR_EQV", "INR 환산", "NUMBER(18,2)", "NN", "", "CRILC"),
    ("BILL_DT", "어음일자", "DATE", "NN", "", ""),
    ("TENOR_DAYS", "만기일수", "NUMBER(4)", "", "", ""),
    ("DUE_DT", "만기일", "DATE", "", "휴일이면 영업일 조정(R-HOL-01)", ""),
    ("SHIPMENT_DT", "선적일", "DATE", "", "회수기한·유전스 기산", "EDPMS/IDPMS"),
    ("REALIZE_DUE_DT", "회수기한", "DATE", "", "수출: 선적일 + TF_EXP_REALIZE_MONTHS", "EDPMS"),
    ("SB_MGMT_NO", "Shipping Bill", "VARCHAR2(16)", "FK", "수출 - TB_TRD_SHIPPING_BILL 연결(EDPMS)", "EDPMS"),
    ("BOE_MGMT_NO", "Bill of Entry", "VARCHAR2(16)", "FK", "수입 - TB_TRD_BOE 연결(IDPMS)", "IDPMS"),
    ("HS_CODE", "HS Code", "VARCHAR2(8)", "", "수입 유전스 상한(자본재 여부) 판정", ""),
    ("PURPOSE_CODE", "Purpose", "VARCHAR2(5)", "", "", "FETERS"),
    ("ECGC_COVER_YN", "ECGC 부보", "VARCHAR2(1)", "", "리스크 프리미엄 차등", ""),
    ("PSL_CAT", "우선대출 분류", "VARCHAR2(10)", "", "EXPORT/MSME/NONE (TReDS MSME 자동)", "PSL"),
    ("BILL_STS", "상태", "VARCHAR2(10)", "NN", "DRAFT/SUBMIT/APPROVED/DISBURSED/MATURED/OVERDUE/REALIZED/CRYSTAL/NPA/CLOSED/CANCEL", ""),
] + AUD, idx=["CUST_NO,BILL_STS", "DUE_DT", "LC_REF"])

table("TB_TF_DOC", "선적서류", "무역금융(TF11)", "B/L·Invoice·P/L·보험·S/B 등 서류 체크리스트(오픈어카운트는 선택).", [
    ("BILL_REF", "어음참조번호", "VARCHAR2(20)", "PK,FK", "", ""),
    ("SEQ", "순번", "NUMBER(3)", "PK", "", ""),
    ("DOC_TYPE", "서류유형", "VARCHAR2(6)", "NN", "BL/AWB/INV/PL/INS/SB/BOE/COO/BOX", ""),
    ("DOC_NO", "서류번호", "VARCHAR2(35)", "", "", ""),
    ("DOC_DT", "서류일자", "DATE", "", "", ""),
    ("ORIG_CNT", "원본 수", "NUMBER(2)", "", "", ""),
    ("COPY_CNT", "사본 수", "NUMBER(2)", "", "", ""),
    ("DOC_STS", "상태", "VARCHAR2(6)", "NN", "RCVD/CHKD/MISS/DISC(하자)", ""),
])

# ---------------------------------------------------------------------------
# 2. 할인 계산 엔진 (M2) - 5단계 금리 + 개설은행 평가 + 정책 프리미엄
# ---------------------------------------------------------------------------
table("TB_TF_DISCOUNT", "할인계산", "무역금융(TF11)", "기준금리·기간스프레드·신용·국가·정책 프리미엄 5단계와 공제액·순지급액. 산출 근거 전부 저장.", [
    ("BILL_REF", "어음참조번호", "VARCHAR2(20)", "PK,FK", "", ""),
    ("RATE_TYPE", "금리유형", "VARCHAR2(6)", "NN", "MCLR/EBLR/SOFR/EURIBOR/FIXED", ""),
    ("BASE_SRC", "기준금리 출처", "VARCHAR2(20)", "NN", "INR: MCLR 1년물(또는 EBLR) / FCY: SOFR·EURIBOR + 조달마진", ""),
    ("BASE_RATE", "① 기준금리", "NUMBER(9,6)", "NN", "적용 하한(R-TF-12)", ""),
    ("TENOR_SPREAD", "② 기간 스프레드", "NUMBER(9,6)", "NN", "≤90일 0.25 / 91~180 0.50 / >180 0.85 (ALCO)", ""),
    ("CREDIT_PREM", "③ 신용 프리미엄", "NUMBER(9,6)", "NN", "LC: 개설은행 등급 / 그 외: 지급인·수출자 등급", ""),
    ("ISSUER_GRADE", "개설은행 등급", "VARCHAR2(1)", "", "S/A/B/C/D/E (TB_TF_ISSUER_EVAL)", ""),
    ("COUNTRY_PREM", "④ 국가 프리미엄", "NUMBER(9,6)", "NN", "OECD CRC 0~7", ""),
    ("POLICY_PREM", "⑤ 정책 프리미엄", "NUMBER(9,6)", "NN", "-0.50 ~ +1.00, 결재선 차등(TB_TF_POLICY_ADJ)", ""),
    ("EFF_RATE", "적용금리", "NUMBER(9,6)", "NN", "= ①+②+③+④+⑤, 하한·상한 검증", ""),
    ("DAYCOUNT", "일수기준", "VARCHAR2(8)", "NN", "INR ACT/365, FCY ACT/360 (관리자 변경 가능)", ""),
    ("DISC_DAYS", "할인일수", "NUMBER(4)", "NN", "", ""),
    ("DISC_AMT", "할인료", "NUMBER(18,2)", "NN", "= 액면 × 적용금리 × 일수 / 일수기준", ""),
    ("COMMISSION", "취급수수료", "NUMBER(18,2)", "NN", "", ""),
    ("GST_AMT", "GST", "NUMBER(18,2)", "NN", "수수료·전신료 과세(R-GST-*). 할인료(이자)는 비과세", "GST"),
    ("MSG_CHG", "전신료", "NUMBER(18,2)", "", "SWIFT/SFMS", ""),
    ("NET_PROCEEDS", "순지급액", "NUMBER(18,2)", "NN", "= 액면 - 할인료 - 수수료 - GST - 전신료 → A/P 계상", ""),
    ("DISC_PCT", "할인비율", "NUMBER(7,4)", "", "", ""),
    ("TOTAL_DED_PCT", "총공제비율", "NUMBER(7,4)", "", "", ""),
    ("NET_PCT", "순지급비율", "NUMBER(7,4)", "", "", ""),
    ("INCOME_RECOG", "수익인식", "VARCHAR2(7)", "NN", "ACCRUAL(미실현할인 일할) / UPFRONT", ""),
] + AUD)

table("TB_TF_ISSUER_EVAL", "개설은행 평가", "무역금융(TF41)", "개설은행 3축(바젤III 40·외부등급 35·거래이력 25) 100점 → S~E 등급. 반기 재평가 + 이벤트 재평가.", [
    ("BANK_BIC", "개설은행 BIC", "VARCHAR2(11)", "PK", "", ""),
    ("EVAL_DT", "평가일", "DATE", "PK", "", ""),
    ("BANK_CTRY", "소재국", "VARCHAR2(2)", "NN", "", ""),
    ("AX1_SCORE", "제1축 재무건전성", "NUMBER(5,2)", "", "CRAR·CET1·NPL·LCR·NSFR·레버리지 (Pillar 3). 자료 없으면 0, 제2축 55%로 상향", ""),
    ("AX2_SCORE", "제2축 외부등급", "NUMBER(5,2)", "", "Moody's/S&P/Fitch/CRISIL, 국가등급 상한", ""),
    ("AX3_SCORE", "제3축 당행 거래이력", "NUMBER(5,2)", "", "최근 24개월 승인불가(IFN/MT 734)·지급지연·분쟁 자동집계", ""),
    ("REFUSAL_CNT_24M", "거절(734) 건수", "NUMBER(4)", "", "TB_TF_LC_PRES·TB_SFMS_MSG(IFN 734)에서 자동", ""),
    ("TOTAL_SCORE", "종합점수", "NUMBER(5,2)", "NN", "", ""),
    ("GRADE", "등급", "VARCHAR2(1)", "NN", "S≥90 / A≥75 / B≥60 / C≥45 / D≥30 / E<30", ""),
    ("BASE_PREMIUM", "기본 프리미엄", "NUMBER(6,4)", "", "S 0.15 / A 0.35 / B 0.60 / C 1.00 / D 1.75 / E 취급불가", ""),
    ("SINGLE_BANK_CAP_PCT", "단일은행 상한(기본자본 %)", "NUMBER(5,2)", "", "S 10 / A 8 / B 5 / C 3 / D 1 / E 0", ""),
    ("NEXT_EVAL_DT", "차기 평가일", "DATE", "NN", "반기. 등급 변동·부도·국가등급 하향 시 즉시", ""),
] + AUD)

table("TB_TF_POLICY_ADJ", "정책 프리미엄 조정 이력", "무역금융(TF11)", "재량 조정의 조정 전후·사유·결재선. 변경 불가 이력 10년 보존.", [
    ("ADJ_ID", "조정ID", "NUMBER(12)", "PK", "", ""),
    ("BILL_REF", "어음참조번호", "VARCHAR2(20)", "FK,NN", "", ""),
    ("RATE_BEFORE", "조정 전 금리", "NUMBER(9,6)", "NN", "", ""),
    ("ADJ_PCT", "조정 폭", "NUMBER(7,4)", "NN", "", ""),
    ("TIER", "결재 구간", "VARCHAR2(2)", "NN", "T1 ±0.10 팀장 / T2 -0.25~+0.30 부지점장 / T3 -0.40~+0.60 지점장 / T4 -0.50~+1.00 지점장+CRO / 초과 차단", ""),
    ("REASON_CD", "사유코드", "VARCHAR2(6)", "NN", "인하 D1~D5 / 인상 U1~U4", ""),
    ("REASON_TXT", "사유 서술", "VARCHAR2(500)", "", "T2 이상 필수", ""),
    ("APPROVER", "결재자", "VARCHAR2(30)", "NN", "", ""),
    ("APPR_TS", "결재시각", "TIMESTAMP", "NN", "", ""),
])

# ---------------------------------------------------------------------------
# 3. 상품별 하위 테이블
# ---------------------------------------------------------------------------
table("TB_TF_COLLECTION", "추심 (D/A·D/P)", "무역금융(TF11)", "URC 522. 발송/도착 추심, 인수·지급, 거절증서, 독촉.", [
    ("BILL_REF", "어음참조번호", "VARCHAR2(20)", "PK,FK", "", ""),
    ("COLL_TYPE", "추심유형", "VARCHAR2(2)", "NN", "DA/DP", ""),
    ("COLL_DIR", "추심방향", "VARCHAR2(8)", "NN", "OUTWARD 발송 / INWARD 도착", ""),
    ("REMITTING_BANK_REF", "송금은행 참조", "VARCHAR2(16)", "", "", ""),
    ("PRESENTING_BANK_BIC", "제시은행", "VARCHAR2(11)", "", "", ""),
    ("PROTEST_YN", "거절증서", "VARCHAR2(1)", "NN", "", ""),
    ("COLL_MSG", "SWIFT/SFMS 메시지", "VARCHAR2(6)", "", "MT400/410/412/416/420/430 (Cat4 MT 유지)", ""),
    ("DISPATCH_DT", "서류발송일", "DATE", "", "", ""),
    ("ACCEPT_DT", "인수일(D/A)", "DATE", "", "인수 후 할인 가능", ""),
    ("TRACER_CNT", "독촉 횟수", "NUMBER(3)", "NN", "EOD ⑧ 독촉 초안", ""),
] + AUD)

table("TB_TF_OPEN_ACCT", "오픈어카운트", "무역금융(TF11)", "후지급·선지급·CAD·위탁·유전스송금. 은행은 AD로서 회수·송금 감독(자금 미제공 시 한도 미차감).", [
    ("BILL_REF", "어음참조번호", "VARCHAR2(20)", "PK,FK", "", ""),
    ("OA_TYPE", "OA 유형", "VARCHAR2(14)", "NN", "DEFERRED_PAYMENT/ADVANCE/CAD/CONSIGNMENT/REMITTANCE_USANCE", ""),
    ("INV_MGMT_NO", "인보이스", "VARCHAR2(16)", "FK", "TB_TRD_INVOICE 연결", ""),
    ("PAY_TERMS", "결제조건", "VARCHAR2(60)", "", "", ""),
    ("PAY_DUE_DT", "결제기한", "DATE", "NN", "", ""),
    ("TRADE_CREDIT_YN", "무역신용 해당", "VARCHAR2(1)", "NN", "수입 결제기한 > 선적일 + 6개월 → Y (R-TF-06)", "TC 보고"),
    ("TC_REPORT_REF", "무역신용 보고 Ref", "VARCHAR2(30)", "", "", "TC 보고"),
] + AUD)

table("TB_TF_TREDS", "TReDS", "무역금융(TF11)", "RXIL·M1xchange·Invoicemart FU 수신·입찰·낙찰·정산. MSME PSL 자동 분류.", [
    ("BILL_REF", "어음참조번호", "VARCHAR2(20)", "PK,FK", "", ""),
    ("PLATFORM", "플랫폼", "VARCHAR2(10)", "NN", "RXIL/M1X/INVOICEMART", "TReDS"),
    ("FU_NO", "FU 번호", "VARCHAR2(30)", "NN", "", "TReDS"),
    ("SELLER_UDYAM", "매출자 Udyam", "VARCHAR2(19)", "NN", "UDYAM-XX-00-0000000", "PSL"),
    ("BUYER_NAME", "매입자", "VARCHAR2(140)", "NN", "", ""),
    ("BUYER_ACCEPT_DT", "매입자 인수일", "DATE", "NN", "인수 후 입찰 가능", ""),
    ("BID_MODE", "입찰방식", "VARCHAR2(4)", "NN", "AUTO/MANU", ""),
    ("BID_RATE", "입찰금리", "NUMBER(9,6)", "", "", ""),
    ("BID_STS", "입찰상태", "VARCHAR2(6)", "NN", "OPEN/WON/LOST/EXPD", ""),
    ("SETTLE_DT", "정산일", "DATE", "", "플랫폼 정산(SFMS 범위 외)", ""),
    ("RECON_STS", "정산대사", "VARCHAR2(5)", "", "EOD ⑦", ""),
] + AUD)

# ---------------------------------------------------------------------------
# 4. 한도 (M3)
# ---------------------------------------------------------------------------
table("TB_TF_LIMIT", "여신한도", "무역금융(TF31)", "전체·상품별·단일거래상대방·국가·일중 5층 한도. 거래 등록 시 전층 동시 검증.", [
    ("LIMIT_ID", "한도ID", "VARCHAR2(16)", "PK", "", ""),
    ("LIMIT_LEVEL", "한도 층", "VARCHAR2(8)", "NN", "OVERALL/PRODUCT/CPTY/COUNTRY/INTRADAY", ""),
    ("PARENT_LIMIT_ID", "상위 한도", "VARCHAR2(16)", "FK", "", ""),
    ("CUST_NO", "고객", "VARCHAR2(8)", "FK", "OVERALL·PRODUCT", ""),
    ("PRODUCT", "상품", "VARCHAR2(14)", "", "BD/BP/LC_IMP/LC_NEGO/BG_FIN/BG_PERF/TREDS ...", ""),
    ("CPTY_KEY", "거래상대방/국가", "VARCHAR2(20)", "", "지급인·개설은행 BIC·국가코드", ""),
    ("CCY", "통화", "VARCHAR2(3)", "NN", "", ""),
    ("SANCTION_AMT", "승인한도", "NUMBER(18,2)", "NN", "", ""),
    ("UTILIZED_AMT", "사용액", "NUMBER(18,2)", "NN", "= Σ TB_TF_LIMIT_UTIL(유효)", ""),
    ("AVAIL_AMT", "가용한도", "NUMBER(18,2)", "NN", "", ""),
    ("APPR_REF", "승인근거", "VARCHAR2(40)", "", "이사회·여신위원회 결의", ""),
    ("VALID_FROM", "시작", "DATE", "NN", "", ""),
    ("VALID_TO", "만료", "DATE", "NN", "", ""),
    ("LIMIT_STS", "상태", "VARCHAR2(6)", "NN", "ACTIVE/EXPIRE/SUSP", ""),
] + AUD)

table("TB_TF_LIMIT_UTIL", "한도 사용 원장", "무역금융(TF31)", "거래별 한도 차감·복원. 취소·만기회수 시 자동 복원.", [
    ("UTIL_ID", "ID", "NUMBER(12)", "PK", "", ""),
    ("LIMIT_ID", "한도ID", "VARCHAR2(16)", "FK,NN", "", ""),
    ("SRC_REF", "거래 Ref", "VARCHAR2(20)", "NN", "BILL_REF/LC_REF/BG_REF", ""),
    ("CCY", "통화", "VARCHAR2(3)", "NN", "", ""),
    ("AMT", "차감액", "NUMBER(18,2)", "NN", "한도통화 환산", ""),
    ("UTIL_STS", "상태", "VARCHAR2(4)", "NN", "HOLD(승인대기)/USED/RLSD(복원)", ""),
])

# ---------------------------------------------------------------------------
# 5. 미지급금(A/P, GL 21100000) - 할인 순지급액·타발 중계 공용
# ---------------------------------------------------------------------------
table("TB_AP_ITEM", "미지급금(A/P) 원장", "공통(TF21/RT22)", "GL 21100000. 할인 순지급액(TF)과 비계좌 타발 보류(RT22)를 하나의 A/P 원장으로 관리. 기존 이체화면에서 차감.", [
    ("AP_ID", "A/P ID", "VARCHAR2(20)", "PK", "", ""),
    ("AP_TYPE", "유형", "VARCHAR2(8)", "NN", "TF_PROC(할인 순지급) / INW_RELAY(타발 중계) / ESCROW_RL", ""),
    ("SRC_REF", "원천 Ref", "VARCHAR2(20)", "NN", "BILL_REF / SUSP_NO", ""),
    ("GL_CODE", "계정", "VARCHAR2(8)", "NN", "21100000", "GL"),
    ("CCY", "통화", "VARCHAR2(3)", "NN", "", ""),
    ("AP_AMT", "계상액", "NUMBER(18,2)", "NN", "", ""),
    ("AP_BAL", "잔액", "NUMBER(18,2)", "NN", "0이면 지급 완료", ""),
    ("PAY_METHOD", "지급방법", "VARCHAR2(6)", "NN", "NEFT/RTGS/INTERNAL/SWIFT", ""),
    ("BENE_NAME", "수취인", "VARCHAR2(140)", "NN", "", "SFMS"),
    ("BENE_IFSC", "IFSC", "VARCHAR2(11)", "", "", "SFMS"),
    ("BENE_ACCT", "수취계좌", "VARCHAR2(34)", "NN", "", "SFMS"),
    ("BOOK_DT", "계상일", "DATE", "NN", "", ""),
    ("AGE_ALERT_DT", "미지급 경고일", "DATE", "", "계상일 + 2영업일", ""),
    ("AP_STS", "상태", "VARCHAR2(6)", "NN", "OPEN/PART/PAID/HOLD", ""),
] + AUD, idx=["AP_STS,BOOK_DT"])

# ---------------------------------------------------------------------------
# 6. SFMS 전문 (M8)
# ---------------------------------------------------------------------------
table("TB_SFMS_MSG", "SFMS 전문", "SFMS(SF11~13)", "RTGS·NEFT(ISO 20022)·IFN 송수신 원문(변경불가, 10년). Velocis 미들웨어 경유.", [
    ("MSG_ID", "메시지ID", "VARCHAR2(20)", "PK", "미들웨어 부여", "SFMS"),
    ("CHANNEL", "채널", "VARCHAR2(4)", "NN", "RTGS/NEFT/IFN", ""),
    ("MSG_TYPE", "전문", "VARCHAR2(10)", "NN", "pacs.008/009/004/002, IFN 700~799, IFN 760/767/760COV", ""),
    ("DIRECTION", "방향", "VARCHAR2(3)", "NN", "OUT/IN", ""),
    ("E2E_ID", "End-to-End ID", "VARCHAR2(35)", "", "BILL_REF", "SFMS"),
    ("UTR", "UTR", "VARCHAR2(22)", "", "", "SFMS"),
    ("AMT", "금액(INR)", "NUMBER(18,2)", "", "", ""),
    ("AP_ID", "A/P", "VARCHAR2(20)", "FK", "", ""),
    ("IDEMPOTENCY_KEY", "멱등성 키", "VARCHAR2(40)", "", "BILL_REF + 시도번호 (R-CTL-002)", ""),
    ("ORIG_MSG_ID", "원전문", "VARCHAR2(20)", "", "재송신 연결", ""),
    ("RESEND_CNT", "재송신 횟수", "NUMBER(2)", "NN", "≤ 3 (거절)", ""),
    ("REASON_CD", "사유코드", "VARCHAR2(6)", "", "pacs.002/004 (AC03, AC04 ...)", ""),
    ("MSG_STS", "상태", "VARCHAR2(16)", "NN", "PENDING_APPROVAL/SENT/ACCEPTED/SETTLED/REJECTED/RETURNED/TIMEOUT/RECEIVED/MATCHED/UNMATCHED", ""),
    ("NEXT_BIZ_DAY_YN", "익영업일 인식", "VARCHAR2(1)", "", "GBS EOD 이후 수신", ""),
    ("RAW_PAYLOAD", "원문", "CLOB", "NN", "변경불가·접근로그", ""),
] + AUD, idx=["E2E_ID", "UTR", "MSG_STS"])

# ---------------------------------------------------------------------------
# 7. 신용장 (LC) - 신규
# ---------------------------------------------------------------------------
table("TB_LC_MASTER", "신용장", "LC(LC11)", "수입LC 개설(일람·유전스), 수출LC 통지·확인·매입, 내국LC(SFMS IFN 700). UCP 600 / ISBP / eUCP.", [
    ("LC_REF", "LC 참조", "VARCHAR2(16)", "PK", "", ""),
    ("LC_ROLE", "당행 역할", "VARCHAR2(8)", "NN", "ISSUE 개설 / ADVISE 통지 / CONFIRM 확인 / NEGO 매입", ""),
    ("LC_SCOPE", "구분", "VARCHAR2(6)", "NN", "IMPORT/EXPORT/INLAND", ""),
    ("MSG_CHANNEL", "전문망", "VARCHAR2(5)", "NN", "SWIFT(MT700) / SFMS(IFN 700)", ""),
    ("LC_NO", "LC 번호", "VARCHAR2(16)", "NN", "", "MT700 :20:"),
    ("CUST_NO", "개설의뢰인/수익자(당행 고객)", "VARCHAR2(8)", "FK,NN", "", ""),
    ("ISSUING_BIC", "개설은행", "VARCHAR2(11)", "NN", "통지·매입 시 평가 대상", ""),
    ("FORM", "형태", "VARCHAR2(12)", "NN", "IRREVOCABLE (UCP 600은 취소불능)", "MT700 :40A:"),
    ("AVAIL_BY", "사용방법", "VARCHAR2(10)", "NN", "PAYMENT/ACCEPTANCE/NEGOTIATION/DEF PAYMENT", "MT700 :41a:"),
    ("CCY", "통화", "VARCHAR2(3)", "NN", "", ""),
    ("LC_AMT", "금액", "NUMBER(18,2)", "NN", "", "MT700 :32B:"),
    ("TOL_PCT", "과부족(%)", "VARCHAR2(7)", "", "", "MT700 :39A:"),
    ("OS_AMT", "미결제 잔액", "NUMBER(18,2)", "NN", "한도 차감 기준", ""),
    ("USANCE_DAYS", "유전스 일수", "NUMBER(4)", "", "수입: 무역신용 판정(선적+6개월 초과 여부)", ""),
    ("LATEST_SHIP_DT", "최종선적일", "DATE", "", "", "MT700 :44C:"),
    ("EXPIRY_DT", "유효기일", "DATE", "NN", "", "MT700 :31D:"),
    ("EXPIRY_PLACE", "유효장소", "VARCHAR2(35)", "", "", ""),
    ("GOODS_DESC", "물품", "VARCHAR2(500)", "", "HS·이중용도 점검", "MT700 :45A:"),
    ("MARGIN_PCT", "증거금(%)", "NUMBER(5,2)", "", "여신 승인조건", ""),
    ("MARGIN_AMT", "증거금", "NUMBER(18,2)", "", "", ""),
    ("REIMB_BIC", "상환은행", "VARCHAR2(11)", "", "RA 발행 시 MT740 (본점 RA는 RA11)", "MT740"),
    ("CONFIRM_YN", "확인 여부", "VARCHAR2(1)", "", "", "MT700 :49:"),
    ("LIMIT_ID", "한도", "VARCHAR2(16)", "FK", "", ""),
    ("LC_STS", "상태", "VARCHAR2(6)", "NN", "DRAFT/ISSUED/AMEND/UTIL/EXPIRED/CLOSED/CANCEL", ""),
] + AUD, idx=["CUST_NO,LC_STS", "EXPIRY_DT"])

table("TB_LC_PRESENTATION", "LC 서류제시·심사", "LC(LC21)", "제시 접수 → 5은행영업일 내 심사 → 일치/하자 → 지급·인수·거절(734). 거절 이력은 개설은행 평가 제3축으로 자동 적재.", [
    ("PRES_ID", "제시ID", "VARCHAR2(16)", "PK", "", ""),
    ("LC_REF", "LC 참조", "VARCHAR2(16)", "FK,NN", "", ""),
    ("PRESENT_DT", "제시일", "DATE", "NN", "", ""),
    ("EXAM_DUE_DT", "심사기한", "DATE", "NN", "제시 다음날부터 5은행영업일 (UCP 600 Art.14(b))", ""),
    ("DOC_AMT", "청구금액", "NUMBER(18,2)", "NN", "≤ OS_AMT(허용치 포함)", ""),
    ("DISCREPANCIES", "하자", "VARCHAR2(1000)", "", "", "MT750/734"),
    ("RESULT", "결과", "VARCHAR2(8)", "NN", "COMPLY/DISCREP/WAIVED/REFUSED", ""),
    ("REFUSAL_MSG", "거절통지", "VARCHAR2(10)", "", "MT734 / IFN 734 (단일 통지, Art.16(c))", ""),
    ("MATURITY_DT", "만기(인수)", "DATE", "", "", "MT732/752"),
    ("SETTLE_DT", "결제일", "DATE", "", "", ""),
    ("BILL_REF", "매입·할인 연결", "VARCHAR2(20)", "FK", "NEGOTIATION → TB_TF_BILL", ""),
] + AUD, idx=["LC_REF", "EXAM_DUE_DT"])

# ---------------------------------------------------------------------------
# 8. 보증 (BG / SBLC) - 신규
# ---------------------------------------------------------------------------
table("TB_BG_MASTER", "보증서(BG/SBLC)", "BG(BG11)", "국내 BG(SFMS IFN 760 + IFN 760COV 필수), 해외 BG(MT760 / 역보증 기반 발행), SBLC. URDG 758 / ISP98.", [
    ("BG_REF", "보증 참조", "VARCHAR2(16)", "PK", "", ""),
    ("BG_SCOPE", "구분", "VARCHAR2(8)", "NN", "INLAND / OUTWARD(해외수익자) / INWARD_CG(역보증 수령 후 국내 발행)", ""),
    ("BG_TYPE", "보증유형", "VARCHAR2(8)", "NN", "FIN 금융 / PERF 이행 / BID 입찰 / ADVPMT 선급금 / SBLC / CUSTOMS", ""),
    ("RULES", "준거규칙", "VARCHAR2(8)", "NN", "URDG758/ISP98/LOCAL(인도계약법)", ""),
    ("CUST_NO", "보증의뢰인", "VARCHAR2(8)", "FK,NN", "", ""),
    ("BENE_NAME", "수익자", "VARCHAR2(140)", "NN", "", ""),
    ("BENE_BANK_IFSC", "수익자 은행 IFSC", "VARCHAR2(11)", "", "국내: IFN 760COV 수신은행", "SFMS"),
    ("BENE_BANK_BIC", "수익자 은행 BIC", "VARCHAR2(11)", "", "해외: MT760", ""),
    ("CCY", "통화", "VARCHAR2(3)", "NN", "", ""),
    ("BG_AMT", "보증금액", "NUMBER(18,2)", "NN", "", ""),
    ("ISSUE_DT", "발행일", "DATE", "NN", "", ""),
    ("EXPIRY_DT", "만료일", "DATE", "NN", "", ""),
    ("CLAIM_EXPIRY_DT", "청구만료일", "DATE", "NN", "인도계약법 s.28 예외: 만료 후 청구기간 1년 이상 (BG_CLAIM_MIN_MONTHS)", ""),
    ("TENOR_MONTHS", "기간(개월)", "NUMBER(4)", "", "", ""),
    ("E_BG_YN", "전자보증(e-BG)", "VARCHAR2(1)", "NN", "NeSL DDE 발행 시 Y", "NeSL"),
    ("NESL_REF", "NeSL 참조", "VARCHAR2(30)", "", "", "NeSL"),
    ("STAMP_DUTY_AMT", "인지세", "NUMBER(18,2)", "", "주별 인지세(UP)·e-stamp 번호", ""),
    ("SFMS_760COV_ID", "IFN 760COV", "VARCHAR2(20)", "FK", "국내 BG 진위확인 전문 (수익자 은행 앞)", "SFMS"),
    ("COUNTER_GTEE_REF", "역보증 참조", "VARCHAR2(30)", "", "INWARD_CG: 해외은행(본점 포함) MT760 역보증", ""),
    ("NR_PARTY_YN", "비거주자 관련", "VARCHAR2(1)", "NN", "FEMA 보증 규정 검증 대상", ""),
    ("FEMA_BASIS", "FEMA 근거", "VARCHAR2(60)", "", "수입 선급금 보증/역보증 기반/비거주자 앞 등", ""),
    ("MARGIN_AMT", "증거금", "NUMBER(18,2)", "", "", ""),
    ("COMMISSION_RATE", "보증료율(연)", "NUMBER(7,4)", "", "기간 비례 + GST", "GST"),
    ("LIMIT_ID", "한도", "VARCHAR2(16)", "FK", "", ""),
    ("BG_STS", "상태", "VARCHAR2(6)", "NN", "DRAFT/ISSUED/AMEND/INVOKE/PAID/EXPIRED/CLAIMX(청구기간 종료)/CANCEL", ""),
] + AUD, idx=["CUST_NO,BG_STS", "CLAIM_EXPIRY_DT"])

table("TB_BG_CLAIM", "보증 청구(Invocation)", "BG(BG21)", "청구 접수·요건 심사·지급·대지급(Devolvement) → 여신 전환.", [
    ("CLAIM_ID", "청구ID", "VARCHAR2(16)", "PK", "", ""),
    ("BG_REF", "보증 참조", "VARCHAR2(16)", "FK,NN", "", ""),
    ("CLAIM_DT", "청구일", "DATE", "NN", "≤ CLAIM_EXPIRY_DT", ""),
    ("CLAIM_AMT", "청구금액", "NUMBER(18,2)", "NN", "≤ 잔여 보증금액", ""),
    ("COMPLIANT_YN", "청구요건 일치", "VARCHAR2(1)", "NN", "URDG 758 Art.15·20 / 보증서 문언", ""),
    ("PAY_DUE_DT", "지급기한", "DATE", "", "URDG: 제시 다음날부터 5영업일 심사", ""),
    ("PAID_DT", "지급일", "DATE", "", "", ""),
    ("DEVOLVE_LOAN_REF", "대지급 여신", "VARCHAR2(20)", "", "의뢰인 미상환 시 여신 전환·자산건전성 대상", ""),
    ("CLAIM_STS", "상태", "VARCHAR2(6)", "NN", "RCVD/EXAM/PAID/REJECT/DEVOLV", ""),
] + AUD)

# ---------------------------------------------------------------------------
# 9. 자산건전성·결정
# ---------------------------------------------------------------------------
table("TB_TF_ASSET_CLASS", "자산건전성 일별", "무역금융(EOD)", "연체일수(DPD) → SMA-0/1/2·NPA, 충당금, CRILC 보고 대상.", [
    ("BILL_REF", "거래 Ref", "VARCHAR2(20)", "PK,FK", "BILL·LC 인수·BG 대지급", ""),
    ("AS_OF_DT", "기준일", "DATE", "PK", "", ""),
    ("DPD", "연체일수", "NUMBER(5)", "NN", "", ""),
    ("ASSET_CLASS", "분류", "VARCHAR2(6)", "NN", "STD/SMA0/SMA1/SMA2/NPA", "CRILC"),
    ("BORROWER_EXPOSURE_INR", "차주 총익스포저", "NUMBER(18,2)", "", "CRILC 기준(₹5 crore 이상) 판정", "CRILC"),
    ("CRILC_RPT_YN", "CRILC 보고 대상", "VARCHAR2(1)", "", "", "CRILC"),
    ("PROVISION_AMT", "충당금", "NUMBER(18,2)", "", "", "GL"),
    ("INCOME_STOP_YN", "이자인식 중단", "VARCHAR2(1)", "", "NPA", ""),
])

table("TB_TF_CRYSTAL", "외화 어음 결정(Crystallisation)", "무역금융(EOD)", "만기 후 미회수 외화 수출어음을 FEDAI 규칙에 따라 INR 채권으로 전환.", [
    ("BILL_REF", "어음참조번호", "VARCHAR2(20)", "PK,FK", "", ""),
    ("CRYSTAL_DT", "결정일", "DATE", "NN", "만기 + 30일 (FEDAI)", ""),
    ("TT_SELL_RATE", "전신환매도율", "NUMBER(15,6)", "NN", "", ""),
    ("INR_AMT", "결정 INR", "NUMBER(18,2)", "NN", "", "GL"),
    ("FX_POS_GL", "환포지션", "VARCHAR2(8)", "NN", "28900000", "GL"),
])

# ---------------------------------------------------------------------------
# 코드 / 파라미터 / 규칙 / 기한 / 화면 / IF / GL
# ---------------------------------------------------------------------------
CODES = [
    ("BILL_CATEGORY", "DISCOUNT", "어음할인 (Bill Discount)", "한도 차감, 할인계산", "N"),
    ("BILL_CATEGORY", "PURCHASE", "어음매입 (Bill Purchase)", "", "N"),
    ("BILL_CATEGORY", "NEGOTIATION", "LC 매입 (Negotiation under LC)", "신용 프리미엄 = 개설은행 등급", "N"),
    ("BILL_CATEGORY", "COLLECTION_DA", "D/A 추심", "URC 522, 한도 미차감(인수 후 할인 시 차감)", "N"),
    ("BILL_CATEGORY", "COLLECTION_DP", "D/P 추심", "URC 522", "N"),
    ("BILL_CATEGORY", "OPEN_ACCOUNT", "오픈어카운트", "자금 미제공 시 한도 미차감", "N"),
    ("BILL_CATEGORY", "TREDS", "TReDS", "MSME PSL", "N"),
    ("ISSUER_GRADE", "S", "90~100 · 0.15% · 기본자본 10%", "", "N"),
    ("ISSUER_GRADE", "A", "75~89 · 0.35% · 8%", "", "N"),
    ("ISSUER_GRADE", "B", "60~74 · 0.60% · 5%", "", "N"),
    ("ISSUER_GRADE", "C", "45~59 · 1.00% · 3% · 여신 사전협의", "", "N"),
    ("ISSUER_GRADE", "D", "30~44 · 1.75% · 1% · 확인LC/ECGC 조건", "", "N"),
    ("ISSUER_GRADE", "E", "<30 · 취급불가", "", "N"),
    ("POLICY_REASON", "D1", "인하 - 경쟁 대응", "", "N"),
    ("POLICY_REASON", "D2", "인하 - 거래관계 고려", "", "N"),
    ("POLICY_REASON", "D3", "인하 - 신규 거래처 유치", "", "N"),
    ("POLICY_REASON", "D4", "인하 - 대형거래 우대", "", "N"),
    ("POLICY_REASON", "D5", "인하 - 본점 전략방침", "", "N"),
    ("POLICY_REASON", "U1", "인상 - 신용우려 가산", "", "N"),
    ("POLICY_REASON", "U2", "인상 - 집중도 관리", "", "N"),
    ("POLICY_REASON", "U3", "인상 - 운영부담 반영", "", "N"),
    ("POLICY_REASON", "U4", "인상 - 시장금리 급변", "", "N"),
    ("ASSET_CLASS", "SMA0", "1~30일", "", "N"),
    ("ASSET_CLASS", "SMA1", "31~60일", "", "N"),
    ("ASSET_CLASS", "SMA2", "61~90일", "", "N"),
    ("ASSET_CLASS", "NPA", "90일 초과", "이자인식 중단·충당금", "N"),
    ("BG_TYPE", "FIN", "금융보증", "", "N"),
    ("BG_TYPE", "PERF", "이행보증", "", "N"),
    ("BG_TYPE", "BID", "입찰보증", "", "N"),
    ("BG_TYPE", "ADVPMT", "선급금 환급보증", "수입 선수금 BG(R-IMP-06)와 연계", "N"),
    ("BG_TYPE", "SBLC", "보증신용장", "ISP98/UCP 600", "N"),
    ("LC_ROLE", "ISSUE", "개설 (수입·내국)", "", "N"),
    ("LC_ROLE", "ADVISE", "통지 (수출)", "", "N"),
    ("LC_ROLE", "CONFIRM", "확인", "개설은행 평가·한도", "N"),
    ("LC_ROLE", "NEGO", "매입 (수출)", "TB_TF_BILL NEGOTIATION", "N"),
]

FLOW = [
    ("TF11 Trade Finance Registration", "거래유형 선택 → 공통 패널(기본·당사자·서류·컴플라이언스·한도) + 동적 패널(할인/추심/OA/TReDS) → 임시저장 → 승인요청", "TB_TF_BILL + 하위"),
    ("TF11 할인계산", "기준금리+기간+신용(개설은행 등급)+국가+정책 → 적용금리 하한·상한 → 할인료·수수료·GST·전신료 → 순지급액", "TB_TF_DISCOUNT"),
    ("TF31 한도", "전체·상품·단일거래상대방·국가·일중 동시 검증, 80% 경고, 100% 차단, 취소·회수 시 복원", "TB_TF_LIMIT(_UTIL)"),
    ("TF91 Checker", "Maker≠Checker, ALCO 기준 초과 Authorizer, 정책 프리미엄 구간 결재선", ""),
    ("A/P → SF11", "승인 시 순지급액 A/P(21100000) 계상 → SFMS 메시지 실행(RTGS/NEFT) → pacs.002/004 → A/P 차감·복원", "TB_AP_ITEM, TB_SFMS_MSG"),
    ("EDPMS·IDPMS", "등록·할인·회수 시점 자동 연계(기존 IF-EDP/IDP), 실패 시 익영업일 재전송", "TB_IF_*"),
    ("TF81 EOD", "① 이자발생 ② 만기도래 ③ 연체전환 ④ SMA·NPA ⑤ 결정 ⑥ EDPMS·IDPMS ⑦ TReDS 대사 ⑧ 추심 독촉 ⑨ 마감요약·잠금 (Dry Run)", "TB_TF_ASSET_CLASS, TB_TF_CRYSTAL"),
]

# ---------------------------------------------------------------------------
# 파라미터 (param_id, 명칭, 값, 단위, 업무, 근거, VERIFY, 비고)
# ---------------------------------------------------------------------------
PARAMS = [
    ("TF_TENOR_SPREAD", "기간 스프레드", "≤90일 0.25 / 91~180일 0.50 / >180일 0.85", "%", "무역금융", "지점 ALCO (요청서 4.2)", "N", "반기 갱신"),
    ("TF_POLICY_TIERS", "정책 프리미엄 결재 구간", "T1 ±0.10 / T2 -0.25~+0.30 / T3 -0.40~+0.60 / T4 -0.50~+1.00", "%", "무역금융", "지점 결재규정 (요청서 4.2.2)", "N", "범위 초과 시스템 차단"),
    ("TF_AUTH_THRESHOLD", "Authorizer 추가승인 기준", "ALCO 결정값", "INR", "무역금융", "지점 ALCO", "N", "건별 INR 환산"),
    ("TF_LIMIT_WARN_PCT", "한도 소진 경고", "80", "%", "무역금융", "지점 여신규정", "N", "100% 신규 차단"),
    ("TF_SINGLE_CPTY_PCT", "단일거래상대방 한도", "10", "% (기본자본)", "무역금융", "지점 기준 (RBI Large Exposures 규정 범위 내)", "Y", "RBI 대규모 익스포저 한도와 정합성 확인"),
    ("TF_AP_AGE_DAYS", "A/P 미지급 경고", "2", "영업일", "무역금융", "요청서 4.4", "N", ""),
    ("TF_CRYSTAL_DAYS", "외화 어음 결정 시점", "30", "일(만기 후)", "무역금융", "FEDAI Rules", "N", "전신환매도율"),
    ("TF_EXP_CREDIT_MAX_DAYS", "수출신용(선적 전·후) 통상 기간", "365", "일", "무역금융", "RBI (Commercial Banks – Credit Facilities) Directions, 2025 Ch.XV", "Y", "2025.11 무역구제(450일)는 2026.6.30 실행분까지 - 이후 소멸 여부 확인"),
    ("TF_EPM_SUBV_CAP_INR", "EPM Niryat Protsahan 이자보전 한도", "5000000", "INR/년·MSME 수출자", "무역금융", "RBI/2025-26/195 (2026.1.19), DGFT TN 20/2025-26", "Y", "IES는 2024.12.31 종료. 2026.4.1부터 EXIM Bank 청구"),
    ("TF_IMP_USANCE_POLICY", "수입 유전스 내부 상한", "일반 180일 / 자본재 360일", "일", "무역금융", "지점 내부정책 (RBI 고정기한 아님)", "N", "초과 시 무역신용 프레임워크 검토"),
    ("TC_MAX_USD", "무역신용 건당 한도", "50000000", "USD", "무역금융", "FEM (Borrowing and Lending) Regs / MD - Trade Credit (2026.2 개정)", "Y", "석유·항공·해운 USD 150M"),
    ("TC_TENOR", "무역신용 기간", "자본재 3년 / 비자본재 1년 또는 영업주기 중 짧은 기간", "-", "무역금융", "MD - Trade Credit", "Y", ""),
    ("BG_CLAIM_MIN_MONTHS", "보증 청구기간(만료 후)", "12", "개월", "보증", "Indian Contract Act s.28 예외 (은행 보증 청구기간 1년 이상)", "Y", ""),
    ("BG_FEMA_RPT_DAYS", "FEMA 보증 분기보고 기한", "30", "일(분기말 후)", "보증", "FEM (Guarantees) Regulations, 2026 (FEMA 8(R)/2026-RB, 2026.1.10 시행)", "Y", "지연보고 수수료(LSF)"),
    ("BG_SFMS_COV_REQUIRED", "국내 BG IFN 760COV 송신", "Y", "Y/N", "보증", "IBA/IFTAS 업계 관행 (RBI 의무 규정 미확인)", "Y", "지점 내부정책으로 운영"),
    ("PCE_MAX_PCT", "부분신용보강 한도", "50", "%", "보증", "RBI (Non-Fund Based Credit Facilities) Directions, 2025", "Y", ""),
    ("LC_EXAM_DAYS", "LC 서류심사 기간", "5", "은행영업일", "LC", "ICC UCP 600 Art.14(b)", "N", "제시 다음날부터"),
    ("URDG_EXAM_DAYS", "보증 청구 심사기간", "5", "영업일", "보증", "ICC URDG 758 Art.20", "N", ""),
    ("SFMS_RESEND_MAX", "SFMS 거절 재송신 한도", "3", "회", "SFMS", "SFMS IF 요건 7.3", "N", ""),
    ("SFMS_TIMEOUT_MIN", "SFMS 응답 시한", "30", "분", "SFMS", "SFMS 화면설계(R-EXC-002)", "N", "초과 시 재송신 잠금"),
]

# (rule_id, 업무, 화면, 규칙명, 조건/로직, 처리, 근거, 파라미터, VERIFY)
RULES = [
    ("R-TF-01", "무역금융", "TF11", "동적 패널", "거래유형별 표시 패널 고정(공통 5 + 할인/추심/OA/TReDS). 미표시 패널 필드는 저장하지 않음", "AUTO", "화면설계서 v6", "", "N"),
    ("R-TF-02", "무역금융", "TF11", "5단계 금리", "적용금리 = 기준 + 기간 + 신용 + 국가 + 정책. 각 요소와 출처를 TB_TF_DISCOUNT에 저장", "AUTO", "요청서 4.2 / RBI (Commercial Banks – Interest Rates on Advances) Directions, 2025", "TF_TENOR_SPREAD", "N"),
    ("R-TF-03", "무역금융", "TF11", "신용 프리미엄 대상", "LC 매입(NEGOTIATION)은 개설은행 등급, 그 외는 지급인·수출자 등급. E등급 개설은행 신규 취급 차단, D등급은 확인LC·ECGC 조건", "ERR", "요청서 4.2.1 / UCP 600", "", "N"),
    ("R-TF-04", "무역금융", "TF11/TF91", "정책 프리미엄 결재", "조정 폭에 따라 T1~T4 결재선, 범위 초과·하한(기준금리 미만)·이사회 상한 초과 차단. 사유코드 필수, 이력 10년", "ERR", "요청서 4.2.2~4.2.3", "TF_POLICY_TIERS", "N"),
    ("R-TF-05", "무역금융", "TF11/TF31", "5층 한도", "전체·상품·단일거래상대방·국가·일중 동시 검증. 80% 경고, 100% 차단, 국가한도는 경고 후 사후보고, 취소·회수 시 자동 복원", "ERR", "요청서 4.3", "TF_LIMIT_WARN_PCT", "N"),
    ("R-TF-06", "무역금융", "TF11", "무역신용 판정", "수입 결제가 계약상 기한을 넘어 해외 공급자·은행 신용을 수반하면 무역신용 프레임워크(건당 USD 50M, 자본재 3년·비자본재 1년/영업주기) 검증·보고. 요청서의 '선적 후 6개월 초과' 기준은 2026 규정 하 재확인 필요", "WARN", "FEM (Borrowing and Lending) Regs, MD - Trade Credit / FEM (Export and Import) Regs 2026", "TC_MAX_USD", "Y"),
    ("R-TF-07", "무역금융", "TF11", "수출 회수기한", "물품: 선적일 + 9개월, 서비스: 인보이스일 + 9개월, INR 인보이스·결제분 12개월. 2026.10.1 이전 선적분은 당시 파라미터 이력 적용", "ERR", "FEM (Export and Import of Goods and Services) Regulations, 2026 as amended by FEMA 23(R)/(1)/2026-RB", "EXP_REALIZE_MONTHS", "Y"),
    ("R-TF-08", "무역금융", "TF11", "수입 유전스", "RBI 고정 상한 없음(계약 기한 기준, AD 연장). 지점 내부 상한 일반 180일·자본재 360일 초과 시 무역신용 검토·여신 승인", "WARN", "FEM (Export and Import) Regs 2026 / 지점 정책", "TF_IMP_USANCE_POLICY", "N"),
    ("R-TF-09", "무역금융", "TF11", "A/P 경유 지급", "승인 시 순지급액을 A/P(21100000)에 계상, 지급은 SF11/이체화면에서 차감. 2영업일 미지급 경고", "AUTO", "요청서 4.4", "TF_AP_AGE_DAYS", "N"),
    ("R-TF-10", "무역금융", "TF81", "자산건전성", "DPD 1~30 SMA-0, 31~60 SMA-1, 61~90 SMA-2, 90 초과 NPA(이자인식 중단·충당금). 자동 판정(수기 변경 불가). 차주 익스포저 ₹5 crore 이상 CRILC 보고", "AUTO", "RBI (Commercial Banks – IRACP) Directions, 2025 / CRILC", "", "Y"),
    ("R-TF-11", "무역금융", "TF81", "외화 결정", "만기 후 30일 미회수 외화 수출어음을 전신환매도율로 INR 결정(환포지션 28900000)", "AUTO", "FEDAI Rules", "TF_CRYSTAL_DAYS", "N"),
    ("R-TF-12", "무역금융", "TF11", "금리 하한·상한", "적용금리 ≥ 기준금리(조달원가), ≤ 이사회 승인 상한. 수출선수금·수입 지연지급 이자는 무역신용 all-in-cost 상한 이내", "ERR", "RBI Interest Rates on Advances Directions 2025 / FEM Export-Import Directions 2026", "", "Y"),
    ("R-TF-13", "무역금융", "TF11", "EPM 이자보전", "MSME 수출자 선적 전·후 수출신용은 EPM Niryat Protsahan 이자보전 대상 여부 판정(연 ₹50 lakh 한도), 청구는 EXIM Bank", "AUTO", "RBI/2025-26/195", "TF_EPM_SUBV_CAP_INR", "Y"),
    ("R-TF-14", "무역금융", "TF11", "TReDS", "매입자 인수 후 입찰, 무소구(without recourse), 보험료는 매출자 부담 불가, 재양도(re-assignment) 이력 관리, MSME PSL 자동 분류", "ERR", "RBI (TReDS) Directions, 2026 (2026.6.23)", "", "Y"),
    ("R-TF-15", "무역금융", "TF11", "결정 전 서류", "선적서류 필수 미접수(보험증권 등) 시 승인요청 차단, OA는 선택", "ERR", "화면설계서 v6", "", "N"),
    ("R-TF-16", "무역금융", "TF11", "EDPMS·IDPMS 연계", "등록·할인·회수 시점 자동 연계(기존 IF-EDP/IF-IDP), 실패 시 익영업일 재전송·알림. ₹10 lakh 이하 감액·미회수는 수출자 신고로 종결 가능(지점 정책)", "AUTO", "FEM Export-Import Directions 2026", "EXP_DECL_CLOSE_INR", "Y"),
    ("R-LC-01", "LC", "LC11", "개설 요건", "개설의뢰인 여신한도·증거금, 물품·HS·이중용도 점검, UCP 600 준거, 유전스가 무역신용에 해당하면 R-TF-06", "ERR", "RBI NFB Credit Facilities Directions 2025 / UCP 600", "", "N"),
    ("R-LC-02", "LC", "LC11", "비자금 여신 대상", "LC·보증 등 비자금 여신은 원칙적으로 당행 자금여신 거래 고객에게만 제공(예외: 역보증·전액담보·자금여신은행 NOC)", "ERR", "RBI (Non-Fund Based Credit Facilities) Directions, 2025 (2026.4.1 시행)", "", "Y"),
    ("R-LC-03", "LC", "LC21", "서류심사 기한", "제시 다음날부터 5은행영업일 내 일치 여부 결정. 하자 시 단일 거절통지(MT734/IFN 734)에 모든 하자 기재", "ALERT", "UCP 600 Art.14(b), 16", "LC_EXAM_DAYS", "N"),
    ("R-LC-04", "LC", "LC21", "거절 이력 연동", "수령한 734 거절은 개설은행 평가 제3축(최근 24개월 승인불가)에 자동 적재 → 재평가 알림", "AUTO", "요청서 4.2.1 / SFMS IF 요건 7.3", "", "N"),
    ("R-LC-05", "LC", "LC11", "전문망", "해외 LC는 SWIFT MT700/707/734/740/742/756(Cat7 MT 유지), 국내 LC는 SFMS IFN 700 계열", "AUTO", "SWIFT / SFMS", "", "N"),
    ("R-BG-01", "보증", "BG11", "발행 요건", "의뢰인 한도·증거금, 보증유형별 문언 승인, 금융보증이 타 규제기관 자금여신을 가능케 하는 경우 차단(무역거래 예외)", "ERR", "RBI (Non-Fund Based Credit Facilities) Directions, 2025", "", "Y"),
    ("R-BG-02", "보증", "BG11", "청구기간", "만료일과 별도로 청구만료일(만료 후 최소 12개월) 지정, 청구만료일 경과 시 자동 해지·한도 복원", "AUTO", "Indian Contract Act s.28", "BG_CLAIM_MIN_MONTHS", "Y"),
    ("R-BG-03", "보증", "BG11", "국내 BG 진위확인", "국내 BG 발행 시 수익자 은행 앞 SFMS IFN 760COV 송신(e-BG 발행 시 NeSL 참조 저장)", "AUTO", "IBA/IFTAS 관행 (지점 정책)", "BG_SFMS_COV_REQUIRED", "Y"),
    ("R-BG-04", "보증", "BG11", "FEMA 보증", "비거주자 관련 보증(해외 수익자, 역보증 기반 발행 등)은 기초거래의 FEMA 적합성 확인 후 발행, 분기말 후 30일 내 AD 보고", "ERR", "FEM (Guarantees) Regulations, 2026", "BG_FEMA_RPT_DAYS", "Y"),
    ("R-BG-05", "보증", "BG21", "청구 처리", "청구요건 일치 심사(5영업일), 법원 금지명령이 없으면 지급, 의뢰인 미상환분은 대지급 여신 전환·자산건전성 편입", "ERR", "URDG 758 / RBI NFB Directions 2025", "URDG_EXAM_DAYS", "N"),
    ("R-VAL-001", "SFMS", "SF11", "IFSC 형식", "수취 IFSC 11자리 형식 오류 → 송신 차단", "ERR", "SFMS IF 요건 7.2", "", "N"),
    ("R-VAL-002", "SFMS", "SF11", "IFSC 마스터", "IFSC 마스터 미존재 → 차단", "ERR", "SFMS IF 요건 7.2", "", "N"),
    ("R-VAL-003", "SFMS", "SF11", "A/P 잔액", "송금액 > A/P 잔액 → 차단", "ERR", "SFMS IF 요건 7.2", "", "N"),
    ("R-VAL-004", "SFMS", "SF11", "문자셋", "수취인명 비허용 문자 → 치환·경고", "WARN", "SFMS IF 요건 7.2", "", "N"),
    ("R-RTE-002", "SFMS", "SF11", "RTGS 최소금액", "RTGS 선택 & 금액 < ₹2 lakh → 차단 (권장 채널과 다르면 사유 필수)", "ERR", "RBI RTGS", "RTGS_MIN_INR", "N"),
    ("R-CTL-001", "SFMS", "SF11", "Maker-Checker", "동일인 승인 차단", "ERR", "SFMS IF 요건 7.2", "", "N"),
    ("R-CTL-002", "SFMS", "SF11", "중복 송신", "멱등성 키(BILL_REF+시도번호) 존재 → 거부", "ERR", "SFMS IF 요건 6.1", "", "N"),
    ("R-CTL-003", "SFMS", "SF11", "AML", "미들웨어 AML hit → 보류·준법 통보", "ERR", "SFMS IF 요건 7.2", "", "N"),
    ("R-EXC-001", "SFMS", "SF12", "반환", "pacs.004 수신 → A/P 자동 복원, 원전문ID 연결 재송신", "AUTO", "SFMS IF 요건 7.3", "", "N"),
    ("R-EXC-002", "SFMS", "SF12", "시한초과", "응답 30분 초과 → 재송신 잠금, 미들웨어 상태조회로 미접수 확인 시에만 재송신(동일 E2E, 신규 MsgId)", "ERR", "SFMS 화면설계", "SFMS_TIMEOUT_MIN", "N"),
    ("R-MCH-002", "SFMS", "SF12", "입금 매칭", "EndToEndId+금액 일치 자동 매칭, 불일치는 수작업 대기열(차액 처리방법 기록)", "AUTO", "SFMS IF 요건 4.3", "", "N"),
    ("R-IFN-001", "SFMS", "SF12", "IFN 734", "IFN 734 수신 → 개설은행 거절 기록·재평가 요청", "AUTO", "SFMS IF 요건 7.3", "", "N"),
]

DEADLINES = [
    ("TF_DUE", "무역금융", "만기일", "만기일(영업일 조정)", "D-7/D-1", "회수", "연체 전환 → DPD", "-", "", "N"),
    ("TF_CRYSTAL", "무역금융", "만기일", "만기 + 30일", "D-5", "회수", "INR 결정 (FEDAI)", "-", "TF_CRYSTAL_DAYS", "N"),
    ("TF_AP_AGE", "무역금융", "A/P 계상일", "+2영업일", "D-day", "A/P 잔액 0", "담당자 경고", "-", "TF_AP_AGE_DAYS", "N"),
    ("LC_EXAM", "LC", "서류 제시일", "다음날부터 5은행영업일", "D-2/D-1", "일치·하자 결정", "UCP 600상 거절권 상실 위험", "-", "LC_EXAM_DAYS", "N"),
    ("LC_EXPIRY", "LC", "개설일", "유효기일", "D-15", "미사용 잔액 해지", "한도 복원", "-", "", "N"),
    ("BG_EXPIRY", "보증", "발행일", "만료일 → 청구만료일(+12개월 이상)", "D-30", "청구만료 도래", "자동 해지·한도 복원", "-", "BG_CLAIM_MIN_MONTHS", "Y"),
    ("BG_FEMA_RPT", "보증", "분기말", "+30일", "D-10", "분기 보고 제출", "LSF 대상", "RBI", "BG_FEMA_RPT_DAYS", "Y"),
    ("EDF_SVC", "수출(서비스)", "인보이스 월말", "+30일", "D-7", "EDF 제출", "보고 지연", "EDPMS", "EXP_SVC_EDF_DAYS", "Y"),
]

SCREENS = [
    ("TF11", "Trade Finance Registration", "기업/개인사업자", "신규", "할인·매입·LC매입·D/A·D/P·OA·TReDS 통합 등록, 동적 패널, 5단계 금리, 한도 헤더, A/P 지급정보", "TB_TF_BILL 외", "tf11"),
    ("TF31", "Limit Management", "기업", "신규", "5층 한도 조회·사용내역·변경이력, 80%/100% 경고·차단", "TB_TF_LIMIT(_UTIL)", "tf31"),
    ("TF41", "Issuing Bank Evaluation", "내부", "신규", "개설은행 3축 평가·등급·프리미엄·단일은행 상한, 734 거절 이력 자동 반영", "TB_TF_ISSUER_EVAL", "tf41"),
    ("TF81", "Trade Finance EOD", "내부", "신규", "9개 배치(이자·만기·연체·SMA/NPA·결정·EDPMS/IDPMS·TReDS·독촉·마감), Dry Run", "TB_TF_ASSET_CLASS, TB_TF_CRYSTAL", "tf81"),
    ("LC11", "Letter of Credit", "기업", "신규", "수입LC 개설·수출LC 통지/확인/매입·내국LC, 서류제시 5영업일 심사, 734 연동", "TB_LC_MASTER, TB_LC_PRESENTATION", "lc11"),
    ("BG11", "Bank Guarantee", "기업", "신규", "BG/SBLC 발행·청구기간·IFN 760COV·e-BG·FEMA 보증 보고·청구(Invocation)", "TB_BG_MASTER, TB_BG_CLAIM", "bg11"),
    ("SF11", "SFMS Message Execution", "내부", "신규", "A/P 지급 전문(pacs.008) 작성·채널·규칙검증·승인, 거절·반환·시한초과·미매칭 재처리, 조회", "TB_SFMS_MSG, TB_AP_ITEM", "sf11"),
]

INTERFACES = [
    ("IF-SFMS-01", "SFMS via Velocis", "OUT", "REST /v1/payments/outward (멱등성 키)", "실시간", "A/P 지급 pacs.008 (RTGS/NEFT)", "TB_SFMS_MSG", "N"),
    ("IF-SFMS-02", "SFMS via Velocis", "IN", "콜백 /v1/payments/inward·return", "실시간", "입금 pacs.008, 반환 pacs.004, 상태 pacs.002", "TB_SFMS_MSG", "N"),
    ("IF-SFMS-03", "SFMS via Velocis", "BOTH", "/v1/trade/ifn (+inbound)", "실시간", "IFN 700/707/730/734/742/750/754/756/799, IFN 760/767/760COV(보증)", "TB_SFMS_MSG, TB_LC_*, TB_BG_*", "N"),
    ("IF-SFMS-04", "SFMS via Velocis", "IN", "/v1/recon/files, SOD/EOD", "1~2시간", "대사 파일·영업일 확정", "TB_SFMS_MSG", "N"),
    ("IF-TREDS-01", "TReDS (RXIL·M1xchange·Invoicemart)", "BOTH", "플랫폼 API", "실시간/일", "FU 수신·입찰·낙찰·정산 대사", "TB_TF_TREDS", "N"),
    ("IF-NESL-01", "NeSL (e-BG)", "BOTH", "NeSL API", "건별", "e-BG 발행·변경·청구", "TB_BG_MASTER", "N"),
    ("IF-CRILC-01", "RBI CRILC", "OUT", "XBRL/포털", "월·주", "₹5 crore 이상 차주 SMA·NPA", "TB_TF_ASSET_CLASS", "Y"),
    ("IF-GTEE-01", "RBI (FEMA 보증 보고)", "OUT", "RBI 지정 포털", "분기", "비거주자 관련 보증 내역", "TB_BG_MASTER", "Y"),
    ("IF-EXIM-01", "EXIM Bank (EPM 이자보전)", "OUT", "포털", "분기", "MSME 수출신용 이자보전 청구", "TB_TF_BILL", "Y"),
]

GL = [
    ("할인 승인", "10610009 할인어음 (액면)", "21100000 미지급금(A/P) 순지급 / 19900000 미실현할인 / 41040000 수수료 / 20800000 GST예수금", "순지급액 A/P 계상"),
    ("A/P 지급", "21100000 미지급금", "10211000 RBI 당좌예금 (RTGS·NEFT)", "SF11 / 이체화면"),
    ("일일 이자발생", "19900000 미실현할인", "41020012 할인수익", "EOD ①, NPA는 중단"),
    ("만기 회수", "10211000 예치환은행 / Nostro", "10610009 할인어음", ""),
    ("외화 결정", "INR 채권 (결정)", "10610009 할인어음 / 28900000 환포지션", "EOD ⑤"),
    ("NPA 충당금", "52010000 대손충당금전입", "10619000 할인어음 충당금", "EOD ④"),
    ("LC 개설 (난외)", "LC 개설 대충계정(Contra)", "LC 지급보증 대충", "증거금 별도 수입"),
    ("LC 결제 (지급)", "수입자 계좌 / 대지급", "Nostro / 상환은행", "미상환 시 여신 전환"),
    ("BG 발행 (난외)", "보증 대충계정(Contra)", "보증채무 대충", "보증료 수익 기간안분 + GST"),
    ("BG 대지급", "BG 대지급금 (여신)", "수익자 지급 Nostro/RTGS", "R-BG-05"),
]

TRACE = [
    ("요청서(2026.8.13) M1", "통합 등록화면·동적 패널", "TF11, R-TF-01", ""),
    ("요청서 M2", "5단계 금리·개설은행 3축·정책 프리미엄·통제장치", "TB_TF_DISCOUNT, TB_TF_ISSUER_EVAL, TB_TF_POLICY_ADJ, TF41, R-TF-02~04,12", "화면설계 v6 예시(2.95%, USD 14,756)는 계산 오류 → 2.99%, USD 14,962.50로 정정"),
    ("요청서 M3", "5층 한도", "TB_TF_LIMIT(_UTIL), TF31, R-TF-05", "단일거래상대방 10%는 RBI 대규모 익스포저 규정과 정합성 확인"),
    ("요청서 M4", "A/P(21100000) 경유 지급", "TB_AP_ITEM(타발 중계 미지급금과 통합), R-TF-09", "기존 설계의 별도 Remittance Payable GL을 21100000으로 통일"),
    ("요청서 M5", "FEMA 자동검증·EDPMS/IDPMS", "R-TF-06~08,16", "회수기한 15개월 → 9개월(INR 12개월), 수입 유전스 '규정 상한' → 내부정책으로 정정"),
    ("요청서 M6", "EOD 9개 배치·Dry Run", "TF81, R-TF-10/11", "SMA-0 포함 CRILC 보고 범위 확인"),
    ("요청서 M7", "TReDS 3개 플랫폼", "TB_TF_TREDS, IF-TREDS-01, R-TF-14", "TReDS Directions 2026 반영(보험료 매출자 부담 불가, 재양도)"),
    ("요청서 M8 / SFMS IF 요건", "SFMS·Velocis, 규칙 엔진, 예외처리", "TB_SFMS_MSG, SF11, R-VAL/RTE/CTL/EXC/MCH/IFN", "기존 RT22 NEFT/RTGS 지급(TB_DOM_PAYOUT)을 동일 SFMS 경로로 통합"),
    ("사용자 요청(2026.11)", "신용장·보증서 포함", "TB_LC_*, TB_BG_*, LC11, BG11, R-LC-*, R-BG-*", "업로드 자료에는 LC 매입만 있고 LC 개설·보증 모듈은 없어 신규 설계(v6 매뉴얼에도 'L/C 별도 모듈'로 표기)"),
]

# ---------------------------------------------------------------------------
# RBI 최신 기준 검토 (2026-11-01 기준)
# (No, 대상, 기존 내용, RBI 최신 기준, 반영, 근거, 확인상태)
# 확인상태: 2차출처 = 법률·회계법인 요약으로 확인, RBI 원문(관보) 대조 필요 / 미확인 = 출처 미발견
# ---------------------------------------------------------------------------
RBI_REVIEW = [
    ("1", "요청서 M5·규제근거 / 기존 설계", "수출대금 회수기한 15개월 (FEMA 23(R)/(7)/2025-RB) / 기존 설계 9개월(한시 15개월)",
     "2026.10.1 시행 FEM (Export and Import of Goods and Services) Regulations, 2026 (FEMA 23(R)/2026-RB) - 2026.9.22 개정(FEMA 23(R)/(1)/2026-RB)으로 9개월, INR 인보이스·결제분 12개월. 물품은 선적일, 서비스는 인보이스일 기산",
     "EXP_REALIZE_MONTHS=9, EXP_REALIZE_MONTHS_INR=12, R-TF-07, R-EXP-05. 2026.10.1 이전 선적분은 파라미터 이력", "FEMA 23(R)/2026-RB, 23(R)/(1)/2026-RB", "2차출처"),
    ("2", "기존 설계", "수입대금 결제 선적 후 6개월(초과 시 연장·무역신용)",
     "고정 6개월 폐지. 계약상 기한에 따르며 AD가 사유 검토 후 연장. 수입 미이행 선수금은 반환",
     "IMP_SETTLE_MONTHS → 지점 모니터링 기준으로 격하, R-IMP-05 수정", "FEM EXIM Regs/Directions 2026", "2차출처"),
    ("3", "기존 설계", "수입 선수금 USD 200,000 초과 시 BG/SBLC 필수",
     "금액 기준 폐지 → AD가 주문의 진정성·선수금 필요성 평가, 필요 시 SBLC/보증 요구(AD 내부정책·SOP)",
     "IMP_ADV_BG_THRESHOLD_USD를 지점 정책값으로 유지(규정 아님), R-IMP-06 수정", "RBI Directions on Export and Import, 2026.1.16", "2차출처"),
    ("4", "기존 설계", "선수금 수입기한 6개월(자본재 3년)+15일 증빙",
     "2026 Directions의 세부 기한 미확인. 수입 미이행 시 반환 의무는 유지",
     "기존 값을 지점 정책으로 유지, VERIFY", "FEM EXIM Directions 2026", "미확인"),
    ("5", "기존 설계", "수출 미회수 상각 자체 5%·AD 10%",
     "% 상한 폐지, AD 이사회 승인 정책·SOP로 감액·상각·상계·제3자 지급 처리. 선적서류·인보이스당 ₹10 lakh 이하는 수출자 신고로 종결 가능",
     "EXP_WRITEOFF_* → AD 정책, EXP_DECL_CLOSE_INR=1,000,000 신설, R-TF-16", "FEM EXIM Regs/Directions 2026", "2차출처"),
    ("6", "기존 설계", "수출선수금 수취 후 선적 1년(한시 3년)",
     "3년 (2025.11 개정이 2026 체계에 승계된 것으로 보도)", "EXP_ADV_SHIP_MONTHS=36, R-EXP-06", "FEMA 23(R)/(7)/2025-RB → 2026 Regs", "2차출처"),
    ("7", "신규", "-", "서비스 수출 EDF(Export Declaration Form): 인보이스 월말 후 30일 내 제출. 무상 수출은 NIL 가액 신고",
     "EXP_SVC_EDF_DAYS=30, 기한관리 EDF_SVC", "FEM EXIM Regs 2026", "2차출처"),
    ("8", "신규", "-", "물품·서비스 간 교차 상계 허용(서면동의 불요), 중계무역 별도 조항, 신규 보고는 FETERS 경유",
     "상계 규칙을 AD 정책 기반으로 일반화, 보고 IF 유지", "FEM EXIM Regs 2026", "2차출처"),
    ("9", "요청서 M5", "수입 유전스 상한 일반 180일·자본재 360일 (RBI 수입 규정)",
     "RBI 고정 상한 아님(계약 기한 기준). 장기 후지급은 무역신용 프레임워크(자본재 3년, 비자본재 1년/영업주기)",
     "TF_IMP_USANCE_POLICY(지점 정책), R-TF-08", "FEM EXIM Regs 2026 / MD - Trade Credit", "2차출처"),
    ("10", "요청서 M5", "OA 결제기한 선적일+6개월 초과 → 무역신용 등록",
     "무역신용: 건당 USD 50M(석유·항공·해운 150M), 자본재 3년·비자본재 1년. 2026.2 ECB 개편 후 무역신용 장은 MD 유지. all-in-cost 상한·보고 방식 확인 필요",
     "TC_MAX_USD, TC_TENOR, R-TF-06 (판정 기준 재확인 VERIFY)", "FEM (Borrowing and Lending) Regs 개정 2026.2", "2차출처"),
    ("11", "요청서", "(이자보전 언급 없음)", "Interest Equalisation Scheme 2024.12.31 종료 → EPM Niryat Protsahan 이자보전(MSME 연 ₹50 lakh 한도), 2026.4.1부터 EXIM Bank 청구",
     "TF_EPM_SUBV_CAP_INR, R-TF-13, IF-EXIM-01", "RBI/2025-26/195 (2026.1.19)", "2차출처"),
    ("12", "참고", "2025.11 무역구제 조치", "RBI (Trade Relief Measures) Directions, 2025: 수출신용 450일·유예 등. 2026.6.30 실행분까지 연장, 이후 연장 미확인",
     "신규 실행은 통상 365일(TF_EXP_CREDIT_MAX_DAYS), 기존 건은 이력 관리", "Trade Relief Measures Directions 2025/2026", "미확인"),
    ("13", "요청서·기존 설계 전반", "Master Circular / MD-KYC 2016 등 인용",
     "2025.11.28 RBI가 마스터서큘러를 기관유형별 'Commercial Banks – … Directions, 2025'로 통합(Credit Facilities, Interest Rates on Advances, IRACP, KYC, Credit Risk Management)",
     "근거 표기 일괄 갱신", "RBI 2025.11.28 통합 Directions", "2차출처"),
    ("14", "신규(LC·BG)", "-", "RBI (Non-Fund Based Credit Facilities) Directions, 2025 (RBI/2025-26/140, 2026.4.1 시행): 비자금여신은 원칙적으로 자금여신 거래 고객에 한정, 타 규제기관 자금여신 지원 보증 금지(무역거래 예외), 청구 시 법원 금지명령 없으면 지급, 부분신용보강 50%",
     "R-LC-02, R-BG-01, R-BG-05, PCE_MAX_PCT", "RBI/2025-26/140", "2차출처"),
    ("15", "신규(BG)", "-", "FEM (Guarantees) Regulations, 2026 (FEMA 8(R)/2026-RB, 2026.1.10 시행): 기초거래가 FEMA에 부합하면 원칙 허용, AD 분기 보고(분기말+30일), 지연보고 수수료",
     "R-BG-04, BG_FEMA_RPT_DAYS, IF-GTEE-01", "FEMA 8(R)/2026-RB", "2차출처"),
    ("16", "요청서 M7", "RBI TReDS 지침(2014·2023)", "RBI (TReDS) Directions, 2026 (2026.6.23): 보험(보험료 매출자 부담 불가), 신용보증, 재양도(2차 시장), 무소구 유지, 온보딩 완화",
     "R-TF-14, TB_TF_TREDS", "RBI/DPSS/2026-27/406", "2차출처"),
    ("17", "요청서 M6", "SMA-0/1/2·NPA, SMA-1 이상 CRILC 보고", "IRACP Directions 2025: SMA-1 31~60, SMA-2 61~90, NPA 90일 초과, NPA 자동 판정 의무. CRILC ₹5 crore 기준·SMA-0 보고 범위는 원문 확인 필요",
     "R-TF-10 (자동판정·수기변경 불가), IF-CRILC-01 VERIFY", "RBI/DOR/2025-26/164", "2차출처/일부 미확인"),
    ("18", "SFMS IF 요건", "국내 LC·보증은 SFMS IFN", "국내 BG IFN 760COV는 IBA·IFTAS 관행으로 정착, 외국계 은행에 대한 RBI 의무 규정은 미확인. e-BG(NeSL)도 RBI 의무화 미확인",
     "BG_SFMS_COV_REQUIRED=Y(지점 정책), IF-NESL-01", "IBA/IFTAS", "미확인"),
    ("19", "SFMS IF 요건", "RTGS ISO 20022 2024.11, NEFT 2025.1 전환", "RTGS는 2013 NG-RTGS부터 ISO 20022 기반. NEFT 전환 시점은 미확인 - 벤더(Velocis)·IFTAS 확인 필요. RTGS ₹2 lakh 최소금액 유지",
     "메시지 버전 설정테이블 관리 유지", "RBI RTGS", "미확인"),
    ("20", "SWIFT", "MX 전환 완료(지점 확인)", "결제 MT(Cat 1·2·9 결제) 2025.11.22 종료. MT9xx 명세서는 수신기관 ISO 지원 2027.11, MT9xx 은퇴 2028.11 예정. camt.110 수신 2026.11, 조사 camt.110/111 의무 2027.11. Cat7(MT7xx)은 MT 유지",
     "MT950 자동발송 유지(2028.11 camt.053 전환 계획), KYC 요청 camt.110 기본", "SWIFT ISO 20022 programme (은행 FAQ)", "2차출처"),
    ("21", "화면설계 v6", "할인비율 2.95%, 할인료 USD 14,756 (USD 500,000, 11.97%, 90일)", "산식대로면 2.99%, USD 14,962.50 (요청서·매뉴얼과 일치)", "화면 예시 정정", "산식 검증", "계산확인"),
    ("23", "화면설계 v6", "USD 어음 할인에 MCLR 1년물(9.97%, INR 기준금리)을 기준금리로 적용", "요청서 4.2①: 외화는 SOFR·EURIBOR + 조달마진. 금리는 이사회 정책(Interest Rates on Advances Directions 2025)",
     "통화별 기준금리 분기 구현(INR MCLR / USD SOFR+마진 / EUR EURIBOR+마진)", "요청서 4.2", "내부정합"),
    ("22", "FDI", "송금인 KYC 보고서", "AD은행은 송금 해외은행으로부터 투자자 KYC 보고서 확보(유지). FC-GPR 배정 후 30일, 배정 60일 유지", "R-KYC-*", "NDI Rules 2019 / MD Foreign Investment", "2차출처"),
]
