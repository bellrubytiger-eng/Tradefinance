"""본점(글로벌사업부) 'IFRS 관련 확인 요청서' — 노이다지점 회신본 생성.

원본 양식(tools/data/ifrs_gap_template.xlsx)을 그대로 열어 노이다지점 칸과 상세 시트 빈 행만 채운다.
뉴욕지점 사전 기재분·서식·드롭다운·수식은 건드리지 않는다.

근거: FY25-26 결산(2026-03-31 TB·AS 19 working·Deferred Tax Computation·Asset Quality·Sch.12),
      GBS 현행 BS/PL(2026-07-07), gaap_model.GAAP_DIFF, decisions_model(G9·F2·L3~L6).
      금액은 INR(현지통화), 기준일 2026-03-31 (별도 표기 시 제외).

python3 tools/build_ifrs_gap.py  →  docs/08_IFRS_갭분석_노이다지점.xlsx
"""
import math
import os
import sys

import openpyxl
from openpyxl.styles import PatternFill

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "tools", "data", "ifrs_gap_template.xlsx")
OUT = os.path.join(ROOT, "docs", "08_IFRS_갭분석_노이다지점.xlsx")
BR = "노이다지점"
BASE = "2026-03-31"
CHECK = PatternFill("solid", fgColor="FFFFF2CC")  # 계약서 대조 등 지점 확인이 남은 칸

# ───────────────────────── 총괄표 I열(노이다지점) ─────────────────────────
SUMMARY = {
    5: "현지 GAAP",
    6: "인도 GAAP: ICAI 회계기준(AS) + RBI 지침(IRACP·투자 Directions 2023·ECL Directions 2026) + "
       "은행법(BR Act) Third Schedule 양식. ※ 은행은 Ind AS(인도판 IFRS) 시행 유예로 미적용",
    7: "일부 수행",
    8: "현행: 대손충당금(본점 기재 내용). 추가 필요: ①대출 취급수수료 EIR 이연(현지 일시인식) "
       "②리스 IFRS 16(현지 AS 19 운용리스) ③임차보증금 현재가치 ④복구충당부채 "
       "⑤보험수리손익 OCI(AS 15) ⑥이연법인세(AS 22→IAS 12). 상세 2_IFRS조정 7건",
    9: "미수행",
    10: "미수행",
    11: "수행",
    12: "수행",
    13: "없음(운전자금 한도·단기 운전·시설 대출 모두 원리금 구조. 후순위·전환·수익연계 조건 없음)",
    14: "4개 유형 12개 계정(FY25-26 실적): ①대출(LOAN APPLICATION FEE 42020701, OTHER FEE 42020729) "
        "②보증(GUARANTEE(Bond) 42010013) ③송금(42020601·602·603·606·629) "
        "④상환·연장·우편(42020401·402·403, 42020307)",
    15: "이연 0개 / 일시인식 12개. 기준: 수취 시 수익(AS 9). 대출 취급수수료는 IFRS 9상 EIR 구성요소라 이연 대상, "
        "보증수수료는 보증기간 안분 여부 확인",
    16: "해당없음",
    17: "해당없음",
    18: "확인 필요",
    19: "계약서 대조 중. AS 19 리스료 스케줄: 매월 13일 지급, 2025-04 +10%(₹14,55,487→₹16,01,036), "
        "2028-04 +10.5%(→₹17,69,145), 종료 2031-03 추정",
    20: "확인 필요",
    21: "확인 필요",
    22: "확인 필요",
    23: "미계상",
    24: "현지 원장에 복구충당부채 계정 없음(AS 29 검토 미실시). 계약 조항 확인 후 의무가 있으면 "
        "인도 GAAP(AS 29, 할인 없음)으로도 계상 대상 → 차세대 CoA 20940000000 신설",
    25: "①NPA 이자 비인식(IRACP) ②지급보증 금융보증·ECL ③법정준비금 25%(자본 내 대체) "
        "④전기오류 AS 5 ⑤AS 19 임대료 정액 인식 여부 ⑥2027-04 RBI ECL·EIR 시행 후 차이 축소. 상세 13_기타",
}
SUMMARY_CHECK = {18, 19, 20, 21, 22}

# ───────────────────────── 2_IFRS조정 ─────────────────────────
# (관련 계정, 현지 GAAP, IFRS, 수행 여부, 주체, 조정 금액, 기준일, 비고)
IFRS_ADJ = [
    ("대출채권 손실충당금\n(10690001 Allowance For Loan Losses-General / 차세대 10380101000)",
     "RBI IRACP: 정상자산 일반충당금(대부분 0.4%) + NPA 특정충당금(연체일수 기준). 2026-03-31 ₹1,52,70,000(전액 정상, NPA 없음)",
     "IFRS 9 기대신용손실(Stage 1·2·3, 12개월/전체기간). 은행 예치·콜론·HTM 채권·부외 약정·보증 포함",
     "수행", "글로벌사업부", "본점 산출", BASE,
     "총괄표 2-1 본점 기재 기준(본점 ECL 산출). 2027-04-01 RBI ECL Directions 2026 시행으로 현지 장부도 "
     "Stage ECL(규제 하한)로 전환 → 하한·모델 차이만 남음. 차세대 CoA Stage별 L5 10380103000~105000"),
    ("대출 취급수수료\n(42020701 LOAN APPLICATION FEE / 차세대 40580010000, 이연 LOF 10370100000)",
     "수취 시 일시 수익(AS 9). FY25-26 ₹26,40,000, FY26-27(~07-07) ₹16,40,000",
     "유효이자율 구성요소로 이연 후 EIR 상각, 이자수익 표시(IFRS 9 B5.4.1~B5.4.3)",
     "미수행", "지점", "산정 필요(대출별 미상각 잔액)", BASE,
     "뉴욕 이슈 4 해당(이연 자체가 없어 차이 더 큼). 2027-04-01부터 RBI ECL Directions 2026의 EIR 적용으로 "
     "현지 장부도 이연 전환 예정 — 그 전까지 본점 조정. GBS가 LOF 이연·EIR 상각 데이터 산출"),
    ("임차료(리스)\n(52032421 RENTAL EXPENSE / 차세대 50940021405) — 사무실 1건·사택 5건",
     "AS 19 운용리스: 리스료 비용(FY25-26 ₹3,88,82,301), 미래 최소리스료 주석 공시. 사용권자산·리스부채 미인식",
     "IFRS 16: 사용권자산·리스부채 인식, 이자비용·감가상각. 12개월 이하 단기·소액 리스 면제 선택 가능",
     "미수행", "미정", "산정 필요(리스 원장)", BASE,
     "현지 원장은 인도 GAAP 유지(은행 Ind AS 116 미적용). 차세대 GBS는 IFRS 16 계정(ROU 107·리스부채 206·이자 "
     "50320500000)을 본점 조정 전용으로 두고 리스 원장 속성으로 월 스케줄 산출. 본점용 리스 템플릿을 이미 "
     "제출 중이면 '수행·지점'으로 정정"),
    ("임차보증금\n(10999002 Security Deposit / 차세대 10490700100)",
     "명목금액 계상 ₹1,31,30,642, 할인 없음",
     "상각후원가(현재가치) 측정. 명목·현재가치 차이는 사용권자산 가산, 할인 상각은 이자수익(IFRS 9·IFRS 16)",
     "미수행", "미정", "산정 필요", BASE,
     "리스와 함께 조정. 차세대 조정 전용 계정 10490700900(할인차금)·40140100000(이자수익)"),
    ("복구충당부채\n(현행 계정 없음 / 차세대 20940000000 신설)",
     "미계상. AS 29는 의무가 있으면 할인 없이 최선 추정치로 인식",
     "IAS 37 현재가치로 인식, 사용권자산 원가 가산, 할인 상각 이자비용",
     "미수행", "지점", "산정 필요", BASE,
     "계약서 원상복구 조항 확인 중(인테리어 Leasehold Improvement 원가 ₹4,78,38,399). 의무가 있으면 인도 GAAP 자체 "
     "보완 사항. 할인 상각 50320200000은 본점 조정 전용"),
    ("퇴직급여(Gratuity)\n(WELFARE EXPENSE 52032xxx / 차세대 50930170100·200)",
     "AS 15: 보험수리평가(2026-03-31 충당부채 ₹13,25,421), 보험수리손익 전액 당기손익",
     "IAS 19: 재측정요소(보험수리손익) 기타포괄손익",
     "미수행", "글로벌사업부", "보험수리 보고서상 손익", BASE,
     "손익↔OCI 분류 차이만 발생, 순자산 영향 없음. 금액 중요성 낮음"),
    ("이연법인세\n(23100101 DEFERRED Tax Liabilities / 차세대 20720xxxxxx)",
     "AS 22: 기간 차이(손익 접근). 2026-03-31 순 DTL ₹20,54,203(감가상각 DTL, 대손충당금·Gratuity DTA), 세율 37.128%",
     "IAS 12: 일시적 차이(재무상태표 접근). 리스·EIR·복구충당부채 등 IFRS 조정분의 이연법인세 추가",
     "미수행", "글로벌사업부", "조정 항목별 × 37.128%", BASE,
     "IFRS 조정의 부수 효과. 리스 이연법인세 계정 10820400000·20720200000은 본점 조정 전용"),
]

# ───────────────────────── 3_금융상품분류 ─────────────────────────
# (유형, 세부, 현지 분류, SPPI 수행, 충족, 사업모형 수행, 사업모형, IFRS 9, 잔액, 비SPPI, 비고)
_LOAN_NOTE = "공식 테스트 미수행. 상품 조건 약식 검토상 충족(원리금 외 현금흐름 없음). 연체이자는 RBI 지침상 이자 아닌 penal charge"
FIN = [
    ("대출채권", "운전자금 한도대출(WCDL) 10610004", "상각후원가(IRACP)", "미수행", "충족", "미수행", "보유", "AC",
     600000000, "없음", _LOAN_NOTE),
    ("대출채권", "단기 운전자금 대출 10610005", "상각후원가(IRACP)", "미수행", "충족", "미수행", "보유", "AC",
     1480000000, "없음", _LOAN_NOTE),
    ("대출채권", "시설자금 기한부 대출 10610007", "상각후원가(IRACP)", "미수행", "충족", "미수행", "보유", "AC",
     987500000, "없음", _LOAN_NOTE),
    ("예치금", "국내 은행 정기예치 10255002", "상각후원가", "미수행", "충족", "미수행", "보유", "AC",
     2018960000, "없음", "은행간 정기예치(총자산 약 35%). 차세대 ECL 범위 포함(결정 L6)"),
    ("예치금", "콜론(Call Loans-Domestic) 10251001", "상각후원가", "미수행", "충족", "미수행", "보유", "AC",
     100000000, "없음", "익일물 은행간 대여"),
    ("예치금", "RBI 당좌·법정예치(BR Act s.11(2)(b)) 10211001·10211002", "상각후원가", "해당없음", "해당없음",
     "해당없음", "보유", "AC", 27388808, "없음", "법정 예치 ₹15,00,000 포함. 중앙은행 예치로 ECL 영향 없음"),
    ("예치금", "Nostro 요구불(국내·해외) 10212001·10212002", "상각후원가", "해당없음", "해당없음", "해당없음",
     "보유", "AC", 20358575, "없음", "결제용 요구불 계좌"),
    ("유가증권", "인도 국채 HTM(SLR 보유분) 10330001", "HTM(RBI 투자 Directions 2023)", "수행", "충족", "수행",
     "보유", "AC", 56886300, "없음",
     "RBI 투자 Directions 2023: HTM 편입 요건에 SPPI 충족·만기보유 의도 포함, 이후 매각 제한 → IFRS 9 AC와 실질 동일"),
    ("매입외환", "수출환어음 매입·할인", "상각후원가(IRACP)", "해당없음", "해당없음", "해당없음", "해당없음", "AC",
     0, "없음", "2026-03-31 잔액 없음. 취급 시 단기 원리금 구조로 AC 예상"),
    ("기타", "지급보증(Bond 보증, 부외)", "우발채무(Sch.12)", "해당없음", "해당없음", "해당없음", "해당없음",
     "해당없음", 5695709, "-", "IFRS 9 금융보증계약: 수수료로 최초 인식 후 ECL과 미상각 잔액 중 큰 금액. 13_기타 2번"),
]

# ───────────────────────── 4-7_COMMISSION ─────────────────────────
# (코드, 계정명, 성격, 관련 상품, 이연/일시, 본점 TB, 당기 인식액, 비고)
COMM = [
    ("42020701", "COMMISSION - LOAN APPLICATION FEE", "대출 신청·심사(Processing) 수수료", "기업 대출(WCDL·단기·시설)",
     "일시인식", "수수료수익", 2640000,
     "IFRS 9상 EIR 구성요소 → 이연·EIR 상각, 이자수익 표시 필요(2_IFRS조정 2번). 2027-04 RBI EIR 적용 예정"),
    ("42020729", "COMMISSION - OTHER FEE", "대출 부대 기타 수수료", "기업 대출", "일시인식", "수수료수익", 217781,
     "건별 성격 확인 후 EIR 구성요소면 이연 대상"),
    ("42010013", "COMMISSION-GUARANTEE(Bond)", "이행·입찰보증(Bond) 발행 수수료", "지급보증", "일시인식", "수수료수익",
     38716, "보증기간 안분 여부 확인. IFRS 9 금융보증은 보증기간에 걸쳐 인식. 기말 보증잔액 ₹56,95,709"),
    ("42020601", "COMMISSION - REMITTANCE OF OUTWARD", "당발송금 취급수수료", "송금", "일시인식", "수수료수익", 193250,
     "용역 제공 시점 인식 — IFRS 15와 차이 없음"),
    ("42020606", "Commission - Remittance Of Outward(OUR)", "당발송금 OUR 조건 중계은행 수수료 회수", "송금", "일시인식",
     "수수료수익", 1557, "실제 중계은행 비용과 금액이 달라 과세 수익(본인)으로 총액 인식 — IFRS 15와 일치"),
    ("42020602", "COMMISSION - REMITTANCE OF INWARD", "타발송금 수수료", "송금", "일시인식", "수수료수익", 1200,
     "차이 없음"),
    ("42020603", "COMMISSION - CABLE AND POSTAGE(REMITTANCE)", "송금 SWIFT 전신료", "송금", "일시인식", "수수료수익",
     18000, "차이 없음"),
    ("42020629", "COMMISSION - OTHERS(REMITTANCE)", "송금 관련 기타 수수료", "송금·외환", "일시인식", "수수료수익",
     1126026, "차이 없음. 차세대에서 취급·중계·SWIFT로 세분(CoA 검토 결정사항)"),
    ("42020401", "COMMISSION - REIMBURSEMENT", "수입 L/C 상환(Reimbursement) 수수료", "무역금융", "일시인식",
     "수수료수익", 85388, "차이 없음"),
    ("42020402", "COMMISSION - CABLE AND POSTAGE(REIMBURSEMENT)", "상환 관련 전신료", "무역금융", "일시인식",
     "수수료수익", 14354, "차이 없음"),
    ("42020403", "COMMISSION - EXTENTION", "인수·상환 만기 연장 수수료", "무역금융", "일시인식", "수수료수익", 22921,
     "신용공여 연장과 관련되면 EIR 구성요소 검토(금액 경미)"),
    ("42020307", "COMMISSION - POSTAGE/DHL(EXPORT)", "수출 서류 우편료 회수", "무역금융", "일시인식", "수수료수익",
     1000, "차이 없음"),
    ("42020704 외", "LOAN COMMITMENT FEE·SYNDICATED LOAN PARTICIPATION/ARRANGE FEE·UNUSED LINE FEE 등",
     "약정·신디케이션 수수료", "신디케이트론·약정", None, None, 0,
     "FY25-26 실적 없음. 발생 시 약정수수료는 인출 가능성에 따라 이연(IFRS 9 B5.4.3) — 차세대 LOF 10370100000"),
]

# ───────────────────────── 8-12_리스·복구 ─────────────────────────
# 열: C 대상, D 용도, E 개시, F rent free, G rent 개시, H 만료, I 기산, J·K 템플릿, L 일치, M 계약 조정, N 템플릿 조정,
#     O 일치, P 옵션, Q 원상복구, R 계상, S 금액, T 미계상 사유, U 비고
_NO_TPL = "-(IFRS 16 템플릿 없음, AS 19 리스료 스케줄 운영)"
LEASE = [
    dict(C="노이다지점 영업점 사무실", D="사무실", E="확인 필요", F="확인 필요", G="확인 필요",
         H="확인 필요(AS 19 스케줄상 2031-03 종료 추정)", J=_NO_TPL, K="AS 19 스케줄 2031-03까지",
         M="확인 필요(계약서 조항)",
         N="AS 19 스케줄: 3년 주기 인상 2025-04 +10%(₹14,55,487→₹16,01,036), 2028-04 +10.5%(→₹17,69,145)",
         P="확인 필요", Q="확인 필요", R="미계상", S="-", T="조항 확인 전. 의무가 있으면 AS 29(할인 없음)로도 계상 대상",
         U="계약서 대조 후 E·F·G·H·M·P·Q 확정. 임차보증금 ₹1,31,30,642(명목). Leasehold Improvement 원가 ₹4,78,38,399"),
    dict(C="주재원 사택 1", D="사택", E="2024-08(첫 임대료 월)", H="2026-12(스케줄)", J=_NO_TPL,
         M="2025-07 월 ₹3,23,000→₹3,25,000", P="확인 필요", Q="확인 필요", R="미계상", S="-",
         T="주거용, 원상복구 비용 경미 예상",
         U="계약기간 12개월 초과 시 IFRS 16 단기리스 면제 불가(갱신옵션 포함 판단)"),
    dict(C="주재원 사택 2", D="사택", E="2025-02(첫 임대료 월)", H="2025-12(종료)", J=_NO_TPL,
         M="조정 없음(월 ₹3,00,000)", P="-", Q="-", R="해당없음", S="-", T="-", U="종료 계약. FY25-26 비용만 해당"),
    dict(C="주재원 사택 3", D="사택", E="2024-04(첫 임대료 월)", H="2026-03(종료)", J=_NO_TPL,
         M="조정 없음(월 ₹2,99,000)", P="-", Q="-", R="해당없음", S="-", T="-", U="종료 계약. FY25-26 비용만 해당"),
    dict(C="주재원 사택 4", D="사택", E="2024-03(첫 임대료 월)", H="2026-11(스케줄)", J=_NO_TPL,
         M="조정 없음(월 ₹2,94,000)", P="확인 필요", Q="확인 필요", R="미계상", S="-",
         T="주거용, 원상복구 비용 경미 예상", U="12개월 초과 — IFRS 16 대상"),
    dict(C="주재원 사택 5", D="사택", E="2026-02(첫 임대료 월)", H="2026-12(스케줄)", J=_NO_TPL,
         M="조정 없음(월 ₹3,83,500)", P="확인 필요", Q="확인 필요", R="미계상", S="-",
         T="주거용, 원상복구 비용 경미 예상", U="계약기간 확인 후 단기리스 면제 여부 판단"),
    dict(C="기타 임차(SUNDRY OFFICE RENTAL EXPENSE 52032261)", D="기타", E="확인 필요", H="확인 필요", J=_NO_TPL,
         P="확인 필요", Q="확인 필요", R="미계상", S="-", T="계약 단위 내역 확인 전",
         U="FY25-26 ₹30,08,161. 계약 단위 내역 확인 후 리스 해당 여부(단기·소액 면제) 판단"),
]

# ───────────────────────── 13_기타 ─────────────────────────
# (관련 계정, 내용, 현지 처리, IFRS 예상 처리, 금액, 처리 방안, 비고)
OTHER = [
    ("대출 이자수익·미수이자", "NPA 이자 인식",
     "RBI IRACP: 90일 연체 등으로 NPA 분류 시 미수이자 환입, 이후 현금주의",
     "IFRS 9 Stage 3: 손상 후 순장부금액 × EIR로 이자 계속 인식", "0(2026-03-31 NPA 없음)",
     "발생 시 본점 조정. GBS가 비인식 이자를 메모 속성으로 보관", "차이 구조만 존재, 현재 금액 없음"),
    ("지급보증·미사용 약정(부외)", "금융보증·대출약정 손실충당금",
     "Sch.12 우발채무 공시. 보증 손실충당금 0(정상)",
     "IFRS 9 금융보증: 최초 수수료로 인식 후 ECL과 미상각 잔액 중 큰 금액. 미사용 약정 ECL",
     "보증잔액 5,695,709",
     "차세대 ECL 범위에 부외(보증·약정) 포함(결정 L6), CCF 속성 관리", "본점 ECL 산출 범위 포함 여부 확인"),
    ("법정준비금(차세대 31110300000)", "법정준비금 적립",
     "은행법 s.17: 순이익 25% 법정준비금 전입(손익 처분). FY25-26 ₹54,20,250",
     "자본 내 이익잉여금 대체, 손익 영향 없음", 5420250,
     "본점 합산 시 이익잉여금으로 통합 표시", "표시 차이만"),
    ("전기 항목(차세대 40890000000·51208000000)", "전기오류 처리",
     "AS 5: 당기 손익에 'Prior period items' 별도 표시",
     "IAS 8: 중요한 오류는 비교표시 재작성(소급)", "-",
     "중요한 건 발생 시 본점 비교표시 재작성", "-"),
    ("임차료 52032421", "AS 19 임대료 정액 인식 여부",
     "AS 19는 리스기간 정액 비용 인식이 원칙. 3년 주기 인상(2025-04·2028-04)분의 균등화(lease equalisation) 반영 여부 확인 필요",
     "IFRS 16 적용 시 사용권자산 상각·이자로 대체되어 해소", "확인 필요",
     "인도 GAAP 자체 점검 사항. 차세대 리스 원장에서 정액 스케줄 산출", "FY25-26 임차료 ₹3,88,82,301"),
    ("대손충당금·대출 수수료", "2027-04-01 RBI ECL·EIR 시행",
     "RBI ECL Directions 2026 시행: Stage ECL(규제 하한), 대출 수수료 EIR 상각. 최초 적용 영향은 이익잉여금 조정",
     "이미 IFRS 9 적용", "-",
     "2027-03-31 기준 차이 재산정 후 본점 조정 항목 축소(대손충당금·EIR 조정이 하한 차이로 축소). 차세대 전환 조정 31110800000",
     "현지 원장이 IFRS에 수렴"),
    ("외화환산·HTM 채권·송금/무역 수수료", "차이 없음 확인 항목",
     "AS 11 외화환산(결산 환율), RBI 투자 Directions 2023 HTM 상각후원가, 송금·무역 수수료 용역 시점 인식",
     "IAS 21·IFRS 9 AC·IFRS 15와 동일", "-", "조정 불요", "확인 완료"),
]

# ───────────────────────── 뉴욕지점_이슈처리 G열 ─────────────────────────
NY_ISSUE = {
    5: ("O", "노이다: 인도 GAAP 장부, 대손충당금 외 conversion 없음"),
    6: ("O", "노이다: SPPI 체계 없음(대출 원리금 구조로 중요 차이 없음). 국채 HTM은 RBI 2023 지침상 SPPI 요건 적용"),
    7: ("X", "노이다: 조건변경 대출 없음(2026-03-31 NPA·재조정 없음)"),
    8: ("O", "노이다: 대출 취급수수료 일시인식(이연 없음) FY25-26 ₹26,40,000 → 본점 기재 'X'를 'O'로 정정"),
    10: ("확인 중", "노이다: IFRS 16 템플릿 없음(AS 19 스케줄). 계약서 대조 중"),
    11: ("확인 중", "노이다: AS 19 스케줄상 인상 2025-04·2028-04, 계약서 조항 대조 중"),
    12: ("확인 중", "노이다: 원상복구 조항 확인 중, 복구충당부채 미계상"),
}


def _lines(text, width):
    if text is None:
        return 1
    chars_per_line = max(4, (width or 9) * 1.1)
    n = 0
    for part in str(text).split("\n"):
        # 한글은 폭이 약 2배
        w = sum(2 if ord(ch) > 0x2E80 else 1 for ch in part)
        n += max(1, math.ceil(w / chars_per_line))
    return n


def _fit(ws, row, min_h=24):
    best = 1
    for c in ws[row]:
        if c.value is None:
            continue
        width = ws.column_dimensions[c.column_letter].width or 9
        best = max(best, _lines(c.value, width))
    ws.row_dimensions[row].height = max(min_h, min(400, best * 12.5 + 6))


def _put(ws, row, col, value, check=False):
    cell = ws.cell(row, col)
    cell.value = value
    if check:
        cell.fill = CHECK
    if isinstance(value, (int, float)):
        cell.number_format = "#,##0"
    return cell


def _first_empty(ws, start=5):
    r = start
    while ws.cell(r, 2).value not in (None, ""):
        r += 1
    return r


def _needs_check(v):
    return isinstance(v, str) and ("확인 필요" in v or v == "확인 중")


def build(src=SRC, out=OUT):
    wb = openpyxl.load_workbook(src)

    ws = wb["총괄표"]
    ws.column_dimensions["I"].width = 42
    for r, v in SUMMARY.items():
        _put(ws, r, 9, v, check=r in SUMMARY_CHECK)
    for r in range(5, 26):
        _fit(ws, r, min_h=ws.row_dimensions[r].height or 24)

    ws = wb["2_IFRS조정"]
    for col, w in (("E", 34), ("H", 16), ("J", 44)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for item in IFRS_ADJ:
        _put(ws, r, 2, BR)
        for j, v in enumerate(item, start=3):
            _put(ws, r, j, v)
        _fit(ws, r)
        r += 1

    ws = wb["3_금융상품분류"]
    for col, w in (("C", 12), ("D", 30), ("E", 18), ("M", 40)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for item in FIN:
        _put(ws, r, 2, BR)
        for j, v in enumerate(item, start=3):
            _put(ws, r, j, v)
        _fit(ws, r)
        r += 1

    ws = wb["4-7_COMMISSION"]
    for col, w in (("D", 34), ("E", 26), ("I", 14), ("N", 44)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for code, name, nature, prod, gi, tb, amt, note in COMM:
        _put(ws, r, 2, BR)
        _put(ws, r, 3, code)
        _put(ws, r, 4, name)
        _put(ws, r, 5, nature)
        _put(ws, r, 6, prod)
        if gi:
            _put(ws, r, 7, gi)
            _put(ws, r, 8, "해당없음")
            _put(ws, r, 9, tb)
            _put(ws, r, 10, "해당없음")
            _put(ws, r, 11, "-")
        _put(ws, r, 12, amt)
        _put(ws, r, 13, 0)
        _put(ws, r, 14, f"FY25-26 실적. {note}")
        _fit(ws, r)
        r += 1

    ws = wb["8-12_리스·복구"]
    for col, w in (("H", 18), ("I", 14), ("K", 16), ("P", 14), ("Q", 14), ("U", 40)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for item in LEASE:
        _put(ws, r, 2, BR)
        for col, v in item.items():
            _put(ws, r, openpyxl.utils.column_index_from_string(col), v, check=_needs_check(v))
        _fit(ws, r)
        r += 1

    ws = wb["13_기타"]
    for col, w in (("E", 34), ("F", 30), ("H", 34), ("I", 20)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for item in OTHER:
        _put(ws, r, 2, BR)
        for j, v in enumerate(item, start=3):
            _put(ws, r, j, v, check=_needs_check(v))
        _fit(ws, r)
        r += 1

    ws = wb["뉴욕지점_이슈처리"]
    for r, (mark, note) in NY_ISSUE.items():
        _put(ws, r, 7, mark, check=mark == "확인 중")
        prev = ws.cell(r, 15).value
        _put(ws, r, 15, f"{prev}\n{note}" if prev else note)
        _fit(ws, r, min_h=ws.row_dimensions[r].height or 42)

    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    print(build(*sys.argv[1:3]))
