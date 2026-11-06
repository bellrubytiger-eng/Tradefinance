"""본점(글로벌사업부) 'IFRS 관련 확인 요청서' — 노이다지점 회신본 생성.

원본 양식(tools/data/ifrs_gap_template.xlsx)을 그대로 열어 노이다지점 칸과 상세 시트 빈 행만 채운다.
뉴욕지점 사전 기재분·서식·드롭다운·수식은 건드리지 않는다.

근거: FY25-26 결산(2026-03-31 TB·AS 19 working·Deferred Tax Computation·Asset Quality·Sch.12),
      GBS 현행 BS/PL(2026-07-07), gaap_model.GAAP_DIFF, decisions_model(G9·F2·L3~L6).
      금액은 INR(현지통화), 기준일 2026-03-31 (별도 표기 시 제외).

python3 tools/build_ifrs_gap.py       →  docs/08_IFRS_갭분석_노이다지점.xlsx
python3 tools/build_ifrs_gap.py en    →  docs/08_IFRS_Gap_Analysis_Noida_EN.xlsx (현지 직원 검토용 영문본, ifrs_gap_en.py)
"""
import math
import os
import sys

import re
import types
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, PatternFill

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "tools", "data", "ifrs_gap_template.xlsx")
OUT = os.path.join(ROOT, "docs", "08_IFRS_갭분석_노이다지점.xlsx")
OUT_EN = os.path.join(ROOT, "docs", "08_IFRS_Gap_Analysis_Noida_EN.xlsx")
HANGUL = re.compile("[\uac00-\ud7a3]")
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
    15: "이연 1개 / 일시인식 11개. 보증수수료(42010013)는 선수수익(20920017 Income On Guarantee Money, 2026-03-31 "
        "₹72,838)으로 보증기간 안분 — 인도 GAAP AS 9(용역 제공 기간에 걸쳐 인식)·IFRS 9 금융보증과 같음. 그 외는 "
        "수취 시 수익(AS 9). 대출 취급수수료는 IFRS 9상 EIR 구성요소라 이연 대상",
    16: "수수료수익",
    17: "정액법",
    18: "Rent free 포함",
    19: "등기 리스계약서(2022-08-02 인지): Commencement(점유) 2022-04-13 / Rent Commencement 2022-08-20"
        "(rent free 4개월 7일, 정의 조항 — 영업개시가 더 빠르면 그날) / Expiration 2031-04-12(9년, 3(a)). "
        "임대료 ₹14,55,487 + GST, 매월 7일 선지급(4(a)(b))",
    20: "해당없음",
    21: "해당없음",
    22: "있음(조항 24 Reinstatement)",
    23: "미계상",
    24: "추정 미실시. 원상복구 의무(조항 24: 종료 시 fit-out 전 warm shell 상태로 반환)가 있어 인도 GAAP(AS 29, "
        "할인 없음)으로도 계상 대상 → 원상복구 견적으로 추정 후 계상 예정. 차세대 CoA 20940000000 신설",
    25: "①NPA 이자 비인식(IRACP) ②지급보증 금융보증·ECL(미사용 약정 없음) ③법정준비금 25%(자본 내 대체) "
        "④전기오류 AS 5 ⑤관리비는 리스 아님(용역) ⑥2027-04 RBI ECL·EIR 시행 후 차이 축소. 상세 13_기타",
}
SUMMARY_CHECK = {15}

# ───────────────────────── 2_IFRS조정 ─────────────────────────
# (관련 계정, 현지 GAAP, IFRS, 수행 여부, 주체, 조정 금액, 기준일, 비고)
IFRS_ADJ = [
    ("대출채권 손실충당금\n(10690001 Allowance For Loan Losses-General / 차세대 10380101000)",
     "RBI IRACP 표준법: 지점 산출. 2026-03-31 ₹1,52,70,000 = 정상자산 ₹306.75 crore × 0.4% ₹1,22,70,000 + "
     "미헤지 외화익스포저(UFCE) 추가충당금 ₹30,00,000(전액 정상, NPA 없음)",
     "IFRS 9 기대신용손실(Stage 1·2·3, 12개월/전체기간). 은행 예치·콜론·HTM 채권·부외 약정·보증 포함",
     "수행", "글로벌사업부", "본점 산출", BASE,
     "총괄표 2-1 본점 기재 기준(본점 ECL 산출). 2027-04-01 RBI ECL Directions 2026 시행으로 현지 장부도 "
     "Stage ECL(규제 하한)로 전환 → 하한·모델 차이만 남음. UFCE 추가충당금은 IFRS 9에 별도 개념이 없어 ECL로 "
     "대체(차세대 결정 L8: Stage 충당금에 흡수, UFCE 태그 관리). 차세대 CoA Stage별 L5 10380103000~105000"),
    ("대출 취급수수료\n(42020701 LOAN APPLICATION FEE / 차세대 40580010000, 이연 LOF 10370100000)",
     "수취 시 일시 수익(AS 9). FY25-26 ₹26,40,000, FY26-27(~07-07) ₹16,40,000",
     "유효이자율 구성요소로 이연 후 EIR 상각, 이자수익 표시(IFRS 9 B5.4.1~B5.4.3)",
     "미수행", "지점", "산정 필요(대출별 미상각 잔액)", BASE,
     "뉴욕 이슈 4 해당(이연 자체가 없어 차이 더 큼). 2027-04-01부터 RBI ECL Directions 2026의 EIR 적용으로 "
     "현지 장부도 이연 전환 예정 — 그 전까지 본점 조정. GBS가 LOF 이연·EIR 상각 데이터 산출"),
    ("임차료(리스)\n(52032421 RENTAL EXPENSE / 차세대 50940021405) — 사무실 1건·사택 5건",
     "AS 19 운용리스: 리스기간 정액 비용 인식(FY25-26 임차료 ₹3,88,82,301), 미래 최소리스료 주석 공시. "
     "사용권자산·리스부채 미인식",
     "IFRS 16: 사용권자산·리스부채 인식, 이자비용·감가상각. 12개월 이하 단기·소액 리스 면제 선택 가능",
     "미수행", "미정", "추정 리스부채(사무실) ₹8,21,57,515", BASE,
     "본점 리스 템플릿은 제출하나 '해당없음'으로 표기(IFRS 16 계산 없음) → 본점 합산에 IFRS 16 미반영. 추정액은 "
     "사무실 잔여 60개월 지급액 ₹10,21,14,084를 할인율 9%(가정)로 현재가치 — 중요성 판단 필요. 현지 원장은 인도 "
     "GAAP 유지(은행 Ind AS 116 미적용). 차세대 GBS는 IFRS 16 계정(ROU 107·리스부채 206·이자 50320500000)을 "
     "본점 조정 전용으로 두고 리스 원장 속성으로 월 스케줄 산출"),
    ("임차보증금\n(10999002 Security Deposit / 차세대 10490700100)",
     "명목금액 계상, 할인 없음. Security Deposit ₹1,31,30,642 중 사무실 무이자 보증금(IFRSD) ₹87,32,922(계약 10(a))",
     "상각후원가(현재가치) 측정. 명목·현재가치 차이는 사용권자산 가산, 할인 상각은 이자수익(IFRS 9·IFRS 16)",
     "미수행", "미정", "산정 필요", BASE,
     "리스와 함께 조정. 차세대 조정 전용 계정 10490700900(할인차금)·40140100000(이자수익)"),
    ("복구충당부채\n(현행 계정 없음 / 차세대 20940000000 신설)",
     "계약상 원상복구 의무 있음, 미계상(추정 미실시). AS 29도 할인 없이 최선 추정치로 인식 요구",
     "IAS 37 현재가치로 인식, 사용권자산 원가 가산, 할인 상각 이자비용",
     "미수행", "지점", "산정 필요", BASE,
     "인도 GAAP 자체 보완 사항 — 원상복구 견적으로 추정해 현지 장부에 AS 29로 계상 예정(인테리어 원가 ₹4,78,38,399 "
     "참고). 본점 IFRS 조정은 현재가치 차이·사용권자산 가산만 남음. 할인 상각 50320200000은 본점 조정 전용"),
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
    ("42010013", "COMMISSION-GUARANTEE(Bond)", "이행·입찰보증(Bond) 발행 수수료", "지급보증", "이연", "수수료수익",
     38716, "수취 시 선수수익(20920017 Income On Guarantee Money)으로 받아 보증기간에 걸쳐 수수료수익 인식 — 인도 GAAP "
     "AS 9(용역 제공 기간 안분)·IFRS 9 금융보증과 같음. 기말 보증잔액 ₹56,95,709. 차세대 선수수익 20420170000, 일할 정액 상각",
     dict(H="수수료수익", J="정액법", K="보증기간", M=72838, check=True)),
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
_NO_TPL = "해당없음(템플릿 '해당없음' 표기 제출)"
LEASE = [
    dict(C="노이다지점 영업점 사무실(등기 리스계약)", D="사무실", E="2022-04-13(점유·fit-out 개시, 3(a))",
         F="4개월 7일", G="2022-08-20(정의 조항: 개시 후 4개월 7일 또는 영업개시 중 빠른 날)",
         H="2031-04-12(9년, 3(a))", I="Rent free 포함", J=_NO_TPL, K=_NO_TPL,
         M="조항 4(f): 3년 주기 2회 — Rent Commencement 후 3년 경과 시 +10%(₹14,55,487→₹16,01,036), 6년 경과 시 "
           "+10.5%(→₹17,69,145). 조항 4(f)는 RCD 기준(2025-08-20·2028-08-20), 정의 조항은 개시일 기준"
           "(2025-04-13·2028-04-13)으로 상충 — 지점은 개시일 기준(2025-04·2028-04) 적용",
         N=_NO_TPL,
         P="임차인 lock-in 36개월(~2025-04-12), 이후 3개월 통지로 해지 가능(12(b)). 임대인은 위반 외 해지 불가(12(d)). "
           "갱신은 상호 합의 신규 계약(13) → 갱신옵션 아님. 해지 미행사 예상으로 9년 전체 반영",
         Q="있음(조항 24 Reinstatement, 12(a)(b))", R="미계상", S="추정 미실시",
         T="추정 미실시 — AS 29상으로도 계상 대상, 원상복구 견적 후 계상 예정",
         U="AS 19 정액 인식 적용(108개월, 무상 4개월 7일·인상분 균등화). 리스 템플릿은 '해당없음' 표기로 제출 → "
           "IFRS 16 미반영(2_IFRS조정 3번). 무이자 보증금 ₹87,32,922(10(a), 3년차 임대료 6개월분). 관리비는 별도 "
           "유지관리 계약. 인테리어 원가 ₹4,78,38,399"),
    dict(C="주재원 사택 1", D="사택", E="2024-08(첫 임대료 월)", H="2026-12(스케줄)", J=_NO_TPL,
         M="2025-07 월 ₹3,23,000→₹3,25,000", P="-", Q="주거용(경미)", R="미계상", S="-",
         T="주거용, 원상복구 비용 경미 예상",
         U="기간은 리스료 스케줄의 첫·마지막 임대료 월(IFRS 16 미반영으로 계약일 불요)"),
    dict(C="주재원 사택 2", D="사택", E="2025-02(첫 임대료 월)", H="2025-12(종료)", J=_NO_TPL,
         M="조정 없음(월 ₹3,00,000)", P="-", Q="-", R="해당없음", S="-", T="-", U="종료 계약. FY25-26 비용만 해당"),
    dict(C="주재원 사택 3", D="사택", E="2024-04(첫 임대료 월)", H="2026-03(종료)", J=_NO_TPL,
         M="조정 없음(월 ₹2,99,000)", P="-", Q="-", R="해당없음", S="-", T="-", U="종료 계약. FY25-26 비용만 해당"),
    dict(C="주재원 사택 4", D="사택", E="2024-03(첫 임대료 월)", H="2026-11(스케줄)", J=_NO_TPL,
         M="조정 없음(월 ₹2,94,000)", P="-", Q="주거용(경미)", R="미계상", S="-",
         T="주거용, 원상복구 비용 경미 예상", U="기간은 리스료 스케줄의 첫·마지막 임대료 월"),
    dict(C="주재원 사택 5", D="사택", E="2026-02(첫 임대료 월)", H="2026-12(스케줄)", J=_NO_TPL,
         M="조정 없음(월 ₹3,83,500)", P="-", Q="주거용(경미)", R="미계상", S="-",
         T="주거용, 원상복구 비용 경미 예상", U="기간은 리스료 스케줄의 첫·마지막 임대료 월"),
    dict(C="관리비(SUNDRY OFFICE RENTAL EXPENSE 52032261)", D="관리비", E="-", H="-", J="해당없음", K="해당없음",
         P="-", Q="-", R="해당없음", S="-", T="-",
         U="리스 아님: 공용공간·사무소 관리 용역(전기료 제외). IFRS 16 비리스 요소로 비용 처리, 조정 불요. "
           "FY25-26 ₹30,08,161. 차세대는 계정명 'RENTAL'을 관리비 계정으로 분리 권고(관리비 예치금 10490700200)"),
]

# ───────────────────────── 13_기타 ─────────────────────────
# (관련 계정, 내용, 현지 처리, IFRS 예상 처리, 금액, 처리 방안, 비고)
OTHER = [
    ("대출 이자수익·미수이자", "NPA 이자 인식",
     "RBI IRACP: 90일 연체 등으로 NPA 분류 시 미수이자 환입, 이후 현금주의",
     "IFRS 9 Stage 3: 손상 후 순장부금액 × EIR로 이자 계속 인식", "0(2026-03-31 NPA 없음)",
     "발생 시 본점 조정. GBS가 비인식 이자를 메모 속성으로 보관", "차이 구조만 존재, 현재 금액 없음"),
    ("지급보증·미사용 약정(부외)", "금융보증·대출약정 손실충당금",
     "Sch.12 우발채무 공시(보증 ₹56,95,709). 보증 손실충당금 0(정상). 미사용 약정 없음(한도 전액 실행)",
     "IFRS 9 금융보증: 최초 수수료로 인식 후 ECL과 미상각 잔액 중 큰 금액",
     "보증잔액 5,695,709",
     "본점 ECL 범위에 보증 포함. 차세대 ECL 범위에 부외(보증·약정) 포함(결정 L6), CCF 속성 관리",
     "미사용 약정이 없어 약정 ECL 해당 없음"),
    ("법정준비금(차세대 31110300000)", "법정준비금 적립",
     "은행법 s.17: 순이익 25% 법정준비금 전입(손익 처분). FY25-26 ₹54,20,250",
     "자본 내 이익잉여금 대체, 손익 영향 없음", 5420250,
     "본점 합산 시 이익잉여금으로 통합 표시", "표시 차이만"),
    ("전기 항목(차세대 40890000000·51208000000)", "전기오류 처리",
     "AS 5: 당기 손익에 'Prior period items' 별도 표시",
     "IAS 8: 중요한 오류는 비교표시 재작성(소급)", "-",
     "중요한 건 발생 시 본점 비교표시 재작성", "-"),
    ("임차료 52032421", "AS 19 임대료 정액 인식",
     "리스기간(2022-04-13~2031-04-12) 정액 비용 인식 적용(무상 4개월 7일·3년 주기 인상분 균등화)",
     "IFRS 16 적용 시 사용권자산 상각·이자로 대체", "-",
     "현지 GAAP 차이 없음. IFRS 16 조정 시 정액 균등화 잔액은 사용권자산과 상계. 차세대 리스 원장이 정액 스케줄 산출",
     "FY25-26 임차료 ₹3,88,82,301"),
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
    10: ("X", "노이다: 리스 템플릿을 '해당없음'으로 표기해 제출(IFRS 16 계산 없음) — 기간 불일치는 없으나 "
              "IFRS 16 미반영 자체가 차이(2_IFRS조정 3번)"),
    11: ("X", "노이다: 템플릿 '해당없음'. 계약 인상 조항은 AS 19 정액 인식에 반영"),
    12: ("O", "노이다: 원상복구 의무 있음, 복구충당부채 미계상 → 견적 후 AS 29로 계상 예정"),
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
    return isinstance(v, str) and any(k in v for k in ("확인 필요", "To be confirmed")) or v in ("확인 중", "Checking")


def _ko():
    g = globals()
    c = types.SimpleNamespace(**{k: g[k] for k in (
        "BR", "SUMMARY", "IFRS_ADJ", "FIN", "COMM", "LEASE", "OTHER", "NY_ISSUE")})
    c.NA, c.COMM_PREFIX, c.CHECKING = "해당없음", "FY25-26 실적.", "확인 중"
    return c


def _en():
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import ifrs_gap_en as c
    c.NA = "N/A"
    return c


def _translate(wb, c):
    """원본 양식 문구·드롭다운·수식을 영문으로 바꾸고 시트명을 바꾼다. 남은 한글이 있으면 실패."""
    for ws in wb:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str):
                    continue
                if v in c.TPL:
                    cell.value = c.TPL[v]
                elif v.startswith("="):
                    for a, b in c.FORMULA_SUBS:
                        v = v.replace(a, b)
                    cell.value = v
        for dv in ws.data_validations.dataValidation:
            if dv.formula1 in c.TPL:
                dv.formula1 = c.TPL[dv.formula1]
    for ws in wb:
        ws.title = c.SHEETS[ws.title]


def _review_sheet(wb, c):
    ws = wb.create_sheet("Review_Request", 0)
    head_src = wb["Summary"]["A4"]
    body_src = wb["Summary"]["I8"]
    ws["A1"] = c.REVIEW_TITLE
    ws["A1"].font = copy(wb["Instructions"]["A1"].font)
    ws.merge_cells("A2:F2")
    ws["A2"] = c.REVIEW_INTRO
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["A2"].font = copy(body_src.font)
    ws.row_dimensions[2].height = 80
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    for j, h in enumerate(c.REVIEW_HEAD, start=1):
        cell = ws.cell(4, j, h)
        for attr in ("font", "fill", "border", "alignment"):
            setattr(cell, attr, copy(getattr(head_src, attr)))
    for i, item in enumerate(c.REVIEW, start=1):
        r = 4 + i
        for j, v in enumerate((i,) + item + ("", ""), start=1):
            cell = ws.cell(r, j, v)
            for attr in ("font", "border", "alignment"):
                setattr(cell, attr, copy(getattr(body_src, attr)))
            if j >= 5:
                cell.fill = CHECK
    for col, w in zip("ABCDEF", (6, 26, 60, 44, 40, 16)):
        ws.column_dimensions[col].width = w
    for r in range(5, 5 + len(c.REVIEW)):
        _fit(ws, r, min_h=30)
    done = getattr(c, "REVIEW_DONE", [])
    if done:
        r = 6 + len(c.REVIEW)
        ws.cell(r, 1, c.REVIEW_DONE_TITLE).font = copy(head_src.font)
        ws.cell(r, 1).fill = copy(head_src.fill)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        for k, (topic, text) in enumerate(done, start=1):
            rr = r + k
            ws.cell(rr, 1, k)
            ws.cell(rr, 2, topic)
            ws.cell(rr, 3, text)
            ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=6)
            for j in range(1, 7):
                cell = ws.cell(rr, j)
                for attr in ("font", "border", "alignment"):
                    setattr(cell, attr, copy(getattr(body_src, attr)))
            ws.row_dimensions[rr].height = max(24, _lines(text, 160) * 12.5 + 6)
    ws.freeze_panes = "A5"
    wb.active = 0


def build(lang="ko", src=SRC, out=None):
    c = _en() if lang == "en" else _ko()
    out = out or (OUT_EN if lang == "en" else OUT)
    BR = c.BR
    wb = openpyxl.load_workbook(src)

    ws = wb["총괄표"]
    ws.column_dimensions["I"].width = 42
    for r, v in c.SUMMARY.items():
        _put(ws, r, 9, v, check=r in SUMMARY_CHECK)
    for r in range(5, 26):
        _fit(ws, r, min_h=ws.row_dimensions[r].height or 24)

    ws = wb["2_IFRS조정"]
    for col, w in (("E", 34), ("H", 16), ("J", 44)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for item in c.IFRS_ADJ:
        _put(ws, r, 2, BR)
        for j, v in enumerate(item, start=3):
            _put(ws, r, j, v)
        _fit(ws, r)
        r += 1

    ws = wb["3_금융상품분류"]
    for col, w in (("C", 12), ("D", 30), ("E", 18), ("M", 40)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for item in c.FIN:
        _put(ws, r, 2, BR)
        for j, v in enumerate(item, start=3):
            _put(ws, r, j, v)
        _fit(ws, r)
        r += 1

    ws = wb["4-7_COMMISSION"]
    for col, w in (("D", 34), ("E", 26), ("I", 14), ("N", 44)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for code, name, nature, prod, gi, tb, amt, note, *ov in c.COMM:
        ov = ov[0] if ov else {}
        _put(ws, r, 2, BR)
        _put(ws, r, 3, code)
        _put(ws, r, 4, name)
        _put(ws, r, 5, nature)
        _put(ws, r, 6, prod)
        if gi:
            _put(ws, r, 7, gi, check=ov.get("check", False))
            _put(ws, r, 8, ov.get("H", c.NA))
            _put(ws, r, 9, tb)
            _put(ws, r, 10, ov.get("J", c.NA))
            _put(ws, r, 11, ov.get("K", "-"))
        _put(ws, r, 12, amt)
        _put(ws, r, 13, ov.get("M", 0))
        _put(ws, r, 14, f"{c.COMM_PREFIX} {note}")
        _fit(ws, r)
        r += 1

    ws = wb["8-12_리스·복구"]
    for col, w in (("H", 18), ("I", 14), ("K", 16), ("P", 14), ("Q", 14), ("U", 40)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for item in c.LEASE:
        _put(ws, r, 2, BR)
        for col, v in item.items():
            _put(ws, r, openpyxl.utils.column_index_from_string(col), v, check=_needs_check(v))
        _fit(ws, r)
        r += 1

    ws = wb["13_기타"]
    for col, w in (("E", 34), ("F", 30), ("H", 34), ("I", 20)):
        ws.column_dimensions[col].width = w
    r = _first_empty(ws)
    for item in c.OTHER:
        _put(ws, r, 2, BR)
        for j, v in enumerate(item, start=3):
            _put(ws, r, j, v, check=_needs_check(v))
        _fit(ws, r)
        r += 1

    ws = wb["뉴욕지점_이슈처리"]
    for r, (mark, note) in c.NY_ISSUE.items():
        _put(ws, r, 7, mark, check=mark == c.CHECKING)
        prev = ws.cell(r, 15).value
        _put(ws, r, 15, f"{prev}\n{note}" if prev else note)
        _fit(ws, r, min_h=ws.row_dimensions[r].height or 42)

    if lang == "en":
        _translate(wb, c)
        _review_sheet(wb, c)
        left = [(ws.title, cell.coordinate) for ws in wb for row in ws.iter_rows() for cell in row
                if isinstance(cell.value, str) and HANGUL.search(cell.value)]
        left += [(ws.title, str(dv.sqref)) for ws in wb for dv in ws.data_validations.dataValidation
                 if dv.formula1 and HANGUL.search(dv.formula1)]
        assert not left, f"untranslated: {left[:10]}"

    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    print(build(*sys.argv[1:2]))
