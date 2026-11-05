# -*- coding: utf-8 -*-
"""시트별 변경 근거 — 본점 분석자료(원안) / 우리측 변경 / 이유·근거 / FY2025-26 결산 증거.

결산 증거 원천: 20260609_Financials_NH_Bank_FY_2025-26.xlsx (노이다 FY2025-26 재무제표 작업파일)
- Grouping: GBS 계정 → Third Schedule 줄 (감사인 결산 grouping)
- Adjustment Entries: 결산 수정분개 (GBS 장부 → 재무제표)
- Law Charges & Insurance: 감사보수·법률비용 수작업 재분류 내역
금액은 FY2025-26(2026-03-31) 기준, INR.
"""

HO = "본점 분석자료(원안)"
OURS = "우리측 변경"
WHY = "이유·근거"
EVID = "FY25-26 결산 증거"

SHEET_DOC = {
    "01 판정기준": [
        (HO, "0930 통합 CoA는 본점(K-IFRS) 기준 구조. 인도 재무제표(Third Schedule) 줄과의 대응표가 없음"),
        (OURS, "GL을 나누는 조건(P1~P3)과 속성으로 두는 조건(P4~P8)을 정하고 Sch.1~16 각 줄을 CoA에 대응"),
        (WHY, "인도 법정 재무제표는 BR Act s.29 Third Schedule 서식. 이 기준이 있어야 계정 신설·속성 판단이 일관됨"),
        (EVID, "결산 grouping 시트가 GBS 계정을 Sch.1~16 줄로 수작업 배정 중 — 계정 단계에서 줄이 정해지면 수작업 제거"),
    ],
    "02 GBS코드충돌": [
        (HO, "0930 'GBS코드' 열(11자리 ↔ 8자리 대응). 예금명세서는 현행 인도 8자리 코드 사용"),
        (OURS, "두 코드를 자동 대조해 13건 의미 충돌·46건 미존재 표시. 2차에서 현행 BS/PL로 0930 GBS열의 오류임을 확인(시트 09)"),
        (WHY, "기준 코드가 11자리로 결정(D1)되어 8자리는 이행 원천으로만 사용 — 0930 GBS열이 틀리면 이행이 틀어짐"),
        (EVID, "결산 grouping의 GBS 계정명(Saving Account (Resident), Fixed Deposit (Resident) 등)이 예금명세서 코드와 일치"),
    ],
    "03 계정검토의견": [
        (HO, "0930 계정 1,168개(현재 명칭 열). 이전 반영안(신규 43·변경 23·비활성 16)"),
        (OURS, "신규·변경·이전안수정·속성전환·비활성·철회·매핑을 계정별로 표시. 차수별 변경은 '이전안 관계' 열"),
        (WHY, "인도 Schedule 줄·RBI 보고·FEMA 상품 구분에 필요한 계정만 신설하고 나머지는 속성(시트 01 원칙)"),
        (EVID, "Sch.1 'Deposit with RBI' ₹15 lakh = Statutory Reserve Account(RBI) → s.11(2)(b) 예치 계정(10104010200) 확인. "
               "Sch.7 I(인도 내)에 해외 Nostro ₹2.01 crore 포함 → 인도 내/외 계정 분리 필요"),
    ],
    "04 예금상품매핑": [
        (HO, "0930 예금 L4 = 고객유형(법인·은행·개인), 저축·적금 구분 없음. 예금명세서 상품 35개(8자리)"),
        (OURS, "상품 → 11자리 L5. 저축 L5 분리, 적금은 정기예금과 동등한 L3(20190000000) 신설 요청"),
        (WHY, "Sch.3은 요구불·저축·정기 + 은행분·기타분. Form A는 저축을 요구불/정기로 다시 나눔 → L5 순도 필요"),
        (EVID, "Sch.3 II 저축 ₹78.7 lakh, III에 적금(Recurring Deposit (Resident)) ₹50,000 실재 → 적금 계정 필요"),
    ],
    "05 RBI보고속성": [
        (HO, "0930에는 속성 정의 없음. 예금명세서 시트 16~18·23에 일부 태그"),
        (OURS, "RBI·FEMA·세무·ECL·자본 보고용 속성 72개를 수준(고객·계좌·거래)과 함께 정의"),
        (WHY, "PSL·BSR·CRILC·LCR 분류는 Schedule 줄이 아니고 수시로 바뀜 → GL이 아닌 속성(P4)"),
        (EVID, "결산 작업파일의 Asset Quality·Concentration·ALM·Top 20 시트를 매번 수작업 집계 — 속성이 있으면 자동 산출"),
    ],
    "06 보고서별요구": [
        (HO, "없음"), (OURS, "보고서 19종별 출처 계정·필요 속성·산출 방식"),
        (WHY, "계정과 속성이 보고서를 실제로 만들 수 있는지 역으로 검증"), (EVID, "Ratios(Form X)·ALM 시트가 수작업 → 이 표의 속성으로 대체"),
    ],
    "07 FN·PSL·Nostro": [
        (HO, "0930: Nostro 10105010000(DOMESTIC)/10105020000(OVERSEA). PSL·NRO 구분 없음"),
        (OURS, "FN은 NRO 부분집합(속성), PSL은 속성, 인도 내 INR 결제계좌 L5 신설"),
        (WHY, "FEMA상 외국인 비거주자 INR 예금은 NRO. Sch.7은 인도 내/외로 나눔"),
        (EVID, "Sch.7 I(인도 내)에 'Due from Oversea Banks on Demand' ₹2.01 crore가 들어가 있음 → 해외 Nostro는 Sch.7 II(인도 외)로 가야 함"),
    ],
    "09 현행GBS이행매핑": [
        (HO, "0930 GBS코드 열이 가리키는 계정('0930 자동대응' 열)"),
        (OURS, "현행 leaf 825개의 이행 대상(11자리)과 판정. 이행 후 B/S·P/L 대차 0 검증"),
        (WHY, "0930 GBS열 오류 57건(저축→법인 요구불, RD→AOB, GST→원천세, 투자 분류 순서 대응 등)"),
        (EVID, "Sch.13 iii에 은행 정기예치 이자 ₹7.93 crore(은행간 자금) → 0930의 CD 매입 대응은 틀림. 결산 수정분개에서 HTM 미수이자 ₹10.2 lakh를 임시코드로 처리 → A.I.R ON STATE TREASURY SECURITIES(10410110000) 사용"),
    ],
    "10 ECL계정": [
        (HO, "ECL 개발명세(8자리 10690003~5, 2096 그룹 등). 0930 부외 충당부채 20910000000~20930000000"),
        (OURS, "Stage별 L5 3개·전환조정 자본·EIR 조정만 신설, 부외 ECL은 0930 기존 계정 사용"),
        (WHY, "ECL Directions 2026(2027-04-01). 명세서의 8자리 10690003~5는 현행 장부에서 사용 중"),
        (EVID, "현재 정상자산 충당금 ₹1.53 crore는 Sch.5 '기타(충당금 포함)'에 표시(IRAC). ECL 전환 후 Stage 충당금으로"),
    ],
    "11 LCR분류·파라미터": [
        (HO, "0930: 예금 L4 고객유형(법인·은행·개인)만"),
        (OURS, "LCR 상대방 분류 11종·유출입률·HQLA·파라미터 테이블. 계정은 은행/비은행까지만"),
        (WHY, "LCR 분류는 그룹 합산·부보·운영예금 여부로 보고일마다 바뀜 → 속성 + 일자 유효 파라미터"),
        (EVID, "ALM 시트를 수작업 작성 중"),
    ],
    "12 세무계정(TDS·TCS·GST)": [
        (HO, "0930: TAX WITHHOLDINGS(20510400000)·PAYABLE VAT(20510500000) L4 두 개뿐. 현행 GBS: 조문별 TDS 30여 개·GST 세목/환전/RCM 계정. TDS 산출물·GST FRS"),
        (OURS, "TDS·TCS: L5 신고서 × L6 납부 성격 21개(TCS 기타 포함). GST: L5 세목 × L6 수수료/환전, RCM·ITC·임시 ITC·전자현금원장·이자가산세. 연말 미경과 이자 TDS 순액법"),
        (WHY, "챌린·신고서 대사 단위가 계정과 맞아야 함(P2). 조문번호 대신 소득 성격 명칭 → 신 소득세법(392·393·394) 조문 변경에도 GL 불변"),
        (EVID, "Sch.5에 TDS 11개·GST 9개 계정을 줄별로 나열. Prepaid TDS on Accrued Interest −₹1.39 lakh를 미지급이자에서 차감 표시(순액법과 같은 결과). 급여 TDS ₹6.98 lakh 결산 수정"),
    ],
    "13 고정자산": [
        (HO, "0930: 기계장치·공구기구·기타·임차개량(가구비품·컴퓨터 구분 없음). FA 코드표: GL 연결 열 공란, 내용연수·상각률 0"),
        (OURS, "Schedule II 분류(단말 3년·서버 6년·전기설비 10년·가구비품 10년) 계정 신설, FA 코드 41개에 GL·내용연수·세무율 연결, 처분손익 계정"),
        (WHY, "인도 장부는 AS + Companies Act Schedule II 내용연수, 세무는 블록 WDV. 은행은 Ind AS 미적용 → ROU는 본점 IFRS 조정용"),
        (EVID, "매년 CBS 감가상각을 전액 되돌리고 FAR 기준으로 재계상(FY25-26 되돌림 ₹2.82 crore·FAR 상각 ₹2.97 crore). GST ITC가 Prepaid VAT에 잘못 기장돼 자산원가로 재분류. 이연법인세(AS 22) ₹37.6 lakh 수작업"),
    ],
    "14 영업비용·Sch16": [
        (HO, "0930 영업비용 L4/L5/L6 그룹(복리비가 사무경비 아래 등). 경비(EP) 코드 103개"),
        (OURS, "계정별 Sch.16 줄 지정, 감사보수·법률·전문가 수수료 L6(SERVICE FEES 아래)·보험료·협회비·CSR L5 신설, 인건비 법정항목·직원 월적립 부채 신설"),
        (WHY, "Sch.16 VII(감사보수)·VIII(법률)은 별도 줄 — GL에서 바로 나와야 함. '기타비용' 중 총수익 1% 초과 항목 별도 공시"),
        (EVID, "감사보수 ₹6.67 lakh·법률비용 ₹13.77 lakh가 Service Commission·GST 불공제에 기장된 뒤 거래별 수작업 재분류(Law Charges & Insurance 시트). "
               "Gratuity는 '복리비(기타인건비성)' ₹4.37 lakh를 보험수리 보고서로 확인. Service Commission ₹3.36 crore = 총수익 ₹29.5 crore의 11%"),
    ],
    "15 송금·무역수수료": [
        (HO, "0930: 송금 L4(당발·타발·전신료·OUR·기타)와 상품별 전신료 L4"),
        (OURS, "중계은행 수수료 고객 회수분 = 과세 수익 40570110000(결정 C1). 취급·SWIFT 전신료는 0930 L4 그대로"),
        (WHY, "받는 금액과 실제 중계은행 차감액이 달라 pure agent 아님 → GST 18% 과세. 정액 OUR와 구분해 마진 관리"),
        (EVID, "Sch.14 송금수수료 중 'Commission - Others(Remittance)' ₹11.26 lakh가 가장 큼(당발 수수료 ₹1.93 lakh의 6배) → 회수분 성격 확인 후 40570110000으로"),
    ],
    "08 결정사항": [
        (HO, "본점 CoA 담당·HQ IT 결정 대상"), (OURS, "결정/권고/미결/확인 요청 상태 관리"),
        (WHY, "결정된 항목만 coa_model.py·분개·DDL에 반영"), (EVID, "-"),
    ],
}

# 결산 증거 → 변경 연결 (결산 자료, 내용, 금액(INR), 시사점, 관련 변경, 시트)
EVIDENCE = [
    ("Grouping Sch.1", "Deposit with RBI(= Statutory Reserve Account RBI)", 1500000, "s.11(2)(b) 예치가 Sch.1 주기 대상", "10104010200 RBI 기타계좌 + encumbrance=S11_2", "03·09"),
    ("Grouping Sch.3", "적금(Recurring Deposit (Resident)) 잔액", 50000, "적금 상품 실재", "적금 L3 20190000000 신설 요청", "04"),
    ("Grouping Sch.5 I", "Prepaid TDS on Accrued Interest를 미지급이자에서 차감", -138767, "연말 미경과 이자 TDS는 결과적으로 순액 표시", "순액법(T4), 10920014는 이행 시 AIP와 상계", "12"),
    ("Grouping Sch.5 III", "정상자산 충당금(Allowance General) Sch.5 표시", 15270000, "현행 IRAC 표시", "ECL Stage 충당금(10380103000~105000)", "10"),
    ("Grouping Sch.6 ii", "콜론(Call Loans - Domestic)을 RBI 기타계좌에 포함", 100000000, "콜론은 Sch.7 I(ii)(a) Money at call", "10107010100 + Sch 줄 매핑", "09"),
    ("Grouping Sch.7 I", "해외 Nostro(Due from Oversea Banks)를 '인도 내'에 포함", 20147539, "Nostro는 Sch.7 II(인도 외)", "10105020000(인도 외)·10105010100(인도 내 INR) 분리", "07"),
    ("Grouping Sch.13 iii", "은행 정기예치 이자를 '은행간 자금 이자'로", 79280961, "은행 정기예치는 Sch.7 '기타예금' — CD 아님", "10106010000·40120100000 (E1)", "09"),
    ("Grouping Sch.14", "Commission - Others(Remittance)", 1126026, "송금수수료 중 최대, 성격 불명확", "중계은행 회수분 40570110000(C1)", "15"),
    ("Grouping Sch.16 VII", "감사보수 = Service Commission 기장 ₹6,51,300 + GST 불공제 ₹15,525", 666825, "계정 없어 거래별 수작업 재분류", "50940021803 AUDITORS' FEES", "14"),
    ("Grouping Sch.16 VIII", "법률비용 = Service Commission ₹12,63,000 + GST 불공제 ₹1,13,670", 1376670, "위와 같음", "50940021804 LAW CHARGES", "14"),
    ("Grouping Sch.16 I", "Gratuity = 복리비(기타인건비성), 보험수리 보고서로 확인", 437408, "Gratuity 비용·충당부채 계정 없음", "50930170100 GRATUITY·20510300400 충당부채", "14"),
    ("Grouping Sch.16 XII", "Service Commission", 33612072, "총수익의 11% → 1% 초과 별도 공시", "외주용역 L6 50940021802 + 공시", "14"),
    ("Grouping Sch.16 XII", "GST 불공제(6개 계정) 합계", 10068090, "XII 표시", "50940010300 + Sch.16 XII", "14"),
    ("Grouping Sch.16", "교육훈련 I · 사무용품 XII · 차량유지 XII · 업무추진비 IV", None, "결산 grouping 실무", "SCH16 매핑을 실무와 일치", "14"),
    ("Adjustment FY25-26", "CBS 감가상각 되돌림(기계·기타·SW·임차개량)", 28178100, "CBS 상각 로직이 인도 기준과 다름", "Schedule II 분류·FA 코드 GL 연결(시트 13)", "13"),
    ("Adjustment FY25-26", "FAR 기준 감가상각 재계상", 29710716, "FAR이 실질 원장", "FA 코드표 내용연수·상각률 채움", "13"),
    ("Adjustment FY25-26", "이연법인세(AS 22) 환입", 3757349, "수작업", "51330000000·20720100000 + 세무상각(블록) 속성", "13"),
    ("Adjustment FY23-24", "고정자산 GST ITC를 Prepaid VAT에서 자산원가로", 35474, "ITC 50% 자본화 규칙 미반영", "GST_ITC_ON_FA 파라미터", "13"),
    ("Adjustment FY25-26", "HTM 국채 미수이자(임시코드 1001)", 1017940, "투자 미수이자 계정 미사용", "10410110000 A.I.R ON STATE TREASURY SECURITIES", "09"),
    ("Adjustment FY25-26", "급여 TDS 결산 수정", 698360, "급여 TDS 정산", "20510400401 TDS 급여", "12"),
]
