# -*- coding: utf-8 -*-
"""통합 표준 CoA(0930) 검토의견 — 인도(Noida) 지점, RBI 보고 관점.

입력
- 통합 표준 CoA 0930 (tools/data/coa_base.json) + 이전 반영안(coa_model.py: ADD/MOD/DEACT)
- GBS Deposit Product Spec (ACPIR 2026 v3, 시트 02·03·04·06·16·17·18·23·32·37·47·49)

판정 원칙 (P)
- GL은 "재무제표 Schedule 줄"과 "잔액 대사 단위"가 갈라질 때만 나눈다.
- 그 밖의 RBI 보고 구분(PSL, BSR 코드, CRILC, LCR 분류, 거주성 세분 등)은 계좌·고객·거래 속성으로 관리하고
  보고 시점에 GL × 속성으로 산출한다.

이 파일은 데이터만 담는다. build_coa_review.py가 Excel을 만든다.
VERIFY='Y'는 규정 조문·서식·주기를 컴플라이언스·재무가 확인해야 하는 항목.
"""
import coa_model as CM

# ---------------------------------------------------------------------------
# 0. 판정 원칙
# ---------------------------------------------------------------------------
PRINCIPLES = [
    ("P1", "GL 분리 조건 ①", "재무제표 Schedule(Third Schedule, BR Act s.29) 표시 줄이 다르면 GL 분리",
     "예: Sch.3 요구불·저축·정기, 은행분·기타분 / Sch.7 인도 내·외 / Sch.6 현금·RBI", "GL"),
    ("P2", "GL 분리 조건 ②", "잔액을 따로 대사·정산해야 하면 GL 분리 (세목별 납부, 결제망별 정산, 제3자 자금)",
     "예: CGST/SGST/IGST, TDS 26Q·27Q, 에스크로 유형별 일일대사, 본점 INR Vostro", "GL"),
    ("P3", "GL 분리 조건 ③", "FEMA상 상품 자체가 다른 계좌(송금성·과세·금리규제가 다름)는 상품 L5로 분리",
     "예: NRE·NRO·FCNR(B)·SNRR·EEFC·RFC, LO/BO/PO", "GL"),
    ("P4", "속성으로 관리", "같은 Schedule 줄 안의 규제 세분, 분기·연간 보고 분류, 시간에 따라 바뀌는 분류는 속성",
     "예: PSL 분류, MSME 규모, BSR 활동·부문 코드, SMA·Stage, LCR 예금자 분류, DICGC 부보, 국가코드", "속성"),
    ("P5", "통화는 GL이 아님", "통화는 GBS 잔액 차원(통화코드)으로 관리. 통화 때문에 GL을 늘리지 않는다",
     "예: FCNR 미지급이자 FCY 전용 GL(20910012*) 불필요, INR Nostro는 통화+속성", "통화차원"),
    ("P6", "Schedule 줄이 속성에서 나오는 경우", "Sch.9 B(담보구분)·C(우선부문/공공/은행/기타)처럼 B/S 줄이라도 "
     "같은 대출이 시점에 따라 줄을 옮기면 GL로 고정하지 않고 속성에서 산출", "PSL·담보는 재분류가 잦아 GL 이동 시 이자·충당금 계정까지 따라 움직여야 함", "속성"),
    ("P7", "L5 순도", "Schedule 매핑은 L5 단위로 한다. L5 하나에 두 Schedule 줄(예: 저축+당좌)이 섞이면 안 됨",
     "이전안 20150300100 'SB/CA'가 위반 → 분리", "GL"),
    ("P8", "통합 CoA 존중", "본점 통합 L1~L4는 바꾸지 않는다. 인도 구분은 L5 신설·표시명 변경·비활성으로만 반영",
     "통합 L4가 고객유형(법인·은행·개인) 축이므로 예금 GL 결정키 = 상품 × 거주성 × 고객유형군", "GL"),
]

# ---------------------------------------------------------------------------
# 1. 인도 재무제표 Schedule ↔ 판정 (GL 필요 줄과 속성 산출 줄)
#    (Schedule, 줄, CoA 대응, 판정, 비고)
# ---------------------------------------------------------------------------
SCHEDULES = [
    ("Sch.1 Capital", "외국은행: s.11(2)(b) RBI 예치액(현금·승인증권) 주기", "31140100000 갑기금", "속성+주석",
     "예치 증권은 투자 GL 그대로, encumbrance=S11_2 속성. SLR 산정에서 제외"),
    ("Sch.2 Reserves", "법정준비금 / 자본준비금 / 투자변동준비금(IFR) / AFS Reserve / 이익잉여금", "31110000000~31130000000",
     "GL", "통합 CoA는 법정준비금이 부채(20510200000)에 있음 → 인도는 자본으로"),
    ("Sch.3 Deposits", "A.I 요구불(은행·기타) / A.II 저축 / A.III 정기(은행·기타)", "20150000000·20160000000 L5", "GL",
     "저축(II)이 통합 CoA에 없음 → L5로 순수 분리. 은행분은 '인도 내 은행'만(Form A banking system)"),
    ("Sch.3 Deposits", "B. 인도 내 지점 예금 / 인도 외 지점 예금", "-", "해당없음", "노이다 지점 예금은 전부 B.(i)"),
    ("Sch.4 Borrowings", "I. 인도 내(RBI·은행·기관) / II. 인도 외", "20220000000·20230000000·20250000000", "GL",
     "Nostro 마이너스 잔액(20140200000)은 월말 '인도 외 차입'으로 표시"),
    ("Sch.5 Other Liabilities", "Bills payable / Inter-office(순액) / 미지급이자 / 기타(충당부채 포함)", "20110200000·21000000000·20410000000·20400000000",
     "GL", "본지점은 순액 표시. Standard asset 충당금 표시 위치는 ACPIR 확인(VERIFY)"),
    ("Sch.6 Cash & RBI", "현금(외화현찰 포함) / RBI 당좌 / RBI 기타계좌", "10101000000~10104010000", "GL",
     "RBI 당좌와 기타계좌 L5 분리"),
    ("Sch.7 Banks & Call", "I. 인도 내: 당좌·기타예치·콜(은행·기타기관) / II. 인도 외: 당좌·기타예치·콜", "10105·10106·10107", "GL(위치) + 속성(통화)",
     "INR·FCY 구분은 통화차원 + nostro_type 속성. 인도 내/외는 기존 L4(Domestic/Oversea)"),
    ("Sch.8 Investments", "인도 내: 국채·승인증권·주식·사채·자회사·기타 / 인도 외", "10200000000", "속성",
     "IFRS 분류(AC·FVOCI·FVPL) GL 유지, RBI 분류(HTM·AFS·FVTPL/HFT)·발행자유형·SLR 여부는 속성"),
    ("Sch.9 Advances A", "(i) 매입·할인어음 (ii) CC·OD·요구불대출 (iii) 기한부대출", "10310·10320·10330·10340", "GL",
     "CC/OD 자산계정이 통합 CoA에 없음 → 신설. Packing Credit 신설"),
    ("Sch.9 Advances B", "(i) 유형자산 담보 (ii) 은행·정부 보증 (iii) 무담보", "-", "속성", "security_class 속성(담보 재평가로 수시 변동)"),
    ("Sch.9 Advances C", "I. 인도 내: 우선부문(PSL)·공공부문·은행·기타 / II. 인도 외", "-", "속성",
     "PSL은 B/S 줄이지만 분기마다 재판정 → psl_flag·psl_category 속성으로 산출"),
    ("Sch.10 Fixed Assets", "건물 / 기타 / 리스자산", "10500000000~10700000000", "GL", "ROU 자산 표시는 Ind AS 미적용 은행 형식 확인(VERIFY)"),
    ("Sch.11 Other Assets", "본지점 순액 / 미수이자 / 선납세금 / 기타(PSL 미달예치·DEAF 청구 등)", "10400000000", "GL",
     "PSL 미달예치금(NABARD·SIDBI 등), DEAF 청구 미수 L4 신설"),
    ("Sch.12 Contingent", "선물환 / 보증(인도 내·외) / 인수·배서 / 기타", "60100000000", "GL(종류) + 속성(인도 내·외, 금융·이행)",
     "인도 내·외는 수익자 위치 속성"),
    ("Sch.13 Interest Earned", "대출·어음 이자 / 투자수익 / RBI·은행간 자금 이자 / 기타", "40100000000~40350000000", "GL", ""),
    ("Sch.14 Other Income", "수수료·환·중개 / 투자매매익 / 재평가익 / 외환매매익 / 배당 / 기타", "40500000000·40700000000", "GL",
     "연체 Penal Charges는 이자가 아닌 수수료(40580030000)"),
    ("Sch.15 Interest Expended", "예금이자 / RBI·은행간 차입이자 / 기타", "50100000000~50300000000", "GL",
     "예금이자는 통합 L4(고객유형) 유지. 거주성·상품은 속성 (51040011* 등 신설 불필요)"),
    ("Sch.16 Operating Exp.", "인건비 / 임차료·세금 / 인쇄 / 광고 / 감가상각 / 감사보수 / 법률 / 통신 / 수선 / 보험 / 기타", "50900000000", "GL",
     "운영비는 본 검토 범위 밖(별도 확인)"),
]

# ---------------------------------------------------------------------------
# 2. 예금 명세서의 '현행 GBS 코드' (8자리) — 0930 CoA의 GBS코드 열과 충돌 여부를 자동 비교
#    (spec 코드, spec 의미, spec 상태)
# ---------------------------------------------------------------------------
SPEC_CODES = [
    ("20121001", "Savings (Resident)", "LIVE"), ("20121002", "Savings (Non-Resident Foreign National)", "Existing"),
    ("20121011", "Savings (NRE)", "NEW*"), ("20121012", "Savings (NRO)", "NEW*"),
    ("20122001", "Term Deposit (Resident)", "LIVE"), ("20122002", "Term (Non-Resident Foreign National)", "Existing"),
    ("20122011", "Term (NRE)", "NEW*"), ("20122012", "Term (NRO)", "NEW*"), ("20122021", "FCNR Term (FCY)", "NEW*"),
    ("20122041", "Term Deposit - Escrow-Linked", "NEW*"), ("20122051", "Term Deposits from Banks", "NEW*"),
    ("20124001", "Recurring Deposit (Resident)", "PARTIAL"), ("20124002", "RD (Non-Resident Foreign National)", "Existing"),
    ("20124011", "RD (NRE)", "NEW*"), ("20124012", "RD (NRO)", "NEW*"),
    ("20125000", "Current Deposit (parent)", "LIVE"), ("20125001", "Current Account (Resident)", "LIVE"),
    ("20125002", "Current (Non-Resident Foreign National)", "Existing"), ("20125003", "EEFC (Resident)", "NEW*"),
    ("20125004", "Collection Account / EEFC non-resident (시트 03·17 이중 사용)", "NEW*"),
    ("20125011", "Current (NRE)", "NEW*"), ("20125012", "Current (NRO)", "NEW*"), ("20125013", "Current (SNRR) - 시트17", "NEW*"),
    ("20125021", "Current (Liaison Office)", "NEW*"), ("20125022", "Current (Project Office)", "NEW*"),
    ("20125023", "Current (Branch Office)", "NEW*"), ("20125031", "FCY (Project Office)", "NEW*"),
    ("20125041", "Escrow Deposits - FEMA", "NEW*"), ("20125042", "Escrow Deposits - Domestic", "NEW*"),
    ("20125043", "Escrow Deposits - PA", "NEW*"), ("20125301", "SNRR - 시트03", "NEW*"),
    ("20910002", "Interest Payable - Demand", "LIVE"), ("20910003", "Interest Payable - Term", "LIVE"),
    ("20910012", "Interest Payable - FCNR (FCY)", "NEW*"), ("20920017", "Unearned Guarantee Fee", "LIVE"),
    ("20990013", "GST Output (CGST)", "LIVE"), ("20990014", "GST Output (SGST)", "LIVE"), ("20990015", "GST Output (IGST)", "LIVE"),
    ("20990016", "GST on FX Revenue", "LIVE"), ("20990019", "GST RCM Payable", "LIVE"),
    ("20990022", "TDS on Deposit Interest (Resident)", "LIVE"), ("20990023", "TDS (Non-Resident, s.195)", "NEW*"),
    ("20990041", "Unidentified Escrow Credits", "NEW*"),
    ("10990013", "GST Input Credit", "LIVE"), ("10990019", "GST RCM Input Credit", "LIVE"),
    ("51040001", "Savings Interest (Resident)", "LIVE"), ("51040002", "Savings Interest (NR Foreign National)", "Existing"),
    ("51040011", "Savings Interest (NRE)", "NEW*"), ("51040012", "Savings Interest (NRO)", "NEW*"),
    ("51040041", "Interest Expense - PA Escrow", "NEW*"),
    ("51050001", "RD Interest (Resident)", "LIVE"), ("51050002", "RD Interest (NR FN)", "Existing"),
    ("51060001", "Term Interest (Resident)", "LIVE"), ("51060002", "Term Interest (NR FN)", "Existing"),
    ("51060014", "Term Interest (NRE)", "NEW*"), ("51060015", "Term Interest (NRO)", "NEW*"),
    ("51060016", "FCNR Interest (FCY)", "NEW*"), ("51060051", "Interest on Term Deposits from Banks", "NEW*"),
    ("10610004", "Loan against Bills (Working Capital)", "LIVE"), ("10610007", "Loan against Bills (Term/Equipment)", "LIVE"),
    ("10630001", "Overdraft (new account)", "NEW*"), ("10420002", "Discounted Bills (Domestic BA)", "LIVE"),
    ("10420003", "Discounted Bills (Overseas BA)", "LIVE"),
    ("10910003", "Interest Receivable - Working Capital", "LIVE"), ("10910005", "Interest Receivable - Term", "LIVE"),
    ("10910009", "Interest Receivable - Discounted Bills", "LIVE"),
    ("41020004", "Interest Income - Working Capital", "LIVE"), ("41020007", "Interest Income - Term", "LIVE"),
    ("41050002", "Interest Income - Discounted Bills (Dom)", "LIVE"), ("41050003", "Interest Income - Discounted Bills (Ovs)", "LIVE"),
    ("42010013", "Guarantee Fee Income", "LIVE"), ("42020601", "Remittance Fee Income", "LIVE"),
    ("42020701", "Loan Processing Fee", "LIVE"), ("42020729", "Other Fees (escrow fee 포함)", "LIVE"),
    ("48090002", "FX Revaluation Gain", "LIVE"), ("53090012", "FX Revaluation Loss", "LIVE"),
    ("71040004", "Guarantees Issued", "LIVE"), ("72040001", "Committed Credit Line (명세서는 LC에 사용)", "LIVE"),
]

# ---------------------------------------------------------------------------
# 3. 계정 검토의견 (구분, 코드, Lv, 상위코드, 영문명, 한글명, 사유·근거, 관련 보고, VERIFY, 이전안 관계)
#    구분: 신규 / 변경 / 이전안수정 / 속성전환 / 비활성 / 매핑
# ---------------------------------------------------------------------------
R = []


def rv(kind, code, lv, parent, en, kr, why, rep, verify="N", prev=""):
    R.append(dict(kind=kind, code=code, lv=lv, parent=parent, en=en, kr=kr, why=why, rep=rep, verify=verify, prev=prev))


# ---- 3-1 현금·RBI·Nostro ---------------------------------------------------
rv("신규", "10104010100", 5, "10104010000", "RBI - CURRENT ACCOUNT (CRR / RTGS SETTLEMENT)", "RBI 당좌(CRR·RTGS 결제)",
   "Sch.6 '(i) in current account'. CRR 유지잔액 일일 점검 대상. NEFT/RTGS·세금납부 결제", "Form A, Sch.6, CRR 일일")
rv("신규", "10104010200", 5, "10104010000", "RBI - OTHER ACCOUNTS", "RBI 기타계좌",
   "Sch.6 '(ii) in other accounts'. s.11(2)(b) 현금예치를 현금으로 할 경우 등. SDF·역레포 표시 위치는 확인 필요",
   "Sch.6, Sch.1 주기", "Y")
rv("변경", "10104010000", 4, "10104000000", "DUE FROM RESERVE BANK OF INDIA", "RBI 예치금(L5 집계)",
   "직접 기장 금지 → L5 사용", "", prev="이전안 MOD 유지(표시명 RBI) + 직접기장 금지 추가")
rv("신규", "10105010100", 5, "10105010000", "CURRENT A/C WITH BANKS IN INDIA - INR (SETTLEMENT)", "인도 내 은행 INR 당좌(결제용)",
   "INR 운용 계좌(청산·NACH 스폰서·현금관리 등)를 FCY Nostro와 분리. Form A '은행시스템 자산'(NDTL 차감), NOP 제외, 은행간 익스포저. "
   "통화만으로도 구분되나 하위원장·대사 단위가 달라 L5로 둠 (선택: 속성만으로 운영 가능)", "Form A, Sch.7 I, LEF(은행간)")
rv("신규", "10105010200", 5, "10105010000", "CURRENT A/C WITH BANKS IN INDIA - FCY", "인도 내 은행 외화 당좌",
   "인도 내 AD은행 외화계좌. Sch.7 I(인도 내)이면서 NOP·LBS 대상(외화)", "Sch.7 I, NOP, LBS")
rv("변경", "10105010000", 4, "10105000000", "DUE FROM BANKS IN INDIA ON DEMAND", "인도 내 은행 요구불(L5 집계)",
   "통합명 DOMESTIC = 지점 소재국(인도). 직접 기장 금지 → L5", "Sch.7 I")
rv("변경", "10105020000", 4, "10105000000", "DUE FROM BANKS OUTSIDE INDIA ON DEMAND (NOSTRO)", "해외 Nostro",
   "통화는 통화차원, Nostro별 하위원장(TB_NOSTRO_ACCT.SUB_LEDGER). nostro_type·국가·은행유형 속성 필수(LBS)", "Sch.7 II, NOP, LBS",
   prev="이전안 MOD 유지")
rv("매핑", "20140200000", 4, "20140000000", "DUE FROM OVERSEA BANKS ON DEMAND O/D", "Nostro 당좌차월",
   "Nostro 마이너스 잔액은 월말 Sch.4 II(인도 외 차입)로 재분류 표시. 일중에는 Nostro 계정 그대로, 재분류 배치로 처리", "Sch.4, Form A(차입)")
rv("매핑", "10107090000", 4, "10107000000", "CCP SWAP MARGIN DEPOSIT → CCIL MARGIN / DEFAULT FUND", "CCIL 증거금·결제기금",
   "인도 CCP(CCIL: 외환·TREPS·국채 결제) 증거금을 통합 CCP 계정에 매핑. 신규 계정 불필요, 속성 ccp_segment", "Sch.11 또는 Sch.7(확인), LCR(담보제공)", "Y")

rv("변경", "10106010000", 4, "10106000000", "TERM DEPOSITS PLACED WITH BANKS IN INDIA", "인도 내 은행 정기예치",
   "(3차) 은행 정기예치 ₹208.9 crore 이행 대상. Sch.7 I(i)(b). 은행 상대 예치만 이 계정 — 법인·기타 대상 예치는 Sch.7이 아니고 대출 성격이므로 인도 미사용. "
   "상대은행 유형(SCB·외국은행 지점·협동조합 등)·만기·운영예치 여부는 LCR 유입률·ECL 범주(vi)·RWA(은행 익스포저) 속성", "Sch.7, Form A(은행시스템 자산), LCR 유입, ECL, RWA")
rv("신규", "10107010400", 5, "10107010000", "TREPS / MONEY MARKET LENDING - OTHER INSTITUTIONS (CCIL)", "TREPS 등 기타기관 콜론(선택)",
   "(3차) Sch.7 '콜·단기 — (b) 기타기관'. CCIL TREPS를 운용할 때만 신설", "Sch.7, LCR 유입")
rv("비활성", "10107050100", 5, "10107050000", "CD PURCHASE - DOMESTIC BANKS", "CD 매입(인도 내 은행)",
   "(3차) 본점 매핑은 '외화정기예치금'이나 영문명은 CD. 인도에서 CD 매입은 투자(Sch.8, 발행자=은행, AC/AFS) → 투자 계정 사용, 은행 예치는 10106010000으로 일원화", "Sch.8")
rv("변경", "10810100000", 4, "10810000000", "ADVANCE INCOME TAX (PREPAID)", "선납법인세(Advance Tax)", "(3차) 선납 분기납부분만. 원천징수당한 TDS는 10810200000", "Sch.11")
rv("신규", "10810200000", 4, "10810000000", "TDS RECEIVABLE (TAX DEDUCTED AT SOURCE BY PAYERS)", "TDS 미수(원천징수당한 세액)",
   "(3차) 현행 TDS Receivable·Prepaid TDS 이행. 26AS/AIS 대사 → 법인세 신고 시 공제. TDS 예수금(20510400xxx, 당행이 원천징수한 부채)과 반대 방향", "Sch.11, 법인세 신고")
rv("비활성", "20180000000", 3, "20100000000", "AOB DEPOSITS (NH ALL ONE BANK)", "AOB(NH올원뱅크) 예금",
   "AOB = NH농협 모바일뱅킹 '올원뱅크'(0930 본점 매핑명 '올원뱅크(국외정산대금)'). 인도지점 미취급 → 1차부터 비활성. 현행 RD 계정이 여기에 잘못 대응돼 있음", "")

# ---- 3-2 여신 --------------------------------------------------------------
rv("신규", "10330304000", 5, "10330300000", "CASH CREDIT / OVERDRAFT", "당좌대월(CC/OD)",
   "Sch.9 A(ii). 통합 CoA에 CC/OD 자산계정 없음(이자수익 40110100000만 존재). 예금명세서 LN-01 제안코드 10630001*은 0930 GBS코드에서 "
   "SYN. WC LOAN(LT)와 충돌 → 본 코드로 대체. Out-of-order 3개 테스트(ACPIR §6(17)) 대상", "Sch.9, BSR-1, CRILC, IRAC")
rv("신규", "10330500000", 4, "10330000000", "EXPORT PACKING CREDIT TOTAL", "선적전 수출금융(Packing Credit)",
   "Sch.9 A(ii) 요구불 성격. 수출대금(EDPMS)으로 상환 대사가 필요하고 수출신용 보고·PSL 수출신용 원천 → 상품 GL", "Sch.9, PSL(수출신용), EDPMS")
rv("신규", "10330501000", 5, "10330500000", "RUPEE PACKING CREDIT (INR)", "원화(INR) 패킹크레딧", "", "")
rv("신규", "10330502000", 5, "10330500000", "PACKING CREDIT IN FOREIGN CURRENCY (PCFC)", "외화 패킹크레딧(PCFC)",
   "외화 수출금융. 금리는 ARR 기준(속성). 통화는 통화차원", "")
rv("신규", "40310920000", 4, "40310000000", "INT. ON PACKING CREDIT (INR / PCFC)", "패킹크레딧 이자", "", "Sch.13")
rv("매핑", "40110100000", 4, "40110000000", "INT. ON OVERDRAFT (CUST)", "당좌대월 이자", "CC/OD 이자수익. A.I.R는 10410030000(OPERATION) 사용", "Sch.13")
rv("매핑", "40580030000", 4, "40580000000", "COMMISSION - LOAN LATE PAYMENT CHARGE → PENAL CHARGES", "연체 Penal Charges",
   "RBI Penal Charges 지침: 연체 부과금은 '이자'로 가산·복리 불가, 수수료로만 인식 → 이자 GL 기장 금지 규칙", "Sch.14, 공정대출관행", "Y")
rv("속성전환", "-", "", "", "PSL / MSME / EXPORT CREDIT CLASSIFICATION", "PSL·MSME·수출신용 분류",
   "GL 분리하지 않음. Sch.9 C 'Priority Sectors' 줄과 PSL 보고는 psl_flag·psl_category로 산출(5. 속성표, 7. 주제 참조)", "PSL 분기·연간, Sch.9 C, BSR-1")

# ---- 3-3 예금 (Sch.3 순도) ---------------------------------------------------
rv("이전안수정", "20150300100", 5, "20150300000", "SB - RESIDENT INDIVIDUAL (INR)", "거주자 개인 저축(SB)",
   "이전안 'SB/CA'는 Sch.3 저축(II)과 요구불(I)을 섞음 → 저축 전용으로 변경. BSBDA·경로·임직원·미성년은 태그", "Sch.3 II, Form A(SB 요구불·정기 부분), BSR-2",
   prev="이전안 ADD 명칭 변경")
rv("신규", "20150300600", 5, "20150300000", "CA - RESIDENT INDIVIDUAL / PROPRIETOR (INR)", "거주자 개인·개인사업자 당좌",
   "20150300100에서 분리. Sch.3 I(ii) 요구불-기타", "Sch.3 I")
rv("신규", "20150300700", 5, "20150300000", "NRE CA (INR)", "NRE 당좌", "예금명세서 NRE-02. 이전안 누락", "Sch.3 I, NRD-CSR")
rv("신규", "20150300800", 5, "20150300000", "NRO CA (INR)", "NRO 당좌",
   "예금명세서 NRO-02·FN-02. 외국인 비거주자(FN) 계좌는 FEMA상 NRO → 같은 GL + nr_subtype=FN", "Sch.3 I, NRD-CSR")
rv("변경", "20150300300", 5, "20150300000", "NRO SB A/C (INR) - INCL. FOREIGN NATIONAL", "NRO 저축(외국인 비거주자 포함)",
   "(2차 권고 확정) 외국인 비거주자(FN)는 NRO의 부분집합 → GL 통합, nr_subtype=FN으로 부분합(시트 07)", "NRD-CSR", prev="이전안 ADD 설명 보강")
rv("신규", "20150100500", 5, "20150100000", "CA - LIAISON OFFICE (INR)", "연락사무소 당좌",
   "FEMA 22(R): 비거주 법인의 인도 사무소, UIN·AAC 관리. 예금명세서 CA-03(20125021*)", "Sch.3 I, AAC 연차보고, FLA")
rv("신규", "20150100600", 5, "20150100000", "CA - BRANCH OFFICE (INR)", "지점사무소 당좌", "예금명세서 CA-06(20125023*). 이익송금 증빙 관리", "Sch.3 I, AAC")
rv("신규", "20150100700", 5, "20150100000", "CA - PROJECT OFFICE (INR)", "프로젝트사무소 당좌", "예금명세서 CA-04(20125022*)", "Sch.3 I, AAC")
rv("신규", "20150100800", 5, "20150100000", "FCY A/C - PROJECT OFFICE", "프로젝트사무소 외화계좌", "예금명세서 CA-05(20125031*). 프로젝트당 1계좌", "Sch.3 I, NOP")
rv("신규", "20150100900", 5, "20150100000", "COLLECTION ACCOUNT (CRM DIR. CH.XIA)", "집금계좌(Collection)",
   "통과성·2영업일 스윕으로 잔액이 0에 가까워 일반 당좌 풀을 왜곡, 반기 10% 점유율 재검증 → 전용 GL(예금명세서 시트16 결론과 동일)", "Sch.3 I, Ch.XIA 반기점검")
rv("이전안수정", "20150100300", 5, "20150100000", "SNRR A/C (INR)", "SNRR 계정",
   "FEMA 5(R)(6)/2026(2026.6.18)으로 사업연계·명칭 요건 삭제 → 사업용/FPI 두 L5 구분 근거 약화. 하나로 통합하고 용도는 snrr_purpose 속성", "NRD-CSR, R-Return",
   prev="이전안 ADD 20150100300·400 통합")
rv("비활성", "20150100400", 5, "20150100000", "SNRR A/C - FPI / PORTFOLIO INVESTMENT (INR)", "SNRR(FPI) — 20150100300에 통합",
   "snrr_purpose=FPI 속성으로 대체", "", prev="이전안 ADD 철회")
rv("이전안수정", "20150400000", 4, "20150000000", "ESCROW DEPOSITS (TOTAL)", "에스크로 예금(L5 집계)",
   "예금명세서 시트16: 유형별 일일 대사(REQ-25) → FEMA·국내무이자·PA 3개 L5. 직접 기장 금지", "Sch.3 I, BLR-1(LCR)", prev="이전안 ADD L4 → 집계계정")
rv("신규", "20150400100", 5, "20150400000", "ESCROW - FEMA (FDI / OPEN OFFER / DEFERRED CONSIDERATION)", "FEMA 에스크로",
   "INR 전용·무이자·여신/상계 금지. escrow_type 속성(FDI·OPO·DEF)", "Sch.3 I, FC-TRS 연계")
rv("신규", "20150400200", 5, "20150400000", "ESCROW - DOMESTIC NON-INTEREST (TRA / COMMERCIAL / RERA)", "국내 무이자 에스크로", "", "Sch.3 I")
rv("신규", "20150400300", 5, "20150400000", "ESCROW - PAYMENT AGGREGATOR", "결제대행업자(PA) 에스크로",
   "유일한 이자부 요구불 에스크로(core portion). 이자비용은 50110100000 + escrow 태그(신규 이자 GL 불필요)", "Sch.3 I, PA 감사인 인증")
rv("신규", "20160100200", 5, "20160100000", "TERM DEPOSIT - ESCROW-LINKED (ESC-TD)", "에스크로 연계 정기예금", "parent_escrow_id 필수", "Sch.3 III")
rv("신규", "20160200100", 5, "20160200000", "TERM DEPOSITS FROM BANKS IN INDIA (IB-01)", "인도 내 은행 정기예금",
   "Sch.3 III(i) 은행분, Form A '은행시스템 부채', DICGC 부보 제외, TDS 없음. 예금명세서 20122051* 대체", "Sch.3, Form A, DICGC")
rv("신규", "20160200200", 5, "20160200000", "TERM DEPOSITS FROM BANKS OUTSIDE INDIA", "인도 외 은행 정기예금",
   "인도 외 은행은 Form A '은행시스템'이 아님 → 인도 내와 분리", "Form A, LBS", "Y")
rv("신규", "20160300600", 5, "20160300000", "RECURRING DEPOSIT - RESIDENT (INR)", "거주자 정기적금(RD)",
   "예금명세서 원칙(product_type이 GL 결정키)에 따라 RD 분리. 법인 RD는 드물어 개인 L4에만 둠", "Sch.3 III, TDS(194A)")
rv("신규", "20160300700", 5, "20160300000", "NRE RECURRING DEPOSIT (INR)", "NRE 정기적금", "", "Sch.3 III, NRD-CSR")
rv("신규", "20160300800", 5, "20160300000", "NRO RECURRING DEPOSIT (INR) - INCL. FOREIGN NATIONAL", "NRO 정기적금", "", "Sch.3 III, NRD-CSR")
rv("변경", "20160300300", 5, "20160300000", "NRO TERM DEPOSIT (INR) - INCL. FOREIGN NATIONAL", "NRO 정기예금(외국인 비거주자 포함)",
   "예금명세서 FN-03 통합", "NRD-CSR", prev="이전안 ADD 설명 보강")
rv("비활성", "20510100000", 4, "20501000000", "DORMANT ACCOUNT", "휴면계좌(부채 이관용)",
   "인도는 휴면·Inoperative 계좌도 예금(Sch.3) 그대로 두고 상태 속성으로 관리, 10년 미청구 시 DEAF 이전. 예금→기타부채 이관 금지", "Sch.3, DEAF 월보")
rv("비활성", "20510200000", 4, "20501000000", "STATUTORY RESERVE (LIABILITY)", "법정준비금(부채측)",
   "인도 법정준비금은 자본(Sch.2) → 31110300000 사용", "Sch.2")

# ---- 3-4 미지급·세금 ----------------------------------------------------------
rv("이전안수정", "20510400100", 5, "20510400000", "TDS PAYABLE - RESIDENT (26Q)", "TDS 예수금(거주자)",
   "이전안 '194A/195' 통합은 26Q(거주자)·27Q(비거주자) 신고서가 달라 대사 불가 → 분리. 조문(194A·194N, 신법 393)은 속성", "TDS 분기신고", prev="이전안 ADD 명칭 변경")
rv("신규", "20510400300", 5, "20510400000", "TDS PAYABLE - NON-RESIDENT (27Q)", "TDS 예수금(비거주자)",
   "NRO·FN 이자(195, 신법 393(2)). 예금명세서 20990023 대응", "TDS 분기신고")
rv("신규", "20510500400", 5, "20510500000", "RCM IGST PAYABLE", "역발행(RCM) IGST",
   "수입 서비스(해외 코레스 수수료, 본점 비용배분 등) RCM. 예금명세서 20990019~21 대응, 이전안 누락", "GSTR-3B")
rv("신규", "20510500500", 5, "20510500000", "RCM CGST PAYABLE", "역발행 CGST", "국내 RCM(법률·보안용역 등)", "GSTR-3B")
rv("신규", "20510500600", 5, "20510500000", "RCM SGST PAYABLE", "역발행 SGST", "", "GSTR-3B")
rv("속성전환", "-", "", "", "GST OUTPUT ON FX CONVERSION (20990016~18)", "환전분 GST 별도 GL",
   "예금명세서는 환전 가치 GST를 별도 GL로 제안 → 세목(CGST/SGST/IGST)이 같으므로 GL 불필요. SAC·과세표준 산식은 거래 속성", "GSTR-1")
rv("신규", "20440900000", 4, "20440000000", "ACCOUNT PAYABLE (UNIDENTIFIED ESCROW CREDITS)", "에스크로 미확인 입금",
   "예금명세서 20990041*. 약정 매칭 전 보관, 원송금인 반환 대상", "")

# ---- 3-5 PSL·DEAF 등 기타자산·수익 ---------------------------------------------
rv("신규", "10490920000", 4, "10490000000", "DEPOSITS IN LIEU OF PSL SHORTFALL (NABARD / SIDBI / NHB / MUDRA)", "PSL 미달 예치금",
   "PSL 목표 미달분을 RIDF 등 기금에 예치. 표시 위치 Sch.11 여부 확인", "Sch.11, PSL 연간", "Y")
rv("신규", "40130980000", 4, "40130000000", "INT. ON PSL SHORTFALL DEPOSITS", "PSL 미달 예치금 이자", "", "Sch.13")
rv("신규", "40591070000", 4, "40591000000", "FEE RECEIVED ON SALE OF PSLC", "PSLC 매도 수수료",
   "PSL 목표 초과분 PSLC(e-Kuber) 매도. 손익 표시 줄 확인", "Sch.14, PSL", "Y")
rv("신규", "50510970000", 4, "50510000000", "FEE PAID ON PURCHASE OF PSLC", "PSLC 매입 수수료",
   "외국은행 지점이 PSL 목표 달성에 흔히 쓰는 수단", "Sch.16, PSL", "Y")
rv("신규", "10490930000", 4, "10490000000", "AMOUNT RECEIVABLE FROM DEAF (CLAIMS PAID)", "DEAF 청구 미수",
   "DEAF 이전 후 예금주에게 지급한 금액을 RBI에 청구. 이전 자체는 예금 직접 차감", "DEAF 월보, Sch.11")

# ---- 3-6 자본·충당금 ----------------------------------------------------------
rv("변경", "31110300000", 4, "31110000000", "STATUTORY RESERVE", "법정준비금(Statutory Reserve)",
   "통합명 RETAINED PROFIT(본점 매핑 '이익준비금' = 한국 상법상 법정적립금) → 인도 표시명 STATUTORY RESERVE(BR Act s.17, RBI는 외국은행 지점도 순이익 25% 이상 적립 요구). "
   "현행 Legal Reserve(28810004)·Statutory Reserves(30110011)를 이리로. 기타준비금은 31110400000", "Sch.2", "Y")
rv("신규", "31110500000", 4, "31110000000", "CAPITAL RESERVE", "자본준비금", "HTM 매각익 등 적립", "Sch.2", "Y")
rv("신규", "31110600000", 4, "31110000000", "INVESTMENT FLUCTUATION RESERVE (IFR)", "투자변동준비금", "AFS·HFT 보유분 기준 적립", "Sch.2, CRAR", "Y")
rv("신규", "31110700000", 4, "31110000000", "REMITTABLE SURPLUS RETAINED IN INDIA FOR CRAR", "자본비율용 인도 유보 송금가능잉여금",
   "외국은행 지점 Tier 1 구성항목", "Sch.2, CRAR", "Y")
rv("변경", "31130000000", 3, "31100000000", "AFS RESERVE (UNREALISED GAIN/LOSS)", "AFS 평가손익",
   "RBI 투자 지침(2023): AFS 평가손익은 AFS Reserve. FVOCI 계정을 매핑", "Sch.2", "Y")
rv("철회", "20990100000", 4, "20990000000", "PROVISION - UNHEDGED FOREIGN CURRENCY EXPOSURE", "미헤지 외화익스포저 충당금",
   "(2차) ECL 명세서 결론 수용: UFCE는 Stage 충당금(10380103000/104000)의 floor 가산으로 흡수, UFCE_BP·RATIO·COMPONENT 태그로 추적", "UFCE 분기, 주석", prev="1차 제안 철회")
rv("철회", "20990200000", 4, "20990000000", "PROVISION - COUNTRY RISK", "국가위험 충당금", "(2차) ECL para 125: 국가위험은 PD·LGD 오버레이로 Stage 충당금에 반영 → 별도 GL 불필요, country_risk_grade 속성", "주석", prev="1차 제안 철회")
rv("철회", "20990300000", 4, "20990000000", "PROVISION - STANDARD ASSETS (STAGE 1·2) IF PRESENTED IN SCH.5", "정상자산 충당금(부채표시분)",
   "(2차) ECL 도입으로 Stage 1·2도 대출 차감 충당금(10380103000/104000). 부외 ECL만 부채(20910000000~20930000000)", "Sch.9, CRAR(Tier 2)", "Y", prev="1차 제안 철회")
rv("비활성", "10380101000", 5, "10380100000", "ALLOWANCE FOR LOAN LOSSES - GENERAL", "대손충당금(일반)",
   "(2차) ECL 명세서 요청대로 Stage별 L5(10380103000~105000) 신설. 2027-03-31까지 사용 후 이관·병행대사 → 비활성", "CRAR, 주석", prev="1차 의견(일반/특정 2개 유지) 변경")
rv("비활성", "10380102000", 5, "10380100000", "ALLOWANCE FOR LOAN LOSSES - SPECIFIC", "대손충당금(특정)", "(2차) 10380105000(Stage 3)로 이관 후 비활성", "")

# ---- 3-7 난외·메모 ------------------------------------------------------------
rv("신규", "60300000000", 2, "", "MEMORANDUM ACCOUNTS (INDIA)", "비망계정(인도)", "대차 영향 없음, 보고·추적용", "")
rv("신규", "60310000000", 3, "60300000000", "INTEREST NOT RECOGNISED - NPA / STAGE 3", "미인식 이자(NPA·Stage 3)",
   "IRAC: NPA 이자는 현금주의. 청구는 하되 수익 미인식분 추적. ACPIR 적용 시 인식기준 확인", "IRAC, 주석", "Y")
rv("신규", "60320000000", 3, "60300000000", "TECHNICALLY WRITTEN-OFF ACCOUNTS", "기술적 상각 채권", "상각 후 회수 추적·공시", "주석, CRILC")
rv("속성전환", "60130400100", 5, "60130400000", "CUST LIABS ON GUARANTEES - FINANCIAL BG", "금융보증 L5 → 속성",
   "Sch.12는 금융·이행을 나누지 않음. CCF(100%/50%)는 bg_nature 속성으로 산출 → L5 철회 권고(유지해도 무방, 본점 선택)", "CRAR, CEOBE",
   prev="이전안 ADD 재검토")
rv("속성전환", "60130400200", 5, "60130400000", "CUST LIABS ON GUARANTEES - PERFORMANCE BG", "이행보증 L5 → 속성", "위와 같음", "", prev="이전안 ADD 재검토")
rv("속성전환", "60120400000", 4, "60120000000", "OUTSTANDING INLAND L/C (SFMS)", "내국신용장 L4 → 속성",
   "Sch.12에 내국/수입 구분 없음. lc_scope 속성(INLAND/IMPORT). 내국 LC는 Sight/Usance L4에 포함", "CEOBE", prev="이전안 ADD 재검토")
rv("매핑", "60120100000", 4, "60120000000", "OUTSTANDING SIGHT L/C", "일람불 LC",
   "예금명세서 TF-02가 LC를 72040001 'Committed Credit Line'에 매핑 → LC는 60120000000 계열이 맞음(정정 요청)", "Sch.12")
rv("매핑", "60130400000", 4, "60130000000", "CUST LIABS ON GUARANTEES (BONDS)", "보증(BG)",
   "예금명세서 LN-03 71040004 → 본 계정. 수익자 인도 내·외는 속성(Sch.12 (a)/(b))", "Sch.12")

rv("신규", "40592040000", 4, "40592000000", "COMMISSION - ESCROW SET-UP & ANNUAL FEE", "에스크로 설정·연간 수수료",
   "예금명세서는 42020729(0930상 대출 기타수수료) 사용 → 대출수수료와 섞임. 계좌관리 서비스 그룹에 신설", "Sch.14, GST")

# ---- 3-8 이자비용·미지급이자 (신설 불필요) --------------------------------------
rv("속성전환", "-", "", "", "INTEREST EXPENSE BY PRODUCT × RESIDENCY (51040011*, 51050011*, 51060014~16*, 51040041*, 51060051*)",
   "상품·거주성별 이자비용 GL",
   "Sch.15는 '예금이자' 한 줄. 통합 L4(50110·50120 고객유형별)에 기장하고 상품·거주성은 계좌 속성. TDS·NRD-CSR도 계좌 기준으로 산출 가능", "Sch.15")
rv("속성전환", "-", "", "", "INTEREST PAYABLE - FCNR (FCY) 20910012*", "FCNR 외화 미지급이자 GL",
   "통화차원으로 충분(P5). 20410030000 + 통화코드", "")
rv("매핑", "20410020000", 4, "20410000000", "A.I.P ON DEMAND DEPOSIT", "요구불(저축 포함) 미지급이자", "예금명세서 20910002와 0930 GBS코드 일치(충돌 없음)", "Sch.5")
rv("매핑", "20410030000", 4, "20410000000", "A.I.P ON TIME DEPOSIT", "정기 미지급이자", "예금명세서 20910003과 일치. FCNR·ESC-TD 포함", "Sch.5")


# ---------------------------------------------------------------------------
# 4. 예금 명세서 상품 → 계정 매핑 (상품, 명세서 원금GL, 제안 원금 GL, 이자비용 GL, 미지급이자 GL, TDS GL, 결정 속성)
# ---------------------------------------------------------------------------
PRODUCT_MAP = [
    ("CA-01", "당좌—일반", "20125001", "법인 20150100100 / 개인 20150300600", "-", "-", "-", "constitution"),
    ("CA-02", "당좌—EEFC", "20125003", "법인 20150100200 / 개인 20150300500", "-", "-", "-", "constitution"),
    ("CA-03", "연락사무소 당좌", "20125021*", "20150100500", "-", "-", "-", "uin, aac_due_date"),
    ("CA-04", "프로젝트사무소 당좌(INR)", "20125022*", "20150100700", "-", "-", "-", "aac_due_date"),
    ("CA-05", "프로젝트사무소 외화", "20125031*", "20150100800", "-", "-", "-", "project_id"),
    ("CA-06", "지점사무소 당좌", "20125023*", "20150100600", "-", "-", "-", "uin, aac_due_date"),
    ("CA-07", "집금계좌", "20125004*", "20150100900", "-", "-", "-", "bank_share_pct, designated_account"),
    ("SB-01~04S", "저축(일반·경로·임직원·BSBDA·소액)", "20121001", "20150300100", "50110300000", "20410020000", "-",
     "segment_tag, feature_tag, kyc_level"),
    ("TD-01", "정기적금(RD)", "20124001", "20160300600", "50120300000", "20410030000", "20510400100", "callability"),
    ("TD-02", "정기예금(FD)", "20122001", "개인 20160300100 / 법인 20160100100", "50120300000 / 50120100000", "20410030000",
     "20510400100", "constitution, callability, amount_slab"),
    ("NRE-01", "NRE 저축", "20121011*", "20150300200", "50110300000", "20410020000", "-(비과세)", "repat_tag"),
    ("NRE-02", "NRE 당좌", "20125011*", "20150300700", "-", "-", "-", ""),
    ("NRE-03", "NRE 정기", "20122011*", "20160300200", "50120300000", "20410030000", "-(비과세)", ""),
    ("NRE-04", "NRE 적금", "20124011*", "20160300700", "50120300000", "20410030000", "-(비과세)", ""),
    ("NRO-01", "NRO 저축", "20121012*", "20150300300", "50110300000", "20410020000", "20510400300", "remit_ytd_usd"),
    ("NRO-02", "NRO 당좌", "20125012*", "20150300800", "-", "-", "-", ""),
    ("NRO-03", "NRO 정기", "20122012*", "20160300300", "50120300000", "20410030000", "20510400300", ""),
    ("NRO-04", "NRO 적금", "20124012*", "20160300800", "50120300000", "20410030000", "20510400300", ""),
    ("FCNR-01", "FCNR(B) 외화정기", "20122021*", "20160300400", "50120300000", "20410030000(+통화)", "-(비과세)", "currency"),
    ("SNRR-01", "SNRR", "20125301* / 20125013*", "20150100300 (개인은 별도 협의)", "-", "-", "-", "snrr_purpose"),
    ("FN-01~04", "외국인 비거주자 저축·당좌·정기·적금", "20121002 / 20125002 / 20122002 / 20124002", "NRO 해당 L5",
     "NRO와 같음", "NRO와 같음", "20510400300", "nr_subtype=FN"),
    ("ESC-FDI/OPO/DEF", "FEMA 에스크로", "20125041*", "20150400100", "-", "-", "-", "escrow_type, agreement_id"),
    ("ESC-TRA/COM/RERA", "국내 무이자 에스크로", "20125042*", "20150400200", "-", "-", "-", "escrow_type, rera_state_code"),
    ("ESC-PA", "PA 에스크로", "20125043*", "20150400300", "50110100000", "20410020000", "20510400100", "pa_scope, core_portion_amt"),
    ("ESC-TD", "에스크로 연계 정기", "20122041*", "20160100200", "50120100000 / 50120300000", "20410030000", "수취인 거주성별",
     "parent_escrow_id"),
    ("IB-01", "은행간 정기예금", "20122051*", "20160200100", "50120200000", "20410030000", "-(은행)", "money_market_eligible"),
    ("LN-01", "CC/OD", "10630001*", "10330304000", "40110100000", "A.I.R 10410030000", "-", "dp_limit, out_of_order"),
    ("LN-01W", "WCDL", "10610004/05", "10330302000 / 10330303000", "40310400000 / 40310500000", "10410030000", "-", "facility_type"),
    ("LN-02", "기한부·신디케이트", "10610007", "10330401000 / 1034xxxxxxx", "40310700000 / 40320xxxxxx", "10410050000·10410060000", "-", "facility_type"),
    ("LN-03", "지급보증", "71040004", "60130400000", "수수료 40510060000", "선수 20420170000", "-", "bg_nature, beneficiary_location"),
    ("TF-02", "신용장", "72040001 (오류)", "60120100000 / 60120200000", "40520020000 등", "선수 20420190000", "-", "lc_scope, tenor_type"),
    ("TF-03", "어음할인", "10420002/03", "10320300000(국내·TReDS) / 10310xxxxxx(수출)", "40350300000 / 40340xxxxxx", "10410090000", "-",
     "bill_origin(INLAND/FOREIGN)"),
]

# ---------------------------------------------------------------------------
# 5. RBI 보고용 속성 (영역, 속성(컬럼), 값·형식, 수준, 사용 보고서, 근거, 현재 설계 반영, VERIFY)
#    수준: 고객 / 계좌 / 거래 / 한도 / 담보 / Nostro
#    현재 설계 반영: 예금명세서(시트) / TF모델 / 송금모델 / 신규
# ---------------------------------------------------------------------------
ATTRS = [
    # 고객 공통
    ("고객", "constitution", "IND·HUF·PROP·PARTNERSHIP·LLP·PVT·PUB·TRUST·SOCIETY·GOVT·PSU·BANK·NBFC·FI·FOREIGN_CO", "고객",
     "예금 GL 결정(L4 고객유형군), BSR-1/2, Sch.9 C(공공부문·은행), LCR", "BSR 코드체계", "부분(personality)", "N"),
    ("고객", "bsr_sector_code", "BSR 기관부문 코드(가계·민간법인·공기업·정부·금융·외국)", "고객", "BSR-1·BSR-2·BSR-7", "RBI BSR Handbook", "신규", "Y"),
    ("고객", "residency_status", "RESIDENT·NRI·PIO/OCI·FOREIGN_NATIONAL·FOREIGN_ENTITY", "고객", "NRD-CSR, FLA, LBS, 예금 GL", "FEMA 2(v)/(w)", "예금명세서 17", "N"),
    ("고객", "country_code", "ISO 3166 (거주국)", "고객", "NRD-CSR, FLA, LBS, 국가위험", "-", "예금명세서 18", "N"),
    ("고객", "pan / lei / ckyc_no", "PAN 10자, LEI 20자, CKYC 14자", "고객", "TDS, CIC, CRILC, 송금 LEI", "IT Act, LEI 지침", "송금모델(LEI)", "N"),
    ("고객", "udyam_no / msme_category / msme_class_date", "MICRO·SMALL·MEDIUM, 분류일·재분류 이력", "고객",
     "PSL(MSME), ACPIR §82 범주(iii)(iv), BSR", "MSMED Act, PSL MD", "TF모델(TReDS만)", "N"),
    ("고객", "group_id", "기업집단 식별자", "고객", "LEF(거액익스포저), CRILC", "LEF 지침", "신규", "N"),
    ("고객", "related_party / director_interest", "Y/N, 관계유형", "고객", "BR Act s.20 이사 관련 대출 제한, 공시", "BR Act s.20", "신규", "N"),
    ("고객", "wilful_defaulter_flag", "Y/N, 등재일", "고객", "ACPIR §101(4) 추가 5%, 신규여신 제한", "Wilful Defaulter MD", "신규", "N"),
    ("고객", "weaker_section_flag", "SC/ST·소농·여성·소수자 등 PSL 약자 기준", "고객", "PSL 약자 하위목표", "PSL MD", "신규", "Y"),
    # 여신
    ("여신", "facility_type", "CC·OD·WCDL·TL·PC·PCFC·BP·BD·TReDS·BC 등", "한도/계좌", "Sch.9 A, BSR-1, Form A(은행신용 내역), IRAC(out-of-order)", "Third Schedule", "부분", "N"),
    ("여신", "security_class", "TANGIBLE·BANK_GOVT_GUARANTEE·UNSECURED", "담보", "Sch.9 B, ACPIR LGD(§98 30%)", "Third Schedule", "신규", "N"),
    ("여신", "sector_class (Sch.9 C)", "PRIORITY·PUBLIC·BANKS·OTHERS (산출값)", "계좌", "Sch.9 C", "Third Schedule", "신규", "N"),
    ("여신", "psl_flag / psl_category / psl_subcategory", "AGRI·MSME(M/S/M)·EXPORT·EDUCATION·HOUSING·SOCIAL_INFRA·RENEWABLE·OTHERS / 세부", "계좌",
     "PSL 분기·연간, ANBC 대비 달성률, Sch.9 C", "PSL Master Directions(2025)", "TF모델(PSL_CAT 3값)", "Y"),
    ("여신", "psl_basis / psl_eval_date", "DIRECT·ON_LENDING(NBFC)·PSLC·INVESTMENT, 판정일", "계좌", "PSL(간접대출 상한), 재판정 이력", "PSL MD", "신규", "Y"),
    ("여신", "export_credit_stage", "PRE_SHIPMENT·POST_SHIPMENT, EDPMS 선적서류 번호", "계좌", "PSL 수출신용, 수출신용 보고, EDPMS 상환대사", "FEMA EXIM", "TF모델", "N"),
    ("여신", "bsr_activity_code", "BSR 활동코드(NIC 기반)", "계좌", "BSR-1, 업종 집중", "BSR Handbook", "신규", "Y"),
    ("여신", "benchmark / spread / reset_date", "EBLR·MCLR·T-bill·SOFR 등", "계좌", "금리 공시, BSR-1(금리), 공정대출관행", "MCLR/EBLR 지침", "부분", "N"),
    ("여신", "acpir_category", "§82 (i)~(xv) 범주", "계좌", "충당금 하한, 범주 혼합 금지(§92)", "ACPIR 2026", "예금명세서 37", "N"),
    ("여신", "stage / irac_class / dpd / sma_class", "STAGE1~3, STD·SUB·DOUBT·LOSS, SMA-0/1/2", "계좌", "CRILC, CIC, 충당금, 주석", "ACPIR, CRILC 지침", "예금명세서 34·35", "N"),
    ("여신", "out_of_order_flags", "한도초과 90일·무입금 90일·이자미충당", "계좌", "CC/OD NPA 판정", "ACPIR §6(17)", "예금명세서 34", "N"),
    ("여신", "restructured / fraud_flag", "Y/N, 일자, 사유", "계좌", "충당금(§101 fraud 전액), 사기보고", "ACPIR, Fraud MD", "부분", "N"),
    ("여신", "crilc_reportable", "그룹 합산 ≥ ₹5crore 자동판정", "고객", "CRILC(분기·주간 SMA)", "CRILC 지침", "예금명세서 23", "Y"),
    ("여신", "ufce_ratio / hedged_amt", "차주 미헤지 외화익스포저, EBID 대비", "고객", "UFCE 추가 충당금·자본", "UFCE 지침", "신규", "Y"),
    ("여신", "ecgc_cover / guarantee_scheme", "ECGC·CGTMSE·NCGTC", "계좌", "Sch.9 B(ii), ACPIR 범주(xiii), CRM", "-", "TF모델(ECGC)", "N"),
    ("여신", "against_own_deposit", "Y/N, 담보 예금 계좌·마진율", "담보", "ACPIR §9(5)·범주(vii)", "ACPIR", "예금명세서 36", "N"),
    ("여신", "sensitive_sector", "CAPITAL_MARKET·REAL_ESTATE·COMMODITY", "계좌", "민감부문 익스포저 공시", "RBI 공시 지침", "신규", "N"),
    # 예금
    ("예금", "product_type / residency / currency", "SB·CA·TD·RD / RES·NRE·NRO·FCNR·RFC·EEFC·SNRR / 통화", "계좌", "GL 결정(시트 4 매핑)", "예금명세서 16·17", "예금명세서 17", "N"),
    ("예금", "nr_subtype", "NRI·PIO·FN", "계좌", "NRD-CSR, 외국인 비거주자 출국 6개월 규칙", "FEMA Deposit Regs", "신규", "N"),
    ("예금", "sb_demand_portion", "저축예금 요구불·정기 부분 산정비율(행태 분석)", "GL", "Form A(저축의 요구불/정기 분할)", "RBI Act s.42", "신규", "Y"),
    ("예금", "lcr_class / stable_amt / less_stable_amt", "RT_S·RT_L·SB_S·SB_L·NFC_OP·NFC_NOP·OLE_NF·FI_OP·FI_NOP·IG_HO·TD_EXCL (시트 11)", "계좌(보고일별)",
     "LCR(BLR-1) 유출률. 고객 분류 + 부보 + 관계로 보고일마다 산출, 스냅샷 저장", "RBI LCR 지침", "신규(예금명세서는 에스크로만)", "Y"),
    ("예금", "operational_flag / operational_agreement_ref / operational_amt", "Y/N, 약정번호, 운영 인정액", "계좌", "LCR 25% 적용분", "RBI LCR 지침", "신규", "Y"),
    ("예금", "withdrawable_30d", "Y/N (callability·잔존만기·중도해지 벌칙에서 산출)", "계좌", "LCR 유출 제외 판정", "RBI LCR 지침", "신규(callability 태그 연계)", "Y"),
    ("고객", "is_financial / fin_entity_type", "Y/N, BANK·NBFC·INSURER·MF·PENSION·OTHER_FI", "고객", "LCR 금융기관 100%, ECL 범주(vi), LEF", "RBI LCR·ECL", "신규", "N"),
    ("고객", "relationship_start_dt / transactional_acct_yn", "최초 거래일, 급여·자동이체 등 거래계좌", "고객/계좌", "소매 안정 판정", "RBI LCR 지침", "신규", "N"),
    ("예금", "digital_enabled_flag", "인터넷·모바일뱅킹 이체 가능 여부", "계좌", "LCR 소매예금 추가 유출률(2.5%p)", "LCR 개정(2026.4.1 시행)", "신규", "Y"),
    ("예금", "dicgc_insurable / insured_amt", "부보대상 여부, 예금자별 한도(₹5 lakh) 내 금액", "고객", "DICGC 보험료 기준, 부보예금 공시", "DICGC Act", "부분", "N"),
    ("예금", "original_maturity / residual_bucket / size_band", "원만기, 잔존만기 구간, 금액구간(₹3crore Bulk)", "계좌", "BSR-2, ALM(SLS), Form A 면제항목, NRD-CSR", "-", "부분", "N"),
    ("예금", "crr_slr_exempt_reason", "면제 사유 코드", "계좌", "Form A 차감항목", "RBI Act s.42", "예금명세서 23", "Y"),
    ("예금", "account_status / inoperative_date / deaf_due_date", "ACTIVE·INOPERATIVE·DORMANT·FROZEN·DEAF", "계좌", "DEAF 월보, 휴면 공시", "DEAF Scheme 2014", "예금명세서 18", "N"),
    ("예금", "escrow_type / agreement_id / parent_escrow_id", "FDI·OPO·DEF·TRA·COM·PA·RERA·TD", "계좌", "에스크로 대사, LCR, PA 인증", "FEMA Sch.5, PA Dir.", "예금명세서 18", "N"),
    ("예금", "repat_tag / remit_ytd_usd", "REPATRIABLE·RESTRICTED / NRO 연 USD 1M 누계", "계좌", "NRO 송금한도, R-Return", "FEMA", "예금명세서 17", "N"),
    ("예금", "snrr_purpose", "BUSINESS·FPI·ECB·OTHERS", "계좌", "NRD-CSR, FPI 지정계좌 통제", "FEMA 5(R)(6)/2026", "송금모델(VOSTRO_PURP)", "Y"),
    # Nostro·은행간
    ("Nostro·은행간", "nostro_type", "FCY_NOSTRO_ABROAD·INR_SETTLEMENT_IN_INDIA·FCY_IN_INDIA·RBI·HO_OWN_BRANCH", "Nostro",
     "Sch.7 I/II, Form A 은행시스템 자산, NOP, LBS", "Third Schedule, RBI Act s.42", "부분(TB_NOSTRO_ACCT)", "N"),
    ("Nostro·은행간", "cpty_bank_type", "RBI·SCB(국내)·FOREIGN_BANK_BR·COOP·RRB·CCIL·OVERSEAS_BANK·OWN_HO/BR", "Nostro/거래",
     "Form A '은행시스템' 해당 여부, LEF 은행간, Sch.7 콜론(은행·기타기관)", "RBI Act s.42", "신규", "Y"),
    ("Nostro·은행간", "cpty_country / currency", "ISO 국가·통화", "Nostro", "LBS(국제은행통계), 국가위험", "RBI LBS", "부분", "Y"),
    ("Nostro·은행간", "encumbrance", "CCIL_MARGIN·LAF·S11_2·RTGS_IDL·NONE", "Nostro/투자", "LCR HQLA 제외, SLR 제외, Sch.1 주기", "-", "신규", "N"),
    ("Nostro·은행간", "recon_status / open_item_age", "대사 상태·미결 경과일", "Nostro", "Nostro 대사(감사·LFAR), 미결항목 처리", "내부·감사", "부분", "N"),
    ("Nostro·은행간", "money_market_eligible", "콜·노티스·텀머니 참가기관 여부", "고객", "IB-01 개설·Form A 은행분", "IRD para 44", "예금명세서 49", "N"),
    # 투자
    ("투자", "rbi_category", "HTM·AFS·FVTPL(HFT 포함)", "보유", "Sch.8, IFR, 자본(시장위험)", "RBI 투자 지침 2023", "신규", "Y"),
    ("투자", "issuer_type / slr_eligible", "GOI·SDL·T-BILL·PSU·CORP·BANK(CD)·CP·FOREIGN / Y·N", "보유", "Sch.8 줄, Form VIII(SLR)", "BR Act s.24", "신규", "N"),
    ("투자", "listed / rated", "Y/N", "보유", "비SLR 투자 한도 공시", "-", "신규", "N"),
    # 난외·외환
    ("난외·외환", "bg_nature", "FINANCIAL·PERFORMANCE·BID", "계좌", "CCF·CEOBE, ACPIR §87", "NFB Directions 2025", "이전안(GL L5)", "N"),
    ("난외·외환", "beneficiary_location", "IN_INDIA·OUTSIDE_INDIA", "계좌", "Sch.12 보증 (a)/(b)", "Third Schedule", "신규", "N"),
    ("난외·외환", "lc_scope / tenor_type", "INLAND·IMPORT / SIGHT·USANCE", "계좌", "CEOBE, IDPMS", "-", "이전안(GL L4)", "N"),
    ("난외·외환", "counter_guarantee", "HO·BANK·NONE, 보증은행", "계좌", "신용위험 경감(CRM), LEF", "Basel III", "신규", "Y"),
    ("난외·외환", "fx_deal_cpty_type / hedge_basis", "MERCHANT·INTERBANK / 실수요 증빙(underlying ref)", "거래", "Sch.12 선물환, NOP·AGL, 외환 위험관리 지침", "FEMA Risk Mgmt Facilities", "부분", "Y"),
    # ECL·자본(RWA) — ECL 명세·RWA 명세에서 요구, GL 아님
    ("ECL·자본", "stage_tag / sicr_rule_hit / low_credit_risk", "1·2·3 / Q1~QL6 / Y·N", "계좌(부외 포함)", "Stage별 충당금 집계, 주석", "ECL Dir. 2026 para 25~37", "ECL 명세 9", "N"),
    ("ECL·자본", "poci_flag / credit_adj_eir", "Y·N / 율", "계좌", "POCI 이자·충당", "ECL Dir. 2026 para 24", "ECL 명세 9", "N"),
    ("ECL·자본", "pd_12m / pd_lifetime / lgd / ead / scenario_weight", "모델값·백스톱(PD 0.03%, LGD 65/70/30%)", "계좌", "ECL 산출, 감사추적", "para 42~100", "ECL 명세 3", "N"),
    ("ECL·자본", "ufce_bp / ufce_ratio / ufce_component", "0·20·40·60·80bp / 비율 / 금액", "고객·계좌", "Floor 가산, RWA +25%p, 공시", "para 126, CRM Dir. Ch.VII", "ECL 명세 9·11", "N"),
    ("ECL·자본", "country_risk_grade / korea_linked_flag", "ECGC 등급 / Y·N", "고객", "SICR(QL5), PD 오버레이", "para 125", "ECL 명세 8", "N"),
    ("ECL·자본", "external_rating / rating_agency / rating_date / rating_scope / inc_flag", "등급·기관·일자·발행/발행자·INC", "고객/한도", "위험가중치, ODR 조정", "SA Dir. 2026 para 25~32", "RWA 명세 9", "N"),
    ("ECL·자본", "ucc_flag / original_maturity / agg_wc_limit_system", "무조건 취소 가능 / 개월 / ₹150crore", "한도", "미사용약정 CCF(5·30·40%·20%)", "SA Dir. Table 19", "RWA 명세 9·13", "N"),
    ("ECL·자본", "guarantee_purpose_code", "금융(100%)·이행(50%)·무역LC(20%) 판정코드", "계좌", "CCF·bg_nature 산출", "SA Dir. 시트13 D", "RWA 명세 13", "N"),
    ("ECL·자본", "collateral_type / value / maturity / currency / revaluation_freq", "담보 원장", "담보", "CRM 헤어컷, LGD 30%", "SA Dir. para 35~37", "RWA 명세 9", "N"),
    ("ECL·자본", "guarantor_id / guarantor_rating / cover_amount", "보증인 원장", "담보", "위험가중치 대체", "SA Dir. para 39", "RWA 명세 9", "N"),
    ("ECL·자본", "s11_2b_deposit_flag", "Y·N", "계정(10104010200)", "본점 익스포저 경감 + 자본 제외(이중계상 방지)", "BR Act s.11(2)(b)(i)", "RWA 명세 9", "Y"),
    # 본지점
    ("본지점", "ho_flag / io_purpose", "Y/N, CAPITAL·FUNDING·SETTLEMENT·RA·VOSTRO", "계좌/거래", "Sch.5/11 순액, FLA, LEF 예외, Form A", "-", "송금모델(TB_IO_TXN)", "N"),
    # 세무
    ("세무", "tds_section / tds_form", "194A·194N·195(신법 393) / 26Q·27Q", "거래", "TDS 납부·신고 대사", "IT Act 2025", "예금명세서 04", "N"),
    ("세무", "form15g_15h / trc_10f / dtaa_rate", "제출여부·유효기간·조약세율", "고객", "TDS 면제·감면", "IT Act", "예금명세서 04", "N"),
    ("세무", "gst: sac / pos_state / gstin / rcm_flag / itc_eligible", "SAC 9971·9997, 공급지, B2B/B2C", "거래", "GSTR-1·3B, ITC 50% 선택", "CGST Act", "부분", "N"),
]

# ---------------------------------------------------------------------------
# 6. 보고서별 요구 (보고서, 주기(확인 필요 표시), 출처 계정, 필요한 속성, 산출 방식, VERIFY)
# ---------------------------------------------------------------------------
RETURNS = [
    ("Form A (RBI Act s.42) - CRR", "격주", "Sch.3 예금 L5, 차입, 은행간 예치·콜, 은행신용",
     "cpty_bank_type, sb_demand_portion, crr_slr_exempt_reason, nostro_type, ho_flag",
     "GL(요구·저축·정기, 은행분) × 속성(인도 은행시스템 여부, 저축 분할, 면제)", "N"),
    ("Form VIII - SLR", "격주/월", "Form A와 동일 + 투자", "issuer_type, slr_eligible, encumbrance", "SLR 증권 − 담보제공분(S11_2·LAF 등)", "N"),
    ("PSL 분기·연간 (CIMS)", "분기·연간", "여신 전체, 투자(PSL 적격), PSLC, 미달예치", "psl_*, msme_*, export_credit_stage, weaker_section_flag",
     "ANBC/CEOBE 산정 후 범주별 달성률. Sch.9 C 우선부문 잔액과 일치 검증", "Y"),
    ("BSR-1 / BSR-2 / BSR-7", "연간·분기", "여신·예금 계좌", "bsr_sector_code, bsr_activity_code, facility_type, size_band, original_maturity, 금리",
     "계좌 단위 추출", "Y"),
    ("CRILC (Main·SMA)", "분기·주간", "여신·난외", "group_id, sma_class, crilc_reportable, pan/lei", "그룹 합산 ≥ ₹5crore", "Y"),
    ("CIC 신용정보 제출", "격주", "여신 전체", "pan, ckyc_no, dpd, irac_class, 한도·잔액", "계좌 단위", "Y"),
    ("NRD-CSR", "월", "NRE·NRO·FCNR·SNRR L5", "residency_status, nr_subtype, country_code, currency, original_maturity, repat_tag",
     "GL(상품 L5) × 국가·통화·만기", "N"),
    ("FLA (외국인 부채·자산)", "연간", "비거주자 예금, 본지점, Nostro", "residency_status, country_code, ho_flag", "", "N"),
    ("LBS (국제은행통계)", "분기", "Nostro, 비거주자 대출·예금, 외화 포지션", "cpty_country, currency, cpty_bank_type, bsr_sector_code",
     "인도 지점의 대외·외화 포지션. Nostro는 국가·통화·상대방 유형이 필수", "Y"),
    ("R-Return / FETERS (CIMS)", "격주", "송금·환전", "purpose_code, fx_flag, io_purpose", "송금모델(TB_REMIT_MST) 기준", "N"),
    ("NOP / AGL (외환 포지션)", "일·월", "외화 자산·부채, 선물환", "currency, nostro_type, fx_deal_cpty_type", "INR 계정 제외, FCY 통화별", "N"),
    ("LCR (BLR-1) / ALM(SLS·IRS)", "월·격주", "예금·여신·HQLA", "lcr_depositor_class, operational_flag, digital_enabled_flag, encumbrance, residual_bucket",
     "", "Y"),
    ("LEF 거액익스포저", "분기", "여신·난외·은행간", "group_id, cpty_bank_type, counter_guarantee, ho_flag", "", "Y"),
    ("Basel III 자본비율", "분기", "전 자산·난외", "bg_nature, security_class, ecgc_cover, counter_guarantee, rating", "CCF·위험가중치", "N"),
    ("UFCE", "분기", "외화 여신 차주", "ufce_ratio, hedged_amt", "추가 충당금 20990100000", "Y"),
    ("DEAF", "월", "예금 전체(휴면), DD·PO", "account_status, inoperative_date, deaf_due_date", "10년 미청구 이전, 청구 시 10490930000", "N"),
    ("DICGC 보험료", "반기", "예금(은행간·정부 등 제외)", "dicgc_insurable, constitution", "", "N"),
    ("TDS 26Q / 27Q, GST 1·3B", "분기·월", "20510400100/300, 20510500xxx", "tds_section, pan, sac, pos_state, rcm_flag", "GL 잔액 = 신고 합계 대사", "N"),
    ("에스크로·PA·RERA 증빙", "수시·분기", "20150400100~300, 20160100200", "escrow_type, agreement_id, pa_scope", "예금명세서 시트 46", "N"),
]

# ---------------------------------------------------------------------------
# 7. 주제별 상세 — PSL, INR Nostro
# ---------------------------------------------------------------------------
TOPICS = [
    *[("FN·NRO", a, b + " — " + c) for a, b, c in __import__("coa_migration_model").FN_NRO],
    ("PSL", "왜 GL이 아닌가",
     "같은 운전자금대출이 차주 Udyam 재분류, 투자·매출 변경, 용도(수출·국내)에 따라 PSL 여부와 범주가 바뀐다. GL로 나누면 재분류마다 원금·미수이자·충당금·선수수익 계정을 "
     "모두 대체 분개해야 하고 기간 손익이 왜곡된다. 반면 PSL 보고는 분기 말 잔액 기준이므로 속성으로 충분하다."),
    ("PSL", "Sch.9 C 줄과의 관계",
     "재무제표 Sch.9 C.I '(i) Priority Sectors'는 B/S 표시 줄이지만, 연말에 psl_flag=Y 잔액을 집계해 표시한다(GL × 속성 산출). PSL 분기보고 합계와 같아야 하므로 검증 규칙을 둔다."),
    ("PSL", "필요한 속성",
     "psl_flag, psl_category(농업·MSME·수출·교육·주택·사회인프라·신재생·기타), psl_subcategory, msme_category(+분류일), psl_basis(직접·NBFC 경유·PSLC), "
     "weaker_section_flag, export_credit_stage, psl_eval_date(재판정 이력). 판정은 한도 승인 시 + 분기 말 재판정."),
    ("PSL", "목표 산정 기초",
     "ANBC(조정 순은행신용)와 CEOBE(난외 신용환산액) 중 큰 금액 기준. 따라서 난외 계정에도 CCF 산출 속성(bg_nature, lc_scope, tenor)이 필요하다. 외국은행 지점의 목표 비율·하위목표는 PSL MD 최신본 확인(VERIFY)."),
    ("PSL", "GL이 필요한 부분",
     "미달 시 NABARD·SIDBI 등 기금 예치(10490920000)와 그 이자(40130980000), PSLC 매매 수수료(40591070000/50510970000)는 실제 자금·손익이 생기므로 GL 신설."),
    ("PSL", "현재 설계와의 차이",
     "TF모델 TB_TF_LIMIT.PSL_CAT은 EXPORT/MSME/NONE 3값뿐 → 범주·하위범주·근거·판정일로 확장하고 여신 공통(고객·계좌) 속성으로 옮긴다. TReDS MSME 자동분류는 유지."),
    ("INR Nostro", "경우 ① 지점이 인도 내 타행에 INR 계좌 운용",
     "청산·NACH 스폰서·현금관리 등으로 INR 계좌를 타행에 두는 경우. Sch.7 I(인도 내 은행 당좌)이고 Form A '은행시스템 자산'으로 NDTL 산정에서 차감, NOP·외화 재평가 제외. "
     "FCY Nostro(Sch.7 II)와 반드시 구분 → L5 10105010100 + nostro_type=INR_SETTLEMENT_IN_INDIA."),
    ("INR Nostro", "경우 ② 인도 내 은행 외화계좌",
     "위치는 인도 내(Sch.7 I)지만 외화 → NOP·LBS 대상. L5 10105010200."),
    ("INR Nostro", "경우 ③ 해외 FCY Nostro",
     "Sch.7 II. 통화는 통화차원, Nostro별 하위원장. LBS용 국가·은행유형 속성, 마이너스 잔액은 Sch.4 II 재분류(20140200000)."),
    ("INR Nostro", "경우 ④ 본점·당행 해외지점이 노이다에 둔 INR 계좌",
     "본점이나 다른 국외지점 입장에서는 'INR Nostro'지만 노이다 장부에서는 같은 법인 → 예금이 아닌 본지점(21040000000). 다른 국외지점 계좌도 같은 원칙 적용 권고(이전안은 본점만 명시)."),
    ("INR Nostro", "경우 ⑤ 타 해외은행의 INR Vostro / SRVA",
     "노이다 장부에서 예금(Sch.3 은행분 요구불, 20150200100/200). Form A상 '인도 외 은행'은 은행시스템이 아니므로 cpty_bank_type으로 구분해 '기타 부채'로 산출(VERIFY)."),
    ("INR Nostro", "결론",
     "INR·FCY 구분은 필수이나 대부분 '위치 L4 + 통화차원 + nostro_type 속성'으로 해결된다. 인도 내 INR 결제계좌만 대사 단위가 달라 L5를 둔다. 본점 계좌는 본지점으로."),
]

# ---------------------------------------------------------------------------
# 8. 결정 필요 (번호, 결정 사항, 선택지, 권고, 결정 주체)
# ---------------------------------------------------------------------------
DECISIONS = [
    ("D1", "기준 계정코드", "① 11자리 통합코드 ② 8자리", "11자리 기준. 8자리는 이행 원천코드(시트 09)로만 사용", "결정(2026-11-05)"),
    ("D2", "예금 GL 결정키", "① 상품×거주성×통화 ② 상품×거주성×고객유형군(통합 L4)",
     "② (통합 L4가 고객유형 축). 현행 거주자 정기·당좌·EEFC는 이행 시 법인/개인 분할(시트 09 '분할')", "권고"),
    ("D3", "외국인 비거주자(FN) vs NRO", "① NRO GL 하나 + nr_subtype 부분합 ② FN 별도 GL ③ NRO 아래 L6",
     "① — FN은 NRO의 부분집합(FEMA). 재무제표·RBI 보고·세율이 같고, GL을 바꾸는 사건은 거주성 변경뿐(시트 07)", "권고"),
    ("D4", "금융/이행보증·내국LC", "① 속성 ② GL L5/L4", "① — RWA 명세(시트13 D)도 보증 목적코드로 CCF를 판정 → 속성과 일치", "권고"),
    ("D5", "정상자산 충당금", "ECL Directions 2026(2027-04-01) 기준 Stage별", "Stage별 L5 3개 + 부외 ECL은 0930 기존 계정 + 전환조정 자본(시트 10)", "결정(2026-11-05)"),
    ("D6", "INR 결제계좌 L5", "① L5 신설 ② 통화+속성만", "① (하위원장 대사·Form A 차감 근거 명확)", "권고"),
    ("D7", "규정 확인(VERIFY=Y)", "-", "PSL 목표, Form A 은행시스템 범위, LBS, 법정준비금·IFR, PSLC 손익 표시, s.11(2)(b) 예치 계좌, ECL 적용 범위(은행예치·HTM)", "미결"),
    ("E1", "은행 정기예치 ₹208.9 crore 계정", "① DUE FROM BANKS ON TIME(10106010000) ② 0930대로 CD PURCHASE", "① — Sch.7 '은행 기타예금'. 은행/비은행은 GL, 상대은행 유형·만기·운영 여부는 LCR·ECL 속성. CD 매입은 투자 계정", "권고"),
    ("L1", "LCR 분류 방식", "① GL(예금 L5를 개인/소규모/법인/금융으로) ② 고객·계좌 속성 + 일자 유효 파라미터", "② — Sch.3에 없는 구분이고 합산·부보·운영 여부로 보고일마다 변함(시트 11)", "권고"),
    ("L2", "LCR 파라미터 값", "-", "유출·유입률, 소규모사업자 한도, 디지털 가산, HQLA 인정비율을 현행 RBI Master Direction 원문으로 확정해 TB_LCR_* 적재", "미결"),
    ("E2", "0930 GBS코드 열 정정", "-", "시트 09 '수정' 57건을 본점 CoA 담당에 전달(8자리→11자리 이행표의 원본)", "요청"),
    ("E3", "ECL 명세서 신규코드", "① 11자리로 신설 ② 8자리 10690003~5", "① — 8자리 10690003~5는 현행 증권·사모사채·콜론 충당금이 이미 사용 중", "권고"),
    ("E4", "EIR 조정계정(10370300000)", "① 신설 ② LOF·LOC 상각만", "LOF·LOC 잔액이 0이면 ②로 시작 가능. ECL팀 확인", "미결"),
]


def collisions():
    """예금명세서 코드 ↔ 0930 GBS코드 열 비교."""
    by_gbs = {}
    for a in CM.BASE:
        if a["gbs"]:
            by_gbs.setdefault(a["gbs"], []).append(a)
    out = []
    for code, meaning, st in SPEC_CODES:
        hits = by_gbs.get(code, [])
        if not hits:
            res, coa = "0930에 없음", ""
        else:
            coa = " / ".join(f'{h["code"]} {h["name"].strip()}' for h in hits)
            res = "확인"
        out.append(dict(code=code, meaning=meaning, st=st, coa=coa, res=res))
    return out


# 사람이 판정한 충돌 결과 (자동 비교 결과에 덮어씀): 같은 계정이면 일치, 다른 계정이면 충돌
SAME = {"20910002", "20910003", "20920017", "10610004", "10610007", "10420002", "10420003", "10910003", "10910005",
        "10910009", "41020004", "41020007", "41050002", "41050003", "42020601", "42020701", "48090002", "53090012"}
# 코드는 같은 계정을 가리키나 용도가 맞지 않는 경우
CAUTION = {"42020729": "0930에서 대출 기타수수료. 에스크로 수수료를 여기에 넣으면 Sch.14 대출수수료로 집계됨 → 40592040000 신설 권고"}


def validate():
    """제안 코드 검증: 신규는 기존·이전안과 중복 금지, 상위코드 존재, 기존 대상(변경·매핑·비활성)은 존재해야 함."""
    known = set(CM.ALL)
    new = {r["code"] for r in R if r["kind"] == "신규"}
    errs = []
    for r in R:
        c = r["code"]
        if c == "-" or r["kind"] == "철회":
            continue
        if r["kind"] == "신규":
            if c in known:
                errs.append(f"신규 코드 중복: {c}")
            if r["parent"] and r["parent"] not in known and r["parent"] not in new:
                errs.append(f"상위코드 없음: {c} → {r['parent']}")
        elif c not in known:
            errs.append(f"대상 계정 없음: {r['kind']} {c}")
    return errs


if __name__ == "__main__":
    e = validate()
    print("검증 오류:", e or "없음")
    print("검토 항목:", len(R), "속성:", len(ATTRS), "보고서:", len(RETURNS))
    col = collisions()
    print("충돌:", sum(1 for x in col if x["res"] == "확인" and x["code"] not in SAME and x["code"] not in CAUTION), "/ 없음:", sum(1 for x in col if x["res"] != "확인"))
