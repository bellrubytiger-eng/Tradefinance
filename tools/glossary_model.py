# -*- coding: utf-8 -*-
"""인도 특수 용어 정의 — 본점 검토자용.

각 시트 표의 행 텍스트에서 아래 패턴을 찾아 '인도 용어 정의' 열에 정의를 붙이고, 전체 목록은 시트 17.
(용어, 분류, 정의, 근거, 탐지 정규식 목록)
"""
import re

G = [
    # 여신·PSL
    ("PSL (Priority Sector Lending)", "여신", "우선부문 대출. RBI가 정한 농업·MSME·수출신용·교육·주택·사회인프라·신재생 등에 대한 대출로, 은행은 ANBC 또는 CEOBE 중 큰 금액의 일정 비율(국내은행·외국은행 모두 기본 40%)을 의무 공급해야 함. 미달 시 NABARD 등 기금 예치",
     "RBI Master Directions – Priority Sector Lending", [r"\bPSL\b", r"우선부문"]),
    ("ANBC / CEOBE", "여신", "ANBC(조정 순은행신용) = 인도 내 은행신용에서 일부 조정한 금액. CEOBE = 난외 익스포저의 신용환산액. PSL 목표는 둘 중 큰 금액 기준", "PSL MD", [r"\bANBC\b", r"\bCEOBE\b"]),
    ("PSLC", "여신", "PSL 증서. 목표 초과 은행이 RBI e-Kuber에서 매도, 미달 은행이 매입해 목표 달성에 산입. 매매 수수료가 손익", "RBI PSLC 지침", [r"\bPSLC\b"]),
    ("RIDF / PSL 미달 예치", "여신", "PSL 미달분을 NABARD(RIDF)·SIDBI·NHB·MUDRA 기금에 저금리로 예치하는 제도", "PSL MD", [r"\bRIDF\b", r"PSL 미달", r"SHORTFALL"]),
    ("MSME / Udyam", "여신", "중소기업. 투자·매출 기준으로 Micro/Small/Medium 구분, Udyam 등록번호로 확인. PSL·충당금 하한·위험가중치에 영향", "MSMED Act 2006", [r"\bMSME\b", r"\bUdyam\b", r"UDYAM"]),
    ("Packing Credit / PCFC", "여신", "선적 전 수출금융. 원화(INR) 패킹크레딧과 외화(PCFC)로 나뉘며 수출대금으로 상환", "RBI 수출신용 지침", [r"PACKING CREDIT", r"\bPCFC\b", r"패킹크레딧"]),
    ("TReDS", "여신", "MSME 매출채권 전자할인 플랫폼(RXIL·M1xchange·Invoicemart). 무소구 할인, MSME PSL", "RBI TReDS 지침", [r"\bTReDS\b", r"TREDS"]),
    ("ECGC", "여신", "인도 수출신용보증공사. 수출금융 보증", "-", [r"\bECGC\b"]),
    ("CC / OD / WCDL", "여신", "CC(Cash Credit)·OD(Overdraft): 한도 내 수시인출 운전자금. WCDL: 운전자금 요구불대출(만기 일시상환)", "-", [r"\bCC/OD\b", r"CASH CREDIT", r"\bWCDL\b", r"당좌대월"]),
    ("Out-of-order", "여신", "CC/OD 계정이 한도초과 90일·무입금 90일·이자 미충당 상태 → NPA 판정 기준", "IRACP / ACPIR §6(17)", [r"[Oo]ut-of-order"]),
    ("IRAC / IRACP", "여신", "현행 자산건전성 분류·충당금·이자인식 지침(Standard·Sub-standard·Doubtful·Loss). 2027-04-01 ECL로 대체", "RBI IRACP Directions 2025", [r"\bIRAC", r"\bIRACP\b"]),
    ("NPA", "여신", "무수익여신. 원리금 90일 초과 연체 등. ECL Stage 3과 같음. 이자는 현금주의", "IRACP", [r"\bNPA\b"]),
    ("SMA", "여신", "조기경보 분류. SMA-0(1~30일)·SMA-1(31~60일)·SMA-2(61~90일) 연체", "RBI 조기경보 체계", [r"\bSMA\b", r"SMA-"]),
    ("ECL Directions 2026 / ACPIR", "여신", "RBI 기대신용손실 기준(자산분류·충당금·이자인식). 2027-04-01 시행, Stage 1·2·3과 규제 하한(floor)", "RBI/DOR/2026-27/398", [r"\bECL\b", r"\bACPIR\b"]),
    ("SICR / Stage", "여신", "신용위험 유의적 증가. Stage 1(12개월 ECL)·Stage 2(SICR, 전체기간)·Stage 3(손상=NPA)", "ECL Dir. para 25~37", [r"\bSICR\b", r"STAGE [123]", r"Stage [123]"]),
    ("UFCE / EBID", "여신", "차주의 미헤지 외화익스포저. 잠재손실/EBID(세후이익+감가상각+이자+리스료) 비율에 따라 추가 충당금·위험가중치", "RBI CRM Directions Ch.VII", [r"\bUFCE\b", r"\bEBID\b"]),
    ("CRILC", "보고", "대형신용정보 중앙저장소. 그룹 합산 ₹5 crore 이상 차주의 익스포저·SMA를 RBI에 보고", "RBI CRILC 지침", [r"\bCRILC\b"]),
    ("CIC", "보고", "신용정보회사(CIBIL 등). 전 차주 신용정보를 격주 제출", "CICRA 2005", [r"\bCIC\b"]),
    ("LEF", "보고", "거액익스포저 한도 체계(동일인·그룹)", "RBI LEF", [r"\bLEF\b"]),
    # 예금·FEMA
    ("NRE", "예금", "비거주 인도인(NRI·OCI)의 INR 예금. 원리금 해외송금 자유, 이자 비과세", "FEMA (Deposit) Regulations 2016", [r"\bNRE\b"]),
    ("NRO", "예금", "비거주자의 인도 내 소득 수취용 INR 예금. 해외송금 연 USD 1M 한도, 이자 원천징수(195). 외국인 비거주자 계좌도 NRO", "FEMA (Deposit) Regs", [r"\bNRO\b"]),
    ("FN (Foreign National)", "예금", "외국 국적 비거주자 개인. INR 예금은 NRO만 가능(NRE 불가)", "FEMA", [r"\bFN\b", r"외국인 비거주자"]),
    ("FCNR(B)", "예금", "비거주자 외화정기예금(1~5년), 금리상한(ARR+스프레드), 이자 비과세", "FEMA (Deposit) Regs", [r"FCNR"]),
    ("RFC", "예금", "귀국 거주자의 외화예금", "FEMA", [r"\bRFC\b"]),
    ("EEFC", "예금", "거주자 외화수입 보유 계정(무이자). 수출·용역대금 일부를 외화로 보유", "FEMA", [r"\bEEFC\b"]),
    ("SNRR", "예금", "특별 비거주자 루피 계정. 비거주자의 인도 내 사업·투자 결제용 INR 계정", "FEMA 5(R)", [r"\bSNRR\b"]),
    ("SRVA / Vostro / RDA", "예금", "SRVA: 해외은행이 INR 무역결제용으로 인도 은행에 둔 특별 루피 Vostro. RDA: 해외 송금회사용 루피 지급 약정", "RBI 2022.7 INR 무역결제 체계", [r"\bSRVA\b", r"\bRDA\b", r"[Vv]ostro", r"VOSTRO"]),
    ("LO / BO / PO, UIN, AAC", "예금", "외국기업의 연락·지점·프로젝트 사무소. UIN(RBI 고유번호), AAC(연차 활동증명서)", "FEMA 22(R)/2016", [r"LIAISON", r"BRANCH OFFICE", r"PROJECT OFFICE", r"\bAAC\b", r"\bUIN\b"]),
    ("Collection Account (Ch.XIA)", "예금", "여신 점유율 10% 미만 은행이 차주에게 열어주는 집금 전용 당좌. 2영업일 내 지정계좌 스윕", "RBI CRM Directions 2025 Ch.XIA", [r"COLLECTION ACCOUNT", r"Collection [Aa]ccount", r"집금"]),
    ("BSBDA", "예금", "기본저축예금(최저잔액 없음, 1인 1계좌)", "RBI BSBDA Directions", [r"\bBSBDA"]),
    ("Bulk Deposit", "예금", "₹3 crore 이상 정기예금. 금액별 차등금리 허용, 매일 공시", "RBI 예금금리 지침", [r"Bulk", r"BULK"]),
    ("DICGC", "예금", "예금보험공사. 예금자·은행별 ₹5 lakh(원리금) 보호, 반기 보험료", "DICGC Act 1961", [r"\bDICGC\b", r"DEPOSIT INSURANCE"]),
    ("DEAF", "예금", "10년 이상 미청구 예금을 RBI 기금으로 이전, 청구 시 은행 지급 후 RBI에 청구", "BR Act s.26A", [r"\bDEAF\b"]),
    ("Escrow (FEMA / RERA / PA)", "예금", "FEMA 주식거래 에스크로(INR·무이자), RERA 부동산 분양대금 지정계좌, PA(결제대행업자) 에스크로", "FEMA Sch.5, RERA 2016, PA Dir.", [r"[Ee]scrow", r"ESCROW", r"\bRERA\b"]),
    ("FEMA / AD Category-I", "외환", "외환관리법. AD Category-I = 모든 외환업무 인가 은행(노이다 지점)", "FEMA 1999", [r"\bFEMA\b", r"\bAD\b Category"]),
    ("LRS", "외환", "거주 개인 해외송금 한도 연 USD 250,000. 송금 시 TCS", "FEMA LRS MD", [r"\bLRS\b"]),
    ("EDPMS / IDPMS", "외환", "RBI 수출·수입대금 추적 시스템(선적·통관 대비 결제 대사)", "RBI", [r"\bEDPMS\b", r"\bIDPMS\b"]),
    ("FIRC / 15CA·15CB", "외환", "FIRC: 타발송금 입금 확인서. 15CA/15CB: 비거주자 송금 시 세금 신고·회계사 증명", "IT Act", [r"\bFIRC\b", r"15CA", r"15CB"]),
    ("R-Return / FETERS / CIMS", "보고", "외환거래 보고(Purpose code별). CIMS = RBI 통합 보고 시스템", "RBI", [r"R-Return", r"FETERS", r"\bCIMS\b"]),
    ("NRD-CSR / FLA / LBS", "보고", "NRD-CSR: 비거주자 예금 월보. FLA: 외국인 부채·자산 연보. LBS: 대외·외화 포지션 국제은행통계", "RBI", [r"NRD-CSR", r"\bFLA\b", r"\bLBS\b"]),
    ("ECB / FPI / LEI", "외환", "ECB: 비거주자로부터의 외화차입. FPI: SEBI 등록 외국 포트폴리오 투자자. LEI: 법인식별기호(₹50 crore 이상 거래 필수)", "FEMA·SEBI·RBI", [r"\bECB\b", r"\bFPI\b", r"\bLEI\b"]),
    # 규제·재무제표
    ("Third Schedule / Sch.", "재무제표", "BR Act s.29 은행 재무제표 서식(Form A 대차대조표·Form B 손익). Sch.1~18이 계정 줄", "BR Act 1949 s.29", [r"Third Schedule", r"Sch\.\s?\d"]),
    ("Form A / CRR / SLR / NDTL", "규제", "Form A: RBI Act s.42 격주 보고. CRR: NDTL 대비 RBI 현금 예치. SLR: NDTL 대비 국채 등 보유. NDTL: 순요구불·정기부채", "RBI Act s.42, BR Act s.24", [r"Form A", r"\bCRR\b", r"\bSLR\b", r"\bNDTL\b"]),
    ("s.11(2)(b) 예치", "규제", "외국은행이 RBI에 예치해야 하는 금액(현금 또는 담보 증권). Sch.1에 주기, SLR·HQLA 제외", "BR Act s.11(2)(b)", [r"s\.11\(2\)", r"S11_2", r"11\(2\)\(b\)"]),
    ("Statutory Reserve / IFR", "규제", "법정준비금(순이익 25% 적립), 투자변동준비금(AFS·HFT 보유분 기준)", "BR Act s.17, RBI", [r"STATUTORY RESERVE", r"Statutory Reserve", r"법정준비금", r"\bIFR\b"]),
    ("HTM / AFS / HFT·FVTPL", "규제", "RBI 투자 분류(만기보유·매도가능·단기매매). IFRS의 AC·FVOCI·FVPL과 유사하나 규정이 다름", "RBI 투자 지침 2023", [r"\bHTM\b", r"\bAFS\b", r"\bHFT\b", r"FVTPL"]),
    ("LCR / HQLA / FALLCR / MSF", "규제", "LCR: 30일 유동성 커버리지(최저 100%). HQLA: 고유동성자산. FALLCR·MSF: SLR 보유분 중 HQLA 인정분", "RBI Basel III LCR", [r"\bLCR\b", r"\bHQLA\b", r"FALLCR", r"\bMSF\b"]),
    ("BSR", "보고", "Basic Statistical Returns. 계좌 단위 여신(BSR-1)·예금(BSR-2)·분기 집계(BSR-7)", "RBI", [r"\bBSR\b", r"BSR-"]),
    ("CCIL / TREPS", "자금", "CCIL: 인도 청산결제기관(외환·국채·단기자금). TREPS: CCIL 담보부 콜 시장", "-", [r"\bCCIL\b", r"\bTREPS\b"]),
    ("Call / Notice / Term money", "자금", "은행간 무담보 자금: 콜(1일)·노티스(2~14일)·텀(15일~1년)", "RBI", [r"Notice money", r"Term money", r"Call money", r"CALL LOAN"]),
    ("NEFT / RTGS / SFMS / NACH", "자금", "인도 지급결제망: NEFT(이연 일괄)·RTGS(실시간 총액)·SFMS(은행간 메시지, 국내 LC)·NACH(자동이체)", "RBI", [r"\bNEFT\b", r"\bRTGS\b", r"\bSFMS\b", r"\bNACH\b"]),
    ("Penal Charges", "여신", "연체 시 이자가 아닌 수수료로만 부과(복리 불가)", "RBI 공정대출관행(Penal Charges) 2023", [r"[Pp]enal [Cc]harge", r"PENAL CHARGE"]),
    # 세무
    ("TDS", "세무", "원천징수. 지급자가 이자·용역대가 등에서 세금을 떼어 정부에 납부. 예금이자는 계좌 입금(미지급이자 계상 포함) 시점이 의무 시점", "Income-tax Act 2025 s.392~393(구 192~195)", [r"\bTDS\b"]),
    ("TCS", "세무", "원천징수(징수형). 판매자·송금은행이 대금과 함께 세금을 받아 납부. LRS 송금·해외여행·차량 등", "IT Act 2025 s.394(구 206C)", [r"\bTCS\b"]),
    ("24Q / 26Q / 27Q / 27EQ / Form 140", "세무", "TDS·TCS 분기 신고서: 24Q 급여, 26Q 거주자 기타, 27Q 비거주자, 27EQ TCS. 신법에서 서식번호 변경(26Q→Form 140 등)", "IT Act", [r"24Q", r"26Q", r"27Q", r"27EQ", r"Form 140"]),
    ("Challan (ITNS-281)", "세무", "TDS·TCS 납부서. 납부 성격 코드별로 납부(익월 7일, 3월분 4/30)", "IT Rules", [r"[Cc]hallan", r"챌린", r"ITNS"]),
    ("15G / 15H / Form 121", "세무", "소득이 과세 기준 미만인 예금자의 원천징수 면제 신고(60세 미만 15G, 이상 15H). 신법 Form 121", "IT Act", [r"15G", r"15H", r"Form 121"]),
    ("26AS / AIS", "세무", "국세청 세액공제 명세. 원천징수당한 세액 확인·대사", "IT Act", [r"26AS", r"\bAIS\b"]),
    ("Senior citizen (TDS)", "세무", "60세 이상 예금자: 예금이자 TDS 한도 ₹1 lakh(일반 ₹50,000)", "IT Act", [r"[Ss]enior", r"SENIOR", r"고령"]),
    ("GST / CGST·SGST·IGST", "세무", "부가가치세. 주 내 거래 = CGST 9% + SGST 9%, 주 간·해외 = IGST 18%. 은행 용역 18%, 이자는 비과세", "CGST·IGST Act 2017", [r"\bGST\b", r"CGST", r"SGST", r"IGST"]),
    ("RCM", "세무", "역발행. 수입 용역·특정 국내 용역은 구매자(은행)가 GST를 납부", "CGST Act s.9(3)·(4)", [r"\bRCM\b"]),
    ("ITC (50% option, s.17(4))", "세무", "매입세액 공제. 은행은 적격 ITC의 50%만 공제하는 선택 가능 — 나머지는 불공제 비용", "CGST Act s.17(4)", [r"\bITC\b", r"17\(4\)"]),
    ("GSTR-1 / 3B / 2B / ISD", "세무", "GST 신고: GSTR-1 매출 명세, 3B 납부, 2B 매입세액 자동명세, ISD 공통비용 배분", "CGST Rules", [r"GSTR", r"\bISD\b"]),
    ("SAC / POS / GSTIN", "세무", "SAC: 용역 분류코드(은행 9971). POS: 공급지(세목 결정). GSTIN: 주별 GST 등록번호", "CGST Act", [r"\bSAC\b", r"\bPOS\b", r"GSTIN"]),
    ("Rule 32(2)(b)", "세무", "환전 GST 과세가액 산식(환전금액 구간별 일정 비율)", "CGST Rules", [r"32\(2\)\(b\)"]),
    ("Pure agent (Rule 33)", "세무", "고객 대신 실비를 지급하고 그대로 회수하는 경우 과세가액에서 제외. 회수액이 실비와 다르면 해당 안 됨", "CGST Rules r.33", [r"[Pp]ure agent", r"PURE AGENT"]),
    ("Income-tax Act 2025", "세무", "2026-04-01 시행 신 소득세법. TDS 192→392, 193~195→393, TCS 206C→394로 조문 재편", "Income-tax Act 2025", [r"소득세법 2025", r"신법", r"(?<![\d,.])39[234](?![\d,])"]),
    # 회계·인사
    ("AS (Indian GAAP) / Ind AS", "회계", "인도 은행은 Ind AS 적용이 연기되어 AS(Companies (Accounting Standards) Rules)를 따름. AS 15 종업원급여, AS 19 리스, AS 22 이연법인세, AS 26 무형자산", "RBI, MCA", [r"\bAS 1[59]\b", r"\bAS 2[26]\b", r"Ind AS"]),
    ("Schedule II (Companies Act)", "회계", "자산별 내용연수 기준(컴퓨터 3년·서버 6년·가구 10년·차량 8년·건물 60년, 잔존가치 5%)", "Companies Act 2013 Sch.II", [r"Schedule II"]),
    ("세무상각 블록(WDV)", "세무", "소득세법 감가상각: 자산 블록별 정률(컴퓨터·SW 40%, 기계 15%, 가구 10%), 사용 180일 미만이면 반액", "IT Act", [r"WDV", r"블록"]),
    ("Gratuity / Leave encashment", "인사", "Gratuity: 5년 이상 근속자 퇴직금(Payment of Gratuity Act 1972). 연차수당: 미사용 연차 현금보상. 둘 다 AS 15 보험수리 평가", "Gratuity Act, AS 15", [r"[Gg]ratuity", r"GRATUITY", r"LEAVE ENCASHMENT", r"연차수당"]),
    ("EPF / EPS / EDLI / ESIC", "인사", "EPF: 퇴직적립(사용자·근로자 각 12%), EPS: 연금, EDLI: 사망보험, ESIC: 의료보험. 익월 15일 납부", "EPF Act 1952, ESI Act 1948", [r"\bEPF\b", r"\bEPS\b", r"\bEDLI\b", r"\bESIC\b", r"\bPF\b"]),
    ("Bonus Act", "인사", "Payment of Bonus Act: 일정 급여 이하 근로자 법정상여(8.33~20%)", "Payment of Bonus Act 1965", [r"Bonus Act", r"법정 상여", r"법정\(Payment"]),
    ("CSR (s.135)", "회계", "순이익 ₹5 crore 이상 등 요건 충족 시 직전 3년 평균 순이익 2%를 Schedule VII 활동에 지출. 미집행분은 별도 계좌·기금 이전", "Companies Act 2013 s.135", [r"\bCSR\b"]),
    ("LFAR", "회계", "Long Form Audit Report. 은행 법정감사인의 상세 보고서", "RBI", [r"\bLFAR\b"]),
    ("crore / lakh", "단위", "1 lakh = 100,000(10만), 1 crore = 10,000,000(1천만) = 100 lakh", "-", [r"crore", r"lakh"]),
    ("IBA / FEDAI / FIMMDA", "기타", "인도은행협회, 외환딜러협회, 채권·단기금융시장 딜러협회 — 은행 회비 대상", "-", [r"\bIBA\b", r"FEDAI", r"FIMMDA"]),
]

_COMPILED = [(t, d, b, [re.compile(p) for p in pats]) for t, _c, d, b, pats in G]


def terms_in(text):
    """텍스트에 나오는 용어의 정의 목록."""
    text = str(text)
    out = []
    for t, d, b, pats in _COMPILED:
        if any(p.search(text) for p in pats):
            out.append(f"■ {t}: {d}")
    return out
