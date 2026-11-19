# -*- coding: utf-8 -*-
"""해외지점 공통 모듈: 한국 기준 ↔ 7개국 비교 맵핑 (기초 조사 초안, 2026-11).
각 칸의 확인 수준: 확인(공식·다수 출처 일치) / 2차(법률·은행 자료) / 확인필요(출처 부족·지점 확인 필요)."""

BRANCHES = [
    # 지점, 국가, 코드, 현지통화, 감독·외환당국, 결제망, 외환통제, 특화 강도, 메모
    ("노이다", "인도", "IN", "INR", "RBI (FEMA)", "RTGS·NEFT (SFMS)", "강함", "높음", "노이다 요건 완료. 공통화 기준 지점"),
    ("북경", "중국", "CN", "CNY", "PBoC·SAFE·NFRA", "CNAPS·CIPS", "강함", "높음", "국제수지 신고·무역 진실성 심사·서비스 대외지급 세무등록"),
    ("하노이", "베트남", "VN", "VND", "SBV", "CITAD·NAPAS", "강함", "높음", "목적별 송금·증빙 보관, 수출대금 국내 입금 의무"),
    ("홍콩", "홍콩", "HK", "HKD", "HKMA", "CHATS·FPS", "없음", "낮음", "AML 송금정보(HKD 8,000), RMB 업무"),
    ("뉴욕", "미국", "US", "USD", "Fed·OCC/NYDFS·OFAC", "Fedwire·CHIPS", "없음", "중간", "개인 송금 소비자보호(Reg E Subpart B), BSA·OFAC"),
    ("런던", "영국", "GB", "GBP", "PRA·FCA·OFSI", "CHAPS·FPS", "없음", "낮음", "CHAPS ISO 20022·하이브리드 주소, UK 제재목록"),
    ("시드니", "호주", "AU", "AUD", "APRA·AUSTRAC", "RITS·NPP", "없음", "중간", "AUSTRAC IFTI 보고(10영업일), 2029 IVTS 전환"),
]
COLS = ["KR", "IN", "CN", "HK", "VN", "US", "GB", "AU"]
COL_NM = {"KR": "한국 (기준)", "IN": "인도 (노이다)", "CN": "중국 (북경)", "HK": "홍콩", "VN": "베트남 (하노이)", "US": "미국 (뉴욕)", "GB": "영국 (런던)", "AU": "호주 (시드니)"}

# 비교 항목: (영역, 항목, 처리구분, {국가: (내용, 확인수준)}, 설계 반영)
# 처리구분: 공통 = 모든 지점 같은 로직 / 정책값 = 같은 로직, 값만 지점별 / 확장 = 해당 국가만 추가 항목·규칙
ROWS = [
    ("외환 규제", "외환통제 수준", "정책값", {
        "KR": ("원칙 자유, 증빙·신고 기반(외국환거래법)", "확인"), "IN": ("FEMA 통제, AD 은행이 증빙 확인", "확인"),
        "CN": ("SAFE 통제, 은행이 진실성 심사", "확인"), "VN": ("SBV 통제, 은행이 증빙 확인·보관", "2차"),
        "HK": ("통제 없음", "확인"), "US": ("통제 없음 (제재·AML 중심)", "확인"), "GB": ("통제 없음 (제재·AML 중심)", "확인"), "AU": ("통제 없음 (AML·보고 중심)", "확인")},
     "지점정책 '외환통제 모드'로 증빙·한도 탭 노출 결정"),
    ("개인 송금", "개인 해외송금 한도", "정책값", {
        "KR": ("증빙 없이 연 USD 10만 (2023.7.4~), 초과 시 증빙", "확인"), "IN": ("LRS 회계연도 USD 25만 (4~3월)", "확인"),
        "CN": ("개인 외화 구매 연 USD 5만 편의한도(역년), 투자 목적 불가", "2차"), "VN": ("목적별 송금만 허용(유학·의료·친족부양 등), 목적별 상한", "2차"),
        "HK": ("한도 없음", "확인"), "US": ("한도 없음", "확인"), "GB": ("한도 없음", "확인"), "AU": ("한도 없음", "확인")},
     "공통 한도 엔진(고객구분·목적별 누계, 집계기간·통화·상한)을 정책값으로"),
    ("개인 송금", "한도 집계기간", "정책값", {
        "KR": ("역년", "확인"), "IN": ("회계연도 4/1~3/31", "확인"), "CN": ("역년", "2차"), "VN": ("목적·건별 (연간 기준 일부)", "확인필요"),
        "HK": ("해당 없음", "확인"), "US": ("해당 없음", "확인"), "GB": ("해당 없음", "확인"), "AU": ("해당 없음", "확인")},
     "IT 설계서 '회계연도 시작월·집계기간' 정책 항목과 동일"),
    ("개인 송금", "개인 송금 고객보호", "확장", {
        "KR": ("-", "확인"), "IN": ("-", "확인"), "CN": ("-", "확인필요"), "VN": ("-", "확인필요"), "HK": ("-", "확인필요"),
        "US": ("CFPB Remittance Transfer Rule(Reg E Subpart B): 사전 고지·영수증, 결제 후 30분 내 취소·환불", "확인"),
        "GB": ("Payment Services Regulations 정보 제공 의무", "확인필요"), "AU": ("-", "확인필요")},
     "US 확장: RT11 개인송금에 사전고지서·영수증 출력, 30분 취소 상태 관리"),
    ("기업 무역송금", "수입대금 송금 증빙 확인", "정책값", {
        "KR": ("지급사유 증빙 확인, 송금방식 수출입 사후관리(국내 2406)", "확인"), "IN": ("Invoice·Bill of Entry, IDPMS 사후관리", "확인"),
        "CN": ("계약·인보이스·세관신고서로 진실성 심사, 무역기업 명부 등록(2024.6.1부터 은행 처리), 건당 USD 10만 이하 은행 재량(2017 기준)", "2차"),
        "VN": ("은행이 실거래 증빙 확인·보관(Circular 20/2022 등), 무역대금은 은행계좌 경유", "2차"),
        "HK": ("규제상 의무 없음 (무역기반 자금세탁 위험평가)", "2차"), "US": ("규제상 의무 없음 (BSA 위험기반)", "2차"),
        "GB": ("규제상 의무 없음 (위험기반)", "2차"), "AU": ("규제상 의무 없음 (위험기반)", "2차")},
     "증빙 요건 매트릭스(필수/옵션/불가)를 지점·고객구분·송금종류별로"),
    ("기업 무역송금", "세관신고 문서", "정책값", {
        "KR": ("수입신고필증 / 수출신고필증", "확인"), "IN": ("Bill of Entry / Shipping Bill", "확인"), "CN": ("수출입화물 통관신고서(报关单)", "확인"),
        "VN": ("세관신고서(Tờ khai hải quan)", "확인"), "HK": ("수출입 신고(TDEC) — 은행 확인 의무 없음", "확인필요"),
        "US": ("CBP 7501 Entry Summary / 수출 EEI — 은행 확인 의무 없음", "확인필요"), "GB": ("CDS 신고(MRN) — 은행 확인 의무 없음", "확인필요"),
        "AU": ("Import Declaration(N10) / EDN — 은행 확인 의무 없음", "확인필요")},
     "공통 증빙유형 '세관신고(CUSTOMS)' + 국가별 명칭·번호형식 코드"),
    ("기업 무역송금", "수입 선급 사후관리 기한", "정책값", {
        "KR": ("송금방식 수입 사후관리 기한 — 원문 확인 필요", "확인필요"), "IN": ("계약 기한·AD 재량 (2026 규정)", "2차"),
        "CN": ("화물무역 모니터링(고정 기한 확인 필요)", "확인필요"), "VN": ("확인 필요", "확인필요"),
        "HK": ("해당 없음", "확인"), "US": ("해당 없음", "확인"), "GB": ("해당 없음", "확인"), "AU": ("해당 없음", "확인")},
     "RT74 사후관리(Due/Receive Date)를 정책 '사후관리 사용 여부·기한'으로"),
    ("송금방식 수출", "수출대금 회수 의무", "정책값", {
        "KR": ("대외채권 회수의무 2017년 전후 폐지(보유현황 보고로 대체)", "2차"), "IN": ("선적 후 9개월(INR 12개월), EDPMS", "2차"),
        "CN": ("수령 시 국제수지 신고, 기한 확인 필요", "확인필요"), "VN": ("수출대금은 국내 외화계좌 입금 의무(외환법령·Decree 70/2014)", "2차"),
        "HK": ("없음", "확인"), "US": ("없음", "확인"), "GB": ("없음", "확인"), "AU": ("없음", "확인")},
     "XP73 수출서류 ↔ 타발 매칭, 회수기한은 정책값(없으면 미사용)"),
    ("보고·코드", "송금 목적·국제수지 코드", "확장", {
        "KR": ("지급등의 사유 코드 (한국은행 보고)", "확인"), "IN": ("RBI Purpose Code (S=지급, P=수취)", "확인"),
        "CN": ("국제수지 신고 거래코드(은행 경유), CNY 송금 목적코드", "2차"), "VN": ("송금 목적 기재 — 코드 체계 확인 필요", "확인필요"),
        "HK": ("없음", "2차"), "US": ("없음", "2차"), "GB": ("없음", "2차"), "AU": ("없음 (IFTI 보고 항목은 별도)", "2차")},
     "Purpose Code 필드 공통 + 코드셋은 지점정책(IT 설계서와 동일)"),
    ("보고·코드", "당국 거래보고", "확장", {
        "KR": ("한국은행 외환전산망", "확인"), "IN": ("IDPMS·EDPMS·CIMS·FETERS", "확인"), "CN": ("국제수지 신고·화물무역 모니터링 시스템", "2차"),
        "VN": ("SBV 정기보고", "2차"), "HK": ("의심거래 보고(JFIU)", "2차"), "US": ("BSA(CTR·SAR), OFAC 보고", "2차"),
        "GB": ("의심거래 보고(NCA), 제재 보고(OFSI)", "2차"), "AU": ("AUSTRAC IFTI 10영업일 내 (2029.3.31부터 IVTS로 전환 예정)", "확인")},
     "보고 인터페이스는 국가별 확장 모듈(송금 원장은 공통)"),
    ("세금", "송금 관련 세금·세무 확인", "확장", {
        "KR": ("-", "확인"), "IN": ("TCS(LRS), 환전 GST, 15CA/15CB", "확인"),
        "CN": ("서비스 등 대외지급 건당 USD 5만 초과 시 세무 등록서 제출", "2차"), "VN": ("외국계약자세(FCT) 확인 필요", "확인필요"),
        "HK": ("-", "2차"), "US": ("-", "2차"), "GB": ("-", "2차"), "AU": ("-", "2차")},
     "세금 항목은 확장필드(IN-TRD처럼 CN·VN 세트)"),
    ("송금정보", "송금인·수취인 정보 (Travel rule)", "공통", {
        "KR": ("특정금융정보법 송금정보", "2차"), "IN": ("RBI KYC Directions", "2차"), "CN": ("인민은행 AML 규정", "확인필요"), "VN": ("SBV AML 규정", "확인필요"),
        "HK": ("HKD 8,000 이상 송금인 정보 완비·검증", "2차"), "US": ("USD 3,000 이상 Travel Rule (31 CFR 1010.410)", "2차"),
        "GB": ("Funds Transfer Regulation 2017", "확인필요"), "AU": ("AML/CTF Act 송금정보", "확인필요")},
     "공통: 송금인 정보 완비 검증(기준금액은 정책값)"),
    ("SWIFT·결제망", "ISO 20022·구조화 주소", "공통", {
        "KR": ("SWIFT CBPR+ 2026.11 하이브리드/구조화 주소", "확인"), "IN": ("CBPR+ 동일, RTGS ISO 기반", "확인"),
        "CN": ("CBPR+ 동일, CIPS ISO 기반", "확인필요"), "VN": ("CBPR+ 동일", "확인필요"), "HK": ("CBPR+ 동일", "확인필요"),
        "US": ("Fedwire ISO 20022 전환 2025.7.14 완료, 2026.11 하이브리드 주소", "확인"),
        "GB": ("CHAPS ISO 20022(2023.6), 2025.11 하이브리드 주소 규칙, 2027 enhanced data 확대", "확인"), "AU": ("CBPR+ 동일", "확인필요")},
     "구조화 주소 입력·검증은 전 지점 공통(현행 Default 탭 보강)"),
    ("SWIFT·결제망", "송금 경로 (Direct / Serial / COV)", "공통", {
        k: ("Nostro·RMA·SSI 기반 판정 동일, 다른 것은 Nostro·마감시각·휴일뿐", "확인") for k in COLS},
     "노이다 경로판정 패널을 Bank Information 개선안으로 공통 적용"),
    ("제재·AML", "제재 목록", "정책값", {
        "KR": ("한국 제재(외교부·기재부) + UN", "2차"), "IN": ("UN·인도 지정", "2차"), "CN": ("UN·중국 반외국제재법", "확인필요"), "VN": ("UN·SBV", "확인필요"),
        "HK": ("UN 제재", "2차"), "US": ("OFAC SDN", "확인"), "GB": ("UK Sanctions List (OFSI)", "2차"), "AU": ("DFAT 통합 목록", "2차")},
     "필터링 엔진 공통 + 적용 목록 세트는 지점정책"),
    ("기타", "LEI 요구", "정책값", {
        "KR": ("-", "확인필요"), "IN": ("INR 50 crore 이상 거래 필수", "확인"), "CN": ("-", "확인필요"), "VN": ("-", "확인필요"),
        "HK": ("-", "확인필요"), "US": ("-", "확인필요"), "GB": ("CHAPS 금융기관 간 지급 LEI 요구(BoE)", "2차"), "AU": ("-", "확인필요")},
     "LEI는 공통 선택 항목 + 필수 조건은 정책값"),
    ("기타", "규제 연도 (회계·보고)", "정책값", {
        "KR": ("역년", "확인"), "IN": ("4월~3월", "확인"), "CN": ("역년", "확인"), "VN": ("역년", "확인"),
        "HK": ("4월~3월", "확인"), "US": ("역년 (연방 회계 10월 시작, 은행 보고와 무관)", "2차"), "GB": ("4월~3월 (세무 4/6)", "2차"), "AU": ("7월~6월", "확인")},
     "조회 기본기간·한도 집계 시작월을 정책값으로 (현행 4/1 고정 제거)"),
]

SOURCES = [
    ("KR", "증빙 없는 해외송금 한도 USD 10만 (2023.7.4)", "https://www.asiae.co.kr/en/article/2023060808321097112"),
    ("KR", "증빙 없는 송금 한도 상향 (KED)", "https://www.kedglobal.com/newsView/ked202302100014"),
    ("KR", "대외채권 회수의무 폐지", "https://www.dt.co.kr/article/10420906"),
    ("CN", "개인 외화 구매 연 USD 5만 (HSBC 중국 안내)", "https://www.hsbc.com.cn/content/dam/hsbc/cn/docs/foreign-exchange/personal-foreign-exchange-business-handling-guide-en.pdf"),
    ("CN", "개인 FX 편의한도·용도 (Hankun)", "https://hankunlaw.com/upload/portal/20260728/5e001c827598c3d4179722d27b57336f.pdf"),
    ("CN", "무역 외환 2024 개편 (SAFE 회람 요약)", "https://www.chinalawandpractice.com/2024/06/05/state-administration-of-foreign-exchange-circular-on-further-optimizing-exchange-control-relating-to-trade"),
    ("CN", "무역기업 명부 은행 처리 (CIB)", "https://www.cib.com.cn/en/aboutCIB/about/news/2024/20240710_3.html"),
    ("CN", "건당 USD 10만 이하 은행 재량 확인 (SAFE 2017)", "https://www.safe.gov.cn/en/2017/0404/1306.html"),
    ("CN", "서비스 대외지급 세무등록 USD 5만 (선전 세무)", "https://www.sz.gov.cn/en_szgov/services/enterprise/taxation/content/post_11478079.html"),
    ("CN", "국제수지 신고 은행 경유 (SAFE)", "https://www.safe.gov.cn/en/2015/0805/765.html"),
    ("CN", "CNY 송금 목적코드 (Nordea)", "https://www.nordea.com/Images/37-370815/China CNH and CNY Purpose of Payment Codes.pdf"),
    ("VN", "Circular 20/2022/TT-NHNN 개인·기관 대외송금", "https://english.luatvietnam.vn/circular-no-20-2022-tt-nhnn-guiding-activities-of-one-way-money-transfer-abroad-from-vietnam-and-240377-doc1.html"),
    ("VN", "KPMG 법률 업데이트 2023.1", "https://advisory.kpmg.us:443/content/dam/kpmg/vn/pdf/Legal-Update/2023/01/legal-technical-update-jan-23-en.pdf"),
    ("VN", "수출입 대금 은행계좌 경유 (SBV)", "https://sbv.gov.vn/en/w/sbvwebapp01sbv073513-1"),
    ("VN", "Decree 70/2014 외화수입 국내 입금", "https://apolatlegal.com/wp-content/uploads/2025/02/Decree-No.-70_2014_ND-CP.pdf"),
    ("HK", "송금인 정보 HKD 8,000 (HKMA 2006)", "https://brdr.hkma.gov.hk/eng/doc-ldg/docId/getPdf/20061122-1-EN/20061122-1-EN.pdf"),
    ("HK", "HKMA AML/CFT 가이드라인 개정안 2023", "https://www.hkma.gov.hk/media/eng/doc/key-information/guidelines-and-circular/2023/20230525e1a2.pdf"),
    ("US", "CFPB Remittance Transfer Rule 해석 (30분 취소)", "https://www.consumerfinance.gov/rules-policy/regulations/1005/interp-34"),
    ("US", "Fedwire ISO 20022 전환 완료 (2025.7)", "https://frbservices.org/news/press-releases/071525-iso20022-migration-announcement"),
    ("US", "Fedwire 2026.11 하이브리드 주소", "https://community.epcor.org/blogs/hoot-e1/2026/07/06/ask-hoot-e-preparing-for-the-next-fedwire-funds-se"),
    ("GB", "CHAPS ISO 20022 enhanced data 의무화 (BoE)", "https://www.bankofengland.co.uk/paper/2024/policy-statement/mandating-iso-20022-enhanced-data-in-chaps"),
    ("GB", "CHAPS enhanced data 2027 확대 (BoE)", "https://www.bankofengland.co.uk/paper/2025/ps/expanding-mandatory-iso-20022-enhanced-data-in-chaps-from-2027"),
    ("AU", "AUSTRAC 개편: IFTI → IVTS (2029)", "https://www.austrac.gov.au/amlctf-reform/reforms-guidance/before-you-start/summary-changes-current-reporting-entities-reform"),
    ("AU", "IFTI 정의 (AUSTRAC)", "https://www.austrac.gov.au/glossary/international-funds-transfer-instruction-ifti"),
]

QUESTIONS = [
    ("공통", "현지통화, 규제 연도(회계·보고 기준 시작월)", "정책값"),
    ("공통", "기본 Nostro(통화별), 송금 마감시각, 휴일 캘린더", "정책값"),
    ("공통", "Purpose Code / 송금 목적 코드 체계 사용 여부와 코드표", "정책값·확장"),
    ("공통", "적용 제재 목록과 필터링 기준", "정책값"),
    ("개인", "개인 해외송금 한도·집계기간·목적 제한, 증빙 없이 가능한 금액", "정책값"),
    ("개인", "개인 송금 고객 고지·영수증·취소권 등 소비자보호 요건", "확장"),
    ("기업 일반", "서비스·기타 대외지급 시 세무 확인(예: 중국 세무등록, 베트남 FCT)", "확장"),
    ("기업 수출입", "수입 송금 시 은행 확인 증빙(Invoice·세관신고)과 확인 기준금액", "정책값"),
    ("기업 수출입", "세관신고 문서 명칭·번호 형식, 은행의 사후관리 의무와 기한", "정책값"),
    ("기업 수출입", "수출대금 회수 의무·기한, 수출서류와 타발 매칭 필요 여부", "정책값"),
    ("보고", "당국 거래보고 대상·주기·시스템(예: AUSTRAC IFTI, 중국 국제수지 신고)", "확장"),
    ("보고", "LEI 필수 조건", "정책값"),
]

SCOPE = [
    ("RT11", "개인·기업 일반 송금", "RT13에서 가운데 4개 탭(Default·Invoice/LRS·Amount Limit·Dealing)을 뺀 축약판", "고객구분 개인·기업일반. 한도·목적코드·경로·구조화 주소 공통 적용"),
    ("RT13", "기업 수출입 송금(취결)", "증빙·한도·딜링 탭 포함", "고객구분 기업수출입. 증빙 매트릭스·세관신고·한도 엔진"),
    ("RT73", "Invoice / 세관신고 원장", "Invoice·Bill of Entry 등록", "세관신고 일반화, 증빙 공통 헤더"),
    ("RT74", "사후 증빙 연결", "송금 후 Invoice·세관신고 추가 연결·접수일", "사후관리 정책(사용 여부·기한)"),
    ("XP73", "수출서류·타발 매칭 (송금방식 수출)", "수출서류 접수 + 수령 타발 매칭", "범위 포함(2026-11-19 확정). 수출 세관신고·회수기한 정책"),
    ("타발송금", "수출대금 수령 (RT21 계열)", "타발 등록·IRM", "범위 포함(2026-11-19 확정). 화면 스크린샷 추가 필요"),
]
