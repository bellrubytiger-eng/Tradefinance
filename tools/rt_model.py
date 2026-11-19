# -*- coding: utf-8 -*-
"""RT11/RT13 공통 모듈 개선 설계 데이터 (현행 화면 기준 개선, 2026-11).
구분: 공통 = 전 지점 같은 로직 / 정책값 = 같은 로직, 값만 지점별 / 확장 = 해당 국가만."""

BR = ["IN", "CN", "VN", "HK", "US", "GB", "AU"]
BR_NM = {"IN": "노이다", "CN": "북경", "VN": "하노이", "HK": "홍콩", "US": "뉴욕", "GB": "런던", "AU": "시드니"}

PRINCIPLES = [
    ("기존 화면 유지", "RT13 배치(상단 · Basic Processing Info · 가운데 4개 탭 · Bank Information · TAG 70/72)를 그대로 두고 항목 단위로 개선한다. 새 화면을 만들지 않는다."),
    ("RT11 = RT13 축약", "RT11과 RT13은 같은 화면 틀을 쓴다. RT11은 가운데 4개 탭을 숨긴 모드이고, 고객구분(개인 · 기업일반 · 기업수출입)으로 화면이 정해진다."),
    ("지점 분기 없음", "소스에 지점 코드 분기를 두지 않는다. 차이는 지점정책표(값)와 확장 필드 세트(국가 항목)로만 표현한다(IT 설계서 V0.3 원칙과 동일)."),
    ("노이다 결과 재사용", "노이다에서 확정한 송금경로(Direct · Serial · COV, 운영자 선택), 복수 Nostro, ISO 20022 구조화 주소는 전 지점 공통으로 올린다."),
    ("특화값은 오픈 전 확정", "국가별 값은 08 비교맵핑의 지점확인질문지로 확정한다. 확정 전에는 노이다 값 또는 '미사용'으로 둔다."),
]

SEGMENTS = [
    # 고객구분, 판정, 화면, 가운데 탭, 한도, 증빙, 비고
    ("개인", "CIF 고객유형 = 개인", "RT11", "숨김", "개인 한도 줄(정책 사용 지점만): 집계기간 누계 · 상한 · 잔여",
     "정책: 한도 초과분 증빙(KR) / LRS(IN) / 목적별 증빙(VN) / 없음(HK·US·GB·AU)", "US: 소비자보호 고지·영수증·30분 취소"),
    ("기업 일반", "CIF 고객유형 = 법인 · 송금목적이 무역 외(서비스·배당·급여 등)", "RT11", "숨김", "없음(정책으로 기업 집계 사용 시 표시만)",
     "정책: 서비스 지급 세무 확인(CN USD 5만 초과) · 증빙 확인(VN)", "Purpose Code가 무역성이면 RT13으로 안내(V09)"),
    ("기업 수출입", "CIF 고객유형 = 법인 · 송금목적이 수입대금(무역)", "RT13", "정책에 따라 표시", "한도 탭(정책): 회계연도 시작월 · 집계기간 · 집계통화",
     "증빙 요건 매트릭스(Product Type 01~04) · Invoice · 세관신고", "XP73 · 타발은 수출 측 설계에서 연결"),
]

# 영역, 현행 항목, 개선 내용, 처리, 구분, RT11, RT13, ISO 20022 / MT, 근거·비고
FIELDS = [
    ("상단", "Ref No", "유지", "유지", "공통", "O", "O", "MsgId · EndToEndId", ""),
    ("상단", "MT / MX", "유지. 기본 MX, 상대은행 MX 미지원 시 MT", "유지", "공통", "O", "O", "pacs.008 / MT103", "CBPR+ 공존기간 종료 후 MX 고정"),
    ("상단", "Product Type (01 Outward Remittance)", "RT11: 송금종류(개인 · 기업일반 코드), RT13: 무역송금종류 01~04. 코드 의미는 지점정책의 증빙 매트릭스가 결정", "변경", "정책값", "O", "O", "-", "IT V0.3 4.3 · 확인 1"),
    ("상단", "(신규) 고객구분", "개인 · 기업일반 · 기업수출입. Cust No 입력 시 CIF 고객유형과 Purpose Code로 자동 표시(입력 불가)", "신규", "공통", "O", "O", "-", "화면 결정 키"),
    ("Basic Processing Info", "Amount / Ccy", "유지", "유지", "공통", "O", "O", "IntrBkSttlmAmt / :32A:", ""),
    ("Basic Processing Info", "Remit comm · In Lieu comm · Cable Charge (INR)", "수수료 통화 'INR' 고정 → 지점 현지통화. 수수료 코드는 지점 수수료표", "변경", "정책값", "O", "O", "-", "IT V0.3 4.4"),
    ("Basic Processing Info", "Trx Type (International)", "유지", "유지", "공통", "O", "O", "-", ""),
    ("Basic Processing Info", "Charge Type (OUR)", "유지. MX 변환: OUR=DEBT, SHA=SHAR, BEN=CRED", "유지", "공통", "O", "O", "ChrgBr / :71A:", ""),
    ("Basic Processing Info", "Cust No", "유지. 조회 시 고객구분 · 구조화 주소 · LEI 자동 채움", "변경", "공통", "O", "O", "Dbtr", ""),
    ("Basic Processing Info", "Remi Acct · Fee Acct", "유지", "유지", "공통", "O", "O", "DbtrAcct", ""),
    ("Basic Processing Info", "Agent Y/N · Agent No", "유지", "유지", "공통", "O", "O", "-", ""),
    ("Basic Processing Info", "Value Date", "유지 + 통화별 마감시각 · 지점/통화 휴일로 자동 제안(변경 가능)", "변경", "정책값", "O", "O", "IntrBkSttlmDt", "노이다 경로엔진 재사용"),
    ("Basic Processing Info", "Remit Type (T/T)", "유지", "유지", "공통", "O", "O", "-", ""),
    ("Basic Processing Info", "Purpose Code [RT07]", "코드셋을 지점정책으로: IN Purpose Code · CN 국제수지 거래코드 · VN 송금목적 분류 · 미사용(HK·US·GB·AU). RT11에서 무역성 코드 선택 시 RT13 안내", "변경", "정책값", "O", "O", "Purp · RgltryRptg / :77B:", "IT V0.3 4.2 · 확인 11"),
    ("Basic Processing Info", "Country · Nationality", "유지", "유지", "공통", "O", "O", "RgltryRptg", ""),
    ("Basic Processing Info", "LEI NO + 만기일 [CH11]", "공통 선택 항목. 필수 조건은 정책(IN INR 50 crore 이상, GB CHAPS 금융기관 지급 등). 만기 경과 시 경고", "변경", "정책값", "O", "O", "Dbtr/Id/OrgId/LEI", "IT V0.3 3.3"),
    ("Basic Processing Info", "IEC CODE", "확장 영역(IN)으로 이동", "이동", "확장", "-", "O", "-", "IN-TRD"),
    ("Basic Processing Info", "TCS Y/N", "확장 영역(IN)으로 이동", "이동", "확장", "O", "O", "-", "IN-TRD"),
    ("확장 영역 (신규 줄)", "(신규) 확장 영역", "Basic Processing Info 아래 1줄. 지점정책의 확장 세트가 있을 때만 표시, 항목은 세트 정의로 동적 표시", "신규", "확장", "O", "O", "-", "IT V0.3 4.4 '확장 영역'"),
    ("확장 영역 (신규 줄)", "IN: IEC · TCS · LRS Y/N", "현행 항목 그대로 이동(LRS는 Invoice/LRS 탭에서 이동). RT11 개인은 LRS · TCS만", "이동", "확장", "O", "O", "-", "노이다 현행"),
    ("확장 영역 (신규 줄)", "CN: 국제수지 거래코드 · 무역기업 명부 확인 · 세무등록번호", "명부 확인은 조회 결과 표시(Y/N). 세무등록번호는 서비스 등 대외지급 건당 USD 5만 초과 시 필수", "신규", "확장", "O", "O", "RgltryRptg", "08 비교맵핑 · 2차 자료, 지점 확인"),
    ("확장 영역 (신규 줄)", "VN: 송금목적 분류 · 증빙 확인 · 보관 확인", "개인은 허용 목적(유학 · 의료 · 친족부양 등)만 선택 가능. 증빙 확인자 · 보관 위치 기록", "신규", "확장", "O", "O", "Purp", "Circular 20/2022 · 2차 자료"),
    ("확장 영역 (신규 줄)", "US: 소비자 송금 고지", "개인 송금만. 사전 고지서 출력 여부 · 영수증 출력 · 취소 가능 시각(결제 후 30분) 표시", "신규", "확장", "O", "-", "-", "Reg E Subpart B"),
    ("Basic Processing Info", "(신규) 개인 한도 줄", "RT11 개인 + 정책 사용 지점만: 집계기간 · 누계 · 금회 · 상한 · 잔여. 초과 시 정책에 따라 증빙 요구 또는 불가", "신규", "정책값", "O", "-", "-", "KR USD 10만 · IN LRS USD 25만 · CN USD 5만"),
    ("Basic Processing Info", "(신규) 송금인 정보 완비 표시", "기준금액 이상이면 송금인 이름 · 계좌 · 주소(또는 ID) 완비 확인. 기준금액은 정책(HK HKD 8,000, US USD 3,000 등)", "신규", "공통", "O", "O", "Dbtr · PstlAdr", "Travel rule"),
    ("Default 탭", "Applicant / Beneficiary: Acct No · Name", "유지. RT11은 Default 탭 내용을 탭 없이 본문에 고정 표시", "유지", "공통", "O", "O", "Dbtr · Cdtr", "RT11 현행 위치 확인 필요"),
    ("Default 탭", "Remitt. Ctry", "국가(ISO 3166) — Ctry", "변경", "공통", "O", "O", "PstlAdr/Ctry", "필수"),
    ("Default 탭", "State Name", "주 · 성 — CtrySubDvsn", "변경", "공통", "O", "O", "PstlAdr/CtrySubDvsn", ""),
    ("Default 탭", "City Name", "도시 — TwnNm", "변경", "공통", "O", "O", "PstlAdr/TwnNm", "필수"),
    ("Default 탭", "Zip Code", "우편번호 — PstCd", "변경", "공통", "O", "O", "PstlAdr/PstCd", ""),
    ("Default 탭", "Address 3줄", "구조화: 도로명 StrtNm · 건물번호 BldgNb · 건물명 BldgNm · 층 Flr + 하이브리드 주소줄 AdrLine 최대 2줄 × 70자", "변경", "공통", "O", "O", "StrtNm · BldgNb · BldgNm · Flr · AdrLine", "CBPR+ 2026.11 하이브리드 · 노이다 적용분"),
    ("Default 탭", "Remark", "유지", "유지", "공통", "O", "O", "-", ""),
    ("증빙 탭 (구 Invoice/LRS)", "LRS Y/N", "IN 확장 영역으로 이동", "이동", "확장", "-", "O", "-", "IT V0.3 4.4"),
    ("증빙 탭 (구 Invoice/LRS)", "Invoice 목록 (No · Date · Input Amt · Balance · Used)", "유지 + 관리번호 · 상태(등록/해제) 열 추가(증빙 공통 헤더)", "변경", "공통", "-", "O", "RmtInf/Strd (선택)", "IT V0.3 4.5 · V01 · V02"),
    ("증빙 탭 (구 Invoice/LRS)", "Total Remittance (Invoice)", "유지. 통화 라벨은 송금통화", "유지", "공통", "-", "O", "-", ""),
    ("증빙 탭 (구 Invoice/LRS)", "(이동) Bill of Entry No · Management Ref No · Registration Date", "Amount Limit 탭에서 증빙 탭 '세관신고' 구역으로 이동. 명칭은 정책(Bill of Entry · 报关单 · Tờ khai · 수입신고필증), 미사용 지점은 숨김", "이동", "정책값", "-", "O", "-", "한도와 무관한 항목이 한도 탭에 있던 것을 정리(제안)"),
    ("증빙 탭 (구 Invoice/LRS)", "(신규) 증빙 요건 표시", "선택한 Product Type의 매트릭스(Invoice · 세관신고 · Dealing 필수/옵션/불가)를 탭 상단에 표시", "신규", "정책값", "-", "O", "-", "V04"),
    ("한도 탭 (구 Amount Limit)", "Total Remittance (Fiscal Year) YEAR · Start · End · Total USD", "회계연도 시작월 · 집계기간(년) · 집계통화를 정책값으로. 상한 적용 지점은 상한 · 잔여 열 추가. 미사용 지점은 탭 숨김", "변경", "정책값", "-", "O", "-", "IT V0.3 4.7 · 확인 2"),
    ("Dealing Information 탭", "Deal Type · Exchange Rate · Date · Deal Ref No", "유지", "유지", "공통", "-", "O", "-", ""),
    ("Dealing Information 탭", "Current / Total Amount (INR) · (USD)", "'INR' → 현지통화, 'USD' → 기준통화(정책)", "변경", "정책값", "-", "O", "-", "IT V0.3 4.4"),
    ("Dealing Information 탭", "Commodity ID · CounterParty ID", "유지. 탭 사용 여부는 정책(7개 지점 확인)", "유지", "정책값", "-", "O", "-", "IT V0.3 4.2"),
    ("Bank Information", "(신규) 송금 경로", "Direct · Serial · COV 판정 결과 표시. Serial과 COV가 모두 가능하면 운영자가 선택하고 선택사유 코드 입력", "신규", "공통", "O", "O", "pacs.008 / pacs.009 COV", "노이다 3.16 확정안"),
    ("Bank Information", "Msg Rcv Bank", "경로에 따라 자동: Direct · COV = 수취은행, Serial = Nostro(또는 중계은행). 수정 시 사유", "변경", "공통", "O", "O", "InstdAgt / 수신 BIC", ""),
    ("Bank Information", "Reimbursement Bank (Nostro) 4001", "통화별 기본 Nostro는 정책값, 같은 통화 복수 Nostro 선택 가능. Nostro BIC 표시", "변경", "정책값", "O", "O", "InstgRmbrsmntAgt · :53:", "노이다 pickNostro"),
    ("Bank Information", "57A / 57D · Acct #", "유지(수취은행). BIC 조회 시 은행명 · 국가 · gpi 여부 표시", "유지", "공통", "O", "O", "CdtrAgt · :57A:", ""),
    ("Bank Information", "Receiver's Corres · 54/D", "Nostro와 경로에서 자동 채움(직접 입력 → 확인)", "변경", "공통", "O", "O", "InstdRmbrsmntAgt · :54:", ""),
    ("Bank Information", "Intermediary Bank · 56/D", "수취은행 SSI에서 자동 제안, 운영자 수정 가능(출처 표시: SSI · 수동)", "변경", "공통", "O", "O", "IntrmyAgt1 · :56:", ""),
    ("TAG 70", "Additional Details", "유지. MX는 RmtInf/Ustrd(140자), 증빙 연결 시 Strd 선택", "유지", "공통", "O", "O", "RmtInf / :70:", ""),
    ("TAG 72", "Additional Instructions", "유지. MX는 InstrForNxtAgt · InstrForCdtrAgt로 나눠 입력", "변경", "공통", "O", "O", "InstrForNxtAgt / :72:", ""),
    ("하단 (신규)", "(신규) 필터링 · 검증 결과", "제재 · 무역거래 필터링 결과(REF-NO) · V01~V18 결과를 한 줄로. 오류가 있으면 취결 버튼 비활성", "신규", "공통", "O", "O", "-", "IT V0.3 확인 10"),
]

# 정책 항목, 설명, {지점: 값}
POLICY = [
    ("현지통화", "수수료 · Dealing 금액 라벨", {"IN": "INR", "CN": "CNY", "VN": "VND", "HK": "HKD", "US": "USD", "GB": "GBP", "AU": "AUD"}),
    ("규제연도 시작월", "한도 집계 · Invoice 기본 조회기간", {"IN": "4", "CN": "1", "VN": "1", "HK": "4", "US": "1", "GB": "4", "AU": "7"}),
    ("Purpose Code 세트", "Purpose Code 코드그룹", {"IN": "IN-PURPOSE", "CN": "CN-BOP", "VN": "VN-PURPOSE (확인)", "HK": "미사용", "US": "미사용", "GB": "미사용", "AU": "미사용"}),
    ("확장 세트", "확장 영역 항목 세트", {"IN": "IN-TRD", "CN": "CN-TRD (확인)", "VN": "VN-TRD (확인)", "HK": "없음", "US": "US-CNS", "GB": "없음", "AU": "없음"}),
    ("개인 한도", "사용 · 상한 · 집계기간", {"IN": "Y · USD 250,000 · 회계연도 (LRS)", "CN": "Y · USD 50,000 · 역년 (확인)", "VN": "목적별 (확인)", "HK": "N", "US": "N", "GB": "N", "AU": "N"}),
    ("개인 한도 초과 시", "증빙 요구 / 불가", {"IN": "불가", "CN": "증빙 요구 (확인)", "VN": "목적 · 증빙 (확인)", "HK": "-", "US": "-", "GB": "-", "AU": "-"}),
    ("기업 한도 집계", "한도 탭: 사용 · 기간 · 통화", {"IN": "Y · 3년 · USD", "CN": "확인", "VN": "확인", "HK": "N", "US": "N", "GB": "N", "AU": "N"}),
    ("증빙 매트릭스", "Product Type별 필수/옵션/불가", {"IN": "IN-TRD-01~04", "CN": "CN-TRD (확인)", "VN": "VN-TRD (확인)", "HK": "COM-OPT", "US": "COM-OPT", "GB": "COM-OPT", "AU": "COM-OPT"}),
    ("세관신고 문서", "증빙 탭 세관신고 구역 명칭 · 사용", {"IN": "Bill of Entry", "CN": "报关单 (확인)", "VN": "Tờ khai hải quan (확인)", "HK": "미사용", "US": "미사용", "GB": "미사용", "AU": "미사용"}),
    ("세관신고 사후관리", "Due · Receive Date 관리", {"IN": "Y", "CN": "확인", "VN": "확인", "HK": "N", "US": "N", "GB": "N", "AU": "N"}),
    ("Dealing 탭", "사용 여부", {"IN": "Y", "CN": "확인", "VN": "확인", "HK": "확인", "US": "확인", "GB": "확인", "AU": "확인"}),
    ("송금인 정보 기준금액", "V11 완비 확인 기준", {"IN": "확인", "CN": "확인", "VN": "확인", "HK": "HKD 8,000", "US": "USD 3,000", "GB": "확인", "AU": "확인"}),
    ("소비자 송금 보호", "개인 고지 · 영수증 · 취소", {"IN": "N", "CN": "N", "VN": "N", "HK": "N", "US": "Y (30분 취소)", "GB": "확인 (PSRs)", "AU": "N"}),
    ("LEI 필수 조건", "V18", {"IN": "INR 50 crore 이상", "CN": "-", "VN": "-", "HK": "-", "US": "-", "GB": "CHAPS 금융기관 지급 (확인)", "AU": "-"}),
    ("제재 목록 세트", "필터링 목록", {"IN": "UN · IN", "CN": "UN · CN", "VN": "UN · VN", "HK": "UN · HK", "US": "OFAC SDN", "GB": "UK Sanctions List", "AU": "DFAT"}),
    ("당국 보고", "송금 원장 연계 보고", {"IN": "CIMS · IDPMS", "CN": "국제수지 신고", "VN": "SBV 보고", "HK": "-", "US": "BSA · OFAC", "GB": "-", "AU": "IFTI 10영업일 (화면 항목 없음, 자동)"}),
    ("기본 USD Nostro", "통화별 기본값 (복수 가능)", {k: "지점 확인" for k in BR} | {"IN": "4001 (현행)"}),
]

# Product Type 매트릭스: 매트릭스ID, Product Type, 의미(추정), Invoice, 세관신고, Dealing, 확장 규칙
MATRIX = [
    ("IN-TRD", "01", "일반 (현행 Outward Remittance)", "옵션", "옵션", "옵션", "LRS 옵션"),
    ("IN-TRD", "02", "수입 선급 (BoE 전, 추정)", "옵션", "불가", "옵션", "LRS 필수 입력, LRS=Y면 Invoice 불가"),
    ("IN-TRD", "03", "수입 후급 (선등록, 추정)", "필수", "필수", "옵션", "LRS 필수 입력, RT73 선등록"),
    ("IN-TRD", "04", "증빙 없음 (추정)", "불가", "불가", "옵션", "LRS 옵션"),
    ("COM-OPT", "01~04", "비통제 지점 공통", "옵션", "숨김", "정책", "증빙은 무역기반 자금세탁 점검용 선택 입력"),
    ("CN-TRD", "01~04", "중국 (확인 후 정의)", "필수 (확인)", "수입 후급 필수 (확인)", "정책", "무역기업 명부 확인 Y 필수"),
    ("VN-TRD", "01~04", "베트남 (확인 후 정의)", "필수 (확인)", "확인", "정책", "증빙 확인자 · 보관 기록 필수"),
]

RULES = [
    ("V01", "송금금액 ≤ 선택한 Invoice 합계", "공통", "RT13", "IT V0.3"),
    ("V02", "Invoice 잔액은 다른 송금에 사용 가능, Used 관리", "공통", "RT13", "IT V0.3"),
    ("V03", "증빙 선행등록 · Cust No 일치 · 세관신고의 Invoice 참조", "공통", "RT13", "IT V0.3"),
    ("V04", "Product Type별 Invoice · 세관신고 필수/옵션/불가", "정책값", "RT13", "IT V0.3 · 증빙 매트릭스"),
    ("V05", "LRS=Y이면 Invoice 입력 불가", "확장 IN", "RT13", "IT V0.3"),
    ("V06", "집계기간 누계(집계통화 환산)", "정책값", "RT13", "IT V0.3 · 상한 적용 확인"),
    ("V07", "사용된 증빙 금액 · 통화 · 고객 변경 및 취소 불가", "공통", "RT73 · RT74", "IT V0.3"),
    ("V08", "RT74 추가 연결 · 해제에 V01~V04 · V07 동일 적용", "공통", "RT74", "IT V0.3"),
    ("V09", "RT11에서 무역성 Purpose Code 또는 기업수출입 고객이면 RT13으로 안내(취결 불가)", "공통", "RT11", "신규"),
    ("V10", "개인 한도: 집계기간 누계 + 금회 ≤ 상한, 초과 시 정책(증빙 요구 / 불가)", "정책값", "RT11", "신규 · 08 비교맵핑"),
    ("V11", "기준금액 이상 송금인 정보(이름 · 계좌 · 주소 또는 ID) 완비", "공통 + 정책값", "RT11 · RT13", "신규"),
    ("V12", "구조화 주소: 국가 · 도시 필수, 주소줄은 최대 2줄 × 70자", "공통", "RT11 · RT13", "노이다 R-ADR"),
    ("V13", "Serial · COV 모두 가능하면 운영자 선택 + 선택사유 필수, COV는 Nostro · 수취은행 SSI 확인", "공통", "RT11 · RT13", "노이다 R-RTG"),
    ("V14", "CN: 서비스 등 대외지급 USD 5만 초과 시 세무등록번호 필수", "확장 CN", "RT11 · RT13", "신규 · 확인"),
    ("V15", "CN: 무역 송금 시 무역기업 명부 확인 = Y", "확장 CN", "RT13", "신규 · 확인"),
    ("V16", "VN: 개인은 허용 목적만, 증빙 확인자 · 보관 기록 필수", "확장 VN", "RT11 · RT13", "신규 · 확인"),
    ("V17", "US: 개인 송금은 사전 고지 출력 후 취결, 결제 후 30분 내 취소 허용", "확장 US", "RT11", "신규"),
    ("V18", "LEI 필수 조건(정책) 충족 · 만기 미경과", "정책값", "RT11 · RT13", "신규"),
]

OPEN = [
    ("RT11 현행 당사자 입력 위치", "RT11에서 Applicant · Beneficiary를 어디서 입력하는지(탭 없이 본문인지) 스크린샷 필요", "RT11 화면 배치"),
    ("Product Type 01~04 정의", "02 · 03 · 04의 업무 의미(추정: 선급 · 후급 · 증빙 없음)와 전 지점 공통 사용 가능 여부", "증빙 매트릭스 (IT 확인 1)"),
    ("세관신고 구역 이동", "BoE 항목을 한도 탭에서 증빙 탭으로 옮기는 제안 수용 여부", "증빙 · 한도 탭"),
    ("개인 한도 초과 처리", "지점별로 증빙 요구인지 취결 불가인지", "V10"),
    ("RT11 개인 · 기업일반 코드", "RT11의 송금종류 코드 체계(현행 코드 확인)", "상단 Product Type"),
    ("CN · VN 확장 항목", "중국 · 베트남 지점 확인 후 확정(08 지점확인질문지)", "확장 세트 CN-TRD · VN-TRD"),
    ("US 소비자 송금 대상", "뉴욕지점의 개인 송금 취급 여부와 고지서 양식", "확장 세트 US-CNS"),
]
