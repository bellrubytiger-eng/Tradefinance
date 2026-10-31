// 설계서 Word 생성: node tools/build_docx.js <model.json> <out.docx>
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType,
  WidthType, ShadingType, BorderStyle, LevelFormat, Footer, Header, PageNumber, TableOfContents, PageBreak,
} = require("docx");

const M = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = process.argv[3];
const FONT = "맑은 고딕";
const W = 9638; // A4, 좌우여백 2cm
const NAVY = "1F4E79";

const run = (t, o = {}) => new TextRun({ text: String(t), font: FONT, size: o.size || 20, bold: o.bold, color: o.color, italics: o.italics });
const P = (t, o = {}) => new Paragraph({ spacing: { after: 100, line: 300 }, alignment: o.align, children: Array.isArray(t) ? t : [run(t, o)] });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { after: 200 }, children: [run(t, { size: 32, bold: true, color: NAVY })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, children: [run(t, { size: 26, bold: true, color: NAVY })] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: 80 }, children: [run(t, { size: 22, bold: true })] });
const B = (t, lvl = 0) => new Paragraph({ numbering: { reference: "bul", level: lvl }, spacing: { after: 60, line: 280 }, children: Array.isArray(t) ? t : [run(t)] });
const N = (t) => new Paragraph({ numbering: { reference: "num", level: 0 }, spacing: { after: 60, line: 280 }, children: [run(t)] });
const KV = (k, v) => B([run(k + " : ", { bold: true }), run(v)]);
const NOTE = (t) => new Paragraph({
  spacing: { before: 80, after: 120 }, shading: { type: ShadingType.CLEAR, fill: "FFF2CC" },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: "C9A227", space: 6 } },
  children: [run(t, { size: 19 })],
});

const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };
function TBL(headers, rows, ratios, opt = {}) {
  const sum = ratios.reduce((a, b) => a + b, 0);
  const widths = ratios.map((r) => Math.floor((W * r) / sum));
  widths[widths.length - 1] += W - widths.reduce((a, b) => a + b, 0);
  const cell = (t, i, hdr, fill) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: hdr ? NAVY : (fill || "FFFFFF") },
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: String(t ?? "").split("\n").map((line) => new Paragraph({ children: [run(line, { size: opt.size || 17, bold: hdr, color: hdr ? "FFFFFF" : undefined })] })),
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((v, i) => cell(v, i, false, opt.verifyCol !== undefined && r[opt.verifyCol] === "Y" ? "FFF2CC" : undefined)) })),
    ],
  });
}
const gap = () => new Paragraph({ spacing: { after: 80 }, children: [] });

const C = []; // body
// ---------------- 표지 ----------------
const cover = [
  new Paragraph({ spacing: { before: 2600 }, children: [] }),
  P([run("GBS India 외환·무역송금 코어뱅킹", { size: 44, bold: true, color: NAVY })], {}),
  P([run("요건정의 및 시스템 설계서", { size: 36, bold: true })]),
  new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 4 } }, spacing: { after: 300 }, children: [] }),
  P([run("Outward / Inward Remittance · LRS · FEMA Trade(IDPMS/EDPMS) · FDI · INR Vostro(MT950) · RBI CIMS", { size: 22 })]),
  new Paragraph({ spacing: { before: 1600 }, children: [] }),
  TBL(["항목", "내용"], [
    ["대상", M.META.branch], ["버전 / 작성일", `${M.META.version} / ${M.META.date}`],
    ["동반 산출물", "02_데이터사전_업무규칙.xlsx (컬럼·규칙·기한·코드), db/ddl_oracle.sql, screens/index.html (화면 시안)"],
    ["기반 자료", M.META.base_docs.join("\n")],
  ], [1, 4], { size: 18 }),
  new Paragraph({ children: [new PageBreak()] }),
  P([run("목차", { size: 28, bold: true, color: NAVY })]),
  new TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-2" }),
  P([run("※ Word에서 목차가 비어 보이면 목차 영역 우클릭 → '필드 업데이트'를 실행하십시오.", { size: 16, italics: true, color: "808080" })]),
];

// ---------------- 1. 개요 ----------------
C.push(H1("1. 개요"));
C.push(H2("1.1 배경"));
C.push(P("노이다지점은 2024.09 '기업 수출입 업무 고도화' 요청을 통해 RT11 단일 화면에서 처리하던 무역/무역외 송금을 분리하고, 상업송장·세관신고서(Bill of Entry) 기반 분할송금 잔액관리와 D/P·D/A 서류관리(Lodgement)를 요청하였다. 이후 2024.11~2025.02 테스트 과정에서 RT13·RT73·XP73 화면에 대한 세부 개선요청이 누적되었다."));
C.push(P("본 설계는 이 요청들을 수용하면서 범위를 다음과 같이 확장하여, 개인·기업 고객 전체의 외환송금을 FEMA·LRS·RBI 보고체계(IDPMS/EDPMS/CIMS)에 맞게 재구성하는 것을 목표로 한다."));
C.push(H2("1.2 범위"));
C.push(TBL(["업무", "개인 고객", "기업 고객(개인사업자 포함)", "주요 화면"], [
  ["당발송금 (Outward)", "LRS 송금(유학·의료·여행·증여·생활비·투자)", "무역외 일반송금, 수입대금(선수금/사후/Direct), 서비스 대가, ODI", "RT11, RT12, RT13"],
  ["타발송금 (Inward)", "개인 송금 수취, NRI 송금", "수출선수금·수출대금 회수(IRM), FDI 수취", "RT21, RT25"],
  ["수출입 서류/신고", "-(개인사업자는 기업과 동일)", "Invoice, Packing List, BoE, Shipping Bill, Lodgement", "RT73, XP73"],
  ["FDI 외환송금", "-", "FDI 수취, 주식배정 추적, 반환송금, INR 금액 확정송금", "FD11, FD12"],
  ["INR Vostro", "-", "해외은행 INR 계좌(일반/SRVA/RDA), MT950 자동발송", "VS11, VS31"],
  ["규제보고/기한관리", "LRS 일일보고(CIMS)", "IDPMS, EDPMS, R-Return, FC-GPR 기한", "MN01, IF01"],
], [1.3, 2, 3, 1.2]));
C.push(H2("1.3 설계 원칙"));
[
  "단일 송금원장(TB_FX_REMIT_MST): 모든 송금 1건 = 1행. 보고서 시스템은 원장 + 상세 테이블만으로 모든 RBI 보고를 재현할 수 있어야 한다.",
  "스냅샷 원칙: PAN·IEC·LEI·고객유형은 송금 시점 값을 원장에 복사 저장한다(고객정보 변경 후에도 과거 보고 재현).",
  "잔액은 배분원장(TB_TRD_PAY_ALLOC)의 합계로만 갱신: Invoice·BoE·Shipping Bill·IRM의 Total / Consumed / Balance 를 하나의 서비스에서 일관 처리.",
  "규정 수치는 하드코딩 금지: TB_FX_PARAM 에 유효기간(EFF_FROM/TO)과 함께 관리하고, 기한 산출 시 사용값을 스냅샷 저장한다(2025.11 RBI 한시완화처럼 선적일 기준 이력 적용).",
  "모든 규제기한은 TB_FX_OBLIGATION 단일 테이블: 기한유형·기산일·기한·상태만 보면 미이행 현황을 산출할 수 있도록 한다.",
  "ISO 20022 우선: 크로스보더 당발은 pacs.008/pacs.009 생성. 기존 화면의 TAG70/TAG72 항목은 MX 필드로 매핑한다.",
].forEach((t) => C.push(B(t)));

// ---------------- 2. 규제 ----------------
C.push(H1("2. 규제 요건 요약"));
C.push(NOTE("본 장의 수치는 2026.10 기준 설계값이며, 노란색(VERIFY) 항목은 운영 반영 전 컴플라이언스 부서의 최신 RBI Master Direction / A.P.(DIR) Circular 확인이 필요하다. 모든 수치는 TB_FX_PARAM으로 관리되어 개정 시 코드 수정 없이 반영된다."));
C.push(H2("2.1 근거 법규"));
C.push(TBL(["구분", "근거", "시스템 반영"], [
  ["외환 일반", "Foreign Exchange Management Act 1999, Current Account Transactions Rules 2000(Schedule I~III)", "Purpose별 허용/금지, 금액한도 검증"],
  ["개인 송금", "RBI Master Direction - Liberalised Remittance Scheme", "RT12, USD 250,000/FY, Form A2, PAN"],
  ["세금", "Income-tax Act s.206C(1G) TCS, s.195 / Rule 37BB (Form 15CA·15CB)", "TCS 자동계산, 15CA/CB 검증 (2026.4 Income-tax Act 2025 조항 대응 확인)"],
  ["수입", "RBI Master Direction - Import of Goods and Services, IDPMS 운영지침", "선수금 기한, BoE 매칭, ORM 보고"],
  ["수출", "RBI Master Direction - Export of Goods and Services, EDPMS 운영지침", "회수기한, IRM, SB 매핑, 연장/상각"],
  ["FDI", "FEMA (Non-Debt Instruments) Rules 2019, Mode of Payment & Reporting Regulations 2019, Press Note 3(2020)", "60일 배정, 75일 반환, FC-GPR 30일, 인접국 정부승인"],
  ["ODI", "FEMA (Overseas Investment) Rules/Regulations 2022", "RT13 Product 05"],
  ["LEI", "RBI circular - LEI for large value cross-border transactions", "비개인 INR 50 crore 이상 LEI 필수"],
  ["Vostro", "RBI - International trade settlement in INR (SRVA, 2022.7)", "SRVA 계좌유형, 사용제한"],
  ["KYC/AML", "RBI Master Direction - KYC 2016, PMLA 2002", "고객 KYC 검증, 제재국 차단, TBML 데이터"],
  ["메시지", "SWIFT CBPR+ (MT/MX 공존기간 2025.11 종료), MT950 Standard", "pacs.008 생성, 구조화 주소, MT950 생성"],
], [1, 3, 2.4]));
C.push(H2("2.2 핵심 규정 파라미터"));
C.push(TBL(["파라미터", "명칭", "값", "근거", "VERIFY"], M.PARAMS.map((p) => [p[0], p[1], `${p[2]} ${p[3]}`, p[5], p[6]]), [2.2, 2.6, 1.1, 2.6, 0.6], { verifyCol: 4, size: 15 }));

// ---------------- 3. 프로세스 ----------------
C.push(H1("3. TO-BE 업무 프로세스"));
const flow = (title, steps) => { C.push(H2(title)); C.push(TBL(["단계", "처리내용", "시스템/데이터"], steps.map((s, i) => [String(i + 1), s[0], s[1]]), [0.5, 4, 2.4])); };
flow("3.1 개인 LRS 당발송금 (RT12)", [
  ["고객 조회: 거주 개인·PAN 검증·KYC 유효 확인", "CH11, R-LRS-01"],
  ["목적 선택(LRS 분류·Purpose Code), 금지목적 차단", "R-LRS-05"],
  ["Form A2 + 타행 기송금 자기신고 입력", "TB_LRS_TXN.DECL_OTHER_BANK_USD"],
  ["FY 누적 조회(행잠금) → USD 250,000 한도 검증", "TB_LRS_FY_SUM, R-LRS-02/07"],
  ["TCS 계산(INR 10 lakh 초과분 × 목적별 세율), 수수료·GST 계산", "R-LRS-04, TB_FX_REMIT_FEE"],
  ["딜링 맵핑 / 환율 확정 → 책임자 승인 → pacs.008 송신", "TB_FX_DEAL_LINK, SWIFT"],
  ["누적 갱신, 익영업일 CIMS LRS 일일보고, 분기 TCS(27EQ)", "TB_CIMS_RETURN"],
]);
flow("3.2 수입 선수금 송금 (RT13 Product 01)", [
  ["RT73에서 Proforma Invoice 등록(유효기간·Incoterms·HS·품목)", "TB_TRD_INVOICE(_ITEM)"],
  ["RT13: IEC(CH11)·LEI 판정, Invoice 돋보기 조회 → 배분액 입력", "R-IMP-01/02/03"],
  ["USD 200,000 초과 시 BG/SBLC 정보 입력", "R-IMP-06"],
  ["승인 → Invoice Consumed/Balance 갱신, ORM(A) 생성", "TB_TRD_PAY_ALLOC, TB_IF_IDPMS_ORM"],
  ["선수금 기한 생성: 송금일+6M(자본재 36M)+15일", "TB_TRD_ADV_IMP, TB_FX_OBLIGATION"],
  ["IDPMS BoE 수신 → 자동/수동 매칭 → 증빙액 갱신 → 완결 시 FULFILLED, Settlement 보고", "IF-IDP-02/03"],
  ["기한 임박 D-30/15/7/0 독촉, 경과 시 OVERDUE 보고", "MN01"],
]);
flow("3.3 수입대금 사후송금 / Direct Import (Product 02·03)", [
  ["BoE 존재 확인(IDPMS 수신분 우선, 미수신 시 RT73 수기등록)", "TB_TRD_BOE"],
  ["RT13 Amount Limit 탭: Invoice Total/Consumed/Balance, BoE Ref/Registration Date/Value 표시", "R-IMP-08/09/10"],
  ["선적일+6개월 결제기한 검증(초과 시 연장근거)", "R-IMP-05"],
  ["수수료: Product 03은 INR 750 정액 + GST 18%, EEFC 출금 시 In lieu of Exchange", "R-IMP-12/13"],
  ["승인 → BoE 미결제액 차감 → IDPMS Settlement 보고, 완결 시 Closure", "TB_IF_IDPMS_SETL"],
]);
flow("3.4 수출: 서류(Lodgement) → 회수(EDPMS)", [
  ["EDPMS로 Shipping Bill 수신 → 회수기한(LEO+9M)·서류제출기한(+21일) 생성", "IF-EDP-01"],
  ["XP73: Lodgement Category(1 FBFP / 2A Collection / 2B Realization / 3 Lodge & Realize) 선택 → 카테고리별 필드", "R-EXP-01"],
  ["RT73 Invoice 조회 연동, 운송·Packing·Amount(FOB/Freight/Insurance/IGST) 입력", "TB_TRD_EXP_LODGE, TB_TRD_LODGE_SB"],
  ["대금 타발 입금(RT21) → 수출 Purpose이면 IRM 자동생성 → EDPMS 송신", "R-EXP-03"],
  ["RT25에서 IRM↔SB 맵핑(다건) → 회수 보고, 잔여분 연장/상각/상계 처리", "R-EXP-04, IF-EDP-03"],
  ["선수금 IRM은 선적이행기한 관리(미이행 시 반환송금)", "R-EXP-06"],
]);
flow("3.5 FDI 수취·배정·반환 (FD11/FD12)", [
  ["타발 입금(RT21, Purpose FDI) → FD11 자동 이관", "TB_FDI_INFLOW"],
  ["해외 송금은행 KYC 수취, 인접국(Press Note 3)·정부승인 검증", "R-FDI-01"],
  ["INR 확정송금: 투자약정 INR 금액 고정 → 필요 FCY 역산·환율 확정(딜링)", "R-FDI-05"],
  ["배정기한(+60)·반환기한(+75) 생성, 배정 등록 시 FC-GPR(+30) 기한 생성", "R-FDI-02"],
  ["미배정분 반환송금(FD12 → Product 06 당발): 원통화·원송금인 앞 반환, 기한경과 시 RBI 승인번호 필수", "R-FDI-03/04"],
]);
flow("3.6 INR Vostro 및 MT950 자동발송 (VS11/VS21/VS31)", [
  ["Vostro 계좌 개설(일반/SRVA/RDA), 명세서 형식·주기·무거래일 발송 설정", "TB_VOSTRO_ACCT"],
  ["입출금: 타발 INR 지급지시·무역결제·SRVA 투자 → 잔액 검증(대월 불가)", "R-VS-01/02/05"],
  ["EOD 마감 → 계좌별 :61: 라인 집계 → 100라인 단위 페이지 분할(60F/60M·62M/62F)", "R-VS-03"],
  ["MT950 생성·검증(기초±거래=기말) → SWIFT 송신 → ACK/NAK 수신, NAK 재발송", "IF-SWF-03/04"],
  ["SRVA 잔액·투자 현황 CIMS 보고", "IF-CIM-03"],
]);
C.push(H3("MT950 생성 예시"));
[":20:VS2610310001", ":25:9001000000123", ":28C:00215/1", ":60F:C261030INR125000000,00",
  ":61:2610311031D5000000,00NTRFINV20261031A//NHNOIDA0001", ":61:2610311031C2500000,00NTRFNONREF//NHNOIDA0002",
  ":62F:C261031INR122500000,00", ":64:C261031INR122500000,00"].forEach((l) => C.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: l, font: "Consolas", size: 18 })] })));
C.push(P("※ MT950은 MT940과 달리 :86: 필드가 없으므로, 거래 부가정보는 :61: Supplementary Details(34x)에 기재한다. :61: = 결제일(YYMMDD)+기장일(MMDD)+D/C+금액+N+거래유형(TRF 등)+계좌주Ref//당행Ref.", { size: 17, italics: true }));

// ---------------- 4. 화면 ----------------
C.push(H1("4. 화면 설계"));
C.push(P("화면 시안은 screens/index.html 에서 확인할 수 있으며(현행 GBS 레이아웃 유지), 주요 검증 로직(LRS 한도·TCS, LEI 판정, Invoice 잔액, BoE 등록일, FDI 기한, MT950 생성)이 동작하도록 구현되어 있다."));
C.push(TBL(["화면ID", "화면명", "대상", "구분", "주요기능"], M.SCREENS.map((s) => [s[0], s[1], s[2], s[3], s[4]]), [0.8, 2, 1.1, 0.6, 4]));
C.push(H2("4.1 공통 컴포넌트"));
[
  "고객 헤더: Customer No / Name / PAN / IEC / LEI + 만료일 (CH11 조회값, 수정불가)",
  "계좌 드롭다운: 고객 보유 CA/SA/EEFC 목록 + 잔액 표시 (XP73 요청사항)",
  "Dealing Information 탭: Deal Type / Date / Ref No(GBS MM, 수기 허용) / FX Rate / Current INR·USD / Total INR·USD",
  "Purpose Code: 마스터 선택 + Others(5자리 코드, Description 100자)",
  "Amount 블록: Total / Consumed / Balance 3필드 표준화",
].forEach((t) => C.push(B(t)));
C.push(H2("4.2 화면별 핵심 변경"));
C.push(H3("RT13 Remittance (Trade Finance)"));
["탭 구성: Default / Product Type / Amount Limit / Dealing Information", "Amount Limit: Total Remittance(Fiscal Year) + Invoice별 Total/Consumed/Balance + Bill of Entry(Ref No, Registration Date, Value)",
  "BoE Product Code ≠ Product Type 오류(현행 rto1300 Line 992) 제거, 정보표시로 변경", "수수료 그리드: Commission(42020601) + GST, TCS Y/N, Commission 수기입력(CCY/Amount) 허용", "Reim Bank(Nostro) 통화별 자동세팅, 명칭 'Receiver's Correspondent' / 'Reimbursement Bank(Nostro)'"].forEach((t) => C.push(B(t)));
C.push(H3("RT73 Invoice / Bill of Entry Mgmt"));
["Invoice Category, Amount/IGST Rate/IGST Amount/Total Amount, Total/Consumed/Balance", "Shipping Date 미래일자 → Packing List 생략 허용(현행 rto7300 Line 989 오류 원인 해소)", "Vessel Name(Flight No)/Voyage No 20자, Port of Shipment/Country of Origin/Consignment/Port of Discharge/Port Code/Place of Destination", "BoE 탭: Home Consumption/Warehousing/Ex-bond, Invoice 값 자동복사, IDPMS 수신 BoE 조회·매칭"].forEach((t) => C.push(B(t)));
C.push(H3("XP73 Doc of Export Management"));
["상단: Lodgement Category / Lodgement Type / Lodgement Ref No / Form No / DA·DP Ref No(XP11)", "Applicant(IE Name/Address, PAN, LEI+만료일, IEC, 계좌) / Beneficiary / Shipping / Amount(FOB·Freight·Insurance, IGST) / Additional(Remarks·Discrepancy 500자)", "Inward Remittance Ref 다건 그리드(GBS 송금 Ref 조회 → 금액 맵핑)"].forEach((t) => C.push(B(t)));
C.push(H3("신규 화면"));
["RT12 LRS: FY 누적·잔여한도 게이지, TCS 미리보기, Form A2 체크", "FD11/FD12: FDI 기한 타임라인(60/75/30), INR 확정 역산 계산기", "VS31: MT950 미리보기·발송상태, 재발송", "MN01: 기한유형별 OPEN/임박/경과 집계 + 독촉 이력", "IF01: IDPMS/EDPMS/CIMS 배치 송수신 결과·오류 재처리"].forEach((t) => C.push(B(t)));

// ---------------- 5. 데이터 ----------------
C.push(H1("5. 데이터 모델 및 관리 규칙"));
C.push(P(`총 ${M.TABLES.length}개 테이블, ${M.TABLES.reduce((a, t) => a + t.cols.length, 0)}개 컬럼. 컬럼 단위 상세(타입·NULL·관리규칙·외부보고 매핑)는 엑셀 '3.컬럼정의' 시트와 db/ddl_oracle.sql 을 기준으로 한다.`));
C.push(H2("5.1 테이블 목록"));
C.push(TBL(["테이블", "한글명", "영역", "설명"], M.TABLES.map((t) => [t.name, t.kor, t.area, t.desc]), [2.2, 1.6, 1.2, 4.2], { size: 15 }));
C.push(H2("5.2 주요 관계"));
[
  "TB_CUST_FX_PROFILE 1 : N TB_FX_REMIT_MST (송금 시 PAN/IEC/LEI 스냅샷)",
  "TB_FX_REMIT_MST 1 : N TB_FX_REMIT_PARTY / TB_FX_REMIT_FEE / TB_FX_DEAL_LINK",
  "TB_FX_REMIT_MST N : M TB_TRD_INVOICE / TB_TRD_BOE / TB_TRD_SHIPPING_BILL / TB_TRD_IRM  ← TB_TRD_PAY_ALLOC 로 해소",
  "TB_FX_REMIT_MST 1 : 0..1 TB_LRS_TXN / TB_TRD_ADV_IMP / TB_FDI_INFLOW",
  "TB_FDI_INFLOW 1 : N TB_FDI_REFUND (반환 1건 = 당발 송금 1건)",
  "TB_VOSTRO_ACCT 1 : N TB_VOSTRO_TXN, 1 : N TB_VOSTRO_STMT (TXN.STMT_ID 로 어느 명세서에 포함됐는지 추적)",
  "모든 업무 테이블 → TB_FX_OBLIGATION (SRC_TABLE + SRC_KEY 다형 참조)",
].forEach((t) => C.push(B(t)));
C.push(H2("5.3 데이터 관리 규칙 (보고서 시스템 연동 관점)"));
C.push(TBL(["항목", "관리 규칙"], [
  ["키 체계", "REMIT_REF_NO(16) = 지점4 + 업무2 + YYMM + 일련6. BoE/SB는 RBI 자연키(번호+일자+Port Code)를 유니크 인덱스로 보장"],
  ["금액", "원통화 FCY_AMT + INR_AMT + USD_EQV_AMT 3종 동시 저장(LRS·BG 기준은 USD, LEI·TCS 기준은 INR). 환산율은 거래 시점 고정"],
  ["일자", "기산일(선적일·LEO·수취일·송금일)은 반드시 별도 컬럼으로 저장, 기한은 OBLIGATION에 산출값과 사용 파라미터를 함께 저장"],
  ["잔액", "Total / Consumed / Balance 는 PAY_ALLOC 합계와 항상 일치(일배치 대사 리포트로 검증). 취소 시 ALLOC_STS=C 로 원복, 물리삭제 금지"],
  ["코드", "Purpose Code는 유효기간 관리, 미등록 코드는 OTHRS + 자유입력으로 받고 사후 정규화"],
  ["이력", "모든 테이블 감사컬럼(등록/수정/승인자). 보고 대상 테이블은 변경이력 테이블(_HIS) 생성 권장"],
  ["보고 재현성", "CIMS/IDPMS/EDPMS 제출 레코드는 TB_CIMS_RETURN_DTL·IF 테이블에 원천키로 남겨, 어떤 송금이 어느 회차로 보고되었는지 역추적"],
  ["개인정보", "PAN·Aadhaar 연계정보는 조회권한 통제 및 화면 마스킹(앞5자리), Aadhaar 번호 자체는 저장하지 않음"],
], [1.2, 5]));
C.push(H2("5.4 인보이스 · BoE 유효기간 및 서류 도착기한"));
C.push(TBL(["대상", "기산일", "기한", "시스템 처리"], [
  ["Proforma Invoice", "Invoice 발행일", "인보이스상 유효일, 미기재 시 +180일(내부정책)", "유효기간 경과 Invoice로 선수금 송금 차단"],
  ["Commercial Invoice", "Invoice 발행일", "+180일 경과 시 경고", "결제기한(선적+6M)과 연계 검증"],
  ["수입 선수금 → 물품 수입", "송금일", "+6개월 (자본재 +3년)", "OBLIGATION IMP_ADV_EVID"],
  ["수입 선수금 → 수입증빙(BoE) 제출", "수입기한", "+15일", "미제출 시 OVERDUE·IDPMS 미결 ORM"],
  ["BoE 등록일 (RT13 검증)", "송금일", "180일 이내", "초과 시 등록 오류"],
  ["수입대금 결제", "선적일(B/L·AWB)", "+6개월", "초과 시 연장근거/Trade Credit"],
  ["수출서류 제출", "LEO(수출일)", "+21일", "지연사유 입력"],
  ["수출대금 회수", "LEO", "+9개월 (해외창고 15개월, 한시완화 확인)", "연장/상각/상계, EDPMS 보고"],
  ["수출선수금 → 선적", "수취일", "+1년 (한시완화 확인)", "미이행 시 반환"],
  ["FDI 배정 / 반환 / FC-GPR", "수취일 / 수취일 / 배정일", "+60일 / +75일 / +30일", "FD11·FD12 기한관리"],
], [1.8, 1.2, 2, 2]));
C.push(H2("5.5 기한관리 매트릭스"));
C.push(TBL(["유형", "기산일", "기한 산식", "알림", "이행조건", "경과 시", "보고", "V"], M.DEADLINES.map((d) => [d[0], d[2], d[3], d[4], d[5], d[6], d[7], d[9]]), [1.2, 1, 1.8, 0.8, 1.5, 1.8, 0.7, 0.3], { verifyCol: 7, size: 14 }));

// ---------------- 6. 업무규칙 ----------------
C.push(H1("6. 업무 규칙 (검증)"));
C.push(P("처리유형: ERR = 거래차단, WARN = 경고 후 진행(승인기록), ALERT = 사후알림, AUTO = 시스템 자동처리"));
C.push(TBL(["ID", "화면", "규칙", "조건 / 로직", "처리", "V"], M.RULES.map((r) => [r[0], r[2], r[3], r[4], r[5], r[8]]), [0.9, 0.9, 1.4, 4.2, 0.6, 0.3], { verifyCol: 5, size: 14 }));

// ---------------- 7. 인터페이스 ----------------
C.push(H1("7. 대외 인터페이스"));
C.push(TBL(["IF ID", "대상", "방향", "방식", "주기", "내용"], M.INTERFACES.map((i) => [i[0], i[1], i[2], i[3], i[4], i[5]]), [0.9, 1.3, 0.6, 1.8, 0.9, 2.3], { size: 15 }));
C.push(H2("7.1 IDPMS (Import Data Processing and Monitoring System)"));
["BOD: ICEGATE BoE 수신 → IEC·AD Code 기준 고객 매핑 → Invoice No/금액으로 송금 자동매칭 후보 제시", "EOD: 당일 승인된 수입 당발 ORM(A/M/C) 생성·송신, 처리결과(ACK/NAK) 반영", "BoE 결제(Settlement)·종결(Closure)·연장·상각 보고, 미결 BoE·ORM 현황은 MN01에서 모니터링"].forEach((t) => C.push(B(t)));
C.push(H2("7.2 EDPMS (Export Data Processing and Monitoring System)"));
["BOD: Shipping Bill(LEO 포함) 수신 → TB_TRD_SHIPPING_BILL·회수기한 생성", "수출 관련 타발 전건 IRM 생성 → EDPMS 송신, e-BRC 연계", "IRM↔SB 맵핑 시 Realization 보고, 연장(EXTN)·상각(WOFF)·상계(SETO)·종결(CLSR) 보고"].forEach((t) => C.push(B(t)));
C.push(H2("7.3 RBI CIMS"));
["LRS 일일보고: 익영업일, PAN별 거래 (V_RPT_LRS_DAILY)", "R-Return(FETERS) 외환거래 Purpose별 보고", "SRVA/Vostro 보고", "TB_CIMS_RETURN에 회차·파일·ACK·반려사유 관리, 반려 시 정정 후 재제출(RESB)"].forEach((t) => C.push(B(t)));
C.push(NOTE("IDPMS·EDPMS·CIMS 파일 레이아웃(XML 스키마, 필드 길이)은 RBI가 AD 은행에 배포하는 최신 기술명세로 확정해야 한다. 본 설계의 IF 테이블은 업무 필드 기준이며, 레이아웃 확정 시 매핑표(엑셀 '3.컬럼정의'의 외부보고 매핑 열)를 상세화한다."));

// ---------------- 8. 회계 ----------------
C.push(H1("8. 회계처리"));
C.push(TBL(["이벤트", "차변", "대변", "비고"], M.GL, [1.6, 1.8, 2.6, 1.4]));

// ---------------- 9. 요청사항 ----------------
C.push(H1("9. 기존 요청사항 반영 및 검토의견"));
C.push(TBL(["출처", "기존 요청", "반영 위치", "검토의견"], M.TRACE, [1, 2.4, 2, 2.4], { size: 15 }));
C.push(H2("9.1 확인이 필요한 주요 검토의견"));
[
  "LEI 기준금액: 요청서 본문 'INR 50,000,000(5 crore)'과 경고문 'INR 50 Crore(500,000,000)'가 불일치. RBI 기준은 INR 50 crore 이므로 파라미터를 500,000,000 으로 설정하였다. 지점 확인 필요.",
  "IEC 자릿수: 요청서 '13 Digits' 표기. DGFT IEC는 PAN 기반 10자리이므로 VARCHAR2(10)으로 설계. 13자리가 다른 코드(예: AD Code 등)를 의미하는지 확인 필요.",
  "TCS: RT13(무역송금)에 TCS Y/N 요청이 있으나 TCS(206C(1G))는 LRS 개인송금 대상이므로 무역송금은 기본 N, 예외 시 수기 Y.",
  "GST: 요청은 IGST(20990016)만 인식이나 현행 CMGL 화면은 CGST/SGST 분개. 고객 GSTIN 주코드 기준 자동판정(R-CM-10)으로 양쪽 모두 지원.",
  "ISO 20022: SWIFT 크로스보더 MT103 공존기간이 2025.11 종료되었으므로 신규 개발은 pacs.008 기준. 2026.11 구조화 주소 의무화 대비 당사자 주소에 Town/Country 필수.",
  "RBI 한시완화(2025.11, 수출 회수 15개월·선수금 선적 3년): 적용 대상기간을 컴플라이언스 확인 후 TB_FX_PARAM 유효기간 행으로 등록.",
].forEach((t) => C.push(N(t)));

// ---------------- 10. 이행 ----------------
C.push(H1("10. 이행 계획(안)"));
C.push(TBL(["단계", "범위", "비고"], [
  ["1단계", "CH11 확장, RT13/RT73/XP73 개선(기존 요청 전체), TB_TRD_PAY_ALLOC 잔액엔진, 수수료·GST 체계", "테스트 중인 개선사항 우선 마무리"],
  ["2단계", "RT12 LRS + TCS + CIMS LRS 일일보고, RT11 일반송금 ISO 20022", "개인 고객 오픈"],
  ["3단계", "IDPMS/EDPMS 인터페이스, RT25, OBLIGATION 엔진·MN01", "RBI UAT 일정 연계"],
  ["4단계", "FD11/FD12 FDI, VS11/VS21/VS31 Vostro·MT950", "코레스 약정·RMA 선행"],
  ["5단계", "보고서 시스템 연동(V_RPT 뷰), TBML 연계, 이행 데이터 이관·대사", "과거 미결 BoE/SB 이관 포함"],
], [0.8, 4, 2]));
C.push(H2("테스트 시나리오(요약)"));
["LRS: FY 경계(3/31→4/1) 누적 리셋, 동시 송금 2건 한도 경합, TCS 기준 초과 직전/직후", "선수금: 자본재/일반재 기한, 부분 BoE 매칭 후 잔여 OVERDUE, BG 기준 경계값(USD 200,000)", "Invoice: 다건 분할송금 후 취소 → Balance 원복, 미래 선적일 Packing List 생략", "LEI: INR 49.99 crore + 만료(경고), 50 crore + 만료(차단)", "FDI: 61일째 배정 불가 경고, 76일째 반환 시 RBI 승인 요구, INR 확정 역산 끝전", "MT950: 0건(무거래), 250건(3페이지 분할), NAK 재발송 시 :28C: 유지"].forEach((t) => C.push(B(t)));

const doc = new Document({
  creator: "GBS India FX Project", title: "외환·무역송금 요건정의 및 설계서",
  styles: {
    default: { document: { run: { font: FONT, size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: FONT, color: NAVY }, paragraph: { spacing: { after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: FONT, color: NAVY }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 22, bold: true, font: FONT }, paragraph: { outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } },
                                 { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 260 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run("GBS India 외환·무역송금 설계서 " + M.META.version, { size: 16, color: "808080" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16 })] })] }) },
    children: [...cover, ...C],
  }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
