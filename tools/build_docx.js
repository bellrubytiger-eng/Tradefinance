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
  ["비거주자 송금", "NRI NRE/NRO/FCNR 송금 (NRO USD 1M/FY)", "외국법인·지점 송금, 외국인 급여", "RT14"],
  ["비계좌 중계", "-", "계좌 없는 수취인: 미지급금 보류 → 부분환전·INR 확정 → NEFT/RTGS", "RT21, RT22"],
  ["FDI 에스크로", "-", "지분양수도·이연대가·공개매수 에스크로 예치/지급/반환", "FD13"],
  ["환전", "외화현찰 매입·판매, 계좌간 환전", "EEFC↔INR, 외화간(KRW·USD·EUR·JPY), 지점 자기계정", "FX11, FX21"],
  ["고객 서식", "당발·타발·환전 확인서", "FIRC(FDI), 당발·타발·환전 확인서, 재발급", "DC01"],
  ["SWIFT 메시지", "-", "MT1xx·2xx·3xx·7xx·9xx ↔ MX 송수신, 수기작성, MT950 자동/수기", "SW11, SW12"],
  ["본지점 거래", "-", "RA 대지급(MT740/742), 본지점 미정산·자동정산, 기준환율·금리, 휴일 차이", "RA11, IO11, IO21, CM21"],
  ["규제보고/기한관리", "LRS 일일보고(CIMS)", "IDPMS, EDPMS, R-Return, FC-GPR 기한", "MN01, IF01"],
], [1.3, 2, 3, 1.2]));
C.push(H2("1.3 설계 원칙"));
[
  "단일 송금원장(TB_FX_REMIT_MST): 모든 송금 1건 = 1행. 보고서 시스템은 원장 + 상세 테이블만으로 모든 RBI 보고를 재현할 수 있어야 한다.",
  "스냅샷 원칙: PAN·IEC·LEI·고객유형은 송금 시점 값을 원장에 복사 저장한다(고객정보 변경 후에도 과거 보고 재현).",
  "잔액은 배분원장(TB_TRD_PAY_ALLOC)의 합계로만 갱신: Invoice·BoE·Shipping Bill·IRM의 Total / Consumed / Balance 를 하나의 서비스에서 일관 처리.",
  "규정 수치는 하드코딩 금지: TB_FX_PARAM 에 유효기간(EFF_FROM/TO)과 함께 관리하고, 기한 산출 시 사용값을 스냅샷 저장한다(수출 회수기한 15개월→9개월 변경처럼 선적일 기준 이력 적용).",
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
  ["수출입 (2026.10.1~)", "FEM (Export and Import of Goods and Services) Regulations, 2026 (FEMA 23(R)/2026-RB, 2026.9.22 개정) + RBI Directions on Export and Import, 2026", "회수기한 9개월(INR 12), 수입 계약기한·AD 재량, ₹10 lakh 신고 종결, EDF, FETERS"],
  ["여신·비자금여신", "RBI (Commercial Banks – Credit Facilities / Interest Rates on Advances / IRACP) Directions, 2025, Non-Fund Based Credit Facilities Directions, 2025", "수출신용, 할인 금리, SMA·NPA, LC·보증"],
  ["보증(FEMA)", "FEM (Guarantees) Regulations, 2026", "비거주자 관련 보증, 분기 보고"],
  ["TReDS", "RBI (TReDS) Directions, 2026", "MSME 매출채권 할인"],
  ["FDI", "FEMA (Non-Debt Instruments) Rules 2019, Mode of Payment & Reporting Regulations 2019, Press Note 3(2020)", "60일 배정, 75일 반환, FC-GPR 30일, 인접국 정부승인"],
  ["ODI", "FEMA (Overseas Investment) Rules/Regulations 2022", "RT13 Product 05"],
  ["LEI", "RBI circular - LEI for large value cross-border transactions", "비개인 INR 50 crore 이상 LEI 필수"],
  ["Vostro", "RBI - International trade settlement in INR (SRVA, 2022.7)", "SRVA 계좌유형, 사용제한"],
  ["KYC/AML", "RBI (Commercial Banks – KYC) Directions, 2025, PMLA 2002", "고객 KYC 검증, 제재국 차단, TBML 데이터"],
  ["메시지", "SWIFT CBPR+ (MT/MX 공존기간 2025.11 종료), MT950 Standard", "pacs.008 생성, 구조화 주소, MT950 생성"],
], [1, 3, 2.4]));
C.push(H2("2.2 핵심 규정 파라미터"));
C.push(TBL(["파라미터", "명칭", "값", "근거", "VERIFY"], M.PARAMS.map((p) => [p[0], p[1], `${p[2]} ${p[3]}`, p[5], p[6]]), [2.2, 2.6, 1.1, 2.6, 0.6], { verifyCol: 4, size: 15 }));

C.push(H2("2.3 RBI 최신 기준 검토 (2026-11-01 기준)"));
C.push(P("업로드된 무역금융 요청서·화면설계 v6·SFMS 요건서와 기존 설계를 RBI 최신 기준과 대조했습니다. 2026년 10월 1일부터 수출입 규정이 전면 개편되었고, 2025년 11월 28일 RBI 마스터서큘러가 'Directions, 2025'로 통합되었으며, 비자금여신(LC·보증)·FEMA 보증·TReDS 규정도 새로 시행되었습니다."));
C.push(NOTE("확인 수준: RBI 웹사이트 원문은 이번 작업 환경에서 열람할 수 없어 법률·회계법인 요약과 은행 공지로 확인했습니다(2차출처). 운영 반영 전 RBI 관보·Directions 원문으로 최종 대조가 필요합니다. '미확인' 항목은 출처를 찾지 못해 기존 값을 지점 정책으로 유지한 것입니다."));
C.push(TBL(["No", "대상", "기존 내용", "RBI 최신 기준", "반영", "확인"], M.RBI_REVIEW.map((r) => [r[0], r[1], r[2], r[3], r[4], r[6]]), [0.3, 0.9, 1.5, 2.6, 1.7, 0.6], { size: 13 }));

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
  ["주문 진정성·선수금 필요성 평가 기록, 지점 정책 기준(USD 200,000) 초과 시 BG/SBLC 요구 (2026 규정상 금액 기준은 폐지)", "R-IMP-06"],
  ["승인 → Invoice Consumed/Balance 갱신, ORM(A) 생성", "TB_TRD_PAY_ALLOC, TB_IF_IDPMS_ORM"],
  ["선수금 기한 생성: 송금일+6M(자본재 36M)+15일", "TB_TRD_ADV_IMP, TB_FX_OBLIGATION"],
  ["IDPMS BoE 수신 → 자동/수동 매칭 → 증빙액 갱신 → 완결 시 FULFILLED, Settlement 보고", "IF-IDP-02/03"],
  ["기한 임박 D-30/15/7/0 독촉, 경과 시 OVERDUE 보고", "MN01"],
]);
flow("3.3 수입대금 사후송금 / Direct Import (Product 02·03)", [
  ["BoE 존재 확인(IDPMS 수신분 우선, 미수신 시 RT73 수기등록)", "TB_TRD_BOE"],
  ["RT13 Amount Limit 탭: Invoice Total/Consumed/Balance, BoE Ref/Registration Date/Value 표시", "R-IMP-08/09/10"],
  ["계약상 결제기한 검증(미기재 시 선적+6개월 모니터링), 초과 시 AD 연장 사유 또는 무역신용 판정", "R-IMP-05"],
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
C.push(H3("FDI 송금인 KYC와 MT199 요청 (R-KYC-01~08)"));
C.push(P("FDI 유입에서 AD은행의 기본 KYC 의무는 송금인(비거주 투자자)에 대한 것입니다. 수취인(피투자법인)은 당행 고객 CDD로 갈음하고, 금액이 크거나 고위험일 때만 추가 실사(EDD)를 합니다."));
C.push(TBL(["단계", "처리", "데이터/메시지"], [
  ["1", "RT21 → FD11 이관 시 송금 전문에 KYC 보고서가 첨부되었는지 확인 (ATTACHED / NONE)", "TB_FDI_KYC.KYC_STS"],
  ["2", "미첨부면 송금은행(pacs.008 DbtrAgt, 없으면 송신은행) 앞 MT199 KYC 요청 초안 자동 생성. :20: KYC+FDI Ref, :21: 상대 Ref, :79: 결제 식별정보와 요청 항목 (35줄 × 50자 검증)", "TB_SWIFT_MSG(MT199), R-KYC-02"],
  ["3", "SW11에서 Checker 승인 후 송신 → SENT. 응답이 없으면 인도 영업일 D+3 재요청, D+7 책임자 보고", "R-KYC-03"],
  ["4", "응답 MT199의 :21:이 요청 :20:과 같으면 FD11에 자동 연결 (RECEIVED). 서명된 KYC 보고서는 TB_FX_DOC(KYC)에 등록", "R-KYC-04"],
  ["5", "검증: 명칭·주소가 전문 송금인과 일치, 설립국·등록번호, 거래개시일, 지분 10% 이상 실소유자와 국적, 자금원천, PEP, 제재 스크리닝. 누락 시 INCOMPLETE → 보완 MT199", "R-KYC-05"],
  ["6", "투자자나 실소유자가 육상국경 인접국이면 Government Route와 승인번호 필수 (Press Note 3)", "R-FDI-01"],
  ["7", "VERIFIED 전에는 FIRC 발급·배정 확인·FC-GPR 지원 불가. 고위험국이거나 FDI_HOLD_UNTIL_KYC=Y면 입금 대신 미지급금 보류", "R-KYC-06"],
  ["8", "INR 10 crore 이상이거나 고위험(인접국 실소유자·PEP)이면 수취인 EDD: 실소유자·투자자금 용도·사업 실체", "R-KYC-07"],
  ["9", "제재 HIT·불일치는 FAILED → 반환(FD12 REJECT)·STR 검토. 반환기한(+75일)까지 미검증이면 반환 대상", "R-KYC-08"],
], [0.4, 5, 1.6], { size: 16 }));
C.push(P("SWIFT는 2025년 11월부터 조사 메시지용 camt.110/111을 도입했습니다. 기본은 요청하신 MT199로 하고, MX 전환 범위는 SWIFT 공지로 확인합니다(VERIFY).", { size: 18, italics: true }));
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

flow("3.7 비계좌 수취인 중계 (RT21 → RT22)", [
  ["타발 수신, 수취인이 당행 계좌 없음 → SETTLE_MODE=SUSP, Remittance Payable(미지급금) 보류", "TB_FX_SUSPENSE, R-SUS-01"],
  ["수취인 식별(PAN/CIN)·이름 확인, 수출·FDI 목적이면 IEC/CIN 확인", "R-SUS-04"],
  ["RT22: 환전 레그 등록(부분환전 가능, INR 확정 레그는 FCY = ROUNDUP(목표 INR/환율, 2)), 레그별 딜 맵핑", "TB_FX_CONV_LEG, R-SUS-02"],
  ["NEFT/RTGS 지급(INR 2 lakh 미만은 NEFT), IFSC·계좌 검증, UTR 수신", "TB_DOM_PAYOUT, R-DOM-01~03"],
  ["잔여 외화: 계속 보류 / 타 AD 외화이체 / 송금인 반환(RT41). 15영업일 경과 시 알림", "R-SUS-03"],
  ["타발송금 확인서(FM-IN-01 relay) 발급, FDI면 FIRC에 UTR·수취은행 표기", "FM-IN-01, FM-FDI-01"],
]);
flow("3.8 FDI 에스크로 (FD13)", [
  ["약정서·당사자 KYC·대상회사 CIN 확인 후 에스크로 계좌 개설(무이자)", "TB_FDI_ESCROW, R-ESC-01"],
  ["이연대가: 예치한도 ≤ 총대가 25%, 만기 ≤ 약정일 + 18개월", "R-ESC-02"],
  ["RT21 SETTLE_MODE=ESCROW로 예치", "TB_FDI_ESCROW_TXN(DEPO)"],
  ["조건 충족: 공동지시서 등 근거 입력 → 거주자 지급(계좌/NEFT/RTGS), FC-TRS 기한(+60일) 생성", "R-ESC-03"],
  ["불성립·만기: 비거주자 원통화 반환(당발 Product 07)", "R-ESC-04"],
]);
flow("3.9 비거주자 송금 (RT14)", [
  ["계좌 선택: NRE·FCNR 잔액은 송금 자유, NRO는 자금 성격 구분(경상소득 / 잔액·자산매각)", "R-NR-01"],
  ["NRO 잔액·자산매각: FY USD 1M 누적 검증(TB_NR_REPAT_SUM), Form 15CA·15CB·원천 증빙", "R-NR-01/02"],
  ["LRS·TCS 비적용, GST 공급장소 = 공급자 소재지(s.13(8)(a)) → CGST+SGST", "R-NR-04, R-GST-04"],
  ["타발 NRE 입금은 해외송금·NRE/FCNR 이체만 허용", "R-NR-03"],
]);
flow("3.10 환전 (FX11)", [
  ["유형 선택: 계좌간(EEFC↔INR) / 외화간(KRW·USD·EUR·JPY) / 외화현찰 매입·판매 / 지점 자기계정", "TB_FX_EXCHANGE"],
  ["EEFC 입금 제한: INR을 매도해 EEFC로 입금 불가. EEFC→INR, EEFC 내 외화간 환전은 허용", "R-FX-02"],
  ["환율: USD 25,000 이상·지점 자기계정은 자금팀 딜, 미만은 고시환율 회차(FX21)", "R-FX-04"],
  ["현찰 판매(여행): USD 3,000/건, PAN·Form A2, LRS 합산 / 현찰 매입: CDF(USD 5,000 초과), INR 현금지급 한도", "R-FX-05/06"],
  ["GST: 환전가치 32(2) 과세표준 × 18%, 외화간은 RBI 기준율 INR 환산액(작은 쪽) 기준", "R-GST-02, R-FX-07"],
  ["환전 확인서(FM-FX-01) 발급, 현찰 매입은 Encashment Certificate 문구 포함", "FM-FX-01"],
]);

flow("3.11 SWIFT 메시지 처리 (SW11 / SW12)", [
  ["업무화면 승인 시 메시지 자동생성(당발 pacs.008, 대지급·정산 pacs.009, MT756, MT950, MT300) → SW11 큐(AUTH 대기)", "TB_SWIFT_MSG, R-SW-01"],
  ["수기 작성은 SW12 템플릿(MT950·MT202/pacs.009·MT300·MT999 등), 필드 형식 검증 후 VERIFY", "R-SW-02/06"],
  ["Checker 승인(작성자≠승인자, 수기 MT950·n99는 2인) → RMA·제재 필터 → 송신", "R-SW-03/04"],
  ["ACK는 완료, NAK는 Repair 큐 → 수정·재검증 → 재발송(PDE)", "R-SW-07"],
  ["수신 메시지 파싱 후 유형별 연결: pacs.008/MT103→RT21, pacs.009/MT202·MT200→VS21·Nostro 대사, camt.054/MT910→본지점·Nostro 대사, MT740/747/742→RA11, MT300→딜 확인", "R-SW-05"],
]);
flow("3.12 RA 대지급과 본지점 정산 (RA11 → IO21)", [
  ["본점 개설 신용장의 MT740 수신 → RA 등록(잔여 수권액 관리), MT747로 변경", "TB_RA_AUTH, R-RA-01"],
  ["MT742 청구 접수 → 수권액·유효기일·청구은행 검증, 지급기한 = 다음날부터 인도 영업일 3일", "TB_RA_CLAIM, R-RA-02/03"],
  ["결제일: USD는 US·IN 모두 영업일, INR은 IN 영업일. 휴일이면 다음 영업일로 이동하고 본점 고객 결제일과의 차이 일수 기록", "R-RA-04, R-HOL-01"],
  ["지급(pacs.009 또는 RTGS) → 본점 MT756 통보 → Due from HO 미정산(TB_IO_TXN) 생성", "R-RA-05"],
  ["IO21: 통화별 OPEN 건수 ≥ 30, 순액 ≥ USD 5백만, 또는 주기(W-FRI) 도래 시 정산배치 생성. Cutoff까지의 OPEN 건만 포함", "TB_IO_SETTLE_BATCH, R-IO-01/02"],
  ["채권-채무 상계 후 순액 청구, 정산통화가 다르면 기준환율 ± 스프레드(IO11) 적용", "R-IO-03/04"],
  ["본점 입금(pacs.009/camt.054) 대사 → 일치 시 SETTLED, 부족 시 PART·DISPUTE. 2회차 이상 미정산은 보고", "R-IO-05/06"],
]);
flow("3.13 휴일 차이 처리 (CM21 → IO21)", [
  ["IN·US·KR(+TARGET2) 캘린더 관리, 인도 임시휴일은 공표 즉시 등록", "TB_HOLIDAY"],
  ["등록 시 해당일 결제 예정 RA 청구·당발·정산배치 자동 재계산(지급기한·결제일)", "R-HOL-02"],
  ["본점 고객 결제일과 지점 결제일 차이 → 이자 = 금액 × (기준금리+스프레드) × 일수 / 일수기준", "TB_IO_FLOAT_ADJ, R-IO-07"],
  ["건별 USD 300 이하: 자금부담 점이 자체손익으로 흡수(월 합산 내부대체 전표) / 초과: 정산배치에 FLOAT 항목으로 본지점 정산", "R-IO-08"],
  ["고객에게 개별 청구하지 않고, 영향 건 목록을 본점에 MT999(또는 정산파일)로 한 번에 통보", "R-HOL-03"],
]);

C.push(H2("3.14 본지점 계정의 기금 구분과 금리"));
C.push(P("본지점 거래는 목적에 따라 세 가지 기금으로 나누고, 모든 본지점 원장(TB_IO_TXN)과 금리표(TB_IO_RATE)에 기금 구분(IO_FUND)을 필수로 둡니다."));
C.push(TBL(["구분", "대상 거래", "금리", "정산", "근거 규칙"], [
  ["갑기금 (Capital)", "본점이 배정한 영업기금", "무이자 (0%)", "정산배치 제외. 증감은 본점 승인 문서 필수, 자본 변동 RBI 보고 대상 확인", "R-IO-11"],
  ["을A (거래성)", "RA 대지급, 본지점 송금·환전, 수수료 배분, 휴일차이 이자", "통화별 기준금리 + 스프레드 (예: USD Term SOFR + 50bp, ACT/360)", "IO21 정산배치 (건수·금액·주기 트리거). 휴일차이와 2회차 초과 미정산 잔액에 이자", "R-IO-12"],
  ["을B (여신성)", "본점 차입·예치 (여신 재원, 유동성)", "기준금리 + 기간구간 스프레드 (≤7일 ~ 1년 초과), 우대조건 충족 시 우대 스프레드", "IO31 거래별 만기 상환·이자 지급(만기/월). 을A 배치와 분리. MT320 확인", "R-IO-13~16"],
], [1, 1.6, 1.8, 2.2, 0.8], { size: 15 }));
C.push(H3("을B 우대조건 판정 (R-IO-14)"));
[
  "조회기간: 차입 시작일 기준 과거 365일 이내에 시작한 머니마켓 거래 (GBS 자금 거래를 TB_MM_DEAL_REF로 매일 복제)",
  "인정 거래: 정기예금, 콜머니, 콜론, 기간물, 통지예금 (취소 거래 제외)",
  "조건: 같은 통화, 금액 ≥ 차입금액, 거래기간 ≥ 30일. 요청서의 '동일 금액'은 '이상'으로 해석했으며 파라미터로 '정확히 동일'로 바꿀 수 있음",
  "충족 거래가 여러 건이면 가장 최근 시작 거래를 근거로 저장(PREF_DEAL_REF), 후보·탈락 사유 전체를 PREF_CHECK_JSON에 스냅샷",
  "연장(rollover)은 새 거래로 등록하고 우대 여부를 다시 판정",
  "예: USD 5,000,000, 92일(T180) → 기준 4.30% + 우대 45bp = 4.75%, 이자 60,694.44 (ACT/360). 이력이 없으면 표준 75bp",
].forEach((t) => C.push(B(t)));

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

C.push(H2("4.3 FDI 업무 화면 이용 가이드"));
C.push(P("FDI 유입은 RT21(타발 등록)에서 시작해 FD11로 자동 이관됩니다. 무역·일반 당발 화면(RT11·RT13)은 FDI 유입에 사용하지 않습니다. 기존 설계서의 'FD13 INR 확정 환산' 표기는 오기였으며, INR 확정 환산은 FD11(당행 계좌)과 RT22(비계좌)에서 처리합니다. FD13은 이번에 FDI 에스크로 화면으로 신설했습니다."));
C.push(TBL(["단계", "상황", "사용 화면", "처리", "생성 데이터/서식"], M.FDI_FLOW, [0.5, 1.6, 1.5, 3, 1.5], { size: 15 }));
C.push(H2("4.4 고객 서식"));
C.push(P("고객 서식 4종의 항목·DB 소스·문구는 별도 문서 '03_고객서식_명세서.docx'와 screens/forms.html(A4 인쇄 시안)에 정리했습니다."));
C.push(TBL(["서식ID", "서식명", "발급 화면", "발급 시점", "용도"], M.FORMS.map((f) => [f.id, f.ko, f.screen, f.when, f.use]), [0.9, 1.6, 1.3, 2.2, 2]));

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
  ["수입 선수금 → 물품 수입", "송금일", "+6개월 (자본재 +3년) - 지점 정책, 2026 Directions 확인", "OBLIGATION IMP_ADV_EVID"],
  ["수입 선수금 → 수입증빙(BoE) 제출", "수입기한", "+15일", "미제출 시 OVERDUE·IDPMS 미결 ORM"],
  ["BoE 등록일 (RT13 검증)", "송금일", "180일 이내", "초과 시 등록 오류"],
  ["수입대금 결제", "계약 결제기한", "계약상 기한 (고정 6개월 폐지)", "AD 연장 사유 / 무역신용"],
  ["수출서류 제출", "LEO(수출일)", "+21일", "지연사유 입력"],
  ["수출대금 회수", "선적(LEO)일 / 서비스 인보이스일", "+9개월 (INR 인보이스·결제 12개월)", "AD 정책 연장·감액·상계, EDPMS 보고"],
  ["수출선수금 → 선적", "수취일", "+3년", "미이행 시 반환"],
  ["서비스 수출 EDF", "인보이스 월말", "+30일", "EDF 제출"],
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
C.push(H2("7.4 SWIFT 메시지 카탈로그"));
C.push(P("요청하신 MT200·MT300·MT900대를 포함해 Vostro·본지점·RA 업무에 필요한 메시지를 모두 지원합니다. 2025년 11월 CBPR+ MT/MX 공존기간 종료 이후 결제·은행이체·입출금 통지는 MX로 송신하고 MT는 수신 번역·레거시 처리용으로 유지합니다. 명세서(MT940/950), 외환·자금 확인(MT300/320), 무역(MT7xx)은 MT로 송수신합니다."));
C.push(TBL(["유형", "Cat", "명칭", "대응", "송신", "자동생성", "수신 연결"], M.SWIFT_TYPES.map((r) => [r[0], r[2], r[3], r[4], r[5], r[7], r[8]]), [0.9, 0.4, 2.2, 1.1, 0.5, 1.8, 1.2], { size: 14 }));
C.push(H2("7.5 Vostro 용도별 처리"));
C.push(TBL(["용도", "계좌주 / 최종고객", "처리"], [
  ["HO 본지점", "본점 INR 계좌", "거래마다 본지점 원장(TB_IO_TXN) 동시 생성 → IO21 정산. RA 대지급 INR 지급과 상계 가능"],
  ["FPI 증권투자", "본국 은행 → 신탁회사·기관투자자", "최종투자자 SEBI FPI 등록번호·커스터디 확인, FPI 지정계좌(SNRR·커스터디)로만 지급 (VERIFY)"],
  ["CORP 법인 송금", "본국 은행 → 한국계 법인", "Vostro 출금으로 인도 지사·자회사에 입금 = 타발송금. Purpose에 따라 FDI(FD11)·IRM·일반 분기, FIRC 발급"],
  ["TRADE / EXCH", "SRVA 코레스 / Exchange House", "INR 무역결제·근로자 송금"],
], [1, 1.8, 4]));
C.push(NOTE("IDPMS·EDPMS·CIMS 파일 레이아웃(XML 스키마, 필드 길이)은 RBI가 AD 은행에 배포하는 최신 기술명세로 확정해야 한다. 본 설계의 IF 테이블은 업무 필드 기준이며, 레이아웃 확정 시 매핑표(엑셀 '3.컬럼정의'의 외부보고 매핑 열)를 상세화한다."));

// ---------------- 8. GST ----------------
C.push(H1("8. GST 처리 기준"));
C.push(P("기존 설계는 '수수료 × 18%, 고객 GSTIN 주코드에 따라 CGST+SGST 또는 IGST'만 정의했습니다. 이번에 환전가치 과세, 비거주자·비계좌 공급장소, 반올림, INR 확정 환전의 별도징수를 추가했습니다."));
C.push(H2("8.1 과세표준"));
C.push(TBL(["항목", "과세 여부", "비고"], [
  ["Commission, Cable charge, In lieu of Exchange, NEFT/RTGS 수수료", "과세 18%", "SAC 9971xx (세무팀 확정)"],
  ["환전가치 (CGST Rules 32(2))", "과세 18%", "방식 B slab (FY 단위 선택, 파라미터 GST_FX_METHOD)"],
  ["TCS, 해외은행 공제수수료, 송금 원금", "비과세", "GST 계산에서 제외"],
  ["EEFC·FCNR 외화 그대로 송금/입금 (무환전)", "환전가치 없음", "수수료분만 과세"],
  ["지점 자기계정 환전", "비대상", "대고객 공급 아님"],
], [3, 1, 2.4]));
C.push(H2("8.2 환전가치 slab (32(2)(b))"));
C.push(TBL(["환전 INR 총액", "과세표준", "예시"], [
  ["≤ INR 1,00,000", "1% (최소 INR 250)", "INR 50,000 → 500"],
  ["1,00,000 ~ 10,00,000", "1,000 + 0.5% × (금액 - 1,00,000)", "INR 3,61,050 → 2,305.25"],
  ["> 10,00,000", "5,500 + 0.1% × (금액 - 10,00,000), 최대 60,000", "INR 88,42,000 → 13,342 / 25 crore → 60,000"],
  ["외화간(INR 미개입)", "두 통화를 RBI 기준율로 INR 환산한 금액 중 작은 금액 기준", "USD 20,000 → KRW: INR 17,68,348 → 6,268.35"],
], [1.6, 3, 2.4]));
C.push(H2("8.3 공급장소와 세목"));
C.push(TBL(["수령자", "공급장소", "세목", "근거"], [
  ["거주자 계좌고객", "은행 기록상 고객 주 (GSTIN 앞 2자리, 없으면 주소)", "UP(09)면 CGST 9% + SGST 9%, 타주면 IGST 18%", "IGST Act s.12(12)"],
  ["거주자 비계좌", "공급자(지점) 소재지", "CGST + SGST", "IGST Act s.12(12)"],
  ["비거주 계좌고객 (NRE/NRO/FCNR, 외국법인)", "공급자(지점) 소재지", "CGST + SGST", "IGST Act s.13(8)(a)"],
  ["해외 비계좌 수령자 (OUR 수수료 청구 해외은행 등)", "국외", "수출용역 영세율 (LUT)", "IGST Act s.13(2), s.16"],
], [1.8, 1.8, 2, 1.2], { size: 16 }));
C.push(H2("8.4 계산 순서와 반올림"));
["과세표준 = Σ과세 수수료 + 환전가치(32(2))", "GST 합계 = ROUND(과세표준 × 18%, 2)", "주내: CGST = ROUND(GST/2, 2), SGST = GST - CGST (0.01 차이 허용) / 주간: IGST = GST 합계",
 "INR 확정 환전·FDI: 수수료와 GST는 별도 출금(FEE_COLLECT_MODE=SEPARATE). 원금 INR에서 공제하지 않음",
 "확인서에 지점 GSTIN·고객 GSTIN·SAC·공급장소를 표기하고, 월별 통합 세금계산서(Rule 54(2)) 대상으로 집계",
 "계산 예시는 엑셀 '7.기한계산기' GST 섹션(수식)과 각 확인서 시안에서 확인할 수 있음"].forEach((t) => C.push(B(t)));
C.push(NOTE("비거주자·해외 수령자 공급장소, SAC 코드, 외화간 환전 과세표준 산정(방식 B 적용 시)은 세무팀 확인이 필요합니다(VERIFY)."));

// ---------------- 9. 회계 ----------------
C.push(H1("9. 회계처리"));
C.push(TBL(["이벤트", "차변", "대변", "비고"], M.GL, [1.6, 1.8, 2.6, 1.4]));

// ---------------- 9. 요청사항 ----------------
C.push(H1("10. 기존 요청사항 반영 및 검토의견"));
C.push(TBL(["출처", "기존 요청", "반영 위치", "검토의견"], M.TRACE, [1, 2.4, 2, 2.4], { size: 15 }));
C.push(H2("10.1 확인이 필요한 주요 검토의견"));
[
  "LEI 기준금액: 요청서 본문 'INR 50,000,000(5 crore)'과 경고문 'INR 50 Crore(500,000,000)'가 불일치. RBI 기준은 INR 50 crore 이므로 파라미터를 500,000,000 으로 설정하였다. 지점 확인 필요.",
  "IEC 자릿수: 요청서 '13 Digits' 표기. DGFT IEC는 PAN 기반 10자리이므로 VARCHAR2(10)으로 설계. 13자리가 다른 코드(예: AD Code 등)를 의미하는지 확인 필요.",
  "TCS: RT13(무역송금)에 TCS Y/N 요청이 있으나 TCS(206C(1G))는 LRS 개인송금 대상이므로 무역송금은 기본 N, 예외 시 수기 Y.",
  "GST: 요청은 IGST(20990016)만 인식이나 현행 CMGL 화면은 CGST/SGST 분개. 고객 GSTIN 주코드 기준 자동판정(R-CM-10)으로 양쪽 모두 지원.",
  "ISO 20022: SWIFT 크로스보더 MT103 공존기간이 2025.11 종료되었으므로 신규 개발은 pacs.008 기준. 2026.11 구조화 주소 의무화 대비 당사자 주소에 Town/Country 필수.",
  "GST: 기존 설계에 환전가치(32(2)) 과세, 비거주자·비계좌 공급장소, 반올림 규칙이 빠져 있어 R-GST-01~08로 보완. 현행 RT13은 Commission에만 GST를 부과하고 Cable charge는 제외되어 있음 → 과세 대상에 포함.",
  "EEFC: INR을 매도해 EEFC로 입금하는 환전은 허용되지 않는 것으로 설계(R-FX-02). 지점 운영 관행과 다르면 확인 필요.",
  "미지급금 보관기한 15영업일(SUSP_MAX_DAYS)은 내부정책 가정값. 지점 기준 확정 필요.",
  "본지점 정산 기준(주기 W-FRI, 자동정산 30건·USD 5백만, 휴일 차이 흡수한도 건당 USD 300, 기준금리·스프레드)은 본점 자금부와 약정으로 확정 필요. 모두 파라미터로 관리.",
  "SWIFT: MT940/950·Category 3·7의 MX 전환 일정은 SWIFT 공지로 재확인 필요(VERIFY). 시스템은 카탈로그의 '송신 형식'만 바꾸면 전환되도록 설계.",
  "FDI KYC: 기본 의무는 송금인. 수취인 EDD 기준(INR 10 crore)과 KYC 확인 전 입금 보류 여부(FDI_HOLD_UNTIL_KYC=N)는 내부정책 가정값. MT199 요청 문안은 컴플라이언스 검토 필요.",
  "본지점 기금: 갑기금 무이자 / 을A 거래성 / 을B 여신성으로 구분. 을B 우대조건의 '동일 금액'은 '이상'으로 해석(파라미터 IO_PREF_AMT_RULE), 기간구간 스프레드 값은 예시이므로 본점 자금부 확정 필요.",
  "RA 휴일 차이: 지점이 늦게 지급한 경우(본점 선수취)와 먼저 지급한 경우(지점 자금부담)를 부호(±)로 구분해 이자 방향을 정함. 고객 청구는 하지 않음(요청하신 방안 반영).",
  "2026.10.1 수출입 규정 전면 개편(9개월·계약기한·AD 재량)은 2차 출처로 확인했으며, RBI 관보 원문(FEMA 23(R)/2026-RB 및 23(R)/(1)/2026-RB)과 Directions 원문 대조가 필요함. 상세는 2.3절.",
].forEach((t) => C.push(N(t)));

// ---------------- 10. 이행 ----------------
C.push(H1("11. 무역금융(여신) 모듈 통합"));
C.push(P("지점이 작성한 무역금융 요청서(어음할인·매입·D/A·D/P·오픈어카운트·TReDS, M1~M8)와 화면설계 v6, SFMS 인터페이스 요건서를 이 설계에 통합했습니다. 업로드 자료는 LC 기반 매입까지만 다루고 LC 개설·보증 모듈은 '별도 모듈'로 남겨 두었으므로, 신용장(LC11)과 보증(BG11)은 RBI 최신 기준으로 새로 설계했습니다."));
C.push(H2("11.1 처리 흐름"));
C.push(TBL(["단계", "처리", "데이터"], M.TF_FLOW, [1.4, 4, 1.6], { size: 16 }));
C.push(H2("11.2 기존 설계와의 통합 지점"));
[
  "A/P 단일화: 할인 순지급액과 비계좌 타발 보류를 모두 GL 21100000 미지급금(TB_AP_ITEM)으로 관리. 기존 'Remittance Payable 신규 GL'은 폐지",
  "SFMS 단일 경로: 무역금융 A/P 지급, RT22 타발 중계, FDI·에스크로 지급 모두 Velocis 미들웨어 → SFMS(TB_SFMS_MSG). 규칙 ID는 지점 SFMS 요건서(R-VAL·RTE·CTL·EXC·MCH·IFN)를 그대로 사용",
  "EDPMS·IDPMS: 무역금융 거래(TB_TF_BILL)는 기존 Shipping Bill·BoE·IRM 원장과 연결해 같은 보고 경로 사용",
  "RA·본지점: 본점 개설 LC의 상환(RA11)은 기존 본지점 정산(IO21)으로, 지점 개설 LC의 상환은 LC11에서 MT740 발행",
  "개설은행 평가: LC 매입 신용 프리미엄과 단일은행 한도에 사용, 734 거절(SWIFT·SFMS) 자동 적재",
  "GST: 할인료(이자)는 비과세, 수수료·전신료·보증료는 R-GST 규칙으로 과세",
].forEach((t) => C.push(B(t)));
C.push(H2("11.3 신용장·보증 핵심 규칙"));
C.push(TBL(["규칙", "내용"], M.RULES.filter((r) => /^R-(LC|BG)-/.test(r[0])).map((r) => [r[0] + " " + r[3], r[4]]), [1.4, 5], { size: 15 }));

C.push(H1("12. 이행 계획(안)"));
C.push(TBL(["단계", "범위", "비고"], [
  ["1단계", "CH11 확장, RT13/RT73/XP73 개선(기존 요청 전체), TB_TRD_PAY_ALLOC 잔액엔진, 수수료·GST 체계", "테스트 중인 개선사항 우선 마무리"],
  ["2단계", "RT12 LRS + TCS + CIMS LRS 일일보고, RT11 일반송금 ISO 20022, RT14 비거주자, FX11/FX21 환전, GST 엔진(R-GST)", "개인 고객 오픈"],
  ["3단계", "IDPMS/EDPMS 인터페이스, RT25, OBLIGATION 엔진·MN01, RT22 미지급금·NEFT/RTGS 연계", "RBI UAT 일정 연계"],
  ["4단계", "FD11/FD12/FD13 FDI·에스크로, VS11/VS21/VS31 Vostro·MT950, SW11/SW12 SWIFT, RA11·IO11·IO21·CM21 본지점, DC01 서식 발급", "코레스 약정·RMA 선행"],
  ["5단계", "보고서 시스템 연동(V_RPT 뷰), TBML 연계, 이행 데이터 이관·대사", "과거 미결 BoE/SB 이관 포함"],
], [0.8, 4, 2]));
C.push(H2("테스트 시나리오(요약)"));
["LRS: FY 경계(3/31→4/1) 누적 리셋, 동시 송금 2건 한도 경합, TCS 기준 초과 직전/직후", "선수금: 자본재/일반재 기한, 부분 BoE 매칭 후 잔여 OVERDUE, 지점 BG 정책 경계값(USD 200,000)", "Invoice: 다건 분할송금 후 취소 → Balance 원복, 미래 선적일 Packing List 생략", "LEI: INR 49.99 crore + 만료(경고), 50 crore + 만료(차단)", "FDI: 61일째 배정 불가 경고, 76일째 반환 시 RBI 승인 요구, INR 확정 역산 끝전", "MT950: 0건(무거래), 250건(3페이지 분할), NAK 재발송 시 :28C: 유지"].forEach((t) => C.push(B(t)));

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
