// 본점 담당자용 — 인도 총계정원장 검토 가이드 (워드)
// node build_guide_docx.js <guide.json> <out.docx>
// 데이터: python3 export_guide_data.py <guide.json>
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType,
  HeadingLevel, TableOfContents, Header, Footer, PageNumber, BorderStyle, LevelFormat, PageBreak, VerticalAlign,
} = require("docx");

const D = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = process.argv[3] || "../docs/07_인도_총계정원장_검토가이드_본점용.docx";
const FONT = "맑은 고딕";
const W = 9638; // A4, 좌우 여백 2cm
const NAVY = "1F4E79", ORANGE = "C55A11", GREY = "595959";
const FILL = { navy: "1F4E79", lnavy: "DDEBF7", lorange: "FCE4D6", lgreen: "E2EFDA", lyellow: "FFF2CC", lgrey: "F2F2F2", white: "FFFFFF" };

// ---------- 기본 요소 ----------
function runs(x, base = {}) {
  if (typeof x === "string") return [new TextRun({ text: x, font: FONT, ...base })];
  return x.map((s) => (typeof s === "string" ? new TextRun({ text: s, font: FONT, ...base }) : new TextRun({ font: FONT, ...base, text: s.t, bold: s.b, color: s.c, italics: s.i })));
}
const P = (x, o = {}) => new Paragraph({ children: runs(x, { size: o.size || 20, color: o.color }), alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 120, line: 300 }, keepNext: o.keepNext });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: FONT })] });
const B = (x, lvl = 0) => new Paragraph({ children: runs(x, { size: 20 }), numbering: { reference: "bul", level: lvl }, spacing: { after: 60, line: 290 } });
const N = (x) => new Paragraph({ children: runs(x, { size: 20 }), numbering: { reference: "num", level: 0 }, spacing: { after: 60, line: 290 } });
const CAP = (t) => new Paragraph({ children: [new TextRun({ text: t, font: FONT, size: 18, bold: true, color: GREY })], spacing: { before: 120, after: 80 }, alignment: AlignmentType.CENTER, keepNext: true });
const BR = () => new Paragraph({ children: [new PageBreak()] });
const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };
const nob = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: nob, bottom: nob, left: nob, right: nob };

function cell(x, width, o = {}) {
  const paras = (Array.isArray(x) && x.length && Array.isArray(x[0]) ? x : [x]).map((t) => new Paragraph({
    children: runs(t === null || t === undefined ? "" : (typeof t === "number" ? t.toLocaleString("en-IN") : t), { size: o.size || 17, bold: o.bold, color: o.color }),
    alignment: o.align, spacing: { after: 20, line: 260 },
  }));
  return new TableCell({
    children: paras, width: { size: width, type: WidthType.DXA }, borders: o.borders || borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 }, verticalAlign: o.valign || VerticalAlign.TOP, columnSpan: o.span,
  });
}
function table(headers, rows, widths, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const hdr = new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { fill: o.hfill || NAVY, color: "FFFFFF", bold: true, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER })) });
  const body = rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((v, i) => cell(v, widths[i], { fill: (o.zebra && ri % 2) ? FILL.lgrey : (o.colFill && o.colFill[i]) })) }));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [hdr, ...body] });
}
function note(x, fill = FILL.lnavy, label = "참고") {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [cell([[{ t: label + "  ", b: true, c: NAVY }, ...(typeof x === "string" ? [x] : x)]], W, { fill, size: 19 })] })] });
}

// ---------- 순서도 (표로 그린 도형: 본점 PC에서 편집 가능) ----------
// item: {box, fill} | {arrow} | {dec, yes, no} | {row:[...], fill}
function flow(items, width = 7200) {
  const side = Math.round((W - width) / 2);
  const cols = [side, width, side];
  const rows = [];
  for (const it of items) {
    if (it.arrow) {
      rows.push(new TableRow({ children: [cell("", side, { borders: noBorders }), cell([[{ t: it.arrow === true ? "▼" : "▼  " + it.arrow, c: GREY }]], width, { borders: noBorders, align: AlignmentType.CENTER, size: 16 }), cell("", side, { borders: noBorders })] }));
    } else if (it.dec) {
      rows.push(new TableRow({ cantSplit: true, children: [
        cell([[{ t: "예 ◀ ", b: true, c: "2E7D32" }, it.yes]], side, { fill: FILL.lgreen, size: 16 }),
        cell([[{ t: "◆ ", c: ORANGE }, { t: it.dec, b: true }]], width, { fill: FILL.lyellow, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER }),
        cell([[it.no, { t: " ▶ 아니오", b: true, c: "C00000" }]], side, { fill: FILL.lorange, size: 16 }),
      ] }));
    } else if (it.row) {
      const n = it.row.length, w = Math.floor(W / n);
      const widths = it.row.map((_, i) => (i === n - 1 ? W - w * (n - 1) : w));
      rows.push(new TableRow({ cantSplit: true, children: it.row.map((t, i) => cell([[{ t, b: i === 0 && it.firstBold }]], widths[i], { fill: it.fill || FILL.lnavy, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER, size: 16 })).concat([]) }));
      continue;
    } else {
      rows.push(new TableRow({ cantSplit: true, children: [cell("", side, { borders: noBorders }), cell([[{ t: it.box, b: !!it.bold }]], width, { fill: it.fill || FILL.lnavy, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER }), cell("", side, { borders: noBorders })] }));
    }
  }
  // row 항목은 열 수가 달라 별도 표로 분리
  return rows;
}
function flowTables(items, width) {
  const out = [];
  let buf = [];
  const push = () => { if (buf.length) { out.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [Math.round((W - (width || 7200)) / 2), width || 7200, W - 2 * Math.round((W - (width || 7200)) / 2)], rows: flow(buf, width) })); buf = []; } };
  for (const it of items) {
    if (it.row) {
      push();
      const n = it.row.length, w = Math.floor(W / n);
      const widths = it.row.map((_, i) => (i === n - 1 ? W - w * (n - 1) : w));
      out.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: [new TableRow({ children: it.row.map((t, i) => cell([[{ t, b: !!it.bold }]], widths[i], { fill: (it.fills && it.fills[i]) || it.fill || FILL.lnavy, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER, size: 16 })) })] }));
    } else buf.push(it);
  }
  push();
  return out;
}

// ---------- 본문 ----------
const C = [];
// 표지
C.push(new Paragraph({ spacing: { before: 2400, after: 200 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "인도(노이다) 지점 총계정원장 검토 가이드", font: FONT, size: 44, bold: true, color: NAVY })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "본점 차세대 GBS 통합 표준 CoA 검토 담당자용 — 인도 회계·세무 기준 비교 자료", font: FONT, size: 24, color: GREY })] }));
C.push(table(["항목", "내용"], [
  ["대상", "본점 CoA·결산·HQ IT 담당자"],
  ["목적", "인도지점 총계정원장을 볼 때 알아야 할 인도 고유 기준(인도 GAAP, GST, TDS·TCS, RBI 규제)을 K-IFRS·한국 세제와 비교해 설명"],
  ["기준일", "2026-11-05 (인도 소득세법 2025 시행 후, RBI ECL 2027-04-01 시행 전)"],
  ["함께 보는 자료", "05_CoA검토_RBI보고속성.xlsx(근거·결산 증거), 06_통합표준CoA_0930_인도지점반영.xlsx(본점 양식 회신본)"],
  ["주의", "규정 수치(세율·한도·비율)는 기준일 현재 일반값. 시행 전 컴플라이언스·세무자문 확인 후 파라미터로 적재"],
], [2000, W - 2000]));
C.push(BR());
C.push(new Paragraph({ children: [new TextRun({ text: "목차", font: FONT, size: 28, bold: true, color: NAVY })], spacing: { after: 120 } }));
C.push(new TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-2" }));
C.push(P([{ t: "※ 목차가 비어 보이면 Word에서 목차를 클릭하고 F9(필드 업데이트)를 누르십시오.", c: GREY }], { size: 16 }));
C.push(BR());

// 1. 요약
C.push(H1("1. 한눈에 보는 차이"));
C.push(P("인도지점 원장은 본점과 같은 GBS·같은 CoA를 쓰지만, 기장 기준은 인도 법령입니다. 본점이 원장을 검토할 때 가장 자주 부딪히는 차이를 먼저 정리합니다."));
C.push(table(["주제", "본점(한국·K-IFRS)", "인도지점", "원장에서 보이는 모습"], [
  ["회계 기준", "K-IFRS", "인도 GAAP(AS) + RBI 지침. 은행은 Ind AS 미적용, 단계적 수렴", "ROU·리스부채 등 IFRS 전용 계정은 인도 원장에 잔액 없음(본점 조정)"],
  ["재무제표 서식", "K-IFRS 공시 서식", "BR Act Third Schedule(Sch.1~18)", "계정마다 Sch 줄이 정해져 있어야 함(예: 해외 Nostro는 Sch.7 '인도 외')"],
  ["부가세", "부가가치세 10%(금융용역 대부분 면세)", "GST 18%, 은행 수수료 과세·이자 비과세, 주 단위 세목(CGST·SGST·IGST)", "세목별·공급 성격별 계정, 매입세액 50%만 공제(나머지 비용)"],
  ["원천징수", "이자소득세 원천징수", "TDS(지급 시 원천징수) + TCS(받을 때 징수). 2026-04 신 소득세법", "TDS 신고서·납부 성격별 계정 21개, 연말 미경과 이자도 TDS"],
  ["거주성", "소득세법 거주자", "FEMA 거주성(계좌 종류)과 소득세법 거주성(세율)이 따로 판정", "NRE·NRO·SNRR·EEFC 등 FEMA 상품별 예금 계정"],
  ["대손충당금", "K-IFRS 1109 ECL", "~2027-03 IRACP(고정 비율), 2027-04 RBI ECL(Stage+규제 하한)", "정상자산 충당금이 지금은 부채(Sch.5)에 표시"],
  ["규제 비율", "BIS·LCR(한국)", "CRR·SLR(격주 Form A), LCR, PSL 40% 의무", "RBI 당좌·국채 잔액, PSL 미달 예치"],
  ["감가상각", "회사 정책", "Companies Act Schedule II 내용연수, 세무는 블록 정률(WDV)", "자산 분류가 Schedule II 기준으로 세분"],
  ["인건비", "퇴직급여(DB/DC)", "Gratuity(AS 15)·EPF·ESIC 법정 부담, 월 적립", "직원 관련 부채 6개 계정"],
  ["단위", "원·억", "lakh(10만)·crore(1천만)", "₹1 crore = ₹100 lakh = 10,000,000"],
], [1500, 2100, 3000, W - 6600], { zebra: true }));

// 2. 회계 기준
C.push(BR());
C.push(H1("2. 인도 회계 기준 — GAAP과 단계적 IFRS 수렴"));
C.push(H2("2.1 기준의 층"));
C.push(P("인도 은행의 장부는 네 가지 규범이 겹쳐서 정해집니다. 아래로 갈수록 은행에만 적용되는 특칙이며, 위의 기준보다 우선합니다."));
C.push(CAP("그림 1. 인도 은행 회계의 네 층 — 아래 층(RBI·세법)이 은행 특칙"));
flowTables([
  { row: ["① 재무제표 서식", "Banking Regulation Act 1949 s.29 — Third Schedule (Form A 대차대조표·Form B 손익, Sch.1~18)"], fills: [FILL.navy, FILL.lnavy], bold: false },
  { row: ["② 인식·측정", "Companies (Accounting Standards) Rules의 AS(인도 GAAP). 은행은 Ind AS 적용이 보류되어 AS를 사용"], fills: [FILL.lnavy, FILL.white] },
  { row: ["③ RBI 특칙", "자산건전성·충당금(IRACP → ECL 2027), 투자 분류(2023 지침), 수익 인식, 공시(Master Direction on Financial Statements)"], fills: [FILL.lnavy, FILL.white] },
  { row: ["④ 세법", "소득세법 2025(TDS·TCS, 법인세, 세무상각), CGST·IGST Act(GST)"], fills: [FILL.lnavy, FILL.white] },
]).forEach((t) => C.push(t));
C.push(H2("2.2 주요 AS와 K-IFRS 대응"));
C.push(table(["인도 AS", "내용", "대응 K-IFRS", "원장 영향"], [
  ["AS 5", "전기 항목(prior period items) 당기 손익에 별도 표시", "K-IFRS 1008(소급 재작성)", "전기오류수정 손익 계정 사용"],
  ["AS 9", "수익 인식", "K-IFRS 1115", "대출 수수료는 RBI 지침에 따라 2027-04부터 유효이자율"],
  ["AS 10", "유형자산", "K-IFRS 1016", "Schedule II 내용연수, 고정자산 GST 50% 원가 가산"],
  ["AS 11", "외화환산", "K-IFRS 1021", "외화 재평가 손익 계정 동일"],
  ["AS 15", "종업원급여(보험수리 손익 당기손익)", "K-IFRS 1019(재측정 OCI)", "Gratuity·연차수당 충당부채, OCI는 본점 조정"],
  ["AS 19", "리스(운용리스는 리스료 비용)", "K-IFRS 1116(사용권자산)", "인도 원장은 운용리스료만, ROU·리스부채는 본점 조정"],
  ["AS 22", "이연법인세(기간 차이)", "K-IFRS 1012(일시적 차이)", "이연법인세 계정 동일, 범위 차이는 본점 조정"],
  ["AS 26", "무형자산", "K-IFRS 1038", "소프트웨어 상각"],
  ["AS 29", "충당부채(할인하지 않음)", "K-IFRS 1037(현재가치)", "복구충당부채 할인 상각은 본점 조정"],
], [1000, 3300, 2100, W - 6400], { zebra: true }));
C.push(H2("2.3 IFRS로의 단계적 수렴"));
C.push(P("RBI는 은행의 Ind AS(인도판 IFRS) 전면 적용을 보류한 대신, 개별 지침을 IFRS 방향으로 바꾸고 있습니다. 원장은 이 순서대로 IFRS에 가까워집니다."));
C.push(CAP("그림 2. 인도 은행의 IFRS 수렴 일정"));
(() => {
  const yrs = ["2018~2019", "2024-04-01", "2025-09-22", "2026-04-01", "2027-04-01", "미정"];
  const ev = ["은행 Ind AS 적용 보류", "투자 분류 지침(HTM·AFS·FVTPL) — IFRS 9과 유사", "GST 2.0(세율 단순화)", "소득세법 2025 시행(TDS 392·393·394)", "ECL 지침·표준법 신용위험 자본 — IFRS 9 ECL·EIR 도입", "Ind AS 116·19·12 등 전면 적용"];
  const w = Math.floor(W / 6), ws = yrs.map((_, i) => (i === 5 ? W - w * 5 : w));
  const mk = (arr, fill, bold, keep) => new TableRow({ cantSplit: true, children: arr.map((t, i) => new TableCell({
    children: [new Paragraph({ keepNext: keep, alignment: AlignmentType.CENTER, children: runs(t, { size: 16, bold }) })],
    width: { size: ws[i], type: WidthType.DXA }, borders, shading: { fill, type: ShadingType.CLEAR, color: "auto" }, margins: { top: 60, bottom: 60, left: 80, right: 80 } })) });
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws, rows: [mk(yrs, FILL.lorange, true, true), mk(ev, FILL.lnavy, false, false)] })];
})().forEach((t) => C.push(t));
C.push(H2("2.4 원장 운영 원칙 — 인도 원장은 인도 GAAP, IFRS는 본점 조정"));
C.push(CAP("그림 3. 인도 원장과 본점 IFRS 조정의 흐름"));
flowTables([
  { box: "거래 (예금·대출·송금·비용)", fill: FILL.lgrey },
  { arrow: true },
  { box: "GBS 인도 원장 — 인도 GAAP (AS + RBI 지침)", fill: FILL.lnavy, bold: true },
  { arrow: true },
  { row: ["법정 재무제표 (Third Schedule)", "RBI 보고 (Form A·CIMS·LCR·PSL)", "세무 (TDS·GST·법인세)", "본점 IFRS 조정 (리스·IAS 19·IAS 12) → 본점 K-IFRS 보고"], fills: [FILL.lgreen, FILL.lgreen, FILL.lgreen, FILL.lorange] },
]).forEach((t) => C.push(t));
C.push(P("0930 CoA의 IFRS 전용 계정은 지우지 않고 남겨 두되, 인도 원장에서는 기장을 막습니다(06 파일 '인도 원장 기장' 열의 '본점 IFRS 조정 전용'). Ind AS가 적용되면 차단만 해제하면 됩니다."));
C.push(table(["코드(접두)", "계정"], D.ifrs_only, [2400, W - 2400], { zebra: true }));
C.push(H2("2.5 인도 GAAP ↔ IFRS 차이와 처리 위치"));
C.push(table(["항목", "인도 원장(인도 GAAP)", "본점 IFRS 조정", "계정", "시점"], D.gaap_diff, [1200, 2600, 2500, 2000, W - 8300], { zebra: true }));
C.push(H2("2.6 재무제표 서식(Third Schedule)과 CoA"));
C.push(P("인도 재무제표의 각 줄은 계정에서 바로 나와야 결산 수작업이 없어집니다. FY2025-26 결산에서는 해외 Nostro가 'Sch.7 인도 내'에, 콜론이 'Sch.6 RBI'에 잘못 배정되어 있었습니다."));
C.push(table(["Schedule", "표시 줄", "CoA 대응", "판정", "비고"], D.schedules, [1500, 2600, 1900, 1200, W - 7200], { zebra: true }));

// 3. GST
C.push(BR());
C.push(H1("3. GST — 구조와 은행 거래"));
C.push(H2("3.1 구조"));
C.push(P("GST는 주(State) 단위 부가가치세입니다. 같은 주 안의 거래는 중앙(CGST)과 주(SGST)가 절반씩, 주 사이·해외 관련 거래는 중앙이 통합(IGST)해 걷습니다. 노이다 지점은 우타르프라데시(UP, 주 코드 09)에 등록(GSTIN)되어 있습니다."));
C.push(table(["구분", "적용", "은행 세율", "납부처", "계정(11자리)"], [
  ["CGST", "주 내 거래(공급지 = UP)", "9%", "중앙정부", "20510500100 (L6: 수수료 101 / 환전 102)"],
  ["SGST", "주 내 거래", "9%", "UP 주정부", "20510500200 (L6: 201 / 202)"],
  ["IGST", "주 간·해외 관련(본점 cross-charge 포함)", "18%", "중앙 → 도착 주", "20510500300 (L6: 301 / 302)"],
  ["RCM", "수입 용역·미등록 임대인 등 — 구매자(은행)가 납부", "18%", "-", "20510500400~600"],
  ["ITC", "매입세액 공제. 은행은 50%만 공제 선택(s.17(4)), 나머지는 비용", "-", "-", "10420100100~700 / 불공제 50940010300"],
  ["비과세", "이자·할인료", "0%", "-", "-"],
], [1000, 3400, 1000, 1300, W - 6700], { zebra: true }));
C.push(H2("3.2 세목 결정 순서도"));
C.push(CAP("그림 4. 은행 수수료의 GST 세목 결정"));
flowTables([
  { box: "은행이 수수료·용역을 공급 (Charge code)", fill: FILL.lgrey },
  { arrow: true },
  { dec: "이자·할인료인가?", yes: "GST 없음 (비과세)", no: "과세 대상" },
  { arrow: true },
  { dec: "환전이 포함되는가?", yes: "과세가액 = Rule 32(2)(b) 산식(아래 표) → 환전 L6", no: "과세가액 = 수수료 금액 → 수수료 L6" },
  { arrow: true },
  { dec: "공급지(POS)가 UP 주인가?", yes: "CGST 9% + SGST 9%", no: "IGST 18% (타 주·본점·해외 수령인 중 과세분)" },
  { arrow: true },
  { box: "세금계산서 발행(GSTIN·SAC 9971) → GSTR-1 보고 → 3B 납부", fill: FILL.lgreen },
], 4800).forEach((t) => C.push(t));
C.push(H2("3.3 환전 과세가액 — Rule 32(2)(b)"));
C.push(table(["환전 금액(INR)", "과세가액", "예: USD 100,000 × 88.42 = ₹88,42,000"], [
  ["₹1 lakh 이하", "1% (최저 ₹250)", ""],
  ["₹1 lakh 초과 ~ ₹10 lakh", "₹1,000 + 초과분의 0.5%", ""],
  ["₹10 lakh 초과", "₹5,500 + 초과분의 0.1% (최고 ₹60,000)", "₹5,500 + 0.1% × ₹78,42,000 = ₹13,342 → GST 18% = ₹2,401.56"],
], [2600, 3000, W - 5600]));
C.push(H2("3.4 은행 거래별 GST"));
C.push(table(["거래", "GST", "계정"], [
  ["송금·LC·보증 수수료", "18% 과세(주 내 CGST+SGST / 그 외 IGST)", "수익 405xxxxxxxx, 매출세액 20510500xxx"],
  ["환전(송금에 포함)", "Rule 32(2)(b) 가액에 18%", "20510500102·202·302"],
  ["중계은행 수수료 고객 회수", "18% 과세(실제 차감액과 달라 pure agent 아님)", "40570110000"],
  ["본점이 지점에 청구하는 수수료(HQ charge)", "수입 용역 → 지점이 RCM IGST 납부", "20510500400 / ITC 10420100400"],
  ["SWIFT 등 해외 서비스 비용", "수입 용역 RCM", "50510600000 + RCM"],
  ["사무실 임차료", "임대인이 GST 부과 → ITC 50%", "50940021405 + ITC / 불공제 50940010300"],
  ["고정자산 구입", "ITC 50% 공제, 50%는 취득원가 가산", "10520xxxxxx"],
], [3200, 3400, W - 6600], { zebra: true }));
C.push(H2("3.5 월간 흐름과 신고"));
C.push(CAP("그림 5. GST 월간 처리 흐름 (신고 기한은 일반 기준)"));
flowTables([
  { row: ["매출(수수료·환전) 세금계산서", "매입 세금계산서 → GSTR-2B와 대사", "ITC 50% 공제 / 50% 불공제 비용", "GSTR-1 (익월 11일) 매출 명세", "GSTR-3B (익월 20일) 납부 — 전자현금원장", "GSTR-9·9C 연간"], fill: FILL.lnavy },
]).forEach((t) => C.push(t));
C.push(note("2B에 안 뜬 매입세액은 임시 ITC(10420100700)에 두었다가 대사 후 세목 계정으로 옮깁니다. 공급자에게 180일 넘게 대금을 안 주면 ITC를 되돌려야 합니다(Rule 37)."));
C.push(H2("3.6 분개 예 — 주 내 법인 고객 당발송금 USD 100,000"));
C.push(table(["차/대", "계정", "금액(INR)", "적요"], [
  ["차", "20150100100 법인 당좌", "88,57,452.50", "원금 + 수수료 + GST"],
  ["대", "11200000000 포지션", "88,42,000.00", "환전"],
  ["대", "40570010000 송금수수료", "11,052.50", ""],
  ["대", "40570030000 SWIFT·전신료", "1,000.00", ""],
  ["대", "20510500101 / 201 CGST·SGST(수수료)", "1,084.73 / 1,084.72", "(11,052.50+1,000) × 9%"],
  ["대", "20510500102 / 202 CGST·SGST(환전)", "1,200.78 / 1,200.78", "₹13,342 × 9%"],
], [700, 4000, 1900, W - 6600]));
C.push(P([{ t: "※ 금액은 설계서 예시(OUT-01)와 같은 가정. 실제 원장은 세목 L5 아래 수수료·환전 L6로 나뉩니다.", c: GREY }], { size: 16 }));

// 4. TDS / TCS
C.push(BR());
C.push(H1("4. TDS·TCS — 원천징수와 거주성"));
C.push(H2("4.1 개념"));
C.push(table(["구분", "의미", "은행의 역할", "예"], [
  ["TDS", "지급자가 지급(또는 계상) 시 세금을 떼어 납부", "예금이자 지급자, 비용 지급자", "정기예금 이자, 임차료, 용역비, 급여, 비거주자 지급"],
  ["TCS", "받는 쪽이 대금과 함께 세금을 더 받아 납부", "LRS 송금 은행", "개인 해외송금(교육·여행·투자)"],
  ["TDS 미수", "남이 은행 수입에서 떼어간 세금 — 법인세에서 공제", "피징수자", "임대수익, 차량 구입 TCS"],
], [1300, 3200, 2200, W - 6700], { zebra: true }));
C.push(H2("4.2 소득세법 2025 (2026-04-01 시행) — 조문 재편"));
C.push(table(["소득", "구 조문(1961법)", "신 조문(2025법)", "신고서"], [
  ["급여", "192", "392", "24Q"], ["예금·기타 이자", "193·194A", "393", "26Q (→ Form 140)"], ["배당", "194", "393", "26Q"],
  ["도급·수수료·임차료·전문용역", "194C·194H·194I·194J", "393", "26Q"], ["현금인출", "194N", "393(3)", "26Q"], ["물품구매·부동산", "194Q·194-IA", "393", "26Q·26QB 계열"],
  ["비거주자 지급", "195·194LC·194E", "393", "27Q"], ["TCS(LRS·여행·차량)", "206C", "394", "27EQ"], ["원천징수 면제 신고", "15G·15H", "Form 121", "-"],
], [3200, 2200, 1800, W - 7200], { zebra: true }));
C.push(note("계정 이름에 조문 번호를 넣지 않고 '소득 성격'으로 두었습니다. 조문이 바뀌어도 계정은 그대로이고, 조문·세율·한도는 파라미터 표(TAX_GL_PARAM)에서 바꿉니다."));
C.push(H2("4.3 거주성 — 두 가지 판정"));
C.push(P("인도에서는 같은 고객에 대해 거주성을 두 번 판정합니다. FEMA 거주성은 어떤 예금 상품(계정)을 쓸 수 있는지를, 소득세법 거주성은 어떤 세율로 원천징수하는지를 정합니다."));
C.push(table(["판정", "기준", "결정하는 것", "원장"], [
  ["FEMA 거주성", "직전 회계연도 인도 체류 182일 초과 + 취업·사업 등 체류 의사", "계좌 종류: 거주자 / NRE·NRO·FCNR(B)·SNRR", "예금 L5 (예: NRO 20150300300)"],
  ["소득세법 거주성", "해당 연도 182일 이상(또는 60일 + 직전 4년 365일 등)", "원천징수 조문·세율: 거주자 393(구 194A) / 비거주자 393(구 195)", "TDS L5·L6 (26Q 계열 / 27Q 계열)"],
], [1700, 3300, 2600, W - 7600], { zebra: true }));
C.push(CAP("그림 6. 예금이자 원천징수 판정 순서도"));
flowTables([
  { box: "고객·계좌 (CKYC·PAN·여권·체류)", fill: FILL.lgrey },
  { arrow: true },
  { dec: "FEMA상 인도 거주자인가?", yes: "거주자 예금(SB·CA·TD·RD) 사용", no: "NRE(인도계)·NRO(모든 비거주자)·FCNR(B)" },
  { arrow: true },
  { dec: "이자가 비과세 상품인가? (NRE·FCNR(B)·은행 저축예금)", yes: "TDS 없음", no: "아래 판정" },
  { arrow: true },
  { dec: "소득세법상 거주자인가?", yes: "거주자: 10%, 연 ₹50,000(고령 ₹1 lakh) 초과분부터, 15G/15H 면제", no: "비거주자: 30% + 4% cess(+surcharge), 한도 없음. TRC·Form 10F 있으면 조세조약 세율" },
  { arrow: true },
  { dec: "PAN이 있는가?", yes: "위 세율", no: "TDS 20% 이상(max(세율, 20%))" },
  { arrow: true },
  { box: "원천징수 → 익월 7일 납부(3월분 4/30) → 분기 신고 → 증명서(16A·16·27D)", fill: FILL.lgreen },
], 4600).forEach((t) => C.push(t));
C.push(note([{ t: "한국인 주재원", b: true }, ": 고용 목적 체류 중에는 FEMA·소득세법 모두 거주자 → 거주자 예금, 10% TDS. 귀임하면 NRO로 재지정(계정 이동)되고 비거주자 세율로 바뀝니다. 국적이 아니라 거주성이 계정을 바꿉니다."], FILL.lyellow, "예"));
C.push(H2("4.4 세율과 계산 예"));
C.push(table(["사례", "계산", "TDS", "계정"], [
  ["거주자 정기예금 이자 연 ₹60,000 (60세 미만, PAN 있음)", "₹50,000 초과 → 전액 10%", "₹6,000", "20510400101"],
  ["거주자 고령자 정기예금 이자 연 ₹90,000", "고령 한도 ₹1 lakh 이하", "0", "-"],
  ["거주자 15G 제출", "면제 신고 유효", "0", "-"],
  ["거주자 PAN 없음 ₹60,000", "max(10%, 20%) = 20%", "₹12,000", "20510400101"],
  ["NRO 정기예금 이자 ₹1,00,000", "30% + 4% cess = 31.2% (surcharge 별도)", "₹31,200", "20510400301"],
  ["NRO, 한국 거주자 TRC + Form 10F", "조세조약 이자 세율 적용(한·인 조약 10%, 확인 필요)", "₹10,000", "20510400301"],
  ["NRE·FCNR(B) 이자", "비과세", "0", "-"],
  ["은행 저축예금 이자", "은행 저축예금 이자는 원천징수 대상 아님", "0", "-"],
], [3300, 3000, 1100, W - 7400], { zebra: true }));
C.push(H2("4.5 TCS — LRS 송금"));
C.push(CAP("그림 7. LRS 송금 TCS 판정 (FY2026-27 세율, 확인 필요)"));
flowTables([
  { box: "거주 개인의 해외송금(LRS, 연 USD 250,000 한도)", fill: FILL.lgrey },
  { arrow: true },
  { dec: "교육비를 금융기관 대출로 송금?", yes: "TCS 없음", no: "아래" },
  { arrow: true },
  { dec: "해외여행 패키지인가?", yes: "금액과 관계없이 2% → 20510400203", no: "아래" },
  { arrow: true },
  { dec: "PAN 기준 연 누계 ₹10 lakh 초과분인가?", yes: "교육·의료 2% (20510400201) / 기타 20% (20510400202)", no: "TCS 없음" },
  { arrow: true },
  { box: "송금액 + TCS 출금 → 익월 7일 납부 → 27EQ 신고 → Form 27D 증명서", fill: FILL.lgreen },
], 4600).forEach((t) => C.push(t));
C.push(P("예: 같은 고객이 같은 해에 투자 목적으로 ₹15 lakh 송금 → ₹10 lakh 초과분 ₹5 lakh × 20% = TCS ₹1 lakh. 새 TCS 유형이 법령으로 생기면 'TCS 기타(20510400299)'에 먼저 기장합니다."));
C.push(H2("4.6 연말(3/31) 미경과 예금이자"));
C.push(P("만기 전이라도 미지급이자를 계상하면 그 시점이 원천징수 시점입니다. 계정 추가 없이 순액법으로 처리합니다."));
C.push(table(["ID", "시점", "차변", "대변", "비고"], D.ye_tds, [700, 1800, 2400, 2400, W - 7300], { zebra: true }));
C.push(H2("4.7 원장 계정 — 신고서(L5) × 납부 성격(L6)"));
C.push(table(["L6 코드", "계정", "한글", "현행 계정·비고"], D.tds_l6, [1500, 3300, 1800, W - 6600], { zebra: true }));
C.push(P(""));
C.push(H3("세금 계정 ↔ 파라미터 (신고서·조문·세율 관리 표)"));
C.push(table(["L5/L6", "세목", "신고서", "구 조문", "신 조문", "파라미터로 관리"], D.tax_param, [1600, 900, 1300, 1900, 900, W - 6600], { zebra: true }));

// 5. 기타
C.push(BR());
C.push(H1("5. 그 밖에 원장에 영향을 주는 인도 기준"));
C.push(table(["주제", "인도 기준", "원장 처리"], [
  ["PSL(우선부문 대출)", "조정 순은행신용(ANBC)의 40%를 농업·MSME·수출 등에 공급. 미달 시 NABARD 등 기금 예치", "계정이 아니라 대출 속성으로 분류. 미달 예치 10490920000, PSLC 수수료만 계정"],
  ["CRR·SLR (Form A)", "순요구불·정기부채(NDTL)의 일정 비율을 RBI 현금·국채로 보유, 격주 보고", "RBI 당좌 10104010100, 은행분/기타분 예금 계정 구분"],
  ["s.11(2)(b) 예치", "외국은행이 RBI에 예치(현금 또는 증권)", "10104010200, Sch.1 주기, SLR·HQLA 제외"],
  ["예금 상품(FEMA)", "NRE·NRO·FCNR(B)·SNRR·EEFC·RFC·SRVA 등 상품별 송금·과세 규칙", "상품별 예금 L5"],
  ["DICGC·DEAF", "예금자당 ₹5 lakh 보호(보험료 반기), 10년 미청구 예금은 RBI 기금 이전", "보험료 51010020000, 휴면도 예금 유지, DEAF 청구 미수 10490930000"],
  ["LCR", "소매 안정 5%·불안정 10%(디지털 +2.5%p), 비금융 40%, 금융 100%", "계정이 아니라 고객·계좌 분류 + 파라미터"],
  ["감가상각", "Schedule II 내용연수(컴퓨터 3년·서버 6년·가구 10년), 세무 블록 정률", "자산 L5 세분, FA 코드표에 내용연수 연결"],
  ["인건비", "Gratuity(5년 이상 근속 퇴직금), EPF 12%+12%, ESIC, 법정 상여", "비용 L5 + 월 적립 부채 20510300100~600"],
  ["CSR", "순자산 ₹500 crore 이상 등이면 3년 평균 순이익의 2% 지출", "50940029100, 미집행 부채 20440910000 (노이다는 FY26-27부터 해당으로 판단)"],
  ["Penal charges", "연체 부과금은 이자가 아닌 수수료(복리 불가)", "40580030000, 이자 계정 기장 금지"],
], [1800, 4200, W - 6000], { zebra: true }));

// 6. 체크리스트
C.push(H1("6. 총계정원장 검토 체크리스트"));
[
  "IFRS 전용 계정(ROU·리스부채 등)에 인도 원장 잔액이 없는가 — 있으면 기장 오류(본점 조정에서 처리)",
  "해외 Nostro가 Sch.7 '인도 외', 콜론이 Sch.7 Money at call로 배정되는가",
  "GST 매출세액이 세목(CGST·SGST·IGST)과 공급 성격(수수료·환전)으로 나뉘고, 3B 납부액과 계정 잔액이 맞는가",
  "GST 불공제(50%)가 50940010300에, 고정자산 GST가 원가에 들어갔는가",
  "TDS 계정 잔액 = 해당 월 챌린 납부액(익월 7일, 3월분 4/30)인가",
  "연말 미경과 이자에 TDS가 계상되었는가(순액법), 저축예금 이자에는 TDS가 없는가",
  "LRS 송금의 TCS가 세율별 L6로 나뉘었는가",
  "비거주자 예금(NRO 등) 이자에 비거주자 세율 또는 조세조약 세율(증빙 있음)이 적용되었는가",
  "감사보수·법률비용이 Service Commission이 아닌 전용 계정에 있는가",
  "Gratuity·연차·EPF·상여가 매월 적립되는가",
  "감가상각이 FAR(Schedule II) 기준과 일치하는가 — 결산 재계상이 없어야 함",
  "정상자산 충당금: 2027-03까지 IRACP, 2027-04부터 Stage별 ECL 계정",
  "s.11(2)(b) RBI 예치가 별도 계정이고 SLR·HQLA에서 빠졌는가",
  "휴면 예금이 예금 계정에 남아 있는가(부채 이관 금지)",
  "금액 단위(lakh·crore) 환산 착오가 없는가",
].forEach((t) => C.push(N(t)));

// 7. 부록
C.push(BR());
C.push(H1("7. 부록 — 인도 용어"));
C.push(table(["용어", "분류", "정의", "근거"], D.glossary, [2200, 800, 4800, W - 7800], { zebra: true }));

// ---------- 문서 ----------
const doc = new Document({
  features: { updateFields: true },
  creator: "NH Bank Noida Branch", title: "인도 총계정원장 검토 가이드(본점용)",
  styles: {
    default: { document: { run: { font: FONT, size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 30, bold: true, color: NAVY }, paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 24, bold: true, color: NAVY }, paragraph: { spacing: { before: 280, after: 120 }, keepNext: true, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 21, bold: true, color: ORANGE }, paragraph: { spacing: { before: 200, after: 80 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 440, hanging: 340 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "NH농협은행 노이다 지점 — 인도 총계정원장 검토 가이드", font: FONT, size: 16, color: GREY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: GREY })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
