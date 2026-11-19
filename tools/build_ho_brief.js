// 본점 협의 설명자료(국문): node tools/build_ho_brief.js <ho.json> <out.docx>
// 데이터: python3 tools/ho_brief_data.py <ho.json>  (06 정정본·coa_model 집계)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType,
  HeadingLevel, LevelFormat, BorderStyle, Header, Footer, PageNumber, TableOfContents, PageBreak,
} = require("docx");

const D = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = process.argv[3];
const FONT = "Malgun Gothic";
const NAVY = "1F3864", ORANGE = "C55A11", GREY = "595959";
const W = 9638; // A4 본문 폭(DXA, 좌우 2cm)
const A = (c) => { const a = D.acct[c]; if (!a) throw new Error("계정 없음 " + c); return a.name; };
const fmt = (n) => Number(n).toLocaleString("ko-KR");

const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: o.size || 20, bold: o.bold, color: o.color, italics: o.it });
const P = (t, o = {}) => new Paragraph({ spacing: { after: o.after ?? 100, line: 300 }, alignment: o.align, children: Array.isArray(t) ? t : [run(t, o)] });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [run(t, { size: 30, bold: true, color: NAVY })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run(t, { size: 24, bold: true, color: NAVY })] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [run(t, { size: 21, bold: true })] });
const B = (t, lvl = 0) => new Paragraph({ numbering: { reference: "bul", level: lvl }, spacing: { after: 60, line: 280 }, children: Array.isArray(t) ? t : [run(t)] });
const N = (t) => new Paragraph({ numbering: { reference: "num", level: 0 }, spacing: { after: 60, line: 280 }, children: Array.isArray(t) ? t : [run(t)] });
const bold = (a, b) => [run(a, { bold: true }), run(b)];
// 핵심 메시지 상자 (문단 음영 + 왼쪽 테두리)
const KEY = (title, lines) => [
  new Paragraph({ shading: { type: ShadingType.CLEAR, fill: "FBE5D6" }, border: { left: { style: BorderStyle.SINGLE, size: 24, color: ORANGE, space: 6 } },
    spacing: { before: 120, after: 0 }, indent: { left: 120, right: 120 }, children: [run(title, { bold: true, color: ORANGE, size: 20 })] }),
  ...lines.map((l, i) => new Paragraph({ shading: { type: ShadingType.CLEAR, fill: "FBE5D6" }, border: { left: { style: BorderStyle.SINGLE, size: 24, color: ORANGE, space: 6 } },
    spacing: { after: i === lines.length - 1 ? 160 : 0, line: 290 }, indent: { left: 120, right: 120 }, children: [run("· " + l, { size: 19 })] })),
];
const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
function TBL(headers, rows, ratios, o = {}) {
  const sum = ratios.reduce((a, b) => a + b, 0);
  const widths = ratios.map((r) => Math.floor((W * r) / sum));
  widths[widths.length - 1] += W - widths.reduce((a, b) => a + b, 0);
  const size = o.size || 17;
  const mk = (txt, i, head, fill) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: head ? NAVY : fill || "FFFFFF" },
    borders: { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder },
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: String(txt ?? "").split("\n").map((line) => new Paragraph({ children: [run(line, { size, bold: head || (o.boldFirst && i === 0), color: head ? "FFFFFF" : undefined })] })),
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => mk(h, i, true)) }),
      ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => mk(c, i, false, o.zebra && ri % 2 ? "F2F2F2" : (o.fill && o.fill(r, i)))) }))],
  });
}
const SP = () => new Paragraph({ spacing: { after: 80 }, children: [] });
const C = [];

// ---------------- 표지 ----------------
C.push(new Paragraph({ spacing: { before: 2600, after: 200 }, children: [run("인도(노이다) 지점", { size: 36, bold: true, color: NAVY })] }));
C.push(new Paragraph({ spacing: { after: 200 }, children: [run("차세대 GBS 통합표준 CoA 인도 반영", { size: 44, bold: true, color: NAVY })] }));
C.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 4 } }, spacing: { after: 300 }, children: [run("본점 협의 설명자료", { size: 30, color: ORANGE, bold: true })] }));
C.push(P("인도 특화 기준 · 중요 기준 · 본점 요청 사항 · 현행 계정 이행 결과", { size: 22, color: GREY }));
C.push(new Paragraph({ spacing: { before: 2400 }, children: [run("2026년 11월 · 인도 노이다 지점", { size: 22 })] }));
C.push(P("첨부: 06_통합표준CoA_0930_인도반영_정정본.xlsx (계정 상세), 02_데이터사전_업무규칙.xlsx 17·19·20·21번 시트", { size: 18, color: GREY }));

// ---------------- 목차 ----------------
C.push(new Paragraph({ children: [new PageBreak()] }));
C.push(P("목차", { size: 28, bold: true, color: NAVY }));
C.push(new TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-2" }));
C.push(P("(Word에서 열면 '필드 업데이트'로 쪽 번호가 채워집니다)", { size: 16, color: GREY }));

// ---------------- 0. 한 장 요약 ----------------
const S = D.summary, grp = ["자산(B/S)", "부채(B/S)", "자본(B/S)", "수익(P/L)", "비용(P/L)", "각주"];
const tot = (k) => grp.reduce((a, g) => a + (S[g][k] || 0), 0);
C.push(H1("0. 한 장 요약 (회의용)"));
C.push(...KEY("이 자료로 본점과 합의하려는 것", [
  "인도는 회계기준(인도 GAAP)·재무제표 서식·외환/예금 규제·세금 체계가 본점과 달라, 같은 0930 CoA를 쓰되 인도에만 필요한 계정과 사용 규칙을 더했습니다.",
  "계정은 꼭 필요한 경우(재무제표 줄·대사 단위·FEMA 상품이 다를 때)만 늘리고, 나머지는 속성(태그)으로 처리해 계정 수를 최소화했습니다.",
  `현행 인도 계정 ${fmt(D.cur.leaf)}개는 빠짐없이 새 CoA로 옮겨지며, 2026-07-07 잔액이 계정별로 정확히 일치합니다(차이 0).`,
  "본점에는 신규 계정 정식 채번, 0930 원안의 오류 정정, IFRS 조정 프로세스 반영 등 10건의 조치를 요청드립니다(4장).",
]));
C.push(H2("숫자로 보는 결과"));
C.push(TBL(["구분", "신규", "변경", "용도지정", "비활성", "철회", "유지", "합계"],
  [...grp.map((g) => [g, ...["신규", "변경", "용도지정", "비활성", "철회", "유지"].map((k) => String(S[g][k] || 0)), String(Object.values(S[g]).reduce((a, b) => a + b, 0))]),
    ["합계", ...["신규", "변경", "용도지정", "비활성", "철회", "유지"].map((k) => String(tot(k))), String(D.total)]],
  [1.6, 0.8, 0.8, 0.9, 0.8, 0.8, 0.8, 0.8], { boldFirst: true, fill: (r) => (r[0] === "합계" ? "DEEAF6" : null) }));
C.push(P("신규: 인도에 필요해 새로 요청하는 계정(임시코드) · 변경: 인도 표시명·사용 방식 변경 · 용도지정: 인도 업무를 기존 계정에 지정 · 비활성: 인도 미사용(삭제 아님) · 철회: 앞선 검토안을 취소", { size: 16, color: GREY }));
C.push(TBL(["항목", "결과"], [
  ["본점 IFRS 조정 전용 계정", `${D.ifrs_only}개 — CoA에는 두되 인도 원장 기장 차단(본점 조정에서만 사용)`],
  ["현행 인도 계정(8자리) 이행", `최하위 ${fmt(D.cur.leaf)}개 = 이관 ${fmt(D.cur_mapped - 1)}개 + 이행 제외 ${D.cur_excluded}개 (모두 잔액 0인 폐지 대상). 06 파일에만 있던 신설 예정 코드 1개 별도`],
  ["잔액 대사", `잔액 보유 ${D.cur.nonzero}개 전부 이관, 대차대조표·손익계산서 대차 0, 계정별 차이 0건`],
  ["06 원본 대비 정정", `${D.log_n}건 (정정이력 시트) — 현행 계정 연결 ${D.log_by_col["현행 인도 GBS(8자리)"] || 0}, 0930 정정 권고 ${D.log_by_col["0930 정정 권고"] || 0}, 상태·기장 ${(D.log_by_col["인도 상태"] || 0) + (D.log_by_col["인도 원장 기장"] || 0)}, 잔액 ${D.log_by_col["현행 잔액 INR(2026-07-07, 차변+)"] || 0}, 행 신설 1`],
], [1.6, 5]));
C.push(H2("회의 진행 제안 (약 90분)"));
C.push(TBL(["순서", "안건", "자료", "목표", "시간"], [
  ["1", "왜 인도는 별도 반영이 필요한가", "1장", "배경 공유", "10분"],
  ["2", "인도 특화 기준 — 인도에만 있는 계정", "2장", "필요성 이해", "25분"],
  ["3", "중요 기준 — 본점과 합의할 원칙", "3장", "원칙 합의", "25분"],
  ["4", "본점 요청 사항 10건", "4장", "담당·일정 확정", "20분"],
  ["5", "이행 결과·추후 고려사항, 질의응답", "5~7장", "공유", "10분"],
], [0.6, 3.2, 0.8, 1.4, 0.7]));

// ---------------- 1. 배경 ----------------
C.push(H1("1. 왜 인도는 별도 반영이 필요한가"));
C.push(...KEY("핵심", [
  "본점 0930 CoA는 K-IFRS 기준입니다. 인도지점 장부는 인도 법령상 인도 회계기준(AS)과 RBI 지침으로 작성해 RBI·세무당국에 제출해야 합니다.",
  "따라서 '본점 CoA를 그대로 쓰되, 인도 법정 장부를 만들 수 있게' 계정과 사용 규칙을 보완한 것이 이번 작업입니다.",
]));
C.push(H2("1.1 본점과 다른 네 가지"));
C.push(TBL(["영역", "본점(0930)", "인도지점", "CoA에 미치는 영향"], [
  ["회계기준", "K-IFRS", "인도 GAAP(AS) + RBI 지침. 은행은 Ind AS(인도판 IFRS) 적용이 연기됨", "리스·ROU·보증금 할인 등 IFRS 전용 계정은 인도 원장에 기장하지 않음"],
  ["재무제표 서식", "본점 서식", "은행법(BR Act) s.29 Third Schedule(Form A/B, Schedule 1~18)", "계정마다 Schedule 줄을 지정해야 함. 줄이 다르면 계정을 나눔"],
  ["외환·예금 규제", "-", "FEMA·RBI: NRE·NRO·FCNR(B)·EEFC·SNRR·SRVA 등 거주성별 예금 상품, CRR·SLR", "상품별로 송금 가능 여부·세금·준비금이 달라 예금 계정을 분리"],
  ["세금", "부가세(VAT)", "GST(CGST·SGST·IGST·역발행), 원천징수 TDS·TCS(조문별 신고)", "세목·신고서·납부 단위로 계정을 나눠야 신고·대사가 가능"],
], [1, 1.1, 2.4, 2.4]));
C.push(H2("1.2 계정을 늘리는 기준 — 'GL은 최소, 나머지는 속성' (결정 G6)"));
C.push(P("계정 수를 늘리면 본점 연결·관리 부담이 커지므로, 다음 세 가지 경우에만 계정(GL)을 나누고 나머지는 거래·고객 속성(태그)으로 관리합니다."));
C.push(TBL(["계정으로 나누는 경우", "예", "속성으로 처리하는 경우", "예"], [
  ["① 인도 재무제표(Schedule) 줄이 다를 때", "가구·비품을 '기타 유형자산'과 분리", "규제 보고용 분류", "우선부문대출(PSL)·MSME 구분, 은행통계(BSR)"],
  ["② 대사(잔액 맞춤) 단위가 다를 때", "RBI 당좌 / RBI 기타계좌", "건전성·위험 구분", "ECL Stage 세부, 국가위험, LCR 분류"],
  ["③ FEMA 상품이 다를 때", "NRE / NRO / FCNR(B) / EEFC", "거래 성격", "금융보증·이행보증, 내국신용장, SNRR 용도"],
], [2, 1.8, 1.6, 2]));
C.push(P("그 결과 신규 175개 중 상당수는 인도 법정 장부에 반드시 필요한 예금·세금·준비금 계정이며, 나머지 규제 분류는 계정을 늘리지 않고 속성으로 해결했습니다.", { size: 19 }));

// ---------------- 2. 인도 특화 기준 ----------------
C.push(H1("2. 인도 특화 기준 — 인도에만 있는 계정"));
C.push(...KEY("읽는 법", [
  "각 항목은 '무엇인가 → 왜 계정이 필요한가 → 대표 계정 → 본점 영향' 순서입니다.",
  "신규 계정 코드는 본점 채번 전 임시코드이며, 본점 통합계정 매핑은 같은 상위 계정 기준의 제안값입니다.",
]));
const sec = (title, what, why, codes, impact) => {
  C.push(H2(title));
  C.push(P([run("무엇인가  ", { bold: true, color: ORANGE }), run(what)]));
  C.push(P([run("왜 계정이 필요한가  ", { bold: true, color: ORANGE }), run(why)]));
  C.push(TBL(["계정", "계정명", "상태"], codes.map(([c, note]) => [c, A(c) + (note ? ` — ${note}` : ""), D.acct[c].status || "유지"]), [1.3, 4.6, 0.8]));
  C.push(P([run("본점 영향  ", { bold: true, color: ORANGE }), run(impact)], { after: 160 }));
};
sec("2.1 예금 상품 (FEMA·RBI)",
  "인도는 예금주의 거주성과 자금 원천에 따라 상품이 법으로 정해져 있습니다. NRE(비거주자, 해외송금 자유·비과세), NRO(비거주자의 인도 내 소득, 송금 연 USD 1백만 한도·원천징수), FCNR(B)(비거주자 외화정기), EEFC(거주자 외화수입 보유), RFC(귀국 거주자 외화), SNRR(비거주자 사업·투자 결제용 루피), SRVA(해외은행의 루피 무역결제 계좌) 등입니다.",
  "상품마다 송금 가능 여부, 이자 과세(원천징수), 예금보험(DICGC ₹5 lakh), 지급준비금(CRR·SLR) 적용이 달라 잔액을 따로 관리·보고해야 합니다. 거주자 당좌·저축도 Third Schedule(Sch.3) 줄이 달라 분리했습니다.",
  [["20150300200"], ["20150300300", "외국인 비거주자도 포함(속성)"], ["20160300400"], ["20150100200"], ["20150100300", "용도는 속성"], ["20150200100"], ["20150400100", "FDI 지분양수도 등"], ["20190000000", "적금 L3 신설"], ["20150300600", "개인 당좌 분리"]],
  "본점 연결 시 이 계정들은 모두 기존 '대외계정' 예금으로 합산됩니다(매핑 제안). 상품 구분은 인도 법정 보고에만 쓰입니다.");
sec("2.2 중앙은행·지급결제",
  "외국은행 지점은 RBI에 당좌계좌(지급준비금 CRR·RTGS 결제)를 두고, 은행법 s.11(2)(b)에 따라 별도 예치금(현금 또는 담보증권)을 RBI에 두어야 합니다. 국내 결제는 NEFT·RTGS로 이루어집니다.",
  "s.11(2)(b) 예치금은 지급준비금·유동성(SLR·LCR) 계산에서 제외되므로 RBI 당좌와 섞이면 규제비율이 틀어집니다. 결제용 국내은행 당좌와 은행 정기예치도 대사 단위가 달라 분리했습니다.",
  [["10104010100"], ["10104010200", "s.11(2)(b) 예치 (담보 표시는 속성)"], ["10105010100"], ["10106010000", "은행 정기예치 ₹208.9 crore"]],
  "본점 매핑은 기존 '해외은행 예치' 계열을 그대로 사용합니다. 본점 보고 영향 없음.");
sec("2.3 세금 — GST·TDS·TCS",
  "인도 GST는 같은 주(州) 안이면 CGST+SGST, 다른 주면 IGST로 나뉘고, 수입 용역 등은 은행이 대신 내는 역발행(RCM)이 있습니다. 은행은 매입세액의 50%만 공제할 수도 있습니다. 원천징수(TDS)와 해외송금 시 징수(TCS)는 조문·신고서(26Q·27Q·27EQ·24Q)별로 신고·납부합니다.",
  "세목·신고서·납부 단위로 계정이 나뉘어야 월별 GST 신고(GSTR)와 분기 TDS 신고를 장부와 맞출 수 있습니다. 환전 수수료에는 환전가치 기준 별도 과세표준(Rule 32(2)(b))이 있어 수수료분·환전분을 나눴습니다.",
  [["20510500101", "수수료분"], ["20510500102", "환전분"], ["20510500400"], ["10420100700"], ["10490940000"], ["20510400101"], ["20510400201", "해외송금(LRS) 시 징수"], ["10810200000"], ["50940010300"]],
  "본점 매핑은 '기타 부채(세금 예수금)' 계열로 합산됩니다. 세목 구분은 인도 신고 전용입니다.");
sec("2.4 여신",
  "인도는 모든 은행에 우선부문대출(PSL, 대출의 40%) 의무가 있고 미달 시 정부기금에 저리 예치해야 합니다. 수출 선적전 금융(Packing Credit), 당좌대월(CC/OD), MSME 매출채권 할인(TReDS), 수출 이자보전 같은 인도 고유 상품이 있습니다.",
  "PSL 분류 자체는 속성으로 처리했지만, 미달 예치금과 그 이자는 실제 자금이므로 계정이 필요합니다. 인도 고유 대출 상품은 Schedule 9 줄과 이자 대사 단위가 달라 분리했습니다.",
  [["10490920000"], ["40130980000"], ["10330500000", "INR·외화(PCFC) L5"], ["10330304000"], ["10490910000"], ["40580030000", "Penal Charges(이자 아님, RBI 2023)"]],
  "본점 매핑은 기존 대출 계열. PSL·MSME 분류는 속성이라 본점 CoA 영향 없음.");
sec("2.5 자본·준비금",
  "인도 은행법은 순이익의 25%를 법정준비금으로 적립하게 하고, RBI는 투자변동준비금(IFR)과 자본비율용 유보 잉여금을 요구합니다. 본점이 배정한 영업기금은 갑기금(무이자)입니다.",
  "준비금은 본점 송금 가능 이익과 자본비율(CRAR) 계산을 좌우하므로 별도 계정이 필요합니다. 부채측 STATUTORY RESERVE(20510200000)는 비활성으로 정리했습니다.",
  [["31110300000"], ["31110600000"], ["31110700000"], ["31110800000", "2027-04 ECL 전환 시"], ["31140100000", "갑기금, 무이자"]],
  "본점 연결 시 자본 계열로 합산. 법정준비금 적립은 본점 이익 송금 가능액을 줄이므로 본점 자금부 공유 필요.");
sec("2.6 인사·비용",
  "인도 노동법상 퇴직금(Gratuity, 5년 이상 근속)·연차수당은 인도 회계기준(AS 15)으로 보험수리 평가하고, 연금(EPF)·의료보험(ESIC) 사용자 부담이 있습니다. 순자산 ₹500 crore 이상 법인은 사회공헌(CSR, 순이익 2%) 의무가 있습니다.",
  "법정 부채·비용이고 감사(LFAR)·세무 대상이라 계정이 필요합니다. 고정자산도 인도 회사법 Schedule II 내용연수(컴퓨터 3년·서버 6년 등)가 달라 분류를 나눴습니다.",
  [["20510300400"], ["50940029100"], ["20440910000"], ["10520100100", "Schedule II 3년"], ["10520400000", "Sch.10 별도 줄"]],
  "본점 매핑은 인건비·기타비용·유형자산 계열로 합산.");
sec("2.7 본지점·외환 송금",
  "본점과의 거래는 갑기금(자본·무이자), 을A(RA 대지급·송금 등 거래성), 을B(차입·예치 등 여신성)로 나눠 금리를 달리 적용합니다. 본점 명의 INR 계좌, 계좌 없는 수취인 앞 송금 보류, 무역금융 순지급도 인도 실무상 따로 관리합니다.",
  "본점 INR 계좌는 같은 법인이라 예금이 아닌 본지점 부채로 두어야 예금·준비금이 부풀지 않습니다. 송금 미지급은 SFMS(RTGS·NEFT) 대사 단위별로 나눴습니다.",
  [["11010000000", "을A"], ["11020000000", "을B"], ["21040000000"], ["20181100100"], ["20440800000"], ["40570110000", "고객 OUR 회수분"]],
  "을A는 본지점 정산배치로 본점과 순액 정산. 본점 측 상대 계정과 금리 기준을 함께 확인 필요.");

// ---------------- 3. 중요 기준 ----------------
C.push(H1("3. 중요 기준 — 본점과 합의할 원칙"));
C.push(...KEY("이 장에서 결정해 주셔야 할 것", [
  "아래 7개 원칙은 인도 장부와 본점 보고가 함께 맞으려면 양측이 같은 기준을 써야 하는 사항입니다.",
  "각 항목의 '본점 영향'과 '요청'을 중심으로 합의해 주시면 됩니다.",
]));
const imp = (title, issue, india, impact, ask) => {
  C.push(H2(title));
  C.push(TBL(["쟁점", "인도안", "본점 영향", "요청"], [[issue, india, impact, ask]], [1.6, 2.4, 1.8, 1.5]));
  C.push(SP());
};
imp("3.1 기준 계정코드와 기장 단위 (G1·G2)", "현행 8자리와 0930 11자리가 공존", "11자리 통합코드만 기준. 8자리는 이행 원천코드로만 쓰고 이행 후 폐기. 기장은 최하위(leaf) 계정에만", "본점 GBS 코드 열(8자리)에 의존한 보고가 있으면 11자리 전환 필요", "11자리 단일 기준 합의");
imp("3.2 회계기준 차이와 IFRS 전용 계정 (G9·F2)", "0930에는 IFRS 16 리스(사용권자산·리스부채) 등 IFRS 전용 계정이 있음", `인도 원장은 인도 GAAP(AS 19 운용리스)로 기장. IFRS 전용 ${D.ifrs_only}개는 CoA에 두되 인도 원장 기장 차단, 본점 K-IFRS 조정에서만 사용. 조정 근거(리스 원장·보험수리·ECL 모델값)는 GBS 속성으로 제공`, "본점 연결 시 IFRS 16·IAS 19·IAS 12 조정을 본점에서 해야 함", "본점 조정 프로세스에 IFRS 16·IAS 19·IAS 12 포함");
imp("3.3 ECL(기대신용손실) 전환 (L3·L6·L8)", "RBI ECL 지침(2026)이 2027-04-01 시행. 현행은 일반·특정 충당금", "2027-03-31까지 일반·특정 충당금(10380101000/102000) 사용 → 시행일에 Stage 1~3(10380103000~105000)으로 이관, 차이는 자본 31110800000. 은행 예치·콜론·만기보유채권도 ECL 대상", "본점 IFRS 9 충당금과 인도 ECL 하한(floor) 차이 조정 필요", "전환일·이관 전표 방식 합의, ECL 명세서 정정(L5)");
imp("3.4 투자 분류 (RBI 2023 투자 지침)", "인도 HTM·AFS·HFT vs 0930 AC·FVOCI·FVPL", "HTM→AC, AFS→FVOCI, HFT→FVPL로 대응. AFS 평가손익은 'AFS Reserve'(31130000000), 투자변동준비금(IFR) 별도. HFT 할인·프리미엄 상각 계정은 0930에 없어 FVPL 이자·공정가치손익으로 흡수", "분류 명칭이 달라도 본점 연결 대응은 1:1", "대응표 확인, HFT 처리 정책 확정(추후)");
imp("3.5 난외 계정 (G5)", "현행 인도는 난외를 6xxx/7xxx 복식으로 기장", "0930 단식 체계 유지. 7xxx 대응계정은 시스템 생성으로 폐지. 인도 비망계정(미인식 이자 60310000000, 기술적 상각 60320000000)만 L2 60300000000 아래 신설", "없음 (0930 체계 그대로)", "비망계정 L2 채번");
imp("3.6 GL 대신 속성으로 처리하는 항목 (G6·L7·P5)", "1차 검토에서 금융/이행보증·내국신용장·SNRR 용도를 계정으로 나눴던 안", "모두 철회하고 속성(bg_nature·lc_scope·snrr_purpose)으로 처리. 규제 환산율(CCF) 계산은 속성 기준", "본점 CoA 변경 없음", "확인만");
imp("3.7 0930 원안 구조·코드 오류 (G3·G4)", "0930 자체의 코드 구조 오류와 GBS코드 열 오대응", "20501000000을 산식상 맞는 20510000000으로 재채번(하위 TDS·GST 계정의 상위). GBS코드 열 57건 + 파생 6건 정정", "0930 원안 수정 필요", "본점 CoA 담당 정정");

// ---------------- 4. 본점 요청 ----------------
C.push(H1("4. 본점 요청 사항"));
C.push(...KEY("요청 요약", [
  "아래 10건은 본점 CoA·보고 프로세스에서 조치가 필요한 사항입니다. 인도지점이 단독으로 처리할 수 없습니다.",
  "우선순위 '높음'은 차세대 GBS 이행 일정에 직접 영향을 줍니다.",
]));
const DEC = Object.fromEntries(D.decisions.map((d) => [d[0], d]));
C.push(TBL(["#", "요청", "내용", "근거", "우선순위"], [
  ["1", "신규 계정 정식 채번", "인도 신규 175개는 임시코드. 본점 채번 후 일괄 교체하며, 본점 통합계정 매핑(안) 확인 요청", "06 신규", "높음"],
  ["2", "OTHER LIABILITIES L3 재채번", "20501000000 → 20510000000 (하위 TDS·GST·직원부채 계정의 상위)", "G3", "높음"],
  ["3", "0930 GBS코드 열 정정", "57건 + 파생 6건(통화스왑·금리스왑, 트레이딩·헤지 뒤바뀜)", "G4", "높음"],
  ["4", "본점 IFRS 조정 프로세스", DEC.G9[7] + ". " + DEC.F2[7], "G9·F2", "높음"],
  ["5", "ECL 관련 정정", DEC.L5[7] + " (EIR 조정계정 신설 요청 철회). 일반·특정 충당금 2027-03-31 기한부 사용 확인", "L3·L5", "중간"],
  ["6", "적금(RD) 계정 신설", DEC.P3[7], "P3", "중간"],
  ["7", "계정 명칭 변경", "10410230000 'TERM PLACEMENTS / CD'(M2), 전신료 'SWIFT / CABLE CHARGES'(C2), 40580030000 'LATE PAYMENT / PENAL CHARGES'(C3)", "M2·C2·C3", "낮음"],
  ["8", "세무 산출물 정정", "TDS 산출물(저축예금 이자 TDS 대상 아님) 정정, GST FRS 코드 정정", "T6·T7", "중간"],
  ["9", "AOB 정의 기재", DEC.G7[7] + " (인도지점 비활성)", "G7", "낮음"],
  ["10", "파생 계정 보완 (취급 개시 전)", "파생 손익 'OTHERS' L4 신설, Day-1 이연손익 계정 신설", "추후", "낮음"],
], [0.35, 1.5, 3.6, 0.8, 0.75], { zebra: true }));

// ---------------- 5. 이행 결과 ----------------
C.push(H1("5. 현행 계정 → 차세대 계정 이행 결과"));
C.push(...KEY("결론", [
  `현행 인도 계정 ${fmt(D.cur.leaf)}개 중 잔액이 있는 ${D.cur.nonzero}개는 모두 새 계정으로 옮겨지고, 2026-07-07 잔액이 계정별로 일치합니다.`,
  "잔액이 0인 계정 중 인도에 없는 상품(MMDA·CP 차입 등)과 비활성 대상만 이행에서 제외합니다.",
]));
C.push(TBL(["항목", "결과", "설명"], [
  ["현행 최하위 계정", fmt(D.cur.leaf), "2026-07-07 대차대조표·손익계산서 999줄 중 합계줄 제외"],
  ["이관", fmt(D.cur_mapped - 1), "새 11자리 최하위 계정으로 연결 (06 파일에만 있는 신설 예정 코드 10990003 제외)"],
  ["이행 제외", String(D.cur_excluded), "모두 잔액 0: 인도 미취급 상품, 비활성 계정 대상 등 (정정본 '현행이행제외' 시트)"],
  ["1:N 분할", String(D.splits), "현행 1개 → 새 계정 여러 개. 고객유형·고정자산 분류·거주성·TCS 유형·경비코드로 분할 ('이행분할기준' 시트)"],
  ["잔액 검증", "차이 0", "대차대조표 대차 0, 손익 대차 0, 계정별 06 잔액 = 현행 잔액"],
], [1.3, 0.9, 4.4]));
C.push(H2("5.1 06 원본 대비 정정한 내용"));
C.push(TBL(["정정", "내용", "근거"], [
  ["상위 계정 → 최하위 계정", "현행 코드가 하위 계정이 있는 상위 계정에 연결된 10건을 최하위 계정으로 이동 (예: SOFTWARE 잔액 ₹3.69 crore)", "G2"],
  ["오류 연결 삭제", "비용 코드(54010091)가 부채 계정에도 연결된 건, 처분손실(53090014) 중복 연결", "G4"],
  ["비활성 대상 정리", "잔액 0인 현행 8개가 비활성 계정에 연결 → 이행 제외로 명시", "G1·G7"],
  ["미연결 현행 계정 보완", "06 이행표에 없던 132개(전부 잔액 0) 중 115개 연결, 17개 제외", "현행 BS/PL 대사"],
  ["충당금 상태", "일반·특정 충당금 '비활성' → '2027-03-31까지 사용' (현행 잔액 ₹-1.59 crore 이관처 확보)", "L3"],
  ["L3 계정 신설", "20510000000 OTHER LIABILITIES", "G3"],
], [1.6, 4.4, 0.8]));

// ---------------- 6. 추후 고려사항 ----------------
C.push(H1("6. 추후 고려사항"));
C.push(P("지금은 잔액이 없어 추천안으로 확정했지만, 아래 시점이 오면 다시 판단할 사항입니다(32건, 정정본 '추후고려사항' 시트)."));
C.push(TBL(["건수", "다시 볼 내용", "재검토 시점", "담당"], D.future.map((f) => [String(f[0]), f[1], f[2], f[3]]), [0.5, 4, 1.5, 1.2]));

// ---------------- 7. 예상 질문 ----------------
C.push(H1("7. 예상 질문과 답변"));
const qa = [
  ["왜 계정이 175개나 늘어나나요?", "인도 법정 장부(Third Schedule)와 세무 신고에 반드시 필요한 예금 상품·세목·준비금·인사 계정입니다. 규제 분류(PSL·MSME·보증 성격 등)는 계정 대신 속성으로 처리해 늘어나는 수를 최소화했습니다. 본점 연결 시에는 모두 기존 통합계정으로 합산됩니다."],
  ["IFRS 전용 계정을 인도에서 안 쓰면 본점 연결은 어떻게 하나요?", "인도 원장은 인도 GAAP으로 작성하고, 본점이 연결할 때 IFRS 16(리스)·IAS 19(퇴직급여)·IAS 12(이연법인세) 차이를 조정합니다. 인도지점은 조정에 필요한 리스 원장·보험수리 결과 등을 GBS 속성으로 제공합니다."],
  ["속성으로 처리하면 본점 보고에 문제가 없나요?", "속성은 같은 계정 안의 세부 구분이라 본점 계정 합계에는 영향이 없습니다. 인도 규제 보고(RBI)에만 속성을 사용합니다."],
  ["현행 8자리 코드는 언제 없어지나요?", "차세대 GBS 이행 때 이행 원천코드로만 쓰고, 이행 검증이 끝나면 폐기합니다(결정 G1)."],
  ["ECL 전환 후 기존 충당금은 어떻게 되나요?", "2027-03-31까지는 현행 일반·특정 충당금을 쓰고, 2027-04-01 시행일에 Stage 1~3 계정으로 옮깁니다. 차이는 ECL 전환 조정(자본, 31110800000)으로 처리합니다."],
  ["신규 계정 코드가 본점 채번과 다르면?", "임시코드는 본점 채번 즉시 일괄 교체하도록 설계돼 있습니다(계정 마스터·분개 규칙·잔액 이관 동시 교체)."],
  ["갑기금 이자 계정을 왜 비활성으로 하나요?", "갑기금은 무이자로 확정했기 때문입니다. 잘못 기장되지 않도록 인도 원장에서 차단합니다."],
];
qa.forEach(([q, a]) => { C.push(P([run("Q. " + q, { bold: true, color: NAVY })], { after: 40 })); C.push(P("A. " + a, { after: 160 })); });

// ---------------- 부록 ----------------
C.push(H1("부록 A. 인도 용어 풀이"));
C.push(P("본점 참고용으로 06 파일 '인도용어정의' 시트를 옮겼습니다.", { size: 18, color: GREY }));
C.push(TBL(["분류", "용어", "뜻", "근거"], D.terms.map((t) => [t[0], t[1], t[2], t[3] || ""]), [0.7, 1.6, 4.2, 1.4], { size: 15, zebra: true }));
C.push(H1("부록 B. 결정사항 전체 목록 (06 파일)"));
C.push(TBL(["ID", "영역", "결정 사항", "확정안", "본점 조치"], D.decisions.map((d) => [d[0], d[1], d[2], String(d[4] || "").slice(0, 260), d[7] || "-"]), [0.45, 0.8, 1.3, 4.2, 1.3], { size: 14, zebra: true }));
C.push(H1("부록 C. 첨부 자료 안내"));
C.push(TBL(["파일", "시트", "내용"], [
  ["06_통합표준CoA_0930_인도반영_정정본.xlsx", "통합표준CoA_계정상태표", "1,361개 계정 전체. A~J 본점 원안(불변), K~S 인도 검토, T·U 정정 내용·근거. 노랑=정정 셀, 주황=신설 행"],
  ["", "정정이력 / 현행이행제외 / 이행분할기준 / 추후고려사항", "06 원본 대비 정정 123건, 이행 제외 25건, 1:N 분할 16건, 추후 고려 32건"],
  ["", "결정사항_모범안 / 인도용어정의", "결정 58건, 인도 용어 77개"],
  ["02_데이터사전_업무규칙.xlsx", "17 / 19 / 20 / 21번 시트", "시스템 설계용 CoA, 이행 맵핑·현행 대사, 검토 의견, 추후 고려사항"],
], [2.4, 2, 3.3]));

const doc = new Document({
  features: { updateFields: true },
  creator: "India Noida Branch", title: "통합표준 CoA 인도 반영 — 본점 협의 설명자료",
  styles: {
    default: { document: { run: { font: FONT, size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 30, bold: true, color: NAVY }, paragraph: { spacing: { before: 240, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 24, bold: true, color: NAVY }, paragraph: { spacing: { before: 220, after: 100 }, keepNext: true, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 21, bold: true }, paragraph: { spacing: { before: 160, after: 80 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 240 } } } },
      { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 240 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run("통합표준 CoA 인도 반영 · 본점 협의 설명자료", { size: 16, color: GREY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
