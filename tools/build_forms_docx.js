// 고객 서식 명세서 Word 생성: node tools/build_forms_docx.js <model.json> <img_dir> <out.docx>
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType,
  WidthType, ShadingType, BorderStyle, LevelFormat, Footer, Header, PageNumber, ImageRun, PageBreak,
} = require("docx");

const M = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const IMG = process.argv[3];
const OUT = process.argv[4];
const FONT = "맑은 고딕";
const W = 9638;
const NAVY = "1F4E79";

const run = (t, o = {}) => new TextRun({ text: String(t), font: FONT, size: o.size || 20, bold: o.bold, color: o.color, italics: o.italics });
const P = (t, o = {}) => new Paragraph({ spacing: { after: 100, line: 300 }, alignment: o.align, children: [run(t, o)] });
const H1 = (t, br = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: br, spacing: { after: 200 }, children: [run(t, { size: 30, bold: true, color: NAVY })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 100 }, children: [run(t, { size: 24, bold: true, color: NAVY })] });
const B = (t) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 60, line: 280 }, children: [run(t)] });
const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };
function TBL(headers, rows, ratios, size = 16) {
  const sum = ratios.reduce((a, b) => a + b, 0);
  const widths = ratios.map((r) => Math.floor((W * r) / sum));
  widths[widths.length - 1] += W - widths.reduce((a, b) => a + b, 0);
  const cell = (t, i, hdr) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: hdr ? NAVY : "FFFFFF" },
    margins: { top: 40, bottom: 40, left: 70, right: 70 },
    children: [new Paragraph({ children: [run(t ?? "", { size, bold: hdr, color: hdr ? "FFFFFF" : undefined })] })],
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((v, i) => cell(v, i, false)) }))],
  });
}

const C = [];
C.push(new Paragraph({ spacing: { before: 2400 }, children: [] }));
C.push(P("GBS India 외환·무역송금", { size: 40, bold: true, color: NAVY }));
C.push(P("고객 서식 명세서", { size: 34, bold: true }));
C.push(P("FIRC(FDI 환전증명서) · 당발송금 확인서 · 타발송금 확인서 · 환전 확인서", { size: 22 }));
C.push(new Paragraph({ spacing: { before: 600 }, children: [] }));
C.push(TBL(["항목", "내용"], [
  ["대상", M.META.branch], ["버전 / 작성일", `${M.META.version} / ${M.META.date}`],
  ["시안", "screens/forms.html (A4 인쇄용, 변형 선택·DB 매핑 표시·DUPLICATE 보기)"],
  ["필드 매핑", "02_데이터사전_업무규칙.xlsx '12.서식필드매핑' 시트"],
], [1, 4], 18));

C.push(H1("1. 공통 원칙"));
[
  "모든 서식은 영문(인도 고객용)으로 발급하고, 은행명·주소·AD Code·SWIFT·GSTIN은 지점 설정값에서 채웁니다(시안의 [BANK NAME] 등 대괄호 값).",
  "원본은 원천 거래 1건당 1회 발급합니다. 이후 발급은 ISSUE_TYPE=DUP로 'DUPLICATE'를 표시하고 사유 입력과 책임자 승인을 받습니다(R-DOC-02).",
  "발급 시점의 표기값 전체를 TB_FX_ADVICE_ISSUE.SNAPSHOT_JSON에 저장해 재발급 시 동일하게 재현합니다. 원거래 정정·취소 시 기존 발급분을 CXL 처리하고 AMD로 재발급합니다(R-DOC-03).",
  "PAN은 앞 5자리·끝 1자리 외 마스킹, 계좌는 끝 4자리만 표시합니다. FIRC의 피투자법인 PAN·CIN은 FC-GPR 용도로 전체 표시합니다(R-DOC-04).",
  "진위확인 코드(VERIFY_CODE)를 QR과 문자로 표시하고 DC01에서 대조합니다(R-DOC-07). 시안의 QR은 패턴 예시이며 운영 시 QR 라이브러리로 생성합니다.",
  "수수료·세금 내역에는 과세 수수료, 환전 과세표준(32(2), 미징수 표시), CGST·SGST 또는 IGST, TCS를 구분해 표기하고, GST Details에 지점·고객 GSTIN, 공급장소, SAC를 표기합니다(R-GST-07).",
  "INR 확정 환전·FDI는 수수료를 별도징수하므로 원금 INR이 감액되지 않았음을 문구로 명시합니다(R-GST-06).",
].forEach((t) => C.push(B(t)));
C.push(H2("발급번호 체계"));
C.push(TBL(["서식", "번호 형식", "예시"], [
  ["FM-FDI-01 FIRC", "FIRC/지점/FY/일련6", "FIRC/NOI/2026-27/000041"],
  ["FM-OUT-01 당발", "ORA/지점/FY/일련6", "ORA/NOI/2026-27/018204"],
  ["FM-IN-01 타발", "IRA/지점/FY/일련6", "IRA/NOI/2026-27/007712"],
  ["FM-FX-01 환전", "FXA/지점/FY/일련6", "FXA/NOI/2026-27/002118"],
], [1.6, 1.6, 2]));

const IMG_OF = { firc: "firc-fdi.png", out: "out-imp.png", in: "in-exp.png", fx: "fx-acct.png" };
M.FORMS.forEach((f, n) => {
  C.push(H1(`${n + 2}. ${f.ko} (${f.id})`));
  C.push(TBL(["항목", "내용"], [
    ["영문 명칭", `${f.title} - ${f.subtitle}`], ["발급 화면", f.screen], ["발급 시점", f.when], ["용도", f.use], ["보관", f.copies],
    ["변형", f.variants.map((v) => v[1]).join(" / ")],
  ], [1, 4], 17));
  const img = path.join(IMG, IMG_OF[f.key]);
  if (fs.existsSync(img)) {
    C.push(new Paragraph({ spacing: { before: 200 }, alignment: AlignmentType.CENTER, children: [
      new ImageRun({ type: "png", data: fs.readFileSync(img), transformation: { width: 440, height: Math.round(440 * 1574 / 1112) },
        altText: { title: f.id, description: `${f.title} 시안`, name: f.id } })] }));
    C.push(P(`그림: ${f.id} 시안 (${f.variants[0][1]}, SPECIMEN)`, { size: 16, italics: true, align: AlignmentType.CENTER }));
  }
  C.push(new Paragraph({ children: [new PageBreak()] }));
  C.push(H2("항목 정의"));
  const vn = Object.fromEntries(f.variants);
  const rows = [];
  f.sections.forEach((s) => s.f.forEach((x) => rows.push([s.t, x[1], x[2], x[3], x[4] ? x[4].map((v) => vn[v]).join(", ") : "전체"])));
  C.push(TBL(["섹션", "영문 라벨", "한글 항목", "DB 소스", "표시 조건"], rows, [1.4, 1.6, 1.1, 2.6, 1.5], 14));
  C.push(H2("문구"));
  C.push(TBL(["No", "문구 (영문)", "표시 조건"], f.statements.map((s, i) => [String(i + 1), s[0], s[1] ? s[1].map((v) => vn[v]).join(", ") : "전체"]), [0.3, 5, 1.4], 15));
  C.push(P(`서명란: ${f.sign.join(" / ")}`, { size: 17 }));
});

const doc = new Document({
  creator: "GBS India FX Project", title: "고객 서식 명세서",
  styles: {
    default: { document: { run: { font: FONT, size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, font: FONT, color: NAVY }, paragraph: { spacing: { after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: FONT, color: NAVY }, paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run("GBS India 고객 서식 명세서 " + M.META.version, { size: 16, color: "808080" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16 })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
