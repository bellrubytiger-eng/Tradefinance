// 화면 흐름도 PPT 생성: node tools/build_pptx.js <model.json> <out.pptx>
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const applyTheme = process.env.PPTX_SKILL_DIR ? require(process.env.PPTX_SKILL_DIR + "/scripts/apply_theme.js").applyTheme : async () => {}; // 테마 색 적용(선택)

const M = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = process.argv[3];

// 업무 영역 색 (노드 채움) - 화면 시안의 메뉴 그룹과 동일한 구분
const DOM = {
  고객: { c: "1D2733", n: "고객" }, 당발송금: { c: "1559A6", n: "당발송금" }, 타발송금: { c: "0F7B6C", n: "타발송금" },
  환전: { c: "A87400", n: "환전·시뮬레이션" }, 수출입서류: { c: "6A5D8C", n: "수출입서류" }, 무역금융: { c: "A3541C", n: "무역금융·LC·보증" },
  FDI: { c: "7A3E9D", n: "FDI" }, "INR Vostro": { c: "2F4858", n: "INR Vostro" }, SWIFT: { c: "3E5C76", n: "SWIFT" },
  본지점: { c: "3D6B35", n: "본지점" }, "서식·보고": { c: "5D6B7C", n: "서식·보고" },
};
const SCR = {}; // id → [name, domain]
const GROUP = { ch11: "고객", rt11: "당발송금", rt12: "당발송금", rt13: "당발송금", rt14: "당발송금", rt21: "타발송금", rt22: "타발송금", rt25: "타발송금",
  fx11: "환전", sm11: "환전", fx21: "환전", rt73: "수출입서류", xp73: "수출입서류", tf11: "무역금융", tf31: "무역금융", tf41: "무역금융", lc11: "무역금융", bg11: "무역금융", tf81: "무역금융", sf11: "무역금융",
  fd11: "FDI", fd12: "FDI", fd13: "FDI", vs11: "INR Vostro", vs31: "INR Vostro", sw11: "SWIFT", sw12: "SWIFT", ra11: "본지점", io11: "본지점", io21: "본지점", io31: "본지점", cm21: "본지점",
  dc01: "서식·보고", mn01: "서식·보고", if01: "서식·보고" };
M.SCREENS.forEach((s) => { SCR[s[0]] = { name: s[1], tag: s[3], func: s[4], dom: GROUP[s[6]] || "서식·보고" }; });

const THEME = {
  name: "GBS Screen Flow", headFontFace: "Malgun Gothic", bodyFontFace: "Malgun Gothic",
  colors: { dk1: "1D2733", lt1: "FFFFFF", dk2: "14324F", lt2: "EEF2F7", accent1: "1559A6", accent2: "0F7B6C", accent3: "A3541C",
    accent4: "7A3E9D", accent5: "3D6B35", accent6: "A87400", hlink: "1559A6", folHlink: "6A5D8C" },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "GBS India 화면 흐름도";
pres.author = "GBS India FX Project";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "TITLE_DARK", background: { color: "14324F" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 11.7, h: 1.4, fontSize: 40, bold: true, color: "FFFFFF", valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.85, w: 11.7, h: 1.6, fontSize: 18, color: "CADCF0", valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.55, y: 0.35, w: 12.2, h: 0.75, fontSize: 30, bold: true, color: "14324F", valign: "middle" }, text: "" } },
    { text: { text: "GBS India 외환·무역송금·무역금융 · 화면 흐름도", options: { x: 0.55, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: "6B7A8C", isTextBox: true } } },
  ],
  slideNumber: { x: 12.3, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: "6B7A8C" },
});

let objN = 0;
const nid = (p) => `${p}-${++objN}`;
// 화면 노드
function node(slide, id, x, y, w = 1.75, h = 0.72, opt = {}) {
  const s = SCR[id.toUpperCase()] || { name: opt.name || id, dom: opt.dom || "서식·보고" };
  const col = (DOM[opt.dom || s.dom] || DOM["서식·보고"]).c;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: opt.ghost ? "FFFFFF" : col }, line: { color: col, width: opt.ghost ? 1.5 : 0.75, dashType: opt.ghost ? "dash" : "solid" }, objectName: nid("node-" + id) });
  slide.addText([
    { text: id.toUpperCase(), options: { bold: true, fontSize: 13, color: opt.ghost ? col : "FFFFFF", breakLine: true } },
    { text: opt.label || s.name, options: { fontSize: 10, color: opt.ghost ? "1D2733" : "FFFFFF" } },
  ], { x, y, w, h, align: "center", valign: "middle", margin: 2, isTextBox: true, objectName: nid("lbl-" + id) });
  return { x, y, w, h };
}
// 외부 시스템/결과 박스
function ext(slide, text, x, y, w = 1.75, h = 0.6) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: "EEF2F7" }, line: { color: "9AA8B9", width: 0.75 }, objectName: nid("ext") });
  slide.addText(text, { x, y, w, h, fontSize: 10, color: "1D2733", align: "center", valign: "middle", margin: 2, isTextBox: true, objectName: nid("ext-t") });
  return { x, y, w, h };
}
function decision(slide, text, x, y, w = 1.9, h = 1.0) {
  slide.addShape(pres.shapes.DIAMOND, { x, y, w, h, fill: { color: "FFF4D6" }, line: { color: "A87400", width: 1 }, objectName: nid("dec") });
  slide.addText(text, { x, y, w, h, fontSize: 10, color: "1D2733", align: "center", valign: "middle", margin: 2, isTextBox: true, objectName: nid("dec-t") });
  return { x, y, w, h };
}
// 화살표: 상자 a의 변 → 상자 b의 변 (side: r,l,t,b)
function pt(b, side) {
  return { r: [b.x + b.w, b.y + b.h / 2], l: [b.x, b.y + b.h / 2], t: [b.x + b.w / 2, b.y], b: [b.x + b.w / 2, b.y + b.h] }[side];
}
function arrow(slide, a, sa, b, sb, label, opt = {}) {
  const [x1, y1] = pt(a, sa), [x2, y2] = pt(b, sb);
  const w = Math.abs(x2 - x1), h = Math.abs(y2 - y1);
  slide.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(w, 0.001), h: Math.max(h, 0.001),
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: opt.color || "52627A", width: opt.width || 1.25, dashType: opt.dash ? "dash" : "solid", endArrowType: "triangle" }, objectName: nid("arr"),
  });
  if (label && Math.hypot(x2 - x1, y2 - y1) >= 1.2) { // 짧은 화살표는 라벨이 노드에 가리므로 생략(설명 상자에 기재)
    const lx = (x1 + x2) / 2, ly = (y1 + y2) / 2;
    slide.addText(label, { x: lx - 0.9 + (opt.dx || 0), y: ly - 0.32 + (opt.dy || 0), w: 1.8, h: 0.28, fontSize: 9, color: opt.color || "52627A", align: "center", isTextBox: true, margin: 0, objectName: nid("arr-l") });
  }
}
function note(slide, text, x, y, w, h) {
  slide.addText(text, { x, y, w, h, fontSize: 11, color: "33475B", valign: "top", isTextBox: true, margin: 4, fill: { color: "F5F8FC" }, objectName: nid("note") });
}
function addContent(title, section, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}

// ---------------- 1. 표지 ----------------
pres.addSection({ title: "개요" });
let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "개요" });
s.addText("화면 흐름도", { placeholder: "title" });
s.addText(`GBS India 노이다지점 · 외환·무역송금·무역금융 TO-BE 화면 ${M.SCREENS.length}종의 연결 관계\n${M.META.date} · ${M.META.version}`, { placeholder: "body" });

// ---------------- 2. 읽는 법 ----------------
s = addContent("읽는 법", "개요", "색은 업무 영역, 실선은 업무 순서, 점선은 조회·참조, 회색 상자는 외부 시스템이나 결과물입니다.");
const doms = Object.entries(DOM);
doms.forEach(([k, v], i) => {
  const col = i % 4, row = Math.floor(i / 4), x = 0.8 + col * 3.05, y = 1.5 + row * 0.85;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 0.55, h: 0.5, rectRadius: 0.06, fill: { color: v.c }, line: { color: v.c }, objectName: nid("leg") });
  const cnt = Object.values(SCR).filter((x) => x.dom === k).length;
  s.addText([{ text: v.n, options: { bold: true, fontSize: 14, breakLine: true } }, { text: `${cnt}개 화면`, options: { fontSize: 11, color: "6B7A8C" } }], { x: x + 0.65, y: y - 0.05, w: 2.3, h: 0.6, isTextBox: true, margin: 0, objectName: nid("leg-t") });
});
const a1 = ext(s, "외부 시스템 · 결과물", 0.8, 4.4, 2.3, 0.6);
const n1 = node(s, "rt21", 4.2, 4.35);
arrow(s, a1, "r", n1, "l", "업무 순서");
const n2 = node(s, "ch11", 7.4, 4.35);
arrow(s, n1, "r", n2, "l", "조회·참조", { dash: true });
decision(s, "분기 조건", 10.3, 4.2, 1.9, 1.0);
note(s, "각 슬라이드의 화면 ID는 화면 시안(screens/index.html)과 엑셀 '1.화면목록' 시트의 ID와 같습니다. NEW 화면과 기존 GBS 변경 화면은 마지막 목록 슬라이드에 구분해 두었습니다.", 0.8, 5.6, 11.7, 0.9);

// ---------------- 3. 전체 지도 ----------------
pres.addSection({ title: "전체 지도" });
s = addContent("전체 지도: 업무 영역과 핵심 연결", "전체 지도", "왼쪽 고객·시뮬레이션에서 시작해 송금·무역금융 처리 화면을 거쳐, 오른쪽의 메시지·지급·서식·보고로 흘러갑니다.");
const clusters = [
  ["고객", ["ch11"], 0.4, 1.35], ["환전", ["sm11", "fx11", "fx21"], 0.4, 2.35],
  ["당발송금", ["rt11", "rt12", "rt13", "rt14"], 2.75, 1.35], ["타발송금", ["rt21", "rt22", "rt25"], 2.75, 3.6],
  ["수출입서류", ["rt73", "xp73"], 5.1, 1.35], ["FDI", ["fd11", "fd12", "fd13"], 5.1, 3.6],
  ["무역금융", ["tf11", "tf31", "tf41", "lc11", "bg11", "tf81", "sf11"], 7.45, 1.35],
  ["SWIFT", ["sw11", "sw12"], 9.8, 1.35], ["INR Vostro", ["vs11", "vs31"], 9.8, 2.95],
  ["본지점", ["ra11", "io11", "io21", "io31", "cm21"], 9.8, 4.4], ["서식·보고", ["dc01", "mn01", "if01"], 0.4, 5.35],
];
const box = {};
clusters.forEach(([d, ids, x, y]) => {
  const col = DOM[d].c, w = 2.15, h = 0.36 + Math.ceil(ids.length / 2) * 0.42;
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: "FFFFFF" }, line: { color: col, width: 1.25 }, objectName: nid("cl") });
  s.addText(DOM[d].n, { x: x + 0.08, y: y + 0.03, w: w - 0.16, h: 0.28, fontSize: 11, bold: true, color: col, isTextBox: true, margin: 0, objectName: nid("cl-t") });
  ids.forEach((id, i) => {
    const bx = x + 0.08 + (i % 2) * 1.02, by = y + 0.36 + Math.floor(i / 2) * 0.42;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: by, w: 0.97, h: 0.34, rectRadius: 0.05, fill: { color: col }, line: { color: col }, objectName: nid("mini") });
    s.addText(id.toUpperCase(), { x: bx, y: by, w: 0.97, h: 0.34, fontSize: 10, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0, objectName: nid("mini-t") });
  });
  box[d] = { x, y, w, h };
});
arrow(s, box["고객"], "r", box["당발송금"], "l");
arrow(s, box["환전"], "r", box["당발송금"], "l");
arrow(s, box["당발송금"], "r", box["수출입서류"], "l");
arrow(s, box["타발송금"], "r", box["FDI"], "l");
arrow(s, box["수출입서류"], "r", box["무역금융"], "l");
arrow(s, box["무역금융"], "r", box["SWIFT"], "l");
arrow(s, box["SWIFT"], "b", box["INR Vostro"], "t");
arrow(s, box["INR Vostro"], "b", box["본지점"], "t", "본점 계좌");
arrow(s, box["FDI"], "r", box["무역금융"], "b", "SFMS 지급", { dx: 0.4 });
note(s, "연결: 시뮬레이션 견적 → 당발 실거래, 당발 ↔ Invoice·BoE, 타발 → FDI 이관, 서류 → 무역금융(EDPMS), 무역금융 → SWIFT(MT7xx·pacs). 모든 송금·무역금융 화면은 CH11 고객 프로필(PAN·IEC·LEI·GSTIN)을 조회하고, 발급 서식은 DC01, 기한은 MN01, RBI 보고는 IF01로 모입니다.", 2.75, 5.7, 6.9, 1.1);

// ---------------- 업무별 흐름 ----------------
pres.addSection({ title: "업무별 흐름" });
// 4. 당발
s = addContent("당발송금: 견적에서 송신·확인서까지", "업무별 흐름", "고객 유형에 따라 RT11·RT12·RT13·RT14 중 하나로 들어가고, 모두 SW11에서 승인·송신됩니다.");
let sm = node(s, "sm11", 0.5, 3.0);
let d1 = decision(s, "고객·목적", 2.6, 2.86);
let r11 = node(s, "rt11", 5.0, 1.3), r12 = node(s, "rt12", 5.0, 2.35), r13 = node(s, "rt13", 5.0, 3.4), r14 = node(s, "rt14", 5.0, 4.45);
arrow(s, sm, "r", d1, "l", "견적");
arrow(s, d1, "t", r11, "l", "기업 일반", { dy: -0.1 }); arrow(s, d1, "r", r12, "l", "개인 LRS", { dy: -0.15 }); arrow(s, d1, "r", r13, "l", "무역·ODI", { dy: 0.3 }); arrow(s, d1, "b", r14, "l", "비거주자", { dy: 0.2 });
let r73 = node(s, "rt73", 7.4, 4.45);
arrow(s, r13, "r", r73, "t", "Invoice·BoE", { dash: true, dx: 0.3 });
let sw = node(s, "sw11", 7.4, 2.35);
[r11, r12, r13].forEach((b) => arrow(s, b, "r", sw, "l"));
let dc = node(s, "dc01", 9.8, 1.3), ifb = node(s, "if01", 9.8, 2.35), mn = node(s, "mn01", 9.8, 3.4);
arrow(s, sw, "r", dc, "l", "FM-OUT-01"); arrow(s, sw, "r", ifb, "l", "ORM·CIMS"); arrow(s, sw, "r", mn, "l", "선수금 기한");
ext(s, "SWIFT pacs.008 (MX)", 9.8, 4.55, 1.75);
arrow(s, sw, "b", { x: 9.8, y: 4.55, w: 1.75, h: 0.6 }, "l", "", { dash: false });
note(s, "RT12 LRS: 회계연도 USD 250,000 한도·TCS 자동계산 / RT13: 인보이스·BoE 잔액 안에서 분할송금, IDPMS ORM 생성 / RT14: NRO 연 USD 1M", 0.5, 5.75, 12.2, 0.75);

// 5. 타발
s = addContent("타발송금: 결제방식별 분기", "업무별 흐름", "RT21에서 Purpose Code와 결제방식(SETTLE_MODE)에 따라 다음 화면이 정해집니다.");
let swi = node(s, "sw11", 0.4, 3.05, 1.75, 0.72, { label: "수신 pacs.008" });
let r21 = node(s, "rt21", 2.6, 3.05);
arrow(s, swi, "r", r21, "l");
let d2 = decision(s, "Purpose ·\n결제방식", 4.75, 2.92);
arrow(s, r21, "r", d2, "l");
const tgt = [["dc01", 1.55, "계좌 입금", "FM-IN-01"], ["rt25", 2.5, "수출대금", "IRM↔SB"], ["fd11", 3.45, "FDI (P0006)", "FDI 등록"], ["rt22", 4.4, "비계좌·보류", "미지급금"], ["fd13", 5.35, "에스크로", "예치"]];
tgt.forEach(([id, y, lab, l2]) => { const b = node(s, id, 7.6, y - 0.2); arrow(s, d2, id === "fd11" ? "r" : y < 3.2 ? "t" : "b", b, "l", lab, { dy: -0.02 }); });
let sf = node(s, "sf11", 10.3, 4.2);
arrow(s, { x: 7.6, y: 4.2, w: 1.75, h: 0.72 }, "r", sf, "l", "RTGS/NEFT");
let ife = node(s, "if01", 10.3, 2.3);
arrow(s, { x: 7.6, y: 2.3, w: 1.75, h: 0.72 }, "r", ife, "l", "EDPMS");
note(s, "Vostro 계좌 출금으로 인도 수취인에게 입금하는 건도 타발송금으로 처리합니다(R-VS-06).", 10.0, 5.25, 2.9, 1.0);

// 6. FDI
pres.addSection({ title: "상세 프로세스" });
s = addContent("FDI: 약정금액만 환전, 초과분은 환전 없이 반환", "상세 프로세스", "투자금액이 확정되어 있으므로 송금인이 더 보내도 약정분만 환전하고, 초과분은 원통화 그대로 원송금인에게 돌려보냅니다(R-FDI-06~08).");
let f21 = node(s, "rt21", 0.4, 2.95);
let f11 = node(s, "fd11", 2.55, 2.95, 1.75, 0.72, { label: "FDI 등록·KYC" });
arrow(s, f21, "r", f11, "l", "P0006");
let kyc = ext(s, "송금인 KYC\ncamt.110 → camt.111", 2.55, 1.3, 1.75, 0.75);
arrow(s, f11, "t", kyc, "b", "", { dash: true });
let d3 = decision(s, "수취액 vs\n약정분", 4.85, 2.82, 1.95, 1.0);
arrow(s, f11, "r", d3, "l");
let conv = ext(s, "약정분만 환전\nINR 입금 (원금 보전)", 7.45, 1.25, 2.1, 0.75);
let exc = node(s, "fd12", 7.45, 2.95, 2.1, 0.72, { label: "초과분 EXCESS 반환" });
let hold = ext(s, "부족: 환전 보류\n부족분 요청·약정 확인", 7.45, 4.6, 2.1, 0.75);
arrow(s, d3, "t", conv, "l", "같거나 많음", { dy: -0.1 });
arrow(s, d3, "r", exc, "l", "초과분");
arrow(s, d3, "b", hold, "l", "부족", { dy: 0.15 });
let firc = node(s, "dc01", 10.35, 1.25, 2.3, 0.75, { label: "FIRC (약정 환전분만)" });
arrow(s, conv, "r", firc, "l");
let ret = ext(s, "pacs.008 원통화 반환\n환전·환전 GST 없음", 10.35, 2.95, 2.3, 0.72);
arrow(s, exc, "r", ret, "l", "5영업일");
note(s, "예: 약정 INR 25 crore, 환율 88.42 → 환전 대상 USD 2,827,414.62. 수취 USD 2,850,000이면 초과 USD 22,585.38은 외화 미지급금으로 분리 후 반환(수수료 USD 25 공제). 배정 60일·미배정 반환 75일·FC-GPR 30일 기한은 FD11에서 함께 관리.", 0.4, 5.75, 12.4, 0.8);

// 7. 무역금융
s = addContent("무역금융: 등록 → 승인 → A/P → SFMS 지급", "상세 프로세스", "할인·매입·LC매입·추심·OA·TReDS는 TF11 한 화면에서 등록하고, 순지급액은 미지급금(21100000)을 거쳐 SF11에서 지급합니다.");
let t31 = node(s, "tf31", 0.4, 1.3), t41 = node(s, "tf41", 0.4, 2.45);
let docs = node(s, "xp73", 0.4, 3.6, 1.75, 0.72, { label: "수출서류 Lodgement" });
let t11 = node(s, "tf11", 3.0, 2.45, 1.9, 0.9);
arrow(s, t31, "r", t11, "t", "5층 한도", { dash: true, dx: -0.3 }); arrow(s, t41, "r", t11, "l", "개설은행 등급", { dash: true }); arrow(s, docs, "r", t11, "b", "서류·S/B", { dash: true, dx: -0.3 });
let chk = ext(s, "Checker / Authorizer\n정책 프리미엄 결재선", 5.55, 2.5, 2.0, 0.8);
arrow(s, t11, "r", chk, "l");
let ap = ext(s, "A/P 21100000\n순지급액 계상", 8.15, 2.5, 1.9, 0.8);
arrow(s, chk, "r", ap, "l", "승인");
let s11 = node(s, "sf11", 10.7, 2.5, 1.9, 0.8);
arrow(s, ap, "r", s11, "l");
let sfms = ext(s, "Velocis → SFMS\npacs.008 / 002 / 004", 10.7, 4.0, 1.9, 0.75);
arrow(s, s11, "b", sfms, "t");
let t81 = node(s, "tf81", 5.55, 4.4, 2.0, 0.72);
arrow(s, t11, "b", t81, "l", "만기·연체", { dash: true, dy: 0.25 });
let rep = node(s, "if01", 8.15, 4.4, 1.9, 0.72, { label: "EDPMS·IDPMS·CRILC" });
arrow(s, t81, "r", rep, "l");
note(s, "D/A·D/P 추심과 오픈어카운트는 자금을 제공하지 않으므로 한도 차감·A/P 없이 추심·EDPMS 관리만 합니다. 반환(pacs.004) 수신 시 A/P가 자동 복원됩니다.", 0.4, 5.75, 12.4, 0.75);

// 8. LC·보증
s = addContent("신용장·보증", "상세 프로세스", "LC·보증은 2026년 4월 시행 비자금여신 규정에 따라 자금여신 거래 고객 원칙을 먼저 확인합니다.");
let lc = node(s, "lc11", 0.5, 1.5, 1.9, 0.8);
let lcm = ext(s, "MT700 (SWIFT)\nIFN 700 (SFMS)", 3.0, 1.5, 1.9, 0.8);
arrow(s, lc, "r", lcm, "l", "개설");
let pres1 = ext(s, "서류 제시\n5은행영업일 심사", 5.5, 1.5, 1.9, 0.8);
arrow(s, lcm, "r", pres1, "l");
let d4 = decision(s, "일치?", 8.0, 1.4, 1.6, 1.0);
arrow(s, pres1, "r", d4, "l");
let pay = ext(s, "결제 / 인수\nMT740 상환", 10.3, 0.95, 2.2, 0.7);
let ref = node(s, "tf41", 10.3, 2.05, 2.2, 0.72, { label: "734 거절 → 개설은행 재평가" });
arrow(s, d4, "r", pay, "l", "일치", { dy: -0.15 }); arrow(s, d4, "b", ref, "l", "하자·거절", { dy: 0.2 });
let bg = node(s, "bg11", 0.5, 3.9, 1.9, 0.8);
let bgm = ext(s, "국내 IFN 760 + 760COV\n해외 MT760", 3.0, 3.9, 1.9, 0.8);
arrow(s, bg, "r", bgm, "l", "발행");
let cl = ext(s, "청구만료일\n만료 + 12개월 이상", 5.5, 3.9, 1.9, 0.8);
arrow(s, bgm, "r", cl, "l");
let inv = decision(s, "청구?", 8.0, 3.8, 1.6, 1.0);
arrow(s, cl, "r", inv, "l");
let dev = node(s, "tf81", 10.3, 3.45, 2.2, 0.72, { label: "대지급 → 자산건전성" });
let fem = node(s, "if01", 10.3, 4.55, 2.2, 0.72, { label: "FEMA 보증 분기 보고" });
arrow(s, inv, "r", dev, "l", "지급", { dy: -0.15 }); arrow(s, bg, "b", fem, "l", "비거주자 관련", { dash: true, dy: 0.25 });
note(s, "LC 매입(NEGOTIATION)은 TF11로 이어져 할인계산·A/P 지급을 탑니다. 본점이 개설한 LC의 상환은 RA11(다음 슬라이드)에서 처리합니다.", 0.5, 5.75, 12.2, 0.75);

// 9. 본지점·SWIFT
s = addContent("본지점: RA 대지급·정산·휴일 차이", "상세 프로세스", "본점 개설 LC의 상환수권(MT740)으로 지점이 대지급하고, 을A 거래로 본점과 정산합니다. 을B 자금거래와 갑기금은 분리합니다.");
let m740 = ext(s, "본점 MT740 수권\nMT742 청구 수신", 0.4, 1.5, 2.0, 0.8);
let ra = node(s, "ra11", 2.9, 1.5, 1.9, 0.8);
arrow(s, m740, "r", ra, "l");
let p9 = node(s, "sw11", 5.3, 1.5, 1.9, 0.8, { label: "pacs.009 지급·MT756" });
arrow(s, ra, "r", p9, "l", "IN 영업일 3일");
let i21 = node(s, "io21", 7.7, 1.5, 1.9, 0.8, { label: "을A 미정산·정산배치" });
arrow(s, p9, "r", i21, "l", "Due from HO");
let hoP = ext(s, "본점 입금\ncamt.054 / pacs.009", 10.2, 1.5, 2.2, 0.8);
arrow(s, i21, "r", hoP, "l", "건수·금액·주기");
let c21 = node(s, "cm21", 2.9, 3.4, 1.9, 0.8);
arrow(s, c21, "t", ra, "b", "임시휴일 재계산", { dash: true, dx: 0.4 });
let i11 = node(s, "io11", 5.3, 3.4, 1.9, 0.8, { label: "갑 0% · 을A · 을B 금리" });
arrow(s, i11, "r", i21, "b", "금리·환율", { dash: true, dx: 0.4 });
let i31 = node(s, "io31", 7.7, 3.4, 1.9, 0.8, { label: "을B 차입·예치" });
arrow(s, i11, "r", i31, "l", "", { dash: true });
let mt320 = ext(s, "MT320 확인\n우대: 1년 내 조달이력", 10.2, 3.4, 2.2, 0.8);
arrow(s, i31, "r", mt320, "l");
let m999 = ext(s, "본점 일괄 통보 MT999", 0.4, 3.4, 2.0, 0.8);
arrow(s, c21, "l", m999, "r");
note(s, "RA11 → SW11: 청구 접수 후 인도 영업일 3일 내 지급 / SW11 → IO21: Due from HO(을A) 기장 / IO21 → 본점: 건수·금액·주기 트리거로 정산배치. 휴일 차이 이자는 고객에게 청구하지 않고, 건당 USD 300 이하는 각 점 자체손익으로 흡수, 초과분은 을A 정산에 포함합니다.", 0.4, 5.55, 12.4, 1.0);

// 10. Vostro·SWIFT·환전
s = addContent("Vostro·SWIFT 메시지·환전", "상세 프로세스", "SW11이 모든 SWIFT 송수신의 허브이고, MT950은 VS31 자동 또는 SW12 수기로 보냅니다. 환전은 FX21 고시환율을 FX11과 SM11이 함께 씁니다.");
let v11 = node(s, "vs11", 0.4, 1.35), v31 = node(s, "vs31", 2.6, 1.35), s12 = node(s, "sw12", 2.6, 2.6);
let hub = node(s, "sw11", 5.0, 2.2, 2.0, 0.9, { label: "송수신 허브 (MT/MX)" });
arrow(s, v11, "r", v31, "l", "명세서 설정"); arrow(s, v31, "r", hub, "l", "MT950 자동"); arrow(s, s12, "r", hub, "l", "수기 작성");
const outs = [["rt21", 1.3, "pacs.008 수신"], ["ra11", 2.3, "MT740/742"], ["io21", 3.3, "camt.054"]];
outs.forEach(([id, y, l]) => { const b = node(s, id, 8.0, y); arrow(s, hub, "r", b, "l", l); });
let f21x = node(s, "fx21", 0.4, 4.55, 1.75, 0.72);
let s11b = node(s, "sm11", 2.6, 4.0), f11b = node(s, "fx11", 2.6, 5.05, 1.75, 0.72);
arrow(s, f21x, "r", s11b, "l", "고시환율", { dash: true }); arrow(s, f21x, "r", f11b, "l", "", { dash: true });
let dcx = node(s, "dc01", 5.0, 5.05, 2.0, 0.72, { label: "FM-FX-01 환전 확인서" });
arrow(s, f11b, "r", dcx, "l");
let tx = ext(s, "RT11~14 · FD11\n실거래 전환", 5.0, 3.95, 2.0, 0.72);
arrow(s, s11b, "r", tx, "l", "견적번호");
note(s, "VS11 명세서 설정 → VS31 EOD MT950 자동, SW12 수기 작성 → SW11 승인·송신. 결제성 메시지는 MX(pacs·camt)로 송신, MT950은 2028년 11월 은퇴 전까지 유지, 신용장·보증(MT7xx)은 MT 유지.", 8.0, 4.4, 4.7, 1.4);

// 11. 서식·보고
s = addContent("서식·기한·RBI 보고의 집결 지점", "상세 프로세스", "처리 화면에서 나온 서식·기한·보고 데이터가 DC01·MN01·IF01로 모입니다.");
const srcs = [["rt12", 1.2], ["rt13", 2.05], ["rt21", 2.9], ["fd11", 3.75], ["fx11", 4.6], ["tf11", 5.45]];
const dcb = node(s, "dc01", 6.0, 1.6, 2.2, 0.85, { label: "서식 4종 발급·재발급" });
const mnb = node(s, "mn01", 6.0, 3.15, 2.2, 0.85, { label: "규제기한 대시보드" });
const ifb2 = node(s, "if01", 6.0, 4.7, 2.2, 0.85, { label: "RBI 인터페이스" });
srcs.forEach(([id, y]) => { const b = node(s, id, 0.6, y - 0.1, 1.75, 0.62); arrow(s, b, "r", y < 2.6 ? dcb : y < 4.2 ? mnb : ifb2, "l", "", { color: "9AA8B9", width: 1 }); });
ext(s, "FIRC · 당발 · 타발 · 환전 확인서", 9.2, 1.7, 3.4, 0.65);
ext(s, "선수금 증빙 · 수출 회수 · FDI 60/75/30일 · 을B 만기", 9.2, 3.25, 3.4, 0.65);
ext(s, "IDPMS · EDPMS · CIMS · CRILC · FEMA 보증 · FIRMS", 9.2, 4.8, 3.4, 0.65);
arrow(s, dcb, "r", { x: 9.2, y: 1.7, w: 3.4, h: 0.65 }, "l"); arrow(s, mnb, "r", { x: 9.2, y: 3.25, w: 3.4, h: 0.65 }, "l"); arrow(s, ifb2, "r", { x: 9.2, y: 4.8, w: 3.4, h: 0.65 }, "l");

// ---------------- 화면 목록 ----------------
pres.addSection({ title: "화면 목록" });
const rows = M.SCREENS.map((x) => [x[0], x[1], (DOM[SCR[x[0]].dom] || {}).n || "", x[3]]);
const half = Math.ceil(rows.length / 2);
[rows.slice(0, half), rows.slice(half)].forEach((part, i) => {
  s = addContent(`화면 목록 (${i + 1}/2)`, "화면 목록", "신규는 이번에 새로 만드는 화면, 변경은 기존 GBS 화면을 고치는 것입니다.");
  const hdr = ["ID", "화면", "영역", "구분"].map((t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: "14324F" }, fontSize: 11 } }));
  s.addTable([hdr, ...part.map((r) => r.map((v, j) => ({ text: v, options: { fontSize: 11, bold: j === 0, color: j === 3 && v === "변경" ? "A3541C" : "1D2733" } })))], {
    x: 0.55, y: 1.3, w: 12.2, colW: [1.1, 6.2, 3.2, 1.7], rowH: 0.29, border: { type: "solid", color: "D9E0E8", pt: 0.75 }, objectName: nid("tbl"),
  });
});

pres.writeFile({ fileName: OUT }).then(async () => { await applyTheme(OUT, THEME); console.log("saved", OUT); });
