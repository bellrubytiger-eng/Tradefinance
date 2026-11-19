# -*- coding: utf-8 -*-
"""RT11/RT13 개선 설계: 엑셀 설계서 + 화면 목업 생성
python3 tools/build_rt_design.py docs/09_RT11_RT13_개선설계.xlsx screens/rt_common.html"""
import json
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import rt_model as R

HERE = os.path.dirname(os.path.abspath(__file__))
F = "Arial"
HDR = PatternFill("solid", fgColor="1F4E79")
KIND = {"공통": "DDEBF7", "정책값": "FFF2CC", "확장": "F8CBAD", "공통 + 정책값": "E2EFDA"}
ACT = {"신규": "C6EFCE", "변경": "FFEB9C", "이동": "F4B084", "유지": "FFFFFF"}
thin = Side(style="thin", color="BFBFBF")
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
WR = Alignment(wrap_text=True, vertical="top")


def fill(c):
    return PatternFill("solid", fgColor=c)


def table(ws, r0, cols, rows, widths, colors=None):
    for c, h in enumerate(cols, 1):
        x = ws.cell(r0, c, h); x.font = Font(name=F, bold=True, color="FFFFFF", size=10); x.fill = HDR; x.border = BD
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, row in enumerate(rows, r0 + 1):
        for c, v in enumerate(row, 1):
            x = ws.cell(i, c, v); x.font = Font(name=F, size=9); x.alignment = WR; x.border = BD
            if colors and c in colors and v in colors[c]:
                x.fill = fill(colors[c][v])
            elif colors and c in colors and isinstance(v, str) and v.startswith("확장"):
                x.fill = fill(KIND["확장"])
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(r0 + 1, 1)


def xlsx(out):
    wb = Workbook()
    ws = wb.active; ws.title = "설계원칙"
    ws["A1"] = "RT11 · RT13 공통 모듈 개선 설계 (현행 화면 기준, 2026-11)"; ws["A1"].font = Font(name=F, bold=True, size=14, color="1F4E79")
    ws["A2"] = "입력: 현행 RT13 화면 캡처, IT 'GBS 글로벌 무역송금 전지점 확대 적용 설계서 V0.3', 08 해외지점 비교맵핑. 화면 목업: screens/rt_common.html"
    ws["A2"].font = Font(name=F, size=9, italic=True)
    table(ws, 4, ["원칙", "내용"], R.PRINCIPLES, [22, 120])
    ws.freeze_panes = None
    r = 6 + len(R.PRINCIPLES)
    ws.cell(r, 1, "구분 색").font = Font(name=F, bold=True, size=10)
    for i, (k, c) in enumerate(list(KIND.items())[:3]):
        ws.cell(r + 1 + i, 1, k).fill = fill(c)
        ws.cell(r + 1 + i, 2, {"공통": "전 지점 같은 로직", "정책값": "같은 로직, 값만 지점별(지점정책표)", "확장": "해당 국가만 추가 항목 · 규칙(확장 세트)"}[k]).font = Font(name=F, size=9)
    n = {}
    for f in R.FIELDS:
        n[f[3]] = n.get(f[3], 0) + 1
    ws.cell(r + 5, 1, "필드 처리 건수").font = Font(name=F, bold=True, size=10)
    for i, k in enumerate(["유지", "변경", "이동", "신규"]):
        ws.cell(r + 6 + i, 1, k).fill = fill(ACT[k]); ws.cell(r + 6 + i, 2, n.get(k, 0))

    s = wb.create_sheet("고객구분")
    table(s, 1, ["고객구분", "판정", "화면", "가운데 4개 탭", "한도", "증빙", "비고"], R.SEGMENTS, [12, 34, 8, 14, 40, 46, 36])

    f = wb.create_sheet("필드개선")
    table(f, 1, ["영역", "현행 항목", "개선 내용", "처리", "구분", "RT11", "RT13", "ISO 20022 / MT", "근거 · 비고"], R.FIELDS,
          [20, 36, 60, 7, 9, 6, 6, 26, 28], colors={4: ACT, 5: KIND})
    f.auto_filter.ref = f"A1:I{len(R.FIELDS) + 1}"

    p = wb.create_sheet("지점정책표")
    rows = [[a, b] + [v.get(k, "") for k in R.BR] for a, b, v in R.POLICY]
    table(p, 1, ["정책 항목", "설명"] + [f"{R.BR_NM[k]} ({k})" for k in R.BR], rows, [18, 24] + [20] * 7)
    for i in range(2, len(rows) + 2):
        for c in range(3, 10):
            v = str(p.cell(i, c).value or "")
            if "확인" in v:
                p.cell(i, c).fill = fill("FCE4D6")
    p.cell(len(rows) + 3, 1, "주황 = 지점 확인 전 가정값. 노이다 값은 현행 확정, 나머지는 08 비교맵핑(2차 자료) 기준 초안").font = Font(name=F, size=9, italic=True)

    m = wb.create_sheet("증빙매트릭스")
    table(m, 1, ["매트릭스 ID", "Product Type", "의미", "Invoice", "세관신고", "Dealing", "확장 규칙"], R.MATRIX, [12, 12, 30, 14, 20, 10, 44])

    v = wb.create_sheet("검증룰")
    table(v, 1, ["ID", "규칙", "구분", "화면", "출처"], R.RULES, [6, 80, 14, 14, 26], colors={3: KIND})

    o = wb.create_sheet("확인필요")
    table(o, 1, ["항목", "확인 내용", "영향"], R.OPEN, [26, 80, 28])
    wb.save(out)
    print("saved", out, len(R.FIELDS), "fields")


MOCK = {
    "IN": dict(ccy="INR", fy=4, purpose="IN Purpose Code (RBI)", ext="IN", plim=dict(cap=250000, used=180000, per="회계연도 4월~3월 (LRS)", over="불가"),
               corp=True, customs="Bill of Entry", dealing="Y", trav=None, cons=False, lei="INR 50 crore 이상 필수", sanc="UN · IN", mtx="IN-TRD"),
    "CN": dict(ccy="CNY", fy=1, purpose="CN 국제수지 거래코드 (확인)", ext="CN", plim=dict(cap=50000, used=20000, per="역년 (외화 구매 편의한도, 확인)", over="증빙 요구 (확인)"),
               corp=None, customs="报关单 (확인)", dealing="확인", trav=None, cons=False, lei=None, sanc="UN · CN", mtx="CN-TRD"),
    "VN": dict(ccy="VND", fy=1, purpose="VN 송금목적 분류 (확인)", ext="VN", plim=None, corp=None, customs="Tờ khai hải quan (확인)", dealing="확인",
               trav=None, cons=False, lei=None, sanc="UN · VN", mtx="VN-TRD"),
    "HK": dict(ccy="HKD", fy=4, purpose=None, ext=None, plim=None, corp=False, customs=None, dealing="확인", trav="HKD 8,000", cons=False, lei=None, sanc="UN · HK", mtx="COM-OPT"),
    "US": dict(ccy="USD", fy=1, purpose=None, ext="US", plim=None, corp=False, customs=None, dealing="확인", trav="USD 3,000", cons=True, lei=None, sanc="OFAC SDN", mtx="COM-OPT"),
    "GB": dict(ccy="GBP", fy=4, purpose=None, ext=None, plim=None, corp=False, customs=None, dealing="확인", trav=None, cons=False, lei="CHAPS 금융기관 지급 (확인)", sanc="UK Sanctions List", mtx="COM-OPT"),
    "AU": dict(ccy="AUD", fy=7, purpose=None, ext=None, plim=None, corp=False, customs=None, dealing="확인", trav=None, cons=False, lei=None, sanc="DFAT", mtx="COM-OPT"),
}


def html(out):
    tpl = open(os.path.join(HERE, "rt_mock_tpl.html"), encoding="utf-8").read()
    data = {"br": R.BR, "nm": R.BR_NM, "mock": MOCK, "matrix": [list(x) for x in R.MATRIX]}
    open(out, "w", encoding="utf-8").write(tpl.replace("/*DATA*/null", json.dumps(data, ensure_ascii=False)))
    print("saved", out)


if __name__ == "__main__":
    xlsx(sys.argv[1])
    html(sys.argv[2])
