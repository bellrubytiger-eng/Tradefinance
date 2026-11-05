# -*- coding: utf-8 -*-
"""본점 양식(통합 표준 CoA 0930) 그대로 인도지점 반영본 생성.

python3 build_ho_format.py [출력경로]   (기본 ../docs/06_통합표준CoA_0930_인도지점반영.xlsx)
build_coa_review.py도 write_status_sheet()로 같은 시트를 검토서에 넣는다.
"""
import sys
from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import ho_format_model as H
import glossary_model as GL

FN = "맑은 고딕"
HO_HDR = PatternFill("solid", fgColor="1F4E79")
IN_HDR = PatternFill("solid", fgColor="C55A11")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ST_FILL = {"신규": "FCE4D6", "변경": "FFF2CC", "용도지정": "E2EFDA", "비활성": "E7E6E6", "철회": "D9D9D9"}

HO_COLS = [("구분", 12), ("업무영역", 18), ("코드", 14), ("Level", 9), ("계정과목(영문/한글)", 60), ("원화계정명", 12),
           ("본점 통합계정코드", 18), ("본점 통합계정과목명", 26), ("비화폐성 여부", 10), ("GBS코드", 12)]
IN_COLS = [("인도 상태", 10), ("인도 표시명 / 신규 계정명", 44), ("현행 인도 GBS(8자리)", 22), ("현행 잔액 INR(2026-07-07, 차변+)", 18),
           ("변경 사유·근거", 70), ("인도 재무제표 줄(참고)", 22), ("0930 정정 권고", 30), ("인도 용어 정의 (본점 참고)", 70)]


def write_status_sheet(ws, rs=None):
    rs = rs or H.rows()
    ws["A1"] = "농협은행 국외지점 통합 표준 CoA — 인도(노이다) 지점 반영"
    ws["A1"].font = Font(name=FN, size=14, bold=True)
    ws.merge_cells("A1:H1")
    n_ho, n_in = len(HO_COLS), len(IN_COLS)
    ws.cell(row=2, column=1, value="본점 원안 (0930, 열 구성·순서 그대로)").font = Font(name=FN, size=10, bold=True, color="1F4E79")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_ho)
    ws.cell(row=2, column=n_ho + 1, value="인도지점 반영 (노이다 검토 결과)").font = Font(name=FN, size=10, bold=True, color="C55A11")
    ws.merge_cells(start_row=2, start_column=n_ho + 1, end_row=2, end_column=n_ho + n_in)
    for i, (h, w) in enumerate(HO_COLS + IN_COLS, 1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = Font(name=FN, size=10, bold=True, color="FFFFFF")
        c.fill = HO_HDR if i <= n_ho else IN_HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[3].height = 32
    for r, x in enumerate(rs, 4):
        vals = [x["grp"], x["area"], x["code"], f'Level {x["lv"]}', x["name"], x["krw"], x["ho"], x["ho_name"], x["nonmon"], x["gbs"],
                x["status"] if x["status"] != "유지" else "", x["india_name"], x["live"], x["amt"], x["why"], x["sch"], x["fix"], x["gloss"]]
        fill = PatternFill("solid", fgColor=ST_FILL[x["status"]]) if x["status"] in ST_FILL else None
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=v if v != "" else None)
            c.font = Font(name=FN, size=9 if i != 18 else 8, bold=(x["lv"] <= 2 and i <= n_ho),
                          strike=(x["status"] == "철회" and i in (5, 12)), color="404040" if i == 18 else "000000")
            c.border = BOX
            c.alignment = Alignment(vertical="top", wrap_text=i in (12, 15, 17, 18),
                                    horizontal="center" if i in (1, 2, 3, 4, 9, 10, 11) else None)
            if fill and (x["status"] in ("신규", "철회") or i in (5, 11, 12)):
                c.fill = fill
            if i == 14:
                c.number_format = '#,##0;[Red]-#,##0'
    ws.freeze_panes = "F4"
    ws.auto_filter.ref = f"A3:{get_column_letter(n_ho + n_in)}{3 + len(rs)}"
    return rs


def write_summary(ws, rs):
    ws["A1"] = "통합 표준 CoA 구조 요약 — 인도(노이다) 지점 반영"
    ws["A1"].font = Font(name=FN, size=14, bold=True)
    hdr = ["구분", "Level1", "Level2", "Level3", "Level4", "Level5", "Level6"]
    for i, h in enumerate(hdr, 1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = Font(name=FN, size=10, bold=True, color="FFFFFF"); c.fill = HO_HDR; c.border = BOX
        c.alignment = Alignment(horizontal="center")
    for r, (g, l1) in enumerate([("자산(B/S)", "1"), ("부채(B/S)", "2"), ("자본(B/S)", "3"), ("수익(P/L)", "4"), ("비용(P/L)", "5"), ("각주", "6")], 4):
        for i, v in enumerate([g, l1, "00", "00", "00", "00", "00"], 1):
            c = ws.cell(row=r, column=i, value=v); c.font = Font(name=FN, size=10, bold=(i == 1)); c.border = BOX
    ws.cell(row=11, column=1, value="인도지점 반영 집계").font = Font(name=FN, size=12, bold=True, color="C55A11")
    cnt = Counter(x["status"] for x in rs)
    grp = Counter((x["grp"], x["status"]) for x in rs)
    groups = ["자산(B/S)", "부채(B/S)", "자본(B/S)", "수익(P/L)", "비용(P/L)", "각주"]
    sts = ["신규", "변경", "용도지정", "비활성", "철회", "유지"]
    for i, h in enumerate(["구분"] + sts + ["합계"], 1):
        c = ws.cell(row=12, column=i, value=h)
        c.font = Font(name=FN, size=10, bold=True, color="FFFFFF"); c.fill = IN_HDR; c.border = BOX
        c.alignment = Alignment(horizontal="center")
    for r, g in enumerate(groups + ["합계"], 13):
        vals = [g] + [(cnt[s] if g == "합계" else grp[(g, s)]) for s in sts]
        vals.append(sum(vals[1:]))
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=v); c.font = Font(name=FN, size=10, bold=(g == "합계" or i == 1)); c.border = BOX
            if i > 1:
                c.alignment = Alignment(horizontal="center")
            if 2 <= i <= 6 and sts[i - 2] in ST_FILL:
                c.fill = PatternFill("solid", fgColor=ST_FILL[sts[i - 2]])
    legend = [
        ("열 구성", "A~J열은 본점 0930 원안 그대로(값·순서 불변). K~R열(주황 머리)이 인도지점 검토 결과"),
        ("신규", "인도에 필요해 신설 요청하는 계정. 코드는 본점 채번 전 임시코드, GBS코드 '채번요청', 본점 통합계정은 같은 상위·같은 레벨 계정 기준 '(안)'"),
        ("변경", "인도 표시명·용도·직접기장 금지 등 변경. 계정과목 열은 본점명 유지, 인도 표시명은 L열"),
        ("용도지정", "인도 업무를 이 기존 계정에 기장하도록 지정(예: CC/OD 이자 → INT. ON OVERDRAFT)"),
        ("비활성", "인도지점 미사용(삭제 아님)"),
        ("철회", "이전 차수 제안을 철회(취소선). 이력 확인용"),
        ("현행 인도 GBS", "현행 8자리 계정 중 이 11자리로 옮기는 것. 잔액은 2026-07-07 BS/PL, 차변 +"),
        ("인도 재무제표 줄", "BR Act Third Schedule 기준 참고 줄(Sch.3 예금, Sch.16 영업비용 등). 최종은 결산 grouping"),
        ("근거 자료", "상세 판단·결산 증거·용어 정의는 검토서(05_CoA검토_RBI보고속성.xlsx) 시트 01~17"),
    ]
    ws.cell(row=21, column=1, value="읽는 법").font = Font(name=FN, size=12, bold=True)
    for r, (k, v) in enumerate(legend, 22):
        a = ws.cell(row=r, column=1, value=k); b = ws.cell(row=r, column=2, value=v)
        a.font = Font(name=FN, size=9, bold=True); b.font = Font(name=FN, size=9)
        if k in ST_FILL:
            a.fill = PatternFill("solid", fgColor=ST_FILL[k])
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        b.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 28
    ws.column_dimensions["A"].width = 16
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 12


def write_glossary(ws):
    ws["A1"] = "인도 특수 용어 정의 (본점 참고)"
    ws["A1"].font = Font(name=FN, size=14, bold=True)
    for i, (h, w) in enumerate([("분류", 10), ("용어", 30), ("정의", 90), ("근거", 32)], 1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = Font(name=FN, size=10, bold=True, color="FFFFFF"); c.fill = HO_HDR; c.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, (t, cat, d, b, _p) in enumerate(GL.G, 4):
        for i, v in enumerate([cat, t, d, b], 1):
            c = ws.cell(row=r, column=i, value=v); c.font = Font(name=FN, size=9); c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A4"


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "../docs/06_통합표준CoA_0930_인도지점반영.xlsx"
    wb = Workbook()
    s1 = wb.active
    s1.title = "요약"
    s2 = wb.create_sheet("통합표준CoA_계정상태표")
    rs = write_status_sheet(s2)
    write_summary(s1, rs)
    write_glossary(wb.create_sheet("인도용어정의"))
    wb.save(out)
    print("saved", out, len(rs), Counter(x["status"] for x in rs))
