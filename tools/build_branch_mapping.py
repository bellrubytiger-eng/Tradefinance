# -*- coding: utf-8 -*-
"""해외지점 비교 맵핑 초안 엑셀: python3 tools/build_branch_mapping.py docs/08_해외지점_비교맵핑_초안.xlsx"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import branch_model as B

F = "Arial"
HDR = PatternFill("solid", fgColor="1F4E79")
KR = PatternFill("solid", fgColor="DDEBF7")
LV = {"확인": PatternFill("solid", fgColor="E2EFDA"), "2차": PatternFill("solid", fgColor="FFF2CC"), "확인필요": PatternFill("solid", fgColor="FCE4D6")}
TP = {"공통": PatternFill("solid", fgColor="DDEBF7"), "정책값": PatternFill("solid", fgColor="FFF2CC"), "확장": PatternFill("solid", fgColor="F8CBAD")}
thin = Side(style="thin", color="BFBFBF")
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
WR = Alignment(wrap_text=True, vertical="top")


def head(ws, r, cols):
    for c, h in enumerate(cols, 1):
        x = ws.cell(r, c, h); x.font = Font(name=F, bold=True, color="FFFFFF", size=10); x.fill = HDR; x.border = BD
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def table(ws, r0, cols, rows, widths):
    head(ws, r0, cols)
    for i, row in enumerate(rows, r0 + 1):
        for c, v in enumerate(row, 1):
            x = ws.cell(i, c, v); x.font = Font(name=F, size=9); x.alignment = WR; x.border = BD
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(r0 + 1, 1)


def main(out):
    wb = Workbook()
    ws = wb.active; ws.title = "요약"
    ws["A1"] = "해외지점 공통 모듈 — 한국 기준 ↔ 7개국 비교 맵핑 (기초 조사 초안, 2026-11)"; ws["A1"].font = Font(name=F, bold=True, size=14, color="1F4E79")
    lines = [
        "목적: 송금방식 수출입·인보이스/세관신고·외국환 한도를 전 지점 공통 모듈로 만들기 위해, 나라별 차이를 공통 / 정책값 / 확장으로 분류",
        "공통 = 모든 지점 같은 로직 · 정책값 = 같은 로직, 값만 지점별 설정 · 확장 = 해당 국가만 추가 항목·규칙 (IT 설계서의 공통 코어 / 정책 설정형 / 확장필드와 같은 구분)",
        "확인 수준: 확인(공식·다수 출처 일치, 초록) / 2차(법률·은행 자료, 노랑) / 확인필요(출처 부족 또는 지점 확인, 주황). 운영 반영 전 지점별 확인 필수",
        "원칙: 기존 화면(RT11·RT13·RT73·RT74·XP73·타발)을 개선하는 방식. 국가 특화가 확정되지 않아도 공통·정책 구조는 먼저 설계 가능하고, 특화값은 지점 오픈 전 확정",
    ]
    for i, t in enumerate(lines, 3):
        ws.cell(i, 1, t).font = Font(name=F, size=10); ws.cell(i, 1).alignment = Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 150
    r = 9
    ws.cell(r, 1, "지점 개요").font = Font(name=F, bold=True, size=12, color="1F4E79")
    table(ws, r + 1, ["지점", "국가", "코드", "현지통화", "감독·외환당국", "결제망", "외환통제", "특화 강도", "메모"], B.BRANCHES, [150, 10, 6, 9, 22, 18, 9, 9, 50])
    ws.column_dimensions["A"].width = 14
    ws.freeze_panes = None

    m = wb.create_sheet("비교맵핑")
    cols = ["영역", "항목", "처리 구분"] + [B.COL_NM[k] for k in B.COLS] + ["설계 반영"]
    head(m, 1, cols)
    for i, (area, item, tp, cell, design) in enumerate(B.ROWS, 2):
        vals = [area, item, tp] + [cell[k][0] for k in B.COLS] + [design]
        for c, v in enumerate(vals, 1):
            x = m.cell(i, c, v); x.font = Font(name=F, size=9, bold=c == 2); x.alignment = WR; x.border = BD
        m.cell(i, 3).fill = TP[tp]
        for j, k in enumerate(B.COLS):
            m.cell(i, 4 + j).fill = LV[cell[k][1]]
    for c, w in enumerate([11, 22, 9] + [26] * 8 + [36], 1):
        m.column_dimensions[get_column_letter(c)].width = w
    m.freeze_panes = "D2"
    n = len(B.ROWS) + 3
    m.cell(n, 1, "범례").font = Font(name=F, bold=True, size=9)
    for i, (k, f) in enumerate(LV.items()):
        m.cell(n + 1 + i, 2, k).fill = f; m.cell(n + 1 + i, 3, {"확인": "공식·다수 출처 일치", "2차": "법률·은행 자료", "확인필요": "지점 확인 필요"}[k]).font = Font(name=F, size=9)

    q = wb.create_sheet("지점확인질문지")
    rows = []
    for k, item, kind in B.QUESTIONS:
        rows.append([k, item, kind] + [""] * 7)
    table(q, 1, ["구분", "확인 항목", "처리 구분", "노이다", "북경", "하노이", "홍콩", "뉴욕", "런던", "시드니"], rows, [10, 52, 10, 18, 18, 18, 18, 18, 18, 18])
    for c in range(4, 11):
        q.cell(2, c).fill = PatternFill()
    for i in range(2, 2 + len(rows)):
        q.cell(i, 4).value = "노이다 요건 완료(설계서 참조)"

    s = wb.create_sheet("범위")
    table(s, 1, ["화면", "업무", "현행", "공통 모듈 방향"], B.SCOPE, [10, 26, 44, 56])

    src = wb.create_sheet("출처")
    table(src, 1, ["국가", "내용", "URL"], B.SOURCES, [8, 46, 90])
    wb.save(out)
    print("saved", out, len(B.ROWS), "rows")


if __name__ == "__main__":
    main(sys.argv[1])
