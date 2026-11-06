# -*- coding: utf-8 -*-
"""06 통합표준 CoA(인도 반영본) 정정본 생성
   python3 tools/build_coa06_fix.py "docs/06_통합표준CoA_0930_인도반영_정정본.xlsx"

원본 서식·열 구성을 유지하고:
- A~J(본점 0930 원안) 값은 바꾸지 않음 (0930 정정 권고는 R열에 추가)
- K~S(인도 검토) 정정 + T·U열에 행별 정정 내용·근거
- 시트 추가: 정정이력 / 현행이행제외 / 이행분할기준 / 추후고려사항
"""
import copy
import os
import sys
from collections import Counter, defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import coa_model as C

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "data", "06_CoA_0930_india_original.xlsx")
FIX = PatternFill("solid", fgColor="FFF2CC")      # 정정 셀
NEWR = PatternFill("solid", fgColor="F4B084")     # 신설 행
HDR = PatternFill("solid", fgColor="C55A11")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F9 = Font(name="Arial", size=9)
TODAY = "2026-11-06"


def main(out):
    wb = load_workbook(SRC)
    ws = wb["통합표준CoA_계정상태표"]
    HR = 3
    log = []  # (code, 열, 전, 후, 근거)
    row_of = {str(ws.cell(r, 3).value): r for r in range(HR + 1, ws.max_row + 1) if ws.cell(r, 3).value}

    def setc(code, col, val, why):
        r = row_of[code]
        cell = ws.cell(r, col)
        old = cell.value
        if (old or "") == (val or ""):
            return
        cell.value = val
        cell.fill = FIX
        log.append((code, ws.cell(HR, col).value, old, val, why))
        note = ws.cell(r, 20)
        note.value = ((note.value + "\n") if note.value else "") + f"{ws.cell(HR, col).value}: {why}"

    # T·U 열 머리
    for col, title in ((20, f"정정 내용 ({TODAY})"), (21, "정정 근거")):
        h = ws.cell(HR, col, title)
        src = ws.cell(HR, 19)
        h.font = copy.copy(src.font); h.fill = HDR; h.alignment = copy.copy(src.alignment); h.border = copy.copy(src.border)
    ws.column_dimensions["T"].width = 48
    ws.column_dimensions["U"].width = 22

    # 1) G3: 20510000000 신설 행 (20501000000 다음 = 20510100000 앞)
    r0 = row_of["20501000000"]
    ws.insert_rows(r0 + 1)
    for c in range(1, 22):
        s = ws.cell(r0, c)
        d = ws.cell(r0 + 1, c)
        d.font = copy.copy(s.font); d.border = copy.copy(s.border); d.alignment = copy.copy(s.alignment); d.number_format = s.number_format
    vals = {1: ws.cell(r0, 1).value, 2: ws.cell(r0, 2).value, 3: "20510000000", 4: "Level 3", 5: "    OTHER LIABILITIES", 7: ws.cell(r0, 7).value,
            8: ws.cell(r0, 8).value, 10: "채번요청", 11: "신규", 12: "사용", 13: "기타부채 L3 (20501000000 재채번)",
            16: "결정 G3: 산식상 하위 L4 20510100000~20510500000(TDS·GST·직원부채 포함)의 상위. 본점 0930의 20501000000을 이 코드로 재채번 요청",
            17: ws.cell(r0, 17).value, 20: "신설 행 (G3 미반영분)", 21: "결정 G3"}
    for c, v in vals.items():
        ws.cell(r0 + 1, c, v)
    for c in range(1, 22):
        ws.cell(r0 + 1, c).fill = NEWR
    log.append(("20510000000", "행", None, "신설", "결정 G3 — 하위 L4 5개의 상위 L3 누락"))
    row_of = {str(ws.cell(r, 3).value): r for r in range(HR + 1, ws.max_row + 1) if ws.cell(r, 3).value}
    setc("20501000000", 11, "변경", "결정 G3")
    setc("20501000000", 13, "→ 20510000000으로 재채번(G3), 이 코드는 미사용", "결정 G3")
    setc("20501000000", 12, "미사용", "결정 G3")
    ws.cell(row_of["20501000000"], 21).value = "결정 G3"

    # 2) 충당금 기한부 사용 (결정 L3)
    for code in ("10380101000", "10380102000"):
        setc(code, 11, "변경", "결정 L3 — 상태 '비활성'과 불일치")
        setc(code, 12, "사용(~2027-03-31)", "결정 L3: 2027-03-31까지 사용 후 ECL Stage 1~3 이관")
        r = row_of[code]
        ws.cell(r, 16).value = ((ws.cell(r, 16).value or "") + " / 기한부 사용: 2027-03-31까지 일반·특정 충당금 기장, ECL 시행일에 10380103000~105000 이관 + 전환 조정 31110800000").strip(" /")
        ws.cell(r, 16).fill = FIX
        ws.cell(r, 21).value = "결정 L3"

    # 3) 현행 8자리(N열) 재작성: 정정 후 최종 맵핑
    inv = defaultdict(list)
    for g, ts in C.FINAL_MAP.items():
        for t in ts:
            inv[t].append(g)
    reasons = defaultdict(list)
    for (g, code), w in C.FINAL_WHY.items():
        reasons[code].append(f"{g} 제거: {w}")
    added = {u["cur"]: u for u in C.CUR_UNMAPPED if u["code"]}
    for code, r in row_of.items():
        if code not in C.BY and code != "20510000000":
            continue
        new = ", ".join(sorted(inv.get(code, []))) or None
        old = ws.cell(r, 14).value
        if (str(old or "")).replace(" ", "") == (new or "").replace(" ", ""):
            continue
        olds = set(str(old or "").replace(" ", "").split(",")) - {""}
        news = set((new or "").replace(" ", "").split(",")) - {""}
        parts = []
        if olds - news:
            gone = sorted(olds - news)
            rs = [x for x in reasons.get(code, [])] or [f"{g}: " + next((e[3] for e in C.EXCLUDED if e[0] == g), "leaf로 이동") for g in gone]
            parts.append("제거 " + ", ".join(gone) + " (" + "; ".join(rs)[:200] + ")")
        if news - olds:
            add = sorted(news - olds)
            src = []
            for g in add:
                if g in added:
                    src.append(f"{g} 현행 미연결 → {added[g]['kind']}")
                else:
                    src.append(f"{g} 상위 계정에서 이동")
            parts.append("추가 " + "; ".join(src)[:250])
        setc(code, 14, new, " / ".join(parts))
        ws.cell(r, 21).value = ws.cell(r, 21).value or ("검토: 현행 BS/PL 대사" if news - olds else "결정 G1·G2")

    # 잔액(O열): 정정 후 최종 맵핑 기준 재계산 (1:1·N:1 대상), 1:N 분할 대상은 원본 유지
    split_t = {t for g, ts in C.FINAL_MAP.items() if len(ts) > 1 for t in ts}
    for code, r in row_of.items():
        if code in split_t:
            continue
        srcs = inv.get(code, [])
        v = round(sum(C.CUR[g]["dr"] for g in srcs if g in C.CUR), 2) if srcs else None
        v = v if v else None
        old = ws.cell(r, 15).value
        if (old or 0) != (v or 0):
            setc(code, 15, v, "잔액 재계산(정정 후 맵핑, 2026-07-07 BS/PL)")

    # 4) 0930 GBS코드 열 정정 권고 (R열) — 파생
    drv = {"10910100300": "GBS 14101001은 통화스왑·트레이딩 → 10910200200, 이 계정(금리스왑·트레이딩)은 14101002",
           "10910200200": "GBS 14101012는 금리스왑·헤지 → 10920100300, 이 계정(통화스왑·트레이딩)은 14101001",
           "10920100300": "GBS 14101002는 금리스왑·트레이딩 → 10910100300, 이 계정(금리스왑·헤지)은 14101012",
           "20810100300": "GBS 24101001은 통화스왑·트레이딩 → 20810200200, 이 계정은 24101002",
           "20810200200": "GBS 24101012는 금리스왑·헤지 → 20820100300, 이 계정은 24101001",
           "20820100300": "GBS 24101011은 통화스왑·헤지 → 20820200200, 이 계정(금리스왑·헤지)은 24101012"}
    for code, t in drv.items():
        if code in row_of:
            r = row_of[code]
            old = ws.cell(r, 18).value
            setc(code, 18, ((old + " / ") if old else "") + "0930 GBS코드 정정: " + t, "현행 파생 코드 규칙(트레이딩·헤지·본지점 × 통화·금리·기타)")

    # 5) 현행 10990003
    r = row_of.get("10490700200")
    if r:
        ws.cell(r, 20).value = ((ws.cell(r, 20).value + "\n") if ws.cell(r, 20).value else "") + "현행 10990003은 2026-07-07 BS에 없음 → 현행 GBS 신설 후 이행(결정 M5)"
        ws.cell(r, 21).value = "결정 M5"

    for r in range(HR + 1, ws.max_row + 1):
        for c in (20, 21):
            x = ws.cell(r, c)
            if x.value:
                x.font = F9; x.alignment = Alignment(wrap_text=True, vertical="top"); x.border = BORDER
    ws.auto_filter.ref = f"A{HR}:U{ws.max_row}"

    # 요약 집계 갱신
    sm = wb["요약"]
    cnt = defaultdict(Counter)
    for r in range(HR + 1, ws.max_row + 1):
        g = ws.cell(r, 1).value
        if g:
            cnt[g][ws.cell(r, 11).value or "유지"] += 1
    for r in range(1, sm.max_row + 1):
        g = sm.cell(r, 1).value
        if g in cnt and isinstance(sm.cell(r, 2).value, int):
            for c, k in enumerate(["신규", "변경", "용도지정", "비활성", "철회", "유지"], 2):
                sm.cell(r, c).value = cnt[g][k]
            sm.cell(r, 8).value = sum(cnt[g].values())
        if g == "합계" and isinstance(sm.cell(r, 2).value, int):
            for c, k in enumerate(["신규", "변경", "용도지정", "비활성", "철회", "유지"], 2):
                sm.cell(r, c).value = sum(cnt[x][k] for x in cnt)
            sm.cell(r, 8).value = sum(sum(v.values()) for v in cnt.values())
    rr = sm.max_row + 2
    sm.cell(rr, 1, f"정정본 {TODAY}").font = Font(name="Arial", bold=True, size=11, color="C55A11")
    for i, t in enumerate([
        "계정상태표 T·U열 = 행별 정정 내용·근거, 정정 셀은 노랑, 신설 행은 주황",
        f"정정 {len(log)}건(정정이력 시트). 현행 8자리 이행: 현행 최하위 {len(C.CUR_LEAF)}개 = 이관 {len(C.FINAL_MAP)}개 + 이행 제외 {len(C.EXCLUDED)}개(06에만 있던 10990003 포함)",
        "잔액 있는 현행 159개 전부 이관, 06 잔액 = 2026-07-07 BS/PL 잔액(차이 0)",
        "1:N 분할 기준은 '이행분할기준', 확정·보류 사항은 '추후고려사항' 시트"], 1):
        sm.cell(rr + i, 1, t).font = Font(name="Arial", size=9)

    # 정정이력
    h = wb.create_sheet("정정이력")
    _table(h, f"정정 이력 ({TODAY}) — 원본 06 대비", ["#", "계정코드", "열", "정정 전", "정정 후", "근거"],
           [(i + 1, a, b, c, d, e) for i, (a, b, c, d, e) in enumerate(log)], [5, 13, 18, 36, 36, 60])
    # 현행 이행 제외
    x = wb.create_sheet("현행이행제외")
    _table(x, "현행 8자리 중 이행하지 않는 계정 (모두 잔액 0)", ["#", "현행 계정", "현행 계정명", "06 연결 대상(삭제)", "사유"],
           [(i + 1, g, n, t, w) for i, (g, n, t, w) in enumerate(sorted(C.EXCLUDED))], [5, 12, 40, 16, 70])
    # 이행 분할 기준
    sp = wb.create_sheet("이행분할기준")
    rows = []
    for g, ts in sorted(C.FINAL_MAP.items()):
        if len(ts) > 1:
            rows.append((g, C.CUR.get(g, {}).get("name", ""), ", ".join(ts), " / ".join(C.acct_name(t) for t in ts), C.SPLIT_RULE.get(g, "속성 기준 분할"),
                         C.CUR.get(g, {}).get("dr", "")))
    _table(sp, "1:N 이행 분할 기준 (현행 1개 → 차세대 여러 계정)", ["현행 계정", "현행 계정명", "대상 11자리", "대상 계정명", "분할 기준", "기준일 잔액"],
           rows, [12, 34, 40, 50, 60, 16])
    # 추후고려사항
    fu = wb.create_sheet("추후고려사항")
    _table(fu, "추후 고려사항 — 추천안 확정 32건 (재검토 시점 도래 시 재판단)", ["#", "현행 계정", "현행 계정명", "적용", "대상 11자리", "대상 계정명", "추후 고려", "재검토 시점", "담당"],
           [(i + 1, u["cur"], u["cur_name"], u["applied"], u["code"], u["name"], u["future"], u["when"], u["owner"]) for i, u in enumerate(C.FUTURE)],
           [5, 12, 34, 11, 13, 30, 60, 22, 18])
    wb.save(out)
    print("saved", out, len(log), "changes", len(C.EXCLUDED), "excluded", len(rows), "splits")


def _table(ws, title, hdr, rows, widths):
    ws.cell(1, 1, title).font = Font(name="Arial", bold=True, size=12, color="1F4E79")
    for c, t in enumerate(hdr, 1):
        x = ws.cell(3, c, t); x.font = Font(name="Arial", bold=True, color="FFFFFF", size=10); x.fill = HDR
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); x.border = BORDER
    for i, row in enumerate(rows, 4):
        for c, v in enumerate(row, 1):
            x = ws.cell(i, c, v); x.font = F9; x.alignment = Alignment(wrap_text=True, vertical="top"); x.border = BORDER
            if isinstance(v, float):
                x.number_format = "#,##0.00"
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(hdr))}{3 + len(rows)}"


if __name__ == "__main__":
    main(sys.argv[1])
