# -*- coding: utf-8 -*-
"""CoA 0930 검토의견 Excel 생성.  python3 build_coa_review.py [출력경로]"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import coa_model as CM
import coa_review_model as M
import coa_migration_model as MG

OUT = sys.argv[1] if len(sys.argv) > 1 else "../docs/05_CoA검토_RBI보고속성.xlsx"

F = "맑은 고딕"
HDR = PatternFill("solid", fgColor="1F3864")
FILL = {
    "신규": "FCE4D6", "이전안수정": "FFF2CC", "변경": "FFF2CC", "속성전환": "DDEBF7", "비활성": "E7E6E6", "매핑": "E2EFDA",
    "철회": "D9D9D9", "수정": "F8CBAD", "신규계정": "FCE4D6", "분할": "FFF2CC", "확인": "FFF2CC", "규칙": "E2EFDA",
    "AUTO 일치": "FFFFFF", "AUTO 확인": "FFF2CC", "미매핑": "E7E6E6", "대응계정": "F2F2F2",
    "신규 수용": "FCE4D6", "0930 기존 계정": "E2EFDA", "병행 후 비활성": "E7E6E6", "범위 확인": "FFF2CC", "신규 수용(선택)": "FCE4D6", "1차 검토 신규": "FCE4D6", "속성": "DDEBF7",
    "충돌": "F8CBAD", "일치": "E2EFDA", "의미주의": "FFF2CC", "0930에 없음": "FCE4D6",
}
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()
wb.remove(wb.active)


def sheet(name, title, note, headers, rows, widths, color_col=None, verify_col=None):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=F, size=14, bold=True)
    ws["A2"] = note
    ws["A2"].font = Font(name=F, size=9, color="595959")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws["A2"].alignment = WRAP
    ws.row_dimensions[2].height = 30
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = Font(name=F, size=10, bold=True, color="FFFFFF")
        c.fill = HDR
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BOX
    for r, row in enumerate(rows, 5):
        fill = None
        if color_col is not None and row[color_col] in FILL:
            fill = PatternFill("solid", fgColor=FILL[row[color_col]])
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = Font(name=F, size=9, bold=(verify_col is not None and i - 1 == verify_col and v == "Y"),
                          color=("C00000" if verify_col is not None and i - 1 == verify_col and v == "Y" else "000000"))
            c.alignment = WRAP
            c.border = BOX
            if fill:
                c.fill = fill
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{4 + len(rows)}"
    return ws


# ---- 판정 결과 ------------------------------------------------------------------
col = M.collisions()
for x in col:
    if x["res"] != "확인":
        x["res"] = "0930에 없음"
    elif x["code"] in M.CAUTION:
        x["res"] = "의미주의"
        x["note"] = M.CAUTION[x["code"]]
    elif x["code"] in M.SAME:
        x["res"] = "일치"
    else:
        x["res"] = "충돌"
nconf = sum(1 for x in col if x["res"] == "충돌")
nmiss = sum(1 for x in col if x["res"] == "0930에 없음")
kinds = {}
for r in M.R:
    kinds[r["kind"]] = kinds.get(r["kind"], 0) + 1

# ---- 이행 매핑 ------------------------------------------------------------------
merr, MIG = MG.validate()
assert not merr, merr
from collections import Counter
mk = Counter(m["kind"] for m in MIG)
nz = [m for m in MIG if m["amt"]]
nz_fix = sum(1 for m in nz if m["kind"] in ("수정", "분할", "신규계정"))
nz_amt_fix = sum(abs(m["amt"]) for m in nz if m["kind"] in ("수정", "분할", "신규계정") and m["code"][0] in "12")
nauto = sum(1 for a in MG.LIVE if a["leaf"] and a["code"] in MG.GBS)
nleaf = sum(1 for a in MG.LIVE if a["leaf"])

# ---- 00 요약 ---------------------------------------------------------------------
summary = [
    ("차수", "2차 (2026-11-05) — 현행 GBS BS/PL(2026-07-07), ECL·신용위험 자본 개발명세 반영. 1차 대비 변경은 '이전안 관계' 열과 '철회' 구분으로 표시"),
    ("결정 반영", "D1 기준코드 = 11자리 통합코드(8자리는 이행 원천). D5 정상자산 충당금 = ECL Directions 2026 Stage별"),
    ("권고 (FN·NRO)", "외국인 비거주자(FN)는 NRO의 부분집합. GL은 NRO 상품별 L5 하나, FN은 nr_subtype 속성으로 부분합. 재무제표·RBI 보고·TDS·송금한도가 같고 GL이 바뀌는 사건은 거주성 변경뿐 (시트 07)"),
    ("이행 매핑", f"현행 인도 GBS leaf 계정 {nleaf}개 중 0930 GBS코드 열로 자동 대응 {nauto}개. 판정: " + " · ".join(f"{k} {v}" for k, v in mk.most_common())
     + f". 잔액 있는 {len(nz)}개 중 {nz_fix}개는 0930 대응을 고쳐야 함 (시트 09, 대차 검증 포함)"),
    ("발견 1 (0930 GBS열 오류)", "0930의 GBS코드 열이 현행 인도 계정을 잘못 가리킴: 저축(20121001)→법인 요구불, 정기(20122001)→법인 정기 전액, RD→AOB, 비거주자 예금→은행 예금, "
     "GST 매출→원천세, TDS 미수→FX SPOT 미수, 비용 미지급→신용카드, 국내송금 미지급→AOB, 법정·기타준비금 뒤바뀜, 투자 Sch.8 분류를 순서대로 대응(주식→외국정부채), 통화·금리스왑 뒤바뀜"),
    ("발견 2 (금액 큼)", "은행 정기예치 ₹208.9 crore(총자산 35%)과 그 이자 ₹6.59 crore·미수 ₹3.74 crore이 0930에서 'CD PURCHASE'(유가증권 성격)에 매핑 → DUE FROM BANKS ON TIME(Sch.7)으로 정정 권고(E1)"),
    ("발견 3 (ECL)", "ECL 명세서가 '미사용'이라 한 10690003~5는 현행 BS에 증권·사모사채·콜론 충당금으로 존재. 부외 ECL 충당부채(2096 그룹)는 0930에 이미 있음(20910000000~20930000000). "
     "11자리로 Stage 1·2·3 L5, 전환조정 자본, EIR 조정만 신설 (시트 10)"),
    ("발견 4 (현행에 이미 있음)", "OD 계정(10610009)·Penal Charges(42020709)·RCM ITC·GST 세목별·TDS 조문별 계정이 현행에 존재 → 예금명세서의 'OD 신규(10630001*)' 요청 불필요, GST·RCM은 11자리 L5로 이어받음"),
    ("1차 의견 수정", "UFCE·국가위험·정상자산(Sch.5) 충당금 신설 제안 철회(ECL 흡수). 일반/특정 충당금은 Stage별로 교체. 0930 GBS열 충돌은 '현행 인도 계정과의 불일치'로 확인됨"),
    ("유지되는 1차 결론", "Sch.3 L5 순도(저축 분리), PSL·BSR·LCR 등은 속성, INR 결제계좌 L5, 에스크로 3개 L5, TDS 신고서별(24Q·26Q·27Q·27EQ) L5 + 조문은 속성(2026.4.1 신법 조문 변경에도 GL 불변)"),
    ("검토의견 건수", " · ".join(f"{k} {v}" for k, v in kinds.items()) + f" / 이번 차수 신규 {len(MG.NEW2)} / 속성 {len(M.ATTRS)}개 / 보고서 {len(M.RETURNS)}종"),
    ("표시", "주황=신규, 노랑=변경·확인, 분홍=0930 대응 수정, 파랑=속성, 회색=비활성·철회·미매핑 / VERIFY=Y(빨강)는 규정 확인 필요"),
    ("다음 단계", "시트 08 결정사항 중 '권고'·'미결' 확정 → coa_model.py 반영, 이행표(시트 09)를 본점 CoA 담당에 전달, ECL·RWA 개발명세의 계정코드를 11자리로 교체"),
]
ws = wb.create_sheet("00 요약")
ws["A1"] = "통합 표준 CoA(0930) 검토의견 — 인도(Noida) 지점 · RBI 보고 관점"
ws["A1"].font = Font(name=F, size=15, bold=True)
ws["A2"] = "작성 기준일 2026-11-05 · 원천: tools/coa_review_model.py (재생성: python3 build_coa_review.py)"
ws["A2"].font = Font(name=F, size=9, color="595959")
for i, (k, v) in enumerate(summary, 4):
    a = ws.cell(row=i, column=1, value=k)
    b = ws.cell(row=i, column=2, value=v)
    a.font = Font(name=F, size=10, bold=True)
    a.fill = PatternFill("solid", fgColor="D9E1F2")
    b.font = Font(name=F, size=10)
    a.alignment = b.alignment = WRAP
    a.border = b.border = BOX
    ws.row_dimensions[i].height = max(18, 15 * (len(v) // 70 + 1))
ws.column_dimensions["A"].width = 16
ws.column_dimensions["B"].width = 120

# ---- 01 판정기준 -------------------------------------------------------------------
sheet("01 판정기준", "GL vs 속성 판정 원칙 및 인도 재무제표 Schedule 대응",
      "GL을 나누는 조건은 P1~P3뿐이다. Schedule 줄이어도 시점마다 재분류되는 구분(P6: PSL·담보)은 속성으로 산출한다.",
      ["ID", "원칙", "내용", "예시", "결론"],
      [p for p in M.PRINCIPLES], [6, 18, 60, 60, 10])
ws = wb["01 판정기준"]
start = 6 + len(M.PRINCIPLES)
ws.cell(row=start, column=1, value="Third Schedule(BR Act s.29) 표시 줄 ↔ CoA").font = Font(name=F, size=12, bold=True)
for i, h in enumerate(["Schedule", "표시 줄", "CoA 대응", "판정", "비고"], 1):
    c = ws.cell(row=start + 1, column=i, value=h)
    c.font = Font(name=F, size=10, bold=True, color="FFFFFF")
    c.fill = HDR
    c.border = BOX
for r, row in enumerate(M.SCHEDULES, start + 2):
    for i, v in enumerate(row, 1):
        c = ws.cell(row=r, column=i, value=v)
        c.font = Font(name=F, size=9)
        c.alignment = WRAP
        c.border = BOX
        if row[3].startswith("속성"):
            c.fill = PatternFill("solid", fgColor=FILL["속성전환"])

# ---- 02 GBS코드 충돌 ---------------------------------------------------------------
rows = [(x["code"], x["meaning"], x["st"], x["coa"] or "-", x["res"], x.get("note", "")) for x in col]
rows.sort(key=lambda r: ({"충돌": 0, "의미주의": 1, "0930에 없음": 2, "일치": 3}[r[4]], r[0]))
sheet("02 GBS코드충돌", "예금명세서 8자리 코드 ↔ 0930 CoA 'GBS코드' 열 대조 (1차, 자동 비교)",
      f"충돌 {nconf}건: 같은 8자리 코드가 0930에서는 다른 계정. 0930에 없음 {nmiss}건: 예금명세서 신규(*) 또는 0930 미채번. "
      "(2차) 현행 BS/PL로 확인한 결과 예금명세서 코드가 현행 인도 계정이 맞고, 0930 GBS열이 잘못 대응한 것 → 전체 이행표는 시트 09.",
      ["예금명세서 코드", "예금명세서 의미", "명세서 상태", "0930 CoA에서 같은 GBS코드를 가진 계정", "판정", "비고"],
      rows, [12, 40, 10, 60, 12, 60], color_col=4)

# ---- 03 계정 검토의견 --------------------------------------------------------------
order = {"신규": 1, "이전안수정": 2, "변경": 3, "매핑": 4, "속성전환": 5, "비활성": 6}
rows = []
for r in M.R:
    cur = CM.ALL.get(r["code"])
    rows.append((r["kind"], r["code"], r["lv"], r["parent"], cur["name"].strip() if cur and r["kind"] != "신규" else "",
                 r["en"], r["kr"], r["why"], r["rep"], r["prev"], r["verify"]))
sheet("03 계정검토의견", "계정 변경·추가·속성전환 의견 (인도지점 적용분)",
      "신규 코드는 본점 채번 전 임시코드(기존 코드·이전안과 중복 없음 검증). '이전안'은 2026-11-01 반영분(coa_model.py). "
      "속성전환은 GL을 만들지 않고 시트 05 속성으로 관리하자는 의견.",
      ["구분", "코드", "Lv", "상위", "현재 명칭(0930/이전안)", "제안 영문명", "한글명", "사유·근거", "관련 보고", "이전안 관계", "VERIFY"],
      rows, [10, 13, 4, 13, 34, 44, 24, 70, 24, 22, 7], color_col=0, verify_col=10)

# ---- 04 예금상품 매핑 -------------------------------------------------------------
sheet("04 예금상품매핑", "예금명세서 상품 → 11자리 계정 매핑 (제안)",
      "원금 GL = 상품 × 거주성 × 고객유형군(통합 L4). 이자비용은 통합 L4(고객유형별)에 기장하고 상품·거주성은 속성 → 명세서의 51040011* 등 신규 이자 GL 불필요.",
      ["상품", "상품명", "명세서 원금 GL(8자리)", "제안 원금 GL(11자리)", "이자비용 GL", "미지급이자 GL", "TDS GL", "함께 관리할 속성"],
      [p for p in M.PRODUCT_MAP], [14, 26, 22, 34, 26, 20, 14, 34])

# ---- 05 RBI 보고 속성 -------------------------------------------------------------
sheet("05 RBI보고속성", "회계계정이 아닌 속성으로 유지해야 할 상세 단계 (RBI·FEMA·세무 보고용)",
      "수준 = 값을 보관하는 단위. '현재 설계 반영'은 예금명세서 시트·TF모델·송금모델에 이미 있는지. 신규는 GBS 고객·계좌 마스터 확장 필요.",
      ["영역", "속성(컬럼)", "값·형식", "수준", "사용 보고서·용도", "근거", "현재 설계 반영", "VERIFY"],
      [a for a in M.ATTRS], [12, 34, 50, 10, 46, 24, 22, 7], verify_col=7)

# ---- 06 보고서별 요구 -------------------------------------------------------------
sheet("06 보고서별요구", "RBI 보고서별 출처 계정과 필요한 속성",
      "주기는 현행 일반 기준이며 VERIFY=Y 항목은 CIMS 리턴 명세로 확인. 산출 방식 'GL × 속성'은 보고 시점에 GL 잔액을 속성으로 나눈다는 뜻.",
      ["보고서", "주기", "출처 계정", "필요한 속성", "산출 방식", "VERIFY"],
      [r for r in M.RETURNS], [30, 10, 40, 60, 50, 7], verify_col=5)

# ---- 07 주제별 상세 ---------------------------------------------------------------
sheet("07 FN·PSL·Nostro", "주제별 상세 — 외국인 비거주자·NRO / PSL 여신 / INR Nostro",
      "FN·NRO는 2차 권고. PSL·INR Nostro는 1차 요청 예시에 대한 판단.",
      ["주제", "항목", "내용"], [t for t in M.TOPICS], [12, 34, 120])

# ---- 08 결정 필요 ----------------------------------------------------------------
sheet("08 결정사항", "본점·재무·컴플라이언스 결정 필요 사항",
      "결정 후 coa_model.py에 반영하고 분개 예시·DDL·설계서를 재생성.",
      ["ID", "결정 사항", "선택지", "권고·결정", "상태"], [d for d in M.DECISIONS], [6, 26, 50, 70, 16])

# ---- 09 이행 매핑 ----------------------------------------------------------------
order9 = {"수정": 0, "분할": 1, "신규계정": 2, "확인": 3, "AUTO 확인": 4, "규칙": 5, "미매핑": 6, "비활성": 7, "AUTO 일치": 8, "대응계정": 9}
rows = sorted(MIG, key=lambda m: (0 if m["amt"] else 1, order9.get(m["kind"], 9), m["code"]))
ws = sheet("09 현행GBS이행매핑", "현행 인도 GBS(8자리, 2026-07-07) → 11자리 통합코드 이행 매핑",
      "leaf 계정 전체. 정렬: 잔액 있는 계정 → 판정(수정·분할·신규·확인…) 순. '0930 자동대응'은 0930 GBS코드 열이 가리키는 계정, '이행 대상'이 권고. "
      "분할 = 고객유형(법인/개인)별로 나눠 이행. 대응계정 = 현행 난외 대(7xxx)는 0930에서 시스템 생성.",
      ["구분", "현행코드", "현행 계정명", "잔액(INR)", "0930 자동대응(GBS코드 열)", "이행 대상(11자리)", "판정", "비고"],
      [({"ASSET": "BS-자산", "LIAB": "BS-부채·자본", "EXP": "PL-비용", "INC": "PL-수익"}[m["side"]], m["code"], m["name"], m["amt"] or None,
        m["auto"], m["tgt"], m["kind"], m["note"]) for m in rows],
      [11, 10, 40, 16, 46, 26, 10, 70], color_col=6)
for r in range(5, 5 + len(rows)):
    ws.cell(row=r, column=4).number_format = '#,##0;[Red]-#,##0'
# 대차 검증: 이행 대상 첫 번째 계정의 첫 자리(1 자산, 2 부채, 3 자본, 4 수익, 5 비용) 기준, 차변 +
sign = {"ASSET": 1, "EXP": 1, "LIAB": -1, "INC": -1}
cls = {}
for m in MIG:
    if m["kind"] == "대응계정" or not m["amt"]:
        continue
    t = MG.targets_of(m["tgt"])
    if not t:
        continue
    k = t[0][0]
    cls[k] = cls.get(k, 0) + sign[m["side"]] * m["amt"]
base = len(rows) + 7
ws.cell(row=base, column=1, value="대차 검증 (이행 대상 계정 기준, 차변 +)").font = Font(name=F, size=11, bold=True)
labels = {"1": "자산(1)", "2": "부채(2)", "3": "자본(3)", "4": "수익(4)", "5": "비용(5)", "6": "난외(6)"}
for i, k in enumerate(sorted(cls)):
    ws.cell(row=base + 1 + i, column=1, value=labels.get(k, k))
    c = ws.cell(row=base + 1 + i, column=4, value=round(cls[k], 2))
    c.number_format = '#,##0.00;[Red]-#,##0.00'
n = len(cls)
ws.cell(row=base + 1 + n, column=1, value="B/S 합계(1+2+3) = 0 이어야 함").font = Font(name=F, bold=True)
ws.cell(row=base + 1 + n, column=4, value=f"=SUM(D{base+1}:D{base+3})").number_format = '#,##0.00'
ws.cell(row=base + 2 + n, column=1, value="P/L 합계(4+5) = 0 이어야 함").font = Font(name=F, bold=True)
ws.cell(row=base + 2 + n, column=4, value=f"=D{base+4}+D{base+5}").number_format = '#,##0.00'
assert abs(cls["1"] + cls["2"] + cls["3"]) < 0.01 and abs(cls["4"] + cls["5"]) < 0.01, cls

# ---- 10 ECL 계정 -----------------------------------------------------------------
ws = sheet("10 ECL계정", "ECL·신용위험 자본 개발명세의 계정 요청 → 11자리",
      "ECL Directions 2026(RBI/DOR/2026-27/398)·SA Directions 2026(397), 2027-04-01 시행. RWA 엔진은 분개가 없고 속성만 필요(시트 05 'ECL·자본').",
      ["명세 코드", "명세 계정", "11자리", "판정", "비고"], [e for e in MG.ECL_MAP], [18, 40, 36, 14, 80], color_col=3)
st = len(MG.ECL_MAP) + 7
ws.cell(row=st, column=1, value="ECL 분개 (11자리)").font = Font(name=F, size=12, bold=True)
for i, h in enumerate(["ID", "사건", "차변", "대변", "비고"], 1):
    c = ws.cell(row=st + 1, column=i, value=h)
    c.font = Font(name=F, size=10, bold=True, color="FFFFFF"); c.fill = HDR; c.border = BOX
for r, row in enumerate(MG.ECL_JE, st + 2):
    for i, v in enumerate(row, 1):
        c = ws.cell(row=r, column=i, value=v); c.font = Font(name=F, size=9); c.alignment = WRAP; c.border = BOX
st2 = st + 3 + len(MG.ECL_JE)
ws.cell(row=st2, column=1, value="이번 차수 신규 계정").font = Font(name=F, size=12, bold=True)
for i, h in enumerate(["코드", "Lv", "상위", "영문명", "한글명·사유"], 1):
    c = ws.cell(row=st2 + 1, column=i, value=h)
    c.font = Font(name=F, size=10, bold=True, color="FFFFFF"); c.fill = HDR; c.border = BOX
for r, n in enumerate(MG.NEW2, st2 + 2):
    for i, v in enumerate([n[0], n[1], n[2], n[3], f"{n[4]} — {n[5]}"], 1):
        c = ws.cell(row=r, column=i, value=v); c.font = Font(name=F, size=9); c.alignment = WRAP; c.border = BOX
        c.fill = PatternFill("solid", fgColor=FILL["신규"])

wb.move_sheet("08 결정사항", offset=2)
wb.save(OUT)
print("saved", OUT)
