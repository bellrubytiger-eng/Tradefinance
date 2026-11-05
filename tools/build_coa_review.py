# -*- coding: utf-8 -*-
"""CoA 0930 검토의견 Excel 생성.  python3 build_coa_review.py [출력경로]"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import coa_model as CM
import coa_review_model as M

OUT = sys.argv[1] if len(sys.argv) > 1 else "../docs/05_CoA검토_RBI보고속성.xlsx"

F = "맑은 고딕"
HDR = PatternFill("solid", fgColor="1F3864")
FILL = {
    "신규": "FCE4D6", "이전안수정": "FFF2CC", "변경": "FFF2CC", "속성전환": "DDEBF7", "비활성": "E7E6E6", "매핑": "E2EFDA",
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

# ---- 00 요약 ---------------------------------------------------------------------
summary = [
    ("검토 대상", "통합 표준 CoA 0930(1,168계정) + 인도지점 이전 반영안(신규 43·변경 23·비활성 16) / GBS Deposit Product Spec ACPIR2026 v3"),
    ("핵심 원칙", "GL은 재무제표 Schedule 줄·대사 단위가 다를 때만 분리. PSL·BSR·CRILC·LCR 등 RBI 보고 세분은 속성으로 관리하고 GL × 속성으로 산출 (시트 01)"),
    ("발견 1 (긴급)", f"예금명세서가 쓰는 '현행 GBS 8자리 코드' 중 {nconf}개가 0930 CoA GBS코드 열에서 다른 계정을 가리킴 "
                    f"(예: 20124001 RD ↔ AOB 예금, 51040001 저축이자 ↔ 법인 요구불이자, 20990013 GST ↔ 원천세). {nmiss}개는 0930에 없음 (시트 02)"),
    ("발견 2", "0930 예금은 고객유형(법인·은행·개인) 축이고 저축(Savings) 구분이 없음. Sch.3·Form A는 요구불/저축/정기 + 은행분/기타분을 요구 → L5 순도 확보 필요. "
              "이전안 20150300100 'SB/CA'는 저축·당좌 혼합이므로 정정"),
    ("발견 3", "예금명세서 매핑 오류: LC → 72040001 'Committed Credit Line'(→ 60120000000이 맞음), CC/OD 제안 10630001*은 신디케이트 대출 코드와 충돌, "
              "에스크로 수수료 42020729는 0930에서 대출 기타수수료"),
    ("발견 4", "0930에 인도 필수 계정 부재: CC/OD 자산, Packing Credit(PCFC), RBI 기타계좌, PSL 미달예치·PSLC, DEAF 청구, 법정준비금(자본)·IFR, GST RCM, TDS 비거주자"),
    ("발견 5", "휴면계좌를 부채(20510100000)로 이관하는 통합 CoA 방식은 인도와 맞지 않음(휴면도 예금, 10년 후 DEAF). 법정준비금이 부채측에 있음"),
    ("PSL", "GL 분리 없이 psl_flag·psl_category·msme_category 등 속성. Sch.9 C 우선부문 줄도 속성에서 산출. 실제 자금·손익(미달예치, PSLC 수수료)만 GL (시트 07)"),
    ("INR Nostro", "인도 내 INR 결제계좌는 L5(10105010100)로 FCY Nostro와 분리(Form A 은행시스템 자산·NOP 제외). 해외 Nostro는 통화차원 + nostro_type·국가 속성(LBS). "
                   "본점·당행 해외지점의 INR 계좌는 예금이 아닌 본지점 (시트 07)"),
    ("이전안 재검토", "속성 원칙에 맞춰 금융/이행보증 L5·내국LC L4는 속성 전환 권고, SNRR 2개 L5는 통합(2026.6 FEMA 개정), 에스크로는 3개 L5로 세분, TDS는 26Q/27Q 분리"),
    ("검토의견 건수", " · ".join(f"{k} {v}" for k, v in kinds.items()) + f" / 속성 {len(M.ATTRS)}개 / 보고서 {len(M.RETURNS)}종"),
    ("표시", "주황=신규, 노랑=변경·이전안수정, 파랑=속성전환, 회색=비활성, 초록=매핑 / VERIFY=Y(빨강)는 규정 조문·주기 확인 필요"),
    ("다음 단계", "시트 08 결정사항(D1~D7) 확정 후 tools/coa_model.py(ADD/MOD/DEACT)·분개 예시·DDL에 반영"),
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
sheet("02 GBS코드충돌", "예금명세서 '현행 GBS 8자리 코드' ↔ 0930 CoA 'GBS코드' 열 대조 (자동 비교)",
      f"충돌 {nconf}건: 같은 8자리 코드가 0930에서는 다른 계정. 0930에 없음 {nmiss}건: 예금명세서 신규(*) 또는 0930 미채번. "
      "→ 11자리 통합코드를 기준키로 하고 8자리는 본점 재채번 후 한 번만 매핑(결정 D1).",
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
sheet("07 PSL·INR Nostro", "주제별 상세 — PSL 여신 / INR Nostro",
      "요청 예시 두 가지(PSL은 계정이 아닌 속성, INR Nostro는 구분 필요)에 대한 상세 판단.",
      ["주제", "항목", "내용"], [t for t in M.TOPICS], [12, 34, 120])

# ---- 08 결정 필요 ----------------------------------------------------------------
sheet("08 결정사항", "본점·재무·컴플라이언스 결정 필요 사항",
      "결정 후 coa_model.py에 반영하고 분개 예시·DDL·설계서를 재생성.",
      ["ID", "결정 사항", "선택지", "권고", "결정 주체"], [d for d in M.DECISIONS], [6, 26, 50, 60, 22])

wb.save(OUT)
print("saved", OUT)
