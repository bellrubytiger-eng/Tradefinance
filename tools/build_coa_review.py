# -*- coding: utf-8 -*-
"""CoA 0930 검토의견 Excel 생성.  python3 build_coa_review.py [출력경로]"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import coa_model as CM
import coa_review_model as M
import coa_migration_model as MG
import lcr_model as L
import tax_fa_model as TF
import sheet_doc_model as SD
import re as _re

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
    ("차수", "7차 (2026-11-05) — 6단계 재분류, 중계은행 회수 과세, 시트별 변경 근거(FY25-26 결산 증거) / 6차 — 연말 미경과 이자 TDS 분개, 송금·무역 수수료 구분 / 5차 — 5단계 체계 반영(L6 → L5 또는 파라미터), TCS 기타·협회비·CSR 부채·직원 월적립 부채 / 4차 — 적금 L3 신설 요청, TDS·TCS·GST 11자리(신 소득세법·GST FRS), 고정자산(Schedule II)·영업비용(Sch.16) 헤드 / 3차 — LCR 분류·파라미터, TDS 미수, 은행 예치 구분 추가 / 2차 — 현행 GBS BS/PL(2026-07-07), ECL·신용위험 자본 개발명세 반영. 1차 대비 변경은 '이전안 관계' 열과 '철회' 구분으로 표시"),
    ("결정 반영", "D1 기준코드 = 11자리 통합코드(8자리는 이행 원천). D5 정상자산 충당금 = ECL Directions 2026 Stage별"),
    ("권고 (FN·NRO)", "외국인 비거주자(FN)는 NRO의 부분집합. GL은 NRO 상품별 L5 하나, FN은 nr_subtype 속성으로 부분합. 재무제표·RBI 보고·TDS·송금한도가 같고 GL이 바뀌는 사건은 거주성 변경뿐 (시트 07)"),
    ("이행 매핑", f"현행 인도 GBS leaf 계정 {nleaf}개 중 0930 GBS코드 열로 자동 대응 {nauto}개. 판정: " + " · ".join(f"{k} {v}" for k, v in mk.most_common())
     + f". 잔액 있는 {len(nz)}개 중 {nz_fix}개는 0930 대응을 고쳐야 함 (시트 09, 대차 검증 포함)"),
    ("발견 1 (0930 GBS열 오류)", "0930 GBS코드 열이 현행 인도 계정을 엉뚱한 계정에 대응시킨 사례(✗ = 0930이 가리킨 계정, → = 권고 이행 대상): "
     "개인 저축 20121001 ✗법인 요구불 → 개인 저축 L5 20150300100 / RD 20124001 ✗AOB(NH올원뱅크) 예금 → RD L5 20160300600 / "
     "GST 매출 ✗원천세 → GST 세목 L5 / TDS 미수 ✗FX SPOT 미수(자금) → TDS 미수 10810200000 / Legal Reserve ✗기타준비금 → Statutory Reserve 31110300000 / "
     "비용 미지급 ✗신용카드, 국내송금 미지급 ✗AOB / 투자 Sch.8 분류 순서 대응(주식 ✗외국정부채) / 통화·금리스왑 뒤바뀜"),
    ("발견 2 (금액 큼)", "은행 정기예치 ₹208.9 crore(총자산 35%)과 그 이자 ₹6.59 crore·미수 ₹3.74 crore이 0930에서 'CD PURCHASE'(유가증권 성격)에 매핑 → DUE FROM BANKS ON TIME(Sch.7)으로 정정 권고(E1)"),
    ("발견 3 (ECL)", "ECL 명세서가 '미사용'이라 한 10690003~5는 현행 BS에 증권·사모사채·콜론 충당금으로 존재. 부외 ECL 충당부채(2096 그룹)는 0930에 이미 있음(20910000000~20930000000). "
     "11자리로 Stage 1·2·3 L5, 전환조정 자본, EIR 조정만 신설 (시트 10)"),
    ("발견 4 (현행에 이미 있음)", "OD 계정(10610009)·Penal Charges(42020709)·RCM ITC·GST 세목별·TDS 조문별 계정이 현행에 존재 → 예금명세서의 'OD 신규(10630001*)' 요청 불필요, GST·RCM은 11자리 L5로 이어받음"),
    ("1차 의견 수정", "UFCE·국가위험·정상자산(Sch.5) 충당금 신설 제안 철회(ECL 흡수). 일반/특정 충당금은 Stage별로 교체. 0930 GBS열 충돌은 '현행 인도 계정과의 불일치'로 확인됨"),
    ("유지되는 1차 결론", "Sch.3 L5 순도(저축 분리), PSL·BSR·LCR 등은 속성, INR 결제계좌 L5, 에스크로 3개 L5, TDS 신고서별(24Q·26Q·27Q·27EQ) L5 + 조문은 속성(2026.4.1 신법 조문 변경에도 GL 불변)"),
    ("검토의견 건수", " · ".join(f"{k} {v}" for k, v in kinds.items()) + f" / 이번 차수 신규 {len(MG.NEW2)} / 속성 {len(M.ATTRS)}개 / 보고서 {len(M.RETURNS)}종"),
    ("표시", "주황=신규, 노랑=변경·확인, 분홍=0930 대응 수정, 파랑=속성, 회색=비활성·철회·미매핑 / VERIFY=Y(빨강)는 규정 확인 필요"),
    ("3차 (LCR·정정)", "은행 정기예치는 은행/비은행만 GL(10106010000), 상대은행 유형·만기·운영 여부는 속성. LCR 분류(개인·소규모사업자·비금융법인·기타법적실체·금융기관·본점)는 "
     "GL이 아닌 계좌 속성 + 일자 유효 파라미터(TB_LCR_*) (시트 11). TDS 미수는 10810200000 신설(선납법인세와 분리). 법정준비금 = Statutory Reserve(31110300000)"),
    ("4차 (적금·AOB)", "적금은 0930에 구분이 없음 → 정기예금과 동등한 L3 20190000000(+이자 50170000000, 미지급이자 20410120000) 신설 요청. 3차까지의 '정기예금 아래 L5' 제안은 철회. "
     "AOB 의미는 본점 확인 요청(0930 본점 매핑명은 올원뱅크) — 인도 미사용은 동일"),
    ("4차 (세무)", "TDS·TCS: 챌린 납부 성격 L5(현행 조문별 계정 30여 개를 1:1·n:1 이행) — 신법(392·393·394) 조문 변경에도 GL 불변. TCS: LRS 교육·의료/기타/해외여행/기타 L5. "
     "GST: 세목 L5(수수료/환전은 파라미터), RCM·ITC·임시 ITC·전자현금원장·이자가산세 신설. GST FRS 신규안 20990040은 현행 TCS 계정과 충돌 (시트 12)"),
    ("4차 (고정자산·비용)", "표시는 BR Act Third Schedule(Sch.10·16), 인식은 인도 AS(은행 Ind AS 미적용), 내용연수는 Companies Act Schedule II, 세무는 감가상각 블록. "
     "컴퓨터 단말·서버·전기설비·가구비품 계정 신설, FA 코드에 GL·내용연수·세무율 연결(시트 13). 감사보수·법률비용이 Service Commission에 섞여 Sch.16 VII·VIII을 GL로 못 냄 → L5 분리, "
     "복리비는 Sch.16 I, Service Commission ₹2.62 crore는 총수익 1% 초과 별도 공시 (시트 14)"),
    ("7차 (6단계·근거)", "6단계 사용 확인 → 세금은 L6(TDS·TCS 납부 성격 21개, GST 수수료/환전), 감사·법률·전문가 수수료는 SERVICE FEES 아래 L6, 0930 L6 36개도 기장 계정(5차 의견 철회). "
     "중계은행 수수료 고객 회수분 = 과세 수익 40570110000(C1 결정, pure agent 통과계정 철회). 모든 시트 상단에 '본점 원안 / 우리측 변경 / 이유 / FY25-26 결산 증거', "
     "계정 표에 '0930 원안' 열 추가, 결산 증거 20건을 시트 16에 정리"),
    ("6차 (TDS·수수료)", "3/31 미경과 예금이자 TDS: 이자비용 차변, 미지급이자(순액)·TDS 예금이자 대변 → 4/30까지 RBI 계좌로 납부. 신규 계정 불필요. "
     "현행 'Prepaid TDS on Accrued Interest'(10920014)는 고객 몫 선급 TDS → 4차의 'TDS 미수' 매핑을 정정, 순액법 전환 시 미지급이자와 상계(시트 12 ④). "
     "수수료: 취급·SWIFT 전신료는 0930 상품별 L4 사용, 중계은행 실비는 pure agent 통과계정 20441100000 또는 과세 수익 40570110000 (시트 15)"),
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

# ---- 11 LCR -----------------------------------------------------------------------
def section(ws, start, title, headers, rows, color_col=None, verify_col=None):
    ws.cell(row=start, column=1, value=title).font = Font(name=F, size=12, bold=True)
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=start + 1, column=i, value=h)
        c.font = Font(name=F, size=10, bold=True, color="FFFFFF"); c.fill = HDR; c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for r, row in enumerate(rows, start + 2):
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            red = verify_col is not None and i - 1 == verify_col and v == "Y"
            c.font = Font(name=F, size=9, bold=red, color="C00000" if red else "000000")
            c.alignment = WRAP; c.border = BOX
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                c.number_format = '#,##0'
            if color_col is not None and str(row[color_col]).startswith("GL"):
                c.fill = PatternFill("solid", fgColor=FILL["신규"])
            elif color_col is not None and row[color_col] == "속성":
                c.fill = PatternFill("solid", fgColor=FILL["속성전환"])
    return start + len(rows) + 3


ws = wb.create_sheet("11 LCR분류·파라미터")
ws["A1"] = "LCR 대비 — 은행/고객 구분, RBI 상대방 분류, 유출·유입률 파라미터"
ws["A1"].font = Font(name=F, size=14, bold=True)
ws["A2"] = ("계정은 재무제표 줄(은행분/기타분, 은행 예치/기타)까지만 나누고, 개인·소규모사업자·비금융법인·금융기관 및 안정/불안정·운영예금은 고객·계좌 속성으로 "
            "보고일마다 산출. 비율은 일자 유효 파라미터(TB_LCR_*)로 관리 — VERIFY=Y 값은 현행 RBI Master Direction으로 확정 후 적재.")
ws["A2"].font = Font(name=F, size=9, color="595959"); ws["A2"].alignment = WRAP
ws.merge_cells("A2:H2"); ws.row_dimensions[2].height = 42
r = 4
r = section(ws, r, "① 계정 vs 속성", ["대상", "판정", "계정·속성", "이유"], L.GL_DECISION, color_col=1)
r = section(ws, r, "② RBI LCR 상대방 분류 (유출)", ["코드", "대분류", "정의", "판정 규칙(속성)", "유출률", "디지털 가산", "비고", "VERIFY"], L.CPTY_CLASS, verify_col=7)
r = section(ws, r, "③ 고객 구성 → 기본 분류 (TB_LCR_CPTY_MAP 초기값)", ["constitution", "기본 분류", "비고", "VERIFY"], L.CONSTITUTION_MAP, verify_col=3)
r = section(ws, r, "④ 유입률", ["대상", "계정", "유입률", "비고", "VERIFY"], L.INFLOW, verify_col=4)
r = section(ws, r, "⑤ 난외 유출", ["대상", "계정", "유출률", "비고", "VERIFY"], L.OFFBS, verify_col=4)
r = section(ws, r, "⑥ HQLA", ["항목", "계정", "수준", "헤어컷·한도", "비고", "VERIFY"], L.HQLA, verify_col=5)
r = section(ws, r, "⑦ 파라미터 (TB_LCR_THRESHOLD)", ["코드", "의미", "초기값", "단위", "근거", "VERIFY"], L.PARAMS, verify_col=5)
r = section(ws, r, "⑧ 파라미터·스냅샷 테이블", ["테이블", "키", "주요 컬럼", "용도"], L.MASTERS)
r = section(ws, r, "⑨ 현행 잔액(2026-07-07) 적용 예", ["항목", "계정(현행 → 11자리)", "잔액(INR)", "LCR 처리", "추가로 필요한 데이터"], L.CURRENT)
for i, w in enumerate([30, 22, 40, 50, 12, 16, 30, 8], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---- 12 세무 계정 --------------------------------------------------------------------
NEWALL = {n[0]: n for n in TF.NEW3}
CUR_NAME = {c: CM.ALL[c].get("br_name") or CM.ALL[c]["name"].strip() for c in CM.ALL}
CUR_NAME.update({n[0]: n[3] for n in MG.NEW2})
CUR_NAME.update({r["code"]: r["en"] for r in M.R if r["code"] != "-" and r["kind"] not in ("철회", "비활성")})
CUR_NAME.update({n[0]: n[3] for n in TF.NEW3})
NEW_ANY = {n[0] for n in TF.NEW3} | {n[0] for n in MG.NEW2} | {r["code"] for r in M.R if r["kind"] == "신규"}


_BASE_BY = {a["code"]: a for a in CM.BASE}


def ho_view(text):
    """문자열 속 11자리 코드마다 0930 원안(명칭·본점 매핑·GBS코드) 또는 신설 표시."""
    out = []
    for c in dict.fromkeys(_re.findall(r"\b\d{11}\b", str(text))):
        if c in _BASE_BY:
            a = _BASE_BY[c]
            out.append(f'{c} {a["name"].strip()} [{a["ho_name"] or "-"} / GBS {a["gbs"] or "-"}]')
        elif c in NEW_ANY:
            anc = next((c[:k].ljust(11, "0") for k in (9, 7, 5, 3) if c[:k].ljust(11, "0") in _BASE_BY and c[:k].ljust(11, "0") != c), None)
            if anc:
                a = _BASE_BY[anc]
                out.append(f'{c} 0930에 없음(신설) — 0930 상위 {anc} {a["name"].strip()} [{a["ho_name"] or "-"} / GBS {a["gbs"] or "-"}]')
            else:
                out.append(f"{c} 0930에 없음(신설)")
    return "\n".join(out)


ws = wb.create_sheet("12 세무계정(TDS·TCS·GST)")
ws["A1"] = "TDS·TCS·GST 계정 — 11자리 (소득세법 2025, GST 2.0)"
ws["A1"].font = Font(name=F, size=14, bold=True)
ws["A2"] = "6단계: TDS·TCS L5 = 신고서, L6 = 챌린 납부 성격 / GST L5 = 세목, L6 = 수수료·환전. 조문(구/신)·세율·한도·고령자 여부는 파라미터(③) — TDS 산출물 Rate Master와 GST Charge Code 마스터(일자 유효)에서 관리하고 GL 이름에 넣지 않음."
ws["A2"].font = Font(name=F, size=9, color="595959"); ws["A2"].alignment = WRAP; ws.merge_cells("A2:F2"); ws.row_dimensions[2].height = 30
r = 4
r = section(ws, r, "① 개정 대응·설계 판단", ["항목", "내용"], TF.TAX_NOTES)
tax_live = [m for m in MIG if m["code"].startswith(("2099", "1099", "1092001", "1099100")) and (m["tgt"].startswith(("2051", "1042", "1081", "1049")))]
rows = []
for m in sorted(tax_live, key=lambda m: m["tgt"]):
    t = MG.targets_of(m["tgt"])
    nm = " / ".join(CUR_NAME.get(x, x) for x in t)
    rows.append((m["code"], m["name"], m["amt"] or None, m["tgt"], nm, ho_view(m["tgt"]), m["kind"], m["note"]))
r = section(ws, r, "② 현행 세무 계정 → 11자리", ["현행코드", "현행 계정명", "잔액(INR)", "11자리", "11자리 계정명", "0930 원안 [본점매핑 / GBS]", "판정", "비고"], rows)
r = section(ws, r, "③ 세금 L5 ↔ 파라미터 (TAX_GL_PARAM, 일자 유효)", ["L5", "세목", "신고서", "구 조문", "신 조문(2025법)", "파라미터로 관리할 하위 구분"], TF.TAX_GL_PARAM)
r = section(ws, r, "④ 연말(3/31) 미경과 예금이자 TDS — 순액법 (예: 거주자 FD 미경과 이자 ₹60,000, TDS 10%)",
            ["ID", "시점", "차변", "대변", "비고"], TF.YE_TDS)
r = section(ws, r, "⑤ 신설 세무 계정", ["코드", "Lv", "상위", "영문명", "한글명", "사유"],
            [n[:6] for n in TF.NEW3 if n[6] == "세무"])
for i, w in enumerate([14, 36, 16, 40, 44, 12, 60], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---- 13 고정자산 ----------------------------------------------------------------------
ws = wb.create_sheet("13 고정자산")
ws["A1"] = "고정자산 — Sch.10 · Schedule II 내용연수 · 세무 블록 · FA 코드 연결"
ws["A1"].font = Font(name=F, size=14, bold=True)
ws["A2"] = "현행 FA 코드표는 GL(Fixed Asset·BS·PL·처분손익) 열이 비어 있고 내용연수·상각률이 0. 아래 분류표로 채우고 1계열(1100~1901)은 3계열(3100~3902, Schedule II 분류)로 통합 후 폐기 권고."
ws["A2"].font = Font(name=F, size=9, color="595959"); ws["A2"].alignment = WRAP; ws.merge_cells("A2:J2"); ws.row_dimensions[2].height = 30
r = 4
r = section(ws, r, "① 자산 분류 → 계정", ["분류키", "분류", "자산 GL", "감가상각누계 GL", "상각비 GL", "Sch.10 줄", "Schedule II 내용연수", "세무 블록률(WDV)", "FA 코드", "비고"], TF.FA_CLASS)
cls_of = {}
for c in TF.FA_CLASS:
    for code in c[8].replace("·", " ").split():
        if code.isdigit():
            cls_of[code] = c
fa_rows = []
for f in TF.FA_LIST:
    c = cls_of.get(f["code"][:4]) or cls_of.get(f["code"])
    if c:
        fa_rows.append((f["code"], f["name"], c[1], c[2], c[3], c[4], "40770000000 / 51209000000" if c[0] not in ("LAND", "LOWVAL", "CWIP") else "-", c[6], c[7]))
    else:
        fa_rows.append((f["code"], f["name"], "(그룹 헤더)", "", "", "", "", "", ""))
r = section(ws, r, "② FA 코드표 채우기 (Fixed Asset / BS / PL / 처분손익 / 내용연수 / 세무율)",
            ["FA 코드", "명칭", "분류", "Fixed Asset A/C", "BS A/C(누계)", "PL A/C(상각비)", "처분 이익/손실", "내용연수", "세무율"], fa_rows)
r = section(ws, r, "③ 파라미터", ["코드", "의미", "값", "근거"], TF.FA_PARAMS)
fa_mig = [m for m in MIG if m["code"][:3] in ("121", "122", "123") or m["code"] in ("52010002", "52010004", "52010005", "52010006", "48010005", "53090004")]
r = section(ws, r, "④ 현행 고정자산 계정 → 11자리", ["현행코드", "현행 계정명", "잔액(INR)", "11자리", "판정", "비고"],
            [(m["code"], m["name"], m["amt"] or None, m["tgt"], m["kind"], m["note"]) for m in fa_mig if m["amt"] or m["kind"] in ("신규계정", "분할", "수정")])
for i, w in enumerate([12, 34, 22, 24, 24, 24, 26, 22, 30, 40], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---- 14 영업비용 · Sch.16 ------------------------------------------------------------
ws = wb.create_sheet("14 영업비용·Sch16")
ws["A1"] = "영업비용 — Third Schedule Sch.16 줄 매핑 · 경비(EP) 코드 연결"
ws["A1"].font = Font(name=F, size=14, bold=True)
ws["A2"] = ("Sch.16 줄은 0930 L4 그룹과 다르므로 계정마다 줄을 지정(속성). 감사보수(VII)·법률비용(VIII)처럼 Sch.16 별도 줄은 GL로 분리. "
            "RBI 공시: Sch.14·16의 '기타' 항목 중 총수익 1% 초과는 별도 공시(VERIFY). 2026-07-07 총수익 ₹23.3 crore → 1% = ₹23 lakh")
ws["A2"].font = Font(name=F, size=9, color="595959"); ws["A2"].alignment = WRAP; ws.merge_cells("A2:H2"); ws.row_dimensions[2].height = 42
r = 4
r = section(ws, r, "① Sch.16 줄", ["줄", "명칭"], list(TF.SCH16_LINES.items()))
mig_by = {m["code"]: m for m in MIG}
_lv = {a["code"]: a["lv"] for a in CM.BASE}
_lv.update({n[0]: n[1] for n in TF.NEW3}); _lv.update({n[0]: n[1] for n in MG.NEW2})


def _is_parent(c):
    """통합 코드 체계 1+2+2+2+2+2: Level n의 유효 자릿수 = 1+2(n-1)."""
    p = c[:1 + 2 * (_lv.get(c, 6) - 1)]
    return any(x != c and x.startswith(p) for x in _lv)


opex = []
for a in CM.BASE + [dict(code=n[0], name=n[3], lv=n[1]) for n in TF.NEW3 if n[6] in ("비용", "고정자산")] + [dict(code=n[0], name=n[3], lv=n[1]) for n in MG.NEW2 if n[0].startswith("509")]:
    c = a["code"]
    if (c.startswith(("509", "5101", "51204", "51211"))) and a.get("lv", 0) >= 4 and not c.startswith(("509202001", "509202002", "509202099", "5101004")) and not _is_parent(c):
        line = TF.sch16_of(c)
        note = next((v for k, v in TF.SCH16_NOTE.items() if c.startswith(k)), "")
        live = [m for m in MIG if c in MG.targets_of(m["tgt"]) and m["stmt"] == "PL"]
        amt = sum(m["amt"] for m in live if MG.targets_of(m["tgt"])[0] == c)  # 분할 계정은 첫 대상에만 (EP 기준 재배분 전)
        opex.append((c, a["name"].strip(), line, TF.SCH16_LINES.get(line, "확인"), ", ".join(m["code"] for m in live), amt or None,
                     "신규" if (c in NEWALL or c in {n[0] for n in MG.NEW2}) else "", note, ho_view(c)))
opex.sort(key=lambda x: (list(TF.SCH16_LINES).index(x[2]) if x[2] in TF.SCH16_LINES else 99, x[0]))
_tot = sum(x[5] or 0 for x in opex)
_live_tot = sum(m["amt"] for m in MIG if m["code"].startswith(("5201", "5202", "5203", "52043")) and m["stmt"] == "PL")
assert abs(_tot - _live_tot) < 1, (_tot, _live_tot)
by_line = {}
for x in opex:
    by_line[x[2]] = by_line.get(x[2], 0) + (x[5] or 0)
r = section(ws, r, f"② Sch.16 줄별 합계 (현행 영업비용 ₹{_live_tot:,.0f} = 이행 후 합계, 분할 계정은 EP 재배분 전)",
            ["줄", "명칭", "금액(INR)"], [(k, TF.SCH16_LINES[k], round(by_line.get(k, 0))) for k in TF.SCH16_LINES])
r = section(ws, r, "③ 11자리 비용 계정 → Sch.16 줄 (현행 잔액 2026-07-07)", ["11자리", "계정명", "줄", "Sch.16 명칭", "현행 계정", "현행 잔액", "구분", "비고", "0930 원안 [본점매핑 / GBS]"], opex)
ep_rows = []
for e in TF.EP_LIST:
    if not e["gl"]:
        continue
    if e["ep"] in TF.EP_OVR:
        tgt, note = TF.EP_OVR[e["ep"]]
    else:
        m = mig_by.get(e["gl"])
        tgt, note = (m["tgt"] if m else ""), (m["note"] if m else "현행 GL 아님")
    t0 = MG.targets_of(tgt)
    line = TF.sch16_of(t0[0]) if t0 else ""
    ep_rows.append((e["ep"], e["name"], e["gl"], tgt, line, round(e["dr"] - e["cr"], 2) or None, note))
r = section(ws, r, "④ 경비(EP) 코드 → 현행 GL → 11자리 → Sch.16 (EP 금액 = 연간 차변 − 대변)", ["EP 코드", "경비명", "현행 GL", "11자리", "Sch.16", "EP 금액", "비고"], ep_rows)
for i, w in enumerate([14, 44, 8, 34, 22, 16, 10, 60], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---- 15 송금·무역 수수료 ------------------------------------------------------------
ws = wb.create_sheet("15 송금·무역수수료")
ws["A1"] = "송금·무역금융 수수료 구분 — 취급수수료 / 중계은행 수수료 / SWIFT 전신료"
ws["A1"].font = Font(name=F, size=14, bold=True)
ws["A2"] = ("Sch.14에서는 모두 'Commission, exchange and brokerage' 한 줄이지만, GST 과세가액(GSTR-1 ↔ GL 대사), pure agent 실비의 과세 제외, "
            "회수·지급 마진 관리 때문에 수수료 종류별 계정이 필요. 0930은 상품별 L4에 취급·전신료가 이미 나뉘어 있어 신설은 중계은행 회수분 1개(과세). 금액·부담자·SAC는 Charge code 파라미터.")
ws["A2"].font = Font(name=F, size=9, color="595959"); ws["A2"].alignment = WRAP; ws.merge_cells("A2:G2"); ws.row_dimensions[2].height = 42
r = section(ws, 4, "① 수수료 종류 × 계정", ["수수료 종류", "상품", "수익 계정", "비용·상대 계정", "GST", "TDS·RCM", "비고", "0930 원안 [본점매핑 / GBS]"],
            [tuple(f) + (ho_view(f[2] + " " + f[3]),) for f in TF.FEE_MATRIX])
r = section(ws, r, "② 신설 계정", ["코드", "Lv", "상위", "영문명", "한글명", "사유"], [n[:6] for n in TF.NEW3 if n[6] == "수수료"])
for i, w in enumerate([30, 26, 40, 40, 24, 30, 50], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---- 16 결산 증거 -------------------------------------------------------------------
ws = wb.create_sheet("16 FY25-26 결산근거")
ws["A1"] = "FY2025-26 재무제표 작업파일에서 확인한 근거 → 변경 연결"
ws["A1"].font = Font(name=F, size=14, bold=True)
ws["A2"] = ("원천: 20260609_Financials_NH_Bank_FY_2025-26.xlsx — Grouping(GBS 계정 → Third Schedule 줄), Adjustment Entries(결산 수정분개), "
            "Law Charges & Insurance(재분류 내역). 결산에서 수작업으로 고치는 항목이 곧 계정 체계가 놓친 부분.")
ws["A2"].font = Font(name=F, size=9, color="595959"); ws["A2"].alignment = WRAP; ws.merge_cells("A2:F2"); ws.row_dimensions[2].height = 30
r = section(ws, 4, "결산 증거 20건", ["결산 자료", "내용", "금액(INR)", "시사점", "관련 변경", "시트"], SD.EVIDENCE)
for i, w in enumerate([22, 50, 16, 40, 46, 8], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

wb.move_sheet("08 결정사항", offset=len(wb.sheetnames) - 1 - wb.sheetnames.index("08 결정사항"))

# ---- 시트별 변경 근거 블록 (상단 삽입) ------------------------------------------------
from openpyxl.worksheet.cell_range import CellRange

DOC_FILL = {SD.HO: "E7E6E6", SD.OURS: "FCE4D6", SD.WHY: "DDEBF7", SD.EVID: "E2EFDA"}


def add_doc(ws, doc):
    n = len(doc) + 1
    merges = [CellRange(str(m)) for m in ws.merged_cells.ranges]
    for m in merges:
        ws.unmerge_cells(str(m))
    heights = {i: d.height for i, d in ws.row_dimensions.items() if d.height}
    ws.insert_rows(2, n)
    for i in sorted(heights, reverse=True):
        if i >= 2:
            ws.row_dimensions[i + n].height = heights[i]
            ws.row_dimensions[i].height = None
    for m in merges:
        if m.min_row >= 2:
            m.shift(row_shift=n)
        ws.merge_cells(m.coord)
    if ws.auto_filter.ref:
        af = CellRange(ws.auto_filter.ref); af.shift(row_shift=n); ws.auto_filter.ref = af.coord
    if ws.freeze_panes:
        fp = ws[ws.freeze_panes]; ws.freeze_panes = ws.cell(row=fp.row + n, column=fp.column).coordinate
    last = max(6, min(ws.max_column, 10))
    for i, (k, v) in enumerate(doc, 2):
        a = ws.cell(row=i, column=1, value=k)
        b = ws.cell(row=i, column=2, value=v)
        a.font = Font(name=F, size=9, bold=True); b.font = Font(name=F, size=9)
        a.fill = b.fill = PatternFill("solid", fgColor=DOC_FILL[k])
        a.alignment = b.alignment = WRAP; a.border = b.border = BOX
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=last)
        ws.row_dimensions[i].height = max(18, 13 * (len(v) // 110 + 1))


for name, doc in SD.SHEET_DOC.items():
    add_doc(wb[name], doc)
wb.save(OUT)
print("saved", OUT)
