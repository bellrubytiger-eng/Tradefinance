# -*- coding: utf-8 -*-
"""데이터사전·업무규칙 Excel 생성: python3 tools/build_xlsx.py docs/<file>.xlsx"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
import model as M

FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
VERIFY_FILL = PatternFill("solid", fgColor="FFF2CC")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
NEW_FILL = PatternFill("solid", fgColor="F4B084")    # 주황: 인도지점 신규 계정
MOD_FILL = PatternFill("solid", fgColor="FCE4D6")    # 연주황: 인도지점 적용 변경
INACT_FILL = PatternFill("solid", fgColor="D9D9D9")  # 회색: 인도지점 비활성
FIX_FILL = PatternFill("solid", fgColor="FFF2CC")    # 노랑: 통합 CoA 정정 권고
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")


def sheet(wb, title, headers, rows, widths, verify_col=None, note=None):
    ws = wb.create_sheet(title)
    r0 = 1
    if note:
        ws.cell(1, 1, note).font = Font(name=FONT, italic=True, size=9, color="595959")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        r0 = 2
    for c, h in enumerate(headers, 1):
        cell = ws.cell(r0, c, h)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    for i, row in enumerate(rows, r0 + 1):
        for c, v in enumerate(row, 1):
            cell = ws.cell(i, c, v)
            if isinstance(v, str) and v.startswith("="):
                cell.data_type = "s"   # 설명문(= AMT + ...)이 수식으로 해석되지 않도록
            cell.font = Font(name=FONT, size=9)
            cell.alignment = WRAP
            cell.border = BORDER
        if verify_col is not None and row[verify_col] == "Y":
            for c in range(1, len(headers) + 1):
                ws.cell(i, c).fill = VERIFY_FILL
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(r0 + 1, 1)
    ws.auto_filter.ref = f"A{r0}:{get_column_letter(len(headers))}{r0 + len(rows)}"
    return ws


def main(out):
    wb = Workbook()
    ws = wb.active
    ws.title = "표지"
    f = lambda **k: Font(name=FONT, **k)
    ws["B2"] = "외환·무역송금 코어뱅킹 데이터사전 및 업무규칙"; ws["B2"].font = f(size=18, bold=True, color="1F4E79")
    ws["B3"] = M.META["project"]; ws["B3"].font = f(size=11)
    info = [("대상", M.META["branch"]), ("버전", M.META["version"]), ("작성일", M.META["date"]),
            ("테이블 수", len(M.T)), ("컬럼 수", sum(len(t["cols"]) for t in M.T)),
            ("업무규칙 수", len(M.RULES)), ("인터페이스 수", len(M.INTERFACES))]
    for i, (k, v) in enumerate(info, 5):
        ws.cell(i, 2, k).font = f(bold=True); ws.cell(i, 3, v).font = f()
    r = 13
    ws.cell(r, 2, "시트 구성").font = f(bold=True, size=12); r += 1
    guide = [("1.화면목록", "TO-BE 화면 ID / 대상고객 / 기능 / 관련 테이블"),
             ("2.테이블목록", "엔티티 목록과 업무영역"),
             ("3.컬럼정의", "전 테이블 컬럼 (타입·제약·관리규칙·외부매핑) - 보고서 시스템 연동 기준"),
             ("4.업무규칙", "검증 규칙(ERR/WARN/ALERT/AUTO)과 근거 규정"),
             ("5.기한관리", "BoE·Invoice 유효기간, 선수금 증빙, 수출회수, FDI 등 규제기한 산식"),
             ("6.규정파라미터", "TB_FX_PARAM 초기값 - 규정 개정 시 이력으로 추가"),
             ("7.기한계산기", "기한·LRS·LEI·GST 계산 예시(노란 칸 입력 → 수식 자동계산)"),
             ("8.코드정의", "공통코드 / Purpose Code 샘플"),
             ("9.인터페이스", "IDPMS·EDPMS·CIMS·SWIFT·GLEIF 등"),
             ("10.회계분개", "이벤트별 분개 요약 (통합 표준 CoA 계정 기준, 18번 시트 요약)"),
             ("11.요청사항추적", "기존 개발요청서 항목별 반영위치·검토의견"),
             ("12.서식필드매핑", "고객 서식 4종(FIRC·당발·타발·환전) 항목별 DB 소스"),
             ("13.FDI화면가이드", "FDI 단계별 사용 화면"),
             ("14.SWIFT카탈로그", "MT1xx~9xx ↔ MX 대응, 송신 형식, 자동생성·수신 연결 화면"),
             ("15.RBI기준검토", "업로드 무역금융 자료·기존 설계의 RBI 최신 기준 대조 결과"),
             ("16.규정근거색인", "규정별 번호·고시일·시행일·관련 파라미터·규칙·출처"),
             ("17.CoA(인도지점)", "06 인도 반영 CoA 전체 1,360계정: 신규(주황)·변경/용도지정/기한부(연주황)·비활성/철회(회색)·IFRS 전용(하늘)·0930 정정권고(노랑), 현행 8자리·잔액"),
             ("18.업무별분개", "업무별 예시 분개(통화별 대차 검증 수식 포함). 금액은 고객 서식 샘플과 일치"),
             ("19.계정이행맵핑", "현행 인도 GBS 8자리 → 차세대 11자리 이행표, 분할 기준·검토 결과·시스템 사용"),
             ("20.CoA검토의견", "06 인도 반영 CoA 검토 의견(추가·수정·삭제)")]
    for k, v in guide:
        ws.cell(r, 2, k).font = f(bold=True); ws.cell(r, 3, v).font = f(); r += 1
    r += 1
    ws.cell(r, 2, "범례").font = f(bold=True, size=12); r += 1
    ws.cell(r, 2, "").fill = VERIFY_FILL
    ws.cell(r, 3, "VERIFY=Y : 규정 개정 가능성이 높은 수치/코드. 운영 반영 전 컴플라이언스 확인 필수").font = f(); r += 1
    ws.cell(r, 2, "").fill = INPUT_FILL
    ws.cell(r, 3, "입력 셀 (7.기한계산기)").font = f(); r += 1
    for fill, txt in [(NEW_FILL, "17·18번 시트: 인도지점 신규 계정 (본점 채번 전 임시코드, 메모 참조)"),
                      (MOD_FILL, "17번 시트: 변경·용도지정·기한부 / 19번 시트: 분할규칙 필요"),
                      (INACT_FILL, "17번 시트: 비활성·철회 (통합 CoA에서 삭제하지 않음)"),
                      (IFRS_FILL, "17번 시트: 본점 IFRS 조정 전용 (인도 원장 기장 차단)"),
                      (FIX_FILL, "17번 시트 0930 정정 권고 / 19번 시트 대상 정정(수정)")]:
        ws.cell(r, 2, "").fill = fill; ws.cell(r, 3, txt).font = f(); r += 1
    r += 1
    ws.cell(r, 2, "기반 자료").font = f(bold=True, size=12); r += 1
    for d in M.META["base_docs"]:
        ws.cell(r, 3, d).font = f(size=9); r += 1
    ws.column_dimensions["B"].width = 18; ws.column_dimensions["C"].width = 90

    sheet(wb, "1.화면목록", ["화면ID", "화면명", "대상", "구분", "주요기능", "관련테이블", "HTML 시안"],
          [(s[0], s[1], s[2], s[3], s[4], s[5], f"screens/index.html#{s[6]}") for s in M.SCREENS],
          [8, 32, 14, 8, 70, 34, 26])

    sheet(wb, "2.테이블목록", ["No", "테이블명", "한글명", "업무영역", "설명", "컬럼수"],
          [(i, t["name"], t["kor"], t["area"], t["desc"], len(t["cols"])) for i, t in enumerate(M.T, 1)],
          [5, 24, 22, 16, 90, 8])

    rows = []
    for t in M.T:
        for i, c in enumerate(t["cols"], 1):
            rows.append((t["name"], t["kor"], i, c[0], c[1], c[2], "Y" if "PK" in c[3] else "",
                         "Y" if "FK" in c[3] else "", "N" if ("NN" in c[3] or "PK" in c[3]) else "Y", c[4], c[5]))
    sheet(wb, "3.컬럼정의", ["테이블명", "테이블 한글명", "순번", "컬럼명", "컬럼 한글명", "데이터타입", "PK", "FK",
                          "NULL 허용", "관리규칙 / 설명", "외부보고 매핑"], rows,
          [22, 18, 5, 22, 22, 15, 4, 4, 6, 70, 16])

    sheet(wb, "4.업무규칙", ["규칙ID", "업무", "화면", "규칙명", "조건 / 로직", "처리", "근거", "파라미터", "VERIFY"],
          M.RULES, [10, 8, 14, 20, 80, 7, 30, 24, 7], verify_col=8,
          note="처리: ERR=거래차단 / WARN=경고 후 진행(승인기록) / ALERT=사후알림·독촉 / AUTO=시스템 자동처리")

    sheet(wb, "5.기한관리", ["기한유형", "업무", "기산일", "기한 산식", "사전알림", "이행(종결) 조건", "기한경과 시 처리",
                          "보고대상", "파라미터", "VERIFY"], M.DEADLINES, [13, 7, 16, 38, 12, 28, 38, 9, 24, 7],
          verify_col=9, note="모든 기한은 TB_FX_OBLIGATION 1행으로 생성되며, 기산일에 유효한 파라미터 값을 스냅샷으로 저장한다.")

    ws = sheet(wb, "6.규정파라미터", ["PARAM_ID", "명칭", "값", "단위", "적용업무", "근거", "VERIFY", "비고"],
               M.PARAMS, [28, 38, 14, 8, 10, 46, 7, 60], verify_col=6)

    build_calc(wb)

    sheet(wb, "8.코드정의", ["코드그룹", "코드", "코드명", "설명", "VERIFY"], M.CODES + [
        ("PURPOSE", p[0], p[2], f"방향:{p[1]} / 업무군:{p[3]}", p[4]) for p in M.PURPOSE_CODES],
          [16, 14, 44, 60, 7], verify_col=4,
          note="Purpose Code는 샘플입니다. 운영 시 RBI 최신 Purpose Code 전체 목록을 TB_CD_PURPOSE로 적재하십시오.")

    sheet(wb, "9.인터페이스", ["IF ID", "대상시스템", "방향", "방식", "주기", "내용", "관련 테이블", "VERIFY"],
          M.INTERFACES, [11, 20, 7, 32, 14, 44, 36, 7], verify_col=7,
          note="RBI 시스템(IDPMS/EDPMS/CIMS) 파일 레이아웃은 RBI가 AD 은행에 배포하는 최신 명세서로 확정해야 합니다(VERIFY).")

    sheet(wb, "10.회계분개", ["이벤트", "차변", "대변", "비고"], M.GL, [26, 34, 56, 30])

    sheet(wb, "11.요청사항추적", ["출처", "기존 요청", "반영 위치", "검토의견"], M.TRACE, [18, 52, 44, 60])

    rows = []
    for fm in M.FORMS:
        vnames = dict(fm["variants"])
        for sec in fm["sections"]:
            for f in sec["f"]:
                cond = "전체" if not f[4] else ", ".join(vnames[v] for v in f[4])
                rows.append((fm["id"], fm["ko"], sec["t"], f[0], f[1], f[2], f[3], cond))
    sheet(wb, "12.서식필드매핑", ["서식ID", "서식명", "섹션", "필드키", "영문 라벨", "한글 항목", "DB 소스", "표시 조건(변형)"], rows,
          [11, 26, 26, 12, 30, 18, 52, 40],
          note="고객 서식(screens/forms.html)의 각 항목이 어느 테이블·컬럼에서 오는지. 발급 시 값 전체를 TB_FX_ADVICE_ISSUE.SNAPSHOT_JSON에 저장")
    sheet(wb, "13.FDI화면가이드", ["단계", "상황", "사용 화면", "처리 내용", "생성 데이터/서식"], M.FDI_FLOW, [6, 34, 30, 70, 40],
          note="FDI 유입은 RT21 → FD11 (비계좌는 RT22, 에스크로는 FD13), 반환은 FD12. RT11·RT13은 FDI 유입에 사용하지 않음")
    sheet(wb, "15.RBI기준검토", ["No", "대상", "기존 내용", "RBI 최신 기준 (2026-11-01)", "반영", "근거", "확인 상태"], M.RBI_REVIEW,
          [5, 22, 40, 70, 44, 34, 14],
          note="확인 상태: 2차출처 = 법률·회계법인 요약으로 확인(RBI 원문 대조 필요) / 미확인 = 출처 미발견, 기존 값을 지점 정책으로 유지 / 계산확인 = 산식 검증")
    sheet(wb, "14.SWIFT카탈로그", ["유형", "표준", "Cat", "명칭", "대응 유형", "현재 송신형식", "방향", "자동생성 원천", "수신 시 연결", "VERIFY"],
          M.SWIFT_TYPES, [11, 6, 5, 40, 18, 10, 7, 34, 20, 7], verify_col=9,
          note="CBPR+ 결제·은행이체·입출금통지는 2025.11 이후 MX로 송신(MT는 수신 번역·레거시용). 보고(MT940/950)·Cat3·Cat7은 MT 유지 - SWIFT 최신 일정 확인(VERIFY)")

    wb.move_sheet("15.RBI기준검토", offset=1)
    sheet(wb, "16.규정근거색인", ["규정", "번호", "고시일", "시행일", "적용 업무", "관련 파라미터", "관련 규칙", "확인 상태", "출처"], M.REG_INDEX,
          [52, 22, 14, 22, 40, 34, 30, 20, 60],
          note="설계에 쓰인 모든 규정·국제규칙의 번호·일자·적용 범위와, 해당 규정을 근거로 하는 파라미터(6번 시트)·업무규칙(4번 시트) 색인")
    build_coa(wb)
    build_journal(wb)
    build_map(wb)
    build_review(wb)
    wb.save(out)
    print("saved", out)


def build_calc(wb):
    """기한 계산기: 파라미터 시트 값을 참조하는 수식."""
    ws = wb.create_sheet("7.기한계산기")
    f = lambda **k: Font(name=FONT, **k)
    p_row = {p[0]: i + 2 for i, p in enumerate(M.PARAMS)}   # 6.규정파라미터 데이터행 (헤더=1행)
    ref = lambda pid: f"'6.규정파라미터'!$C${p_row[pid]}"
    ws["A1"] = "규제기한 계산기 (노란 칸 입력, 나머지는 수식)"; ws["A1"].font = f(bold=True, size=13, color="1F4E79")
    ws["A2"] = "파라미터는 '6.규정파라미터' 시트를 참조합니다. 값 변경 시 자동 재계산됩니다."; ws["A2"].font = f(size=9, italic=True)
    rows = [
        ("수입 선수금", None, None),
        ("선수금 송금일", "2026-04-15", "input_date"),
        ("자본재 여부(Y/N)", "N", "input"),
        ("송금액(USD)", 250000, "input_num"),
        ("수입기한", f"=EDATE(B5,IF(B6=\"Y\",VALUE({ref('IMP_ADV_IMPORT_MONTHS_CAP')}),VALUE({ref('IMP_ADV_IMPORT_MONTHS')})))", "date"),
        ("수입증빙(BoE) 제출기한", f"=B8+VALUE({ref('IMP_ADV_EVIDENCE_GRACE_DAYS')})", "date"),
        ("BG/SBLC 요구 (지점 정책)", f"=IF(B7>VALUE({ref('IMP_ADV_BG_THRESHOLD_USD')}),\"필요\",\"불요\")", "text"),
        ("BoE 등록 허용 최종일", f"=B5+VALUE({ref('BOE_REG_LIMIT_DAYS')})", "date"),
        ("", None, None),
        ("수출", None, None),
        ("LEO(수출일)", "2026-03-20", "input_date"),
        ("해외창고 수출(Y/N)", "N", "input"),
        ("서류제출기한", f"=B14+VALUE({ref('EXP_DOC_SUBMIT_DAYS')})", "date"),
        ("대금회수기한", f"=EDATE(B14,IF(B15=\"Y\",VALUE({ref('EXP_REALIZE_MONTHS_WH')}),VALUE({ref('EXP_REALIZE_MONTHS')})))", "date"),
        ("", None, None),
        ("FDI", None, None),
        ("FDI 수취일", "2026-09-01", "input_date"),
        ("배정기한(+60)", f"=B20+VALUE({ref('FDI_ALLOT_DAYS')})", "date"),
        ("반환기한(+75)", f"=B21+VALUE({ref('FDI_REFUND_DAYS')})", "date"),
        ("배정일", "2026-10-10", "input_date"),
        ("FC-GPR 신고기한", f"=B23+VALUE({ref('FDI_FCGPR_DAYS')})", "date"),
        ("", None, None),
        ("LRS / LEI", None, None),
        ("FY 기송금 누적(USD, 당행+신고)", 180000, "input_num"),
        ("금회 송금(USD)", 50000, "input_num"),
        ("LRS 잔여한도(USD)", f"=VALUE({ref('LRS_LIMIT_USD')})-B27-B28", "num"),
        ("한도 판정", "=IF(B29>=0,\"가능\",\"한도초과-차단\")", "text"),
        ("기업송금 INR 환산액", 420000000, "input_num"),
        ("LEI 필수 여부", f"=IF(B31>=VALUE({ref('LEI_THRESHOLD_INR')}),\"필수(만료 시 차단)\",\"선택(만료 시 경고)\")", "text"),
        ("", None, None),
        ("GST (수수료 + 환전가치, R-GST-01~05)", None, None),
        ("환전 INR 금액 (무환전=0)", 8842000, "input_num"),
        ("과세 수수료 합계 (Commission+Cable 등)", 12052.5, "input_num"),
        ("고객 주코드 (GSTIN 앞 2자리)", "09", "input"),
        ("GST 수령자 유형", "R_ACCT", "input"),
        ("환전 과세표준 32(2)(b)", "=IF(B35<=0,0,IF(B35<=100000,MAX(250,ROUND(B35*0.01,2)),IF(B35<=1000000,ROUND(1000+(B35-100000)*0.005,2),MIN(60000,ROUND(5500+(B35-1000000)*0.001,2)))))", "num2"),
        ("과세표준 합계", "=B36+B39", "num2"),
        ("공급 구분", f"=IF(B38=\"NR_NONACCT\",\"국외(영세율)\",IF(OR(B38<>\"R_ACCT\",B37={ref('BRANCH_STATE_CODE')}),\"주내\",\"주간\"))", "text"),
        ("GST 18% 합계", "=IF(B41=\"국외(영세율)\",0,ROUND(B40*0.18,2))", "num2"),
        ("CGST 9%", "=IF(B41=\"주내\",ROUND(B42/2,2),0)", "num2"),
        ("SGST 9% (= 합계 - CGST)", "=IF(B41=\"주내\",B42-B43,0)", "num2"),
        ("IGST 18%", "=IF(B41=\"주간\",B42,0)", "num2"),
        ("", None, None),
        ("본지점 을B 금리 (R-IO-13/14/16)", None, None),
        ("차입 원금", 5000000, "input_num"),
        ("기간(일)", 92, "input_num"),
        ("기준금리 %", 4.30, "input_num"),
        ("표준 스프레드 (bp, 기간구간)", 75, "input_num"),
        ("우대 스프레드 (bp, 기간구간)", 45, "input_num"),
        ("우대조건 충족 (Y/N)", "Y", "input"),
        ("일수 기준 (360/365)", 360, "input_num"),
        ("적용금리 %", "=B50+IF(B53=\"Y\",B52,B51)/100", "num4"),
        ("이자", "=ROUND(B48*B55/100*B49/B54,2)", "num2"),
        ("기간구간", "=IF(B49<=7,\"T007\",IF(B49<=30,\"T030\",IF(B49<=90,\"T090\",IF(B49<=180,\"T180\",IF(B49<=365,\"T365\",\"T999 (본점 별도승인)\")))))", "text"),
        ("", None, None),
        ("FDI 약정금액·초과입금 (R-FDI-06~08)", None, None),
        ("약정 INR (발행가 × 주식수)", 250000000, "input_num"),
        ("딜 환율 (INR per USD)", 88.42, "input_num"),
        ("수취 외화 (USD)", 2850000, "input_num"),
        ("환전 대상 외화 = ROUNDUP(약정/환율, 2)", "=ROUNDUP(B60/B61,2)", "num2"),
        ("초과 외화 (환전 없이 원통화 반환)", "=ROUND(MAX(0,B62-B63),2)", "num2"),
        ("부족 외화 (환전 보류)", "=ROUND(MAX(0,B63-B62),2)", "num2"),
        ("입금 INR", "=IF(B65>0,0,ROUND(B63*B61,2))", "num2"),
        ("처리", "=IF(B65>0,\"부족 → 환전 보류·부족분 요청\",IF(B64>0,\"약정분만 환전, 초과분 원통화 반환\",\"전액 환전\"))", "text"),
    ]
    from datetime import date
    for i, (label, val, kind) in enumerate(rows, 4):
        a = ws.cell(i, 1, label)
        a.font = f(bold=(val is None and label != ""), size=10, color="1F4E79" if val is None else "000000")
        if val is None:
            continue
        if kind == "input_date":
            y, m, d = map(int, val.split("-")); val = date(y, m, d)
        b = ws.cell(i, 2, val)
        b.font = f(size=10, color="0000FF" if kind.startswith("input") else "000000")
        b.border = BORDER
        if kind.startswith("input"):
            b.fill = INPUT_FILL
        if kind in ("date", "input_date"):
            b.number_format = "yyyy-mm-dd"
        if kind in ("num", "input_num"):
            b.number_format = "#,##0;(#,##0);-"
        if kind == "num4":
            b.number_format = "0.0000"
        if kind == "num2" or (kind == "input_num" and isinstance(val, float)):
            b.number_format = "#,##0.00;(#,##0.00);-"
    ws.column_dimensions["A"].width = 34; ws.column_dimensions["B"].width = 24
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
    ws.add_data_validation(dv)
    for c in ("B6", "B15"):
        dv.add(c)
    dv2 = DataValidation(type="list", formula1='"R_ACCT,NR_ACCT,R_NONACCT,NR_NONACCT"', allow_blank=False)
    ws.add_data_validation(dv2)
    dv2.add("B38")
    dv.add("B53")
    ws["A71"] = "갑기금 = 무이자(계산 없음) · 을A = 기준금리+스프레드(휴일차이·정산지연) · 을B = 기간구간 금리, 우대조건: 과거 1년 내 동일통화·금액 이상·30일 이상 조달이력(정기예금·콜머니·콜론 등)"
    ws["A71"].font = Font(name=FONT, size=9, italic=True)
    ws["A70"] = "R_ACCT 거주자 계좌고객(고객 주 기준) · NR_ACCT 비거주 계좌고객 · R_NONACCT 비계좌 · NR_NONACCT 해외 비계좌(영세율). 근거: IGST Act s.12(12), s.13(8)(a), s.13(2)"
    ws["A70"].font = Font(name=FONT, size=9, italic=True)




IFRS_FILL = PatternFill("solid", fgColor="DDEBF7")   # 하늘: 본점 IFRS 조정 전용(인도 원장 차단)
ERR_FILL = PatternFill("solid", fgColor="F8CBAD")


def _hdr(ws, r, hdr):
    for c, h in enumerate(hdr, 1):
        x = ws.cell(r, c, h); x.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); x.fill = HDR_FILL
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); x.border = BORDER


def build_coa(wb):
    ws = wb.create_sheet("17.CoA(인도지점)")
    rows = M.as_dict()["COA"]
    cnt = lambda k: sum(1 for r in rows if r["status"] == k)
    ws.cell(1, 1, f"차세대 GBS 통합 표준 CoA(0930) · 인도(노이다) 반영본(06 파일) — {len(rows)}계정 / 신규 {cnt('신규')} (주황) · 변경 {cnt('변경')}·용도지정 {cnt('용도지정')} (연주황) · "
                  f"비활성 {cnt('비활성')}·철회 {cnt('철회')} (회색) · 기한부 {cnt('기한부')} · IFRS 전용 {sum(1 for r in rows if 'IFRS' in r['post'])} (하늘)").font = Font(name=FONT, bold=True, size=11, color="1F4E79")
    ws.cell(2, 1, "A~J 본점 0930 원안, K열부터 인도 검토(06 파일). 기장은 leaf·사용 계정만(결정 G2). '기한부'는 이번 검토로 06 파일 상태를 정정한 계정(결정 L3). 계정코드 셀 메모에 사유.").font = Font(name=FONT, italic=True, size=9, color="595959")
    hdr = ["계정코드", "Lv", "구분", "업무영역", "계정과목(영문)", "인도 표시명", "본점 통합계정", "본점 통합계정명", "비화폐성", "GBS코드(0930)",
           "인도 상태", "인도 원장 기장", "leaf", "현행 인도 GBS(8자리)", "현행 잔액 INR(2026-07-07)", "변경 사유·근거", "인도 재무제표 줄", "0930 정정 권고"]
    _hdr(ws, 4, hdr)
    for i, a in enumerate(rows, 5):
        vals = [a["code"], a["lv"], a["grp"], a["area"], "  " * (a["lv"] - 1) + a["name"], a["br_name"], a["ho"], a["ho_name"], a["nonmon"], a["gbs"],
                a["status"], a["post"], "Y" if a["leaf"] else "", a["cur"], a["bal"], a["memo"], a["sch"], a["fix"]]
        st = a["status"]
        fill = NEW_FILL if st == "신규" else MOD_FILL if st in ("변경", "용도지정", "기한부") else INACT_FILL if st in ("비활성", "철회") else IFRS_FILL if "IFRS" in a["post"] else None
        for c, v in enumerate(vals, 1):
            x = ws.cell(i, c, v); x.border = BORDER
            x.font = Font(name=FONT, size=9, bold=(a["lv"] <= 2 or st == "신규"), color="7F7F7F" if st in ("비활성", "철회") else "000000", strike=(st == "철회"))
            x.alignment = Alignment(vertical="top", wrap_text=c in (5, 6, 8, 16, 17, 18))
            if c == 15:
                x.number_format = "#,##0.00"
            if fill:
                x.fill = fill
        if a["fix"]:
            ws.cell(i, 18).fill = FIX_FILL
        if st:
            cm = Comment(f"[{st}] {a['memo'] or ''}"[:600], "CBS 설계"); cm.width, cm.height = 340, 120
            ws.cell(i, 1).comment = cm
    for c, w in enumerate([13, 4, 9, 13, 44, 30, 15, 20, 7, 10, 9, 12, 5, 16, 15, 46, 22, 30], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A4:R{4 + len(rows)}"
    r = 6 + len(rows)
    ws.cell(r, 1, "잔액 합계 검증 (차변+, 0이면 현행 잔액 전체가 이관)").font = Font(name=FONT, bold=True, size=9)
    ws.cell(r, 15, f"=SUM(O5:O{4 + len(rows)})").number_format = "#,##0.00"


def build_map(wb):
    ws = wb.create_sheet("19.계정이행맵핑")
    d = M.as_dict()
    rows = d["COA_MAP"]
    ws.cell(1, 1, "현행 인도 GBS 8자리 → 차세대 11자리 이행 맵핑 (06 파일 기준 + 검토)").font = Font(name=FONT, bold=True, size=13, color="1F4E79")
    sev = {k: sum(1 for m in rows if m["sev"] == k) for k in ("OK", "분할규칙", "수정", "오류")}
    ws.cell(2, 1, f"현행 계정 {len({m['cur'] for m in rows})}개 · 연결 {len(rows)}행 — OK {sev['OK']} / 분할규칙 {sev['분할규칙']} / 수정 {sev['수정']} / 오류 {sev['오류']}. "
                  "'권고 대상'이 06 파일 대상과 다르면 정정 제안. 시스템 사용 열은 당 설계의 화면·분개가 기장하는 계정.").font = Font(name=FONT, italic=True, size=9, color="595959")
    used = {}
    for jr in d["JOURNAL"]:
        if jr[7]:
            used.setdefault(jr[7], set()).add(jr[1])
    hdr = ["현행 계정(8자리)", "유형", "06 대상(11자리)", "대상 계정명", "Lv", "인도 상태", "인도 원장 기장", "기준일 잔액(INR, 차변+)", "검토", "권고 대상", "검토 의견·분할 기준", "시스템 사용(분개 ID)"]
    _hdr(ws, 4, hdr)
    for i, m in enumerate(rows, 5):
        vals = [m["cur"], m["kind"], m["code"], m["name"], m["lv"], m["status"], m["post"], m["bal"], m["sev"], m["target"], m["issue"],
                ", ".join(sorted(used.get(m["target"], set()) | used.get(m["code"], set())))]
        f = {"수정": FIX_FILL, "분할규칙": MOD_FILL, "오류": ERR_FILL}.get(m["sev"])
        for c, v in enumerate(vals, 1):
            x = ws.cell(i, c, v); x.border = BORDER; x.font = Font(name=FONT, size=9, bold=(c == 9 and m["sev"] != "OK"))
            x.alignment = Alignment(vertical="top", wrap_text=c in (4, 11, 12))
            if c == 8:
                x.number_format = "#,##0.00"
            if f and c in (9, 10, 11):
                x.fill = f
    last = 4 + len(rows)
    ws.cell(last + 2, 1, "잔액 합계 (중복 대상 제외 전, 참고)").font = Font(name=FONT, bold=True, size=9)
    ws.cell(last + 2, 8, f"=SUM(H5:H{last})").number_format = "#,##0.00"
    for c, w in enumerate([12, 14, 13, 34, 4, 9, 12, 16, 9, 13, 64, 18], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A4:L{last}"


def build_review(wb):
    ws = wb.create_sheet("20.CoA검토의견")
    ws.cell(1, 1, "06 통합표준 CoA(인도 반영본) 검토 의견 — 추가·수정·삭제").font = Font(name=FONT, bold=True, size=13, color="1F4E79")
    ws.cell(2, 1, "결정 열은 06 파일 '결정사항_모범안' ID. '반영'은 당 설계(화면·분개)에 이미 적용한 사항.").font = Font(name=FONT, italic=True, size=9, color="595959")
    _hdr(ws, 4, ["구분", "중요도", "대상", "내용", "권고", "관련 결정"])
    for i, r in enumerate(M.as_dict()["COA_REVIEW"], 5):
        for c, v in enumerate(r, 1):
            x = ws.cell(i, c, v); x.border = BORDER; x.font = Font(name=FONT, size=9, bold=c in (1, 2)); x.alignment = WRAP
        ws.cell(i, 2).fill = {"높음": ERR_FILL, "중간": FIX_FILL}.get(r[1], PatternFill())
    for c, w in enumerate([7, 7, 30, 70, 60, 12], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A5"


def build_journal(wb):
    ws = wb.create_sheet("18.업무별분개")
    d = M.as_dict()
    status = {r["code"]: r["status"] for r in d["COA"]}
    ws.cell(1, 1, "업무별 예시 분개 (통합 표준 CoA 기준)").font = Font(name=FONT, bold=True, size=13, color="1F4E79")
    ws.cell(2, 1, "환전 거래는 통화별로 대차를 맞추고 11200000000(통화별 포지션)으로 연결. 난외(OB)는 대차검증 제외. "
                  "주황 계정 = 인도지점 신규(임시코드). 각 단계 아래 '검증' 행은 SUMIFS 수식으로 통화별 차변=대변 확인.").font = Font(name=FONT, italic=True, size=9, color="595959")
    hdr = ["업무", "시나리오", "내용", "화면", "단계", "일자", "차/대", "계정코드", "계정명", "통화", "차변", "대변", "적요", "근거", "단계키"]
    R0 = 4
    for c, h in enumerate(hdr, 1):
        x = ws.cell(R0, c, h); x.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); x.fill = HDR_FILL
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); x.border = BORDER
    rows = d["JOURNAL"]
    total = R0 + len(rows) + 200
    rng = lambda col: f"${col}${R0 + 1}:${col}${total}"
    r = R0 + 1
    prev_biz, prev_key = None, None
    groups = []
    for row in rows:
        key = f"{row[1]}|{row[4]}"
        groups.append((key, row))
    i = 0
    while i < len(groups):
        key, row = groups[i]
        if row[0] != prev_biz:
            x = ws.cell(r, 1, f"■ {row[0]}"); x.font = Font(name=FONT, bold=True, size=11, color="FFFFFF")
            for c in range(1, 16):
                ws.cell(r, c).fill = PatternFill("solid", fgColor="2E75B6")
            r += 1; prev_biz = row[0]
        block = []
        while i < len(groups) and groups[i][0] == key:
            block.append(groups[i][1]); i += 1
        for b in block:
            vals = list(b[:13]) + [b[13], key]
            for c, v in enumerate(vals, 1):
                x = ws.cell(r, c, v); x.font = Font(name=FONT, size=9); x.border = BORDER
                x.alignment = Alignment(vertical="top", wrap_text=c in (3, 9, 13, 14))
                if c in (11, 12):
                    x.number_format = "#,##0.00"
            if b[7] and status.get(b[7]) == "신규":
                ws.cell(r, 8).fill = NEW_FILL; ws.cell(r, 9).fill = NEW_FILL
            if b[6] in ("난외+", "난외-"):
                for c in (7, 8, 9, 10, 11, 12):
                    ws.cell(r, c).font = Font(name=FONT, size=9, italic=True, color="7F7F7F")
            r += 1
        for ccy in dict.fromkeys(b[9] for b in block if b[6] in ("차변", "대변")):
            ws.cell(r, 7, "검증").font = Font(name=FONT, size=9, bold=True)
            ws.cell(r, 10, ccy).font = Font(name=FONT, size=9, bold=True)
            ws.cell(r, 11, f'=SUMIFS({rng("K")},{rng("O")},"{key}",{rng("J")},"{ccy}",{rng("G")},"차변")')
            ws.cell(r, 12, f'=SUMIFS({rng("L")},{rng("O")},"{key}",{rng("J")},"{ccy}",{rng("G")},"대변")')
            ws.cell(r, 13, f'=IF(ROUND(K{r}-L{r},2)=0,"OK","불일치")')
            for c in range(7, 14):
                x = ws.cell(r, c); x.fill = SUB_FILL; x.border = BORDER
                if c in (11, 12):
                    x.number_format = "#,##0.00"; x.font = Font(name=FONT, size=9, bold=True)
            ws.cell(r, 13).font = Font(name=FONT, size=9, bold=True, color="006100")
            r += 1
    for c, w in enumerate([10, 9, 34, 12, 26, 11, 7, 13, 36, 6, 16, 16, 46, 22, 18], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.column_dimensions["O"].hidden = True
    ws.freeze_panes = f"A{R0 + 1}"

if __name__ == "__main__":
    main(sys.argv[1])
