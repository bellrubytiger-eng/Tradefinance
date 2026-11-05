# -*- coding: utf-8 -*-
"""차세대 GBS 통합 표준 CoA(0930) — 인도(노이다) 지점 반영본(06 파일) + 현행 8자리 → 차세대 11자리 이행 맵핑 + 업무별 예시 분개.

- 원천: tools/data/coa_india_0930.json  (06_..._CoA_0930 인도 반영본, 1,360계정, 결정사항 58건)
- ROWS    : 계정상태표 (A~J 본점 원안 + K~S 인도 검토)
- MAPPING : 현행 인도 GBS 8자리 → 11자리 (1:1 / N:1 / 1:N 분할), 분할 기준·검토 결과
- REVIEW  : 06 파일 검토 의견 (추가·수정·삭제 권고)
- JOURNAL : 업무별 예시 분개 (leaf·사용 계정만, 금액은 forms_model 샘플과 일치)
"""
import json
import os
import re
from collections import defaultdict
from forms_model import r2, gst_fx_value, legs_fixed

_HERE = os.path.dirname(os.path.abspath(__file__))
_SRC = json.load(open(os.path.join(_HERE, "data", "coa_india_0930.json"), encoding="utf-8"))
DECISIONS = _SRC["decisions"]
_SEG = [1, 3, 5, 7, 9, 11]


def level(c):
    segs = [c[0:1], c[1:3], c[3:5], c[5:7], c[7:9], c[9:11]]
    for i in range(5, -1, -1):
        if segs[i] not in ("0", "00"):
            return i + 1
    return 1


def parent_of(c):
    L = level(c)
    return None if L == 1 else c[:_SEG[L - 2]].ljust(11, "0")


BASE = []
for r in _SRC["rows"]:
    r = list(r) + [None] * (19 - len(r))
    BASE.append({"grp": r[0], "area": r[1], "code": str(r[2]), "lv": level(str(r[2])), "name": (r[4] or "").strip(), "krw": r[5] or "",
                 "ho": str(r[6] or ""), "ho_name": r[7] or "", "nonmon": r[8] or "", "gbs": str(r[9] or ""),
                 "status": r[10] or "", "post": r[11] or "사용", "br_name": (r[12] or "").strip(), "cur": str(r[13] or ""),
                 "bal": r[14], "memo": r[15] or "", "sch": r[16] or "", "fix": r[17] or "", "terms": r[18] or ""})
BY = {a["code"]: a for a in BASE}

# 검토 결과 반영: 06 결정 L3(일반·특정 충당금은 2027-03-31까지 사용)와 상태 열(비활성)이 불일치 → 기한부 사용으로 처리
TEMP_USE = {"10380101000": "2027-03-31", "10380102000": "2027-03-31"}


def usable(code, on=None):
    a = BY.get(code)
    if not a:
        return False
    if code in TEMP_USE:
        return (on or "2026-11-05") <= TEMP_USE[code]
    return a["post"] == "사용"


def _active_children(c):
    pre = c[:_SEG[level(c) - 1]]
    return [o for o in BY if o != c and o.startswith(pre) and level(o) == level(c) + 1 and BY[o]["post"] != "미사용"]


PARENTS = {c for c in BY if level(c) >= 2 and _active_children(c)}
LEAF = set(BY) - PARENTS


def acct_name(code):
    a = BY[code]
    return a["br_name"] or a["name"]


def coa_rows():
    """엑셀·DDL용 행 (코드순). status: 신규/변경/용도지정/비활성/철회/'' , post: 사용/미사용/본점 IFRS 조정 전용"""
    out = []
    for a in sorted(BASE, key=lambda x: x["code"]):
        r = dict(a)
        if a["code"] in TEMP_USE:
            r["status"] = "기한부"
            r["post"] = f"사용(~{TEMP_USE[a['code']]})"
            r["memo"] = (a["memo"] + " / " if a["memo"] else "") + f"검토: 결정 L3에 따라 {TEMP_USE[a['code']]}까지 사용 후 ECL Stage 계정으로 이관"
        r["leaf"] = a["code"] in LEAF
        out.append(r)
    return out


ALL = {r["code"]: r for r in coa_rows()}

# 하위호환(설계서 표): 신규·비활성·정정권고
ADD = [(a["code"], a["lv"], a["name"], a["br_name"], "", a["memo"]) for a in BASE if a["status"] == "신규"]
DEACT = [(a["code"], (a["br_name"] or a["name"]) + " — " + (a["memo"] or a["status"])) for a in BASE if a["status"] in ("비활성", "철회")]
FIX = [(a["code"], a["fix"]) for a in BASE if a["fix"]]

# ---------------------------------------------------------------------------
# 현행 8자리 → 차세대 11자리 이행 맵핑
# ---------------------------------------------------------------------------
_cur = defaultdict(list)
for a in BASE:
    for g in re.split(r"[,\s/;]+", a["cur"]):
        if re.fullmatch(r"\d{8}", g):
            _cur[g].append(a["code"])

SPLIT_RULE = {
    "12110002": "FAR 자산분류(단말/서버·네트워크/전기설비) - 결정 F1·F3", "12310002": "FAR 자산분류 (취득원가 계정과 동일 비율)", "52010002": "FAR 자산분류 (상각 스케줄)",
    "12110004": "FAR 자산분류(장비 / 가구·비품)", "12110005": "FAR 자산분류(가구·비품 / 기타)", "12310004": "FAR 자산분류(장비 / 가구·비품)",
    "12310005": "FAR 자산분류(가구·비품 / 기타)", "52010004": "FAR 자산분류(장비 / 가구·비품)", "52010005": "FAR 자산분류(가구·비품 / 기타)",
    "20125001": "고객유형(법인 → 20150100100 / 개인·개인사업자 → 20150300600) - 결정 P1", "20125003": "고객유형(법인 / 개인)",
    "20122001": "고객유형(법인 / 개인) - 결정 P1", "51060001": "이자 원천 예금의 고객유형(정기예금 분할과 동일)",
    "20990030": "수취인 거주성(거주자 199 / 비거주자 399) + 원천징수 조문", "20990040": "TCS 유형(LRS 교육·의료 / LRS 기타 / 해외여행 패키지)",
    "52032481": "경비(EP) 코드(노무 / 감사 / 법률 / 전문가) - 결정 X2·X6",
    "54010091": "오류: 비용 코드(5계열)가 부채 충당부채 20990000000에도 연결됨 → 비용 50690000000만, 부채측은 현행 부채 계정에서 이행",
    "53090014": "0930 GBS 중복코드: 매매손실 50430201000 / 처분손실 50440000000(상위 계정) → 처분손실 leaf 지정 필요",
}
LEAF_FIX = {  # 상위(비기장) 계정에 연결된 현행 코드 → 권고 leaf
    "10610000000": ("10610100000", "SOFTWARE L4 → leaf L5"), "10650000000": ("10650100000", "OTHERS L4 → leaf L5"),
    "20181100000": ("20181100100", "송금 미지급 L5 분리(TB_AP_ITEM 유형별: 타발 보류 100 / FDI 초과 200 / 반송 300)"),
    "31130000000": ("31130100000", "AFS Reserve: 평가이익 100 / 평가손실 200 (부호 기준 분할)"),
    "49900000000": ("49910000000", "손익이체 합계 → leaf 49910000000"), "59900000000": ("59910000000", "당기순이익 이체 잔액 ₹7.80 crore: 결산 마감계정 → leaf 59910000000 (이행일 기준 결산 후 0 확인)"),
    "50440000000": ("50440100000", "처분손실 L3 → leaf (채권 100 / 주식 200)"), "50930020000": ("50930020100", "급여 L4 → 기본급 leaf (수당 200·300 분리)"),
    "50930170000": ("50930170100", "퇴직·법정 인건비 → Gratuity 100 / 연차 200 / EPF 300 / ESIC 400"),
    "60170000000": ("60170300000", "약정 L3 → leaf (성격 확인: IRS 100 / 신용공여 약정 300)"),
}


def _map_rows():
    out = []
    for g in sorted(_cur):
        tg = _cur[g]
        kind = "1:1" if len(tg) == 1 else "1:N 분할"
        for c in tg:
            a = BY[c]
            n_src = sum(1 for k, v in _cur.items() if c in v)
            k2 = kind if n_src == 1 else (kind + " · N:1 통합")
            issue, sev, sug = "", "OK", c
            if c in TEMP_USE:
                issue, sev = f"대상 계정 상태 '비활성'이나 결정 L3상 {TEMP_USE[c]}까지 사용 — 기한부 사용으로 수정 필요", "수정"
            elif a["post"] != "사용":
                issue, sev, sug = f"대상 계정이 '{a['status'] or a['post']}' — 잔액 {'있음' if a['bal'] else '없음'}: 현행 계정 폐지(이행 제외)로 명시 권고", "수정", "이행 제외"
            elif c in PARENTS:
                lf = LEAF_FIX.get(c, (sorted(_active_children(c))[0], "상위 계정 → leaf 지정"))
                issue, sev, sug = f"상위(비기장) 계정에 연결 → {lf[1]}", "수정", lf[0]
            if g in SPLIT_RULE:
                issue = (issue + " / " if issue else "") + "분할 기준: " + SPLIT_RULE[g]
                if sev == "OK":
                    sev = "오류" if g == "54010091" else "분할규칙"
            out.append({"cur": g, "code": c, "name": acct_name(c), "lv": a["lv"], "status": a["status"], "post": a["post"], "bal": a["bal"],
                        "kind": k2, "sev": sev, "issue": issue, "target": sug})
    return out


MAPPING = _map_rows()

# ---------------------------------------------------------------------------
# 06 파일 검토 의견 (추가·수정·삭제)
# ---------------------------------------------------------------------------
REVIEW = [
    ("수정", "높음", "20510000000 (OTHER LIABILITIES L3)", "결정 G3(20501000000 → 20510000000 재채번)이 계정표에 미반영. 하위 L4 20510100000~500000 5개가 상위 없이 존재(TDS·GST·직원부채 포함)", "계정표에 20510000000 행 추가, 20501000000은 '철회'로 표시", "G3"),
    ("수정", "높음", "10380101000 / 10380102000 (일반·특정 충당금)", "상태가 '비활성·미사용'인데 결정 L3는 2027-03-31까지 사용. 현행 10690001·10690004 잔액 ₹-1.59 crore가 미사용 계정으로 이관되어 이행 시 갈 곳이 없음", "상태 '기한부 사용(~2027-03-31)', ECL 시행일에 Stage 1~3(10380103000~105000)으로 이관 전표 + 전환 조정 31110800000", "L3"),
    ("수정", "중간", "현행 20401001~3·42020801·48010020·53010014·51290001·53090092", "잔액 0인 현행 계정을 비활성 11자리(STATE FUNDS·AOB·TOKEN·갑기금 이자)에 연결", "현행 이행표에서 '이행 제외(현행 폐지)'로 명시, 대상 칸 비움", "G1·G7"),
    ("수정", "높음", "현행 54010091", "비용(5계열) 코드가 부채 20990000000(PROVISION FOR OTHERS)과 비용 50690000000에 동시 연결. 0930 GBS 열 중복을 그대로 가져온 것으로 보임", "비용 50690000000만 연결. 부채 20990000000은 현행 부채측 충당부채 코드에서 이행", "G4"),
    ("수정", "중간", "현행 53090014", "매매손실 50430201000과 처분손실 50440000000(상위 계정)에 동시 연결(0930 GBS 중복)", "성격별 1개 leaf로 지정(처분손실이면 50440100000/200000)", "G4"),
    ("수정", "중간", "상위 계정 연결 10건", "현행 코드가 하위 계정이 있는 상위 계정에 연결(10610000000·20181100000·31130000000·59900000000 등). G2 '기장은 leaf'와 불일치", "19.계정이행맵핑 시트의 '권고 대상' 열 leaf로 정정", "G2"),
    ("수정", "중간", "1:N 분할 18건", "현행 1개 → 차세대 여러 계정(예: 20125001 → 법인 당좌/개인 당좌, 20990040 → TCS 3종). 이행 프로그램에 분할 키가 없으면 잔액 배분 불가", "건별 분할 기준(고객유형·FAR 자산분류·거주성·TCS 유형·EP 코드)을 이행 규칙으로 확정", "P1·F1·T1·X6"),
    ("추가", "중간", "현행 CoA 전체 목록", "06 파일은 잔액 있는 현행 계정(145개, 합계 0 ✓)과 연결된 697개만 보여줌. 잔액 0인 현행 계정 중 누락 여부는 확인 불가", "현행 GBS 8자리 계정 마스터 전체 제공 → 미연결 코드 0건 확인", "G1"),
    ("수정", "낮음", "당발 OUR 수수료", "결정 C1: OUR 정액 40570060000, 중계은행 수수료 고객 회수 40570110000(과세). 송금 화면 수수료 엔진에 OUR 부과 항목이 아직 없음", "송금 화면에 OUR 수수료 행 추가(정액/실비 회수 구분)", "C1"),
    ("확인", "낮음", "ECL 시행일", "결정 L3는 2027-04-01 ECL 전환 전제. RBI ECL 최종 지침의 적용일·대상(외국은행 지점 포함 여부) 확인 필요", "최종 고시 확인 후 TEMP_USE 기한 확정", "L3·L6"),
    ("반영", "-", "시스템 기장 계정 (당 설계)", "분개·화면이 쓰던 상위·철회 계정 8종을 06 파일 leaf로 변경: RBI 10104010100, 에스크로 20150400100, TCS 20510400201, GST L6(수수료 x01 / 환전 x02), 내국 LC → 60120100000+속성, 금융·이행보증 → 60130400000+속성, SNRR(FPI) → 20150100300+속성", "적용 완료 (R-ACC-01 leaf 검증 통과)", "G2·L7·P5·T7"),
]


# ---------------------------------------------------------------------------
# 5. 업무별 예시 분개
#    leg = (차/대, 계정, 통화, 금액, 금액기준, 적요)
#    D/C = 차변/대변, OB+/OB- = 난외 증가/감소(대차검증 제외)
#    환전은 통화별로 대차를 맞추고 11200000000 포지션으로 연결
# ---------------------------------------------------------------------------
POS, NOSTRO, RBI = "11200000000", "10105020000", "10104010100"   # RBI 당좌(CRR·RTGS 결제) L5 - 06 파일
# GST 매출세액: 세목 L5 × L6(01 수수료 / 02 환전) - 결정 T7
GSTA = {"CGST": ("20510500101", "20510500102"), "SGST": ("20510500201", "20510500202"), "IGST": ("20510500301", "20510500302")}


def _gst(fee, pos, cv=0.0):
    """세액은 과세표준 합계로 계산(확인서와 동일)하고 L6(수수료분/환전분)로 배분"""
    base = r2(fee + cv)
    g = r2(base * 0.18)
    heads = [("CGST", r2(g / 2)), ("SGST", r2(g - r2(g / 2)))] if pos == "CS" else [("IGST", g)]
    out = []
    for h, amt in heads:
        fx = r2(amt * cv / base) if base else 0.0
        if amt - fx:
            out.append(("C", GSTA[h][0], "INR", r2(amt - fx), h, f"{h} (수수료분)"))
        if fx:
            out.append(("C", GSTA[h][1], "INR", fx, h, f"{h} (환전분 32(2)(b))"))
    return out


def _gsum(legs):
    return r2(sum(l[3] for l in legs))


J = []  # (업무, 시나리오ID, 시나리오, 화면, 단계, 일자, 근거, legs)


def je(biz, sid, title, scr, step, dt, ref, legs):
    J.append((biz, sid, title, scr, step, dt, ref, legs))


# ---- 당발송금 -------------------------------------------------------------
fee = [("C", "40570010000", "INR", 11052.50, "COMM", "송금수수료"), ("C", "40570030000", "INR", 1000.00, "CABLE", "전신료")]
g = _gst(11052.50 + 1000, "CS", gst_fx_value(8842000.00))
tot = r2(8842000.00 + _gsum(fee) + _gsum(g))
je("당발송금", "OUT-01", "수입대금 선급 USD 100,000 (법인 CA 출금, OUR)", "RT13", "① 출금·환전·송신", "2026-10-31", "R-IMP, GST 32(2)(b)", [
    ("D", "20150100100", "INR", tot, "TOTAL_DR", "고객 CA 출금(원금+수수료+GST)"),
    ("C", POS, "INR", 8842000.00, "INR_AMT", "환전 INR (88.42)"),
    *fee, *g,
    ("D", POS, "USD", 100000.00, "FCY_AMT", "환전 USD"),
    ("C", NOSTRO, "USD", 100000.00, "FCY_AMT", "Nostro 출금 (pacs.008)"),
])
je("당발송금", "OUT-01", "수입대금 선급 USD 100,000 (법인 CA 출금, OUR)", "RT13", "② 해외은행 수수료 청구(OUR)", "2026-11-05", "MT191/camt.106", [
    ("D", "50510300000", "USD", 25.00, "CORR_CHG", "중계·수취은행 수수료"),
    ("C", NOSTRO, "USD", 25.00, "CORR_CHG", "Nostro 차감"),
])
fee = [("C", "40570010000", "INR", 500.00, "COMM", "송금수수료"), ("C", "40570030000", "INR", 1000.00, "CABLE", "전신료")]
g = _gst(1500, "CS", gst_fx_value(3978900.00))
tcs = ("C", "20510400201", "INR", 198945.00, "TCS", "TCS LRS 교육·의료")
tot = r2(3978900.00 + _gsum(fee) + _gsum(g) + 198945.00)
je("당발송금", "OUT-02", "LRS 유학비 USD 45,000 (개인 SB, TCS)", "RT12", "① 출금·환전·TCS", "2026-10-31", "R-LRS, 206C(1G)", [
    ("D", "20150300100", "INR", tot, "TOTAL_DR", "개인 SB 출금"), ("C", POS, "INR", 3978900.00, "INR_AMT", "환전 INR"),
    *fee, *g, tcs,
    ("D", POS, "USD", 45000.00, "FCY_AMT", ""), ("C", NOSTRO, "USD", 45000.00, "FCY_AMT", "Nostro 출금"),
])
je("당발송금", "OUT-02", "LRS 유학비 USD 45,000 (개인 SB, TCS)", "RT12", "② TCS 납부(익월 7일)", "2026-11-07", "27EQ", [
    ("D", "20510400201", "INR", 198945.00, "TCS", "TCS 예수금 정리"), ("C", RBI, "INR", 198945.00, "TCS", "정부 납부(RBI 경유)"),
])
je("당발송금", "OUT-03", "EEFC 출금 송금 USD 20,000 (무환전)", "RT11", "① EEFC 출금·송신 + 환전대체수수료", "2026-10-31", "R-FX, 40570100000", [
    ("D", "20150100200", "USD", 20000.00, "FCY_AMT", "EEFC 출금"), ("C", NOSTRO, "USD", 20000.00, "FCY_AMT", "Nostro 출금"),
    ("D", "20150100100", "INR", 2360.00, "TOTAL_DR", "수수료 INR 계좌 출금"),
    ("C", "40570100000", "INR", 1000.00, "COMM", "In lieu of exchange (요율표 예시)"),
    ("C", "40570030000", "INR", 1000.00, "CABLE", "전신료"), *_gst(2000.00, "CS"),
])

# 결제경로별 결제계정 차이 (R-RTG-07): 고객 측 분개는 같고 외화 레그의 대변 계정만 달라짐
fee = [("C", "40570010000", "INR", 5000.00, "COMM", "송금수수료"), ("C", "40570030000", "INR", 1000.00, "CABLE", "전신료")]
g = _gst(6000, "CS", gst_fx_value(r2(250000 * 88.42)))
je("당발송금", "OUT-04", "Direct(본점): 수취인 계좌가 본점 서울 USD 250,000", "RT11", "① 출금·환전·본점 앞 pacs.008", "2026-11-04", "R-RTG-01 ③, R-RTG-07", [
    ("D", "20150100100", "INR", r2(r2(250000 * 88.42) + 6000 + _gsum(g)), "TOTAL_DR", "고객 CA 출금"),
    ("C", POS, "INR", r2(250000 * 88.42), "INR_AMT", ""), *fee, *g,
    ("D", POS, "USD", 250000.00, "FCY_AMT", ""),
    ("C", "21010000000", "USD", 250000.00, "FCY_AMT", "을A Due to HO (본점이 수취인 입금) → IO21 정산 대상"),
])
je("당발송금", "OUT-05", "Cover: 수취은행 직접 pacs.008 + Nostro 앞 pacs.009 COV USD 620,000", "RT11/SW11", "① Nostro 결제(pacs.009 COV 기준)", "2026-11-04", "R-RTG-02/04", [
    ("D", POS, "USD", 620000.00, "FCY_AMT", "고객 측 INR 분개는 OUT-01과 동일(생략)"),
    ("C", NOSTRO, "USD", 620000.00, "FCY_AMT", "Nostro는 pacs.009 COV로 1회 차기. pacs.008은 정보 전달(결제 없음)"),
])

# ---- 타발송금 -------------------------------------------------------------
g = _gst(250, "IG", gst_fx_value(7639488.00))
je("타발송금", "IN-01", "수출선수금 USD 86,400 → 법인 CA (타주 고객 IGST)", "RT21", "① 수취·환전·입금", "2026-10-31", "R-IN, IRM", [
    ("D", NOSTRO, "USD", 86400.00, "FCY_AMT", "Nostro 입금 (pacs.008)"), ("C", POS, "USD", 86400.00, "FCY_AMT", ""),
    ("D", POS, "INR", 7639488.00, "INR_AMT", "환전 INR (88.42)"),
    ("C", "20150100100", "INR", r2(7639488.00 - 250 - _gsum(g)), "NET_CR", "고객 CA 입금(수수료·GST 차감)"),
    ("C", "40570020000", "INR", 250.00, "COMM", "타발수수료"), *g,
])
je("타발송금", "IN-02", "외화 그대로 EEFC 입금 USD 30,000", "RT21", "① 수취·EEFC 입금", "2026-10-31", "R-IN", [
    ("D", NOSTRO, "USD", 30000.00, "FCY_AMT", "Nostro 입금"), ("C", "20150100200", "USD", 30000.00, "FCY_AMT", "EEFC 입금(무환전)"),
])
fee = [("C", "40570020000", "INR", 500.00, "COMM", "중계 타발수수료"), ("C", "40570070000", "INR", 25.00, "RTGS_CHG", "RTGS 수수료")]
g = _gst(525, "CS", gst_fx_value(4599400.00))
pay = r2(4599400.00 - _gsum(fee) - _gsum(g))
je("타발송금", "IN-03", "비계좌 수취인 중계 USD 52,000 → RTGS", "RT21→RT22→SF11", "① 수취·보류(10-29)", "2026-10-29", "R-SUS-01", [
    ("D", NOSTRO, "USD", 52000.00, "FCY_AMT", "Nostro 입금"), ("C", "20181100100", "USD", 52000.00, "FCY_AMT", "타발 미지급(외화) 보류"),
])
je("타발송금", "IN-03", "비계좌 수취인 중계 USD 52,000 → RTGS", "RT21→RT22→SF11", "② 환전(10-31, 88.45)", "2026-10-31", "R-SUS, 레그별 딜", [
    ("D", "20181100100", "USD", 52000.00, "FCY_AMT", "외화 보류 차감"), ("C", POS, "USD", 52000.00, "FCY_AMT", ""),
    ("D", POS, "INR", 4599400.00, "INR_AMT", ""), ("C", "20181100100", "INR", 4599400.00, "INR_AMT", "INR 보류 계상"),
])
je("타발송금", "IN-03", "비계좌 수취인 중계 USD 52,000 → RTGS", "RT21→RT22→SF11", "③ RTGS 지급", "2026-10-31", "SFMS pacs.008", [
    ("D", "20181100100", "INR", 4599400.00, "INR_AMT", "INR 보류 정리"),
    ("C", RBI, "INR", pay, "NET_CR", "RTGS 지급(UTR)"), *fee, *g,
])
je("타발송금", "IN-03", "비계좌 수취인 중계 USD 52,000 → RTGS", "RT21→RT22→SF11", "④ (예외) RTGS 반송", "2026-11-01", "pacs.004", [
    ("D", RBI, "INR", pay, "NET_CR", "반송 입금"), ("C", "20181100300", "INR", pay, "NET_CR", "반송 미지급 복원 → 재지급/반환"),
])

# ---- FDI ------------------------------------------------------------------
conv_fcy, conv_inr = legs_fixed(250000000.00, 88.42)
exc = r2(2850000.00 - conv_fcy)
g = _gst(7500, "IG", gst_fx_value(conv_inr))
je("FDI", "FDI-01", "FDI 약정 INR 25 crore, 수취 USD 2,850,000 (초과 반환)", "FD11", "① 수취·약정분 환전·초과분 분리", "2026-09-01", "R-FDI-06/07", [
    ("D", NOSTRO, "USD", 2850000.00, "FCY_AMT", "Nostro 입금"),
    ("C", POS, "USD", conv_fcy, "CONV_FCY", f"약정분만 환전 USD {conv_fcy:,.2f}"),
    ("C", "20181100200", "USD", exc, "EXCESS_FCY", "초과분 환전 금지, 반환대기"),
    ("D", POS, "INR", conv_inr, "INR_AMT", "환전 INR (88.42)"),
    ("C", "20150100100", "INR", conv_inr, "INR_AMT", "피투자법인 CA (Share application money 관리)"),
])
je("FDI", "FDI-01", "FDI 약정 INR 25 crore, 수취 USD 2,850,000 (초과 반환)", "FD11", "② 수수료·GST (별도 출금)", "2026-09-01", "GST s.12(12)", [
    ("D", "20150100100", "INR", r2(7500 + _gsum(g)), "TOTAL_DR", "피투자법인 CA"), ("C", "40570020000", "INR", 7500.00, "COMM", "FDI 수수료"), *g,
])
je("FDI", "FDI-01", "FDI 약정 INR 25 crore, 수취 USD 2,850,000 (초과 반환)", "FD12", "③ 초과분 원통화 반환(5영업일, BEN)", "2026-09-03", "R-FDI-07, FDI_EXCESS_RETURN_DAYS", [
    ("D", "20181100200", "USD", exc, "EXCESS_FCY", "반환대기 정리"),
    ("C", NOSTRO, "USD", r2(exc - 25), "NET_FCY", "송금인에게 반환(pacs.008)"),
    ("C", "40570090000", "USD", 25.00, "COMM", "반환 수수료(BEN 차감, 외화수익). 해외 송금인 대상 → GST 영세(세무 확인)"),
])
je("FDI", "FDI-02", "FDI 미배정 반환 INR 5 crore (60일 경과)", "FD12", "① 반환 송금", "2026-10-20", "R-FDI-05, 75일", [
    ("D", "20150100100", "INR", r2(50000000 + 6000 + r2((6000 + gst_fx_value(50000000)) * 0.18)), "TOTAL_DR", "피투자법인 CA"),
    ("C", POS, "INR", 50000000.00, "INR_AMT", ""), ("C", "40570010000", "INR", 5000.00, "COMM", ""),
    ("C", "40570030000", "INR", 1000.00, "CABLE", ""), *_gst(6000, "IG", gst_fx_value(50000000)),
    ("D", POS, "USD", 565482.92, "FCY_AMT", ""), ("C", NOSTRO, "USD", 565482.92, "FCY_AMT", "원송금인 반환"),
])
esc_fcy, esc_inr = legs_fixed(100000000.00, 88.40)
je("FDI", "FDI-03", "FDI 에스크로 INR 10 crore (지분 인수)", "FD13", "① 수취·환전·에스크로 예치", "2026-10-15", "FEM(NDI) 에스크로", [
    ("D", NOSTRO, "USD", esc_fcy, "FCY_AMT", ""), ("C", POS, "USD", esc_fcy, "FCY_AMT", ""),
    ("D", POS, "INR", esc_inr, "INR_AMT", ""), ("C", "20150400100", "INR", esc_inr, "INR_AMT", "에스크로 예치(무이자)"),
])
je("FDI", "FDI-03", "FDI 에스크로 INR 10 crore (지분 인수)", "FD13", "② 조건 충족 → 매도인 RTGS 지급", "2026-11-20", "에스크로 약정", [
    ("D", "20150400100", "INR", esc_inr, "INR_AMT", "에스크로 해지"), ("C", RBI, "INR", esc_inr, "INR_AMT", "매도인 RTGS"),
])

# ---- 환전 -----------------------------------------------------------------
g = _gst(0, "CS", gst_fx_value(4419000.00))
je("환전", "FX-01", "EEFC USD 50,000 → INR CA (88.38)", "FX11", "① 계좌 간 환전", "2026-10-31", "R-FX", [
    ("D", "20150100200", "USD", 50000.00, "FCY_AMT", "EEFC 출금"), ("C", POS, "USD", 50000.00, "FCY_AMT", ""),
    ("D", POS, "INR", 4419000.00, "INR_AMT", ""), ("C", "20150100100", "INR", 4419000.00, "INR_AMT", "CA 입금"),
    ("D", "20150100100", "INR", _gsum(g), "GST", "환전 GST 출금"), *g,
])
g = _gst(0, "CS", gst_fx_value(104340.00))
je("환전", "FX-02", "외화현찰 매입 USD 1,200 (워크인, 86.95)", "FX11", "① 현찰 매입", "2026-10-31", "Encashment Certificate", [
    ("D", "10103000000", "USD", 1200.00, "FCY_AMT", "외국통화 시재"), ("C", POS, "USD", 1200.00, "FCY_AMT", ""),
    ("D", POS, "INR", 104340.00, "INR_AMT", ""), ("C", "10101000000", "INR", r2(104340 - _gsum(g)), "NET_CR", "INR 현금 지급(GST 차감)"), *g,
])
je("환전", "FX-03", "지점 자기계정 USD 1,000,000 매도 (은행간, 88.40)", "FX11(BRANCH)", "① 결제일(T+2) 결제", "2026-11-03", "R-FX-08, CCIL", [
    ("D", RBI, "INR", 88400000.00, "INR_AMT", "INR 수령(RBI 당좌)"), ("C", POS, "INR", 88400000.00, "INR_AMT", ""),
    ("D", POS, "USD", 1000000.00, "FCY_AMT", ""), ("C", NOSTRO, "USD", 1000000.00, "FCY_AMT", "USD 인도"),
])
je("환전", "FX-04", "일말 포지션 손익 대체 (IN-01 매입분을 88.50에 커버)", "FX21", "① 커버 거래", "2026-10-31", "", [
    ("D", RBI, "INR", 7646400.00, "INR_AMT", "은행간 매도 INR 수령"), ("C", POS, "INR", 7646400.00, "INR_AMT", ""),
    ("D", POS, "USD", 86400.00, "FCY_AMT", ""), ("C", NOSTRO, "USD", 86400.00, "FCY_AMT", ""),
])
je("환전", "FX-04", "일말 포지션 손익 대체 (IN-01 매입분을 88.50에 커버)", "FX21", "② EOD 포지션 손익 대체", "2026-10-31", "USD 포지션 0, INR 잔액 = 손익", [
    ("D", POS, "INR", 6912.00, "PNL", "INR 포지션 잔액(7,646,400 - 7,639,488)"), ("C", "40760100000", "INR", 6912.00, "PNL", "외환매매이익"),
])

# ---- Vostro ---------------------------------------------------------------
je("Vostro", "VS-01", "SRVA: 인도 수입자 INR 결제 → 해외 수출자 은행 Vostro", "VS11", "① SRVA 입금", "2026-10-31", "RBI SRVA 체계", [
    ("D", "20150100100", "INR", 5000000.00, "INR_AMT", "수입자 CA 출금"), ("C", "20150200100", "INR", 5000000.00, "INR_AMT", "SRVA 입금 → MT950"),
])
je("Vostro", "VS-01", "SRVA: 해외 수입자 결제 → 인도 수출자", "VS11", "② SRVA 출금·RTGS", "2026-10-31", "R-VS-06", [
    ("D", "20150200100", "INR", 3000000.00, "INR_AMT", "SRVA 출금"), ("C", RBI, "INR", 3000000.00, "INR_AMT", "수출자 거래은행 RTGS"),
])
je("Vostro", "VS-01", "SRVA 계좌유지·MT950 수수료", "VS31", "③ 월 수수료", "2026-10-31", "GST s.13(8)(a)", [
    ("D", "20150200100", "INR", 5900.00, "TOTAL_DR", "SRVA 출금"), ("C", "40592030000", "INR", 5000.00, "COMM", "계좌유지·명세서"), *_gst(5000.00, "CS"),
])
je("Vostro", "VS-02", "본점 INR Vostro: 본점이 USD 1,000,000 송금 → INR 예치", "VS11", "① 본점 Vostro 입금", "2026-10-31", "본지점 환율", [
    ("D", NOSTRO, "USD", 1000000.00, "FCY_AMT", "본점 → Nostro"), ("C", POS, "USD", 1000000.00, "FCY_AMT", ""),
    ("D", POS, "INR", 88420000.00, "INR_AMT", ""), ("C", "21040000000", "INR", 88420000.00, "INR_AMT", "본점 INR Vostro 입금 → MT950"),
])
je("Vostro", "VS-02", "본점 INR Vostro → 한국계 법인 인도 자회사 지급", "VS11/RT21", "② 본점 Vostro 출금", "2026-10-31", "R-VS-06, IRM", [
    ("D", "21040000000", "INR", 20000000.00, "INR_AMT", "본점 Vostro 출금(수수료 없음)"), ("C", "20150100100", "INR", 20000000.00, "INR_AMT", "자회사 CA 입금"),
])

# ---- 본지점·RA ------------------------------------------------------------
je("본지점·RA", "IO-01", "갑기금 USD 10,000,000 수령", "IO11", "① 영업기금 수령", "2026-10-01", "R-IO 갑", [
    ("D", NOSTRO, "USD", 10000000.00, "FCY_AMT", ""), ("C", "31140100000", "USD", 10000000.00, "FCY_AMT", "갑기금(무이자, 정산 제외)"),
])
je("본지점·RA", "RA-01", "본점 LC 상환청구(MT742) USD 250,000", "RA11", "① 대지급(인도 영업일 3일 내)", "2026-10-29", "URR 725, R-RA", [
    ("D", "11010000000", "USD", 250100.00, "FCY_AMT", "을A Due from HO (원금+상환수수료)"),
    ("C", NOSTRO, "USD", 250000.00, "FCY_AMT", "청구은행 지급(pacs.009)"),
    ("C", "40550010000", "USD", 100.00, "COMM", "상환수수료(본점 부담)"),
])
je("본지점·RA", "RA-01", "본점 LC 상환청구(MT742) USD 250,000", "IO21", "② 정산배치 입금(금요일)", "2026-10-30", "건수·금액·주기 트리거", [
    ("D", NOSTRO, "USD", 250100.00, "FCY_AMT", "본점 정산 입금"), ("C", "11010000000", "USD", 250100.00, "FCY_AMT", "을A 정산"),
])
je("본지점·RA", "RA-02", "휴일차이 이자 USD 412.33 (USD 300 초과 → SETTLE)", "CM21/IO21", "① 정산 포함", "2026-10-30", "IO_FLOAT_ABSORB_USD", [
    ("D", "11010000000", "USD", 412.33, "INT", "을A 채권"), ("C", "40620000000", "USD", 412.33, "INT", "을A 이자수익"),
])
je("본지점·RA", "RA-02", "휴일차이 이자 USD 180.00 (USD 300 이하 → ABSORB)", "CM21", "① 자체 흡수", "2026-10-30", "IO_FLOAT_ABSORB_USD", [])
r = r2(5000000 * 0.0485 * 92 / 360)
m1 = r2(5000000 * 0.0485 * 30 / 360)
je("본지점·RA", "IO-02", "을B 본점 차입 USD 5,000,000 · 92일 · 4.85%", "IO31", "① 차입", "2026-10-01", "MT320, R-IO 을B", [
    ("D", NOSTRO, "USD", 5000000.00, "FCY_AMT", ""), ("C", "21020000000", "USD", 5000000.00, "FCY_AMT", "을B 차입"),
])
je("본지점·RA", "IO-02", "을B 본점 차입 USD 5,000,000 · 92일 · 4.85%", "IO31", "② 월말 이자 발생(30일)", "2026-10-31", "ACT/360", [
    ("D", "51130000000", "USD", m1, "INT", "을B 이자비용"), ("C", "20410080000", "USD", m1, "INT", "본지점 미지급이자"),
])
je("본지점·RA", "IO-02", "을B 본점 차입 USD 5,000,000 · 92일 · 4.85%", "IO31", "③ 만기 상환(잔여 62일 이자 포함)", "2027-01-01", "", [
    ("D", "21020000000", "USD", 5000000.00, "FCY_AMT", ""), ("D", "20410080000", "USD", m1, "INT", "기발생 미지급이자"),
    ("D", "51130000000", "USD", r2(r - m1), "INT", "잔여 기간 이자"), ("C", NOSTRO, "USD", r2(5000000 + r), "FCY_AMT", "원리금 상환"),
])

# ---- 무역금융 -------------------------------------------------------------
da_int = r2(200000 * 0.061 * 90 / 360)
da_net = r2(200000 - da_int)
da_inr = r2(da_net * 88.42)
je("무역금융", "TF-01", "수출 D/A 어음 매입 USD 200,000 · 90일 · 6.10%(ARR+스프레드)", "TF11", "① 매입·환전·순지급 A/P", "2026-10-31", "R-TF-09, RBI 검토 #23", [
    ("D", "10310200000", "USD", 200000.00, "FACE", "D/A 기한부 매입어음(액면)"),
    ("C", "20420080000", "USD", da_int, "DISC_INT", "선취 할인이자(미경과)"),
    ("C", POS, "USD", da_net, "NET_FCY", ""),
    ("D", POS, "INR", da_inr, "INR_AMT", "88.42"), ("C", "20440800000", "INR", da_inr, "INR_AMT", "순지급액 A/P"),
    ("D", "20150100100", "INR", 2360.00, "TOTAL_DR", "매입수수료 출금"), ("C", "40560020000", "INR", 2000.00, "COMM", "D/A 매입·추심 수수료"), *_gst(2000.00, "CS"),
])
je("무역금융", "TF-01", "수출 D/A 어음 매입", "SF11", "② A/P 지급(수출자 CA)", "2026-10-31", "", [
    ("D", "20440800000", "INR", da_inr, "INR_AMT", "A/P 정리"), ("C", "20150100100", "INR", da_inr, "INR_AMT", "수출자 CA 입금"),
])
je("무역금융", "TF-01", "수출 D/A 어음 매입", "EOD", "③ 일일 이자 인식(1일분)", "2026-11-01", "EOD ①, NPA 시 중단", [
    ("D", "20420080000", "USD", r2(da_int / 90), "DISC_INT", "미경과 → 경과"), ("C", "40340200000", "USD", r2(da_int / 90), "DISC_INT", "기한부 매입어음 이자"),
])
je("무역금융", "TF-01", "수출 D/A 어음 매입", "TF81", "④ 만기 회수", "2027-01-29", "", [
    ("D", NOSTRO, "USD", 200000.00, "FACE", "수입자 은행 결제"), ("C", "10310200000", "USD", 200000.00, "FACE", "매입어음 정리"),
])
ts = r2(10000000 * 0.0875 * 60 / 365)
je("무역금융", "TF-02", "TReDS 매출채권 할인 INR 1 crore · 60일 · 8.75%", "TF11", "① 할인·A/P", "2026-10-31", "TReDS Directions 2026", [
    ("D", "10320300000", "INR", 10000000.00, "FACE", "국내 매출채권 할인"),
    ("C", "20420090000", "INR", ts, "DISC_INT", "선취 할인이자"), ("C", "20440800000", "INR", r2(10000000 - ts), "NET_CR", "순지급액 A/P"),
])
je("무역금융", "TF-02", "TReDS 매출채권 할인", "SF11", "② TReDS 결제(RTGS)", "2026-10-31", "", [
    ("D", "20440800000", "INR", r2(10000000 - ts), "NET_CR", ""), ("C", RBI, "INR", r2(10000000 - ts), "NET_CR", "판매기업 RTGS"),
])
je("무역금융", "TF-02", "TReDS 매출채권 할인", "TF81", "③ 만기 회수", "2026-12-30", "", [
    ("D", RBI, "INR", 10000000.00, "FACE", "구매기업 결제"), ("C", "10320300000", "INR", 10000000.00, "FACE", ""),
    ("D", "20420090000", "INR", ts, "DISC_INT", "잔여 미경과 일괄(예시)"), ("C", "40350300000", "INR", ts, "DISC_INT", "할인이자"),
])
je("무역금융", "TF-03", "수출 이자보전 청구 INR 34,000", "TF81", "① 청구", "2026-10-31", "EPM Niryat Protsahan", [
    ("D", "10490910000", "INR", 34000.00, "INT", "정부 미수"), ("C", "40340500000", "INR", 34000.00, "INT", "이자보전 수익"),
])
je("무역금융", "TF-04", "할인어음 NPA 충당금", "TF81", "① 충당금 전입", "2026-10-31", "IRACP", [
    ("D", "50640000000", "INR", 1768400.00, "PROV", "대손충당금 전입(예시 10%)"), ("C", "10380102000", "INR", 1768400.00, "PROV", "개별 충당금"),
])

# ---- 신용장·보증 ----------------------------------------------------------
lc_fee = 7500.00
je("신용장·보증", "LC-01", "수입 일람불 LC USD 300,000 개설 (증거금 10%)", "LC11", "① 개설·증거금·수수료", "2026-10-15", "R-LC", [
    ("OB+", "60120100000", "USD", 300000.00, "LC_AMT", "LC 미결제 난외 증가"),
    ("D", "20150100100", "INR", 2652600.00, "MARGIN", "증거금 출금"), ("C", "20260100000", "INR", 2652600.00, "MARGIN", "수입보증금(일람불)"),
    ("D", "20150100100", "INR", r2(lc_fee * 1.18), "TOTAL_DR", ""), ("C", "40520020000", "INR", lc_fee, "COMM", "개설수수료"), *_gst(lc_fee, "CS"),
])
je("신용장·보증", "LC-01", "수입 일람불 LC USD 300,000", "LC11", "② 서류 결제(대지급)", "2026-11-10", "UCP 600", [
    ("D", "10360100000", "USD", 300000.00, "FCY_AMT", "LC 대지급"), ("C", NOSTRO, "USD", 300000.00, "FCY_AMT", "네고은행 지급"),
    ("OB-", "60120100000", "USD", 300000.00, "LC_AMT", "난외 감소"),
])
je("신용장·보증", "LC-01", "수입 일람불 LC USD 300,000", "LC11", "③ 고객 상환(증거금 충당)", "2026-11-10", "", [
    ("D", "20260100000", "INR", 2652600.00, "MARGIN", "증거금 충당"), ("D", "20150100100", "INR", 23873400.00, "INR_AMT", "잔액 CA 출금"),
    ("C", POS, "INR", 26526000.00, "INR_AMT", "88.42"),
    ("D", POS, "USD", 300000.00, "FCY_AMT", ""), ("C", "10360100000", "USD", 300000.00, "FCY_AMT", "대지급 회수"),
])
je("신용장·보증", "LC-02", "내국신용장 INR 2 crore 개설 (SFMS IFN 700)", "LC11", "① 개설", "2026-10-20", "IFN 700", [
    ("OB+", "60120100000", "INR", 20000000.00, "LC_AMT", "내국 LC 난외 증가 (속성 lc_scope=INLAND, 결정 L7)"),
])
bg_c = 750000.00
je("신용장·보증", "BG-01", "금융보증 INR 5 crore · 1년 · 1.5%", "BG11", "① 발행·보증료 선수", "2026-10-31", "NFB Directions, FEM(Guarantees)", [
    ("OB+", "60130400000", "INR", 50000000.00, "BG_AMT", "보증 난외 (속성 bg_nature=FINANCIAL, CCF 100%)"),
    ("D", "20150100100", "INR", r2(bg_c * 1.18), "TOTAL_DR", ""), ("C", "20420170000", "INR", bg_c, "COMM", "보증료 선수(기간안분)"), *_gst(bg_c, "CS"),
])
je("신용장·보증", "BG-01", "금융보증 INR 5 crore", "EOD", "② 월 보증료 안분", "2026-11-30", "", [
    ("D", "20420170000", "INR", 62500.00, "COMM", ""), ("C", "40510060000", "INR", 62500.00, "COMM", "보증료 수익"),
])
je("신용장·보증", "BG-02", "이행보증 INR 1 crore 발행", "BG11", "① 발행", "2026-10-31", "", [
    ("OB+", "60130400000", "INR", 10000000.00, "BG_AMT", "보증 난외 (속성 bg_nature=PERFORMANCE, CCF 50%)"),
])
je("신용장·보증", "BG-02", "이행보증 청구 INR 1 crore", "BG11", "② 청구 지급(대지급)", "2027-03-15", "R-BG-05", [
    ("D", "10360300000", "INR", 10000000.00, "INR_AMT", "보증 대지급금"), ("C", RBI, "INR", 10000000.00, "INR_AMT", "수익자 RTGS"),
    ("OB-", "60130400000", "INR", 10000000.00, "BG_AMT", "난외 감소 (PERFORMANCE)"),
])


def check():
    """단계별·통화별 대차 검증. 실패 시 AssertionError"""
    for biz, sid, title, scr, step, dt, ref, legs in J:
        for a in legs:
            assert a[1] in BY, ("CoA에 없는 계정", sid, a[1])
            assert usable(a[1], dt), ("기장 불가 계정(미사용·IFRS 전용)", sid, a[1])
            assert a[1] in LEAF, ("상위 계정 직접 기장 금지", sid, a[1])
        ccys = {l[2] for l in legs if l[0] in "DC"}
        for c in ccys:
            d = r2(sum(l[3] for l in legs if l[0] == "D" and l[2] == c))
            cr = r2(sum(l[3] for l in legs if l[0] == "C" and l[2] == c))
            assert abs(d - cr) < 0.005, (sid, step, c, d, cr)
    return len(J)


check()


def journal_rows():
    out = []
    for biz, sid, title, scr, step, dt, ref, legs in J:
        if not legs:
            out.append((biz, sid, title, scr, step, dt, "", "", "", "", None, None, "분개 없음: TB_HOL_FLOAT에 ABSORB 기록, 관리회계 보고만", ref))
        for i, (dc, code, ccy, amt, basis, memo) in enumerate(legs, 1):
            dr = amt if dc in ("D", "OB+") else None
            cr = amt if dc in ("C", "OB-") else None
            kind = {"D": "차변", "C": "대변", "OB+": "난외+", "OB-": "난외-"}[dc]
            out.append((biz, sid, title, scr, step, dt, kind, code, acct_name(code), ccy, dr, cr, memo, ref))
    return out


def rule_rows():
    """TB_ACCT_RULE 초기적재: (이벤트, 순번, 차대, 계정, 통화기준, 금액기준)"""
    out = []
    for biz, sid, title, scr, step, dt, ref, legs in J:
        ev = f"{sid}-{step[0]}"
        for i, (dc, code, ccy, amt, basis, memo) in enumerate(legs, 1):
            out.append((ev, i, dc, code, "FCY" if ccy != "INR" else "INR", basis, scr, f"{title} / {step}"))
    return out


def gl_summary():
    """설계서·엑셀 회계분개 요약 (이벤트, 차변, 대변, 비고)"""
    out = []
    for biz, sid, title, scr, step, dt, ref, legs in J:
        if not legs:
            out.append((f"[{biz}] {title} {step}", "-", "-", "분개 없음(ABSORB)"))
            continue
        f = lambda k: " / ".join(dict.fromkeys(f"{l[1]} {acct_name(l[1])}" for l in legs if l[0] in k))
        ob = f(("OB+", "OB-"))
        out.append((f"[{biz}] {title} {step}", f(("D",)) or "-", f(("C",)) or "-", ("난외: " + ob + " · " if ob else "") + scr))
    return out


if __name__ == "__main__":
    print(check(), "steps OK;", len(journal_rows()), "lines;", len(BASE), "accounts;", len(MAPPING), "mapping rows;",
          sum(1 for m in MAPPING if m["sev"] != "OK"), "to review")
