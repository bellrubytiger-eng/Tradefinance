# -*- coding: utf-8 -*-
"""차세대 GBS 통합 표준 CoA(0930 기준) + 인도(Noida) 지점 적용분 + 업무별 예시 분개.

- BASE  : 통합 표준 CoA 원본 (tools/data/coa_base.json, 업로드 원본에서 추출)
- ADD   : 인도지점 신규 계정 (주황 표기)        → 본점 채번 전 임시코드
- MOD   : 인도지점 적용 시 표시명·사용방식 변경 (주황 표기)
- DEACT : 인도지점 비활성 (회색, 삭제 아님)
- FIX   : 통합 CoA 자체 정정 권고 (노랑)
- JOURNAL : 업무별 예시 분개 (금액은 forms_model 샘플과 일치)
"""
import json
import os
from forms_model import r2, gst_fx_value, legs_fixed

_HERE = os.path.dirname(os.path.abspath(__file__))
BASE = json.load(open(os.path.join(_HERE, "data", "coa_base.json"), encoding="utf-8"))
BY = {a["code"]: a for a in BASE}

# ---------------------------------------------------------------------------
# 1. 신규 계정 (code, level, 영문명, 한글명, 본점 매핑 복사 기준 계정, 메모)
# ---------------------------------------------------------------------------
ADD = [
    # 예금: 규제 구분 L5 (결정 3)
    ("20150100100", 5, "CA - RESIDENT CORPORATE (INR)", "거주자 법인 당좌(INR)", "20150100000", "CRR·SLR 대상, TDS 해당 없음(무이자)"),
    ("20150100200", 5, "EEFC A/C - CORPORATE (FCY)", "EEFC 계정(법인, 외화)", "20150100000", "FEM(Deposit) Regs. 무이자, 외화 그대로 보유"),
    ("20150100300", 5, "SNRR A/C - NON-RESIDENT BUSINESS (INR)", "SNRR 계정(비거주 사업용)", "20150100000", "비거주자 INR 사업결제용. 7년 기한(사업용), CIMS 보고"),
    ("20150100400", 5, "SNRR A/C - FPI / PORTFOLIO INVESTMENT (INR)", "SNRR 계정(FPI·증권투자)", "20150100000", "한국 신탁·기관투자자(FPI) 증권투자 결제. 본점 Vostro와 분리"),
    ("20150200100", 5, "VOSTRO - SRVA (INR, RUPEE TRADE SETTLEMENT)", "특별루피 보스트로(SRVA)", "20150200000", "타 해외은행의 INR 무역결제 계좌(RBI 2022.7 체계). 잉여잔액 국채투자 허용"),
    ("20150200200", 5, "VOSTRO - INR (CORRESPONDENT, NON-SRVA)", "일반 INR 보스트로(타 해외은행)", "20150200000", "SRVA 외 일반 코레스 INR 계좌. 본점 계좌는 여기 아님(21040000000)"),
    ("20150200300", 5, "CA - DOMESTIC BANKS (INR)", "국내은행 당좌(INR)", "20150200000", "인도 국내은행 예치금(은행간)"),
    ("20150300100", 5, "SB/CA - RESIDENT INDIVIDUAL (INR)", "거주자 개인 보통·당좌(INR)", "20150300000", "LRS 송금 출금계좌. 이자 TDS(194A)"),
    ("20150300200", 5, "NRE SB A/C (INR)", "NRE 보통예금", "20150300000", "송금 가능(repatriable), 이자 비과세"),
    ("20150300300", 5, "NRO SB A/C (INR)", "NRO 보통예금", "20150300000", "송금 연 USD 1M 한도(RT14), 이자 TDS(195)"),
    ("20150300400", 5, "RFC / RFC(D) A/C (FCY)", "RFC 계정(귀국 거주자, 외화)", "20150300000", "귀국 거주자 외화 보유"),
    ("20150300500", 5, "EEFC A/C - INDIVIDUAL (FCY)", "EEFC 계정(개인 전문직 등)", "20150300000", "개인 서비스 수출대금 외화 보유"),
    ("20150400000", 4, "ESCROW DEPOSIT (FDI / OTHERS)", "에스크로 예금(FDI 등)", "20150100000", "FD13. FEM(NDI) 에스크로, 무이자, 인출조건 관리"),
    ("20160100100", 5, "TD - RESIDENT CORPORATE (INR)", "거주자 법인 정기예금", "20160100000", ""),
    ("20160100900", 5, "TD - OTHER CORPORATE", "기타 법인 정기예금", "20160100000", ""),
    ("20160300100", 5, "TD - RESIDENT INDIVIDUAL (INR)", "거주자 개인 정기예금", "20160300000", "이자 TDS(194A)"),
    ("20160300200", 5, "NRE TERM DEPOSIT (INR)", "NRE 정기예금", "20160300000", "만기 1년 이상"),
    ("20160300300", 5, "NRO TERM DEPOSIT (INR)", "NRO 정기예금", "20160300000", "이자 TDS(195)"),
    ("20160300400", 5, "FCNR(B) TERM DEPOSIT (FCY)", "FCNR(B) 외화정기예금", "20160300000", "1~5년, 금리상한 규제(ARR+스프레드)"),
    ("20160300500", 5, "RFC TERM DEPOSIT (FCY)", "RFC 외화정기예금", "20160300000", ""),
    # 본점 INR Vostro (결정 1)
    ("21040000000", 3, "INTEROFFICE INR VOSTRO - HEAD OFFICE", "본점 INR 보스트로(본지점)", "21010000000", "본점 명의 INR 계좌. 같은 법인이므로 예금이 아닌 본지점 부채. MT950 대상, 을A 정산배치·이자에서 제외. 마이너스 잔액 불허"),
    # 미지급 (결정 2: CoA대로 분리)
    ("20181100100", 5, "REMITTANCE PAYABLE - INWARD PENDING (NON-A/C, KYC)", "타발 미지급(비계좌·지시대기)", "20181100000", "RT21 SETTLE_MODE=SUSP → RT22. 통화별 하위원장(TB_AP_ITEM INW_RELAY)"),
    ("20181100200", 5, "REMITTANCE PAYABLE - FDI EXCESS TO BE RETURNED", "FDI 초과입금 반환대기(외화)", "20181100000", "R-FDI-07. 환전 금지, 5영업일 내 원통화 반환"),
    ("20181100300", 5, "REMITTANCE PAYABLE - RETURNED DOMESTIC PAYOUT", "NEFT/RTGS 반송 미지급", "20181100000", "pacs.004 수신분 복원"),
    ("20440800000", 4, "ACCOUNT PAYABLE (TRADE FINANCE DISBURSEMENT)", "무역금융 순지급 미지급금", "20440700000", "TF11 승인 순지급액. SF11 RTGS/NEFT 지급 시 차감 (R-TF-09)"),
    # 세금
    ("20510400100", 5, "TDS PAYABLE (194A / 195)", "TDS 예수금", "20510400000", "26Q/27Q 분기신고, 익월 7일 납부"),
    ("20510400200", 5, "TCS PAYABLE - LRS (206C(1G))", "TCS 예수금(LRS)", "20510400000", "27EQ 분기신고, 익월 7일 납부"),
    ("20510500100", 5, "OUTPUT CGST PAYABLE", "GST 매출세액 CGST", "20510500000", "GSTR-1/3B 세목별. 지점 등록주(UP 09)"),
    ("20510500200", 5, "OUTPUT SGST PAYABLE", "GST 매출세액 SGST", "20510500000", ""),
    ("20510500300", 5, "OUTPUT IGST PAYABLE", "GST 매출세액 IGST", "20510500000", "타 주 고객·수입서비스 RCM"),
    ("10420100100", 5, "INPUT TAX CREDIT - CGST", "GST 매입세액공제 CGST", "10420100000", "은행 ITC 50% 제한(s.17(4)) 반영"),
    ("10420100200", 5, "INPUT TAX CREDIT - SGST", "GST 매입세액공제 SGST", "10420100000", ""),
    ("10420100300", 5, "INPUT TAX CREDIT - IGST", "GST 매입세액공제 IGST", "10420100000", ""),
    # 무역금융
    ("10320300000", 4, "BILLS DISCOUNTED - DOMESTIC TRADE (INVOICE / TReDS, INR)", "국내 매출채권 할인(Invoice·TReDS)", "10320100000", "TF11 OA·TReDS. 본점 매핑 확인 필요(현재 은행인수어음과 동일 매핑)"),
    ("40350300000", 4, "INT. ON BILLS DISCOUNTED - DOMESTIC TRADE", "국내 매출채권 할인이자", "40350100000", ""),
    ("10490910000", 4, "INTEREST SUBVENTION RECEIVABLE (GOI)", "이자보전 미수금(정부)", "10490100000", "EPM Niryat Protsahan 등 정부 청구분"),
    ("40340500000", 4, "INTEREST SUBVENTION ON EXPORT CREDIT", "수출금융 이자보전 수익", "40340300000", ""),
    ("60120400000", 4, "OUTSTANDING INLAND L/C (SFMS IFN 700)", "내국신용장 미결제(난외)", "60120100000", "국내 LC 구분 보고"),
    ("60130400100", 5, "CUST LIABS ON GUARANTEES - FINANCIAL BG", "금융보증(난외)", "60130400000", "NFB Directions 2025: CCF 100%"),
    ("60130400200", 5, "CUST LIABS ON GUARANTEES - PERFORMANCE BG", "이행보증(난외)", "60130400000", "NFB Directions 2025: CCF 50%"),
    ("40510060000", 4, "COMMISSION - BANK GUARANTEE (FIN./PERF.)", "보증료(금융·이행보증)", "40510050000", "선수 후 기간안분(20420170000)"),
    # 수수료
    ("40570100000", 4, "COMMISSION - IN LIEU OF EXCHANGE", "환전 대체 수수료(무환전 송금)", "40570010000", "EEFC·외화계정 출입 송금"),
    ("40592030000", 4, "COMMISSION - VOSTRO A/C MAINTENANCE & STATEMENT", "보스트로 계좌유지·명세서 수수료", "40592020000", "타 해외은행 SRVA·일반 Vostro. 본점 계좌 미부과"),
]

# ---------------------------------------------------------------------------
# 2. 인도지점 적용 변경 (code, 인도 표시명 또는 None, 메모)
# ---------------------------------------------------------------------------
MOD = [
    ("10104010000", "DUE FROM RESERVE BANK OF INDIA (RTGS SETTLEMENT / CRR)", "통합 명칭 STATE BANK = 중앙은행. 인도지점 표시명 RBI. NEFT/RTGS·TCS/TDS 납부·이자보전 수령"),
    ("10101000000", "CASH ON HAND (INR)", "인도지점 자국통화 현금. 외국통화는 10103000000"),
    ("10105020000", None, "Nostro(USD·EUR·JPY·KRW). 통화는 계정이 아닌 통화코드 차원으로 관리"),
    ("11200000000", None, "통화별 환포지션. 분개는 이 계정 하나로, B/S 표시는 잔액 방향에 따라 11200000000/21300000000"),
    ("20150100000", None, "직접 기장 금지 → L5(거주자 당좌·EEFC·SNRR) 사용"),
    ("20150200000", None, "직접 기장 금지 → L5(SRVA·일반 Vostro·국내은행) 사용"),
    ("20150300000", None, "직접 기장 금지 → L5(거주자·NRE·NRO·RFC·EEFC) 사용"),
    ("20160100000", None, "직접 기장 금지 → L5 사용"),
    ("20160300000", None, "직접 기장 금지 → L5(NRE·NRO·FCNR(B)·RFC) 사용"),
    ("20181100000", None, "송금 미지급. 직접 기장 금지 → L5(타발 보류·FDI 초과·반송) 사용"),
    ("20440700000", None, "일반 미지급. 무역금융 순지급은 20440800000"),
    ("20510400000", None, "직접 기장 금지 → TDS/TCS L5"),
    ("20510500000", "OUTPUT GST PAYABLE (CGST/SGST/IGST)", "VAT → 인도 GST. 직접 기장 금지 → 세목 L5"),
    ("10420100000", "INPUT TAX CREDIT - GST", "VAT 선급 → GST ITC. 세목 L5 사용"),
    ("60130400000", None, "직접 기장 금지 → 금융/이행보증 L5"),
    ("11010000000", None, "을A Due from HO: RA 대지급·본지점 송금. 정산배치 대상"),
    ("21010000000", None, "을A Due to HO"),
    ("11020000000", None, "을B 본점 대여(예치). 기간·우대금리"),
    ("21020000000", None, "을B 본점 차입. MT320 확인"),
    ("31140100000", None, "갑기금(영업기금). 무이자·정산 제외"),
    ("40620000000", None, "을A 이자수익 + 휴일차이 이자 SETTLE분"),
    ("40630000000", None, "을B 이자수익"),
    ("51130000000", None, "을B 이자비용"),
]

# ---------------------------------------------------------------------------
# 3. 인도지점 비활성 (삭제 아님)
# ---------------------------------------------------------------------------
DEACT = [
    ("10490200000", "NAPAS(타국 결제망) - 인도지점 미사용"),
    ("20440500000", "NAPAS - 인도지점 미사용"),
    ("40570080000", "NAPAS - 인도지점 미사용"),
    ("50510960000", "NAPAS - 인도지점 미사용"),
    ("20180100000", "AOB - 인도지점 미사용"),
    ("20180200000", "AOB - 인도지점 미사용"),
    ("20440400000", "AOB - 인도지점 미사용"),
    ("40590010000", "AOB - 인도지점 미사용"),
    ("50510950000", "AOB - 인도지점 미사용"),
    ("40750000000", "TOKEN - 인도지점 미사용"),
    ("51210000000", "TOKEN - 인도지점 미사용"),
    ("40610000000", "갑기금 무이자 확정 → 기장 금지"),
    ("51110000000", "갑기금 무이자 확정 → 기장 금지"),
    ("20210100000", "STATE FUNDS - 콜머니(20220100000~)와 중복 여부 확인 후 비활성"),
    ("20210200000", "STATE FUNDS - 확인 후 비활성"),
    ("20210300000", "STATE FUNDS - 확인 후 비활성"),
]

# ---------------------------------------------------------------------------
# 4. 통합 CoA 정정 권고 (본점 CoA 담당 전달)
# ---------------------------------------------------------------------------
FIX = [
    ("40510020000", "GBS코드 42010002가 40510030000과 중복"),
    ("40510030000", "GBS코드 42010002가 40510020000과 중복"),
    ("40310000000", "GBS코드 41020000이 상위 40300000000과 중복"),
    ("20900000000", "GBS코드 54010000이 비용 50600000000과 중복(부채 충당부채)"),
    ("20720000000", "GBS코드 23100000이 상위 20700000000과 중복"),
    ("40340100000", "본점 매핑 불일치: 일람불 매입이자 '기타외환수입수수료' vs 네고 '수출관련외환수입이자'"),
    ("40340200000", "본점 매핑 불일치: 기한부(D/A) 매입이자도 '수출관련외환수입이자'로 통일 권고"),
    ("40310100000", "본점 매핑 불일치: 자산 '수입결제외화대출' vs 이자 '일반외화대출금이자(일시)'"),
    ("31120100000", "본점 매핑 명칭 불일치: 이익 '외화본지점을' vs 결손 '외화본지점(국외을)'"),
    ("40520080000", "L/G 수수료 중복(40510050000). 한 계정으로 통일"),
    ("40520070000", "인수 수수료 중복(40510010000). 한 계정으로 통일"),
    ("20120100000", "본점 매핑 '해외은행' - 고객 미결제 성격과 불일치 확인"),
    ("50240020000", "코드 패턴 오류: 50240200000이 맞음"),
    ("20500000000", "L2 명칭 중복(OTHER LIABILITIES TOTAL, 20400000000과 동일)"),
    ("51000000000", "L2 명칭 중복(OTHER OPERATING EXPENSE TOTAL, 50900000000과 동일)"),
    ("40600000000", "정렬 순서: 40700000000 뒤에 위치"),
    ("10107050100", "오탈자 DOMESITC → DOMESTIC"),
    ("10420090000", "오탈자 CERITIFICATE → CERTIFICATE"),
    ("40410000000", "오탈자 FIANACIAL → FINANCIAL"),
    ("40400000000", "오탈자 REVERSAL REVERSAL"),
    ("40510020000", "오탈자 COMMISIION"),
    ("40810000000", "오탈자 INCONE → INCOME"),
    ("40550030000", "오탈자 EXTENTION → EXTENSION"),
    ("20181100000", "오탈자 REMMITTANCE → REMITTANCE"),
    ("10410230000", "오탈자 PURCASE → PURCHASE"),
    ("51207000000", "오탈자 REMEMPTION → REDEMPTION"),
    ("10107070000", "GBS코드 누락(우리 업무 관련)"),
]


def coa_rows():
    """CoA 전체 + 인도지점 상태. (row dict 목록, 코드순)"""
    mod = {c: (n, m) for c, n, m in MOD}
    de = dict(DEACT)
    fx = {}
    for c, m in FIX:
        fx.setdefault(c, []).append(m)
    rows = []
    for a in BASE:
        r = dict(a, br_name="", status="", memo="")
        if a["code"] in mod:
            r["br_name"] = mod[a["code"]][0] or ""
            r["status"], r["memo"] = "변경", mod[a["code"]][1]
        if a["code"] in de:
            r["status"], r["memo"] = "비활성", de[a["code"]]
        if a["code"] in fx:
            r["fix"] = " / ".join(fx[a["code"]])
        rows.append(r)
    for code, lv, en, kr, sib, memo in ADD:
        s = BY[sib]
        rows.append({"grp": s["grp"], "area": s["area"], "code": code, "lv": lv, "name": en, "krw": kr,
                     "ho": s["ho"], "ho_name": s["ho_name"] + " (안)", "nonmon": s["nonmon"], "gbs": "채번요청",
                     "br_name": "", "status": "신규", "memo": memo})
    rows.sort(key=lambda r: r["code"])
    return rows


ALL = {r["code"]: r for r in coa_rows()}


def acct_name(code):
    r = ALL[code]
    return r.get("br_name") or r["name"]


# ---------------------------------------------------------------------------
# 5. 업무별 예시 분개
#    leg = (차/대, 계정, 통화, 금액, 금액기준, 적요)
#    D/C = 차변/대변, OB+/OB- = 난외 증가/감소(대차검증 제외)
#    환전은 통화별로 대차를 맞추고 11200000000 포지션으로 연결
# ---------------------------------------------------------------------------
POS, NOSTRO, RBI = "11200000000", "10105020000", "10104010000"
CGST, SGST, IGST = "20510500100", "20510500200", "20510500300"


def _gst(base, pos):
    g = r2(base * 0.18)
    if pos == "CS":
        c = r2(g / 2)
        return [("C", CGST, "INR", c, "CGST", "CGST 9%"), ("C", SGST, "INR", r2(g - c), "SGST", "SGST 9%")]
    return [("C", IGST, "INR", g, "IGST", "IGST 18%")]


def _gsum(legs):
    return r2(sum(l[3] for l in legs))


J = []  # (업무, 시나리오ID, 시나리오, 화면, 단계, 일자, 근거, legs)


def je(biz, sid, title, scr, step, dt, ref, legs):
    J.append((biz, sid, title, scr, step, dt, ref, legs))


# ---- 당발송금 -------------------------------------------------------------
fee = [("C", "40570010000", "INR", 11052.50, "COMM", "송금수수료"), ("C", "40570030000", "INR", 1000.00, "CABLE", "전신료")]
g = _gst(11052.50 + 1000 + gst_fx_value(8842000.00), "CS")
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
g = _gst(1500 + gst_fx_value(3978900.00), "CS")
tcs = ("C", "20510400200", "INR", 198945.00, "TCS", "TCS 5% (교육, 자기자금)")
tot = r2(3978900.00 + _gsum(fee) + _gsum(g) + 198945.00)
je("당발송금", "OUT-02", "LRS 유학비 USD 45,000 (개인 SB, TCS)", "RT12", "① 출금·환전·TCS", "2026-10-31", "R-LRS, 206C(1G)", [
    ("D", "20150300100", "INR", tot, "TOTAL_DR", "개인 SB 출금"), ("C", POS, "INR", 3978900.00, "INR_AMT", "환전 INR"),
    *fee, *g, tcs,
    ("D", POS, "USD", 45000.00, "FCY_AMT", ""), ("C", NOSTRO, "USD", 45000.00, "FCY_AMT", "Nostro 출금"),
])
je("당발송금", "OUT-02", "LRS 유학비 USD 45,000 (개인 SB, TCS)", "RT12", "② TCS 납부(익월 7일)", "2026-11-07", "27EQ", [
    ("D", "20510400200", "INR", 198945.00, "TCS", "TCS 예수금 정리"), ("C", RBI, "INR", 198945.00, "TCS", "정부 납부(RBI 경유)"),
])
je("당발송금", "OUT-03", "EEFC 출금 송금 USD 20,000 (무환전)", "RT11", "① EEFC 출금·송신 + 환전대체수수료", "2026-10-31", "R-FX, 40570100000", [
    ("D", "20150100200", "USD", 20000.00, "FCY_AMT", "EEFC 출금"), ("C", NOSTRO, "USD", 20000.00, "FCY_AMT", "Nostro 출금"),
    ("D", "20150100100", "INR", 2360.00, "TOTAL_DR", "수수료 INR 계좌 출금"),
    ("C", "40570100000", "INR", 1000.00, "COMM", "In lieu of exchange (요율표 예시)"),
    ("C", "40570030000", "INR", 1000.00, "CABLE", "전신료"), *_gst(2000.00, "CS"),
])

# ---- 타발송금 -------------------------------------------------------------
g = _gst(250 + gst_fx_value(7639488.00), "IG")
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
g = _gst(525 + gst_fx_value(4599400.00), "CS")
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
g = _gst(7500 + gst_fx_value(conv_inr), "IG")
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
    ("C", "40570030000", "INR", 1000.00, "CABLE", ""), *_gst(6000 + gst_fx_value(50000000), "IG"),
    ("D", POS, "USD", 565482.92, "FCY_AMT", ""), ("C", NOSTRO, "USD", 565482.92, "FCY_AMT", "원송금인 반환"),
])
esc_fcy, esc_inr = legs_fixed(100000000.00, 88.40)
je("FDI", "FDI-03", "FDI 에스크로 INR 10 crore (지분 인수)", "FD13", "① 수취·환전·에스크로 예치", "2026-10-15", "FEM(NDI) 에스크로", [
    ("D", NOSTRO, "USD", esc_fcy, "FCY_AMT", ""), ("C", POS, "USD", esc_fcy, "FCY_AMT", ""),
    ("D", POS, "INR", esc_inr, "INR_AMT", ""), ("C", "20150400000", "INR", esc_inr, "INR_AMT", "에스크로 예치(무이자)"),
])
je("FDI", "FDI-03", "FDI 에스크로 INR 10 crore (지분 인수)", "FD13", "② 조건 충족 → 매도인 RTGS 지급", "2026-11-20", "에스크로 약정", [
    ("D", "20150400000", "INR", esc_inr, "INR_AMT", "에스크로 해지"), ("C", RBI, "INR", esc_inr, "INR_AMT", "매도인 RTGS"),
])

# ---- 환전 -----------------------------------------------------------------
g = _gst(gst_fx_value(4419000.00), "CS")
je("환전", "FX-01", "EEFC USD 50,000 → INR CA (88.38)", "FX11", "① 계좌 간 환전", "2026-10-31", "R-FX", [
    ("D", "20150100200", "USD", 50000.00, "FCY_AMT", "EEFC 출금"), ("C", POS, "USD", 50000.00, "FCY_AMT", ""),
    ("D", POS, "INR", 4419000.00, "INR_AMT", ""), ("C", "20150100100", "INR", 4419000.00, "INR_AMT", "CA 입금"),
    ("D", "20150100100", "INR", _gsum(g), "GST", "환전 GST 출금"), *g,
])
g = _gst(gst_fx_value(104340.00), "CS")
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
    ("OB+", "60120400000", "INR", 20000000.00, "LC_AMT", "내국 LC 난외 증가"),
])
bg_c = 750000.00
je("신용장·보증", "BG-01", "금융보증 INR 5 crore · 1년 · 1.5%", "BG11", "① 발행·보증료 선수", "2026-10-31", "NFB Directions, FEM(Guarantees)", [
    ("OB+", "60130400100", "INR", 50000000.00, "BG_AMT", "금융보증 난외(CCF 100%)"),
    ("D", "20150100100", "INR", r2(bg_c * 1.18), "TOTAL_DR", ""), ("C", "20420170000", "INR", bg_c, "COMM", "보증료 선수(기간안분)"), *_gst(bg_c, "CS"),
])
je("신용장·보증", "BG-01", "금융보증 INR 5 crore", "EOD", "② 월 보증료 안분", "2026-11-30", "", [
    ("D", "20420170000", "INR", 62500.00, "COMM", ""), ("C", "40510060000", "INR", 62500.00, "COMM", "보증료 수익"),
])
je("신용장·보증", "BG-02", "이행보증 INR 1 crore 발행", "BG11", "① 발행", "2026-10-31", "", [
    ("OB+", "60130400200", "INR", 10000000.00, "BG_AMT", "이행보증 난외(CCF 50%)"),
])
je("신용장·보증", "BG-02", "이행보증 청구 INR 1 crore", "BG11", "② 청구 지급(대지급)", "2027-03-15", "R-BG-05", [
    ("D", "10360300000", "INR", 10000000.00, "INR_AMT", "보증 대지급금"), ("C", RBI, "INR", 10000000.00, "INR_AMT", "수익자 RTGS"),
    ("OB-", "60130400200", "INR", 10000000.00, "BG_AMT", "난외 감소"),
])


def _parents():
    rows = coa_rows()
    out = set()
    for r in rows:
        if r["lv"] >= 3:
            pre = r["code"][: 1 + 2 * (r["lv"] - 1)]
            if any(x["lv"] > r["lv"] and x["code"].startswith(pre) for x in rows):
                out.add(r["code"])
    return out


PARENTS = _parents()


def check():
    """단계별·통화별 대차 검증. 실패 시 AssertionError"""
    for biz, sid, title, scr, step, dt, ref, legs in J:
        for a in legs:
            assert a[1] in ALL, (sid, a[1])
            assert ALL[a[1]]["status"] != "비활성", (sid, a[1])
            assert a[1] not in PARENTS, ("직접 기장 금지 계정", sid, a[1])
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
    print(check(), "steps OK;", len(journal_rows()), "lines;", len([r for r in coa_rows() if r["status"]]), "marked")
