# -*- coding: utf-8 -*-
"""
고객 서식 3종 정의 (HTML 서식, Word 서식명세, Excel 필드매핑의 단일 원천)
필드: (key, 영문 라벨, 한글 항목, DB 소스, 표시조건)  표시조건 None = 모든 변형
샘플 데이터는 가상 고객이며, 화면에는 SPECIMEN 표기와 함께 렌더링된다.
"""

MST = "TB_FX_REMIT_MST"
FDI = "TB_FDI_INFLOW"
ISS = "TB_FX_ADVICE_ISSUE"
PTY = "TB_FX_REMIT_PARTY"
CUS = "TB_CUST_FX_PROFILE"

FORMS = [
    {
        "id": "FM-FDI-01", "key": "firc",
        "title": "Foreign Inward Remittance Certificate",
        "subtitle": "cum Encashment Certificate · Foreign Direct Investment",
        "ko": "FDI 외화수취·환전증명서 (FIRC)",
        "screen": "FD11 (재발급 DC01)",
        "when": "FD11 등록 + 입금완료 + 투자자 KYC 수취 후 (R-DOC-01)",
        "copies": "원본 1부 고객 교부, PDF·발급내용 스냅샷 보관",
        "use": "피투자법인이 RBI FIRMS에 Form FC-GPR 신고 시 첨부",
        "variants": [("fdi", "당행 계좌 입금"), ("fdi_rtgs", "비계좌 · 부분환전 후 RTGS 지급")],
        "sections": [
            {"t": "Certificate", "f": [
                ("cert_no", "Certificate No.", "증명서 번호", f"{ISS}.ISSUE_NO", None),
                ("issue_dt", "Date of Issue", "발급일", f"{ISS}.ISSUE_DT", None),
                ("fdi_ref", "FDI Reference", "FDI 관리번호", f"{FDI}.FDI_REF_NO", None),
                ("our_ref", "Bank Reference", "당행 송금참조", f"{MST}.REMIT_REF_NO", None),
            ]},
            {"t": "Beneficiary · Indian Investee Company", "f": [
                ("ben_name", "Name", "피투자법인명", f"{CUS}.(INVESTEE_CUST_NO) 고객명", None),
                ("ben_cin", "CIN", "법인등록번호", f"{FDI}.INVESTEE_CIN", None),
                ("ben_pan", "PAN", "PAN", f"{MST}.PAN_SNAP", None),
                ("ben_acct", "Account Credited", "입금계좌", f"{MST}.SRC_ACCT_NO", None),
                ("ben_addr", "Registered Office", "등록주소", f"{PTY}(BENE).ADDR1/ADDR2/TOWN", None),
            ]},
            {"t": "Remitter · Non-resident Investor", "f": [
                ("rem_name", "Name", "투자자명", f"{FDI}.INVESTOR_NAME", None),
                ("rem_addr", "Address", "주소", f"{PTY}(APPL).ADDR1/TOWN", None),
                ("rem_ctry", "Country", "국가", f"{FDI}.INVESTOR_CTRY", None),
                ("rem_bank", "Remitting Bank", "송금은행 / BIC", f"{PTY}(ORDINS).PARTY_NAME, BIC", None),
                ("their_ref", "Sender's Reference", "송금은행 Ref", f"{MST}.THEIR_REF_NO", None),
                ("uetr", "UETR", "UETR", f"{MST}.UETR", None),
            ]},
            {"t": "Remittance and Conversion", "f": [
                ("recv_dt", "Date of Receipt", "수취일", f"{FDI}.RECEIPT_DT", None),
                ("instd", "Amount Remitted", "송금 지시금액", f"{MST}.CCY + INSTD_AMT", None),
                ("fchg", "Charges Deducted Abroad", "해외 공제수수료", f"{MST}.FOREIGN_CHG_AMT", None),
                ("fcy", "Net Foreign Currency Received", "수취 외화금액", f"{MST}.CCY + FCY_AMT", None),
                ("conv", "Converted (agreed investment)", "약정 환전 외화", f"{FDI}.CONV_FCY_AMT", ["fdi"]),
                ("excess", "Excess Returned, Not Converted", "초과 반환(미환전)", f"{FDI}.EXCESS_FCY_AMT, EXCESS_REFUND_REF", ["fdi"]),
                ("rate", "Exchange Rate", "적용환율", f"{MST}.FX_RATE", ["fdi"]),
                ("conv_dt", "Date of Conversion", "환전일", f"{MST}.VALUE_DT", ["fdi"]),
                ("deal_ref", "Dealing Reference", "딜 Ref", "TB_FX_DEAL_LINK.DEAL_REF_NO", ["fdi"]),
                ("inr", "INR Amount Credited", "입금 INR", f"{MST}.INR_AMT", None),
                ("inr_words", "Amount in Words", "금액 한글/영문", "INR_AMT → 인도식 영문(Crore/Lakh)", None),
                ("inr_fixed", "INR Fixed Conversion", "INR 확정 환전", f"{FDI}.INR_FIXED_YN, TARGET_INR_AMT", ["fdi_rtgs"]),
                ("conv_legs", "Conversion Details", "환전 레그", "TB_FX_CONV_LEG (CONV_DT, FCY_AMT, FX_RATE, INR_AMT, DEAL_REF_NO)", ["fdi_rtgs"]),
                ("payout", "Paid to Investee by", "타행 지급", "TB_DOM_PAYOUT (RAIL, UTR, BENE_IFSC, BENE_ACCT)", ["fdi_rtgs"]),
                ("charges", "Bank Charges", "수수료 (별도징수)", "TB_FX_REMIT_FEE, FEE_COLLECT_MODE=SEPARATE", None),
            ]},
            {"t": "Purpose and Investment", "f": [
                ("purpose", "Purpose Code", "Purpose Code", f"{MST}.PURPOSE_CODE + TB_CD_PURPOSE.DESC_EN", None),
                ("instr", "Instrument", "투자수단", f"{FDI}.INSTRUMENT", None),
                ("route", "Route", "투자경로 / 승인번호", f"{FDI}.ROUTE, GOVT_APPROVAL_REF", None),
                ("kyc", "KYC of Remitter (Investor)", "송금인 KYC", "TB_FDI_KYC.KYC_STS, REQ_SENT_DT, RESP_RCV_DT", None),
                ("allot_due", "Allotment Due By", "배정기한", f"{FDI}.ALLOT_DUE_DT", None),
                ("refund_due", "Refund Due By (if not allotted)", "반환기한", f"{FDI}.REFUND_DUE_DT", None),
            ]},
        ],
        "statements": [
            ("This is to certify that the foreign currency amount stated above has been received in India through normal banking channels from the remitter named above for credit to the beneficiary, and has been converted into Indian Rupees at the exchange rate stated.", None),
            ("The remittance has been received towards subscription to capital instruments of the beneficiary company under Foreign Direct Investment and is recorded under the purpose code stated above.", None),
            ("The KYC report on the non-resident investor has been received from the remitting bank.", None),
            ("Capital instruments are to be issued within 60 days of receipt of the amount. Any amount not so issued is to be refunded within 15 days thereafter, under the FEMA (Non-Debt Instruments) Rules, 2019.", None),
            ("The investee company does not hold an account with us. The amount was held in remittance payable, converted on the dates shown and paid in full by RTGS to the investee company's own account stated above.", ["fdi_rtgs"]),
            ("Only the agreed investment amount was converted into Indian Rupees. The excess amount shown was returned to the remitter in the original currency without conversion and does not form part of this investment.", ["fdi"]),
            ("Bank charges and GST were collected separately; the INR amount credited towards the investment has not been reduced by any charge.", None),
            ("This certificate is issued for submission with Form FC-GPR on the RBI FIRMS portal. A copy issued after the original is marked DUPLICATE.", None),
        ],
        "sign": ["Authorised Signatory", "Checked by", "Branch Seal"],
    },
    {
        "id": "FM-OUT-01", "key": "out",
        "title": "Outward Remittance Advice",
        "subtitle": "Debit Advice · Customer Copy",
        "ko": "당발송금 확인서",
        "screen": "RT11 / RT12 / RT13 / FD12 (재발급 DC01)",
        "when": "전문 송신완료(STATUS ≥ 30) 후 자동생성 (R-DOC-05)",
        "copies": "인터넷뱅킹 PDF, 창구 요청 시 출력",
        "use": "송금 사실 증빙, LRS TCS 확인, 수입 선수금 증빙기한 안내",
        "variants": [("imp", "기업 · 수입선수금 (RT13)"), ("lrs", "개인 · LRS (RT12)"), ("nro", "비거주자 · NRO 송금 (RT14)"), ("fdr", "기업 · FDI 반환송금 (FD12)")],
        "sections": [
            {"t": "Advice", "f": [
                ("adv_no", "Advice No.", "확인서 번호", f"{ISS}.ISSUE_NO", None),
                ("adv_dt", "Date", "발급일", f"{ISS}.ISSUE_DT", None),
                ("our_ref", "Transaction Reference", "송금참조번호", f"{MST}.REMIT_REF_NO", None),
                ("txn_type", "Transaction Type", "거래유형", f"{MST}.SCREEN_ID + PRODUCT_TYPE", None),
                ("value_dt", "Value Date", "결제일", f"{MST}.VALUE_DT", None),
                ("uetr", "UETR (SWIFT gpi)", "UETR", f"{MST}.UETR", None),
            ]},
            {"t": "Remitter", "f": [
                ("rem_name", "Name", "송금인", f"{PTY}(APPL).PARTY_NAME", None),
                ("cust_no", "Customer No.", "고객번호", f"{MST}.CUST_NO", None),
                ("rem_acct", "Account Debited", "출금계좌", f"{MST}.SRC_ACCT_NO (끝4자리)", None),
                ("rem_pan", "PAN", "PAN", f"{MST}.PAN_SNAP (마스킹)", None),
                ("rem_iec", "IEC", "IEC", f"{MST}.IEC_SNAP", ["imp"]),
                ("rem_lei", "LEI", "LEI", f"{MST}.LEI_SNAP", ["imp", "fdr"]),
            ]},
            {"t": "Beneficiary", "f": [
                ("ben_name", "Name", "수취인", f"{PTY}(BENE).PARTY_NAME", None),
                ("ben_acct", "Account / IBAN", "수취계좌", f"{PTY}(BENE).ACCT_NO", None),
                ("ben_addr", "Address", "주소", f"{PTY}(BENE).ADDR1/TOWN/CTRY", None),
                ("ben_bank", "Beneficiary Bank", "수취은행 / BIC", f"{PTY}(BENBNK).PARTY_NAME, BIC", None),
                ("intm", "Intermediary Bank", "중계은행", f"{PTY}(INTRMD).BIC", None),
            ]},
            {"t": "Amount", "f": [
                ("amt", "Remittance Amount", "송금액", f"{MST}.CCY + FCY_AMT", None),
                ("rate", "Exchange Rate", "적용환율", f"{MST}.FX_RATE (RATE_TYPE)", None),
                ("inr", "INR Equivalent", "INR 환산", f"{MST}.INR_AMT", None),
                ("chg_type", "Charges", "수수료 부담", f"{MST}.CHARGE_TYPE", None),
                ("fees", "Bank Charges and Taxes", "수수료·세금 내역", "TB_FX_REMIT_FEE (FEE_TYPE별)", None),
                ("gst_info", "GST Details", "GST 정보", "지점 GSTIN, 고객 GSTIN, SAC, GST_POS_STATE, GST_RCPT_TYPE", None),
                ("total", "Total Amount Debited", "총 출금액", "INR_AMT + Σ TB_FX_REMIT_FEE.AMT", None),
                ("total_words", "Amount in Words", "금액(영문)", "총 출금액 → 인도식 영문", None),
            ]},
            {"t": "Purpose and Regulatory", "f": [
                ("purpose", "Purpose Code", "Purpose Code", f"{MST}.PURPOSE_CODE + DESC", None),
                ("rmt_info", "Remittance Information", "송금 메시지", f"{MST}.TAG70_INFO", None),
                ("form_a2", "Form A2", "Form A2", f"{MST}.FORM_A2_YN", None),
                ("lrs_fy", "LRS Remitted in FY (incl. this)", "LRS FY 누적", "TB_LRS_TXN.FY, CUM_USD_AFTER", ["lrs"]),
                ("tcs", "TCS Collected (s.206C(1G))", "TCS 징수", "TB_LRS_TXN.TCS_RATE, TCS_AMT", ["lrs"]),
                ("inv_ref", "Invoice", "인보이스", "TB_TRD_PAY_ALLOC → TB_TRD_INVOICE.INV_NO, INV_DT", ["imp"]),
                ("orm", "IDPMS ORM No.", "ORM 번호", "TB_IF_IDPMS_ORM.ORM_NO", ["imp"]),
                ("evid_due", "Bill of Entry Due By", "수입증빙 제출기한", "TB_TRD_ADV_IMP.EVIDENCE_DUE_DT", ["imp"]),
                ("nro_fy", "NRO Repatriation in FY (incl. this)", "NRO FY 누적", "TB_NR_REPAT_SUM.CUM_USD_CAPITAL", ["nro"]),
                ("f15ca", "Form 15CA / 15CB", "15CA / 15CB", f"{MST}.F15CA_ACK_NO, F15CB_UDIN", ["imp", "fdr", "nro"]),
                ("fdi_ref", "Original FDI Reference", "원 FDI 관리번호", "TB_FDI_REFUND.FDI_REF_NO", ["fdr"]),
                ("refund_rsn", "Refund Reason", "반환사유", "TB_FDI_REFUND.REFUND_TYPE", ["fdr"]),
            ]},
        ],
        "statements": [
            ("We have remitted the above amount as per your instructions and debited your account with the total amount shown, including bank charges and applicable GST.", None),
            ("Charges of correspondent or beneficiary banks may be deducted from the amount paid to the beneficiary when charges are SHA or BEN. The payment can be tracked with the UETR.", None),
            ("This remittance is made under the Liberalised Remittance Scheme. Your total LRS remittances through all banks must not exceed USD 250,000 in a financial year (April to March). TCS collected will appear in Form 26AS; a TCS certificate (Form 27D) will be issued for the quarter.", ["lrs"]),
            ("This is an advance remittance for import. Please submit the Bill of Entry (evidence of import) to the bank by the due date shown. Outstanding advance remittances are reported to RBI through IDPMS.", ["imp"]),
            ("This remittance is made from your NRO account under the facility of USD 1 million per financial year for balances and sale proceeds of assets, on the basis of Form 15CA and the chartered accountant's certificate in Form 15CB.", ["nro"]),
            ("This is a refund of share application money not allotted within 60 days of receipt, made to the original remitter in the original currency under the FEMA (Non-Debt Instruments) Rules, 2019.", ["fdr"]),
        ],
        "sign": ["Authorised Signatory", "Branch Seal"],
    },
    {
        "id": "FM-IN-01", "key": "in",
        "title": "Inward Remittance Advice",
        "subtitle": "Credit Advice · Customer Copy",
        "ko": "타발송금 확인서",
        "screen": "RT21 (재발급 DC01)",
        "when": "입금완료 후 자동생성, 창구·인터넷뱅킹 발급 (R-DOC-06)",
        "copies": "인터넷뱅킹 PDF, 창구 요청 시 출력",
        "use": "입금 사실 증빙. 수출대금은 IRM 번호 안내. FDI용 FIRC 아님",
        "variants": [("exp", "기업 · 수출선수금 (RT21 → EDPMS)"), ("ind", "개인 · 송금 수취 (RT21)"), ("nre", "비거주자 · NRE 입금 (RT21)"), ("relay", "비계좌 · 미지급금 → RTGS 지급 (RT22)"), ("fdi", "기업 · FDI 입금 일반확인서 (RT21)")],
        "sections": [
            {"t": "Advice", "f": [
                ("adv_no", "Advice No.", "확인서 번호", f"{ISS}.ISSUE_NO", None),
                ("adv_dt", "Date", "발급일", f"{ISS}.ISSUE_DT", None),
                ("our_ref", "Our Reference", "당행 참조번호", f"{MST}.REMIT_REF_NO", None),
                ("their_ref", "Sender's Reference", "송금은행 Ref", f"{MST}.THEIR_REF_NO", None),
                ("value_dt", "Value Date", "결제일", f"{MST}.VALUE_DT", None),
                ("uetr", "UETR (SWIFT gpi)", "UETR", f"{MST}.UETR", None),
            ]},
            {"t": "Beneficiary · Our Customer", "f": [
                ("ben_name", "Name", "수취인", f"{PTY}(BENE).PARTY_NAME", None),
                ("cust_no", "Customer No.", "고객번호", f"{MST}.CUST_NO", None),
                ("ben_acct", "Account Credited", "입금계좌", f"{MST}.SRC_ACCT_NO (끝4자리), SRC_ACCT_TYPE", None),
                ("ben_pan", "PAN", "PAN", f"{MST}.PAN_SNAP (마스킹)", None),
                ("ben_iec", "IEC", "IEC", f"{MST}.IEC_SNAP", ["exp"]),
            ]},
            {"t": "Remitter", "f": [
                ("rem_name", "Name", "송금인", f"{PTY}(APPL).PARTY_NAME", None),
                ("rem_addr", "Address", "주소", f"{PTY}(APPL).ADDR1/TOWN", None),
                ("rem_ctry", "Country", "국가", f"{MST}.REMIT_CTRY", None),
                ("rem_bank", "Remitting Bank", "송금은행 / BIC", f"{PTY}(ORDINS).PARTY_NAME, BIC", None),
            ]},
            {"t": "Amount", "f": [
                ("instd", "Amount Remitted", "송금 지시금액", f"{MST}.CCY + INSTD_AMT", None),
                ("fchg", "Charges Deducted Abroad", "해외 공제수수료", f"{MST}.FOREIGN_CHG_AMT", None),
                ("fcy", "Amount Received", "수취 외화금액", f"{MST}.CCY + FCY_AMT", None),
                ("rate", "Exchange Rate", "적용환율", f"{MST}.FX_RATE", None),
                ("conv_legs", "Conversion Details", "환전 레그", "TB_FX_CONV_LEG", ["relay"]),
                ("fees", "Bank Charges and Taxes", "당행 수수료·세금", "TB_FX_REMIT_FEE", None),
                ("gst_info", "GST Details", "GST 정보", "지점 GSTIN, 고객 GSTIN, SAC, GST_POS_STATE, GST_RCPT_TYPE", None),
                ("net", "Net Amount Credited", "입금액", f"{MST}.INR_AMT - Σ FEE", None),
                ("net_words", "Amount in Words", "금액(영문)", "입금액 → 인도식 영문", None),
            ]},
            {"t": "Purpose and Regulatory", "f": [
                ("purpose", "Purpose Code", "Purpose Code", f"{MST}.PURPOSE_CODE + DESC", None),
                ("rmt_info", "Remittance Information", "송금 메시지", f"{MST}.TAG70_INFO", None),
                ("irm", "EDPMS IRM No.", "IRM 번호", "TB_TRD_IRM.IRM_NO", ["exp"]),
                ("ship_due", "Shipment Due By (advance)", "선적이행기한", "TB_TRD_IRM.SHIP_DUE_DT", ["exp"]),
                ("fdi_ref", "FDI Reference", "FDI 관리번호", f"{FDI}.FDI_REF_NO", ["fdi"]),
                ("excess", "Excess Returned, Not Converted", "초과 반환(미환전)", f"{FDI}.EXCESS_FCY_AMT", ["fdi"]),
                ("nr_acct", "Account Type", "비거주자 계좌", f"{MST}.NR_ACCT_TYPE", ["nre"]),
                ("payout", "Paid to You by", "타행 지급", "TB_DOM_PAYOUT (RAIL, UTR, BENE_IFSC, BENE_ACCT)", ["relay"]),
                ("susp", "Held in Remittance Payable", "미지급금 보류", "TB_FX_SUSPENSE (SUSP_NO, PARK_DT, SUSP_RSN)", ["relay"]),
            ]},
        ],
        "statements": [
            ("We have received the above remittance and credited your account with the net amount shown.", None),
            ("This remittance is recorded in EDPMS under the IRM number shown. Goods against this advance must be shipped by the due date; otherwise the advance is to be refunded or an extension obtained through the bank. The e-BRC is generated on the DGFT portal after the IRM is mapped to the shipping bill.", ["exp"]),
            ("This advice may be used as proof of an inward remittance. A Foreign Inward Remittance Certificate can be issued on request.", ["ind"]),
            ("The amount has been credited to your NRE account. NRE balances and interest are freely repatriable.", ["nre"]),
            ("You do not hold an account with us. The amount was held in remittance payable pending your instructions, converted on the dates shown and paid by RTGS to the account stated. Please confirm the credit with your bank using the UTR.", ["relay"]),
            ("This advice is not a Foreign Inward Remittance Certificate (FIRC) and cannot be used for Form FC-GPR. The FDI certificate (form FM-FDI-01) will be issued separately once the investor's KYC report is received.", ["fdi"]),
        ],
        "sign": ["Authorised Signatory", "Branch Seal"],
    },
    {
        "id": "FM-FX-01", "key": "fx",
        "title": "Foreign Exchange Advice",
        "subtitle": "Currency Exchange · Encashment Certificate for notes purchased",
        "ko": "환전 확인서 (Encashment Certificate 겸용)",
        "screen": "FX11 (재발급 DC01)",
        "when": "환전 승인 즉시. 현찰·TT 매입은 고객 요청 시 Encashment Certificate 문구 포함",
        "copies": "고객 1부, PDF·스냅샷 보관",
        "use": "환전 증빙, 외국인 출국 시 잔여 INR 재환전 근거, GST 세금계산서 정보",
        "variants": [("acct", "계좌 · EEFC USD → INR"), ("cross", "외화간 · USD → KRW (EEFC)"), ("cash", "현찰 매입 · 외국인 방문객")],
        "sections": [
            {"t": "Advice", "f": [
                ("adv_no", "Advice No.", "확인서 번호", f"{ISS}.ISSUE_NO", None),
                ("adv_dt", "Date", "발급일", f"{ISS}.ISSUE_DT", None),
                ("fx_ref", "Exchange Reference", "환전 참조번호", "TB_FX_EXCHANGE.FX_REF_NO", None),
                ("fx_type", "Type", "환전유형", "TB_FX_EXCHANGE.FX_TYPE", None),
            ]},
            {"t": "Customer", "f": [
                ("cust_name", "Name", "고객명", f"{CUS}.고객명 / 비계좌 성명", None),
                ("cust_no", "Customer No.", "고객번호", "TB_FX_EXCHANGE.CUST_NO", ["acct", "cross"]),
                ("passport", "Passport / Nationality", "여권 / 국적", "TB_FX_EXCHANGE.PASSPORT_NO", ["cash"]),
                ("cust_pan", "PAN", "PAN", "TB_FX_EXCHANGE.PAN_SNAP (마스킹)", ["acct", "cross"]),
                ("sell_acct", "Account Debited", "출금계좌", "TB_FX_EXCHANGE.SELL_ACCT_NO (끝4자리)", ["acct", "cross"]),
                ("buy_acct", "Account Credited", "입금계좌", "TB_FX_EXCHANGE.BUY_ACCT_NO (끝4자리)", ["acct", "cross"]),
                ("paid_by", "Paid by", "지급방법", "현금(한도 내) / 계좌", ["cash"]),
            ]},
            {"t": "Exchange", "f": [
                ("sell", "You Gave", "고객 지급", "TB_FX_EXCHANGE.SELL_CCY, SELL_AMT", None),
                ("rate", "Exchange Rate", "적용환율", "TB_FX_EXCHANGE.RATE, RATE_SRC", None),
                ("rbi_ref", "RBI Reference Rates (for GST)", "RBI 기준율", "TB_FX_EXCHANGE.RBI_REF_RATE", ["cross"]),
                ("deal_ref", "Dealing Reference", "딜 Ref", "TB_FX_EXCHANGE.DEAL_REF_NO", ["acct", "cross"]),
                ("notes", "Notes Purchased", "매입 권종", "TB_FX_CASH_STOCK (권종·수량)", ["cash"]),
                ("cdf", "Currency Declaration Form", "CDF", "TB_FX_EXCHANGE.CDF_NO", ["cash"]),
                ("fees", "Bank Charges and Taxes", "수수료·세금", "TB_FX_REMIT_FEE (FX_REF_NO)", None),
                ("gst_info", "GST Details", "GST 정보", "지점 GSTIN, 고객 GSTIN, SAC, 공급장소", None),
            ]},
        ],
        "statements": [
            ("We have exchanged the currency shown at the rate stated. Charges and GST are shown separately; GST on currency exchange is charged on the value determined under rule 32(2) of the CGST Rules.", None),
            ("The INR proceeds have been credited to your account. Balances in an EEFC account can be converted to INR at any time; INR cannot be converted back into the EEFC account.", ["acct"]),
            ("This is a cross-currency exchange within your EEFC account. No Indian Rupees were involved; the GST value is based on the RBI reference rates shown.", ["cross"]),
            ("Encashment Certificate: certified that we have purchased the foreign currency notes shown from the person named above. Please keep this certificate; it is required to reconvert unspent Indian Rupees into foreign currency on departure.", ["cash"]),
        ],
        "sign": ["Authorised Signatory", "Branch Seal"],
    },
]

BRANCH = {
    "bank": "[BANK NAME]", "branch": "India Noida Branch",
    "addr": "[Branch address], Sector 18, Noida, Uttar Pradesh 201301, India",
    "ad_code": "[AD Code]", "bic": "[BANKINBBNOI]", "gstin": "09[XXXXXXXXXX]1Z[X]", "state": "09",
}


# ---------- GST 계산 헬퍼 (R-GST-01~05) ----------
def r2(x):
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(str(x)).quantize(Decimal("0.01"), ROUND_HALF_UP))


def fin(x):
    """인도식 자릿수 (12,34,567.89)"""
    neg, x = x < 0, abs(r2(x))
    i, d = f"{x:.2f}".split(".")
    if len(i) > 3:
        head, tail = i[:-3], i[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:]); head = head[:-2]
        if head:
            parts.insert(0, head)
        i = ",".join(parts) + "," + tail
    return ("-" if neg else "") + i + "." + d


def charges_text(lines, how):
    tot = r2(sum(a for _, a in lines if a))
    parts = " + ".join(f"{l.split(' on ')[0]} {fin(a)}" for l, a in lines if a)
    return f"INR {fin(tot)} {how} ({parts})"


def gst_fx_value(inr):
    """CGST Rules 32(2)(b) 환전 과세표준"""
    if inr <= 0:
        return 0.0
    if inr <= 100000:
        return r2(max(250.0, inr * 0.01))
    if inr <= 1000000:
        return r2(1000 + (inr - 100000) * 0.005)
    return r2(min(60000.0, 5500 + (inr - 1000000) * 0.001))


def fee_lines(charges, conv_inr, pos, extra=()):
    """charges: [(라벨, 금액)] 과세 수수료, conv_inr: 환전 INR(0=무환전), pos: CS/IG/ZR"""
    lines = list(charges)
    cv = gst_fx_value(conv_inr)
    base = r2(sum(a for _, a in charges) + cv)
    if cv:
        lines.append((f"Currency conversion value u/r 32(2)(b) on INR {fin(conv_inr)}: INR {fin(cv)} (taxable value only, not charged)", None))
    gst = r2(base * 0.18)
    if pos == "CS":
        c = r2(gst / 2)
        lines += [(f"CGST 9% on INR {fin(base)}", c), (f"SGST 9% on INR {fin(base)}", r2(gst - c))]
    elif pos == "IG":
        lines += [(f"IGST 18% on INR {fin(base)}", gst)]
    else:
        lines += [("GST: zero-rated export of service (LUT)", 0.0)]
    return lines + list(extra)


def legs_fixed(target_inr, rate):
    import math
    fcy = math.ceil(target_inr / rate * 100) / 100
    return fcy, r2(fcy * rate)


GST_UP = "Place of supply: Uttar Pradesh (09) · intra-state"
SAC = "SAC 9971xx (foreign exchange services)"

# 비계좌 FDI: 1차 INR 10 crore 확정 환전 + 2차 잔여 환전
_L1_FCY, _L1_INR = legs_fixed(100000000, 88.40)
_TOTAL_FCY = 2000000.00
_L2_FCY = r2(_TOTAL_FCY - _L1_FCY)
_L2_INR = r2(_L2_FCY * 88.47)
_RELAY_FCY, _RELAY_INR = 52000.00, r2(52000 * 88.45)
_FDI_FEES_IG = fee_lines([("Commission (FDI)", 7500.00)], 250000000.70, "IG")
_FDI_FEES_CS = fee_lines([("Commission (FDI)", 7500.00)], r2(_L1_INR + _L2_INR), "CS")

SAMPLES = {
    "firc": {
        "fdi": {
            "cert_no": "FIRC/NOI/2026-27/000041", "issue_dt": "2026-09-03", "fdi_ref": "FDI2609010007", "our_ref": "NOI021RI2609000031",
            "ben_name": "SEOUL PRECISION PARTS INDIA LTD", "ben_cin": "U29304UP2019FTC118254", "ben_pan": "AABCS7710Q",
            "ben_acct": "6015907000412 (Current Account)", "ben_addr": "Plot 14, Sector 34, Gurugram, Haryana 122001, India",
            "rem_name": "SEOUL PRECISION CO LTD", "rem_addr": "88 Gasan Digital 1-ro, Geumcheon-gu, Seoul", "rem_ctry": "KR - Korea, Republic of",
            "rem_bank": "[Remitting bank], Seoul · [BIC]", "their_ref": "FDI-SPC-20260901", "uetr": "4b9e3c51-77a2-4f2e-9d0c-3a1f6c2e8b40",
            "recv_dt": "2026-09-01", "ccy": "USD", "instd": 2850030.00, "fchg": 30.00, "fcy": 2850000.00, "rate": 88.4200,
            "conv": "USD 2,827,414.62 (agreed INR 25,00,00,000.00 at 88.4200, rounded up)",
            "excess": "USD 22,585.38 returned unconverted on 2026-09-03 (ref FDR2609030001), charges USD 25.00 deducted",
            "conv_dt": "2026-09-01", "deal_ref": "MM2608290112 (SPOT)", "inr": 250000000.70,
            "inr_fixed": "Yes - target INR 25,00,00,000.00 agreed; required USD 2,827,414.62 at 88.4200 (rounded up)",
            "charges": charges_text(_FDI_FEES_IG, "debited separately on 2026-09-01; customer GSTIN 06AABCS7710Q1ZT, inter-state"),
            "purpose": "P0006 - Foreign direct investment in India (equity)", "instr": "Equity Shares", "route": "Automatic Route",
            "kyc": "Requested by camt.110 on 2026-09-01; report from remitting bank received 2026-09-02, verified", "allot_due": "2026-10-31", "refund_due": "2026-11-15",
            "copy": "ORIGINAL", "verify": "7F3A-91C2-0D5E-B844"},
        "fdi_rtgs": {
            "cert_no": "FIRC/NOI/2026-27/000047", "issue_dt": "2026-10-09", "fdi_ref": "FDI2610050002", "our_ref": "NOI035RI2610000009",
            "ben_name": "GANGNAM MOBILITY INDIA PVT LTD", "ben_cin": "U34100KA2025FTC201447", "ben_pan": "AAJCG5512M",
            "ben_acct": "Not an account holder with us", "ben_addr": "12th Floor, Prestige Tower, MG Road, Bengaluru, Karnataka 560001, India",
            "rem_name": "GANGNAM MOBILITY CO LTD", "rem_addr": "301 Teheran-ro, Gangnam-gu, Seoul", "rem_ctry": "KR - Korea, Republic of",
            "rem_bank": "[Remitting bank], Seoul · [BIC]", "their_ref": "GM-FDI-2610-01", "uetr": "9d27e4b0-1c5a-4f36-8e02-7b4d3a1c6f58",
            "recv_dt": "2026-10-05", "ccy": "USD", "instd": _TOTAL_FCY + 25, "fchg": 25.00, "fcy": _TOTAL_FCY,
            "rate": "Two tranches - see conversion details", "conv_dt": "2026-10-06 and 2026-10-08", "deal_ref": "MM2610060031 / MM2610080044",
            "inr": r2(_L1_INR + _L2_INR),
            "inr_fixed": "Tranche 1 fixed at INR 10,00,00,000.00 (USD rounded up); balance converted at market rate",
            "conv_legs": [("2026-10-06", _L1_FCY, 88.40, _L1_INR, "MM2610060031 fixed"), ("2026-10-08", _L2_FCY, 88.47, _L2_INR, "MM2610080044")],
            "payout": f"RTGS · UTR [UTR] · IFSC [XXXX0XXXXXX] · A/c XXXXXXXX7731 in the name of GANGNAM MOBILITY INDIA PVT LTD · INR {fin(_L1_INR + _L2_INR)} on 2026-10-08",
            "charges": charges_text(_FDI_FEES_CS, "paid separately by the investee by NEFT on 2026-10-08; non-account holder, place of supply UP"),
            "purpose": "P0006 - Foreign direct investment in India (equity)", "instr": "Equity Shares", "route": "Automatic Route",
            "kyc": "Requested by camt.110 on 2026-10-05; report from remitting bank received 2026-10-06, verified", "allot_due": "2026-12-04", "refund_due": "2026-12-19",
            "copy": "ORIGINAL", "verify": "3E7B-08D1-C64A-29F5"},
    },
    "out": {
        "imp": {"adv_no": "ORA/NOI/2026-27/018204", "adv_dt": "2026-10-31", "our_ref": "NOI013RO2610000184", "txn_type": "RT13 · 01 Advance Import Remittance",
                "value_dt": "2026-10-31", "uetr": "e1a7c0d2-5b3f-4e8a-a9d1-0f62b7c4d915",
                "rem_name": "HANBIT AGRO INDIA PVT LTD", "cust_no": "00000038", "rem_acct": "XXXXXXXXX0250 (Current Account)", "rem_pan": "AAACH****K",
                "rem_iec": "AAACH4821K", "rem_lei": "5493001KJTIIGC8Y1R12",
                "ben_name": "HANBIT SEED CO LTD", "ben_acct": "301-0221-7781-21", "ben_addr": "45 Seed Valley-ro, Anseong, KR", "ben_bank": "[Beneficiary bank], Seoul · [BIC]",
                "intm": "[Intermediary bank], New York · [BIC]", "ccy": "USD", "amt": 100000.00, "rate": 88.4200, "inr": 8842000.00, "chg_type": "OUR",
                "fees": fee_lines([("Commission (GL 40570010000)", 11052.50), ("Cable charge (GL 40570030000)", 1000.00)], 8842000.00, "CS"),
                "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Customer GSTIN 09AAACH4821K1Z5 · {GST_UP} · {SAC}",
                "purpose": "S0101 - Advance payment against imports", "rmt_info": "ADVANCE PAYMENT PI INV-2026-0912",
                "form_a2": "Received", "inv_ref": "INV-2026-0912 dated 2026-09-12 (Proforma)", "orm": "0291402ORM2610310184",
                "evid_due": "2027-05-15", "f15ca": "Not applicable (import of goods)", "copy": "ORIGINAL", "verify": "2C81-6D0F-A4E9-1173"},
        "lrs": {"adv_no": "ORA/NOI/2026-27/018219", "adv_dt": "2026-10-31", "our_ref": "NOI012RO2610000077", "txn_type": "RT12 · LRS - Education",
                "value_dt": "2026-10-31", "uetr": "0c5d2e91-3f4a-4b77-8e6d-21a9f0b3c6e2",
                "rem_name": "PRIYA MEHTA", "cust_no": "00000107", "rem_acct": "XXXXXXXXX0871 (Savings Account)", "rem_pan": "BQRPM****L",
                "ben_name": "UNIVERSITY OF TORONTO", "ben_acct": "00012-004-1234567", "ben_addr": "27 King's College Circle, Toronto, CA", "ben_bank": "[Beneficiary bank], Toronto · [BIC]",
                "intm": "-", "ccy": "USD", "amt": 45000.00, "rate": 88.4200, "inr": 3978900.00, "chg_type": "OUR",
                "fees": fee_lines([("Commission", 500.00), ("Cable charge", 1000.00)], 3978900.00, "CS", [("TCS 5% u/s 206C(1G) (not subject to GST)", 198945.00)]),
                "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Customer unregistered, address in Uttar Pradesh · {GST_UP} · {SAC}",
                "purpose": "S0305 - Travel for education (tuition fees)", "rmt_info": "TUITION FEE STUDENT ID 1009876543",
                "form_a2": "Received, with declaration of FY remittances", "lrs_fy": "FY 2026-27 · USD 239,500.00 of USD 250,000.00",
                "tcs": "5% (education, own funds) · INR 1,98,945.00", "copy": "ORIGINAL", "verify": "9B02-E7A1-55C3-0F6D"},
        "nro": {"adv_no": "ORA/NOI/2026-27/018241", "adv_dt": "2026-10-31", "our_ref": "NOI022RO2610000015", "txn_type": "RT14 · 22 Non-Resident Outward (NRO)",
                "value_dt": "2026-10-31", "uetr": "5a8c1e3f-9b20-4d7e-a6c4-0e3b9f2d7a11",
                "rem_name": "ANAND KRISHNAN (Non-Resident)", "cust_no": "00000233", "rem_acct": "XXXXXXXXX1196 (NRO Savings)", "rem_pan": "CKLPK****D",
                "ben_name": "ANAND KRISHNAN", "ben_acct": "0721-334190-001", "ben_addr": "18 Marina Way, Singapore", "ben_bank": "[Beneficiary bank], Singapore · [BIC]",
                "intm": "-", "ccy": "USD", "amt": 120000.00, "rate": 88.4200, "inr": 10610400.00, "chg_type": "SHA",
                "fees": fee_lines([("Commission", 2500.00), ("Cable charge", 1000.00)], 10610400.00, "CS"),
                "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Non-resident account holder: place of supply = location of supplier (IGST s.13(8)(a)) · {GST_UP} · {SAC}",
                "purpose": "S1305 - Repatriation of sale proceeds of immovable property (code per RBI master)", "rmt_info": "SALE PROCEEDS FLAT 4B NOIDA",
                "form_a2": "Received", "nro_fy": "FY 2026-27 · USD 120,000.00 of USD 1,000,000.00 (sale proceeds)",
                "f15ca": "15CA Part C Ack [Ack No.] · 15CB UDIN [UDIN]", "copy": "ORIGINAL", "verify": "6B4D-2F80-E91A-C357"},
        "fdr": {"adv_no": "ORA/NOI/2026-27/018230", "adv_dt": "2026-10-20", "our_ref": "NOI061RO2610000012", "txn_type": "FD12 · 06 FDI Refund Remittance",
                "value_dt": "2026-10-20", "uetr": "6f1b8d3a-2c9e-4a50-b7f4-98e1d2c0a5b3",
                "rem_name": "SEOUL PRECISION PARTS INDIA LTD", "cust_no": "00000052", "rem_acct": "XXXXXXXXX0412 (Current Account)", "rem_pan": "AABCS****Q",
                "rem_lei": "984500A1B2C3D4E5F627",
                "ben_name": "SEOUL PRECISION CO LTD", "ben_acct": "110-482-339901", "ben_addr": "88 Gasan Digital 1-ro, Seoul, KR", "ben_bank": "[Remitting bank of original FDI], Seoul · [BIC]",
                "intm": "-", "ccy": "USD", "amt": 565482.92, "rate": 88.4200, "inr": 50000000.00, "chg_type": "OUR",
                "fees": fee_lines([("Commission", 5000.00), ("Cable charge", 1000.00)], 50000000.00, "IG"),
                "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Customer GSTIN 06AABCS7710Q1ZT (Haryana) · Place of supply: Haryana (06) · inter-state · {SAC}",
                "purpose": "Refund of FDI (purpose code per RBI master)", "rmt_info": "REFUND SHARE APPLICATION MONEY FDI2608100003",
                "form_a2": "Received", "f15ca": "Not taxable - refund of capital", "fdi_ref": "FDI2608100003 (received 2026-08-10)",
                "refund_rsn": "Not allotted within 60 days", "copy": "ORIGINAL", "verify": "D4E0-3A7B-C915-6E28"},
    },
    "in": {
        "exp": {"adv_no": "IRA/NOI/2026-27/007712", "adv_dt": "2026-10-31", "our_ref": "NOI021RI2610000188", "their_ref": "KEBH261031778812",
                "value_dt": "2026-10-31", "uetr": "8a562c67-ca16-48ba-b074-65581be6f001",
                "ben_name": "SEOUL PRECISION PARTS INDIA LTD", "cust_no": "00000052", "ben_acct": "XXXXXXXXX0412 (Current Account)", "ben_pan": "AABCS****Q", "ben_iec": "AABCS7710Q",
                "rem_name": "DAESUNG FOODS CO LTD", "rem_addr": "21 Teheran-ro, Gangnam-gu, Seoul", "rem_ctry": "KR - Korea, Republic of", "rem_bank": "[Remitting bank], Seoul · [BIC]",
                "ccy": "USD", "instd": 86425.00, "fchg": 25.00, "fcy": 86400.00, "rate": 88.4200, "inr": 7639488.00,
                "fees": fee_lines([("Inward commission", 250.00)], 7639488.00, "IG"),
                "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Customer GSTIN 06AABCS7710Q1ZT (Haryana) · Place of supply: Haryana (06) · inter-state · {SAC}",
                "purpose": "P0103 - Advance receipts against export contracts", "rmt_info": "ADVANCE PO DS-2026-114",
                "irm": "0291402261031000017", "ship_due": "2027-10-31", "copy": "ORIGINAL", "verify": "51AF-08C7-E2D3-9B60"},
        "ind": {"adv_no": "IRA/NOI/2026-27/007718", "adv_dt": "2026-10-31", "our_ref": "NOI021RI2610000191", "their_ref": "FT26304XK0091",
                "value_dt": "2026-10-31", "uetr": "c3e9a1f0-6d2b-4c85-b0a7-5e4f1d8c2a96",
                "ben_name": "PRIYA MEHTA", "cust_no": "00000107", "ben_acct": "XXXXXXXXX0871 (Savings Account)", "ben_pan": "BQRPM****L",
                "rem_name": "RAHUL MEHTA", "rem_addr": "Flat 12, Al Barsha 1, Dubai", "rem_ctry": "AE - United Arab Emirates", "rem_bank": "[Remitting bank], Dubai · [BIC]",
                "ccy": "AED", "instd": 15000.00, "fchg": 0.00, "fcy": 15000.00, "rate": 24.0700, "inr": 361050.00,
                "fees": fee_lines([], 361050.00, "CS"),
                "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Customer unregistered, address in Uttar Pradesh · {GST_UP} · {SAC}",
                "purpose": "P1301 - Inward remittance for family maintenance", "rmt_info": "FAMILY SUPPORT OCT", "copy": "ORIGINAL", "verify": "E06B-4F19-7A2C-D385"},
        "nre": {"adv_no": "IRA/NOI/2026-27/007725", "adv_dt": "2026-10-31", "our_ref": "NOI021RI2610000203", "their_ref": "SGX2610310044",
                "value_dt": "2026-10-31", "uetr": "f2b6d091-4e3c-47a8-9c15-8d0a6e3b2c47",
                "ben_name": "ANAND KRISHNAN (Non-Resident)", "cust_no": "00000233", "ben_acct": "XXXXXXXXX1188 (NRE Savings)", "ben_pan": "CKLPK****D",
                "rem_name": "ANAND KRISHNAN", "rem_addr": "18 Marina Way, Singapore", "rem_ctry": "SG - Singapore", "rem_bank": "[Remitting bank], Singapore · [BIC]",
                "ccy": "USD", "instd": 30000.00, "fchg": 0.00, "fcy": 30000.00, "rate": 88.4200, "inr": 2652600.00,
                "fees": fee_lines([], 2652600.00, "CS"),
                "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Non-resident account holder: place of supply = location of supplier (IGST s.13(8)(a)) · {GST_UP} · {SAC}",
                "purpose": "P1302 - Personal gifts and donations / own funds to NRE (code per RBI master)", "rmt_info": "OWN FUNDS TO NRE",
                "nr_acct": "NRE - repatriable", "copy": "ORIGINAL", "verify": "0D93-B5E7-41C8-FA26"},
        "relay": {"adv_no": "IRA/NOI/2026-27/007731", "adv_dt": "2026-10-31", "our_ref": "NOI035RI2610000014", "their_ref": "HKB26102900871",
                  "value_dt": "2026-10-29", "uetr": "7c3a9e15-0b48-4d62-a1f7-3e5c9b0d8a24",
                  "ben_name": "KAVERI TEXTILES PVT LTD", "cust_no": "WALK-IN", "ben_acct": "Not an account holder with us", "ben_pan": "AAECK****R",
                  "rem_name": "HANA APPAREL CO LTD", "rem_addr": "55 Myeongdong-gil, Jung-gu, Seoul", "rem_ctry": "KR - Korea, Republic of", "rem_bank": "[Remitting bank], Seoul · [BIC]",
                  "ccy": "USD", "instd": _RELAY_FCY + 20, "fchg": 20.00, "fcy": _RELAY_FCY, "rate": 88.4500, "inr": _RELAY_INR,
                  "conv_legs": [("2026-10-31", _RELAY_FCY, 88.45, _RELAY_INR, "MM2610310057")],
                  "fees": fee_lines([("Inward commission (relay)", 500.00), ("RTGS charge", 25.00)], _RELAY_INR, "CS"),
                  "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Customer GSTIN 33AAECK6620R1Z8 · Non-account holder: place of supply = location of supplier (IGST s.12(12)) · {GST_UP} · {SAC}",
                  "purpose": "P0102 - Realisation of export bills (goods)", "rmt_info": "INV KT-2026-0417",
                  "susp": "SP2610290004 · parked 2026-10-29 · reason: no account with us",
                  "payout": "RTGS · UTR [UTR] · IFSC [XXXX0XXXXXX] · A/c XXXXXXXX4410 in the name of KAVERI TEXTILES PVT LTD · 2026-10-31",
                  "copy": "ORIGINAL", "verify": "B81E-6C24-0F5A-D973"},
        "fdi": {"adv_no": "IRA/NOI/2026-27/006590", "adv_dt": "2026-09-01", "our_ref": "NOI021RI2609000031", "their_ref": "FDI-SPC-20260901",
                "value_dt": "2026-09-01", "uetr": "4b9e3c51-77a2-4f2e-9d0c-3a1f6c2e8b40",
                "ben_name": "SEOUL PRECISION PARTS INDIA LTD", "cust_no": "00000052", "ben_acct": "XXXXXXXXX0412 (Current Account)", "ben_pan": "AABCS****Q",
                "rem_name": "SEOUL PRECISION CO LTD", "rem_addr": "88 Gasan Digital 1-ro, Seoul", "rem_ctry": "KR - Korea, Republic of", "rem_bank": "[Remitting bank], Seoul · [BIC]",
                "ccy": "USD", "instd": 2850030.00, "fchg": 30.00, "fcy": 2850000.00, "rate": 88.4200, "inr": 250000000.70,
                "excess": "USD 22,585.38 not converted - returned to remitter (FIRC covers USD 2,827,414.62 only)",
                "fees": _FDI_FEES_IG, "fee_mode": "SEPARATE",
                "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Customer GSTIN 06AABCS7710Q1ZT (Haryana) · inter-state · charges debited separately (principal not reduced) · {SAC}",
                "purpose": "P0006 - Foreign direct investment in India (equity)", "rmt_info": "SHARE SUBSCRIPTION SPPIL", "fdi_ref": "FDI2609010007",
                "copy": "ORIGINAL", "verify": "A7C5-2E90-13FB-6D48"},
    },
}

_FX_ACCT_INR = r2(50000 * 88.38)
_FX_CROSS_KRW = r2(20000 * 1392.40)
_FX_CROSS_INR_EQV = min(r2(20000 * 88.42), r2(_FX_CROSS_KRW * 0.06350))
_FX_CASH_INR = r2(1200 * 86.95)

SAMPLES["fx"] = {
    "acct": {"adv_no": "FXA/NOI/2026-27/002118", "adv_dt": "2026-10-31", "fx_ref": "NOI051FX2610000031", "fx_type": "Account exchange · EEFC (USD) to Current Account (INR)",
             "cust_name": "HANBIT AGRO INDIA PVT LTD", "cust_no": "00000038", "cust_pan": "AAACH****K",
             "sell_acct": "XXXXXXXXX0013 (EEFC USD)", "buy_acct": "XXXXXXXXX0250 (Current Account INR)",
             "sell": "USD 50,000.00", "rate": "88.3800 INR per USD (dealing rate)", "deal_ref": "MM2610310066 (SPOT)",
             "buy": f"INR {fin(_FX_ACCT_INR)}", "buy_inr": _FX_ACCT_INR,
             "fees": fee_lines([], _FX_ACCT_INR, "CS"),
             "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Customer GSTIN 09AAACH4821K1Z5 · {GST_UP} · GST debited from INR account · {SAC}",
             "copy": "ORIGINAL", "verify": "41C7-9E0A-B263-F8D5"},
    "cross": {"adv_no": "FXA/NOI/2026-27/002124", "adv_dt": "2026-10-31", "fx_ref": "NOI052FX2610000007", "fx_type": "Cross-currency · USD to KRW within EEFC",
              "cust_name": "HANBIT AGRO INDIA PVT LTD", "cust_no": "00000038", "cust_pan": "AAACH****K",
              "sell_acct": "XXXXXXXXX0013 (EEFC USD)", "buy_acct": "XXXXXXXXX0021 (EEFC KRW)",
              "sell": "USD 20,000.00", "rate": "1,392.40 KRW per USD (dealing rate)", "deal_ref": "MM2610310071",
              "rbi_ref": f"INR 88.4200 per USD · INR 0.0635 per KRW → lesser INR value {fin(_FX_CROSS_INR_EQV)}",
              "buy": f"KRW {_FX_CROSS_KRW:,.0f}", "buy_inr": None,
              "fees": fee_lines([], _FX_CROSS_INR_EQV, "CS"),
              "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Customer GSTIN 09AAACH4821K1Z5 · {GST_UP} · method B slab applied on INR equivalent (R-FX-07) · GST debited from INR account · {SAC}",
              "copy": "ORIGINAL", "verify": "C2E8-17B4-6A0D-93F1"},
    "cash": {"adv_no": "FXA/NOI/2026-27/002131", "adv_dt": "2026-10-31", "fx_ref": "NOI053FX2610000019", "fx_type": "Foreign currency notes purchased (walk-in)",
             "cust_name": "MIN-JUN PARK", "passport": "M[XXXXXXXX] · Republic of Korea",
             "paid_by": "Cash INR (within cash payment limit)",
             "sell": "USD 1,200.00 (notes)", "rate": "86.9500 INR per USD (cash buying rate)",
             "notes": "USD 100 × 12", "cdf": "Not required (below USD 5,000)",
             "buy": f"INR {fin(_FX_CASH_INR)}", "buy_inr": _FX_CASH_INR,
             "fees": fee_lines([], _FX_CASH_INR, "CS"),
             "gst_info": f"Bank GSTIN 09[XXXXXXXXXX]1Z[X] · Unregistered customer · {GST_UP} · GST deducted from INR paid · {SAC}",
             "copy": "ORIGINAL", "verify": "8D51-F3A6-0C9E-27B4"},
}
