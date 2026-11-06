"""English edition of the HO IFRS gap-analysis reply (Noida Branch) — for review by local finance staff.

Same structures as build_ifrs_gap.py (Korean). build_ifrs_gap.build(lang="en") uses:
  * the content below for the Noida cells, and
  * TPL / SHEETS / FORMULA_SUBS to translate the HO template text, NY pre-filled cells and dropdown lists.
Amounts in INR, as of 2026-03-31 unless stated.
"""

BR = "Noida Branch"
BASE = "2026-03-31"
TBC = "To be confirmed"
CHECKING = "Checking"

# ───────────────────────── Summary sheet, column I ─────────────────────────
SUMMARY = {
    5: "Local GAAP",
    6: "Indian GAAP: ICAI Accounting Standards (AS) + RBI Directions (IRACP, Investment Directions 2023, "
       "ECL Directions 2026) + formats of the Third Schedule to the Banking Regulation Act. "
       "Note: Ind AS (Indian IFRS) is deferred for banks and not applied",
    7: "Partially performed",
    8: "Current: loan-loss allowance (as noted by HO). Also needed: (1) EIR deferral of loan processing fees "
       "(recognised upfront locally) (2) IFRS 16 leases (AS 19 operating leases locally) (3) PV of lease security "
       "deposit (4) restoration provision (5) actuarial gains/losses to OCI (AS 15) (6) deferred tax (AS 22 → IAS 12). "
       "7 items in detail in 2_IFRS_Adjustments",
    9: "Not performed",
    10: "Not performed",
    11: "Performed",
    12: "Performed",
    13: "No (working-capital line, short-term and term loans are all plain principal-and-interest. "
        "No subordination, conversion or profit-linked terms)",
    14: "4 types, 12 accounts with FY25-26 income: (1) loans (LOAN APPLICATION FEE 42020701, OTHER FEE 42020729) "
        "(2) guarantees (GUARANTEE(Bond) 42010013) (3) remittance (42020601/602/603/606/629) "
        "(4) reimbursement, extension, postage (42020401/402/403, 42020307)",
    15: "1 deferred / 11 upfront. Guarantee commission (42010013) is taken to unearned income (20920017 Income On "
        "Guarantee Money, ₹72,838 at 2026-03-31 → ₹39,021 at 2026-10-31 with guarantees unchanged) and spread over the guarantee period — same as Indian GAAP AS 9 "
        "(recognised as the service is performed) and IFRS 9 financial guarantees. Others are recognised when received "
        "(AS 9). Loan processing fees form part of the EIR under IFRS 9 and should be deferred",
    16: "Fee income",
    17: "Straight-line",
    18: "Incl. rent-free",
    19: "Registered lease deed (stamped 2022-08-02): Commencement (possession) 2022-04-13 / Rent Commencement "
        "2022-08-20 (rent-free 4 months 7 days per the definitions, or earlier if operations started earlier) / "
        "Expiration 2031-04-12 (9 years, cl.3(a)). Rent ₹14,55,487 + GST, payable in advance by the 7th (cl.4(a)(b))",
    20: "N/A",
    21: "N/A",
    22: "Yes (cl.24 Reinstatement)",
    23: "Not recognised",
    24: "Not yet estimated. A reinstatement obligation exists (cl.24: return in warm-shell condition as handed over for "
        "fit-out), so it must also be provided under Indian GAAP "
        "(AS 29, undiscounted) → to be recognised from a reinstatement quotation. New account 20940000000 in the "
        "next-gen CoA",
    25: "(1) No interest recognised on NPAs (IRACP) (2) guarantees as financial guarantee contracts / ECL; nature of the "
        "₹30.01 crore Reserved Credit Line Commitment at 2026-10-31 to be confirmed (3) 25% statutory reserve (transfer "
        "within equity) (4) prior-period items under AS 5 (5) AS 19 rent not straight-lined (maintenance charges are not "
        "a lease) (6) gap narrows after RBI ECL/EIR from 2027-04. Detail in 13_Others",
}

# ───────────────────────── 2_IFRS_Adjustments ─────────────────────────
IFRS_ADJ = [
    ("Loan-loss allowance\n(10690001 Allowance For Loan Losses-General / next-gen 10380101000)",
     "RBI IRACP standard method, computed by the branch. 2026-03-31: ₹1,52,70,000 = standard advances ₹306.75 crore "
     "× 0.4% ₹1,22,70,000 + incremental provision for unhedged foreign-currency exposure (UFCE) ₹30,00,000 "
     "(all standard, no NPA)",
     "IFRS 9 expected credit loss (Stage 1/2/3, 12-month/lifetime). Includes bank placements, call loans, "
     "HTM bonds, off-balance-sheet commitments and guarantees",
     "Performed", "Global Business Dept", "Computed by HO", BASE,
     "Per HO note in Summary 2-1 (ECL computed by HO). From 2027-04-01 the RBI ECL Directions 2026 move the local "
     "books to stage-based ECL (with regulatory floors) → only floor/model differences remain. The UFCE provision "
     "has no IFRS 9 equivalent and is replaced by ECL (next-gen decision L8: absorbed into stage allowance, UFCE tag). "
     "Next-gen CoA stage L5 10380103000~105000"),
    ("Loan processing fees\n(42020701 LOAN APPLICATION FEE / next-gen 40580010000, deferral LOF 10370100000)",
     "Recognised upfront when received (AS 9). FY25-26 ₹26,40,000; FY26-27 (to 07-07) ₹16,40,000",
     "Deferred as part of the effective interest rate and amortised by EIR, presented as interest income "
     "(IFRS 9 B5.4.1–B5.4.3)",
     "Not performed", "Branch", "To be computed (unamortised balance per loan)", BASE,
     "NY issue 4 applies (larger gap than NY: no deferral at all). From 2027-04-01 EIR under the RBI ECL "
     "Directions 2026 moves the local books to deferral; until then an HO adjustment is needed. "
     "GBS to produce LOF deferral and EIR amortisation data"),
    ("Rent (leases)\n(52032421 RENTAL EXPENSE / next-gen 50940021405) — 1 office, 5 staff residences",
     "AS 19 operating leases: rent expensed as billed, not straight-lined (FY25-26 ₹3,88,82,301), "
     "future minimum lease payments disclosed. No right-of-use asset or lease liability",
     "IFRS 16: right-of-use asset and lease liability, interest expense and depreciation. "
     "Optional exemptions for leases of 12 months or less and low-value assets",
     "Not performed", "TBD", "Estimated lease liability (office) ₹8,21,57,515", BASE,
     "The HO lease template is submitted but marked 'N/A' (no IFRS 16 calculation) → IFRS 16 not reflected in HO "
     "consolidation. Estimate = PV of the remaining 60 office payments of ₹10,21,14,084 at 9% (assumed rate) — "
     "materiality to be assessed. Local ledger stays on Indian GAAP (Ind AS 116 not applied to banks). Next-gen GBS "
     "keeps IFRS 16 accounts (ROU 107, lease liability 206, interest 50320500000) for HO adjustments only and produces "
     "the monthly schedule from lease-register attributes"),
    ("Lease security deposit\n(10999002 Security Deposit / next-gen 10490700100)",
     "Carried at nominal, not discounted. Security Deposit ₹1,31,30,642 includes the office interest-free refundable "
     "deposit (IFRSD) ₹87,32,922 (deed cl.10(a))",
     "Amortised cost (present value). Difference between nominal and PV added to ROU asset; "
     "unwinding is interest income (IFRS 9 / IFRS 16)",
     "Not performed", "TBD", "To be computed", BASE,
     "Adjust together with leases. Next-gen HO-only accounts 10490700900 (discount) and 40140100000 (interest income)"),
    ("Restoration provision\n(no current account / next-gen 20940000000 new)",
     "Reinstatement obligation exists under the lease; not recognised (not yet estimated). AS 29 also requires a "
     "best estimate, undiscounted",
     "IAS 37 at present value, added to ROU asset cost, unwinding as interest expense",
     "Not performed", "Branch", "To be computed", BASE,
     "Fix under Indian GAAP — estimate from a reinstatement quotation and provide in the local books under AS 29 "
     "(leasehold improvements at cost ₹4,78,38,399 for reference). HO IFRS adjustment then covers only discounting and "
     "the ROU addition. Unwinding 50320200000 is HO-adjustment only"),
    ("Gratuity\n(WELFARE EXPENSE 52032xxx / next-gen 50930170100·200)",
     "AS 15: actuarial valuation (liability 2026-03-31 ₹13,25,421); all actuarial gains/losses in profit or loss",
     "IAS 19: remeasurements (actuarial gains/losses) in OCI",
     "Not performed", "Global Business Dept", "Actuarial gain/loss per actuary's report", BASE,
     "Only a P&L↔OCI classification difference, no effect on net assets. Low materiality"),
    ("Deferred tax\n(23100101 DEFERRED Tax Liabilities / next-gen 20720xxxxxx)",
     "AS 22: timing differences (income-statement approach). Net DTL 2026-03-31 ₹20,54,203 (DTL on depreciation; "
     "DTA on loan-loss provision and gratuity), rate 37.128%",
     "IAS 12: temporary differences (balance-sheet approach). Add deferred tax on IFRS adjustments "
     "(leases, EIR, restoration provision)",
     "Not performed", "Global Business Dept", "Per adjustment × 37.128%", BASE,
     "Knock-on effect of the IFRS adjustments. Lease deferred-tax accounts 10820400000 / 20720200000 are HO-only"),
]

# ───────────────────────── 3_Fin_Instruments ─────────────────────────
_LOAN_NOTE = ("No formal test. A desk review of the terms indicates SPPI is met (no cash flows other than principal "
              "and interest). Penal charges are not interest under RBI rules")
FIN = [
    ("Loans", "Working-capital demand loan (WCDL) 10610004", "Amortised cost (IRACP)", "Not performed", "Met",
     "Not performed", "Hold to collect", "AC", 600000000, "None", _LOAN_NOTE),
    ("Loans", "Short-term working-capital loans 10610005", "Amortised cost (IRACP)", "Not performed", "Met",
     "Not performed", "Hold to collect", "AC", 1480000000, "None", _LOAN_NOTE),
    ("Loans", "Term loans (equipment) 10610007", "Amortised cost (IRACP)", "Not performed", "Met",
     "Not performed", "Hold to collect", "AC", 987500000, "None", _LOAN_NOTE),
    ("Placements", "Term deposits with domestic banks 10255002", "Amortised cost", "Not performed", "Met",
     "Not performed", "Hold to collect", "AC", 2018960000, "None",
     "Inter-bank term deposits (about 35% of total assets). In next-gen ECL scope (decision L6)"),
    ("Placements", "Call loans - domestic 10251001", "Amortised cost", "Not performed", "Met", "Not performed",
     "Hold to collect", "AC", 100000000, "None", "Overnight inter-bank lending"),
    ("Placements", "RBI current account and statutory deposit (BR Act s.11(2)(b)) 10211001·10211002",
     "Amortised cost", "N/A", "N/A", "N/A", "Hold to collect", "AC", 27388808, "None",
     "Includes statutory deposit ₹15,00,000. Central-bank balance, no ECL effect"),
    ("Placements", "Nostro on demand (domestic and overseas) 10212001·10212002", "Amortised cost", "N/A", "N/A",
     "N/A", "Hold to collect", "AC", 20358575, "None", "Settlement accounts"),
    ("Securities", "Government of India securities, HTM (SLR) 10330001", "HTM (RBI Investment Directions 2023)",
     "Performed", "Met", "Performed", "Hold to collect", "AC", 56886300, "None",
     "RBI Investment Directions 2023: HTM requires SPPI cash flows and intent to hold to maturity, with sale "
     "limits → in substance the same as IFRS 9 AC"),
    ("Bills purchased", "Export bills purchased / discounted", "Amortised cost (IRACP)", "N/A", "N/A", "N/A",
     "N/A", "AC", 0, "None", "No balance at 2026-03-31. Expected AC when written (short-term, principal and interest)"),
    ("Other", "Guarantees (bond guarantees, off-balance-sheet)", "Contingent liability (Sch.12)", "N/A", "N/A",
     "N/A", "N/A", "N/A", 5695709, "-",
     "IFRS 9 financial guarantee: initially at the fee, then higher of ECL and unamortised amount. See 13_Others no.2"),
    ("Other", "Loan commitment (Reserved Credit Line Commitment, off-balance-sheet)", "Commitment (off-B/S, 62040001)",
     "N/A", "N/A", "N/A", "N/A", "N/A", 300104291, "-",
     "Balance at 2026-10-31 (nil in Sch.12 'other items' at 2026-03-31). If an undrawn sanctioned limit, within IFRS 9 "
     "commitment ECL. See 13_Others no.2"),
]

# ───────────────────────── 4-7_COMMISSION ─────────────────────────
COMM = [
    ("42020701", "COMMISSION - LOAN APPLICATION FEE", "Loan application / processing fee",
     "Corporate loans (WCDL, short-term, term)", "Upfront", "Fee income", 2640000,
     "Part of EIR under IFRS 9 → defer, amortise by EIR, present as interest income (2_IFRS_Adjustments no.2). "
     "RBI EIR expected from 2027-04"),
    ("42020729", "COMMISSION - OTHER FEE", "Other loan-related fees", "Corporate loans", "Upfront", "Fee income",
     217781, "Check the nature of each item; defer if part of EIR"),
    ("42010013", "COMMISSION-GUARANTEE(Bond)", "Performance / bid bond issuance commission", "Guarantees",
     "Deferred", "Fee income", 38716,
     "Received into unearned income (20920017 Income On Guarantee Money) and recognised as fee income over the "
     "guarantee period — same as Indian GAAP AS 9 (service performed over time) and IFRS 9 financial guarantees. "
     "Unearned balance fell to ₹39,021 at 2026-10-31 with guarantees unchanged at ₹56,95,709 — spreading confirmed. "
     "Next-gen unearned income 20420170000, daily straight-line",
     dict(H="Fee income", J="Straight-line", K="Guarantee period", M=72838)),
    ("42020601", "COMMISSION - REMITTANCE OF OUTWARD", "Outward remittance handling fee", "Remittance", "Upfront",
     "Fee income", 193250, "Recognised when the service is rendered — same as IFRS 15"),
    ("42020606", "Commission - Remittance Of Outward(OUR)", "Recovery of intermediary-bank charges on OUR remittances",
     "Remittance", "Upfront", "Fee income", 1557,
     "Amount differs from the actual intermediary charge, so taxable income as principal (gross) — same as IFRS 15"),
    ("42020602", "COMMISSION - REMITTANCE OF INWARD", "Inward remittance fee", "Remittance", "Upfront",
     "Fee income", 1200, "No difference"),
    ("42020603", "COMMISSION - CABLE AND POSTAGE(REMITTANCE)", "SWIFT charges on remittances", "Remittance",
     "Upfront", "Fee income", 18000, "No difference"),
    ("42020629", "COMMISSION - OTHERS(REMITTANCE)", "Other remittance-related fees", "Remittance / FX", "Upfront",
     "Fee income", 1126026,
     "No difference. Next-gen splits into handling / intermediary / SWIFT (CoA review decision)"),
    ("42020401", "COMMISSION - REIMBURSEMENT", "Import L/C reimbursement commission", "Trade finance", "Upfront",
     "Fee income", 85388, "No difference"),
    ("42020402", "COMMISSION - CABLE AND POSTAGE(REIMBURSEMENT)", "Cable charges on reimbursements",
     "Trade finance", "Upfront", "Fee income", 14354, "No difference"),
    ("42020403", "COMMISSION - EXTENTION", "Extension fee on acceptance / reimbursement maturity", "Trade finance",
     "Upfront", "Fee income", 22921, "Review as part of EIR if it relates to extending credit (small amount)"),
    ("42020307", "COMMISSION - POSTAGE/DHL(EXPORT)", "Recovery of courier charges on export documents",
     "Trade finance", "Upfront", "Fee income", 1000, "No difference"),
    ("42020704 etc.", "LOAN COMMITMENT FEE, SYNDICATED LOAN PARTICIPATION/ARRANGE FEE, UNUSED LINE FEE etc.",
     "Commitment / syndication fees", "Syndicated loans / commitments", None, None, 0,
     "No income in FY25-26. If earned, defer commitment fees depending on likelihood of drawdown "
     "(IFRS 9 B5.4.3) — next-gen LOF 10370100000"),
]
COMM_PREFIX = "FY25-26 income."

# ───────────────────────── 8-12_Lease_Restoration ─────────────────────────
_NO_TPL = "N/A (template submitted marked 'N/A')"
_RES = "Residential; restoration cost expected to be minor"
LEASE = [
    dict(C="Noida Branch office premises (registered lease deed)", D="Office",
         E="2022-04-13 (possession for fit-out, cl.3(a))", F="4 months 7 days",
         G="2022-08-20 (definitions: 4 months 7 days after commencement, or start of operations if earlier)",
         H="2031-04-12 (9 years, cl.3(a))", I="Incl. rent-free", J=_NO_TPL, K=_NO_TPL,
         M="Cl.4(f): two 3-yearly steps — +10% after 3 years from Rent Commencement (₹14,55,487→₹16,01,036), +10.5% "
           "after 6 years (→₹17,69,145). Cl.4(f) counts from Rent Commencement (2025-08-20, 2028-08-20) but the "
           "definition counts from Lease Commencement (2025-04-13, 2028-04-13) — branch applies the commencement basis "
           "(2025-04, 2028-04)",
         N=_NO_TPL,
         P="Lessee lock-in 36 months (to 2025-04-12), then may terminate on 3 months' notice (cl.12(b)). Lessor cannot "
           "terminate except for breach (cl.12(d)). Renewal only by a fresh, mutually agreed deed (cl.13) → not a renewal "
           "option. Full 9 years used (termination not expected)",
         Q="Yes (cl.24 Reinstatement; cl.12(a)(b))", R="Not recognised", S="Not yet estimated",
         T="Not yet estimated — also required under AS 29; to be recognised from a reinstatement quotation",
         U="AS 19 straight-lining not applied (rent expensed as billed, 13_Others no.5). Lease template "
           "submitted marked 'N/A' → IFRS 16 not reflected (2_IFRS_Adjustments no.3). Interest-free deposit ₹87,32,922 "
           "(cl.10(a), 6 months' rent). Maintenance under a separate agreement. Leasehold improvements at cost ₹4,78,38,399"),
    dict(C="Staff residence 1", D="Residence", E="2024-08 (first rent month)", H="2026-12 (schedule)", J=_NO_TPL,
         M="2025-07 monthly ₹3,23,000→₹3,25,000", P="-", Q="Residential (minor)", R="Not recognised", S="-", T=_RES,
         U="Period = first and last rent month in the rent schedule (contract dates not needed as IFRS 16 is not applied)"),
    dict(C="Staff residence 2", D="Residence", E="2025-02 (first rent month)", H="2025-12 (ended)", J=_NO_TPL,
         M="No escalation (₹3,00,000/month)", P="-", Q="-", R="N/A", S="-", T="-",
         U="Ended. FY25-26 expense only"),
    dict(C="Staff residence 3", D="Residence", E="2024-04 (first rent month)", H="2026-03 (ended)", J=_NO_TPL,
         M="No escalation (₹2,99,000/month)", P="-", Q="-", R="N/A", S="-", T="-",
         U="Ended. FY25-26 expense only"),
    dict(C="Staff residence 4", D="Residence", E="2024-03 (first rent month)", H="2026-11 (schedule)", J=_NO_TPL,
         M="No escalation (₹2,94,000/month)", P="-", Q="Residential (minor)", R="Not recognised", S="-", T=_RES,
         U="Period = first and last rent month in the rent schedule"),
    dict(C="Staff residence 5", D="Residence", E="2026-02 (first rent month)", H="2026-12 (schedule)", J=_NO_TPL,
         M="No escalation (₹3,83,500/month)", P="-", Q="Residential (minor)", R="Not recognised", S="-", T=_RES,
         U="Period = first and last rent month in the rent schedule"),
    dict(C="Maintenance charges (SUNDRY OFFICE RENTAL EXPENSE 52032261)", D="Maintenance", E="-", H="-", J="N/A",
         K="N/A", P="-", Q="-", R="N/A", S="-", T="-",
         U="Not a lease: maintenance of common areas and the office (electricity excluded). Non-lease (service) "
           "component under IFRS 16, expensed; no adjustment. FY25-26 ₹30,08,161. Next-gen: move out of the 'RENTAL' "
           "account into a maintenance-charges account (maintenance deposit 10490700200)"),
]

# ───────────────────────── 13_Others ─────────────────────────
OTHER = [
    ("Loan interest income / accrued interest", "Interest on NPAs",
     "RBI IRACP: on NPA classification (e.g. 90 days past due) accrued interest is reversed, then cash basis",
     "IFRS 9 Stage 3: interest continues on net carrying amount × EIR", "0 (no NPA at 2026-03-31)",
     "HO adjustment if it arises. GBS keeps unrecognised interest as a memo attribute",
     "Structural difference only; no amount now"),
    ("Guarantees and undrawn commitments (off-balance-sheet)", "Allowance for financial guarantees and loan commitments",
     "Disclosed as contingent liabilities (Sch.12; guarantees ₹56,95,709). No guarantee provision (standard). "
     "Commitments: nil at 2026-03-31 (Sch.12 other items 0); Reserved Credit Line Commitment ₹30,01,04,291 at "
     "2026-10-31 (off-B/S 62040001). IRACP sets no provision on undrawn commitments",
     "IFRS 9 financial guarantee: initially at the fee, then higher of ECL and unamortised amount. ECL on undrawn "
     "loan commitments",
     "Guarantees 5,695,709 / commitment 300,104,291 (2026-10-31)",
     "Include guarantees and commitments in HO ECL. Next-gen ECL scope includes off-balance-sheet guarantees and "
     "commitments (decision L6), with CCF attributes",
     "To be confirmed: nature of the ₹30.01 crore commitment (undrawn sanctioned limit?)"),
    ("Statutory reserve (next-gen 31110300000)", "Transfer to statutory reserve",
     "BR Act s.17: 25% of net profit to statutory reserve (appropriation). FY25-26 ₹54,20,250",
     "Transfer within retained earnings in equity; no P&L effect", 5420250,
     "Present within retained earnings on HO consolidation", "Presentation only"),
    ("Prior-period items (next-gen 40890000000·51208000000)", "Prior-period errors",
     "AS 5: shown separately as 'prior period items' in current-year P&L",
     "IAS 8: material errors corrected retrospectively (restate comparatives)", "-",
     "Restate HO comparatives if a material item arises", "-"),
    ("Rent 52032421", "AS 19 straight-lining not applied",
     "Monthly billed rent is expensed as is. Rent-free 4 months 7 days and 3-yearly escalations are not equalised — "
     "no lease-equalisation liability in the 2026-10-31 BS",
     "Resolved under IFRS 16 (ROU depreciation and interest replace rent)",
     "Cumulative gap about ₹56–90 lakh at 2026-03-31 (depending on the actual rent start)",
     "AS 19 para 23 requires straight-line expense over the lease term → Indian GAAP fix. Agree timing and treatment "
     "of the cumulative amount with the auditor; next-gen lease register produces the straight-line schedule and "
     "equalisation liability",
     "FY25-26 rent ₹3,88,82,301. Straight-line ₹15,51,550/month (deed terms) vs. current billing ₹16,01,036"),
    ("Loan-loss allowance / loan fees", "RBI ECL and EIR from 2027-04-01",
     "RBI ECL Directions 2026: stage-based ECL (regulatory floors) and EIR amortisation of loan fees. "
     "First-time impact through retained earnings",
     "IFRS 9 already applied", "-",
     "Recompute the gap at 2027-03-31 and reduce HO adjustments (allowance and EIR adjustments shrink to floor "
     "differences). Next-gen transition account 31110800000",
     "Local ledger converges on IFRS"),
    ("FX translation, HTM bonds, remittance/trade fees", "Items confirmed as no difference",
     "AS 11 FX translation (closing rate), RBI Investment Directions 2023 HTM at amortised cost, "
     "remittance/trade fees recognised when the service is rendered",
     "Same as IAS 21, IFRS 9 AC and IFRS 15", "-", "No adjustment needed", "Confirmed"),
]

# ───────────────────────── NY_Issues, column G ─────────────────────────
NY_ISSUE = {
    5: ("O", "Noida: Indian GAAP books; no conversion other than the loan-loss allowance"),
    6: ("O", "Noida: no SPPI framework (loans are plain principal and interest, no material gap). "
             "HTM G-secs follow the SPPI condition in the RBI 2023 Directions"),
    7: ("X", "Noida: no modified loans (no NPA or restructuring at 2026-03-31)"),
    8: ("O", "Noida: loan processing fees recognised upfront (no deferral), FY25-26 ₹26,40,000 → "
             "HO entry 'X' corrected to 'O'"),
    10: ("X", "Noida: lease template submitted marked 'N/A' (no IFRS 16 calculation) — no term mismatch, but not "
              "reflecting IFRS 16 is itself a gap (2_IFRS_Adjustments no.3)"),
    11: ("X", "Noida: template 'N/A'. Escalation clause reflected in AS 19 straight-line rent"),
    12: ("O", "Noida: reinstatement obligation exists, no restoration provision → to be recognised under AS 29 "
              "from a quotation"),
}

# ───────────────────────── Review request sheet (local staff) ─────────────────────────
REVIEW_TITLE = "Review request — Noida Branch reply to HO IFRS gap analysis"
REVIEW_INTRO = (
    "HO (Global Business Department, with Deloitte Anjin as reviewer) asked each overseas branch to report the "
    "differences between local GAAP and IFRS used for HO's K-IFRS consolidated statements. This English copy is the "
    "draft reply of the Noida Branch. Please review the facts and the Indian GAAP treatment, and answer the questions "
    "below. Cells shaded yellow in the other sheets still need checking against contracts or records. "
    "Source: FY25-26 audited financials (2026-03-31 TB, AS 19 working, deferred tax computation, asset quality, Sch.12) "
    "and GBS balances at 2026-07-07. Amounts in INR."
)
REVIEW_HEAD = ("No.", "Sheet / cell", "Question for local review", "Draft answer in this file", "Reviewer response",
               "Reviewer / date")
REVIEW = [
    ("Summary I19, 8-12 row 6 (F, G), 13_Others no.5",
     "Actual rent start: the deed gives 4 months 7 days rent-free (to 2022-08-19), or until operations start if earlier. "
     "The AS 19 rent schedule in the FY25-26 file shows rent from 2022-04-13. What was the date of the first rent "
     "invoice? It sets the straight-line gap (about ₹90 lakh if rent started 2022-08-20, about ₹56 lakh if 2022-04-13).",
     "Deed basis: Rent Commencement 2022-08-20."),
    ("13_Others no.5",
     "Straight-lining is not applied (no equalisation liability in the 2026-10-31 BS). Discuss with the auditor whether to "
     "book the cumulative equalisation liability and from when.",
     "Reported as an Indian GAAP gap."),
    ("13_Others no.2, 3_Fin_Instruments (commitment row)",
     "The 2026-10-31 BS shows 62040001 RESERVED CREDIT LINE COMMITMENT ₹30,01,04,291 off-balance-sheet. The branch said "
     "all limits are fully drawn. What is this line (undrawn sanctioned limit, or something else)?",
     "Shown as a commitment at 2026-10-31, nature to be confirmed."),
    ("2_IFRS_Adjustments (leases)",
     "RENTAL EXPENSE for Apr–Oct 2026 is ₹3,31,79,797 (about ₹47 lakh/month), but known rents (office + 3 residences) "
     "are about ₹26 lakh/month. Which other contracts or charges are in this account (new residences, GST, arrears)?",
     "Not reflected; lease list may be incomplete."),
    ("8-12 row 6 (S), 2_IFRS_Adjustments (restoration)",
     "Reinstatement cost estimate (quotation) for the AS 29 provision — removal of fit-outs back to warm shell (cl.24).",
     "Obligation confirmed by cl.24; amount not estimated."),
    ("Summary I7–I8, 2_IFRS_Adjustments (allowance)",
     "Does HO compute IFRS 9 ECL for Noida, or does the branch send any ECL figure?",
     "Per HO's own note: 'Performed / Global Business Dept'."),
    ("2_IFRS_Adjustments (leases)",
     "The lease template goes to HO marked 'N/A'. Should we give IFRS 16 figures this time? Estimated office lease "
     "liability ₹8,21,57,515 (9% assumed rate) — HO/Deloitte may consider it material.",
     "Left as 'Not performed / TBD' with the estimate."),
]
REVIEW_DONE_TITLE = "Confirmed by the branch (2026-11-06)"
REVIEW_DONE = [
    ("Office lease", "Rent expensed as billed (2026-10-31 BS has no equalisation liability). From the registered deed: term 2022-04-13 to 2031-04-12 (cl.3(a)); "
                     "rent-free 4 months 7 days; +10% / +10.5% after 3 and 6 years (cl.4(f)); lock-in 36 months; "
                     "reinstatement to warm shell (cl.24); IFRSD ₹87,32,922 (cl.10(a))."),
    ("Lease template", "Submitted to HO but marked 'N/A'."),
    ("Other rentals", "SUNDRY OFFICE RENTAL EXPENSE is maintenance of common areas and the office, excluding electricity — not a lease."),
    ("Staff residences", "Exact contract dates not needed: IFRS 16 is not applied; rent-schedule months are shown."),
    ("Loan processing fees", "Recognised upfront, not deferred."),
    ("Loan-loss allowance", "Branch computes standard assets × 0.4% (standard method); the extra ₹30,00,000 is the UFCE "
                            "incremental provision."),
    ("Guarantee commission", "Spread over the guarantee period: unearned income 20920017 fell from ₹72,838 "
                             "(2026-03-31) to ₹39,021 (2026-10-31) with guarantees unchanged."),
]

# ───────────────────────── template translation ─────────────────────────
SHEETS = {
    "작성안내": "Instructions",
    "총괄표": "Summary",
    "2_IFRS조정": "2_IFRS_Adjustments",
    "3_금융상품분류": "3_Fin_Instruments",
    "4-7_COMMISSION": "4-7_COMMISSION",
    "8-12_리스·복구": "8-12_Lease_Restoration",
    "13_기타": "13_Others",
    "뉴욕지점_이슈처리": "NY_Issues",
}
FORMULA_SUBS = [('"*기재 요망*"', '"*to be provided*"'), ('"확인 필요"', '"To be confirmed"')]

TPL = {
    # Instructions
    "해외지점 및 해외종속법인 IFRS 관련 확인 요청서": "IFRS Confirmation Request — Overseas Branches and Subsidiaries",
    "별첨 1 I 작성 주관: 글로벌사업부 I 작성 기준 : 회신일 현재":
        "Attachment 1 | Owner: Global Business Department | As of: reply date",
    "1. 작성 목적": "1. Purpose",
    "뉴욕지점 재무제표 점검(2026년 10월) 결과 확인된 사항이 다른 해외지점 및 종속법인에도 공통적으로 존재하는지 확인하고, 본점 재무제표(IFRS) 작성을 위한 GAAP 차이 및 시정이 필요한 사항을 사전에 파악하기 위함입니다.":
        "To check whether the findings from the review of the New York Branch financial statements (October 2026) also "
        "exist at other overseas branches and subsidiaries, and to identify in advance the GAAP differences and "
        "corrections needed for HO's (IFRS) financial statements.",
    "2. 작성 대상": "2. Scope",
    "뉴욕, 북경, 하노이, 노이다, 런던, 홍콩, 시드니지점(7개) 및 미얀마, 캄보디아 종속법인(2개), 총 9개 대상":
        "New York, Beijing, Hanoi, Noida, London, Hong Kong and Sydney branches (7) and the Myanmar and Cambodia "
        "subsidiaries (2) — 9 in total",
    "3. 시트 구성": "3. Sheets",
    "① 총괄표: 항목 1~13에 대한 지점·법인별 요약 응답\n② 2_IFRS조정: 현지 GAAP과 IFRS 간 조정(conversion) 항목 상세\n③ 3_금융상품분류: 보유 금융상품별 SPPI·사업모형 테스트 수행 내역\n④ 4-7_COMMISSION: CoA 상 수수료(COMMISSION) 계정별 이연/일시인식, 인식계정, 상각방법\n⑤ 8-12_리스·복구: 임대차계약별 리스기간, 템플릿 일치 여부, 임대료 조정시점, 복구의무 및 복구충당부채\n⑥ 13_기타: 그 밖에 IFRS와 차이가 있거나 확인이 필요한 사항\n⑦ 뉴욕지점_이슈처리: 뉴욕지점 점검 이슈별 타 지점 해당 여부 및 처리 주체 의견":
        "① Summary: summary answers to items 1–13 by branch/subsidiary\n"
        "② 2_IFRS_Adjustments: detail of local GAAP → IFRS conversion items\n"
        "③ 3_Fin_Instruments: SPPI and business-model tests by financial instrument held\n"
        "④ 4-7_COMMISSION: deferral/upfront recognition, income account and amortisation method by COMMISSION account in the CoA\n"
        "⑤ 8-12_Lease_Restoration: lease term per lease contract, agreement with template, rent escalation dates, "
        "restoration obligation and provision\n"
        "⑥ 13_Others: other IFRS differences or items needing confirmation\n"
        "⑦ NY_Issues: whether each NY review issue applies to other branches, and who should handle it",
    "4. 작성 방법": "4. How to complete",
    "• 흰색 셀에 입력하며, 선택형 항목은 셀 우측 드롭다운 목록에서 선택해 주시기 바랍니다.\n• 총괄표는 요약 응답이며, 상세 내역은 각 상세 시트에 계정·계약 단위로 기재해 주시기 바랍니다.\n• 해당 사항이 없는 경우 '해당없음'으로 기재하고, 빈칸으로 두지 않도록 해 주시기 바랍니다.\n• 판단이 어렵거나 확인 중인 항목은 '확인 중'으로 기재하고 비고란에 사유를 적어 주시기 바랍니다.":
        "• Enter in white cells; for choice items use the drop-down list at the right of the cell.\n"
        "• The Summary holds summary answers; enter details by account or contract in each detail sheet.\n"
        "• If not applicable, enter 'N/A' — do not leave blanks.\n"
        "• If unclear or still being checked, enter 'Checking' and give the reason in Remarks.",
    "5. 뉴욕지점 사전 기재분": "5. Pre-filled for New York",
    "연녹색 셀은 뉴욕지점 점검 결과를 토대로 미리 기재한 내용입니다. 뉴욕지점에서 내용을 확인하시고, 다른 부분이 있으면 수정해 주시기 바랍니다.":
        "Light-green cells were pre-filled from the New York review. New York Branch to confirm and correct any differences.",
    "6. 첨부 요청 자료": "6. Attachments requested",
    "• 현지 재무제표(최근 결산 기준) 및 본점 합산용 TB\n• 현지 회계정책 또는 결산 매뉴얼(수수료 인식, 리스, 충당부채 관련)\n• 임대차계약서 사본 및 리스 계산 템플릿(사용권자산·리스부채)\n• CoA 중 COMMISSION 관련 계정 목록":
        "• Local financial statements (latest year end) and TB for HO consolidation\n"
        "• Local accounting policies or closing manual (fee recognition, leases, provisions)\n"
        "• Copies of lease deeds and lease calculation template (ROU asset / lease liability)\n"
        "• List of COMMISSION accounts in the CoA",
    "7. 회신 기한 및 회신처": "7. Deadline and reply to",
    "회신 기한: 2026년 11 월  30 일(월)까지\n회신처: 딜로이트 안진회계법인 농협은행 감사팀 (담당: 신대환 / 연락처: 010-3440-2538)":
        "Deadline: Monday, 30 November 2026\n"
        "Reply to: Deloitte Anjin LLC, NongHyup Bank audit team (contact: Daehwan Shin / 010-3440-2538)",
    "범례": "Legend",
    "입력 셀(지점·법인 작성)": "Input cell (branch/subsidiary to complete)",
    "뉴욕지점 점검 결과 기준 사전 기재(확인 후 수정)": "Pre-filled from the NY review (confirm and correct)",
    "구분/항목 영역(수정 불요)": "Category/item area (do not change)",
    # Summary
    "총괄표 I 지점·법인별 요약 응답": "Summary | Answers by branch / subsidiary",
    "항목별 응답을 기재하고, 계정·계약 단위 상세는 각 상세 시트에 작성해 주시기 바랍니다.":
        "Enter an answer per item; give account- or contract-level detail in each detail sheet.",
    "구분": "Category", "확인 항목": "Item", "작성 방법 / 선택지": "How to answer / options",
    "관련 상세시트": "Detail sheet",
    "뉴욕지점": "New York Branch", "북경지점": "Beijing Branch", "하노이지점": "Hanoi Branch",
    "노이다지점": "Noida Branch", "런던지점": "London Branch", "홍콩지점": "Hong Kong Branch",
    "시드니지점": "Sydney Branch", "미얀마 법인": "Myanmar subsidiary", "캄보디아 법인": "Cambodia subsidiary",
    "재무제표 작성기준": "Reporting framework",
    "현지 재무제표 작성 기준": "Framework of local financial statements",
    "IFRS / 현지 GAAP(기준명 기재) / 기타": "IFRS / Local GAAP (name it) / Other",
    "현지 GAAP(US GAAP)": "Local GAAP (US GAAP)",
    "현지 GAAP": "Local GAAP",
    "현지 GAAP인 경우 적용 기준명": "If local GAAP, name of the framework",
    "예: US GAAP, PRC GAAP, VAS, Ind AS, UK GAAP(FRS 102), HKFRS, AASB 등":
        "e.g. US GAAP, PRC GAAP, VAS, Ind AS, UK GAAP (FRS 102), HKFRS, AASB",
    "본점 합산 시 IFRS conversion 수행 여부": "IFRS conversion performed for HO consolidation?",
    "수행 / 미수행 / 일부 수행": "Performed / Not performed / Partially performed",
    "2_IFRS조정": "2_IFRS_Adjustments",
    "미수행": "Not performed",
    "일부수행": "Partially performed",
    "IFRS conversion 항목(수행 중이거나 필요하다고 판단되는 항목)": "IFRS conversion items (performed or considered necessary)",
    "항목명 나열, 상세는 2_IFRS조정 시트 기재": "List items; detail in 2_IFRS_Adjustments",
    "현재 conversion 없음. 점검 시 식별된 차이 가능 항목: 이연대출부대수익 상각방법, 대출채권 조건변경, 복구충당부채":
        "No conversion now. Possible differences found in review: amortisation of deferred loan fees, "
        "loan modifications, restoration provision",
    "대손 충당금 ": "Loan-loss allowance",
    "리스 및 대손충당금": "Leases and loan-loss allowance",
    "금융상품 분류": "Classification of financial instruments",
    "대출채권 SPPI 테스트 수행 여부": "SPPI test performed for loans?",
    "수행 / 미수행 / 해당없음": "Performed / Not performed / N/A",
    "3_금융상품분류": "3_Fin_Instruments",
    "대출채권 사업모형 테스트 수행 여부": "Business-model test performed for loans?",
    "유가증권 SPPI 테스트 수행 여부": "SPPI test performed for securities?",
    "해당없음(유가증권 미보유)": "N/A (no securities held)",
    "유가증권 사업모형 테스트 수행 여부": "Business-model test performed for securities?",
    "SPPI 미충족 가능 상품 보유 여부(후순위 대출·트랜치, 옵션부 대출, 수익연계 조건 등)":
        "Any products that may fail SPPI (subordinated loans/tranches, loans with options, profit-linked terms, etc.)?",
    "있음(상품 기재) / 없음": "Yes (name products) / No",
    "없음(후순위·옵션부 대출 미보유)": "No (no subordinated or option-embedded loans)",
    "수수료 인식": "Fee recognition",
    "CoA 상 COMMISSION 계정 유형": "Types of COMMISSION accounts in the CoA",
    "계정 유형 나열, 계정별 상세는 4-7_COMMISSION 시트 기재": "List types; detail per account in 4-7_COMMISSION",
    "신디케이트론 참여수수료(Commitment Fee) 등 – 전체 계정 목록 기재 요망":
        "Syndicated-loan participation fee (commitment fee) etc. — full account list to be provided",
    "COMMISSION 항목 중 이연/일시인식 분류 현황": "COMMISSION items: deferred vs. upfront",
    "이연 계정 수 / 일시인식 계정 수, 분류 기준 기재": "Number deferred / number upfront, and the basis",
    "Commitment Fee: 이연 / 그 외 계정 기재 요망": "Commitment fee: deferred / other accounts to be provided",
    "이연항목 상각액의 인식 계정": "Account where amortisation of deferred items is recognised",
    "수수료수익 / 이자수익 / 혼재": "Fee income / Interest income / Mixed",
    "혼재(현지 재무제표: 수수료수익, 본점 합산 TB: 이자수익)":
        "Mixed (local FS: fee income; HO consolidation TB: interest income)",
    "이연항목 상각 방법": "Amortisation method for deferred items",
    "정액법 / 유효이자율법 / 혼재": "Straight-line / EIR method / Mixed",
    "정액법": "Straight-line",
    "리스": "Leases",
    "임대차계약서 상 리스기간 기산 시점": "Start of lease term per lease deed",
    "Rent free 기간 포함(Commencement Date 기산) / Rent free 이후(Rent Commencement Date 기산) / Rent free 없음":
        "Incl. rent-free (from Commencement Date) / After rent-free (from Rent Commencement Date) / No rent-free",
    "8-12_리스·복구": "8-12_Lease_Restoration",
    "Rent free 포함": "Incl. rent-free",
    "Rent free 기간 및 계약서 상 주요 일자": "Rent-free period and key dates in the deed",
    "Commencement / Rent Commencement / Expiration 일자 기재": "Enter Commencement / Rent Commencement / Expiration dates",
    "Commencement 2023.9.1 / Rent Commencement 2024.9.1(12개월 rent free) / Expiration 2034.8.31":
        "Commencement 2023-09-01 / Rent Commencement 2024-09-01 (12 months rent-free) / Expiration 2034-08-31",
    "리스 템플릿 상 계약기간과 계약서 일치 여부": "Lease term in template agrees with deed?",
    "일치 / 불일치(내용 기재) / 해당없음": "Match / Mismatch (describe) / N/A",
    "불일치 → 시정 완료(템플릿 만료일 2033년 → 2034.8.31 반영)":
        "Mismatch → corrected (template expiry 2033 → 2034-08-31)",
    "임대료 조정시점의 계약서 상 조정시점과 매칭 여부": "Rent escalation dates agree with deed?",
    "일치 / 불일치(내용 기재) / 조정조항 없음": "Match / Mismatch (describe) / No escalation clause",
    "불일치 → 시정 완료(조정시점 2028년 → 2029년 반영)": "Mismatch → corrected (escalation 2028 → 2029)",
    "복구충당부채": "Restoration provision",
    "계약서 상 원상복구의무 존재 여부": "Restoration (reinstatement) obligation in the deed?",
    "있음(조항번호 기재) / 없음": "Yes (clause no.) / No",
    "있음(Tower 45 임대차계약서 10.01)": "Yes (Tower 45 lease, clause 10.01)",
    "복구충당부채 계상 여부": "Restoration provision recognised?",
    "계상 / 미계상": "Recognised / Not recognised",
    "미계상": "Not recognised",
    "미계상 시 사유": "If not recognised, why",
    "예: 금액 중요성 낮음, 현지 GAAP 상 미인식, 추정 불가 등":
        "e.g. immaterial, not recognised under local GAAP, cannot be estimated",
    "금액 중요성 낮음": "Immaterial",
    "기타": "Other",
    "그 밖에 IFRS와 차이가 있거나 확인이 필요한 사항": "Other IFRS differences or items to confirm",
    "주요 내용 요약, 상세는 13_기타 시트 기재": "Summarise; detail in 13_Others",
    "13_기타": "13_Others",
    "작성 완료 항목 수": "Items completed",
    "작성률": "Completion rate",
    "※ 뉴욕지점 '확인 필요', '기재 요망' 표시 항목은 지점 확인 후 내용을 확정해 주시기 바랍니다.":
        "※ New York: items marked 'To be confirmed' or 'To be provided' are to be finalised after branch confirmation.",
    # 2_IFRS_Adjustments
    "항목 2 I IFRS conversion 항목 상세": "Item 2 | IFRS conversion items in detail",
    "현지 GAAP과 IFRS 간 회계처리 차이가 있는 항목을 기재해 주시기 바랍니다(현재 조정 수행 중인 항목 및 조정이 필요하다고 판단되는 항목 모두 포함).":
        "List items where local GAAP and IFRS treatment differ (both items already adjusted and items considered to need adjustment).",
    "지점·법인": "Branch / subsidiary", "관련 계정": "Related account",
    "현지 GAAP 회계처리": "Local GAAP treatment", "IFRS 회계처리": "IFRS treatment",
    "현재 conversion 수행 여부": "Conversion currently performed?", "조정 수행 주체": "Adjusted by",
    "조정 금액(현지통화)": "Adjustment amount (local currency)", "기준일": "As of", "비고": "Remarks",
    "이연대출부대수익(Commitment Fee)": "Deferred loan fees (commitment fee)",
    "이연 후 정액법 상각": "Deferred and amortised straight-line",
    "유효이자율법에 따라 유효수익 조정(IFRS 9 B5.4.2)": "Adjust effective yield under the EIR method (IFRS 9 B5.4.2)",
    "미정": "TBD",
    "양적·질적 중요성에 따른 적용 여부 결정 필요": "Decide based on quantitative and qualitative materiality",
    "대출채권(조건변경)": "Loans (modification)",
    "유효이자율 재산정 후 향후 이자수익 반영(ASC 310-20)":
        "Recalculate EIR and reflect in future interest income (ASC 310-20)",
    "최초 유효이자율로 할인하여 변경손익 즉시 인식(IFRS 9 5.4.3)":
        "Discount at the original EIR and recognise modification gain/loss immediately (IFRS 9 5.4.3)",
    "제거요건 미충족 조건변경 발생 시 중요성에 따라 검토":
        "Review by materiality when non-derecognition modifications occur",
    "원상복구의무에 대한 충당부채 인식(IAS 37), 사용권자산 원가 가산":
        "Recognise provision for restoration (IAS 37), added to ROU asset cost",
    "지점": "Branch",
    # 3_Fin_Instruments
    "항목 3 I 보유 금융상품 분류 및 SPPI·사업모형 테스트":
        "Item 3 | Classification of financial instruments held and SPPI / business-model tests",
    "보유 금융상품 유형별로 현지 분류와 IFRS 9 분류, 테스트 수행 여부를 기재해 주시기 바랍니다.":
        "For each type of instrument held, enter the local classification, IFRS 9 classification and whether tests were performed.",
    "상품 유형": "Type", "세부 상품명": "Product", "현지 GAAP 분류": "Local GAAP classification",
    "SPPI 테스트 수행": "SPPI test performed", "SPPI 충족 여부": "SPPI met?",
    "사업모형 테스트 수행": "Business-model test performed", "사업모형": "Business model",
    "IFRS 9 분류": "IFRS 9 classification", "기말 잔액(현지통화)": "Year-end balance (local currency)",
    "비SPPI 가능 조건(후순위, 옵션, 수익연계 등)": "Possible non-SPPI terms (subordination, options, profit-linked)",
    "대출채권": "Loans",
    "CB·IB 대출(신디케이트론 포함)": "CB/IB loans (incl. syndicated loans)",
    "상각후원가": "Amortised cost",
    "없음(점검 시 후순위·옵션부 대출 미확인)": "None (no subordinated or option-embedded loans found in review)",
    "향후 IB 대출 중 후순위 트랜치 등 발생 시 모니터링 권고":
        "Monitor if subordinated tranches etc. arise in future IB loans",
    "유가증권": "Securities",
    "해당없음": "N/A",
    "유가증권 미보유": "No securities held",
    # 4-7_COMMISSION
    "항목 4~7 │ COMMISSION 계정별 인식 현황": "Items 4–7 │ Recognition by COMMISSION account",
    "CoA 상 COMMISSION(수수료) 관련 계정을 모두 기재하고, 계정별 이연/일시인식 여부, 상각액 인식 계정, 상각 방법을 기재해 주시기 바랍니다.":
        "List all COMMISSION (fee) accounts in the CoA, with deferral/upfront recognition, the account for amortisation "
        "and the amortisation method for each.",
    "계정코드": "Account code", "계정명(CoA)": "Account name (CoA)", "수수료 성격(수취 사유)": "Nature of fee",
    "관련 상품(대출 유형 등)": "Related product (loan type etc.)", "이연/일시인식": "Deferred / upfront",
    "이연 시 상각액 인식 계정(현지)": "Account for amortisation if deferred (local)",
    "본점 합산 TB 상 계정": "Account in HO consolidation TB", "상각 방법": "Amortisation method",
    "상각 기간": "Amortisation period", "당기 인식액(현지통화)": "Recognised this year (local currency)",
    "기말 이연잔액(현지통화)": "Deferred balance at year end (local currency)",
    "Commitment Fee(신디케이트론 참여수수료)": "Commitment fee (syndicated-loan participation fee)",
    "신디케이트론 참여에 따른 약정수수료(주선용역 없음)": "Commitment fee for joining a syndicated loan (no arranging service)",
    "신디케이트론": "Syndicated loans",
    "이연": "Deferred",
    "수수료수익": "Fee income",
    "이자수익": "Interest income",
    "특정 대출 발생가능성이 거의 확실하여 유효이자율법 적용 필요(회계업무지침서 2.1.2, IFRS 9 B5.4.2)":
        "Drawdown of the specific loan is virtually certain, so the EIR method applies (accounting manual 2.1.2, IFRS 9 B5.4.2)",
    # 8-12
    "항목 8~12 I 임대차계약별 리스기간·임대료 조정·복구의무":
        "Items 8–12 | Lease term, rent escalation and restoration obligation by lease contract",
    "지점·법인이 체결한 임대차계약(사무실, 사택, 주차장 등)별로 기재해 주시기 바랍니다. 계약서 기재 내용과 리스 템플릿 반영 내용을 나란히 기재하여 비교합니다.":
        "Enter one row per lease (office, staff residence, parking, etc.). Show the deed terms next to the lease template "
        "to compare them.",
    "임차 대상(주소/건물명)": "Leased premises (address / building)", "용도": "Use",
    "계약서 Commencement Date": "Deed Commencement Date", "Rent free 기간(개월)": "Rent-free period (months)",
    "계약서 Rent Commencement Date": "Deed Rent Commencement Date", "계약서 Expiration Date": "Deed Expiration Date",
    "리스기간 기산 기준": "Lease term starts from", "템플릿 상 리스 개시일": "Lease start in template",
    "템플릿 상 리스 만료일": "Lease end in template", "계약기간 일치 여부": "Term matches?",
    "계약서 상 임대료 조정 시점(조항)": "Rent escalation per deed (clause)",
    "템플릿 상 임대료 조정 시점": "Rent escalation in template", "조정시점 일치 여부": "Escalation matches?",
    "연장·해지 옵션 및 반영 여부": "Extension/termination options and treatment",
    "원상복구의무(조항)": "Restoration obligation (clause)",
    "복구충당부채 금액(현지통화)": "Restoration provision (local currency)", "미계상 사유": "Reason not recognised",
    "시정조치 내역 및 비고": "Corrective action and remarks",
    "사무실": "Office",
    "2033년(수정 전)": "2033 (before correction)",
    "불일치": "Mismatch",
    "Rent Commencement Date 기준 5년(2029년)까지 USD 47,798.33, 이후 USD 51,212.50(조항 4.01)":
        "USD 47,798.33 for 5 years from Rent Commencement Date (to 2029), then USD 51,212.50 (clause 4.01)",
    "2028년(수정 전)": "2028 (before correction)",
    "있음(조항 10.01)": "Yes (clause 10.01)",
    "기재 요망": "To be provided",
    "리스 만료일 2034.8.31 및 임대료 조정시점 2029년으로 템플릿 수정 완료":
        "Template corrected to expiry 2034-08-31 and escalation in 2029",
    # 13
    "항목 13 I 기타 확인 사항": "Item 13 | Other items",
    "그 밖에 현지 GAAP과 IFRS 간 차이가 있거나, 본점 합산 과정에서 확인이 필요하다고 판단되는 사항을 기재해 주시기 바랍니다(예: 손실충당금, 대출채권 조건변경, 무역금융·지급보증, 이연법인세, 퇴직급여, 외화환산 등).":
        "Enter any other local GAAP vs. IFRS differences or items to confirm for HO consolidation (e.g. loss allowance, "
        "loan modifications, trade finance / guarantees, deferred tax, employee benefits, FX translation).",
    "내용": "Description", "현지 회계처리": "Local treatment", "IFRS 상 예상 처리": "Expected IFRS treatment",
    "금액(현지통화)": "Amount (local currency)", "처리 방안(안)": "Proposed handling",
    # NY issues
    "뉴욕지점 점검 이슈별 해당 여부 및 처리 주체 의견": "NY review issues — applicability and who handles them",
    "뉴욕지점 점검 이슈가 지점·법인에도 해당되는지 표시하고(O/X/확인 중), 처리 주체에 대한 글로벌사업부 의견을 기재해 주시기 바랍니다.":
        "Mark whether each NY issue also applies (O / X / Checking), and the Global Business Department's view on who handles it.",
    "뉴욕지점 점검 이슈": "NY review issue", "뉴욕지점 조치 현황": "NY action status",
    "처리 주체(안)": "Handled by (proposed)", "처리 방안 및 일정": "Approach and timing",
    "해당 지점·법인 수": "No. of branches/subsidiaries affected",
    "GAAP 조정(일반)": "GAAP adjustment (general)",
    "현지 GAAP 기준 작성 후 본점 합산 시 별도 IFRS conversion 미수행":
        "Prepared under local GAAP; no separate IFRS conversion on HO consolidation",
    "중요 GAAP 차이 식별·분석 필요": "Identify and analyse material GAAP differences",
    "혼합(항목별 구분)": "Mixed (by item)",
    "대출채권 분류": "Loan classification",
    "SPPI·사업모형 기준 분류 판단 체계 부재(현재 중요 차이는 없을 것으로 예상)":
        "No SPPI / business-model classification framework (no material difference expected now)",
    "지점 자체 모니터링 권고": "Branch monitoring recommended",
    "대출채권 조건변경": "Loan modifications",
    "조건변경 시 IFRS(변경손익 즉시 인식)와 US GAAP(유효이자율 재산정) 처리 차이":
        "On modification, IFRS (gain/loss immediately) vs. US GAAP (recalculate EIR)",
    "발생 시 중요성에 따라 conversion 검토": "Consider conversion by materiality when it occurs",
    "이연대출부대수익 상각방법": "Amortisation of deferred loan fees",
    "신디케이트론 참여수수료를 정액법으로 상각(유효이자율법 적용 필요)":
        "Syndicated-loan participation fees amortised straight-line (EIR method needed)",
    "중요성에 따른 적용 여부 결정 필요": "Decide based on materiality",
    "수수료 계정 분류": "Fee account classification",
    "현지 재무제표(수수료수익)와 본점 합산 TB(이자수익) 간 분류 차이":
        "Classification differs between local FS (fee income) and HO consolidation TB (interest income)",
    "본점 합산 TB 상 이자수익 표시 중": "Shown as interest income in HO consolidation TB",
    "리스기간": "Lease term",
    "리스 템플릿 만료일(2033년)과 계약서 만료일(2034.8.31) 불일치":
        "Template expiry (2033) differs from deed expiry (2034-08-31)",
    "시정 완료": "Corrected",
    "임대료 조정시점": "Rent escalation date",
    "임대료 조정시점(템플릿 2028년, 계약서 2029년) 불일치": "Escalation date differs (template 2028, deed 2029)",
    "계약서 상 원상복구의무 존재하나 복구충당부채 미계상":
        "Restoration obligation in the deed but no restoration provision recognised",
    # dropdown lists
    '"IFRS,현지 GAAP,기타"': '"IFRS,Local GAAP,Other"',
    '"수행,미수행,일부 수행"': '"Performed,Not performed,Partially performed"',
    '"수행,미수행,해당없음"': '"Performed,Not performed,N/A"',
    '"수수료수익,이자수익,혼재,해당없음"': '"Fee income,Interest income,Mixed,N/A"',
    '"정액법,유효이자율법,혼재,해당없음"': '"Straight-line,EIR method,Mixed,N/A"',
    '"Rent free 포함,Rent free 이후,Rent free 없음,해당없음"': '"Incl. rent-free,After rent-free,No rent-free,N/A"',
    '"일치,불일치,해당없음"': '"Match,Mismatch,N/A"',
    '"일치,불일치,조정조항 없음,해당없음"': '"Match,Mismatch,No escalation clause,N/A"',
    '"있음,없음,해당없음"': '"Yes,No,N/A"',
    '"계상,미계상,해당없음"': '"Recognised,Not recognised,N/A"',
    '"뉴욕지점,북경지점,하노이지점,노이다지점,런던지점,양곤지점,홍콩지점,시드니지점,미얀마 법인,캄보디아 법인"':
        '"New York Branch,Beijing Branch,Hanoi Branch,Noida Branch,London Branch,Yangon Branch,Hong Kong Branch,'
        'Sydney Branch,Myanmar subsidiary,Cambodia subsidiary"',
    '"지점,글로벌사업부,미정"': '"Branch,Global Business Dept,TBD"',
    '"대출채권,유가증권,예치금,매입외환,기타"': '"Loans,Securities,Placements,Bills purchased,Other"',
    '"충족,미충족,해당없음"': '"Met,Not met,N/A"',
    '"보유,보유 및 매도,기타,해당없음"': '"Hold to collect,Hold to collect & sell,Other,N/A"',
    '"AC,FVOCI,FVPL,해당없음"': '"AC,FVOCI,FVPL,N/A"',
    '"이연,일시인식"': '"Deferred,Upfront"',
    '"수수료수익,이자수익,해당없음"': '"Fee income,Interest income,N/A"',
    '"정액법,유효이자율법,해당없음"': '"Straight-line,EIR method,N/A"',
    '"Rent free 포함,Rent free 이후,Rent free 없음"': '"Incl. rent-free,After rent-free,No rent-free"',
    '"일치,불일치"': '"Match,Mismatch"',
    '"일치,불일치,조정조항 없음"': '"Match,Mismatch,No escalation clause"',
    '"각 지점·법인 처리,글로벌사업부 주관 conversion,혼합(항목별 구분),미정"':
        '"Each branch/subsidiary,Conversion led by Global Business Dept,Mixed (by item),TBD"',
    '"O,X,확인 중"': '"O,X,Checking"',
}
