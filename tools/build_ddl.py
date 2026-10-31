# -*- coding: utf-8 -*-
"""Oracle DDL 생성: python3 tools/build_ddl.py db/ddl_oracle.sql"""
import sys
import model as M

FK_TARGET = {  # 컬럼명 → 참조 테이블(PK 단일컬럼)
    "CUST_NO": "TB_CUST_FX_PROFILE", "INVESTEE_CUST_NO": "TB_CUST_FX_PROFILE", "GUARDIAN_CUST_NO": "TB_CUST_FX_PROFILE",
    "REMIT_REF_NO": "TB_FX_REMIT_MST", "OUT_REMIT_REF_NO": "TB_FX_REMIT_MST", "SRC_REMIT_REF_NO": "TB_FX_REMIT_MST",
    "INV_MGMT_NO": "TB_TRD_INVOICE", "BOE_MGMT_NO": "TB_TRD_BOE", "SB_MGMT_NO": "TB_TRD_SHIPPING_BILL",
    "IRM_NO": "TB_TRD_IRM", "FDI_REF_NO": "TB_FDI_INFLOW", "VOSTRO_ACCT_NO": "TB_VOSTRO_ACCT",
    "LODGE_REF_NO": "TB_TRD_EXP_LODGE", "RETURN_ID": "TB_CIMS_RETURN", "STMT_ID": "TB_VOSTRO_STMT",
    "R_PARTY_CUST_NO": "TB_CUST_FX_PROFILE", "PAYOUT_ID": "TB_DOM_PAYOUT", "ESCROW_ACCT_NO": "TB_FDI_ESCROW",
    "MSG_TYPE": "TB_SWIFT_MSG_TYPE", "SWIFT_MSG_ID": "TB_SWIFT_MSG", "PAY_MSG_ID": "TB_SWIFT_MSG",
    "STMT_MSG_ID": "TB_SWIFT_MSG", "SETTLE_MSG_ID": "TB_SWIFT_MSG", "RA_NO": "TB_RA_AUTH",
    "IO_REF": "TB_IO_TXN", "BATCH_ID": "TB_IO_SETTLE_BATCH",
    "PREF_DEAL_REF": "TB_MM_DEAL_REF", "MT320_MSG_ID": "TB_SWIFT_MSG",
    "REQ_MSG_ID": "TB_SWIFT_MSG", "RESP_MSG_ID": "TB_SWIFT_MSG", "KYC_DOC_ID": "TB_FX_DOC",
}


def main(out):
    L = [f"-- {M.META['project']} / {M.META['version']} / generated from tools/model.py - DO NOT EDIT BY HAND", ""]
    fks = []
    for t in M.T:
        L.append(f"-- {t['kor']} : {t['desc']}")
        L.append(f"CREATE TABLE {t['name']} (")
        defs = []
        for c in t["cols"]:
            nn = " NOT NULL" if ("NN" in c[3] or "PK" in c[3]) else ""
            defs.append(f"  {c[0]:<24} {c[2]}{nn}")
        pk = [c[0] for c in t["cols"] if "PK" in c[3]]
        defs.append(f"  CONSTRAINT PK_{t['name'][3:]} PRIMARY KEY ({', '.join(pk)})")
        L.append(",\n".join(defs))
        L.append(");")
        L.append(f"COMMENT ON TABLE {t['name']} IS '{t['kor']}';")
        for c in t["cols"]:
            cm = (c[1] + (" - " + c[4] if c[4] else "")).replace("'", "''")
            L.append(f"COMMENT ON COLUMN {t['name']}.{c[0]} IS '{cm}';")
        for i, ix in enumerate(t["idx"], 1):
            L.append(f"CREATE INDEX IX_{t['name'][3:]}_{i} ON {t['name']} ({ix});")
        for c in t["cols"]:
            tgt = FK_TARGET.get(c[0])
            if "FK" in c[3] and tgt and tgt != t["name"]:
                tgt_pk = [x[0] for x in next(x for x in M.T if x["name"] == tgt)["cols"] if "PK" in x[3]]
                if len(tgt_pk) == 1:
                    fks.append(f"ALTER TABLE {t['name']} ADD CONSTRAINT FK_{t['name'][3:]}_{c[0]} "
                               f"FOREIGN KEY ({c[0]}) REFERENCES {tgt} ({tgt_pk[0]});")
        L.append("")
    L.append("-- Foreign keys")
    L += fks
    L.append("")
    L.append("-- 규정 파라미터 초기적재")
    for p in M.PARAMS:
        v = p[2].replace("'", "''")
        L.append(f"INSERT INTO TB_FX_PARAM (PARAM_ID, EFF_FROM_DT, EFF_TO_DT, PARAM_NAME, PARAM_VALUE, UNIT, REG_SOURCE, REG_USER_ID, REG_TS) "
                 f"VALUES ('{p[0]}', DATE '2000-01-01', DATE '9999-12-31', '{p[1]}', '{v}', '{p[3]}', '{p[5]}', 'SYSTEM', SYSTIMESTAMP);")
    L.append("")
    L.append("-- 보고서 시스템 연동용 뷰 예시: 기한 미이행 현황")
    L.append("""CREATE OR REPLACE VIEW V_RPT_OBLIGATION_OPEN AS
SELECT o.OBLIG_TYPE, o.CUST_NO, c.IEC_NO, c.PAN_NO, o.SRC_TABLE, o.SRC_KEY, o.BASE_DT,
       NVL(o.EXT_DUE_DT, o.DUE_DT) AS EFFECTIVE_DUE_DT,
       o.TARGET_AMT, o.FULFILLED_AMT, o.TARGET_AMT - NVL(o.FULFILLED_AMT,0) AS PENDING_AMT,
       TRUNC(SYSDATE) - NVL(o.EXT_DUE_DT, o.DUE_DT) AS OVERDUE_DAYS, o.OBLIG_STS
  FROM TB_FX_OBLIGATION o JOIN TB_CUST_FX_PROFILE c ON c.CUST_NO = o.CUST_NO
 WHERE o.OBLIG_STS NOT IN ('FULFILLED','WAIVED');""")
    L.append("")
    L.append("""CREATE OR REPLACE VIEW V_RPT_LRS_DAILY AS
SELECT l.FY, l.PAN_NO, m.TRX_DT, m.REMIT_REF_NO, m.PURPOSE_CODE, l.LRS_CAT, m.CCY, m.FCY_AMT,
       l.USD_EQV_AMT, m.INR_AMT, m.REMIT_CTRY, l.TCS_AMT, l.CIMS_RPT_STS
  FROM TB_LRS_TXN l JOIN TB_FX_REMIT_MST m ON m.REMIT_REF_NO = l.REMIT_REF_NO
 WHERE m.STATUS = '40';""")
    open(out, "w", encoding="utf-8").write("\n".join(L) + "\n")
    print("saved", out, len(M.T), "tables", len(fks), "fks")


if __name__ == "__main__":
    main(sys.argv[1])
