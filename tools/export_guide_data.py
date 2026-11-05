# -*- coding: utf-8 -*-
"""본점용 워드 가이드(build_guide_docx.js)에 넣을 데이터 내보내기."""
import json
import sys

import coa_review_model as RM
import gaap_model as GA
import glossary_model as GL
import lease_model as LM
import tax_fa_model as TF

out = {
    "gaap_diff": GA.GAAP_DIFF,
    "ifrs_only": GA.IFRS_ONLY,
    "schedules": RM.SCHEDULES,
    "tax_param": TF.TAX_GL_PARAM,
    "ye_tds": TF.YE_TDS,
    "fee_matrix": TF.FEE_MATRIX,
    "fa_class": TF.FA_CLASS,
    "lease_je": LM.LEASE_JE,
    "glossary": [(t, c, d, b) for t, c, d, b, _p in GL.G],
    "tds_l6": [(n[0], n[3], n[4], n[5]) for n in TF.NEW3 if n[0].startswith("2051040") and n[1] == 6],
}
json.dump(out, open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("exported", sys.argv[1])
