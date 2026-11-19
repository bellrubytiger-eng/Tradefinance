# -*- coding: utf-8 -*-
"""본점 협의 설명자료용 데이터: python3 tools/ho_brief_data.py out.json
06 정정본(docs/06_..._정정본.xlsx) 집계 + coa_model 결과."""
import json
import os
import sys
from collections import Counter, defaultdict
from openpyxl import load_workbook
import coa_model as C

HERE = os.path.dirname(os.path.abspath(__file__))
FIXED = os.path.join(HERE, "..", "docs", "06_통합표준CoA_0930_인도반영_정정본.xlsx")


def main(out):
    wb = load_workbook(FIXED, read_only=True, data_only=True)
    ws = wb["통합표준CoA_계정상태표"]
    rows = [r for r in ws.iter_rows(min_row=4, values_only=True) if r[2]]
    acct = {str(r[2]): {"name": (r[12] or (r[4] or "").strip()), "en": (r[4] or "").strip(), "status": r[10] or "", "post": r[11] or "", "area": r[1]} for r in rows}
    summary = defaultdict(Counter)
    for r in rows:
        summary[r[0]][r[10] or "유지"] += 1
    new_area = Counter(r[1] for r in rows if r[10] == "신규")
    log = [r for r in wb["정정이력"].iter_rows(min_row=4, values_only=True) if r[0]]
    terms = [r[:4] for r in load_workbook(os.path.join(HERE, "data", "06_CoA_0930_india_original.xlsx"), read_only=True)["인도용어정의"].iter_rows(min_row=4, values_only=True) if r[1]]
    fut = defaultdict(lambda: [0, "", "", ""])
    for u in C.FUTURE:
        k = (u["future"], u["when"])
        fut[k][0] += 1; fut[k][1] = u["future"]; fut[k][2] = u["when"]; fut[k][3] = u["owner"]
    d = {
        "acct": acct,
        "summary": {g: dict(c) for g, c in summary.items()},
        "total": len(rows),
        "new_area": dict(new_area.most_common()),
        "decisions": [list(x) for x in C.DECISIONS],
        "review": [list(x) for x in C.REVIEW],
        "future": list(fut.values()),
        "cur": C.CUR_STATS, "cur_mapped": len(C.FINAL_MAP), "cur_excluded": len(C.EXCLUDED),
        "splits": sum(1 for v in C.FINAL_MAP.values() if len(v) > 1),
        "log_n": len(log), "log_by_col": dict(Counter(r[2] for r in log)),
        "terms": [list(t) for t in terms],
        "ifrs_only": sum(1 for a in acct.values() if "IFRS" in a["post"]),
        "bal_rows": sum(1 for a in C.BASE if a["bal"] not in (None, "")),
    }
    json.dump(d, open(out, "w", encoding="utf-8"), ensure_ascii=False, default=str)
    print("saved", out, d["total"], "accounts")


if __name__ == "__main__":
    main(sys.argv[1])
