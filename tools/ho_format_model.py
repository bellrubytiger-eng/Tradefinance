# -*- coding: utf-8 -*-
"""본점 양식(통합표준CoA_계정상태표, 0930) 그대로 + 인도지점 반영 열.

본점 원본 10열(구분·업무영역·코드·Level·계정과목·원화계정명·본점 통합계정코드·본점 통합계정과목명·비화폐성·GBS코드)을
그대로 두고, 오른쪽에 인도지점 검토 결과를 붙인다. 신설 계정은 코드 순서 위치에 끼워 넣는다.
"""
import re

import coa_model as CM
import coa_review_model as RM
import coa_migration_model as MG
import tax_fa_model as TF
import glossary_model as GL
import gaap_model as GA

BASE_BY = {a["code"]: a for a in CM.BASE}

# 표시 순서가 정해진 상태
ORDER = ["신규", "변경", "용도지정", "비활성", "철회", "유지"]


def _sibling_ho(code, lv):
    """같은 상위 아래 같은 Level 기존 계정의 본점 매핑(가장 흔한 값)."""
    k = 1 + 2 * (lv - 2)
    sib = [a for a in CM.BASE if a["lv"] == lv and a["code"][:k] == code[:k] and a["ho_name"]]
    if not sib:
        return None
    from collections import Counter
    name = Counter(a["ho_name"] for a in sib).most_common(1)[0][0]
    return next(a for a in sib if a["ho_name"] == name)


def _ancestor(code):
    for k in (9, 7, 5, 3, 1):
        c = code[:k].ljust(11, "0")
        if c != code and c in BASE_BY:
            return BASE_BY[c]
    return None


def sch_line(code):
    """인도 재무제표(Third Schedule) 줄 — 참고용 접두 매핑."""
    if code.startswith(("509", "5101001", "5101002", "5101003", "51204", "51211")):
        ln = TF.sch16_of(code)
        return f"Sch.16 {ln}" if ln else "Sch.16"
    P = [("10101", "Sch.6 I 현금"), ("10102", "Sch.6 I 현금"), ("10103", "Sch.6 I 현금"), ("10104", "Sch.6 II RBI"),
         ("10105", "Sch.7 은행 당좌(인도 내/외)"), ("10106", "Sch.7 은행 기타예금"), ("101070", "Sch.7 Money at call"), ("10107", "Sch.7"),
         ("10108", "Sch.7(충당금 차감)"), ("102", "Sch.8 투자"), ("103", "Sch.9 대출"), ("104", "Sch.11 기타자산"),
         ("105", "Sch.10 고정자산"), ("106", "Sch.10 고정자산"), ("107", "해당없음(인도 AS 19)"), ("108", "Sch.11 선납세금·이연법인세"),
         ("109", "Sch.11 / 파생"), ("110", "Sch.11 본지점(순액)"), ("1120", "포지션(B/S 미표시)"),
         ("2011", "Sch.5 Bills payable 등"), ("2012", "Sch.5"), ("2013", "Sch.4 차입(당좌차월)"), ("2014", "Sch.4 차입(당좌차월)"),
         ("2015", "Sch.3 요구불·저축"), ("2016", "Sch.3 III 정기"), ("2017", "Sch.3 III 정기(CD)"), ("2018", "Sch.5 기타"),
         ("2019", "Sch.3 III 정기(적금)"), ("202", "Sch.4 차입 / Sch.5"), ("203", "해당없음"), ("204", "Sch.5 기타부채"),
         ("205", "Sch.5 기타부채"), ("206", "해당없음(인도 AS 19)"), ("207", "Sch.5"), ("208", "Sch.5 / 파생"), ("209", "Sch.5 충당부채"),
         ("210", "Sch.5 본지점(순액)"), ("211", "Sch.5"), ("2120", "Sch.5"), ("2130", "포지션"),
         ("3114", "Sch.1 자본"), ("311", "Sch.2 준비금·잉여금"),
         ("4011", "Sch.13 i"), ("4012", "Sch.13 iii"), ("4013", "Sch.13 iii"), ("402", "Sch.13 ii"), ("403", "Sch.13 i"),
         ("404", "충당금 환입(P&L)"), ("405", "Sch.14 i"), ("406", "Sch.13 iv"), ("4076", "Sch.14 v 외환"), ("4077", "Sch.14 iv 처분손익"),
         ("407", "Sch.14 vii 기타"), ("408", "Sch.14 vii 기타"), ("409", "Sch.14"), ("410", "Sch.14"), ("499", "손익"),
         ("501", "Sch.15 i 예금이자"), ("502", "Sch.15 ii 차입이자"), ("503", "Sch.15 iii"), ("504", "Sch.14 ii(순액)"),
         ("505", "Sch.16 XII / 수수료 차감"), ("506", "충당금(P&L)"), ("507", "Sch.14"), ("508", "Sch.14"), ("5101004", "Sch.14 v 외환(순액)"),
         ("5111", "Sch.15"), ("512", "Sch.16 XII / Sch.14"), ("513", "법인세"), ("599", "손익"),
         ("601", "Sch.12 우발채무"), ("602", "Sch.12 선물환"), ("603", "비망(대차 없음)")]
    best = ""
    for p, v in P:
        if code.startswith(p) and len(p) > len(best.split("|")[0] if best else ""):
            best = p + "|" + v
    return best.split("|")[1] if best else ""


def _live_index():
    idx = {}
    for m in MG.migrate():
        for i, t in enumerate(MG.targets_of(m["tgt"])):
            e = idx.setdefault(t, {"codes": [], "amt": 0.0})
            e["codes"].append(m["code"])
            if i == 0:
                e["amt"] += m["amt"] if m["side"] in ("ASSET", "EXP") else -m["amt"] if m["side"] in ("LIAB", "INC") else 0
    return idx


def rows():
    """본점 양식 행 목록 (dict)."""
    st = {}  # code -> dict(status, name_in, why, src)

    def put(code, status, name="", why="", src=""):
        cur = st.get(code)
        # 뒤에 온 판단이 우선(차수 순서대로 호출)
        st[code] = dict(status=status, name=name or (cur["name"] if cur else ""), why=why or (cur["why"] if cur else ""), src=src)

    # 1차 이전안(coa_model)
    for c, n, m in CM.MOD:
        put(c, "변경", n or "", m, "1차")
    for c, m in CM.DEACT:
        put(c, "비활성", "", m, "1차")
    for c, lv, en, kr, sib, memo in CM.ADD:
        put(c, "신규", en, memo, "1차")
    # 검토의견(1~8차)
    kmap = {"신규": "신규", "변경": "변경", "이전안수정": "변경", "매핑": "변경", "비활성": "비활성", "철회": "철회", "속성전환": "변경"}
    for r in RM.R:
        if r["code"] == "-":
            continue
        k = kmap[r["kind"]]
        if r["kind"] == "속성전환" and r["code"] in BASE_BY:
            k = "변경"
        if r["kind"] == "속성전환" and r["code"] not in BASE_BY:
            k = "철회"
        if r["kind"] == "매핑":
            k = "용도지정" if r["code"] in BASE_BY else "신규"
        put(r["code"], k, r["en"], r["why"], "검토의견")
    for n in MG.NEW2:
        put(n[0], "신규", n[3], n[5], "이행")
    for n in TF.NEW3:
        put(n[0], "신규", n[3], n[5], "세무·고정자산·비용")
    # 이전안 신규 중 철회된 것: 기존에 없는 코드가 비활성이면 철회
    for c, v in st.items():
        if c not in BASE_BY and v["status"] == "비활성":
            v["status"] = "철회"

    kr = {x[0]: x[3] for x in CM.ADD}
    kr.update({r["code"]: r["kr"] for r in RM.R if r["code"] != "-"})
    kr.update({n[0]: n[4] for n in MG.NEW2})
    kr.update({n[0]: n[4] for n in TF.NEW3})
    fix = {}
    for c, m in CM.FIX:
        fix.setdefault(c, []).append(m)
    for a in CM.BASE:
        p = parent_of(a["code"])
        if p and p not in BASE_BY:
            fix.setdefault(a["code"], []).append(
                f"코드 구조: 산식상 상위 {p}가 없음(실제 상위 L3는 20501000000 OTHER LIABILITIES). L3 코드를 20510000000으로 바꾸거나 하위를 2050xxxxxxx로 재채번 요청")
    live = _live_index()
    lvl = {}
    out = []
    for a in CM.BASE:
        lvl[a["code"]] = a["lv"]
    for c, v in st.items():
        if c not in BASE_BY:
            src_lv = next((n[1] for n in TF.NEW3 if n[0] == c), None) or next((n[1] for n in MG.NEW2 if n[0] == c), None) \
                or next((x[1] for x in CM.ADD if x[0] == c), None) or next((r["lv"] for r in RM.R if r["code"] == c and r["lv"]), 5)
            lvl[c] = int(src_lv)
    codes = sorted(set(BASE_BY) | {c for c in st if c not in BASE_BY})
    for c in codes:
        a = BASE_BY.get(c)
        v = st.get(c)
        status = v["status"] if v else "유지"
        if a is None and status == "철회":
            pass
        if a:
            grp, area, name, krw, ho, ho_name, nonmon, gbs = a["grp"], a["area"], a["name"].strip(), a["krw"], a["ho"], a["ho_name"], a["nonmon"], a["gbs"]
        else:
            anc = _ancestor(c) or {}
            grp, area = anc.get("grp", ""), anc.get("area", "")
            name, krw = v["name"], ""
            src_ho = anc if anc.get("ho_name") else (_sibling_ho(c, lvl[c]) or {})
            if not src_ho.get("ho_name") and c[:4] in ("2019", "5017"):  # 적금 = 정기예금과 같은 본점 매핑(결정 P3)
                tc = {"2019": "2016", "5017": "5012"}[c[:4]] + c[4:]
                twin = BASE_BY.get(tc) or _sibling_ho(tc, lvl[c]) or _ancestor(tc)
                src_ho = twin if twin and twin.get("ho_name") else {}
            ho, ho_name = src_ho.get("ho", ""), (src_ho.get("ho_name", "") + " (안)") if src_ho.get("ho_name") else "(본점 매핑 요청)"
            nonmon, gbs = anc.get("nonmon", ""), "채번요청"
        india_name = (v["name"] if v and a and v["name"] and v["name"].strip() != name else "")
        if a is None:
            india_name = kr.get(c, "")
        lv = lvl[c]
        e = live.get(c, {"codes": [], "amt": 0.0})
        why = v["why"] if v else ""
        text = " ".join([name, india_name, why])
        out.append(dict(
            grp=grp, area=area, code=c, lv=lv, name=("  " * (lv - 1)) + name, krw=krw, ho=ho, ho_name=ho_name, nonmon=nonmon, gbs=gbs,
            status=status, ledger=GA.ledger_use(c, status), india_name=india_name, live=", ".join(e["codes"][:6]) + (" 외" if len(e["codes"]) > 6 else ""),
            amt=round(e["amt"], 2) if e["amt"] else None, why=why, sch=sch_line(c), fix=" / ".join(fix.get(c, [])),
            gloss="\n".join(GL.terms_in(text)), src=(v["src"] if v else "")))
    return out


def level_of(code):
    """본점 Level 산식과 동일: 끝에서부터 '00'이 아닌 첫 구간 (1+2+2+2+2+2)."""
    segs = [code[0]] + [code[i:i + 2] for i in range(1, 11, 2)]
    n = 6
    while n > 1 and segs[n - 1] == "00":
        n -= 1
    return n


def parent_of(code):
    lv = level_of(code)
    k = 1 + 2 * (lv - 2)
    return code[:k].ljust(11, "0") if lv > 1 else None


def check(rs=None):
    """(우리 신설 오류, 0930 원안 구조 오류). Level은 본점 산식, 상위는 산식상 상위 코드 존재 여부."""
    rs = rs or rows()
    codes = {r["code"] for r in rs}
    ours, base = [], []
    for r in rs:
        tgt = base if r["code"] in BASE_BY else ours
        if r["lv"] != level_of(r["code"]):
            tgt.append(f'Level 불일치 {r["code"]} 선언 {r["lv"]} / 산식 {level_of(r["code"])}')
        p = parent_of(r["code"])
        if p and p not in codes:
            tgt.append(f'상위 없음 {r["code"]} → 산식상 상위 {p}')
    return ours, base


if __name__ == "__main__":
    o, b = check()
    print("신설 계정 Level·상위:", o or "오류 없음")
    print("0930 원안 구조:", b)
    from collections import Counter
    rs = rows()
    print(len(rs), Counter(r["status"] for r in rs))
