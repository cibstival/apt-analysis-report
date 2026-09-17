"""공통 로직: 설정 로드, 실거래 정규화, 지표 계산, 텍스트 플레이스홀더 치환.

build_report.py(엑셀 생성)와 metrics.py(분석 문장 작성용 수치 출력)가 함께 사용한다.
엑셀 셀은 수식으로 계산하고, 여기서는 같은 값을 파이썬으로 계산해 문장에 넣는다.
"""
import csv, json, math, re, statistics
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

PY_M2 = 3.3058  # 1평 = 3.3058㎡


# ---------------------------------------------------------------- 로드
def load_json(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def parse_date(s):
    if isinstance(s, date):
        return s
    s = str(s).strip().replace(".", "-").replace("/", "-")
    if re.fullmatch(r"\d{8}", s):
        s = f"{s[:4]}-{s[4:6]}-{s[6:]}"
    return datetime.strptime(s[:10], "%Y-%m-%d").date()


def half_key(d):
    return f"{d.year}.{'06' if d.month <= 6 else '12'}"


def area_to_pyeong(area, types):
    """전용면적 → 설정의 평형 라벨(가장 가까운 전용면적). 같은 면적이 여러 평형이면 첫 번째."""
    best = min(types, key=lambda t: abs(float(t["area"]) - float(area)))
    if abs(float(best["area"]) - float(area)) > 3:
        return None
    return best["pyeong"]


def load_trades(apt, base_dir):
    """표준 CSV(date,pyeong,area,price,floor,dong,include,note) → dict 리스트.
    pyeong이 비어 있으면 area로 매핑, include가 비어 있으면 이상거래 규칙으로 자동 판정."""
    p = Path(apt["trades_csv"])
    if not p.is_absolute():
        cands = [Path(base_dir) / p]
        if apt.get("_config_dir"):
            cands.insert(0, Path(apt["_config_dir"]) / p)
        p = next((c for c in cands if c.exists()), cands[-1])
    types = apt["types"]
    rows = []
    with open(p, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if not r.get("date") or not r.get("price"):
                continue
            area = float(r["area"]) if r.get("area") else None
            py = r.get("pyeong")
            py = int(float(py)) if py not in (None, "") else (area_to_pyeong(area, types) if area else None)
            if py is None:
                continue
            if area is None:
                area = next(float(t["area"]) for t in types if t["pyeong"] == py)
            inc = r.get("include")
            rows.append({
                "date": parse_date(r["date"]), "pyeong": py, "area": area,
                "price": int(float(str(r["price"]).replace(",", ""))),
                "floor": int(float(r["floor"])) if r.get("floor") not in (None, "") else None,
                "dong": (r.get("dong") or "").strip(),
                "include": None if inc in (None, "") else int(float(inc)),
                "note": r.get("note") or "",
            })
    rows.sort(key=lambda x: x["date"])
    auto_flag_outliers(rows, apt.get("outlier_threshold", 0.25))
    for r in rows:
        r["half"] = half_key(r["date"])
        r["ppp"] = r["price"] / (r["area"] / PY_M2)
    return rows


def auto_flag_outliers(rows, thr):
    """include가 비어 있는 거래만 판정: 같은 평형의 ±6개월 거래 중앙값(자기 제외, 최소 2건)보다 thr 이상 낮으면 0.
    직거래·특수관계 등 저가 거래가 평균을 왜곡하는 것을 막기 위함. 판정 결과는 note에 남긴다."""
    for r in rows:
        if r["include"] is not None:
            continue
        peers = [x["price"] for x in rows if x is not r and x["pyeong"] == r["pyeong"] and abs((x["date"] - r["date"]).days) <= 183]
        if len(peers) >= 2 and r["price"] < statistics.median(peers) * (1 - thr):
            r["include"] = 0
            r["note"] = (r["note"] + " / " if r["note"] else "") + f"자동 제외: 동일평형 ±6개월 중앙값 대비 {thr:.0%} 이상 저가"
        else:
            r["include"] = 1


# ---------------------------------------------------------------- 포맷
def eok(v):
    """만원 → '8억 8,700만원' 스타일 (음수는 '-' 접두)"""
    v = int(round(v))
    if v < 0:
        return "-" + eok(-v)
    e, m = divmod(v, 10000)
    if e and m:
        return f"{e}억 {m:,}만원"
    if e:
        return f"{e}억원"
    return f"{m:,}만원"


def eok_short(v):
    return f"{v/10000:.2f}억".replace(".00억", "억").replace("-0억", "0억")


def pct(v, sign=True, nd=1):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "-"
    return f"{v:+.{nd}%}" if sign else f"{v:.{nd}%}"


# ---------------------------------------------------------------- 지표
def corr(a, b):
    n = len(a)
    if n < 3:
        return float("nan")
    ma, mb = sum(a) / n, sum(b) / n
    sa = math.sqrt(sum((x - ma) ** 2 for x in a))
    sb = math.sqrt(sum((y - mb) ** 2 for y in b))
    if sa == 0 or sb == 0:
        return float("nan")
    return sum((x - ma) * (y - mb) for x, y in zip(a, b)) / (sa * sb)


def rate_at(events, d):
    val = None
    for e in events:
        if parse_date(e["date"]) <= d:
            val = e["rate"]
    return val


# ---------------------------------------------------------------- 정책(세율·요율·기본값)
_POLICY = None


def load_policy(path=None):
    """config/policy.json 로드(캐시). 세율·요율·기본값은 코드가 아니라 이 파일에서 온다."""
    global _POLICY
    if path is None and _POLICY is not None:
        return _POLICY
    path = path or Path(__file__).resolve().parent.parent / "config" / "policy.json"
    _POLICY = load_json(path)
    return _POLICY


def acquisition_tax_rate(p, pol=None):
    """1주택 취득세율(지방교육세 제외). low_upto 이하 low_rate, high_from 초과 high_rate, 사이는 선형."""
    t = (pol or load_policy())["acquisition_tax"]
    if p <= t["low_upto"]:
        return t["low_rate"]
    if p <= t["high_from"]:
        return t["low_rate"] + (p - t["low_upto"]) / (t["high_from"] - t["low_upto"]) * (t["high_rate"] - t["low_rate"])
    return t["high_rate"]


def acquisition_tax_formula(P, pol=None):
    """acquisition_tax_rate와 같은 값을 내는 엑셀 수식. P는 거래가 셀 참조."""
    t = (pol or load_policy())["acquisition_tax"]
    L, H, lr, hr = t["low_upto"], t["high_from"], t["low_rate"], t["high_rate"]
    return f"IF({P}<={L},{lr},IF({P}<={H},{lr}+({P}-{L})/({H}-{L})*({hr}-{lr}),{hr}))"


def brokerage_rate(p, pol=None):
    """중개보수 상한요율(부가세 제외). brackets는 upto 오름차순, 마지막은 upto null."""
    for b in (pol or load_policy())["brokerage"]["brackets"]:
        if b["upto"] is None or p <= b["upto"]:
            return b["rate"]
    return 0.0


def ltcg_rate(pol=None, years=3, one_home=True, resident=True):
    """장기보유특별공제율. 1세대 1주택(표2)은 보유+거주(거주 2년 이상일 때), 아니면 일반(표1)."""
    cg = (pol or load_policy())["capital_gains"]
    y = str(int(years))
    if one_home and resident:
        return cg["ltcg_one_home_hold"].get(y, 0.0) + cg["ltcg_one_home_reside"].get(y, 0.0)
    return cg["ltcg_general"].get(y, 0.0)


def capital_gains_tax(sale, gain, pol=None, one_home=True, ltcg=None):
    """양도세+지방소득세(만원). gain = 양도차익(양도가 − 취득가 − 필요경비). 1세대 1주택이면 고가주택 초과분만 과세."""
    cg = (pol or load_policy())["capital_gains"]
    hv, bd = cg["high_value_threshold"], cg["basic_deduction"]
    if ltcg is None:
        ltcg = ltcg_rate(pol, 3, one_home, True)
    taxable = (0 if sale <= hv else gain * (sale - hv) / sale) if one_home else gain
    tb = max(0.0, taxable * (1 - ltcg) - bd)
    tax = max(tb * b["rate"] - b["deduction"] for b in cg["brackets"])
    return max(0.0, tax) * (1 + cg["local_tax_ratio"])


def capital_gains_formula(SALE, GAIN, HOME1, HV, LTCG, BDED, pol=None):
    """capital_gains_tax와 같은 값을 내는 엑셀 수식. 인자는 셀 참조 문자열."""
    cg = (pol or load_policy())["capital_gains"]
    taxable = f"IF({HOME1}=1,IF({SALE}<={HV},0,{GAIN}*({SALE}-{HV})/{SALE}),{GAIN})"
    tb = f"MAX(0,{taxable}*(1-{LTCG})-{BDED})"
    prog = ",".join(f"{tb}*{b['rate']}-{b['deduction']}" for b in cg["brackets"])
    return f"MAX(0,{prog})*{1 + cg['local_tax_ratio']}"


def brokerage_formula(P, pol=None):
    """brokerage_rate와 같은 값을 내는 엑셀 수식(중첩 IF)."""
    br = (pol or load_policy())["brokerage"]["brackets"]
    expr = str(br[-1]["rate"])
    for b in reversed(br[:-1]):
        expr = f"IF({P}<={b['upto']},{b['rate']},{expr})"
    return expr


def compute_metrics(apt, mkt, trades):
    M = {}
    types = {t["pyeong"]: t for t in apt["types"]}
    focus = apt["user_deal"]["pyeong"]
    comp = apt.get("compare_pyeong")
    user = apt["user_deal"]["price"]
    ft = types[focus]
    inc = [t for t in trades if t["include"]]

    def of_type(py):
        return [t for t in inc if t["pyeong"] == py]

    M["apt_name"] = apt["name"]
    M["focus_label"] = f"{focus}평"
    M["user_price"] = eok(user)
    M["user_price_short"] = eok_short(user)
    M["user_ppp"] = f"{user/(ft['area']/PY_M2):,.0f}만원"
    ftr = of_type(focus)
    if ftr:
        mx = max(ftr, key=lambda t: (t["price"], t["date"]))
        M["focus_max_price"] = eok(mx["price"]); M["focus_max_date"] = mx["date"].isoformat()
        M["focus_max_floor"] = f"{mx['floor']}층" if mx["floor"] else ""
        M["user_vs_max"] = pct(user / mx["price"] - 1)
        M["_focus_max"] = mx
    # 사용자 물건(동·호·층·매수/매도)과 같은 동 실거래 비교
    ud = apt["user_deal"]
    dong = str(ud.get("dong") or "").strip(); ho = str(ud.get("ho") or "").strip()
    floor = ud.get("floor") or (int(ho[:-2]) if ho[:-2].isdigit() else None)
    M["user_dong"] = f"{dong}동" if dong else ""; M["user_ho"] = f"{ho}호" if ho else ""
    M["user_floor"] = f"{floor}층" if floor else ""
    M["user_unit"] = " ".join(x for x in (M["user_dong"], M["user_ho"] or M["user_floor"]) if x) or "동·호 미입력"
    M["user_side"] = ud.get("side", "미확인")
    M["_user_dong"] = dong; M["_user_floor"] = floor
    if dong:
        dtr = [t for t in ftr if str(t.get("dong") or "") == dong]
        M["dong_n"] = str(len(dtr))
        if dtr:
            dm = max(dtr, key=lambda t: (t["price"], t["date"]))
            M["dong_max_price"] = eok(dm["price"]); M["dong_max_date"] = dm["date"].isoformat()
            M["dong_max_floor"] = f"{dm['floor']}층" if dm["floor"] else ""
            M["dong_avg_price"] = eok(sum(t["price"] for t in dtr) / len(dtr))
            M["user_vs_dong_max"] = pct(user / dm["price"] - 1)
            M["_dong_max"] = dm
    if ft.get("kb_price"):
        M["kb_focus"] = eok(ft["kb_price"]); M["user_vs_kb"] = pct(user / ft["kb_price"] - 1)
    if ft.get("kb_jeonse"):
        M["jeonse_focus"] = eok(ft["kb_jeonse"]); M["jeonse_ratio"] = pct(ft["kb_jeonse"] / user, sign=False)
    pr = apt.get("price_refs", {})
    if pr.get("private_price"):
        M["private_price"] = eok(pr["private_price"]["value"]); M["user_vs_private"] = pct(user / pr["private_price"]["value"] - 1)
    nb = pr.get("new_build")
    if nb:
        M["newbuild_member"] = eok(nb["value"])
        if nb.get("member_ratio"):
            g = nb["value"] / nb["member_ratio"]
            M["newbuild_general"] = eok(g); M["user_vs_newbuild_general"] = pct(user / g - 1)
    if comp and of_type(comp):
        cm = max(of_type(comp), key=lambda t: (t["price"], t["date"]))
        M["compare_label"] = f"{comp}평"; M["compare_max_price"] = eok(cm["price"])
        cppp = cm["price"] / (types[comp]["area"] / PY_M2)
        M["compare_max_ppp"] = f"{cppp:,.0f}만원"
        M["user_ppp_vs_compare"] = pct((user / (ft["area"] / PY_M2)) / cppp - 1)

    # 연도별
    years = defaultdict(lambda: {"all": 0})
    for t in trades:
        years[t["date"].year]["all"] += 1
    for t in inc:
        years[t["date"].year].setdefault(t["pyeong"], []).append(t["price"])
    ystats = {}
    for y in sorted(years):
        row = {"total": years[y]["all"]}
        for py in types:
            v = years[y].get(py)
            if v:
                row[f"{py}평"] = {"n": len(v), "avg": round(sum(v) / len(v)), "max": max(v), "min": min(v)}
        ystats[y] = row
    M["_yearly"] = ystats
    if ftr:
        # 고점(개별 최고) 이후 최저 → 최대 낙폭
        peak = None; mdd = 0; pk = None; tr = None
        for t in sorted(ftr, key=lambda x: x["date"]):
            if peak is None or t["price"] > peak["price"]:
                peak = t
            dd = t["price"] / peak["price"] - 1
            if dd < mdd:
                mdd = dd; pk = peak; tr = t
        if pk:
            M["focus_max_drawdown"] = pct(mdd)
            M["focus_drawdown_desc"] = f"{eok_short(pk['price'])}({pk['date']:%Y.%m}) → {eok_short(tr['price'])}({tr['date']:%Y.%m})"
    # 반기
    halves = defaultdict(lambda: {"n": 0, "ppp": []})
    for t in trades:
        halves[t["half"]]["n"] += 1
    for t in inc:
        halves[t["half"]]["ppp"].append(t["ppp"])
    M["_halves"] = {k: {"n": v["n"], "avg_ppp": round(sum(v["ppp"]) / len(v["ppp"])) if v["ppp"] else None} for k, v in sorted(halves.items())}

    # 금리
    kr, us = mkt["rates"]["kr"], mkt["rates"]["us"]
    asof = parse_date(mkt["as_of"])
    krn, usn = rate_at(kr, asof), rate_at(us, asof)
    M["kr_rate_now"] = f"{krn:.2%}"; M["us_rate_now"] = f"{usn:.2%}"; M["spread_now"] = f"{(krn-usn)*100:+.2f}%p"

    # 금리차-거래량 상관
    per = [p for p in mkt["periods"] if not p.get("partial")]
    sp, krl, cnt = [], [], []
    for p in per:
        d = parse_date(p["date"])
        k, u = rate_at(kr, d), rate_at(us, d)
        sp.append((k - u) * 100); krl.append(k * 100); cnt.append(halves[p["label"]]["n"] if p["label"] in halves else 0)
    M["corr_spread_volume"] = f"{corr(sp, cnt):+.2f}"
    M["corr_krrate_volume"] = f"{corr(krl, cnt):+.2f}"
    dk = [krl[i] - krl[i - 1] for i in range(1, len(krl))]; dv = [cnt[i] - cnt[i - 1] for i in range(1, len(cnt))]
    M["corr_dkr_dvolume"] = f"{corr(dk, dv):+.2f}"

    # 시나리오·손익분기 (엑셀 수식과 동일 로직)
    ol = apt.get("text", {}).get("outlook", {})
    sc = ol.get("scenarios", [])
    if sc:
        M["exp_1y"] = pct(sum(s[1] * s[2] for s in sc)); M["exp_3y"] = pct(sum(s[1] * s[3] for s in sc))
    md = ol.get("model", {})
    if md:
        pol = load_policy(); ld = pol["loan"]; vat = 1 + pol["brokerage"]["vat_ratio"]; edu = 1 + pol["acquisition_tax"]["local_education_tax_ratio"]
        sell = pol["brokerage"]["sell_rate_simplified"]
        loan = min(user * md.get("ltv", ld["ltv_default"]), md.get("loan_cap", ld["loan_cap_default"]))
        acq = user * acquisition_tax_rate(user, pol) * edu + user * brokerage_rate(user, pol) * vat + md.get("etc_cost", ld["etc_cost_default"])
        costs = user + acq + loan * md.get("rate", ld["rate_default"]) * 3 + md.get("holding_tax", ld["holding_tax_default"]) * 3 - md.get("housing_saving", 0) * 3
        be = costs / (1 - sell) / user - 1
        cash = user - loan + acq
        be2 = (costs + cash * md.get("deposit_rate", ld["deposit_rate_default"]) * 3) / (1 - sell) / user - 1
        M["breakeven_3y"] = pct(be); M["breakeven_3y_opp"] = pct(be2)
        M["loan_amount"] = eok(loan); M["cash_needed"] = eok(cash); M["acq_cost"] = eok(acq)
        # 시나리오별 양도세(12억 초과분 과세)와 세후 순손익
        cg = pol["capital_gains"]; ud = apt["user_deal"]
        one_home = md.get("one_home", 1) == 1
        ltcg = md["ltcg"] if md.get("ltcg") is not None else ltcg_rate(pol, 3, one_home, bool(ud.get("resident", True)))
        M["ltcg_rate"] = pct(ltcg, sign=False, nd=0); M["hv_threshold"] = eok_short(cg["high_value_threshold"])
        names = {0: "bull", 1: "base", 2: "bear", 3: "stress"}
        parts = []
        for i, sc_ in enumerate(sc):
            sale = user * (1 + sc_[3]); sell_cost = sale * brokerage_rate(sale, pol) * vat
            gain = sale - user - acq - sell_cost
            tax = capital_gains_tax(sale, gain, pol, one_home, ltcg)
            net = sale - sell_cost - user - acq - loan * md.get("rate", ld["rate_default"]) * 3 - md.get("holding_tax", ld["holding_tax_default"]) * 3 + md.get("housing_saving", 0) * 3 - tax
            k = names.get(i, f"s{i}")
            M[f"sale_{k}"] = eok(sale); M[f"cgt_{k}"] = eok(tax) if tax >= 1 else "0원"; M[f"net_{k}"] = eok(net)
            parts.append(f"{sc_[0]} {eok_short(sale)} → 양도세 {eok_short(tax) if tax >= 1 else '0'}")
        M["cgt_scenarios"] = " / ".join(parts)
    return M


# ---------------------------------------------------------------- 치환
_PH = re.compile(r"<<([a-zA-Z0-9_]+)>>")


def fill(obj, M):
    """문자열 안의 <<키>>를 지표 값으로 치환(재귀). 없는 키는 그대로 남겨 검토 시 드러나게 한다."""
    if isinstance(obj, str):
        return _PH.sub(lambda m: str(M.get(m.group(1), m.group(0))), obj)
    if isinstance(obj, list):
        return [fill(x, M) for x in obj]
    if isinstance(obj, dict):
        return {k: fill(v, M) for k, v in obj.items()}
    return obj


def find_unfilled(obj, path="text"):
    out = []
    if isinstance(obj, str):
        out += [(path, m) for m in _PH.findall(obj)]
    elif isinstance(obj, list):
        for i, x in enumerate(obj):
            out += find_unfilled(x, f"{path}[{i}]")
    elif isinstance(obj, dict):
        for k, v in obj.items():
            out += find_unfilled(v, f"{path}.{k}")
    return out


def safe_sheet(name, suffix):
    n = re.sub(r"[\[\]:*?/\\]", "", name)
    return (n[: 31 - len(suffix)] + suffix)
