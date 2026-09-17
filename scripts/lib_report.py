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
    """만원 → '8억 8,700만원' 스타일"""
    v = int(round(v))
    e, m = divmod(v, 10000)
    if e and m:
        return f"{e}억 {m:,}만원"
    if e:
        return f"{e}억원"
    return f"{m:,}만원"


def eok_short(v):
    return f"{v/10000:.2f}억".replace(".00억", "억")


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


def acquisition_tax_rate(p):
    """1주택 취득세율(지방교육세 제외). 6억 이하 1%, 6~9억 산식, 9억 초과 3%."""
    if p <= 60000:
        return 0.01
    if p <= 90000:
        return (p / 10000 * 2 / 3 - 3) / 100
    return 0.03


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
        loan = min(user * md.get("ltv", 0.4), md.get("loan_cap", 60000))
        acq = user * acquisition_tax_rate(user) * 1.1 + user * 0.004 * 1.1 + md.get("etc_cost", 300)
        costs = user + acq + loan * md.get("rate", 0.05) * 3 + md.get("holding_tax", 150) * 3 - md.get("housing_saving", 0) * 3
        be = costs / (1 - 0.0044) / user - 1
        cash = user - loan + acq
        be2 = (costs + cash * md.get("deposit_rate", 0.035) * 3) / (1 - 0.0044) / user - 1
        M["breakeven_3y"] = pct(be); M["breakeven_3y_opp"] = pct(be2)
        M["loan_amount"] = eok(loan); M["cash_needed"] = eok(cash); M["acq_cost"] = eok(acq)
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
