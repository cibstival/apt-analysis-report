#!/usr/bin/env python3
"""국토부 실거래가 공개시스템(rt.molit.go.kr)에서 API 키 없이 단지 실거래 CSV를 직접 내려받아 표준 CSV로 저장.

사용법:
  python scripts/fetch_rtms_direct.py --apt config/apartments/<단지>.json [--from 2017] [--to 2026-09-17] [--rent]
  python scripts/fetch_rtms_direct.py --apt <설정> --list        # 해당 읍면동의 단지 목록만 출력(단지명 확인용)

동작
  1) 설정의 region.sido/gu/dong 으로 시도→시군구→읍면동 코드를 조회하고, match.apt_name_contains 로 단지 코드를 찾는다.
  2) 사이트 제한(시군구 단위 1년)에 맞춰 연도별로 CSV(ptXlsCSVDown.do)를 POST 요청한다.
  3) 받은 파일을 import_rtms_csv.py 와 같은 규칙으로 정규화해 trades_csv 에 저장한다(해제 거래 제외, 동·층 포함).
  --rent 는 전월세 자료를 data/<단지>_rent.csv 로 저장한다(전세 시세 확인용).
사이트 구조가 바뀌면 실패할 수 있다. 그때는 fetch_molit_trades.py(API 키) 또는 수동 다운로드 + import_rtms_csv.py 를 쓴다.
"""
import argparse, csv, io, sys, time
from datetime import date
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib_report import area_to_pyeong, load_json

try:
    import requests
except ImportError:
    sys.exit("pip install requests 필요")

BASE = "https://rt.molit.go.kr"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120"


def session():
    s = requests.Session(); s.headers["User-Agent"] = UA
    s.get(BASE + "/pt/xls/xls.do?mobileAt=", timeout=30)
    return s


def resolve_codes(s, sido, gu, dong):
    sidos = s.post(BASE + "/data/sido.do", timeout=30).json()
    sd = next((x for x in sidos if sido[:2] in x["ctprvnNm"]), None)
    if not sd: sys.exit(f"시도 '{sido}' 를 찾지 못함: {[x['ctprvnNm'] for x in sidos]}")
    sggs = s.post(BASE + "/data/sgg.do", data={"signguCode": sd["signguCode"][:2]}, timeout=30).json()
    sg = next((x for x in sggs if x["signguNm"] == gu), None) or next((x for x in sggs if gu in x["signguNm"]), None)
    if not sg: sys.exit(f"시군구 '{gu}' 를 찾지 못함: {[x['signguNm'] for x in sggs]}")
    form = base_form(sd["signguCode"], sg["signguCode"])
    # 읍면동·단지 목록은 '해당 기간에 신고 자료가 있는 것만' 반환하므로 최근 1년을 조건으로 준다
    td = date.today(); form["srhFromDt"] = f"{td.year - 1}-{td.month:02d}-01"; form["srhToDt"] = td.isoformat()
    emds = s.post(BASE + "/cmm/ptEmdList.do", data=form, timeout=30).json()["emdList"]
    em = next((x for x in emds if x["ladNm"] == dong), None) or next((x for x in emds if dong in x["ladNm"]), None)
    if not em: sys.exit(f"읍면동 '{dong}' 를 찾지 못함: {[x['ladNm'] for x in emds]}")
    form["srhEmdCd"] = em["emdCode"]
    hs = s.post(BASE + "/cmm/ptHsmpList.do", data=form, timeout=30).json()["hsmpList"]
    return form, dict(sido=sd["ctprvnNm"], gu=sg["signguNm"], dong=em["ladNm"]), hs


def base_form(sido_cd, sgg_cd):
    return {"srhThingNo": "A", "srhDelngSecd": "1", "srhAddrGbn": "1", "srhLfstsSecd": "1", "srhNewRonSecd": "1",
            "srhSidoCd": sido_cd, "srhSggCd": sgg_cd, "srhEmdCd": "", "srhHsmpCd": "", "srhLoadCd": "",
            "srhFromDt": "", "srhToDt": "", "srhArea": "", "srhLrArea": "", "srhRoadNm": "", "srhFromAmount": "", "srhToAmount": "",
            "sidoNm": "", "sggNm": "", "emdNm": "", "hsmpNm": "", "areaNm": "전체", "loadNm": "전체", "mobileAt": ""}


def download_years(s, form, y0, to_dt, rent=False):
    form = dict(form); form["srhDelngSecd"] = "2" if rent else "1"
    texts = []
    for y in range(y0, to_dt.year + 1):
        form["srhFromDt"] = f"{y}-01-01"; form["srhToDt"] = f"{y}-12-31" if y < to_dt.year else to_dt.isoformat()
        r = s.post(BASE + "/pt/xls/ptXlsCSVDown.do", data=form, timeout=90)
        txt = r.content.decode("cp949", errors="ignore")
        n = sum(1 for l in txt.splitlines() if l.count(",") > 8)
        print(f"  {y}: {max(n - 1, 0)}건", file=sys.stderr); texts.append(txt); time.sleep(1.2)
    return texts


def parse(txt):
    lines = txt.splitlines()
    try:
        h = next(i for i, l in enumerate(lines) if "단지명" in l and ("거래금액" in l or "보증금" in l))
    except StopIteration:
        return []
    return list(csv.DictReader(io.StringIO("\n".join(lines[h:]))))


def col(r, *names):
    for k in r:
        if k and any(n in k for n in names): return (r[k] or "").strip()
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apt", required=True); ap.add_argument("--from", dest="y0", type=int, default=2017)
    ap.add_argument("--to", default=date.today().isoformat()); ap.add_argument("--rent", action="store_true"); ap.add_argument("--list", action="store_true")
    a = ap.parse_args()
    apt = load_json(a.apt); root = Path(a.apt).resolve().parent.parent.parent
    rg = apt["region"]; want = apt.get("match", {}).get("apt_name_contains", apt["name"]).replace(" ", "")
    s = session()
    form, names, hs = resolve_codes(s, rg["sido"], rg["gu"], rg["dong"])
    if a.list:
        for x in hs: print(x["aprpnHsmpCode"], x["aprpnHsmpNm"])
        return
    cand = [x for x in hs if want in x["aprpnHsmpNm"].replace(" ", "")]
    if len(cand) != 1:
        sys.exit(f"단지 '{want}' 후보 {len(cand)}개: {[x['aprpnHsmpNm'] for x in cand] or [x['aprpnHsmpNm'] for x in hs]}\n→ match.apt_name_contains 를 국토부 표기와 맞추세요 (--list 로 확인).")
    form.update({"srhHsmpCd": cand[0]["aprpnHsmpCode"], "sidoNm": names["sido"], "sggNm": names["gu"], "emdNm": names["dong"], "hsmpNm": cand[0]["aprpnHsmpNm"]})
    print(f"단지: {cand[0]['aprpnHsmpNm']} (코드 {cand[0]['aprpnHsmpCode']}), {names['sido']} {names['gu']} {names['dong']}", file=sys.stderr)
    to_dt = date.fromisoformat(a.to)
    rows = [r for t in download_years(s, form, a.y0, to_dt, rent=a.rent) for r in parse(t)]
    jibun = str(apt.get("match", {}).get("jibun", ""))
    if jibun: rows = [r for r in rows if jibun in col(r, "번지")]

    if a.rent:
        out = []
        for r in rows:
            out.append({"date": f"{col(r,'계약년월')[:4]}-{col(r,'계약년월')[4:6]}-{int(col(r,'계약일') or 1):02d}", "kind": col(r, "전월세구분"),
                        "area": col(r, "전용면적"), "deposit": col(r, "보증금").replace(",", ""), "rent": col(r, "월세금").replace(",", ""),
                        "floor": col(r, "층"), "contract": col(r, "계약구분"), "renewal": col(r, "갱신요구권")})
        out.sort(key=lambda x: x["date"])
        p = root / "data" / (Path(apt["trades_csv"]).stem.replace("_trades", "") + "_rent.csv")
        with open(p, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(out[0].keys()) if out else ["date"]); w.writeheader(); w.writerows(out)
        print(f"{len(out)}건 저장 → {p}"); return

    out = []
    for r in rows:
        if col(r, "해제사유").strip("- "): continue
        area = float(col(r, "전용면적")); py = area_to_pyeong(area, apt["types"])
        if py is None:
            print(f"[경고] 전용 {area}㎡ 가 types 에 없음 → 건너뜀 (types 에 평형 추가)", file=sys.stderr); continue
        ym = col(r, "계약년월"); dd = int(col(r, "계약일"))
        out.append({"date": f"{ym[:4]}-{ym[4:6]}-{dd:02d}", "pyeong": py, "area": area, "price": col(r, "거래금액").replace(",", ""),
                    "floor": col(r, "층"), "dong": (r.get("동") or "").strip("- "), "include": "",
                    "note": "직거래" if "직거래" in col(r, "거래유형") else ""})
    out.sort(key=lambda x: x["date"])
    p = root / apt["trades_csv"]; p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["date", "pyeong", "area", "price", "floor", "dong", "include", "note"]); w.writeheader(); w.writerows(out)
    print(f"{len(out)}건 저장 → {p}  (동 표기 {sum(1 for x in out if x['dong'])}건)")


if __name__ == "__main__":
    main()
