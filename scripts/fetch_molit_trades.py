#!/usr/bin/env python3
"""국토교통부 아파트 매매 실거래가 API(공공데이터포털) → 표준 실거래 CSV.

필요: 공공데이터포털(data.go.kr)에서 '국토교통부_아파트 매매 실거래가 상세 자료' 활용신청 후
      일반 인증키(Decoding)를 환경변수 MOLIT_API_KEY 에 설정.

사용법:
  python scripts/fetch_molit_trades.py --apt config/apartments/<단지>.json --from 201701 --to 202609
  (lawd_cd·단지명·법정동·지번은 아파트 설정의 region / match 항목을 사용)

동작:
  - 월별로 호출해 단지명(공백 무시 부분일치)·법정동·지번으로 필터
  - 해제(취소) 거래 제외, 전용면적 → 설정 types의 평형으로 매핑
  - 기존 CSV가 있으면 병합(같은 날짜·면적·층·가격은 중복으로 보고 기존 행의 dong/include/note 유지)

주의: API 응답 필드명은 개편될 수 있다. 첫 실행 시 --debug 로 원본 item 1건을 출력해 필드명을 확인하고,
      FIELD 매핑에 없는 이름이면 아래 FIELD 사전에 추가한다.
"""
import argparse, csv, json, os, sys, time, urllib.parse, urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib_report import area_to_pyeong, load_json

URL = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev"
FIELD = {  # 표준키: 후보 필드명(신규 영문 → 구 한글)
    "name": ["aptNm", "아파트"], "price": ["dealAmount", "거래금액"], "year": ["dealYear", "년"], "month": ["dealMonth", "월"],
    "day": ["dealDay", "일"], "area": ["excluUseAr", "전용면적"], "floor": ["floor", "층"], "umd": ["umdNm", "법정동"],
    "jibun": ["jibun", "지번"], "cancel": ["cdealType", "해제여부"], "dong": ["aptDong", "동"], "gbn": ["dealingGbn", "거래유형"],
}


def get(item, key):
    for k in FIELD[key]:
        el = item.find(k)
        if el is not None and el.text is not None:
            return el.text.strip()
    return ""


def months(a, b):
    y, m = int(a[:4]), int(a[4:]); Y, Mo = int(b[:4]), int(b[4:])
    while (y, m) <= (Y, Mo):
        yield f"{y}{m:02d}"; m += 1
        if m == 13: y += 1; m = 1


def call(key, lawd, ym, page, debug=False):
    q = urllib.parse.urlencode({"serviceKey": key, "LAWD_CD": lawd, "DEAL_YMD": ym, "pageNo": page, "numOfRows": 1000})
    with urllib.request.urlopen(f"{URL}?{q}", timeout=30) as r:
        root = ET.fromstring(r.read())
    code = (root.findtext(".//resultCode") or "").strip()
    if code not in ("00", "000", ""):
        raise RuntimeError(f"API 오류 {code}: {root.findtext('.//resultMsg')}")
    items = root.findall(".//item")
    if debug and items:
        print(ET.tostring(items[0], encoding="unicode")); debug = False
    total = int(root.findtext(".//totalCount") or 0)
    return items, total


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apt", required=True); ap.add_argument("--from", dest="frm", required=True); ap.add_argument("--to", required=True)
    ap.add_argument("--debug", action="store_true")
    a = ap.parse_args()
    key = os.environ.get("MOLIT_API_KEY")
    if not key: sys.exit("환경변수 MOLIT_API_KEY 가 없습니다. data.go.kr 인증키(Decoding)를 설정하세요.")
    apt = load_json(a.apt); apt["_config_dir"] = str(Path(a.apt).resolve().parent); root = Path(a.apt).resolve().parent.parent.parent
    reg, match = apt["region"], apt.get("match", {})
    want = match.get("apt_name_contains", apt["name"]).replace(" ", "")
    out_p = root / apt["trades_csv"]
    rows = []
    for ym in months(a.frm, a.to):
        page = 1
        while True:
            items, total = call(key, reg["lawd_cd"], ym, page, a.debug); a.debug = False
            for it in items:
                if want not in get(it, "name").replace(" ", ""): continue
                if match.get("umd") and match["umd"] not in get(it, "umd"): continue
                if match.get("jibun") and get(it, "jibun") != str(match["jibun"]): continue
                if get(it, "cancel") in ("O", "o", "Y"): continue
                area = float(get(it, "area")); py = area_to_pyeong(area, apt["types"])
                if py is None:
                    print(f"[경고] 평형 매핑 실패 전용 {area}㎡ → config types 확인"); continue
                d = f"{get(it,'year')}-{int(get(it,'month')):02d}-{int(get(it,'day')):02d}"
                rows.append({"date": d, "pyeong": py, "area": area, "price": get(it, "price").replace(",", ""),
                             "floor": get(it, "floor"), "dong": get(it, "dong"), "include": "",
                             "note": "직거래" if "직거래" in get(it, "gbn") else ""})
            if page * 1000 >= total: break
            page += 1
        time.sleep(0.15)
    # 병합
    existing = {}
    if out_p.exists():
        with open(out_p, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                existing[(r["date"], str(float(r["area"])), r["floor"], r["price"])] = r
    for r in rows:
        k = (r["date"], str(float(r["area"])), r["floor"], r["price"])
        if k not in existing: existing[k] = r
    merged = sorted(existing.values(), key=lambda r: r["date"])
    out_p.parent.mkdir(parents=True, exist_ok=True)
    with open(out_p, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["date", "pyeong", "area", "price", "floor", "dong", "include", "note"]); w.writeheader()
        for r in merged: w.writerow({k: r.get(k, "") for k in w.fieldnames})
    print(f"신규 {len(rows)}건 조회, 총 {len(merged)}건 저장 → {out_p}")


if __name__ == "__main__":
    main()
