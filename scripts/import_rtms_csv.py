#!/usr/bin/env python3
"""국토부 실거래가 공개시스템(rt.molit.go.kr)에서 내려받은 CSV/XLSX → 표준 실거래 CSV.

API 키가 없을 때 쓰는 대체 경로. 사이트에서 '아파트 매매 → 시군구 선택 → 기간 → 다운로드'.
사용법: python scripts/import_rtms_csv.py --apt config/apartments/<단지>.json 파일1.csv [파일2.csv ...]
헤더 앞의 안내문 줄은 자동으로 건너뛰고, 단지명·전용면적·계약년월·계약일·거래금액·층·동·해제사유발생일 열을 찾는다.
"""
import argparse, csv, io, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib_report import area_to_pyeong, load_json


def read_rows(p):
    raw = Path(p).read_bytes()
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try: text = raw.decode(enc); break
        except UnicodeDecodeError: continue
    lines = text.splitlines()
    h = next(i for i, l in enumerate(lines) if "단지명" in l and "거래금액" in l)
    return list(csv.DictReader(io.StringIO("\n".join(lines[h:]))))


def col(r, *names):
    for k in r:
        if k and any(n in k for n in names): return (r[k] or "").strip()
    return ""


ap = argparse.ArgumentParser(); ap.add_argument("--apt", required=True); ap.add_argument("files", nargs="+")
a = ap.parse_args()
apt = load_json(a.apt); apt["_config_dir"] = str(Path(a.apt).resolve().parent); root = Path(a.apt).resolve().parent.parent.parent
want = apt.get("match", {}).get("apt_name_contains", apt["name"]).replace(" ", "")
jibun = str(apt.get("match", {}).get("jibun", ""))
out = []
for f in a.files:
    for r in read_rows(f):
        if want not in col(r, "단지명").replace(" ", ""): continue
        if jibun and jibun not in col(r, "번지"): continue
        if col(r, "해제사유"): continue
        area = float(col(r, "전용면적")); py = area_to_pyeong(area, apt["types"])
        if py is None: continue
        ym = col(r, "계약년월"); dd = int(col(r, "계약일"))
        out.append({"date": f"{ym[:4]}-{ym[4:6]}-{dd:02d}", "pyeong": py, "area": area, "price": col(r, "거래금액").replace(",", ""),
                    "floor": col(r, "층"), "dong": col(r, "동") if "동" in r else "", "include": "",
                    "note": "직거래" if "직거래" in col(r, "거래유형") else ""})
out.sort(key=lambda x: x["date"])
p = root / apt["trades_csv"]; p.parent.mkdir(parents=True, exist_ok=True)
with open(p, "w", newline="", encoding="utf-8-sig") as fh:
    w = csv.DictWriter(fh, fieldnames=["date", "pyeong", "area", "price", "floor", "dong", "include", "note"]); w.writeheader(); w.writerows(out)
print(f"{len(out)}건 저장 → {p}")
