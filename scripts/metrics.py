#!/usr/bin/env python3
"""분석 문장 작성용 수치 출력.

사용법: python scripts/metrics.py --apt config/apartments/<단지>.json [--market config/market.json]

출력(JSON):
  placeholders : 텍스트에 <<키>>로 넣을 수 있는 값 (예: <<user_vs_max>> → +10.9%)
  yearly       : 연도별·평형별 거래 건수/평균/최고/최저 (만원)
  halves       : 반기별 거래건수·평균 전용평단가
  unfilled     : 설정 텍스트 중 치환되지 않은 <<키>> 목록
분석 문장을 쓰기 전에 반드시 이 출력을 보고 숫자를 맞춘다.
"""
import argparse, json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib_report import compute_metrics, fill, find_unfilled, load_json, load_trades

ap = argparse.ArgumentParser()
ap.add_argument("--apt", required=True)
ap.add_argument("--market", default=str(Path(__file__).resolve().parent.parent / "config" / "market.json"))
a = ap.parse_args()
apt = load_json(a.apt); apt["_config_dir"] = str(Path(a.apt).resolve().parent); mkt = load_json(a.market)
root = Path(a.apt).resolve().parent.parent.parent
tr = load_trades(apt, root)
M = compute_metrics(apt, mkt, tr)
out = {
    "placeholders": {k: v for k, v in M.items() if not k.startswith("_")},
    "yearly": M["_yearly"],
    "halves": M["_halves"],
    "trades_total": len(tr),
    "excluded": [dict(date=t["date"].isoformat(), pyeong=t["pyeong"], price=t["price"], note=t["note"]) for t in tr if not t["include"]],
    "recent": [dict(date=t["date"].isoformat(), pyeong=t["pyeong"], floor=t["floor"], price=t["price"]) for t in tr[-12:]],
    "unfilled": find_unfilled(fill(apt.get("text", {}), M)),
}
print(json.dumps(out, ensure_ascii=False, indent=1, default=str))
