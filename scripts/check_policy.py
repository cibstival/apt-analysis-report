#!/usr/bin/env python3
"""정책 기준(config/policy.json) 유효기간 점검·갱신.

사용법:
  python scripts/check_policy.py                 # 상태 출력. 오래됐으면 exit 1
  python scripts/check_policy.py --touch --note "확인 결과 변경 없음" --source "정책브리핑 2026.10.2"
  python scripts/check_policy.py --touch --note "취득세 구간 변경 반영" --source "지방세법 개정 2027.1.1"
  python scripts/check_policy.py --json          # 기계 판독용

--touch 는 checked_at을 오늘로 바꾸고 changelog에 한 줄을 남긴다.
값 자체(세율·요율·규제 내용)는 policy.json을 직접 고친 뒤 --touch로 기록한다.
"""
import argparse, json, sys
from datetime import date, datetime
from pathlib import Path

POLICY = Path(__file__).resolve().parent.parent / "config" / "policy.json"
MARKET = Path(__file__).resolve().parent.parent / "config" / "market.json"


def load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save(p, obj):
    with open(p, "w", encoding="utf-8", newline="\n") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
        f.write("\n")


def status(pol, mkt, today):
    checked = datetime.strptime(pol["checked_at"], "%Y-%m-%d").date()
    age = (today - checked).days
    limit = int(pol.get("check_interval_days", 30))
    items = []
    for key in ("acquisition_tax", "brokerage", "capital_gains", "loan", "regulation", "tax_reform"):
        sec = pol.get(key)
        if sec:
            items.append(dict(key=key, label=sec.get("label", key), effective_from=sec.get("effective_from", ""),
                              source=sec.get("source", ""), status=sec.get("status", "")))
    mk_as_of = datetime.strptime(mkt["as_of"], "%Y-%m-%d").date()
    rates_checked = mkt.get("rates_checked_at", mkt["as_of"])
    return dict(today=today.isoformat(), checked_at=pol["checked_at"], age_days=age, limit_days=limit, stale=age > limit,
                market_as_of=mkt["as_of"], market_age_days=(today - mk_as_of).days, rates_checked_at=rates_checked,
                last_kr=mkt["rates"]["kr"][-1], last_us=mkt["rates"]["us"][-1], items=items,
                changelog_tail=pol.get("changelog", [])[-3:])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--touch", action="store_true", help="checked_at을 오늘로 갱신하고 changelog 추가")
    ap.add_argument("--note", default="", help="--touch 시 changelog 메모")
    ap.add_argument("--source", default="", help="--touch 시 근거 출처")
    ap.add_argument("--date", default=None, help="오늘 날짜 대체(YYYY-MM-DD, 테스트용)")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    today = datetime.strptime(a.date, "%Y-%m-%d").date() if a.date else date.today()
    pol = load(POLICY); mkt = load(MARKET)

    if a.touch:
        if not a.note:
            sys.exit("--touch 에는 --note 가 필요합니다(무엇을 확인했는지 남기세요).")
        pol["checked_at"] = today.isoformat()
        pol.setdefault("changelog", []).append({"date": today.isoformat(), "note": a.note, "source": a.source})
        save(POLICY, pol)
        print(f"policy.json checked_at → {today}, changelog 추가: {a.note}")
        return

    st = status(pol, mkt, today)
    if a.json:
        print(json.dumps(st, ensure_ascii=False, indent=1)); sys.exit(1 if st["stale"] else 0)

    flag = "⚠ 재확인 필요" if st["stale"] else "OK"
    print(f"[정책 기준] 확인일 {st['checked_at']} ({st['age_days']}일 경과, 기준 {st['limit_days']}일) → {flag}")
    for it in st["items"]:
        extra = f"  [{it['status']}]" if it["status"] else ""
        print(f"  - {it['label']}: 시행 {it['effective_from']}~  출처 {it['source']}{extra}")
    print(f"[시장 데이터] market.json as_of {st['market_as_of']} ({st['market_age_days']}일 경과), 금리 자동확인 {st['rates_checked_at']}")
    print(f"  - 한국 기준금리 마지막: {st['last_kr']['date']} {st['last_kr']['rate']:.2%}  / 미국 상단: {st['last_us']['date']} {st['last_us']['rate']:.2%}")
    if st["changelog_tail"]:
        print("[최근 변경]")
        for c in st["changelog_tail"]:
            print(f"  - {c['date']} {c['note']} ({c.get('source', '')})")
    if st["stale"]:
        print("\n→ references/research_guide.md '7. 정책 점검' 순서대로 확인한 뒤 --touch 로 기록하세요.")
        sys.exit(1)


if __name__ == "__main__":
    main()
