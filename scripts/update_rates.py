#!/usr/bin/env python3
"""한·미 기준금리를 공식 소스에서 읽어 config/market.json의 rates를 자동 갱신.

사용법:
  python scripts/update_rates.py            # 변경분 있으면 market.json 수정, 없으면 그대로
  python scripts/update_rates.py --dry-run  # 수정하지 않고 결과만 출력
  python scripts/update_rates.py --check    # 소스 최신값과 market.json 마지막 값이 다르면 exit 1

소스
  미국: FRED DFEDTARU/DFEDTARL (연방기금금리 목표범위 상단/하단, 일별). 키 불필요.
        FRED는 효력일(결정 다음날) 기준이므로 결정일 = 효력일 - 1일로 기록한다.
  한국: 환경변수 ECOS_API_KEY가 있으면 한국은행 ECOS API(722Y001/0101000),
        없으면 한국은행 기준금리 공개 페이지를 파싱한다(페이지 구조가 바뀌면 실패 → 경고만).

규칙
  - market.json의 마지막 결정일 이후 이벤트만 추가한다. 마지막 값과 같은 금리는 추가하지 않는다.
  - as_of는 건드리지 않는다(periods·groups 등 수작업 항목과 묶여 있음). 대신 rates_checked_at을 갱신한다.
  - sources_rows의 '한국 기준금리'·'미국 기준금리' 행 문구를 마지막 이벤트 기준으로 다시 쓴다.
GitHub Actions(.github/workflows/update-rates.yml)가 매주 실행해 커밋한다.
"""
import argparse, csv, io, json, os, re, sys, urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

MARKET = Path(__file__).resolve().parent.parent / "config" / "market.json"
UA = {"User-Agent": "Mozilla/5.0 (apt-analysis-report rate updater)"}


def get(url, timeout=60, tries=3):
    """GET with 재시도(FRED가 CI 환경에서 간헐적으로 느리다)."""
    import time
    last = None
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read().decode("utf-8", errors="ignore")
        except Exception as e:  # noqa
            last = e; time.sleep(5 * (i + 1))
    raise last


def fred_series(sid):
    """FRED CSV → [(date, float)] (결측 '.' 제외)."""
    txt = get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={sid}")
    out = []
    for row in csv.DictReader(io.StringIO(txt)):
        d, v = row.get("observation_date") or row.get("DATE"), row.get(sid)
        if v and v != ".":
            out.append((datetime.strptime(d, "%Y-%m-%d").date(), float(v)))
    return out


def us_events():
    """FRED 상단/하단 일별 → 변경일 이벤트. rate=상단(소수), note='하단~상단% (효력 YYYY-MM-DD)'."""
    up = fred_series("DFEDTARU"); lo = dict(fred_series("DFEDTARL"))
    ev, prev = [], None
    for d, v in up:
        if v != prev:
            low = lo.get(d, v - 0.25)
            ev.append(dict(date=(d - timedelta(days=1)).isoformat(), rate=round(v / 100, 4),
                           note=f"{low:.2f}~{v:.2f}% (효력 {d.isoformat()})"))
            prev = v
    return ev, up[-1]


def kr_events_ecos(key):
    start = "20170101"; end = date.today().strftime("%Y%m%d")
    url = f"https://ecos.bok.or.kr/api/StatisticSearch/{key}/json/kr/1/10000/722Y001/D/{start}/{end}/0101000"
    js = json.loads(get(url))
    rows = js["StatisticSearch"]["row"]
    ev, prev = [], None
    for r in rows:
        t, v = r["TIME"], float(r["DATA_VALUE"])
        if v != prev:
            ev.append(dict(date=f"{t[:4]}-{t[4:6]}-{t[6:8]}", rate=round(v / 100, 4), note=""))
            prev = v
    return ev, (ev[-1]["date"], ev[-1]["rate"] * 100)


def kr_events_page():
    """한국은행 기준금리 페이지 표(년 / 월일 / 금리) 파싱."""
    import html as _h
    s = get("https://www.bok.or.kr/portal/singl/baseRate/list.do?dataSeCd=01&menuNo=200643")
    ev = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", s, re.S):
        cells = [_h.unescape(re.sub(r"<[^>]+>", "", c)).strip() for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S)]
        if len(cells) >= 3 and re.fullmatch(r"\d{4}", cells[0]):
            m = re.match(r"(\d{1,2})\D+(\d{1,2})", cells[1])
            try:
                rate = float(re.sub(r"[^\d.]", "", cells[2]))
            except ValueError:
                continue
            if m:
                ev.append(dict(date=f"{cells[0]}-{int(m.group(1)):02d}-{int(m.group(2)):02d}", rate=round(rate / 100, 4), note=""))
    ev.sort(key=lambda e: e["date"])
    if not ev:
        raise RuntimeError("한국은행 페이지에서 기준금리 표를 찾지 못했습니다(구조 변경?)")
    return ev, (ev[-1]["date"], ev[-1]["rate"] * 100)


def kr_events():
    key = os.environ.get("ECOS_API_KEY")
    if key:
        try:
            return kr_events_ecos(key)
        except Exception as e:  # noqa
            print(f"[경고] ECOS 실패({e}) → 공개 페이지로 대체", file=sys.stderr)
    return kr_events_page()


def add_kr_notes(ev):
    for i, e in enumerate(ev):
        if e["note"]:
            continue
        if i == 0:
            e["note"] = ""
        else:
            d = (e["rate"] - ev[i - 1]["rate"]) * 100
            e["note"] = f"{d:+.2f}%p" if abs(d) > 1e-9 else "동결"
    return ev


def merge(existing, fresh):
    """existing 마지막 결정일 이후의 fresh 이벤트 중 금리가 바뀐 것만 추가. 반환: (추가된 목록)"""
    last_d = datetime.strptime(existing[-1]["date"][:10], "%Y-%m-%d").date()
    last_r = existing[-1]["rate"]
    added = []
    for e in fresh:
        d = datetime.strptime(e["date"], "%Y-%m-%d").date()
        if d > last_d and abs(e["rate"] - last_r) > 1e-9:
            existing.append(e); added.append(e); last_r = e["rate"]; last_d = d
    return added


def refresh_sources_rows(mkt):
    kr, us = mkt["rates"]["kr"][-1], mkt["rates"]["us"][-1]
    kd = kr["date"].replace("-0", ".").replace("-", "."); ud = us["date"].replace("-0", ".").replace("-", ".")
    us_range = us["note"].split("%")[0] + "%" if "~" in us["note"] else f"{us['rate']:.2%}"
    new = {"한국 기준금리": f"한국은행 금융통화위원회 결정(실제값), {kd} {kr['rate']:.2%}까지 — https://www.bok.or.kr",
           "미국 기준금리": f"FOMC 연방기금금리 목표범위 상단(실제값), {ud} {us_range}까지 — https://www.federalreserve.gov"}
    for row in mkt.get("sources_rows", []):
        if row and row[0] in new:
            row[1] = new[row[0]]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--check", action="store_true", help="최신값 불일치 시 exit 1 (수정 안 함)")
    ap.add_argument("--market", default=str(MARKET))
    a = ap.parse_args()
    with open(a.market, encoding="utf-8") as f:
        mkt = json.load(f)

    report, changed, mismatch, failed = [], False, False, False
    for cc, fn in (("kr", kr_events), ("us", us_events)):
        try:
            fresh, latest = fn()
        except Exception as e:  # noqa
            report.append(f"[{cc}] 조회 실패: {e}"); failed = True; continue
        if cc == "kr":
            fresh = add_kr_notes(fresh)
        added = merge(mkt["rates"][cc], fresh)
        last = mkt["rates"][cc][-1]
        report.append(f"[{cc}] 소스 최신 {latest[0]} {latest[1]:.2f}% / market.json 마지막 {last['date']} {last['rate']*100:.2f}%"
                      + (f" / 추가 {len(added)}건: " + ", ".join(f"{e['date']} {e['rate']*100:.2f}%" for e in added) if added else " / 변경 없음"))
        if added:
            changed = True
        if abs(float(latest[1]) - last["rate"] * 100) > 1e-6:
            # 새 이벤트를 추가했는데도 다르면(FRED 지연 등) 또는 소스가 뒤처지면 경고
            report.append(f"[{cc}] ⚠ 소스 최신값({latest[1]:.2f}%)과 market.json 마지막 값({last['rate']*100:.2f}%)이 다릅니다. "
                          "소스 지연(FRED는 효력일 익일 반영)이거나 수기 입력 오류일 수 있으니 확인하세요.")
            mismatch = True

    print("\n".join(report))
    if a.check:
        sys.exit(1 if mismatch else 0)
    if a.dry_run:
        return
    if failed and not changed:
        print("일부 소스 조회 실패, 추가 이벤트 없음 → market.json 변경하지 않음"); sys.exit(2)
    if not failed:
        mkt["rates_checked_at"] = date.today().isoformat()
    refresh_sources_rows(mkt)
    with open(a.market, "w", encoding="utf-8", newline="\n") as f:
        json.dump(mkt, f, ensure_ascii=False, indent=1); f.write("\n")
    print(f"market.json 저장 (rates_checked_at={mkt.get('rates_checked_at')}, 이벤트 추가={'있음' if changed else '없음'})")


if __name__ == "__main__":
    main()
