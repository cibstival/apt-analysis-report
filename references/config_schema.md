# 설정 파일 스키마

## 목차
0. policy.json (정책·세제 기준)
1. market.json (시장 공통)
2. apartments/<단지>.json (단지별)
3. 실거래 CSV
4. 플레이스홀더 목록

---
## 0. policy.json
코드는 세율·요율·기본값을 여기서만 읽는다. `scripts/check_policy.py`가 `checked_at` 경과를 점검한다.

| 키 | 설명 |
|---|---|
| `checked_at`, `check_interval_days` | 마지막 확인일과 재확인 주기(일). `check_policy.py --touch`로 갱신 |
| `changelog[]` | `{date, note, source}` 확인·변경 이력. 출처_가정 시트에 최근 3건 표시 |
| `acquisition_tax` | 1주택 취득세: `low_upto`(만원) 이하 `low_rate`, `high_from` 초과 `high_rate`, 사이는 선형. `local_education_tax_ratio`(지방교육세/취득세) |
| `brokerage` | `brackets[{upto, rate}]` 오름차순, 마지막 `upto:null`. `vat_ratio`, `sell_rate_simplified`(손익분기 수식용 매도보수) |
| `capital_gains` | `assumption`(출처 시트), `report_note`(손익 모델 주석) |
| `loan` | `*_default`(단지 JSON `outlook.model` 누락 시 기본값), `current{}`(현행 규제 요약, 출처 시트) |
| `regulation`, `tax_reform` | `current{}`, `status`, `source`, `effective_from`. 출처 시트에 자동 기재 |

## 1. market.json
| 키 | 설명 |
|---|---|
| `as_of` | 작성 기준일 `YYYY-MM-DD`. 금리 조회·보고서 표기에 사용. 수작업 항목(periods·groups·volume)과 함께 갱신 |
| `rates_checked_at` | `update_rates.py`가 마지막으로 공식 소스와 대조한 날. 자동 기록 |
| `rates.kr[] / rates.us[]` | `{date, rate(소수, 3%=0.03), note}` 결정일 순. 미국은 목표범위 **상단** |
| `periods[]` | `{label:"2026.06", date:"2026-06-30"}` 반기말. 마지막 부분 반기는 `partial:true`, label에 `*` |
| `groups.high / groups.low` | `{name, members, ppp[], anchors{label:근거}, is_estimate}` ppp는 periods와 같은 길이(만원/3.3㎡, 전용) |
| `volume.rows[]` | `{label, seoul, high_share, low_share}` 맨 앞 1행은 증감률 기준 반기, 이후 periods와 1:1. 부분 반기는 값 없이 label만 |
| `sources_rows[]` | `[구분, 내용]` 출처_가정 시트 상단 |

## 2. apartments/<단지>.json
| 키 | 필수 | 설명 |
|---|---|---|
| `name` | ○ | 보고서·시트 표시 이름 |
| `sheet_prefix` | | 시트 이름 앞부분(기본 name). `_실거래`, `_분석`이 붙음, 31자 제한 |
| `address, households, built` | ○ | 표지·개요 |
| `region.{sido,gu,dong,lawd_cd}` | ○ | lawd_cd = 법정동코드 앞 5자리(시군구). 지역 분석 제목에 dong 사용 |
| `match.{apt_name_contains,umd,jibun}` | | 국토부 데이터 필터 |
| `trades_csv` | ○ | 설정 파일 위치 또는 스킬 루트 기준 상대경로 |
| `trades_note` | | 실거래 시트 안내문 |
| `outlier_threshold` | | 자동 제외 기준(기본 0.25) |
| `kb_asof` | | KB시세 기준월 표기 |
| `types[]` | ○ | `{pyeong, area, households, kb_price, kb_jeonse, jeonse_note}` pyeong=공급평형 정수, area=전용㎡ |
| `compare_pyeong` | | 비교용 두 번째 평형(null 가능) |
| `user_deal.{price,pyeong,note,side}` | ○ | 거래가(만원)·평형·매수/매도/보유/미확인. 거래가 없으면 기준가 사용 후 note에 명시 |
| `user_deal.{dong,ho,floor}` | | 물건 동·호·층(문자열/정수). `floor`가 없으면 호수 앞자리로 계산. 동이 있으면 Ⅱ장에 "같은 동 최고가·평균" 행과 `dong_*` 플레이스홀더가 생김 |
| `user_deal.{resident,first_home,loan_plan,unit_note}` | | 실거주 여부·무주택 여부·대출 계획(만원)·물건 특이사항. 텍스트·checklist·model 작성에 참고 |
| `price_refs.private_price` | | `{label, value, source}` 민간 시세 |
| `price_refs.new_build` | | `{label, value, area, member_ratio, general_label, note, source, ratio_source}` member_ratio가 있으면 일반분양가 환산 행 자동 추가 |
| `sources_rows[]` | | 출처_가정 시트에 추가될 단지별 행 |
| `text.chart_notes[]` | | 요약_차트 하단 `[제목, 본문]` |
| `text.summary[]` | ○ | 분석 시트 Ⅰ장 결론 bullet |
| `text.price_callout` | | Ⅱ장 해석 박스 |
| `text.complex.overview[]` | | `[항목, 내용]` |
| `text.complex.phases[]` | | `[국면, 금리·정책, 가격·거래, 해석]` |
| `text.complex.swot` | | `{S,W,O,T}` |
| `text.episode` | | null이면 장 생략. `{title, year, fact_periods[≤4], callout, trade_events{날짜:이벤트}, factors[[요인,시기,기사요지,함의]], insights[]}` |
| `text.spread` | | `{callout, interpretation[]}` 상관계수 표는 자동 |
| `text.region` | | null이면 장 생략. `{title, callout, timeline[[시기,내용]], pipeline[[사업,규모,단계,관계]], pipeline_source, structure[], outlook[[기간,국면,흐름,영향]]}` |
| `text.outlook.drivers[]` | | `[요인, 현재, ↑/↓/→, 강/중/약, 근거]` 방향에 따라 색 자동 |
| `text.outlook.scenarios[]` | ○ | `[이름, 확률, 1년 변동, 3년 누적, 전제]` 확률 합 1 |
| `text.outlook.model` | | `ltv, loan_cap, rate, etc_cost, holding_tax, deposit_rate, housing_saving` + 각 `_note` |
| `text.outlook.{drivers_callout, model_callout, checklist[], monitoring[[지표,현재,경고,긍정]]}` | | |
| `text.sources[]` | | Ⅷ장 출처 bullet |

## 3. 실거래 CSV (UTF-8)
`date,pyeong,area,price,floor,dong,include,note`
- date `YYYY-MM-DD`, price 만원 정수, pyeong 비우면 area로 자동 매핑
- include 비우면 자동 판정(동일평형 ±6개월 중앙값 대비 threshold 이상 저가 → 0), 0/1로 쓰면 그대로 사용
- dong은 확인된 경우만(예: 101)

## 4. 플레이스홀더 (`<<키>>`)
`python scripts/metrics.py --apt ...`의 `placeholders`가 최신 목록이다. 주요 키:
`apt_name, focus_label, user_price, user_price_short, user_ppp, focus_max_price, focus_max_date, focus_max_floor, user_vs_max, kb_focus, user_vs_kb, jeonse_focus, jeonse_ratio, private_price, user_vs_private, newbuild_member, newbuild_general, user_vs_newbuild_general, compare_label, compare_max_price, compare_max_ppp, user_ppp_vs_compare, focus_max_drawdown, focus_drawdown_desc, kr_rate_now, us_rate_now, spread_now, corr_spread_volume, corr_krrate_volume, corr_dkr_dvolume, exp_1y, exp_3y, breakeven_3y, breakeven_3y_opp, loan_amount, cash_needed, acq_cost`

### 물건(동·호) 관련 플레이스홀더
| 키 | 값 예 | 비고 |
|---|---|---|
| `user_unit` | 101동 1102호 | 동·호 없으면 "동·호 미입력" |
| `user_dong`, `user_ho`, `user_floor` | 101동 / 1102호 / 11층 | 각각 없으면 빈 문자열 |
| `user_side` | 매수 | `user_deal.side` |
| `dong_n` | 19 | 같은 동·같은 평형 실거래 건수(동 표기 거래만) |
| `dong_max_price`, `dong_max_date`, `dong_max_floor` | 7억 4,000만원 / 2026-07-01 / 9층 | 같은 동 최고가. 거래 없으면 키 없음 |
| `dong_avg_price`, `user_vs_dong_max` | 6억 4,929만원 / +13.5% | 같은 동 평균·최고가 대비 |
