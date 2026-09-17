---
name: apt-analysis-report
description: 아파트 단지명(과 선택적으로 거래가·평형)만 받으면 국토부 실거래·KB시세·한미 기준금리·지역 개발 이슈를 조사해 동일한 구조의 엑셀 분석 보고서(평단가/거래량 추세 차트, 금리차 상관분석, 과거 이상 구간 기사 분석, 지역 심층 분석, 가격 시나리오·3년 손익 모델)를 생성한다. 사용자가 특정 아파트·단지의 시세 분석, 가격 전망, 매수·매도 검토, "○○아파트 보고서/엑셀 만들어줘", "이 단지로 바꿔서 다시 뽑아줘", 창신쌍용 보고서 같은 형식을 다른 단지에 적용하려 할 때 반드시 이 스킬을 사용한다.
---

# 아파트 단지 분석 보고서 생성

코드와 레이아웃은 고정돼 있고, 단지마다 달라지는 것은 **설정 JSON + 실거래 CSV**뿐이다.
따라서 새 단지 요청이 오면 할 일은 ① 데이터를 모으고 ② 조사해서 설정을 채우고 ③ 빌드·검증하는 것이다.
엑셀 구조를 새로 설계하거나 스크립트를 크게 고치지 않는다(버그 수정은 예외).

```
apt-analysis-report/
├── scripts/build_report.py      # 설정 → 엑셀 (시트 7개, 수식·차트 포함)
├── scripts/metrics.py           # 문장 작성용 수치·플레이스홀더 출력
├── scripts/fetch_molit_trades.py# 국토부 API → 실거래 CSV (MOLIT_API_KEY 필요)
├── scripts/import_rtms_csv.py   # rt.molit.go.kr 다운로드 파일 → 실거래 CSV
├── scripts/verify_xlsx.py       # LibreOffice로 수식 오류 검사(없으면 생략)
├── config/market.json           # 시장 공통: 한미 금리, 반기 목록, 비교 그룹 평단가·거래량
├── config/apartments/_template.json
├── config/apartments/changsin_ssangyong1.json  # 완성 예시(항상 참고)
├── data/*.csv                   # 단지별 실거래
└── references/                  # config_schema.md, research_guide.md
```

## 작업 순서

### 0) 입력 확인
- 필수: 단지명. 선택: 거래가(만원)·평형·매수/매도 여부.
- 동명이단지가 흔하다(예: ○○현대). 시·구·동이 없으면 한 번만 물어본다.
- 거래가가 없으면 멈추지 말고 최근 실거래 최고가를 `user_deal.price`로 쓰고 note에 "기준가(사용자 거래가 없음)"라고 적는다.

### 1) 시장 설정 최신화 (`config/market.json`)
`as_of`가 오늘보다 1개월 이상 오래됐으면 먼저 갱신한다. 방법은 `references/research_guide.md`의 '시장 데이터' 절.
- 금리: 한국은행·FOMC 결정을 검색해 `rates` 끝에 추가.
- `periods` 마지막 부분 반기(`partial: true`)의 label·date를 오늘 기준으로 수정하고, 완성된 반기는 일반 행으로 추가.
- `groups.*.ppp`와 `volume.rows`는 periods와 **같은 길이·순서**를 유지한다(volume은 맨 앞 기준 반기 1개가 더 있음). 새 값은 공표치 근거를 anchors/sources_rows에 남긴다.
- 서울 밖 단지면 비교 그룹을 해당 지역에 맞게 바꿀지 사용자에게 한 줄로 확인한다.

### 2) 단지 기본정보 → 설정 파일 생성
`_template.json`을 `config/apartments/<영문이름>.json`으로 복사하고 채운다.
- 공식 단지명, 지번 주소, 법정동, 시군구코드 5자리(`region.lawd_cd`), 세대수, 입주연월, 동 수, 용적률.
- `types`: 공급평형·전용㎡·세대수·KB 매매시세·KB 전세시세. 전용면적은 국토부 데이터와 매칭되므로 소수점까지 정확히.
- `compare_pyeong`: 거래가 많은 다른 대표 평형(없으면 null).
- `match`: 국토부 데이터의 단지명 일부·법정동·지번(동명이단지 걸러내기용).

### 3) 실거래 수집
우선순위: (a) `MOLIT_API_KEY`가 있으면 `python scripts/fetch_molit_trades.py --apt <설정> --from 201701 --to <이번달YYYYMM> --debug`
(첫 호출에 `--debug`로 응답 필드명을 확인하고, 스크립트 FIELD 사전에 없으면 추가) →
(b) 사용자가 rt.molit.go.kr 다운로드 파일을 주면 `import_rtms_csv.py` →
(c) 둘 다 없으면 공개 단지 페이지(아파트위키 등)의 실거래 목록을 CSV로 옮긴다(출처·기준일을 `trades_note`에 기록).

그다음 `python scripts/metrics.py --apt <설정>`을 실행해 `excluded`(자동 제외 거래)와 `recent`를 확인한다.
자동 규칙은 보수적이라, 같은 시기 시세보다 20% 이상 낮은 직거래·가족간 거래로 보이면 CSV `include`를 0으로 직접 바꾸고 note에 이유를 쓴다. 반대로 정상 거래가 제외됐으면 1로 둔다.

### 4) 조사 후 텍스트 채우기
`metrics.py` 출력의 `yearly`, `halves`, `placeholders`를 옆에 두고 쓴다. 섹션별 검색어·출처 기준은 `references/research_guide.md`.
- **숫자는 가능하면 `<<키>>` 플레이스홀더**로 쓴다(예: `<<user_vs_max>>`). 데이터가 바뀌어도 문장이 틀리지 않는다.
- `complex.phases`: 연도별 가격·거래 흐름을 금리 국면과 묶어 6~8개 국면으로.
- `episode`: `halves`에서 가장 설명이 필요한 구간(금리와 거래량이 반대로 움직인 반기, 급락, 거래절벽 등)을 골라 당시 기사로 원인을 분석한다. 마땅한 구간이 없으면 `"episode": null`(섹션 자동 생략).
- `region`: 해당 동/생활권의 정비사업·교통·공급·정책 파이프라인과 단기/중기/장기 전망. 해당 없으면 null.
- `price_refs.new_build`: 인근 신축·분양 비교치를 넣을 때 **조합원분양가/일반분양가/입주 후 시세 중 무엇인지, 어느 시점 추정인지**를 label과 source에 반드시 명시한다(과거에 조합원분양가를 일반분양가로 잘못 표기한 사례가 있음). 조합원가면 `member_ratio`를 기사 근거로 넣는다.
- `outlook.drivers`·`model`: 실행 시점의 주담대 금리, LTV·대출한도, 토지거래허가·규제지역 여부, 세제 변경을 **해당 단지 위치 기준으로 다시 검색**한다. 규제는 자주 바뀐다.
- `checklist`: 매수/매도가 불명확하면 양쪽 관점을 모두 넣는다.

### 5) 빌드·검증
```bash
python scripts/build_report.py --apt config/apartments/<단지>.json --out output/<단지>_분석보고서_<YYYY-MM-DD>.xlsx
python scripts/verify_xlsx.py output/<파일>.xlsx
python scripts/metrics.py --apt config/apartments/<단지>.json   # "unfilled"가 빈 리스트인지 확인
```
- 빌드 로그에 `[경고] 치환되지 않은 플레이스홀더`가 나오면 metrics 키 이름을 고친다.
- 가능하면 LibreOffice로 PDF 렌더링해 차트·표가 깨지지 않는지 눈으로 확인한다.
- 사용자에게는 핵심 결론 3~5줄, 추정치로 들어간 항목, 확인하지 못한 항목(예: 동 정보, KB시세 날짜)을 짧게 보고한다.

## 반드시 지킬 원칙
- **실측 vs 추정 구분**: 금리·국토부 실거래는 실측, 구 그룹 평단가·거래량·시나리오 확률은 추정. 셀 색(노란 칸)·비고·출처 시트에 추정임을 남긴다. 추정치를 실측처럼 쓰지 않는다.
- **동 단위 분석**: 데이터에 동 정보가 실제로 있을 때만 동별 결론을 낸다. 없으면 단지 전체 기준임을 명시한다.
- **출처 표기**: 기사 내용은 요약해 쓰고 "(매체 YYYY.M.D)"로 표기한다. 원문 문장을 길게 옮기지 않는다.
- **투자 조언 아님**: 매수·매도 권유 표현을 쓰지 말고, 판단 근거와 조건(실거주·보유기간·금리)을 제시한다. 세금·대출은 전문가 확인이 필요하다고 남긴다.
- **구조 일관성**: 시트 이름·섹션 순서·수식 구조는 유지한다. 새 분석이 필요하면 설정 텍스트로 추가하고, 레이아웃 변경은 사용자가 요청할 때만 한다.

## 자주 생기는 문제
- 평형 매핑 실패: API 전용면적이 `types.area`와 3㎡ 이상 차이 → 해당 평형을 types에 추가.
- 반기 거래 0건: 차트는 앞뒤 보간, 데이터 시트 비고에 자동 기록됨.
- `MAXIFS`/`MINIFS`는 엑셀 2019+ 필요(구버전 엑셀이면 사용자에게 안내).
- 서울 외 지역: `market.json`의 sources_rows·groups 이름이 서울 기준이므로 함께 수정.
