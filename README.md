# apt-saramara — 아파트 살까말까 (Claude Code 스킬)

"○○아파트 101동 1102호 8억에 사는 거 어때?" 한 줄이면 국토부 실거래·전세, 한·미 기준금리, 대출·세제 정책, 지역 개발 이슈를 조사해
**같은 구조의 엑셀 분석 보고서**(요약 차트 · 단지 종합 분석 · 같은 동 실거래 비교 · 금리차 상관 · 과거 이상 구간 분석 · 지역 전망 · 가격 시나리오 · 3년 손익 모델 · 출처)를 만들어 주는 Claude Code 스킬입니다.

## 1. 설치
**사전 준비**: [Claude Code](https://claude.com/claude-code) 설치 + Python 3.10 이상

```bash
# 프로젝트 단위(권장): 스킬을 쓸 프로젝트 폴더에서
git clone https://github.com/cibstival/apt-saramara .claude/skills/apt-saramara

# 또는 모든 프로젝트에서 쓰려면 (Windows는 %USERPROFILE%\.claude\skills\)
git clone https://github.com/cibstival/apt-saramara ~/.claude/skills/apt-saramara

pip install openpyxl requests   # 필수
# 선택: LibreOffice (수식 오류 자동 검사·PDF 미리보기용)
```
설치 후 해당 폴더에서 Claude Code를 (재)시작하면 `/apt-saramara` 스킬이 자동으로 잡힙니다.
git이 없으면 GitHub 페이지의 **Code → Download ZIP**을 받아 위 경로에 풀어도 됩니다(폴더 바로 아래에 `SKILL.md`가 오도록).

업데이트: 설치 폴더에서 `git pull` (스킬이 실행 때마다 스스로 `git pull --ff-only`를 시도합니다)

## 2. 사용
### 가장 간단한 방법
```
/apt-saramara
```
단지명 · 거래가 · 동·호 · 매수/매도/보유·실거주 여부를 **한 번에** 묻고 바로 시작합니다. 모르는 항목은 비워두면 됩니다.

### 한 줄로 바로 시작
```
/apt-saramara 홍은현대아파트 101동 1102호 8.4억 매수
```
```
홍은현대아파트 101동 1102호 8억 4천에 사는 거 어떨까?
```
```
잠실엘스 33평 가지고 있는데 지금 27억에 팔아도 될까? 실거주 중이야.
```
```
은마아파트 분석 보고서 만들어줘.
```
- **필수**: 단지명, 거래가. 빠지면 그것만 되묻습니다. "그냥 만들어줘"라고 하면 최근 최고 실거래를 기준가로 씁니다.
- **선택**: 동·호(→층 자동), 평형, 매수/매도/보유, 실거주 여부, 무주택 여부, 대출 계획. 동을 알려주면 보고서에 **같은 동 실거래 최고가·평균 대비** 비교가 들어갑니다.
- 동명이단지(○○현대 등)는 구·동을 같이 적어 주세요.

### 실행되면 벌어지는 일
1. **최신화 게이트** — `git pull`, 한·미 기준금리 자동 확인, 정책 기준 확인일 점검(30일 초과 시 웹검색으로 재확인)
2. **실거래·전세 수집** — 국토부 실거래가 공개시스템에서 키 없이 직접 다운로드(동·층 포함)
3. **조사** — 단지 정보, 지역 정비사업·교통, 시장 동향, 주담대 금리 등을 웹검색해 설정 JSON 작성
4. **빌드·검증** — 엑셀 생성 후 미치환 항목·수식 오류 검사, **프로젝트 폴더의 `apt_saramara/`** 에 저장

5~10분 정도 걸리고, 마지막에 핵심 결론 3~5줄과 추정치·미확인 항목을 알려줍니다.

## 3. 자동 최신화 구조
| 항목 | 어디에 | 어떻게 갱신 |
|---|---|---|
| 한·미 기준금리 | `config/market.json` | GitHub Actions가 매주 월요일 한국은행·연준 공식 페이지에서 읽어 커밋. 스킬 실행 시에도 `scripts/update_rates.py`로 재확인 |
| 취득세·중개보수·양도세 가정·LTV·규제지역 | `config/policy.json` | 확인일(`checked_at`)이 30일 넘으면 `scripts/check_policy.py`가 경고 → 스킬이 정책 변경을 검색해 JSON과 변경이력을 갱신. 코드에는 세율 상수가 없음 |
| 주담대 금리·LTV·규제 여부 (단지 위치 기준) | 단지별 JSON `outlook.model` | 실행 때마다 웹검색으로 새로 채움 |

적용된 정책 기준과 확인일은 보고서 '출처_가정' 시트에 자동으로 들어갑니다.
선택: 저장소 Secrets에 `ECOS_API_KEY`(한국은행 ECOS)를 넣으면 공개 페이지 파싱 대신 API를 씁니다.

## 4. 실거래 데이터 경로
1. **기본**: `scripts/fetch_rtms_direct.py` — rt.molit.go.kr에서 연도별 CSV를 직접 받음. 키 불필요. `--rent`로 전월세도 받아 전세 시세로 사용
2. 예비: `MOLIT_API_KEY`가 있으면 `scripts/fetch_molit_trades.py` (data.go.kr "아파트 매매 실거래가 상세 자료" 활용신청 후 환경변수 설정)
3. 예비: rt.molit.go.kr에서 직접 내려받은 파일을 `scripts/import_rtms_csv.py`로 변환

## 5. 직접 실행
```bash
cd .claude/skills/apt-saramara
python scripts/check_policy.py                                              # 정책 기준 확인일 점검
python scripts/update_rates.py --dry-run                                    # 기준금리 소스 대조
python scripts/fetch_rtms_direct.py --apt config/apartments/hongeun_hyundai.json   # 실거래 다운로드
python scripts/metrics.py --apt config/apartments/hongeun_hyundai.json      # 문장에 쓸 수치·플레이스홀더
python scripts/build_report.py --apt config/apartments/hongeun_hyundai.json   # → <프로젝트 루트>/apt_saramara/홍은현대아파트_분석보고서_<오늘>.xlsx
python scripts/verify_xlsx.py ../../../apt_saramara/홍은현대아파트_분석보고서_<오늘>.xlsx   # LibreOffice 있을 때 수식 검사
```

## 6. 무엇이 고정이고 무엇이 바뀌나
| 고정 (코드) | 단지·물건마다 바뀜 (설정·데이터) |
|---|---|
| 시트 7개 구성, 목차·카드·섹션 디자인 | 단지명·주소·평형·세대수·전세 시세 |
| 평단가·거래량·금리 차트, 상관계수·구간표 수식 | 실거래 CSV (`data/`) |
| 같은 동 비교, 시나리오 표, 3년 손익 모델·손익분기 수식 | 물건 정보(동·호·거래가·매수/매도), 요약·해석 문장, SWOT |
| 세율·요율 수식 (값은 `policy.json`에서) | 과거 이상 구간(에피소드) 기사 분석 |
| 플레이스홀더 치환(`<<user_vs_max>>`, `<<dong_max_price>>` 등) | 지역 개발 분석·전망, 드라이버, 체크리스트 |

예시 설정: `config/apartments/changsin_ssangyong1.json`, `hongeun_hyundai.json`. 필드 설명은 `references/config_schema.md`, 조사 방법은 `references/research_guide.md`.

## 7. 주의
- 구 그룹 평단가·거래량, 시나리오 확률, 보유세는 추정치입니다(엑셀 노란 칸·출처 시트에 표시).
- 국토부 공개시스템·API 구조가 바뀌면 다운로드가 실패할 수 있습니다. 그때는 4절의 예비 경로를 씁니다.
- 보고서는 정보 제공용이며 투자·세무·법률 자문이 아닙니다.
