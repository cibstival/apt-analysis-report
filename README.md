# apt-analysis-report (Claude Code 스킬)

아파트 단지명만 바꾸면 같은 구조의 엑셀 분석 보고서를 다시 만들어 주는 Claude Code 스킬입니다.
창신쌍용1단지 보고서(요약 차트, 단지 종합 분석, 데이터, 거래량, 실거래, 금리 이력, 출처)가 기본 형태입니다.

## 1. 설치
**사전 준비**: [Claude Code](https://claude.com/claude-code) 설치 + Python 3.10 이상

```bash
# 프로젝트 단위(권장): 스킬을 쓸 프로젝트 폴더에서
git clone https://github.com/cibstival/apt-analysis-report .claude/skills/apt-analysis-report

# 또는 모든 프로젝트에서 쓰려면 (Windows는 %USERPROFILE%\.claude\skills\)
git clone https://github.com/cibstival/apt-analysis-report ~/.claude/skills/apt-analysis-report

pip install openpyxl          # 필수
# 선택: LibreOffice (수식 오류 자동 검사·PDF 미리보기용)
```
설치 후 해당 폴더에서 Claude Code를 (재)시작하면 `/apt-analysis-report` 스킬이 자동으로 잡힙니다.
git이 없으면 GitHub 페이지의 **Code → Download ZIP**을 받아 위 경로에 풀어도 됩니다(폴더 바로 아래에 `SKILL.md`가 오도록).

업데이트: 설치 폴더에서 `git pull` (스킬이 실행 때마다 스스로 `git pull --ff-only`를 시도합니다)

### 정책·금리 자동 최신화
- **기준금리**: GitHub Actions가 매주 월요일 FRED·한국은행에서 한·미 기준금리를 읽어 `config/market.json`을 갱신·커밋합니다. 스킬도 실행 시 `scripts/update_rates.py`를 돌려 그 자리에서 다시 확인합니다.
- **세제·대출·규제**: `config/policy.json`에 취득세·중개보수·양도세 가정·LTV·규제지역이 확인일과 함께 들어 있습니다. 확인일이 30일을 넘으면 스킬이 실행 전에 정책 변경을 검색해 확인하고(`scripts/check_policy.py`), 바뀐 값은 JSON과 변경이력에 기록합니다. 코드에는 세율 상수가 없습니다.
- 보고서 '출처_가정' 시트에 적용된 정책 기준과 확인일이 자동으로 들어갑니다.
- 선택: 저장소 Secrets에 `ECOS_API_KEY`(한국은행 ECOS)를 넣으면 공개 페이지 파싱 대신 API를 씁니다.

### 국토부 실거래 API 키 (선택이지만 권장)
1. data.go.kr 에서 "국토교통부_아파트 매매 실거래가 상세 자료" 활용신청
2. 발급된 일반 인증키(Decoding)를 환경변수로 설정
```bash
export MOLIT_API_KEY="발급받은키"                      # macOS/Linux
[Environment]::SetEnvironmentVariable("MOLIT_API_KEY","발급받은키","User")   # Windows PowerShell
```
키가 없으면 rt.molit.go.kr 에서 내려받은 파일을 넘겨주거나, Claude가 공개 페이지의 실거래 목록을 옮겨 적습니다.

## 2. 사용 (Claude Code에 이렇게 말하면 됩니다)
```
은마아파트 분석 보고서 만들어줘. 31평 26억에 매수 검토 중이야.
```
```
창신쌍용1단지 보고서 형식 그대로, 성산시영 전용 50㎡로 바꿔서 뽑아줘.
```
```
market.json 최신 금리로 업데이트하고 창신쌍용1단지 보고서 다시 빌드해줘.
```
Claude가 SKILL.md 순서대로 ① 시장 데이터 갱신 ② 단지 정보·실거래 수집 ③ 기사 조사 후 설정 작성 ④ 빌드·검증을 진행하고 `output/` 에 엑셀을 저장합니다.

## 3. 직접 실행
```bash
cd .claude/skills/apt-analysis-report
python scripts/build_report.py --apt config/apartments/changsin_ssangyong1.json --out output/창신쌍용1단지.xlsx
python scripts/metrics.py --apt config/apartments/changsin_ssangyong1.json     # 문장에 쓸 수치 확인
python scripts/verify_xlsx.py output/창신쌍용1단지.xlsx                        # 수식 오류 검사
```

## 4. 무엇이 고정이고 무엇이 바뀌나
| 고정 (코드) | 단지마다 바뀜 (설정·데이터) |
|---|---|
| 시트 7개 구성, 목차·카드·섹션 디자인 | 단지명·주소·평형·세대수·KB시세 |
| 평단가·거래량·금리 차트, 상관계수·구간표 수식 | 실거래 CSV |
| 시나리오 표·3년 손익 모델·손익분기 수식 | 요약·해석 문장, 사이클 해석, SWOT |
| 연도별 실거래 요약 수식 | 과거 이상 구간(에피소드) 기사 분석 |
| 플레이스홀더 치환(<<user_vs_max>> 등) | 지역(동) 개발 분석·전망, 드라이버, 체크리스트 |

시장 공통 데이터(한미 금리, 비교 그룹 평단가·거래량)는 `config/market.json` 하나로 모든 단지가 공유합니다.

## 5. 주의
- 구 그룹 평단가·거래량, 시나리오 확률은 추정치입니다(엑셀 노란 칸·출처 시트에 표시).
- 국토부 API 필드명은 개편될 수 있어 첫 실행 때 `--debug`로 확인하도록 되어 있습니다.
- 보고서는 정보 제공용이며 투자·세무·법률 자문이 아닙니다.
