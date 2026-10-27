# 나라장터·D2B 입찰공고 자동 수집기

파일 역할: 설치, 실행, 설정과 개발 방법을 안내하는 사용자·개발자 문서입니다.

마지막 업데이트: 2026-10-23

작업자: SMI 장지수

하나의 `main.py` 실행으로 **나라장터(G2B) 학술·연구용역 입찰공고**와 **방위사업청 D2B 국내 경쟁입찰공고**를 별도로 수집합니다. 결과 파일은 같은 컬럼 구조의 Excel 두 개로 저장됩니다. Excel 서식 저장을 위해 **openpyxl**을 사용합니다. 처음 실행하기 전에 이 폴더에서 `python -m pip install -r requirements.txt`를 실행하세요.

## 1. 폴더 구조

```text
bid_notice_collector/
├── main.py               # 두 API 순차 실행 (진입점)
├── config.py             # .env, settings.json, KST 날짜 관리
├── http_client.py        # 통신 재시도, 전체 페이지 수집
├── exporter.py           # 키워드 제외 및 Excel 저장
├── collectors/
│   ├── __init__.py
│   ├── nara.py            # 나라장터 용역 공고 API 및 필드 매핑
│   └── d2b.py             # 방위사업청 D2B 공고 API 및 필드 매핑
├── settings.json         # 검색조건·키워드·표시명 설정
├── .env.example          # 인증키 설정 템플릿 (실제 키 없음)
├── .gitignore            # 인증키와 결과물 Git 추적 제외
├── tests/
│   └── test_collector.py # 서버 호출 없는 자동 테스트
└── output/               # 생성된 Excel 저장 위치
```

## 2. 실행 준비

- **Python 3.9 이상** (권장: Python 3.11 이상)
- 공공데이터포털에서 해당 API 서비스의 **활용 신청 및 사용 승인**
- 각 서비스에 사용할 **일반 인증키(Decoding 형태)**. 인코딩된 키를 입력해도 내부에서 한 번 디코딩합니다.

### ① 환경변수 파일 만들기

프로젝트 폴더에서 `.env.example`을 복사하여 **`.env`** 파일을 만들고 인증키를 입력하세요.

```dotenv
NARA_API_KEY=나라장터용_인증키
D2B_API_KEY=D2B용_인증키
```

두 API가 동일한 공공데이터포털 인증키로 승인되었다면 다음처럼 공통 키 하나로 설정해도 됩니다.

```dotenv
API_KEY=공통_인증키
```

환경변수(운영체제/Codespaces Secrets 등)에 동일한 이름을 설정하면 `.env`보다 우선 적용됩니다. **`.env`를 메일/메신저로 전송하거나 Git에 업로드하지 마세요.** ZIP 배포본에는 실제 인증키가 없습니다. `.gitignore`는 Git을 이용한 실수 예방용이며 ZIP/수동 복사 시 비밀정보 유출을 막아주지는 않습니다.

- Windows 명령 프롬프트: `copy .env.example .env`
- Windows PowerShell: `Copy-Item .env.example .env`
- macOS/Linux/Codespaces: `cp .env.example .env`

### ② 프로그램 실행

```bash
python main.py
```

Windows에서 `python` 명령이 작동하지 않으면 `py main.py`를 시도할 수 있습니다.

기본 설정에서는 **실행 시점의 한국시간(KST) 기준 어제부터 오늘까지 (양쪽 날짜 포함)** 조회합니다. 시간별 자르기는 하지 않습니다.

- 예: 2026-10-16 실행 → 2026-10-15 ~ 2026-10-16 공고 조회
- 나라장터 요청은 시작일 `0000`, 종료일 `2359`로 전송합니다.
- D2B 요청은 시작일·종료일을 `YYYYMMDD` 형식으로 전송합니다.
- 두 API에 적용되는 업종/마감 조건은 기존과 같습니다.

**원하는 날짜 범위로 변경하려면 `settings.json`의 `date_range`를 수정합니다.**

```json
"date_range": {
  "start_date": "2026-10-12",
  "end_date": "2026-10-14"
}
```

이 경우 **10월 12일, 13일, 14일의 공고**를 모두 조회합니다. 기본 자동 모드로 되돌리려면 두 값을 모두 `null`로 바꾸세요.

```json
"date_range": {
  "start_date": null,
  "end_date": null
}
```

`settings.json` 수정 없이 일회성으로 기간만 바꿀 수도 있습니다. 아래 CLI 날짜는 설정 파일보다 우선합니다.

```bash
python main.py --start-date 2026-10-12 --end-date 2026-10-14
```

기존의 하루 조회 옵션도 호환을 위해 유지합니다.

```bash
python main.py --date 2026-10-15
```

두 날짜 중 하나만 입력하거나, 잘못된 날짜/역순 날짜를 지정하면 실행 전에 설정 오류가 발생합니다.

### ③ 결과 확인

```text
output/
├── 나라장터_20261015_20261016.xlsx
└── D2B_20261015_20261016.xlsx
```

두 파일 모두 다음 **10개 컬럼을 같은 순서**로 갖습니다.

`공고위치, 공고번호, 공고구분, 공고일자, 공고명, 공고기관(발주기관), 수요기관, 추정가격(원), 입찰참가등록 마감일시, 상세보기 주소`

- 결과는 **Excel 통합 문서(.xlsx)**로 저장됩니다.
- 공고가 0건이라도 헤더만 있는 Excel를 생성합니다.
- 한 API만 실패하면 다른 API는 계속 처리합니다. 실패한 API의 기존 결과 파일을 덮어쓰지 않습니다. 프로세스 종료코드가 1이면 하나 이상 실패한 것입니다.
- 같은 날짜 범위로 다시 실행하면 해당 범위의 결과 Excel를 최신 결과로 **덮어씁니다**.

## 3. 검색 및 제외 키워드 수정

`settings.json`을 메모장 또는 VS Code로 열어서 수정합니다.

```json
{
  "date_range": {
    "start_date": null,
    "end_date": null
  },
  "page_size": 100,
  "output_directory": "output",
  "exclude_keywords": ["청소", "경비", "보험"],
  "keyword_search_columns": ["공고명"],
  "nara": {
    "display_name": "나라장터",
    "industry_code": "1169",
    "exclude_closed": "Y"
  },
  "d2b": {
    "display_name": "조달청"
  }
}
```

> 위 코드는 **키워드 설정 예시**입니다. 기본 제공 `settings.json`의 `exclude_keywords`는 `[]`(제외 없음)이므로 반드시 원하는 키워드를 입력하세요.

- `exclude_keywords`: **어느 하나라도 포함되면 붉은색으로 표시(OR)**, 영문 대소문자 무시. 공고를 삭제하지 않습니다. UI와 CLI가 같은 설정을 사용하며, `[]`이면 표시하지 않습니다.
- `keyword_search_columns`: 키워드를 검사할 출력 컬럼. 기본은 **`공고명`만** 검사합니다. 기관명에도 적용하려면 `["공고명", "공고기관(발주기관)", "수요기관"]`로 변경합니다.
- `date_range.start_date`, `date_range.end_date`: 둘 다 `null`이면 **어제~오늘 자동 조회**, 둘 다 `YYYY-MM-DD`로 입력하면 **지정한 시작일~종료일 조회**. 양 끝 날짜를 모두 포함합니다.
- `page_size`: API 호출당 최대 결과 수 (1~999, 기본 100).
- `nara.industry_code`: 업종코드 `1169`(학술·연구용역).
- `nara.exclude_closed`: `Y`면 나라장터 **입찰마감제외** API 조건 사용, `N`이면 사용하지 않음.
- `d2b.display_name`: 요청하신 출력 라벨 `조달청`입니다. **실제 이 두 번째 API는 조달청이 아니라 방위사업청 D2B**입니다. 오해 방지를 원하면 `"방위사업청(D2B)"`로 바꾸세요.

**주의**: `exclude_keywords`는 API 요청조건이 아니라 **수집된 공고를 붉은색으로 표시하는 로컬 조건**입니다. API 호출량은 줄이지 않으며 공고도 삭제하지 않습니다. 공고 본문·첨부파일을 다운로드해 검사하지 않으며, 기본적으로 **공고명**에서만 검색합니다.

## 4. 데이터 출처와 빈칸 설명

| Excel 컬럼 | 나라장터 API | D2B API |
|---|---|---|
| 공고위치 | 설정 표시명 `나라장터` | 설정 표시명 `조달청` (실제 출처: 방위사업청) |
| 공고번호 | `bidNtceNo`-`bidNtceOrd` | `g2bPblancNo`-`g2bPblancOdr` (없으면 `pblancNo`-`pblancOdr`) |
| 공고구분 | `ntceKindNm` | `pblancSe` |
| 공고일자 | `bidNtceDt` (일시 포함) | `pblancDate` |
| 공고명 | `bidNtceNm` | `bidNm` |
| 공고기관(발주기관) | `ntceInsttNm` | `ornt` |
| 수요기관 | `dminsttNm` | **미제공 → 빈칸** |
| 추정가격(원) | `presmptPrce` | `bsicExpt` (기초예가) |
| 입찰참가등록 마감일시 | `bidQlfctRgstDt` | `bidPartcptRegistClosDt` |
| 상세보기 주소 | `bidNtceDtlUrl` | **미제공 → 빈칸** |

**D2B의 `bsicExpt`는 기초예가입니다.** 공통 출력 컬럼에 별도의 기초예가 열이 없어 `추정가격(원)` 컬럼에 표시하며, API 응답에 값이 없으면 빈칸으로 둡니다. D2B 목록 API 응답에 수요기관/상세 링크도 없어 임의 URL을 만들지 않습니다.

나라장터는 **용역 검색조건 전용 API** (`getBidPblancListInfoServcPPSSrch`)로 게시일시/업종코드/입찰마감제외를 요청합니다. D2B는 **국내 경쟁입찰 목록 API** (`getDmstcCmpetBidPblancList`)로 공고일자를 요청한 뒤, 응답의 `busiDivs`가 `용역`인 공고만 저장합니다. **D2B에는 나라장터의 1169 업종 필터나 입찰마감제외 조건이 자동 적용되지 않습니다.**

나라장터 `bidClseExcpYn=Y` 결과는 실행시점의 마감 상황에 따라 바뀔 수 있습니다. 공고일자에 대한 과거 스냅샷이 필요하다면 해당 옵션을 `N`으로 바꿔 조회하세요.

## 5. 코드 수정 / 인수인계 가이드

| 변경 목적 | 수정 파일 |
|---|---|
| 인증키 변경 | `.env` (코드 수정 불필요) |
| 키워드 추가/삭제 | `settings.json` → `exclude_keywords` |
| 키워드 검사 대상 변경 | `settings.json` → `keyword_search_columns` |
| 조회 시작일·종료일 변경 | `settings.json` → `date_range` (`null`/`null` = 어제~오늘) |
| 나라장터 검색 업종/마감 제외 변경 | `settings.json` → `nara` |
| 표시할 기관명 변경 | `settings.json` → `display_name` |
| Excel 컬럼 이름/순서 | `config.py` → `COLUMNS`, 각 `collectors/*.py` → `convert` |
| API URL/요청 파라미터 | `collectors/nara.py` 또는 `collectors/d2b.py` |
| Excel 출력 규칙 | `exporter.py` |
| 전체 실행 순서 / 파일명 | `main.py` |

### 테스트

실제 API 호출 없이 필드 매핑과 Excel 출력 등을 확인합니다.

```bash
python -m unittest discover -s tests -v
```

실제 운영 전에는 두 API 활용신청 상태를 확인하고 **인증키를 넣어 한 번 직접 실행**한 다음 나라장터/방위사업청 사이트에서 표본 공고를 대조해야 합니다. API 상세 스키마나 응답 방식이 바뀌면 해당 수집기만 수정하면 됩니다.

## 6. 개발 환경

- **GitHub Codespaces**: 프로젝트를 저장소에 넣고 터미널에서 실행합니다. `.env`는 저장소에 커밋하지 마세요. 권장: Codespaces Secrets로 인증키를 환경변수로 설정.
- **Google Colab**: ZIP을 업로드·압축 해제하고 `main.py`를 실행합니다. Colab 런타임이 초기화되면 저장 파일이 사라질 수 있어 결과물을 별도 저장해야 합니다.
- **일반 Python**: 네트워크 접근이 가능한 Python 3.9+에서 `python main.py`로 실행합니다. PC에 Python 자체가 없다면 브라우저의 Codespaces/Colab을 이용하세요.

※ ZIP 배포본에는 계정·비밀키 및 수집된 원본 공고 데이터가 포함되지 않습니다.


## 포함 키워드 검색 (`include_keywords`)

`settings.json`에서 다음과 같이 설정합니다. 기본값 `[]`은 포함 키워드 제한 없이 전체 공고를 대상으로 합니다.

```json
"include_keywords": ["연구", "조사", "분석"],
"exclude_keywords": ["청소"],
"keyword_search_columns": ["공고명"]
```

- 포함 키워드는 **OR 조건**으로 적용합니다. 예: 공고명에 `연구`, `조사`, `분석` 중 하나라도 포함되면 통과합니다.
- 이후 기존 제외 키워드(하나라도 일치하면 붉은색 표시)가 적용됩니다.
- 비교는 대소문자를 구분하지 않습니다. API 호출 후 결과에 대해 필터링하며, API 요청의 서버측 검색 파라미터로 전달되지는 않습니다.
- `keyword_search_columns`에 지정한 출력 컬럼을 포함 조건과 붉은색 표시 조건 모두에 사용합니다.
- 향후 웹/EXE UI를 제작하면 UI에서 입력한 포함 키워드를 이 설정에 대응시키거나 실행 시 인자로 전달하도록 연결할 수 있습니다.


## 공고명 노란색 강조

`main.py`는 `output/`에 나라장터와 D2B 결과를 각각 `.xlsx` 파일로 저장합니다.
`settings.json`의 `highlight_keywords`에 지정한 단어가 공고명에 하나라도 포함되면
해당 **공고명 셀 전체의 배경을 노란색**으로 표시합니다.

```json
"highlight_keywords": ["사이버", "국방", "지휘통제", "AI", "인공지능", "5G", "특화망", "네트워크"]
```

영문 대소문자를 구분하지 않는 부분 문자열 검색입니다(`ai`, `5g`도 일치).
강조 조건에 맞지 않는 공고도 기존 포함/제외 필터를 통과하면 그대로 저장됩니다.
강조를 모두 끄려면 `highlight_keywords`를 `[]`, `highlight_rules`를 `{}`로 설정하세요. 기존 CSV 파일은 자동 변환하거나 삭제하지 않습니다.


## 참여 검토 공고를 넓게 강조하는 규칙

누락을 줄이기 위해 공고명에서 **관심 분야 또는 과업 성격 중 하나만 일치해도**
노란색으로 표시합니다. `settings.json`의 `highlight_rules`에서 분류별 표현을 수정할 수 있습니다.
국방·안보, AI·디지털·통신, 항공우주·무인·첨단기술과 연구·기획·전략,
조사·분석·평가, 시험·검증·번역을 포함합니다. 기존 `highlight_keywords`도 함께 적용됩니다.

비교 시 영문 대소문자, 공백, 전각/반각 차이를 정규화합니다. 부분 문자열 검색이므로
관련 없는 공고도 강조될 수 있습니다. 운영·유지보수·구매 등의 표현으로 자동 제외하지 않습니다.
이는 제목에 기반한 규칙이며 AI 의미 분류나 참여 자격 판정은 아닙니다.

Excel의 기존 10개 열 뒤에 **강조 근거** 열이 추가됩니다.
예: `관심 분야: 국방·안보: 국방; 과업 성격: 연구·기획·전략: 연구`.
노란색은 사람의 검토 후보이며, 강조하지 않은 공고도 저장합니다.
기관명만으로 강조하지 않으므로 D2B의 모든 공고가 군 기관이라는 이유로 강조되지는 않습니다.

`include_keywords`는 저장할 행을 필터링하고, `exclude_keywords`는 해당 행을 삭제하지 않고 붉은색으로 표시합니다.
누락 방지가 목적이면 두 목록을 기본값 `[]`으로 유지하세요. API 자체의 조회기간·업종·마감 조건은
그대로 적용되므로 강조 규칙이 조회 대상 범위를 넓히지는 않습니다.


## 파일별 역할과 관리 정보

아래 파일의 마지막 업데이트는 **2026-10-23**, 작업자는 **SMI 장지수**입니다. 이번 업데이트는 주석·설명 보강입니다.

| 파일 | 역할 |
|---|---|
| `main.py` | 프로그램 진입점. 날짜·인증키를 읽고 두 API 수집, 키워드 필터링, Excel 저장을 순서대로 실행합니다. |
| `config.py` | 공통 출력 컬럼, 인증키 로딩, 설정 검증과 한국시간(KST) 기준 조회기간 계산을 담당합니다. |
| `http_client.py` | HTTPS 요청의 일시적 오류를 재시도하고 API의 전체 페이지를 수집·검증합니다. |
| `highlighter.py` | 공고명의 관심 분야·과업 표현을 비교하여 노란색 강조 여부와 근거를 생성합니다. |
| `exporter.py` | 포함·제외 키워드 필터링, 강조 공고 우선 정렬, Excel 서식 저장 및 CSV 저장을 담당합니다. |
| `collectors/nara.py` | 나라장터 용역 공고를 날짜·업종·마감 조건으로 조회하고 JSON 응답을 공통 컬럼으로 변환합니다. |
| `collectors/d2b.py` | D2B 국내 경쟁입찰 공고를 날짜 범위로 조회하고 XML 응답을 공통 컬럼으로 변환합니다. |
| `__init__.py` | bid_notice_collector 디렉터리를 Python 패키지로 표시합니다. 실행 진입점은 main.py입니다. |
| `collectors/__init__.py` | 출처별 API 수집기를 모은 collectors 디렉터리를 Python 패키지로 표시합니다. |
| `tests/test_collector.py` | 실제 API 호출 없이 날짜, 설정, 응답 매핑, 필터, 강조·서식과 파일 저장 흐름을 검증합니다. |
| `requirements.txt` | 프로젝트 실행에 필요한 외부 Python 패키지와 버전을 지정합니다. |
| `.gitignore` | 인증키 파일, Python 캐시와 수집 결과가 Git에 포함되지 않도록 지정합니다. |
| `.env.example` | 나라장터·D2B 인증키 설정 예시입니다. .env로 복사하거나 환경변수로 설정합니다. |
| `output/.gitkeep` | 결과가 없는 상태에서도 Git이 output 디렉터리를 유지하도록 하는 자리표시자입니다. |
| `settings.json` | 조회조건·필터·강조 설정. JSON은 주석을 지원하지 않아 `_documentation`에 설명·날짜·작업자를 기록합니다. |
| `README.md` | 설치·사용·개발 안내와 파일별 역할을 설명합니다. |

현재 `settings.json`에는 사용자가 지정한 제외 키워드가 있습니다. 해당 공고는 저장하되 붉은색으로 표시합니다. 생성된 Excel·CSV와 인증키가 들어 있는 로컬 `.env`는 관리 주석 대상이 아닙니다.


### main에서 추가한 UI 파일

아래 파일은 `origin/main`에서 가져왔습니다. 마지막 업데이트: **2026-10-23**, 작업자: **SMI 장지수**.

| 파일 | 역할 |
|---|---|
| `.gitignore` | 저장소 루트의 생성 결과 파일을 Git 추적에서 제외합니다. |
| `README_UI.md` | 웹·Windows UI의 설치, 실행과 배포 방법을 안내합니다. |
| `build_exe.bat` | Windows에서 의존성을 설치하고 Tkinter 앱을 단일 EXE로 빌드합니다. |
| `desktop_app.py` | Tkinter 화면에서 조회조건을 입력하고 백그라운드 수집 결과를 표시합니다. |
| `requirements-build.txt` | Windows EXE 빌드용 PyInstaller와 공통 실행 의존성을 지정합니다. |
| `requirements-web.txt` | 웹 UI용 Streamlit과 공통 실행 의존성을 지정합니다. |
| `tests/test_ui.py` | 페이지 이벤트와 UI 공통 실행기의 키워드 필터·Excel 저장을 모의 응답으로 검증합니다. |
| `ui_runner.py` | 웹·데스크톱 UI가 공유하는 수집 실행기입니다. 페이지 이벤트 전달, 필터링과 Excel 저장을 담당합니다. |
| `web_app.py` | Streamlit 화면에서 공고를 조회하고 진행 상태, 표와 Excel 다운로드를 제공합니다. |

`main.py`와 웹·데스크톱 UI는 동일한 저장기로 노란색 강조·근거가 포함된 Excel을 저장합니다. UI의 페이지 이벤트 지원과 EXE 설정 경로 처리를 연결했습니다. 기존 파일의 다른 main 변경(입찰구분 컬럼, D2B 용역 필터)은 이번 파일 추가에 포함하지 않았습니다.


### UI의 강조·정렬과 Excel 저장

웹·데스크톱 모두 `settings.json`의 `highlight_keywords`와 `highlight_rules`를 사용합니다.
화면에 강조 근거를 표시하고 강조 공고를 먼저 모읍니다. 웹은 공고명 셀, 데스크톱은
Treeview 제약으로 행 전체에 노란색을 표시합니다. 두 UI가 저장하는 Excel은 동일하게
**공고명 셀만 노란색**, 강조 공고 우선 안정 정렬, 강조 근거 열, 회색 헤더, 검은 테두리,
첫 행 고정, 자동필터, 열 너비를 적용합니다. 웹 다운로드도 `.xlsx`입니다.

화면은 수집 중인 임시 결과를 표시하며 API의 전체 수집 검증이 끝난 뒤에만 파일을 저장합니다.
포함 키워드는 UI 입력값을 사용하고 제외 키워드는 설정 파일을 따릅니다.
