# 웹 + Windows EXE 사용 안내 (추가 기능)

기존 `main.py` 명령행 수집은 그대로 동작합니다. 두 UI는 `ui_runner.py`의 공통 수집기를 사용합니다.

## 공통 준비
- 프로젝트 루트에 `settings.json`, `.env` 두 파일을 유지합니다. `.env.example`을 복사해 `.env`를 만들고 API 키를 입력하세요.
- 웹/EXE에서 입력한 **포함 키워드**는 쉼표로 구분하며 하나라도 일치하면 포함합니다(OR). 비어 있으면 전체 공고를 표시합니다. 기존 `settings.json`의 **제외 키워드**는 계속 적용합니다. 키워드 판정 대상 컬럼은 `keyword_search_columns`입니다(기본: 공고명).
- 날짜는 한국시간 기준 전일~당일이 기본값이며 UI에서 변경할 수 있습니다(YYYY-MM-DD). 두 날짜 모두 포함됩니다.
- 나라장터는 1169 업종과 `bidClseExcpYn=Y` 검색 조건을 유지합니다. D2B의 상세보기 주소는 개별 페이지가 아닌 사용자가 지정한 목록 주소입니다.
- 페이지 하나가 도착할 때마다 화면에 결과가 추가됩니다. 따라서 같은 페이지의 레코드는 묶여 나타납니다. API 완료 후 각각 `output/나라장터_시작일_종료일.csv`, `output/D2B_시작일_종료일.csv`로 저장합니다.
- API 오류가 발생한 수집원은 파일을 덮어쓰지 않으며, 다른 수집원은 계속 실행합니다.

## 웹 UI (Streamlit)

```sh
python -m pip install -r requirements-web.txt
python -m streamlit run web_app.py
```

웹 서버가 실행되는 컴퓨터에만 Python 및 Streamlit이 필요합니다. 브라우저는 실행된 웹 앱 주소에 접속합니다. 이 버전은 개인/내부망 사용을 전제로 하며 로그인 기능이 없습니다. **공용 인터넷에 바로 노출하지 마세요.**

## Windows EXE (Tkinter)

Windows PC에서 Python 3.10 이상으로 프로젝트 폴더에 들어가 `build_exe.bat`를 실행합니다. 결과물: `dist/BidNoticeCollector.exe`. GUI 자체는 Tkinter(표준 라이브러리)를 사용하며 EXE 빌드에만 PyInstaller를 설치합니다.

**배포할 때는 EXE 외에 `settings.json`과 `.env`를 EXE와 같은 폴더에 배치**합니다. `.env`에는 해당 사용자의 API 키가 있어야 합니다. `output/`은 자동 생성됩니다. 외부 인증키를 EXE 내부에 넣지 않습니다. 윈도우 PC에서 빌드해야 Windows EXE가 생성됩니다. 해당 PC의 방화벽 및 인터넷 연결이 필요합니다.

## 핵심 파일
- `ui_runner.py`: 공통 실행기, 포함/제외 키워드 판정, API별 CSV 저장, 페이지 이벤트
- `web_app.py`: Streamlit 인터페이스
- `desktop_app.py`: Tkinter 인터페이스
- `http_client.py`, `collectors/nara.py`, `collectors/d2b.py`: 페이지 도착 이벤트 지원
- `config.py`: EXE에서 `settings.json`/`.env`를 실행 파일 옆에서 읽도록 경로 처리

## 데이터 정확도 / 검증
현재 테스트는 네트워크 없는 모의 API 응답으로 이루어집니다. 실제 인증키를 사용한 엔드투엔드 호출과 EXE 빌드, 실제 운영 PC 배포는 사용자 환경에서 별도 검증이 필요합니다. 수집 중 서버 데이터가 바뀌면 페이지 결과가 변경될 수 있습니다.
