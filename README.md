# 네이버 카페 모니터링

엑셀에 정리된 네이버 카페 게시글 URL을 순회하며 조회수/댓글수/좋아요 등을 수집하고,
등록한 닉네임이 댓글을 단 게시글이 있으면 댓글 내용을 엑셀에 함께 기록해주는 도구입니다.
로그인 없이 공개 카페 게시글만 대상으로 합니다.

## 1. 새 PC에서 세팅하기

### 1) 필수 프로그램 설치

- **Python 3.10 이상** ([python.org](https://www.python.org/downloads/)에서 설치, 설치 시 "Add python.exe to PATH" 체크)
- **Google Chrome 브라우저** (Selenium이 실제 크롬을 제어하는 방식이라 반드시 설치되어 있어야 함)

> 참고: Windows에 Python을 설치하면 `python` 명령어 대신 `py` 런처만 정상 동작하는 경우가 있습니다.
> 이 경우 아래 모든 `python` 명령어를 `py`로 바꿔서 실행하면 됩니다.

### 2) 프로젝트 받기

```powershell
git clone https://github.com/ljg60042/naver_Cafe_Monitoring.git
cd naver_Cafe_Monitoring
```

### 3) 필요한 패키지 설치

```powershell
pip install -r requirements.txt
```

(설치되는 패키지: `selenium`, `webdriver-manager`, `openpyxl`, `pandas`)

### 4) 실행

```powershell
python naver_cafe_monitor.py
```

- 최초 실행 시 인터넷에 연결되어 있어야 합니다 (`webdriver-manager`가 설치된 Chrome 버전에 맞는
  드라이버를 자동으로 다운로드합니다).

## 2. 사용 방법

1. **네이버 아이디(닉네임) 등록**: 입력창에 추적할 닉네임을 입력하고 "추가" 클릭 (Enter 키도 가능).
   여러 개 등록 가능하며, 목록에서 선택 후 "선택 삭제"로 제거할 수 있습니다.
   등록한 닉네임 목록은 `nicknames.txt`에 자동 저장되어 다음 실행 시에도 유지됩니다.
2. **엑셀 파일 선택**: "엑셀 파일 선택" 버튼으로 카페 게시글 URL이 담긴 엑셀 파일을 선택합니다.
   엑셀에는 `업로드 링크` 컬럼이 있어야 하며, 각 행에 `cafe.naver.com` 또는 `naver.me` URL을 넣어둡니다.
3. **모니터링 시작**: 닉네임을 1개 이상 등록하고 엑셀 파일을 선택하면 버튼이 활성화됩니다.
   시작하면 다른 버튼/입력창은 처리가 끝날 때까지 잠깁니다.
4. 처리가 끝나면 결과를 저장할 위치를 지정합니다. 결과 엑셀에는 카페명, 게시판, 제목, 조회수, 댓글수,
   좋아요 수와 함께 등록된 닉네임이 작성한 댓글이 있으면 **댓글순번 / 댓글닉네임 / 댓글내용** 컬럼에
   기록됩니다. 한 게시글에 일치하는 댓글이 여러 개면 그만큼 행이 나뉘어 표시됩니다.

## 3. exe로 빌드하기

배포용 실행파일이 필요하면 PyInstaller로 빌드합니다.

```powershell
pip install pyinstaller
python -m PyInstaller naver_cafe_monitor.spec
```

- 빌드 결과물은 `dist\naver_cafe_monitor.exe`에 생성됩니다 (단일 exe 파일, 콘솔창 없음).
- exe 파일 하나만 다른 PC로 옮겨도 실행 가능합니다. 단, 받는 PC에도 **Chrome 브라우저**와
  **인터넷 연결**은 필요합니다.
- 서명되지 않은 exe라서 Windows SmartScreen 경고가 뜰 수 있습니다 → "추가 정보" → "실행"으로 진행.
- `nicknames.txt`는 exe가 위치한 폴더를 기준으로 자동 생성/저장됩니다.

## 4. 참고 사항

- `build/`, `dist/`, `__pycache__/`, `debug_iframe.html`, `nicknames.txt`는 `.gitignore`에 등록되어
  있어 깃허브에는 올라가지 않습니다. (빌드 산출물 및 사용자별 로컬 데이터이기 때문)
- `debug_iframe.html`은 게시글 처리 중 디버깅용으로 남기는 페이지 소스 스냅샷입니다. 정상 동작에는
  영향 없습니다.
