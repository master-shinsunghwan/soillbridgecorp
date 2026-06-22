# Workhub 인계서

## 전달 목적

이 압축본은 (주)소일브릿지 업무자동화 프로그램의 현재 작업본을 다른 PC에서 검토하고 이어서 개발할 수 있도록 만든 로컬 실행용 패키지입니다.

## 현재 기준

- 저장소: `https://github.com/master-shinsunghwan/soillbridgecorp`
- 브랜치: `main`
- 앱 동작 검증 기준 커밋: `826c3b8 Fix Workhub lucide icon import`
- 주요 실행 파일: `scripts/workhub_delivery_app.py`
- 로컬 실행 주소: `http://127.0.0.1:8770/`
- 로그인 테스트 계정: `admin / admin1234`

## 포함된 주요 데이터

- `config/workhub.db`: 현재 업무 DB
- `templates/`: 엑셀 출력 양식
- `scripts/`: 업무자동화 서버 및 변환 로직
- `static/`: 화면 스타일 파일
- `tests/`: 회귀 테스트
- `requirements.txt`: 실행에 필요한 Python 패키지 목록

보안상 네이버 메일 비밀번호, 서버 토큰, 개인 PC 전용 비밀키 파일은 압축본에서 제외하는 것을 권장합니다. 다른 PC에서 메일 발송까지 테스트하려면 관리자 화면에서 네이버 메일 정보를 다시 저장해야 합니다.

## 다른 PC에서 실행하는 방법

1. 압축 파일을 원하는 폴더에 풉니다.
2. Python 3.11 이상이 설치되어 있는지 확인합니다.
3. `업무허브 실행.cmd`를 더블클릭합니다.
4. 브라우저가 열리면 `http://127.0.0.1:8770/` 또는 `/login`으로 접속합니다.
5. `admin / admin1234`로 로그인합니다.

필요 패키지가 없으면 실행 스크립트가 `requirements.txt` 기준으로 설치를 시도합니다.

## 우선 확인할 화면

- 로그인 후 첫 진입 화면
- 매출현황 및 관리
- 매출표 업로드
- 통합관리대장 관리
- CS 처리대장
- 관리자
- 보조 도구 > 업무 파일 자료실

## 최근 처리된 핵심 내용

- 다른 PC 작업 ZIP에서 복원한 DB 기준으로 로컬 실행 안정화
- `product_sales` 관련 DB 스키마 보강
- 통합관리대장/CS 처리대장 업로드 구분 및 관리자 권한 흐름 보강
- 매출현황 화면 구성 및 매입처별 총합계 금액 카드 반영
- 로그인 후 화면이 멈추던 lucide 아이콘 import 오류 수정
- 브라우저에서 로그인 후 메뉴 클릭 동작 확인
- 전체 테스트 통과: `55 tests OK`

## 테스트 방법

PowerShell에서 프로젝트 폴더로 이동한 뒤 실행합니다.

```powershell
python -m pip install -r requirements.txt
python -m unittest discover -s tests
```

정상 기준:

```text
Ran 55 tests
OK
```

## 주의사항

- `config/workhub.db`에는 실제 업무 데이터가 들어 있을 수 있으므로 외부 공유 시 주의해야 합니다.
- 네이버 메일 발송 테스트는 네이버 SMTP 설정과 계정 보안 설정이 필요합니다.
- Google Drive/rclone 백업은 VPS 또는 해당 PC에 rclone 연결을 별도로 설정해야 합니다.
- GitHub와 이어서 작업하려면 압축본보다 저장소를 clone/pull 하는 방식이 더 안전합니다.
