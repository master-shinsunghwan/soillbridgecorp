# 업무허브 로컬 PC 테스트 패키지

## 실행 방법

1. 이 ZIP 파일을 원하는 폴더에 압축 해제합니다.
2. `로컬 테스트 실행.cmd`를 더블클릭합니다.
3. 처음 실행할 때는 Python 가상환경을 만들고 필요한 패키지를 설치하므로 시간이 조금 걸릴 수 있습니다.
4. 브라우저가 자동으로 열리면 아래 주소로 접속됩니다.

```text
http://127.0.0.1:8765/
```

## 기본 로그인

```text
관리자: admin / admin1234
사용자: user / user1234
```

## 테스트 데이터 저장 위치

이 패키지는 NAS가 아니라 압축 해제한 폴더 안의 `local_data`를 사용합니다.

```text
local_data/config/workhub.db
local_data/output/workhub_app
local_data/backups
```

## 주의사항

- Python 3.10 이상이 설치되어 있어야 합니다.
- 처음 실행 시 `pip install`이 필요하므로 인터넷 연결이 필요할 수 있습니다.
- 이 패키지는 로컬 테스트용입니다. 운영/NAS 공용 저장소 테스트는 기존 `workhub_run.cmd` 설정을 사용하세요.
