# (주)소일브릿지 발주 업무자동화

엑셀 기반 일일 발주/택배/CS 업무를 로컬 PC와 NAS에서 처리하기 위한 업무 자동화 앱입니다.

## 주요 기능

- 개별 택배건 정리
- 송장번호 추출
- 롯데택배 발주서 변환
- 차량인수증 생성
- 업체 CS 요청 메일 작성/전송
- CS 처리대장 DB 관리

## 로컬 실행

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\run_workhub_delivery_app.ps1
```

브라우저에서 아래 주소로 접속합니다.

```text
http://127.0.0.1:8765/
```

## PC용 실행파일 만들기

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File build_workhub_exe.ps1
```

완성된 파일은 `output` 폴더에 생성됩니다.

## NAS 배포

Synology NAS 배포는 [NAS_DEPLOY_README.md](NAS_DEPLOY_README.md)를 참고합니다.

## GitHub 제외 대상

다음 파일은 GitHub에 올리지 않습니다.

- `.build`, `output`, `outputs`
- `node_modules`
- DB 파일
- 메일 비밀번호/설정 파일
- 생성된 `exe`, `zip` 파일
