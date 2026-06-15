# (주)소일브릿지 발주 업무자동화

엑셀 기반 일일 발주/택배/CS 업무를 로컬 PC와 NAS에서 처리하기 위한 업무 자동화 앱입니다.

## 주요 기능

- 개별 택배건 정리
- 송장번호 추출
- 롯데택배 발주서 변환
- 차량인수증 생성
- 업체 CS 요청 메일 작성/전송
- CS 처리대장 DB 관리
- 회사 포털
- 업무관리
- 수출입 업무
- CRM 메신저 웹훅 수신 준비

## 로컬 실행

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\run_workhub_delivery_app.ps1
```

브라우저에서 아래 주소로 접속합니다.

```text
http://127.0.0.1:8765/
```

## CRM 메신저 웹훅

CRM 메신저 연동 화면에서 `X-Workhub-Webhook-Token` 값을 확인한 뒤, 카카오 챗봇 스킬 또는 공통 웹훅에서 아래 경로로 POST 요청을 보냅니다.

```text
/api/crm-messenger-webhook
```

운영 환경에서는 `WORKHUB_PUBLIC_BASE_URL`로 공개 HTTPS 주소를 고정할 수 있습니다.

```text
WORKHUB_PUBLIC_BASE_URL=https://업무도메인.example.com
```

`WORKHUB_CRM_WEBHOOK_TOKEN` 환경변수로 토큰을 고정할 수 있습니다. 이 경우 화면의 토큰 재발급 기능은 비활성 운영 방식으로 보고 사용할 수 없습니다.

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
