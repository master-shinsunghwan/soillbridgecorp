# Workhub VPS Hosting Design

## Goal

Hostinger Ubuntu VPS에서 Workhub를 `https://erp.soilbridgecorp.cloud`로 운영할 수 있도록 배포 기준 파일, 운영 문서, 보안 기본값을 정리한다.

## Architecture

Workhub Python 서버는 VPS 내부 `127.0.0.1:8770`에만 바인딩한다. 외부 사용자는 Nginx HTTPS Reverse Proxy를 통해 접속한다. 데이터와 백업은 Git 저장소 밖의 `/opt/workhub/data`, `/opt/workhub/backups`에 저장한다.

## Security

코드에는 기본 관리자 비밀번호를 저장하지 않는다. 새 운영 DB의 첫 관리자 계정은 `/opt/workhub/.env`의 `WORKHUB_INITIAL_ADMIN_*` 환경변수로만 생성한다. 세션 쿠키는 운영 환경에서 `HttpOnly`, `SameSite=Lax`, `Secure`를 적용한다.

## Deployment Files

배포 파일은 `deploy/systemd`, `deploy/nginx`, `deploy/scripts` 아래에 둔다. 운영자는 `README_DEPLOY.md`와 `DEBUG_GUIDE.md`를 따라 최초 배포, 업데이트, 백업, 롤백, 오류 보고를 진행한다.

## Testing

배포 산출물의 존재와 주요 경로, 기본 비밀번호 제거, 운영 초기 관리자 생성을 `tests/test_vps_deployment.py`에서 검증한다.
