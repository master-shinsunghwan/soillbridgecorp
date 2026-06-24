# Workhub VPS 배포 및 인계서

이 문서는 `(주)소일브릿지 업무자동화 Workhub`를 다른 PC에서 수정하고, GitHub 저장 후 Hostinger VPS에 재배포하기 위한 기준 문서입니다.

주의: 이 문서에는 비밀번호, 개인 SSH 키, 메일 비밀번호, DB 원본 파일을 넣지 않습니다. 공개키와 접속 경로, 배포 절차만 기록합니다.

## 1. 기본 정보

- GitHub 저장소: `https://github.com/master-shinsunghwan/soillbridgecorp.git`
- 운영 브랜치: `main`
- 로컬 기준 작업 폴더: `C:\Users\ssh19\OneDrive\Documents\Codex\soillbridgecorp`
- 주요 실행 파일: `scripts/workhub_delivery_app.py`
- 주요 테스트 파일: `tests/test_workhub_app_feature_parity.py`
- 운영 주소: `https://workhub.soilbridgecorp.cloud/`
- VPS IP: `31.97.107.175`
- VPS 사용자: `root`
- VPS 소스 위치: `/opt/soilbridgecorp`
- VPS Docker Compose 위치: `/opt/company-erp`
- 운영 컨테이너명: `soilbridge-workhub`
- 운영 DB/업로드 데이터 위치: `/opt/workhub/data`
- 운영 백업 ZIP 위치: `/opt/workhub/backups`
- 운영 환경설정 파일: `/opt/company-erp/soilbridge-workhub.env`

## 2. 인증서 및 접속 키 정보

### SSH 접속 키

현재 이 PC의 SSH 개인키 경로:

```text
C:\Users\ssh19\.ssh\workhub_hostinger_ed25519
```

현재 이 PC의 SSH 공개키 내용:

```text
ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAIMYRN4+IVTGTvLLa6pon9WhJOLyQJIHzPw1HGrYKxWev workhub-hostinger-codex
```

다른 PC에서 작업할 때는 둘 중 하나를 선택합니다.

1. 기존 개인키를 안전하게 복사해서 같은 경로 또는 원하는 경로에 저장
2. 다른 PC에서 새 SSH 키를 만들고, 새 공개키를 VPS의 `/root/.ssh/authorized_keys`에 추가

개인키 파일은 절대 GitHub, 카카오톡, 일반 문서, 공유 폴더에 그대로 올리지 않습니다.

### HTTPS/SSL 인증서

운영 주소 `https://workhub.soilbridgecorp.cloud/`는 VPS의 Caddy 컨테이너가 HTTPS 인증서를 자동 관리합니다.

- Caddy 컨테이너명: `company-erp-caddy`
- Caddy 설정 위치: `/opt/company-erp/Caddyfile`
- 인증서 원본/개인키는 Caddy 내부 볼륨에서 관리되며, 일반 배포 문서에 복사하지 않습니다.
- 인증서 갱신은 Caddy가 자동 처리합니다.

확인 명령:

```bash
docker logs --tail 80 company-erp-caddy
docker exec company-erp-caddy cat /etc/caddy/Caddyfile
```

## 3. 다른 PC에서 처음 작업 준비

### GitHub에서 프로젝트 받기

```powershell
cd C:\Users\사용자명\Documents
git clone https://github.com/master-shinsunghwan/soillbridgecorp.git
cd soillbridgecorp
git checkout main
git pull --ff-only origin main
```

### Python 패키지 설치

```powershell
python -m pip install -r requirements.txt
```

### 로컬 실행

```powershell
python scripts\workhub_delivery_app.py
```

브라우저에서 아래 주소로 확인합니다.

```text
http://127.0.0.1:8781/
```

만약 8781 포트가 이미 사용 중이면 다른 포트로 실행하거나 기존 Workhub 프로세스를 종료합니다.

## 4. 수정 후 로컬 테스트

배포 전에 최소한 아래 검증은 진행합니다.

```powershell
python -m py_compile scripts\workhub_delivery_app.py
python -m unittest discover tests
```

브라우저에서 실제 화면도 확인합니다.

- 로그인 가능 여부
- 통합관리대장 화면
- CS처리대장 화면
- 관리자 화면 스크롤
- 연차 관리 화면
- 최근 수정한 버튼/팝업/필터 기능

## 5. GitHub 저장 절차

수정 후 아래 순서로 GitHub에 저장합니다.

```powershell
git status --short --branch
git add scripts/workhub_delivery_app.py tests/test_workhub_app_feature_parity.py
git commit -m "변경 내용 요약"
git push origin main
```

수정 파일이 더 있으면 `git add`에 같이 포함합니다.

아래 파일/폴더는 GitHub에 올리지 않습니다.

- `config/`
- `backups/`
- `output/`
- `sales_reports/`
- `node_modules/`
- `.env`
- `*.db`
- `*.zip`
- 실제 업무 엑셀 원본
- 개인키/비밀번호 파일

## 6. VPS 재배포 절차

### 방법 A: 다른 PC에서 SSH로 바로 배포

SSH 키가 등록되어 있으면 PowerShell에서 실행합니다.

```powershell
ssh -i $env:USERPROFILE\.ssh\workhub_hostinger_ed25519 root@31.97.107.175 "cd /opt/soilbridgecorp && git fetch origin main && git pull --ff-only origin main"
ssh -i $env:USERPROFILE\.ssh\workhub_hostinger_ed25519 root@31.97.107.175 "cd /opt/company-erp && docker compose -f docker-compose.hostinger.yml up -d --build workhub"
```

### 방법 B: Hostinger Terminal에서 배포

Hostinger VPS Terminal을 열고 아래 명령을 실행합니다.

```bash
cd /opt/soilbridgecorp
git fetch origin main
git pull --ff-only origin main

cd /opt/company-erp
docker compose -f docker-compose.hostinger.yml up -d --build workhub
```

## 7. 배포 후 확인

컨테이너 상태 확인:

```bash
docker ps --format "table {{.Names}}\t{{.Status}}\t{{.Ports}}"
```

Workhub 로그 확인:

```bash
docker logs --tail 80 soilbridge-workhub
```

운영 서버 내부 응답 확인:

```bash
docker exec company-erp-caddy wget -q -O- http://workhub:8787/ | head -40
```

최종 브라우저 확인:

```text
https://workhub.soilbridgecorp.cloud/?_fresh=배포확인
```

브라우저 캐시 때문에 화면이 예전처럼 보이면 `Ctrl + F5`로 강력 새로고침하거나, 주소 뒤에 `_fresh=숫자`를 붙여 확인합니다.

## 8. 운영 DB와 백업 주의사항

운영 DB는 VPS 내부 데이터 폴더에 저장됩니다.

```text
/opt/workhub/data
```

Docker 재배포는 코드를 새로 빌드하는 작업이며, 위 데이터 폴더를 지우지 않는 한 DB는 유지됩니다.

그래도 큰 업로드, 대량 수정, 배포 전에는 백업을 먼저 생성합니다.

```bash
ls -lh /opt/workhub/backups
```

Workhub 관리자 화면의 백업 메뉴에서도 백업 ZIP 생성 및 오프라인 다운로드를 확인합니다.

Google Drive 자동 백업은 VPS의 `rclone` 설정과 연결되어야 합니다.

확인 예시:

```bash
rclone listremotes
rclone lsd 원격드라이브명:
```

## 9. SSH 키를 새 PC에 등록하는 방법

### 새 키 만들기

다른 PC PowerShell에서 실행합니다.

```powershell
ssh-keygen -t ed25519 -C "workhub-hostinger-new-pc" -f $env:USERPROFILE\.ssh\workhub_hostinger_ed25519
```

생성된 공개키 확인:

```powershell
Get-Content $env:USERPROFILE\.ssh\workhub_hostinger_ed25519.pub
```

### VPS에 공개키 추가

Hostinger Terminal에서 아래 명령으로 공개키를 추가합니다.

```bash
mkdir -p /root/.ssh
chmod 700 /root/.ssh
nano /root/.ssh/authorized_keys
chmod 600 /root/.ssh/authorized_keys
```

`authorized_keys`에 새 PC의 `.pub` 내용을 한 줄로 붙여넣습니다.

### 접속 확인

다른 PC에서 확인합니다.

```powershell
ssh -i $env:USERPROFILE\.ssh\workhub_hostinger_ed25519 root@31.97.107.175 "hostname && docker ps --format 'table {{.Names}}\t{{.Status}}'"
```

## 10. 문제 발생 시 빠른 점검

### GitHub 최신 코드가 VPS에 반영됐는지 확인

```bash
cd /opt/soilbridgecorp
git status --short --branch
git log -1 --oneline
```

### 컨테이너 안의 코드가 최신인지 확인

```bash
docker exec soilbridge-workhub sh -lc "grep -n '헤르메스' /app/scripts/workhub_delivery_app.py | head"
```

확인할 문구는 최근 수정한 기능명으로 바꿔도 됩니다.

### 화면은 안 바뀌는데 배포는 된 것 같을 때

1. 브라우저에서 `Ctrl + F5`
2. 주소 뒤에 `?_fresh=현재시간` 붙이기
3. 컨테이너 재빌드 확인
4. Caddy가 `workhub:8787`로 연결 중인지 확인

```bash
docker exec company-erp-caddy cat /etc/caddy/Caddyfile
docker logs --tail 80 company-erp-caddy
docker logs --tail 80 soilbridge-workhub
```

## 11. 현재 운영 원칙

- 로컬에서 충분히 테스트한 뒤 배포합니다.
- 사용자가 “배포하자”라고 승인하기 전에는 VPS에 바로 반영하지 않습니다.
- 운영 데이터는 관리자 승인 없이 삭제하지 않습니다.
- DB, 백업 ZIP, 개인키, 메일 비밀번호는 GitHub에 올리지 않습니다.
- 다른 PC에서 작업해도 최종 저장 기준은 GitHub `main`입니다.
- VPS는 GitHub `main`을 받아서 Docker로 다시 빌드하는 방식으로 운영합니다.

