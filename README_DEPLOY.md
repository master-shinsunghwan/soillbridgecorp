# Workhub Hostinger VPS 배포 안내

이 문서는 `(주)소일브릿지 업무자동화`를 Hostinger Ubuntu VPS에서 운영하는 기준입니다.

운영 주소는 `https://erp.soilbridgecorp.cloud`를 기준으로 설명합니다. Python Workhub 서버는 외부에 직접 열지 않고 VPS 내부 `127.0.0.1:8770`에서만 실행하며, 외부 접속은 Nginx HTTPS 프록시가 받습니다.

## 운영 구조

```text
직원 PC / 노트북
-> https://erp.soilbridgecorp.cloud
-> Nginx HTTPS Reverse Proxy
-> 127.0.0.1:8770 Python Workhub Server
-> /opt/workhub/data SQLite DB 및 업로드 파일
-> /opt/workhub/backups 백업 파일
```

## 1. 서버 준비

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y git python3 python3-venv python3-pip nginx certbot python3-certbot-nginx sqlite3
sudo adduser --system --group --home /opt/workhub workhub
sudo mkdir -p /opt/workhub/data /opt/workhub/backups
```

## 2. 프로젝트 받기

```bash
cd /opt
sudo git clone https://github.com/master-shinsunghwan/soillbridgecorp.git
sudo chown -R workhub:workhub /opt/soillbridgecorp /opt/workhub
```

## 3. Python 및 화면 빌드 준비

```bash
sudo -u workhub bash
cd /opt/soillbridgecorp
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
exit
```

Node.js가 필요한 화면 빌드 작업을 서버에서 직접 할 경우:

```bash
cd /opt/soillbridgecorp
npm ci
npm run build:css
npm run build
```

## 4. 운영 환경 파일 작성

```bash
sudo cp /opt/soillbridgecorp/.env.example /opt/workhub/.env
sudo nano /opt/workhub/.env
sudo chown workhub:workhub /opt/workhub/.env
sudo chmod 600 /opt/workhub/.env
```

최초 운영 DB 생성 전에 반드시 아래 값을 채웁니다.

```env
WORKHUB_ENV=production
WORKHUB_HOST=127.0.0.1
WORKHUB_PORT=8770
WORKHUB_DATA_DIR=/opt/workhub/data
WORKHUB_BACKUP_DIR=/opt/workhub/backups
WORKHUB_COOKIE_SECURE=true
WORKHUB_SECRET_KEY=길고-랜덤한-문자열
WORKHUB_INITIAL_ADMIN_USERNAME=관리자아이디
WORKHUB_INITIAL_ADMIN_NAME=관리자이름
WORKHUB_INITIAL_ADMIN_PASSWORD=안전한초기비밀번호
```

첫 관리자 계정은 DB에 사용자가 하나도 없을 때만 생성됩니다. 이후에는 관리자 메뉴에서 계정을 관리합니다.

## 5. systemd 서비스 등록

```bash
sudo cp /opt/soillbridgecorp/deploy/systemd/workhub.service /etc/systemd/system/workhub.service
sudo systemctl daemon-reload
sudo systemctl enable workhub
sudo systemctl start workhub
sudo systemctl status workhub --no-pager
```

운영 명령:

```bash
sudo systemctl start workhub
sudo systemctl stop workhub
sudo systemctl restart workhub
sudo systemctl status workhub
sudo journalctl -u workhub -f
```

## 6. Nginx 및 SSL 적용

DNS에서 `erp.soilbridgecorp.cloud`의 A 레코드를 VPS 공인 IP로 연결합니다.

```bash
sudo cp /opt/soillbridgecorp/deploy/nginx/workhub.conf /etc/nginx/sites-available/workhub.conf
sudo ln -s /etc/nginx/sites-available/workhub.conf /etc/nginx/sites-enabled/workhub.conf
sudo nginx -t
sudo systemctl reload nginx
sudo certbot --nginx -d erp.soilbridgecorp.cloud
sudo nginx -t
sudo systemctl reload nginx
```

## 7. 업데이트

운영 서버에서 코드를 직접 수정하지 않습니다. 로컬/Codex에서 수정 후 GitHub에 올리고, VPS에서는 최신 코드만 가져옵니다.

```bash
cd /opt/soillbridgecorp
sudo deploy/scripts/update.sh
```

기본 흐름:

```text
Codex 또는 로컬 PC에서 수정
-> 로컬 테스트
-> main 브랜치에 병합 및 push
-> VPS에서 git pull
-> 서비스 재시작
```

운영이 안정되면 `dev` 브랜치에서 개발하고, 검증 후 `main`에 병합하는 방식을 권장합니다.

## 8. 백업

수동 백업:

```bash
sudo /opt/soillbridgecorp/deploy/scripts/backup.sh
```

매일 새벽 3시 자동 백업 예시:

```bash
sudo crontab -e
```

```cron
0 3 * * * /opt/soillbridgecorp/deploy/scripts/backup.sh >> /var/log/workhub-backup.log 2>&1
```

Google Drive 백업은 VPS에 rclone을 설치하고 Workhub 관리자 화면의 백업 설정에서 rclone 원격지를 연결하는 방식이 가장 현실적입니다.

```bash
sudo apt install -y rclone
rclone config
```

Docker 배포에서는 Workhub 컨테이너 안에서 rclone을 실행하므로, VPS의 rclone 설정을 컨테이너에 읽기 전용으로 연결해야 합니다.

```yaml
volumes:
  - /root/.config/rclone:/root/.config/rclone:ro
```

Workhub 백업 ZIP에는 `config/workhub.db`, 주요 설정 파일, `output/workhub_app`, `shared_files`, `sales_reports`가 포함됩니다. 장애 시 이 ZIP 하나로 업무 DB와 업로드/다운로드 산출물, 공유 자료, 매출 업로드 자료를 함께 복원할 수 있습니다.

## 9. 롤백

코드 롤백은 아래처럼 실행합니다.

```bash
cd /opt/soillbridgecorp
sudo deploy/scripts/rollback.sh <commit_hash>
```

DB 복원은 자동으로 하지 않습니다. 반드시 백업 파일을 먼저 확인하고 수동 절차로 진행합니다.

## 10. 주의사항

- 운영 DB, 업로드 파일, 백업 파일, `.env`는 GitHub에 올리지 않습니다.
- VPS에서는 Python 서버를 외부 IP에 직접 열지 않습니다.
- 외부 접속은 반드시 Nginx HTTPS를 통합니다.
- 운영 DB를 수정하기 전에는 백업을 먼저 만듭니다.
- 기본 관리자 비밀번호는 코드에 저장하지 않고 `/opt/workhub/.env`에만 둡니다.
