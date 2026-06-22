# Workhub 오류 확인 안내

운영 중 문제가 생기면 아래 순서로 확인합니다.

## 오류 보고 양식

```text
1. 오류가 발생한 메뉴:
예: CRM 업무관리 > 업무보드

2. 실행한 동작:
예: TASK-0007에서 완료 버튼 클릭

3. 기대한 결과:
예: 카드가 완료 칸으로 이동해야 함

4. 실제 결과:
예: 상태만 바뀐 것처럼 보이고 카드가 이동하지 않음

5. 화면 캡처:
첨부

6. 브라우저 Console 오류:
붙여넣기

7. 서버 로그:
sudo journalctl -u workhub -n 100 결과 붙여넣기

8. 발생 시간:
예: 2026-06-22 15:30
```

## 브라우저 오류 확인

브라우저 개발자도구를 열고 Console, Network 탭을 확인합니다.

```text
Console: 빨간 오류 메시지 확인
Network: 실패한 요청의 상태 코드와 응답 확인
```

## Workhub 서비스 상태 확인

```bash
sudo systemctl status workhub --no-pager
sudo journalctl -u workhub -n 100 --no-pager
sudo journalctl -u workhub -f
```

## Nginx 상태 확인

```bash
sudo nginx -t
sudo systemctl status nginx --no-pager
sudo journalctl -u nginx -n 100 --no-pager
```

## 포트 확인

```bash
sudo ss -tulpn | grep 8770
```

정상이라면 Python 서버가 `127.0.0.1:8770`에서 실행 중이어야 합니다.

## 서비스 재시작

```bash
sudo systemctl restart workhub
sudo systemctl status workhub --no-pager
```

## 최근 업데이트 확인

```bash
cd /opt/soillbridgecorp
git log --oneline -5
git status
```

## 백업 확인

```bash
ls -lh /opt/workhub/backups
```

DB 복원은 자동으로 하지 않습니다. 복원 전 현재 DB와 업로드 파일을 별도로 보관한 뒤 진행합니다.
