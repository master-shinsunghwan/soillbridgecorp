# Workhub 설치 안내

업무허브는 Windows PC에서 바로가기로 실행하는 로컬 프로그램입니다.

## 설치 위치

```text
C:\Users\신성환\AppData\Local\Workhub
```

## 실행 방법

바탕화면의 `Workhub` 바로가기를 더블클릭합니다.

또는 시작 메뉴에서 `Workhub`를 실행합니다.

실행하면 브라우저가 열리고 아래 주소로 접속합니다.

```text
http://127.0.0.1:8765
```

## 포함 기능

- 개별 택배건 정리
- 송장번호 추출
- 롯데택배 발주서 변환
- 차량인수증 생성

## 재설치

작업 폴더에서 아래 파일을 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File install_workhub_app.ps1
```

## 삭제

작업 폴더에서 아래 파일을 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File uninstall_workhub_app.ps1
```

## 다른 PC에서 개발 이어가기

GitHub에서 프로젝트를 내려받은 뒤 작업 폴더에서 아래 순서로 진행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\run_workhub_delivery_app.ps1
```

설치형 실행 파일이 필요할 때는 아래 파일을 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File build_workhub_exe.ps1
```

주의: `output`, `.build`, `node_modules`, `config`, DB 파일, 메일 비밀번호 파일은 GitHub에 올리지 않습니다.
