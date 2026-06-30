# Workhub Windows PC 앱 배포

이 배포본은 직원이 웹 주소를 입력하지 않고 `(주)소일브릿지 업무자동화`를 Windows 앱처럼 실행하도록 만든 전용 실행파일입니다.

## 앱 방식

- 실행파일: `SoilbridgeWorkhub.exe`
- 표시 이름: `(주)소일브릿지 업무자동화`
- 접속 대상: `https://workhub.soilbridgecorp.cloud/`
- 화면 방식: Windows WebView2 앱 창
- 주소창, 브라우저 탭, 브라우저 뒤로가기 UI 없음
- 로그인 세션은 `%LOCALAPPDATA%\SoilbridgeWorkhubDesktop\WebViewData`에 저장

## 빌드

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build_workhub_desktop_app.ps1
```

완료되면 아래 위치에 배포용 ZIP이 생성됩니다.

```text
output\SoilbridgeWorkhub_Desktop_yyyyMMdd_HHmm.zip
```

## 직원 PC 설치

1. ZIP 파일을 직원 PC에 전달합니다.
2. 압축을 풉니다.
3. `설치.cmd`를 실행합니다.
4. 바탕화면 또는 시작 메뉴의 `(주)소일브릿지 업무자동화` 바로가기를 실행합니다.

설치 없이 바로 실행하려면 압축 해제 폴더의 `SoilbridgeWorkhub.exe`를 더블클릭하면 됩니다.

## 삭제

압축 해제 폴더의 `삭제.cmd`를 실행하면 바로가기와 설치 폴더가 삭제됩니다.

## 운영 참고

- 이 앱은 서버와 DB를 PC에 복사하지 않습니다. 기존 VPS Workhub를 앱 창으로 불러옵니다.
- VPS가 내려가 있거나 인터넷 연결이 끊기면 앱 안에서 연결 실패 화면과 다시 연결 버튼이 표시됩니다.
- 코드서명 인증서가 없으면 Windows SmartScreen이 첫 실행 시 경고를 표시할 수 있습니다. 사내 배포량이 늘어나면 코드서명 인증서 적용을 권장합니다.
