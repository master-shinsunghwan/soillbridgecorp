# (주)소일브릿지 발주 업무자동화 NAS 서버 설치 안내

대상 NAS: Synology DS920+ 계열

## 1. NAS 준비

1. DSM에 관리자 계정으로 접속합니다.
2. 패키지 센터에서 `Container Manager`를 설치합니다.
   - DSM 7에서는 Container Manager
   - DSM 6에서는 Docker 이름으로 보일 수 있습니다.
3. File Station에서 공유 폴더를 하나 만듭니다.
   - 예: `docker/workhub`

## 2. 파일 업로드

아래 파일/폴더를 NAS의 `docker/workhub` 폴더에 올립니다.

- `Dockerfile`
- `docker-compose.synology.yml`
- `requirements.txt`
- `scripts` 폴더
- `templates` 폴더

## 3. 비밀번호 암호화 키 변경

`docker-compose.synology.yml` 파일에서 아래 값을 임의의 긴 문자열로 바꿉니다.

```yaml
WORKHUB_SECRET_KEY: "CHANGE_ME_SOILBRIDGE_WORKHUB_SECRET"
```

예:

```yaml
WORKHUB_SECRET_KEY: "soilbridge-2026-workhub-private-key-very-long"
```

이 값은 네이버 메일 비밀번호 저장에 사용됩니다.  
값이 바뀌면 저장된 네이버 비밀번호는 다시 입력해야 합니다.

## 4. Container Manager에서 실행

1. Container Manager를 엽니다.
2. `Project` 또는 `프로젝트` 메뉴로 이동합니다.
3. 새 프로젝트를 만듭니다.
4. 경로는 `docker/workhub` 폴더를 선택합니다.
5. compose 파일은 `docker-compose.synology.yml`을 선택합니다.
6. 실행합니다.

## 5. 접속 주소

사무실 PC에서 브라우저를 열고 아래 주소로 접속합니다.

```text
http://NAS_IP주소:8765
```

예:

```text
http://192.168.0.50:8765
```

NAS IP 주소는 DSM 제어판의 네트워크 정보에서 확인할 수 있습니다.

## 6. 저장되는 데이터

NAS의 `workhub-data` 폴더에 저장됩니다.

- 업체 메일 주소록
- 네이버 메일 설정
- 업로드 파일
- 생성된 엑셀 파일

## 7. 주의사항

- 외부 인터넷에서 접속하게 열기 전에는 반드시 계정/접근 제한을 추가하는 것이 좋습니다.
- 네이버 메일 발송은 네이버 메일의 SMTP/외부메일 사용 설정이 켜져 있어야 합니다.
- 여러 PC에서 동시에 접속할 수 있지만, 같은 업체 주소록을 동시에 저장하는 작업은 순차적으로 사용하는 것을 권장합니다.
