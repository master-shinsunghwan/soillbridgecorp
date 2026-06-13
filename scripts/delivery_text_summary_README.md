# 개별 택배건 전달 텍스트 생성기

`scripts/delivery_text_summary.py`는 주소일브릿지 Excel 파일을 읽어서 아래 형태의 텍스트를 생성합니다.

```text
★6월12일(금) 개별 택배건 전달드립니다★

상품명 - 1개 (3건)
상품명 - 2개 (1건)
```

## 기본 실행

```powershell
python scripts\delivery_text_summary.py "C:\path\to\주소일브릿지.xlsx"
```

결과 파일은 기본적으로 아래 폴더에 저장됩니다.

```text
output\delivery_text
```

## 정렬 방식

기본값은 상품명순입니다.

```powershell
python scripts\delivery_text_summary.py "C:\path\to\주소일브릿지.xlsx" --sort name
python scripts\delivery_text_summary.py "C:\path\to\주소일브릿지.xlsx" --sort count
python scripts\delivery_text_summary.py "C:\path\to\주소일브릿지.xlsx" --sort first
```

- `name`: 상품명순
- `count`: 건수 많은 순
- `first`: 엑셀에 나온 순서

## 자동 인식 컬럼

아래 컬럼명을 자동으로 찾습니다.

- 상품명: `제 품 명`, `제품명`, `상품명`, `품명`, `주문상품명`
- 수량: `수 량`, `수량`, `주문수량`, `개수`
- 주문번호: `주문번호`

주문번호에 `260612`처럼 날짜가 포함되어 있으면 제목의 날짜를 자동으로 만듭니다.
