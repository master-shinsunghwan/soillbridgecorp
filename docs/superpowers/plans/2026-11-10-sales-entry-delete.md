# Sales Entry Delete Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 날짜별 매출 집계에서 상품, 매출처, 매입처 행을 선택해 안전하게 삭제한다.

**Architecture:** 서버의 고정된 종류 매핑으로 공통 조회/삭제 함수를 만들고 기존 상품 함수는 호환 래퍼로 유지한다. 화면은 세 가지 모드를 가진 단일 삭제 도구로 바꾸고 공통 API를 호출한다.

**Tech Stack:** Python, SQLite, vanilla JavaScript, unittest

**Spec:** `docs/superpowers/specs/2026-11-10-sales-entry-delete-design.md`

## Global Constraints

- 조회에는 유효한 ISO 매출일자가 필요하다.
- 일반 업로드는 선택한 행만, 수기 등록은 연결 묶음 전체를 삭제한다.
- 테이블과 열 선택은 서버의 허용 목록만 사용한다.

## Review Focus

- 잘못된 종류 값이 SQL에 전달되지 않고 오류 처리되는지 확인한다.
- 일반 업로드의 한 행 삭제가 다른 집계 행이나 업로드 기록을 지우지 않는지 확인한다.
- 수기 등록의 업체 행 삭제가 연결된 모든 집계를 정리하는지 확인한다.
- 매입처 화면에 매출액이 아닌 공급금액이 표시되는지 확인한다.
- 삭제 후 현재 목록과 대시보드가 함께 갱신되는지 확인한다.

---

### Task 1: 공통 조회와 삭제

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Test: `tests/test_sales_report_uploads.py`

**Interfaces:**
- Produces: `list_sales_report_entries(kind, report_date, query)` and `delete_sales_report_entry(kind, row_id)`

- [ ] 업체별 조회와 삭제 범위 테스트를 작성하고 실패를 확인한다.
- [ ] 서버 고정 매핑과 공통 조회/삭제 함수를 구현한다.
- [ ] 수기 등록 연쇄 삭제 및 잘못된 종류 테스트를 통과시킨다.

### Task 2: 삭제 화면과 API 연결

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Test: `tests/test_workhub_app_feature_parity.py`

**Interfaces:**
- Consumes: Task 1의 공통 조회/삭제 함수
- Produces: `/api/sales-report-entries`, `/api/sales-report-entry-delete`

- [ ] 세 모드와 공통 API가 렌더링되는지 실패 테스트를 작성한다.
- [ ] 전환 버튼, 동적 검색 문구, 공통 결과 목록과 확인창을 구현한다.
- [ ] 대상 테스트와 전체 테스트를 실행한다.

### Task 3: 저장

**Files:**
- Commit only files changed by this plan.

- [ ] 문법 검사, 전체 테스트, diff 검사를 실행한다.
- [ ] 관련 파일만 커밋하고 GitHub에 푸시한다.

