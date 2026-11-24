# 다음 연도 연차 선사용 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 신성환 실장이 직원별 다음 연도 연차 선사용을 확정 등록하고 새해 연차 발생 시 자동으로 한 번 차감하게 한다.

**Architecture:** SQLite 전용 장부로 선사용 상태를 관리하고 기존 연차 발생 함수가 미적용 선사용을 승인 완료 연차 이력과 잔여일수에 원자적으로 반영한다. 연차 API 진입 시 현재 연도 발생을 보장하며, 화면과 API 모두 신성환 실장 또는 관리자만 선사용 관리 기능에 접근한다.

**Tech Stack:** Python stdlib, SQLite, vanilla JavaScript, unittest

**Spec:** `docs/superpowers/specs/2026-11-24-next-year-leave-advance-design.md`

## Global Constraints

- 선사용 입력 권한은 신성환 실장과 관리자에게만 부여한다.
- 선사용 날짜는 입력 시점의 다음 연도만 허용한다.
- 연차 발생과 선사용 차감은 하나의 트랜잭션으로 처리한다.
- 적용 완료 행은 삭제하지 않고 감사 이력으로 유지한다.
- 기존 `static/workhub.css` 변경과 `runtime/` 폴더는 건드리지 않는다.

## Review Focus

- 표시명과 사용자명 변형에서도 신성환 실장을 정확히 식별하는지 확인한다.
- 일반 직원이 API를 직접 호출해도 등록·삭제가 거부되는지 확인한다.
- 반차와 중복 날짜 검증이 정확한지 확인한다.
- 연차 발생 재실행 시 선사용이 중복 차감되지 않는지 확인한다.
- 선사용 합계가 총 연차를 초과하면 일부 직원만 반영되는 상태 없이 전체 롤백되는지 확인한다.

---

### Task 1: 선사용 장부와 권한

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Create: `tests/test_leave_advance_usage.py`

**Interfaces:**
- Produces: `can_manage_next_year_leave(actor)`, `add_next_year_leave_usage(payload, actor)`, `list_next_year_leave_usage(actor)`, `delete_next_year_leave_usage(id, actor)`

- [ ] 권한, 다음 연도 날짜, 반차, 중복 입력, 대기 건 삭제 테스트를 작성한다.
- [ ] 테스트가 함수/테이블 부재로 실패하는지 확인한다.
- [ ] 장부 마이그레이션과 서비스 함수를 구현한다.
- [ ] 대상 테스트를 통과시킨다.

### Task 2: 연도 전환 자동 차감

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Test: `tests/test_leave_advance_usage.py`

**Interfaces:**
- Consumes: Task 1의 선사용 장부
- Produces: `apply_annual_leave_accrual()`의 선사용 차감 및 `ensure_current_year_leave_accrual()`

- [ ] 1회 차감, 승인 완료 이력 생성, 초과 시 롤백 테스트를 작성한다.
- [ ] 기존 연차 발생 함수에 미적용 선사용 반영을 추가한다.
- [ ] 연차 API 진입 시 현재 연도 자동 발생을 보장한다.
- [ ] 대상 테스트를 통과시킨다.

### Task 3: 전용 화면과 API

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Modify: `tests/test_workhub_app_feature_parity.py`

**Interfaces:**
- Produces: `/api/leave-advance`, `/api/leave-advance-delete`, `다음 연도 선사용` 탭

- [ ] 권한 플래그, 입력 필드, 목록, 삭제 버튼과 API 경로의 실패 테스트를 작성한다.
- [ ] 신성환 실장·관리자에게만 보이는 탭과 등록/삭제 요청을 구현한다.
- [ ] 화면 및 API 테스트를 통과시킨다.

### Task 4: 검증과 저장

**Files:**
- Commit only files in this plan.

- [ ] `python -m py_compile scripts/workhub_delivery_app.py`를 실행한다.
- [ ] 전체 unittest를 실행한다.
- [ ] diff 검사를 실행하고 관련 파일만 커밋·푸시한다.
