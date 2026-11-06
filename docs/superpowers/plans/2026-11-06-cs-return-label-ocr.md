# CS Return Label OCR Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add reviewed OCR import of return-label photos into the CS ledger with editable KakaoTalk text output.

**Architecture:** Put pure OCR text parsing, candidate scoring, and Kakao text formatting in a focused module. Keep HTTP endpoints, SQLite persistence, and the modal UI in the existing Workhub app. Analysis is read-only; a separate apply endpoint performs approved updates and attachment linking transactionally.

**Tech Stack:** Python 3.12, SQLite, Pillow, pytesseract/Tesseract kor+eng, vanilla HTML/CSS/JavaScript

**Spec:** `docs/superpowers/specs/2026-11-06-cs-return-label-ocr-design.md`

## Global Constraints

- No paid OCR API or external image transmission.
- Never mutate CS data during analysis.
- Reject unsupported or oversized uploads and duplicate return invoice application.
- Preserve unrelated changes in `static/workhub.css` and `runtime/`.

## Review Focus

- A label containing multiple 12-digit barcodes must prefer the number nearest the `운송장번호` heading.
- Masked Korean names such as `김*란` must match the full CS recipient without matching unrelated names.
- Weak or tied candidates must remain `확인 필요` rather than auto-selecting a case.
- An unmatched result must require explicit `신규 등록` action.
- Duplicate invoice numbers must not be applied to two CS cases.

---

### Task 1: OCR parsing, matching, and Kakao formatting

**Files:**
- Create: `scripts/workhub_cs_return_ocr.py`
- Create: `tests/test_cs_return_label_ocr.py`
- Modify: `requirements.txt`

**Interfaces:**
- Produces: `extract_return_label_fields(text)`, `match_return_label(parsed, cases)`, `format_kakao_return_text(rows, title_date)` and `ocr_return_label_image(data)`.

- [ ] Write failing tests for invoice preference, masked names, tied candidates, and Kakao output.
- [ ] Run `python -m pytest tests/test_cs_return_label_ocr.py -q` and confirm feature failures.
- [ ] Implement the minimal pure functions and image OCR rotation fallback.
- [ ] Add `Pillow==11.0.0` and rerun the tests.

### Task 2: Persistence and HTTP endpoints

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Modify: `tests/test_cs_return_label_ocr.py`

**Interfaces:**
- Consumes: Task 1 parsing/matching functions.
- Produces: `/api/cs-return-label-analyze`, `/api/cs-return-label-apply`, `cs_case_attachments` persistence.

- [ ] Write failing tests for analysis payloads, matched updates, explicit new-case creation, attachment linking, and duplicate rejection.
- [ ] Run the focused tests and confirm failures.
- [ ] Add schema, storage helpers, analysis/apply services, and authenticated endpoints.
- [ ] Rerun focused tests.

### Task 3: CS ledger review UI

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Modify: `tests/test_cs_return_label_ocr.py`

**Interfaces:**
- Consumes: Task 2 endpoints.
- Produces: upload modal, editable result table, progress, original preview, apply actions, and Kakao copy editor.

- [ ] Write failing static UI tests for controls and endpoint wiring.
- [ ] Run focused tests and confirm failures.
- [ ] Implement the responsive modal and client workflow.
- [ ] Run focused tests, `python -m py_compile scripts/workhub_delivery_app.py scripts/workhub_cs_return_ocr.py`, and the full test suite.

