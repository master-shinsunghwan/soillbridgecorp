# daisyUI Progressive UI Design

## Product Context

(주)소일브릿지 업무자동화는 발주업무, 통합관리대장, CS처리대장, 차량인수증, 연차관리, 백업/권한관리까지 포함한 내부 업무 도구다. 현재 앱은 `scripts/workhub_delivery_app.py`의 Python HTTP server 안에 HTML, CSS, JavaScript가 함께 들어 있는 구조다.

## Design Goal

daisyUI and Tailwind CSS should improve the existing operational UI without disrupting the current Python app architecture or ledger workflows. The goal is not a marketing redesign; it is a quieter, denser, more reliable back-office interface for repeated data entry and review.

## Chosen Approach

Use a progressive adoption path:

1. Add a Tailwind/daisyUI build pipeline that produces a local static CSS file.
2. Serve the generated CSS from the existing Python app.
3. Keep the current custom CSS and behavior in place while daisyUI is introduced.
4. Convert low-risk primitives first: buttons, inputs, selects, textareas, badges, modal surfaces, and dashboard cards.
5. Convert high-risk large tables last: 통합관리대장 and CS처리대장.

## Non-Goals

- Do not split the frontend into React/Vite in this phase.
- Do not replace the ledger table editing behavior while adding the CSS foundation.
- Do not depend on public CDN assets for production-like local testing.
- Do not redesign the app into a landing page or card-heavy marketing layout.

## Interaction Requirement

Full interactivity must remain intact. Existing upload, search, filter, download, table edit, CS 접수, vehicle receipt, leave request, admin, backup, and update flows must continue to work.

## Visual Direction

Use daisyUI with a restrained business theme. Keep the current dense operational layout, sidebar navigation, top bar, and large ledger workspace. Favor readable controls, clear selected states, compact tables, and predictable modal behavior.
