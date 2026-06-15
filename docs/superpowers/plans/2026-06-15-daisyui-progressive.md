# daisyUI Progressive UI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add daisyUI/Tailwind CSS to the existing Workhub Python app without breaking current workflows.

**Architecture:** Keep `scripts/workhub_delivery_app.py` as the application entrypoint and add a small Node-based CSS build. The generated CSS is served locally by the Python HTTP server, while current custom CSS remains in place during progressive migration.

**Tech Stack:** Python stdlib HTTP server, SQLite, openpyxl, Node.js, Tailwind CSS v4, daisyUI v5.

---

### Task 1: Add CSS Build Foundation

**Files:**
- Create: `package.json`
- Create: `src/workhub.css`
- Create: `tests/test_daisyui_foundation.py`
- Modify: `scripts/workhub_delivery_app.py`

- [ ] **Step 1: Write the failing foundation test**

Create `tests/test_daisyui_foundation.py` with assertions that the project has a Tailwind/daisyUI CSS input, a build script, a generated static CSS target, and an app stylesheet link.

- [ ] **Step 2: Run the test to verify it fails**

Run: `python tests/test_daisyui_foundation.py`
Expected: fail because the CSS build files and app stylesheet link are missing.

- [ ] **Step 3: Add package and CSS input**

Add `package.json` with `build:css` and `watch:css` scripts using `@tailwindcss/cli`, `tailwindcss`, and `daisyui`.

Add `src/workhub.css`:

```css
@import "tailwindcss";

@plugin "daisyui" {
  themes: corporate --default, business --prefersdark;
}

@source "../scripts/workhub_delivery_app.py";
```

- [ ] **Step 4: Serve static CSS from Python**

Add `STATIC_DIR = ROOT / "static"` and update `do_GET` so `/static/...` files are served by `serve_file`.

Add `<link rel="stylesheet" href="/static/workhub.css" />` to both the main app HTML and login HTML.

- [ ] **Step 5: Build CSS**

Run: `npm install`
Run: `npm run build:css`
Expected: `static/workhub.css` exists.

- [ ] **Step 6: Verify**

Run: `python tests/test_daisyui_foundation.py`
Run: `python -m py_compile scripts/workhub_delivery_app.py`
Start app and request `http://127.0.0.1:8765/static/workhub.css`.

### Task 2: Convert Low-Risk Controls

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Extend: `tests/test_daisyui_foundation.py`

- [ ] **Step 1: Add tests for class presence**

Assert buttons and common inputs include daisyUI-compatible classes in the HTML string.

- [ ] **Step 2: Apply daisyUI classes to buttons and inputs**

Add classes such as `btn`, `btn-primary`, `btn-error`, `input`, `select`, and `textarea` without removing existing classes.

- [ ] **Step 3: Verify app pages**

Run static tests, compile Python, open the app, and verify dashboard, upload modal, 통합관리대장, CS처리대장, and 연차관리 still render.

### Task 3: Convert Cards and Modal Surfaces

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Extend: `tests/test_daisyui_foundation.py`

- [ ] **Step 1: Add tests for card and modal compatibility**

Assert dashboard cards and modal surfaces keep existing IDs and gain daisyUI-compatible class hooks.

- [ ] **Step 2: Add daisyUI card/modal classes**

Add `card`, `card-body`, `modal-box`, and related classes while preserving existing custom classes and JS selectors.

- [ ] **Step 3: Visual verification**

Open dashboard, upload modal, notice popup, and CS popup. Confirm controls remain clickable and text does not overflow.

### Task 4: Prepare Ledger Table Migration

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Extend: `tests/test_daisyui_foundation.py`

- [ ] **Step 1: Add tests around ledger selectors**

Assert all existing table IDs and cell-edit IDs remain unchanged.

- [ ] **Step 2: Add non-invasive daisyUI table classes**

Add table classes without changing row generation, cell editing, filter popovers, or event listeners.

- [ ] **Step 3: Verify ledger workflows**

Upload sample data, select editable cells, apply a top edit bar change, save, refresh, and confirm data remains.

## Self-Review

- Spec coverage: The plan covers progressive setup, safe serving, low-risk controls, cards/modals, and delayed ledger table migration.
- Placeholder scan: No undefined implementation placeholders remain; later tasks define the target selectors and verification activities.
- Type consistency: File paths and script names match the current project structure.
