# VPS Hosting Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Workhub ready for Hostinger VPS operation through Nginx HTTPS, systemd, safe data paths, backup scripts, and initial-admin environment variables.

**Architecture:** Keep the Python server bound to `127.0.0.1:8770`, proxy all public traffic through Nginx, and store mutable data outside Git under `/opt/workhub`. Bootstrap the first admin account from environment variables only when the users table is empty.

**Tech Stack:** Python stdlib HTTP server, SQLite, systemd, Nginx, Bash deployment scripts, unittest.

---

### Task 1: Add Deployment Regression Tests

**Files:**
- Create: `tests/test_vps_deployment.py`

- [x] **Step 1: Write failing tests**

Add tests for required VPS files, `.env.example`, `.gitignore`, default password removal, and production initial admin creation.

- [x] **Step 2: Run tests and verify failure**

Run: `python -m unittest tests.test_vps_deployment -v`

Expected: FAIL because deployment files and bootstrap behavior are not implemented yet.

### Task 2: Add Initial Admin Environment Bootstrap

**Files:**
- Modify: `scripts/workhub_delivery_app.py`

- [x] **Step 1: Remove hardcoded default passwords**

Replace hardcoded default user passwords with environment-based first-admin creation.

- [x] **Step 2: Add environment-driven session settings**

Read `WORKHUB_SESSION_SECONDS`, `WORKHUB_SESSION_IDLE_SECONDS`, `WORKHUB_COOKIE_SECURE`, and `WORKHUB_COOKIE_SAMESITE`.

### Task 3: Add VPS Deployment Artifacts

**Files:**
- Create: `.env.example`
- Create: `deploy/systemd/workhub.service`
- Create: `deploy/nginx/workhub.conf`
- Create: `deploy/scripts/backup.sh`
- Create: `deploy/scripts/update.sh`
- Create: `deploy/scripts/rollback.sh`
- Create: `README_DEPLOY.md`
- Create: `DEBUG_GUIDE.md`
- Modify: `.gitignore`

- [x] **Step 1: Add files with production paths**

Use `/opt/soillbridgecorp` for code and `/opt/workhub` for data/backups.

- [x] **Step 2: Keep real env files ignored**

Ignore `.env` and `.env.*`, but allow `.env.example`.

### Task 4: Verify and Save

**Files:**
- All changed files

- [ ] **Step 1: Run deployment tests**

Run: `python -m unittest tests.test_vps_deployment -v`

- [ ] **Step 2: Run core Workhub checks**

Run: `python -m py_compile scripts/workhub_delivery_app.py`

- [ ] **Step 3: Commit and push**

Commit the verified changes and push to `origin/main`.
