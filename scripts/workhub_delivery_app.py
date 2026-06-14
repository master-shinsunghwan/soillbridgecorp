from __future__ import annotations

import json
import mimetypes
import os
import re
import smtplib
import ssl
import sqlite3
import subprocess
import sys
import time
import threading
import base64
import ctypes
import hashlib
import hmac
import secrets
import tempfile
import zipfile
from copy import copy
from io import BytesIO
from datetime import date, datetime
from email.message import EmailMessage
from email.utils import formataddr
from html import escape as html_escape
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlsplit

from openpyxl import Workbook, load_workbook

from delivery_text_summary import summarize_workbook
from invoice_number_exporter import export_invoice_numbers, extract_invoice_rows
from lotte_order_form_converter import convert_lotte_order_form
from vehicle_receipt_generator import generate_vehicle_receipt


if getattr(sys, "frozen", False):
    ROOT = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    RUNTIME_ROOT = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "Workhub"
else:
    ROOT = Path(__file__).resolve().parents[1]
    RUNTIME_ROOT = ROOT

RUNTIME_ROOT = Path(os.environ.get("WORKHUB_DATA_DIR", str(RUNTIME_ROOT)))

OUTPUT_DIR = RUNTIME_ROOT / "output" / "workhub_app"
UPLOAD_DIR = OUTPUT_DIR / "uploads"
DOWNLOAD_DIR = OUTPUT_DIR / "downloads"
CONFIG_DIR = RUNTIME_ROOT / "config"
DB_PATH = CONFIG_DIR / "workhub.db"
MAIL_SETTINGS_PATH = CONFIG_DIR / "mail_settings.json"
VENDOR_CONTACTS_PATH = CONFIG_DIR / "vendor_contacts.json"
BACKUP_DIR = Path(os.environ.get("WORKHUB_BACKUP_DIR", str(RUNTIME_ROOT / "backups")))
LUCIDE_DIR = ROOT / "node_modules" / "lucide"
LOTTE_TEMPLATE = ROOT / "templates" / "lotte_order_form_template.xlsx"
MANAGEMENT_EXPORT_TEMPLATE = ROOT / "templates" / "management_ledger_export_template.xlsx"
NAVER_SMTP_HOST = "smtp.naver.com"
NAVER_SMTP_PORT = 465
SECRET_KEY_PATH = CONFIG_DIR / "secret.key"
TOKEN_PREFIX_DPAPI = "dpapi:"
TOKEN_PREFIX_KEY = "key1:"
SESSION_COOKIE_NAME = "workhub_session"
SESSION_SECONDS = 60 * 60 * 16
BACKUP_RETENTION_DAYS = 90
AUTO_BACKUP_HOUR = 3
_BACKUP_SCHEDULER_STARTED = False
_SYSTEM_UPDATE_LOCK = threading.Lock()
DEFAULT_USERS = (
    ("admin", "관리자", "admin", "admin1234"),
    ("user", "사용자", "user", "user1234"),
)
PERMISSION_DEFINITIONS = (
    ("ledger_delete", "대장 삭제", "통합관리대장/CS처리대장 선택 삭제"),
    ("notice_manage", "공지사항 관리", "공지사항 작성/수정/삭제"),
    ("ledger_edit", "대장 수정", "통합관리대장/CS처리대장 내용 저장"),
    ("excel_upload", "엑셀 업로드", "업로드/대량 등록 기능 사용"),
    ("excel_download", "엑셀 다운로드", "대장 및 변환 결과 다운로드"),
    ("cs_receive", "CS접수", "통합관리대장에서 CS 처리대장 접수"),
    ("mail_send", "메일 발송", "업체 CS 메일 발송"),
    ("user_admin", "사용자 관리", "계정 추가/수정/권한 변경"),
    ("backup_manage", "백업 관리", "수동/자동 백업 파일 관리"),
    ("system_update", "시스템 업데이트", "GitHub 업데이트 확인/적용"),
    ("leave_view", "연차 조회", "연차 내역 조회"),
    ("leave_approve", "연차 승인", "연차 신청 승인/반려"),
    ("leave_manage", "연차 관리", "연차 등록/수정/삭제"),
)
ALL_PERMISSIONS = tuple(key for key, _, _ in PERMISSION_DEFINITIONS)
DEFAULT_ROLE_PERMISSIONS = {
    "admin": ALL_PERMISSIONS,
    "user": ("ledger_edit", "excel_download", "cs_receive", "leave_view"),
}
LUCIDE_FALLBACK_JS = """
export function createIcons() {}
export const BriefcaseBusiness = {};
export const Home = {};
export const MessageCircle = {};
export const Info = {};
export const ChevronDown = {};
export const ChevronRight = {};
export const PlusSquare = {};
export const RefreshCw = {};
export const Ellipsis = {};
export const Headphones = {};
export const Package = {};
export const ClipboardCheck = {};
export const CircleDollarSign = {};
export const FileText = {};
export const FileSpreadsheet = {};
export const ClipboardList = {};
export const BarChart3 = {};
export const CopyCheck = {};
export const Bell = {};
export const Download = {};
export const Truck = {};
export const Mail = {};
export const Upload = {};
export const Database = {};
export const CalendarDays = {};
export const X = {};
""".strip()


HTML = r"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>(주)소일브릿지 발주 업무자동화</title>
  <style>
    :root {
      --bg: #f5f7fb;
      --panel: #ffffff;
      --line: #e3e8f2;
      --text: #111827;
      --muted: #667085;
      --navy: #071a3b;
      --navy-2: #10285a;
      --blue: #2563eb;
      --blue-soft: #dbeafe;
      --green: #079455;
      --green-soft: #dcfae6;
      --orange: #d97706;
      --orange-soft: #fef0c7;
      --purple: #7c3aed;
      --purple-soft: #ede9fe;
      --red: #dc2626;
      --red-soft: #fee2e2;
      --shadow: 0 10px 28px rgba(15, 23, 42, .08);
      font-family: Pretendard, Inter, "Noto Sans KR", "Malgun Gothic", Arial, sans-serif;
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      color: var(--text);
      background: var(--bg);
      letter-spacing: 0;
    }

    .app { min-height: 100vh; display: grid; grid-template-columns: 232px minmax(0, 1fr); }
    body.standalone .app { grid-template-columns: minmax(0, 1fr); }
    body.standalone .sidebar,
    body.standalone .top-search,
    body.standalone .top-tools { display: none; }
    body.standalone .topbar { grid-template-columns: 1fr; }
    .sidebar {
      background: linear-gradient(180deg, var(--navy), #081430);
      color: white;
      padding: 22px 16px;
    }
    .brand-icon {
      width: 38px; height: 38px; border-radius: 9px;
      display: grid; place-items: center;
      background: linear-gradient(145deg, #2f6df6, #7c3aed);
      color: white;
      box-shadow: 0 8px 18px rgba(15, 35, 70, .18);
      margin: 0;
      flex: 0 0 auto;
    }
    .brand-icon svg { width: 21px; height: 21px; }
    .brand {
      display: flex;
      align-items: center;
      gap: 11px;
      margin-bottom: 26px;
      padding: 0 4px;
    }
    .brand-label { font-size: 18px; font-weight: 900; line-height: 1.32; margin: 0; }
    .notice-board {
      min-height: 44px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      padding: 13px 15px;
      color: #344054;
      overflow: hidden;
      box-shadow: var(--shadow);
    }
    .notice-board-kicker {
      font-size: 11px;
      font-weight: 900;
      color: #155bc8;
      margin-bottom: 5px;
    }
    .notice-board-title {
      font-size: 16px;
      font-weight: 900;
      line-height: 1.32;
    }
    .notice-board-meta {
      font-size: 12px;
      color: #667085;
      margin-top: 3px;
    }
    .notice-board-body {
      font-size: 13px;
      line-height: 1.48;
      color: #475467;
      margin-top: 7px;
      max-height: 72px;
      overflow: hidden;
      white-space: pre-line;
    }
    .nav-item, .nav-section, .app-add {
      display: flex; align-items: center; gap: 13px;
      min-height: 43px; padding: 0 12px; border-radius: 8px;
      font-size: 14px; font-weight: 750; color: #d9e3ff;
      margin-bottom: 6px;
      border: 0;
      background: transparent;
      width: 100%;
      font-family: inherit;
      text-align: left;
      cursor: pointer;
    }
    .nav-item.active {
      color: white;
      background: rgba(72, 118, 255, .28);
      box-shadow: inset 0 0 0 1px rgba(255,255,255,.08);
    }
    .nav-item svg { width: 18px; height: 18px; flex: 0 0 auto; }
    .nav-item .nav-label {
      display: flex;
      align-items: center;
      gap: 13px;
      min-width: 0;
    }
    .nav-item .nav-chevron {
      margin-left: auto;
      width: 16px;
      height: 16px;
      transition: transform .16s ease;
    }
    .nav-group.open .nav-chevron { transform: rotate(90deg); }
    .nav-submenu {
      display: none;
      margin: -2px 0 8px 31px;
      padding-left: 10px;
      border-left: 1px solid rgba(255,255,255,.14);
    }
    .nav-group.open .nav-submenu { display: grid; gap: 4px; }
    .nav-subitem {
      min-height: 34px;
      padding: 0 10px;
      border: 0;
      border-radius: 7px;
      background: transparent;
      color: #c9d6f4;
      font-family: inherit;
      font-size: 13px;
      font-weight: 750;
      text-align: left;
      cursor: pointer;
    }
    .nav-subitem:hover,
    .nav-item:hover {
      background: rgba(255,255,255,.08);
      color: white;
    }
    .nav-section {
      min-height: auto;
      padding: 0 8px 8px;
      margin: 18px 0 0;
      color: #9fb0d3;
      font-size: 12px;
      font-weight: 850;
    }
    .hash { font-size: 15px; width: 22px; text-align: center; color: #d9e3ff; font-weight: 900; }
    .divider { height: 1px; background: rgba(255,255,255,.12); margin: 18px 6px; }

    main { min-width: 0; }
    .topbar {
      height: 74px;
      display: grid;
      grid-template-columns: 1fr minmax(260px, 430px) auto;
      align-items: center;
      gap: 18px;
      padding: 0 22px;
    }
    .title-wrap { display: grid; gap: 8px; }
    .title { display: flex; align-items: center; gap: 10px; font-size: 25px; font-weight: 900; line-height: 1.2; }
    .subtitle { display: none; }
    .top-actions { display: flex; gap: 12px; }
    .top-button {
      height: 38px; padding: 0 13px; border-radius: 8px; border: 1px solid var(--line);
      background: #fff; display: flex; align-items: center; gap: 8px;
      font-size: 13px; font-weight: 850; color: #344054;
    }
    .top-search {
      height: 42px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      display: flex;
      align-items: center;
      gap: 9px;
      padding: 0 14px;
      color: #98a2b3;
      font-size: 13px;
      font-weight: 700;
    }
    .top-tools {
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .icon-button {
      width: 38px;
      height: 38px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      display: grid;
      place-items: center;
      color: #344054;
      cursor: default;
    }
    .icon-button svg,
    .top-search svg { width: 17px; height: 17px; }
    .user-chip {
      height: 40px;
      display: flex;
      align-items: center;
      gap: 9px;
      padding: 0 11px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      font-size: 13px;
      font-weight: 850;
      white-space: nowrap;
    }
    .logout-button {
      height: 40px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 0 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      color: #475467;
      font-size: 13px;
      font-weight: 850;
      text-decoration: none;
    }
    .avatar {
      width: 27px;
      height: 27px;
      border-radius: 50%;
      background: linear-gradient(145deg, #155bc8, #08a66c);
    }

    .content { padding: 0 22px 24px; display: grid; gap: 16px; }
    .card {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }
    .summary { padding: 18px 21px 26px; }
    .section-title { margin: 0 0 26px; font-size: 24px; font-weight: 850; }
    .metrics { display: grid; grid-template-columns: repeat(4, 1fr); gap: 22px; align-items: center; }
    .metric { display: grid; grid-template-columns: 84px 1fr; gap: 18px; align-items: center; position: relative; }
    .metric:not(:last-child)::after {
      content: ""; position: absolute; right: -11px; top: 5px; height: 76px; width: 1px; background: #e0e4eb;
    }
    .metric-icon { width: 84px; height: 84px; border-radius: 17px; display: grid; place-items: center; }
    .metric-icon.red { background: var(--red-soft); color: var(--red); }
    .metric-icon.orange { background: var(--orange-soft); color: #7a430d; }
    .metric-icon.green { background: var(--green-soft); color: #0c5b34; }
    .metric-icon.blue { background: var(--blue-soft); color: var(--blue); }
    .metric-label { font-size: 17px; margin-bottom: 9px; }
    .metric-value { font-size: 40px; line-height: 1; font-weight: 800; color: var(--blue); }
    .metric-value.red { color: var(--red); }
    .metric-value.orange { color: #d16b0b; }
    .metric-value.green { color: var(--green); }

    .stat-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
    }
    .stat-card {
      padding: 18px 18px 17px;
      min-height: 116px;
      display: grid;
      grid-template-columns: 1fr 56px;
      gap: 10px;
      align-items: center;
    }
    .stat-label {
      color: #344054;
      font-size: 13px;
      font-weight: 850;
      margin-bottom: 10px;
    }
    .stat-value {
      font-size: 28px;
      line-height: 1;
      font-weight: 950;
    }
    .stat-trend {
      margin-top: 12px;
      font-size: 12px;
      color: var(--green);
      font-weight: 850;
    }
    .stat-trend.red { color: var(--red); }
    .stat-icon {
      width: 54px;
      height: 54px;
      border-radius: 16px;
      display: grid;
      place-items: center;
      font-size: 24px;
      font-weight: 950;
    }
    .stat-icon svg { width: 24px; height: 24px; }
    .stat-icon.blue { color: var(--blue); background: var(--blue-soft); }
    .stat-icon.green { color: var(--green); background: var(--green-soft); }
    .stat-icon.orange { color: var(--orange); background: var(--orange-soft); }
    .stat-icon.purple { color: var(--purple); background: var(--purple-soft); }

    .dashboard-card { padding: 0; }
    .dashboard-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 13px;
    }
    .dashboard-title {
      font-size: 17px;
      font-weight: 950;
    }
    .notice-template {
      display: grid;
      gap: 10px;
      padding: 14px;
    }
    .notice-template-grid {
      display: grid;
      grid-template-columns: 150px 1fr 160px;
      gap: 8px;
    }
    .notice-template input,
    .notice-template textarea {
      border: 1px solid #cbd5e1;
      border-radius: 7px;
      padding: 0 10px;
      font-family: inherit;
      font-size: 13px;
      background: white;
    }
    .notice-template input { height: 34px; }
    .notice-template textarea {
      min-height: 78px;
      padding-top: 9px;
      resize: vertical;
    }
    .notice-template-actions {
      display: flex;
      justify-content: flex-end;
      gap: 8px;
    }
    .notice-preview {
      border: 1px solid #e5e7eb;
      border-radius: 7px;
      padding: 10px 12px;
      background: #f8fafc;
      font-size: 13px;
      line-height: 1.5;
      color: #344054;
    }
    .notice-preview strong {
      display: block;
      color: #111827;
      font-size: 14px;
      margin-bottom: 3px;
    }
    .notice-popup-backdrop {
      position: fixed;
      inset: 0;
      display: none;
      place-items: center;
      background: rgba(15, 23, 42, .28);
      z-index: 70;
      padding: 18px;
    }
    .notice-popup-backdrop.open { display: grid; }
    .notice-popup {
      width: min(680px, 100%);
      border-radius: 10px;
      border: 1px solid #d7dce5;
      background: white;
      box-shadow: 0 22px 50px rgba(15, 23, 42, .24);
    }
    .notice-popup-head {
      height: 52px;
      padding: 0 16px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      border-bottom: 1px solid #e5e7eb;
      font-weight: 950;
    }
    .action-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; }
    .action-card {
      min-height: 150px;
      padding: 16px;
      display: grid;
      grid-template-columns: 46px minmax(0, 1fr);
      gap: 12px;
      align-items: start;
      transition: border-color .16s ease, box-shadow .16s ease, transform .16s ease;
    }
    .action-card:hover {
      border-color: #b7c1d1;
      box-shadow: 0 12px 28px rgba(15, 23, 42, .08);
      transform: translateY(-1px);
    }
    .action-main { min-width: 0; }
    .action-top { display: contents; }
    .action-icon { width: 42px; height: 42px; border-radius: 10px; display: grid; place-items: center; flex: 0 0 auto; }
    .action-icon svg { width: 20px; height: 20px; }
    .action-icon.blue { background: var(--blue-soft); color: var(--blue); }
    .action-icon.green { background: var(--green-soft); color: var(--green); }
    .action-icon.orange { background: var(--orange-soft); color: var(--orange); }
    .action-icon.purple { background: var(--purple-soft); color: var(--purple); }
    .action-kicker {
      height: 23px;
      padding: 0 12px;
      display: inline-flex;
      align-items: center;
      border-radius: 999px;
      font-size: 11px;
      font-weight: 850;
      background: #f3f6fb;
      color: #566174;
      white-space: nowrap;
      margin-bottom: 9px;
    }
    .action-kicker.blue { background: var(--blue-soft); color: var(--blue); }
    .action-kicker.green { background: var(--green-soft); color: var(--green); }
    .action-kicker.orange { background: var(--orange-soft); color: var(--orange); }
    .action-kicker.purple { background: var(--purple-soft); color: var(--purple); }
    .action-title { font-size: 15px; font-weight: 950; margin: 0 0 6px; }
    .action-sub { min-height: 38px; color: var(--muted); font-size: 12px; line-height: 1.45; margin: 0 0 12px; }
    .action-button {
      height: 31px;
      min-width: 88px;
      padding: 0 12px;
      border-radius: 7px;
      border: 1px solid #cbd5e1;
      background: white;
      color: #1f2937;
      font-size: 12px;
      font-weight: 900;
      cursor: pointer;
    }
    .action-button.green { color: #1f2937; border-color: #cbd5e1; background: white; }
    .action-button.orange { color: #1f2937; border-color: #cbd5e1; background: white; }
    .action-button.purple { color: #1f2937; border-color: #cbd5e1; background: white; }
    .action-button:hover { filter: brightness(.985); }

    .lower { display: grid; grid-template-columns: 2fr 1fr; gap: 34px; }
    .panel { padding: 18px 22px 20px; min-height: 270px; }
    .list-item { display: grid; grid-template-columns: 57px 1fr auto; gap: 17px; align-items: center; padding: 10px 0 18px; border-bottom: 1px solid #e4e7ed; }
    .small-icon { width: 57px; height: 57px; border-radius: 10px; display: grid; place-items: center; }
    .small-icon.blue { background: var(--blue-soft); color: var(--blue); }
    .small-icon.green { background: var(--green-soft); color: var(--green); }
    .small-icon.purple { background: var(--purple-soft); color: var(--purple); }
    .item-title { font-size: 18px; font-weight: 800; margin-bottom: 4px; }
    .item-sub { font-size: 16px; color: var(--muted); }
    .more { width: 100%; height: 53px; border: 1px solid #d7dce5; border-radius: 8px; background: white; font-size: 17px; color: #343b46; margin-top: 12px; }
    .status { padding: 8px 13px; border-radius: 7px; background: #dff5e7; color: #0a6c3a; font-weight: 800; }
    .status.warn { background: #ffe1e1; color: #d20c1c; }

    .modal-backdrop {
      position: fixed; inset: 0; display: none; place-items: center;
      background: rgba(249, 250, 252, .42);
      z-index: 20;
    }
    .modal-backdrop.open { display: grid; }
    .modal {
      width: min(620px, calc(100vw - 38px));
      max-height: calc(100vh - 38px);
      overflow-y: auto;
      background: white; border: 1px solid #b7bdc8; border-radius: 12px;
      box-shadow: var(--shadow); padding: 24px 28px 26px;
      position: relative;
    }
    .modal.ledger-modal {
      width: calc(100vw - 18px);
      height: calc(100vh - 18px);
      max-height: calc(100vh - 18px);
      padding: 18px 18px 20px;
      overflow: hidden;
    }
    .modal.ledger-modal #uploadForm {
      height: calc(100% - 54px);
      display: flex;
      flex-direction: column;
      min-height: 0;
    }
    .modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 24px; }
    .modal-title { font-size: 25px; font-weight: 850; }
    .close { border: 0; background: transparent; color: #3f4650; cursor: pointer; padding: 4px; }
    .field-label { display: block; font-size: 18px; font-weight: 750; margin-bottom: 10px; }
    .dropzone {
      border: 1px dashed #9aa4b2; border-radius: 8px; background: #fbfcff;
      padding: 22px; min-height: 112px; display: grid; gap: 8px; align-content: center;
      cursor: pointer;
      transition: border-color .15s ease, background .15s ease, box-shadow .15s ease;
    }
    .dropzone.dragover {
      border-color: var(--blue);
      background: #eef5ff;
      box-shadow: 0 0 0 3px rgba(21, 91, 200, .12);
    }
    .drop-main { font-size: 17px; font-weight: 750; color: #1a2230; }
    .drop-sub { font-size: 14px; color: var(--muted); }
    input[type="file"] { display: none; }
    .options { margin-top: 16px; display: flex; gap: 12px; align-items: center; }
    select {
      height: 41px; border: 1px solid #aab2bf; border-radius: 7px; padding: 0 12px;
      font-size: 15px; background: white; min-width: 180px;
    }
    .modal-actions { display: flex; justify-content: flex-end; gap: 14px; margin-top: 24px; }
    .btn {
      height: 50px; min-width: 106px; border-radius: 8px; border: 1px solid #b9c0ca;
      background: white; font-size: 18px; font-weight: 760; cursor: pointer;
    }
    .btn.primary { background: linear-gradient(180deg, #08a66c, #047a4d); border-color: #08794f; color: white; }
    .btn.blue { background: linear-gradient(180deg, #1f73e8, #145bc8); border-color: #145bc8; color: white; }
    .result {
      margin-top: 18px; display: none;
    }
    .result.open { display: block; }
    textarea {
      width: 100%; height: 240px; resize: vertical; border: 1px solid #ccd3dd;
      border-radius: 8px; padding: 14px; font-size: 15px; line-height: 1.55;
      font-family: "Malgun Gothic", "Noto Sans KR", sans-serif;
    }
    .result-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 10px; }
    .notice { margin-top: 10px; min-height: 20px; color: var(--muted); font-size: 14px; }
    .admin-panel {
      display: grid;
      gap: 14px;
      padding: 16px;
      background: white;
      border: 1px solid var(--line);
      border-radius: 8px;
    }
    .admin-form {
      display: grid;
      grid-template-columns: 1.1fr 1.1fr .8fr 1fr auto auto;
      gap: 10px;
      align-items: end;
    }
    .admin-form label {
      display: grid;
      gap: 6px;
      font-size: 12px;
      font-weight: 850;
      color: #344054;
    }
    .admin-form input,
    .admin-form select {
      height: 34px;
      border: 1px solid #cfd6e2;
      border-radius: 7px;
      padding: 0 9px;
      font-size: 13px;
      font-weight: 700;
      background: white;
    }
    .admin-check {
      height: 34px;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 13px;
      font-weight: 850;
      color: #344054;
    }
    .admin-check input { width: 16px; height: 16px; }
    .permission-grid {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 8px;
      padding: 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #f8fafc;
    }
    .permission-item {
      min-height: 42px;
      display: flex;
      align-items: flex-start;
      gap: 7px;
      padding: 8px;
      border: 1px solid #e5e7eb;
      border-radius: 7px;
      background: white;
      font-size: 12px;
      font-weight: 850;
      color: #344054;
    }
    .permission-item input { width: 15px; height: 15px; margin-top: 1px; }
    .permission-item small {
      display: block;
      margin-top: 3px;
      color: #667085;
      font-size: 11px;
      font-weight: 650;
      line-height: 1.35;
      white-space: normal;
    }
    .admin-table-wrap {
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
    }
    .admin-table {
      width: 100%;
      min-width: 760px;
      border-collapse: collapse;
      font-size: 13px;
    }
    .admin-table th {
      height: 34px;
      padding: 0 10px;
      background: #eef6ff;
      border-bottom: 1px solid var(--line);
      text-align: left;
      font-weight: 900;
    }
    .admin-table td {
      height: 38px;
      padding: 4px 10px;
      border-bottom: 1px solid #edf1f7;
      white-space: nowrap;
    }
    .admin-table tr:last-child td { border-bottom: 0; }
    .admin-action {
      height: 30px;
      border: 1px solid #cfd6e2;
      border-radius: 7px;
      background: white;
      font-size: 12px;
      font-weight: 850;
      cursor: pointer;
    }
    .admin-message { min-height: 20px; color: var(--muted); font-size: 13px; font-weight: 750; }
    .permission-hidden { display: none !important; }
    .leave-panel { display: grid; gap: 14px; }
    .leave-summary-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; }
    .leave-summary-card {
      min-height: 86px;
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      box-shadow: var(--shadow);
    }
    .leave-summary-card span { display: block; color: var(--muted); font-size: 13px; font-weight: 850; }
    .leave-summary-card strong { display: block; margin-top: 8px; font-size: 28px; line-height: 1; }
    .leave-summary-card.accent strong { color: var(--green); }
    .leave-tabs { display: flex; flex-wrap: wrap; gap: 8px; }
    .leave-tab {
      height: 36px;
      padding: 0 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      color: #344054;
      font-size: 13px;
      font-weight: 900;
      cursor: pointer;
    }
    .leave-tab.active { border-color: var(--blue); color: var(--blue); background: #eff6ff; }
    .leave-tab-panel { display: none; }
    .leave-tab-panel.active { display: block; }
    .leave-grid { display: grid; gap: 14px; }
    .leave-grid.two { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    .leave-card {
      min-width: 0;
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      box-shadow: var(--shadow);
    }
    .leave-card.narrow { max-width: 720px; }
    .leave-card-title { margin-bottom: 12px; font-size: 16px; font-weight: 950; }
    .leave-form { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; align-items: end; }
    .leave-form label { display: grid; gap: 6px; font-size: 12px; font-weight: 900; color: #344054; }
    .leave-form label.wide { grid-column: 1 / -1; }
    .leave-form input,
    .leave-form select,
    .leave-form textarea {
      width: 100%;
      min-height: 36px;
      border: 1px solid #cfd6e2;
      border-radius: 7px;
      padding: 8px 10px;
      font-size: 13px;
      font-weight: 700;
      background: white;
    }
    .leave-form textarea { min-height: 92px; resize: vertical; }
    .leave-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    .leave-table th {
      height: 34px;
      padding: 0 10px;
      background: #eef6ff;
      border-bottom: 1px solid var(--line);
      text-align: left;
      font-weight: 900;
      white-space: nowrap;
    }
    .leave-table td {
      min-height: 38px;
      padding: 8px 10px;
      border-bottom: 1px solid #edf1f7;
      vertical-align: top;
    }
    .leave-table tr:last-child td { border-bottom: 0; }
    .leave-action-row { display: flex; gap: 6px; }
    .leave-action {
      height: 30px;
      border: 1px solid #cfd6e2;
      border-radius: 7px;
      background: white;
      font-size: 12px;
      font-weight: 850;
      cursor: pointer;
    }
    .leave-action.approve { border-color: #0b8f55; color: #067647; }
    .leave-action.reject { border-color: #f4a7a7; color: #b42318; }
    .leave-message { min-height: 20px; color: var(--muted); font-size: 13px; font-weight: 750; }
    .backup-panel { display: grid; gap: 14px; }
    .backup-summary-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; }
    .backup-summary-card {
      min-height: 86px;
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      box-shadow: var(--shadow);
    }
    .backup-summary-card span { display: block; color: var(--muted); font-size: 13px; font-weight: 850; }
    .backup-summary-card strong {
      display: block;
      margin-top: 8px;
      font-size: 18px;
      line-height: 1.35;
      word-break: break-all;
    }
    .backup-card {
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      box-shadow: var(--shadow);
    }
    .backup-note {
      margin: 0 0 12px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 750;
      line-height: 1.6;
    }
    .backup-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    .backup-table th {
      height: 34px;
      padding: 0 10px;
      background: #eef6ff;
      border-bottom: 1px solid var(--line);
      text-align: left;
      font-weight: 900;
      white-space: nowrap;
    }
    .backup-table td {
      height: 40px;
      padding: 6px 10px;
      border-bottom: 1px solid #edf1f7;
      vertical-align: middle;
      white-space: nowrap;
    }
    .backup-table tr:last-child td { border-bottom: 0; }
    .backup-actions { display: flex; gap: 6px; }
    .backup-action {
      height: 30px;
      padding: 0 10px;
      border: 1px solid #cfd6e2;
      border-radius: 7px;
      background: white;
      font-size: 12px;
      font-weight: 850;
      cursor: pointer;
    }
    .backup-action.danger { border-color: #f4a7a7; color: #b42318; }
    .backup-message { min-height: 20px; color: var(--muted); font-size: 13px; font-weight: 750; }
    .backup-restore-row {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
      margin: 0 0 12px;
    }
    .backup-restore-row input { display: none; }
    #backupRestoreInput { display: none; }
    .system-panel { display: grid; gap: 14px; }
    .system-summary-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; }
    .system-summary-card {
      min-height: 92px;
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      box-shadow: var(--shadow);
    }
    .system-summary-card span { display: block; color: var(--muted); font-size: 13px; font-weight: 850; }
    .system-summary-card strong {
      display: block;
      margin-top: 8px;
      font-size: 17px;
      line-height: 1.35;
      word-break: break-all;
    }
    .system-card {
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: white;
      box-shadow: var(--shadow);
    }
    .system-note {
      margin: 0 0 12px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 750;
      line-height: 1.6;
    }
    .system-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    .system-table th {
      height: 34px;
      padding: 0 10px;
      background: #eef6ff;
      border-bottom: 1px solid var(--line);
      text-align: left;
      font-weight: 900;
      white-space: nowrap;
    }
    .system-table td {
      min-height: 40px;
      padding: 8px 10px;
      border-bottom: 1px solid #edf1f7;
      vertical-align: top;
      white-space: nowrap;
    }
    .system-table tr:last-child td { border-bottom: 0; }
    .system-message { min-height: 20px; color: var(--muted); font-size: 13px; font-weight: 750; }
    @media (max-width: 1100px) {
      .admin-form { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .permission-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .leave-grid.two,
      .leave-summary-grid,
      .backup-summary-grid,
      .system-summary-grid { grid-template-columns: 1fr; }
    }
    .vehicle-fields { display: none; }
    .cs-fields { display: none; }
    .ledger-fields { display: none; }
    .management-fields { display: none; }
    .ledger-cs-popup-head { display: none; }
    .modal.ledger-modal .cs-fields.ledger-cs-popup {
      position: absolute;
      z-index: 35;
      top: 70px;
      right: 22px;
      bottom: 72px;
      display: block !important;
      width: min(560px, calc(100vw - 62px));
      overflow: auto;
      padding: 18px;
      border: 1px solid #9aa4b2;
      border-radius: 10px;
      background: white;
      box-shadow: 0 18px 44px rgba(15, 23, 42, .28);
    }
    .modal.ledger-modal .cs-fields.ledger-cs-popup .ledger-cs-popup-head {
      position: sticky;
      top: -18px;
      z-index: 2;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin: -18px -18px 14px;
      padding: 15px 18px;
      border-bottom: 1px solid #d7dce5;
      background: white;
    }
    .ledger-cs-popup-title {
      font-size: 20px;
      font-weight: 850;
    }
    .ledger-cs-popup-close {
      height: 34px;
      min-width: 64px;
      border: 1px solid #aab2bf;
      border-radius: 6px;
      background: white;
      font-family: inherit;
      font-weight: 800;
      cursor: pointer;
    }
    .text-field { margin-top: 14px; }
    .text-field input,
    .text-field select,
    .text-field textarea {
      width: 100%;
      border: 1px solid #aab2bf;
      border-radius: 7px;
      font-size: 16px;
      font-family: inherit;
      background: #fff;
    }
    .text-field input,
    .text-field select {
      height: 48px;
      padding: 0 13px;
    }
    .text-field textarea {
      min-height: 78px;
      padding: 12px 13px;
      resize: vertical;
      line-height: 1.45;
    }
    .product-table {
      margin-top: 12px;
      display: grid;
      gap: 8px;
      max-height: 280px;
      overflow: auto;
      padding-right: 4px;
    }
    .product-row {
      display: grid;
      grid-template-columns: 1fr 95px 95px;
      gap: 9px;
    }
    .product-row input {
      height: 42px;
      border: 1px solid #aab2bf;
      border-radius: 7px;
      padding: 0 11px;
      font-size: 15px;
      font-family: inherit;
    }
    .add-row {
      margin-top: 10px;
      height: 40px;
      border: 1px solid #a084c8;
      color: var(--purple);
      background: #fdfbff;
      border-radius: 7px;
      font-weight: 750;
      cursor: pointer;
    }
    .checkbox-field {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-top: 12px;
      color: #323b48;
      font-size: 15px;
      font-weight: 650;
    }
    .checkbox-field input { width: 17px; height: 17px; }
    .cs-toolbar {
      display: flex;
      gap: 10px;
      margin-top: 12px;
    }
    .cs-save-button {
      height: 40px;
      border: 1px solid #155bc8;
      color: #155bc8;
      background: #f4f8ff;
      border-radius: 7px;
      font-weight: 750;
      cursor: pointer;
      padding: 0 14px;
    }
    .cs-case-list {
      margin-top: 16px;
      border: 1px solid #d7dce5;
      border-radius: 8px;
      overflow: hidden;
      background: #fbfcff;
    }
    .cs-case-head {
      padding: 10px 12px;
      font-size: 15px;
      font-weight: 850;
      border-bottom: 1px solid #d7dce5;
      background: #f3f6fb;
    }
    .cs-case-item {
      padding: 10px 12px;
      border-bottom: 1px solid #e5e9f0;
      font-size: 14px;
      line-height: 1.45;
      color: #253041;
    }
    .cs-case-item:last-child { border-bottom: 0; }
    .cs-case-meta {
      color: #687385;
      font-size: 13px;
      margin-top: 3px;
    }
    .ledger-toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 7px;
      margin-bottom: 10px;
      align-items: center;
    }
    .ledger-toolbar input { flex: 1 1 280px; min-width: 220px; }
    .ledger-toolbar input,
    .ledger-toolbar select {
      height: 34px;
      border: 1px solid #aab2bf;
      border-radius: 7px;
      padding: 0 9px;
      font-size: 12px;
      font-family: inherit;
      background: white;
    }
    .ledger-toolbar select { flex: 0 0 auto; max-width: 160px; }
    .ledger-toolbar .btn {
      height: 34px;
      min-width: auto;
      padding: 0 11px;
      border-radius: 7px;
      font-size: 12px;
    }
    .ledger-import-button {
      height: 34px;
      min-width: 86px;
      border: 1px solid #0d6ddf;
      border-radius: 7px;
      background: #eef6ff;
      color: #0d56c2;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      font-size: 12px;
      font-weight: 850;
      cursor: pointer;
    }
    .ledger-import-button svg {
      width: 14px;
      height: 14px;
    }
    .ledger-count {
      height: 34px;
      display: inline-flex;
      align-items: center;
      padding: 0 8px;
      border: 1px solid #d7dce5;
      border-radius: 7px;
      background: #f8fafc;
      color: #475467;
      font-size: 12px;
      font-weight: 800;
      white-space: nowrap;
    }
    .ledger-import-button input { display: none; }
    .management-month-tabs {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-top: 8px;
      padding: 8px 0 0;
    }
    .management-month-tab {
      height: 30px;
      min-width: 54px;
      padding: 0 10px;
      border: 1px solid #cfd6e2;
      border-radius: 7px;
      background: white;
      color: #344054;
      font-size: 12px;
      font-weight: 850;
      cursor: pointer;
    }
    .management-month-tab.active {
      border-color: #145bc8;
      background: #eef6ff;
      color: #145bc8;
    }
    .management-download-group {
      display: inline-flex;
      flex-wrap: wrap;
      gap: 5px;
      align-items: center;
    }
    .ledger-wrap {
      border: 1px solid #d7dce5;
      border-radius: 8px;
      overflow-x: scroll;
      overflow-y: auto;
      max-height: calc(100vh - 230px);
      max-width: 100%;
      background: white;
      scrollbar-gutter: stable both-edges;
    }
    .modal.ledger-modal.ledger-view .ledger-fields {
      display: flex !important;
      flex-direction: column;
      flex: 1;
      min-height: 0;
    }
    .modal.ledger-modal.management-view .management-fields {
      display: flex !important;
      flex-direction: column;
      flex: 1;
      min-height: 0;
    }
    .modal.ledger-modal .ledger-wrap {
      flex: 1;
      min-height: 0;
      max-height: none;
    }
    .modal.ledger-modal .management-wrap {
      flex: 1;
      min-height: 0;
      max-height: none;
    }
    .ledger-table {
      width: 100%;
      min-width: 2600px;
      border-collapse: collapse;
      font-size: 12px;
    }
    .ledger-table th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: #8ecf45;
      border-bottom: 1px solid #6baa2d;
      color: #111827;
      font-weight: 850;
      padding: 5px 6px;
      text-align: center;
      white-space: nowrap;
      vertical-align: bottom;
      line-height: 1.18;
    }
    .ledger-table th.has-filter {
      padding-right: 26px;
    }
    .ledger-th-title {
      display: block;
      line-height: 1.2;
    }
    .ledger-filter-trigger {
      position: absolute;
      right: 5px;
      top: 50%;
      transform: translateY(-50%);
      width: 19px;
      height: 19px;
      border: 1px solid rgba(17, 24, 39, .42);
      border-radius: 3px;
      background: rgba(255,255,255,.88);
      color: #111827;
      font-size: 11px;
      line-height: 1;
      cursor: pointer;
      padding: 0;
    }
    .ledger-filter-trigger.active {
      background: #155bc8;
      border-color: #0f4aaa;
      color: white;
    }
    .ledger-filter-popover {
      position: fixed;
      z-index: 40;
      display: none;
      width: 260px;
      max-height: 390px;
      padding: 10px;
      border: 1px solid #98a2b3;
      border-radius: 7px;
      background: white;
      box-shadow: 0 14px 32px rgba(15, 23, 42, .24);
    }
    .ledger-filter-popover.open { display: block; }
    .ledger-filter-title {
      font-size: 13px;
      font-weight: 850;
      margin-bottom: 8px;
    }
    .ledger-filter-search {
      width: 100%;
      height: 34px;
      border: 1px solid #aab2bf;
      border-radius: 5px;
      padding: 0 8px;
      font-family: inherit;
      font-size: 13px;
      margin-bottom: 8px;
    }
    .ledger-filter-option-list {
      max-height: 230px;
      overflow: auto;
      border: 1px solid #e2e6ee;
      border-radius: 5px;
      margin-bottom: 9px;
    }
    .ledger-filter-option {
      width: 100%;
      min-height: 28px;
      border: 0;
      border-bottom: 1px solid #eef1f5;
      background: white;
      padding: 5px 7px;
      text-align: left;
      font-family: inherit;
      font-size: 12px;
      cursor: pointer;
    }
    .ledger-filter-option:hover { background: #eef6ff; }
    .ledger-filter-actions {
      display: flex;
      justify-content: flex-end;
      gap: 7px;
    }
    .ledger-filter-actions button {
      height: 30px;
      border-radius: 5px;
      border: 1px solid #aab2bf;
      background: white;
      font-family: inherit;
      font-weight: 800;
      cursor: pointer;
      padding: 0 10px;
    }
    .ledger-filter-actions .apply {
      border-color: #087a46;
      background: #087a46;
      color: white;
    }
    .ledger-table th.invoice-head {
      background: #f3b21d;
      border-bottom-color: #c98e10;
    }
    .ledger-table td {
      border-bottom: 1px solid #e6eaf0;
      padding: 4px 6px;
      vertical-align: middle;
      text-align: center;
      color: #1f2937;
      font-size: 11px;
      line-height: 1.22;
    }
    .ledger-table tr.completed-cs td {
      background: #fff8d8;
    }
    .ledger-table tr.row-dirty td {
      box-shadow: inset 0 0 0 9999px rgba(37, 99, 235, .035);
    }
    .ledger-table tr.management-duplicate td {
      background: var(--duplicate-row-color, #eef6ff);
    }
    .ledger-table td.left { text-align: left; }
    .ledger-status {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 64px;
      height: 26px;
      border-radius: 999px;
      background: #eaf2ff;
      color: #155bc8;
      font-weight: 800;
      font-size: 12px;
    }
    .ledger-edit,
    .ledger-status-select {
      width: 100%;
      min-width: 118px;
      height: 28px;
      border: 1px solid #aab2bf;
      border-radius: 6px;
      padding: 0 6px;
      font-size: 11px;
      font-family: inherit;
      background: white;
    }
    .ledger-status-select { min-width: 128px; }
    .ledger-save {
      height: 28px;
      min-width: 48px;
      border: 1px solid #087a46;
      border-radius: 6px;
      background: #eefaf3;
      color: #087a46;
      font-weight: 850;
      font-size: 11px;
      cursor: pointer;
    }
    .ledger-check {
      width: 16px;
      height: 16px;
      accent-color: #155bc8;
      cursor: pointer;
    }
    .management-edit {
      width: 100%;
      min-width: 86px;
      height: 26px;
      border: 1px solid transparent;
      border-radius: 5px;
      padding: 0 5px;
      background: rgba(255,255,255,.72);
      font-family: inherit;
      font-size: 11px;
    }
    .management-edit:focus {
      border-color: #2563eb;
      background: white;
      outline: none;
      box-shadow: 0 0 0 2px rgba(37, 99, 235, .12);
    }
    .management-edit.wide { min-width: 260px; }
    .management-cs-button {
      height: 28px;
      min-width: 58px;
      border: 1px solid #155bc8;
      border-radius: 6px;
      background: #eef6ff;
      color: #155bc8;
      font-size: 11px;
      font-weight: 850;
      cursor: pointer;
    }
    .management-cs-button:disabled {
      border-color: #7a5a00;
      background: rgba(255,255,255,.46);
      color: #5f4300;
      cursor: default;
    }
    .ledger-table tr.management-cs-received td {
      background: #ffc000 !important;
    }
    .management-cs-received .management-edit {
      background: rgba(255,255,255,.48);
    }
    .workspace-view {
      display: none;
      flex-direction: column;
      gap: 12px;
      min-height: 0;
      height: calc(100vh - 74px);
      padding: 0 22px 24px;
    }
    .workspace-view.active { display: flex; }
    .workspace-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      min-height: 48px;
    }
    .workspace-title {
      font-size: 18px;
      font-weight: 950;
    }
    .workspace-actions {
      display: flex;
      gap: 8px;
      align-items: center;
    }
    .workspace-button {
      height: 34px;
      border: 1px solid #cbd5e1;
      border-radius: 7px;
      background: white;
      color: #1f2937;
      font-family: inherit;
      font-size: 12px;
      font-weight: 900;
      padding: 0 12px;
      cursor: pointer;
    }
    .workspace-button.danger {
      border-color: #f1a7a7;
      background: #fff1f1;
      color: #b42318;
    }
    .workspace-mount {
      flex: 1;
      min-height: 0;
      display: flex;
      flex-direction: column;
    }
    .workspace-view.active .ledger-fields,
    .workspace-view.active .management-fields {
      display: flex !important;
      flex-direction: column;
      flex: 1;
      min-height: 0;
    }
    .workspace-view.active .ledger-wrap {
      flex: 1;
      min-height: 0;
      max-height: none;
    }

    @media (max-width: 1180px) {
      .app { grid-template-columns: 76px minmax(0, 1fr); }
      .brand { justify-content: center; padding: 0; }
      .brand-label, .nav-item > span, .nav-label span, .nav-section, .nav-submenu { display: none; }
      .nav-item { justify-content: center; padding: 0; }
      .stat-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .action-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    }
    @media (max-width: 760px) {
      .app { grid-template-columns: 1fr; }
      .sidebar { display: none; }
      .topbar {
        height: auto;
        padding: 16px 12px;
        grid-template-columns: 1fr;
      }
      .top-tools { display: none; }
      .content { padding: 0 12px 18px; }
      .title { font-size: 22px; }
      .stat-grid,
      .action-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div class="app">
    <aside class="sidebar">
      <div class="brand">
        <div class="brand-icon"><i data-lucide="briefcase-business"></i></div>
        <div class="brand-label">(주)소일브릿지<br>업무자동화</div>
      </div>
      <div class="nav-section">MAIN</div>
      <div class="nav-group open" id="noticeNavGroup">
        <button class="nav-item active" id="noticeNavToggle" type="button" data-view="dashboard">
          <span class="nav-label"><i data-lucide="home"></i> <span>공지사항</span></span>
          <i class="nav-chevron" data-lucide="chevron-right"></i>
        </button>
        <div class="nav-submenu">
          <button class="nav-subitem" id="noticeInputOpen" type="button">공지사항 입력</button>
        </div>
      </div>
      <div class="nav-group" id="orderNavGroup">
        <button class="nav-item" id="orderNavToggle" type="button">
          <span class="nav-label"><i data-lucide="clipboard-list"></i> <span>발주업무</span></span>
          <i class="nav-chevron" data-lucide="chevron-right"></i>
        </button>
        <div class="nav-submenu">
          <button class="nav-subitem" type="button" data-open="delivery">개별 택배건 정리</button>
          <button class="nav-subitem" type="button" data-open="invoice">송장번호 추출</button>
          <button class="nav-subitem" type="button" data-open="lotte">롯데택배 발주서 변환</button>
          <button class="nav-subitem" type="button" data-open="vehicle">차량인수증</button>
        </div>
      </div>
      <button class="nav-item" type="button" data-open="management"><i data-lucide="database"></i> <span>통합관리대장 관리</span></button>
      <button class="nav-item" type="button" data-open="ledger"><i data-lucide="clipboard-check"></i> <span>CS 처리대장</span></button>
      <button class="nav-item" type="button" data-open="cs"><i data-lucide="mail"></i> <span>업체 메일</span></button>
      __LEAVE_NAV__
      __ADMIN_NAV__
      __BACKUP_NAV__
      __SYSTEM_NAV__
      <div class="nav-section">TOOLS</div>
      <button class="nav-item" type="button" data-open="invoice"><i data-lucide="file-spreadsheet"></i> <span>송장 추출</span></button>
      <button class="nav-item" type="button" data-open="vehicle"><i data-lucide="truck"></i> <span>차량인수증</span></button>
    </aside>

    <main>
      <header class="topbar">
        <div class="title-wrap">
          <div class="title">금일 공지사항 <i data-lucide="chevron-down"></i></div>
          <p class="subtitle">발주 파일 변환과 인수증 생성을 한 곳에서 처리합니다.</p>
        </div>
        <div class="top-search"><i data-lucide="file-text"></i> 파일명, 수령인, 송장번호, CS내용 검색</div>
        <div class="top-tools">
          <button class="icon-button" type="button"><i data-lucide="bell"></i></button>
          <button class="icon-button" type="button"><i data-lucide="refresh-cw"></i></button>
          <div class="user-chip"><span class="avatar"></span><span>__USER_DISPLAY__</span></div>
          <a class="logout-button" href="/logout">로그아웃</a>
        </div>
      </header>

      <section class="content" id="dashboardContent">
        <div class="stat-grid">
          <article class="card stat-card">
            <div>
              <div class="stat-label">오늘 택배건</div>
              <div class="stat-value">352</div>
              <div class="stat-trend">전일 대비 ▲ 12.5%</div>
            </div>
            <div class="stat-icon blue"><i data-lucide="package"></i></div>
          </article>
          <article class="card stat-card">
            <div>
              <div class="stat-label">송장 추출</div>
              <div class="stat-value">87</div>
              <div class="stat-trend">완료율 ▲ 5.3%</div>
            </div>
            <div class="stat-icon purple"><i data-lucide="file-spreadsheet"></i></div>
          </article>
          <article class="card stat-card">
            <div>
              <div class="stat-label">CS 미처리</div>
              <div class="stat-value">23</div>
              <div class="stat-trend red">전일 대비 ▲ 3.1%</div>
            </div>
            <div class="stat-icon orange"><i data-lucide="headphones"></i></div>
          </article>
          <article class="card stat-card">
            <div>
              <div class="stat-label">완료 처리</div>
              <div class="stat-value">128</div>
              <div class="stat-trend">전일 대비 ▲ 18.2%</div>
            </div>
            <div class="stat-icon green"><i data-lucide="copy-check"></i></div>
          </article>
        </div>

        <section class="notice-board" id="sidebarNoticePreview">
          <div class="notice-board-kicker">금일 공지사항</div>
          <div class="notice-board-title">등록된 공지 없음</div>
          <div class="notice-board-body">공지사항 입력 메뉴를 눌러 내용을 입력해주세요.</div>
        </section>

        <section class="dashboard-card">
          <div class="dashboard-head">
            <div class="dashboard-title">빠른 실행</div>
          </div>
          <div class="action-grid">
          <article class="card action-card">
            <div class="action-icon blue"><i data-lucide="file-text"></i></div>
            <div class="action-main">
                <span class="action-kicker blue">텍스트</span>
              <h2 class="action-title">개별 택배건 정리</h2>
              <p class="action-sub">주소일브릿지 엑셀을 전달용 텍스트로 변환합니다.</p>
              <button class="action-button" data-open="delivery">엑셀 업로드</button>
            </div>
          </article>

          <article class="card action-card">
            <div class="action-icon green"><i data-lucide="file-spreadsheet"></i></div>
            <div class="action-main">
                <span class="action-kicker green">엑셀</span>
              <h2 class="action-title">송장번호 추출</h2>
              <p class="action-sub">출고송장 엑셀에서 수하인별 송장번호 엑셀을 생성합니다.</p>
              <button class="action-button green" data-open="invoice">엑셀 업로드</button>
            </div>
          </article>

          <article class="card action-card">
            <div class="action-icon orange"><i data-lucide="clipboard-list"></i></div>
            <div class="action-main">
                <span class="action-kicker orange">양식</span>
              <h2 class="action-title">롯데택배 발주서 변환</h2>
              <p class="action-sub">주소일브릿지 원본을 업로드해 지정 양식으로 출력합니다.</p>
              <button class="action-button orange" data-open="lotte">엑셀 업로드</button>
            </div>
          </article>

          <article class="card action-card">
            <div class="action-icon purple"><i data-lucide="truck"></i></div>
            <div class="action-main">
                <span class="action-kicker purple">직접입력</span>
              <h2 class="action-title">차량인수증 생성</h2>
              <p class="action-sub">직접 입력한 제품 내역을 인수증 양식으로 출력합니다.</p>
              <button class="action-button purple" data-open="vehicle">인수증 입력</button>
            </div>
          </article>

          <article class="card action-card">
            <div class="action-icon blue"><i data-lucide="mail"></i></div>
            <div class="action-main">
                <span class="action-kicker blue">메일</span>
              <h2 class="action-title">업체 CS 요청</h2>
              <p class="action-sub">업체별 CS 요청 내용을 작성해 네이버 메일로 전송합니다.</p>
              <button class="action-button" data-open="cs">메일 작성</button>
            </div>
          </article>

          <article class="card action-card">
            <div class="action-icon green"><i data-lucide="clipboard-check"></i></div>
            <div class="action-main">
                <span class="action-kicker green">DB</span>
              <h2 class="action-title">CS 처리대장</h2>
              <p class="action-sub">자동화 DB에 저장된 CS건을 처리대장 형태로 조회합니다.</p>
              <button class="action-button green" data-open="ledger">처리대장 보기</button>
            </div>
          </article>
          </div>
        </section>
      </section>
      <section class="workspace-view" id="managementWorkspace">
        <div class="workspace-head">
          <div class="workspace-title">통합관리대장 관리</div>
          <div class="workspace-actions">
            <button class="workspace-button" type="button" id="managementSaveAll">해당 내용 저장</button>
            <button class="workspace-button danger" type="button" id="managementDeleteSelected">선택 주문 삭제</button>
            <button class="workspace-button" type="button" data-open-window="management">새창으로 열기</button>
          </div>
        </div>
        <div class="workspace-mount" id="managementWorkspaceMount"></div>
      </section>
      <section class="workspace-view" id="ledgerWorkspace">
        <div class="workspace-head">
          <div class="workspace-title">CS 처리대장</div>
          <div class="workspace-actions">
            <button class="workspace-button" type="button" id="ledgerSaveAll">해당 내용 저장</button>
            <button class="workspace-button danger" type="button" id="ledgerDeleteSelected">선택 주문 삭제</button>
            <button class="workspace-button" type="button" data-open-window="ledger">새창으로 열기</button>
          </div>
        </div>
        <div class="workspace-mount" id="ledgerWorkspaceMount"></div>
      </section>
      __LEAVE_WORKSPACE__
      __ADMIN_WORKSPACE__
      __BACKUP_WORKSPACE__
      __SYSTEM_WORKSPACE__
    </main>
  </div>

  <div class="notice-popup-backdrop" id="noticePopup">
    <div class="notice-popup" role="dialog" aria-modal="true">
      <div class="notice-popup-head">
        <span>공지사항 입력</span>
        <button class="close" id="noticePopupClose" type="button" aria-label="닫기"><i data-lucide="x"></i></button>
      </div>
      <div class="notice-template">
        <div class="notice-template-grid">
          <input id="noticeDateInput" type="date" />
          <input id="noticeTitleInput" type="text" placeholder="공지 제목" />
          <input id="noticeOwnerInput" type="text" placeholder="담당자" />
        </div>
        <textarea id="noticeBodyInput" placeholder="공지 내용을 입력해주세요. 예) 금일 출고 마감 시간 / 택배 특이사항 / 업체 회신 필요 건"></textarea>
        <div class="notice-template-actions">
          <button class="workspace-button" type="button" id="noticeClearButton">초기화</button>
          <button class="workspace-button" type="button" id="noticeSaveButton">공지 저장</button>
        </div>
        <div class="notice-preview" id="noticePreview">
          <strong>저장된 공지사항이 없습니다.</strong>
          공지사항을 입력하고 저장하면 이곳에서 미리 볼 수 있습니다.
        </div>
      </div>
    </div>
  </div>

  <div class="modal-backdrop" id="modal">
    <div class="modal" role="dialog" aria-modal="true">
      <div class="modal-head">
        <div class="modal-title" id="modalTitle">파일 업로드</div>
        <button class="close" id="closeModal" aria-label="닫기"><i data-lucide="x"></i></button>
      </div>
      <form id="uploadForm">
        <label class="field-label" id="fileLabel">엑셀 파일 선택</label>
        <label class="dropzone" for="fileInput">
          <span class="drop-main" id="dropMain">파일을 선택하거나 여기에 올려주세요.</span>
          <span class="drop-sub" id="dropSub">xlsx 파일만 처리합니다.</span>
          <input id="fileInput" name="file" type="file" accept=".xlsx,.xlsm" />
        </label>
        <div id="templateUpload" style="display:none; margin-top:16px;">
          <label class="field-label" for="templateInput">양식 파일 선택</label>
          <label class="dropzone" for="templateInput">
            <span class="drop-main" id="templateDropMain">롯데택배 발주서 양식을 선택해주세요.</span>
            <span class="drop-sub">출력 파일은 이 양식의 서식을 따라갑니다.</span>
            <input id="templateInput" name="template" type="file" accept=".xlsx,.xlsm" />
          </label>
        </div>
        <div class="options" id="deliveryOptions">
          <label for="sortMode">정렬</label>
          <select id="sortMode" name="sort">
            <option value="name">상품명순</option>
            <option value="count">건수 많은 순</option>
            <option value="first">엑셀 순서</option>
          </select>
        </div>
        <div class="vehicle-fields" id="vehicleFields">
          <div class="text-field">
            <label class="field-label" for="receiptTypeSelect">차량인수증 타입</label>
            <select id="receiptTypeSelect" name="receipt_type">
              <option value="일반">일반</option>
              <option value="모드니 전용">모드니 전용</option>
            </select>
          </div>
          <div class="text-field">
            <label class="field-label" for="supplierInput">공급받는자</label>
            <input id="supplierInput" name="supplier" type="text" placeholder="예) 모드니" />
          </div>
          <div class="text-field">
            <label class="field-label" for="receiptDateInput">일자</label>
            <input id="receiptDateInput" name="receipt_date" type="date" />
          </div>
          <div class="text-field">
            <label class="field-label" for="freightPaymentSelect">운임비용</label>
            <select id="freightPaymentSelect" name="freight_payment">
              <option value="선불">선불</option>
              <option value="후불">후불</option>
            </select>
          </div>
          <label class="field-label" style="margin-top:16px;">제품명 / 수량 / 입수량</label>
          <div class="product-table" id="productTable"></div>
          <button class="add-row" id="addProductRow" type="button">제품 줄 추가</button>
          <div class="text-field">
            <label class="field-label" for="requestNoteInput">요청사항</label>
            <textarea id="requestNoteInput" name="request_note" placeholder="예) 파손주의 / 현장 사진 문자 요청"></textarea>
          </div>
          <div class="text-field">
            <label class="field-label" for="deliveryPlaceInput">납품장소</label>
            <input id="deliveryPlaceInput" name="delivery_place" type="text" placeholder="예) 모드니 물류센터" />
          </div>
          <div class="text-field">
            <label class="field-label" for="managerInput">담당자명</label>
            <input id="managerInput" name="manager" type="text" placeholder="예) 홍길동 / 010-0000-0000" />
          </div>
        </div>
        <div class="cs-fields" id="csFields">
          <div class="ledger-cs-popup-head">
            <strong class="ledger-cs-popup-title">CS 추가</strong>
            <button class="ledger-cs-popup-close" id="ledgerCsPopupClose" type="button">닫기</button>
          </div>
          <div class="text-field">
            <label class="field-label" for="naverEmailInput">네이버 메일 아이디</label>
            <input id="naverEmailInput" name="naver_email" type="text" placeholder="예) soilbridge@naver.com" />
          </div>
          <div class="text-field">
            <label class="field-label" for="naverPasswordInput">네이버 메일 비밀번호</label>
            <input id="naverPasswordInput" name="naver_password" type="password" placeholder="저장된 비밀번호가 없으면 입력" autocomplete="current-password" />
          </div>
          <label class="checkbox-field">
            <input id="saveMailCredentials" type="checkbox" checked />
            <span>아이디/비밀번호 저장</span>
          </label>
          <div class="text-field">
            <label class="field-label" for="vendorContactSelect">업체 선택</label>
            <select id="vendorContactSelect">
              <option value="">업체를 선택해주세요</option>
            </select>
          </div>
          <div class="text-field">
            <label class="field-label">업체 메일 주소록 엑셀 업로드</label>
            <label class="dropzone" for="vendorContactsFileInput">
              <span class="drop-main" id="vendorContactsDropMain">업체명/메일주소 엑셀을 선택해주세요.</span>
              <span class="drop-sub">헤더 예시: 업체명, 메일주소</span>
              <input id="vendorContactsFileInput" name="vendor_contacts" type="file" accept=".xlsx,.xlsm" />
            </label>
          </div>
          <div class="text-field">
            <label class="field-label" for="recipientEmailInput">받는 업체 메일</label>
            <input id="recipientEmailInput" name="recipient_email" type="email" placeholder="예) vendor@example.com" />
          </div>
          <div class="text-field">
            <label class="field-label" for="vendorNameInput">업체명</label>
            <input id="vendorNameInput" name="vendor_name" type="text" placeholder="예) 키친쿡" />
          </div>
          <button class="add-row" id="saveVendorContact" type="button">업체 메일 저장/업데이트</button>
          <div class="text-field">
            <label class="field-label" for="csOriginInput">원출고일 및 원송장번호</label>
            <input id="csOriginInput" name="cs_origin" type="text" placeholder="예) 2026-06-13 / 1234567890" />
          </div>
          <div class="text-field">
            <label class="field-label" for="csProductInput">상품명</label>
            <input id="csProductInput" name="cs_product" type="text" />
          </div>
          <div class="text-field">
            <label class="field-label" for="csReceiverInput">수령인</label>
            <input id="csReceiverInput" name="cs_receiver" type="text" />
          </div>
          <div class="text-field">
            <label class="field-label" for="csPhoneInput">수령인 연락처</label>
            <input id="csPhoneInput" name="cs_phone" type="text" />
          </div>
          <div class="text-field">
            <label class="field-label" for="csAddressInput">수령인 주소</label>
            <input id="csAddressInput" name="cs_address" type="text" />
          </div>
          <div class="text-field">
            <label class="field-label" for="csTypeInput">CS타입</label>
            <select id="csTypeInput" name="cs_type">
              <option value="">선택</option>
              <option value="변심반품">변심반품</option>
              <option value="불량반품">불량반품</option>
              <option value="불량교환">불량교환</option>
              <option value="불량재출고(미회수)">불량재출고(미회수)</option>
              <option value="오출고(오배송)">오출고(오배송)</option>
            </select>
          </div>
          <div class="text-field">
            <label class="field-label" for="csContentInput">CS내용</label>
            <textarea id="csContentInput" name="cs_content" placeholder="예) 파손 / 오배송 / 누락 / 반품 접수 요청"></textarea>
          </div>
          <div class="text-field">
            <label class="field-label" for="csSubjectInput">메일 제목</label>
            <input id="csSubjectInput" name="subject" type="text" />
          </div>
          <div class="text-field">
            <label class="field-label" for="csBodyInput">요청 내용</label>
            <textarea id="csBodyInput" name="body"></textarea>
          </div>
          <div class="cs-toolbar">
            <button class="cs-save-button" id="saveCsCase" type="button">CS건 DB 저장</button>
          </div>
          <div class="cs-case-list">
            <div class="cs-case-head">최근 저장 CS</div>
            <div id="csCaseList"></div>
          </div>
        </div>
        <div class="ledger-fields" id="ledgerFields">
          <div class="ledger-toolbar">
            <input id="ledgerSearchInput" type="text" placeholder="업체명, 수령인, 상품명, 원송장, CS내용 검색" />
            <select id="ledgerStatusFilter">
              <option value="">상태 전체</option>
              <option value="회수지시">회수지시</option>
              <option value="회수 완료">회수 완료</option>
              <option value="재발송 완료">재발송 완료</option>
              <option value="전체 처리완료">전체 처리완료</option>
            </select>
            <select id="ledgerYearFilter">
              <option value="">년도별로 보기</option>
            </select>
            <select id="ledgerMonthFilter">
              <option value="">월별로 보기</option>
            </select>
            <button class="btn blue" id="ledgerRefresh" type="button">조회</button>
            <select id="ledgerPageSize">
              <option value="100">100개씩 보기</option>
              <option value="500">500개씩 보기</option>
              <option value="1000">1,000개씩 보기</option>
              <option value="2000">2,000개씩 보기</option>
              <option value="5000">5,000개씩 보기</option>
            </select>
            <span class="ledger-count" id="ledgerCountLabel">표시 0건</span>
            <button class="btn blue" id="ledgerDownloadExcel" type="button">엑셀 다운로드</button>
            <button class="btn primary" id="ledgerAddCs" type="button">CS 추가</button>
            <label class="ledger-import-button" for="ledgerImportInput">
              <i data-lucide="upload"></i>
              <span id="ledgerImportDropMain">업로드</span>
              <input id="ledgerImportInput" name="ledger_import" type="file" accept=".xlsx,.xlsm" />
            </label>
          </div>
          <div class="ledger-wrap">
            <table class="ledger-table">
              <thead>
                <tr>
                  <th>선택</th>
                  <th class="has-filter"><span class="ledger-th-title">날짜</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="occurred_at" data-label="날짜">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">매출거래처</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="sales_vendor" data-label="매출거래처">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">매입거래처</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="purchase_vendor" data-label="매입거래처">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">처리진행상태</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="status" data-label="처리진행상태">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">완료일</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="completed_at" data-label="완료일">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">처리내용</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="cs_type" data-label="처리내용">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">C/S 내용</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="cs_content" data-label="C/S 내용">▼</button></th>
                  <th class="invoice-head has-filter"><span class="ledger-th-title">재발송운송장번호</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="reship_invoice" data-label="재발송운송장번호">▼</button></th>
                  <th class="invoice-head has-filter"><span class="ledger-th-title">회수운송장번호</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="return_invoice" data-label="회수운송장번호">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">주문일자</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="order_date" data-label="주문일자">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">출고일</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="ship_date" data-label="출고일">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">주문자</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="orderer_name" data-label="주문자">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">연락처</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="orderer_phone" data-label="주문자 연락처">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">수령자</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="receiver_name" data-label="수령자">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">연락처</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="receiver_phone" data-label="수령자 연락처">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">제품명</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="product_name" data-label="제품명">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">수량</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="quantity" data-label="수량">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">상세주소</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="receiver_address" data-label="상세주소">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">택배사</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="courier" data-label="택배사">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">송장번호</span><button class="ledger-filter-trigger" type="button" data-ledger-filter-button="original_invoice" data-label="송장번호">▼</button></th>
                </tr>
              </thead>
              <tbody id="ledgerBody"></tbody>
            </table>
          </div>
        </div>
        <div class="management-fields" id="managementFields">
          <div class="ledger-toolbar">
            <input id="managementSearchInput" type="text" placeholder="거래처, 수령자, 상품명, 주소, 송장번호 검색" />
            <select id="managementYearFilter" hidden>
              <option value="">년도 선택</option>
            </select>
            <select id="managementMonthFilter">
              <option value="">월별로 보기</option>
            </select>
            <button class="btn blue" id="managementRefresh" type="button">조회</button>
            <select id="managementPageSize">
              <option value="100">100개씩 보기</option>
              <option value="500" selected>500개씩 보기</option>
              <option value="1000">1,000개씩 보기</option>
              <option value="2000">2,000개씩 보기</option>
              <option value="5000">5,000개씩 보기</option>
            </select>
            <div class="management-download-group">
              <button class="btn blue" id="managementDownloadSelected" type="button">선택 다운로드</button>
              <button class="btn blue" id="managementDownloadMonth" type="button">월별 다운로드</button>
              <button class="btn blue" id="managementDownloadYear" type="button">년별 다운로드</button>
              <button class="btn blue" id="managementDownloadAll" type="button">전체 다운로드</button>
            </div>
            <label class="ledger-import-button" for="managementImportInput">
              <i data-lucide="upload"></i>
              <span id="managementImportDropMain">업로드</span>
              <input id="managementImportInput" name="management_import" type="file" accept=".xlsx,.xlsm" />
            </label>
          </div>
          <div class="ledger-wrap management-wrap">
            <table class="ledger-table">
              <thead>
                <tr>
                  <th>선택</th>
                  <th class="has-filter"><span class="ledger-th-title">주문일자</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="order_date" data-label="주문일자">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">출고일</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="ship_date" data-label="출고일">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">매입거래처</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="purchase_vendor" data-label="매입거래처">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">매출거래처</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="sales_vendor" data-label="매출거래처">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">거래구분</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="transaction_type" data-label="거래구분">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">장부입력확인</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="ledger_checked" data-label="장부입력확인">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">주문자</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="orderer_name" data-label="주문자">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">발신자연락처</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="sender_phone" data-label="발신자연락처">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">수령자</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="receiver_name" data-label="수령자">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">수령자연락처</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="receiver_phone" data-label="수령자연락처">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">제품명</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="product_name" data-label="제품명">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">수량</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="quantity" data-label="수량">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">상세주소</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="receiver_address" data-label="상세주소">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">택배사</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="courier" data-label="택배사">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">운송장번호</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="invoice_number" data-label="운송장번호">▼</button></th>
                  <th class="has-filter"><span class="ledger-th-title">특이사항</span><button class="ledger-filter-trigger" type="button" data-management-filter-button="memo" data-label="특이사항">▼</button></th>
                  <th>CS접수</th>
                </tr>
              </thead>
              <tbody id="managementBody"></tbody>
            </table>
          </div>
          <div class="management-month-tabs" id="managementMonthTabs"></div>
        </div>
        <div class="ledger-filter-popover" id="ledgerFilterPopover">
          <div class="ledger-filter-title" id="ledgerFilterTitle">필터</div>
          <input class="ledger-filter-search" id="ledgerFilterSearch" type="text" placeholder="검색어 입력" />
          <div class="ledger-filter-option-list" id="ledgerFilterOptions"></div>
          <div class="ledger-filter-actions">
            <button type="button" id="ledgerFilterClear">전체</button>
            <button class="apply" type="button" id="ledgerFilterApply">적용</button>
          </div>
        </div>
        <div class="notice" id="notice"></div>
        <div class="modal-actions">
          <button class="btn" type="button" id="cancel">취소</button>
          <button class="btn primary" id="submitButton" type="submit">생성</button>
        </div>
      </form>
      <div class="result" id="result">
        <textarea id="resultText" readonly></textarea>
        <div class="result-actions">
          <button class="btn" type="button" id="copyResult">복사</button>
          <button class="btn blue" type="button" id="downloadText">텍스트 저장</button>
        </div>
      </div>
    </div>
  </div>

  <script type="module">
    import { createIcons, BriefcaseBusiness, Home, MessageCircle, Info, ChevronDown, ChevronRight, PlusSquare, RefreshCw, Ellipsis, Headphones, Package, ClipboardCheck, CircleDollarSign, FileText, FileSpreadsheet, ClipboardList, BarChart3, CopyCheck, Bell, Download, Truck, Mail, Upload, Database, CalendarDays, X } from "/lucide/dist/esm/lucide.js";
    createIcons({ icons: { BriefcaseBusiness, Home, MessageCircle, Info, ChevronDown, ChevronRight, PlusSquare, RefreshCw, Ellipsis, Headphones, Package, ClipboardCheck, CircleDollarSign, FileText, FileSpreadsheet, ClipboardList, BarChart3, CopyCheck, Bell, Download, Truck, Mail, Upload, Database, CalendarDays, X } });
    if (new URLSearchParams(window.location.search).get("standalone") === "1") {
      document.body.classList.add("standalone");
    }
    const currentUserPermissions = new Set(__USER_PERMISSIONS__);
    const permissionLabels = __PERMISSION_LABELS__;

    function can(permission) {
      return currentUserPermissions.has(permission);
    }

    function permissionLabel(permission) {
      return permissionLabels[permission] || permission;
    }

    function setHidden(element, hidden) {
      if (element) element.classList.toggle("permission-hidden", Boolean(hidden));
    }

    function applyStaticPermissions() {
      setHidden(document.querySelector("#noticeInputOpen"), !can("notice_manage"));
      setHidden(managementDeleteSelected, !can("ledger_delete"));
      setHidden(ledgerDeleteSelected, !can("ledger_delete"));
      setHidden(managementSaveAll, !can("ledger_edit"));
      setHidden(ledgerSaveAll, !can("ledger_edit"));
      setHidden(ledgerAddCs, !can("ledger_edit"));
      setHidden(saveCsCaseButton, !can("ledger_edit"));
      setHidden(document.querySelector("label[for='ledgerImportInput']"), !can("excel_upload"));
      setHidden(document.querySelector("label[for='managementImportInput']"), !can("excel_upload"));
      setHidden(ledgerDownloadExcel, !can("excel_download"));
      managementDownloadButtons.forEach((button) => setHidden(button, !can("excel_download")));
      setHidden(document.querySelector("label[for='vendorContactsFileInput']"), !can("excel_upload"));
      setHidden(saveVendorContactButton, !can("mail_send"));
      document.querySelectorAll('[data-open="cs"]').forEach((button) => setHidden(button, !can("mail_send")));
      if (!can("notice_manage")) {
        setHidden(noticeSaveButton, true);
        setHidden(noticeClearButton, true);
      }
    }

    const modal = document.querySelector("#modal");
    const modalTitle = document.querySelector("#modalTitle");
    const uploadForm = document.querySelector("#uploadForm");
    const fileLabel = document.querySelector("#fileLabel");
    const fileInput = document.querySelector("#fileInput");
    const templateInput = document.querySelector("#templateInput");
    const dropMain = document.querySelector("#dropMain");
    const dropSub = document.querySelector("#dropSub");
    const templateUpload = document.querySelector("#templateUpload");
    const templateDropMain = document.querySelector("#templateDropMain");
    const deliveryOptions = document.querySelector("#deliveryOptions");
    const vehicleFields = document.querySelector("#vehicleFields");
    const csFields = document.querySelector("#csFields");
    const ledgerFields = document.querySelector("#ledgerFields");
    const managementFields = document.querySelector("#managementFields");
    const productTable = document.querySelector("#productTable");
    const receiptTypeSelect = document.querySelector("#receiptTypeSelect");
    const supplierInput = document.querySelector("#supplierInput");
    const receiptDateInput = document.querySelector("#receiptDateInput");
    const freightPaymentSelect = document.querySelector("#freightPaymentSelect");
    const requestNoteInput = document.querySelector("#requestNoteInput");
    const deliveryPlaceInput = document.querySelector("#deliveryPlaceInput");
    const managerInput = document.querySelector("#managerInput");
    const naverEmailInput = document.querySelector("#naverEmailInput");
    const naverPasswordInput = document.querySelector("#naverPasswordInput");
    const saveMailCredentials = document.querySelector("#saveMailCredentials");
    const vendorContactSelect = document.querySelector("#vendorContactSelect");
    const vendorContactsFileInput = document.querySelector("#vendorContactsFileInput");
    const vendorContactsDropMain = document.querySelector("#vendorContactsDropMain");
    const recipientEmailInput = document.querySelector("#recipientEmailInput");
    const vendorNameInput = document.querySelector("#vendorNameInput");
    const saveVendorContactButton = document.querySelector("#saveVendorContact");
    const csOriginInput = document.querySelector("#csOriginInput");
    const csProductInput = document.querySelector("#csProductInput");
    const csReceiverInput = document.querySelector("#csReceiverInput");
    const csPhoneInput = document.querySelector("#csPhoneInput");
    const csAddressInput = document.querySelector("#csAddressInput");
    const csTypeInput = document.querySelector("#csTypeInput");
    const csContentInput = document.querySelector("#csContentInput");
    const csSubjectInput = document.querySelector("#csSubjectInput");
    const csBodyInput = document.querySelector("#csBodyInput");
    const saveCsCaseButton = document.querySelector("#saveCsCase");
    const csCaseList = document.querySelector("#csCaseList");
    const ledgerCsPopupClose = document.querySelector("#ledgerCsPopupClose");
    const ledgerSearchInput = document.querySelector("#ledgerSearchInput");
    const ledgerStatusFilter = document.querySelector("#ledgerStatusFilter");
    const ledgerYearFilter = document.querySelector("#ledgerYearFilter");
    const ledgerMonthFilter = document.querySelector("#ledgerMonthFilter");
    const ledgerRefresh = document.querySelector("#ledgerRefresh");
    const ledgerPageSize = document.querySelector("#ledgerPageSize");
    const ledgerDownloadExcel = document.querySelector("#ledgerDownloadExcel");
    const ledgerCountLabel = document.querySelector("#ledgerCountLabel");
    const ledgerAddCs = document.querySelector("#ledgerAddCs");
    const ledgerBody = document.querySelector("#ledgerBody");
    const ledgerImportInput = document.querySelector("#ledgerImportInput");
    const ledgerImportDropMain = document.querySelector("#ledgerImportDropMain");
    const managementSearchInput = document.querySelector("#managementSearchInput");
    const managementYearFilter = document.querySelector("#managementYearFilter");
    const managementMonthFilter = document.querySelector("#managementMonthFilter");
    const managementRefresh = document.querySelector("#managementRefresh");
    const managementPageSize = document.querySelector("#managementPageSize");
    const managementDownloadSelected = document.querySelector("#managementDownloadSelected");
    const managementDownloadMonth = document.querySelector("#managementDownloadMonth");
    const managementDownloadYear = document.querySelector("#managementDownloadYear");
    const managementDownloadAll = document.querySelector("#managementDownloadAll");
    const managementDownloadButtons = [
      managementDownloadSelected,
      managementDownloadMonth,
      managementDownloadYear,
      managementDownloadAll,
    ].filter(Boolean);
    const managementImportInput = document.querySelector("#managementImportInput");
    const managementImportDropMain = document.querySelector("#managementImportDropMain");
    const managementBody = document.querySelector("#managementBody");
    const managementMonthTabs = document.querySelector("#managementMonthTabs");
    const managementSaveAll = document.querySelector("#managementSaveAll");
    const ledgerSaveAll = document.querySelector("#ledgerSaveAll");
    const managementDeleteSelected = document.querySelector("#managementDeleteSelected");
    const ledgerDeleteSelected = document.querySelector("#ledgerDeleteSelected");
    const ledgerFilterButtons = Array.from(document.querySelectorAll("[data-ledger-filter-button]"));
    const managementFilterButtons = Array.from(document.querySelectorAll("[data-management-filter-button]"));
    const ledgerFilterPopover = document.querySelector("#ledgerFilterPopover");
    const ledgerFilterTitle = document.querySelector("#ledgerFilterTitle");
    const ledgerFilterSearch = document.querySelector("#ledgerFilterSearch");
    const ledgerFilterOptions = document.querySelector("#ledgerFilterOptions");
    const ledgerFilterClear = document.querySelector("#ledgerFilterClear");
    const ledgerFilterApply = document.querySelector("#ledgerFilterApply");
    const result = document.querySelector("#result");
    const resultText = document.querySelector("#resultText");
    const notice = document.querySelector("#notice");
    const submitButton = document.querySelector("#submitButton");
    const pageTitle = document.querySelector(".title");
    const dashboardContent = document.querySelector("#dashboardContent");
    const noticeDateInput = document.querySelector("#noticeDateInput");
    const noticeTitleInput = document.querySelector("#noticeTitleInput");
    const noticeOwnerInput = document.querySelector("#noticeOwnerInput");
    const noticeBodyInput = document.querySelector("#noticeBodyInput");
    const noticeSaveButton = document.querySelector("#noticeSaveButton");
    const noticeClearButton = document.querySelector("#noticeClearButton");
    const noticePreview = document.querySelector("#noticePreview");
    const sidebarNoticePreview = document.querySelector("#sidebarNoticePreview");
    const noticePopup = document.querySelector("#noticePopup");
    const noticePopupClose = document.querySelector("#noticePopupClose");
    const managementWorkspace = document.querySelector("#managementWorkspace");
    const ledgerWorkspace = document.querySelector("#ledgerWorkspace");
    const leaveWorkspace = document.querySelector("#leaveWorkspace");
    const userAdminWorkspace = document.querySelector("#userAdminWorkspace");
    const backupWorkspace = document.querySelector("#backupWorkspace");
    const managementWorkspaceMount = document.querySelector("#managementWorkspaceMount");
    const ledgerWorkspaceMount = document.querySelector("#ledgerWorkspaceMount");
    const userAdminRefresh = document.querySelector("#userAdminRefresh");
    const userAdminId = document.querySelector("#userAdminId");
    const userAdminUsername = document.querySelector("#userAdminUsername");
    const userAdminDisplayName = document.querySelector("#userAdminDisplayName");
    const userAdminRole = document.querySelector("#userAdminRole");
    const userAdminPassword = document.querySelector("#userAdminPassword");
    const userAdminActive = document.querySelector("#userAdminActive");
    const userAdminSave = document.querySelector("#userAdminSave");
    const userAdminBody = document.querySelector("#userAdminBody");
    const userAdminMessage = document.querySelector("#userAdminMessage");
    const userAdminPermissionChecks = Array.from(document.querySelectorAll("[data-permission-check]"));
    const leaveRefresh = document.querySelector("#leaveRefresh");
    const leaveTotalDays = document.querySelector("#leaveTotalDays");
    const leaveUsedDays = document.querySelector("#leaveUsedDays");
    const leaveRemainingDays = document.querySelector("#leaveRemainingDays");
    const leaveMessage = document.querySelector("#leaveMessage");
    const leaveBalanceBody = document.querySelector("#leaveBalanceBody");
    const leaveHistoryBody = document.querySelector("#leaveHistoryBody");
    const leaveApprovalBody = document.querySelector("#leaveApprovalBody");
    const leaveTypeSelect = document.querySelector("#leaveTypeSelect");
    const leaveUnitSelect = document.querySelector("#leaveUnitSelect");
    const leaveStartDate = document.querySelector("#leaveStartDate");
    const leaveEndDate = document.querySelector("#leaveEndDate");
    const leaveReasonInput = document.querySelector("#leaveReasonInput");
    const leaveRequestSubmit = document.querySelector("#leaveRequestSubmit");
    const leaveAdminUserSelect = document.querySelector("#leaveAdminUserSelect");
    const leaveAdminTotalInput = document.querySelector("#leaveAdminTotalInput");
    const leaveAdminUsedInput = document.querySelector("#leaveAdminUsedInput");
    const leaveBalanceSave = document.querySelector("#leaveBalanceSave");
    const leaveUsageUserSelect = document.querySelector("#leaveUsageUserSelect");
    const leaveUsageDatesInput = document.querySelector("#leaveUsageDatesInput");
    const leaveUsageNoteInput = document.querySelector("#leaveUsageNoteInput");
    const leaveUsageSave = document.querySelector("#leaveUsageSave");
    const leaveAdminBalanceBody = document.querySelector("#leaveAdminBalanceBody");
    const leaveTabs = Array.from(document.querySelectorAll("[data-leave-tab]"));
    const backupRefresh = document.querySelector("#backupRefresh");
    const backupCreate = document.querySelector("#backupCreate");
    const backupRestoreInput = document.querySelector("#backupRestoreInput");
    const backupPath = document.querySelector("#backupPath");
    const backupBody = document.querySelector("#backupBody");
    const backupMessage = document.querySelector("#backupMessage");
    const systemUpdateWorkspace = document.querySelector("#systemUpdateWorkspace");
    const systemUpdateRefresh = document.querySelector("#systemUpdateRefresh");
    const systemUpdateCheck = document.querySelector("#systemUpdateCheck");
    const systemUpdateApply = document.querySelector("#systemUpdateApply");
    const systemUpdateSource = document.querySelector("#systemUpdateSource");
    const systemUpdateBranch = document.querySelector("#systemUpdateBranch");
    const systemUpdateCurrent = document.querySelector("#systemUpdateCurrent");
    const systemUpdateState = document.querySelector("#systemUpdateState");
    const systemUpdateMessage = document.querySelector("#systemUpdateMessage");
    const systemUpdateHistoryBody = document.querySelector("#systemUpdateHistoryBody");
    let currentMode = "dashboard";
    let vendorContacts = [];
    let ledgerCases = [];
    let managementRecords = [];
    let managementPeriods = [];
    let userAccounts = [];
    let activeLedgerFilterField = "";
    let activeManagementFilterField = "";
    const ledgerFilters = {};
    const managementFilters = {};
    let isBulkSaving = false;

    if (managementWorkspaceMount && managementFields) managementWorkspaceMount.appendChild(managementFields);
    if (ledgerWorkspaceMount && ledgerFields) ledgerWorkspaceMount.appendChild(ledgerFields);
    if (ledgerFilterPopover) document.body.appendChild(ledgerFilterPopover);
    fillPeriodSelects(ledgerYearFilter, ledgerMonthFilter);
    renderManagementPeriodControls();
    applyStaticPermissions();
    loadNoticeTemplate();

    function addProductRow(productName = "", quantity = "", packQuantity = "") {
      const row = document.createElement("div");
      row.className = "product-row";
      row.innerHTML = `
        <input class="product-name" type="text" placeholder="제품명" value="${productName}">
        <input class="product-quantity" type="text" placeholder="수량" value="${quantity}">
        <input class="product-pack-quantity" type="text" placeholder="입수량" value="${packQuantity}">
      `;
      productTable.appendChild(row);
    }

    function defaultProductRowCount() {
      return receiptTypeSelect.value === "모드니 전용" ? 15 : 5;
    }

    function resetProductRows() {
      productTable.innerHTML = "";
      for (let i = 0; i < defaultProductRowCount(); i += 1) addProductRow();
    }

    function todayString() {
      const now = new Date();
      const offset = now.getTimezoneOffset() * 60000;
      return new Date(now.getTime() - offset).toISOString().slice(0, 10);
    }

    function fillPeriodSelects(yearSelect, monthSelect) {
      const currentYear = new Date().getFullYear();
      const startYear = 2023;
      for (let year = currentYear + 1; year >= startYear; year -= 1) {
        const option = document.createElement("option");
        option.value = String(year);
        option.textContent = `${year}년`;
        yearSelect.appendChild(option);
      }
      for (let month = 1; month <= 12; month += 1) {
        const option = document.createElement("option");
        option.value = String(month).padStart(2, "0");
        option.textContent = `${month}월`;
        monthSelect.appendChild(option);
      }
    }

    function ensureYearForMonth(yearSelect, monthSelect) {
      if (monthSelect.value && !yearSelect.value) yearSelect.value = String(new Date().getFullYear());
    }

    function managementPeriodYears() {
      return Array.from(new Set(managementPeriods.map((period) => String(period.year || "")).filter(Boolean)))
        .sort((left, right) => right.localeCompare(left));
    }

    function managementMonthsForYear(year) {
      return managementPeriods
        .filter((period) => String(period.year || "") === String(year || ""))
        .map((period) => String(period.month || "").padStart(2, "0"))
        .filter(Boolean)
        .sort();
    }

    function selectedManagementPeriod() {
      const selectedOption = managementMonthFilter?.selectedOptions?.[0];
      const optionYear = selectedOption?.dataset?.year || "";
      const optionMonth = selectedOption?.dataset?.month || "";
      if (optionYear && optionMonth) return { year: optionYear, month: optionMonth };
      const value = managementMonthFilter?.value || "";
      if (!value) return { year: "", month: "" };
      const match = value.match(/^(\d{4})-(\d{2})$/);
      if (match) return { year: match[1], month: match[2] };
      return { year: managementYearFilter?.value || "", month: "" };
    }

    function setManagementPeriod(year, month) {
      const normalizedYear = String(year || "");
      const normalizedMonth = String(month || "").padStart(2, "0");
      managementYearFilter.value = normalizedYear;
      managementMonthFilter.value = normalizedYear && normalizedMonth ? `${normalizedYear}-${normalizedMonth}` : "";
    }

    function renderManagementPeriodControls() {
      if (!managementYearFilter || !managementMonthFilter) return;
      const years = managementPeriodYears();
      const previousPeriod = selectedManagementPeriod();
      managementYearFilter.innerHTML = `<option value="">년도 선택</option>${years.map((year) => (
        `<option value="${escapeHtml(year)}">${escapeHtml(year)}년</option>`
      )).join("")}`;
      if (previousPeriod.year && years.includes(previousPeriod.year)) managementYearFilter.value = previousPeriod.year;
      else managementYearFilter.value = "";

      const periodOptions = managementPeriods.map((period) => {
        const year = String(period.year || "");
        const month = String(period.month || "").padStart(2, "0");
        return `<option value="${escapeHtml(`${year}-${month}`)}" data-year="${escapeHtml(year)}" data-month="${escapeHtml(month)}">${escapeHtml(year)}년 ${Number(month)}월</option>`;
      }).join("");
      managementMonthFilter.innerHTML = `<option value="">월별로 보기</option>${periodOptions}`;
      if (previousPeriod.year && previousPeriod.month) {
        const value = `${previousPeriod.year}-${previousPeriod.month}`;
        managementMonthFilter.value = Array.from(managementMonthFilter.options).some((option) => option.value === value) ? value : "";
      } else {
        managementMonthFilter.value = "";
      }

      renderManagementMonthTabs();
    }

    function renderManagementMonthTabs() {
      if (!managementMonthTabs) return;
      const activePeriod = selectedManagementPeriod();
      const activeMonth = activePeriod.month || "";
      const activeYear = activePeriod.year || managementPeriodYears()[0] || "";
      const months = activeYear ? managementMonthsForYear(activeYear) : [];
      managementMonthTabs.innerHTML = months.length ? months.map((month) => {
        const active = month === activeMonth ? " active" : "";
        return `<button class="management-month-tab${active}" type="button" data-management-year="${escapeHtml(activeYear)}" data-management-month="${escapeHtml(month)}">${escapeHtml(activeYear)}년 ${Number(month)}월</button>`;
      }).join("") : `<button class="management-month-tab" type="button" disabled>생성된 월 없음</button>`;
    }

    function roleText(role) {
      return role === "admin" ? "관리자" : "사용자";
    }

    function resetUserAdminForm() {
      if (!userAdminId) return;
      userAdminId.value = "";
      userAdminUsername.value = "";
      userAdminUsername.disabled = false;
      userAdminDisplayName.value = "";
      userAdminRole.value = "user";
      userAdminPassword.value = "";
      userAdminActive.checked = true;
      userAdminPermissionChecks.forEach((checkbox) => {
        checkbox.checked = ["ledger_edit", "excel_download", "cs_receive", "leave_view"].includes(checkbox.value);
      });
      userAdminMessage.textContent = "신규 사용자는 아이디와 비밀번호를 입력한 뒤 저장하세요.";
    }

    function syncPermissionChecksForRole() {
      if (!userAdminRole) return;
      if (userAdminRole.value === "admin") {
        userAdminPermissionChecks.forEach((checkbox) => {
          checkbox.checked = true;
        });
      }
    }

    function renderUserAccounts() {
      if (!userAdminBody) return;
      if (!userAccounts.length) {
        userAdminBody.innerHTML = `<tr><td colspan="7">등록된 사용자가 없습니다.</td></tr>`;
        return;
      }
      userAdminBody.innerHTML = userAccounts.map((user) => `
        <tr data-user-id="${user.id}">
          <td>${escapeHtml(user.username)}</td>
          <td>${escapeHtml(user.display_name)}</td>
          <td>${roleText(user.role)}</td>
          <td>${(user.permissions || []).map((permission) => permissionLabel(permission)).join(", ")}</td>
          <td>${user.active ? "사용" : "중지"}</td>
          <td>${escapeHtml(user.created_at || "")}</td>
          <td><button class="admin-action" type="button" data-user-edit="${user.id}">수정</button></td>
        </tr>
      `).join("");
    }

    function editUserAccount(userId) {
      const user = userAccounts.find((item) => String(item.id) === String(userId));
      if (!user || !userAdminId) return;
      userAdminId.value = user.id;
      userAdminUsername.value = user.username || "";
      userAdminUsername.disabled = false;
      userAdminDisplayName.value = user.display_name || "";
      userAdminRole.value = user.role || "user";
      userAdminPassword.value = "";
      userAdminActive.checked = Boolean(user.active);
      const permissions = new Set(user.permissions || []);
      userAdminPermissionChecks.forEach((checkbox) => {
        checkbox.checked = permissions.has(checkbox.value);
      });
      userAdminMessage.textContent = `${user.username} 계정 수정 중입니다. 비밀번호는 변경할 때만 입력하세요.`;
      userAdminUsername.focus();
    }

    async function loadUserAccounts() {
      if (!userAdminBody) return;
      userAdminMessage.textContent = "사용자 목록을 불러오는 중입니다.";
      try {
        const response = await fetch("/api/users");
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "사용자 목록을 불러오지 못했습니다.");
        userAccounts = data.users || [];
        renderUserAccounts();
        if (!userAdminId.value) resetUserAdminForm();
      } catch (error) {
        userAdminMessage.textContent = error.message;
      }
    }

    async function saveUserAccount() {
      if (!userAdminSave) return;
      const payload = {
        id: userAdminId.value,
        username: userAdminUsername.value.trim(),
        display_name: userAdminDisplayName.value.trim(),
        role: userAdminRole.value,
        password: userAdminPassword.value,
        active: userAdminActive.checked,
        permissions: userAdminPermissionChecks.filter((checkbox) => checkbox.checked).map((checkbox) => checkbox.value),
      };
      userAdminMessage.textContent = "사용자 계정을 저장하는 중입니다.";
      userAdminSave.disabled = true;
      try {
        const response = await fetch("/api/users-save", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "사용자 계정을 저장하지 못했습니다.");
        userAccounts = data.users || [];
        renderUserAccounts();
        resetUserAdminForm();
        userAdminMessage.textContent = data.message || "사용자 계정을 저장했습니다.";
      } catch (error) {
        userAdminMessage.textContent = error.message;
      } finally {
        userAdminSave.disabled = false;
      }
    }

    function backupSizeText(size) {
      const number = Number(size || 0);
      if (number >= 1024 * 1024) return `${(number / 1024 / 1024).toFixed(1)} MB`;
      if (number >= 1024) return `${(number / 1024).toFixed(1)} KB`;
      return `${number} B`;
    }

    function renderBackups(data) {
      if (!backupBody) return;
      backupPath.textContent = data.backup_dir || "-";
      const backups = data.backups || [];
      if (!backups.length) {
        backupBody.innerHTML = `<tr><td colspan="4">생성된 백업 파일이 없습니다.</td></tr>`;
        return;
      }
      backupBody.innerHTML = backups.map((backup) => `
        <tr>
          <td>${escapeHtml(backup.name)}</td>
          <td>${escapeHtml(backup.created_at || "")}</td>
          <td>${backupSizeText(backup.size)}</td>
          <td>
            <div class="backup-actions">
              <button class="backup-action" type="button" data-backup-restore="${escapeHtml(backup.name)}">복원</button>
              <button class="backup-action" type="button" data-backup-download="${escapeHtml(backup.name)}">다운로드</button>
              <button class="backup-action danger" type="button" data-backup-delete="${escapeHtml(backup.name)}">삭제</button>
            </div>
          </td>
        </tr>
      `).join("");
    }

    async function loadBackups() {
      if (!backupWorkspace) return;
      backupMessage.textContent = "백업 목록을 불러오는 중입니다.";
      try {
        const response = await fetch("/api/backups");
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "백업 목록을 불러오지 못했습니다.");
        renderBackups(data);
        backupMessage.textContent = data.last_backup ? `마지막 백업: ${data.last_backup}` : "아직 백업이 없습니다.";
      } catch (error) {
        backupMessage.textContent = error.message;
      }
    }

    async function createBackupNow() {
      if (!backupCreate) return;
      backupCreate.disabled = true;
      backupMessage.textContent = "백업 파일을 생성하는 중입니다.";
      try {
        const response = await fetch("/api/backup-create", { method: "POST" });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "백업 생성에 실패했습니다.");
        backupMessage.textContent = data.message || "백업 파일을 생성했습니다.";
        await loadBackups();
      } catch (error) {
        backupMessage.textContent = error.message;
      } finally {
        backupCreate.disabled = false;
      }
    }

    function downloadBackup(name) {
      const url = `/api/backup-download?name=${encodeURIComponent(name)}`;
      window.location.href = url;
    }

    async function deleteBackup(name) {
      if (!confirm(`${name} 백업 파일을 삭제할까요?`)) return;
      backupMessage.textContent = "백업 파일을 삭제하는 중입니다.";
      try {
        const response = await fetch("/api/backup-delete", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ name }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "백업 삭제에 실패했습니다.");
        backupMessage.textContent = data.message || "백업 파일을 삭제했습니다.";
        await loadBackups();
      } catch (error) {
        backupMessage.textContent = error.message;
      }
    }

    function restoreCompleteMessage(data) {
      return `${data.message || "백업 데이터 복원이 완료되었습니다."} 프로그램을 다시 실행하거나 로그아웃 후 다시 로그인해주세요.`;
    }

    async function restoreBackupByName(name) {
      if (!confirm(`${name} 백업 데이터로 현재 업무 데이터를 복원할까요?\n\n복원 전 현재 데이터는 자동으로 예비 백업됩니다.`)) return;
      backupMessage.textContent = "백업 데이터를 복원하는 중입니다.";
      try {
        const response = await fetch("/api/backup-restore", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ name }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "백업 복원에 실패했습니다.");
        backupMessage.textContent = restoreCompleteMessage(data);
      } catch (error) {
        backupMessage.textContent = error.message;
      }
    }

    async function restoreBackupFromUpload() {
      if (!backupRestoreInput || !backupRestoreInput.files[0]) return;
      const file = backupRestoreInput.files[0];
      if (!confirm(`${file.name} 파일로 현재 업무 데이터를 복원할까요?\n\n복원 전 현재 데이터는 자동으로 예비 백업됩니다.`)) {
        backupRestoreInput.value = "";
        return;
      }
      const formData = new FormData();
      formData.append("file", file);
      backupMessage.textContent = "업로드한 백업 데이터를 복원하는 중입니다.";
      try {
        const response = await fetch("/api/backup-restore-upload", {
          method: "POST",
          body: formData,
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "백업 복원에 실패했습니다.");
        backupMessage.textContent = restoreCompleteMessage(data);
      } catch (error) {
        backupMessage.textContent = error.message;
      } finally {
        backupRestoreInput.value = "";
      }
    }

    function renderSystemUpdate(data) {
      if (!systemUpdateWorkspace) return;
      const status = data.status || {};
      systemUpdateSource.textContent = status.source_dir || "Git 저장소 연결 없음";
      systemUpdateBranch.textContent = status.branch || "-";
      systemUpdateCurrent.textContent = status.current_short || "-";
      if (!status.available) {
        systemUpdateState.textContent = status.message || "Git 저장소 연결 없음";
      } else if (status.behind > 0) {
        systemUpdateState.textContent = `${status.behind}개 업데이트 가능`;
      } else {
        systemUpdateState.textContent = "최신 상태";
      }
      systemUpdateMessage.textContent = status.message || "";
      const history = data.history || [];
      systemUpdateHistoryBody.innerHTML = history.length
        ? history.map((row) => `
          <tr>
            <td>${escapeHtml(row.created_at || "")}</td>
            <td>${escapeHtml(row.action || "")}</td>
            <td>${escapeHtml(row.status || "")}</td>
            <td>${escapeHtml(row.before_commit || "")}</td>
            <td>${escapeHtml(row.after_commit || "")}</td>
            <td>${escapeHtml(row.backup_name || "")}</td>
            <td>${escapeHtml(row.message || "")}</td>
          </tr>
        `).join("")
        : `<tr><td colspan="7">업데이트 이력이 없습니다.</td></tr>`;
      if (systemUpdateApply) systemUpdateApply.disabled = !status.available || Boolean(status.dirty);
      if (systemUpdateCheck) systemUpdateCheck.disabled = !status.available;
    }

    async function loadSystemUpdateStatus() {
      if (!systemUpdateWorkspace) return;
      systemUpdateMessage.textContent = "업데이트 정보를 불러오는 중입니다.";
      try {
        const response = await fetch("/api/system-update");
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "업데이트 정보를 불러오지 못했습니다.");
        renderSystemUpdate(data);
      } catch (error) {
        systemUpdateMessage.textContent = error.message;
      }
    }

    async function checkSystemUpdate() {
      if (!systemUpdateCheck) return;
      systemUpdateCheck.disabled = true;
      systemUpdateMessage.textContent = "GitHub 업데이트를 확인하는 중입니다.";
      try {
        const response = await fetch("/api/system-update-check", { method: "POST" });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "업데이트 확인에 실패했습니다.");
        renderSystemUpdate(data);
      } catch (error) {
        systemUpdateMessage.textContent = error.message;
      } finally {
        systemUpdateCheck.disabled = false;
      }
    }

    async function applySystemUpdate() {
      if (!confirm("업데이트 적용 전 현재 데이터를 자동 백업합니다.\nGitHub 최신 코드를 적용할까요?")) return;
      systemUpdateApply.disabled = true;
      systemUpdateMessage.textContent = "업데이트 적용 중입니다. 창을 닫지 마세요.";
      try {
        const response = await fetch("/api/system-update-apply", { method: "POST" });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "업데이트 적용에 실패했습니다.");
        renderSystemUpdate(data);
        systemUpdateMessage.textContent = `${data.message || "업데이트 적용이 완료되었습니다."} 프로그램/나스 서비스를 다시 시작해주세요.`;
      } catch (error) {
        systemUpdateMessage.textContent = error.message;
      } finally {
        systemUpdateApply.disabled = false;
      }
    }

    function dayText(value) {
      const number = Number(value || 0);
      return `${Number.isInteger(number) ? number : number.toFixed(1)}일`;
    }

    function renderLeaveSelectOptions(select, rows, valueField = "id", labelField = "name") {
      if (!select) return;
      select.innerHTML = rows.map((row) => `<option value="${row[valueField]}">${escapeHtml(row[labelField])}</option>`).join("");
    }

    function setLeaveTab(tabName) {
      leaveTabs.forEach((button) => button.classList.toggle("active", button.dataset.leaveTab === tabName));
      ["mine", "request", "approvals", "admin"].forEach((name) => {
        const panel = document.querySelector(`#leaveTab${name[0].toUpperCase()}${name.slice(1)}`);
        if (panel) panel.classList.toggle("active", name === tabName);
      });
    }

    function renderLeaveData(data) {
      if (!leaveWorkspace) return;
      leaveTotalDays.textContent = dayText(data.summary?.total_days);
      leaveUsedDays.textContent = dayText(data.summary?.used_days);
      leaveRemainingDays.textContent = dayText(data.summary?.remaining_days);
      leaveBalanceBody.innerHTML = (data.balances || []).length
        ? data.balances.map((row) => `
          <tr><td>${escapeHtml(row.name)}</td><td>${dayText(row.total_days)}</td><td>${dayText(row.used_days)}</td><td><strong>${dayText(row.remaining_days)}</strong></td></tr>
        `).join("")
        : `<tr><td colspan="4">연차 기준이 아직 설정되지 않았습니다.</td></tr>`;
      leaveHistoryBody.innerHTML = (data.requests || []).length
        ? data.requests.map((row) => `
          <tr>
            <td>${escapeHtml(row.start_date)}${row.start_date === row.end_date ? "" : ` ~ ${escapeHtml(row.end_date)}`}</td>
            <td>${escapeHtml(row.unit_label)}</td>
            <td>${dayText(row.requested_days)}</td>
            <td>${escapeHtml(row.status_label)}</td>
            <td>${escapeHtml(row.reason)}</td>
          </tr>
        `).join("")
        : `<tr><td colspan="5">연차 사용/신청 이력이 없습니다.</td></tr>`;
      renderLeaveSelectOptions(leaveTypeSelect, data.leave_types || []);
      const userOptions = (data.users || []).map((user) => ({
        id: user.id,
        name: `${user.display_name} (${user.username})`,
      }));
      renderLeaveSelectOptions(leaveAdminUserSelect, userOptions);
      renderLeaveSelectOptions(leaveUsageUserSelect, userOptions);
      leaveApprovalBody.innerHTML = (data.pending_requests || []).length
        ? data.pending_requests.map((row) => `
          <tr>
            <td>${escapeHtml(row.requester)}</td>
            <td>${escapeHtml(row.start_date)}${row.start_date === row.end_date ? "" : ` ~ ${escapeHtml(row.end_date)}`}</td>
            <td>${escapeHtml(row.unit_label)}</td>
            <td>${dayText(row.requested_days)}</td>
            <td>${escapeHtml(row.reason)}</td>
            <td>
              <div class="leave-action-row">
                <button class="leave-action approve" type="button" data-leave-decision="approve" data-leave-id="${row.id}">승인</button>
                <button class="leave-action reject" type="button" data-leave-decision="reject" data-leave-id="${row.id}">반려</button>
              </div>
            </td>
          </tr>
        `).join("")
        : `<tr><td colspan="6">승인 대기 중인 연차 신청이 없습니다.</td></tr>`;
      leaveAdminBalanceBody.innerHTML = (data.admin_balances || []).length
        ? data.admin_balances.map((row) => `
          <tr><td>${escapeHtml(row.display_name)}</td><td>${escapeHtml(row.username)}</td><td>${dayText(row.total_days)}</td><td>${dayText(row.used_days)}</td><td><strong>${dayText(row.remaining_days)}</strong></td></tr>
        `).join("")
        : `<tr><td colspan="5">직원 연차 현황이 없습니다.</td></tr>`;
      document.querySelector('[data-leave-tab="approvals"]')?.classList.toggle("permission-hidden", !data.can_approve);
      document.querySelector('[data-leave-tab="admin"]')?.classList.toggle("permission-hidden", !data.can_manage);
    }

    async function loadLeaveData() {
      if (!leaveWorkspace) return;
      leaveMessage.textContent = "연차 정보를 불러오는 중입니다.";
      try {
        const response = await fetch("/api/leaves");
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "연차 정보를 불러오지 못했습니다.");
        renderLeaveData(data);
        leaveMessage.textContent = "";
      } catch (error) {
        leaveMessage.textContent = error.message;
      }
    }

    async function submitLeaveRequest() {
      leaveMessage.textContent = "연차 신청을 저장하는 중입니다.";
      try {
        const payload = {
          leave_type_id: leaveTypeSelect.value,
          unit: leaveUnitSelect.value,
          start_date: leaveStartDate.value,
          end_date: leaveEndDate.value,
          reason: leaveReasonInput.value.trim(),
        };
        const response = await fetch("/api/leave-request", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "연차 신청에 실패했습니다.");
        leaveReasonInput.value = "";
        leaveMessage.textContent = data.message || "연차 신청이 저장되었습니다.";
        await loadLeaveData();
        setLeaveTab("mine");
      } catch (error) {
        leaveMessage.textContent = error.message;
      }
    }

    function syncHalfDayDates() {
      if (!leaveUnitSelect || !leaveStartDate || !leaveEndDate) return;
      const isHalfDay = leaveUnitSelect.value === "HALF_DAY";
      if (isHalfDay && leaveStartDate.value) {
        leaveEndDate.value = leaveStartDate.value;
      }
      leaveEndDate.disabled = isHalfDay;
    }

    async function decideLeaveRequest(requestId, decision) {
      const comment = decision === "reject" ? window.prompt("반려 사유를 입력해주세요.", "반려") || "반려" : "";
      leaveMessage.textContent = "연차 신청을 처리하는 중입니다.";
      try {
        const response = await fetch("/api/leave-decision", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ request_id: requestId, decision, comment }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "연차 신청 처리에 실패했습니다.");
        leaveMessage.textContent = data.message || "처리되었습니다.";
        await loadLeaveData();
      } catch (error) {
        leaveMessage.textContent = error.message;
      }
    }

    async function saveLeaveBalance() {
      leaveMessage.textContent = "직원 연차 기준을 저장하는 중입니다.";
      try {
        const response = await fetch("/api/leave-balance", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            user_id: leaveAdminUserSelect.value,
            total_days: leaveAdminTotalInput.value,
            used_days: leaveAdminUsedInput.value,
          }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "연차 기준 저장에 실패했습니다.");
        leaveMessage.textContent = data.message || "연차 기준을 저장했습니다.";
        await loadLeaveData();
      } catch (error) {
        leaveMessage.textContent = error.message;
      }
    }

    async function saveHistoricalLeaveUsage() {
      leaveMessage.textContent = "기존 사용 연차를 등록하는 중입니다.";
      try {
        const response = await fetch("/api/leave-historical", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            user_id: leaveUsageUserSelect.value,
            usage_dates: leaveUsageDatesInput.value,
            note: leaveUsageNoteInput.value,
          }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "기존 사용 연차 등록에 실패했습니다.");
        leaveUsageDatesInput.value = "";
        leaveUsageNoteInput.value = "";
        leaveMessage.textContent = data.message || "기존 사용 연차를 등록했습니다.";
        await loadLeaveData();
      } catch (error) {
        leaveMessage.textContent = error.message;
      }
    }

    function loadNoticeTemplate() {
      let saved = {};
      try {
        saved = JSON.parse(localStorage.getItem("workhub_notice_template") || "{}");
      } catch {
        saved = {};
      }
      noticeDateInput.value = saved.date || todayString();
      noticeTitleInput.value = saved.title || "";
      noticeOwnerInput.value = saved.owner || "";
      noticeBodyInput.value = saved.body || "";
      renderNoticePreview();
    }

    function noticePayload() {
      return {
        date: noticeDateInput.value || todayString(),
        title: noticeTitleInput.value.trim(),
        owner: noticeOwnerInput.value.trim(),
        body: noticeBodyInput.value.trim(),
      };
    }

    function renderNoticePreview() {
      const payload = noticePayload();
      if (!payload.title && !payload.body) {
        noticePreview.innerHTML = `<strong>저장된 공지사항이 없습니다.</strong>공지사항을 입력하고 저장하면 이곳에서 미리 볼 수 있습니다.`;
        sidebarNoticePreview.innerHTML = `
          <div class="notice-board-kicker">금일 공지사항</div>
          <div class="notice-board-title">등록된 공지 없음</div>
          <div class="notice-board-body">공지사항 입력 메뉴를 눌러 내용을 입력해주세요.</div>
        `;
        return;
      }
      const meta = [shortKoreanDate(payload.date), payload.owner ? `담당 ${escapeHtml(payload.owner)}` : ""].filter(Boolean).join(" / ");
      noticePreview.innerHTML = `
        <strong>${escapeHtml(payload.title || "제목 없음")}</strong>
        <div>${escapeHtml(meta)}</div>
        <div>${escapeHtml(payload.body).replaceAll("\n", "<br>")}</div>
      `;
      sidebarNoticePreview.innerHTML = `
        <div class="notice-board-kicker">금일 공지사항</div>
        <div class="notice-board-title">${escapeHtml(payload.title || "제목 없음")}</div>
        <div class="notice-board-meta">${escapeHtml(meta)}</div>
        <div class="notice-board-body">${escapeHtml(payload.body || "내용 없음")}</div>
      `;
    }

    function openNoticePopup() {
      if (!can("notice_manage")) {
        notice.textContent = "공지사항 관리 권한이 없습니다.";
        return;
      }
      loadNoticeTemplate();
      noticePopup.classList.add("open");
      setTimeout(() => noticeTitleInput.focus(), 0);
    }

    function closeNoticePopup() {
      noticePopup.classList.remove("open");
    }

    function saveNoticeTemplate() {
      if (!can("notice_manage")) {
        notice.textContent = "공지사항 관리 권한이 없습니다.";
        return;
      }
      const payload = noticePayload();
      localStorage.setItem("workhub_notice_template", JSON.stringify(payload));
      renderNoticePreview();
      notice.textContent = "공지사항을 저장했습니다.";
      closeNoticePopup();
    }

    function clearNoticeTemplate() {
      if (!can("notice_manage")) {
        notice.textContent = "공지사항 관리 권한이 없습니다.";
        return;
      }
      localStorage.removeItem("workhub_notice_template");
      noticeTitleInput.value = "";
      noticeOwnerInput.value = "";
      noticeBodyInput.value = "";
      noticeDateInput.value = todayString();
      renderNoticePreview();
      notice.textContent = "공지사항 입력 내용을 초기화했습니다.";
    }

    function collectVehiclePayload() {
      const items = [...productTable.querySelectorAll(".product-row")].map((row) => ({
        product_name: row.querySelector(".product-name").value.trim(),
        quantity: row.querySelector(".product-quantity").value.trim(),
        pack_quantity: row.querySelector(".product-pack-quantity").value.trim(),
      })).filter((item) => item.product_name || item.quantity || item.pack_quantity);
      return {
        receipt_type: receiptTypeSelect.value,
        supplier: supplierInput.value.trim(),
        receipt_date: receiptDateInput.value,
        freight_payment: freightPaymentSelect.value,
        request_note: requestNoteInput.value.trim(),
        delivery_place: deliveryPlaceInput.value.trim(),
        manager: managerInput.value.trim(),
        items,
      };
    }

    function defaultCsSubject(vendorName = "") {
      return `[CS 요청] ${vendorName ? `${vendorName} ` : ""}확인 부탁드립니다`;
    }

    function defaultCsBody() {
      return `안녕하세요. (주)소일브릿지 입니다.\n\n\n\n- 원출고일  및 원송장번호 : ${csOriginInput.value.trim()}\n\n\n\n- 상품명 : ${csProductInput.value.trim()} \n\n- 수령인 : ${csReceiverInput.value.trim()}\n\n- 수령인 연락처 : ${csPhoneInput.value.trim()}\n\n- 수령인 주소 : ${csAddressInput.value.trim()}\n\n\n\n- CS내용 : ${csContentInput.value.trim()}\n\n \n\nCS건을 보내드립니다.\n\n \n\n★반품 접수 후 일주일 이상 회신 없으실 경우 자체 환불 및 정산 반영 예정이오니, 처리 결과 꼭 회신 부탁드립니다.\n\n `;
    }

    function refreshCsBody() {
      csBodyInput.value = defaultCsBody();
    }

    async function loadMailSettings() {
      try {
        const response = await fetch("/api/mail-settings");
        if (!response.ok) return;
        const data = await response.json();
        if (currentMode !== "cs") return;
        naverEmailInput.value = data.naver_email || "";
        naverPasswordInput.placeholder = data.has_password ? "저장된 비밀번호 사용" : "저장된 비밀번호가 없으면 입력";
      } catch {
        // 저장된 메일 설정이 없어도 CS 작성은 계속 가능합니다.
      }
    }

    function renderVendorContacts() {
      vendorContactSelect.innerHTML = "";
      const emptyOption = document.createElement("option");
      emptyOption.value = "";
      emptyOption.textContent = vendorContacts.length ? "업체를 선택해주세요" : "저장된 업체가 없습니다";
      vendorContactSelect.appendChild(emptyOption);

      vendorContacts.forEach((contact) => {
        const option = document.createElement("option");
        option.value = contact.vendor_name;
        option.textContent = `${contact.vendor_name} / ${contact.email}`;
        vendorContactSelect.appendChild(option);
      });
    }

    async function loadVendorContacts() {
      try {
        const response = await fetch("/api/vendor-contacts");
        if (!response.ok) return;
        const data = await response.json();
        vendorContacts = data.contacts || [];
        renderVendorContacts();
      } catch {
        vendorContacts = [];
        renderVendorContacts();
      }
    }

    function applySelectedVendor() {
      const selected = vendorContacts.find((contact) => contact.vendor_name === vendorContactSelect.value);
      if (!selected) return;
      vendorNameInput.value = selected.vendor_name;
      recipientEmailInput.value = selected.email;
      csSubjectInput.value = defaultCsSubject(selected.vendor_name);
    }

    async function saveCurrentVendorContact() {
      const vendorName = vendorNameInput.value.trim();
      const email = recipientEmailInput.value.trim();
      if (!vendorName || !email) {
        notice.textContent = "업체명과 받는 업체 메일을 입력해주세요.";
        return;
      }
      try {
        saveVendorContactButton.disabled = true;
        const response = await fetch("/api/vendor-contact", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ vendor_name: vendorName, email }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "업체 메일 저장에 실패했습니다.");
        vendorContacts = data.contacts || [];
        renderVendorContacts();
        vendorContactSelect.value = vendorName;
        notice.textContent = "업체 메일 주소를 저장했습니다.";
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        saveVendorContactButton.disabled = false;
      }
    }

    async function uploadVendorContactsWorkbook() {
      const file = vendorContactsFileInput.files[0];
      if (!file) return;
      vendorContactsDropMain.textContent = file.name;
      const formData = new FormData();
      formData.append("file", file);
      notice.textContent = "업체 메일 주소록을 저장 중입니다.";
      try {
        const response = await fetch("/api/vendor-contacts-import", {
          method: "POST",
          body: formData,
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "업체 메일 주소록 저장에 실패했습니다.");
        vendorContacts = data.contacts || [];
        renderVendorContacts();
        notice.textContent = data.message || "업체 메일 주소록을 저장했습니다.";
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        vendorContactsFileInput.value = "";
      }
    }

    function collectCsPayload() {
      return {
        naver_email: naverEmailInput.value.trim(),
        naver_password: naverPasswordInput.value,
        save_credentials: saveMailCredentials.checked,
        recipient_email: recipientEmailInput.value.trim(),
        vendor_name: vendorNameInput.value.trim(),
        cs_origin: csOriginInput.value.trim(),
        cs_product: csProductInput.value.trim(),
        cs_receiver: csReceiverInput.value.trim(),
        cs_phone: csPhoneInput.value.trim(),
        cs_address: csAddressInput.value.trim(),
        cs_type: csTypeInput.value.trim(),
        cs_content: csContentInput.value.trim(),
        subject: csSubjectInput.value.trim(),
        body: csBodyInput.value.trim(),
      };
    }

    function renderCsCases(cases) {
      csCaseList.innerHTML = "";
      if (!cases || cases.length === 0) {
        const empty = document.createElement("div");
        empty.className = "cs-case-item";
        empty.textContent = "저장된 CS건이 없습니다.";
        csCaseList.appendChild(empty);
        return;
      }
      cases.forEach((csCase) => {
        const item = document.createElement("div");
        item.className = "cs-case-item";
        const title = [
          `#${csCase.id}`,
          csCase.status || "접수",
          csCase.vendor_name || "업체 미입력",
          csCase.product_name || "상품 미입력",
        ].join(" · ");
        const meta = [
          csCase.receiver_name || "수령인 미입력",
          csCase.receiver_phone || "",
          csCase.original_invoice ? `원송장 ${csCase.original_invoice}` : "",
          csCase.mail_sent_at ? `메일 ${csCase.mail_sent_at}` : "",
        ].filter(Boolean).join(" / ");
        item.innerHTML = `<strong>${title}</strong><div class="cs-case-meta">${meta}</div>`;
        csCaseList.appendChild(item);
      });
    }

    function escapeHtml(value) {
      return String(value || "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
    }

    function todayStatusDate() {
      const now = new Date();
      return `${now.getMonth() + 1}/${now.getDate()}`;
    }

    function parseDateParts(value) {
      const text = String(value || "").trim();
      if (!text) return null;
      let match = text.match(/(\d{4})[-./년\s]+(\d{1,2})[-./월\s]+(\d{1,2})/);
      if (match) {
        return {
          year: match[1],
          month: String(Number(match[2])).padStart(2, "0"),
          day: String(Number(match[3])).padStart(2, "0"),
        };
      }
      match = text.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일?/);
      if (match) {
        return {
          year: "",
          month: String(Number(match[1])).padStart(2, "0"),
          day: String(Number(match[2])).padStart(2, "0"),
        };
      }
      match = text.match(/^(\d{1,2})[./-](\d{1,2})$/);
      if (match) {
        return {
          year: "",
          month: String(Number(match[1])).padStart(2, "0"),
          day: String(Number(match[2])).padStart(2, "0"),
        };
      }
      return null;
    }

    function shortKoreanDate(value) {
      const parts = parseDateParts(value);
      return parts ? `${Number(parts.month)}월 ${Number(parts.day)}일` : String(value || "");
    }

    function fullDateForSave(displayValue, rawValue) {
      const displayParts = parseDateParts(displayValue);
      if (!displayParts) return String(displayValue || "").trim();
      const rawParts = parseDateParts(rawValue);
      const year = displayParts.year || rawParts?.year || String(new Date().getFullYear());
      return `${year}-${displayParts.month}-${displayParts.day}`;
    }

    function ledgerStatusOptions(currentStatus = "") {
      const dateText = todayStatusDate();
      const statuses = [
        `회수지시(${dateText})`,
        `회수 완료(${dateText})`,
        `재발송 완료(${dateText})`,
        `재발송 완료(${dateText})/회수지시(${dateText})`,
        "전체 처리완료",
      ];
      const normalizedCurrent = currentStatus || statuses[0];
      return statuses.includes(normalizedCurrent) ? statuses : [normalizedCurrent, ...statuses];
    }

    const csTypeOptions = ["변심반품", "불량반품", "불량교환", "불량재출고(미회수)", "오출고(오배송)"];

    function selectOptions(options, currentValue = "") {
      const normalizedCurrent = currentValue || "";
      const optionList = normalizedCurrent && !options.includes(normalizedCurrent)
        ? [normalizedCurrent, ...options]
        : options;
      return optionList.map((option) => (
        `<option value="${escapeHtml(option)}" ${option === normalizedCurrent ? "selected" : ""}>${escapeHtml(option)}</option>`
      )).join("");
    }

    function isCompletedByValues(typeValue, statusValue) {
      const type = String(typeValue || "").replaceAll(" ", "").trim();
      const status = String(statusValue || "").replaceAll(" ", "").trim();
      if (status.includes("전체처리완료")) return true;
      if ((type === "변심반품" || type === "변신반품" || type === "불량반품") && status.includes("회수완료")) return true;
      if ((type === "불량교환" || type === "오출고(오배송)") && status.includes("전체처리완료")) return true;
      if (type === "불량재출고(미회수)" && status.includes("재발송완료")) return true;
      return false;
    }

    function isCompletedCsCase(csCase) {
      return isCompletedByValues(csCase.cs_type, csCase.status);
    }

    function updateLedgerRowCompletion(row) {
      if (!row) return;
      const status = row.querySelector('[data-field="status"]')?.value || "";
      const csType = row.querySelector('[data-field="cs_type"]')?.value || "";
      row.classList.toggle("completed-cs", isCompletedByValues(csType, status));
    }

    function ledgerFieldValue(csCase, field) {
      if (field === "purchase_vendor") return csCase.purchase_vendor || csCase.vendor_name || "";
      if (field === "occurred_at") return csCase.occurred_at || csCase.created_at || "";
      if (field === "original_invoice") return csCase.original_invoice || csCase.original_info || "";
      return csCase[field] || "";
    }

    function matchesLedgerFilters(csCase) {
      const toolbarStatus = ledgerStatusFilter.value.trim().toLowerCase();
      if (toolbarStatus && !String(csCase.status || "").toLowerCase().includes(toolbarStatus)) return false;
      return Object.entries(ledgerFilters).every(([field, filterValue]) => {
        const value = String(filterValue || "").trim().toLowerCase();
        if (!value) return true;
        return String(ledgerFieldValue(csCase, field)).toLowerCase().includes(value);
      });
    }

    function applyLedgerFilters() {
      const filtered = ledgerCases.filter(matchesLedgerFilters);
      renderLedger(filtered);
      ledgerFilterButtons.forEach((button) => {
        const field = button.dataset.ledgerFilterButton;
        button.classList.toggle("active", Boolean(ledgerFilters[field]));
      });
      ledgerCountLabel.textContent = `불러온 ${ledgerCases.length}건 / 표시 ${filtered.length}건`;
      if (currentMode === "ledger") notice.textContent = `${filtered.length}건 조회되었습니다.`;
    }

    function managementFieldValue(record, field) {
      return record[field] || "";
    }

    function matchesManagementFilters(record) {
      return Object.entries(managementFilters).every(([field, filterValue]) => {
        const value = String(filterValue || "").trim().toLowerCase();
        if (!value) return true;
        return String(managementFieldValue(record, field)).toLowerCase().includes(value);
      });
    }

    function applyManagementFilters() {
      const filtered = managementRecords.filter(matchesManagementFilters);
      renderManagement(filtered);
      managementFilterButtons.forEach((button) => {
        const field = button.dataset.managementFilterButton;
        button.classList.toggle("active", Boolean(managementFilters[field]));
      });
      if (currentMode === "management") notice.textContent = `${filtered.length}건 조회되었습니다.`;
    }

    function renderLedgerFilterOptions(field, searchText = "") {
      const normalizedSearch = searchText.trim().toLowerCase();
      const values = Array.from(new Set(
        ledgerCases
          .map((csCase) => String(ledgerFieldValue(csCase, field) || "").trim())
          .filter(Boolean)
      )).sort((left, right) => left.localeCompare(right, "ko"));
      const filteredValues = values
        .filter((value) => !normalizedSearch || value.toLowerCase().includes(normalizedSearch))
        .slice(0, 220);
      ledgerFilterOptions.innerHTML = filteredValues.length
        ? filteredValues.map((value) => (
          `<button class="ledger-filter-option" type="button" data-filter-value="${escapeHtml(value)}">${escapeHtml(value)}</button>`
        )).join("")
        : `<button class="ledger-filter-option" type="button" disabled>표시할 값이 없습니다.</button>`;
    }

    function renderManagementFilterOptions(field, searchText = "") {
      const normalizedSearch = searchText.trim().toLowerCase();
      const values = Array.from(new Set(
        managementRecords
          .map((record) => String(managementFieldValue(record, field) || "").trim())
          .filter(Boolean)
      )).sort((left, right) => left.localeCompare(right, "ko"));
      const filteredValues = values
        .filter((value) => !normalizedSearch || value.toLowerCase().includes(normalizedSearch))
        .slice(0, 220);
      ledgerFilterOptions.innerHTML = filteredValues.length
        ? filteredValues.map((value) => (
          `<button class="ledger-filter-option" type="button" data-filter-value="${escapeHtml(value)}">${escapeHtml(value)}</button>`
        )).join("")
        : `<button class="ledger-filter-option" type="button" disabled>표시할 값이 없습니다.</button>`;
    }

    function openLedgerFilter(button) {
      activeManagementFilterField = "";
      activeLedgerFilterField = button.dataset.ledgerFilterButton || "";
      ledgerFilterTitle.textContent = `${button.dataset.label || "필터"} 필터`;
      ledgerFilterSearch.value = ledgerFilters[activeLedgerFilterField] || "";
      renderLedgerFilterOptions(activeLedgerFilterField, ledgerFilterSearch.value);
      const rect = button.getBoundingClientRect();
      ledgerFilterPopover.style.left = `${Math.min(rect.left, window.innerWidth - 280)}px`;
      ledgerFilterPopover.style.top = `${Math.min(rect.bottom + 6, window.innerHeight - 410)}px`;
      ledgerFilterPopover.classList.add("open");
      ledgerFilterSearch.focus();
      ledgerFilterSearch.select();
    }

    function openManagementFilter(button) {
      activeLedgerFilterField = "";
      activeManagementFilterField = button.dataset.managementFilterButton || "";
      ledgerFilterTitle.textContent = `${button.dataset.label || "필터"} 필터`;
      ledgerFilterSearch.value = managementFilters[activeManagementFilterField] || "";
      renderManagementFilterOptions(activeManagementFilterField, ledgerFilterSearch.value);
      const rect = button.getBoundingClientRect();
      ledgerFilterPopover.style.left = `${Math.min(rect.left, window.innerWidth - 280)}px`;
      ledgerFilterPopover.style.top = `${Math.min(rect.bottom + 6, window.innerHeight - 410)}px`;
      ledgerFilterPopover.classList.add("open");
      ledgerFilterSearch.focus();
      ledgerFilterSearch.select();
    }

    function closeLedgerFilter() {
      ledgerFilterPopover.classList.remove("open");
      activeLedgerFilterField = "";
      activeManagementFilterField = "";
    }

    function markRowDirty(row, dirty = true) {
      if (!row || !row.dataset) return;
      if (dirty) {
        row.dataset.dirty = "1";
        row.classList.add("row-dirty");
      } else {
        delete row.dataset.dirty;
        row.classList.remove("row-dirty");
      }
    }

    function applyRowPermissions(row) {
      if (!row) return;
      if (!can("ledger_edit")) {
        row.querySelectorAll("[data-field], [data-management-field]").forEach((input) => {
          input.disabled = true;
        });
      }
      row.querySelectorAll(".management-cs-button").forEach((button) => {
        setHidden(button, !can("cs_receive"));
      });
    }

    function dirtyRows(container, rowSelector) {
      return Array.from(container.querySelectorAll(`${rowSelector}[data-dirty="1"]`));
    }

    function selectedRows(container, rowSelector) {
      return Array.from(container.querySelectorAll(rowSelector))
        .filter((row) => row.querySelector("[data-row-check]")?.checked);
    }

    function setLedgerFilter(value) {
      if (!activeLedgerFilterField) return;
      const normalized = String(value || "").trim();
      if (normalized) ledgerFilters[activeLedgerFilterField] = normalized;
      else delete ledgerFilters[activeLedgerFilterField];
      applyLedgerFilters();
      closeLedgerFilter();
    }

    function setManagementFilter(value) {
      if (!activeManagementFilterField) return;
      const normalized = String(value || "").trim();
      if (normalized) managementFilters[activeManagementFilterField] = normalized;
      else delete managementFilters[activeManagementFilterField];
      applyManagementFilters();
      closeLedgerFilter();
    }

    function setActivePopoverFilter(value) {
      if (activeManagementFilterField) setManagementFilter(value);
      else setLedgerFilter(value);
    }

    function refreshActiveFilterOptions() {
      if (activeManagementFilterField) {
        renderManagementFilterOptions(activeManagementFilterField, ledgerFilterSearch.value);
      } else {
        renderLedgerFilterOptions(activeLedgerFilterField, ledgerFilterSearch.value);
      }
    }

    function resetCsFormInputs() {
      naverEmailInput.value = "";
      naverPasswordInput.value = "";
      naverPasswordInput.placeholder = "저장된 비밀번호가 없으면 입력";
      saveMailCredentials.checked = true;
      vendorContactSelect.value = "";
      vendorContactsFileInput.value = "";
      vendorContactsDropMain.textContent = "업체명/메일주소 엑셀을 선택해주세요.";
      recipientEmailInput.value = "";
      vendorNameInput.value = "";
      csOriginInput.value = "";
      csProductInput.value = "";
      csReceiverInput.value = "";
      csPhoneInput.value = "";
      csAddressInput.value = "";
      csTypeInput.value = "";
      csContentInput.value = "";
      csSubjectInput.value = defaultCsSubject();
      refreshCsBody();
    }

    function openLedgerCsPopup() {
      closeLedgerFilter();
      resetCsFormInputs();
      csFields.classList.add("ledger-cs-popup");
      csFields.style.display = "block";
      loadMailSettings();
      loadVendorContacts();
      loadCsCases();
      notice.textContent = "새 CS 내용을 입력한 뒤 CS건 DB 저장을 눌러주세요.";
      setTimeout(() => vendorNameInput.focus(), 0);
    }

    function closeLedgerCsPopup() {
      csFields.classList.remove("ledger-cs-popup");
      if (currentMode === "ledger") csFields.style.display = "none";
    }

    function renderLedger(cases) {
      ledgerBody.innerHTML = "";
      if (!cases || cases.length === 0) {
        const row = document.createElement("tr");
        row.innerHTML = `<td colspan="21">조회된 CS건이 없습니다.</td>`;
        ledgerBody.appendChild(row);
        return;
      }
      cases.forEach((csCase) => {
        const row = document.createElement("tr");
        row.dataset.caseId = csCase.id;
        const statusOptions = selectOptions(ledgerStatusOptions(csCase.status), csCase.status || ledgerStatusOptions()[0]);
        const csTypeSelectOptions = `<option value="" ${csCase.cs_type ? "" : "selected"}>선택</option>${selectOptions(csTypeOptions, csCase.cs_type)}`;
        if (isCompletedCsCase(csCase)) row.classList.add("completed-cs");
        row.innerHTML = `
          <td><input class="ledger-check" type="checkbox" data-row-check /></td>
          <td>${escapeHtml(csCase.occurred_at || csCase.created_at)}</td>
          <td>${escapeHtml(csCase.sales_vendor)}</td>
          <td>${escapeHtml(csCase.purchase_vendor || csCase.vendor_name)}</td>
          <td><select class="ledger-status-select" data-field="status">${statusOptions}</select></td>
          <td>${escapeHtml(csCase.completed_at)}</td>
          <td><select class="ledger-status-select" data-field="cs_type">${csTypeSelectOptions}</select></td>
          <td class="left">${escapeHtml(csCase.cs_content)}</td>
          <td><input class="ledger-edit" data-field="reship_invoice" value="${escapeHtml(csCase.reship_invoice)}" /></td>
          <td><input class="ledger-edit" data-field="return_invoice" value="${escapeHtml(csCase.return_invoice)}" /></td>
          <td data-full-date="${escapeHtml(csCase.order_date)}">${escapeHtml(shortKoreanDate(csCase.order_date))}</td>
          <td data-full-date="${escapeHtml(csCase.ship_date)}">${escapeHtml(shortKoreanDate(csCase.ship_date))}</td>
          <td>${escapeHtml(csCase.orderer_name)}</td>
          <td>${escapeHtml(csCase.orderer_phone)}</td>
          <td>${escapeHtml(csCase.receiver_name)}</td>
          <td>${escapeHtml(csCase.receiver_phone)}</td>
          <td class="left">${escapeHtml(csCase.product_name)}</td>
          <td>${escapeHtml(csCase.quantity)}</td>
          <td class="left">${escapeHtml(csCase.receiver_address)}</td>
          <td>${escapeHtml(csCase.courier)}</td>
          <td>${escapeHtml(csCase.original_invoice || csCase.original_info)}</td>
        `;
        applyRowPermissions(row);
        ledgerBody.appendChild(row);
      });
    }

    async function loadCsCases() {
      try {
        const response = await fetch("/api/cs-cases");
        if (!response.ok) return;
        const data = await response.json();
        renderCsCases(data.cases || []);
      } catch {
        renderCsCases([]);
      }
    }

    async function loadLedgerCases() {
      const query = ledgerSearchInput.value.trim();
      const params = new URLSearchParams({ limit: ledgerPageSize.value || "100" });
      if (query) params.set("q", query);
      if (ledgerYearFilter.value) params.set("year", ledgerYearFilter.value);
      if (ledgerMonthFilter.value) params.set("month", ledgerMonthFilter.value);
      const url = `/api/cs-cases?${params.toString()}`;
      try {
        const response = await fetch(url);
        if (!response.ok) return;
        const data = await response.json();
        ledgerCases = data.cases || [];
        applyLedgerFilters();
      } catch {
        ledgerCases = [];
        renderLedger([]);
        notice.textContent = "CS 처리대장을 불러오지 못했습니다.";
      }
    }

    async function uploadLedgerWorkbook() {
      const file = ledgerImportInput.files[0];
      if (!file) return;
      ledgerImportDropMain.textContent = file.name;
      const formData = new FormData();
      formData.append("file", file);
      notice.textContent = "CS 처리대장 데이터를 DB에 업로드 중입니다.";
      try {
        const response = await fetch("/api/cs-cases-import", {
          method: "POST",
          body: formData,
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "CS 처리대장 업로드에 실패했습니다.");
        notice.textContent = data.message || "CS 처리대장 데이터를 업로드했습니다.";
        await loadLedgerCases();
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        ledgerImportInput.value = "";
        ledgerImportDropMain.textContent = "업로드";
      }
    }

    function renderManagement(records) {
      managementBody.innerHTML = "";
      if (!records || records.length === 0) {
        const row = document.createElement("tr");
        row.innerHTML = `<td colspan="18">조회된 통합관리대장 데이터가 없습니다.</td>`;
        managementBody.appendChild(row);
        return;
      }
      const duplicateColors = [
        "#fff7d6",
        "#e8f7ee",
        "#e8f1ff",
        "#f4eaff",
        "#ffecef",
        "#e8faf8",
        "#fff0df",
        "#eef2ff",
      ];
      const duplicateCounts = new Map();
      const duplicateColorByKey = new Map();
      records.forEach((record) => {
        const dateKey = String(record.order_date || record.ship_date || "").trim();
        const invoiceKey = String(record.invoice_number || "").trim();
        if (!dateKey || !invoiceKey) return;
        const key = `${dateKey}||${invoiceKey}`;
        duplicateCounts.set(key, (duplicateCounts.get(key) || 0) + 1);
      });
      let duplicateGroupIndex = 0;
      records.forEach((record) => {
        const row = document.createElement("tr");
        row.dataset.recordId = record.id;
        const dateKey = String(record.order_date || record.ship_date || "").trim();
        const invoiceKey = String(record.invoice_number || "").trim();
        const duplicateKey = dateKey && invoiceKey ? `${dateKey}||${invoiceKey}` : "";
        if (duplicateKey && duplicateCounts.get(duplicateKey) > 1) {
          if (!duplicateColorByKey.has(duplicateKey)) {
            duplicateColorByKey.set(
              duplicateKey,
              duplicateColors[duplicateGroupIndex % duplicateColors.length]
            );
            duplicateGroupIndex += 1;
          }
          row.classList.add("management-duplicate");
          row.style.setProperty("--duplicate-row-color", duplicateColorByKey.get(duplicateKey));
        }
        const csReceived = Boolean(record.cs_received_at);
        if (csReceived) row.classList.add("management-cs-received");
        row.innerHTML = `
          <td><input class="ledger-check" type="checkbox" data-row-check /></td>
          <td><input class="management-edit" data-management-field="order_date" data-raw-date="${escapeHtml(record.order_date)}" value="${escapeHtml(shortKoreanDate(record.order_date))}" /></td>
          <td><input class="management-edit" data-management-field="ship_date" data-raw-date="${escapeHtml(record.ship_date)}" value="${escapeHtml(shortKoreanDate(record.ship_date))}" /></td>
          <td><input class="management-edit" data-management-field="purchase_vendor" value="${escapeHtml(record.purchase_vendor)}" /></td>
          <td><input class="management-edit" data-management-field="sales_vendor" value="${escapeHtml(record.sales_vendor)}" /></td>
          <td><input class="management-edit" data-management-field="transaction_type" value="${escapeHtml(record.transaction_type)}" /></td>
          <td><input class="management-edit" data-management-field="ledger_checked" value="${escapeHtml(record.ledger_checked)}" /></td>
          <td><input class="management-edit" data-management-field="orderer_name" value="${escapeHtml(record.orderer_name)}" /></td>
          <td><input class="management-edit" data-management-field="sender_phone" value="${escapeHtml(record.sender_phone)}" /></td>
          <td><input class="management-edit" data-management-field="receiver_name" value="${escapeHtml(record.receiver_name)}" /></td>
          <td><input class="management-edit" data-management-field="receiver_phone" value="${escapeHtml(record.receiver_phone)}" /></td>
          <td class="left"><input class="management-edit wide" data-management-field="product_name" value="${escapeHtml(record.product_name)}" /></td>
          <td><input class="management-edit" data-management-field="quantity" value="${escapeHtml(record.quantity)}" /></td>
          <td class="left"><input class="management-edit wide" data-management-field="receiver_address" value="${escapeHtml(record.receiver_address)}" /></td>
          <td><input class="management-edit" data-management-field="courier" value="${escapeHtml(record.courier)}" /></td>
          <td><input class="management-edit" data-management-field="invoice_number" value="${escapeHtml(record.invoice_number)}" /></td>
          <td class="left"><input class="management-edit wide" data-management-field="memo" value="${escapeHtml(record.memo)}" /></td>
          <td><button class="management-cs-button" type="button" ${csReceived ? "disabled" : ""}>${csReceived ? "접수완료" : "CS접수"}</button></td>
        `;
        applyRowPermissions(row);
        managementBody.appendChild(row);
      });
    }

    async function loadManagementRecords() {
      const query = managementSearchInput.value.trim();
      const params = new URLSearchParams({ limit: managementPageSize.value || "500" });
      const period = selectedManagementPeriod();
      if (query) params.set("q", query);
      if (period.year) params.set("year", period.year);
      if (period.month) params.set("month", period.month);
      renderManagementMonthTabs();
      try {
        const response = await fetch(`/api/management-records?${params.toString()}`);
        if (!response.ok) return;
        const data = await response.json();
        managementRecords = data.records || [];
        applyManagementFilters();
      } catch {
        managementRecords = [];
        applyManagementFilters();
        notice.textContent = "통합관리대장을 불러오지 못했습니다.";
      }
    }

    async function loadManagementPeriods() {
      if (!managementYearFilter || !managementMonthFilter) return;
      try {
        const response = await fetch("/api/management-periods");
        if (!response.ok) return;
        const data = await response.json();
        managementPeriods = data.periods || [];
        renderManagementPeriodControls();
      } catch {
        managementPeriods = [];
        renderManagementPeriodControls();
      }
    }

    async function loadManagementWorkspaceData() {
      await loadManagementPeriods();
      await loadManagementRecords();
    }

    async function uploadManagementWorkbook() {
      const file = managementImportInput.files[0];
      if (!file) return;
      managementImportDropMain.textContent = file.name;
      const formData = new FormData();
      formData.append("file", file);
      notice.textContent = "통합관리대장 데이터를 DB에 업로드 중입니다.";
      try {
        const response = await fetch("/api/management-import", {
          method: "POST",
          body: formData,
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "통합관리대장 업로드에 실패했습니다.");
        notice.textContent = data.message || "통합관리대장 데이터를 업로드했습니다.";
        await loadManagementWorkspaceData();
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        managementImportInput.value = "";
        managementImportDropMain.textContent = "업로드";
      }
    }

    function collectManagementRow(row) {
      const payload = { id: row.dataset.recordId };
      row.querySelectorAll("[data-management-field]").forEach((input) => {
        const field = input.dataset.managementField;
        payload[field] = field === "order_date" || field === "ship_date"
          ? fullDateForSave(input.value.trim(), input.dataset.rawDate || "")
          : input.value.trim();
      });
      return payload;
    }

    async function saveManagementPayload(payload) {
      const response = await fetch("/api/management-record-update", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "통합관리대장 저장에 실패했습니다.");
      return data;
    }

    function updateManagementRecordCache(payload) {
      const record = managementRecords.find((item) => String(item.id) === String(payload.id));
      if (!record) return;
      Object.entries(payload).forEach(([key, value]) => {
        if (key !== "id") record[key] = value;
      });
    }

    async function saveManagementRow(button) {
      const row = button.closest("tr");
      if (!row) return;
      const payload = collectManagementRow(row);
      try {
        button.disabled = true;
        const data = await saveManagementPayload(payload);
        notice.textContent = data.message || "통합관리대장 행을 저장했습니다.";
        updateManagementRecordCache(payload);
        markRowDirty(row, false);
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        button.disabled = false;
      }
    }

    async function receiveManagementCs(button) {
      const row = button.closest("tr");
      if (!row) return;
      const payload = collectManagementRow(row);
      try {
        button.disabled = true;
        const saveResponse = await fetch("/api/management-record-update", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
        const saveData = await saveResponse.json();
        if (!saveResponse.ok) throw new Error(saveData.error || "통합관리대장 저장에 실패했습니다.");
        const response = await fetch("/api/management-to-cs", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: payload.id }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "CS 처리대장 접수에 실패했습니다.");
        notice.textContent = data.message || "CS 처리대장에 접수했습니다.";
        showWorkspace("ledger");
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        button.disabled = false;
      }
    }

    function collectLedgerRow(row) {
      return {
        id: row.dataset.caseId,
        status: row.querySelector('[data-field="status"]')?.value || "",
        cs_type: row.querySelector('[data-field="cs_type"]')?.value.trim() || "",
        return_invoice: row.querySelector('[data-field="return_invoice"]')?.value.trim() || "",
        reship_invoice: row.querySelector('[data-field="reship_invoice"]')?.value.trim() || "",
      };
    }

    async function saveLedgerPayload(payload) {
      const response = await fetch("/api/cs-case-update", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "CS 처리내용 저장에 실패했습니다.");
      return data;
    }

    function updateLedgerCaseCache(payload) {
      const savedCase = ledgerCases.find((item) => String(item.id) === String(payload.id));
      if (!savedCase) return;
      savedCase.status = payload.status;
      savedCase.cs_type = payload.cs_type;
      savedCase.return_invoice = payload.return_invoice;
      savedCase.reship_invoice = payload.reship_invoice;
    }

    async function saveLedgerRow(button) {
      const row = button.closest("tr");
      if (!row) return;
      const payload = collectLedgerRow(row);
      try {
        button.disabled = true;
        const data = await saveLedgerPayload(payload);
        notice.textContent = data.message || "CS 처리내용을 저장했습니다.";
        updateLedgerCaseCache(payload);
        updateLedgerRowCompletion(row);
        markRowDirty(row, false);
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        button.disabled = false;
      }
    }

    async function saveVisibleManagementRows({ silent = false, selectedOnly = false } = {}) {
      const rows = selectedOnly
        ? selectedRows(managementBody, "tr[data-record-id]")
        : dirtyRows(managementBody, "tr[data-record-id]");
      if (rows.length === 0) {
        if (!silent) notice.textContent = selectedOnly ? "체크된 통합관리대장 행이 없습니다." : "저장할 변경 내용이 없습니다.";
        return 0;
      }
      for (const row of rows) {
        const payload = collectManagementRow(row);
        await saveManagementPayload(payload);
        updateManagementRecordCache(payload);
        markRowDirty(row, false);
        const checkbox = row.querySelector("[data-row-check]");
        if (checkbox) checkbox.checked = false;
      }
      if (!silent) notice.textContent = `통합관리대장 ${rows.length}건 저장 완료`;
      return rows.length;
    }

    async function saveVisibleLedgerRows({ silent = false, selectedOnly = false } = {}) {
      const rows = selectedOnly
        ? selectedRows(ledgerBody, "tr[data-case-id]")
        : dirtyRows(ledgerBody, "tr[data-case-id]");
      if (rows.length === 0) {
        if (!silent) notice.textContent = selectedOnly ? "체크된 CS 처리대장 행이 없습니다." : "저장할 변경 내용이 없습니다.";
        return 0;
      }
      for (const row of rows) {
        const payload = collectLedgerRow(row);
        await saveLedgerPayload(payload);
        updateLedgerCaseCache(payload);
        updateLedgerRowCompletion(row);
        markRowDirty(row, false);
        const checkbox = row.querySelector("[data-row-check]");
        if (checkbox) checkbox.checked = false;
      }
      if (!silent) notice.textContent = `CS 처리대장 ${rows.length}건 저장 완료`;
      return rows.length;
    }

    async function saveCurrentWorkspaceRows({ silent = false, mode = currentMode, selectedOnly = false } = {}) {
      if (isBulkSaving) return 0;
      if (mode !== "management" && mode !== "ledger") return 0;
      try {
        isBulkSaving = true;
        if (mode === "management") return await saveVisibleManagementRows({ silent, selectedOnly });
        return await saveVisibleLedgerRows({ silent, selectedOnly });
      } catch (error) {
        if (!silent) notice.textContent = error.message;
        return 0;
      } finally {
        isBulkSaving = false;
      }
    }

    function selectedIds(container, rowSelector, idName) {
      return selectedRows(container, rowSelector)
        .map((row) => row.dataset[idName])
        .filter(Boolean);
    }

    async function deleteSelectedRows(mode) {
      const isManagement = mode === "management";
      const ids = isManagement
        ? selectedIds(managementBody, "tr[data-record-id]", "recordId")
        : selectedIds(ledgerBody, "tr[data-case-id]", "caseId");
      if (ids.length === 0) {
        notice.textContent = isManagement ? "삭제할 통합관리대장 행을 체크해주세요." : "삭제할 CS 처리대장 행을 체크해주세요.";
        return;
      }
      const label = isManagement ? "통합관리대장" : "CS 처리대장";
      if (!window.confirm(`${label} 선택 주문 ${ids.length}건을 삭제할까요?`)) return;
      const endpoint = isManagement ? "/api/management-records-delete" : "/api/cs-cases-delete";
      try {
        const response = await fetch(endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ ids }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "선택 주문 삭제에 실패했습니다.");
        notice.textContent = data.message || `${label} 선택 주문 ${ids.length}건을 삭제했습니다.`;
        if (isManagement) await loadManagementWorkspaceData();
        else await loadLedgerCases();
      } catch (error) {
        notice.textContent = error.message;
      }
    }

    function collectManagementExportRows() {
      return Array.from(managementBody.querySelectorAll("tr[data-record-id]")).map((row) => collectManagementRow(row));
    }

    function collectSelectedManagementExportRows() {
      return selectedRows(managementBody, "tr[data-record-id]").map((row) => collectManagementRow(row));
    }

    function textFromCell(row, index) {
      return row.children[index]?.textContent.trim() || "";
    }

    function fullDateFromCell(row, index) {
      return row.children[index]?.dataset.fullDate || textFromCell(row, index);
    }

    function collectLedgerExportRows() {
      return Array.from(ledgerBody.querySelectorAll("tr[data-case-id]")).map((row) => ({
        occurred_at: textFromCell(row, 1),
        sales_vendor: textFromCell(row, 2),
        purchase_vendor: textFromCell(row, 3),
        status: row.querySelector('[data-field="status"]')?.value || "",
        completed_at: textFromCell(row, 5),
        cs_type: row.querySelector('[data-field="cs_type"]')?.value || "",
        cs_content: textFromCell(row, 7),
        reship_invoice: row.querySelector('[data-field="reship_invoice"]')?.value.trim() || "",
        return_invoice: row.querySelector('[data-field="return_invoice"]')?.value.trim() || "",
        order_date: fullDateFromCell(row, 10),
        ship_date: fullDateFromCell(row, 11),
        orderer_name: textFromCell(row, 12),
        orderer_phone: textFromCell(row, 13),
        receiver_name: textFromCell(row, 14),
        receiver_phone: textFromCell(row, 15),
        product_name: textFromCell(row, 16),
        quantity: textFromCell(row, 17),
        receiver_address: textFromCell(row, 18),
        courier: textFromCell(row, 19),
        original_invoice: textFromCell(row, 20),
      }));
    }

    async function downloadExcel(endpoint, rows, fallbackName, button) {
      if (!rows.length) {
        notice.textContent = "다운로드할 데이터가 없습니다.";
        return;
      }
      try {
        button.disabled = true;
        const response = await fetch(endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ rows }),
        });
        if (!response.ok) {
          const data = await response.json();
          throw new Error(data.error || "엑셀 다운로드에 실패했습니다.");
        }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = filenameFromResponse(response, fallbackName);
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        notice.textContent = "엑셀 다운로드가 시작되었습니다.";
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        button.disabled = false;
      }
    }

    async function downloadManagementExcel(scope, button) {
      const payload = { scope };
      let fallbackName = "통합관리대장.xlsx";
      if (scope === "selected") {
        payload.rows = collectSelectedManagementExportRows();
        if (!payload.rows.length) {
          notice.textContent = "다운로드할 행을 체크해주세요.";
          return;
        }
        fallbackName = "통합관리대장_선택.xlsx";
      } else if (scope === "month") {
        const period = selectedManagementPeriod();
        if (!period.year || !period.month) {
          notice.textContent = "월별 다운로드는 년도와 월을 선택해주세요.";
          return;
        }
        payload.year = period.year;
        payload.month = period.month;
        fallbackName = `통합관리대장_${payload.year}년_${Number(payload.month)}월.xlsx`;
      } else if (scope === "year") {
        const period = selectedManagementPeriod();
        if (!period.year) {
          notice.textContent = "년별 다운로드는 년도를 선택해주세요.";
          return;
        }
        payload.year = period.year;
        fallbackName = `통합관리대장_${payload.year}년.xlsx`;
      } else {
        fallbackName = "통합관리대장_전체.xlsx";
      }

      try {
        button.disabled = true;
        const response = await fetch("/api/management-export", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
        if (!response.ok) {
          const data = await response.json();
          throw new Error(data.error || "통합관리대장 엑셀 다운로드에 실패했습니다.");
        }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = filenameFromResponse(response, fallbackName);
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        notice.textContent = "통합관리대장 엑셀 다운로드가 시작되었습니다.";
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        button.disabled = false;
      }
    }

    async function saveCurrentCsCase() {
      refreshCsBody();
      const payload = collectCsPayload();
      if (!payload.vendor_name && !payload.cs_receiver && !payload.cs_product && !payload.cs_content) {
        notice.textContent = "업체명, 수령인, 상품명, CS내용 중 하나 이상 입력해주세요.";
        return;
      }
      try {
        saveCsCaseButton.disabled = true;
        const response = await fetch("/api/cs-case", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "CS건 저장에 실패했습니다.");
        notice.textContent = data.message || "CS건을 DB에 저장했습니다.";
        if (currentMode === "ledger") {
          closeLedgerCsPopup();
          await loadLedgerCases();
        } else {
          await loadCsCases();
        }
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        saveCsCaseButton.disabled = false;
      }
    }

    function filenameFromResponse(response, fallback) {
      const disposition = response.headers.get("Content-Disposition") || "";
      const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
      if (utf8Match) {
        try {
          return decodeURIComponent(utf8Match[1]);
        } catch {
          return utf8Match[1];
        }
      }
      const asciiMatch = disposition.match(/filename="?([^";]+)"?/i);
      return asciiMatch ? asciiMatch[1] : fallback;
    }

    function openModal(mode) {
      currentMode = mode;
      closeLedgerCsPopup();
      modal.classList.add("open");
      const modalPanel = modal.querySelector(".modal");
      modalPanel.classList.toggle("ledger-modal", mode === "ledger" || mode === "management");
      modalPanel.classList.toggle("ledger-view", mode === "ledger");
      modalPanel.classList.toggle("management-view", mode === "management");
      result.classList.remove("open");
      resultText.value = "";
      fileInput.value = "";
      templateInput.value = "";
      receiptTypeSelect.value = "일반";
      supplierInput.value = "";
      receiptDateInput.value = todayString();
      freightPaymentSelect.value = "선불";
      requestNoteInput.value = "";
      deliveryPlaceInput.value = "";
      managerInput.value = "";
      naverEmailInput.value = "";
      naverPasswordInput.value = "";
      naverPasswordInput.placeholder = "저장된 비밀번호가 없으면 입력";
      saveMailCredentials.checked = true;
      vendorContactSelect.value = "";
      vendorContactsFileInput.value = "";
      vendorContactsDropMain.textContent = "업체명/메일주소 엑셀을 선택해주세요.";
      recipientEmailInput.value = "";
      vendorNameInput.value = "";
      csOriginInput.value = "";
      csProductInput.value = "";
      csReceiverInput.value = "";
      csPhoneInput.value = "";
      csAddressInput.value = "";
      csTypeInput.value = "";
      csContentInput.value = "";
      csSubjectInput.value = defaultCsSubject();
      refreshCsBody();
      resetProductRows();
      notice.textContent = "";
      dropMain.textContent = "파일을 선택하거나 여기에 올려주세요.";
      templateDropMain.textContent = "롯데택배 발주서 양식을 선택해주세요.";
      if (mode === "delivery") {
        modalTitle.textContent = "택배건 요약";
        fileLabel.textContent = "주소일브릿지 엑셀 선택";
        submitButton.textContent = "생성";
        submitButton.className = "btn primary";
        deliveryOptions.style.display = "flex";
        templateUpload.style.display = "none";
        vehicleFields.style.display = "none";
        csFields.style.display = "none";
        ledgerFields.style.display = "none";
        managementFields.style.display = "none";
        fileInput.required = false;
        templateInput.required = false;
        dropSub.textContent = "주소일브릿지 엑셀을 업로드하면 전달용 텍스트를 만듭니다.";
      } else if (mode === "invoice") {
        modalTitle.textContent = "송장번호 추출";
        fileLabel.textContent = "출고송장 엑셀 선택";
        submitButton.textContent = "엑셀 생성";
        submitButton.className = "btn blue";
        deliveryOptions.style.display = "none";
        templateUpload.style.display = "none";
        vehicleFields.style.display = "none";
        csFields.style.display = "none";
        ledgerFields.style.display = "none";
        managementFields.style.display = "none";
        fileInput.required = false;
        templateInput.required = false;
        dropSub.textContent = "출고송장 엑셀을 업로드하면 수하인별 송장번호 엑셀을 다운로드합니다.";
      } else if (mode === "lotte") {
        modalTitle.textContent = "롯데택배 발주서 변환";
        fileLabel.textContent = "주소일브릿지 원본 선택";
        submitButton.textContent = "엑셀 생성";
        submitButton.className = "btn primary";
        deliveryOptions.style.display = "none";
        templateUpload.style.display = "none";
        vehicleFields.style.display = "none";
        csFields.style.display = "none";
        ledgerFields.style.display = "none";
        managementFields.style.display = "none";
        fileInput.required = false;
        templateInput.required = false;
        dropSub.textContent = "주소일브릿지 원본을 업로드하면 지정된 롯데택배 발주서 양식으로 출력합니다.";
      } else if (mode === "vehicle") {
        modalTitle.textContent = "차량인수증 생성";
        submitButton.textContent = "인수증 생성";
        submitButton.className = "btn primary";
        document.querySelector("label[for='fileInput']").style.display = "none";
        deliveryOptions.style.display = "none";
        templateUpload.style.display = "none";
        vehicleFields.style.display = "block";
        csFields.style.display = "none";
        ledgerFields.style.display = "none";
        managementFields.style.display = "none";
        fileInput.required = false;
        templateInput.required = false;
      } else if (mode === "cs") {
        modalTitle.textContent = "업체 CS 요청";
        submitButton.textContent = "메일 전송";
        submitButton.className = "btn primary";
        deliveryOptions.style.display = "none";
        templateUpload.style.display = "none";
        vehicleFields.style.display = "none";
        csFields.style.display = "block";
        ledgerFields.style.display = "none";
        managementFields.style.display = "none";
        fileInput.required = false;
        templateInput.required = false;
        loadMailSettings();
        loadVendorContacts();
        loadCsCases();
      } else if (mode === "ledger") {
        modalTitle.textContent = "CS 처리대장";
        submitButton.textContent = "닫기";
        submitButton.className = "btn";
        deliveryOptions.style.display = "none";
        templateUpload.style.display = "none";
        vehicleFields.style.display = "none";
        csFields.style.display = "none";
        ledgerFields.style.display = "block";
        managementFields.style.display = "none";
        fileInput.required = false;
        templateInput.required = false;
        ledgerSearchInput.value = "";
        ledgerStatusFilter.value = "";
        ledgerYearFilter.value = "";
        ledgerMonthFilter.value = "";
        ledgerImportInput.value = "";
        ledgerImportDropMain.textContent = "업로드";
        Object.keys(ledgerFilters).forEach((key) => delete ledgerFilters[key]);
        closeLedgerFilter();
        loadLedgerCases();
      } else {
        modalTitle.textContent = "통합관리대장 관리";
        submitButton.textContent = "닫기";
        submitButton.className = "btn";
        deliveryOptions.style.display = "none";
        templateUpload.style.display = "none";
        vehicleFields.style.display = "none";
        csFields.style.display = "none";
        ledgerFields.style.display = "none";
        managementFields.style.display = "block";
        fileInput.required = false;
        templateInput.required = false;
        managementSearchInput.value = "";
        managementYearFilter.value = "";
        managementMonthFilter.value = "";
        managementPageSize.value = "500";
        managementImportInput.value = "";
        managementImportDropMain.textContent = "업로드";
        Object.keys(managementFilters).forEach((key) => delete managementFilters[key]);
        closeLedgerFilter();
        loadManagementWorkspaceData();
      }

      const fileDrop = document.querySelector("label[for='fileInput']");
      fileDrop.style.display = mode === "vehicle" || mode === "cs" || mode === "ledger" || mode === "management" ? "none" : "grid";
      fileLabel.style.display = mode === "vehicle" || mode === "cs" || mode === "ledger" || mode === "management" ? "none" : "block";
    }

    function closeModal() {
      closeLedgerCsPopup();
      closeLedgerFilter();
      modal.classList.remove("open");
    }

    function setPageTitle(text) {
      if (!pageTitle) return;
      const firstNode = pageTitle.childNodes[0];
      if (firstNode && firstNode.nodeType === Node.TEXT_NODE) {
        firstNode.nodeValue = `${text} `;
      } else {
        pageTitle.prepend(document.createTextNode(`${text} `));
      }
    }

    function setActiveNav(mode) {
      document.querySelectorAll(".nav-item").forEach((item) => item.classList.remove("active"));
      const selector = mode === "dashboard" ? '[data-view="dashboard"]' : `[data-open="${mode}"]`;
      const activeItem = document.querySelector(selector);
      if (activeItem) activeItem.classList.add("active");
    }

    function showWorkspace(mode) {
      closeModal();
      if (mode === "userAdmin" && !userAdminWorkspace) mode = "dashboard";
      if (mode === "leave" && !leaveWorkspace) mode = "dashboard";
      if (mode === "backup" && !backupWorkspace) mode = "dashboard";
      if (mode === "systemUpdate" && !systemUpdateWorkspace) mode = "dashboard";
      currentMode = mode;
      const showManagement = mode === "management";
      const showLedger = mode === "ledger";
      const showLeave = mode === "leave";
      const showUserAdmin = mode === "userAdmin" && Boolean(userAdminWorkspace);
      const showBackup = mode === "backup" && Boolean(backupWorkspace);
      const showSystemUpdate = mode === "systemUpdate" && Boolean(systemUpdateWorkspace);
      dashboardContent.style.display = mode === "dashboard" ? "" : "none";
      managementWorkspace.classList.toggle("active", showManagement);
      ledgerWorkspace.classList.toggle("active", showLedger);
      if (leaveWorkspace) leaveWorkspace.classList.toggle("active", showLeave);
      if (userAdminWorkspace) userAdminWorkspace.classList.toggle("active", showUserAdmin);
      if (backupWorkspace) backupWorkspace.classList.toggle("active", showBackup);
      if (systemUpdateWorkspace) systemUpdateWorkspace.classList.toggle("active", showSystemUpdate);
      setActiveNav(mode);
      if (showManagement) {
        setPageTitle("통합관리대장 관리");
        managementSearchInput.value = "";
        managementYearFilter.value = "";
        managementMonthFilter.value = "";
        managementPageSize.value = "500";
        managementImportInput.value = "";
        managementImportDropMain.textContent = "업로드";
        closeLedgerFilter();
        loadManagementWorkspaceData();
      } else if (showLedger) {
        setPageTitle("CS 처리대장");
        ledgerSearchInput.value = "";
        ledgerStatusFilter.value = "";
        ledgerImportInput.value = "";
        ledgerImportDropMain.textContent = "업로드";
        Object.keys(ledgerFilters).forEach((key) => delete ledgerFilters[key]);
        closeLedgerFilter();
        loadLedgerCases();
      } else if (showLeave) {
        setPageTitle("연차 관리");
        closeLedgerFilter();
        loadLeaveData();
      } else if (showUserAdmin) {
        setPageTitle("사용자 관리");
        closeLedgerFilter();
        resetUserAdminForm();
        loadUserAccounts();
      } else if (showBackup) {
        setPageTitle("백업 관리");
        closeLedgerFilter();
        loadBackups();
      } else if (showSystemUpdate) {
        setPageTitle("업데이트 관리");
        closeLedgerFilter();
        loadSystemUpdateStatus();
      } else {
        currentMode = "dashboard";
        setPageTitle("금일 공지사항");
        closeLedgerFilter();
      }
    }

    function openWorkspaceWindow(mode) {
      const url = new URL(window.location.href);
      url.searchParams.set("view", mode);
      url.searchParams.set("standalone", "1");
      window.open(url.toString(), "_blank", "width=1480,height=920");
    }

    document.querySelectorAll("[data-open]").forEach((button) => {
      button.addEventListener("click", () => {
        const mode = button.dataset.open;
        if (mode === "management" || mode === "ledger" || mode === "userAdmin" || mode === "leave" || mode === "backup" || mode === "systemUpdate") {
          showWorkspace(mode);
          return;
        }
        openModal(mode);
      });
    });
    document.querySelectorAll("[data-view]").forEach((button) => {
      button.addEventListener("click", () => {
        showWorkspace(button.dataset.view);
        if (button.id === "noticeNavToggle") document.querySelector("#noticeNavGroup").classList.toggle("open");
      });
    });
    document.querySelectorAll("[data-open-window]").forEach((button) => {
      button.addEventListener("click", () => openWorkspaceWindow(button.dataset.openWindow));
    });
    document.querySelector("#orderNavToggle").addEventListener("click", () => {
      document.querySelector("#orderNavGroup").classList.toggle("open");
    });
    document.querySelector("#noticeInputOpen").addEventListener("click", openNoticePopup);
    noticePopupClose.addEventListener("click", closeNoticePopup);
    noticePopup.addEventListener("click", (event) => {
      if (event.target === noticePopup) closeNoticePopup();
    });
    document.querySelector("#closeModal").addEventListener("click", closeModal);
    document.querySelector("#cancel").addEventListener("click", closeModal);
    modal.addEventListener("click", (event) => {
      if (event.target === modal) closeModal();
    });
    fileInput.addEventListener("change", () => {
      dropMain.textContent = fileInput.files[0] ? fileInput.files[0].name : "파일을 선택하거나 여기에 올려주세요.";
    });
    templateInput.addEventListener("change", () => {
      templateDropMain.textContent = templateInput.files[0] ? templateInput.files[0].name : "롯데택배 발주서 양식을 선택해주세요.";
    });

    function setupDropzone(dropzone, input, label, fallbackText) {
      ["dragenter", "dragover"].forEach((eventName) => {
        dropzone.addEventListener(eventName, (event) => {
          event.preventDefault();
          event.stopPropagation();
          dropzone.classList.add("dragover");
        });
      });
      ["dragleave", "drop"].forEach((eventName) => {
        dropzone.addEventListener(eventName, (event) => {
          event.preventDefault();
          event.stopPropagation();
          dropzone.classList.remove("dragover");
        });
      });
      dropzone.addEventListener("drop", (event) => {
        const files = event.dataTransfer.files;
        if (!files || files.length === 0) return;
        input.files = files;
        label.textContent = files[0] ? files[0].name : fallbackText;
        input.dispatchEvent(new Event("change", { bubbles: true }));
      });
    }

    setupDropzone(
      document.querySelector("label[for='fileInput']"),
      fileInput,
      dropMain,
      "파일을 선택하거나 여기에 올려주세요."
    );
    setupDropzone(
      document.querySelector("label[for='templateInput']"),
      templateInput,
      templateDropMain,
      "롯데택배 발주서 양식을 선택해주세요."
    );
    setupDropzone(
      document.querySelector("label[for='vendorContactsFileInput']"),
      vendorContactsFileInput,
      vendorContactsDropMain,
      "업체명/메일주소 엑셀을 선택해주세요."
    );
    setupDropzone(
      document.querySelector("label[for='ledgerImportInput']"),
      ledgerImportInput,
      ledgerImportDropMain,
      "업로드"
    );
    setupDropzone(
      document.querySelector("label[for='managementImportInput']"),
      managementImportInput,
      managementImportDropMain,
      "업로드"
    );
    document.querySelector("#addProductRow").addEventListener("click", () => addProductRow());
    noticeSaveButton.addEventListener("click", saveNoticeTemplate);
    noticeClearButton.addEventListener("click", clearNoticeTemplate);
    [noticeDateInput, noticeTitleInput, noticeOwnerInput, noticeBodyInput]
      .forEach((input) => input.addEventListener("input", renderNoticePreview));
    receiptTypeSelect.addEventListener("change", resetProductRows);
    vendorContactSelect.addEventListener("change", applySelectedVendor);
    saveVendorContactButton.addEventListener("click", saveCurrentVendorContact);
    vendorContactsFileInput.addEventListener("change", uploadVendorContactsWorkbook);
    saveCsCaseButton.addEventListener("click", saveCurrentCsCase);
    ledgerRefresh.addEventListener("click", loadLedgerCases);
    ledgerStatusFilter.addEventListener("change", applyLedgerFilters);
    ledgerFilterButtons.forEach((button) => {
      button.addEventListener("click", (event) => {
        event.stopPropagation();
        openLedgerFilter(button);
      });
    });
    managementFilterButtons.forEach((button) => {
      button.addEventListener("click", (event) => {
        event.stopPropagation();
        openManagementFilter(button);
      });
    });
    ledgerFilterSearch.addEventListener("input", () => {
      refreshActiveFilterOptions();
    });
    ledgerFilterSearch.addEventListener("keydown", (event) => {
      if (event.key === "Enter") setActivePopoverFilter(ledgerFilterSearch.value);
      if (event.key === "Escape") closeLedgerFilter();
    });
    ledgerFilterApply.addEventListener("click", () => setActivePopoverFilter(ledgerFilterSearch.value));
    ledgerFilterClear.addEventListener("click", () => setActivePopoverFilter(""));
    ledgerFilterOptions.addEventListener("click", (event) => {
      const option = event.target.closest("[data-filter-value]");
      if (option) setActivePopoverFilter(option.dataset.filterValue || "");
    });
    document.addEventListener("click", (event) => {
      if (
        !ledgerFilterPopover.contains(event.target)
        && !event.target.closest("[data-ledger-filter-button]")
        && !event.target.closest("[data-management-filter-button]")
      ) {
        closeLedgerFilter();
      }
    });
    ledgerAddCs.addEventListener("click", openLedgerCsPopup);
    ledgerCsPopupClose.addEventListener("click", closeLedgerCsPopup);
    ledgerImportInput.addEventListener("change", uploadLedgerWorkbook);
    managementRefresh.addEventListener("click", loadManagementRecords);
    managementImportInput.addEventListener("change", uploadManagementWorkbook);
    managementSaveAll.addEventListener("click", () => saveCurrentWorkspaceRows({ mode: "management", selectedOnly: true }));
    ledgerSaveAll.addEventListener("click", () => saveCurrentWorkspaceRows({ mode: "ledger", selectedOnly: true }));
    managementDeleteSelected.addEventListener("click", () => deleteSelectedRows("management"));
    ledgerDeleteSelected.addEventListener("click", () => deleteSelectedRows("ledger"));
    if (userAdminRefresh) userAdminRefresh.addEventListener("click", loadUserAccounts);
    if (userAdminSave) userAdminSave.addEventListener("click", saveUserAccount);
    if (userAdminRole) userAdminRole.addEventListener("change", syncPermissionChecksForRole);
    if (backupRefresh) backupRefresh.addEventListener("click", loadBackups);
    if (backupCreate) backupCreate.addEventListener("click", createBackupNow);
    if (backupRestoreInput) backupRestoreInput.addEventListener("change", restoreBackupFromUpload);
    if (systemUpdateRefresh) systemUpdateRefresh.addEventListener("click", loadSystemUpdateStatus);
    if (systemUpdateCheck) systemUpdateCheck.addEventListener("click", checkSystemUpdate);
    if (systemUpdateApply) systemUpdateApply.addEventListener("click", applySystemUpdate);
    if (backupBody) {
      backupBody.addEventListener("click", (event) => {
        const restoreButton = event.target.closest("[data-backup-restore]");
        const downloadButton = event.target.closest("[data-backup-download]");
        const deleteButton = event.target.closest("[data-backup-delete]");
        if (restoreButton) restoreBackupByName(restoreButton.dataset.backupRestore);
        if (downloadButton) downloadBackup(downloadButton.dataset.backupDownload);
        if (deleteButton) deleteBackup(deleteButton.dataset.backupDelete);
      });
    }
    if (userAdminBody) {
      userAdminBody.addEventListener("click", (event) => {
        const editButton = event.target.closest("[data-user-edit]");
        if (editButton) editUserAccount(editButton.dataset.userEdit);
      });
    }
    leaveTabs.forEach((button) => {
      button.addEventListener("click", () => setLeaveTab(button.dataset.leaveTab));
    });
    if (leaveRefresh) leaveRefresh.addEventListener("click", loadLeaveData);
    if (leaveRequestSubmit) leaveRequestSubmit.addEventListener("click", submitLeaveRequest);
    if (leaveUnitSelect) leaveUnitSelect.addEventListener("change", syncHalfDayDates);
    if (leaveStartDate) leaveStartDate.addEventListener("change", syncHalfDayDates);
    if (leaveBalanceSave) leaveBalanceSave.addEventListener("click", saveLeaveBalance);
    if (leaveUsageSave) leaveUsageSave.addEventListener("click", saveHistoricalLeaveUsage);
    if (leaveApprovalBody) {
      leaveApprovalBody.addEventListener("click", (event) => {
        const button = event.target.closest("[data-leave-decision]");
        if (button) decideLeaveRequest(button.dataset.leaveId, button.dataset.leaveDecision);
      });
    }
    managementPageSize.addEventListener("change", loadManagementRecords);
    ledgerPageSize.addEventListener("change", loadLedgerCases);
    ledgerYearFilter.addEventListener("change", loadLedgerCases);
    managementYearFilter.addEventListener("change", () => {
      renderManagementPeriodControls();
      loadManagementRecords();
    });
    ledgerMonthFilter.addEventListener("change", () => {
      ensureYearForMonth(ledgerYearFilter, ledgerMonthFilter);
      loadLedgerCases();
    });
    managementMonthFilter.addEventListener("change", () => {
      const period = selectedManagementPeriod();
      managementYearFilter.value = period.year || "";
      renderManagementMonthTabs();
      loadManagementRecords();
    });
    if (managementMonthTabs) {
      managementMonthTabs.addEventListener("click", (event) => {
        const button = event.target.closest("[data-management-month]");
        if (!button) return;
        setManagementPeriod(button.dataset.managementYear || "", button.dataset.managementMonth || "");
        renderManagementPeriodControls();
        loadManagementRecords();
      });
    }
    managementDownloadSelected.addEventListener("click", () => downloadManagementExcel("selected", managementDownloadSelected));
    managementDownloadMonth.addEventListener("click", () => downloadManagementExcel("month", managementDownloadMonth));
    managementDownloadYear.addEventListener("click", () => downloadManagementExcel("year", managementDownloadYear));
    managementDownloadAll.addEventListener("click", () => downloadManagementExcel("all", managementDownloadAll));
    ledgerDownloadExcel.addEventListener("click", () => {
      downloadExcel("/api/cs-cases-export", collectLedgerExportRows(), "CS처리대장.xlsx", ledgerDownloadExcel);
    });
    managementBody.addEventListener("input", (event) => {
      if (event.target.closest("[data-management-field]")) markRowDirty(event.target.closest("tr"));
    });
    managementBody.addEventListener("click", (event) => {
      const csButton = event.target.closest(".management-cs-button");
      if (csButton) receiveManagementCs(csButton);
    });
    ledgerBody.addEventListener("input", (event) => {
      if (event.target.closest("[data-field]")) markRowDirty(event.target.closest("tr"));
    });
    ledgerBody.addEventListener("change", (event) => {
      const field = event.target.closest('[data-field="status"], [data-field="cs_type"]');
      if (field) {
        markRowDirty(field.closest("tr"));
        updateLedgerRowCompletion(field.closest("tr"));
      }
    });
    ledgerSearchInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        loadLedgerCases();
      }
    });
    managementSearchInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        loadManagementRecords();
      }
    });
    vendorNameInput.addEventListener("input", () => {
      csSubjectInput.value = defaultCsSubject(vendorNameInput.value.trim());
    });
    [csOriginInput, csProductInput, csReceiverInput, csPhoneInput, csAddressInput, csContentInput]
      .forEach((input) => input.addEventListener("input", refreshCsBody));

    setInterval(() => {
      saveCurrentWorkspaceRows({ silent: true });
    }, 10 * 60 * 1000);

    uploadForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const formData = new FormData(uploadForm);
      notice.textContent = "처리 중입니다.";
      submitButton.disabled = true;
      try {
        if (currentMode === "ledger" || currentMode === "management") {
          closeModal();
        } else if (currentMode === "cs") {
          refreshCsBody();
          const payload = collectCsPayload();
          if (!payload.recipient_email || !payload.subject || !payload.body) {
            throw new Error("받는 업체 메일, 제목, 요청 내용을 입력해주세요.");
          }
          const response = await fetch("/api/cs-mail", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
          });
          const data = await response.json();
          if (!response.ok) throw new Error(data.error || "메일 전송에 실패했습니다.");
          notice.textContent = data.message || "메일 전송이 완료되었습니다.";
          await loadCsCases();
        } else if (currentMode === "vehicle") {
          const payload = collectVehiclePayload();
          if (!payload.supplier || !payload.delivery_place || !payload.manager || payload.items.length === 0) {
            throw new Error("공급받는자, 제품, 납품장소, 담당자명을 입력해주세요.");
          }
          const response = await fetch("/api/vehicle-receipt", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
          });
          if (!response.ok) {
            const data = await response.json();
            throw new Error(data.error || "처리에 실패했습니다.");
          }
          const blob = await response.blob();
          const url = URL.createObjectURL(blob);
          const link = document.createElement("a");
          link.href = url;
          link.download = filenameFromResponse(response, "차량인수증.xlsx");
          document.body.appendChild(link);
          link.click();
          link.remove();
          URL.revokeObjectURL(url);
          notice.textContent = "차량인수증 다운로드가 시작되었습니다.";
        } else if (currentMode === "delivery") {
          if (!fileInput.files[0]) throw new Error("주소일브릿지 엑셀 파일을 선택해주세요.");
          const response = await fetch("/api/delivery-summary", { method: "POST", body: formData });
          const data = await response.json();
          if (!response.ok) throw new Error(data.error || "처리에 실패했습니다.");
          resultText.value = data.text;
          result.classList.add("open");
          notice.textContent = `${data.line_count}개 묶음이 생성되었습니다.`;
        } else {
          if (!fileInput.files[0]) {
            throw new Error(currentMode === "invoice" ? "출고송장 엑셀 파일을 선택해주세요." : "주소일브릿지 원본 엑셀 파일을 선택해주세요.");
          }
          const endpoint = currentMode === "invoice" ? "/api/invoice-export" : "/api/lotte-order-form";
          const response = await fetch(endpoint, { method: "POST", body: formData });
          if (!response.ok) {
            const data = await response.json();
            throw new Error(data.error || "처리에 실패했습니다.");
          }
          const blob = await response.blob();
          const url = URL.createObjectURL(blob);
          const link = document.createElement("a");
          link.href = url;
          link.download = filenameFromResponse(
            response,
            currentMode === "invoice" ? "송장번호_추출.xlsx" : "롯데택배_발주서.xlsx"
          );
          document.body.appendChild(link);
          link.click();
          link.remove();
          URL.revokeObjectURL(url);
          notice.textContent = "엑셀 파일 다운로드가 시작되었습니다.";
        }
      } catch (error) {
        notice.textContent = error.message;
      } finally {
        submitButton.disabled = false;
      }
    });

    document.querySelector("#copyResult").addEventListener("click", async () => {
      await navigator.clipboard.writeText(resultText.value);
      notice.textContent = "텍스트를 복사했습니다.";
    });

    document.querySelector("#downloadText").addEventListener("click", () => {
      const blob = new Blob([resultText.value], { type: "text/plain;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "개별_택배건_요약.txt";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    });

    const initialView = new URLSearchParams(window.location.search).get("view");
    showWorkspace(["management", "ledger", "leave", "userAdmin", "backup", "systemUpdate"].includes(initialView) ? initialView : "dashboard");
  </script>
</body>
</html>
"""

LOGIN_HTML = r"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>(주)소일브릿지 업무자동화 로그인</title>
  <style>
    :root {
      --bg: #f5f7fb;
      --panel: #ffffff;
      --line: #d8dee9;
      --text: #111827;
      --muted: #667085;
      --blue: #2563eb;
      --green: #079455;
      font-family: Pretendard, Inter, "Noto Sans KR", "Malgun Gothic", Arial, sans-serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      background:
        radial-gradient(circle at 20% 20%, rgba(37, 99, 235, .13), transparent 28%),
        radial-gradient(circle at 80% 10%, rgba(7, 148, 85, .12), transparent 25%),
        var(--bg);
      color: var(--text);
    }
    .login-shell {
      width: min(420px, calc(100vw - 32px));
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 14px;
      box-shadow: 0 22px 60px rgba(15, 23, 42, .16);
      padding: 34px 32px 30px;
    }
    .brand {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 26px;
    }
    .brand-mark {
      width: 44px;
      height: 44px;
      border-radius: 12px;
      display: grid;
      place-items: center;
      background: linear-gradient(145deg, #2563eb, #079455);
      color: white;
      font-weight: 900;
    }
    .brand-title { font-size: 20px; font-weight: 900; line-height: 1.32; }
    .brand-sub { margin-top: 4px; color: var(--muted); font-size: 13px; font-weight: 700; }
    h1 { margin: 0 0 8px; font-size: 26px; }
    .lead { margin: 0 0 24px; color: var(--muted); font-size: 14px; line-height: 1.55; }
    label { display: block; margin: 14px 0 7px; font-size: 13px; font-weight: 850; color: #344054; }
    input {
      width: 100%;
      height: 44px;
      border: 1px solid #cfd6e2;
      border-radius: 9px;
      padding: 0 13px;
      font-size: 15px;
      font-weight: 700;
      outline: none;
    }
    input:focus {
      border-color: var(--blue);
      box-shadow: 0 0 0 3px rgba(37, 99, 235, .13);
    }
    .error {
      display: __ERROR_DISPLAY__;
      margin: 0 0 14px;
      padding: 11px 12px;
      border-radius: 8px;
      background: #fee2e2;
      color: #b42318;
      font-size: 13px;
      font-weight: 850;
    }
    button {
      width: 100%;
      height: 46px;
      margin-top: 22px;
      border: 0;
      border-radius: 9px;
      background: linear-gradient(135deg, #2563eb, #079455);
      color: white;
      font-size: 15px;
      font-weight: 900;
      cursor: pointer;
    }
    .hint {
      margin-top: 18px;
      padding: 12px;
      border-radius: 8px;
      background: #f8fafc;
      border: 1px solid #e5e7eb;
      color: #475467;
      font-size: 12px;
      line-height: 1.55;
    }
    .hint strong { color: #111827; }
  </style>
</head>
<body>
  <main class="login-shell">
    <div class="brand">
      <div class="brand-mark">SB</div>
      <div>
        <div class="brand-title">(주)소일브릿지<br>업무자동화</div>
        <div class="brand-sub">관리자 / 사용자 로그인</div>
      </div>
    </div>
    <h1>로그인</h1>
    <p class="lead">업무 화면을 사용하려면 계정으로 로그인해주세요.</p>
    <div class="error">아이디 또는 비밀번호가 올바르지 않습니다.</div>
    <form method="post" action="/login">
      <label for="username">아이디</label>
      <input id="username" name="username" type="text" autocomplete="username" autofocus />
      <label for="password">비밀번호</label>
      <input id="password" name="password" type="password" autocomplete="current-password" />
      <button type="submit">로그인</button>
    </form>
    <div class="hint">
      기본 계정: <strong>admin / admin1234</strong><br>
      사용자 계정: <strong>user / user1234</strong>
    </div>
  </main>
</body>
</html>
"""

ADMIN_NAV_HTML = r"""
      <button class="nav-item" type="button" data-open="userAdmin"><i data-lucide="info"></i> <span>사용자 관리</span></button>
"""

BACKUP_NAV_HTML = r"""
      <button class="nav-item" type="button" data-open="backup"><i data-lucide="database"></i> <span>백업 관리</span></button>
"""

SYSTEM_NAV_HTML = r"""
      <button class="nav-item" type="button" data-open="systemUpdate"><i data-lucide="refresh-cw"></i> <span>업데이트 관리</span></button>
"""

LEAVE_NAV_HTML = r"""
      <button class="nav-item" type="button" data-open="leave"><i data-lucide="calendar-days"></i> <span>연차 관리</span></button>
"""

LEAVE_WORKSPACE_HTML = r"""
      <section class="workspace-view" id="leaveWorkspace">
        <div class="workspace-head">
          <div class="workspace-title">연차 관리</div>
          <div class="workspace-actions">
            <button class="workspace-button" type="button" id="leaveRefresh">새로고침</button>
          </div>
        </div>
        <div class="workspace-mount">
          <div class="leave-panel">
            <div class="leave-summary-grid">
              <article class="leave-summary-card">
                <span>시작 연차</span>
                <strong id="leaveTotalDays">0일</strong>
              </article>
              <article class="leave-summary-card">
                <span>사용 연차</span>
                <strong id="leaveUsedDays">0일</strong>
              </article>
              <article class="leave-summary-card accent">
                <span>남은 연차</span>
                <strong id="leaveRemainingDays">0일</strong>
              </article>
            </div>
            <div class="leave-tabs">
              <button class="leave-tab active" type="button" data-leave-tab="mine">내 연차</button>
              <button class="leave-tab" type="button" data-leave-tab="request">연차 신청</button>
              <button class="leave-tab" type="button" data-leave-tab="approvals">승인 대기</button>
              <button class="leave-tab" type="button" data-leave-tab="admin">직원별 관리</button>
            </div>
            <div class="leave-message" id="leaveMessage"></div>
            <section class="leave-tab-panel active" id="leaveTabMine">
              <div class="leave-grid two">
                <div class="leave-card">
                  <div class="leave-card-title">연차 잔여 현황</div>
                  <table class="leave-table">
                    <thead><tr><th>유형</th><th>시작</th><th>사용</th><th>잔여</th></tr></thead>
                    <tbody id="leaveBalanceBody"></tbody>
                  </table>
                </div>
                <div class="leave-card">
                  <div class="leave-card-title">연차 사용/신청 이력</div>
                  <table class="leave-table">
                    <thead><tr><th>기간</th><th>구분</th><th>수량</th><th>상태</th><th>사유</th></tr></thead>
                    <tbody id="leaveHistoryBody"></tbody>
                  </table>
                </div>
              </div>
            </section>
            <section class="leave-tab-panel" id="leaveTabRequest">
              <div class="leave-card narrow">
                <div class="leave-card-title">연차 신청</div>
                <div class="leave-form">
                  <label>휴가 유형<select id="leaveTypeSelect"></select></label>
                  <label>단위<select id="leaveUnitSelect"><option value="FULL_DAY">연차</option><option value="HALF_DAY">반차</option></select></label>
                  <label>시작일<input id="leaveStartDate" type="date" /></label>
                  <label>종료일<input id="leaveEndDate" type="date" /></label>
                  <label class="wide">사유<textarea id="leaveReasonInput" placeholder="연차 사유를 입력해주세요."></textarea></label>
                  <button class="workspace-button" type="button" id="leaveRequestSubmit">신청하기</button>
                </div>
              </div>
            </section>
            <section class="leave-tab-panel" id="leaveTabApprovals">
              <div class="leave-card">
                <div class="leave-card-title">승인 대기</div>
                <table class="leave-table">
                  <thead><tr><th>신청자</th><th>기간</th><th>구분</th><th>수량</th><th>사유</th><th>처리</th></tr></thead>
                  <tbody id="leaveApprovalBody"></tbody>
                </table>
              </div>
            </section>
            <section class="leave-tab-panel" id="leaveTabAdmin">
              <div class="leave-grid two">
                <div class="leave-card">
                  <div class="leave-card-title">직원 연차 기준 설정</div>
                  <div class="leave-form">
                    <label>직원<select id="leaveAdminUserSelect"></select></label>
                    <label>시작 연차<input id="leaveAdminTotalInput" type="number" min="0" step="0.5" value="10" /></label>
                    <label>기존 사용 연차<input id="leaveAdminUsedInput" type="number" min="0" step="0.5" value="0" /></label>
                    <button class="workspace-button" type="button" id="leaveBalanceSave">연차 기준 저장</button>
                  </div>
                </div>
                <div class="leave-card">
                  <div class="leave-card-title">기존 사용 연차 날짜별 입력</div>
                  <div class="leave-form">
                    <label>직원<select id="leaveUsageUserSelect"></select></label>
                    <label class="wide">사용 일자 목록<textarea id="leaveUsageDatesInput" placeholder="2026-01-03&#10;2026-02-14 반차&#10;2026-03-05 0.5"></textarea></label>
                    <label class="wide">메모<input id="leaveUsageNoteInput" type="text" placeholder="예: 시스템 도입 전 사용분" /></label>
                    <button class="workspace-button" type="button" id="leaveUsageSave">사용 일자 등록</button>
                  </div>
                </div>
              </div>
              <div class="leave-card">
                <div class="leave-card-title">직원별 연차 현황</div>
                <table class="leave-table">
                  <thead><tr><th>직원</th><th>아이디</th><th>시작</th><th>사용</th><th>잔여</th></tr></thead>
                  <tbody id="leaveAdminBalanceBody"></tbody>
                </table>
              </div>
            </section>
          </div>
        </div>
      </section>
"""

ADMIN_WORKSPACE_HTML = r"""
      <section class="workspace-view" id="userAdminWorkspace">
        <div class="workspace-head">
          <div class="workspace-title">사용자 관리</div>
          <div class="workspace-actions">
            <button class="workspace-button" type="button" id="userAdminRefresh">새로고침</button>
          </div>
        </div>
        <div class="workspace-mount">
          <div class="admin-panel">
            <div class="admin-form">
              <input id="userAdminId" type="hidden" />
              <label>아이디
                <input id="userAdminUsername" type="text" placeholder="예) hong" />
              </label>
              <label>표시 이름
                <input id="userAdminDisplayName" type="text" placeholder="예) 홍길동" />
              </label>
              <label>권한
                <select id="userAdminRole">
                  <option value="user">사용자</option>
                  <option value="admin">관리자</option>
                </select>
              </label>
              <label>비밀번호
                <input id="userAdminPassword" type="password" placeholder="신규/변경 시 입력" />
              </label>
              <label class="admin-check"><input id="userAdminActive" type="checkbox" checked /> 사용</label>
              <button class="workspace-button" type="button" id="userAdminSave">저장</button>
            </div>
            <div class="permission-grid" id="userAdminPermissions">
              __PERMISSION_CHECKBOXES__
            </div>
            <div class="admin-message" id="userAdminMessage"></div>
            <div class="admin-table-wrap">
              <table class="admin-table">
                <thead>
                  <tr>
                    <th>아이디</th>
                    <th>표시 이름</th>
                    <th>권한</th>
                    <th>세부권한</th>
                    <th>상태</th>
                    <th>생성일</th>
                    <th>수정</th>
                  </tr>
                </thead>
                <tbody id="userAdminBody"></tbody>
              </table>
            </div>
          </div>
        </div>
      </section>
"""

BACKUP_WORKSPACE_HTML = r"""
      <section class="workspace-view" id="backupWorkspace">
        <div class="workspace-head">
          <div class="workspace-title">백업 관리</div>
          <div class="workspace-actions">
            <button class="workspace-button" type="button" id="backupRefresh">새로고침</button>
            <button class="workspace-button" type="button" id="backupCreate">지금 백업하기</button>
            <label class="workspace-button" for="backupRestoreInput">백업 파일 복원</label>
            <input id="backupRestoreInput" type="file" accept=".zip" />
          </div>
        </div>
        <div class="workspace-mount">
          <div class="backup-panel">
            <div class="backup-summary-grid">
              <article class="backup-summary-card">
                <span>자동 백업</span>
                <strong>매일 03:00</strong>
              </article>
              <article class="backup-summary-card">
                <span>보관 기준</span>
                <strong>최근 90일</strong>
              </article>
              <article class="backup-summary-card">
                <span>백업 위치</span>
                <strong id="backupPath">-</strong>
              </article>
            </div>
            <div class="backup-card">
              <p class="backup-note">
                백업에는 업무 DB, 메일 설정, 업체 주소록, 암호화 키가 포함됩니다. 나스에서는 이 백업 폴더를 Synology Hyper Backup 대상으로 잡으면 됩니다.
                복원 전에는 현재 데이터를 자동으로 한 번 더 백업합니다.
              </p>
              <div class="backup-message" id="backupMessage"></div>
              <table class="backup-table">
                <thead>
                  <tr>
                    <th>백업 파일</th>
                    <th>생성일</th>
                    <th>크기</th>
                    <th>관리</th>
                  </tr>
                </thead>
                <tbody id="backupBody"></tbody>
              </table>
            </div>
          </div>
        </div>
      </section>
"""

SYSTEM_WORKSPACE_HTML = r"""
      <section class="workspace-view" id="systemUpdateWorkspace">
        <div class="workspace-head">
          <div class="workspace-title">업데이트 관리</div>
          <div class="workspace-actions">
            <button class="workspace-button" type="button" id="systemUpdateRefresh">새로고침</button>
            <button class="workspace-button" type="button" id="systemUpdateCheck">업데이트 확인</button>
            <button class="workspace-button" type="button" id="systemUpdateApply">업데이트 적용</button>
          </div>
        </div>
        <div class="workspace-mount">
          <div class="system-panel">
            <div class="system-summary-grid">
              <article class="system-summary-card">
                <span>실행 위치</span>
                <strong id="systemUpdateSource">-</strong>
              </article>
              <article class="system-summary-card">
                <span>브랜치</span>
                <strong id="systemUpdateBranch">-</strong>
              </article>
              <article class="system-summary-card">
                <span>현재 버전</span>
                <strong id="systemUpdateCurrent">-</strong>
              </article>
              <article class="system-summary-card">
                <span>업데이트 상태</span>
                <strong id="systemUpdateState">-</strong>
              </article>
            </div>
            <div class="system-card">
              <p class="system-note">
                업데이트 적용 시 현재 업무 데이터가 자동 백업된 뒤 GitHub 최신 코드가 적용됩니다. 적용 후에는 프로그램/나스 서비스를 다시 시작해주세요.
              </p>
              <div class="system-message" id="systemUpdateMessage"></div>
              <table class="system-table">
                <thead>
                  <tr>
                    <th>일시</th>
                    <th>작업</th>
                    <th>결과</th>
                    <th>이전 버전</th>
                    <th>적용 버전</th>
                    <th>백업</th>
                    <th>메시지</th>
                  </tr>
                </thead>
                <tbody id="systemUpdateHistoryBody"></tbody>
              </table>
            </div>
          </div>
        </div>
      </section>
"""


def safe_filename(filename: str) -> str:
    filename = Path(filename).name
    filename = re.sub(r'[<>:"/\\|?*]+', "_", filename)
    return filename or "upload.xlsx"


def original_uploaded_filename(filename: str) -> str:
    return re.sub(r"^\d{10,}_", "", filename)


def parse_multipart(headers, rfile) -> dict[str, tuple[str, bytes] | str]:
    content_type = headers.get("Content-Type", "")
    boundary_match = re.search(r"boundary=(.+)", content_type)
    if not boundary_match:
        raise ValueError("업로드 형식이 올바르지 않습니다.")

    boundary = boundary_match.group(1).strip().strip('"').encode()
    length = int(headers.get("Content-Length", "0"))
    body = rfile.read(length)
    fields: dict[str, tuple[str, bytes] | str] = {}

    for part in body.split(b"--" + boundary):
        part = part.strip()
        if not part or part == b"--":
            continue
        if part.endswith(b"--"):
            part = part[:-2].strip()
        if b"\r\n\r\n" not in part:
            continue

        raw_headers, data = part.split(b"\r\n\r\n", 1)
        data = data.rstrip(b"\r\n")
        disposition = ""
        for line in raw_headers.decode("utf-8", errors="ignore").split("\r\n"):
            if line.lower().startswith("content-disposition:"):
                disposition = line
                break

        name_match = re.search(r'name="([^"]+)"', disposition)
        if not name_match:
            continue
        field_name = name_match.group(1)
        filename_match = re.search(r'filename="([^"]*)"', disposition)
        if filename_match:
            fields[field_name] = (safe_filename(filename_match.group(1)), data)
        else:
            fields[field_name] = data.decode("utf-8", errors="ignore")

    return fields


def save_uploaded_file(fields: dict[str, tuple[str, bytes] | str], field_name: str = "file") -> Path:
    uploaded = fields.get(field_name)
    if not isinstance(uploaded, tuple):
        raise ValueError("업로드된 파일이 없습니다.")

    filename, data = uploaded
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("xlsx 또는 xlsm 파일만 업로드할 수 있습니다.")

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    path = UPLOAD_DIR / f"{int(time.time() * 1000)}_{filename}"
    path.write_bytes(data)
    return path


def save_uploaded_backup_zip(fields: dict[str, tuple[str, bytes] | str], field_name: str = "file") -> Path:
    uploaded = fields.get(field_name)
    if not isinstance(uploaded, tuple):
        raise ValueError("업로드된 백업 파일이 없습니다.")

    filename, data = uploaded
    if not filename.lower().endswith(".zip"):
        raise ValueError("zip 백업 파일만 업로드할 수 있습니다.")

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    path = UPLOAD_DIR / f"{int(time.time() * 1000)}_{safe_filename(filename)}"
    path.write_bytes(data)
    return path


def parse_receipt_date(value: object) -> date:
    if isinstance(value, str) and value:
        try:
            return date.fromisoformat(value)
        except ValueError:
            pass
    return date.today()


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clean_payload_text(payload: dict, key: str) -> str:
    return str(payload.get(key, "") or "").strip()


def default_permissions_for_role(role: str) -> list[str]:
    return list(DEFAULT_ROLE_PERMISSIONS.get(role, DEFAULT_ROLE_PERMISSIONS["user"]))


def normalize_permissions(value: object, role: str = "user") -> list[str]:
    if role == "admin":
        return list(ALL_PERMISSIONS)
    if isinstance(value, str):
        try:
            value = json.loads(value) if value else []
        except json.JSONDecodeError:
            value = []
    if not isinstance(value, list):
        value = default_permissions_for_role(role)
    allowed = {str(item) for item in value if str(item) in ALL_PERMISSIONS}
    return [key for key in ALL_PERMISSIONS if key in allowed]


def permissions_html() -> str:
    return "\n".join(
        f"""<label class="permission-item"><input type="checkbox" value="{key}" data-permission-check /> <span>{label}<small>{description}</small></span></label>"""
        for key, label, description in PERMISSION_DEFINITIONS
    )


def role_label(role: str) -> str:
    return "관리자" if role == "admin" else "사용자"


def render_app_html(user: dict[str, str]) -> str:
    display_name = user.get("display_name") or user.get("username") or "사용자"
    role = role_label(user.get("role", "user"))
    display = display_name if display_name == role else f"{display_name} · {role}"
    permissions = normalize_permissions(user.get("permissions", []), user.get("role", "user"))
    leave_enabled = any(permission in permissions for permission in ("leave_view", "leave_approve", "leave_manage"))
    leave_nav = LEAVE_NAV_HTML if leave_enabled else ""
    leave_workspace = LEAVE_WORKSPACE_HTML if leave_enabled else ""
    admin_nav = ADMIN_NAV_HTML if "user_admin" in permissions else ""
    admin_workspace = ADMIN_WORKSPACE_HTML.replace("__PERMISSION_CHECKBOXES__", permissions_html()) if "user_admin" in permissions else ""
    backup_nav = BACKUP_NAV_HTML if "backup_manage" in permissions else ""
    backup_workspace = BACKUP_WORKSPACE_HTML if "backup_manage" in permissions else ""
    system_nav = SYSTEM_NAV_HTML if "system_update" in permissions else ""
    system_workspace = SYSTEM_WORKSPACE_HTML if "system_update" in permissions else ""
    return (
        HTML
        .replace("__USER_DISPLAY__", html_escape(display))
        .replace("__LEAVE_NAV__", leave_nav)
        .replace("__LEAVE_WORKSPACE__", leave_workspace)
        .replace("__ADMIN_NAV__", admin_nav)
        .replace("__ADMIN_WORKSPACE__", admin_workspace)
        .replace("__BACKUP_NAV__", backup_nav)
        .replace("__BACKUP_WORKSPACE__", backup_workspace)
        .replace("__SYSTEM_NAV__", system_nav)
        .replace("__SYSTEM_WORKSPACE__", system_workspace)
        .replace("__USER_PERMISSIONS__", json.dumps(permissions, ensure_ascii=False))
        .replace("__PERMISSION_LABELS__", json.dumps({key: label for key, label, _ in PERMISSION_DEFINITIONS}, ensure_ascii=False))
    )


def render_login_html(show_error: bool = False) -> str:
    return LOGIN_HTML.replace("__ERROR_DISPLAY__", "block" if show_error else "none")


def password_hash(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    iterations = 260000
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), iterations)
    return f"pbkdf2_sha256${iterations}${salt}${digest.hex()}"


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        algorithm, iterations_text, salt, expected = stored_hash.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        iterations = int(iterations_text)
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), iterations)
        return hmac.compare_digest(digest.hex(), expected)
    except ValueError:
        return False


def token_digest(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def parse_iso_date(value: object) -> date:
    try:
        return date.fromisoformat(str(value or "").strip())
    except ValueError as exc:
        raise ValueError("날짜 형식이 올바르지 않습니다.") from exc


def is_half_step(value: float) -> bool:
    return abs(value * 2 - round(value * 2)) < 0.000001


def clean_leave_days(value: object, label: str) -> float:
    try:
        days = float(str(value).strip())
    except ValueError as exc:
        raise ValueError(f"{label}는 숫자로 입력해주세요.") from exc
    if days < 0 or not is_half_step(days):
        raise ValueError(f"{label}는 0 이상, 0.5일 단위로 입력해주세요.")
    return round(days, 2)


def calculate_leave_days(start: date, end: date, unit: str) -> float:
    if start > end:
        raise ValueError("종료일은 시작일보다 빠를 수 없습니다.")
    if unit == "HALF_DAY":
        if start != end:
            raise ValueError("반차는 하루 안에서만 신청할 수 있습니다.")
        if start.weekday() >= 5:
            raise ValueError("주말에는 반차를 신청할 수 없습니다.")
        return 0.5
    days = 0
    cursor = start
    while cursor <= end:
        if cursor.weekday() < 5:
            days += 1
        cursor = date.fromordinal(cursor.toordinal() + 1)
    if days <= 0:
        raise ValueError("신청 가능한 평일이 없습니다.")
    return float(days)


def leave_status_label(status: str) -> str:
    return {
        "PENDING": "승인대기",
        "APPROVED": "승인완료",
        "REJECTED": "반려",
        "CANCELED": "취소",
        "HISTORICAL": "기존사용",
    }.get(status, status)


def connect_db() -> sqlite3.Connection:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def init_db() -> None:
    connection = connect_db()
    try:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                display_name TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                permissions TEXT,
                password_hash TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        user_columns = {
            row["name"] for row in connection.execute("PRAGMA table_info(users)").fetchall()
        }
        if "permissions" not in user_columns:
            connection.execute("ALTER TABLE users ADD COLUMN permissions TEXT")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS login_sessions (
                token_hash TEXT PRIMARY KEY,
                username TEXT NOT NULL,
                created_at REAL NOT NULL,
                expires_at REAL NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS system_update_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                action TEXT NOT NULL,
                before_commit TEXT,
                after_commit TEXT,
                status TEXT NOT NULL,
                message TEXT,
                backup_name TEXT,
                created_at TEXT NOT NULL
            )
            """
        )
        now = now_text()
        for username, display_name, role, default_password in DEFAULT_USERS:
            exists = connection.execute("SELECT id, permissions FROM users WHERE username = ?", (username,)).fetchone()
            if not exists:
                connection.execute(
                    """
                    INSERT INTO users (username, display_name, role, permissions, password_hash, active, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, 1, ?, ?)
                    """,
                    (
                        username,
                        display_name,
                        role,
                        json.dumps(default_permissions_for_role(role), ensure_ascii=False),
                        password_hash(default_password),
                        now,
                        now,
                    ),
                )
            elif not exists["permissions"]:
                connection.execute(
                    "UPDATE users SET permissions = ?, updated_at = ? WHERE username = ?",
                    (json.dumps(default_permissions_for_role(role), ensure_ascii=False), now, username),
                )
        for row in connection.execute("SELECT username, role FROM users WHERE permissions IS NULL OR permissions = ''").fetchall():
            connection.execute(
                "UPDATE users SET permissions = ?, updated_at = ? WHERE username = ?",
                (
                    json.dumps(default_permissions_for_role(row["role"]), ensure_ascii=False),
                    now,
                    row["username"],
                ),
            )
        connection.execute("DELETE FROM login_sessions WHERE expires_at < ?", (time.time(),))
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS cs_cases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT '접수',
                vendor_name TEXT,
                vendor_email TEXT,
                original_info TEXT,
                original_invoice TEXT,
                product_name TEXT,
                orderer_name TEXT,
                orderer_phone TEXT,
                receiver_name TEXT,
                receiver_phone TEXT,
                receiver_address TEXT,
                cs_type TEXT,
                cs_content TEXT,
                return_invoice TEXT,
                reship_invoice TEXT,
                mail_subject TEXT,
                mail_body TEXT,
                mail_sent_at TEXT,
                source_file TEXT,
                source_sheet TEXT,
                source_row INTEGER,
                occurred_at TEXT,
                completed_at TEXT,
                order_date TEXT,
                ship_date TEXT,
                sales_vendor TEXT,
                purchase_vendor TEXT,
                courier TEXT,
                quantity TEXT
            )
            """
        )
        existing_columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(cs_cases)").fetchall()
        }
        extra_columns = {
            "source_file": "TEXT",
            "source_sheet": "TEXT",
            "source_row": "INTEGER",
            "occurred_at": "TEXT",
            "completed_at": "TEXT",
            "order_date": "TEXT",
            "ship_date": "TEXT",
            "orderer_name": "TEXT",
            "orderer_phone": "TEXT",
            "sales_vendor": "TEXT",
            "purchase_vendor": "TEXT",
            "courier": "TEXT",
            "quantity": "TEXT",
            "cs_type": "TEXT",
        }
        for column, column_type in extra_columns.items():
            if column not in existing_columns:
                connection.execute(f"ALTER TABLE cs_cases ADD COLUMN {column} {column_type}")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_cs_cases_status ON cs_cases(status)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_cs_cases_original_invoice ON cs_cases(original_invoice)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_cs_cases_receiver_phone ON cs_cases(receiver_phone)")
        connection.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_cs_cases_source ON cs_cases(source_file, source_sheet, source_row)")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS management_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                source_file TEXT NOT NULL,
                source_sheet TEXT NOT NULL,
                source_row INTEGER NOT NULL,
                purchase_vendor TEXT,
                sales_vendor TEXT,
                transaction_type TEXT,
                ledger_checked TEXT,
                order_date TEXT,
                ship_date TEXT,
                orderer_name TEXT,
                sender_phone TEXT,
                receiver_name TEXT,
                receiver_phone TEXT,
                product_name TEXT,
                quantity TEXT,
                receiver_address TEXT,
                courier TEXT,
                invoice_number TEXT,
                memo TEXT,
                cs_received_at TEXT
            )
            """
        )
        management_columns = {
            row["name"] for row in connection.execute("PRAGMA table_info(management_records)").fetchall()
        }
        if "cs_received_at" not in management_columns:
            connection.execute("ALTER TABLE management_records ADD COLUMN cs_received_at TEXT")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_management_invoice ON management_records(invoice_number)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_management_receiver_phone ON management_records(receiver_phone)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_management_order_date ON management_records(order_date)")
        connection.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_management_source ON management_records(source_file, source_sheet, source_row)")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS leave_types (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL,
                is_paid INTEGER NOT NULL DEFAULT 1,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS leave_balances (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                leave_type_id INTEGER NOT NULL,
                total_days REAL NOT NULL DEFAULT 0,
                used_days REAL NOT NULL DEFAULT 0,
                remaining_days REAL NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(user_id, leave_type_id)
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS leave_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                leave_type_id INTEGER NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                unit TEXT NOT NULL,
                requested_days REAL NOT NULL,
                reason TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'PENDING',
                rejected_reason TEXT,
                decided_by INTEGER,
                finalized_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS leave_balance_ledger (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                balance_id INTEGER NOT NULL,
                user_id INTEGER NOT NULL,
                actor_id INTEGER,
                delta_days REAL NOT NULL,
                reason TEXT NOT NULL,
                request_id INTEGER,
                created_at TEXT NOT NULL
            )
            """
        )
        timestamp = now_text()
        for code, name in (("annual", "연차"), ("special", "특별휴가"), ("unpaid", "무급휴가"), ("sick", "병가")):
            connection.execute(
                """
                INSERT OR IGNORE INTO leave_types (code, name, is_paid, is_active, created_at, updated_at)
                VALUES (?, ?, ?, 1, ?, ?)
                """,
                (code, name, 0 if code == "unpaid" else 1, timestamp, timestamp),
            )
        connection.execute("CREATE INDEX IF NOT EXISTS idx_leave_requests_user ON leave_requests(user_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_leave_requests_status ON leave_requests(status)")
        connection.commit()
    finally:
        connection.close()


def authenticate_user(username: str, password: str) -> dict[str, str] | None:
    init_db()
    normalized = username.strip()
    if not normalized or not password:
        return None
    connection = connect_db()
    try:
        row = connection.execute(
            """
            SELECT id, username, display_name, role, permissions, password_hash
              FROM users
             WHERE username = ? AND active = 1
            """,
            (normalized,),
        ).fetchone()
        if not row or not verify_password(password, row["password_hash"]):
            return None
        return {
            "id": str(row["id"]),
            "username": row["username"],
            "display_name": row["display_name"],
            "role": row["role"],
            "permissions": normalize_permissions(row["permissions"], row["role"]),
        }
    finally:
        connection.close()


def create_login_session(username: str) -> str:
    init_db()
    token = secrets.token_urlsafe(32)
    now = time.time()
    connection = connect_db()
    try:
        connection.execute(
            """
            INSERT INTO login_sessions (token_hash, username, created_at, expires_at)
            VALUES (?, ?, ?, ?)
            """,
            (token_digest(token), username, now, now + SESSION_SECONDS),
        )
        connection.commit()
        return token
    finally:
        connection.close()


def current_user_from_token(token: str) -> dict[str, str] | None:
    if not token:
        return None
    init_db()
    connection = connect_db()
    try:
        row = connection.execute(
            """
            SELECT users.id, users.username, users.display_name, users.role, users.permissions, login_sessions.expires_at
              FROM login_sessions
              JOIN users ON users.username = login_sessions.username
             WHERE login_sessions.token_hash = ?
               AND users.active = 1
            """,
            (token_digest(token),),
        ).fetchone()
        if not row:
            return None
        if float(row["expires_at"]) < time.time():
            connection.execute("DELETE FROM login_sessions WHERE token_hash = ?", (token_digest(token),))
            connection.commit()
            return None
        return {
            "id": str(row["id"]),
            "username": row["username"],
            "display_name": row["display_name"],
            "role": row["role"],
            "permissions": normalize_permissions(row["permissions"], row["role"]),
        }
    finally:
        connection.close()


def delete_login_session(token: str) -> None:
    if not token:
        return
    init_db()
    connection = connect_db()
    try:
        connection.execute("DELETE FROM login_sessions WHERE token_hash = ?", (token_digest(token),))
        connection.commit()
    finally:
        connection.close()


def list_users() -> list[dict[str, str | int]]:
    init_db()
    connection = connect_db()
    try:
        rows = connection.execute(
            """
            SELECT id, username, display_name, role, permissions, active, created_at, updated_at
              FROM users
             ORDER BY role = 'admin' DESC, username COLLATE NOCASE
            """
        ).fetchall()
        users = []
        for row in rows:
            item = dict(row)
            item["permissions"] = normalize_permissions(item.get("permissions"), str(item.get("role", "user")))
            users.append(item)
        return users
    finally:
        connection.close()


def save_user_account(payload: dict, actor: dict[str, str]) -> int:
    init_db()
    user_id = int(payload.get("id", 0) or 0)
    username = clean_payload_text(payload, "username")
    display_name = clean_payload_text(payload, "display_name")
    role = clean_payload_text(payload, "role") or "user"
    password = str(payload.get("password", "") or "")
    active = 1 if payload.get("active", True) else 0
    permissions = normalize_permissions(payload.get("permissions", []), role)
    if role not in {"admin", "user"}:
        raise ValueError("권한은 관리자 또는 사용자만 선택할 수 있습니다.")
    if not re.fullmatch(r"[A-Za-z0-9_.-]{3,32}", username):
        raise ValueError("아이디는 영문/숫자/._- 조합 3~32자로 입력해주세요.")
    if not display_name:
        display_name = username
    if not user_id and not password:
        raise ValueError("신규 사용자는 비밀번호를 입력해주세요.")
    now = now_text()
    connection = connect_db()
    try:
        if user_id:
            existing = connection.execute("SELECT username FROM users WHERE id = ?", (user_id,)).fetchone()
            if not existing:
                raise ValueError("수정할 사용자를 찾지 못했습니다.")
            if existing["username"] == actor.get("username") and not active:
                raise ValueError("현재 로그인한 관리자 계정은 사용 중지할 수 없습니다.")
            columns = [
                "username = ?",
                "display_name = ?",
                "role = ?",
                "permissions = ?",
                "active = ?",
                "updated_at = ?",
            ]
            values: list[object] = [username, display_name, role, json.dumps(permissions, ensure_ascii=False), active, now]
            if password:
                columns.append("password_hash = ?")
                values.append(password_hash(password))
            values.append(user_id)
            connection.execute(f"UPDATE users SET {', '.join(columns)} WHERE id = ?", values)
            saved_id = user_id
        else:
            cursor = connection.execute(
                """
                INSERT INTO users (username, display_name, role, permissions, password_hash, active, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (username, display_name, role, json.dumps(permissions, ensure_ascii=False), password_hash(password), active, now, now),
            )
            saved_id = int(cursor.lastrowid)
        connection.commit()
        return saved_id
    except sqlite3.IntegrityError as exc:
        raise ValueError("이미 사용 중인 아이디입니다.") from exc
    finally:
        connection.close()


def user_has_permission(user: dict[str, str], permission: str) -> bool:
    return permission in normalize_permissions(user.get("permissions", []), user.get("role", "user"))


RESTORE_ALLOWED_FILES = {
    "config/workhub.db": DB_PATH,
    "config/mail_settings.json": MAIL_SETTINGS_PATH,
    "config/vendor_contacts.json": VENDOR_CONTACTS_PATH,
    "config/secret.key": SECRET_KEY_PATH,
}


def valid_backup_name(name: str) -> bool:
    return bool(re.fullmatch(r"workhub_backup_\d{8}_\d{6}(?:_\d+)?\.zip", name))


def backup_path_from_name(name: str) -> Path:
    name = Path(name).name
    if not valid_backup_name(name):
        raise ValueError("백업 파일명이 올바르지 않습니다.")
    path = (BACKUP_DIR / name).resolve()
    if BACKUP_DIR.resolve() not in path.parents:
        raise ValueError("백업 파일 경로가 올바르지 않습니다.")
    return path


def backup_file_info(path: Path) -> dict[str, str | int]:
    stat = path.stat()
    return {
        "name": path.name,
        "size": stat.st_size,
        "created_at": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
    }


def list_backup_files() -> list[dict[str, str | int]]:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    files = [path for path in BACKUP_DIR.glob("workhub_backup_*.zip") if path.is_file() and valid_backup_name(path.name)]
    files.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    return [backup_file_info(path) for path in files]


def cleanup_backup_retention() -> None:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    cutoff = time.time() - (BACKUP_RETENTION_DAYS * 24 * 60 * 60)
    for path in BACKUP_DIR.glob("workhub_backup_*.zip"):
        if path.is_file() and valid_backup_name(path.name) and path.stat().st_mtime < cutoff:
            path.unlink(missing_ok=True)


def safe_sqlite_copy(target_db: Path) -> None:
    init_db()
    source = connect_db()
    target = sqlite3.connect(target_db)
    try:
        source.backup(target)
        target.commit()
    finally:
        target.close()
        source.close()


def create_workhub_backup(reason: str = "manual") -> dict[str, str | int]:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = BACKUP_DIR / f"workhub_backup_{timestamp}.zip"
    counter = 1
    while output_path.exists():
        output_path = BACKUP_DIR / f"workhub_backup_{timestamp}_{counter}.zip"
        counter += 1
    with tempfile.TemporaryDirectory(prefix="workhub_backup_", dir=BACKUP_DIR) as temp_dir:
        temp_db = Path(temp_dir) / "workhub.db"
        safe_sqlite_copy(temp_db)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(temp_db, "config/workhub.db")
            for path in (MAIL_SETTINGS_PATH, VENDOR_CONTACTS_PATH, SECRET_KEY_PATH):
                if path.exists() and path.is_file():
                    archive.write(path, f"config/{path.name}")
            archive.writestr(
                "manifest.json",
                json.dumps(
                    {
                        "created_at": now_text(),
                        "reason": reason,
                        "data_dir": str(RUNTIME_ROOT),
                        "backup_retention_days": BACKUP_RETENTION_DAYS,
                        "included": [
                            "config/workhub.db",
                            "config/mail_settings.json",
                            "config/vendor_contacts.json",
                            "config/secret.key",
                        ],
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
            )
    cleanup_backup_retention()
    return backup_file_info(output_path)


def validate_restored_db(path: Path) -> None:
    connection = sqlite3.connect(path)
    try:
        result = connection.execute("PRAGMA quick_check").fetchone()
        if not result or str(result[0]).lower() != "ok":
            raise ValueError("백업 DB 파일 검증에 실패했습니다.")
    finally:
        connection.close()


def restore_workhub_backup(source_zip: Path) -> dict[str, object]:
    if not source_zip.exists() or not zipfile.is_zipfile(source_zip):
        raise ValueError("올바른 백업 zip 파일이 아닙니다.")

    pre_restore = create_workhub_backup("pre-restore")
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="workhub_restore_", dir=BACKUP_DIR) as temp_dir:
        temp_root = Path(temp_dir)
        extracted: dict[str, Path] = {}
        with zipfile.ZipFile(source_zip) as archive:
            names = {item.filename.replace("\\", "/").lstrip("/") for item in archive.infolist()}
            if "config/workhub.db" not in names:
                raise ValueError("백업 파일 안에 config/workhub.db가 없습니다.")
            for member in archive.infolist():
                member_name = member.filename.replace("\\", "/").lstrip("/")
                if member_name not in RESTORE_ALLOWED_FILES:
                    continue
                output = temp_root / member_name
                output.parent.mkdir(parents=True, exist_ok=True)
                output.write_bytes(archive.read(member))
                extracted[member_name] = output

        validate_restored_db(extracted["config/workhub.db"])

        for backup_name, target in RESTORE_ALLOWED_FILES.items():
            source = extracted.get(backup_name)
            if source:
                target.parent.mkdir(parents=True, exist_ok=True)
                source.replace(target)
            elif target.exists():
                target.unlink()

    return {
        "message": "백업 데이터 복원이 완료되었습니다.",
        "restored_from": source_zip.name,
        "pre_restore_backup": pre_restore,
    }


def has_backup_for_date(date_token: str) -> bool:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    return any(BACKUP_DIR.glob(f"workhub_backup_{date_token}_*.zip"))


def backup_scheduler_loop() -> None:
    while True:
        try:
            now = datetime.now()
            if now.hour >= AUTO_BACKUP_HOUR and not has_backup_for_date(now.strftime("%Y%m%d")):
                created = create_workhub_backup("auto")
                print(f"자동 백업 완료: {created['name']}")
        except Exception as exc:  # noqa: BLE001
            print(f"자동 백업 실패: {exc}")
        time.sleep(5 * 60)


def start_backup_scheduler() -> None:
    global _BACKUP_SCHEDULER_STARTED
    if _BACKUP_SCHEDULER_STARTED:
        return
    _BACKUP_SCHEDULER_STARTED = True
    thread = threading.Thread(target=backup_scheduler_loop, name="workhub-backup-scheduler", daemon=True)
    thread.start()


def git_source_dir() -> Path | None:
    candidates: list[Path] = []
    env_source = os.environ.get("WORKHUB_SOURCE_DIR")
    if env_source:
        candidates.append(Path(env_source))
    candidates.extend([ROOT, RUNTIME_ROOT, Path.cwd()])
    for candidate in candidates:
        try:
            path = candidate.resolve()
        except OSError:
            continue
        if (path / ".git").exists():
            return path
    return None


def run_git(args: list[str], source_dir: Path, timeout: int = 30) -> tuple[int, str, str]:
    try:
        completed = subprocess.run(
            ["git", "-C", str(source_dir), *args],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
        )
        return completed.returncode, completed.stdout.strip(), completed.stderr.strip()
    except FileNotFoundError:
        return 127, "", "git 명령을 찾지 못했습니다."
    except subprocess.TimeoutExpired:
        return 124, "", "git 명령 시간이 초과되었습니다."


def git_output(args: list[str], source_dir: Path, default: str = "") -> str:
    code, stdout, _ = run_git(args, source_dir)
    return stdout if code == 0 else default


def system_update_status(fetch: bool = False) -> dict[str, object]:
    source_dir = git_source_dir()
    if not source_dir:
        return {
            "available": False,
            "source_dir": "",
            "branch": "",
            "current_commit": "",
            "current_short": "",
            "upstream_commit": "",
            "remote_url": "",
            "ahead": 0,
            "behind": 0,
            "dirty": False,
            "message": "현재 실행 위치가 Git 저장소가 아닙니다. 나스 배포 시 GitHub 저장소 폴더에서 실행하면 업데이트 기능을 사용할 수 있습니다.",
        }
    if fetch:
        run_git(["fetch", "--all", "--prune"], source_dir, timeout=60)
    current_commit = git_output(["rev-parse", "HEAD"], source_dir)
    current_short = git_output(["rev-parse", "--short", "HEAD"], source_dir)
    branch = git_output(["rev-parse", "--abbrev-ref", "HEAD"], source_dir)
    upstream = git_output(["rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}"], source_dir)
    upstream_commit = git_output(["rev-parse", "@{u}"], source_dir) if upstream else ""
    ahead = 0
    behind = 0
    if upstream:
        counts = git_output(["rev-list", "--left-right", "--count", "HEAD...@{u}"], source_dir)
        parts = counts.split()
        if len(parts) == 2 and all(part.isdigit() for part in parts):
            ahead = int(parts[0])
            behind = int(parts[1])
    dirty = bool(git_output(["status", "--porcelain"], source_dir))
    message = ""
    if not upstream:
        message = "원격 추적 브랜치가 설정되어 있지 않습니다."
    elif dirty:
        message = "저장되지 않은 코드 변경이 있어 자동 업데이트를 적용하지 않습니다."
    elif behind > 0:
        message = f"GitHub에 적용 가능한 업데이트 {behind}개가 있습니다."
    else:
        message = "최신 상태입니다."
    return {
        "available": bool(upstream),
        "source_dir": str(source_dir),
        "branch": branch,
        "current_commit": current_commit,
        "current_short": current_short,
        "upstream": upstream,
        "upstream_commit": upstream_commit,
        "remote_url": git_output(["config", "--get", "remote.origin.url"], source_dir),
        "ahead": ahead,
        "behind": behind,
        "dirty": dirty,
        "message": message,
    }


def list_system_update_history(limit: int = 30) -> list[dict[str, str | int]]:
    init_db()
    connection = connect_db()
    try:
        rows = connection.execute(
            """
            SELECT id, action, before_commit, after_commit, status, message, backup_name, created_at
              FROM system_update_history
             ORDER BY id DESC
             LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        connection.close()


def record_system_update(action: str, before_commit: str, after_commit: str, status: str, message: str, backup_name: str = "") -> None:
    init_db()
    connection = connect_db()
    try:
        connection.execute(
            """
            INSERT INTO system_update_history
                (action, before_commit, after_commit, status, message, backup_name, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (action, before_commit, after_commit, status, message, backup_name, now_text()),
        )
        connection.commit()
    finally:
        connection.close()


def system_update_payload(fetch: bool = False) -> dict[str, object]:
    return {
        "status": system_update_status(fetch=fetch),
        "history": list_system_update_history(),
    }


def apply_system_update() -> dict[str, object]:
    if not _SYSTEM_UPDATE_LOCK.acquire(blocking=False):
        raise ValueError("이미 업데이트가 진행 중입니다.")
    backup_name = ""
    before_commit = ""
    after_commit = ""
    try:
        status = system_update_status(fetch=True)
        before_commit = str(status.get("current_short") or status.get("current_commit") or "")
        if not status.get("available"):
            raise ValueError(str(status.get("message") or "업데이트 기능을 사용할 수 없습니다."))
        if status.get("dirty"):
            raise ValueError("저장되지 않은 코드 변경이 있어 업데이트를 중단했습니다.")
        if int(status.get("behind") or 0) <= 0:
            record_system_update("apply", before_commit, before_commit, "skip", "이미 최신 상태입니다.", "")
            payload = system_update_payload(fetch=False)
            payload["message"] = "이미 최신 상태입니다."
            return payload

        backup = create_workhub_backup("pre-update")
        backup_name = str(backup["name"])
        source_dir = Path(str(status["source_dir"]))
        code, stdout, stderr = run_git(["pull", "--ff-only"], source_dir, timeout=120)
        if code != 0:
            message = stderr or stdout or "git pull 실패"
            record_system_update("apply", before_commit, before_commit, "fail", message, backup_name)
            raise ValueError(message)
        new_status = system_update_status(fetch=False)
        after_commit = str(new_status.get("current_short") or new_status.get("current_commit") or "")
        message = stdout or "업데이트 적용 완료"
        record_system_update("apply", before_commit, after_commit, "success", message, backup_name)
        payload = system_update_payload(fetch=False)
        payload["message"] = "업데이트 적용이 완료되었습니다."
        payload["backup_name"] = backup_name
        payload["restart_required"] = True
        return payload
    finally:
        _SYSTEM_UPDATE_LOCK.release()


def get_leave_type_id(code: str = "annual") -> int:
    connection = connect_db()
    try:
        row = connection.execute("SELECT id FROM leave_types WHERE code = ?", (code,)).fetchone()
        if not row:
            raise ValueError("연차 유형을 찾지 못했습니다.")
        return int(row["id"])
    finally:
        connection.close()


def ensure_leave_balance(connection: sqlite3.Connection, user_id: int, leave_type_id: int) -> sqlite3.Row:
    now = now_text()
    connection.execute(
        """
        INSERT OR IGNORE INTO leave_balances
            (user_id, leave_type_id, total_days, used_days, remaining_days, created_at, updated_at)
        VALUES (?, ?, 0, 0, 0, ?, ?)
        """,
        (user_id, leave_type_id, now, now),
    )
    row = connection.execute(
        "SELECT * FROM leave_balances WHERE user_id = ? AND leave_type_id = ?",
        (user_id, leave_type_id),
    ).fetchone()
    if not row:
        raise ValueError("연차 잔여 정보를 만들지 못했습니다.")
    return row


def list_leave_types() -> list[dict[str, str | int]]:
    connection = connect_db()
    try:
        rows = connection.execute(
            "SELECT id, code, name, is_paid, is_active FROM leave_types WHERE is_active = 1 ORDER BY id"
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        connection.close()


def list_active_users_for_leave() -> list[dict[str, str | int]]:
    connection = connect_db()
    try:
        rows = connection.execute(
            """
            SELECT id, username, display_name
              FROM users
             WHERE active = 1
             ORDER BY display_name COLLATE NOCASE, username COLLATE NOCASE
            """
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        connection.close()


def leave_balance_rows(connection: sqlite3.Connection, user_id: int) -> list[dict[str, str | float]]:
    annual_type_id = get_leave_type_id("annual")
    ensure_leave_balance(connection, user_id, annual_type_id)
    rows = connection.execute(
        """
        SELECT leave_types.code, leave_types.name, leave_balances.total_days,
               leave_balances.used_days, leave_balances.remaining_days
          FROM leave_balances
          JOIN leave_types ON leave_types.id = leave_balances.leave_type_id
         WHERE leave_balances.user_id = ?
         ORDER BY leave_types.id
        """,
        (user_id,),
    ).fetchall()
    return [
        {
            "code": row["code"],
            "name": row["name"],
            "total_days": round(float(row["total_days"] or 0), 2),
            "used_days": round(float(row["used_days"] or 0), 2),
            "remaining_days": round(float(row["remaining_days"] or 0), 2),
        }
        for row in rows
    ]


def list_leave_requests_for_user(connection: sqlite3.Connection, user_id: int) -> list[dict[str, str | float | int]]:
    rows = connection.execute(
        """
        SELECT leave_requests.*, leave_types.name AS leave_type_name, users.display_name
          FROM leave_requests
          JOIN leave_types ON leave_types.id = leave_requests.leave_type_id
          JOIN users ON users.id = leave_requests.user_id
         WHERE leave_requests.user_id = ?
         ORDER BY leave_requests.start_date DESC, leave_requests.id DESC
        """,
        (user_id,),
    ).fetchall()
    return [leave_request_dict(row) for row in rows]


def leave_request_dict(row: sqlite3.Row) -> dict[str, str | float | int]:
    return {
        "id": row["id"],
        "user_id": row["user_id"],
        "requester": row["display_name"] if "display_name" in row.keys() else "",
        "leave_type_name": row["leave_type_name"] if "leave_type_name" in row.keys() else "연차",
        "start_date": row["start_date"],
        "end_date": row["end_date"],
        "unit": row["unit"],
        "unit_label": "반차" if row["unit"] == "HALF_DAY" else "연차",
        "requested_days": round(float(row["requested_days"] or 0), 2),
        "reason": row["reason"],
        "status": row["status"],
        "status_label": leave_status_label(row["status"]),
        "rejected_reason": row["rejected_reason"] or "",
        "created_at": row["created_at"],
    }


def list_pending_leave_requests(connection: sqlite3.Connection) -> list[dict[str, str | float | int]]:
    rows = connection.execute(
        """
        SELECT leave_requests.*, leave_types.name AS leave_type_name, users.display_name
          FROM leave_requests
          JOIN leave_types ON leave_types.id = leave_requests.leave_type_id
          JOIN users ON users.id = leave_requests.user_id
         WHERE leave_requests.status = 'PENDING'
         ORDER BY leave_requests.created_at ASC
        """
    ).fetchall()
    return [leave_request_dict(row) for row in rows]


def list_leave_admin_balances(connection: sqlite3.Connection) -> list[dict[str, str | float | int]]:
    annual_type_id = get_leave_type_id("annual")
    users = list_active_users_for_leave()
    rows = []
    for user in users:
        balance = ensure_leave_balance(connection, int(user["id"]), annual_type_id)
        rows.append({
            "user_id": user["id"],
            "username": user["username"],
            "display_name": user["display_name"],
            "total_days": round(float(balance["total_days"] or 0), 2),
            "used_days": round(float(balance["used_days"] or 0), 2),
            "remaining_days": round(float(balance["remaining_days"] or 0), 2),
        })
    return rows


def leave_payload(user: dict[str, str]) -> dict:
    user_id = int(user["id"])
    connection = connect_db()
    try:
        balances = leave_balance_rows(connection, user_id)
        annual = next((row for row in balances if row["code"] == "annual"), balances[0] if balances else {})
        payload = {
            "summary": {
                "total_days": annual.get("total_days", 0),
                "used_days": annual.get("used_days", 0),
                "remaining_days": annual.get("remaining_days", 0),
            },
            "balances": balances,
            "requests": list_leave_requests_for_user(connection, user_id),
            "leave_types": list_leave_types(),
            "can_approve": user_has_permission(user, "leave_approve") or user_has_permission(user, "leave_manage"),
            "can_manage": user_has_permission(user, "leave_manage"),
            "pending_requests": [],
            "users": [],
            "admin_balances": [],
        }
        if payload["can_approve"]:
            payload["pending_requests"] = list_pending_leave_requests(connection)
        if payload["can_manage"]:
            payload["users"] = list_active_users_for_leave()
            payload["admin_balances"] = list_leave_admin_balances(connection)
        connection.commit()
        return payload
    finally:
        connection.close()


def create_leave_request(user: dict[str, str], payload: dict) -> int:
    init_db()
    user_id = int(user["id"])
    leave_type_id = int(payload.get("leave_type_id") or get_leave_type_id("annual"))
    start = parse_iso_date(payload.get("start_date"))
    end = parse_iso_date(payload.get("end_date"))
    unit = clean_payload_text(payload, "unit") or "FULL_DAY"
    if unit not in {"FULL_DAY", "HALF_DAY"}:
        raise ValueError("연차 단위가 올바르지 않습니다.")
    requested_days = calculate_leave_days(start, end, unit)
    reason = clean_payload_text(payload, "reason")
    if not reason:
        raise ValueError("연차 사유를 입력해주세요.")
    now = now_text()
    connection = connect_db()
    try:
        balance = ensure_leave_balance(connection, user_id, leave_type_id)
        if float(balance["remaining_days"] or 0) < requested_days:
            raise ValueError("남은 연차보다 신청 일수가 많습니다.")
        cursor = connection.execute(
            """
            INSERT INTO leave_requests
                (user_id, leave_type_id, start_date, end_date, unit, requested_days, reason,
                 status, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, 'PENDING', ?, ?)
            """,
            (user_id, leave_type_id, start.isoformat(), end.isoformat(), unit, requested_days, reason, now, now),
        )
        connection.commit()
        return int(cursor.lastrowid)
    finally:
        connection.close()


def decide_leave_request(request_id: int, actor: dict[str, str], decision: str, comment: str = "") -> None:
    init_db()
    if decision not in {"approve", "reject"}:
        raise ValueError("처리 값이 올바르지 않습니다.")
    connection = connect_db()
    try:
        row = connection.execute("SELECT * FROM leave_requests WHERE id = ?", (request_id,)).fetchone()
        if not row:
            raise ValueError("연차 신청건을 찾지 못했습니다.")
        if row["status"] != "PENDING":
            raise ValueError("이미 처리된 연차 신청입니다.")
        now = now_text()
        if decision == "reject":
            connection.execute(
                """
                UPDATE leave_requests
                   SET status = 'REJECTED', rejected_reason = ?, decided_by = ?,
                       finalized_at = ?, updated_at = ?
                 WHERE id = ?
                """,
                (comment or "반려", int(actor["id"]), now, now, request_id),
            )
        else:
            balance = ensure_leave_balance(connection, int(row["user_id"]), int(row["leave_type_id"]))
            requested_days = float(row["requested_days"] or 0)
            if float(balance["remaining_days"] or 0) < requested_days:
                raise ValueError("잔여 연차가 부족해 승인할 수 없습니다.")
            used = round(float(balance["used_days"] or 0) + requested_days, 2)
            remaining = round(float(balance["remaining_days"] or 0) - requested_days, 2)
            connection.execute(
                "UPDATE leave_balances SET used_days = ?, remaining_days = ?, updated_at = ? WHERE id = ?",
                (used, remaining, now, balance["id"]),
            )
            connection.execute(
                """
                INSERT INTO leave_balance_ledger
                    (balance_id, user_id, actor_id, delta_days, reason, request_id, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (balance["id"], row["user_id"], int(actor["id"]), -requested_days, "연차 승인 차감", request_id, now),
            )
            connection.execute(
                """
                UPDATE leave_requests
                   SET status = 'APPROVED', decided_by = ?, finalized_at = ?, updated_at = ?
                 WHERE id = ?
                """,
                (int(actor["id"]), now, now, request_id),
            )
        connection.commit()
    finally:
        connection.close()


def set_leave_balance(payload: dict, actor: dict[str, str]) -> None:
    init_db()
    user_id = int(payload.get("user_id") or 0)
    if not user_id:
        raise ValueError("직원을 선택해주세요.")
    total = clean_leave_days(payload.get("total_days"), "시작 연차")
    used = clean_leave_days(payload.get("used_days"), "기존 사용 연차")
    if used > total:
        raise ValueError("기존 사용 연차는 시작 연차보다 클 수 없습니다.")
    remaining = round(total - used, 2)
    leave_type_id = get_leave_type_id("annual")
    now = now_text()
    connection = connect_db()
    try:
        balance = ensure_leave_balance(connection, user_id, leave_type_id)
        connection.execute(
            """
            UPDATE leave_balances
               SET total_days = ?, used_days = ?, remaining_days = ?, updated_at = ?
             WHERE id = ?
            """,
            (total, used, remaining, now, balance["id"]),
        )
        connection.execute(
            """
            INSERT INTO leave_balance_ledger
                (balance_id, user_id, actor_id, delta_days, reason, request_id, created_at)
            VALUES (?, ?, ?, ?, ?, NULL, ?)
            """,
            (balance["id"], user_id, int(actor["id"]), 0, f"초기 연차 설정: 시작 {total}일, 기존 사용 {used}일", now),
        )
        connection.commit()
    finally:
        connection.close()


def parse_historical_leave_lines(text: str) -> list[tuple[date, str, float]]:
    entries: list[tuple[date, str, float]] = []
    seen: set[tuple[str, float]] = set()
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        match = re.match(r"^(\d{4}-\d{2}-\d{2})(?:[\s,]+(반차|0\.5|1|1\.0|연차|일))?$", line)
        if not match:
            raise ValueError(f"사용 일자 형식이 올바르지 않습니다: {line}")
        used_date = parse_iso_date(match.group(1))
        unit_text = match.group(2) or "연차"
        unit = "HALF_DAY" if unit_text in {"반차", "0.5"} else "FULL_DAY"
        days = 0.5 if unit == "HALF_DAY" else 1.0
        key = (used_date.isoformat(), days)
        if key in seen:
            raise ValueError(f"같은 사용 일자가 중복 입력됐습니다: {used_date.isoformat()}")
        seen.add(key)
        entries.append((used_date, unit, days))
    if not entries:
        raise ValueError("사용 일자를 1개 이상 입력해주세요.")
    return entries


def add_historical_leave_usage(payload: dict, actor: dict[str, str]) -> int:
    init_db()
    user_id = int(payload.get("user_id") or 0)
    if not user_id:
        raise ValueError("직원을 선택해주세요.")
    entries = parse_historical_leave_lines(str(payload.get("usage_dates", "") or ""))
    note = clean_payload_text(payload, "note")
    leave_type_id = get_leave_type_id("annual")
    now = now_text()
    connection = connect_db()
    try:
        balance = ensure_leave_balance(connection, user_id, leave_type_id)
        total_days = round(sum(entry[2] for entry in entries), 2)
        if float(balance["remaining_days"] or 0) < total_days:
            raise ValueError("잔여 연차보다 기존 사용 연차가 많습니다.")
        for used_date, unit, days in entries:
            reason = f"기존 사용 연차 등록{': ' + note if note else ''}"
            cursor = connection.execute(
                """
                INSERT INTO leave_requests
                    (user_id, leave_type_id, start_date, end_date, unit, requested_days, reason,
                     status, decided_by, finalized_at, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'APPROVED', ?, ?, ?, ?)
                """,
                (user_id, leave_type_id, used_date.isoformat(), used_date.isoformat(), unit, days, reason, int(actor["id"]), now, now, now),
            )
            connection.execute(
                """
                INSERT INTO leave_balance_ledger
                    (balance_id, user_id, actor_id, delta_days, reason, request_id, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (balance["id"], user_id, int(actor["id"]), -days, reason, int(cursor.lastrowid), now),
            )
        used = round(float(balance["used_days"] or 0) + total_days, 2)
        remaining = round(float(balance["remaining_days"] or 0) - total_days, 2)
        connection.execute(
            "UPDATE leave_balances SET used_days = ?, remaining_days = ?, updated_at = ? WHERE id = ?",
            (used, remaining, now, balance["id"]),
        )
        connection.commit()
        return len(entries)
    finally:
        connection.close()


def extract_invoice_number(value: str) -> str:
    numbers = re.findall(r"\d{6,}", value.replace("-", "").replace(" ", ""))
    return numbers[-1] if numbers else ""


def normalize_compact(value: str) -> str:
    return re.sub(r"[\s/_-]+", "", str(value or "").strip())


def status_date_text(*values: object) -> str:
    for value in values:
        text = clean_cell(value)
        if not text:
            continue
        month_day = re.search(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일?", text)
        if month_day:
            return f"{int(month_day.group(1))}/{int(month_day.group(2))}"
        slash_day = re.search(r"(?<!\d)(\d{1,2})[./-](\d{1,2})(?!\d)", text)
        if slash_day:
            return f"{int(slash_day.group(1))}/{int(slash_day.group(2))}"
        iso_day = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", text)
        if iso_day:
            return f"{int(iso_day.group(2))}/{int(iso_day.group(3))}"
    now = datetime.now()
    return f"{now.month}/{now.day}"


def normalize_cs_type_value(cs_type: str, cs_content: str, status: str = "") -> str:
    existing = clean_cell(cs_type)
    if existing in {"변심반품", "변신반품", "불량반품", "불량교환", "불량재출고(미회수)", "오출고(오배송)"}:
        return "변심반품" if existing == "변신반품" else existing

    text = normalize_compact(f"{existing} {cs_content} {status}")
    if not text:
        return ""

    if any(keyword in text for keyword in ["오배송", "오출고", "오발주", "착오", "중복발주", "중복출고", "이중발주", "이중출고", "주소오류", "주소오기재"]):
        return "오출고(오배송)"
    if any(keyword in text for keyword in ["출고취소", "출고전취소", "출소취소"]):
        return "오출고(오배송)"
    if any(keyword in text for keyword in ["변심", "단순변심"]):
        return "변심반품"
    if any(keyword in text for keyword in ["맞교환", "맞효관", "불량교환", "제품불량교환", "교환요청", "교환원", "교환진행", "교환"]):
        return "불량교환"
    if any(keyword in text for keyword in ["상품누락", "누락재발송", "누락재배송", "부분재발송", "부분재출고", "부족분재발송", "일부재발송", "일부분재배송", "미수령재배송", "재출고", "대체출고", "재발송요청", "배출고", "누락배송", "부분누락", "부분발송", "부분미발송", "미발송", "추가배송", "추가발송", "부품구매", "부품출고", "추가구매", "제품미수령", "화물추적", "배송추적", "배송확인", "운송장회신", "충전기분실", "재발송", "발송완료"]):
        return "불량재출고(미회수)"
    if any(keyword in text for keyword in ["불량반품", "제품불량반품", "부분반품", "반품회수", "불량", "파손배송", "파손불량", "환불", "회수완료", "회수요청", "코팅벗겨짐", "입고사유확인"]):
        return "불량반품"
    if "반품" in text:
        return "변심반품"
    return ""


def normalize_progress_status(status: str, completed_at: str = "", occurred_at: str = "") -> str:
    raw = clean_cell(status)
    text = normalize_compact(raw)
    date_text = status_date_text(raw, completed_at, occurred_at)

    has_return_complete = any(keyword in text for keyword in ["회수완료", "입고완료", "반품완료", "수거완료", "반송완료"])
    has_return_request = any(keyword in text for keyword in ["회수요청", "회수지시", "회수등록", "회수접수", "업체처리요청", "자동수거", "직접수거", "업체측회수", "회수"])
    has_reship_complete = any(keyword in text for keyword in ["재발송완료", "재출고완료", "발송완료", "출고완료", "재발송", "재출고", "배출고"])
    has_all_done = any(keyword in text for keyword in ["전체처리완료", "처리완료", "업체처리완료", "정산반영완료", "사고처리완료", "확인완료", "수령확인완료", "화물추적완료"])

    if has_all_done or (has_reship_complete and has_return_complete):
        return "전체 처리완료"
    if has_reship_complete and has_return_request:
        return f"재발송 완료({date_text})/회수지시({date_text})"
    if has_return_complete:
        return f"회수 완료({date_text})"
    if has_reship_complete or any(keyword in text for keyword in ["직접보내주심", "직접보냄", "업체에서직접보냄"]):
        return f"재발송 완료({date_text})"
    if has_return_request or any(keyword in text for keyword in ["반품", "회수철회", "반품철회"]):
        return f"회수지시({date_text})"
    return f"회수지시({date_text})"


def cs_case_from_payload(payload: dict, status: str = "접수", mail_sent: bool = False) -> dict[str, str]:
    original_info = clean_payload_text(payload, "cs_origin")
    return {
        "status": status,
        "vendor_name": clean_payload_text(payload, "vendor_name"),
        "vendor_email": clean_payload_text(payload, "recipient_email"),
        "original_info": original_info,
        "original_invoice": clean_payload_text(payload, "original_invoice") or extract_invoice_number(original_info),
        "product_name": clean_payload_text(payload, "cs_product"),
        "orderer_name": clean_payload_text(payload, "orderer_name"),
        "orderer_phone": clean_payload_text(payload, "orderer_phone"),
        "receiver_name": clean_payload_text(payload, "cs_receiver"),
        "receiver_phone": clean_payload_text(payload, "cs_phone"),
        "receiver_address": clean_payload_text(payload, "cs_address"),
        "cs_type": clean_payload_text(payload, "cs_type"),
        "cs_content": clean_payload_text(payload, "cs_content"),
        "return_invoice": clean_payload_text(payload, "return_invoice"),
        "reship_invoice": clean_payload_text(payload, "reship_invoice"),
        "mail_subject": clean_payload_text(payload, "subject"),
        "mail_body": clean_payload_text(payload, "body"),
        "mail_sent_at": now_text() if mail_sent else clean_payload_text(payload, "mail_sent_at"),
    }


def save_cs_case(payload: dict, status: str = "접수", mail_sent: bool = False) -> int:
    init_db()
    case = cs_case_from_payload(payload, status=status, mail_sent=mail_sent)
    timestamp = now_text()
    columns = [
        "created_at",
        "updated_at",
        "status",
        "vendor_name",
        "vendor_email",
        "original_info",
        "original_invoice",
        "product_name",
        "orderer_name",
        "orderer_phone",
        "receiver_name",
        "receiver_phone",
        "receiver_address",
        "cs_type",
        "cs_content",
        "return_invoice",
        "reship_invoice",
        "mail_subject",
        "mail_body",
        "mail_sent_at",
    ]
    values = {
        "created_at": timestamp,
        "updated_at": timestamp,
        **case,
    }
    placeholders = ", ".join("?" for _ in columns)
    connection = connect_db()
    try:
        cursor = connection.execute(
            f"INSERT INTO cs_cases ({', '.join(columns)}) VALUES ({placeholders})",
            [values[column] for column in columns],
        )
        connection.commit()
        return int(cursor.lastrowid)
    finally:
        connection.close()


def date_period_condition(fields: list[str], year: str = "", month: str = "") -> tuple[str, list[str]]:
    year = clean_cell(year)
    month = clean_cell(month).zfill(2) if clean_cell(month) else ""
    if not re.fullmatch(r"\d{4}", year):
        year = ""
    if not re.fullmatch(r"\d{2}", month) or not (1 <= int(month) <= 12):
        month = ""
    if not year and not month:
        return "", []
    if year and month:
        pattern = f"{year}-{month}%"
    elif year:
        pattern = f"{year}%"
    else:
        pattern = f"%-{month}-%"
    return "(" + " OR ".join(f"{field} LIKE ?" for field in fields) + ")", [pattern] * len(fields)


def list_cs_cases(query: str = "", status: str = "", limit: int = 20, year: str = "", month: str = "") -> list[dict[str, str | int]]:
    init_db()
    query = query.strip()
    status = status.strip()
    params: list[object] = []
    conditions: list[str] = []
    if query:
        pattern = f"%{query}%"
        conditions.append(
            """
            (
                vendor_name LIKE ?
                OR vendor_email LIKE ?
                OR sales_vendor LIKE ?
                OR purchase_vendor LIKE ?
                OR original_info LIKE ?
                OR original_invoice LIKE ?
                OR product_name LIKE ?
                OR orderer_name LIKE ?
                OR orderer_phone LIKE ?
                OR receiver_name LIKE ?
                OR receiver_phone LIKE ?
                OR receiver_address LIKE ?
                OR cs_type LIKE ?
                OR cs_content LIKE ?
                OR return_invoice LIKE ?
                OR reship_invoice LIKE ?
            )
            """
        )
        params = [pattern] * 16
    if status:
        conditions.append("status = ?")
        params.append(status)
    period_condition, period_params = date_period_condition(
        ["occurred_at", "order_date", "ship_date", "created_at"],
        year,
        month,
    )
    if period_condition:
        conditions.append(period_condition)
        params.extend(period_params)
    where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
    params.append(limit)
    connection = connect_db()
    try:
        rows = connection.execute(
            f"""
            SELECT id, created_at, updated_at, status, vendor_name, vendor_email,
                   original_info, original_invoice, product_name, orderer_name, orderer_phone, receiver_name,
                   receiver_phone, receiver_address, cs_type, cs_content, return_invoice,
                   reship_invoice, mail_subject, mail_sent_at, occurred_at, completed_at, order_date,
                   ship_date, sales_vendor, purchase_vendor, courier, quantity
              FROM cs_cases
              {where}
             ORDER BY id DESC
             LIMIT ?
            """,
            params,
        ).fetchall()
    finally:
        connection.close()
    return [dict(row) for row in rows]


def update_cs_case(case_id: int, payload: dict) -> None:
    status = clean_payload_text(payload, "status")
    cs_type = clean_payload_text(payload, "cs_type")
    return_invoice = clean_payload_text(payload, "return_invoice")
    reship_invoice = clean_payload_text(payload, "reship_invoice")

    init_db()
    connection = connect_db()
    try:
        cursor = connection.execute(
            """
            UPDATE cs_cases
               SET status = COALESCE(NULLIF(?, ''), status),
                   cs_type = ?,
                   return_invoice = ?,
                   reship_invoice = ?,
                   updated_at = ?
             WHERE id = ?
            """,
            [status, cs_type, return_invoice, reship_invoice, now_text(), case_id],
        )
        connection.commit()
        if cursor.rowcount == 0:
            raise ValueError("수정할 CS건을 찾지 못했습니다.")
    finally:
        connection.close()


def delete_cs_cases(case_ids: list[int]) -> int:
    init_db()
    if not case_ids:
        return 0
    placeholders = ", ".join("?" for _ in case_ids)
    connection = connect_db()
    try:
        cursor = connection.execute(f"DELETE FROM cs_cases WHERE id IN ({placeholders})", case_ids)
        connection.commit()
        return int(cursor.rowcount or 0)
    finally:
        connection.close()


def management_query_conditions(query: str = "", year: str = "", month: str = "") -> tuple[str, list[object]]:
    query = query.strip()
    params: list[object] = []
    conditions: list[str] = []
    if query:
        pattern = f"%{query}%"
        conditions.append(
            """
            (
                purchase_vendor LIKE ?
                OR sales_vendor LIKE ?
                OR transaction_type LIKE ?
                OR ledger_checked LIKE ?
                OR order_date LIKE ?
                OR ship_date LIKE ?
                OR orderer_name LIKE ?
                OR sender_phone LIKE ?
                OR receiver_name LIKE ?
                OR receiver_phone LIKE ?
                OR product_name LIKE ?
                OR receiver_address LIKE ?
                OR courier LIKE ?
                OR invoice_number LIKE ?
                OR memo LIKE ?
            )
            """
        )
        params = [pattern] * 15
    period_condition, period_params = date_period_condition(["order_date", "ship_date"], year, month)
    if period_condition:
        conditions.append(period_condition)
        params.extend(period_params)
    where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
    return where, params


def list_management_records(query: str = "", limit: int | None = 300, year: str = "", month: str = "") -> list[dict[str, str | int]]:
    init_db()
    where, params = management_query_conditions(query, year, month)
    limit_sql = ""
    if limit is not None:
        limit_sql = "LIMIT ?"
        params.append(limit)
    connection = connect_db()
    try:
        rows = connection.execute(
            f"""
            SELECT id, created_at, source_file, source_sheet, source_row, purchase_vendor,
                   sales_vendor, transaction_type, ledger_checked, order_date, ship_date,
                   orderer_name, sender_phone, receiver_name, receiver_phone, product_name,
                   quantity, receiver_address, courier, invoice_number, memo, cs_received_at
              FROM management_records
              {where}
             ORDER BY order_date DESC, id DESC
             {limit_sql}
            """,
            params,
        ).fetchall()
    finally:
        connection.close()
    return [dict(row) for row in rows]


def list_management_records_by_ids(record_ids: list[int]) -> list[dict[str, str | int]]:
    init_db()
    if not record_ids:
        return []
    placeholders = ", ".join("?" for _ in record_ids)
    order_case = " ".join(f"WHEN {record_id} THEN {index}" for index, record_id in enumerate(record_ids))
    connection = connect_db()
    try:
        rows = connection.execute(
            f"""
            SELECT id, created_at, source_file, source_sheet, source_row, purchase_vendor,
                   sales_vendor, transaction_type, ledger_checked, order_date, ship_date,
                   orderer_name, sender_phone, receiver_name, receiver_phone, product_name,
                   quantity, receiver_address, courier, invoice_number, memo, cs_received_at
              FROM management_records
             WHERE id IN ({placeholders})
             ORDER BY CASE id {order_case} END
            """,
            record_ids,
        ).fetchall()
    finally:
        connection.close()
    return [dict(row) for row in rows]


def list_management_periods() -> list[dict[str, str | int]]:
    init_db()
    connection = connect_db()
    try:
        rows = connection.execute(
            """
            SELECT order_date, ship_date
              FROM management_records
             WHERE COALESCE(order_date, '') <> '' OR COALESCE(ship_date, '') <> ''
            """
        ).fetchall()
    finally:
        connection.close()

    counts: dict[tuple[str, str], int] = {}
    for row in rows:
        year_month = extract_year_month(row["order_date"], row["ship_date"])
        if not year_month:
            continue
        counts[year_month] = counts.get(year_month, 0) + 1
    return [
        {"year": year, "month": month, "count": count}
        for (year, month), count in sorted(counts.items(), key=lambda item: (item[0][0], item[0][1]), reverse=True)
    ]


MANAGEMENT_EDIT_COLUMNS = [
    "purchase_vendor",
    "sales_vendor",
    "transaction_type",
    "ledger_checked",
    "order_date",
    "ship_date",
    "orderer_name",
    "sender_phone",
    "receiver_name",
    "receiver_phone",
    "product_name",
    "quantity",
    "receiver_address",
    "courier",
    "invoice_number",
    "memo",
]


def get_management_record(record_id: int) -> dict[str, str | int]:
    init_db()
    connection = connect_db()
    try:
        row = connection.execute(
            """
            SELECT id, created_at, source_file, source_sheet, source_row, purchase_vendor,
                   sales_vendor, transaction_type, ledger_checked, order_date, ship_date,
                   orderer_name, sender_phone, receiver_name, receiver_phone, product_name,
                   quantity, receiver_address, courier, invoice_number, memo, cs_received_at
              FROM management_records
             WHERE id = ?
            """,
            [record_id],
        ).fetchone()
    finally:
        connection.close()
    if not row:
        raise ValueError("통합관리대장 행을 찾지 못했습니다.")
    return dict(row)


def update_management_record(record_id: int, payload: dict) -> None:
    init_db()
    values = [clean_payload_text(payload, column) for column in MANAGEMENT_EDIT_COLUMNS]
    assignments = ", ".join(f"{column} = ?" for column in MANAGEMENT_EDIT_COLUMNS)
    connection = connect_db()
    try:
        cursor = connection.execute(
            f"UPDATE management_records SET {assignments} WHERE id = ?",
            [*values, record_id],
        )
        connection.commit()
        if cursor.rowcount == 0:
            raise ValueError("수정할 통합관리대장 행을 찾지 못했습니다.")
    finally:
        connection.close()


def delete_management_records(record_ids: list[int]) -> int:
    init_db()
    if not record_ids:
        return 0
    placeholders = ", ".join("?" for _ in record_ids)
    connection = connect_db()
    try:
        cursor = connection.execute(f"DELETE FROM management_records WHERE id IN ({placeholders})", record_ids)
        connection.commit()
        return int(cursor.rowcount or 0)
    finally:
        connection.close()


def create_cs_case_from_management(record_id: int) -> int:
    record = get_management_record(record_id)
    timestamp = now_text()
    source_file = f"통합관리대장:{record.get('source_file', '')}"
    source_sheet = str(record.get("source_sheet", "") or "")
    source_row = int(record.get("source_row", 0) or 0)
    connection = connect_db()
    try:
        cursor = connection.execute(
            """
            INSERT OR IGNORE INTO cs_cases (
                created_at, updated_at, status, vendor_name, original_info, original_invoice,
                product_name, orderer_name, orderer_phone, receiver_name, receiver_phone,
                receiver_address, cs_type, cs_content, source_file, source_sheet, source_row,
                occurred_at, order_date, ship_date, sales_vendor, purchase_vendor, courier, quantity
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                timestamp,
                timestamp,
                "접수",
                record.get("purchase_vendor", ""),
                f"{record.get('ship_date', '')} / {record.get('invoice_number', '')}".strip(" /"),
                record.get("invoice_number", ""),
                record.get("product_name", ""),
                record.get("orderer_name", ""),
                record.get("sender_phone", ""),
                record.get("receiver_name", ""),
                record.get("receiver_phone", ""),
                record.get("receiver_address", ""),
                "",
                "",
                source_file,
                source_sheet,
                source_row,
                record.get("order_date", "") or record.get("ship_date", ""),
                record.get("order_date", ""),
                record.get("ship_date", ""),
                record.get("sales_vendor", ""),
                record.get("purchase_vendor", ""),
                record.get("courier", ""),
                record.get("quantity", ""),
            ],
        )
        if cursor.rowcount:
            case_id = int(cursor.lastrowid)
        else:
            existing = connection.execute(
                "SELECT id FROM cs_cases WHERE source_file = ? AND source_sheet = ? AND source_row = ?",
                [source_file, source_sheet, source_row],
            ).fetchone()
            case_id = int(existing["id"]) if existing else 0
        connection.execute(
            """
            UPDATE management_records
               SET cs_received_at = COALESCE(NULLIF(cs_received_at, ''), ?)
             WHERE id = ?
            """,
            [timestamp, record_id],
        )
        connection.commit()
    finally:
        connection.close()
    if not case_id:
        raise ValueError("CS 처리대장 접수에 실패했습니다.")
    return case_id


def normalized_header(value: object) -> str:
    return re.sub(r"[\s\r\n/]+", "", str(value or "")).strip().lower()


def clean_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = str(value).replace("\r\n", "\n").strip()
    return re.sub(r"[ \t]+", " ", text)


MANAGEMENT_EXPORT_COLUMNS = [
    ("purchase_vendor", "매입거래처"),
    ("sales_vendor", "매출거래처"),
    ("transaction_type", "거래구분"),
    ("ledger_checked", "장부입력확인"),
    ("order_date", "주문일자"),
    ("ship_date", "출고일"),
    ("orderer_name", "주문자"),
    ("sender_phone", "발신자연락처"),
    ("receiver_name", "수령자"),
    ("receiver_phone", "수령자연락처"),
    ("product_name", "제 품 명"),
    ("quantity", "수량"),
    ("receiver_address", "상 세 주 소"),
    ("courier", "택배사"),
    ("invoice_number", "운송장번호"),
    ("memo", "특이사항"),
]


LEDGER_EXPORT_COLUMNS = [
    ("occurred_at", "날짜"),
    ("sales_vendor", "매출거래처"),
    ("purchase_vendor", "매입거래처"),
    ("status", "처리진행상태"),
    ("completed_at", "완료일"),
    ("cs_type", "처리내용"),
    ("cs_content", "C/S 내용"),
    ("reship_invoice", "재발송운송장번호"),
    ("return_invoice", "회수운송장번호"),
    ("order_date", "주문일자"),
    ("ship_date", "출고일"),
    ("orderer_name", "주문자"),
    ("orderer_phone", "연락처"),
    ("receiver_name", "수령자"),
    ("receiver_phone", "연락처"),
    ("product_name", "제품명"),
    ("quantity", "수량"),
    ("receiver_address", "상세주소"),
    ("courier", "택배사"),
    ("original_invoice", "송장번호"),
]


def workbook_bytes_from_rows(rows: list[dict], columns: list[tuple[str, str]], sheet_name: str) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name[:31]
    worksheet.append([header for _, header in columns])
    for row in rows:
        worksheet.append([clean_cell(row.get(key, "")) for key, _ in columns])
    for column_cells in worksheet.columns:
        max_length = max((len(clean_cell(cell.value)) for cell in column_cells), default=0)
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 10), 42)
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def extract_year_month(*values: object) -> tuple[str, str] | None:
    for value in values:
        text = clean_cell(value)
        if not text:
            continue
        match = re.search(r"(\d{4})[-./년\s]+(\d{1,2})", text)
        if match:
            return match.group(1), f"{int(match.group(2)):02d}"
    return None


def grouped_management_rows(rows: list[dict]) -> list[tuple[str, list[dict]]]:
    groups: dict[str, list[dict]] = {}
    for row in rows:
        year_month = extract_year_month(row.get("order_date"), row.get("ship_date"))
        key = f"{year_month[0]}년 {year_month[1]}월" if year_month else "다운로드"
        groups.setdefault(key, []).append(row)
    return sorted(groups.items(), key=lambda item: item[0])


def cell_style_snapshot(cell) -> dict:
    return {
        "font": copy(cell.font),
        "fill": copy(cell.fill),
        "border": copy(cell.border),
        "alignment": copy(cell.alignment),
        "protection": copy(cell.protection),
        "number_format": cell.number_format,
    }


def apply_cell_style(cell, style: dict) -> None:
    cell.font = copy(style["font"])
    cell.fill = copy(style["fill"])
    cell.border = copy(style["border"])
    cell.alignment = copy(style["alignment"])
    cell.protection = copy(style["protection"])
    cell.number_format = style["number_format"]


def clear_management_template_sheet(worksheet, style_row: list[dict], row_height: float | None) -> None:
    if worksheet.max_row > 2:
        worksheet.delete_rows(3, worksheet.max_row - 2)
    for column_index, (_, header) in enumerate(MANAGEMENT_EXPORT_COLUMNS, start=1):
        worksheet.cell(2, column_index, header)
    if row_height:
        worksheet.row_dimensions[3].height = row_height
    for column_index, style in enumerate(style_row, start=1):
        apply_cell_style(worksheet.cell(3, column_index), style)


def populate_management_template_sheet(worksheet, title: str, rows: list[dict], style_row: list[dict], row_height: float | None) -> None:
    worksheet.title = title[:31]
    worksheet.cell(1, 1, "(주)소일브릿지(SOILBRIDGE) 월별 발주 리스트")
    for row_offset, row in enumerate(rows, start=3):
        if row_height:
            worksheet.row_dimensions[row_offset].height = row_height
        for column_index, (key, _) in enumerate(MANAGEMENT_EXPORT_COLUMNS, start=1):
            cell = worksheet.cell(row_offset, column_index, clean_cell(row.get(key, "")))
            apply_cell_style(cell, style_row[column_index - 1])
    last_row = max(2, len(rows) + 2)
    worksheet.auto_filter.ref = f"A2:P{last_row}"
    worksheet.freeze_panes = "A3"


def management_workbook_bytes_from_template(rows: list[dict]) -> bytes:
    if not MANAGEMENT_EXPORT_TEMPLATE.exists():
        return workbook_bytes_from_rows(rows, MANAGEMENT_EXPORT_COLUMNS, "통합관리대장")
    workbook = load_workbook(MANAGEMENT_EXPORT_TEMPLATE)
    base_sheet = workbook.worksheets[0]
    max_columns = len(MANAGEMENT_EXPORT_COLUMNS)
    style_row = [cell_style_snapshot(base_sheet.cell(3, column_index)) for column_index in range(1, max_columns + 1)]
    row_height = base_sheet.row_dimensions[3].height
    for sheet in list(workbook.worksheets)[1:]:
        workbook.remove(sheet)
    clear_management_template_sheet(base_sheet, style_row, row_height)
    groups = grouped_management_rows(rows) or [("다운로드", rows)]
    sheets = [base_sheet]
    for _title, _rows in groups[1:]:
        sheets.append(workbook.copy_worksheet(base_sheet))
    for worksheet, (title, group_rows) in zip(sheets, groups, strict=False):
        populate_management_template_sheet(worksheet, title, group_rows, style_row, row_height)
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def management_export_rows_from_payload(payload: dict) -> tuple[list[dict], str]:
    scope = clean_payload_text(payload, "scope") or "selected"
    if scope == "selected":
        rows = payload.get("rows", [])
        if isinstance(rows, list) and rows:
            return rows, "통합관리대장_선택"
        record_ids = [int(value) for value in payload.get("ids", []) if str(value).isdigit()]
        rows = list_management_records_by_ids(record_ids)
        return rows, "통합관리대장_선택"
    if scope == "month":
        year = clean_payload_text(payload, "year")
        month = clean_payload_text(payload, "month").zfill(2)
        if not re.fullmatch(r"\d{4}", year) or not re.fullmatch(r"\d{2}", month) or not (1 <= int(month) <= 12):
            raise ValueError("월별 다운로드는 년도와 월을 선택해주세요.")
        rows = list_management_records(limit=None, year=year, month=month)
        return rows, f"통합관리대장_{year}년_{int(month)}월"
    if scope == "year":
        year = clean_payload_text(payload, "year")
        if not re.fullmatch(r"\d{4}", year):
            raise ValueError("년별 다운로드는 년도를 선택해주세요.")
        rows = list_management_records(limit=None, year=year)
        return rows, f"통합관리대장_{year}년"
    if scope == "all":
        rows = list_management_records(limit=None)
        return rows, "통합관리대장_전체"
    rows = payload.get("rows", [])
    return rows if isinstance(rows, list) else [], "통합관리대장"


def find_header(headers: list[str], names: set[str]) -> int | None:
    for idx, header in enumerate(headers):
        if header in names:
            return idx
    return None


def find_header_contains(headers: list[str], *parts: str) -> int | None:
    for idx, header in enumerate(headers):
        if all(part in header for part in parts):
            return idx
    return None


def row_value(row: tuple, idx: int | None) -> str:
    if idx is None or idx >= len(row):
        return ""
    return clean_cell(row[idx])


def contact_index_after(headers: list[str], anchor_idx: int | None, before_idx: int | None = None) -> int | None:
    contact_indexes = [idx for idx, header in enumerate(headers) if "연락처" in header]
    if not contact_indexes:
        return None
    if anchor_idx is not None:
        for idx in contact_indexes:
            if idx > anchor_idx and (before_idx is None or idx < before_idx):
                return idx
    return None


def receiver_phone_index(headers: list[str], receiver_idx: int | None) -> int | None:
    contact_indexes = [idx for idx, header in enumerate(headers) if "연락처" in header]
    if not contact_indexes:
        return None
    if receiver_idx is not None:
        for idx in contact_indexes:
            if idx > receiver_idx:
                return idx
    return contact_indexes[-1]


def management_header_indexes(headers: list[str]) -> dict[str, int | None]:
    return {
        "purchase_vendor": find_header(headers, {"매입거래처"}),
        "sales_vendor": find_header(headers, {"매출거래처"}),
        "transaction_type": find_header(headers, {"거래구분"}),
        "ledger_checked": find_header(headers, {"장부입력확인"}),
        "order_date": find_header(headers, {"주문일자"}),
        "ship_date": find_header(headers, {"출고일"}),
        "orderer_name": find_header(headers, {"주문자"}),
        "sender_phone": find_header(headers, {"발신자연락처", "주문자연락처"}),
        "receiver_name": find_header(headers, {"수령자", "수취인", "수하인"}),
        "receiver_phone": find_header(headers, {"수령자연락처", "수취인연락처", "수하인연락처"}),
        "product_name": find_header(headers, {"제품명", "상품명", "품명"}),
        "quantity": find_header(headers, {"수량"}),
        "receiver_address": find_header(headers, {"상세주소", "주소"}),
        "courier": find_header(headers, {"택배사"}),
        "invoice_number": find_header(headers, {"운송장번호", "송장번호"}),
        "memo": find_header(headers, {"특이사항", "배송메세지", "배송메시지", "비고"}),
    }


def import_management_workbook(path: Path) -> tuple[int, int]:
    init_db()
    workbook = load_workbook(path, data_only=True, read_only=True)
    inserted = 0
    skipped = 0
    timestamp = now_text()
    source_file = original_uploaded_filename(path.name)
    columns = [
        "created_at",
        "source_file",
        "source_sheet",
        "source_row",
        "purchase_vendor",
        "sales_vendor",
        "transaction_type",
        "ledger_checked",
        "order_date",
        "ship_date",
        "orderer_name",
        "sender_phone",
        "receiver_name",
        "receiver_phone",
        "product_name",
        "quantity",
        "receiver_address",
        "courier",
        "invoice_number",
        "memo",
    ]
    placeholders = ", ".join("?" for _ in columns)
    connection = connect_db()
    try:
        for worksheet in workbook.worksheets:
            rows = worksheet.iter_rows(max_col=80, values_only=True)
            next(rows, None)
            header_row = next(rows, None)
            if not header_row:
                continue
            headers = [normalized_header(value) for value in header_row]
            indexes = management_header_indexes(headers)
            if indexes["receiver_name"] is None or indexes["product_name"] is None:
                continue
            for excel_row_number, row in enumerate(rows, start=3):
                record = {
                    "created_at": timestamp,
                    "source_file": source_file,
                    "source_sheet": worksheet.title,
                    "source_row": excel_row_number,
                    "purchase_vendor": row_value(row, indexes["purchase_vendor"]),
                    "sales_vendor": row_value(row, indexes["sales_vendor"]),
                    "transaction_type": row_value(row, indexes["transaction_type"]),
                    "ledger_checked": row_value(row, indexes["ledger_checked"]),
                    "order_date": row_value(row, indexes["order_date"]),
                    "ship_date": row_value(row, indexes["ship_date"]),
                    "orderer_name": row_value(row, indexes["orderer_name"]),
                    "sender_phone": row_value(row, indexes["sender_phone"]),
                    "receiver_name": row_value(row, indexes["receiver_name"]),
                    "receiver_phone": row_value(row, indexes["receiver_phone"]),
                    "product_name": row_value(row, indexes["product_name"]),
                    "quantity": row_value(row, indexes["quantity"]),
                    "receiver_address": row_value(row, indexes["receiver_address"]),
                    "courier": row_value(row, indexes["courier"]),
                    "invoice_number": row_value(row, indexes["invoice_number"]),
                    "memo": row_value(row, indexes["memo"]),
                }
                if not any(record[key] for key in ("receiver_name", "product_name", "invoice_number", "receiver_address")):
                    continue
                cursor = connection.execute(
                    f"INSERT OR IGNORE INTO management_records ({', '.join(columns)}) VALUES ({placeholders})",
                    [record[column] for column in columns],
                )
                if cursor.rowcount:
                    inserted += 1
                else:
                    skipped += 1
        connection.commit()
    finally:
        connection.close()
        workbook.close()
    return inserted, skipped


def import_cs_cases_from_workbook(path: Path) -> tuple[int, int]:
    init_db()
    workbook = load_workbook(path, data_only=True, read_only=True)
    inserted = 0
    skipped = 0
    timestamp = now_text()
    source_file = path.name

    insert_sql = """
        INSERT OR IGNORE INTO cs_cases (
            created_at, updated_at, status, vendor_name, vendor_email,
            original_info, original_invoice, product_name, orderer_name, orderer_phone, receiver_name,
            receiver_phone, receiver_address, cs_type, cs_content, return_invoice,
            reship_invoice, mail_subject, mail_body, mail_sent_at,
            source_file, source_sheet, source_row, occurred_at, completed_at,
            order_date, ship_date, sales_vendor, purchase_vendor, courier, quantity
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    connection = connect_db()
    try:
        for worksheet in workbook.worksheets:
            rows = worksheet.iter_rows(max_col=80, values_only=True)
            title_row = next(rows, None)
            header_row = next(rows, None)
            if not header_row:
                continue

            headers = [normalized_header(value) for value in header_row]
            date_idx = find_header(headers, {"발생일", "날짜"})
            sales_idx = find_header(headers, {"발생거래처", "매출거래처"})
            purchase_idx = find_header(headers, {"처리거래처", "매입거래처"})
            status_idx = find_header_contains(headers, "처리", "상태")
            completed_idx = find_header(headers, {"처리완료일", "완료일"})
            order_idx = find_header(headers, {"주문일자"})
            ship_idx = find_header(headers, {"출고일"})
            orderer_idx = find_header(headers, {"주문자", "주문인", "구매자"})
            receiver_idx = find_header(headers, {"수령자", "수취인", "수하인", "받는분", "받으시는분"})
            orderer_phone_idx = contact_index_after(headers, orderer_idx, receiver_idx)
            phone_idx = receiver_phone_index(headers, receiver_idx)
            product_idx = find_header(headers, {"제품명", "상품명", "품명"})
            quantity_idx = find_header(headers, {"수량"})
            address_idx = find_header_contains(headers, "상세", "주소")
            courier_idx = find_header(headers, {"택배사"})
            original_idx = find_header(headers, {"기존운송장번호", "송장번호"})
            request_idx = find_header(headers, {"요구내용", "처리내용"})
            cs_content_idx = find_header(headers, {"cs내용", "c/s내용"})
            return_idx = find_header(headers, {"회수운송장번호"})
            reship_idx = find_header(headers, {"재발송운송장번호", "재발송송장번호"})
            if address_idx is None and quantity_idx is not None and courier_idx is not None and quantity_idx + 1 < courier_idx:
                address_idx = quantity_idx + 1

            for excel_row_number, row in enumerate(rows, start=3):
                row = tuple(row)
                if not any(clean_cell(value) for value in row):
                    continue

                sales_vendor = row_value(row, sales_idx)
                purchase_vendor = row_value(row, purchase_idx)
                vendor_name = purchase_vendor or sales_vendor
                original_invoice = row_value(row, original_idx)
                original_info = " / ".join(
                    value for value in [row_value(row, ship_idx), original_invoice] if value
                )
                request_text = row_value(row, request_idx)
                cs_text = row_value(row, cs_content_idx)
                cs_content = cs_text or request_text
                raw_status = row_value(row, status_idx)
                cs_type = normalize_cs_type_value(request_text, cs_content, raw_status)
                product_name = row_value(row, product_idx)
                orderer_name = row_value(row, orderer_idx)
                orderer_phone = row_value(row, orderer_phone_idx)
                receiver_name = row_value(row, receiver_idx)
                receiver_phone = row_value(row, phone_idx)

                if not any([vendor_name, original_invoice, product_name, receiver_name, receiver_phone, cs_content]):
                    continue

                values = [
                    timestamp,
                    timestamp,
                    normalize_progress_status(raw_status, row_value(row, completed_idx), row_value(row, date_idx)),
                    vendor_name,
                    "",
                    original_info,
                    extract_invoice_number(original_invoice) or original_invoice,
                    product_name,
                    orderer_name,
                    orderer_phone,
                    receiver_name,
                    receiver_phone,
                    row_value(row, address_idx),
                    cs_type,
                    cs_content,
                    row_value(row, return_idx),
                    row_value(row, reship_idx),
                    "",
                    "",
                    "",
                    source_file,
                    worksheet.title,
                    excel_row_number,
                    row_value(row, date_idx),
                    row_value(row, completed_idx),
                    row_value(row, order_idx),
                    row_value(row, ship_idx),
                    sales_vendor,
                    purchase_vendor,
                    row_value(row, courier_idx),
                    row_value(row, quantity_idx),
                ]
                cursor = connection.execute(insert_sql, values)
                if cursor.rowcount:
                    inserted += 1
                else:
                    skipped += 1
        connection.commit()
    finally:
        connection.close()
        workbook.close()

    return inserted, skipped


class DataBlob(ctypes.Structure):
    _fields_ = [
        ("cbData", ctypes.c_ulong),
        ("pbData", ctypes.POINTER(ctypes.c_char)),
    ]


def protect_text(value: str) -> str:
    if not value:
        return ""
    if os.name != "nt":
        return TOKEN_PREFIX_KEY + protect_text_with_secret(value)
    raw = value.encode("utf-8")
    in_buffer = ctypes.create_string_buffer(raw)
    in_blob = DataBlob(len(raw), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_char)))
    out_blob = DataBlob()
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(out_blob),
    ):
        raise OSError("메일 비밀번호 저장 암호화에 실패했습니다.")
    try:
        encrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        return TOKEN_PREFIX_DPAPI + base64.b64encode(encrypted).decode("ascii")
    finally:
        ctypes.windll.kernel32.LocalFree(out_blob.pbData)


def unprotect_text(value: str) -> str:
    if not value:
        return ""
    if value.startswith(TOKEN_PREFIX_KEY):
        return unprotect_text_with_secret(value.removeprefix(TOKEN_PREFIX_KEY))
    if value.startswith(TOKEN_PREFIX_DPAPI):
        value = value.removeprefix(TOKEN_PREFIX_DPAPI)
    if os.name != "nt":
        return unprotect_text_with_secret(value)
    raw = base64.b64decode(value.encode("ascii"))
    in_buffer = ctypes.create_string_buffer(raw)
    in_blob = DataBlob(len(raw), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_char)))
    out_blob = DataBlob()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(out_blob),
    ):
        raise OSError("저장된 메일 비밀번호를 읽지 못했습니다.")
    try:
        decrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        return decrypted.decode("utf-8")
    finally:
        ctypes.windll.kernel32.LocalFree(out_blob.pbData)


def get_server_secret() -> bytes:
    env_secret = os.environ.get("WORKHUB_SECRET_KEY", "").strip()
    if env_secret:
        return hashlib.sha256(env_secret.encode("utf-8")).digest()

    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if SECRET_KEY_PATH.exists():
        return base64.b64decode(SECRET_KEY_PATH.read_text(encoding="utf-8").strip())

    secret = secrets.token_bytes(32)
    SECRET_KEY_PATH.write_text(base64.b64encode(secret).decode("ascii"), encoding="utf-8")
    try:
        os.chmod(SECRET_KEY_PATH, 0o600)
    except OSError:
        pass
    return secret


def key_stream(secret: bytes, nonce: bytes, size: int) -> bytes:
    chunks: list[bytes] = []
    counter = 0
    while sum(len(chunk) for chunk in chunks) < size:
        chunks.append(hashlib.sha256(secret + nonce + counter.to_bytes(4, "big")).digest())
        counter += 1
    return b"".join(chunks)[:size]


def protect_text_with_secret(value: str) -> str:
    secret = get_server_secret()
    nonce = secrets.token_bytes(16)
    raw = value.encode("utf-8")
    encrypted = bytes(left ^ right for left, right in zip(raw, key_stream(secret, nonce, len(raw))))
    signature = hmac.new(secret, nonce + encrypted, hashlib.sha256).digest()
    return base64.b64encode(nonce + signature + encrypted).decode("ascii")


def unprotect_text_with_secret(value: str) -> str:
    secret = get_server_secret()
    raw = base64.b64decode(value.encode("ascii"))
    nonce = raw[:16]
    signature = raw[16:48]
    encrypted = raw[48:]
    expected = hmac.new(secret, nonce + encrypted, hashlib.sha256).digest()
    if not hmac.compare_digest(signature, expected):
        raise OSError("저장된 메일 비밀번호 검증에 실패했습니다.")
    decrypted = bytes(left ^ right for left, right in zip(encrypted, key_stream(secret, nonce, len(encrypted))))
    return decrypted.decode("utf-8")


def normalize_naver_email(value: object) -> str:
    email = str(value or "").strip()
    if email and "@" not in email:
        email = f"{email}@naver.com"
    return email


def load_mail_settings(include_password: bool = False) -> dict[str, str | bool]:
    if not MAIL_SETTINGS_PATH.exists():
        return {"naver_email": "", "has_password": False}
    try:
        settings = json.loads(MAIL_SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"naver_email": "", "has_password": False}
    password_token = str(settings.get("password_token", ""))
    loaded: dict[str, str | bool] = {
        "naver_email": str(settings.get("naver_email", "")),
        "has_password": bool(password_token),
    }
    if include_password and password_token:
        loaded["naver_password"] = unprotect_text(password_token)
    return loaded


def save_mail_settings(naver_email: str, naver_password: str | None = None) -> None:
    settings = load_mail_settings(include_password=False)
    settings["naver_email"] = naver_email
    if naver_password:
        settings["password_token"] = protect_text(naver_password)
        settings["has_password"] = True
    else:
        existing = {}
        if MAIL_SETTINGS_PATH.exists():
            try:
                existing = json.loads(MAIL_SETTINGS_PATH.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                existing = {}
        if existing.get("password_token"):
            settings["password_token"] = existing["password_token"]
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    MAIL_SETTINGS_PATH.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def load_vendor_contacts() -> list[dict[str, str]]:
    if not VENDOR_CONTACTS_PATH.exists():
        return []
    try:
        raw_contacts = json.loads(VENDOR_CONTACTS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    contacts: list[dict[str, str]] = []
    if isinstance(raw_contacts, list):
        for raw in raw_contacts:
            if not isinstance(raw, dict):
                continue
            vendor_name = str(raw.get("vendor_name", "")).strip()
            email = str(raw.get("email", "")).strip()
            if vendor_name and email:
                contacts.append({"vendor_name": vendor_name, "email": email})
    return sorted(contacts, key=lambda item: item["vendor_name"])


def save_vendor_contact(vendor_name: str, email: str) -> list[dict[str, str]]:
    vendor_name = vendor_name.strip()
    email = email.strip()
    if not vendor_name:
        raise ValueError("업체명을 입력해주세요.")
    if not email or "@" not in email:
        raise ValueError("업체 메일주소를 올바르게 입력해주세요.")

    contacts = load_vendor_contacts()
    replaced = False
    for contact in contacts:
        if contact["vendor_name"] == vendor_name:
            contact["email"] = email
            replaced = True
            break
    if not replaced:
        contacts.append({"vendor_name": vendor_name, "email": email})
    contacts = sorted(contacts, key=lambda item: item["vendor_name"])
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    VENDOR_CONTACTS_PATH.write_text(json.dumps(contacts, ensure_ascii=False, indent=2), encoding="utf-8")
    return contacts


def save_vendor_contacts_bulk(new_contacts: list[dict[str, str]]) -> tuple[list[dict[str, str]], int]:
    contacts_by_name = {contact["vendor_name"]: contact["email"] for contact in load_vendor_contacts()}
    saved_count = 0
    for contact in new_contacts:
        vendor_name = str(contact.get("vendor_name", "")).strip()
        email = str(contact.get("email", "")).strip()
        if not vendor_name or not email or "@" not in email:
            continue
        contacts_by_name[vendor_name] = email
        saved_count += 1

    contacts = [
        {"vendor_name": vendor_name, "email": email}
        for vendor_name, email in contacts_by_name.items()
    ]
    contacts = sorted(contacts, key=lambda item: item["vendor_name"])
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    VENDOR_CONTACTS_PATH.write_text(json.dumps(contacts, ensure_ascii=False, indent=2), encoding="utf-8")
    return contacts, saved_count


def header_text(value: object) -> str:
    return re.sub(r"\s+", "", str(value or "")).strip().lower()


def import_vendor_contacts_from_workbook(path: Path) -> tuple[list[dict[str, str]], int]:
    workbook = load_workbook(path, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("업체 메일 주소록 엑셀에 데이터가 없습니다.")

    header = [header_text(value) for value in rows[0]]
    vendor_headers = {"업체명", "거래처명", "회사명", "업체", "거래처"}
    email_headers = {"메일", "메일주소", "이메일", "email", "e-mail", "emailaddress"}
    vendor_idx = next((idx for idx, value in enumerate(header) if value in vendor_headers), None)
    email_idx = next((idx for idx, value in enumerate(header) if value in email_headers), None)

    data_rows = rows[1:] if vendor_idx is not None and email_idx is not None else rows
    if vendor_idx is None or email_idx is None:
        vendor_idx = 0
        email_idx = 1

    imported: list[dict[str, str]] = []
    for row in data_rows:
        vendor_name = str(row[vendor_idx] or "").strip() if len(row) > vendor_idx else ""
        email = str(row[email_idx] or "").strip() if len(row) > email_idx else ""
        if vendor_name and email and "@" in email:
            imported.append({"vendor_name": vendor_name, "email": email})

    if not imported:
        raise ValueError("저장할 업체명/메일주소를 찾지 못했습니다. 엑셀에 업체명과 메일주소 열을 넣어주세요.")

    return save_vendor_contacts_bulk(imported)


def send_cs_mail(payload: dict) -> None:
    saved = load_mail_settings(include_password=True)
    naver_email = normalize_naver_email(payload.get("naver_email")) or str(saved.get("naver_email", ""))
    naver_password = str(payload.get("naver_password") or saved.get("naver_password", ""))
    recipient = str(payload.get("recipient_email", "")).strip()
    subject = str(payload.get("subject", "")).strip()
    body = str(payload.get("body", "")).strip()

    if not naver_email:
        raise ValueError("네이버 메일 아이디를 입력해주세요.")
    if not naver_password:
        raise ValueError("네이버 메일 비밀번호를 입력해주세요.")
    if not recipient:
        raise ValueError("받는 업체 메일을 입력해주세요.")
    if not subject:
        raise ValueError("메일 제목을 입력해주세요.")
    if not body:
        raise ValueError("메일 내용을 입력해주세요.")

    if payload.get("save_credentials", True):
        save_mail_settings(naver_email, naver_password)

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = formataddr(("(주)소일브릿지", naver_email))
    message["To"] = recipient
    message.set_content(body)

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL(NAVER_SMTP_HOST, NAVER_SMTP_PORT, context=context, timeout=20) as smtp:
            smtp.login(naver_email, naver_password)
            smtp.send_message(message)
    except smtplib.SMTPAuthenticationError as exc:
        raise ValueError("네이버 메일 로그인에 실패했습니다. 아이디/비밀번호와 네이버 메일 SMTP 사용 설정을 확인해주세요.") from exc


class WorkhubHandler(BaseHTTPRequestHandler):
    def cookie_value(self, name: str) -> str:
        cookies = self.headers.get("Cookie", "")
        for part in cookies.split(";"):
            if "=" not in part:
                continue
            key, value = part.strip().split("=", 1)
            if key == name:
                return unquote(value)
        return ""

    def current_user(self) -> dict[str, str] | None:
        return current_user_from_token(self.cookie_value(SESSION_COOKIE_NAME))

    def send_redirect(self, location: str, status: int = 303) -> None:
        self.send_response(status)
        self.send_header("Location", location)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()

    def set_session_cookie(self, token: str) -> None:
        self.send_header(
            "Set-Cookie",
            f"{SESSION_COOKIE_NAME}={quote(token)}; Path=/; Max-Age={SESSION_SECONDS}; HttpOnly; SameSite=Lax",
        )

    def clear_session_cookie(self) -> None:
        self.send_header(
            "Set-Cookie",
            f"{SESSION_COOKIE_NAME}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax",
        )

    def require_permission(self, user: dict[str, str], permission: str, label: str) -> bool:
        if user_has_permission(user, permission):
            return True
        self.send_json({"error": f"{label} 권한이 없습니다."}, status=403)
        return False

    def do_GET(self) -> None:
        if self.path.startswith("/lucide/"):
            relative = unquote(self.path.removeprefix("/lucide/"))
            target = (LUCIDE_DIR / relative).resolve()
            if LUCIDE_DIR.resolve() in target.parents and target.is_file():
                content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
                self.send_bytes(target.read_bytes(), content_type)
                return
            if relative == "dist/esm/lucide.js":
                self.send_bytes(LUCIDE_FALLBACK_JS.encode("utf-8"), "application/javascript; charset=utf-8")
                return

        if self.path.startswith("/login"):
            parsed = urlsplit(self.path)
            params = parse_qs(parsed.query)
            self.send_bytes(render_login_html(show_error=params.get("error", [""])[0] == "1").encode("utf-8"), "text/html; charset=utf-8")
            return

        if self.path == "/logout":
            delete_login_session(self.cookie_value(SESSION_COOKIE_NAME))
            self.send_response(303)
            self.clear_session_cookie()
            self.send_header("Location", "/login")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            return

        user = self.current_user()
        if not user:
            if self.path.startswith("/api/"):
                self.send_json({"error": "로그인이 필요합니다."}, status=401)
                return
            self.send_redirect("/login")
            return

        if self.path == "/" or self.path.startswith("/?"):
            self.send_bytes(render_app_html(user).encode("utf-8"), "text/html; charset=utf-8")
            return

        if self.path == "/api/users":
            if not self.require_permission(user, "user_admin", "사용자 관리"):
                return
            self.send_json({"users": list_users()})
            return

        if self.path == "/api/backups":
            if not self.require_permission(user, "backup_manage", "백업 관리"):
                return
            backups = list_backup_files()
            self.send_json({
                "backup_dir": str(BACKUP_DIR),
                "backups": backups,
                "last_backup": backups[0]["created_at"] if backups else "",
                "retention_days": BACKUP_RETENTION_DAYS,
                "auto_backup_hour": AUTO_BACKUP_HOUR,
            })
            return

        if self.path == "/api/system-update":
            if not self.require_permission(user, "system_update", "시스템 업데이트"):
                return
            self.send_json(system_update_payload(fetch=False))
            return

        if self.path.startswith("/api/backup-download"):
            if not self.require_permission(user, "backup_manage", "백업 관리"):
                return
            parsed = urlsplit(self.path)
            params = parse_qs(parsed.query)
            try:
                path = backup_path_from_name(params.get("name", [""])[0])
                if not path.exists():
                    raise FileNotFoundError("백업 파일을 찾지 못했습니다.")
            except Exception as exc:  # noqa: BLE001
                self.send_json({"error": str(exc)}, status=400)
                return
            data = path.read_bytes()
            filename = quote(path.name)
            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{filename}")
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            self.send_header("Pragma", "no-cache")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        if self.path == "/api/leaves":
            if not any(user_has_permission(user, permission) for permission in ("leave_view", "leave_approve", "leave_manage")):
                self.send_json({"error": "연차 조회 권한이 없습니다."}, status=403)
                return
            self.send_json(leave_payload(user))
            return

        if self.path == "/api/mail-settings":
            if not self.require_permission(user, "mail_send", "메일 발송"):
                return
            self.send_json(load_mail_settings(include_password=False))
            return

        if self.path == "/api/vendor-contacts":
            if not self.require_permission(user, "mail_send", "메일 발송"):
                return
            self.send_json({"contacts": load_vendor_contacts()})
            return

        if self.path.startswith("/api/cs-cases"):
            parsed = urlsplit(self.path)
            params = parse_qs(parsed.query)
            query = params.get("q", [""])[0]
            status = params.get("status", [""])[0]
            year = params.get("year", [""])[0]
            month = params.get("month", [""])[0]
            try:
                limit = min(max(int(params.get("limit", ["100"])[0]), 1), 5000)
            except ValueError:
                limit = 100
            self.send_json({"cases": list_cs_cases(query=query, status=status, limit=limit, year=year, month=month)})
            return

        if self.path.startswith("/api/management-records"):
            parsed = urlsplit(self.path)
            params = parse_qs(parsed.query)
            query = params.get("q", [""])[0]
            year = params.get("year", [""])[0]
            month = params.get("month", [""])[0]
            try:
                limit = min(max(int(params.get("limit", ["100"])[0]), 1), 5000)
            except ValueError:
                limit = 100
            self.send_json({"records": list_management_records(query=query, limit=limit, year=year, month=month)})
            return

        if self.path == "/api/management-periods":
            self.send_json({"periods": list_management_periods()})
            return

        self.send_error(404)

    def do_POST(self) -> None:
        try:
            if self.path == "/login":
                length = int(self.headers.get("Content-Length", "0"))
                raw_body = self.rfile.read(length).decode("utf-8")
                payload = parse_qs(raw_body)
                user = authenticate_user(
                    payload.get("username", [""])[0],
                    payload.get("password", [""])[0],
                )
                if not user:
                    self.send_redirect("/login?error=1")
                    return
                token = create_login_session(user["username"])
                self.send_response(303)
                self.set_session_cookie(token)
                self.send_header("Location", "/")
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                return

            user = self.current_user()
            if not user:
                self.send_json({"error": "로그인이 필요합니다."}, status=401)
                return

            if self.path == "/api/users-save":
                if not self.require_permission(user, "user_admin", "사용자 관리"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                user_id = save_user_account(payload, user)
                self.send_json({"message": "사용자 계정을 저장했습니다.", "user_id": user_id, "users": list_users()})
                return

            if self.path == "/api/backup-create":
                if not self.require_permission(user, "backup_manage", "백업 관리"):
                    return
                backup = create_workhub_backup("manual")
                self.send_json({"message": "백업 파일을 생성했습니다.", "backup": backup})
                return

            if self.path == "/api/system-update-check":
                if not self.require_permission(user, "system_update", "시스템 업데이트"):
                    return
                self.send_json(system_update_payload(fetch=True))
                return

            if self.path == "/api/system-update-apply":
                if not self.require_permission(user, "system_update", "시스템 업데이트"):
                    return
                self.send_json(apply_system_update())
                return

            if self.path == "/api/backup-delete":
                if not self.require_permission(user, "backup_manage", "백업 관리"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                path = backup_path_from_name(clean_payload_text(payload, "name"))
                if not path.exists():
                    raise FileNotFoundError("삭제할 백업 파일을 찾지 못했습니다.")
                path.unlink()
                self.send_json({"message": "백업 파일을 삭제했습니다.", "name": path.name})
                return

            if self.path == "/api/backup-restore":
                if not self.require_permission(user, "backup_manage", "백업 관리"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                path = backup_path_from_name(clean_payload_text(payload, "name"))
                if not path.exists():
                    raise FileNotFoundError("복원할 백업 파일을 찾지 못했습니다.")
                self.send_json(restore_workhub_backup(path))
                return

            if self.path == "/api/leave-request":
                if not self.require_permission(user, "leave_view", "연차 조회"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                request_id = create_leave_request(user, payload)
                self.send_json({"message": "연차 신청이 저장되었습니다.", "request_id": request_id})
                return

            if self.path == "/api/leave-decision":
                if not (user_has_permission(user, "leave_approve") or user_has_permission(user, "leave_manage")):
                    self.send_json({"error": "연차 승인 권한이 없습니다."}, status=403)
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                decide_leave_request(int(payload.get("request_id", 0)), user, clean_payload_text(payload, "decision"), clean_payload_text(payload, "comment"))
                self.send_json({"message": "연차 신청을 처리했습니다."})
                return

            if self.path == "/api/leave-balance":
                if not self.require_permission(user, "leave_manage", "연차 관리"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                set_leave_balance(payload, user)
                self.send_json({"message": "직원 연차 기준을 저장했습니다."})
                return

            if self.path == "/api/leave-historical":
                if not self.require_permission(user, "leave_manage", "연차 관리"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                count = add_historical_leave_usage(payload, user)
                self.send_json({"message": f"기존 사용 연차 {count}건을 등록했습니다.", "count": count})
                return

            if self.path == "/api/cs-mail":
                if not self.require_permission(user, "mail_send", "메일 발송"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                send_cs_mail(payload)
                case_id = save_cs_case(payload, status="메일발송", mail_sent=True)
                self.send_json({"message": "CS 요청 메일 전송 및 DB 저장이 완료되었습니다.", "case_id": case_id})
                return

            if self.path == "/api/cs-case":
                if not self.require_permission(user, "ledger_edit", "대장 수정"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                case_id = save_cs_case(payload, status=clean_payload_text(payload, "status") or "접수")
                self.send_json({"message": "CS건을 DB에 저장했습니다.", "case_id": case_id})
                return

            if self.path == "/api/cs-case-update":
                if not self.require_permission(user, "ledger_edit", "대장 수정"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                case_id = int(payload.get("id", 0))
                if not case_id:
                    raise ValueError("수정할 CS건 ID가 없습니다.")
                update_cs_case(case_id, payload)
                self.send_json({"message": "CS 처리내용을 저장했습니다.", "case_id": case_id})
                return

            if self.path == "/api/cs-cases-delete":
                if not self.require_permission(user, "ledger_delete", "대장 삭제"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                case_ids = [int(value) for value in payload.get("ids", []) if str(value).isdigit()]
                if not case_ids:
                    raise ValueError("삭제할 CS 처리대장 행이 없습니다.")
                deleted = delete_cs_cases(case_ids)
                self.send_json({"message": f"CS 처리대장 선택 주문 {deleted}건을 삭제했습니다.", "deleted": deleted})
                return

            if self.path == "/api/management-record-update":
                if not self.require_permission(user, "ledger_edit", "대장 수정"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                record_id = int(payload.get("id", 0))
                if not record_id:
                    raise ValueError("수정할 통합관리대장 행 ID가 없습니다.")
                update_management_record(record_id, payload)
                self.send_json({"message": "통합관리대장 행을 저장했습니다.", "record_id": record_id})
                return

            if self.path == "/api/management-records-delete":
                if not self.require_permission(user, "ledger_delete", "대장 삭제"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                record_ids = [int(value) for value in payload.get("ids", []) if str(value).isdigit()]
                if not record_ids:
                    raise ValueError("삭제할 통합관리대장 행이 없습니다.")
                deleted = delete_management_records(record_ids)
                self.send_json({"message": f"통합관리대장 선택 주문 {deleted}건을 삭제했습니다.", "deleted": deleted})
                return

            if self.path == "/api/management-to-cs":
                if not self.require_permission(user, "cs_receive", "CS접수"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                record_id = int(payload.get("id", 0))
                if not record_id:
                    raise ValueError("CS접수할 통합관리대장 행 ID가 없습니다.")
                case_id = create_cs_case_from_management(record_id)
                self.send_json({"message": "CS 처리대장에 접수했습니다.", "case_id": case_id})
                return

            if self.path == "/api/management-export":
                if not self.require_permission(user, "excel_download", "엑셀 다운로드"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                rows, filename_stem = management_export_rows_from_payload(payload)
                if not rows:
                    raise ValueError("엑셀로 다운로드할 통합관리대장 데이터가 없습니다.")
                data = management_workbook_bytes_from_template(rows)
                filename = quote(f"{filename_stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{filename}")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return

            if self.path == "/api/cs-cases-export":
                if not self.require_permission(user, "excel_download", "엑셀 다운로드"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                rows = payload.get("rows", [])
                if not isinstance(rows, list) or not rows:
                    raise ValueError("엑셀로 다운로드할 CS 처리대장 데이터가 없습니다.")
                data = workbook_bytes_from_rows(rows, LEDGER_EXPORT_COLUMNS, "CS처리대장")
                filename = quote(f"CS처리대장_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{filename}")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return

            if self.path == "/api/vendor-contact":
                if not self.require_permission(user, "mail_send", "메일 발송"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                contacts = save_vendor_contact(
                    str(payload.get("vendor_name", "")),
                    str(payload.get("email", "")),
                )
                self.send_json({"message": "업체 메일 주소를 저장했습니다.", "contacts": contacts})
                return

            if self.path == "/api/vehicle-receipt":
                if not self.require_permission(user, "excel_download", "엑셀 다운로드"):
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                output_path = generate_vehicle_receipt(
                    supplier=payload.get("supplier", ""),
                    items=payload.get("items", []),
                    freight_payment=payload.get("freight_payment", "선불"),
                    receipt_type=payload.get("receipt_type", "일반"),
                    request_note=payload.get("request_note", ""),
                    delivery_place=payload.get("delivery_place", ""),
                    manager=payload.get("manager", ""),
                    output_dir=DOWNLOAD_DIR,
                    output_date=parse_receipt_date(payload.get("receipt_date", "")),
                )
                filename = quote(output_path.name)
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                self.send_header(
                    "Content-Disposition",
                    f"attachment; filename*=UTF-8''{filename}",
                )
                data = output_path.read_bytes()
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return

            fields = parse_multipart(self.headers, self.rfile)

            if self.path == "/api/backup-restore-upload":
                if not self.require_permission(user, "backup_manage", "백업 관리"):
                    return
                upload_path = save_uploaded_backup_zip(fields, "file")
                self.send_json(restore_workhub_backup(upload_path))
                return

            if self.path == "/api/vendor-contacts-import":
                if not self.require_permission(user, "excel_upload", "엑셀 업로드"):
                    return
                upload_path = save_uploaded_file(fields, "file")
                contacts, saved_count = import_vendor_contacts_from_workbook(upload_path)
                self.send_json({
                    "message": f"업체 메일 주소 {saved_count}건을 저장했습니다.",
                    "saved_count": saved_count,
                    "contacts": contacts,
                })
                return

            if self.path == "/api/cs-cases-import":
                if not self.require_permission(user, "excel_upload", "엑셀 업로드"):
                    return
                upload_path = save_uploaded_file(fields, "file")
                inserted, skipped = import_cs_cases_from_workbook(upload_path)
                self.send_json({
                    "message": f"CS 처리대장 {inserted}건 업로드 완료, 중복 {skipped}건 제외했습니다.",
                    "inserted": inserted,
                    "skipped": skipped,
                })
                return

            if self.path == "/api/management-import":
                if not self.require_permission(user, "excel_upload", "엑셀 업로드"):
                    return
                upload_path = save_uploaded_file(fields, "file")
                inserted, skipped = import_management_workbook(upload_path)
                self.send_json({
                    "message": f"통합관리대장 {inserted}건 업로드 완료, 중복 {skipped}건 제외했습니다.",
                    "inserted": inserted,
                    "skipped": skipped,
                })
                return

            if self.path == "/api/delivery-summary":
                if not self.require_permission(user, "excel_upload", "엑셀 업로드"):
                    return
                upload_path = save_uploaded_file(fields)
                sort_mode = fields.get("sort", "name")
                if sort_mode not in {"name", "count", "first"}:
                    sort_mode = "name"
                text, _ = summarize_workbook(upload_path, sort_mode=str(sort_mode))
                lines = [line for line in text.splitlines() if " - " in line]
                self.send_json({"text": text, "line_count": len(lines)})
                return

            if self.path == "/api/invoice-export":
                if not self.require_permission(user, "excel_upload", "엑셀 업로드"):
                    return
                if not self.require_permission(user, "excel_download", "엑셀 다운로드"):
                    return
                upload_path = save_uploaded_file(fields)
                preview_rows = extract_invoice_rows(upload_path)
                output_path = export_invoice_numbers(upload_path, DOWNLOAD_DIR)
                filename = quote(output_path.name)
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                self.send_header(
                    "Content-Disposition",
                    f"attachment; filename*=UTF-8''{filename}",
                )
                self.send_header("X-Row-Count", str(len(preview_rows)))
                data = output_path.read_bytes()
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return

            if self.path == "/api/lotte-order-form":
                if not self.require_permission(user, "excel_upload", "엑셀 업로드"):
                    return
                if not self.require_permission(user, "excel_download", "엑셀 다운로드"):
                    return
                source_path = save_uploaded_file(fields, "file")
                if not LOTTE_TEMPLATE.exists():
                    raise FileNotFoundError(f"롯데택배 발주서 양식을 찾지 못했습니다: {LOTTE_TEMPLATE}")
                output_path = convert_lotte_order_form(source_path, LOTTE_TEMPLATE, DOWNLOAD_DIR)
                filename = quote(output_path.name)
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                self.send_header(
                    "Content-Disposition",
                    f"attachment; filename*=UTF-8''{filename}",
                )
                data = output_path.read_bytes()
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return

            self.send_error(404)
        except Exception as exc:  # noqa: BLE001
            self.send_json({"error": str(exc)}, status=400)

    def send_json(self, payload: dict, status: int = 200) -> None:
        self.send_bytes(
            json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            "application/json; charset=utf-8",
            status,
        )

    def send_bytes(self, data: bytes, content_type: str, status: int = 200) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, format: str, *args) -> None:
        return


def run(host: str = "127.0.0.1", port: int = 8765) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    init_db()
    start_backup_scheduler()
    server = ThreadingHTTPServer((host, port), WorkhubHandler)
    print(f"(주)소일브릿지 발주 업무자동화 앱 실행 중: http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    selected_port = int(os.environ.get("WORKHUB_PORT", sys.argv[1] if len(sys.argv) > 1 else 8765))
    selected_host = os.environ.get("WORKHUB_HOST", "127.0.0.1")
    run(host=selected_host, port=selected_port)
