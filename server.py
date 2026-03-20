from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import io
import json
import os
import re
import secrets
import sqlite3
import threading
from datetime import date, datetime, timedelta, time
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote, unquote, urlparse
import urllib.request

ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
DATA_DIR = ROOT / "data"
DB_PATH = DATA_DIR / "approval.db"
EDITOR_PROVIDER_INTERNAL = "internal"
EDITOR_PROVIDER_GOOGLE_DOCS = "google_docs"
GOOGLE_DOC_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{20,}$")
DEFAULT_GOOGLE_SCOPE = "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/documents.readonly"
USERNAME_PATTERN = re.compile(r"^[a-zA-Z0-9_.-]{3,40}$")
APPROVAL_TEMPLATE_TOTAL_SLOTS = 4
MAX_APPROVER_STEPS = APPROVAL_TEMPLATE_TOTAL_SLOTS - 1
LEAVE_APPROVAL_TEMPLATE_TOTAL_SLOTS = 2
DOC_VISIBILITY_VALUES = {"public", "private", "department"}
LEAVE_TEMPLATE_TYPE = "leave"
OVERTIME_TEMPLATE_TYPE = "overtime"
BUSINESS_TRIP_TEMPLATE_TYPE = "business_trip"
BUSINESS_TRIP_RESULT_TEMPLATE_TYPE = "business_trip_result"
EDUCATION_TEMPLATE_TYPE = "education"
EDUCATION_RESULT_TEMPLATE_TYPE = "education_result"
WORK_START_TIME = time(9, 0)
LUNCH_START_TIME = time(12, 0)
LUNCH_END_TIME = time(13, 0)
WORK_END_TIME = time(18, 0)
WORK_HOURS_PER_DAY = 8.0
ARCHIVE_DELETED_FILES_FOLDER_ID = "1NpYmr1xTdSrappZRC8_UiinoBE73LkTP"
BUSINESS_TRIP_RESULT_TEMPLATE_DOC_ID = "1Cn1G54wChPQcSuU4CFarLViqKoDkvlGp5fzYRHuu7Is"
BUSINESS_TRIP_RESULT_OUTPUT_FOLDER_ID = "1K0UaEg-t39aSLUaXwO4g91snBHBtdFhz"
EDUCATION_DRAFT_OUTPUT_FOLDER_ID = "1MDucU8ZCrWl9O8PEomu4ig_OoVuK6-Yv"
EDUCATION_RESULT_TEMPLATE_DOC_ID = "1zEAqQ3C-xOtt5Duy8Oz49acEBQ3uvue_2x2ELuMP1J0"
EDUCATION_RESULT_OUTPUT_FOLDER_ID = "1MDucU8ZCrWl9O8PEomu4ig_OoVuK6-Yv"
ISSUE_DEPARTMENT_OPTIONS = {
    "복지관",
    "함께배움팀",
    "성장이음팀",
    "건강채움팀",
    "같이돌봄팀",
    "내일이룸팀",
    "미래그림팀",
    "행복동네팀",
    "일상동행팀",
    "누구나운동센터",
}

CORS_ALLOWED_ORIGINS = os.getenv("CORS_ALLOWED_ORIGINS", "").split(",")
CORS_ALLOWED_ORIGINS = [o.strip() for o in CORS_ALLOWED_ORIGINS if o.strip()]

SESSIONS: dict[str, int] = {}
SESSIONS_LOCK = threading.Lock()

DEFAULT_USERS = [
    {"username": "admin", "password": "admin123!", "full_name": "시스템 관리자", "role": "admin", "department": "경영지원"},
    {"username": "ceo", "password": "ceo123!", "full_name": "대표 이민호", "role": "executive", "department": "대표실"},
    {"username": "kim", "password": "kim123!", "full_name": "김지훈", "role": "employee", "department": "사업기획팀"},
    {"username": "lee", "password": "lee123!", "full_name": "이수현", "role": "employee", "department": "인사총무팀"},
    {"username": "park", "password": "park123!", "full_name": "박은영", "role": "employee", "department": "재무회계팀"},
]

DEFAULT_NOTICES = [
    {"title": "전자결재 운영 가이드", "content": "신규 문서는 임시저장 후 결재선을 확인하고 상신하세요. 반려 시 코멘트를 확인해 재상신할 수 있습니다.", "pinned": 1},
    {"title": "사내 공용회의실 예약 정책", "content": "회의실 자원은 일정 메뉴에서 선착순으로 예약되며, 종료 후 10분 내 정리 정돈이 필요합니다.", "pinned": 0},
]


class ApiError(Exception):
    def __init__(self, status: int, message: str) -> None:
        super().__init__(message)
        self.status = status
        self.message = message


def now_ts() -> str:
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")


def today_str() -> str:
    return date.today().isoformat()


def hash_password(password: str) -> str:
    pepper = "ktbizoffice-approval-demo"
    return hashlib.sha256(f"{pepper}:{password}".encode("utf-8")).hexdigest()


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def ensure_schema() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                full_name TEXT NOT NULL,
                role TEXT NOT NULL,
                department TEXT NOT NULL,
                approval_stamp_image_url TEXT,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                template_type TEXT NOT NULL,
                content TEXT NOT NULL,
                editor_provider TEXT NOT NULL DEFAULT 'internal',
                external_doc_id TEXT,
                external_doc_url TEXT,
                status TEXT NOT NULL,
                priority TEXT NOT NULL DEFAULT 'normal',
                due_date TEXT,
                leave_type TEXT,
                leave_start_date TEXT,
                leave_end_date TEXT,
                leave_days REAL,
                leave_reason TEXT,
                leave_substitute_name TEXT,
                leave_substitute_work TEXT,
                overtime_type TEXT,
                overtime_start_date TEXT,
                overtime_end_date TEXT,
                overtime_hours REAL,
                overtime_content TEXT,
                overtime_etc TEXT,
                trip_department TEXT,
                trip_job_title TEXT,
                trip_name TEXT,
                trip_type TEXT,
                trip_destination TEXT,
                trip_start_date TEXT,
                trip_end_date TEXT,
                trip_transportation TEXT,
                trip_expense TEXT,
                trip_purpose TEXT,
                education_department TEXT,
                education_job_title TEXT,
                education_name TEXT,
                education_title TEXT,
                education_category TEXT,
                education_provider TEXT,
                education_location TEXT,
                education_start_date TEXT,
                education_end_date TEXT,
                education_purpose TEXT,
                education_tuition_detail TEXT,
                education_tuition_amount REAL,
                education_material_detail TEXT,
                education_material_amount REAL,
                education_transport_detail TEXT,
                education_transport_amount REAL,
                education_other_detail TEXT,
                education_other_amount REAL,
                education_budget_subject TEXT,
                education_funding_source TEXT,
                education_payment_method TEXT,
                education_support_budget REAL,
                education_used_budget REAL,
                education_remain_budget REAL,
                education_companion TEXT,
                education_ordered TEXT,
                education_suggestion TEXT,
                trip_result TEXT,
                source_trip_document_id INTEGER,
                issue_code TEXT,
                issue_department TEXT,
                issue_year TEXT,
                recipient_text TEXT,
                visibility_scope TEXT NOT NULL DEFAULT 'private',
                attachments_json TEXT NOT NULL DEFAULT '[]',
                drafter_id INTEGER NOT NULL,
                submitted_at TEXT,
                completed_at TEXT,
                reference_ids TEXT NOT NULL DEFAULT '[]',
                is_deleted INTEGER NOT NULL DEFAULT 0,
                deleted_at TEXT,
                 deleted_by INTEGER,
                 print_snapshot_file_id TEXT,
                 print_snapshot_url TEXT,
                 print_snapshot_generated_at TEXT,
                 edit_request_status TEXT NOT NULL DEFAULT 'none',
                edit_request_requested_by INTEGER,
                edit_request_reason TEXT,
                edit_request_requested_at TEXT,
                edit_request_reviewer_id INTEGER,
                edit_request_decided_by INTEGER,
                edit_request_decided_at TEXT,
                delete_request_status TEXT NOT NULL DEFAULT 'none',
                delete_request_requested_by INTEGER,
                delete_request_reason TEXT,
                delete_request_requested_at TEXT,
                delete_request_reviewer_id INTEGER,
                delete_request_decided_by INTEGER,
                delete_request_decided_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY (drafter_id) REFERENCES users(id)
            );
            CREATE TABLE IF NOT EXISTS approval_steps (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                step_order INTEGER NOT NULL,
                approver_id INTEGER NOT NULL,
                status TEXT NOT NULL,
                acted_at TEXT,
                comment TEXT,
                UNIQUE (document_id, step_order),
                FOREIGN KEY (document_id) REFERENCES documents(id) ON DELETE CASCADE,
                FOREIGN KEY (approver_id) REFERENCES users(id)
            );
            CREATE TABLE IF NOT EXISTS document_comments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                user_id INTEGER NOT NULL,
                comment TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (document_id) REFERENCES documents(id) ON DELETE CASCADE,
                FOREIGN KEY (user_id) REFERENCES users(id)
            );
            CREATE TABLE IF NOT EXISTS notices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                author_id INTEGER NOT NULL,
                pinned INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                FOREIGN KEY (author_id) REFERENCES users(id)
            );
            CREATE TABLE IF NOT EXISTS schedules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                event_type TEXT NOT NULL,
                owner_id INTEGER NOT NULL,
                resource_name TEXT,
                status TEXT NOT NULL DEFAULT 'active',
                created_at TEXT NOT NULL,
                FOREIGN KEY (owner_id) REFERENCES users(id)
            );
            CREATE TABLE IF NOT EXISTS notifications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                message TEXT NOT NULL,
                link TEXT,
                is_read INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                FOREIGN KEY (user_id) REFERENCES users(id)
            );
            CREATE TABLE IF NOT EXISTS department_issue_sequences (
                department TEXT PRIMARY KEY,
                last_seq INTEGER NOT NULL DEFAULT 0,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS department_issue_sequences_yearly (
                department TEXT NOT NULL,
                issue_year TEXT NOT NULL,
                last_seq INTEGER NOT NULL DEFAULT 0,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (department, issue_year)
            );
            CREATE TABLE IF NOT EXISTS leave_usages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL UNIQUE,
                user_id INTEGER NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                leave_days REAL NOT NULL,
                status TEXT NOT NULL DEFAULT 'active',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY (document_id) REFERENCES documents(id) ON DELETE CASCADE,
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS user_ui_preferences (
                user_id INTEGER PRIMARY KEY,
                tab_order_json TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL,
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value_json TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL
            );
            """
        )
        ensure_document_columns(conn)
        ensure_user_columns(conn)
        # One-time migration: carry legacy admin-based tab order into global setting.
        has_global_tab_order = conn.execute(
            "SELECT 1 FROM app_settings WHERE key='global_tab_order'"
        ).fetchone()
        if not has_global_tab_order:
            legacy_tab_row = conn.execute(
                """
                SELECT p.tab_order_json
                FROM user_ui_preferences p
                JOIN users u ON u.id = p.user_id
                WHERE u.role = 'admin'
                ORDER BY p.updated_at DESC, p.user_id DESC
                LIMIT 1
                """
            ).fetchone()
            if legacy_tab_row and legacy_tab_row["tab_order_json"]:
                conn.execute(
                    """
                    INSERT INTO app_settings (key, value_json, updated_at)
                    VALUES ('global_tab_order', ?, ?)
                    """,
                    (legacy_tab_row["tab_order_json"], now_ts()),
                )

        if conn.execute("SELECT COUNT(*) c FROM users").fetchone()["c"] == 0:
            for u in DEFAULT_USERS:
                conn.execute(
                    "INSERT INTO users (username, password_hash, full_name, role, department, job_title, total_leave, used_leave, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (u["username"], hash_password(u["password"]), u["full_name"], u["role"], u["department"], u.get("job_title", ""), u.get("total_leave", 0), u.get("used_leave", 0), now_ts()),
                )

        if conn.execute("SELECT COUNT(*) c FROM notices").fetchone()["c"] == 0:
            author = conn.execute("SELECT id FROM users WHERE username='admin'").fetchone()
            if not author:
                author = conn.execute("SELECT id FROM users ORDER BY id ASC LIMIT 1").fetchone()
            
            if author:
                aid = author["id"]
                for n in DEFAULT_NOTICES:
                    conn.execute(
                        "INSERT INTO notices (title, content, author_id, pinned, created_at) VALUES (?, ?, ?, ?, ?)",
                        (n["title"], n["content"], aid, n["pinned"], now_ts()),
                    )

        if conn.execute("SELECT COUNT(*) c FROM schedules").fetchone()["c"] == 0:
            owner = conn.execute("SELECT id FROM users WHERE username='ceo'").fetchone()
            if not owner:
                owner = conn.execute("SELECT id FROM users ORDER BY id ASC LIMIT 1").fetchone()

            if owner:
                oid = owner["id"]
                conn.execute(
                    "INSERT INTO schedules (title, start_date, end_date, event_type, owner_id, resource_name, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    ("주간 임원 회의", today_str(), today_str(), "meeting", oid, "대회의실 A", now_ts()),
                )


def ensure_document_columns(conn: sqlite3.Connection) -> None:
    columns = {row["name"] for row in conn.execute("PRAGMA table_info(documents)").fetchall()}
    if "editor_provider" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN editor_provider TEXT NOT NULL DEFAULT 'internal'")
    if "external_doc_id" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN external_doc_id TEXT")
    if "external_doc_url" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN external_doc_url TEXT")
    if "due_date" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN due_date TEXT")
    if "leave_type" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN leave_type TEXT")
    if "leave_start_date" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN leave_start_date TEXT")
    if "leave_end_date" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN leave_end_date TEXT")
    if "leave_days" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN leave_days REAL")
    if "leave_reason" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN leave_reason TEXT")
    if "leave_substitute_name" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN leave_substitute_name TEXT")
    if "leave_substitute_work" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN leave_substitute_work TEXT")
    if "overtime_type" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN overtime_type TEXT")
    if "overtime_start_date" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN overtime_start_date TEXT")
    if "overtime_end_date" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN overtime_end_date TEXT")
    if "overtime_hours" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN overtime_hours REAL")
    if "overtime_content" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN overtime_content TEXT")
    if "overtime_etc" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN overtime_etc TEXT")
    if "trip_department" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_department TEXT")
    if "trip_job_title" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_job_title TEXT")
    if "trip_name" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_name TEXT")
    if "trip_type" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_type TEXT")
    if "trip_destination" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_destination TEXT")
    if "trip_start_date" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_start_date TEXT")
    if "trip_end_date" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_end_date TEXT")
    if "trip_transportation" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_transportation TEXT")
    if "trip_expense" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_expense TEXT")
    if "trip_purpose" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_purpose TEXT")
    if "education_department" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_department TEXT")
    if "education_job_title" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_job_title TEXT")
    if "education_name" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_name TEXT")
    if "education_title" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_title TEXT")
    if "education_category" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_category TEXT")
    if "education_provider" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_provider TEXT")
    if "education_location" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_location TEXT")
    if "education_start_date" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_start_date TEXT")
    if "education_end_date" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_end_date TEXT")
    if "education_purpose" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_purpose TEXT")
    if "education_tuition_detail" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_tuition_detail TEXT")
    if "education_tuition_amount" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_tuition_amount REAL")
    if "education_material_detail" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_material_detail TEXT")
    if "education_material_amount" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_material_amount REAL")
    if "education_transport_detail" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_transport_detail TEXT")
    if "education_transport_amount" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_transport_amount REAL")
    if "education_other_detail" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_other_detail TEXT")
    if "education_other_amount" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_other_amount REAL")
    if "education_budget_subject" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_budget_subject TEXT")
    if "education_funding_source" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_funding_source TEXT")
    if "education_payment_method" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_payment_method TEXT")
    if "education_support_budget" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_support_budget REAL")
    if "education_used_budget" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_used_budget REAL")
    if "education_remain_budget" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_remain_budget REAL")
    if "education_companion" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_companion TEXT")
    if "education_ordered" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_ordered TEXT")
    if "education_suggestion" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN education_suggestion TEXT")
    if "trip_result" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN trip_result TEXT")
    if "source_trip_document_id" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN source_trip_document_id INTEGER")
    if "issue_code" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN issue_code TEXT")
    if "recipient_text" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN recipient_text TEXT")
    if "issue_department" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN issue_department TEXT")
    if "issue_year" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN issue_year TEXT")
    if "visibility_scope" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN visibility_scope TEXT NOT NULL DEFAULT 'private'")
    if "attachments_json" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN attachments_json TEXT NOT NULL DEFAULT '[]'")
    if "is_deleted" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN is_deleted INTEGER NOT NULL DEFAULT 0")
    if "deleted_at" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN deleted_at TEXT")
    if "deleted_by" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN deleted_by INTEGER")
    if "print_snapshot_file_id" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN print_snapshot_file_id TEXT")
    if "print_snapshot_url" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN print_snapshot_url TEXT")
    if "print_snapshot_generated_at" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN print_snapshot_generated_at TEXT")
    if "edit_request_status" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN edit_request_status TEXT NOT NULL DEFAULT 'none'")
    if "edit_request_requested_by" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN edit_request_requested_by INTEGER")
    if "edit_request_reason" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN edit_request_reason TEXT")
    if "edit_request_requested_at" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN edit_request_requested_at TEXT")
    if "edit_request_reviewer_id" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN edit_request_reviewer_id INTEGER")
    if "edit_request_decided_by" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN edit_request_decided_by INTEGER")
    if "edit_request_decided_at" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN edit_request_decided_at TEXT")
    if "delete_request_status" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN delete_request_status TEXT NOT NULL DEFAULT 'none'")
    if "delete_request_requested_by" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN delete_request_requested_by INTEGER")
    if "delete_request_reason" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN delete_request_reason TEXT")
    if "delete_request_requested_at" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN delete_request_requested_at TEXT")
    if "delete_request_reviewer_id" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN delete_request_reviewer_id INTEGER")
    if "delete_request_decided_by" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN delete_request_decided_by INTEGER")
    if "delete_request_decided_at" not in columns:
        conn.execute("ALTER TABLE documents ADD COLUMN delete_request_decided_at TEXT")


def ensure_user_columns(conn: sqlite3.Connection) -> None:
    columns = {row["name"] for row in conn.execute("PRAGMA table_info(users)").fetchall()}
    if "job_title" not in columns:
        conn.execute("ALTER TABLE users ADD COLUMN job_title TEXT NOT NULL DEFAULT ''")
    if "total_leave" not in columns:
        conn.execute("ALTER TABLE users ADD COLUMN total_leave REAL NOT NULL DEFAULT 0")
    if "used_leave" not in columns:
        conn.execute("ALTER TABLE users ADD COLUMN used_leave REAL NOT NULL DEFAULT 0")
    if "profile_image_url" not in columns:
        conn.execute("ALTER TABLE users ADD COLUMN profile_image_url TEXT")
    if "approval_stamp_image_url" not in columns:
        conn.execute("ALTER TABLE users ADD COLUMN approval_stamp_image_url TEXT")
    if "auth_provider" not in columns:
        conn.execute("ALTER TABLE users ADD COLUMN auth_provider TEXT NOT NULL DEFAULT 'local'")
    if "google_sub" not in columns:
        conn.execute("ALTER TABLE users ADD COLUMN google_sub TEXT")
    if "google_email" not in columns:
        conn.execute("ALTER TABLE users ADD COLUMN google_email TEXT")


def extract_google_doc_id(value: str) -> str | None:
    candidate = value.strip()
    if not candidate:
        return None
    if GOOGLE_DOC_ID_PATTERN.fullmatch(candidate):
        return candidate
    match = re.search(r"/document/d/([A-Za-z0-9_-]{20,})", candidate)
    if match:
        return match.group(1)
    parsed = urlparse(candidate)
    if parsed.netloc.endswith("google.com"):
        query_id = parse_qs(parsed.query).get("id", [None])[0]
        if query_id and GOOGLE_DOC_ID_PATTERN.fullmatch(query_id):
            return query_id
    return None


def build_google_doc_urls(doc_id: str) -> dict[str, str]:
    base = f"https://docs.google.com/document/d/{doc_id}"
    return {"edit_url": f"{base}/edit", "preview_url": f"{base}/preview"}


def document_form_type_label(value: str | None) -> str:
    v = (value or "").strip()
    mapping = {
        "internal": "내부문서",
        "outbound": "외부발신",
        "general": "일반",
        "expense": "지출결의",
        "leave": "휴가계",
        "overtime": "연장근로",
        "business_trip": "출장신청서",
        "business_trip_result": "출장보고서",
        "education": "교육신청서",
        "education_result": "교육보고서",
        "purchase": "구매품의",
    }
    return mapping.get(v, v or "내부문서")


def approval_template_total_slots_for_type(template_type: str | None) -> int:
    return LEAVE_APPROVAL_TEMPLATE_TOTAL_SLOTS if (template_type or "").strip() == LEAVE_TEMPLATE_TYPE else APPROVAL_TEMPLATE_TOTAL_SLOTS


def max_approver_steps_for_type(template_type: str | None) -> int:
    return max(1, approval_template_total_slots_for_type(template_type) - 1)


def google_integration_config() -> dict[str, Any]:
    client_id = (os.environ.get("GOOGLE_OAUTH_CLIENT_ID") or "").strip()
    api_key = (os.environ.get("GOOGLE_API_KEY") or "").strip()
    app_id = (os.environ.get("GOOGLE_CLOUD_APP_ID") or "").strip()
    scope = (os.environ.get("GOOGLE_DRIVE_SCOPE") or DEFAULT_GOOGLE_SCOPE).strip()
    enabled = bool(client_id and api_key)
    return {
        "enabled": enabled,
        "client_id": client_id,
        "api_key": api_key,
        "app_id": app_id,
        "scope": scope,
    }

# Google Apps Script Integration
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwC5EkT8lyKoW_7SNOdepuQ_kG-C6msfdCpZv8GO8fgHjRoWxTsyK0k0w4xYBUs_CII/exec"

def log_to_sheet(username: str, ip: str) -> None:
    if not GOOGLE_SCRIPT_URL: return
    try:
        data = {"action": "log_login", "username": username, "ip": ip}
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL, 
            data=json.dumps(data).encode('utf-8'),
            headers={'Content-Type': 'application/json'}
        )
        # Fire and forget: swallow timeout/network errors so they do not print thread tracebacks.
        def _send_log_request() -> None:
            try:
                with urllib.request.urlopen(req, timeout=5):
                    pass
            except Exception as thread_err:
                print(f"log_to_sheet background request skipped: {thread_err}")
        threading.Thread(target=_send_log_request, daemon=True).start()
    except Exception as e:
        print(f"Failed to log to sheet: {e}")

def sync_user_to_sheet(
    user: dict[str, Any],
    profile_image: dict[str, Any] | None = None,
    approval_stamp_image: dict[str, Any] | None = None,
) -> dict[str, str | None]:
    result_urls: dict[str, str | None] = {
        "profile_image_url": None,
        "approval_stamp_image_url": None,
    }
    if not GOOGLE_SCRIPT_URL:
        return result_urls
    try:
        data: dict[str, Any] = {"action": "sync_user", "user": user}
        if profile_image:
            data["profile_image"] = profile_image
        if approval_stamp_image:
            data["approval_stamp_image"] = approval_stamp_image
        payload = json.dumps(data).encode('utf-8')
        print(
            f"[sync_user_to_sheet] Sending {len(payload)} bytes "
            f"(profile_image={'yes' if profile_image else 'no'}, stamp_image={'yes' if approval_stamp_image else 'no'})"
        )
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=payload,
            headers={'Content-Type': 'application/json'}
        )
        with urllib.request.urlopen(req, timeout=30) as response:
            res_body = response.read().decode('utf-8')
            print(f"[sync_user_to_sheet] Response status={response.status}, length={len(res_body)}")
            print(f"[sync_user_to_sheet] Response body (first 500 chars): {res_body[:500]}")
            # Try parsing JSON directly
            try:
                res_json = json.loads(res_body)
            except json.JSONDecodeError:
                # GAS may wrap response in HTML; try to extract JSON
                import re as _re
                m = _re.search(r'\{.*\}', res_body, _re.DOTALL)
                if m:
                    res_json = json.loads(m.group())
                else:
                    print(f"[sync_user_to_sheet] Could not parse response as JSON")
                    return result_urls
            print(f"[sync_user_to_sheet] Parsed JSON: {json.dumps(res_json, ensure_ascii=False)[:300]}")
            data_obj = res_json.get("data") if isinstance(res_json.get("data"), dict) else {}
            if not isinstance(data_obj, dict):
                data_obj = {}
            result_urls["profile_image_url"] = (
                data_obj.get("profile_image_url")
                or res_json.get("profile_image_url")
            )
            result_urls["approval_stamp_image_url"] = (
                data_obj.get("approval_stamp_image_url")
                or res_json.get("approval_stamp_image_url")
            )
            if data_obj.get("image_error"):
                print(f"[sync_user_to_sheet] PROFILE IMAGE UPLOAD ERROR from GAS: {data_obj.get('image_error')}")
            if data_obj.get("approval_stamp_error"):
                print(f"[sync_user_to_sheet] STAMP IMAGE UPLOAD ERROR from GAS: {data_obj.get('approval_stamp_error')}")
            return result_urls
    except Exception as e:
        print(f"[sync_user_to_sheet] ERROR: {type(e).__name__}: {e}")
        return result_urls

def delete_user_from_sheet(username: str) -> None:
    if not GOOGLE_SCRIPT_URL: return
    try:
        data = {"action": "delete_user", "username": username}
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps(data).encode('utf-8'),
            headers={'Content-Type': 'application/json'}
        )
        threading.Thread(target=lambda: urllib.request.urlopen(req, timeout=5)).start()
    except Exception as e:
        print(f"Failed to delete user from sheet: {e}")


def parse_gas_json_response(raw: str, context: str) -> dict[str, Any]:
    text = (raw or "").strip()
    if not text:
        raise ApiError(400, f"{context}: Apps Script 응답이 비어 있습니다. 웹앱 배포/권한을 확인해 주세요.")
    try:
        parsed = json.loads(text)
        if isinstance(parsed, dict):
            return parsed
        raise ApiError(400, f"{context}: Apps Script 응답 형식이 올바르지 않습니다.")
    except json.JSONDecodeError:
        import re as _re
        m = _re.search(r"\{.*\}", text, _re.DOTALL)
        if m:
            try:
                parsed = json.loads(m.group())
                if isinstance(parsed, dict):
                    return parsed
            except json.JSONDecodeError:
                pass
        head = text[:120].replace("\n", " ").replace("\r", " ")
        if head.startswith("<"):
            raise ApiError(400, f"{context}: Apps Script가 JSON 대신 HTML을 반환했습니다. 웹앱 재배포/공개권한(Anyone) 확인 필요")
        raise ApiError(400, f"{context}: Apps Script 응답 파싱 실패 ({head})")


def validate_approval_template_tokens(doc_id: str, required_slots: int) -> dict[str, Any]:
    if not GOOGLE_SCRIPT_URL:
        raise ApiError(500, "Google Apps Script 연동 URL이 설정되어 있지 않습니다.")
    try:
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps(
                {
                    "action": "validate_approval_template",
                    "doc_id": doc_id,
                    "required_slots": required_slots,
                }
            ).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=20) as response:
            raw = response.read().decode("utf-8")
        payload = parse_gas_json_response(raw, "결재 템플릿 검증")
        if payload.get("status") != "success":
            raise ApiError(400, f"결재 템플릿 검증 실패: {payload.get('message') or 'unknown error'}")
        data = payload.get("data") if isinstance(payload.get("data"), dict) else {}
        if not data.get("valid"):
            missing = [str(x) for x in (data.get("missing_tokens") or [])]
            preview = ", ".join(missing[:8]) if missing else "필수 토큰"
            extra = "" if len(missing) <= 8 else f" 외 {len(missing) - 8}개"
            raise ApiError(400, f"결재 템플릿 토큰이 부족합니다. 누락: {preview}{extra}")
        return data
    except ApiError:
        raise
    except Exception as e:
        raise ApiError(400, f"결재 템플릿 검증 호출 실패: {type(e).__name__}: {e}") from e


def reset_google_approval_doc_slots(doc_id: str, total_slots: int = APPROVAL_TEMPLATE_TOTAL_SLOTS) -> dict[str, Any]:
    if not GOOGLE_SCRIPT_URL:
        raise ApiError(500, "Google Apps Script 연동 URL이 설정되어 있지 않습니다.")
    req = urllib.request.Request(
        GOOGLE_SCRIPT_URL,
        data=json.dumps(
            {
                "action": "reset_approval_doc",
                "doc_id": doc_id,
                "total_slots": int(total_slots or APPROVAL_TEMPLATE_TOTAL_SLOTS),
            }
        ).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )

    timeout_attempts = [30, 60, 120]
    last_timeout_err: Exception | None = None
    for attempt_index, timeout_seconds in enumerate(timeout_attempts, start=1):
        try:
            with urllib.request.urlopen(req, timeout=timeout_seconds) as response:
                raw = response.read().decode("utf-8")
            payload = parse_gas_json_response(raw, "결재칸 초기화")
            if payload.get("status") != "success":
                raise ApiError(400, f"결재칸 초기화 실패: {payload.get('message') or 'unknown error'}")
            return payload.get("data") if isinstance(payload.get("data"), dict) else {}
        except ApiError:
            raise
        except Exception as e:
            msg = str(e).lower()
            is_timeout = isinstance(e, TimeoutError) or ("timed out" in msg)
            if not is_timeout:
                raise ApiError(400, f"결재칸 초기화 호출 실패: {type(e).__name__}: {e}") from e
            last_timeout_err = e
            print(
                f"[reset_approval_doc] timeout doc={doc_id} attempt={attempt_index}/{len(timeout_attempts)} "
                f"timeout={timeout_seconds}s: {type(e).__name__}: {e}"
            )
            if attempt_index >= len(timeout_attempts):
                break
    assert last_timeout_err is not None
    raise ApiError(
        400,
        f"결재칸 초기화 호출 실패(재시도 초과): {type(last_timeout_err).__name__}: {last_timeout_err}",
    ) from last_timeout_err


def populate_google_doc_draft_fields(
    doc_id: str,
    *,
    title: str,
    issue_date: str | None,
    issue_code: str | None,
    doc_form_type: str | None = None,
    recipient_text: str | None = None,
    leave_dept: str | None = None,
    leave_name: str | None = None,
    leave_job_title: str | None = None,
    leave_type: str | None = None,
    leave_period: str | None = None,
    leave_total_days: str | None = None,
    leave_used_days: str | None = None,
    leave_remain_days: str | None = None,
    leave_reason: str | None = None,
    leave_substitute_name: str | None = None,
    leave_substitute_work: str | None = None,
    overtime_dept: str | None = None,
    overtime_name: str | None = None,
    overtime_job_title: str | None = None,
    overtime_type: str | None = None,
    overtime_time: str | None = None,
    overtime_hours: str | None = None,
    overtime_content: str | None = None,
    overtime_etc: str | None = None,
    trip_department: str | None = None,
    trip_job_title: str | None = None,
    trip_name: str | None = None,
    trip_type: str | None = None,
    trip_destination: str | None = None,
    trip_period: str | None = None,
    trip_transportation: str | None = None,
    trip_expense: str | None = None,
    trip_purpose: str | None = None,
    education_department: str | None = None,
    education_job_title: str | None = None,
    education_name: str | None = None,
    education_title: str | None = None,
    education_category: str | None = None,
    education_provider: str | None = None,
    education_location: str | None = None,
    education_period: str | None = None,
    education_purpose: str | None = None,
    education_tuition_detail: str | None = None,
    education_tuition_amount: str | None = None,
    education_material_detail: str | None = None,
    education_material_amount: str | None = None,
    education_transport_detail: str | None = None,
    education_transport_amount: str | None = None,
    education_other_detail: str | None = None,
    education_other_amount: str | None = None,
    education_budget_subject: str | None = None,
    education_funding_source: str | None = None,
    education_payment_method: str | None = None,
    education_support_budget: str | None = None,
    education_used_budget: str | None = None,
    education_remain_budget: str | None = None,
    education_companion: str | None = None,
    education_ordered: str | None = None,
    education_suggestion: str | None = None,
    education_content: str | None = None,
    education_apply_point: str | None = None,
    trip_result: str | None = None,
    draft_date: str | None = None,
) -> dict[str, Any]:
    if not GOOGLE_SCRIPT_URL or not doc_id:
        return {}
    payload_body = {
        "action": "populate_draft_doc_fields",
        "doc_id": doc_id,
        "title": title or "",
        "issue_date": issue_date or "",
        "issue_code": issue_code or "",
        "doc_form_type": doc_form_type or "",
        "recipient_text": recipient_text or "",
        "leave_dept": leave_dept or "",
        "leave_name": leave_name or "",
        "leave_job_title": leave_job_title or "",
        "leave_type": leave_type or "",
        "leave_period": leave_period or "",
        "leave_total_days": leave_total_days or "",
        "leave_used_days": leave_used_days or "",
        "leave_remain_days": leave_remain_days or "",
        "leave_reason": leave_reason or "",
        "leave_substitute_name": leave_substitute_name or "",
        "leave_substitute_work": leave_substitute_work or "",
        "overtime_dept": overtime_dept or "",
        "overtime_name": overtime_name or "",
        "overtime_job_title": overtime_job_title or "",
        "overtime_type": overtime_type or "",
        "overtime_time": overtime_time or "",
        "overtime_hours": overtime_hours or "",
        "overtime_content": overtime_content or "",
        "overtime_etc": overtime_etc or "",
        "trip_department": trip_department or "",
        "trip_job_title": trip_job_title or "",
        "trip_name": trip_name or "",
        "trip_type": trip_type or "",
        "trip_destination": trip_destination or "",
        "trip_period": trip_period or "",
        "trip_transportation": trip_transportation or "",
        "trip_expense": trip_expense or "",
        "trip_purpose": trip_purpose or "",
        "education_department": education_department or "",
        "education_job_title": education_job_title or "",
        "education_name": education_name or "",
        "education_title": education_title or "",
        "education_category": education_category or "",
        "education_provider": education_provider or "",
        "education_location": education_location or "",
        "education_period": education_period or "",
        "education_purpose": education_purpose or "",
        "education_tuition_detail": education_tuition_detail or "",
        "education_tuition_amount": education_tuition_amount or "",
        "education_material_detail": education_material_detail or "",
        "education_material_amount": education_material_amount or "",
        "education_transport_detail": education_transport_detail or "",
        "education_transport_amount": education_transport_amount or "",
        "education_other_detail": education_other_detail or "",
        "education_other_amount": education_other_amount or "",
        "education_budget_subject": education_budget_subject or "",
        "education_funding_source": education_funding_source or "",
        "education_payment_method": education_payment_method or "",
        "education_support_budget": education_support_budget or "",
        "education_used_budget": education_used_budget or "",
        "education_remain_budget": education_remain_budget or "",
        "education_companion": education_companion or "",
        "education_ordered": education_ordered or "",
        "education_suggestion": education_suggestion or "",
        "education_content": education_content or "",
        "education_apply_point": education_apply_point or "",
        "trip_result": trip_result or "",
        "draft_date": draft_date or "",
    }

    def _request_once(timeout_seconds: int) -> dict[str, Any]:
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps(payload_body).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=timeout_seconds) as response:
            raw = response.read().decode("utf-8")
        payload = parse_gas_json_response(raw, "기안 문서 필드 채움")
        if payload.get("status") != "success":
            raise ApiError(400, f"기안 문서 자동입력 실패: {payload.get('message') or 'unknown error'}")
        return payload.get("data") if isinstance(payload.get("data"), dict) else {}

    try:
        try:
            return _request_once(45)
        except TimeoutError:
            # Apps Script may take longer for larger/outbound templates; retry once with a longer timeout.
            return _request_once(90)
    except ApiError:
        raise
    except Exception as e:
        raise ApiError(400, f"기안 문서 자동입력 호출 실패: {type(e).__name__}: {e}") from e


def upload_document_attachments_to_drive(files: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not files:
        return []
    if not GOOGLE_SCRIPT_URL:
        raise ApiError(500, "Google Apps Script 연동 URL이 설정되어 있지 않습니다.")
    try:
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps(
                {
                    "action": "upload_attachments",
                    "files": files,
                }
            ).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=60) as response:
            raw = response.read().decode("utf-8")
        payload = parse_gas_json_response(raw, "첨부파일 업로드")
        if payload.get("status") != "success":
            raise ApiError(400, f"첨부파일 업로드 실패: {payload.get('message') or 'unknown error'}")
        data = payload.get("data") if isinstance(payload.get("data"), dict) else {}
        attachments = data.get("attachments") if isinstance(data.get("attachments"), list) else []
        normalized: list[dict[str, Any]] = []
        for item in attachments:
            if not isinstance(item, dict):
                continue
            file_id = str(item.get("file_id") or item.get("id") or "").strip()
            name = str(item.get("name") or "").strip()
            if not file_id or not name:
                continue
            normalized.append(
                {
                    "file_id": file_id,
                    "name": name,
                    "mime_type": str(item.get("mime_type") or item.get("mimeType") or ""),
                    "size": int(item.get("size") or 0),
                    "web_view_url": str(item.get("web_view_url") or item.get("webViewUrl") or item.get("url") or ""),
                }
            )
        return normalized
    except ApiError:
        raise
    except Exception as e:
        raise ApiError(400, f"첨부파일 업로드 호출 실패: {type(e).__name__}: {e}") from e


def delete_google_drive_file(doc_id: str) -> dict[str, Any]:
    if not GOOGLE_SCRIPT_URL or not doc_id:
        return {"skipped": True, "reason": "Google Apps Script URL not configured or doc_id missing"}
    try:
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps({"action": "delete_drive_file", "doc_id": doc_id}).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=20) as response:
            raw = response.read().decode("utf-8")
        payload = parse_gas_json_response(raw, "Google Drive 문서 삭제")
        if payload.get("status") != "success":
            message = str(payload.get("message") or "unknown error")
            raise ApiError(400, f"Google Drive 문서 삭제 실패: {message}")
        data = payload.get("data") if isinstance(payload.get("data"), dict) else {}
        print(f"[delete_google_drive_file] OK doc_id={doc_id} data={json.dumps(data, ensure_ascii=False)[:200]}")
        return data
    except ApiError:
        raise
    except Exception as e:
        print(f"[delete_google_drive_file] ERROR doc_id={doc_id}: {type(e).__name__}: {e}")
        raise ApiError(400, f"Google Drive 문서 삭제 호출 실패: {type(e).__name__}: {e}") from e


def move_google_drive_file(doc_id: str, target_folder_id: str) -> dict[str, Any]:
    if not GOOGLE_SCRIPT_URL or not doc_id:
        return {"skipped": True, "reason": "Google Apps Script URL not configured or doc_id missing"}
    if not target_folder_id:
        raise ApiError(500, "삭제 보관 폴더 ID가 설정되어 있지 않습니다.")
    try:
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps(
                {"action": "move_drive_file", "doc_id": doc_id, "target_folder_id": target_folder_id}
            ).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=30) as response:
            raw = response.read().decode("utf-8")
        payload = parse_gas_json_response(raw, "Google Drive 파일 이동")
        if payload.get("status") != "success":
            message = str(payload.get("message") or "unknown error")
            raise ApiError(400, f"Google Drive 파일 이동 실패: {message}")
        data = payload.get("data") if isinstance(payload.get("data"), dict) else {}
        print(
            f"[move_google_drive_file] OK doc_id={doc_id} folder={target_folder_id} "
            f"data={json.dumps(data, ensure_ascii=False)[:200]}"
        )
        return data
    except ApiError:
        raise
    except Exception as e:
        print(f"[move_google_drive_file] ERROR doc_id={doc_id}: {type(e).__name__}: {e}")
        raise ApiError(400, f"Google Drive 파일 이동 호출 실패: {type(e).__name__}: {e}") from e


def copy_google_doc_template_to_folder(template_doc_id: str, target_folder_id: str, title: str) -> dict[str, str]:
    if not GOOGLE_SCRIPT_URL:
        raise ApiError(500, "Google Apps Script 연동 URL이 설정되어 있지 않습니다.")
    if not template_doc_id:
        raise ApiError(500, "출장결과 템플릿 문서 ID가 설정되어 있지 않습니다.")
    if not target_folder_id:
        raise ApiError(500, "출장결과 저장 폴더 ID가 설정되어 있지 않습니다.")
    try:
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps(
                {
                    "action": "copy_template_doc",
                    "template_doc_id": template_doc_id,
                    "target_folder_id": target_folder_id,
                    "title": title or "출장보고서",
                }
            ).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=60) as response:
            raw = response.read().decode("utf-8")
        payload = parse_gas_json_response(raw, "출장결과 문서 사본 생성")
        if payload.get("status") != "success":
            raise ApiError(400, f"출장결과 문서 사본 생성 실패: {payload.get('message') or 'unknown error'}")
        data = payload.get("data") if isinstance(payload.get("data"), dict) else {}
        copy_data = data.get("copy") if isinstance(data.get("copy"), dict) else data
        doc_id = str(copy_data.get("doc_id") or copy_data.get("file_id") or copy_data.get("id") or "").strip()
        if not doc_id:
            raise ApiError(400, "출장결과 문서 사본 생성 실패: 문서 ID를 받지 못했습니다.")
        urls = build_google_doc_urls(doc_id)
        edit_url = str(copy_data.get("edit_url") or copy_data.get("web_view_url") or copy_data.get("url") or urls["edit_url"])
        name = str(copy_data.get("name") or title or "출장보고서")
        return {"doc_id": doc_id, "edit_url": edit_url, "name": name}
    except ApiError:
        raise
    except Exception as e:
        raise ApiError(400, f"출장결과 문서 사본 생성 호출 실패: {type(e).__name__}: {e}") from e


def delete_document_related_drive_files(doc_row: sqlite3.Row) -> None:
    if "external_doc_id" in doc_row.keys() and doc_row["external_doc_id"]:
        delete_google_drive_file(str(doc_row["external_doc_id"]))
    if "print_snapshot_file_id" in doc_row.keys() and doc_row["print_snapshot_file_id"]:
        delete_google_drive_file(str(doc_row["print_snapshot_file_id"]))
    attachments_to_delete = json.loads(doc_row["attachments_json"] or "[]") if "attachments_json" in doc_row.keys() else []
    if isinstance(attachments_to_delete, list):
        for att in attachments_to_delete:
            if isinstance(att, dict) and att.get("file_id"):
                delete_google_drive_file(str(att["file_id"]))


def move_document_related_files_to_archive_folder(doc_row: sqlite3.Row) -> None:
    file_ids: list[str] = []
    seen: set[str] = set()

    def _add_file_id(raw_id: Any) -> None:
        file_id = str(raw_id or "").strip()
        if not file_id or file_id in seen:
            return
        seen.add(file_id)
        file_ids.append(file_id)

    if "external_doc_id" in doc_row.keys():
        _add_file_id(doc_row["external_doc_id"])
    if "print_snapshot_file_id" in doc_row.keys():
        _add_file_id(doc_row["print_snapshot_file_id"])
    if "attachments_json" in doc_row.keys():
        try:
            attachments = json.loads(doc_row["attachments_json"] or "[]")
        except Exception:
            attachments = []
        if isinstance(attachments, list):
            for att in attachments:
                if isinstance(att, dict):
                    _add_file_id(att.get("file_id"))

    for file_id in file_ids:
        move_google_drive_file(file_id, ARCHIVE_DELETED_FILES_FOLDER_ID)


def apply_delete_mode_to_document(
    conn: sqlite3.Connection,
    *,
    document_id: int,
    doc_row: sqlite3.Row,
    actor_id: int,
    mode: str,
) -> None:
    now = now_ts()
    if mode == "archive":
        move_document_related_files_to_archive_folder(doc_row)
        conn.execute(
            "UPDATE documents SET is_deleted=1, deleted_at=?, deleted_by=?, updated_at=? WHERE id=?",
            (now, actor_id, now, document_id),
        )
        sync_leave_usage_for_document(conn, document_id)
        return
    if mode == "purge":
        conn.execute("UPDATE documents SET is_deleted=1, updated_at=? WHERE id=?", (now, document_id))
        sync_leave_usage_for_document(conn, document_id)
        delete_document_related_drive_files(doc_row)
        conn.execute("DELETE FROM documents WHERE id=?", (document_id,))
        return
    raise ApiError(400, "삭제 모드는 archive 또는 purge 여야 합니다.")


def apply_delete_mode_for_linked_trip_result_documents(
    conn: sqlite3.Connection,
    *,
    source_document_id: int,
    actor_id: int,
    mode: str,
) -> list[int]:
    child_rows = conn.execute(
        """
        SELECT id, title, external_doc_id, attachments_json, print_snapshot_file_id
        FROM documents
        WHERE source_trip_document_id=?
        """,
        (source_document_id,),
    ).fetchall()
    deleted_ids: list[int] = []
    for child in child_rows:
        child_id = int(child["id"])
        apply_delete_mode_to_document(
            conn,
            document_id=child_id,
            doc_row=child,
            actor_id=actor_id,
            mode=mode,
        )
        deleted_ids.append(child_id)
    return deleted_ids


def create_google_doc_pdf_snapshot(
    doc_id: str,
    *,
    snapshot_name: str | None = None,
    replace_file_id: str | None = None,
) -> dict[str, Any]:
    if not GOOGLE_SCRIPT_URL or not doc_id:
        raise ApiError(500, "Google Apps Script 연동 URL 또는 문서 ID가 없습니다.")
    body: dict[str, Any] = {"action": "create_pdf_snapshot", "doc_id": doc_id}
    if snapshot_name:
        body["snapshot_name"] = snapshot_name
    if replace_file_id:
        body["replace_file_id"] = replace_file_id
    try:
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=120) as response:
            raw = response.read().decode("utf-8")
        payload = parse_gas_json_response(raw, "문서 인쇄용 PDF 생성")
        if payload.get("status") != "success":
            raise ApiError(400, f"문서 인쇄용 PDF 생성 실패: {payload.get('message') or 'unknown error'}")
        data = payload.get("data") if isinstance(payload.get("data"), dict) else {}
        info = data.get("snapshot") if isinstance(data.get("snapshot"), dict) else data
        file_id = str(info.get("file_id") or info.get("id") or "").strip() if isinstance(info, dict) else ""
        if not file_id:
            raise ApiError(400, "문서 인쇄용 PDF 생성 실패: 파일 ID를 받지 못했습니다.")
        return {
            "file_id": file_id,
            "web_view_url": str(info.get("web_view_url") or info.get("webViewUrl") or info.get("url") or f"https://drive.google.com/file/d/{file_id}/view"),
            "name": str(info.get("name") or ""),
            "mime_type": str(info.get("mime_type") or info.get("mimeType") or "application/pdf"),
            "size": int(info.get("size") or 0),
        }
    except ApiError:
        raise
    except Exception as e:
        raise ApiError(400, f"문서 인쇄용 PDF 생성 호출 실패: {type(e).__name__}: {e}") from e


def read_google_drive_file_bytes(file_id: str) -> dict[str, Any]:
    if not GOOGLE_SCRIPT_URL or not file_id:
        raise ApiError(500, "Google Apps Script 연동 URL 또는 파일 ID가 없습니다.")
    try:
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps({"action": "read_drive_file_base64", "file_id": file_id}).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=120) as response:
            raw = response.read().decode("utf-8")
        payload = parse_gas_json_response(raw, "인쇄용 PDF 파일 읽기")
        if payload.get("status") != "success":
            raise ApiError(400, f"인쇄용 PDF 파일 읽기 실패: {payload.get('message') or 'unknown error'}")
        data = payload.get("data") if isinstance(payload.get("data"), dict) else {}
        file_info = data.get("file") if isinstance(data.get("file"), dict) else data
        b64 = str(file_info.get("data_base64") or file_info.get("data") or "").strip() if isinstance(file_info, dict) else ""
        if not b64:
            raise ApiError(400, "인쇄용 PDF 파일 읽기 실패: 파일 데이터가 비어 있습니다.")
        try:
            content = base64.b64decode(b64)
        except Exception as e:
            raise ApiError(400, f"인쇄용 PDF 파일 디코딩 실패: {type(e).__name__}: {e}") from e
        return {
            "bytes": content,
            "mime_type": str(file_info.get("mime_type") or file_info.get("mimeType") or "application/pdf"),
            "name": str(file_info.get("name") or "print.pdf"),
            "file_id": str(file_info.get("file_id") or file_info.get("id") or file_id),
        }
    except ApiError:
        raise
    except Exception as e:
        raise ApiError(400, f"인쇄용 PDF 파일 읽기 호출 실패: {type(e).__name__}: {e}") from e


def ensure_document_print_snapshot(conn: sqlite3.Connection, document_id: int, *, force: bool = False) -> dict[str, Any] | None:
    row = fetch_document(conn, document_id)
    if not row:
        return None
    if (row["editor_provider"] or EDITOR_PROVIDER_INTERNAL) != EDITOR_PROVIDER_GOOGLE_DOCS:
        return None
    if not row["external_doc_id"]:
        return None
    if ("is_deleted" in row.keys()) and int(row["is_deleted"] or 0):
        return None
    if (row["status"] or "") not in {"approved", "rejected"}:
        return None

    existing_file_id = str(row["print_snapshot_file_id"] or "").strip() if "print_snapshot_file_id" in row.keys() else ""
    existing_url = str(row["print_snapshot_url"] or "").strip() if "print_snapshot_url" in row.keys() else ""
    if existing_file_id and existing_url and not force:
        return {
            "file_id": existing_file_id,
            "web_view_url": existing_url,
            "generated_at": row["print_snapshot_generated_at"] if "print_snapshot_generated_at" in row.keys() else None,
        }

    snapshot_name = f"{row['title']}_인쇄본_{document_id}.pdf"
    snapshot = create_google_doc_pdf_snapshot(
        str(row["external_doc_id"]),
        snapshot_name=snapshot_name,
        replace_file_id=existing_file_id or None,
    )
    now = now_ts()
    conn.execute(
        "UPDATE documents SET print_snapshot_file_id=?, print_snapshot_url=?, print_snapshot_generated_at=?, updated_at=? WHERE id=?",
        (snapshot["file_id"], snapshot["web_view_url"], now, now, document_id),
    )
    snapshot["generated_at"] = now
    return snapshot


def allocate_department_issue_code(
    conn: sqlite3.Connection,
    department: str,
    issue_date_value: str | None = None,
    issue_year_value: str | None = None,
) -> str:
    dept = (department or "").strip() or "기타"
    if issue_year_value and re.fullmatch(r"\d{4}", str(issue_year_value).strip()):
        issue_year = str(issue_year_value).strip()
    else:
        issue_year = str(date.fromisoformat(issue_date_value).year) if issue_date_value else str(date.today().year)
    prefix = f"{dept}-{issue_year}-"
    rows = conn.execute(
        """
        SELECT issue_code
        FROM documents
        WHERE COALESCE(is_deleted,0)=0
          AND issue_code IS NOT NULL
          AND issue_code LIKE ?
        """,
        (f"{prefix}%",),
    ).fetchall()

    next_seq = 1
    pattern = re.compile(rf"^{re.escape(dept)}-{re.escape(issue_year)}-(\d+)$")
    for row in rows:
        code = str(row["issue_code"] or "").strip()
        m = pattern.match(code)
        if not m:
            continue
        try:
            seq = int(m.group(1))
        except ValueError:
            continue
        if seq >= next_seq:
            next_seq = seq + 1
    return f"{dept}-{issue_year}-{next_seq:02d}"


def sync_approval_signatures_to_google_doc(conn: sqlite3.Connection, document_id: int) -> None:
    if not GOOGLE_SCRIPT_URL:
        return
    try:
        doc = fetch_document(conn, document_id)
        if not doc:
            return
        if (doc["editor_provider"] or EDITOR_PROVIDER_INTERNAL) != EDITOR_PROVIDER_GOOGLE_DOCS:
            return
        if not doc["external_doc_id"]:
            return

        drafter = conn.execute(
            "SELECT full_name, approval_stamp_image_url FROM users WHERE id=?",
            (doc["drafter_id"],),
        ).fetchone()
        if not drafter:
            return

        slots: list[dict[str, Any]] = [
            {
                "slot_index": 1,
                "name": drafter["full_name"],
                "stamp_url": drafter["approval_stamp_image_url"] or "",
                "kind": "drafter",
            }
        ]

        step_rows = conn.execute(
            """
            SELECT s.step_order, s.status, u.full_name, u.approval_stamp_image_url
            FROM approval_steps s
            JOIN users u ON u.id = s.approver_id
            WHERE s.document_id = ?
            ORDER BY s.step_order ASC
            """,
            (document_id,),
        ).fetchall()
        for row in step_rows:
            slots.append(
                {
                    "slot_index": int(row["step_order"]) + 1,
                    "name": row["full_name"],
                    "stamp_url": row["approval_stamp_image_url"] or "",
                    "status": row["status"],
                    "kind": "approver",
                }
            )

        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL,
            data=json.dumps(
                {
                    "action": "update_approval_doc",
                    "doc_id": doc["external_doc_id"],
                    "document_id": document_id,
                    "total_slots": approval_template_total_slots_for_type(doc["template_type"] if "template_type" in doc.keys() else None),
                    "used_slots": len(slots),
                    "slots": slots,
                }
            ).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=20) as response:
            res_body = response.read().decode("utf-8")
            print(f"[sync_approval_signatures] doc={document_id} status={response.status} body={res_body[:400]}")
    except Exception as e:
        # Best effort: approval flow should continue even if Google Docs update fails.
        print(f"[sync_approval_signatures] ERROR doc={document_id}: {type(e).__name__}: {e}")



def dict_user(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "username": row["username"],
        "full_name": row["full_name"],
        "role": row["role"],
        "department": row["department"],
        "job_title": row["job_title"],
        "total_leave": row["total_leave"],
        "used_leave": row["used_leave"],
        "auth_provider": row["auth_provider"] if "auth_provider" in row.keys() else "local",
        "google_email": row["google_email"] if "google_email" in row.keys() else None,
        "profile_image_url": row["profile_image_url"] if "profile_image_url" in row.keys() else None,
        "approval_stamp_image_url": row["approval_stamp_image_url"] if "approval_stamp_image_url" in row.keys() else None,
    }


def get_user_by_id(conn: sqlite3.Connection, user_id: int) -> sqlite3.Row | None:
    return conn.execute(
        "SELECT id, username, full_name, role, department, job_title, total_leave, used_leave, auth_provider, google_email, profile_image_url, approval_stamp_image_url FROM users WHERE id = ?",
        (user_id,),
    ).fetchone()


def get_token(headers: Any) -> str | None:
    auth = headers.get("Authorization", "")
    if auth.startswith("Bearer "):
        return auth[7:].strip() or None
    return None


def auth_user(conn: sqlite3.Connection, headers: Any) -> sqlite3.Row | None:
    token = get_token(headers)
    if not token:
        return None
    with SESSIONS_LOCK:
        user_id = SESSIONS.get(token)
    if not user_id:
        return None
    return get_user_by_id(conn, user_id)


def require_user(conn: sqlite3.Connection, headers: Any) -> sqlite3.Row:
    user = auth_user(conn, headers)
    if not user:
        raise ApiError(401, "로그인이 필요합니다.")
    return user


def new_session(user_id: int) -> str:
    token = secrets.token_hex(24)
    with SESSIONS_LOCK:
        SESSIONS[token] = user_id
    return token


def drop_session(headers: Any) -> None:
    token = get_token(headers)
    if not token:
        return
    with SESSIONS_LOCK:
        SESSIONS.pop(token, None)


def google_auth_config_public() -> dict[str, Any]:
    # Keep backward compatibility: prefer OAuth client id used by integration flow,
    # then fallback to legacy GOOGLE_CLIENT_ID if present.
    client_id = (
        os.environ.get("GOOGLE_OAUTH_CLIENT_ID")
        or os.environ.get("GOOGLE_CLIENT_ID")
        or ""
    ).strip()
    return {"enabled": bool(client_id), "client_id": client_id}


def verify_google_id_token(id_token: str, expected_client_id: str) -> dict[str, Any]:
    token = (id_token or "").strip()
    if not token:
        raise ApiError(400, "Google 인증 토큰이 필요합니다.")
    if not expected_client_id:
        raise ApiError(500, "서버에 Google Client ID가 설정되지 않았습니다.")

    verify_url = f"https://oauth2.googleapis.com/tokeninfo?id_token={quote(token)}"
    try:
        with urllib.request.urlopen(verify_url, timeout=15) as response:
            raw = response.read().decode("utf-8")
    except Exception as e:
        raise ApiError(401, f"Google 토큰 검증 실패: {e}")

    try:
        payload = json.loads(raw)
    except json.JSONDecodeError:
        raise ApiError(401, "Google 토큰 검증 응답 파싱에 실패했습니다.")

    if not isinstance(payload, dict):
        raise ApiError(401, "Google 토큰 검증 응답 형식이 올바르지 않습니다.")
    if payload.get("error_description") or payload.get("error"):
        raise ApiError(401, f"Google 인증 실패: {payload.get('error_description') or payload.get('error')}")

    aud = str(payload.get("aud") or "").strip()
    if aud != expected_client_id:
        raise ApiError(401, "Google 인증 대상(Client ID)이 일치하지 않습니다.")

    email = str(payload.get("email") or "").strip().lower()
    if not email:
        raise ApiError(401, "Google 계정 이메일을 확인할 수 없습니다.")
    email_verified = str(payload.get("email_verified") or "").strip().lower()
    if email_verified not in {"true", "1"}:
        raise ApiError(401, "이메일 인증이 완료된 Google 계정만 가입할 수 있습니다.")

    sub = str(payload.get("sub") or "").strip()
    if not sub:
        raise ApiError(401, "Google 계정 고유 식별자(sub)가 없습니다.")

    return payload


def build_google_username(conn: sqlite3.Connection, email: str, sub: str) -> str:
    local = email.split("@", 1)[0].strip().lower()
    local = re.sub(r"[^a-z0-9_.-]", ".", local)
    local = re.sub(r"\.{2,}", ".", local).strip("._-")
    if len(local) < 3:
        tail = re.sub(r"[^a-z0-9]", "", sub.lower())[-8:] or "googleusr"
        local = f"g.{tail}"
    base = local[:40]
    if not USERNAME_PATTERN.fullmatch(base):
        base = re.sub(r"[^a-zA-Z0-9_.-]", "", base)
    if len(base) < 3 or not USERNAME_PATTERN.fullmatch(base):
        tail = re.sub(r"[^a-z0-9]", "", sub.lower())[-10:] or secrets.token_hex(5)
        base = f"guser{tail}"[:40]

    candidate = base
    idx = 1
    while True:
        exists = conn.execute("SELECT 1 FROM users WHERE username=?", (candidate,)).fetchone()
        if not exists:
            return candidate
        suffix = f".{idx}"
        candidate = f"{base[: max(1, 40 - len(suffix))]}{suffix}"
        idx += 1
        if idx > 9999:
            raise ApiError(500, "Google 계정용 사용자 아이디 생성에 실패했습니다.")


def _add_cors_headers(handler: BaseHTTPRequestHandler) -> None:
    origin = handler.headers.get("Origin", "")
    if CORS_ALLOWED_ORIGINS and origin in CORS_ALLOWED_ORIGINS:
        handler.send_header("Access-Control-Allow-Origin", origin)
        handler.send_header("Access-Control-Allow-Credentials", "true")
        handler.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        handler.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")


def json_response(handler: BaseHTTPRequestHandler, status: int, payload: Any) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    _add_cors_headers(handler)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def text_response(handler: BaseHTTPRequestHandler, status: int, text: str, content_type: str) -> None:
    body = text.encode("utf-8")
    handler.send_response(status)
    _add_cors_headers(handler)
    handler.send_header("Content-Type", content_type)
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def bytes_response(
    handler: BaseHTTPRequestHandler,
    status: int,
    data: bytes,
    content_type: str,
    filename: str | None = None,
    disposition: str = "inline",
) -> None:
    handler.send_response(status)
    _add_cors_headers(handler)
    handler.send_header("Content-Type", content_type)
    handler.send_header("Content-Length", str(len(data)))
    if filename:
        safe_name = filename.replace("\r", "").replace("\n", "").strip() or "print.pdf"
        try:
            safe_name.encode("latin-1")
            handler.send_header("Content-Disposition", f'{disposition}; filename="{safe_name}"')
        except UnicodeEncodeError:
            ascii_fallback = "print.pdf"
            handler.send_header(
                "Content-Disposition",
                f"{disposition}; filename=\"{ascii_fallback}\"; filename*=UTF-8''{quote(safe_name)}",
            )
    handler.end_headers()
    handler.wfile.write(data)


def parse_json_body(handler: BaseHTTPRequestHandler) -> dict[str, Any]:
    content_len = int(handler.headers.get("Content-Length", "0"))
    if content_len <= 0:
        return {}
    raw = handler.rfile.read(content_len)
    if not raw:
        return {}
    try:
        return json.loads(raw.decode("utf-8"))
    except json.JSONDecodeError as exc:
        raise ApiError(400, f"JSON 파싱 오류: {exc.msg}") from exc


def parse_json_body_optional(handler: BaseHTTPRequestHandler) -> dict[str, Any]:
    try:
        return parse_json_body(handler)
    except ApiError:
        return {}


def create_notification(conn: sqlite3.Connection, user_id: int, message: str, link: str | None = None) -> None:
    conn.execute(
        "INSERT INTO notifications (user_id, message, link, is_read, created_at) VALUES (?, ?, ?, 0, ?)",
        (user_id, message, link, now_ts()),
    )


def notify_admins(conn: sqlite3.Connection, message: str, link: str | None = None, exclude_user_id: int | None = None) -> None:
    rows = conn.execute("SELECT id FROM users WHERE role='admin'").fetchall()
    for r in rows:
        if exclude_user_id is not None and int(r["id"]) == int(exclude_user_id):
            continue
        create_notification(conn, int(r["id"]), message, link)


def fetch_steps(conn: sqlite3.Connection, document_id: int) -> list[dict[str, Any]]:
    rows = conn.execute(
        """
        SELECT s.id, s.step_order, s.status, s.acted_at, s.comment,
               u.id approver_id, u.full_name approver_name, u.department
        FROM approval_steps s JOIN users u ON u.id = s.approver_id
        WHERE s.document_id = ? ORDER BY s.step_order ASC
        """,
        (document_id,),
    ).fetchall()
    return [
        {
            "id": r["id"],
            "step_order": r["step_order"],
            "status": r["status"],
            "acted_at": r["acted_at"],
            "comment": r["comment"],
            "approver": {"id": r["approver_id"], "name": r["approver_name"], "department": r["department"]},
        }
        for r in rows
    ]


def fetch_comments(conn: sqlite3.Connection, document_id: int) -> list[dict[str, Any]]:
    rows = conn.execute(
        """
        SELECT c.id, c.comment, c.created_at, u.id user_id, u.full_name
        FROM document_comments c JOIN users u ON u.id = c.user_id
        WHERE c.document_id = ? ORDER BY c.created_at ASC
        """,
        (document_id,),
    ).fetchall()
    return [
        {"id": r["id"], "comment": r["comment"], "created_at": r["created_at"], "user": {"id": r["user_id"], "name": r["full_name"]}}
        for r in rows
    ]


def fetch_document(conn: sqlite3.Connection, document_id: int) -> sqlite3.Row | None:
    return conn.execute(
        """
        SELECT d.*, u.full_name drafter_name, u.department drafter_department
        FROM documents d JOIN users u ON u.id = d.drafter_id
        WHERE d.id = ?
        """,
        (document_id,),
    ).fetchone()


def dict_document(conn: sqlite3.Connection, row: sqlite3.Row, include_detail: bool = False) -> dict[str, Any]:
    current = conn.execute(
        """
        SELECT s.step_order, u.full_name approver_name
        FROM approval_steps s JOIN users u ON u.id = s.approver_id
        WHERE s.document_id = ? AND s.status = 'pending'
        ORDER BY s.step_order ASC LIMIT 1
        """,
        (row["id"],),
    ).fetchone()
    latest_rejected = conn.execute(
        """
        SELECT s.step_order, s.acted_at, s.comment,
               u.id approver_id, u.full_name approver_name, u.department approver_department
        FROM approval_steps s
        JOIN users u ON u.id = s.approver_id
        WHERE s.document_id = ? AND s.status = 'rejected'
        ORDER BY s.acted_at DESC, s.step_order DESC
        LIMIT 1
        """,
        (row["id"],),
    ).fetchone()
    out = {
        "id": row["id"],
        "title": row["title"],
        "template_type": row["template_type"],
        "editor_provider": row["editor_provider"] or EDITOR_PROVIDER_INTERNAL,
        "status": row["status"],
        "priority": row["priority"],
        "due_date": row["due_date"],
        "leave_type": row["leave_type"] if "leave_type" in row.keys() else None,
        "leave_start_date": row["leave_start_date"] if "leave_start_date" in row.keys() else None,
        "leave_end_date": row["leave_end_date"] if "leave_end_date" in row.keys() else None,
        "leave_days": row["leave_days"] if "leave_days" in row.keys() else None,
        "leave_reason": row["leave_reason"] if "leave_reason" in row.keys() else None,
        "leave_substitute_name": row["leave_substitute_name"] if "leave_substitute_name" in row.keys() else None,
        "leave_substitute_work": row["leave_substitute_work"] if "leave_substitute_work" in row.keys() else None,
        "overtime_type": row["overtime_type"] if "overtime_type" in row.keys() else None,
        "overtime_start_date": row["overtime_start_date"] if "overtime_start_date" in row.keys() else None,
        "overtime_end_date": row["overtime_end_date"] if "overtime_end_date" in row.keys() else None,
        "overtime_hours": row["overtime_hours"] if "overtime_hours" in row.keys() else None,
        "overtime_content": row["overtime_content"] if "overtime_content" in row.keys() else None,
        "overtime_etc": row["overtime_etc"] if "overtime_etc" in row.keys() else None,
        "trip_department": row["trip_department"] if "trip_department" in row.keys() else None,
        "trip_job_title": row["trip_job_title"] if "trip_job_title" in row.keys() else None,
        "trip_name": row["trip_name"] if "trip_name" in row.keys() else None,
        "trip_type": row["trip_type"] if "trip_type" in row.keys() else None,
        "trip_destination": row["trip_destination"] if "trip_destination" in row.keys() else None,
        "trip_start_date": row["trip_start_date"] if "trip_start_date" in row.keys() else None,
        "trip_end_date": row["trip_end_date"] if "trip_end_date" in row.keys() else None,
        "trip_transportation": row["trip_transportation"] if "trip_transportation" in row.keys() else None,
        "trip_expense": row["trip_expense"] if "trip_expense" in row.keys() else None,
        "trip_purpose": row["trip_purpose"] if "trip_purpose" in row.keys() else None,
        "education_department": row["education_department"] if "education_department" in row.keys() else None,
        "education_job_title": row["education_job_title"] if "education_job_title" in row.keys() else None,
        "education_name": row["education_name"] if "education_name" in row.keys() else None,
        "education_title": row["education_title"] if "education_title" in row.keys() else None,
        "education_category": row["education_category"] if "education_category" in row.keys() else None,
        "education_provider": row["education_provider"] if "education_provider" in row.keys() else None,
        "education_location": row["education_location"] if "education_location" in row.keys() else None,
        "education_start_date": row["education_start_date"] if "education_start_date" in row.keys() else None,
        "education_end_date": row["education_end_date"] if "education_end_date" in row.keys() else None,
        "education_purpose": row["education_purpose"] if "education_purpose" in row.keys() else None,
        "education_tuition_detail": row["education_tuition_detail"] if "education_tuition_detail" in row.keys() else None,
        "education_tuition_amount": row["education_tuition_amount"] if "education_tuition_amount" in row.keys() else None,
        "education_material_detail": row["education_material_detail"] if "education_material_detail" in row.keys() else None,
        "education_material_amount": row["education_material_amount"] if "education_material_amount" in row.keys() else None,
        "education_transport_detail": row["education_transport_detail"] if "education_transport_detail" in row.keys() else None,
        "education_transport_amount": row["education_transport_amount"] if "education_transport_amount" in row.keys() else None,
        "education_other_detail": row["education_other_detail"] if "education_other_detail" in row.keys() else None,
        "education_other_amount": row["education_other_amount"] if "education_other_amount" in row.keys() else None,
        "education_budget_subject": row["education_budget_subject"] if "education_budget_subject" in row.keys() else None,
        "education_funding_source": row["education_funding_source"] if "education_funding_source" in row.keys() else None,
        "education_payment_method": row["education_payment_method"] if "education_payment_method" in row.keys() else None,
        "education_support_budget": row["education_support_budget"] if "education_support_budget" in row.keys() else None,
        "education_used_budget": row["education_used_budget"] if "education_used_budget" in row.keys() else None,
        "education_remain_budget": row["education_remain_budget"] if "education_remain_budget" in row.keys() else None,
        "education_companion": row["education_companion"] if "education_companion" in row.keys() else None,
        "education_ordered": row["education_ordered"] if "education_ordered" in row.keys() else None,
        "education_suggestion": row["education_suggestion"] if "education_suggestion" in row.keys() else None,
        "trip_result": row["trip_result"] if "trip_result" in row.keys() else None,
        "source_trip_document_id": row["source_trip_document_id"] if "source_trip_document_id" in row.keys() else None,
        "issue_code": row["issue_code"] if "issue_code" in row.keys() else None,
        "issue_department": row["issue_department"] if "issue_department" in row.keys() else None,
        "issue_year": row["issue_year"] if "issue_year" in row.keys() else None,
        "recipient_text": row["recipient_text"] if "recipient_text" in row.keys() else None,
        "visibility_scope": row["visibility_scope"] if "visibility_scope" in row.keys() else "private",
        "drafter": {"id": row["drafter_id"], "name": row["drafter_name"], "department": row["drafter_department"]},
        "submitted_at": row["submitted_at"],
        "completed_at": row["completed_at"],
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
        "reference_ids": json.loads(row["reference_ids"] or "[]"),
        "attachments": json.loads(row["attachments_json"] or "[]") if "attachments_json" in row.keys() else [],
        "current_step": None,
        "latest_rejection": None,
        "returned_for_resubmit": False,
        "external_doc": None,
        "print_snapshot": None,
        "is_deleted": bool(row["is_deleted"]) if "is_deleted" in row.keys() else False,
        "deleted_at": row["deleted_at"] if "deleted_at" in row.keys() else None,
        "deleted_by": row["deleted_by"] if "deleted_by" in row.keys() else None,
        "edit_request": {
            "status": (row["edit_request_status"] if "edit_request_status" in row.keys() else "none") or "none",
            "requested_by": row["edit_request_requested_by"] if "edit_request_requested_by" in row.keys() else None,
            "reason": row["edit_request_reason"] if "edit_request_reason" in row.keys() else None,
            "requested_at": row["edit_request_requested_at"] if "edit_request_requested_at" in row.keys() else None,
            "reviewer_id": row["edit_request_reviewer_id"] if "edit_request_reviewer_id" in row.keys() else None,
            "decided_by": row["edit_request_decided_by"] if "edit_request_decided_by" in row.keys() else None,
            "decided_at": row["edit_request_decided_at"] if "edit_request_decided_at" in row.keys() else None,
        },
        "delete_request": {
            "status": (row["delete_request_status"] if "delete_request_status" in row.keys() else "none") or "none",
            "requested_by": row["delete_request_requested_by"] if "delete_request_requested_by" in row.keys() else None,
            "reason": row["delete_request_reason"] if "delete_request_reason" in row.keys() else None,
            "requested_at": row["delete_request_requested_at"] if "delete_request_requested_at" in row.keys() else None,
            "reviewer_id": row["delete_request_reviewer_id"] if "delete_request_reviewer_id" in row.keys() else None,
            "decided_by": row["delete_request_decided_by"] if "delete_request_decided_by" in row.keys() else None,
            "decided_at": row["delete_request_decided_at"] if "delete_request_decided_at" in row.keys() else None,
        },
    }
    if out["editor_provider"] == EDITOR_PROVIDER_GOOGLE_DOCS and row["external_doc_id"]:
        urls = build_google_doc_urls(row["external_doc_id"])
        out["external_doc"] = {
            "doc_id": row["external_doc_id"],
            "edit_url": row["external_doc_url"] or urls["edit_url"],
            "preview_url": urls["preview_url"],
        }
    if "print_snapshot_file_id" in row.keys() and row["print_snapshot_file_id"]:
        out["print_snapshot"] = {
            "file_id": row["print_snapshot_file_id"],
            "web_view_url": row["print_snapshot_url"] or f"https://drive.google.com/file/d/{row['print_snapshot_file_id']}/view",
            "generated_at": row["print_snapshot_generated_at"] if "print_snapshot_generated_at" in row.keys() else None,
        }
    if current:
        out["current_step"] = {"order": current["step_order"], "approver_name": current["approver_name"]}
    if latest_rejected:
        out["latest_rejection"] = {
            "step_order": latest_rejected["step_order"],
            "acted_at": latest_rejected["acted_at"],
            "comment": latest_rejected["comment"],
            "approver": {
                "id": latest_rejected["approver_id"],
                "name": latest_rejected["approver_name"],
                "department": latest_rejected["approver_department"],
            },
        }
        if out["status"] == "in_review" and out["current_step"] is None:
            out["returned_for_resubmit"] = True
    req = out.get("edit_request") if isinstance(out.get("edit_request"), dict) else None
    if req:
        if req.get("requested_by"):
            u = conn.execute("SELECT id, full_name FROM users WHERE id=?", (int(req["requested_by"]),)).fetchone()
            if u:
                req["requested_by_name"] = u["full_name"]
        if req.get("reviewer_id"):
            u = conn.execute("SELECT id, full_name FROM users WHERE id=?", (int(req["reviewer_id"]),)).fetchone()
            if u:
                req["reviewer_name"] = u["full_name"]
        if req.get("decided_by"):
            u = conn.execute("SELECT id, full_name FROM users WHERE id=?", (int(req["decided_by"]),)).fetchone()
            if u:
                req["decided_by_name"] = u["full_name"]
    del_req = out.get("delete_request") if isinstance(out.get("delete_request"), dict) else None
    if del_req:
        if del_req.get("requested_by"):
            u = conn.execute("SELECT id, full_name FROM users WHERE id=?", (int(del_req["requested_by"]),)).fetchone()
            if u:
                del_req["requested_by_name"] = u["full_name"]
        if del_req.get("reviewer_id"):
            u = conn.execute("SELECT id, full_name FROM users WHERE id=?", (int(del_req["reviewer_id"]),)).fetchone()
            if u:
                del_req["reviewer_name"] = u["full_name"]
        if del_req.get("decided_by"):
            u = conn.execute("SELECT id, full_name FROM users WHERE id=?", (int(del_req["decided_by"]),)).fetchone()
            if u:
                del_req["decided_by_name"] = u["full_name"]
    if include_detail:
        out["content"] = row["content"]
        out["approval_steps"] = fetch_steps(conn, row["id"])
        out["comments"] = fetch_comments(conn, row["id"])
    return out


def can_view_document_row(conn: sqlite3.Connection, row: sqlite3.Row, viewer: sqlite3.Row) -> bool:
    if viewer["role"] == "admin":
        return True
    if row["drafter_id"] == viewer["id"]:
        return True

    # Approval participants and explicit references must be able to view the document regardless of public scope.
    ref_ids = []
    try:
        ref_ids = json.loads(row["reference_ids"] or "[]")
    except Exception:
        ref_ids = []
    if isinstance(ref_ids, list) and viewer["id"] in ref_ids:
        return True
    participant = conn.execute(
        "SELECT 1 FROM approval_steps WHERE document_id=? AND approver_id=? LIMIT 1",
        (row["id"], viewer["id"]),
    ).fetchone()
    if participant:
        return True

    scope = (row["visibility_scope"] if "visibility_scope" in row.keys() else "private") or "private"
    if scope == "public":
        return True
    if scope == "department":
        return (row["drafter_department"] or "") == (viewer["department"] or "")
    return False


def can_open_original_doc_for_viewer(row: sqlite3.Row, viewer: sqlite3.Row) -> bool:
    status = (row["status"] if "status" in row.keys() else "") or ""
    if status not in {"approved", "rejected"}:
        return True
    if status == "rejected" and int(row["drafter_id"]) == int(viewer["id"]):
        # Rejected documents must be editable by the drafter for resubmission.
        return True
    req_status = (row["edit_request_status"] if "edit_request_status" in row.keys() else "none") or "none"
    req_user_id = row["edit_request_requested_by"] if "edit_request_requested_by" in row.keys() else None
    reviewer_id = row["edit_request_reviewer_id"] if "edit_request_reviewer_id" in row.keys() else None
    if reviewer_id is not None and int(reviewer_id) == int(viewer["id"]) and req_status == "pending":
        return True
    return req_status == "approved" and req_user_id is not None and int(req_user_id) == int(viewer["id"])


def request_completed_document_edit(conn: sqlite3.Connection, document_id: int, actor: sqlite3.Row, reason: str, reviewer_id: int) -> dict[str, Any]:
    doc = fetch_document(conn, document_id)
    if not doc:
        raise ApiError(404, "문서를 찾을 수 없습니다.")
    if ("is_deleted" in doc.keys()) and doc["is_deleted"]:
        raise ApiError(409, "보관삭제된 문서는 수정요청할 수 없습니다.")
    if not can_view_document_row(conn, doc, actor):
        raise ApiError(403, "해당 문서를 열람할 권한이 없습니다.")
    if (doc["status"] or "") not in {"approved", "rejected"}:
        raise ApiError(409, "완료된 문서에서만 수정요청할 수 있습니다.")
    if not reason.strip():
        raise ApiError(400, "수정요청 사유를 입력해 주세요.")
    if reviewer_id <= 0:
        raise ApiError(400, "수정요청 결재자를 선택해 주세요.")
    if int(actor["id"]) == int(reviewer_id):
        raise ApiError(400, "본인을 수정요청 결재자로 지정할 수 없습니다.")
    reviewer = conn.execute("SELECT id, role, full_name FROM users WHERE id=?", (reviewer_id,)).fetchone()
    if not reviewer:
        raise ApiError(400, "선택한 수정요청 결재자를 찾을 수 없습니다.")
    is_doc_approver = conn.execute(
        "SELECT 1 FROM approval_steps WHERE document_id=? AND approver_id=? LIMIT 1",
        (document_id, reviewer_id),
    ).fetchone()
    if reviewer["role"] != "admin" and not is_doc_approver:
        raise ApiError(400, "수정요청 결재자는 해당 문서의 결재선 사용자 또는 관리자여야 합니다.")
    current_status = (doc["edit_request_status"] if "edit_request_status" in doc.keys() else "none") or "none"
    current_requester = doc["edit_request_requested_by"] if "edit_request_requested_by" in doc.keys() else None
    delete_req_status = (doc["delete_request_status"] if "delete_request_status" in doc.keys() else "none") or "none"
    if delete_req_status == "pending":
        raise ApiError(409, "삭제요청이 대기 중인 문서는 수정요청할 수 없습니다.")
    if current_status == "pending":
        if current_requester and int(current_requester) == int(actor["id"]):
            raise ApiError(409, "이미 수정요청이 접수되었습니다.")
        raise ApiError(409, "이미 다른 수정요청이 대기 중입니다.")

    now = now_ts()
    conn.execute(
        """
        UPDATE documents
        SET edit_request_status='pending',
            edit_request_requested_by=?,
            edit_request_reason=?,
            edit_request_requested_at=?,
            edit_request_reviewer_id=?,
            edit_request_decided_by=NULL,
            edit_request_decided_at=NULL,
            updated_at=?
        WHERE id=?
        """,
        (actor["id"], reason.strip(), now, reviewer_id, now, document_id),
    )
    create_notification(conn, int(reviewer_id), f"완료 문서 수정요청 결재: {doc['title']}", f"/documents/{document_id}")
    notify_admins(conn, f"완료 문서 수정요청 접수: {doc['title']}", f"/documents/{document_id}", exclude_user_id=int(actor["id"]))
    refreshed = fetch_document(conn, document_id)
    if not refreshed:
        raise ApiError(500, "문서를 다시 불러오지 못했습니다.")
    return dict_document(conn, refreshed, include_detail=True)


def request_completed_document_delete(conn: sqlite3.Connection, document_id: int, actor: sqlite3.Row, reason: str, reviewer_id: int) -> dict[str, Any]:
    doc = fetch_document(conn, document_id)
    if not doc:
        raise ApiError(404, "문서를 찾을 수 없습니다.")
    if ("is_deleted" in doc.keys()) and doc["is_deleted"]:
        raise ApiError(409, "보관삭제된 문서는 삭제요청할 수 없습니다.")
    if not can_view_document_row(conn, doc, actor):
        raise ApiError(403, "해당 문서를 열람할 권한이 없습니다.")
    if (doc["status"] or "") not in {"approved", "rejected"}:
        raise ApiError(409, "완료된 문서에서만 삭제요청할 수 있습니다.")
    if not reason.strip():
        raise ApiError(400, "삭제요청 사유를 입력해 주세요.")
    if reviewer_id <= 0:
        raise ApiError(400, "삭제요청 결재자를 선택해 주세요.")
    if int(actor["id"]) == int(reviewer_id):
        raise ApiError(400, "본인을 삭제요청 결재자로 지정할 수 없습니다.")
    reviewer = conn.execute("SELECT id, role, full_name FROM users WHERE id=?", (reviewer_id,)).fetchone()
    if not reviewer:
        raise ApiError(400, "선택한 삭제요청 결재자를 찾을 수 없습니다.")
    is_doc_approver = conn.execute(
        "SELECT 1 FROM approval_steps WHERE document_id=? AND approver_id=? LIMIT 1",
        (document_id, reviewer_id),
    ).fetchone()
    if reviewer["role"] != "admin" and not is_doc_approver:
        raise ApiError(400, "삭제요청 결재자는 해당 문서의 결재선 사용자 또는 관리자여야 합니다.")
    edit_req_status = (doc["edit_request_status"] if "edit_request_status" in doc.keys() else "none") or "none"
    if edit_req_status == "pending":
        raise ApiError(409, "수정요청이 대기 중인 문서는 삭제요청할 수 없습니다.")
    current_status = (doc["delete_request_status"] if "delete_request_status" in doc.keys() else "none") or "none"
    current_requester = doc["delete_request_requested_by"] if "delete_request_requested_by" in doc.keys() else None
    if current_status == "pending":
        if current_requester and int(current_requester) == int(actor["id"]):
            raise ApiError(409, "이미 삭제요청이 접수되었습니다.")
        raise ApiError(409, "이미 다른 삭제요청이 대기 중입니다.")

    now = now_ts()
    conn.execute(
        """
        UPDATE documents
        SET delete_request_status='pending',
            delete_request_requested_by=?,
            delete_request_reason=?,
            delete_request_requested_at=?,
            delete_request_reviewer_id=?,
            delete_request_decided_by=NULL,
            delete_request_decided_at=NULL,
            updated_at=?
        WHERE id=?
        """,
        (actor["id"], reason.strip(), now, reviewer_id, now, document_id),
    )
    create_notification(conn, int(reviewer_id), f"완료 문서 삭제요청 결재: {doc['title']}", f"/documents/{document_id}")
    notify_admins(conn, f"완료 문서 삭제요청 접수: {doc['title']}", f"/documents/{document_id}", exclude_user_id=int(actor["id"]))
    refreshed = fetch_document(conn, document_id)
    if not refreshed:
        raise ApiError(500, "문서를 다시 불러오지 못했습니다.")
    return dict_document(conn, refreshed, include_detail=True)


def decide_completed_document_delete_request(conn: sqlite3.Connection, document_id: int, actor: sqlite3.Row, approve: bool) -> dict[str, Any]:
    doc = fetch_document(conn, document_id)
    if not doc:
        raise ApiError(404, "문서를 찾을 수 없습니다.")
    if ("is_deleted" in doc.keys()) and doc["is_deleted"]:
        raise ApiError(409, "보관삭제된 문서는 삭제요청을 처리할 수 없습니다.")
    current_status = (doc["delete_request_status"] if "delete_request_status" in doc.keys() else "none") or "none"
    if current_status != "pending":
        raise ApiError(409, "처리할 삭제요청이 없습니다.")
    req_user_id = doc["delete_request_requested_by"] if "delete_request_requested_by" in doc.keys() else None
    reviewer_id = doc["delete_request_reviewer_id"] if "delete_request_reviewer_id" in doc.keys() else None
    if actor["role"] != "admin":
        if reviewer_id is None or int(reviewer_id) != int(actor["id"]):
            raise ApiError(403, "지정된 삭제요청 결재자 또는 관리자만 처리할 수 있습니다.")
    now = now_ts()
    if approve:
        move_document_related_files_to_archive_folder(doc)
    next_status = "approved" if approve else "rejected"
    conn.execute(
        "UPDATE documents SET delete_request_status=?, delete_request_decided_by=?, delete_request_decided_at=?, updated_at=? WHERE id=?",
        (next_status, actor["id"], now, now, document_id),
    )
    if approve:
        conn.execute(
            "UPDATE documents SET is_deleted=1, deleted_at=?, deleted_by=?, updated_at=? WHERE id=?",
            (now, actor["id"], now, document_id),
        )
        sync_leave_usage_for_document(conn, document_id)
    if req_user_id:
        create_notification(
            conn,
            int(req_user_id),
            f"문서 '{doc['title']}' 삭제요청이 {'수락' if approve else '거절'}되었습니다.",
            f"/documents/{document_id}",
        )
    refreshed = fetch_document(conn, document_id)
    if not refreshed:
        # 승인 후 보관삭제됐어도 row 자체는 남아 있어야 정상
        raise ApiError(500, "문서를 다시 불러오지 못했습니다.")
    return dict_document(conn, refreshed, include_detail=True)


def decide_completed_document_edit_request(conn: sqlite3.Connection, document_id: int, actor: sqlite3.Row, approve: bool) -> dict[str, Any]:
    doc = fetch_document(conn, document_id)
    if not doc:
        raise ApiError(404, "문서를 찾을 수 없습니다.")
    if ("is_deleted" in doc.keys()) and doc["is_deleted"]:
        raise ApiError(409, "보관삭제된 문서는 수정요청을 처리할 수 없습니다.")
    current_status = (doc["edit_request_status"] if "edit_request_status" in doc.keys() else "none") or "none"
    if current_status != "pending":
        raise ApiError(409, "처리할 수정요청이 없습니다.")
    req_user_id = doc["edit_request_requested_by"] if "edit_request_requested_by" in doc.keys() else None
    reviewer_id = doc["edit_request_reviewer_id"] if "edit_request_reviewer_id" in doc.keys() else None
    if actor["role"] != "admin":
        if reviewer_id is None or int(reviewer_id) != int(actor["id"]):
            raise ApiError(403, "지정된 수정요청 결재자 또는 관리자만 처리할 수 있습니다.")
    now = now_ts()
    next_status = "approved" if approve else "rejected"
    conn.execute(
        "UPDATE documents SET edit_request_status=?, edit_request_decided_by=?, edit_request_decided_at=?, updated_at=? WHERE id=?",
        (next_status, actor["id"], now, now, document_id),
    )
    if req_user_id:
        create_notification(
            conn,
            int(req_user_id),
            f"문서 '{doc['title']}' 수정요청이 {'수락' if approve else '거절'}되었습니다.",
            f"/documents/{document_id}",
        )
    refreshed = fetch_document(conn, document_id)
    if not refreshed:
        raise ApiError(500, "문서를 다시 불러오지 못했습니다.")
    return dict_document(conn, refreshed, include_detail=True)


def complete_completed_document_edit_request(conn: sqlite3.Connection, document_id: int, actor: sqlite3.Row) -> dict[str, Any]:
    doc = fetch_document(conn, document_id)
    if not doc:
        raise ApiError(404, "문서를 찾을 수 없습니다.")
    if ("is_deleted" in doc.keys()) and doc["is_deleted"]:
        raise ApiError(409, "보관삭제된 문서는 수정완료 저장할 수 없습니다.")
    if (doc["status"] or "") not in {"approved", "rejected"}:
        raise ApiError(409, "완료 문서에서만 수정완료 저장할 수 있습니다.")
    req_status = (doc["edit_request_status"] if "edit_request_status" in doc.keys() else "none") or "none"
    req_user_id = doc["edit_request_requested_by"] if "edit_request_requested_by" in doc.keys() else None
    if req_status != "approved":
        raise ApiError(409, "수락된 수정요청만 수정완료 저장할 수 있습니다.")
    if req_user_id is None or int(req_user_id) != int(actor["id"]):
        raise ApiError(403, "수정요청 요청자만 수정완료 저장할 수 있습니다.")

    now = now_ts()
    conn.execute(
        "UPDATE documents SET edit_request_status='closed', updated_at=? WHERE id=?",
        (now, document_id),
    )
    reviewer_id = doc["edit_request_reviewer_id"] if "edit_request_reviewer_id" in doc.keys() else None
    if reviewer_id:
        create_notification(conn, int(reviewer_id), f"문서 '{doc['title']}' 수정완료 저장 처리됨", f"/documents/{document_id}")
    notify_admins(conn, f"문서 '{doc['title']}' 수정완료 저장 처리됨", f"/documents/{document_id}", exclude_user_id=int(actor["id"]))
    try:
        ensure_document_print_snapshot(conn, document_id, force=True)
    except Exception as e:
        print(f"[print_snapshot] WARN complete edit request doc={document_id}: {type(e).__name__}: {e}")
    refreshed = fetch_document(conn, document_id)
    if not refreshed:
        raise ApiError(500, "문서를 다시 불러오지 못했습니다.")
    return dict_document(conn, refreshed, include_detail=True)


def parse_leave_days_value(raw_value: Any) -> float | None:
    if raw_value is None or raw_value == "":
        return None
    try:
        value = float(raw_value)
    except (TypeError, ValueError) as exc:
        raise ApiError(400, "휴가 사용일수는 숫자로 입력해 주세요.") from exc
    if value <= 0:
        raise ApiError(400, "휴가 사용일수는 0보다 커야 합니다.")
    return round(value, 2)


def format_leave_number(value: Any) -> str:
    if value is None or value == "":
        return ""
    try:
        num = float(value)
    except (TypeError, ValueError):
        return str(value)
    if abs(num - round(num)) < 1e-9:
        return str(int(round(num)))
    return f"{num:.2f}".rstrip("0").rstrip(".")


def validate_leave_datetime(value: Any, field_name: str) -> str | None:
    if value is None:
        return None
    raw = str(value).strip()
    if not raw:
        return None
    normalized = raw.replace(" ", "T")
    # Accept legacy date-only and new datetime-local values.
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", normalized):
        try:
            date.fromisoformat(normalized)
            return normalized
        except ValueError as exc:
            raise ApiError(400, f"{field_name} 형식이 올바르지 않습니다.") from exc
    try:
        dt = datetime.fromisoformat(normalized)
        return dt.strftime("%Y-%m-%dT%H:%M")
    except ValueError as exc:
        raise ApiError(400, f"{field_name} 형식은 YYYY-MM-DDTHH:MM 이어야 합니다.") from exc


def _parse_leave_bound_datetime(value: str, *, end_bound: bool) -> datetime:
    normalized = str(value).strip().replace(" ", "T")
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", normalized):
        d = date.fromisoformat(normalized)
        return datetime.combine(d, WORK_END_TIME if end_bound else WORK_START_TIME)
    return datetime.fromisoformat(normalized)


def _overlap_seconds(start_a: datetime, end_a: datetime, start_b: datetime, end_b: datetime) -> float:
    start = max(start_a, start_b)
    end = min(end_a, end_b)
    if end <= start:
        return 0.0
    return (end - start).total_seconds()


def calculate_leave_days_from_datetimes(leave_start_value: str, leave_end_value: str) -> float:
    start_dt = _parse_leave_bound_datetime(leave_start_value, end_bound=False)
    end_dt = _parse_leave_bound_datetime(leave_end_value, end_bound=True)
    if end_dt <= start_dt:
        raise ApiError(400, "휴가 종료일시는 시작일시보다 뒤여야 합니다.")

    total_seconds = 0.0
    cur_day = start_dt.date()
    last_day = end_dt.date()
    while cur_day <= last_day:
        morning_start = datetime.combine(cur_day, WORK_START_TIME)
        morning_end = datetime.combine(cur_day, LUNCH_START_TIME)
        afternoon_start = datetime.combine(cur_day, LUNCH_END_TIME)
        afternoon_end = datetime.combine(cur_day, WORK_END_TIME)
        total_seconds += _overlap_seconds(start_dt, end_dt, morning_start, morning_end)
        total_seconds += _overlap_seconds(start_dt, end_dt, afternoon_start, afternoon_end)
        cur_day += timedelta(days=1)

    if total_seconds <= 0:
        raise ApiError(400, "휴가시간이 정근로시간(09:00~18:00, 점심 12:00~13:00 제외)과 겹치지 않습니다.")
    leave_days = total_seconds / 3600.0 / WORK_HOURS_PER_DAY
    return round(leave_days, 2)


def format_leave_period_text(leave_start_value: str | None, leave_end_value: str | None) -> str:
    if not leave_start_value or not leave_end_value:
        return ""
    start_text = str(leave_start_value).replace("T", " ")
    end_text = str(leave_end_value).replace("T", " ")
    return start_text if start_text == end_text else f"{start_text} ~ {end_text}"


def validate_leave_request_fields(template_type: str, leave_start_date: str | None, leave_end_date: str | None, leave_days: float | None) -> tuple[str | None, str | None, float | None]:
    if template_type != LEAVE_TEMPLATE_TYPE:
        return None, None, None
    if not leave_start_date:
        raise ApiError(400, "휴가계 문서는 휴가 시작일시를 입력해 주세요.")
    if not leave_end_date:
        raise ApiError(400, "휴가계 문서는 휴가 종료일시를 입력해 주세요.")
    normalized_start = validate_leave_datetime(leave_start_date, "휴가 시작일시")
    normalized_end = validate_leave_datetime(leave_end_date, "휴가 종료일시")
    if not normalized_start or not normalized_end:
        raise ApiError(400, "휴가 일시를 확인해 주세요.")
    computed_days = calculate_leave_days_from_datetimes(normalized_start, normalized_end)
    return normalized_start, normalized_end, computed_days


def parse_overtime_hours_value(raw_value: Any) -> float | None:
    if raw_value is None or raw_value == "":
        return None
    try:
        value = float(raw_value)
    except (TypeError, ValueError) as exc:
        raise ApiError(400, "연장근로 시간은 숫자로 입력해 주세요.") from exc
    if value <= 0:
        raise ApiError(400, "연장근로 시간은 0보다 커야 합니다.")
    return round(value, 2)


def calculate_overtime_hours_from_datetimes(start_value: str, end_value: str) -> float:
    start_dt = _parse_leave_bound_datetime(start_value, end_bound=False)
    end_dt = _parse_leave_bound_datetime(end_value, end_bound=True)
    if end_dt <= start_dt:
        raise ApiError(400, "연장근로 종료일시는 시작일시보다 뒤여야 합니다.")
    hours = (end_dt - start_dt).total_seconds() / 3600.0
    if hours <= 0:
        raise ApiError(400, "연장근로 시간이 올바르지 않습니다.")
    return round(hours, 2)


def format_overtime_period_text(start_value: str | None, end_value: str | None, hours_value: float | None = None) -> str:
    if not start_value or not end_value:
        return ""
    start_text = str(start_value).replace("T", " ")
    end_text = str(end_value).replace("T", " ")
    base = start_text if start_text == end_text else f"{start_text} ~ {end_text}"
    if hours_value and hours_value > 0:
        return f"{base} ({format_leave_number(hours_value)}시간)"
    return base


def format_trip_period_text(start_value: str | None, end_value: str | None) -> str:
    if not start_value or not end_value:
        return ""
    start_text = str(start_value).replace("T", " ")
    end_text = str(end_value).replace("T", " ")
    return start_text if start_text == end_text else f"{start_text} ~ {end_text}"


def format_education_period_text(start_value: str | None, end_value: str | None) -> str:
    if not start_value and not end_value:
        return ""
    if start_value and end_value:
        start_text = str(start_value).replace("T", " ")
        end_text = str(end_value).replace("T", " ")
        return start_text if start_text == end_text else f"{start_text} ~ {end_text}"
    return str(start_value or end_value or "").replace("T", " ")


def validate_overtime_request_fields(template_type: str, overtime_start_date: str | None, overtime_end_date: str | None, overtime_hours: float | None) -> tuple[str | None, str | None, float | None]:
    if template_type != OVERTIME_TEMPLATE_TYPE:
        return None, None, None
    if not overtime_start_date:
        raise ApiError(400, "연장근로 문서는 연장근로 시작일시를 입력해 주세요.")
    if not overtime_end_date:
        raise ApiError(400, "연장근로 문서는 연장근로 종료일시를 입력해 주세요.")
    normalized_start = validate_leave_datetime(overtime_start_date, "연장근로 시작일시")
    normalized_end = validate_leave_datetime(overtime_end_date, "연장근로 종료일시")
    if not normalized_start or not normalized_end:
        raise ApiError(400, "연장근로 일시를 확인해 주세요.")
    computed_hours = calculate_overtime_hours_from_datetimes(normalized_start, normalized_end)
    return normalized_start, normalized_end, computed_hours


def validate_business_trip_request_fields(
    template_type: str,
    trip_start_date: str | None,
    trip_end_date: str | None,
) -> tuple[str | None, str | None]:
    if template_type != BUSINESS_TRIP_TEMPLATE_TYPE:
        return None, None
    if not trip_start_date:
        raise ApiError(400, "출장신청서 문서는 출장 시작일시를 입력해 주세요.")
    if not trip_end_date:
        raise ApiError(400, "출장신청서 문서는 출장 종료일시를 입력해 주세요.")
    normalized_start = validate_leave_datetime(trip_start_date, "출장 시작일시")
    normalized_end = validate_leave_datetime(trip_end_date, "출장 종료일시")
    if not normalized_start or not normalized_end:
        raise ApiError(400, "출장 일시를 확인해 주세요.")
    start_dt = _parse_leave_bound_datetime(normalized_start, end_bound=False)
    end_dt = _parse_leave_bound_datetime(normalized_end, end_bound=True)
    if end_dt <= start_dt:
        raise ApiError(400, "출장 종료일시는 시작일시보다 뒤여야 합니다.")
    return normalized_start, normalized_end


def validate_education_request_fields(
    template_type: str,
    education_start_date: str | None,
    education_end_date: str | None,
) -> tuple[str | None, str | None]:
    if template_type != EDUCATION_TEMPLATE_TYPE:
        return None, None
    if not education_start_date:
        raise ApiError(400, "교육신청서 문서는 교육 시작일시를 입력해 주세요.")
    if not education_end_date:
        raise ApiError(400, "교육신청서 문서는 교육 종료일시를 입력해 주세요.")
    normalized_start = validate_leave_datetime(education_start_date, "교육 시작일시")
    normalized_end = validate_leave_datetime(education_end_date, "교육 종료일시")
    if not normalized_start or not normalized_end:
        raise ApiError(400, "교육 일시를 확인해 주세요.")
    start_dt = _parse_leave_bound_datetime(normalized_start, end_bound=False)
    end_dt = _parse_leave_bound_datetime(normalized_end, end_bound=True)
    if end_dt <= start_dt:
        raise ApiError(400, "교육 종료일시는 시작일시보다 뒤여야 합니다.")
    return normalized_start, normalized_end


def calculate_trip_hours(start_value: str | None, end_value: str | None) -> float:
    if not start_value or not end_value:
        return 0.0
    start_dt = _parse_leave_bound_datetime(start_value, end_bound=False)
    end_dt = _parse_leave_bound_datetime(end_value, end_bound=True)
    if end_dt <= start_dt:
        return 0.0
    return round((end_dt - start_dt).total_seconds() / 3600.0, 2)


def parse_trip_expense_amount(raw_value: Any) -> float:
    if raw_value is None:
        return 0.0
    text = str(raw_value).strip()
    if not text:
        return 0.0
    # Keep only number-like characters for robust parsing of "10,000원" style values.
    cleaned = re.sub(r"[^\d\.\-]", "", text.replace(",", ""))
    if cleaned in {"", "-", ".", "-."}:
        return 0.0
    try:
        return round(float(cleaned), 2)
    except (TypeError, ValueError):
        return 0.0


def parse_education_money_value(raw_value: Any) -> float | None:
    if raw_value in (None, ""):
        return None
    text = str(raw_value).strip()
    if not text:
        return None
    cleaned = re.sub(r"[^\d\.\-]", "", text.replace(",", ""))
    if cleaned in {"", "-", ".", "-."}:
        return None
    try:
        value = float(cleaned)
    except (TypeError, ValueError) as exc:
        raise ApiError(400, "교육비 항목 금액은 숫자로 입력해 주세요.") from exc
    return round(value, 2)


def calculate_education_hours(start_value: str | None, end_value: str | None) -> float:
    if not start_value or not end_value:
        return 0.0
    start_dt = _parse_leave_bound_datetime(start_value, end_bound=False)
    end_dt = _parse_leave_bound_datetime(end_value, end_bound=True)
    if end_dt <= start_dt:
        return 0.0
    return round((end_dt - start_dt).total_seconds() / 3600.0, 2)


def _adjust_user_used_leave(conn: sqlite3.Connection, user_id: int, delta: float) -> None:
    if abs(delta) < 1e-9:
        return
    row = conn.execute("SELECT used_leave FROM users WHERE id=?", (user_id,)).fetchone()
    if not row:
        return
    current = float(row["used_leave"] or 0)
    new_value = round(max(0.0, current + float(delta)), 2)
    conn.execute("UPDATE users SET used_leave=? WHERE id=?", (new_value, user_id))


def sync_leave_usage_for_document(conn: sqlite3.Connection, document_id: int) -> None:
    doc = fetch_document(conn, document_id)
    if not doc:
        return
    if (doc["template_type"] or "") != LEAVE_TEMPLATE_TYPE:
        return

    leave_start = doc["leave_start_date"] if "leave_start_date" in doc.keys() else None
    leave_end = doc["leave_end_date"] if "leave_end_date" in doc.keys() else None
    leave_days = float(doc["leave_days"] or 0) if "leave_days" in doc.keys() and doc["leave_days"] is not None else 0.0
    row = conn.execute("SELECT * FROM leave_usages WHERE document_id=?", (document_id,)).fetchone()

    should_be_active = (
        not (("is_deleted" in doc.keys()) and int(doc["is_deleted"] or 0))
        and (doc["status"] in ("in_review", "approved"))
        and bool(leave_start and leave_end and leave_days > 0)
    )

    if not should_be_active:
        if row and str(row["status"] or "") == "active":
            _adjust_user_used_leave(conn, int(row["user_id"]), -float(row["leave_days"] or 0))
            conn.execute("UPDATE leave_usages SET status='cancelled', updated_at=? WHERE document_id=?", (now_ts(), document_id))
        return

    if not row:
        conn.execute(
            """
            INSERT INTO leave_usages (document_id, user_id, start_date, end_date, leave_days, status, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, 'active', ?, ?)
            """,
            (document_id, doc["drafter_id"], leave_start, leave_end, leave_days, now_ts(), now_ts()),
        )
        _adjust_user_used_leave(conn, int(doc["drafter_id"]), leave_days)
        return

    prev_status = str(row["status"] or "cancelled")
    prev_days = float(row["leave_days"] or 0)
    conn.execute(
        "UPDATE leave_usages SET user_id=?, start_date=?, end_date=?, leave_days=?, status='active', updated_at=? WHERE document_id=?",
        (doc["drafter_id"], leave_start, leave_end, leave_days, now_ts(), document_id),
    )
    if prev_status != "active":
        _adjust_user_used_leave(conn, int(doc["drafter_id"]), leave_days)
    else:
        delta = round(leave_days - prev_days, 2)
        if abs(delta) > 1e-9:
            _adjust_user_used_leave(conn, int(doc["drafter_id"]), delta)


def list_leave_usages(conn: sqlite3.Connection, viewer: sqlite3.Row) -> dict[str, Any]:
    me = conn.execute("SELECT total_leave, used_leave FROM users WHERE id=?", (viewer["id"],)).fetchone()
    rows = conn.execute(
        """
        SELECT lu.*, d.title document_title, d.status document_status, d.leave_type document_leave_type, COALESCE(d.is_deleted,0) is_deleted,
               u.full_name user_name, u.department user_department
        FROM leave_usages lu
        JOIN documents d ON d.id = lu.document_id
        JOIN users u ON u.id = lu.user_id
        WHERE COALESCE(d.is_deleted,0)=0
          AND lu.status='active'
        ORDER BY lu.updated_at DESC, lu.id DESC
        LIMIT 500
        """
    ).fetchall()
    records = [
        {
            "id": r["id"],
            "document_id": r["document_id"],
            "document_title": r["document_title"],
            "document_status": r["document_status"],
            "leave_type": (r["document_leave_type"] if "document_leave_type" in r.keys() else None),
            "user": {"id": r["user_id"], "name": r["user_name"], "department": r["user_department"]},
            "start_date": r["start_date"],
            "end_date": r["end_date"],
            "leave_days": float(r["leave_days"] or 0),
            "status": r["status"],
            "updated_at": r["updated_at"],
        }
        for r in rows
    ]
    return {
        "summary": {
            "total_leave": float(me["total_leave"] or 0) if me else 0.0,
            "used_leave": float(me["used_leave"] or 0) if me else 0.0,
        },
        "my_records": [r for r in records if r["user"]["id"] == viewer["id"]],
        "other_records": [r for r in records if r["user"]["id"] != viewer["id"]],
    }


def list_overtime_usages(
    conn: sqlite3.Connection,
    viewer: sqlite3.Row,
    selected_year: int | None = None,
    selected_month: int | None = None,
) -> dict[str, Any]:
    now_dt = datetime.now()
    year = int(selected_year or now_dt.year)
    month = int(selected_month or now_dt.month)
    if year < 2000 or year > 2100:
        year = now_dt.year
    if month < 1 or month > 12:
        month = now_dt.month

    overtime_rows = conn.execute(
        """
        SELECT d.id, d.title, d.status, d.template_type, d.overtime_type, d.overtime_start_date, d.overtime_end_date,
               d.overtime_hours, d.updated_at, d.created_at,
               u.full_name user_name, u.department user_department
        FROM documents d
        JOIN users u ON u.id = d.drafter_id
        WHERE d.drafter_id = ?
          AND d.template_type = ?
          AND d.status = 'approved'
          AND COALESCE(d.is_deleted,0)=0
        ORDER BY COALESCE(d.overtime_start_date, d.updated_at, d.created_at) DESC, d.id DESC
        LIMIT 1000
        """,
        (viewer["id"], OVERTIME_TEMPLATE_TYPE),
    ).fetchall()

    leave_rows = conn.execute(
        """
        SELECT d.leave_start_date, d.leave_end_date
        FROM documents d
        WHERE d.drafter_id = ?
          AND d.template_type = ?
          AND d.status = 'approved'
          AND COALESCE(d.is_deleted,0)=0
          AND d.leave_start_date IS NOT NULL
          AND d.leave_end_date IS NOT NULL
        """,
        (viewer["id"], LEAVE_TEMPLATE_TYPE),
    ).fetchall()

    def _parse_iso_dt(value: Any) -> datetime | None:
        raw = str(value or "").strip()
        if not raw:
            return None
        try:
            return datetime.fromisoformat(raw)
        except Exception:
            return None

    def _calc_business_overlap_hours(start_dt: datetime, end_dt: datetime) -> float:
        if end_dt <= start_dt:
            return 0.0
        total_seconds = 0.0
        cur_day = start_dt.date()
        last_day = end_dt.date()
        while cur_day <= last_day:
            morning_start = datetime.combine(cur_day, WORK_START_TIME)
            morning_end = datetime.combine(cur_day, LUNCH_START_TIME)
            afternoon_start = datetime.combine(cur_day, LUNCH_END_TIME)
            afternoon_end = datetime.combine(cur_day, WORK_END_TIME)
            total_seconds += _overlap_seconds(start_dt, end_dt, morning_start, morning_end)
            total_seconds += _overlap_seconds(start_dt, end_dt, afternoon_start, afternoon_end)
            cur_day += timedelta(days=1)
        return round(max(0.0, total_seconds / 3600.0), 4)

    def _split_interval_by_day(start_dt: datetime, end_dt: datetime) -> list[tuple[datetime, datetime, float]]:
        if end_dt <= start_dt:
            return []
        out: list[tuple[datetime, datetime, float]] = []
        cursor = start_dt
        while cursor < end_dt:
            next_day = datetime.combine(cursor.date() + timedelta(days=1), time(0, 0))
            seg_end = min(end_dt, next_day)
            hours = (seg_end - cursor).total_seconds() / 3600.0
            if hours > 0:
                out.append((cursor, seg_end, round(hours, 4)))
            cursor = seg_end
        return out

    def _week_start_monday(d: date) -> date:
        return d - timedelta(days=d.weekday())

    leave_hours_by_day: dict[str, float] = {}
    for lr in leave_rows:
        sdt = _parse_iso_dt(lr["leave_start_date"])
        edt = _parse_iso_dt(lr["leave_end_date"])
        if not sdt or not edt or edt <= sdt:
            continue
        cur_day = sdt.date()
        last_day = edt.date()
        while cur_day <= last_day:
            day_start = datetime.combine(cur_day, time(0, 0))
            day_end = datetime.combine(cur_day + timedelta(days=1), time(0, 0))
            overlap_h = _calc_business_overlap_hours(max(sdt, day_start), min(edt, day_end))
            if overlap_h > 0:
                key = cur_day.isoformat()
                leave_hours_by_day[key] = round(leave_hours_by_day.get(key, 0.0) + overlap_h, 4)
            cur_day += timedelta(days=1)

    raw_records: list[dict[str, Any]] = []
    available_years: set[int] = set()
    for r in overtime_rows:
        start_raw = (r["overtime_start_date"] or "").strip()
        end_raw = (r["overtime_end_date"] or "").strip()
        start_dt = _parse_iso_dt(start_raw)
        end_dt = _parse_iso_dt(end_raw)
        if not start_dt or not end_dt or end_dt <= start_dt:
            continue
        rec_year = start_dt.year
        rec_month = start_dt.month
        available_years.add(rec_year)
        hours_val = float(r["overtime_hours"] or 0)
        raw_records.append({
            "document_id": r["id"],
            "document_title": r["title"],
            "document_status": r["status"],
            "user": {"id": int(viewer["id"]), "name": r["user_name"], "department": r["user_department"]},
            "overtime_type": r["overtime_type"],
            "start_date": start_raw or None,
            "end_date": end_raw or None,
            "overtime_hours": hours_val,
            "updated_at": r["updated_at"],
            "_start_dt": start_dt,
            "_end_dt": end_dt,
        })

    overtime_segments: list[dict[str, Any]] = []
    for rec in raw_records:
        for seg_start, seg_end, seg_hours in _split_interval_by_day(rec["_start_dt"], rec["_end_dt"]):
            overtime_segments.append({
                "document_id": rec["document_id"],
                "start_dt": seg_start,
                "end_dt": seg_end,
                "hours": seg_hours,
            })
    overtime_segments.sort(key=lambda s: (s["start_dt"], s["document_id"]))

    weekly_actual_hours: dict[str, float] = {}
    daily_leave_100_remaining: dict[str, float] = {}
    monthly_hours: dict[tuple[int, int], float] = {}
    monthly_hours_100: dict[tuple[int, int], float] = {}
    monthly_hours_150: dict[tuple[int, int], float] = {}
    monthly_counts: dict[tuple[int, int], int] = {}
    daily_hours: dict[str, float] = {}
    daily_hours_100: dict[str, float] = {}
    daily_hours_150: dict[str, float] = {}
    daily_counts: dict[str, int] = {}
    seen_doc_days: set[tuple[int, str]] = set()
    seen_doc_months: set[tuple[int, int, int]] = set()
    record_split_map: dict[int, dict[str, float]] = {}

    def _get_week_regular_base_hours(monday: date) -> float:
        week_key = monday.isoformat()
        if week_key in weekly_actual_hours:
            return weekly_actual_hours[week_key]
        total = 0.0
        for i in range(5):  # Mon-Fri only
            d = monday + timedelta(days=i)
            leave_h = min(float(leave_hours_by_day.get(d.isoformat(), 0.0)), WORK_HOURS_PER_DAY)
            total += max(0.0, WORK_HOURS_PER_DAY - leave_h)
        weekly_actual_hours[week_key] = round(total, 4)
        return weekly_actual_hours[week_key]

    for seg in overtime_segments:
        seg_start = seg["start_dt"]
        seg_hours = float(seg["hours"] or 0)
        if seg_hours <= 0:
            continue
        day_key = seg_start.date().isoformat()
        week_monday = _week_start_monday(seg_start.date())
        week_key = week_monday.isoformat()
        _get_week_regular_base_hours(week_monday)

        alloc_100 = 0.0
        alloc_150 = 0.0
        weekday = seg_start.weekday()  # Mon=0
        if weekday <= 4:
            if day_key not in daily_leave_100_remaining:
                daily_leave_100_remaining[day_key] = round(min(float(leave_hours_by_day.get(day_key, 0.0)), WORK_HOURS_PER_DAY), 4)
            remain_100 = max(0.0, float(daily_leave_100_remaining.get(day_key, 0.0)))
            alloc_100 = min(seg_hours, remain_100)
            alloc_150 = max(0.0, seg_hours - alloc_100)
            daily_leave_100_remaining[day_key] = round(max(0.0, remain_100 - alloc_100), 4)
        else:
            before_actual = float(weekly_actual_hours.get(week_key, 0.0))
            remain_to_40 = max(0.0, 40.0 - before_actual)
            alloc_100 = min(seg_hours, remain_to_40)
            alloc_150 = max(0.0, seg_hours - alloc_100)

        weekly_actual_hours[week_key] = round(float(weekly_actual_hours.get(week_key, 0.0)) + seg_hours, 4)

        month_key = (seg_start.year, seg_start.month)
        monthly_hours[month_key] = round(monthly_hours.get(month_key, 0.0) + seg_hours, 4)
        monthly_hours_100[month_key] = round(monthly_hours_100.get(month_key, 0.0) + alloc_100, 4)
        monthly_hours_150[month_key] = round(monthly_hours_150.get(month_key, 0.0) + alloc_150, 4)
        day_iso = seg_start.date().isoformat()
        daily_hours[day_iso] = round(daily_hours.get(day_iso, 0.0) + seg_hours, 4)
        daily_hours_100[day_iso] = round(daily_hours_100.get(day_iso, 0.0) + alloc_100, 4)
        daily_hours_150[day_iso] = round(daily_hours_150.get(day_iso, 0.0) + alloc_150, 4)
        doc_day_key = (seg["document_id"], day_iso)
        if doc_day_key not in seen_doc_days:
            daily_counts[day_iso] = daily_counts.get(day_iso, 0) + 1
            seen_doc_days.add(doc_day_key)
        record_split = record_split_map.setdefault(seg["document_id"], {"hours_100": 0.0, "hours_150": 0.0})
        record_split["hours_100"] = round(record_split["hours_100"] + alloc_100, 4)
        record_split["hours_150"] = round(record_split["hours_150"] + alloc_150, 4)

    records: list[dict[str, Any]] = []
    for rec in raw_records:
        start_dt = rec["_start_dt"]
        rec_month_key = (start_dt.year, start_dt.month)
        doc_month_key = (rec["document_id"], start_dt.year, start_dt.month)
        if doc_month_key not in seen_doc_months:
            monthly_counts[rec_month_key] = monthly_counts.get(rec_month_key, 0) + 1
            seen_doc_months.add(doc_month_key)
        split = record_split_map.get(rec["document_id"], {"hours_100": 0.0, "hours_150": 0.0})
        rec["overtime_100_hours"] = round(split.get("hours_100", 0.0), 2)
        rec["overtime_150_hours"] = round(split.get("hours_150", 0.0), 2)
        rec.pop("_start_dt", None)
        rec.pop("_end_dt", None)
        records.append(rec)

    year_monthly = []
    for m in range(1, 13):
        key = (year, m)
        hours_val = round(monthly_hours.get(key, 0.0), 2)
        year_monthly.append({
            "month": m,
            "hours": hours_val,
            "hours_100": round(monthly_hours_100.get(key, 0.0), 2),
            "hours_150": round(monthly_hours_150.get(key, 0.0), 2),
            "count": monthly_counts.get(key, 0),
        })

    month_records = []
    for rec in records:
        start_raw = rec["start_date"] or ""
        try:
            start_dt = datetime.fromisoformat(start_raw) if start_raw else None
        except Exception:
            start_dt = None
        if not start_dt:
            continue
        if start_dt.year == year and start_dt.month == month:
            month_records.append(rec)

    month_records.sort(key=lambda r: (r.get("start_date") or "", r.get("document_id") or 0), reverse=True)
    month_used = round(sum(float(r.get("overtime_hours") or 0) for r in month_records), 2)
    month_used_100 = round(sum(float(r.get("overtime_100_hours") or 0) for r in month_records), 2)
    month_used_150 = round(sum(float(r.get("overtime_150_hours") or 0) for r in month_records), 2)
    monthly_cap = 15.0

    if not available_years:
        available_years.add(year)

    month_start = date(year, month, 1)
    if month == 12:
        next_month_start = date(year + 1, 1, 1)
    else:
        next_month_start = date(year, month + 1, 1)
    day_count = (next_month_start - month_start).days
    month_daily = []
    for d in range(1, day_count + 1):
        day_dt = date(year, month, d)
        key = day_dt.isoformat()
        hours_val = round(daily_hours.get(key, 0.0), 2)
        month_daily.append({
            "day": d,
            "date": key,
            "hours": hours_val,
            "hours_100": round(daily_hours_100.get(key, 0.0), 2),
            "hours_150": round(daily_hours_150.get(key, 0.0), 2),
            "count": daily_counts.get(key, 0),
        })

    return {
        "selection": {"year": year, "month": month},
        "summary": {
            "monthly_cap_hours": monthly_cap,
            "month_used_hours": month_used,
            "month_remaining_hours": round(max(0.0, monthly_cap - month_used), 2),
            "month_used_100_hours": month_used_100,
            "month_used_150_hours": month_used_150,
            "month_weighted_hours": round(month_used_100 + (month_used_150 * 1.5), 2),
            "year_total_hours": round(sum(item["hours"] for item in year_monthly), 2),
        },
        "available_years": sorted(available_years, reverse=True),
        "year_monthly": year_monthly,
        "month_daily": month_daily,
        "month_records": month_records,
    }


def list_business_trip_usages(
    conn: sqlite3.Connection,
    viewer: sqlite3.Row,
    selected_year: int | None = None,
    selected_month: int | None = None,
) -> dict[str, Any]:
    # Defensive migration guard: keeps API working even if server started from an older DB schema.
    ensure_document_columns(conn)

    now_dt = datetime.now()
    year = int(selected_year or now_dt.year)
    month = int(selected_month or now_dt.month)
    if year < 2000 or year > 2100:
        year = now_dt.year
    if month < 1 or month > 12:
        month = now_dt.month

    try:
        rows = conn.execute(
            """
            SELECT d.id, d.title, d.status, d.trip_type, d.trip_destination, d.trip_start_date, d.trip_end_date,
                   d.trip_transportation, d.trip_expense, d.trip_purpose,
                   d.updated_at, d.created_at, u.full_name user_name, u.department user_department,
                   (
                     SELECT COUNT(1)
                     FROM documents r
                     WHERE r.source_trip_document_id=d.id
                       AND COALESCE(r.is_deleted,0)=0
                   ) result_doc_count,
                   (
                     SELECT r.id
                     FROM documents r
                     WHERE r.source_trip_document_id=d.id
                       AND COALESCE(r.is_deleted,0)=0
                     ORDER BY COALESCE(r.updated_at, r.created_at) DESC, r.id DESC
                     LIMIT 1
                   ) latest_result_doc_id,
                   (
                     SELECT r.status
                     FROM documents r
                     WHERE r.source_trip_document_id=d.id
                       AND COALESCE(r.is_deleted,0)=0
                     ORDER BY COALESCE(r.updated_at, r.created_at) DESC, r.id DESC
                     LIMIT 1
                   ) latest_result_doc_status,
                   (
                     SELECT r.title
                     FROM documents r
                     WHERE r.source_trip_document_id=d.id
                       AND COALESCE(r.is_deleted,0)=0
                     ORDER BY COALESCE(r.updated_at, r.created_at) DESC, r.id DESC
                     LIMIT 1
                   ) latest_result_doc_title,
                   (
                     SELECT COALESCE(r.submitted_at, r.updated_at, r.created_at)
                     FROM documents r
                     WHERE r.source_trip_document_id=d.id
                       AND COALESCE(r.is_deleted,0)=0
                     ORDER BY COALESCE(r.updated_at, r.created_at) DESC, r.id DESC
                     LIMIT 1
                   ) latest_result_doc_time
            FROM documents d
            JOIN users u ON u.id = d.drafter_id
            WHERE d.drafter_id = ?
              AND d.template_type = ?
              AND d.status IN ('in_review', 'approved', 'rejected')
              AND COALESCE(d.is_deleted,0)=0
            ORDER BY COALESCE(d.trip_start_date, d.updated_at, d.created_at) DESC, d.id DESC
            LIMIT 1000
            """,
            (viewer["id"], BUSINESS_TRIP_TEMPLATE_TYPE),
        ).fetchall()
    except sqlite3.OperationalError as e:
        # Return empty payload with warning instead of failing entire dashboard render.
        return {
            "selection": {"year": year, "month": month},
            "summary": {
                "month_count": 0,
                "month_total_hours": 0,
                "month_avg_hours": 0,
                "month_total_expense": 0,
                "month_avg_expense": 0,
            },
            "status_summary": [{"key": "in_review", "count": 0}, {"key": "approved", "count": 0}, {"key": "rejected", "count": 0}],
            "progress_summary": [{"key": "scheduled", "count": 0}, {"key": "ongoing", "count": 0}, {"key": "finished", "count": 0}, {"key": "rejected", "count": 0}],
            "type_summary": [],
            "month_records": [],
            "result_history": [],
            "warning": f"출장현황 조회 중 스키마 오류: {e}",
        }

    records: list[dict[str, Any]] = []
    type_counter: dict[str, int] = {}
    status_counter: dict[str, int] = {"in_review": 0, "approved": 0, "rejected": 0}
    progress_counter: dict[str, int] = {"scheduled": 0, "ongoing": 0, "finished": 0, "rejected": 0}
    now_local = datetime.now()
    for r in rows:
        start_raw = str(r["trip_start_date"] or "").strip()
        end_raw = str(r["trip_end_date"] or "").strip()
        try:
            start_dt = datetime.fromisoformat(start_raw) if start_raw else None
        except Exception:
            start_dt = None
        if not start_dt:
            try:
                start_dt = datetime.fromisoformat(str(r["created_at"]).replace(" ", "T"))
            except Exception:
                start_dt = None
        if not start_dt:
            continue
        if start_dt.year != year or start_dt.month != month:
            continue
        status_value = str(r["status"] or "").strip()
        if status_value in status_counter:
            status_counter[status_value] += 1
        trip_type = str(r["trip_type"] or "").strip() or "기타"
        type_counter[trip_type] = type_counter.get(trip_type, 0) + 1
        end_dt: datetime | None = None
        if end_raw:
            try:
                end_dt = datetime.fromisoformat(end_raw)
            except Exception:
                end_dt = None
        if not end_dt:
            end_dt = start_dt

        progress_token = "finished"
        progress_label = "종료"
        if status_value == "rejected":
            progress_token = "rejected"
            progress_label = "반려"
        elif now_local < start_dt:
            progress_token = "scheduled"
            progress_label = "예정"
        elif start_dt <= now_local <= end_dt:
            progress_token = "ongoing"
            progress_label = "진행중"
        progress_counter[progress_token] = progress_counter.get(progress_token, 0) + 1

        trip_hours = calculate_trip_hours(r["trip_start_date"], r["trip_end_date"])
        expense_amount = parse_trip_expense_amount(r["trip_expense"])
        records.append(
            {
                "document_id": r["id"],
                "document_title": r["title"],
                "document_status": r["status"],
                "trip_type": r["trip_type"],
                "trip_destination": r["trip_destination"],
                "trip_start_date": r["trip_start_date"],
                "trip_end_date": r["trip_end_date"],
                "trip_transportation": r["trip_transportation"],
                "trip_expense": r["trip_expense"],
                "trip_expense_amount": expense_amount,
                "trip_purpose": r["trip_purpose"],
                "trip_hours": trip_hours,
                "result_doc_count": int(r["result_doc_count"] or 0) if "result_doc_count" in r.keys() else 0,
                "latest_result_doc_id": int(r["latest_result_doc_id"] or 0) if "latest_result_doc_id" in r.keys() and r["latest_result_doc_id"] else None,
                "latest_result_doc_status": str(r["latest_result_doc_status"] or "").strip() if "latest_result_doc_status" in r.keys() else "",
                "latest_result_doc_title": str(r["latest_result_doc_title"] or "").strip() if "latest_result_doc_title" in r.keys() else "",
                "latest_result_doc_time": str(r["latest_result_doc_time"] or "").strip() if "latest_result_doc_time" in r.keys() else "",
                "progress_token": progress_token,
                "progress_label": progress_label,
                "updated_at": r["updated_at"],
                "user": {"id": int(viewer["id"]), "name": r["user_name"], "department": r["user_department"]},
            }
        )

    records.sort(key=lambda rec: str(rec.get("trip_start_date") or ""), reverse=True)
    month_total_hours = round(sum(float(r.get("trip_hours") or 0) for r in records), 2)
    month_total_expense = round(sum(float(r.get("trip_expense_amount") or 0) for r in records), 2)
    type_summary = [
        {"type": t, "count": c}
        for t, c in sorted(type_counter.items(), key=lambda kv: (-kv[1], kv[0]))
    ]
    status_summary = [
        {"key": key, "count": int(status_counter.get(key, 0))}
        for key in ("in_review", "approved", "rejected")
    ]
    progress_summary = [
        {"key": key, "count": int(progress_counter.get(key, 0))}
        for key in ("scheduled", "ongoing", "finished", "rejected")
    ]
    result_history = [
        {
            "source_document_id": rec.get("document_id"),
            "source_document_title": rec.get("document_title"),
            "trip_start_date": rec.get("trip_start_date"),
            "trip_end_date": rec.get("trip_end_date"),
            "result_doc_count": rec.get("result_doc_count", 0),
            "latest_result_doc_id": rec.get("latest_result_doc_id"),
            "latest_result_doc_status": rec.get("latest_result_doc_status"),
            "latest_result_doc_title": rec.get("latest_result_doc_title"),
            "latest_result_doc_time": rec.get("latest_result_doc_time"),
        }
        for rec in records
        if int(rec.get("result_doc_count") or 0) > 0
    ]
    result_history.sort(key=lambda item: str(item.get("latest_result_doc_time") or ""), reverse=True)
    month_count = len(records)
    return {
        "selection": {"year": year, "month": month},
        "summary": {
            "month_count": month_count,
            "month_total_hours": month_total_hours,
            "month_avg_hours": round((month_total_hours / month_count), 2) if month_count else 0.0,
            "month_total_expense": month_total_expense,
            "month_avg_expense": round((month_total_expense / month_count), 2) if month_count else 0.0,
        },
        "status_summary": status_summary,
        "progress_summary": progress_summary,
        "type_summary": type_summary,
        "month_records": records,
        "result_history": result_history,
    }


def list_education_usages(
    conn: sqlite3.Connection,
    viewer: sqlite3.Row,
    selected_year: int | None = None,
    selected_month: int | None = None,
) -> dict[str, Any]:
    ensure_document_columns(conn)

    now_dt = datetime.now()
    year = int(selected_year or now_dt.year)
    month = int(selected_month or now_dt.month)
    if year < 2000 or year > 2100:
        year = now_dt.year
    if month < 1 or month > 12:
        month = now_dt.month

    rows = conn.execute(
        """
        SELECT d.id, d.title, d.status, d.updated_at, d.created_at,
               d.education_department, d.education_job_title, d.education_name,
               d.education_title, d.education_category, d.education_provider, d.education_location,
               d.education_start_date, d.education_end_date, d.education_purpose,
               d.education_tuition_detail, d.education_tuition_amount,
               d.education_material_detail, d.education_material_amount,
               d.education_transport_detail, d.education_transport_amount,
               d.education_other_detail, d.education_other_amount,
               d.education_budget_subject, d.education_funding_source, d.education_payment_method,
               d.education_support_budget, d.education_used_budget, d.education_remain_budget,
               d.education_companion, d.education_ordered, d.education_suggestion,
               u.full_name user_name, u.department user_department,
               (
                 SELECT COUNT(1)
                 FROM documents r
                 WHERE r.source_trip_document_id=d.id
                   AND r.template_type=?
                   AND COALESCE(r.is_deleted,0)=0
               ) result_doc_count,
               (
                 SELECT r.id
                 FROM documents r
                 WHERE r.source_trip_document_id=d.id
                   AND r.template_type=?
                   AND COALESCE(r.is_deleted,0)=0
                 ORDER BY COALESCE(r.updated_at, r.created_at) DESC, r.id DESC
                 LIMIT 1
               ) latest_result_doc_id,
               (
                 SELECT r.status
                 FROM documents r
                 WHERE r.source_trip_document_id=d.id
                   AND r.template_type=?
                   AND COALESCE(r.is_deleted,0)=0
                 ORDER BY COALESCE(r.updated_at, r.created_at) DESC, r.id DESC
                 LIMIT 1
               ) latest_result_doc_status,
               (
                 SELECT r.title
                 FROM documents r
                 WHERE r.source_trip_document_id=d.id
                   AND r.template_type=?
                   AND COALESCE(r.is_deleted,0)=0
                 ORDER BY COALESCE(r.updated_at, r.created_at) DESC, r.id DESC
                 LIMIT 1
               ) latest_result_doc_title,
               (
                 SELECT COALESCE(r.submitted_at, r.updated_at, r.created_at)
                 FROM documents r
                 WHERE r.source_trip_document_id=d.id
                   AND r.template_type=?
                   AND COALESCE(r.is_deleted,0)=0
                 ORDER BY COALESCE(r.updated_at, r.created_at) DESC, r.id DESC
                 LIMIT 1
               ) latest_result_doc_time
        FROM documents d
        JOIN users u ON u.id = d.drafter_id
        WHERE d.drafter_id = ?
          AND d.template_type = ?
          AND d.status IN ('in_review', 'approved', 'rejected')
          AND COALESCE(d.is_deleted,0)=0
        ORDER BY COALESCE(d.education_start_date, d.updated_at, d.created_at) DESC, d.id DESC
        LIMIT 1000
        """,
        (
            EDUCATION_RESULT_TEMPLATE_TYPE,
            EDUCATION_RESULT_TEMPLATE_TYPE,
            EDUCATION_RESULT_TEMPLATE_TYPE,
            EDUCATION_RESULT_TEMPLATE_TYPE,
            EDUCATION_RESULT_TEMPLATE_TYPE,
            viewer["id"],
            EDUCATION_TEMPLATE_TYPE,
        ),
    ).fetchall()

    records: list[dict[str, Any]] = []
    status_counter: dict[str, int] = {"in_review": 0, "approved": 0, "rejected": 0}
    category_counter: dict[str, int] = {}
    available_years: set[int] = set()
    for r in rows:
        start_raw = str(r["education_start_date"] or "").strip()
        base_raw = start_raw or str(r["updated_at"] or r["created_at"] or "").strip()
        base_dt: datetime | None = None
        if base_raw:
            try:
                base_dt = datetime.fromisoformat(base_raw.replace(" ", "T"))
            except Exception:
                base_dt = None
        if not base_dt:
            continue
        available_years.add(base_dt.year)
        if base_dt.year != year or base_dt.month != month:
            continue

        status_value = str(r["status"] or "").strip()
        if status_value in status_counter:
            status_counter[status_value] += 1

        category = str(r["education_category"] or "").strip() or "기타"
        category_counter[category] = category_counter.get(category, 0) + 1

        tuition_amount = float(r["education_tuition_amount"] or 0)
        material_amount = float(r["education_material_amount"] or 0)
        transport_amount = float(r["education_transport_amount"] or 0)
        other_amount = float(r["education_other_amount"] or 0)
        total_amount = round(tuition_amount + material_amount + transport_amount + other_amount, 2)

        records.append(
            {
                "document_id": int(r["id"]),
                "document_title": r["title"],
                "document_status": status_value,
                "education_department": r["education_department"],
                "education_job_title": r["education_job_title"],
                "education_name": r["education_name"],
                "education_title": r["education_title"],
                "education_category": r["education_category"],
                "education_provider": r["education_provider"],
                "education_location": r["education_location"],
                "education_start_date": r["education_start_date"],
                "education_end_date": r["education_end_date"],
                "education_period": format_education_period_text(r["education_start_date"], r["education_end_date"]),
                "education_hours": calculate_education_hours(r["education_start_date"], r["education_end_date"]),
                "education_purpose": r["education_purpose"],
                "education_tuition_detail": r["education_tuition_detail"],
                "education_tuition_amount": tuition_amount,
                "education_material_detail": r["education_material_detail"],
                "education_material_amount": material_amount,
                "education_transport_detail": r["education_transport_detail"],
                "education_transport_amount": transport_amount,
                "education_other_detail": r["education_other_detail"],
                "education_other_amount": other_amount,
                "education_total_amount": total_amount,
                "education_budget_subject": r["education_budget_subject"],
                "education_funding_source": r["education_funding_source"],
                "education_payment_method": r["education_payment_method"],
                "education_support_budget": float(r["education_support_budget"] or 0),
                "education_used_budget": float(r["education_used_budget"] or 0),
                "education_remain_budget": float(r["education_remain_budget"] or 0),
                "education_companion": r["education_companion"],
                "education_ordered": r["education_ordered"],
                "education_suggestion": r["education_suggestion"],
                "result_doc_count": int(r["result_doc_count"] or 0) if "result_doc_count" in r.keys() else 0,
                "latest_result_doc_id": int(r["latest_result_doc_id"] or 0) if "latest_result_doc_id" in r.keys() and r["latest_result_doc_id"] else None,
                "latest_result_doc_status": str(r["latest_result_doc_status"] or "").strip() if "latest_result_doc_status" in r.keys() else "",
                "latest_result_doc_title": str(r["latest_result_doc_title"] or "").strip() if "latest_result_doc_title" in r.keys() else "",
                "latest_result_doc_time": str(r["latest_result_doc_time"] or "").strip() if "latest_result_doc_time" in r.keys() else "",
                "updated_at": r["updated_at"],
                "user": {"id": int(viewer["id"]), "name": r["user_name"], "department": r["user_department"]},
            }
        )

    records.sort(key=lambda rec: str(rec.get("education_start_date") or rec.get("updated_at") or ""), reverse=True)
    month_count = len(records)
    month_total_hours = round(sum(float(item.get("education_hours") or 0) for item in records), 2)
    month_total_amount = round(sum(float(item.get("education_total_amount") or 0) for item in records), 2)
    if not available_years:
        available_years.add(year)
    result_history = [
        {
            "source_document_id": rec.get("document_id"),
            "source_document_title": rec.get("document_title"),
            "education_start_date": rec.get("education_start_date"),
            "education_end_date": rec.get("education_end_date"),
            "result_doc_count": rec.get("result_doc_count", 0),
            "latest_result_doc_id": rec.get("latest_result_doc_id"),
            "latest_result_doc_status": rec.get("latest_result_doc_status"),
            "latest_result_doc_title": rec.get("latest_result_doc_title"),
            "latest_result_doc_time": rec.get("latest_result_doc_time"),
        }
        for rec in records
        if int(rec.get("result_doc_count") or 0) > 0
    ]
    result_history.sort(key=lambda item: str(item.get("latest_result_doc_time") or ""), reverse=True)

    return {
        "selection": {"year": year, "month": month},
        "summary": {
            "month_count": month_count,
            "month_total_hours": month_total_hours,
            "month_avg_hours": round((month_total_hours / month_count), 2) if month_count else 0.0,
            "month_total_amount": month_total_amount,
            "month_avg_amount": round((month_total_amount / month_count), 2) if month_count else 0.0,
        },
        "status_summary": [
            {"key": key, "count": int(status_counter.get(key, 0))}
            for key in ("in_review", "approved", "rejected")
        ],
        "category_summary": [
            {"category": k, "count": v}
            for k, v in sorted(category_counter.items(), key=lambda kv: (-kv[1], kv[0]))
        ],
        "available_years": sorted(available_years, reverse=True),
        "month_records": records,
        "result_history": result_history,
    }


def submit_business_trip_result_document(
    conn: sqlite3.Connection,
    *,
    source_document_id: int,
    actor: sqlite3.Row,
    approver_ids: list[int],
    reference_ids: list[int] | None = None,
    trip_result_text: str,
) -> tuple[dict[str, Any], list[str]]:
    source = fetch_document(conn, source_document_id)
    if not source:
        raise ApiError(404, "출장 원문서를 찾을 수 없습니다.")
    if (source["template_type"] or "") != BUSINESS_TRIP_TEMPLATE_TYPE:
        raise ApiError(400, "출장신청서 문서에서만 출장결과 상신이 가능합니다.")
    if int(source["drafter_id"]) != int(actor["id"]):
        raise ApiError(403, "출장결과는 원문 기안자만 상신할 수 있습니다.")
    if "is_deleted" in source.keys() and int(source["is_deleted"] or 0):
        raise ApiError(409, "보관삭제된 문서는 출장결과 상신이 불가능합니다.")

    result_text = str(trip_result_text or "").strip()
    if not result_text:
        raise ApiError(400, "출장결과 내용을 입력해 주세요.")

    normalized_approvers: list[int] = []
    for uid in approver_ids or []:
        try:
            iid = int(uid)
        except (TypeError, ValueError) as exc:
            raise ApiError(400, "출장결과 결재선 사용자 ID가 잘못되었습니다.") from exc
        if iid not in normalized_approvers:
            normalized_approvers.append(iid)
    if not normalized_approvers:
        raise ApiError(400, "출장결과 결재선을 최소 1명 이상 지정해 주세요.")
    max_steps_for_template = max_approver_steps_for_type(BUSINESS_TRIP_RESULT_TEMPLATE_TYPE)
    if len(normalized_approvers) > max_steps_for_template:
        raise ApiError(
            400,
            f"출장결과 결재선은 최대 {max_steps_for_template}명까지 지정할 수 있습니다.",
        )
    if int(actor["id"]) in normalized_approvers:
        raise ApiError(400, "본인은 출장결과 결재자로 지정할 수 없습니다.")
    existing_approvers = {
        int(r["id"])
        for r in conn.execute(
            "SELECT id FROM users WHERE id IN ({})".format(",".join("?" for _ in normalized_approvers)),
            tuple(normalized_approvers),
        ).fetchall()
    }
    missing_approvers = [uid for uid in normalized_approvers if uid not in existing_approvers]
    if missing_approvers:
        raise ApiError(400, f"존재하지 않는 출장결과 결재자 ID가 있습니다: {missing_approvers}")

    normalized_refs: list[int] = []
    for rid in reference_ids or []:
        try:
            iid = int(rid)
        except (TypeError, ValueError):
            continue
        if iid not in normalized_refs:
            normalized_refs.append(iid)
    if normalized_refs:
        existing_refs = {
            int(r["id"])
            for r in conn.execute(
                "SELECT id FROM users WHERE id IN ({})".format(",".join("?" for _ in normalized_refs)),
                tuple(normalized_refs),
            ).fetchall()
        }
        normalized_refs = [rid for rid in normalized_refs if rid in existing_refs]

    copy_title = f"{source['title']}_출장보고서_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    copied = copy_google_doc_template_to_folder(
        BUSINESS_TRIP_RESULT_TEMPLATE_DOC_ID,
        BUSINESS_TRIP_RESULT_OUTPUT_FOLDER_ID,
        copy_title,
    )
    external_doc_id = copied["doc_id"]
    urls = build_google_doc_urls(external_doc_id)
    external_doc_url = urls["edit_url"]
    content = f"[Google Docs 문서] {external_doc_url}"

    issue_date_for_doc = today_str()
    default_issue_dept = str(source["issue_department"] or source["trip_department"] or actor["department"] or "기타").strip() or "기타"
    issue_year_raw = str(source["issue_year"] or "").strip()
    effective_issue_year = issue_year_raw if re.fullmatch(r"\d{4}", issue_year_raw) else str(date.fromisoformat(issue_date_for_doc).year)
    issue_code = allocate_department_issue_code(
        conn,
        default_issue_dept,
        issue_date_for_doc,
        effective_issue_year,
    )
    visibility_scope = (source["visibility_scope"] or "private") if "visibility_scope" in source.keys() else "private"
    if visibility_scope not in DOC_VISIBILITY_VALUES:
        visibility_scope = "private"

    now = now_ts()
    trip_start = validate_leave_datetime(source["trip_start_date"], "출장 시작일시")
    trip_end = validate_leave_datetime(source["trip_end_date"], "출장 종료일시")
    trip_period = format_trip_period_text(trip_start, trip_end)
    source_refs: list[int] = []
    if normalized_refs:
        source_refs = normalized_refs
    else:
        try:
            loaded = json.loads(source["reference_ids"] or "[]")
            if isinstance(loaded, list):
                for rid in loaded:
                    try:
                        source_refs.append(int(rid))
                    except (TypeError, ValueError):
                        continue
        except Exception:
            source_refs = []

    cur = conn.execute(
        """
        INSERT INTO documents (title, template_type, content, editor_provider, external_doc_id, external_doc_url, status, priority, due_date, leave_type, leave_start_date, leave_end_date, leave_days, leave_reason, leave_substitute_name, leave_substitute_work, overtime_type, overtime_start_date, overtime_end_date, overtime_hours, overtime_content, overtime_etc, trip_department, trip_job_title, trip_name, trip_type, trip_destination, trip_start_date, trip_end_date, trip_transportation, trip_expense, trip_purpose, trip_result, source_trip_document_id, issue_code, issue_department, issue_year, recipient_text, visibility_scope, attachments_json, drafter_id, submitted_at, completed_at, reference_ids, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, 'draft', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?)
        """,
        (
            f"{source['title']} 출장결과",
            BUSINESS_TRIP_RESULT_TEMPLATE_TYPE,
            content,
            EDITOR_PROVIDER_GOOGLE_DOCS,
            external_doc_id,
            external_doc_url,
            source["priority"] or "normal",
            source["due_date"],
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            source["trip_department"] if "trip_department" in source.keys() else None,
            source["trip_job_title"] if "trip_job_title" in source.keys() else None,
            source["trip_name"] if "trip_name" in source.keys() else None,
            source["trip_type"] if "trip_type" in source.keys() else None,
            source["trip_destination"] if "trip_destination" in source.keys() else None,
            trip_start,
            trip_end,
            source["trip_transportation"] if "trip_transportation" in source.keys() else None,
            source["trip_expense"] if "trip_expense" in source.keys() else None,
            source["trip_purpose"] if "trip_purpose" in source.keys() else None,
            result_text,
            source_document_id,
            issue_code,
            default_issue_dept,
            effective_issue_year,
            source["recipient_text"] if "recipient_text" in source.keys() else "",
            visibility_scope,
            "[]",
            actor["id"],
            json.dumps(source_refs, ensure_ascii=False),
            now,
            now,
        ),
    )
    new_doc_id = int(cur.lastrowid or 0)
    if new_doc_id <= 0:
        raise ApiError(500, "출장결과 문서 생성에 실패했습니다.")
    for idx, approver_id in enumerate(normalized_approvers, start=1):
        conn.execute(
            "INSERT INTO approval_steps (document_id, step_order, approver_id, status) VALUES (?, ?, ?, 'waiting')",
            (new_doc_id, idx, approver_id),
        )

    warnings: list[str] = []
    try:
        fill_data = populate_google_doc_draft_fields(
            external_doc_id,
            title=f"{source['title']} 출장결과",
            issue_date=issue_date_for_doc,
            issue_code=issue_code,
            doc_form_type=document_form_type_label(BUSINESS_TRIP_RESULT_TEMPLATE_TYPE),
            recipient_text=source["recipient_text"] if "recipient_text" in source.keys() else "",
            trip_department=(source["trip_department"] if "trip_department" in source.keys() else "") or "",
            trip_job_title=(source["trip_job_title"] if "trip_job_title" in source.keys() else "") or "",
            trip_name=(source["trip_name"] if "trip_name" in source.keys() else "") or "",
            trip_type=(source["trip_type"] if "trip_type" in source.keys() else "") or "",
            trip_destination=(source["trip_destination"] if "trip_destination" in source.keys() else "") or "",
            trip_period=trip_period,
            trip_transportation=(source["trip_transportation"] if "trip_transportation" in source.keys() else "") or "",
            trip_expense=(source["trip_expense"] if "trip_expense" in source.keys() else "") or "",
            trip_purpose=(source["trip_purpose"] if "trip_purpose" in source.keys() else "") or "",
            trip_result=result_text,
            draft_date=issue_date_for_doc,
        )
        result_map = fill_data.get("result") if isinstance(fill_data.get("result"), dict) else {}
        updated_count = 0
        for key in (
            "trip_department",
            "trip_job_title",
            "trip_name",
            "trip_type",
            "trip_destination",
            "trip_period",
            "trip_transportation",
            "trip_expense",
            "trip_purpose",
            "trip_result",
        ):
            item = result_map.get(key) if isinstance(result_map, dict) else None
            if isinstance(item, dict) and item.get("updated"):
                updated_count += 1
        if updated_count == 0:
            warnings.append("출장결과 문서 자동입력 결과가 없습니다. 템플릿 토큰(TRIP_*)을 확인해 주세요.")
    except Exception as e:
        warnings.append(f"출장결과 문서 자동입력 경고: {e}")
        print(f"[populate_trip_result_fields] WARN doc={new_doc_id}: {type(e).__name__}: {e}")

    submitted = submit_document(conn, new_doc_id, int(actor["id"]))
    return submitted, warnings


def submit_education_result_document(
    conn: sqlite3.Connection,
    *,
    source_document_id: int,
    actor: sqlite3.Row,
    approver_ids: list[int],
    reference_ids: list[int] | None = None,
    education_content_text: str,
    education_apply_point_text: str,
) -> tuple[dict[str, Any], list[str]]:
    source = fetch_document(conn, source_document_id)
    if not source:
        raise ApiError(404, "교육 원문서를 찾을 수 없습니다.")
    if (source["template_type"] or "") != EDUCATION_TEMPLATE_TYPE:
        raise ApiError(400, "교육신청서 문서에서만 교육결과 상신이 가능합니다.")
    if int(source["drafter_id"]) != int(actor["id"]):
        raise ApiError(403, "교육결과는 원문 기안자만 상신할 수 있습니다.")
    if "is_deleted" in source.keys() and int(source["is_deleted"] or 0):
        raise ApiError(409, "보관삭제된 문서는 교육결과 상신이 불가능합니다.")

    content_text = str(education_content_text or "").strip()
    apply_point_text = str(education_apply_point_text or "").strip()
    if not content_text:
        raise ApiError(400, "교육내용을 입력해 주세요.")
    if not apply_point_text:
        raise ApiError(400, "적용점을 입력해 주세요.")

    normalized_approvers: list[int] = []
    for uid in approver_ids or []:
        try:
            iid = int(uid)
        except (TypeError, ValueError) as exc:
            raise ApiError(400, "교육결과 결재선 사용자 ID가 잘못되었습니다.") from exc
        if iid not in normalized_approvers:
            normalized_approvers.append(iid)
    if not normalized_approvers:
        raise ApiError(400, "교육결과 결재선을 최소 1명 이상 지정해 주세요.")
    max_steps_for_template = max_approver_steps_for_type(EDUCATION_RESULT_TEMPLATE_TYPE)
    if len(normalized_approvers) > max_steps_for_template:
        raise ApiError(
            400,
            f"교육결과 결재선은 최대 {max_steps_for_template}명까지 지정할 수 있습니다.",
        )
    if int(actor["id"]) in normalized_approvers:
        raise ApiError(400, "본인은 교육결과 결재자로 지정할 수 없습니다.")
    existing_approvers = {
        int(r["id"])
        for r in conn.execute(
            "SELECT id FROM users WHERE id IN ({})".format(",".join("?" for _ in normalized_approvers)),
            tuple(normalized_approvers),
        ).fetchall()
    }
    missing_approvers = [uid for uid in normalized_approvers if uid not in existing_approvers]
    if missing_approvers:
        raise ApiError(400, f"존재하지 않는 교육결과 결재자 ID가 있습니다: {missing_approvers}")

    normalized_refs: list[int] = []
    for rid in reference_ids or []:
        try:
            iid = int(rid)
        except (TypeError, ValueError):
            continue
        if iid not in normalized_refs:
            normalized_refs.append(iid)
    if normalized_refs:
        existing_refs = {
            int(r["id"])
            for r in conn.execute(
                "SELECT id FROM users WHERE id IN ({})".format(",".join("?" for _ in normalized_refs)),
                tuple(normalized_refs),
            ).fetchall()
        }
        normalized_refs = [rid for rid in normalized_refs if rid in existing_refs]

    copy_title = f"{source['title']}_교육보고서_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    copied = copy_google_doc_template_to_folder(
        EDUCATION_RESULT_TEMPLATE_DOC_ID,
        EDUCATION_RESULT_OUTPUT_FOLDER_ID,
        copy_title,
    )
    external_doc_id = copied["doc_id"]
    urls = build_google_doc_urls(external_doc_id)
    external_doc_url = urls["edit_url"]
    content = f"[Google Docs 문서] {external_doc_url}"

    issue_date_for_doc = today_str()
    default_issue_dept = str(source["issue_department"] or source["education_department"] or actor["department"] or "기타").strip() or "기타"
    issue_year_raw = str(source["issue_year"] or "").strip()
    effective_issue_year = issue_year_raw if re.fullmatch(r"\d{4}", issue_year_raw) else str(date.fromisoformat(issue_date_for_doc).year)
    issue_code = allocate_department_issue_code(
        conn,
        default_issue_dept,
        issue_date_for_doc,
        effective_issue_year,
    )
    visibility_scope = (source["visibility_scope"] or "private") if "visibility_scope" in source.keys() else "private"
    if visibility_scope not in DOC_VISIBILITY_VALUES:
        visibility_scope = "private"

    now = now_ts()
    education_start = validate_leave_datetime(source["education_start_date"], "교육 시작일시")
    education_end = validate_leave_datetime(source["education_end_date"], "교육 종료일시")
    education_period = format_education_period_text(education_start, education_end)
    source_refs: list[int] = []
    if normalized_refs:
        source_refs = normalized_refs
    else:
        try:
            loaded = json.loads(source["reference_ids"] or "[]")
            if isinstance(loaded, list):
                for rid in loaded:
                    try:
                        source_refs.append(int(rid))
                    except (TypeError, ValueError):
                        continue
        except Exception:
            source_refs = []

    cur = conn.execute(
        """
        INSERT INTO documents (title, template_type, content, editor_provider, external_doc_id, external_doc_url, status, priority, due_date, leave_type, leave_start_date, leave_end_date, leave_days, leave_reason, leave_substitute_name, leave_substitute_work, overtime_type, overtime_start_date, overtime_end_date, overtime_hours, overtime_content, overtime_etc, trip_department, trip_job_title, trip_name, trip_type, trip_destination, trip_start_date, trip_end_date, trip_transportation, trip_expense, trip_purpose, education_department, education_job_title, education_name, education_title, education_category, education_provider, education_location, education_start_date, education_end_date, education_purpose, education_tuition_detail, education_tuition_amount, education_material_detail, education_material_amount, education_transport_detail, education_transport_amount, education_other_detail, education_other_amount, education_budget_subject, education_funding_source, education_payment_method, education_support_budget, education_used_budget, education_remain_budget, education_companion, education_ordered, education_suggestion, source_trip_document_id, issue_code, issue_department, issue_year, recipient_text, visibility_scope, attachments_json, drafter_id, submitted_at, completed_at, reference_ids, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, 'draft', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?)
        """,
        (
            f"{source['title']} 교육결과",
            EDUCATION_RESULT_TEMPLATE_TYPE,
            content,
            EDITOR_PROVIDER_GOOGLE_DOCS,
            external_doc_id,
            external_doc_url,
            source["priority"] or "normal",
            source["due_date"],
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            source["education_department"] if "education_department" in source.keys() else None,
            source["education_job_title"] if "education_job_title" in source.keys() else None,
            source["education_name"] if "education_name" in source.keys() else None,
            source["education_title"] if "education_title" in source.keys() else None,
            source["education_category"] if "education_category" in source.keys() else None,
            source["education_provider"] if "education_provider" in source.keys() else None,
            source["education_location"] if "education_location" in source.keys() else None,
            education_start,
            education_end,
            source["education_purpose"] if "education_purpose" in source.keys() else None,
            source["education_tuition_detail"] if "education_tuition_detail" in source.keys() else None,
            source["education_tuition_amount"] if "education_tuition_amount" in source.keys() else None,
            source["education_material_detail"] if "education_material_detail" in source.keys() else None,
            source["education_material_amount"] if "education_material_amount" in source.keys() else None,
            source["education_transport_detail"] if "education_transport_detail" in source.keys() else None,
            source["education_transport_amount"] if "education_transport_amount" in source.keys() else None,
            source["education_other_detail"] if "education_other_detail" in source.keys() else None,
            source["education_other_amount"] if "education_other_amount" in source.keys() else None,
            source["education_budget_subject"] if "education_budget_subject" in source.keys() else None,
            source["education_funding_source"] if "education_funding_source" in source.keys() else None,
            source["education_payment_method"] if "education_payment_method" in source.keys() else None,
            source["education_support_budget"] if "education_support_budget" in source.keys() else None,
            source["education_used_budget"] if "education_used_budget" in source.keys() else None,
            source["education_remain_budget"] if "education_remain_budget" in source.keys() else None,
            source["education_companion"] if "education_companion" in source.keys() else None,
            source["education_ordered"] if "education_ordered" in source.keys() else None,
            source["education_suggestion"] if "education_suggestion" in source.keys() else None,
            source_document_id,
            issue_code,
            default_issue_dept,
            effective_issue_year,
            source["recipient_text"] if "recipient_text" in source.keys() else "",
            visibility_scope,
            "[]",
            actor["id"],
            json.dumps(source_refs, ensure_ascii=False),
            now,
            now,
        ),
    )
    new_doc_id = int(cur.lastrowid or 0)
    if new_doc_id <= 0:
        raise ApiError(500, "교육결과 문서 생성에 실패했습니다.")
    for idx, approver_id in enumerate(normalized_approvers, start=1):
        conn.execute(
            "INSERT INTO approval_steps (document_id, step_order, approver_id, status) VALUES (?, ?, ?, 'waiting')",
            (new_doc_id, idx, approver_id),
        )

    warnings: list[str] = []
    try:
        fill_data = populate_google_doc_draft_fields(
            external_doc_id,
            title=f"{source['title']} 교육결과",
            issue_date=issue_date_for_doc,
            issue_code=issue_code,
            doc_form_type=document_form_type_label(EDUCATION_RESULT_TEMPLATE_TYPE),
            recipient_text=source["recipient_text"] if "recipient_text" in source.keys() else "",
            education_department=(source["education_department"] if "education_department" in source.keys() else "") or "",
            education_job_title=(source["education_job_title"] if "education_job_title" in source.keys() else "") or "",
            education_name=(source["education_name"] if "education_name" in source.keys() else "") or "",
            education_title=(source["education_title"] if "education_title" in source.keys() else "") or "",
            education_category=(source["education_category"] if "education_category" in source.keys() else "") or "",
            education_provider=(source["education_provider"] if "education_provider" in source.keys() else "") or "",
            education_location=(source["education_location"] if "education_location" in source.keys() else "") or "",
            education_period=education_period,
            education_purpose=(source["education_purpose"] if "education_purpose" in source.keys() else "") or "",
            education_tuition_detail=(source["education_tuition_detail"] if "education_tuition_detail" in source.keys() else "") or "",
            education_tuition_amount=format_leave_number(source["education_tuition_amount"] if "education_tuition_amount" in source.keys() else None),
            education_material_detail=(source["education_material_detail"] if "education_material_detail" in source.keys() else "") or "",
            education_material_amount=format_leave_number(source["education_material_amount"] if "education_material_amount" in source.keys() else None),
            education_transport_detail=(source["education_transport_detail"] if "education_transport_detail" in source.keys() else "") or "",
            education_transport_amount=format_leave_number(source["education_transport_amount"] if "education_transport_amount" in source.keys() else None),
            education_other_detail=(source["education_other_detail"] if "education_other_detail" in source.keys() else "") or "",
            education_other_amount=format_leave_number(source["education_other_amount"] if "education_other_amount" in source.keys() else None),
            education_budget_subject=(source["education_budget_subject"] if "education_budget_subject" in source.keys() else "") or "",
            education_funding_source=(source["education_funding_source"] if "education_funding_source" in source.keys() else "") or "",
            education_payment_method=(source["education_payment_method"] if "education_payment_method" in source.keys() else "") or "",
            education_support_budget=format_leave_number(source["education_support_budget"] if "education_support_budget" in source.keys() else None),
            education_used_budget=format_leave_number(source["education_used_budget"] if "education_used_budget" in source.keys() else None),
            education_remain_budget=format_leave_number(source["education_remain_budget"] if "education_remain_budget" in source.keys() else None),
            education_companion=(source["education_companion"] if "education_companion" in source.keys() else "") or "",
            education_ordered=(source["education_ordered"] if "education_ordered" in source.keys() else "") or "",
            education_suggestion=(source["education_suggestion"] if "education_suggestion" in source.keys() else "") or "",
            education_content=content_text,
            education_apply_point=apply_point_text,
            draft_date=issue_date_for_doc,
        )
        result_map = fill_data.get("result") if isinstance(fill_data.get("result"), dict) else {}
        updated_count = 0
        for key in (
            "education_department",
            "education_job_title",
            "education_name",
            "education_title",
            "education_category",
            "education_provider",
            "education_location",
            "education_period",
            "education_purpose",
            "education_content",
            "education_apply_point",
        ):
            item = result_map.get(key) if isinstance(result_map, dict) else None
            if isinstance(item, dict) and item.get("updated"):
                updated_count += 1
        if updated_count == 0:
            warnings.append("교육결과 문서 자동입력 결과가 없습니다. 템플릿 토큰(EDU_*)을 확인해 주세요.")
    except Exception as e:
        warnings.append(f"교육결과 문서 자동입력 경고: {e}")
        print(f"[populate_education_result_fields] WARN doc={new_doc_id}: {type(e).__name__}: {e}")

    submitted = submit_document(conn, new_doc_id, int(actor["id"]))
    return submitted, warnings


def build_overtime_csv_export(payload: dict[str, Any]) -> bytes:
    selection = payload.get("selection") or {}
    summary = payload.get("summary") or {}
    records = payload.get("month_records") or []
    year = int(selection.get("year") or datetime.now().year)
    month = int(selection.get("month") or datetime.now().month)
    sio = io.StringIO()
    writer = csv.writer(sio)
    writer.writerow(["연장근로 내역", f"{year}년 {month}월"])
    writer.writerow([])
    writer.writerow(["월 기준(시간)", summary.get("monthly_cap_hours", 15)])
    writer.writerow(["사용(시간)", summary.get("month_used_hours", 0)])
    writer.writerow(["잔여(시간)", summary.get("month_remaining_hours", 0)])
    writer.writerow(["100%(시간)", summary.get("month_used_100_hours", 0)])
    writer.writerow(["150%(시간)", summary.get("month_used_150_hours", 0)])
    writer.writerow(["환산(1.5x)", summary.get("month_weighted_hours", 0)])
    writer.writerow([])
    writer.writerow(["형태", "시작일시", "종료일시", "시간", "100%", "150%", "문서상태", "문서제목"])
    for r in records:
        writer.writerow([
            r.get("overtime_type") or "",
            r.get("start_date") or "",
            r.get("end_date") or "",
            r.get("overtime_hours") or 0,
            r.get("overtime_100_hours") or 0,
            r.get("overtime_150_hours") or 0,
            r.get("document_status") or "",
            r.get("document_title") or "",
        ])
    csv_text = sio.getvalue()
    return ("\ufeff" + csv_text).encode("utf-8")


def submit_document(conn: sqlite3.Connection, document_id: int, actor_id: int) -> dict[str, Any]:
    doc = fetch_document(conn, document_id)
    if not doc:
        raise ApiError(404, "문서를 찾을 수 없습니다.")
    if "is_deleted" in doc.keys() and doc["is_deleted"]:
        raise ApiError(409, "보관삭제된 문서는 상신할 수 없습니다.")
    if doc["drafter_id"] != actor_id:
        raise ApiError(403, "문서 기안자만 상신할 수 있습니다.")
    if doc["status"] != "draft":
        raise ApiError(409, "임시저장 상태의 문서만 상신할 수 있습니다.")

    steps = conn.execute(
        "SELECT id, approver_id FROM approval_steps WHERE document_id = ? ORDER BY step_order ASC",
        (document_id,),
    ).fetchall()
    if not steps:
        raise ApiError(400, "결재선이 없습니다. 결재자를 지정해 주세요.")
    total_slots = approval_template_total_slots_for_type(doc["template_type"] if "template_type" in doc.keys() else None)
    max_steps = max_approver_steps_for_type(doc["template_type"] if "template_type" in doc.keys() else None)
    if len(steps) > max_steps:
        raise ApiError(400, f"현재 템플릿은 결재자 최대 {max_steps}명까지 지원합니다. (기안자 포함 {total_slots}칸)")
    if (doc["editor_provider"] or EDITOR_PROVIDER_INTERNAL) == EDITOR_PROVIDER_GOOGLE_DOCS and doc["external_doc_id"]:
        validate_approval_template_tokens(str(doc["external_doc_id"]), len(steps) + 1)

    first = steps[0]
    now = now_ts()
    conn.execute("UPDATE documents SET status='in_review', submitted_at=?, updated_at=? WHERE id=?", (now, now, document_id))
    conn.execute("UPDATE approval_steps SET status='pending' WHERE id=?", (first["id"],))
    create_notification(conn, first["approver_id"], f"새 결재 요청: {doc['title']}", f"/documents/{document_id}")
    sync_leave_usage_for_document(conn, document_id)
    sync_approval_signatures_to_google_doc(conn, document_id)

    updated = fetch_document(conn, document_id)
    if not updated:
        raise ApiError(500, "상신 처리 후 문서를 확인할 수 없습니다.")
    return dict_document(conn, updated, include_detail=True)


def resubmit_rejected_document(conn: sqlite3.Connection, document_id: int, actor_id: int) -> dict[str, Any]:
    doc = fetch_document(conn, document_id)
    if not doc:
        raise ApiError(404, "문서를 찾을 수 없습니다.")
    if "is_deleted" in doc.keys() and doc["is_deleted"]:
        raise ApiError(409, "보관삭제된 문서는 재기안 상신할 수 없습니다.")
    if doc["drafter_id"] != actor_id:
        raise ApiError(403, "문서 기안자만 재기안 상신할 수 있습니다.")
    if doc["status"] != "rejected":
        # Compatibility path for previously "returned" documents that were kept in in_review.
        if doc["status"] != "in_review":
            raise ApiError(409, "반려 문서만 재기안 상신할 수 있습니다.")
        has_rejected = conn.execute(
            "SELECT 1 FROM approval_steps WHERE document_id=? AND status='rejected' LIMIT 1",
            (document_id,),
        ).fetchone()
        current_pending = conn.execute(
            "SELECT 1 FROM approval_steps WHERE document_id=? AND status='pending' LIMIT 1",
            (document_id,),
        ).fetchone()
        # Allow resubmission if there is a recorded rejection, or if workflow is effectively returned/stuck
        # with no pending approver (legacy behavior compatibility).
        if not has_rejected and current_pending:
            raise ApiError(409, "반려 문서만 재기안 상신할 수 있습니다.")

    steps = conn.execute(
        "SELECT id, approver_id FROM approval_steps WHERE document_id = ? ORDER BY step_order ASC",
        (document_id,),
    ).fetchall()
    if not steps:
        raise ApiError(400, "결재선이 없습니다. 결재자를 지정해 주세요.")
    total_slots = approval_template_total_slots_for_type(doc["template_type"] if "template_type" in doc.keys() else None)
    max_steps = max_approver_steps_for_type(doc["template_type"] if "template_type" in doc.keys() else None)
    if len(steps) > max_steps:
        raise ApiError(400, f"현재 템플릿은 결재자 최대 {max_steps}명까지 지원합니다. (기안자 포함 {total_slots}칸)")

    # Re-submission uses the same copied Google Doc, so token placeholders may no longer exist.
    # We skip template token validation here and rely on slot-based updates.
    if (doc["editor_provider"] or EDITOR_PROVIDER_INTERNAL) == EDITOR_PROVIDER_GOOGLE_DOCS and doc["external_doc_id"]:
        total_slots = approval_template_total_slots_for_type(doc["template_type"] if "template_type" in doc.keys() else None)
        try:
            reset_google_approval_doc_slots(str(doc["external_doc_id"]), total_slots)
        except Exception as e:
            # Best effort for legacy docs whose placeholders were removed; sync_approval_signatures may still rebuild via layout.
            print(f"[resubmit_reset_approval_doc] WARN doc={document_id}: {type(e).__name__}: {e}")

    now = now_ts()
    conn.execute(
        "UPDATE documents SET status='in_review', submitted_at=?, completed_at=NULL, updated_at=? WHERE id=?",
        (now, now, document_id),
    )
    conn.execute(
        "UPDATE approval_steps SET status='waiting', acted_at=NULL, comment=NULL WHERE document_id=?",
        (document_id,),
    )
    first = steps[0]
    conn.execute("UPDATE approval_steps SET status='pending' WHERE id=?", (first["id"],))
    create_notification(conn, first["approver_id"], f"재기안 결재 요청: {doc['title']}", f"/documents/{document_id}")
    sync_leave_usage_for_document(conn, document_id)
    sync_approval_signatures_to_google_doc(conn, document_id)

    refreshed = fetch_document(conn, document_id)
    if not refreshed:
        raise ApiError(500, "재기안 상신 후 문서를 확인할 수 없습니다.")
    return dict_document(conn, refreshed, include_detail=True)


def apply_approval_action(conn: sqlite3.Connection, document_id: int, actor: sqlite3.Row, action: str, comment: str) -> dict[str, Any]:
    doc = fetch_document(conn, document_id)
    if not doc:
        raise ApiError(404, "문서를 찾을 수 없습니다.")
    if "is_deleted" in doc.keys() and doc["is_deleted"]:
        raise ApiError(409, "보관삭제된 문서는 결재 처리할 수 없습니다.")
    if not can_view_document_row(conn, doc, actor):
        raise ApiError(403, "해당 문서를 열람할 권한이 없습니다.")

    if action == "comment":
        if not comment.strip():
            raise ApiError(400, "코멘트 내용을 입력해 주세요.")
        conn.execute(
            "INSERT INTO document_comments (document_id, user_id, comment, created_at) VALUES (?, ?, ?, ?)",
            (document_id, actor["id"], comment.strip(), now_ts()),
        )
        if doc["drafter_id"] != actor["id"]:
            create_notification(conn, doc["drafter_id"], f"문서 '{doc['title']}'에 새 코멘트가 등록되었습니다.", f"/documents/{document_id}")
        refreshed = fetch_document(conn, document_id)
        if not refreshed:
            raise ApiError(500, "문서를 다시 불러오지 못했습니다.")
        return dict_document(conn, refreshed, include_detail=True)

    if doc["status"] != "in_review":
        raise ApiError(409, "진행 중인 결재 문서에서만 승인/반려할 수 있습니다.")

    pending = conn.execute(
        """
        SELECT id, step_order
        FROM approval_steps
        WHERE document_id=? AND approver_id=? AND status='pending'
        ORDER BY step_order ASC LIMIT 1
        """,
        (document_id, actor["id"]),
    ).fetchone()
    if not pending:
        raise ApiError(403, "현재 사용자에게 대기 중인 결재 단계가 없습니다.")

    now = now_ts()
    if action == "approve":
        conn.execute("UPDATE approval_steps SET status='approved', acted_at=?, comment=? WHERE id=?", (now, comment.strip() or None, pending["id"]))
        nxt = conn.execute(
            "SELECT id, approver_id FROM approval_steps WHERE document_id=? AND step_order > ? AND status IN ('waiting','rejected') ORDER BY step_order ASC LIMIT 1",
            (document_id, pending["step_order"]),
        ).fetchone()
        if nxt:
            conn.execute("UPDATE approval_steps SET status='pending', acted_at=NULL, comment=NULL WHERE id=?", (nxt["id"],))
            conn.execute("UPDATE documents SET updated_at=? WHERE id=?", (now, document_id))
            create_notification(conn, nxt["approver_id"], f"결재 차례가 되었습니다: {doc['title']}", f"/documents/{document_id}")
        else:
            conn.execute("UPDATE documents SET status='approved', completed_at=?, updated_at=? WHERE id=?", (now, now, document_id))
            conn.execute("UPDATE approval_steps SET status='skipped' WHERE document_id=? AND status='waiting'", (document_id,))
            create_notification(conn, doc["drafter_id"], f"문서 '{doc['title']}'가 최종 승인되었습니다.", f"/documents/{document_id}")
            if doc["template_type"] == "leave":
                owner = conn.execute("SELECT full_name FROM users WHERE id=?", (doc["drafter_id"],)).fetchone()
                conn.execute(
                    "INSERT INTO schedules (title, start_date, end_date, event_type, owner_id, resource_name, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (f"{owner['full_name']} 휴가", doc["due_date"] or today_str(), doc["due_date"] or today_str(), "leave", doc["drafter_id"], None, now),
                )
    elif action == "reject":
        if not comment.strip():
            raise ApiError(400, "반려 시 반려 사유를 입력해 주세요.")
        conn.execute("UPDATE approval_steps SET status='rejected', acted_at=?, comment=? WHERE id=?", (now, comment.strip(), pending["id"]))
        # Rejection always returns the document to the drafter for edit/resubmission.
        conn.execute("UPDATE approval_steps SET status='skipped' WHERE document_id=? AND status IN ('waiting','pending')", (document_id,))
        conn.execute("UPDATE documents SET status='rejected', completed_at=?, updated_at=? WHERE id=?", (now, now, document_id))
        create_notification(conn, doc["drafter_id"], f"문서 '{doc['title']}'가 반려되었습니다. 수정 후 재기안 상신해 주세요.", f"/documents/{document_id}")
    else:
        raise ApiError(400, "지원하지 않는 액션입니다.")

    sync_leave_usage_for_document(conn, document_id)
    sync_approval_signatures_to_google_doc(conn, document_id)
    refreshed = fetch_document(conn, document_id)
    if not refreshed:
        raise ApiError(500, "결재 처리 후 문서를 확인할 수 없습니다.")
    if action in {"approve", "reject"} and (refreshed["status"] or "") in {"approved", "rejected"}:
        try:
            ensure_document_print_snapshot(conn, document_id, force=True)
            refreshed = fetch_document(conn, document_id) or refreshed
        except Exception as e:
            print(f"[print_snapshot] WARN approval action doc={document_id}: {type(e).__name__}: {e}")
    return dict_document(conn, refreshed, include_detail=True)


def validate_date(date_value: str | None, field_name: str) -> str | None:
    if not date_value:
        return None
    try:
        date.fromisoformat(date_value)
        return date_value
    except ValueError as exc:
        raise ApiError(400, f"{field_name} 형식은 YYYY-MM-DD 이어야 합니다.") from exc


class ApprovalHandler(BaseHTTPRequestHandler):
    server_version = "ApprovalServer/1.0"

    def do_GET(self) -> None:
        self.route_request("GET")

    def do_POST(self) -> None:
        self.route_request("POST")

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        _add_cors_headers(self)
        self.end_headers()

    def route_request(self, method: str) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)
        if path.startswith("/api/"):
            self.handle_api(method, path, query)
            return
        self.serve_static(path)

    def handle_api(self, method: str, path: str, query: dict[str, list[str]]) -> None:
        try:
            with get_conn() as conn:
                if method == "GET" and path == "/api/health":
                    json_response(self, 200, {"ok": True, "server_time": now_ts()})
                    return

                if method == "POST" and path == "/api/auth/login":
                    body = parse_json_body(self)
                    username = (body.get("username") or "").strip()
                    password = body.get("password") or ""
                    if not username or not password:
                        raise ApiError(400, "아이디와 비밀번호를 입력해 주세요.")
                    row = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
                    if not row:
                        raise ApiError(401, "로그인 정보가 올바르지 않습니다.")
                    if (
                        (row["auth_provider"] if "auth_provider" in row.keys() else "local") == "google"
                        and row["password_hash"] != hash_password(password)
                    ):
                        raise ApiError(401, "Google 계정으로 가입된 사용자입니다. 아래 Google 로그인 버튼을 사용해 주세요.")
                    if row["password_hash"] != hash_password(password):
                        raise ApiError(401, "로그인 정보가 올바르지 않습니다.")
                    token = new_session(row["id"])
                    user = get_user_by_id(conn, row["id"])
                    if not user:
                        raise ApiError(500, "사용자 정보를 조회하지 못했습니다.")
                    log_to_sheet(user["username"], self.client_address[0])
                    json_response(self, 200, {"token": token, "user": dict_user(user)})
                    return

                if method == "GET" and path == "/api/auth/google/config":
                    json_response(self, 200, google_auth_config_public())
                    return

                if method == "POST" and path == "/api/auth/google":
                    body = parse_json_body(self)
                    id_token = str(body.get("id_token") or "").strip()
                    cfg = google_auth_config_public()
                    if not cfg.get("enabled"):
                        raise ApiError(503, "Google 회원가입/로그인이 비활성화되어 있습니다.")

                    claims = verify_google_id_token(id_token, str(cfg.get("client_id") or ""))
                    google_sub = str(claims.get("sub") or "").strip()
                    google_email = str(claims.get("email") or "").strip().lower()
                    full_name = str(claims.get("name") or "").strip() or google_email.split("@", 1)[0]

                    found = conn.execute(
                        "SELECT id FROM users WHERE google_sub=?",
                        (google_sub,),
                    ).fetchone()
                    if not found:
                        found = conn.execute(
                            "SELECT id FROM users WHERE lower(COALESCE(google_email,''))=lower(?)",
                            (google_email,),
                        ).fetchone()

                    created = False
                    if found:
                        user_id = int(found["id"])
                    else:
                        username = build_google_username(conn, google_email, google_sub)
                        temp_password = secrets.token_urlsafe(24)
                        cur = conn.execute(
                            """
                            INSERT INTO users (
                                username, password_hash, full_name, role, department,
                                job_title, total_leave, used_leave, created_at,
                                auth_provider, google_sub, google_email
                            )
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                username,
                                hash_password(temp_password),
                                full_name,
                                "employee",
                                "미지정",
                                "",
                                15.0,
                                0.0,
                                now_ts(),
                                "google",
                                google_sub,
                                google_email,
                            ),
                        )
                        user_id = int(cur.lastrowid)
                        created = True

                    user_row = get_user_by_id(conn, user_id)
                    if not user_row:
                        raise ApiError(500, "Google 사용자 정보를 조회하지 못했습니다.")

                    if created:
                        sync_user_to_sheet(dict_user(user_row), None, None)
                    token = new_session(user_id)
                    log_to_sheet(user_row["username"], self.client_address[0])
                    json_response(
                        self,
                        200,
                        {"token": token, "user": dict_user(user_row), "created": created},
                    )
                    return

                if method == "POST" and path == "/api/auth/logout":
                    drop_session(self.headers)
                    json_response(self, 200, {"ok": True})
                    return

                user = require_user(conn, self.headers)

                if method == "GET" and path == "/api/integrations/google/config":
                    json_response(self, 200, {"config": google_integration_config()})
                    return
                if method == "POST" and path == "/api/integrations/google/validate-template":
                    body = parse_json_body(self)
                    doc_id = extract_google_doc_id((body.get("doc_id") or body.get("doc_url") or "").strip() if isinstance(body.get("doc_id") or body.get("doc_url"), str) else str(body.get("doc_id") or body.get("doc_url") or ""))
                    if not doc_id:
                        raise ApiError(400, "검증할 Google Docs 문서 ID가 필요합니다.")
                    required_slots = int(body.get("required_slots") or APPROVAL_TEMPLATE_TOTAL_SLOTS)
                    result = validate_approval_template_tokens(doc_id, required_slots)
                    json_response(self, 200, {"result": result})
                    return
                if method == "POST" and path == "/api/integrations/google/reset-approval-doc":
                    body = parse_json_body(self)
                    doc_id = extract_google_doc_id(
                        (body.get("doc_id") or body.get("doc_url") or "").strip()
                        if isinstance(body.get("doc_id") or body.get("doc_url"), str)
                        else str(body.get("doc_id") or body.get("doc_url") or "")
                    )
                    if not doc_id:
                        raise ApiError(400, "초기화할 Google Docs 문서 ID가 필요합니다.")
                    total_slots = int(body.get("total_slots") or APPROVAL_TEMPLATE_TOTAL_SLOTS)
                    result = reset_google_approval_doc_slots(doc_id, total_slots)
                    json_response(self, 200, {"result": result})
                    return
                if method == "POST" and path == "/api/integrations/google/upload-attachments":
                    body = parse_json_body(self)
                    files = body.get("files") if isinstance(body.get("files"), list) else []
                    if not files:
                        json_response(self, 200, {"attachments": []})
                        return
                    if len(files) > 10:
                        raise ApiError(400, "첨부파일은 최대 10개까지 업로드할 수 있습니다.")
                    attachments = upload_document_attachments_to_drive([f for f in files if isinstance(f, dict)])
                    json_response(self, 200, {"attachments": attachments})
                    return

                if method == "GET" and path == "/api/auth/me":
                    json_response(self, 200, {"user": dict_user(user)})
                    return

                if method == "GET" and path == "/api/ui/tab-order":
                    row = conn.execute(
                        "SELECT value_json FROM app_settings WHERE key = 'global_tab_order'"
                    ).fetchone()
                    if not row or not row["value_json"]:
                        # Legacy fallback: last saved admin preference
                        row = conn.execute(
                            """
                            SELECT p.tab_order_json AS value_json
                            FROM user_ui_preferences p
                            JOIN users u ON u.id = p.user_id
                            WHERE u.role = 'admin'
                            ORDER BY p.updated_at DESC, p.user_id DESC
                            LIMIT 1
                            """
                        ).fetchone()
                    tab_order: list[str] = []
                    if row and row["value_json"]:
                        try:
                            parsed = json.loads(row["value_json"])
                            if isinstance(parsed, list):
                                for item in parsed:
                                    key = str(item or "").strip()
                                    if not key or key in tab_order:
                                        continue
                                    if len(key) > 64:
                                        continue
                                    tab_order.append(key)
                        except Exception:
                            tab_order = []
                    json_response(self, 200, {"tab_order": tab_order})
                    return

                if method == "POST" and path == "/api/ui/tab-order":
                    if user["role"] != "admin":
                        raise ApiError(403, "탭 순서 저장은 관리자만 가능합니다.")
                    body = parse_json_body(self)
                    raw_order = body.get("tab_order")
                    if not isinstance(raw_order, list) or not raw_order:
                        raise ApiError(400, "저장할 탭 순서가 없습니다.")
                    normalized: list[str] = []
                    for item in raw_order:
                        key = str(item or "").strip()
                        if not key:
                            continue
                        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]{0,63}", key):
                            raise ApiError(400, "탭 식별자 형식이 올바르지 않습니다.")
                        if key in normalized:
                            continue
                        normalized.append(key)
                    if not normalized:
                        raise ApiError(400, "저장할 탭 순서가 없습니다.")
                    if len(normalized) > 20:
                        raise ApiError(400, "탭 순서 항목 수가 너무 많습니다.")
                    conn.execute(
                        """
                        INSERT INTO user_ui_preferences (user_id, tab_order_json, updated_at)
                        VALUES (?, ?, ?)
                        ON CONFLICT(user_id) DO UPDATE SET
                            tab_order_json=excluded.tab_order_json,
                            updated_at=excluded.updated_at
                        """,
                        (user["id"], json.dumps(normalized, ensure_ascii=False), now_ts()),
                    )
                    conn.execute(
                        """
                        INSERT INTO app_settings (key, value_json, updated_at)
                        VALUES ('global_tab_order', ?, ?)
                        ON CONFLICT(key) DO UPDATE SET
                            value_json=excluded.value_json,
                            updated_at=excluded.updated_at
                        """,
                        (json.dumps(normalized, ensure_ascii=False), now_ts()),
                    )
                    json_response(self, 200, {"tab_order": normalized})
                    return

                if method == "GET" and path == "/api/users":
                    users = conn.execute("SELECT * FROM users ORDER BY full_name").fetchall()
                    json_response(self, 200, {"users": [dict_user(u) for u in users]})
                    return

                if method == "POST" and path == "/api/users":
                    if user["role"] != "admin":
                        raise ApiError(403, "사용자 등록은 관리자만 가능합니다.")
                    body = parse_json_body(self)
                    username = (body.get("username") or "").strip()
                    password = body.get("password") or ""
                    full_name = (body.get("full_name") or "").strip()
                    department = (body.get("department") or "").strip()
                    role = (body.get("role") or "employee").strip() or "employee"

                    if not USERNAME_PATTERN.fullmatch(username):
                        raise ApiError(400, "아이디는 영문/숫자/._- 조합 3~40자로 입력해 주세요.")
                    if len(password) < 6:
                        raise ApiError(400, "비밀번호는 6자 이상이어야 합니다.")
                    if not full_name:
                        raise ApiError(400, "이름을 입력해 주세요.")
                    if not department:
                        raise ApiError(400, "부서를 입력해 주세요.")
                    if role not in ("admin", "executive", "employee"):
                        raise ApiError(400, "권한 값이 올바르지 않습니다.")

                    exists = conn.execute("SELECT id FROM users WHERE username=?", (username,)).fetchone()
                    if exists:
                        raise ApiError(409, "이미 존재하는 아이디입니다.")

                    job_title = (body.get("job_title") or "").strip()
                    total_leave = float(body.get("total_leave") or 0)
                    used_leave = float(body.get("used_leave") or 0)

                    cur = conn.execute(
                        """
                        INSERT INTO users (username, password_hash, full_name, role, department, job_title, total_leave, used_leave, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (username, hash_password(password), full_name, role, department, job_title, total_leave, used_leave, now_ts()),
                    )
                    user_id = cur.lastrowid
                    created = get_user_by_id(conn, int(user_id))
                    if not created:
                        raise ApiError(500, "생성한 사용자를 조회하지 못했습니다.")
                    
                    profile_image = body.get("profile_image")
                    approval_stamp_image = body.get("approval_stamp_image")
                    upload_urls = sync_user_to_sheet(dict_user(created), profile_image, approval_stamp_image)

                    if upload_urls.get("profile_image_url") or upload_urls.get("approval_stamp_image_url"):
                        conn.execute(
                            """
                            UPDATE users
                            SET profile_image_url = COALESCE(?, profile_image_url),
                                approval_stamp_image_url = COALESCE(?, approval_stamp_image_url)
                            WHERE id = ?
                            """,
                            (upload_urls.get("profile_image_url"), upload_urls.get("approval_stamp_image_url"), user_id),
                        )
                        created = get_user_by_id(conn, int(user_id)) # Reload with URL

                    json_response(self, 201, {"user": dict_user(created)})
                    return

                m_user_delete = re.fullmatch(r"/api/users/(\d+)/delete", path)
                if method == "POST" and m_user_delete:
                    if user["role"] != "admin":
                        raise ApiError(403, "사용자 삭제는 관리자만 가능합니다.")
                    target_id = int(m_user_delete.group(1))
                    if target_id == user["id"]:
                        raise ApiError(400, "현재 로그인한 관리자는 삭제할 수 없습니다.")

                    target = conn.execute("SELECT id, username FROM users WHERE id=?", (target_id,)).fetchone()
                    if not target:
                        raise ApiError(404, "삭제할 사용자를 찾을 수 없습니다.")

                    # Pre-delete cleanup
                    # 1. Identify Admin
                    admin_row = conn.execute("SELECT id FROM users WHERE role='admin' ORDER BY id ASC LIMIT 1").fetchone()
                    admin_id = admin_row["id"] if admin_row else user["id"]

                    # 2. Delete non-critical
                    conn.execute("DELETE FROM notifications WHERE user_id=?", (target_id,))
                    conn.execute("DELETE FROM schedules WHERE owner_id=?", (target_id,))
                    conn.execute("DELETE FROM document_comments WHERE user_id=?", (target_id,))
                    
                    # 3. Reassign critical
                    conn.execute("UPDATE notices SET author_id=? WHERE author_id=?", (admin_id, target_id))
                    conn.execute("UPDATE documents SET drafter_id=? WHERE drafter_id=?", (admin_id, target_id))
                    conn.execute("UPDATE approval_steps SET approver_id=? WHERE approver_id=?", (admin_id, target_id))

                    # 4. Delete User
                    conn.execute("DELETE FROM users WHERE id=?", (target_id,))
                    
                    delete_user_from_sheet(target["username"])
                    
                    json_response(self, 200, {"ok": True, "deleted_user_id": target_id})
                    return

                m_user_update = re.fullmatch(r"/api/users/(\d+)/update", path)
                if method == "POST" and m_user_update:
                    if user["role"] != "admin":
                        raise ApiError(403, "사용자 수정은 관리자만 가능합니다.")
                    target_id = int(m_user_update.group(1))
                    
                    body = parse_json_body(self)
                    password = body.get("password") or ""
                    full_name = (body.get("full_name") or "").strip()
                    department = (body.get("department") or "").strip()
                    role = (body.get("role") or "employee").strip()
                    job_title = (body.get("job_title") or "").strip()
                    total_leave = float(body.get("total_leave") or 0)
                    used_leave = float(body.get("used_leave") or 0)

                    if not full_name or not department:
                        raise ApiError(400, "이름과 부서는 필수입니다.")

                    try:
                        if password and len(password) >= 6:
                            conn.execute(
                                "UPDATE users SET password_hash=?, full_name=?, department=?, role=?, job_title=?, total_leave=?, used_leave=? WHERE id=?",
                                (hash_password(password), full_name, department, role, job_title, total_leave, used_leave, target_id)
                            )
                        else:
                            conn.execute(
                                "UPDATE users SET full_name=?, department=?, role=?, job_title=?, total_leave=?, used_leave=? WHERE id=?",
                                (full_name, department, role, job_title, total_leave, used_leave, target_id)
                            )
                    except sqlite3.IntegrityError:
                        raise ApiError(409, "데이터베이스 오류가 발생했습니다.")
                        
                    updated = conn.execute("SELECT * FROM users WHERE id=?", (target_id,)).fetchone()
                    if not updated:
                        raise ApiError(404, "사용자를 찾을 수 없습니다.")
                    
                    profile_image = body.get("profile_image")
                    approval_stamp_image = body.get("approval_stamp_image")
                    print(f"[user_update] profile_image present: {profile_image is not None}, type: {type(profile_image).__name__ if profile_image else 'None'}")
                    if profile_image:
                        print(f"[user_update] image keys: {list(profile_image.keys()) if isinstance(profile_image, dict) else 'not a dict'}, data length: {len(profile_image.get('data','')) if isinstance(profile_image, dict) else 0}")
                    if approval_stamp_image:
                        print(f"[user_update] approval_stamp_image keys: {list(approval_stamp_image.keys()) if isinstance(approval_stamp_image, dict) else 'not a dict'}, data length: {len(approval_stamp_image.get('data','')) if isinstance(approval_stamp_image, dict) else 0}")
                    upload_urls = sync_user_to_sheet(dict_user(updated), profile_image, approval_stamp_image)

                    if upload_urls.get("profile_image_url") or upload_urls.get("approval_stamp_image_url"):
                        conn.execute(
                            """
                            UPDATE users
                            SET profile_image_url = COALESCE(?, profile_image_url),
                                approval_stamp_image_url = COALESCE(?, approval_stamp_image_url)
                            WHERE id = ?
                            """,
                            (upload_urls.get("profile_image_url"), upload_urls.get("approval_stamp_image_url"), target_id),
                        )
                        updated = conn.execute("SELECT * FROM users WHERE id=?", (target_id,)).fetchone()

                    json_response(self, 200, {"user": dict_user(updated)})
                    return

                if method == "GET" and path == "/api/dashboard":
                    pending = conn.execute(
                        """
                        SELECT COUNT(*) c
                        FROM documents d
                        WHERE COALESCE(d.is_deleted,0)=0
                          AND (
                            EXISTS (
                              SELECT 1 FROM approval_steps s
                              WHERE s.document_id=d.id AND s.approver_id=? AND s.status='pending'
                            )
                            OR (
                              COALESCE(d.edit_request_status,'none')='pending'
                              AND d.edit_request_reviewer_id=?
                            )
                            OR (
                              COALESCE(d.delete_request_status,'none')='pending'
                              AND d.delete_request_reviewer_id=?
                            )
                          )
                        """,
                        (user["id"], user["id"], user["id"]),
                    ).fetchone()["c"]
                    drafts = conn.execute("SELECT COUNT(*) c FROM documents WHERE drafter_id=? AND status='draft' AND COALESCE(is_deleted,0)=0", (user["id"],)).fetchone()["c"]
                    reviewing = conn.execute("SELECT COUNT(*) c FROM documents WHERE drafter_id=? AND status IN ('in_review','rejected') AND COALESCE(is_deleted,0)=0", (user["id"],)).fetchone()["c"]
                    done = conn.execute("SELECT COUNT(*) c FROM documents WHERE drafter_id=? AND status='approved' AND COALESCE(is_deleted,0)=0", (user["id"],)).fetchone()["c"]
                    unread = conn.execute("SELECT COUNT(*) c FROM notifications WHERE user_id=? AND is_read=0", (user["id"],)).fetchone()["c"]
                    week = (date.today() + timedelta(days=7)).isoformat()
                    upcoming = conn.execute("SELECT COUNT(*) c FROM schedules WHERE start_date BETWEEN ? AND ? AND status='active'", (today_str(), week)).fetchone()["c"]
                    draft_rows = conn.execute(
                        """
                        SELECT d.*, u.full_name drafter_name, u.department drafter_department
                        FROM documents d JOIN users u ON u.id = d.drafter_id
                        WHERE d.drafter_id = ? AND d.status = 'draft' AND COALESCE(d.is_deleted,0)=0
                        ORDER BY d.updated_at DESC
                        LIMIT 20
                        """,
                        (user["id"],),
                    ).fetchall()
                    reviewing_rows = conn.execute(
                        """
                        SELECT d.*, u.full_name drafter_name, u.department drafter_department
                        FROM documents d JOIN users u ON u.id = d.drafter_id
                        WHERE d.drafter_id = ? AND d.status IN ('in_review','rejected') AND COALESCE(d.is_deleted,0)=0
                        ORDER BY d.updated_at DESC
                        LIMIT 20
                        """,
                        (user["id"],),
                    ).fetchall()
                    pending_rows = conn.execute(
                        """
                        SELECT d.*, u.full_name drafter_name, u.department drafter_department
                        FROM documents d
                        JOIN users u ON u.id = d.drafter_id
                        WHERE COALESCE(d.is_deleted,0)=0
                          AND (
                            EXISTS (
                              SELECT 1 FROM approval_steps s
                              WHERE s.document_id=d.id AND s.approver_id=? AND s.status='pending'
                            )
                            OR (
                              COALESCE(d.edit_request_status,'none')='pending'
                              AND d.edit_request_reviewer_id=?
                            )
                            OR (
                              COALESCE(d.delete_request_status,'none')='pending'
                              AND d.delete_request_reviewer_id=?
                            )
                          )
                        ORDER BY d.updated_at DESC
                        LIMIT 20
                        """,
                        (user["id"], user["id"], user["id"]),
                    ).fetchall()
                    completed_rows = conn.execute(
                        """
                        SELECT d.*, u.full_name drafter_name, u.department drafter_department
                        FROM documents d JOIN users u ON u.id = d.drafter_id
                        WHERE d.drafter_id = ? AND d.status = 'approved' AND COALESCE(d.is_deleted,0)=0
                        ORDER BY COALESCE(d.completed_at, d.updated_at) DESC
                        LIMIT 20
                        """,
                        (user["id"],),
                    ).fetchall()
                    json_response(self, 200, {
                        "pending_count": pending,
                        "my_draft_count": drafts,
                        "my_in_review_count": reviewing,
                        "my_pending_approval_count": pending,
                        "my_completed_count": done,
                        "unread_notifications": unread,
                        "upcoming_schedule_count": upcoming,
                        "my_drafts": [dict_document(conn, r, include_detail=False) for r in draft_rows],
                        "my_in_review": [dict_document(conn, r, include_detail=False) for r in reviewing_rows],
                        "my_pending_approvals": [dict_document(conn, r, include_detail=False) for r in pending_rows],
                        "my_completed": [dict_document(conn, r, include_detail=False) for r in completed_rows],
                    })
                    return

                if method == "GET" and path == "/api/leaves":
                    json_response(self, 200, list_leave_usages(conn, user))
                    return

                if method == "GET" and path == "/api/overtimes":
                    try:
                        year = int((query.get("year", [""])[0] or "").strip()) if (query.get("year", [""])[0] or "").strip() else None
                    except Exception:
                        year = None
                    try:
                        month = int((query.get("month", [""])[0] or "").strip()) if (query.get("month", [""])[0] or "").strip() else None
                    except Exception:
                        month = None
                    json_response(self, 200, list_overtime_usages(conn, user, year, month))
                    return

                if method == "GET" and path == "/api/business-trips":
                    try:
                        year = int((query.get("year", [""])[0] or "").strip()) if (query.get("year", [""])[0] or "").strip() else None
                    except Exception:
                        year = None
                    try:
                        month = int((query.get("month", [""])[0] or "").strip()) if (query.get("month", [""])[0] or "").strip() else None
                    except Exception:
                        month = None
                    json_response(self, 200, list_business_trip_usages(conn, user, year, month))
                    return

                if method == "GET" and path == "/api/educations":
                    try:
                        year = int((query.get("year", [""])[0] or "").strip()) if (query.get("year", [""])[0] or "").strip() else None
                    except Exception:
                        year = None
                    try:
                        month = int((query.get("month", [""])[0] or "").strip()) if (query.get("month", [""])[0] or "").strip() else None
                    except Exception:
                        month = None
                    json_response(self, 200, list_education_usages(conn, user, year, month))
                    return

                m_trip_result_submit = re.fullmatch(r"/api/business-trips/(\d+)/result", path)
                if method == "POST" and m_trip_result_submit:
                    body = parse_json_body(self)
                    raw_approver_ids = body.get("approver_ids")
                    # Backward compatibility for older clients that still send reviewer_id.
                    if raw_approver_ids is None:
                        legacy_reviewer_id = body.get("reviewer_id")
                        raw_approver_ids = [legacy_reviewer_id] if legacy_reviewer_id not in (None, "", 0, "0") else []
                    if not isinstance(raw_approver_ids, list):
                        raise ApiError(400, "출장결과 결재선 형식이 올바르지 않습니다.")
                    approver_ids: list[int] = []
                    for uid in raw_approver_ids:
                        try:
                            approver_ids.append(int(uid))
                        except (TypeError, ValueError):
                            raise ApiError(400, "출장결과 결재선 사용자 ID가 올바르지 않습니다.")

                    raw_reference_ids = body.get("reference_ids") or []
                    if not isinstance(raw_reference_ids, list):
                        raise ApiError(400, "출장결과 참조자 형식이 올바르지 않습니다.")
                    reference_ids: list[int] = []
                    for rid in raw_reference_ids:
                        try:
                            reference_ids.append(int(rid))
                        except (TypeError, ValueError):
                            raise ApiError(400, "출장결과 참조자 사용자 ID가 올바르지 않습니다.")
                    result_text = str(body.get("trip_result") or "").strip()
                    doc_result, warnings = submit_business_trip_result_document(
                        conn,
                        source_document_id=int(m_trip_result_submit.group(1)),
                        actor=user,
                        approver_ids=approver_ids,
                        reference_ids=reference_ids,
                        trip_result_text=result_text,
                    )
                    json_response(self, 201, {"document": doc_result, "warnings": warnings})
                    return

                m_education_result_submit = re.fullmatch(r"/api/educations/(\d+)/result", path)
                if method == "POST" and m_education_result_submit:
                    body = parse_json_body(self)
                    raw_approver_ids = body.get("approver_ids")
                    # Backward compatibility for older clients that still send reviewer_id.
                    if raw_approver_ids is None:
                        legacy_reviewer_id = body.get("reviewer_id")
                        raw_approver_ids = [legacy_reviewer_id] if legacy_reviewer_id not in (None, "", 0, "0") else []
                    if not isinstance(raw_approver_ids, list):
                        raise ApiError(400, "교육결과 결재선 형식이 올바르지 않습니다.")
                    approver_ids: list[int] = []
                    for uid in raw_approver_ids:
                        try:
                            approver_ids.append(int(uid))
                        except (TypeError, ValueError):
                            raise ApiError(400, "교육결과 결재선 사용자 ID가 올바르지 않습니다.")

                    raw_reference_ids = body.get("reference_ids") or []
                    if not isinstance(raw_reference_ids, list):
                        raise ApiError(400, "교육결과 참조자 형식이 올바르지 않습니다.")
                    reference_ids: list[int] = []
                    for rid in raw_reference_ids:
                        try:
                            reference_ids.append(int(rid))
                        except (TypeError, ValueError):
                            raise ApiError(400, "교육결과 참조자 사용자 ID가 올바르지 않습니다.")
                    education_content_text = str(body.get("education_content") or "").strip()
                    education_apply_point_text = str(body.get("education_apply_point") or "").strip()
                    doc_result, warnings = submit_education_result_document(
                        conn,
                        source_document_id=int(m_education_result_submit.group(1)),
                        actor=user,
                        approver_ids=approver_ids,
                        reference_ids=reference_ids,
                        education_content_text=education_content_text,
                        education_apply_point_text=education_apply_point_text,
                    )
                    json_response(self, 201, {"document": doc_result, "warnings": warnings})
                    return

                if method == "GET" and path == "/api/overtimes/export":
                    try:
                        year = int((query.get("year", [""])[0] or "").strip()) if (query.get("year", [""])[0] or "").strip() else None
                    except Exception:
                        year = None
                    try:
                        month = int((query.get("month", [""])[0] or "").strip()) if (query.get("month", [""])[0] or "").strip() else None
                    except Exception:
                        month = None
                    payload = list_overtime_usages(conn, user, year, month)
                    selection = payload.get("selection") or {}
                    y = int(selection.get("year") or datetime.now().year)
                    m = int(selection.get("month") or datetime.now().month)
                    filename = f"overtime_{y}_{m:02d}.csv"
                    bytes_response(
                        self,
                        200,
                        build_overtime_csv_export(payload),
                        "text/csv; charset=utf-8",
                        filename=filename,
                        disposition="attachment",
                    )
                    return

                if method == "GET" and path == "/api/documents":
                    where: list[str] = []
                    params: list[Any] = []
                    mine = query.get("mine", ["0"])[0] == "1"
                    pending_me = query.get("pending_me", ["0"])[0] == "1"
                    all_completed = query.get("all_completed", ["0"])[0] == "1"
                    statuses = [s.strip() for s in query.get("status", [""])[0].split(",") if s.strip()]
                    search = (query.get("search", [""])[0] or "").strip()
                    archived = query.get("archived", ["0"])[0] == "1"

                    if all_completed:
                        where.append("d.status IN ('approved', 'rejected')")
                    else:
                        if mine:
                            where.append("d.drafter_id = ?")
                            params.append(user["id"])
                        if pending_me:
                            where.append(
                                "("
                                "EXISTS (SELECT 1 FROM approval_steps s WHERE s.document_id=d.id AND s.approver_id=? AND s.status='pending') "
                                "OR (COALESCE(d.edit_request_status,'none')='pending' AND d.edit_request_reviewer_id=?) "
                                "OR (COALESCE(d.delete_request_status,'none')='pending' AND d.delete_request_reviewer_id=?)"
                                ")"
                            )
                            params.extend([user["id"], user["id"], user["id"]])
                    if archived:
                        if user["role"] != "admin":
                            raise ApiError(403, "보관 문서 조회는 관리자만 가능합니다.")
                        where.append("COALESCE(d.is_deleted,0)=1")
                    else:
                        where.append("COALESCE(d.is_deleted,0)=0")
                    if statuses and not all_completed:
                        ph = ",".join("?" for _ in statuses)
                        where.append(f"d.status IN ({ph})")
                        params.extend(statuses)
                    if search:
                        where.append("(d.title LIKE ? OR d.content LIKE ? OR d.external_doc_url LIKE ?)")
                        params.extend([f"%{search}%", f"%{search}%", f"%{search}%"])

                    where_sql = "WHERE " + " AND ".join(where) if where else ""
                    rows = conn.execute(
                        f"""
                        SELECT d.*, u.full_name drafter_name, u.department drafter_department
                        FROM documents d JOIN users u ON u.id = d.drafter_id
                        {where_sql}
                        ORDER BY d.created_at DESC LIMIT 300
                        """,
                        params,
                    ).fetchall()
                    out_docs: list[dict[str, Any]] = []
                    for r in rows:
                        can_open = can_view_document_row(conn, r, user)
                        if (not all_completed) and (not can_open):
                            continue
                        item = dict_document(conn, r, include_detail=False)
                        item["can_open"] = can_open
                        out_docs.append(item)
                    json_response(self, 200, {"documents": out_docs})
                    return

                m_print_snapshot = re.fullmatch(r"/api/documents/(\d+)/print-snapshot", path)
                if method == "POST" and m_print_snapshot:
                    document_id = int(m_print_snapshot.group(1))
                    body = parse_json_body(self)
                    doc = fetch_document(conn, document_id)
                    if not doc:
                        raise ApiError(404, "문서를 찾을 수 없습니다.")
                    if ("is_deleted" in doc.keys() and doc["is_deleted"]) and user["role"] != "admin":
                        raise ApiError(404, "문서를 찾을 수 없습니다.")
                    if not can_view_document_row(conn, doc, user):
                        raise ApiError(403, "해당 문서를 열람할 권한이 없습니다.")
                    if (doc["status"] or "") not in {"approved", "rejected"}:
                        raise ApiError(409, "완료 문서만 인쇄용 파일을 생성할 수 있습니다.")
                    force_snapshot = bool(body.get("force")) if isinstance(body, dict) else False
                    snapshot = ensure_document_print_snapshot(conn, document_id, force=force_snapshot)
                    if not snapshot:
                        raise ApiError(400, "인쇄용 파일을 생성할 수 없는 문서입니다.")
                    refreshed = fetch_document(conn, document_id) or doc
                    out_doc = dict_document(conn, refreshed, include_detail=False)
                    if isinstance(out_doc.get("external_doc"), dict):
                        out_doc["external_doc"]["can_open_original"] = can_open_original_doc_for_viewer(refreshed, user)
                    json_response(self, 200, {"snapshot": snapshot, "document": out_doc})
                    return

                m_print_binary = re.fullmatch(r"/api/documents/(\d+)/print-binary", path)
                if method == "POST" and m_print_binary:
                    document_id = int(m_print_binary.group(1))
                    body = parse_json_body_optional(self)
                    doc = fetch_document(conn, document_id)
                    if not doc:
                        raise ApiError(404, "문서를 찾을 수 없습니다.")
                    if ("is_deleted" in doc.keys() and doc["is_deleted"]) and user["role"] != "admin":
                        raise ApiError(404, "문서를 찾을 수 없습니다.")
                    if not can_view_document_row(conn, doc, user):
                        raise ApiError(403, "해당 문서를 열람할 권한이 없습니다.")
                    if (doc["status"] or "") not in {"approved", "rejected"}:
                        raise ApiError(409, "완료 문서만 인쇄할 수 있습니다.")
                    force_snapshot = bool(body.get("force")) if isinstance(body, dict) else False
                    snapshot = ensure_document_print_snapshot(conn, document_id, force=force_snapshot)
                    if not snapshot or not snapshot.get("file_id"):
                        raise ApiError(400, "인쇄용 PDF 파일을 생성하지 못했습니다.")
                    file_data = read_google_drive_file_bytes(str(snapshot["file_id"]))
                    bytes_response(
                        self,
                        200,
                        file_data["bytes"],
                        file_data.get("mime_type") or "application/pdf",
                        file_data.get("name") or f"document_{document_id}.pdf",
                    )
                    return

                m_detail = re.fullmatch(r"/api/documents/(\d+)", path)
                if method == "GET" and m_detail:
                    doc = fetch_document(conn, int(m_detail.group(1)))
                    if not doc:
                        raise ApiError(404, "문서를 찾을 수 없습니다.")
                    if ("is_deleted" in doc.keys() and doc["is_deleted"]) and user["role"] != "admin":
                        raise ApiError(404, "문서를 찾을 수 없습니다.")
                    if not can_view_document_row(conn, doc, user):
                        raise ApiError(403, "해당 문서를 열람할 권한이 없습니다.")
                    out_doc = dict_document(conn, doc, include_detail=True)
                    if isinstance(out_doc.get("external_doc"), dict):
                        out_doc["external_doc"]["can_open_original"] = can_open_original_doc_for_viewer(doc, user)
                    json_response(self, 200, {"document": out_doc})
                    return

                if method == "POST" and path == "/api/documents":
                    body = parse_json_body(self)
                    warnings: list[str] = []
                    title = (body.get("title") or "").strip()
                    template_type = (body.get("template_type") or "general").strip() or "general"
                    recipient_text = (body.get("recipient_text") or "").strip()
                    visibility_scope = (body.get("visibility_scope") or "private").strip() or "private"
                    issue_department = (body.get("issue_department") or user["department"] or "").strip()
                    issue_year = str(body.get("issue_year") or "").strip()
                    leave_type = (body.get("leave_type") or "").strip()
                    leave_start_date = validate_leave_datetime(body.get("leave_start_date"), "휴가 시작일시")
                    leave_end_date = validate_leave_datetime(body.get("leave_end_date"), "휴가 종료일시")
                    leave_days = parse_leave_days_value(body.get("leave_days"))
                    leave_reason = (body.get("leave_reason") or "").strip()
                    leave_substitute_name = (body.get("leave_substitute_name") or "").strip()
                    leave_substitute_work = (body.get("leave_substitute_work") or "").strip()
                    overtime_type = (body.get("overtime_type") or "").strip()
                    overtime_start_date = validate_leave_datetime(body.get("overtime_start_date"), "연장근로 시작일시")
                    overtime_end_date = validate_leave_datetime(body.get("overtime_end_date"), "연장근로 종료일시")
                    overtime_hours = parse_overtime_hours_value(body.get("overtime_hours"))
                    overtime_content = (body.get("overtime_content") or "").strip()
                    overtime_etc = (body.get("overtime_etc") or "").strip()
                    trip_department = (body.get("trip_department") or "").strip()
                    trip_job_title = (body.get("trip_job_title") or "").strip()
                    trip_name = (body.get("trip_name") or "").strip()
                    trip_type = (body.get("trip_type") or "").strip()
                    trip_destination = (body.get("trip_destination") or "").strip()
                    trip_start_date = validate_leave_datetime(body.get("trip_start_date"), "출장 시작일시")
                    trip_end_date = validate_leave_datetime(body.get("trip_end_date"), "출장 종료일시")
                    trip_transportation = (body.get("trip_transportation") or "").strip()
                    trip_expense = (body.get("trip_expense") or "").strip()
                    trip_purpose = (body.get("trip_purpose") or "").strip()
                    education_department = (body.get("education_department") or "").strip()
                    education_job_title = (body.get("education_job_title") or "").strip()
                    education_name = (body.get("education_name") or "").strip()
                    education_title = (body.get("education_title") or "").strip()
                    education_category = (body.get("education_category") or "").strip()
                    education_provider = (body.get("education_provider") or "").strip()
                    education_location = (body.get("education_location") or "").strip()
                    education_start_date = validate_leave_datetime(body.get("education_start_date"), "교육 시작일시")
                    education_end_date = validate_leave_datetime(body.get("education_end_date"), "교육 종료일시")
                    education_purpose = (body.get("education_purpose") or "").strip()
                    education_tuition_detail = (body.get("education_tuition_detail") or "").strip()
                    education_tuition_amount = parse_education_money_value(body.get("education_tuition_amount"))
                    education_material_detail = (body.get("education_material_detail") or "").strip()
                    education_material_amount = parse_education_money_value(body.get("education_material_amount"))
                    education_transport_detail = (body.get("education_transport_detail") or "").strip()
                    education_transport_amount = parse_education_money_value(body.get("education_transport_amount"))
                    education_other_detail = (body.get("education_other_detail") or "").strip()
                    education_other_amount = parse_education_money_value(body.get("education_other_amount"))
                    education_budget_subject = (body.get("education_budget_subject") or "").strip()
                    education_funding_source = (body.get("education_funding_source") or "").strip()
                    education_payment_method = (body.get("education_payment_method") or "").strip()
                    education_support_budget = parse_education_money_value(body.get("education_support_budget"))
                    education_used_budget = parse_education_money_value(body.get("education_used_budget"))
                    education_remain_budget = parse_education_money_value(body.get("education_remain_budget"))
                    education_companion = (body.get("education_companion") or "").strip()
                    education_ordered = (body.get("education_ordered") or "").strip()
                    education_suggestion = (body.get("education_suggestion") or "").strip()
                    content = (body.get("content") or "").strip()
                    editor_provider = EDITOR_PROVIDER_GOOGLE_DOCS
                    external_doc_url_input = (body.get("external_doc_url") or "").strip()
                    external_doc_id_input = (body.get("external_doc_id") or "").strip()
                    priority = (body.get("priority") or "normal").strip() or "normal"
                    due_date = validate_date(body.get("due_date"), "기한일")
                    submit_now = bool(body.get("submit"))
                    if not title:
                        raise ApiError(400, "문서 제목을 입력해 주세요.")
                    if visibility_scope not in DOC_VISIBILITY_VALUES:
                        raise ApiError(400, "문서 공개 여부 값이 올바르지 않습니다.")
                    if issue_department and issue_department not in ISSUE_DEPARTMENT_OPTIONS:
                        raise ApiError(400, "시행부서 값이 올바르지 않습니다.")
                    if issue_year and not re.fullmatch(r"\d{4}", issue_year):
                        raise ApiError(400, "시행년도는 4자리 숫자로 입력해 주세요.")
                    leave_start_date, leave_end_date, leave_days = validate_leave_request_fields(
                        template_type,
                        leave_start_date,
                        leave_end_date,
                        leave_days,
                    )
                    if template_type == LEAVE_TEMPLATE_TYPE and not leave_type:
                        raise ApiError(400, "휴가계 문서는 휴가형태를 입력해 주세요.")
                    overtime_start_date, overtime_end_date, overtime_hours = validate_overtime_request_fields(
                        template_type,
                        overtime_start_date,
                        overtime_end_date,
                        overtime_hours,
                    )
                    if template_type == OVERTIME_TEMPLATE_TYPE and not overtime_type:
                        raise ApiError(400, "연장근로 문서는 형태를 입력해 주세요.")
                    if template_type == OVERTIME_TEMPLATE_TYPE and not overtime_content:
                        raise ApiError(400, "연장근로 문서는 연장근로 내용을 입력해 주세요.")
                    trip_start_date, trip_end_date = validate_business_trip_request_fields(
                        template_type,
                        trip_start_date,
                        trip_end_date,
                    )
                    if template_type == BUSINESS_TRIP_TEMPLATE_TYPE and not trip_type:
                        raise ApiError(400, "출장신청서 문서는 출장종류를 입력해 주세요.")
                    if template_type == BUSINESS_TRIP_TEMPLATE_TYPE and not trip_destination:
                        raise ApiError(400, "출장신청서 문서는 출장지를 입력해 주세요.")
                    if template_type == BUSINESS_TRIP_TEMPLATE_TYPE and not trip_purpose:
                        raise ApiError(400, "출장신청서 문서는 출장목적을 입력해 주세요.")
                    education_start_date, education_end_date = validate_education_request_fields(
                        template_type,
                        education_start_date,
                        education_end_date,
                    )
                    if template_type == EDUCATION_TEMPLATE_TYPE:
                        if not education_department:
                            education_department = (user["department"] or "").strip() if "department" in user.keys() else ""
                        if not education_job_title:
                            education_job_title = (user["job_title"] or "").strip() if "job_title" in user.keys() else ""
                        if not education_name:
                            education_name = (user["full_name"] or "").strip() if "full_name" in user.keys() else ""
                        if not education_title:
                            raise ApiError(400, "교육신청서 문서는 교육명을 입력해 주세요.")
                        if not education_category:
                            raise ApiError(400, "교육신청서 문서는 교육분류를 입력해 주세요.")
                        if not education_provider:
                            raise ApiError(400, "교육신청서 문서는 교육기관을 입력해 주세요.")
                        if not education_location:
                            raise ApiError(400, "교육신청서 문서는 교육장소를 입력해 주세요.")
                        if not education_purpose:
                            raise ApiError(400, "교육신청서 문서는 교육목적을 입력해 주세요.")
                    external_doc_id: str | None = None
                    external_doc_url: str | None = None
                    external_doc_id = extract_google_doc_id(external_doc_id_input or external_doc_url_input)
                    if not external_doc_id:
                        raise ApiError(400, "Google Docs 링크 또는 문서 ID를 입력해 주세요.")
                    urls = build_google_doc_urls(external_doc_id)
                    external_doc_url = urls["edit_url"]
                    if not content:
                        content = f"[Google Docs 문서] {external_doc_url}"

                    approver_ids = body.get("approver_ids") or []
                    if not isinstance(approver_ids, list) or not approver_ids:
                        raise ApiError(400, "결재선을 최소 1명 이상 지정해 주세요.")
                    normalized: list[int] = []
                    for uid in approver_ids:
                        try:
                            iid = int(uid)
                        except (TypeError, ValueError) as exc:
                            raise ApiError(400, "결재선 사용자 ID가 잘못되었습니다.") from exc
                        if iid not in normalized:
                            normalized.append(iid)
                    total_slots_for_template = approval_template_total_slots_for_type(template_type)
                    max_steps_for_template = max_approver_steps_for_type(template_type)
                    if len(normalized) > max_steps_for_template:
                        raise ApiError(400, f"현재 템플릿은 결재자 최대 {max_steps_for_template}명까지 지원합니다. (기안자 포함 {total_slots_for_template}칸)")

                    refs = body.get("reference_ids") or []
                    ref_ids: list[int] = []
                    if isinstance(refs, list):
                        for rid in refs:
                            try:
                                ref_ids.append(int(rid))
                            except (TypeError, ValueError):
                                pass
                    attachments_raw = body.get("attachments") or []
                    attachments: list[dict[str, Any]] = []
                    if isinstance(attachments_raw, list):
                        for item in attachments_raw:
                            if not isinstance(item, dict):
                                continue
                            file_id = str(item.get("file_id") or item.get("id") or "").strip()
                            name = str(item.get("name") or "").strip()
                            if not file_id or not name:
                                continue
                            attachments.append(
                                {
                                    "file_id": file_id,
                                    "name": name,
                                    "mime_type": str(item.get("mime_type") or item.get("mimeType") or ""),
                                    "size": int(item.get("size") or 0),
                                    "web_view_url": str(item.get("web_view_url") or item.get("webViewUrl") or ""),
                                }
                            )

                    existing = {
                        r["id"]
                        for r in conn.execute(
                            "SELECT id FROM users WHERE id IN ({})".format(",".join("?" for _ in normalized)),
                            normalized,
                        ).fetchall()
                    }
                    missing = [uid for uid in normalized if uid not in existing]
                    if missing:
                        raise ApiError(400, f"존재하지 않는 결재자 ID가 있습니다: {missing}")

                    now = now_ts()
                    issue_date_for_doc = due_date or today_str()
                    effective_issue_department = issue_department or user["department"] or "기타"
                    effective_issue_year = issue_year or str(date.fromisoformat(issue_date_for_doc).year)
                    issue_code = allocate_department_issue_code(
                        conn,
                        effective_issue_department,
                        issue_date_for_doc,
                        effective_issue_year,
                    )
                    cur = conn.execute(
                        """
                        INSERT INTO documents (title, template_type, content, editor_provider, external_doc_id, external_doc_url, status, priority, due_date, leave_type, leave_start_date, leave_end_date, leave_days, leave_reason, leave_substitute_name, leave_substitute_work, overtime_type, overtime_start_date, overtime_end_date, overtime_hours, overtime_content, overtime_etc, trip_department, trip_job_title, trip_name, trip_type, trip_destination, trip_start_date, trip_end_date, trip_transportation, trip_expense, trip_purpose, education_department, education_job_title, education_name, education_title, education_category, education_provider, education_location, education_start_date, education_end_date, education_purpose, education_tuition_detail, education_tuition_amount, education_material_detail, education_material_amount, education_transport_detail, education_transport_amount, education_other_detail, education_other_amount, education_budget_subject, education_funding_source, education_payment_method, education_support_budget, education_used_budget, education_remain_budget, education_companion, education_ordered, education_suggestion, issue_code, issue_department, issue_year, recipient_text, visibility_scope, attachments_json, drafter_id, submitted_at, completed_at, reference_ids, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, 'draft', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?)
                        """,
                        (
                            title,
                            template_type,
                            content,
                            editor_provider,
                            external_doc_id,
                            external_doc_url,
                            priority,
                            due_date,
                            leave_type or None,
                            leave_start_date,
                            leave_end_date,
                            leave_days,
                            leave_reason or None,
                            leave_substitute_name or None,
                            leave_substitute_work or None,
                            overtime_type or None,
                            overtime_start_date,
                            overtime_end_date,
                            overtime_hours,
                            overtime_content or None,
                            overtime_etc or None,
                            trip_department or None,
                            trip_job_title or None,
                            trip_name or None,
                            trip_type or None,
                            trip_destination or None,
                            trip_start_date,
                            trip_end_date,
                            trip_transportation or None,
                            trip_expense or None,
                            trip_purpose or None,
                            education_department or None,
                            education_job_title or None,
                            education_name or None,
                            education_title or None,
                            education_category or None,
                            education_provider or None,
                            education_location or None,
                            education_start_date,
                            education_end_date,
                            education_purpose or None,
                            education_tuition_detail or None,
                            education_tuition_amount,
                            education_material_detail or None,
                            education_material_amount,
                            education_transport_detail or None,
                            education_transport_amount,
                            education_other_detail or None,
                            education_other_amount,
                            education_budget_subject or None,
                            education_funding_source or None,
                            education_payment_method or None,
                            education_support_budget,
                            education_used_budget,
                            education_remain_budget,
                            education_companion or None,
                            education_ordered or None,
                            education_suggestion or None,
                            issue_code,
                            effective_issue_department,
                            effective_issue_year,
                            recipient_text,
                            visibility_scope,
                            json.dumps(attachments, ensure_ascii=False),
                            user["id"],
                            json.dumps(ref_ids, ensure_ascii=False),
                            now,
                            now,
                        ),
                    )
                    doc_id = cur.lastrowid
                    if not doc_id:
                        raise ApiError(500, "문서 생성에 실패했습니다.")

                    for idx, aid in enumerate(normalized, start=1):
                        conn.execute("INSERT INTO approval_steps (document_id, step_order, approver_id, status) VALUES (?, ?, ?, 'waiting')", (doc_id, idx, aid))

                    # Fill draft metadata into the copied Google Docs template (title / issue date / issue code).
                    leave_period = format_leave_period_text(leave_start_date, leave_end_date)
                    overtime_time_text = format_overtime_period_text(overtime_start_date, overtime_end_date, overtime_hours)
                    trip_period = format_trip_period_text(trip_start_date, trip_end_date)
                    education_period = format_education_period_text(education_start_date, education_end_date)
                    user_total_leave = float(user["total_leave"] or 0) if "total_leave" in user.keys() else 0.0
                    user_used_leave = float(user["used_leave"] or 0) if "used_leave" in user.keys() else 0.0
                    req_leave_days = float(leave_days or 0) if template_type == LEAVE_TEMPLATE_TYPE else 0.0
                    projected_remain = max(0.0, user_total_leave - user_used_leave - req_leave_days)
                    try:
                        fill_data = populate_google_doc_draft_fields(
                            external_doc_id,
                            title=title,
                            issue_date=issue_date_for_doc,
                            issue_code=issue_code,
                            doc_form_type=document_form_type_label(template_type),
                            recipient_text=recipient_text,
                            leave_dept=((user["department"] or "") if "department" in user.keys() else "") if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_name=((user["full_name"] or "") if "full_name" in user.keys() else "") if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_job_title=((user["job_title"] or "") if "job_title" in user.keys() else "") if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_type=leave_type if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_period=leave_period if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_total_days=format_leave_number(user_total_leave) if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_used_days=format_leave_number(req_leave_days) if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_remain_days=format_leave_number(projected_remain) if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_reason=leave_reason if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_substitute_name=leave_substitute_name if template_type == LEAVE_TEMPLATE_TYPE else "",
                            leave_substitute_work=leave_substitute_work if template_type == LEAVE_TEMPLATE_TYPE else "",
                            overtime_dept=((user["department"] or "") if "department" in user.keys() else "") if template_type == OVERTIME_TEMPLATE_TYPE else "",
                            overtime_name=((user["full_name"] or "") if "full_name" in user.keys() else "") if template_type == OVERTIME_TEMPLATE_TYPE else "",
                            overtime_job_title=((user["job_title"] or "") if "job_title" in user.keys() else "") if template_type == OVERTIME_TEMPLATE_TYPE else "",
                            overtime_type=overtime_type if template_type == OVERTIME_TEMPLATE_TYPE else "",
                            overtime_time=overtime_time_text if template_type == OVERTIME_TEMPLATE_TYPE else "",
                            overtime_hours=format_leave_number(overtime_hours) if template_type == OVERTIME_TEMPLATE_TYPE else "",
                            overtime_content=overtime_content if template_type == OVERTIME_TEMPLATE_TYPE else "",
                            overtime_etc=overtime_etc if template_type == OVERTIME_TEMPLATE_TYPE else "",
                            trip_department=(trip_department or ((user["department"] or "") if "department" in user.keys() else "")) if template_type == BUSINESS_TRIP_TEMPLATE_TYPE else "",
                            trip_job_title=(trip_job_title or ((user["job_title"] or "") if "job_title" in user.keys() else "")) if template_type == BUSINESS_TRIP_TEMPLATE_TYPE else "",
                            trip_name=(trip_name or ((user["full_name"] or "") if "full_name" in user.keys() else "")) if template_type == BUSINESS_TRIP_TEMPLATE_TYPE else "",
                            trip_type=trip_type if template_type == BUSINESS_TRIP_TEMPLATE_TYPE else "",
                            trip_destination=trip_destination if template_type == BUSINESS_TRIP_TEMPLATE_TYPE else "",
                            trip_period=trip_period if template_type == BUSINESS_TRIP_TEMPLATE_TYPE else "",
                            trip_transportation=trip_transportation if template_type == BUSINESS_TRIP_TEMPLATE_TYPE else "",
                            trip_expense=trip_expense if template_type == BUSINESS_TRIP_TEMPLATE_TYPE else "",
                            trip_purpose=trip_purpose if template_type == BUSINESS_TRIP_TEMPLATE_TYPE else "",
                            education_department=education_department if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_job_title=education_job_title if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_name=education_name if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_title=education_title if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_category=education_category if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_provider=education_provider if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_location=education_location if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_period=education_period if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_purpose=education_purpose if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_tuition_detail=education_tuition_detail if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_tuition_amount=format_leave_number(education_tuition_amount) if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_material_detail=education_material_detail if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_material_amount=format_leave_number(education_material_amount) if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_transport_detail=education_transport_detail if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_transport_amount=format_leave_number(education_transport_amount) if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_other_detail=education_other_detail if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_other_amount=format_leave_number(education_other_amount) if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_budget_subject=education_budget_subject if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_funding_source=education_funding_source if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_payment_method=education_payment_method if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_support_budget=format_leave_number(education_support_budget) if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_used_budget=format_leave_number(education_used_budget) if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_remain_budget=format_leave_number(education_remain_budget) if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_companion=education_companion if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_ordered=education_ordered if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            education_suggestion=education_suggestion if template_type == EDUCATION_TEMPLATE_TYPE else "",
                            draft_date=issue_date_for_doc,
                        )
                        if template_type == BUSINESS_TRIP_TEMPLATE_TYPE:
                            result_map = fill_data.get("result") if isinstance(fill_data.get("result"), dict) else {}
                            trip_keys = [
                                "trip_department",
                                "trip_job_title",
                                "trip_name",
                                "trip_type",
                                "trip_destination",
                                "trip_period",
                                "trip_transportation",
                                "trip_expense",
                                "trip_purpose",
                            ]
                            updated_count = 0
                            for key in trip_keys:
                                item = result_map.get(key) if isinstance(result_map, dict) else None
                                if isinstance(item, dict) and item.get("updated"):
                                    updated_count += 1
                            if updated_count == 0:
                                warnings.append(
                                    "출장 문서 자동입력 결과가 없습니다. Apps Script 최신 버전 재배포 및 토큰명(TRIP_*)을 확인해 주세요."
                                )
                        if template_type == EDUCATION_TEMPLATE_TYPE:
                            result_map = fill_data.get("result") if isinstance(fill_data.get("result"), dict) else {}
                            education_keys = [
                                "education_department",
                                "education_job_title",
                                "education_name",
                                "education_title",
                                "education_category",
                                "education_provider",
                                "education_location",
                                "education_period",
                                "education_purpose",
                                "education_tuition_detail",
                                "education_tuition_amount",
                                "education_material_detail",
                                "education_material_amount",
                                "education_transport_detail",
                                "education_transport_amount",
                                "education_other_detail",
                                "education_other_amount",
                                "education_budget_subject",
                                "education_funding_source",
                                "education_payment_method",
                                "education_support_budget",
                                "education_used_budget",
                                "education_remain_budget",
                                "education_companion",
                                "education_ordered",
                                "education_suggestion",
                            ]
                            updated_count = 0
                            for key in education_keys:
                                item = result_map.get(key) if isinstance(result_map, dict) else None
                                if isinstance(item, dict) and item.get("updated"):
                                    updated_count += 1
                            if updated_count == 0:
                                warnings.append(
                                    "교육 문서 자동입력 결과가 없습니다. Apps Script 최신 버전 재배포 및 토큰명(EDU_*)을 확인해 주세요."
                                )
                    except Exception as e:
                        # Best effort: draft creation/submission should continue even if template auto-fill times out.
                        print(f"[populate_google_doc_draft_fields] WARN doc={doc_id}: {type(e).__name__}: {e}")
                        warnings.append(f"문서 자동입력 경고: {e}")

                    if template_type == EDUCATION_TEMPLATE_TYPE and external_doc_id:
                        try:
                            move_google_drive_file(external_doc_id, EDUCATION_DRAFT_OUTPUT_FOLDER_ID)
                        except Exception as e:
                            print(f"[move_google_drive_file] WARN education doc={doc_id}: {type(e).__name__}: {e}")
                            warnings.append(f"교육 문서 저장 폴더 이동 경고: {e}")

                    if submit_now:
                        result = submit_document(conn, int(doc_id), user["id"])
                    else:
                        created = fetch_document(conn, int(doc_id))
                        if not created:
                            raise ApiError(500, "생성한 문서를 찾지 못했습니다.")
                        result = dict_document(conn, created, include_detail=True)

                    json_response(self, 201, {"document": result, "warnings": warnings})
                    return

                m_submit = re.fullmatch(r"/api/documents/(\d+)/submit", path)
                if method == "POST" and m_submit:
                    result = submit_document(conn, int(m_submit.group(1)), user["id"])
                    json_response(self, 200, {"document": result})
                    return

                m_resubmit = re.fullmatch(r"/api/documents/(\d+)/resubmit", path)
                if method == "POST" and m_resubmit:
                    result = resubmit_rejected_document(conn, int(m_resubmit.group(1)), user["id"])
                    json_response(self, 200, {"document": result})
                    return

                m_action = re.fullmatch(r"/api/documents/(\d+)/actions", path)
                if method == "POST" and m_action:
                    body = parse_json_body(self)
                    action = (body.get("action") or "").strip()
                    comment = (body.get("comment") or "").strip()
                    result = apply_approval_action(conn, int(m_action.group(1)), user, action, comment)
                    json_response(self, 200, {"document": result})
                    return

                m_edit_request = re.fullmatch(r"/api/documents/(\d+)/edit-request", path)
                if method == "POST" and m_edit_request:
                    body = parse_json_body(self)
                    reason = str(body.get("reason") or "").strip()
                    try:
                        reviewer_id = int(body.get("reviewer_id") or 0)
                    except (TypeError, ValueError):
                        raise ApiError(400, "수정요청 결재자 ID가 올바르지 않습니다.")
                    result = request_completed_document_edit(conn, int(m_edit_request.group(1)), user, reason, reviewer_id)
                    json_response(self, 200, {"document": result})
                    return

                m_edit_request_decision = re.fullmatch(r"/api/documents/(\d+)/edit-request/decision", path)
                if method == "POST" and m_edit_request_decision:
                    body = parse_json_body(self)
                    approve = bool(body.get("approve"))
                    result = decide_completed_document_edit_request(conn, int(m_edit_request_decision.group(1)), user, approve)
                    json_response(self, 200, {"document": result})
                    return

                m_edit_request_complete = re.fullmatch(r"/api/documents/(\d+)/edit-request/complete", path)
                if method == "POST" and m_edit_request_complete:
                    result = complete_completed_document_edit_request(conn, int(m_edit_request_complete.group(1)), user)
                    json_response(self, 200, {"document": result})
                    return

                m_delete_request = re.fullmatch(r"/api/documents/(\d+)/delete-request", path)
                if method == "POST" and m_delete_request:
                    body = parse_json_body(self)
                    reason = str(body.get("reason") or "").strip()
                    try:
                        reviewer_id = int(body.get("reviewer_id") or 0)
                    except (TypeError, ValueError):
                        raise ApiError(400, "삭제요청 결재자 ID가 올바르지 않습니다.")
                    result = request_completed_document_delete(conn, int(m_delete_request.group(1)), user, reason, reviewer_id)
                    json_response(self, 200, {"document": result})
                    return

                m_delete_request_decision = re.fullmatch(r"/api/documents/(\d+)/delete-request/decision", path)
                if method == "POST" and m_delete_request_decision:
                    body = parse_json_body(self)
                    approve = bool(body.get("approve"))
                    result = decide_completed_document_delete_request(conn, int(m_delete_request_decision.group(1)), user, approve)
                    json_response(self, 200, {"document": result})
                    return

                m_doc_delete = re.fullmatch(r"/api/documents/(\d+)/delete", path)
                if method == "POST" and m_doc_delete:
                    if user["role"] != "admin":
                        raise ApiError(403, "문서 삭제는 관리자만 가능합니다.")
                    document_id = int(m_doc_delete.group(1))
                    body = parse_json_body(self)
                    mode = (body.get("mode") or "archive").strip().lower()
                    doc_row = conn.execute("SELECT id, title, external_doc_id, attachments_json, print_snapshot_file_id FROM documents WHERE id=?", (document_id,)).fetchone()
                    if not doc_row:
                        raise ApiError(404, "문서를 찾을 수 없습니다.")
                    if mode not in {"archive", "purge"}:
                        raise ApiError(400, "삭제 모드는 archive 또는 purge 여야 합니다.")

                    linked_deleted_ids = apply_delete_mode_for_linked_trip_result_documents(
                        conn,
                        source_document_id=document_id,
                        actor_id=int(user["id"]),
                        mode=mode,
                    )
                    apply_delete_mode_to_document(
                        conn,
                        document_id=document_id,
                        doc_row=doc_row,
                        actor_id=int(user["id"]),
                        mode=mode,
                    )
                    json_response(
                        self,
                        200,
                        {
                            "ok": True,
                            "mode": mode,
                            "document_id": document_id,
                            "linked_trip_result_document_ids": linked_deleted_ids,
                        },
                    )
                    return

                m_doc_restore = re.fullmatch(r"/api/documents/(\d+)/restore", path)
                if method == "POST" and m_doc_restore:
                    if user["role"] != "admin":
                        raise ApiError(403, "문서 복원은 관리자만 가능합니다.")
                    document_id = int(m_doc_restore.group(1))
                    row = conn.execute("SELECT id, is_deleted FROM documents WHERE id=?", (document_id,)).fetchone()
                    if not row:
                        raise ApiError(404, "문서를 찾을 수 없습니다.")
                    if not int(row["is_deleted"] or 0):
                        raise ApiError(400, "보관삭제된 문서가 아닙니다.")
                    conn.execute(
                        "UPDATE documents SET is_deleted=0, deleted_at=NULL, deleted_by=NULL, updated_at=? WHERE id=?",
                        (now_ts(), document_id),
                    )
                    sync_leave_usage_for_document(conn, document_id)
                    json_response(self, 200, {"ok": True, "mode": "restore", "document_id": document_id})
                    return

                if method == "GET" and path == "/api/approvals/pending":
                    rows = conn.execute(
                        """
                        SELECT d.*, u.full_name drafter_name, u.department drafter_department
                        FROM documents d JOIN users u ON u.id = d.drafter_id
                        WHERE COALESCE(d.is_deleted,0)=0 AND EXISTS (
                            SELECT 1 FROM approval_steps s
                            WHERE s.document_id=d.id AND s.approver_id=? AND s.status='pending'
                        )
                        ORDER BY d.updated_at DESC
                        """,
                        (user["id"],),
                    ).fetchall()
                    visible_rows = [r for r in rows if can_view_document_row(conn, r, user)]
                    json_response(self, 200, {"documents": [dict_document(conn, r, include_detail=False) for r in visible_rows]})
                    return

                if method == "GET" and path == "/api/notices":
                    rows = conn.execute(
                        """
                        SELECT n.id, n.title, n.content, n.pinned, n.created_at, u.full_name author_name
                        FROM notices n JOIN users u ON u.id = n.author_id
                        ORDER BY n.pinned DESC, n.created_at DESC LIMIT 200
                        """
                    ).fetchall()
                    notices = [
                        {
                            "id": r["id"],
                            "title": r["title"],
                            "content": r["content"],
                            "pinned": bool(r["pinned"]),
                            "created_at": r["created_at"],
                            "author_name": r["author_name"],
                        }
                        for r in rows
                    ]
                    json_response(self, 200, {"notices": notices})
                    return

                if method == "POST" and path == "/api/notices":
                    if user["role"] != "admin":
                        raise ApiError(403, "공지 등록은 관리자만 가능합니다.")
                    body = parse_json_body(self)
                    title = (body.get("title") or "").strip()
                    content = (body.get("content") or "").strip()
                    pinned = 1 if body.get("pinned") else 0
                    if not title or not content:
                        raise ApiError(400, "공지 제목과 내용을 입력해 주세요.")
                    conn.execute(
                        "INSERT INTO notices (title, content, author_id, pinned, created_at) VALUES (?, ?, ?, ?, ?)",
                        (title, content, user["id"], pinned, now_ts()),
                    )
                    json_response(self, 201, {"ok": True})
                    return

                if method == "GET" and path == "/api/schedules":
                    start = query.get("start", [today_str()])[0]
                    end = query.get("end", [(date.today() + timedelta(days=30)).isoformat()])[0]
                    validate_date(start, "시작일")
                    validate_date(end, "종료일")
                    rows = conn.execute(
                        """
                        SELECT s.id, s.title, s.start_date, s.end_date, s.event_type, s.resource_name,
                               s.status, s.created_at, u.id owner_id, u.full_name owner_name
                        FROM schedules s JOIN users u ON u.id = s.owner_id
                        WHERE s.start_date <= ? AND s.end_date >= ?
                        ORDER BY s.start_date ASC, s.created_at DESC
                        """,
                        (end, start),
                    ).fetchall()
                    schedules = [
                        {
                            "id": r["id"],
                            "title": r["title"],
                            "start_date": r["start_date"],
                            "end_date": r["end_date"],
                            "event_type": r["event_type"],
                            "resource_name": r["resource_name"],
                            "status": r["status"],
                            "created_at": r["created_at"],
                            "owner": {"id": r["owner_id"], "name": r["owner_name"]},
                        }
                        for r in rows
                    ]
                    json_response(self, 200, {"schedules": schedules})
                    return

                if method == "POST" and path == "/api/schedules":
                    body = parse_json_body(self)
                    title = (body.get("title") or "").strip()
                    event_type = (body.get("event_type") or "general").strip() or "general"
                    resource_name = (body.get("resource_name") or "").strip() or None
                    start_date = validate_date(body.get("start_date"), "시작일")
                    end_date = validate_date(body.get("end_date"), "종료일")
                    if not title:
                        raise ApiError(400, "일정 제목을 입력해 주세요.")
                    if not start_date or not end_date:
                        raise ApiError(400, "시작일과 종료일을 입력해 주세요.")
                    if start_date > end_date:
                        raise ApiError(400, "종료일은 시작일보다 이전일 수 없습니다.")
                    conn.execute(
                        "INSERT INTO schedules (title, start_date, end_date, event_type, owner_id, resource_name, status, created_at) VALUES (?, ?, ?, ?, ?, ?, 'active', ?)",
                        (title, start_date, end_date, event_type, user["id"], resource_name, now_ts()),
                    )
                    json_response(self, 201, {"ok": True})
                    return

                if method == "GET" and path == "/api/notifications":
                    rows = conn.execute(
                        "SELECT id, message, link, is_read, created_at FROM notifications WHERE user_id=? ORDER BY created_at DESC LIMIT 100",
                        (user["id"],),
                    ).fetchall()
                    json_response(self, 200, {
                        "notifications": [
                            {
                                "id": r["id"],
                                "message": r["message"],
                                "link": r["link"],
                                "is_read": bool(r["is_read"]),
                                "created_at": r["created_at"],
                            }
                            for r in rows
                        ]
                    })
                    return

                if method == "POST" and path == "/api/notifications/read":
                    body = parse_json_body(self)
                    ids = body.get("ids")
                    if isinstance(ids, list) and ids:
                        normalized: list[int] = []
                        for item in ids:
                            try:
                                normalized.append(int(item))
                            except (TypeError, ValueError):
                                pass
                        if normalized:
                            ph = ",".join("?" for _ in normalized)
                            conn.execute(f"UPDATE notifications SET is_read=1 WHERE user_id=? AND id IN ({ph})", [user["id"], *normalized])
                    else:
                        conn.execute("UPDATE notifications SET is_read=1 WHERE user_id=?", (user["id"],))
                    json_response(self, 200, {"ok": True})
                    return

                raise ApiError(404, "지원하지 않는 API 경로입니다.")

        except ApiError as exc:
            json_response(self, exc.status, {"error": exc.message})
        except sqlite3.DatabaseError as exc:
            json_response(self, 500, {"error": f"데이터베이스 오류: {exc}"})
        except Exception as exc:
            json_response(self, 500, {"error": f"서버 오류: {exc}"})

    def serve_static(self, path: str) -> None:
        requested = "/index.html" if path in ("", "/") else unquote(path)
        safe_path = Path(requested.lstrip("/"))
        target = (STATIC_DIR / safe_path).resolve()
        try:
            target.relative_to(STATIC_DIR)
        except ValueError:
            text_response(self, 403, "Forbidden", "text/plain; charset=utf-8")
            return

        if not target.exists() or target.is_dir():
            target = (STATIC_DIR / "index.html").resolve()
            if not target.exists():
                text_response(self, 404, "index.html not found", "text/plain; charset=utf-8")
                return

        mimes = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".json": "application/json; charset=utf-8",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".svg": "image/svg+xml",
            ".ico": "image/x-icon",
        }
        raw = target.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mimes.get(target.suffix.lower(), "application/octet-stream"))
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def log_message(self, format: str, *args: Any) -> None:
        print(f"[{now_ts()}] {self.address_string()} {format % args}")


def run(host: str, port: int) -> None:
    ensure_schema()
    server = ThreadingHTTPServer((host, port), ApprovalHandler)
    print(f"Web 전자결재 서버 시작: http://{host}:{port}")
    print("초기 계정: admin/admin123!, ceo/ceo123!, kim/kim123!, lee/lee123!, park/park123!")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버를 종료합니다.")
    finally:
        server.server_close()


def main() -> None:
    parser = argparse.ArgumentParser(description="KT BizOffice 스타일 웹 전자결재 시스템")
    parser.add_argument("--host", default=os.environ.get("HOST", "127.0.0.1"), help="서버 바인드 주소")
    parser.add_argument("--port", type=int, default=int(os.environ.get("PORT", "8080")), help="서버 포트")
    args = parser.parse_args()
    run(args.host, args.port)


if __name__ == "__main__":
    main()
