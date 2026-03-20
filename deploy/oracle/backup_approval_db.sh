#!/usr/bin/env bash
set -euo pipefail

APP_DIR="/opt/eapproval"
BACKUP_DIR="/opt/eapproval/backups"
TS="$(date +%Y%m%d_%H%M%S)"
DB_FILE="$APP_DIR/data/approval.db"

mkdir -p "$BACKUP_DIR"

if [ ! -f "$DB_FILE" ]; then
  echo "[backup] DB not found: $DB_FILE"
  exit 1
fi

cp "$DB_FILE" "$BACKUP_DIR/approval_$TS.db"
find "$BACKUP_DIR" -type f -name 'approval_*.db' -mtime +14 -delete

echo "[backup] OK: $BACKUP_DIR/approval_$TS.db"
