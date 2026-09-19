#!/usr/bin/env bash
set -euo pipefail
APP_ROOT="/home/camai/ip-reports.propackhub.com/app"
ENV_FILE="/home/camai/.camai.env"
DOCKER_ENV="/home/camai/ip-reports.propackhub.com/docker/.env"
cd "$APP_ROOT"
# shellcheck disable=SC1090
[[ -f "$ENV_FILE" ]] && source "$ENV_FILE"
[[ -f "$DOCKER_ENV" ]] && source "$DOCKER_ENV"
export NOOR_DB_HOST="${NOOR_DB_HOST:-127.0.0.1}"
export NOOR_DB_PORT="${NOOR_DB_PORT:-5437}"
export NOOR_DB_USER="${NOOR_DB_USER:-warehouse}"
export NOOR_DB_PASSWORD="${NOOR_WAREHOUSE_POSTGRES_PASSWORD:-}"
export NOOR_PG_SCHEMA=noor
export IP_REPORTS_PORT="${IP_REPORTS_PORT:-3010}"
export FLASK_SECRET_KEY="${FLASK_SECRET_KEY:-${IP_REPORTS_FLASK_SECRET:-ip-reports-flask-secret}}"
exec ./venv/bin/python run.py
