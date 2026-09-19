#!/usr/bin/env bash
# Run on camai after Noor pushes to GitHub (or after manual git pull).
set -euo pipefail
APP_DIR="/home/camai/ip-reports.propackhub.com/app"
cd "$APP_DIR"
git pull --ff-only origin main
./venv/bin/pip install -q -r requirements.txt
pm2 restart ip-reports-app
pm2 save
