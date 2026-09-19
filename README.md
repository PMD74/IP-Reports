# IP Reports — Noor dashboard (Flask)

## Run locally (dev)

```bash
cd ops/ip-reports/app
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
export NOOR_DB_HOST=127.0.0.1 NOOR_DB_PORT=5437 NOOR_DB_USER=warehouse NOOR_DB_PASSWORD=...
python run.py
```

## Production (camai)

- App: PM2 `ip-reports-app` → `127.0.0.1:3010`
- Data: PostgreSQL `noor-warehouse` schema `noor`
- Public: https://ip-reports.propackhub.com (nginx → Flask; `/ingest/` → ingest :3011)

## GitHub

Repository: https://github.com/PMD74/IP-Reports

Push updates:

```bash
git add -A && git commit -m "..." && git push origin main
```

After push, on camai:

```bash
cd /home/camai/ip-reports.propackhub.com/app
./deploy/git-pull-and-restart.sh
```

Or from Camille’s Mac: `ops/ip-reports/deploy-from-github.sh`
