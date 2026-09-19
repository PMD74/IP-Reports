#!/usr/bin/env python3
"""Run IP Reports Flask app (Postgres via pg_compat)."""
import os

import pg_compat  # noqa: F401 — must load before app imports mysql.connector

# Schema name used in INFORMATION_SCHEMA queries in app.py
os.environ.setdefault("NOOR_PG_SCHEMA", "noor")

from app import app  # noqa: E402

if __name__ == "__main__":
    port = int(os.environ.get("IP_REPORTS_PORT", "3010"))
    app.run(host="127.0.0.1", port=port, debug=False)
