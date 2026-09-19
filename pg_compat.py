"""MySQL connector API shim over PostgreSQL (noor-warehouse)."""
from __future__ import annotations

import os
import re
import sys

import psycopg2
import psycopg2.extras


class Error(Exception):
    pass


def _sql_mysql_to_pg(sql: str) -> str:
    if not sql:
        return sql
    s = re.sub(r"`([^`]+)`", r'"\1"', sql)
    # Unquoted MySQL table names -> noor."NAME"
    for tbl in (
        "HDSLS10916",
        "HDSLS10916MI",
        "HFACR200",
        "HDINV90196",
        "HISFC90196",
        "SALES_SUMMARY",
        "SALES_SUMMARY_MI",
        "SALES_SUMMARY_STATUS",
        "SALES_SUMMARY_STATUS_MI",
        "UsersList",
        "UsersByDivision",
        "DashboardMaster",
        "DashboardDivisionAccess",
    ):
        s = re.sub(
            rf"\bFROM\s+{tbl}\b",
            f'FROM noor."{tbl}"',
            s,
            flags=re.IGNORECASE,
        )
        s = re.sub(
            rf"\bJOIN\s+{tbl}\b",
            f'JOIN noor."{tbl}"',
            s,
            flags=re.IGNORECASE,
        )
        s = re.sub(
            rf"\bINTO\s+{tbl}\b",
            f'INTO noor."{tbl}"',
            s,
            flags=re.IGNORECASE,
        )
        s = re.sub(
            rf"\bUPDATE\s+{tbl}\b",
            f'UPDATE noor."{tbl}"',
            s,
            flags=re.IGNORECASE,
        )
        s = re.sub(
            rf"\bDELETE\s+FROM\s+{tbl}\b",
            f'DELETE FROM noor."{tbl}"',
            s,
            flags=re.IGNORECASE,
        )
    return s


def _row_upper(row: dict | None) -> dict | None:
    if row is None:
        return None
    return {str(k).upper(): v for k, v in row.items()}


class Cursor:
    def __init__(self, conn, dictionary: bool = False):
        self._dictionary = dictionary
        self._cur = conn._pg.cursor(
            cursor_factory=psycopg2.extras.RealDictCursor if dictionary else None
        )

    def execute(self, sql, params=None):
        try:
            self._cur.execute(_sql_mysql_to_pg(sql), params or None)
        except psycopg2.Error as e:
            raise Error(str(e)) from e

    def fetchone(self):
        row = self._cur.fetchone()
        if self._dictionary and row is not None:
            return _row_upper(dict(row))
        return row

    def fetchall(self):
        rows = self._cur.fetchall()
        if self._dictionary:
            return [_row_upper(dict(r)) for r in rows]
        return rows

    def close(self):
        self._cur.close()


class Connection:
    def __init__(self, pg_conn):
        self._pg = pg_conn
        self.autocommit = True

    def cursor(self, dictionary=False):
        return Cursor(self, dictionary=dictionary)

    def commit(self):
        self._pg.commit()

    def close(self):
        self._pg.close()


def connect(**kwargs):
    host = kwargs.get("host") or os.environ.get("NOOR_DB_HOST", "127.0.0.1")
    port = int(kwargs.get("port") or os.environ.get("NOOR_DB_PORT", "5437"))
    dbname = kwargs.get("dbname") or "noor_warehouse"
    legacy = kwargs.get("database")
    if legacy in ("cloudvirtualdb", "noor"):
        dbname = "noor_warehouse"
    user = kwargs.get("user") or os.environ.get("NOOR_DB_USER", "warehouse")
    password = kwargs.get("password") or os.environ.get("NOOR_DB_PASSWORD", "")
    try:
        pg = psycopg2.connect(
            host=host,
            port=port,
            dbname=dbname,
            user=user,
            password=password,
            options="-c search_path=noor,public",
        )
        pg.autocommit = True
        return Connection(pg)
    except psycopg2.Error as e:
        raise Error(str(e)) from e


# Install as mysql.connector
_mod = sys.modules.setdefault("mysql", type(sys)("mysql"))
_mod.connector = sys.modules[__name__]
sys.modules["mysql.connector"] = sys.modules[__name__]
