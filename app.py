
from flask import Flask, render_template, request, jsonify, session, send_file, redirect
from flask import request, jsonify
import pandas as pd
from werkzeug.utils import secure_filename
import os
import mysql.connector
from mysql.connector import Error
from datetime import date, datetime, timedelta
from io import BytesIO
import logging
import calendar
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from io import BytesIO
from datetime import datetime


def get_db_connection():

    return mysql.connector.connect(
        **DB_CONFIG
    )


UPLOAD_FOLDER = "uploads"

ALLOWED_EXTENSIONS = {
    "xlsx",
    "xls"
}

os.makedirs(
    UPLOAD_FOLDER,
    exist_ok=True
)


# ============================================================
# OPTIONAL EXCEL SUPPORT
# ============================================================

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None


# ============================================================
# APPLICATION
# ============================================================

app = Flask(__name__)

app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-me-in-production")

app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"


# ============================================================
# DATABASE
# ============================================================

DB_CONFIG = {
    "host": os.environ.get("NOOR_DB_HOST", "127.0.0.1"),
    "port": int(os.environ.get("NOOR_DB_PORT", "5437")),
    "database": os.environ.get("NOOR_PG_SCHEMA", "noor"),
    "dbname": "noor_warehouse",
    "user": os.environ.get("NOOR_DB_USER", "warehouse"),
    "password": os.environ.get("NOOR_DB_PASSWORD", ""),
    "autocommit": True,
}


# ============================================================
# PHYSICAL SALES TABLES
# ============================================================

PRIMARY_SALES_TABLE = "HDSLS10916"
MI_SALES_TABLE = "HDSLS10916MI"


# ============================================================
# NON-MI DIVISIONS
# ============================================================

NON_MI_DIVISIONS = {
    "TF",
    "BF",
    "SB",
    "TP",
    "FP",
}


# ============================================================
# GROUP FIELDS
# ============================================================

GROUP_FIELDS = {
    "category": "SELECTIONCODEDESCRIPTION",
    "customer": "CUSTOMERNAME",
    "country": "COUNTRYNAME",
    "salesrep": "SALESREPNAME",
    "item": "ITEMCODE",
    "producttype": "PRODUCTTYPE",
    "itemgroup": "ITEMGROUPDESCRIPTION",
}


# ============================================================
# CATEGORY FIELD RULE
#
# IMPORTANT:
#
# MI:
#     ITEMGROUPDESCRIPTION
#
# All other divisions:
#     SELECTIONCODEDESCRIPTION
#
# ALL:
#     MI records      -> ITEMGROUPDESCRIPTION
#     Non-MI records  -> SELECTIONCODEDESCRIPTION
# ============================================================

def get_category_field(division=None):
    """
    Return the physical database field that must be treated
    as Category for the requested division.

    MI:
        ITEMGROUPDESCRIPTION

    All other divisions:
        SELECTIONCODEDESCRIPTION
    """

    if division is not None:

        division_value = str(
            division
        ).strip().upper()

        if division_value == "MI":
            return "ITEMGROUPDESCRIPTION"

    return "SELECTIONCODEDESCRIPTION"


def get_category_expression(division=None):
    """
    Return the SQL expression to use as Category.

    For a specific division:

        MI  -> ITEMGROUPDESCRIPTION
        ALL -> division-aware CASE expression
        Other -> SELECTIONCODEDESCRIPTION

    The ALL expression is important because the sales source
    can contain MI and non-MI records together.
    """

    if division is not None:

        division_value = str(
            division
        ).strip().upper()

        if division_value == "MI":

            return """
                COALESCE(
                    NULLIF(
                        TRIM(`ITEMGROUPDESCRIPTION`),
                        ''
                    ),
                    'Not Specified'
                )
            """

        if division_value != "ALL":

            return """
                COALESCE(
                    NULLIF(
                        TRIM(`SELECTIONCODEDESCRIPTION`),
                        ''
                    ),
                    'Not Specified'
                )
            """

    # --------------------------------------------------------
    # ALL DIVISIONS
    #
    # MI records:
    #     ITEMGROUPDESCRIPTION
    #
    # Other records:
    #     SELECTIONCODEDESCRIPTION
    # --------------------------------------------------------

    return """
        COALESCE(
            NULLIF(
                TRIM(
                    CASE
                        WHEN UPPER(
                            TRIM(
                                COALESCE(`DIVISION`, '')
                            )
                        ) = 'MI'
                        THEN `ITEMGROUPDESCRIPTION`
                        ELSE `SELECTIONCODEDESCRIPTION`
                    END
                ),
                ''
            ),
            'Not Specified'
        )
    """


# ============================================================
# DATABASE CONNECTION
# ============================================================

def conn():
    """
    Create a new MySQL connection.
    """

    return mysql.connector.connect(
        **DB_CONFIG
    )


def rows(sql, params=()):
    """
    Execute SELECT and return dictionary rows.
    """

    c = conn()
    cur = c.cursor(dictionary=True)

    try:

        cur.execute(
            sql,
            params
        )

        return cur.fetchall()

    finally:

        cur.close()
        c.close()


# ============================================================
# SAFE SQL IDENTIFIER
# ============================================================

def quote_identifier(name):
    """
    Quote a MySQL identifier.

    Only use this for identifiers that originate from our
    controlled configuration/table metadata.
    """

    return "`" + str(name).replace("`", "``") + "`"


# ============================================================
# TABLE HELPERS
# ============================================================

def table_name(actual_name):
    """
    Find the actual table name in INFORMATION_SCHEMA.
    """

    c = conn()
    cur = c.cursor()

    try:

        cur.execute(
            """
            SELECT TABLE_NAME
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_SCHEMA = %s
              AND LOWER(TABLE_NAME) = LOWER(%s)
            LIMIT 1
            """,
            (
                DB_CONFIG["database"],
                actual_name,
            ),
        )

        r = cur.fetchone()

        if not r:

            raise Exception(
                f"Table '{actual_name}' was not found in database "
                f"'{DB_CONFIG['database']}'."
            )

        return r[0]

    finally:

        cur.close()
        c.close()


def table_columns(table):
    """
    Return columns for a table.
    """

    c = conn()
    cur = c.cursor()

    try:

        cur.execute(
            """
            SELECT COLUMN_NAME
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_SCHEMA = %s
              AND TABLE_NAME = %s
            ORDER BY ORDINAL_POSITION
            """,
            (
                DB_CONFIG["database"],
                table,
            ),
        )

        return [
            r[0]
            for r in cur.fetchall()
        ]

    finally:

        cur.close()
        c.close()


def find_column(columns, candidates):
    """
    Find a column using exact case-insensitive matching.
    """

    lookup = {
        str(x).upper(): x
        for x in columns
    }

    for candidate in candidates:

        if candidate.upper() in lookup:

            return lookup[
                candidate.upper()
            ]

    return None


def find_column_by_pattern(columns, patterns):
    """
    Find a column using partial matching.
    """

    for column in columns:

        name = (
            str(column)
            .upper()
            .replace("_", "")
        )

        for pattern in patterns:

            if pattern.upper() in name:

                return column

    return None


# ============================================================
# USER TABLE
# ============================================================

def user_table_columns():

    table = table_name(
        "UsersList"
    )

    columns = table_columns(
        table
    )

    username_col = find_column(
        columns,
        [
            "USERNAME",
            "USER_NAME",
            "LOGINNAME",
            "LOGIN_NAME",
            "LOGINID",
            "LOGIN_ID",
            "USERID",
            "USER_ID",
            "USERCODE",
            "USER_CODE",
            "USER",
        ],
    )

    password_col = find_column(
        columns,
        [
            "PASSWORD",
            "PASS",
            "PASSWD",
            "PWD",
            "USERPASSWORD",
            "USER_PASSWORD",
            "USERPWD",
            "USER_PWD",
            "LOGINPASSWORD",
            "LOGIN_PASSWORD",
            "LOGINPWD",
            "LOGIN_PWD",
        ],
    )

    if not username_col:

        username_col = find_column_by_pattern(
            columns,
            [
                "USERNAME",
                "LOGINNAME",
                "LOGINID",
                "USERCODE",
            ],
        )

    if not username_col:

        username_col = find_column_by_pattern(
            columns,
            [
                "USER"
            ],
        )

    if not password_col:

        password_col = find_column_by_pattern(
            columns,
            [
                "PASSWORD",
                "PASSWD",
                "USERPWD",
                "LOGINPWD",
                "PWD",
            ],
        )

    if not username_col or not password_col:

        raise Exception(
            "UsersList login columns could not be identified. "
            "Available columns are: "
            + ", ".join(
                str(x)
                for x in columns
            )
        )

    return (
        table,
        username_col,
        password_col,
    )


# ============================================================
# USER DIVISION TABLE
# ============================================================

def division_table_columns():

    table = table_name(
        "UsersByDivision"
    )

    columns = table_columns(
        table
    )

    username_col = find_column(
        columns,
        [
            "USERNAME",
            "USER_NAME",
            "LOGINNAME",
            "LOGIN_NAME",
            "LOGINID",
            "LOGIN_ID",
            "USERID",
            "USER_ID",
            "USERCODE",
            "USER_CODE",
            "USER",
        ],
    )

    division_col = find_column(
        columns,
        [
            "DIVISION",
            "DIVISIONCODE",
            "DIVISION_CODE",
            "DIVISIONNAME",
            "DIVISION_NAME",
        ],
    )

    if not username_col:

        username_col = find_column_by_pattern(
            columns,
            [
                "USERNAME",
                "LOGINNAME",
                "LOGINID",
                "USERCODE",
                "USER",
            ],
        )

    if not division_col:

        division_col = find_column_by_pattern(
            columns,
            [
                "DIVISION"
            ],
        )

    if not username_col or not division_col:

        raise Exception(
            "UsersByDivision columns could not be identified. "
            "Available columns are: "
            + ", ".join(
                str(x)
                for x in columns
            )
        )

    return (
        table,
        username_col,
        division_col,
    )


# ============================================================
# AUTHENTICATION
# ============================================================

def authenticate_user(
    username,
    password
):

    table, username_col, password_col = (
        user_table_columns()
    )

    sql = f"""
        SELECT *
        FROM {quote_identifier(table)}
        WHERE CAST({quote_identifier(username_col)} AS CHAR) = %s
          AND CAST({quote_identifier(password_col)} AS CHAR) = %s
        LIMIT 1
    """

    result = rows(
        sql,
        (
            username,
            password,
        ),
    )

    return (
        result[0]
        if result
        else None
    )


# ============================================================
# GET USER DIVISIONS
# ============================================================

def get_user_divisions(username):

    table, username_col, division_col = (
        division_table_columns()
    )

    sql = f"""
        SELECT DISTINCT
            TRIM(
                CAST({quote_identifier(division_col)} AS CHAR)
            ) AS division
        FROM {quote_identifier(table)}
        WHERE CAST({quote_identifier(username_col)} AS CHAR) = %s
          AND {quote_identifier(division_col)} IS NOT NULL
          AND TRIM(
                CAST({quote_identifier(division_col)} AS CHAR)
              ) <> ''
        ORDER BY
            TRIM(
                CAST({quote_identifier(division_col)} AS CHAR)
            )
    """

    result = rows(
        sql,
        (username,),
    )

    return [
        str(
            r["division"]
        ).strip()
        for r in result
        if r.get("division") is not None
    ]


# ============================================================
# CURRENT USER
# ============================================================

def current_user():

    username = session.get(
        "username"
    )

    if not username:

        return None

    return {
        "username": username,
        "divisions": session.get(
            "divisions",
            [],
        ),
    }


# ============================================================
# NORMALIZE DIVISIONS
# ============================================================

def normalize_divisions(divisions):

    result = []

    existing = set()

    for division in divisions or []:

        value = str(
            division
        ).strip()

        if not value:
            continue

        upper_value = value.upper()

        if upper_value not in existing:

            result.append(value)

            existing.add(
                upper_value
            )

    return result


# ============================================================
# DEFAULT DIVISION
# ============================================================

def get_default_division(
    divisions=None
):
    """
    Return the first division in the user's permission order.

    If the user has ALL permission, the first actual report
    division is returned.

    IMPORTANT:
    This is used as the default filter instead of ALL.
    """

    if divisions is None:

        user = current_user()

        if not user:
            return None

        divisions = user.get(
            "divisions",
            [],
        )

    divisions = normalize_divisions(
        divisions
    )

    if not divisions:
        return None

    real_divisions = [
        x
        for x in divisions
        if x.upper() != "ALL"
    ]

    if real_divisions:

        return real_divisions[0]

    return "ALL"


# ============================================================
# DEFAULT PERIOD
# ============================================================

def get_default_period():
    """
    Default reporting period is always the current month
    and current year.
    """

    today = date.today()

    return {
        "start_year": today.year,
        "start_month": today.month,
        "end_year": today.year,
        "end_month": today.month,
    }


# ============================================================
# DIVISION PERMISSION
# ============================================================

def division_allowed(
    requested_division
):

    user = current_user()

    if not user:
        return False

    allowed = [
        str(x).strip().upper()
        for x in user.get(
            "divisions",
            []
        )
    ]

    requested = str(
        requested_division or ""
    ).strip().upper()

    if not requested:
        return False

    if requested == "ALL":

        return "ALL" in allowed

    return requested in allowed


# ============================================================
# SALES SOURCE BUILDER
# ============================================================

def build_sales_source(
    allowed_divisions=None,
    selected_divisions=None,
):
    """
    Build a user-specific UNION source.

    Examples:

        User = TF

            HDSLS10916
            WHERE DIVISION = 'TF'

        User = MI

            HDSLS10916MI

        User = TF, MI

            HDSLS10916MI
            UNION ALL
            HDSLS10916 WHERE DIVISION = 'TF'

        User = TF, BF, MI

            HDSLS10916MI
            UNION ALL
            HDSLS10916 WHERE DIVISION IN ('TF','BF')
    """

    if allowed_divisions is None:

        user = current_user()

        if not user:

            raise Exception(
                "User session is not available."
            )

        allowed_divisions = user.get(
            "divisions",
            [],
        )

    allowed_divisions = normalize_divisions(
        allowed_divisions
    )

    selected_divisions = normalize_divisions(
        selected_divisions
    )

    allowed_upper = {
        str(x).strip().upper()
        for x in allowed_divisions
    }

    # --------------------------------------------------------
    # ALL ACCESS
    # --------------------------------------------------------

    if "ALL" in allowed_upper:

        source = f"""
        (
            SELECT *
            FROM {quote_identifier(MI_SALES_TABLE)}

            UNION ALL

            SELECT *
            FROM {quote_identifier(PRIMARY_SALES_TABLE)}
        ) AS sales_data
        """

        return source

    # --------------------------------------------------------
    # PERMITTED NON-MI DIVISIONS
    # --------------------------------------------------------

    permitted_non_mi = [
        division
        for division in allowed_divisions
        if division.upper() != "MI"
    ]

    permitted_non_mi = [
        division
        for division in permitted_non_mi
        if division.upper() in NON_MI_DIVISIONS
    ]

    # --------------------------------------------------------
    # SELECTED DIVISIONS
    # --------------------------------------------------------

    if selected_divisions:

        selected_upper = {
            x.upper()
            for x in selected_divisions
        }

        permitted_non_mi = [
            division
            for division in permitted_non_mi
            if division.upper()
            in selected_upper
        ]

        include_mi = (
            "MI" in selected_upper
            and
            "MI" in allowed_upper
        )

    else:

        include_mi = (
            "MI"
            in allowed_upper
        )

    source_parts = []

    # --------------------------------------------------------
    # MI SOURCE
    # --------------------------------------------------------

    if include_mi:

        source_parts.append(
            f"""
            SELECT *
            FROM {quote_identifier(MI_SALES_TABLE)}
            """
        )

    # --------------------------------------------------------
    # NON-MI SOURCE
    # --------------------------------------------------------

    if permitted_non_mi:

        safe_values = []

        for division in permitted_non_mi:

            safe_value = (
                str(division)
                .replace("\\", "\\\\")
                .replace("'", "''")
            )

            safe_values.append(
                f"'{safe_value}'"
            )

        division_list = ",".join(
            safe_values
        )

        source_parts.append(
            f"""
            SELECT *
            FROM {quote_identifier(PRIMARY_SALES_TABLE)}
            WHERE `DIVISION` IN ({division_list})
            """
        )

    # --------------------------------------------------------
    # NO SOURCE
    # --------------------------------------------------------

    if not source_parts:

        raise Exception(
            "No sales source is available for the user's "
            "division permissions."
        )

    # --------------------------------------------------------
    # SINGLE SOURCE
    # --------------------------------------------------------

    if len(source_parts) == 1:

        return f"""
        (
            {source_parts[0]}
        ) AS sales_data
        """

    # --------------------------------------------------------
    # UNION SOURCE
    # --------------------------------------------------------

    return f"""
    (
        {' UNION ALL '.join(source_parts)}
    ) AS sales_data
    """


# ============================================================
# GET REPORT DIVISIONS
# ============================================================

def get_report_divisions():

    user = current_user()

    if not user:
        return []

    allowed = normalize_divisions(
        user.get(
            "divisions",
            []
        )
    )

    allowed_upper = {
        x.upper()
        for x in allowed
    }

    # --------------------------------------------------------
    # ALL ACCESS
    # --------------------------------------------------------

    if "ALL" in allowed_upper:

        result = rows(
            f"""
            SELECT DISTINCT
                TRIM(`DIVISION`) AS division
            FROM {quote_identifier(PRIMARY_SALES_TABLE)}
            WHERE `DIVISION` IS NOT NULL
              AND TRIM(`DIVISION`) <> ''

            UNION

            SELECT DISTINCT
                TRIM(`DIVISION`) AS division
            FROM {quote_identifier(MI_SALES_TABLE)}
            WHERE `DIVISION` IS NOT NULL
              AND TRIM(`DIVISION`) <> ''

            ORDER BY division
            """
        )

        return [
            str(
                r["division"]
            ).strip()
            for r in result
            if r.get("division") is not None
        ]

    # --------------------------------------------------------
    # USER-SPECIFIC ACCESS
    # --------------------------------------------------------

    return allowed


# ============================================================
# GET DEFAULT REPORT DIVISION
# ============================================================

def get_default_report_division():

    divisions = get_report_divisions()

    if not divisions:
        return None

    return divisions[0]


# ============================================================
# MARGIN
# ============================================================

MARGIN = """
(
    COALESCE(`INVOICEDAMOUNT`, 0)
    -
    (
        COALESCE(`DELIVEREDQUANTITY`, 0)
        *
        COALESCE(`MATERIALCOST`, 0)
    )
)
"""


# ============================================================
# PERIOD HELPERS
# ============================================================

def period(
    sy,
    sm,
    ey,
    em
):

    sy = int(sy)
    sm = int(sm)
    ey = int(ey)
    em = int(em)

    if not 1 <= sm <= 12:

        raise ValueError(
            "Start month must be 1-12."
        )

    if not 1 <= em <= 12:

        raise ValueError(
            "End month must be 1-12."
        )

    if (sy, sm) > (ey, em):

        raise ValueError(
            "Start period cannot be after end period."
        )

    return (
        sy,
        sm,
        ey,
        em,
    )


def build_date_where(
    sy,
    sm,
    ey,
    em
):

    start_date = (
        f"{sy:04d}-{sm:02d}-01"
    )

    if em == 12:

        end_year = ey + 1
        end_month = 1

    else:

        end_year = ey
        end_month = em + 1

    end_date = (
        f"{end_year:04d}-{end_month:02d}-01"
    )

    where = """
        `INVOICEDATE` IS NOT NULL
        AND `INVOICEDATE` >= %s
        AND `INVOICEDATE` < %s
    """

    return (
        where,
        [
            start_date,
            end_date,
        ],
    )


def add_division_filter(
    where,
    params,
    division,
):

    if (
        division
        and
        division.upper() != "ALL"
    ):

        where += """
            AND `DIVISION` = %s
        """

        params.append(
            division
        )

    return (
        where,
        params
    )


# ============================================================
# DYNAMIC DASHBOARD ACCESS
# ============================================================
# Dashboard permissions are controlled by DashboardMaster and
# DashboardDivisionAccess. Users continue to use UsersByDivision.
# A dashboard is available when at least one of the user's divisions
# has access to it. Therefore the same division automatically shares
# the same dashboard permissions.
# ============================================================

DASHBOARD_MASTER_TABLE = "DashboardMaster"
DASHBOARD_ACCESS_TABLE = "DashboardDivisionAccess"

DASHBOARD_ROUTE_MAP = {
    "SALES": "/sales-dashboard",
    "SALES SUMMARY": "/sales-summary",
    "SALES_SUMMARY": "/sales-summary",
    "PRODUCTION": "/production-dashboard",
    "INVENTORY": "/inventory-dashboard",
    "FINANCE": "/Finance-dashboard",

}

DASHBOARD_DEFAULTS = {
    "SALES": {
        "name": "Sales Dashboard",
        "description": "Sales KPIs, sales summary and performance analysis",
        "icon": "sales",
        "sort_order": 10,
    },
    "SALES SUMMARY": {
        "name": "Sales Summary",
        "description": "MTD / YTD Sales Summary",
        "icon": "sales-summary",
        "sort_order": 15,
    },
    "PRODUCTION": {
        "name": "Finance Ageing Dashboard",
        "description": "Finance ageing, receivables and customer information",
        "icon": "production",
        "sort_order": 20,
    },
    "INVENTORY": {
        "name": "Inventory Dashboard",
        "description": "Inventory, stock and material information",
        "icon": "inventory",
        "sort_order": 30,
    },
    "FINANCE": {
        "name": "Finance Dashboard",
        "description": "Finance, Ageing and Customer information",
        "icon": "Finance",
        "sort_order": 40,
    },

}


def allowed_file(filename):

    return (
        "." in filename
        and
        filename.rsplit(
            ".",
            1
        )[1].lower()
        in ALLOWED_EXTENSIONS
    )


def _dashboard_tables_exist():
    try:
        table_name(DASHBOARD_MASTER_TABLE)
        table_name(DASHBOARD_ACCESS_TABLE)
        return True
    except Exception:
        return False


def _dashboard_master_columns():
    table = table_name(DASHBOARD_MASTER_TABLE)
    columns = table_columns(table)
    code_col = find_column(columns, ["DASHBOARD_CODE", "CODE", "DASHBOARDCODE"])
    name_col = find_column(columns, ["DASHBOARD_NAME", "NAME", "DASHBOARDNAME"])
    route_col = find_column(columns, ["ROUTE", "URL", "PATH"])
    icon_col = find_column(columns, ["ICON"])
    desc_col = find_column(columns, ["DESCRIPTION", "DESC"])
    sort_col = find_column(columns, ["SORT_ORDER", "SORTORDER", "DISPLAY_ORDER", "DISPLAYORDER"])
    enabled_col = find_column(columns, ["ENABLED", "ACTIVE", "IS_ACTIVE", "ISACTIVE"])
    if not code_col or not name_col or not route_col:
        raise Exception("DashboardMaster must contain DASHBOARD_CODE, DASHBOARD_NAME and ROUTE.")
    return table, code_col, name_col, route_col, icon_col, desc_col, sort_col, enabled_col


def _dashboard_access_columns():
    table = table_name(DASHBOARD_ACCESS_TABLE)
    columns = table_columns(table)
    code_col = find_column(columns, ["DASHBOARD_CODE", "CODE", "DASHBOARDCODE"])
    division_col = find_column(columns, ["DIVISION", "DIVISION_CODE", "DIVISIONCODE", "DIVISION_NAME", "DIVISIONNAME"])
    enabled_col = find_column(columns, ["ENABLED", "ACTIVE", "IS_ACTIVE", "ISACTIVE"])
    if not code_col or not division_col:
        raise Exception("DashboardDivisionAccess must contain DASHBOARD_CODE and DIVISION.")
    return table, code_col, division_col, enabled_col


def get_user_dashboards(user=None):
    user = user or current_user()
    if not user:
        return []

    allowed_divisions = normalize_divisions(user.get("divisions", []))
    allowed_upper = {x.upper() for x in allowed_divisions}

    # Compatibility fallback until the new tables are created.
    if not _dashboard_tables_exist():
        if allowed_divisions:
            d = DASHBOARD_DEFAULTS["SALES"].copy()
            d.update({"dashboard_code": "SALES", "route": "/sales-dashboard", "enabled": True})
            return [d]
        return []

    mt, mc, mn, mr, mi, md, ms, me = _dashboard_master_columns()
    at, ac, ad, ae = _dashboard_access_columns()
    master_enabled = f"AND COALESCE(m.{quote_identifier(me)}, 1) = 1" if me else ""
    access_enabled = f"AND COALESCE(a.{quote_identifier(ae)}, 1) = 1" if ae else ""

    select_sql = f"""
        SELECT DISTINCT
            m.{quote_identifier(mc)} AS dashboard_code,
            m.{quote_identifier(mn)} AS dashboard_name,
            m.{quote_identifier(mr)} AS route,
            {('m.' + quote_identifier(mi)) if mi else 'NULL'} AS icon,
            {('m.' + quote_identifier(md)) if md else 'NULL'} AS description,
            {('m.' + quote_identifier(ms)) if ms else '0'} AS sort_order
        FROM {quote_identifier(mt)} m
        INNER JOIN {quote_identifier(at)} a
          ON UPPER(TRIM(CAST(a.{quote_identifier(ac)} AS CHAR))) =
             UPPER(TRIM(CAST(m.{quote_identifier(mc)} AS CHAR)))
        WHERE 1=1 {master_enabled} {access_enabled}
    """

    if "ALL" in allowed_upper:
        result = rows(select_sql + " ORDER BY sort_order, dashboard_name")
    else:
        if not allowed_divisions:
            return []
        placeholders = ",".join(["%s"] * len(allowed_divisions))
        sql = select_sql + f"""
          AND UPPER(TRIM(CAST(a.{quote_identifier(ad)} AS CHAR))) IN ({placeholders})
          ORDER BY sort_order, dashboard_name
        """
        result = rows(sql, tuple(allowed_divisions))

    output = []
    for r in result:
        code = str(r.get("dashboard_code") or "").strip().upper()
        # Accept either database naming convention for Sales Summary.
        if code == "SALES_SUMMARY":
            code = "SALES SUMMARY"
        route = DASHBOARD_ROUTE_MAP.get(code)
        if not code or not route:
            continue
        default = DASHBOARD_DEFAULTS.get(code, {})
        output.append({
            "dashboard_code": code,
            "dashboard_name": str(r.get("dashboard_name") or default.get("name") or code),
            "route": route,
            "icon": str(r.get("icon") or default.get("icon") or "dashboard"),
            "description": str(r.get("description") or default.get("description") or ""),
            "sort_order": int(r.get("sort_order") or default.get("sort_order") or 999),
            "enabled": True,
        })
    return output


def dashboard_allowed(code):
    code = str(code or "").strip().upper()
    return any(d["dashboard_code"] == code for d in get_user_dashboards())


def get_landing_url(user=None):
    dashboards = get_user_dashboards(user)
    if len(dashboards) == 1:
        return dashboards[0]["route"]
    return "/" if dashboards else None


# ============================================================
# HOME / MAIN MENU
# ============================================================

@app.route("/")
def home():
    """
    Application entry point.

    Unauthenticated users see index.html, which contains the login form.
    After successful login, the application always comes back here first.
    Authenticated users are then shown the dashboard selection/placeholder
    page. That page uses /api/dashboards to display only the dashboards
    permitted for the user's assigned divisions.
    """
    user = current_user()

    if not user:
        today = date.today()
        return render_template(
            "index.html",
            year=today.year,
            month=today.month,
            current_year=today.year,
            current_month=today.month,
        )

    dashboards = get_user_dashboards(user)

    if not dashboards:
        session.clear()
        today = date.today()
        return render_template(
            "index.html",
            year=today.year,
            month=today.month,
            current_year=today.year,
            current_month=today.month,
        )

    return render_template(
        "dashboard_placeholder.html",
        dashboard_name="Dashboard",
        username=user["username"],
        divisions=user["divisions"],
        dashboards=dashboards,
    )


@app.get("/api/dashboards")
def api_dashboards():
    user = current_user()
    if not user:
        return jsonify({"authenticated": False, "error": "Session expired. Please login again."}), 401
    dashboards = get_user_dashboards(user)
    return jsonify({"authenticated": True, "username": user["username"],
                    "divisions": user["divisions"], "dashboards": dashboards,
                    "count": len(dashboards), "landing_url": get_landing_url(user)})


@app.route("/main-menu")
def main_menu():
    """Return to the authenticated dashboard selection page."""
    user = current_user()
    if not user:
        return redirect("/")

    dashboards = get_user_dashboards(user)
    if not dashboards:
        session.clear()
        return redirect("/")

    return render_template(
        "dashboard_placeholder.html",
        dashboard_name="Dashboard",
        username=user["username"],
        divisions=user["divisions"],
        dashboards=dashboards,
    )


# ============================================================
# SALES DASHBOARD PAGE
# ============================================================

@app.route("/sales-dashboard")
def sales_dashboard_page():
    """Full Sales KPI dashboard."""
    if not current_user():
        return redirect("/")

    if not dashboard_allowed("SALES"):
        return jsonify({
            "error": "You do not have permission to access the Sales Dashboard."
        }), 403

    today = date.today()

    return render_template(
        "index.html",
        year=today.year,
        month=today.month,
        current_year=today.year,
        current_month=today.month,
        sales_dashboard=True,
    )


@app.route("/sales-summary")
def category_sales_report_page():
    """Sales Summary / MTD-YTD report."""
    if not current_user():
        return redirect("/")

    # Sales Summary can be enabled either as its own module or under SALES.
    if not (
        dashboard_allowed("SALES SUMMARY")
        or dashboard_allowed("SALES_SUMMARY")
        or dashboard_allowed("SALES")
    ):
        return jsonify({
            "error": "You do not have permission to access the Sales Summary."
        }), 403

    today = date.today()

    return render_template(
        "sales_summary.html",
        year=today.year,
        month=today.month,
        current_year=today.year,
        current_month=today.month,
    )


# ============================================================
# PRODUCTION DASHBOARD
# ============================================================

PRODUCTION_TABLE = "HISFC90196"

def production_allowed_divisions():
    user = current_user()
    if not user:
        return []
    return normalize_divisions(user.get("divisions", []))

def production_source_and_filter(selected_division=None):
    table = table_name(PRODUCTION_TABLE)
    allowed = production_allowed_divisions()
    allowed_upper = {str(x).strip().upper() for x in allowed}
    selected = str(selected_division or "").strip()
    selected_upper = selected.upper()
    if not allowed:
        raise Exception("No division permission is assigned to this user.")
    where = ["1=1"]
    params = []
    if "ALL" not in allowed_upper:
        if selected and selected_upper != "ALL":
            if selected_upper not in allowed_upper:
                raise Exception("You do not have permission for the selected division.")
            where.append("UPPER(TRIM(CAST(`DIVISION` AS CHAR))) = %s")
            params.append(selected_upper)
        else:
            placeholders = ",".join(["%s"] * len(allowed_upper))
            where.append(f"UPPER(TRIM(CAST(`DIVISION` AS CHAR))) IN ({placeholders})")
            params.extend(sorted(allowed_upper))
    elif selected and selected_upper != "ALL":
        where.append("UPPER(TRIM(CAST(`DIVISION` AS CHAR))) = %s")
        params.append(selected_upper)
    return table, " AND ".join(where), params

def production_period_where(sy, sm, ey, em):
    sy, sm, ey, em = period(sy, sm, ey, em)
    return ("((`YEAR` > %s OR (`YEAR` = %s AND `MONTH` >= %s)) AND "
            "(`YEAR` < %s OR (`YEAR` = %s AND `MONTH` <= %s)))",
            [sy, sy, sm, ey, ey, em])

def production_rows(sql, params=()):
    return rows(sql, tuple(params))

@app.route("/production-dashboard")
def production_dashboard_page():
    """
    Production Dashboard route now serves the Finance Ageing dashboard.
    The existing dashboard permission code remains unchanged:
        Dashboard Code = PRODUCTION
        Route         = /production-dashboard
    """
    if not current_user():
        return redirect("/")

    if not dashboard_allowed("PRODUCTION"):
        return jsonify({
            "error": "You do not have permission to access the Production Dashboard."
        }), 403

    return render_template("production_dashboard.html")

# ============================================================
# FINANCE AGEING / CUSTOMER INFORMATION
# Served through the existing /Finance-dashboard menu item
#
# SECURITY RULE:
# UsersByDivision is the ONLY authority for Finance division access.
#
# The browser/HTML is NEVER trusted for division security.
# Every Finance API request obtains the logged-in user's permitted
# divisions and applies them directly to HFACR200.
# ============================================================

FINANCE_AGEING_TABLE = "HFACR200"


# ============================================================
# GET FRESH FINANCE DIVISION PERMISSIONS
# ============================================================

def finance_allowed_divisions():
    """
    Get the divisions assigned to the currently logged-in user.

    IMPORTANT:
    Do NOT rely only on session["divisions"] here.

    UsersByDivision is the master source for division permissions.
    This function reads it again for Finance requests so that the
    Finance dashboard cannot accidentally expose another division.
    """

    user = current_user()

    if not user:
        return []

    username = str(
        user.get("username", "")
    ).strip()

    if not username:
        return []

    try:
        divisions = get_user_divisions(username)

        return normalize_divisions(
            divisions
        )

    except Exception:
        logging.exception(
            "Unable to retrieve Finance division permissions "
            "from UsersByDivision for user %s",
            username
        )
        return []


# ============================================================
# FINANCE SECURITY WHERE CLAUSE
# ============================================================

def finance_where(
    selected_division=None,
    selected_customer=None,
    selected_salesman=None
):
    """
    Build the secure Finance Ageing WHERE clause.

    UsersByDivision controls the permitted divisions.

    Examples:

        User assigned:
            MI

        User sees:
            MI only

        User assigned:
            MI, FP

        User sees:
            MI and FP

        User assigned:
            ALL

        User sees:
            all Finance divisions

    If the browser requests a division that is not assigned to
    the user, an exception is raised and the request is rejected.
    """

    allowed = finance_allowed_divisions()

    if not allowed:
        raise PermissionError(
            "No division permission is assigned to this user."
        )

    # --------------------------------------------------------
    # Normalize permitted divisions
    # --------------------------------------------------------

    allowed_upper = {
        str(x).strip().upper()
        for x in allowed
        if str(x).strip()
    }

    if not allowed_upper:
        raise PermissionError(
            "No valid division permission is assigned to this user."
        )

    has_all_permission = (
        "ALL" in allowed_upper
    )

    # --------------------------------------------------------
    # Base WHERE
    # --------------------------------------------------------

    where = [
        "1=1"
    ]

    params = []

    # --------------------------------------------------------
    # DIVISION SECURITY
    #
    # This is the most important part.
    #
    # If user does NOT have ALL:
    #
    #     HFACR200.DIVISION must be one of
    #     UsersByDivision divisions.
    # --------------------------------------------------------

    if not has_all_permission:

        placeholders = ",".join(
            ["%s"] * len(allowed_upper)
        )

        where.append(
            "UPPER(TRIM(CAST(`DIVISION` AS CHAR))) "
            f"IN ({placeholders})"
        )

        params.extend(
            sorted(allowed_upper)
        )

    # --------------------------------------------------------
    # USER SELECTED DIVISION
    # --------------------------------------------------------

    division = str(
        selected_division or ""
    ).strip()

    if division:

        division_upper = division.upper()

        # "ALL" means all divisions assigned to this user,
        # NOT all divisions in HFACR200.
        if division_upper == "ALL":

            if not has_all_permission:

                # User is NOT permitted to select ALL.
                # Do not allow the browser to bypass security.
                raise PermissionError(
                    "You do not have permission to select ALL divisions."
                )

            # User really has ALL permission.
            # No additional division condition is necessary.

        else:

            # ------------------------------------------------
            # SECURITY CHECK
            # ------------------------------------------------

            if (
                not has_all_permission
                and division_upper not in allowed_upper
            ):
                raise PermissionError(
                    "You do not have permission for the selected "
                    f"division: {division}"
                )

            # ------------------------------------------------
            # Apply selected division
            # ------------------------------------------------

            where.append(
                "UPPER(TRIM(CAST(`DIVISION` AS CHAR))) = %s"
            )

            params.append(
                division_upper
            )

    # --------------------------------------------------------
    # CUSTOMER FILTER
    # --------------------------------------------------------

    customer = str(
        selected_customer or ""
    ).strip()

    if customer:

        where.append(
            "UPPER(TRIM(CAST(`CUSTOMERCODE` AS CHAR))) = %s"
        )

        params.append(
            customer.upper()
        )

    # --------------------------------------------------------
    # SALESMAN FILTER
    # --------------------------------------------------------

    salesman = str(
        selected_salesman or ""
    ).strip()

    if salesman:

        where.append(
            "UPPER(TRIM(CAST(`SALESMANNAME` AS CHAR))) = %s"
        )

        params.append(
            salesman.upper()
        )

    return (
        " AND ".join(where),
        params
    )


# ============================================================
# FINANCE AGEING DATA API
# ============================================================

@app.get("/api/finance-ageing/data")
def api_finance_ageing_data():

    user = current_user()

    # --------------------------------------------------------
    # LOGIN CHECK
    # --------------------------------------------------------

    if not user:
        return jsonify({
            "success": False,
            "error": "Session expired. Please login again."
        }), 401

    # --------------------------------------------------------
    # DASHBOARD ACCESS CHECK
    # --------------------------------------------------------

    if not dashboard_allowed("FINANCE"):
        return jsonify({
            "success": False,
            "error":
                "You do not have permission to access "
                "the Finance Ageing Dashboard."
        }), 403

    try:

        # ----------------------------------------------------
        # TABLE
        # ----------------------------------------------------

        table = quote_identifier(
            table_name(
                FINANCE_AGEING_TABLE
            )
        )

        # ----------------------------------------------------
        # BROWSER FILTERS
        # ----------------------------------------------------

        division = request.args.get(
            "division",
            ""
        ).strip()

        customer = request.args.get(
            "customer",
            ""
        ).strip()

        salesman = request.args.get(
            "salesman",
            ""
        ).strip()

        # ----------------------------------------------------
        # SECURE WHERE
        #
        # This applies UsersByDivision security FIRST.
        # ----------------------------------------------------

        where, params = finance_where(
            selected_division=division,
            selected_customer=customer,
            selected_salesman=salesman
        )

        # ----------------------------------------------------
        # QUERY
        # ----------------------------------------------------

        sql = f"""
            SELECT
                `CUSTOMERCODE`,
                `CUSTOMERNAME`,
                `SALESMANNAME`,
                `PAYMENTTERMS`,

                COALESCE(
                    `CREDITLIMIT`, 0
                ) AS `CREDITLIMIT`,

                COALESCE(
                    `ACTUALOVERDUE`, 0
                ) AS `ACTUALOVERDUE`,

                COALESCE(
                    `DAYS30`, 0
                ) AS `DAYS30`,

                COALESCE(
                    `DAYS60`, 0
                ) AS `DAYS60`,

                COALESCE(
                    `DAYS90`, 0
                ) AS `DAYS90`,

                COALESCE(
                    `DAYS120`, 0
                ) AS `DAYS120`,

                COALESCE(
                    `DAYS150`, 0
                ) AS `DAYS150`,

                COALESCE(
                    `DAYS180`, 0
                ) AS `DAYS180`,

                COALESCE(
                    `DAYS270`, 0
                ) AS `DAYS270`,

                COALESCE(
                    `DAYS365`, 0
                ) AS `DAYS365`,

                COALESCE(
                    `DAYSABOVE365`, 0
                ) AS `DAYSABOVE365`,

                COALESCE(
                    `PDCBALANCE`, 0
                ) AS `PDCBALANCE`,

                COALESCE(
                    `OUTSTANDING`, 0
                ) AS `OUTSTANDING`,

                `DIVISION`,

                COALESCE(
                    `BALANCE`, 0
                ) AS `BALANCE`,

                COALESCE(
                    `TERMSDAYS`, 0
                ) AS `TERMSDAYS`

            FROM {table}

            WHERE {where}

            ORDER BY
                COALESCE(
                    `OUTSTANDING`, 0
                ) DESC
        """

        result = rows(
            sql,
            tuple(params)
        )

        # ----------------------------------------------------
        # GET ACTUAL PERMITTED DIVISIONS
        # ----------------------------------------------------

        allowed = finance_allowed_divisions()

        # ----------------------------------------------------
        # RETURN DATA
        # ----------------------------------------------------

        return jsonify({

            "success": True,

            "username":
                user["username"],

            "divisions":
                allowed,

            "count":
                len(result),

            "data":
                result
        })

    except PermissionError as e:

        logging.warning(
            "Finance permission denied for %s: %s",
            user.get("username"),
            str(e)
        )

        return jsonify({
            "success": False,
            "error": str(e)
        }), 403

    except Exception as e:

        logging.exception(
            "Finance ageing data error"
        )

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# ============================================================
# FINANCE FILTER API
#
# IMPORTANT:
# Division, Customer and Salesman are all generated from the
# SECURED HFACR200 dataset.
# ============================================================

@app.get("/api/finance-ageing/filters")
def api_finance_ageing_filters():

    user = current_user()

    # --------------------------------------------------------
    # LOGIN CHECK
    # --------------------------------------------------------

    if not user:

        return jsonify({
            "success": False,
            "error":
                "Session expired. Please login again."
        }), 401

    # --------------------------------------------------------
    # DASHBOARD ACCESS CHECK
    # --------------------------------------------------------

    if not dashboard_allowed("FINANCE"):

        return jsonify({
            "success": False,
            "error":
                "You do not have permission to access "
                "the Finance Ageing Dashboard."
        }), 403

    try:

        # ----------------------------------------------------
        # GET ACTUAL USERSBYDIVISION PERMISSION
        # ----------------------------------------------------

        allowed = finance_allowed_divisions()

        if not allowed:

            return jsonify({
                "success": False,
                "error":
                    "No Finance division permission is assigned "
                    "to this user."
            }), 403

        # ----------------------------------------------------
        # SECURED WHERE
        #
        # NO browser parameter is used here.
        #
        # Therefore this always starts with the user's
        # UsersByDivision permissions.
        # ----------------------------------------------------

        table = quote_identifier(
            table_name(
                FINANCE_AGEING_TABLE
            )
        )

        where, params = finance_where()

        # ----------------------------------------------------
        # DIVISIONS
        # ----------------------------------------------------

        divisions = rows(
            f"""
                SELECT DISTINCT
                    TRIM(
                        CAST(`DIVISION` AS CHAR)
                    ) AS division

                FROM {table}

                WHERE {where}

                  AND `DIVISION` IS NOT NULL

                  AND TRIM(
                        CAST(`DIVISION` AS CHAR)
                      ) <> ''

                ORDER BY division
            """,
            tuple(params)
        )

        # ----------------------------------------------------
        # CUSTOMERS
        #
        # Because "where" already contains the division
        # permission, customers are automatically restricted
        # to the user's permitted divisions.
        # ----------------------------------------------------

        customers = rows(
            f"""
                SELECT DISTINCT

                    TRIM(
                        CAST(`CUSTOMERCODE` AS CHAR)
                    ) AS customercode,

                    TRIM(
                        CAST(`CUSTOMERNAME` AS CHAR)
                    ) AS customername

                FROM {table}

                WHERE {where}

                  AND `CUSTOMERCODE` IS NOT NULL

                  AND TRIM(
                        CAST(`CUSTOMERCODE` AS CHAR)
                      ) <> ''

                ORDER BY
                    customername,
                    customercode
            """,
            tuple(params)
        )

        # ----------------------------------------------------
        # SALESMEN
        #
        # Again restricted by the same division WHERE clause.
        # ----------------------------------------------------

        salesmen = rows(
            f"""
                SELECT DISTINCT

                    TRIM(
                        CAST(`SALESMANNAME` AS CHAR)
                    ) AS salesman

                FROM {table}

                WHERE {where}

                  AND `SALESMANNAME` IS NOT NULL

                  AND TRIM(
                        CAST(`SALESMANNAME` AS CHAR)
                      ) <> ''

                ORDER BY salesman
            """,
            tuple(params)
        )

        # ----------------------------------------------------
        # CLEAN DIVISION LIST
        # ----------------------------------------------------

        finance_divisions = []

        for row in divisions:

            value = str(
                row.get("division", "")
            ).strip()

            if not value:
                continue

            if value.upper() == "ALL":
                continue

            if value.upper() not in {
                str(x).strip().upper()
                for x in allowed
            } and "ALL" not in {
                str(x).strip().upper()
                for x in allowed
            }:
                continue

            finance_divisions.append(
                value
            )

        # ----------------------------------------------------
        # SORT / REMOVE DUPLICATES
        # ----------------------------------------------------

        finance_divisions = normalize_divisions(
            finance_divisions
        )

        # ----------------------------------------------------
        # ONLY SHOW "ALL" IF USER HAS EXPLICIT ALL PERMISSION
        # ----------------------------------------------------

        allowed_upper = {
            str(x).strip().upper()
            for x in allowed
        }

        if "ALL" in allowed_upper:

            division_options = [
                "ALL"
            ] + [
                x for x in finance_divisions
                if x.upper() != "ALL"
            ]

        else:

            division_options = finance_divisions

        # ----------------------------------------------------
        # DEFAULT DIVISION
        #
        # Do NOT force "ALL" for normal users.
        # ----------------------------------------------------

        default_division = get_default_division(
            allowed
        )

        if (
            default_division == "ALL"
            and "ALL" not in allowed_upper
            and division_options
        ):
            default_division = division_options[0]

        # ----------------------------------------------------
        # RETURN
        # ----------------------------------------------------

        return jsonify({

            "success": True,

            "username":
                user["username"],

            # This is the actual UsersByDivision permission.
            "allowed_divisions":
                allowed,

            # This is what HTML is allowed to display.
            "divisions":
                division_options,

            "default_division":
                default_division,

            "customers":
                customers,

            "salesmen": [
                str(
                    r["salesman"]
                ).strip()

                for r in salesmen

                if r.get("salesman") is not None
            ]
        })

    except Exception as e:

        logging.exception(
            "Finance ageing filter error"
        )

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# ============================================================
# FINANCE DASHBOARD PAGE ROUTES
#
# DashboardMaster route = /Finance-dashboard
# Keep aliases for compatibility.
# ============================================================

@app.route("/Finance-dashboard")
@app.route("/finance-dashboard")
@app.route("/finance-ageing")
def finance_ageing():

    # --------------------------------------------------------
    # LOGIN
    # --------------------------------------------------------

    if not current_user():
        return redirect("/")

    # --------------------------------------------------------
    # DASHBOARD ACCESS
    # --------------------------------------------------------

    if not dashboard_allowed("FINANCE"):

        return jsonify({
            "success": False,
            "error":
                "You do not have permission to access "
                "the Finance Dashboard."
        }), 403

    # --------------------------------------------------------
    # SERVE HTML
    #
    # HTML does NOT determine permissions.
    # The APIs above do.
    # --------------------------------------------------------

    return render_template(
        "finance_ageing.html"
    )


@app.get("/api/production-divisions")
def api_production_divisions():
    if not current_user():
        return jsonify({"error": "Session expired. Please login again."}), 401
    if not dashboard_allowed("PRODUCTION"):
        return jsonify({"error": "You do not have permission to access the Production Dashboard."}), 403
    try:
        table, where, params = production_source_and_filter()
        data = production_rows(f"SELECT DISTINCT TRIM(CAST(`DIVISION` AS CHAR)) AS division FROM {quote_identifier(table)} WHERE {where} AND `DIVISION` IS NOT NULL AND TRIM(CAST(`DIVISION` AS CHAR)) <> '' ORDER BY division", params)
        divisions = [str(x['division']).strip() for x in data if x.get('division') is not None]
        allowed = production_allowed_divisions()
        if any(str(x).upper() == 'ALL' for x in allowed):
            return jsonify({"divisions": ["ALL"] + divisions})
        return jsonify({"divisions": divisions})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.get("/api/production-dashboard")
def api_production_dashboard():
    if not current_user():
        return jsonify({"error": "Session expired. Please login again."}), 401
    if not dashboard_allowed("PRODUCTION"):
        return jsonify({"error": "You do not have permission to access the Production Dashboard."}), 403
    try:
        sy = request.args.get("start_year", type=int) or date.today().year
        sm = request.args.get("start_month", type=int) or date.today().month
        ey = request.args.get("end_year", type=int) or sy
        em = request.args.get("end_month", type=int) or sm
        division = request.args.get("division", "")
        top_n = max(1, min(request.args.get("top_n", default=10, type=int) or 10, 100))
        table, div_where, div_params = production_source_and_filter(division)
        period_where, period_params = production_period_where(sy, sm, ey, em)
        where = f"{div_where} AND {period_where}"
        params = div_params + period_params
        qtable = quote_identifier(table)
        receipt = "UPPER(TRIM(COALESCE(`TRANSACTIONTYPE`, ''))) = 'RECEIPT'"
        issue = "UPPER(TRIM(COALESCE(`TRANSACTIONTYPE`, ''))) = 'ISSUE'"
        kpi_sql = f"""SELECT
          COALESCE(SUM(CASE WHEN {receipt} THEN COALESCE(`QUANTITYINUNITS`,0) ELSE 0 END),0) produced_units,
          COALESCE(SUM(CASE WHEN {receipt} THEN COALESCE(`QUANTITYINKGS`,0) ELSE 0 END),0) produced_kgs,
          COALESCE(SUM(CASE WHEN {receipt} THEN ABS(COALESCE(`MATERIALCOST`,0)) ELSE 0 END),0) production_cost,
          COALESCE(SUM(CASE WHEN {issue} THEN ABS(COALESCE(`QUANTITYINUNITS`,0)) ELSE 0 END),0) consumed_units,
          COALESCE(SUM(CASE WHEN {issue} THEN ABS(COALESCE(`QUANTITYINKGS`,0)) ELSE 0 END),0) consumed_kgs,
          COALESCE(SUM(CASE WHEN {issue} THEN ABS(COALESCE(`MATERIALCOST`,0)) ELSE 0 END),0) consumption_cost,
          COUNT(DISTINCT CASE WHEN {receipt} THEN NULLIF(TRIM(CAST(`ORIGINALITEM` AS CHAR)), '') END) produced_items
          FROM {qtable} WHERE {where}"""
        kpi = production_rows(kpi_sql, params)[0]
        produced_kgs=float(kpi.get('produced_kgs') or 0); consumed_kgs=float(kpi.get('consumed_kgs') or 0)
        recovery_pct=(produced_kgs/consumed_kgs*100) if consumed_kgs else 0
        def grouped(name_expr, condition, order_expr, limit=True):
            lim = f" LIMIT {top_n}" if limit else ""
            sql=f"""SELECT {name_expr} AS name,
              COALESCE(SUM(ABS(COALESCE(`QUANTITYINUNITS`,0))),0) AS units,
              COALESCE(SUM(ABS(COALESCE(`QUANTITYINKGS`,0))),0) AS kgs,
              COALESCE(SUM(ABS(COALESCE(`MATERIALCOST`,0))),0) AS cost
              FROM {qtable} WHERE {where} AND {condition}
              GROUP BY name ORDER BY {order_expr} DESC{lim}"""
            return production_rows(sql, params)
        top_produced=grouped("COALESCE(NULLIF(TRIM(CAST(`DESCRIPTION` AS CHAR)),''), NULLIF(TRIM(CAST(`ORIGINALITEM` AS CHAR)),''),'Not Specified')", receipt, "kgs")
        top_consumed=grouped("COALESCE(NULLIF(TRIM(CAST(`DESCRIPTION` AS CHAR)),''), NULLIF(TRIM(CAST(`ORIGINALITEM` AS CHAR)),''),'Not Specified')", issue, "cost")
        by_category=grouped("COALESCE(NULLIF(TRIM(CAST(`PRODUCTCATEGORY` AS CHAR)),''),'Not Specified')", receipt, "kgs")
        by_division=grouped("COALESCE(NULLIF(TRIM(CAST(`DIVISION` AS CHAR)),''),'Not Specified')", receipt, "kgs", False)
        monthly_sql=f"""SELECT `YEAR` AS year, `MONTH` AS month,
          COALESCE(SUM(CASE WHEN {receipt} THEN ABS(COALESCE(`QUANTITYINUNITS`,0)) ELSE 0 END),0) AS produced_units,
          COALESCE(SUM(CASE WHEN {receipt} THEN ABS(COALESCE(`QUANTITYINKGS`,0)) ELSE 0 END),0) AS produced_kgs,
          COALESCE(SUM(CASE WHEN {issue} THEN ABS(COALESCE(`QUANTITYINKGS`,0)) ELSE 0 END),0) AS consumed_kgs
          FROM {qtable} WHERE {where} GROUP BY `YEAR`,`MONTH` ORDER BY `YEAR`,`MONTH`"""
        monthly=production_rows(monthly_sql, params)
        balance_sql=f"""SELECT COALESCE(NULLIF(TRIM(CAST(`ORIGINALITEM` AS CHAR)),''),'Not Specified') AS item,
          MAX(COALESCE(NULLIF(TRIM(CAST(`DESCRIPTION` AS CHAR)),''),'')) AS description,
          COALESCE(SUM(CASE WHEN {receipt} THEN ABS(COALESCE(`QUANTITYINKGS`,0)) ELSE 0 END),0) AS receipt_kgs,
          COALESCE(SUM(CASE WHEN {issue} THEN ABS(COALESCE(`QUANTITYINKGS`,0)) ELSE 0 END),0) AS issue_kgs,
          COALESCE(SUM(CASE WHEN {receipt} THEN ABS(COALESCE(`QUANTITYINUNITS`,0)) ELSE 0 END),0) AS receipt_units,
          COALESCE(SUM(CASE WHEN {issue} THEN ABS(COALESCE(`QUANTITYINUNITS`,0)) ELSE 0 END),0) AS issue_units
          FROM {qtable} WHERE {where} GROUP BY item HAVING receipt_kgs > 0 OR issue_kgs > 0
          ORDER BY receipt_kgs DESC, issue_kgs DESC LIMIT {top_n}"""
        balance=production_rows(balance_sql, params)
        for r in balance:
            issue_k=float(r.get('issue_kgs') or 0); receipt_k=float(r.get('receipt_kgs') or 0)
            r['recovery_pct']=(receipt_k/issue_k*100) if issue_k else None
            r['net_kgs']=receipt_k-issue_k
        return jsonify({"kpi": {"produced_units": kpi['produced_units'], "produced_kgs": kpi['produced_kgs'], "production_cost": kpi['production_cost'], "consumed_units": kpi['consumed_units'], "consumed_kgs": kpi['consumed_kgs'], "consumption_cost": kpi['consumption_cost'], "produced_items": kpi['produced_items'], "recovery_pct": recovery_pct}, "top_produced": top_produced, "top_consumed": top_consumed, "by_category": by_category, "by_division": by_division, "monthly": monthly, "issue_receipt": balance})
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500



# ============================================================
# INVENTORY DASHBOARD PAGE
# ============================================================

INVENTORY_TABLE = "HDINV90196"

@app.route("/inventory-dashboard")
def inventory_dashboard_page():

    try:

        user = current_user()

        if not user:

            return redirect("/")


        if not dashboard_allowed("INVENTORY"):

            return jsonify({
                "error":
                    "You do not have permission to access "
                    "the Inventory Dashboard."
            }), 403


        return render_template(
            "inventory_dashboard.html"
        )


    except Exception as e:

        return jsonify({
            "error": str(e)
        }), 500

# ============================================================
# INVENTORY DIVISIONS
# ============================================================

@app.get("/api/inventory-divisions")
def api_inventory_divisions():

    try:

        user = current_user()


        if not user:

            return jsonify({

                "error":
                    "Session expired. Please login again."

            }), 401


        if not dashboard_allowed("INVENTORY"):

            return jsonify({

                "error":
                    "You do not have permission to access "
                    "the Inventory Dashboard."

            }), 403


        divisions = normalize_divisions(
            user.get(
                "divisions",
                []
            )
        )


        if not divisions:

            return jsonify({

                "divisions": [],

                "default_division":
                    None

            })


        default_division = (
            get_default_division(
                divisions
            )
        )


        return jsonify({

            "divisions":
                divisions,

            "default_division":
                default_division

        })


    except Exception as e:

        return jsonify({
            "error": str(e)
        }), 500


# ============================================================
# INVENTORY WHERE BUILDER
# ============================================================

def build_inventory_where(
    start_year,
    start_month,
    end_year,
    end_month,
    division,
    user_divisions
):

    where = """

        `YEAR` IS NOT NULL

        AND `MONTH` IS NOT NULL

        AND (

            (`YEAR` > %s)

            OR

            (

                `YEAR` = %s

                AND

                `MONTH` >= %s

            )

        )

        AND (

            (`YEAR` < %s)

            OR

            (

                `YEAR` = %s

                AND

                `MONTH` <= %s

            )

        )

    """


    params = [

        start_year,

        start_year,

        start_month,

        end_year,

        end_year,

        end_month

    ]


    allowed_upper = {

        str(x)
        .strip()
        .upper()

        for x in user_divisions

    }


    # --------------------------------------------------------
    # USER HAS ALL ACCESS
    # --------------------------------------------------------

    if (
        "ALL"
        not in allowed_upper
    ):

        allowed = [

            str(x)
            .strip()

            for x in user_divisions

            if str(x)
            .strip()
            .upper() != "ALL"

        ]


        if not allowed:

            where += """

                AND 1 = 0

            """

            return where, params


        placeholders = ",".join(

            ["%s"] *
            len(allowed)

        )


        where += f"""

            AND `DIVISION` IN
            ({placeholders})

        """


        params.extend(
            allowed
        )


    # --------------------------------------------------------
    # SELECTED DIVISION
    # --------------------------------------------------------

    if (

        division

        and

        str(division)
        .strip()
        .upper() != "ALL"

    ):

        division_value = (
            str(division)
            .strip()
        )


        # Prevent access to an unpermitted division.

        if (

            "ALL"
            not in allowed_upper

            and

            division_value
            .upper()

            not in allowed_upper

        ):

            where += """

                AND 1 = 0

            """

        else:

            where += """

                AND `DIVISION` = %s

            """


            params.append(
                division_value
            )


    return (

        where,

        params

    )


# ============================================================
# INVENTORY DASHBOARD API
# ============================================================

@app.get("/api/inventory-dashboard")
def api_inventory_dashboard():

    try:


        # ----------------------------------------------------
        # USER
        # ----------------------------------------------------

        user = current_user()


        if not user:

            return jsonify({

                "error":
                    "Session expired. Please login again."

            }), 401


        # ----------------------------------------------------
        # PERMISSION
        # ----------------------------------------------------

        if not dashboard_allowed("INVENTORY"):

            return jsonify({

                "error":
                    "You do not have permission to access "
                    "the Inventory Dashboard."

            }), 403


        # ----------------------------------------------------
        # PARAMETERS
        # ----------------------------------------------------

        start_year = int(

            request.args.get(
                "start_year"
            )

        )


        start_month = int(

            request.args.get(
                "start_month"
            )

        )


        end_year = int(

            request.args.get(
                "end_year"
            )

        )


        end_month = int(

            request.args.get(
                "end_month"
            )

        )


        division = (

            request.args.get(
                "division"
            )

            or ""

        ).strip()


        top_n = int(

            request.args.get(
                "top_n",
                10
            )

        )


        # ----------------------------------------------------
        # VALIDATE PERIOD
        # ----------------------------------------------------

        if (

            start_month < 1

            or

            start_month > 12

            or

            end_month < 1

            or

            end_month > 12

        ):

            raise ValueError(
                "Month must be between 1 and 12."
            )


        if (

            start_year,
            start_month

        ) > (

            end_year,
            end_month

        ):

            raise ValueError(
                "Start period cannot be after "
                "end period."
            )


        # ----------------------------------------------------
        # LIMIT TOP N
        # ----------------------------------------------------

        if top_n < 1:

            top_n = 10


        if top_n > 100:

            top_n = 100


        # ----------------------------------------------------
        # USER DIVISIONS
        # ----------------------------------------------------

        user_divisions = normalize_divisions(

            user.get(
                "divisions",
                []
            )

        )


        # ----------------------------------------------------
        # BUILD WHERE
        # ----------------------------------------------------

        where, params = (

            build_inventory_where(

                start_year,

                start_month,

                end_year,

                end_month,

                division,

                user_divisions

            )

        )


        # ----------------------------------------------------
        # TABLE
        # ----------------------------------------------------

        table = quote_identifier(
            INVENTORY_TABLE
        )


        # ====================================================
        # KPI
        # ====================================================

        kpi_sql = f"""

            SELECT


                COUNT(*) AS transactions,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'RECEIPT'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS receipts,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'ISSUE'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS issues,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'RECEIPT'

                            THEN

                                COALESCE(
                                    `QUANTITYINUNITS`,
                                    0
                                )

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'ISSUE'

                            THEN

                                COALESCE(
                                    `QUANTITYINUNITS`,
                                    0
                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS net_movement,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `MATERIALCOST`,
                                0
                            )

                        )

                    ),

                    0

                ) AS transaction_value,


                COALESCE(

                    SUM(

                        CASE

                            WHEN

                                UPPER(

                                    TRIM(

                                        COALESCE(
                                            `TRANSACTIONTYPE`,
                                            ''
                                        )

                                    )

                                ) = 'ISSUE'

                            AND

                                UPPER(

                                    TRIM(

                                        COALESCE(
                                            `TYPEOFORDER`,
                                            ''
                                        )

                                    )

                                )

                                LIKE '%PRODUCTION%'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS production_consumption,


                COUNT(

                    DISTINCT

                    NULLIF(

                        TRIM(

                            COALESCE(
                                `ORIGINALITEM`,
                                ''
                            )

                        ),

                        ''

                    )

                ) AS active_items,


                COUNT(

                    DISTINCT

                    NULLIF(

                        TRIM(

                            COALESCE(
                                `PRODUCTCATEGORY`,
                                ''
                            )

                        ),

                        ''

                    )

                ) AS categories


            FROM {table}


            WHERE {where}

        """


        kpi_result = rows(

            kpi_sql,

            tuple(params)

        )


        kpi = (

            kpi_result[0]

            if kpi_result

            else {}

        )


        # ====================================================
        # TRANSACTION / ORDER TYPE
        # ====================================================

        transaction_sql = f"""

            SELECT


                COALESCE(

                    NULLIF(

                        TRIM(

                            `TYPEOFORDER`

                        ),

                        ''

                    ),

                    'Not Specified'

                ) AS order_type,


                COUNT(*) AS transactions,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'RECEIPT'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS receipts,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'ISSUE'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS issues,


                COALESCE(

                    SUM(

                        COALESCE(
                            `QUANTITYINUNITS`,
                            0
                        )

                    ),

                    0

                ) AS net_movement,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `MATERIALCOST`,
                                0
                            )

                        )

                    ),

                    0

                ) AS transaction_value


            FROM {table}


            WHERE {where}


            GROUP BY

                COALESCE(

                    NULLIF(

                        TRIM(
                            `TYPEOFORDER`
                        ),

                        ''

                    ),

                    'Not Specified'

                )


            ORDER BY

                transaction_value DESC

        """


        transaction_types = rows(

            transaction_sql,

            tuple(params)

        )


        # ====================================================
        # TOP CONSUMED ITEMS
        # ====================================================

        consumed_sql = f"""

            SELECT


                COALESCE(

                    NULLIF(

                        TRIM(`ORIGINALITEM`),

                        ''

                    ),

                    'Not Specified'

                ) AS item,


                MAX(

                    TRIM(

                        COALESCE(
                            `DESCRIPTION`,
                            ''
                        )

                    )

                ) AS description,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `QUANTITYINUNITS`,
                                0
                            )

                        )

                    ),

                    0

                ) AS issued_units,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `QUANTITYINKGS`,
                                0
                            )

                        )

                    ),

                    0

                ) AS issued_kgs,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `MATERIALCOST`,
                                0
                            )

                        )

                    ),

                    0

                ) AS consumption_cost


            FROM {table}


            WHERE {where}


            AND

                UPPER(

                    TRIM(

                        COALESCE(
                            `TRANSACTIONTYPE`,
                            ''
                        )

                    )

                ) = 'ISSUE'


            GROUP BY

                COALESCE(

                    NULLIF(

                        TRIM(`ORIGINALITEM`),

                        ''

                    ),

                    'Not Specified'

                )


            ORDER BY

                issued_units DESC


            LIMIT %s

        """


        consumed_params = (

            list(params)

            +

            [top_n]

        )


        top_consumed_items = rows(

            consumed_sql,

            tuple(consumed_params)

        )


        # ====================================================
        # TOP PRODUCTION MATERIALS
        # ====================================================

        production_material_sql = f"""

            SELECT


                COALESCE(

                    NULLIF(

                        TRIM(`ORIGINALITEM`),

                        ''

                    ),

                    'Not Specified'

                ) AS item,


                MAX(

                    TRIM(

                        COALESCE(
                            `DESCRIPTION`,
                            ''
                        )

                    )

                ) AS description,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `QUANTITYINUNITS`,
                                0
                            )

                        )

                    ),

                    0

                ) AS issued_units,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `QUANTITYINKGS`,
                                0
                            )

                        )

                    ),

                    0

                ) AS issued_kgs,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `MATERIALCOST`,
                                0
                            )

                        )

                    ),

                    0

                ) AS consumption_cost


            FROM {table}


            WHERE {where}


            AND

                UPPER(

                    TRIM(

                        COALESCE(
                            `TRANSACTIONTYPE`,
                            ''
                        )

                    )

                ) = 'ISSUE'


            AND

                UPPER(

                    TRIM(

                        COALESCE(
                            `TYPEOFORDER`,
                            ''
                        )

                    )

                )

                LIKE '%PRODUCTION%'


            GROUP BY

                COALESCE(

                    NULLIF(

                        TRIM(`ORIGINALITEM`),

                        ''

                    ),

                    'Not Specified'

                )


            ORDER BY

                consumption_cost DESC


            LIMIT %s

        """


        production_material_params = (

            list(params)

            +

            [top_n]

        )


        top_production_materials = rows(

            production_material_sql,

            tuple(
                production_material_params
            )

        )


        # ====================================================
        # PRODUCT CATEGORY
        # ====================================================

        category_sql = f"""

            SELECT


                COALESCE(

                    NULLIF(

                        TRIM(`PRODUCTCATEGORY`),

                        ''

                    ),

                    'Not Specified'

                ) AS product_category,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'RECEIPT'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS receipts,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'ISSUE'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS issues,


                COALESCE(

                    SUM(

                        COALESCE(
                            `QUANTITYINUNITS`,
                            0
                        )

                    ),

                    0

                ) AS net_movement,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `MATERIALCOST`,
                                0
                            )

                        )

                    ),

                    0

                ) AS transaction_value


            FROM {table}


            WHERE {where}


            GROUP BY

                COALESCE(

                    NULLIF(

                        TRIM(`PRODUCTCATEGORY`),

                        ''

                    ),

                    'Not Specified'

                )


            ORDER BY

                transaction_value DESC

        """


        category_data = rows(

            category_sql,

            tuple(params)

        )


        # ====================================================
        # DIVISION ANALYSIS
        # ====================================================

        division_sql = f"""

            SELECT


                COALESCE(

                    NULLIF(

                        TRIM(`DIVISION`),

                        ''

                    ),

                    'Not Specified'

                ) AS division,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'RECEIPT'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS receipts,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'ISSUE'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS issues,


                COALESCE(

                    SUM(

                        COALESCE(
                            `QUANTITYINUNITS`,
                            0
                        )

                    ),

                    0

                ) AS net_movement,


                COALESCE(

                    SUM(

                        ABS(

                            COALESCE(
                                `MATERIALCOST`,
                                0
                            )

                        )

                    ),

                    0

                ) AS transaction_value


            FROM {table}


            WHERE {where}


            GROUP BY

                COALESCE(

                    NULLIF(

                        TRIM(`DIVISION`),

                        ''

                    ),

                    'Not Specified'

                )


            ORDER BY

                transaction_value DESC

        """


        division_data = rows(

            division_sql,

            tuple(params)

        )


        # ====================================================
        # MONTHLY MOVEMENT
        # ====================================================

        monthly_sql = f"""

            SELECT


                `YEART` AS year,


                `MONTH` AS month,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'RECEIPT'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS receipts,


                COALESCE(

                    SUM(

                        CASE

                            WHEN UPPER(

                                TRIM(

                                    COALESCE(
                                        `TRANSACTIONTYPE`,
                                        ''
                                    )

                                )

                            ) = 'ISSUE'

                            THEN

                                ABS(

                                    COALESCE(
                                        `QUANTITYINUNITS`,
                                        0
                                    )

                                )

                            ELSE 0

                        END

                    ),

                    0

                ) AS issues,


                COALESCE(

                    SUM(

                        COALESCE(
                            `QUANTITYINUNITS`,
                            0
                        )

                    ),

                    0

                ) AS net_movement,


                COUNT(*) AS transactions


            FROM {table}


            WHERE {where}


            GROUP BY

                `YEART`,

                `MONTH`


            ORDER BY

                `YEART`,

                `MONTH`

        """


        monthly_data = rows(

            monthly_sql,

            tuple(params)

        )


        # ====================================================
        # RESPONSE
        # ====================================================

        return jsonify({

            "kpi": {

                "transactions":

                    int(
                        kpi.get(
                            "transactions",
                            0
                        )
                        or 0
                    ),


                "receipts":

                    float(
                        kpi.get(
                            "receipts",
                            0
                        )
                        or 0
                    ),


                "issues":

                    float(
                        kpi.get(
                            "issues",
                            0
                        )
                        or 0
                    ),


                "net_movement":

                    float(
                        kpi.get(
                            "net_movement",
                            0
                        )
                        or 0
                    ),


                "transaction_value":

                    float(
                        kpi.get(
                            "transaction_value",
                            0
                        )
                        or 0
                    ),


                "production_consumption":

                    float(
                        kpi.get(
                            "production_consumption",
                            0
                        )
                        or 0
                    ),


                "active_items":

                    int(
                        kpi.get(
                            "active_items",
                            0
                        )
                        or 0
                    ),


                "categories":

                    int(
                        kpi.get(
                            "categories",
                            0
                        )
                        or 0
                    )

            },


            "transaction_types":

                transaction_types,


            "top_consumed_items":

                top_consumed_items,


            "top_production_materials":

                top_production_materials,


            "categories":

                category_data,


            "divisions":

                division_data,


            "monthly":

                monthly_data

        })


    except ValueError as e:

        return jsonify({

            "error":
                str(e)

        }), 400


    except Exception as e:

        logging.exception(
            "Inventory Dashboard Error"
        )


        return jsonify({

            "error":
                str(e)

        }), 500





# ============================================================
# LOGIN
# ============================================================

@app.post("/api/login")
def login():

    try:

        data = request.get_json(
            silent=True
        ) or {}

        username = str(
            data.get(
                "username",
                ""
            )
        ).strip()

        password = str(
            data.get(
                "password",
                ""
            )
        )

        if not username or not password:

            return jsonify({
                "success": False,
                "error":
                    "Please enter username and password.",
            }), 400

        user = authenticate_user(
            username,
            password,
        )

        if not user:

            return jsonify({
                "success": False,
                "error":
                    "Invalid username or password.",
            }), 401

        divisions = get_user_divisions(
            username
        )

        if not divisions:

            return jsonify({
                "success": False,
                "error":
                    "Login is valid, but no division permission "
                    "is assigned to this user.",
            }), 403

        session.clear()

        session["username"] = username
        session["divisions"] = divisions

        default_period = get_default_period()

        default_division = get_default_division(
            divisions
        )

        session["default_division"] = (
            default_division
        )

        session["default_start_year"] = (
            default_period["start_year"]
        )

        session["default_start_month"] = (
            default_period["start_month"]
        )

        session["default_end_year"] = (
            default_period["end_year"]
        )

        session["default_end_month"] = (
            default_period["end_month"]
        )

        dashboards = get_user_dashboards({"username": username, "divisions": divisions})

        return jsonify({
            "success": True,
            "username": username,
            "divisions": divisions,
            "default_division": default_division,
            "default_period": default_period,
            "dashboards": dashboards,
            "dashboard_count": len(dashboards),
            "landing_url": dashboards[0]["route"] if len(dashboards) == 1 else "/",
        })

    except Error as e:

        return jsonify({
            "success": False,
            "error":
                f"Database error: {e}",
        }), 500

    except Exception as e:

        return jsonify({
            "success": False,
            "error":
                f"Application error: {e}",
        }), 500


# ============================================================
# SESSION
# ============================================================

@app.get("/api/session")
def api_session():

    user = current_user()

    if not user:

        return jsonify({
            "authenticated": False,
        }), 401

    default_period = get_default_period()

    default_division = get_default_division(
        user["divisions"]
    )

    dashboards = get_user_dashboards(user)

    return jsonify({
        "authenticated": True,
        "username": user["username"],
        "divisions": user["divisions"],
        "default_division": default_division,
        "default_period": default_period,
        "dashboards": dashboards,
        "dashboard_count": len(dashboards),
        "landing_url": get_landing_url(user),
    })


# ============================================================
# LOGOUT
# ============================================================

@app.post("/api/logout")
def logout():

    session.clear()

    return jsonify({
        "success": True,
    })


# ============================================================
# DASHBOARD API
# ============================================================

@app.get("/api/dashboard")
def dashboard():

    try:

        user = current_user()

        if not user:

            return jsonify({
                "error":
                    "Session expired. Please login again.",
            }), 401

        # ====================================================
        # CURRENT DATE
        # ====================================================

        today = date.today()

        # ====================================================
        # PERIOD
        # ====================================================

        raw_sy = request.args.get(
            "start_year"
        )

        raw_sm = request.args.get(
            "start_month"
        )

        raw_ey = request.args.get(
            "end_year"
        )

        raw_em = request.args.get(
            "end_month"
        )

        if (
            raw_sy is None
            or str(raw_sy).strip() == ""
            or raw_sm is None
            or str(raw_sm).strip() == ""
            or raw_ey is None
            or str(raw_ey).strip() == ""
            or raw_em is None
            or str(raw_em).strip() == ""
        ):

            default_period = get_default_period()

            start_year = default_period[
                "start_year"
            ]

            start_month = default_period[
                "start_month"
            ]

            end_year = default_period[
                "end_year"
            ]

            end_month = default_period[
                "end_month"
            ]

        else:

            start_year = raw_sy
            start_month = raw_sm
            end_year = raw_ey
            end_month = raw_em

        sy, sm, ey, em = period(
            start_year,
            start_month,
            end_year,
            end_month,
        )

        # ====================================================
        # TOP N
        # ====================================================

        top = (
            20
            if request.args.get(
                "top_n"
            ) == "20"
            else 10
        )

        # ====================================================
        # GROUP
        # ====================================================

        group_key = request.args.get(
            "category_field",
            "category",
        )

        group = GROUP_FIELDS.get(
            group_key,
            "SELECTIONCODEDESCRIPTION",
        )

        # ====================================================
        # DIVISION
        # ====================================================

        raw_division = request.args.get(
            "division"
        )

        if (
            raw_division is None
            or str(raw_division).strip() == ""
        ):

            division = get_default_division(
                user["divisions"]
            )

        else:

            division = str(
                raw_division
            ).strip()

        if not division:

            return jsonify({
                "error":
                    "No default division is available "
                    "for this user.",
            }), 403

        # ====================================================
        # SECURITY
        # ====================================================

        if not division_allowed(
            division
        ):

            return jsonify({
                "error":
                    "You do not have permission to view division: "
                    + division,
            }), 403

        # ====================================================
        # USER-SPECIFIC SALES SOURCE
        # ====================================================

        sales_source = build_sales_source(
            allowed_divisions=user[
                "divisions"
            ],
            selected_divisions=(
                [division]
                if division.upper() != "ALL"
                else None
            ),
        )

        # ====================================================
        # DATE WHERE
        # ====================================================

        where, params = build_date_where(
            sy,
            sm,
            ey,
            em,
        )

        where, params = add_division_filter(
            where,
            params,
            division,
        )

        params = tuple(params)

        # ====================================================
        # KPI
        # ====================================================

        k = rows(
            f"""
            SELECT

                COALESCE(
                    SUM(`INVOICEDAMOUNT`),
                    0
                ) sales,

                COALESCE(
                    SUM(`DELIVEREDQUANTITY`),
                    0
                ) units,

                COALESCE(
                    SUM(`DELIVEREDQUANTITYINKGS`),
                    0
                ) kgs,

                COALESCE(
                    SUM({MARGIN}),
                    0
                ) margin,

                COUNT(
                    DISTINCT `CUSTOMER`
                ) customers,

                COUNT(
                    DISTINCT `ITEMCODE`
                ) items,

                COUNT(
                    DISTINCT `INVOICENO`
                ) invoices,

                COALESCE(
                    SUM(`ORDERAMOUNT`),
                    0
                ) order_amount,

                COALESCE(
                    SUM(`ORDERDISCOUNT`),
                    0
                ) discounts,

                COALESCE(
                    SUM(`BACKORDERQUANTITY`),
                    0
                ) backorder,

                COALESCE(
                    SUM(`STOCKONHAND`),
                    0
                ) stock

            FROM {sales_source}

            WHERE {where}
            """,
            params,
        )[0]

        sales = float(
            k["sales"] or 0
        )

        units = float(
            k["units"] or 0
        )

        kgs = float(
            k["kgs"] or 0
        )

        margin = float(
            k["margin"] or 0
        )

        customers_count = int(
            k["customers"] or 0
        )

        items_count = int(
            k["items"] or 0
        )

        invoices_count = int(
            k["invoices"] or 0
        )

        order_amount = float(
            k["order_amount"] or 0
        )

        discounts = float(
            k["discounts"] or 0
        )

        backorder = float(
            k["backorder"] or 0
        )

        stock = float(
            k["stock"] or 0
        )

        margin_pct = (
            margin / sales * 100
            if sales
            else 0
        )

        rate = (
            sales / kgs
            if kgs
            else 0
        )

        avg_invoice = (
            sales / invoices_count
            if invoices_count
            else 0
        )

        avg_customer_sales = (
            sales / customers_count
            if customers_count
            else 0
        )

        avg_item_sales = (
            sales / items_count
            if items_count
            else 0
        )

        kpi = {

            "sales":
                sales,

            "units":
                units,

            "kgs":
                kgs,

            "rate":
                rate,

            "margin":
                margin,

            "margin_pct":
                margin_pct,

            "customers":
                customers_count,

            "items":
                items_count,

            "invoices":
                invoices_count,

            "avg_invoice":
                avg_invoice,

            "avg_customer_sales":
                avg_customer_sales,

            "avg_item_sales":
                avg_item_sales,

            "order_amount":
                order_amount,

            "discounts":
                discounts,

            "backorder":
                backorder,

            "stock":
                stock,
        }

        # ====================================================
        # GROUPED DATA
        # ====================================================

        def grouped(
            field,
            limit="",
        ):

            # ------------------------------------------------
            # CATEGORY SPECIAL RULE
            #
            # MI -> ITEMGROUPDESCRIPTION
            # Others -> SELECTIONCODEDESCRIPTION
            # ------------------------------------------------

            if (
                group_key == "category"
                and field == "SELECTIONCODEDESCRIPTION"
            ):

                label_expression = (
                    get_category_expression(
                        division
                    )
                )

            else:

                label_expression = f"""
                    COALESCE(
                        NULLIF(
                            TRIM(`{field}`),
                            ''
                        ),
                        'Not Specified'
                    )
                """

            return rows(
                f"""
                SELECT

                    {label_expression}
                    AS label,

                    COALESCE(
                        SUM(`INVOICEDAMOUNT`),
                        0
                    ) AS sales,

                    COALESCE(
                        SUM(`DELIVEREDQUANTITY`),
                        0
                    ) AS units,

                    COALESCE(
                        SUM(`DELIVEREDQUANTITYINKGS`),
                        0
                    ) AS kgs,

                    COALESCE(
                        SUM({MARGIN}),
                        0
                    ) AS margin

                FROM {sales_source}

                WHERE {where}

                GROUP BY
                    {label_expression}

                ORDER BY
                    sales DESC

                {limit}
                """,
                params,
            )

        def add_metrics(data):

            for r in data:

                current_sales = float(
                    r["sales"] or 0
                )

                current_margin = float(
                    r["margin"] or 0
                )

                current_kgs = float(
                    r["kgs"] or 0
                )

                r["rate"] = (
                    current_sales /
                    current_kgs
                    if current_kgs
                    else 0
                )

                r["margin_pct"] = (
                    current_margin /
                    current_sales * 100
                    if current_sales
                    else 0
                )

            return data

        category = add_metrics(
            grouped(group)
        )

        customers = add_metrics(
            grouped(
                "CUSTOMERNAME",
                f"LIMIT {top}",
            )
        )

        countries = add_metrics(
            grouped(
                "COUNTRYNAME"
            )
        )

        salesreps = add_metrics(
            grouped(
                "SALESREPNAME"
            )
        )

        # ====================================================
        # MONTHLY
        # ====================================================

        monthly = rows(
            f"""
            SELECT

                YEAR(`INVOICEDATE`) AS year,

                MONTH(`INVOICEDATE`) AS month,

                DATE_FORMAT(
                    MIN(`INVOICEDATE`),
                    '%%Y-%%m'
                ) AS period,

                COALESCE(
                    SUM(`INVOICEDAMOUNT`),
                    0
                ) sales,

                COALESCE(
                    SUM(`DELIVEREDQUANTITY`),
                    0
                ) units,

                COALESCE(
                    SUM(`DELIVEREDQUANTITYINKGS`),
                    0
                ) kgs,

                COALESCE(
                    SUM({MARGIN}),
                    0
                ) margin

            FROM {sales_source}

            WHERE {where}

            GROUP BY
                YEAR(`INVOICEDATE`),
                MONTH(`INVOICEDATE`)

            ORDER BY
                YEAR(`INVOICEDATE`),
                MONTH(`INVOICEDATE`)
            """,
            params,
        )

        cumulative = 0

        for r in monthly:

            current_sales = float(
                r["sales"] or 0
            )

            current_margin = float(
                r["margin"] or 0
            )

            current_kgs = float(
                r["kgs"] or 0
            )

            cumulative += current_sales

            r["cumulative"] = cumulative

            r["rate"] = (
                current_sales /
                current_kgs
                if current_kgs
                else 0
            )

            r["margin_pct"] = (
                current_margin /
                current_sales * 100
                if current_sales
                else 0
            )

        # ====================================================
        # ITEMS BY CATEGORY
        #
        # IMPORTANT:
        #
        # MI:
        #     category = ITEMGROUPDESCRIPTION
        #
        # Other:
        #     category = SELECTIONCODEDESCRIPTION
        #
        # ALL:
        #     category changes according to DIVISION
        # ====================================================

        item_category_expression = (
            get_category_expression(
                division
            )
        )

        top_items_category = rows(
            f"""
            SELECT

                {item_category_expression}
                AS category,

                COALESCE(
                    NULLIF(
                        TRIM(`ITEMCODE`),
                        ''
                    ),
                    'Not Specified'
                ) itemcode,

                COALESCE(
                    NULLIF(
                        TRIM(`ITEMDESCRIPTION`),
                        ''
                    ),
                    'Not Specified'
                ) description,

                COALESCE(
                    SUM(`INVOICEDAMOUNT`),
                    0
                ) sales,

                COALESCE(
                    SUM(`DELIVEREDQUANTITY`),
                    0
                ) units,

                COALESCE(
                    SUM(`DELIVEREDQUANTITYINKGS`),
                    0
                ) kgs,

                COALESCE(
                    SUM({MARGIN}),
                    0
                ) margin

            FROM {sales_source}

            WHERE {where}

            GROUP BY
                {item_category_expression},
                `ITEMCODE`,
                `ITEMDESCRIPTION`

            ORDER BY
                category,
                sales DESC
            """,
            params,
        )

        category_rank = None
        category_position = 0

        for r in top_items_category:

            if (
                r["category"]
                != category_rank
            ):

                category_rank = r[
                    "category"
                ]

                category_position = 1

            else:

                category_position += 1

            r["rank"] = (
                category_position
            )

            current_sales = float(
                r["sales"] or 0
            )

            current_margin = float(
                r["margin"] or 0
            )

            current_kgs = float(
                r["kgs"] or 0
            )

            r["rate"] = (
                current_sales /
                current_kgs
                if current_kgs
                else 0
            )

            r["margin_pct"] = (
                current_margin /
                current_sales * 100
                if current_sales
                else 0
            )

        top5_items = [
            r
            for r in top_items_category
            if r["rank"] <= 5
        ]

        # ====================================================
        # TOTAL ROWS
        # ====================================================

        total_row = {

            "label":
                "TOTAL",

            "sales":
                sales,

            "units":
                units,

            "kgs":
                kgs,

            "rate":
                rate,

            "margin":
                margin,

            "margin_pct":
                margin_pct,
        }

        monthly_total = {

            "period":
                "TOTAL",

            "sales":
                sales,

            "cumulative":
                sales,

            "units":
                units,

            "kgs":
                kgs,

            "rate":
                rate,

            "margin":
                margin,

            "margin_pct":
                margin_pct,
        }

        customers_all = add_metrics(
            grouped(
                "CUSTOMERNAME"
            )
        )

        category_all = list(
            category
        )

        items_all = list(
            top_items_category
        )

        # ====================================================
        # RESPONSE
        # ====================================================

        return jsonify({

            "kpi":
                kpi,

            "monthly":
                monthly,

            "monthly_total":
                monthly_total,

            "category":
                category,

            "category_all":
                category_all,

            "category_total":
                total_row,

            "customers":
                customers,

            "customers_all":
                customers_all,

            "customers_total":
                total_row,

            "countries":
                countries,

            "countries_total":
                total_row,

            "salesreps":
                salesreps,

            "salesreps_total":
                total_row,

            "top5_items":
                top5_items,

            "items_all":
                items_all,

            "top5_items_total":
                total_row,

            # ------------------------------------------------
            # ACTIVE FILTERS
            # ------------------------------------------------

            "filters": {

                "start_year":
                    sy,

                "start_month":
                    sm,

                "end_year":
                    ey,

                "end_month":
                    em,

                "division":
                    division,

                "category_field":
                    group_key,

                "top_n":
                    top,
            },

            "default_filters": {

                "start_year":
                    today.year,

                "start_month":
                    today.month,

                "end_year":
                    today.year,

                "end_month":
                    today.month,

                "division":
                    get_default_division(
                        user["divisions"]
                    ),
            },

            "period": {

                "start_year":
                    sy,

                "start_month":
                    sm,

                "end_year":
                    ey,

                "end_month":
                    em,
            },

            "division":
                division,

            # ------------------------------------------------
            # NEW INFORMATION
            # ------------------------------------------------
            # This tells the frontend which physical field
            # is being used as Category.
            # ------------------------------------------------

            "category_source":
                (
                    "ITEMGROUPDESCRIPTION"
                    if (
                        division.upper() == "MI"
                    )
                    else
                    "SELECTIONCODEDESCRIPTION"
                ),
        })

    except ValueError as e:

        return jsonify({
            "error": str(e),
        }), 400

    except Error as e:

        return jsonify({
            "error":
                f"Database error: {e}",
        }), 500

    except Exception as e:

        return jsonify({
            "error":
                f"Application error: {e}",
        }), 500


# ============================================================
# SALES CATEGORY REPORT - MTD AND YTD
# ============================================================

@app.get("/api/category-sales-report")
def api_category_sales_report():

    try:

        # ====================================================
        # USER / SESSION
        # ====================================================

        user = current_user()

        if not user:

            return jsonify({
                "error":
                    "Session expired. Please login again."
            }), 401


        # ====================================================
        # AVAILABLE DIVISIONS
        # ====================================================

        available_divisions = get_report_divisions()


        # ====================================================
        # REQUESTED DIVISIONS
        # ====================================================

        requested = request.args.get("divisions")

        if requested:

            selected_divisions = [
                x.strip()
                for x in requested.split(",")
                if x.strip()
            ]

        else:

            default_division = get_default_report_division()

            selected_divisions = [
                default_division
            ]


        # ====================================================
        # REMOVE DUPLICATES
        # ====================================================

        selected_divisions = list(
            dict.fromkeys(
                selected_divisions
            )
        )


        # ====================================================
        # SECURITY
        # ====================================================

        permitted = {
            str(x).strip().upper()
            for x in user["divisions"]
        }

        all_allowed = "ALL" in permitted


        # ====================================================
        # ALL DIVISIONS REQUEST
        # ====================================================

        if "ALL" in [
            str(x).upper()
            for x in selected_divisions
        ]:

            if not all_allowed:

                return jsonify({
                    "error":
                        "You do not have permission "
                        "to view all divisions."
                }), 403

            selected_divisions = [
                x
                for x in available_divisions
                if str(x).upper() != "ALL"
            ]


        # ====================================================
        # CHECK INDIVIDUAL DIVISION PERMISSION
        # ====================================================

        if not all_allowed:

            for division in selected_divisions:

                if (
                    str(division).strip().upper()
                    not in permitted
                ):

                    return jsonify({
                        "error":
                            "You do not have permission "
                            "to view division: "
                            + str(division)
                    }), 403


        # ====================================================
        # NOTHING SELECTED
        # ====================================================

        if not selected_divisions:

            return jsonify({
                "success": True,
                "today": date.today().isoformat(),
                "available_divisions":
                    available_divisions,
                "selected_divisions": [],
                "divisions": [],
                "grand_total": {},
                "initial_load_single_division": True
            })


        # ====================================================
        # MYSQL CONNECTION
        # ====================================================

        connection = conn()

        cursor = connection.cursor(
            dictionary=True
        )

        today = date.today()


        # ====================================================
        # STORAGE
        # ====================================================

        report_rows = []


        # ====================================================
        # GRAND TOTAL ACCUMULATOR
        #
        # This total is calculated from the actual
        # selected category rows.
        #
        # Price/Kg is recalculated AFTER all Sales
        # and Kgs have been accumulated.
        # ====================================================

        grand_total = {

            "mtd_local": 0.0,
            "mtd_export": 0.0,
            "mtd_total": 0.0,
            "mtd_kgs": 0.0,
            "mtd_material_cost": 0.0,

            "pm_total": 0.0,
            "pm_kgs": 0.0,

            "ytd_local": 0.0,
            "ytd_export": 0.0,
            "ytd_total": 0.0,
            "ytd_kgs": 0.0,
            "ytd_material_cost": 0.0,

            "lytd_total": 0.0,
            "lytd_kgs": 0.0
        }


        # ====================================================
        # 1. CHECK WHETHER MI IS SELECTED
        # ====================================================

        mi_selected = any(
            str(div).strip().upper() == "MI"
            for div in selected_divisions
        )


        # ====================================================
        # 2. GET MI DATA
        #
        # IMPORTANT:
        # SUMMARY_TYPE is read.
        #
        # Stored TOTAL rows are NOT added to report_rows,
        # otherwise they would become categories.
        # ====================================================

        if mi_selected:

            sql_mi = """
                SELECT

                    REPORT_DATE,
                    DIVISION,
                    CATEGORY,
                    SUMMARY_TYPE,

                    MTD_LOCAL,
                    MTD_EXPORT,
                    MTD_TOTAL,
                    MTD_PRICE_KG,
                    MTD_MATERIAL_COST_KG,

                    PM_TOTAL,
                    PM_PRICE_KG,

                    YTD_LOCAL,
                    YTD_EXPORT,
                    YTD_TOTAL,
                    YTD_PRICE_KG,
                    YTD_MATERIAL_COST_KG,

                    LYTD_TOTAL,
                    LYTD_PRICE_KG,

                    MTD_KGS,
                    MTD_MATERIAL_COST,

                    PM_KGS,

                    YTD_KGS,
                    YTD_MATERIAL_COST,

                    LYTD_KGS

                FROM SALES_SUMMARY_MI

                WHERE REPORT_DATE = %s

                  AND DIVISION = 'MI'

                  AND COALESCE(
                        SUMMARY_TYPE,
                        'DETAIL'
                      ) <> 'TOTAL'

                ORDER BY
                    MTD_TOTAL DESC,
                    CATEGORY
            """

            cursor.execute(
                sql_mi,
                (today,)
            )

            mi_rows = cursor.fetchall()

            report_rows.extend(
                mi_rows
            )


        # ====================================================
        # 3. GET NON-MI DIVISIONS
        # ====================================================

        non_mi_divisions = [

            div

            for div in selected_divisions

            if (
                str(div).strip().upper()
                != "MI"
            )
        ]


        # ====================================================
        # GET NON-MI DATA
        # ====================================================

        if non_mi_divisions:

            placeholders = ", ".join(
                ["%s"] * len(
                    non_mi_divisions
                )
            )

            sql_other = f"""

                SELECT

                    REPORT_DATE,
                    DIVISION,
                    CATEGORY,
                    SUMMARY_TYPE,

                    MTD_LOCAL,
                    MTD_EXPORT,
                    MTD_TOTAL,
                    MTD_PRICE_KG,
                    MTD_MATERIAL_COST_KG,

                    PM_TOTAL,
                    PM_PRICE_KG,

                    YTD_LOCAL,
                    YTD_EXPORT,
                    YTD_TOTAL,
                    YTD_PRICE_KG,
                    YTD_MATERIAL_COST_KG,

                    LYTD_TOTAL,
                    LYTD_PRICE_KG,

                    MTD_KGS,
                    MTD_MATERIAL_COST,

                    PM_KGS,

                    YTD_KGS,
                    YTD_MATERIAL_COST,

                    LYTD_KGS

                FROM SALES_SUMMARY

                WHERE REPORT_DATE = %s

                  AND DIVISION IN (
                      {placeholders}
                  )

                  AND COALESCE(
                        SUMMARY_TYPE,
                        'DETAIL'
                      ) <> 'TOTAL'

                ORDER BY
                    DIVISION,
                    MTD_TOTAL DESC,
                    CATEGORY
            """

            params = [
                today
            ]

            params.extend(
                non_mi_divisions
            )

            cursor.execute(
                sql_other,
                params
            )

            other_rows = cursor.fetchall()

            report_rows.extend(
                other_rows
            )


        # ====================================================
        # CLOSE MYSQL
        # ====================================================

        cursor.close()
        connection.close()

        cursor = None
        connection = None


        # ====================================================
        # 4. FINAL SORT
        # ====================================================

        report_rows.sort(
            key=lambda x: (
                str(
                    x.get(
                        "DIVISION",
                        ""
                    )
                ),

                -(
                    float(
                        x.get(
                            "MTD_TOTAL"
                        ) or 0
                    )
                ),

                str(
                    x.get(
                        "CATEGORY",
                        ""
                    )
                )
            )
        )


        # ====================================================
        # 5. HELPER FUNCTION
        #
        # Converts database row values safely to float.
        # ====================================================

        def num(value):

            try:

                return float(
                    value or 0
                )

            except (
                TypeError,
                ValueError
            ):

                return 0.0


        # ====================================================
        # 6. BUILD DIVISION STRUCTURE
        # ====================================================

        divisions = {}


        for row in report_rows:

            division = str(
                row.get(
                    "DIVISION",
                    ""
                )
            ).strip()


            # ------------------------------------------------
            # CREATE DIVISION
            # ------------------------------------------------

            if division not in divisions:

                divisions[division] = {

                    "division": division,

                    "categories": [],

                    "total": {

                        "mtd_local": 0.0,
                        "mtd_export": 0.0,
                        "mtd_total": 0.0,
                        "mtd_price_kg": 0.0,
                        "mtd_material_cost_kg": 0.0,

                        "pm_total": 0.0,
                        "pm_price_kg": 0.0,

                        "ytd_local": 0.0,
                        "ytd_export": 0.0,
                        "ytd_total": 0.0,
                        "ytd_price_kg": 0.0,
                        "ytd_material_cost_kg": 0.0,

                        "lytd_total": 0.0,
                        "lytd_price_kg": 0.0,

                        "mtd_kgs": 0.0,
                        "mtd_material_cost": 0.0,

                        "pm_kgs": 0.0,

                        "ytd_kgs": 0.0,
                        "ytd_material_cost": 0.0,

                        "lytd_kgs": 0.0
                    }
                }


            # ------------------------------------------------
            # READ VALUES
            # ------------------------------------------------

            mtd_local = num(
                row.get("MTD_LOCAL")
            )

            mtd_export = num(
                row.get("MTD_EXPORT")
            )

            mtd_total = num(
                row.get("MTD_TOTAL")
            )

            mtd_kgs = num(
                row.get("MTD_KGS")
            )

            mtd_material_cost = num(
                row.get("MTD_MATERIAL_COST")
            )

            pm_total = num(
                row.get("PM_TOTAL")
            )

            pm_kgs = num(
                row.get("PM_KGS")
            )

            ytd_local = num(
                row.get("YTD_LOCAL")
            )

            ytd_export = num(
                row.get("YTD_EXPORT")
            )

            ytd_total = num(
                row.get("YTD_TOTAL")
            )

            ytd_kgs = num(
                row.get("YTD_KGS")
            )

            ytd_material_cost = num(
                row.get("YTD_MATERIAL_COST")
            )

            lytd_total = num(
                row.get("LYTD_TOTAL")
            )

            lytd_kgs = num(
                row.get("LYTD_KGS")
            )


            # =================================================
            # CATEGORY PRICE / KG
            # =================================================

            mtd_price_kg = (
                mtd_total / mtd_kgs
                if mtd_kgs
                else 0.0
            )

            mtd_material_cost_kg = (
                mtd_material_cost / mtd_kgs
                if mtd_kgs
                else 0.0
            )

            pm_price_kg = (
                pm_total / pm_kgs
                if pm_kgs
                else 0.0
            )

            ytd_price_kg = (
                ytd_total / ytd_kgs
                if ytd_kgs
                else 0.0
            )

            ytd_material_cost_kg = (
                ytd_material_cost / ytd_kgs
                if ytd_kgs
                else 0.0
            )

            lytd_price_kg = (
                lytd_total / lytd_kgs
                if lytd_kgs
                else 0.0
            )


            # =================================================
            # ADD CATEGORY
            # =================================================

            divisions[division]["categories"].append({

                "category":
                    row.get("CATEGORY"),

                "mtd_local":
                    mtd_local,

                "mtd_export":
                    mtd_export,

                "mtd_total":
                    mtd_total,

                "mtd_price_kg":
                    mtd_price_kg,

                "mtd_material_cost_kg":
                    mtd_material_cost_kg,

                "pm_total":
                    pm_total,

                "pm_price_kg":
                    pm_price_kg,

                "ytd_local":
                    ytd_local,

                "ytd_export":
                    ytd_export,

                "ytd_total":
                    ytd_total,

                "ytd_price_kg":
                    ytd_price_kg,

                "ytd_material_cost_kg":
                    ytd_material_cost_kg,

                "lytd_total":
                    lytd_total,

                "lytd_price_kg":
                    lytd_price_kg,

                "mtd_kgs":
                    mtd_kgs,

                "mtd_material_cost":
                    mtd_material_cost,

                "pm_kgs":
                    pm_kgs,

                "ytd_kgs":
                    ytd_kgs,

                "ytd_material_cost":
                    ytd_material_cost,

                "lytd_kgs":
                    lytd_kgs
            })


            # =================================================
            # ACCUMULATE DIVISION TOTAL
            #
            # IMPORTANT:
            # This belongs ONLY to the current division.
            # =================================================

            division_total = divisions[
                division
            ]["total"]


            division_total["mtd_local"] += mtd_local
            division_total["mtd_export"] += mtd_export
            division_total["mtd_total"] += mtd_total

            division_total["mtd_kgs"] += mtd_kgs
            division_total["mtd_material_cost"] += (
                mtd_material_cost
            )

            division_total["pm_total"] += pm_total
            division_total["pm_kgs"] += pm_kgs

            division_total["ytd_local"] += ytd_local
            division_total["ytd_export"] += ytd_export
            division_total["ytd_total"] += ytd_total

            division_total["ytd_kgs"] += ytd_kgs
            division_total["ytd_material_cost"] += (
                ytd_material_cost
            )

            division_total["lytd_total"] += lytd_total
            division_total["lytd_kgs"] += lytd_kgs


            # =================================================
            # ACCUMULATE GRAND TOTAL
            #
            # This combines ALL selected divisions.
            # =================================================

            grand_total["mtd_local"] += mtd_local
            grand_total["mtd_export"] += mtd_export
            grand_total["mtd_total"] += mtd_total

            grand_total["mtd_kgs"] += mtd_kgs
            grand_total["mtd_material_cost"] += (
                mtd_material_cost
            )

            grand_total["pm_total"] += pm_total
            grand_total["pm_kgs"] += pm_kgs

            grand_total["ytd_local"] += ytd_local
            grand_total["ytd_export"] += ytd_export
            grand_total["ytd_total"] += ytd_total

            grand_total["ytd_kgs"] += ytd_kgs
            grand_total["ytd_material_cost"] += (
                ytd_material_cost
            )

            grand_total["lytd_total"] += lytd_total
            grand_total["lytd_kgs"] += lytd_kgs


        # ====================================================
        # 7. CALCULATE EACH DIVISION PRICE / KG
        # ====================================================

        for division_data in divisions.values():

            t = division_data["total"]


            t["mtd_price_kg"] = (
                t["mtd_total"] / t["mtd_kgs"]
                if t["mtd_kgs"]
                else 0.0
            )

            t["mtd_material_cost_kg"] = (
                t["mtd_material_cost"] / t["mtd_kgs"]
                if t["mtd_kgs"]
                else 0.0
            )

            t["pm_price_kg"] = (
                t["pm_total"] / t["pm_kgs"]
                if t["pm_kgs"]
                else 0.0
            )

            t["ytd_price_kg"] = (
                t["ytd_total"] / t["ytd_kgs"]
                if t["ytd_kgs"]
                else 0.0
            )

            t["ytd_material_cost_kg"] = (
                t["ytd_material_cost"] / t["ytd_kgs"]
                if t["ytd_kgs"]
                else 0.0
            )

            t["lytd_price_kg"] = (
                t["lytd_total"] / t["lytd_kgs"]
                if t["lytd_kgs"]
                else 0.0
            )


        # ====================================================
        # 8. CALCULATE GRAND TOTAL PRICE / KG
        #
        # IMPORTANT:
        #
        # We DO NOT average category Price/Kg.
        #
        # We calculate:
        #
        # Total Sales / Total Kgs
        #
        # This gives the correct weighted Price/Kg.
        # ====================================================

        grand_mtd_price_kg = (

            grand_total["mtd_total"]
            /
            grand_total["mtd_kgs"]

            if grand_total["mtd_kgs"]

            else 0.0
        )


        # ====================================================
        # GRAND TOTAL MTD MATERIAL COST / KG
        # ====================================================

        grand_mtd_material_cost_kg = (

            grand_total["mtd_material_cost"]
            /
            grand_total["mtd_kgs"]

            if grand_total["mtd_kgs"]

            else 0.0
        )


        # ====================================================
        # GRAND TOTAL PM PRICE / KG
        # ====================================================

        grand_pm_price_kg = (

            grand_total["pm_total"]
            /
            grand_total["pm_kgs"]

            if grand_total["pm_kgs"]

            else 0.0
        )


        # ====================================================
        # GRAND TOTAL YTD PRICE / KG
        # ====================================================

        grand_ytd_price_kg = (

            grand_total["ytd_total"]
            /
            grand_total["ytd_kgs"]

            if grand_total["ytd_kgs"]

            else 0.0
        )


        # ====================================================
        # GRAND TOTAL YTD MATERIAL COST / KG
        # ====================================================

        grand_ytd_material_cost_kg = (

            grand_total["ytd_material_cost"]
            /
            grand_total["ytd_kgs"]

            if grand_total["ytd_kgs"]

            else 0.0
        )


        # ====================================================
        # GRAND TOTAL LYTD PRICE / KG
        # ====================================================

        grand_lytd_price_kg = (

            grand_total["lytd_total"]
            /
            grand_total["lytd_kgs"]

            if grand_total["lytd_kgs"]

            else 0.0
        )


        # ====================================================
        # 9. CREATE GRAND TOTAL OBJECT
        # ====================================================

        grand_total_response = {

            "division":
                "TOTAL",

            "category":
                "TOTAL",


            # ------------------------------------------------
            # MTD
            # ------------------------------------------------

            "mtd_local":
                grand_total["mtd_local"],

            "mtd_export":
                grand_total["mtd_export"],

            "mtd_total":
                grand_total["mtd_total"],

            "mtd_price_kg":
                grand_mtd_price_kg,

            "mtd_material_cost_kg":
                grand_mtd_material_cost_kg,


            # ------------------------------------------------
            # PM
            # ------------------------------------------------

            "pm_total":
                grand_total["pm_total"],

            "pm_price_kg":
                grand_pm_price_kg,


            # ------------------------------------------------
            # YTD
            # ------------------------------------------------

            "ytd_local":
                grand_total["ytd_local"],

            "ytd_export":
                grand_total["ytd_export"],

            "ytd_total":
                grand_total["ytd_total"],

            "ytd_price_kg":
                grand_ytd_price_kg,

            "ytd_material_cost_kg":
                grand_ytd_material_cost_kg,


            # ------------------------------------------------
            # LYTD
            # ------------------------------------------------

            "lytd_total":
                grand_total["lytd_total"],

            "lytd_price_kg":
                grand_lytd_price_kg,


            # ------------------------------------------------
            # QUANTITIES
            # ------------------------------------------------

            "mtd_kgs":
                grand_total["mtd_kgs"],

            "mtd_material_cost":
                grand_total["mtd_material_cost"],

            "pm_kgs":
                grand_total["pm_kgs"],

            "ytd_kgs":
                grand_total["ytd_kgs"],

            "ytd_material_cost":
                grand_total["ytd_material_cost"],

            "lytd_kgs":
                grand_total["lytd_kgs"]
        }


        # ====================================================
        # 10. RESPONSE DIVISIONS
        #
        # IMPORTANT:
        #
        # DO NOT replace division["total"] with the
        # grand total.
        #
        # Each division keeps its OWN total.
        # ====================================================

        response_divisions = list(
            divisions.values()
        )


        # ====================================================
        # 11. RESPONSE
        # ====================================================

        return jsonify({

            "success":
                True,

            "today":
                today.isoformat(),

            "available_divisions":
                available_divisions,

            "selected_divisions":
                selected_divisions,

            "divisions":
                response_divisions,


            # =================================================
            # GLOBAL GRAND TOTAL
            # =================================================

            "grand_total":
                grand_total_response,


            "initial_load_single_division":
                len(
                    selected_divisions
                ) == 1
        })


    # ========================================================
    # DATABASE ERROR
    # ========================================================

    except Error as e:

        try:

            if cursor:
                cursor.close()

            if connection:
                connection.close()

        except Exception:
            pass


        logging.exception(
            "CATEGORY SALES REPORT DATABASE ERROR"
        )


        return jsonify({

            "error":
                f"Database error: {e}"

        }), 500


    # ========================================================
    # GENERAL ERROR
    # ========================================================

    except Exception as e:

        try:

            if cursor:
                cursor.close()

            if connection:
                connection.close()

        except Exception:
            pass


        logging.exception(
            "CATEGORY SALES REPORT APPLICATION ERROR"
        )


        return jsonify({

            "error":
                f"Application error: {e}"

        }), 500



# =========================================================
# COMPARISON API
# =========================================================

@app.get("/api/comparison")
def comparison():

    try:

        # =====================================================
        # USER / SESSION
        # =====================================================

        user = current_user()

        if not user:
            return jsonify({
                "error":
                    "Session expired. Please login again."
            }), 401


        # =====================================================
        # YEARS
        # =====================================================

        years_text = request.args.get(
            "years",
            ""
        ).strip()

        if not years_text:
            return jsonify({
                "error":
                    "Please select at least one year."
            }), 400


        years = []

        for value in years_text.split(","):

            value = value.strip()

            if not value:
                continue

            try:

                year = int(value)

                if year not in years:
                    years.append(year)

            except ValueError:
                continue


        if not years:
            return jsonify({
                "error":
                    "Invalid comparison years."
            }), 400


        years.sort()


        # =====================================================
        # DIVISION
        # =====================================================

        division = request.args.get(
            "division",
            "ALL"
        ).strip()

        if not division:
            division = "ALL"


        if not division_allowed(
            division
        ):
            return jsonify({
                "error":
                    "You do not have permission to view division: "
                    + division
            }), 403


        division_upper = division.upper()


        # =====================================================
        # DIMENSION
        # =====================================================

        dimension_key = request.args.get(
            "dimension",
            "category"
        ).strip().lower()


        # -----------------------------------------------------
        # Validate dimension
        # -----------------------------------------------------

        if dimension_key not in GROUP_FIELDS:

            return jsonify({
                "error":
                    "Invalid comparison dimension."
            }), 400


        # =====================================================
        # DIMENSION SQL EXPRESSION
        # =====================================================

        if dimension_key == "category":

            # -------------------------------------------------
            # IMPORTANT:
            #
            # MI
            #     ITEMGROUPDESCRIPTION
            #
            # Non-MI
            #     SELECTIONCODEDESCRIPTION
            #
            # ALL
            #     Division-aware CASE expression
            # -------------------------------------------------

            dimension_expression = (
                get_category_expression(
                    division
                )
            )

        else:

            dimension_expression = (
                f"`{GROUP_FIELDS[dimension_key]}`"
            )


        # =====================================================
        # MEASURE
        # =====================================================

        measure = request.args.get(
            "measure",
            "sales"
        ).strip().lower()


        if measure == "units":

            measure_sql = """
                COALESCE(
                    SUM(
                        `DELIVEREDQUANTITY`
                    ),
                    0
                )
            """

        elif measure == "kgs":

            measure_sql = """
                COALESCE(
                    SUM(
                        `DELIVEREDQUANTITYINKGS`
                    ),
                    0
                )
            """

        else:

            measure = "sales"

            measure_sql = """
                COALESCE(
                    SUM(
                        `INVOICEDAMOUNT`
                    ),
                    0
                )
            """


        # =====================================================
        # PERIOD MODE
        # =====================================================

        by_period = (
            request.args.get(
                "period",
                "no"
            ).strip().lower()
            == "yes"
        )


        # =====================================================
        # USER-SPECIFIC SALES SOURCE
        # =====================================================
        #
        # DO NOT USE:
        #
        #     SALES_SOURCE
        #
        # The application builds the source dynamically.
        #
        # =====================================================

        sales_source = build_sales_source(

            allowed_divisions=user.get(
                "divisions",
                []
            ),

            selected_divisions=(
                [division]
                if division_upper != "ALL"
                else None
            )
        )


        # =====================================================
        # WHERE
        # =====================================================

        placeholders = ",".join(
            ["%s"] * len(years)
        )


        where = f"""
            `INVOICEDATE` IS NOT NULL
            AND YEAR(`INVOICEDATE`)
                IN ({placeholders})
        """


        params = list(years)


        # =====================================================
        # DIVISION FILTER
        # =====================================================

        if division_upper != "ALL":

            where += """
                AND UPPER(
                    TRIM(
                        `DIVISION`
                    )
                ) = UPPER(%s)
            """

            params.append(
                division
            )


        # =====================================================
        # SQL
        # =====================================================

        if by_period:

            # -------------------------------------------------
            # MONTHLY COMPARISON
            # -------------------------------------------------
            # IMPORTANT:
            # MySQL ONLY_FULL_GROUP_BY requires the complete
            # month CASE expression to be included in GROUP BY.
            # -------------------------------------------------

            month_expression = """
                CASE
                    WHEN MONTH(`INVOICEDATE`) = 1 THEN 'Jan'
                    WHEN MONTH(`INVOICEDATE`) = 2 THEN 'Feb'
                    WHEN MONTH(`INVOICEDATE`) = 3 THEN 'Mar'
                    WHEN MONTH(`INVOICEDATE`) = 4 THEN 'Apr'
                    WHEN MONTH(`INVOICEDATE`) = 5 THEN 'May'
                    WHEN MONTH(`INVOICEDATE`) = 6 THEN 'Jun'
                    WHEN MONTH(`INVOICEDATE`) = 7 THEN 'Jul'
                    WHEN MONTH(`INVOICEDATE`) = 8 THEN 'Aug'
                    WHEN MONTH(`INVOICEDATE`) = 9 THEN 'Sep'
                    WHEN MONTH(`INVOICEDATE`) = 10 THEN 'Oct'
                    WHEN MONTH(`INVOICEDATE`) = 11 THEN 'Nov'
                    WHEN MONTH(`INVOICEDATE`) = 12 THEN 'Dec'
                END
            """

            sql = f"""
                SELECT

                    YEAR(`INVOICEDATE`) AS year,

                    MONTH(`INVOICEDATE`) AS month,

                    {month_expression} AS period,

                    {dimension_expression} AS label,

                    {measure_sql} AS value

                FROM {sales_source}

                WHERE {where}

                GROUP BY

                    YEAR(`INVOICEDATE`),

                    MONTH(`INVOICEDATE`),

                    {month_expression},

                    {dimension_expression}

                ORDER BY

                    YEAR(`INVOICEDATE`),

                    MONTH(`INVOICEDATE`),

                    value DESC
            """
        else:

            # -------------------------------------------------
            # YEARLY COMPARISON
            # -------------------------------------------------

            sql = f"""
                SELECT

                    YEAR(`INVOICEDATE`) AS year,

                    {dimension_expression}
                        AS label,

                    {measure_sql}
                        AS value

                FROM {sales_source}

                WHERE {where}

                GROUP BY

                    YEAR(`INVOICEDATE`),

                    {dimension_expression}

                ORDER BY

                    YEAR(`INVOICEDATE`),

                    value DESC
            """


        # =====================================================
        # EXECUTE
        # =====================================================

        result = rows(
            sql,
            tuple(params)
        )


        # =====================================================
        # NORMALIZE RESULT
        # =====================================================

        for row in result:

            row["label"] = (
                str(
                    row.get(
                        "label"
                    )
                    or "Not Specified"
                ).strip()
            )

            row["value"] = float(
                row.get(
                    "value"
                )
                or 0
            )


        # =====================================================
        # LABELS
        # =====================================================

        labels = []

        for row in result:

            label = row["label"]

            if label not in labels:
                labels.append(label)


        # =====================================================
        # PERIOD OUTPUT
        # =====================================================

        if by_period:

            periods = []

            for row in result:

                key = (
                    int(row["year"]),
                    int(row["month"])
                )

                if key not in periods:
                    periods.append(key)


            periods.sort()


            output = []


            for year, month in periods:

                row = {
                    "period":
                        f"{year:04d}-{month:02d}"
                }


                # -------------------------------------------------
                # Initialize every category/item/etc. to ZERO
                # -------------------------------------------------

                for label in labels:

                    row[label] = 0


                # -------------------------------------------------
                # Fill actual values
                # -------------------------------------------------

                for result_row in result:

                    if (
                        int(result_row["year"]) == year
                        and
                        int(result_row["month"]) == month
                    ):

                        row[
                            result_row["label"]
                        ] = float(
                            result_row["value"]
                            or 0
                        )


                output.append(row)


        # =====================================================
        # YEAR OUTPUT
        # =====================================================

        else:

            output = []


            for year in years:

                row = {
                    "year": year
                }


                # -------------------------------------------------
                # Initialize every label
                # -------------------------------------------------

                for label in labels:

                    row[label] = 0


                # -------------------------------------------------
                # Fill actual values
                # -------------------------------------------------

                for result_row in result:

                    if (
                        int(result_row["year"])
                        == year
                    ):

                        row[
                            result_row["label"]
                        ] = float(
                            result_row["value"]
                            or 0
                        )


                output.append(row)


        # =====================================================
        # RESPONSE
        # =====================================================

        return jsonify({

            "success":
                True,

            "years":
                years,

            "dimension":
                dimension_key,

            "measure":
                measure,

            "period":
                by_period,

            "labels":
                labels,

            "rows":
                output,

            "raw":
                result,

            "division":
                division,

            "category_source":
                (
                    "ITEMGROUPDESCRIPTION"
                    if division_upper == "MI"
                    else
                    (
                        "ITEMGROUPDESCRIPTION / "
                        "SELECTIONCODEDESCRIPTION"
                        if division_upper == "ALL"
                        else
                        "SELECTIONCODEDESCRIPTION"
                    )
                )

        })


    # =========================================================
    # DATABASE ERROR
    # =========================================================

    except Error as e:

        return jsonify({
            "error":
                f"Database error: {e}"
        }), 500


    # =========================================================
    # APPLICATION ERROR
    # =========================================================

    except Exception as e:

        return jsonify({
            "error":
                f"Application error: {e}"
        }), 500




# =========================================================
# EXCEL EXPORT
# =========================================================

@app.post("/api/export/excel")
def export_excel():

    try:

        # =====================================================
        # SESSION CHECK
        # =====================================================

        if not current_user():

            return jsonify({
                "error": "Session expired. Please login again."
            }), 401


        # =====================================================
        # OPENPYXL CHECK
        # =====================================================

        if Workbook is None:

            return jsonify({
                "error":
                    "Excel export requires openpyxl. "
                    "Please install it with: "
                    "pip install openpyxl"
            }), 500


        # =====================================================
        # PAYLOAD
        # =====================================================

        payload = request.get_json(
            silent=True
        ) or {}

        data = payload.get(
            "data"
        ) or {}


        # =====================================================
        # DIVISION
        # =====================================================

        division = str(
            payload.get(
                "division",
                "ALL"
            )
        ).strip()


        if not division_allowed(
            division
        ):

            return jsonify({
                "error":
                    "You do not have permission to export division: "
                    + division
            }), 403


        # =====================================================
        # WORKBOOK
        # =====================================================

        wb = Workbook()

        default_ws = wb.active

        wb.remove(
            default_ws
        )


        # =====================================================
        # STYLES
        # =====================================================

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        header_font = Font(
            color="FFFFFF",
            bold=True
        )

        total_fill = PatternFill(
            fill_type="solid",
            fgColor="FFF2CC"
        )

        total_font = Font(
            bold=True
        )

        title_font = Font(
            size=14,
            bold=True
        )

        thin = Side(
            style="thin",
            color="B7B7B7"
        )

        border = Border(
            bottom=thin
        )


        # =====================================================
        # TITLE
        # =====================================================

        def add_title(
            ws,
            title
        ):

            ws.cell(
                row=1,
                column=1,
                value=title
            )

            ws.cell(
                row=1,
                column=1
            ).font = title_font

            ws.merge_cells(
                start_row=1,
                start_column=1,
                end_row=1,
                end_column=10
            )


            ws.cell(
                row=2,
                column=1,
                value="User"
            )

            ws.cell(
                row=2,
                column=2,
                value=current_user()["username"]
            )


            ws.cell(
                row=3,
                column=1,
                value="Division"
            )

            ws.cell(
                row=3,
                column=2,
                value=division
            )


            ws.cell(
                row=4,
                column=1,
                value="Period"
            )

            ws.cell(
                row=4,
                column=2,
                value=str(
                    payload.get(
                        "period",
                        ""
                    )
                )
            )


        # =====================================================
        # HEADER STYLE
        # =====================================================

        def style_header(
            ws,
            row,
            start_col,
            end_col
        ):

            for col in range(
                start_col,
                end_col + 1
            ):

                cell = ws.cell(
                    row=row,
                    column=col
                )

                cell.fill = header_fill
                cell.font = header_font

                cell.alignment = Alignment(
                    horizontal="center"
                )

                cell.border = border


        # =====================================================
        # AUTOSIZE
        # =====================================================

        def autosize(ws):

            for col_cells in ws.columns:

                max_len = 0

                col_index = (
                    col_cells[0].column
                )

                for cell in col_cells:

                    value = (
                        ""
                        if cell.value is None
                        else str(cell.value)
                    )

                    max_len = max(
                        max_len,
                        len(value)
                    )

                ws.column_dimensions[
                    get_column_letter(
                        col_index
                    )
                ].width = min(
                    max(max_len + 2, 12),
                    35
                )


        # =====================================================
        # SUMMARY
        # =====================================================

        def add_summary():

            kpi = data.get(
                "kpi",
                {}
            )

            ws = wb.create_sheet(
                "KPI Summary"
            )

            add_title(
                ws,
                "SalesMaster KPI Summary"
            )


            headers = [
                "KPI",
                "Value"
            ]


            for c, value in enumerate(
                headers,
                1
            ):

                ws.cell(
                    row=6,
                    column=c,
                    value=value
                )


            style_header(
                ws,
                6,
                1,
                2
            )


            rows_data = [

                (
                    "Sales Value",
                    kpi.get("sales", 0)
                ),

                (
                    "Delivered Units",
                    kpi.get("units", 0)
                ),

                (
                    "Delivered KG",
                    kpi.get("kgs", 0)
                ),

                (
                    "Avg KG Rate",
                    kpi.get("rate", 0)
                ),

                (
                    "Margin",
                    kpi.get("margin", 0)
                ),

                (
                    "Margin %",
                    kpi.get("margin_pct", 0)
                ),

                (
                    "Customers",
                    kpi.get("customers", 0)
                ),

                (
                    "Items",
                    kpi.get("items", 0)
                ),

                (
                    "Invoices",
                    kpi.get("invoices", 0)
                ),

                (
                    "Average Invoice Value",
                    kpi.get("avg_invoice", 0)
                ),

                (
                    "Average Customer Sales",
                    kpi.get("avg_customer_sales", 0)
                ),

                (
                    "Average Item Sales",
                    kpi.get("avg_item_sales", 0)
                ),

                (
                    "Order Amount",
                    kpi.get("order_amount", 0)
                ),

                (
                    "Discounts",
                    kpi.get("discounts", 0)
                ),

                (
                    "Backorder Quantity",
                    kpi.get("backorder", 0)
                ),

                (
                    "Stock on Hand",
                    kpi.get("stock", 0)
                ),
            ]


            for r, (label, value) in enumerate(
                rows_data,
                7
            ):

                ws.cell(
                    row=r,
                    column=1,
                    value=label
                )

                ws.cell(
                    row=r,
                    column=2,
                    value=float(value or 0)
                )

                ws.cell(
                    row=r,
                    column=2
                ).number_format = '#,##0.00'


            ws.freeze_panes = "A7"

            autosize(ws)


        # =====================================================
        # STANDARD REPORT SHEET
        # =====================================================

        def add_report_sheet(
            sheet_name,
            title,
            records,
            monthly=False,
            total=None
        ):

            ws = wb.create_sheet(
                sheet_name[:31]
            )

            add_title(
                ws,
                title
            )


            if monthly:

                headers = [
                    "Month",
                    "Sales Value",
                    "Cumulative Sales",
                    "Units",
                    "KG",
                    "Avg KG Rate",
                    "Margin",
                    "Margin %"
                ]

            else:

                headers = [
                    "Name",
                    "Sales Value",
                    "Units",
                    "KG",
                    "Avg KG Rate",
                    "Margin",
                    "Margin %"
                ]


            header_row = 6


            for c, value in enumerate(
                headers,
                1
            ):

                ws.cell(
                    row=header_row,
                    column=c,
                    value=value
                )


            style_header(
                ws,
                header_row,
                1,
                len(headers)
            )


            row_no = 7


            for record in records or []:

                if monthly:

                    values = [

                        record.get(
                            "period",
                            ""
                        ),

                        record.get(
                            "sales",
                            0
                        ),

                        record.get(
                            "cumulative",
                            0
                        ),

                        record.get(
                            "units",
                            0
                        ),

                        record.get(
                            "kgs",
                            0
                        ),

                        record.get(
                            "rate",
                            0
                        ),

                        record.get(
                            "margin",
                            0
                        ),

                        record.get(
                            "margin_pct",
                            0
                        ),
                    ]

                else:

                    values = [

                        record.get(
                            "label",
                            ""
                        ),

                        record.get(
                            "sales",
                            0
                        ),

                        record.get(
                            "units",
                            0
                        ),

                        record.get(
                            "kgs",
                            0
                        ),

                        record.get(
                            "rate",
                            0
                        ),

                        record.get(
                            "margin",
                            0
                        ),

                        record.get(
                            "margin_pct",
                            0
                        ),
                    ]


                for c, value in enumerate(
                    values,
                    1
                ):

                    cell = ws.cell(
                        row=row_no,
                        column=c,
                        value=value
                    )

                    if (
                        c > 1
                        and isinstance(
                            value,
                            (int, float)
                        )
                    ):

                        cell.number_format = '#,##0.00'


                row_no += 1


            # =================================================
            # TOTAL
            # =================================================

            if total:

                if monthly:

                    values = [

                        "TOTAL",

                        total.get(
                            "sales",
                            0
                        ),

                        total.get(
                            "cumulative",
                            0
                        ),

                        total.get(
                            "units",
                            0
                        ),

                        total.get(
                            "kgs",
                            0
                        ),

                        total.get(
                            "rate",
                            0
                        ),

                        total.get(
                            "margin",
                            0
                        ),

                        total.get(
                            "margin_pct",
                            0
                        ),
                    ]

                else:

                    values = [

                        "TOTAL",

                        total.get(
                            "sales",
                            0
                        ),

                        total.get(
                            "units",
                            0
                        ),

                        total.get(
                            "kgs",
                            0
                        ),

                        total.get(
                            "rate",
                            0
                        ),

                        total.get(
                            "margin",
                            0
                        ),

                        total.get(
                            "margin_pct",
                            0
                        ),
                    ]


                for c, value in enumerate(
                    values,
                    1
                ):

                    cell = ws.cell(
                        row=row_no,
                        column=c,
                        value=value
                    )

                    cell.font = total_font
                    cell.fill = total_fill

                    if (
                        c > 1
                        and isinstance(
                            value,
                            (int, float)
                        )
                    ):

                        cell.number_format = '#,##0.00'


            ws.freeze_panes = "A7"


            if row_no > 7:

                ws.auto_filter.ref = (
                    f"A6:"
                    f"{get_column_letter(len(headers))}"
                    f"{row_no - 1}"
                )


            autosize(ws)


        # =====================================================
        # ITEMS BY CATEGORY
        # COMPLETE LIST - NO TOP 5
        # =====================================================

        def add_item_sheet(
            records,
            title="All Items by Category",
            sheet_name="Items by Category"
        ):

            ws = wb.create_sheet(
                sheet_name[:31]
            )

            add_title(
                ws,
                title
            )


            headers = [
                "Category",
                "Item Code",
                "Item Description",
                "Sales Value",
                "Quantity",
                "Quantity KG",
                "Avg KG Rate",
                "Margin",
                "Margin %"
            ]


            for c, value in enumerate(
                headers,
                1
            ):

                ws.cell(
                    row=6,
                    column=c,
                    value=value
                )


            style_header(
                ws,
                6,
                1,
                len(headers)
            )


            row_no = 7


            # =================================================
            # WRITE EVERY ITEM
            # =================================================

            for r in records or []:

                values = [

                    r.get(
                        "category",
                        ""
                    ),

                    r.get(
                        "itemcode",
                        ""
                    ),

                    r.get(
                        "description",
                        ""
                    ),

                    r.get(
                        "sales",
                        0
                    ),

                    r.get(
                        "units",
                        0
                    ),

                    r.get(
                        "kgs",
                        0
                    ),

                    r.get(
                        "rate",
                        0
                    ),

                    r.get(
                        "margin",
                        0
                    ),

                    r.get(
                        "margin_pct",
                        0
                    )
                ]


                for c, value in enumerate(
                    values,
                    1
                ):

                    cell = ws.cell(
                        row=row_no,
                        column=c,
                        value=value
                    )

                    if c >= 4:

                        cell.number_format = (
                            '#,##0.00'
                        )


                row_no += 1


            ws.freeze_panes = "A7"


            if row_no > 7:

                ws.auto_filter.ref = (
                    f"A6:I{row_no - 1}"
                )


            autosize(ws)


        # =====================================================
        # ADD KPI SUMMARY
        # =====================================================

        add_summary()


        # =====================================================
        # MONTHLY SALES
        # =====================================================

        add_report_sheet(
            "Monthly Sales",
            "Monthly Sales",
            data.get(
                "monthly",
                []
            ),
            True,
            data.get(
                "monthly_total"
            )
        )


        # =====================================================
        # CUSTOMERS
        # ALWAYS COMPLETE LIST
        # =====================================================

        customer_records = data.get(
            "customers_all",
            data.get(
                "customers",
                []
            )
        )


        add_report_sheet(
            "Customers",
            "Sales by Customer - Complete List",
            customer_records,
            False,
            data.get(
                "customers_total"
            )
        )


        # =====================================================
        # CATEGORY
        # ALWAYS COMPLETE LIST
        # =====================================================

        category_records = data.get(
            "category_all",
            data.get(
                "category",
                []
            )
        )


        add_report_sheet(
            "Category",
            "Sales by Category - Complete List",
            category_records,
            False,
            data.get(
                "category_total"
            )
        )


        # =====================================================
        # COUNTRIES
        # =====================================================

        add_report_sheet(
            "Countries",
            "Sales by Country",
            data.get(
                "countries",
                []
            ),
            False,
            data.get(
                "countries_total"
            )
        )


        # =====================================================
        # SALES REPRESENTATIVES
        # =====================================================

        add_report_sheet(
            "Sales Reps",
            "Sales by Sales Representative",
            data.get(
                "salesreps",
                []
            ),
            False,
            data.get(
                "salesreps_total"
            )
        )


        # =====================================================
        # ITEMS
        # ALWAYS COMPLETE LIST
        # =====================================================

        item_records = data.get(
            "items_all",
            data.get(
                "top_items_category",
                []
            )
        )


        add_item_sheet(
            item_records,
            "All Items by Category - Complete List",
            "Items by Category"
        )


        # =====================================================
        # COMPARISON DATA
        # =====================================================

        comparison_data = payload.get(
            "comparison"
        )


        if comparison_data:

            ws = wb.create_sheet(
                "Comparison"
            )

            add_title(
                ws,
                "Sales Comparison"
            )


            labels = comparison_data.get(
                "labels",
                []
            )


            headers = []


            if comparison_data.get(
                "period"
            ):

                headers.append(
                    "Period"
                )

            else:

                headers.append(
                    "Year"
                )


            headers.extend(
                labels
            )


            for c, value in enumerate(
                headers,
                1
            ):

                ws.cell(
                    row=6,
                    column=c,
                    value=value
                )


            style_header(
                ws,
                6,
                1,
                len(headers)
            )


            row_no = 7


            for record in comparison_data.get(
                "rows",
                []
            ):

                if comparison_data.get(
                    "period"
                ):

                    values = [
                        record.get(
                            "period",
                            ""
                        )
                    ]

                else:

                    values = [
                        record.get(
                            "year",
                            ""
                        )
                    ]


                for label in labels:

                    values.append(
                        record.get(
                            label,
                            0
                        )
                    )


                for c, value in enumerate(
                    values,
                    1
                ):

                    cell = ws.cell(
                        row=row_no,
                        column=c,
                        value=value
                    )


                    if c > 1:

                        cell.number_format = (
                            '#,##0.00'
                        )


                row_no += 1


            autosize(ws)


        # =====================================================
        # SAVE EXCEL
        # =====================================================

        output = BytesIO()

        wb.save(
            output
        )

        output.seek(0)


        filename = (
            "SalesMaster_KPI_Report.xlsx"
        )


        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )


    # =========================================================
    # ERROR
    # =========================================================



    except Exception as e:

        return jsonify({
            "error":
                f"Excel export error: {e}"
        }), 500



# =========================================================
# FINANCE AGEING - EXPORT EXCEL
# =========================================================

@app.route(
    "/api/finance-ageing/export",
    methods=["GET"]
)
def export_finance_ageing():

    conn = None
    cursor = None

    try:

        # =====================================================
        # GET FILTER PARAMETERS
        # =====================================================

        division = request.args.get(
            "division",
            ""
        ).strip()


        customer = request.args.get(
            "customer",
            ""
        ).strip()


        salesman = request.args.get(
            "salesman",
            ""
        ).strip()


        search = request.args.get(
            "search",
            ""
        ).strip()


        # =====================================================
        # DATABASE CONNECTION
        # =====================================================

        conn = get_db_connection()


        cursor = conn.cursor(
            dictionary=True
        )


        # =====================================================
        # BASE SQL
        #
        # IMPORTANT:
        #
        # USE THE SAME TABLE / VIEW USED BY:
        #
        # /api/finance-ageing/data
        # =====================================================

        sql = """

            SELECT

                DIVISION,

                CUSTOMERCODE,

                CUSTOMERNAME,

                SALESMANNAME,

                CREDITLIMIT,

                ACTUALOVERDUE,

                DAYS30,

                DAYS60,

                DAYS90,

                DAYS120,

                DAYS150,

                DAYS180,

                DAYS270,

                DAYS365,

                DAYSABOVE365,

                OUTSTANDING,

                TERMSDAYS

            FROM HFACR200

            WHERE 1 = 1

        """


        params = []


        # =====================================================
        # DIVISION FILTER
        # =====================================================

        if division:

            sql += """

                AND DIVISION = %s

            """

            params.append(
                division
            )


        # =====================================================
        # CUSTOMER FILTER
        # =====================================================

        if customer:

            sql += """

                AND CUSTOMERCODE = %s

            """

            params.append(
                customer
            )


        # =====================================================
        # SALESMAN FILTER
        # =====================================================

        if salesman:

            sql += """

                AND SALESMANNAME = %s

            """

            params.append(
                salesman
            )


        # =====================================================
        # CUSTOMER SEARCH
        # =====================================================

        if search:

            sql += """

                AND (

                    LOWER(
                        COALESCE(
                            CUSTOMERCODE,
                            ''
                        )
                    ) LIKE %s

                    OR

                    LOWER(
                        COALESCE(
                            CUSTOMERNAME,
                            ''
                        )
                    ) LIKE %s

                )

            """


            search_value = (

                "%"

                +

                search.lower()

                +

                "%"

            )


            params.append(
                search_value
            )


            params.append(
                search_value
            )


        # =====================================================
        # ORDER BY
        # =====================================================

        sql += """

            ORDER BY

                OUTSTANDING DESC

        """


        # =====================================================
        # EXECUTE QUERY
        # =====================================================

        cursor.execute(
            sql,
            params
        )


        rows = cursor.fetchall()


        # =====================================================
        # NO DATA
        # =====================================================

        if not rows:

            return jsonify({

                "error":

                "No Finance Ageing records found for the selected filters."

            }), 404


        # =====================================================
        # CREATE DATAFRAME
        # =====================================================

        df = pd.DataFrame(
            rows
        )


        # =====================================================
        # NUMERIC COLUMNS
        # =====================================================

        numeric_columns = [

            "CREDITLIMIT",

            "ACTUALOVERDUE",

            "DAYS30",

            "DAYS60",

            "DAYS90",

            "DAYS120",

            "DAYS150",

            "DAYS180",

            "DAYS270",

            "DAYS365",

            "DAYSABOVE365",

            "OUTSTANDING",

            "TERMSDAYS"

        ]


        for column in numeric_columns:

            if column in df.columns:

                df[column] = pd.to_numeric(

                    df[column],

                    errors="coerce"

                ).fillna(0)


        # =====================================================
        # CREATE CALCULATED AGEING COLUMNS
        # =====================================================

        df["AGEING_91_180"] = (

            df["DAYS120"]

            +

            df["DAYS150"]

            +

            df["DAYS180"]

        )


        df["AGEING_181_365"] = (

            df["DAYS270"]

            +

            df["DAYS365"]

        )


        df["AGEING_ABOVE_180"] = (

            df["DAYS270"]

            +

            df["DAYS365"]

            +

            df["DAYSABOVE365"]

        )


        # =====================================================
        # CREATE EXPORT DATAFRAME
        #
        # SAME STRUCTURE AS DASHBOARD TABLE
        # =====================================================

        export_df = pd.DataFrame({

            "Division":

                df["DIVISION"],


            "Customer Code":

                df["CUSTOMERCODE"],


            "Customer Name":

                df["CUSTOMERNAME"],


            "Salesman":

                df["SALESMANNAME"],


            "Credit Limit":

                df["CREDITLIMIT"],


            "Actual Overdue":

                df["ACTUALOVERDUE"],


            "0-30 Days":

                df["DAYS30"],


            "31-60 Days":

                df["DAYS60"],


            "61-90 Days":

                df["DAYS90"],


            "91-180 Days":

                df["AGEING_91_180"],


            "181-365 Days":

                df["AGEING_181_365"],


            "Above 365 Days":

                df["DAYSABOVE365"],


            "Above 180 Days":

                df["AGEING_ABOVE_180"],


            "Outstanding":

                df["OUTSTANDING"],


            "Terms Days":

                df["TERMSDAYS"]

        })


        # =====================================================
        # KPI CALCULATIONS
        # =====================================================

        total_outstanding = float(

            df["OUTSTANDING"].sum()

        )


        total_overdue = float(

            df["ACTUALOVERDUE"].sum()

        )


        total_credit_limit = float(

            df["CREDITLIMIT"].sum()

        )


        total_customers = len(df)


        overdue_percentage = (

            total_overdue

            /

            total_outstanding

            *

            100

            if total_outstanding != 0

            else 0

        )


        credit_utilisation = (

            total_outstanding

            /

            total_credit_limit

            *

            100

            if total_credit_limit != 0

            else 0

        )


        # =====================================================
        # SUMMARY DATA
        # =====================================================

        summary_data = {

            "Description": [

                "Total Outstanding",

                "Actual Overdue",

                "Overdue Percentage",

                "Total Credit Limit",

                "Credit Utilisation",

                "Total Customers",

                "0-30 Days",

                "31-60 Days",

                "61-90 Days",

                "91-180 Days",

                "Above 180 Days",

                "Export Date"

            ],

            "Value": [

                total_outstanding,

                total_overdue,

                overdue_percentage / 100,

                total_credit_limit,

                credit_utilisation / 100,

                total_customers,

                float(
                    df["DAYS30"].sum()
                ),

                float(
                    df["DAYS60"].sum()
                ),

                float(
                    df["DAYS90"].sum()
                ),

                float(
                    df["AGEING_91_180"].sum()
                ),

                float(
                    df["AGEING_ABOVE_180"].sum()
                ),

                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

            ]

        }


        summary_df = pd.DataFrame(

            summary_data

        )


        # =====================================================
        # FILTER INFORMATION
        # =====================================================

        filter_data = {

            "Filter": [

                "Division",

                "Customer",

                "Salesman",

                "Customer Search"

            ],

            "Selected Value": [

                division

                if division

                else "All Divisions",


                customer

                if customer

                else "All Customers",


                salesman

                if salesman

                else "All Salesmen",


                search

                if search

                else "No Search"

            ]

        }


        filter_df = pd.DataFrame(

            filter_data

        )


        # =====================================================
        # CREATE EXCEL FILE
        # =====================================================

        output = BytesIO()


        with pd.ExcelWriter(

            output,

            engine="openpyxl"

        ) as writer:


            # =================================================
            # SUMMARY SHEET
            # =================================================

            summary_df.to_excel(

                writer,

                sheet_name="Summary",

                index=False

            )


            # =================================================
            # FILTER SHEET
            # =================================================

            filter_df.to_excel(

                writer,

                sheet_name="Filters",

                index=False

            )


            # =================================================
            # DETAIL SHEET
            # =================================================

            export_df.to_excel(

                writer,

                sheet_name="Finance Ageing",

                index=False

            )


            # =================================================
            # WORKBOOK
            # =================================================

            workbook = writer.book


            # =================================================
            # SUMMARY SHEET
            # =================================================

            summary_sheet = workbook[
                "Summary"
            ]


            summary_sheet.freeze_panes = "A2"


            summary_sheet.auto_filter.ref = (

                summary_sheet.dimensions

            )


            summary_sheet.column_dimensions[
                "A"
            ].width = 30


            summary_sheet.column_dimensions[
                "B"
            ].width = 22


            # =================================================
            # FILTER SHEET
            # =================================================

            filter_sheet = workbook[
                "Filters"
            ]


            filter_sheet.freeze_panes = "A2"


            filter_sheet.auto_filter.ref = (

                filter_sheet.dimensions

            )


            filter_sheet.column_dimensions[
                "A"
            ].width = 25


            filter_sheet.column_dimensions[
                "B"
            ].width = 35


            # =================================================
            # DETAIL SHEET
            # =================================================

            detail_sheet = workbook[
                "Finance Ageing"
            ]


            # =================================================
            # FREEZE HEADER
            # =================================================

            detail_sheet.freeze_panes = "A2"


            # =================================================
            # AUTO FILTER
            # =================================================

            detail_sheet.auto_filter.ref = (

                detail_sheet.dimensions

            )


            # =================================================
            # FORMAT HEADERS
            # =================================================

            from openpyxl.styles import Font


            for cell in summary_sheet[1]:

                cell.font = Font(
                    bold=True
                )


            for cell in filter_sheet[1]:

                cell.font = Font(
                    bold=True
                )


            for cell in detail_sheet[1]:

                cell.font = Font(
                    bold=True
                )


            # =================================================
            # AUTO COLUMN WIDTH
            # =================================================

            for column_cells in detail_sheet.columns:


                max_length = 0


                column_letter = (

                    column_cells[0].column_letter

                )


                for cell in column_cells:


                    try:


                        value_length = len(

                            str(
                                cell.value
                                if cell.value is not None
                                else ""
                            )

                        )


                        if value_length > max_length:

                            max_length = value_length


                    except Exception:

                        pass


                adjusted_width = min(

                    max(
                        max_length + 2,
                        12
                    ),

                    35

                )


                detail_sheet.column_dimensions[
                    column_letter
                ].width = adjusted_width


            # =================================================
            # NUMBER FORMAT - DETAIL
            # =================================================

            numeric_headers = [

                "Credit Limit",

                "Actual Overdue",

                "0-30 Days",

                "31-60 Days",

                "61-90 Days",

                "91-180 Days",

                "181-365 Days",

                "Above 365 Days",

                "Above 180 Days",

                "Outstanding"

            ]


            for row in detail_sheet.iter_rows(

                min_row=2

            ):


                for cell in row:


                    header = detail_sheet.cell(

                        row=1,

                        column=cell.column

                    ).value


                    if header in numeric_headers:

                        cell.number_format = (

                            '#,##0.00'

                        )


            # =================================================
            # SUMMARY NUMBER FORMATS
            # =================================================

            summary_sheet["B2"].number_format = (
                '#,##0.00'
            )


            summary_sheet["B3"].number_format = (
                '#,##0.00'
            )


            summary_sheet["B4"].number_format = (
                '0.00%'
            )


            summary_sheet["B5"].number_format = (
                '#,##0.00'
            )


            summary_sheet["B6"].number_format = (
                '0.00%'
            )


            summary_sheet["B7"].number_format = (
                '#,##0'
            )


            for row_number in range(

                8,

                13

            ):

                summary_sheet.cell(

                    row=row_number,

                    column=2

                ).number_format = (

                    '#,##0.00'

                )


        # =====================================================
        # CLOSE DATABASE
        # =====================================================

        if cursor:

            cursor.close()


        if conn:

            conn.close()


        # =====================================================
        # RESET FILE POINTER
        # =====================================================

        output.seek(0)


        # =====================================================
        # FILE NAME
        # =====================================================

        filename = (

            "Finance_Ageing_"

            +

            datetime.now().strftime(

                "%Y%m%d_%H%M%S"

            )

            +

            ".xlsx"

        )


        # =====================================================
        # SEND FILE
        # =====================================================

        return send_file(

            output,

            as_attachment=True,

            download_name=filename,

            mimetype=(

                "application/vnd.openxmlformats-officedocument."

                "spreadsheetml.sheet"

            )

        )


    except Exception as e:


        print(

            "FINANCE AGEING EXPORT ERROR:",

            str(e)

        )


        # =====================================================
        # CLOSE CONNECTION ON ERROR
        # =====================================================

        try:

            if cursor:

                cursor.close()


        except Exception:

            pass


        try:

            if conn:

                conn.close()


        except Exception:

            pass


        return jsonify({

            "error":

            str(e)

        }), 500




@app.route(
    "/api/finance-ageing/import",
    methods=["POST"]
)
def import_finance_ageing():

    try:

        if "file" not in request.files:

            return jsonify({

                "error":
                "No Excel file selected"

            }), 400


        file = request.files["file"]


        if file.filename == "":

            return jsonify({

                "error":
                "No file selected"

            }), 400


        if not allowed_file(
            file.filename
        ):

            return jsonify({

                "error":
                "Only Excel files (.xlsx, .xls) are allowed"

            }), 400


        filename = secure_filename(
            file.filename
        )


        filepath = os.path.join(

            UPLOAD_FOLDER,

            filename

        )


        file.save(filepath)


        # =====================================
        # READ EXCEL
        # =====================================

        df = pd.read_excel(
            filepath
        )


        # =====================================
        # CLEAN COLUMN NAMES
        # =====================================

        df.columns = (

            df.columns

            .astype(str)

            .str.strip()

            .str.upper()

            .str.replace(
                " ",
                ""
            )

        )


        # =====================================
        # REQUIRED COLUMNS
        # =====================================

        required_columns = [

            "DIVISION",

            "CUSTOMERCODE",

            "CUSTOMERNAME",

            "SALESMANNAME",

            "CREDITLIMIT",

            "ACTUALOVERDUE",

            "DAYS30",

            "DAYS60",

            "DAYS90",

            "DAYS120",

            "DAYS150",

            "DAYS180",

            "DAYS270",

            "DAYS365",

            "DAYSABOVE365",

            "OUTSTANDING",

            "TERMSDAYS"

        ]


        missing_columns = [

            column

            for column in required_columns

            if column not in df.columns

        ]


        if missing_columns:

            return jsonify({

                "error":
                "Missing Excel columns",

                "missing_columns":
                missing_columns

            }), 400


        # =====================================
        # REPLACE NaN WITH ZERO
        # =====================================

        df = df.where(
            pd.notnull(df),
            None
        )


        # =====================================
        # DATABASE CONNECTION
        # =====================================

        conn = get_db_connection()


        cursor = conn.cursor()


        # =====================================
        # CLEAR OLD DATA
        # =====================================

        cursor.execute(

            """
            DELETE FROM finance_ageing
            """

        )


        # =====================================
        # INSERT EXCEL DATA
        # =====================================

        insert_sql = """

        INSERT INTO finance_ageing (

            DIVISION,

            CUSTOMERCODE,

            CUSTOMERNAME,

            SALESMANNAME,

            CREDITLIMIT,

            ACTUALOVERDUE,

            DAYS30,

            DAYS60,

            DAYS90,

            DAYS120,

            DAYS150,

            DAYS180,

            DAYS270,

            DAYS365,

            DAYSABOVE365,

            OUTSTANDING,

            TERMSDAYS

        )

        VALUES (

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s,

            %s

        )

        """


        records = []


        for _, row in df.iterrows():

            records.append(

                (

                    row["DIVISION"],

                    row["CUSTOMERCODE"],

                    row["CUSTOMERNAME"],

                    row["SALESMANNAME"],

                    row["CREDITLIMIT"],

                    row["ACTUALOVERDUE"],

                    row["DAYS30"],

                    row["DAYS60"],

                    row["DAYS90"],

                    row["DAYS120"],

                    row["DAYS150"],

                    row["DAYS180"],

                    row["DAYS270"],

                    row["DAYS365"],

                    row["DAYSABOVE365"],

                    row["OUTSTANDING"],

                    row["TERMSDAYS"]

                )

            )


        cursor.executemany(

            insert_sql,

            records

        )


        conn.commit()


        cursor.close()

        conn.close()


        # =====================================
        # DELETE TEMP FILE
        # =====================================

        if os.path.exists(
            filepath
        ):

            os.remove(
                filepath
            )


        return jsonify({

            "success":
            True,

            "message":
            "Finance Ageing Excel imported successfully",

            "records":
            len(records)

        })


    except Exception as e:


        print(
            "FINANCE EXCEL IMPORT ERROR:",
            str(e)
        )


        return jsonify({

            "error":
            str(e)

        }), 500



# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=False,
    )