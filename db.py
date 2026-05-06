"""Database connection + schema helpers.

Secrets are read from .env (kept out of git by .gitignore).
Never log DB_PASSWORD or the full DB_URL.
"""

import os
import secrets
from urllib.parse import quote_plus
from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

load_dotenv()

_REQUIRED = ("DB_HOST", "DB_PORT", "DB_USER", "DB_PASSWORD", "DB_NAME")


def _require(key: str) -> str:
    val = os.environ.get(key)
    if not val:
        raise RuntimeError(f"Missing required env var: {key} (check your .env)")
    return val


def _db_url(database: str | None = None) -> str:
    """Build a SQLAlchemy URL. Password is URL-encoded so special chars are safe."""
    host = _require("DB_HOST")
    port = _require("DB_PORT")
    user = _require("DB_USER")
    password = quote_plus(_require("DB_PASSWORD"))
    name = database if database is not None else _require("DB_NAME")
    return f"mysql+pymysql://{user}:{password}@{host}:{port}/{name}?charset=utf8mb4"


_engine: Engine | None = None


def get_engine() -> Engine:
    """Return the shared SQLAlchemy engine, creating it on first use."""
    global _engine
    if _engine is None:
        _engine = create_engine(_db_url(), pool_pre_ping=True, pool_recycle=3600)
    return _engine


def ensure_database() -> None:
    """Create the database if it doesn't exist. Safe to call repeatedly."""
    db_name = _require("DB_NAME")
    # Connect without selecting a database, then create it
    server_engine = create_engine(_db_url(database=""), pool_pre_ping=True)
    with server_engine.connect() as conn:
        conn.execute(text(
            f"CREATE DATABASE IF NOT EXISTS `{db_name}` "
            f"CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        ))
        conn.commit()
    server_engine.dispose()


# Drop the legacy invoice-keyed tables before creating the project-keyed ones.
# Order matters: children (FKs) before parent.
LEGACY_DROPS = [
    "DROP TABLE IF EXISTS invoice_photos",
    "DROP TABLE IF EXISTS invoice_phones",
    "DROP TABLE IF EXISTS invoices",
]


SCHEMA = [
    # ---------- projects ----------
    """
    CREATE TABLE IF NOT EXISTS projects (
        project_number   VARCHAR(64)   NOT NULL,
        vendor_key       VARCHAR(32)   NOT NULL,
        vendor_name      VARCHAR(128)  NOT NULL,
        share_token      VARCHAR(64)   NOT NULL,
        return_count     INT           NOT NULL DEFAULT 0,
        notes            TEXT,
        created_at       DATETIME      NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at       DATETIME      NOT NULL DEFAULT CURRENT_TIMESTAMP
                                       ON UPDATE CURRENT_TIMESTAMP,
        PRIMARY KEY (project_number),
        UNIQUE KEY uq_share_token (share_token),
        KEY idx_vendor (vendor_key)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
    """,

    # ---------- phones in each project ----------
    """
    CREATE TABLE IF NOT EXISTS project_phones (
        id                BIGINT AUTO_INCREMENT PRIMARY KEY,
        project_number    VARCHAR(64)   NOT NULL,
        invoice_number    VARCHAR(64),
        wid               VARCHAR(64),
        imei              VARCHAR(32),
        brand             VARCHAR(64),
        model             VARCHAR(128),
        carrier           VARCHAR(64),
        carrier_gsx       VARCHAR(64),
        grade             VARCHAR(32),
        qc_error_code     VARCHAR(128),
        condition_text    VARCHAR(255),
        cost              DECIMAL(10,2),
        reasons           TEXT,
        not_sure_reasons  TEXT,
        detail_lock       VARCHAR(32),
        scanned           TINYINT(1)    NOT NULL DEFAULT 0,
        scanned_at        DATETIME      NULL,
        created_at        DATETIME      NOT NULL DEFAULT CURRENT_TIMESTAMP,
        KEY idx_project (project_number),
        KEY idx_invoice (invoice_number),
        KEY idx_imei (imei),
        KEY idx_wid (wid),
        CONSTRAINT fk_phone_project FOREIGN KEY (project_number)
            REFERENCES projects (project_number) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
    """,

    # ---------- photo links per project ----------
    """
    CREATE TABLE IF NOT EXISTS project_photos (
        id               BIGINT AUTO_INCREMENT PRIMARY KEY,
        project_number   VARCHAR(64)   NOT NULL,
        photo_url        TEXT          NOT NULL,
        label            VARCHAR(255),
        created_at       DATETIME      NOT NULL DEFAULT CURRENT_TIMESTAMP,
        KEY idx_project (project_number),
        CONSTRAINT fk_photo_project FOREIGN KEY (project_number)
            REFERENCES projects (project_number) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
    """,
]


def init_schema() -> None:
    """Drop legacy invoice tables and create the project tables. Idempotent."""
    engine = get_engine()
    with engine.begin() as conn:
        for stmt in LEGACY_DROPS:
            conn.execute(text(stmt))
        for stmt in SCHEMA:
            conn.execute(text(stmt))


def health_check() -> bool:
    """Return True if we can reach the database."""
    try:
        with get_engine().connect() as conn:
            conn.execute(text("SELECT 1"))
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Repository helpers
# ---------------------------------------------------------------------------

def _gen_share_token() -> str:
    return secrets.token_urlsafe(12)


def _clean(val):
    """Normalize pandas/numpy NaN to None. Leaves everything else as-is."""
    if val is None:
        return None
    try:
        import pandas as pd
        if pd.isna(val):
            return None
    except Exception:
        pass
    return val


def upsert_project(project_number: str, vendor_key: str, vendor_name: str,
                   return_count: int) -> str:
    """Create or update a project row. Returns its share_token (stable across updates)."""
    with get_engine().begin() as conn:
        existing = conn.execute(
            text("SELECT share_token FROM projects WHERE project_number = :pn"),
            {"pn": project_number},
        ).fetchone()
        if existing:
            conn.execute(text("""
                UPDATE projects
                   SET vendor_key=:vk, vendor_name=:vn, return_count=:rc
                 WHERE project_number=:pn
            """), {"vk": vendor_key, "vn": vendor_name,
                   "rc": return_count, "pn": project_number})
            return existing[0]
        token = _gen_share_token()
        conn.execute(text("""
            INSERT INTO projects
                (project_number, vendor_key, vendor_name, share_token, return_count)
            VALUES (:pn, :vk, :vn, :tok, :rc)
        """), {"pn": project_number, "vk": vendor_key, "vn": vendor_name,
               "tok": token, "rc": return_count})
        return token


def replace_phones(project_number: str, rows: list[dict]) -> None:
    """Replace the phone list for a given project with a fresh set of rows."""
    with get_engine().begin() as conn:
        conn.execute(
            text("DELETE FROM project_phones WHERE project_number = :pn"),
            {"pn": project_number},
        )
        if not rows:
            return
        payload = [{
            "pn": project_number,
            "inv": _clean(r.get("Invoice")),
            "wid": _clean(r.get("WID")),
            "imei": _clean(r.get("IMEI")),
            "brand": _clean(r.get("Brand")),
            "model": _clean(r.get("Model")),
            "carrier": _clean(r.get("Carrier")),
            "carrier_gsx": _clean(r.get("Carrier GSX")),
            "grade": _clean(r.get("Vendor Condition")),
            "qc": _clean(r.get("QC Error Code")),
            "cond": _clean(r.get("Condition")),
            "cost": _clean(r.get("Cost")),
            "reasons": _clean(r.get("Reason(s)")),
            "ns": _clean(r.get("Not Sure / Need to Test")),
            "detail": _clean(r.get("Detail Lock Status")),
        } for r in rows]
        conn.execute(text("""
            INSERT INTO project_phones
                (project_number, invoice_number, wid, imei, brand, model,
                 carrier, carrier_gsx, grade, qc_error_code, condition_text,
                 cost, reasons, not_sure_reasons, detail_lock)
            VALUES (:pn, :inv, :wid, :imei, :brand, :model,
                    :carrier, :carrier_gsx, :grade, :qc, :cond,
                    :cost, :reasons, :ns, :detail)
        """), payload)


def get_project_by_token(token: str) -> dict | None:
    with get_engine().connect() as conn:
        row = conn.execute(
            text("SELECT * FROM projects WHERE share_token = :tok"),
            {"tok": token},
        ).mappings().fetchone()
        return dict(row) if row else None


def get_project_by_number(project_number: str) -> dict | None:
    with get_engine().connect() as conn:
        row = conn.execute(
            text("SELECT * FROM projects WHERE project_number = :pn"),
            {"pn": project_number},
        ).mappings().fetchone()
        return dict(row) if row else None


def get_phones_for_project(project_number: str) -> list[dict]:
    with get_engine().connect() as conn:
        rows = conn.execute(text("""
            SELECT id, invoice_number, wid, imei, brand, model, carrier, carrier_gsx,
                   grade, qc_error_code, condition_text, cost, reasons,
                   not_sure_reasons, detail_lock, scanned, scanned_at
              FROM project_phones
             WHERE project_number = :pn
          ORDER BY id
        """), {"pn": project_number}).mappings().all()
        return [dict(r) for r in rows]


def list_photos(project_number: str) -> list[dict]:
    with get_engine().connect() as conn:
        rows = conn.execute(text("""
            SELECT id, photo_url, label, created_at
              FROM project_photos
             WHERE project_number = :pn
          ORDER BY id
        """), {"pn": project_number}).mappings().all()
        return [dict(r) for r in rows]


def add_photo(project_number: str, photo_url: str, label: str = "") -> int:
    with get_engine().begin() as conn:
        result = conn.execute(text("""
            INSERT INTO project_photos (project_number, photo_url, label)
            VALUES (:pn, :url, :label)
        """), {"pn": project_number, "url": photo_url, "label": label or None})
        return result.lastrowid


def delete_photo(photo_id: int, project_number: str) -> bool:
    """Delete a photo, but only if it belongs to the given project. Returns True on delete."""
    with get_engine().begin() as conn:
        result = conn.execute(text("""
            DELETE FROM project_photos
             WHERE id = :id AND project_number = :pn
        """), {"id": photo_id, "pn": project_number})
        return result.rowcount > 0
