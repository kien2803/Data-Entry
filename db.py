"""SQLite persistence for extraction history — stdlib only, no ORM.

Each upload is stored under data/uploads/ and one row records the outcome,
so results can be re-exported or re-extracted without another LLM call.
"""

import json
import sqlite3
from datetime import datetime
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
DB_PATH = DATA_DIR / "app.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS extractions(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  file_name TEXT,
  stored_path TEXT,
  mode TEXT,
  ok INTEGER,
  duration_ms INTEGER,
  invoice_json TEXT,
  warnings_json TEXT,
  error TEXT,
  created_at TEXT
);
"""


def _conn() -> sqlite3.Connection:
    DATA_DIR.mkdir(exist_ok=True)
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def init_db():
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    with _conn() as c:
        c.execute(SCHEMA)


def _row_to_dict(r: sqlite3.Row) -> dict:
    return dict(r)


def save_extraction(*, file_name, stored_path, mode, ok, duration_ms=None,
                    invoice=None, warnings=None, error=None) -> int:
    with _conn() as c:
        cur = c.execute(
            "INSERT INTO extractions(file_name, stored_path, mode, ok,"
            " duration_ms, invoice_json, warnings_json, error, created_at)"
            " VALUES (?,?,?,?,?,?,?,?,?)",
            (file_name, str(stored_path), mode, int(ok), duration_ms,
             json.dumps(invoice, ensure_ascii=False) if invoice else None,
             json.dumps(warnings, ensure_ascii=False) if warnings else None,
             error, datetime.now().isoformat(timespec="seconds")))
        return cur.lastrowid


def update_extraction(eid: int, *, ok, duration_ms=None,
                      invoice=None, warnings=None, error=None):
    with _conn() as c:
        c.execute(
            "UPDATE extractions SET ok=?, duration_ms=?, invoice_json=?,"
            " warnings_json=?, error=? WHERE id=?",
            (int(ok), duration_ms,
             json.dumps(invoice, ensure_ascii=False) if invoice else None,
             json.dumps(warnings, ensure_ascii=False) if warnings else None,
             error, eid))


def list_extractions(limit: int = 50) -> list[dict]:
    with _conn() as c:
        rows = c.execute(
            "SELECT id, file_name, mode, ok, duration_ms, error, created_at,"
            " (SELECT count(*) FROM json_each(warnings_json)) AS n_warnings"
            " FROM extractions ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
    return [_row_to_dict(r) for r in rows]


def get_extraction(eid: int) -> dict | None:
    with _conn() as c:
        r = c.execute("SELECT * FROM extractions WHERE id=?", (eid,)).fetchone()
    return _row_to_dict(r) if r else None


def delete_extraction(eid: int) -> dict | None:
    row = get_extraction(eid)
    if not row:
        return None
    with _conn() as c:
        c.execute("DELETE FROM extractions WHERE id=?", (eid,))
    return row
