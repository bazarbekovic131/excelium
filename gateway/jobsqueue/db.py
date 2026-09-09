"""SQLite база шлюза: 
Функции: 
  1. Ведение очереди заданий.
  2. Ведение реестра выданных файлов.
"""
import sqlite3
from pathlib import Path

_SCHEMA = """
CREATE TABLE IF NOT EXISTS jobs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  consumer TEXT NOT NULL DEFAULT 'docv',
  producer TEXT NOT NULL,
  idempotency_key TEXT,
  type TEXT NOT NULL,
  payload TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending','leased','acked')),
  attempts INTEGER NOT NULL DEFAULT 0,
  leased_until TEXT,
  created_at TEXT NOT NULL,
  acked_at TEXT,
  UNIQUE (producer, idempotency_key)
);
CREATE INDEX IF NOT EXISTS ix_jobs_lease ON jobs(consumer, status, leased_until);
CREATE TABLE IF NOT EXISTS files (
  token TEXT PRIMARY KEY,
  orig_name TEXT NOT NULL,
  suffix TEXT NOT NULL,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS typst_templates (
  name TEXT PRIMARY KEY,
  source TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS typst_template_history (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  source TEXT NOT NULL,
  saved_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_typst_history ON typst_template_history(name, id);
CREATE TABLE IF NOT EXISTS typst_assets (
  name TEXT PRIMARY KEY,
  data BLOB NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS directories (
  name TEXT NOT NULL,
  uid TEXT NOT NULL,
  data TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  PRIMARY KEY (name, uid)
);
CREATE TABLE IF NOT EXISTS signer_people (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  fio TEXT NOT NULL,
  fio_key TEXT NOT NULL UNIQUE,
  position TEXT NOT NULL DEFAULT '',
  docv_uid TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS signer_positions (
  name TEXT PRIMARY KEY,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS signer_sets (
  name TEXT NOT NULL,
  ord INTEGER NOT NULL,
  person_id INTEGER REFERENCES signer_people(id) ON DELETE CASCADE,
  position TEXT NOT NULL DEFAULT '',
  position_ref TEXT NOT NULL DEFAULT '',
  dept_ref TEXT NOT NULL DEFAULT '',
  print_company TEXT NOT NULL DEFAULT '',
  mark TEXT NOT NULL DEFAULT '',
  skip_expense_types TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (name, ord)
);
CREATE TABLE IF NOT EXISTS signer_bindings (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  company TEXT NOT NULL,
  company_key TEXT NOT NULL,
  object_name TEXT NOT NULL DEFAULT '',
  object_key TEXT NOT NULL DEFAULT '',
  set_name TEXT NOT NULL,
  soglasovano_id INTEGER REFERENCES signer_people(id),
  soglasovano_position TEXT NOT NULL DEFAULT '',
  soglasovano_company TEXT NOT NULL DEFAULT '',
  soglasovano_ref TEXT NOT NULL DEFAULT '',
  soglasovano_dept TEXT NOT NULL DEFAULT '',
  utverzhdayu_id INTEGER REFERENCES signer_people(id),
  utverzhdayu_position TEXT NOT NULL DEFAULT '',
  utverzhdayu_company TEXT NOT NULL DEFAULT '',
  utverzhdayu_ref TEXT NOT NULL DEFAULT '',
  utverzhdayu_dept TEXT NOT NULL DEFAULT '',
  UNIQUE (company_key, object_key)
);
CREATE TABLE IF NOT EXISTS heartbeat (
  kind TEXT PRIMARY KEY,
  seen_at TEXT NOT NULL
);
"""

# default: PRAGMA journal_mode=WAL
# - 
# default: PRAGMA foreign_keys=ON
#
def connect(path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(path, timeout=5)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

# Колонки, добавленные после первого выпуска: база на сервере уже
# создана, поэтому CREATE TABLE их не заведёт — дописываем на месте.
_ADDED_COLUMNS = {
    "signer_sets": [("position_ref", "TEXT NOT NULL DEFAULT \'\'"),
                    ("dept_ref", "TEXT NOT NULL DEFAULT \'\'")],
    "signer_bindings": [("soglasovano_ref", "TEXT NOT NULL DEFAULT \'\'"),
                        ("soglasovano_dept", "TEXT NOT NULL DEFAULT \'\'"),
                        ("utverzhdayu_ref", "TEXT NOT NULL DEFAULT \'\'"),
                        ("utverzhdayu_dept", "TEXT NOT NULL DEFAULT \'\'")],
}


def init_db(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with connect(path) as conn:
        conn.executescript(_SCHEMA)
        for table, columns in _ADDED_COLUMNS.items():
            have = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
            for name, decl in columns:
                if name not in have:
                    conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {decl}")
    path.chmod(0o600)