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
  created_at TEXT NOT NULL,
  pinned INTEGER NOT NULL DEFAULT 0
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
CREATE TABLE IF NOT EXISTS signer_roles (
  name TEXT PRIMARY KEY,
  title TEXT NOT NULL DEFAULT '',
  holder_uid TEXT NOT NULL DEFAULT '',
  holder_name TEXT NOT NULL DEFAULT '',
  enabled INTEGER NOT NULL DEFAULT 1,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS signer_set_lines (
  set_name TEXT NOT NULL,
  ord INTEGER NOT NULL,
  role TEXT NOT NULL,
  print_company TEXT NOT NULL DEFAULT '',
  mark TEXT NOT NULL DEFAULT '',
  skip_expense_types TEXT NOT NULL DEFAULT '',
  enabled INTEGER NOT NULL DEFAULT 1,
  PRIMARY KEY (set_name, ord)
);
CREATE TABLE IF NOT EXISTS signer_rules (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  company TEXT NOT NULL,
  company_key TEXT NOT NULL,
  object_name TEXT NOT NULL DEFAULT '',
  object_key TEXT NOT NULL DEFAULT '',
  set_name TEXT NOT NULL,
  soglasovano_role TEXT NOT NULL DEFAULT '',
  soglasovano_company TEXT NOT NULL DEFAULT '',
  utverzhdayu_role TEXT NOT NULL DEFAULT '',
  utverzhdayu_company TEXT NOT NULL DEFAULT '',
  enabled INTEGER NOT NULL DEFAULT 1,
  UNIQUE (company_key, object_key)
);
CREATE TABLE IF NOT EXISTS ops_runs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  op TEXT NOT NULL,
  source TEXT NOT NULL DEFAULT 'ui',
  params_json TEXT NOT NULL DEFAULT '{}',
  ok INTEGER NOT NULL,
  exit_code INTEGER NOT NULL,
  duration_ms INTEGER NOT NULL,
  stdout TEXT NOT NULL DEFAULT '',
  stderr TEXT NOT NULL DEFAULT '',
  error TEXT NOT NULL DEFAULT '',
  files_json TEXT NOT NULL DEFAULT '[]',
  started_at TEXT NOT NULL,
  ip TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS ix_ops_runs_op ON ops_runs(op, id);
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

# Колонки, добавленные после первого выпуска: на сервере база уже есть,
# и CREATE TABLE IF NOT EXISTS их не заведёт — дописываем на месте.
_ADDED_COLUMNS = {
    "files": [("pinned", "INTEGER NOT NULL DEFAULT 0")],
    "signer_roles": [("enabled", "INTEGER NOT NULL DEFAULT 1")],
    "signer_set_lines": [("enabled", "INTEGER NOT NULL DEFAULT 1")],
    "signer_rules": [("enabled", "INTEGER NOT NULL DEFAULT 1")],
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