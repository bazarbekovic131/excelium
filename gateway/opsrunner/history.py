"""История запусков операций.

Результат запуска раньше жил только на странице ответа и терялся при
обновлении. Теперь каждый запуск — строка в базе: параметры, вывод,
файлы, кто и откуда запустил. Отсюда же операция повторяется с теми же
параметрами. Старые записи чистятся вместе с остальными сметками.
"""
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from ..jobsqueue.db import connect

OUTPUT_LIMIT = 64 * 1024   # stdout и stderr хранятся до 64 КБ каждый
KEEP_DAYS = 30
KEEP_ROWS = 500


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class OpsHistory:
    def __init__(self, db_path: Path):
        self.db_path = db_path

    def add(self, result: dict, params: dict, *, ip: str = "", source: str = "ui") -> int:
        with connect(self.db_path) as conn:
            cur = conn.execute(
                "INSERT INTO ops_runs (op, source, params_json, ok, exit_code, duration_ms,"
                " stdout, stderr, error, files_json, started_at, ip)"
                " VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                (result["op"], source, json.dumps(params, ensure_ascii=False),
                 1 if result["ok"] else 0, int(result["exit_code"]),
                 int(result["duration_ms"]), str(result.get("stdout") or "")[:OUTPUT_LIMIT],
                 str(result.get("stderr") or "")[:OUTPUT_LIMIT], str(result.get("error") or ""),
                 json.dumps(result.get("files") or [], ensure_ascii=False), _now(), ip))
            return cur.lastrowid

    def get(self, run_id: int) -> dict | None:
        with connect(self.db_path) as conn:
            row = conn.execute("SELECT * FROM ops_runs WHERE id = ?", (run_id,)).fetchone()
        return self._row(row, with_output=True) if row else None

    def recent(self, op: str | None = None, limit: int = 20) -> list[dict]:
        """Свежие первыми, без тел вывода — для списков."""
        sql = ("SELECT id, op, source, params_json, ok, exit_code, duration_ms, error,"
               " files_json, started_at, ip FROM ops_runs")
        args: tuple = ()
        if op:
            sql += " WHERE op = ?"
            args = (op,)
        sql += " ORDER BY id DESC LIMIT ?"
        with connect(self.db_path) as conn:
            rows = conn.execute(sql, (*args, limit)).fetchall()
        return [self._row(r, with_output=False) for r in rows]

    def last_by_op(self) -> dict[str, dict]:
        """Последний запуск каждой операции — для строки под карточкой."""
        out: dict[str, dict] = {}
        for run in self.recent(limit=KEEP_ROWS):
            out.setdefault(run["op"], run)
        return out

    def sweep(self) -> int:
        cutoff = (datetime.now(timezone.utc) - timedelta(days=KEEP_DAYS)).isoformat(
            timespec="seconds")
        with connect(self.db_path) as conn:
            removed = conn.execute("DELETE FROM ops_runs WHERE started_at < ?",
                                   (cutoff,)).rowcount
            removed += conn.execute(
                "DELETE FROM ops_runs WHERE id NOT IN"
                " (SELECT id FROM ops_runs ORDER BY id DESC LIMIT ?)", (KEEP_ROWS,)).rowcount
        return removed

    @staticmethod
    def _row(row, *, with_output: bool) -> dict:
        out = {"id": row["id"], "op": row["op"], "source": row["source"],
               "params": json.loads(row["params_json"] or "{}"),
               "ok": bool(row["ok"]), "exit_code": row["exit_code"],
               "duration_ms": row["duration_ms"], "error": row["error"],
               "files": json.loads(row["files_json"] or "[]"),
               "started_at": row["started_at"], "ip": row["ip"]}
        if with_output:
            out["stdout"] = row["stdout"]
            out["stderr"] = row["stderr"]
        return out
