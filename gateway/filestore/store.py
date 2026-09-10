"""Выдача файлов по непере­бираемым токенам с TTL.

Файл лежит в var/files/<token><suffix>; человекочитаемое имя в таблице files той же базе шлюза.
Доступно по адресу address/ui/files (Файлы)
"""
import logging
import re
import secrets
from datetime import datetime, timedelta, timezone
from pathlib import Path

from ..config import Settings
from ..jobsqueue.db import connect

log = logging.getLogger(__name__)

TOKEN_RE = re.compile(r"^[A-Za-z0-9_-]{20,50}$")
_SUFFIX_RE = re.compile(r"^\.[A-Za-z0-9.]{1,10}$")

# returns the current time with timezone awareness
def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class FileStore:
    """
    Initialized with settings (defined in config)
    gateway/config.py

    Creates a file directory if not present.
    """
    def __init__(self, settings: Settings):
        self.settings = settings
        self.files_dir = settings.files_dir
        self.files_dir.mkdir(parents=True, exist_ok=True)

    def save_bytes(self, data: bytes, suffix: str, orig_name: str) -> str:
        """
        Generates a token which acts together with suffix as a file name. Saves the file and returns
        the token. 

        Called in save_file method of FileStore class.
        """
        if not _SUFFIX_RE.fullmatch(suffix):
            raise ValueError(f"bad suffix: {suffix!r}") # what-s bad suffix?
        token = secrets.token_urlsafe(24) # I feel like this is useless kinda, simple cipher.
        (self.files_dir / f"{token}{suffix}").write_bytes(data) # write file to the given directory
        with connect(self.settings.db_path) as conn:
            conn.execute(
                "INSERT INTO files (token, orig_name, suffix, created_at) VALUES (?,?,?,?)",
                (token, orig_name, suffix, _now()),
            )
        return token

    def save_file(self, path: Path, orig_name: str | None = None) -> str:
        return self.save_bytes(path.read_bytes(), path.suffix, orig_name or path.name)

    def resolve(self, token: str) -> tuple[Path, str] | None:
        """token -> (file path on a disk, имя для скачивания) or None."""
        if not TOKEN_RE.fullmatch(token):
            return None
        with connect(self.settings.db_path) as conn:
            row = conn.execute(
                "SELECT orig_name, suffix FROM files WHERE token = ?", (token,)
            ).fetchone()
        if row is None:
            return None
        path = self.files_dir / f"{token}{row['suffix']}"
        if not path.is_file():
            return None
        return path, row["orig_name"]

    def download_url(self, token: str) -> str:
        return f"{self.settings.base_url}/files/{token}"

    def list_files(self) -> list[dict]:
        with connect(self.settings.db_path) as conn:
            rows = conn.execute(
                "SELECT token, orig_name, suffix, created_at, pinned FROM files"
                " ORDER BY created_at DESC").fetchall()
        out = []
        for row in rows:
            path = self.files_dir / f"{row['token']}{row['suffix']}"
            out.append({"token": row["token"], "orig_name": row["orig_name"],
                        "suffix": row["suffix"], "created_at": row["created_at"],
                        "pinned": bool(row["pinned"]),
                        "size": path.stat().st_size if path.is_file() else 0})
        return out

    def set_pinned(self, tokens: list[str], pinned: bool) -> int:
        """Закреплённый файл переживает срок хранения: бланки, эталоны,
        резервные копии не должны исчезать через трое суток."""
        tokens = [t for t in tokens if TOKEN_RE.fullmatch(t)]
        if not tokens:
            return 0
        marks = ",".join("?" * len(tokens))
        with connect(self.settings.db_path) as conn:
            cur = conn.execute(f"UPDATE files SET pinned = ? WHERE token IN ({marks})",
                               (1 if pinned else 0, *tokens))
        return cur.rowcount

    def delete_many(self, tokens: list[str]) -> int:
        return sum(1 for t in tokens if self.delete(t))

    def rename(self, token: str, new_name: str) -> bool:
        """Меняет отображаемое имя (имя на диске остаётся токеном)."""
        if not TOKEN_RE.fullmatch(token) or not new_name.strip():
            return False
        with connect(self.settings.db_path) as conn:
            cur = conn.execute("UPDATE files SET orig_name = ? WHERE token = ?",
                               (new_name.strip(), token))
        return cur.rowcount > 0

    def delete(self, token: str) -> bool:
        if not TOKEN_RE.fullmatch(token):
            return False
        with connect(self.settings.db_path) as conn:
            row = conn.execute("SELECT suffix FROM files WHERE token = ?", (token,)).fetchone()
            if row is None:
                return False
            conn.execute("DELETE FROM files WHERE token = ?", (token,))
        (self.files_dir / f"{token}{row['suffix']}").unlink(missing_ok=True)
        return True

    def sweep(self) -> int:
        """Удаляет файлы старше TTL (кроме закреплённых) и осиротевшие файлы без записи."""
        cutoff = (
            datetime.now(timezone.utc) - timedelta(hours=self.settings.file_ttl_hours)
        ).isoformat(timespec="seconds")
        removed = 0
        with connect(self.settings.db_path) as conn:
            rows = conn.execute(
                "SELECT token, suffix FROM files WHERE created_at < ? AND pinned = 0",
                (cutoff,)).fetchall()
            for row in rows:
                (self.files_dir / f"{row['token']}{row['suffix']}").unlink(missing_ok=True)
                removed += 1
            conn.execute("DELETE FROM files WHERE created_at < ? AND pinned = 0", (cutoff,))
            known = {r["token"] + r["suffix"] for r in conn.execute(
                "SELECT token, suffix FROM files").fetchall()}
        for f in self.files_dir.iterdir():
            if f.is_file() and f.name not in known:
                f.unlink(missing_ok=True)
                removed += 1
        if removed:
            log.info("filestore sweep", extra={"data": {"removed": removed}})
        return removed
