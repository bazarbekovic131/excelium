"""Логирование: json-строки в stderr (подхватывает systemd) + отдельный audit-лог."""
import json
import logging
import logging.handlers
from datetime import datetime, timezone
from pathlib import Path


class JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        entry = {
            "ts": datetime.now(timezone.utc).isoformat(timespec="milliseconds"),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }
        if record.exc_info:
            entry["exc"] = self.formatException(record.exc_info)
        extra = getattr(record, "data", None)
        if extra:
            entry.update(extra)
        return json.dumps(entry, ensure_ascii=False)


def setup_logging(var_dir: Path) -> None:
    root = logging.getLogger()
    if not any(getattr(h, "_gateway", False) for h in root.handlers):
        root.setLevel(logging.INFO)
        stream = logging.StreamHandler()
        stream.setFormatter(JsonFormatter())
        stream._gateway = True
        root.addHandler(stream)

    # Файл журнала привязан к var_dir: при повторном вызове с другим
    # каталогом (тесты, смена настроек) прежний хендлер закрывается,
    # иначе события уезжали бы в чужой каталог.
    var_dir.mkdir(parents=True, exist_ok=True)
    target = (var_dir / "audit.log").resolve()
    audit = logging.getLogger("audit")
    audit.setLevel(logging.INFO)
    audit.propagate = False
    for handler in list(audit.handlers):
        if not getattr(handler, "_gateway", False):
            continue
        if Path(getattr(handler, "baseFilename", "")).resolve() == target:
            return
        audit.removeHandler(handler)
        handler.close()
    # Ротация по размеру самим процессом: uvicorn здесь один воркер, а
    # внешний logrotate для этого файла не нужен и не должен настраиваться.
    # Страница журнала читает audit.log и audit.log.1…5 (gateway/auditlog.py).
    fh = logging.handlers.RotatingFileHandler(
        target, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8")
    fh.setFormatter(JsonFormatter())
    fh._gateway = True
    audit.addHandler(fh)


def audit_log(event: str, **data) -> None:
    logging.getLogger("audit").info(event, extra={"data": data})
