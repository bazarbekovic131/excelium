"""Чтение журнала событий (var/audit.log) для интерфейса.

Пишет в журнал logging_setup.audit_log — по JSON-объекту на строку.
Файл ротируется по размеру (audit.log, audit.log.1 … .5), поэтому
чтение идёт по всем частям, свежие первыми. Названия событий и их
окраска живут здесь же, чтобы обзор и страница журнала совпадали.
"""
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterator

ALMATY = timezone(timedelta(hours=5))
BACKUPS = 5   # столько же, сколько у RotatingFileHandler в logging_setup

EVENT_TITLES = {
    "deny_ip": "Отклонён запрос с чужого адреса",
    "deny_token": "Отклонён запрос с неверным токеном",
    "deny_ui": "Отклонён вход в панель",
    "rate_limited": "Запрос отклонён по лимиту частоты",
    "ui_login": "Вход в панель",
    "ui_login_failed": "Неверный админ-токен",
    "job_enqueued": "Задание поставлено в очередь",
    "jobs_acked": "Doc-V подтвердил задания",
    "job_many_attempts": "Задание выдаётся слишком часто",
    "ops_start": "Запущена операция",
    "ops_finish": "Операция завершена",
    "ops_repeat": "Операция повторена",
    "ops_saved": "Сохранена операция",
    "ops_deleted": "Удалена операция",
    "directory_replaced": "Обновлён справочник из Doc-V",
    "typst_template_saved": "Сохранён шаблон Typst",
    "typst_template_restored": "Возвращена версия шаблона",
    "typst_template_deleted": "Удалён шаблон Typst",
    "typst_asset_uploaded": "Загружена картинка",
    "typst_asset_deleted": "Удалена картинка",
    "file_to_assets": "Файл перенесён в картинки",
    "ui_file_uploaded": "Загружен файл",
    "ui_file_renamed": "Файл переименован",
    "ui_file_deleted": "Файл удалён",
    "ui_files_deleted": "Файлы удалены пачкой",
    "ui_file_pinned": "Файл закреплён или откреплён",
    "ui_files_pinned": "Файлы закреплены пачкой",
    "ui_files_zip": "Файлы скачаны архивом",
    "ui_job_ack": "Задание подтверждено вручную",
    "ui_job_enqueued": "Создано тестовое задание",
    "settings_saved": "Изменены настройки",
    "signers_rule_saved": "Сохранено правило подписей",
    "signers_rule_deleted": "Удалено правило подписей",
    "signers_rules_deleted": "Удалены правила подписей",
    "signers_rules_toggled": "Правила подписей включены или отключены",
    "signers_company_applied": "Проставлены подписи по компании",
    "signers_set_saved": "Сохранён набор согласующих",
    "signers_set_deleted": "Удалён набор согласующих",
    "signers_roles_saved": "Сохранён каталог должностей",
    "signers_role_deleted": "Удалена должность",
    "signers_roles_deleted": "Удалены должности",
    "signers_roles_toggled": "Должности включены или отключены",
    "signers_roles_linked": "Должности сопоставлены со Структурой",
    "signers_exported": "Выгружен состав подписантов",
    "signers_imported": "Загружен состав подписантов",
}
META_KEYS = ("ts", "level", "logger", "message")


def tone(message: str, details: dict) -> str:
    """Окраска события: danger — отказы, warn — сбои, off — служебные
    действия в панели, ok — всё остальное."""
    if message.startswith("deny") or message == "rate_limited":
        return "danger"
    if message.endswith("_failed") or details.get("ok") is False:
        return "warn"
    if message.startswith("ui_"):
        return "off"
    return "ok"


def parse(line: str) -> dict | None:
    try:
        raw = json.loads(line)
    except ValueError:
        return None
    if not isinstance(raw, dict):
        return None
    message = str(raw.get("message", ""))
    details = {k: v for k, v in raw.items() if k not in META_KEYS}
    title = EVENT_TITLES.get(message, message)
    return {"ts": str(raw.get("ts", "")), "event": message, "title": title,
            "tone": tone(message, details), "details": details,
            "ip": str(details.get("ip", "")),
            "text": title + (" · " + json.dumps(details, ensure_ascii=False)
                             if details else "")}


def log_files(var_dir: Path) -> list[Path]:
    """audit.log, потом audit.log.1 … — от свежего к старому."""
    base = var_dir / "audit.log"
    files = [base] + [var_dir / f"audit.log.{i}" for i in range(1, BACKUPS + 1)]
    return [f for f in files if f.is_file()]


def iter_events(var_dir: Path) -> Iterator[dict]:
    """Свежие первыми: каждый файл читается целиком и разворачивается."""
    for path in log_files(var_dir):
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
        for line in reversed(lines):
            event = parse(line)
            if event:
                yield event


def query(var_dir: Path, *, search: str = "", tone_filter: str = "", event: str = "",
          days: int = 0, limit: int = 50) -> tuple[list[dict], bool]:
    """-> (события, есть_ли_ещё). days: 1 — с полуночи по Астане, 7 — неделя,
    0 — всё. Файлы хронологические, поэтому при выходе за срок — стоп."""
    needle = search.strip().casefold()
    cutoff = ""
    if days > 0:
        start = datetime.now(ALMATY).replace(hour=0, minute=0, second=0, microsecond=0)
        if days > 1:
            start -= timedelta(days=days - 1)
        cutoff = start.astimezone(timezone.utc).isoformat(timespec="milliseconds")
    out: list[dict] = []
    for item in iter_events(var_dir):
        if cutoff and item["ts"] < cutoff:
            break
        if tone_filter and item["tone"] != tone_filter:
            continue
        if event and item["event"] != event:
            continue
        if needle and needle not in item["text"].casefold():
            continue
        if len(out) >= limit:
            return out, True
        out.append(item)
    return out, False
