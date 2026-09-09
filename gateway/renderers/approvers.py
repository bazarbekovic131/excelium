"""Матрица подписантов реестра платежей (data/approvers.yaml).

Источник исторический — utils/firmen_und_objekte.py; идентификаторы
расшифровывает лист СПР_ПОДПИСАНТОВ шаблона template.xlsx (VLOOKUP).
Словари строятся один раз при загрузке. Неизвестная пара компания/объект
даёт реестр с пустыми блоками подписей — с warning в лог, а не молча.

Планируемая замена (веха M8): Doc-V присылает согласующих прямо в
payload, и этот модуль перестаёт быть нужным.
"""
import logging
import re
import unicodedata
from pathlib import Path

import yaml

log = logging.getLogger(__name__)

EMPTY = ([0, 0], [0] * 8)  # нули гасятся IFERROR в формулах шаблона

# Doc-V выводит название организации по-разному: то с кавычками-ёлочками,
# то с прямыми, то с довеском вроде «(KZT)» или «(OLD)». Точное сравнение
# на этом ломалось и давало реестр с нулями вместо подписантов, поэтому
# ключи и запрос приводятся к общему виду.
_QUOTES = dict.fromkeys(map(ord, '«»"\'“”„‘’'), None)
_KZ_FOLD = str.maketrans("ұүқғңәөіһҰҮҚҒҢӘӨІҺ", "уукгнаоихуукгнаоих")


def normalize_name(value) -> str:
    text = unicodedata.normalize("NFC", str(value or ""))
    text = re.sub(r"\s*\([^()]*\)\s*$", "", text)   # хвост в скобках
    text = text.translate(_QUOTES).translate(_KZ_FOLD)
    return re.sub(r"[\s\-]+", " ", text).strip().casefold()


def warn_if_stale(yaml_path: Path, source_path: Path) -> bool:
    """Матрицу правят в utils/firmen_und_objekte.py, а шлюз читает YAML.

    Если исходник свежее собранного справочника, реестр уйдёт со старыми
    подписантами и никто этого не заметит — поэтому предупреждаем на старте.
    """
    if not source_path.exists() or not yaml_path.exists():
        return False
    if source_path.stat().st_mtime <= yaml_path.stat().st_mtime:
        return False
    log.warning(
        "матрица подписантов свежее справочника шлюза — пересоберите его: "
        "python3 scripts/approvers_from_py.py > data/approvers.yaml",
        extra={"data": {"source": str(source_path), "yaml": str(yaml_path)}},
    )
    return True


class ApproverMatrix:
    def __init__(self, path: Path):
        raw = yaml.safe_load(path.read_text(encoding="utf-8"))
        lists = raw["approver_lists"]
        self._pairs: dict[tuple[str, str], tuple[list[int], list[int]]] = {}
        for rule in raw["rules"]:
            value = (rule["directors"], lists[rule["approvers"]])
            for obj in rule["objects"]:
                self._pairs[(normalize_name(rule["company"]), normalize_name(obj))] = value
        self._fallbacks = {
            normalize_name(fb["company"]): (fb["directors"], lists[fb["approvers"]])
            for fb in raw.get("company_fallbacks", [])
        }
        excl = raw.get("expense_type_exclusions", {})
        self._excl_types = {t.lower() for t in excl.get("expense_types", [])}
        self._excl_ids = set(excl.get("remove_ids", []))

    def lookup(self, company: str, object_name: str) -> tuple[list[int], list[int]]:
        found = self._pairs.get((normalize_name(company), normalize_name(object_name)))
        if found is None:
            found = self._fallbacks.get(normalize_name(company))
        if found is None:
            log.warning(
                "нет подписантов для пары компания/объект — реестр выйдет без подписей",
                extra={"data": {"company": company, "object": object_name}},
            )
            return EMPTY
        directors, approvers = found
        return list(directors), list(approvers)

    def filter_for_expense_type(self, approvers: list[int], expense_type: str) -> list[int]:
        """Для коммерческих расходов/зарплаты/налогов часть согласующих исключается."""
        if (expense_type or "").strip().lower() in self._excl_types:
            return [a for a in approvers if a not in self._excl_ids]
        return list(approvers)
