"""Кто подписывает реестр: собственный справочник шлюза.

Раньше состав подписей жил в двух местах сразу — матрица на сервере
задавала номера строк, а ФИО и должности лежали на скрытом листе
СПР_ПОДПИСАНТОВ внутри шаблона Excel. Реестр из-за этого менялся задним
числом: правка листа переписывала подписи в уже выпущенных файлах.

Теперь состав знает шлюз, а в Excel уходят готовые строки. Doc-V ничего
про подписи не присылает — наоборот, может спросить: GET /signers.

Модель:
- люди — ФИО и должность по умолчанию, связываются со справочником
  Структуры из Doc-V по фамилии с инициалами, и тогда ФИО обновляется
  само;
- наборы — упорядоченные списки согласующих под таблицей;
- привязки — какой набор и какие двое руководителей идут у компании
  (при необходимости у отдельного объекта компании).

Должность и компания хранятся в привязке, а не у человека: один и тот
же человек — генеральный директор двух десятков компаний, и в Структуре
Doc-V у него одна должность, а печатать надо ту, под которой он
подписывает этот реестр.
"""
import logging
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

from .jobsqueue.db import connect

log = logging.getLogger(__name__)

ROLE_SOGLASOVANO = "soglasovano"
ROLE_UTVERZHDAYU = "utverzhdayu"
AGREED_MARK_ID = 3  # перед этим подписантом печаталась строка «СОГЛАСОВАНО»

# Doc-V выводит название организации по-разному: то с кавычками-ёлочками,
# то с прямыми, то с довеском вроде «(KZT)», то с русской К вместо
# казахской Қ. Точное сравнение на этом ломалось и давало реестр без
# подписей, поэтому ключи и запрос приводятся к общему виду.
_QUOTES = dict.fromkeys(map(ord, '«»"\'“”„‘’'), None)
_KZ_FOLD = str.maketrans("ұүқғңәөіһҰҮҚҒҢӘӨІҺ", "уукгнаоихуукгнаоих")


def normalize_name(value) -> str:
    text = unicodedata.normalize("NFC", str(value or ""))
    text = re.sub(r"\s*\([^()]*\)\s*$", "", text)   # хвост в скобках
    text = text.translate(_QUOTES).translate(_KZ_FOLD)
    return re.sub(r"[\s\-]+", " ", text).strip().casefold()


def fio_key(value) -> str:
    """Фамилия и инициалы: «Аманов Б.Ш.» и «Аманов Бауыржан Шарипович»
    должны сойтись — в шаблоне писали инициалами, Doc-V шлёт полностью."""
    text = re.sub(r"[.,]", " ", unicodedata.normalize("NFC", str(value or "")))
    parts = [p for p in re.split(r"\s+", text.translate(_KZ_FOLD)) if p]
    if not parts:
        return ""
    initials = "".join(p[0] for p in parts[1:3])
    return f"{parts[0]} {initials}".strip().casefold()


def _person(fio: str, position: str, company: str, mark: str = "") -> dict:
    out = {"fio": fio, "position": position, "company": company}
    if mark:
        out["mark"] = mark
    return out


class SignerStore:
    def __init__(self, db_path: Path):
        self.db_path = db_path

    # --- чтение ---------------------------------------------------------

    def people(self) -> list[dict]:
        with connect(self.db_path) as conn:
            return [dict(r) for r in conn.execute(
                "SELECT * FROM signer_people ORDER BY fio").fetchall()]

    def sets(self) -> dict[str, list[dict]]:
        out: dict[str, list[dict]] = {}
        with connect(self.db_path) as conn:
            for r in conn.execute(
                    "SELECT s.*, p.fio, p.position AS default_position"
                    " FROM signer_sets s JOIN signer_people p ON p.id = s.person_id"
                    " ORDER BY s.name, s.ord").fetchall():
                out.setdefault(r["name"], []).append(dict(r))
        return out

    def bindings(self) -> list[dict]:
        with connect(self.db_path) as conn:
            return [dict(r) for r in conn.execute(
                "SELECT * FROM signer_bindings ORDER BY company, object_name").fetchall()]

    def stats(self) -> dict[str, int]:
        with connect(self.db_path) as conn:
            return {
                "people": conn.execute("SELECT COUNT(*) c FROM signer_people").fetchone()["c"],
                "sets": conn.execute(
                    "SELECT COUNT(DISTINCT name) c FROM signer_sets").fetchone()["c"],
                "bindings": conn.execute(
                    "SELECT COUNT(*) c FROM signer_bindings").fetchone()["c"],
                "linked": conn.execute(
                    "SELECT COUNT(*) c FROM signer_people WHERE docv_uid <> ''").fetchone()["c"],
            }

    # --- подбор ---------------------------------------------------------

    def resolve(self, company: str, object_name: str = "",
                expense_type: str = "") -> dict:
        """-> {"soglasovano": …|None, "utverzhdayu": …|None,
                "coordinators": [...], "source": …}."""
        company_key = normalize_name(company)
        object_key = normalize_name(object_name)
        with connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT * FROM signer_bindings WHERE company_key = ? AND object_key = ?",
                (company_key, object_key)).fetchone()
            source = "объект"
            if row is None:
                row = conn.execute(
                    "SELECT * FROM signer_bindings WHERE company_key = ? AND object_key = ''",
                    (company_key,)).fetchone()
                source = "компания"
            if row is None:
                log.warning("нет подписантов: реестр выйдет без подписей",
                            extra={"data": {"company": company, "object": object_name}})
                return {"soglasovano": None, "utverzhdayu": None,
                        "coordinators": [], "source": "нет привязки"}
            names = self._names(conn)
            entries = conn.execute(
                "SELECT * FROM signer_sets WHERE name = ? ORDER BY ord",
                (row["set_name"],)).fetchall()

        expense = str(expense_type or "").strip().casefold()
        coordinators = []
        for entry in entries:
            skip = [t.strip().casefold() for t in (entry["skip_expense_types"] or "").split(",")]
            if expense and expense in [t for t in skip if t]:
                continue
            fio, default_position = names.get(entry["person_id"], ("", ""))
            coordinators.append(_person(fio, entry["position"] or default_position,
                                        entry["print_company"], entry["mark"]))

        def top(role: str) -> dict | None:
            person_id = row[f"{role}_id"]
            if not person_id:
                return None
            fio, default_position = names.get(person_id, ("", ""))
            if not fio:
                return None
            return _person(fio, row[f"{role}_position"] or default_position,
                           row[f"{role}_company"] or row["company"])

        return {"soglasovano": top(ROLE_SOGLASOVANO),
                "utverzhdayu": top(ROLE_UTVERZHDAYU),
                "coordinators": coordinators, "source": source}

    def _names(self, conn) -> dict[int, tuple[str, str]]:
        return {r["id"]: (r["fio"], r["position"])
                for r in conn.execute("SELECT id, fio, position FROM signer_people")}

    # --- связь со Структурой Doc-V ---------------------------------------

    def link_directory(self, directories: dict[str, dict[str, dict]]) -> int:
        """Сопоставляет людей с выгрузкой Структуры по фамилии и инициалам
        и подтягивает оттуда ФИО: состав Doc-V меняется чаще справочника."""
        by_key: dict[str, tuple[str, str]] = {}
        for items in directories.values():
            for uid, data in items.items():
                name = data.get("name") or data.get("display_name") or ""
                key = fio_key(name)
                if key and key not in by_key:
                    by_key[key] = (uid, str(name).strip())
        if not by_key:
            return 0
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        linked = 0
        with connect(self.db_path) as conn:
            for row in conn.execute("SELECT id, fio, fio_key FROM signer_people").fetchall():
                found = by_key.get(row["fio_key"])
                if not found or found[1] == row["fio"]:
                    continue
                conn.execute(
                    "UPDATE signer_people SET docv_uid = ?, fio = ?, updated_at = ?"
                    " WHERE id = ?", (found[0], found[1], now, row["id"]))
                linked += 1
        if linked:
            log.info("подписанты связаны со Структурой",
                     extra={"data": {"count": linked}})
        return linked

    # --- правка из /ui ----------------------------------------------------

    def save_person(self, person_id: int | None, fio: str, position: str) -> int:
        fio = fio.strip()
        if not fio:
            raise ValueError("ФИО обязательно")
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        with connect(self.db_path) as conn:
            if person_id:
                conn.execute(
                    "UPDATE signer_people SET fio = ?, fio_key = ?, position = ?,"
                    " updated_at = ? WHERE id = ?",
                    (fio, fio_key(fio), position.strip(), now, person_id))
                return person_id
            cur = conn.execute(
                "INSERT INTO signer_people (fio, fio_key, position, docv_uid, updated_at)"
                " VALUES (?,?,?,'',?) ON CONFLICT(fio_key) DO UPDATE SET fio = excluded.fio,"
                " position = excluded.position, updated_at = excluded.updated_at",
                (fio, fio_key(fio), position.strip(), now))
            if cur.lastrowid:
                return cur.lastrowid
            return conn.execute("SELECT id FROM signer_people WHERE fio_key = ?",
                                (fio_key(fio),)).fetchone()["id"]

    def save_binding(self, *, company: str, object_name: str, set_name: str,
                     soglasovano: dict | None, utverzhdayu: dict | None,
                     binding_id: int | None = None) -> int:
        """soglasovano/utverzhdayu: {"person_id", "position", "company"}."""
        company = company.strip()
        if not company:
            raise ValueError("компания обязательна")
        object_name = object_name.strip()
        left = soglasovano or {}
        right = utverzhdayu or {}
        values = (company, normalize_name(company), object_name, normalize_name(object_name),
                  set_name.strip(),
                  left.get("person_id") or None, (left.get("position") or "").strip(),
                  (left.get("company") or "").strip(),
                  right.get("person_id") or None, (right.get("position") or "").strip(),
                  (right.get("company") or "").strip())
        with connect(self.db_path) as conn:
            if binding_id:
                conn.execute(
                    "UPDATE signer_bindings SET company=?, company_key=?, object_name=?,"
                    " object_key=?, set_name=?, soglasovano_id=?, soglasovano_position=?,"
                    " soglasovano_company=?, utverzhdayu_id=?, utverzhdayu_position=?,"
                    " utverzhdayu_company=? WHERE id=?", (*values, binding_id))
                return binding_id
            cur = conn.execute(
                "INSERT INTO signer_bindings (company, company_key, object_name, object_key,"
                " set_name, soglasovano_id, soglasovano_position, soglasovano_company,"
                " utverzhdayu_id, utverzhdayu_position, utverzhdayu_company)"
                " VALUES (?,?,?,?,?,?,?,?,?,?,?)"
                " ON CONFLICT(company_key, object_key) DO UPDATE SET"
                " company=excluded.company, object_name=excluded.object_name,"
                " set_name=excluded.set_name, soglasovano_id=excluded.soglasovano_id,"
                " soglasovano_position=excluded.soglasovano_position,"
                " soglasovano_company=excluded.soglasovano_company,"
                " utverzhdayu_id=excluded.utverzhdayu_id,"
                " utverzhdayu_position=excluded.utverzhdayu_position,"
                " utverzhdayu_company=excluded.utverzhdayu_company", values)
            return cur.lastrowid or 0

    def delete_binding(self, binding_id: int) -> None:
        with connect(self.db_path) as conn:
            conn.execute("DELETE FROM signer_bindings WHERE id = ?", (binding_id,))

    def save_set(self, name: str, entries: list[dict]) -> int:
        """entries: [{"person_id", "position", "company", "mark",
        "skip_expense_types"}] — порядок списка и есть порядок подписей."""
        name = name.strip()
        if not name:
            raise ValueError("имя набора обязательно")
        with connect(self.db_path) as conn:
            conn.execute("DELETE FROM signer_sets WHERE name = ?", (name,))
            conn.executemany(
                "INSERT INTO signer_sets (name, ord, person_id, position, print_company,"
                " mark, skip_expense_types) VALUES (?,?,?,?,?,?,?)",
                [(name, i, e["person_id"], (e.get("position") or "").strip(),
                  (e.get("company") or "").strip(), (e.get("mark") or "").strip(),
                  (e.get("skip_expense_types") or "").strip())
                 for i, e in enumerate(entries) if e.get("person_id")])
        return len(entries)

    # --- первичное наполнение --------------------------------------------

    def is_empty(self) -> bool:
        with connect(self.db_path) as conn:
            return conn.execute(
                "SELECT COUNT(*) c FROM signer_bindings").fetchone()["c"] == 0

    def seed(self, matrix_path: Path, template_path: Path) -> int:
        """Разовый перенос состава из старых источников: структура — из
        data/signers_seed.yaml, ФИО и должности — с листа СПР_ПОДПИСАНТОВ
        шаблона. Дальше состав правится в /ui, а эти файлы не нужны."""
        import openpyxl
        import yaml

        if not matrix_path.exists() or not template_path.exists():
            return 0
        raw = yaml.safe_load(matrix_path.read_text(encoding="utf-8"))
        sheet = openpyxl.load_workbook(template_path, read_only=True)["СПР_ПОДПИСАНТОВ"]
        spr: dict[int, tuple[str, str, str]] = {}
        for row in sheet.iter_rows(min_col=2, max_col=10, values_only=True):
            number, fio, company, position = row[0], row[4], row[6], row[8]
            if isinstance(number, int) and fio:
                spr[number] = (str(fio).strip(), str(position or "").strip(),
                               str(company or "").strip())

        excl = raw.get("expense_type_exclusions", {})
        skip_types = ",".join(excl.get("expense_types", []))
        skip_ids = set(excl.get("remove_ids", []))

        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        with connect(self.db_path) as conn:
            ids: dict[int, int] = {}
            for number, (fio, position, _company) in spr.items():
                key = fio_key(fio)
                found = conn.execute("SELECT id FROM signer_people WHERE fio_key = ?",
                                     (key,)).fetchone()
                if found:
                    ids[number] = found["id"]
                    continue
                cur = conn.execute(
                    "INSERT INTO signer_people (fio, fio_key, position, docv_uid, updated_at)"
                    " VALUES (?,?,?,'',?)", (fio, key, position, now))
                ids[number] = cur.lastrowid

            for name, numbers in raw.get("approver_lists", {}).items():
                rows = []
                for i, number in enumerate(numbers):
                    if number not in spr:
                        continue
                    _fio, position, company = spr[number]
                    rows.append((name, i, ids[number], position, company,
                                 "СОГЛАСОВАНО" if number == AGREED_MARK_ID else "",
                                 skip_types if number in skip_ids else ""))
                conn.execute("DELETE FROM signer_sets WHERE name = ?", (name,))
                conn.executemany(
                    "INSERT INTO signer_sets (name, ord, person_id, position, print_company,"
                    " mark, skip_expense_types) VALUES (?,?,?,?,?,?,?)", rows)

            def director(number: int) -> tuple:
                if number not in spr:
                    return (None, "", "")
                _fio, position, company = spr[number]
                return (ids[number], position, company)

            count = 0
            pairs = [(rule["company"], obj, rule["approvers"], rule["directors"])
                     for rule in raw.get("rules", []) for obj in rule["objects"]]
            pairs += [(fb["company"], "", fb["approvers"], fb["directors"])
                      for fb in raw.get("company_fallbacks", [])]
            for company, obj, set_name, directors in pairs:
                left = director(directors[0]) if directors else (None, "", "")
                right = director(directors[1]) if len(directors) > 1 else (None, "", "")
                conn.execute(
                    "INSERT OR REPLACE INTO signer_bindings (company, company_key,"
                    " object_name, object_key, set_name, soglasovano_id,"
                    " soglasovano_position, soglasovano_company, utverzhdayu_id,"
                    " utverzhdayu_position, utverzhdayu_company)"
                    " VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                    (company, normalize_name(company), obj, normalize_name(obj),
                     set_name, *left, *right))
                count += 1
        log.info("справочник подписантов наполнен",
                 extra={"data": {"bindings": count, "people": len(spr)}})
        return count

