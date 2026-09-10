"""Кто подписывает реестр: справочник подписантов шлюза.

Doc-V про подписи ничего не присылает — наоборот, может спросить
(GET /signers). Источников правды два, и они не пересекаются:

- **люди** — Структура Doc-V: кто существует, как зовут, uid.
  Своего списка людей у шлюза нет;
- **должности шлюза** — в каком качестве человек подписывает: ключ,
  печатаемое название и кто занимает (сотрудник Структуры). У должности
  ровно один человек: подпись в реестре — одна строка, один человек.
  Если печатаемое название общее, а люди разные, это разные должности
  с одинаковым названием («Гендиректор Шар-Құрылыс» и «Гендиректор
  дочерних» обе печатаются как «Генеральный директор»).

Наборы согласующих и правила компаний ссылаются только на должности.
Своё у них одно — компания в подписи. Смена подписанта делается в одном
месте: у должности меняется человек.
"""
import json
import logging
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

from .jobsqueue.db import connect

log = logging.getLogger(__name__)

TOP_ROLES = ("soglasovano", "utverzhdayu")
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


def normalize_object(value) -> str:
    """Объект нормализуется мягче компании: у объектов хвост в скобках —
    это и есть различие («Школа (Нұра)» и «Школа (Тельман)» — разные
    стройки), а у компаний это мусор вроде «(KZT)»."""
    text = unicodedata.normalize("NFC", str(value or ""))
    text = text.translate(_QUOTES).translate(_KZ_FOLD)
    return re.sub(r"[\s\-]+", " ", text).strip().casefold()


def fio_key(value) -> str:
    """Фамилия и инициалы: «Аманов Б.Ш.» и «Аманов Бауыржан Шарипович»
    должны сойтись — в шаблоне писали инициалами, Doc-V шлёт полностью."""
    text = re.sub(r"[.,]", " ", print_name(value))
    parts = [p for p in re.split(r"\s+", text.translate(_KZ_FOLD)) if p]
    if not parts:
        return ""
    initials = "".join(p[0] for p in parts[1:3])
    return f"{parts[0]} {initials}".strip().casefold()


def print_name(value) -> str:
    """ФИО как печатать: Структура шлёт «Фамилия И.О. (Должность)», а
    в подпись должность идёт отдельной строкой — хвост в скобках долой."""
    text = unicodedata.normalize("NFC", str(value or "")).strip()
    return re.sub(r"\s*\([^()]*\)\s*$", "", text).strip()


def _person(fio: str, position: str, company: str, mark: str = "") -> dict:
    out = {"fio": fio, "position": position, "company": company}
    if mark:
        out["mark"] = mark
    return out


class SignerStore:
    def __init__(self, db_path: Path):
        self.db_path = db_path

    # --- Структура Doc-V ------------------------------------------------

    def staff(self) -> list[dict]:
        """Сотрудники из выгрузки Doc-V. Записи без имени пропускаются:
        подписывать нечем."""
        out = []
        with connect(self.db_path) as conn:
            rows = conn.execute("SELECT name, uid, data FROM directories").fetchall()
        for row in rows:
            data = json.loads(row["data"])
            raw = str(data.get("name") or data.get("display_name") or "").strip()
            if not raw:
                continue
            out.append({"uid": row["uid"], "fio": print_name(raw), "raw_name": raw,
                        "position": str(data.get("position") or "").strip(),
                        "department": str(data.get("department") or "").strip()})
        return sorted(out, key=lambda p: p["fio"])

    # --- должности шлюза --------------------------------------------------

    def roles(self, ctx: dict | None = None) -> list[dict]:
        """Все должности с тем, кто их занимает, и числом подписей,
        которые на них держатся."""
        ctx = ctx or self.context()
        with connect(self.db_path) as conn:
            rows = conn.execute("SELECT * FROM signer_roles ORDER BY name").fetchall()
            usage = self._usage(conn)
        out = []
        for row in rows:
            fio = self._holder_fio(row, ctx)
            out.append({"name": row["name"], "title": row["title"] or row["name"],
                        "holder_uid": row["holder_uid"], "holder_name": row["holder_name"],
                        "holder": fio, "vacant": not fio,
                        "in_structura": bool(row["holder_uid"]
                                             and row["holder_uid"] in ctx["staff"]),
                        "enabled": bool(row["enabled"]),
                        "used": usage.get(row["name"], 0)})
        return out

    def _usage(self, conn) -> dict[str, int]:
        counts: dict[str, int] = {}
        for column in ("soglasovano_role", "utverzhdayu_role"):
            for r in conn.execute(f"SELECT {column} AS role, COUNT(*) n FROM signer_rules"
                                  f" WHERE {column} <> '' GROUP BY {column}"):
                counts[r["role"]] = counts.get(r["role"], 0) + r["n"]
        for r in conn.execute("SELECT role, COUNT(*) n FROM signer_set_lines GROUP BY role"):
            counts[r["role"]] = counts.get(r["role"], 0) + r["n"]
        return counts

    def role_usage(self, name: str) -> int:
        with connect(self.db_path) as conn:
            return self._usage(conn).get(name, 0)

    def save_role(self, name: str, *, title: str = "", holder_uid: str = "",
                  holder_name: str = "", enabled: bool | None = None) -> str:
        """enabled=None — флаг не трогаем: сохранение каталога не должно
        возвращать из отпуска тех, кого отключили."""
        name = str(name or "").strip()
        if not name:
            raise ValueError("у должности должно быть имя")
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        flag = 1 if enabled is None or enabled else 0
        keep = "signer_roles.enabled" if enabled is None else "excluded.enabled"
        with connect(self.db_path) as conn:
            conn.execute(
                "INSERT INTO signer_roles (name, title, holder_uid, holder_name, enabled,"
                " updated_at) VALUES (?,?,?,?,?,?) ON CONFLICT(name) DO UPDATE SET"
                " title = excluded.title, holder_uid = excluded.holder_uid,"
                f" holder_name = excluded.holder_name, enabled = {keep},"
                " updated_at = excluded.updated_at",
                (name, str(title or "").strip(), str(holder_uid or "").strip(),
                 print_name(holder_name), flag, now))
        return name

    def set_roles_enabled(self, names: list[str], enabled: bool) -> int:
        """«В отпуске»: должность остаётся во всех правилах и наборах, но
        подпись от неё не печатается, пока её не включат обратно."""
        names = [str(n).strip() for n in names if str(n).strip()]
        if not names:
            return 0
        marks = ",".join("?" * len(names))
        with connect(self.db_path) as conn:
            cur = conn.execute(
                f"UPDATE signer_roles SET enabled = ? WHERE name IN ({marks})",
                (1 if enabled else 0, *names))
        return cur.rowcount

    def delete_roles(self, names: list[str]) -> tuple[int, list[str]]:
        """-> (удалено, пропущено): на используемую должность ссылаются
        подписи, её удалять нельзя."""
        deleted, skipped = 0, []
        with connect(self.db_path) as conn:
            usage = self._usage(conn)
            for name in names:
                name = str(name).strip()
                if not name:
                    continue
                if usage.get(name):
                    skipped.append(name)
                    continue
                deleted += conn.execute("DELETE FROM signer_roles WHERE name = ?",
                                        (name,)).rowcount
        return deleted, skipped

    def rename_role(self, old: str, new: str) -> None:
        """Переименование тянет за собой все ссылки: ключ — это и есть
        связь, других идентификаторов у должности нет."""
        new = str(new or "").strip()
        if not new or new == old:
            return
        with connect(self.db_path) as conn:
            if conn.execute("SELECT 1 FROM signer_roles WHERE name = ?", (new,)).fetchone():
                raise ValueError(f"должность «{new}» уже есть")
            conn.execute("UPDATE signer_roles SET name = ? WHERE name = ?", (new, old))
            conn.execute("UPDATE signer_set_lines SET role = ? WHERE role = ?", (new, old))
            for column in ("soglasovano_role", "utverzhdayu_role"):
                conn.execute(f"UPDATE signer_rules SET {column} = ? WHERE {column} = ?",
                             (new, old))

    def delete_role(self, name: str) -> None:
        if self.role_usage(name):
            raise ValueError(f"на должность «{name}» ссылаются подписи")
        with connect(self.db_path) as conn:
            conn.execute("DELETE FROM signer_roles WHERE name = ?", (name,))

    def link_by_name(self) -> dict:
        """Разовая помощь после переноса: должности, у которых человек
        вписан только именем, получают uid из Структуры по фамилии с
        инициалами. Дальше имя из Структуры главнее вписанного."""
        by_key: dict[str, str] = {}
        for person in self.staff():
            key = fio_key(person["fio"])
            if key and key not in by_key:
                by_key[key] = person["uid"]
        if not by_key:
            return {"staff": 0, "linked": 0, "left": 0}
        linked, left = 0, []
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        with connect(self.db_path) as conn:
            for row in conn.execute(
                    "SELECT name, holder_name FROM signer_roles"
                    " WHERE holder_uid = '' AND holder_name <> ''").fetchall():
                uid = by_key.get(fio_key(row["holder_name"]))
                if uid:
                    conn.execute("UPDATE signer_roles SET holder_uid = ?, updated_at = ?"
                                 " WHERE name = ?", (uid, now, row["name"]))
                    linked += 1
                else:
                    left.append(row["holder_name"])
        result = {"staff": len(by_key), "linked": linked, "left": len(left)}
        log.info("должности сопоставлены со Структурой",
                 extra={"data": {**result, "no_match": left[:10]}})
        return result

    # --- наборы ---------------------------------------------------------

    def sets(self, ctx: dict | None = None) -> dict[str, list[dict]]:
        ctx = ctx or self.context()
        out: dict[str, list[dict]] = {}
        with connect(self.db_path) as conn:
            names = [r["set_name"] for r in conn.execute(
                "SELECT DISTINCT set_name FROM signer_set_lines"
                " UNION SELECT DISTINCT set_name FROM signer_rules ORDER BY set_name")]
            for name in names:
                out[name] = []
            for r in conn.execute("SELECT * FROM signer_set_lines"
                                  " ORDER BY set_name, ord").fetchall():
                role = ctx["roles"].get(r["role"])
                out.setdefault(r["set_name"], []).append({
                    **dict(r), "title": (role["title"] or role["name"]) if role else "",
                    "holder": self._holder_fio(role, ctx) if role else "",
                    "missing": role is None})
        return out

    def save_set(self, name: str, lines: list[dict]) -> int:
        """lines: [{"role", "company", "mark", "skip_expense_types", "enabled"}]
        — порядок списка и есть порядок подписей. Выключенная строка
        остаётся на месте, но в реестр не печатается."""
        name = str(name or "").strip()
        if not name:
            raise ValueError("имя набора обязательно")
        rows = [(name, i, e["role"].strip(), (e.get("company") or "").strip(),
                 (e.get("mark") or "").strip(), (e.get("skip_expense_types") or "").strip(),
                 0 if str(e.get("enabled", "1")) in ("0", "False", "") else 1)
                for i, e in enumerate(e for e in lines if str(e.get("role") or "").strip())]
        with connect(self.db_path) as conn:
            conn.execute("DELETE FROM signer_set_lines WHERE set_name = ?", (name,))
            conn.executemany(
                "INSERT INTO signer_set_lines (set_name, ord, role, print_company, mark,"
                " skip_expense_types, enabled) VALUES (?,?,?,?,?,?,?)", rows)
        return len(rows)

    def delete_set(self, name: str) -> None:
        with connect(self.db_path) as conn:
            used = conn.execute("SELECT COUNT(*) c FROM signer_rules WHERE set_name = ?",
                                (name,)).fetchone()["c"]
            if used:
                raise ValueError(f"набор «{name}» используют {used} правил")
            conn.execute("DELETE FROM signer_set_lines WHERE set_name = ?", (name,))

    # --- правила компаний -------------------------------------------------

    def rules(self) -> list[dict]:
        with connect(self.db_path) as conn:
            return [dict(r) for r in conn.execute(
                "SELECT * FROM signer_rules ORDER BY company, object_name").fetchall()]

    def rule(self, rule_id: int) -> dict | None:
        with connect(self.db_path) as conn:
            row = conn.execute("SELECT * FROM signer_rules WHERE id = ?",
                               (rule_id,)).fetchone()
            return dict(row) if row else None

    def company_rules(self, company: str) -> list[dict]:
        with connect(self.db_path) as conn:
            return [dict(r) for r in conn.execute(
                "SELECT * FROM signer_rules WHERE company_key = ? ORDER BY object_name",
                (normalize_name(company),)).fetchall()]

    def save_rule(self, *, company: str, object_name: str, set_name: str,
                  soglasovano: str = "", soglasovano_company: str = "",
                  utverzhdayu: str = "", utverzhdayu_company: str = "",
                  rule_id: int | None = None, enabled: bool = True) -> int:
        company = str(company or "").strip()
        if not company:
            raise ValueError("компания обязательна")
        object_name = str(object_name or "").strip()
        values = (company, normalize_name(company), object_name, normalize_object(object_name),
                  str(set_name or "").strip(), str(soglasovano or "").strip(),
                  str(soglasovano_company or "").strip(), str(utverzhdayu or "").strip(),
                  str(utverzhdayu_company or "").strip(), 1 if enabled else 0)
        with connect(self.db_path) as conn:
            if rule_id:
                conn.execute(
                    "UPDATE signer_rules SET company=?, company_key=?, object_name=?,"
                    " object_key=?, set_name=?, soglasovano_role=?, soglasovano_company=?,"
                    " utverzhdayu_role=?, utverzhdayu_company=?, enabled=? WHERE id=?",
                    (*values, rule_id))
                return rule_id
            cur = conn.execute(
                "INSERT INTO signer_rules (company, company_key, object_name, object_key,"
                " set_name, soglasovano_role, soglasovano_company, utverzhdayu_role,"
                " utverzhdayu_company, enabled) VALUES (?,?,?,?,?,?,?,?,?,?)"
                " ON CONFLICT(company_key, object_key) DO UPDATE SET"
                " company=excluded.company, object_name=excluded.object_name,"
                " set_name=excluded.set_name, soglasovano_role=excluded.soglasovano_role,"
                " soglasovano_company=excluded.soglasovano_company,"
                " utverzhdayu_role=excluded.utverzhdayu_role,"
                " utverzhdayu_company=excluded.utverzhdayu_company,"
                " enabled=excluded.enabled", values)
            return cur.lastrowid or 0

    def delete_rule(self, rule_id: int) -> None:
        with connect(self.db_path) as conn:
            conn.execute("DELETE FROM signer_rules WHERE id = ?", (rule_id,))

    def _ids(self, ids) -> list[int]:
        return [int(i) for i in ids if str(i).strip().lstrip("-").isdigit()]

    def set_rules_enabled(self, ids: list, enabled: bool) -> int:
        """Отключённое правило объекта уступает правилу компании, а
        отключённое правило компании — как будто его нет."""
        ids = self._ids(ids)
        if not ids:
            return 0
        marks = ",".join("?" * len(ids))
        with connect(self.db_path) as conn:
            cur = conn.execute(f"UPDATE signer_rules SET enabled = ? WHERE id IN ({marks})",
                               (1 if enabled else 0, *ids))
        return cur.rowcount

    def delete_rules(self, ids: list) -> int:
        ids = self._ids(ids)
        if not ids:
            return 0
        marks = ",".join("?" * len(ids))
        with connect(self.db_path) as conn:
            cur = conn.execute(f"DELETE FROM signer_rules WHERE id IN ({marks})", ids)
        return cur.rowcount

    def apply_to_company(self, company: str, *, soglasovano: dict | None = None,
                         utverzhdayu: dict | None = None,
                         set_name: str | None = None,
                         rule_ids: list | None = None) -> int:
        """Проставить одно и то же во все правила компании (или только в
        выбранные rule_ids). Утверждающий у компании один на все объекты,
        и правкой по одной строке это занятие на полдня. None — поле не
        трогаем; {"role": "", …} — очистить."""
        sets_sql, values = [], []
        for role, slot in (("soglasovano", soglasovano), ("utverzhdayu", utverzhdayu)):
            if slot is None:
                continue
            sets_sql += [f"{role}_role=?", f"{role}_company=?"]
            values += [str(slot.get("role") or "").strip(),
                       str(slot.get("company") or "").strip()]
        if set_name:
            sets_sql.append("set_name=?")
            values.append(set_name.strip())
        if not sets_sql:
            return 0
        where, params = "company_key = ?", [normalize_name(company)]
        ids = self._ids(rule_ids or [])
        if ids:
            where += f" AND id IN ({','.join('?' * len(ids))})"
            params += ids
        with connect(self.db_path) as conn:
            cur = conn.execute(
                f"UPDATE signer_rules SET {', '.join(sets_sql)} WHERE {where}",
                (*values, *params))
            return cur.rowcount

    # --- подбор ---------------------------------------------------------

    def context(self) -> dict:
        """Разово прочитанные Структура и должности — чтобы сводная
        таблица на сотню строк не ходила в базу за каждой."""
        with connect(self.db_path) as conn:
            roles = {r["name"]: dict(r) for r in conn.execute("SELECT * FROM signer_roles")}
        return {"staff": {p["uid"]: p for p in self.staff()}, "roles": roles}

    def _holder_fio(self, role: dict | None, ctx: dict) -> str:
        """Имя из Структуры главнее вписанного руками: вписанное — это
        запас на случай, если человека в Doc-V нет."""
        if not role:
            return ""
        person = ctx["staff"].get(role["holder_uid"]) if role["holder_uid"] else None
        return person["fio"] if person else print_name(role["holder_name"])

    def _signature(self, role_name: str, company: str, ctx: dict,
                   mark: str = "") -> dict | None:
        role = ctx["roles"].get(role_name)
        if role is None:
            log.warning("должности нет в справочнике — подпись останется пустой",
                        extra={"data": {"role": role_name}})
            return None
        if not role.get("enabled", 1):
            log.info("должность отключена — подпись не печатается",
                     extra={"data": {"role": role_name}})
            return None
        fio = self._holder_fio(role, ctx)
        if not fio:
            log.warning("на должность никто не назначен — подпись останется пустой",
                        extra={"data": {"role": role_name}})
            return None
        return _person(fio, role["title"] or role["name"], company, mark)

    def resolve(self, company: str, object_name: str = "",
                expense_type: str = "", ctx: dict | None = None) -> dict:
        """-> {"soglasovano": …|None, "utverzhdayu": …|None,
                "coordinators": [...], "source": …}."""
        company_key = normalize_name(company)
        object_key = normalize_object(object_name)
        with connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT * FROM signer_rules WHERE company_key = ? AND object_key = ?"
                " AND enabled = 1", (company_key, object_key)).fetchone()
            source = "объект"
            if row is None:
                row = conn.execute(
                    "SELECT * FROM signer_rules WHERE company_key = ? AND object_key = ''"
                    " AND enabled = 1", (company_key,)).fetchone()
                source = "компания"
            if row is None:
                log.warning("нет правила подписей: реестр выйдет без подписей",
                            extra={"data": {"company": company, "object": object_name}})
                return {"soglasovano": None, "utverzhdayu": None,
                        "coordinators": [], "source": "нет правила"}
            whole = row
            if source == "объект":
                whole = conn.execute(
                    "SELECT * FROM signer_rules WHERE company_key = ? AND object_key = ''"
                    " AND enabled = 1", (company_key,)).fetchone() or row
            lines = conn.execute(
                "SELECT * FROM signer_set_lines WHERE set_name = ? AND enabled = 1"
                " ORDER BY ord", (row["set_name"],)).fetchall()
        ctx = ctx or self.context()

        def active(role_name: str) -> bool:
            role = ctx["roles"].get(role_name)
            return bool(role and role.get("enabled", 1))

        expense = str(expense_type or "").strip().casefold()
        coordinators = []
        for line in lines:
            skip = {t.strip().casefold() for t in (line["skip_expense_types"] or "").split(",")
                    if t.strip()}
            if expense and expense in skip:
                continue
            person = self._signature(line["role"], line["print_company"], ctx, line["mark"])
            if person:
                coordinators.append(person)

        def top(role: str) -> dict | None:
            # у объекта не задано или должность в отпуске — берём правило компании
            source_row = row if row[f"{role}_role"] and active(row[f"{role}_role"]) else whole
            if not source_row[f"{role}_role"]:
                return None
            return self._signature(source_row[f"{role}_role"],
                                   source_row[f"{role}_company"] or row["company"], ctx)

        return {"soglasovano": top("soglasovano"), "utverzhdayu": top("utverzhdayu"),
                "coordinators": coordinators, "source": source}

    def stats(self) -> dict[str, int]:
        with connect(self.db_path) as conn:
            return {
                "roles": conn.execute("SELECT COUNT(*) c FROM signer_roles").fetchone()["c"],
                "sets": conn.execute(
                    "SELECT COUNT(DISTINCT set_name) c FROM signer_set_lines").fetchone()["c"],
                "rules": conn.execute("SELECT COUNT(*) c FROM signer_rules").fetchone()["c"],
                "vacant": sum(1 for r in self.roles() if r["vacant"]),
                "disabled_roles": conn.execute(
                    "SELECT COUNT(*) c FROM signer_roles WHERE enabled = 0").fetchone()["c"],
            }

    # --- экспорт и импорт -------------------------------------------------

    def export_json(self) -> dict:
        """Весь состав одним словарём без служебных ключей: ключи
        компаний и объектов пересчитываются при импорте."""
        with connect(self.db_path) as conn:
            roles = [{"name": r["name"], "title": r["title"], "holder_uid": r["holder_uid"],
                      "holder_name": r["holder_name"], "enabled": bool(r["enabled"])}
                     for r in conn.execute("SELECT * FROM signer_roles ORDER BY name")]
            sets: dict[str, list[dict]] = {}
            for r in conn.execute("SELECT * FROM signer_set_lines ORDER BY set_name, ord"):
                sets.setdefault(r["set_name"], []).append({
                    "role": r["role"], "print_company": r["print_company"],
                    "mark": r["mark"], "skip_expense_types": r["skip_expense_types"],
                    "enabled": bool(r["enabled"])})
            rules = [{"company": r["company"], "object_name": r["object_name"],
                      "set_name": r["set_name"],
                      "soglasovano_role": r["soglasovano_role"],
                      "soglasovano_company": r["soglasovano_company"],
                      "utverzhdayu_role": r["utverzhdayu_role"],
                      "utverzhdayu_company": r["utverzhdayu_company"],
                      "enabled": bool(r["enabled"])}
                     for r in conn.execute("SELECT * FROM signer_rules"
                                           " ORDER BY company, object_name")]
        return {"version": 1,
                "exported_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "roles": roles, "sets": sets, "rules": rules}

    def import_json(self, data: dict, *, mode: str = "replace") -> dict[str, int]:
        """Полная замена состава содержимым файла. Только replace: слияние
        двух составов без общего идентификатора неоднозначно."""
        if mode != "replace":
            raise ValueError("поддерживается только полная замена")
        if not isinstance(data, dict) or data.get("version") != 1:
            raise ValueError("ожидается файл экспорта версии 1")
        roles, sets, rules = data.get("roles"), data.get("sets"), data.get("rules")
        if not isinstance(roles, list) or not isinstance(sets, dict) \
                or not isinstance(rules, list):
            raise ValueError("в файле должны быть roles, sets и rules")
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")

        def flag(item: dict) -> int:
            return 0 if item.get("enabled") in (False, 0, "0") else 1

        with connect(self.db_path) as conn:
            for table in ("signer_rules", "signer_set_lines", "signer_roles"):
                conn.execute(f"DELETE FROM {table}")
            conn.executemany(
                "INSERT INTO signer_roles (name, title, holder_uid, holder_name, enabled,"
                " updated_at) VALUES (?,?,?,?,?,?)",
                [(str(r.get("name") or "").strip(), str(r.get("title") or ""),
                  str(r.get("holder_uid") or ""), print_name(r.get("holder_name")),
                  flag(r), now) for r in roles if str(r.get("name") or "").strip()])
            lines = []
            for set_name, items in sets.items():
                for i, e in enumerate(items if isinstance(items, list) else []):
                    if str(e.get("role") or "").strip():
                        lines.append((set_name, i, str(e["role"]).strip(),
                                      str(e.get("print_company") or ""),
                                      str(e.get("mark") or ""),
                                      str(e.get("skip_expense_types") or ""), flag(e)))
            conn.executemany(
                "INSERT INTO signer_set_lines (set_name, ord, role, print_company, mark,"
                " skip_expense_types, enabled) VALUES (?,?,?,?,?,?,?)", lines)
            rows = []
            for r in rules:
                company = str(r.get("company") or "").strip()
                if not company:
                    continue
                obj = str(r.get("object_name") or "").strip()
                rows.append((company, normalize_name(company), obj, normalize_object(obj),
                             str(r.get("set_name") or ""), str(r.get("soglasovano_role") or ""),
                             str(r.get("soglasovano_company") or ""),
                             str(r.get("utverzhdayu_role") or ""),
                             str(r.get("utverzhdayu_company") or ""), flag(r)))
            conn.executemany(
                "INSERT OR REPLACE INTO signer_rules (company, company_key, object_name,"
                " object_key, set_name, soglasovano_role, soglasovano_company,"
                " utverzhdayu_role, utverzhdayu_company, enabled)"
                " VALUES (?,?,?,?,?,?,?,?,?,?)", rows)
            counts = {"roles": conn.execute("SELECT COUNT(*) c FROM signer_roles").fetchone()["c"],
                      "sets": len({ln[0] for ln in lines}),
                      "rules": conn.execute("SELECT COUNT(*) c FROM signer_rules").fetchone()["c"]}
        log.info("состав подписантов импортирован", extra={"data": counts})
        return counts

    # --- наполнение для тестов --------------------------------------------

    def seed(self, matrix_path: Path, template_path: Path) -> int:
        """Наполнение из старых источников: структура — из YAML, ФИО и
        должности — с листа СПР_ПОДПИСАНТОВ шаблона. Боевой состав давно
        перенесён и живёт в базе; сюда ходят только тесты
        (tests/data/signers_seed.yaml), чтобы работать на реальном объёме.

        Должность получает ключ по названию, если в старом справочнике
        её занимал один человек, иначе — с фамилией в скобках: одиннадцать
        генеральных директоров становятся одиннадцатью должностями с
        одним печатаемым названием."""
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

        holders: dict[str, set[str]] = {}
        for fio, position, _company in spr.values():
            holders.setdefault(position, set()).add(fio)

        def role_key(number: int) -> str:
            fio, position, _company = spr[number]
            return position if len(holders[position]) == 1 else f"{position} ({fio})"

        excl = raw.get("expense_type_exclusions", {})
        skip_types = ",".join(excl.get("expense_types", []))
        skip_ids = set(excl.get("remove_ids", []))
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")

        with connect(self.db_path) as conn:
            for number, (fio, position, _company) in spr.items():
                conn.execute(
                    "INSERT OR IGNORE INTO signer_roles (name, title, holder_uid,"
                    " holder_name, updated_at) VALUES (?,?,?,?,?)",
                    (role_key(number), position, "", fio, now))

            for name, numbers in raw.get("approver_lists", {}).items():
                conn.execute("DELETE FROM signer_set_lines WHERE set_name = ?", (name,))
                conn.executemany(
                    "INSERT INTO signer_set_lines (set_name, ord, role, print_company,"
                    " mark, skip_expense_types) VALUES (?,?,?,?,?,?)",
                    [(name, i, role_key(n), spr[n][2],
                      "СОГЛАСОВАНО" if n == AGREED_MARK_ID else "",
                      skip_types if n in skip_ids else "")
                     for i, n in enumerate(numbers) if n in spr])

            def director(number: int) -> tuple[str, str]:
                return (role_key(number), spr[number][2]) if number in spr else ("", "")

            pairs = [(rule["company"], obj, rule["approvers"], rule["directors"])
                     for rule in raw.get("rules", []) for obj in rule["objects"]]
            pairs += [(fb["company"], "", fb["approvers"], fb["directors"])
                      for fb in raw.get("company_fallbacks", [])]
            for company, obj, set_name, directors in pairs:
                left = director(directors[0]) if directors else ("", "")
                right = director(directors[1]) if len(directors) > 1 else ("", "")
                conn.execute(
                    "INSERT OR REPLACE INTO signer_rules (company, company_key, object_name,"
                    " object_key, set_name, soglasovano_role, soglasovano_company,"
                    " utverzhdayu_role, utverzhdayu_company) VALUES (?,?,?,?,?,?,?,?,?)",
                    (company, normalize_name(company), obj, normalize_object(obj),
                     set_name, *left, *right))
            count = conn.execute("SELECT COUNT(*) c FROM signer_rules").fetchone()["c"]
            roles = conn.execute("SELECT COUNT(*) c FROM signer_roles").fetchone()["c"]
        log.info("справочник подписантов наполнен",
                 extra={"data": {"rules": count, "roles": roles, "spr_rows": len(spr)}})
        return count
