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
import json
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


SLOT_SEP = "||"


GW_PREFIX = "gw:"   # должность шлюза; uid и должности Структуры так не начинаются


def gw_slot(name: str) -> str:
    """Выбор «должность шлюза»: подпись идёт от того, кого на неё
    назначили, и меняется в одном месте для всех компаний сразу."""
    return slot_value(f"{GW_PREFIX}{name}")


def slot_value(ref: str = "", dept: str = "", person_id=None) -> str:
    """Одно поле формы на выбор подписанта: должность шлюза, должность
    или сотрудник из Структуры, либо человек из справочника. Ссылка
    сильнее — с неё и начинаем."""
    if ref:
        return f"r:{ref}{SLOT_SEP}{dept or ''}"
    if person_id:
        return f"p:{person_id}"
    return ""


def parse_slot(value: str) -> dict:
    value = str(value or "").strip()
    if value.startswith("r:"):
        ref, _, dept = value[2:].partition(SLOT_SEP)
        return {"ref": ref, "dept": dept, "person_id": None}
    if value.startswith("p:") and value[2:].isdigit():
        return {"ref": "", "dept": "", "person_id": int(value[2:])}
    return {"ref": "", "dept": "", "person_id": None}


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

    # --- должности шлюза --------------------------------------------------

    def gateway_positions(self) -> list[str]:
        """Названия должностей шлюза. В подпись печатается именно оно, а не
        должность из Структуры: у Doc-V она одна на человека, а подписывать
        он может как директор разных компаний."""
        with connect(self.db_path) as conn:
            return [r["name"] for r in conn.execute(
                "SELECT name FROM signer_positions ORDER BY name")]

    def position_cards(self) -> list[dict]:
        """Должности шлюза вместе с назначенным человеком и числом мест,
        где на должность ссылаются подписи."""
        staff = {p["uid"]: p for p in self.staff()}
        with connect(self.db_path) as conn:
            people = {r["id"]: r["fio"] for r in
                      conn.execute("SELECT id, fio FROM signer_people")}
            rows = conn.execute("SELECT * FROM signer_positions ORDER BY name").fetchall()
            out = []
            for row in rows:
                ref = f"{GW_PREFIX}{row['name']}"
                used = conn.execute(
                    "SELECT (SELECT COUNT(*) FROM signer_bindings"
                    "  WHERE soglasovano_ref = ? OR utverzhdayu_ref = ?)"
                    " + (SELECT COUNT(*) FROM signer_sets WHERE position_ref = ?) AS n",
                    (ref, ref, ref)).fetchone()["n"]
                holder = staff.get(row["holder_uid"])
                out.append({
                    "name": row["name"],
                    "value": slot_value(ref),
                    "holder_uid": row["holder_uid"],
                    "holder_person_id": row["holder_person_id"],
                    "holder": (holder["fio"] if holder
                               else people.get(row["holder_person_id"], "")),
                    "holder_slot": slot_value(row["holder_uid"], "",
                                              row["holder_person_id"]),
                    "vacant": not holder and not people.get(row["holder_person_id"]),
                    "used": used,
                })
        return out

    def assign_position(self, name: str, *, holder_uid: str = "",
                        holder_person_id=None) -> None:
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        with connect(self.db_path) as conn:
            conn.execute(
                "INSERT INTO signer_positions (name, holder_uid, holder_person_id,"
                " updated_at) VALUES (?,?,?,?) ON CONFLICT(name) DO UPDATE SET"
                " holder_uid = excluded.holder_uid,"
                " holder_person_id = excluded.holder_person_id,"
                " updated_at = excluded.updated_at",
                (name.strip(), holder_uid.strip(), holder_person_id or None, now))

    def position_usage(self, name: str) -> int:
        ref = f"{GW_PREFIX}{name}"
        with connect(self.db_path) as conn:
            return conn.execute(
                "SELECT (SELECT COUNT(*) FROM signer_bindings"
                "  WHERE soglasovano_ref = ? OR utverzhdayu_ref = ?)"
                " + (SELECT COUNT(*) FROM signer_sets WHERE position_ref = ?) AS n",
                (ref, ref, ref)).fetchone()["n"]

    def add_position(self, name: str) -> str:
        name = str(name or "").strip()
        if not name:
            return ""
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        with connect(self.db_path) as conn:
            conn.execute("INSERT OR IGNORE INTO signer_positions (name, updated_at)"
                         " VALUES (?,?)", (name, now))
        return name

    def delete_position(self, name: str) -> None:
        with connect(self.db_path) as conn:
            conn.execute("DELETE FROM signer_positions WHERE name = ?", (name,))

    def _remember_positions(self, *names) -> None:
        """Должность, набранную руками, каталог подхватывает сам — иначе
        в ста тринадцати привязках развелись бы опечатки."""
        for name in names:
            self.add_position(name)

    # --- Структура Doc-V: кто сейчас занимает должность -------------------

    def staff(self) -> list[dict]:
        """Сотрудники из выгрузки Doc-V. Записи без имени пропускаются:
        подписывать нечем."""
        out = []
        with connect(self.db_path) as conn:
            rows = conn.execute("SELECT name, uid, data FROM directories").fetchall()
        for row in rows:
            data = json.loads(row["data"])
            fio = str(data.get("name") or data.get("display_name") or "").strip()
            if not fio:
                continue
            out.append({"uid": row["uid"], "fio": fio,
                        "position": str(data.get("position") or "").strip(),
                        "department": str(data.get("department") or "").strip(),
                        "department_uid": str(data.get("department_uid") or "").strip(),
                        "directory": row["name"]})
        return sorted(out, key=lambda p: p["fio"])

    def positions(self) -> list[dict]:
        """Должности из Структуры для выпадающих списков: сама должность,
        отдел и кто её сейчас занимает."""
        grouped: dict[tuple[str, str], list[str]] = {}
        for person in self.staff():
            if not person["position"]:
                continue
            grouped.setdefault((person["position"], person["department"]), []).append(
                person["fio"])
        return [{"position": position, "department": department,
                 "holders": sorted(x for x in holders if x)}
                for (position, department), holders in sorted(grouped.items())]

    def _holder(self, staff: list[dict], ref: str, dept: str,
                ctx: dict | None = None) -> dict | None:
        """Кого подставить по ссылке. Ссылкой может быть uid сотрудника —
        тогда подпись всегда от него, а ФИО обновляется вслед за Doc-V —
        либо название должности: тогда подставляется тот, кто занимает её
        сейчас, и подпись переезжает вместе с назначением."""
        ref = str(ref or "").strip()
        if not ref:
            return None
        dept = str(dept or "").strip()
        if ref.startswith(GW_PREFIX):
            return self._gw_holder(staff, ref[len(GW_PREFIX):], ctx)
        by_uid = [p for p in staff if p["uid"] == ref]
        if by_uid:
            return by_uid[0]
        found = [p for p in staff if p["position"] == ref
                 and (not dept or p["department"] == dept)]
        if not found:
            log.warning("в Структуре никого нет по этой ссылке — подпись останется пустой",
                        extra={"data": {"ref": ref, "department": dept}})
            return None
        if len(found) > 1:
            log.warning("должность занимают несколько человек — беру первого по алфавиту",
                        extra={"data": {"position_ref": ref, "department": dept,
                                        "holders": [p["fio"] for p in found]}})
        return sorted(found, key=lambda p: p["fio"])[0]

    def _gw_holder(self, staff: list[dict], name: str,
                   ctx: dict | None = None) -> dict | None:
        """Кого назначили на должность шлюза. Должность и есть то, что
        печатается, поэтому она же идёт в position."""
        cached = (ctx or {}).get("gw")
        if cached is not None:
            row = cached.get(name)
            if row is None:
                log.warning("должности шлюза нет в каталоге — подпись останется пустой",
                            extra={"data": {"position": name}})
                return None
            fio = ""
            if row["holder_uid"]:
                found = [p for p in staff if p["uid"] == row["holder_uid"]]
                fio = found[0]["fio"] if found else ""
            if not fio and row["holder_person_id"]:
                fio = (ctx or {}).get("people", {}).get(row["holder_person_id"], "")
            if not fio:
                log.warning("на должность шлюза никто не назначен —"
                            " подпись останется пустой",
                            extra={"data": {"position": name}})
                return None
            return {"uid": row["holder_uid"], "fio": fio, "position": name,
                    "department": ""}
        with connect(self.db_path) as conn:
            row = conn.execute("SELECT * FROM signer_positions WHERE name = ?",
                               (name,)).fetchone()
            if row is None:
                log.warning("должности шлюза нет в каталоге — подпись останется пустой",
                            extra={"data": {"position": name}})
                return None
            fio = ""
            if row["holder_uid"]:
                found = [p for p in staff if p["uid"] == row["holder_uid"]]
                fio = found[0]["fio"] if found else ""
            if not fio and row["holder_person_id"]:
                person = conn.execute("SELECT fio FROM signer_people WHERE id = ?",
                                      (row["holder_person_id"],)).fetchone()
                fio = person["fio"] if person else ""
        if not fio:
            log.warning("на должность шлюза никто не назначен — подпись останется пустой",
                        extra={"data": {"position": name}})
            return None
        return {"uid": row["holder_uid"], "fio": fio, "position": name, "department": ""}

    # --- подбор ---------------------------------------------------------

    def context(self) -> dict:
        """Разово прочитанные Структура и каталог должностей — чтобы
        сводная таблица на сотню строк не ходила в базу за каждой."""
        with connect(self.db_path) as conn:
            gw = {r["name"]: dict(r) for r in
                  conn.execute("SELECT * FROM signer_positions")}
            people = {r["id"]: r["fio"] for r in
                      conn.execute("SELECT id, fio FROM signer_people")}
        return {"staff": self.staff(), "gw": gw, "people": people}

    def resolve(self, company: str, object_name: str = "",
                expense_type: str = "", ctx: dict | None = None) -> dict:
        """-> {"soglasovano": …|None, "utverzhdayu": …|None,
                "coordinators": [...], "source": …}."""
        company_key = normalize_name(company)
        object_key = normalize_object(object_name)
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
            whole = None
            if source == "объект":
                whole = conn.execute(
                    "SELECT * FROM signer_bindings WHERE company_key = ?"
                    " AND object_key = ''", (company_key,)).fetchone()
            names = self._names(conn)
            entries = conn.execute(
                "SELECT * FROM signer_sets WHERE name = ? ORDER BY ord",
                (row["set_name"],)).fetchall()
        ctx = ctx or {}
        staff = ctx.get("staff")
        if staff is None:
            staff = self.staff()

        def fio_of(person_id, ref, dept) -> tuple[str, str]:
            """-> (ФИО, должность по умолчанию). Ссылка на должность
            сильнее записанного человека: она и нужна, чтобы подпись
            менялась вместе с составом Doc-V."""
            holder = self._holder(staff, ref, dept, ctx)
            if holder:
                return holder["fio"], holder["position"]
            return names.get(person_id, ("", ""))

        expense = str(expense_type or "").strip().casefold()
        coordinators = []
        for entry in entries:
            skip = [t.strip().casefold() for t in (entry["skip_expense_types"] or "").split(",")]
            if expense and expense in [t for t in skip if t]:
                continue
            fio, default_position = fio_of(entry["person_id"], entry["position_ref"],
                                           entry["dept_ref"])
            if not fio:
                continue
            coordinators.append(_person(fio, entry["position"] or default_position,
                                        entry["print_company"], entry["mark"]))

        def top(role: str) -> dict | None:
            source = row
            if not row[f"{role}_id"] and not row[f"{role}_ref"] and whole is not None:
                source = whole   # у объекта не задано — берём правило компании
            person_id = source[f"{role}_id"]
            ref = source[f"{role}_ref"]
            if not person_id and not ref:
                return None
            fio, default_position = fio_of(person_id, ref, source[f"{role}_dept"])
            if not fio:
                return None
            return _person(fio, source[f"{role}_position"] or default_position,
                           source[f"{role}_company"] or row["company"])

        return {"soglasovano": top(ROLE_SOGLASOVANO),
                "utverzhdayu": top(ROLE_UTVERZHDAYU),
                "coordinators": coordinators, "source": source}

    def _names(self, conn) -> dict[int, tuple[str, str]]:
        return {r["id"]: (r["fio"], r["position"])
                for r in conn.execute("SELECT id, fio, position FROM signer_people")}

    # --- связь со Структурой Doc-V ---------------------------------------

    def link_directory(self, directories: dict[str, dict[str, dict]] | None = None) -> dict:
        """Сверяет людей справочника с выгрузкой Структуры по фамилии с
        инициалами: запоминает uid и подтягивает написание ФИО. Связь
        нужна и сама по себе — по ней видно, кого Doc-V уже не знает."""
        by_key: dict[str, tuple[str, str]] = {}
        for person in self.staff():
            key = fio_key(person["fio"])
            if key and key not in by_key:
                by_key[key] = (person["uid"], person["fio"])
        if not by_key:
            return {"staff": 0, "matched": 0, "renamed": 0, "unmatched": 0}
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        matched = renamed = 0
        unmatched = []
        with connect(self.db_path) as conn:
            for row in conn.execute(
                    "SELECT id, fio, fio_key, docv_uid FROM signer_people").fetchall():
                found = by_key.get(row["fio_key"])
                if not found:
                    unmatched.append(row["fio"])
                    continue
                matched += 1
                if found[0] == row["docv_uid"] and found[1] == row["fio"]:
                    continue
                if found[1] != row["fio"]:
                    renamed += 1
                conn.execute(
                    "UPDATE signer_people SET docv_uid = ?, fio = ?, updated_at = ?"
                    " WHERE id = ?", (found[0], found[1], now, row["id"]))
        result = {"staff": len(by_key), "matched": matched, "renamed": renamed,
                  "unmatched": len(unmatched)}
        log.info("подписанты сверены со Структурой",
                 extra={"data": {**result, "no_match": unmatched[:10]}})
        return result

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

    def people_usage(self) -> dict[int, int]:
        """Сколько подписей держится на каждом человеке — чтобы в списке
        было сразу видно, кого удалять нельзя."""
        counts: dict[int, int] = {}
        with connect(self.db_path) as conn:
            for column in ("soglasovano_id", "utverzhdayu_id"):
                for row in conn.execute(
                        f"SELECT {column} AS id, COUNT(*) AS n FROM signer_bindings"
                        f" WHERE {column} IS NOT NULL GROUP BY {column}"):
                    counts[row["id"]] = counts.get(row["id"], 0) + row["n"]
            for row in conn.execute(
                    "SELECT person_id AS id, COUNT(*) AS n FROM signer_sets"
                    " WHERE person_id IS NOT NULL GROUP BY person_id"):
                counts[row["id"]] = counts.get(row["id"], 0) + row["n"]
        return counts

    def person_usage(self, person_id: int) -> list[str]:
        """Где человек подписывает: удалять его вслепую нельзя, иначе
        реестр молча выйдет без подписи."""
        with connect(self.db_path) as conn:
            used = [f"{r['company']} · {r['object_name'] or 'вся компания'}"
                    for r in conn.execute(
                        "SELECT company, object_name FROM signer_bindings"
                        " WHERE soglasovano_id = ? OR utverzhdayu_id = ?",
                        (person_id, person_id)).fetchall()]
            used += [f"набор {r['name']}" for r in conn.execute(
                "SELECT DISTINCT name FROM signer_sets WHERE person_id = ?",
                (person_id,)).fetchall()]
        return used

    def delete_person(self, person_id: int) -> None:
        with connect(self.db_path) as conn:
            conn.execute("DELETE FROM signer_people WHERE id = ?", (person_id,))

    def save_binding(self, *, company: str, object_name: str, set_name: str,
                     soglasovano: dict | None, utverzhdayu: dict | None,
                     binding_id: int | None = None) -> int:
        """soglasovano/utverzhdayu: {"person_id", "position", "company",
        "ref", "dept"} — ref указывает должность в Структуре Doc-V, и тогда
        подпись идёт от того, кто её занимает сейчас."""
        company = company.strip()
        if not company:
            raise ValueError("компания обязательна")
        object_name = object_name.strip()
        left = soglasovano or {}
        right = utverzhdayu or {}
        values = (company, normalize_name(company), object_name, normalize_object(object_name),
                  set_name.strip(),
                  left.get("person_id") or None, (left.get("position") or "").strip(),
                  (left.get("company") or "").strip(), (left.get("ref") or "").strip(),
                  (left.get("dept") or "").strip(),
                  right.get("person_id") or None, (right.get("position") or "").strip(),
                  (right.get("company") or "").strip(), (right.get("ref") or "").strip(),
                  (right.get("dept") or "").strip())
        self._remember_positions(left.get("position"), right.get("position"))
        with connect(self.db_path) as conn:
            if binding_id:
                conn.execute(
                    "UPDATE signer_bindings SET company=?, company_key=?, object_name=?,"
                    " object_key=?, set_name=?, soglasovano_id=?, soglasovano_position=?,"
                    " soglasovano_company=?, soglasovano_ref=?, soglasovano_dept=?,"
                    " utverzhdayu_id=?, utverzhdayu_position=?, utverzhdayu_company=?,"
                    " utverzhdayu_ref=?, utverzhdayu_dept=? WHERE id=?",
                    (*values, binding_id))
                return binding_id
            cur = conn.execute(
                "INSERT INTO signer_bindings (company, company_key, object_name, object_key,"
                " set_name, soglasovano_id, soglasovano_position, soglasovano_company,"
                " soglasovano_ref, soglasovano_dept, utverzhdayu_id, utverzhdayu_position,"
                " utverzhdayu_company, utverzhdayu_ref, utverzhdayu_dept)"
                " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                " ON CONFLICT(company_key, object_key) DO UPDATE SET"
                " company=excluded.company, object_name=excluded.object_name,"
                " set_name=excluded.set_name, soglasovano_id=excluded.soglasovano_id,"
                " soglasovano_position=excluded.soglasovano_position,"
                " soglasovano_company=excluded.soglasovano_company,"
                " soglasovano_ref=excluded.soglasovano_ref,"
                " soglasovano_dept=excluded.soglasovano_dept,"
                " utverzhdayu_id=excluded.utverzhdayu_id,"
                " utverzhdayu_position=excluded.utverzhdayu_position,"
                " utverzhdayu_company=excluded.utverzhdayu_company,"
                " utverzhdayu_ref=excluded.utverzhdayu_ref,"
                " utverzhdayu_dept=excluded.utverzhdayu_dept", values)
            return cur.lastrowid or 0

    def apply_to_company(self, company: str, *, soglasovano: dict | None = None,
                         utverzhdayu: dict | None = None,
                         set_name: str | None = None) -> int:
        """Проставить одно и то же во все привязки компании. Утверждающий
        у компании один на все объекты, и правкой по одной строке это
        занятие на полдня. None — поле не трогаем."""
        company_key = normalize_name(company)
        sets_sql, values = [], []
        for role, slot in (("soglasovano", soglasovano), ("utverzhdayu", utverzhdayu)):
            if slot is None:
                continue
            sets_sql += [f"{role}_id=?", f"{role}_ref=?", f"{role}_dept=?",
                         f"{role}_position=?", f"{role}_company=?"]
            values += [slot.get("person_id") or None, (slot.get("ref") or "").strip(),
                       (slot.get("dept") or "").strip(),
                       (slot.get("position") or "").strip(),
                       (slot.get("company") or "").strip()]
            self._remember_positions(slot.get("position"))
        if set_name:
            sets_sql.append("set_name=?")
            values.append(set_name.strip())
        if not sets_sql:
            return 0
        with connect(self.db_path) as conn:
            cur = conn.execute(
                f"UPDATE signer_bindings SET {', '.join(sets_sql)} WHERE company_key = ?",
                (*values, company_key))
            return cur.rowcount

    def company_bindings(self, company: str) -> list[dict]:
        with connect(self.db_path) as conn:
            return [dict(r) for r in conn.execute(
                "SELECT * FROM signer_bindings WHERE company_key = ?"
                " ORDER BY object_name", (normalize_name(company),)).fetchall()]

    def delete_binding(self, binding_id: int) -> None:
        with connect(self.db_path) as conn:
            conn.execute("DELETE FROM signer_bindings WHERE id = ?", (binding_id,))

    def save_set(self, name: str, entries: list[dict]) -> int:
        """entries: [{"person_id"|"ref", "dept", "position", "company", "mark",
        "skip_expense_types"}] — порядок списка и есть порядок подписей."""
        name = name.strip()
        if not name:
            raise ValueError("имя набора обязательно")
        self._remember_positions(*[e.get("position") for e in entries])
        with connect(self.db_path) as conn:
            conn.execute("DELETE FROM signer_sets WHERE name = ?", (name,))
            conn.executemany(
                "INSERT INTO signer_sets (name, ord, person_id, position, position_ref,"
                " dept_ref, print_company, mark, skip_expense_types)"
                " VALUES (?,?,?,?,?,?,?,?,?)",
                [(name, i, e.get("person_id") or None, (e.get("position") or "").strip(),
                  (e.get("ref") or "").strip(), (e.get("dept") or "").strip(),
                  (e.get("company") or "").strip(), (e.get("mark") or "").strip(),
                  (e.get("skip_expense_types") or "").strip())
                 for i, e in enumerate(entries) if e.get("person_id") or e.get("ref")])
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
                    rows.append((name, i, ids[number], position, "", "", company,
                                 "СОГЛАСОВАНО" if number == AGREED_MARK_ID else "",
                                 skip_types if number in skip_ids else ""))
                conn.execute("DELETE FROM signer_sets WHERE name = ?", (name,))
                conn.executemany(
                    "INSERT INTO signer_sets (name, ord, person_id, position, position_ref,"
                    " dept_ref, print_company, mark, skip_expense_types)"
                    " VALUES (?,?,?,?,?,?,?,?,?)", rows)

            def director(number: int) -> tuple:
                if number not in spr:
                    return (None, "", "")
                _fio, position, company = spr[number]
                return (ids[number], position, company)

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
                    (company, normalize_name(company), obj, normalize_object(obj),
                     set_name, *left, *right))
            # Должность, которую в старом справочнике занимал ровно один
            # человек, сразу получает его: каталог становится рабочим без
            # ручного назначения. «Генеральный директор» у двух десятков
            # компаний разный, поэтому остаётся вакантным.
            holders: dict[str, set[int]] = {}
            for number, (_fio, position, _company) in spr.items():
                if position:
                    holders.setdefault(position, set()).add(ids[number])
            conn.executemany(
                "INSERT OR IGNORE INTO signer_positions (name, holder_person_id,"
                " updated_at) VALUES (?,?,?)",
                [(position, next(iter(people_ids)) if len(people_ids) == 1 else None, now)
                 for position, people_ids in holders.items()])
            count = conn.execute(
                "SELECT COUNT(*) c FROM signer_bindings").fetchone()["c"]
            people = conn.execute(
                "SELECT COUNT(*) c FROM signer_people").fetchone()["c"]
        log.info("справочник подписантов наполнен",
                 extra={"data": {"bindings": count, "people": people,
                                 "spr_rows": len(spr)}})
        return count

