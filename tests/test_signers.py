"""Справочник подписантов шлюза: подбор, должности по компании, связь
со Структурой Doc-V."""
import io
import json
from pathlib import Path

import openpyxl

from conftest import docv_headers
from gateway.signers import fio_key

MODEL = json.loads((Path(__file__).parent / "data" / "model.json").read_text(encoding="utf-8"))


def _render(client, kind, payload):
    r = client.post(f"/render/registry/{kind}", json=payload, headers=docv_headers())
    assert r.status_code == 200, r.text
    token = r.json()["download_url"].rsplit("/", 1)[1]
    return openpyxl.load_workbook(io.BytesIO(client.get(f"/files/{token}").content))


def test_seed_filled_from_old_sources(client):
    stats = client.app.state.signers.stats()
    assert stats["bindings"] > 50 and stats["people"] > 20 and stats["sets"] >= 5


def test_one_person_signs_for_many_companies_with_own_position(client):
    """Один и тот же человек — гендиректор десятков компаний. Должность и
    компания в подписи берутся из привязки, а не из карточки человека."""
    store = client.app.state.signers
    seen = {}
    for binding in store.bindings():
        resolved = store.resolve(binding["company"], binding["object_name"])
        person = resolved["utverzhdayu"]
        if person:
            seen.setdefault(person["fio"], set()).add(person["company"])
    multi = {fio: companies for fio, companies in seen.items() if len(companies) > 1}
    assert multi, "в справочнике должен быть подписант сразу нескольких компаний"
    for companies in multi.values():
        assert all(companies), "у каждой подписи своя компания"


def test_resolve_endpoint_serves_docv(client):
    r = client.get("/signers", params={"company": "ТОО «Шар-Кұрылыс»",
                                       "object": "Администрация"},
                   headers=docv_headers())
    assert r.status_code == 200
    body = r.json()
    assert body["utverzhdayu"]["fio"]
    assert body["utverzhdayu"]["position"]
    assert body["coordinators"]
    assert body["source"] in ("объект", "компания")


def test_resolve_endpoint_needs_token(client):
    assert client.get("/signers").status_code == 403


def test_object_binding_wins_over_company(client):
    store = client.app.state.signers
    person = store.people()[0]
    store.save_binding(company="ТОО «Тест»", object_name="", set_name="list_1",
                       soglasovano=None,
                       utverzhdayu={"person_id": person["id"], "position": "Директор",
                                    "company": "ТОО «Тест»"})
    store.save_binding(company="ТОО «Тест»", object_name="Объект", set_name="list_1",
                       soglasovano=None,
                       utverzhdayu={"person_id": person["id"],
                                    "position": "Управляющий директор",
                                    "company": "ТОО «Тест»"})
    assert store.resolve("ТОО «Тест»", "Объект")["utverzhdayu"]["position"] == \
        "Управляющий директор"
    assert store.resolve("ТОО «Тест»", "Другой")["utverzhdayu"]["position"] == "Директор"
    assert store.resolve("ТОО «Тест»", "Другой")["source"] == "компания"


def test_directory_link_refreshes_fio(client):
    """Структура Doc-V шлёт ФИО полностью, в шаблоне были инициалы —
    после связывания печатается то, что в Doc-V."""
    store = client.app.state.signers
    before = store.resolve("ТОО «Шар-Кұрылыс»", "Администрация")["utverzhdayu"]["fio"]
    assert before == "Аманов Б.Ш."
    client.post("/directory/structura", headers=docv_headers(), json={"items": [
        {"uid": "u-1", "display_name": "Аманов Бауыржан Шарипович",
         "position": "Генеральный директор", "department": "Дирекция"}]})
    result = store.link_directory()
    assert result["matched"] == 1 and result["renamed"] == 1
    after = store.resolve("ТОО «Шар-Кұрылыс»", "Администрация")["utverzhdayu"]
    assert after["fio"] == "Аманов Бауыржан Шарипович"
    # должность осталась той, под которой он подписывает эту компанию
    assert after["position"] == "Генеральный директор"
    assert after["company"]


def test_fio_key_matches_initials_and_full_name():
    assert fio_key("Аманов Б.Ш.") == fio_key("Аманов Бауыржан Шарипович")
    assert fio_key("Татин Ә. Ж") == fio_key("Татин Адилет Жанович")
    assert fio_key("Аманов Б.Ш.") != fio_key("Аманов Дархан Ерланович")
    assert fio_key("") == ""


def test_priority_registry_takes_coordinators_from_store(client):
    entry = dict(MODEL["request"][0], organization="ТОО «Шар-Кұрылыс»",
                 object_name="Администрация", payment_number="1", status="в работе")
    wb = _render(client, "priority", {"request": [entry]})
    sheet = wb[wb.sheetnames[0]]
    lines = [str(sheet.cell(row=r, column=2).value)
             for r in range(9, sheet.max_row + 1) if sheet.cell(row=r, column=2).value]
    expected = client.app.state.signers.resolve("ТОО «Шар-Кұрылыс»", "Администрация")
    assert lines and len(lines) == len(expected["coordinators"])
    assert lines[0] == expected["coordinators"][0]["position"]


def test_person_edit_changes_every_signature(client):
    store = client.app.state.signers
    person = next(p for p in store.people() if p["fio"] == "Омарова Г.А.")
    store.save_person(person["id"], "Омарова Гульнара Алиевна", "Главный бухгалтер")
    resolved = store.resolve("ТОО «Шар-Кұрылыс»", "Администрация")
    assert any(c["fio"] == "Омарова Гульнара Алиевна" for c in resolved["coordinators"])


def test_objects_with_brackets_stay_distinct(client):
    """У объектов скобки — это различие: «Школа (Нұра)» и «Школа (Тельман)»
    разные стройки. У компаний скобки — мусор вроде «(KZT)»."""
    store = client.app.state.signers
    objects = {b["object_name"] for b in store.bindings()
               if b["company"] == 'ТОО "СМУ Аргон"'}
    assert {"Школа (Нұра)", "Школа (Тельман)", "Школа (Уркер)"} <= objects
    person = store.people()[0]
    store.save_binding(company='ТОО "СМУ Аргон"', object_name="Школа (Нұра)",
                       set_name="list_1", soglasovano=None,
                       utverzhdayu={"person_id": person["id"], "position": "Прораб",
                                    "company": "тест"})
    assert store.resolve('ТОО "СМУ Аргон"', "Школа (Нұра)")["utverzhdayu"]["position"] \
        == "Прораб"
    other = store.resolve('ТОО "СМУ Аргон"', "Школа (Тельман)")
    assert other["utverzhdayu"] is None or other["utverzhdayu"]["position"] != "Прораб"


def _structura(client, items):
    r = client.post("/directory/structura", headers=docv_headers(), json={"items": items})
    assert r.status_code == 200


def test_signature_follows_whoever_holds_the_position(client):
    """Главное: привязка указывает должность, а подпись идёт от того, кто
    её занимает сейчас. Сменился сотрудник в Doc-V — сменилась подпись."""
    store = client.app.state.signers
    _structura(client, [{"uid": "u-10", "display_name": "Петров Пётр Петрович",
                         "position": "Финансовый директор", "department": "Финансы"}])
    store.save_binding(company="ТОО «Смена»", object_name="", set_name="list_1",
                       soglasovano=None,
                       utverzhdayu={"ref": "Финансовый директор", "dept": "Финансы",
                                    "position": "Генеральный директор",
                                    "company": "ТОО «Смена»"})
    first = store.resolve("ТОО «Смена»")["utverzhdayu"]
    assert first["fio"] == "Петров Пётр Петрович"
    # печатается должность из привязки, а не из Структуры
    assert first["position"] == "Генеральный директор"

    _structura(client, [{"uid": "u-11", "display_name": "Сидорова Анна Ивановна",
                         "position": "Финансовый директор", "department": "Финансы"}])
    second = store.resolve("ТОО «Смена»")["utverzhdayu"]
    assert second["fio"] == "Сидорова Анна Ивановна"
    assert second["position"] == "Генеральный директор"


def test_position_reference_can_use_uid(client):
    """Ссылкой годится и шифр записи: Doc-V шлёт то название, то uid."""
    store = client.app.state.signers
    _structura(client, [{"uid": "uid-777", "display_name": "Ким Олег Сергеевич",
                         "position": "Начальник ЮО", "department": "Юротдел"}])
    store.save_binding(company="ТОО «Шифр»", object_name="", set_name="list_1",
                       soglasovano={"ref": "uid-777", "position": "Начальник ЮО",
                                    "company": "ТОО «Шифр»"},
                       utverzhdayu=None)
    assert store.resolve("ТОО «Шифр»")["soglasovano"]["fio"] == "Ким Олег Сергеевич"


def test_vacant_position_leaves_signature_empty(client, caplog):
    store = client.app.state.signers
    _structura(client, [{"uid": "u-12", "display_name": "Кто-то", "position": "Кладовщик"}])
    store.save_binding(company="ТОО «Пусто»", object_name="", set_name="list_1",
                       soglasovano=None,
                       utverzhdayu={"ref": "Такой должности нет", "position": "Директор",
                                    "company": "ТОО «Пусто»"})
    assert store.resolve("ТОО «Пусто»")["utverzhdayu"] is None


def test_positions_list_shows_holders(client):
    _structura(client, [
        {"uid": "u-20", "display_name": "Первый И.И.", "position": "Прораб",
         "department": "СМР"},
        {"uid": "u-21", "display_name": "Второй П.П.", "position": "Прораб",
         "department": "СМР"}])
    positions = client.app.state.signers.positions()
    prorab = next(p for p in positions if p["position"] == "Прораб")
    assert prorab["department"] == "СМР"
    assert prorab["holders"] == ["Второй П.П.", "Первый И.И."]


def test_employee_reference_keeps_person_and_gateway_position(client):
    """Второй способ: выбрать сотрудника по uid и дать ему должность
    шлюза. Печатается должность шлюза, а не та, что в Структуре."""
    store = client.app.state.signers
    _structura(client, [{"uid": "05ab617e-6fff", "display_name": "ДРС",
                         "position": "Начальник участка",
                         "department": "МЖК New Line",
                         "department_uid": "a0227bc0-580a"}])
    store.save_binding(company="ТОО «Уид»", object_name="", set_name="list_1",
                       soglasovano=None,
                       utverzhdayu={"ref": "05ab617e-6fff",
                                    "position": "Генеральный директор",
                                    "company": "ТОО «Уид»"})
    signer = store.resolve("ТОО «Уид»")["utverzhdayu"]
    assert signer["fio"] == "ДРС"
    assert signer["position"] == "Генеральный директор"   # должность шлюза
    assert signer["company"] == "ТОО «Уид»"

    # Doc-V переименовал сотрудника — подпись идёт за ним
    _structura(client, [{"uid": "05ab617e-6fff", "display_name": "Дюсенов Р.С.",
                         "position": "Заместитель директора",
                         "department": "МЖК New Line"}])
    again = store.resolve("ТОО «Уид»")["utverzhdayu"]
    assert again["fio"] == "Дюсенов Р.С."
    assert again["position"] == "Генеральный директор"


def test_gateway_position_catalogue(client):
    store = client.app.state.signers
    catalogue = store.gateway_positions()
    assert "Генеральный директор" in catalogue and "Главный бухгалтер" in catalogue
    # должность, набранная руками, попадает в каталог сама
    person = store.people()[0]
    store.save_binding(company="ТОО «Каталог»", object_name="", set_name="list_1",
                       soglasovano=None,
                       utverzhdayu={"person_id": person["id"],
                                    "position": "Управляющий партнёр",
                                    "company": "ТОО «Каталог»"})
    assert "Управляющий партнёр" in store.gateway_positions()
    store.delete_position("Управляющий партнёр")
    assert "Управляющий партнёр" not in store.gateway_positions()
    # из каталога убрали, а подпись осталась прежней
    assert store.resolve("ТОО «Каталог»")["utverzhdayu"]["position"] == "Управляющий партнёр"


def test_staff_without_position_still_selectable(client):
    """Запись без должности годится как сотрудник, но не как должность."""
    _structura(client, [{"uid": "u-30", "display_name": "Безлошадный Б.Б."}])
    store = client.app.state.signers
    assert any(p["uid"] == "u-30" for p in store.staff())
    assert not any(p["position"] == "" for p in store.positions())
