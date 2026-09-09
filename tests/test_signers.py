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
    linked = store.link_directory(client.app.state.directory.all())
    assert linked == 1
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
