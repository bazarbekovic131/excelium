import io
import json
from pathlib import Path

import openpyxl

from conftest import docv_headers

MODEL = json.loads((Path(__file__).parent / "data" / "model.json").read_text(encoding="utf-8"))


def _render(client, payload=None):
    r = client.post("/render/registry/inner", json=payload or MODEL, headers=docv_headers())
    assert r.status_code == 200, r.text
    url = r.json()["download_url"]
    token = url.rsplit("/", 1)[1]
    resp = client.get(f"/files/{token}")
    assert resp.status_code == 200
    return openpyxl.load_workbook(io.BytesIO(resp.content))


def _signature_lines(sheet) -> list[str]:
    return [str(sheet.cell(row=r, column=6).value)
            for r in range(18, sheet.max_row + 1)
            if sheet.cell(row=r, column=6).value]


def test_contract(client):
    wb = _render(client)
    # служебные листы в выдачу не попадают, REESTR удалён
    assert "REESTR" not in wb.sheetnames
    assert not [s for s in wb.sheetnames if s.startswith("СПР_")]
    # 13 позиций дают лист на каждую пару компания+объект
    pairs = {(d["organization"], d["object_name"]) for d in MODEL["request"]}
    assert len(wb.sheetnames) == len(pairs)

    sheet = next(wb[s] for s in wb.sheetnames if s.endswith("Администрация")
                 and wb[s]["F17"].value and "Шар-Кұрылыс" in wb[s]["F17"].value)
    assert sheet["G11"].value == "Администрация"
    assert str(sheet["F7"].value).startswith("РЕЕСТР ПЛАТЕЖЕЙ №20/")
    # позиция: стороны, сумма числом, назначение собрано
    assert sheet["F17"].value.startswith("Заявитель: ")
    assert isinstance(sheet["H17"].value, (int, float))
    assert sheet["I17"].value
    # подписи готовыми значениями: ни формул, ни номеров строк справочника
    assert sheet["B2"].value is None and sheet["B4"].value is None
    assert sheet["I3"].value and "Генеральный директор" in sheet["I3"].value
    assert sheet["I5"].value.endswith("Аманов Б.Ш.")
    lines = _signature_lines(sheet)
    assert lines, "нет строк подписантов"
    assert not [x for x in lines if x.startswith("=")]
    assert any("Главный бухгалтер" in x for x in lines)
    assert sheet.print_area


def test_empty_zatraty_does_not_crash(client):
    payload = {"request": [dict(MODEL["request"][0], zatraty=None)]}
    wb = _render(client, payload)
    assert wb  # раньше падало на .lower() от None


def test_expense_type_excludes_approvers(client):
    base = dict(MODEL["request"][0], organization='ТОО "СМУ Аргон"',
                object_name='ЖК "New Line"')  # набор list_3 содержит начальника ПТО
    wb_normal = _render(client, {"request": [dict(base, zatraty="СМР")]})
    wb_salary = _render(client, {"request": [dict(base, zatraty="Зарплата")]})

    def positions(wb):
        return _signature_lines(wb[wb.sheetnames[0]])

    excluded = "Начальник производственно-технического отдела"
    assert any(excluded in x for x in positions(wb_normal))
    assert not any(excluded in x for x in positions(wb_salary))


def test_unknown_company_yields_blank_signatures(client):
    payload = {"request": [dict(MODEL["request"][0], organization="ТОО «Никто»",
                                object_name="Нигде")]}
    wb = _render(client, payload)
    sheet = wb[wb.sheetnames[0]]
    # пустая строка при сохранении становится None — важно, что пусто
    assert not sheet["F2"].value and not sheet["I2"].value
    assert not sheet["F5"].value and not sheet["I5"].value
    assert not _signature_lines(sheet)


def test_company_name_variations_still_find_signers(client):
    """Doc-V меняет вид названия организации; подписанты не должны пропадать."""
    store = client.app.state.signers
    expected = store.resolve("ТОО «Шар-Кұрылыс»", "Администрация")
    assert expected["utverzhdayu"]
    for variant in ('ТОО «Шар-Кұрылыс» (KZT)', 'ТОО "Шар-Кұрылыс"',
                    'ТОО «Шар-Кұрылыс» ', 'ТОО  «Шар-Кұрылыс»(OLD)',
                    'тоо «шар-курылыс»'):
        assert store.resolve(variant, "Администрация") == expected, variant
    # объект тоже: другой стиль кавычек
    smu = store.resolve('ТОО "СМУ Аргон"', 'ЖК "Багыстан-1"')
    assert smu == store.resolve('ТОО «СМУ Аргон»', 'ЖК «Багыстан-1»')
    # неизвестная компания по-прежнему честно даёт пустой блок
    unknown = store.resolve("ТОО «Никто»", "Нигде")
    assert unknown["coordinators"] == [] and unknown["utverzhdayu"] is None
