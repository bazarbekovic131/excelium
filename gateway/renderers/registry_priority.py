"""Реестр предстоящих оплат (по приоритету): лист на объект.

Порт models/priority_registry.py; прямые обращения по ключам заменены
на .get() с пустыми значениями, дубли импортов убраны. Подписи под
таблицей раньше были вписаны в код двумя фамилиями — теперь берутся из
справочника подписантов шлюза, как и во внутреннем реестре.
"""
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Font

from .common import add_colontituls, format_row, set_print_area

DATA_COLS = ["A", "B", "C", "D", "E", "F", "G"]
START_ROW = 7


def render_priority(entries: list[dict], template_path, signers=None):
    workbook = openpyxl.load_workbook(template_path)
    template_sheet = workbook["REESTR"]
    font = Font(name="Arial", size=12)

    # Подписи под таблицей — из справочника шлюза по компании реестра;
    # объект у каждого листа свой, поэтому подбор идёт по первой записи.
    first = entries[0] if entries else {}
    resolved = signers.resolve(first.get("organization", ""),
                               first.get("object_name", "")) if signers else {}
    coordinators = resolved.get("coordinators") or []

    groups: dict[str, list[dict]] = defaultdict(list)
    for entry in entries:
        groups[str(entry.get("object_name") or "Без объекта")].append(entry)

    for object_name, group in groups.items():
        sheet = workbook.copy_worksheet(template_sheet)
        sheet.title = object_name[:31]

        sheet["A4"] = f"Объект: {object_name}"
        sheet["A5"] = f"Дата подачи заявки: {datetime.today().strftime('%d.%m.%y')}г."

        total = 0.0
        for idx, item in enumerate(group, start=1):
            row = START_ROW + idx - 1
            sheet[f"A{row}"] = idx
            sheet[f"B{row}"] = item.get("payment_number") or ""
            sheet[f"C{row}"] = item.get("status") or ""
            sheet[f"D{row}"] = item.get("counteragent") or ""
            sheet[f"E{row}"] = float(item.get("payment_sum") or 0)
            sheet[f"F{row}"] = item.get("TRU") or ""
            sheet[f"G{row}"] = item.get("object_name") or ""
            total += float(item.get("payment_sum") or 0)
            format_row(sheet, row, DATA_COLS, height=75)

        total_row = START_ROW + len(group)
        sheet[f"D{total_row}"] = "ВСЕГО"
        sheet[f"E{total_row}"] = total
        styled = [f"D{total_row}", f"E{total_row}"]

        for i, person in enumerate(coordinators):
            row = total_row + 2 + i * 2
            sheet[f"B{row}"] = str(person.get("position") or "")
            sheet[f"D{row}"] = f"___________________   {person.get('fio') or ''}"
            styled += [f"B{row}", f"D{row}"]
        for ref in styled:
            sheet[ref].font = font

    workbook.remove(template_sheet)
    for name in workbook.sheetnames:
        set_print_area(workbook[name], anchor_col="B", anchor_col_index=2, area="A1:I{row}")
        add_colontituls(workbook[name])
    return workbook
