"""Внутренний реестр платежей: группировка по паре компания+объект,
лист на каждую пару, блок подписей из справочника подписантов шлюза.

Порт models/inner_registry.py. Отличия от оригинала:
- подписи печатаются готовыми строками, а не номерами строк с VLOOKUP:
  служебные листы СПР_ПОДПИСАНТОВ и СПР_ОБЪЕКТОВ в выдачу не попадают,
  и выпущенный реестр больше не меняется задним числом;
- компания/вид затрат больше не гоняются через служебные ячейки H11/H12
  (данные передаются напрямую, ячейки и так затирались в конце);
- удалена formula_b — она вычислялась и никогда не записывалась,
  а её MATCH не совпадал с именами листа СПР_ОБЪЕКТОВ;
- пустой «Вид затрат» не роняет реестр (раньше .lower() на None).
"""
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font

from ..signers import SignerStore
from .common import (add_colontituls, create_concatenated_info, find_last_row_in_col,
                     format_row, set_cell_properties, set_print_area)

DATA_COLS = ["F", "G", "H", "I"]
START_ROW = 17
UNDERLINE = "_" * 28
SPR_SHEETS = ("СПР_ПОДПИСАНТОВ", "СПР_ОБЪЕКТОВ")


def render_inner(entries: list[dict], template_path, signers: SignerStore):
    workbook = openpyxl.load_workbook(template_path)

    groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for item in entries:
        groups[(item.get("organization", ""), item.get("object_name", ""))].append(item)

    unique_companies = sorted({company for company, _ in groups})
    company_numbers = {c: i + 1 for i, c in enumerate(unique_companies)}

    template_sheet = workbook["REESTR"]
    for sub, ((company, object_name), group) in enumerate(groups.items(), start=1):
        title = f"C{company_numbers[company]}_{object_name}"[:31]
        sheet = workbook.copy_worksheet(template_sheet)
        sheet.title = title

        sheet["G11"] = object_name
        sheet["G10"] = datetime.today()
        sheet["F7"] = f"{group[0].get('registry_name', '')}/{sub}"

        row = START_ROW
        for entry in group:
            sheet[f"F{row}"] = (f"Заявитель: {entry.get('organization', '')}\n\n"
                                f"Кому: {entry.get('counteragent', '')}")
            sheet[f"G{row}"] = entry.get("zatraty") or ""
            sheet[f"H{row}"] = float(entry.get("payment_sum") or 0)
            sheet[f"I{row}"] = create_concatenated_info(entry)
            format_row(sheet, row, DATA_COLS, height=100, left_align=("F",))
            row += 1

        resolved = signers.resolve(company, object_name,
                                   expense_type=str(group[0].get("zatraty") or ""))
        _add_signatures(sheet, resolved)
        set_print_area(sheet, anchor_col="F", anchor_col_index=6, area="F1:I{row}")
        add_colontituls(sheet)

    workbook.remove(template_sheet)
    for name in SPR_SHEETS:
        if name in workbook.sheetnames:
            workbook.remove(workbook[name])
    return workbook


def _line(person: dict | None) -> tuple[str, str]:
    """-> (должность с компанией, ФИО)."""
    if not isinstance(person, dict):
        return "", ""
    line = " ".join(x for x in (str(person.get("position") or "").strip(),
                                str(person.get("company") or "").strip()) if x)
    return line, str(person.get("fio") or "").strip()


def _add_signatures(sheet, resolved: dict) -> None:
    """Подписи готовыми значениями в те же ячейки, что раньше заполняли
    формулы шаблона. Верх: слева «СОГЛАСОВАНО:» (F2/F3/F5), справа
    «УТВЕРЖДАЮ» (I2/I3/I5). Ниже — согласующие с шагом в три строки;
    запись с «mark» даёт строку-заголовок и подпись двумя ниже."""
    final_row = find_last_row_in_col(sheet, 6) or START_ROW - 1

    left_line, left_fio = _line(resolved.get("soglasovano"))
    right_line, right_fio = _line(resolved.get("utverzhdayu"))
    sheet["F2"] = "СОГЛАСОВАНО:" if left_fio else ""
    sheet["F3"] = left_line
    sheet["F5"] = f"{UNDERLINE} {left_fio}" if left_fio else ""
    if not right_fio:
        sheet["I2"] = ""  # статичный «УТВЕРЖДАЮ» шаблона не должен висеть над пустотой
    sheet["I3"] = right_line
    sheet["I5"] = f"{UNDERLINE} {right_fio}" if right_fio else ""
    # номера строк справочника больше не пишутся: ячейки очищаются,
    # чтобы в файле не осталось следов прежней схемы с VLOOKUP
    sheet["B2"] = None
    sheet["B4"] = None

    left_align = Alignment(horizontal="left")
    right_align = Alignment(horizontal="right")
    bold14 = Font(size=14, bold=True)
    for i, person in enumerate(resolved.get("coordinators") or [], start=1):
        row = final_row + i * 3
        line, fio = _line(person)
        mark = str(person.get("mark") or "").strip()
        if mark:
            set_cell_properties(sheet, row, 6, mark, None, left_align, Font(bold=False))
            row += 2
        set_cell_properties(sheet, row, 6, line, None, left_align, bold14)
        set_cell_properties(sheet, row, 9, fio, None, right_align, bold14)
