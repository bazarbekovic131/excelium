from datetime import datetime
import logging
# from shutil import copy2
from utils.scripts import format_row, set_border, find_last_row_in_col, load_excel, hide_sheets, create_concatenated_info, set_print_area, add_colontituls, set_cell_properties
import re
import utils.firmen_und_objekte as firmobj
from openpyxl.styles import Alignment, Font, Border, Side

EMPTY_MATCH = ([0] * 2, [0] * 8)


def _variants(value):
    """Написания названия, которые стоит попробовать в матрице.

    Матрица сравнивает строки точно, поэтому «ТОО «Шар-Кұрылыс» (KZT)»,
    прямые кавычки или лишний пробел дают нули вместо согласующих.
    Doc-V со временем меняет способ вывода поля организации, и подбор
    не должен от этого разваливаться.
    """
    seen, out = set(), []
    base = str(value or "").strip()
    for text in (base,
                 re.sub(r"\s*\([^()]*\)\s*$", "", base).strip(),  # хвост «(KZT)», «(OLD)»
                 re.sub(r"\s+", " ", base)):
        for quoted in (text,
                       text.replace('"', "«", 1).replace('"', "»", 1),
                       text.replace("«", '"').replace("»", '"')):
            candidate = quoted.strip()
            if candidate and candidate not in seen:
                seen.add(candidate)
                out.append(candidate)
    return out


def lookup_coordinators(company, object_name):
    """Подбор подписантов, устойчивый к написанию названий."""
    for c in _variants(company):
        for o in _variants(object_name):
            found = firmobj.check_company_object_pair(c, o)
            if found != EMPTY_MATCH:
                return found
    logging.warning(
        "подписанты не найдены: компания %r, объект %r — реестр выйдет с нулями",
        company, object_name)
    return EMPTY_MATCH

def add_coordinators_v4(sheet):
    '''
    Adds coordinators to the specified sheet.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet to add coordinators to.

    Returns:
        None
    '''

    col_index = 6 #F
    formula_f = '=IFERROR(IF(ISNUMBER(VALUE(INDIRECT("B" & ROW()))), VLOOKUP(VALUE(INDIRECT("B" & ROW())), СПР_ПОДПИСАНТОВ!$B$14:$K$100, 9, 0) & " " & VLOOKUP(VALUE(INDIRECT("B" & ROW())), СПР_ПОДПИСАНТОВ!$B$14:$K$100, 7, 0), ""),"")'
    formula_i = '=IFERROR(IF(ISNUMBER(VALUE(INDIRECT("B" & ROW()))), VLOOKUP(VALUE(INDIRECT("B" & ROW())), СПР_ПОДПИСАНТОВ!$B$14:$K$100, 5, 0), ""),"")'

    final_row = find_last_row_in_col(sheet, col_index)

    if final_row:
        logging.info(f"The last non-empty cell in column {chr(64 + col_index)} of sheet '{sheet.title}' is in row {final_row}.")
    else:
        logging.info(f"No non-empty cells found in column {chr(64 + col_index)} of sheet '{sheet.title}'.")
        final_row = 0 # never happens

    company = sheet['H11'].value # Get the company name from the sheet
    object_name = sheet['G11'].value # Get the object name from the sheet
    doctype = sheet['H12'].value

    directors, coordinators_list = lookup_coordinators(company, object_name)
    specified_types = ['коммерческие расходы', 'зарплата', 'налоги']
    if sheet['G17'].value.lower() in specified_types:
        coordinators_list = [i for i in coordinators_list if i not in [4, 6]] # remove selected approvers in payments of commercial expenses
    else:
        pass # continue without changing

    # Turned off
    # if sheet['H12'].value.lower() == 'заявка на налоги': # for tax related payments
    #     coordinators_list =firmobj.return_administration_approvers() # approvers always correspond to List 1 approvers
    # TODO: remove head of PTO from approval of salary. code of PTO: 6

    n = len(coordinators_list) # Get the number of coordinators
    logging.info(f'Company: {company}, Object: {object_name} Number of coordinators: {n}; coordinators: {coordinators_list}')
    for i in range(1,n+1):
        formula_b = f'=INDEX(СПР_ОБЪЕКТОВ!$B$7:$K$80, MATCH($G11, СПР_ОБЪЕКТОВ!$B$7:$B$80, 0), {3 + i})'
        row = final_row + i * 3

        if coordinators_list[i-1] != 3:
            set_cell_properties(sheet, row, 2, coordinators_list[i-1], set_border('thin'))
            set_cell_properties(sheet, row, 6, formula_f, None, Alignment(horizontal='left'), Font(size=14, bold=True))
            set_cell_properties(sheet, row, 9, formula_i, None, Alignment(horizontal='right'), Font(size=14, bold=True))
        else:
            set_cell_properties(sheet, row, 6, "СОГЛАСОВАНО", None, Alignment(horizontal='left'), Font(bold=False))
            set_cell_properties(sheet, row + 2, 2, 3, set_border('thin'))
            set_cell_properties(sheet, row + 2, 6, formula_f, None, Alignment(horizontal='left'), Font(size=14, bold=True))
            set_cell_properties(sheet, row + 2, 9, formula_i, None, Alignment(horizontal='right'), Font(size=14, bold=True))

    # Add directors (final piece)
    sheet['B2'] = directors[0]
    sheet['B4'] = directors[1]
    sheet['H11'] = '' # remove the temporary company name
    sheet['H12'] = '' # remove the temporary document type

def loop_json(json_data, workbook):
    '''
    This function works with the loaded json and with the copied workbook.
    Groups entries by (organization, object_name) pairs so each company 
    gets its own sheet per object.
    '''
    from collections import defaultdict
    
    cols = ['F', 'G', 'H', 'I']
    
    for key_title, data in json_data.items():
        print(f"Processing {len(data)} documents")
        
        # Group entries by (company, object) pairs
        groups = defaultdict(list)
        for item in data:
            key = (item.get('organization', ''), item.get('object_name', ''))
            groups[key].append(item)
        
        # Create company number mapping for sheet naming
        unique_companies = sorted(set(company for company, obj in groups.keys()))
        company_numbers = {company: idx + 1 for idx, company in enumerate(unique_companies)}
        
        sub = 0  # sub-nomer reestra
        
        for (company, object_name), entries in groups.items():
            sub += 1
            
            # Create sheet name using company number (to fit Excel's 31-char limit)
            company_num = company_numbers[company]
            # Format: "C{num}_{object_name}" - e.g., "C1_Администрация"
            sheet_title = f"C{company_num}_{object_name}"
            if len(sheet_title) > 31:
                sheet_title = sheet_title[:31]
            
            # Create a copy of the source sheet
            source_sheet = workbook['REESTR']
            new_sheet = workbook.copy_worksheet(source_sheet)
            new_sheet.title = sheet_title
            
            # Set sheet metadata
            workbook[sheet_title]['G11'] = object_name
            workbook[sheet_title]['G10'] = datetime.today()
            registry_name = entries[0].get('registry_name', '')
            workbook[sheet_title]['F7'] = f'{registry_name}/{sub}'
            
            # Store company info for coordinators (H11, H12 are used by add_coordinators_v4)
            workbook[sheet_title]['H11'] = company
            workbook[sheet_title]['H12'] = entries[0].get('doctype', '')
            
            # Write all entries for this (company, object) group
            start_row = 17
            for entry in entries:
                row = start_row
                
                sides_str = f'Заявитель: {entry["organization"]}' + '\n\n' + f'Кому: {entry["counteragent"]}'
                workbook[sheet_title][f'F{row}'] = sides_str
                workbook[sheet_title][f'G{row}'] = entry.get('zatraty', '')
                workbook[sheet_title][f'H{row}'] = float(entry.get('payment_sum', 0))
                
                i_cell_str = create_concatenated_info(entry)
                workbook[sheet_title][f'I{row}'] = i_cell_str
                
                # Format the row
                format_row(workbook[sheet_title], row, cols)
                workbook[sheet_title][f'F{row}'].alignment = Alignment(
                    horizontal='left', vertical='center', wrap_text=True
                )
                
                start_row += 1

def format_excel_inner(json_data):
    logging.info('Opening template.xlsx')
    workbook = load_excel('excel_templates/template.xlsx')
    initial_sheets = ['REESTR', 'СПР_ОБЪЕКТОВ', 'СПР_ПОДПИСАНТОВ']

    logging.info('Reading JSON file')
    loop_json(json_data, workbook)

    logging.info('Removing REESTR sheet')
    if 'REESTR' in workbook.sheetnames:
        reestr_sheet = workbook['REESTR']
        workbook.remove(reestr_sheet)

    # Update the initial_sheets list after removing REESTR
    initial_sheets.remove('REESTR')
    
    hide_sheets(workbook, initial_sheets)
    for sheet in workbook.sheetnames:
        if sheet not in initial_sheets:
            add_coordinators_v4(workbook[sheet])
            set_print_area(workbook[sheet])
            add_colontituls(workbook[sheet])
        else:
            continue

    return workbook
