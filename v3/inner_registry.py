from datetime import datetime
import logging
from shutil import copy2
from scripts import format_row, set_border, find_last_row_in_col, load_excel, hide_sheets, create_concatenated_info, set_print_area, add_colontituls, set_cell_properties
import openpyxl
import firmen_und_objekten as firmobj
from openpyxl.styles import Alignment, Font, Border, Side



def add_coordinators(sheet):
    '''
    Adds coordinators to the specified sheet.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet to add coordinators to.

    Returns:
        None
    '''

    col_index = 6 #F
    formula_f = '=IF(ISNUMBER(VALUE(INDIRECT("B" & ROW()))), VLOOKUP(VALUE(INDIRECT("B" & ROW())), СПР_ПОДПИСАНТОВ!$B$14:$K$100, 9, 0) & " " & VLOOKUP(VALUE(INDIRECT("B" & ROW())), СПР_ПОДПИСАНТОВ!$B$14:$K$100, 7, 0), "")'
    formula_i = '=IF(ISNUMBER(VALUE(INDIRECT("B" & ROW()))), VLOOKUP(VALUE(INDIRECT("B" & ROW())), СПР_ПОДПИСАНТОВ!$B$14:$K$100, 5, 0), "")'

    final_row = find_last_row_in_col(sheet, col_index)

    if final_row:
        logging.info(f"The last non-empty cell in column {chr(64 + col_index)} of sheet '{sheet.title}' is in row {final_row}.")
    else:
        logging.info(f"No non-empty cells found in column {chr(64 + col_index)} of sheet '{sheet.title}'.")

    # Loop through the specified rows (8 times)
    for i in range(1,8):
        formula_b = f'=INDEX(СПР_ОБЪЕКТОВ!$B$7:$K$80, MATCH($G11, СПР_ОБЪЕКТОВ!$B$7:$B$80, 0), {3 + i})'
        row = final_row + i * 3
        set_cell_properties(sheet, row, 2, formula_b, set_border('thin'))
        set_cell_properties(sheet, row, 6, formula_f, None, Alignment(horizontal='left'), Font(size=14, bold=True))
        set_cell_properties(sheet, row, 9, formula_i, None, Alignment(horizontal='right'), Font(size=14, bold=True))

    last_row = final_row + 24
    set_cell_properties(sheet, last_row, 6, "СОГЛАСОВАНО", None, Alignment(horizontal='left'), Font(bold=False))
    set_cell_properties(sheet, last_row + 2, 2, 3, set_border('thin'))
    set_cell_properties(sheet, last_row + 2, 6, formula_f, None, Alignment(horizontal='left'), Font(size=14, bold=True))
    set_cell_properties(sheet, last_row + 2, 9, formula_i, None, Alignment(horizontal='right'), Font(size=14, bold=True))

def add_coordinators_v4(sheet):
    '''
    Adds coordinators to the specified sheet.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet to add coordinators to.

    Returns:
        None
    '''

    col_index = 6 #F
    formula_f = '=IF(ISNUMBER(VALUE(INDIRECT("B" & ROW()))), VLOOKUP(VALUE(INDIRECT("B" & ROW())), СПР_ПОДПИСАНТОВ!$B$14:$K$100, 9, 0) & " " & VLOOKUP(VALUE(INDIRECT("B" & ROW())), СПР_ПОДПИСАНТОВ!$B$14:$K$100, 7, 0), "")'
    formula_i = '=IF(ISNUMBER(VALUE(INDIRECT("B" & ROW()))), VLOOKUP(VALUE(INDIRECT("B" & ROW())), СПР_ПОДПИСАНТОВ!$B$14:$K$100, 5, 0), "")'

    final_row = find_last_row_in_col(sheet, col_index)

    if final_row:
        logging.info(f"The last non-empty cell in column {chr(64 + col_index)} of sheet '{sheet.title}' is in row {final_row}.")
    else:
        logging.info(f"No non-empty cells found in column {chr(64 + col_index)} of sheet '{sheet.title}'.")

    company = sheet['H11'].value # Get the company name from the sheet
    object_name = sheet['G11'].value # Get the object name from the sheet

    directors, coordinators_list = firmobj.check_company_object_pair(company, object_name) # Get the coordinators for the company and object

    n = len(coordinators_list) # Get the number of coordinators
    print(f'Company: {company}, Object: {object_name} Number of coordinators: {n}; coordinators: {coordinators_list}')
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

def loop_json(json_data, workbook):
    '''
    This function works with the loaded json and with the copied workbook
    '''
    cols = ['F', 'G', 'H', 'I']
    for key_title, data in json_data.items():

        print(len(data)) #how many documents are fetched

        for i in range(len(data)):
            if data[i]['object_name'] not in workbook.sheetnames:

                start_row = 17 # starting row for the writing
                row = start_row
                source_sheet = workbook['REESTR']

                # Create a copy of the source sheet with the desired name
                new_sheet = workbook.copy_worksheet(source_sheet)

                #listname feststelln
                new_sheet.title = data[i]['object_name']

                #objektname speichern
                object_name = data[i]['object_name']
                workbook[object_name][f'G11'] = object_name
                workbook[object_name][f'G10'] = datetime.today()
                workbook[object_name][f'F7'] = data[i]['registry_name']

                #Zeile formatieren
                format_row(workbook[object_name], row, cols)

                # On dr JSON datei bekommn

                workbook[object_name][f'H11'] = data[i]["organization"]

                sides_str = f'Заявитель: {data[i]["organization"]}'+'\n\n'+f'Кому: {data[i]["counteragent"]}'

                workbook[object_name][f'F{row}'] = sides_str

                workbook[object_name][f'G{row}'] = data[i]['zatraty'] # Zatraty po DDS

                workbook[object_name][f'H{row}'] = float(data[i]['payment_sum']) # Gebuhr

                i_cell_str = create_concatenated_info(data[i])

                workbook[object_name][f'I{row}'] = i_cell_str

                start_row += 1
            else:
                # test
                row = start_row
                object_name = data[i]['object_name']

                #Zeile formatieren
                format_row(workbook[object_name], row, cols)

                sides_str = f'Заявитель: {data[i]["organization"]}'+'\n\n'+f'Кому: {data[i]["counteragent"]}'

                workbook[object_name][f'F{row}'] = sides_str

                workbook[object_name][f'G{row}'] = data[i]['zatraty'] # Zatraty po DDS

                workbook[object_name][f'H{row}'] = float(data[i]['payment_sum']) # Gebuhr

                #workbook[object_name][f'M{row}'] = data[i]['sluzhebnaja_zapiska'] # Objektname

                i_cell_str = create_concatenated_info(data[i])
                workbook[object_name][f'I{row}'] = i_cell_str

                start_row += 1

def format_excel_inner(json_data):
    logging.info('Opening template.xlsx')
    workbook = load_excel('template.xlsx')

    initial_sheets = ['REESTR', 'СПР_ОБЪЕКТОВ', 'СПР_ПОДПИСАНТОВ']

    logging.info('Reading JSON file')
    # json_data = read_json() # read json
    # print(json_data)

    print('Looping through JSON. Adding documents')
    loop_json(json_data, workbook)

    for sheet in workbook.sheetnames:
        if sheet not in initial_sheets:
            add_coordinators_v4(workbook[sheet])
            set_print_area(workbook[sheet])
            add_colontituls(workbook[sheet])
    hide_sheets(workbook, initial_sheets)

    return workbook
