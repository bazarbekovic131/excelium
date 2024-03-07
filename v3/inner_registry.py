from datetime import datetime
import logging
from shutil import copy2
from scripts import format_row, set_border, find_last_row_in_col, load_excel, hide_sheets, create_concatenated_info, set_print_area
import openpyxl
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
        # Calculate the current row
        row = final_row + i * 3 # it was i*2
        formula_b = f'=INDEX(СПР_ОБЪЕКТОВ!$B$7:$K$80, MATCH($G11, СПР_ОБЪЕКТОВ!$B$7:$B$80, 0), {3 + i})'

        # Add thin borders to cell B{row}
        sheet.cell(row=row, column=2, value=formula_b)
        sheet.cell(row=row, column=2).border = set_border('thin')

        # Add formula to cell F{row}
        sheet.cell(row=row, column=6, value=formula_f)
        sheet.cell(row=row, column=6).alignment = Alignment(horizontal='left')
        sheet.cell(row=row, column=6).font = Font(size=14, bold=True)

        # Add formula to cell I{row}
        sheet.cell(row=row, column=9, value=formula_i)
        sheet.cell(row=row, column=6).alignment = Alignment(horizontal='left')
        sheet.cell(row=row, column=6).font = Font(bold=True, size = 14)

    # Add the word "СОГЛАСОВАНО" in cell F{row + 2}
    last_row = final_row + 16  # Assuming the last row in the cycle is 16 rows from the starting row
    sheet.cell(row=last_row+1, column=6, value="СОГЛАСОВАНО")
    sheet.cell(row=last_row+1, column=6).font = Font(bold=True)

    sheet.cell(row=last_row + 3, column=2, value=3)
    sheet.cell(row=last_row + 3, column=2).border = set_border('thin')

    sheet.cell(row=last_row + 3, column=6, value=formula_f)
    sheet.cell(row=last_row + 3, column=6).alignment = Alignment(horizontal='left')
    sheet.cell(row=last_row + 3, column=6).font = Font(bold = True, size = 14)

    sheet.cell(row=last_row + 3, column=9, value=formula_i)
    sheet.cell(row=last_row + 3, column=9).alignment = Alignment(horizontal='left')
    sheet.cell(row=last_row + 3, column=9).font = Font(bold = True, size = 14)

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
            add_coordinators(workbook[sheet])
            set_print_area(workbook[sheet])
    hide_sheets(workbook, initial_sheets)

    return workbook
