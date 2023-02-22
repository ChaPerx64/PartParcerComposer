# Модуль с допфункциями, помогающими работать с эксель
from openpyxl import load_workbook


def load_first_sheet(xl_path, remove_enters=True):
    try:
        xl_workbook = load_workbook(xl_path, read_only=False, data_only=True)
    except:
        raise RuntimeError('Ошибка чтения файла:', xl_path)
    active_sheet = xl_workbook[xl_workbook.sheetnames[0]]
    for row in active_sheet.rows:
        for cell in row:
            if remove_enters:
                if cell.value:
                    new_value = str(cell.value).replace('\n', '')
                    cell.value = new_value
    return active_sheet
