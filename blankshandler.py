# Модуль, который сканирует бланки в папке бланков и создает из них процедуры

import os
from openpyxl import load_workbook


# Класс BlankHandler, занимающийся сбором и анализом xlsx-бланков заказов
class BlanksHandler:
    # При инициализации объекта класса, нужно скормить ему объект класса ConfigParser модуля сonfigparser,
    # инициализированный через ConfigParser.read('path')
    def __init__(self, params: dict):
        self.search_range = params.get("blanks_path")
        self.PP_HEADER = params.get("PP_HEADER")

    # Ищет файлы бланков в директории, прописанной в конфиге
    def get_blanks_paths(self):
        blanks_pathlist = []
        with os.scandir(path=self.search_range) as directory:
            for entry in directory:
                if ".xlsx" in entry.name and not '~$' in entry.name:
                    joint_path = '/'.join((self.search_range, entry.name))
                    if self._pp_check(joint_path):
                        blanks_pathlist.append(joint_path)
        if len(blanks_pathlist) == 0:
            raise FileNotFoundError("Файлов бланков не найдено!")
        return blanks_pathlist

    # Проверяет xlsx-файлы на соответствие критериям
    # Сейчас это: первая ячейка первого листа должна содержать XL_HEADER
    def _pp_check(self, path):
        xl_workbook = load_workbook(path, read_only=True)
        if xl_workbook[xl_workbook.sheetnames[0]]['A1'].value == self.PP_HEADER:
            return True
        return None

    # Метод, который анализирует xlsx-файл по данному пути и возвращает dict c парами заголовок-фильтр
    @staticmethod
    def analyze_blank(path):
        try:
            xl_workbook = load_workbook(path, read_only=True)
        except Exception:
            raise FileNotFoundError('Файл ' + path + 'не существует!')
        active_sheet = xl_workbook[xl_workbook.sheetnames[0]]
        headers_list = list()
        filters_list = list()
        forming_dict = dict()
        count = 1
        for row in active_sheet.rows:
            if count == 3:  # Заголовки в третьей строке
                for cell in row:
                    if cell.value is None:
                        headers_list.append(None)
                        # Отдельная обработка None, потому что оно конвертируется в строку 'none' командой ниже
                    else:
                        headers_list.append(str(cell.value).lower())
            if count == 4:  # Фильтры в четвертой
                for cell in row:
                    if cell.value is None:
                        filters_list.append(None)  # См. выше
                    else:
                        filters_list.append(str(cell.value).lower())
                break
            count += 1
        for header, subfilter in zip(headers_list, filters_list):
            forming_dict.update({header: subfilter})
        if not forming_dict:
            raise FileNotFoundError('Ошибка формирования словаря фильтров: ' + str(path))
        return forming_dict


# Для тестов
from configparser import ConfigParser

configur = ConfigParser()
configur.read('config.ini')
configur_now = {'blanks_path': configur.get("paths", "blanks_path"),
                'PP_HEADER': configur.get("variables", "PP_HEADER")}
bhandler = BlanksHandler(configur_now)
xl_path = bhandler.get_blanks_paths()[0]
print(xl_path)
print(bhandler.analyze_blank(xl_path))
