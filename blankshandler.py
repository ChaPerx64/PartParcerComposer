# Модуль, который сканирует бланки в папке бланков и создает из них процедуры

import os
from openpyxl import load_workbook

# from configparser import ConfigParser
XL_HEADER = 'PP_BLANK'


class BlanksHandler:
    def __init__(self, configur):
        self.search_range = configur.get("paths", "blanks_path")

    # Ищет файлы бланков в директории, прописанной в конфиге
    def get_blanks_paths(self):
        blanks_pathlist = []
        with os.scandir(path=self.search_range) as center:
            for entry in center:
                if ".xlsx" in entry.name and not '~$' in entry.name:
                    joint_path = '/'.join((self.search_range, entry.name))
                    if self.pp_checker(joint_path):
                        blanks_pathlist.append(joint_path)
        if len(blanks_pathlist) == 0:
            raise FileNotFoundError("Файлов бланков не найдено!")
        return blanks_pathlist

    # Проверяет xlsx-файлы на соответствие критериям
    # Сейчас это: первая ячейка первого листа должна содержать XL_HEADER
    @staticmethod
    def pp_checker(path):
        xl_workbook = load_workbook(path, read_only=True)
        if xl_workbook[xl_workbook.sheetnames[0]]['A1'].value == XL_HEADER:
            return True
        return None

#
# print(xlsx_finder('./Бланки'))
