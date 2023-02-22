# Модуль, который сканирует бланки в папке бланков и создает из них процедуры

import os, shutil
from openpyxl import load_workbook
from specparser import SpecParser
from aux_code import load_first_sheet


class BlanksParser:
    def __init__(self, xl_path, pp_header):
        self.path = xl_path
        self.sheet = self._load_first_sheet()
        if not self.is_pp(pp_header):
            raise FileNotFoundError('Данный файл не является файлом ПП-бланка:', self.path)

    def _load_first_sheet(self):
        return load_first_sheet(self.path)

    # Метод, который анализирует xlsx-файл по данному пути и возвращает dict c парами заголовок-фильтр
    def get_filters(self):
        forming_dict = dict()
        for j in range(1, self.sheet.max_column):
            header = self.sheet.cell(row=3, column=j).value
            subfilter = self.sheet.cell(row=4, column=j).value
            if header:                              # Чтобы None не превращалось в 'none'
                header = str(header).lower()
            if subfilter:                           # Чтобы None не превращалось в 'none'
                subfilter = str(subfilter).lower()
            forming_dict.update({header: subfilter})
        if not forming_dict:
            raise FileNotFoundError('Ошибка формирования словаря фильтров: ' + str(path))
        return forming_dict

    # Проверяет xlsx-файлы на соответствие критериям
    # Сейчас это: первая ячейка первого листа должна содержать XL_HEADER
    def is_pp(self, pp_header):
        if self.sheet.cell(row=1, column=1).value == pp_header:
            return True
        return None


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
                    parser = BlanksParser(joint_path, self.PP_HEADER)
                    if parser.is_pp(self.PP_HEADER):
                        blanks_pathlist.append(joint_path)
        if len(blanks_pathlist) == 0:
            raise FileNotFoundError("Файлов бланков не найдено!")
        return blanks_pathlist

    def copy_blank(self, path_source: str, path_destination: str):
        shutil.copyfile(path_source, path_destination)

    def fill_blank(self, path):
        # parser =
        pass

    def form_blank(self, path_source, path_destination):
        pass  # копировать шаблон
        pass  # наполнить его


class BlankFiller:
    def __init__(self, xl_path):
        load

    pass


# Для тестов
if __name__ == '__main__':
    from configparser import ConfigParser

    configur = ConfigParser()
    configur.read('config.ini')
    configur_now = {'blanks_path': configur.get("paths", "blanks_path"),
                    'PP_HEADER': configur.get("variables", "PP_HEADER")}
    bhandler = BlanksHandler(configur_now)
    bparser = BlanksParser(bhandler.get_blanks_paths()[0], configur_now['PP_HEADER'])
    print(bparser.get_filters())
