# Модуль, который сканирует бланки в папке бланков и создает из них процедуры

import datetime
import os
import time

from blanksparser import BlanksParser
from specparser import SpecParser
from entrymatcher import EntryMatcher
from aux_code import filename_from_path


class BlanksHandler:
    """
    Класс BlankHandler, занимающийся сбором и анализом xlsx-бланков заказов
    """

    def __init__(self, params: dict):
        """
        При инициализации объекта класса, нужно скормить ему объект класса ConfigParser модуля сonfigparser,
        инициализированный через ConfigParser.read('path')
        """
        self.params = params
        self.found_blanks = self.get_blanks_paths()
        self.blanks_names = self.get_blanks_names()

    def get_blanks_paths(self):
        """
        Ищет файлы бланков в директории, прописанной в конфиге
        """
        blanks_pathlist = []
        with os.scandir(path=self.search_range) as directory:
            for entry in directory:
                if ".xlsx" in entry.name and not '~$' in entry.name:
                    joint_path = '/'.join((self.search_range, entry.name))
                    parser = BlanksParser(joint_path, self.params)
                    if parser.is_pp():
                        blanks_pathlist.append(joint_path)
        if len(blanks_pathlist) == 0:
            raise FileNotFoundError("Файлов бланков не найдено!")
        return blanks_pathlist

    def get_blanks_names(self):
        out_ls = list()
        for item in self.found_blanks:
            out_ls.append(filename_from_path(item))
        return out_ls

    def get_copy_wbws(self, source_path):
        bparser = BlanksParser(source_path, self.params)
        new_wb_ws = bparser.get_copy()
        return new_wb_ws

    def save_formatted_copy(self, source_path, path_destination: str):
        new_ws = self.get_copy_wbws(source_path)
        new_ws = self.format_sheet(new_ws)
        new_ws.parent.save(path_destination)

    def format_sheet(self, sheet):
        sheet.delete_rows(1)
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=sheet.max_column)
        for row in sheet:
            for cell in row:
                if cell.value:
                    for key, value in self.params.items():
                        cell.value = str(cell.value).replace(key, value)
        return sheet

    def fill_blank(self, b_path, spec_path):
        sparser = SpecParser(spec_path, self.params)
        bparser = BlanksParser(b_path, self.params)
        order_sheet = bparser.get_copy()
        matcher = EntryMatcher(bparser.get_filters(), sparser.get_flat_unformatted(), self.params)
        if not matcher.errors:
            i = 4
            for entry in matcher.get_matches():
                order_sheet.cell(row=i, column=1).value = i - 3
                j = 2
                for svalue in entry.values():
                    order_sheet.cell(row=i, column=j).value = svalue
                    j += 1
                i += 1
            return order_sheet
        else:
            out_list = list()
            for item in matcher.errors:
                out_list.append({self.params.get('KEY_1'): item})
            return out_list

    def form_order_fromind(self, list_no: int, spec_path, path_destination):
        order_sheet = self.fill_blank(self.found_blanks[list_no], spec_path)
        order_sheet = self.format_sheet(order_sheet)
        order_sheet.parent.save(path_destination)

    def form_order_fromname(self, blankname, spec_path, path_destination):
        if blankname in self.blanks_names:
            return self.form_order_fromind(self.blanks_names.index(blankname), spec_path, path_destination)




# Для тестов
if __name__ == '__main__':
    from configparser import ConfigParser

    configur = ConfigParser()
    configur.read('config.ini')
    configur_now = {'blanks_path': configur.get("paths", "blanks_path"),
                    'PP_HEADER': configur.get("variables", "PP_HEADER"),
                    '<PRODUCT>': 'Стеллаж',
                    '<AUTHOR>': 'Пар-оол Ч.А.',
                    '<DOCS_FOLDER>': configur.get("common folders", "folder_2"),
                    '<DATE_OF_ISSUE>': str(datetime.date.today()),
                    '<OTHER>': '',
                    'PP_NUMBER': 'Номер',
                    'POSITION_MARKER': 'Поз.',
                    'QUANTITY_MARKER': 'Кол-во',
                    'KEY_1': 'Обозначение',
                    'KEY_2': 'Наименование'
                    }
    bhandler = BlanksHandler(configur_now)
    index_yo = 5
    spec_path = './Пример спеки/Familia Rack.xlsx'
    new_ls = spec_path.split('.')
    new_ls[1] = new_ls[1] + ' ' + '_'.join(
        (str(time.gmtime().tm_hour), str(time.gmtime().tm_min), str(time.gmtime().tm_sec)))
    new_path = '.'.join(new_ls)
    bhandler.form_order_fromind(index_yo, spec_path, new_path)
