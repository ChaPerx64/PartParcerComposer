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
        self._params = params
        self.found_blanks = self.get_blanks_paths()
        self.blanks_names = self.get_blanks_names()

    def get_blanks_paths(self):
        """
        Ищет файлы бланков в директории, прописанной в конфиге
        """
        blanks_pathlist = []
        with os.scandir(path=self._params.get('BLANKS_PATH')) as directory:
            for entry in directory:
                if ".xlsx" in entry.name and not '~$' in entry.name:
                    joint_path = '/'.join((self._params.get('BLANKS_PATH'), entry.name))
                    parser = BlanksParser(joint_path, self._params)
                    if parser.is_pp():
                        blanks_pathlist.append(joint_path)
        if not blanks_pathlist:
            raise FileNotFoundError("Файлов бланков не найдено!")
        return blanks_pathlist

    def get_blanks_names(self):
        out_ls = list()
        for item in self.found_blanks:
            out_ls.append(filename_from_path(item))
        return out_ls

    def get_copy_wbws(self, source_path):
        bparser = BlanksParser(source_path, self._params)
        new_wb_ws = bparser.get_copy()
        return new_wb_ws

    def save_formatted_copy(self, source_path, path_destination: str):
        new_ws = self.get_copy_wbws(source_path)
        new_ws = self.format_sheet(new_ws)
        new_ws.parent.save(path_destination)

    def format_sheet(self, sheet):
        sheet.delete_rows(1)
        if str(self._params['MERGETOCOLUMN']).isdigit():
            merge_edge = int(self._params['MERGETOCOLUMN'])
            if merge_edge != 0:
                sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=merge_edge)
        else:
            sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=sheet.max_column)
        for row in sheet:
            for cell in row:
                if cell.value:
                    for key, value in self._params['MASKS_VALUES'].items():
                        if value:
                            cell.value = str(cell.value).replace(key, value)
                        else:
                            cell.value = str(cell.value).replace(key, '')
        return sheet

    def fill_blank(self, b_path, spec_path):
        sparser = SpecParser(spec_path, self._params)
        bparser = BlanksParser(b_path, self._params)
        order_sheet = bparser.get_copy()
        matcher = EntryMatcher(bparser.get_filters(), sparser.get_flat_unformatted(), self._params)
        if matcher.errors:
            return matcher.errors
        else:
            i = 4
            for entry in matcher.get_matches():
                order_sheet.cell(row=i, column=1).value = i - 3
                j = 2
                for svalue in entry.values():
                    order_sheet.cell(row=i, column=j).value = svalue
                    j += 1
                i += 1
            return order_sheet

    def form_order_fromind(self, list_no: int, spec_path, path_destination):
        order_sheet = self.fill_blank(self.found_blanks[list_no], spec_path)
        if isinstance(order_sheet, list):
            return order_sheet
        order_sheet = self.format_sheet(order_sheet)
        order_sheet.parent.save(path_destination)

    def form_order_fromname(self, blankname, spec_path, path_destination):
        if blankname in self.blanks_names:
            return self.form_order_fromind(self.blanks_names.index(blankname), spec_path, path_destination)
