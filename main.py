import PySimpleGUI as sg
import os.path
import datetime
import PartParser
import sys
import GCPathHandler as ph
from blankshandler import BlanksHandler
from confighandler import ConfigHandler, touch_config
from winconstructor import WinConstructor
from aux_code import curr_dt_subfolder, new_filepath
from guireactor import GUIReactor

SPEC_TYPES = ("Плоская спецификация",)

# Create layout


CONFIG = touch_config()
config_dict = {'blanks_path': CONFIG.get("paths", "blanks_path"),
               'PP_HEADER': CONFIG.get("variables", "PP_HEADER"),
               '<PRODUCT>': 'Стеллаж',
               '<AUTHOR>': CONFIG.get("userdata", "AUTHOR"),
               '<DOCS_FOLDER>': CONFIG.get("common folders", "folder_2"),
               '<DATE_OF_ISSUE>': str(datetime.date.today()),
               '<OTHER>': '',
               'PP_NUMBER': 'Номер',
               'POSITION_MARKER': 'Поз.',
               'QUANTITY_MARKER': 'Кол-во',
               'KEY_1': 'Обозначение',
               'KEY_2': 'Наименование'
               }


# Load or create the configuration file
if __name__ == '__main__':
    reactor = GUIReactor()
    WINCONSTRUCT = True
    while True:
        if config_dict is None:
            sg.popup_error('Config не выполнен!')
            break
        if WINCONSTRUCT:
            bhandler = BlanksHandler(config_dict)
            wcons = WinConstructor(bhandler)
            window = wcons.get_main_window()
            WINCONSTRUCT = False
        event, values = window.read()
        br_flag, con_flag = reactor.react(event, values, window, bhandler, config_dict)
        if con_flag:
            continue
        if br_flag:
            break
    window.close()
