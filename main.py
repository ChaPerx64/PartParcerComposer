import datetime

import PySimpleGUI as sg

from blankshandler import BlanksHandler
from confighandler import touch_config
from guireactor import GUIReactor
from winconstructor import WinConstructor

SPEC_TYPES = ("Плоская спецификация",)

# Create layout


CONFIG = touch_config()
config_dict = {'blanks_path': CONFIG.get("paths", "blanks_path"),
               'PP_HEADER': CONFIG.get("markers", "PP_HEADER"),
               '<AUTHOR>': CONFIG.get("userdata", "<AUTHOR>"),
               '<DOCS_FOLDER>': CONFIG.get("common folders", "folder_2"),
               '<DATE_OF_ISSUE>': str(datetime.date.today()),
               '<OTHER>': CONFIG.get('userdata', '<OTHER>'),
               'NUMBER_MARKER': CONFIG.get('markers', 'NUMBER_MARKER'),
               'POSITION_MARKER': CONFIG.get('markers', 'POSITION_MARKER'),
               'QUANTITY_MARKER': CONFIG.get('markers', 'QUANTITY_MARKER'),
               'SPEC_SORT_1': CONFIG.get('SPEC sorting order', 'SPEC_SORT_1'),
               'SPEC_SORT_2': CONFIG.get('SPEC sorting order', 'SPEC_SORT_2')
               }


# Load or create the configuration file
if __name__ == '__main__':
    reactor = GUIReactor()
    bhandler = BlanksHandler(config_dict)
    wcons = WinConstructor(bhandler)
    newwindow = 'main'
    window = None
    while True:
        if config_dict is None:
            sg.popup_error('Config не выполнен!')
            break
        if newwindow:
            window = wcons.get_window(newwindow, config_dict)
            newwindow = None
        event, values = window.read(timeout=20)
        br_flag, newwindow = reactor.react(event, values, window, config_dict)
        if br_flag:
            break
        if newwindow:
            window.close()
    window.close()
