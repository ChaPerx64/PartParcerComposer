import PySimpleGUI as sg
import os.path
import PartParser
import sys
import GCPathHandler as ph
from blankshandler import BlanksHandler
from confighandler import ConfigHandler, touch_config

# Constants for style)
H1_FONT = ("Segoe UI", 22)
H2_FONT = ("Segoe UI", 16)
DEFAULT_FONT = ('Segoe UI', 10)
SMALL_FONT = ("Segoe UI", 8)
SPEC_TYPES = ("Плоская спецификация",)


# Create layout
layout = [
    [sg.Push(), sg.Text("Welcome to PartParser!", font=H1_FONT), sg.Push()],
    [sg.HSeparator()],
    [sg.Text("Формирование заказа", font=H2_FONT)],
    [sg.Text("Спецификация сборки"), sg.In(enable_events=True, key='-XLSX-'), sg.FileBrowse(key='XBrowse')],
    [sg.Text("Бланки спецификации"), sg.Listbox(SPEC_TYPES,
                                                no_scrollbar=True,
                                                size=(60, len(SPEC_TYPES)),
                                                font=SMALL_FONT,
                                                select_mode="LISTBOX_SELECT_MODE_SINGLE",
                                                )],
    [sg.HSeparator()],
    [sg.Text("Область поиска", font=H2_FONT)],
    [sg.Text("Папка поиска"), sg.In(enable_events=True,
                                    key="-GCFolder-",
                                    ), sg.FolderBrowse("Изменить", key='FBrowse')],
    [sg.Checkbox("Искать в общих папках (из config.ini)"),
     sg.Checkbox("Искать в 'Design Standarts'"),
     sg.Checkbox("Искать по всему 'GrabCAD'")],
    [sg.HSeparator()],
    [sg.Text("Параметры поиска", font=H2_FONT)],
    [sg.Checkbox("Пробовать поиск групповых чертежей", default=True), sg.Checkbox("Выводить log-файл", default=True)],
    [sg.HSeparator()],
    [sg.Push(), sg.Button("Искать"), sg.Push(), sg.Button('config')]
]

# create window
window = sg.Window(
    "PartParserComposer",
    layout,
    size=(600, 600),
    element_justification="left",
    font=DEFAULT_FONT,
    finalize=True
)


# Load or create the configuration file



if __name__ == '__main__':
    CONFIG = touch_config()
    print(CONFIG)
    # bhandler = BlanksHandler(CONFIG)
    # print(bhandler.get_blanks_paths())
    # window['-GCFolder-'].update(CONFIG['GRABCAD_PATH'])
    while True:
        if CONFIG is None:
            sg.popup_error('Config не выполнен!')
            break
        event, values = window.read()
        if event in ("Exit", sg.WIN_CLOSED):
            break
        if event == "config":
            config_temp = touch_config(True)
            if config_temp is not None:
                CONFIG = config_temp
        # window['-GCFolder-'].update(CONFIG['GRABCAD_PATH'])
        if event == 'XBrowse':
            print('boo')
            print(values)
            # window['-XLSX-'].update(ph.get_project_pathway())
    window.close()
    # path = GCPathHandler.CDPath(CONFIG['GRABCAD_PATH'])
    # print(GCPathHandler.get_project_pathway(CONFIG['GRABCAD_PATH']))
