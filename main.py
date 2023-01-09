import PySimpleGUI as sg
import os.path
import PartParser
import sys
import GCPathHandler as ph
from configparser import ConfigParser

# Load or create the configuration file
CONFIG = {}
try:
    configur = ConfigParser()
    configur.read("config.ini")
    CONFIG.update({"GRABCAD_PATH": configur.get("paths", "grabcad_path")})
    # GRABCAD_PATH = )
except:
    configur = None


# Constants for style)
H1_FONT = ("Segoe UI", 22)
H2_FONT = ("Segoe UI", 18)
DEFAULT_FONT = ('Segoe UI', 11)
SMALL_FONT = ("Segoe UI", 8)
SPEC_TYPES = ("Плоская спецификация",)

# Create layout
layout = [
    [sg.Push(), sg.Text("Welcome to PartParser!", font=H1_FONT), sg.Push()],
    [sg.HSeparator()],
    [sg.Text("Спецификация для парсинга", font=H2_FONT)],
    [sg.Text("Выбери Excel-файл"), sg.In(enable_events=True, key='-XLSX-'), sg.FileBrowse(key='XBrowse')],
    [sg.Text("Выбери тип спецификации"), sg.Listbox(SPEC_TYPES,
                                             no_scrollbar=True,
                                             size=(60, 2),
                                             font=SMALL_FONT,
                                             select_mode="LISTBOX_SELECT_MODE_SINGLE",
                                             )],
    [sg.HSeparator()],
    [sg.Text("Область поиска", font=H2_FONT)],
    [sg.Text("Папка поиска"), sg.In(enable_events=True,
                                    key="-GCFolder-",
                                    ), sg.FolderBrowse("Изменить", key='FBrowse')],
    [sg.Checkbox("Искать в 'Sorters'"),
     sg.Checkbox("Искать в 'Design Standarts'"),
     sg.Checkbox("Искать по всему 'GrabCAD'")],
    [sg.HSeparator()],
    [sg.Text("Параметры поиска", font=H2_FONT)],
    [sg.Checkbox("Пробовать поиск групповых чертежей", default=True), sg.Checkbox("Выводить log-файл", default=True)],
    [sg.HSeparator()],
    [sg.Push(), sg.Button("Искать"), sg.Push()]
]

# create window
window = sg.Window(
    "PartParserComposer",
    layout,
    size=(700, 500),
    element_justification="left",
    font=DEFAULT_FONT,
    finalize=True
)


def create_config(config):
    configur = ConfigParser()
    gc_path = sg.popup_get_folder('Config-файл не найден. Необходима настройка. Укажи корневую папку GrabCAD.')
    configur.add_section("paths")
    configur.set("paths", "grabcad_path", gc_path)
    with open('config.ini', 'w') as conf:
        configur.write(conf)


if __name__ == '__main__':
    if configur is None:
        create_config(CONFIG)
    # window['-GCFolder-'].update(CONFIG['GRABCAD_PATH'])
    while True:
        event, values = window.read()
        if event in ("Exit", sg.WIN_CLOSED):
            break
        # window['-GCFolder-'].update(CONFIG['GRABCAD_PATH'])
        # if event is :
        # window[]
    window.close()
    # path = GCPathHandler.CDPath(CONFIG['GRABCAD_PATH'])
    # print(GCPathHandler.get_project_pathway(CONFIG['GRABCAD_PATH']))
