import PySimpleGUI as sg
import os.path
import PartParser
import sys
import GCPathHandler as ph
from configparser import ConfigParser
from blankshandler import BlanksHandler

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
    [sg.Text("Спецификация для парсинга", font=H2_FONT)],
    [sg.Text("Выбери Excel-файл"), sg.In(enable_events=True, key='-XLSX-'), sg.FileBrowse(key='XBrowse')],
    [sg.Text("Выбери тип спецификации"), sg.Listbox(SPEC_TYPES,
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
def touch_config(force_create=False):
    configur = ConfigParser()
    config_out = {}
    if force_create:
        try:
            sg.popup_ok('Настройка приложения.')
            gc_path = sg.popup_get_folder('Укажи корневую папку КД.')
            configur.add_section("paths")
            configur.set("paths", "grabcad_path", gc_path)
            blanks_path = sg.popup_get_folder('Укажи папку бланков.')
            configur.set("paths", "blanks_path", blanks_path)
            configur.add_section('common folders')
            try:
                i = 1
                while sg.popup_ok_cancel('Добавить общую папку ' + str(i) +'?') == 'OK':
                    configur.set('common folders', '_'.join(('folder', str(i))), sg.popup_get_folder('Укажи путь'))
                    i += 1
            except:
                sg.popup_ok('Добавление папки отменено')
            with open('config.ini', 'w') as conf:
                configur.write(conf)
            sg.popup_annoying('Программа сконфигурирована!')
        except:
            return None
    else:
        try:
            configur.read("config.ini")
            configur.get("paths", "blanks_path")
            configur.get("paths", "grabcad_path")
        except:
            return touch_config(force_create=True)
    return configur


if __name__ == '__main__':
    CONFIG = touch_config()
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
