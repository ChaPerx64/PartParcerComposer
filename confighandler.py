from configparser import ConfigParser
import PySimpleGUI as sg

class ConfigHandler:
    def __init__(self, path: str):
        self.path = path


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
                while sg.popup_ok_cancel('Добавить общую папку ' + str(i) + '?') == 'OK':
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