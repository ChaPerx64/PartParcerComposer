import PySimpleGUI as sg
from blankshandler import BlanksHandler
from aux_code import curr_dt_subfolder, new_filepath, filename_from_path
from cadpathhandler import CADPathHandler
from confighandler import config_save, touch_config
import os


class GUIReactor:
    def __init__(self):
        self._wincons = None
        self._event = None
        self._values = dict()
        self._window = None
        self._bhandler = None
        self._params = dict()
        self._br_flag = None
        self._newwindow = None
        self._savedisabled = True

    def react(self, event: str, values: dict, window: sg.Window, params: dict):
        self._event = event
        self._values = values
        self._window = window
        self._params = params
        self._br_flag = None
        self._newwindow = None
        if event == '-DEFSAVE-':
            self._flip_save()
        if event == 'Парсить':
            err_message = self._parse_check()
            if not err_message:
                self._parse()
            else:
                sg.popup_error(err_message)
        if event == "Config":
            self._newwindow = 'config'
        if event == "-GOHOME-":
            self._newwindow = 'main'
        if event in ("Exit", sg.WIN_CLOSED):
            self._br_flag = True
        if event == "-SAVECONFIG-":
            self._save_config()
        if event == 'Help':
            self._newwindow = 'Help'
        if event == '-CONFIG_DEFAULT-':
            self._config_default()
            self._newwindow = 'config'
        if event == '-HELP_CONFIG-':
            self._newwindow = 'config_help'
        return self._br_flag, self._newwindow

    def _flip_save(self):
        self._savedisabled = not self._savedisabled
        self._window['-SAVEFOLDER-'].update(disabled=self._savedisabled)
        self._window['-FBrowse-'].update(disabled=self._savedisabled)

    def _parse_check(self):
        """Функция, проверяющая ввод данных пользователем"""

        chandler = CADPathHandler(self._params.get('CADFOLDER_PATH'))
        masks = self._params['MASKS']
        masks_values = self._params['MASKS_VALUES']

        '''Проверка ввода данных, используемых для масок'''
        if self._values['-PRODUCTNAME-']:
            masks_values.update({masks['PRODUCTNAME_MASK']: self._values['-PRODUCTNAME-']})
        else:
            return 'Введи наименование изделия'

        '''Проверка ввода пути исходной спеки'''
        if self._values['-XLSX-']:
            self._params.update({'XLSXPATH': self._values['-XLSX-']})
        else:
            return 'Укажи исходную спецификацию'

        '''Проверка ввода пути к модели, для маски в таблице'''
        if self._values['-DOCSFOLDER-']:
            if self._params['CADFOLDER_PATH'] in self._values['-DOCSFOLDER-']:
                masks_values.update({masks['DOCSFOLDER_MASK']: chandler.strip_to_cadfolder(self._values['-DOCSFOLDER-'])})
            else:
                masks_values.update({masks['DOCSFOLDER_MASK']: self._values['-DOCSFOLDER-']})
        else:
            return 'Не введён путь к документации (справочный)'

        '''Проверка врода данных о пути сохранения заказов'''
        if self._values['-DEFSAVE-']:
            self._params.update({'SAVEPATH': curr_dt_subfolder(self._values['-XLSX-'])})
        elif self._values['-SAVEFOLDER-']:
            self._params.update({'SAVEPATH': str(self._values['-SAVEFOLDER-']) + '/'})
        else:
            return 'Выбери путь для сохранения'

        '''Проверка выбора бланками для сохранения'''
        parse_list = list()
        self._bhandler = BlanksHandler(self._params)
        # Проверяет чекбоксы со спеками
        for blankname in self._bhandler.get_blanks_names():
            if self._values[blankname]:
                parse_list.append(blankname)
        if not (parse_list or self._values['Кастомный бланк:']):
            return 'Не выбран ни один бланк для заполнения'
        if self._values['Кастомный бланк:']:
            self._params.update({'CUSTOMBLANK': self._values['-CUSTOMBLANK-']})
        if parse_list:
            self._params.update({'PARSELIST': parse_list})

        if not self._params.get('AUTHOR'):
            return 'Не настроено имя автора (Параметр "AUTHOR")'

        if not self._params.get('CADFOLDER_PATH'):
            return 'Не настроен путь к хранилищу КД (Параметр "CADFOLDER_PATH")'

        '''Возвращает None, если все условия выше не сработали (Ошибки не обнаружены)'''
        return None

    def _parse(self):
        status_string = 'Результаты парсинга:\n'
        os.makedirs(self._params.get('SAVEPATH'), exist_ok=True)
        if self._params.get('PARSELIST'):
            for blank in self._params.get('PARSELIST'):
                status_string += '\n' + str(blank) + ':\n'
                new_path = new_filepath(
                    self._params['MASKS_VALUES'][self._params.get('MASKS').get('PRODUCTNAME_MASK')],
                    self._params.get('SAVEPATH'),
                    blank
                )
                out = self._bhandler.form_order_fromname(blank, self._params.get('XLSXPATH'), new_path)
                if out:
                    for item in out:
                        status_string += str(item) + '\n'
                else:
                    status_string += 'Спарсено успешно!\n'
        if self._params.get('CUSTOMBLANK'):
            status_string += '\nКастомный бланк:\n'
            new_path = new_filepath(
                self._params['MASKS_VALUES'][self._params.get('MASKS').get('PRODUCTNAME_MASK')],
                self._params.get('SAVEPATH'),
                filename_from_path(self._params.get('CUSTOMBLANK'))
            )
            out = self._bhandler.form_order_from_path(
                self._params.get('CUSTOMBLANK'), self._params.get('XLSXPATH'), new_path
            )
            if out:
                for item in out:
                    status_string += str(item) + '\n'
            else:
                status_string += 'Спарсено успешно!\n'
        sg.popup_ok(status_string)
        self._br_flag = True

    def _save_config(self):
        for label, values in self._params.items():
            if 'MASKS' not in label:
                self._params.update({label: self._values[label]})
        for label, values in self._params.get('MASKS').items():
            self._params['MASKS'].update({label: self._values[label]})
        config_save(self._params)
        sg.popup_ok('Конфиг сохранён')

    @staticmethod
    def _config_default():
        touch_config(force_create=True)
