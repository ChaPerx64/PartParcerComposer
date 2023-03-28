import PySimpleGUI as sg
from blankshandler import BlanksHandler
from aux_code import curr_dt_subfolder, new_filepath
from confighandler import touch_config
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
            self.flip_save()
        if event == 'Парсить':
            err_message = self.parse_check()
            if not err_message:
                self.parse()
            else:
                sg.popup_error(err_message)
        if event == "config":
            self._newwindow = 'config'
        if event == "-GOHOME-":
            self._newwindow = 'main'
        if event in ("Exit", sg.WIN_CLOSED):
            self._br_flag = True
        return self._br_flag, self._newwindow

    def flip_save(self):
        self._savedisabled = not self._savedisabled
        self._window['-SAVEFOLDER-'].update(disabled=self._savedisabled)
        self._window['-FBrowse-'].update(disabled=self._savedisabled)

    def parse_check(self):
        """Функция, проверяющая ввод данных пользователем"""

        '''Проверка ввода данных, используемых для масок'''
        if self._values['-PRODUCTNAME-']:
            self._params.update({'<PRODUCTNAME>': self._values['-PRODUCTNAME-']})
        else:
            return 'Введи наименование изделия'

        '''Проверка ввода пути исходной спеки'''
        if self._values['-XLSX-']:
            self._params.update({'XLSXPATH': self._values['-XLSX-']})
        else:
            return 'Укажи исходную спецификацию'

        '''Проверка ввода пути к модели, для маски в таблице'''
        if self._values.get('-CADPATH-'):
            self._params.update({'CADPATH': self._values['-CADPATH-']})
        else:
            return 'Не введён путь к документации (справочный)'

        '''Проверка врода данных о пути сохранения заказов'''
        if self._values['-DEFSAVE-']:
            self._params.update({'SAVEPATH': curr_dt_subfolder(self._values['-XLSX-'])})
        elif self._values['-SAVEFOLDER-']:
            self._params.update({'SAVEPATH': self._values['-SAVEFOLDER-']})
        else:
            return 'Выбери путь для сохранения'

        '''Проверка выбора путей для сохранения'''
        parse_list = list()
        self._bhandler = BlanksHandler(self._params)
        # Проверяет чекбоксы со спеками
        for blankname in self._bhandler.get_blanks_names():
            if self._values[blankname]:
                parse_list.append(blankname)
        if not (parse_list or self._values['Кастомный бланк:']):
            return 'Не выбран ни один бланк для заполнения'
        if self._values['Кастомный бланк:']:
            parse_list.append(self._values['-CUSTOMBLANK-'])
        if parse_list:
            self._params.update({'PARSELIST': parse_list})

        '''Возвращает None, если все условия выше не сработали (Ошибки не обнаружены)'''
        return None

    def parse(self):
        status_string = 'Результаты парсинга:\n'
        os.makedirs(self._params.get('SAVEPATH'), exist_ok=True)
        for blank in self._params.get('PARSELIST'):
            status_string += '\n' + str(blank) + ':\n'
            new_path = new_filepath(self._params.get('<PRODUCTNAME>'), self._params.get('SAVEPATH'), blank)
            out = self._bhandler.form_order_fromname(blank, self._params.get('XLSXPATH'), new_path)
            if out:
                for item in out:
                    status_string += str(item) + '\n'
            else:
                status_string += 'Спарсено успешно!\n'
        sg.popup_ok(status_string)
        self._br_flag = True

    def t_config(self):
        self._newwindow = 'config'
