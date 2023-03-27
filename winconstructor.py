import PySimpleGUI as sg
from blankshandler import BlanksHandler

_H1_FONT = ("Segoe UI", 22)
_H2_FONT = ("Segoe UI", 12)
_DEFAULT_FONT = ('Segoe UI', 10)
_SMALL_FONT = ("Segoe UI", 8)
_PAGE_1 = True
_SPECTIP = 'Спецификация должна быть экспортирована из Солида "смещенной" с "подробной нумерацией"'
_BLANKTIP = 'Бланк должен соответствовать правилам. См. файл "INFO"'


class WinConstructor:
    def __init__(self, bhandler: BlanksHandler):
        self._config_dict = dict()
        self.bhandler = bhandler

    def _b_cb_layout(self):
        out_list = list()
        i = 0
        for entry in self.bhandler.get_blanks_names():
            out_list.append([sg.Checkbox(entry, key=entry, font=_SMALL_FONT)])
            i += 1
        return out_list

    def _params_list_layout(self):
        in_list = [
            # Label  param.keyword    event_key
            ['Автор', '<AUTHOR>', '-CONFIG_AUTHOR-'],
            ['Путь к шаблонам', 'blanks_path', '-CONFIG_BLANKS_PATH-'],
            ['Сортировка 1', 'SPEC_SORT_1', '-CONFIG_SPEC_SORT_1-'],
            ['Сортировка 2', 'SPEC_SORT_2', '-CONFIG_SPEC_SORT_2-'],
        ]
        out_list = list()
        for label, mask, keyword in in_list:
            out_list.append([sg.Text(label), sg.In(self._config_dict.get(mask), key=keyword)])
        return out_list

    def get_main_layout(self):
        return [
            [sg.VPush()],
            [sg.Push(), sg.Text("Welcome to PartParser!", font=_H1_FONT), sg.Push()],
            [sg.Push(), sg.Text("Формирование заказа", font=_H2_FONT), sg.Push()],
            [sg.HSeparator()],
            [sg.Text('Введи наименование изделия', font=_H2_FONT)],
            [sg.In(key='-PROJECTNAME-', expand_x=True, tooltip=_SPECTIP)],
            [sg.HSeparator()],
            [sg.Text("Выбери эксель-файл спецификации", font=_H2_FONT)],
            [sg.In(key='-XLSX-', expand_x=True, tooltip=_SPECTIP), sg.FileBrowse(key='-XBrowse-')],
            [sg.HSeparator()],
            [sg.Text("Выбери бланки для заполнения из стандартных", font=_H2_FONT)],
            self._b_cb_layout(),
            [sg.Checkbox("Кастомный бланк:", font=_SMALL_FONT, key='-CUSTOMBLANK_CB-'),
             sg.In(key='-SPECSAVE-', expand_x=True, tooltip=_BLANKTIP, font=_SMALL_FONT),
             sg.FileBrowse(key='-SBrowse-')],
            [sg.HSeparator()],
            [sg.Text('Данные для заполнения', font=_H2_FONT)],
            [sg.Text('Путь к документации', font=_SMALL_FONT),
             sg.In(key='-SPECSAVE-', expand_x=True, tooltip=_BLANKTIP, font=_SMALL_FONT),
             sg.FileBrowse()
             ],
            [sg.HSeparator()],
            [sg.Text("Сохранение результата", font=_H2_FONT)],
            [sg.Checkbox('Сохранить в подпапке "/дата-время"', key='-DEFSAVE-', enable_events=True, default=True)],
            [sg.Text("Или выбери другое место")],
            [sg.In(enable_events=True, key="-SaveFolder-", expand_x=True, disabled=True),
             sg.FolderBrowse(key='-FBrowse-', disabled=True)],
            [sg.Push(), sg.Button("Парсить"), sg.Push(), sg.Button('config')],
            [sg.VPush()]
        ]

    def get_config_layout(self):
        return [
            [sg.VPush()],
            [sg.Push(), sg.Text('Настройки PartParser', font=_H1_FONT), sg.Push()],
            [sg.VPush()],
            self._params_list_layout(),
            [sg.VPush()],
            [sg.Push(), sg.Button('Сохранить', key='-SAVECONFIG-'), sg.Button('На главную', key='-GOHOME-')]
        ]

    def get_window(self, winname: str, config_dict: dict):
        self._config_dict = config_dict
        layout = []
        if winname == 'main':
            layout = self.get_main_layout()
            el_just = 'left'
        elif winname == 'config':
            layout = self.get_config_layout()
            el_just = 'right'
        return sg.Window(
            "PartParserComposer",
            layout,
            # size=(600, 800),
            resizable=True,
            margins=(30, 30),
            element_justification=el_just,
            font=_DEFAULT_FONT,
            finalize=True
        )
