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
    def __init__(self):
        self._params = dict()
        self._bhandler = None

    def _main_rcolumn_lout(self):
        out_list = [
            [sg.Text("Выбери бланки для заполнения из стандартных", font=_H2_FONT)]
        ]
        i = 0
        for entry in self._bhandler.get_blanks_names():
            out_list.append([sg.Checkbox(entry, key=entry, font=_SMALL_FONT)])
            i += 1
        out_list.append(
            [sg.Checkbox("Кастомный бланк:", key="Кастомный бланк:", font=_SMALL_FONT),
             sg.In(key='-CUSTOMBLANK-', expand_x=True, tooltip=_BLANKTIP, font=_SMALL_FONT),
             sg.FileBrowse(key='-SBrowse-')]
        )
        out_list.append([sg.VPush()])
        return out_list

    def _params_list_layout(self):
        out_list = list()
        for label, value in self._params.items():
            if 'MASK' not in label:
                out_list.append([sg.Text(label), sg.In(value, key=label)])
        return out_list

    def _masks_layout(self):
        out_list = list()
        for label, value in self._params['MASKS'].items():
            out_list.append([sg.Text(label), sg.In(value, key=label)])
        return out_list

    def get_main_layout(self):
        return [
            [sg.VPush()],
            [
                sg.Push(),
                sg.Column(
                    [
                        [sg.Push(), sg.Text("Welcome to PartParser!", font=_H1_FONT), sg.Push()],
                        [sg.Push(), sg.Text("Формирование заказа", font=_H2_FONT), sg.Push()],
                        [sg.HSeparator()],
                        [sg.Text('Введи наименование изделия', font=_H2_FONT)],
                        [sg.In(key='-PRODUCTNAME-', expand_x=True, tooltip=_SPECTIP)],
                        [sg.HSeparator()],
                        [sg.Text("Выбери эксель-файл спецификации", font=_H2_FONT)],
                        [sg.In(key='-XLSX-', expand_x=True, tooltip=_SPECTIP), sg.FileBrowse(key='-XBrowse-')],
                        [sg.HSeparator()],
                        [sg.Text('Данные для заполнения', font=_H2_FONT)],
                        [
                            sg.Text('Путь к документации', font=_SMALL_FONT),
                            sg.In(key='-DOCSFOLDER-', expand_x=True, tooltip=_BLANKTIP, font=_SMALL_FONT),
                            sg.FolderBrowse()
                        ],
                        [sg.HSeparator()],
                        [sg.Text("Сохранение результата", font=_H2_FONT)],
                        [sg.Checkbox('Сохранить в подпапке "/дата-время"', key='-DEFSAVE-', enable_events=True, default=True)],
                        [sg.Text("Или выбери другое место")],
                        [
                            sg.In(enable_events=True, key="-SAVEFOLDER-", expand_x=True, disabled=True),
                            sg.FolderBrowse(key='-FBrowse-', disabled=True)
                        ]
                    ]
                ),
                sg.VSeperator(),
                sg.Column(self._main_rcolumn_lout()),
                sg.Push()
            ],
            [sg.Push(), sg.Button("Парсить"), sg.Button('Config')],
            [sg.VPush()]
        ]

    def get_config_layout(self):
        return [
            [sg.VPush()],
            [sg.Push(), sg.Text('Настройки PartParser', font=_H1_FONT), sg.Push()],
            [sg.VPush()],
            [sg.Text('Параметры')],
            [sg.Push(), *self._params_list_layout(), sg.Push()],
            [sg.HSeparator()],
            [sg.Text('Маски')],
            [sg.Push(), *self._masks_layout(), sg.Push()],
            [sg.VPush()],
            [sg.Push(), sg.Button('Сохранить', key='-SAVECONFIG-'), sg.Button('На главную', key='-GOHOME-')]
        ]

    def get_window(self, winname: str, params: dict, bhandler: BlanksHandler):
        self._params = params
        self._bhandler = bhandler
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
            resizable=True,
            margins=(30, 30),
            element_justification=el_just,
            font=_DEFAULT_FONT,
            finalize=True
        )
