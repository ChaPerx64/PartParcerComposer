import PySimpleGUI as sg
from blankshandler import BlanksHandler

_H1_FONT = ("Segoe UI", 22)
_H2_FONT = ("Segoe UI", 14)
_H3_FONT = ("Segoe UI", 12)
_DEFAULT_FONT = ('Segoe UI', 10)
_SMALL_FONT = ("Segoe UI", 8)
_PAGE_1 = True
_SPECTIP = 'Спецификация должна быть экспортирована из Солида "смещенной" с "подробной нумерацией"'
_BLANKTIP = 'Бланк должен соответствовать правилам. См. файл "INFO"'
_HELP_0 = '''SpecParser работает по следующему принципу:
Он читает выгруженную из SolidWorks сдвинутую спецификацию с подробной нумерацией формата .xlsx.
На основании этой спецификации заполняет "бланки" формата .xlsx.
Эти "бланки" являются шаблонами (пустыми таблицами) с некоторыми фильтрами (требованиями к объектам,
которые должны в него войти).
Заполненный бланк SpecParser сохраняет, как "заказ".
'''
_HELP_1 = '''Наименование изделия - имя, которое будет вставляться в заголовки спек заказов и в имя файла заказа.
Эксель-файл спецификации - спецификация, экспортированная из SW, сдвинутая, с подробной нумерацией.
Путь к документации - путь к моделям для справки (для исполнителей заказа).

Note 1:
SpecParser автоматически обрезает строку "Путь в документации" до корневой папки КД, указанной в конфиге,
если этот каталог является подпапкой каталога КД
Note 2:
Вставка в окошки путей напрямую запрещена, чтобы не вызывать конфликт из-за разных разделителей
подкаталогов. Вставлять скопированные пути можно в диалоговом окне 'Browse'.
'''
_HELP_2 = '''Здесь нужно выбрать место сохранения спек заказов.
По умолчанию выбрана опция: Сохранить рядом со спекой в подпапке "/дата-время".
С этой опцией заказы сохраняются в подпапку "Заказы ГГГГ_ММ_ДД__ЧЧ_ММ_СС".
Если галочку снять, можно будет выбрать произвольную папку для сохранения.
'''
_HELP_3 = '''Этот список формируется автомаатически на основании файлов бланков, найденных в каталоге бланков.
Предполагается, что этот каталог будет находиться на каком-то общем ресурсе (сетевом или синхронизируемом, как GrabCAD),
Поэтому здесь тажке присутсвует возможность выбрать кастомный бланк, находящийся в произвольном месте, на случай, если
ты пользуешься каким-то шаблоном, применимым только к твоему случаю.
'''
_CONFIG_HELP_1 = '''CADFOLDER_PATH - Путь к хранилищу КД. Пока что используется только для
вычисления строки "Путь к КД", вставляемой в заказ.

BLANKS_PATH - Путь к папке с бланками

PP_HEADER - Тег, по которому SpecParser определяет файл бланка (см. файл INFO в папке с бланками)

POSITION_MARKER - Заголовок столбца позиций в спеке SW
QUANTITY_MARKER - Заголовок столбца количества в cпеке SW
NUMBER_MARKER - Заголовок номеров позиций в спеке SW

Заказ сортируется по двум позициям в восходящем порядке по символам Unicode
(case-sensitive, заглавные буквы впереди маленьких)
SPEC_SORT_1 - Заголовок столбца, у которого первый приоритет сортировки
SPEC_SORT_1 - Заголовок столбца, у которого второй приоритет сортировки
MERGETOCOLUMN - Число. Номер столбца, до которого надо объединить ячейки в шапке заказа
AUTHOR - Имя, которым будет подписана спека заказа
OTHER - Поле для ввода дополнительных данных в бланк с помощью маски <OTHER>
'''


class WinConstructor:
    def __init__(self):
        self._params = dict()
        self._bhandler = None


    @staticmethod
    def _main_lcolumn_lout():
        return [
            [sg.Text('Данные для заполнения', font=_H3_FONT)],
            [sg.Text('Введи наименование изделия')],
            [sg.In(key='-PRODUCTNAME-', expand_x=True, tooltip=_SPECTIP)],
            [sg.Text("Выбери эксель-файл спецификации")],
            [
                sg.In(key='-XLSX-', expand_x=True, tooltip=_SPECTIP, readonly=True),
                sg.FileBrowse(key='-XBrowse-')
            ],
            [sg.Text('Путь к документации')],
            [
                sg.In(key='-DOCSFOLDER-', expand_x=True, tooltip=_BLANKTIP, font=_SMALL_FONT, readonly=True),
                sg.FolderBrowse()
            ],
            [sg.HSeparator()],
            [sg.Text("Сохранение результата", font=_H3_FONT)],
            [
                sg.Checkbox(
                    'Сохранить рядом со спекой в подпапке "/дата-время"', key='-DEFSAVE-', enable_events=True,
                    default=True
                )
            ],
            [sg.Text("Или в другом месте", font=_SMALL_FONT)],
            [
                sg.In(enable_events=True, key="-SAVEFOLDER-", expand_x=True, disabled=True, readonly=True),
                sg.FolderBrowse(key='-FBrowse-', disabled=True)
            ]
        ]

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
            [sg.Push(), sg.Text("SpecParser", font=_H1_FONT), sg.Push()],
            [sg.Push(), sg.Text("Формирование заказа", font=_H2_FONT), sg.Push()],
            [
                sg.Push(),
                sg.Column(self._main_lcolumn_lout()),
                sg.VSeperator(),
                sg.Column(self._main_rcolumn_lout()),
                sg.Push()
            ],
            [sg.Push(), sg.Button("Парсить"), sg.Button('Config'), sg.Help()],
            [sg.VPush()]
        ]

    def get_config_layout(self) -> list[list]:
        return [
            [sg.VPush()],
            [sg.Push(), sg.Text('Настройки SpecParser', font=_H1_FONT), sg.Push()],
            [sg.VPush()],
            [sg.Text('ВНИМАНИЕ! Пути вводить в виде: C:/Папка/Подпапка')],
            [sg.Text('Параметры')],
            [sg.Push(), *self._params_list_layout(), sg.Push()],
            [sg.HSeparator()],
            [sg.Text('Маски')],
            [sg.Push(), *self._masks_layout(), sg.Push()],
            [sg.VPush()],
            [
                sg.Button('На главную', key='-GOHOME-'),
                sg.Button('Help', key='-HELP_CONFIG-'),
                sg.Push(),
                sg.Button('Сохранить', key='-SAVECONFIG-'),
                sg.Button('Сбросить', key='-CONFIG_DEFAULT-')
            ]
        ]

    @staticmethod
    def get_main_help_layout():
        return [
            [sg.Text('Справка', font=_H1_FONT), sg.Push()],
            [sg.Text(_HELP_0)],
            [sg.Text('Данные для ввода', font=_H3_FONT), sg.Push()],
            [sg.Text(_HELP_1)],
            [sg.Text('Сохранение результата', font=_H3_FONT), sg.Push()],
            [sg.Text(_HELP_2)],
            [sg.Text('Выбор бланков для заполнения', font=_H3_FONT), sg.Push()],
            [sg.Text(_HELP_3)],
            [sg.Button('На главную', key='-GOHOME-')]
        ]

    @staticmethod
    def get_config_help_layout():
        return [
            [sg.Text('Справка конфига', font=_H1_FONT)],
            [sg.Text(_CONFIG_HELP_1)],
            [sg.Push(), sg.Button('Config')]
        ]

    def get_window(self, winname: str, params: dict, bhandler: BlanksHandler) -> sg.Window:
        self._params = params
        self._bhandler = bhandler
        layout = []
        if winname == 'main':
            layout = self.get_main_layout()
            el_just = 'left'
        elif winname == 'config':
            layout = self.get_config_layout()
            el_just = 'right'
        elif winname == 'Help':
            layout = self.get_main_help_layout()
            el_just = 'left'
        elif winname == 'config_help':
            layout = self.get_config_help_layout()
            el_just = 'left'
        return sg.Window(
            "PartParserComposer",
            layout,
            resizable=True,
            margins=(30, 30),
            element_justification=el_just,
            font=_DEFAULT_FONT,
            finalize=True
        )
