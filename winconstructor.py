import PySimpleGUI as sg
from blankshandler import BlanksHandler

H1_FONT = ("Segoe UI", 22)
H2_FONT = ("Segoe UI", 12)
DEFAULT_FONT = ('Segoe UI', 10)
SMALL_FONT = ("Segoe UI", 8)
SPECTIP = 'Спецификация должна быть экспортирована из Солида "смещенной" с "подробной нумерацией"'
BLANKTIP = 'Бланк должен соответствовать правилам. См. файл "INFO"'


class WinConstructor:
    def __init__(self, bhandler: BlanksHandler):
        self.bhandler = bhandler
        pass

    def b_cb_layout(self):
        out_list = list()
        i = 0
        for entry in self.bhandler.get_blanks_names():
            out_list.append([sg.Checkbox(entry, key=entry, font=SMALL_FONT)])
            i += 1
        return out_list

    def get_main_layout(self):
        main_layout = [
            [sg.VPush()],
            [sg.Push(), sg.Text("Welcome to PartParser!", font=H1_FONT), sg.Push()],
            [sg.Push(), sg.Text("Формирование заказа", font=H2_FONT), sg.Push()],
            [sg.HSeparator()],
            [sg.Text('Введи наименование изделия', font=H2_FONT)],
            [sg.In(key='-PROJECTNAME-', expand_x=True, tooltip=SPECTIP)],
            [sg.HSeparator()],
            [sg.Text("Выбери эксель-файл спецификации", font=H2_FONT)],
            [sg.In(key='-XLSX-', expand_x=True, tooltip=SPECTIP), sg.FileBrowse(key='-XBrowse-')],
            [sg.HSeparator()],
            [sg.Text("Выбери бланки для заполнения из стандартных", font=H2_FONT)],
            self.b_cb_layout(),
            [sg.Text("Можешь добавить кастомный бланк для заполнения", font=H2_FONT)],
            [sg.In(key='-SPECSAVE-', expand_x=True, tooltip=BLANKTIP), sg.FileBrowse(key='-SBrowse-')],
            [sg.HSeparator()],
            [sg.Text("Сохранение результата", font=H2_FONT)],
            [sg.Checkbox('Сохранить в подпапке "/дата-время"', key='-DEFSAVE-', enable_events=True, default=True)],
            [sg.Text("Или выбери другое место")],
            [sg.In(enable_events=True, key="-SaveFolder-", expand_x=True, disabled=True), sg.FolderBrowse(key='-FBrowse-', disabled=True)],
            [sg.Push(), sg.Button("Парсить"), sg.Push(), sg.Button('config')],
            [sg.VPush()]
        ]
        return main_layout

    def get_main_window(self):
        return sg.Window(
            "PartParserComposer",
            self.get_main_layout(),
            size=(600, 800),
            margins=(30, 30),
            element_justification="left",
            font=DEFAULT_FONT,
            finalize=True
        )
