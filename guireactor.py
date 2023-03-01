import PySimpleGUI as sg
from blankshandler import BlanksHandler
from aux_code import curr_dt_subfolder, new_filepath
from confighandler import touch_config
import os

class GUIReactor:
    def __init__(self):
        self.event = None
        self.values = None
        self.window = None
        self.bhandler = None
        self.config = None
        self.br_flag = None
        self.con_flag = None
        self._savedisabled = True

    def react(self, event: str, values: dict, window: sg.Window, bhandler: BlanksHandler, config: dict):
        self.event = event
        self.values = values
        self.window = window
        self.bhandler = bhandler
        if self.event == '-DEFSAVE-':
            self.flip_save()
        if self.event == 'Парсить':
            self.parse()
        if event == "config":
            self.t_config()
        if event in ("Exit", sg.WIN_CLOSED):
            self.br_flag = True
        return self.br_flag, self.con_flag


    def flip_save(self):
        self._savedisabled = not self._savedisabled
        self.window['-SaveFolder-'].update(disabled=self._savedisabled)
        self.window['-FBrowse-'].update(disabled=self._savedisabled)

    def parse(self):
        parse_list = list()
        for blankname in self.bhandler.get_blanks_names():
            if self.values[blankname]:
                parse_list.append(blankname)
        if self.values['-DEFSAVE-']:
            new_folder = curr_dt_subfolder(self.values['-XLSX-'])
            os.makedirs(new_folder, exist_ok=True)
            for blank in parse_list:
                new_path = new_filepath(self.values['-PROJECTNAME-'], new_folder, blank)
                self.bhandler.form_order_fromname(blank, self.values['-XLSX-'], new_path)
        sg.popup_ok('Спека спарсена!')
        self.br_flag = True

    def t_config(self):
        config_temp = touch_config(True)
        if config_temp:
            CONFIG = config_temp