import PySimpleGUI as sg

from blankshandler import BlanksHandler
from confighandler import touch_config
from guireactor import GUIReactor
from winconstructor import WinConstructor


if __name__ == '__main__':
    params = touch_config()
    reactor = GUIReactor()
    bhandler = BlanksHandler(params)
    wcons = WinConstructor()
    newwindow = 'main'
    window = None
    while True:
        if params is None:
            sg.popup_error('Config не выполнен!')
            break
        if newwindow:
            window = wcons.get_window(newwindow, params, bhandler)
            newwindow = None
        event, values = window.read(timeout=20)
        br_flag, newwindow = reactor.react(event, values, window, params)
        if br_flag:
            break
        if newwindow:
            window.close()
    window.close()
