import PySimpleGUI as sg
from blankshandler import BlanksHandler
from confighandler import touch_config
from guireactor import GUIReactor
from winconstructor import WinConstructor


if __name__ == '__main__':
    try:
        reactor = GUIReactor()
        wcons = WinConstructor()
        newwindow = 'main'
        window = None
        while True:
            if newwindow:
                try:
                    params = touch_config()
                    bhandler = BlanksHandler(params)
                    window = wcons.get_window(newwindow, params, bhandler)
                    newwindow = None
                except Exception as e:
                    sg.popup_error(
                        'Произошла ошибка при чтении config.ini:\n\n' +
                        str(e)
                    )
                    break
            event, values = window.read(timeout=20)
            br_flag, newwindow = reactor.react(event, values, window, params)
            if br_flag:
                break
            if newwindow:
                window.close()
        if window:
            window.close()
    except Exception as e:
        sg.popup_error(
            'Обнаружена криитическая ошибка:\n\n' +
            str(e)
        )
