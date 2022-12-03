import PySimpleGUI as sg
import os.path
# sg.Window(title="Hello World", layout=[[]])
H1_FONT = ("Arial", 20)

# create layout
layout = [
    [sg.Text("Welcome to PartParser!", font=H1_FONT)],
    [sg.Text("Выбери Excel-файл")],
    [[sg.In(enable_events=True)], [sg.FileBrowse()]],
    [sg.Button("Check")]
]

# sg.Graph
# sg.

window = sg.Window(
    "PartParserComposer",
    layout,
    margins=(100, 300),
    resizable=True,
    element_justification="center"
    # finalize=True
    # no_titlebar=True
)

if __name__ == '__main__':
    # sg.Window("PartParserComposer", layout, margins=(100, 300), no_titlebar=True).read()
    print('foo')
    while True:
        event, values = window.read()
        if event in ("Exit", sg.WIN_CLOSED):
            break
    window.close()