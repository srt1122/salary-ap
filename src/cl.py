import FreeSimpleGUI as sg
import sys

def input_years():

    sg.theme('DarkBlue')

    layout = [[sg.Text('半角英数字で年月を入力して下さい。《例：2023/05》')],
              [sg.Text('月日'), sg.InputText(size=(10,1), key='text')],
              [sg.Button('OK', key='ok')]]

    window = sg.Window('入力', layout)

    while True:

      global csv_years, folder_years

      event, values = window.read()
      if event == sg.WIN_CLOSED:
        break
      elif event == 'ok':
        years = values['text']
        sp_years = years.split('/')
        try:
          csv_years = sp_years[0]+sp_years[1]
        except IndexError:
          ms = value = sg.popup_error('年月を入力して下さい。')
          if ms == 'Error':
            sys.exit()
        folder_years = sp_years[0]+'.'+sp_years[1]

        break

    window.close()

    return csv_years, folder_years,

def create_error_message(word):

    ms = sg.popup_error(word)
    if ms == 'Error':
      sys.exit()

def create_ok_message():

  ms = sg.popup("出力しました！")

