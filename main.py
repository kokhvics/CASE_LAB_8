from openpyxl import load_workbook
import PySimpleGUI as sg

#лабораторная работа с графическим интерфейсом GUI
elements = [[sg.Text('Поставщик'), sg.Push(), sg.Input(key='post')],
          [sg.Text('Номер партии'), sg.Push(), sg.Input(key='part')],
          [sg.Text('Название модели'), sg.Push(), sg.Input(key='name')],
          [sg.Text('Пол'), sg.Push(), sg.Input(key='sex')],
          [sg.Text('Размерная сетка'), sg.Push(), sg.Input(key='size')],
          [sg.Text('Количество пар'), sg.Push(), sg.Input(key='numb')],
          [sg.Button('Добавить'), sg.Button('Закрыть')]]

window = sg.Window('База данных', elements, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Закрыть':
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('base.xlsx')
            sheet = wb['Лист1']
            ID = len(sheet['ID']) + 1


            data = [ID, values['post'], values['part'], values['name'], values['sex'], values['size'],values['numb']]

            sheet.append(data)

            wb.save('base.xlsx')

            window['post'].update(value='')
            window['part'].update(value='')
            window['name'].update(value='')
            window['sex'].update(value='')
            window['size'].update(value='')
            window['numb'].update(value='')
            window['name'].set_focus()

            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User.\nPlease try again later.')

window.close()
