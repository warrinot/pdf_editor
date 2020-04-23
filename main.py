import PySimpleGUI as sg
import re
import PyPDF2
import os


def find_pages(string_of_pages: list()):
    if '-' in string_of_pages:
        page_nums = string_of_pages.split('-')
        page_nums = list(range(int(page_nums[0]) - 1, int(page_nums[1])))

    else:
        page_nums = re.findall(r'\d+', string_of_pages)
        page_nums = [int(page) - 1 for page in page_nums]
    return page_nums


def split_doc(filename: str, pages_to_split: list, new_name: str, onefile: bool):
    pdf_file = open(filename, 'rb')  # open pdf_file
    pdf_reader = PyPDF2.PdfFileReader(pdf_file, strict=False)  # init pdf_reader
    folder = os.path.split(filename)[0]  # folder to save in original file's folder

    num_pages = pdf_reader.numPages

    if not pages_to_split:  # full doc, if pages not specified
        pages_to_split = range(num_pages)

    if new_name:  # new_name if specified
        new_file_name = os.path.join(folder, new_name)
        print(new_file_name)
    else:  # original name
        new_file_name = os.path.splitext(filename)[0] + '_new'
        print(new_file_name)

    if onefile:
        pdf_writer = PyPDF2.PdfFileWriter()
        for page in range(num_pages):
            if page not in pages_to_split:
                continue
            page_object = pdf_reader.getPage(page)
            pdf_writer.addPage(page_object)
        new_file = open(f'{new_file_name}.pdf', 'wb')
        pdf_writer.write(new_file)
        new_file.close()

    else:  # multiple files
        for page in range(num_pages):
            if page not in pages_to_split:
                continue
            page_object = pdf_reader.getPage(page)
            pdf_writer = PyPDF2.PdfFileWriter()
            pdf_writer.addPage(page_object)
            new_file = open(f'{new_file_name}_page{page+1}.pdf', 'wb')
            pdf_writer.write(new_file)
            new_file.close()

    pdf_file.close()


def main():
    sg.theme('Dark Blue 3')
    layout = [[sg.Text('Выберите действие:')],
              [sg.Button('Split')],
              [sg.Button('Merge(under development)', disabled=True)]
              ]
    window = sg.Window('Pdf things', layout)

    while True:
        event, values = window.read()

        if event is None or event == 'Exit':
            break
        elif event == 'Split':
            window.Hide()
            split_window()
            window.UnHide()

        elif event == 'Merge':
            window.Hide()
            merge_window()
            window.UnHide()

    window.close()


def split_window():
    layout_split = [[sg.Text('Выберите pdf файл, который нужно разделить')],

                    [sg.In(disabled=True), sg.FileBrowse(file_types=(("Pdf-file", "*.pdf"),),
                                                         key='file_to_split')],

                    [sg.Text('Введите номера страниц, которые нужно отделить')],

                    [sg.Input(key='-IN-', enable_events=True,
                              tooltip=('Введите номера страниц через запятую или оставьте пустым, '
                                       'чтобы разделить весь документ\n'))],

                    [sg.Radio('Сохранить страницы отдельно', "RADIO1", default=True, key='radio1'),
                     sg.Radio('Сохранить одним файлом', "RADIO1", key='radio2')],

                    [sg.Text('Имя нового файла'), sg.Input(disabled=False, key='new_name')],

                    [sg.Button('Split'), sg.Exit()]]

    window_split = sg.Window('Pdf Split', layout_split)

    while True:
        event, values = window_split.read(timeout=100)
        # print(event, values)
        if event is None or event == 'Exit':
            break

        if event == '-IN-' and values['-IN-'] and values['-IN-'][-1] not in ('0123456789, -'):
            window_split['-IN-'].update(values['-IN-'][:-1])

        elif event == 'Split':

            print(event, values)
            file_to_split = values['file_to_split']
            pages_to_split = find_pages(values['-IN-'])
            print(pages_to_split)
            new_file_name = values['new_name']
            onefile = values['radio2']
            try:
                split_doc(file_to_split, pages_to_split, new_file_name, onefile)
            except FileNotFoundError:
                sg.popup_ok('Не выбран файл.')

        # if values['radio2']:
            # window_split['new_name'].Update(disabled=False)
        # else:
            # window_split['new_name'].Update(disabled=True)
    window_split.close()


def merge_window():
    layout_merge = [[sg.Text('Выберите pdf файл')]]
    window_merge = sg.Window('Pdf merge', layout_merge)

    while True:
        event, values = window_merge.read()
        if event is None or event == 'Exit':
            break

        elif event == 'Merge':
            pass

    window_merge.close()


if __name__ == '__main__':
    main()
