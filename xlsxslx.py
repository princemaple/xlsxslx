from pathlib import Path
import win32com.client as win32


def convert(excel, path: Path):
    '''
    FileFormat = 51 is for .xlsx extension
    FileFormat = 56 is for .xls extension
    '''
    wb = excel.Workbooks.Open(str(path))
    wb.Password = ''
    wb.SaveAs(
        str(path.with_suffix('.xlsx' if path.suffix == '.xls' else '.xls')),
        FileFormat=51 if path.suffix == '.xls' else 56,
    )
    wb.Close()


def main(paths, glob='*.xls'):
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    for path in paths:
        path = Path(path).absolute()
        if path.is_file():
            convert(excel, path)
        elif path.is_dir():
            for path in path.glob(glob):
                convert(excel, path)

    excel.Application.Quit()  # type: ignore


if __name__ == '__main__':
    import PySimpleGUI as sg

    window = sg.Window(
        'Transform XLS <-> XLSX',
        [
            [sg.Button('xls->xlsx', key=1), sg.Button('xlsx->xls', key=2)],
            [sg.Button('xls=>>xlsx', key=11), sg.Button('xlsx=>>xls', key=12)],
        ],
        size=(200, 100),
    )

    command, _values = window.read()

    glob = ''
    if command == 1:
        path = sg.popup_get_file(
            'XLS files', 'Select files', file_types=(('XLS', '*.xls'),), multiple_files=True
        )
    elif command == 2:
        path = sg.popup_get_file(
            'XLSX files', 'Select files', file_types=(('XLSX', '*.xlsx'),), multiple_files=True
        )
    elif command == 11:
        path, glob = sg.popup_get_folder('XLS folder', 'Select foler'), '*.xls'
    elif command == 12:
        path, glob = sg.popup_get_folder('XLSX folder', 'Select folder'), '*.xlsx'
    else:
        raise Exception('Unknown')

    if path:
        main(path.split(';'), glob)
