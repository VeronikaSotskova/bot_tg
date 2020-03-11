import apiclient.discovery


def rowcol_to_a1(row, col):
    if col < 0 or col > 25:
        print('error. Invalid column value')
        return Exception
    r = chr(col + 65)
    res = ''
    res += str(r)
    res += str(row + 1)
    return res


def a1_to_rowcol(label: str):
    if label[0] < 'A' or label[0] > 'Z':
        print('error. Invalid column value')
        return Exception
    col = ord(label[0]) - 65
    try:
        row = int(label[1:] - 1)
    except Exception:
        print('invalid cell value')
    return row, col


def get_a1_notation(sheetTitle: str, endCell: tuple, startCell: tuple = (0, 0)):
    sheet_range = str(sheetTitle) + '!'
    cell_start = rowcol_to_a1(startCell[0], startCell[-1])
    cell_end = rowcol_to_a1(endCell[0], endCell[-1])
    sheet_range += cell_start + ':' + cell_end
    return sheet_range


def create_spreadsheet(service, title, title_1_list='Лист номер один', rowCount=100, colCount=15):
    return service.spreadsheets().create(body={
        'properties': {'title': title, 'locale': 'ru_RU'},
        'sheets': [{'properties': {'sheetType': 'GRID',
                                   'sheetId': 0,
                                   'title': title_1_list,
                                   'gridProperties': {'rowCount': rowCount, 'columnCount': colCount}}}]
    }).execute()


def add_sheet(service, spreadsheetId, title, rowCount=20, colCount=12):
    return service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheetId,
        body=
        {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": title,
                            "gridProperties": {
                                "rowCount": rowCount,
                                "columnCount": colCount
                            }
                        }
                    }
                }
            ]
        }).execute()


def allow_permission(httpAuth, spreadsheetId, email):
    # Выбираем работу с Google Drive и 3 версию API
    driveService = apiclient.discovery.build('drive', 'v3', http=httpAuth)
    # Открываем доступ на редактирование
    access = driveService.permissions().create(
        fileId=spreadsheetId,
        body={'type': 'user', 'role': 'writer', 'emailAddress': email},
        fields='id'
    ).execute()
    return access


def get_lists_by_id(service, spreadsheetId):
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    return spreadsheet.get('sheets')


def get_id_by_sheet(sheet):
    return sheet['properties']['sheetId']


def get_title_by_sheet(sheet):
    return sheet['properties']['title']


def add_values_to_column(service, spreadsheetId, data: list, sheetTitle, startCell: tuple = (0, 0)):
    cell_end = (len(data), startCell[-1])
    sheet_range = get_a1_notation(sheetTitle, startCell=startCell, endCell=cell_end)

    return service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
        "valueInputOption": "USER_ENTERED",
        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
        "data": [
            {"range": sheet_range,
             # Сначала заполнять столбцы, затем строки
             "majorDimension": "COLUMNS",
             "values": [
                 data
             ]}
        ]
    }).execute()


def draw_border(service, spreadsheetId, sheetTitle, endCell: tuple, width: int = 1, startCell: tuple = (0, 0)):
    # Рисуем рамку

    sheet_range = get_a1_notation(sheetTitle, startCell=startCell, endCell=endCell)
    results = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheetId,
        body={
            "requests": [
                {'updateBorders': {
                    'range': sheet_range,
                    # Задаем стиль для верхней границы
                    'bottom': {
                        # Сплошная линия
                        'style': 'SOLID',
                        # Шириной 1 пиксель
                        'width': width,
                        # Черный цвет
                        'color': {
                            'red': 0,
                            'green': 0,
                            'blue': 0,
                            'alpha': 1
                        }
                    },
                    # Задаем стиль для нижней границы
                    'top': {
                        'style': 'SOLID',
                        'width': width,
                        'color': {
                            'red': 0,
                            'green': 0,
                            'blue': 0,
                            'alpha': 1
                        }
                    },
                    # Задаем стиль для левой границы
                    'left': {
                        'style': 'SOLID',
                        'width': width,
                        'color': {
                            'red': 0,
                            'green': 0,
                            'blue': 0,
                            'alpha': 1
                        }
                    },
                    # Задаем стиль для правой границы
                    'right': {
                        'style': 'SOLID',
                        'width': width,
                        'color': {
                            'red': 0,
                            'green': 0,
                            'blue': 0,
                            'alpha': 1
                        }
                    },
                    # Задаем стиль для внутренних горизонтальных линий
                    'innerHorizontal': {
                        'style': 'SOLID',
                        'width': width,
                        'color': {
                            'red': 0,
                            'green': 0,
                            'blue': 0,
                            'alpha': 1
                        }
                    },
                    # Задаем стиль для внутренних вертикальных линий
                    'innerVertical': {
                        'style': 'SOLID',
                        'width': width,
                        'color': {
                            'red': 0,
                            'green': 0,
                            'blue': 0,
                            'alpha': 1
                        }
                    }
                }}
            ]
        }).execute()


def add_title(service, spreadsheetId, title_of_col: list, sheetTitle, startCell: tuple = (0, 0)):
    cell_end = (startCell[0], startCell[-1] + len(title_of_col))
    sheet_range = get_a1_notation(sheetTitle, startCell=startCell, endCell=cell_end)

    service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
        "valueInputOption": "USER_ENTERED",
        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
        "data": [
            {"range": sheet_range,
             # Сначала заполнять столбцы, затем строки
             "majorDimension": "ROWS",
             "values": [
                 title_of_col
             ]}
        ]
    }).execute()



def add_row(service, spreadsheetId, sheetTitle, object_data, startCell: tuple = (0, 1)):
    cell_end = (startCell[0] + 1, startCell[-1] + len(object_data))
    sheet_range = get_a1_notation(sheetTitle, startCell=startCell, endCell=cell_end)

    return service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
        "valueInputOption": "USER_ENTERED",
        "data": [
            {"range": sheet_range,
             # Сначала заполнять строки, затем столбцы
             "majorDimension": "ROWS",
             "values": [
                 object_data
             ]}
        ]
    }).execute()
