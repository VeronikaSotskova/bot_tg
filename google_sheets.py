# Подключаем библиотеки
import httplib2
from oauth2client.service_account import ServiceAccountCredentials
from sheets_util import *
import constants

ROWS_COUNT = 0

CREDENTIALS_FILE = constants.credentials

# Читаем ключи из файла
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                               ['https://www.googleapis.com/auth/spreadsheets',
                                                                'https://www.googleapis.com/auth/drive'])

# Авторизуемся в системе
httpAuth = credentials.authorize(httplib2.Http())

# Выбираем работу с таблицами и 4 версию API
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

'''
Создаем гугл док с таблицей
'''

spreadsheet = create_spreadsheet(service, 'test')

# сохраняем идентификатор файла
spreadsheetId = spreadsheet['spreadsheetId']
print('https://docs.google.com/spreadsheets/d/' + spreadsheetId)

# Выдаем разрешение на файл
allow_permission(httpAuth, spreadsheetId, 'verasos13@gmail.com')

# Добавление листа
results = add_sheet(service, spreadsheetId, '2 List')

# Получаем список листов, их Id и название
sheetList = get_lists_by_id(service, spreadsheetId)
for sheet in sheetList:
    print(get_id_by_sheet(sheet), get_title_by_sheet(sheet))

sheetId = sheetList[0]['properties']['sheetId']
sheetTitle = get_title_by_sheet(sheetList[0])
print('Title', sheetTitle)

print('Мы будем использовать лист с Id = ', sheetId)

'''
Заполняем ячейки
'''
data_title = ['Имя', 'Фамилия', 'Возраст']
add_title(service, spreadsheetId, data_title, sheetTitle)
ROWS_COUNT += 1

data_people1 = ['123', '----', 19]
add_row(service, spreadsheetId, sheetTitle, data_people1, (ROWS_COUNT, 0))
ROWS_COUNT += 1

data_people2 = ['asda', 'dass', 2]
add_row(service, spreadsheetId, sheetTitle, data_people2, (ROWS_COUNT, 0))
ROWS_COUNT += 1

data_people3 = ['asddsa', 'hfasass', 24]
add_row(service, spreadsheetId, sheetTitle, data_people3, (ROWS_COUNT, 0))
ROWS_COUNT += 1



