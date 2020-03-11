import apiclient
import httplib2
from oauth2client.service_account import ServiceAccountCredentials

import constants
import telebot
from telebot import apihelper, types

from sheets_util import *

token = constants.token
proxy = constants.proxy
INFO_FILE = 'table_info.txt'
CREDENTIALS_FILE = constants.credentials
TEMP_FILE = 'temp.txt'

bot = telebot.TeleBot(token)
apihelper.proxy = {'https': f'socks5://{proxy}'}

credentials = ''
httpAuth = ''
service = ''
SPREAD_NAME = 'test'
LIST_NAME = 'LIST_1'
data_title = ['Имя', 'Фамилия', 'Возраст']


def start_init_table():
    global credentials
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                   ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])
    # Авторизуемся в системе
    global httpAuth
    httpAuth = credentials.authorize(httplib2.Http())

    # Выбираем работу с таблицами и 4 версию API
    global service
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
    spreadsheet = create_spreadsheet(service, SPREAD_NAME, LIST_NAME)
    allow_permission(httpAuth, spreadsheet['spreadsheetId'], 'verasos13@gmail.com')

    # инициализируем начальные данные
    write_to_file(0, spreadsheet['spreadsheetId'], LIST_NAME)
    add_title(service, spreadsheet['spreadsheetId'], data_title, LIST_NAME)
    print('https://docs.google.com/spreadsheets/d/' + spreadsheet['spreadsheetId'])


@bot.message_handler(content_types=['text'])
def start(message):
    if message.text == '/reg':
        bot.send_message(message.from_user.id, "Как тебя зовут?")
        bot.register_next_step_handler(message, get_name)  # следующий шаг – функция get_name
    elif message.text == '/init':
        start_init_table()
        bot.send_message(message.from_user.id, 'Напиши /reg чтобы пройти опрос')
    else:
        bot.send_message(message.from_user.id, 'Напиши /reg чтобы пройти опрос' + '\n' +
                         'Напиши /init чтобы создать гугл док')


def get_name(message):  # получаем фамилию
    f = open(TEMP_FILE, 'w')
    name = message.text
    name_to_file = str(name) + '\n'
    f.write(name_to_file)
    f.close()
    bot.send_message(message.from_user.id, 'Какая у тебя фамилия?')
    bot.register_next_step_handler(message, get_surname)


def get_surname(message):
    f = open(TEMP_FILE, 'a')
    surname = message.text
    surname_to_file = str(surname) + '\n'
    f.write(surname_to_file)
    f.close()
    bot.send_message(message.from_user.id, 'Сколько тебе лет?')
    bot.register_next_step_handler(message, get_age)


def get_age(message):
    try:
        f = open(TEMP_FILE, 'a')
        age = int(message.text)
        age_to_file = str(age) + '\n'
        f.write(str(age))
        f.close()
    except Exception:
        bot.send_message(message.from_user.id, 'Цифрами пожалуйста')
    # наша клавиатура
    keyboard = types.InlineKeyboardMarkup()
    # кнопка «Да»
    key_yes = types.InlineKeyboardButton(text='Да', callback_data='yes')
    # добавляем кнопку в клавиатуру
    keyboard.add(key_yes)
    key_no = types.InlineKeyboardButton(text='Нет', callback_data='no')
    keyboard.add(key_no)
    data_temp = read_temp()
    question = 'Тебе ' + str(data_temp[2]) + ' лет, тебя зовут ' + data_temp[0] + ' ' + data_temp[1] + '?'
    bot.send_message(message.from_user.id, text=question, reply_markup=keyboard)


@bot.callback_query_handler(func=lambda call: True)
def callback_worker(call):
    # call.data это callback_data, которую мы указали при объявлении кнопки
    if call.data == "yes":
        # сохраняем данные
        data_people = read_temp()
        info_table = read_table_info()
        add_row(service, info_table[1], info_table[2], data_people, (info_table[0] + 1, 0))
        update_info_file_1()
        bot.send_message(call.message.chat.id, 'Запомню : )' + '\n' +
                         'https://docs.google.com/spreadsheets/d/' + str(read_table_info()[1]))

    elif call.data == "no":
        pass


def read_temp():
    arr = []
    f = open(TEMP_FILE, 'r')
    for line in f:
        arr.append(line.split('\n')[0])
    f.close()
    return arr[0], arr[1], int(arr[2])


def read_table_info():
    """
    1 строка - ROWS_TABLE
    2 строка - spreadsheetId
    3 строка - название листа
    """
    arr = []
    f = open(INFO_FILE, 'r')
    for line in f:
        arr.append(line.split('\n')[0])
    return int(arr[0]), arr[1], arr[2]


def write_to_file(rows_count, spreadsheet_id, list_name):
    f = open(INFO_FILE, 'w')
    f.write(str(rows_count) + '\n' + str(spreadsheet_id) + '\n' + str(list_name))
    f.close()


def update_info_file_1():
    info = read_table_info()
    rows = int(info[0])
    spread_id = info[1]
    title = info[2]
    rows += 1
    write_to_file(rows, spread_id, title)


# Запускаем постоянный опрос бота в Телеграме
bot.polling(none_stop=True, interval=0)
