from imaplib import IMAP4_SSL
from re import search as re_search
from datetime import datetime, timedelta
from email import utils, message_from_string
from sys import argv
from sys import exit as sys_exit
from os import system as os_system
from sys import platform

# размеры окна
if platform == 'win32':
    os_system("mode con cols=150 lines=30")

# проверка параметров
if len(argv) < 7 or '-s' not in argv or '-l' not in argv or '-p' not in argv:
    print("Неверное число параметров, ожидалось -s 'server' -l 'email' -p 'password'")
    input('\n=== нажмите Enter для выхода ===')
    sys_exit()
# разбор параметров
server = argv[argv.index('-s') + 1]
login = argv[argv.index('-l') + 1]
password = argv[argv.index('-p') + 1]

current_date = datetime.today()
days_ahead = argv.index('-d') + 1 if '-d' in argv else 2  # на сколько дней вперёд
dates = [current_date]
dates.extend([current_date + timedelta(days=i) for i in range(1, days_ahead + 1)])
for date in dates:
    dates[dates.index(date)] = date.strftime("%d.%m.%Y")

are_lessons_found = False
subject_types = {'Лекция': 'лекция',
                 'Практическое занятие': 'практика',
                 'Лабораторное занятие': 'лаба'}
count = 50  # сколько писем проверять от последнего

try:
    imap = IMAP4_SSL('imap.' + server)
    imap.login(login, password)
    print('Авторизация успешна, поиск...\n')
    imap.select("inbox")

    _, data = imap.search(None, 'ALL')
    ids = data[0]  # Получаем строку номеров писем
    id_list = ids.split()  # Разделяем ID писем

    for i in range(1, count):
        email_id = id_list[-i]
        _, data = imap.fetch(email_id, "(RFC822)")
        raw_email = data[0][1].decode('utf-8')

        sender = utils.parseaddr(message_from_string(raw_email)['From'])[1]
        if sender == 'process@isu.ifmo.ru':

            for date in dates:
                current_date_index = raw_email.find(date)
                if raw_email.find(date) != -1:  # если нашлась текущая дата цикла
                    are_lessons_found = True
                    lesson_time = re_search("([0-1]?[0-9]|2[0-3]):[0-5][0-9]", raw_email[current_date_index:]).group()
                    link = re_search(r"(?P<url>https?://itmo.zoom[^\s]+)",
                                     raw_email[current_date_index:]).group("url")[:-1]
                    subject_start_index = raw_email.find('дисциплине') + len('дисциплине') + 3
                    subject_end_index = raw_email.find('преподаватель') - 4
                    subject_name = ' '.join(raw_email[subject_start_index:subject_end_index].split())

                    # определяем тип занятия
                    subject_type = None
                    for subj_type_raw, subj_type in subject_types.items():
                        if ' '.join(raw_email.split()).find(subj_type_raw) != -1:
                            subject_type = subj_type
                            break

                    print(f'{subject_name} ({subject_type}) {date} в {lesson_time}: {link}')

    imap.logout()

except Exception as exc:
    if str(exc).startswith('[Errno 11001]'):
        print(f'Неверный адрес сервера ({exc})')
    elif str(exc).lower().find('authentication') != 0:
        print(f'Неверный email или пароль ({exc})')
else:
    if are_lessons_found:
        print('\nСкопируйте нужную ссылку в буфер обмена')
    else:
        print('Ссылки на ближайшие пары не найдены')
finally:
    input('=== нажмите Enter для выхода ===')

# pyinstaller --onefile main.py
