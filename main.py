from imaplib import IMAP4_SSL
from re import search as re_search
from datetime import datetime, timedelta
from email import utils, message_from_string
from sys import argv
from sys import exit as sys_exit
from os import system as os_system

os_system("mode con cols=150 lines=30")  # размеры окна

if len(argv) != 7 or '-s' not in argv or '-l' not in argv or '-p' not in argv:
    print("Неверное число параметров, ожидалось -s 'server' -l 'email' -p 'password'")
    input('\n=== нажмите Enter для выхода ===')
    sys_exit()

server = argv[argv.index('-s') + 1]
login = argv[argv.index('-l') + 1]
password = argv[argv.index('-p') + 1]

try:
    imap = IMAP4_SSL('imap.' + server)
    imap.login(login, password)
    print('Авторизация успешна, поиск...\n')
    imap.select("inbox")

    _, data = imap.search(None, 'ALL')

    ids = data[0]  # Получаем строку номеров писем
    id_list = ids.split()  # Разделяем ID писем

    current_date = datetime.now()
    is_lessons_found = False

    for date in (current_date, current_date + timedelta(days=1)):
        date = date.strftime("%d.%m.%Y")

        for i in range(1, 50):
            email_id = id_list[-i]
            _, data = imap.fetch(email_id, "(RFC822)")
            raw_email = data[0][1].decode('utf-8')
            who_sent = utils.parseaddr(message_from_string(raw_email)['From'])[1]
            if who_sent == 'process@isu.ifmo.ru':
                current_date_index = raw_email.find(date)
                if current_date_index != -1:  # если нашлась текущая дата
                    is_lessons_found = True
                    lesson_time = re_search("([0-1]?[0-9]|2[0-3]):[0-5][0-9]", raw_email[current_date_index:]).group()
                    link = re_search("(?P<url>https?://itmo.zoom[^\s]+)", raw_email[current_date_index:]).group("url")[
                           :-1]
                    subject_start_index = raw_email.find('дисциплине') + len('дисциплине') + 3
                    subject_end_index = raw_email.find('преподаватель') - 4
                    subject_name_raw = raw_email[subject_start_index:subject_end_index].split()
                    subject_name = ' '.join(subject_name_raw)
                    print(f'{subject_name} {date} в {lesson_time}: {link}')

    imap.logout()

except Exception as exc:
    if str(exc).startswith('[Errno 11001]'):
        print(f'Неверный адрес сервера ({exc})')
    elif str(exc).lower().find('authentication') != 0:
        print(f'Неверный email или пароль ({exc})')
else:
    if is_lessons_found:
        print('\nСкопируйте нужную ссылку в буфер обмена')
    else:
        print('Ссылки на сегодняшние и завтрашние пары не найдены')
finally:
    input('\n=== нажмите Enter для выхода ===')

# pyinstaller --onefile main.py
