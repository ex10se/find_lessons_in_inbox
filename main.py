import email
import os
import re
from datetime import datetime, timedelta
from email.header import decode_header
from imaplib import IMAP4_SSL
from sys import argv
from sys import platform
import getpass

import pandas as pd

# размеры окна
if platform == 'win32':
    os.system("mode con cols=150 lines=20")

# проверка параметров
if len(argv) < 7 or '-s' not in argv or '-e' not in argv or '-p' not in argv:
    print("Неверное число параметров, ожидалось -s 'server' -e 'email' -p 'password'")
    print("Закройте программу или введите данные")
    SERVER = input("Server: ")
    EMAIL = input("Email: ")
    PASSWORD = getpass.getpass("Password (no echo): ")
else:
    # разбор параметров
    SERVER = argv[argv.index('-s') + 1]
    EMAIL = argv[argv.index('-e') + 1]
    PASSWORD = argv[argv.index('-p') + 1]

DAYS_AHEAD = int(argv[argv.index('-d') + 1]) if '-d' in argv else 3  # на сколько дней вперёд

dates = [(datetime.today() + timedelta(days=i)).strftime("%d.%m.%Y") for i in range(DAYS_AHEAD + 1)]
are_lessons_found = False
count = 50  # сколько писем проверять от последнего
result_rows = []

subject_types = {
    'Практическое занятие': 'Практика',
    'Лабораторное занятие': 'Лаба',
}


def decompose_letter(body: str) -> dict:
    lesson_date = re.search(r'(?<=состоится\s).*?(?=\s)', body).group()
    lesson_url = re.search(r'(?<=Вы\sможете\sпо\s<a\shref=").*?(?="\s)', body)
    if not lesson_url:
        lesson_url = re.search(r'https://itmo\.zoom\.us/.*?(?=<)', body)
    return {
        'Дата': lesson_date,
        'Время': re.search(rf'(?<={re.escape(lesson_date)} ).*?(?=\sв)', body).group(),
        'Тип': re.search(r'(?<=Вас,\sчто\s)[\w\W]*?(?=\sпо)', body).group(),
        'Предмет': re.search(r'(?<=по\sдисциплине\s).*?(?=,)', body).group(),
        'Ссылка': lesson_url.group(),
    }


def print_results(rows) -> None:
    df = pd.DataFrame(rows)
    df['Дата'] = pd.to_datetime(df['Дата'], dayfirst=True)
    df.sort_values(by=['Дата', 'Время'], inplace=True)
    for row in df.values:
        # сокращение названия предмета
        subject_name_words = row[3].split(' ')
        subject_name = ""
        k = 6  # максимальная длина слов
        for i in range(len(subject_name_words)):
            subject_name += f"{subject_name_words[i][:k]}"
            subject_name += '.' if len(subject_name_words[i]) > 1 else ''
            subject_name += '' if i + 1 == len(subject_name_words) else ' '
        # вывод результата
        print(f'{row[0].strftime("%d.%m")} в {row[1]}: '
              f'{subject_types.get(row[2]) if subject_types.get(row[2]) else row[2]}'  # сокращает тип пары
              f' по "{subject_name}" {row[4]}')


if __name__ == '__main__':
    try:
        with IMAP4_SSL(f'imap.{SERVER}') as imap:
            imap.login(EMAIL, PASSWORD)
            print('Авторизация успешна, поиск...\n')
            imap.select("inbox")
            id_list = imap.search(None, 'ALL')[1][0].split()[::-1][:count]

            for email_id in id_list:  # берем по письму, начиная с самого последнего
                data = imap.fetch(email_id, "(RFC822)")[1][0][1]
                msg = email.message_from_bytes(data)

                if msg['Return-path'] == '<process@isu.ifmo.ru>' and \
                        decode_header(msg['Subject'])[0][0].decode('utf-8') == 'ИСУ ИТМО - Дистанционное обучение':
                    payload = msg.get_payload()[1].get_payload()  # письмо в сыром виде
                    body_no_breaks = re.sub(r'^\s+|\n|\r|\s+$', '', payload)  # удаляем переносы строк
                    letter_details = decompose_letter(body_no_breaks)
                    if letter_details['Дата'] in dates:
                        # если дата не сегодняшняя, либо сегодняшняя, но
                        # между текущим временем и временем пары менее трех часов
                        try:
                            lesson_time = datetime.strptime(letter_details['Время'], '%H:%M')
                        except ValueError:
                            if letter_details['Дата'] != dates[0] or letter_details['Дата'] == dates[0]:
                                letter_details['Время'] = letter_details['Время'].replace('<b>', '').replace('</b>', '')
                                result_rows.append(letter_details)
                                are_lessons_found = True
                        else:
                            if letter_details['Дата'] != dates[0] or letter_details['Дата'] == dates[0] and \
                                    datetime.now().hour - lesson_time.hour < 3:
                                result_rows.append(letter_details)
                                are_lessons_found = True

    except Exception as exc:
        if str(exc).startswith('[Errno 11001]'):
            print(f'Неверный адрес сервера ({exc})')
        elif str(exc).lower().find('authentication') != 0:
            print(f'Неверный email или пароль ({exc})')
    else:
        if are_lessons_found:
            print_results(result_rows)
            print('\nСкопируйте нужную ссылку в буфер обмена')
        else:
            print('Ссылки на ближайшие пары не найдены')
    finally:
        input('\n=== нажмите Enter для выхода ===')

    # pyinstaller --onefile main.py
