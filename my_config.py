import datetime


#------------ Работа с Excel ------------#

# название файла Excel, который нужно обновить, обязательно с расширением
name_excel_for_update = 'sample_data_excel.xlsx'

# список подключений, которые нужно обновить
# если список будет пустым -> то будут обновлены все подключения в Excel
list_connections = [
    'sample_data_cvs_1',
    'sample_data_cvs_2',
    'sample_data_cvs_3',
    'sample_data_cvs_4',
    'sample_data_cvs_5'
    ]

# пауза, которая берется после обновления подключения, чтобы Excel "отдохнул", в секундах
pause_after_upd = 10

# если после какого-либо обновления подключения трубуется больше времени на обновление сводных таблиц в Excel
# укажи названия этих подключений и после обновления этих подключений пауза будет дольше
long_list_connections = []

# увеличенная пауза после обновления подключений из списка long_list_connections, в секундах
long_pause_after_upd = 100

# пауза, которая берется после падения обновления Excel, в секундах
pause_after_error = 30


#------------ Работа с почтой (Outlook) ------------#

# нужно искать письмо с алертом?
# search_mail = True # если нужно
search_mail = False # если НЕ нужно

# получатель письма по почте, можно несколько
# регистр неважен -> работа будет в нижнем регистре
recipient_mail = ['irkvins-ivan@yandex.ru', 'nazvan.marketolog.sait@yandex.ru']

# пауза, которая берется на обновление (синхронизации) почты в Outlook, в секундах
outlook_wait_time = 20

# переменная, которая определяет нужный аккаунт в Outlook, указывать как адрес почты
# с этого адреса будут отправлять письма
# регистр неважен -> работа будет в нижнем регистре
account_outlook = 'irkvins-ivan@yandex.ru'

# в какой папке в почте искать письмо
# регистр неважен -> поиск названия папки будет в нижнем регистре
alert_folder = 'алерты'

# CSV файл, в который вставляется дата, когда обновлялся Excel
filename_csv = 'check_update_excel.csv'

# тема письма для алертов, чтобы можно было складывать их в отдельную папку
theme_mail = 'Обновление Excel // .py'

# выбери нужный формат записи даты
today = datetime.date.today().strftime('%d.%m.%y') # дата формата '03.09.25'
# today = datetime.date.today().strftime('%d.%m.%Y') # дата формата '03.09.2025'
# today = datetime.date.today().strftime('%Y-%m-%d') # дата формата '2025-09-03'
# today = datetime.date.today().strftime('%d-%b-%Y').upper() # дата формата '03-SEP-2025'
# today = datetime.date.today().strftime('%d-%b-%y').upper() # дата формата '03-SEP-25'
# today = datetime.date.today().strftime('%d-%b-%Y').lstrip('0').upper() # дата формата '3-SEP-2025'
# today = datetime.date.today().strftime('%d-%b-%y').lstrip('0').upper() # дата формата '3-SEP-25'

# текст (патерн поиска письма), который нужно искать в письме
# регистр неважен -> поиск текста письма будет в нижнем регистре
# сам текст письма тоже будет приведен к нижнему регистру
find_text_in_mail = [f'PROCEDURE TEST_TEST_TEST {today}',
                     'Витрина успешно сформирована'
                     ]

# отправлять уведомление по почте, если письмо из переменной find_text_in_mail НЕ нашлось?
send_mail_if_mail_no_found = True # если нужно
# send_mail_if_mail_no_found = False # если НЕ нужно


#------------ Не трогать, для тестов ------------#

# raise Exception('Это специально вызванная ошибка')
