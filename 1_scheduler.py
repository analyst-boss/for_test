# чтобы не использовать кэшированные обёртки для COM
import sys
sys.modules['win32com.gen_py'] = None

# импорты
import datetime
import os
import pandas as pd
import subprocess

# импорты функций из файла main_functions.py
import main_functions as m_func

# импорт значений из файла my_config.py
import my_config as config

# текущая дата
today = datetime.datetime.now().strftime('%Y-%m-%d')

# определение директории
try:
    main_dir = os.path.dirname(os.path.abspath(__file__)) # для .py
except NameError:
    main_dir = os.getcwd() # для .ipynb

# работаем с логами
# путь до папки с логами txt
log_dir = os.path.join(main_dir, 'log-files')

# формируем лог файл по маске в папке с логами
log_filename = f'log-file_{today}.txt'

# путь до нового сформированного txt с логами
new_log_path = os.path.join(log_dir, log_filename)

# название файла Excel, который нужно обновить
name_excel_for_update = config.name_excel_for_update

# CSV файл, в который вставляется DT, когда обновлялся Excel
filename_csv = config.filename_csv
path_csv = os.path.join(main_dir, filename_csv)

# проверяем налииче файла check_update_excel.csv
# если нет -> создать с дефолтными данными
# если есть -> ничего не делать
m_func.create_update_excel_today(name_excel_for_update, path_csv, new_log_path)

try:

    # читаем файл check_update_excel.csv
    df = pd.read_csv(path_csv, encoding = 'windows-1251')

    # берем из файла check_update_excel.csv последнюю дату обновления (дата последней записи)
    last_update_dt = df['update_dt'].max()

    # если текущая дата НЕ равна последнему обновлению Excel -> то обовлять Excel нужно
    if today != last_update_dt:

        # если нужно проверять письмо с алертом
        if config.search_mail is True:

            # путь до нужного файла .py
            script_run = os.path.join(main_dir, '2_check_update_excel_and_mail.py')

            # запускаем скрипт для проверки на обновление Excel
            subprocess.run(['python', script_run])
        
        # если переменная config.search_mail is False -> проверять письмо с алертом не нужно
        else:

            # путь до нужного файла .py
            script_run = os.path.join(main_dir, '2_check_update_excel.py')

            # запускаем скрипт для проверки на обновление Excel
            subprocess.run(['python', script_run])    

except Exception as e:
    
    m_func.log(new_log_path, f'Ошибка при отработке скрипта "1_scheduler.py": {str(e)}')

    # отправляем письмо
    m_func.send_mail_outlook(recipient_mail = config.recipient_mail,
                             subject_mail = config.theme_mail,
                             body_mail = f'Ошибка при отработке скрипта "1_scheduler.py": {str(e)}.\nЧитай файл с логами.',
                             attach_files = [new_log_path],
                             account_outlook = config.account_outlook.lower(),
                             new_log_path = new_log_path
                             )