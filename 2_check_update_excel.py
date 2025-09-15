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

# название файла Excel, который нужно обновить
name_excel_for_update = config.name_excel_for_update

# системное уведомление windows
m_func.show_notification_windows(title = name_excel_for_update,
                                 message = 'Начало работы "2_check_update_excel.py"'
                                 )

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

# список файлов с логами txt
list_log_files_txt = os.listdir(log_dir)

# путь до папки с архивами логов внутри папки с логами
zip_log_dir = os.path.join(log_dir,'zip-log-files.zip')

m_func.log(new_log_path, f'==============================================\n')
m_func.log(new_log_path, f'Работа скрипта "2_check_update_excel.py" за дату: {today}')
m_func.log(new_log_path, f'Начало работы скрипта: {datetime.datetime.now()}.')
           
# функция для архивации лог файлов
m_func.zip_log_files(new_log_path, zip_log_dir, list_log_files_txt, log_dir)

# функция для удаления лог файлов, которые до жтого были добавлены в архив
m_func.del_log_files(new_log_path, list_log_files_txt, log_dir)

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
    if today == last_update_dt:
    
        m_func.log(new_log_path, f'Excel сегодня уже обновлялся.')
        m_func.log(new_log_path, f'Письмо в Outlook отправленно.\n')
                
        # отправляем письмо
        m_func.send_mail_outlook(recipient_mail = config.recipient_mail,
                                subject_mail = config.theme_mail,
                                body_mail = f'Excel "{name_excel_for_update}" за сегодня УЖЕ обновлялся.\nДата записи о последнем автоматическом обновлении {df['update_dt'].max()}.\nЧитай файл с логами.',
                                attach_files = [new_log_path],
                                account_outlook = config.account_outlook.lower(),
                                new_log_path = new_log_path
                                )

    # если текущая дата НЕ равна последнему обновлению Excel -> то обовлять Excel нужно
    else:
    
        m_func.log(new_log_path, f'Excel сегодня HE обновлялся.')
        m_func.log(new_log_path, f'Запуск скрипта "3_update_excel.py" на обновление Excel.\n')
                
        # путь до нужного файла .py
        script_run = os.path.join(main_dir, '3_update_excel.py')
                                
        # запускаем скрип для обнолвения Excel
        # сохраняем успех выполнения скрипта в переменную result_script_run
        result_script_run = subprocess.run(['python', script_run])

        # если вызываемый скрипт выполнился без каких то ошибок
        if result_script_run.returncode == 0:
        
            # вносим запись, что сегодня Excel обновился
            m_func.insert_update_excel_today(name_excel_for_update, path_csv, new_log_path)
            m_func.log(new_log_path, f'Файл Excel обновлен.')
            m_func.log(new_log_path, f'Файл {filename_csv} обновлен.')
            m_func.log(new_log_path, f'Конец работы.\n')
                    
            # отправляем письмо
            m_func.send_mail_outlook(recipient_mail = config.recipient_mail,
                                    subject_mail = config.theme_mail,
                                    body_mail = f'Excel "{name_excel_for_update}" обновился.\nЧитай файл с логами.',
                                    attach_files = [new_log_path],
                                    account_outlook = config.account_outlook.lower(),
                                    new_log_path = new_log_path
                                    )

        # если вызываемый скрипт выполнился с ошиками
        else:
            m_func.log(new_log_path, f'Возникла ошибка при выполнении обновления Excel.\n')
            
            # отправляем письмо
            m_func.send_mail_outlook(recipient_mail = config.recipient_mail,
                                        subject_mail = config.theme_mail,
                                        body_mail = f'Возникла ошибка при выполнение обновления Excel "{name_excel_for_update}".\nЧитай файл с логами.',
                                        attach_files = [new_log_path],
                                        account_outlook = config.account_outlook.lower(),
                                        new_log_path = new_log_path
                                        )


except Exception as e:

    m_func.log(new_log_path, f'Ошибка при отработке скрипта "2_check_update_excel.py": {str(e)}\n')

    # отправляем письмо
    m_func.send_mail_outlook(recipient_mail = config.recipient_mail,
                             subject_mail = config.theme_mail,
                             body_mail = f'Excel "{name_excel_for_update}" HE обновился.\nОшибка при отработке скрипта "2_check_update_excel.py": {str(e)}.\nЧитай файл с логами.',
                             attach_files = [new_log_path],
                             account_outlook = config.account_outlook.lower(),
                             new_log_path = new_log_path
                             )
finally:

    # системное уведомление windows
    m_func.show_notification_windows(title = name_excel_for_update,
                                     message = 'Конец работы "2_check_update_excel.py"'
                                     )