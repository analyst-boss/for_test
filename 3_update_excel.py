# чтобы не использовать кэшированные обёртки для COM
import sys
sys.modules['win32com.gen_py'] = None

# импорты
import datetime
import os
import shutil

# импорты функций из файла main_functions.py
import main_functions as m_func

# импорт значений из файла my_config.py
import my_config as config

# фиксируем начало работы скрипта
start_play = datetime.datetime.now()

# текущая дата
today = datetime.datetime.now().strftime('%Y-%m-%d')

# название файла Excel, который нужно обновить
name_excel_for_update = config.name_excel_for_update

# системное уведомление windows
m_func.show_notification_windows(title = name_excel_for_update,
                                 message = 'Начало обновления'
                                 )

try:

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
    zip_log_dir = os.path.join(log_dir, 'zip-log-files-excel.zip')
                               
    # логируем время начала работы скрипта
    m_func.log(new_log_path, f'**********************************************\n')
    m_func.log(new_log_path, f'Начало работы скрипта "3_update_excel.py" за дату: {str(start_play)}\n')

    # функция для архивации лог файлов
    m_func.zip_log_files(new_log_path, zip_log_dir, list_log_files_txt, log_dir)

    # функция для удаления лог файлов, которые до этого были добавлены в архив
    m_func.del_log_files(new_log_path, list_log_files_txt, log_dir)

    # работаем с Excel
    # проверяем, чтобы название Excel из my_config.py совпадало с названием Excel в папке
    # папка с Excel
    excel_folder = os.path.join(main_dir, 'excel')

    # находим самое короткое название файла, потому что к копиям Excel будет добавлена дата обновления
    # получаем список всех Excel файлов в папке
    xlsx_files = [i for i in os.listdir(excel_folder) if i.lower().endswith(('.xls', '.xlsx'))]

    # проверка, что Excel файл для обновления сущетсвует
    if len(xlsx_files) == 0:

        m_func.log(new_log_path, f'Отсутствует Excel для обновления. Проверяй.')

        # конец работы скрипта
        finish_play = datetime.datetime.now()

        m_func.log(new_log_path, f'\nВремя работы скрипта: {str(finish_play - start_play)}\n')
        m_func.log(new_log_path, f'**********************************************\n')
        
        # чтобы передать падение в вызываемый скрипт для падения самого скрипта
        sys.exit(1)

    else:

        original_excel = min(xlsx_files, key = len)

        # проверяем, что названия совпадают
        if original_excel != name_excel_for_update:
            m_func.log(new_log_path, f'Название Excel из my_config.py НЕ совпадает с названием Excel в папке. Проверяй. Работа скрипта "3_update_excel.py" остановлена.')

            # конец работы скрипта
            finish_play = datetime.datetime.now()

            m_func.log(new_log_path, f'\nВремя работы скрипта: {str(finish_play - start_play)}\n')
            m_func.log(new_log_path, f'**********************************************\n')
            
            # чтобы передать падение в вызываемый скрипт для падения самого скрипта
            sys.exit(1)

        else:

            # путь к исходному файлу Excel
            excel_path = os.path.join(main_dir, 'excel', name_excel_for_update)

            # получаем директорию исходного файла с Excel
            excel_dir_path = os.path.dirname(excel_path)

            # создаем имя для нового файла Excel по маске
            new_excel_filename = os.path.basename(excel_path).split('.')[0] + f'_upd_{today}.' + os.path.basename(excel_path).split('.')[1]

            # полный путь до копии файла Excel
            new_excel_path = os.path.join(excel_dir_path, new_excel_filename)

            # копируем файл Excel
            shutil.copy2(excel_path, new_excel_path)

            # путь до папки с архивами Excel внутри папки с Excel
            zip_excel_dir = os.path.join(excel_dir_path, 'zip-excel.zip')
                                        
            # список файлов с Excel
            list_excel_files_txt = os.listdir(excel_dir_path)

            m_func.log(new_log_path, f'Excel скопирован')
            m_func.log(new_log_path, f'Путь до файла:')
            m_func.log(new_log_path, f'{new_excel_path}\n')

            # функция для архивации Excel файлов
            m_func.zip_excel_files(new_log_path, list_excel_files_txt, excel_dir_path, zip_excel_dir)

            # функция для удаления Excel файлов, которые до этого были добавлены в архив
            m_func.del_excel_files(new_log_path, list_excel_files_txt, excel_dir_path)

            # список подключений, которые нужно обновить
            list_connections = config.list_connections

            # если список подключений в my_config.py пустой -> то создается список из всех существующих подключений в Excel
            if len(list_connections) == 0:
                list_connections = m_func.get_connections_excel(new_log_path, new_excel_path)
                m_func.log(new_log_path, f'Так как в файле my_config.py не задан список подключений -> будут обновлены все.')

            # фиксируем начало обновления подключений
            m_func.log(new_log_path, f'Начало обновления подключений: {str(datetime.datetime.today())}\n')

            # переменные для функции m_func.update_connect
            pause_after_upd = config.pause_after_upd # пауза, которая берется после обновления подключения
            long_list_connections = config.long_list_connections # список подключений, после которых требуется больше паузы
            long_pause_after_upd = config.long_pause_after_upd # увеличенная пауза после обновления покдлючений из списка long_list_connections
            pause_after_error = config.pause_after_error # пауза, которая берется после падения обновления Excel
            # главная функция по обновлению excel
            m_func.update_connect(list_connections, new_log_path, new_excel_path, pause_after_upd, long_list_connections, long_pause_after_upd, pause_after_error)

            # конец работы скрипта
            finish_play = datetime.datetime.now()

            m_func.log(new_log_path, f'\nВремя работы скрипта: {str(finish_play - start_play)}\n')
            m_func.log(new_log_path, f'**********************************************\n')

except Exception as e:
    m_func.log(new_log_path, f'Ошибка при отработке скрипта "3_update_excel.py": {str(e)}.\nЧитай файл с логами.\n')

    # чтобы передать падение в вызываемый скрипт для падения самого скрипта
    sys.exit(1)
                                                                
finally:
    # системное уведомление windows
    m_func.show_notification_windows(title = name_excel_for_update,
                                     message = 'Конец обновления'
                                     )