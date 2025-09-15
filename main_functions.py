# импорты
import win32com.client as win32 # для работы с приложениями windows
import pythoncom # для работы с приложениями windows
from plyer import notification # для поп-ап уведомлений в windows
import datetime # для работы с датами
import time # для задержек в выполнении скрипта
import os # для работы с директориями
import zipfile # для архивации файлов
import pandas as pd # для работы с таблицами
import sys # для проверки успешности выполнения другого скрипта

# импорт значений из файла my_config.py
import my_config as config

# название файла Excel, который нужно обновить
name_excel_for_update = config.name_excel_for_update

# пауза, которая берется на обновление почты в outlook, в секундах
outlook_wait_time = config.outlook_wait_time

# нужный аккаунт в Outlook
account_outlook = config.account_outlook.lower()


####################################################################################

# функция отображает системное уведомление
def show_notification_windows(title, message):
    """
    что делает:
        функция отображает системное уведомление
    аргументы:
        title : заголовок уведомления (можно прописать самостоятельно, изначально стоят значения по умолчанию)
        message : текст уведомления (можно прописать самостоятельно, изначально стоят значения по умолчанию)
    возвращает:
        всплывающее уведомление windows
    """
    # создаем объект уведомления
    notification.notify(
        title = title,
        message = message
        )

####################################################################################

# функция для логирования одновременно в консоль и в файл с сохранением
def log(new_log_path, message):
    """
    что делает:
        функция для логирования одновременно в консоль и в файл с сохранением
        идет сохранение после каждого инсерта данных
        это на случай, если программа вылетит из-за тяжелого датасета
        так как может багануть не само подключение, а сам редактор кода
    аргументы:
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
        message : указанный текст
    возвращает:
        вывод текста в консоль + запись текста в файл
    """
    with open(new_log_path, 'a', encoding = 'windows-1251') as f:
        f.write(message + '\n')
        f.flush() # сбрасывает буфер python в OC (еще не на диск)
        os.fsync(f.fileno()) # принудительно записывает файл на диск
    """
    принудительно закрывать txt файл с логами не нужно
    конструкцию new_log_path.close() можно не использовать
    """
    print(message)

####################################################################################

# функция, которая получает список лог файлов для удаления и их кол-во
def list_zip_log_files(list_log_files_txt, new_log_path):
    """
    что делает:
        функция, которая получает список лог файлов для удаления и их количество
        находим год и номер текущего месяца
        берем все файлы, в которых дата отличается от текущей
        формируем список таких файлов
    аргументы:
        list_log_files_txt : список файлов с логами txt (формируется динамически, исходя из структуры проекта и папок в нем)
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        list_for_zip : список файлов для архивации
        counter_zip_files : кол-во файлов для архивации
    """
    try:
        # год и номер текущего месяца
        month_rec = datetime.datetime.now().strftime('%Y-%m')
        # список для названия файлов для архивирования
        list_for_zip = []
        # вносим в список list_for_zip названия файлов, в названии которых нет текущего месяца
        for i in list_log_files_txt:
            # не учитываем файл 'zip-log-files.zip'
            # учитываем файлы с расширением .txt
            if ('zip-log-files.zip' not in i) and (i.lower().endswith('.txt')):
                # i.split('_')[-1] --> тут идет строгая привязка к названию файлов с логами strftime('%Y-%m-%d')
                if month_rec not in i.split('_')[-1]:
                    list_for_zip.append(i)
        # кол-во файлов для архивации
        counter_zip_files = len(list_for_zip)
        return list_for_zip, counter_zip_files
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции list_zip_log_files: {str(e)}\n')
        return [], 0 # чтобы дальше вызывающие функции не ломались

####################################################################################

# функция для архивации лог файлов
def zip_log_files(new_log_path, zip_log_dir, list_log_files_txt, log_dir):
    """
    что делает:
        функция для архивации лог файлов
        берем список файлов, архивируем их по очереди
    аргументы:
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
        zip_log_dir : путь до папки с архивами логов внутри палки с логами (формируется динамически, исходя из структуры проекта и папок в нем)
        list_log_files_txt : список файлов с логами txt (формируется динамически, исходя из структуры проекта и папок в нем)
        log_dir : путь до папки с логами txt (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        архивирует лог файлы
    """
    try:
        files_to_zip, count_to_zip = list_zip_log_files(list_log_files_txt, new_log_path)
        if count_to_zip > 0:
            with zipfile.ZipFile(zip_log_dir, mode = 'a', compression = zipfile.ZIP_DEFLATED, allowZip64 = True) as archive:
                for i in files_to_zip:
                    archive.write(os.path.join(log_dir, i), arcname = i)
            log(new_log_path, f'Занесено в архив: {count_to_zip} файлов с логами')
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции zip_log_files: {str(e)}\n')

####################################################################################

# функция для удаления лог файлов, которые до этого были добавлены в архив
def del_log_files(new_log_path, list_log_files_txt, log_dir):
    """
    что делает:
        функция для удаления лог файлов, которые до этого были добавлены в архив
        берем список файлов, удаляем их по очереди
    аргументы:
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
        list_log_files_txt : список файлов с логами txt (формируется динамически, исходя из структуры проекта и папок в нем)
        log_dir : путь до папки с логами txt (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        удаляет лог файлы
    """
    try:
        files_to_del, count_to_del = list_zip_log_files(list_log_files_txt, new_log_path)
        if count_to_del > 0:
            for i in files_to_del:
                os.remove(os.path.join(log_dir, i))
            log(new_log_path, f'Удалено {count_to_del} файлов с логами\n')
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции del_log_files: {str(e)}\n')

####################################################################################

# функция, которая получает список Excel файлов для удаления и их кол-во
def list_zip_excel_files(list_excel_files_txt, new_log_path, name_excel_for_update):
    """
    что делает:
        функция, которая получает список Excel файлов для удаления и их кол-во
        находим год и номер текущего месяца
        берем все файлы в которых дата отличается от текущей
        формируем список таких файлов
    аргументы:
        list_excel_files_txt : список файлов с Excel (формируется динамически, исходя из структуры проекта и папок в нем)
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        list_for_zip : список файлов для архивации
        counter_zip_files : кол-во файлов для архивации
    """
    try:
        # год и номер текущего месяца
        month_rec = datetime.datetime.now().strftime('%Y-%m')
        # список для названия файлов для архивирования
        list_for_zip = []
        # вносим в список list_for_zip названия файлов, в названии которых нет текущего месяца
        for i in list_excel_files_txt:
            # не учитываем файл 'zip-excel.zip'
            # учитываем файлы с расширением .xls и .xlsx
            if ('zip-excel.zip' not in i) and (name_excel_for_update not in i) and (i.lower().endswith(('.xls', '.xlsx'))):
                # i.split('_')[-1] --> тут идет строгая привязка к названию файлов с логами strftime('%Y-%m-%d')
                if month_rec not in i.split('_')[-1]:
                    list_for_zip.append(i)
        # кол-во файлов для архивации
        counter_zip_files = len(list_for_zip)
        return list_for_zip, counter_zip_files
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции list_zip_excel_files: {str(e)}\n')
        return [], 0 # чтобы дальше вызывающие функции не ломались

####################################################################################

# функция для архивации Excel файлов
def zip_excel_files(new_log_path, list_excel_files_txt, excel_dir_path, zip_excel_dir):
    """
    что делает:
        функция для архивации Excel файлов
        берем список файлов, архивируем их по очереди
    аргументы:
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
        list_excel_files_txt : список файлов с Excel (формируется динамически, исходя из структуры проекта и папок в нем)
        excel_dir_path : директория исходного файла с Excel (формируется динамически, исходя из структуры проекта и папок в нем)
        zip_excel_dir : путь до папки с архивами Excel внутри папки с Excel (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        архивирует Excel
    """
    try:
        files_for_zip, counter_zip_files = list_zip_excel_files(list_excel_files_txt, new_log_path, name_excel_for_update)
        if counter_zip_files > 0:
            with zipfile.ZipFile(zip_excel_dir, mode = 'a', compression = zipfile.ZIP_DEFLATED, allowZip64 = True) as archive:
                for i in files_for_zip:
                    archive.write(os.path.join(excel_dir_path, i), arcname=i)
            log(new_log_path, f'Занесено в архив: {counter_zip_files} файлов с Excel')
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции zip_excel_files: {str(e)}\n')

####################################################################################

# функция для удаления Excel файлов, которые до этого добавлены в архив
def del_excel_files(new_log_path, list_excel_files_txt, excel_dir_path):
    """
    что делает:
        функция для удаления Excel файлов, которые до этого добавлены в архив
        берем список файлов, удаляем их по очереди
    аргументы:
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
        list_excel_files_txt : список файлов с Excel (формируется динамически, исходя из структуры проекта и папок в нем)
        excel_dir_path : директория исходного файла с Excel (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        удаляет Excel
    """
    try:
        files_to_del, count_to_del = list_zip_excel_files(list_excel_files_txt, new_log_path, name_excel_for_update)
        if count_to_del > 0:
            for i in files_to_del:
                os.remove(os.path.join(excel_dir_path, i))
            log(new_log_path, f'Удалено {count_to_del} файлов с Excel\n')
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции del_log_files: {str(e)}\n')

####################################################################################

# главная функция по обновлению Excel
def update_connect(list_connections, new_log_path, new_excel_path, pause_after_upd, long_list_connections, long_pause_after_upd, pause_after_error):
    """
    что делает:
        обновляет подключение Excel
        1) открывает Excel
        2) обновляет по очереди подключения Excel
        3) сохраняет и закрывает Excel
    аргументы:
        list_connections : списко подключений для обновления (задается в my_config.py)
        new_log_path : путь до файла с логами (формируется динамически, исходя из структуры проекта и папок в нем)
        new_excel_path : путь до Excel файла, который нужно обновить (формируется динамически, исходя из структуры проекта и папок в нем)
        pause_after_upd : пауза, которая берется после обновления подключения (задается в my_config.py)
        long_list_connections : список подключений, после которых требуется больше паузы (задается в my_config.py)
        long_pause_after_upd : увеличенная пауза после обновления подключений из списка long_list_connections (задается в my_config.py)
        pause_after_error : пауза, которая берется после падения обновления Excel (задается в my_config.py)
    возвращает:
        обновляет подключение Excel
    ------------------------------
    отличие Dispatch от DispatchEx:
    Dispatch : используется для получения ссылки на уже существующий объект COM
    если объект уже запущен в другом процессе, то Dispatch подключается к этому существующему процессу
    DispatchEx : используется для получения ссылки на объект COM, создавая новый экземпляр приложения
    -) использовать Dispatch, когда нужно подключиться к уже работающему приложению или получить доступ к объекту, который уже существует
    -) использовать DispatchEx, когда нужно запустить новое приложение или получить доступ к объекту, даже если приложение еще не запущено
    sys.exit(1) нужно, чтобы передать падение функции в вызываемый скрипт для падения самого скрипта
    """
    try:
        # оперделяем переменные
        excel, excel_wb = None, None
        # делаем новый поток СОМ для этого скрипта (важно для многопоточности)
        pythoncom.CoInitialize()
        # открываем Excel
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False # False, чтобы НЕ видеть Excel
        # excel.Visible = True # True, чтобы видеть Excel
        excel.DisplayAlerts = False # False, чтобы НЕ видеть всплывающие диалоги Excel
        # excel.DisplayAlerts = True # True, чтобы видеть всплывающие диалоги Excel
        # указываем путь до файла Excel
        excel_wb = excel.Workbooks.Open(new_excel_path)
        # выбираем список подключений, которые нужно обновить
        connections_to_update = list_connections
        # создаем копию списка для последующей фильтрации (рабочий список)
        connections_process_update = connections_to_update.copy()
        # список подключений, которые успешно обновились
        connections_done_update = []
        # счетчик количества подключений, которые планируется обновить
        count_connect = len(connections_to_update)
        # счетчик количества обновленных подключений
        counter_connect = 0
        # итерация по копии списка подключений, чтобы безопасно удалять элементы
        for con_name in connections_process_update.copy():
            # счетчик попыток обновления у подключений
            attempt = 1
            # флаг успешного обновления подключения
            success = False        
            # даем максимум 2 попытки обновить подключение
            while attempt <= 2:
                try:
                    log(new_log_path, f'Обновление подключения: {con_name} | Попытка: {attempt}')
                    # начало обновления подключения
                    start_task = datetime.datetime.now()
                    # ищем подключение по имени в файле
                    connection = None
                    for conn in excel_wb.Connections:
                        if conn.Name == con_name:
                            connection = conn
                            break
                    # если подключение не найдено - выходим из цикла
                    if not connection:
                        log(new_log_path, f'Подключение {con_name} НЕ найдено в Excel')
                        break
                    # обновление подключения
                    connection.Refresh()
                    # сохраняем изменения
                    excel_wb.Save()
                    log(new_log_path, f'Сохранение Excel после обновления подключения: {con_name}')
                    # конец работы таска
                    end_task = datetime.datetime.now()
                    # лог успешного обновления
                    log(new_log_path, f'Подключение "{con_name}" успешно обновлено')
                    log(new_log_path, f'Время работы обновления "{con_name}" : {str(end_task - start_task)}\n')
                    # добавляем это подключение в список успешно обновленных
                    connections_done_update.append(con_name)
                    # увеличиваем счетчик успешно обновленных подключений
                    counter_connect += 1
                    # меняем флаг успешного обновления
                    success = True
                    # делаем паузу взависимости от названия подключения
                    time.sleep(long_pause_after_upd if con_name in long_list_connections else pause_after_upd)
                    # выходим из попыток обновления и переходим к следующему обновлению
                    break
                except Exception as e:
                    log(new_log_path, f'Ошибка при обновлении подключения "{con_name}" на попытке {attempt}: {str(e)}')
                    # увеличиваем счетчик попыток обновления у подключений
                    attempt += 1
                    # пауза после падения обновления подключения
                    time.sleep(pause_after_error)
            # если не удалось обновить подключение после 2ух попыток
            if not success:
                log(new_log_path, f'Подключение "{con_name}" не обновилось после 2ух попыток')
                # удаляем подключение из рабочего списка
                connections_process_update.remove(con_name)
                log(new_log_path, f'Переходим к обновлению следующего подключения\n')
        # после окончания цикла проверяем статус всех подключений
        if counter_connect == count_connect:
            log(new_log_path, f'Все подключения успешно обновлены')
        else:
            log(new_log_path, f'Некоторые подключения не обновились')
            log(new_log_path, f'Планировалось обновить: {count_connect}')
            log(new_log_path, f'Успешно обновлено: {counter_connect}')
            log(new_log_path, f'He удалось обновить: {set(connections_to_update) - set(connections_done_update)}\n')
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции update_connect: {str(e)}')
        sys.exit(1)
    finally:
        if excel_wb is not None:
            # сохраняем изменения
            excel_wb.Save()
            log(new_log_path, f'Финальное сохранение Excel')
            # закрываем книгу и Excel.
            excel_wb.Close(SaveChanges = True)
        if excel is not None:
            # закрываем Excel
            excel.Quit()
        # завершаем работу СОМ для текущего потока
        pythoncom.CoUninitialize()
        log(new_log_path, f'Excel закрыт')
        log(new_log_path, f'Конец работы функции update_connect')

####################################################################################

# функция для получения списка всех подключений в Excel
def get_connections_excel(new_log_path, new_excel_path):
    """
    что делает:
        функция для получения списка всех подключений в Excel
        будет срабатывать, если в my_config.py не указан список подключений руками
    аргументы:
        new_log_path : путь до файла с логами (формируется динамически, исходя из структуры проекта и папок в нем)
        new_excel_path : путь до Excel файла, который нужно обновить (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        list_connections : вернет список подключений
    """
    try:
        list_connections = []
        # оперделяем переменные
        excel, excel_wb = None, None
        pythoncom.CoInitialize()
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel_wb = excel.Workbooks.Open(new_excel_path)
        connections = excel_wb.Connections
        if connections.Count > 0:
            for i in connections:
                list_connections.append(i.Name)
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции get_connections_excel: {str(e)}\n')
        sys.exit(1)
    finally:
        if excel_wb is not None:
            # сохраняем изменения
            excel_wb.Save()
            log(new_log_path, f'Финальное сохранение Excel')
            # закрываем книгу и Excel.
            excel_wb.Close(SaveChanges = True)
        if excel is not None:
            # закрываем Excel
            excel.Quit()
        # завершаем работу СОМ для текущего потока
        pythoncom.CoUninitialize()
        return list_connections

####################################################################################

# функция ищет корневую папку указанного аккаунта в Outlook
def get_outlook_account(namespace, account_name, new_log_path):
    """
    что делает:
        ищет корневую папку конкретного аккаунта в Outlook
    аргументы:
        namespace : outlook.GetNamespace('MAPI') (задается в функциях с Outlook)
        account_name : адрес или имя нужного аккаунта (задается в my_config.py)
        new_log_path : путь до лог файла (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        объект папки аккаунта или None
    """
    try:
        for i in range(1, namespace.Folders.Count + 1):
            folder = namespace.Folders.Item(i)
            if folder.Name.lower() == account_name.lower():
                return folder
        log(new_log_path, f'Аккаунт {account_name} не найден в Outlook')
        return None
    except Exception as e:
        log(new_log_path, f'Ошибка get_outlook_account: {str(e)}\n')
        return None

####################################################################################

# функция обновляет все письма в Outlook (синхронизирует папки и письма в них)
def update_outlook_mail(new_log_path, outlook_wait_time):
    """
    что делает:
        обновляет все письма в Outlook (синхронизирует папки и письма)
        особенно необходимо, если подключена почта через IMAP (например, Яндекс)
        так как при выключенном Outlook почта не синхронизируется
        сам объект Outlook в этой функции тоже нужно создавать, так как именно тут идет запуск синхронизации 
    аргументы:
        new_log_path : путь до лог файла (формируется динамически, исходя из структуры проекта и папок в нем)
        outlook_wait_time : кол-во секунд на обновление (задается в my_config.py)
    возвращает:
        ничего, делает синхронизацию
    """
    try:
        # создаем обект приложения Outlook
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        # делаем поиск только по нужному аккаунту
        account = get_outlook_account(namespace, account_outlook, new_log_path)
        if not account:
            raise ValueError(f"Аккаунт {account_outlook} не найден в Outlook")
        # выбираем SyncObjects, фильтруя по имени аккаунта
        sync_objects = namespace.SyncObjects
        if sync_objects.Count > 0:
            # берем первый SyncObject (обычно он синхронизирует все папки)
            sync = sync_objects.Item(1)
            # запуск синхронизации
            sync.Start()
            time.sleep(outlook_wait_time)           
        else:
            log(new_log_path, 'SyncObjects не найдено, синхронизация невозможна. Возможные причины:\n\
                1. Учетная запись подключена только через POP\n\
                2. Учетная запись еще не настроена или отключена\n\
                3. Особый режим запуска Outlook, который не создает объекты SyncObjects\n\
                В этом случае новые письма будут доступны только после ручного обновления в Outlook\n')
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции update_outlook_mail: {str(e)}\n')

####################################################################################

# функция для отправки письма в Outlook
def send_mail_outlook(recipient_mail, subject_mail, body_mail, attach_files, account_outlook, new_log_path):
    """
    что делает:
        отправка письма в Outlook
    аргументы:
        recipient_mail : указываем список получателей (задается в my_config.py)
        subject_mail : тема письма (можно прописать самостоятельно, изначально стоят значения по умолчанию)
        body_mail : текст письма (можно прописать самостоятельно, изначально стоят значения по умолчанию)
        attach_files : вложения письма (можно прописать самостоятельно, изначально стоят значения по умолчанию)
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
        account_outlook : email нужного аккаунта в Outlook (задается в my_config.py)
    возвращает:
        ничего, отправляет письмо в Outlook
    """
    try:
        # сначала обновляем все письма в outlook
        # особенно необходимо, если подключена почта через IMAP (например, Яндекс)
        update_outlook_mail(new_log_path, outlook_wait_time)
        # создаем объект приложения Outlook
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        # находим нужный аккаунт именно в Accounts (а не в Folders)
        account = None
        for i in namespace.Accounts:
            if i.DisplayName.lower() == account_outlook.lower() or i.SmtpAddress.lower() == account_outlook.lower():
                account = i
                break
        if not account:
            raise ValueError(f'Аккаунт {account_outlook} не найден среди namespace.Accounts')
        # создаем новый элемент письма
        mail = outlook.CreateItem(0)
        # указываем список получателей
        mail.To = '; '.join(recipient_mail).lower()
        # тема письма
        mail.Subject = subject_mail
        # текст письма
        mail.Body = body_mail
        # указываем аккаунт отправителя
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  
        # вложения письма
        for j in attach_files:
            mail.Attachments.Add(j)
        # отправляем письмо
        mail.Send()
        # пауза для отправки письма в outlook 
        # особенно необходимо, если подключена почта через IMAP (например, Яндекс)
        update_outlook_mail(new_log_path, outlook_wait_time)
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции send_mail_outlook: {str(e)}\n')

####################################################################################

# рекурсивный поиск папки во всем каталоге Outlook (для одного пользователя)
def find_folder_recursive(folder, target_name, new_log_path):
    """
    что делает:
        рекурсивно ищет папку с именем target_name внутри данной папки folder
        проходит по всем вложенным подпапкам и возвращает найденную папку или None
    аргументы:
        folder : текущая папка, внутри которой выполняется поиск (точка входа в рекурсию) (задается в процессе работы функций с Outlook)
        target_name : имя папки, которую нужно найти (передается переменная mail_folder, а она задается в my_config.py)
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        найденную папку или None
    """
    try:
        for subfolder in folder.Folders:
            if subfolder.Name.lower() == target_name.lower():
                return subfolder
            # рекурсивно ищем в подпапках
            found = find_folder_recursive(subfolder, target_name, new_log_path)
            if found:
                return found
        return None
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции find_folder_recursive: {str(e)}\n')

####################################################################################

# функция для поиска папки в Outlook
# учитывает случаи, если работа ведется не только в корпоратиыном Outlook
# а например, авторизация яндекс почты в Outlook
# в таком случае, нужная папка будет не во "Входящих", а в любом другом месте (из-за ограничения сервера IMAP)
# работает только для одного пользователя
def find_folder_outlook(mail_folder, new_log_path):
    """
    что делает:
        ищет папку с именем mail_folder во всем дереве папок почтового ящика Outlook
        начинает поиск с корневой папки почты и вызывает рекурсивный обход
        возвращает объект папки или None, если папка не найдена
    аргументы:
        mail_folder : название папки, в которой нужно искать письмо (задается в my_config.py)
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        объект папки если она есть, если нет -> то None
    """
    try:
        # сначала обновляем все письма в outlook
        # особенно необходимо, если подключена почта через IMAP (например, Яндекс)
        update_outlook_mail(new_log_path, outlook_wait_time)
        # outlook = win32.DispatchEx('Outlook.Application')
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        # делаем поиск только по нужному аккаунту
        root_folder = get_outlook_account(namespace, account_outlook, new_log_path)
        if not root_folder:
            raise ValueError(f"Аккаунт {account_outlook} не найден в Outlook") 
        # пробуем найти папку рекурсивно по всему дереву папок
        found_folder = find_folder_recursive(root_folder, mail_folder, new_log_path)
        return found_folder
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции find_folder_outlook: {str(e)}\n')

####################################################################################

# функция проверяет наличие письма в папке Outlook
def find_mail_outlook(text_in_mail, messages, new_log_path):
    """
    что делает:
        функция проверяет наличие письма в папке Outlook по условиям:
        1) в теле письма должны содержаться все подстроки из text_in_mail
        2) письмо должно быть за текущую дату
    аргументы:
        text_in_mail : текст, который нужно искать в письме (задается в my_config.py)
        messages : переменная со всеми письмами в папке Outlook (задается в процессе работы функций с Outlook)
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        True, если письмо найдено, иначе False.
    """
    try:
        # сначала обновляем все письма в outlook
        # особенно необходимо, если подключена почта через IMAP (например, Яндекс)
        update_outlook_mail(new_log_path, outlook_wait_time)
        # переменная наличия нужного письма
        found_mail = False
        # проходимся по письмам в заданной папке и ищем нужное письмо по маске
        for message in messages:
            # ищем все совпадения в одном письме + письмо должно быть за текущую дату
            if all(i.lower() in message.Body.lower() for i in text_in_mail) \
                and datetime.date.today().strftime('%Y-%m-%d') in message.SentOn.strftime('%Y-%m-%d'):
                # после того как письмо нашлось -> поиск останавливается
                found_mail = True
                break
        return found_mail
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции find_mail_outlook: {str(e)}\n')

####################################################################################

# функция для создании CSV таблицы check_update_excel.csv
def create_update_excel_today(name_excel_for_update, path_csv, new_log_path):
    """
    что делает:
        функция для создании CSV таблицы check_update_excel.csv
    аргументы:
        name_excel_for_update : название файла Excel, который нужно обновить (задается в my_config.py)
        path_csv : CSV файл, в который вставляется DТ, когда обновлялся Excel (формируется динамически, исходя из структуры проекта и папок в нем)
        new_log_path : путь до нового сформированного txt с логами (формируется динамически, исходя из структуры проекта и папок в нем)
    возвращает:
        если CSV уже есть -> ничего, если нет CSV -> создает его
    """
    try:
        if os.path.exists(path_csv):
            pass
        else:
            df = pd.DataFrame(columns = ['row_number', 'name_excel', 'update_dt'])
            # заполняем дефолтными данным
            next_row = {'row_number' : 1,
                        'name_excel' : name_excel_for_update,
                        'update_dt' : '1970-01-01'
                        }
            df = pd.concat([df, pd.DataFrame([next_row])], ignore_index = True)
            df.to_csv(path_csv, index = False, encoding = 'windows-1251')
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции create_update_excel_today: {str(e)}\n')

####################################################################################

# функция для вставки данных в таблицу check_update_excel.csv об обновлении за текущую дату
def insert_update_excel_today(name_excel_for_update, path_csv, new_log_path):
    """
    что делает:
        функция для вставки данных в таблицу check_update_excel.csv об обновлении за текущую дату
    аргументы:
        name_excel_for_update : название файла Excel, который нужно обновить (задается в my_config.py)
        path_csv : CSV файл, в который вставляется DТ, когда обновлялся Excel (формируется динамически, исходя из структуры проекта и папок в нем)
        new_log_path : путь до нового сформированного txt с логами  (формируется динамически, исходя из структуры проекта и папок в нем)
        название файла : check_update_excel.csv (задается в my_config.py)
        поля:
            row_number : порядковый номер записи
            name_excel : название excel файла, берется из переменной name_excel_for_update
            update_dt : дата вставки данных в таблицу
    возвращает:
        добавляет запись в CSV
    """
    try:
        # filename = *check_update_excel.csv’
        now = datetime.datetime.now().strftime('%Y-%m-%d')
        if os.path.exists(path_csv):
            # если файл существует -> читаем ero
            df = pd.read_csv(path_csv, encoding = 'windows-1251')
            # получаем следующий индекс
            next_index = df['row_number'].max() + 1
        else:
            # если файла нет -> то создаем его
            create_update_excel_today(path_csv)
            next_index = 1
        # добавляем новую запись
        next_row = {'row_number' : next_index,
                    'name_excel' : name_excel_for_update,
                    'update_dt' : now
                    }
        # объединяем с прошлыми записями
        df = pd.concat([df, pd.DataFrame([next_row])], ignore_index = True)
        # сохраняем обратно в CSV
        df.to_csv(path_csv, index = False, encoding = 'windows-1251')
    except Exception as e:
        log(new_log_path, f'Ошибка в работе функции insert_update_excel_today: {str(e)}\n')

####################################################################################