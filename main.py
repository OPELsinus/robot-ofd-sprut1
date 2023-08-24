import datetime
import os
import shutil
import time

import Levenshtein
import pandas as pd
from time import sleep

import win32com.client as win32
import psycopg2 as psycopg2
from pywinauto import keyboard

from config import logger, download_path, robot_name, db_host, db_port, db_name, db_user, db_pass, tg_token, chat_id, smtp_host, smtp_author, jadyra_path, ardak_path, mapping_path
from core import Sprut
from tools.clipboard import clipboard_get
from tools.smtp import smtp_send
from tools.tg import tg_send

months = [
    'Январь',
    'Февраль',
    'Март',
    'Апрель',
    'Май',
    'Июнь',
    'Июль',
    'Август',
    'Сентябрь',
    'Октябрь',
    'Ноябрь',
    'Декабрь'
]


def sql_create_table():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
        CREATE TABLE IF NOT EXISTS ROBOT.{robot_name.replace("-", "_")} (
            started_time timestamp,
            ended_time timestamp,
            store_name text,
            short_name text,
            status text,
            responsible text,
            found_difference text,
            count int,
            error_reason text,
            error_saved_path text,
            execution_time text
            )
        '''
    c = conn.cursor()
    c.execute(table_create_query)

    conn.commit()
    c.close()
    conn.close()


def sql_delete_table():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
        DROP TABLE ROBOT.{robot_name.replace("-", "_")}'''
    c = conn.cursor()
    c.execute(table_create_query)

    conn.commit()
    c.close()
    conn.close()


def insert_data_in_db(started_time, store_name, short_name, status, responsible, found_difference, count, error_reason, error_saved_path, execution_time):
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)

    query_delete = f"""
        delete from ROBOT.{robot_name.replace("-", "_")} where store_name = '{store_name}'
    """

    query = f"""
        INSERT INTO ROBOT.{robot_name.replace("-", "_")}
        (started_time, ended_time, store_name, short_name, status, responsible, found_difference, count, error_reason, error_saved_path, execution_time)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """

    values = (
        started_time.strftime("%d.%m.%Y %H:%M:%S.%f"),
        datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"),
        str(store_name),
        str(short_name),
        str(status),
        str(responsible),
        str(found_difference),
        str(count),
        str(error_reason),
        str(error_saved_path),
        str(execution_time)
    )

    cursor = conn.cursor()

    conn.autocommit = True
    try:
        cursor.execute(query_delete)
    except Exception as e:
        logger.info(f'GOVNO {e}')
        pass

    try:
        cursor.execute(query, values)

    except Exception as e:
        conn.rollback()
        logger.info(f"Error: {e}")

    conn.commit()

    cursor.close()
    conn.close()


def get_all_data():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            order by started_time desc
            '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())
    df1.columns = ['started_time', 'ended_time', 'store_name', 'short_name', 'status', 'responsible', 'found_difference', 'count', 'error_reason', 'error_saved_path', 'execution_time']

    cur.close()
    conn.close()

    return df1


def write_branches_in_their_big_excels(end_date_):
    logger.info('Started write_branches_in_their_big_excels')

    # ? Create new page
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    end_date_ = datetime.datetime.strptime(end_date_, '%d.%m.%Y')
    year = str(end_date_.year)
    month = end_date_.month

    def open_excel(path):
        found = False
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(path)
        last_sheet_index = wb.Worksheets.Count
        ws = None

        for sheet in wb.Worksheets:
            if months[month - 1].lower() in str(sheet.Name).lower() and (year in str(sheet.Name) or year[2:] in str(sheet.Name)):
                logger.info(sheet.Name)
                ws = wb.Worksheets(sheet.Name)
                found = True

        if not found:
            ws = wb.Worksheets.Add(Before=wb.Worksheets(last_sheet_index))

            ws.Name = months[month - 1] + ' ' + year

            ws1 = wb.Worksheets(last_sheet_index - 2)

            ws1.Range('A1:AA2').Copy()
            ws.Range('A1').PasteSpecial()

            header_range = ws.Range('A1:Z2')
            header_range.AutoFilter(Field=1)
            header_range.AutoFilter(Field=2)

        return [wb, ws]

    def check_one_branch(ws2, short_name, single_branch):

        at_least_one_found = False
        count = 0

        for ind in range(11, 1000):
            if 'Итого' in str(df.loc[ind, '№ Кассы']):
                break

            copy = False
            if df.loc[ind, 'Нал Разница'] != 0 or df.loc[ind, 'Безнал Разница'] or df.loc[ind, 'Разница']:
                copy = True
            if df.loc[ind, 'ООФД - Z - отчет - СПРУТ'] is not True:
                copy = True

            if copy:
                at_least_one_found = True
                count += 1

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False

                wb0 = excel.Workbooks.Open(os.path.join(os.path.join(download_path, 'reports'), single_branch))
                ws0 = wb0.Worksheets(1)

                empty_row = ws2.Cells.SpecialCells(win32.constants.xlCellTypeLastCell).Row + 1

                ws0.Range(f'A{ind + 2}:W{ind + 2}').Copy()
                ws2.Range(f'C{empty_row}').PasteSpecial()

                ws2.Cells(empty_row, 1).Value = end_date_.strftime('%d.%m.%Y')
                ws2.Columns(1).ColumnWidth = 15

                ws2.Cells(empty_row, 2).Value = short_name

                ws2.Cells(empty_row, 1).HorizontalAlignment = win32.constants.xlCenter
                ws2.Cells(empty_row, 1).VerticalAlignment = win32.constants.xlCenter

                ws2.Cells(empty_row, 2).HorizontalAlignment = win32.constants.xlCenter
                ws2.Cells(empty_row, 2).VerticalAlignment = win32.constants.xlCenter

                wb0.Close(False)

        return [ws2, at_least_one_found, count]

    baishukova_wb, baishukova_ws = open_excel(ardak_path)
    nusipova_wb, nusipova_ws = open_excel(jadyra_path)

    df1 = pd.read_excel(mapping_path)

    # ? Проверка каждого экселя на наличие расхождений
    for branch in os.listdir(os.path.join(download_path, 'reports')):

        df = pd.read_excel(os.path.join(os.path.join(download_path, 'reports'), branch))

        df.columns = ['№ Кассы', 'Регистр. № кассы', '№', 'Итог продаж', 'Возвраты: (нал,безнал, бонус)', 'Возвраты Бонусы', 'Итого за минусом возвратов:', 'безнал', 'Итого наличных', 'итого наличных', 'Сертификаты подаренные', 'Сертификаты, реализованные частным лицам', 'Сертификаты, реализованные юр/ лицам', 'Сертификаты, созданные при возврате товара', 'Чеки по акции "Счастливый чек" (Бесплатные чеки)', 'Нехватка разменных монет', 'Оплата Бонусами', 'безнал', 'Итого продаж', 'Нал Разница', 'Безнал Разница', 'Разница', 'ООФД - Z - отчет - СПРУТ']

        title = df['Регистр. № кассы'].iloc[5]
        try:
            started_time = datetime.datetime.now()
            start_time = time.time()
            if 'nusipova' in str(df1[df1['Название филиала в Спруте'] == title]['Сотрудник'].iloc[0]).lower():
                short_name = df1[df1['Название филиала в Спруте'] == title]['Короткое название филиала'].iloc[0]
                try:
                    nusipova_ws, found, count = check_one_branch(nusipova_ws, short_name, branch)
                    end_time = time.time()
                    insert_data_in_db(started_time, branch, short_name, 'success', 'Nusipova', found, count, '', '', str(end_time - start_time))
                except Exception as ex:
                    end_time = time.time()
                    tg_send(f'FAILED: {short_name} | ({branch})', bot_token=tg_token, chat_id=chat_id)
                    insert_data_in_db(started_time, branch, short_name, 'failed', 'Nusipova', '', 0, str(ex), '', str(end_time - start_time))

            else:
                short_name = df1[df1['Название филиала в Спруте'] == title]['Короткое название филиала'].iloc[0]

                try:
                    baishukova_ws, found, count = check_one_branch(baishukova_ws, title, branch)
                    end_time = time.time()
                    insert_data_in_db(started_time, branch, short_name, 'success', 'Baishukova', found, count, '', '', str(end_time - start_time))
                except Exception as ex:
                    end_time = time.time()
                    tg_send(f'FAILED: {short_name} | ({branch})', bot_token=tg_token, chat_id=chat_id)
                    insert_data_in_db(started_time, branch, short_name, 'failed', 'Baishukova', '', 0, str(ex), '', str(end_time - start_time))

        except Exception as error:
            logger.info(error)
    print('Finishing1')
    logger.info('Finishing1')
    empty_row = baishukova_ws.Cells.SpecialCells(win32.constants.xlCellTypeLastCell).Row
    baishukova_ws.Cells(empty_row, 1).EntireRow.Interior.ColorIndex = 40

    empty_row = nusipova_ws.Cells.SpecialCells(win32.constants.xlCellTypeLastCell).Row
    nusipova_ws.Cells(empty_row, 1).EntireRow.Interior.ColorIndex = 40

    baishukova_wb.Save()
    baishukova_wb.Close()

    nusipova_wb.Save()
    nusipova_wb.Close()
    excel.Application.Quit()
    logger.info('Finishing2')
    print('Finishing2')


def send_in_cache(sprut, today):
    sprut.open("Контроль передачи данных", switch=False)

    logger.info('Switching')
    sprut.parent_switch({"title_re": ".Контроль передачи данных.", "class_name": "Tcontrolcache_fm_main",
                         "control_type": "Window", "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()
    logger.info('Switched')
    sprut.find_element({"title": "Журналы", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 2}).click()

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    keyboard.send_keys("{TAB}")
    keyboard.send_keys("{TAB}")
    keyboard.send_keys("^%{F11} ")

    sprut.parent_switch({"title": "Выборка по запросу", "class_name": "Tvms_modifier_fm_builder", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()

    sprut.find_element({"title": "Дата чека", "class_name": "", "control_type": "ListItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()
    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(today)

    sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()
    sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()

    logger.info('Clicked list')
    keyboard.send_keys("{DOWN}" * 4)
    keyboard.send_keys("{ENTER}")

    logger.info('Clicked item')
    sprut.find_element({"title": "", "class_name": "TvmsBitBtn", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "И", "class_name": "TvmsBitBtn", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()
    logger.info('KEKUS')
    sprut.find_element({"title": "Торговая площадка", "class_name": "", "control_type": "ListItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()
    sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys("{F5}")

    sprut.parent_switch({"title": "Поиск", "class_name": "Tvms_search_fm_builder", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()

    branches = ['Торговый зал_ОПТ АФ №55', 'Торговый зал_ОПТ ШФ №35', 'ТОРГ']

    for branch in branches:
        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('^N')

        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(f'%{branch}%', sprut.keys.ENTER, protect_first=True)

        sprut.find_element({"title": "", "class_name": "", "control_type": "Button",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).click()

    sprut.find_element({"title": "Выбрать", "class_name": "TvmsBitBtn", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()
    sprut.parent_back(1)

    while True:
        try:
            sprut.find_element({"title": "", "class_name": "TvmsBitBtn", "control_type": "Button",
                                "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=1).click()
            break
        except:
            pass
    logger.info('clicked')
    while True:
        try:
            sprut.find_element({"title": "Ввод", "class_name": "TvmsBitBtn", "control_type": "Button",
                                "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1).click()
            break
        except:
            pass
    logger.info('clicked1')

    sprut.parent_back(1)

    sprut.find_element({"title": " ", "class_name": "TPanel", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(20, 17))

    sprut.find_element({"title": "Журналы", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 2}).click()

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()

    sprut.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.parent_back(1)


def create_z_reports(branches, start_date, end_date):

    for branch in branches[::]:
        for i in range(5):
            try:
                logger.info(f'Started {branch}')

                sprut = Sprut("MAGNUM")
                sprut.run()

                sprut.open("Отчеты")

                keyboard.send_keys("{F5}")

                sprut.parent_switch({"title": "Поиск", "class_name": "Tvms_search_fm_builder", "control_type": "Window",
                                     "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                sprut.find_element({"title": "Название отчета", "class_name": "TvmsComboBox", "control_type": "Pane",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                keyboard.send_keys("{UP}" * 4)
                keyboard.send_keys("{ENTER}")
                # sprut.find_element({"title": "", "class_name": "", "control_type": "ListItem",
                #                     "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('3303')

                keyboard.send_keys("{ENTER}")

                sprut.find_element({"title": "Перейти", "class_name": "TvmsBitBtn", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                # ? ---------------------------------------------------------
                # sprut.get_pane(1).type_keys(sprut.Keys.F9)

                sprut.parent_back(1)


                sprut.find_element({"title": "", "class_name": "", "control_type": "SplitButton",
                                    "visible_only": True, "enabled_only": True, "found_index": 4}).click()

                sprut.parent_switch({"title": "N100912-Сверка Z отчётов и оборота Спрут", "class_name": "TfrmParams", "control_type": "Window",
                                     "visible_only": True, "enabled_only": True, "found_index": 0})

                sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 1}).click(double=True)
                keyboard.send_keys("{BACKSPACE}" * 20)
                sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(start_date)

                sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).click(double=True)
                keyboard.send_keys("{BACKSPACE}" * 20)
                sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(end_date)

                # ? Search for 1 branch
                sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 1}).click()
                sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 1}).click(double=True)

                if sprut.wait_element({"title": "Отчеты", "class_name": "#32770", "control_type": "Window",
                                       "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=8):
                    sprut.find_element({"title": "ОК", "class_name": "Button", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys("{F5}")

                sprut.parent_switch({"title": "Поиск", "class_name": "Tvms_search_fm_builder", "control_type": "Window",
                                     "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()

                try:
                    sprut.find_element({"title": "", "class_name": "", "control_type": "Button",
                                        "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=10).click()
                except:
                    pass

                sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('^N')

                # branch = 'Алматинский филиал №1 ТОО "Magnum Cash&Carry"'

                sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(f'%{branch}%', sprut.keys.ENTER, protect_first=True)

                sprut.find_element({"title": "", "class_name": "", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 1}).click()

                sprut.find_element({"title": "Выбрать", "class_name": "TvmsBitBtn", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).click()
                sprut.parent_back(1)

                sprut.find_element({"title": "Ввод", "class_name": "TvmsFooterButton", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).click(double=True)

                sleep(1)

                if sprut.wait_element({"title": "Ввод", "class_name": "TvmsFooterButton", "control_type": "Button",
                                       "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=3):
                    sprut.find_element({"title": "Ввод", "class_name": "TvmsFooterButton", "control_type": "Button",
                                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

                wait_loading(branch)

                # sprut.parent_back(1).set_focus()

                sprut.quit()

                print('Finished branch')
                logger.info('Finished branch')

                break

            except Exception as exc:
                logger.info(f'Error occured at {branch}: {exc}')
                sleep(10)

    print('-----------------------------------------------------------------------')
    print('Finished CREATING Z REPORTS')
    print('-----------------------------------------------------------------------')

    logger.info('Finished CREATING Z REPORTS')


def wait_loading(branch):
    print('Started loading')
    logger.info('Started loading')
    branch = branch.replace('.', '').replace('"', '')
    found = False
    while True:
        for file in os.listdir(download_path):
            sleep(.1)
            creation_time = os.path.getctime(os.path.join(download_path, file))
            current_time = datetime.datetime.now().timestamp()
            time_difference = current_time - creation_time
            days_since_creation = time_difference / (60 * 60 * 24)

            if int(days_since_creation) <= 1 and file[0] != '$' and '.' in file and 'xl' in file and '100912' in file:
                logger.info(file)
                type = '.' + file.split('.')[1]
                shutil.move(os.path.join(download_path, file), os.path.join(os.path.join(download_path, 'reports'), branch + type))
                found = True
                break
        if found:
            break
    print('Finished loading')
    logger.info('Finished loading')


def get_all_existing_branches_from_sprut(sprut):
    # ? Collecting all existing branches from the sprut

    sprut.open("Отчеты")

    keyboard.send_keys("{F5}")

    sprut.parent_switch({"title": "Поиск", "class_name": "Tvms_search_fm_builder", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "Название отчета", "class_name": "TvmsComboBox", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    keyboard.send_keys("{UP}" * 4)
    keyboard.send_keys("{ENTER}")
    # sprut.find_element({"title": "", "class_name": "", "control_type": "ListItem",
    #                     "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('3303')

    keyboard.send_keys("{ENTER}")

    sprut.find_element({"title": "Перейти", "class_name": "TvmsBitBtn", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    # ? ---------------------------------------------------------
    # sprut.get_pane(1).type_keys(sprut.Keys.F9)

    sprut.parent_back(1)

    sprut.find_element({"title": "", "class_name": "", "control_type": "SplitButton",
                        "visible_only": True, "enabled_only": True, "found_index": 4}).click()

    sprut.parent_switch({"title": "N100912-Сверка Z отчётов и оборота Спрут", "class_name": "TfrmParams", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0})

    # ? Search for 1 branch
    sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()
    sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click(double=True)

    if sprut.wait_element({"title": "Отчеты", "class_name": "#32770", "control_type": "Window",
                           "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=8):
        sprut.find_element({"title": "ОК", "class_name": "Button", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()
    sleep(.3)
    sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys("{F5}")

    sprut.parent_switch({"title": "Поиск", "class_name": "Tvms_search_fm_builder", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()

    try:
        sprut.find_element({"title": "", "class_name": "", "control_type": "Button",
                            "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=10).click()
    except:
        pass

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('^N')

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('%%', sprut.keys.ENTER, protect_first=True)

    keyboard.send_keys("{RIGHT}")  # ? Uncomment
    keyboard.send_keys("^A")
    keyboard.send_keys("^{VK_INSERT}")
    sleep(.1)
    keyboard.send_keys("^{VK_INSERT}")

    sleep(.1)
    clipboard_data = clipboard_get()

    data_lines = clipboard_data.strip().split('\n')
    data = [line.split(',') for line in data_lines]

    # df2 = pd.DataFrame(data)

    # def apply_replacements(cell):
    #     return cell.replace('.', '').replace('"', '')
    #
    # data_frame = data_frame.applymap(apply_replacements)

    # df2 = df2.transpose()
    #
    # df2.columns = ['stores']

    # ? Close the window and return to tme root Sprut
    sprut.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()
    sprut.parent_back(1)

    sprut.find_element({"title": "Отмена", "class_name": "TvmsFooterButton", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.parent_switch({"title": "\"Отчеты\"", "class_name": "Treport_frm_main", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0})

    sprut.find_element({"title": "Приложение", "class_name": "", "control_type": "MenuBar",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(270, 20))

    sprut.parent_switch({"title": "\"Главное меню ПС СПРУТ\"", "class_name": "Tsprut_fm_Main", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0})
    return data[0]


def get_branches_to_execute(df1, branches_with_quote):
    def apply_replacements(cell):
        return cell.replace('.', '').replace('"', '')

    branches_without_quote = pd.DataFrame(branches_with_quote)
    branches_without_quote = branches_without_quote.applymap(apply_replacements)
    branches_without_quote.columns = ['stores']

    def apply_replacements1(cell):
        return cell.replace('.xlsx', '')

    df1['store_name'] = df1['store_name'].apply(apply_replacements1)
    logger.info(df1['store_name'])

    import numpy as np

    logger.info(f"{len(np.setdiff1d(np.asarray(branches_without_quote['stores']), np.asarray(df1['store_name'])))}")

    skipped_branches = np.setdiff1d(np.asarray(branches_without_quote['stores']), np.asarray(df1['store_name']))
    logger.info(skipped_branches)
    logger.info('-------------------------------------------------------------------------')

    branches_to_execute_ = []

    for branch in branches_with_quote:
        branch_ = str(branch).lower().replace('.', '')

        for branch1 in skipped_branches:
            branch1_ = str(branch1).lower()

            diff = Levenshtein.distance(branch_, branch1_)
            if diff <= 2:
                logger.info(f'TO EXECUTE: {diff} | {branch}, {branch1}')
                branches_to_execute_.append(branch)
                break

    logger.info(branches_to_execute_)

    return branches_to_execute_


if __name__ == '__main__':

    failed = False

    today = datetime.datetime.today()

    start_date = (today - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
    end_date = (today - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
    today = today.strftime('%d.%m.%Y')

    print(start_date, end_date)

    for i in range(5):
        try:
            try:
                sql_delete_table()
            except:
                pass

            sql_create_table()

            sprut = Sprut("MAGNUM")
            sprut.run()
            try:
                df2 = get_all_existing_branches_from_sprut(sprut)
            except Exception as e:
                logger.info(e)
                sleep(1000)

            try:
                df = get_all_data()

                branches_to_execute = get_branches_to_execute(df, df2)

                print(df)

            except:
                branches_to_execute = df2

            branches = ['Алматинский филиал №1 ТОО "Magnum Cash&Carry"', 'Товарищество с ограниченной ответственностью Magnum Cash&Carry(777)', 'Алматинский филиал №2 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №3  ТОО "Magnum Cash&Carry"', 'Карагандинский Филиал №1 ТОО "Magnum Cash&Carry"', 'Филиал №1 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №4 ТОО "Magnum Cash&Carry" в г. Алматы', 'Филиал ТОО "Magnum Cash&Carry" №5 в г. Алматы', 'Алматинский филиал №6 ТОО "Magnum Cash&Carry"', 'Филиал Тест ТОО "Magnum cash&carry"', 'Алматинский филиал №7 ТОО "Magnum Cash&Carry"', 'Филиал ТОО "Magnum cash&carry" в г. Шымкент', 'Алматинский филиал №8 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №10 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №9 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №11 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №12 ТОО "Magnum Cash&Carry"', 'Филиал №2 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал ТОО "Magnum cash&carry" в г. Талдыкорган',
                        'Алматинский филиал №14 ТОО "Magnum Cash&Carry"', 'Филиал №2 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №3 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №2 ТОО "Magnum Cash&Carry" в г.Талдыкорган', 'Алматинский филиал №16 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №15 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №17 ТОО "Magnum Cash&Carry"', 'Филиал №1 ТОО "Magnum Cash&Carry" в г.Каскелен', 'Алматинский филиал №20 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №18 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №19 ТОО "Magnum Cash&Carry"', 'Филиал №4 ТОО "Magnum Cash&Carry" в г.Шымкент', 'Карагандинский филиал №2 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №21 ТОО "Magnum Cash&Carry"', 'Филиал №1 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Алматинский филиал №22 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №23 ТОО "Magnum Cash&Carry"', 'Филиал №3 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Алматинский филиал №24 ТОО "Magnum Cash&Carry"',
                        'Филиал №4 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №5 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Алматинский филиал №25 ТОО "Magnum Cash&Carry"', 'Филиал №1 в г. Кызылорда ТОО "Magnum Cash&Carry"', 'Алматинский филиал №26 ТОО "Magnum Cash&Carry"', 'Филиал №6 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №7 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №8 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №9 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №10 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №1 ТОО "Magnum Cash&Carry" в г. Тараз', 'Алматинский филиал №32 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №28 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №29 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №30 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №31 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №33 ТОО Magnum Cash&Carry', 'Алматинский филиал №34 ТОО Magnum Cash&Carry', 'Алматинский филиал №35 ТОО Magnum Cash&Carry',
                        'Филиал №36 ТОО "Magnum Cash&Carry" в г Алматы',
                        'Филиал №37 ТОО "Magnum Cash&Carry" в г. Алматы', 'Филиал №38 ТОО Magnum Cash&Carry в г. Алматы', 'Филиал №5 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №11 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №12 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №13 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №13 ТОО "Magnum Cash&Carry" в г.Алматы', 'Филиал №39 ТОО "Magnum Cash&Carry" в г.Алматы', 'Филиал №15 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №40 ТОО "MAGNUM CASH&CARRY" в г.Алматы', 'Алматинский филиал №41 ТОО "Magnum Cash&Carry"', 'Филиал №42 ТОО "Magnum Cash&Carry" в г.Алматы', 'Алматинский филиал №43 ТОО "Magnum Cash&Carry"', 'Филиал №14 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №6 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №7 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал РЦ №1 ТОО "Magnum Cash&Carry" в г.Астана', 'Филиал РЦ №2 ТОО "Magnum Cash&Carry" в г.Шымкент', 'Филиал №16 ТОО "MAGNUM CASH&CARRY" в г.Астана',
                        'Филиал №17 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Карагандинский филиал №4 ТОО "Magnum Cash&Carry"', 'Карагандинский филиал №3 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №44 ТОО "Magnum Cash&Carry"', 'Филиал №8 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №2 ТОО "Magnum Cash&Carry" в г. Тараз', 'Карагандинский филиал №5 ТОО "Magnum Cash&Carry"', 'Филиал №45 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №1 ТОО "Magnum Cash&Carry" в г.Есик', 'Филиал №19 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №46 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №24 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Алматинский филиал №49 ТОО "Magnum Cash&Carry"', 'Филиал №21 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №9 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №48 ТОО «MAGNUM СASH&CARRY» в г.Алматы', 'Филиал №10 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №20 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №56 ТОО «MAGNUM CASH&CARRY» в г. Алматы',
                        'Филиал №28 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №50 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №53 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №1 ТОО «МAGNUM СASH&CARRY» в г. Туркестан', 'Филиал №22 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №7 ТОО «МAGNUM СASH&CARRY» в г.Караганда', 'Филиал №51 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №23 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №18 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №52 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №25 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №54 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №55 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №26 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №27 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №29 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №30 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №60 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №2 в г. Кызылорда ТОО "Magnum Cash&Carry"',
                        'Карагандинский филиал №6 ТОО "Magnum Cash&Carry"', 'Дискаунтер Реалист №11', 'Филиал №59 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №58 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №1 ТОО "MAGNUM CASH&CARRY"  в г. Усть-Каменогорск', 'Филиал №2 ТОО "MAGNUM CASH&CARRY"  в г. Усть-Каменогорск', 'Филиал №31 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'ДУЦП ТОО «Magnum Cash&Carry»', 'Филиал №33 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №35 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №32 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №41 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №34 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №36 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №37 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №2 ТОО "Magnum Cash&Carry" в г.Каскелен', 'Филиал №47 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №2 ТОО «МAGNUM СASH&CARRY» в г. Туркестан', 'Филиал №61 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №38 ТОО "MAGNUM CASH&CARRY" в г.Астана',
                        'Филиал №39 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №40 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №42 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №51 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №48 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №49 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №43 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №44 ТОО "MAGNUM CASH&CARRY" г.Астана', 'Филиал №53 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №45 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №57 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №46 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №47 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №50 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №52 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №11 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №56 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №54 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №55 ТОО "MAGNUM CASH&CARRY" в г.Астана',
                        'Филиал №62 ТОО «MAGNUM CASH&CARRY» в г. Алматы',
                        'Филиал №63 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №12 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №68 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №3 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №14 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №67 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Распределительный центр №3 в Алматинской области', 'Филиал №66 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №69 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №63 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №64 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №57 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №62 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №15 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Алматинский филиал №71 ТОО "Magnum Cash&Carry"', 'Филиал №20 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №17 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №73 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №72 ТОО «MAGNUM СASH&CARRY» в г. Алматы',
                        'Филиал №18 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №19 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №65 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №3 ТОО «МAGNUM СASH&CARRY» по Туркестанской области', 'Филиал №61 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №20 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №21 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №58 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №1 ТОО "Magnum Cash&Carry" в г.Конаев', 'Филиал №3 ТОО "MAGNUM CASH&CARRY"  в г. Усть-Каменогорск', 'Филиал №19 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №22 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №64 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №65 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №17 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал РЦ №4 ТОО "Magnum Cash&Carry" в г.Петропавловск', 'Филиал №2 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №11 ТОО "Magnum Cash&Carry" в г. Петропавловск',
                        'Филиал №4 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №5 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №15 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №7 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №8 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №6 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №13 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №18 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №10 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №12 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №9 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №16 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №14 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №3 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №59 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №13 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №75 ТОО "Magnum Сash&Сarry" в г. Алматы',
                        'Филиал №60 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №21 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №22 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №23 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №70 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №24 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №25 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №26 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №27 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №28 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №29 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №30 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №31 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №32 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №33 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №34 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №4 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №5 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №6 ТОО "Magnum Cash&Carry" в г. Тараз',
                        'Филиал №7 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №8 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №9 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №10 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №66 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №67 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №35 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №23 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №76 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №1 Маркет холл ТОО "Magnum Cash&Carry" в г. Алматы', 'Филиал №68 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №69 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №71 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №73 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №70 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №74 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №72 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №75 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №77 ТОО "Magnum Сash&Сarry" в г. Алматы',
                        'Филиал №78 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №76 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №77 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №79 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Алматинский филиал №80 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №81 ТОО "Magnum Cash&Carry"', 'Филиал №82 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №83 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №4 ТОО «МAGNUM СASH&CARRY» в г. Туркестан', 'Филиал №79 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №84 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №85 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №80 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №81 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №86 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №82 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №83 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №74 ТОО "Magnum Сash&Сarry" в г. Алматы'
                        ]

            print(branches_to_execute, len(branches_to_execute))
            logger.info(branches_to_execute)
            send_in_cache(sprut, today)
            sprut.quit()

            create_z_reports(branches_to_execute, start_date, end_date)

            logger.info('Finishing')
            print('Finishing')
            write_branches_in_their_big_excels(end_date)
            print('Finished')
            logger.info('Finished')

            smtp_send(r"""Добрый день!
Расхождения, выявленные в отчете 100912 отражены в сводной таблице. Готовые сводные таблицы размещены на сетевой папке M:\Stuff\_06_Бухгалтерия\1. ОК и ЗО\алмата\отчет по контролю касс 2022г\Жадыра Робот; M:\Stuff\_06_Бухгалтерия\1. ОК и ЗО\алмата\отчет по контролю касс 2022г\Ардак Робот""",
                      to=['Abdykarim.D@magnum.kz', 'Mukhtarova@magnum.kz', 'Sakpankulova@magnum.kz', 'Nusipova@magnum.kz', 'Baishukova@magnum.kz'],
                      subject=f'Сверка чеков ОФД-Спрут robot за {today}', username=smtp_author, url=smtp_host)

            exit()

        except Exception as error:
            if i == 4:
                failed = True
            logger.info(f'Error occured: {error}\nRetried times: {i + 1}')
            # sleep(2000)

    smtp_send(r"""Добрый день!
    Робот не отработал ни одну из 5 попыток""",
              to=['Abdykarim.D@magnum.kz', 'Mukhtarova@magnum.kz', 'Sakpankulova@magnum.kz', 'Nusipova@magnum.kz', 'Baishukova@magnum.kz'],
              subject=f'Сверка чеков ОФД-Спрут robot за {today}', username=smtp_author, url=smtp_host)
