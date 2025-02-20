import datetime
import os
import shutil
import time
import traceback
from contextlib import suppress
from pathlib import Path

import Levenshtein
import pandas as pd
from time import sleep

import win32com
import win32com.client as win32
import psycopg2 as psycopg2
from pywinauto import keyboard

from config import logger, download_path, robot_name, db_host, db_port, db_name, db_user, db_pass, tg_token, chat_id, smtp_host, smtp_author, jadyra_path, ardak_path, mapping_path, global_password, global_username, saving_path, ip_address, main_executor, today_, end_date_
from core import Sprut
from tools.clipboard import clipboard_get
from tools.net_use import net_use
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

    ended_time = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f") if status != 'processing' else datetime.datetime.now()

    values = (
        started_time.strftime("%d.%m.%Y %H:%M:%S.%f"),
        ended_time,
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
        print(f'GOVNO {e}')
        pass

    try:
        cursor.execute(query, values)

    except Exception as e:
        conn.rollback()
        print(f"Error: {e}")

    conn.commit()

    cursor.close()
    conn.close()


def update_data_in_db(started_time, store_name, short_name, status, responsible, found_difference, count, error_reason, error_saved_path, execution_time):
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)

    query_delete = f"""
        delete from ROBOT.{robot_name.replace("-", "_")} where store_name = '{store_name}'
    """

    query = f"""        
        INSERT INTO ROBOT.{robot_name.replace("-", "_")}
            (started_time, ended_time, store_name, short_name, status, responsible, found_difference, count, error_reason, error_saved_path, execution_time)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE
            status = %s,
            started_time = %s,
            ended_time = %s,
            short_name = %s,
            responsible = %s,
            found_difference = %s,
            count = %s,
            error_reason = %s,
            error_saved_path = %s,
            execution_time = %s;
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
        print(f'GOVNO {e}')
        pass

    try:
        cursor.execute(query, values)

    except Exception as e:
        conn.rollback()
        print(f"Error: {e}")

    conn.commit()

    cursor.close()
    conn.close()


def check_if_store_in_db(store_name):
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            where store_name = '{store_name}' and status = 'success'
            '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())

    cur.close()
    conn.close()
    return True if len(df1) >= 1 else False


def get_data_to_execute():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            where (status != 'success' and status != 'processing')
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

    print('Started writing branches in their big excels')

    with suppress(Exception):
        os.system('taskkill /im excel.exe')

    # ? Create new page
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    # excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    end_date_ = datetime.datetime.strptime(end_date_, '%d.%m.%Y')
    year = str(end_date_.year)
    month = end_date_.month

    print(year, month)

    def open_excel(path):

        found = False

        excel1 = win32.gencache.EnsureDispatch('Excel.Application')
        excel1.Visible = False
        excel1.DisplayAlerts = False
        wb = excel1.Workbooks.Open(path)
        last_sheet_index = wb.Worksheets.Count

        ws = None

        for sheet in wb.Worksheets:
            if months[month - 1].lower() in str(sheet.Name).lower() and (year in str(sheet.Name) or year[2:] in str(sheet.Name)):
                print('sheet name:', sheet.Name)
                ws = wb.Worksheets(sheet.Name)
                found = True
                print('Нашёл:', sheet.Name)

        if not found:
            print('Started creating new sheet')
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

                excel_ = win32.gencache.EnsureDispatch('Excel.Application')
                excel_.Visible = False
                excel_.DisplayAlerts = False

                wb0 = excel_.Workbooks.Open(os.path.join(saving_path, single_branch))
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
    for branch in os.listdir(saving_path):

        if branch == 'Secondary machine finished.txt':
            continue

        df = pd.read_excel(os.path.join(saving_path, branch))

        try:
            df.columns = ['№ Кассы', 'Регистр. № кассы', '№', 'Итог продаж', 'Возвраты: (нал,безнал, бонус)', 'Возвраты Бонусы', 'Итого за минусом возвратов:', 'безнал', 'Итого наличных', 'итого наличных', 'Сертификаты подаренные', 'Сертификаты, реализованные частным лицам', 'Сертификаты, реализованные юр/ лицам', 'Сертификаты, созданные при возврате товара', 'Чеки по акции "Счастливый чек" (Бесплатные чеки)', 'Нехватка разменных монет', 'Оплата Бонусами', 'безнал', 'Итого продаж', 'Нал Разница', 'Безнал Разница', 'Разница', 'ООФД - Z - отчет - СПРУТ']
        except:
            df.columns = ['№ Кассы', 'Регистр. № кассы', '№', 'Итог продаж', 'Возвраты: (нал,безнал, бонус)', 'Возвраты Бонусы', 'Итого за минусом возвратов:', 'безнал', 'Итого наличных', 'итого наличных', 'Сертификаты подаренные', 'Сертификаты, реализованные частным лицам', 'Сертификаты, реализованные юр/ лицам', 'Сертификаты, созданные при возврате товара', 'Чеки по акции "Счастливый чек" (Бесплатные чеки)', 'Нехватка разменных монет', 'Оплата Бонусами', 'безнал', 'Итого продаж', 'Нал Разница', 'Безнал Разница', 'Разница', 'ООФД - Z - отчет - СПРУТ', 'Сумма скидки итого', 'Итого продаж за минусом скидки и возвратов, предоставлены в ОФД']

        title = df['Регистр. № кассы'].iloc[5]
        try:
            started_time = datetime.datetime.now()
            start_time = time.time()
            if 'nusipova' in str(df1[df1['Название филиала в Спруте'] == title]['Сотрудник'].iloc[0]).lower():
                short_name = df1[df1['Название филиала в Спруте'] == title]['Короткое название филиала'].iloc[0]
                try:
                    nusipova_ws, found, count = check_one_branch(nusipova_ws, short_name, branch)
                    end_time = time.time()
                    # insert_data_in_db(started_time, branch, short_name, 'success', 'Nusipova', found, count, '', '', str(end_time - start_time))
                except Exception as ex:
                    end_time = time.time()
                    tg_send(f'FAILED: {short_name} | ({branch})', bot_token=tg_token, chat_id=chat_id)
                    # insert_data_in_db(started_time, branch, short_name, 'failed', 'Nusipova', '', 0, str(ex), '', str(end_time - start_time))

            else:
                short_name = df1[df1['Название филиала в Спруте'] == title]['Короткое название филиала'].iloc[0]

                try:
                    baishukova_ws, found, count = check_one_branch(baishukova_ws, title, branch)
                    end_time = time.time()
                    # insert_data_in_db(started_time, branch, short_name, 'success', 'Baishukova', found, count, '', '', str(end_time - start_time))
                except Exception as ex:
                    end_time = time.time()
                    tg_send(f'FAILED: {short_name} | ({branch})', bot_token=tg_token, chat_id=chat_id)
                    # insert_data_in_db(started_time, branch, short_name, 'failed', 'Baishukova', '', 0, str(ex), '', str(end_time - start_time))

        except Exception as error:
            print('error:', error)
    print('Finishing1')
    print('Заканчиваем1')
    # empty_row = baishukova_ws.Cells.SpecialCells(win32.constants.xlCellTypeLastCell).Row
    # baishukova_ws.Cells(empty_row, 1).EntireRow.Interior.ColorIndex = 40

    # empty_row = nusipova_ws.Cells.SpecialCells(win32.constants.xlCellTypeLastCell).Row
    # nusipova_ws.Cells(empty_row, 1).EntireRow.Interior.ColorIndex = 40

    baishukova_wb.Save()
    baishukova_wb.Close()

    nusipova_wb.Save()
    nusipova_wb.Close()
    excel.Application.Quit()

    print('Заканчиваем2')
    print('Finishing2')


def send_in_cache(today):

    sprut = Sprut("MAGNUM")
    sprut.run()
    sprut.open("Контроль передачи данных", switch=False)

    print('Switching')
    sprut.parent_switch({"title_re": ".Контроль передачи данных.", "class_name": "Tcontrolcache_fm_main",
                         "control_type": "Window", "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()
    print('Switched')
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

    print('Clicked list')
    keyboard.send_keys("{DOWN}" * 4)
    keyboard.send_keys("{ENTER}")

    print('Clicked item')
    sprut.find_element({"title": "", "class_name": "TvmsBitBtn", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "И", "class_name": "TvmsBitBtn", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()
    print('KEKUS')
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

    print('clicked')
    while True:
        try:
            sprut.find_element({"title": "Ввод", "class_name": "TvmsBitBtn", "control_type": "Button",
                                "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1).click()
            break
        except:
            pass

    print('clicked1')

    sprut.parent_back(1)

    sprut.find_element({"title": " ", "class_name": "TPanel", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=600).click(coords=(20, 17))

    sprut.find_element({"title": "Журналы", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 2}).click()

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()

    sprut.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    # sprut.parent_back(1)

    sprut.quit()


def create_z_reports(branches, start_date, end_date):

    for ind_, branch in enumerate(branches):

        if str(branch) == 'nan':
            continue

        if check_if_store_in_db(branch):
            continue

        # found = False
        # branch_ = branch.replace('.', '').replace('"', '').replace('«', '').replace('»', '')
        # for file_ in os.listdir(saving_path):
        #     if branch_ in file_:
        #         found = True
        #         break
        # if found:
        #     continue

        insert_data_in_db(datetime.datetime.now(), branch, '', 'processing', '', '', 0, '', '', '')

        for _ in range(5):
            try:
                print(f'Started branch: {ind_}, {branch}')

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
                                    "visible_only": True, "enabled_only": True, "found_index": 0}).click(double=True, set_focus=True)

                sleep(1)

                while True:
                    try:
                        sprut.find_element({"title": "Ввод", "class_name": "TvmsFooterButton", "control_type": "Button",
                                            "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=0.5).click()
                    except:
                        break
                # sleep(0.1)
                #
                # with suppress(Exception):
                #     sprut.find_element({"title": "Ввод", "class_name": "TvmsFooterButton", "control_type": "Button",
                #                         "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=0.1).click(set_focus=True)

                wait_loading(branch)

                # sprut.parent_back(1).set_focus()

                sprut.quit()

                print('Finished branch')

                insert_data_in_db(datetime.datetime.now(), branch, '', 'success', '', '', 0, '', '', '')

                break

            except Exception as exc:
                print(f'Error occured at {branch}: {traceback.print_exc()}')
                sleep(1)
        print('-----------------------------------------------------------------------')
    print('Finished CREATING Z REPORTS')
    print('-----------------------------------------------------------------------')


def wait_loading(branch):

    print('Started loading')
    branch = branch.replace('.', '').replace('"', '').replace('«', '').replace('»', '')
    found = False

    count = 0

    while True:

        for file in os.listdir(download_path):
            sleep(.1)
            creation_time = os.path.getctime(os.path.join(download_path, file))
            current_time = datetime.datetime.now().timestamp()
            time_difference = current_time - creation_time
            minutes_since_creation = time_difference / 60

            if int(minutes_since_creation) <= 2 and file[0] != '$' and '.' in file and 'xl' in file and '100912' in file:
                print(file)
                type = '.' + file.split('.')[1]
                shutil.move(os.path.join(download_path, file), os.path.join(saving_path, branch + type))
                found = True
                break
        if found:
            break

        sleep(5)
        count += 1

        if count >= 120:  # 10 minutes (600 seconds)
            print('----------------------------------- Ne dozhdalsya govna -----------------------------------')
            print('----------------------------------- Ne dozhdalsya govna -----------------------------------')
            print('----------------------------------- Ne dozhdalsya govna -----------------------------------')
            break

    if not found:
        raise Exception("Error: Could not find excel")
    print('Finished loading')
    print('Finished loading')


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
    print(df1['store_name'])

    import numpy as np

    print(f"{len(np.setdiff1d(np.asarray(branches_without_quote['stores']), np.asarray(df1['store_name'])))}")

    skipped_branches = np.setdiff1d(np.asarray(branches_without_quote['stores']), np.asarray(df1['store_name']))
    print(skipped_branches)
    print('-------------------------------------------------------------------------')

    branches_to_execute_ = []

    for branch in branches_with_quote:
        branch_ = str(branch).lower().replace('.', '')

        for branch1 in skipped_branches:
            branch1_ = str(branch1).lower()

            diff = Levenshtein.distance(branch_, branch1_)
            if diff <= 2:
                print(f'TO EXECUTE: {diff} | {branch}, {branch1}')
                branches_to_execute_.append(branch)
                break

    print(branches_to_execute_)

    return branches_to_execute_


def archive_files(prev_date):

    try:
        os.makedirs(os.path.join(saving_path.parent, f'reports_ofd_zip'))
    except:
        pass
    destination_folder = os.path.join(saving_path.parent, f'reports_ofd_zip')

    zip_file_name = f'Выгрузка сверки чеков за {prev_date}'
    zip_file_path = os.path.join(destination_folder, zip_file_name)

    shutil.make_archive(zip_file_path, 'zip', saving_path)

    return zip_file_path


def wait_until_main_machine_finished():

    found = False

    while True:
        if not os.path.isfile(os.path.join(saving_path, 'Secondary machine finished.txt')):
            break
        # for file_ in os.listdir(saving_path):
        #     if file_ == 'Secondary machine finished.txt':
        #         found = True
        #         break
        #
        # if not found:
        #     break

        sleep(10)


def wait_until_secondary_machine_finished():

    # found = False

    while True:
        if os.path.isfile(os.path.join(saving_path, 'Secondary machine finished.txt')):
            break

        # for file_ in os.listdir(saving_path):
        #     if file_ == 'Secondary machine finished.txt':
        #         found = True
        #         break
        #
        # if not found:
        #     break

        sleep(10)


if __name__ == '__main__':

    for day in range(1):

        failed = False
        today = datetime.datetime.today()

        if end_date_ == '':
            start_date = (today - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
            end_date = (today - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
            today = today.strftime('%d.%m.%Y')

        else:
            start_date = end_date_
            end_date = end_date_
            today = today_

        # Пример
        # start_date = '24.03.2024'  # ? Comment in prod
        # end_date = '24.03.2024'  # * Comment in prod
        # today = '25.03.2024'  # * Comment in prod

        # save_date = (datetime.date(int(end_date.split('.')[2]), int(end_date.split('.')[1]), int(end_date.split('.')[0])) - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
        save_date = end_date
   
        print(start_date, end_date, '|', save_date)

        net_use(Path(ardak_path).parent.parent, global_username, global_password)
        net_use(ardak_path, global_username, global_password)
        net_use(jadyra_path, global_username, global_password)

        # main_executor = '192.168.0.110'
        # * -----
        # main_executor = '10.70.2.9'  # '172.20.1.24'

        logger.info(end_date)
        logger.warning(f'Робот запустился за дату {today} на машине {ip_address}, дата сохранения отчётов {save_date}')

        with suppress(Exception):

            if ip_address == main_executor:

                with suppress(Exception):
                    sql_delete_table()

                sql_create_table()

        for i in range(5):

            try:

                df = pd.read_excel(mapping_path)

                branches_to_execute = list(df[df['Сотрудник'] == 'Baishukova@magnum.kz']['Название филиала в Спруте'])

                if ip_address == main_executor:
                    branches_to_execute = list(df[df['Сотрудник'] == 'Nusipova@magnum.kz']['Название филиала в Спруте'])

                print(len(branches_to_execute), branches_to_execute)

                send_in_cache(end_date)

                print('Started Z Reports')

                create_z_reports(branches_to_execute, end_date, end_date)

                print('Finishing0')

                if ip_address == main_executor:

                    wait_until_secondary_machine_finished()

                    print('Finishing')
                    for tries in range(5):
                        logger.warning(f'Запись в файлы сбора. Попытка {tries + 1} / 5')
                        try:
                            write_branches_in_their_big_excels(save_date)
                            break
                        except Exception as error:
                            traceback.print_exc()
                            smtp_send(r"""Робот сломался при записи в главные эксели""",
                                      to=['Abdykarim.D@magnum.kz'],
                                      subject=f'Сбор расхождений по чекам за {end_date}', username=smtp_author, url=smtp_host)
                            logger.warning(f'ERROR OCCURED: {error}')
                            with suppress(Exception):
                                os.system('taskkill /im excel.exe')
                            sleep(10)
                            pass

                    with suppress(Exception):
                        Path.unlink(Path(os.path.join(saving_path, 'Secondary machine finished.txt')))

                    archive_files(end_date)

                    smtp_send(r"""Добрый день!
                                Расхождения, выявленные в отчете 100912 отражены в сводной таблице. Готовые сводные таблицы размещены на сетевой папке M:\Stuff\_06_Бухгалтерия\1. ОК и ЗО\алмата\отчет по контролю касс 2022г\Жадыра Робот; M:\Stuff\_06_Бухгалтерия\1. ОК и ЗО\алмата\отчет по контролю касс 2022г\Ардак Робот""",
                              to=['Abdykarim.D@magnum.kz', 'Sakpankulova@magnum.kz', 'ABITAKYN@magnum.kz', 'Baishukova@magnum.kz'],
                              subject=f'Сбор расхождений по чекам за {today}', username=smtp_author, url=smtp_host)
                    logger.info('Процесс закончился успешно')
                    failed = False

                    logger.info(f'Закончили на дату {end_date}')
                    logger.warning(f'Робот успешно завершился на дату {end_date}')

                else:

                    with open(os.path.join(saving_path, 'Secondary machine finished.txt'), 'w') as file:
                        file.write('kek')

                break

            except Exception as error:
                traceback.print_exc()
                with suppress(Exception):
                    os.system('taskkill /im excel.exe')
                # if i == 4:
                #     failed = True
                # print(f'Error occured: {error}\nRetried times: {i + 1}')
                # logger.warning(f'Error occured: {error}\nRetried times: {i + 1}')
                # sleep(2000)
        if failed and ip_address == main_executor:
            # logger.info(f'Робот сломался')
            smtp_send(r"""Добрый день!
            Робот не отработал ни одну из 5 попыток""",
                      to=['Abdykarim.D@magnum.kz', 'Sakpankulova@magnum.kz', 'ABITAKYN@magnum.kz', 'Baishukova@magnum.kz'],
                      subject=f'Сбор расхождений по чекам за {end_date}', username=smtp_author, url=smtp_host)

            logger.warning(f'Робот сломался на дату {end_date} на машине {ip_address}')

        elif failed and ip_address != main_executor:
            logger.warning(f'Робот сломался на дату {end_date} на машине {ip_address}')

    with suppress(Exception):
        sql_delete_table()


