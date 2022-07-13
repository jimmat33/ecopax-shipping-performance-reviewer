'''
docstr
'''
import os
import sqlite3 as sl
import re
from datetime import date
# pylint: disable=W0703


def db_connect():
    '''
    docstr
    '''
    db_file_name = os.path.abspath('performance-review.db')
    db_conn = None

    try:
        db_conn = sl.connect(db_file_name)
    except sl.Error as conn_err:
        print('\n' + conn_err + '\n')

    return db_conn


def db_add_performance_entry(performance_prop_list, add_type):
    '''
    docstr
    '''
    db_connection = db_connect()

    with db_connection:
        try:
            cur = db_connection.cursor()

            if add_type == 'team':

                team_names_str = ''
                team_names = performance_prop_list[0]

                for name_val in team_names:
                    if name_val != '':
                        if team_names.index(name_val) == 0:
                            name_val = re.sub('[\"\']', '', name_val)
                            team_names_str = name_val.capitalize()
                        else:
                            name_val = re.sub('[\"\']', '', name_val)
                            team_names_str = team_names_str + '<' + name_val.capitalize()

                performance_prop_list[0] = team_names_str

                team_add_sql_statement = ''' INSERT INTO TeamPerformanceTable
                                            (TeamNames, JobDate, JobType, TimeWorking,
                                            TransportRefrence, Filepath)
                                            VALUES(?,?,?,?,?,?) '''

                cur.execute(team_add_sql_statement, performance_prop_list)
            else:

                performance_prop_list[0] = performance_prop_list[0].capitalize().strip()
                if performance_prop_list[0] != '':
                    check_sql_statement = ''' SELECT WorkerName, TransportRefrence
                                              FROM IndividualPerformanceTable
                                              WHERE WorkerName =? AND TransportRefrence =? '''
                    cur.execute(check_sql_statement, [performance_prop_list[0],
                                                      performance_prop_list[4]])
                    rows = cur.fetchall()

                    performance_prop_list[0] = re.sub('[\"\']', '', performance_prop_list[0])
                    performance_prop_list[0] = performance_prop_list[0].capitalize().strip()

                    if len(rows) == 0:
                        individual_add_sql_satatement = ''' INSERT INTO IndividualPerformanceTable
                                                            (WorkerName, JobDate, JobType,
                                                            JobTimeWorking, TransportRefrence,
                                                            WorkerJob, NumJobMembers, FilePath)
                                                            VALUES (?,?,?,?,?,?,?,?) '''
                        cur.execute(individual_add_sql_satatement, performance_prop_list)

        except Exception:
            pass

        db_connection.commit()

    db_connection.close()


def db_get_individual_data():
    '''
    docstr
    '''
    db_connection = db_connect()
    ret_list = []
    month_lst = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

    with db_connection:
        cur = db_connection.cursor()

        get_sql_statement = ''' SELECT WorkerName, JobDate, JobType, JobTimeWorking,
                                WorkerJob
                                FROM IndividualPerformanceTable '''

        cur.execute(get_sql_statement)

        ret_list = cur.fetchall()

        db_connection.commit()
    db_connection.close()

    i = 0
    while i < len(ret_list):
        ret_list[i] = list(ret_list[i])

        date_str = ret_list[i][1]
        date_lst = date_str.split('/')
        if len(date_lst[2]) > 4:
            date_lst[2] = date_lst[2][0:4]
        ret_list[i][1] = date(int(date_lst[2]), int(date_lst[0]), int(date_lst[1]))

        month_str = month_lst[int(ret_list[i][1].month) - 1]
        ret_list[i].append(410)
        ret_list[i].append(month_str)
        i += 1

    return ret_list


def db_get_team_data():
    '''
    docstr
    '''
    db_connection = db_connect()
    ret_list = []
    month_lst = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

    with db_connection:
        cur = db_connection.cursor()

        get_sql_statement = ''' SELECT TeamNames, JobDate, JobType, TimeWorking
                                FROM TeamPerformanceTable '''

        cur.execute(get_sql_statement)

        ret_list = cur.fetchall()

        db_connection.commit()
    db_connection.close()

    i = 0
    while i < len(ret_list):
        ret_list[i] = list(ret_list[i])

        date_str = ret_list[i][1]
        date_lst = date_str.split('/')

        if len(date_lst[2]) > 4:
            date_lst[2] = date_lst[2][0:4]

        ret_list[i][1] = date(int(date_lst[2]), int(date_lst[0]), int(date_lst[1]))

        month_str = month_lst[int(ret_list[i][1].month) - 1]
        new_str = ret_list[i][0].replace('<', ', ')
        ret_list[i][0] = new_str
        ret_list[i].append(month_str)
        i += 1

    return ret_list


def db_add_excel_file(file_props):
    '''
    docstr
    '''
    db_connection = db_connect()

    with db_connection:
        cur = db_connection.cursor()

        add_file_sql_statement = ''' INSERT INTO ExcelFileTable(Filepath, FileName, DateRange)
                                     VALUES(?,?,?) '''

        try:
            cur.execute(add_file_sql_statement, file_props)
        except Exception:
            pass

        db_connection.commit()

    db_connection.close()


def db_get_excel_file(filename):
    '''
    docstr
    '''
    db_connection = db_connect()
    file_list = []

    with db_connection:
        cur = db_connection.cursor()

        get_file_sql_statement = ''' SELECT FilePath, FileName, DateRange FROM ExcelFileTable
                                     WHERE FilePath =? '''

        try:
            cur.execute(get_file_sql_statement, [filename])

        except Exception:
            pass

        file_list = cur.fetchall()
        db_connection.commit()

    db_connection.close()

    return file_list


def db_get_all_excel_filepaths():
    '''
    docstr
    '''
    db_connection = db_connect()
    file_list = []

    with db_connection:
        cur = db_connection.cursor()

        get_file_sql_statement = ''' SELECT FilePath FROM ExcelFileTable'''

        try:
            cur.execute(get_file_sql_statement)

        except Exception:
            pass

        file_list = cur.fetchall()
        db_connection.commit()

    db_connection.close()

    return file_list


def db_remove_excel_file(filepath):
    '''
    docstr
    '''
    db_connection = db_connect()

    with db_connection:
        cur = db_connection.cursor()

        remove_file_sql_statement_1 = ''' DELETE FROM IndividualPerformanceTable
                                          WHERE FilePath =? '''
        remove_file_sql_statement_2 = ''' DELETE FROM TeamPerformanceTable
                                          WHERE FilePath =? '''
        remove_file_sql_statement_3 = ''' DELETE FROM ExcelFileTable
                                          WHERE FilePath =? '''

        try:
            cur.execute(remove_file_sql_statement_1, [filepath])
            cur.execute(remove_file_sql_statement_2, [filepath])
            cur.execute(remove_file_sql_statement_3, [filepath])
        except Exception:
            pass

        db_connection.commit()

    db_connection.close()


def db_clear_database():
    '''
    docstr
    '''
    db_connection = db_connect()

    with db_connection:
        cur = db_connection.cursor()

        remove_file_sql_statement_1 = ''' DELETE FROM IndividualPerformanceTable '''
        remove_file_sql_statement_2 = ''' DELETE FROM TeamPerformanceTable '''
        remove_file_sql_statement_3 = ''' DELETE FROM ExcelFileTable '''

        try:
            cur.execute(remove_file_sql_statement_1)
            cur.execute(remove_file_sql_statement_2)
            cur.execute(remove_file_sql_statement_3)
        except Exception:
            pass

        db_connection.commit()

    db_connection.close()
