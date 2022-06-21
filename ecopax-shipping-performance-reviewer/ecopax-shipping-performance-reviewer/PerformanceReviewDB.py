import os
import sqlite3 as sl
from datetime import datetime

def db_connect():
    db_file_name = os.path.abspath('performance-review.db')
    db_conn = None

    try: 
        db_conn = sl.connect(db_file_name)
    except sl.Error as e:
        print('\n')
        print(e)
        print('\n')

    return db_conn

def db_add_performance_entry(performance_prop_list, add_type):
    db_connection = db_connect()

    with db_connection:
        try:
            cur = db_connection.cursor()

            if add_type == 'team':

                team_names_str = ''
                team_names = performance_prop_list[0].strip()

                for name_val in team_names:
                    if team_names.index(name_val) == 0:
                        team_names_str = name_val.strip().capitalize()
                    else:
                        team_names_str = team_names_str + '<' + name_val.strip().capitalize()

                performance_prop_list[0] = team_names_str.strip()

                team_add_sql_statement = ''' INSERT INTO TeamPerformanceTable(TeamNames, JobDate, JobType, TimeWorking, TransportRefrence, Filepath) VALUES(?,?,?,?,?,?) '''

                cur.execute(team_add_sql_statement, performance_prop_list)
            else:

                performance_prop_list[0] = performance_prop_list[0].capitalize().strip()
                check_sql_statement = ''' SELECT WorkerName, TransportRefrence FROM IndividualPerformanceTable WHERE WorkerName =? AND TransportRefrence =? '''
                cur.execute(check_sql_statement, [performance_prop_list[0], performance_prop_list[4]])
                rows = cur.fetchall()

                performance_prop_list[0] = performance_prop_list[0].capitalize().strip()

                if len(rows) == 0:
                    individual_add_sql_satatement = ''' INSERT INTO IndividualPerformanceTable(WorkerName, JobDate, JobType, JobTimeWorking, TransportRefrence, WorkerJob, IndividualTime, NumJobMembers, FilePath) VALUES (?,?,?,?,?,?,?,?,?) '''
                    cur.execute(individual_add_sql_satatement, performance_prop_list)


        except Exception:
            pass

        db_connection.commit()


    db_connection.close()


def db_get_data_from_column(column_name, table_name):
    db_connection = db_connect()
    ret_list = []

    with db_connection:
        cur = db_connection.cursor()

        if table_name == 'team':
            sql_statement = f''' SELECT {column_name} FROM TeamPerformanceTable '''
        else:
            sql_statement = f''' SELECT {column_name} FROM IndividualPerformanceTable '''


        cur.execute(sql_statement)

        cur_list = cur.fetchall()

        db_connection.commit()

    db_connection.close()

    for item in cur_list:
            ret_list.append(item[0])

    if column_name == 'TeamNames':
        new_ret_list = []
        for nameset in ret_list:
            if len(nameset.split('<')) != 1:
                new_ret_str = nameset.replace('<', ', ')
                new_ret_list.append(new_ret_str)

        return new_ret_list
    else:
        return ret_list


def db_add_excel_file(file_props):
    db_connection = db_connect()

    with db_connection:
        cur = db_connection.cursor()

        add_file_sql_statement = ''' INSERT INTO ExcelFileTable(Filepath, FileName, DateRange) VALUES(?,?,?) '''

        try:
            cur.execute(add_file_sql_statement, file_props)
        except Exception:
            pass

        db_connection.commit()

    db_connection.close()


def db_get_excel_file(filename):
    db_connection = db_connect()
    file_list = []

    with db_connection:
        cur = db_connection.cursor()

        get_file_sql_statement = ''' SELECT FilePath, FileName, DateRange FROM ExcelFileTable WHERE FilePath =? '''

        try:
            cur.execute(get_file_sql_statement, [filename])

        except Exception:
            pass
        
        
        file_list = cur.fetchall()
        db_connection.commit()

    db_connection.close()

    return file_list


def db_get_all_excel_filepaths():
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
    db_connection = db_connect()

    with db_connection:
        cur = db_connection.cursor()

        remove_file_sql_statement_1 = ''' DELETE FROM IndividualPerformanceTable WHERE FilePath =? '''
        remove_file_sql_statement_2 = ''' DELETE FROM TeamPerformanceTable WHERE FilePath =? '''
        remove_file_sql_statement_3 = ''' DELETE FROM ExcelFileTable WHERE FilePath =? '''

        try:
            cur.execute(remove_file_sql_statement_1, [filepath])
            cur.execute(remove_file_sql_statement_2, [filepath])
            cur.execute(remove_file_sql_statement_3, [filepath])
        except Exception:
            pass

        db_connection.commit()

    db_connection.close()
        






