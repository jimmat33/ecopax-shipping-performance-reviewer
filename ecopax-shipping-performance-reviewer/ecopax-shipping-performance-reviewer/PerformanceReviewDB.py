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
                team_names = performance_prop_list[0]

                for name_val in team_names:
                    if team_names.index(name_val) == 0:
                        team_names_str = name_val.capitalize()
                    else:
                        team_names_str = team_names_str + '<' +name_val.capitalize()

                performance_prop_list[0] = team_names_str

                team_add_sql_statement = ''' INSERT INTO TeamPerformanceTable(TeamNames, JobDate, JobType, TimeWorking, TransportRefrence) VALUES(?,?,?,?,?) '''

                cur.execute(team_add_sql_statement, performance_prop_list)
            else:
                check_sql_statement = ''' SELECT WorkerName, TransportRefrence FROM IndividualPerformanceTable WHERE WorkerName =? AND TransportRefrence =? '''
                cur.execute(check_sql_statement, [performance_prop_list[0], performance_prop_list[4]])
                rows = cur.fetchall()

                performance_prop_list[0] = performance_prop_list[0].capitalize()

                if len(rows) == 0:
                    individual_add_sql_satatement = ''' INSERT INTO IndividualPerformanceTable(WorkerName, JobDate, JobType, JobTimeWorking, TransportRefrence, WorkerJob, IndividualTime, NumJobMembers) VALUES (?,?,?,?,?,?,?,?) '''
                    cur.execute(individual_add_sql_satatement, performance_prop_list)


        except Exception:
            pass

        db_connection.commit()


    db_connection.close()







