import pandas as pd
from pathlib import Path
from PerformanceReviewDB import *
import re
import datetime

class TWExportExcelFile(object):
    def __init__(self, old_filepath):
        self.old_filepath = old_filepath
        self._filepath = self.fix_file_ext()

        self.process_file()
        

    def fix_file_ext(self):
        base_file = self.old_filepath
        name, ext = base_file.split('.')
        new_file = '{}.{}'.format(name, 'csv')

        try:
            with open(base_file , encoding="utf_16") as f1:
                with open(new_file, 'w') as f2:
                    f2.write(f1.read())
        except:
            pass

        return new_file


    def get_team_names(self,team_names_lst):
        if len(team_names_lst) == 0:
            return []
        
        fst_lst = team_names_lst.split('&')
        fst_lst[-1] = fst_lst[-1].split('/')[0].strip()

        for name in fst_lst:
            name = re.sub('[\"\']', '', name)
        

        return fst_lst

    
    def process_file(self):
        filename = self._filepath.split('/')[-1].split('.')[0]

        date_range_list = self.process_data()
        date_range = date_range_list[0] + ' - ' + date_range_list[1]

        db_add_excel_file([self._filepath, filename, date_range])



    def process_data(self):
        date_range_list = []
        fin_date_range_list = ['','']
        file_path = Path(self._filepath)
        file_extension = file_path.suffix.lower()[1:]
        total_job_time = 0

        if file_extension == 'xlsx':
            df = pd.read_excel(self._filepath, engine='openpyxl')
        elif file_extension == 'xls':
            df = pd.read_excel(self._filepath)
        elif file_extension == 'csv':
            df = pd.read_csv(self._filepath)
        else:
            raise Exception("File not supported")

        column_pd = list(df.columns)
        column_pd = column_pd[0]
        column_list = column_pd.split('\t')

        team_names_column = column_list.index('SL - Remarks')
        booking_date_column = column_list.index('SL - Date and time of booking')
        job_type_column = column_list.index('SL - Resource name')
        transport_refrence_column = column_list.index('SL - Transport reference')
        job_start_time_column = column_list.index('"SL - Custom timestamp 2 date and time "')
        job_end_time_column = column_list.index('"SL - Custom timestamp 3 date and time "')
        
        df = df.reset_index()
        
        for row_set in df.itertuples():
            row_str = row_set[2]
            row = row_str.split('\t')

            team_names = self.get_team_names(row[team_names_column])

            booking_date = row[booking_date_column]
            if len(booking_date) > 10:
                booking_date = booking_date[:-9].strip()
        
            job_type = row[job_type_column]
            job_type = re.sub('[\"\']', '', job_type)
            job_type = job_type.strip()

            transport_refrence = row[transport_refrence_column]

            job_start_time = row[job_start_time_column][-8:-3].strip()
            job_end_time = row[job_end_time_column][-8:-3].strip()

            if len(job_start_time[:-3]) == 1:
                job_start_time = '0' + job_start_time

            if len(job_end_time[:-3]) == 1:
                job_end_time = '0' + job_end_time

            try:
                job_start_time = datetime.datetime.strptime(job_start_time, '%H:%M')
                job_end_time = datetime.datetime.strptime(job_end_time, '%H:%M')

                total_job_time = job_end_time - job_start_time
                total_job_time = int(total_job_time.total_seconds() / 60)
                converted_total_job_time = str(total_job_time)
            except Exception:
                converted_total_job_time = 'error'


            if isinstance(team_names, list) and len(team_names) != 0:
                db_add_performance_entry([team_names, booking_date, job_type, converted_total_job_time, transport_refrence, self._filepath], 'team')

                for person_name in team_names:
                    if team_names.index(person_name) == 0:
                        worker_job = 'Driver'
                    else:
                        worker_job = 'Unloader'

                    time_divided = int(total_job_time / len(team_names))

                    try:
                        converted_individual_job_time = str(time_divided)
                    except Exception:
                        converted_individual_job_time = 'error'

                    db_add_performance_entry([person_name.strip(), booking_date, job_type,converted_total_job_time, transport_refrence,worker_job, converted_individual_job_time, len(team_names), self._filepath], 'individual')
                    date_range_list.append(booking_date)


        i = 0
        while i < len(date_range_list):
            try:
                date_range_list[i] = datetime.datetime.strptime(date_range_list[i],'%m/%d/%Y')
            except Exception:
                pass
            i += 1

        earliest_date = min(date_range_list)
        latest_date = max(date_range_list)
        try:
            fin_date_range_list[0] = datetime.datetime.strftime(earliest_date, '%m/%d/%Y')
            fin_date_range_list[1] = datetime.datetime.strftime(latest_date, '%m/%d/%Y')
        except Exception:
            pass
        return fin_date_range_list
        








