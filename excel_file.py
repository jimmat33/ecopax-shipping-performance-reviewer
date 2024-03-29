'''
docstr
'''
import datetime
import re
import os
import traceback
import csv
from pathlib import Path
import pandas as pd
from performance_review_db import db_add_excel_file, db_add_performance_entry
# pylint: disable=W0703


class TWExportExcelFile():
    '''
    docstr
    '''
    def __init__(self, old_filepath):
        '''
        docstr
        '''
        self.old_filepath = old_filepath
        self._filepath = self.fix_file_ext()
        self.date_range_list = []
        self.process_data()
        self.process_file()
        os.remove(self._filepath)

    def fix_file_ext(self):
        '''
        doctr
        '''
        base_file = self.old_filepath
        name = base_file.split('.')
        new_file = f'{name[0]}.csv'

        try:
            with open(base_file, encoding="utf_16") as f_1:
                with open(new_file, 'w', encoding="utf_8") as f_2:
                    f_2.write(f_1.read())
        except Exception:
            pass

        return new_file

    def get_team_names(self, team_names_lst):
        '''
        docstr
        '''
        if len(team_names_lst) == 0:
            return []

        fst_lst = team_names_lst.split('&')
        fst_lst[-1] = fst_lst[-1].split('/')[0].strip()

        for name in fst_lst:
            name = re.sub('[\"\']', '', name)

        return fst_lst

    def process_file(self):
        '''
        docstr
        '''
        filename = self._filepath.split('/')[-1].split('.')[0]

        date_range_list = self.get_date_range()
        date_range = date_range_list[0] + ' - ' + date_range_list[1]

        db_add_excel_file([self._filepath, filename, date_range])

    def process_data(self):
        '''
        docstr
        '''
        try:
            data_frame = self.get_data_frame()
            col_index_lst = self.get_data_frame_cols(data_frame)
            self.get_values_from_row(data_frame, col_index_lst)
        except Exception:
            traceback.print_exc()

    def get_data_frame(self):
        '''
        docstr
        '''
        file_path = Path(self._filepath)
        file_extension = file_path.suffix.lower()[1:]

        if file_extension == 'xlsx':
            data_frame = pd.read_excel(self._filepath, engine='openpyxl')
        elif file_extension == 'xls':
            data_frame = pd.read_excel(self._filepath)
        elif file_extension == 'csv':
            data_frame = pd.read_csv(self._filepath, engine='python', sep=',',
                                     quoting=csv.QUOTE_NONE, quotechar='"',
                                     on_bad_lines='skip')
        else:
            raise Exception("File not supported")

        return data_frame

    def get_data_frame_cols(self, data_frame):
        '''
        docstr
        '''
        column_pd = list(data_frame.columns)
        column_pd = column_pd[0]
        column_list = column_pd.split('\t')
        formatted_col_list = []

        for column_name in column_list:
            col_name = re.sub('[\"\']', '', column_name)
            col_name.strip()
            formatted_col_list.append(col_name)

        team_names_column = formatted_col_list.index('SL - Remarks')
        booking_date_column = formatted_col_list.index('SL - Date and time of booking')
        job_type_column = formatted_col_list.index('SL - Resource name')
        transport_refrence_column = formatted_col_list.index('SL - Transport reference')
        job_start_time_column = formatted_col_list.index('SL - Custom timestamp 2 date and time ')
        job_end_time_column = formatted_col_list.index('SL - Custom timestamp 3 date and time ')
        in_or_out_column = formatted_col_list.index('SL - Location')

        return [team_names_column, booking_date_column, job_type_column, transport_refrence_column,
                job_start_time_column, job_end_time_column, in_or_out_column]

    def get_values_from_row(self, data_frame, col_index_lst):
        '''
        docstr
        '''
        data_frame = data_frame.reset_index()
        total_job_time = 0

        for row_set in data_frame.itertuples():
            row_str = row_set[2]
            row = row_str.split('\t')
            if len(row) > 6:
                team_names = self.get_team_names(row[col_index_lst[0]])

                booking_date = row[col_index_lst[1]]
                if len(booking_date) > 10:
                    booking_date = booking_date.split(' ')[0]

                job_type = row[col_index_lst[2]]
                job_type = re.sub('[\"\']', '', job_type)
                job_type = job_type.strip()

                transport_refrence = row[col_index_lst[3]]
                try:
                    if (row[col_index_lst[4]].find('AM') != -1
                       or row[col_index_lst[4]].find('PM') != -1):
                        row[col_index_lst[4]] = row[col_index_lst[4]][:-2].strip()

                    if (row[col_index_lst[5]].find('AM') != -1
                       or row[col_index_lst[5]].find('PM') != -1):
                        row[col_index_lst[5]] = row[col_index_lst[5]][:-2].strip()
                except Exception:
                    pass

                try:
                    job_start_time = datetime.datetime.strptime(row[col_index_lst[4]],
                                                                '%m/%d/%Y %H:%M')
                    job_end_time = datetime.datetime.strptime(row[col_index_lst[5]],
                                                              '%m/%d/%Y %H:%M')

                    if ((job_end_time < job_start_time) and
                            (job_end_time.date() == job_start_time.date())):

                        job_end_time = job_end_time + datetime.timedelta(hours=12)

                    total_job_time = job_end_time - job_start_time
                    total_job_time = int(total_job_time.total_seconds() / 60)
                    converted_total_job_time = str(total_job_time)
                except Exception:
                    converted_total_job_time = 'error'
                    # traceback.print_exc()

                try:
                    in_or_out_val = row[col_index_lst[6]]
                    if in_or_out_val.find('delivery') != -1:
                        in_or_out = 'Inbound'
                    else:
                        in_or_out = 'Outbound'
                except Exception:
                    pass

            if len(team_names) > 1 and job_type.find('20') != -1:
                job_type = 'Ocean Containers - 40 HQ'

            row_vals = [team_names, booking_date, job_type, converted_total_job_time,
                        transport_refrence, in_or_out]

            if ((total_job_time <= 480 and total_job_time >= 1)
               and converted_total_job_time != 'error'):
                self.add_entry_to_db(row_vals)

    def add_entry_to_db(self, row_vals):
        '''
        docstr
        '''
        team_names = row_vals[0]
        booking_date = row_vals[1]
        job_type = row_vals[2]
        converted_total_job_time = row_vals[3]
        transport_refrence = row_vals[4]
        in_or_out = row_vals[5]

        if isinstance(team_names, list) and len(team_names) != 0:
            db_add_performance_entry([team_names, booking_date, job_type, converted_total_job_time,
                                      transport_refrence, self._filepath, in_or_out], 'team')

            for person_name in team_names:
                if team_names.index(person_name) == 0:
                    worker_job = 'Driver'
                else:
                    worker_job = 'Unloader'

                db_add_performance_entry([person_name.strip(), booking_date, job_type,
                                          converted_total_job_time, transport_refrence,
                                          worker_job, len(team_names), self._filepath, in_or_out],
                                         'individual')
                self.date_range_list.append(booking_date)

    def get_date_range(self):
        '''
        docstr
        '''
        fin_date_range_list = ['', '']

        i = 0
        while i < len(self.date_range_list):
            try:
                self.date_range_list[i] = datetime.datetime.strptime(self.date_range_list[i],
                                                                     '%m/%d/%Y')
            except Exception:
                pass
            i += 1

        try:
            earliest_date = min(self.date_range_list)
            latest_date = max(self.date_range_list)

            fin_date_range_list[0] = datetime.datetime.strftime(earliest_date, '%m/%d/%Y')
            fin_date_range_list[1] = datetime.datetime.strftime(latest_date, '%m/%d/%Y')
        except Exception:
            print(self.date_range_list)

        return fin_date_range_list
