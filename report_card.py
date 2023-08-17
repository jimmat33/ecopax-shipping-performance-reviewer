import pandas as pd
import csv
from pathlib import Path
import datetime
import re
import os


def get_data_from_tw_excel_sheet(user_input_fp):
    """
    """
    # Fixing the file
    file_to_read = fix_file_ext(user_input_fp)

    data_frame = get_data_frame(file_to_read)
    col_index_dict = get_default_tw_columns()
    file_data_lst = get_values_from_row(data_frame, col_index_dict)


def get_data_frame(excel_fp):
    """
    """
    file_path = Path(excel_fp)
    file_extension = file_path.suffix.lower()[1:]

    if file_extension == 'xlsx':
        data_frame = pd.read_excel(excel_fp, engine='openpyxl')
    elif file_extension == 'xls':
        data_frame = pd.read_excel(excel_fp, engine='openpyxl')
    elif file_extension == 'csv':
        data_frame = pd.read_csv(excel_fp, engine='python', sep=',',
                                 quoting=csv.QUOTE_NONE, quotechar='"',
                                 on_bad_lines='skip')
    else:
        raise Exception("File not supported")

    return data_frame


def fix_file_ext(filepath):
    """
    """
    name = filepath.split('.')
    file_name_wo_ext = ''.join(name[:-1])
    new_file = f'{file_name_wo_ext}.csv'

    with open(filepath, encoding="utf_16") as f_1:
        with open(new_file, 'w', encoding="utf_8") as f_2:
            f_2.write(f_1.read())

    return new_file


def get_default_tw_columns():
    """
    """
    alphabet = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k',
                'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v',
                'w', 'x', 'y', 'z']

    file_cols_dict = {
        'Team Names': -1,
        'Booking Date': -1,
        'Job Type': -1,
        'Job Start Time': -1,
        'Job End Time': -1,
        'Pickup or Delivery': -1}

    tw_columns_fp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tw_ship_tool_def_columns.txt')

    file_lines = []
    with open(tw_columns_fp, 'r') as tw_file:
        file_data = tw_file.read()
        file_lines = file_data.split('\n')

    for lin in file_lines:
        if 'Team Names' in lin:
            line_val = str(lin.split('=')[1]).strip()
            if line_val.isnumeric():
                file_cols_dict['Team Names'] = int(line_val)
            else:
                file_cols_dict['Team Names'] = alphabet.index(line_val.lower())
        elif 'Booking Date' in lin:
            line_val = str(lin.split('=')[1]).strip()
            if line_val.isnumeric():
                file_cols_dict['Booking Date'] = int(line_val)
            else:
                file_cols_dict['Booking Date'] = alphabet.index(line_val.lower())
        elif 'Job Type' in lin:
            line_val = str(lin.split('=')[1]).strip()
            if line_val.isnumeric():
                file_cols_dict['Job Type'] = int(line_val)
            else:
                file_cols_dict['Job Type'] = alphabet.index(line_val.lower())
        elif 'Job Start Time' in lin:
            line_val = str(lin.split('=')[1]).strip()
            if line_val.isnumeric():
                file_cols_dict['Job Start Time'] = int(line_val)
            else:
                file_cols_dict['Job Start Time'] = alphabet.index(line_val.lower())
        elif 'Job End Time' in lin:
            line_val = str(lin.split('=')[1]).strip()
            if line_val.isnumeric():
                file_cols_dict['Job End Time'] = int(line_val)
            else:
                file_cols_dict['Job End Time'] = alphabet.index(line_val.lower())
        elif 'Pickup or Delivery' in lin:
            line_val = str(lin.split('=')[1]).strip()
            if line_val.isnumeric():
                file_cols_dict['Pickup or Delivery'] = int(line_val)
            else:
                file_cols_dict['Pickup or Delivery'] = alphabet.index(line_val.lower())

    return file_cols_dict


def get_team_names(team_names_lst):
    """
    """
    if len(team_names_lst) == 0:
        return []

    fst_lst = team_names_lst.split('&')
    fst_lst[-1] = fst_lst[-1].split('/')[0].strip()

    for name in fst_lst:
        name = re.sub('[\"\']', '', name)

    return fst_lst


def get_values_from_row(data_frame, col_index_lst):
    """
    """
    data_frame = data_frame.reset_index()
    total_job_time = 0

    for row_set in data_frame.itertuples():
        row_str = row_set[2]
        row = row_str.split('\t')
        if len(row) > 6:
            team_names = get_team_names(row[col_index_lst['Team Names']])

            booking_date = row[col_index_lst['Booking Date']]
            if len(booking_date) > 10:
                booking_date = booking_date.split(' ')[0]

            job_type = row[col_index_lst['Job Type']]
            job_type = re.sub('[\"\']', '', job_type)
            job_type = job_type.strip()

            try:
                if (row[col_index_lst[4]].find('AM') != -1
                    or row[col_index_lst[4]].find('PM') != -1):
                    row[col_index_lst[4]] = row[col_index_lst['Job Start Time']][:-2].strip()

                if (row[col_index_lst[5]].find('AM') != -1
                    or row[col_index_lst[5]].find('PM') != -1):
                    row[col_index_lst[5]] = row[col_index_lst['Job End Time']][:-2].strip()
            except Exception:
                pass

            try:
                job_start_time = datetime.datetime.strptime(row[col_index_lst['Job Start Time']],
                                                            '%m/%d/%Y %H:%M')
                job_end_time = datetime.datetime.strptime(row[col_index_lst['Job End Time']],
                                                          '%m/%d/%Y %H:%M')

                if ((job_end_time < job_start_time)
                    and (job_end_time.date() == job_start_time.date())):
                    job_end_time = job_end_time + datetime.timedelta(hours=12)

                total_job_time = job_end_time - job_start_time
                total_job_time = int(total_job_time.total_seconds() / 60)
                converted_total_job_time = str(total_job_time)
            except Exception:
                converted_total_job_time = 'error'
                # traceback.print_exc()

            try:
                in_or_out_val = row[col_index_lst['Pickup or Delivery']]
                if in_or_out_val.find('delivery') != -1:
                    in_or_out = 'Inbound'
                else:
                    in_or_out = 'Outbound'
            except Exception:
                pass

        if len(team_names) > 1 and job_type.find('20') != -1:
            job_type = 'Ocean Containers - 40 HQ'

        row_vals = [team_names, booking_date, job_type, converted_total_job_time, in_or_out]

        if ((total_job_time <= 480 and total_job_time >= 1)
            and converted_total_job_time != 'error'):
            print()
