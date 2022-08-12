'''
docstr
'''
import datetime as dt
from datetime import date, timedelta
import timeit
import matplotlib.pyplot as plt
import multiprocessing as mp
import os
import re
from pathlib import Path
import numpy as np
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
from dateutil.relativedelta import relativedelta
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from reportlab.lib.units import inch
import traceback
import io
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from performance_review_db import db_get_individual_data, db_get_team_data

pdf_filepaths = []

def create_report_cards():
    '''
    docstr
    '''
    start = timeit.default_timer()

    data_lst = get_report_data()
    for data_entry in data_lst:
        try:
            create_report_pages(data_entry)
        except Exception:
            traceback.print_exc()

    final_report_fp = create_final_report(pdf_filepaths)
    stop = timeit.default_timer()
    print(f'\n\nDone, Time Ran: {(stop-start)/60} minutes')
    return final_report_fp


def allsundays(year):
    """This code was provided in the previous answer! It's not mine!"""
    d = date(year, 1, 1)
    d += timedelta(days=6 - d.weekday())
    while d.year == year:
        yield d
        d += timedelta(days=7)


def get_week_dates(year):
    '''
    docstr
    '''
    Dict = {}
    for w_n, d in enumerate(allsundays(year)):
        # This is my only contribution!
        Dict[w_n + 1] = [(d + timedelta(days=k)).isoformat() for k in range(0,7) ]

    return Dict


def get_report_data():
    '''
    docstr
    '''
    report_lst = []
    all_team_data = db_get_team_data()
    all_individual_data = db_get_individual_data()

    current_date = dt.date.today()

    week_start = current_date - timedelta(days=7)
    month_start = current_date - timedelta(days=30)
    year_start = current_date - timedelta(days=365)

    current_date_str = current_date.strftime('%m/%d/%y')
    week_start_str = week_start.strftime('%m/%d/%y')
    month_start_str = month_start.strftime('%m/%d/%y')
    year_start_str = year_start.strftime('%m/%d/%y')

    names_set = set()
    {names_set.add((entry[0], entry[5])) for entry in all_individual_data}

    for name_entry in names_set:
        name_entry = list(name_entry)

        data_lst_entry = {'Name': '', 'Job': '', 'Date': '',
                          'Week-Date-Range': '', 'Week-Data': [],
                          'Month-Date-Range': '', 'Month-Data': [],
                          'Year-Date-Range': '', 'Year-Data': []}

        data_lst_entry['Name'] = name_entry[0]
        data_lst_entry['Job'] = name_entry[1]
        data_lst_entry['Date'] = dt.date.today()

        data_lst_entry['Week-Date-Range'] = week_start_str + '-' + current_date_str
        data_lst_entry['Month-Date-Range'] = month_start_str + '-' + current_date_str
        data_lst_entry['Year-Date-Range'] = year_start_str + '-' + current_date_str

        week_data_lst = format_week_data(all_individual_data, name_entry[0], week_start)
        month_data_lst = format_month_data(all_individual_data, name_entry[0], month_start)
        team_data_lst = format_team_data(all_team_data, name_entry[0], month_start)
        year_data_lst = format_year_data(all_individual_data, name_entry[0], year_start)

        data_lst_entry['Week-Data'] = week_data_lst
        data_lst_entry['Month-Data'] = month_data_lst
        data_lst_entry['Team-Data'] = team_data_lst
        data_lst_entry['Year-Data'] = year_data_lst

        report_lst.append(data_lst_entry)

    return report_lst


def get_value_avgs(format_dict):
    '''
    docstr
    '''
    dict_keys = [*format_dict]
    dict_list = []
    for key in dict_keys:
        time_lst = []
        inner_dict = {'Job-Type': '', 'Average': '', 'Num-Jobs': '', 'Rank': ''}
        inner_dict['Job-Type'] = key

        for entry in format_dict[key]:
            time_lst.append(int(entry))

        num_jobs = len(time_lst)
        avg_time = round((sum(time_lst) / num_jobs), )

        inner_dict['Average'] = str(avg_time)
        inner_dict['Num-Jobs'] = num_jobs

        dict_list.append(inner_dict)

    return dict_list


def fix_team_list(dict_list):
    '''
    docstr
    '''
    sorted_job_lst = (sorted(dict_list, key=lambda item: (item['Num-Jobs'], item['Average'])))
    sorted_job_lst.reverse()
    return sorted_job_lst[:5]


def get_team_value_avgs(format_dict):
    '''
    docstr
    '''
    dict_keys = [*format_dict]
    dict_list = []
    for key in dict_keys:
        time_lst = []
        inner_dict = {'Team-Names': '', 'Average': '', 'Num-Jobs': ''}

        name_str = ''
        for name in key:
            if key.index(name) == len(key) - 1:
                name_str += name
            else:
                name_str += (name + ', ')

        inner_dict['Team-Names'] = name_str

        for entry in format_dict[key]:
            time_lst.append(int(entry))

        num_jobs = len(time_lst)
        avg_time = round((sum(time_lst) / num_jobs), )

        inner_dict['Average'] = str(avg_time)
        inner_dict['Num-Jobs'] = num_jobs

        dict_list.append(inner_dict)

    if len(dict_list) > 8:
        shortened_lst = fix_team_list(dict_list)
        return shortened_lst

    return dict_list


def get_year_value_avgs(format_dict):
    '''
    docstr
    '''
    ret_dict = {}
    for key in format_dict:
        ret_dict[key] = []
        month_avg_dict = {}
        for entry in format_dict[key]:
            entry_month = dt.datetime.strftime(entry[1], '%b %Y')
            if entry_month not in month_avg_dict:
                month_avg_dict[entry_month] = [int(entry[0])]
            else:
                month_avg_dict[entry_month].append(int(entry[0]))
        for month_key in month_avg_dict:
            avg_time = round((sum(month_avg_dict[month_key]) / len(month_avg_dict[month_key])), 1)
            month_avg_dict[month_key] = avg_time

        ret_dict[key].append(month_avg_dict)

    return ret_dict


def format_week_data(all_individual_data, name_entry, week_start):
    '''
    docstr
    '''
    ret_lst = []
    entry_dict = {}
    current_date = dt.date.today()

    for entry in all_individual_data:
        entry_dt = entry[1]

        if entry[0] == name_entry and (entry_dt >= week_start and entry_dt <= current_date):
            if entry[2] not in entry_dict:
                entry_dict[entry[2]] = [entry[4]]
            else:
                entry_dict[entry[2]].append(entry[4])

    ret_lst = get_value_avgs(entry_dict)
    return ret_lst


def format_month_data(all_individual_data, name_entry, month_start):
    '''
    docstr
    '''
    ret_lst = []
    entry_dict = {}
    current_date = dt.date.today()

    for entry in all_individual_data:
        entry_dt = entry[1]

        if entry[0] == name_entry and (entry_dt >= month_start and entry_dt <= current_date):
            if entry[2] not in entry_dict:
                entry_dict[entry[2]] = [entry[4]]
            else:
                entry_dict[entry[2]].append(entry[4])

    ret_lst = get_value_avgs(entry_dict)
    return ret_lst


def format_team_data(all_team_data, name_entry, month_start):
    '''
    pass
    '''
    ret_lst = []
    entry_dict = {}
    current_date = dt.date.today()

    for entry in all_team_data:
        entry_dt = entry[1]

        sanitzed_entries = []
        entry_names_old_lst = entry[0].split(',')
        for name_e in entry_names_old_lst:
            clean_name = name_e.strip().lower()
            sanitzed_entries.append(clean_name)

        entry_names_lst = []
        for tw_name in sanitzed_entries:
            clean_name = tw_name.strip()
            new_name = get_real_name(clean_name)
            entry_names_lst.append(new_name)

        if (name_entry.lower() in sanitzed_entries and len(entry_names_lst) > 1) and (entry_dt >= month_start and entry_dt <= current_date):
            sorted_entries = tuple(sorted(entry_names_lst))
            

            if sorted_entries not in entry_dict:
                entry_dict[sorted_entries] = [entry[4]]
            else:
                entry_dict[sorted_entries].append(entry[4])

    ret_lst = get_team_value_avgs(entry_dict)
    return ret_lst


def format_year_data(all_individual_data, name_entry, year_start):
    '''
    docstr
    '''
    months = []
    dt_months = []
    ret_lst = []
    entry_dict = {}
    sorted_ret_lst = []
    current_date = dt.date.today()

    for entry in all_individual_data:
        entry_dt = entry[1]

        if entry[0] == name_entry and (entry_dt >= year_start and entry_dt <= current_date):
            if entry[2] not in entry_dict:
                entry_dict[entry[2]] = [[entry[4], entry_dt]]
            else:
                entry_dict[entry[2]].append([entry[4], entry_dt])

    ret_lst = get_year_value_avgs(entry_dict)

    current_month = dt.datetime.today()

    for i in range(11, 0, -1):
        next_month = current_month + relativedelta(months=i)
        dt_months.append(next_month)
    for month in dt_months:
        months.append(month.strftime('%b'))

    months.reverse()
    months.append(current_month.strftime('%b'))

    sorted_ret_lst = dict_sort(ret_lst)

    return sorted_ret_lst


def dict_sort(unsort_dict):
    '''
    pass
    '''
    sorted_ret_dict = {}
    sorted_month_order = [None] * 12
    dt_months = []
    currentDate = dt.date.today()
    months_start_date = dt.date(currentDate.year, currentDate.month, 1)
    for i in range(11, 0, -1):
        next_month = months_start_date - relativedelta(months=i)
        dt_months.append(next_month)
    dt_months.append(months_start_date)
    dt_months.reverse()

    for job_type_entry in unsort_dict:
        inner_dict_lst = []
        inner_dict = unsort_dict[job_type_entry][0]

        months_in_dict = [*inner_dict]

        dt_data_months_lst = []
        [dt_data_months_lst.append(dt.datetime.strptime(month_ent, '%b %Y')) for month_ent in months_in_dict]

        dt_data_months_lst.sort()
        if len(dt_data_months_lst) > 12:
            overage = len(dt_data_months_lst) - 12
            dt_data_months_lst = dt_data_months_lst[overage:]


        for month_entry in dt_data_months_lst:
          try:
            sorted_index = dt_months.index(month_entry.date())
            if sorted_index != -1:
              sorted_month_order[sorted_index] = month_entry
          except Exception:
            traceback.print_exc()
  
        for entry in sorted_month_order:
          try:
            month_as_str = dt.datetime.strftime(entry, '%b %Y')
            
            data_val = unsort_dict[job_type_entry][0][month_as_str]
            inner_dict_lst.append(data_val)
          except Exception:
            inner_dict_lst.append(None)
  
        sorted_ret_dict[job_type_entry] = inner_dict_lst

    return sorted_ret_dict


def create_report_pages(page_data):
    '''
    docstr
    '''
    if page_data['Name'] != 'Company':
        front_canvas = create_front_page(page_data)
        full_report = merge_single_report(front_canvas, page_data)

        pdf_filepaths.append(full_report)
    else:
        full_report = create_company_report()
        pdf_filepaths.insert(0, full_report)


def get_sorted_months():
    '''
    doctsr
    '''
    dt_months = []
    months = []
    current_month = dt.datetime.today()

    for i in range(11, 0, -1):
        next_month = current_month + relativedelta(months=i)
        dt_months.append(next_month)
    for month in dt_months:
        months.append(month.strftime('%b'))

    months.reverse()
    months.append(current_month.strftime('%b'))

    return months


def get_real_name(tw_name_str):
    '''
    docstr
    '''
    name_dict = {
        'dee': 'Dee Afflick',
        'hector': 'Hector Rodriguez',
        'armany': 'Armany Marrero',
        'jose': 'Jose DeLeon',
        'richardc': 'Richard Cuadros',
        'collin': 'Collin Sager',
        'carlos1': 'Carlos Rodriguez Escobar',
        'carlos2': 'Carlos Rubio',
        'john': 'John Pacheco',
        'juan': 'Juan Delgado',
        'sebastian': 'Sebastian Rojas',
        'cesar': 'Cesar Contreras Sepulveda',
        'deury': 'Deury Lopez Jimenez',
        'manuel': 'Manuel Brito Camacho',
        'max': 'Max Ramos',
        'olieser': 'Olieser Hernandez Lopez',
        'oliver': 'Oliver Gonzalez Lopez',
        'orvin': 'Orvin Irizarry Hernandez',
        'ivan': 'Ivan Bonilla',
        'julio': 'Julio Gamboa',
        'kevin': 'Kevin Bonilla',
        'luis': 'Luis Otero Castro',
        'andres': 'Andres'
    }

    try:
        return name_dict[tw_name_str.lower()]
    except Exception:
        return tw_name_str


def create_front_page(page_data):
    '''
    docstr
    '''
    if os.path.exists(os.path.abspath
                      ('Ecopax-Performance-Reviewer-Program-Files\\Report Card Cache')):
        f_path = os.path.abspath(
            'Ecopax-Performance-Reviewer-Program-Files\\Report Card Cache')
    else:
        f_path = os.path.abspath('Report Card Cache')

    packet = io.BytesIO()
    output_fp = f_path + '\\' + page_data['Name'] + '-front.pdf'
    report_canvas = canvas.Canvas(packet, pagesize=letter)

    pdfmetrics.registerFont(TTFont('Calibri', 'Calibri.ttf'))

    real_name_str = get_real_name(page_data['Name'])

    report_canvas.setFont('Calibri', 12)
    report_canvas.drawString(95, 697, real_name_str)
    report_canvas.drawString(77, 674.5, page_data['Job'])
    report_canvas.drawString(87, 657.5, dt.datetime.strftime(page_data['Date'], '%m/%d/%Y'))

    report_canvas.setFont('Calibri', 15)
    report_canvas.drawString(135, 623, page_data['Week-Date-Range'])
    report_canvas.drawString(140, 489, page_data['Month-Date-Range'])
    report_canvas.drawString(127, 232, page_data['Year-Date-Range'])

    report_canvas.setFont('Calibri', 12)
    week_start_pixel_y = 582

    for week_data in page_data['Week-Data']:
        report_canvas.drawString(54, week_start_pixel_y, week_data['Job-Type'])
        report_canvas.drawString(287, week_start_pixel_y, str(week_data['Average']))
        report_canvas.drawString(422, week_start_pixel_y, str(week_data['Num-Jobs']))
        report_canvas.drawString(494, week_start_pixel_y, week_data['Rank'])
        week_start_pixel_y -= 30

    month_start_pixel_y = 447

    for month_data in page_data['Month-Data']:
        report_canvas.drawString(54, month_start_pixel_y, month_data['Job-Type'])
        report_canvas.drawString(422, month_start_pixel_y, str(month_data['Average']))
        report_canvas.drawString(514, month_start_pixel_y, str(month_data['Num-Jobs']))
        report_canvas.drawString(494, month_start_pixel_y, month_data['Rank'])
        month_start_pixel_y -= 30

    for team_data in page_data['Team-Data']:
        report_canvas.drawString(54, month_start_pixel_y, team_data['Team-Names'])
        report_canvas.drawString(422, month_start_pixel_y, str(team_data['Average']))
        report_canvas.drawString(514, month_start_pixel_y, str(team_data['Num-Jobs']))
        month_start_pixel_y -= 20

    months_to_write = get_sorted_months()

    report_canvas.setFont('Calibri', 10)
    month_str_pixel_x = 140
    for month_str in months_to_write:
        report_canvas.drawString(month_str_pixel_x, 183, month_str)
        month_str_pixel_x += 40

    year_start_pixel_y = 127
    for key in page_data['Year-Data']:
        report_canvas.drawString(13, year_start_pixel_y, key)
        year_data_pixel_x = 580
        year_avg_data = page_data['Year-Data'][key]
        for month_key in year_avg_data:
            if isinstance(month_key, float):
                report_canvas.drawString(year_data_pixel_x, year_start_pixel_y, str(month_key))
            year_data_pixel_x -= 40
        year_start_pixel_y -= 30

    # make modifications here
    report_canvas.save()
    packet.seek(0)

    if os.path.exists(os.path.abspath
                      ('Ecopax-Performance-Reviewer-Program-Files\\program-dependables\\(PDF)Transwide Report Card PDF.pdf')):
        template_pdf_fp = os.path.abspath(
            'Ecopax-Performance-Reviewer-Program-Files\\program-dependables\\(PDF)Transwide Report Card PDF.pdf')
    else:
        template_pdf_fp = os.path.abspath('program-dependables\\(PDF)Transwide Report Card PDF.pdf')

    new_pdf = PdfFileReader(packet)
    try:
        exisiting_pdf = PdfFileReader(open(template_pdf_fp, 'rb'))
    except Exception:
        pass

    output = PdfFileWriter()

    template_page = exisiting_pdf.getPage(0)
    template_page.mergePage(new_pdf.getPage(0))

    output.addPage(template_page)

    try:
        output_stream = open(output_fp, 'wb')
        output.write(output_stream)
        output_stream.close()
    except Exception:
        pass

    return output_fp


def get_back_page():
    '''
    return back page filepath
    '''
    if os.path.exists(os.path.abspath
                      ('Ecopax-Performance-Reviewer-Program-Files\\program-dependables\\(PDF)Transwide Report Card Back.pdf')):
        output_fp = os.path.abspath(
            'Ecopax-Performance-Reviewer-Program-Files\\program-dependables\\(PDF)Transwide Report Card Back.pdf')
    else:
        output_fp = os.path.abspath('program-dependables\\(PDF)Transwide Report Card Back.pdf')
    return output_fp


def create_final_report(pdf_filepaths):
    '''
    return fulepath
    '''
    pdf_merger = PdfFileMerger()

    if os.path.exists(os.path.abspath
                      ('Ecopax-Performance-Reviewer-Program-Files\\Report Card Reports')):
        f_path = os.path.abspath(
            'Ecopax-Performance-Reviewer-Program-Files\\Report Card Reports')
    else:
        f_path = os.path.abspath('Report Card Reports')

    f_name = '\\ReportCardReport-' + dt.datetime.now().strftime('%m-%d-%y-%H-%M-%S') + '.pdf'

    output_fp = f_path + f_name
    pdf_filepaths = sorted(list(set(pdf_filepaths)))
    for pdf_fp in pdf_filepaths:
        try:
            pdf_merger.append(pdf_fp)
        except Exception:
            pass

    with Path(output_fp).open(mode='wb') as output_file:
        pdf_merger.write(output_file)

    return output_fp


def merge_single_report(front_canvas, page_data):
    '''
    return filepath
    '''
    back_page = get_back_page()

    pdf_merger = PdfFileMerger()
    if os.path.exists(os.path.abspath
                    ('Ecopax-Performance-Reviewer-Program-Files\\Report Card Cache')):
        f_path = os.path.abspath(
            'Ecopax-Performance-Reviewer-Program-Files\\Report Card Cache')
    else:
        f_path = os.path.abspath('Report Card Cache')

    f_name = page_data['Name'] + '-fullreport.pdf'

    output_fp = f_path + '\\' + f_name
    try:
        pdf_merger.append(front_canvas)
        pdf_merger.append(back_page)

        with Path(output_fp).open(mode='wb') as output_file:
            pdf_merger.write(output_file)
    except Exception:
        pass

    return output_fp


def create_company_report():
    '''
    return filepath
    '''
    output_fp = ''
    return output_fp
