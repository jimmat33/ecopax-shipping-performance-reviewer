'''
docstr
'''
import datetime as dt
from datetime import date, timedelta
import matplotlib.pyplot as plt
import multiprocessing as mp
import os
import re
from pathlib import Path
import numpy as np
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from reportlab.lib.units import inch
import io
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from performance_review_db import db_get_individual_data, db_get_team_data

pdf_filepaths = []


def create_report_cards():
    '''
    docstr
    '''
    page_data_full = get_report_data()
    page_data = page_data_full[0]

    create_report_pages(page_data)


    '''
    pdf_filepaths.clear()
    data_lst = get_report_data()
    process_lst = []

    for data_entry in data_lst:
        pdf_proc = mp.Process(target=create_report_pages, args=(data_entry,))
        process_lst.append(pdf_proc)

    num_cores = mp.cpu_count

    split_proc_list = np.array_split(process_lst, num_cores)

    for lst in split_proc_list:
        for pdf_process in lst:
            pdf_process.start()

    for lst in split_proc_list:
        for pdf_process in lst:
            pdf_process.join()

    final_report_fp = create_final_report()
    return final_report_fp
    '''


def allsundays(year):
    """This code was provided in the previous answer! It's not mine!"""
    d = date(year, 1, 1)                    # January 1st                                                          
    d += timedelta(days=6 - d.weekday())  # First Sunday                                                         
    while d.year == year:
        yield d
        d += timedelta(days=7)


def get_week_dates(year):
    '''
    docstr
    '''
    Dict = {}
    for wn,d in enumerate(allsundays(year)):
        # This is my only contribution!
        Dict[wn+1] = [(d + timedelta(days=k)).isoformat() for k in range(0,7) ]

    return Dict


def get_report_data():
    '''
    docstr
    '''
    all_team_data = db_get_team_data()
    all_individual_data = db_get_individual_data()

    current_date = dt.date.today()

    week_start = current_date - timedelta(days=30)
    month_start = current_date - timedelta(days=30)
    year_start = current_date - timedelta(days=365)

    current_date_str = current_date.strftime('%m/%d/%y')
    week_start_str = week_start.strftime('%m/%d/%y')
    month_start_str = month_start.strftime('%m/%d/%y')
    year_start_str = year_start.strftime('%m/%d/%y')

    names_set = set()
    {names_set.add([entry[0], entry[5]]) for entry in all_individual_data}

    for name_entry in names_set:
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

        

    data_lst = [{'Name': 'Jimmy Mattison',
                 'Job': 'Driver',
                 'Date': '7/18/2022',
                 'Week-Date-Range': '12/17/22-12/23/22',
                 'Week-Data': [
                        {'Job-Type': '40\' Ocean Container', 'Average': '90', 'Num-Jobs': '4', 'Rank': '1'},
                        {'Job-Type': '20\' Ocean Container', 'Average': '60', 'Num-Jobs': '2', 'Rank': '4'},
                        {'Job-Type': 'Pallet Load', 'Average': '40', 'Num-Jobs': '12', 'Rank': '1'}],
                 'Month-Date-Range': '13/10/22-13/20/22',
                 'Month-Data': [
                        {'Job-Type': '40\' Ocean Container', 'Average': '90', 'Num-Jobs': '4', 'Rank': '1'},
                        {'Job-Type': '20\' Ocean Container', 'Average': '60', 'Num-Jobs': '2', 'Rank': '4'},
                        {'Job-Type': 'Pallet Load', 'Average': '40', 'Num-Jobs': '12', 'Rank': '1'}],
                 'Team-Data': [
                        {'Team-Names': 'Jimmy, James, Dean', 'Average': '90', 'Num-Jobs': '4'},
                        {'Team-Names': 'Jimmy, Mike, Melanie', 'Average': '80', 'Num-Jobs': '2'}],
                 'Year-Date-Range': '13/10/22-13/20/22',
                 'Year-Data': [
                        {'Job-Type': '40\' Ocean Container', 'Month': 'Jan', 'Avg-Time': '150'},
                        {'Job-Type': '40\' Ocean Container', 'Month': 'Feb', 'Avg-Time': '147'},
                        {'Job-Type': '40\' Ocean Container', 'Month': 'Mar', 'Avg-Time': '130'},
                        {'Job-Type': '40\' Ocean Container', 'Month': 'Apr', 'Avg-Time': '150'},
                        {'Job-Type': '40\' Ocean Container', 'Month': 'May', 'Avg-Time': '140'},
                        {'Job-Type': '40\' Ocean Container', 'Month': 'Jun', 'Avg-Time': '145'},
                        {'Job-Type': '40\' Ocean Container', 'Month': 'Jul', 'Avg-Time': '120'},
                        {'Job-Type': '20\' Ocean Container', 'Month': 'Jan', 'Avg-Time': '70'},
                        {'Job-Type': '20\' Ocean Container', 'Month': 'Feb', 'Avg-Time': '60'},
                        {'Job-Type': '20\' Ocean Container', 'Month': 'Mar', 'Avg-Time': '65'},
                        {'Job-Type': '20\' Ocean Container', 'Month': 'Apr', 'Avg-Time': '78'},
                        {'Job-Type': '20\' Ocean Container', 'Month': 'May', 'Avg-Time': '60'},
                        {'Job-Type': '20\' Ocean Container', 'Month': 'Jun', 'Avg-Time': '55'},
                        {'Job-Type': '20\' Ocean Container', 'Month': 'Jul', 'Avg-Time': '40'},
                        {'Job-Type': 'Pallet Load', 'Month': 'Jan', 'Avg-Time': '50'},
                        {'Job-Type': 'Pallet Load', 'Month': 'Feb', 'Avg-Time': '44'},
                        {'Job-Type': 'Pallet Load', 'Month': 'Mar', 'Avg-Time': '42'},
                        {'Job-Type': 'Pallet Load', 'Month': 'Apr', 'Avg-Time': '61'},
                        {'Job-Type': 'Pallet Load', 'Month': 'May', 'Avg-Time': '47'},
                        {'Job-Type': 'Pallet Load', 'Month': 'Jun', 'Avg-Time': '38'},
                        {'Job-Type': 'Pallet Load', 'Month': 'Jul', 'Avg-Time': '30'}]}]

    return data_lst


def format_week_data(all_individual_data, name_entry, week_start):
    '''
    docstr
    '''
    return []


def format_month_data(all_individual_data, name_entry, month_start):
    '''
    docstr
    '''
    return []


def format_team_data(all_team_data, name_entry, month_start):
    '''
    pass
    '''
    return []


def format_year_data(all_individual_data, name_entry, year_start):
    '''
    docstr
    '''
    return []


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


def create_front_page(page_data):
    '''
    docstr
    '''
    if os.path.exists(os.path.abspath
                      ('Ecopax-Performance-Reviwer-Program-Files\\Report Card Cache')):
        f_path = os.path.abspath(
            'Ecopax-Performance-Reviwer-Program-Files\\Report Card Cache')
    else:
        f_path = os.path.abspath('Report Card Cache')

    packet = io.BytesIO()
    output_fp = f_path + '\\' + page_data['Name'] + '-front.pdf'
    report_canvas = canvas.Canvas(packet, pagesize=letter)

    pdfmetrics.registerFont(TTFont('Calibri', 'Calibri.ttf'))

    report_canvas.setFont('Calibri', 12)
    report_canvas.drawString(95, 697, page_data['Name'])
    report_canvas.drawString(77, 674.5, page_data['Job'])
    report_canvas.drawString(87, 657.5, page_data['Date'])

    report_canvas.setFont('Calibri', 15)
    report_canvas.drawString(135, 623, page_data['Week-Date-Range'])
    report_canvas.drawString(140, 489, page_data['Month-Date-Range'])
    report_canvas.drawString(127, 232, page_data['Year-Date-Range'])

    report_canvas.setFont('Calibri', 12)
    week_start_pixel_y = 582

    for week_data in page_data['Week-Data']:
        report_canvas.drawString(54, week_start_pixel_y, week_data['Job-Type'])
        report_canvas.drawString(287, week_start_pixel_y, week_data['Average'])
        report_canvas.drawString(422, week_start_pixel_y, week_data['Num-Jobs'])
        report_canvas.drawString(494, week_start_pixel_y, week_data['Rank'])
        week_start_pixel_y -= 30

    month_start_pixel_y = 447

    for month_data in page_data['Month-Data']:
        report_canvas.drawString(54, month_start_pixel_y, month_data['Job-Type'])
        report_canvas.drawString(287, month_start_pixel_y, month_data['Average'])
        report_canvas.drawString(422, month_start_pixel_y, month_data['Num-Jobs'])
        report_canvas.drawString(494, month_start_pixel_y, month_data['Rank'])
        month_start_pixel_y -= 30

    for team_data in page_data['Team-Data']:
        report_canvas.drawString(54, month_start_pixel_y, team_data['Team-Names'])
        report_canvas.drawString(287, month_start_pixel_y, team_data['Average'])
        report_canvas.drawString(422, month_start_pixel_y, team_data['Num-Jobs'])
        month_start_pixel_y -= 20

    graph_fp = create_graph(page_data)
    report_canvas.drawImage(graph_fp, 0.55 * inch, 0.4 * inch)

    # make modifications here
    report_canvas.save()
    packet.seek(0)

    if os.path.exists(os.path.abspath
                      ('Ecopax-Performance-Reviwer-Program-Files\\program-dependables\\(PDF)Transwide Report Card PDF.pdf')):
        template_pdf_fp = os.path.abspath(
            'Ecopax-Performance-Reviwer-Program-Files\\program-dependables\\(PDF)Transwide Report Card PDF.pdf')
    else:
        template_pdf_fp = os.path.abspath('program-dependables\\(PDF)Transwide Report Card PDF.pdf')

    new_pdf = PdfFileReader(packet)
    exisiting_pdf = PdfFileReader(open(template_pdf_fp, 'rb'))

    output = PdfFileWriter()

    template_page = exisiting_pdf.getPage(0)
    template_page.mergePage(new_pdf.getPage(0))

    output.addPage(template_page)

    output_stream = open(output_fp, 'wb')
    output.write(output_stream)
    output_stream.close()

    return output_fp


def get_back_page():
    '''
    return back page filepath
    '''
    if os.path.exists(os.path.abspath
                      ('Ecopax-Performance-Reviwer-Program-Files\\program-dependables\\(PDF)Transwide Report Card Back.pdf')):
        output_fp = os.path.abspath(
            'Ecopax-Performance-Reviwer-Program-Files\\program-dependables\\(PDF)Transwide Report Card Back.pdf')
    else:
        output_fp = os.path.abspath('program-dependables\\(PDF)Transwide Report Card Back.pdf')
    return output_fp


def create_company_report():
    '''
    return fulepath
    '''
    output_fp = ''
    return output_fp


def merge_single_report(front_canvas, page_data):
    '''
    return filepath
    '''
    back_page = get_back_page()

    pdf_merger = PdfFileMerger()
    if os.path.exists(os.path.abspath
                      ('Ecopax-Performance-Reviwer-Program-Files\\Report Card Cache')):
        f_path = os.path.abspath(
            'Ecopax-Performance-Reviwer-Program-Files\\Report Card Cache')
    else:
        f_path = os.path.abspath('Report Card Cache')

    f_name = page_data['Name'] + '-fullreport.pdf'

    output_fp = f_path + '\\' + f_name

    pdf_merger.append(front_canvas)
    pdf_merger.append(back_page)

    with Path(output_fp).open(mode='wb') as output_file:
        pdf_merger.write(output_file)

    return output_fp

def create_final_report():
    '''
    return filepath
    '''
    output_fp = ''
    return output_fp


def create_graph(page_data):
    '''
    return fp
    '''
    if os.path.exists(os.path.abspath
                      ('Ecopax-Performance-Reviwer-Program-Files\\Report Card Cache')):
        f_path = os.path.abspath(
            'Ecopax-Performance-Reviwer-Program-Files\\Report Card Cache')
    else:
        f_path = os.path.abspath('Report Card Cache')

    output_fp = f_path + '\\' + page_data['Name'] + '-graph.png'

    graph_data = page_data['Year-Data']
    formatted_data = format_graph_data(graph_data)

    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    acceptable_formats = ['-', '-.', ':', '--', '-', '-.', ':', '--']
    acceptable_markers = ['o', 'v', 's', '*', '+', 'o', 'v', 's', '*', '+']

    plt.figure(figsize=(5.75, 2))
    for dataset in formatted_data:
        type_name = re.sub('[\"]', '', dataset[0])
        dataset_vals = dataset[1:]
        plt.plot(np.array(months[:len(dataset_vals)]), np.array(dataset_vals), acceptable_formats[formatted_data.index(dataset)], label=type_name, marker=acceptable_markers[formatted_data.index(dataset)], color='black')

    plt.legend(prop={'size': 7})
    plt.savefig(output_fp)

    return output_fp


def format_graph_data(graph_data):
    '''
    docstr
    '''
    job_types = set()
    lst_cmp = {job_types.add(entry['Job-Type']) for entry in graph_data}
    num_job_type = len(job_types)

    data_lst = []

    for i in range(num_job_type):
        current_job_type = list(job_types)[i]
        inner_lst = []
        inner_lst.append(current_job_type)
        for item in graph_data:
            if item['Job-Type'] == current_job_type:
                inner_lst.append(int(item['Avg-Time']))
        data_lst.append(inner_lst)

    return data_lst
