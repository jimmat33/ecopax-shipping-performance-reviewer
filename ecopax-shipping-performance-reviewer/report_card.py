'''
docstr
'''
import datetime as dt
import multiprocessing as mp
import numpy as np
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileMerger


def create_report_cards():
    '''
    docstr
    '''
    data_lst = get_report_data()
    current_timestamp = dt.datetime.now()
    process_lst = []

    for data_entry in data_lst:
        pdf_proc = mp.Process(target=create_report_pages, args=(data_entry, current_timestamp,))
        process_lst.append(pdf_proc)

    num_cores = mp.cpu_count

    split_proc_list = np.array_split(process_lst, num_cores)

    for lst in split_proc_list:
        for pdf_process in lst:
            pdf_process.start()

    for lst in split_proc_list:
        for pdf_process in lst:
            pdf_process.join()


def get_report_data():
    '''
    docstr
    '''
    data_lst = [{'Name': 'Jimmy Mattison',
                 'Job': 'Driver',
                 'Date': '7/18/2022',
                 'Week-Data': [
                        {'Job-Type': '40\' Ocean Container', 'Average': '90', 'Num-Jobs': '4', 'Rank': '1'},
                        {'Job-Type': '20\' Ocean Container', 'Average': '60', 'Num-Jobs': '2', 'Rank': '4'},
                        {'Job-Type': 'Pallet Load', 'Average': '40', 'Num-Jobs': '12', 'Rank': '1'}],
                 'Month-Data': [
                        {'Job-Type': '40\' Ocean Container', 'Average': '90', 'Num-Jobs': '4', 'Rank': '1'},
                        {'Job-Type': '20\' Ocean Container', 'Average': '60', 'Num-Jobs': '2', 'Rank': '4'},
                        {'Job-Type': 'Pallet Load', 'Average': '40', 'Num-Jobs': '12', 'Rank': '1'}],
                 'Team-Data': [
                        {'Team-Names': 'Jimmy, James, Dean', 'Average': '90', 'Num-Jobs': '4'},
                        {'Team-Names': 'Jimmy, Mike, Melanie', 'Average': '80', 'Num-Jobs': '2'}],
                 'Year-Data': [
                        {'Job-Type': '40\' Ocean Container', 'Month': 'Jan', 'Avg-Time': '150'},
                        {'Job-Type': '40\' Ocean Container', 'Month': 'Feb', 'Avg-Time': '147'}]}]

    return data_lst

def create_report_pages(page_data, current_timestamp):
    '''
    docstr
    '''
    f_path  = ''
    canvas_name = page_data['Name'] + '-report-front'
    canvas_path = f_path + '\\' + canvas_name
    report_canvas = canvas.Canvas(canvas_path, pagesize=letter)
    back_page = get_back_page()

    if page_data['Name'] != 'Company':
        create_front_page(report_canvas)
    else:
        create_company_report()

def create_front_page(report_canvas):
    '''
    docstr
    '''
    pass


def get_back_page():
    '''
    docstr
    '''
    pass


def create_company_report():
    '''
    docstr
    '''
    pass