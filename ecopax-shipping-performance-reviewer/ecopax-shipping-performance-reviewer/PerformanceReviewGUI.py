import tkinter as tk
from tkinter import *
import tkinter.font
import os
from tkinter.tix import Select
from tkinter.ttk import *
from  tkinter import ttk
from tkcalendar import Calendar, DateEntry
from tkinter import filedialog
from TWExportExcelFile import *
import time
from PerformanceReviewDB import *

class PerformanceReviewGUI(object):
    def __init__(self):
        self.root = Tk()
        self.root.geometry('800x600')
        self.root.title("Shipping Performance Reviewer")
        self.root.resizable(False, False)
        img = tk.PhotoImage(file= (os.path.abspath('gui_icon.png')))
        self.root.tk.call('wm', 'iconphoto', self.root._w, img)
        self.excel_file_list = []
        self.excel_index = 1
        self.progress_bar = 0
        self.progress_bar_label = 0
        self.excel_sheet_frame = 0
        self.worker_job_options = [] #get from database
        self.worker_team_options = [] #get from database
        self.job_type_options = [] #get from database
        self.individual_worker_options = [] #get from database


    def run_gui(self):
        self.init_widgits()
        self.root.mainloop()


    def init_widgits(self):
        widget_font = tk.font.Font(family = 'MS Shell Dlg 2', size = 10)
        large_widget_font = tk.font.Font(family = 'MS Shell Dlg 2', size = 14)
        dropdown_font = tk.font.Font(family = 'MS Shell Dlg 2', size = 10)

        initial_dropdown = StringVar(value='Click to Select a Value')

#import spreadsheet button
        self.import_sheet_button = Button(self.root, text = 'Import Spreadsheet', state = 'normal', command = self.import_spreadsheet_btn_click)
        self.import_sheet_button.place(x = 25, y = 35, width = 220, height = 55)


#remove spreadsheet button
        self.remove_sheet_button = Button(self.root, text = 'Remove Spreadsheet', state = 'normal', command = self.remove_spreadsheet_btn_click)
        self.remove_sheet_button.place(x = 25, y = 105, width = 220, height = 55)

        '''
 #table select all checkbox
        self.check_var = IntVar(value=0)
        self.check_all_excel_chkbox = Checkbutton(self.root, text = 'Select all Excel files',variable = self.check_var, onvalue = 1, offvalue = 0, command=self.select_all_excel)
        self.check_all_excel_chkbox.place(x = 300, y = 10, width = 400, height = 25)
        #if deleting all, do popup check
        '''

#excel sheet table
        self.excel_sheet_frame = Frame(self.root)
        self.excel_sheet_frame.place(x = 300, y = 35, width = 485, height = 240)

            #table scrollbars
        excel_vertical_scroll = Scrollbar(self.excel_sheet_frame)
        excel_vertical_scroll.pack(side=RIGHT, fill=Y)

        excel_horizontal_scroll = Scrollbar(self.excel_sheet_frame,orient='horizontal')
        excel_horizontal_scroll.pack(side= BOTTOM,fill=X)

        self.excel_sheet_frame = ttk.Treeview(self.excel_sheet_frame,yscrollcommand=excel_vertical_scroll.set, xscrollcommand =excel_horizontal_scroll.set)

        excel_vertical_scroll.config(command=self.excel_sheet_frame.yview)
        excel_horizontal_scroll.config(command=self.excel_sheet_frame.xview)
        
            #setting up table
        self.excel_sheet_frame['columns'] = ('file_name', 'file_location', 'date_range')
        self.excel_sheet_frame.column('#0', width=0, stretch=NO)
        self.excel_sheet_frame.column('file_name', anchor=CENTER, width = 130)
        self.excel_sheet_frame.column('file_location', anchor=CENTER, width = 205)
        self.excel_sheet_frame.column('date_range', anchor=CENTER, width = 130)

        self.excel_sheet_frame.heading('#0', text = '', anchor=CENTER)
        self.excel_sheet_frame.heading('file_name', text = 'File Name', anchor=CENTER)
        self.excel_sheet_frame.heading('file_location', text = 'File Location', anchor=CENTER)
        self.excel_sheet_frame.heading('date_range', text = 'Date Range', anchor=CENTER)

        self.excel_sheet_frame.pack()
            #check database for things to import here if james wants persistent data


#report settings label
        self.report_settings_label = Label(self.root, text= 'Generate Report Queries:', font = large_widget_font, state = 'normal')
        self.report_settings_label.place(x = 25, y = 285, height = 25)


#job type dropdown
        self.job_type_label = Label(self.root, text= 'Job Type: ', font = dropdown_font, state = 'normal')
        self.job_type_label.place(x = 25, y = 320, height = 25)

        self.job_type_dropdown = Combobox(self.root, values=self.job_type_options)
        self.job_type_dropdown.place(x = 90, y = 320, width = 180, height = 25)


#worker team dropdown
        self.worker_team_label = Label(self.root, text= 'Worker Team: ', font = dropdown_font, state = 'normal')
        self.worker_team_label.place(x = 25, y = 355, height = 25)

        self.worker_team_dropdown = Combobox(self.root, values=self.worker_team_options)
        self.worker_team_dropdown.place(x = 120, y = 355, width = 180, height = 25)


#individual worker dropdown
        self.individual_worker_label = Label(self.root, text= 'Individual Worker: ', font = dropdown_font, state = 'normal')
        self.individual_worker_label.place(x = 25, y = 390, height = 25)

        self.individual_worker_dropdown = Combobox(self.root, values=self.individual_worker_options)
        self.individual_worker_dropdown.place(x = 140, y = 390, width = 180, height = 25)


#worker job dropdown
        self.worker_job_label = Label(self.root, text= 'Worker Job: ', font = dropdown_font, state = 'normal')
        self.worker_job_label.place(x = 25, y = 425, height = 25)

        self.worker_job_dropdown = Combobox(self.root, values=self.worker_job_options)
        self.worker_job_dropdown.place(x = 108, y = 425, width = 180, height = 25)


#date range picker
        self.worker_job_label = Label(self.root, text= 'Date Range: ', font = dropdown_font, state = 'normal')
        self.worker_job_label.place(x = 25, y = 460, height = 25)

        self.start_date = DateEntry(self.root, width= 7, foreground= "white",bd=2)
        self.start_date.delete(0, 'end')
        self.start_date.place(x = 105, y = 460, height = 25)

        self.to_label = Label(self.root, text= 'to', font = dropdown_font, state = 'normal')
        self.to_label.place(x = 178, y = 460, height = 25)

        self.end_date = DateEntry(self.root, width= 7, foreground= "white",bd=2)
        self.end_date.delete(0, 'end')
        self.end_date.place(x = 200, y = 460, height = 25)

        self.ytd_button = Button(self.root, text='YTD', state = 'normal')
        self.ytd_button.place(x = 277, y = 460, width = 65, height = 27)


#Dataview type dropdown
        self.dataview_type_label = Label(self.root, text= 'Dataview Type: ', font = dropdown_font, state = 'normal')
        self.dataview_type_label.place(x = 25, y = 495, height = 25)

        dataview_type_options = ['Percent Share(Pie)', 'Amount Comparison(Bar)', 'Trend(Scatter)'] 

        self.dataview_type_dropdown = Combobox(self.root, values=dataview_type_options)
        self.dataview_type_dropdown.place(x = 126, y = 495, width = 180, height = 25)


#genrate query button
        self.generate_query_button = Button(self.root, text = 'Generate Query', state = 'normal')
        self.generate_query_button.place(x = 25, y = 530, width = 150, height = 45)

#delete selected queries button
        self.delete_selected_queries_button = Button(self.root, text = 'Delete Selected Queries', state = 'normal')
        self.delete_selected_queries_button.place(x = 190, y = 530, width = 150, height = 45)

#go to file loc button
        self.goto_report_loc_button = Button(self.root, text = 'Go To Report Folder', state = 'normal')
        self.goto_report_loc_button.place(x = 550, y = 280, width = 185, height = 40)


#Generate report button
        self.generate_report_button = Button(self.root, text = 'Generate Report', state = 'normal')
        self.generate_report_button.place(x = 350, y = 280, width = 185, height = 40)
        #add msgbox popup with filename

        '''
#select all queries checkbox
        self.check_all_queries_chkbox = Checkbutton(self.root, text = 'Select all Queries',variable = self.check_var, onvalue = 1, offvalue = 0)
        self.check_all_queries_chkbox.place(x = 350, y = 340, width = 150, height = 25)
        #if deleting all, do popup check
        '''

#report table
        self.report_frame = Frame(self.root)
        self.report_frame.place(x = 350, y = 365, width = 435, height = 205)

            #table scrollbars
        report_vertical_scroll = Scrollbar(self.report_frame)
        report_vertical_scroll.pack(side=RIGHT, fill=Y)

        report_horizontal_scroll = Scrollbar(self.report_frame,orient='horizontal')
        report_horizontal_scroll.pack(side= BOTTOM,fill=X)

        self.report_frame = ttk.Treeview(self.report_frame,yscrollcommand=report_vertical_scroll.set, xscrollcommand =report_horizontal_scroll.set)

        report_vertical_scroll.config(command=self.report_frame.yview)
        report_horizontal_scroll.config(command=self.report_frame.xview)
        
            #setting up table
        self.report_frame['columns'] = ('query', 'query_type')
        self.report_frame.column('#0', width=0, stretch=NO)
        self.report_frame.column('query', anchor=CENTER, width = 285)
        self.report_frame.column('query_type', anchor=CENTER, width = 175)

        self.report_frame.heading('#0', text = '', anchor=CENTER)
        self.report_frame.heading('query', text = 'Query', anchor=CENTER)
        self.report_frame.heading('query_type', text = 'Query Type', anchor=CENTER)

        self.report_frame.pack()

        count = 1
        self.report_frame.insert(parent = '', index = 'end', iid = count, text = '', values = ('Select Driver, Dee, Time Share, YTD', 'Percent Time Share(Pie)'))
        
        '''
#progress bar
        self.progress_bar = Progressbar(self.root, maximum=100)
        self.progress_bar.place(x = 0, y = 588, height = 12, width = 540)
        #self.progress_bar['value'] = 0


#progress bar label
        #self.progress_bar_label = Label(self.root, text= 'Progress label placeholder', font = widget_font, state = 'normal')
        self.progress_bar_label = Label(self.root, text= '', font = widget_font, state = 'normal')
        self.progress_bar_label.place(x = 555, y = 580, height = 20)

        '''


    def import_spreadsheet_btn_click(self):
        filename_list = filedialog.askopenfilenames(initialdir = "", title = "Select a File", filetypes = (("all files", "*.*"),))

        
        #just for now, will multiprocess this
        for filename in filename_list:
            excel_sheet = TWExportExcelFile(filename)


            filename_spl = filename.split('.')
            new_filename = filename_spl[0] +'.csv'

            file_lst = db_get_excel_file(new_filename)


            filepath_parts = file_lst[0][0].split('/')
            formatted_correct = filepath_parts[3:-1]
            filepath_str = ''

            for part in formatted_correct:
                if part == formatted_correct[0]:
                    filepath_str = filepath_str + part
                else:
                    filepath_str = filepath_str + '/' + part

            formatted_filename = filepath_parts[-1].split('.')[0]


            self.excel_sheet_frame.insert(parent = '', index = 'end', iid = self.excel_index, text = '', values = (formatted_filename, filepath_str, file_lst[0][2]))
            self.excel_index += 1

            '''
            self.worker_job_options = [] #get from database
            self.worker_team_options = [] #get from database
            self.job_type_options = [] #get from database
            self. individual_worker_options = [] #get from database
            '''
            self.worker_team_options = self.worker_team_options + list(set(db_get_data_from_column('TeamNames','team')))
            self.worker_job_options = self.worker_job_options + list(set(db_get_data_from_column('WorkerJob','individual')))
            self.job_type_options =  self.job_type_options + list(set(db_get_data_from_column('JobType','team')))
            self.individual_worker_options = self.individual_worker_options + list(set(db_get_data_from_column('WorkerName','individual')))

            self.job_type_dropdown['values'] = self.job_type_options
            self.worker_team_dropdown['values'] = self.worker_team_options
            self.individual_worker_dropdown['values'] = self.individual_worker_options
            self.worker_job_dropdown['values'] = self.worker_job_options


    def remove_spreadsheet_btn_click(self):
        try:
            selected_item_index = self.excel_sheet_frame.focus()
            item = self.excel_sheet_frame.item(selected_item_index)
            if selected_item_index != '':

                file_name = str(item['values'][0]) + '.csv' 
                filepath_list = db_get_all_excel_filepaths()
                filepath = ''

                for data in filepath_list:
                    if data[0].find(file_name) != -1:
                        filepath = data[0]

                db_remove_excel_file(filepath)
                
                self.excel_sheet_frame.delete(selected_item_index)

        except:
            pass

        
        
        
    '''
    def update_progress_bar(self):
        if self.progress_bar['value'] == 100:
            self.progress_bar['value'] == 0
        else:
            self.progress_bar['value'] += 25
    '''



