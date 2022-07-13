'''
docstr
'''
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.ttk import Scrollbar, Button, Frame
from create_report import create_report
from performance_review_db import (db_get_excel_file, db_get_all_excel_filepaths,
                                   db_remove_excel_file, db_clear_database)
import excel_file


class PerformanceReviewGUI():
    '''
    docstr
    '''
    # pylint: disable=R0902
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry('800x400')
        self.root.title("Shipping Performance Reviewer")
        self.root.resizable(False, False)
        img = tk.PhotoImage(file=(os.path.abspath('gui_icon.png')))
        self.root.tk.call('wm', 'iconphoto', self.root._w, img)
        self.excel_index = 1

        self.import_sheet_button = Button(self.root, text='Import Spreadsheet', state='normal',
                                          command=self.import_spreadsheet_btn_click)

        self.remove_sheet_button = Button(self.root, text='Remove Spreadsheet', state='normal',
                                          command=self.remove_spreadsheet_btn_click)

        self.goto_report_loc_button = Button(self.root, text='Go To Report Folder', state='normal',
                                             command=self.open_file_loc)

        self.generate_report_button = Button(self.root, text='Generate Report', state='normal',
                                             command=self.generate_report)

        self.excel_sheet_frame = Frame(self.root)

    def run_gui(self):
        '''
        docstr
        '''
        self.init_widgits()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

    def init_widgits(self):
        '''
        docstr
        '''
        self.import_sheet_button.place(x=25, y=35, width=220, height=55)
        self.remove_sheet_button.place(x=25, y=105, width=220, height=55)
        self.excel_sheet_frame.place(x=300, y=35, width=485, height=240)
        self.goto_report_loc_button.place(x=585, y=280, width=185, height=40)
        self.generate_report_button.place(x=385, y=280, width=185, height=40)

        excel_vertical_scroll = Scrollbar(self.excel_sheet_frame)
        excel_vertical_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        excel_horizontal_scroll = Scrollbar(self.excel_sheet_frame, orient='horizontal')
        excel_horizontal_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.excel_sheet_frame = ttk.Treeview(self.excel_sheet_frame,
                                              yscrollcommand=excel_vertical_scroll.set,
                                              xscrollcommand=excel_horizontal_scroll.set)

        excel_vertical_scroll.config(command=self.excel_sheet_frame.yview)
        excel_horizontal_scroll.config(command=self.excel_sheet_frame.xview)

        self.excel_sheet_frame['columns'] = ('file_name', 'file_location', 'date_range')
        self.excel_sheet_frame.column('#0', width=0, stretch=tk.NO)
        self.excel_sheet_frame.column('file_name', anchor=tk.CENTER, width=130)
        self.excel_sheet_frame.column('file_location', anchor=tk.CENTER, width=205)
        self.excel_sheet_frame.column('date_range', anchor=tk.CENTER, width=130)

        self.excel_sheet_frame.heading('#0', text='', anchor=tk.CENTER)
        self.excel_sheet_frame.heading('file_name', text='File Name', anchor=tk.CENTER)
        self.excel_sheet_frame.heading('file_location', text='File Location', anchor=tk.CENTER)
        self.excel_sheet_frame.heading('date_range', text='Date Range', anchor=tk.CENTER)

        self.excel_sheet_frame.pack()

    def open_file_loc(self):
        '''
        docstr
        '''
        filebrowser_path = os.path.join(os.getenv('WINDIR'), 'explorer.exe')
        path = os.path.abspath('Performance Review Reports')

        if os.path.isdir(path):
            subprocess.run([filebrowser_path, path], check=True)
        elif os.path.isfile(path):
            subprocess.run([filebrowser_path, '/select,', os.path.normpath(path)], check=True)

    def import_spreadsheet_btn_click(self):
        '''
        docstr
        '''
        # pylint: disable=E1102
        filename_list = filedialog.askopenfilenames(initialdir="", title="Select a File",
                                                    filetypes=(("all files", "*.*"),))

        for filename in filename_list:
            excel_file(filename)

            filename_spl = filename.split('.')
            new_filename = filename_spl[0] + '.csv'

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

            self.excel_sheet_frame.insert(parent='', index='end', iid=self.excel_index, text='',
                                          values=(formatted_filename, filepath_str, file_lst[0][2]))
            self.excel_index += 1

    def remove_spreadsheet_btn_click(self):
        '''
        docstr
        '''
        # pylint: disable=W0703
        # pylint: disable=E1111
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

        except Exception:
            pass

    def generate_report(self):
        '''
        docstr
        '''
        create_report()

    def on_closing(self):
        '''
        docstr
        '''
        db_clear_database()
        self.root.destroy()
