'''
docstr
'''
import os
import string
from datetime import date, datetime
from pathlib import Path
import win32com.client as win32
import xlsxwriter
from performance_review_db import db_get_individual_data, db_get_team_data
win32c = win32.constants
win32.gencache.is_readonly = False
win32.gencache.GetGeneratePath()
# pylint: disable=W0703
# pylint: disable=R0914


def create_report():
    '''
    docstr
    '''
    create_excel_file()


def create_excel_file():
    '''
    docstr
    '''
    folder_fp = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
    if os.path.exists(os.path.abspath('Ecopax-Performance-Reviwer-Program-Files\\Performance Review Reports')):
        folder_fp = os.path.abspath('Ecopax-Performance-Reviwer-Program-Files\\Performance Review Reports')
    else:
        folder_fp = os.path.abspath('Performance Review Reports')

    f_name = '\\PerformanceReport-' + datetime.now().strftime('%m-%d-%y-%H-%M-%S') + '.xlsx'

    created_fp = folder_fp + f_name

    ind_data = db_get_individual_data()
    team_data = db_get_team_data()

    if len(ind_data) == 0 or len(team_data) == 0:
        return

    workbook = xlsxwriter.Workbook(created_fp)
    ind_sheet = workbook.add_worksheet('Individual Data Master')
    team_sheet = workbook.add_worksheet('Team Data Master')

    populate_worksheet(ind_data, ind_sheet, workbook)
    populate_worksheet(team_data, team_sheet, workbook)

    workbook.close()

    run_excel(created_fp)


def populate_worksheet(data_lst, sheet_obj, wrk_book):
    '''
    docstr
    '''
    date_format = {'num_format': 'mm/dd/yyyy'}
    fmt = wrk_book.add_format(date_format)

    if sheet_obj.name == 'Individual Data Master':
        sheet_obj.write('A1', 'Worker Name')
        sheet_obj.write('B1', 'Job Date')
        sheet_obj.write('C1', 'Job Type')
        sheet_obj.write('D1', 'Total Job Time')
        sheet_obj.write('E1', 'Worker Job')
        sheet_obj.write('F1', 'Minutes Per Day')
        sheet_obj.write('G1', 'Month')

    elif sheet_obj.name == 'Team Data Master':
        sheet_obj.write('A1', 'Team Names')
        sheet_obj.write('B1', 'Job Date')
        sheet_obj.write('C1', 'Job Type')
        sheet_obj.write('D1', 'Total Job Time')
        sheet_obj.write('E1', 'Month')

    i = 0

    while i < len(data_lst):
        j = 0
        while j < len(data_lst[i]):
            cell_index = f'{string.ascii_uppercase[j]}{i + 2}'
            if isinstance(data_lst[i][j], date):
                sheet_obj.write(cell_index, data_lst[i][j], fmt)
            elif str(data_lst[i][j]).isnumeric():
                sheet_obj.write(cell_index, int(data_lst[i][j]))
            else:
                sheet_obj.write(cell_index, data_lst[i][j])
            j += 1
        i += 1

    if sheet_obj.name == 'Individual Data Master':
        sheet_obj.autofilter(f'A1:G{i + 1}')
    elif sheet_obj.name == 'Team Data Master':
        sheet_obj.autofilter(f'A1:E{i + 1}')


def run_excel(str_path):
    '''
    docstr
    '''
    f_path = Path(str_path)
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    try:
        wrk_book = excel.Workbooks.Open(f_path)
    except Exception:
        print(f'Failed to open spreadsheet.  Invalid filename or location: {f_path}')

    # set worksheet
    ind_master = wrk_book.Sheets('Individual Data Master')
    team_master = wrk_book.Sheets('Team Data Master')

    team_mod_name = 'Team Modifiable Pivot'
    wrk_book.Sheets.Add().Name = team_mod_name
    team_mod_sheet = wrk_book.Sheets(team_mod_name)

    ind_mod_name = 'Individual Modifiable Pivot'
    wrk_book.Sheets.Add().Name = ind_mod_name
    ind_mod_sheet = wrk_book.Sheets(ind_mod_name)

    loading_v_nonloading_name = 'Loading vs. Non-Loading'
    wrk_book.Sheets.Add().Name = loading_v_nonloading_name
    loading_v_nonloading_sheet = wrk_book.Sheets(loading_v_nonloading_name)

    team_ltt_name = 'Team Loading Time Trend'
    wrk_book.Sheets.Add().Name = team_ltt_name
    team_ltt_sheet = wrk_book.Sheets(team_ltt_name)

    ind_ltt_name = 'Individual Loading Time Trend'
    wrk_book.Sheets.Add().Name = ind_ltt_name
    ind_ltt_sheet = wrk_book.Sheets(ind_ltt_name)

    ind_modifiable_pivot(wrk_book, ind_master, ind_mod_sheet)
    team_modifiable_pivot(wrk_book, team_master, team_mod_sheet)
    loading_vs_nonloading_pivot(wrk_book, ind_master, loading_v_nonloading_sheet)
    team_loading_time_trend_pivot(wrk_book, team_master, team_ltt_sheet)
    ind_loading_time_trend_pivot(wrk_book, ind_master, ind_ltt_sheet)

    wrk_book.Close(True)
    excel.Quit()


def ind_loading_time_trend_pivot(wrk_book, data_sheet, pivot_sheet):
    '''
    docstr
    '''
    table_name = 'Individual Loading Time Trend'
    table_rows = ['Worker Name', 'Month', 'Job Date']
    table_cols = []
    table_filters = ['Job Type']
    table_calcs = [['Total Job Time', 'Total Job Time Sum', win32c.xlSum, '##0'],
                   ['Total Job Time', 'Total Job Time Avg', win32c.xlAverage, '##0.00'],
                   ['Total Job Time', 'Number of Jobs Done', win32c.xlCount, '##0']]

    # pivot table location
    pivot_cell_loc = pivot_sheet.Cells(3, 1)
    pivot_target_range = pivot_sheet.Range(pivot_cell_loc, pivot_cell_loc)

    # grab the pivot table source data
    pivot_cache = wrk_book.PivotCaches().Create(SourceType=win32c.xlDatabase,
                                                SourceData=data_sheet.UsedRange)

    # create the pivot table object
    pivot_cache.CreatePivotTable(TableDestination=pivot_target_range, TableName=table_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pivot_sheet.Select()
    pivot_sheet.Cells(3, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in ((table_filters, win32c.xlPageField),
                                (table_rows, win32c.xlRowField),
                                (table_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pivot_sheet.PivotTables(table_name).PivotFields(value).Orientation = field_r
            pivot_sheet.PivotTables(table_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in table_calcs:
        pivot_sheet.PivotTables(table_name).AddDataField(
            pivot_sheet.PivotTables(table_name)
            .PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pivot_sheet.PivotTables(table_name).ShowValuesRow = True
    pivot_sheet.PivotTables(table_name).ColumnGrand = True

    pivot_sheet.Shapes.AddChart2(201)


def team_loading_time_trend_pivot(wrk_book, data_sheet, pivot_sheet):
    '''
    docstr
    '''
    table_name = 'Team Loading Time Trend'
    table_rows = ['Team Names', 'Month', 'Job Date']
    table_cols = []
    table_filters = ['Job Type']
    table_calcs = [['Total Job Time', 'Total Job Time Sum', win32c.xlSum, '##0'],
                   ['Total Job Time', 'Total Job Time Avg', win32c.xlAverage, '##0.00'],
                   ['Total Job Time', 'Number of Jobs Done', win32c.xlCount, '##0']]

    # pivot table location
    pivot_cell_loc = pivot_sheet.Cells(3, 1)
    pivot_target_range = pivot_sheet.Range(pivot_cell_loc, pivot_cell_loc)

    # grab the pivot table source data
    pivot_cache = wrk_book.PivotCaches().Create(SourceType=win32c.xlDatabase,
                                                SourceData=data_sheet.UsedRange)

    # create the pivot table object
    pivot_cache.CreatePivotTable(TableDestination=pivot_target_range, TableName=table_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pivot_sheet.Select()
    pivot_sheet.Cells(3, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in ((table_filters, win32c.xlPageField),
                                (table_rows, win32c.xlRowField),
                                (table_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pivot_sheet.PivotTables(table_name).PivotFields(value).Orientation = field_r
            pivot_sheet.PivotTables(table_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in table_calcs:
        pivot_sheet.PivotTables(table_name).AddDataField(
            pivot_sheet.PivotTables(table_name)
            .PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pivot_sheet.PivotTables(table_name).ShowValuesRow = True
    pivot_sheet.PivotTables(table_name).ColumnGrand = True

    pivot_sheet.Shapes.AddChart2(201)


def loading_vs_nonloading_pivot(wrk_book, data_sheet, pivot_sheet):
    '''
    doscstr
    '''
    table_name = 'Individual Loading Vs. Non-Loading Time'
    table_rows = ['Worker Name', 'Month', 'Job Date']
    table_cols = []
    table_filters = ['Job Type', 'Worker Job']
    table_calcs = [['Total Job Time', 'Job Time Sum', win32c.xlSum, '##0'],
                   ['Minutes Per Day', 'Num Minutes Per Day', win32c.xlMin, '##0']]

    # pivot table location
    pivot_cell_loc = pivot_sheet.Cells(3, 1)
    pivot_target_range = pivot_sheet.Range(pivot_cell_loc, pivot_cell_loc)

    # grab the pivot table source data
    pivot_cache = wrk_book.PivotCaches().Create(SourceType=win32c.xlDatabase,
                                                SourceData=data_sheet.UsedRange)

    # create the pivot table object
    pivot_cache.CreatePivotTable(TableDestination=pivot_target_range, TableName=table_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pivot_sheet.Select()
    pivot_sheet.Cells(3, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in ((table_filters, win32c.xlPageField),
                                (table_rows, win32c.xlRowField),
                                (table_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pivot_sheet.PivotTables(table_name).PivotFields(value).Orientation = field_r
            pivot_sheet.PivotTables(table_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in table_calcs:
        pivot_sheet.PivotTables(table_name).AddDataField(
            pivot_sheet.PivotTables(table_name)
            .PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pivot_sheet.PivotTables(table_name).ShowValuesRow = True
    pivot_sheet.PivotTables(table_name).ColumnGrand = True

    pivot_sheet.Shapes.AddChart(53)


def ind_modifiable_pivot(wrk_book, data_sheet, pivot_sheet):
    '''
    docstr
    '''
    table_name = 'Individual Modifiable Table'
    table_rows = ['Worker Name', 'Month', 'Job Date']
    table_cols = []
    table_filters = []
    table_calcs = []

    # pivot table location
    pivot_cell_loc = pivot_sheet.Cells(3, 1)
    pivot_target_range = pivot_sheet.Range(pivot_cell_loc, pivot_cell_loc)

    # grab the pivot table source data
    pivot_cache = wrk_book.PivotCaches().Create(SourceType=win32c.xlDatabase,
                                                SourceData=data_sheet.UsedRange)

    # create the pivot table object
    pivot_cache.CreatePivotTable(TableDestination=pivot_target_range, TableName=table_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pivot_sheet.Select()
    pivot_sheet.Cells(3, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in ((table_filters, win32c.xlPageField),
                                (table_rows, win32c.xlRowField),
                                (table_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pivot_sheet.PivotTables(table_name).PivotFields(value).Orientation = field_r
            pivot_sheet.PivotTables(table_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in table_calcs:
        pivot_sheet.PivotTables(table_name).AddDataField(
            pivot_sheet.PivotTables(table_name)
            .PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pivot_sheet.PivotTables(table_name).ShowValuesRow = True
    pivot_sheet.PivotTables(table_name).ColumnGrand = True

    pivot_sheet.Shapes.AddChart2(201)


def team_modifiable_pivot(wrk_book, data_sheet, pivot_sheet):
    '''
    docstr
    '''
    table_name = 'Team Modifiable Table'
    table_rows = ['Team Names', 'Month', 'Job Date']
    table_cols = []
    table_filters = []
    table_calcs = []

    # pivot table location
    pivot_cell_loc = pivot_sheet.Cells(3, 1)
    pivot_target_range = pivot_sheet.Range(pivot_cell_loc, pivot_cell_loc)

    # grab the pivot table source data
    pivot_cache = wrk_book.PivotCaches().Create(SourceType=win32c.xlDatabase,
                                                SourceData=data_sheet.UsedRange)

    # create the pivot table object
    pivot_cache.CreatePivotTable(TableDestination=pivot_target_range, TableName=table_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pivot_sheet.Select()
    pivot_sheet.Cells(3, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in ((table_filters, win32c.xlPageField),
                                (table_rows, win32c.xlRowField),
                                (table_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pivot_sheet.PivotTables(table_name).PivotFields(value).Orientation = field_r
            pivot_sheet.PivotTables(table_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in table_calcs:
        pivot_sheet.PivotTables(table_name).AddDataField(
            pivot_sheet.PivotTables(table_name)
            .PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pivot_sheet.PivotTables(table_name).ShowValuesRow = True
    pivot_sheet.PivotTables(table_name).ColumnGrand = True

    pivot_sheet.Shapes.AddChart2(201)
