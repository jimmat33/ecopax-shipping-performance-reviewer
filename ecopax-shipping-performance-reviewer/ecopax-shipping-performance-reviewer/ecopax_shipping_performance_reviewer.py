from TWExportExcelFile import *

def main():
    test_sheet = TWExportExcelFile(r'C:\Users\jmattison\Desktop\performance-reviewer\data-1655127929240.xls')
    test_sheet.process_data()

if __name__ == '__main__':
    main()
