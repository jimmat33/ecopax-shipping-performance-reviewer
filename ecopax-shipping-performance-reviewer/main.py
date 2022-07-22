'''
This file is used to create a gui instance and freeze multiprocessing
when the file is compiled in exe format

Use: pylint <filename> for static code quality checks
Use: radon <cc, mi, raw, hal> <filename>.py for complexity checking
'''
import multiprocessing as mp
from gui_application import PerformanceReviewGUI

if __name__ == '__main__':
    mp.freeze_support()
    manager = mp.Manager()
    gui_obj = PerformanceReviewGUI(manager)
    gui_obj.run_gui()
