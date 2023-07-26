"""
docstr
"""
import timeit
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinter.ttk import Scrollbar, Button, Frame
from create_report import create_report
from excel_file import TWExportExcelFile
from report_card import create_report_cards


def main():
    root = tk.Tk()
    root.geometry("800x400")
    root.title("Shipping Performance Reviewer")
    root.resizable(False, False)

    img = tk.PhotoImage(file=(os.path.abspath("gui_icon.png")))
    root.tk.call("wm", "iconphoto", root._w, img)

    excel_index = 1


    goto_report_loc_button = Button(
        root,
        text="Go To Excel Report Folder",
        state="normal",
        command=open_file_loc,
    )

    generate_report_button = Button(
        root, text="Generate Reports", state="normal", command=generate_report
    )

    report_card_button = Button(
        root,
        text="Go To Report Card Folder",
        state="normal",
        command=go_to_report_cards_btn_click,
    )

    excel_sheet_frame = Frame(root)

    excel_sheet_frame.place(x=300, y=35, width=485, height=240)
    goto_report_loc_button.place(x=25, y=245, width=220, height=55)
    generate_report_button.place(x=25, y=175, width=220, height=55)
    report_card_button.place(x=25, y=315, width=220, height=55)

    excel_vertical_scroll = Scrollbar(excel_sheet_frame)
    excel_vertical_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    excel_horizontal_scroll = Scrollbar(excel_sheet_frame, orient="horizontal")
    excel_horizontal_scroll.pack(side=tk.BOTTOM, fill=tk.X)

    excel_sheet_frame = ttk.Treeview(
        excel_sheet_frame,
        yscrollcommand=excel_vertical_scroll.set,
        xscrollcommand=excel_horizontal_scroll.set,
    )

    excel_vertical_scroll.config(command=excel_sheet_frame.yview)
    excel_horizontal_scroll.config(command=excel_sheet_frame.xview)

    excel_sheet_frame["columns"] = ("file_name", "file_location", "date_range")
    excel_sheet_frame.column("#0", width=0, stretch=tk.NO)
    excel_sheet_frame.column("file_name", anchor=tk.CENTER, width=130)
    excel_sheet_frame.column("file_location", anchor=tk.CENTER, width=205)
    excel_sheet_frame.column("date_range", anchor=tk.CENTER, width=130)

    excel_sheet_frame.heading("#0", text="", anchor=tk.CENTER)
    excel_sheet_frame.heading("file_name", text="File Name", anchor=tk.CENTER)
    excel_sheet_frame.heading(
        "file_location", text="File Location", anchor=tk.CENTER
    )
    excel_sheet_frame.heading("date_range", text="Date Range", anchor=tk.CENTER)

    excel_sheet_frame.pack()

    run_gui()


def run_gui():
    """
    docstr
    """
    init_widgits()
    root.after(1, on_open)
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


def open_file_loc():
    """
    docstr
    """
    filebrowser_path = os.path.join(os.getenv("WINDIR"), "explorer.exe")

    if os.path.exists(
        os.path.abspath(
            "ecopax-shipping-performance-reviewer\\Performance Review Reports"
        )
    ):
        path = os.path.abspath(
            "ecopax-shipping-performance-reviewer\\Performance Review Reports"
        )
    else:
        path = os.path.abspath("Performance Review Reports")
    try:
        if os.path.isdir(path):
            subprocess.run([filebrowser_path, path], check=True)
        elif os.path.isfile(path):
            subprocess.run(
                [filebrowser_path, "/select,", os.path.normpath(path)], check=True
            )
    except Exception:
        pass


def import_spreadsheet_btn_click():
    """
    docstr
    """
    try:
        # pylint: disable=E1102
        filename_list = filedialog.askopenfilenames(
            initialdir="", title="Select a File", filetypes=(("all files", "*.*"),)
        )
        start = timeit.default_timer()
        root.withdraw()
        for filename in filename_list:
            TWExportExcelFile(filename)

            filename_spl = filename.split(".")
            new_filename = filename_spl[0] + ".csv"

            file_lst = db_get_excel_file(new_filename)

            filepath_parts = file_lst[0][0].split("/")
            formatted_correct = filepath_parts[3:-1]
            filepath_str = ""

            for part in formatted_correct:
                if part == formatted_correct[0]:
                    filepath_str = filepath_str + part
                else:
                    filepath_str = filepath_str + "/" + part

            formatted_filename = filepath_parts[-1].split(".")[0]

            excel_sheet_frame.insert(
                parent="",
                index="end",
                iid=excel_index,
                text="",
                values=(formatted_filename, filepath_str, file_lst[0][2]),
            )
            excel_index += 1

            stop = timeit.default_timer()
            print(f"\n\nDone, Time Ran: {(stop - start)/60} minutes")

        root.deiconify()
    except Exception:
        messagebox.showerror(
            "Import Error",
            "There was an error importing your file. Please try another file, or contact James Mattison",
        )


def remove_spreadsheet_btn_click():
    """
    docstr
    """
    # pylint: disable=W0703
    # pylint: disable=E1111
    try:
        selected_item_index = excel_sheet_frame.focus()
        item = excel_sheet_frame.item(selected_item_index)
        if selected_item_index != "":
            file_name = str(item["values"][0]) + ".csv"
            filepath_list = db_get_all_excel_filepaths()
            filepath = ""

            for data in filepath_list:
                if data[0].find(file_name) != -1:
                    filepath = data[0]

            db_remove_excel_file(filepath)

            excel_sheet_frame.delete(selected_item_index)

    except Exception:
        messagebox.showerror(
            "Removal Error",
            "There was an error removing your file. Please try another file, or contact James Mattison",
        )


def go_to_report_cards_btn_click():
    """
    docstr
    """
    filebrowser_path = os.path.join(os.getenv("WINDIR"), "explorer.exe")

    if os.path.exists(
        os.path.abspath("ecopax-shipping-performance-reviewer\\Report Card Reports")
    ):
        path = os.path.abspath(
            "ecopax-shipping-performance-reviewer\\Report Card Reports"
        )
    else:
        path = os.path.abspath("Report Card Reports")
    try:
        if os.path.isdir(path):
            subprocess.run([filebrowser_path, path], check=True)
        elif os.path.isfile(path):
            subprocess.run(
                [filebrowser_path, "/select,", os.path.normpath(path)], check=True
            )
    except Exception:
        pass


def generate_report():
    """
    docstr
    """
    try:
        root.withdraw()

        create_report()
        create_report_cards()

        root.deiconify()
    except Exception:
        messagebox.showerror(
            "Report Error",
            "There was an error creating your report files. Please try another Transwide file, or contact James Mattison",
        )


def on_closing():
    """
    docstr
    """
    db_clear_database()
    root.destroy()


def on_open():
    """
    docstr
    """
    db_clear_database()

    if os.path.exists(
        os.path.abspath("ecopax-shipping-performance-reviewer\\Report Card Cache")
    ):
        f_path = os.path.abspath(
            "ecopax-shipping-performance-reviewer\\Report Card Cache"
        )
    else:
        f_path = os.path.abspath("Report Card Cache")

    for f in os.listdir(f_path):
        os.remove(os.path.join(f_path, f))
