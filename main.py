import PySimpleGUI as sg

BACKGROUND_GREY_HEX = '#B3B3B3'
ECOPAX_GREEN_HEX = '#8DC63F'
WHITE_HEX = '#FFFFFF'
BLACK_HEX = '#000000'

TITLE_FONT = ('Arial', 18)


def gui_command_loop():
    """
    """
    layout = build_gui()
    sg.theme_background_color(BACKGROUND_GREY_HEX)

    window = sg.Window('Ecopax Shipping Performance Analysis Tool', layout)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()

        if event == '-import_sheet_btn-':
            # Get file name from file browser
            # verify that it is a valid file
            # send it to the function to handle imports
            # pull values from drop downs regarding column location
            # set status text
            pass

        if event == '-generate_report_btn-':
            # check that a file has been imported
            # generate report with that file
            pass

        if event == '-go_to_report_btn-':
            # Open file location for the reports, based on current user settings
            pass

        if event == '-default_column_loc_btn-':
            # open text file that stores the default column locations
            pass

        if event == '-report_output_loc_btn-':
            # set the output directory for the report
            pass

        if event == '-open_logs_btn-':
            # open the folder containing the logs, store this in a specific spot on the machine so it doesn't get deleted
            pass

        # If user closes window or clicks cancel
        if event == sg.WIN_CLOSED or event == 'Cancel':
            break


def build_gui() -> list:
    """
    """
    # ===========================================================================================
    #   The Main Tab (tab1)
    # ===========================================================================================

    title_txt =                             sg.Text('Ecopax Shipping Performance Analysis Tool',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX,
                                                    font=TITLE_FONT,
                                                    pad=(5, (5, 20)))

    file_txt_input =                        sg.Input(readonly=True,
                                                     size=(70, 10),
                                                     pad=(5, 0),
                                                     key='-file_txt_input-')

    tw_file_browser =                       sg.FileBrowse(button_color=ECOPAX_GREEN_HEX,
                                                          key='-tw_file_browser-')

    import_sheet_btn =                      sg.Button('Import Transwide Sheet',
                                                      button_color=ECOPAX_GREEN_HEX,
                                                      pad=(5, (5, 20)),
                                                      size=(20, 2),
                                                      key='-import_sheet_btn-')

    file_input_status_txt =                 sg.Text('',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX,
                                                    font=TITLE_FONT,
                                                    pad=(5, (5, 20)),
                                                    key='-file_input_status_txt-')

    generate_report_btn =                   sg.Button('Generate Report',
                                                      button_color=ECOPAX_GREEN_HEX,
                                                      pad=(5, (20, 20)),
                                                      size=(20, 2),
                                                      key='-generate_report_btn-')

    go_to_report_btn =                      sg.Button('Go To Report Folder',
                                                      button_color=ECOPAX_GREEN_HEX,
                                                      pad=(20, (20, 20)),
                                                      size=(20, 2),
                                                      key='-go_to_report_btn-')

    # ===========================================================================================
    #   Column Selection Tab (tab2)
    # ===========================================================================================

    team_names_combo_txt =                  sg.Text('Team Name Column: ',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX)

    team_names_combo =                      sg.Combo(values=[],
                                                     readonly=True,
                                                     button_background_color=ECOPAX_GREEN_HEX,
                                                     key='-team_names_combo-')

    booking_date_combo_txt =                sg.Text('Booking Date Column: ',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX)

    booking_date_combo =                    sg.Combo(values=[],
                                                     readonly=True,
                                                     button_background_color=ECOPAX_GREEN_HEX,
                                                     key='-booking_date_combo-')

    job_type_combo_txt =                    sg.Text('Job Type Column: ',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX)

    job_type_combo =                        sg.Combo(values=[],
                                                     readonly=True,
                                                     button_background_color=ECOPAX_GREEN_HEX,
                                                     key='-job_date_combo-')

    transport_reference_combo_txt =         sg.Text('Transport Reference Column: ',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX)

    transport_reference_combo =             sg.Combo(values=[],
                                                     readonly=True,
                                                     button_background_color=ECOPAX_GREEN_HEX,
                                                     key='-transport_reference_combo-')

    job_start_time_combo_txt =              sg.Text('Job Start Time Column: ',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX)

    job_start_time_combo =                  sg.Combo(values=[],
                                                     readonly=True,
                                                     button_background_color=ECOPAX_GREEN_HEX,
                                                     key='-job_start_time_combo-')

    job_end_time_combo_txt =                sg.Text('Job End Time Column: ',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX)

    job_end_time_combo =                    sg.Combo(values=[],
                                                     readonly=True,
                                                     button_background_color=ECOPAX_GREEN_HEX,
                                                     key='-job_end_time_combo-')

    in_or_out_combo_txt =                   sg.Text('Ingoing or Outgoing Column: ',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX)

    in_or_out_combo =                       sg.Combo(values=[],
                                                     readonly=True,
                                                     button_background_color=ECOPAX_GREEN_HEX,
                                                     key='-in_or_out_combo-')

    # ===========================================================================================
    #   User Settings Tab (tab3)
    # ===========================================================================================

    default_column_loc_txt =                sg.Text('Modify Default Column Locations: ',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX)

    default_column_loc_btn =                sg.Button('Open Default Column File',
                                                      button_color=ECOPAX_GREEN_HEX,
                                                      key='-default_column_loc_btn-')

    report_output_loc_txt =                 sg.Text('Report Output Location: ',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX,
                                                    pad=(5, (35, 20)))

    report_out_txt_input =                  sg.Input(readonly=True,
                                                     size=(50, 10),
                                                     pad=(5, (35, 20)),
                                                     key='-report_out_txt_input-')

    report_output_file_browser =            sg.FileBrowse(button_color=ECOPAX_GREEN_HEX,
                                                          pad=(5, (35, 20)),
                                                          key='-report_output_file_browser-')

    report_output_loc_btn =                 sg.Button('Change Report Output Location',
                                                      button_color=ECOPAX_GREEN_HEX,
                                                      pad=(5, (5, 20)),
                                                      key='-report_output_loc_btn-')

    report_output_chng_status_txt =         sg.Text('',
                                                    background_color=BACKGROUND_GREY_HEX,
                                                    text_color=BLACK_HEX,
                                                    pad=(5, (5, 20)),
                                                    key='-report_output_chng_status_txt-')

    open_logs_btn =                         sg.Button('Open Error Logs',
                                                      button_color=ECOPAX_GREEN_HEX,
                                                      size=(20, 2),
                                                      pad=(5, (5, 20)),
                                                      key='-open_logs_btn-')

    # ===========================================================================================
    #   Tab and Window Layout
    # ===========================================================================================

    tab1_layout = [[title_txt],
                   [file_txt_input, tw_file_browser],
                   [import_sheet_btn, file_input_status_txt],
                   [generate_report_btn, go_to_report_btn]]

    tab2_layout = [[team_names_combo_txt, team_names_combo],
                   [booking_date_combo_txt, booking_date_combo],
                   [job_type_combo_txt, job_type_combo],
                   [transport_reference_combo_txt, transport_reference_combo],
                   [job_start_time_combo_txt, job_start_time_combo],
                   [job_end_time_combo_txt, job_end_time_combo],
                   [in_or_out_combo_txt, in_or_out_combo]]

    tab3_layout = [[default_column_loc_txt, default_column_loc_btn],
                   [report_output_loc_txt, report_out_txt_input, report_output_file_browser],
                   [report_output_loc_btn],
                   [report_output_chng_status_txt],
                   [open_logs_btn]]

    tab_group = [sg.TabGroup(background_color=BACKGROUND_GREY_HEX,
                             selected_title_color=BLACK_HEX,
                             selected_background_color=ECOPAX_GREEN_HEX,
                             title_color=BLACK_HEX,
                             layout=[[sg.Tab(title="Run Program", background_color=BACKGROUND_GREY_HEX, layout=tab1_layout),
                                     sg.Tab(title="Select Columns", background_color=BACKGROUND_GREY_HEX, layout=tab2_layout),
                                     sg.Tab(title="User Settings", background_color=BACKGROUND_GREY_HEX, layout=tab3_layout)]])]

    return [tab_group]


if __name__ == '__main__':
    gui_command_loop()
