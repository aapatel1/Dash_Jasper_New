import PySimpleGUI as sg
from datetime import timedelta, date
import threading
from multiprocessing import Pool
from openpyxl import load_workbook
import os
import queue
from main_dash import run_download_report
from summary_report import run_combined_summary
from split_summary import run_split_summary
from concurrent.futures import ThreadPoolExecutor


sg.theme('DarkBlue3')
gui_queue = queue.Queue()


def download_dash(report_date, output_path, client, reports):
    dash = run_download_report()
    dash_status = dash.run_download(gui_queue, report_date, output_path, client, reports)
    return dash_status


def dash_summary(report_date, output_path, client, summary_file, summary_type):
    if summary_type == "Combined":
        dash = run_combined_summary()
        dash.gui_queue = gui_queue
        dash_status = dash.run(report_date, output_path, client, summary_file)
        return dash_status
    else:
        dash = run_split_summary()
        dash.gui_queue = gui_queue
        dash_status = dash.run(report_date, output_path, client, summary_file)
        return dash_status


def load_setting():
    setting = 'Dash_SettingSheet.xlsx'
    settingWB = load_workbook(setting, data_only=True, read_only=True)
    creds_sheet = settingWB['Creds']
    creds_values = [list(row) for row in creds_sheet.values]
    clients = [str(item[0]).strip() for item in creds_values[1:]]

    report_sheet = settingWB['Dash_Reports']
    report_data = [list(row) for row in report_sheet.values if row[0]]
    client_report_dict = {}
    for row in report_data[1:]:
        client_name = str(row[0]).strip()
        report_name = str(row[1]).strip()
        if not client_report_dict or client_name not in client_report_dict.keys():
            client_report_dict[client_name] = [report_name]
        else:
            client_report_dict[client_name].append(report_name)

    return clients, client_report_dict


def run_process(clients, report_date, output_path, reports):
    with ThreadPoolExecutor(max_workers=2) as executor:
        for client in clients:
            results = executor.submit(download_dash, report_date, output_path, client, reports)

        # for result in results:
        #     print(result)


def run_gui(thread=None):
    clients, client_report_dict = load_setting()
    default_client = clients[0]
    reports = client_report_dict.get(default_client) or []

    yesterday = (date.today().replace(day=1) - timedelta(days=1)).strftime('%m/%d/%Y')

    layout = [
        [
            sg.Text('Dash Jasper - Unearned Revenue Summary',
                    size=(65, 1),
                    font=('Corbel', 18),
                    justification='center',
                    pad=((0, 0), (5, 10)))
        ],
        [
            sg.CalendarButton("Report Date", size=(12, 1), format='%m/%d/%Y', key='report_date_btn',
                              enable_events=True, pad=((10, 0), (0, 0))),
            sg.Input(yesterday, size=(12, 1), font=('Corbel', 11), key='report_date', disabled=True,
                     justification='center', enable_events=True, readonly=True, pad=((10, 0), (0, 0))),
        ],
        [
            sg.Text('Select Clients: ', size=(12, 1), auto_size_text=False, justification='left',
                    pad=((10, 0), (30, 30))),
            sg.Listbox(clients, size=(30, 5), font=('Corbel', 11), key='client', disabled=False,
                       enable_events=True, select_mode='extended', pad=((10, 10), (30, 30))),
            sg.Text('Select Reports: ', size=(12, 1), auto_size_text=False, justification='left',
                    pad=((30, 0), (30, 30))),
            sg.Listbox(reports, size=(30, 5), font=('Corbel', 11), key='report_list', disabled=False,
                       enable_events=True, select_mode='extended', default_values=reports, pad=((10, 10), (30, 30))),
        ],
        [
            sg.Text('Download Path: ', size=(12, 1), auto_size_text=False, justification='left', pad=((10, 0), (0, 0))),
            sg.InputText(os.getcwd(), size=(40, 1), key='download_path', readonly=True, pad=((10, 0), (0, 0))),
            sg.FolderBrowse(initial_folder=os.getcwd(), size=(10, 1), pad=((10, 0), (0, 0)))
        ],
        [
            sg.OK('Report Download', key='report_download', size=(16, 1), font=('Corbel', 10), pad=((25, 5), (20, 20))),
            sg.OK('Combined Summary', key='summary', size=(16, 1), font=('Corbel', 10), pad=((10, 5), (20, 20))),
            sg.Exit('Exit', key='exit', size=(15, 1), font=('Corbel', 10), pad=((10, 5), (20, 20))),
        ],
        [
            sg.Text("Status :", size=(15, 1), justification='left', font=('Corbel', 11)),
        ],
        [
            sg.Multiline(size=(100, 7), font='courier 10', background_color='white', text_color='black', key='status',
                         autoscroll=True, enable_events=True, change_submits=False, pad=((10, 10), (0, 0)))
        ],
    ]

    window = sg.Window('Dash-Jasper M.28',
                       element_justification='left',
                       text_justification='left',
                       auto_size_text=True).Layout(layout).Finalize()
    window.Maximize()

    while True:
        event, values = window.Read(timeout=1000)
        window.refresh()

        if event:
            if event == "location_summary":
                window['summary'].Update("Location Summary")
                window.refresh()

            elif event == "combined_summary":
                window['summary'].Update("Combined Summary")
                window.refresh()

        if event in ('Exit', None) or event == sg.WIN_CLOSED:
            window.close()
            break

        elif event == 'report_download':
            report_date = values['report_date']
            clients = values['client']
            output_path = values['download_path']
            reports = values['report_list']
            if not reports:
                window['status'].print('Error: Please select at least one report to process.\n')
                continue

            window['status'].print('Dash Report Download Processing...\n')
            window['report_download'].Update(disabled=True)
            window['summary'].Update(disabled=True)


            # run_process(clients, report_date, output_path, reports)
            # # manager = multiprocessing.Manager()
            # # return_dict = manager.dict()
            # jobs = []
            #
            #
            #
            # # for client in clients:
            # #
            # #     p = multiprocessing.Process(target=download_dash,
            # #                                 args=(report_date, output_path, client, reports))
            # #     jobs.append(p)
            # #     p.start()
            # #
            # # for proc in jobs:
            # #     proc.join()
            # # print(return_dict.values())

            for client in clients:
                thread = threading.Thread(target=download_dash,
                                          args=(report_date, output_path, client, reports))
                thread.start()

        elif event == 'summary':
            if values['combined_summary']:
                summary_type = "Combined"
            else:
                summary_type = "Location Wise"

            report_date = values['report_date']
            client = values['client']
            output_path = values['download_path']
            output_path = os.path.join(output_path, "Downloads", client)
            reportDate = str(report_date).replace('/', '-')
            summary_file = os.path.join(output_path, f'Unearned Revenue - Summary {reportDate}.xlsx')

            window['status'].print('Dash Summary Processing...\n')
            window['report_download'].Update(disabled=True)
            window['summary'].Update(disabled=True)
            thread = threading.Thread(target=dash_summary,
                                      args=(report_date, output_path, client, summary_file, summary_type))
            thread.start()

        elif event == "exit":
            window.close()
            break

        if thread:
            if not thread.is_alive():
                window['report_download'].Update(disabled=False)
                window['summary'].Update(disabled=False)
                window.refresh()

        try:
            message = gui_queue.get_nowait()
        except:
            message = None
        if message:
            for key, value in message.items():
                if key == 'status':
                    window['status'].print(value)
                    window.refresh()
                if key == 'Success':
                    sg.Popup(value, title='Status')
            window.refresh()


if __name__ == '__main__':
    # main function
    run_gui()
