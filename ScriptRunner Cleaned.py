import PySimpleGUI as sg
import datetime as dt
from os import listdir, getcwd
from os.path import isfile, join, splitext, basename
import queue
import threading
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import re

MAX_NUM_SCRIPTS = 100

# Update this to use the OS.current directory or something
main_filepath = 'PUT PATH HERE'
icons_dict = {'check': {'regular': 'check.png', '24': 'check24.png'},
              'more': {'regular': 'more.png', '24': 'more24.png'},
              'more-or-less': {'regular': 'more-or-less.png', '24': 'more-or-less24.png'},
              'remove': {'regular': 'remove.png', '24': 'remove24.png'}}

def getReportingDefaults():
    right_now = dt.datetime.now()

    # if we are past the 20th of a month it is safe to assume we are ready for the next month's close
    if right_now.day >= 20:
        month = f'{right_now.month:02d}'
        year = f'{right_now.year}'
    else:
        month = f'{(right_now - dt.timedelta(days=30)).month:02d}'
        year = f'{(right_now - dt.timedelta(days=30)).year}'
    return month, year


def getBillingSystems(b_systems, values):
    selected_systems = []
    for bs in b_systems:
        if values[bs]:
            selected_systems.append(bs)
    return selected_systems


def getBSScripts(b_systems, script_path):
    scripts = []
    for file in listdir(script_path):
        if isfile(join(script_path, file)) and splitext(file)[-1].lower() == '.sql':
            for bs in b_systems:
                if(bs == "PPMMAC" and "ppm" in file.lower()):
                    scripts.append(join(script_path, file))
                elif(bs == "Sheridan_ECW" and "ecw" in file.lower() and "jupiter" not in file.lower()):
                    scripts.append(join(script_path, file))
                elif bs.lower() in file.lower():
                    scripts.append(join(script_path, file))
    return scripts


def getAllBSScripts(b_systems, script_paths):
    all_scripts = []
    for sp in script_paths:
        all_scripts.extend(getBSScripts(b_systems, sp))
    return all_scripts


def main():
    sg.theme('Reddit')

    icons = [join(getcwd(), 'assets', icons_dict['more']['24']),
             join(getcwd(), 'assets', icons_dict['more-or-less']['24']),
             join(getcwd(), 'assets', icons_dict['check']['24']),
             join(getcwd(), 'assets', icons_dict['remove']['24'])]

    script_path = f'{main_filepath}scripts\\'
    reporting_months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    reporting_years = ['2020', '2021', '2022', '2023', '2024', '2025']
    billing_systems = ['Acute',
                       'AS400',
                       'GE',
                       'Jupiter_ECW',
                       'McKesson',
                       'MediCorp',
                       'NFPHI',
                       'Northside',
                       'POP',
                       'POPEast',
                       'POPSouth',
                       'POPTejas',
                       'PPMMAC',
                       'PracticeMax',
                       'Reventics',
                       'Sheridan_ECW',
                       'Valley',
                       'Zotec']

    default_month, default_year = getReportingDefaults()

    layout1 = [[sg.Text('Get ready to run some Scripts!')],
               [sg.Text('Reporting Month: '),
                sg.Combo(reporting_months, key='rmonth', default_value=default_month, s=(5, 20), enable_events=False, readonly=True)],
               [sg.Text('Reporting Year:   '),
                sg.Combo(reporting_years, key='ryear', default_value=default_year, s=(5, 20), enable_events=False, readonly=True)],
               [sg.VPush()],
               [sg.VPush()],
               [sg.Column([[sg.Checkbox('Acute', key='Acute')],
                           [sg.Checkbox('Jupiter_ECW', key='Jupiter_ECW')],
                           [sg.Checkbox('NFPHI', key='NFPHI')],
                           [sg.Checkbox('POPEast', key='POPEast')],
                           [sg.Checkbox('PPMMAC', key='PPMMAC')],
                           [sg.Checkbox('Sheridan_ECW', key='Sheridan_ECW')]]), sg.Push(),
                sg.Column([[sg.Checkbox('AS400', key='AS400')],
                           [sg.Checkbox('McKesson', key='McKesson')],
                           [sg.Checkbox('Northside', key='Northside')],
                           [sg.Checkbox('POPSouth', key='POPSouth')],
                           [sg.Checkbox('PracticeMax', key='PracticeMax')],
                           [sg.Checkbox('Valley', key='Valley')]]), sg.Push(),
                sg.Column([[sg.Checkbox('GE', key='GE')],
                           [sg.Checkbox('MediCorp', key='MediCorp')],
                           [sg.Checkbox('POP', key='POP')],
                           [sg.Checkbox('POPTejas', key='POPTejas')],
                           [sg.Checkbox('Reventics', key='Reventics')],
                           [sg.Checkbox('Zotec', key='Zotec')]]), sg.Push()],
               [sg.VPush()],
               [sg.VPush()],
               [sg.Button('Go'), sg.Button('Exit')]]

    hidden_elements = [[sg.Text('', key=f'-FILE{i}-', visible=False), sg.Image(icons[0], key=f'-IMAGE{i}-', visible=False)] for i in range(MAX_NUM_SCRIPTS)]
    layout2 = [[sg.Text('Initializing...', key='-STATUS-')],
               [sg.ProgressBar(100, 'h', size=(80, 20), k='-PROGRESS-', expand_x=True)],
               [sg.VPush()]]
    layout2.extend(hidden_elements)

    layout = [[sg.Column(layout1, key='-COL1-'),
               sg.Column(layout2, visible=False, key='-COL2-')]]

    window = sg.Window('Month End Script Runner', layout, finalize=True, relative_location=(-300, -100))

    current_layout = 1
    work_id = 0
    scripts = []
    current_script = 0
    gui_queue = queue.Queue()
    script_lock = False

    # Event Loop
    while True:
        event, values = window.Read(timeout=100)

        if event is None or event == 'Exit':
            break
        if event == 'Go':
            script_paths = [join(main_filepath, 'Actuals', f'{values["ryear"]}-{values["rmonth"]}'),
                            join(main_filepath, 'Aging by Payor', f'{values["ryear"]}-{values["rmonth"]}'),
                            join(main_filepath, 'AR GT365 Lost Reimb Impact', f'{values["ryear"]}-{values["rmonth"]}'),
                            join(main_filepath, 'Consolidated Insurance', f'{values["ryear"]}-{values["rmonth"]}'),
                            join(main_filepath, 'Estimates', f'{values["ryear"]}-{values["rmonth"]}'),
                            join(main_filepath, 'GE Patient Level Refund Liability', f'{values["ryear"]}-{values["rmonth"]}'),
                            join(main_filepath, 'Reserve Valuation', f'{values["ryear"]}-{values["rmonth"]}')]
            bsystems = getBillingSystems(billing_systems, values)
            scripts = getAllBSScripts(bsystems, script_paths)

            window[f'-COL{current_layout}-'].update(visible=False)
            current_layout = 2
            window[f'-COL{current_layout}-'].update(visible=True)
            window.refresh()

            for i in range(len(scripts)):
                window[f'-FILE{i}-'].update(scripts[i], visible=True)
                window[f'-IMAGE{i}-'].update(visible=True)
            window[f'-STATUS-'].update('Connecting to the database...')
            window.refresh()

            script_lock = True

        if script_lock and current_script < len(scripts):
            script_lock = False

            window[f'-STATUS-'].update(f'Running the {scripts[current_script]} script...')
            window[f'-IMAGE{current_script}-'].update(icons[1])
            window.refresh()

            thread_id = threading.Thread(target=runScript, args=(scripts[current_script], work_id, gui_queue), daemon=True)
            thread_id.start()
            work_id += 1

        try:
            message = gui_queue.get_nowait()
        except queue.Empty:
            message = None

        if message is not None:
            if 'ERROR' in message:
                window[f'-IMAGE{current_script}-'].update(icons[3])
            else:
                window[f'-IMAGE{current_script}-'].update(icons[2])

            window['-PROGRESS-'].update_bar(current_script+1, len(scripts))
            window.refresh()

            current_script += 1
            script_lock = True

        if current_script == len(scripts):
            window[f'-STATUS-'].update('All done! Have a great day!')
            window.refresh()

    window.close()


def runScript(scriptname, work_id, gui_queue):
    try:
        driver = 'DRIVER'
        server = r'SERVER'
        database = 'DATABASE'
        credentials = 'Trusted_Connection=yes;'

        connection_string = f'{driver}{server}{database}{credentials}'
        connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
        engine = create_engine(connection_url)

        with open(scriptname) as s:
            script = re.sub(r'USE DATABASE;', '', s.read(), flags=re.IGNORECASE)

        output = pd.read_sql(script, engine)
        output = output.fillna("NULL")
        output.to_csv(join('outputs', basename(scriptname).replace('.sql', '.txt')), index=False, sep='\t')

        gui_queue.put(f'{work_id} ::: DONE')
    except:
        gui_queue.put(f'{work_id} ::: ERROR')

    return


if __name__ == '__main__':
    main()
