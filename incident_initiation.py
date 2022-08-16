#!/usr/bin/env python3

# File name: incident_initiation.py
# Author: Kyle LaPolice
# Date created: 1 - Jul - 2022
# Date last modified: 16 - Aug - 2022
# Script Version: 1.0

from pathlib import Path

import PySimpleGUI as sg # pip install PySimpleGUI
import pandas as pd # pip install pandas openpyxL
from docxtpl import DocxTemplate # pip install docxtpl

#remove data validation warnings from excel doc
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

window_name = 'Incident Initiation' # Change name of window per project

# Layout of window
layout = [
    [sg.Text("Select Incident Log File:")],
    [sg.Input(key='EXCEL', enable_events = True), sg.FileBrowse(key = 'eBROWSE')],

    [sg.Text("Select Incidents to Initiate:")],
    [sg.Listbox(values=[], key = 'LISTBOX', size = (0,10), expand_x = True, select_mode = sg.LISTBOX_SELECT_MODE_MULTIPLE)],

    [sg.Text("Enter Name:")],
    [sg.Input(key='NAME', enable_events = True, default_text = "")],

    [sg.OK(key = 'GO'), sg.Cancel(key = 'Exit')]
    ]

# Creates window that user interacts with
window = sg.Window(window_name, layout)

# Run the Event Loop
while True:
    event, values = window.read()

    # Close window if events
    if event == 'Exit' or event == sg.WIN_CLOSED:
        break

    # updates Listbox with Excel sheets
    if event == 'EXCEL':
        # create dataframe with just incident num colum and event date
        e_sheets = pd.read_excel(values.get('EXCEL'), sheet_name="Equipment Incident Log", usecols=['Equipment Incident #','Event Date'])

        # drop the empty rows, remove event date column, reverse the order
        e_dropna = e_sheets.dropna()
        e_nums = e_dropna.drop(columns = 'Event Date')
        e_reverse = e_nums.iloc[::-1]

        # convert dataframe to list and send to listbox
        window['LISTBOX'].update(values = list(e_reverse.values.tolist()))
        continue

    # executes program
    if event =='GO':

        # gets data from excel sheet chosen from Listbox
        e_sheets = pd.read_excel(values.get('EXCEL'), sheet_name = 'Equipment Incident Log')
        keep_list = values.get('LISTBOX')

        # converts from list of lists to just a list
        joined_list = []
        for items in values.get('LISTBOX'):
            joined_list.append(items[0])

        # only selects rows that were selected in the listbox
        cleaned = e_sheets[e_sheets['Equipment Incident #'].isin(joined_list)]

        #add user name input
        cleaned = cleaned.assign(Name=values.get('NAME'))
        # cleaned.loc[:,'Name'] = values.get('NAME')
        # cleaned['Name'] = values.get('NAME')
        print(cleaned)

        # removes from the column headers "spaces" and "#", and then removes "/" from the entire dataframe
        cleaned.columns = cleaned.columns.str.replace(' ','_')
        cleaned.columns = cleaned.columns.str.replace('#','')
        cleaned = cleaned.replace("/","", regex=True)

        # converts date format
        cleaned['Event_Date'] = cleaned['Event_Date'].dt.strftime('%d-%b-%Y')

        # sets directories
        base_dir =  Path(__file__).parent
        template_path = base_dir / '_Template' / 'FRM-1297_B Incident Form.docx'

        # itterates over template to create word docs from templates
        for record in cleaned.to_dict(orient="records"):
            doc = DocxTemplate(template_path)
            doc.render(record)

            # dicrectory to new incident folder lableing it via incident, strips whitespace at ends
            output_dir = base_dir / f"{record['Equipment_Incident_']} ~ {record['EQ_Description']}".strip()

            # checks if file already exists, if it does not exist creates the folder and saves the docx in it
            if not output_dir.is_dir():
                output_dir.mkdir(exist_ok=True)
                output_path = output_dir / 'FRM-1297_B Incident Form.docx'
                doc.save(output_path)
        continue

window.close()
