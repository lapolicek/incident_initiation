#!/usr/bin/env python3

# File name: Incident_Initiation.py
# Author: Kyle LaPolice
# Date created: 1 - Jul - 2022
# Date last modified: 20 - Jan - 2022
# Script Version: 1.3

from pathlib import Path
from typing import Any

import PySimpleGUI as sg # pip install PySimpleGUI
import pandas as pd # pip install pandas openpyxL
from docxtpl import DocxTemplate # pip install docxtpl

#remove data validation warnings from excel doc
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Change name of window per project
window_name = 'Incident Initiation'

# Layout of window
layout = [
    [sg.Text("Select Incident Log File:")],
    [sg.Input(key = 'EXCEL', enable_events = True), sg.FileBrowse(key = 'eBROWSE')],

    [sg.Text("Select Incidents to Initiate: (Hold CTRL to select multiple rows)")],
    [sg.Table(values = [],
              headings = ['Incident #', 'Event Date', 'Equipment #', '     EQ Description     ', 'Assigned to?'],
              key = 'TABLE',
              expand_x = True,
              justification = 'left',
              enable_events = True,
              select_mode=sg.TABLE_SELECT_MODE_EXTENDED)],

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

    # updates table with Excel sheet
    if event == 'EXCEL':
        # create dataframe with incident num colum, event date, eq num, eq desc, assigned
        e_sheets = pd.read_excel(values.get('EXCEL'), sheet_name = "Equipment Incident Log")

        e_dropna = e_sheets.dropna(subset=['Event Date'])   # drop the rows with empty cells under Event Date
        e_dropna['Event Date'] = e_dropna['Event Date'].dt.strftime('%d-%b-%Y') # reformat event date column to dd/mmm/yyyy
        e_reverse = e_dropna.iloc[::-1] # reverse the order
        e_index = e_reverse.reset_index(drop=True) # resets the index to new pruned and reversed dataframe

        # create dataframe with columns used in table
        table_values = e_index[['Equipment Incident #','Event Date', 'EQ #', 'EQ Description', 'Assigned to?']]

        # convert dataframe to list and send to table
        window['TABLE'].update(values = list(table_values.values.tolist()))
        continue

    # executes program
    if event =='GO':

        # converts from selected table row to specific eq incident num and adds to list
        keep_list = values.get('TABLE')
        joined_list = []
        for items in values.get('TABLE'):
            row = e_index['Equipment Incident #'][items]
            joined_list.append(row)

        # creates dataframe from rows that were selected in the TABLE
        cleaned = e_index[e_index['Equipment Incident #'].isin(joined_list)]

        #add user name input to dataframe
        cleaned = cleaned.assign(Name = values.get('NAME'))
        print(cleaned)

        # removes from the column headers "spaces" and "#", and then removes "/" from the entire dataframe
        cleaned.columns = cleaned.columns.str.replace(' ','_')
        cleaned.columns = cleaned.columns.str.replace('#','')

        # removes "/" and """ from the entire dataframe
        cleaned = cleaned.replace("/","", regex=True)
        cleaned = cleaned.replace("\"","", regex=True)
        cleaned = cleaned.replace("\t","", regex=True)

        # sets directories
        base_dir =  Path(__file__).parent
        template_path = base_dir / '_Template' / 'FRM-1297_C Incident Form.docx'

        # itterates over template to create word docs from templates
        for record in cleaned.to_dict(orient="records"):
            doc = DocxTemplate(template_path)
            doc.render(record)

            # dicrectory to new incident folder lableing it via incident, strips whitespace at ends
            output_dir = base_dir / f"{record['Equipment_Incident_']} ~ {record['EQ_Description']}".strip()

            # checks if file already exists, if it does not exist creates the folder and saves the docx in it
            if not output_dir.is_dir():
                output_dir.mkdir(exist_ok=True)
                output_path = output_dir / 'FRM-1297_C Incident Form.docx'
                doc.save(output_path)
        continue

window.close()
