import pandas as pd
from pandas import ExcelWriter
import numpy as np
import os
from datetime import date, datetime
import shutil
import PySimpleGUI as sg
import warnings

try:
    import pyi_splash
    pyi_splash.update_text('Can take upto a minute for the app to load.')
    pyi_splash.update_text('Loading.')
    pyi_splash.update_text('Loading..')
    pyi_splash.update_text('Loading...')
    pyi_splash.update_text('Loading....')
    pyi_splash.update_text('UI Loaded.')
    pyi_splash.close()
except:
    pass

sg.theme("LightGrey1")
sg.set_options(font=(14), border_width = 0)

def extract_data(directory_path, columns_to_extract, date_columns=[], rows_to_skip=0, output_filename='output.csv'):
    """
    Extracts data from all Excel and CSV files in the specified directory and its subdirectories that contain all the
    specified columns.

    :param directory_path: The path to the directory to search for Excel and CSV files.
    :param columns_to_extract: A list of column names to extract from each file.
    :param date_columns: A list of column names to parse as dates using pd.to_datetime().
    :param output_filename: The path and filename to save the extracted data in a csv
    :return: A DataFrame containing the extracted data from all Excel and CSV files that contain the specified columns,
    or None if no files contain the specified columns.
    """
    warnings.simplefilter(action='ignore', category=UserWarning)
    start_time = datetime.now()
    extracted_data = pd.DataFrame()
    columns_not_found = []
    files_with_no_columns = []
    sheet_read = ''
    extracted_columns = columns_to_extract.copy()
    extracted_columns.extend(['File Name','SubDir','CreatedDate','LastModifiedDate'])

    for root, dirs, files in os.walk(directory_path):
        for file in files:
            sg.Print('Started reading ' + file + ' at: {}'.format(datetime.now()) + ' ...')
            if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.csv'):
                file_path = os.path.join(root, file)
                columns_found_outer = False
                if file.endswith('.csv'):
                    df = pd.read_csv(file_path,skiprows=rows_to_skip,encoding_errors= 'replace')
                    for col in columns_to_extract:
                        if col in df.columns:
                            columns_found_outer = True
                            #sg.Print(col + 'in csv true loop')
                        else:
                            columns_found_outer = False
                            #sg.Print(col + 'in false loop')
                            break
                else:
                    dfs = pd.read_excel(file_path, sheet_name=None,skiprows=rows_to_skip)
                    for key in dfs.keys():
                        columns_found = False
                        for col in columns_to_extract:
                            if col in dfs[key].columns:
                                columns_found = True
                                #sg.Print(col + 'in true loop')
                            else:
                                columns_found = False
                                columns_not_found.append(col + " not found in sheet: " + key + " of file ---> " + file)
                                #sg.Print(col + 'in false loop')
                                break
                        if columns_found:
                            columns_found_outer = True
                            df = dfs[key]
                            sheet_read = key
                            break
                if columns_found_outer:
                    if len(date_columns) > 0:
                        for col in date_columns:
                            if col in df.columns:
                                df[col] = pd.to_datetime(df[col])
                    df['File Name'] = file
                    df['SubDir'] = os.path.basename(root)

                    createx = modx = os.path.getctime(file_path)
                    xcreate = datetime.fromtimestamp(modx)
                    df['CreatedDate'] = xcreate

                    modx = os.path.getmtime(file_path)
                    xmod = datetime.fromtimestamp(modx)
                    df['LastModifiedDate'] = xmod
                    extracted_data = pd.concat([extracted_data, df[extracted_columns]], ignore_index=True)
                    #sg.Print(xmod)
                    if file.endswith('.csv'):
                        sg.Print(file + ' has been read',text_color='white', background_color='green')
                    else:
                        sg.Print(file + ' has been read and it was last modified on ' + xmod.strftime('%Y-%m-%d') + '. The name of the sheet that was read is: ' + sheet_read,text_color='white', background_color='green')
                else:
                    files_with_no_columns.append(file_path)    
    if len(files_with_no_columns) > 0:
        sg.Print("The following files do not contain the specified columns:",text_color='white', background_color='red')
        for file_path in files_with_no_columns:
            sg.Print(file_path,text_color='red')
        if len(columns_not_found) > 0:
            for col in columns_not_found:
                sg.Print(col,text_color='blue')
    if not extracted_data.empty:
        extracted_data = extracted_data.applymap(lambda s: s.upper() if type(s) == str else s).fillna('')
        extracted_data.to_csv(output_filename)
        sg.Print('Started at: {}'.format(start_time) + '. \nEnded at: {}'.format(datetime.now()) + '. \nTime elapsed (hh:mm:ss.ms) {}'.format(datetime.now() - start_time))
        return extracted_data
    else:
        sg.Print("Specified columns do not exist in any file in the provided directory.",text_color='white', background_color='red')
        sg.Print('Started at: {}'.format(start_time) + '. \nEnded at: {}'.format(datetime.now()) + '. \nTime elapsed (hh:mm:ss.ms) {}'.format(datetime.now() - start_time))
        return None

#### Selction pop-up
def GUI_POPUP(text, data):
    layout = [
        [sg.Text(text, font = (any, 14, 'bold'))],
        [sg.Listbox(values=data, select_mode='multiple', size=(50,10), key='SELECTED')],
        [sg.Button('OK')],
    ]    
    window = sg.Window('Select Columns', layout,).Finalize()    
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'OK':
            #selected_options = values['SELECTED']
            #sg.popup(f'Selected options: {selected_options}')
            break
        else:
            sg.Print('OVER')
    window.close()

    if values and values['SELECTED']:
        return values['SELECTED']

def get_columns(upload_file):
    opt = []
    try:
        df = pd.read_csv(upload_file, nrows = 5, encoding_errors= 'replace')
        for key in df.keys():
            opt.append(key)
    except IOError as e:
            sg.popup_error("""
It appears that the file from which columns are to be selected is open.
Please close that csv file before closing this message to continue.
            """,no_titlebar = True, grab_anywhere = True,background_color = 'black', text_color='white')
    return opt
    
# Show result
def show_result(df):
    #sg.theme("LightGrey1")
    layout_1 = [[sg.Button('Close')],
                [sg.Table(key='-FINAL_TBL-',values = df.values.tolist(), headings = list(df.keys()), 
        auto_size_columns=True, vertical_scroll_only = False, size=(1000,800), 
        col_widths=list(map(lambda x:len(x)+1, list(df.keys()))))],
        ]
    window_result = sg.Window('Extracted Output', layout_1,resizable=True).Finalize()
    window_result.Maximize()
    while True:
        event, values = window_result.read()
        if event in (sg.WIN_CLOSED, 'Close'):
            break
    #sg.theme("LightGrey1")
    window_result.close(); del window_result
    
## Form Layout

column_layout = []    

#Sample
first_row_col1 = [
    [sg.Text('Select a sample csv file', font = (any, 14, 'bold'))],
    [sg.Input(key='-FIELD1-'), sg.FileBrowse(file_types=(("CSV Files", "*.csv"),),)],
]

#Columns from Sample
second_row_col1 = [
    [sg.Text('Columns to extract:', font = (any, 14, 'bold'))],
    [sg.Input('',key='-S1-',readonly=True), sg.Button('Select',key='SEL1',)],
]

#Date Columns
second_row_col2 = [
    [sg.Text('Columns to parse as dates:', font = (any, 14, 'bold'))],
    [sg.Input('',key='-S2-',readonly=True), sg.Button('Select',key='SEL2',)],
]

#Path for directory to be read - directory_path_value
third_row_col1 = [
    [sg.Text('Directory to read:', font = (any, 14, 'bold'))],
    [sg.Input(key='-READ_FOLDER-'), sg.FolderBrowse()],
]

#Save Location
third_row_col2 = [
    [sg.Text('Location to save output:', font = (any, 14, 'bold'))],
    [sg.Input(key='-WRITE_FOLDER-'), sg.FolderBrowse()],
]

#Provide Rows to skip
fourth_row_col1 = [
    [sg.Text('Rows to skip:', font = (any, 14, 'bold'))],
    [sg.Input(key='-SKIP_ROWS-',default_text='0')],
]

layout = [
            [sg.Column(first_row_col1,key='-Row1-')],
            [sg.Column(second_row_col1,key='-Row2-'),sg.Column(second_row_col2,key='-Row2Col2-')],
            [sg.Column(third_row_col1,key='-Row3-'),sg.Column(third_row_col2,key='-Row3Col2-')],
            [sg.Column(fourth_row_col1,key='-Row4-'),],
            [sg.Button('Extract Data',key='SUBMIT',button_color=('white', 'green')), sg.Button('Clear'), sg.Button('Exit', key='Exit',button_color=('white', 'red'))],
        ]

columns_to_extract = []
date_columns = []
directory_path_value = ''
output_filename_value = ''
rows_to_skip_value = 0

window = sg.Window('Recursively Extract Excel / CSV Data', 
                   layout, 
                   enable_close_attempted_event=True,
                   resizable=True,
                   element_justification = 'c',
                  )
while True:  # Event Loop
    event, values = window.read()
    if (event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == 'Exit') and \
    sg.popup_yes_no('Do you really want to exit?',no_titlebar = True, grab_anywhere = True,background_color = 'black', text_color='white') == 'Yes':
        break
        
    if event == 'SEL1' and values['-FIELD1-']:
        selected = GUI_POPUP('Select Columns', get_columns(values['-FIELD1-']))
        print('selected:', selected)
        if selected:
            window['-S1-'].update(selected)
            columns_to_extract = selected

    if event == 'SEL2' and values['-S1-']:
        selected = GUI_POPUP('Select Columns', columns_to_extract)
        print('selected:', selected)
        if selected:
            window['-S2-'].update(selected)
            date_columns = selected
            
    if event == 'Clear':
        window['-FIELD1-'].update('')
        window['-S1-'].update('')
        window['-S2-'].update('')
        columns_to_extract = []
        date_columns = []
        window['-READ_FOLDER-'].update('')
        directory_path_value = ''
        window['-WRITE_FOLDER-'].update('')
        output_filename_value = ''
        window['-SKIP_ROWS-'].update('0')
        rows_to_skip_value = 0
        print(columns_to_extract)
        print(date_columns)
        print(directory_path_value)
        print(output_filename_value)
        print(rows_to_skip_value)
    
    if event == 'SUBMIT':
        if values['-S1-']:
            sg.Print('columns_to_extract has value {}'.format(columns_to_extract))
            columns_flag = True
        else:
            columns_flag = False
        if values['-S2-']:
            sg.Print('date_columns has value {}'.format(date_columns))
        if values['-READ_FOLDER-']:
            directory_path_value = values['-READ_FOLDER-']
            sg.Print(directory_path_value)
            directory_flag = True
        else:
            directory_flag = False
        if values['-WRITE_FOLDER-']:
            output_filename_value = os.path.join(values['-WRITE_FOLDER-'],'extracted_data_'+str(date.today())+'.csv')
            print(output_filename_value)
            output_file_flag = True
        else:
            output_file_flag = False
        if values['-SKIP_ROWS-']:
            try:
                rows_to_skip_value = int(values['-SKIP_ROWS-'])
                if 0 <= rows_to_skip_value <= 50:
                    skiprows_flag = True
                else:
                    sg.popup('Invalid input: enter an integer between 1 and 50')
                    skiprows_flag = False
            except ValueError:
                sg.popup('Invalid input: enter an integer')
                skiprows_flag = False
        if columns_flag and directory_flag and output_file_flag and skiprows_flag:
            print('Ready to call extract_data function')
            df = extract_data(directory_path_value,columns_to_extract, date_columns, rows_to_skip_value, output_filename_value)
            if isinstance(df, pd.DataFrame):
                show_result(df)
                sg.popup("Output File:", "extracted_data_"+str(date.today())+".csv", "Saved at: ",values['-WRITE_FOLDER-'])
            
window.close(); del window