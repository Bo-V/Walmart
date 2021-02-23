#
# Create a simple graphic interface
#
# Added in OS check


import PySimpleGUI as sg
import pandas as pd
import numpy as np
import sys
import os
import os.path
import tkinter as tk
import time
import requests
import openpyxl

from sys import platform
from os.path import expanduser



##########################################################################################
##########################################################################################


def exporta11y():
    myHeaderFont = 'Arial 22'
    myTextFont1 = 'Arial 15'
    myInputFont1 = 'Arial 13'
    myTextFont2 = 'Arial 14'
    myInputFont2 = 'Arial 12'
    myTextFont3 = 'Arial 12'
    myInputFont3 = 'Arial 11'
    myTextFont4 = 'Arial 11'        

    # Check device
    myosis = platform
    if myosis == "darwin":
        myosis = "Mac"
    if myosis == "win32":
        myosis = "Win"

    # print(platform)
    # print(myosis)

    sg.theme('Dark Blue 17')    # Keep things interesting for your users

    layout = [[sg.Text('Select the CSV file you have exported from platform',
                       font=myHeaderFont, pad=((5, 5), (0, 5)))],
              
        [sg.Text('Export status:', font=myTextFont2, pad=((5, 5), (0, 20))), sg.Text("                                                         " ,
            key="feedback", font=myInputFont2, pad=((5, 5), (0, 20)))],

        [sg.Text('Platform Export:', font=myTextFont1, size=(14, 0), pad=((5, 0), (0, 20))), 
        sg.InputText(size=(70, 0), font=myInputFont1, pad=((0, 0), (0, 20))), 
        sg.FileBrowse(key='input_file', font= myTextFont4, size=(8, 1), pad=((15, 5), (0, 20)), 
                      file_types= (('CSV Files', '*.csv'),('All Files', '*.*')),
                      button_color=('black', 'lightblue'))],

        [sg.Text('Enter the Applause cycle ID:', font=myTextFont3, size=(22, 0), pad=((5, 5), (0, 15))), 
        sg.Input(key='cycleID', size=(15, 0), font=myInputFont3, pad=((5, 5), (0, 20))),

        sg.Text('Related ticket in Jira (Story):', size=(22, 0), font=myTextFont3, pad=((35, 5), (0, 15))),
        sg.Input(default_text='CEAQA-xxx',key='story_jira', size=(14, 0), font=myInputFont3, pad=((5, 5), (0, 20))),
        sg.Text()],


        [sg.Text('Auto-populated labels:', font=myTextFont2, size=(18, 0), pad=((5, 5), (0, 10))),
        sg.Input(default_text='Applause-Cycle-[XXXXXX] WCAG_[X.X.X] Applause-[Man/Auto]',key='auto_labels', size=(79, 1), font=myInputFont2, pad=((5, 5), (0, 15)))],

        [sg.Text('Check default labels:', size=(17, 0), font=myTextFont2, pad=((5, 5), (0, 10))),
        sg.Input(default_text='ADA Applause applause_accessibility_cycle',key='standard_labels', size=(80, 1), font=myInputFont2, pad=((5, 5), (0, 15)))],

        [sg.Text('Labels separated by space:', size=(22, 0), font=myTextFont2, pad=((5, 5), (0, 10))),
        sg.Input(default_text=' ',key='labels_from_jira', size=(74, 0), font=myInputFont2, pad=((5, 5), (0, 15)))],
     
        [sg.Button('Export', font= myTextFont4, size=(25, 1), border_width=3, pad=((250, 5), (10, 5)),
                   button_color=('black', 'green')),
         sg.Button("Quit", size=(10,1),font= myTextFont4, border_width=3,pad=((20, 5), (10, 5)), 
                   button_color=('black', 'red'))]
        ]

    #root = tk.Tk()
    # eliminate the titlebar
    #root.overrideredirect(1)
    # Create the window
    window = sg.Window("Walmart A11y Exporter", layout, margins=(15, 15))
    myosis = myosis.lower()
    # print(myosis)


    while True:
        event, values = window.Read()
        if event is None or event == 'Quit':
            window.close()
            manimenu()

        window['feedback'].Update("Ready to Export") # show the button in the feedback text
        if event == 'Export':
            if len(sys.argv) == 1:
                fname = values[0]
            

        else:
            fname = sys.argv[1]

        if not fname:
            sg.Popup("Cancel", "No filename supplied")
            
            window.close()
            exporta11y()
            raise SystemExit("Cancelling: no filename was supplied ")
        else:
            sg.Popup('Exporting Using This File:', 
                      fname)
            

        csv_file = values['input_file']
        cycle = values['cycleID']
        

        # Bo
        mypath = expanduser("~")
        ostype = myosis #values['MyOS']
        ostype = ostype.lower()
        # print(ostype)

        # Original 

        xlsx_file = cycle + 'tempfile.xlsx'

        pd.read_csv(csv_file).to_excel(xlsx_file, engine='xlsxwriter')

        df = pd.read_excel(xlsx_file)

        df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '').str.replace('?', '')
        
        # print(df.columns)
    #   Automaticly update caught with automation label
    #
        dfman_auto = df['caught_with_automation']
        autoy = ' Applause_Auto'
        auton = ' Applause-Man'
        #print(len(dfman_auto))
        man_auto1 = [None] * len(dfman_auto)
        labels = [None] * len(dfman_auto)
        
        for i in range(0, len(dfman_auto)):
            if dfman_auto[i] == 'Yes':
                man_auto1[i] = autoy 
            #    print(str(man_auto1[i]))
                story_jira = ' ' + values['story_jira']
                auto_labels = 'Applause-Cycle-' + cycle + str(man_auto1[i]) + story_jira
                labels[i] = auto_labels + ' ' + values['standard_labels'] + ' ' + values['labels_from_jira']   
            #    print(labels[i]) 
            
            if dfman_auto[i] == 'No':
                man_auto1[i] = auton
                #    print(str(man_auto1[i])) 
                story_jira = ' ' + values['story_jira']
                auto_labels = 'Applause-Cycle-' + cycle + str(man_auto1[i]) + story_jira
                labels[i] = auto_labels + ' ' + values['standard_labels'] + ' ' + values['labels_from_jira']   
                #    print(labels[i])
                
                # we're pulling just our ID and attachment columns to split them into their own columns
                df['bugid'] = df['id']
                links = df[['id',
                            'bugid',
                            'attachments']]

                links.set_index('id', inplace=True)
                split = links['attachments'].str.split('%3D', expand=True)

                # after splitting our attachments - remove the last empty column
                col_count = len(split.columns) - 1
                split.reset_index(inplace=True)

                split.drop(col_count,
                        axis='columns',
                        inplace=True)

                # put our columns into rows
                split = pd.melt(split,
                                id_vars=['id'],
                                var_name='column_number',
                                value_name='url')

                split.drop('column_number',
                        axis='columns',
                        inplace=True)

                # now taking our url column and pulling out the attachment name
                split2 = split['url'].str.split(': ', expand=True)
                split2[1] = split2[1] + '%3D'

                split2.replace(r'\n',
                            '',
                            regex=True,
                            inplace=True)

                # put our attachment names and URLs to lists - filter out None values to be safe
                title = split2[0].tolist()
                titleFiltered = list(filter(None,
                                            title))
                url = split2[1].tolist()
                # urlFiltered = list(filter(None, url))
                urlFiltered = [i for i in url if i is not np.nan]
                #print(titleFiltered)
                #print(urlFiltered)
                urls = len(urlFiltered)
                #print(len(urlFiltered))
                #print(len(titleFiltered))

        # save dataframe to a text file - this is only useful for quick troubleshooting
        #txt = np.savetxt('attachments.txt',
        #                split2.values,
        #                delimiter=' ',
        #                fmt='%s')
        
        # Checks for users base path plus /downloads/Walmart_exports, if available 
        # We switch to working directory, if not available we create it
        #
        if os.path.exists(mypath + '/downloads/Walmart_exports'):
            os.chdir(mypath + '/downloads/Walmart_exports')    
        else:
            os.makedirs(mypath + '/downloads/Walmart_exports')
        # print(mypath)
        # exit()
        #mypath = os.path.expanduser(path)
        #
        # Mac needs directory set - Bo
        #
        if myosis == "mac":
            os.chdir(mypath + '/downloads/Walmart_exports')
        if os.path.exists('attachment_downloads'):
            os.rename('attachment_downloads',
                    'attachments' + str(time.time()))
            os.makedirs('attachment_downloads')
        else:
            os.makedirs('attachment_downloads')
        # Mac needs directory set - Bo
        #
        if myosis == "mac":
            os.chdir(mypath + '/downloads/Walmart_exports/attachment_downloads/')

        filecounter = 0
        window['feedback'].Update("Downloading Attachments") # show the status in the feedback text

        # loop through attachment list and download each file with the
        for i in range(urls):
            local_filename = titleFiltered[i]

            filecounter = filecounter+1
            window['feedback'].Update("Downloading Attachment #: " + str(filecounter)) # show the button in the feedback text

            with requests.get(urlFiltered[i], stream=True) as r:
                r.raise_for_status()
                
                #print('downloading attachment ' + str([i]) + ' ' + str(titleFiltered[i]))

                #print(current_dir)
                #print('status code returned ' + str(r.status_code))              

                # Win - For Windows    
                if myosis == "win":
                    #print(myosis)
                    #print("Windows")
                    window.Refresh()
                    os.chdir(mypath + '/downloads/Walmart_exports/')
                    with open(r'attachment_downloads\\' + local_filename, 'wb') as f:

                        for chunk in r.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                                f.flush()
                                    
                # mac - When using a Mac
                if myosis == "mac":
                    with open(local_filename, 'wb') as f:
                        #print(os.curdir)
                        window.Refresh()
                        
                        
                        for chunk in r.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                                f.flush()
                        
                        
        window['feedback'].Update("Renaming Attachments") # show the button in the feedback text

        # rename all files really quick
        # Mac needs directory set - Bo
        if myosis == "mac":
            os.chdir(mypath + '/downloads/Walmart_exports')
                
        filenames = os.listdir('attachment_downloads')
        directory = 'attachment_downloads'
        remove = 'Bug'
        replacement = ''
        for filename in filenames:
            # print("File name: " + filename)
            window.refresh()

            # mac - When using a Mac
            if myosis == "mac":
            #    print('Renaming files Mac')
                os.chdir(mypath + '/downloads/Walmart_exports')
                os.rename( directory + "/" + filename, directory + "/" + filename.replace(remove, replacement))
            
            # Win - For Windows
            if myosis == "win":
            #    print('Renaming files Win')
                os.chdir(mypath + '/downloads/Walmart_exports/')
                os.rename(r'attachment_downloads\\' + filename, directory + '\\' + filename.replace(remove, replacement))



        #print(df.dtypes)
        #print(list(man_auto1))
        

        # create new dataframe from concatenated fields
        jira_id = df['id'].map(str)

        jira_summary = df['title'].map(str)

        jira_description = '*Action Performed:*' + '\r\n' + \
                        df['action_performed'].map(str) + '\r\n' + '\r\n' + \
                        '*Expected Result:* ' + '\r\n' + df['expected_result'].map(str) + '\r\n' + '\r\n' + \
                        '*Actual Result:* ' + '\r\n' + df['actual_result'].map(str) + '\r\n' + '\r\n' + \
                        '*Suggested resolutions:* ' + '\r\n' + df['suggested_resolutions'].map(str) + '\r\n' + '\r\n' + \
                        'Area issue was found: ' + df['area_issue_was_found'].map(str) + '\r\n' + \
                        'Failed WCAG 2.1 checkpoint(s):' + df['failed_wcag_2.1_checkpoints'].map(str) + '\r\n' + \
                        'Additional Info: ' + '\r\n' + df['additional_environment_info'].map(str) + '\r\n' + '\r\n'+\
                        'Applause ID: ' + df['id'].map(str) + '\r\n' +\
                        'Applause URL: ' + 'https://platform.utest.com/testcycles/' + cycle + '/issues/' + df['id'].map(str) + '/'
        jira_labels = labels
        #print(labels)
        jira_priority = df['priority']
        jira_automation = df['caught_with_automation']
        jira_wcag = df['failed_wcag_2.1_checkpoints']

        d = {'bug_id': jira_id,
            'priority': jira_priority,
            'summary': jira_summary,
            'description': jira_description,
            'caught_with_autom': jira_automation,
            'wcag': jira_wcag,
            'wcag_label': 'WCAG_' + jira_wcag,
            'labels': jira_labels}

        # Writes data to Jira_imports
        jira_import = pd.DataFrame(data=d)
        
        # Sorts sheet on Priority column
        #jira_import = jira_import.sort_values(by=['priority'])

        # Use these output files for import
        
        
        excel_file = jira_import.to_excel(cycle+'_walmart_export_ready.xlsx')
        excel_file = cycle + '_walmart_export_ready.xlsx'
        csv_file = cycle + '_walmart_export_ready.csv'
        pd.read_excel(excel_file).to_csv(csv_file)
        window['feedback'].Update("Renaming Directory") # show the button in the feedback text
        window.refresh()
        # Renames existing attahment directory if matching exisits
        if os.path.exists(cycle + '_attachments'):
            os.rename(cycle + '_attachments',
                        cycle + '_attachments' + str(time.time()))
        # Renames temp attachments directory to match cycle number
        os.rename('attachment_downloads',
                        cycle + '_attachments')

        sourcefile = mypath + '/downloads/Walmart_exports/' + cycle + '_walmart_export_ready.xlsx'
        destinationfile = mypath + '/downloads/Walmart_exports/' + cycle + '_attachments/' + cycle + '_walmart_export_ready.xlsx'
        os.rename(sourcefile, destinationfile)

        sourcefile1 = mypath + '/downloads/Walmart_exports/' + cycle + '_walmart_export_ready.csv'
        destinationfile1 = mypath + '/downloads/Walmart_exports/' + cycle + '_attachments/' + cycle + '_walmart_export_ready.csv'
        os.rename(sourcefile1, destinationfile1)


        window['feedback'].Update("Export Complete") # show the button in the feedback text    
        window.refresh()
        #if event is None or event == '      Quit      ':
        #    break
        # print(event, values)
            

        #
    window.Close()

    manimenu()



    ##########################################################################################
    ##########################################################################################


    #  Function for Glass Exports

def exportGlass():


    sg.theme('DarkBlue2')    # Keep things interesting for your users

    layout = [[sg.Text('Select the csv with your bug exports')],
            [sg.Text('Platform Export',
                    size=(15, 1)),
                    sg.InputText(),
                    sg.FileBrowse(key='input_file')],
            [sg.Text('Please enter the cycle ID:')],
            [sg.Input(key='cycleID')],

            [sg.Text('Auto-populated labels (no update from UI):')],
            [sg.Input(default_text='Applause-Cycle-[XXXXXX] WCAG_[X.X.X] Applause-[Man/Auto]',key='auto_labels')],

            [sg.Text('Check default labels:')],
    #          [sg.Input(default_text='Glass-Mobile-Platform Glass-Android-Bug ADA Applause applause_accessibility_cycle',key='standard_labels')],
            [sg.Input(default_text='Applause Applause_Glass_Ordering glass-production-issue',key='standard_labels')],

            [sg.Text('Enter labels from Jira separated by space:')],
            [sg.Input(default_text='',key='labels_from_jira')],

            [sg.Text('Related ticket in Jira (Story):')],
            [sg.Input(default_text='',key='story_jira')],
            [sg.Submit(),
            sg.Exit()]]

    window = sg.Window('Walmart Fn Exporter',
                    layout)

    while True:
        event, values = window.Read()

        if event == 'Submit':
            if len(sys.argv) == 1:
                fname = values[0]

            else:
                fname = sys.argv[1]

            if not fname:
                sg.Popup("Cancel", "No filename supplied")
                window.close()
                exportGlass()
                raise SystemExit("Cancelling: no filename was supplied")
            else:
                sg.Popup('The filename you chose was', fname)

            csv_file = values['input_file']
            cycle = values['cycleID']
            
            xlsx_file = cycle + 'tempfile.xlsx'
            os.remove('tempfile.xlsx')
            pd.read_csv(csv_file).to_excel(xlsx_file, engine='xlsxwriter')
            df = pd.read_excel(xlsx_file)
            df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '').str.replace('?', '')

    #        print (df['caught_with_automation'])
    #        if df['caught_with_automation'] == 'Yes':
    #          man_auto = "Applause-Auto"
    #        else:
    #        man_auto = " Applause-Man"
            story_jira = ' ' + values['story_jira']
            
            auto_labels = 'Applause-Cycle-' + cycle + story_jira
            labels = auto_labels + ' ' + values['standard_labels'] + ' ' + values['labels_from_jira']

            # we're pulling just our ID and attachment columns to split them into their own columns
            df['bugid'] = df['id']
            links = df[['id',
                        'bugid',
                        'attachments']]

            links.set_index('id', inplace=True)
            split = links['attachments'].str.split('%3D', expand=True)

            # after splitting our our attachments - remove the last empty column
            col_count = len(split.columns) - 1
            split.reset_index(inplace=True)

            split.drop(col_count,
                    axis='columns',
                    inplace=True)

            # put our columns into rows
            split = pd.melt(split,
                            id_vars=['id'],
                            var_name='column_number',
                            value_name='url')

            split.drop('column_number',
                    axis='columns',
                    inplace=True)

            # now taking our url column and pulling out the attachment name
            split2 = split['url'].str.split(': ', expand=True)
            split2[1] = split2[1] + '%3D'

            split2.replace(r'\n',
                        '',
                        regex=True,
                        inplace=True)

            # put our attachment names and URLs to lists - filter out None values to be safe
            title = split2[0].tolist()
            titleFiltered = list(filter(None,
                                        title))
            url = split2[1].tolist()
            # urlFiltered = list(filter(None, url))
            urlFiltered = [i for i in url if i is not np.nan]
            print(titleFiltered)
            print(urlFiltered)
            urls = len(urlFiltered)
            print(len(urlFiltered))
            print(len(titleFiltered))

            # save dataframe to a text file - this is only useful for quick troubleshooting
            #txt = np.savetxt('attachments.txt',
            #                split2.values,
            #                delimiter=' ',
            #                fmt='%s')

            if os.path.exists('attachment_downloads'):
                os.rename('attachment_downloads',
                        'attachments' + str(time.time()))
                os.makedirs('attachment_downloads')
            else:
                os.makedirs('attachment_downloads')

            # loop through attachment list and download each file with the
            for i in range(urls):
                local_filename = titleFiltered[i]
                with requests.get(urlFiltered[i], stream=True) as r:
                    r.raise_for_status()
                    

                    print('downloading attachment ' + str([i]) + ' ' + str(titleFiltered[i]))
                    
                    


                    print('status code returned ' + str(r.status_code))
                    with open(r'attachment_downloads\\' + local_filename, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                                f.flush()

            # rename all files really quick
            filenames = os.listdir('attachment_downloads')
            directory = 'attachment_downloads'
            remove = 'Bug'
            replacement = ''
            for filename in filenames:
                print(filename)
                os.rename(r'attachment_downloads\\' + filename, directory + '\\' + filename.replace(remove, replacement))
            print(df.dtypes)

            # create new dataframe from concatenated fields
            jira_id = df['id'].map(str)

            jira_summary = df['title'].map(str)

            jira_description = '*Action Performed:*' + '\r\n' + \
                            df['action_performed'].map(str) + '\r\n' + '\r\n' + \
                            '*Expected Result:* ' + '\r\n' + df['expected_result'].map(str) + '\r\n' + '\r\n' + \
                            '*Actual Result:* ' + '\r\n' + df['actual_result'].map(str) + '\r\n' + '\r\n' + \
                            '*Error message:* ' + '\r\n' + df['error_message'].map(str) + '\r\n' + '\r\n' + \
                            '*Environment:* ' + '\r\n' + df['environment'].map(str) + '\r\n' + df['community_reproductions'].map(str) + '\r\n' + '\r\n' + \
                            'Additional Info: ' + '\r\n' + df['additional_environment_info'].map(str) + '\r\n' + '\r\n'+\
                            'Applause ID: ' + df['id'].map(str) + '\r\n' +\
                            'Applause URL: ' + 'https://platform.utest.com/testcycles/' + cycle + '/issues/' + df['id'].map(str) + '/'
            jira_labels = labels
    #        jira_priority = df['priority']
    #        jira_automation = df['caught_with_automation']
    #        jira_wcag = df['failed_wcag_2.1_checkpoints']

            d = {'bug_id': jira_id,
    #             'priority': jira_priority,
                'summary': jira_summary,
                'description': jira_description,
    #             'caught_with_autom': jira_automation,
    #             'wcag': jira_wcag,
    #             'wcag_label': 'WCAG_' + jira_wcag,
            'labels': jira_labels}

            jira_import = pd.DataFrame(data=d)


            # Use these output files for import
            excel_file = jira_import.to_excel(cycle+'_walmart_export_ready.xlsx')
            excel_file = cycle + '_walmart_export_ready.xlsx'
            csv_file = cycle + '_walmart_export_ready.csv'
            pd.read_excel(excel_file).to_csv(csv_file)

        if event is None or event == 'Exit':
            break
        # print(event, values)

    window.Close()
    manimenu()

##########################################################################################
##########################################################################################



def manimenu():
    # Main Menu
    # 
    # 
    #  
    appversion = "1.02"
    sg.theme('DarkBlue17')
    layout = [ [sg.Text("    Please Select Issue Type:                           " , key="feedback", pad=((5,5),(10, 0)), font='Arial 12')],
            
            [sg.Button("A11y Issues", size=(35,1),font='Arial 12', border_width=3,pad=(105, 20),
                       button_color=('black', 'lightblue'))],
            [sg.Button("Fn Issues", size=(35,1),font='Arial 12', border_width=3,pad=(105, 8), 
                       button_color=('black', 'lightblue'))],
            [sg.Button("QUIT", size=(15,1),font='Arial 12', border_width=3,pad=(160, 40), 
                       button_color=('black', 'red'))]
            ]

    # Create the Window
    window = sg.Window('    Walmart Issues Export Script      Version: ' + appversion, layout,
                       grab_anywhere=False,
                       keep_on_top=True,
                       no_titlebar=False, 
                       size=(450, 275),
                       font='Arial 12')

    # Event Loop to process "events"
    while True:
        event, values = window.read()
        window['feedback'].Update(event) # show the button in the feedback text
        if event == ('A11y Issues'):
            window.close()
            exporta11y()
        if event == ('Fn Issues'):
            window.close()
            exportGlass()
        if event in (None, 'QUIT'):
            exit()
        
    window.close()
manimenu() 
