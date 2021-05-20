# -*- coding: utf-8 -*-
"""
Created on Tue Dec  1 15:53:34 2020

@author: Zoren.Liu

Download eTender spreadsheets.

Agency UUID

TfNSW (Roads and Maritime Projects)     9A049771-BCF1-66C0-DC82BD14117A76E8
Transport for NSW                       5C6E81DB-F27C-49D1-EE4FDD8E268C3100
Transport NSW - Corporate               9A059A35-A5E2-AC7D-E2CE015DBD1A6792
Transport NSW - Transport Services      E3282140-F3F9-C752-16E52B022B0F34BB

Export to .exe

pyinstaller --onefile FILENAME.py

"""

import numpy as np
from datetime import datetime
import subprocess
import os
# import time
import pandas as pd

#%% Messages header

print("==================================================================================")
print("                             eTender Export Tool v1.3                             ")
print("==================================================================================")
print("                              Created by Zoren Liu                                ")
print(" ")
print("Downloads are saved to folder ETENDER_DOWNLOAD, combined file saved to")
print("folder ETENDER_COMBINED, in current directory.")
print("")
print("Version history:")
print("v1.3 -  [2021-03-28] Added 'ExportDateTime' column to combined eTender csv file.")
print("        Column records the date and time of file creation.")
print("        Changed combined file name convention to YYYYMMDD")
print("v1.2 -  [2021-02-05] Added 'Source Name' column to combined eTender csv file.")
print("        Column records the name of original downloaded spreadsheet.")
print("v1.1 -  [2021-02-02] Channged combined file name to date of download, format: ")
print("        DDMMYYYY. Added version history")
print("v1.0 -  [2020-12-18] New release")
print("==================================================================================")
print(" ")

#%% Definitions

agencyUUID = np.array(["9A049771%2DBCF1%2D66C0%2DDC82BD14117A76E8",
                        "5C6E81DB%2DF27C%2D49D1%2DEE4FDD8E268C3100",
                        "9A059A35%2DA5E2%2DAC7D%2DE2CE015DBD1A6792",
                        "E3282140%2DF3F9%2DC752%2D16E52B022B0F34BB"
                        ])

agency_name = np.array(["TfNSW_RMS",
                        "TfNSW",
                        "TfNSW_Corp",
                        "TfNSW_Tran"
    ])

dl_dates = np.array([["1%2DJul%2D", "31%2DDec%2D"],
                      ["1%2DJan%2D", "30%2DJun%2D"]
                      ])

year_s = 2000
year_e = int(datetime.today().strftime('%Y'))

url_main = "https://www.tenders.nsw.gov.au/?event=public.advancedsearch.cnDownload&decorator=XLS&agencyUUID=AGENCY_UUID&agencyStatus=%2D1&keyword=&publishFrom=PUB_START&publishTo=PUB_END&valueFrom=&valueTo=&supplierName=&supplierABN=&RFTID=&contractFrom=&contractTo=&category=&Postcode=&piggyback=&download="

#%% Utility functions

def check_date_range():
    '''
    Check whether today's date falls in the first half or the second half of the year
    
    Inputs
    - N/A
    
    Returns
    - True - If today's date falls between 1-Jan-YYYY and 30-Jun-YYYY [bool]
    - False - If today's date falls between 1-Jul-YYYY and 31-Dec-YYYY [bool]
    '''
    return datetime(int(datetime.today().strftime('%Y')), 7, 1) > datetime.today()

def mod_url(url, search_list, mod_list):
    '''
    Modifies the template download url to replace the variable names with
    specified parameters
    
    Inputs
    - url - the template url [str]
    - search_list - list of variable names to be replace [list, str]
    - mod_list - list of parameters to replace the variables [list, str]
    
    Outputs
    - url - final modified url [str]
    '''
    for n in range(len(search_list)):
        url = url.replace(search_list[n], mod_list[n])
    return url

#%% Generating urls

date_range = check_date_range()

curl_func = []

search_list = np.array(["AGENCY_UUID", "PUB_START", "PUB_END"])

for i in range(len(agencyUUID)):
    file_count = 1
    for y in range(year_e, year_s-1, -1):
        for p in range(len(dl_dates)):
            if y == int(datetime.today().strftime('%Y')) and date_range == True and p == 0:
                continue
            else:        
                period_start = dl_dates[p][0] + str(y)
                period_end = dl_dates[p][1] + str(y)
                temp_url = mod_url(url_main, search_list, [agencyUUID[i], period_start, period_end])
                
                line = 'curl "' + temp_url + '" > eTender_' + agency_name[i] + '_' + period_end.replace('%2D', '-') + '_' + period_start.replace('%2D', '-') + '_' + str(file_count) + '.xls'
                file_count += 1
                curl_func.append(line)


#%% Creating folders and working directories

folders = ['ETENDER_DOWNLOAD', 'ETENDER_COMBINED']
wds = [os.getcwd()]

for folder in folders:

    process = subprocess.Popen("mkdir " + folder,
                          stdout=subprocess.PIPE, 
                          stderr=subprocess.PIPE,
                          shell = True)
    
    wds.append(os.getcwd() + '\\' + folder)
    
    flag = 1
    while folder not in os.listdir():
        if flag == 1:
            print('Creating folder ' + folder + '...')
            print(" ")
            flag = 2
        continue

#%% Download eTender files using CURL on cmd and saving to directory

os.chdir(wds[1])

print("Downloading eTender files...")
print(" ")
for n in curl_func:
    process = subprocess.Popen(n,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE,
                                shell = True)

downloadDateTime = datetime.today().strftime('%d-%b-%Y %I:%M %p')

print("Downloading eTender files complete")
print(" ")
#%% Combining eTender files

files = [n for n in os.listdir() if 'eTender_' in n]

flag = True
for file in files:
    if flag == True:
        df_eTender = pd.read_excel(file, header=2)
        df_eTender['Source Name'] = file
        df_eTender['ExportDateTime'] = downloadDateTime
        flag = False
        continue
    
    temp_df = pd.read_excel(file, header=2)
    temp_df['Source Name'] = file
    df_eTender['ExportDateTime'] = downloadDateTime
    df_eTender = df_eTender.append(temp_df, ignore_index=True)
    
#%% Re-formatting data

# Change currency object to float
# df_eTender['Estimated amount payable to the contractor (including GST)']=df_eTender['Estimated amount payable to the contractor (including GST)'].replace('[\$,]', '', regex=True).astype('float')

#%% Saving combined file to directory

os.chdir(wds[2])

print("Combining eTender files...")
print(" ")
df_eTender.to_csv(datetime.today().strftime('%Y%m%d')+'.csv', index=False)
print("Combining eTender files complete")
print(" ")

#%% Go back to original directory
os.chdir(wds[0])

print("Exporting complete. Press Enter to exit.")
print(" ")
print("==================================================================================")



_ = input('')

# time.sleep(4)

