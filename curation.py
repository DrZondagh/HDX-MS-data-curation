"""
Created on Tue Jun  6 16:43:03 2017

@author: J.Zondagh
"""
import pandas as pd
import glob
import xlwt
import csv
import os



# This saves every csv file (in specific folder) into a seperate worksheet in the same xls file (saved in same folder)
excel_writer = pd.ExcelWriter(r"PATH")

for filename in glob.glob(r"PATH*"):
    wb = xlwt.Workbook()
    (f_path, f_name) = os.path.split(filename)

    print(f_path)
    print(r"PATH*.csv"%f_name) 

    for filename_2 in glob.glob(r"PATH"%f_name):
        (f_path_2, f_name_2) = os.path.split(filename_2)

        print(f_path_2)
        print(f_name_2)

        (f_short_name, f_extension) = os.path.splitext(f_name_2)
        ws = wb.add_sheet(f_short_name)
        spamReader = csv.reader(open(filename_2, 'r'))
        for rowx, row in enumerate(spamReader):
            for colx, value in enumerate(row):
                ws.write(rowx, colx, value)
    wb.save(f_path_2+"\Mono_Pep_Files_merged.xls")
    
    
    # The following will merge all the worksheets (from all the files in filename_2) into a single worksheet (saved in desktop)
    workbook = pd.read_excel(f_path_2+"\Mono_Pep_Files_merged.xls", sheetname=None,header=None)
    merged = pd.concat(workbook, axis=1, ignore_index=False)
    merged.to_excel(excel_writer,sheet_name=f_name)
